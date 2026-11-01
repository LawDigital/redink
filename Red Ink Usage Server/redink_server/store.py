"""Transactional SQLite persistence. A monthly event is not a presence signal."""
from __future__ import annotations

import hashlib
import json
import logging
import os
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path

from .config import Settings
from .protocol import ProtocolError, safe_product_id
from .retention import effective_cutoff


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat(timespec="microseconds").replace("+00:00", "Z")


CURRENT_SCHEMA_VERSION = 3

SCHEMA = """
CREATE TABLE IF NOT EXISTS billing_usage (
 billing_event_id TEXT PRIMARY KEY, month TEXT NOT NULL, network_key TEXT NOT NULL,
 network_scope TEXT NOT NULL, allowed_network_ids TEXT NOT NULL,
 product_id TEXT NOT NULL, product_id_safe TEXT NOT NULL, user_key TEXT NOT NULL,
 anonymous INTEGER NOT NULL, user_obf TEXT NOT NULL, first_use_utc TEXT NOT NULL,
 first_received_utc TEXT NOT NULL, last_received_utc TEXT NOT NULL, source TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS billing_month ON billing_usage(month,network_key,product_id);
CREATE TABLE IF NOT EXISTS host_usage (
 host_event_id TEXT PRIMARY KEY, billing_event_id TEXT NOT NULL REFERENCES billing_usage(billing_event_id),
 host_id TEXT NOT NULL, first_received_utc TEXT NOT NULL, last_received_utc TEXT NOT NULL,
 last_use_utc TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS host_billing ON host_usage(billing_event_id,host_id);
CREATE TABLE IF NOT EXISTS billing_hosts (
 billing_event_id TEXT NOT NULL REFERENCES billing_usage(billing_event_id), host_id TEXT NOT NULL,
 PRIMARY KEY(billing_event_id,host_id)
);
CREATE TABLE IF NOT EXISTS daily_usage (
 billing_event_id TEXT NOT NULL REFERENCES billing_usage(billing_event_id), loaded_date TEXT NOT NULL,
 first_received_utc TEXT NOT NULL, last_received_utc TEXT NOT NULL, source TEXT NOT NULL,
 PRIMARY KEY(billing_event_id,loaded_date)
);
CREATE INDEX IF NOT EXISTS daily_date ON daily_usage(loaded_date,billing_event_id);
CREATE TABLE IF NOT EXISTS recent_reports (
 host_event_id TEXT PRIMARY KEY REFERENCES host_usage(host_event_id),
 last_use_utc TEXT NOT NULL, reported_utc TEXT NOT NULL, received_utc TEXT NOT NULL,
 report_interval_minutes INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS recent_seen ON recent_reports(last_use_utc);
CREATE TABLE IF NOT EXISTS collector_errors (
 error_event_id TEXT PRIMARY KEY, month TEXT NOT NULL, error_code TEXT NOT NULL,
 related_file TEXT NOT NULL, detail TEXT NOT NULL, first_received_utc TEXT NOT NULL,
 last_received_utc TEXT NOT NULL, source TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS month_revisions (
 month TEXT PRIMARY KEY, revision INTEGER NOT NULL, changed_utc TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS log_events (
 record_id TEXT PRIMARY KEY, date TEXT NOT NULL, timestamp_local TEXT NOT NULL,
 host_id TEXT NOT NULL, user_hash TEXT NOT NULL, version TEXT NOT NULL,
 feature TEXT NOT NULL, invoked INTEGER NOT NULL, source TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS log_date ON log_events(date,host_id);
CREATE TABLE IF NOT EXISTS source_status (
 kind TEXT NOT NULL, path TEXT NOT NULL, last_scan_utc TEXT NOT NULL,
 records INTEGER NOT NULL DEFAULT 0, errors INTEGER NOT NULL DEFAULT 0,
 skipped INTEGER NOT NULL DEFAULT 0, detail TEXT NOT NULL DEFAULT '',
 PRIMARY KEY(kind,path)
);
CREATE TABLE IF NOT EXISTS imported_files (
 kind TEXT NOT NULL, path TEXT NOT NULL, size INTEGER NOT NULL, mtime_ns INTEGER NOT NULL,
 records INTEGER NOT NULL, errors INTEGER NOT NULL, skipped INTEGER NOT NULL,
 last_scan_utc TEXT NOT NULL, detail TEXT NOT NULL, PRIMARY KEY(kind,path)
);
CREATE TABLE IF NOT EXISTS log_scan_cursor (
 path TEXT PRIMARY KEY, offset_bytes INTEGER NOT NULL,
 line_number INTEGER NOT NULL, head_signature TEXT NOT NULL, tail_signature TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS schema_migrations (
 version INTEGER PRIMARY KEY, applied_utc TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS report_snapshots (
 report_id TEXT PRIMARY KEY, created_utc TEXT NOT NULL, content_json TEXT NOT NULL,
 sha256 TEXT NOT NULL
);
"""


class Store:
    def __init__(self, settings: Settings, clock=utc_now):
        self.settings = settings
        self.path = settings.database_path
        self.clock = clock
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._migrate_database()

    def _migrate_database(self) -> None:
        """Automatic, transactional, backed-up upgrade when a newer release starts.

        No separate migration command is required. Only new/existing versioned
        schemas are supported; newer versions fail closed instead of downgrading.
        """
        with self.connect() as connection:
            version = int(connection.execute("PRAGMA user_version").fetchone()[0])
            if version > CURRENT_SCHEMA_VERSION:
                raise RuntimeError(
                    f"Database version {version} is newer than supported version "
                    f"{CURRENT_SCHEMA_VERSION}. Start a compatible server version."
                )
            if version < 0:
                raise RuntimeError("Invalid database schema version.")
            # Retain the original WAL journaling behaviour for concurrent
            # collector reads/writes; the backup API includes committed WAL pages.
            connection.execute("PRAGMA journal_mode=WAL")
            # Legacy databases sometimes have user_version=0 but already
            # contain v1 tables. Do not mistake them for a fresh database.
            has_user_tables = bool(connection.execute("""SELECT 1 FROM sqlite_master
                WHERE type='table' AND name NOT LIKE 'sqlite_%' LIMIT 1""").fetchone())
            migration_needed = version < CURRENT_SCHEMA_VERSION
            if not migration_needed:
                return
            if connection.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                raise RuntimeError("Database integrity check failed before migration; no changes were made.")
            backup = self._migration_backup(connection, version) if has_user_tables else None
            try:
                # Important: execute the complete migration as ONE transaction.
                # The schema is additive and preserves old monthly evidence,
                # daily observations, log events and immutable snapshots.
                connection.executescript("BEGIN IMMEDIATE;\n" + SCHEMA)
                connection.execute("""INSERT OR IGNORE INTO billing_hosts
                    SELECT billing_event_id, host_id FROM host_usage""")
                connection.execute("INSERT OR IGNORE INTO schema_migrations VALUES (?,?)",
                                   (CURRENT_SCHEMA_VERSION, iso(self.clock())))
                connection.execute(f"PRAGMA user_version={CURRENT_SCHEMA_VERSION}")
                if connection.execute("PRAGMA foreign_key_check").fetchone():
                    raise RuntimeError("Foreign-key validation failed during migration.")
                if connection.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise RuntimeError("Integrity validation failed during migration.")
                connection.execute("COMMIT")
            except Exception:
                connection.rollback()
                raise
            logging.getLogger(__name__).info(
                "Database schema migrated from v%s to v%s%s", version,
                CURRENT_SCHEMA_VERSION,
                f"; pre-migration backup: {backup}" if backup else "; fresh database")

    def _migration_backup(self, source: sqlite3.Connection, old_version: int) -> Path:
        """SQLite-online backup, including committed WAL pages, before altering schema."""
        directory = self.settings.data_dir / "backups"
        directory.mkdir(parents=True, exist_ok=True)
        destination_path = directory / (
            f"usage-before-schema-v{old_version}-to-v{CURRENT_SCHEMA_VERSION}-"
            f"{self.clock().astimezone(timezone.utc):%Y%m%dT%H%M%SZ}-"
            f"{secrets.token_hex(4)}.sqlite3")
        # Pre-create with restrictive permissions and exclusive creation;
        # sqlite3.connect() can then open the empty file normally.
        fd = os.open(destination_path, os.O_CREAT | os.O_EXCL | os.O_RDWR, 0o600)
        os.close(fd)
        try:
            target = sqlite3.connect(destination_path)
            try:
                source.backup(target)
                if target.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise RuntimeError("Pre-migration backup verification failed.")
            finally:
                target.close()
            return destination_path
        except Exception:
            destination_path.unlink(missing_ok=True)
            raise

    def retention_cutoff(self) -> str:
        return effective_cutoff(self.settings.cutoff_date, self.settings.keep_months,
                                self.clock().astimezone().date())

    @staticmethod
    def _is_expired(event: dict, cutoff: str) -> bool:
        if not cutoff:
            return False
        month = event["month"]
        if month < cutoff[:7]:
            return True
        if month > cutoff[:7]:
            return False
        # Boundary month: a legacy marker without explicit days is not
        # evidence that a client loaded the add-in after the cutoff.
        dates = event.get("loaded_dates") or []
        if any(day >= cutoff for day in dates):
            return False
        first = event.get("first_use_utc", "")
        return not first or first[:10] < cutoff

    @contextmanager
    def connect(self):
        connection = sqlite3.connect(self.path, timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=15000")
        connection.execute("PRAGMA synchronous=FULL")
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    @staticmethod
    def _revision(connection, month, stamp):
        connection.execute("INSERT INTO month_revisions VALUES (?,1,?) ON CONFLICT(month) DO UPDATE SET revision=revision+1,changed_utc=excluded.changed_utc", (month, stamp))

    def ingest(self, event: dict, source: str = "HTTP collector", *, live_request: bool = False) -> dict:
        """Commit monthly identity and explicit load dates before acknowledging.

        live_request is retained for import/HTTP caller compatibility. Receipt
        time and transport never manufacture a load observation.
        """
        now = self.clock()
        stamp = iso(now)
        kind = event["event_type"]
        cutoff = self.retention_cutoff()
        if cutoff and self._is_expired(event, cutoff):
            # Acknowledge but do not persist. A negative response would cause
            # old offline markers to retry forever, regardless of retention.
            return {"accepted": True, "discarded": True, "reason": "outside_retention",
                    "cutoff_date": cutoff, "event_type": kind, "server_utc": stamp}
        with self.connect() as connection:
            # Serialize read-validate-write across concurrent Waitress requests.
            connection.execute("BEGIN IMMEDIATE")
            if kind == "E":
                existing = connection.execute("SELECT error_code,month FROM collector_errors WHERE error_event_id=?", (event["error_event_id"],)).fetchone()
                if existing and (existing["error_code"] != event["error_code"] or existing["month"] != event["month"]):
                    raise ProtocolError("An error event ID was reused for different contents.")
                connection.execute("""INSERT INTO collector_errors VALUES (?,?,?,?,?,?,?,?)
                 ON CONFLICT(error_event_id) DO UPDATE SET last_received_utc=excluded.last_received_utc""",
                 (event["error_event_id"], event["month"], event["error_code"], event.get("related_file", ""),
                  event.get("detail", ""), stamp, stamp, source))
                result = {"accepted": True, "duplicate": existing is not None, "event_type": "E", "server_utc": stamp}
            elif kind == "U":
                # Month is client-local in the established contract. Allow one month
                # ahead at UTC boundary but never assign a replay to receipt month.
                next_month = (now.replace(day=28) + timedelta(days=4)).strftime("%Y-%m")
                if event["month"] > next_month:
                    raise ProtocolError("Usage month is implausibly far in the future.")
                loaded_dates = [day for day in event.get("loaded_dates", []) if not cutoff or day >= cutoff]
                latest_local_day = (now + timedelta(hours=14)).date().isoformat()
                if any(day > latest_local_day for day in loaded_dates):
                    raise ProtocolError("An add-in load date is implausibly far in the future.")
                billing = event["billing_event_id"]
                existing = connection.execute("SELECT * FROM billing_usage WHERE billing_event_id=?", (billing,)).fetchone()
                host_row = connection.execute("SELECT * FROM host_usage WHERE host_event_id=?", (event["host_event_id"],)).fetchone()
                if host_row and (host_row["billing_event_id"] != billing or host_row["host_id"] != event["host_id"]):
                    raise ProtocolError("A host event ID was reused for a different identity.")
                changed = existing is None or host_row is None
                if existing:
                    if any(existing[k] != event[k] for k in ("month", "network_key", "user_key")):
                        raise ProtocolError("A billing event ID was reused for a different identity.")
                    if safe_product_id(existing["product_id"]) != safe_product_id(event["product_id"]):
                        raise ProtocolError("A billing event ID was reused for a different product.")
                    product = existing["product_id"]
                    # v1 PATH and QUERY only carry a sanitized product. A later
                    # JSON/marker record may supply the original spelling.
                    if event["product_id"] != event["product_id_safe"] and product == existing["product_id_safe"]:
                        product = event["product_id"]
                    scope = event["network_scope"] or existing["network_scope"]
                    networks = sorted(set(json.loads(existing["allowed_network_ids"])) | set(event["allowed_network_ids"]))
                    first_use = min(filter(None, (existing["first_use_utc"], event["first_use_utc"])), default="")
                    # An anonymized replay removes reversible identity information.
                    anonymous = bool(existing["anonymous"]) or event["anonymous"]
                    obf = "" if anonymous else (event["user_obf"] or existing["user_obf"])
                    changed = changed or product != existing["product_id"] or scope != existing["network_scope"] or first_use != existing["first_use_utc"] or json.dumps(networks) != existing["allowed_network_ids"] or anonymous != bool(existing["anonymous"]) or obf != existing["user_obf"]
                    connection.execute("""UPDATE billing_usage SET product_id=?,network_scope=?,allowed_network_ids=?,
                     first_use_utc=?,last_received_utc=?,anonymous=?,user_obf=? WHERE billing_event_id=?""",
                     (product, scope, json.dumps(networks), first_use, stamp, int(anonymous), obf, billing))
                else:
                    connection.execute("INSERT INTO billing_usage VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                      (billing, event["month"], event["network_key"], event["network_scope"], json.dumps(sorted(set(event["allowed_network_ids"]))),
                       event["product_id"], event["product_id_safe"], event["user_key"], int(event["anonymous"]), event["user_obf"],
                       event["first_use_utc"], stamp, stamp, source))
                connection.execute("""INSERT INTO host_usage VALUES (?,?,?,?,?,?) ON CONFLICT(host_event_id)
                  DO UPDATE SET last_received_utc=excluded.last_received_utc,
                  last_use_utc=MAX(host_usage.last_use_utc,excluded.last_use_utc)""",
                  (event["host_event_id"], billing, event["host_id"], stamp, stamp, event.get("last_use_utc", "")))
                hosts_added = 0
                for host in event.get("loaded_hosts", [event["host_id"]]):
                    hosts_added += connection.execute("INSERT OR IGNORE INTO billing_hosts VALUES (?,?)", (billing, host)).rowcount
                dates_added = 0
                for day in loaded_dates:
                    dates_added += connection.execute("INSERT OR IGNORE INTO daily_usage VALUES (?,?,?,?,?)", (billing, day, stamp, stamp, source)).rowcount
                    connection.execute("UPDATE daily_usage SET last_received_utc=? WHERE billing_event_id=? AND loaded_date=?", (stamp, billing, day))
                changed = changed or bool(dates_added) or bool(hosts_added)
                if changed:
                    self._revision(connection, event["month"], stamp)
                result = {"accepted": True, "duplicate": existing is not None and host_row is not None,
                          "event_type": "U", "month": event["month"], "server_utc": stamp,
                          "loaded_dates_added": dates_added, "recent_report_updated": False}
            else:
                raise ProtocolError("Unsupported normalized event type.")
            connection.execute("INSERT INTO server_meta VALUES ('last_ingest_utc',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (stamp,))
        return result

    def monthly_rows(self, start: str = "0001-01", end: str = "9999-12") -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute("""WITH daily AS (
              SELECT month,network_key,product_id,COUNT(*) AS observed_user_days FROM (
                SELECT DISTINCT b.month,b.network_key,b.product_id,b.user_key,d.loaded_date
                FROM daily_usage d JOIN billing_usage b USING(billing_event_id)
              ) GROUP BY month,network_key,product_id
            )
            SELECT b.month,b.network_key,MAX(b.network_scope) AS network_scope,b.product_id,
             COUNT(DISTINCT b.user_key) AS users,
             COUNT(DISTINCT CASE WHEN h.host_id='WD' THEN b.user_key END) AS word_users,
             COUNT(DISTINCT CASE WHEN h.host_id='EX' THEN b.user_key END) AS excel_users,
             COUNT(DISTINCT CASE WHEN h.host_id='OL' THEN b.user_key END) AS outlook_users,
             COUNT(DISTINCT CASE WHEN EXISTS(SELECT 1 FROM daily_usage d WHERE d.billing_event_id=b.billing_event_id) THEN b.user_key END) AS users_with_daily_evidence,
             COALESCE(MAX(daily.observed_user_days),0) AS observed_user_days,
             MIN(b.first_received_utc) AS first_received_utc,MAX(b.last_received_utc) AS last_received_utc,
             MAX(hu.last_use_utc) AS last_use_utc,r.revision,r.changed_utc
             FROM billing_usage b LEFT JOIN billing_hosts h USING(billing_event_id)
             LEFT JOIN (SELECT billing_event_id,MAX(last_use_utc) AS last_use_utc FROM host_usage GROUP BY billing_event_id) hu USING(billing_event_id)
             LEFT JOIN daily ON daily.month=b.month AND daily.network_key=b.network_key AND daily.product_id=b.product_id
             JOIN month_revisions r ON r.month=b.month WHERE b.month BETWEEN ? AND ?
             GROUP BY b.month,b.network_key,b.product_id ORDER BY b.month,b.network_key,b.product_id""", (start, end)).fetchall()
        result = [dict(row) for row in rows]
        for row in result:
            row["monthly_only_users"] = row["users"] - row["users_with_daily_evidence"]
        return result

    def user_rows(self, month: str) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute("""SELECT b.month,b.network_key,b.network_scope,b.product_id,b.user_key,
             GROUP_CONCAT(DISTINCT h.host_id) AS hosts,MIN(b.first_use_utc) AS first_use_utc,
             MIN(b.first_received_utc) AS first_received_utc,MAX(b.last_received_utc) AS last_received_utc,
             MAX(hu.last_use_utc) AS last_use_utc,
             GROUP_CONCAT(DISTINCT d.loaded_date) AS loaded_dates,COUNT(DISTINCT d.loaded_date) AS loaded_days
             FROM billing_usage b LEFT JOIN billing_hosts h USING(billing_event_id)
             LEFT JOIN (SELECT billing_event_id,MAX(last_use_utc) AS last_use_utc FROM host_usage GROUP BY billing_event_id) hu USING(billing_event_id)
             LEFT JOIN daily_usage d USING(billing_event_id) WHERE b.month=?
             GROUP BY b.month,b.network_key,b.product_id,b.user_key ORDER BY b.network_key,b.product_id,b.user_key""", (month,)).fetchall()
        result = [dict(row) for row in rows]
        for row in result:
            row["hosts"] = ",".join(sorted((row["hosts"] or "").split(",")))
            row["loaded_dates"] = ",".join(sorted((row["loaded_dates"] or "").split(",")))
        return result

    def daily_rows(self, start: str, end: str) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute("""SELECT d.loaded_date AS date,b.network_key,MAX(b.network_scope) AS network_scope,
             b.product_id,COUNT(DISTINCT b.user_key) AS users,
             MIN(d.first_received_utc) AS first_received_utc,MAX(d.last_received_utc) AS last_received_utc
             FROM daily_usage d JOIN billing_usage b USING(billing_event_id)
             WHERE d.loaded_date BETWEEN ? AND ?
             GROUP BY d.loaded_date,b.network_key,b.product_id ORDER BY d.loaded_date,b.network_key,b.product_id""", (start, end)).fetchall()
        return [dict(row) for row in rows]

    def dashboard(self, month: str, day: str | None = None) -> dict:
        now = self.clock()
        day = day or now.astimezone().date().isoformat()
        with self.connect() as connection:
            events = connection.execute("SELECT COUNT(*) FROM host_usage JOIN billing_usage USING(billing_event_id) WHERE month=?", (month,)).fetchone()[0]
            errors = connection.execute("SELECT COUNT(*) FROM collector_errors").fetchone()[0]
            last = connection.execute("SELECT value FROM server_meta WHERE key='last_ingest_utc'").fetchone()
            sources = [dict(r) for r in connection.execute("SELECT * FROM source_status ORDER BY kind,path")]
            recent_errors = [dict(r) for r in connection.execute("SELECT error_code,month,first_received_utc,last_received_utc FROM collector_errors ORDER BY last_received_utc DESC LIMIT 20")]
        monthly = self.monthly_rows()
        selected = [r for r in monthly if r["month"] == month]
        history = {}
        for row in monthly:
            h = history.setdefault(row["month"], {"month": row["month"], "users": 0, "host_users": 0})
            h["users"] += row["users"]
            h["host_users"] += row["word_users"] + row["excel_users"] + row["outlook_users"]
        daily_history = {}
        for row in self.daily_rows(month + "-01", month + "-31"):
            daily_history[row["date"]] = daily_history.get(row["date"], 0) + row["users"]
        return {
            "day": {"date": day, "users": sum(r["users"] for r in self.daily_rows(day, day)),
                    "basis": "Explicit client-local add-in load dates; one user per scope and product per day.",
                    "as_of": iso(now)},
            "month": {"month": month, "users": sum(r["users"] for r in selected), "events": events,
                      "users_with_daily_evidence": sum(r["users_with_daily_evidence"] for r in selected),
                      "monthly_only_users": sum(r["monthly_only_users"] for r in selected),
                      "observed_user_days": sum(r["observed_user_days"] for r in selected),
                      "hosts": {h: sum(r[k] for r in selected) for h, k in (("WD", "word_users"), ("EX", "excel_users"), ("OL", "outlook_users"))}},
            "daily_history": [{"date": day, "users": count} for day, count in daily_history.items()],
            "history": list(history.values()), "monthly_rows": selected,
            "health": {"database": "Ready", "status": "attention" if errors or any(s["errors"] for s in sources) else "ok",
                       "last_ingest_utc": last[0] if last else None, "pending_errors": errors},
            "sources": sources, "client_errors": recent_errors, "logs": self.log_summary(month),
        }

    def add_log_records(self, records: list[tuple[str, dict]], source: str) -> int:
        added = 0
        with self.connect() as connection:
            cutoff = self.retention_cutoff()
            for key, record in records:
                if cutoff and record["date"] < cutoff:
                    continue
                cursor = connection.execute("INSERT OR IGNORE INTO log_events VALUES (?,?,?,?,?,?,?,?,?)",
                 (key, record["date"], record["timestamp_local"], record["host_id"], record["user_hash"],
                  record["version"], record["feature"], int(record["invoked"]), source))
                added += cursor.rowcount
        return added

    def log_summary(self, month: str = "") -> dict:
        where = " WHERE date LIKE ?" if month else ""
        args = (month + "%",) if month else ()
        with self.connect() as connection:
            records, users = connection.execute("SELECT COUNT(*),COUNT(DISTINCT user_hash) FROM log_events" + where, args).fetchone()
            files = connection.execute("SELECT COUNT(*) FROM imported_files WHERE kind='log' AND records>0").fetchone()[0]
            errors = connection.execute("SELECT COALESCE(SUM(errors),0) FROM source_status WHERE kind='log'").fetchone()[0]
            result = {"files": files, "records": records, "users": users, "errors": errors, "date_basis": "client-local",
                      "identity_basis": "SharedLogger username hash; unrelated to licence user keys"}
            for column, alias, output in (("host_id", "host", "by_host"), ("version", "version", "by_version"),
                                          ("date", "date", "by_day"), ("feature", "feature", "by_feature")):
                sql = f"SELECT {column} AS {alias},COUNT(*) AS count FROM log_events{where} GROUP BY {column} ORDER BY " + (column if column == "date" else "count DESC") + " LIMIT 500"
                result[output] = [dict(row) for row in connection.execute(sql, args)]
        return result

    def log_analysis(self, period: str = "30d", host: str = "ALL") -> dict:
        """Metadata-only log trends, independent of billing/loadedDates.

        Client-local days are never reinterpreted as UTC timestamps. Missing
        dates are unobserved; no synthetic zeros are added to trends.
        """
        from collections import Counter
        from .log_taxonomy import load_taxonomy
        from .usage_semantics import describe, primary_family

        windows = {"7d": 7, "30d": 30, "90d": 90, "365d": 365, "all": None}
        if period not in windows or host not in {"ALL", "WD", "EX", "OL"}:
            raise ValueError("Invalid log period or Office host.")
        taxonomy = load_taxonomy(self.settings.data_dir)
        days = windows[period]
        date_from = (self.clock().date() - timedelta(days=days - 1)).isoformat() if days else None
        where = []
        args = []
        if date_from:
            where.append("date >= ?")
            args.append(date_from)
            where.append("date <= ?")
            args.append(self.clock().date().isoformat())
        if host != "ALL":
            where.append("host_id = ?")
            args.append(host)
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        with self.connect() as connection:
            totals = connection.execute("SELECT COUNT(*),COUNT(DISTINCT user_hash),MAX(timestamp_local) FROM log_events" + clause, args).fetchone()
            inv_clause = clause + (" AND invoked=1" if clause else " WHERE invoked=1")
            rows = connection.execute(
                "SELECT feature,host_id,date,COUNT(*) AS count FROM log_events" + inv_clause +
                " GROUP BY feature,host_id,date ORDER BY date,feature,host_id", args).fetchall()
            versions = [dict(r) for r in connection.execute(
                "SELECT version,COUNT(*) AS count FROM log_events" + clause +
                " GROUP BY version ORDER BY count DESC,version LIMIT 100", args)]
            source_rows = [dict(r) for r in connection.execute(
                "SELECT path,last_scan_utc,records,errors,detail FROM source_status WHERE kind='log' ORDER BY path")]
            total_files = connection.execute("SELECT COUNT(*) FROM imported_files WHERE kind='log' AND records>0").fetchone()[0]
        features = Counter()
        agents = Counter()
        families = Counter()
        hosts = Counter()
        host_agents = Counter()
        daily = Counter()
        daily_agents = Counter()
        observed = 0
        agent_total = 0
        secondary = Counter()
        kinds = Counter()
        raw_invocations = 0
        members = {}
        for row in rows:
            count = row["count"]
            raw_invocations += count
            event = describe(row['feature'])
            kinds[event.kind] += count
            if not event.primary:
                secondary[(event.kind,event.label)] += count
                continue
            observed += count
            feature = event.label
            group = primary_family(event, taxonomy)
            members.setdefault(group['id'],Counter())[feature] += count
            features[(feature, group["id"], group["label"], group["agentic"])] += count
            families[(group["id"], group["label"], group["agentic"])] += count
            hosts[row["host_id"]] += count
            daily[row["date"]] += count
            if group["agentic"]:
                agent_total += count
                agents[feature] += count
                host_agents[row["host_id"]] += count
                daily_agents[row["date"]] += count
        grouped_features = [dict(feature=key[0],group=key[1],group_label=key[2],agentic=key[3],count=count)
                            for key,count in features.most_common(200)]
        return {
            "period": period, "host_filter": host, "from_date": date_from,
            "to_date": self.clock().date().isoformat() if days else None,
            "date_basis": "client-local; timezone not included in SharedLogger lines",
            "identity_basis": "username hash from SharedLogger, NOT a license identity",
            "files": total_files, "scanned_sources": len(source_rows),
            "source_status": source_rows,
            "import_errors": sum(row["errors"] for row in source_rows),
            "entries": totals[0], "invocations": observed,
            "other_entries": totals[0] - raw_invocations,
            "raw_invocations": raw_invocations, "event_kinds": dict(kinds),
            "tool_details": [dict(kind=k[0],feature=k[1],count=n) for k,n in secondary.most_common() if k[0] in ('tool','skill','agent')],
            "legacy_markers": [dict(feature=k[1],count=n) for k,n in secondary.most_common() if k[0]=='legacy'],
            "users": totals[1], "last_entry_local": totals[2],
            "agentic_invocations": agent_total,
            "standard_invocations": observed - agent_total,
            "agentic_basis": "Explicit top-level tooling-loop starts only; invocation is not completion. Older logs do not establish loop starts.",
            "features": grouped_features,
            "agents": [dict(feature=name,count=count) for name,count in agents.most_common(60)],
            "families": [dict(id=k[0],label=k[1],agentic=k[2],count=v,commands=[dict(feature=f,count=n) for f,n in members[k[0]].most_common()]) for k,v in families.most_common()],
            "hosts": [dict(host=h, count=hosts[h], agentic=host_agents[h],standard=hosts[h]-host_agents[h])
                      for h in ("WD", "EX", "OL")],
            "days": [dict(date=d,count=daily[d],agentic=daily_agents[d],standard=daily[d]-daily_agents[d])
                     for d in sorted(daily)],
            "versions": versions,
        }

    def log_scan_cursor(self, path: Path) -> dict | None:
        with self.connect() as connection:
            row=connection.execute("SELECT * FROM log_scan_cursor WHERE path=?",(str(path),)).fetchone()
        return dict(row) if row else None

    def save_log_progress(self, path: Path, stat, *, rows: list[tuple[str,dict]],
                          legacy_keys: list[tuple[str,str,dict]] | None = None,
                          offset: int, line_number: int, head: str, tail: str,
                          records: int, errors: int, skipped: int, detail: str) -> None:
        """Commit inserted events and the byte cursor in one transaction."""
        with self.connect() as connection:
            # One-time migration of v1.1's occurrence-based log IDs to stable
            # line-based IDs, without double-counting previously imported data.
            # This also refreshes old parser classifications/versions in place.
            cutoff = self.retention_cutoff()
            for old, new, record in (legacy_keys or ()):
                if cutoff and record["date"] < cutoff:
                    connection.execute("DELETE FROM log_events WHERE record_id IN (?,?)", (old, new))
                    continue
                current = connection.execute("SELECT 1 FROM log_events WHERE record_id=?",(old,)).fetchone()
                if current:
                    already = connection.execute("SELECT 1 FROM log_events WHERE record_id=?",(new,)).fetchone()
                    if already and old != new:
                        connection.execute("DELETE FROM log_events WHERE record_id=?",(old,))
                    else:
                        connection.execute("""UPDATE log_events SET record_id=?,date=?,timestamp_local=?,host_id=?,
                            user_hash=?,version=?,feature=?,invoked=? WHERE record_id=?""",
                            (new,record["date"],record["timestamp_local"],record["host_id"],
                             record["user_hash"],record["version"],record["feature"],int(record["invoked"]),old))
            for key, record in rows:
                if cutoff and record["date"] < cutoff:
                    continue
                connection.execute("INSERT OR IGNORE INTO log_events VALUES (?,?,?,?,?,?,?,?,?)",
                    (key, record["date"],record["timestamp_local"],record["host_id"],record["user_hash"],
                     record["version"],record["feature"],int(record["invoked"]),str(path)))
            connection.execute("INSERT OR REPLACE INTO log_scan_cursor VALUES (?,?,?,?,?)",
                (str(path),offset,line_number,head,tail))
            connection.execute("INSERT OR REPLACE INTO imported_files VALUES (?,?,?,?,?,?,?,?,?)",
                ("log",str(path),stat.st_size,stat.st_mtime_ns,records,errors,skipped,iso(self.clock()),detail[:2000]))

    def source_status(self, kind, path, records=0, errors=0, skipped=0, detail=""):
        with self.connect() as connection:
            connection.execute("INSERT OR REPLACE INTO source_status VALUES (?,?,?,?,?,?,?)",
                               (kind, str(path), iso(self.clock()), records, errors, skipped, detail[:2000]))

    def file_status(self, kind, path):
        with self.connect() as connection:
            row = connection.execute("SELECT * FROM imported_files WHERE kind=? AND path=?", (kind, str(path))).fetchone()
        return dict(row) if row else None

    def save_file_status(self, kind, path, stat, records, errors, skipped, detail):
        with self.connect() as connection:
            connection.execute("INSERT OR REPLACE INTO imported_files VALUES (?,?,?,?,?,?,?,?,?)",
             (kind, str(path), stat.st_size, stat.st_mtime_ns, records, errors, skipped, iso(self.clock()), detail[:2000]))

    def retention_preview(self) -> dict:
        """Exact count for the current effective cutoff, no database changes."""
        cutoff = self.retention_cutoff()
        result = {"cutoff_date": cutoff, "enabled": bool(cutoff), "counts": {},
                  "warning": "Deleting saved evidence cannot be undone without a backup."}
        keys = ("billing_usage", "daily_usage", "log_events", "collector_errors",
                "report_snapshots", "host_usage", "billing_hosts", "recent_reports", "month_revisions")
        result["counts"] = dict.fromkeys(keys, 0)
        if not cutoff:
            return result
        with self.connect() as con:
            # A boundary-month identity survives only with an explicit newer day
            # or a first-use date on/after the configured cutoff.
            expired_billing = """b.month < ? OR (b.month=? AND substr(b.first_use_utc,1,10) < ?
                    AND NOT EXISTS (SELECT 1 FROM daily_usage d WHERE d.billing_event_id=b.billing_event_id AND d.loaded_date >= ?))"""
            args = (cutoff[:7], cutoff[:7], cutoff, cutoff)
            result["counts"]["billing_usage"] = con.execute("SELECT COUNT(*) FROM billing_usage b WHERE " + expired_billing,args).fetchone()[0]
            result["counts"]["daily_usage"] = con.execute("SELECT COUNT(*) FROM daily_usage d WHERE d.loaded_date < ? OR d.billing_event_id IN (SELECT b.billing_event_id FROM billing_usage b WHERE " + expired_billing + ")", (cutoff,*args)).fetchone()[0]
            for table,column in (("log_events","date"),("collector_errors","month"),("month_revisions","month")):
                result["counts"][table] = con.execute(f"SELECT COUNT(*) FROM {table} WHERE {column} < ?", (cutoff if column=='date' else cutoff[:7],)).fetchone()[0]
            for table in ("host_usage", "billing_hosts", "recent_reports"):
                if table == "recent_reports":
                    expr = "host_event_id IN (SELECT host_event_id FROM host_usage WHERE billing_event_id IN (SELECT b.billing_event_id FROM billing_usage b WHERE " + expired_billing + "))"
                else:
                    expr = "billing_event_id IN (SELECT b.billing_event_id FROM billing_usage b WHERE " + expired_billing + ")"
                result["counts"][table] = con.execute("SELECT COUNT(*) FROM " + table + " WHERE " + expr, args).fetchone()[0]
            # Snapshots whose period ended before the boundary month, or which
            # were themselves generated before the cutoff, are eligible.
            result["counts"]["report_snapshots"] = sum(self._snapshot_is_expired(row,cutoff) for row in
                con.execute("SELECT created_utc,content_json FROM report_snapshots"))
        result["total"] = sum(result["counts"].values())
        return result

    @staticmethod
    def _snapshot_is_expired(row, cutoff: str) -> bool:
        if row["created_utc"][:10] < cutoff:
            return True
        try:
            end = json.loads(row["content_json"]).get("to_month", "")
            return bool(end and end < cutoff[:7])
        except (ValueError, AttributeError, TypeError):
            return False

    def clean_retention(self, expected_cutoff: str) -> dict:
        """Delete old stored evidence in one atomic transaction; keep scan cursors."""
        cutoff = self.retention_cutoff()
        if not cutoff or cutoff != expected_cutoff:
            raise ValueError("Retention cutoff has changed. Refresh the preview before deleting.")
        snapshot = self.retention_preview()
        with self.connect() as con:
            con.execute("BEGIN IMMEDIATE")
            expired_billing = """month < ? OR (month=? AND substr(first_use_utc,1,10) < ?
                    AND NOT EXISTS (SELECT 1 FROM daily_usage d WHERE d.billing_event_id=billing_usage.billing_event_id AND d.loaded_date >= ?))"""
            con.execute("DELETE FROM log_events WHERE date < ?", (cutoff,))
            con.execute("DELETE FROM daily_usage WHERE loaded_date < ?", (cutoff,))
            # In boundary month, first-use before cutoff with a qualifying
            # newer observed day must be retained. Check after removing old days.
            args = (cutoff[:7],cutoff[:7],cutoff,cutoff)
            ids = [r[0] for r in con.execute("SELECT billing_event_id FROM billing_usage WHERE " + expired_billing,args)]
            for billing in ids:
                con.execute("DELETE FROM recent_reports WHERE host_event_id IN (SELECT host_event_id FROM host_usage WHERE billing_event_id=?)",(billing,))
                con.execute("DELETE FROM billing_hosts WHERE billing_event_id=?",(billing,))
                con.execute("DELETE FROM host_usage WHERE billing_event_id=?",(billing,))
                con.execute("DELETE FROM daily_usage WHERE billing_event_id=?",(billing,))
                con.execute("DELETE FROM billing_usage WHERE billing_event_id=?",(billing,))
            con.execute("DELETE FROM collector_errors WHERE month < ?", (cutoff[:7],))
            con.execute("DELETE FROM month_revisions WHERE month < ?", (cutoff[:7],))
            for row in con.execute("SELECT report_id,created_utc,content_json FROM report_snapshots"):
                if self._snapshot_is_expired(row,cutoff):
                    con.execute("DELETE FROM report_snapshots WHERE report_id=?",(row["report_id"],))
            con.execute("INSERT INTO server_meta VALUES ('last_retention_cleanup_utc',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",(iso(self.clock()),))
        return {"cleaned": True, "cutoff_date": cutoff, "preview_counts": snapshot["counts"],
                "remaining": self.retention_preview()["counts"]}

    def maintenance(self):
        with self.connect() as connection:
            # The former recent-report cache is no longer used for counting.
            # Preserve migrated evidence and all explicit daily/monthly history.
            connection.execute("PRAGMA optimize")
        if self.settings.auto_cleanup and self.retention_cutoff():
            self.clean_retention(self.retention_cutoff())

    def check(self) -> str:
        with self.connect() as connection:
            return connection.execute("PRAGMA quick_check").fetchone()[0]

    def backup(self, output: Path) -> None:
        output = output.resolve()
        if output.exists() or output == self.path.resolve():
            raise ValueError("Backup destination already exists. Choose a new file.")
        output.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as source:
            destination = sqlite3.connect(output)
            try:
                source.backup(destination)
                if destination.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise RuntimeError("Backup verification failed.")
            finally:
                destination.close()

    def save_report_snapshot(self, report: dict) -> dict:
        content = json.dumps(report, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
        digest = hashlib.sha256(content.encode("utf-8")).hexdigest()
        report_id = "RI-" + digest[:20]
        with self.connect() as connection:
            connection.execute("INSERT OR IGNORE INTO report_snapshots VALUES (?,?,?,?)", (report_id, iso(self.clock()), content, digest))
        return {"report_id": report_id, "sha256": digest, "report": report}
