
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
CREATE TABLE IF NOT EXISTS server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS report_snapshots (
 report_id TEXT PRIMARY KEY, created_utc TEXT NOT NULL, content_json TEXT NOT NULL,
 sha256 TEXT NOT NULL
);
