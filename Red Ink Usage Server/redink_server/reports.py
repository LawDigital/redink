"""Auditable usage reports with optional exact decimal pricing, without tax assumptions."""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from .store import Store, iso


def valid_month(value: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[1-9][0-9]{3}-(?:0[1-9]|1[0-2])", value):
        raise ValueError("Month must use yyyy-MM.")
    return value


def range_months(start: str, end: str) -> tuple[str, str]:
    start, end = valid_month(start), valid_month(end)
    if start > end:
        raise ValueError("The first month must be on or before the last month.")
    return start, end


def valid_day(value: str) -> str:
    try:
        if not isinstance(value, str) or not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value) or date.fromisoformat(value).isoformat() != value:
            raise ValueError()
    except ValueError:
        raise ValueError("Date must be a real calendar date in yyyy-MM-dd format.") from None
    return value


def monthly_report(store: Store, start: str, end: str) -> dict:
    start, end = range_months(start, end)
    rows = store.monthly_rows(start, end)
    price = store.settings.unit_price
    for row in rows:
        row["coverage"] = "observed; late reports may revise this month"
        if price:
            row["unit_price"] = price
            row["currency"] = store.settings.currency
            row["amount"] = str((Decimal(price) * row["users"]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    result = {
        "schema_version": 2, "generated_utc": iso(store.clock()), "from_month": start, "to_month": end,
        "basis": "Distinct user keys with an add-in loaded during the month, per network scope and product; Office hosts are deduplicated. AI commands and daily totals do not determine monthly billing units.",
        "coverage": "Observed monthly usage, not proof of completeness. Late offline reports revise historical months. An absent month means no evidence, not confirmed zero usage.",
        "daily_evidence": "Only explicit client-local loadedDates contribute to daily counts. Legacy monthly markers remain valid monthly evidence; dates are never inferred from receipt time or timestamp ranges.",
        "identity": "Add-in licence user identity; changing licence scope, user ID or salt can change identity.",
        "tax": "No tax, contract minimums, rounding rules beyond two amount decimals, or licence entitlement decisions are inferred.",
        "units": sum(row["users"] for row in rows), "rows": rows,
    }
    if price:
        result.update(currency=store.settings.currency, amount=str(sum((Decimal(row["amount"]) for row in rows), Decimal("0.00"))))
    payload = json.dumps(result, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    result["content_sha256"] = hashlib.sha256(payload).hexdigest()
    return result


def csv_cell(value):
    """Neutralize spreadsheet formula injection, including leading whitespace."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return value
    value = str(value)
    if value.lstrip(" \t\r\n").startswith(("=", "+", "-", "@")) or value.startswith(("\t", "\r", "\n")):
        return "'" + value
    return value


def csv_text(rows: list[dict], columns: list[str]) -> str:
    output = io.StringIO(newline="")
    writer = csv.writer(output, lineterminator="\r\n")
    writer.writerow(columns)
    for row in rows:
        writer.writerow([csv_cell(row.get(key, "")) for key in columns])
    return "\ufeff" + output.getvalue()


def report_csv(report: dict) -> str:
    rows = [dict(row, generated_utc=report["generated_utc"], content_sha256=report["content_sha256"]) for row in report["rows"]]
    columns = ["month", "network_key", "network_scope", "product_id", "users", "word_users", "excel_users", "outlook_users",
               "users_with_daily_evidence", "monthly_only_users", "observed_user_days",
               "revision", "changed_utc", "first_received_utc", "last_received_utc", "last_use_utc", "coverage", "generated_utc", "content_sha256"]
    if "amount" in report:
        columns += ["unit_price", "currency", "amount"]
    return csv_text(rows, columns)


def users_csv(store: Store, month: str) -> str:
    rows = store.user_rows(valid_month(month))
    return csv_text(rows, ["month", "network_key", "network_scope", "product_id", "user_key", "hosts", "first_use_utc",
                           "loaded_days", "loaded_dates", "first_received_utc", "last_received_utc", "last_use_utc"])


def daily_csv(store: Store, month: str) -> str:
    month = valid_month(month)
    rows = [dict(row, date_basis="client-local add-in load date", coverage="observed; late offline reports may add evidence")
            for row in store.daily_rows(month + "-01", month + "-31")]
    return csv_text(rows, ["date", "network_key", "network_scope", "product_id", "users", "date_basis",
                           "first_received_utc", "last_received_utc", "coverage"])
