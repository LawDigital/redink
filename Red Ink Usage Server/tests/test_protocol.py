import unittest
from urllib.parse import quote, urlencode
from redink_server.protocol import ProtocolError, normalize_event, parse_legacy_path, parse_legacy_query


def usage(**changes):
    event = {"v": 1, "t": "U", "month": "2026-09", "networkKey": "n0123456789abcdef", "networkScope": "EXAMPLE.ORG|SUB.EXAMPLE.ORG", "allowedNetworkIds": ["EXAMPLE.ORG", "SUB.EXAMPLE.ORG"], "productId": "PRO", "productIdSafe": "PRO", "hostId": "WD", "userKey": "u" + "1" * 32, "billingEventId": "b" + "2" * 32, "hostEventId": "e" + "3" * 32, "firstUseUtc": "2026-09-01T11:22:33.1234567Z", "anon": True}
    event.update(changes)
    return event


def compact(**changes):
    event = {"ri_lc_v": 1, "t": "U", "m": "2026-09", "n": "n0123456789abcdef", "ns": "EXAMPLE.ORG|SUB.EXAMPLE.ORG", "p": "PRO", "h": "WD", "u": "u" + "1" * 32, "b": "b" + "2" * 32, "e": "e" + "3" * 32, "a": 1}
    event.update(changes)
    return event


def path_event(**changes):
    event = compact(**changes)
    keys = ["m", "n", "ns", "p", "h", "u", "b", "e", "a"]
    return "/collect/RI_LC_v1/" + "/".join(k + "/" + quote(str(event[k]), safe="") for k in keys) + "/r/abcdef01"


class ProtocolTests(unittest.TestCase):
    def test_marker_and_compact_contracts_match(self):
        full, short = normalize_event(usage()), normalize_event(compact())
        for key in ("event_type", "month", "network_key", "network_scope", "allowed_network_ids", "product_id", "host_id", "user_key", "billing_event_id", "host_event_id", "anonymous"):
            self.assertEqual(full[key], short[key], key)

    def test_path_and_query_access_log_formats(self):
        a = parse_legacy_path('127.0.0.1 - - [30/Oct/2026] "GET ' + path_event() + ' HTTP/1.1" 200 2')
        b = parse_legacy_query('2026-10-30 10:00:00 GET /collect ' + urlencode(compact()) + ' 200')
        self.assertEqual(a["month"], "2026-09")
        self.assertEqual(a["billing_event_id"], b["billing_event_id"])
        self.assertEqual(a["allowed_network_ids"], ["EXAMPLE.ORG", "SUB.EXAMPLE.ORG"])

    def test_form_allowed_networks_are_split(self):
        row = parse_legacy_query(urlencode(compact(ani="EXAMPLE.ORG|SUB.EXAMPLE.ORG")))
        self.assertEqual(row["allowed_network_ids"], ["EXAMPLE.ORG", "SUB.EXAMPLE.ORG"])

    def test_invalid_path_url_encoding_is_protocol_error(self):
        with self.assertRaises(ProtocolError): parse_legacy_path(path_event(ns="PLACEHOLDER").replace("PLACEHOLDER", "%FF"))

    def test_sanitized_path_product_preserves_event_ids(self):
        full = normalize_event(usage(productId="PRO / 2026", productIdSafe="PRO___2026"))
        path = parse_legacy_path(path_event(p="PRO___2026"))
        self.assertEqual(path["billing_event_id"], full["billing_event_id"])
        self.assertNotEqual(path["product_id"], full["product_id"])

    def test_reject_incomplete_and_conflicting_aliases(self):
        for field in ("month", "networkKey", "productId", "hostId", "userKey", "billingEventId", "hostEventId"):
            value = usage(); value.pop(field)
            if field == "productId": value.pop("productIdSafe")
            with self.subTest(field=field), self.assertRaises(ProtocolError): normalize_event(value)
        with self.assertRaisesRegex(ProtocolError, "Conflicting"): normalize_event(usage(m="2025-01"))

    def test_reject_invalid_month_hash_version_boolean(self):
        for change in ({"month": "2026-13"}, {"networkKey": "foo"}, {"v": 3}, {"anon": "maybe"}, {"hostId": "BAD"}, {"billingEventId": "b1"}, {"month": "0000-01"}, {"productId": "a\r\nb"}):
            with self.subTest(change=change), self.assertRaises(ProtocolError): normalize_event(usage(**change))

    def test_legacy_timestamp_is_not_recent_activity(self):
        event = normalize_event(usage())
        self.assertEqual(event["last_use_utc"], "")
        self.assertEqual(event["reported_utc"], "")

    def test_periodic_monthly_marker_timestamps(self):
        record = usage(lastUseUtc="2026-10-30T10:11:12.1234567Z", reportedUtc="2026-10-30T11:11:12+01:00")
        event = normalize_event(record)
        self.assertEqual(event["version"], 1)
        self.assertEqual(event["event_type"], "U")
        self.assertEqual(event["billing_event_id"], normalize_event(usage())["billing_event_id"])
        self.assertEqual(event["last_use_utc"], "2026-10-30T10:11:12.123456Z")
        self.assertEqual(event["reported_utc"], "2026-10-30T10:11:12.000000Z")
        for change in ({"lastUseUtc": "2026-10-30T10:11:12"}, {"reportedUtc": "invalid"}, {"v": 2}, {"t": "H"}):
            with self.subTest(change=change), self.assertRaises(ProtocolError):
                normalize_event(dict(record, **change))

    def test_reporting_interval_bounds(self):
        self.assertEqual(normalize_event(usage())["report_interval_minutes"], 0)
        for interval in (0, 1, 5, 360, 10080):
            self.assertEqual(normalize_event(usage(reportIntervalMinutes=interval))["report_interval_minutes"], interval)
        for interval in (-1, 10081, True, "2.5"):
            with self.subTest(interval=interval), self.assertRaises(ProtocolError):
                normalize_event(usage(reportIntervalMinutes=interval))

    def test_obfuscation_not_output_identity(self):
        event = normalize_event(usage(anon=False, userObf="o1_YWxpY2U"))
        self.assertEqual(event["user_obf"], "o1_YWxpY2U")
        self.assertNotIn("alice", event.values())
        self.assertEqual(normalize_event(usage(userObf="o1_YWxpY2U"))["user_obf"], "")

    def test_error_query_matches_full_event(self):
        full = {"v": 1, "t": "E", "errorEventId": "x" + "a"*32, "errorCode": "local_marker_missing", "relatedMonth": "2026-09", "relatedFile": "RI_LC_x.json"}
        query = {"ri_lc_v": 1, "t": "E", "id": full["errorEventId"], "ec": full["errorCode"], "m": "2026-09", "f": full["relatedFile"]}
        self.assertEqual(normalize_event(full), normalize_event(query))

    def test_malformed_path_never_recovers_prefix(self):
        with self.assertRaises(ProtocolError): parse_legacy_path(path_event() + "/unexpected/x")
        with self.assertRaises(ProtocolError): parse_legacy_path("/RI_LC_v1/m/2026-09/n/short")
        self.assertIsNone(parse_legacy_path("/health"))

    def test_duplicate_query_values_rejected(self):
        with self.assertRaises(ProtocolError): parse_legacy_query(urlencode(compact()) + "&m=2026-08")


if __name__ == "__main__": unittest.main()
