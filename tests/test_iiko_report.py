import os
import re
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import iiko_report as report


class KpiStatusTests(unittest.TestCase):
    def test_barista_boundaries(self):
        cases = [
            (0, None, 1200),
            (1199, None, 1200),
            (1200, 1200, 1350),
            (1300, 1200, 1350),
            (1349, 1200, 1350),
            (1350, 1350, 1600),
            (1599, 1350, 1600),
            (1600, 1600, None),
        ]
        for value, expected_current, expected_next in cases:
            with self.subTest(value=value):
                current, next_level = report.get_kpi_status(value, "barista")
                self.assertEqual(
                    current["threshold"] if current else None,
                    expected_current,
                )
                self.assertEqual(
                    next_level["threshold"] if next_level else None,
                    expected_next,
                )

    def test_unknown_role_has_no_levels(self):
        self.assertEqual(report.get_kpi_status(9999, "other"), (None, None))


class CategoryFormulaTests(unittest.TestCase):
    def test_dops_sum_all_four_components(self):
        employee = {
            "id": "employee-1",
            "name": "Тестовый Сотрудник",
            "role": "barista",
            "worked_today": True,
            "worked_hours": 8.0,
            "avg_check": 1000.0,
            "revenue": 10000.0,
            "orders_count": 10,
            "dop_bar": 100.0,
            "dop_kitchen": 200.0,
            "dop_water": 300.0,
            "dop_alcohol": 400.0,
        }
        month_data = report.update_month_data({}, [employee], date(2026, 10, 10))
        day = month_data["employee-1"]["days"][0]
        self.assertEqual(day["dops"], 1000.0)
        self.assertEqual(
            report.calc_monthly_category_pct([day], "dops"),
            0.1,
        )

    def test_build_stats_maps_all_four_iiko_categories(self):
        with (
            mock.patch.object(report, "BARISTA_NAMES", ["Тестовый Сотрудник"]),
            mock.patch.object(report, "WAITER_NAMES", []),
        ):
            stats = report.build_stats(
                employees=[],
                hours_data=None,
                sales_data={"data": [{
                    "WaiterName": "Тестовый Сотрудник",
                    "DishDiscountSumInt": 10000,
                    "UniqOrderId": 10,
                    "DishDiscountSumInt.average": 1000,
                }]},
                category_data={"Тестовый Сотрудник": {
                    "Допы Бар": 100,
                    "Допы Кухня": 200,
                    "Бутилированные напитки": 300,
                    "Алкоголь": 400,
                }},
            )
        self.assertEqual(stats[0]["dop_bar"], 100)
        self.assertEqual(stats[0]["dop_kitchen"], 200)
        self.assertEqual(stats[0]["dop_water"], 300)
        self.assertEqual(stats[0]["dop_alcohol"], 400)


class OrganizationFilterTests(unittest.TestCase):
    def test_month_history_key_is_scoped_to_branch(self):
        with (
            mock.patch.object(report, "ORG_NAME", "Волоколамское шоссе 95/2с8"),
            mock.patch.object(report, "ORG_ID", ""),
        ):
            key = report.get_month_key(date(2026, 10, 10))
        self.assertRegex(key, r"^2026-10_[0-9a-f]{16}$")
        self.assertNotEqual(key, "2026-10")

    def test_resolves_exact_active_department_by_name(self):
        organizations = [
            {"id": "wrong", "name": report.ORG_NAME, "type": "CORPORATION", "active": True},
            {"id": "closed", "name": report.ORG_NAME, "type": "DEPARTMENT", "active": False},
            {"id": "target", "name": report.ORG_NAME, "type": "DEPARTMENT", "active": True},
        ]
        with mock.patch.object(report, "ORG_ID", ""):
            self.assertEqual(report.resolve_organization(organizations)["id"], "target")

    def test_ambiguous_department_fails_closed(self):
        organizations = [
            {"id": "one", "name": report.ORG_NAME, "type": "DEPARTMENT", "active": True},
            {"id": "two", "name": report.ORG_NAME, "type": "DEPARTMENT", "active": True},
        ]
        with mock.patch.object(report, "ORG_ID", ""):
            with self.assertRaisesRegex(RuntimeError, "неоднозначно"):
                report.resolve_organization(organizations)

    def test_sales_filter_uses_department_uuid(self):
        filters = report._sales_filters(
            "department-uuid",
            "2026-10-10T00:00:00",
            "2026-10-10T23:59:59",
        )
        self.assertEqual(
            filters[report.OLAP_DEPARTMENT_FIELD],
            {"filterType": "IncludeValues", "values": ["department-uuid"]},
        )


class SecurityTests(unittest.TestCase):
    def test_source_has_no_literal_passwords_or_disabled_tls(self):
        source = Path(report.__file__).read_text(encoding="utf-8")
        forbidden = [
            r'API_PASSWORD\s*=\s*["\'][^"\']+["\']',
            r'PG_PASS\s*=\s*["\'][^"\']+["\']',
            r"CERT_NONE",
            r"check_hostname\s*=\s*False",
            r"verify\s*=\s*False",
            r"token\s*\[:\s*8\s*\]",
        ]
        for pattern in forbidden:
            with self.subTest(pattern=pattern):
                self.assertIsNone(re.search(pattern, source))

    def test_tls_verification_is_default(self):
        with mock.patch.object(report, "IIKO_CA_BUNDLE", ""):
            session = report._session()
        self.assertIs(session.verify, True)

    def test_missing_iiko_configuration_fails_before_network(self):
        with mock.patch.multiple(
            report,
            IIKO_SERVER="",
            API_LOGIN="",
            API_PASSWORD="",
            ORG_ID="",
        ):
            with self.assertRaisesRegex(RuntimeError, "IIKO_SERVER"):
                report.get_token()

    def test_employee_name_is_html_escaped(self):
        employee = {
            "id": "employee-1",
            "name": "<script>alert(1)</script>",
            "role": "barista",
            "worked_today": True,
            "worked_hours": 8.0,
            "open_time": "08:00",
            "close_time": "22:00",
            "revenue": 1200.0,
            "orders_count": 1,
            "avg_check": 1200.0,
            "planned_open": "08:00",
            "planned_close": "22:00",
            "dop_bar": 0.0,
            "dop_kitchen": 0.0,
            "desserts": 0.0,
        }
        month_data = {
            "employee-1": {
                "name": employee["name"],
                "role": "barista",
                "days": [{
                    "date": "2026-10-10",
                    "worked": True,
                    "hours": 8.0,
                    "open_time": "08:00",
                    "close_time": "22:00",
                    "revenue": 1200.0,
                    "orders": 1,
                    "avg_check": 1200.0,
                    "dops": 0.0,
                    "desserts": 0.0,
                }],
                "violations": [],
            }
        }
        output = report.generate_html(
            [employee],
            month_data,
            date(2026, 10, 10),
            "",
        )
        self.assertNotIn("<script>alert(1)</script>", output)
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", output)

    def test_partial_source_failure_does_not_write_history(self):
        cached = {
            "employee-1": {
                "name": "Тестовый Сотрудник",
                "role": "barista",
                "days": [],
                "violations": [],
            }
        }
        with (
            mock.patch.object(report, "load_month_data_pg", return_value=cached),
            mock.patch.object(report, "get_token", return_value="token"),
            mock.patch.object(report, "get_organizations", return_value=[]),
            mock.patch.object(report, "get_employees", return_value=[]),
            mock.patch.object(report, "get_sales", return_value=None),
            mock.patch.object(report, "get_shift_times", return_value=None),
            mock.patch.object(report, "get_category_sales", return_value=None),
            mock.patch.object(report, "load_attendance", return_value={}),
            mock.patch.object(report, "load_schedule", return_value={}),
            mock.patch.object(report, "save_month_data_pg") as save_mock,
        ):
            output = report.generate_report_html()
        save_mock.assert_not_called()
        self.assertIn("Тестовый Сотрудник", output)

    def test_each_required_olap_failure_fails_closed(self):
        cached = {
            "employee-1": {
                "name": "Тестовый Сотрудник",
                "role": "barista",
                "days": [],
                "violations": [],
            }
        }
        organization = {
            "id": "department-uuid",
            "name": report.ORG_NAME,
            "type": "DEPARTMENT",
            "active": True,
        }
        for failed_source in ("sales", "shift", "category"):
            with self.subTest(failed_source=failed_source):
                with (
                    mock.patch.object(report, "ORG_ID", ""),
                    mock.patch.object(report, "load_month_data_pg", return_value=cached),
                    mock.patch.object(report, "get_token", return_value="token"),
                    mock.patch.object(report, "get_organizations", return_value=[organization]),
                    mock.patch.object(report, "validate_olap_department_filter"),
                    mock.patch.object(report, "get_employees", return_value=[{"id": "employee-1"}]),
                    mock.patch.object(
                        report,
                        "get_sales",
                        return_value=None if failed_source == "sales" else {"data": []},
                    ),
                    mock.patch.object(report, "get_worked_hours", return_value=None),
                    mock.patch.object(
                        report,
                        "get_shift_times",
                        return_value=None if failed_source == "shift" else {"data": []},
                    ),
                    mock.patch.object(
                        report,
                        "get_category_sales",
                        return_value=None if failed_source == "category" else {},
                    ),
                    mock.patch.object(report, "load_attendance", return_value={}),
                    mock.patch.object(report, "load_schedule", return_value={}),
                    mock.patch.object(report, "build_stats") as build_mock,
                    mock.patch.object(report, "save_month_data_pg") as save_mock,
                ):
                    output = report.generate_report_html()
                build_mock.assert_not_called()
                save_mock.assert_not_called()
                self.assertIn("Тестовый Сотрудник", output)

    def test_malformed_olap_payload_is_rejected(self):
        self.assertFalse(report._olap_response_ok(None))
        self.assertFalse(report._olap_response_ok({}))
        self.assertFalse(report._olap_response_ok({"data": {}}))
        self.assertTrue(report._olap_response_ok({"data": []}))


class TimezoneTests(unittest.TestCase):
    def test_default_timezone_is_moscow(self):
        self.assertEqual(report.APP_TIMEZONE.key, os.environ.get("APP_TIMEZONE", "Europe/Moscow"))


if __name__ == "__main__":
    unittest.main()

