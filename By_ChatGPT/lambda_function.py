from __future__ import annotations

"""
Single-file AWS monthly billing report Lambda (no CUR, no S3).

Dependencies supplied by Lambda runtime/layer:
- boto3 / botocore (normally included in the managed Python runtime)
- openpyxl (attach your existing Lambda layer)

Required environment variables:
- SES_SENDER
- SES_RECIPIENTS            comma-separated

Optional environment variables:
- SES_CC                    comma-separated
- SES_REGION                default: AWS_REGION or us-east-1
- COST_EXPLORER_REGION      default: us-east-1
- ORGANIZATIONS_REGION      default: us-east-1
- COST_METRIC               default: UnblendedCost
- CURRENCY_SYMBOL           default: $
- SPP_RECORD_TYPES          default: Solution Provider Program Discount
- BUNDLED_RECORD_TYPES      default: Bundled Discount,BundledDiscount
- FAIL_ON_ESTIMATED         default: true
- FAIL_ON_RECONCILIATION    default: true
- RECONCILIATION_TOLERANCE default: 0.02
- SES_RAW_EMAIL_MAX_BYTES   default: 9500000
- LOG_LEVEL                 default: INFO

Manual test event examples:
    {"report_month": "2026-06", "send_email": true}
    {"send_email": true}  # previous calendar month

The workbook is created in Lambda /tmp, attached directly through SES, and
deleted before the invocation completes. It is never uploaded to S3.
"""

import json
import logging
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from email import policy
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import boto3
from botocore.config import Config

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

LOGGER = logging.getLogger(__name__)

ZERO = Decimal("0")
CENT = Decimal("0.01")


def d(value: Any) -> Decimal:
    """Convert AWS numeric strings safely to Decimal."""
    if value in (None, ""):
        return ZERO
    return Decimal(str(value))


def q(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def normalize(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def parse_csv_env(name: str, default: Sequence[str]) -> tuple[str, ...]:
    raw = os.getenv(name)
    values = default if not raw else [part.strip() for part in raw.split(",")]
    return tuple(value for value in values if value)


def parse_bool_env(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass(frozen=True)
class ReportConfig:
    metric: str = "UnblendedCost"
    currency_symbol: str = "$"
    tolerance: Decimal = Decimal("0.02")
    fail_on_estimated: bool = True
    fail_on_reconciliation: bool = True
    spp_record_types: tuple[str, ...] = (
        "Solution Provider Program Discount",
    )
    bundled_record_types: tuple[str, ...] = (
        "Bundled Discount",
        "BundledDiscount",
    )
    credit_record_types: tuple[str, ...] = ("Credit",)
    refund_record_types: tuple[str, ...] = ("Refund",)
    tax_record_types: tuple[str, ...] = ("Tax",)
    savings_plan_record_types: tuple[str, ...] = (
        "SavingsPlanCoveredUsage",
        "Savings Plan Covered Usage",
        "SavingsPlanNegation",
        "Savings Plan Negation",
        "SavingsPlanRecurringFee",
        "Savings Plan Recurring Fee",
        "SavingsPlanUpfrontFee",
        "Savings Plan Upfront Fee",
    )
    generic_discount_record_types: tuple[str, ...] = ("Discount",)

    @classmethod
    def from_env(cls) -> "ReportConfig":
        return cls(
            metric=os.getenv("COST_METRIC", "UnblendedCost"),
            currency_symbol=os.getenv("CURRENCY_SYMBOL", "$"),
            tolerance=d(os.getenv("RECONCILIATION_TOLERANCE", "0.02")),
            fail_on_estimated=parse_bool_env("FAIL_ON_ESTIMATED", True),
            fail_on_reconciliation=parse_bool_env("FAIL_ON_RECONCILIATION", True),
            spp_record_types=parse_csv_env(
                "SPP_RECORD_TYPES", ("Solution Provider Program Discount",)
            ),
            bundled_record_types=parse_csv_env(
                "BUNDLED_RECORD_TYPES", ("Bundled Discount", "BundledDiscount")
            ),
            credit_record_types=parse_csv_env("CREDIT_RECORD_TYPES", ("Credit",)),
            refund_record_types=parse_csv_env("REFUND_RECORD_TYPES", ("Refund",)),
            tax_record_types=parse_csv_env("TAX_RECORD_TYPES", ("Tax",)),
            savings_plan_record_types=parse_csv_env(
                "SAVINGS_PLAN_RECORD_TYPES",
                (
                    "SavingsPlanCoveredUsage",
                    "Savings Plan Covered Usage",
                    "SavingsPlanNegation",
                    "Savings Plan Negation",
                    "SavingsPlanRecurringFee",
                    "Savings Plan Recurring Fee",
                    "SavingsPlanUpfrontFee",
                    "Savings Plan Upfront Fee",
                ),
            ),
            generic_discount_record_types=parse_csv_env(
                "GENERIC_DISCOUNT_RECORD_TYPES", ("Discount",)
            ),
        )


@dataclass
class Account:
    account_id: str
    name: str
    state: str = "ACTIVE"


@dataclass
class AccountBreakdown:
    account: Account
    record_types: dict[str, Decimal] = field(default_factory=dict)
    service_record_types: dict[str, dict[str, Decimal]] = field(default_factory=dict)
    direct_total: Decimal = ZERO
    unit: str = "USD"


@dataclass
class SavingsPlanDetail:
    savings_plan_arn: str
    attributes: dict[str, str]
    total_commitment: Decimal = ZERO
    used_commitment: Decimal = ZERO
    unused_commitment: Decimal = ZERO
    utilization_percentage: Decimal = ZERO
    amortized_recurring_commitment: Decimal = ZERO
    amortized_upfront_commitment: Decimal = ZERO
    total_amortized_commitment: Decimal = ZERO
    net_savings: Decimal = ZERO
    on_demand_cost_equivalent: Decimal = ZERO


@dataclass
class BillingDataset:
    start_date: str
    end_date: str
    month_label: str
    accounts: list[Account]
    account_breakdowns: dict[str, AccountBreakdown]
    discovered_record_types: list[str]
    record_type_totals: dict[str, Decimal]
    savings_plan_details: list[SavingsPlanDetail]
    estimated: bool = False
    warnings: list[str] = field(default_factory=list)


class RecordTypeClassifier:
    def __init__(self, config: ReportConfig):
        self.config = config
        self._sets = {
            "spp": {normalize(x) for x in config.spp_record_types},
            "bundled": {normalize(x) for x in config.bundled_record_types},
            "credit": {normalize(x) for x in config.credit_record_types},
            "refund": {normalize(x) for x in config.refund_record_types},
            "tax": {normalize(x) for x in config.tax_record_types},
            "savings_plans": {
                normalize(x) for x in config.savings_plan_record_types
            },
            "other_discount": {
                normalize(x) for x in config.generic_discount_record_types
            },
        }

    def category(self, record_type: str) -> str:
        key = normalize(record_type)
        for category, values in self._sets.items():
            if key in values:
                return category
        return "base"

    def configured_types(self, category: str) -> tuple[str, ...]:
        return {
            "spp": self.config.spp_record_types,
            "bundled": self.config.bundled_record_types,
            "credit": self.config.credit_record_types,
            "refund": self.config.refund_record_types,
            "tax": self.config.tax_record_types,
            "savings_plans": self.config.savings_plan_record_types,
            "other_discount": self.config.generic_discount_record_types,
        }[category]


class CostExplorerCollector:
    def __init__(self, ce_client: Any, organizations_client: Any, config: ReportConfig):
        self.ce = ce_client
        self.organizations = organizations_client
        self.config = config

    def list_active_accounts(self) -> list[Account]:
        accounts: list[Account] = []
        next_token: str | None = None
        while True:
            request: dict[str, Any] = {"MaxResults": 20}
            if next_token:
                request["NextToken"] = next_token
            response = self.organizations.list_accounts(**request)
            for item in response.get("Accounts", []):
                state = item.get("State") or item.get("Status") or "UNKNOWN"
                if state == "ACTIVE":
                    accounts.append(
                        Account(
                            account_id=item["Id"],
                            name=item.get("Name") or item["Id"],
                            state=state,
                        )
                    )
            next_token = response.get("NextToken")
            if not next_token:
                break
        return sorted(accounts, key=lambda x: (x.name.lower(), x.account_id))

    def discover_record_types(self, start_date: str, end_date: str) -> list[str]:
        values: list[str] = []
        next_token: str | None = None
        while True:
            request: dict[str, Any] = {
                "TimePeriod": {"Start": start_date, "End": end_date},
                "Dimension": "RECORD_TYPE",
                "Context": "COST_AND_USAGE",
                "MaxResults": 1000,
            }
            if next_token:
                request["NextPageToken"] = next_token
            response = self.ce.get_dimension_values(**request)
            values.extend(
                item["Value"]
                for item in response.get("DimensionValues", [])
                if item.get("Value")
            )
            next_token = response.get("NextPageToken")
            if not next_token:
                break
        return sorted(set(values), key=str.lower)

    def _get_cost_and_usage_pages(self, **request: Any) -> Iterable[dict[str, Any]]:
        next_token: str | None = None
        while True:
            page_request = dict(request)
            if next_token:
                page_request["NextPageToken"] = next_token
            response = self.ce.get_cost_and_usage(**page_request)
            yield response
            next_token = response.get("NextPageToken")
            if not next_token:
                break

    def get_account_direct_totals(
        self, start_date: str, end_date: str
    ) -> tuple[dict[str, Decimal], dict[str, str], bool]:
        totals: dict[str, Decimal] = defaultdict(lambda: ZERO)
        units: dict[str, str] = {}
        estimated = False
        request = {
            "TimePeriod": {"Start": start_date, "End": end_date},
            "Granularity": "MONTHLY",
            "Metrics": [self.config.metric],
            "GroupBy": [{"Type": "DIMENSION", "Key": "LINKED_ACCOUNT"}],
        }
        for response in self._get_cost_and_usage_pages(**request):
            for period in response.get("ResultsByTime", []):
                estimated = estimated or bool(period.get("Estimated"))
                for group in period.get("Groups", []):
                    if not group.get("Keys"):
                        continue
                    account_id = group["Keys"][0]
                    metric = group.get("Metrics", {}).get(self.config.metric, {})
                    totals[account_id] += d(metric.get("Amount"))
                    if metric.get("Unit"):
                        units[account_id] = metric["Unit"]
        return dict(totals), units, estimated

    def get_account_record_types(
        self, start_date: str, end_date: str
    ) -> tuple[dict[str, dict[str, Decimal]], dict[str, Decimal], bool]:
        by_account: dict[str, dict[str, Decimal]] = defaultdict(
            lambda: defaultdict(lambda: ZERO)
        )
        org_totals: dict[str, Decimal] = defaultdict(lambda: ZERO)
        estimated = False
        request = {
            "TimePeriod": {"Start": start_date, "End": end_date},
            "Granularity": "MONTHLY",
            "Metrics": [self.config.metric],
            "GroupBy": [
                {"Type": "DIMENSION", "Key": "LINKED_ACCOUNT"},
                {"Type": "DIMENSION", "Key": "RECORD_TYPE"},
            ],
        }
        for response in self._get_cost_and_usage_pages(**request):
            for period in response.get("ResultsByTime", []):
                estimated = estimated or bool(period.get("Estimated"))
                for group in period.get("Groups", []):
                    keys = group.get("Keys", [])
                    if len(keys) < 2:
                        continue
                    account_id, record_type = keys[0], keys[1]
                    amount = d(
                        group.get("Metrics", {})
                        .get(self.config.metric, {})
                        .get("Amount")
                    )
                    by_account[account_id][record_type] += amount
                    org_totals[record_type] += amount
        return (
            {key: dict(value) for key, value in by_account.items()},
            dict(org_totals),
            estimated,
        )

    def get_service_record_types_for_account(
        self, account_id: str, start_date: str, end_date: str
    ) -> tuple[dict[str, dict[str, Decimal]], bool]:
        by_service: dict[str, dict[str, Decimal]] = defaultdict(
            lambda: defaultdict(lambda: ZERO)
        )
        estimated = False
        request = {
            "TimePeriod": {"Start": start_date, "End": end_date},
            "Granularity": "MONTHLY",
            "Metrics": [self.config.metric],
            "Filter": {
                "Dimensions": {"Key": "LINKED_ACCOUNT", "Values": [account_id]}
            },
            "GroupBy": [
                {"Type": "DIMENSION", "Key": "SERVICE"},
                {"Type": "DIMENSION", "Key": "RECORD_TYPE"},
            ],
        }
        for response in self._get_cost_and_usage_pages(**request):
            for period in response.get("ResultsByTime", []):
                estimated = estimated or bool(period.get("Estimated"))
                for group in period.get("Groups", []):
                    keys = group.get("Keys", [])
                    if len(keys) < 2:
                        continue
                    service, record_type = keys[0], keys[1]
                    amount = d(
                        group.get("Metrics", {})
                        .get(self.config.metric, {})
                        .get("Amount")
                    )
                    by_service[service][record_type] += amount
        return {key: dict(value) for key, value in by_service.items()}, estimated

    def get_savings_plan_details(
        self, start_date: str, end_date: str
    ) -> list[SavingsPlanDetail]:
        details: list[SavingsPlanDetail] = []
        next_token: str | None = None
        while True:
            request: dict[str, Any] = {
                "TimePeriod": {"Start": start_date, "End": end_date},
                "DataType": [
                    "ATTRIBUTES",
                    "UTILIZATION",
                    "AMORTIZED_COMMITMENT",
                    "SAVINGS",
                ],
                "MaxResults": 100,
            }
            if next_token:
                request["NextToken"] = next_token
            try:
                response = self.ce.get_savings_plans_utilization_details(**request)
            except self.ce.exceptions.DataUnavailableException:
                LOGGER.warning("Savings Plans utilization data is unavailable")
                return []
            for item in response.get("SavingsPlansUtilizationDetails", []):
                utilization = item.get("Utilization", {})
                amortized = item.get("AmortizedCommitment", {})
                savings = item.get("Savings", {})
                details.append(
                    SavingsPlanDetail(
                        savings_plan_arn=item.get("SavingsPlanArn", ""),
                        attributes={
                            str(key): str(value)
                            for key, value in item.get("Attributes", {}).items()
                        },
                        total_commitment=d(utilization.get("TotalCommitment")),
                        used_commitment=d(utilization.get("UsedCommitment")),
                        unused_commitment=d(utilization.get("UnusedCommitment")),
                        utilization_percentage=d(
                            utilization.get("UtilizationPercentage")
                        ),
                        amortized_recurring_commitment=d(
                            amortized.get("AmortizedRecurringCommitment")
                        ),
                        amortized_upfront_commitment=d(
                            amortized.get("AmortizedUpfrontCommitment")
                        ),
                        total_amortized_commitment=d(
                            amortized.get("TotalAmortizedCommitment")
                        ),
                        net_savings=d(savings.get("NetSavings")),
                        on_demand_cost_equivalent=d(
                            savings.get("OnDemandCostEquivalent")
                        ),
                    )
                )
            next_token = response.get("NextToken")
            if not next_token:
                break
        return details

    def collect(self, start_date: str, end_date: str, month_label: str) -> BillingDataset:
        accounts = self.list_active_accounts()
        account_map = {account.account_id: account for account in accounts}
        discovered = self.discover_record_types(start_date, end_date)
        direct_totals, units, estimated_a = self.get_account_direct_totals(
            start_date, end_date
        )
        account_record_types, record_type_totals, estimated_b = (
            self.get_account_record_types(start_date, end_date)
        )

        # Include closed/moved accounts that still appear in the billing data.
        for account_id in sorted(set(direct_totals) | set(account_record_types)):
            if account_id not in account_map:
                account = Account(account_id=account_id, name=f"Account {account_id}")
                accounts.append(account)
                account_map[account_id] = account

        breakdowns: dict[str, AccountBreakdown] = {}
        estimated = estimated_a or estimated_b
        for account in sorted(accounts, key=lambda x: (x.name.lower(), x.account_id)):
            service_records, service_estimated = self.get_service_record_types_for_account(
                account.account_id, start_date, end_date
            )
            estimated = estimated or service_estimated
            breakdowns[account.account_id] = AccountBreakdown(
                account=account,
                record_types=account_record_types.get(account.account_id, {}),
                service_record_types=service_records,
                direct_total=direct_totals.get(account.account_id, ZERO),
                unit=units.get(account.account_id, "USD"),
            )

        warnings: list[str] = []
        classifier = RecordTypeClassifier(self.config)
        normalized_discovered = {normalize(item) for item in discovered}
        for category in ("spp", "bundled"):
            configured = classifier.configured_types(category)
            if configured and not any(
                normalize(item) in normalized_discovered for item in configured
            ):
                warnings.append(
                    f"No configured {category.upper()} RECORD_TYPE was discovered. "
                    f"Configured values: {', '.join(configured)}. Review the Record Types sheet."
                )

        generic_discount_total = sum(
            (
                amount
                for record_type, amount in record_type_totals.items()
                if classifier.category(record_type) == "other_discount"
            ),
            ZERO,
        )
        if generic_discount_total != ZERO:
            warnings.append(
                "A generic Discount RECORD_TYPE has a non-zero amount. Cost Explorer "
                "cannot prove whether it is SPP, bundled, or another discount; it is "
                "reported as Other Discount."
            )

        savings_plan_details = self.get_savings_plan_details(start_date, end_date)
        return BillingDataset(
            start_date=start_date,
            end_date=end_date,
            month_label=month_label,
            accounts=sorted(accounts, key=lambda x: (x.name.lower(), x.account_id)),
            account_breakdowns=breakdowns,
            discovered_record_types=discovered,
            record_type_totals=record_type_totals,
            savings_plan_details=savings_plan_details,
            estimated=estimated,
            warnings=warnings,
        )


class ExcelReportBuilder:
    TITLE_FILL = PatternFill("solid", fgColor="FFFFFF")
    HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
    SUBHEADER_FILL = PatternFill("solid", fgColor="E2F0D9")
    TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
    WARNING_FILL = PatternFill("solid", fgColor="FFF2CC")
    ERROR_FILL = PatternFill("solid", fgColor="F4CCCC")
    THIN_GRAY = Side(style="thin", color="B7B7B7")
    TOP_BORDER = Border(top=Side(style="thin", color="000000"))

    def __init__(self, config: ReportConfig):
        self.config = config
        self.classifier = RecordTypeClassifier(config)
        self.currency_format = (
            f'{config.currency_symbol}#,##0.00;[Red]('
            f'{config.currency_symbol}#,##0.00);-'
        )
        self.percent_format = "0.00%"

    def _load_template(self, template_path: str | None) -> Workbook:
        if template_path and Path(template_path).exists():
            return load_workbook(template_path)
        return Workbook()

    @staticmethod
    def _clear_sheet(ws: Any) -> None:
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
        ws.merged_cells.ranges = []
        ws.freeze_panes = None
        ws.auto_filter.ref = None
        for table_name in list(ws.tables.keys()):
            del ws.tables[table_name]

    @staticmethod
    def _unique_sheet_name(wb: Workbook, desired: str) -> str:
        invalid = r"[\\/*?:\[\]]"
        base = re.sub(invalid, "-", desired).strip() or "Account"
        base = base[:31]
        candidate = base
        index = 2
        while candidate in wb.sheetnames:
            suffix = f"-{index}"
            candidate = f"{base[:31-len(suffix)]}{suffix}"
            index += 1
        return candidate

    def _get_or_create(self, wb: Workbook, name: str) -> Any:
        if name in wb.sheetnames:
            ws = wb[name]
            self._clear_sheet(ws)
            return ws
        return wb.create_sheet(name)

    def _style_title(self, ws: Any, title: str, end_column: int) -> None:
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=end_column)
        cell = ws.cell(2, 2, title)
        cell.font = Font(name="Calibri", size=16, bold=True)
        cell.alignment = Alignment(horizontal="left")

    def _style_headers(self, ws: Any, row: int, start_col: int, headers: Sequence[str]) -> None:
        for offset, header in enumerate(headers):
            cell = ws.cell(row, start_col + offset, header)
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.fill = self.HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=self.THIN_GRAY)
        ws.row_dimensions[row].height = 22

    def _apply_currency(self, ws: Any, rows: Iterable[int], columns: Iterable[int]) -> None:
        for row in rows:
            for column in columns:
                ws.cell(row, column).number_format = self.currency_format
                ws.cell(row, column).alignment = Alignment(horizontal="right")

    @staticmethod
    def _set_widths(ws: Any, widths: Mapping[int, float]) -> None:
        for column, width in widths.items():
            ws.column_dimensions[get_column_letter(column)].width = width

    def _add_table(self, ws: Any, ref: str, name: str) -> None:
        if ref.endswith(":"):
            return
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

    def _account_components(self, breakdown: AccountBreakdown) -> dict[str, Decimal]:
        result = {
            "base": ZERO,
            "savings_plans": ZERO,
            "spp": ZERO,
            "bundled": ZERO,
            "credit": ZERO,
            "refund": ZERO,
            "tax": ZERO,
            "other_discount": ZERO,
        }
        for record_type, amount in breakdown.record_types.items():
            result[self.classifier.category(record_type)] += amount
        result["final"] = sum(breakdown.record_types.values(), ZERO)
        return result

    def _service_components(self, records: Mapping[str, Decimal]) -> dict[str, Decimal]:
        result = {
            "base": ZERO,
            "savings_plans": ZERO,
            "spp": ZERO,
            "bundled": ZERO,
            "credit": ZERO,
            "refund": ZERO,
            "tax": ZERO,
            "other_discount": ZERO,
        }
        for record_type, amount in records.items():
            result[self.classifier.category(record_type)] += amount
        result["final"] = sum(records.values(), ZERO)
        return result

    def _write_all_total(self, wb: Workbook, dataset: BillingDataset) -> None:
        ws = self._get_or_create(wb, "All Total")
        self._style_title(ws, f"{dataset.month_label} AWS Costs for all accounts", 5)
        headers = ["Account Name", "Account ID", "Final Cost", "Currency"]
        self._style_headers(ws, 3, 2, headers)
        row = 4
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            ws.cell(row, 2, account.name)
            ws.cell(row, 3, account.account_id)
            ws.cell(row, 4, float(breakdown.direct_total))
            ws.cell(row, 5, breakdown.unit)
            row += 1
        total_row = row
        ws.cell(total_row, 2, "Total").font = Font(bold=True)
        ws.cell(total_row, 4, f"=SUM(D4:D{total_row-1})")
        ws.cell(total_row, 4).font = Font(bold=True)
        for col in range(2, 6):
            ws.cell(total_row, col).fill = self.TOTAL_FILL
            ws.cell(total_row, col).border = self.TOP_BORDER
        self._apply_currency(ws, range(4, total_row + 1), [4])
        self._set_widths(ws, {2: 35, 3: 18, 4: 18, 5: 12})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:E{total_row-1}"

    def _write_components(self, wb: Workbook, dataset: BillingDataset) -> None:
        ws = self._get_or_create(wb, "Cost+SPP+Bundle Discount")
        self._style_title(ws, f"{dataset.month_label} AWS Cost Components", 12)
        headers = [
            "Account Name",
            "Account ID",
            "Base Cost",
            "Savings Plans",
            "SPP",
            "Bundled Discount",
            "Credits",
            "Refunds",
            "Tax",
            "Other Discount",
            "Final Cost",
        ]
        self._style_headers(ws, 3, 2, headers)
        row = 4
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            components = self._account_components(breakdown)
            values = [
                account.name,
                account.account_id,
                components["base"],
                components["savings_plans"],
                components["spp"],
                components["bundled"],
                components["credit"],
                components["refund"],
                components["tax"],
                components["other_discount"],
                breakdown.direct_total,
            ]
            for col, value in enumerate(values, 2):
                ws.cell(row, col, float(value) if isinstance(value, Decimal) else value)
            row += 1
        total_row = row
        ws.cell(total_row, 2, "Total").font = Font(bold=True)
        for col in range(4, 13):
            letter = get_column_letter(col)
            ws.cell(total_row, col, f"=SUM({letter}4:{letter}{total_row-1})")
            ws.cell(total_row, col).font = Font(bold=True)
        for col in range(2, 13):
            ws.cell(total_row, col).fill = self.TOTAL_FILL
            ws.cell(total_row, col).border = self.TOP_BORDER
        self._apply_currency(ws, range(4, total_row + 1), range(4, 13))
        self._set_widths(ws, {2: 35, 3: 18, **{col: 18 for col in range(4, 13)}})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:L{total_row-1}"

    def _write_simple_account_summary(
        self,
        wb: Workbook,
        dataset: BillingDataset,
        sheet_name: str,
        title: str,
        category: str,
        amount_header: str,
        include_magnitude: bool = False,
    ) -> None:
        ws = self._get_or_create(wb, sheet_name)
        end_col = 6 if include_magnitude else 5
        self._style_title(ws, title, end_col)
        headers = ["Account Name", "Account ID", amount_header, "Currency"]
        if include_magnitude:
            headers.insert(3, "Display Magnitude")
        self._style_headers(ws, 3, 2, headers)
        row = 4
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            amount = self._account_components(breakdown)[category]
            ws.cell(row, 2, account.name)
            ws.cell(row, 3, account.account_id)
            ws.cell(row, 4, float(amount))
            if include_magnitude:
                ws.cell(row, 5, float(abs(amount)))
                ws.cell(row, 6, breakdown.unit)
            else:
                ws.cell(row, 5, breakdown.unit)
            row += 1
        total_row = row
        ws.cell(total_row, 2, "Total").font = Font(bold=True)
        ws.cell(total_row, 4, f"=SUM(D4:D{total_row-1})")
        ws.cell(total_row, 4).font = Font(bold=True)
        if include_magnitude:
            ws.cell(total_row, 5, f"=SUM(E4:E{total_row-1})")
            ws.cell(total_row, 5).font = Font(bold=True)
        for col in range(2, end_col + 1):
            ws.cell(total_row, col).fill = self.TOTAL_FILL
            ws.cell(total_row, col).border = self.TOP_BORDER
        currency_cols = [4, 5] if include_magnitude else [4]
        self._apply_currency(ws, range(4, total_row + 1), currency_cols)
        self._set_widths(ws, {2: 35, 3: 18, 4: 20, 5: 20, 6: 12})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:{get_column_letter(end_col)}{total_row-1}"

    def _write_category_detail(
        self,
        wb: Workbook,
        dataset: BillingDataset,
        sheet_name: str,
        title: str,
        category: str,
    ) -> None:
        ws = self._get_or_create(wb, sheet_name)
        self._style_title(ws, title, 7)
        headers = [
            "Account Name",
            "Account ID",
            "Record Type",
            "Bill Impact",
            "Display Magnitude",
            "Currency",
        ]
        self._style_headers(ws, 3, 2, headers)
        row = 4
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            for record_type, amount in sorted(
                breakdown.record_types.items(), key=lambda item: item[0].lower()
            ):
                if self.classifier.category(record_type) != category:
                    continue
                ws.cell(row, 2, account.name)
                ws.cell(row, 3, account.account_id)
                ws.cell(row, 4, record_type)
                ws.cell(row, 5, float(amount))
                ws.cell(row, 6, float(abs(amount)))
                ws.cell(row, 7, breakdown.unit)
                row += 1
        if row == 4:
            ws.cell(4, 2, "No matching record types were returned for this period.")
            ws.merge_cells("B4:G4")
            ws["B4"].fill = self.WARNING_FILL
            row = 5
        total_row = row
        ws.cell(total_row, 2, "Total").font = Font(bold=True)
        ws.cell(total_row, 5, f"=SUM(E4:E{total_row-1})")
        ws.cell(total_row, 6, f"=SUM(F4:F{total_row-1})")
        for col in range(2, 8):
            ws.cell(total_row, col).fill = self.TOTAL_FILL
            ws.cell(total_row, col).border = self.TOP_BORDER
        self._apply_currency(ws, range(4, total_row + 1), [5, 6])
        self._set_widths(ws, {2: 35, 3: 18, 4: 36, 5: 20, 6: 20, 7: 12})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:G{max(3, total_row-1)}"

    def _write_savings_plan_utilization(self, wb: Workbook, dataset: BillingDataset) -> None:
        ws = self._get_or_create(wb, "SP Utilization")
        self._style_title(ws, f"{dataset.month_label} Savings Plans Utilization", 17)
        headers = [
            "Savings Plan ARN",
            "Linked Account",
            "SP Type",
            "Payment Option",
            "Region",
            "Instance Family",
            "Total Commitment",
            "Used Commitment",
            "Unused Commitment",
            "Utilization %",
            "Amortized Recurring",
            "Amortized Upfront",
            "Total Amortized",
            "Net Savings",
            "On-Demand Equivalent",
            "All Attributes",
        ]
        self._style_headers(ws, 3, 2, headers)
        row = 4
        for detail in dataset.savings_plan_details:
            attrs = {normalize(k): v for k, v in detail.attributes.items()}
            linked = (
                attrs.get("linkedaccount")
                or attrs.get("linkedaccountid")
                or attrs.get("accountid")
                or ""
            )
            sp_type = attrs.get("savingsplanstype") or attrs.get("savingsplantype") or ""
            payment = attrs.get("paymentoption") or ""
            region = attrs.get("region") or ""
            family = attrs.get("instancetypefamily") or attrs.get("instancefamily") or ""
            values = [
                detail.savings_plan_arn,
                linked,
                sp_type,
                payment,
                region,
                family,
                detail.total_commitment,
                detail.used_commitment,
                detail.unused_commitment,
                detail.utilization_percentage / Decimal("100"),
                detail.amortized_recurring_commitment,
                detail.amortized_upfront_commitment,
                detail.total_amortized_commitment,
                detail.net_savings,
                detail.on_demand_cost_equivalent,
                json.dumps(detail.attributes, sort_keys=True),
            ]
            for col, value in enumerate(values, 2):
                ws.cell(row, col, float(value) if isinstance(value, Decimal) else value)
            row += 1
        if row == 4:
            ws.cell(4, 2, "No Savings Plans utilization details returned.")
            ws.merge_cells("B4:Q4")
            row = 5
        self._apply_currency(ws, range(4, row), range(8, 11))
        self._apply_currency(ws, range(4, row), range(12, 17))
        for r in range(4, row):
            ws.cell(r, 11).number_format = self.percent_format
        self._set_widths(
            ws,
            {
                2: 52,
                3: 18,
                4: 20,
                5: 18,
                6: 15,
                7: 18,
                **{col: 18 for col in range(8, 17)},
                17: 50,
            },
        )
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:Q{max(3, row-1)}"


    def _write_metadata(self, wb: Workbook, dataset: BillingDataset) -> None:
        ws = self._get_or_create(wb, "Report Metadata")
        self._style_title(ws, f"{dataset.month_label} Report Metadata", 5)
        self._style_headers(ws, 3, 2, ["Setting", "Value", "Source / Notes", "Documentation"])
        rows = [
            (
                "Billing period",
                f"{dataset.start_date} to {dataset.end_date} (end exclusive)",
                "Calendar-month Cost Explorer query",
                "https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetCostAndUsage.html",
            ),
            (
                "Cost metric",
                self.config.metric,
                "Signed AWS billing amounts; discounts and credits are normally negative",
                "https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetCostAndUsage.html",
            ),
            (
                "Accounts source",
                "AWS Organizations ListAccounts",
                "Current active organization accounts plus any billed account IDs returned by Cost Explorer",
                "https://docs.aws.amazon.com/organizations/latest/APIReference/API_ListAccounts.html",
            ),
            (
                "Record-type source",
                "Cost Explorer GetDimensionValues / GetCostAndUsage",
                "Use the Record Types sheet to verify SPP and bundled mappings",
                "https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetDimensionValues.html",
            ),
            (
                "Savings Plans source",
                "GetSavingsPlansUtilizationDetails",
                "Savings Plan ARN utilization and savings details",
                "https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetSavingsPlansUtilizationDetails.html",
            ),
            (
                "SPP RECORD_TYPE mapping",
                ", ".join(self.config.spp_record_types),
                "Exact normalized matching only",
                "Review against your organization before production scheduling",
            ),
            (
                "Bundled RECORD_TYPE mapping",
                ", ".join(self.config.bundled_record_types),
                "Exact normalized matching only",
                "Review against your organization before production scheduling",
            ),
            (
                "Estimated",
                "Yes" if dataset.estimated else "No",
                "At least one Cost Explorer result was marked Estimated" if dataset.estimated else "All queried periods returned Estimated=false",
                "https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetCostAndUsage.html",
            ),
            (
                "Generated UTC",
                datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "Lambda execution time",
                "",
            ),
        ]
        for row_idx, values in enumerate(rows, 4):
            for col_idx, value in enumerate(values, 2):
                ws.cell(row_idx, col_idx, value)
                ws.cell(row_idx, col_idx).alignment = Alignment(vertical="top", wrap_text=True)
        self._set_widths(ws, {2: 28, 3: 48, 4: 60, 5: 85})
        ws.freeze_panes = "B4"

    def _write_record_types(self, wb: Workbook, dataset: BillingDataset) -> None:
        ws = self._get_or_create(wb, "Record Types")
        self._style_title(ws, f"{dataset.month_label} Cost Explorer Record Types", 6)
        headers = ["Record Type", "Mapped Category", "Organization Amount", "Configured/Discovered"]
        self._style_headers(ws, 3, 2, headers)
        all_types = sorted(
            set(dataset.discovered_record_types) | set(dataset.record_type_totals),
            key=str.lower,
        )
        row = 4
        for record_type in all_types:
            ws.cell(row, 2, record_type)
            ws.cell(row, 3, self.classifier.category(record_type))
            ws.cell(row, 4, float(dataset.record_type_totals.get(record_type, ZERO)))
            ws.cell(row, 5, "Discovered")
            row += 1
        configured: set[str] = set()
        for category in (
            "spp",
            "bundled",
            "credit",
            "refund",
            "tax",
            "savings_plans",
            "other_discount",
        ):
            for record_type in self.classifier.configured_types(category):
                if normalize(record_type) in {normalize(item) for item in all_types}:
                    continue
                key = normalize(record_type)
                if key in configured:
                    continue
                configured.add(key)
                ws.cell(row, 2, record_type)
                ws.cell(row, 3, category)
                ws.cell(row, 4, 0)
                ws.cell(row, 5, "Configured but not discovered")
                for col in range(2, 6):
                    ws.cell(row, col).fill = self.WARNING_FILL
                row += 1
        self._apply_currency(ws, range(4, row), [4])
        self._set_widths(ws, {2: 42, 3: 22, 4: 22, 5: 32})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:E{max(3, row-1)}"

    def _write_reconciliation(self, wb: Workbook, dataset: BillingDataset) -> list[str]:
        ws = self._get_or_create(wb, "Reconciliation")
        self._style_title(ws, f"{dataset.month_label} Reconciliation", 9)
        headers = [
            "Account Name",
            "Account ID",
            "Direct CE Total",
            "Record-Type Sum",
            "Service/Record Sum",
            "Direct vs Record Diff",
            "Direct vs Service Diff",
            "Status",
        ]
        self._style_headers(ws, 3, 2, headers)
        errors: list[str] = []
        row = 4
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            record_sum = sum(breakdown.record_types.values(), ZERO)
            service_sum = sum(
                (
                    amount
                    for records in breakdown.service_record_types.values()
                    for amount in records.values()
                ),
                ZERO,
            )
            record_diff = breakdown.direct_total - record_sum
            service_diff = breakdown.direct_total - service_sum
            ok = (
                abs(record_diff) <= self.config.tolerance
                and abs(service_diff) <= self.config.tolerance
            )
            status = "PASS" if ok else "FAIL"
            if not ok:
                errors.append(
                    f"{account.name} ({account.account_id}) reconciliation failed: "
                    f"direct-record={record_diff}, direct-service={service_diff}"
                )
            values = [
                account.name,
                account.account_id,
                breakdown.direct_total,
                record_sum,
                service_sum,
                record_diff,
                service_diff,
                status,
            ]
            for col, value in enumerate(values, 2):
                ws.cell(row, col, float(value) if isinstance(value, Decimal) else value)
            ws.cell(row, 9).fill = self.SUBHEADER_FILL if ok else self.ERROR_FILL
            row += 1
        total_row = row
        ws.cell(total_row, 2, "Organization Total").font = Font(bold=True)
        for col in range(4, 9):
            letter = get_column_letter(col)
            ws.cell(total_row, col, f"=SUM({letter}4:{letter}{total_row-1})")
            ws.cell(total_row, col).font = Font(bold=True)
        ws.cell(total_row, 9, "PASS" if not errors else "FAIL")
        ws.cell(total_row, 9).fill = self.SUBHEADER_FILL if not errors else self.ERROR_FILL
        status_fill = self.SUBHEADER_FILL if not errors else self.ERROR_FILL
        for col in range(2, 10):
            ws.cell(total_row, col).border = self.TOP_BORDER
            ws.cell(total_row, col).fill = status_fill if col == 9 else self.TOTAL_FILL
        self._apply_currency(ws, range(4, total_row + 1), range(4, 9))
        self._set_widths(ws, {2: 35, 3: 18, **{col: 22 for col in range(4, 9)}, 9: 12})
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"B3:I{total_row-1}"

        warning_row = total_row + 3
        ws.cell(warning_row, 2, "Report Warnings").font = Font(bold=True)
        ws.cell(warning_row, 2).fill = self.WARNING_FILL
        for warning in dataset.warnings:
            warning_row += 1
            ws.cell(warning_row, 2, warning)
            ws.merge_cells(start_row=warning_row, start_column=2, end_row=warning_row, end_column=9)
            ws.cell(warning_row, 2).alignment = Alignment(wrap_text=True)
            ws.cell(warning_row, 2).fill = self.WARNING_FILL
        if dataset.estimated:
            warning_row += 1
            ws.cell(warning_row, 2, "Cost Explorer marked at least one result as Estimated.")
            ws.merge_cells(start_row=warning_row, start_column=2, end_row=warning_row, end_column=9)
            ws.cell(warning_row, 2).fill = self.ERROR_FILL
        return errors

    def _write_account_sheets(self, wb: Workbook, dataset: BillingDataset) -> None:
        # Remove a placeholder sheet if present.
        if "Account-1" in wb.sheetnames:
            del wb["Account-1"]
        for account in dataset.accounts:
            breakdown = dataset.account_breakdowns[account.account_id]
            sheet_name = self._unique_sheet_name(wb, f"{account.name}-{account.account_id[-4:]}")
            ws = wb.create_sheet(sheet_name)
            self._style_title(
                ws,
                f"{dataset.month_label} - {account.name} ({account.account_id})",
                11,
            )
            headers = [
                "Service",
                "Base Cost",
                "Savings Plans",
                "SPP",
                "Bundled Discount",
                "Credits",
                "Refunds",
                "Tax",
                "Other Discount",
                "Final Service Cost",
            ]
            self._style_headers(ws, 3, 2, headers)
            row = 4
            for service, records in sorted(
                breakdown.service_record_types.items(), key=lambda item: item[0].lower()
            ):
                components = self._service_components(records)
                values = [
                    service,
                    components["base"],
                    components["savings_plans"],
                    components["spp"],
                    components["bundled"],
                    components["credit"],
                    components["refund"],
                    components["tax"],
                    components["other_discount"],
                    components["final"],
                ]
                for col, value in enumerate(values, 2):
                    ws.cell(row, col, float(value) if isinstance(value, Decimal) else value)
                row += 1
            if row == 4:
                ws.cell(row, 2, "No service-level costs returned for this period.")
                row += 1
            total_row = row
            ws.cell(total_row, 2, "Total").font = Font(bold=True)
            for col in range(3, 12):
                letter = get_column_letter(col)
                ws.cell(total_row, col, f"=SUM({letter}4:{letter}{total_row-1})")
                ws.cell(total_row, col).font = Font(bold=True)
            for col in range(2, 12):
                ws.cell(total_row, col).fill = self.TOTAL_FILL
                ws.cell(total_row, col).border = self.TOP_BORDER
            self._apply_currency(ws, range(4, total_row + 1), range(3, 12))
            self._set_widths(ws, {2: 45, **{col: 18 for col in range(3, 12)}})
            ws.freeze_panes = "B4"
            ws.auto_filter.ref = f"B3:K{max(3, total_row-1)}"

            record_start = total_row + 3
            ws.cell(record_start, 2, "Record Type Audit").font = Font(bold=True)
            ws.cell(record_start, 2).fill = self.SUBHEADER_FILL
            self._style_headers(
                ws,
                record_start + 1,
                2,
                ["Record Type", "Mapped Category", "Amount"],
            )
            audit_row = record_start + 2
            for record_type, amount in sorted(
                breakdown.record_types.items(), key=lambda item: item[0].lower()
            ):
                ws.cell(audit_row, 2, record_type)
                ws.cell(audit_row, 3, self.classifier.category(record_type))
                ws.cell(audit_row, 4, float(amount))
                audit_row += 1
            self._apply_currency(ws, range(record_start + 2, audit_row), [4])

    def build(
        self,
        dataset: BillingDataset,
        output_path: str,
        template_path: str | None = None,
    ) -> list[str]:
        wb = self._load_template(template_path)
        if not wb.sheetnames:
            wb.create_sheet("All Total")
        if wb.sheetnames == ["Sheet"]:
            wb["Sheet"].title = "All Total"

        self._write_all_total(wb, dataset)
        self._write_components(wb, dataset)
        self._write_simple_account_summary(
            wb,
            dataset,
            "Total Cost for All Accounts",
            f"{dataset.month_label} Base Cost for all accounts",
            "base",
            "Base Cost",
        )
        self._write_simple_account_summary(
            wb,
            dataset,
            "SPP for All Accounts",
            f"{dataset.month_label} Solution Provider Program Discounts",
            "spp",
            "SPP Bill Impact",
            include_magnitude=True,
        )
        self._write_simple_account_summary(
            wb,
            dataset,
            "Bundled_Discount",
            f"{dataset.month_label} Bundled Discounts",
            "bundled",
            "Bundled Bill Impact",
            include_magnitude=True,
        )
        self._write_category_detail(
            wb,
            dataset,
            "Savings Plans",
            f"{dataset.month_label} Savings Plans Billing Records",
            "savings_plans",
        )
        self._write_savings_plan_utilization(wb, dataset)
        self._write_category_detail(
            wb,
            dataset,
            "SPP Detail",
            f"{dataset.month_label} SPP Billing Records",
            "spp",
        )
        self._write_category_detail(
            wb,
            dataset,
            "Bundled Detail",
            f"{dataset.month_label} Bundled Discount Billing Records",
            "bundled",
        )
        self._write_category_detail(
            wb,
            dataset,
            "Credits",
            f"{dataset.month_label} Credits",
            "credit",
        )
        self._write_category_detail(
            wb,
            dataset,
            "Refunds",
            f"{dataset.month_label} Refunds",
            "refund",
        )
        self._write_category_detail(
            wb,
            dataset,
            "Other Discounts",
            f"{dataset.month_label} Unclassified Generic Discounts",
            "other_discount",
        )
        self._write_metadata(wb, dataset)
        self._write_record_types(wb, dataset)
        reconciliation_errors = self._write_reconciliation(wb, dataset)
        self._write_account_sheets(wb, dataset)

        # Order important sheets first.
        preferred_order = [
            "All Total",
            "Cost+SPP+Bundle Discount",
            "Total Cost for All Accounts",
            "Savings Plans",
            "SP Utilization",
            "SPP for All Accounts",
            "SPP Detail",
            "Bundled_Discount",
            "Bundled Detail",
            "Credits",
            "Refunds",
            "Other Discounts",
            "Report Metadata",
            "Record Types",
            "Reconciliation",
        ]
        ordered = [wb[name] for name in preferred_order if name in wb.sheetnames]
        ordered.extend(ws for ws in wb.worksheets if ws not in ordered)
        wb._sheets = ordered

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        return reconciliation_errors


def previous_month_period(today: date | None = None) -> tuple[str, str, str]:
    today = today or datetime.now(timezone.utc).date()
    current_start = date(today.year, today.month, 1)
    previous_end = current_start
    if current_start.month == 1:
        previous_start = date(current_start.year - 1, 12, 1)
    else:
        previous_start = date(current_start.year, current_start.month - 1, 1)
    return (
        previous_start.isoformat(),
        previous_end.isoformat(),
        previous_start.strftime("%b-%Y"),
    )


def period_from_month(month: str) -> tuple[str, str, str]:
    match = re.fullmatch(r"(\d{4})-(\d{2})", month)
    if not match:
        raise ValueError("report_month must use YYYY-MM format")
    year, month_num = int(match.group(1)), int(match.group(2))
    start = date(year, month_num, 1)
    if month_num == 12:
        end = date(year + 1, 1, 1)
    else:
        end = date(year, month_num + 1, 1)
    return start.isoformat(), end.isoformat(), start.strftime("%b-%Y")

# ---------------------------------------------------------------------------
# Lambda entry point and direct SES delivery
# ---------------------------------------------------------------------------

BOTO_CONFIG = Config(
    retries={"max_attempts": 10, "mode": "adaptive"},
    connect_timeout=10,
    read_timeout=120,
)

# SES v1 SendRawEmail accepts a maximum 10 MB raw MIME message, including
# attachment encoding and headers. The lower default provides safety margin.
SES_ABSOLUTE_MAX_BYTES = 10_000_000
SES_DEFAULT_SAFE_MAX_BYTES = 9_500_000


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def split_addresses(raw: str) -> list[str]:
    return [item.strip() for item in raw.replace(";", ",").split(",") if item.strip()]


def send_report_email(
    ses_client: Any,
    sender: str,
    recipients: list[str],
    cc: list[str],
    subject: str,
    body_html: str,
    report_path: str,
) -> dict[str, Any]:
    """Attach the generated workbook and send it directly with SES."""
    all_recipients = recipients + cc
    if not recipients:
        raise ValueError("At least one SES recipient is required")
    if len(all_recipients) > 50:
        raise ValueError("SES SendRawEmail supports at most 50 total recipients")

    try:
        configured_max = int(
            os.getenv("SES_RAW_EMAIL_MAX_BYTES", str(SES_DEFAULT_SAFE_MAX_BYTES))
        )
    except ValueError as exc:
        raise ValueError("SES_RAW_EMAIL_MAX_BYTES must be an integer") from exc
    if configured_max <= 0:
        raise ValueError("SES_RAW_EMAIL_MAX_BYTES must be greater than zero")
    max_raw_bytes = min(configured_max, SES_ABSOLUTE_MAX_BYTES)

    report = Path(report_path)
    if not report.exists():
        raise FileNotFoundError(f"Billing workbook does not exist: {report_path}")

    message = MIMEMultipart("mixed")
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    if cc:
        message["Cc"] = ", ".join(cc)

    alternative = MIMEMultipart("alternative")
    alternative.attach(
        MIMEText(
            "The AWS monthly billing workbook is attached to this email.",
            "plain",
            "utf-8",
        )
    )
    alternative.attach(MIMEText(body_html, "html", "utf-8"))
    message.attach(alternative)

    with report.open("rb") as report_file:
        attachment = MIMEApplication(
            report_file.read(),
            _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    attachment.add_header(
        "Content-Disposition",
        "attachment",
        filename=report.name,
    )
    message.attach(attachment)

    raw_message = message.as_bytes(policy=policy.SMTP)
    raw_size = len(raw_message)
    if raw_size > max_raw_bytes:
        raise ValueError(
            "The generated SES email is too large: "
            f"{raw_size:,} bytes; configured limit is {max_raw_bytes:,} bytes. "
            "No email was sent. Reduce the workbook size or split the report."
        )

    LOGGER.info(
        "Sending workbook through SES: filename=%s attachment_bytes=%s raw_bytes=%s",
        report.name,
        report.stat().st_size,
        raw_size,
    )
    response = ses_client.send_raw_email(
        Source=sender,
        Destinations=all_recipients,
        RawMessage={"Data": raw_message},
    )
    return {
        "MessageId": response.get("MessageId"),
        "RawMessageBytes": raw_size,
        "AttachmentBytes": report.stat().st_size,
    }


def lambda_handler(event: dict[str, Any] | None, context: Any) -> dict[str, Any]:
    """Create the previous-month report, or use event['report_month']=YYYY-MM."""
    event = event or {}
    config = ReportConfig.from_env()

    if event.get("report_month"):
        start_date, end_date, month_label = period_from_month(
            str(event["report_month"])
        )
    else:
        start_date, end_date, month_label = previous_month_period()

    ce_client = boto3.client(
        "ce",
        region_name=os.getenv("COST_EXPLORER_REGION", "us-east-1"),
        config=BOTO_CONFIG,
    )
    organizations_client = boto3.client(
        "organizations",
        region_name=os.getenv("ORGANIZATIONS_REGION", "us-east-1"),
        config=BOTO_CONFIG,
    )
    ses_client = boto3.client(
        "ses",
        region_name=os.getenv("SES_REGION", os.getenv("AWS_REGION", "us-east-1")),
        config=BOTO_CONFIG,
    )

    sender = required_env("SES_SENDER")
    recipients = split_addresses(required_env("SES_RECIPIENTS"))
    cc = split_addresses(os.getenv("SES_CC", ""))

    output_name = f"{month_label}_AWS_Cost_Report.xlsx"
    output_path = Path("/tmp") / output_name

    try:
        collector = CostExplorerCollector(ce_client, organizations_client, config)
        dataset = collector.collect(start_date, end_date, month_label)

        if (
            dataset.estimated
            and config.fail_on_estimated
            and not bool(event.get("force_estimated", False))
        ):
            raise RuntimeError(
                "Cost Explorer marked the billing data as Estimated. "
                "Wait until the month is finalized, or set force_estimated=true "
                "only when you deliberately want a provisional report."
            )

        # No external workbook template is required. The builder creates and
        # formats every worksheet from code.
        builder = ExcelReportBuilder(config)
        reconciliation_errors = builder.build(
            dataset=dataset,
            output_path=str(output_path),
            template_path=None,
        )

        if reconciliation_errors and config.fail_on_reconciliation:
            raise RuntimeError(
                "Reconciliation failed; no email was sent: "
                + " | ".join(reconciliation_errors[:10])
            )

        warning_prefix = "[WARNING] " if dataset.warnings else ""
        subject = f"{warning_prefix}{month_label} AWS Monthly Billing Report"
        body_html = f"""
        <html><body>
          <p>The AWS monthly billing workbook for <strong>{month_label}</strong>
             is attached.</p>
          <p>Billing period: {start_date} through {end_date}
             (the end date is exclusive).</p>
          <p>Accounts included: {len(dataset.accounts)}</p>
          <p>Cost Explorer estimated data: {'Yes' if dataset.estimated else 'No'}</p>
          <p>Warnings: {len(dataset.warnings)}. Review the
             <strong>Record Types</strong> and <strong>Reconciliation</strong>
             worksheets.</p>
          <p>The workbook was sent directly through Amazon SES and was not stored
             in Amazon S3.</p>
        </body></html>
        """

        send_email = bool(event.get("send_email", True))
        message_id: str | None = None
        raw_message_bytes: int | None = None
        attachment_bytes = output_path.stat().st_size

        if send_email:
            email_result = send_report_email(
                ses_client=ses_client,
                sender=sender,
                recipients=recipients,
                cc=cc,
                subject=subject,
                body_html=body_html,
                report_path=str(output_path),
            )
            message_id = email_result.get("MessageId")
            raw_message_bytes = email_result.get("RawMessageBytes")
            attachment_bytes = email_result.get("AttachmentBytes")
        else:
            LOGGER.warning(
                "send_email=false: workbook was generated in /tmp only and will be deleted"
            )

        result = {
            "status": "ok",
            "billing_period": {"start": start_date, "end": end_date},
            "month_label": month_label,
            "accounts": len(dataset.accounts),
            "estimated": dataset.estimated,
            "warnings": dataset.warnings,
            "reconciliation_errors": reconciliation_errors,
            "report_filename": output_name,
            "attachment_bytes": attachment_bytes,
            "raw_message_bytes": raw_message_bytes,
            "email_sent": send_email,
            "ses_message_id": message_id,
            "persistent_storage": False,
        }
        LOGGER.info("Billing report result: %s", json.dumps(result, default=str))
        return result
    finally:
        try:
            output_path.unlink(missing_ok=True)
        except OSError:
            LOGGER.exception("Could not delete temporary workbook: %s", output_path)
