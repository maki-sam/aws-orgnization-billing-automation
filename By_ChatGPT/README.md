# AWS Organizations Monthly Billing Report — Lambda

A single-file AWS Lambda that builds a multi-sheet Excel workbook of the
previous month's AWS Organizations costs and emails it as an attachment via
Amazon SES. No Cost and Usage Report (CUR), no S3 — the workbook is generated
in `/tmp` and sent directly.

> **File:** [`lambda_function.py`](lambda_function.py) · **Handler:** `lambda_handler`

---

## What it does

1. Lists every **ACTIVE** account in the Organization (`organizations:ListAccounts`).
2. Pulls last month's costs from **Cost Explorer**:
   - per account × `RECORD_TYPE` (usage, tax, credits, SPP, bundled discount, Savings Plans …) — the **account total**;
   - per account × `SERVICE` × `USAGE_TYPE` — the **per-service breakdown**.
3. Classifies record types (SPP, bundled, credit, refund, tax, Savings Plans, other discount, base).
4. Builds an `.xlsx` workbook with [openpyxl](https://openpyxl.readthedocs.io/).
5. Runs a **reconciliation** safety check, then emails the workbook via `ses:SendRawEmail`.
6. Deletes the temp file from `/tmp` (always, in a `finally`).

Cost Explorer's **End date is exclusive** — for June 2026 the query period is
`Start=2026-06-01, End=2026-07-01`.

---

## Quick start

### 1. Prerequisites

| Requirement | Detail |
|---|---|
| **Runtime** | Python 3.9+ Lambda |
| **Layer** | `openpyxl` (and its dependency `et_xmlfile`) must be provided as a **Lambda layer** — they are *not* bundled in this file |
| **SES** | `SES_SENDER` must be a **verified** identity. In the SES sandbox, every recipient must be verified too, or request production access |
| **Account** | Deploy in the **management account** (or a delegated Cost Explorer/Organizations account) |

### 2. IAM permissions

```json
{
  "Version": "2012-10-17",
  "Statement": [
    { "Effect": "Allow",
      "Action": [
        "ce:GetCostAndUsage",
        "ce:GetDimensionValues",
        "organizations:ListAccounts",
        "ses:SendRawEmail"
      ],
      "Resource": "*" }
  ]
}
```

Plus the standard `AWSLambdaBasicExecutionRole` for CloudWatch Logs.

### 3. Minimum configuration

Set two environment variables and deploy:

```
SES_SENDER=reports@yourdomain.com
SES_RECIPIENTS=finance@yourbank.com,ops@yourbank.com
```

### 4. Test it (without emailing)

Invoke with this test event — it builds the workbook but **does not send**:

```json
{ "report_month": "2026-06", "send_email": false }
```

Then check the CloudWatch log line `Billing report result: {...}` for
`accounts`, `estimated`, `warnings`, and `reconciliation_errors`.

### 5. Schedule it

Trigger monthly with **EventBridge Scheduler** (e.g. 2nd of each month) — no
event payload is needed; it defaults to the **previous** calendar month.

```
cron(0 6 2 * ? *)   # 06:00 UTC on the 2nd
```

---

## Configuration

### Required

| Variable | Description |
|---|---|
| `SES_SENDER` | Verified SES "From" address |
| `SES_RECIPIENTS` | Comma/semicolon-separated "To" addresses |

### Recommended / optional

| Variable | Default | Description |
|---|---|---|
| `SES_REGION` | `AWS_REGION` or `us-east-1` | SES region |
| `SES_CC` | *(none)* | Comma-separated CC addresses |
| `COST_EXPLORER_REGION` | `us-east-1` | Cost Explorer is global but the endpoint lives here |
| `ORGANIZATIONS_REGION` | `us-east-1` | Organizations endpoint |
| `COMPONENT_METRIC` | `UnblendedCost` | Metric for account / record-type totals |
| `SERVICE_METRIC` | `NetUnblendedCost` | Metric for the per-service breakdown |
| `CURRENCY_SYMBOL` | `$` | Prefix in the Excel money format |
| `SPP_RECORD_TYPES` | `Solution Provider Program Discount` | Record types treated as SPP |
| `BUNDLED_RECORD_TYPES` | `Bundled Discount,BundledDiscount` | Record types treated as bundled discount |
| `CREDIT_RECORD_TYPES` | `Credit` | Credit record types |
| `REFUND_RECORD_TYPES` | `Refund` | Refund record types |
| `TAX_RECORD_TYPES` | `Tax` | Tax record types |
| `SAVINGS_PLAN_RECORD_TYPES` | *(SP covered usage / negation / fees)* | Savings Plans record types |
| `GENERIC_DISCOUNT_RECORD_TYPES` | `Discount` | Any other discount record type |
| `RECLASSIFY_DATA_TRANSFER` | `true` | Move data-transfer usage into a "Data Transfer" service row |
| `DATA_TRANSFER_USAGE_PATTERNS` | `DataTransfer,DataXfer` | Usage-type markers for data transfer |
| `SERVICE_NAME_MAP_JSON` | *(none)* | JSON object of `{ "CE service name": "Display name" }` overrides |
| `FAIL_ON_ESTIMATED` | `true` | Abort if Cost Explorer marks the period as Estimated |
| `FAIL_ON_RECONCILIATION` | `true` | Abort (don't email) if reconciliation fails |
| `RECONCILIATION_TOLERANCE` | `0.02` | Absolute reconciliation floor, in currency units |
| `RECONCILIATION_REL_TOLERANCE` | `0.005` | Relative reconciliation allowance (fraction of the account total) |
| `SES_RAW_EMAIL_MAX_BYTES` | `9500000` | Reject the email above this size (capped at the 10 MB SES hard limit) |
| `LOG_LEVEL` | `INFO` | Logger level |

### Event flags

| Field | Default | Effect |
|---|---|---|
| `report_month` | previous month | `"YYYY-MM"` to report a specific month |
| `send_email` | `true` | **An empty event `{}` sends the email.** Set `false` to build only |
| `force_estimated` | `false` | Build a provisional report even when the period is still Estimated |

---

## The workbook

Sheets, in order:

| # | Sheet | Columns | Notes |
|---|---|---|---|
| 1 | **All Total** | Account Name · Account ID · Total Cost | `Total Cost = Cost + SPP + Bundled` — see [Cost model](#cost-model) |
| 2 | **Cost+SPP+Bundle Discount** | Cost · SPP Charges · Bundled Charges | Cost = net total; SPP/Bundled shown **positive**; `Total Cost = Cost + SPP + Bundled` (merged cell) |
| 3 | **Total Cost for All Accounts** | Cost | Per-account **base** cost (usage only) |
| 4 | **SPP for All Accounts** | SPP Discounts | Positive magnitudes |
| 5 | **Bundled_Discount** | Bundled Discount | Positive magnitudes |
| 6+ | **One sheet per account** | Service · Cost | Service breakdown (`SERVICE`+`USAGE_TYPE`, `NetUnblendedCost`), with data-transfer reclassification and display-name mapping |

Every total is a live Excel formula (`=SUM(...)`), and the workbook is saved
with `fullCalcOnLoad` so totals recalc when opened.

---

## Cost model

Two Cost Explorer metrics are combined:

- **Account total** (`COMPONENT_METRIC`, default `UnblendedCost`): the sum of
  all record types for the account — this is the **net** amount AWS bills.
- **Service breakdown** (`SERVICE_METRIC`, default `NetUnblendedCost`): used
  for the per-account Service/Cost sheets.

On the **All Total** and **Cost+SPP+Bundle Discount** sheets:

```
Cost      = account net total (direct_total)
SPP       = SPP discount as a POSITIVE magnitude
Bundled   = bundled discount as a POSITIVE magnitude
Total Cost = Cost + SPP + Bundled     ← GROSS, i.e. before SPP/bundled discounts
```

> ⚠️ **Read this before sharing with finance.** The **Total Cost** column on
> the All Total sheet is the **gross** figure (net cost *plus* the discount
> amounts), **not** the net amount invoiced by AWS. The net bill is the per-
> account **Cost** column on the Cost+SPP+Bundle Discount sheet. Label
> accordingly for your audience.

Both `Cost` and the `Cost/SPP/Bundled` triple come from one source in code
(`ExcelReportBuilder._cost_spp_bundled`), so the two sheets always agree.

### Reconciliation

For each account the **net** account total is compared against the sum of its
displayed service costs (different CE metrics, so a small spend-proportional
gap is expected). The per-account tolerance is:

```
max(RECONCILIATION_TOLERANCE, RECONCILIATION_REL_TOLERANCE * |account total|)
```

If any account exceeds it and `FAIL_ON_RECONCILIATION=true`, the run raises and
**no email is sent** (the reason is in the CloudWatch logs).

---

## Response

`lambda_handler` returns (and logs) a summary:

```json
{
  "status": "ok",
  "billing_period": { "start": "2026-06-01", "end": "2026-07-01" },
  "month_label": "Jun-2026",
  "accounts": 12,
  "estimated": false,
  "warnings": [],
  "reconciliation_errors": [],
  "report_filename": "Jun-2026_AWS_Cost_Report.xlsx",
  "attachment_bytes": 24576,
  "raw_message_bytes": 33012,
  "email_sent": true,
  "ses_message_id": "0100018f...",
  "persistent_storage": false
}
```

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `Missing required environment variable: SES_SENDER` | Set `SES_SENDER` / `SES_RECIPIENTS` |
| `Unable to import module ... openpyxl` | The openpyxl **layer** isn't attached |
| Raises "marked this period as Estimated" | Month not finalized — wait, or pass `force_estimated: true` |
| "Reconciliation failed; email was not sent" | An account's net vs service sum exceeded tolerance — inspect the logged diffs; adjust `RECONCILIATION_*` if legitimate |
| "SES email is too large" | Attachment over the limit — raise `SES_RAW_EMAIL_MAX_BYTES` (≤ 10 MB) or reduce scope |
| Email not received | Sender/recipient not SES-verified, or still in the SES sandbox |
| `[WARNING]` subject prefix | A configured SPP/bundled record type wasn't found this month — check the logged record types |

---

## Notes for maintainers

- **Not runnable without AWS + the openpyxl layer.** After any change, do a
  backfill smoke test with `{ "report_month": "<closed month>", "send_email": false }`
  and confirm the log summary before re-enabling the schedule.
- **SES client uses no automatic retries** (`SES_BOTO_CONFIG`) because
  `SendRawEmail` is not idempotent — a retry after a timeout could send twice.
- All money math uses `Decimal`; values are converted to `float` only when
  written into cells.
- Sibling variant [`../lambda_function.py`](../lambda_function.py) is the
  console-matching (`COST_BASIS`) design; this file is the record-type design.
