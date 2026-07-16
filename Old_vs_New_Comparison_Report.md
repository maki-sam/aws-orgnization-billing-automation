# Billing Lambda — Old vs New Comparison Report

**Project:** CB Bank Billing Automation — Bug Fix
**Date:** 2026-07-15
**File:** `lambda_function.py`
**Goal of the change:** Report totals must match the billing values shown in the AWS Console (Billing → Bills page / Cost Explorer default view).

---

## 1. Executive Summary

The old Lambda was built to match the **invoice PDF**, which is **gross of credits** (credits arrive as separate credit memos). The AWS Console, however, shows costs **net of credits and refunds**. That single design decision was the main reason the report never matched the console. The new version defaults to the **console basis** (net of credits/refunds) and keeps the old invoice basis available behind a switch.

Two secondary defects also affected the numbers (rounding drift, silently dropped discount types), and one operational bug (hardcoded test dates) would have made every scheduled run report June 2026 forever.

---

## 2. Changes That Affect the Numbers

| # | Area | Old version | New version | Impact on total |
|---|------|-------------|-------------|-----------------|
| 1 | **Credits & refunds** | `Credit`/`Refund` record types excluded from all queries → report is **gross** | Included by default (`COST_BASIS=console`) as a visible **"Credits & Refunds" row** on each account sheet, inside the SubTotal → report is **net**, matching the console | Usually the **entire mismatch**: report was higher than console by the month's credit amount |
| 2 | **Rounding** | Every service amount written as `round(x, 2)` **before** Excel sums the rows | Full-precision values written; the `#,##0.00` number format handles display | Cent-level drift (grows with number of accounts × services) |
| 3 | **Unknown discount types** | Every record type containing "discount" was excluded from the main query, but only **SPP** and **Bundled** were added back — an EDP or Distributor discount would vanish from the report | Only the four record types shown on dedicated rows (Tax, Credit/Refund, SPP, Bundled) are pulled out; **everything else stays in the per-service data** | Zero today, but a silent future mismatch the moment AWS applies any other discount type |
| 4 | **Pagination** | `discover_record_types()` and `get_tax_and_discounts()` did **not** paginate (`NextPageToken` ignored) | All Cost Explorer calls go through one paginating helper `_ce_query()` | Potential truncated/missing rows once the org grows past one response page |

---

## 3. Operational Bug Fixed

| Area | Old version | New version |
|------|-------------|-------------|
| **Manual test dates** | `MANUAL_START = "2026-06-01"`, `MANUAL_END = "2026-07-01"` left set — every scheduled run would report **June 2026 forever**, regardless of the actual month | Both reset to `""` (normal auto mode: previous month) |

---

## 4. New Configuration

| Item | Old version | New version |
|------|-------------|-------------|
| `COST_BASIS` env var | — (did not exist) | `"console"` (default) = net of credits/refunds, matches AWS Console. `"invoice"` = old gross behaviour, matches invoice PDF. Invalid values raise an error instead of running with wrong numbers. |
| Event override | `{"start", "end"}` only | `{"start", "end", "cost_basis"}` — basis can be switched per invocation for backfills/reconciliation |

---

## 5. Code Structure Changes

| Old function | New function | What changed |
|--------------|--------------|--------------|
| inline pagination loop (only in `get_costs_by_account_service`) | `_ce_query(start, end, group_by, flt)` | One shared generator; every CE query now paginates |
| `discover_record_types()` | `discover_record_types()` | Same purpose, now paginated via `_ce_query` |
| — | `classify_record_types()` | New: splits the month's record types into Tax / Credit+Refund / SPP / Bundled buckets in one place |
| `get_tax_and_discounts()` → `(tax, spp, bundled)` | `get_special_records()` → `(tax, credits, spp, bundled)` | Also fetches per-account credits/refunds; paginated |
| `get_costs_by_account_service()` | `get_costs_by_account_service()` | Same logic, but excludes **only** the four reported-separately buckets instead of every "*discount*" type |
| `build_workbook(label, accounts, costs, tax, spp, bundled)` | `build_workbook(..., credits, ..., cost_basis)` | Two new parameters; writes full-precision values; adds Credits & Refunds row in console mode |
| `send_report(wb_bytes, label, total_hint)` | `send_report(..., cost_basis)` | Email body states which basis the total uses |
| unused imports `calendar`, `get_column_letter` | removed | Cleanup only |

---

## 6. Excel Report Layout Changes

**Per-account sheets (console mode only):**

```
OLD                          NEW (COST_BASIS=console)
---------------------        -------------------------
Service rows...              Service rows...
Tax                          Tax
                             Credits & Refunds   <- NEW row (negative amount)
SubTotal  (services+tax)     SubTotal  (services+tax+credits)
SPP                          SPP
Total                        Total               <- now equals the console
                                                    per-account value
```

- The Credits & Refunds row sits **inside** the SubTotal because the console's per-account figure is net of credits. All rollup sheets (Total Cost for All Accounts, Cost+SPP+Bundle Discount, All Total) pull from SubTotal/SPP anchors, so they inherit the fix automatically — their formulas are unchanged.
- With `COST_BASIS=invoice` the layout is **identical to the old report** (no credits row, gross numbers).
- Display note: because cells now hold full precision, rows displayed at 2 dp can *appear* one cent off from the displayed SubTotal. That is the trade-off for the totals matching the console exactly. (The old version had the reverse problem: rows visually added up, but the total was wrong vs the console.)

**Unchanged:** the 13-sheet structure, tab order, styling, sheet-name collision handling, account sorting (`ACCOUNT_SORT`), Bundled_Discount sheet behaviour.

---

## 7. Email & Handler Response Changes

**Email body:**

| Old | New |
|-----|-----|
| `Total (as invoiced, before credits): USD x` | `Total (net of credits/refunds, as shown in the AWS Console): USD x` — wording follows the active basis |

**Handler return value:**

```jsonc
// OLD
{"status": "sent", "file": "...", "period": "...", "total": 12345.67}

// NEW
{
  "status": "sent",
  "file": "...",
  "period": "...",
  "cost_basis": "console",
  "total": 12345.67,             // compare this to the AWS Console
  "credits_refunds": -230.45,    // the difference between the two bases
  "gross_before_credits": 12576.12  // compare this to the invoice PDF
}
```

The response now shows **both** totals plus the credit amount, so a console-vs-invoice discrepancy can be explained at a glance without re-running anything.

---

## 8. What Did NOT Change

- Metric: `UnblendedCost` (same metric the console uses) — unchanged.
- Tax still reported as its own row per account.
- SPP and Bundled discounts still reported on their own rows/sheets.
- Account discovery from `organizations:ListAccounts` (no hardcoded accounts).
- SES email delivery, attachment naming (`CB_Bank_<Mon-YYYY>_AWS_Cost_Report.xlsx`).
- Schedule assumption: run on the 2nd of each month for the previous month.
- IAM permissions: `ce:GetCostAndUsage`, `organizations:ListAccounts`, `ses:SendRawEmail`.

---

## 9. How to Verify the Fix

1. Deploy the new `lambda_function.py`.
2. Invoke with a **closed** month (the console's current month is an estimate):
   ```json
   {"start": "2026-06-01", "end": "2026-07-01"}
   ```
3. Compare the returned `total` with **Billing → Bills → June 2026** in the management account.
4. If comparing with Cost Explorer instead, ensure its filters **include credits and refunds** (that is the console default).
5. To reconcile against the invoice PDF, invoke with `{"cost_basis": "invoice", ...}` and compare `gross_before_credits`.
