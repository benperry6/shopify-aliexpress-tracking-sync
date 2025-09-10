👉 **Full article (context, problems, and solutions):** https://media.prostrike.io/auto-add-aliexpress-tracking-to-fulfilled-shopify-orders/

# Auto-add AliExpress Tracking to Fulfilled Shopify Orders (Google Apps Script)

> Sync AliExpress shipping emails → extract Ali order IDs & tracking → match to Shopify fulfilled orders → push tracking via GraphQL → log to Google Sheet, with 72-hour no-tracking alerts.

## What this does
- Scans Gmail for AliExpress **confirmation/shipping** emails.
- Normalizes addresses and matches them to Shopify orders (fulfilled).
- Pulls tracking via `aliexpress.ds.order.tracking.get` (TOP) with HMAC-SHA256.
- Updates Shopify via **GraphQL** `fulfillmentTrackingInfoUpdate` (multi-numbers, dedup).
- Writes Ali order IDs + tracking to **Google Sheet** (safe-write, protects formulas).
- Sends a **72h alert** if no tracking shows up after confirmation.

## Prereqs
- Google Workspace account (Apps Script + Gmail + Sheets)
- Shopify store(s) with a **private app** or custom app OAuth (scopes listed below)
- AliExpress Developer account (app key/secret + tokens)
- A Gmail inbox receiving AliExpress transactional emails

## Install (15 minutes)
1. **Spreadsheet**
   - Create a sheet named: `Consolidated - Balance Sheet`
   - Ensure headers exist (anywhere in first 5 rows):
     - `Order #ID`, `From store?`
     - The script will auto-create `AE ID(s)` and `Tracking ID(s)` if missing.

2. **Apps Script**
   - In Google Drive: `New → More → Apps Script`
   - Create a standalone project, paste `src/sync_from_emails.gs`.
   - Set your `CFG.SPREADSHEET_ID` and other **placeholders**.

3. **Shopify OAuth**
   - Deploy the script as a **Web app** (Project Settings → Script URL).
   - In Shopify Partners, set App URL + Redirect URL to the script URL.
   - Install the app per shop; tokens are stored in **ScriptProperties**.

4. **AliExpress tokens**
   - Use the provided helpers (`AliAuth_ExchangeCodeOnce` / `AliAuth_Refresh`) to populate tokens in **ScriptProperties**.
   - Do **not** hard-code secrets.

5. **Gmail label/filters (optional but recommended)**
   - Filter `from:transaction@notice.aliexpress.com subject:(commande)` for clarity.

6. **Run**
   - Execute `Sync_FromEmails_Main()` or set a time-based trigger (e.g., every hour).

## Shopify scopes
read_orders,read_all_orders,read_customers,
read_fulfillments,write_fulfillments,
read_merchant_managed_fulfillment_orders,write_…

## Configuration
Update the `CFG` block with your values (Spreadsheet ID, store mapping, alert recipients). Keep secrets in **ScriptProperties**.

## Notes & Gotchas
- Only **fulfilled** Shopify orders are fetched (GraphQL `query: "fulfillment_status:fulfilled ..."`)  
- Robust matching uses `startsWith`, Levenshtein ≤ 6, or Jaccard ≥ 0.85 on normalized addresses.  
- Safe-write blocks updates to columns containing formulas (prevents breaking `ARRAYFORMULA`).  
- AliExpress TOP uses **Asia/Shanghai** timestamp; signature is **HMAC-SHA256** with strict key sorting.

## License
MIT
