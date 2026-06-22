# HubSpot Bulk Recipient Export (ZIP)

A Chrome extension that adds a one‑click **"Download selected as ZIP"** button to the
HubSpot Marketing Email list. Tick the email rows you want, name the ZIP, and it pulls
every recipient list as **XLSX** and packages them into a single ZIP — instead of the
~6 manual clicks per email (open → Export → Recipients list → change to XLSX → Export →
notification bell → download).

## What it does, step by step

1. Reads the rows you've ticked (HubSpot's normal checkboxes).
2. Looks up each email's internal export id in one batched request.
3. Asks HubSpot to generate an **XLSX** recipient list for each (the "Basic" list = sent/delivered roster, all events).
4. Watches the notifications for each file to finish generating.
5. Downloads each finished file and zips them under the name you typed.

No data leaves your machine except the normal calls to HubSpot that the website itself makes.

## Install (load unpacked)

1. Open Chrome and go to `chrome://extensions`.
2. Turn on **Developer mode** (top‑right toggle).
3. Click **Load unpacked**.
4. Select this folder: `C:\Users\yutianhuang\hubspot-bulk-recipient-export`
5. Make sure you're **logged in to HubSpot** in the same Chrome profile.

To update after any code change: come back to `chrome://extensions` and click the **reload ⟳**
icon on this extension's card.

## Use

1. Go to the marketing email list: `https://app.hubspot.com/email/<portal>/manage/state/sent`
   (or any `…/manage/…` view).
2. A panel appears at the bottom‑right: **Bulk recipient ZIP export**.
3. Tick the checkbox on each email row you want. The panel shows a live count.
4. Edit the **ZIP file name** if you like (defaults to `hubspot-recipient-lists-<date>-<time>`).
5. Click **Download selected as ZIP**.
6. Keep this tab open while files generate (it polls HubSpot's notifications for the finished files).
7. The ZIP downloads to your **Downloads** folder. Each file inside is named after its email.

Each `.xlsx` inside the ZIP is exactly the file HubSpot's own "Export recipients list → XLSX → Basic"
produces.

## Notes & troubleshooting

- **Keep the tab focused** while it works. Generation usually takes seconds but can take a minute or
  two for very large lists; the tool waits up to 5 minutes per batch.
- **"… not ready"** in the result means some files hadn't finished generating before the timeout —
  just run those again.
- **"not a valid XLSX"** for a file means HubSpot returned a login/error page instead of the file —
  reload the HubSpot tab (to refresh your session) and try again.
- Works on whatever rows are **currently visible and ticked** on the page. To grab more than one
  page, do one page at a time, or increase "per page" at the bottom of the HubSpot list.
- Selection uses HubSpot's native row checkboxes, so "select all" on the page works too.

## How it works (for maintainers)

- `content.js` (runs on the HubSpot page) does everything: builds the panel, reads ticked rows,
  resolves each email's `hs_root_mic_id` via one `crm-search` on object type `0-29`
  (filter `hs_origin_asset_id IN [...]`), fires `EmailExports/exportEventDigests` per email
  (prefixing `emailName` with a per‑run token — at the FRONT, because HubSpot truncates the name to
  60 chars — so it can match the right notification), then polls
  `notification-station/general/v3/notifications/list` for each finished export's `ctaProxyUrl`.
- It then fetches each CTA (same‑origin → session cookie is sent; it 302‑redirects to a
  CloudFront‑signed file on `*.hubspotusercontent-na1.net`, readable cross‑origin thanks to
  `host_permissions` + the `rules.json` declarativeNetRequest CORS header), builds a ZIP
  (store method, no external library), and downloads it via a blob URL.
- All internal HubSpot POSTs include the `X-HubSpot-CSRF-hubspotapi` header (value of the
  `csrf.app` cookie).

If HubSpot changes their internal API, the most likely things to update are the object type
(`0-29`), the property names (`hs_root_mic_id` / `hs_origin_asset_id`), or the export RPC path.
