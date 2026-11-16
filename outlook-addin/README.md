# Convergent Method - Microsoft 365 Outlook Signature Plugin

This directory contains the production-ready Microsoft 365 Outlook Add-in that automatically inserts the corporate HTML signature and disclaimer (`signatures/Mark Email Signature.html`) when users compose new emails, replies, or calendar meetings.

---

## Architecture Overview

* **Event-Based Activation**: Automatically triggers on `OnNewMessageCompose` and `OnNewAppointmentOrganizer` via Office.js `setSignatureAsync`.
* **Zero Client Maintenance**: Deployed centrally by an M365 Administrator. Works across Outlook on Windows (both classic and New Outlook), Outlook for Mac, and Outlook on the web.
* **Hosted via GitHub Pages**: All static assets and scripts are served directly from `https://convergentmethod.com/outlook-addin/` over HTTPS.

---

## Files Included

| File | Purpose |
| :--- | :--- |
| `manifest.xml` | The Microsoft 365 Add-in XML manifest defining capabilities, icons, URLs, and event triggers. |
| `launchevent.html` | Background HTML runtime loader for web, Mac, and New Outlook. |
| `launchevent.js` | The JavaScript event handler that detects the sender, builds the signature HTML, and calls `setSignatureAsync()`. |
| `index.html` | Interactive portal for admins to preview the signature, test formatting, and copy deployment links. |
| `assets/` | High-resolution square icons (16px, 32px, 64px, 80px, 128px) required by Microsoft 365. |

---

## Organization-Wide Deployment (Microsoft 365 Admin Center)

To deploy this add-in across your entire organization without requiring employees to install anything:

1. Go to the [Microsoft 365 Admin Center](https://admin.microsoft.com) and log in with your Admin credentials.
2. In the navigation sidebar, expand **Settings** and select **Integrated apps**.
3. Near the top of the page, click **Upload custom apps**.
4. In the drawer:
   * **App type**: Select **Office Add-in**.
   * Choose **Provide link to manifest file** and enter:
     ```
     https://convergentmethod.com/outlook-addin/manifest.xml
     ```
     *(Alternatively, you can upload `manifest.xml` directly from your local drive).*
5. Click **Next**.
6. On the **Assign users** page:
   * Select **Entire organization** (or assign to specific groups or users if conducting a test phase first).
7. Review the app permissions (`ReadWriteItem`) and click **Next**, then **Finish deployment**.

> **Note on Propagation Time:** Outlook on the web will display the signature within a few minutes after a browser refresh. Outlook for Windows / Mac desktop clients typically sync within 6 to 24 hours depending on the client update cycle.

---

## Customizing User Information

In `launchevent.js`, the add-in looks up user information using `Office.context.mailbox.userProfile`:

* **Exact Matches:** Pre-configured directory for Mark Stemmans and Amanda Stemmans with exact credentials, roles, direct phone numbers, and formatting.
* **New Employees / Dynamic Fallback:** If a new team member joins, the add-in dynamically extracts their `displayName` and `emailAddress` from their Microsoft 365 profile and applies the Convergent Method layout.
* **Enforce Mark's Signature Everywhere:** If you want all emails across the tenant to literally use Mark Stemmans' signature, simply toggle:
  ```javascript
  var FORCE_MARK_SIGNATURE = true;
  ```
  in `launchevent.js`.

---

## Alternative: Exchange Online Server-Side Disclaimer

If you prefer server-side stamping (applies when messages are sent across all email apps including iOS Mail and Android), use the template at [signatures/Exchange-Mail-Flow-Rule.html](file:///c:/Users/ethan/My%20Drive/Convergent%20Method/GitHub/ConvergentMethod.github.io/signatures/Exchange-Mail-Flow-Rule.html):

1. Go to the [Exchange admin center](https://admin.exchange.microsoft.com).
2. Go to **Mail flow** -> **Rules**.
3. Click **Add a rule** -> **Apply a disclaimer**.
4. Set condition: **The sender is located... Inside the organization** and **The recipient is located... Outside the organization**.
5. Set action: **Apply a disclaimer to the message... Append a disclaimer**.
6. Paste the HTML content from `signatures/Exchange-Mail-Flow-Rule.html`.
