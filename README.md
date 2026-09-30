# Opening Balance Calculator

A simple page that helps clients work out their suspense account opening balance for an accounting go-live. When they click **Calculate Opening Balances**, the figures are saved as a new row in the **Opening Balance Call** tab of the Accounting Onboarding spreadsheet.

## What's in this repo

- `index.html`: the calculator page (this is all clients see)
- `OpeningBalance.gs`: the Apps Script code that saves each calculation to the spreadsheet. This is not used by the website itself, it is kept here as a backup and is pasted into Apps Script.

## How it works

1. The client fills in the form. Client Name, both dates and Client Account Closing Balance are required.
2. They click **Calculate Opening Balances**.
3. The results show on screen and the figures are sent to the spreadsheet.
4. A message confirms the save worked (or offers a **Try again** button if it did not).

Each click of Calculate adds a new row. Changing figures after calculating updates the results on screen only; they need to click Calculate again to save the new figures.

## Setting up the spreadsheet side

1. Open the Accounting Onboarding spreadsheet, then **Extensions > Apps Script**.
2. Add a new file called `OpeningBalance` and paste in the contents of `OpeningBalance.gs`. Leave `Code.gs` as it is.
3. **Deploy > Manage deployments**, click the pencil on the existing deployment, choose **New version**, then **Deploy**. The web app URL stays the same.
4. Approve the permissions if Google asks.

The "Opening Balance Call" tab is created automatically the first time someone saves a calculation.

## Publishing the page (GitHub Pages)

1. Create a new repository on GitHub and upload these files.
2. Go to **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, pick `main` and `/ (root)`, then **Save**.
4. After a minute or so GitHub gives you the live link.

## Changing things later

- **New web app URL:** update `WEB_APP_URL` near the top of the script section in `index.html`.
- **Different tab name:** change `OB_SHEET_NAME` in `OpeningBalance.gs` (and redeploy).
