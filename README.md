# Digital A-Z Word Bank App Template

This is the clean source template for each student's individual App.

## Files to configure

### GitHub App

In `index.html`, replace only:

1. `PASTE_PUBLISHED_SHEET_URL_HERE`
2. `PASTE_SPREADSHEET_ID_HERE`
3. `PASTE_BRIDGE_WEB_APP_URL_HERE`

Do not change `DISCOVERY_URL`, `courseDays`, or `EXTRA_TABS`.

### Apps Script Bridge

Copy `setup/Bridge.gs` into the student's standalone Apps Script Bridge project. Replace only:

1. `PASTE_SPREADSHEET_ID_HERE`
2. `PASTE_GITHUB_PAGES_APP_URL_HERE`

Run `setupBridge`, deploy the Bridge as a web app, and put its `/exec` URL into `index.html`. After GitHub Pages is live, run `getPersonalAccessLink` to obtain the student's personal connection link.

Never put the access token directly in GitHub.

## Day access

Change `Admin Settings!B1` in the student's Google Sheet. Study and Add & Manage will then allow Day 1 through that number. The A-Z Word Bank, Everyday Phrasal Verbs, Business Phrasal Verbs, and Proverbs remain available for study.

## Publishing

Do not deploy this template repository itself. Create each student's repository with **Use this template**, configure the placeholders, and then enable GitHub Pages in that student's repository.
