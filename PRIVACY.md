# Privacy policy

For [Goalie](https://andrewmacheret.com/projects/goalie) and [Miler Lite](https://andrewmacheret.com/projects/miler-lite), which sign in with Google as the "Sheets Tracker" app. Last updated November 24, 2026.

## The short version

Both apps are web pages that read and write your own Google Sheet, straight from your browser. There is no server or database behind them. Your data is not collected, stored, sold or shared by the developer.

## What the apps do with your Google account

When you click Authorize, Google asks you to let the app see, edit, create and delete your Google Sheets spreadsheets. The apps use that access only to:

- read the settings tab of the tracker spreadsheet the page is set up for, to build the form
- add one row to that spreadsheet each time you click Submit

They don't open, read or change any other spreadsheet, and they don't see your Google password. Google gives the page a short-lived access token, which stays in the open browser tab, is only sent to Google's Sheets API, and is gone when you close the tab. Sign Out revokes it.

You can remove the apps' access at any time from your [Google account's third-party connections](https://myaccount.google.com/connections).

The apps' use and transfer of information received from Google APIs will adhere to the [Google API Services User Data Policy](https://developers.google.com/terms/api-services-user-data-policy), including the Limited Use requirements.

## What the apps store

Each app saves a copy of its spreadsheet's settings in your browser's local storage, on your device only, so the form shows up faster next time. For Goalie that's the questions and their choices. For Miler Lite it's the businesses, presets and latest odometer reading. Clearing your browser's site data for andrewmacheret.com deletes it.

## What the apps do not do

- No analytics, tracking, ads or cookies
- Your spreadsheet data is never sent to the developer or anyone besides Google
- No data is used for advertising or to train AI models

## Other services

The pages are hosted on Amazon Web Services, which keeps standard access logs (IP address, page address, browser type and time) for 30 days. The pages also load Bootstrap and jQuery from public CDNs, which see your IP address the same way. Google's handling of your data is covered by [Google's privacy policy](https://policies.google.com/privacy).

## Changes

If this policy changes, the new version will be posted here with a new date. The apps' source code is public in the [goalie](https://github.com/andrewmacheret/goalie) and [miler-lite](https://github.com/andrewmacheret/miler-lite) repos, so you can check what they do.

## Contact

Questions? Email [andrew.macheret@proton.me](mailto:andrew.macheret@proton.me).
