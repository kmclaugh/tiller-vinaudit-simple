# vinAudit (Simple) Automation

The tiller-vinaudit-simple script is a simple Google Apps Script to scrape
[vinAudit](https://www.vinaudit.com/vehicle-market-value-api) estimates of a car
values and insert them into the Balance History sheet of Tiller-enabled
spreadsheets. This workflow is especially useful for tracking net worth.

Visit [Tiller HQ](https://tillerhq.com) to learn more about Tiller.

## Adivsory

_This simple script is designed for intermediate users and includes only
lightweight error checking. We hope it meets your needs out of the box, but
further tweaks may be required to get it working in your environment. As a
one-off Tiller-Labs release, Tiller offers no warranties or support for this
solution._

## Setting Up Your Project

To configure your project, implement the following steps in Google Sheets:

1. Find your vehichle's
   [vin](https://www.txdmv.gov/motorists/how-to-find-the-vin)
2. Open the Tiller Google-Sheets spreadsheet you'd like to add your vehichle's
   value to.
3. Click on Tools -> Script editor to open the spreadsheet's bound scripts.
4. Copy the contents of
   [vinaudit.js](https://raw.githubusercontent.com/TillerHQ/tiller-vinaudit-simple/master/vinaudit.js)
   from this repo into the Google Sheets script editor.
5. Set the `vins` variable (at the start of the code) equal to your vins.
6. Save the script file and name it `carEstimator`.
7. Click on the `Select Function` dropdown in the control bar and select the
   `vinAuditInsert` function.
8. Click the run/play button.
9. Select `Review Permissions` and select the google accounts associated with
   your Tiller Account.
10. You'll get a warning saying "This app isn't verified". Select `Advanced` and
    `Go to carEstimator (unsafe)`.
11. Select Allow

If you've completed all the steps successfully, you should have a new entry in
your Balance History reflecting the viAudit estimate of the vehichle.

## Run Monthly

1. In the script editor select `Edit > Current Project's Triggers`.
2. Select event source = Time-driven.
3. Select your desired options.

## Taking This Solution Further...

Consider adding:

- An
  [onOpen()](https://developers.google.com/apps-script/guides/triggers/#onopene)
  function to create a menu item to execute an update
- [Triggers](https://developers.google.com/apps-script/guides/triggers/) to
  automate script execution
