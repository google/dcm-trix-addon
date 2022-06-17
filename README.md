**_Disclaimer: This is not an official Google product._**

# DCM Trix Addon

## Overview

DCM Trix Addon allows users to download reports from DoubleClick Campaign
Manager on a scheduled basis into a Google Sheets Spreadsheet. Features of this
addon are:

-   One time fetch of reports to spreadsheet.
-   Scheduled fetch of reports in background (Hourly, Daily, Weekly).

## Initial Setup

1.  Go to drive.google.com and create a new spreadsheet by going to `New >
    Google Sheets`.
1.  Copy the doc id from the URL. This is the value between the `d` and `edit` as shown in bold below:
    <pre>https://docs.google.com/spreadsheets/d/<b><u>1p3IHbvgz0vpibMcW47HfzYijl4sLnECKVYist3ByZXU</u></b>/edit#gid=0</pre>
1.  Give the spreadsheet a name (Ex: "DCM Reporting")

## Upload the script files

#### Upload the files in the script using command line tool - Option 1

If you are familiar with using command line tools, you can use clasp. Else go to
Option 2.

1.  Install [clasp](https://developers.google.com/apps-script/guides/clasp) as
    per the instructions.
1.  Download the github project using

    ```bash
    git clone https://github.com/google/dcm-trix-addon
    ```

1.  Change to the project directory.

    ```bash
    cd dcm-trix-addon
    ```

1.  Authenticate and allow clasp to upload files on your behalf.

    ```bash
    clasp login
    ```

    If running on a remote machine, run the following instead and follow the
    instructions.

    ```bash
    clasp login --no-localhost
    ```

1.  Create a new apps script project in the spreadsheet created earlier. Replace
    `[doc id]` with the value obtained earlier.

    ```bash
    clasp create dcm-trix-addon [doc id]
    ```

1.  Deploy the chagnes to the apps script project.

    ```bash
    clasp push
    ```

1.  Open the script editor by going to the spreadsheet window and selecting
    `Tools > Script Editor`. Verify that all the files have been copied over.

After completing this step, continue with [Setting up
authentication](#setting-up-authentication).

#### Upload files manually - Option 2

1.  Download all the files in the repo using the download as zip option.
1.  Extract the zip and open all the files in a text editor (Sublime text,
    notepad++ etc).
1.  Open script editor by going to the spreadsheet and going to `Tools > Script
    Editor`.
1.  Give the script a name (Ex: dcm-trix-addon).
1.  Create each individual file in the script editor with the same names as the
    files downloaded and copy paste the contents. Do this for all the files
    except `appsscript.json` which is done in the next step.
1.  Go to `View > Show manifest file` in the script editor. Copy the contents of
    `appsscript.json` and replace the contents of this file.

Once completed, move to the next section to continue the setup.

## Setting up Time Zone

1. Go to `Project Settings` and change the Time Zone to the desired one. This will directly influence in the Scheduling function of the code.
1. If no changes are made in this section, the Timezone selected for the Project will be GMT-03:00.

## Setting up authentication

1.  Go to `File > project properties` and copy the script id. You will need this
    later.
1.  In the same script editor go to `Resources > Cloud platform project` and
    click on the project link.
1.  Click on the button at the top left corner of the page (three horizontal
    lines) to open the project menu.
1.  Go to `APIs & Services > Dashboard`
1.  Click on `Enable APIS and Services` at the top of the page.
1.  Search for "DCM" and click on `DCM/DFA Reporting And Trafficking API`.
1.  Click on `Enable`.
1.  Go to `APIs & Services > Credentials`
1.  Create new credentials by going to `Create credentials -> Oauth client id ->
    Web application`
1.  Give it a name (Ex: dcm-trix-addon).
1.  Under `Authorized redirect URIs` which is the second field in
    `Restrictions`, paste `https://script.google.com/macros/d/{SCRIPT
    ID}/usercallback` with `{SCRIPT ID}` replaced with the value obtained
    earlier.
1.  Hit 'create' - copy the client id and client secret values and paste them in
    `Constants.gs` file for `CLIENT_ID` and `CLIENT_SECRET` respectively.
1.  Go back to the spreadsheet and refresh the document. Verify that the addon
    is visible in `Add-ons > dcm-trix-addon`. The name may be different if you
    had given a different name earlier.

## First time run configuration

1.  Go to `Add-ons > dcm-trix-addon > Select DCM Report`. You should get a popup
    for authorization.
1.  In the popup click on `Continue`.
1.  Select the account you want to use this addon with.
1.  In the next screen, click on `Advanced`, scroll to the bottom and click on
    `Go to dcm-trix-addon (unsafe)`
1.  Review the permissions and scroll to the bottom and click on `Allow`.
1.  Close the popup that's open in the spreadsheet and go to `Add-ons >
    dcm-trix-addon > Select DCM Report` again.
1.  Click on `Authorize` and follow the steps again. If all goes well, you
    should see a page that says `Success! You can close this tab and open the
    Dialog box again.` Sometimes you will see a blank page. If that's the case,
    try switching to another tab and switching back again and see if shows the
    message now. If it still doesn't show, proceed to the next step to try and
    see if authentication was successful.
1.  This authentication needs to be done only the first time for each sheet.
1.  Back in the spreadsheet, close the popup and go to `Add-ons >
    dcm-trix-addon > Select DCM Report` again. You should see a popup that says
    `DCM Report Selection`. If that's the case, the configuration has been
    successful and you can now use the addon.

If you encounter any issues, review the setup and check whether you have
followed all the steps.

## Limitations

-   This addon can only download files that are less than 10MB in size. This is
    a limitation of Apps Script.
