# Library update setup

## Goal

The master Google Apps Script project contains the full system code as a Library.
Every cloned spreadsheet project contains only `Bootstrap_Client.gs` and references the master Library as `YTTools`.

This keeps cloned spreadsheets small and lets you release new versions from one master project.

## Master Library project

1. Open the master Apps Script project.
2. Import/copy these files into it:
   - `00_CoreConfig.gs`
   - `01_ApiKeyManager.gs`
   - `02_MenuGuideDialogs.gs`
   - `03_AdminPanel.gs`
   - `04_YouTubeApiUtils.gs`
   - `05_VideoUpdate.gs`
   - `06_SubtitleTranscript.gs`
   - `07_ChannelUpdateAndFetch.gs`
   - `08_SearchByTopic.gs`
   - `09_Cleanup.gs`
   - `10_UtilsData.gs`
   - `11_OAuthAnalytics.gs`
   - `12_AIAnalyzer.gs`
   - `13_UpdateSystem.gs`
3. Save the project.
4. Open Project Settings and copy the Script ID.
5. Create a new Library version from Deploy > Manage deployments.

## Cloned spreadsheet project

1. Open Apps Script for the cloned spreadsheet.
2. Delete all copied system source files from the cloned project.
3. Add the master Library:
   - Click `Libraries` plus button.
   - Paste the master Script ID.
   - Select the released version.
   - Set identifier exactly to `YTTools`.
4. Add one local script file named `Bootstrap_Client.gs`.
5. Copy the whole content of this folder's `Bootstrap_Client.gs` into that local file.
6. Save, run `onOpen`, authorize, then reload the spreadsheet.

## Update manifest

Host a JSON manifest at a stable HTTPS URL. You can start from `manifest.example.json`.

Required fields:

- `version`: semantic/date version shown to users.
- `libraryScriptId`: the master Library Script ID.
- `libraryVersion`: the Library version number users should select.
- `releaseNotes`: human-readable changes.

In any cloned spreadsheet:

1. Open menu `YouTube Tools`.
2. Choose `Kiểm tra cập nhật hệ thống`.
3. Paste the manifest URL and save.
4. Click `Kiểm tra cập nhật`.
5. If an update is available, change the `YTTools` Library version in the Apps Script left sidebar.
6. Save and reload the spreadsheet.

## Important limits

Apps Script Libraries are versioned dependencies. A script cannot safely and universally switch its own Library version without using the Apps Script API and additional OAuth scopes.
This setup therefore uses a safe update center that checks the manifest, logs update checks, and tells the user exactly which Library version to select.

Do not store user sheet data, API keys, or history in the Library project. Those stay in each cloned spreadsheet.
