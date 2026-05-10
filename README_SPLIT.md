# Apps Script split source

This folder contains the split `.gs` version of `C:\Users\jacyg\Desktop\APP\Appscript.txt`.

Validation performed after splitting:

- Original line count: 8376
- Joined split line count: 8376
- Joined split content matches the original line-by-line
- JavaScript syntax check passed on the joined split source

After the update-system module was added, syntax checks were also run on:

- Library source files `00_*.gs` through `13_*.gs`
- Client wrapper file `Bootstrap_Client.gs`

Import rule:

- Import `00_*.gs` through `13_*.gs` into the master Library Google Apps Script project.
- Do not import the original monolithic `Appscript.txt` into the same project at the same time, because that would duplicate global `const` and function names.
- In cloned spreadsheet projects, do not import all Library source files. Add the Library with identifier `YTTools`, then keep only `Bootstrap_Client.gs` as local code.

Files:

- `00_CoreConfig.gs`: global constants, task progress cache, shared dialog header, video column map.
- `01_ApiKeyManager.gs`: API key sheet read/write and API key inline view.
- `02_MenuGuideDialogs.gs`: menu, guide, range/single-row dialogs, quota and health helpers.
- `03_AdminPanel.gs`: admin dashboard and inline views.
- `04_YouTubeApiUtils.gs`: YouTube quota reminder, fetch retry/batch helpers, shared batch write/delete helpers.
- `05_VideoUpdate.gs`: video sheet update functions.
- `06_SubtitleTranscript.gs`: subtitle/transcript fallback system and subtitle guide/retry/clear.
- `07_ChannelUpdateAndFetch.gs`: channel sheet update and fetch recent videos from channels.
- `08_SearchByTopic.gs`: search videos/channels by topic.
- `09_Cleanup.gs`: cleanup tools and STT recalculation.
- `10_UtilsData.gs`: country/timezone map, formatting, parsing, YouTube channel helpers.
- `11_OAuthAnalytics.gs`: OAuth token flow and YouTube Analytics functions.
- `12_AIAnalyzer.gs`: AI sheet analyzer and history.
- `13_UpdateSystem.gs`: version/update manifest checker and update center dialog.

Support files:

- `Bootstrap_Client.gs`: wrapper file for cloned spreadsheet projects using the Library identifier `YTTools`.
- `manifest.example.json`: example update manifest that can be hosted on GitHub raw, Drive web endpoint, or another stable HTTPS URL.
