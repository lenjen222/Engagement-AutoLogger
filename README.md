# Engagement-AutoLogger
A Macro that takes relevant info from a master downloaded file and puts that information into a different open excel file.

At Thomson Reuters, associates are stuck doing repetitive, error-prone tasks. One of the most frustrating was maintaining the engagement log: every binder required manually copy-pasting three fields (client name, binder ID, page count) from a downloaded spreadsheet into a log Excel file.

Each entry = three Ctrl-C / Ctrl-V cycles.

Even downloading the “pending binders” export only reduced the steps slightly.

It was tedious, archaic, and ripe for mistakes.

Following the tax extension season, I've decided to build a binder autologger that will solve this problem for Thomson Reuters, or any other individual who needs to copy columns of information from one xlsx file to another. I estimate this autologger will save up to 5 minutes a day per associate, or 14 man hours a day during peak tax season.


HOW TO USE:
1) Open your engagement log

Open your engagement log in LibreOffice Calc.

If you don’t have a special sheet for the log, the macro will use the active sheet.

Recommended sheet name: EngagementLog

Recommended header row (Row 1): Client | Binder ID | Pages

2) Add the macro to the file

In Calc, press Alt + F11 to open the Macro editor.

In the left panel, select your engagement log file (not “My Macros”).

Click New → give the module a name like AutoLogger.

Paste the macro code you were given into the right pane and click Save.

If your export uses slightly different header names (like “Binder ID #” or “Total Pages”), edit the header list at the top of the macro. It matches case-insensitively and even partial names.

3) Enable macros (one-time)

Go to Tools → Options → LibreOffice → Security → Macro Security…

Set to a level that allows macros from documents you trust (e.g., “Medium”).

Click OK.

4) Run the macro

With your engagement log open, go to Tools → Macros → Run Macro…

Choose your document → the module you created → ImportPendingBinders → Run.

A file picker will appear. Select your latest pending binders .xlsx file.

The rows will be added to the first empty line under your headers.

5) (Optional) Add a toolbar button

View → Toolbars → Customize… → Toolbars tab

Pick a toolbar (Standard), click Add…, find your macro, and add it.

Now you can click one button to run it anytime.

How to use (everyday steps)

Download the day’s pending binders Excel file from your dashboard.

Open your engagement log in LibreOffice Calc.

Click your Import Pending Binders button (or run the macro).

Select the downloaded .xlsx file.

Done! The three columns are copied in automatically.

Notes & Tips

Column names: The macro expects headers like “Client”, “Binder ID”, “Pages.” If your export labels are slightly different, update the header strings at the top of the macro.

Duplicates: If you sometimes re-import the same binder, you can enable the simple “dedupe by Binder ID” line in the macro (there’s a comment next to it).

Sheet name: By default the macro writes to a sheet named EngagementLog. You can change that constant at the top of the macro if your sheet is named differently.

Header row: The macro assumes your headers are on the first row. If your headers are on row 2, change the HEADER_ROW value in the macro to 1.

Troubleshooting

It says “Header not found.”
Open the pending binders file and check the exact header text. Update the header strings in the macro (top of the file). The match is not case-sensitive and allows partial matches, but the words should be close.

It pasted in the wrong place.
Make sure your engagement log has headers in Row 1, and that Columns A, B, C are Client | Binder ID | Pages. The macro appends to the first empty row under your data.

Macro won’t run / blocked.
Check Tools → Options → LibreOffice → Security → Macro Security and allow macros from trusted documents.

Wrong sheet.
Change the TARGET_SHEET_NAME setting at the top of the macro to the sheet you want to use.

Numbers look weird (Pages).
The macro converts Pages to whole numbers. If your source has blank or text entries in that column, they’ll be treated as 0.

Safety & Privacy

The macro runs on your local computer.

It doesn’t upload files anywhere.

Only run macros from documents you trust. (You can trust me :) )

License

Use it, share it, improve it. If you find a bug or want a feature (like a simple de-duplication toggle, or a “Pick output file” dialog), open an issue or PR.
