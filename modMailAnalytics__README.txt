================================================================================
EMAIL ANALYTICS - README
================================================================================

WHAT IT IS
--------------------------------------------------------------------------------
A single-module Excel/VBA tool that scans one or more Outlook mailboxes and
produces a multi-sheet Excel report: global statistics, per-mailbox volumes,
sent/received split, internal/external split, distribution by weekday and by
hour, monthly/half-yearly/annual evolution, heatmaps, top senders/recipients,
recipient-count distribution and (optionally) sent-message length analysis.

It is a standalone VBA project (separate from the Outlook toolkit) and ships as
three code files plus this README:

| File                                         | Description                  |
|----------------------------------------------|------------------------------|
| modEmailAnalytics_v23     | Single module (analysis +    |
|                                              | Excel reporting)             |
| modEmailAnalytics__frmConfig_Code_V23    | Configuration form code      |
| modEmailAnalytics__frmProgress_Code_V23  | Progress form code           |
| modEmailAnalytics__README_V23            | This file                    |

The module is organized into 12 clearly labelled sections (search "[SECTION N]"):
  [SECTION 1]  Public constants and global variables
  [SECTION 2]  Entry point (RunEmailAnalysis)
  [SECTION 3]  Analysis helpers (init, counters, folder cache)
  [SECTION 4]  Folder and email processing (hot path)
  [SECTION 5]  SMTP / address resolution
  [SECTION 6]  Generic helpers (dictionaries, domains)
  [SECTION 7]  Excel generation - entry point + Dashboard / Summary
  [SECTION 8]  Excel generation - By weekday / By hour
  [SECTION 9]  Excel generation - Evolution / Heatmaps
  [SECTION 10] Excel generation - Top 10 / Contacts / DATA
  [SECTION 11] Excel generation - Message length
  [SECTION 12] Excel generation - Time charts

REQUIREMENTS
--------------------------------------------------------------------------------
- Outlook Classic (desktop) with the mailbox(es) to analyze connected.
- Excel with VBA. Late binding is used for the Outlook objects.
- Two UserForms named exactly: frmConfig and frmProgress.
  (Unlike the Outlook toolkit, this project does use UserForms.)

INSTALLATION
--------------------------------------------------------------------------------
1. In the Excel VBA editor (Alt+F11), create two UserForms named frmConfig and
   frmProgress, with the controls the form code expects (list of mailboxes,
   colour combo, period pickers, test/resolution/length/debug toggles, progress
   labels and bar). Control names are referenced by the code as-is.
2. Import or paste modEmailAnalytics_v23 into a standard
   module (name the module modEmailAnalytics).
3. Paste the frmConfig / frmProgress code into the matching form code windows.
4. Debug > Compile, then run RunEmailAnalysis (Alt+F8).
5. For maximum speed, leave "Exchange resolution" OFF.

RECOMMENDED SETTINGS FOR MAXIMUM SPEED (in frmConfig)
--------------------------------------------------------------------------------
  [x] Select the mailboxes to analyze
  [ ] Exchange resolution   <-- turn OFF (very important)
  [ ] Length analysis       <-- off if not needed
  [ ] Test mode             <-- off for a full run
  [ ] Debug mode            <-- OFF for best performance

PERFORMANCE
--------------------------------------------------------------------------------
~400 to 800 emails/min with Exchange resolution OFF. Projection for ~80,000
emails: roughly 2.5 to 3.5 hours.

Key optimizations: index loop instead of For Each; MessageClass Like "IPM.Note*"
instead of TypeName; COM property caching (Sender, SenderEmailType, Subject,
ReceivedTime); direct Year/Month/Day/Hour instead of Format; excluded-folder
cache; ScreenUpdating off during generation; block (array) writes instead of
cell-by-cell; QuickSort O(n log n) for top-N; dynamic year detection.

SIZE NOTE
--------------------------------------------------------------------------------
The module is ~3350 lines (~2400 code, ~950 comments). VBA's practical per-module
limit is well below the theoretical 65k lines. If you add many features and the
effective code grows past ~5000 lines, consider splitting analysis and reporting
into two modules. Warning signs: "Procedure too large", slow compilation,
out-of-memory at runtime.

TROUBLESHOOTING
--------------------------------------------------------------------------------
"Ambiguous name detected" / duplicate procedure
  -> An old copy of the module is still present. Remove it and re-import.

"Constant TEST_SAMPLING not found"
  -> Only part of the code was imported. Re-import the whole module.

Compiles but "Object variable not set" at runtime
  -> Make sure RunEmailAnalysis is the entry point you launch, and that
     frmConfig and frmProgress both exist in the project.

NOTE ON LOCALE
--------------------------------------------------------------------------------
Folder detection matches French Outlook folder names (e.g. the Sent and Inbox
folders). On a non-French Outlook, adjust the folder-name tests in the code
(search for the Inbox / Sent matching logic) to your locale.

================================================================================
END OF GUIDE
================================================================================
