================================================================================
MAIL ANALYTICS (modMailAnalytics) - README
================================================================================
Version : V25.3 (2026-11-16). Rewritten for audit 11/2026 (F9-13): the
previous notice still described the standalone Excel tool of July 2026.

WHAT IT IS
--------------------------------------------------------------------------------
A module of the ClassicOutlook-Toolkit VBA project (Outlook Classic), with two
UserForms (frmConfig, frmProgress). It scans one or more mailboxes and writes
a multi-sheet Excel report: dashboard, summary, per hour / weekday / month /
year, evolution by mailbox, heatmaps, top contacts (internal / external /
unknown), number of recipients, To / Cc / Indirect split, reply rate and
reply delay, optional sent-message length, raw DATA sample, and a hidden
_ETAT sheet that stores the counters for later incremental updates.

Files:
  Outlook/modMailAnalytics.txt                      the module
  Interface/Formulaires/frmConfig.frm + .frx        launch form (layout + code)
  Interface/Formulaires/frmProgress.frm + .frx      progress form (layout + code)
  Interface/Formulaires/modMailAnalytics__frmConfig_Code.txt    code of frmConfig
  Interface/Formulaires/modMailAnalytics__frmProgress_Code.txt  code of frmProgress
The .frm files carry the same code as the *_Code.txt files (V25.3). The form
LAYOUT only exists in the binary .frx files: keep each .frx next to its .frm.

COMMANDS
--------------------------------------------------------------------------------
RunEmailAnalysis       ribbon button "Statistiques emails" (Alt+F8 also works).
                       Full analysis, or a period ("since year", custom dates).
RunEmailAnalysisDelta  update of the latest report (Alt+F8; a ribbon button
                       "Mettre a jour" next to "Statistiques emails" is
                       planned). No file dialog: it takes the most recent
                       Stats_Emails_FULL_*.xlsx or Stats_Emails_DELTA_*.xlsx of
                       the output folder, names it in the launch form, and
                       adds only the mail received or sent AFTER the exact end
                       of that report, up to now. If there is no usable report
                       (none, TEST report, or a report made before V25.3), it
                       offers to run a full analysis right away.

Only one analysis can run at a time (a second click shows a message), and the
toolkit's heavy deferred tasks wait until the analysis is over.

CONFIGURATION (config.ini of the toolkit)
--------------------------------------------------------------------------------
  StatMailboxN = name_fragment ; keyword ; #RRGGBB [; user_email [; YYYY-MM-DD]]
    name_fragment : part of the mailbox name as shown in Outlook
    keyword       : fallback to recognise the user when no address is known
    #RRGGBB       : mailbox colour in the report
    user_email    : address of the mailbox (recommended)
    YYYY-MM-DD    : ignore mail older than this date for this mailbox
  StatMailboxDefault = #RRGGBB   colour of mailboxes not listed
  InternalDomains = domain1|domain2   (toolkit key) internal domains

How the user is recognised (sent mail, To / Cc): by ADDRESS whenever one is
known - the mailbox address, and for the main mailbox (or a mailbox whose
address is one of the user's) all of the user's addresses (profile accounts and
Exchange aliases). The keyword is only a fallback for a mailbox without a known
address. The Address / Keyword option of the form is no longer read.

Internal / external: for an internal mailbox (no domain, or one of
InternalDomains), "internal" means any InternalDomains domain, sub-domains
included; for another mailbox, the domain of that mailbox. Addresses are
resolved by modSmtp (no directory call per mail): the former "Exchange
resolution" option is gone (hidden in the form).

System folders (deleted items, junk, sync issues, conflicts, local / server
failures, outbox, drafts, RSS) are recognised by identity, in any language;
non-mail folders (calendar, contacts, tasks, notes, journal) are skipped. Sent
Items and its sub-folders count as sent.

OUTPUT
--------------------------------------------------------------------------------
Folder: <toolkit folder>\Sorties (modConfig.MacroVBADir), fallback Desktop.
  Stats_Emails_FULL_YYYYMMDD_HHMMSS.xlsx    full analysis (or period)
  Stats_Emails_DELTA_YYYYMMDD_HHMMSS.xlsx   update (cumulated counters)
  Stats_Emails_TEST_YYYYMMDD_HHMMSS.xlsx    test mode (1 mail in 10), never
                                            used as an update base
  EmailAnalytics_trace.log                  detailed trace of the last run
                                            (one line per folder, every 250
                                            items, errors "!!!")
General log lines go to debug.log under the filter "Analytics".

If the report cannot be saved (output folder refused...), it stays OPEN in
Excel with a message: save it by hand. If the mailbox scan fails, a message
gives the error and the trace file.

Update (delta) rules: the _ETAT sheet stores PERIOD_START (start of the
cumulated period), PERIOD_END (exact end, included in the counters) and
STATE_FORMAT = 2. An update counts mail strictly after PERIOD_END, so no mail
is counted twice. In an update report, the DATA and Length sheets only cover
the added mail; the Summary sheet shows the cumulated period and what the
update added.

Reply rate and delay ("Replies" sheet): read on each received mail from the
last action Exchange records on it (Reply / Reply All, and its time). A
forward is not a reply. Replies made from a phone may not be recorded: to be
checked on the workstation.

PERFORMANCE
--------------------------------------------------------------------------------
Measured on 08/07/2026: 285,695 mails in 22 min 24 s (about 12,700 mails/min).
A restricted period or an update only opens the mail of that period (date
filter on each folder), so a monthly update takes a fraction of a full run.
Outlook stays responsive by bursts during the scan (DoEvents every 500 mails);
the report is written by blocks in a private Excel instance.

INSTALLATION / UPDATE IN THE VBA EDITOR
--------------------------------------------------------------------------------
1. Module: paste Outlook/modMailAnalytics.txt into the modMailAnalytics module.
2. Forms already present: open the code window of frmConfig and frmProgress
   and replace their code with the matching *_Code.txt file.
   Forms missing: File > Import frmConfig.frm / frmProgress.frm (with their
   .frx) - import may be blocked on a locked-down workstation.
3. Debug > Compile.

TROUBLESHOOTING
--------------------------------------------------------------------------------
"Ambiguous name detected"
  -> an old copy of the module (modEmailAnalytics...) is still in the project.
Nothing happens / message about the report
  -> read EmailAnalytics_trace.log in the output folder.
Many mails in "Indirect" (To / Cc sheet) or wrong sent / received split
  -> check the user_email field of StatMailboxN for that mailbox.

================================================================================
END OF GUIDE
================================================================================
