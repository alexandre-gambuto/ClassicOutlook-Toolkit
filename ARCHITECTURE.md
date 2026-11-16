# ARCHITECTURE

A map of the whole VBA codebase (Outlook Classic project + the code of the three
companion workbooks). Read this before any change that touches several modules, the
event routing, the referential file paths, or the ribbon entry points.

State described: November 2026 (audit corrections of 2026-11-16).

---

## 1. Hard environment constraints

- **Outlook Classic only** (not New Outlook, not Outlook on the web). Main hook:
  `ThisOutlookSession` + `WithEvents` classes.
- **64-bit Office**: every `Declare` uses `PtrSafe` and `LongPtr` for handles.
- **No admin rights**: nothing may require admin (no `HKLM` key, no MSI, no globally
  registered COM). Registry writes are limited to `HKCU` via `SaveSetting` /
  `GetSetting`.
- **No VBA Project Object Model access**: code cannot manipulate VBA modules
  programmatically. Corollary: importing `.bas` / `.cls` is impossible on a locked
  machine - everything is pasted by hand in the VBE.
- **Late binding only** for Excel, Word, ADODB, WScript.Shell, Shell.Application,
  MSXML, Scripting.Dictionary. Never add a COM reference under Tools > References.
- **Option Explicit** on line 1 of every module.
- **EDR on the workstation**: no PowerShell and no `cmd.exe` launched from Outlook.
  Processes the code does start: `mshta.exe` (toast notifications, `modNotify`),
  `notepad.exe` (log, reports, `config.ini`), `explorer.exe` (open a folder),
  plus `Shell.Application.ShellExecute` (open a web page or a file with its default
  application). The clipboard is written through the Windows API (`modSendToLLM`).
- **Dialogs**: `MsgBox` / `InputBox` through `modUi` and `modMsgBoxCustom`, toasts
  through `modNotify`. The only UserForms of the Outlook project are the two windows
  of `modMailAnalytics` (`frmConfig`, `frmProgress`); no other form is added.
- **Deferred work**: only `modDefer` (Windows timer). `Application.OnTime` does not
  exist in Outlook.
- **UI strings** are ASCII in the source: accents come from the `A()` decoder (markers
  `~e`, `~a`...); bilingual text comes from `T(fr, en)`, which applies `A()` to both
  branches. Log strings stay ASCII and untranslated.

---

## 2. Module topology

### Orchestration
- `ThisOutlookSession`: single entry point: `Application_Startup` / `Application_Quit`,
  global Outlook events (Sent Items, calendar, reminders, `ItemSend`, new mail),
  role windows, command channel, a few ribbon shims.
- `clsAppEvents`: `WithEvents` on `Application.Inspectors`, on the mail Explorer
  (inline responses) and on `Application.Explorers` (first window of a session started
  without one); relays to `modNewMeetingDefault`, `modReplyInPlace`,
  `modReplyGreeting`, `modHideBcc`, `modMaximizeWindow`, `modMeetingToAppointment` and
  `ThisOutlookSession.OnFirstExplorer`.
- `clsApptWatcher`: per-appointment window watcher (Teams block removal on a meeting
  I organize without third-party invitee, conversion toast).
- `clsExplorerWatcher`: per-Explorer watcher pinning one window to the mail role and
  one to the calendar role.
- `clsInboxWatcher`: watches the Inbox of one mailbox for new-mail alerts
  (`modInboxAlert`).
- `clsFolderWatcher`: watches one `Folders` collection so that folders created,
  renamed, moved or deleted in Outlook enter the QuickMove list at once.

### Infrastructure
- `modConfig`: loads `config.ini`, getters (always call the getter, not a `g_*`
  global), `T()` / `A()`, toolkit folder (`ToolkitDir`, `MacroVBADir` = `Sorties\`,
  `WorkbooksDir` = `Classeurs\`, `VBADir` = `Code\`), `SetToolkitFolder`.
- `modDebugLog`: 4-level logger with per-module filter, `LogClip` (shortened text),
  `LastLineTime` (end of the previous session); rotation `debug.log` ->
  `debug.log.old` -> `debug.log.old2`.
- `modDefer`: deferred actions (see section 4), "one heavy task at a time" guard
  (`BeginHeavy` / `EndHeavy`).
- `modCommon`: stateless helpers (file names, HTML escaping, folder by path, safe
  property reads, DASL date filter `FormatDaslDate`, 64-bit tick counter `TickMs` /
  `ElapsedMs`, real wait `SleepMs`).
- `modUi`: dialogs (`Inform`, `Ask`, `Fail`, `AskText`, `AskTextEx` which tells
  Cancel from an empty answer, `Toast`), foreground guarantees.
- `modMsgBoxCustom`: `MsgBox` with renamed buttons and an auto-close countdown.
- `modNotify`: auto-closing toasts (`mshta`, JavaScript), multi-monitor aware,
  `IsRectVisibleOnScreen`.
- `modSmtp`: SMTP resolution (sender, recipients, organizer, external guests),
  `MyAddresses` / `IsMyAddress` / `IsFromMe`, `IsInternalDomain`,
  `HasExternalRecipient`, message kind (`MessageClassKind`: receipt vs non-delivery),
  auto-reply detection (`IsAutoReplyItem`).
- `modExcelBridge`: Excel COM lifecycle (ghost instance, stealth open with retries,
  open workbook found by name when OneDrive shows it as a web address, block reads,
  last data row under a filter, safe save and close).
- `modInspectorQueue`: keeps Inspector objects alive between `NewInspector` and a
  deferred pass.
- `modStoreContext`: multi-mailbox context (store of the displayed folder, account to
  send with, system folders of a given store).
- `modDevTools`: `Reinit` (rebinds events after a VBE save), `ReinitSilent`,
  `PrepareForEdit`.
- `modMaintenance`: maintenance menu (`Maintenance`, `MaintenanceQuick`), log level
  and module filter (pages of 30 modules, "+" = next page), paths diagnosis (two
  windows), module versions (full report written to `Sorties\versions-modules.txt`
  and opened in Notepad), bindings check (`AssertBindings`, 6 bindings). "Open
  config" without a `config.ini` starts the quick start guide.
- `modQuickStart`: first-run guide (environment check, five questions, minimal
  `config.ini`).
- `modBackupVBA` / `modBackupAuto`: manual backup of the VBA project; monthly
  automatic backup (VBA project, ribbons, workbooks, registry sections, caches) with a
  size alert on the VBA project file.

### Mail
`modQuickMove` (filing, folder index of all mailboxes, recent folders, "Go to folder",
undo of the last filing), `modReplyInPlace` (files sent replies and forwards next to
the original), `modReplyGreeting` (greeting line in replies), `modReplySuspicion`
("this mail seems to await a reply" proposal), `modMailResponseTracker` (tracking of
sent mail awaiting replies, Excel dashboard, morning reminder),
`modMailConfidentiality` (confidentiality / urgency banner cycle, quoted-history
boundary `OwnZoneEnd`), `modMailCampaign` (campaign engine driven by the campaigns
workbook), `modHideBcc`, `modInboxAlert` (per-mailbox new-mail alerts), `modSnooze`
(postpone mails, catch-up of missed reminders), `modSearch` ("sent / received this
week" searches and back), `modMailExport` (export to PDF / Word), `modMarkAllRead`
(mailbox of the displayed folder), `modThreadAttachments` (re-attach files from the
thread), `modSendToLLM` (send the text being written to a browser writing assistant),
`modFolderScan` (auto-archiving of inactive project subfolders), `modFolderMerge`,
`modFolderSort` (A-Z order of sibling folders), `modFolderRename` (CSV review of
folder names), `modViewGuard` ("Standardize view", manual only),
`modMailAnalytics` (mail statistics, see below).

### Calendar and project tagging
`modProjectTag` (project / contract / opportunity tag in appointment subjects, choice
window, remembered on/off switch of the automatic question `AutoAskEnabled` /
`ToggleAutoAsk`, `StripProjectTag` for external guests, batch mode
`CheckAssignment(appt, batchMode)` / `BatchStopRequested` for "tag recent
appointments", projects loaded silently when the load is automatic),
`modMigrateTags` (refuses a property name still in use in `config.ini`),
`modMeetingColor` (category by attendee type), `modPurgeCategories` (removes only the
automatic categories of past appointments), `modAntiOverlap` (+15 min shift of a short
plain appointment that collides; `MarkAsUserPlaced` protects an appointment placed by
the -15 / +15 min or actual-duration buttons), `modNewMeetingDefault` (default
duration when Outlook's own slot is kept), `modMeetingToAppointment` (meeting without
third party -> plain appointment, Teams block, `InspectorHasTeamsLink`),
`modAppointmentPrivacy`, `modShiftAppointment` (-15 / +15 min, actual
duration), `modFlashAppointment`, `modMaximizeWindow`.

### Contacts
`modContactCommon` (shared helpers and the undo file of bulk changes, command
`ContactUndoLastRun`; calls `modConfig`, `modDebugLog`, `modUi`, `modMsgBoxCustom`,
`modDefer`), `modContactFromMail`, `modContactEnrich` (fill in e-mail / phones from
mail), `modContactNormalize`, `modContactBirthday` (birthday calendar and its view),
`modContactRelation` ("cooling relationships" report, reads all mail folders),
`modContactNote`, `modContactExport` / `modContactImport`.

### Mail statistics (Outlook project)
`modMailAnalytics` runs inside the Outlook VBA project with its two UserForms
(`frmConfig`, `frmProgress`; layout in the `.frm` / `.frx`, code also in
`Interface\Formulaires\*_Code.txt`). Commands `RunEmailAnalysis` (full analysis)
and `RunEmailAnalysisDelta` (update of the latest report found in `Sorties\`). It
writes an Excel report (`Stats_Emails_FULL_*` / `_DELTA_*`). See
`Outlook\modMailAnalytics__README.txt`.

### Archived (kept in the repository, not in the installed project)
`modOneNote` and `modOneNoteProbe` (OneNote reorganization, blocked on the target
workstation by a broken COM registration). Their header says
`ARCHIVE (11/2026)`; `validate.py` skips them and reports any remaining reference.
To take them out of an installed project: paste `modMaintenance` v4.6 first, then
*File > Remove* `modOneNote` and `modOneNoteProbe` (no export), *Debug > Compile*.

### Workbook code (separate Excel VBA projects)
- `Excel\Suivi de temps\`: `modSuiviTemps` (time tracking, private, French) and its
  `ThisWorkbook`.
- `Excel\Suivi des mails\`: `modMailTracker`, `ThisWorkbook`, sheet code.
- `Excel\Suivi des Campagnes\`: `modCampaignBook`, `ThisWorkbook`.

Workbook code cannot call the Outlook modules: it talks to Outlook through the command
channel (section 4) and reads the toolkit folder from `HKCU`. Each workbook has its
own ribbon tab (`customUI14.xml` inside the file; for the time-tracking workbook, the
tab appears once its prepared `customUI14.xml` is injected, its sheet buttons stay
meanwhile).

---

## 3. Outlook startup flow

`ThisOutlookSession.Application_Startup`, immediate part:

1. Reads the time of the last line of `debug.log` (`modDebugLog.LastLineTime`)
   before anything is written, then `modConfig.LoadConfig` and the `SESSION START`
   line.
2. Previous shutdown: the `HKCU Toolkit\Session\LastEnd` mark written by
   `Application_Quit` is read and consumed; the log says "clean" or "abnormal"
   (with the time of the last line). A restart less than 30 minutes after a clean
   shutdown skips the catch-up tasks (colours, sent items, archiving).
3. Safe start: every start increments `HKCU Toolkit\Session\PendingStarts`, reset by
   a clean shutdown or by `StartupConfirm` (250 s). After three consecutive starts
   that were neither closed cleanly nor confirmed, the deferred startup tasks are not
   armed (events stay bound); an error line and a toast say so; macro
   `ResumeStartupTasks` arms them by hand.
4. Binds `mSentItems` (default Sent Items), starts `modInboxAlert`, binds the command
   channel (`Notes\ClassicOutlook-Toolkit`), reads the remembered state of the
   automatic tagging question (`modProjectTag.AutoAskEnabled`), binds
   `CalendarItems` (default calendar).
5. `modNewMeetingDefault.InitEvents` (creates `clsAppEvents`), before the role
   windows so that the reading-pane listener binds to the real mail window.
6. Role windows: each open Explorer gets the role of what it shows (mail or
   calendar); the missing window is created at the first switch. Outlook started
   without any window (by a workbook): the first window that opens gets its role
   then (`OnFirstExplorer`). Then `modReplyInPlace.EnsureInit`.
7. Deferred tasks armed (table below), unless safe start; `StartupConfirm` armed.
8. Bindings check (`modMaintenance.AssertBindings`, silent if all is well) only when
   a mail window exists (Outlook started by a workbook has none yet).

Deferred part (`modDefer`, seconds after startup; `*` = skipped after a quick restart):

| Delay | Key | Task |
|---|---|---|
| 8 s | `QuickStart` | first-run guide, only without `config.ini` |
| 15 s* | `CatchUpColor` | colours appointments synchronised while Outlook was closed |
| 20 s | `RemindOverdue` | morning reminder of overdue tracked mails (reads `Sorties\suivi-echeances.txt`, no Excel; light recalculation once a day at the first idle moment) |
| 25 s* | `SentCatchUp` | files mails sent from a phone while Outlook was closed |
| 45 s | `TagPreload` | loads the project list (silent if the workbook is missing) |
| 60 s | `QMWatch` | folder watchers for the QuickMove list |
| 120 s* | `AutoArchive` | auto-archiving of inactive project subfolders |
| 180 s | `SnoozeCatchUp` | brings back postponed mails whose reminder was missed |
| 210 s | `MonthlyBackup` | monthly backup (first session of the month) |
| 240 s | `Reindex` | QuickMove reindex in small batches, only if the last full scan is older than 7 days |
| 250 s | `StartupConfirm` | confirms the start (resets the safe-start counter) |

Heavy tasks run one at a time (`modDefer` postpones a heavy task while another one
runs). Keep delays short: a timer still armed during a VBE recompilation can crash
Outlook; in the first minutes after startup, run `PrepareForEdit` before editing code.

`Application_Quit` writes the `LastEnd` mark, resets `PendingStarts`, then writes the
`SESSION END` line. `RebindEvents` rebinds the events after a VBE save; it is called
by `modDevTools.Reinit` and by the bindings watchdog (`modDevTools.ReinitSilent`),
which runs at each send, each new mail and each reminder.

---

## 4. Post-startup event flow

### `Application.Inspectors.NewInspector` (via `clsAppEvents`)
- `modNewMeetingDefault.OnNewInspector`: a new appointment that kept Outlook's own
  default slot (`OutlookDefaultSlot`, 30 min) gets the toolkit default duration; a
  range chosen with the mouse or an imported event keeps its duration.
- `modReplyInPlace.OnNewInspector`: for a Reply / Reply all / Forward, remembers
  where the original message lives.
- `modReplyGreeting`, `modHideBcc`, `modMaximizeWindow`, Teams pre-cleaning:
  deferred passes (`modDefer`).
Each listener is isolated (`On Error Resume Next` in the relay).

### `mSentItems_ItemAdd` (an item lands in Sent Items)
1. Campaign mail -> `modMailCampaign.HandleSentItem` (filed in the campaign folder,
   sheet updated later through `CampaignSent`); never offered for manual filing.
2. `modReplySuspicion.CaptureSentItem` (before any move).
3. `modReplyInPlace.TryMoveSentItem`: a reply or forward is filed next to the
   original (through `modQuickMove.MoveItemTo`, which keeps the moved item, tells
   `modReplySuspicion` its new identifier and remembers the move for
   `QuickMoveUndo`).
4. Otherwise, code detected in the mail (subject, attachments, body) and exactly one
   folder whose last segment carries that code, not glued to other digits
   (`modQuickMove.FindCodeMatches`; a non-numeric code may fall back on the full path)
   -> moved by `MoveItemTo`, toast.
5. Otherwise the mail joins the queue of sent mails awaiting filing (session memory)
   and a toast says "Mail envoyé non classé (n en attente)". No window opens by itself
   after a send. The command **"Classer les envoyés en attente"**
   (`modQuickMove.QuickMoveSentPending` -> `ThisOutlookSession.FileSentPending`)
   opens the filing window for each mail still in Sent Items (search pre-filled with
   the detected code, three recent folders `r1`..`r3`; empty = leave it in Sent Items,
   Cancel = stop the series). The window runs as a heavy task (no deferred heavy task
   starts meanwhile).

### `Application_ItemSend`
- Meeting request: the subject sent is the subject of the appointment (tag
  included), **without the project tag when an invitee is external**
  (`modSmtp.HasExternalRecipient`, `modProjectTag.StripProjectTag`; a recipient whose
  SMTP address cannot be found counts as external). Cancellations, responses and
  forwards keep Outlook's prefix; their tag is also stripped for external recipients.
  The organizer's appointment keeps its tag.
- Campaign mail: no detection, no pop-up; `modMailCampaign.CampaignOnItemSend` notes a
  send from a secondary account (its Sent Items are read 30 s later).
- Other mails: urgent banner -> tracking line (`UrgentFlush`; a banner without a date
  asks before sending); follow-up sent from the tracking workbook -> `FollowupSent`;
  only for a mail that is neither of these,
  `modReplySuspicion.EvaluateAndStore` (score on the text written by the user, quoted
  history and signature excluded; proposal 2 s after the move).

### `Application_Reminder`
- A legacy QuickMove reindex marker appointment (pre-v3) is deleted.
- Mail reminder: `modSnooze.HandleSnoozeReminder` (back to the original folder,
  unread).

### `CalendarItems_ItemAdd` / `ItemChange`
- `ItemAdd`: on an appointment I organize, `modProjectTag.CheckAssignment` opens the
  project search directly (Cancel = no tag) while the automatic question is on (switch
  `ToggleAutoAsk`, remembered in the registry); then colour, anti-overlap, privacy,
  meeting-to-appointment conversion.
- `ItemChange`: re-tag if the subject lost its tag (`CheckReTag`), offer to tag an
  accepted invitation once, then the same chain.
- Tagging an occurrence of a recurring meeting asks "whole series / this occurrence
  only" (the series is tagged on its master).
- Events that arrive while the tagging question is open are queued and replayed
  (`CalReplay`) without the question.

### Role windows (`clsExplorerWatcher` -> `HandleExplorerSwitch`)
When the separate-calendar-window feature is on, a mail window that switches to the
calendar is sent back to the Inbox and the calendar window is activated (and the
other way round), with a 1.5 s anti-bounce guard (64-bit tick counter). A calendar
window created by the toolkit takes the position remembered by
`RememberCalendarPosition` (ignored when it would be off screen); the user may then
move or resize it freely during the session.

### Command channel (workbooks -> Outlook)
A workbook writes a note `TOOLKIT|<command>|<timestamp>|...` in
`Notes\ClassicOutlook-Toolkit`; `mCmdItems_ItemAdd` / `ItemChange` route `REFRESH` /
`REPAIR` (mail tracking, acknowledged by rewriting the note `ACK|...`) and
`CAMPAIGN_*` (campaigns, `CampaignCmd`). The last timestamp handled is kept in
`HKCU Toolkit\CmdChannel\LastStamp`: a command is never run twice, even after a
restart. A command older than 120 s is ignored (toast "relancez-la" when it is less
than 15 minutes old). Nothing heavy runs inside the event: everything goes through
`modDefer`.

### `modDefer` dispatch keys
Every `modDefer.Arm*("Key")` must have a matching `Case "Key"` in
`modDefer.Dispatch`. Heavy keys (one at a time) are marked **H**.

| Key | Target | Armed by |
|---|---|---|
| `ReplySusp` | `modReplySuspicion.ProposeDeferredFollowup` | "awaiting reply" proposal |
| `MeetingAdjust` | `modNewMeetingDefault.DelayedAdjust` | new appointment |
| `AutoArchive` **H** | `modFolderScan.AutoArchiveMove_OnTime` | startup |
| `CatchUpColor` **H** | `modMeetingColor.CatchUpUncolored_OnTime` | startup |
| `RemindOverdue` **H** | `modMailResponseTracker.RemindOverdue` | startup |
| `FollowupSent` | `modMailResponseTracker.RegisterFollowupSent` | follow-up sent |
| `UrgentFlush` | `modMailResponseTracker.FlushUrgentTracking` | urgent mail sent |
| `DashRefresh` **H** | `modMailResponseTracker.RefreshDashboard` | workbook Refresh / Repair, daily light recalculation |
| `MaxWindow` | `modMaximizeWindow.MaximizePass` | new window |
| `SentCatchUp` **H** | `modReplyInPlace.CatchUpRecentSent` | startup |
| `ActivateRole` | `ThisOutlookSession.ActivateRoleDeferred` | role windows |
| `TeamsPreClean` | `modMeetingToAppointment.DelayedTeamsPreClean` | reply with a meeting |
| `ReplyGreeting` | `modReplyGreeting.InsertGreetingDeferred` | reply draft |
| `TagPreload` **H** | `modProjectTag.InitProjectTag` | startup, `Reinit` (+10 s) |
| `MonthlyBackup` **H** | `modBackupAuto.RunMonthlyBackupIfNeeded` | startup |
| `CampaignSent` | `modMailCampaign.CampaignSentFlush` | campaign mail sent (+10 s, pushed back at each send; +30 s after a send from a secondary account) |
| `HideBcc` | `modHideBcc.HideBccPass` | compose window |
| `InboxFlush` | `modInboxAlert.InboxAlertFlush` | new mail |
| `BirthdayView` **H** | `ThisOutlookSession.SetupBirthdayViewDeferred` | first display of the calendar |
| `Reindex` **H** | `modQuickMove.ReindexStep` | startup, folder rescan |
| `QMWatch` | `modQuickMove.FolderWatchStart` / `FolderWatchStep` | startup |
| `CampaignNext` | `modMailCampaign.CampaignSendNext` | campaign send (one row per tick) |
| `CampaignCmd` **H** | `modMailCampaign.CampaignRunPending` | campaign command from the workbook |
| `QuickStart` | `modQuickStart.QuickStartAuto` | startup without `config.ini` |
| `CalReplay` | `ThisOutlookSession.CalendarReplayDeferred` | calendar events queued during tagging |
| `SnoozeCatchUp` | `modSnooze.SnoozeCatchUp` | startup |
| `StartupConfirm` | `ThisOutlookSession.ConfirmStartupDeferred` | startup |

Removed in November 2026: the monthly flag cleanup (`CleanFlags`, `CleanOldFlags`)
and `SentPrompt` (the filing window of sent mails opens only on command).

---

## 5. Public entry points (ribbon bindings)

All names below are referenced by the `.exportedUI` files with the `Projet.` prefix
(the Outlook VBA project is named `Projet`). **Do not rename without regenerating the
ribbon exports.** Ribbon labels are French only (an exported ribbon does not follow
`Language=EN`).

| Window / tab | Macros |
|---|---|
| Quick access toolbar | `MarkAllRead`, `SnoozeList`, `FolderSortAZ`, `QuickMoveUndo` |
| Home (mail) | `QuickMove`, `QuickMoveSentPending`, `GoToFolder`, `SnoozeEmail`, `MarkMailForTracking`, `OpenTrackingDashboard`, `FollowUpOverdue`, `SentMailsThisWeek`, `ReceivedMailsThisWeek`, `ResetSearch`, `ExportSelectedMails`, `MergeFolders`, `ContactCreateFromMail` |
| Calendar | `CreateFlashAppointment15`, `ToggleAutoAsk` |
| Selected appointment | `TagSelection`, `CreateFolderFromAppointment`, `ShiftAppointmentMinus15`, `ShiftAppointmentPlus15`, `SetActualDuration` |
| Contacts | `ContactAddNote`, `ContactNormalizeSelected`, `ContactNormalizeAll`, `ContactEnrichFromMail`, `ContactUndoLastRun`, `ContactBirthdaySync`, `RelationScan`, `ContactExportAll`, `ContactImportCsv` |
| Folder | `ExportProjectFolderNames`, `RenameProjectFoldersFromCSV` |
| View | `StandardizeView`, `RememberCalendarPosition` |
| Emaileur (campaigns) | `CampaignHelp`, `CampaignDiagnostic`, `CampaignCleanDrafts`, `CampaignOpenDrafts`, `CampaignOpenWorkbook`, `CampaignMenu`, `CampaignCheck`, `CampaignChangeTemplate`, `CampaignCreateSheet`, `CampaignGenerateRow`, `CampaignGenerateBatch`, `CampaignResetStatuses`, `CampaignSend`, `CampaignStop` |
| Statistiques | `RunEmailAnalysis`, `RunEmailAnalysisDelta` |
| Maintenance | `Maintenance`, `MaintenanceQuick`, `Reinit`, `QuickMoveReload`, `OpenLog`, `ClearLog`, `MaintCycleLogLevel`, `MaintFilterModules`, `MaintOpenConfig`, `ReloadConfig`, `SetToolkitFolder`, `QuickStart`, `MaintOpenVBAFolder`, `BackupVBA`, `RunBackupNow`, `OpenBackupsFolder`, `MaintDiagnosePaths`, `MaintModuleVersions` |
| Read message | `QuickMove`, `MarkMailForTracking`, `ContactCreateFromMail` |
| New message | `MarkMailForTracking`, `ThisOutlookSession.Sensitivity_Confidential`, `AddThreadAttachments`, `SendToChatGPT`, `SendToClaude`, `SendToMistral` |
| Appointment window | `TagProject` |

Where they live: `modQuickMove` (`QuickMove`, `QuickMoveSentPending`, `GoToFolder`,
`QuickMoveReload`, `QuickMoveUndo`), `modSnooze`, `modMailResponseTracker`
(`MarkMailForTracking`, `OpenTrackingDashboard`, `FollowUpOverdue`: opens the
dashboard with overdue rows ticked, follow-ups are prepared by the workbook's
Relancer button), `modSearch`, `modMailExport`, `modFolderMerge`,
`modContactFromMail`, `modFlashAppointment`, `modProjectTag` (`TagSelection`,
`TagProject`, `ToggleAutoAsk`, `CreateFolderFromAppointment`), `modShiftAppointment`,
contact modules (`ContactUndoLastRun` in `modContactCommon`), `modFolderRename`,
`modViewGuard`, `modMailCampaign`, `modMailAnalytics`, `modMaintenance`,
`modDevTools` (`Reinit`), `modDebugLog` (`OpenLog`, `ClearLog`), `modConfig`
(`ReloadConfig`, `SetToolkitFolder`), `modQuickStart`, `modBackupVBA` (`BackupVBA`),
`modBackupAuto` (`RunBackupNow`, `OpenBackupsFolder`), `modMarkAllRead`,
`modFolderSort`, `modThreadAttachments`, `modSendToLLM`, `ThisOutlookSession`
(`Sensitivity_Confidential`, the only confidentiality entry: cycle none ->
CONFIDENTIEL -> URGENT -> PERSONNEL -> none; `RememberCalendarPosition`).

Useful commands without a button (Alt+F8): `ResumeStartupTasks`,
`ForgetCalendarPosition`, `SnoozeCatchUp`,
`RepairTrackerRows`, `PurgePastCategories`, `MigrateProjectTags`,
`CleanTeamsSelection`, `ColorAndTagRecentItems`, `QuickStartPreview` /
`QuickStartReset`, `ContactBirthdayResetViewFlag`, `PrepareForEdit`, `ReinitSilent`,
`DumpHTML`.

---

## 6. External referential files

**Toolkit folder** (`modConfig.ToolkitDir`): one folder named
`ClassicOutlook-Toolkit`, found at startup (remembered folder, else the Desktop or a
OneDrive root, or one level below) and remembered in `HKCU` for the Excel projects:

    ClassicOutlook-Toolkit\
      config.ini      configuration, read at startup
      Classeurs\      workbooks (time tracking, mail tracking, campaigns)
      Sorties\        outputs (see below), Sauvegardes\ (backups)
      Ressources\     runtime files (alert sound...)
      Code\           sources to paste into the VBE, by application: Outlook\,
                      Excel\ (one folder per workbook), Interface\ (Rubans\ =
                      ribbons, Formulaires\ = UserForms), Docs\, Outils\
                      (validate.py)

The former flat layout (`<Desktop>\Macro VBA\Outlook\`, workbooks on the Desktop) is
still recognized as is. Config values may use `%TOOLKIT%` (toolkit folder) and
`%NAME%` (Windows environment variable).

**Files written in `Sorties\`**:
- `debug.log` (+ `debug.log.old`, `debug.log.old2`); `debug_suivi_temps.log` (time
  tracking workbook); `EmailAnalytics_trace.log` (statistics).
- `versions-modules.txt`: module versions report (Maintenance > Versions des modules,
  opened in Notepad).
- `suivi-echeances.txt`: overdue summary of the mail-tracking workbook, written by
  Outlook and by the workbook at each save; read by the morning reminder without
  opening Excel.
- Contact reports (`ContactEnrich-*.txt`, normalisation, import, birthdays, relations)
  and their undo files `*-undo.tsv` (renamed `*-undo-done.tsv` once restored by
  `ContactUndoLastRun`, which writes `ContactUndo-*.txt`).
- Statistics reports `Stats_Emails_FULL_*.xlsx` / `Stats_Emails_DELTA_*.xlsx`.
- `Sauvegardes\`: manual and monthly backups (VBA project, ribbon exports read from
  `Code\Interface\Rubans`, the three workbooks, registry sections, QuickMove cache).

**Workbooks** (paths in `config.ini`, default `Classeurs\`):
- **Time tracking** (`Suivi de temps.xlsm`): source of truth for projects, contracts,
  partner names and hours per month. Read by column index (A to E), sheet
  "Temps par mois". Read-only for loading, read-write when creating a project entry.
- **Mail tracking** (`Suivi des mails.xlsm`): one row per tracked sent mail; hidden
  `_Schema` sheet written by `modMailResponseTracker` (logical column name -> index),
  read by the workbook's own VBA.
- **Campaigns** (`Suivi des Campagnes.xlsm`): one sheet per campaign; see
  `Docs\GUIDE-Campagnes.md`.

**Other**:
- `VbaProject.OTM` (`%APPDATA%\Microsoft\Outlook`): the file holding all VBA code;
  backed up by `modBackupVBA` and `modBackupAuto` (size alert above 25 MB).
- Caches (`%APPDATA%\Microsoft\Outlook\`): QuickMove folder index (`#QM3` format:
  store, path, EntryID, StoreID) and recent folders.
- Registry (`HKCU\Software\VB and VBA Program Settings\`): `ClassicOutlookToolkit\
  Paths`, `\Options` (`AutoAsk`: automatic tagging question), `\Recents`;
  `Toolkit\Session` (`LastEnd`, `PendingStarts`), `\InboxAlert`, `\QuickStart`,
  `\ContactEnrich`, `\CmdChannel` (`LastStamp`); `MailResponseTracker\Rappel`
  (`LastLightRefresh`), `\Relink` (resume point of the mail re-identification),
  `\RepairFail`; `QuickMove`, `Outlook_Calendrier\Position`, `BackupAuto\Schedule`,
  `FolderScan\AutoArchive`. All are saved by the monthly backup.
- UserProperties written on items (names never translated): the project tag
  (`UpTag` in `config.ini`, default `ProjectTag`; the time-tracking workbook writes
  the same name), `ProjectAsk`, the anti-overlap mark (`UpShift`), `EmaileurCampagne`,
  `FollowupConvID`, `SnoozeOriginId`, `BirthdayKey`, `BirthdayStore`.

---

## 7. Business conventions

### Codes (prefixes and patterns are configurable in `config.ini`)
- **Project code**: a configurable letter prefix + 4 digits (e.g. `XXNN`). A reserved
  high value means "contract without opportunity"; `...0` means "cross-cutting".
- **Contract**: `20YY_NNNNN` (5 digits after the year).
- **Opportunity**: `20YY_OT_NNNN` (or a `_DT_` variant), 4 digits.
- **Cross-cutting**: a configurable prefix letter + optional 2 digits (e.g. `A`, `A01`),
  handled by `ExtractCrossCode`.

### Tags in appointment titles (by `modProjectTag`)
- With opportunity: `[<project>/<name>] title (20YY_OT_NNNN)`
- Opportunity without project code: `[<opportunity>/<name>/<partner>] title`
- Without opportunity: `[<project>/<name>] title`
- Contract: `[<contract>/<name>] title`
- Cross-cutting: `title [<prefix>/<prefix>01]`
- The tag stays on the organizer's appointment; it is removed from the subject of the
  invitation sent when an invitee is external.

### Auto-applied Outlook categories
- Cross-cutting task → internal category.
- Opportunity (`_OT_` / `_DT_`) → commercial category.
- A specific contract can be mapped to a dedicated category (configurable).
- Meetings: internal / external category by attendee type (`modMeetingColor`); these
  two are the only ones removed by `PurgePastCategories`.
- Personal, private or confidential appointments: excluded from the automatic
  processing.

---

## 8. Code delivery convention (reminder)

Modules are delivered as complete, ready-to-paste files: ASCII, CRLF line endings,
starting directly with `Option Explicit`, with no `Attribute` / `VERSION` / `BEGIN` /
`END` lines (the VBE regenerates those). Each module has `MOD_NAME`, `MOD_VERSION` and
`ModVersion()`; every delivery gets a new version number, higher than any number
already used, and starts from the file in `Code\`. UI strings go through `T()`; log
strings stay ASCII and untranslated; functional strings (category names, status
values, folder-name matches, config keys, registry keys) are never translated.
Run `Outils\validate.py` before pasting.

**Paste together.** Modules delivered on the same day call each other's new (or
removed) procedures: paste the whole set in one session, then *Debug > Compile* once;
a partial paste does not compile. For the 2026-11-16 delivery, a safe sequence: the
foundations (`modCommon`, `modConfig`, `modDebugLog`, `modDefer`, `modExcelBridge`,
`modSmtp`, `modUi`, `modNotify`, `modMsgBoxCustom`, `modInboxAlert`), then
`modContactCommon` before the other contact modules, `modQuickMove` with
`modReplySuspicion` and `modReplyInPlace`, `modMailCampaign` after `modSmtp`,
`modExcelBridge` and `modQuickMove`, `modProjectTag` before `modDevTools`,
`modMailAnalytics` with the code of its two forms, `modSendToLLM` with
`modMailConfidentiality`, `modMaintenance` (it reads the versions of
`modMailAnalytics` and `modSendToLLM`), then `ThisOutlookSession`, `clsAppEvents`,
`modDefer`, `modQuickMove` and `modMailConfidentiality` must all be in place before
compiling (they call each other's new procedures, or no longer call removed ones).
Remove the two OneNote modules only after `modMaintenance` v4.6 is in place. Workbook
code is pasted in each workbook (the time-tracking workbook gets a `ThisWorkbook`
module for its new ribbon tab). Ribbons: import each `.exportedUI` file once, from a
window of the same kind, after exporting the current customization (one file per
window, see `README.md`); the workbook ribbons (`customUI14.xml`) are injected
separately.
