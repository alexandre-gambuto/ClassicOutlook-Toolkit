# ARCHITECTURE

A map of the whole VBA codebase (Outlook Classic + the standalone Excel tool +
OneNote helpers). Read this before any change that touches several modules, the event
routing, the referential file paths, or the ribbon entry points.

---

## 1. Hard environment constraints

- **Outlook Classic only** (not New Outlook, not Outlook on the web). Main hook:
  `ThisOutlookSession` + `WithEvents` classes.
- **64-bit Office**: every `Declare` uses `PtrSafe` and `LongPtr` for handles.
- **No admin rights**: nothing may require admin (no `HKLM` key, no MSI, no globally
  registered COM). Registry writes are limited to `HKCU` via `SaveSetting` /
  `GetSetting` / `Shell("reg.exe ...")`.
- **No VBA Project Object Model access**: code cannot manipulate VBA modules
  programmatically. Corollary: importing `.cls` is impossible on a locked machine -
  everything is pasted by hand in the VBE.
- **Late binding only** for Excel, OneNote, ADODB, WScript.Shell, MSXML,
  Scripting.Dictionary. Never add a COM reference under Tools > References.
- **Option Explicit** at the top of every module.
- **UI strings** (`MsgBox`, `InputBox`) are ASCII, no accents: accents come from the
  `A()` decoder; bilingual text comes from `T(fr, en)`. Comments may use UTF-8 accents.

---

## 2. Module topology

### Orchestration
- `ThisOutlookSession`: single entry point: `Application_Startup`, global Outlook
  events, ribbon shims.
- `clsAppEvents`: `WithEvents` on `Application.Inspectors`, routes to modules.
- `clsApptWatcher`: per-appointment watcher (anti-overlap / colouring).
- `clsExplorerWatcher`: per-Explorer watcher pinning one window to the mail role and one to the calendar role.

### Outlook business modules
Mail: `modProjectTag`, `modMigrateTags`, `modQuickMove`, `modReplyInPlace`,
`modMailCampaign`, `modMailConfidentiality`, `modMailResponseTracker`,
`modReplySuspicion`, `modMarkAllRead`, `modSearch`, `modMailExport`, `modFolderScan`,
`modFolderMerge`, `modPurgeCategories`, `modSendToLLM`, `modReplyGreeting`.

Calendar: `modMeetingColor`, `modAntiOverlap`, `modNewMeetingDefault`,
`modMeetingToAppointment`, `modAppointmentPrivacy`, `modShiftAppointment`,
`modFlashAppointment`, `modSnooze`, `modMaximizeWindow`.

Infrastructure: `modConfig`, `modDebugLog`, `modDefer`, `modCommon`,
`modMsgBoxCustom`, `modDevTools`, `modMaintenance`, `modBackupVBA`,
`modBackupAuto`, `modNotify`, and since pack v3: `modSmtp` (SMTP resolution with
a session cache), `modExcelBridge` (shared Excel COM lifecycle), `modContactCommon`
(stateless helpers of the six contact modules), `modInspectorQueue` (Inspector
objects kept alive between `NewInspector` and a deferred `modDefer` pass).

### Excel / OneNote
- `modMailAnalytics`: standalone Excel analytics project (own `frmConfig` /
  `frmProgress` UserForms, own README).
- `modOneNote`: OneNote notebook inventory / reorganization.
- `modOneNoteProbe`: standalone diagnostic probe for the OneNote COM chain.

---

## 3. Outlook startup flow

`ThisOutlookSession.Application_Startup` runs, in order:

1. **Bind `mSentItems`**: `WithEvents` on the default Sent folder.
2. **`modFolderScan.AutoArchiveMove`**: two-way archiving of project / partnership
   subfolders (1-month threshold).
3. **Weekly QuickMove reindex**: on the first launch of the week (any weekday), via
   `modDefer.Arm "Reindex", 120` (pack v3: the hidden technical appointment and its
   `Application_Reminder` round-trip are gone).
4. **`CleanOldFlags`**: monthly cleanup of orphaned technical flags in the Inbox
   (flag != Snooze, older than a month).
5. **`modProjectTag.InitProjectTag`**: loads tagging preferences (auto-ask toggle).
6. **Bind `CalendarItems`**: `WithEvents` on the default calendar for auto-tagging.
7. **Bind `olMainExplorer`**: `WithEvents` on the main Explorer (calendar redirect to
   a separate window).
8. **`modNewMeetingDefault.InitEvents`**: instantiates `clsAppEvents` and subscribes
   to `Application.Inspectors`.
9. **`modReplyInPlace.EnsureInit`**: initializes the conversation→folder dict and
   caches the default store.

---

## 4. Post-startup event flow

### `Application.Inspectors.NewInspector` (via `clsAppEvents`)
On every inspector window (new mail, new appointment, Reply, Forward, ...):
- `modNewMeetingDefault.OnNewInspector`: for a blank, non-all-day, non-recurring
  appointment, forces `Duration = 15`.
- `modReplyInPlace.OnNewInspector`: for a Reply/ReplyAll/Forward, records
  `ConversationID -> FolderPath` (where the original message lived).

The two listeners are independent: an error in one does not stop the other
(`On Error Resume Next` in the relay).

### `mSentItems_ItemAdd` (an item lands in Sent)
1. **`modReplyInPlace.TryMoveSentItem`**: if the ConversationID is known, move the
   mail into the parent folder and return; done.
2. Otherwise **QuickMove auto**: extract a prefill (project / contract / opportunity)
   via `ExtractPrefill`; if exactly one cached match, move automatically.
3. Otherwise **QuickMove manual**: prompt the user with a filing `InputBox`
   (progressive refinement, numeric choice).

### `Application_Reminder`
- A legacy QuickMove reindex marker appointment (pre-v3) is simply deleted.
- If it's a `MailItem`: delegate to `modSnooze.HandleSnoozeReminder` (move back to the
  original folder, or Inbox as fallback).

### `modDefer` dispatch keys
Every `modDefer.Arm*("Key")` must have a matching `Case "Key"` in `modDefer.Dispatch`.
Keys in use: `MeetingAdjust`, `MaxWindow`, `HideBcc`, `TeamsPreClean`, `ReplyGreeting`,
`InboxFlush`, `BirthdayView`, and since pack v3 `Reindex` (weekly QuickMove index) and
`CampaignNext` (asynchronous campaign send, one row per tick).

### `CalendarItems_ItemAdd` / `ItemChange`
- `ItemAdd`: on an appointment the user created (not re-accepted), trigger
  `modProjectTag.CheckAssignment` and mark it as "asked".
- `ItemChange`: re-tag automatically (if the subject lost its tag) and offer to tag
  invitation responses (Accepted / Tentative).

### `olMainExplorer_FolderSwitch`
If the separate-calendar-window feature is on and the user switches to an appointment
folder in the main window: redirect to a dedicated window and send the main Explorer
back to the Inbox.

### `Application_ItemSend`
Intercepts outgoing `MeetingItem`s: applies `CheckAssignment` to the associated
appointment and syncs the tagged subject between MeetingItem and AppointmentItem.

---

## 5. Public entry points (ribbon bindings)

All names below are referenced by the `.exportedUI` files with the `Projet.` prefix
(the Outlook VBAProject is named `Projet`). **Do not rename without regenerating the
ribbon exports.**

- `ThisOutlookSession`: `Sensitivity_Normal`, `Sensitivity_Personal`,
  `Sensitivity_Private`, `Sensitivity_Confidential`, `RememberCalendarPosition`,
  `ForgetCalendarPosition`.
- `modQuickMove`: `QuickMove`, `GoToFolder`, `GoToMailFolder`, `QuickMoveReload`.
- `modProjectTag`: `TagProject`, `TagSelection`, `ToggleAutoAsk`,
  `CreateFolderFromAppointment`.
- `modSnooze`: `SnoozeEmail`, `SnoozeList`.
- `modMailResponseTracker`: `MarkMailForTracking`, `OpenTrackingDashboard`.
- `modSearch`: `SentMailsThisWeek`, `ReceivedMailsThisWeek`, `ResetSearch`.
- `modFlashAppointment`: `CreateFlashAppointment15`.
- `modShiftAppointment`: `ShiftAppointmentPlus15`, `ShiftAppointmentMinus15`.
- `modMarkAllRead`: `MarkAllRead`.
- `modBackupVBA`: `BackupVBA`, `RunBackupNow`.
- `modOneNote`: `RunReorganization`, `ExecutePlanFromFile`.
- `modMailCampaign`: `Campaign*` (menu, create sheet, generate row/batch, check, reset,
  send, clean drafts, ...) and `CampaignStop` (pack v3: interrupts an asynchronous send).
- `modSendToLLM`: `SendToClaude`, `SendToMistral`, `SendToChatGPT`.
- `modEmailAnalytics` (Excel project): `RunEmailAnalysis`.

---

## 6. External referential files

Paths are environment-specific and live in `config.ini` (not hard-coded). Typical setup:

- **Project / contract referential** (an `.xlsm` workbook): source of truth for
  projects, contracts and partner names.
  - Sheet with headers on row 1; columns for opportunity, contract, project code,
    project name, partner, then hours per month.
  - Read read-only for loading, read-write when creating a project entry.
- **Mail-tracking workbook** (an `.xlsm`, with its own dashboard VBA and action
  buttons): tracks replies to multi-recipient sent mail
  (id, send date, subject, recipients, recipient count, delay, deadline, reply count,
  responders, status, ConversationID, origin folder).
- **`VbaProject.OTM`** (`%APPDATA%\Microsoft\Outlook`): the file holding all VBA code;
  backed up by `modBackupVBA` to a timestamped copy.
- **Caches** (`%APPDATA%\Microsoft\Outlook\`): a full folder index (`#QM3` format since
  pack v3: store, path, EntryID, StoreID: rebinding by `GetFolderFromID`, stale paths
  self-healed from the live folder) and a short most-recently-used folder list for
  QuickMove.
- **Mail-tracking workbook `_Schema` sheet** (pack v3): hidden sheet written by
  `modMailResponseTracker` at each open (logical column name -> index); the workbook's
  own VBA reads its column indices by name from it, with the historical constants as
  fallback.

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
- Without opportunity: `[<project>/<name>] title`
- Contract: `[<contract>/<name>] title`
- Cross-cutting: `title [<prefix>/<prefix>01]`

### Auto-applied Outlook categories
- Cross-cutting task → internal category.
- Opportunity (`_OT_` / `_DT_`) → commercial category.
- A specific contract can be mapped to a dedicated category (configurable).
- Otherwise: none.

---

## 8. Code delivery convention (reminder)

Modules are delivered as complete, ready-to-paste files: UTF-8, CRLF line endings,
starting directly with `Option Explicit`, with no `Attribute` / `VERSION` / `BEGIN` /
`END` lines (the VBE regenerates those). UI strings go through `T()`; log strings stay
ASCII and untranslated; functional strings (category names, status values, folder-name
matches, config keys, registry keys) are never translated.
