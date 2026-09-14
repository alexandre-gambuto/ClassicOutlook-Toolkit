# Troubleshooting and known pitfalls

Every entry below comes from a real incident of this project. Read it before opening an
issue: most "it does not work" reports end up here.

## The events stopped firing (nothing happens on send, on new appointment...)

- **Ctrl+S in the VBA editor recompiles the project and resets all module-level
  variables.** Every `WithEvents` binding is lost until the next startup. Run
  *Maintenance > Reinit* (or `Reinit` from the Immediate window) instead of restarting
  Outlook.
- **Tests with the VBA editor open are not representative.** Some events and the
  external toast process behave differently while the editor has the focus. Close the
  editor, then test.
- *Maintenance > Check bindings* (`AssertBindings`) tells which of the six bindings is
  dead and rebinds it.

## Outlook crashes 20 to 30 seconds after startup

One module does not compile, and the failure surfaces inside the deferred-action timer
callback, out of reach of any error handler. Run *Debug > Compile Project* in the
editor: fix the reported module. The usual causes are a module-level declaration placed
after a procedure, a duplicate `Dim`, or a variable named like a global helper (`t`,
`a`). `validate.py` catches all three before pasting.

## Nothing runs "in N seconds"

`Application.OnTime` does not exist in Outlook: the call fails silently. Every deferred
action goes through `modDefer` (`SetTimer` based). If a deferred action never fires,
check that its key has a matching `Case` in `modDefer.Dispatch` and look for
`Dispatch` lines in the log.

## Accented characters show as `~e`, `fen~xtre`

The sources are ASCII; accents are decoded at run time by `modConfig.A()` from `~x`
style markers. A marker outside `T()` or `A()`, or an unknown marker, is displayed as
is. `validate.py` reports both (`ACCENT`, `MARKER`).

## The desktop path is wrong (log, caches or config not found)

On a OneDrive for Business tenant the desktop is redirected; `Environ("USERPROFILE") &
"\Desktop"` is wrong. The toolkit resolves the desktop through the shell special
folder and OneDrive variables (`modConfig.DesktopPath()`). Set `DesktopPathOverride`
in `config.ini` if the automatic resolution still fails, and check the working folder
`<Desktop>\Macro VBA\Outlook\`.

## config.ini seems ignored

- The file is read once at startup; reload it from *Maintenance > Configuration*.
- A key read directly as a global (`g_...`) before the load returns an empty value: the
  code uses getters, which trigger the load. Do the same in your own additions.
- Files saved with LF-only line endings are handled since v2; older versions required
  CRLF.
- `python validate.py --config config.ini` lists unknown keys (typos).

## QuickMove does not find a folder, or proposes a stale one

The folder index is cached in the working folder. It rebinds by folder identity at
startup and repairs renamed paths by itself; if a large part of the tree changed,
*QuickMove > Reload* rebuilds it. Mailboxes listed in `QuickMoveExcludeStores` are never
indexed.

## Replies are not filed, or filed in the wrong mailbox

`modReplyInPlace` memorises the parent folder of the mail being replied to, in any
mailbox. It never files into system folders (Inbox, Sent, Drafts, Deleted, Junk,
Outbox) nor into public folders. Filing happens when the sent item lands in Sent Items:
on Exchange Online the move is asynchronous, so allow a second or two.

## Excel workbooks: "Excel is empty", hidden windows, errors 462

- A workbook saved while its window was hidden is reopened hidden by Excel ("Excel is
  empty"). The toolkit re-shows the window before every save; if a workbook is already
  in that state, open Excel, *View > Unhide*, save.
- A leftover `EXCEL.EXE` (ghost instance) answers `GetObject` but fails at the first
  real call (error 462). The toolkit tests the instance with a full round-trip and
  ignores ghosts; kill the leftover process if it accumulates.
- Excel busy (a cell being edited, a modal dialog) rejects COM calls; the toolkit
  retries with a growing delay. Finish the edit and retry.

## Mail-tracking dashboard: buttons act on the wrong column

The sheet code reads its column layout from the hidden `_Schema` sheet written by
Outlook at each opening. If the sheet is missing (workbook never reopened by Outlook
since version 3), the historical layout is assumed. Open the dashboard from Outlook
once.

## Toasts do not appear

Toasts are small HTML windows launched as an external process (`mshta`). They are
blocked when launched from the VBA editor on some EDR setups, and they respect the
"skip if foreground" and grouping settings of the inbox alerts. Check the log for
`Notify` lines and test with the editor closed.

## Birthday calendar view is never applied

`BirthdayCalendarAutoShow` defaults to 0. Set it to 1 in `config.ini`; the log shows an
explicit line when the view setup is skipped because of that key.

## Campaign send seems stuck

Since version 3 the send runs in the background: one mail per tick, with the throttle
and batch pauses from `config.ini`. The Excel status bar shows the progress; `CampaignStop`
interrupts; `CampaignResetSend` clears a stale state after a VBA recompilation.

## OneNote commands fail (TYPE_E_CANTLOADLIBRARY)

The OneNote COM type library is not registered correctly on the machine (64-bit
OneNote registry issue). The module compiles but cannot run until IT repairs the
registration; `modOneNoteProbe` produces a verdict report to hand to IT.

## New Outlook

The toolkit does not run in New Outlook, which has no VBA. Use Outlook Classic.
