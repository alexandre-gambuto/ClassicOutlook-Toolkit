# Installation guide

The toolkit is pasted by hand into the Outlook VBA editor. This is deliberate: the
target environment blocks the VBA Project Object Model, so `.bas` import is not
available and no installer is provided. Count one to two hours for a first install.

## 1. Prerequisites

- Outlook Classic (desktop), 64-bit Office, Windows.
- Macros enabled for Outlook: *File > Options > Trust Center > Trust Center Settings >
  Macro Settings*, "Notifications for all macros" or "Enable all macros". A managed
  tenant may impose its own policy.
- Excel with VBA if you use the time-tracking referential, the mail-tracking dashboard,
  the campaigns or the analytics tool.
- Optional: Python 3 to run `validate.py`.

## 2. Validate the sources (optional, recommended)

```
python validate.py                 # every .txt of the folder
python validate.py modQuickMove.txt
python validate.py --config config.ini
```

Zero error is required before pasting. Warnings are informative. The validator also
reports config keys that the code does not know (typos).

## 3. Paste the modules

1. `Alt+F11` opens the VBA editor. In the *Project* pane, expand the Outlook project.
2. **Infrastructure first**, so that the other modules compile:
   `modConfig`, `modDebugLog`, `modCommon`, `modUi`, `modDefer`, `modSmtp`,
   `modExcelBridge`, `modContactCommon`, `modInspectorQueue`, `modNotify`,
   `modMsgBoxCustom`, `modStoreContext`.
3. For each remaining `.txt` file: *Insert > Module*, paste the file content, then rename
   the module to the file name (properties window, `F4`). A module name must match the
   file name exactly: other modules call it by that name.
4. Class modules: *Insert > Class Module* for `clsAppEvents`, `clsApptWatcher`,
   `clsInboxWatcher`, `clsExplorerWatcher`.
5. `ThisOutlookSession`: open the existing object and paste the content of
   `ThisOutlookSession.txt` over it.
6. Files start with `Option Explicit` and carry no `Attribute` line; paste them as they
   are.

Encoding: the files are ASCII; accented characters in the UI are produced at run time
by the `A()` decoder from `~e`, `~x` style markers. Do not "fix" them in the editor.

## 4. Configuration

Shortcut: restart Outlook once with the modules pasted and no `config.ini`. Eight
seconds after startup the quick start guide offers to check the environment, ask five
questions and write a minimal file in the right place (`Maintenance > Configuration >
5` runs it again). The manual way follows.

1. Create the working folder used by the toolkit, by default
   `<Desktop>\Macro VBA\Outlook\` (the desktop path is resolved through OneDrive
   redirection when present).
2. Copy `config.ini.example` to `config.ini` in that folder and edit it: paths of the
   referential workbooks, project code prefix and patterns, internal domains, colours,
   `Language=FR` or `EN`. Every key is documented in `docs/CONFIGURATION.md`.
3. Keys you do not set fall back to the defaults coded in `modConfig`. The file is
   optional but strongly recommended.
4. Line endings: CRLF, CR or LF are all accepted.

## 5. Referential workbooks

- **Time-tracking workbook** (`TimeTrackingFilename`, sheet `TimeTrackingSheet`): one
  row per project or contract, headers on row 1, columns read by position
  (opportunity, contract, project code, folder name, partner, then hours per month).
- **Mail-tracking workbook** (`MailTrackingFilename`): created automatically at first
  use. Paste the content of `Suivi_des_mails__Feuille_Suivi.txt` into the VBA module of
  its tracking sheet to get the action buttons; the sheet code reads its column layout
  from the hidden `_Schema` sheet written by Outlook.
- **Campaign workbook**: created by the `CampaignCreateSheet` command.

## 6. Compile and start

1. *Debug > Compile Project*. Fix any error before going further: one non-compiling
   module makes the deferred-action scheduler unusable and can crash Outlook a few
   seconds after startup.
2. Close the VBA editor. Event tests with the editor open are not representative.
3. Restart Outlook completely. `Application_Startup` binds the events and schedules
   the startup tasks.
4. Run the maintenance command *Check bindings* (ribbon Maintenance, or `AssertBindings`
   from the Immediate window): six bindings must be alive.

## 7. Ribbon

Import the `*.exportedUI` files via *File > Options > Customize Ribbon > Import/Export >
Import customization file*. The buttons call the public macros by name; do not rename
the public procedures without regenerating the exports.

## 8. Upgrading from an earlier version

- **Version 3 adds four modules** (`modSmtp`, `modExcelBridge`, `modContactCommon`,
  `modInspectorQueue`): paste them before updating the others.
- Replace each existing module by the new content (select all, paste). Do not keep two
  copies of a module under different names.
- New `config.ini` keys are listed in the "PACK v3" block of `config.ini.example`; the
  defaults reproduce the previous behaviour, except `BirthdayCalendarAutoShow` which you
  set to `1` yourself if you want the birthday calendar view applied at startup.
- The QuickMove folder cache is migrated transparently to its new format at the first
  save. The mail-tracking workbook receives its hidden `_Schema` sheet at the first
  opening by Outlook.
- Legacy technical appointment `QuickMove_BackgroundReindex` (pre-v3) is deleted
  automatically when its reminder fires; you may also delete it by hand.
- Compile, close the editor, restart Outlook.

## 9. Uninstall

Delete the modules from the VBA editor (right-click > Remove, "No" to export), remove
the ribbon customizations, delete the working folder. Nothing else is written outside
`HKCU\Software\VB and VBA Program Settings`.
