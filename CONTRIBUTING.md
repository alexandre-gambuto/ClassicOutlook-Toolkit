# Contributing

Contributions are welcome, within the constraints that shaped the project: the code
must keep running on a locked-down corporate workstation, without admin rights, without
UserForms, without external components.

## Ground rules

- **Outlook Classic, 64-bit only.** Every `Declare` uses `PtrSafe` and `LongPtr`; no
  `#If VBA7` branches.
- **Late binding everywhere.** `CreateObject` for Excel, Word, OneNote, Scripting,
  WScript, MSXML. Never add a reference under *Tools > References*.
- **No UserForms in the Outlook project.** Dialogs go through `modUi`,
  `modMsgBoxCustom` and the toast notifications.
- **No process spawning from Outlook** (no `cmd`, `powershell`, `reg.exe`, `SendKeys`).
  EDR agents flag it, and the toolkit has pure-VBA alternatives (`GetAllSettings`,
  `Shell.Application`, the Word editor of an inspector).
- **Everything environment-specific lives in `config.ini`**, read through the getters of
  `modConfig` (never through the `g_` globals directly). Add new keys to
  `config.ini.example`; `python validate.py --gen-config` lists them.
- **No hard-coded French in the code.** UI strings use `T(fr, en)`; accented characters
  use the `~` markers decoded by `A()` (`~e` e acute, `~f` e grave, `~a` a grave, `~x` e
  circumflex, `~c` c cedilla, `~u` u circumflex, `~U` u grave, `~A` a circumflex,
  `~y` i with diaeresis, `~E` E acute, `~i` i circumflex, `~o` o circumflex, `~I` I
  circumflex). Functional strings (category names, status values, DASL, config keys)
  and log lines are never translated.
- **Identifiers in English.** Module names `modXxx` / `clsXxx`; a module never carries
  the name of one of its procedures. Public ribbon procedures are a stable API: never
  rename them without regenerating the `.exportedUI` files.
- **Logging through `modDebugLog`** (`DError`, `DInfo`, `DDebug`, `DTrace` with the
  module short name). No `Debug.Print`, no diagnostic `MsgBox`. Never log mail bodies
  or secrets.
- **Deferred execution through `modDefer`** only. `Application.OnTime` does not exist
  in Outlook. Every armed key needs a matching `Case` in `Dispatch`.
- **Event handlers never raise.** `On Error Resume Next` at the top, errors logged.

## Source file format

- One `.txt` file per module, starting directly with `Option Explicit`, no
  `Attribute` or `VERSION` lines, ASCII only, CRLF line endings.
- Module-level declarations (`Const`, `Dim`, `Declare`, `Type`, `Enum`) before the first
  procedure. VBA does not compile otherwise, and the failure can crash Outlook at
  startup.
- Never name a local variable `t` or `a` in a procedure that calls `T()` or `A()`:
  VBA is case-insensitive and the local hides the helper.
- Collections such as `Folder.Folders` or `Folder.Items` are assigned to a variable
  before a loop, never re-evaluated at each iteration.
- Each module has `MOD_NAME` (6 to 10 characters) and `MOD_VERSION` (`vN.N - YYYY-MM-DD`)
  exposed by `ModVersion()`; `python validate.py --gen-versions` regenerates the
  version inventory of `modMaintenance`.

## Before submitting

1. `python validate.py` on the changed files: zero error. Warnings should be
   understood, not silenced.
2. *Debug > Compile* in the Outlook VBA editor, then a restart of Outlook with the
   editor closed.
3. A `CHANGELOG.txt` entry at the top, newest first:
   `[YYYY-MM-DD] modXxx vN - short description`, followed by bullet points. Keep the
   history below untouched.
4. Complete files only. No diffs, no fragments: the reviewer pastes the whole module.
5. New `config.ini` keys documented in `config.ini.example` and
   `docs/CONFIGURATION.md`.

## Testing without a debugger

The environment has no debugger session worth trusting for events. The project method
is: write a small probe first (a module with one `Sub` that exercises the mechanism and
logs the verdict), run it from the ribbon with the editor closed, read the log, then
build the feature on what the probe proved. Keep the probe in the repository if the
mechanism is fragile (`modOneNoteProbe` is an example).
