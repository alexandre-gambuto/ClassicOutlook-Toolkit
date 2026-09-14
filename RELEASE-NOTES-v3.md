# Release notes, version 3

Version 3 is a full review of the code base: every module was read, and the resulting
fixes, performance work, refactoring and cleanup were delivered as one pack. Module
version constants were kept as they were; the pack is the version.

## What users see

- **Replies are filed in shared mailboxes too.** A reply to a mail filed in another
  mailbox now joins its folder; previously anything outside the default mailbox stayed
  in Sent Items.
- **Campaigns no longer freeze Outlook.** The throttled send runs in the background,
  one mail per tick, with a new `CampaignStop` command.
- **The folder index heals itself.** Folders are rebound by identity at startup; a
  renamed or moved folder keeps working and the cache is rewritten. The weekly reindex no
  longer needs a technical appointment in the calendar.
- **Faster sends.** The urgent-banner check no longer reads the whole HTML body of every
  outgoing mail.
- **Greeting from the directory.** The first name of a reply greeting comes from the
  company directory when available.
- **Dialogs stay in front.** Custom message boxes are raised to the foreground and their
  countdown pauses while Outlook is in the background.
- **Configurable vocabulary and categories.** Reply-suspicion keywords, category names
  (personal, internal task, commercial) and referential labels moved to `config.ini`.
- **Accents.** Seven new accent markers; twenty-seven labels that displayed the wrong
  accent are fixed.
- **Registry backup covers every setting branch**, exported in pure VBA (no shell).
- **Mail-tracking dashboard**: the workbook reads its column layout from a hidden
  `_Schema` sheet written by Outlook, so a column change on one side can no longer
  trigger the wrong action on the other.

## Under the hood

- Four new infrastructure modules: `modSmtp` (address resolution with a session
  cache), `modExcelBridge` (Excel COM lifecycle), `modContactCommon` (shared contact
  helpers), `modInspectorQueue` (inspectors kept alive for deferred passes).
- One code scanner in QuickMove instead of seven copies of the same three passes; one
  folder picker shared by GoTo and folder merge; one project constructor and one
  five-field wizard in the tagger.
- COM collections evaluated once per loop; folder activity computed with MAPI
  restrictions instead of sorts; block reads of Excel ranges; a filtered flag cleanup.
- `validate.py`: unknown accent markers, wrong circumflex forms, duplicate `Dim`,
  collections re-evaluated in loops, `--config`, `--gen-versions`, `--gen-config`.
- `config.ini.example` completed with 92 keys that the code read but the template did
  not list.

## Upgrade notes

- Paste the four new modules before updating the others.
- New keys have defaults that reproduce the previous behaviour; set
  `BirthdayCalendarAutoShow=1` if you want the birthday view at startup.
- Deferred-action keys added: `Reindex`, `CampaignNext`.
- Not included in this version, on purpose: repair of the tracker through
  `AdvancedSearch`, an OLEDB reader for the referential, log buffering, a `GetTable`
  based analytics scan, the writing-assistant bridge (left untouched).

The detailed list is the `[2026-09-02] PACK v3` entry of `CHANGELOG.txt`.
