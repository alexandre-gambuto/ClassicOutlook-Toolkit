# config.ini reference

All environment-specific values live in `config.ini`, read once at startup by
`modConfig` and exposed through typed getters. `config.ini.example` is the commented
template; this page lists every key the code reads, grouped by the module that reads
it, with the default applied when the key is absent.

Format: `Key=value`, one per line; `;` or `#` starts a comment; `[Section]` lines are
ignored; CRLF, CR and LF line endings are accepted. Booleans accept `1/0`,
`true/false`, `yes/no`, `oui/non`. Lists use `|` as separator.

After editing: ribbon *Maintenance > Configuration > Reload*, or restart Outlook.

`python validate.py --config config.ini` reports keys unknown to the code.

Generated from the sources with `python validate.py --gen-config`; descriptions were
added by hand for the most useful keys.

## Core (`modConfig`)

| Key | Default | Description |
|-----|---------|-------------|
| `AnthropicApiKey` | *(empty)* | API key stored locally for the assistant bridge. Never commit config.ini. |
| `ArchiveSubfolderName` | `_Archives` | Name of the archive subfolder under each root. |
| `AutoArchiveFolders` | *(empty)* | Root folders whose inactive subfolders are archived (pipe-separated paths). |
| `AutoArchiveMonths` | `1` | Inactivity threshold in months. |
| `BackupRetentionMonths` | `12` |  |
| `CampaignBatchSize` | `50` | Mails per batch before a long pause. |
| `CampaignCap` | `200` | Maximum mails per send. |
| `CampaignPauseSec` | `180` | Long pause between batches, in seconds. |
| `CampaignStepSec` | `20` | Seconds between two campaign mails. |
| `CommercialCategory` | `Commercial` | Category applied to opportunities. |
| `ConfidBandColorConf` | `#C65911` |  |
| `ConfidBandColorPersonal` | `#2E75B6` |  |
| `ConfidBandColorPrivate` | `#BF8F00` |  |
| `ConfidBandColorUrgent` | `#D13438` |  |
| `ConfidTextColorConf` | `#C00000` |  |
| `ConfidTextColorPersonal` | `#1F4E79` |  |
| `ConfidTextColorPrivate` | `#9C6500` |  |
| `ConfidTextColorUrgent` | `#B00020` |  |
| `ContractLabelDefault` | `Operations commerciales et support` | Contract label written when a project has no contract. |
| `ContractLabelOpportunity` | `Opportunite CRM` | Contract label written in the referential for an opportunity. |
| `ContractNumberLength` | `5` | Number of digits after the underscore in a contract code. |
| `ContractYearPrefix` | `202` | Leading characters of contract codes (e.g. 20). |
| `CrossCodeLength` | `1` | Length of a cross-cutting code. |
| `CrossPrefix` | `TR` | Prefix of cross-cutting (non-project) task codes. |
| `DefaultAppointmentDuration` | `15` | Default duration (minutes) of new appointments. |
| `DefaultLogLevel` | `1` | Log level at startup: ERROR, INFO, DEBUG or TRACE. |
| `ExternalMeetingCategory` | `Reunion externe` |  |
| `InboxAlertDefaultColor` | `#1565C0` |  |
| `InboxAlertDefaultSound` | *(empty)* |  |
| `InboxAlertDurationMs` | `5000` | Alert display time. |
| `InboxAlertEnabled` | `False` | Master switch of per-mailbox new-mail alerts. |
| `InboxAlertGroupMs` | `1500` | Bursts of mails within this delay are grouped in one toast. |
| `InboxAlertSkipIfForeground` | `False` | No alert when Outlook is already the foreground window. |
| `InboxAlertUnreadOnly` | `True` | Alert only for unread items. |
| `InternalDomains` | `domaine1.fr\|domaine2.fr` | Mail domains considered internal (pipe-separated). Drives meeting colouring and reply suspicion. |
| `InternalMeetingCategory` | `Reunion interne` |  |
| `InternalTaskCategory` | `Interne` | Category applied to cross-cutting tasks. |
| `KeywordsCap` | `3` |  |
| `Language` | `FR` | UI language: FR or EN. Every dialog string goes through T(fr, en). |
| `MailTrackingFilename` | `Suivi des mails.xlsx` | File name of the mail-tracking dashboard workbook. |
| `MailTrackingPath` | *(empty)* |  |
| `MailTrackingSheet` | `Suivi` | Tracking sheet name. |
| `MeetingDuration` | `15` |  |
| `OneNoteNotebook1` | `Bloc-notes 1` |  |
| `OneNoteNotebook2` | `Bloc-notes 2` |  |
| `OneNoteNotebook3` | `Bloc-notes 3` |  |
| `OpportunityNumberLength` | `4` | Number of digits at the end of an opportunity code. |
| `OpportunityTypes` | `OP` | Opportunity code types (pipe-separated, e.g. OT|DT). |
| `PersonalCategory` | `Personnel` | Category name of personal appointments (excluded from tagging and colouring, set private). |
| `PersonalCategoryAlt` | `Perso` | Short alternative accepted for PersonalCategory. |
| `ProjectCodeLength` | `4` | Number of digits in a project code. |
| `ProjectPrefix` | `DOSS` | Prefix of project codes (e.g. DV). Used by tag detection and QuickMove. |
| `ReplyDelayDays` | `7` |  |
| `ShiftMinutes` | `15` |  |
| `SignatureBookmark` | `_MailAutoSig` | Marker that starts the signature block (stripped before sending to an assistant). |
| `SignatureTextMarker` | *(empty)* |  |
| `SpecialContractCategory` | *(empty)* | Category applied when SpecialContractCode matches. |
| `SpecialContractCode` | *(empty)* | Contract code that forces a specific Outlook category. |
| `SuspicionKeywords` | *(empty)* | Words that mark a sent mail as awaiting a reply (pipe-separated). Empty = built-in French list. |
| `SuspicionThreshold` | `3` |  |
| `TagPageSize` | `12` |  |
| `TimeTrackingFilename` | `Suivi de temps.xlsm` | File name of the project/contract referential workbook (in the working folder). |
| `TimeTrackingPath` | *(empty)* |  |
| `TimeTrackingSheet` | `Temps par mois` | Sheet of the referential read by the tagger. |
| `ToastColorError` | `#C62828` | Border colour of error toasts. |
| `ToastColorFiled` | `#2E7D32` | Border colour of the "filed" toast (hex). |
| `ToastColorInfo` | `#1565C0` | Border colour of information toasts. |
| `ToastColorKept` | `#1565C0` |  |
| `ToastColorSuccess` | `#2E7D32` | Border colour of success toasts. |
| `ToastColorText` | `#1565C0` |  |
| `ToastColorWarn` | `#EF6C00` | Border colour of warning toasts. |
| `ToastDurationMs` | `3000` | Toast display time in milliseconds. |
| `ToastEnabled` | `True` |  |
| `ToastKeptInSent` | `False` |  |
| `UpShift` | `AntiChevDecale` |  |
| `UpTag` | `ProjectTag` |  |
| `VBADir` | *(empty)* |  |
| `YearExclusionMax` | `2032` | See YearExclusionMin. |
| `YearExclusionMin` | `2023` | Four-digit numbers between Min and Max are treated as years, not codes. |

## `modAppointmentPrivacy`

| Key | Default | Description |
|-----|---------|-------------|
| `PrivateOrganizers` | *(empty)* | Organizers (SMTP, pipe-separated) whose appointments are set to private. |
| `PrivateStripCategories` | `auto` | auto = remove automatic categories on private appointments. |

## `modContactBirthday`

| Key | Default | Description |
|-----|---------|-------------|
| `AnniversaryEnabled` | `True` |  |
| `AnniversaryFolders` | `*` |  |
| `AnniversaryLabel` | *(empty)* |  |
| `AnniversaryReminderDaysBefore` | *(empty)* |  |
| `BirthdayBodyDetails` | `True` |  |
| `BirthdayCalendarAutoShow` | `False` | Apply the birthday calendar view at startup (1/0). Defaults to 0. |
| `BirthdayCalendarKeepDefaultOf` | *(empty)* |  |
| `BirthdayCalendarKeepPaths` | *(empty)* |  |
| `BirthdayCalendarName` | *(empty)* |  |
| `BirthdayCalendarOverlay` | `True` |  |
| `BirthdayCalendarPath` | *(empty)* | Outlook path of the birthday calendar. |
| `BirthdayCalendarUncheckOthers` | `True` |  |
| `BirthdayCalendarWeekView` | `True` |  |
| `BirthdayCategory` | *(empty)* |  |
| `BirthdayCategoryColor` | `3` |  |
| `BirthdayLabel` | *(empty)* |  |
| `BirthdayPrivate` | `False` |  |
| `BirthdayReminderDaysBefore` | `1` |  |
| `BirthdayReminderHour` | `9` |  |
| `BirthdaySubjectFormat` | `{name} - {label} - {date}` |  |
| `ContactSkipFolders` | *(empty)* | Contact folder name patterns ignored by the contact tools (pipe-separated, wildcards allowed). |
| `NameDayEnabled` | `True` |  |
| `NameDayFolders` | `*Amis*\|*Famille*` |  |
| `NameDayLabel` | *(empty)* |  |
| `NameDayReminderDaysBefore` | *(empty)* |  |
| `UnknownBirthYear` | `1604` |  |
| `UnknownBirthYearMax` | *(empty)* |  |

## `modContactExport`

| Key | Default | Description |
|-----|---------|-------------|
| `ContactCsvSeparator` | `;` |  |
| `ContactExportDir` | *(empty)* |  |
| `ContactExportPhotos` | `True` |  |

## `modContactFromMail`

| Key | Default | Description |
|-----|---------|-------------|
| `ContactFromMailCategory` | *(empty)* |  |
| `ContactFromMailDefaultFolder` | *(empty)* |  |
| `PhoneDefaultCountryCode` | `33` | Country code assumed for national phone numbers. |

## `modContactImport`

| Key | Default | Description |
|-----|---------|-------------|
| `ContactImportAllowMove` | `True` |  |
| `ContactImportCreateFolders` | `True` |  |

## `modContactNormalize`

| Key | Default | Description |
|-----|---------|-------------|
| `AddressCityUppercase` | `True` |  |
| `AddressCountryUppercase` | `True` |  |
| `AddressDefaultCountry` | *(empty)* |  |
| `AddressLowercaseWords` | *(empty)* |  |
| `AddressUppercaseWords` | *(empty)* |  |
| `ClearCategoriesPrompt` | `True` |  |
| `ContactFileAsFormat` | `LastFirst` |  |
| `NameDayExtra` | *(empty)* |  |
| `NameParticles` | *(empty)* |  |
| `NormalizeAddresses` | `True` |  |
| `NormalizeEmails` | `True` |  |
| `NormalizeNameDay` | `True` |  |
| `NormalizeNames` | `True` |  |
| `NormalizePhones` | `True` |  |
| `PhoneOutputSpaced` | `False` |  |

## `modContactRelation`

| Key | Default | Description |
|-----|---------|-------------|
| `RelationScanExcludeDomains` | *(empty)* |  |
| `RelationScanMonths` | `48` |  |
| `RelationScanStores` | *(empty)* |  |
| `RelationScanThresholdDays` | `180` |  |

## `modFolderRename`

| Key | Default | Description |
|-----|---------|-------------|
| `FolderRenameRoot` | `DV` | Root folder of the mass-rename tool. |

## `modHideBcc`

| Key | Default | Description |
|-----|---------|-------------|
| `HideBccOnCompose` | `True` | Hide the Bcc field on every new compose window. |

## `modInboxAlert`

| Key | Default | Description |
|-----|---------|-------------|
| `InboxAlertColor` | *(empty)* |  |
| `InboxAlertSound` | *(empty)* |  |
| `InboxAlertSound1` | *(empty)* | Sound file for that mailbox (empty = none). |
| `InboxAlertStore` | *(empty)* |  |

## `modMailAnalytics__frmConfig_Code`

| Key | Default | Description |
|-----|---------|-------------|
| `AddressMode` | *(empty)* |  |
| `ColorIndex` | *(empty)* |  |
| `ColorRGB` | *(empty)* |  |
| `DateMin` | *(empty)* |  |
| `DateMinEnabled` | *(empty)* |  |
| `Domain` | *(empty)* |  |
| `EmailUser` | *(empty)* |  |
| `Keyword` | *(empty)* |  |

## `modMaintenance`

| Key | Default | Description |
|-----|---------|-------------|
| `KnownModules` | *(empty)* |  |

## `modMeetingToAppointment`

| Key | Default | Description |
|-----|---------|-------------|
| `PersonalSlotToast` | `True` | Toast when a meeting is converted back to a personal slot. |
| `TeamsPreCleanEnabled` | `True` | Remove the Teams block when a new appointment opens. |

## `modMigrateTags`

| Key | Default | Description |
|-----|---------|-------------|
| `TagMigrationMap` | *(empty)* | Legacy to current user-property names for the one-shot tag migration (Old>New;Old2>New2). |

## `modQuickMove`

| Key | Default | Description |
|-----|---------|-------------|
| `NewFolderCreateMissingParents` | `0` | Create intermediate folders when a target folder is created on the fly. |
| `QuickMoveExcludeStores` | *(empty)* | Mailboxes excluded from the QuickMove index (display names, pipe-separated). |

## `modReplyGreeting`

| Key | Default | Description |
|-----|---------|-------------|
| `GreetingEnabled` | `True` | Automatic greeting line on single-recipient replies. |
| `GreetingWord` | *(empty)* | Greeting word (empty = Bonjour / Hello depending on Language). |

## `modSearch`

| Key | Default | Description |
|-----|---------|-------------|
| `ViewCompact` | `Compacter` | Compact view setting applied to the calendar window. |
| `ViewPreview` | `Aper` | Preview pane setting applied to the calendar window. |

## `modSendToLLM`

| Key | Default | Description |
|-----|---------|-------------|
| `LLM_AutoEnter` | `True` |  |
| `LLM_DelayChatGPTMs` | *(empty)* |  |
| `LLM_DelayClaudeMs` | *(empty)* |  |
| `LLM_DelayMistralMs` | *(empty)* |  |
| `LLM_PromptPrefix` | *(empty)* |  |
