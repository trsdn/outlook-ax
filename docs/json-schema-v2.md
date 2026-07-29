# JSON Schema v2 — outlook-ax

All responses produced by `outlook-ax --json` follow the JSON v2 envelope format
described in this document. Existing consumers that rely on the v1 flat-map format
can request it explicitly with `--json-version 1`.

## Version selection

| Flag                         | v1.1 behavior | v2.0 behavior |
|------------------------------|---------------|---------------|
| `--json`                     | v1 (default)  | v2 (default)  |
| `--json` `--json-version 2`  | v2            | v2            |
| `--json` `--json-version 1`  | v1            | v1 (compat)   |

v1 remains available through all 2.x releases.

## v2 Success Envelope

```json
{
  "schemaVersion": 2,
  "ok": true,
  "command": "<command-name>",
  "data": { ... }
}
```

| Field           | Type    | Description                              |
|-----------------|---------|------------------------------------------|
| `schemaVersion` | integer | Always `2` for v2 responses.            |
| `ok`            | boolean | `true` on success.                       |
| `command`       | string  | Command name (e.g. `"notifications"`).  |
| `data`          | object  | Command-specific payload (see below).   |

## v2 Error Envelope

```json
{
  "schemaVersion": 2,
  "ok": false,
  "command": "<command-name>",
  "error": {
    "code": "<machine-code>",
    "message": "<human-readable>",
    "details": {}
  }
}
```

| Field             | Type    | Description                                     |
|-------------------|---------|-------------------------------------------------|
| `ok`              | boolean | Always `false` for errors.                     |
| `error.code`      | string  | Machine-readable error code (see below).        |
| `error.message`   | string  | Human-readable explanation.                     |
| `error.details`   | object  | Optional structured context.                    |

### Error codes

| Code                          | Exit code | Meaning                                    |
|-------------------------------|-----------|---------------------------------------------|
| `accessibilityPermissionDenied` | 77      | TCC/Accessibility access not granted.       |
| `outlookNotRunning`           | 69        | Outlook process not found.                  |
| `noAccessibleWindows`         | 69        | Outlook running but no AX windows.          |
| `notInCalendarView`           | 69        | Outlook is not showing the Calendar tab.    |
| `notInMailView`               | 69        | Outlook is not showing the Mail tab.        |
| `viewSwitchFailed`            | 1         | UI view switch timed out or was rejected.   |
| `eventsTableNotFound`         | 69        | Calendar list-view table not found.         |
| `eventRowNotFound`            | 1         | Named event not found in the list.          |
| `ambiguousEvent`              | 1         | Multiple events matched; cannot disambiguate.|
| `messageListNotFound`         | 69        | Mail message-list table not found.          |
| `inboxNotSelected`            | 69        | Inbox not currently selected.               |
| `folderAmbiguous`             | 1         | Folder name matches multiple accounts.      |
| `formFieldNotFound`           | 1         | A required form field was absent.           |
| `formFieldWriteFailed`        | 1         | AX write to a form field failed.            |
| `sendBlockedByFieldFailure`   | 1         | Send/Save blocked after a field failure.    |

## Stdout contract

Each invocation emits **exactly one** JSON document on stdout. Diagnostic
messages use stderr and never appear in the JSON document.

## Native types

v2 preserves native JSON types throughout:
- Counts are **integers** (not strings): `"count": 3`
- Flags are **booleans** (not strings): `"isAllDay": true`, `"running": false`

The v1 format stringified all values; v2 does not.

## Command payloads

### `notifications`

```json
{
  "schemaVersion": 2,
  "ok": true,
  "command": "notifications",
  "data": {
    "count": 3
  }
}
```

### `status`

```json
{
  "schemaVersion": 2,
  "ok": true,
  "command": "status",
  "data": {
    "notifications": 0,
    "running": true,
    "view": "calendar",
    "window": "Calendar"
  }
}
```

### `mail inbox`

```json
{
  "schemaVersion": 2,
  "ok": true,
  "command": "mail.inbox",
  "data": [
    {
      "date": "Mon Jul 21",
      "from": "Alice",
      "preview": "Hi there…",
      "subject": "Hello"
    }
  ]
}
```

### `calendar today`

```json
{
  "schemaVersion": 2,
  "ok": true,
  "command": "calendar.today",
  "data": [
    {
      "calendar": "Work",
      "categories": [],
      "date": "Wednesday, 15 April",
      "end": "10:00",
      "isAllDay": false,
      "myResponse": "accepted",
      "organizer": "Bob",
      "start": "09:00",
      "status": "Busy",
      "title": "Team Meeting"
    }
  ]
}
```

## Migration guide (v1 → v2)

1. Add `--json-version 2` to all invocations to opt into v2 during v1.1.
2. Update consumers to read `data` from the success envelope instead of the
   flat map. Check `ok === true` before reading `data`.
3. Update notification count consumers to read `data.count` as an integer.
4. Update status consumers to read `data.running` as a boolean.
5. When v2.0 ships, remove the `--json-version` flag (bare `--json` → v2).
6. Keep `--json-version 1` available for a one-release grace period.
