# Calendar Temporal Contract — outlook-ax

This document specifies the normalized temporal values in `CalendarTemporal`
and the rules that govern them. These values are separate from the localized
display strings returned in `CalendarEvent.date`, `.start`, and `.end`
(which are deprecated display fields).

## Goals

- Never silently infer a missing year, ambiguous DST wall time, or unconfirmed time zone.
- Provide machine-readable ISO 8601 dates and, where safe, RFC 3339 timestamps.
- Allow consumers to detect when temporal data is incomplete rather than fabricated.

## Field reference

| Field               | Type     | When present                                      |
|---------------------|----------|---------------------------------------------------|
| `isoDate`           | string   | Always when date can be resolved (YYYY-MM-DD).    |
| `localStartTime`    | string   | Timed events only (HH:MM, 24-hour).               |
| `localEndTime`      | string   | Timed events only (HH:MM, 24-hour).               |
| `startTimestamp`    | string   | Resolved timed events only (RFC 3339).            |
| `endTimestamp`      | string   | Resolved timed events only (RFC 3339).            |
| `isAllDay`          | boolean  | Always.                                           |
| `year`              | integer  | When the year is explicitly known (not inferred). |
| `resolution`        | string   | Always. See resolution values.                    |
| `unresolvedReason`  | string   | Only when resolution = "unresolved".              |
| `timeZoneSource`    | string   | Always. See TZ source values.                     |
| `timeZoneIdentifier`| string   | When resolution = "resolved".                     |

## Resolution values

| Value       | Meaning                                                               |
|-------------|-----------------------------------------------------------------------|
| `resolved`  | Date, local time, and time zone are all known. RFC 3339 is available.|
| `localOnly` | Date and local time are known, time zone is uncertain.               |
| `unresolved`| Date or time could not be parsed. Check `unresolvedReason`.          |

## Unresolved reasons

| Value          | Meaning                                                     |
|----------------|-------------------------------------------------------------|
| `missingYear`  | No year context was available; year cannot be inferred.    |
| `missingTime`  | Event has no time component and no all-day flag.            |
| `ambiguousDST` | Wall time is ambiguous during a DST transition.             |
| `parseError`   | The date/time string did not match any supported format.    |

## Time zone sources

| Value           | Meaning                                                      |
|-----------------|--------------------------------------------------------------|
| `axLabel`       | Time zone was explicitly identified in the AX tree.          |
| `systemFallback`| System local time zone used (may differ from event's TZ).    |
| `absent`        | No time zone information available (all-day events, etc.).   |

## Rules

### Year

- The year MUST be explicitly known (from the AX tree or from the UI anchor date).
- The year MUST NOT be inferred from the current year.
- Events near a year boundary (December/January) are especially sensitive.
- If the year cannot be determined, `resolution` is `"unresolved"` and
  `unresolvedReason` is `"missingYear"`.

### All-day events

- All-day events have `isAllDay: true` and MUST NOT have `startTimestamp`
  or `endTimestamp` (no synthetic midnight timestamps).
- The `isoDate` is the start date (YYYY-MM-DD).
- For multi-day all-day events, `endTimestamp` is the **exclusive** end date
  (start date + n days), not a synthetic midnight timestamp.
- `localStartTime` and `localEndTime` are always absent for all-day events.

### Timed events

- `localStartTime` and `localEndTime` are in 24-hour format (HH:MM).
- Both 12-hour (AM/PM) and 24-hour time strings from the UI are normalized
  to 24-hour before being stored.
- If both time and time zone are known, `startTimestamp` is a valid RFC 3339
  timestamp (e.g. `"2026-04-15T09:00:00+02:00"`).
- If the time zone is unknown, `startTimestamp` is absent and
  `resolution` is `"localOnly"`.

### DST

- Ambiguous wall times during a DST transition remain `"unresolved"` with
  `unresolvedReason: "ambiguousDST"`.
- The system MUST NOT choose arbitrarily between the two possible UTC offsets.

### Hyphenated titles

- Hyphenated text in event titles (e.g. "Stand-up") MUST NOT be parsed as
  a time range. Only verified AX time/date child elements are used as input
  to temporal parsing.

### Categories vs. calendar

- `categories` is a separate field containing category labels assigned to
  the event.
- `calendar` (CalendarEvent.calendar) is populated only from verified calendar
  identity and MUST NOT contain category labels.
- When the calendar cannot be determined, `calendar` is the empty string.

## Backward-compatible display fields

The following `CalendarEvent` fields are retained for source compatibility
and are deprecated for machine use. Use `CalendarTemporal` for new consumers:

| Deprecated field | Replacement           | Notes                                     |
|------------------|-----------------------|-------------------------------------------|
| `date`           | `temporal.isoDate`    | Localized display header; language varies |
| `start`          | `temporal.localStartTime` | "HH:MM" or localized all-day label   |
| `end`            | `temporal.localEndTime`   | "HH:MM" or localized date            |
