# Goal: Implement the complete audit remediation

## User Request

Implement the approved implementation plan for all repository findings and GitHub issues #1 through #29.

## Refined Goal

Implement every issue in https://github.com/trsdn/outlook-ax/issues/1 through https://github.com/trsdn/outlook-ax/issues/29 using the dependency-ordered specification at `/Users/torstenmahr/.copilot/session-state/24c7cdc9-72c1-4e23-8680-fcd298411bac/files/plan/architecture-outlook-ax-remediation-1.md`. First land fail-closed safety fixes, then replace the duplicated root CLI/library implementations with one tested SwiftPM core plus an `outlook-ax` executable. Complete the typed CLI/JSON, localization, temporal-model, mail, calendar, CI, and documentation migrations without adding third-party runtime dependencies.

## Acceptance Criteria

- [ ] Criterion 1: GitHub-hosted macOS CI builds the release CLI, runs all Swift tests and CLI subprocess tests, and validates repository/architecture scripts.
- [ ] Criterion 2: `.claude/settings.json` is the effective SessionStart hook configuration, and skill synchronization works when invoked outside the repository root.
- [ ] Criterion 3: No command closes or discards a window that existed before the operation; event detail/editor windows are selected through before/after identities and ambiguity errors.
- [ ] Criterion 4: Mail compose and calendar create fail closed on missing fields, failed AX actions, or failed postconditions; Send and Save are never invoked after a requested-field failure.
- [ ] Criterion 5: `mail inbox --limit` strictly accepts only `1...100`; all malformed limits return a usage error without reaching `Collection.prefix`.
- [ ] Criterion 6: Inbox reads verify the selected Inbox and parse only validated message rows under a verified message-list table.
- [ ] Criterion 7: `Sources/OutlookAX` is the only AX, connection, localization, parser, model, and domain implementation; the root `outlook-ax.swift` duplicate is removed after parity is proven.
- [ ] Criterion 8: `Package.swift` exposes the existing `OutlookAX` library and an `outlook-ax` executable; `make build` and `make install` still produce/install one executable without third-party runtime dependencies.
- [ ] Criterion 9: Passive status/read commands do not launch, activate, or unminimize Outlook; errors distinguish permission denial, Outlook not running, and no accessible windows.
- [ ] Criterion 10: No production command assumes `wins[0]`; view and window selection use semantic classifiers, stable identities, and explicit ambiguity failures.
- [ ] Criterion 11: CLI parsing is testable; explicit help exits `0`, usage errors exit `64`, permission denial exits `77`, unavailable Outlook/windows exit `69`, and each invocation emits exactly one result.
- [ ] Criterion 12: JSON v2 uses typed Codable success/error envelopes with native numbers and booleans, while an explicit documented v1 compatibility path remains available through the 2.x migration.
- [ ] Criterion 13: One shared localization catalog contains fixture-verified command-critical labels for de, en, fr, es, and it; localized UI literals outside that catalog are rejected by tests.
- [ ] Criterion 14: Calendar models separate localized display strings from normalized temporal values with year, all-day range semantics, time-zone source, and explicit unresolved/local-only states.
- [ ] Criterion 15: Calendar parsing supports 12-hour and 24-hour times, never parses hyphenated titles as ranges, keeps categories separate from calendar identity, and passes locale/year/DST fixtures.
- [ ] Criterion 16: `calendar today` returns only the verified current date in chronological order; detail lookup uses stable row references or unique normalized fallback matching and parses location semantically.
- [ ] Criterion 17: Mail folder/calendar selectors preserve account/group identity and reject ambiguous bare names; body reads are faithful; search returns parsed results, no-results, or timeout.
- [ ] Criterion 18: Calendar view switching, timescale, date navigation, and calendar toggling report success only after verified state changes and work across nested menus, months, years, and duplicate names.
- [ ] Criterion 19: README, AGENTS.md, architecture, AX-path, JSON-schema, temporal-contract, and changelog documentation match the final SwiftPM architecture and behavior.
- [ ] Criterion 20: Automated tests include pure parser fixtures, fake AX/process/clock/input tests, form safety tests, localization coverage, JSON schema tests, CLI subprocess tests, architecture guards, and optional non-destructive live smoke instructions.
- [ ] Criterion 21: `make build`, `swift test`, all repository check scripts, and the clean-checkout CI-equivalent gate pass with no unresolved changes or generated tracked artifacts.
- [ ] Criterion 22: Every issue #1-#29 has its acceptance criteria implemented and covered; issue closure must not precede passing verification.

## Scope Boundaries

**In scope:**
- All code, test, build, workflow, script, localization, public model, compatibility, and documentation changes required by issues #1-#29.
- The architecture and migration sequence defined in the approved implementation-plan artifact.
- Small temporary stop-loss changes in the root CLI before its removal.
- Backward-compatible public-library and JSON migration shims.
- GitHub issue updates or closure only after independent verification demonstrates each issue is complete.

**Out of scope:**
- Third-party runtime dependencies.
- New Outlook commands unrelated to issues #1-#29.
- Destructive live Outlook testing, real email/event sending, deletion, draft discard, or closing unrelated user windows.
- Silently inferring missing calendar years, time zones, or ambiguous resource identities.
- Rewriting unrelated git history or modifying unrelated user changes.

## Applicable Project Conventions

**Quality gate command:**
- `make build`
- `swift test`
- `bash scripts/check-repository-layout.sh` after it is introduced
- `bash scripts/check-architecture.sh` after it is introduced
- Run the CI-equivalent aggregate defined by the final workflow before completion.

**Commit convention:**
- Conventional Commits based on repository history: `feat:`, `fix:`, `docs:`, `ci:`, `chore:`, and scoped variants.
- Builder iteration title: `type(scope): [B] description`, at most 72 characters.
- Inspector iteration title: `chore(scope): [I] description`, at most 72 characters.
- Assisted-by trailer required: `Assisted-by: Claude:Sonnet-4.6` for Builder commits and `Assisted-by: Claude:Haiku-4.5` for Inspector commits.
- Also include `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>` and `Copilot-Session: 7e536eb6-0929-432f-a58d-d0c08475dc8c`.

**Guidelines:**
- `AGENTS.md`
- `/Users/torstenmahr/.copilot/session-state/24c7cdc9-72c1-4e23-8680-fcd298411bac/files/plan/architecture-outlook-ax-remediation-1.md`
- GitHub issue bodies #1-#29

**Rules:**
- Preserve the no-third-party-runtime-dependency and one-executable distribution goals.
- Use localized shared labels for every UI match and locale-neutral/English-normalized machine output.
- Use `ok()`/`fail()` semantics only until the typed runner replaces them; never add silent failures.
- Never press Send, Delete, Discard, or broad Close in discovery, fixtures, hosted CI, or optional live smoke tests.
- Date/time keyboard simulation is allowed only in an explicitly owned interactive editor and must be postcondition-verified.
- The existing single-file/no-Package guidance in `AGENTS.md` is obsolete under issue #28 and the approved plan; update it rather than preserving duplicated implementations.
