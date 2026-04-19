---
name: changelog
description: Maintain CHANGELOG.md and create GitHub releases. Use when shipping new features, fixing bugs, or cutting a release. Triggers on "update changelog", "new release", "cut a release", "bump version", or after completing a feature/fix.
---

# Changelog & Release Workflow

## Format

This project uses [Keep a Changelog](https://keepachangelog.com/) with
[Semantic Versioning](https://semver.org/).

### Version Bumping Rules

| Change Type | Bump | Example |
|---|---|---|
| Breaking CLI interface change (renamed command, changed JSON schema) | **major** | 1.0.0 → 2.0.0 |
| New command, new flag, new L10n language, new feature | **minor** | 1.0.0 → 1.1.0 |
| Bug fix, L10n correction, doc update, internal refactor | **patch** | 1.0.0 → 1.0.1 |

### Entry Categories

Use exactly these headings under each version (omit empty ones):

- `### Added` — new commands, flags, L10n languages, skills, docs
- `### Changed` — behavior changes, refactored internals, updated dependencies
- `### Fixed` — bug fixes, L10n corrections, AX path updates
- `### Removed` — removed commands, deprecated features
- `### Security` — vulnerability fixes

### Entry Style

- Lead with **bold component name** — then a dash and description
- Be specific: name the command, the L10n key, the AX path
- One logical change per bullet (not one commit per bullet)

Good:
```markdown
- **`calendar create`** — added `--location` flag for setting event location
- **L10n** — added Japanese (ja) labels for all 124 keys
```

Bad:
```markdown
- Updated calendar stuff
- Fixed a bug
```

## Phase 1 — Update CHANGELOG.md

1. Read `CHANGELOG.md`
2. Add entries under `## [Unreleased]` using the categories above
3. If cutting a release, move `[Unreleased]` entries to a new version heading:

```markdown
## [Unreleased]

## [1.1.0] — 2026-04-20

### Added
- **`mail snooze`** — snooze current email with `--until` date/time flag
```

4. Update the compare links at the bottom of the file:

```markdown
[Unreleased]: https://github.com/trsdn/outlook-ax/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/trsdn/outlook-ax/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/trsdn/outlook-ax/releases/tag/v1.0.0
```

## Phase 2 — Create Git Tag and GitHub Release

Only when the user explicitly asks to "cut a release" or "create a release".

### Pre-flight

- [ ] All changes committed and pushed
- [ ] CHANGELOG.md has the new version heading with today's date
- [ ] Version string follows semver (X.Y.Z)

### Tag and Release

```bash
# 1. Create annotated tag
git tag -a v1.1.0 -m "v1.1.0"

# 2. Push tag
git push origin v1.1.0

# 3. Extract release notes from CHANGELOG.md (everything between this version
#    and the previous version heading), write to temp file
# 4. Create GitHub release
gh release create v1.1.0 --title "v1.1.0" --notes-file /tmp/outlook-ax-release-notes.md
```

**Important:** Always use `--notes-file`, never inline `--notes` with heredoc
(shell quoting breaks on markdown content).

### Post-release

- [ ] Verify release appears at https://github.com/trsdn/outlook-ax/releases
- [ ] If this repo is a submodule, update the parent repo's submodule pointer

## Safety Rules

- Never create a release without the user explicitly asking for it
- Never skip the tag step (releases without tags break the compare links)
- Never backdate a changelog entry — use today's date
- If unsure about the version bump level, ask the user
