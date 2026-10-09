# Changelog

All notable changes to Ribon Studio are documented here.

## [1.0.0] - 2026-10-10

### Added

- Added Vietnamese and English interface language selection, persisted per Windows user.
- Added application and window icon assets.
- Added Ribbon options for built-in tab visibility, `startFromScratch`, and custom tab placement.
- Added text-only Ribbon groups using `labelControl`.

### Changed

- Refined the add-button workflow to distinguish existing tabs, new tabs, groups, and placement options.
- Added a dedicated **Add label** workflow that hides macro, icon, and button-detail settings that do not apply to text-only groups.
- Kept Ribbon customizations embedded in the selected workbook; language preferences remain local to the current Windows user.

### Fixed

- Retried rejected Excel COM calls while Excel is busy.
- Avoided control access while WPF initializes the add-button window.
