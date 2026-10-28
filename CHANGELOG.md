# Changelog

## 1.2.0 — October 28, 2026

### Fixed

- Searchable and Automatic category menus now open and select content even when host handlers cancel native click, Enter, or Space actions. The controls explicitly handle activation and isolate chooser interactions from canvas selection handlers.
- Browser regressions cover normal and guarded host behavior, thresholds, search, selection, focus, empty results, and cleanup.

### Added

- Update notices for site owners and page editors, with the loaded version, newer-release download, release notes, manual checks, and conditional reload guidance.
- A setting to disable release checks, one-hour shared session caching, request timeouts, and rate-limit handling.
- Separate permission, network, version, URL-validation, and update lifecycle regression tests.

### Removed

- The original-project GitHub link in the addon and current README. Plain-text original-author attribution remains.

### Upgrade

Install the new `modern-hillbilly-tabs.sppkg` from the release in the SharePoint App Catalog and enable the update. Refresh the affected page and check **Updates** in edit mode for loaded version **1.2.0**. Reloading alone does not install the package. Once this version is deployed, future stable GitHub releases with their installation package attached can trigger notices.

## 1.1.0 — October 28, 2026

### Added

- Six tab designs: Classic underline, Rounded pills, Outlined tabs, Segmented control, Solid blocks, and Minimal.
- Left, Centered, and Full width alignment settings for each web part.
- A searchable category menu that keeps the current category visible and closes after a new selection.
- An Automatic navigation layout that uses the category menu when there are more than eight categories.
- Interactive sample navigation in page edit mode, so editors can preview their settings before publishing.
- SharePoint page and section theme support, including dark backgrounds, readable filled-tab text, and high-contrast focus indicators.
- Illustrated configuration instructions and 43 automated regression tests.

### Changed

- The production package command cleans previous build assets before bundling.
- New web parts default to Automatic navigation. Existing saved web parts retain their tab bar and Classic underline design until an editor changes their settings.

### Upgrade

Upload `modern-hillbilly-tabs.sppkg` to the SharePoint App Catalog, replace the existing package, and enable the update. Original app and component IDs are unchanged; saved labels and web part references remain compatible.

For a page with many categories, edit the tabs web part and choose **Appearance → Navigation layout → Searchable category menu**, then republish the page. **Tab design** and **Tab alignment** are in the same group.

The update is credited to **Dr. Ghassan Abukar, Ph.D.** Original project attribution to **Mark Rackley** is retained.

## 1.0.1 — October 28, 2026

- Repaired tab discovery for legacy and newer SharePoint canvas markup patterns covered by the tests.
- Added handling for delayed and replaced web parts.
- Preserved SharePoint-owned DOM elements while changing tab visibility.
- Added keyboard navigation, content restoration, and 15 regression tests.
- Updated the build to SPFx 1.21.1, TypeScript 5.3, and Node.js 22.
