# Modern Hillbilly Tabs 365

Place neighboring SharePoint page web parts into selectable tabs. This update repairs web part discovery and tab switching for the newer canvas markup patterns covered by the regression tests, while preserving existing tab labels and saved web part references.

## About

**SharePoint compatibility update by Dr. Ghassan Abukar, Ph.D.**

Original Modern Hillbilly Tabs project by **Mark Rackley**: <https://github.com/mrackley/Modern_Hillbilly_Tabs>.

Version: **1.0.1** · SharePoint solution version: **1.0.1.0**.

## Install or update

Download the installation package and complete source archive from [GitHub Releases](https://github.com/drghassanabukar/modern-hillbilly-tabs-365/releases/latest).

1. Use the included `modern-hillbilly-tabs.sppkg` from the download, or build it as described below.
2. Upload it to your SharePoint App Catalog. For an existing installation, replace the previous package and deploy the update. The original solution and web part IDs are preserved.
3. If SharePoint offers an app update on the target site, apply it. Refresh the page after the updated assets become available.
4. Test on a page before using it broadly: select every tab, edit the page, publish, and reopen it. Existing labels should remain.
5. If a tab reports an unavailable web part, open **Manage Tab Labels** and reselect the intended web part from the same section.

Solution ID: `86a6b96a-7f98-4508-827d-b8bd6409f60b`

Web part ID: `6bcd9bfc-425b-47c2-8e5e-c17eb1c864c5`

## Configure

1. Use a standard page section. Place the tabs web part above the content web parts in the same column and section.
2. Edit the tabs web part, then choose **Manage Tab Labels**.
3. Choose a web part and enter a tab label for each row. Rows determine tab order.
4. Publish the page. Only the selected content appears. Content remains visible in edit mode.

The legacy `sectionClass` and `webpartClass` properties remain as optional class-name overrides. Standard markup is detected automatically. Each content web part should belong to one tab set. Content remains in its original canvas location; the update does not move web parts between columns or sections.

Tab buttons support Left/Right Arrow, Home, End, Enter, and Space. Unavailable or already-owned content is disabled with an explanatory message. Labels are rendered as text.

## What changed

- Detects legacy `ControlZone` wrappers, `CanvasControl` automation attributes, and feature instance IDs.
- Matches saved GUIDs and legacy `control_<guid>` references to their newer identity attributes.
- Responds to late rendering and replaced controls for the web part's lifetime.
- Changes visibility in place, avoiding the original approach of moving SharePoint-owned DOM nodes.
- Restores original visibility and accessibility attributes on re-render, edit mode, and disposal.
- Limits discovery to the current section and prevents two tab sets from controlling the same element.
- Replaces the global jQuery tab plugin with instance-scoped tab buttons and keyboard support.
- Updates the build to SPFx 1.21.1, TypeScript 5.3, and Node.js 22.

## Build

Use **Node.js 22.14 or newer in the 22.x line**. The `.nvmrc` file records the selected runtime. No global gulp installation is required.

```sh
npm ci
npm test
npm run package
```

The production package is created at:

```text
sharepoint/solution/modern-hillbilly-tabs.sppkg
```

For local development:

```sh
npm run build
npx gulp trust-dev-cert
npx gulp serve
```

## Validation and compatibility limits

The automated DOM regression suite exercises the actual web part renderer against simulated legacy and newer SharePoint page markup. The original implementation failed the newer-markup and delayed-rendering scenarios before this repair.

**This update has not been verified in your SharePoint tenant.** Confirm the deployment on your affected page. Flexible sections and future canvas markup changes require separate validation; use a standard section for this release.

SharePoint Online supports older SPFx releases, so the old SDK version alone does not explain the runtime failure. Microsoft also documents that the page DOM and CSS are implementation details rather than a supported extensibility API. This solution still depends on those details to control neighboring web parts, so future SharePoint changes may require another compatibility update.

Primary references:

- [Microsoft SPFx compatibility reference](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility)
- [SharePoint Framework overview and page DOM guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft SPFx 1.21 release notes](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/release-1.21)
- [Upstream report: all tab content remains visible](https://github.com/mrackley/Modern_Hillbilly_Tabs/issues/6)

## Disclaimer

Provided as is, without warranty. This repository retains attribution to the original project.
