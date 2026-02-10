# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SharePoint Framework (SPFx) 1.22.2 List View Command Set extension for SharePoint Online. The extension adds custom commands to SharePoint document library command bars for metadata extraction.

## Build Commands

- **`pnpm run build`** — Clean build, run tests, and produce the .sppkg package (production)
- **`pnpm run start`** — Start the local dev server (port 4321, HTTPS) for debugging against a SharePoint site
- **`pnpm run clean`** — Clean build artifacts

The build system uses `@rushstack/heft` with `@microsoft/spfx-web-build-rig`. Heft orchestrates TypeScript compilation, testing, and SPFx packaging. The rig pattern means most build config is inherited rather than local — see `config/rig.json`.

**Node requirement:** >=22.14.0 <23.0.0

## Dependencies

Dependencies are installed using `pnpm i --shamefully-hoist`.

## Testing

Tests use Jest via Heft. Test files go alongside source or in `__tests__/` directories using `*.test.ts` or `*.spec.ts` naming. `heft test` runs tests as part of the build. Type definitions come from `@types/heft-jest`.

## Architecture

This is a **ListViewCommandSet** extension (not a web part). The single extension lives at:

```
src/extensions/spfxMetadataExtraction/
├── SpfxMetadataExtractionCommandSet.ts       # Main component
├── SpfxMetadataExtractionCommandSet.manifest.json
└── loc/                                       # Localization strings (en-us)
```

**Key patterns:**

- The command set class extends `BaseListViewCommandSet<ISpfxMetadataExtractionCommandSetProperties>`
- Commands are defined in the manifest JSON and referenced by ID (e.g., `COMMAND_1`, `COMMAND_2`) in code
- Command visibility is toggled reactively via `listViewStateChangedEvent` — call `this.raiseOnChange()` after updating visibility
- Properties are passed via `ClientSideComponentProperties` JSON (configured in `config/serve.json` for local dev, or via SharePoint deployment manifests in `sharepoint/assets/`)
- Dialogs use `@microsoft/sp-dialog`
- SharePoint specific libraries, models, services, classes, etc should be extracted away as soon as possible in favor of domain specific representations.  Ideally, these domain specific representations are created within the Extension.ts (e.g. MetadataExtractionCommandSet.ts which is the bridge between SharePoint Online and custom code) and passed as properties/parameters to custom functionality.

**Service boundaries:**

- All SharePoint REST calls (reads and writes) must go through a service class (e.g., `MetadataExtractionService`). Dialogs and React components must never import or use `ISharePointRestClient` directly.
- Do not add abstract methods to `FieldBase` unless they are called from production code. Prompt generation lives in `MetadataExtractionField.getExtractionHint()`, not on individual field subclasses.
- Never mutate objects inside React state updaters. When updating a `MetadataExtractionField` in state, call `ef.clone()` before modifying properties, and return the clone.

**Component ID:** `c2fbb0ac-b2e6-48ff-8b6a-af3065224b39`

## UI

SPFx components, generally, should feel like part of the SharePoint Online environment (page, list/library, etc).  The Fluent 2 design system should be used when possible and since we are using React, the associated React components: `@fluentui/react` (reference: https://developer.microsoft.com/en-us/fluentui#/controls/web).

## Configuration Files

- `config/serve.json` — Local dev server config; update `{tenantDomain}` with your actual SharePoint tenant URL before running `pnpm start`
- `config/package-solution.json` — Solution packaging, feature definitions, and solution ID
- `config/config.json` — Bundle entry points and localized resource mappings
- `.eslintrc.js` — ESLint config extending `@microsoft/eslint-config-spfx` with extensive rule overrides

## Deployment

Production build produces `sharepoint/solution/spfx-metadata-extraction.sppkg`. This package is uploaded to the SharePoint App Catalog. The extension deploys via CustomAction defined in `sharepoint/assets/elements.xml`, targeting document libraries (Registration ID 101).
