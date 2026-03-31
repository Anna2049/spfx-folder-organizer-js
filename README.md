# List Folder Organizer (SPFx 1.13 — Pure JavaScript)

An SPFx web part that organizes SharePoint list items into folder hierarchies based on configurable grouping rules. Built with **pure JavaScript** targeting **SPFx 1.13** for maximum compatibility with corporate npm registries (Artifactory, Nexus, etc.).

## Features

- Displays all custom lists on the current site
- Supports three grouping strategies:
  - **By Date** — Year / Month / Day (1–3 levels)
  - **By Initial Letter** — First letter / First two / First three (1–3 levels)
  - **By Choice Value** — Groups items by choice field value (1 level)
- Auto-enables folder creation on lists that don't have it
- Creates missing folders automatically
- Moves items into target folders
- Shows real-time progress and activity log

## Technical Details

| Component | Version |
|---|---|
| SPFx runtime | 1.13.0 |
| SPFx build tools | 1.12.1 |
| React | 16.13.1 |
| Fluent UI React | 8.48.0 |
| TypeScript (compiler only) | 4.5.5 |
| Source code style | **Pure JavaScript** (`.ts`/`.tsx` wrappers required by SPFx) |

> **Note:** All source code is written in plain JavaScript with `var`, `function`, `prototype` patterns, and `React.createElement` instead of JSX. The `.ts`/`.tsx` file extensions are required by the SPFx build pipeline, but no TypeScript-specific syntax (interfaces, generics, type annotations, enums, etc.) is used.

## Project Structure

```
spfx-folder-organizer-js/
│
├── post-install-patches.js       # Patches node_modules after install
├── gulpfile.js                   # Build pipeline, reads .env
├── package.json
├── tsconfig.json
│
├── config/                       # SPFx build configuration
│   ├── config.json
│   ├── copy-assets.json
│   ├── package-solution.json
│   ├── serve.json
│   └── write-manifests.json
│
└── src/
    └── webparts/
        └── listFolderOrganizer/
            ├── ListFolderOrganizerWebPart.ts            # Web part entry point
            ├── ListFolderOrganizerWebPart.manifest.json
            ├── components/
            │   ├── ListFolderOrganizer.tsx               # Main UI component
            │   ├── ListFolderOrganizer.module.scss
            │   └── IListFolderOrganizerProps.ts
            ├── models/
            │   ├── GroupingType.ts
            │   └── IListInfo.ts
            ├── services/
            │   └── SharePointService.ts
            └── loc/
                ├── en-us.js
                └── mystrings.d.ts
```

## Setup

### Prerequisites

- **Node.js** 14+ (tested on Node 14, 16, 18, 22)
- **npm** 6+
- Access to npm registry (public or corporate Artifactory)

### Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/Anna2049/spfx-folder-organizer-js.git
cd spfx-folder-organizer-js

# 2. Create .env file with your SharePoint site URL
cp .env.example .env
# Edit .env and set: SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/yoursite

# 3. Install dependencies (--ignore-scripts required on Node 17+)
npm install --ignore-scripts

# 4. Apply post-install patches (node-sass, SPBuildRig, rush-stack-compiler shim)
node post-install-patches.js

# 5. Build
npm run build        # Debug bundle
npm run ship         # Production bundle + .sppkg package
```

### Corporate Artifactory Setup

If your corporate environment uses an npm proxy (e.g., JPMorgan Chase Artifactory), configure it before installing:

```bash
npm config set registry https://your-artifactory-url/api/npm/npm
```

The following packages must be available in your Artifactory:

| Package | Version |
|---|---|
| `@microsoft/sp-core-library` | 1.13.0 |
| `@microsoft/sp-http` | 1.13.0 |
| `@microsoft/sp-webpart-base` | 1.13.0 |
| `@microsoft/sp-component-base` | 1.13.0 |
| `@microsoft/sp-property-pane` | 1.13.0 |
| `@microsoft/sp-lodash-subset` | 1.13.0 |
| `@microsoft/sp-build-web` | 1.12.1 |
| `@microsoft/sp-module-interfaces` | 1.13.0 |
| `@fluentui/react` | 8.48.0 |
| `react` | 16.13.1 |
| `react-dom` | 16.13.1 |
| `typescript` | 4.5.5 |
| `sass` | 1.32.x+ |
| `gulp` | 4.0.2 |
| `dotenv` | 16.4.7 |

### What the Post-Install Patches Do

The `post-install-patches.js` script applies three patches to node_modules:

1. **node-sass → dart-sass redirect**: Replaces `node-sass/lib/index.js` with a shim that delegates to the `sass` (dart-sass) package. This avoids the native C++ build that node-sass requires (which needs Python and a C++ compiler).

2. **SPBuildRig.js Node version check**: The SPFx 1.12 build tools throw an error on Node 17+. The patch converts the `throw` into a `console.warn`.

3. **@microsoft/rush-stack-compiler-3.2 shim**: The SPFx build pipeline requires this package for TypeScript compilation, but it's not available in many corporate Artifactories. The patch creates a minimal shim package that wraps the local TypeScript installation.

## npm Scripts

| Script | Description |
|---|---|
| `npm run build` | Debug bundle (`gulp bundle`) |
| `npm run bundle` | Production bundle (`gulp bundle --ship`) |
| `npm run package` | Package solution (`gulp package-solution --ship`) |
| `npm run ship` | Production bundle + package in one step |
| `npm run serve` | Start dev server with SharePoint workbench |
| `npm run clean` | Clean build artifacts |

> All build scripts include `--openssl-legacy-provider` to support webpack 4 on Node 17+.

## Build & Deploy

```bash
# Production build
npm run ship

# The .sppkg file is at:
# sharepoint/solution/spfx-folder-organizer-js.sppkg
```

Upload the `.sppkg` to your SharePoint **App Catalog**, then add the **List Folder Organizer** web part to any page.

### Using the Workbench (Local Testing)

```bash
npm run serve
```

1. You must be **signed in** to your SharePoint Online tenant in the browser
2. Accept the self-signed certificate warning on first run
3. On the workbench page, click **+** to add a web part → find **"List Folder Organizer"**
4. The workbench uses **real SharePoint data** — lists, fields, and items are all live

## How It Works

1. The web part loads all custom lists (BaseTemplate 100) from the current site
2. User selects lists, picks a grouping strategy, depth level, and source field
3. Source field dropdown filters by type (DateTime for date grouping, Text/Choice for others)
4. On **Organize Selected**:
   - Enables folder creation on the list if not already enabled
   - Fetches all list items
   - Calculates target folder path per item (e.g., `2024/03 - March/15`)
   - Creates any missing folders in the hierarchy
   - Moves items that aren't already in the correct folder
5. Progress and results are shown inline and in the activity log

## Troubleshooting

### `error:0308010C:digital envelope routines::unsupported`
Your Node version is 17+. The npm scripts already handle this with `--openssl-legacy-provider`. If running gulp directly, set the env var first:
```bash
set NODE_OPTIONS=--openssl-legacy-provider
npx gulp bundle
```

### `Cannot find module '@microsoft/rush-stack-compiler-3.2'`
Run `node post-install-patches.js` — it creates the shim package.

### `Node Sass does not yet support your current environment`
Run `node post-install-patches.js` — it redirects node-sass to dart-sass.

### `Your dev environment is running NodeJS version ... which does not meet the requirements`
Run `node post-install-patches.js` — it disables the version check.
