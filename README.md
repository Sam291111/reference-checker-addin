# Reference Checker Excel Add-in

This is an Office.js task-pane add-in for Excel on Mac. It checks the current workbook against an imported lookup CSV and gives non-blocking warnings when a paper is already in the lookup.

## What It Checks

- Data table sheets: columns named `reference_1`, `reference_2`, `reference_3`, and so on.
- Work doc sheets: rows with `doi`, `authors` or `author`, `year`, and optionally `title`.
- Lookup CSV: expects columns named `title`, `doi`, and `author`.

The add-in does not assume that `_1` maps to `a`, `_2` maps to `b`, etc. It uses the lookup table as the source of truth.

If a typed reference is not found in the lookup table, nothing is changed.

## Hosted Setup

The normal setup is to publish the static add-in files to GitHub Pages. Excel then loads the add-in from GitHub's HTTPS URL, so you do not need to start the local server each time.

1. Create a GitHub repository for this folder.
2. Push this folder to the repository's `main` branch.
3. Wait for the `Deploy GitHub Pages` workflow to finish. It publishes the built add-in to the `gh-pages` branch.
4. In GitHub, go to `Settings > Pages`, set the source to `Deploy from a branch`, choose `gh-pages` and `/ (root)`, then save.
5. Install the hosted manifest locally:

```sh
cd "/Users/samwinter/Documents/Shared/University/Summer Internships/Year 2/DAWN - UoB/Work/Excel Extention Development/reference-checker-addin"
ADDIN_BASE_URL=https://USERNAME.github.io/REPOSITORY npm run install:pages-manifest
```

Replace `USERNAME` and `REPOSITORY` with the GitHub Pages URL shown by the workflow. For example:

```text
https://samwinter.github.io/reference-checker-addin
```

Restart Excel after installing the hosted manifest. The lookup CSV is still uploaded by the user inside Excel; it is not committed to GitHub or hosted on the server. When you switch from `localhost` to GitHub Pages, import the lookup CSV again because Excel treats them as different add-in storage locations.

## Local Fallback

Use this when you want to test changes locally or when the hosted add-in is unavailable.

1. Install the local manifest:

```sh
cd "/Users/samwinter/Documents/Shared/University/Summer Internships/Year 2/DAWN - UoB/Work/Excel Extention Development/reference-checker-addin"
npm run install:local-manifest
```

2. Start the local server:

```sh
npm start
```

The local task pane is served from:

```text
https://localhost:3000/taskpane.html
```

Stop the local server with `Ctrl-C`.

## Sideload On Excel For Mac

Microsoft's Mac sideloading guide uses Excel's `wef` folder:
<https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac>

1. Install either the hosted manifest or the local fallback manifest using the commands above.
2. Open Excel.
3. Open the workbook you want to check.
4. Use `Insert > Add-ins > My Add-ins`, then open `Reference Checker`.
5. After it loads, Excel should show a `References` group on the Home ribbon with an `Open Checker` button.

If Excel does not show the add-in immediately, restart Excel after copying the manifest.

If Terminal hangs while creating or opening the Excel container folder, use Finder instead:

1. Press `Cmd + Shift + G`.
2. Paste this folder path:

```text
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

3. Create the `wef` folder if Finder says it does not exist.
4. Copy either `dist/manifest.xml` for hosted mode or `manifest.xml` for local fallback mode into that folder as `reference-checker-manifest.xml`.

## Usage

1. Open the Reference Checker task pane.
   - Use `Home > References > Open Checker` if the ribbon button is visible.
   - Use `Insert > Add-ins > My Add-ins > Reference Checker` if Excel only shows it through My Add-ins.
2. Import `Look-Up.csv`.
3. Open the movable review window with either:
   - `Home > References > Review Window`, or
   - `Open Review Window` in the task pane.
4. Move the review window to another screen if useful.
5. Click `Enable / Test Sound` in the review window once. This unlocks audio in that separate Office window.
6. Edit cells as usual.

When a match is found, the add-in adds a suggestion to the review queue and shows a short toast alert. If sound is enabled, it also plays a two-note warning sound. It does not change cells while detecting matches.

Settings:

- `Show toast alerts`: toggles short non-blocking alerts in the task pane.
- `Play warning sound`: toggles the audible warning when new suggestions are added.
- `Enable / Test Sound`: unlocks and tests audio for that pane/window. The review window has its own sound unlock because Office opens it as a separate browser surface.

Review queue actions:

- `Check Active Sheet`: scans the current sheet and adds suggestions.
- `Select All`: selects every visible suggestion in the queue.
- `Apply Selected`: updates selected, unambiguous citation suggestions.
- `Remove Rows`: deletes selected work-doc rows that are already present in the lookup.
- `Remove Selected`: removes selected suggestions from the queue.
- `Clear Queue`: removes all suggestions from the queue.
- `Hide Pane`: hides the task pane on supported Excel versions while the shared runtime keeps running.

The add-in checks columns named `author` and `reference_1`, `reference_2`, `reference_3`, and so on. A typed `author year_number` value is replaced only after you select the queue item and apply it.

Ambiguous matches can be reviewed and removed from the queue, but are never applied automatically. For example, if the lookup contains both `rockstrom et al 2009a` and `rockstrom et al 2009b`, a typed `rockstrom et al 2009_1` will be queued for manual review.

The `Work-Doc` format is detected too: the add-in checks DOI first, then title/year and author/year. Work-doc matches are queued as `Already in lookup - skip this paper` items, and the matching row is softly highlighted in the sheet. They are not applied as cell replacements. Select those items and use `Remove Rows` if you want Excel to delete the duplicate listing row entirely.

## Test

```sh
cd "/Users/samwinter/Documents/Shared/University/Summer Internships/Year 2/DAWN - UoB/Work/Excel Extention Development/reference-checker-addin"
npm test
```

Build the GitHub Pages output locally:

```sh
ADDIN_BASE_URL=https://USERNAME.github.io/REPOSITORY npm run build:pages
```
