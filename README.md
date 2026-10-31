# Google Sheet Health Manager

This is a Google Apps Script project that provides a read-only web API for the
health data stored in the "health" Google Sheet (Karasapo nutrition records,
TANITA body composition, steps from Apple Health, and Mounjaro medication
records).

## Features

- Provides a JSON and JSONP web API to read the sheet's content.
- Can return all sheets at once, or a single sheet via the `sheet` parameter.
- No authentication for now (same style as `list-of-stuff-gas`).
- Includes a local testing setup using Vitest.
- Easy to build and deploy with npm scripts.
- Provides an OpenAPI specification in `openapi.yaml`.

## How it Works

### `doGet()`

This function works as a web API. When you access the script's deployment URL,
it reads data from the sheets in the `health` spreadsheet (`overview`,
`daily_nutrition`, `steps`, `meals`, `period_summaries`, `body_composition`,
`body_segments`, `medication_plan`, `medication_log`, `goals_settings`,
`data_dictionary`) and returns them as a single JSON object keyed by sheet
name. The first row of each sheet is used as the keys.

- Add `?sheet=steps` to get only one sheet (returned as an array of rows).
- Add `?callback=myFunction` for a JSONP response.

There is no authentication for now, so anyone who knows the deployment URL
can read the data. Keep the URL private.

## Deployment

- Apps Script project: `health` (bound to the "health" spreadsheet)
- Web app URL: https://script.google.com/macros/s/AKfycbxGka-swbhzKbZsVUVzCbIpFHkkBf1Wr8DDsyBsHJtWhZv0LLWjlMqOSQbHHP2Cc3-O/exec
- Access: ANYONE, no authentication for now. Keep the URL private.

## API Specification

An OpenAPI 3.0 specification is available in [`openapi.yaml`](openapi.yaml).
It documents the `doGet()` endpoint, including the optional `sheet`, `token`,
and `callback` parameters.

## Local Development

This project uses Node.js and npm for local testing and building.

### Setup

1.  Clone this repository to your computer.
2.  Install the necessary developer packages by running:
    ```bash
    npm install
    ```

### Available Scripts

-   **Run tests:**
    ```bash
    npm test
    ```
    This command runs all tests with Vitest and shows a code coverage report.

-   **Build the script for deployment:**
    ```bash
    npm run build
    ```
    This command takes the source code from `src/Code.js` and creates a Google
    Apps Script compatible file at `dist/Code.gs`.

-   **Deploy the script:**
    ```bash
    npm run deploy
    ```
    This command uses `clasp` to upload the contents of the `dist/` folder to
    your Google Apps Script project. You must have `clasp` installed and
    configured for this to work.
