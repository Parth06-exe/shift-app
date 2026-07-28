# Azure Private Schedule Setup

This app uses Azure Storage Static Website for the public frontend and Azure Functions for the private schedule lookup API.

## What You Already Have

- Storage account: `shiftappstore`
- Public website endpoint: `https://shiftappstore.z49.web.core.windows.net/`
- Public website container: `$web`
- Private schedule container should be: `schedules`
- Monthly schedule file should be: `shift.xlsx`

## Required Azure Resources

1. Azure Storage Static Website
   - Holds only public files like `index.html`.
   - Do not put `shift.xlsx` here.

2. Private Blob Container
   - Name: `schedules`
   - Access level: Private
   - Contains: `shift.xlsx`

3. Azure Function App
   - Runtime: Node.js
   - Recommended plan: Consumption (Windows)
   - Region: same as storage account, preferably `koreacentral`

## Required Excel Format

The schedule must include an employee code column in the same header row as `Ime`.

Accepted code column names:

```text
Code
PIN
Employee Code
Sifra
Šifra
```

Each employee row needs a private code value. Share each code only with that employee.

## Function App Environment Variables

Add these in Function App > Settings > Environment variables / Configuration:

```text
SHIFT_STORAGE_CONNECTION_STRING = <storage account connection string>
SHIFT_CONTAINER = schedules
SHIFT_BLOB = shift.xlsx
ALLOWED_ORIGIN = https://shiftappstore.z49.web.core.windows.net
REQUIRE_EMPLOYEE_CODE = true
SHIFT_CACHE_TTL_MS = 300000
SHIFT_MAX_BLOB_BYTES = 1048576
```

Notes:

- `REQUIRE_EMPLOYEE_CODE` now defaults to `true` in code. Keep it set to `true` in Azure so the portal configuration is explicit.
- `SHIFT_CACHE_TTL_MS` controls how long the Function keeps the parsed workbook in memory. `300000` means 5 minutes.
- `SHIFT_MAX_BLOB_BYTES` rejects unexpectedly large schedule files before parsing. `1048576` means 1 MB.

## Frontend API URL

After deploying the Function App, update this line in `index.html` if your Function URL changes:

```js
var API_URL='https://YOUR-FUNCTION-APP.azurewebsites.net/api/lookup';
```

Current production value:

```js
var API_URL='https://shift-gyere2btfvgsb7cb.austriaeast-01.azurewebsites.net/api/lookup';
```

Then upload the updated `index.html` to the `$web` container if you are deploying the static site manually.

## Important Security Rule

Never upload `shift.xlsx` into:

- `$web`
- GitHub Pages public files
- any public blob container

The browser should only call the Function API. The Function API reads the private Excel file and returns only the matching employee shifts after both exact full name and employee code match.

## Remaining Parser Risk

The backend still uses the `xlsx` npm package to read the workbook. That package has known advisories and the npm package is not fully maintained. This is acceptable only if the private Blob upload path is admin-only and employees cannot upload or replace `shift.xlsx`.

For stronger hardening later, replace `xlsx` with a maintained parser or move the monthly schedule into a database table instead of parsing Excel at request time.
