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

## Function App Environment Variables

Add these in Function App > Settings > Environment variables / Configuration:

```text
SHIFT_STORAGE_CONNECTION_STRING = <storage account connection string>
SHIFT_CONTAINER = schedules
SHIFT_BLOB = shift.xlsx
ALLOWED_ORIGIN = https://shiftappstore.z49.web.core.windows.net
```

Optional later, if the Excel file has a `Code`, `PIN`, `Sifra`, or `Šifra` column:

```text
REQUIRE_EMPLOYEE_CODE = true
```

If this is true, the API requires both name and code.

## Frontend API URL

After deploying the Function App, update this line in `index.html`:

```js
var API_URL='https://YOUR-FUNCTION-APP.azurewebsites.net/api/lookup';
```

Replace it with your real Function App URL, for example:

```js
var API_URL='https://shift-app-api.azurewebsites.net/api/lookup';
```

Then upload the updated `index.html` to the `$web` container.

## Important Security Rule

Never upload `shift.xlsx` into:

- `$web`
- GitHub Pages public files
- any public blob container

The browser should only call the Function API. The Function API reads the private Excel file and returns only the matching employee shifts.
