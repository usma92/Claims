# Claims — Azure AD App Registration & First-Run Auth

## Why this is needed

`download_planner.py` calls the Microsoft Graph API using **delegated permissions**
(it acts as you, with your Planner access). Azure AD needs a registered application
to issue tokens. This is a one-time setup — once done, the script runs silently via
cached refresh token.

---

## Step 1 — Register an Azure AD app (IT or self-service)

If you have access to the Azure portal (portal.azure.com) under the Fred Jones tenant:

1. Go to **Azure Active Directory → App registrations → New registration**
2. Name: `Claims Dashboard Downloader`
3. Supported account types: **Single tenant** (this org only)
4. Redirect URI: leave blank (public client — we use device flow)
5. Click **Register**
6. Copy the **Application (client) ID** — paste it into `download_planner.py` as `CLIENT_ID`
7. Go to **API permissions → Add a permission → Microsoft Graph → Delegated**
   - Add: `Tasks.Read`
   - Add: `User.Read`
8. Click **Grant admin consent** (or ask IT to do this step)

If you don't have portal access, send IT this ticket template:

> **Request:** Register an Azure AD app for an internal automation script.
> - Name: `Claims Dashboard Downloader`
> - Type: Public client / native app (no secret needed)
> - Delegated permissions: `Tasks.Read`, `User.Read`
> - Tenant: `4455dc99-bca8-4daf-bd45-2a50a7ceb65f`
> Please share the Application (client) ID when done.

---

## Step 2 — Paste the CLIENT_ID into the script

Open `scripts/download_planner.py`, find line:

```python
CLIENT_ID = "YOUR_APP_CLIENT_ID"
```

Replace `YOUR_APP_CLIENT_ID` with the GUID IT provided.

---

## Step 3 — Install dependencies

```powershell
cd "C:\Users\dbalan\Documents\Claude\Projects\Team Dashboards\2. Claims"
.venv\Scripts\activate.bat
pip install msal requests
```

---

## Step 4 — First run (interactive auth)

Run the script once from the terminal while logged in to Windows:

```powershell
cd "C:\Users\dbalan\Documents\Claude\Projects\Team Dashboards\2. Claims"
.venv\Scripts\activate.bat
python scripts\download_planner.py
```

It will print something like:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXX-XXXX to authenticate.
```

Open that URL in any browser, sign in with your Fred Jones account, enter the code.
The script completes, writes `data/raw/Ford_Claims.xlsx` and `data/raw/Chrysler_Claims.xlsx`,
and saves a refresh token to `data/.token_cache.json`.

**All subsequent runs (including Task Scheduler) are fully silent** — the cached
refresh token handles authentication without any browser interaction.

---

## Step 5 — Register in Task Scheduler (monthly)

```powershell
$proj = "C:\Users\dbalan\Documents\Claude\Projects\Team Dashboards\2. Claims"
$bat  = "$proj\scripts\run_pipeline.bat"

$action  = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$bat`" >> `"$proj\logs\download.log`" 2>&1"
$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "06:30AM"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask `
    -TaskName "Claims — Monthly Download" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Pulls Ford and Chrysler Claims from MS Planner via Graph API"
```

Creates a task that fires on the 1st of each month at 06:30. Adjust `-DaysOfMonth`
and `-At` to match your preferred timing.

---

## Token cache location

`data/.token_cache.json` — contains a refresh token tied to your identity.
- **Do not commit this file to git** (it's already in `.gitignore` via the `data/` exclusion)
- If the token ever expires or you switch accounts, delete the cache file and re-run
  Step 4 (device flow auth) once.
