# Export Portal Users – ANCPI

Application that extracts the **complete list of members** of an ArcGIS
Enterprise / Portal and saves it to **CSV** and **Excel (.xlsx)**.

It was built starting from the original script (`cod.txt`, based on the
`arcgis` library) and rewritten to talk **directly to the Portal REST API**
using only the Python standard library — so it does not depend on the heavy
`arcgis` package.

---

## 1. What it does

1. Authenticates against the portal and obtains a temporary token
   (`/sharing/rest/generateToken`).
2. Pages through all members (`/sharing/rest/portals/self/users`).
3. Optionally, for each member, reads the groups it belongs to
   (`/sharing/rest/community/users/{user}`).
4. Writes the result to `portal_users.csv` and `portal_users.xlsx`.

### Generated columns

| Column        | Description                                                  |
|---------------|--------------------------------------------------------------|
| Username      | The user name                                                |
| LicenseType   | License type (e.g. `GISProfessionalAdvUT`, `viewerUT`)       |
| FullName      | Full name                                                    |
| Email         | Email address                                                |
| Role          | Role (e.g. `org_user`, `org_admin`)                          |
| Level         | Account level (`1` / `2`)                                    |
| Disabled      | Account disabled? (`Yes` / `No`)                             |
| Provider      | Identity provider (e.g. `arcgis`)                            |
| Created       | Account creation date (`YYYY-MM-DD HH:MM:SS`)                |
| LastLogin     | Last sign-in (`Never` if the user never logged in)           |
| Modified      | Date the account was last modified                           |
| Groups        | Groups the member belongs to (separated by `;`)             |

The `.xlsx` file additionally has: bold header, frozen header row (freeze
pane), auto-filter on columns and auto-adjusted column widths.

---

## 2. Requirements

- **Windows** (tested on Windows 11).
- For the executable: **nothing** — `ExportPortalUsers.exe` runs standalone,
  without Python installed.
- For running from source: **Python 3.9+** (tested on 3.14.5) and the
  `openpyxl` package (installed automatically by `start.bat`).
- Network access to the portal: `https://gis.ancpi.ro/portal`.
- A portal account with permission to list the organization's members.

---

## 3. How to run

At startup, the application interactively asks for four pieces of information
(press **Enter** at any prompt to keep the default value shown in brackets):

1. **Portal GIS URL** – e.g. `https://gis.ancpi.ro/portal`
2. **User (username)** – e.g. `ginu.popescu`
3. **Password** – not shown while typing, kept in memory only
4. **Save path** for the CSV and XLSX files

There are three ways to launch it, in order of simplicity:

### Option A – Executable (recommended)

Double-click:

```
dist\ExportPortalUsers.exe
```

The application asks the four questions above in the console, then generates
the files. No Python installation required.

### Option B – `start.bat` launcher

Double-click `start.bat`. It detects Python, installs `openpyxl` automatically
if missing, and starts the application.

### Option C – Directly from source

In a terminal (PowerShell / CMD):

```powershell
cd C:\#_LUCRU\Aplicatii\ExportPortalUsers
python export_portal_users.py
```

---

## 4. Password – how it is supplied

The password is **never stored** in code or in files.

- **By default:** the application asks for it interactively at run time
  (`Parola pentru ginu.popescu:`). The entered text is hidden and kept only in
  memory for the duration of the run.
- **Optional (automated runs):** if the `PORTAL_PASSWORD` environment variable
  is set, it is used without prompting:

  ```powershell
  $env:PORTAL_PASSWORD = "password"    # current session only (PowerShell)
  ```

> Security note: the interactive prompt (`getpass`) reads directly from the
> console, so it must be run in a real terminal, not through a pipe or input
> redirection.

---

## 5. Configuration

The portal URL, the user and the save path are requested **interactively** at
run time. The values in the table are only the **defaults** (proposed at the
prompt) and can be changed at the top of `export_portal_users.py`:

| Variable               | Default                                   | Purpose                                                    |
|------------------------|-------------------------------------------|------------------------------------------------------------|
| `DEFAULT_PORTAL_URL`   | `https://gis.ancpi.ro/portal`             | Portal URL proposed at the prompt                          |
| `DEFAULT_USERNAME`     | `ginu.popescu`                            | User proposed at the prompt                                |
| `DEFAULT_OUTPUT_PATH`  | `C:\#_LUCRU\Aplicatii\ExportPortalUsers`  | Save path proposed at the prompt                           |
| `INCLUDE_GROUPS`       | `True`                                    | Include the Groups column (one extra REST call per member) |
| `WRITE_XLSX`           | `True`                                    | Also write the Excel file                                  |
| `FILTER_LIKE_ORIGINAL` | `False`                                   | `True` = only viewer licenses, exclude `esri_*` accounts   |
| `INSECURE_SSL`         | `False`                                   | `True` = skip TLS certificate verification                 |
| `PAGE_SIZE`            | `100`                                     | Members per REST call                                      |

> At the prompt, press **Enter** to accept the default shown.
> `INCLUDE_GROUPS = False` speeds up the run significantly (removes one REST
> call per member).

---

## 6. Folder structure

```
ExportPortalUsers\
├─ dist\
│  └─ ExportPortalUsers.exe   the standalone executable (with icon)
├─ export_portal_users.py     the main script
├─ start.bat                  launcher (Python + auto-install openpyxl)
├─ make_icon.py               the icon generator
├─ app.ico                    the application icon
├─ cod.txt                    the original script (starting point)
├─ README.md                  documentation (Romanian)
├─ README_EN.md               documentation (English)
├─ portal_users.csv           CSV output
└─ portal_users.xlsx          Excel output
```

---

## 7. Rebuilding the executable

The executable was created with **PyInstaller**. To rebuild it after changing
the code:

```powershell
cd C:\#_LUCRU\Aplicatii\ExportPortalUsers

# (once) install the tools
python -m pip install pyinstaller openpyxl pillow

# (optional) regenerate the icon
python make_icon.py

# build the exe
python -m PyInstaller --onefile --console ^
  --name ExportPortalUsers ^
  --icon app.ico --collect-all openpyxl ^
  --distpath dist --workpath build --specpath build --noconfirm ^
  export_portal_users.py
```

The executable is produced in `dist\ExportPortalUsers.exe`. The `build\` folder
and the `.spec` file are temporary and can be deleted.

---

## 8. Troubleshooting

| Symptom                                             | Cause / fix                                                                     |
|-----------------------------------------------------|----------------------------------------------------------------------------------|
| "Windows protected your PC" when starting the exe   | SmartScreen false positive (unsigned exe). *More info → Run anyway*.             |
| Antivirus flags the exe                             | Common false positive for unsigned PyInstaller executables.                     |
| `Token request failed` / authentication error       | Wrong password/user, or the account cannot list the organization's members.     |
| TLS certificate error (`CERTIFICATE_VERIFY_FAILED`) | Internal/self-signed certificate → set `INSECURE_SSL = True`.                   |
| `python` is not recognized                          | Python not installed / not in PATH. See https://www.python.org/downloads/.      |
| The password prompt "does not appear" / hangs       | Run in a real terminal, not via a pipe; or use `PORTAL_PASSWORD`.               |
| The `Groups` column is empty                        | Members really have no groups, or `INCLUDE_GROUPS = False`.                      |

---

## 9. Technical notes

- The CSV file is written with **UTF-8 with BOM** encoding (`utf-8-sig`) for
  correct display of diacritics when opened directly in Excel.
- The `Created`, `LastLogin`, `Modified` values come from the portal as
  epoch-milliseconds and are converted to readable local time.
- The application does not depend on the `arcgis` library; it uses only the
  Python standard library + `openpyxl` (only for the `.xlsx` export).
