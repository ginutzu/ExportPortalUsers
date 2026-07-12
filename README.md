# Export Portal Users – ANCPI

Aplicație care extrage **lista completă a membrilor** dintr-un portal ArcGIS
Enterprise și o salvează în format **CSV** și **Excel (.xlsx)**.

Construită pornind de la scriptul original `cod.txt` (bazat pe biblioteca
`arcgis`), rescrisă să folosească **direct API-ul REST al portalului** prin
biblioteca standard Python — deci nu depinde de pachetul greu `arcgis`.

---

## 1. Ce face

1. Se autentifică pe portal și obține un token temporar
   (`/sharing/rest/generateToken`).
2. Parcurge paginat toți membrii (`/sharing/rest/portals/self/users`).
3. Pentru fiecare membru, opțional, citește grupurile din care face parte
   (`/sharing/rest/community/users/{user}`).
4. Scrie rezultatul în `portal_users.csv` și `portal_users.xlsx`.

### Coloane generate

| Coloană       | Descriere                                                    |
|---------------|--------------------------------------------------------------|
| Username      | Numele de utilizator                                         |
| LicenseType   | Tipul de licență (ex. `GISProfessionalAdvUT`, `viewerUT`)    |
| FullName      | Numele complet                                               |
| Email         | Adresa de email                                              |
| Role          | Rolul (ex. `org_user`, `org_admin`)                          |
| Level         | Nivelul contului (`1` / `2`)                                 |
| Disabled      | Cont dezactivat? (`Yes` / `No`)                              |
| Provider      | Furnizorul de identitate (ex. `arcgis`)                      |
| Created       | Data creării contului (`AAAA-LL-ZZ HH:MM:SS`)                |
| LastLogin     | Ultima autentificare (`Never` dacă nu s-a logat niciodată)   |
| Modified      | Data ultimei modificări a contului                           |
| Groups        | Grupurile din care face parte (separate prin `;`)            |

Fișierul `.xlsx` are în plus: antet îngroșat, rând înghețat (freeze pane),
auto-filtru pe coloane și lățimi auto-ajustate.

---

## 2. Cerințe

- **Windows** (testat pe Windows 11).
- Pentru executabil: **nimic** — `ExportPortalUsers.exe` rulează de sine
  stătător, fără Python instalat.
- Pentru rularea din sursă: **Python 3.9+** (testat pe 3.14.5) și pachetul
  `openpyxl` (instalat automat de `start.bat`).
- Acces de rețea la portal: `https://gis.ancpi.ro/portal`.
- Un cont de portal cu drept de a lista membrii organizației.

---

## 3. Cum se rulează

La pornire, aplicația cere interactiv patru informații (apasă **Enter** la
oricare pentru a păstra valoarea implicită afișată în paranteze):

1. **URL portal GIS** – ex. `https://gis.ancpi.ro/portal`
2. **Utilizator (username)** – ex. `ginu.popescu`
3. **Parola** – nu se afișează la tastare, rămâne doar în memorie
4. **Calea de salvare** a fișierelor CSV și XLSX

Există trei variante de lansare, în ordinea simplității:

### Varianta A – Executabil (recomandat)

Dublu-click pe:

```
dist\ExportPortalUsers.exe
```

Aplicația pune cele patru întrebări de mai sus în consolă, apoi generează
fișierele. Nu necesită Python instalat.

### Varianta B – Launcher `start.bat`

Dublu-click pe `start.bat`. Acesta:

1. detectează Python (`py` sau `python`),
2. instalează automat `openpyxl` dacă lipsește,
3. pornește aplicația.

### Varianta C – Direct din sursă

Într-un terminal (PowerShell / CMD):

```powershell
cd C:\#_LUCRU\Aplicatii\ExportPortalUsers
python export_portal_users.py
```

---

## 4. Parola – cum se furnizează

Parola **nu este stocată** în cod sau în fișiere.

- **Implicit:** aplicația o cere interactiv la rulare
  (`Password for ginu.popescu:`). Textul introdus este ascuns și rămâne doar
  în memorie pe durata rulării.
- **Opțional (rulări automate):** dacă este setată variabila de mediu
  `PORTAL_PASSWORD`, aceasta este folosită fără a mai întreba:

  ```powershell
  $env:PORTAL_PASSWORD = "parola"      # doar pentru sesiunea curentă
  ```

> Notă de securitate: promptul interactiv (`getpass`) citește direct din
> consolă, deci trebuie rulat într-un terminal real, nu printr-un pipe sau
> redirecționare de input.

---

## 5. Configurare

URL-ul portalului, utilizatorul și calea de salvare se cer **interactiv** la
rulare. Valorile din tabel sunt doar **valorile implicite** (propuse la prompt)
și pot fi modificate în capul fișierului `export_portal_users.py`:

| Variabilă              | Implicit                                  | Rol                                                        |
|------------------------|-------------------------------------------|------------------------------------------------------------|
| `DEFAULT_PORTAL_URL`   | `https://gis.ancpi.ro/portal`             | URL portal propus la prompt                                |
| `DEFAULT_USERNAME`     | `ginu.popescu`                            | Utilizator propus la prompt                                |
| `DEFAULT_OUTPUT_PATH`  | `C:\#_LUCRU\Aplicatii\ExportPortalUsers`  | Calea de salvare propusă la prompt                         |
| `INCLUDE_GROUPS`       | `True`                                    | Include coloana Groups (un apel REST în plus per membru)   |
| `WRITE_XLSX`           | `True`                                    | Scrie și fișierul Excel                                    |
| `FILTER_LIKE_ORIGINAL` | `False`                                   | `True` = doar licențe viewer, exclude conturile `esri_*`   |
| `INSECURE_SSL`         | `False`                                   | `True` = ignoră verificarea certificatului TLS             |
| `PAGE_SIZE`            | `100`                                     | Membri per apel REST                                       |

> La prompt, apasă **Enter** pentru a accepta valoarea implicită afișată.
> `INCLUDE_GROUPS = False` accelerează semnificativ rularea (elimină un apel
> REST pentru fiecare membru).

---

## 6. Structura folderului

```
ExportPortalUsers\
├─ dist\
│  └─ ExportPortalUsers.exe   executabilul standalone (cu iconiță)
├─ export_portal_users.py     scriptul principal
├─ start.bat                  launcher (Python + auto-install openpyxl)
├─ make_icon.py               generatorul iconiței
├─ app.ico                    iconița aplicației
├─ cod.txt                    scriptul original (sursă de pornire)
├─ README.md                  această documentație
├─ portal_users.csv           output CSV
└─ portal_users.xlsx          output Excel
```

---

## 7. Reconstruirea executabilului

Executabilul a fost creat cu **PyInstaller**. Pentru a-l reconstrui după
modificări în cod:

```powershell
cd C:\#_LUCRU\Aplicatii\ExportPortalUsers

# (o singură dată) instalarea uneltelor
python -m pip install pyinstaller openpyxl pillow

# (opțional) regenerarea iconiței
python make_icon.py

# construirea exe-ului
python -m PyInstaller --onefile --console ^
  --name ExportPortalUsers ^
  --icon "C:\#_LUCRU\Aplicatii\ExportPortalUsers\app.ico" ^
  --collect-all openpyxl ^
  --distpath dist --workpath build --specpath build --noconfirm ^
  export_portal_users.py
```

Executabilul rezultă în `dist\ExportPortalUsers.exe`. Folderele `build\` și
fișierul `.spec` sunt temporare și pot fi șterse.

---

## 8. Probleme frecvente

| Simptom                                             | Cauză / soluție                                                                 |
|-----------------------------------------------------|----------------------------------------------------------------------------------|
| „Windows protected your PC" la pornirea exe-ului    | Fals pozitiv SmartScreen (exe nesemnat). *More info → Run anyway*.               |
| Antivirusul semnalează exe-ul                       | Fals pozitiv frecvent la executabile PyInstaller nesemnate.                     |
| `Token request failed` / eroare de autentificare    | Parolă/utilizator greșit, sau contul nu are drepturi de listare a membrilor.    |
| Eroare de certificat TLS (`CERTIFICATE_VERIFY_FAILED`) | Certificat intern/self-signed → setează `INSECURE_SSL = True`.                |
| `python` nu este recunoscut                         | Python nu este instalat / nu e în PATH. Vezi https://www.python.org/downloads/. |
| Promptul de parolă „nu apare" / se blochează        | Rulează într-un terminal real, nu prin pipe; sau folosește `PORTAL_PASSWORD`.   |
| Coloana `Groups` e goală                            | Membrii chiar nu au grupuri, sau `INCLUDE_GROUPS = False`.                       |

---

## 9. Note tehnice

- Fișierul CSV este scris cu codare **UTF-8 cu BOM** (`utf-8-sig`) pentru
  afișarea corectă a diacriticelor la deschiderea directă în Excel.
- Datele `Created`, `LastLogin`, `Modified` vin de la portal ca
  epoch-milisecunde și sunt convertite în timp local citibil.
- Aplicația nu depinde de biblioteca `arcgis`; folosește doar biblioteca
  standard Python + `openpyxl` (doar pentru exportul `.xlsx`).
