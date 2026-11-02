# Hardware Lab Inventory Management System

Inventory register for the semiconductor validation lab. It tracks AMD-specific and Wipro-specific engineering assets: motherboards, add-in cards, debuggers, SoC boards, power supplies, and cables.

About 500 engineers use it. Each person has their own employee ID. Two or three inventory owners hold Super Admin. Custodians inward new hardware and record transfers. An engineer sees only the assets currently in their name.

The system runs on a standard Wipro Windows laptop. Microsoft Edge, Chrome, and Excel are enough. Nothing else is installed, and no administrator rights are required.

## How it is built

- The practice page is `web_app/index.html`. Open it in Edge or Chrome. Saves stay in that browser on that laptop.
- The lab system is Excel. Each person has their own `IMS_Frontend.xlsm`, so the lab never shares one open workbook.
- One database, `IMS_Backend.accdb`, sits on a normal shared folder. Every Excel copy reads and writes that same file.
- Excel creates the database itself through Office. The Access program is not installed.
- The folder must be a normal file share, for example `\\fileserver\HardwareLab\IMS\`. Keep it out of OneDrive, SharePoint, and any folder that syncs copies by itself. The lab group needs permission to create files there, because Excel writes a small lock file beside the database while someone is saving.

## How it works

- Sign in with an employee ID and a lab access code.
- Super Admin has complete access, including accounts. Admin is the inventory custodian and is the only role that inwards and allocates. User is the engineer and sees only their own assets.
- Create account always makes a User.
- Inward: scan the barcode and press Enter. Type and barcode are required. Ownership and purchase order are required. A known family can fill the empty fields.
- Transfer: move a board to an engineer, back to the cage, into maintenance, or to Scrapped/Returned. A scrap or return needs a remark. That is the outward record.
- The dashboard filters AMD-Specific and Wipro-Specific and downloads a CSV.
- When the custodian saves, the next person who opens Excel sees that save. The practice page does not do this.

## Options

Sir asked for a common place, three access levels, an inward page, a transfer page, and a view page. These are the three ways to do that on the laptops the lab already has.

| Option | What sir sees | Where it runs | Fit for 500 engineers |
| --- | --- | --- | --- |
| 1. Excel screens, one shared database | The same assets the inventory custodian just saved | Each laptop opens Excel. The database sits on a lab shared folder. | This is the one to deploy. |
| 2. Browser page on an intranet server | The same assets, in the browser | Needs a server the company already operates, plus a small service on that server. The laptop cannot host 500 people. | Later, if IT provides that server. |
| 3. Browser page opened from the laptop | The screen design, with sample data on that PC only | Double-click `web_app/index.html` in Edge. | Demonstration of the pages. An inward saved here does not appear on sir's laptop. |

Option 1 matches the note: VB screens, access control, inward, transfer, assignment view, download. Super Admin has full control, Admin inwards and transfers, User views only their own assets. AMD-specific and Wipro-specific stay separate, and the consolidated view and the audit export are on the dashboard.

## Run the walkthrough on this Windows laptop

This is the demonstration you can open in the room.

1. Copy this project folder onto the laptop.
2. Open `web_app/index.html` in Microsoft Edge or Chrome. Double-click the file, or use File → Open.
3. Sign in with a sample account from the table below, or choose **Create account**.

A self-registered person is an engineer. They see only their own assigned hardware. They cannot inward stock or transfer it. A Super Admin promotes a custodian from **Accounts**.

Use this walkthrough to show the screens: sign-in, registration, barcode inward, transfer, the AMD / Wipro filter, and CSV export. The walkthrough stores its data in that browser on that PC. The shared register for all 500 engineers is the Excel deployment in the next section. Each engineer still gets their own login there.

### Demonstration sequence

1. Sign in as **EN1042** / **3001**. The dashboard lists only the boards checked out to that engineer. Inward and Transfer are absent.
2. Sign out. Sign in as **AD010** / **2001**.
3. Open **Inward**. Click the barcode box, type `MB-7800X-009`, and press Enter. Make and model fill from the existing H13SSL family. Choose **AMD-Specific**, enter a purchase order, and save.
4. Open **Transfer**. Enter `MB-7800X-009`, press Enter, then assign it to **EN1108** with a short remark.
5. Sign out. Sign in as **EN1108** / **3002** and confirm the new board is on that engineer's list.
6. Sign in as **SA001** / **1001**. Open **Accounts** to show role changes. **Create account** on the sign-in screen adds another engineer ID.

Barcode prefixes the inward screen understands: `MB` motherboard, `AIC` add-in card, `DBG` debugger, `SOC` SoC board, `PSU` power supply, `CBL` cable.

## Shared register for the whole lab

This is one inventory. When the inventory custodian saves an inward, a Super Admin sees that asset on their own screen. An engineer who signs in sees the rows assigned to their employee ID. Nobody keeps a private copy of the asset list.

Each person opens their own copy of the Excel screens. All of those copies read and write the same database on a Wipro shared folder:

`\\<lab-share>\HardwareLab\IMS\IMS_Backend.accdb`

That folder is the common place. The workbook connects, saves, and disconnects. People do not all open that one database file inside Excel at the same time, because Excel would lock it. They use the screens, and the screens share the data. Two custodians can update different assets at the same time. A transfer writes the asset row and the audit row together; if the audit row fails, the asset change is rolled back.

Setup on the Windows laptop, using Excel:

1. Open `vba_app/dist/Lab_Inventory_System_SEED.xlsx`.
2. Save it as a macro-enabled workbook named `IMS_Frontend.xlsm`. If Excel shows a security bar, choose **Enable Content**.
3. Press Alt+F11 and import the modules and forms in `vba_app/dist/windows_import`, including `modSecurity.bas`, then paste `ThisWorkbook_Code.txt` into ThisWorkbook. The click path is in [docs/WINDOWS_EXCEL_SETUP.md](docs/WINDOWS_EXCEL_SETUP.md). If an older workbook was already imported, import the modules again so sign-in can use the updated database.
4. Press Alt+F8 and run `IMS_EditBackendPath`. Point it at the `.accdb` path on the lab share.
5. Run `IMS_InitializeBackend` once. That creates the database, the tables, and the sample engineers.
6. Run `IMS_Start`.

Give every engineer a copy of that `IMS_Frontend.xlsm`. Every copy points at the same database path. Engineers register from the sign-in screen, or a Super Admin creates them. SA001 then promotes the two or three inventory owners and leaves everyone else as User.

Place the database on a normal shared folder the lab can already write to. Keep it off OneDrive, SharePoint, and any folder that syncs. Access also writes a small lock file beside the database, so the lab group needs permission to create files in that folder. How the short connections stay safe with the full lab online is written in [docs/CONCURRENCY_AND_INTERNET.md](docs/CONCURRENCY_AND_INTERNET.md).

Publishing `web_app/index.html` on an intranet link lets someone open the same screens from a browser. That page is the walkthrough. The inventory all 500 people add to is the database on the lab share, reached from Excel on each Wipro laptop.

## Roles

| Role | Who | What they can do |
| --- | --- | --- |
| Super Admin | Two or three inventory owners | Everything a custodian can do, plus account roles and a sample-data reset |
| Admin | Inventory custodian | Inward hardware, edit metadata, transfer, export the audit |
| User | Validation engineer | View and export only the assets assigned to their employee ID |

Self-registration always creates a User. It cannot create an Admin. Sample accounts:

| Employee ID | Access code | Role |
| --- | --- | --- |
| SA001 | 1001 | Super Admin |
| AD010 | 2001 | Admin |
| EN1042 | 3001 | User |
| EN1108 | 3002 | User |

These access codes belong to the sample data. They are not Windows passwords. The shared database stores a hash of the code. The practice page still shows the sample buttons so the screens can be demonstrated. Before the lab goes live, set `SHOW_DEMO_ACCOUNTS` in `web_app/index.html` and `IMS_SHOW_DEMO_FILL` in `vba_app/modGlobals.bas` to false so those buttons disappear. Real engineers register their own code. A Super Admin turns an old sample account off from Accounts.

## Rules the screens enforce

- A custodian follows the lab flow. Available can go to an engineer, to maintenance, or to scrapped/returned. A scrapped or returned board is put back only by a Super Admin, who still has complete access.
- A transfer that does not change the person or the status is rejected, so the audit stays meaningful.
- A manufacturer serial number, once entered, cannot be used on a second asset.
- The name on the list comes from the user directory, so a renamed person is current on the next view.
- A date that is past, or inside the next 30 days, shows Expired or Expiring soon beside the status.
- If two people save the same barcode, the second save asks them to scan it again.
- Reset copies the database to `IMS_Backend_backup_yyyymmdd_hhnnss.accdb` in the same folder before it reloads the sample set. Type RESET to confirm. Do not reset a folder that already holds real assets.
- A Super Admin cannot turn off their own account. The last active Super Admin cannot be removed.
- A CSV cell that starts with `=`, `+`, `-`, or `@` is written as text.

## What is in this folder

| Path | Purpose |
| --- | --- |
| `web_app/index.html` | Walkthrough. Opens in Edge or Chrome. |
| `vba_app/dist/Lab_Inventory_System_SEED.xlsx` | Excel shell. Save as `.xlsm` on the Windows laptop. |
| `vba_app/dist/windows_import/` | VBA modules and forms to import into that workbook. |
| `vba_app/Assets_Master_Schema.csv` | Table definitions for assets, transfers, and users. |
| `docs/WINDOWS_EXCEL_SETUP.md` | Excel setup on the Windows laptop. |
| `docs/CONCURRENCY_AND_INTERNET.md` | How 500 people share one database, and what the browser page is for. |

Asset ownership is either `AMD-Specific` or `Wipro-Specific`. Status is `Available in Lab`, `Assigned to Engineer`, `Under Maintenance`, or `Scrapped/Returned`. Every transfer stores the time, the barcode, the previous employee, the new employee, who recorded it, and the remark.

## Nothing to install

The Windows laptop uses programs it already has.

| Piece | Where it comes from |
| --- | --- |
| Browser walkthrough | Edge or Chrome, already on the laptop |
| Screens for the shared register | Excel, already on the laptop |
| Database file | Created by Excel through the Office data provider `Microsoft.ACE.OLEDB.12.0` |
| Network calls | `ADODB`, `ADOX`, and `Scripting.FileSystemObject`, which ship with Windows and Office |

No Python runtime, Node.js, Docker, SQL Server, add-in, or custom program is installed on the laptop. The Access application is not required. Do not install a separate Access Database Engine next to Office; a mismatched 32-bit and 64-bit pair conflicts with Excel.

Excel must be allowed to run macros in `IMS_Frontend.xlsm` (Enable Content, or a folder Excel already trusts). That is a Trust Center choice. If site policy blocks every macro, the browser walkthrough still opens, and the shared register waits until this workbook is trusted.
