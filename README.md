# Hardware Lab Inventory Management System

Inventory register for the semiconductor validation lab. It tracks AMD-specific and Wipro-specific engineering assets: motherboards, add-in cards, debuggers, SoC boards, power supplies, and cables.

About 500 engineers use it. Each person signs in with their own alias ID. Two Super Admins approve access. The inventory custodians inward new hardware and record transfers. A lead sees the team's hardware, and an engineer sees only the assets currently in their name.

The system runs on a standard Wipro Windows laptop. Microsoft Edge, Chrome, and Excel are enough. Nothing else is installed, and no administrator rights are required.

## How it is built

- The practice page is `web_app/index.html`. Open it in Edge or Chrome. Saves stay in that browser on that laptop.
- The lab system is Excel. Each person has their own `IMS_Frontend.xlsm`, so the lab never shares one open workbook.
- One database, `IMS_Backend.accdb`, sits on a normal shared folder. Every Excel copy reads and writes that same file.
- Excel creates the database itself through Office. The Access program is not installed.
- The folder must be a normal file share, for example `\\fileserver\HardwareLab\IMS\`. Keep it out of OneDrive, SharePoint, and any folder that syncs copies by itself. The lab group needs permission to create files there, because Excel writes a small lock file beside the database while someone is saving.

## How it works

- Sign in with an alias ID and a lab access code.
- Super Admin has complete access, including accounts. Admin is the inventory custodian and is the only role that inwards and allocates. Lead Admin sees the team's hardware. User is the engineer and sees only their own assets.
- Create account sends a request. The alias, mail ID and full name must match the lab directory, and the person signs in only after a Super Admin approves the request and chooses the role.
- Inward: scan the barcode and press Enter. Barcode, ownership, type and purchase order are required (marked *). A known family fills the empty fields.
- Transfer: move a board to an engineer, back to the cage, into maintenance, or to Scrapped/Returned. A scrap or return needs a remark. That is the outward record.
- The dashboard filters AMD-Specific and Wipro-Specific and downloads a CSV.
- When the custodian saves, the next person who opens Excel sees that save. The practice page does not do this.

## What each screen shows

Every screen has one line under its title that says what to do next, so an engineer can use it without training.

| Screen | What it adds |
| --- | --- |
| Dashboard | Count cards: total, available, assigned, under maintenance, AMD-Specific, Wipro-Specific. Each person's counts cover only what their role may see. |
| Dashboard | A **Needs your attention** box. A Super Admin sees access requests waiting for approval. Everyone sees the expired boards, and those expiring within 30 days, that are in their view. |
| Dashboard | **Recent activity** for Admin and Super Admin: the last 10 inwards and transfers from the audit log. |
| Dashboard | One search box for barcode, manufacturer serial, model, PO, person, or description. |
| Transfer | The full history of the scanned barcode: every earlier holder, the date, who recorded it, and the remark. |
| My account | Alias, name, role, mail ID and lead, plus **Change my access code**. Everyone has this screen. |
| Back up now | Super Admin only. Copies the database to the backup folder, which is a different folder from the database. |

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

A new request waits for a Super Admin. The Super Admin chooses the real role in **Accounts**. Until then the person cannot sign in.

Use this walkthrough to show the screens: sign-in, access requests, barcode inward, transfer with history, the dashboard cards and attention box, the AMD / Wipro filter, and CSV export. The walkthrough stores its data in that browser on that PC. The shared register for all 500 engineers is the Excel deployment in the next section. Each engineer still gets their own login there.

### Demonstration sequence

1. Sign in as **SA001** / **1001**. The cards show the whole lab. **Needs your attention** lists the request from EN2301 and two boards that are expired or expiring soon. **Recent activity** shows the latest moves.
2. Open **Accounts**. EN2301 is at the top of the list. Choose the role **User**, tick **Approve / Active**, and save. The request leaves the attention box.
3. Sign out. Sign in as **AD010** / **2001**. Open **Inward**, type `MB-7800X-009` in the barcode box, and press Enter. Make and model fill from the existing H13SSL family. Choose **AMD-Specific**, enter a purchase order, and save.
4. Open **Transfer**. Enter `MB-7800X-009`, press Enter, and the history shows the inward. Assign it to **EN1108** with a short remark. The history and the dashboard's recent activity both show the move.
5. On the dashboard, type a manufacturer serial such as `BCM-57508` into search and press Enter. The one matching board is listed.
6. Sign out. Sign in as **LD200** / **4001**. The lead sees the team's hardware only, with no Inward or Transfer.
7. Sign out. Sign in as **EN1108** / **3002**. The new board is on that engineer's list. Open **My account** to show the details and the change-code form.

Barcode prefixes the inward screen understands: `MB` motherboard, `AIC` add-in card, `DBG` debugger, `SOC` SoC board, `PSU` power supply, `CBL` cable.

## Shared register for the whole lab

This is one inventory. When the inventory custodian saves an inward, a Super Admin sees that asset on their own screen. An engineer who signs in sees the rows assigned to their employee ID. Nobody keeps a private copy of the asset list.

Each person opens their own copy of the Excel screens. All of those copies read and write the same database on a Wipro shared folder:

`\\<lab-share>\HardwareLab\IMS\IMS_Backend.accdb`

That folder is the common place. The workbook connects, saves, and disconnects. People do not all open that one database file inside Excel at the same time, because Excel would lock it. They use the screens, and the screens share the data. Two custodians can update different assets at the same time. A transfer writes the asset row and the audit row together; if the audit row fails, the asset change is rolled back.

Setup on the Windows laptop, using Excel:

1. Open `vba_app/dist/Lab_Inventory_System_SEED.xlsx`.
2. Save it as a macro-enabled workbook named `IMS_Frontend.xlsm`. If Excel shows a security bar, choose **Enable Content**.
3. Press Alt+F11 and import every module and form in `vba_app/dist/windows_import`, then paste `ThisWorkbook_Code.txt` into ThisWorkbook. The click path is in [docs/WINDOWS_EXCEL_SETUP.md](docs/WINDOWS_EXCEL_SETUP.md). Each screen builds its own buttons and boxes when it opens, so a form needs only its code.
4. Press Alt+F8 and run `IMS_EditBackendPath`. Point it at the `.accdb` path on the lab share.
5. Run `IMS_EditBackupFolder` and choose a different folder for backups.
6. Run `IMS_InitializeBackend` once. That creates the database, the tables, and the sample accounts.
7. Run `IMS_Start`.

Give every engineer a copy of that `IMS_Frontend.xlsm`. Every copy points at the same database path. A Super Admin adds each person to the lab directory, the person requests access from the sign-in screen, and a Super Admin approves the request with the right role.

Place the database on a normal shared folder the lab can already write to. Keep it off OneDrive, SharePoint, and any folder that syncs. Access also writes a small lock file beside the database, so the lab group needs permission to create files in that folder. How the short connections stay safe with the full lab online is written in [docs/CONCURRENCY_AND_INTERNET.md](docs/CONCURRENCY_AND_INTERNET.md).

Publishing `web_app/index.html` on an intranet link lets someone open the same screens from a browser. That page is the walkthrough. The inventory all 500 people add to is the database on the lab share, reached from Excel on each Wipro laptop.

## Roles

| Role | Who | What they can do |
| --- | --- | --- |
| Super Admin | You and your manager | Everything a custodian can do, plus approving access and choosing the real role |
| Admin | Satish Sir and Noor Sir | Inward hardware, edit metadata, transfer, export the audit. Noor Sir is the custodian who allocates |
| Lead Admin | A lead with a team | View and export only the assets assigned to that lead's engineers. Cannot inward or transfer |
| User | Validation engineer | View and export only the assets assigned to their alias |

A new person can request access only when their alias and mail ID are already on the lab directory. The request stays closed until a Super Admin approves it and chooses the role. Sample accounts:

| Alias ID | Access code | Role |
| --- | --- | --- |
| SA001 | 1001 | Super Admin |
| AD010 | 2001 | Admin |
| LD200 | 4001 | Lead Admin |
| EN1042 | 3001 | User, on LD200's team |
| EN1108 | 3002 | User, on LD200's team |

EN2301 (Sample Requesting Engineer) has already requested access and is waiting for approval. EN2200 / en2200@lab.example / Sample New Engineer is on the directory and has no account yet, so requesting that alias shows the request step from the start. The sample mail IDs are placeholders until the company list arrives.

These access codes belong to the sample data. They are not Windows passwords. The shared database stores a hash of the code, and each person can change their own code from **My account**. The practice page still shows the sample buttons so the screens can be demonstrated. Before the lab goes live, set `SHOW_DEMO_ACCOUNTS` in `web_app/index.html` and `IMS_SHOW_DEMO_FILL` in `vba_app/modGlobals.bas` to false so those buttons disappear. Real engineers choose their own code when they request access. A Super Admin turns an old sample account off from Accounts.

## Rules the screens enforce

- A custodian follows the lab flow. Available can go to an engineer, to maintenance, or to scrapped/returned. A scrapped or returned board is put back only by a Super Admin, who still has complete access.
- A transfer that does not change the person or the status is rejected, so the audit stays meaningful.
- A manufacturer serial number, once entered, cannot be used on a second asset.
- The name on the list comes from the user directory, so a renamed person is current on the next view.
- A date that is past, or inside the next 30 days, shows Expired or Expiring soon beside the status.
- If two people save the same barcode, the second save asks them to scan it again.
- **Back up now** and Reset both copy the database to `IMS_Backend_backup_yyyymmdd_hhnnss.accdb` in the backup folder. That folder is `BackupFolder` on the Config sheet, or `IMS_Backups` next to the database folder when none is set. A backup folder that is the same as the database folder is refused.
- Reset reloads the sample set after the backup. Type RESET to confirm. Do not reset a database that already holds real assets.
- Changing an access code needs the current code, and the new code typed twice.
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

Asset ownership is either `AMD-Specific` or `Wipro-Specific`. Status is `Available in Lab`, `Assigned to Engineer`, `Under Maintenance`, or `Scrapped/Returned`. Every transfer stores the time, the barcode, the previous holder, the new holder, who recorded it, and the remark. The Transfer screen and the dashboard read that same log for the history and recent activity.

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
