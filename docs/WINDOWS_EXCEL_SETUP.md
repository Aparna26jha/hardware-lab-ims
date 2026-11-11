# Windows Excel setup

Use these files, already in the project folder:

- `vba_app/dist/Lab_Inventory_System_SEED.xlsx`
- `vba_app/dist/windows_import/` (every `.bas`, both `.txt` files, and the `forms` folder)

Do this in Excel on the Wipro Windows laptop. No extra software is installed.

## 1. Turn the seed file into a macro workbook

1. Open `Lab_Inventory_System_SEED.xlsx` in Excel.
2. File, Save As, and choose **Excel Macro-Enabled Workbook (*.xlsm)**.
3. Name it `IMS_Frontend.xlsm`.
4. If Excel shows a security bar, choose **Enable Content**. You need this once on each machine that will open the file. It is a Trust Center choice, not an installation.

The sheets `User_Directory`, `Assets_Master`, and `Transfer_Audit_Log` are a readable copy of the sample. Editing them does not change the live lab database. The live database is the `.accdb` created in the next steps. The name `BackendPath` points at Config cell B2, and `BackupFolder` points at Config cell B3.

## 2. Import the modules

1. Press Alt+F11.
2. If the project explorer is hidden, press Ctrl+R.
3. Right-click the workbook project, then **Import File**, and import every `.bas` file in `windows_import`:
   - `modGlobals.bas`
   - `modUi.bas`
   - `modSecurity.bas`
   - `modADODB_Backend.bas`
   - `modSeedData.bas`
   - `Module_AccessControl.bas`
   - `Module_BarcodeHandler.bas`
   - `modExport.bas`
   - `modInitDatabase.bas`
   - `modMain.bas`
4. Import every form in `windows_import/forms`:
   - `frmAuth.frm`
   - `frmRegister.frm`
   - `frmDashboard.frm`
   - `frmInward.frm`
   - `frmTransfer.frm`
   - `frmUserAdmin.frm`
   - `frmAccount.frm`
5. In the project explorer, double-click **ThisWorkbook**. Replace that code window with the body of `ThisWorkbook_Code.txt` (the procedure only, not the instruction lines at the top).
6. Choose **Debug, Compile VBAProject**.

Each screen creates its own labels, boxes, lists, and buttons in code when it opens. An imported form therefore looks blank in the VBA designer. That is expected: run `IMS_Start` to see the screens.

If Import is unavailable for a form, or Excel reports an error while importing one, create it by hand:

1. **Insert, UserForm**.
2. In the Properties window (F4), set **(Name)** to the file name without `.frm`, for example `frmDashboard`.
3. Double-click the empty form to open its code window, delete what is there, and paste that form's code from `UserForms_Code.txt` (from `Option Explicit` to the end of that form's section).

No control has to be drawn by hand.

Do not set a reference under Tools, References. The code creates ADODB, ADOX, and FileSystemObject late, so a missing library checkbox cannot break the workbook when it moves to another PC. The Microsoft Forms library that UserForms use is added by Excel itself when the first form is inserted or imported.

On Inward and Transfer, no button is the Default button. A scanner sends Enter after each barcode, and the barcode box uses that Enter to look the tag up.

## 3. Point the workbook at the share

1. Press Alt+F8, run `IMS_EditBackendPath`.
2. Enter a UNC path the lab can write, for example `\\fileserver\HardwareLab\IMS\IMS_Backend.accdb`.
3. The macro writes that path into `BackendPath` and saves the workbook.
4. Run `IMS_EditBackupFolder` and enter a different folder, for example `\\fileserver\HardwareLab\IMS_Backups`. If you skip this, backups go to `IMS_Backups` beside the database folder.

Use a real file server path. Do not put the `.accdb` in OneDrive, SharePoint, or a folder that syncs. Those clients lock the file in a way Access cannot share.

The folder needs to allow the lab group to create files, because Access writes a lock file named `IMS_Backend.laccdb` beside the database while someone is connected. Modify permission on that folder is the usual fit. That is a folder permission you already use for lab shares, not a software install.

## 4. Create the tables

1. Press Alt+F8 and run `IMS_InitializeBackend`.
2. The macro creates the `.accdb` if it is missing, creates `User_Directory`, `Allowed_Directory`, `Assets_Master`, and `Transfer_Audit_Log`, and loads the sample accounts, assets, and three audit rows.
3. Run `IMS_TestConnection`. It should report 6 user accounts.
4. Run it again later only to confirm the path. A second initialize leaves existing users in place and does not duplicate the sample.

`IMS_ResetBackend` is also on Alt+F8. It asks you to type RESET. Only a signed-in Super Admin can run it, so start the app and sign in as SA001 before using it. The dashboard button calls the same routine.

## 5. Sign in

Run `IMS_Start`, or close and reopen the workbook. `Workbook_Open` calls `IMS_Start`.

- SA001 / 1001 sees the whole lab, the waiting request from EN2301, recent activity, Inward, Transfer, Accounts, Back up now, both exports, and reset.
- AD010 / 2001 sees the whole lab, recent activity, Inward, Transfer, and both exports. Accounts, Back up now, and reset stay hidden.
- LD200 / 4001 sees the hardware held by LD200 and by that lead's team, and can export that list.
- EN1042 / 3001 sees only assets assigned to EN1042, and can export that list.
- Everyone has **My account**, where they can change their own access code.
- Create account on the sign-in form sends a request. The alias, mail ID, and full name must match the lab directory. A Super Admin approves it in Accounts and chooses the role.

Scan a barcode into Inward or Transfer and let the scanner send Enter. Typed codes work the same way if you press Enter in the barcode box. On the dashboard, type in the search box and press Enter.

Save the finished `IMS_Frontend.xlsm` and give each person a copy. Every copy must contain the same `BackendPath`. People do not share one open workbook. They share the `.accdb`.

## 6. What to hide before the floor uses it

In `modGlobals.bas`, set `IMS_SHOW_DEMO_FILL` to `False` and compile again. Turn the sample accounts off from Accounts, or run a reset only on a database that is still sample data. Reset saves a backup copy of the database file first. Do not reset a database that already holds real assets.
