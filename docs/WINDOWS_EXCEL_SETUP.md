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

The sheets `User_Directory`, `Assets_Master`, and `Transfer_Audit_Log` are a readable copy of the sample. Editing them does not change the live lab database. The live database is the `.accdb` created in the next steps. The name `BackendPath` points at Config cell B2.

## 2. Import the modules

1. Press Alt+F11.
2. If the project explorer is hidden, press Ctrl+R.
3. Right-click the workbook project, then **Import File**, and import every `.bas` file in `windows_import`:
   - `modGlobals.bas`
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
   - `frmInward.frm`
   - `frmTransfer.frm`
   - `frmDashboard.frm`
   - `frmUserAdmin.frm`
5. In the project explorer, double-click **ThisWorkbook**. Replace that code window with the body of `ThisWorkbook_Code.txt` (the procedure only, not the instruction lines at the top).
6. Choose **Debug, Compile VBAProject**. Fix a missing control only if Excel reports one. Each control's name is the token after `Begin MSForms...` in the matching `.frm` file. The code reads those names.

Do not set a reference under Tools, References. The code creates ADODB, ADOX, and FileSystemObject late, so a missing library checkbox cannot break the workbook when it moves to another PC.

If File, Import is unavailable, create each UserForm with **Insert, UserForm**, rename it in the properties window to the file name without `.frm`, add the controls named in that file, and paste the code from `UserForms_Code.txt` into the form. The imported `.frm` already contains both the controls and the code, so import is the shorter path.

On `frmInward` and `frmTransfer`, do not mark a button as the form's Default button. A scanner sends Enter, and a default button would steal that key. `frmAuth` is the exception: Sign in is the default button.

## 3. Point the workbook at the share

1. Press Alt+F8, run `IMS_EditBackendPath`.
2. Enter a UNC path the lab can write, for example `\\fileserver\HardwareLab\IMS\IMS_Backend.accdb`.
3. The macro writes that path into `BackendPath` and saves the workbook.

Use a real file server path. Do not put the `.accdb` in OneDrive, SharePoint, or a folder that syncs. Those clients lock the file in a way Access cannot share.

The folder needs to allow the lab group to create files, because Access writes a lock file named `IMS_Backend.laccdb` beside the database while someone is connected. Modify permission on that folder is the usual fit. That is a folder permission you already use for lab shares, not a software install.

## 4. Create the tables

1. Press Alt+F8 and run `IMS_InitializeBackend`.
2. The macro creates the `.accdb` if it is missing, creates `User_Directory`, `Assets_Master`, and `Transfer_Audit_Log`, and loads the sample users, assets, and three audit rows.
3. Run `IMS_TestConnection`. It should report 4 user accounts.
4. Run it again later only to confirm the path. A second initialize leaves existing users in place and does not duplicate the sample.

`IMS_ResetBackend` is also on Alt+F8. It asks you to type RESET. Only a signed-in Super Admin can run it, so start the app and sign in as SA001 before using it. The dashboard button calls the same routine.

## 5. Sign in

Run `IMS_Start`, or close and reopen the workbook. `Workbook_Open` calls `IMS_Start`.

- SA001 / 1001 sees inward, transfer, accounts, export, and reset.
- AD010 / 2001 sees inward, transfer, and both exports. Accounts and reset stay hidden.
- EN1042 / 3001 sees only assets assigned to EN1042, and can export that list.
- Create account on the sign-in form adds a User. It cannot create an Admin. SA001 promotes people from Accounts.

Scan a barcode into Inward or Transfer and let the scanner send Enter. Typed codes work the same way if you press Enter in the barcode box.

Save the finished `IMS_Frontend.xlsm` and give each person a copy. Every copy must contain the same `BackendPath`. People do not share one open workbook. They share the `.accdb`.

## 6. What to hide before the floor uses it

In `modGlobals.bas`, set `IMS_SHOW_DEMO_FILL` to `False` and compile again. Change the sample access codes from Accounts, or replace the seed and run a reset only on a database that is still sample data. Do not reset a database that already holds real assets.
