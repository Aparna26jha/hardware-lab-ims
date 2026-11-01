# Shared use on the lab floor, and the browser copy

## Why one shared workbook fails

If 500 people open the same `.xlsm`, Excel takes a write lock on that file. Everyone else is read-only or gets a sharing violation, and a crash while the file is open can corrupt the workbook. The floor copy avoids that split:

- Each person runs `IMS_Frontend.xlsm` from their own disk or profile. The forms never bind to a range in a shared workbook.
- The only shared file is `IMS_Backend.accdb` on an SMB share.
- A save or a search opens an ADODB connection, runs one statement, and closes the connection before the form waits for the next scan.

## What the connection string is doing

`modADODB_Backend` opens ACE with three choices that matter under load:

- `OLE DB Services=-2` turns connection pooling off. Pooling is useful for SQL Server, where a pool keeps a network session warm. On an Access file it keeps the `.laccdb` lock after `Close`, so the next person appears to collide with a session that the VBA has already finished. Closing the connection is what releases the file, and pooling would undo that. The design is short connections, not a reused pool.
- `Mode=16` is Share Deny None, so one open connection does not demand an exclusive lock.
- `Jet OLEDB:Database Locking Mode=1` is row-level locking. Two custodians can update different assets at the same time. They only contend if they update the same row.

Transfer and inward do the asset update and the audit insert inside one `BeginTrans` / `CommitTrans`. If the audit insert fails, the asset change rolls back. A lock error (Jet 3218, 3260, and the related provider codes) is retried up to four times with a short pause. The window is the time of one update, not the time a form stays open.

Practical limits that keep the file healthy:

- Do not open the `.accdb` in the Access application during the day. An exclusive open blocks the floor.
- Do not store it in OneDrive, SharePoint, or any sync folder.
- Compact the file from one admin PC after hours if it grows for months. That is a maintenance copy, not an extra product.
- Back up the `.accdb` when nobody is in it. Copying the file is enough.
- The role checks in the forms stop an engineer from inwarding hardware. They are not a security boundary against someone who can open the tables. The share permissions are the boundary. The access code is a lab PIN. People must not reuse a Windows or email password, because the code is stored in the user table for this workbook to check.

## Browser edition on a locked-down laptop

1. Copy `web_app/index.html` to the laptop. A network share you can already read is enough.
2. Double-click the file, or open Edge or Chrome and choose that file.
3. No runtime, extension, or admin approval is involved. The page does not load a CDN.
4. Sign in with a sample account, or choose Create account. A self-registered person is a User and sees only the dashboard.
5. SA001 can open Accounts, change a role, and restore the sample set by typing RESET.
6. Inventory and audit export download a CSV through the browser's normal download. The file is not linked back to the page.

Data lives in `localStorage` for that browser profile on that Windows user. Clearing site data, or opening the file in another browser, starts from the sample set again. Two engineers on two PCs do not see each other's scans.

## Putting the same page on an intranet

The laptop still installs nothing. Someone who already administers an internal site copies the single HTML file into a directory that site already serves. Typical cases:

1. An existing IIS or Apache folder the web team already owns. They add `index.html` the same way they add any other static page. Users open `https://intranet.example/lab-ims/`.
2. If the company already allows a static host you can publish to without new software, upload the same one file. Do not stand up a new server for this.

What that URL does, and does not do:

- Sign-in and registration work. They write the user list into the browser that submitted the form.
- The page is still a private notebook for each browser. It does not become a shared database because it was opened from `https`.
- Do not put the file on the public internet. There is no server-side session, and the sample access codes are in the page source until you remove the seed and the demo buttons.

Set `SHOW_DEMO_ACCOUNTS` to `false` in `web_app/index.html` before that file is placed on an intranet, and replace the sample access codes in the seed only if this browser copy is meant to hold real names. Real shared custody of AMD and Wipro assets stays on the Access file.

## Which edition to give the floor

| Need | Use |
| --- | --- |
| Walkthrough of the screens on a Wipro Windows laptop | `web_app/index.html` |
| An intranet link that must open with no install | The same HTML file, knowing each browser keeps its own data |
| Custodians and engineers updating one inventory together | `IMS_Frontend.xlsm` plus `IMS_Backend.accdb` on the SMB share |
