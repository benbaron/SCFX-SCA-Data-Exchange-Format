# README — Installing `SCLX_Ledger_IO_v13_reviewed.bas` into an Excel Workbook

## Purpose

This README explains how to install the SCLX VBA importer/exporter into an Excel workbook so the workbook can:

* export workbook data to **SCLX 1.3 JSON**
* import SCLX JSON back into the workbook
* use the supporting JSON parser library required by the VBA code

This setup is intended for the workbook based on the **SCA Exchequer Report** layout and the VBA module:

* `SCLX_Ledger_IO_v13_reviewed.bas`

It also requires the JSON parsing library module:

* `JsonConverter.bas`

---

## Files you need

Install these VBA modules into the workbook:

1. **Main SCLX VBA module**

   * `SCLX_Ledger_IO_v13_reviewed.bas`

2. **JSON parser dependency**

   * `JsonConverter.bas`

`JsonConverter.bas` comes from the VBA-JSON project.

---

## Before you begin

### 1. Save the workbook as a macro-enabled workbook

If the workbook is currently `.xlsx`, save it as:

* **Excel Macro-Enabled Workbook (`.xlsm`)**

Use:

* **File**
* **Save As**
* choose **Excel Macro-Enabled Workbook (*.xlsm)**

This is required because standard `.xlsx` files cannot store VBA code.

---

## Installation method A — import the `.bas` files

This is the preferred method.

### Step 1. Open the workbook

Open the Excel workbook that will contain the importer/exporter.

### Step 2. Open the VBA editor

Press:

```text
Alt + F11
```

This opens the Visual Basic Editor.

### Step 3. Open the Project Explorer

If Project Explorer is not visible, press:

```text
Ctrl + R
```

Locate your workbook in the tree on the left.

### Step 4. Import `JsonConverter.bas`

In the VBA editor:

* right-click the workbook project
* choose **Import File...**
* select `JsonConverter.bas`

This adds the JSON parser module to the workbook.

### Step 5. Import `SCLX_Ledger_IO_v13_reviewed.bas`

Again:

* right-click the workbook project
* choose **Import File...**
* select `SCLX_Ledger_IO_v13_reviewed.bas`

This adds the SCLX import/export module.

### Step 6. Save the workbook

Press:

```text
Ctrl + S
```

Save the workbook as `.xlsm`.

---

## Installation method B — paste the code manually

Use this only if importing the `.bas` file is not possible.

### Step 1. Open the VBA editor

```text
Alt + F11
```

### Step 2. Insert a new standard module

In the VBA editor:

* right-click the workbook project
* choose **Insert**
* choose **Module**

### Step 3. Paste the code

Open the `.bas` file in a text editor and paste its contents into the new module.

### Important note about `Attribute VB_Name`

If the file begins with a line like:

```vb
Attribute VB_Name = "SCLX_Ledger_IO"
```

delete that line before pasting into the editor.

That line belongs in exported `.bas` files, but usually should not be pasted directly into the code window.

### Step 4. Repeat for `JsonConverter.bas`

Insert another standard module and paste the JSON converter code into it.

### Step 5. Save the workbook

Save as `.xlsm`.

---

## JSON library requirement

The SCLX module depends on `JsonConverter.bas`.

Without it, these calls will fail:

* `JsonConverter.ParseJson(...)`
* `JsonConverter.ConvertToJson(...)`

So the SCLX module will **not work** unless `JsonConverter.bas` is installed in the same workbook VBA project.

---

## Optional reference setting

The reviewed SCLX module uses late binding for dictionaries, so it does **not** require a compile-time reference to Microsoft Scripting Runtime for the main logic.

However, the JSON library may sometimes recommend certain references depending on the exact version you are using.

In many setups, no extra reference is required.

If your `JsonConverter.bas` version instructs you to enable a reference, do so through:

* VBA editor
* **Tools**
* **References**

Then enable the required library.

---

## Compile check after installation

After importing both modules, compile the project.

In the VBA editor:

* choose **Debug**
* choose **Compile VBAProject**

If there is no error, the code is installed correctly.

If Excel highlights a line, fix that issue before running the macros.

---

## Macro security

Excel may block macros by default.

To allow the code to run:

* open Excel
* go to **File**
* **Options**
* **Trust Center**
* **Trust Center Settings**
* **Macro Settings**

Use a setting appropriate for your environment.

A common option while developing is:

* **Disable VBA macros with notification**

Then reopen the workbook and click **Enable Content** when prompted.

---

## How to run the exporter/importer

After installation, the main entry points are:

* `ExportSCLX`
* `ImportSCLX`

### To run them

In Excel:

* press `Alt + F8`
* select:

  * `ExportSCLX` to export workbook data to JSON
  * `ImportSCLX` to import JSON into workbook tabs
* click **Run**

---

## What the installed module expects

The reviewed module is written against the workbook structure it was mapped to, including these sheets:

* `Summary`
* `Ledger`
* `Outstanding`
* `Assets&Inventory`
* `Supplies`

It also expects the workbook to use the mapped row/column layout from the reviewed version.

### Important

The reviewed module is based on the actual workbook layout that was analyzed, including ledger split columns such as:

* split row number in `AG / BG / CG / DG`
* split amount in `AH / BH / CH / DH`
* income category in `AI / BI / CI / DI`
* expense category in `AJ / BJ / CJ / DJ`
* used for in `AK / BK / CK / DK`
* item number in `AL / BL / CL / DL`
* quantity in `AM / BM / CM / DM`

If the workbook layout changes later, the VBA mappings must be updated.

---

## Recommended workbook backup procedure

Before first import or export testing:

1. Make a copy of the workbook.
2. Install the VBA only into the copy.
3. Test `ExportSCLX`.
4. Test `ImportSCLX` with a small sample file.
5. Compare the workbook before and after import.

This protects the production workbook from accidental data changes during testing.

---

## Common installation problems

### Problem: “Sub or Function not defined” on `JsonConverter`

Cause:

* `JsonConverter.bas` is missing

Fix:

* import `JsonConverter.bas` into the same workbook VBA project

---

### Problem: “Expected: end of statement” or similar on `Attribute VB_Name`

Cause:

* you pasted the exported `.bas` file directly into the editor, including the `Attribute` line

Fix:

* delete the `Attribute VB_Name = ...` line before pasting

---

### Problem: macros do not appear in `Alt + F8`

Cause:

* code is not in a standard module
* workbook not saved as `.xlsm`
* project did not compile
* macros are disabled

Fix:

* ensure both modules are in standard modules
* save as `.xlsm`
* compile the project
* enable macros

---

### Problem: compile error in the imported VBA

Cause:

* partial copy/paste
* wrong or incomplete JSON library version
* corrupted line breaks during paste

Fix:

* re-import the `.bas` files directly rather than pasting
* compile again

---

### Problem: runtime error when importing or exporting JSON

Cause:

* malformed JSON
* missing workbook sheet
* workbook layout does not match expected structure
* template/default rows interpreted incorrectly in an older module version

Fix:

* use the reviewed 1.3 module
* confirm sheet names exactly match
* validate the JSON file
* test with a small known-good file first

---

## Windows and Mac notes

### Windows

This setup is best supported in desktop Excel on Windows.

### Mac

The VBA may work on Mac desktop Excel, but compatibility depends on the JSON library version and object support used by that environment.

If using Mac:

* test the JSON library first
* confirm `JsonConverter.bas` works in a simple parse/serialize test

### Excel for the web

VBA macros do **not** run in Excel for the web.

You must use desktop Excel.

---

## Recommended module names

For clarity, keep the modules named approximately like this:

* `JsonConverter`
* `SCLX_Ledger_IO`

The internal module name does not have to match the file name exactly, but keeping it close helps maintenance.

---

## Suggested smoke test

After installation:

1. Open the workbook.
2. Run `ExportSCLX`.
3. Save the JSON file.
4. Open the JSON in a text editor.
5. Confirm you see:

   * `"version": "1.3"`
   * top-level arrays such as `transactions`, `chartOfAccounts`, `outstandingItems`, `assets`, `supplies`
6. Reopen a copy of the workbook.
7. Run `ImportSCLX` with the exported file.
8. Verify rows append correctly in:

   * Ledger
   * Outstanding
   * Assets&Inventory
   * Supplies

---

## Files to distribute with the workbook

For a reusable package, keep these together:

* the macro-enabled workbook (`.xlsm`)
* `SCLX_Ledger_IO_v13_reviewed.bas`
* `JsonConverter.bas`
* the SCLX 1.3 schema
* the SCLX 1.3 validator rules
* the integrator manual

That gives you both the executable workbook integration and the interchange specification.

---

## Summary

To install the SCLX Excel integration:

1. Save workbook as `.xlsm`
2. Open VBA editor with `Alt + F11`
3. Import `JsonConverter.bas`
4. Import `SCLX_Ledger_IO_v13_reviewed.bas`
5. Compile the project
6. Enable macros
7. Run `ExportSCLX` / `ImportSCLX`

The installation is complete once the workbook compiles and the macros appear in `Alt + F8`.
