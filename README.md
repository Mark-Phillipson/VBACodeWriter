# VBACodeWriter

VBACodeWriter is a Microsoft Access COM add-in that helps you write VBA faster in the VBE by searching Access objects and code symbols, then inserting or opening the selected item.

## What it offers

The add-in adds a **VBA Code Writer** menu to the Access VBA editor (VBE) with these commands:

- **List Forms** - list form names from the current database.
- **List Reports** - list report names.
- **List Controls** - list controls from the previously selected form/report.
- **List Tables** - list table names.
- **List Queries** - list query names.
- **List Fields** - list fields from the currently selected table/query.
- **List Variables** - list variables visible in the active procedure/module declarations.
- **Dimension Variable** - inserts a `Dim` statement for a selected variable name using naming-prefix inference (for example `str* -> As String`, `lng* -> As Long`, `rs* -> As DAO.Recordset`).
- **List Procedures** - list procedures in the active module.
- **List Modules** - list modules and open one, then list its procedures.
- **List All Procedures** - list procedures across all modules in `Module.Procedure` format.
- **Parse SQL** - formats SQL text into VBA string-building code and can insert it into the active module.
- **Comment Block** / **UnComment Block** - adds built-in VBE command bar actions for comment/uncomment.

## Search form behavior

Most commands open the same search dialog where you can:

- Filter with a search box.
- Select from list results.
- Insert the selected value into the current cursor/selection in the active code pane.
- Copy the selected value to clipboard.
- Optionally open the selected object (used for modules/all procedures).
- For tables/queries, optionally continue to a field list.

Quick select buttons (**Select Top**, **Select 2nd** ... **Select 9th**) are included to pick top matches quickly.

## How to use

1. Open your Access database and press `Alt+F11` to open the VBA editor.
2. In the VBE menu, click **VBA Code Writer**.
3. Choose a command (for example **List Tables**).
4. In the search form, type to filter and select an item.
5. Keep **Insert Into Code** checked to paste into code at the current cursor position, or use clipboard/open options as needed.
6. For object drill-down, run **List Tables** or **List Queries** and enable **Show List of Fields**.

## Parse SQL workflow

1. Copy SQL text (or paste into the unformatted SQL box).
2. Click **Generate** to create VBA-ready SQL string code.
3. Optionally enable **Declare Variable** to include `Dim stringSQLText As String`.
4. Click **Insert Code** to insert the generated SQL code at the current line in the active module.

## Build and setup notes

- Project type: VB.NET class library targeting **.NET Framework 4.7**.
- Office interop references are configured for Access/DAO/VBE.
- COM interop registration is enabled in project settings.

### Verified build status (CLI)

`dotnet build` currently fails in a clean environment unless prerequisites are installed:

1. Install the **.NET Framework 4.7 Developer Pack** (targeting pack).
2. Note that `VBACodeWriterSetup.vdproj` is a Visual Studio Setup Project and is **not supported by MSBuild/dotnet CLI**.

Use full Visual Studio (with the setup project extension if needed) for installer packaging.

## Repository structure

- `VBACodeWriter/` - add-in source code (menu commands, search UI, SQL parser).
- `VBACodeWriterSetup/` - setup project (`.vdproj`) for installer packaging.
- `VBACodeWriter.sln` - solution file.
