# VBACodeWriter 🚀

VBACodeWriter is a VBA productivity toolkit for Microsoft Access that now follows a VS Code extension-based workflow while still supporting the Access/VBE tooling and installer packaging needed for local deployment.

The project combines the Access add-in experience with a newer extension-focused setup and a dedicated setup project for packaging and installation.

## ✨ What it offers

The add-in adds a **VBA Code Writer** menu to the Access VBA editor (VBE) with these commands:

- **List Forms** 📄 - list form names from the current database.
- **List Reports** 🧾 - list report names.
- **List Controls** 🧩 - list controls from the previously selected form/report.
- **List Tables** 🗃️ - list table names.
- **List Queries** 🔎 - list query names.
- **List Fields** 🏷️ - list fields from the currently selected table/query.
- **List Variables** 🧠 - list variables visible in the active procedure/module declarations.
- **Dimension Variable** 📝 - inserts a `Dim` statement for a selected variable name using naming-prefix inference (for example `str* -> As String`, `lng* -> As Long`, `rs* -> As DAO.Recordset`).
- **List Procedures** 🧭 - list procedures in the active module.
- **List Modules** 📦 - list modules and open one, then list its procedures.
- **List All Procedures** 🌍 - list procedures across all modules in `Module.Procedure` format.
- **Parse SQL** 🧪 - formats SQL text into VBA-ready string-building code and can insert it directly into the active module.
- **Add Error Handler** 🛡️ - inserts a standard VBA error-handling block for the selected procedure, with support for both `Sub` and `Function` procedures.
- **Select Procedure** 🎯 - selects the whole procedure in the VBE so you can review or edit it immediately.
- **Comment Block** / **UnComment Block** 💬 - adds built-in VBE command bar actions for comment/uncomment.

## 🔤 Search form behavior

Most commands open the same search dialog where you can:

- Filter with a search box.
- Select from list results.
- Insert the selected value into the current cursor/selection in the active code pane.
- Copy the selected value to clipboard.
- Open the selected object when relevant (used for modules/all procedures).
- For tables/queries, optionally continue to a field list.
- Use quick-select buttons (**Select Top**, **Select 2nd** ... **Select 9th**) to jump to common results.
- For procedure work, add an error handler or select the full procedure directly from the search form.

## 🚀 How to use

1. Open your Access database and press `Alt+F11` to open the VBA editor.
2. In the VBE menu, click **VBA Code Writer**.
3. Choose a command (for example **List Tables** or **List Procedures**).
4. In the search form, type to filter and select an item.
5. Keep **Insert Into Code** checked to paste into code at the current cursor position, or use clipboard/open options as needed.
6. For object drill-down, run **List Tables** or **List Queries** and enable **Show List of Fields**.
7. For procedure cleanup, select a procedure and use **Add Error Handler** or **Select Procedure**.

## 🧪 Parse SQL workflow

1. Copy SQL text (or paste into the unformatted SQL box).
2. Click **Generate** to create VBA-ready SQL string code.
3. Optionally enable **Declare Variable** to include `Dim stringSQLText As String`.
4. Click **Insert Code** to insert the generated SQL code at the current line in the active module.
5. The generated output is optimized for Access/VBA string concatenation and keeps SQL easier to read and maintain inside the module.

## 🛠️ Latest enhancements

Recent improvements include:

- **Procedure-aware error handling insertion** ✅
  - inserts `On Error GoTo HandleError` directly after the procedure declaration
  - supports `Sub` and `Function` templates
  - adds a standard VBA `ExitHere` / `HandleError` block with a professional message box pattern
- **Procedure selection from the search form** ✅
  - select a procedure and highlight the full procedure body in the VBA editor
  - keeps the selection active after the dialog closes
- **Parse SQL enhancements** ✅
  - cleaner VBA string generation
  - easier insertion into the active code pane
  - improved workflow for SQL-heavy Access projects

## 🏗️ Build and setup notes

- Project type: VB.NET add-in/class library targeting **.NET Framework 4.7** with a newer VS Code extension-oriented workflow.
- Office interop references are configured for Access/DAO/VBE.
- COM interop registration is enabled in project settings.
- A dedicated setup project was added to support packaging and deployment workflows.

### ✅ Verified build status (CLI)

`dotnet build` currently fails in a clean environment unless prerequisites are installed:

1. Install the **.NET Framework 4.7 Developer Pack** (targeting pack).
2. Use full Visual Studio for any setup or installer work because the setup packages are Visual Studio project types and are **not supported by MSBuild/dotnet CLI**.
3. If using the installer extension, install the **Visual Studio Installer Projects** extension to build the `.vdproj` packaging projects correctly.

## 📁 Repository structure

- `VBACodeWriter/` - Access/VBE add-in source code (menu commands, search UI, SQL parser, procedure helpers).
- `SetupNew/` - new setup project for the updated packaging and installation flow.
- `VBACodeWriterSetup/` - existing installer project (`.vdproj`) for packaging and deployment support.
- `VBACodeWriter.sln` - solution file for the current project set.
