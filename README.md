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
- **Procedure Chat** - opens a modeless OpenAI chat sidebar for the procedure being edited.

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

## Procedure Chat

1. Open an editable code module and place the cursor inside a `Sub`, `Function`, or `Property` procedure.
2. Choose **Procedure Chat** from the VBE menu bar. The sidebar sits inside the right edge of the editing workspace, leaving the editor usable.
3. The add-in reads `OPENAI_API_KEY` from the Access process environment. Set it before starting Access, or enter your key directly in the masked **API Key** field and click **Use Key**. A manually entered key is kept in memory for this add-in session only; it is not written to settings, the registry, or logs. Do not put keys in chat messages.
4. Enter a message and click **Send** (or press `Ctrl+Enter`). Each send captures a fresh snapshot of the active procedure, including its module/name and property accessor kind. The default model is `gpt-4.1-mini`; the **Model** field accepts another text/chat-compatible OpenAI model available to your account.
5. Responses display basic Markdown: headings, emphasis, lists, quotes, links as text, and monospaced code blocks. HTML remains inert text; scripts and remote images are not executed or loaded.
6. **Copy Response** copies the complete latest response as Markdown, including responses with no code. Select a code block in the dropdown and use **Copy Code** for just that snippet, without Markdown fences. Paste it manually wherever needed, or use **Insert Code** to insert it before the current VBE cursor line after confirming the target module and line. Insertion never replaces existing code automatically.

**Cancel** stops a pending request; a request also times out after two minutes. Failed or cancelled requests leave your message available for retry. **New Chat** clears conversation history without clearing the session key. **Close** hides the sidebar and restores the editor width; reopening it preserves the chat and key until the add-in is unloaded or Access exits.

The sidebar is a Windows child panel attached to the VBE's native MDI workspace, not a native draggable VBE tool window. It follows workspace resizing and temporarily hides when the workspace is too small (below approximately 650 x 420 logical pixels). This requires an MDI-based Windows VBE and access to the VBA project object model. Protected projects and other VBE versions may not support all actions.

### Privacy and limits

Sending transmits the current procedure, your message, and previous turns (including earlier procedure snapshots) to `https://api.openai.com/v1/chat/completions`. Only send code you are authorized to share; remove secrets and sensitive data first. API billing and OpenAI's data policies apply. Requests specify `store: false`, which is not a guarantee of zero provider-side retention. No requests are sent merely by opening the panel.

Conversation text is limited to 120,000 characters per request and responses to 4,096 completion tokens. Truncated responses are labelled. Start a new chat when the context limit is reached. Review and back up code before insertion; generated code is never executed by the chat panel.

### Build, reload, and verify

Restore packages with NuGet and build from a Visual Studio Developer Command Prompt:

```powershell
nuget restore VBACodeWriter\packages.config -PackagesDirectory packages
MSBuild VBACodeWriter\VBACodeWriter.vbproj /p:RegisterForComInterop=false /p:RunCodeAnalysis=false
```

Close Access before rebuilding the DLL it has loaded, then reopen Access to get the new **Procedure Chat** command. This command assumes the add-in is already registered. For installer deployment, refresh the setup project's detected dependencies and include `Markdig.dll`, `System.Memory.dll`, `System.Buffers.dll`, `System.Numerics.Vectors.dll`, and `System.Runtime.CompilerServices.Unsafe.dll` alongside the add-in.

For validation while Access remains open, build to a separate output folder and run the no-network smoke harness:

```powershell
MSBuild tests\ChatSmokeTests.vbproj /p:RegisterForComInterop=false /p:RunCodeAnalysis=false /p:OutputPath=bin\ChatValidation\
.\tests\bin\ChatValidation\ChatSmokeTests.exe --inspect-access
```

The harness checks mocked OpenAI requests, cancellation, error handling, Markdown, snippet extraction, key handling, and sidebar layout/cleanup in a temporary window. The optional `--inspect-access` flag checks only the running VBE's window classes; it does not read procedure text, modify a database, or call OpenAI. A final manual check in Access is still needed for procedure capture, editor focus/keyboard handling, insertion, and a real API response.

### Connection diagnostics

The HTTP client explicitly uses TLS 1.2 because Office-hosted .NET can inherit legacy TLS defaults. This setting applies only to the chat client; it does not change Access-wide networking settings or disable certificate validation. Network errors distinguish TLS handshake, certificate trust, DNS, proxy resolution, and connection failures without exposing keys or request contents.

After building the smoke harness, run this optional connectivity check:

```powershell
.\tests\bin\ChatValidation\ChatSmokeTests.exe --check-network
```

Unlike the default smoke tests, this command sends unauthenticated HTTPS GET requests to OpenAI's `/v1/models` endpoint. It sends no API key, chat, or procedure text. It compares system defaults, legacy TLS 1.0, and TLS 1.2, then checks the production client under simulated legacy host defaults. **HTTP 401 is expected** and confirms the HTTPS connection reached the API; it does not test your key or model access. Save and close Access, run the **Build and deploy VBACodeWriter locally** task, and reopen Access before retrying a chat with a revised build.

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

- Project type: VB.NET add-in/class library targeting **.NET Framework 4.8** with a newer VS Code extension-oriented workflow.
- Office interop references are configured for Access/DAO/VBE.
- COM interop registration is enabled in project settings.
- A dedicated setup project was added to support packaging and deployment workflows.

### ✅ Verified build status (CLI)

`dotnet build` currently fails in a clean environment unless prerequisites are installed:

1. Install the **.NET Framework 4.8 Developer Pack** (targeting pack).
2. Use full Visual Studio for any setup or installer work because the setup packages are Visual Studio project types and are **not supported by MSBuild/dotnet CLI**.
3. If using the installer extension, install the **Visual Studio Installer Projects** extension to build the `.vdproj` packaging projects correctly.

## 📁 Repository structure

- `VBACodeWriter/` -  the jukeAccess/VBE add-in source code (menu commands, search UI, SQL parser, procedure helpers).
- `SetupNew/` - new setup project for the updated packaging and installation flow.
- `VBACodeWriterSetup/` - existing installer project (`.vdproj`) for packaging and deployment support.
- `VBACodeWriter.sln` - solution file for the current project set.
