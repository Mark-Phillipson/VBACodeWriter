Imports Extensibility
Imports System.Runtime.InteropServices
Imports Access = Microsoft.Office.Interop.Access
#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the VBACodeWriterSetup project,
' right click the project in the Solution Explorer, then choose install.
' The setup projects no longer exist or work with the latest version of visual studio therefore:
' You will need to run the setup file "C:\Users\MPhil\Source\Repos\VBACodeWriterHelper\VBACodeWriterSetup\Debug\VBACodeWriterSetup.msi"
#End Region

<GuidAttribute("F3CF5FDA-1C85-479C-9A9E-4EFFCE6B2293"), ProgIdAttribute("VBACodeWriter.Connect")> _
Public Class Connect

    Implements Extensibility.IDTExtensibility2

    Private applicationObject As Access.Application
    Private addInInstance As Microsoft.Office.Core.COMAddIn
    Private WithEvents MenuCommandForms As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandReports As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandControls As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandVariables As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandTables As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandQueries As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandFields As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandParseSQL As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandDimensionVariable As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandListProcedures As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandListModules As Microsoft.Office.Core.CommandBarButton
    Private WithEvents MenuCommandListAllProcedures As Microsoft.Office.Core.CommandBarButton
    Private ObjectSetting As New ObjectSettings


    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub

    Private Sub DestroyComObject(MyObject As Object)
        Dim IntegerReferenceCount As Integer
        Do
            IntegerReferenceCount = _
             System.Runtime.InteropServices.Marshal.ReleaseComObject(MyObject)
        Loop While IntegerReferenceCount > 0
        MyObject = Nothing

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
        If Not MenuCommandControls Is Nothing Then MenuCommandControls.Delete() : DestroyComObject(MenuCommandControls)
        If Not MenuCommandDimensionVariable Is Nothing Then MenuCommandDimensionVariable.Delete() : DestroyComObject(MenuCommandDimensionVariable)
        If Not MenuCommandFields Is Nothing Then MenuCommandFields.Delete() : DestroyComObject(MenuCommandFields)
        If Not MenuCommandForms Is Nothing Then MenuCommandForms.Delete() : DestroyComObject(MenuCommandForms)
        If Not MenuCommandParseSQL Is Nothing Then MenuCommandParseSQL.Delete() : DestroyComObject(MenuCommandParseSQL)
        If Not MenuCommandQueries Is Nothing Then MenuCommandQueries.Delete() : DestroyComObject(MenuCommandQueries)
        If Not MenuCommandReports Is Nothing Then MenuCommandReports.Delete() : DestroyComObject(MenuCommandReports)
        If Not MenuCommandTables Is Nothing Then MenuCommandTables.Delete() : DestroyComObject(MenuCommandTables)
        If Not MenuCommandVariables Is Nothing Then MenuCommandVariables.Delete() : DestroyComObject(MenuCommandVariables)
        If Not MenuCommandListProcedures Is Nothing Then MenuCommandListProcedures.Delete() : DestroyComObject(MenuCommandListProcedures)
        If Not MenuCommandListModules Is Nothing Then MenuCommandListModules.Delete() : DestroyComObject(MenuCommandListModules)
        If Not MenuCommandListAllProcedures Is Nothing Then MenuCommandListAllProcedures.Delete() : DestroyComObject(MenuCommandListAllProcedures)
        'If Not SearchForm Is Nothing Then SearchForm.Close() : SearchForm = Nothing
        If Not ObjectSetting Is Nothing Then ObjectSetting = Nothing


        'If Not basInsertCode.ObjectSetting Is Nothing Then
        '    basInsertCode.ObjectSetting = Nothing
        'End If
        If Not applicationObject Is Nothing Then
            DestroyComObject(applicationObject.VBE)            'applicationObject.CurrentProject.CloseConnection()
            applicationObject.Quit(Access.AcQuitOption.acQuitSaveAll)
            DestroyComObject(applicationObject)
        End If
        DestroyComObject(addInInstance)

        'If Not basInsertCode.applicationObject Is Nothing Then
        '    basInsertCode.applicationObject = Nothing
        'End If
        'applicationObject.VBE.ActiveWindow.Close()

        'applicationObject.CurrentProject.CloseConnection()
        'DestroyComObject(applicationObject.CurrentProject)f
        'DestroyComObject(applicationObject.Forms)
        'DestroyComObject(applicationObject.Reports)
        'DestroyComObject(applicationObject.DoCmd)
        'DestroyComObject(applicationObject.DBEngine)

        Dim integerCounter As Integer
        Dim p As System.Diagnostics.Process
        Dim stringTemporary As String = ""

        For Each p In System.Diagnostics.Process.GetProcesses()
            '            lstProcesses.Items.Add(p.ProcessName & " - " & p.Id.ToString())

            stringTemporary = stringTemporary & p.ProcessName & " / "
            If LCase(p.ProcessName) = "msaccess" Then
                On Error Resume Next
                If p.MainWindowTitle.Length = 0 Then
                    p.Kill()
                    'Else
                    'If MessageBox.Show("Kill " & p.MainWindowTitle & " Confirm?", "Question: " & p.MainWindowTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) _
                    '        = Windows.Forms.DialogResult.Yes Then
                    '    p.Kill()
                    '    integerCounter = integerCounter + 1
                    'End If
                End If
                On Error GoTo 0
            End If

        Next


    End Sub

    Public Sub OnCononection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        applicationObject = CType(application, Microsoft.Office.Interop.Access.Application)
        addInInstance = CType(addInInst, Microsoft.Office.Core.COMAddIn)
        addInInstance.Object = Me
        Dim CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                addInInst, "List Forms", "List Forms", 0, "List Forms")
        MenuCommandForms = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                addInInst, "List Reports", "List Reports", 0, "List Reports")
        MenuCommandReports = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                addInInst, "List Controls", "List Controls", 0, "List Controls")
        MenuCommandControls = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                addInInst, "List Variables", "List Variables", 0, "List Variables")
        MenuCommandVariables = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                                                             addInInst, "Dimension Variable", "Dimension Variable", 0, "&Dimension Variable")
        MenuCommandDimensionVariable = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                         addInInst, "List Tables", "List Tables", 0, "List Tables")
        MenuCommandTables = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                         addInInst, "List Queries", "List Queries", 0, "List Queries")
        MenuCommandQueries = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                         addInInst, "List Fields", "List Fields", 0, "List Fields")
        MenuCommandFields = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
                         addInInst, "Parse SQL", "Parse SQL", 0, "Parse SQL")
        MenuCommandParseSQL = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
        addInInst, "List Procedures", "List Procedures", 0, "List Procedures")
        MenuCommandListProcedures = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)
        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
        addInInst, "List Modules", "List Modules", 0, "List Modules")
        MenuCommandListModules = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)

        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
        addInInst, "List All Procedures", "List All Procedures", 0, "List All Procedures")
        MenuCommandListAllProcedures = CType(CommandBar, Microsoft.Office.Core.CommandBarButton)

        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
        addInInst, "Comment Block", "Comment Block", 0, "Comment Block")

        CommandBar = AddMenuCommand(CType(applicationObject.VBE.CommandBars("Menu Bar"), Microsoft.Office.Core.CommandBar),
        addInInst, "UnComment Block", "UnComment Block", 0, "UnComment Block")

    End Sub

    Public Sub ListForms(ByVal BooleanVBA As Boolean)
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("Form")

        searchForm.SetFormOrReport("Form")

        'ObjectSetting = New ObjectSettings
        ObjectSetting.SetFormOrReport("Form")
        ObjectSetting.SetObjectType("Form")
        ShowSearchForm(BooleanVBA, searchForm)

    End Sub
    Private Sub MenuCommandForms_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandForms.Click

        ListForms(True)
    End Sub

    Private Sub ShowSearchForm(ByVal BooleanVBA As Boolean, ByRef searchForm As SearchForm2, Optional stringModuleName As String = "")
        searchForm.AccessInstance1 = applicationObject
        searchForm.StringDatabaseFilename1 = applicationObject.CurrentDb.Name
        searchForm.ObjectSetting1 = ObjectSetting

        searchForm.ShowDialog()
        If searchForm.ObjectType = "Procedure" And searchForm.InsertIntoCodeCheckbox.Checked = False And stringModuleName.Length > 0 Then
            applicationObject.DoCmd.OpenModule(stringModuleName, searchForm.ObjectsListbox.Text)
        End If
        If searchForm.ObjectType <> "Control" And searchForm.ObjectType <> "Field" And searchForm.ObjectType <> "Variable" And searchForm.ObjectType <> "Module" Then
            ObjectSetting.SetObjectName(searchForm.ObjectsListbox.Text)
        Else
            ObjectSetting.SetObjectName(searchForm.GetObjectName)
        End If
        If searchForm.ObjectType = "Module" And searchForm.ObjectsListbox.Text.Length > 0 Then
            applicationObject.DoCmd.OpenModule(searchForm.ObjectsListbox.Text)
            Exit Sub
        End If
        If searchForm.ObjectType = "AllProcedure" Then
            If searchForm.OpenObjectCheckbox.Checked Then
                applicationObject.DoCmd.OpenModule(Left(searchForm.ObjectsListbox.Text, InStr(searchForm.ObjectsListbox.Text, ".") - 1),
                                                   Mid(searchForm.ObjectsListbox.Text, InStr(searchForm.ObjectsListbox.Text, ".") + 1))
                Exit Sub
            End If
        End If
        If searchForm.InsertIntoCodeCheckbox.Checked And BooleanVBA Then
            If searchForm.ObjectType = "AllProcedure" Then
                InsertValueIntoCode((Mid(searchForm.ObjectsListbox.Text, InStr(searchForm.ObjectsListbox.Text, ".") + 1)), applicationObject, ObjectSetting)
            ElseIf searchForm.ObjectType = "Procedure" Then
                applicationObject.DoCmd.OpenModule(ObjectSetting.GetOriginalModule())
                InsertValueIntoCode((searchForm.ObjectsListbox.Text), applicationObject, ObjectSetting)
            Else
                InsertValueIntoCode((searchForm.ObjectsListbox.Text), applicationObject, ObjectSetting)

            End If
        ElseIf searchForm.InsertIntoCodeCheckbox.Checked = False Then
            'AccessInst.Screen.ActiveControl = searchForm.ListBoxObjects.Text
        Else
            'SendKeys.SendWait(searchForm.ListBoxObjects.Text)
        End If
        If searchForm.PlaceinClipboardCheckbox.Checked Then
            On Error Resume Next
            Clipboard.Clear()
            Clipboard.SetText(searchForm.ObjectsListbox.Text)
            On Error GoTo 0
        End If
        'searchForm.Visible = True
        'searchForm.Close()
        'searchForm = Nothing

    End Sub

    Public Sub Listreports(ByVal BooleanVBA As Boolean)
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        searchForm.SetObjectName(ObjectSetting.GetObjectname)

        searchForm.SetObjectType("Report")
        searchForm.SetFormOrReport("Report")
        ObjectSetting = New ObjectSettings
        ObjectSetting.SetFormOrReport("Report")
        ObjectSetting.SetObjectType("Report")
        ShowSearchForm(BooleanVBA, searchForm)
    End Sub
    Private Sub MenuCommandReports_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandReports.Click
        Listreports(True)

    End Sub
    Public Sub ListTables(ByVal BooleanVBA As Boolean)

        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        searchForm.SetObjectName("")
        searchForm.SetObjectType("Table")
        searchForm.SetFormOrReport("")

        'ObjectSetting = New ObjectSettings
        ObjectSetting.SetObjectType("Table")
        ObjectSetting.LastTableorQueryObjectType = "Table"
        ShowSearchForm(BooleanVBA, searchForm)
        If searchForm.ShowFieldsCheckbox.Checked Then
            ListFields(BooleanVBA)
        End If
    End Sub
    Private Sub MenuCommandTables_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandTables.Click
        ListTables(True)
    End Sub

    Public Sub ListQueries(ByVal BooleanVBA As Boolean)
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        searchForm.SetObjectName("")
        searchForm.SetObjectType("Query")
        searchForm.SetFormOrReport("")
        ObjectSetting = New ObjectSettings
        ObjectSetting.SetObjectType("Query")
        ObjectSetting.LastTableorQueryObjectType = "Query"
        ShowSearchForm(BooleanVBA, searchForm)
        If searchForm.ShowFieldsCheckbox.Checked Then
            ListFields(BooleanVBA)
        End If
    End Sub
    Private Sub MenuCommandQueries_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandQueries.Click
        ListQueries(True)
    End Sub

    Public Sub ListControls(ByVal BooleanVBA As Boolean)
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        If ObjectSetting.GetFormorReport Is Nothing Then Exit Sub
        If ObjectSetting.GetObjectname Is Nothing Then Exit Sub
        If Not ObjectSetting.GetObjectType = "Form" And Not ObjectSetting.GetObjectType = "Report" Then Exit Sub
        searchForm.SetFormOrReport(ObjectSetting.GetFormorReport)
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("Control")
        ShowSearchForm(BooleanVBA, searchForm)

    End Sub
    Private Sub MenuCommandControls_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandControls.Click

        ListControls(True)
    End Sub

    Public Sub ListFields(ByVal BooleanVBA As Boolean)
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        ' Need new property to record the last table or query object type
        If ObjectSetting.GetObjectname Is Nothing Then Exit Sub
        ObjectSetting.SetObjectType(ObjectSetting.LastTableorQueryObjectType)
        If Not ObjectSetting.GetObjectType = "Table" And Not ObjectSetting.GetObjectType = "Query" Then Exit Sub
        searchForm.SetFormOrReport(ObjectSetting.GetFormorReport)
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("Field")
        searchForm.SetFormOrReport("")
        ShowSearchForm(BooleanVBA, searchForm)
    End Sub
    Private Sub MenuCommandFields_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandFields.Click
        ListFields(True)

    End Sub

    Private Sub MenuCommandVariables_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandVariables.Click
        Dim CM As Microsoft.Vbe.Interop.CodeModule
        Dim startLine As Integer
        Dim startColumn As Integer
        Dim EndLine As Integer
        Dim EndColumn As Integer
        Dim strProcedure As String
        Dim lngLine As Integer
        Dim strVariable As String
        Dim varArray(1000) As String
        Dim searchForm As New SearchForm2
        'ObjectSetting = New ObjectSettings
        'MessageBox.Show(ObjectSetting.GetObjectname)
        CM = applicationObject.VBE.ActiveCodePane.CodeModule
        applicationObject.VBE.ActiveCodePane.GetSelection(startLine, startColumn, EndLine, EndColumn)
        strProcedure = CM.ProcOfLine(startLine, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)
        If Len(strProcedure) = 0 Then

            MsgBox("To list variables the cursor has to be inside a procedure." & vbCr & "" & vbCr & "Process will now abort!" _
            , MsgBoxStyle.Information _
            , "Process is Now Aborting...")
            Exit Sub
        End If
        For k = 1 To UBound(varArray)
            varArray(k) = Nothing
        Next
        ' Get variables from the module declaration area before the procedure
        lngLine = 1
        Do Until Trim(CM.Lines(lngLine, 1)).Contains("Sub") Or Trim(CM.Lines(lngLine, 1)).Contains("Function") Or Trim(CM.Lines(lngLine, 1)).Contains("Property")
            strVariable = ""
            ExtractVariable(CM.Lines(lngLine, 1), varArray)
            lngLine = lngLine + 1
        Loop
        lngLine = CM.ProcStartLine(strProcedure, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)
        'lngLine = lngLine - 1
        'frmListOfVariables.cboVariables.Clear()
        Do Until Trim(CM.Lines(lngLine, 1)) = "End Sub" Or Trim(CM.Lines(lngLine, 1)) = "End Function"
            strVariable = ""
            ExtractVariable(CM.Lines(lngLine, 1), varArray)
            lngLine = lngLine + 1

        Loop
        searchForm = New SearchForm2
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("Variable")
        searchForm.SetArray(varArray)


        ObjectSetting.SetObjectType("Variable")

        ShowSearchForm(True, searchForm)

        CM = Nothing



    End Sub

    Private Sub AddVarItem(ByVal strVar As String, ByVal varArray As Object)
        ObjectSetting.LineNumber = ObjectSetting.LineNumber + 1
        If InStr(strVar, ")") > 0 And InStr(strVar, "(") > 0 Then
            strVar = Left(strVar, InStr(strVar, "(") - 1)
            varArray(ObjectSetting.LineNumber) = strVar & "()"
            Exit Sub
        End If
        If InStr(strVar, "(") > 0 Then Exit Sub
        If InStr(strVar, ")") > 0 Then Exit Sub
        If InStr(strVar, " ") > 0 Then Exit Sub
        varArray(ObjectSetting.LineNumber) = strVar
    End Sub

    Private Sub MenuCommandParseSQL_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandParseSQL.Click
        Dim FormParseSQL As New FormParseSQL
        FormParseSQL.AccessInstance1 = applicationObject
        FormParseSQL.ShowDialog()
        FormParseSQL = Nothing
    End Sub

    Private Sub MenuCommandDimensionVariable_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandDimensionVariable.Click
        Dim strSelection As String
        Dim CM As Microsoft.Vbe.Interop.CodeModule
        Dim startLine As Integer
        Dim startColumn As Integer
        Dim EndLine As Integer
        Dim EndColumn As Integer
        Dim strProcedure As String
        Dim lngLine As Integer
        Dim strAsSection As String

        On Error GoTo HandleErr

        'Get active codemodule
        CM = applicationObject.VBE.ActiveCodePane.CodeModule
        applicationObject.VBE.ActiveCodePane.GetSelection(startLine, startColumn, EndLine, EndColumn)
        strSelection = applicationObject.VBE.ActiveCodePane.CodeModule.Lines(startLine, EndLine + 1 - startLine)
        strSelection = Trim(Mid(strSelection, startColumn, EndColumn - startColumn))
        If Right(strSelection, 1) = "=" Then strSelection = Left(strSelection, Len(strSelection) - 1)
        If Len(strSelection) = 0 Then
            MsgBox("Please select a variable that has not been declared before using this function", vbInformation, "Dimension a Selected Variable")
            Exit Sub
        End If
        strProcedure = CM.ProcOfLine(startLine, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)

        lngLine = CM.ProcStartLine(strProcedure, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)
        Do Until (Len(Trim(CM.Lines(lngLine, 1))) > 0 And Left(CM.Lines(lngLine, 1), 1) <> "'" And Right(CM.Lines(lngLine, 1), 1) <> "_")
            lngLine = lngLine + 1
        Loop
        Select Case Left(strSelection.ToLower, 3)
            Case "str"
                strAsSection = "As String"
            Case "lng", "lon"
                strAsSection = "As Long"
            Case "int"
                strAsSection = "As Integer"
            Case "dbl", "dou"
                strAsSection = "As Double"
            Case "bln", "boo"
                strAsSection = "As Boolean"
            Case "var"
                strAsSection = "As Variant"
            Case "sng", "sin"
                strAsSection = "As Single"
            Case "obj"
                strAsSection = "As Object"
            Case "frm", "for"
                strAsSection = "As Access.Form"
            Case "rpt", "rep"
                strAsSection = "As Access.Report"
            Case "cur"
                strAsSection = "As Currency"
            Case "dte", "dat"
                strAsSection = "As Date"
            Case "cmd", "com"
                strAsSection = "As ADODB.Command"
            Case Else
                Select Case strSelection.ToLower ' Specific variables
                    Case "k", "l", "m", "i"
                        strAsSection = "As Integer"
                    Case "rs", "rst", "rd"
                        strAsSection = "As DAO.Recordset"
                    Case "db"
                        strAsSection = "As DAO.Database"
                    Case Else
                        strAsSection = "As Variant"
                End Select
                Select Case Left(strSelection, 2)
                    Case "rs"
                        strAsSection = "As DAO.Recordset"
                End Select
        End Select
        CM.InsertLines(lngLine + 1, "    Dim " & strSelection.Trim & " " & Trim(strAsSection))
        '& " '" & Format(My.Computer.Clock.LocalTime, "hh:nn"))

ExitHere:
        Exit Sub

HandleErr:
        Select Case Err.Number
            'Case # '
            Case Else
                MsgBox("Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "Dimension Variable")
                Resume ExitHere
        End Select
        Resume 'Debug only

    End Sub
    Public Function TestCalling() As String
        MessageBox.Show("Test")


        Return Now.ToString
    End Function
    Private Sub ExtractVariable(ByRef StringLine As String, ByVal varArray() As Object)
        Dim stringVariable As String = ""
        Dim stringTemporary As String = ""
        Dim stringTemporary2 As String = ""
        If Left(Trim(StringLine), 4) = "Dim " Then

            stringVariable = Mid(Trim(StringLine), 5)

        End If
        If Left(Trim(StringLine), 7) = "Static " Then

            stringVariable = Mid(Trim(StringLine), 7)

        End If
        If StringLine.StartsWith("Global") Then
            stringVariable = StringLine.Substring(7)
        End If
        If StringLine.StartsWith("Private") And (Not StringLine.Contains("Sub") And Not StringLine.Contains("Function")) Then
            stringVariable = StringLine.Substring(8)
        End If
        If InStr(Trim(StringLine), " As ") > 0 Then
            If Len(stringVariable) = 0 Then stringVariable = Trim(StringLine)
            If InStr(stringVariable, " As ") > 0 Then
                stringTemporary = Mid(stringVariable, InStr(stringVariable, " As ") + 4)
                If InStr(stringTemporary, ",") > 0 Then stringTemporary = Trim(Mid(stringTemporary, InStr(stringTemporary, ",") + 1))
                stringVariable = Left(stringVariable, InStr(stringVariable, " As ") - 1)
                Do While InStr(stringTemporary, " As ") > 0
                    stringTemporary2 = Left(stringTemporary, InStr(stringTemporary, " As ") - 1)
                    AddVarItem(stringTemporary2, varArray)
                    If InStr(stringTemporary, " As ") > 0 Then
                        stringTemporary = Mid(stringTemporary, InStr(stringTemporary, " As ") + 4)
                    Else
                        Exit Do
                    End If
                    If InStr(stringTemporary, ",") > 0 Then stringTemporary = Trim(Mid(stringTemporary, InStr(stringTemporary, ",") + 1))
                Loop
            End If
            If InStr(stringVariable, "Function ") > 0 And InStr(stringVariable, "(") > 0 Then
                stringVariable = Mid(stringVariable, InStr(stringVariable, "(") + 1)
            End If
            If InStr(stringVariable, "Sub ") > 0 And InStr(stringVariable, "(") > 0 Then
                stringVariable = Mid(stringVariable, InStr(stringVariable, "(") + 1)
            End If
            AddVarItem(Trim(stringVariable), varArray)
        Else 'variable may have been defined without a type
            If Len(stringVariable) > 0 Then

                AddVarItem(Trim(stringVariable), varArray)
            End If
        End If
    End Sub


    Private Sub MenuCommandListProcedures_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandListProcedures.Click
        ListProcedures()
    End Sub

    Private Sub ListProcedures()
        Dim varArray(1000) As String
        Dim searchForm As New SearchForm2

        Dim CurrentModule As Microsoft.Vbe.Interop.CodeModule
        Dim lngCount As Long, lngCountDecl As Long, lngI As Long
        Dim strProcName As String, astrProcNames() As String
        Dim intI As Integer
        Dim lngR As Long
        CurrentModule = applicationObject.VBE.ActiveCodePane.CodeModule
        Dim stringModuleName As String = CurrentModule.Name
        ' Open specified Module object.
        If Left(stringModuleName, 5) = "Form_" Then
            applicationObject.DoCmd.OpenForm(Mid(stringModuleName, 6), Microsoft.Office.Interop.Access.AcFormView.acDesign)
        ElseIf Left(stringModuleName, 7) = "Report_" Then
            applicationObject.DoCmd.OpenReport(Mid(stringModuleName, 8), Microsoft.Office.Interop.Access.AcView.acViewDesign)
        End If
        applicationObject.DoCmd.OpenModule(stringModuleName)
        ' Return reference to Module object.
        'CurrentModule = applicationObject.VBE.Modules(stringModuleName)
        ' Count lines in module.
        lngCount = CurrentModule.CountOfLines
        ' Count lines in Declaration section in module.
        lngCountDecl = CurrentModule.CountOfDeclarationLines
        ' Determine name of first procedure.
        strProcName = CurrentModule.ProcOfLine(lngCountDecl + 1, lngR)
        ' Initialize counter variable.
        intI = 0
        ' Redimension array.
        ReDim Preserve astrProcNames(intI)
        ' Store name of first procedure in array.
        astrProcNames(intI) = strProcName
        ' Determine procedure name for each line after declarations.
        For lngI = lngCountDecl + 1 To lngCount
            ' Compare procedure name with ProcOfLine property value.
            If strProcName <> CurrentModule.ProcOfLine(lngI, lngR) Then
                ' Increment counter.
                intI = intI + 1
                strProcName = CurrentModule.ProcOfLine(lngI, lngR)
                ReDim Preserve astrProcNames(intI)
                ' Assign unique procedure names to array.
                astrProcNames(intI) = strProcName
            End If
        Next lngI

        'Call adh_accSortStringArray(astrProcNames())
        'astrProcNames
        'For intI = 0 To UBound(astrProcNames)
        '    strMsg = strMsg & astrProcNames(intI) & ";"
        'Next intI
        varArray = astrProcNames




        'varArray = basInsertCode.AllProcs(CM.Name, addInInstance)
        searchForm = New SearchForm2
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("Procedure")
        searchForm.SetArray(varArray)


        ObjectSetting.SetObjectType("Procedure")

        ShowSearchForm(True, searchForm, stringModuleName)


        CurrentModule = Nothing
    End Sub

    Private Sub MenuCommandListModules_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandListModules.Click
        Dim searchForm As New SearchForm2
        searchForm = New SearchForm2
        searchForm.SetObjectName("")
        searchForm.SetObjectType("Module")
        searchForm.SetFormOrReport("")

        'ObjectSetting = New ObjectSettings
        Dim stringActiveModuleName As String = applicationObject.VBE.ActiveCodePane.CodeModule.Name
        ObjectSetting.SetOriginalModule(stringActiveModuleName)
        ObjectSetting.SetObjectType("Module")
        ObjectSetting.LastTableorQueryObjectType = ""
        ShowSearchForm(False, searchForm)
        ListProcedures()
    End Sub

    Private Sub MenuCommandListAllProcedures_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles MenuCommandListAllProcedures.Click
        Dim varArray(1000) As String
        Dim searchForm As New SearchForm2

        'Dim CurrentModule As Microsoft.Vbe.Interop.CodeModule
        Dim CurrentModule As Microsoft.Office.Interop.Access.Module
        Dim lngCount As Long, lngCountDecl As Long, lngI As Long
        Dim strProcName As String, astrProcNames() As String
        Dim intI As Integer
        Dim lngR As Long
        Dim stringActiveModuleName As String = applicationObject.VBE.ActiveCodePane.CodeModule.Name
        Dim AccessObject As Microsoft.Office.Interop.Access.AccessObject
        For Each AccessObject In applicationObject.CurrentProject.AllModules
            Dim stringModuleName As String = AccessObject.FullName
            applicationObject.DoCmd.OpenModule(stringModuleName)
            CurrentModule = applicationObject.Modules(AccessObject.FullName)
            ' Open specified Module object.
            'If Left(stringModuleName, 5) = "Form_" Then
            '    applicationObject.DoCmd.OpenForm(Mid(stringModuleName, 6), Microsoft.Office.Interop.Access.AcFormView.acDesign)
            'ElseIf Left(stringModuleName, 7) = "Report_" Then
            '    applicationObject.DoCmd.OpenReport(Mid(stringModuleName, 8), Microsoft.Office.Interop.Access.AcView.acViewDesign)
            'End If
            ' Return reference to Module object.
            'CurrentModule = applicationObject.VBE.Modules(stringModuleName)
            ' Count lines in module.
            lngCount = CurrentModule.CountOfLines
            ' Count lines in Declaration section in module.
            lngCountDecl = CurrentModule.CountOfDeclarationLines
            ' Determine name of first procedure.
            strProcName = CurrentModule.ProcOfLine(lngCountDecl + 1, lngR)
            ' Initialize counter variable.
            'intI = 0
            ' Redimension array.
            ReDim Preserve astrProcNames(intI)
            ' Store name of first procedure in array.
            astrProcNames(intI) = stringModuleName & "." & strProcName
            ' Determine procedure name for each line after declarations.
            For lngI = lngCountDecl + 1 To lngCount
                ' Compare procedure name with ProcOfLine property value.
                If strProcName <> CurrentModule.ProcOfLine(lngI, lngR) Then
                    ' Increment counter.
                    intI = intI + 1
                    strProcName = CurrentModule.ProcOfLine(lngI, lngR)
                    ReDim Preserve astrProcNames(intI)
                    ' Assign unique procedure names to array.
                    astrProcNames(intI) = stringModuleName & "." & strProcName
                End If
            Next lngI

            'Call adh_accSortStringArray(astrProcNames())
            'astrProcNames
            'For intI = 0 To UBound(astrProcNames)
            '    strMsg = strMsg & astrProcNames(intI) & ";"
            'Next intI
            If Not stringModuleName = stringActiveModuleName Then
                applicationObject.DoCmd.Close(Microsoft.Office.Interop.Access.AcObjectType.acModule, stringModuleName, Access.AcCloseSave.acSaveYes)
            End If
        Next
        varArray = astrProcNames




        'varArray = basInsertCode.AllProcs(CM.Name, addInInstance)
        searchForm = New SearchForm2
        searchForm.SetObjectName(ObjectSetting.GetObjectname)
        searchForm.SetObjectType("AllProcedure")
        searchForm.SetArray(varArray)


        ObjectSetting.SetObjectType("AllProcedure")

        ShowSearchForm(True, searchForm)
        'applicationObject.DoCmd.OpenModule(stringModuleName, searchForm.ObjectsListbox.Text)
        CurrentModule = Nothing

    End Sub
End Class
