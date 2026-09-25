Public Class SearchForm2
    Private strObjectName As String
    Private stringFormOrReport As String
    Private strObjectType As String
    Private varArray(0 To 3000) As String
    Private AccessInstance As Microsoft.Office.Interop.Access.Application
    Private StringDatabaseFilename As String
    Private ObjectSetting As New ObjectSettings
    Private WithEvents AddErrorHandlerButton As System.Windows.Forms.Button
    Private WithEvents SelectProcedureButton As System.Windows.Forms.Button
    Private blnSelectProcedureRequested As Boolean

    Public ReadOnly Property SelectProcedureRequested() As Boolean
        Get
            Return blnSelectProcedureRequested
        End Get
    End Property

    Public Property ObjectSetting1() As ObjectSettings
        Get
            Return ObjectSetting
        End Get
        Set(ByVal value As ObjectSettings)
            ObjectSetting = value
        End Set
    End Property
    Public Property AccessInstance1() As Microsoft.Office.Interop.Access.Application
        Get
            Return AccessInstance
        End Get
        Set(ByVal value As Microsoft.Office.Interop.Access.Application)
            AccessInstance = value
        End Set
    End Property
    Public Property ObjectType() As String
        Get
            Return strObjectType
        End Get
        Set(ByVal value As String)
            strObjectType = value
        End Set
    End Property

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles DoCanceButton.Click

        Me.Close()
    End Sub

    Sub SetObjectName(StringObjectNameIn As String)
        strObjectName = StringObjectNameIn
    End Sub

    Public Function GetObjectName() As String
        Return strObjectName
    End Function

    Sub SetObjectType(StringObjectTypein As String)
        strObjectType = StringObjectTypein
    End Sub

    Sub SetFormOrReport(StringFormorReportIn As String)
        stringFormOrReport = StringFormorReportIn
    End Sub

    Public Function GetFormorReport() As String
        Return stringFormOrReport
    End Function

    Private Sub OkayButton_Click(sender As Object, e As EventArgs) Handles OkayButton.Click
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If
    End Sub

    Private Sub SearchForm2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        varArray = Nothing
    End Sub

    Private Sub UpdateProcedureActionVisibility()
        Dim blnVisible As Boolean = (strObjectType = "Procedure" Or strObjectType = "AllProcedure")
        If AddErrorHandlerButton IsNot Nothing Then
            AddErrorHandlerButton.Visible = blnVisible
            AddErrorHandlerButton.Enabled = blnVisible AndAlso Me.ObjectsListbox.SelectedIndex >= 0
        End If
        If SelectProcedureButton IsNot Nothing Then
            SelectProcedureButton.Visible = blnVisible
            SelectProcedureButton.Enabled = blnVisible AndAlso Me.ObjectsListbox.SelectedIndex >= 0
        End If
    End Sub

    Private Sub SearchForm2_Load(sender As Object, e As EventArgs) Handles Me.Load
        If AddErrorHandlerButton Is Nothing Then
            AddErrorHandlerButton = New System.Windows.Forms.Button()
            AddErrorHandlerButton.Text = "Add Error Handler"
            AddErrorHandlerButton.Size = New System.Drawing.Size(152, 36)
            AddErrorHandlerButton.Location = New System.Drawing.Point(711, 340)
            AddErrorHandlerButton.BackColor = System.Drawing.SystemColors.Highlight
            AddErrorHandlerButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
            AddErrorHandlerButton.ForeColor = System.Drawing.Color.White
            Me.Controls.Add(AddErrorHandlerButton)
            AddHandler AddErrorHandlerButton.Click, AddressOf AddErrorHandlerButton_Click
        End If
        If SelectProcedureButton Is Nothing Then
            SelectProcedureButton = New System.Windows.Forms.Button()
            SelectProcedureButton.Text = "Select Procedure"
            SelectProcedureButton.Size = New System.Drawing.Size(152, 36)
            SelectProcedureButton.Location = New System.Drawing.Point(711, 385)
            SelectProcedureButton.BackColor = System.Drawing.SystemColors.Highlight
            SelectProcedureButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
            SelectProcedureButton.ForeColor = System.Drawing.Color.White
            Me.Controls.Add(SelectProcedureButton)
            AddHandler SelectProcedureButton.Click, AddressOf SelectProcedureButton_Click
        End If
        AddErrorHandlerButton.Visible = False
        AddErrorHandlerButton.Enabled = False
        SelectProcedureButton.Visible = False
        SelectProcedureButton.Enabled = False

        Dim aob As Microsoft.Office.Interop.Access.AccessObject 'This crashes when you use the full object don't know why
        Dim Table As Microsoft.Office.Interop.Access.Dao.TableDef
        Dim query As Microsoft.Office.Interop.Access.Dao.QueryDef
        Dim i As Integer
        Dim blnArrayFilled As Boolean
        'Dim objCollection As System.Collections.IEnumerable = AccessInstance.CurrentProject.AllMacros
        Dim ObjCollectionAllObjects As Microsoft.Office.Interop.Access.AllObjects = AccessInstance.CurrentProject.AllForms
        Dim ObjCollectionFields As Microsoft.Office.Interop.Access.Dao.Fields
        Dim DataBase As Microsoft.Office.Interop.Access.Dao.Database
        Me.OkayButton.Enabled = True
        DataBase = AccessInstance.Application.CurrentDb()
        If Not strObjectName Is Nothing Then
            If strObjectName.Length + 0 > 1 Then
                Me.TableQueryTextBox.Text = strObjectName
            Else
                Me.TableQueryTextBox.Text = ""
            End If
        End If

        ' If the exit sub Inserted here then the add-in closes correctly
        'If the ObjCollectionAllObjects is not set to all forms then Microsoft Access window does not close why?
        'Why does this couse the add-in to hang?
        'DataBase = AccessInstance.DBEngine.Workspaces(0).OpenDatabase(StringDatabaseFilename1)
        Me.SearchTextBox.Text = ""
        Me.ObjectsListbox.Items.Clear()
        If strObjectType Is Nothing Then
            strObjectType = ObjectSetting1.GetObjectType
        End If
        Me.LastObjectTypeTextBox.Text = ObjectSetting1.LastObjectType
        ObjectSetting1.LastObjectType = strObjectType
        Me.ShowFieldsCheckbox.Enabled = False
        Me.InsertIntoCodeCheckbox.Enabled = True
        Me.PlaceinClipboardCheckbox.Enabled = True
        Me.OpenObjectCheckbox.Enabled = False
        Select Case strObjectType
            Case "Form"
                ObjCollectionAllObjects = AccessInstance.CurrentProject.AllForms
            Case "Report"
                ObjCollectionAllObjects = AccessInstance.CurrentProject.AllReports
            Case "Control"
                If strObjectName.Length = 0 Then
                    MessageBox.Show("Form or report Name Required")
                    Exit Sub
                End If
                If stringFormOrReport = "Form" Then
                    AccessInstance.DoCmd.OpenForm(strObjectName, Microsoft.Office.Interop.Access.AcFormView.acDesign, , , , Microsoft.Office.Interop.Access.AcWindowMode.acHidden)
                    For i = 1 To UBound(varArray)
                        varArray(i) = Nothing
                    Next
                    Dim control As Microsoft.Office.Interop.Access.Control
                    i = 0
                    For Each control In AccessInstance.Forms("[" & strObjectName & "]").Controls
                        i = i + 1
                        varArray(i) = control.Name
                    Next

                    blnArrayFilled = True
                    AccessInstance.DoCmd.Close(Microsoft.Office.Interop.Access.AcObjectType.acForm, strObjectName, Microsoft.Office.Interop.Access.AcCloseSave.acSaveYes)
                ElseIf stringFormOrReport = "Report" Then
                    AccessInstance.DoCmd.OpenReport(strObjectName, Microsoft.Office.Interop.Access.AcView.acViewDesign, , , Microsoft.Office.Interop.Access.AcWindowMode.acHidden)
                    For i = 1 To UBound(varArray)
                        varArray(i) = Nothing
                    Next
                    Dim control As Microsoft.Office.Interop.Access.Control
                    i = 0
                    For Each control In AccessInstance.Reports("[" & strObjectName & "]").Controls
                        i = i + 1
                        varArray(i) = control.Name
                    Next

                    blnArrayFilled = True
                    AccessInstance.DoCmd.Close(Microsoft.Office.Interop.Access.AcObjectType.acReport, strObjectName, Microsoft.Office.Interop.Access.AcCloseSave.acSaveYes)
                End If
            Case "Field"
                If ObjectSetting1.GetObjectType = "Table" And strObjectName.Length = 0 Then
                    MessageBox.Show("Table or Query Name Required")
                    Exit Sub
                End If
                If ObjectSetting1.GetObjectType = "Table" Then
                    ObjCollectionFields = DataBase.TableDefs("[" & strObjectName & "]").Fields
                ElseIf ObjectSetting1.GetObjectType = "Query" Then
                    ObjCollectionFields = DataBase.QueryDefs("[" & strObjectName & "]").Fields
                End If
            Case "Table"
                'db.TableDefs.Refresh()
                i = 0
                For Each Table In DataBase.TableDefs
                    i = i + 1
                    varArray(i) = Table.Name

                Next
                Me.ShowFieldsCheckbox.Enabled = True
                blnArrayFilled = True
                'ObjCollectionAllObjects = AccessInstance.CurrentProject.AllForms
            Case "Query"
                i = 0
                For Each query In DataBase.QueryDefs
                    i = i + 1
                    If Not query.Name.Substring(0, 1) = "~" Then ' Do not include general SQL statements that are not queries as such
                        varArray(i) = query.Name
                    End If
                Next
                blnArrayFilled = True
                Me.ShowFieldsCheckbox.Enabled = True
                'ObjCollectionAllObjects = AccessInstance.CurrentProject.AllForms
            Case "Module"
                ObjCollectionAllObjects = AccessInstance.CurrentProject.AllModules
                Me.InsertIntoCodeCheckbox.Checked = False
                Me.InsertIntoCodeCheckbox.Enabled = False
                Me.PlaceinClipboardCheckbox.Enabled = False
                Me.OpenObjectCheckbox.Checked = True
            Case "Variable"
                blnArrayFilled = True
                'ObjCollectionAllObjects = AccessInstance.CurrentProject.AllForms
            Case "Procedure"
                Me.InsertIntoCodeCheckbox.Checked = False
                Me.InsertIntoCodeCheckbox.Enabled = True
                Me.PlaceinClipboardCheckbox.Enabled = True
                blnArrayFilled = True
            Case "AllProcedure"
                Me.InsertIntoCodeCheckbox.Checked = False
                Me.InsertIntoCodeCheckbox.Enabled = True
                Me.PlaceinClipboardCheckbox.Enabled = True
                Me.OpenObjectCheckbox.Enabled = True
                Me.OpenObjectCheckbox.Checked = True
                blnArrayFilled = True
        End Select
        If Not blnArrayFilled Then
            i = 0
            If Not ObjCollectionFields Is Nothing Then
                For Each field In ObjCollectionFields
                    i = i + 1
                    varArray(i) = CStr(field.Name)
                Next
            Else
                If Not ObjCollectionAllObjects Is Nothing Then
                    For Each aob In ObjCollectionAllObjects
                        i = i + 1
                        varArray(i) = CStr(aob.Name)
                    Next aob
                End If
            End If

        End If
        Array.Sort(varArray)
        'BubbleSort1(varArray)
        For i = 0 To UBound(varArray)
            If Not IsNothing(varArray(i)) Then
                Me.ObjectsListbox.Items.Add(varArray(i))
            End If
        Next
        UpdateProcedureActionVisibility()
        Me.Text = "Search for " & strObjectType
        Me.SearchTextBox.Focus()


        'db.Close()

        DestroyComObject(DataBase)
        DestroyComObject(aob)
        DestroyComObject(Table)
        DestroyComObject(query)
        DestroyComObject(ObjCollectionAllObjects)
        DestroyComObject(ObjCollectionFields)
    End Sub
    Private Sub DestroyComObject(MyObject As Object)
        Dim IntegerReferenceCount As Integer
        If MyObject Is Nothing Then Exit Sub
        Do
            IntegerReferenceCount = _
             System.Runtime.InteropServices.Marshal.ReleaseComObject(MyObject)
        Loop While IntegerReferenceCount > 0
        MyObject = Nothing

    End Sub
    Private Sub SearchTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchTextBox.KeyDown
        Dim KeyCode As Short = CType(e.KeyCode, Short)
        Dim Shift As Short = CType(e.KeyData \ &H10000, Short)
        On Error GoTo ErrorHandler
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        If KeyCode = System.Windows.Forms.Keys.Down Then
            Me.ObjectsListbox.Focus()
            Me.ObjectsListbox.SetSelected(0, True)
        End If
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.OkayButton.Enabled = True
            Me.SelectTopButton.Enabled = True

        Else
            Me.OkayButton.Enabled = False
        End If
ExitHere:
        Exit Sub
ErrorHandler:
        Select Case Err.Number
            Case 381 'Invalid property array index
                MessageBox.Show("A match has not been found in this case please type something else.", "No Match", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
                Me.SearchTextBox.Focus()
                Resume ExitHere
            Case Else
                MessageBox.Show("The following error has occurred and the current procedure will now abort." & " " & Err.Description, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)

                Resume ExitHere
        End Select
        Resume

    End Sub

    Private Sub SearchTextBox_TextChanged(sender As Object, e As EventArgs) Handles SearchTextBox.TextChanged
        Dim i As Integer
        Me.ObjectsListbox.Items.Clear()
        For i = 0 To UBound(varArray)
            If Not IsNothing(varArray(i)) Then
                'Me.cboControls.AddItem varArrayCtls(i)
                If InStr(UCase(CStr(varArray(i))), UCase(Me.SearchTextBox.Text)) > 0 Then
                    Me.ObjectsListbox.Items.Add(varArray(i))
                End If
            End If
        Next

    End Sub
    Public Sub SetArray(ByVal varArrayIn() As String)
        varArray = varArrayIn
    End Sub


    Private Sub SelectTopButton_Click(sender As Object, e As EventArgs) Handles SelectTopButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(0, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub
    Public Property StringDatabaseFilename1() As String
        Get
            Return StringDatabaseFilename
        End Get
        Set(ByVal value As String)
            StringDatabaseFilename = value
        End Set
    End Property


    Private Sub ShowFieldsCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles ShowFieldsCheckbox.CheckedChanged
        If Me.ShowFieldsCheckbox.Checked Then
            Me.InsertIntoCodeCheckbox.Checked = False
        End If
    End Sub

    Private Sub SelectProcedureButton_Click(sender As Object, e As EventArgs) Handles SelectProcedureButton.Click
        If Me.ObjectsListbox.SelectedIndex < 0 Then Exit Sub

        Dim procedureName As String = Trim(Me.ObjectsListbox.Text)
        If Len(procedureName) = 0 Then Exit Sub
        blnSelectProcedureRequested = True
        Me.Close()
    End Sub

    Private Function ResolveSelectedProcedure(ByRef moduleName As String, ByRef procedureName As String) As Boolean
        Dim selectedText As String = Trim(Me.ObjectsListbox.Text)
        moduleName = ""
        procedureName = ""

        If Len(selectedText) = 0 Then Return False

        If strObjectType = "AllProcedure" Then
            Dim dotPos As Integer = InStr(selectedText, ".")
            If dotPos <= 1 Then Return False
            moduleName = Microsoft.VisualBasic.Left(selectedText, dotPos - 1)
            procedureName = Mid(selectedText, dotPos + 1)
        Else
            procedureName = selectedText
            moduleName = AccessInstance.VBE.ActiveCodePane.CodeModule.Name
        End If

        Return Len(moduleName) > 0 AndAlso Len(procedureName) > 0
    End Function

    Private Function GetProcedureEndLine(ByVal CM As Object, ByVal procStartLine As Integer, ByVal procCount As Integer) As Integer
        Dim i As Integer
        For i = procStartLine + procCount - 1 To procStartLine Step -1
            Dim lineText As String = UCase(Trim(CStr(CM.Lines(i, 1))))
            If lineText = "END SUB" OrElse lineText = "END FUNCTION" Then
                Return i
            End If
        Next

        Return procStartLine + procCount - 1
    End Function

    Private Function GetProcedureDeclarationEndLine(ByVal CM As Object, ByVal procStartLine As Integer, ByVal procCount As Integer) As Integer
        Dim i As Integer
        Dim searchEnd As Integer = procStartLine + procCount - 1
        If searchEnd > procStartLine + 20 Then
            searchEnd = procStartLine + 20
        End If

        Dim declStart As Integer = procStartLine
        For i = procStartLine To searchEnd
            Dim lineText As String = UCase(Trim(CStr(CM.Lines(i, 1))))
            If lineText Like "* SUB *" OrElse lineText Like "* FUNCTION *" Then
                declStart = i
                Exit For
            End If
        Next

        Dim declEnd As Integer = declStart
        Do While declEnd < searchEnd
            Dim currentLine As String = Trim(CStr(CM.Lines(declEnd, 1)))
            If Not currentLine.EndsWith("_") Then Exit Do
            declEnd = declEnd + 1
        Loop

        Return declEnd
    End Function

    Private Sub AddErrorHandlerButton_Click(sender As Object, e As EventArgs) Handles AddErrorHandlerButton.Click
        If Me.ObjectsListbox.SelectedIndex < 0 Then
            MessageBox.Show("Select a procedure before adding an error handler.", "Procedure Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim moduleName As String = ""
        Dim procedureName As String = ""
        If Not ResolveSelectedProcedure(moduleName, procedureName) Then
            MessageBox.Show("Select a procedure before adding an error handler.", "Procedure Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Try
            If AccessInstance.VBE.ActiveCodePane.CodeModule.Name <> moduleName Then
                AccessInstance.DoCmd.OpenModule(moduleName)
            End If

            Dim CM As Object = AccessInstance.VBE.ActiveCodePane.CodeModule
            Dim procStartLine As Integer = CM.ProcStartLine(procedureName, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)
            Dim procCount As Integer = CM.ProcCountLines(procedureName, Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc)

            If procStartLine <= 0 Or procCount <= 0 Then
                MessageBox.Show("The selected procedure could not be found in the active module.", "Procedure Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim procEndLine As Integer = GetProcedureEndLine(CM, procStartLine, procCount)
            Dim hasHandler As Boolean = False
            Dim i As Integer

            For i = procStartLine To procEndLine
                Dim lineText As String = Trim(CStr(CM.Lines(i, 1)))
                If (lineText.ToUpper()) = "ON ERROR GOTO HANDLEERROR" OrElse lineText.ToUpper() = "HANDLEERROR:" Then
                    hasHandler = True
                    Exit For
                End If
            Next

            If hasHandler Then
                ' MessageBox.Show("This procedure already contains an error handler.", "Error Handler Exists", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim declEndLine As Integer = GetProcedureDeclarationEndLine(CM, procStartLine, procCount)
            Dim firstLine As String = Trim(CStr(CM.Lines(declEndLine, 1)))
            Dim isFunction As Boolean = InStr(1, UCase(firstLine), "FUNCTION", vbTextCompare) > 0
            Dim exitStatement As String = IIf(isFunction, "Exit Function", "Exit Sub")
            Dim literalProcedureName As String = """" & procedureName & """"

            CM.InsertLines(declEndLine + 1, "    On Error GoTo HandleError")

            Dim insertText As String = _
                "ExitHere:" & vbCrLf & _
                "    " & exitStatement & vbCrLf & vbCrLf & _
                "HandleError:" & vbCrLf & _
                "    Select Case Err.Number" & vbCrLf & _
                "    ' Case 9" & vbCrLf & _
                "       'Do Something" & vbCrLf & _
                "       'Resume ExitHere" & vbCrLf & _
                "    Case Else" & vbCrLf & _
                "        MsgBox ""Unexpected Error:"" _" & vbCrLf & _
                "            & vbCrLf & ""Error "" & Err.Number & "" "" & Err.Description _" & vbCrLf & _
                "            & "" in procedure "" & " & literalProcedureName & " _" & vbCrLf & _
                "            , vbCritical, ""Please Investigate""" & vbCrLf & _
                "        Resume ExitHere" & vbCrLf & _
                "        Resume 'For Debug Only" & vbCrLf & _
                "    End Select"

            CM.InsertLines(procEndLine + 1, vbCrLf & insertText)

            Me.Visible = False
        Catch ex As Exception
            MessageBox.Show("Unable to insert the error handler: " & ex.Message, "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ObjectsListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ObjectsListbox.SelectedIndexChanged
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.SelectTopButton.Enabled = True
            Me.OkayButton.Enabled = True
        Else
            Me.OkayButton.Enabled = False
        End If
        UpdateProcedureActionVisibility()
    End Sub

    Private Sub SelectSecondButton_Click(sender As Object, e As EventArgs) Handles SelectSecondButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(1, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select3rdButton_Click(sender As Object, e As EventArgs) Handles Select3rdButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(2, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select4thButton_Click(sender As Object, e As EventArgs) Handles Select4thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(3, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select5thButton_Click(sender As Object, e As EventArgs) Handles Select5thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(4, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select6thButton_Click(sender As Object, e As EventArgs) Handles Select6thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(5, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select7thButton_Click(sender As Object, e As EventArgs) Handles Select7thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(6, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select8thButton_Click(sender As Object, e As EventArgs) Handles Select8thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(7, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub

    Private Sub Select9thButton_Click(sender As Object, e As EventArgs) Handles Select9thButton.Click
        Me.ObjectsListbox.Focus()
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        Me.ObjectsListbox.SetSelected(8, True)
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.Visible = False
        End If

    End Sub
End Class