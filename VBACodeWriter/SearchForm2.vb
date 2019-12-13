Public Class SearchForm2
    Private strObjectName As String
    Private stringFormOrReport As String
    Private strObjectType As String
    Private varArray(0 To 3000) As String
    Private AccessInstance As Microsoft.Office.Interop.Access.Application
    Private StringDatabaseFilename As String
    Private ObjectSetting As New ObjectSettings

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

    Private Sub SearchForm2_Load(sender As Object, e As EventArgs) Handles Me.Load
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

    Private Sub ObjectsListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ObjectsListbox.SelectedIndexChanged
        If Me.ObjectsListbox.Items.Count = 0 Then Exit Sub
        If Len(Me.ObjectsListbox.Text) > 0 Then
            Me.SelectTopButton.Enabled = True
            Me.OkayButton.Enabled = True
        Else
            Me.OkayButton.Enabled = False
        End If
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