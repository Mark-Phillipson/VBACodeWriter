Imports System.Text.RegularExpressions

Public Class FormParseSQL
    Dim intNoOfLines As Integer
    Dim intLongestLine As Integer
    Private AccessInstance As Microsoft.Office.Interop.Access.Application
    Private Function FormatSQLForEmbeddedCode(ByVal BooleanDeclareVariable As Boolean) As String


        Const intKEYWORDMAX As Short = 12
        Const strLine As String = "    stringSQLText = stringSQLText & "
        Const intLINEUP As Short = 12
        Const strCRTEXT As String = " & vbCrLf"

        Dim strSQL As String
        Dim lngChar As Integer
        Dim blnkeyWord As Boolean
        Dim intKeyWord As Short
        Dim strQ As String
        Dim lngEnd As Integer
        Dim strOut As String

        Dim strKeyWord(intKEYWORDMAX) As String

        ' Should do this by grabbing one word "Element" at a time delimited by
        'spaces or commas.  Then determine what the word is and break
        ' to the start of the line on keywords and long elements.
        ' What is the new continuation character?  make the form allow both
        ' continuations and/or & concatination
        'Dim DataBase As Microsoft.Office.Interop.Access.Dao.Database
        'DataBase = AccessInstance.Application.CurrentDb()
        'Dim objCollection As System.Collections.IEnumerable = AccessInstance.Reports
        'Dim objCollection As System.Collections.IEnumerable = AccessInstance.CurrentProject.AllReports
        'DestroyComObject(objCollection)
        'DestroyComObject(DataBase)
        strKeyWord(1) = "INNER JOIN"
        strKeyWord(2) = "LEFT JOIN"
        strKeyWord(3) = "RIGHT JOIN"
        strKeyWord(4) = "WHERE"
        strKeyWord(5) = "GROUP BY"
        strKeyWord(6) = "ORDER BY"
        strKeyWord(7) = "HAVING"
        strKeyWord(8) = "ON"
        strKeyWord(9) = "FROM"
        strKeyWord(10) = ","
        strKeyWord(11) = "AND"
        strKeyWord(12) = "OR"

        strQ = Chr(34)

        If BooleanDeclareVariable Then
            strOut = "    Dim stringSQLText As String" & vbCrLf
        Else
            strOut = ""
        End If
        strOut = strOut & "    stringSQLText = " & strQ
        If Not Me.TextBoxUnformattedSQL.Text.Length = 0 Then
            Dim stringTemporary As String
            stringTemporary = Me.TextBoxUnformattedSQL.Text
            strSQL = Regex.Replace(stringTemporary, " {2,}", " ")
            lngChar = 1
            intNoOfLines = 0
            Do Until lngChar > Len(strSQL)
                blnkeyWord = False
                For intKeyWord = 1 To intKEYWORDMAX
                    If Mid(strSQL, lngChar, Len(strKeyWord(intKeyWord)) + 1) = (strKeyWord(intKeyWord) & " ") Then
                        blnkeyWord = True
                        If intKeyWord = 8 Then
                            blnkeyWord = Asc(Mid(strSQL, lngChar - 1, 1)) = 32
                        End If
                        If intKeyWord = 11 Then
                            blnkeyWord = Asc(Mid(strSQL, lngChar - 1, 1)) = 32
                        End If
                        If intKeyWord = 12 Then
                            blnkeyWord = Asc(Mid(strSQL, lngChar - 1, 1)) = 32
                        End If
                        Exit For
                    End If
                Next intKeyWord
                If blnkeyWord Then
                    strOut = strOut & strQ & strCRTEXT & vbCrLf & strLine & strQ & Space(intLINEUP - Len(strKeyWord(intKeyWord))) & strKeyWord(intKeyWord)
                    lngChar = lngChar + Len(strKeyWord(intKeyWord))
                    intNoOfLines = intNoOfLines + 1
                    If Len(strLine & strQ & Space(intLINEUP - Len(strKeyWord(intKeyWord))) & strKeyWord(intKeyWord)) > intLongestLine Then
                        intLongestLine = Len(strLine & strQ & Space(intLINEUP - Len(strKeyWord(intKeyWord))) & strKeyWord(intKeyWord))
                    End If
                ElseIf Asc(Mid(strSQL, lngChar, 1)) = 13 Or Asc(Mid(strSQL, lngChar, 1)) = 10 Then
                    lngChar = lngChar + 1
                Else
                    Select Case Asc(Mid(strSQL, lngChar, 1))
                        Case 39
                            lngEnd = InStr(lngChar + 1, strSQL, Mid(strSQL, lngChar, 1))
                            strOut = strOut & Mid(strSQL, lngChar, lngEnd - lngChar + 1)
                            lngChar = lngEnd + 1
                        Case 34
                            lngEnd = InStr(lngChar + 1, strSQL, Mid(strSQL, lngChar, 1))
                            strOut = strOut & strQ & Mid(strSQL, lngChar, lngEnd - lngChar + 1) & strQ
                            lngChar = lngEnd + 1
                        Case 91
                            lngEnd = InStr(lngChar + 1, strSQL, "]")
                            strOut = strOut & Mid(strSQL, lngChar, lngEnd - lngChar + 1)
                            lngChar = lngEnd + 1
                        Case Else
                            lngEnd = lngChar
                            'Debug.Print strOut
                            strOut = strOut & Mid(strSQL, lngChar, lngEnd - lngChar + 1)
                            'Debug.Assert InStr(strOut, "strSQl") > 0

                            lngChar = lngEnd + 1
                    End Select
                End If
            Loop
        End If

        Return strOut & strQ

    End Function

    Private Sub ButtonGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGenerate.Click
        Me.TextBoxFormattedSQL.Text = FormatSQLForEmbeddedCode(Me.CheckBoxDeclareVariable.Checked)
    End Sub

    Private Sub FormParseSQL_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.TextBoxUnformattedSQL.Text = My.Computer.Clipboard.GetText
    End Sub

    Private Sub ButtonInsertCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonInsertCode.Click
        Dim CM As Microsoft.Vbe.Interop.CodeModule
        Dim startLine As Integer
        Dim startColumn As Integer
        Dim EndLine As Integer
        Dim EndColumn As Integer
        Me.TextBoxFormattedSQL.Text = FormatSQLForEmbeddedCode(Me.CheckBoxDeclareVariable.Checked)
        'Get active codemodule
        CM = AccessInstance.VBE.ActiveCodePane.CodeModule
        AccessInstance.VBE.ActiveCodePane.GetSelection(startLine, startColumn, EndLine, EndColumn)
        CM.InsertLines(startLine, Me.TextBoxFormattedSQL.Text)
        If Me.CheckBoxDeclareVariable.Checked Then
            intNoOfLines = intNoOfLines + 1
        End If
        AccessInstance.VBE.ActiveCodePane.SetSelection(startLine, 1, startLine + intNoOfLines, 1000) 'CLng(intLongestLine)
        Me.Visible = False

    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub ButtonReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReset.Click
        Me.TextBoxUnformattedSQL.Text = ""
        Me.TextBoxUnformattedSQL.Focus()
    End Sub
    Public Property AccessInstance1() As Microsoft.Office.Interop.Access.Application
        Get
            Return AccessInstance
        End Get
        Set(ByVal value As Microsoft.Office.Interop.Access.Application)
            AccessInstance = value
        End Set
    End Property

    Private Sub DestroyComObject(MyObject As Object)
        Dim IntegerReferenceCount As Integer
        If MyObject Is Nothing Then Exit Sub
        Do
            IntegerReferenceCount =
             System.Runtime.InteropServices.Marshal.ReleaseComObject(MyObject)
        Loop While IntegerReferenceCount > 0
        MyObject = Nothing

    End Sub
End Class