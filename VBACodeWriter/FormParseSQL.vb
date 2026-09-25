Imports System.Text.RegularExpressions

Public Class FormParseSQL
    Dim intNoOfLines As Integer
    Dim intLongestLine As Integer
    Private AccessInstance As Microsoft.Office.Interop.Access.Application
    Private Function FormatSQLForEmbeddedCode(ByVal BooleanDeclareVariable As Boolean) As String


        Const intKEYWORDMAX As Short = 47
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
        Dim strPostKeyword As String

        Dim strKeyWord() As String = {
            "LEFT OUTER JOIN",
            "RIGHT OUTER JOIN",
            "FULL OUTER JOIN",
            "INNER JOIN",
            "LEFT JOIN",
            "RIGHT JOIN",
            "FULL JOIN",
            "CROSS JOIN",
            "OUTER APPLY",
            "CROSS APPLY",
            "SELECT",
            "INSERT",
            "UPDATE",
            "DELETE",
            "MERGE",
            "FROM",
            "WHERE",
            "GROUP BY",
            "ORDER BY",
            "HAVING",
            "ON",
            "JOIN",
            "UNION",
            "UNION ALL",
            "INTERSECT",
            "EXCEPT",
            "CASE",
            "WHEN",
            "THEN",
            "ELSE",
            "END",
            "AS",
            "DISTINCT",
            "TOP",
            "INTO",
            "VALUES",
            "SET",
            "IN",
            "LIKE",
            "BETWEEN",
            "IS",
            "NOT",
            "NULL",
            "EXISTS",
            "ALL",
            "ANY",
            "SOME",
            "WITH",
            "OVER",
            "PARTITION BY",
            "WINDOW",
            "RETURNING",
            "DECLARE",
            "BEGIN",
            "END",
            "IF",
            "ELSEIF",
            "WHILE",
            "FOR",
            "NEXT",
            "TRUNCATE",
            "BEGIN TRAN",
            "COMMIT",
            "ROLLBACK",
            "ALTER",
            "CREATE",
            "DROP",
            "RENAME",
            "EXEC",
            "EXECUTE",
            "GO",
            "PIVOT",
            "UNPIVOT",
            "MATCH",
            "BY",
            "THROW",
            "TRY",
            "CATCH",
            "AND",
            "OR",
            ","
        }

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
                For intKeyWord = 0 To strKeyWord.GetUpperBound(0)
                    Dim strKW As String = strKeyWord(intKeyWord)
                    If Len(strKW) = 0 Then Continue For

                    Dim strCandidate As String = UCase$(Mid(strSQL, lngChar, Len(strKW)))
                    Dim strKeywordUpper As String = UCase$(strKW)
                    Dim strNextChar As String = ""
                    Dim strPreviousChar As String = ""

                    If lngChar + Len(strKW) <= Len(strSQL) Then
                        strNextChar = Mid(strSQL, lngChar + Len(strKW), 1)
                    End If
                    If lngChar > 1 Then
                        strPreviousChar = Mid(strSQL, lngChar - 1, 1)
                    End If

                    Dim blnPreviousBoundary As Boolean =
                        lngChar = 1 OrElse
                        strPreviousChar = " " OrElse
                        strPreviousChar = vbTab OrElse
                        strPreviousChar = vbCr OrElse
                        strPreviousChar = vbLf OrElse
                        strPreviousChar = "," OrElse
                        strPreviousChar = ";" OrElse
                        strPreviousChar = "(" OrElse
                        strPreviousChar = ")"

                    Dim blnNextBoundary As Boolean =
                        lngChar + Len(strKW) > Len(strSQL) OrElse
                        strNextChar = "" OrElse
                        strNextChar = " " OrElse
                        strNextChar = vbTab OrElse
                        strNextChar = vbCr OrElse
                        strNextChar = vbLf OrElse
                        strNextChar = "," OrElse
                        strNextChar = ";" OrElse
                        strNextChar = "(" OrElse
                        strNextChar = ")"

                    If strCandidate = strKeywordUpper AndAlso blnPreviousBoundary AndAlso blnNextBoundary Then
                        blnkeyWord = True
                        Exit For
                    End If
                Next intKeyWord
                If blnkeyWord Then
                    Dim strKeywordOut As String = strKeyWord(intKeyWord)
                    Dim intPadding As Integer = Math.Max(1, intLINEUP - Len(strKeywordOut))
                    strOut = strOut & strQ & strCRTEXT & vbCrLf & strLine & strQ & Space(intPadding) & strKeywordOut
                    lngChar = lngChar + Len(strKeywordOut)
                    intNoOfLines = intNoOfLines + 1
                    If Len(strLine & strQ & Space(intPadding) & strKeywordOut) > intLongestLine Then
                        intLongestLine = Len(strLine & strQ & Space(intPadding) & strKeywordOut)
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