Option Strict Off
Option Explicit On
Module basInsertCode

    Public Sub InsertValueIntoCode(ByRef strValue As String, AccessInst As Microsoft.Office.Interop.Access.Application, ObjectSetting As ObjectSettings)
        'Test
        Dim CM As Object 'CodeModule
        Dim startLine As Integer
        Dim startColumn As Integer
        Dim EndLine As Integer
        Dim EndColumn As Integer
        Dim strSelection As String
        Dim blnSelected As Boolean
        Dim strLine As String
        Dim strLeftOfCursor As String
        Dim strRightOfCursor As String
        CM = AccessInst.VBE.ActiveCodePane.CodeModule
        AccessInst.VBE.ActiveCodePane.GetSelection(startLine, startColumn, EndLine, EndColumn)
        'Debug.Print(VB6.TabLayout(startLine, startColumn, EndLine, EndColumn))
        If startColumn = EndColumn Then
            blnSelected = False
        Else
            blnSelected = True
        End If
        strSelection = AccessInst.VBE.ActiveCodePane.CodeModule.Lines(startLine, EndLine + 1 - startLine)
        strLine = AccessInst.VBE.ActiveCodePane.CodeModule.Lines(startLine, EndLine + 1 - startLine)
        strLeftOfCursor = Left(strLine, EndColumn)
        strRightOfCursor = Mid(strLine, EndColumn + 1)
        'strSelection = Trim(Mid(strSelection, startColumn, EndColumn - startColumn))
        If Not blnSelected Then
            'strSelection = RTrim(strLeftOfCursor & " " & strValue & " " & strRightOfCursor)
            strSelection = Left(strLine, startColumn - 1) & strValue & Mid(strLine, EndColumn)
        Else
            'strSelection = Replace(strLine, Mid(strLine, startColumn, EndColumn - startColumn), strValue, 1, 1)
            strSelection = Left(strLine, startColumn - 1) & strValue & Mid(strLine, EndColumn)

        End If

        If Len(strValue) > 0 Then 'Enter the value on the line
            'CM.InsertLines startLine, cboVariables.Text
            'UPGRADE_WARNING: Couldn't resolve default property of object CM.ReplaceLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Try
                CM.ReplaceLine(startLine, strSelection)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try



        End If
        'Set Selection
        If blnSelected Then
            strLine = AccessInst.VBE.ActiveCodePane.CodeModule.Lines(startLine, EndLine + 1 - startLine)
            startColumn = InStr(strLine, strValue)
            AccessInst.VBE.ActiveCodePane.SetSelection(startLine, startColumn, EndLine, startColumn + Len(strValue))
        Else
            AccessInst.VBE.ActiveCodePane.SetSelection(startLine, startColumn + Len(strSelection), EndLine, EndColumn + Len(strSelection))
        End If
    End Sub
    Public Sub BubbleSort1(ByRef pvarArray() As Object)
        Dim i As Integer
        Dim iMin As Integer
        Dim iMax As Integer
        Dim varSwap As Object
        Dim blnSwapped As Boolean

        iMin = LBound(pvarArray)
        iMax = UBound(pvarArray) - 1
        Do
            blnSwapped = False
            For i = iMin To iMax
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If pvarArray(i) > pvarArray(i + 1) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object varSwap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    varSwap = pvarArray(i)
                    'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pvarArray(i) = pvarArray(i + 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object varSwap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pvarArray(i + 1) = varSwap
                    blnSwapped = True
                End If
            Next
            iMax = iMax - 1
        Loop Until Not blnSwapped
    End Sub

End Module