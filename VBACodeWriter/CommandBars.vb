
'Option Strict Off
Option Explicit On

Module CommandBars


    'Private Const conMenuIDTools = 30007


    Function AddMenuCommand(ByRef cbrMenu As Microsoft.Office.Core.CommandBar,
                            ByRef AddInInst As Object, Optional ByRef strProp As String = "", Optional _
                            ByRef strValue As String = "", Optional ByRef lngFaceid As Integer = 0, _
                            Optional ByRef conMenuName As String = "&VBA Code Writer") As Microsoft.Office.Core.CommandBarControl


        Dim cbcMsgBox As Microsoft.Office.Core.CommandBarControl
        Dim cbcFormat As Microsoft.Office.Core.CommandBarControl
        Dim c As Microsoft.Office.Core.CommandBarControl
        Dim cbr As Microsoft.Office.Core.CommandBar
        Dim cbc As Microsoft.Office.Core.CommandBarControl
        Dim blnValue As Boolean
        On Error GoTo HandleErr
        AddMenuCommand = Nothing
        If cbrMenu.Name = "Tools" Then
            ' Get a pointer to the Tools menu
            'Set cbcTools = cbrMenu.FindControl( _
            ''Name:="Code Window", Recursive:=False)

            ' If we found the Tools menu then add
            ' a new menu command (but only if it doesn't
            ' already exist!)
            If Not cbrMenu Is Nothing Then

                ' Try to find the command based on its tag
                cbcMsgBox = cbrMenu.FindControl(Tag:=conMenuName, Recursive:=False)

                ' If we didn't find it, add a new command
                If cbcMsgBox Is Nothing Then
                    If conMenuName = "Comment Block" Then
                        cbcMsgBox = cbrMenu.Controls.Add(Type:=Microsoft.Office.Core.MsoControlType.msoControlButton, Id:=192)
                    ElseIf conMenuName = "UnComment Block" Then
                        cbcMsgBox = cbrMenu.Controls.Add(Type:=Microsoft.Office.Core.MsoControlType.msoControlButton, Id:=2552)
                    Else
                        cbcMsgBox = cbrMenu.Controls.Add(Type:=Microsoft.Office.Core.MsoControlType.msoControlButton)
                    End If

                    With cbcMsgBox
                            .Caption = conMenuName


                        .Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonCaption
                        .Tag = conMenuName
                        .FaceId = lngFaceid

                        ' This enables demand loading
                        '.OnAction = "!<" & AddInInst.ProgId & ">"
                        .BeginGroup = False
                        .TooltipText = "Click here to get " & strValue
                        .DescriptionText = strValue
                        .Visible = True
                    End With
                End If

                ' Return pointer to menu command
                AddMenuCommand = cbcMsgBox
                'Debug.Print cbcMsgBox.Caption
            End If
        Else
            cbr = cbrMenu
            cbc = cbr.FindControl(Tag:=strValue)
            If strValue = "True" Then
                blnValue = True
            Else
                blnValue = False
            End If
            If cbc Is Nothing Then
                If conMenuName = "Comment Block" Then
                    cbc = cbr.Controls.Add(Id:=192)
                ElseIf conMenuName = "UnComment Block" Then
                    cbc = cbr.Controls.Add(Id:=2552)
                Else
                    cbc = cbr.Controls.Add(Type:=Microsoft.Office.Core.MsoControlType.msoControlButton)

                    With cbc
                        .Caption = strValue
                        .Tag = .Caption
                        '.OnAction = "=MercVBACodeWriter.ToggleProperty(" & Chr(34) & strProp & Chr(34) & ", True)"
                        .BeginGroup = False
                        .TooltipText = .Caption
                        .DescriptionText = .Caption
                        .Visible = True
                        'UPGRADE_WARNING: Couldn't resolve default property of object cbc.FaceId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .FaceId = lngFaceid
                        'UPGRADE_WARNING: Couldn't resolve default property of object cbc.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonCaption

                    End With
                End If
            End If
            cbcFormat = cbr.FindControl(Tag:=strValue)
            AddMenuCommand = cbcFormat

        End If



ExitHere:
        Exit Function

        ' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
        ' Automatic error
HandleErr:
        Select Case Err.Number
            'Case # '
            ' MsgBox "", vbExclamation, ""
            Case Else
                MsgBox("Error " & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Unexpected Error in basCommandBars.AddMenuCommand") 'ErrorHandler:$$N=basCommandBars.AddMenuCommand
                Resume ExitHere
        End Select
        Resume
    End Function


End Module
