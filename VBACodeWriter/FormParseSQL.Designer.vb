<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormParseSQL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBoxDeclareVariable = New System.Windows.Forms.CheckBox()
        Me.TextBoxUnformattedSQL = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxFormattedSQL = New System.Windows.Forms.TextBox()
        Me.ButtonGenerate = New System.Windows.Forms.Button()
        Me.ButtonReset = New System.Windows.Forms.Button()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.ButtonInsertCode = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(20, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(268, 70)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "INSTRUCTIONS" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1) Paste your SQL into the left text box." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2) Click Generate to c" &
    "reate the formatted SQL." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3) Click the Insert Code button."
        '
        'CheckBoxDeclareVariable
        '
        Me.CheckBoxDeclareVariable.AccessibleDescription = "clicked hundred and 31"
        Me.CheckBoxDeclareVariable.AutoSize = True
        Me.CheckBoxDeclareVariable.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CheckBoxDeclareVariable.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxDeclareVariable.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.CheckBoxDeclareVariable.Location = New System.Drawing.Point(20, 153)
        Me.CheckBoxDeclareVariable.Name = "CheckBoxDeclareVariable"
        Me.CheckBoxDeclareVariable.Size = New System.Drawing.Size(112, 18)
        Me.CheckBoxDeclareVariable.TabIndex = 1
        Me.CheckBoxDeclareVariable.Text = "Declare Variable"
        Me.CheckBoxDeclareVariable.UseVisualStyleBackColor = False
        '
        'TextBoxUnformattedSQL
        '
        Me.TextBoxUnformattedSQL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBoxUnformattedSQL.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TextBoxUnformattedSQL.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxUnformattedSQL.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBoxUnformattedSQL.Location = New System.Drawing.Point(20, 213)
        Me.TextBoxUnformattedSQL.Multiline = True
        Me.TextBoxUnformattedSQL.Name = "TextBoxUnformattedSQL"
        Me.TextBoxUnformattedSQL.Size = New System.Drawing.Size(338, 317)
        Me.TextBoxUnformattedSQL.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label2.Location = New System.Drawing.Point(21, 186)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(137, 14)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Paste Unformatted SQL"
        '
        'TextBoxFormattedSQL
        '
        Me.TextBoxFormattedSQL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxFormattedSQL.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TextBoxFormattedSQL.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxFormattedSQL.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBoxFormattedSQL.Location = New System.Drawing.Point(370, 212)
        Me.TextBoxFormattedSQL.Multiline = True
        Me.TextBoxFormattedSQL.Name = "TextBoxFormattedSQL"
        Me.TextBoxFormattedSQL.Size = New System.Drawing.Size(338, 317)
        Me.TextBoxFormattedSQL.TabIndex = 4
        '
        'ButtonGenerate
        '
        Me.ButtonGenerate.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ButtonGenerate.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonGenerate.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonGenerate.Location = New System.Drawing.Point(468, 34)
        Me.ButtonGenerate.Name = "ButtonGenerate"
        Me.ButtonGenerate.Size = New System.Drawing.Size(191, 37)
        Me.ButtonGenerate.TabIndex = 5
        Me.ButtonGenerate.Text = "Generate"
        Me.ButtonGenerate.UseVisualStyleBackColor = False
        '
        'ButtonReset
        '
        Me.ButtonReset.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ButtonReset.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonReset.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonReset.Location = New System.Drawing.Point(468, 71)
        Me.ButtonReset.Name = "ButtonReset"
        Me.ButtonReset.Size = New System.Drawing.Size(191, 37)
        Me.ButtonReset.TabIndex = 6
        Me.ButtonReset.Text = "Reset"
        Me.ButtonReset.UseVisualStyleBackColor = False
        '
        'ButtonClose
        '
        Me.ButtonClose.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ButtonClose.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonClose.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonClose.Location = New System.Drawing.Point(468, 108)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(191, 37)
        Me.ButtonClose.TabIndex = 7
        Me.ButtonClose.Text = "Close"
        Me.ButtonClose.UseVisualStyleBackColor = False
        '
        'ButtonInsertCode
        '
        Me.ButtonInsertCode.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ButtonInsertCode.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonInsertCode.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonInsertCode.Location = New System.Drawing.Point(468, 145)
        Me.ButtonInsertCode.Name = "ButtonInsertCode"
        Me.ButtonInsertCode.Size = New System.Drawing.Size(191, 37)
        Me.ButtonInsertCode.TabIndex = 8
        Me.ButtonInsertCode.Text = "Insert Code"
        Me.ButtonInsertCode.UseVisualStyleBackColor = False
        '
        'FormParseSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(723, 547)
        Me.Controls.Add(Me.ButtonInsertCode)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.ButtonReset)
        Me.Controls.Add(Me.ButtonGenerate)
        Me.Controls.Add(Me.TextBoxFormattedSQL)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxUnformattedSQL)
        Me.Controls.Add(Me.CheckBoxDeclareVariable)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Name = "FormParseSQL"
        Me.Text = "Parse SQL"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxDeclareVariable As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxUnformattedSQL As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBoxFormattedSQL As System.Windows.Forms.TextBox
    Friend WithEvents ButtonGenerate As System.Windows.Forms.Button
    Friend WithEvents ButtonReset As System.Windows.Forms.Button
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents ButtonInsertCode As System.Windows.Forms.Button
End Class
