﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SearchForm2
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
        Me.SearchTextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ObjectsListbox = New System.Windows.Forms.ListBox()
        Me.SelectTopButton = New System.Windows.Forms.Button()
        Me.OkayButton = New System.Windows.Forms.Button()
        Me.DoCanceButton = New System.Windows.Forms.Button()
        Me.InsertIntoCodeCheckbox = New System.Windows.Forms.CheckBox()
        Me.PlaceinClipboardCheckbox = New System.Windows.Forms.CheckBox()
        Me.TableQueryTextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ShowFieldsCheckbox = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LastObjectTypeTextBox = New System.Windows.Forms.TextBox()
        Me.SelectSecondButton = New System.Windows.Forms.Button()
        Me.OpenObjectCheckbox = New System.Windows.Forms.CheckBox()
        Me.Select3rdButton = New System.Windows.Forms.Button()
        Me.Select4thButton = New System.Windows.Forms.Button()
        Me.Select5thButton = New System.Windows.Forms.Button()
        Me.Select6thButton = New System.Windows.Forms.Button()
        Me.Select7thButton = New System.Windows.Forms.Button()
        Me.Select8thButton = New System.Windows.Forms.Button()
        Me.Select9thButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'SearchTextBox
        '
        Me.SearchTextBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.SearchTextBox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchTextBox.ForeColor = System.Drawing.Color.White
        Me.SearchTextBox.Location = New System.Drawing.Point(141, 39)
        Me.SearchTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SearchTextBox.Name = "SearchTextBox"
        Me.SearchTextBox.Size = New System.Drawing.Size(543, 26)
        Me.SearchTextBox.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(12, 43)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "&Search"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(12, 443)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "&Name"
        '
        'ObjectsListbox
        '
        Me.ObjectsListbox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ObjectsListbox.Font = New System.Drawing.Font("Calibri", 23.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ObjectsListbox.ForeColor = System.Drawing.Color.White
        Me.ObjectsListbox.FormattingEnabled = True
        Me.ObjectsListbox.ItemHeight = 38
        Me.ObjectsListbox.Location = New System.Drawing.Point(140, 90)
        Me.ObjectsListbox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ObjectsListbox.Name = "ObjectsListbox"
        Me.ObjectsListbox.Size = New System.Drawing.Size(545, 346)
        Me.ObjectsListbox.TabIndex = 3
        '
        'SelectTopButton
        '
        Me.SelectTopButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.SelectTopButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.SelectTopButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectTopButton.ForeColor = System.Drawing.Color.White
        Me.SelectTopButton.Location = New System.Drawing.Point(37, 87)
        Me.SelectTopButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SelectTopButton.Name = "SelectTopButton"
        Me.SelectTopButton.Size = New System.Drawing.Size(96, 42)
        Me.SelectTopButton.TabIndex = 4
        Me.SelectTopButton.Text = "&Select Top"
        Me.SelectTopButton.UseVisualStyleBackColor = False
        '
        'OkayButton
        '
        Me.OkayButton.BackColor = System.Drawing.SystemColors.Highlight
        Me.OkayButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.OkayButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OkayButton.ForeColor = System.Drawing.Color.White
        Me.OkayButton.Location = New System.Drawing.Point(711, 36)
        Me.OkayButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OkayButton.Name = "OkayButton"
        Me.OkayButton.Size = New System.Drawing.Size(152, 36)
        Me.OkayButton.TabIndex = 5
        Me.OkayButton.Text = "&Okay"
        Me.OkayButton.UseVisualStyleBackColor = False
        '
        'DoCanceButton
        '
        Me.DoCanceButton.BackColor = System.Drawing.SystemColors.Highlight
        Me.DoCanceButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.DoCanceButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.DoCanceButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DoCanceButton.ForeColor = System.Drawing.Color.White
        Me.DoCanceButton.Location = New System.Drawing.Point(711, 91)
        Me.DoCanceButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DoCanceButton.Name = "DoCanceButton"
        Me.DoCanceButton.Size = New System.Drawing.Size(152, 36)
        Me.DoCanceButton.TabIndex = 6
        Me.DoCanceButton.Text = "&Cancel"
        Me.DoCanceButton.UseVisualStyleBackColor = False
        '
        'InsertIntoCodeCheckbox
        '
        Me.InsertIntoCodeCheckbox.AutoSize = True
        Me.InsertIntoCodeCheckbox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.InsertIntoCodeCheckbox.Checked = True
        Me.InsertIntoCodeCheckbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.InsertIntoCodeCheckbox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InsertIntoCodeCheckbox.ForeColor = System.Drawing.Color.White
        Me.InsertIntoCodeCheckbox.Location = New System.Drawing.Point(711, 151)
        Me.InsertIntoCodeCheckbox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.InsertIntoCodeCheckbox.Name = "InsertIntoCodeCheckbox"
        Me.InsertIntoCodeCheckbox.Size = New System.Drawing.Size(126, 22)
        Me.InsertIntoCodeCheckbox.TabIndex = 7
        Me.InsertIntoCodeCheckbox.Text = "Insert Into Code"
        Me.InsertIntoCodeCheckbox.UseVisualStyleBackColor = False
        '
        'PlaceinClipboardCheckbox
        '
        Me.PlaceinClipboardCheckbox.AutoSize = True
        Me.PlaceinClipboardCheckbox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.PlaceinClipboardCheckbox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PlaceinClipboardCheckbox.ForeColor = System.Drawing.Color.White
        Me.PlaceinClipboardCheckbox.Location = New System.Drawing.Point(711, 198)
        Me.PlaceinClipboardCheckbox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PlaceinClipboardCheckbox.Name = "PlaceinClipboardCheckbox"
        Me.PlaceinClipboardCheckbox.Size = New System.Drawing.Size(138, 22)
        Me.PlaceinClipboardCheckbox.TabIndex = 8
        Me.PlaceinClipboardCheckbox.Text = "Place in Clipboard"
        Me.PlaceinClipboardCheckbox.UseVisualStyleBackColor = False
        '
        'TableQueryTextBox
        '
        Me.TableQueryTextBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.TableQueryTextBox.Enabled = False
        Me.TableQueryTextBox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TableQueryTextBox.ForeColor = System.Drawing.Color.White
        Me.TableQueryTextBox.Location = New System.Drawing.Point(171, 475)
        Me.TableQueryTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TableQueryTextBox.Name = "TableQueryTextBox"
        Me.TableQueryTextBox.Size = New System.Drawing.Size(419, 26)
        Me.TableQueryTextBox.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(12, 479)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 18)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Last Object Name"
        '
        'ShowFieldsCheckbox
        '
        Me.ShowFieldsCheckbox.AutoSize = True
        Me.ShowFieldsCheckbox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ShowFieldsCheckbox.Enabled = False
        Me.ShowFieldsCheckbox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShowFieldsCheckbox.ForeColor = System.Drawing.Color.White
        Me.ShowFieldsCheckbox.Location = New System.Drawing.Point(711, 245)
        Me.ShowFieldsCheckbox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ShowFieldsCheckbox.Name = "ShowFieldsCheckbox"
        Me.ShowFieldsCheckbox.Size = New System.Drawing.Size(141, 22)
        Me.ShowFieldsCheckbox.TabIndex = 11
        Me.ShowFieldsCheckbox.Text = "Show List of Fields"
        Me.ShowFieldsCheckbox.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label4.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(604, 482)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 18)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Last Object Type"
        '
        'LastObjectTypeTextBox
        '
        Me.LastObjectTypeTextBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.LastObjectTypeTextBox.Enabled = False
        Me.LastObjectTypeTextBox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LastObjectTypeTextBox.ForeColor = System.Drawing.Color.White
        Me.LastObjectTypeTextBox.Location = New System.Drawing.Point(763, 476)
        Me.LastObjectTypeTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.LastObjectTypeTextBox.Name = "LastObjectTypeTextBox"
        Me.LastObjectTypeTextBox.Size = New System.Drawing.Size(163, 26)
        Me.LastObjectTypeTextBox.TabIndex = 12
        '
        'SelectSecondButton
        '
        Me.SelectSecondButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.SelectSecondButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.SelectSecondButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectSecondButton.ForeColor = System.Drawing.Color.White
        Me.SelectSecondButton.Location = New System.Drawing.Point(37, 125)
        Me.SelectSecondButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SelectSecondButton.Name = "SelectSecondButton"
        Me.SelectSecondButton.Size = New System.Drawing.Size(96, 42)
        Me.SelectSecondButton.TabIndex = 14
        Me.SelectSecondButton.Text = "Select 2nd"
        Me.SelectSecondButton.UseVisualStyleBackColor = False
        '
        'OpenObjectCheckbox
        '
        Me.OpenObjectCheckbox.AutoSize = True
        Me.OpenObjectCheckbox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.OpenObjectCheckbox.Enabled = False
        Me.OpenObjectCheckbox.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenObjectCheckbox.ForeColor = System.Drawing.Color.White
        Me.OpenObjectCheckbox.Location = New System.Drawing.Point(711, 292)
        Me.OpenObjectCheckbox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OpenObjectCheckbox.Name = "OpenObjectCheckbox"
        Me.OpenObjectCheckbox.Size = New System.Drawing.Size(105, 22)
        Me.OpenObjectCheckbox.TabIndex = 15
        Me.OpenObjectCheckbox.Text = "Open Object"
        Me.OpenObjectCheckbox.UseVisualStyleBackColor = False
        '
        'Select3rdButton
        '
        Me.Select3rdButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select3rdButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select3rdButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select3rdButton.ForeColor = System.Drawing.Color.White
        Me.Select3rdButton.Location = New System.Drawing.Point(37, 162)
        Me.Select3rdButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select3rdButton.Name = "Select3rdButton"
        Me.Select3rdButton.Size = New System.Drawing.Size(96, 42)
        Me.Select3rdButton.TabIndex = 16
        Me.Select3rdButton.Text = "Select 3rd"
        Me.Select3rdButton.UseVisualStyleBackColor = False
        '
        'Select4thButton
        '
        Me.Select4thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select4thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select4thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select4thButton.ForeColor = System.Drawing.Color.White
        Me.Select4thButton.Location = New System.Drawing.Point(37, 199)
        Me.Select4thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select4thButton.Name = "Select4thButton"
        Me.Select4thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select4thButton.TabIndex = 17
        Me.Select4thButton.Text = "Select 4th"
        Me.Select4thButton.UseVisualStyleBackColor = False
        '
        'Select5thButton
        '
        Me.Select5thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select5thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select5thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select5thButton.ForeColor = System.Drawing.Color.White
        Me.Select5thButton.Location = New System.Drawing.Point(37, 237)
        Me.Select5thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select5thButton.Name = "Select5thButton"
        Me.Select5thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select5thButton.TabIndex = 18
        Me.Select5thButton.Text = "Select 5th"
        Me.Select5thButton.UseVisualStyleBackColor = False
        '
        'Select6thButton
        '
        Me.Select6thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select6thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select6thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select6thButton.ForeColor = System.Drawing.Color.White
        Me.Select6thButton.Location = New System.Drawing.Point(37, 276)
        Me.Select6thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select6thButton.Name = "Select6thButton"
        Me.Select6thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select6thButton.TabIndex = 19
        Me.Select6thButton.Text = "Select 6th"
        Me.Select6thButton.UseVisualStyleBackColor = False
        '
        'Select7thButton
        '
        Me.Select7thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select7thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select7thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select7thButton.ForeColor = System.Drawing.Color.White
        Me.Select7thButton.Location = New System.Drawing.Point(37, 313)
        Me.Select7thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select7thButton.Name = "Select7thButton"
        Me.Select7thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select7thButton.TabIndex = 20
        Me.Select7thButton.Text = "Select 7th"
        Me.Select7thButton.UseVisualStyleBackColor = False
        '
        'Select8thButton
        '
        Me.Select8thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select8thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select8thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select8thButton.ForeColor = System.Drawing.Color.White
        Me.Select8thButton.Location = New System.Drawing.Point(37, 352)
        Me.Select8thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select8thButton.Name = "Select8thButton"
        Me.Select8thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select8thButton.TabIndex = 21
        Me.Select8thButton.Text = "Select 8th"
        Me.Select8thButton.UseVisualStyleBackColor = False
        '
        'Select9thButton
        '
        Me.Select9thButton.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.Select9thButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Select9thButton.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Select9thButton.ForeColor = System.Drawing.Color.White
        Me.Select9thButton.Location = New System.Drawing.Point(37, 390)
        Me.Select9thButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Select9thButton.Name = "Select9thButton"
        Me.Select9thButton.Size = New System.Drawing.Size(96, 42)
        Me.Select9thButton.TabIndex = 22
        Me.Select9thButton.Text = "Select 9th"
        Me.Select9thButton.UseVisualStyleBackColor = False
        '
        'SearchForm2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(12, Byte), Integer), CType(CType(12, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(945, 523)
        Me.Controls.Add(Me.Select9thButton)
        Me.Controls.Add(Me.Select8thButton)
        Me.Controls.Add(Me.Select7thButton)
        Me.Controls.Add(Me.Select6thButton)
        Me.Controls.Add(Me.Select5thButton)
        Me.Controls.Add(Me.Select4thButton)
        Me.Controls.Add(Me.Select3rdButton)
        Me.Controls.Add(Me.OpenObjectCheckbox)
        Me.Controls.Add(Me.SelectSecondButton)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LastObjectTypeTextBox)
        Me.Controls.Add(Me.ShowFieldsCheckbox)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TableQueryTextBox)
        Me.Controls.Add(Me.PlaceinClipboardCheckbox)
        Me.Controls.Add(Me.InsertIntoCodeCheckbox)
        Me.Controls.Add(Me.DoCanceButton)
        Me.Controls.Add(Me.OkayButton)
        Me.Controls.Add(Me.SelectTopButton)
        Me.Controls.Add(Me.ObjectsListbox)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SearchTextBox)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "SearchForm2"
        Me.Text = "SearchForm2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SearchTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ObjectsListbox As System.Windows.Forms.ListBox
    Friend WithEvents SelectTopButton As System.Windows.Forms.Button
    Friend WithEvents OkayButton As System.Windows.Forms.Button
    Friend WithEvents DoCanceButton As System.Windows.Forms.Button
    Friend WithEvents InsertIntoCodeCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents PlaceinClipboardCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents TableQueryTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ShowFieldsCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LastObjectTypeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SelectSecondButton As System.Windows.Forms.Button
    Friend WithEvents OpenObjectCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents Select3rdButton As System.Windows.Forms.Button
    Friend WithEvents Select4thButton As System.Windows.Forms.Button
    Friend WithEvents Select5thButton As System.Windows.Forms.Button
    Friend WithEvents Select6thButton As Button
    Friend WithEvents Select7thButton As Button
    Friend WithEvents Select8thButton As Button
    Friend WithEvents Select9thButton As Button
End Class
