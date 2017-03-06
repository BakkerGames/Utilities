<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormExternalCompare
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
        Me.TextBoxExternalCompareApp = New System.Windows.Forms.TextBox()
        Me.LabelUseExternalCompare = New System.Windows.Forms.Label()
        Me.LabelExternalCompareApp = New System.Windows.Forms.Label()
        Me.CheckBoxUseExternalCompare = New System.Windows.Forms.CheckBox()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBoxExternalCompareApp
        '
        Me.TextBoxExternalCompareApp.Location = New System.Drawing.Point(175, 38)
        Me.TextBoxExternalCompareApp.Name = "TextBoxExternalCompareApp"
        Me.TextBoxExternalCompareApp.Size = New System.Drawing.Size(542, 20)
        Me.TextBoxExternalCompareApp.TabIndex = 3
        '
        'LabelUseExternalCompare
        '
        Me.LabelUseExternalCompare.AutoSize = True
        Me.LabelUseExternalCompare.Location = New System.Drawing.Point(12, 15)
        Me.LabelUseExternalCompare.Name = "LabelUseExternalCompare"
        Me.LabelUseExternalCompare.Size = New System.Drawing.Size(152, 13)
        Me.LabelUseExternalCompare.TabIndex = 0
        Me.LabelUseExternalCompare.Text = "Use External compare program"
        '
        'LabelExternalCompareApp
        '
        Me.LabelExternalCompareApp.AutoSize = True
        Me.LabelExternalCompareApp.Location = New System.Drawing.Point(12, 41)
        Me.LabelExternalCompareApp.Name = "LabelExternalCompareApp"
        Me.LabelExternalCompareApp.Size = New System.Drawing.Size(157, 13)
        Me.LabelExternalCompareApp.TabIndex = 2
        Me.LabelExternalCompareApp.Text = "External compare program path:"
        '
        'CheckBoxUseExternalCompare
        '
        Me.CheckBoxUseExternalCompare.AutoSize = True
        Me.CheckBoxUseExternalCompare.Location = New System.Drawing.Point(175, 14)
        Me.CheckBoxUseExternalCompare.Name = "CheckBoxUseExternalCompare"
        Me.CheckBoxUseExternalCompare.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxUseExternalCompare.TabIndex = 1
        Me.CheckBoxUseExternalCompare.UseVisualStyleBackColor = True
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(286, 64)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 23)
        Me.ButtonSave.TabIndex = 4
        Me.ButtonSave.Text = "Save"
        Me.ButtonSave.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(367, 64)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 5
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'FormExternalCompare
        '
        Me.AcceptButton = Me.ButtonSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(729, 99)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonSave)
        Me.Controls.Add(Me.CheckBoxUseExternalCompare)
        Me.Controls.Add(Me.LabelExternalCompareApp)
        Me.Controls.Add(Me.LabelUseExternalCompare)
        Me.Controls.Add(Me.TextBoxExternalCompareApp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormExternalCompare"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "External Compare Program Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBoxExternalCompareApp As System.Windows.Forms.TextBox
    Friend WithEvents LabelUseExternalCompare As System.Windows.Forms.Label
    Friend WithEvents LabelExternalCompareApp As System.Windows.Forms.Label
    Friend WithEvents CheckBoxUseExternalCompare As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
End Class
