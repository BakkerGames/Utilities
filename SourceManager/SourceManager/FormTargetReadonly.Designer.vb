<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTargetReadonly
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
        Me.LabelFilename = New System.Windows.Forms.Label()
        Me.ButtonIgnore = New System.Windows.Forms.Button()
        Me.ButtonIgnoreAll = New System.Windows.Forms.Button()
        Me.ButtonRetry = New System.Windows.Forms.Button()
        Me.ButtonAbort = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LabelFilename
        '
        Me.LabelFilename.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelFilename.Location = New System.Drawing.Point(12, 9)
        Me.LabelFilename.Name = "LabelFilename"
        Me.LabelFilename.Size = New System.Drawing.Size(399, 73)
        Me.LabelFilename.TabIndex = 0
        Me.LabelFilename.Text = "<Filename>"
        Me.LabelFilename.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ButtonIgnore
        '
        Me.ButtonIgnore.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonIgnore.Location = New System.Drawing.Point(214, 85)
        Me.ButtonIgnore.Name = "ButtonIgnore"
        Me.ButtonIgnore.Size = New System.Drawing.Size(75, 23)
        Me.ButtonIgnore.TabIndex = 3
        Me.ButtonIgnore.Text = "&Ignore"
        Me.ButtonIgnore.UseVisualStyleBackColor = True
        '
        'ButtonIgnoreAll
        '
        Me.ButtonIgnoreAll.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonIgnoreAll.Location = New System.Drawing.Point(295, 85)
        Me.ButtonIgnoreAll.Name = "ButtonIgnoreAll"
        Me.ButtonIgnoreAll.Size = New System.Drawing.Size(75, 23)
        Me.ButtonIgnoreAll.TabIndex = 4
        Me.ButtonIgnoreAll.Text = "Ignore &All"
        Me.ButtonIgnoreAll.UseVisualStyleBackColor = True
        '
        'ButtonRetry
        '
        Me.ButtonRetry.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonRetry.Location = New System.Drawing.Point(133, 85)
        Me.ButtonRetry.Name = "ButtonRetry"
        Me.ButtonRetry.Size = New System.Drawing.Size(75, 23)
        Me.ButtonRetry.TabIndex = 2
        Me.ButtonRetry.Text = "&Retry"
        Me.ButtonRetry.UseVisualStyleBackColor = True
        '
        'ButtonAbort
        '
        Me.ButtonAbort.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonAbort.Location = New System.Drawing.Point(52, 85)
        Me.ButtonAbort.Name = "ButtonAbort"
        Me.ButtonAbort.Size = New System.Drawing.Size(75, 23)
        Me.ButtonAbort.TabIndex = 1
        Me.ButtonAbort.Text = "&Abort"
        Me.ButtonAbort.UseVisualStyleBackColor = True
        '
        'FormTargetReadonly
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonAbort
        Me.ClientSize = New System.Drawing.Size(423, 120)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonAbort)
        Me.Controls.Add(Me.ButtonRetry)
        Me.Controls.Add(Me.ButtonIgnoreAll)
        Me.Controls.Add(Me.ButtonIgnore)
        Me.Controls.Add(Me.LabelFilename)
        Me.Name = "FormTargetReadonly"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Target file is Read-Only"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelFilename As System.Windows.Forms.Label
    Friend WithEvents ButtonIgnore As System.Windows.Forms.Button
    Friend WithEvents ButtonIgnoreAll As System.Windows.Forms.Button
    Friend WithEvents ButtonRetry As System.Windows.Forms.Button
    Friend WithEvents ButtonAbort As System.Windows.Forms.Button
End Class
