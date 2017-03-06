<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTargetNewer
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
        Me.ButtonNo = New System.Windows.Forms.Button()
        Me.ButtonNoToAll = New System.Windows.Forms.Button()
        Me.ButtonYes = New System.Windows.Forms.Button()
        Me.ButtonYesToAll = New System.Windows.Forms.Button()
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
        Me.LabelFilename.TabIndex = 1
        Me.LabelFilename.Text = "<Filename>"
        Me.LabelFilename.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ButtonNo
        '
        Me.ButtonNo.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonNo.Location = New System.Drawing.Point(174, 85)
        Me.ButtonNo.Name = "ButtonNo"
        Me.ButtonNo.Size = New System.Drawing.Size(75, 23)
        Me.ButtonNo.TabIndex = 2
        Me.ButtonNo.Text = "&No"
        Me.ButtonNo.UseVisualStyleBackColor = True
        '
        'ButtonNoToAll
        '
        Me.ButtonNoToAll.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonNoToAll.Location = New System.Drawing.Point(255, 85)
        Me.ButtonNoToAll.Name = "ButtonNoToAll"
        Me.ButtonNoToAll.Size = New System.Drawing.Size(75, 23)
        Me.ButtonNoToAll.TabIndex = 3
        Me.ButtonNoToAll.Text = "No To All"
        Me.ButtonNoToAll.UseVisualStyleBackColor = True
        '
        'ButtonYes
        '
        Me.ButtonYes.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonYes.Location = New System.Drawing.Point(12, 85)
        Me.ButtonYes.Name = "ButtonYes"
        Me.ButtonYes.Size = New System.Drawing.Size(75, 23)
        Me.ButtonYes.TabIndex = 5
        Me.ButtonYes.Text = "&Yes"
        Me.ButtonYes.UseVisualStyleBackColor = True
        '
        'ButtonYesToAll
        '
        Me.ButtonYesToAll.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonYesToAll.Location = New System.Drawing.Point(93, 85)
        Me.ButtonYesToAll.Name = "ButtonYesToAll"
        Me.ButtonYesToAll.Size = New System.Drawing.Size(75, 23)
        Me.ButtonYesToAll.TabIndex = 6
        Me.ButtonYesToAll.Text = "Yes To All"
        Me.ButtonYesToAll.UseVisualStyleBackColor = True
        '
        'ButtonAbort
        '
        Me.ButtonAbort.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonAbort.Location = New System.Drawing.Point(336, 85)
        Me.ButtonAbort.Name = "ButtonAbort"
        Me.ButtonAbort.Size = New System.Drawing.Size(75, 23)
        Me.ButtonAbort.TabIndex = 4
        Me.ButtonAbort.Text = "Abort"
        Me.ButtonAbort.UseVisualStyleBackColor = True
        '
        'FormTargetNewer
        '
        Me.AcceptButton = Me.ButtonNo
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonAbort
        Me.ClientSize = New System.Drawing.Size(423, 120)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonAbort)
        Me.Controls.Add(Me.ButtonYesToAll)
        Me.Controls.Add(Me.ButtonYes)
        Me.Controls.Add(Me.ButtonNoToAll)
        Me.Controls.Add(Me.ButtonNo)
        Me.Controls.Add(Me.LabelFilename)
        Me.Name = "FormTargetNewer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Target file is newer. Overwrite?"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelFilename As System.Windows.Forms.Label
    Friend WithEvents ButtonNo As System.Windows.Forms.Button
    Friend WithEvents ButtonNoToAll As System.Windows.Forms.Button
    Friend WithEvents ButtonYes As System.Windows.Forms.Button
    Friend WithEvents ButtonYesToAll As System.Windows.Forms.Button
    Friend WithEvents ButtonAbort As System.Windows.Forms.Button
End Class
