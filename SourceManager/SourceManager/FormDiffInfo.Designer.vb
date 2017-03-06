<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormDiffInfo
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
        Me.ButtonOk = New System.Windows.Forms.Button()
        Me.LabelFrom = New System.Windows.Forms.Label()
        Me.LabelFromNewerOlder = New System.Windows.Forms.Label()
        Me.LabelFromSize = New System.Windows.Forms.Label()
        Me.LabelFromDate = New System.Windows.Forms.Label()
        Me.LabelToDate = New System.Windows.Forms.Label()
        Me.LabelToSize = New System.Windows.Forms.Label()
        Me.LabelToNewerOlder = New System.Windows.Forms.Label()
        Me.LabelTo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ButtonOk
        '
        Me.ButtonOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonOk.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonOk.Location = New System.Drawing.Point(158, 217)
        Me.ButtonOk.Name = "ButtonOk"
        Me.ButtonOk.Size = New System.Drawing.Size(75, 23)
        Me.ButtonOk.TabIndex = 8
        Me.ButtonOk.Text = "OK"
        Me.ButtonOk.UseVisualStyleBackColor = True
        '
        'LabelFrom
        '
        Me.LabelFrom.AutoSize = True
        Me.LabelFrom.Location = New System.Drawing.Point(12, 9)
        Me.LabelFrom.Name = "LabelFrom"
        Me.LabelFrom.Size = New System.Drawing.Size(33, 13)
        Me.LabelFrom.TabIndex = 0
        Me.LabelFrom.Text = "From:"
        '
        'LabelFromNewerOlder
        '
        Me.LabelFromNewerOlder.AutoSize = True
        Me.LabelFromNewerOlder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFromNewerOlder.Location = New System.Drawing.Point(51, 9)
        Me.LabelFromNewerOlder.Name = "LabelFromNewerOlder"
        Me.LabelFromNewerOlder.Size = New System.Drawing.Size(43, 13)
        Me.LabelFromNewerOlder.TabIndex = 1
        Me.LabelFromNewerOlder.Text = "Newer"
        '
        'LabelFromSize
        '
        Me.LabelFromSize.AutoSize = True
        Me.LabelFromSize.Location = New System.Drawing.Point(51, 61)
        Me.LabelFromSize.Name = "LabelFromSize"
        Me.LabelFromSize.Size = New System.Drawing.Size(50, 13)
        Me.LabelFromSize.TabIndex = 3
        Me.LabelFromSize.Text = "FromSize"
        '
        'LabelFromDate
        '
        Me.LabelFromDate.AutoSize = True
        Me.LabelFromDate.Location = New System.Drawing.Point(51, 35)
        Me.LabelFromDate.Name = "LabelFromDate"
        Me.LabelFromDate.Size = New System.Drawing.Size(53, 13)
        Me.LabelFromDate.TabIndex = 2
        Me.LabelFromDate.Text = "FromDate"
        '
        'LabelToDate
        '
        Me.LabelToDate.AutoSize = True
        Me.LabelToDate.Location = New System.Drawing.Point(51, 139)
        Me.LabelToDate.Name = "LabelToDate"
        Me.LabelToDate.Size = New System.Drawing.Size(43, 13)
        Me.LabelToDate.TabIndex = 6
        Me.LabelToDate.Text = "ToDate"
        '
        'LabelToSize
        '
        Me.LabelToSize.AutoSize = True
        Me.LabelToSize.Location = New System.Drawing.Point(51, 165)
        Me.LabelToSize.Name = "LabelToSize"
        Me.LabelToSize.Size = New System.Drawing.Size(40, 13)
        Me.LabelToSize.TabIndex = 7
        Me.LabelToSize.Text = "ToSize"
        '
        'LabelToNewerOlder
        '
        Me.LabelToNewerOlder.AutoSize = True
        Me.LabelToNewerOlder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelToNewerOlder.Location = New System.Drawing.Point(51, 113)
        Me.LabelToNewerOlder.Name = "LabelToNewerOlder"
        Me.LabelToNewerOlder.Size = New System.Drawing.Size(43, 13)
        Me.LabelToNewerOlder.TabIndex = 5
        Me.LabelToNewerOlder.Text = "Newer"
        '
        'LabelTo
        '
        Me.LabelTo.AutoSize = True
        Me.LabelTo.Location = New System.Drawing.Point(12, 113)
        Me.LabelTo.Name = "LabelTo"
        Me.LabelTo.Size = New System.Drawing.Size(23, 13)
        Me.LabelTo.TabIndex = 4
        Me.LabelTo.Text = "To:"
        '
        'FormDiffInfo
        '
        Me.AcceptButton = Me.ButtonOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonOk
        Me.ClientSize = New System.Drawing.Size(390, 252)
        Me.ControlBox = False
        Me.Controls.Add(Me.LabelToDate)
        Me.Controls.Add(Me.LabelToSize)
        Me.Controls.Add(Me.LabelToNewerOlder)
        Me.Controls.Add(Me.LabelTo)
        Me.Controls.Add(Me.LabelFromDate)
        Me.Controls.Add(Me.LabelFromSize)
        Me.Controls.Add(Me.LabelFromNewerOlder)
        Me.Controls.Add(Me.LabelFrom)
        Me.Controls.Add(Me.ButtonOk)
        Me.Name = "FormDiffInfo"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "FormDiffInfo"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonOk As System.Windows.Forms.Button
    Friend WithEvents LabelFrom As System.Windows.Forms.Label
    Friend WithEvents LabelFromNewerOlder As System.Windows.Forms.Label
    Friend WithEvents LabelFromSize As System.Windows.Forms.Label
    Friend WithEvents LabelFromDate As System.Windows.Forms.Label
    Friend WithEvents LabelToDate As System.Windows.Forms.Label
    Friend WithEvents LabelToSize As System.Windows.Forms.Label
    Friend WithEvents LabelToNewerOlder As System.Windows.Forms.Label
    Friend WithEvents LabelTo As System.Windows.Forms.Label
End Class
