<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormDiff
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormDiff))
        Me.PanelTop = New System.Windows.Forms.Panel()
        Me.LabelFont = New System.Windows.Forms.Label()
        Me.ComboBoxFont = New System.Windows.Forms.ComboBox()
        Me.PanelMain = New System.Windows.Forms.Panel()
        Me.TextBoxDiff = New System.Windows.Forms.TextBox()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.PanelTop.SuspendLayout()
        Me.PanelMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelTop
        '
        Me.PanelTop.Controls.Add(Me.LabelFont)
        Me.PanelTop.Controls.Add(Me.ComboBoxFont)
        Me.PanelTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelTop.Location = New System.Drawing.Point(0, 0)
        Me.PanelTop.Name = "PanelTop"
        Me.PanelTop.Size = New System.Drawing.Size(792, 23)
        Me.PanelTop.TabIndex = 1
        '
        'LabelFont
        '
        Me.LabelFont.AutoSize = True
        Me.LabelFont.Location = New System.Drawing.Point(13, 5)
        Me.LabelFont.Name = "LabelFont"
        Me.LabelFont.Size = New System.Drawing.Size(31, 13)
        Me.LabelFont.TabIndex = 0
        Me.LabelFont.Text = "Font:"
        '
        'ComboBoxFont
        '
        Me.ComboBoxFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFont.FormattingEnabled = True
        Me.ComboBoxFont.Items.AddRange(New Object() {"Proportional", "Fixed"})
        Me.ComboBoxFont.Location = New System.Drawing.Point(50, 1)
        Me.ComboBoxFont.Name = "ComboBoxFont"
        Me.ComboBoxFont.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxFont.TabIndex = 1
        '
        'PanelMain
        '
        Me.PanelMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelMain.Controls.Add(Me.TextBoxDiff)
        Me.PanelMain.Location = New System.Drawing.Point(0, 23)
        Me.PanelMain.Name = "PanelMain"
        Me.PanelMain.Size = New System.Drawing.Size(792, 500)
        Me.PanelMain.TabIndex = 0
        '
        'TextBoxDiff
        '
        Me.TextBoxDiff.BackColor = System.Drawing.Color.White
        Me.TextBoxDiff.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBoxDiff.Location = New System.Drawing.Point(0, 0)
        Me.TextBoxDiff.Multiline = True
        Me.TextBoxDiff.Name = "TextBoxDiff"
        Me.TextBoxDiff.ReadOnly = True
        Me.TextBoxDiff.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBoxDiff.Size = New System.Drawing.Size(792, 500)
        Me.TextBoxDiff.TabIndex = 0
        Me.TextBoxDiff.WordWrap = False
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(359, 250)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 3
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'FormDiff
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 523)
        Me.Controls.Add(Me.PanelTop)
        Me.Controls.Add(Me.PanelMain)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "FormDiff"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "FormDiff"
        Me.PanelTop.ResumeLayout(False)
        Me.PanelTop.PerformLayout()
        Me.PanelMain.ResumeLayout(False)
        Me.PanelMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelTop As System.Windows.Forms.Panel
    Friend WithEvents LabelFont As System.Windows.Forms.Label
    Friend WithEvents ComboBoxFont As System.Windows.Forms.ComboBox
    Friend WithEvents PanelMain As System.Windows.Forms.Panel
    Friend WithEvents TextBoxDiff As System.Windows.Forms.TextBox
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
End Class
