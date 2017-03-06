<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddApplicationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelMain = New System.Windows.Forms.ToolStripStatusLabel()
        Me.PanelMain = New System.Windows.Forms.Panel()
        Me.LabelApplication = New System.Windows.Forms.Label()
        Me.TextBoxSourceDir = New System.Windows.Forms.TextBox()
        Me.TextBoxTransferDir = New System.Windows.Forms.TextBox()
        Me.TextBoxMergeDir = New System.Windows.Forms.TextBox()
        Me.ButtonUnzipChanges = New System.Windows.Forms.Button()
        Me.ButtonZipChanges = New System.Windows.Forms.Button()
        Me.LabelTransferDir = New System.Windows.Forms.Label()
        Me.LabelMergeDir = New System.Windows.Forms.Label()
        Me.LabelSourceDir = New System.Windows.Forms.Label()
        Me.ComboBoxApplication = New System.Windows.Forms.ComboBox()
        Me.ButtonCompare = New System.Windows.Forms.Button()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        Me.PanelMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.OptionsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(416, 24)
        Me.MenuStripMain.TabIndex = 0
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddApplicationToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'AddApplicationToolStripMenuItem
        '
        Me.AddApplicationToolStripMenuItem.Name = "AddApplicationToolStripMenuItem"
        Me.AddApplicationToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.AddApplicationToolStripMenuItem.Text = "&Add Application"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.AboutToolStripMenuItem.Text = "&About..."
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelMain})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 182)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(416, 22)
        Me.StatusStripMain.SizingGrip = False
        Me.StatusStripMain.TabIndex = 2
        '
        'ToolStripStatusLabelMain
        '
        Me.ToolStripStatusLabelMain.Name = "ToolStripStatusLabelMain"
        Me.ToolStripStatusLabelMain.Size = New System.Drawing.Size(401, 17)
        Me.ToolStripStatusLabelMain.Spring = True
        Me.ToolStripStatusLabelMain.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PanelMain
        '
        Me.PanelMain.Controls.Add(Me.ButtonCompare)
        Me.PanelMain.Controls.Add(Me.LabelApplication)
        Me.PanelMain.Controls.Add(Me.TextBoxSourceDir)
        Me.PanelMain.Controls.Add(Me.TextBoxTransferDir)
        Me.PanelMain.Controls.Add(Me.TextBoxMergeDir)
        Me.PanelMain.Controls.Add(Me.ButtonUnzipChanges)
        Me.PanelMain.Controls.Add(Me.ButtonZipChanges)
        Me.PanelMain.Controls.Add(Me.LabelTransferDir)
        Me.PanelMain.Controls.Add(Me.LabelMergeDir)
        Me.PanelMain.Controls.Add(Me.LabelSourceDir)
        Me.PanelMain.Controls.Add(Me.ComboBoxApplication)
        Me.PanelMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelMain.Location = New System.Drawing.Point(0, 24)
        Me.PanelMain.Name = "PanelMain"
        Me.PanelMain.Size = New System.Drawing.Size(416, 158)
        Me.PanelMain.TabIndex = 1
        '
        'LabelApplication
        '
        Me.LabelApplication.AutoSize = True
        Me.LabelApplication.Location = New System.Drawing.Point(12, 18)
        Me.LabelApplication.Name = "LabelApplication"
        Me.LabelApplication.Size = New System.Drawing.Size(62, 13)
        Me.LabelApplication.TabIndex = 0
        Me.LabelApplication.Text = "Application:"
        '
        'TextBoxSourceDir
        '
        Me.TextBoxSourceDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxSourceDir.Location = New System.Drawing.Point(116, 42)
        Me.TextBoxSourceDir.Name = "TextBoxSourceDir"
        Me.TextBoxSourceDir.Size = New System.Drawing.Size(288, 20)
        Me.TextBoxSourceDir.TabIndex = 3
        '
        'TextBoxTransferDir
        '
        Me.TextBoxTransferDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxTransferDir.Location = New System.Drawing.Point(116, 96)
        Me.TextBoxTransferDir.Name = "TextBoxTransferDir"
        Me.TextBoxTransferDir.Size = New System.Drawing.Size(288, 20)
        Me.TextBoxTransferDir.TabIndex = 7
        '
        'TextBoxMergeDir
        '
        Me.TextBoxMergeDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxMergeDir.Location = New System.Drawing.Point(116, 69)
        Me.TextBoxMergeDir.Name = "TextBoxMergeDir"
        Me.TextBoxMergeDir.Size = New System.Drawing.Size(288, 20)
        Me.TextBoxMergeDir.TabIndex = 5
        '
        'ButtonUnzipChanges
        '
        Me.ButtonUnzipChanges.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonUnzipChanges.Location = New System.Drawing.Point(264, 124)
        Me.ButtonUnzipChanges.Name = "ButtonUnzipChanges"
        Me.ButtonUnzipChanges.Size = New System.Drawing.Size(100, 23)
        Me.ButtonUnzipChanges.TabIndex = 10
        Me.ButtonUnzipChanges.Text = "Unzip Changes"
        Me.ButtonUnzipChanges.UseVisualStyleBackColor = True
        '
        'ButtonZipChanges
        '
        Me.ButtonZipChanges.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonZipChanges.Location = New System.Drawing.Point(158, 124)
        Me.ButtonZipChanges.Name = "ButtonZipChanges"
        Me.ButtonZipChanges.Size = New System.Drawing.Size(100, 23)
        Me.ButtonZipChanges.TabIndex = 9
        Me.ButtonZipChanges.Text = "Zip Changes"
        Me.ButtonZipChanges.UseVisualStyleBackColor = True
        '
        'LabelTransferDir
        '
        Me.LabelTransferDir.AutoSize = True
        Me.LabelTransferDir.Location = New System.Drawing.Point(12, 99)
        Me.LabelTransferDir.Name = "LabelTransferDir"
        Me.LabelTransferDir.Size = New System.Drawing.Size(94, 13)
        Me.LabelTransferDir.TabIndex = 6
        Me.LabelTransferDir.Text = "Transfer Directory:"
        '
        'LabelMergeDir
        '
        Me.LabelMergeDir.AutoSize = True
        Me.LabelMergeDir.Location = New System.Drawing.Point(12, 72)
        Me.LabelMergeDir.Name = "LabelMergeDir"
        Me.LabelMergeDir.Size = New System.Drawing.Size(85, 13)
        Me.LabelMergeDir.TabIndex = 4
        Me.LabelMergeDir.Text = "Merge Directory:"
        '
        'LabelSourceDir
        '
        Me.LabelSourceDir.AutoSize = True
        Me.LabelSourceDir.Location = New System.Drawing.Point(12, 45)
        Me.LabelSourceDir.Name = "LabelSourceDir"
        Me.LabelSourceDir.Size = New System.Drawing.Size(89, 13)
        Me.LabelSourceDir.TabIndex = 2
        Me.LabelSourceDir.Text = "Source Directory:"
        '
        'ComboBoxApplication
        '
        Me.ComboBoxApplication.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxApplication.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxApplication.FormattingEnabled = True
        Me.ComboBoxApplication.Location = New System.Drawing.Point(116, 15)
        Me.ComboBoxApplication.Name = "ComboBoxApplication"
        Me.ComboBoxApplication.Size = New System.Drawing.Size(288, 21)
        Me.ComboBoxApplication.Sorted = True
        Me.ComboBoxApplication.TabIndex = 1
        '
        'ButtonCompare
        '
        Me.ButtonCompare.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonCompare.Location = New System.Drawing.Point(52, 124)
        Me.ButtonCompare.Name = "ButtonCompare"
        Me.ButtonCompare.Size = New System.Drawing.Size(100, 23)
        Me.ButtonCompare.TabIndex = 8
        Me.ButtonCompare.Text = "Compare"
        Me.ButtonCompare.UseVisualStyleBackColor = True
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(416, 204)
        Me.Controls.Add(Me.PanelMain)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.MenuStripMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transfer Files"
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.PanelMain.ResumeLayout(False)
        Me.PanelMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents PanelMain As System.Windows.Forms.Panel
    Friend WithEvents LabelTransferDir As System.Windows.Forms.Label
    Friend WithEvents LabelMergeDir As System.Windows.Forms.Label
    Friend WithEvents LabelSourceDir As System.Windows.Forms.Label
    Friend WithEvents ComboBoxApplication As System.Windows.Forms.ComboBox
    Friend WithEvents ButtonUnzipChanges As System.Windows.Forms.Button
    Friend WithEvents ButtonZipChanges As System.Windows.Forms.Button
    Friend WithEvents ToolStripStatusLabelMain As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TextBoxTransferDir As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxMergeDir As System.Windows.Forms.TextBox
    Friend WithEvents LabelApplication As System.Windows.Forms.Label
    Friend WithEvents TextBoxSourceDir As System.Windows.Forms.TextBox
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddApplicationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ButtonCompare As System.Windows.Forms.Button

End Class
