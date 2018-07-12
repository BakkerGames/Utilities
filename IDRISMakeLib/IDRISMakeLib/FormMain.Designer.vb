<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Me.StatusLabel = New System.Windows.Forms.Label()
        Me.EnvCombo = New System.Windows.Forms.ComboBox()
        Me.EnvLabel = New System.Windows.Forms.Label()
        Me.CheckCompileLibs = New System.Windows.Forms.CheckBox()
        Me.CheckAddComments = New System.Windows.Forms.CheckBox()
        Me.VB6Prog = New System.Windows.Forms.RichTextBox()
        Me.CadolProg = New System.Windows.Forms.RichTextBox()
        Me.StartButton = New System.Windows.Forms.Button()
        Me.LibraryCombo = New System.Windows.Forms.ComboBox()
        Me.LibraryLabel = New System.Windows.Forms.Label()
        Me.TargetLabel = New System.Windows.Forms.Label()
        Me.SourceLabel = New System.Windows.Forms.Label()
        Me.VolumeCombo = New System.Windows.Forms.ComboBox()
        Me.VolumeLabel = New System.Windows.Forms.Label()
        Me.TargetPath = New System.Windows.Forms.TextBox()
        Me.SourcePath = New System.Windows.Forms.TextBox()
        Me.StopCompiling = New System.Windows.Forms.Button()
        Me.CheckChangedOnly = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CancelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IndexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BrowseEnvironment = New System.Windows.Forms.OpenFileDialog()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusLabel
        '
        Me.StatusLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StatusLabel.Location = New System.Drawing.Point(1, 418)
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(481, 15)
        Me.StatusLabel.TabIndex = 15
        Me.StatusLabel.Text = "Select Environment"
        '
        'EnvCombo
        '
        Me.EnvCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.EnvCombo.FormattingEnabled = True
        Me.EnvCombo.Items.AddRange(New Object() {"PC", "PC-D", "Local", "Test", "Accept", "Prod", "FIS", "FISTest", "EOY", "<Browse>"})
        Me.EnvCombo.Location = New System.Drawing.Point(69, 34)
        Me.EnvCombo.Name = "EnvCombo"
        Me.EnvCombo.Size = New System.Drawing.Size(85, 21)
        Me.EnvCombo.TabIndex = 1
        '
        'EnvLabel
        '
        Me.EnvLabel.AutoSize = True
        Me.EnvLabel.Location = New System.Drawing.Point(1, 37)
        Me.EnvLabel.Name = "EnvLabel"
        Me.EnvLabel.Size = New System.Drawing.Size(66, 13)
        Me.EnvLabel.TabIndex = 0
        Me.EnvLabel.Text = "Environment"
        '
        'CheckCompileLibs
        '
        Me.CheckCompileLibs.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.CheckCompileLibs.AutoSize = True
        Me.CheckCompileLibs.Checked = True
        Me.CheckCompileLibs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckCompileLibs.Location = New System.Drawing.Point(304, 92)
        Me.CheckCompileLibs.Name = "CheckCompileLibs"
        Me.CheckCompileLibs.Size = New System.Drawing.Size(85, 17)
        Me.CheckCompileLibs.TabIndex = 12
        Me.CheckCompileLibs.Text = "Compile Libs"
        Me.CheckCompileLibs.UseVisualStyleBackColor = True
        '
        'CheckAddComments
        '
        Me.CheckAddComments.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.CheckAddComments.AutoSize = True
        Me.CheckAddComments.Enabled = False
        Me.CheckAddComments.Location = New System.Drawing.Point(201, 92)
        Me.CheckAddComments.Name = "CheckAddComments"
        Me.CheckAddComments.Size = New System.Drawing.Size(97, 17)
        Me.CheckAddComments.TabIndex = 11
        Me.CheckAddComments.Text = "Add Comments"
        Me.CheckAddComments.UseVisualStyleBackColor = True
        '
        'VB6Prog
        '
        Me.VB6Prog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.VB6Prog.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VB6Prog.Location = New System.Drawing.Point(243, 140)
        Me.VB6Prog.Name = "VB6Prog"
        Me.VB6Prog.Size = New System.Drawing.Size(239, 273)
        Me.VB6Prog.TabIndex = 16
        Me.VB6Prog.Text = ""
        Me.VB6Prog.WordWrap = False
        '
        'CadolProg
        '
        Me.CadolProg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CadolProg.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CadolProg.Location = New System.Drawing.Point(2, 140)
        Me.CadolProg.Name = "CadolProg"
        Me.CadolProg.Size = New System.Drawing.Size(239, 273)
        Me.CadolProg.TabIndex = 15
        Me.CadolProg.Text = ""
        Me.CadolProg.WordWrap = False
        '
        'StartButton
        '
        Me.StartButton.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.StartButton.Location = New System.Drawing.Point(161, 113)
        Me.StartButton.Name = "StartButton"
        Me.StartButton.Size = New System.Drawing.Size(79, 23)
        Me.StartButton.TabIndex = 13
        Me.StartButton.Text = "Start"
        Me.StartButton.UseVisualStyleBackColor = True
        '
        'LibraryCombo
        '
        Me.LibraryCombo.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LibraryCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LibraryCombo.FormattingEnabled = True
        Me.LibraryCombo.Location = New System.Drawing.Point(294, 67)
        Me.LibraryCombo.Name = "LibraryCombo"
        Me.LibraryCombo.Size = New System.Drawing.Size(118, 21)
        Me.LibraryCombo.Sorted = True
        Me.LibraryCombo.TabIndex = 9
        '
        'LibraryLabel
        '
        Me.LibraryLabel.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LibraryLabel.AutoSize = True
        Me.LibraryLabel.Location = New System.Drawing.Point(247, 70)
        Me.LibraryLabel.Name = "LibraryLabel"
        Me.LibraryLabel.Size = New System.Drawing.Size(38, 13)
        Me.LibraryLabel.TabIndex = 8
        Me.LibraryLabel.Text = "Library"
        '
        'TargetLabel
        '
        Me.TargetLabel.AutoSize = True
        Me.TargetLabel.Location = New System.Drawing.Point(160, 49)
        Me.TargetLabel.Name = "TargetLabel"
        Me.TargetLabel.Size = New System.Drawing.Size(63, 13)
        Me.TargetLabel.TabIndex = 4
        Me.TargetLabel.Text = "Target Path"
        '
        'SourceLabel
        '
        Me.SourceLabel.AutoSize = True
        Me.SourceLabel.Location = New System.Drawing.Point(160, 28)
        Me.SourceLabel.Name = "SourceLabel"
        Me.SourceLabel.Size = New System.Drawing.Size(66, 13)
        Me.SourceLabel.TabIndex = 2
        Me.SourceLabel.Text = "Source Path"
        '
        'VolumeCombo
        '
        Me.VolumeCombo.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.VolumeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.VolumeCombo.FormattingEnabled = True
        Me.VolumeCombo.Location = New System.Drawing.Point(120, 67)
        Me.VolumeCombo.Name = "VolumeCombo"
        Me.VolumeCombo.Size = New System.Drawing.Size(118, 21)
        Me.VolumeCombo.Sorted = True
        Me.VolumeCombo.TabIndex = 7
        '
        'VolumeLabel
        '
        Me.VolumeLabel.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.VolumeLabel.AutoSize = True
        Me.VolumeLabel.Location = New System.Drawing.Point(73, 70)
        Me.VolumeLabel.Name = "VolumeLabel"
        Me.VolumeLabel.Size = New System.Drawing.Size(42, 13)
        Me.VolumeLabel.TabIndex = 6
        Me.VolumeLabel.Text = "Volume"
        '
        'TargetPath
        '
        Me.TargetPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TargetPath.Location = New System.Drawing.Point(228, 46)
        Me.TargetPath.Name = "TargetPath"
        Me.TargetPath.Size = New System.Drawing.Size(254, 20)
        Me.TargetPath.TabIndex = 5
        '
        'SourcePath
        '
        Me.SourcePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SourcePath.Location = New System.Drawing.Point(228, 25)
        Me.SourcePath.Name = "SourcePath"
        Me.SourcePath.Size = New System.Drawing.Size(254, 20)
        Me.SourcePath.TabIndex = 3
        '
        'StopCompiling
        '
        Me.StopCompiling.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.StopCompiling.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.StopCompiling.Enabled = False
        Me.StopCompiling.Location = New System.Drawing.Point(244, 113)
        Me.StopCompiling.Name = "StopCompiling"
        Me.StopCompiling.Size = New System.Drawing.Size(79, 23)
        Me.StopCompiling.TabIndex = 14
        Me.StopCompiling.Text = "Cancel"
        Me.StopCompiling.UseVisualStyleBackColor = True
        '
        'CheckChangedOnly
        '
        Me.CheckChangedOnly.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.CheckChangedOnly.AutoSize = True
        Me.CheckChangedOnly.Checked = True
        Me.CheckChangedOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckChangedOnly.Location = New System.Drawing.Point(98, 92)
        Me.CheckChangedOnly.Name = "CheckChangedOnly"
        Me.CheckChangedOnly.Size = New System.Drawing.Size(93, 17)
        Me.CheckChangedOnly.TabIndex = 10
        Me.CheckChangedOnly.Text = "Changed Only"
        Me.CheckChangedOnly.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.CompileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(486, 24)
        Me.MenuStrip1.TabIndex = 17
        Me.MenuStrip1.Text = "MenuStrip1"
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
        'CompileToolStripMenuItem
        '
        Me.CompileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StartToolStripMenuItem, Me.CancelToolStripMenuItem})
        Me.CompileToolStripMenuItem.Name = "CompileToolStripMenuItem"
        Me.CompileToolStripMenuItem.Size = New System.Drawing.Size(64, 20)
        Me.CompileToolStripMenuItem.Text = "&Compile"
        '
        'StartToolStripMenuItem
        '
        Me.StartToolStripMenuItem.Name = "StartToolStripMenuItem"
        Me.StartToolStripMenuItem.Size = New System.Drawing.Size(110, 22)
        Me.StartToolStripMenuItem.Text = "&Start"
        '
        'CancelToolStripMenuItem
        '
        Me.CancelToolStripMenuItem.Enabled = False
        Me.CancelToolStripMenuItem.Name = "CancelToolStripMenuItem"
        Me.CancelToolStripMenuItem.Size = New System.Drawing.Size(110, 22)
        Me.CancelToolStripMenuItem.Text = "&Cancel"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomizeToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'CustomizeToolStripMenuItem
        '
        Me.CustomizeToolStripMenuItem.Name = "CustomizeToolStripMenuItem"
        Me.CustomizeToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.CustomizeToolStripMenuItem.Text = "&Customize"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ContentsToolStripMenuItem, Me.IndexToolStripMenuItem, Me.SearchToolStripMenuItem, Me.toolStripSeparator5, Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'ContentsToolStripMenuItem
        '
        Me.ContentsToolStripMenuItem.Name = "ContentsToolStripMenuItem"
        Me.ContentsToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.ContentsToolStripMenuItem.Text = "&Contents"
        '
        'IndexToolStripMenuItem
        '
        Me.IndexToolStripMenuItem.Name = "IndexToolStripMenuItem"
        Me.IndexToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.IndexToolStripMenuItem.Text = "&Index"
        '
        'SearchToolStripMenuItem
        '
        Me.SearchToolStripMenuItem.Name = "SearchToolStripMenuItem"
        Me.SearchToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.SearchToolStripMenuItem.Text = "&Search"
        '
        'toolStripSeparator5
        '
        Me.toolStripSeparator5.Name = "toolStripSeparator5"
        Me.toolStripSeparator5.Size = New System.Drawing.Size(119, 6)
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'BrowseEnvironment
        '
        Me.BrowseEnvironment.DefaultExt = "ini"
        Me.BrowseEnvironment.Filter = "Connection Info|Connect.ini"
        Me.BrowseEnvironment.InitialDirectory = "C:\"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'FormMain
        '
        Me.AcceptButton = Me.StartButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.StopCompiling
        Me.ClientSize = New System.Drawing.Size(486, 435)
        Me.Controls.Add(Me.CheckChangedOnly)
        Me.Controls.Add(Me.StopCompiling)
        Me.Controls.Add(Me.StatusLabel)
        Me.Controls.Add(Me.EnvCombo)
        Me.Controls.Add(Me.EnvLabel)
        Me.Controls.Add(Me.CheckCompileLibs)
        Me.Controls.Add(Me.CheckAddComments)
        Me.Controls.Add(Me.VB6Prog)
        Me.Controls.Add(Me.CadolProg)
        Me.Controls.Add(Me.StartButton)
        Me.Controls.Add(Me.LibraryCombo)
        Me.Controls.Add(Me.LibraryLabel)
        Me.Controls.Add(Me.TargetLabel)
        Me.Controls.Add(Me.SourceLabel)
        Me.Controls.Add(Me.VolumeCombo)
        Me.Controls.Add(Me.VolumeLabel)
        Me.Controls.Add(Me.TargetPath)
        Me.Controls.Add(Me.SourcePath)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormMain"
        Me.Text = "Compile IDRIS Library"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusLabel As System.Windows.Forms.Label
    Friend WithEvents EnvCombo As System.Windows.Forms.ComboBox
    Friend WithEvents EnvLabel As System.Windows.Forms.Label
    Friend WithEvents CheckCompileLibs As System.Windows.Forms.CheckBox
    Friend WithEvents CheckAddComments As System.Windows.Forms.CheckBox
    Friend WithEvents VB6Prog As System.Windows.Forms.RichTextBox
    Friend WithEvents CadolProg As System.Windows.Forms.RichTextBox
    Friend WithEvents StartButton As System.Windows.Forms.Button
    Friend WithEvents LibraryCombo As System.Windows.Forms.ComboBox
    Friend WithEvents LibraryLabel As System.Windows.Forms.Label
    Friend WithEvents TargetLabel As System.Windows.Forms.Label
    Friend WithEvents SourceLabel As System.Windows.Forms.Label
    Friend WithEvents VolumeCombo As System.Windows.Forms.ComboBox
    Friend WithEvents VolumeLabel As System.Windows.Forms.Label
    Friend WithEvents TargetPath As System.Windows.Forms.TextBox
    Friend WithEvents SourcePath As System.Windows.Forms.TextBox
    Friend WithEvents StopCompiling As System.Windows.Forms.Button
    Friend WithEvents CheckChangedOnly As System.Windows.Forms.CheckBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CompileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CancelToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomizeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IndexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BrowseEnvironment As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker

End Class
