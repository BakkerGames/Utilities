<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemCompare = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddAppToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UsernameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.IgnoreSpacesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IgnoreVersionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemQuickCompareBinary = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemIncludeTestProj = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemIgnoreMissingDirectoryContents = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcludeFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExternalCompareProgramToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.StatusLabelCounts = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LabelFromDir = New System.Windows.Forms.Label()
        Me.ComboFromDir = New System.Windows.Forms.ComboBox()
        Me.ComboToDir = New System.Windows.Forms.ComboBox()
        Me.LabelToDir = New System.Windows.Forms.Label()
        Me.ButtonCompare = New System.Windows.Forms.Button()
        Me.TabControlMain = New System.Windows.Forms.TabControl()
        Me.TabPageDiff = New System.Windows.Forms.TabPage()
        Me.ListBoxDiff = New System.Windows.Forms.ListBox()
        Me.TabPageProjDiff = New System.Windows.Forms.TabPage()
        Me.ListBoxProjDiff = New System.Windows.Forms.ListBox()
        Me.TabPageFrom = New System.Windows.Forms.TabPage()
        Me.ListBoxFrom = New System.Windows.Forms.ListBox()
        Me.TabPageTo = New System.Windows.Forms.TabPage()
        Me.ListBoxTo = New System.Windows.Forms.ListBox()
        Me.ButtonCopyFromTo = New System.Windows.Forms.Button()
        Me.ButtonCopyToFrom = New System.Windows.Forms.Button()
        Me.ButtonShowDiffs = New System.Windows.Forms.Button()
        Me.ButtonDeleteFromOnly = New System.Windows.Forms.Button()
        Me.ButtonDeleteToOnly = New System.Windows.Forms.Button()
        Me.ComboApplication = New System.Windows.Forms.ComboBox()
        Me.LabelApplication = New System.Windows.Forms.Label()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.ButtonSelectAll = New System.Windows.Forms.Button()
        Me.ToolStripMenuItemIncludePackages = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        Me.TabControlMain.SuspendLayout()
        Me.TabPageDiff.SuspendLayout()
        Me.TabPageProjDiff.SuspendLayout()
        Me.TabPageFrom.SuspendLayout()
        Me.TabPageTo.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.OptionsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(792, 24)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemCompare, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ToolStripMenuItemCompare
        '
        Me.ToolStripMenuItemCompare.Name = "ToolStripMenuItemCompare"
        Me.ToolStripMenuItemCompare.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.ToolStripMenuItemCompare.Size = New System.Drawing.Size(142, 22)
        Me.ToolStripMenuItemCompare.Text = "Compare"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(139, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectAllToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.SelectAllToolStripMenuItem.Text = "Select &All"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddAppToolStripMenuItem, Me.UsernameToolStripMenuItem, Me.ToolStripSeparator3, Me.IgnoreSpacesToolStripMenuItem, Me.IgnoreVersionsToolStripMenuItem, Me.ToolStripMenuItemQuickCompareBinary, Me.ToolStripMenuItemIncludeTestProj, Me.ToolStripMenuItemIgnoreMissingDirectoryContents, Me.ToolStripMenuItemIncludePackages, Me.ExcludeFilesToolStripMenuItem, Me.ToolStripSeparator1, Me.ExternalCompareProgramToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'AddAppToolStripMenuItem
        '
        Me.AddAppToolStripMenuItem.Name = "AddAppToolStripMenuItem"
        Me.AddAppToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.AddAppToolStripMenuItem.Text = "&Add Application"
        '
        'UsernameToolStripMenuItem
        '
        Me.UsernameToolStripMenuItem.Name = "UsernameToolStripMenuItem"
        Me.UsernameToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.UsernameToolStripMenuItem.Text = "&Username"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(251, 6)
        '
        'IgnoreSpacesToolStripMenuItem
        '
        Me.IgnoreSpacesToolStripMenuItem.Name = "IgnoreSpacesToolStripMenuItem"
        Me.IgnoreSpacesToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.IgnoreSpacesToolStripMenuItem.Text = "Ignore &Spaces"
        '
        'IgnoreVersionsToolStripMenuItem
        '
        Me.IgnoreVersionsToolStripMenuItem.Name = "IgnoreVersionsToolStripMenuItem"
        Me.IgnoreVersionsToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.IgnoreVersionsToolStripMenuItem.Text = "Ignore &Versions"
        '
        'ToolStripMenuItemQuickCompareBinary
        '
        Me.ToolStripMenuItemQuickCompareBinary.Name = "ToolStripMenuItemQuickCompareBinary"
        Me.ToolStripMenuItemQuickCompareBinary.Size = New System.Drawing.Size(254, 22)
        Me.ToolStripMenuItemQuickCompareBinary.Text = "&Quick Compare Binary"
        '
        'ToolStripMenuItemIncludeTestProj
        '
        Me.ToolStripMenuItemIncludeTestProj.Name = "ToolStripMenuItemIncludeTestProj"
        Me.ToolStripMenuItemIncludeTestProj.Size = New System.Drawing.Size(254, 22)
        Me.ToolStripMenuItemIncludeTestProj.Text = "Include &Test Projects"
        '
        'ToolStripMenuItemIgnoreMissingDirectoryContents
        '
        Me.ToolStripMenuItemIgnoreMissingDirectoryContents.Name = "ToolStripMenuItemIgnoreMissingDirectoryContents"
        Me.ToolStripMenuItemIgnoreMissingDirectoryContents.Size = New System.Drawing.Size(254, 22)
        Me.ToolStripMenuItemIgnoreMissingDirectoryContents.Text = "Ignore Missing &Directory Contents"
        '
        'ExcludeFilesToolStripMenuItem
        '
        Me.ExcludeFilesToolStripMenuItem.Name = "ExcludeFilesToolStripMenuItem"
        Me.ExcludeFilesToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.ExcludeFilesToolStripMenuItem.Text = "E&xclude Files..."
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(251, 6)
        '
        'ExternalCompareProgramToolStripMenuItem
        '
        Me.ExternalCompareProgramToolStripMenuItem.Name = "ExternalCompareProgramToolStripMenuItem"
        Me.ExternalCompareProgramToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.ExternalCompareProgramToolStripMenuItem.Text = "&External Compare Program..."
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
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabelCounts})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 501)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(792, 22)
        Me.StatusStripMain.SizingGrip = False
        Me.StatusStripMain.TabIndex = 15
        '
        'StatusLabelCounts
        '
        Me.StatusLabelCounts.Name = "StatusLabelCounts"
        Me.StatusLabelCounts.Size = New System.Drawing.Size(105, 17)
        Me.StatusLabelCounts.Text = "StatusLabelCounts"
        '
        'LabelFromDir
        '
        Me.LabelFromDir.AutoSize = True
        Me.LabelFromDir.Location = New System.Drawing.Point(246, 34)
        Me.LabelFromDir.Name = "LabelFromDir"
        Me.LabelFromDir.Size = New System.Drawing.Size(49, 13)
        Me.LabelFromDir.TabIndex = 3
        Me.LabelFromDir.Text = "From Dir:"
        '
        'ComboFromDir
        '
        Me.ComboFromDir.FormattingEnabled = True
        Me.ComboFromDir.Location = New System.Drawing.Point(301, 31)
        Me.ComboFromDir.Name = "ComboFromDir"
        Me.ComboFromDir.Size = New System.Drawing.Size(159, 21)
        Me.ComboFromDir.Sorted = True
        Me.ComboFromDir.TabIndex = 4
        '
        'ComboToDir
        '
        Me.ComboToDir.FormattingEnabled = True
        Me.ComboToDir.Location = New System.Drawing.Point(511, 31)
        Me.ComboToDir.Name = "ComboToDir"
        Me.ComboToDir.Size = New System.Drawing.Size(159, 21)
        Me.ComboToDir.Sorted = True
        Me.ComboToDir.TabIndex = 6
        '
        'LabelToDir
        '
        Me.LabelToDir.AutoSize = True
        Me.LabelToDir.Location = New System.Drawing.Point(466, 34)
        Me.LabelToDir.Name = "LabelToDir"
        Me.LabelToDir.Size = New System.Drawing.Size(39, 13)
        Me.LabelToDir.TabIndex = 5
        Me.LabelToDir.Text = "To Dir:"
        '
        'ButtonCompare
        '
        Me.ButtonCompare.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonCompare.Enabled = False
        Me.ButtonCompare.Location = New System.Drawing.Point(705, 29)
        Me.ButtonCompare.Name = "ButtonCompare"
        Me.ButtonCompare.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCompare.TabIndex = 7
        Me.ButtonCompare.Text = "Compare"
        Me.ButtonCompare.UseVisualStyleBackColor = True
        '
        'TabControlMain
        '
        Me.TabControlMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControlMain.Controls.Add(Me.TabPageDiff)
        Me.TabControlMain.Controls.Add(Me.TabPageProjDiff)
        Me.TabControlMain.Controls.Add(Me.TabPageFrom)
        Me.TabControlMain.Controls.Add(Me.TabPageTo)
        Me.TabControlMain.Location = New System.Drawing.Point(12, 58)
        Me.TabControlMain.Name = "TabControlMain"
        Me.TabControlMain.SelectedIndex = 0
        Me.TabControlMain.Size = New System.Drawing.Size(768, 408)
        Me.TabControlMain.TabIndex = 8
        '
        'TabPageDiff
        '
        Me.TabPageDiff.Controls.Add(Me.ListBoxDiff)
        Me.TabPageDiff.Location = New System.Drawing.Point(4, 22)
        Me.TabPageDiff.Name = "TabPageDiff"
        Me.TabPageDiff.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageDiff.Size = New System.Drawing.Size(760, 382)
        Me.TabPageDiff.TabIndex = 0
        Me.TabPageDiff.Text = "Differences"
        Me.TabPageDiff.UseVisualStyleBackColor = True
        '
        'ListBoxDiff
        '
        Me.ListBoxDiff.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBoxDiff.FormattingEnabled = True
        Me.ListBoxDiff.IntegralHeight = False
        Me.ListBoxDiff.Location = New System.Drawing.Point(3, 3)
        Me.ListBoxDiff.Name = "ListBoxDiff"
        Me.ListBoxDiff.ScrollAlwaysVisible = True
        Me.ListBoxDiff.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxDiff.Size = New System.Drawing.Size(754, 376)
        Me.ListBoxDiff.Sorted = True
        Me.ListBoxDiff.TabIndex = 0
        '
        'TabPageProjDiff
        '
        Me.TabPageProjDiff.Controls.Add(Me.ListBoxProjDiff)
        Me.TabPageProjDiff.Location = New System.Drawing.Point(4, 22)
        Me.TabPageProjDiff.Name = "TabPageProjDiff"
        Me.TabPageProjDiff.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageProjDiff.Size = New System.Drawing.Size(760, 382)
        Me.TabPageProjDiff.TabIndex = 3
        Me.TabPageProjDiff.Text = "Project Differences"
        Me.TabPageProjDiff.UseVisualStyleBackColor = True
        '
        'ListBoxProjDiff
        '
        Me.ListBoxProjDiff.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBoxProjDiff.FormattingEnabled = True
        Me.ListBoxProjDiff.IntegralHeight = False
        Me.ListBoxProjDiff.Location = New System.Drawing.Point(3, 3)
        Me.ListBoxProjDiff.Name = "ListBoxProjDiff"
        Me.ListBoxProjDiff.ScrollAlwaysVisible = True
        Me.ListBoxProjDiff.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxProjDiff.Size = New System.Drawing.Size(754, 376)
        Me.ListBoxProjDiff.Sorted = True
        Me.ListBoxProjDiff.TabIndex = 1
        '
        'TabPageFrom
        '
        Me.TabPageFrom.Controls.Add(Me.ListBoxFrom)
        Me.TabPageFrom.Location = New System.Drawing.Point(4, 22)
        Me.TabPageFrom.Name = "TabPageFrom"
        Me.TabPageFrom.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageFrom.Size = New System.Drawing.Size(760, 382)
        Me.TabPageFrom.TabIndex = 1
        Me.TabPageFrom.Text = "In From Dir Only"
        Me.TabPageFrom.UseVisualStyleBackColor = True
        '
        'ListBoxFrom
        '
        Me.ListBoxFrom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBoxFrom.FormattingEnabled = True
        Me.ListBoxFrom.IntegralHeight = False
        Me.ListBoxFrom.Location = New System.Drawing.Point(3, 3)
        Me.ListBoxFrom.Name = "ListBoxFrom"
        Me.ListBoxFrom.ScrollAlwaysVisible = True
        Me.ListBoxFrom.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxFrom.Size = New System.Drawing.Size(754, 376)
        Me.ListBoxFrom.Sorted = True
        Me.ListBoxFrom.TabIndex = 1
        '
        'TabPageTo
        '
        Me.TabPageTo.Controls.Add(Me.ListBoxTo)
        Me.TabPageTo.Location = New System.Drawing.Point(4, 22)
        Me.TabPageTo.Name = "TabPageTo"
        Me.TabPageTo.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageTo.Size = New System.Drawing.Size(760, 382)
        Me.TabPageTo.TabIndex = 2
        Me.TabPageTo.Text = "In To Dir Only"
        Me.TabPageTo.UseVisualStyleBackColor = True
        '
        'ListBoxTo
        '
        Me.ListBoxTo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBoxTo.FormattingEnabled = True
        Me.ListBoxTo.IntegralHeight = False
        Me.ListBoxTo.Location = New System.Drawing.Point(3, 3)
        Me.ListBoxTo.Name = "ListBoxTo"
        Me.ListBoxTo.ScrollAlwaysVisible = True
        Me.ListBoxTo.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxTo.Size = New System.Drawing.Size(754, 376)
        Me.ListBoxTo.Sorted = True
        Me.ListBoxTo.TabIndex = 2
        '
        'ButtonCopyFromTo
        '
        Me.ButtonCopyFromTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ButtonCopyFromTo.Location = New System.Drawing.Point(12, 472)
        Me.ButtonCopyFromTo.Name = "ButtonCopyFromTo"
        Me.ButtonCopyFromTo.Size = New System.Drawing.Size(111, 23)
        Me.ButtonCopyFromTo.TabIndex = 9
        Me.ButtonCopyFromTo.Text = "Copy From -> To"
        Me.ButtonCopyFromTo.UseVisualStyleBackColor = True
        '
        'ButtonCopyToFrom
        '
        Me.ButtonCopyToFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonCopyToFrom.Location = New System.Drawing.Point(669, 472)
        Me.ButtonCopyToFrom.Name = "ButtonCopyToFrom"
        Me.ButtonCopyToFrom.Size = New System.Drawing.Size(111, 23)
        Me.ButtonCopyToFrom.TabIndex = 14
        Me.ButtonCopyToFrom.Text = "Copy From <- To"
        Me.ButtonCopyToFrom.UseVisualStyleBackColor = True
        '
        'ButtonShowDiffs
        '
        Me.ButtonShowDiffs.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonShowDiffs.Location = New System.Drawing.Point(300, 472)
        Me.ButtonShowDiffs.Name = "ButtonShowDiffs"
        Me.ButtonShowDiffs.Size = New System.Drawing.Size(111, 23)
        Me.ButtonShowDiffs.TabIndex = 11
        Me.ButtonShowDiffs.Text = "Show Differences"
        Me.ButtonShowDiffs.UseVisualStyleBackColor = True
        '
        'ButtonDeleteFromOnly
        '
        Me.ButtonDeleteFromOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ButtonDeleteFromOnly.Location = New System.Drawing.Point(129, 472)
        Me.ButtonDeleteFromOnly.Name = "ButtonDeleteFromOnly"
        Me.ButtonDeleteFromOnly.Size = New System.Drawing.Size(75, 23)
        Me.ButtonDeleteFromOnly.TabIndex = 10
        Me.ButtonDeleteFromOnly.Text = "Delete"
        Me.ButtonDeleteFromOnly.UseVisualStyleBackColor = True
        '
        'ButtonDeleteToOnly
        '
        Me.ButtonDeleteToOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonDeleteToOnly.Location = New System.Drawing.Point(588, 472)
        Me.ButtonDeleteToOnly.Name = "ButtonDeleteToOnly"
        Me.ButtonDeleteToOnly.Size = New System.Drawing.Size(75, 23)
        Me.ButtonDeleteToOnly.TabIndex = 13
        Me.ButtonDeleteToOnly.Text = "Delete"
        Me.ButtonDeleteToOnly.UseVisualStyleBackColor = True
        '
        'ComboApplication
        '
        Me.ComboApplication.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboApplication.FormattingEnabled = True
        Me.ComboApplication.Location = New System.Drawing.Point(81, 31)
        Me.ComboApplication.Name = "ComboApplication"
        Me.ComboApplication.Size = New System.Drawing.Size(159, 21)
        Me.ComboApplication.Sorted = True
        Me.ComboApplication.TabIndex = 2
        '
        'LabelApplication
        '
        Me.LabelApplication.AutoSize = True
        Me.LabelApplication.Location = New System.Drawing.Point(13, 34)
        Me.LabelApplication.Name = "LabelApplication"
        Me.LabelApplication.Size = New System.Drawing.Size(62, 13)
        Me.LabelApplication.TabIndex = 1
        Me.LabelApplication.Text = "Application:"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonCancel.Location = New System.Drawing.Point(705, 29)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 8
        Me.ButtonCancel.TabStop = False
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        Me.ButtonCancel.Visible = False
        '
        'ButtonSelectAll
        '
        Me.ButtonSelectAll.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonSelectAll.Location = New System.Drawing.Point(417, 472)
        Me.ButtonSelectAll.Name = "ButtonSelectAll"
        Me.ButtonSelectAll.Size = New System.Drawing.Size(75, 23)
        Me.ButtonSelectAll.TabIndex = 12
        Me.ButtonSelectAll.Text = "Select All"
        Me.ButtonSelectAll.UseVisualStyleBackColor = True
        '
        'ToolStripMenuItemIncludePackages
        '
        Me.ToolStripMenuItemIncludePackages.Name = "ToolStripMenuItemIncludePackages"
        Me.ToolStripMenuItemIncludePackages.Size = New System.Drawing.Size(254, 22)
        Me.ToolStripMenuItemIncludePackages.Text = "Include Packages Directories"
        '
        'FormMain
        '
        Me.AcceptButton = Me.ButtonCompare
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 523)
        Me.Controls.Add(Me.ButtonSelectAll)
        Me.Controls.Add(Me.ComboApplication)
        Me.Controls.Add(Me.LabelApplication)
        Me.Controls.Add(Me.ButtonDeleteToOnly)
        Me.Controls.Add(Me.ButtonDeleteFromOnly)
        Me.Controls.Add(Me.ButtonShowDiffs)
        Me.Controls.Add(Me.ButtonCopyToFrom)
        Me.Controls.Add(Me.ButtonCopyFromTo)
        Me.Controls.Add(Me.TabControlMain)
        Me.Controls.Add(Me.ComboToDir)
        Me.Controls.Add(Me.LabelToDir)
        Me.Controls.Add(Me.ComboFromDir)
        Me.Controls.Add(Me.LabelFromDir)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.MenuStripMain)
        Me.Controls.Add(Me.ButtonCompare)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStripMain
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Source Manager"
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.TabControlMain.ResumeLayout(False)
        Me.TabPageDiff.ResumeLayout(False)
        Me.TabPageProjDiff.ResumeLayout(False)
        Me.TabPageFrom.ResumeLayout(False)
        Me.TabPageTo.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LabelFromDir As System.Windows.Forms.Label
    Friend WithEvents ComboFromDir As System.Windows.Forms.ComboBox
    Friend WithEvents ComboToDir As System.Windows.Forms.ComboBox
    Friend WithEvents LabelToDir As System.Windows.Forms.Label
    Friend WithEvents ButtonCompare As System.Windows.Forms.Button
    Friend WithEvents TabControlMain As System.Windows.Forms.TabControl
    Friend WithEvents TabPageDiff As System.Windows.Forms.TabPage
    Friend WithEvents TabPageFrom As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTo As System.Windows.Forms.TabPage
    Friend WithEvents ListBoxDiff As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxFrom As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxTo As System.Windows.Forms.ListBox
    Friend WithEvents StatusLabelCounts As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ButtonCopyFromTo As System.Windows.Forms.Button
    Friend WithEvents ButtonCopyToFrom As System.Windows.Forms.Button
    Friend WithEvents ButtonShowDiffs As System.Windows.Forms.Button
    Friend WithEvents ButtonDeleteFromOnly As System.Windows.Forms.Button
    Friend WithEvents ButtonDeleteToOnly As System.Windows.Forms.Button
    Friend WithEvents ComboApplication As System.Windows.Forms.ComboBox
    Friend WithEvents LabelApplication As System.Windows.Forms.Label
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents TabPageProjDiff As System.Windows.Forms.TabPage
    Friend WithEvents ListBoxProjDiff As System.Windows.Forms.ListBox
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsernameToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ButtonSelectAll As System.Windows.Forms.Button
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IgnoreSpacesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddAppToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IgnoreVersionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExternalCompareProgramToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItemQuickCompareBinary As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemCompare As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItemIncludeTestProj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemIgnoreMissingDirectoryContents As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ExcludeFilesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemIncludePackages As ToolStripMenuItem
End Class
