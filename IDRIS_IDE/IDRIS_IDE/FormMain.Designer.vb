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
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelMain = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintPreviewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UndoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RedoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.FindToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FindAgainToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FindAndReplaceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PanelEnvironment = New System.Windows.Forms.Panel()
        Me.ComboBoxLibrary = New System.Windows.Forms.ComboBox()
        Me.ComboBoxVolume = New System.Windows.Forms.ComboBox()
        Me.ComboBoxDevice = New System.Windows.Forms.ComboBox()
        Me.ComboBoxEnvironment = New System.Windows.Forms.ComboBox()
        Me.ComboBoxServer = New System.Windows.Forms.ComboBox()
        Me.PanelFileButtons = New System.Windows.Forms.Panel()
        Me.ButtonMoveRight = New System.Windows.Forms.Button()
        Me.ButtonMoveLeft = New System.Windows.Forms.Button()
        Me.TextBoxMain = New System.Windows.Forms.TextBox()
        Me.PanelSortOptions = New System.Windows.Forms.Panel()
        Me.RadioButtonByNumber = New System.Windows.Forms.RadioButton()
        Me.RadioButtonByName = New System.Windows.Forms.RadioButton()
        Me.ListBoxFiles = New System.Windows.Forms.ListBox()
        Me.ButtonSearchAll = New System.Windows.Forms.Button()
        Me.ButtonShowAll = New System.Windows.Forms.Button()
        Me.StatusStripMain.SuspendLayout()
        Me.MenuStripMain.SuspendLayout()
        Me.PanelEnvironment.SuspendLayout()
        Me.PanelFileButtons.SuspendLayout()
        Me.PanelSortOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelMain})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 520)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(844, 22)
        Me.StatusStripMain.TabIndex = 6
        Me.StatusStripMain.Text = "StatusStrip1"
        '
        'ToolStripStatusLabelMain
        '
        Me.ToolStripStatusLabelMain.Name = "ToolStripStatusLabelMain"
        Me.ToolStripStatusLabelMain.Size = New System.Drawing.Size(829, 17)
        Me.ToolStripStatusLabelMain.Spring = True
        Me.ToolStripStatusLabelMain.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(844, 24)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.toolStripSeparator, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.toolStripSeparator1, Me.PrintToolStripMenuItem, Me.PrintPreviewToolStripMenuItem, Me.toolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Image = CType(resources.GetObject("NewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.NewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Image = CType(resources.GetObject("OpenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.OpenToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.OpenToolStripMenuItem.Text = "&Open"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(143, 6)
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Image = CType(resources.GetObject("SaveToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.SaveToolStripMenuItem.Text = "&Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.SaveAsToolStripMenuItem.Text = "Save &As"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(143, 6)
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.PrintToolStripMenuItem.Text = "&Print"
        '
        'PrintPreviewToolStripMenuItem
        '
        Me.PrintPreviewToolStripMenuItem.Image = CType(resources.GetObject("PrintPreviewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintPreviewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintPreviewToolStripMenuItem.Name = "PrintPreviewToolStripMenuItem"
        Me.PrintPreviewToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.PrintPreviewToolStripMenuItem.Text = "Print Pre&view"
        '
        'toolStripSeparator2
        '
        Me.toolStripSeparator2.Name = "toolStripSeparator2"
        Me.toolStripSeparator2.Size = New System.Drawing.Size(143, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UndoToolStripMenuItem, Me.RedoToolStripMenuItem, Me.toolStripSeparator3, Me.CutToolStripMenuItem, Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem, Me.toolStripSeparator4, Me.SelectAllToolStripMenuItem, Me.ToolStripSeparator5, Me.FindToolStripMenuItem, Me.FindAgainToolStripMenuItem, Me.FindAndReplaceToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'UndoToolStripMenuItem
        '
        Me.UndoToolStripMenuItem.Name = "UndoToolStripMenuItem"
        Me.UndoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.UndoToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.UndoToolStripMenuItem.Text = "&Undo"
        '
        'RedoToolStripMenuItem
        '
        Me.RedoToolStripMenuItem.Name = "RedoToolStripMenuItem"
        Me.RedoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.RedoToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.RedoToolStripMenuItem.Text = "&Redo"
        '
        'toolStripSeparator3
        '
        Me.toolStripSeparator3.Name = "toolStripSeparator3"
        Me.toolStripSeparator3.Size = New System.Drawing.Size(204, 6)
        '
        'CutToolStripMenuItem
        '
        Me.CutToolStripMenuItem.Image = CType(resources.GetObject("CutToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CutToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripMenuItem.Name = "CutToolStripMenuItem"
        Me.CutToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.CutToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.CutToolStripMenuItem.Text = "Cu&t"
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Image = CType(resources.GetObject("CopyToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CopyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.CopyToolStripMenuItem.Text = "&Copy"
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Image = CType(resources.GetObject("PasteToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PasteToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.PasteToolStripMenuItem.Text = "&Paste"
        '
        'toolStripSeparator4
        '
        Me.toolStripSeparator4.Name = "toolStripSeparator4"
        Me.toolStripSeparator4.Size = New System.Drawing.Size(204, 6)
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.SelectAllToolStripMenuItem.Text = "Select &All"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(204, 6)
        '
        'FindToolStripMenuItem
        '
        Me.FindToolStripMenuItem.Name = "FindToolStripMenuItem"
        Me.FindToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.FindToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.FindToolStripMenuItem.Text = "&Find"
        '
        'FindAgainToolStripMenuItem
        '
        Me.FindAgainToolStripMenuItem.Name = "FindAgainToolStripMenuItem"
        Me.FindAgainToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.FindAgainToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.FindAgainToolStripMenuItem.Text = "Find &Again"
        '
        'FindAndReplaceToolStripMenuItem
        '
        Me.FindAndReplaceToolStripMenuItem.Name = "FindAndReplaceToolStripMenuItem"
        Me.FindAndReplaceToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
        Me.FindAndReplaceToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.FindAndReplaceToolStripMenuItem.Text = "Find and &Replace"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomizeToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
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
        'PanelEnvironment
        '
        Me.PanelEnvironment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.PanelEnvironment.Controls.Add(Me.ComboBoxLibrary)
        Me.PanelEnvironment.Controls.Add(Me.ComboBoxVolume)
        Me.PanelEnvironment.Controls.Add(Me.ComboBoxDevice)
        Me.PanelEnvironment.Controls.Add(Me.ComboBoxEnvironment)
        Me.PanelEnvironment.Controls.Add(Me.ComboBoxServer)
        Me.PanelEnvironment.Location = New System.Drawing.Point(0, 27)
        Me.PanelEnvironment.Name = "PanelEnvironment"
        Me.PanelEnvironment.Size = New System.Drawing.Size(644, 27)
        Me.PanelEnvironment.TabIndex = 1
        '
        'ComboBoxLibrary
        '
        Me.ComboBoxLibrary.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxLibrary.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxLibrary.FormattingEnabled = True
        Me.ComboBoxLibrary.Location = New System.Drawing.Point(516, 3)
        Me.ComboBoxLibrary.Name = "ComboBoxLibrary"
        Me.ComboBoxLibrary.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxLibrary.Sorted = True
        Me.ComboBoxLibrary.TabIndex = 4
        '
        'ComboBoxVolume
        '
        Me.ComboBoxVolume.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxVolume.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxVolume.FormattingEnabled = True
        Me.ComboBoxVolume.Location = New System.Drawing.Point(389, 3)
        Me.ComboBoxVolume.Name = "ComboBoxVolume"
        Me.ComboBoxVolume.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxVolume.Sorted = True
        Me.ComboBoxVolume.TabIndex = 3
        '
        'ComboBoxDevice
        '
        Me.ComboBoxDevice.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxDevice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxDevice.FormattingEnabled = True
        Me.ComboBoxDevice.Location = New System.Drawing.Point(262, 3)
        Me.ComboBoxDevice.Name = "ComboBoxDevice"
        Me.ComboBoxDevice.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxDevice.Sorted = True
        Me.ComboBoxDevice.TabIndex = 2
        '
        'ComboBoxEnvironment
        '
        Me.ComboBoxEnvironment.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxEnvironment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxEnvironment.FormattingEnabled = True
        Me.ComboBoxEnvironment.Location = New System.Drawing.Point(135, 3)
        Me.ComboBoxEnvironment.Name = "ComboBoxEnvironment"
        Me.ComboBoxEnvironment.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxEnvironment.Sorted = True
        Me.ComboBoxEnvironment.TabIndex = 1
        '
        'ComboBoxServer
        '
        Me.ComboBoxServer.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxServer.FormattingEnabled = True
        Me.ComboBoxServer.Location = New System.Drawing.Point(8, 3)
        Me.ComboBoxServer.Name = "ComboBoxServer"
        Me.ComboBoxServer.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxServer.TabIndex = 0
        '
        'PanelFileButtons
        '
        Me.PanelFileButtons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelFileButtons.BackColor = System.Drawing.SystemColors.Control
        Me.PanelFileButtons.Controls.Add(Me.ButtonMoveRight)
        Me.PanelFileButtons.Controls.Add(Me.ButtonMoveLeft)
        Me.PanelFileButtons.Location = New System.Drawing.Point(0, 54)
        Me.PanelFileButtons.Name = "PanelFileButtons"
        Me.PanelFileButtons.Size = New System.Drawing.Size(644, 23)
        Me.PanelFileButtons.TabIndex = 2
        '
        'ButtonMoveRight
        '
        Me.ButtonMoveRight.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonMoveRight.Enabled = False
        Me.ButtonMoveRight.Location = New System.Drawing.Point(621, 0)
        Me.ButtonMoveRight.Name = "ButtonMoveRight"
        Me.ButtonMoveRight.Size = New System.Drawing.Size(23, 23)
        Me.ButtonMoveRight.TabIndex = 1
        Me.ButtonMoveRight.Text = ">"
        Me.ButtonMoveRight.UseVisualStyleBackColor = True
        '
        'ButtonMoveLeft
        '
        Me.ButtonMoveLeft.Enabled = False
        Me.ButtonMoveLeft.Location = New System.Drawing.Point(0, 0)
        Me.ButtonMoveLeft.Name = "ButtonMoveLeft"
        Me.ButtonMoveLeft.Size = New System.Drawing.Size(23, 23)
        Me.ButtonMoveLeft.TabIndex = 0
        Me.ButtonMoveLeft.Text = "<"
        Me.ButtonMoveLeft.UseVisualStyleBackColor = True
        '
        'TextBoxMain
        '
        Me.TextBoxMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxMain.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxMain.Location = New System.Drawing.Point(0, 77)
        Me.TextBoxMain.Multiline = True
        Me.TextBoxMain.Name = "TextBoxMain"
        Me.TextBoxMain.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBoxMain.Size = New System.Drawing.Size(644, 443)
        Me.TextBoxMain.TabIndex = 3
        '
        'PanelSortOptions
        '
        Me.PanelSortOptions.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelSortOptions.Controls.Add(Me.RadioButtonByNumber)
        Me.PanelSortOptions.Controls.Add(Me.RadioButtonByName)
        Me.PanelSortOptions.Location = New System.Drawing.Point(644, 27)
        Me.PanelSortOptions.Name = "PanelSortOptions"
        Me.PanelSortOptions.Size = New System.Drawing.Size(200, 27)
        Me.PanelSortOptions.TabIndex = 4
        '
        'RadioButtonByNumber
        '
        Me.RadioButtonByNumber.AutoSize = True
        Me.RadioButtonByNumber.Location = New System.Drawing.Point(107, 4)
        Me.RadioButtonByNumber.Name = "RadioButtonByNumber"
        Me.RadioButtonByNumber.Size = New System.Drawing.Size(77, 17)
        Me.RadioButtonByNumber.TabIndex = 1
        Me.RadioButtonByNumber.TabStop = True
        Me.RadioButtonByNumber.Text = "By Number"
        Me.RadioButtonByNumber.UseVisualStyleBackColor = True
        '
        'RadioButtonByName
        '
        Me.RadioButtonByName.AutoSize = True
        Me.RadioButtonByName.Checked = True
        Me.RadioButtonByName.Location = New System.Drawing.Point(17, 4)
        Me.RadioButtonByName.Name = "RadioButtonByName"
        Me.RadioButtonByName.Size = New System.Drawing.Size(68, 17)
        Me.RadioButtonByName.TabIndex = 0
        Me.RadioButtonByName.TabStop = True
        Me.RadioButtonByName.Text = "By Name"
        Me.RadioButtonByName.UseVisualStyleBackColor = True
        '
        'ListBoxFiles
        '
        Me.ListBoxFiles.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBoxFiles.FormattingEnabled = True
        Me.ListBoxFiles.IntegralHeight = False
        Me.ListBoxFiles.ItemHeight = 16
        Me.ListBoxFiles.Location = New System.Drawing.Point(644, 77)
        Me.ListBoxFiles.Name = "ListBoxFiles"
        Me.ListBoxFiles.ScrollAlwaysVisible = True
        Me.ListBoxFiles.Size = New System.Drawing.Size(200, 443)
        Me.ListBoxFiles.Sorted = True
        Me.ListBoxFiles.TabIndex = 5
        '
        'ButtonSearchAll
        '
        Me.ButtonSearchAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonSearchAll.Location = New System.Drawing.Point(744, 54)
        Me.ButtonSearchAll.Name = "ButtonSearchAll"
        Me.ButtonSearchAll.Size = New System.Drawing.Size(100, 23)
        Me.ButtonSearchAll.TabIndex = 7
        Me.ButtonSearchAll.Text = "Search All"
        Me.ButtonSearchAll.UseVisualStyleBackColor = True
        '
        'ButtonShowAll
        '
        Me.ButtonShowAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonShowAll.Enabled = False
        Me.ButtonShowAll.Location = New System.Drawing.Point(644, 54)
        Me.ButtonShowAll.Name = "ButtonShowAll"
        Me.ButtonShowAll.Size = New System.Drawing.Size(100, 23)
        Me.ButtonShowAll.TabIndex = 8
        Me.ButtonShowAll.Text = "Show All"
        Me.ButtonShowAll.UseVisualStyleBackColor = True
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(844, 542)
        Me.Controls.Add(Me.ButtonShowAll)
        Me.Controls.Add(Me.ButtonSearchAll)
        Me.Controls.Add(Me.ListBoxFiles)
        Me.Controls.Add(Me.PanelSortOptions)
        Me.Controls.Add(Me.TextBoxMain)
        Me.Controls.Add(Me.PanelFileButtons)
        Me.Controls.Add(Me.PanelEnvironment)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.MenuStripMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStripMain
        Me.MinimumSize = New System.Drawing.Size(860, 580)
        Me.Name = "FormMain"
        Me.Text = "IDRIS_IDE"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.PanelEnvironment.ResumeLayout(False)
        Me.PanelFileButtons.ResumeLayout(False)
        Me.PanelSortOptions.ResumeLayout(False)
        Me.PanelSortOptions.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintPreviewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UndoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RedoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SelectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomizeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PanelEnvironment As System.Windows.Forms.Panel
    Friend WithEvents ComboBoxLibrary As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxVolume As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxDevice As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxEnvironment As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxServer As System.Windows.Forms.ComboBox
    Friend WithEvents PanelFileButtons As System.Windows.Forms.Panel
    Friend WithEvents ButtonMoveRight As System.Windows.Forms.Button
    Friend WithEvents ButtonMoveLeft As System.Windows.Forms.Button
    Friend WithEvents TextBoxMain As System.Windows.Forms.TextBox
    Friend WithEvents ToolStripStatusLabelMain As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PanelSortOptions As System.Windows.Forms.Panel
    Friend WithEvents RadioButtonByNumber As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonByName As System.Windows.Forms.RadioButton
    Friend WithEvents ListBoxFiles As System.Windows.Forms.ListBox
    Friend WithEvents ButtonSearchAll As System.Windows.Forms.Button
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents FindToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FindAndReplaceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ButtonShowAll As System.Windows.Forms.Button
    Friend WithEvents FindAgainToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
