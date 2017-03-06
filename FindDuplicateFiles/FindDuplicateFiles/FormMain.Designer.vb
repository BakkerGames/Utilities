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
        Me.GroupBoxMain = New System.Windows.Forms.GroupBox()
        Me.ButtonSearch = New System.Windows.Forms.Button()
        Me.RadioButtonMatchingFilenames = New System.Windows.Forms.RadioButton()
        Me.RadioButtonMatchingContents = New System.Windows.Forms.RadioButton()
        Me.LabelSearchDirectory = New System.Windows.Forms.Label()
        Me.TextBoxSearchDirectory = New System.Windows.Forms.TextBox()
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelMain = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DataGridViewMain = New System.Windows.Forms.DataGridView()
        Me.DGVMain_MD5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DGVMain_Filename = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBoxMain.SuspendLayout()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        CType(Me.DataGridViewMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBoxMain
        '
        Me.GroupBoxMain.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBoxMain.Controls.Add(Me.ButtonSearch)
        Me.GroupBoxMain.Controls.Add(Me.RadioButtonMatchingFilenames)
        Me.GroupBoxMain.Controls.Add(Me.RadioButtonMatchingContents)
        Me.GroupBoxMain.Controls.Add(Me.LabelSearchDirectory)
        Me.GroupBoxMain.Controls.Add(Me.TextBoxSearchDirectory)
        Me.GroupBoxMain.Location = New System.Drawing.Point(60, 27)
        Me.GroupBoxMain.Name = "GroupBoxMain"
        Me.GroupBoxMain.Size = New System.Drawing.Size(528, 76)
        Me.GroupBoxMain.TabIndex = 1
        Me.GroupBoxMain.TabStop = False
        '
        'ButtonSearch
        '
        Me.ButtonSearch.Location = New System.Drawing.Point(440, 45)
        Me.ButtonSearch.Name = "ButtonSearch"
        Me.ButtonSearch.Size = New System.Drawing.Size(75, 23)
        Me.ButtonSearch.TabIndex = 4
        Me.ButtonSearch.Text = "Search"
        Me.ButtonSearch.UseVisualStyleBackColor = True
        '
        'RadioButtonMatchingFilenames
        '
        Me.RadioButtonMatchingFilenames.AutoSize = True
        Me.RadioButtonMatchingFilenames.Location = New System.Drawing.Point(221, 48)
        Me.RadioButtonMatchingFilenames.Name = "RadioButtonMatchingFilenames"
        Me.RadioButtonMatchingFilenames.Size = New System.Drawing.Size(119, 17)
        Me.RadioButtonMatchingFilenames.TabIndex = 3
        Me.RadioButtonMatchingFilenames.Text = "Matching Filenames"
        Me.RadioButtonMatchingFilenames.UseVisualStyleBackColor = True
        '
        'RadioButtonMatchingContents
        '
        Me.RadioButtonMatchingContents.AutoSize = True
        Me.RadioButtonMatchingContents.Checked = True
        Me.RadioButtonMatchingContents.Location = New System.Drawing.Point(101, 48)
        Me.RadioButtonMatchingContents.Name = "RadioButtonMatchingContents"
        Me.RadioButtonMatchingContents.Size = New System.Drawing.Size(114, 17)
        Me.RadioButtonMatchingContents.TabIndex = 2
        Me.RadioButtonMatchingContents.TabStop = True
        Me.RadioButtonMatchingContents.Text = "Matching Contents"
        Me.RadioButtonMatchingContents.UseVisualStyleBackColor = True
        '
        'LabelSearchDirectory
        '
        Me.LabelSearchDirectory.AutoSize = True
        Me.LabelSearchDirectory.Location = New System.Drawing.Point(6, 23)
        Me.LabelSearchDirectory.Name = "LabelSearchDirectory"
        Me.LabelSearchDirectory.Size = New System.Drawing.Size(89, 13)
        Me.LabelSearchDirectory.TabIndex = 0
        Me.LabelSearchDirectory.Text = "Search Directory:"
        '
        'TextBoxSearchDirectory
        '
        Me.TextBoxSearchDirectory.AllowDrop = True
        Me.TextBoxSearchDirectory.Location = New System.Drawing.Point(101, 19)
        Me.TextBoxSearchDirectory.Name = "TextBoxSearchDirectory"
        Me.TextBoxSearchDirectory.Size = New System.Drawing.Size(414, 20)
        Me.TextBoxSearchDirectory.TabIndex = 1
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(649, 24)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStrip1"
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
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelMain})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 377)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(649, 22)
        Me.StatusStripMain.TabIndex = 3
        Me.StatusStripMain.Text = "StatusStrip1"
        '
        'ToolStripStatusLabelMain
        '
        Me.ToolStripStatusLabelMain.Name = "ToolStripStatusLabelMain"
        Me.ToolStripStatusLabelMain.Size = New System.Drawing.Size(634, 17)
        Me.ToolStripStatusLabelMain.Spring = True
        Me.ToolStripStatusLabelMain.Text = "Searching..."
        '
        'DataGridViewMain
        '
        Me.DataGridViewMain.AllowUserToAddRows = False
        Me.DataGridViewMain.AllowUserToDeleteRows = False
        Me.DataGridViewMain.AllowUserToResizeColumns = False
        Me.DataGridViewMain.AllowUserToResizeRows = False
        Me.DataGridViewMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridViewMain.BackgroundColor = System.Drawing.SystemColors.Control
        Me.DataGridViewMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewMain.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DGVMain_MD5, Me.DGVMain_Filename})
        Me.DataGridViewMain.Location = New System.Drawing.Point(12, 109)
        Me.DataGridViewMain.Name = "DataGridViewMain"
        Me.DataGridViewMain.ReadOnly = True
        Me.DataGridViewMain.RowHeadersVisible = False
        Me.DataGridViewMain.Size = New System.Drawing.Size(625, 265)
        Me.DataGridViewMain.TabIndex = 2
        '
        'DGVMain_MD5
        '
        Me.DGVMain_MD5.HeaderText = "MD5"
        Me.DGVMain_MD5.MinimumWidth = 210
        Me.DGVMain_MD5.Name = "DGVMain_MD5"
        Me.DGVMain_MD5.ReadOnly = True
        Me.DGVMain_MD5.Width = 210
        '
        'DGVMain_Filename
        '
        Me.DGVMain_Filename.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DGVMain_Filename.HeaderText = "Filename"
        Me.DGVMain_Filename.Name = "DGVMain_Filename"
        Me.DGVMain_Filename.ReadOnly = True
        '
        'FormMain
        '
        Me.AcceptButton = Me.ButtonSearch
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 399)
        Me.Controls.Add(Me.DataGridViewMain)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.GroupBoxMain)
        Me.Controls.Add(Me.MenuStripMain)
        Me.MainMenuStrip = Me.MenuStripMain
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Duplicate Files"
        Me.GroupBoxMain.ResumeLayout(False)
        Me.GroupBoxMain.PerformLayout()
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        CType(Me.DataGridViewMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBoxMain As System.Windows.Forms.GroupBox
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabelMain As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ButtonSearch As System.Windows.Forms.Button
    Friend WithEvents RadioButtonMatchingFilenames As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonMatchingContents As System.Windows.Forms.RadioButton
    Friend WithEvents LabelSearchDirectory As System.Windows.Forms.Label
    Friend WithEvents TextBoxSearchDirectory As System.Windows.Forms.TextBox
    Friend WithEvents DataGridViewMain As System.Windows.Forms.DataGridView
    Friend WithEvents DGVMain_MD5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DGVMain_Filename As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
