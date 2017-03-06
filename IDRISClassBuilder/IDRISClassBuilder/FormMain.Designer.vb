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
        Me.TextFields = New System.Windows.Forms.TextBox()
        Me.TextOutput = New System.Windows.Forms.TextBox()
        Me.TextClassName = New System.Windows.Forms.TextBox()
        Me.TextDatabaseName = New System.Windows.Forms.TextBox()
        Me.TextInput = New System.Windows.Forms.TextBox()
        Me.TextFromPath = New System.Windows.Forms.TextBox()
        Me.TextToPath = New System.Windows.Forms.TextBox()
        Me.ButtonBuildAll = New System.Windows.Forms.Button()
        Me.TextConnName = New System.Windows.Forms.TextBox()
        Me.LabelConnName = New System.Windows.Forms.Label()
        Me.LabelDatabase = New System.Windows.Forms.Label()
        Me.TextDatabase = New System.Windows.Forms.TextBox()
        Me.LabelDate = New System.Windows.Forms.Label()
        Me.TextDate = New System.Windows.Forms.TextBox()
        Me.ButtonBuildHist = New System.Windows.Forms.Button()
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBoxApp = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripComboBoxDrive = New System.Windows.Forms.ToolStripComboBox()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelMain = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextFields
        '
        Me.TextFields.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextFields.Location = New System.Drawing.Point(11, 290)
        Me.TextFields.Multiline = True
        Me.TextFields.Name = "TextFields"
        Me.TextFields.ReadOnly = True
        Me.TextFields.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextFields.Size = New System.Drawing.Size(774, 160)
        Me.TextFields.TabIndex = 13
        '
        'TextOutput
        '
        Me.TextOutput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextOutput.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextOutput.Location = New System.Drawing.Point(11, 456)
        Me.TextOutput.Multiline = True
        Me.TextOutput.Name = "TextOutput"
        Me.TextOutput.ReadOnly = True
        Me.TextOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextOutput.Size = New System.Drawing.Size(774, 160)
        Me.TextOutput.TabIndex = 14
        Me.TextOutput.WordWrap = False
        '
        'TextClassName
        '
        Me.TextClassName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextClassName.Location = New System.Drawing.Point(401, 263)
        Me.TextClassName.Name = "TextClassName"
        Me.TextClassName.ReadOnly = True
        Me.TextClassName.Size = New System.Drawing.Size(384, 20)
        Me.TextClassName.TabIndex = 12
        '
        'TextDatabaseName
        '
        Me.TextDatabaseName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextDatabaseName.Location = New System.Drawing.Point(11, 263)
        Me.TextDatabaseName.Name = "TextDatabaseName"
        Me.TextDatabaseName.ReadOnly = True
        Me.TextDatabaseName.Size = New System.Drawing.Size(384, 20)
        Me.TextDatabaseName.TabIndex = 11
        '
        'TextInput
        '
        Me.TextInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextInput.Location = New System.Drawing.Point(12, 62)
        Me.TextInput.Multiline = True
        Me.TextInput.Name = "TextInput"
        Me.TextInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextInput.Size = New System.Drawing.Size(774, 166)
        Me.TextInput.TabIndex = 2
        '
        'TextFromPath
        '
        Me.TextFromPath.Location = New System.Drawing.Point(12, 36)
        Me.TextFromPath.Name = "TextFromPath"
        Me.TextFromPath.Size = New System.Drawing.Size(384, 20)
        Me.TextFromPath.TabIndex = 0
        '
        'TextToPath
        '
        Me.TextToPath.Location = New System.Drawing.Point(402, 36)
        Me.TextToPath.Name = "TextToPath"
        Me.TextToPath.Size = New System.Drawing.Size(384, 20)
        Me.TextToPath.TabIndex = 1
        '
        'ButtonBuildAll
        '
        Me.ButtonBuildAll.Enabled = False
        Me.ButtonBuildAll.Location = New System.Drawing.Point(527, 234)
        Me.ButtonBuildAll.Name = "ButtonBuildAll"
        Me.ButtonBuildAll.Size = New System.Drawing.Size(82, 23)
        Me.ButtonBuildAll.TabIndex = 9
        Me.ButtonBuildAll.Text = "Build All"
        Me.ButtonBuildAll.UseVisualStyleBackColor = True
        '
        'TextConnName
        '
        Me.TextConnName.Location = New System.Drawing.Point(418, 236)
        Me.TextConnName.Name = "TextConnName"
        Me.TextConnName.Size = New System.Drawing.Size(100, 20)
        Me.TextConnName.TabIndex = 8
        '
        'LabelConnName
        '
        Me.LabelConnName.AutoSize = True
        Me.LabelConnName.Location = New System.Drawing.Point(349, 239)
        Me.LabelConnName.Name = "LabelConnName"
        Me.LabelConnName.Size = New System.Drawing.Size(63, 13)
        Me.LabelConnName.TabIndex = 7
        Me.LabelConnName.Text = "ConnName:"
        '
        'LabelDatabase
        '
        Me.LabelDatabase.AutoSize = True
        Me.LabelDatabase.Location = New System.Drawing.Point(169, 239)
        Me.LabelDatabase.Name = "LabelDatabase"
        Me.LabelDatabase.Size = New System.Drawing.Size(56, 13)
        Me.LabelDatabase.TabIndex = 5
        Me.LabelDatabase.Text = "Database:"
        '
        'TextDatabase
        '
        Me.TextDatabase.Location = New System.Drawing.Point(231, 236)
        Me.TextDatabase.Name = "TextDatabase"
        Me.TextDatabase.Size = New System.Drawing.Size(100, 20)
        Me.TextDatabase.TabIndex = 6
        '
        'LabelDate
        '
        Me.LabelDate.AutoSize = True
        Me.LabelDate.Location = New System.Drawing.Point(10, 239)
        Me.LabelDate.Name = "LabelDate"
        Me.LabelDate.Size = New System.Drawing.Size(33, 13)
        Me.LabelDate.TabIndex = 3
        Me.LabelDate.Text = "Date:"
        '
        'TextDate
        '
        Me.TextDate.Location = New System.Drawing.Point(49, 236)
        Me.TextDate.Name = "TextDate"
        Me.TextDate.Size = New System.Drawing.Size(100, 20)
        Me.TextDate.TabIndex = 4
        '
        'ButtonBuildHist
        '
        Me.ButtonBuildHist.Enabled = False
        Me.ButtonBuildHist.Location = New System.Drawing.Point(703, 234)
        Me.ButtonBuildHist.Name = "ButtonBuildHist"
        Me.ButtonBuildHist.Size = New System.Drawing.Size(82, 23)
        Me.ButtonBuildHist.TabIndex = 10
        Me.ButtonBuildHist.Text = "Build History"
        Me.ButtonBuildHist.UseVisualStyleBackColor = True
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem, Me.ToolStripComboBoxApp, Me.ToolStripComboBoxDrive})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(797, 27)
        Me.MenuStripMain.TabIndex = 15
        Me.MenuStripMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 23)
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
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 23)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.AboutToolStripMenuItem.Text = "&About..."
        '
        'ToolStripComboBoxApp
        '
        Me.ToolStripComboBoxApp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBoxApp.Name = "ToolStripComboBoxApp"
        Me.ToolStripComboBoxApp.Size = New System.Drawing.Size(121, 23)
        '
        'ToolStripComboBoxDrive
        '
        Me.ToolStripComboBoxDrive.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBoxDrive.Items.AddRange(New Object() {"C:", "M:", "N:", "P:", "Y:"})
        Me.ToolStripComboBoxDrive.Name = "ToolStripComboBoxDrive"
        Me.ToolStripComboBoxDrive.Size = New System.Drawing.Size(121, 23)
        Me.ToolStripComboBoxDrive.Sorted = True
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelMain})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 628)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(797, 22)
        Me.StatusStripMain.TabIndex = 16
        Me.StatusStripMain.Text = "StatusStripMain"
        '
        'ToolStripStatusLabelMain
        '
        Me.ToolStripStatusLabelMain.Name = "ToolStripStatusLabelMain"
        Me.ToolStripStatusLabelMain.Size = New System.Drawing.Size(782, 17)
        Me.ToolStripStatusLabelMain.Spring = True
        Me.ToolStripStatusLabelMain.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(797, 650)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.ButtonBuildHist)
        Me.Controls.Add(Me.LabelDate)
        Me.Controls.Add(Me.TextDate)
        Me.Controls.Add(Me.LabelDatabase)
        Me.Controls.Add(Me.TextDatabase)
        Me.Controls.Add(Me.LabelConnName)
        Me.Controls.Add(Me.TextConnName)
        Me.Controls.Add(Me.ButtonBuildAll)
        Me.Controls.Add(Me.TextToPath)
        Me.Controls.Add(Me.TextFromPath)
        Me.Controls.Add(Me.TextInput)
        Me.Controls.Add(Me.TextDatabaseName)
        Me.Controls.Add(Me.TextClassName)
        Me.Controls.Add(Me.TextOutput)
        Me.Controls.Add(Me.TextFields)
        Me.Controls.Add(Me.MenuStripMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStripMain
        Me.MaximizeBox = False
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "IDRIS Class Builder"
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextFields As System.Windows.Forms.TextBox
    Friend WithEvents TextOutput As System.Windows.Forms.TextBox
    Friend WithEvents TextClassName As System.Windows.Forms.TextBox
    Friend WithEvents TextDatabaseName As System.Windows.Forms.TextBox
    Friend WithEvents TextInput As System.Windows.Forms.TextBox
    Friend WithEvents TextFromPath As System.Windows.Forms.TextBox
    Friend WithEvents TextToPath As System.Windows.Forms.TextBox
    Friend WithEvents ButtonBuildAll As System.Windows.Forms.Button
    Friend WithEvents TextConnName As System.Windows.Forms.TextBox
    Friend WithEvents LabelConnName As System.Windows.Forms.Label
    Friend WithEvents LabelDatabase As System.Windows.Forms.Label
    Friend WithEvents TextDatabase As System.Windows.Forms.TextBox
    Friend WithEvents LabelDate As System.Windows.Forms.Label
    Friend WithEvents TextDate As System.Windows.Forms.TextBox
    Friend WithEvents ButtonBuildHist As System.Windows.Forms.Button
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabelMain As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripComboBoxApp As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents ToolStripComboBoxDrive As System.Windows.Forms.ToolStripComboBox

End Class
