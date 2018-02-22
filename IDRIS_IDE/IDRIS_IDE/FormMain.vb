' --------------------------------
' --- FormMain.vb - 02/22/2018 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 02/22/2018 - SBakker
'            - Fixed out-of-bounds with CurrentSearchPos.
' 01/22/2018 - SBakker
'            - Added PC and AltPC paths.
' 01/05/2018 - SBakker
'            - Changed server paths to match current servers.
'            - Gracefully handle missing paths.
' 09/28/2017 - SBakker
'            - Switched to Arena.Common.Bootstrap.
' 05/29/2014 - SBakker
'            - Working on adding features to IDRIS_IDE.
'            - Made all filenames in list lowercase and a better font.
'            - Added "Search All" and "Show All" buttons.
'            - Added "Find" logic.
' 03/12/2012 - SBakker
'            - Added unchanged and changed Disk icons on the buttons.
' 02/03/2011 - SBakker
'            - Started working on IDRIS_IDE program in VB.NET.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports Arena.Common.Bootstrap

Public Class FormMain

#Region " Internal Variables "

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private FullPath_Server As String = ""
    Private FullPath_Environment As String = ""
    Private FullPath_Device As String = ""
    Private FullPath_Volume As String = ""
    Private FullPath_Library As String = ""

    Private IDRISFileList As List(Of IDRISFile) = Nothing
    Private CurrIDRISFile As IDRISFile = Nothing

    Private FileButtonList As New List(Of Button)
    Private NextButtonStart As Integer = 0
    Private WithEvents CurrFileButton As Button

    Private SearchString As String = ""
    Private ReplaceString As String = ""
    Private CurrentSearchPos As Integer = 0

#End Region

#Region " Form Routines "

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try

            Try
                If Bootstrapper.MustBootstrap Then
                    Me.Close()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
                Me.Close()
                Exit Sub
            End Try

            ' --- Update the settings from the last version ---
            If My.Settings.CallUpgrade Then
                My.Settings.Upgrade()
                My.Settings.CallUpgrade = False
                My.Settings.Save()
            End If

            Me.Show()
            Application.DoEvents()

            ' --- Set ByName or ByNumber first ---
            If My.Settings.ShowByName Then
                RadioButtonByName.Checked = True
            Else
                RadioButtonByNumber.Checked = True
            End If

            ' --- Fill Server List ---
            Dim ServerList() As String = My.Settings.ServerList.ToUpper.Split(";"c)
            Dim ServerPaths() As String = My.Settings.ServerPaths.Split(";"c)
            If ServerList.GetUpperBound(0) <> ServerPaths.GetUpperBound(0) Then
                MessageBox.Show("Number of configured Servers and ServerPaths don't match", My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End If
            For ServerIndex As Integer = 0 To ServerList.GetUpperBound(0)
                Dim ServerName As String = ServerList(ServerIndex)
                ComboBoxServer.Items.Add(ServerName)
            Next
            If My.Settings.LastServer <> "" Then
                For TempIndex As Integer = 0 To ComboBoxServer.Items.Count - 1
                    If CStr(ComboBoxServer.Items(TempIndex)) = My.Settings.LastServer Then
                        ComboBoxServer.SelectedIndex = TempIndex
                        Exit For
                    End If
                Next
            Else
                If ComboBoxServer.Items.Count = 1 Then
                    ComboBoxServer.SelectedIndex = 0
                End If
            End If

        Catch ex As Exception

            MessageBox.Show("Error starting " + My.Application.Info.AssemblyName +
                vbCrLf + vbCrLf + ex.Message,
                Me.Text, MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            Me.Close()

        End Try

    End Sub

#End Region

#Region " ComboBox SelectedIndex Changed "

    Private Sub ComboBoxServer_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxServer.SelectedIndexChanged
        ComboBoxEnvironment.Items.Clear()
        ComboBoxDevice.Items.Clear()
        ComboBoxVolume.Items.Clear()
        ComboBoxLibrary.Items.Clear()
        ListBoxFiles.Items.Clear()
        FullPath_Environment = ""
        FullPath_Device = ""
        FullPath_Volume = ""
        FullPath_Library = ""
        ToolStripStatusLabelMain.Text = ""
        If ComboBoxServer.SelectedIndex < 0 Then Exit Sub
        If ComboBoxServer.Items.Count = 0 Then Exit Sub
        Dim ServerPaths() As String = My.Settings.ServerPaths.Split(";"c)
        If CStr(ComboBoxServer.SelectedItem) = "LOCAL" AndAlso Directory.Exists(My.Settings.AltLocalPath) Then
            ' --- Everything is fine ---
        ElseIf CStr(ComboBoxServer.SelectedItem) = "PC" AndAlso Directory.Exists(My.Settings.AltPCPath) Then
            ' --- Everything is fine ---
        Else
            Application.DoEvents()
            Cursor.Current = Cursors.WaitCursor
            If Not Directory.Exists(ServerPaths(ComboBoxServer.SelectedIndex)) Then
                Cursor.Current = Cursors.Default
                MessageBox.Show("Directory not found: " + ServerPaths(ComboBoxServer.SelectedIndex), My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Cursor.Current = Cursors.Default
        End If
        FullPath_Server = ServerPaths(ComboBoxServer.SelectedIndex)
        My.Settings.LastServer = CStr(ComboBoxServer.SelectedItem)
        My.Settings.Save()
        If CStr(ComboBoxServer.SelectedItem) = "PROD" Then
            ComboBoxEnvironment.Items.Add("PROD")
            Exit Sub
        End If
        Dim DirList() As String = Nothing
        If CStr(ComboBoxServer.SelectedItem) = "LOCAL" AndAlso Directory.Exists(My.Settings.AltLocalPath) Then
            DirList = Directory.GetDirectories(My.Settings.AltLocalPath)
            FullPath_Server = My.Settings.AltLocalPath
        ElseIf CStr(ComboBoxServer.SelectedItem) = "PC" AndAlso Directory.Exists(My.Settings.AltPCPath) Then
            DirList = Directory.GetDirectories(My.Settings.AltPCPath)
            FullPath_Server = My.Settings.AltPCPath
        Else
            DirList = Directory.GetDirectories(ServerPaths(ComboBoxServer.SelectedIndex))
        End If
        For Each TempDir As String In DirList
            If Directory.Exists(TempDir) Then
                TempDir = TempDir.Substring(TempDir.LastIndexOf("\"c) + 1, TempDir.Length - (TempDir.LastIndexOf("\"c) + 1)).ToUpper
                ComboBoxEnvironment.Items.Add(TempDir)
            End If
        Next
        ToolStripStatusLabelMain.Text = FullPath_Server
        If My.Settings.LastEnvironment <> "" Then
            For TempIndex As Integer = 0 To ComboBoxEnvironment.Items.Count - 1
                If CStr(ComboBoxEnvironment.Items(TempIndex)) = My.Settings.LastEnvironment Then
                    ComboBoxEnvironment.SelectedIndex = TempIndex
                    Exit For
                End If
            Next
        Else
            If ComboBoxEnvironment.Items.Count = 1 Then
                ComboBoxEnvironment.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub ComboBoxEnvironment_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxEnvironment.SelectedIndexChanged
        ComboBoxDevice.Items.Clear()
        FullPath_Device = ""
        If CStr(ComboBoxServer.SelectedItem) = "PROD" Then
            FullPath_Environment = "SOURCE\CADOL"
        Else
            FullPath_Environment = CStr(ComboBoxEnvironment.SelectedItem) + "\CADOLSRC"
        End If
        If Not Directory.Exists(FullPath_Server + "\" + FullPath_Environment) Then
            ComboBoxVolume.Items.Clear()
            ComboBoxLibrary.Items.Clear()
            Exit Sub
        End If
        My.Settings.LastEnvironment = CStr(ComboBoxEnvironment.SelectedItem)
        My.Settings.Save()
        Dim DirList() As String = Directory.GetDirectories(FullPath_Server + "\" + FullPath_Environment)
        For Each TempDir As String In DirList
            TempDir = TempDir.Substring(TempDir.LastIndexOf("\"c) + 1, TempDir.Length - (TempDir.LastIndexOf("\"c) + 1)).ToUpper
            ComboBoxDevice.Items.Add(TempDir)
        Next
        ToolStripStatusLabelMain.Text = FullPath_Server + "\" + FullPath_Environment
        If My.Settings.LastDevice <> "" Then
            For TempIndex As Integer = 0 To ComboBoxDevice.Items.Count - 1
                If CStr(ComboBoxDevice.Items(TempIndex)) = My.Settings.LastDevice Then
                    ComboBoxDevice.SelectedIndex = TempIndex
                    Exit For
                End If
            Next
        Else
            If ComboBoxDevice.Items.Count = 1 Then
                ComboBoxDevice.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub ComboBoxDevice_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxDevice.SelectedIndexChanged
        ComboBoxVolume.Items.Clear()
        FullPath_Volume = ""
        FullPath_Device = CStr(ComboBoxDevice.SelectedItem)
        My.Settings.LastDevice = CStr(ComboBoxDevice.SelectedItem)
        My.Settings.Save()
        Dim DirList() As String = Directory.GetDirectories(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device)
        For Each TempDir As String In DirList
            TempDir = TempDir.Substring(TempDir.LastIndexOf("\"c) + 1, TempDir.Length - (TempDir.LastIndexOf("\"c) + 1)).ToUpper
            ComboBoxVolume.Items.Add(TempDir)
        Next
        ToolStripStatusLabelMain.Text = FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device
        If My.Settings.LastVolume <> "" Then
            For TempIndex As Integer = 0 To ComboBoxVolume.Items.Count - 1
                If CStr(ComboBoxVolume.Items(TempIndex)) = My.Settings.LastVolume Then
                    ComboBoxVolume.SelectedIndex = TempIndex
                    Exit For
                End If
            Next
        Else
            If ComboBoxVolume.Items.Count = 1 Then
                ComboBoxVolume.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub ComboBoxVolume_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxVolume.SelectedIndexChanged
        ComboBoxLibrary.Items.Clear()
        FullPath_Library = ""
        If ComboBoxVolume.SelectedIndex < 0 Then
            FullPath_Volume = ""
            Exit Sub
        End If
        FullPath_Volume = CStr(ComboBoxVolume.SelectedItem)
        My.Settings.LastVolume = CStr(ComboBoxVolume.SelectedItem)
        My.Settings.Save()
        Dim DirList() As String = Directory.GetDirectories(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume)
        For Each TempDir As String In DirList
            TempDir = TempDir.Substring(TempDir.LastIndexOf("\"c) + 1, TempDir.Length - (TempDir.LastIndexOf("\"c) + 1)).ToUpper
            If TempDir = "_IDRISYS" Then Continue For
            ComboBoxLibrary.Items.Add(TempDir)
        Next
        ToolStripStatusLabelMain.Text = FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume
        If My.Settings.LastLibrary <> "" Then
            For TempIndex As Integer = 0 To ComboBoxLibrary.Items.Count - 1
                If CStr(ComboBoxLibrary.Items(TempIndex)) = My.Settings.LastLibrary Then
                    ComboBoxLibrary.SelectedIndex = TempIndex
                    Exit For
                End If
            Next
        Else
            If ComboBoxLibrary.Items.Count = 1 Then
                ComboBoxLibrary.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub ComboBoxLibrary_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxLibrary.SelectedIndexChanged
        FullPath_Library = CStr(ComboBoxLibrary.SelectedItem)
        My.Settings.LastLibrary = CStr(ComboBoxLibrary.SelectedItem)
        My.Settings.Save()
        IDRISFileList = New List(Of IDRISFile)
        CurrIDRISFile = Nothing
        ClearButtonList()
        FillListBoxFiles()
        ToolStripStatusLabelMain.Text = FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library
    End Sub

#End Region

#Region " Main Menu Routines "

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
    End Sub

#End Region

    Private Sub ListBoxFiles_DoubleClick(sender As Object, e As System.EventArgs) Handles ListBoxFiles.DoubleClick

        If ListBoxFiles.SelectedIndex < 0 Then Exit Sub

        Dim TempFilename As String = CStr(ListBoxFiles.SelectedItem)
        If TempFilename.Substring(3, 3) = " - " Then
            TempFilename = TempFilename.Substring(6)
        End If

        CurrentSearchPos = 0

        If HandleFileSwitching(TempFilename) Then
            Exit Sub
        End If

        ' --- Create new file ---
        CurrIDRISFile = New IDRISFile
        With CurrIDRISFile
            .FullPath = FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library
            .FileName = TempFilename
            .FileText = File.ReadAllText(.FullPath + "\" + .FileName)
            TextBoxMain.Text = .FileText
            TextBoxMain.SelectionStart = .SelectionStart
            TextBoxMain.SelectionLength = 0
            .Changed = False ' Must come last, after loading TextBoxMain
        End With

        ' --- Add this into the list of open files ---
        IDRISFileList.Add(CurrIDRISFile)

        ' --- Show and store the new button for this file ---
        CurrFileButton = New Button
        With CurrFileButton
            .AutoSize = True
            .TextImageRelation = TextImageRelation.ImageBeforeText
            .Text = TempFilename
            .ForeColor = Color.Black
            .Image = My.Resources.DisabledDisk
            .Left = NextButtonStart
        End With
        PanelFileButtons.Controls.Add(CurrFileButton)
        FileButtonList.Add(CurrFileButton)
        NextButtonStart += CurrFileButton.Width
        AddHandler CurrFileButton.Click, AddressOf FileButton_Click

        ' --- Make sure the text box is editable ---
        'If TextBoxMain.ReadOnly Then
        '    TextBoxMain.ReadOnly = False
        'End If

        TextBoxMain.Focus()
        TextBoxMain.SelectionStart = 0
        TextBoxMain.SelectionLength = 0
        TextBoxMain.ScrollToCaret()

    End Sub

    Private Sub TextBoxMain_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBoxMain.TextChanged
        If CurrIDRISFile Is Nothing Then Exit Sub
        CurrIDRISFile.Changed = True
        If CurrFileButton Is Nothing Then Exit Sub
        CurrFileButton.Image = My.Resources.EnabledDisk
    End Sub

    Private Sub FileButton_Click(sender As System.Object, e As System.EventArgs)

        CurrFileButton = CType(sender, Button)
        Dim TempFilename As String = CurrFileButton.Text

        If HandleFileSwitching(TempFilename) Then
            Exit Sub
        End If

        Throw New SystemException("Error re-displaying file: " + TempFilename)

    End Sub

    Private Function HandleFileSwitching(ByVal TempFileName As String) As Boolean

        ' --- Save text from TextBoxMain back into record ---
        If CurrIDRISFile IsNot Nothing Then
            ' --- Check if opening the same file already open ---
            If CurrIDRISFile.FileName = TempFileName Then
                Return True
            End If
            CurrIDRISFile.SelectionStart = TextBoxMain.SelectionStart
            If CurrIDRISFile.Changed Then
                If CurrIDRISFile.FileText = TextBoxMain.Text Then
                    ' --- Changed and then undid change ---
                    CurrIDRISFile.Changed = False
                Else
                    CurrIDRISFile.FileText = TextBoxMain.Text
                End If
            End If
        End If

        CurrFileButton = Nothing

        ' --- See if file is already in the list ---
        For Each CurrIDRISFile In IDRISFileList
            If CurrIDRISFile.FileName = TempFileName Then
                ' --- Load the text into TextBoxmain ---
                TextBoxMain.Text = CurrIDRISFile.FileText
                TextBoxMain.Focus()
                TextBoxMain.SelectionStart = CurrIDRISFile.SelectionStart
                TextBoxMain.SelectionLength = 0
                TextBoxMain.ScrollToCaret()
                For Each CurrObj As Control In PanelFileButtons.Controls
                    If TypeOf CurrObj Is Button Then
                        If CType(CurrObj, Button).Text = TempFileName Then
                            CurrFileButton = CType(CurrObj, Button)
                            Exit For
                        End If
                    End If
                Next
                Return True
            End If
        Next

        Return False

    End Function

    Private Sub RadioButtonByName_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButtonByName.CheckedChanged
        If ComboBoxLibrary.Items Is Nothing OrElse ComboBoxLibrary.Items.Count = 0 Then
            Exit Sub
        End If
        If Not RadioButtonByName.Checked Then Exit Sub
        My.Settings.ShowByName = True
        My.Settings.Save()
        FillListBoxFiles()
    End Sub

    Private Sub RadioButtonByNumber_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButtonByNumber.CheckedChanged
        If ComboBoxLibrary.Items Is Nothing OrElse ComboBoxLibrary.Items.Count = 0 Then
            Exit Sub
        End If
        If Not RadioButtonByNumber.Checked Then Exit Sub
        My.Settings.ShowByName = False
        My.Settings.Save()
        FillListBoxFiles()
    End Sub

    Private Sub FillListBoxFiles()
        ListBoxFiles.Items.Clear()
        Dim FileList() As String = Directory.GetFiles(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library)
        For Each TempFile As String In FileList
            TempFile = TempFile.Substring(TempFile.LastIndexOf("\"c) + 1, TempFile.Length - (TempFile.LastIndexOf("\"c) + 1)).ToLower
            ' --- Only include valid Cadol files ---
            If TempFile.EndsWith(".k") Then
                If RadioButtonByNumber.Checked Then
                    Dim ProgNum As Integer = GetProgramNumber(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library + "\" + TempFile)
                    ListBoxFiles.Items.Add(ProgNum.ToString("000") + " - " + TempFile)
                Else
                    ListBoxFiles.Items.Add(TempFile)
                End If
            ElseIf TempFile.EndsWith(".i") OrElse TempFile.EndsWith(".fmt") Then
                ListBoxFiles.Items.Add(TempFile)
            End If
        Next
        ButtonShowAll.Enabled = False
    End Sub

    Private Sub ClearButtonList()
        TextBoxMain.Text = ""
        TextBoxMain.ReadOnly = True
        ButtonMoveLeft.Enabled = False
        ButtonMoveRight.Enabled = False
        For TempIndex As Integer = PanelFileButtons.Controls.Count - 1 To 2 Step -1
            PanelFileButtons.Controls.RemoveAt(TempIndex)
        Next
        FileButtonList.Clear()
        NextButtonStart = ButtonMoveLeft.Width
    End Sub

    Private Function GetProgramNumber(ByVal FilePath As String) As Integer
        Dim FileLines() As String = File.ReadAllLines(FilePath)
        Dim EndLineNum As Integer = FileLines.GetUpperBound(0)
        Do While Not FileLines(EndLineNum).ToUpper.StartsWith("END ")
            EndLineNum -= 1
            If EndLineNum < 0 Then Return 999
        Loop
        Return CInt(FileLines(EndLineNum).Substring(4).Trim)
    End Function

    Private Sub ButtonSearchAll_Click(sender As Object, e As EventArgs) Handles ButtonSearchAll.Click
        Dim FullFileText As String
        Dim TempFile As String
        ' -----------------------------
        If Clipboard.ContainsText Then
            SearchString = Clipboard.GetText
        End If
        SearchString = InputBox("Enter text for search", "Search All", SearchString)
        If String.IsNullOrWhiteSpace(SearchString) Then Exit Sub
        SearchString = SearchString.ToLower
        ' --- Find SearchString in files ---
        ListBoxFiles.Items.Clear()
        Dim FileList() As String = Directory.GetFiles(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library)
        For Each FullFilename As String In FileList
            TempFile = FullFilename.Substring(FullFilename.LastIndexOf("\"c) + 1, FullFilename.Length - (FullFilename.LastIndexOf("\"c) + 1)).ToLower
            ' --- Only include valid Cadol files ---
            If TempFile.EndsWith(".k") OrElse TempFile.EndsWith(".i") OrElse TempFile.EndsWith(".fmt") Then
                FullFileText = File.ReadAllText(FullFilename).ToLower
                If Not FullFileText.Contains(SearchString) Then
                    Continue For
                End If
            End If
            If TempFile.EndsWith(".k") Then
                If RadioButtonByNumber.Checked Then
                    Dim ProgNum As Integer = GetProgramNumber(FullPath_Server + "\" + FullPath_Environment + "\" + FullPath_Device + "\" + FullPath_Volume + "\" + FullPath_Library + "\" + TempFile)
                    ListBoxFiles.Items.Add(ProgNum.ToString("000") + " - " + TempFile)
                Else
                    ListBoxFiles.Items.Add(TempFile)
                End If
            ElseIf TempFile.EndsWith(".i") OrElse TempFile.EndsWith(".fmt") Then
                ListBoxFiles.Items.Add(TempFile)
            End If
        Next
        ButtonShowAll.Enabled = True
    End Sub

    Private Sub ButtonShowAll_Click(sender As Object, e As EventArgs) Handles ButtonShowAll.Click
        FillListBoxFiles()
    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Dim TempPos As Integer
        ' --------------------
        If Clipboard.ContainsText Then
            SearchString = Clipboard.GetText
        End If
        SearchString = InputBox("Enter text to find", "Find", SearchString)
        If String.IsNullOrWhiteSpace(SearchString) Then Exit Sub
        SearchString = SearchString.ToLower
        Clipboard.SetText(SearchString)
        ' --- Look for text in current file ---
        CurrentSearchPos = TextBoxMain.SelectionStart + TextBoxMain.SelectionLength
        TempPos = TextBoxMain.Text.ToLower.IndexOf(SearchString, CurrentSearchPos)
        If TempPos < 0 Then
            TempPos = TextBoxMain.Text.ToLower.IndexOf(SearchString)
        End If
        If TempPos >= 0 Then
            TextBoxMain.Focus()
            TextBoxMain.SelectionStart = TempPos
            TextBoxMain.SelectionLength = SearchString.Length
            TextBoxMain.ScrollToCaret()
            CurrentSearchPos = TempPos + SearchString.Length
        Else
            MessageBox.Show("Text not found: " + SearchString, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub FindAgainToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindAgainToolStripMenuItem.Click
        Dim TempPos As Integer
        ' --------------------
        If CurrentSearchPos >= TextBoxMain.Text.Length Then
            CurrentSearchPos = 0
        End If
        TempPos = TextBoxMain.Text.ToLower.IndexOf(SearchString, CurrentSearchPos)
        If TempPos < 0 Then
            TempPos = TextBoxMain.Text.ToLower.IndexOf(SearchString)
        End If
        If TempPos >= 0 Then
            TextBoxMain.Focus()
            TextBoxMain.SelectionStart = TempPos
            TextBoxMain.SelectionLength = SearchString.Length
            TextBoxMain.ScrollToCaret()
            CurrentSearchPos = TempPos + SearchString.Length
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        TextBoxMain.SelectionStart = 0
        TextBoxMain.SelectionLength = TextBoxMain.Text.Length
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If TextBoxMain.SelectionLength > 0 Then
            Clipboard.Clear()
            Clipboard.SetText(TextBoxMain.SelectedText)
        End If
    End Sub

End Class
