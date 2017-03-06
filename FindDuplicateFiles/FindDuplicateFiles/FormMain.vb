' --------------------------------
' --- FormMain.vb - 01/30/2015 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 01/30/2015 - SBakker
'            - Started FindDuplicateFiles project.
' ----------------------------------------------------------------------------------------------------

Imports MD5ClassLibrary.MD5Utilities
Imports System.IO

Public Class FormMain

#Region " Internal Variables "

    Private FileList As New ArrayList

    Private DirectoryCount As Integer
    Private FileCount As Integer

#End Region

#Region " Form Events "

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try

            ' --- Update the settings from the last version ---
            If My.Settings.CallUpgrade Then
                My.Settings.Upgrade()
                My.Settings.CallUpgrade = False
                My.Settings.Save()
            End If

            TextBoxSearchDirectory.Text = My.Settings.SearchDirectory
            If My.Settings.MatchingOption = 0 Then
                RadioButtonMatchingContents.Checked = True
            Else
                RadioButtonMatchingFilenames.Checked = True
            End If

            TextBoxSearchDirectory.Focus()

        Catch ex As Exception

            MessageBox.Show("Error starting " + My.Application.Info.AssemblyName + _
                vbCrLf + vbCrLf + ex.Message, _
                Me.Text, MessageBoxButtons.OK, _
                MessageBoxIcon.Error)
            Me.Close()

        End Try

    End Sub

#End Region

#Region " MainMenu Events "

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
    End Sub

#End Region

#Region " TextBox Events "

    Private Sub TextBoxSearchDirectory_DragEnter(sender As Object, e As DragEventArgs) Handles TextBoxSearchDirectory.DragEnter
        ' Check the format of the data being dropped.
        If e.Data.GetDataPresent(DataFormats.Text) OrElse e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' --- Display the copy cursor ---
            e.Effect = DragDropEffects.Copy
        Else
            ' --- Display the no-drop cursor ---
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBoxSearchDirectory_DragDrop(sender As Object, e As DragEventArgs) Handles TextBoxSearchDirectory.DragDrop
        If e.Data.GetDataPresent(DataFormats.Text) Then
            ' --- Paste the text ---
            TextBoxSearchDirectory.Text = CStr(e.Data.GetData(DataFormats.Text)).Replace("""", "")
            TextBoxSearchDirectory.SelectionStart = TextBoxSearchDirectory.Text.Length
        ElseIf e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' --- Get the list of filenames being dragged ---
            Dim MyFiles() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If MyFiles.Count > 1 Then
                TextBoxSearchDirectory.Text = ""
                Exit Sub
            End If
            ' --- Paste the filename ---
            TextBoxSearchDirectory.Text = MyFiles(0)
            TextBoxSearchDirectory.SelectionStart = TextBoxSearchDirectory.Text.Length
        End If
    End Sub

    Private Sub TextBoxSearchDirectory_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSearchDirectory.TextChanged
        If TextBoxSearchDirectory.Text.Contains("""") Then
            TextBoxSearchDirectory.Text = TextBoxSearchDirectory.Text.Replace("""", "")
            TextBoxSearchDirectory.SelectionStart = TextBoxSearchDirectory.Text.Length
        End If
    End Sub

#End Region

#Region " Button Events "

    Private Sub ButtonSearch_Click(sender As Object, e As EventArgs) Handles ButtonSearch.Click

        If String.IsNullOrWhiteSpace(TextBoxSearchDirectory.Text) Then Exit Sub
        If Not Directory.Exists(TextBoxSearchDirectory.Text) Then
            MessageBox.Show("Directory not found!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        My.Settings.SearchDirectory = TextBoxSearchDirectory.Text
        If RadioButtonMatchingContents.Checked Then
            My.Settings.MatchingOption = 0
        Else
            My.Settings.MatchingOption = 1
        End If
        My.Settings.Save()

        TextBoxSearchDirectory.Enabled = False
        RadioButtonMatchingContents.Enabled = False
        RadioButtonMatchingFilenames.Enabled = False
        ButtonSearch.Enabled = False

        SearchForMatchingFiles()

        TextBoxSearchDirectory.Enabled = True
        RadioButtonMatchingContents.Enabled = True
        RadioButtonMatchingFilenames.Enabled = True
        ButtonSearch.Enabled = True

    End Sub

#End Region

#Region " StatusLine Routines "

    Private Sub UpdateStatusLine(ByVal DoneFlag As Boolean)
        If DoneFlag AndAlso ToolStripStatusLabelMain.Text.EndsWith("...") Then
            ToolStripStatusLabelMain.Text += " Done"
        Else
            ToolStripStatusLabelMain.Text = "Directories: " + DirectoryCount.ToString +
                                            " - Files Found: " + FileCount.ToString
            If DoneFlag Then
                ToolStripStatusLabelMain.Text += " - Done"
            End If
        End If
    End Sub

#End Region

#Region " Internal Routines "

    Private Sub SearchForMatchingFiles()
        DataGridViewMain.Rows.Clear()
        FileList.Clear()
        DirectoryCount = 0
        FileCount = 0
        UpdateStatusLine(False)
        Application.DoEvents()
        FillFileListRecursive(TextBoxSearchDirectory.Text)
        FileList.Sort()
        ShowMatchingFiles()
        UpdateStatusLine(True)
    End Sub

    Private Sub FillFileListRecursive(ByVal Pathname As String)
        Dim FileMD5 As String
        Dim DirInfo As New DirectoryInfo(Pathname)
        ' ----------------------------------------
        DirectoryCount += 1
        UpdateStatusLine(False)
        Application.DoEvents()
        ' --- Go through all files first ---
        For Each CurrSourceFile As FileInfo In DirInfo.GetFiles
            FileCount += 1
            UpdateStatusLine(False)
            Application.DoEvents()
            If (CurrSourceFile.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            If IgnoreFiles(CurrSourceFile.FullName) Then Continue For
            ' --- Add file FullName and MD5 to FileList ---
            FileMD5 = CalcMD5(CurrSourceFile.FullName)
            FileList.Add(FileMD5 + vbTab + CurrSourceFile.FullName)
        Next
        ' --- Go through all subdirectories next ---
        For Each CurrSubDir As DirectoryInfo In DirInfo.GetDirectories
            If (CurrSubDir.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            If IgnoreDir(CurrSubDir.FullName) Then Continue For
            FillFileListRecursive(CurrSubDir.FullName)
        Next
    End Sub

    Private Sub ShowMatchingFiles()
        Dim CurrItem As String()
        Dim CurrMD5 As String
        Dim CurrFilename As String
        Dim LastMD5 As String = ""
        Dim LastFilename As String = ""
        Dim LastAddedFlag As Boolean = False
        Dim RowIndex As Integer = -1
        Dim GroupIndex As Integer = -1
        ' ----------------------------------
        For CurrIndex As Integer = 0 To FileList.Count - 1
            CurrItem = CStr(FileList(CurrIndex)).Split(CChar(vbTab))
            CurrMD5 = CurrItem(0)
            CurrFilename = CurrItem(1)
            If CurrMD5 = LastMD5 Then
                If Not LastAddedFlag Then
                    DataGridViewMain.Rows.Add()
                    RowIndex += 1
                    GroupIndex += 1
                    With DataGridViewMain.Rows(RowIndex)
                        If GroupIndex Mod 2 = 0 Then
                            .DefaultCellStyle.BackColor = Color.White
                        Else
                            .DefaultCellStyle.BackColor = Color.LightGray
                        End If
                        .Cells(DGVMain_MD5.Index).Value = LastMD5
                        .Cells(DGVMain_Filename.Index).Value = LastFilename
                    End With
                    LastAddedFlag = True
                End If
                DataGridViewMain.Rows.Add()
                RowIndex += 1
                With DataGridViewMain.Rows(RowIndex)
                    If GroupIndex Mod 2 = 0 Then
                        .DefaultCellStyle.BackColor = Color.White
                    Else
                        .DefaultCellStyle.BackColor = Color.LightGray
                    End If
                    .Cells(DGVMain_MD5.Index).Value = CurrMD5
                    .Cells(DGVMain_Filename.Index).Value = CurrFilename
                End With
            Else
                LastAddedFlag = False
            End If
            LastMD5 = CurrMD5
            LastFilename = CurrFilename
        Next
    End Sub

    Private Function IgnoreDir(ByVal DirPath As String) As Boolean
        ' --- Want to include the main Bin directory ---
        If DirPath.EndsWith("\Bin") Then Return False
        ' --- Check for others using lowercase ---
        DirPath = DirPath.ToLower
        If DirPath.EndsWith("\bin") Then Return True
        If DirPath.EndsWith("\obj") Then Return True
        Return False
    End Function

    Private Function IgnoreFiles(ByVal Filename As String) As Boolean
        Filename = Filename.ToLower
        If Filename.EndsWith(".application") Then Return True
        If Filename.EndsWith(".bak") Then Return True
        If Filename.EndsWith(".cache") Then Return True
        If Filename.EndsWith(".com") Then Return True
        If Filename.EndsWith(".db") Then Return True
        If Filename.EndsWith(".deploy") Then Return True
        If Filename.EndsWith(".lnk") Then Return True
        If Filename.EndsWith(".log") Then Return True
        If Filename.EndsWith(".ocx") Then Return True
        If Filename.EndsWith(".par2") Then Return True
        If Filename.EndsWith(".pdb") Then Return True
        If Filename.EndsWith(".sav") Then Return True
        If Filename.EndsWith(".scc") Then Return True
        If Filename.EndsWith(".suo") Then Return True
        If Filename.EndsWith(".tmp") Then Return True
        If Filename.EndsWith(".udl") Then Return True
        If Filename.EndsWith(".user") Then Return True
        If Filename.EndsWith(".vbw") Then Return True
        If Filename.EndsWith(".vspscc") Then Return True
        If Filename.EndsWith(".vssscc") Then Return True
        ' --- Check for batch files that vary by location ---
        If Filename.EndsWith("\buildall.bat") Then Return True
        If Filename.EndsWith("\buildtest.bat") Then Return True
        If Filename.EndsWith("\publishall.bat") Then Return True
        If Filename.EndsWith("\clearuserprograms.bat") Then Return True
        If Filename.EndsWith("\clearusersettings.bat") Then Return True
        ' --- Find files in specific directories ---
        If Filename.Contains("\bin\") AndAlso Filename.EndsWith(".exe") Then Return True
        If Filename.Contains("\bin\") AndAlso Filename.EndsWith(".dll") Then Return True
        If Filename.Contains("\bin\") AndAlso Filename.EndsWith(".xml") Then Return True
        If Filename.Contains("\bin\") AndAlso Filename.EndsWith(".config") Then Return True
        If Filename.Contains("\bin\") AndAlso Filename.EndsWith(".settings") Then Return True
        Return False
    End Function

#End Region

End Class
