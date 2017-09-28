' --------------------------------
' --- FormMain.vb - 09/28/2017 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/28/2017 - SBakker
'            - Switched to Arena.Common.Bootstrap.
' 09/26/2016 - SBakker
'            - Trying yet again to propogate deleted files correctly.
' 03/31/2016 - SBakker
'            - Fixed debug file compare program path to be handled better.
' 12/11/2015 - SBakker
'            - Run SaveMergeTransferList() before BuildZipFile(). It wasn't deleting files until next
'              time around, because the changes weren't written out until after the Zip file.
' 10/30/2015 - SBakker
'            - Made the "No zip files found" message be on the status bar instead of a MessageBox.
' 07/15/2015 - SBakker
'            - Fixed LocalStagingDir checking to build directory when it is not found. (Had been in
'              there, but lost sometime in the past.)
' 07/10/2015 - SBakker
'            - Fixed checking for directories to be consistent everywhere.
'            - Don't check for directories that aren't needed for the current operation.
' 06/24/2015 - SBakker
'            - Trap error when directory can't be created. Drive mapping may not exist anymore.
' 06/08/2015 - SBakker
'            - Added SaveCurrentSettings() on all three button clicks.
' 04/28/2015 - SBakker
'            - Added automatic deletion of files from target system which were deleted on sending
'              system.
' 04/22/2015 - SBakker
'            - Fixed so it will copy older files instead of skipping them. Wasn't handling them right!
' 04/09/2015 - SBakker
'            - Update and save TransferFilesData after unzipping and copying, so that files don't do a
'              double-bounce between servers.
' 03/16/2015 - SBakker
'            - Disable ButtonCompare when it shouldn't be clickable.
' 03/12/2015 - SBakker
'            - Added Compare button so that the Source and Merge dirs can be compared at any time.
' 03/05/2015 - SBakker
'            - Changing to use a merge directory. Changed files will be copied there and can then be
'              manually merged into or out of the source directory.
' 03/04/2015 - SBakker
'            - Only ask once when the Temp directory already exists.
' 02/13/2015 - SBakker
'            - Don't error out if a file is open by another process. Just ignore it and continue.
' 02/05/2015 - SBakker
'            - Added new DeletedMD5Tag (all zeroes) so that missing files can be identified.
' 01/15/2015 - SBakker
'            - Start up SourceManager.exe if they want to view differences. This is done while zipping
'              or unzipping and the Temp directory already exists. SourceManager will automatically
'              use the information sent and start comparing.
'            - Send the local copy of the zip file to the recycle bin, not the original. The Transfer
'              drive might not have a recycle bin if it is remote.
' 01/12/2015 - SBakker
'            - Added computer name to the ZIP filename when zipping files.
'            - Ignore any zip files with this computer name in them when unzipping. This will prevent
'              unzipping to the same machine that zipped it. Only works between exactly two machines!
'            - Changed TransferFilesList to only use MD5 tags!!! This will take care of any datetime
'              issues, which seem to happen a lot.
'            - Added DeleteFilesRecursive() to delete all files in the TempDir before deleteing the
'              directories themselves. Sometimes hidden or readonly files exist when copied by hand.
'            - Added message "ZIP file not found" when unzipping and no zip files exist.
'            - Added "\bin\*.settings" to ignored files list.
'            - Added icons to most message boxes.
' 12/19/2014 - SBakker
'            - Turn off FormCompare for now, until it has more functionallity.
' 11/25/2014 - SBakker
'            - Changed "cannot exist" to "already exists". Sounds better.
' 09/29/2014 - SBakker
'            - Did some cleanup on the Unzip code.
' 09/26/2014 - SBakker
'            - Adjust file DateTime on all unzipped files to be correct for the current timezone.
'            - Move Zip files to Recycle Bin instead of deleting permenently.
'            - When adjusting file DateTime, also hide files that start with ".", which are special.
' 09/24/2014 - SBakker
'            - Adding TransferFilesList and a hidden file ".TransferFilesData" to keep track of sizes
'              and dates of files. This will allow any file, no matter the datetime, to be copied if
'              it has changed. Doesn't matter if the file is older or newer, which "LastCopyDate"
'              couldn't handle.
'            - Removed "LastCopyDate.txt" handling, because a single date isn't good enough for file
'              comparison.
' 09/23/2014 - SBakker
'            - Creating TransferFiles utility.
'            - Set file attributes to Normal after copy to Temp, so they can be deleted later.
'            - Zip into local temp directory, then copy to target. Copy from target to local temp dir
'              then unzip from there. This should make zip/unzip faster on slow servers.
'            - Added better handling when no zip files exist or when no files are copied.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports Arena.Common.Bootstrap
Imports MD5ClassLibrary.MD5Utilities

Public Class FormMain

#Region " Internal Variables "

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private Const LocalTempDir As String = "C:\Temp"
    Private Const FileCompareProgramName As String = "SourceManager.exe"

    Private DirectoryCount As Integer
    Private FilesFound As Integer
    Private FilesCopied As Integer
    Private LocalStagingDir As String

    Friend Const TransferFilesDataName As String = ".TransferFilesData"
    Friend Const DeleteFilesDataName As String = ".DeleteFilesData"
    Friend Const Local2UTDOffsetFilename As String = ".Local2UTCOffset"
    Friend Const DateTimeFormatString As String = "yyyyMMdd_HHmmss"
    Friend Const DeletedMD5Tag As String = "00000000000000000000000000000000"

    Friend TransferFilesList As SortedList(Of String, String)

#End Region

#Region " Form Events "

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try

            Try
                If Bootstrapper.MustBootstrap Then
                    Me.Close()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(FuncName + vbCrLf + ex.Message,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.Close()
                Exit Sub
            End Try

            ' --- Update the settings from the last version ---
            If My.Settings.CallUpgrade Then
                My.Settings.Upgrade()
                My.Settings.CallUpgrade = False
                My.Settings.Save()
            End If

            ' --- Make sure the local temp directory exists ---
            If Not Directory.Exists(LocalTempDir) Then
                Directory.CreateDirectory(LocalTempDir)
            End If

            If Not String.IsNullOrWhiteSpace(My.Settings.Application) Then
                Dim TempApps() As String = My.Settings.Application.Split("|"c)
                ComboBoxApplication.Items.Clear()
                For Each CurrApp As String In TempApps
                    ComboBoxApplication.Items.Add(CurrApp)
                Next
                If Not String.IsNullOrWhiteSpace(My.Settings.LastApp) Then
                    For CurrIndex As Integer = 0 To ComboBoxApplication.Items.Count - 1
                        If My.Settings.LastApp = CStr(ComboBoxApplication.Items(CurrIndex)) Then
                            ' --- This will trigger all the other boxes to be filled ---
                            ComboBoxApplication.SelectedIndex = CurrIndex
                            Exit For
                        End If
                    Next
                End If
            End If

        Catch ex As Exception

            MessageBox.Show("Error starting " + My.Application.Info.AssemblyName + vbCrLf + vbCrLf + ex.Message,
                            Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()

        End Try

    End Sub

#End Region

#Region " Main Menu Events "

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AddApplicationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddApplicationToolStripMenuItem.Click
        Dim NewApp As String
        ' ------------------
        NewApp = InputBox("Enter new application name: ", My.Application.Info.AssemblyName, "")
        If String.IsNullOrWhiteSpace(NewApp) Then Exit Sub
        NewApp = NewApp.Trim
        If NewApp.Contains("=") OrElse NewApp.Contains("|") Then
            MessageBox.Show("Application name contains invalid characters",
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        For Each AppName As String In ComboBoxApplication.Items
            If AppName.ToUpper = NewApp.ToUpper Then
                MessageBox.Show("Application already in list",
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Next
        ComboBoxApplication.Items.Add(NewApp)
        If String.IsNullOrWhiteSpace(My.Settings.Application) Then
            My.Settings.Application = NewApp
        Else
            My.Settings.Application += "|" + NewApp
        End If
        My.Settings.Save()
        For CurrIndex As Integer = 0 To ComboBoxApplication.Items.Count - 1
            If CStr(ComboBoxApplication.Items(CurrIndex)) = NewApp Then
                ComboBoxApplication.SelectedIndex = CurrIndex
                Exit For
            End If
        Next
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAboutMain As New AboutMain
        TempAboutMain.ShowDialog()
    End Sub

#End Region

#Region " ComboBox Events "

    Private Sub ComboBoxApplication_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxApplication.SelectedIndexChanged

        ToolStripStatusLabelMain.Text = ""

        If ComboBoxApplication.SelectedIndex < 0 Then Exit Sub

        My.Settings.LastApp = CStr(ComboBoxApplication.Items(ComboBoxApplication.SelectedIndex))

        TextBoxSourceDir.Text = ""
        If Not String.IsNullOrWhiteSpace(My.Settings.SourceDir) Then
            Dim SourceDirList() As String = My.Settings.SourceDir.Split("|"c)
            For Each CurrSourceDir As String In SourceDirList
                If CurrSourceDir.StartsWith(My.Settings.LastApp + "=") Then
                    TextBoxSourceDir.Text = CurrSourceDir.Substring((My.Settings.LastApp + "=").Length)
                    Exit For
                End If
            Next
        End If

        TextBoxMergeDir.Text = ""
        If Not String.IsNullOrWhiteSpace(My.Settings.MergeDir) Then
            Dim TempDirList() As String = My.Settings.MergeDir.Split("|"c)
            For Each CurrMergeDir As String In TempDirList
                If CurrMergeDir.StartsWith(My.Settings.LastApp + "=") Then
                    TextBoxMergeDir.Text = CurrMergeDir.Substring((My.Settings.LastApp + "=").Length)
                    Exit For
                End If
            Next
        End If

        TextBoxTransferDir.Text = ""
        If Not String.IsNullOrWhiteSpace(My.Settings.TransferDir) Then
            Dim TransferDirList() As String = My.Settings.TransferDir.Split("|"c)
            For Each CurrTransferDir As String In TransferDirList
                If CurrTransferDir.StartsWith(My.Settings.LastApp + "=") Then
                    TextBoxTransferDir.Text = CurrTransferDir.Substring((My.Settings.LastApp + "=").Length)
                    Exit For
                End If
            Next
        End If

        My.Settings.Save()

        LocalStagingDir = Path.Combine(Path.Combine(Path.GetTempPath, My.Application.Info.AssemblyName), My.Settings.LastApp)

    End Sub

#End Region

#Region " Button Events "

    Private Sub ButtonCompare_Click(sender As Object, e As EventArgs) Handles ButtonCompare.Click

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim AppPath As String
        Dim Parameters As String
        Dim SourceDirName As String = TextBoxSourceDir.Text
        Dim MergeDirName As String = TextBoxMergeDir.Text
        ' -------------------------------------------------

        ToolStripStatusLabelMain.Text = ""

        ' --- Check for missing info ---
        If ComboBoxApplication.SelectedIndex < 0 Then Exit Sub

        Try
            If String.IsNullOrWhiteSpace(SourceDirName) Then Exit Sub
            If Not Directory.Exists(SourceDirName) Then
                MessageBox.Show("Source directory not found: " + SourceDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + SourceDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        Try
            If String.IsNullOrWhiteSpace(MergeDirName) Then Exit Sub
            If Not Directory.Exists(MergeDirName) Then
                MessageBox.Show("Merge directory not found: " + MergeDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + MergeDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        SaveCurrentSettings()

#If DEBUG Then
        AppPath = "C:\Utilities\Bin\" + FileCompareProgramName
        If Not File.Exists(AppPath) Then
            AppPath = "Y:\Utilities\Bin\" + FileCompareProgramName
        End If
        If Not File.Exists(AppPath) Then
            AppPath = "P:\Utilities\Bin\" + FileCompareProgramName
        End If
#Else
        AppPath = My.Application.Info.DirectoryPath + "\" + FileCompareProgramName
#End If

        Parameters = """" + ComboBoxApplication.Text + """ " +
                     """" + SourceDirName + """ " +
                     """" + MergeDirName + """"

        ' --- Disable controls ---
        ComboBoxApplication.Enabled = False
        TextBoxSourceDir.Enabled = False
        TextBoxMergeDir.Enabled = False
        TextBoxTransferDir.Enabled = False
        ButtonCompare.Enabled = False
        ButtonZipChanges.Enabled = False
        ButtonUnzipChanges.Enabled = False

        ' --- Now show differences that were just unzipped ---
        If File.Exists(AppPath) Then
            ' --- Start the application ---
            If Process.Start(AppPath, Parameters) Is Nothing Then
                MessageBox.Show(FuncName + vbCrLf + "Cannot start process: " + AppPath + RTrim(" " + Parameters),
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        ' --- Enable controls ---
        ComboBoxApplication.Enabled = True
        TextBoxSourceDir.Enabled = True
        TextBoxMergeDir.Enabled = True
        TextBoxTransferDir.Enabled = True
        ButtonCompare.Enabled = True
        ButtonZipChanges.Enabled = True
        ButtonUnzipChanges.Enabled = True

        ComboBoxApplication.Focus() ' Otherwise doesn't seem to refresh
        Application.DoEvents()
        ButtonCompare.Focus()
        Application.DoEvents()

    End Sub

    Private Sub ButtonZipChanges_Click(sender As Object, e As EventArgs) Handles ButtonZipChanges.Click

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim MergeDirName As String = TextBoxMergeDir.Text
        Dim TransferDirName As String = TextBoxTransferDir.Text
        ' -----------------------------------------------------

        ToolStripStatusLabelMain.Text = ""

        ' --- Check for missing info ---
        If ComboBoxApplication.SelectedIndex < 0 Then Exit Sub

        Try
            If String.IsNullOrWhiteSpace(MergeDirName) Then Exit Sub
            If Not Directory.Exists(MergeDirName) Then
                MessageBox.Show("Merge directory not found: " + MergeDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + MergeDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        Try
            If String.IsNullOrWhiteSpace(TransferDirName) Then Exit Sub
            If Not Directory.Exists(TransferDirName) Then
                MessageBox.Show("Transfer directory not found: " + TransferDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + TransferDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        SaveCurrentSettings()

        ' --- Disable controls ---
        ComboBoxApplication.Enabled = False
        TextBoxSourceDir.Enabled = False
        TextBoxMergeDir.Enabled = False
        TextBoxTransferDir.Enabled = False
        ButtonCompare.Enabled = False
        ButtonZipChanges.Enabled = False
        ButtonUnzipChanges.Enabled = False

        ' --- Copy files and build Zip file ---
        GetMergeTransferList()
        CopyMergeToTemp()
        SaveMergeTransferList()
        If FilesCopied > 0 Then
            BuildZipFile()
        End If
        UpdateStatusLine(True)

        ' --- Remove Temp directory to be ready for the next time ---
        Try
            If Directory.Exists(LocalStagingDir) Then
                Directory.Delete(LocalStagingDir, True)
            End If
        Catch ex As Exception
            MessageBox.Show("Error deleting directory: " + LocalStagingDir + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        ' --- Enable controls ---
        ComboBoxApplication.Enabled = True
        TextBoxSourceDir.Enabled = True
        TextBoxMergeDir.Enabled = True
        TextBoxTransferDir.Enabled = True
        ButtonCompare.Enabled = True
        ButtonZipChanges.Enabled = True
        ButtonUnzipChanges.Enabled = True

        ComboBoxApplication.Focus() ' Otherwise doesn't seem to refresh
        Application.DoEvents()
        ButtonZipChanges.Focus()
        Application.DoEvents()

    End Sub

    Private Sub ButtonUnzipChanges_Click(sender As Object, e As EventArgs) Handles ButtonUnzipChanges.Click

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim ZipFileCount As Integer = 0
        Dim AppName As String = My.Settings.LastApp
        Dim SourceDirName As String = TextBoxSourceDir.Text
        Dim MergeDirName As String = TextBoxMergeDir.Text
        Dim TransferDirName As String = TextBoxTransferDir.Text
        ' -----------------------------------------------------

        ToolStripStatusLabelMain.Text = ""

        ' --- Check for missing info ---
        If ComboBoxApplication.SelectedIndex < 0 Then Exit Sub

        Try
            If String.IsNullOrWhiteSpace(SourceDirName) Then Exit Sub
            If Not Directory.Exists(SourceDirName) Then
                MessageBox.Show("Source directory not found: " + SourceDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + SourceDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        Try
            If String.IsNullOrWhiteSpace(MergeDirName) Then Exit Sub
            If Not Directory.Exists(MergeDirName) Then
                MessageBox.Show("Merge directory not found: " + MergeDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + MergeDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        Try
            If String.IsNullOrWhiteSpace(TransferDirName) Then Exit Sub
            If Not Directory.Exists(TransferDirName) Then
                MessageBox.Show("Transfer directory not found: " + TransferDirName,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + TransferDirName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        Try
            If String.IsNullOrWhiteSpace(LocalStagingDir) Then Exit Sub
            If Not Directory.Exists(LocalStagingDir) Then
                Try
                    Directory.CreateDirectory(LocalStagingDir)
                Catch ex As Exception
                    MessageBox.Show("Local staging directory not found: " + LocalStagingDir,
                                    My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show("Error accessing directory: " + LocalStagingDir + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try

        SaveCurrentSettings()

        ' --- Count the number of Zip files waiting ---
        Dim TransferDirInfo As New DirectoryInfo(TransferDirName)
        For Each CurrFile As FileInfo In TransferDirInfo.GetFiles(AppName + "_*.zip")
            ' --- Ignore files which came from this computer ---
            If CurrFile.Name.EndsWith("_" + My.Computer.Name.ToUpper + ".zip") Then
                Continue For
            End If
            ' --- Count the files found ---
            ZipFileCount += 1
        Next

        If ZipFileCount = 0 Then
            ToolStripStatusLabelMain.Text = "No zip files found"
            Exit Sub
        End If

        Dim AppPath As String
        Dim Parameters As String

#If DEBUG Then
        AppPath = "C:\Utilities\Bin\" + FileCompareProgramName
        If Not File.Exists(AppPath) Then
            AppPath = "Y:\Utilities\Bin\" + FileCompareProgramName
        End If
        If Not File.Exists(AppPath) Then
            AppPath = "P:\Utilities\Bin\" + FileCompareProgramName
        End If
#Else
        AppPath = My.Application.Info.DirectoryPath + "\" + FileCompareProgramName
#End If

        Parameters = """" + ComboBoxApplication.Text + """ " +
                     """" + SourceDirName + """ " +
                     """" + MergeDirName + """"

        ' --- Disable controls ---
        ComboBoxApplication.Enabled = False
        TextBoxSourceDir.Enabled = False
        TextBoxMergeDir.Enabled = False
        TextBoxTransferDir.Enabled = False
        ButtonCompare.Enabled = False
        ButtonZipChanges.Enabled = False
        ButtonUnzipChanges.Enabled = False

        ' --- Unzip files ---
        GetMergeTransferList()
        SaveCurrentSettings()

        If ExtractZipFile() Then

            ' --- Get rid of files which were deleted on sending system ---
            Try
                If File.Exists(LocalStagingDir + "\" + DeleteFilesDataName) Then
                    Dim DeleteFilesList() As String = File.ReadAllLines(LocalStagingDir + "\" + DeleteFilesDataName)
                    For Each TempFilenameToDelete As String In DeleteFilesList
                        If File.Exists(MergeDirName + "\" + TempFilenameToDelete) Then
                            File.SetAttributes(MergeDirName + "\" + TempFilenameToDelete, FileAttributes.Normal)
                            File.Delete(MergeDirName + "\" + TempFilenameToDelete)
                        End If
                    Next
                    File.SetAttributes(LocalStagingDir + "\" + DeleteFilesDataName, FileAttributes.Normal)
                    File.Delete(LocalStagingDir + "\" + DeleteFilesDataName)
                End If
            Catch ex As Exception
                ' --- Ignore delete errors  ---
            End Try

            ' --- Copy unzipped files to MergeDir ---
            DirectoryCount = 0
            FilesFound = 0
            FilesCopied = 0
            CopyFilesRecursive(LocalStagingDir, MergeDirName, LocalStagingDir.Length)
            SaveMergeTransferList()

            ' --- Now show differences that were just unzipped ---
            If File.Exists(AppPath) Then
                ' --- Start the application ---
                If Process.Start(AppPath, Parameters) Is Nothing Then
                    MessageBox.Show(FuncName + vbCrLf + "Cannot start process: " + AppPath + RTrim(" " + Parameters),
                                    My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        End If

        ' --- Enable controls ---
        ComboBoxApplication.Enabled = True
        TextBoxSourceDir.Enabled = True
        TextBoxMergeDir.Enabled = True
        TextBoxTransferDir.Enabled = True
        ButtonCompare.Enabled = True
        ButtonZipChanges.Enabled = True
        ButtonUnzipChanges.Enabled = True

        ComboBoxApplication.Focus() ' Otherwise doesn't seem to refresh
        Application.DoEvents()
        ButtonUnzipChanges.Focus()
        Application.DoEvents()

    End Sub

#End Region

#Region " StatusLine Routines "

    Private Sub UpdateStatusLine(ByVal DoneFlag As Boolean)
        If DoneFlag AndAlso ToolStripStatusLabelMain.Text.EndsWith("...") Then
            ToolStripStatusLabelMain.Text += " Done"
        Else
            ToolStripStatusLabelMain.Text = "Directories: " + DirectoryCount.ToString +
                                            " - Files Found: " + FilesFound.ToString +
                                            " - Files Copied: " + FilesCopied.ToString
            If DoneFlag Then
                ToolStripStatusLabelMain.Text += " - Done"
            End If
        End If
    End Sub

#End Region

#Region " Settings Routines "

    Private Sub SaveCurrentSettings()

        Dim Found As Boolean
        Dim NewSetting As StringBuilder
        ' -----------------------------

        ' --- SourceDir ---
        Dim SourceDirList() As String = My.Settings.SourceDir.Split("|"c)
        Dim CurrSourceDir As String
        NewSetting = New StringBuilder
        Found = False
        TextBoxSourceDir.Text = TextBoxSourceDir.Text.Trim
        For CurrIndex As Integer = 0 To SourceDirList.Count - 1
            CurrSourceDir = SourceDirList(CurrIndex)
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            If CurrSourceDir.StartsWith(My.Settings.LastApp + "=") Then
                NewSetting.Append(My.Settings.LastApp + "=" + TextBoxSourceDir.Text)
                Found = True
            Else
                NewSetting.Append(SourceDirList(CurrIndex))
            End If
        Next
        If Not Found Then
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            NewSetting.Append(My.Settings.LastApp + "=" + TextBoxSourceDir.Text)
        End If
        My.Settings.SourceDir = NewSetting.ToString

        ' --- MergeDir ---
        Dim MergeDirList() As String = My.Settings.MergeDir.Split("|"c)
        Dim CurrMergeDir As String
        NewSetting = New StringBuilder
        Found = False
        TextBoxMergeDir.Text = TextBoxMergeDir.Text.Trim
        For CurrIndex As Integer = 0 To MergeDirList.Count - 1
            CurrMergeDir = MergeDirList(CurrIndex)
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            If CurrMergeDir.StartsWith(My.Settings.LastApp + "=") Then
                NewSetting.Append(My.Settings.LastApp + "=" + TextBoxMergeDir.Text)
                Found = True
            Else
                NewSetting.Append(MergeDirList(CurrIndex))
            End If
        Next
        If Not Found Then
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            NewSetting.Append(My.Settings.LastApp + "=" + TextBoxMergeDir.Text)
        End If
        My.Settings.MergeDir = NewSetting.ToString

        ' --- TransferDir ---
        Dim TransferDirList() As String = My.Settings.TransferDir.Split("|"c)
        Dim CurrTransferDir As String
        NewSetting = New StringBuilder
        Found = False
        TextBoxTransferDir.Text = TextBoxTransferDir.Text.Trim
        For CurrIndex As Integer = 0 To TransferDirList.Count - 1
            CurrTransferDir = TransferDirList(CurrIndex)
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            If CurrTransferDir.StartsWith(My.Settings.LastApp + "=") Then
                NewSetting.Append(My.Settings.LastApp + "=" + TextBoxTransferDir.Text)
                Found = True
            Else
                NewSetting.Append(TransferDirList(CurrIndex))
            End If
        Next
        If Not Found Then
            If NewSetting.Length > 0 Then
                NewSetting.Append("|")
            End If
            NewSetting.Append(My.Settings.LastApp + "=" + TextBoxTransferDir.Text)
        End If
        My.Settings.TransferDir = NewSetting.ToString

        ' --- Save the settings ---
        My.Settings.Save()

    End Sub

#End Region

#Region " Copy Routines "

    Private Sub CopyMergeToTemp()
        ' --- Delete temp directory if it already exists ---
        If Directory.Exists(LocalStagingDir) Then
            Try
                DeleteFilesRecursive(LocalStagingDir)
                Directory.Delete(LocalStagingDir, True)
            Catch ex As Exception
                MessageBox.Show("Error deleting directory: " + LocalStagingDir + vbCrLf + ex.Message,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
        End If
        ' --- Copy all changed files ---
        DirectoryCount = 0
        FilesFound = 0
        FilesCopied = 0
        UpdateStatusLine(False)
        Application.DoEvents()
        CopyFilesRecursive(TextBoxMergeDir.Text, LocalStagingDir, TextBoxMergeDir.Text.Length)
    End Sub

    Private Sub CopyFilesRecursive(ByVal FromDir As String, ByVal ToDir As String, ByVal BaseDirLength As Integer)
        Dim RelativeFilename As String
        Dim FromFileDateTimeSize As String
        Dim FromFileMD5 As String
        Dim FromDirInfo As New DirectoryInfo(FromDir)
        ' -------------------------------------------
        DirectoryCount += 1
        UpdateStatusLine(False)
        Application.DoEvents()
        ' --- Go through all files first ---
        For Each CurrFromFile As FileInfo In FromDirInfo.GetFiles
            FilesFound += 1
            UpdateStatusLine(False)
            Application.DoEvents()
            If (CurrFromFile.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            If IgnoreFiles(CurrFromFile.FullName) Then Continue For
            ' --- Check if file needs to be copied ---
            RelativeFilename = CurrFromFile.FullName.Substring(BaseDirLength + 1).ToLower
            Try
                FromFileMD5 = CalcMD5(CurrFromFile.FullName)
            Catch ex As Exception
                Continue For
            End Try
            If TransferFilesList.ContainsKey(RelativeFilename) Then
                Dim ItemValue As String = TransferFilesList.Item(RelativeFilename)
                If ItemValue = FromFileMD5 Then
                    Continue For
                End If
                If ItemValue.Contains(" ") Then
                    ' --- Old TransferFilesList format, "Date_Time Size" ---
                    FromFileDateTimeSize = CurrFromFile.LastWriteTimeUtc.ToString(DateTimeFormatString) + " " + CurrFromFile.Length.ToString
                    If ItemValue = FromFileDateTimeSize Then
                        ' --- File matches, just update information and continue ---
                        TransferFilesList.Item(RelativeFilename) = FromFileMD5
                        Continue For
                    End If
                End If
                ' --- Needs to be copied, file has changed ---
                TransferFilesList.Item(RelativeFilename) = FromFileMD5
            Else
                TransferFilesList.Add(RelativeFilename, FromFileMD5)
            End If
            ' --- Copy file ---
            If Not Directory.Exists(ToDir) Then
                Directory.CreateDirectory(ToDir)
            End If
            If File.Exists(ToDir + "\" + CurrFromFile.Name) Then
                ''If File.GetLastWriteTimeUtc(ToDir + "\" + CurrFromFile.Name) >= CurrFromFile.LastWriteTimeUtc Then
                ''    Continue For
                ''End If
                File.Delete(ToDir + "\" + CurrFromFile.Name)
            End If
            File.Copy(CurrFromFile.FullName, ToDir + "\" + CurrFromFile.Name)
            File.SetAttributes(ToDir + "\" + CurrFromFile.Name, FileAttributes.Normal)
            FilesCopied += 1
            UpdateStatusLine(False)
            Application.DoEvents()
        Next
        ' --- Go through all subdirectories next ---
        For Each CurrSubDir As DirectoryInfo In FromDirInfo.GetDirectories
            If (CurrSubDir.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            If IgnoreDir(CurrSubDir.FullName) Then Continue For
            CopyFilesRecursive(CurrSubDir.FullName, ToDir + "\" + CurrSubDir.Name, BaseDirLength)
        Next
    End Sub

    Private Function IgnoreDir(ByVal DirPath As String) As Boolean
        ' --- Ignore known configuration directories ---
        If DirPath.StartsWith(".") Then Return True
        If DirPath.Contains("\.") Then Return True
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
        ' --- Ignore known configuration files ---
        If Filename.StartsWith(".") Then Return True
        If Filename.Contains("\.") Then Return True
        ' --- Ignore specific file extensions ---
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

#Region " Delete Routines "

    Private Sub DeleteFilesRecursive(TempDir As String)
        Dim TempDirInfo As New DirectoryInfo(TempDir)
        ' -------------------------------------------
        ' --- Go through all files first ---
        For Each CurrFile As FileInfo In TempDirInfo.GetFiles
            If CurrFile.Attributes <> FileAttributes.Normal Then
                CurrFile.Attributes = FileAttributes.Normal
            End If
            File.Delete(CurrFile.FullName)
        Next
        ' --- Go through all subdirectories next ---
        For Each CurrSubDir As DirectoryInfo In TempDirInfo.GetDirectories
            If CurrSubDir.Attributes <> FileAttributes.Directory Then
                CurrSubDir.Attributes = FileAttributes.Directory
            End If
            DeleteFilesRecursive(TempDir + "\" + CurrSubDir.Name)
        Next
    End Sub

#End Region

#Region " Zip Routines "

    Private Sub BuildZipFile()

        Dim ZipFileBaseName As String
        Dim ZipFileFullName As String
        Dim ZipFileLocalName As String
        Dim AppName As String = My.Settings.LastApp
        Dim TransferDirName As String = TextBoxTransferDir.Text
        Dim CurrDateTime As DateTime = Now
        Dim CurrUTCOffset As New DateTimeOffset(CurrDateTime)
        ' ---------------------------------------------------

        File.WriteAllText(LocalStagingDir + "\" + Local2UTDOffsetFilename, CurrUTCOffset.Offset.ToString)
        File.SetAttributes(LocalStagingDir + "\" + Local2UTDOffsetFilename, FileAttributes.Hidden)

        BuildDeleteFilesList()

        ' --- Include computer name so this computer won't try to unzip it ---
        ZipFileBaseName = AppName + "_" + CurrDateTime.ToUniversalTime.ToString(DateTimeFormatString) + "_" + My.Computer.Name.ToUpper + ".zip"
        ZipFileFullName = TransferDirName + "\" + ZipFileBaseName
        ZipFileLocalName = LocalTempDir + "\" + ZipFileBaseName

        ToolStripStatusLabelMain.Text += " - Zipping..."
        Application.DoEvents()
        Cursor.Current = Cursors.WaitCursor
        Try
            ZipFile.CreateFromDirectory(LocalStagingDir, ZipFileLocalName)
            File.Copy(ZipFileLocalName, ZipFileFullName)
            File.Delete(ZipFileLocalName)
        Catch ex As Exception
            MessageBox.Show("Error zipping " + LocalStagingDir + " to " + ZipFileFullName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Private Function ExtractZipFile() As Boolean

        Dim ZipFileFullName As String = ""
        Dim ZipFileLocalName As String = ""
        Dim CurrUTCOffset As New DateTimeOffset(Now)
        Dim ZipOffsetTimespan As TimeSpan
        Dim ZipFilesOffset As String = ""
        Dim AppName As String = My.Settings.LastApp
        Dim TransferDirName As String = TextBoxTransferDir.Text
        ' -----------------------------------------------------

        DirectoryCount = 0
        FilesFound = 0
        FilesCopied = 0

        Dim TransferDirInfo As New DirectoryInfo(TransferDirName)

        For Each CurrFile As FileInfo In TransferDirInfo.GetFiles(AppName + "_*.zip")
            ' --- Ignore files which came from this computer ---
            If CurrFile.Name.EndsWith("_" + My.Computer.Name.ToUpper + ".zip") Then
                Continue For
            End If
            ' --- Get the lowest zip file name ---
            If String.IsNullOrWhiteSpace(ZipFileFullName) OrElse CurrFile.FullName < ZipFileFullName Then
                ZipFileFullName = CurrFile.FullName
                ZipFileLocalName = LocalTempDir + "\" + CurrFile.Name
            End If
        Next

        If String.IsNullOrWhiteSpace(ZipFileFullName) Then
            ToolStripStatusLabelMain.Text = "Done"
            MessageBox.Show("No zip files found: " + TransferDirName + "\" + ComboBoxApplication.Text,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        ' --- Delete local staging directory ---
        If Directory.Exists(LocalStagingDir) Then
            Try
                DeleteFilesRecursive(LocalStagingDir)
                Directory.Delete(LocalStagingDir, True)
            Catch ex As Exception
                MessageBox.Show("Error deleting directory: " + LocalStagingDir + vbCrLf + ex.Message,
                                My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End Try
        End If

        ToolStripStatusLabelMain.Text = "Unzipping..."
        Application.DoEvents()

        Try
            Cursor.Current = Cursors.WaitCursor
            File.Copy(ZipFileFullName, ZipFileLocalName)
            ZipFile.ExtractToDirectory(ZipFileLocalName, LocalStagingDir)
            If File.Exists(LocalStagingDir + "\" + Local2UTDOffsetFilename) Then
                ZipFilesOffset = File.ReadAllText(LocalStagingDir + "\" + Local2UTDOffsetFilename)
            End If
            ' --- Reset all file datetimes based on offset difference ---
            If Not String.IsNullOrWhiteSpace(ZipFilesOffset) AndAlso ZipFilesOffset <> CurrUTCOffset.Offset.ToString Then
                ' --- Get the difference between the two time zones ---
                ZipOffsetTimespan = TimeSpan.Parse(CurrUTCOffset.Offset.ToString.Replace("+", ""))
                ZipOffsetTimespan = ZipOffsetTimespan.Subtract(TimeSpan.Parse(ZipFilesOffset.Replace("+", "")))
                ' --- Adjust all files unzipped by this time zone difference ---
                DirectoryCount = 0
                FilesFound = 0
                FilesCopied = 0
                UpdateStatusLine(False)
                Application.DoEvents()
                UpdateFileDateTimeRecursive(LocalStagingDir, ZipOffsetTimespan)
            End If
            Cursor.Current = Cursors.Default
            ''FilesCopied += 1
        Catch ex As Exception
            MessageBox.Show("Error unzipping " + ZipFileFullName + " to " + LocalStagingDir + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

        Try
            My.Computer.FileSystem.DeleteFile(ZipFileLocalName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            File.Delete(ZipFileFullName) ' no need to recycle, might fail anyway
        Catch ex As Exception
            MessageBox.Show("Error deleting " + ZipFileFullName + vbCrLf + ex.Message,
                            My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        UpdateStatusLine(True)

        Return True

    End Function

    Private Sub UpdateFileDateTimeRecursive(ByVal TempDir As String, ByVal ZipOffsetTimespan As TimeSpan)
        Dim TempDirInfo As New DirectoryInfo(TempDir)
        Dim UpdatedLastWriteTimeUTC As DateTimeOffset
        ' -------------------------------------------
        DirectoryCount += 1
        UpdateStatusLine(False)
        Application.DoEvents()
        ' --- Go through all files first ---
        For Each CurrTempFile As FileInfo In TempDirInfo.GetFiles
            FilesFound += 1
            UpdateStatusLine(False)
            Application.DoEvents()
            ' --- Timey-wimey stuff ---
            UpdatedLastWriteTimeUTC = DateTimeOffset.Parse(CurrTempFile.LastWriteTimeUtc.ToString + " +00:00")
            UpdatedLastWriteTimeUTC = UpdatedLastWriteTimeUTC.Add(ZipOffsetTimespan)
            CurrTempFile.LastWriteTimeUtc = CDate(UpdatedLastWriteTimeUTC.ToString)
            If CurrTempFile.Name.StartsWith(".") Then
                File.SetAttributes(CurrTempFile.FullName, FileAttributes.Hidden)
            End If
        Next
        ' --- Go through all subdirectories next ---
        For Each CurrSubDir As DirectoryInfo In TempDirInfo.GetDirectories
            If (CurrSubDir.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            If IgnoreDir(CurrSubDir.FullName) Then Continue For
            UpdateFileDateTimeRecursive(CurrSubDir.FullName, ZipOffsetTimespan)
        Next
    End Sub

#End Region

#Region " TransferFilesList Routines "

    Private Sub GetMergeTransferList()
        Dim TabPos As Integer
        Dim TransferListFilename As String
        ' --------------------------------
        TransferListFilename = TextBoxMergeDir.Text + "\" + TransferFilesDataName
        TransferFilesList = New SortedList(Of String, String)
        If File.Exists(TransferListFilename) Then
            For Each CurrLine As String In File.ReadLines(TransferListFilename, Encoding.UTF8)
                If String.IsNullOrWhiteSpace(CurrLine) Then Continue For
                TabPos = CurrLine.IndexOf(vbTab)
                If TabPos < 0 Then Continue For
                TransferFilesList.Add(CurrLine.Substring(0, TabPos), CurrLine.Substring(TabPos + 1))
            Next
        End If
    End Sub

    Private Sub SaveMergeTransferList()
        Dim Result As New StringBuilder
        Dim TransferListFilename As String
        ' --------------------------------
        TransferListFilename = TextBoxMergeDir.Text + "\" + TransferFilesDataName
        For CurrIndex As Integer = 0 To TransferFilesList.Count - 1
            If File.Exists(TextBoxMergeDir.Text + "\" + TransferFilesList.Keys(CurrIndex)) Then
                If TransferFilesList.Values(CurrIndex) = DeletedMD5Tag Then
                    Throw New SystemException("Existing file has deleted MD5")
                End If
            Else
                ' --- Mark the current file as deleted in memory also ---
                TransferFilesList(TransferFilesList.Keys(CurrIndex)) = DeletedMD5Tag
            End If
            Result.Append(TransferFilesList.Keys(CurrIndex))
            Result.Append(vbTab)
            Result.AppendLine(TransferFilesList.Values(CurrIndex))
        Next
        If File.Exists(TransferListFilename) Then
            File.SetAttributes(TransferListFilename, FileAttributes.Normal)
        End If
        File.WriteAllText(TransferListFilename, Result.ToString, Encoding.UTF8)
        File.SetAttributes(TransferListFilename, FileAttributes.Hidden)
    End Sub

    Private Sub BuildDeleteFilesList()
        Dim Count As Integer = 0
        Dim Result As New StringBuilder
        Dim DeleteListFilename As String
        ' ------------------------------
        DeleteListFilename = LocalStagingDir + "\" + DeleteFilesDataName
        For CurrIndex As Integer = 0 To TransferFilesList.Count - 1
            If TransferFilesList.Values(CurrIndex) = DeletedMD5Tag Then
                Count += 1
                Result.AppendLine(TransferFilesList.Keys(CurrIndex))
            End If
        Next
        If Count > 0 Then
            File.WriteAllText(DeleteListFilename, Result.ToString, Encoding.UTF8)
            File.SetAttributes(DeleteListFilename, FileAttributes.Hidden)
        End If
    End Sub

#End Region

End Class
