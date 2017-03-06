' --------------------------------
' --- FormMain.vb - 02/06/2017 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 02/06/2017 - SBakker
'            - Changed settings to use SVRTEST2 instead of SVRTEST.
' 12/05/2016 - SBakker
'            - Changed setting name from INIPathFIS to INIPathFIS2, so it could point to SVRREPORT.
' 07/26/2016 - SBakker
'            - Added environment "PC" to the dropdown.
' 04/11/2016 - SBakker
'            - Switched two checks for VB6 to end up with Program Files (x86) so error is logical.
' 03/15/2016 - SBakker
'            - Added item "FISTest" for Steve to do some SQL 2014 testing.
' 04/28/2015 - SBakker
'            - Added My.Settings.INIPathLocalAlt2 so that the full UNC path of the user home directory
'              can be used.
'            - Replace "*" with GetUserName() in all local paths, in case the path is of the format
'              "\\DHFILE\H_*\...". This matches how SourceManager handles local path settings.
' 03/03/3014 - SBakker
'            - Updated AboutMain to show the current executable path at the bottom.
' 02/24/2014 - SBakker
'            - Added Bootstrap loading all programs to another location, and then running from there.
' 10/23/2013 - SBakker
'            - Renamed AboutBox1 to AboutBoxMain.
' 09/17/2013 - SBakker
'            - Added additional error information.
' 03/09/2012 - SBakker
'            - Fixed errors in use of LastEnv and LastVolume. They got reversed in some
'              places.
' 02/01/2012 - SBakker
'            - Added EOY environment.
' 01/31/2012 - SBakker
'            - Disabled "Add Comments" checkbox - it hasn't been completely written yet.
' 12/02/2010 - SBakker
'            - Modified to put the VB6EXE command into a batch file, then run the batch file
'              in the background. For some reason, running it directly causes a "File Not
'              Found" error, no matter how it is done. It appears to be a relative path
'              issue. But using a batch file works. This is only a problem with .NET 4.0.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 09/09/2009 - SBakker
'            - Updated to use new version of CompileCadol.
' 05/01/2009 - SBakker - URD 11236
'            - Move IDRIS Acceptance to SVRTEST.
'            - Fixed Option Strict On issues.
' 04/06/2009 - SBAKKER
'            - Changed TEST to be on SVRTEST, merged all the three FIS's into
'              one option, FIS on SVRFIS.
' 10/08/2008 - SBAKKER - URD 11164
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
'            - Handle cross compiling between 32-bit and 64-bit systems.
'            - Added SVRFIS to list of environments.
' 08/19/2008 - SBAKKER
'            - Added LastVolume and LastLibrary settings.
' 06/11/2008 - SBAKKER
'            - Removed unused Path definitions. They are now in Settings.
'            - Added My.Settings.LocalCompilePath, in case it needs changing.
'            - Added My.Settings.CallUpgrade to pull previous version settings.
' 06/09/2008 - SBAKKER
'            - Switched to use CompileCadol.dll.
' 03/03/2008 - SBAKKER
'            - Added setting INIPathLocalAlt, using "Y:\IDRIS\...", in case the
'              IDRIS files are on the Y: drive. (Mine are!)
' 08/27/2007 - SBAKKER
'            - Added error checking that will show the proper error when an
'              expression can't be evaluated.
'            - Build "ToDir" if not found.
' 04/27/2007 - SBAKKER - URD 10950
'            - Fixed to allow new Production SourceSafe folder structure to be
'              used for source and object files.
'            - Changed INIPath* variables into Application settings.
' 01/18/2007 - sbakker - Leave "%" in EXE and VBP names, while replacing with
'              "_" for directory names.
' 01/10/2007 - sbakker - Added logic to handle both %SYSVOL and _SYSVOL,
'              %IDRISYS and _IDRISYS.
' 01/03/2007 - Added compiling to a local directory, then moving the EXE to the
'              target directory. This will speed up compiles on network drives.
' 11/17/2006 - Don't turn off read-only flag on any output files. Instead, just
'              show an error and allow them to abort/retry/ignore. Files need to
'              be properly checked out of SourceSafe.
'            - Added background thread for performing the VB6 compile. This
'              makes the front end much more responsive.
' 10/31/2006 - Added checking for "FromPathLocal", etc, if the computer name
'              matches the start of the "FromPath", etc, name. This will make
'              the compiles go so much faster, if it is not using UNC paths.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.SystemUtils
Imports CompileCadol
Imports CadolSourceDataClass
Imports System.ComponentModel
Imports System.IO
Imports System.Threading

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    ' --- Local compile path speeds up the compile process ---
    Private LocalCompilePath As String = "C:\Temp"

    Private INIFilename As String = ""
    Private CommonPath As String = ""
    Private VB6EXE As String = ""
    Private Compiling As Boolean = False
    Private Cancelled As Boolean = False
    Private TotalErrors As Integer = 0
    Private TotalCompiled As Integer = 0
    Private TotalProjectsBuilt As Integer = 0
    Private TotalLibsCompiled As Integer = 0

    Private Sub FormMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try
            If Arena_Bootstrap.BootstrapClass.CopyProgramsToLaunchPath Then
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Me.Close()
            Exit Sub
        End Try

        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
            My.Settings.Save()
        End If

        Application.DoEvents()

        If My.Settings.LastEnv <> "" Then
            EnvCombo.Text = My.Settings.LastEnv
        End If

        If My.Settings.LocalCompilePath <> "" Then
            LocalCompilePath = My.Settings.LocalCompilePath
        End If

    End Sub

    Private Sub FormMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If BackgroundWorker1.IsBusy Then
                e.Cancel = True
                MessageBox.Show("Library is still being compiled - Cannot exit", "Cannot Exit", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If Compiling Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        My.Settings.Save()
        Cancelled = True
    End Sub

    Private Sub FormMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState <> FormWindowState.Minimized Then
            CadolProg.Width = (Me.ClientSize.Width - 8) \ 2
            VB6Prog.Left = (Me.ClientSize.Width + 2) \ 2
            VB6Prog.Width = (Me.ClientSize.Width - 8) \ 2
        End If
    End Sub

    Private Sub VolumeCombo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VolumeCombo.SelectedIndexChanged
        Dim TempName As String
        Dim Libraries() As String
        ' -----------------------
        LibraryCombo.Items.Clear()
        If VolumeCombo.SelectedIndex <= 0 Then ' selected all
            Exit Sub
        End If
        LibraryCombo.Items.Add("--- ALL ---")
        If SourcePath.Text <> "" AndAlso VolumeCombo.SelectedIndex >= 0 Then
            If VolumeCombo.SelectedIndex <= 0 Then
                My.Settings.LastVolume = ""
            Else
                My.Settings.LastVolume = VolumeCombo.Text
            End If
            My.Settings.Save()
            TempName = SourcePath.Text + "\" + CStr(VolumeCombo.Items(VolumeCombo.SelectedIndex))
            Libraries = Directory.GetDirectories(TempName)
            For Each TempName In Libraries
                TempName = TempName.Substring(TempName.LastIndexOf("\") + 1).ToUpper
                If TempName <> "GENS" AndAlso TempName <> "INCLUDE" Then
                    LibraryCombo.Items.Add(TempName)
                End If
            Next
            If My.Settings.LastLibrary <> "" Then
                LibraryCombo.Text = My.Settings.LastLibrary
            Else
                LibraryCombo.SelectedIndex = 0
            End If
        Else
            LibraryCombo.SelectedIndex = 0
        End If
    End Sub

    Private Sub StartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                Handles StartButton.Click, StartToolStripMenuItem.Click
        If EnvCombo.Text = "" Then Exit Sub
        If SourcePath.Text = "" Then Exit Sub
        If TargetPath.Text = "" Then Exit Sub
        If VolumeCombo.Items.Count = 0 Then Exit Sub
        If VolumeCombo.Text = "" Then Exit Sub
        If VolumeCombo.Text <> "--- ALL ---" Then
            If LibraryCombo.Text = "" Then Exit Sub
        End If
        ' --- prevent any changes while compiling ---
        For Each c As Control In Me.Controls
            If c.GetType Is GetType(Label) Then Continue For
            If c.GetType Is GetType(RichTextBox) Then Continue For
            c.Enabled = False
        Next
        Cancelled = False
        Compiling = True
        StopCompiling.Enabled = True
        TotalErrors = 0
        TotalCompiled = 0
        TotalProjectsBuilt = 0
        TotalLibsCompiled = 0
        ' --- compile the programs ---
        CadolProg.Text = ""
        VB6Prog.Text = ""
        StatusLabel.Text = "Starting..."
        Cursor.Current = Cursors.WaitCursor
        If VolumeCombo.SelectedIndex = 0 Then ' all volumes
            CompileAllVolumes(SourcePath.Text, TargetPath.Text)
        Else
            If LibraryCombo.Items.Count > 0 Then
                If LibraryCombo.SelectedIndex = 0 Then
                    CompileAllLibs(SourcePath.Text + "\" + VolumeCombo.Text, _
                                   TargetPath.Text + "\" + VolumeCombo.Text)
                Else
                    CompileSingleLib(SourcePath.Text + "\" + VolumeCombo.Text, _
                                     TargetPath.Text + "\" + VolumeCombo.Text, _
                                     LibraryCombo.Text)
                End If
            End If
        End If
        Cursor.Current = Cursors.Default
        If TotalErrors > 0 Then
            StatusLabel.Text = "Total Errors Found: " + TotalErrors.ToString
            MessageBox.Show("Total Errors Found: " + TotalErrors.ToString, "IDRISMakeLib")
        ElseIf Cancelled Then
            StatusLabel.Text = "Cancelled"
            MessageBox.Show("Cancelled", "IDRISMakeLib")
        ElseIf CheckChangedOnly.Checked AndAlso TotalCompiled = 0 AndAlso _
               TotalProjectsBuilt = 0 AndAlso TotalLibsCompiled = 0 Then
            StatusLabel.Text = "No compiles needed"
            MessageBox.Show("No compiles needed", "IDRISMakeLib")
        Else
            StatusLabel.Text = "Total programs compiled: " + TotalCompiled.ToString + _
                               ", Total projects built: " + TotalProjectsBuilt.ToString + _
                               ", Total libraries compiled: " + TotalLibsCompiled.ToString
            MessageBox.Show("Total programs compiled: " + TotalCompiled.ToString + vbCrLf + _
                            "Total projects built: " + TotalProjectsBuilt.ToString + vbCrLf + _
                            "Total libraries compiled: " + TotalLibsCompiled.ToString, _
                            "IDRISMakeLib")
        End If
        ' --- re-enable all the controls ---
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
        Cancelled = False
        Compiling = False
        StopCompiling.Enabled = False
        CancelToolStripMenuItem.Enabled = False
        StatusLabel.Text = ""
    End Sub

    Private Sub CompileAllVolumes(ByVal FromPath As String, ByVal ToPath As String)

        Dim VolumeName As String
        Dim TempName As String
        Dim Libraries() As String
        Dim LibraryName As String
        Dim LibraryNum1 As Integer
        Dim LibraryNum2 As Integer
        Dim TempLibName As String
        ' ------------------------

        For Each VolumeName In VolumeCombo.Items
            If VolumeName = "--- ALL ---" Then Continue For
            TempName = SourcePath.Text + "\" + VolumeName
            Libraries = Directory.GetDirectories(TempName)
            ' --- sort the library entries. bubble sort, but there aren't many libraries ---
            For LibraryNum1 = 0 To Libraries.GetUpperBound(0) - 1
                For LibraryNum2 = LibraryNum1 + 1 To Libraries.GetUpperBound(0)
                    If Libraries(LibraryNum1) > Libraries(LibraryNum2) Then
                        TempLibName = Libraries(LibraryNum1)
                        Libraries(LibraryNum1) = Libraries(LibraryNum2)
                        Libraries(LibraryNum2) = TempLibName
                    End If
                Next
            Next
            ' --- compile the libraries ---
            For Each LibraryName In Libraries
                System.Windows.Forms.Application.DoEvents()
                If Cancelled Then Exit Sub
                LibraryName = LibraryName.ToUpper
                If LibraryName = "--- ALL ---" Then Continue For
                If LibraryName = "GENS" Then Continue For
                If LibraryName = "INCLUDE" Then Continue For
                LibraryName = LibraryName.Substring(LibraryName.LastIndexOf("\") + 1)
                CompileSingleLib(FromPath + "\" + VolumeName, ToPath + "\" + VolumeName, LibraryName)
            Next
        Next

    End Sub

    Private Sub CompileAllLibs(ByVal FromPath As String, ByVal ToPath As String)

        Dim LibraryName As String
        ' -----------------------

        For Each LibraryName In LibraryCombo.Items
            System.Windows.Forms.Application.DoEvents()
            If Cancelled Then Exit Sub
            LibraryName = LibraryName.ToUpper
            If LibraryName = "--- ALL ---" Then Continue For
            If LibraryName = "GENS" Then Continue For
            If LibraryName = "INCLUDE" Then Continue For
            CompileSingleLib(FromPath, ToPath, LibraryName)
        Next

    End Sub

    Private Sub CompileSingleLib(ByVal FromPath As String, _
                                 ByVal ToPath As String, _
                                 ByVal LibraryName As String)

        Dim Filename As String
        Dim ShortFilename As String
        Dim MyCadol As New Compiler
        Dim FileNum As Integer
        Dim FileFound(255) As Boolean
        Dim ObjFileName As String
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim Files() As String
        Dim BasFiles() As String
        Dim SysFiles() As String
        Dim CurrInfo As FileInfo
        Dim Answer As DialogResult
        Dim DidCompile As Boolean
        Dim LatestLibDateTime As Date
        Dim VolumeName As String
        Dim FixedLibName As String
        Dim LibraryErrors As Integer = 0
        Dim BuildProject As Boolean = False
        Dim LibraryCompiled As Boolean = False
        Dim NeedDummyProg As Boolean = False
        Dim IDRISYSPath As String = ""
        Dim RuntimePath As String = ""
        Dim TempDirName As String = ""
        Dim ProgBefore As String
        Dim ProgAfter As String
        ' --------------------------------------

        LibraryName = LibraryName.ToUpper
        If LibraryName = "--- ALL ---" Then Exit Sub
        If LibraryName = "GENS" Then Exit Sub
        If LibraryName = "INCLUDE" Then Exit Sub

        System.Windows.Forms.Application.DoEvents()

        FixedLibName = LibraryName.Replace("/"c, "_"c).Replace("%"c, "_"c)

        ' --- build output directory ---
        Try
            Directory.CreateDirectory(ToPath + "\" + LibraryName)
        Catch
        End Try

        MyCadol.FromPath = FromPath + "\" + LibraryName
        MyCadol.ToPath = ToPath + "\" + LibraryName
        MyCadol.KeepComments = CheckAddComments.Checked

        For FileNum = 0 To 255
            FileFound(FileNum) = False
        Next

        LibraryErrors = 0
        BuildProject = False
        LibraryCompiled = False
        LatestLibDateTime = #1/1/1900#

        VolumeName = FromPath.Substring(FromPath.LastIndexOf("\") + 1)
        Files = Directory.GetFiles(MyCadol.FromPath, "*.k")
        For Each Filename In Files
            System.Windows.Forms.Application.DoEvents()
            If Cancelled Then Exit Sub
            ShortFilename = Filename.Substring(Filename.LastIndexOf("\") + 1)
            StatusLabel.Text = (VolumeName + "\" + LibraryName + "\" + ShortFilename).ToUpper
            System.Windows.Forms.Application.DoEvents()
            Try
                ProgBefore = File.ReadAllText(Filename)
                ProgAfter = CadolReformat.ReformatCadolProgram(ProgBefore)
            Catch ex As Exception
                LibraryErrors += 1
                TotalErrors += 1
                CadolProg.LoadFile(MyCadol.FromPath + "\" + ShortFilename, RichTextBoxStreamType.PlainText)
                VB6Prog.Text = ShortFilename + vbCrLf + "   *** " + ex.Message + " ***"
                Answer = MessageBox.Show("Error found - Press OK to continue or Cancel to abort", _
                                         "Error found", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                If Answer = DialogResult.Cancel Then
                    Cancelled = True
                    Compiling = False
                    Exit Sub
                End If
                Continue For
            End Try
            FileNum = MyCadol.CompileToIL(ShortFilename, CheckChangedOnly.Checked, DidCompile)
            If MyCadol.ErrorCount > 0 Then
                LibraryErrors += MyCadol.ErrorCount
                TotalErrors += MyCadol.ErrorCount
                CadolProg.LoadFile(MyCadol.FromPath + "\" + ShortFilename, RichTextBoxStreamType.PlainText)
                VB6Prog.Text = MyCadol.ErrorList
                Answer = MessageBox.Show("Error found - Press OK to continue or Cancel to abort",
                                         "Error found", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                If Answer = DialogResult.Cancel Then
                    Cancelled = True
                    Compiling = False
                    Exit Sub
                End If
                Continue For
            End If
            If DidCompile Then
                TotalCompiled += 1
                LibraryCompiled = True
                CadolProg.LoadFile(MyCadol.FromPath + "\" + ShortFilename, RichTextBoxStreamType.PlainText)
                If LibraryName <> "_IDRISYS" Then
                    ObjFileName = "modProg" + FileNum.ToString.PadLeft(3, "0"c) + ".bas"
                Else
                    ObjFileName = "modSysProg" + FileNum.ToString.PadLeft(3, "0"c) + ".bas"
                End If
                VB6Prog.LoadFile(MyCadol.ToPath + "\" + ObjFileName, RichTextBoxStreamType.PlainText)
            End If
            If LatestLibDateTime < MyCadol.LatestDateTime Then
                LatestLibDateTime = MyCadol.LatestDateTime
            End If
            System.Windows.Forms.Application.DoEvents()
            FileFound(FileNum) = True
        Next

        If LibraryErrors > 0 Then Exit Sub

        If LibraryName = "_IDRISYS" Then
            BasFiles = Directory.GetFiles(MyCadol.FromPath, "*.bas")
            For Each Filename In BasFiles
                ObjFileName = Filename.Substring(Filename.LastIndexOf("\") + 1)
                If File.Exists(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                    If File.GetLastWriteTimeUtc(Filename) <= File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                        Continue For
                    End If
                    ' --- check if the file read-only ---
                    CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\" + ObjFileName)
                    Do While CurrInfo.IsReadOnly
                        Answer = MessageBox.Show("""" + ObjFileName + """ is Read-Only",
                                                 "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                        ' --- abort ---
                        If Answer = DialogResult.Abort Then
                            Cancelled = True
                            Compiling = False
                            TotalErrors += 1
                            Exit Sub
                        End If
                        ' --- ignore ---
                        If Answer = DialogResult.Ignore Then
                            TotalErrors += 1
                            Exit Sub
                        End If
                        ' --- retry ---
                        CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\" + ObjFileName)
                    Loop
                End If
                StatusLabel.Text = (VolumeName + "\" + LibraryName + "\" + ObjFileName).ToUpper
                System.Windows.Forms.Application.DoEvents()
                My.Computer.FileSystem.CopyFile(Filename, ToPath + "\" + LibraryName + "\" + ObjFileName, True)
                ' --- mark the file read-write after copy ---
                CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\" + ObjFileName)
                If CurrInfo.IsReadOnly Then
                    CurrInfo.IsReadOnly = False
                End If
            Next
            Exit Sub
        End If

        If String.Equals(GetINIValue(FromPath + "\" + LibraryName + "\KEEPOPTS.INI",
                         "Compile", "Gens"), "TRUE", StringComparison.OrdinalIgnoreCase) Then
            MyCadol.FromPath = FromPath + "\" + "GENS"
            MyCadol.ToPath = ToPath + "\" + LibraryName
            For FileNum = 0 To 255
                If Not FileFound(FileNum) Then
                    ShortFilename = FileNum.ToString.PadLeft(3, "0"c) + "GEN.K"
                    If File.Exists(MyCadol.FromPath + "\" + ShortFilename) Then
                        StatusLabel.Text = (VolumeName + "\" + LibraryName + "\" + ShortFilename).ToUpper
                        FileNum = MyCadol.CompileToIL(ShortFilename, CheckChangedOnly.Checked, DidCompile)
                        If MyCadol.ErrorCount > 0 Then
                            LibraryErrors += MyCadol.ErrorCount
                            TotalErrors += MyCadol.ErrorCount
                            CadolProg.LoadFile(MyCadol.FromPath + "\" + ShortFilename, RichTextBoxStreamType.PlainText)
                            VB6Prog.Text = MyCadol.ErrorList
                            Answer = MessageBox.Show("Error found - Press OK to continue or Cancel to abort",
                                                     "Error found", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                            If Answer = DialogResult.Cancel Then
                                Cancelled = True
                                Compiling = False
                                TotalErrors += 1
                                Exit Sub
                            End If
                            Continue For
                        End If
                        If DidCompile Then
                            TotalCompiled += 1
                            LibraryCompiled = True
                            CadolProg.LoadFile(MyCadol.FromPath + "\" + ShortFilename, RichTextBoxStreamType.PlainText)
                            VB6Prog.LoadFile(MyCadol.ToPath + "\modProg" + FileNum.ToString.PadLeft(3, "0"c) + ".bas", RichTextBoxStreamType.PlainText)
                        End If
                        If LatestLibDateTime < MyCadol.LatestDateTime Then
                            LatestLibDateTime = MyCadol.LatestDateTime
                        End If
                        System.Windows.Forms.Application.DoEvents()
                        FileFound(FileNum) = True
                    End If
                End If
            Next
        End If

        If LibraryErrors > 0 Then Exit Sub

        ' --- check compiled module files and dummy modules ---
        For FileNum = 0 To 255
            System.Windows.Forms.Application.DoEvents()
            If Cancelled Then Exit Sub
            ObjFileName = "modProg" + FileNum.ToString.PadLeft(3, "0"c) + ".bas"
            StatusLabel.Text = (VolumeName + "\" + LibraryName + "\" + ObjFileName).ToUpper
            System.Windows.Forms.Application.DoEvents()
            ' --- check if compiled module has later date ---
            If FileFound(FileNum) Then
                If File.Exists(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                    If LatestLibDateTime < File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                        LatestLibDateTime = File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\" + ObjFileName)
                    End If
                End If
                Continue For
            End If
            ' --- check if dummy program exists and is correct ---
            NeedDummyProg = True
            If File.Exists(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                NeedDummyProg = False
                sr = New StreamReader(ToPath + "\" + LibraryName + "\" + ObjFileName)
                If sr.ReadLine <> "Attribute VB_Name = ""modProg" + FileNum.ToString.PadLeft(3, "0"c) + """" Then
                    NeedDummyProg = True
                ElseIf sr.ReadLine <> "Option Explicit" Then
                    NeedDummyProg = True
                ElseIf sr.ReadLine <> "" Then
                    NeedDummyProg = True
                ElseIf sr.ReadLine <> "Public Sub PROG_" + FileNum.ToString.PadLeft(3, "0"c) + "(ByVal JUMPPOINT As Long)" Then
                    NeedDummyProg = True
                ElseIf sr.ReadLine <> "       FATALERROR ""MISSING PROGRAM NUMBER: " + FileNum.ToString.PadLeft(3, "0"c) + """" Then
                    NeedDummyProg = True
                ElseIf sr.ReadLine <> "End Sub" Then
                    NeedDummyProg = True
                End If
                sr.Close()
                If NeedDummyProg Then
                    ' --- check if the file read-only ---
                    CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\" + ObjFileName)
                    Do While CurrInfo.IsReadOnly
                        Answer = MessageBox.Show("""" + ObjFileName + """ is Read-Only",
                                                 "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                        ' --- abort ---
                        If Answer = DialogResult.Abort Then
                            Cancelled = True
                            Compiling = False
                            TotalErrors += 1
                            Exit Sub
                        End If
                        ' --- ignore ---
                        If Answer = DialogResult.Ignore Then
                            TotalErrors += 1
                            Exit Sub
                        End If
                        ' --- retry ---
                        CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\" + ObjFileName)
                    Loop
                End If
            End If
            If NeedDummyProg Then
                TotalCompiled += 1
                LibraryCompiled = True
                sw = New StreamWriter(ToPath + "\" + LibraryName + "\" + ObjFileName)
                sw.WriteLine("Attribute VB_Name = ""modProg" + FileNum.ToString.PadLeft(3, "0"c) + """")
                sw.WriteLine("Option Explicit")
                sw.WriteLine("")
                sw.WriteLine("Public Sub PROG_" + FileNum.ToString.PadLeft(3, "0"c) + "(ByVal JUMPPOINT As Long)")
                sw.WriteLine("       FATALERROR ""MISSING PROGRAM NUMBER: " + FileNum.ToString.PadLeft(3, "0"c) + """")
                sw.WriteLine("End Sub")
                sw.Close()
                CadolProg.Text = ""
                If LatestLibDateTime < File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\" + ObjFileName) Then
                    LatestLibDateTime = File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\" + ObjFileName)
                End If
                VB6Prog.LoadFile(ToPath + "\" + LibraryName + "\" + ObjFileName, RichTextBoxStreamType.PlainText)
            End If
        Next

        ' --- check if project file is out of date ---
        If File.Exists((ToPath + "\" + LibraryName + "\LIB_" + FixedLibName).ToUpper + ".vbp") Then
            ' --- get datetime of VB6 project file ---
            If LatestLibDateTime <= File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp") Then
                LatestLibDateTime = File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp")
                If LibraryCompiled Then
                    BuildProject = True
                End If
            Else
                BuildProject = True
                LibraryCompiled = True
            End If
        Else
            BuildProject = True
            LibraryCompiled = True
        End If

        ' --- if still don't need to compile, check _IDRISYS modules ---
        If Not LibraryCompiled AndAlso CheckCompileLibs.Checked Then
            If LatestLibDateTime < File.GetLastWriteTimeUtc(ToPath + "\LIB_" + FixedLibName + ".exe") Then
                IDRISYSPath = ToPath + "\..\..\DEVICE00"
                ' --- look for "_SYSVOL" ---
                IDRISYSPath += "\_SYSVOL"
                ' --- look for "_IDRISYS" ---
                IDRISYSPath += "\_IDRISYS"
                SysFiles = Directory.GetFiles(IDRISYSPath, "modSysProg*.bas")
                For Each Filename In SysFiles
                    System.Windows.Forms.Application.DoEvents()
                    If Cancelled Then Exit Sub
                    If LatestLibDateTime < File.GetLastWriteTimeUtc(Filename) Then
                        LatestLibDateTime = File.GetLastWriteTimeUtc(Filename)
                        BuildProject = True
                        LibraryCompiled = True
                    End If
                Next
                If LatestLibDateTime > File.GetLastWriteTimeUtc(ToPath + "\LIB_" + FixedLibName + ".exe") Then
                    LibraryCompiled = True
                End If
            Else
                LibraryCompiled = True
            End If
        End If

        ' --- if still don't need to compile, check runtime modules ---
        If Not LibraryCompiled AndAlso CheckCompileLibs.Checked Then
            If LatestLibDateTime < File.GetLastWriteTimeUtc(ToPath + "\LIB_" + FixedLibName + ".exe") Then
                RuntimePath = ToPath + "\..\..\COMMON"
                SysFiles = Directory.GetFiles(RuntimePath, "rt*.*")
                For Each Filename In SysFiles
                    System.Windows.Forms.Application.DoEvents()
                    If Cancelled Then Exit Sub
                    If LatestLibDateTime < File.GetLastWriteTimeUtc(Filename) Then
                        LatestLibDateTime = File.GetLastWriteTimeUtc(Filename)
                        BuildProject = True
                        LibraryCompiled = True
                    End If
                Next
                If LatestLibDateTime > File.GetLastWriteTimeUtc(ToPath + "\LIB_" + FixedLibName + ".exe") Then
                    BuildProject = True
                    LibraryCompiled = True
                End If
            Else
                BuildProject = True
                LibraryCompiled = True
            End If
        End If

        ' --- check if project file is read-only ---
        System.Windows.Forms.Application.DoEvents()
        If Cancelled Then Exit Sub
        If File.Exists(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp") Then
            ' --- get datetime of VB6 project file ---
            If BuildProject OrElse
               LatestLibDateTime > File.GetLastWriteTimeUtc(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp") Then
                ' --- check if the file read-only ---
                CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp")
                Do While CurrInfo.IsReadOnly
                    Answer = MessageBox.Show("""LIB_" + FixedLibName + ".vbp"" is Read-Only",
                                             "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                    ' --- abort ---
                    If Answer = DialogResult.Abort Then
                        Cancelled = True
                        Compiling = False
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- ignore ---
                    If Answer = DialogResult.Ignore Then
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- retry ---
                    CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp")
                Loop
                BuildProject = True
                LibraryCompiled = True
            End If
        Else
            BuildProject = True
            LibraryCompiled = True
        End If

        If Not LibraryCompiled Then Exit Sub

        If BuildProject Then
            StatusLabel.Text = "Building Project " + (VolumeName + "\" + LibraryName + "\LIB_" + FixedLibName).ToUpper + ".vbp"
            System.Windows.Forms.Application.DoEvents()
            MyCadol.BuildVB6ProjectFile(ToPath, LibraryName, CommonPath)
            TotalProjectsBuilt += 1
        End If

        System.Windows.Forms.Application.DoEvents()
        If Cancelled Then Exit Sub
        If CheckCompileLibs.Checked Then
            If File.Exists(ToPath + "\LIB_" + FixedLibName + ".exe") Then
                ' --- check if the file read-only ---
                CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\LIB_" + FixedLibName + ".exe")
                Do While CurrInfo.IsReadOnly
                    Answer = MessageBox.Show("""LIB_" + FixedLibName + ".exe"" is Read-Only",
                                             "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                    ' --- abort ---
                    If Answer = DialogResult.Abort Then
                        Cancelled = True
                        Compiling = False
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- ignore ---
                    If Answer = DialogResult.Ignore Then
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- retry ---
                    CurrInfo = My.Computer.FileSystem.GetFileInfo(ToPath + "\LIB_" + FixedLibName + ".exe")
                Loop
            End If
            ' --- delete EXE file before compiling ---
            Do While File.Exists(ToPath + "\LIB_" + FixedLibName + ".exe")
                Try
                    File.Delete(ToPath + "\LIB_" + FixedLibName + ".exe")
                Catch ex As Exception
                    Answer = MessageBox.Show("""LIB_" + FixedLibName + ".exe"" is in use and cannot be compiled",
                                             "File in use", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                    ' --- abort ---
                    If Answer = DialogResult.Abort Then
                        Cancelled = True
                        Compiling = False
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- ignore ---
                    If Answer = DialogResult.Ignore Then
                        TotalErrors += 1
                        Exit Sub
                    End If
                End Try
            Loop
            ' --- create a temporary subdirectory ---
            Try
                If Not Directory.Exists(LocalCompilePath) Then
                    Directory.CreateDirectory(LocalCompilePath)
                End If
                TempDirName = LocalCompilePath + "\" + Now.Ticks.ToString
                Do While Directory.Exists(TempDirName)
                    Thread.Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                    TempDirName = LocalCompilePath + "\" + Now.Ticks.ToString
                Loop
                Directory.CreateDirectory(TempDirName)
            Catch ex As Exception
                Answer = MessageBox.Show("Cannot create temporary directory",
                                         "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                ' --- abort ---
                If Answer = DialogResult.Abort Then
                    Cancelled = True
                    Compiling = False
                    TotalErrors += 1
                    Exit Sub
                End If
                ' --- ignore ---
                If Answer = DialogResult.Ignore Then
                    TotalErrors += 1
                    Exit Sub
                End If
            End Try
            ' --- compile EXE file ---
            StatusLabel.Text = "Compiling " + (VolumeName + "\LIB_" + FixedLibName).ToUpper + ".exe"
            System.Windows.Forms.Application.DoEvents()
            ' --- this command will take a while to run ---
            Dim RunCmd As String = """" + VB6EXE + """ " +
                                   "/make """ + ToPath + "\" + LibraryName + "\LIB_" + FixedLibName + ".vbp"" " +
                                   "/outdir """ + TempDirName + """"
            ' --- Build a batch file to run the command - needed for VB6EXE to work ---
            Dim BatchFileName As String = TempDirName + "\BuildLib.bat"
            File.WriteAllText(BatchFileName, RunCmd + vbCrLf)
            ' --- Run job in the background ---
            BackgroundWorker1.RunWorkerAsync(BatchFileName)
            Do While BackgroundWorker1.IsBusy
                ' --- prevent excess CPU usage while waiting ---
                Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop
            ' --- Delete batch file when done ---
            File.Delete(BatchFileName)
            ' --- copy file from temporary directory to actual directory ---
            If File.Exists(TempDirName + "\LIB_" + FixedLibName + ".exe") Then
                Try
                    File.Move(TempDirName + "\LIB_" + FixedLibName + ".exe", ToPath + "\LIB_" + FixedLibName + ".exe")
                Catch ex As Exception
                    Answer = MessageBox.Show("""LIB_" + FixedLibName + ".exe"" cannot be moved",
                                             "Move Failed", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                    ' --- abort ---
                    If Answer = DialogResult.Abort Then
                        Cancelled = True
                        Compiling = False
                        TotalErrors += 1
                        Exit Sub
                    End If
                    ' --- ignore ---
                    If Answer = DialogResult.Ignore Then
                        TotalErrors += 1
                        Exit Sub
                    End If
                End Try
            End If
            Try
                Directory.Delete(TempDirName)
            Catch ex As Exception
                ' --- not fatal if can't delete ---
            End Try
        End If

    End Sub

    Private Sub FillVolumeCombo()
        Dim Volumes() As String
        Dim TempName As String
        ' ---------------------
        VolumeCombo.Items.Clear()
        VolumeCombo.Items.Add("--- ALL ---")
        If SourcePath.Text <> "" Then
            Volumes = Directory.GetDirectories(SourcePath.Text)
            For Each TempName In Volumes
                TempName = TempName.Substring(TempName.LastIndexOf("\") + 1).ToUpper
                VolumeCombo.Items.Add(TempName)
            Next
            If My.Settings.LastVolume <> "" Then
                VolumeCombo.Text = My.Settings.LastVolume
            Else
                VolumeCombo.SelectedIndex = 0
            End If
        Else
            VolumeCombo.SelectedIndex = 0
        End If
    End Sub

    Private Sub EnvCombo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnvCombo.SelectedIndexChanged
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim TempPath As String
        ' --------------------
        Select Case EnvCombo.Text.ToUpper
            Case "PC"
                INIFilename = My.Settings.INIPathPC.Replace("*", GetUserName())
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "LOCAL"
                INIFilename = My.Settings.INIPathLocal.Replace("*", GetUserName())
                If Not File.Exists(INIFilename) Then
                    INIFilename = My.Settings.INIPathLocalAlt.Replace("*", GetUserName())
                End If
                If Not File.Exists(INIFilename) Then
                    INIFilename = My.Settings.INIPathLocalAlt2.Replace("*", GetUserName())
                End If
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "TEST"
                INIFilename = My.Settings.INIPathTest
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "ACCEPT"
                INIFilename = My.Settings.INIPathAccept
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "PROD"
                INIFilename = My.Settings.INIPathProd
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "FIS"
                INIFilename = My.Settings.INIPathFIS2
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "FISTEST"
                INIFilename = My.Settings.INIPathFISTest
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "EOY"
                INIFilename = My.Settings.INIPathEOY
                My.Settings.LastEnv = EnvCombo.Text
                My.Settings.Save()
            Case "<BROWSE>"
                If BrowseEnvironment.ShowDialog() = DialogResult.OK Then
                    INIFilename = BrowseEnvironment.FileName
                Else
                    INIFilename = ""
                End If
            Case Else
                INIFilename = ""
        End Select
        If INIFilename = "" Then
            SourcePath.Text = ""
            TargetPath.Text = ""
            CommonPath = ""
            VB6EXE = ""
            Exit Sub
        End If
        Try
            ' --- check if local environment is setup ---
            If Not File.Exists(INIFilename) Then
                If EnvCombo.Text.ToUpper = "LOCAL" Then
                    MessageBox.Show("Your PC is not setup with a local IDRIS environment")
                Else
                    MessageBox.Show("INI File not found: " + INIFilename)
                End If
                Exit Sub
            End If
            ' --- get source directory path ---
            TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "FromDir").ToUpper
            If TempPath.StartsWith("\\" + Environment.GetEnvironmentVariable("COMPUTERNAME").ToUpper + "\") Then
                TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "FromDirLocal").ToUpper
                If TempPath = "" Then
                    TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "FromDir").ToUpper
                End If
            End If
            If TempPath.LastIndexOf("\") > TempPath.LastIndexOf("\DEVICE") Then
                TempPath = TempPath.Substring(0, TempPath.LastIndexOf("\DEVICE") + 9)
            End If
            If TempPath = "" OrElse Not Directory.Exists(TempPath) Then
                MessageBox.Show("FromDir not specified or not found: " + TempPath)
                Exit Sub
            End If
            SourcePath.Text = TempPath
            ' --- get target directory path ---
            TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "ToDir").ToUpper
            If TempPath.StartsWith("\\" + Environment.GetEnvironmentVariable("COMPUTERNAME").ToUpper + "\") Then
                TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "ToDirLocal").ToUpper
                If TempPath = "" Then
                    TempPath = GetINIValue(INIFilename, "IDRISMakeLib", "ToDir").ToUpper
                End If
            End If
            If TempPath.LastIndexOf("\") > TempPath.LastIndexOf("\DEVICE") Then
                TempPath = TempPath.Substring(0, TempPath.LastIndexOf("\DEVICE") + 9)
            End If
            If TempPath = "" OrElse Not Directory.Exists(TempPath) Then
                ' --- build output directory ---
                Try
                    Directory.CreateDirectory(TempPath)
                Catch
                    MessageBox.Show("ToDir not specified or not found: " + TempPath)
                    Exit Sub
                End Try
            End If
            TargetPath.Text = TempPath
            ' --- get common directory path ---
            CommonPath = GetINIValue(INIFilename, "IDRISMakeLib", "CommonDir").ToUpper
            If CommonPath.StartsWith("\\" + Environment.GetEnvironmentVariable("COMPUTERNAME").ToUpper + "\") Then
                CommonPath = GetINIValue(INIFilename, "IDRISMakeLib", "CommonDirLocal").ToUpper
                If CommonPath = "" Then
                    CommonPath = GetINIValue(INIFilename, "IDRISMakeLib", "CommonDir").ToUpper
                End If
            End If
            If CommonPath = "" OrElse Not Directory.Exists(CommonPath) Then
                MessageBox.Show("CommonPath not specified or not found: " + CommonPath)
                Exit Sub
            End If
            ' --- get VB6 Compiler path and executable name ---
            VB6EXE = GetINIValue(INIFilename, "IDRISMakeLib", "VB6").ToUpper
            ' --- adjust for 64-bit INI file run from a 32-bit system ---
            If VB6EXE <> "" AndAlso Not File.Exists(VB6EXE) Then
                VB6EXE = VB6EXE.Replace("\PROGRAM FILES (X86)\", "\PROGRAM FILES\")
            End If
            ' --- adjust for 32-bit INI file run from a 64-bit system ---
            If VB6EXE <> "" AndAlso Not File.Exists(VB6EXE) Then
                VB6EXE = VB6EXE.Replace("\PROGRAM FILES\", "\PROGRAM FILES (X86)\")
            End If
            ' --- check if vb6 compiler exists ---
            If VB6EXE = "" OrElse Not File.Exists(VB6EXE) Then
                MessageBox.Show("VB6EXE not specified or not found: " + VB6EXE + vbCrLf + _
                                "Programs will not be compiled into IDRIS Libraries.", _
                                "Can't compile IDRIS Libraries", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                CheckCompileLibs.Checked = False
                CheckCompileLibs.Enabled = False
            Else
                CheckCompileLibs.Enabled = True
            End If
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Unknown error occurred trying to open INI file" + vbCrLf + ex.Message)
        End Try
        FillVolumeCombo()
    End Sub

    Private Sub StopCompiling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                Handles StopCompiling.Click, CancelToolStripMenuItem.Click
        Cancelled = True
        Compiling = False
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutMain.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Cancelled = True
        If Compiling Then
            Compiling = False
            Exit Sub
        End If
        Me.Close()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Try
            Shell(CStr(e.Argument), AppWinStyle.Hide, True)
            e.Result = "" ' ok
            TotalLibsCompiled += 1
        Catch ex As Exception
            e.Result = ex.Message
            worker.CancelAsync()
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ' First, handle the case where an exception was thrown.
        If (e.Error IsNot Nothing) Then
            MessageBox.Show(e.Error.Message)
        ElseIf e.Result.ToString <> "" Then
            MessageBox.Show(e.Result.ToString)
        End If
    End Sub

    Private Sub LibraryCombo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LibraryCombo.SelectedIndexChanged
        If LibraryCombo.SelectedIndex <= 0 Then
            My.Settings.LastLibrary = ""
        Else
            My.Settings.LastLibrary = LibraryCombo.Text
        End If
        My.Settings.Save()
    End Sub

End Class
