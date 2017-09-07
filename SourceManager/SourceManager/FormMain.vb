' --------------------------------
' --- FormMain.vb - 09/07/2017 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/07/2017 - SBakker
'            - Don't ignore ".gitignore" files.
' 07/31/2017 - SBakker
'            - Ignore "testresults" directories.
'            - Handle "arena2" the same as "arena".
' 07/24/2017 - SBakker
'            - Stop ignoring "packages" folder. Want NuGet in Arena now.
'            - Include ".dll" files in "...\packages\..." folders.
' 06/26/2017 - SBakker
'            - Ignore folder "packages". Can't easily be copied.
' 05/12/2017 - SBakker
'            - Make sure DoCancel is set to True when closing form, so compare stops running.
' 02/21/2017 - SBakker
'            - Added ".frx" to KnownBinaryFile().
' 01/20/2017 - SBakker
'            - Ignore ".jfm" database project cache files.
' 12/21/2016 - SBakker
'            - Ignore ".dbmdl" database project cache files.
' 12/19/2016 - SBakker
'            - Added ".pfx" files to be handled under the "Project Differences" tab.
'            - Removed extra .ToLower modifiers on already-lowered strings.
' 09/22/2016 - SBakker
'            - Added SQL project files to ProjectFile() function.
' 09/06/2016 - SBakker
'            - Fixing problems with saving ExcludedFileList.
' 08/22/2016 - SBakker
'            - Removed special handling for CodeSigning.pfx files. They are now handled the same as
'              other known binary files.
' 07/20/2016 - SBakker
'            - Removed all special handling of environments, drives, paths, and thumbprints.
'            - Added C# project and properties files to ProjectFile() function.
' 06/30/2016 - SBakker
'            - Added VVBackupForce.bat to ignore list.
' 06/23/2016 - SBakker
'            - Added setting ExcludeFileList so that any desired files can be excluded.
' 05/25/2016 - SBakker
'            - Don't run External Compare program modally, let SourceManager continue on.
' 05/12/2016 - SBakker
'            - Ignore VVBackup.bat and VVCompare.bat in comparisons.
' 02/16/2016 - SBakker
'            - Fixed some issues with External Compare.
' 02/09/2016 - SBakker
'            - Disable ComboFromDir and ComboToDir until an Application is selected.
' 02/01/2016 - SBakker
'            - Ignore VaultBackup.bat and VaultCompare.bat in comparisons.
' 01/11/2016 - SBakker
'            - Added option to ignore PFX files. They aren't tracked by Git.
'            - Fixed to be compatable with .NET 4.5.2. There were some reference issues.
' 12/31/2015 - SBakker
'            - After deleting files, also remove all empty parent directories.
' 12/01/2015 - SBakker
'            - After deleting files, remove empty directories.
' 10/21/2015 - SBakker
'            - Don't show empty directories on InFromDirOnly or InToDirOnly.
' 10/08/2015 - SBakker
'            - Changed so the C: drive uses the environment "PC" instead of "Local". Avoids confusion
'              when the Y: drive is also "Local" and working in the office.
' 07/02/2015 - SBakker
'            - Ignore files currently open by another process which has them locked, such as MS Word.
' 05/02/2015 - SBakker
'            - Handle situation where a top-level path (C:\, \\server\share\) is entered. They have no
'              parent directory, so ThisDir.Parent Is Nothing and ThisDir.Parent.Name fails.
' 04/22/2015 - SBakker
'            - Compare normal files as binary first. This will go very quickly since most files should
'              match anyway. Nothing else has been changed.
' 04/13/2015 - SBakker
'            - Added more binary file extensions.
' 03/06/2015 - SBakker
'            - Ignore files and directories that start with ".".
' 01/15/2015 - SBakker
'            - Changed ComboApplication combo box to be a Drop Down List.
'            - Added handling of command line parameters: Application FromDir ToDir. Immediately click
'              ButtonCompare if all three parameters are OK.
' 12/19/2014 - SBakker
'            - Don't fill ComboFromDir/ComboToDir boxes from Last settings if the directories don't
'              exist.
' 10/08/2014 - SBakker
'            - Ignore known configuration files, which have a filename starting with ".".
' 08/29/2014 - SBakker
'            - Added new "clearuserprograms.bat" file to ignore file list.
'            - Removed "CleanAll.bat" from list of special files.
' 08/22/2014 - SBakker
'            - Ignore certain directories which are created automatically within projects, but may or
'              may not have any files in them.
' 08/19/2014 - SBakker
'            - Added My.Settings.IgnoreMissingDirectories to ignore the contents of directories which
'              are missing on the From or To side. They will be shown as "C:\Directory\..." in the
'              list, and no files or subdirectories will be shown. This is to simplify comparing
'              when there are missing directories on one side or the other with many files.
' 08/12/2014 - SBakker
'            - Disable "Compare" button when ComboFromDir.Text = ComboToDir.Text or either is blank.
' 06/30/2014 - SBakker
'            - Remmove value for My.Settings.ExternalCompareApp if the program doesn't exist.
'            - Cleared default value for My.Settings.ExternalCompareApp, and instead search for known
'              locations of the WinMerge program. Fill in if found.
' 06/23/2014 - SBakker
'            - Added new "clearusersettings.bat" file to ignore file list.
' 05/28/2014 - SBakker
'            - Fixed spelling of "Canceled".
' 05/23/2014 - SBakker
'            - Added SkyDrive/OneDrive to have the same handling as Dropbox.
' 05/22/2014 - SBakker
'            - Removed error for missing <PlatformTarget> lines in *.vbproj files. Old source doesn't
'              have them.
' 05/05/2014 - SBakker
'            - Call ListFilesTo() for ToDir, instead of ListFilesFrom().
' 05/04/2014 - SBakker
'            - Fixed checking of non-VBNET files non-binary. They got handed off to the size and date
'              routines, and never used the compare options. Ha! Been chasing this one down for a long
'              time!
' 04/30/2014 - SBakker
'            - Ignore hidden directories, even the ones provided.
' 04/15/2014 - SBakker
'            - Only include ".settings" files in the "Applications" directory.
' 04/09/2014 - SBakker
'            - Added VB6's "*.vbp" files to the list of ProjectFile() so they will get added to the
'              ListBoxProjDiff instead of ListBoxDiff.
' 03/24/2014 - SBakker
'            - Added special handling for "Arena\Bin" and "Utilities\Bin", so that they don't copy
'              files, but do check subdirectories. Cory's new folder of "MenuIcons" needs to be
'              copied, and any other new subdirectories, but not the executable files in "Arena\Bin".
' 03/17/2014 - SBakker
'            - Removed special handling just for Arena_Finance. Changed to be N:\Finance\Arena, which
'              helps everything work better.
' 03/14/2014 - SBakker
'            - Use new generic BinaryCompareClass.BinaryFilesMatch() rather than having one here.
'            - Added "CleanAll.bat" to list of special files.
' 02/25/2014 - SBakker
'            - Added option to include "TEST_" projects.
' 02/24/2014 - SBakker
'            - Added Bootstrap loading all programs to another location, and then running from there.
' 02/18/2014 - SBakker
'            - Added ToolStripMenuItemCompare so that F5 can be bound to the same code as clicking
'              "Compare".
' 02/11/2014 - SBakker
'            - Changed "<DebugType>full</DebugType>" to "<DebugType>Full</DebugType>" (case change).
' 02/03/2014 - SBakker
'            - Don't ignore "zip" and "zipx" files while comparing Templates.
'            - Only check for missing PlatformTarget lines in ".vbproj" files.
' 01/31/2014 - SBakker
'            - Changed PlatformTarget = AnyCPU to PlatformTarget = x86 during file copy. Using "x86"
'              makes programs load faster!
' 01/28/2014 - SBakker
'            - Exclude UDL files from comparisons.
' 12/09/2013 - SBakker
'            - Added handling for Quick Binary Compare, that will only check the file sizes and dates.
'            - Added more binary file types.
' 12/05/2013 - SBakker - URD 12229
'            - Added special handling for the new AcceptDN environment.
' 11/11/2013 - SBakker
'            - Ignore <ProductName> with "Microsoft" and "Report Viewer" when adding environments.
' 10/23/2013 - SBakker
'            - Switch to Arena versions of ConfigInfo, DataConn, and Utilities.
' 10/18/2013 - SBakker
'            - Ignore directories that start with "Test_". Those will only stay on the machine where
'              they are created.
' 10/15/2013 - SBakker
'            - Don't add duplicate items to ComboFromDir or ComboToDir.
'            - Make sure to add either "Arena\" or "Arena_Finance\" to the <PublishURL> path whenever
'              the Finance environment is involved. Had to add "FromPath" as a parameter to FixFile so
'              this information would be known.
'            - Added "Report" environment.
' 09/25/2013 - SBakker
'            - Added special handling for the new Finance environment.
' 04/01/2013 - SBakker
'            - Don't remove specific version info from Microsoft references.
'            - Put project files onto Project tab, not Differences tab.
' 02/27/2013 - SBakker
'            - Removed unused local variable.
' 02/20/2013 - SBakker
'            - Added StatusLine messages instead of MessageBoxes for any "Completed" msgs.
'              They are annoying to have to find and click on.
'            - Added checking that the file to be deleted is Read-Only. Allow Abort, Retry,
'              and Ignore.
' 10/09/2012 - SBakker
'            - Added KnownBinaryFiles() to prevent running compares that will never work,
'              and thus prevent skipping files that should be compared by time/date/size.
' 08/24/2012 - SBakker
'            - Make sure to set all the properties of MyFileCompare before calling FixFile.
'              Otherwise the VBNET setting isn't turned on properly.
' 02/15/2012 - SBakker
'            - Fixed differences between <PublishUrl> needing "%24" and all others needing
'              "$" in any UNC path like "\\localhost\c$\".
' 02/13/2012 - SBakker
'            - Fixed FileCompareVBNET and FindReplace sections to handle local computer name
'              replacing "LOCALHOST". Now it will work with any computer!
' 01/12/2012 - SBakker
'            - Added FormExternalCompare and the new settings UseExternalCompare and
'              ExternalCompareApp. This lets an outside program be used to compare the files
'              instead of the built-in file compare.
' 09/09/2011 - SBakker
'            - Added checks for "CurrLine IsNot Nothing AndAlso" in FixFile to prevent some
'              "Object not set" errors.
' 08/31/2011 - SBakker
'            - Force OptionStrict = On and OptionInfer = Off. This isn't getting done in
'              new projects!
' 06/28/2011 - SBakker
'            - Check for EOL errors, with either CRCR or LFLF patterns. These cause issues
'              in a variety of places, including VSS file compares. Now it shows these as
'              different, and fixes the target file after copying.
'            - Fixed logic for IgnoreAll, YesToAll, NoToAll to still update the lists of
'              files actually copied. This has never really worked right, and might still be
'              incorrect, but it looks better.
' 06/21/2011 - SBakker
'            - Removed checking for "_#ENV". Caused issues in unexpected places.
' 06/15/2011 - SBakker
'            - Added checking for "_#ENV" to go along with " - #ENV#".
' 06/05/2011 - SBakker
'            - Added ".zipx" files to IgnoreFile list.
' 05/23/2011 - SBakker
'            - Added OptionIgnoreVersions setting, for showing only actual changes to
'              project and assembly files.
' 05/20/2011 - SBakker
'            - Added ".log" files to IgnoreFile list.
' 05/19/2011 - SBakker
'            - Fixed issue where "File is read-only: Ignore" was getting treated like they
'              selected "Abort".
'            - Fixed issue where "File is read-only: Abort" wasn't redisplaying the list of
'              files remaining (wasn't removing the ones copied).
'            - Added ".TabsToSpaces" and ".TrimBlanks" to the list of things controlled by
'              My.Settings.OptionIgnoreSpaces.
'            - Disable Copy and Delete while comparing. Still allow ShowDiffs and ViewFile.
'            - Moved Enum OverwriteResult into FormMain.vb.
' 04/13/2011 - SBakker
'            - Replace DebugType = pdb-only with None, so Release project compiles will have
'              no debugging info and can then be optimized.
' 04/07/2011 - SBakker
'            - Made everything that's not Local, Test, Accept, or Prod be a PC.
'            - Fixed enabling of buttons to happen on the first difference found. Also made
'              them not enable unless they could be used, like not enabling Copy if no files
'              are selected.
' 04/05/2011 - SBakker
'            - Fixed issue with replacing " - #ENV#" from Production. Now it's not a one-way
'              conversion. Also make sure " - Prod" gets removed if found.
'            - Disable the three combo boxes during comparison. Changing one would crash the
'              program!
'            - Ignore .NET Framework product names.
' 03/29/2011 - SBakker
'            - Don't replace " - Local", etc, with " - Prod", instead replace with blank.
'              This makes conversions one-way, but is nicer for Click-Once applications.
' 03/18/2011 - SBakker
'            - Added FormTargetNewer to handle responses like NoToAll and YesToAll.
'            - Fixed removal of all filenames actually copied or deleted to work correctly.
' 02/08/2011 - SBakker
'            - Added File Length to information shown on FormDiffInfo.
' 01/20/2011 - SBakker
'            - Added checks for PublicKeyTokens during file comparisons.
' 01/19/2011 - SBakker
'            - Added ".bak" and ".sav" to files to be ignored during comparison.
' 01/13/2011 - SBakker
'            - Added use of new FirstDiffOnly property in the FileCompareDataClass. This is
'              quicker, as it stops as soon as it knows any differences exist. Only used now
'              for special files.
'            - Added some known binary, temporary, or work files to the list of files to be
'              ignored.
'            - Added check to see if CurrLine is Nothing. File.ReadAllLines() can return a
'              Nothing line in some cases, usually at the end of the file.
' 01/12/2011 - SBakker
'            - Added menu option that allows adding new application names.
'            - Changed all three combo boxes to be sorted.
'            - Removed IDRIS_Programs and IDRIS_VBNET from list. I'm the only one who would
'              use them.
' 01/07/2011 - SBakker
'            - Switched to using FileCompareVBNET which inherits from FileCompareClass.
'            - Use UtilitiesDataClass.FileUtils.GetFileEncoding() to determine the file's
'              encoding, to prevent character translation issues either reading or writing.
' 01/05/2011 - SBakker
'            - Added option to ignore spaces and blank lines while comparing.
' 12/29/2010 - SBakker
'            - Ignore hidden and system files while comparing directories.
' 12/14/2010 - SBakker
'            - Only add directories to dropdowns that actually exist for the current user or
'              computer.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 11/16/2010 - SBakker
'            - Ignore BuildAll.bat and BuildTest.bat in comparisons.
' 11/15/2010 - SBakker
'            - Fixed comparison of LocalPath to properly uppercase the username and to allow
'              for an adjusted username (My.Settings.AltUserName).
'            - Ignore PublishAll.bat in comparisons.
' 11/10/2010 - SBakker
'            - Added error checking that the directories entered/selected actually exist.
' 10/19/2010 - SBakker
'            - Use whatever is defined in CFileCompare.ExcludeLines to determine which
'              lines must not be altered. Avoids mismatches between these two programs.
'            - Stop disabling everything during compare. Views and copies should still work
'              even though the compare continues.
' 10/06/2010 - SBakker
'            - Added "IDRIS_IDE" as a special project. "IDIRSMakeLib" was already one.
' 10/01/2010 - SBakker
'            - Clear the StatusLabelCounts whenever any of the combo boxes change.
' 09/30/2010 - SBakker
'            - Moved the Applications and Paths into Settings. Added a user setting for
'              manually-added paths, so they only need to be entered once.
'            - Set the cursor to WaitCursor while comparing files.
' 09/15/2010 - SBakker
'            - Removed "Arena_Utils" from the lists. They are all in "Utilities" now.
'            - Updated icon to be something new.
'            - Added "IDRIS_Programs" to list of applications.
' 07/22/2010 - SBakker
'            - Remove extra information after "<Reference Include=" project names. It isn't
'              needed in VS 2010, and may be out of date or incorrect. Also remove the
'              <SpecificVersion> "False" lines, and squish out removed lines when saving.
' 07/16/2010 - SBakker
'            - Added CommonFuncts to hold routines used in multiple classes.
' 07/15/2010 - SBakker
'            - Ignore "c:\windows" and "c:\temp" directories when comparing.
' 07/09/2010 - SBakker
'            - Fixed to include some files which were getting excluded.
' 07/08/2010 - SBakker
'            - Removed "ProductionOnly" function. All directories will now be
'              environment-level, and Production will only compile from P:.
'            - Added "Utilities" to the dropdown, with all the proper paths.
'            - Added check for PublicKeyTokens. Wasn't getting updated during the
'              copy. However, this is less necessary in VS 2010, as the current
'              info is read directly from the DLL instead of the project file.
' 07/02/2010 - SBakker
'            - Added N:\Arena_Scripts and P:\Arena_Scripts into dropdown lists.
' 06/22/2010 - SBakker
'            - Fixed so that PC drive and path get replaced with Local Drive and
'              path. All programs will be published from Local, not from C.
' 06/17/2010 - SBakker
'            - Changed from "C$" to just "C" for the localhost path. The "$"
'              causes issues with XML files. There shouldn't be any publishing on
'              the C drive anyway.
'            - Exclude this program from any environmental changes!!! (Duh!)
' 06/09/2010 - SBakker
'            - Fixed LocalPath as a replace parameter to subsitutute the username.
' 06/01/2010 - SBakker
'            - Exclude Visual Studio 2010 "*.vssscc" and ".vspscc" files.
'            - Switched back to drive-for-drive and path-for-path. Cannot publish
'              anymore to a network drive, only a network path.
' 05/24/2010 - SBakker
'            - Added "V:\" Production SourceSafe path into all compare and replace
'              logic.
' 05/17/2010 - SBakker
'            - Switch to using drive names instead of path names. Was causing
'              problems.
'            - Added "PublishAll.bat" as a special file.
' 05/03/2010 - SBakker
'            - Ignore directory BuildProcessTemplates. This comes from TS 2010.
' 04/27/2010 - SBakker
'            - Ignore any Arena_Utils projects when copying between environments
'              and converting settings. They are all Production routines and
'              shouldn't be altered.
'            - Added options for IDRIS_VB6 and IDRIS_VB.NET.
' 04/06/2010 - SBakker
'            - Added proper handling of using the PC's C: drive for source.
' 02/05/2010 - SBakker
'            - Show the number of differences in the status bar, not in a popup at
'              the end. This helps make it faster when comparing.
' 01/25/2010 - SBakker
'            - Added a tab for Project Differences. Makes it easier to tell when
'              a file has really changed, or just when a version number changed.
' 01/14/2010 - SBakker
'            - Make sure the Code Signing Files are in the "\Security\" directory.
' 12/28/2009 - SBakker
'            - Ignore ".cache" files.
' 12/17/2009 - SBakker
'            - Added IDRIS to list of applications.
'            - Added extra paths to the C: drive.
' 11/30/2009 - SBakker
'            - Added a DiffCount variable and message at the end.
'            - Compare CodeSigning files, but with their own proper local version.
' 11/19/2009 - SBakker
'            - Check if target file is newer and ask if they want to overwrite it.
'            - Added FormDiffInfo to show newer/older and actual dates.
'            - Added "View File" ability to FromDirOnly and ToDirOnly files.
'            - Allow any value to be typed into ComboFromDir and ComboToDir. Now
'              it is generic and can be used for any two directories. Yeah!
'            - Added Cancel button in case the wrong comparisons are being made.
' 10/08/2009 - SBakker
'            - Fixed to include files in the top level directory.
'            - Added "Arena_Scripts" to list of applications.
' 10/06/2009 - SBakker
'            - Build target directories if they don't exist, before copying files.
'            - Added "C:\Arena" to list for "Arena" comparisons.
'            - Don't disable controls for invalid selections.
'            - Added copying of files to the C:\ drive.
' 10/02/2009 - SBakker
'            - Added File Copy (From -> To and To -> From).
'            - Added fixing target file after it is done copying.
'            - Use special checking for files which might have environment diffs.
'              This is slower, so not used for all files.
'            - Added Application, so many can be added to this same manager.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FileUtils
Imports Arena_Utilities.StringUtils
Imports Arena_Utilities.SystemUtils
Imports FileCompareDataClass
Imports System.IO
Imports System.Text

Public Enum OverwriteResult
    Unknown
    No
    NoToAll
    Yes
    YesToAll
    Abort
    Retry
    Ignore
    IgnoreAll
End Enum

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private DirCount As Integer = 0
    Private FileCount As Integer = 0
    Private DiffCount As Integer = 0
    Private CurrEnabled As Boolean = True
    Private DoCancel As Boolean = True

    Private AppPathList As New List(Of String)
    Private AppPathExtra As New List(Of String)

    Private MyFileCompare As New FileCompareVBNET

    Private Sub FormMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

        ' --- Get settings from previous version ---
        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
            My.Settings.Save()
        End If

        If Not String.IsNullOrWhiteSpace(My.Settings.ExternalCompareApp) Then
            If My.Settings.ExternalCompareApp.ToUpper.Contains("C:\PROGRAM FILES (X86)\WINMERGE\WINMERGEU.EXE") AndAlso
                Not File.Exists("C:\PROGRAM FILES (X86)\WINMERGE\WINMERGEU.EXE") Then
                My.Settings.UseExternalCompare = False
                My.Settings.Save()
            ElseIf My.Settings.ExternalCompareApp.ToUpper.Contains("S:\WINMERGE\WINMERGEU.EXE") AndAlso
                Not File.Exists("S:\WINMERGE\WINMERGEU.EXE") Then
                My.Settings.UseExternalCompare = False
                My.Settings.Save()
            End If
        End If

        If String.IsNullOrWhiteSpace(My.Settings.ExternalCompareApp) Then
            If File.Exists("C:\Program Files (x86)\WinMerge\WinMergeU.exe") Then
                My.Settings.ExternalCompareApp = """C:\Program Files (x86)\WinMerge\WinMergeU.exe"" /e /u /maximize"
                My.Settings.Save()
            ElseIf File.Exists("S:\WinMerge\WinMergeU.exe") Then
                My.Settings.ExternalCompareApp = """S:\WinMerge\WinMergeU.exe"" /e /u /maximize"
                My.Settings.Save()
            End If
        End If

        ' --- initialization code here ---
        DirCount = 0
        FileCount = 0
        DiffCount = 0
        ShowStatus()
        MarkEnabled(True)
        ComboFromDir.Enabled = False
        ComboToDir.Enabled = False

        FillComboApplication()

        ' --- Check Command-line arguments ---
        Dim CurrArg As String
        Dim CmdLineApplication As String = Nothing
        Dim CmdLineFromDir As String = Nothing
        Dim CmdLineToDir As String = Nothing

        For i As Integer = 0 To CmdLineArgs.Count - 1
            CurrArg = CmdLineArgs.Arg(i)
            If CurrArg.StartsWith("/") OrElse CurrArg.StartsWith("-") Then
                ' '' --- Options ---
                ''Select Case CurrArg.Substring(1).ToUpper
                ''    Case "S"
                ''        DoSubdirs = True
                ''    Case "I"
                ''        IgnoreCase = True
                ''    Case "R"
                ''        UseRegEx = True
                ''    Case "Q"
                ''        QuietMode = True
                ''    Case "V"
                ''        VerboseMode = True
                ''    Case "UTF8"
                ''        OutputUTF8 = True
                ''    Case "ASCII"
                ''        OutputASCII = True
                ''    Case Else
                ''        ShowSyntax()
                ''        Exit Sub
                ''End Select
                Continue For
            End If
            If CmdLineApplication Is Nothing Then
                CmdLineApplication = CurrArg
            ElseIf CmdLineFromDir Is Nothing Then
                CmdLineFromDir = CurrArg
            ElseIf CmdLineToDir Is Nothing Then
                CmdLineToDir = CurrArg
            End If
        Next

        ' --- Set combo box values from settings or command line ---
        If String.IsNullOrWhiteSpace(CmdLineApplication) Then
            ' --- Select the same values as last time ---
            ComboApplication.Text = My.Settings.LastAppName
            If Directory.Exists(My.Settings.LastFromDir) Then
                ComboFromDir.Text = My.Settings.LastFromDir
            End If
            If Directory.Exists(My.Settings.LastToDir) Then
                ComboToDir.Text = My.Settings.LastToDir
            End If
        Else
            ComboApplication.Text = CmdLineApplication
            If String.IsNullOrWhiteSpace(ComboApplication.Text) Then
                ' --- Create new application if it doesn't exist yet ---
                ComboApplication.Items.Add(CmdLineApplication)
                ComboApplication.Text = CmdLineApplication
            End If
            If Not String.IsNullOrWhiteSpace(CmdLineFromDir) AndAlso Directory.Exists(CmdLineFromDir) Then
                ComboFromDir.Text = CmdLineFromDir
                If Not String.IsNullOrWhiteSpace(CmdLineToDir) AndAlso Directory.Exists(CmdLineToDir) Then
                    ComboToDir.Text = CmdLineToDir
                End If
            End If
        End If

        ' --- Set checkboxes from saved settings ---
        IgnoreSpacesToolStripMenuItem.Checked = My.Settings.OptionIgnoreSpaces
        IgnoreVersionsToolStripMenuItem.Checked = My.Settings.OptionIgnoreVersions
        ToolStripMenuItemQuickCompareBinary.Checked = My.Settings.QuickCompareBinary
        ToolStripMenuItemIncludeTestProj.Checked = My.Settings.IncludeTestProjects
        ToolStripMenuItemIgnoreMissingDirectoryContents.Checked = My.Settings.IgnoreMissingDirectories

        Me.Show()
        Application.DoEvents()

        ' --- Immediately run if called from the command line ---
        If Not String.IsNullOrWhiteSpace(CmdLineToDir) AndAlso ButtonCompare.Enabled Then
            ButtonCompare.PerformClick()
        End If

    End Sub

    Private Sub FormMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        DoCancel = True
    End Sub

    Private Sub FillMyFileCompareSettings()

        With MyFileCompare

            ''.LocalSecKey = My.Settings.LocalSecKey
            ''.TestSecKey = My.Settings.TestSecKey
            ''.AcceptSecKey = My.Settings.AcceptSecKey
            ''.AcceptDNSecKey = My.Settings.AcceptSecKey ' Same as Accept
            ''.FinanceSecKey = My.Settings.FinanceSecKey
            ''.ProdSecKey = My.Settings.ProdSecKey
            ''.ReportSecKey = My.Settings.ProdSecKey ' Same as Prod

            ''.LocalPublicKeyToken = My.Settings.LocalPublicKeyToken
            ''.TestPublicKeyToken = My.Settings.TestPublicKeyToken
            ''.AcceptPublicKeyToken = My.Settings.AcceptPublicKeyToken
            ''.AcceptDNPublicKeyToken = My.Settings.AcceptPublicKeyToken
            ''.FinancePublicKeyToken = My.Settings.FinancePublicKeyToken
            ''.ProdPublicKeyToken = My.Settings.ProdPublicKeyToken
            ''.ReportPublicKeyToken = My.Settings.ProdPublicKeyToken ' Same as Prod

            ''.PCDrive = My.Settings.PCDrive
            ''.LocalDrive = My.Settings.LocalDrive
            ''.TestDrive = My.Settings.TestDrive
            ''.AcceptDrive = My.Settings.AcceptDrive
            ''.AcceptDNDrive = My.Settings.AcceptDNDrive
            ''.FinanceDrive = My.Settings.FinanceDrive
            ''.ProdDrive = My.Settings.ProdDrive
            ''.ProdDrive2 = My.Settings.ProdDrive2
            ''.ReportDrive = My.Settings.ReportDrive

            ''.PCPath = My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower
            ''.LocalPath = My.Settings.LocalPath.Replace("*", GetUserNameAdj)
            ''.TestPath = My.Settings.TestPath
            ''.AcceptPath = My.Settings.AcceptPath
            ''.AcceptDNPath = My.Settings.AcceptDNPath
            ''.FinancePath = My.Settings.FinancePath
            ''.ProdPath = My.Settings.ProdPath
            ''.ProdPath2 = My.Settings.ProdPath2
            ''.ReportPath = My.Settings.ReportPath

        End With

    End Sub

    Private Sub ButtonCompare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCompare.Click
        InitCompare()
    End Sub

    Private Sub InitCompare()
        If String.IsNullOrEmpty(ComboFromDir.Text) Then Exit Sub
        If String.IsNullOrEmpty(ComboToDir.Text) Then Exit Sub
        If Not Directory.Exists(ComboFromDir.Text) Then
            MessageBox.Show("Directory not found: " + ComboFromDir.Text, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Exit Sub
        End If
        If Not Directory.Exists(ComboToDir.Text) Then
            MessageBox.Show("Directory not found: " + ComboToDir.Text, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Exit Sub
        End If
        My.Settings.LastAppName = ComboApplication.Text
        My.Settings.LastFromDir = ComboFromDir.Text
        My.Settings.LastToDir = ComboToDir.Text
        AddAppPathExtra(ComboApplication.Text, ComboFromDir.Text)
        AddAppPathExtra(ComboApplication.Text, ComboToDir.Text)
        My.Settings.Save()
        StartCompare()
    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        DoCancel = True
    End Sub

    Private Sub MarkEnabled(ByVal Value As Boolean)
        ' --- Enable/disable controls based on value ---
        ComboApplication.Enabled = Value
        ComboFromDir.Enabled = Value
        ComboToDir.Enabled = Value
        ' --- Change bottom buttons based on tab page and list box contents ---
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            ButtonCopyFromTo.Enabled = ((ListBoxDiff.SelectedItems.Count > 0) And Value)
            ButtonDeleteFromOnly.Enabled = False
            ButtonShowDiffs.Text = "Show Differences"
            ButtonShowDiffs.Enabled = (ListBoxDiff.SelectedItems.Count = 1)
            ButtonSelectAll.Enabled = ((ListBoxDiff.Items.Count > 0) And Value)
            ButtonDeleteToOnly.Enabled = False
            ButtonCopyToFrom.Enabled = ((ListBoxDiff.SelectedItems.Count > 0) And Value)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            ButtonCopyFromTo.Enabled = ((ListBoxProjDiff.SelectedItems.Count > 0) And Value)
            ButtonDeleteFromOnly.Enabled = False
            ButtonShowDiffs.Text = "Show Differences"
            ButtonShowDiffs.Enabled = (ListBoxProjDiff.SelectedItems.Count = 1)
            ButtonSelectAll.Enabled = ((ListBoxProjDiff.Items.Count > 0) And Value)
            ButtonDeleteToOnly.Enabled = False
            ButtonCopyToFrom.Enabled = ((ListBoxProjDiff.SelectedItems.Count > 0) And Value)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageFrom" Then
            ButtonCopyFromTo.Enabled = ((ListBoxFrom.SelectedItems.Count > 0) And Value)
            ButtonDeleteFromOnly.Enabled = ((ListBoxFrom.SelectedItems.Count > 0) And Value)
            ButtonShowDiffs.Text = "View File"
            ButtonShowDiffs.Enabled = (ListBoxFrom.SelectedItems.Count = 1)
            ButtonSelectAll.Enabled = ((ListBoxFrom.Items.Count > 0) And Value)
            ButtonDeleteToOnly.Enabled = False
            ButtonCopyToFrom.Enabled = False
        ElseIf TabControlMain.SelectedTab.Name = "TabPageTo" Then
            ButtonCopyFromTo.Enabled = False
            ButtonDeleteFromOnly.Enabled = False
            ButtonShowDiffs.Text = "View File"
            ButtonShowDiffs.Enabled = (ListBoxTo.SelectedItems.Count = 1)
            ButtonSelectAll.Enabled = ((ListBoxTo.Items.Count > 0) And Value)
            ButtonDeleteToOnly.Enabled = ((ListBoxTo.SelectedItems.Count > 0) And Value)
            ButtonCopyToFrom.Enabled = ((ListBoxTo.SelectedItems.Count > 0) And Value)
        End If
        CurrEnabled = Value
        My.Application.DoEvents()
    End Sub

    Private Sub IncDirCount()
        DirCount += 1
        ShowStatus()
    End Sub

    Private Sub IncFileCount()
        FileCount += 1
        ShowStatus()
    End Sub

    Private Sub IncDiffCount()
        DiffCount += 1
        ShowStatus()
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ShowStatus()
        StatusLabelCounts.Text = "Directories Checked = " + DirCount.ToString +
                                 ", Files Checked = " + FileCount.ToString +
                                 ", Differences Found = " + DiffCount.ToString
        Application.DoEvents()
    End Sub

    Private Sub ListFilesFrom(ByVal ThisDir As DirectoryInfo)
        Dim CompFilename As String
        ' ------------------------
        Application.DoEvents()
        If DoCancel Then Exit Sub
        IncDirCount()
        Dim ThisDirParentName As String = ""
        If ThisDir.Parent IsNot Nothing Then ' Top-level directories have no parent!
            ThisDirParentName = ThisDir.Parent.Name
        End If
        If IgnoreDirectory(ThisDir.Name.ToLower, ThisDirParentName.ToLower) Then Exit Sub
        If Not IgnoreDirectoryFiles(ThisDir.Name.ToLower, ThisDirParentName.ToLower) Then
            If My.Settings.IgnoreMissingDirectories Then
                If Not Directory.Exists(ReplaceIgnoreCase(ThisDir.FullName, ComboFromDir.Text.ToLower, ComboToDir.Text)) Then
                    ' --- compare file doesn't exist ---
                    If Not IgnoreMissingDirectoryName(ThisDir.Name.ToLower) Then
                        If ThisDir.GetFiles.Count > 0 OrElse ThisDir.GetDirectories.Count > 0 Then
                            ListBoxFrom.Items.Add(ThisDir.FullName + "\...")
                            IncDiffCount()
                        End If
                    End If
                    Exit Sub
                End If
            End If
            For Each TempFile As FileInfo In ThisDir.GetFiles
                Application.DoEvents()
                If DoCancel Then Exit Sub
                If (TempFile.Attributes And FileAttributes.Hidden) <> 0 Then Continue For
                If (TempFile.Attributes And FileAttributes.System) <> 0 Then Continue For
                IncFileCount()
                Dim RelFilename As String = TempFile.FullName.Substring(Len(ComboFromDir.Text) + 1)
                Dim TempFilename As String = TempFile.FullName.ToLower
                If IgnoreFile(TempFilename) Then Continue For
                Try
                    CompFilename = ReplaceIgnoreCase(TempFilename, ComboFromDir.Text.ToLower, ComboToDir.Text)
                    Dim CompFile As New FileInfo(CompFilename)
                    If CompFile.Exists Then
                        ' --- Compare as binary first ---
                        If Not SpecialFile(TempFilename) AndAlso Not ProjectFile(TempFilename) Then
                            If BinaryCompareClass.BinaryFilesMatch(TempFilename, CompFilename) Then
                                Continue For
                            End If
                            If KnownBinaryFile(TempFilename) Then
                                ListBoxDiff.Items.Add(RelFilename)
                                IncDiffCount()
                                Continue For
                            End If
                        End If
                        ' --- Compare special project files or ignore spaces in files ---
                        If Not KnownBinaryFile(TempFilename) Then
                            With MyFileCompare
                                .Clear()
                                .ResetFlags()
                                .FirstDiffOnly = True
                                .TabsToSpaces = My.Settings.OptionIgnoreSpaces
                                .SquishSpaces = My.Settings.OptionIgnoreSpaces
                                .SquishLines = My.Settings.OptionIgnoreSpaces
                                .TrimBlanks = My.Settings.OptionIgnoreSpaces
                                .IgnoreVersionNumbers = My.Settings.OptionIgnoreVersions
                            End With
                            Try
                                MyFileCompare.DoCompare(TempFilename, CompFilename)
                                If MyFileCompare.DiffCount <> 0 Then
                                    If SpecialFile(TempFilename) OrElse ProjectFile(TempFilename) Then
                                        ListBoxProjDiff.Items.Add(RelFilename)
                                    Else
                                        ListBoxDiff.Items.Add(RelFilename)
                                    End If
                                    IncDiffCount()
                                End If
                                Continue For
                            Catch ex As Exception
                                ' --- Must be a binary file ---
                            End Try
                        End If
                        ' --- Compare file sizes ---
                        If TempFile.Length <> CompFile.Length Then
                            If SpecialFile(TempFilename) OrElse ProjectFile(TempFilename) Then
                                ListBoxProjDiff.Items.Add(RelFilename)
                            Else
                                ListBoxDiff.Items.Add(RelFilename)
                            End If
                            IncDiffCount()
                            Continue For
                        End If
                        ' --- Compare file modification datetime ---
                        If TempFile.LastWriteTimeUtc = CompFile.LastWriteTimeUtc Then
                            Continue For
                        End If
                        If Not My.Settings.QuickCompareBinary Then
                            ' --- Compare file contents byte by byte ---
                            If Not BinaryCompareClass.BinaryFilesMatch(TempFilename, CompFilename) Then
                                If SpecialFile(TempFilename) OrElse ProjectFile(TempFilename) Then
                                    ListBoxProjDiff.Items.Add(RelFilename)
                                Else
                                    ListBoxDiff.Items.Add(RelFilename)
                                End If
                                IncDiffCount()
                                Continue For
                            End If
                            ' --- Check for EOL errors ---
                            If FileEOLErrors(TempFilename) Then
                                If SpecialFile(TempFilename) OrElse ProjectFile(TempFilename) Then
                                    ListBoxProjDiff.Items.Add(RelFilename)
                                Else
                                    ListBoxDiff.Items.Add(RelFilename)
                                End If
                                IncDiffCount()
                                Continue For
                            End If
                        End If
                    Else
                        ' --- compare file doesn't exist ---
                        ListBoxFrom.Items.Add(RelFilename)
                        IncDiffCount()
                    End If
                Catch ex As Exception
                    ' --- Ignore if open, such as with MS Word ---
                    If ex.Message.ToLower.Contains("being used by another process") Then Continue For
                    ' --- error! ---
                    ListBoxFrom.Items.Add(RelFilename)
                    IncDiffCount()
                End Try
            Next
        End If
        For Each TempDir As DirectoryInfo In ThisDir.GetDirectories
            If (TempDir.Attributes And FileAttributes.Hidden) <> 0 Then
                Continue For
            End If
            Application.DoEvents()
            If DoCancel Then Exit Sub
            ListFilesFrom(TempDir)
        Next
    End Sub

    Private Sub ListFilesTo(ByVal ThisDir As DirectoryInfo)
        Dim CompFilename As String
        ' ------------------------
        Application.DoEvents()
        If DoCancel Then Exit Sub
        IncDirCount()
        Dim ThisDirParentName As String = ""
        If ThisDir.Parent IsNot Nothing Then ' Top-level directories have no parent!
            ThisDirParentName = ThisDir.Parent.Name
        End If
        If IgnoreDirectory(ThisDir.Name.ToLower, ThisDirParentName.ToLower) Then Exit Sub
        If Not IgnoreDirectoryFiles(ThisDir.Name.ToLower, ThisDirParentName.ToLower) Then
            If My.Settings.IgnoreMissingDirectories Then
                If Not Directory.Exists(ReplaceIgnoreCase(ThisDir.FullName, ComboToDir.Text.ToLower, ComboFromDir.Text)) Then
                    ' --- compare file doesn't exist ---
                    If Not IgnoreMissingDirectoryName(ThisDir.Name.ToLower) Then
                        If ThisDir.GetFiles.Count > 0 OrElse ThisDir.GetDirectories.Count > 0 Then
                            ListBoxTo.Items.Add(ThisDir.FullName + "\...")
                            IncDiffCount()
                        End If
                    End If
                    Exit Sub
                End If
            End If
            For Each TempFile As FileInfo In ThisDir.GetFiles
                Application.DoEvents()
                If DoCancel Then Exit Sub
                If (TempFile.Attributes And FileAttributes.Hidden) <> 0 Then Continue For
                If (TempFile.Attributes And FileAttributes.System) <> 0 Then Continue For
                IncFileCount()
                Dim RelFilename As String = TempFile.FullName.Substring(Len(ComboToDir.Text) + 1)
                Dim TempFilename As String = TempFile.FullName.ToLower
                If IgnoreFile(TempFilename) Then Continue For
                Try
                    CompFilename = ReplaceIgnoreCase(TempFilename, ComboToDir.Text.ToLower, ComboFromDir.Text)
                    Dim CompFile As New FileInfo(CompFilename)
                    If CompFile.Exists Then
                        ' --- do nothing, already checked ---
                    Else
                        ' --- compare file doesn't exist ---
                        ListBoxTo.Items.Add(RelFilename)
                        IncDiffCount()
                    End If
                Catch ex As Exception
                    ' --- Ignore if open, such as with MS Word ---
                    If ex.Message.ToLower.Contains("being used by another process") Then Continue For
                    ' --- error! ---
                    ListBoxTo.Items.Add(RelFilename)
                    IncDiffCount()
                End Try
            Next
        End If
        For Each TempDir As DirectoryInfo In ThisDir.GetDirectories
            If (TempDir.Attributes And FileAttributes.Hidden) <> 0 Then
                Continue For
            End If
            Application.DoEvents()
            If DoCancel Then Exit Sub
            ListFilesTo(TempDir)
        Next
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
    End Sub

    Private Sub ListBoxDiff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxDiff.MouseDown
        If e.Button = MouseButtons.Right Then
            If ListBoxDiff.SelectedItems.Count = 1 Then
                Dim TempFormDiffInfo As New FormDiffInfo
                Dim TempFilename As String = CType(ListBoxDiff.Items(ListBoxDiff.SelectedIndices(0)), System.String)
                Dim FromDate As DateTime = File.GetLastWriteTime(ComboFromDir.Text + "\" + TempFilename)
                Dim ToDate As DateTime = File.GetLastWriteTime(ComboToDir.Text + "\" + TempFilename)
                With TempFormDiffInfo
                    .Text = TempFilename.Substring(TempFilename.LastIndexOf("\"c) + 1)
                    If FromDate < ToDate Then
                        .LabelFromNewerOlder.Text = "Older"
                        .LabelToNewerOlder.Text = "Newer"
                    Else
                        .LabelFromNewerOlder.Text = "Newer"
                        .LabelToNewerOlder.Text = "Older"
                    End If
                    .LabelFromDate.Text = FromDate.ToString
                    .LabelToDate.Text = ToDate.ToString
                    .LabelFromSize.Text = New FileInfo(ComboFromDir.Text + "\" + TempFilename).Length.ToString("#,#")
                    .LabelToSize.Text = New FileInfo(ComboToDir.Text + "\" + TempFilename).Length.ToString("#,#")
                    .ShowDialog()
                End With
            End If
        End If
    End Sub

    Private Sub ListBoxDiff_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxDiff.MouseDoubleClick
        ShowDifferences()
    End Sub

    Private Sub ListBoxProjDiff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxProjDiff.MouseDown
        If e.Button = MouseButtons.Right Then
            If ListBoxProjDiff.SelectedItems.Count = 1 Then
                Dim TempFormDiffInfo As New FormDiffInfo
                Dim TempFilename As String = CType(ListBoxProjDiff.Items(ListBoxProjDiff.SelectedIndices(0)), System.String)
                Dim FromDate As DateTime = File.GetLastWriteTime(ComboFromDir.Text + "\" + TempFilename)
                Dim ToDate As DateTime = File.GetLastWriteTime(ComboToDir.Text + "\" + TempFilename)
                With TempFormDiffInfo
                    .Text = TempFilename.Substring(TempFilename.LastIndexOf("\"c) + 1)
                    If FromDate < ToDate Then
                        .LabelFromNewerOlder.Text = "Older"
                        .LabelToNewerOlder.Text = "Newer"
                    Else
                        .LabelFromNewerOlder.Text = "Newer"
                        .LabelToNewerOlder.Text = "Older"
                    End If
                    .LabelFromDate.Text = FromDate.ToString
                    .LabelToDate.Text = ToDate.ToString
                    .LabelFromSize.Text = ""
                    .LabelToSize.Text = ""
                    .ShowDialog()
                End With
            End If
        End If
    End Sub

    Private Sub ListBoxProjDiff_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxProjDiff.MouseDoubleClick
        ShowDifferences()
    End Sub

    Private Sub ListBoxFrom_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxFrom.MouseDoubleClick
        ViewFromOnlyFile()
    End Sub

    Private Sub ListBoxTo_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxTo.MouseDoubleClick
        ViewToOnlyFile()
    End Sub

    ''' <summary>
    ''' Both directory names must be sent as lowercase.
    ''' </summary>
    Private Function IgnoreDirectory(ByVal DirName As String, ByVal ParentName As String) As Boolean
        If DirName.StartsWith(".") Then Return True
        If DirName = "bin" Then
            If ParentName = "arena" Then Return False
            If ParentName = "arena2" Then Return False
            If ParentName = "utilities" Then Return False
            Return True
        End If
        If DirName = "obj" Then Return True
        If DirName = "install" Then Return True
        ''If DirName = "packages" Then Return True
        If DirName = "publish" Then Return True
        If DirName = "testresults" Then Return True
        If DirName = "buildprocesstemplates" Then Return True
        If Not My.Settings.IncludeTestProjects Then
            If DirName.StartsWith("test_") Then Return True
        End If
        If ParentName = "bin" Then
            If DirName = "debug" Then Return True
            If DirName = "release" Then Return True
            If DirName = "x86" Then Return True
            If DirName = "security" Then Return True
            Return False
        End If
        Return False
    End Function

    ''' <summary>
    ''' Both directory names must be sent as lowercase.
    ''' </summary>
    Private Function IgnoreDirectoryFiles(ByVal DirName As String, ByVal ParentName As String) As Boolean
        If DirName = "bin" Then
            If ParentName = "arena" Then Return True
            If ParentName = "arena2" Then Return True
            If ParentName = "utilities" Then Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Directory name must be sent as lowercase.
    ''' </summary>
    Private Function IgnoreMissingDirectoryName(ByVal DirName As String) As Boolean
        If DirName = "service references" Then Return True
        If DirName = "datasources" Then Return True
        Return False
    End Function

    Private Function IgnoreFile(ByVal FileName As String) As Boolean
        FileName = FileName.ToLower
        ' --- Ignore everything in Applications except for Settings ---
        If FileName.Contains("\applications\") Then
            If FileName.EndsWith(".settings") Then Return False
            Return True
        End If
        ' --- Ignore known configuration files ---
        If FileName.StartsWith(".") OrElse FileName.Contains("\.") Then
            If FileName.EndsWith(".gitignore") Then Return False
            Return True
        End If
        ' --- Ignore known and common binary, temporary, or work files ---
        If FileName.EndsWith(".application") Then Return True
        If FileName.EndsWith(".bak") Then Return True
        If FileName.EndsWith(".cache") Then Return True
        If FileName.EndsWith(".com") Then Return True
        If FileName.EndsWith(".db") Then Return True
        If FileName.EndsWith(".deploy") Then Return True
        If FileName.EndsWith(".dll") Then
            If FileName.Contains("\packages\") Then Return False
            Return True
        End If
        If FileName.EndsWith(".dbmdl") Then Return True
        If FileName.EndsWith(".exe") Then Return True
        If FileName.EndsWith(".jfm") Then Return True
        If FileName.EndsWith(".lnk") Then Return True
        If FileName.EndsWith(".log") Then Return True
        If FileName.EndsWith(".ocx") Then Return True
        If FileName.EndsWith(".pdb") Then Return True
        If FileName.EndsWith(".sav") Then Return True
        If FileName.EndsWith(".scc") Then Return True
        If FileName.EndsWith(".suo") Then Return True
        If FileName.EndsWith(".tmp") Then Return True
        If FileName.EndsWith(".udl") Then Return True
        If FileName.EndsWith(".user") Then Return True
        If FileName.EndsWith(".vbw") Then Return True
        If FileName.EndsWith(".vspscc") Then Return True
        If FileName.EndsWith(".vssscc") Then Return True
        ' --- Ignore batch files which have paths in them ---
        If FileName.EndsWith("buildall.bat") Then Return True
        If FileName.EndsWith("buildtest.bat") Then Return True
        If FileName.EndsWith("publishall.bat") Then Return True
        If FileName.EndsWith("vaultbackup.bat") Then Return True
        If FileName.EndsWith("vaultcompare.bat") Then Return True
        If FileName.EndsWith("vvbackup.bat") Then Return True
        If FileName.EndsWith("vvbackupforce.bat") Then Return True
        If FileName.EndsWith("vvcompare.bat") Then Return True
        If FileName.EndsWith("clearuserprograms.bat") Then Return True
        If FileName.EndsWith("clearusersettings.bat") Then Return True
        ' --- Use ExcludeFileList ---
        If Not String.IsNullOrEmpty(My.Settings.ExcludeFileList) Then
            Dim TempList As String() = My.Settings.ExcludeFileList.ToLower.Split(";"c)
            For Each TempItem As String In TempList
                TempItem = TempItem.Trim
                If Not String.IsNullOrEmpty(TempItem) Then
                    If FileName.EndsWith(TempItem) Then Return True
                End If
            Next
        End If
        Return False
    End Function

    Private Function KnownBinaryFile(ByVal Filename As String) As Boolean
        Filename = Filename.ToLower
        ' --- Check for known non-special files that might be flagged below due to names ---
        If Filename.Contains(".azw") Then Return True
        If Filename.EndsWith(".doc") Then Return True
        If Filename.EndsWith(".docx") Then Return True
        If Filename.EndsWith(".epub") Then Return True
        If Filename.EndsWith(".flac") Then Return True
        If Filename.EndsWith(".frx") Then Return True
        If Filename.EndsWith(".gif") Then Return True
        If Filename.EndsWith(".ico") Then Return True
        If Filename.EndsWith(".jpg") Then Return True
        If Filename.EndsWith(".mobi") Then Return True
        If Filename.EndsWith(".mp3") Then Return True
        If Filename.EndsWith(".pdf") Then Return True
        If Filename.EndsWith(".pfx") Then Return True
        If Filename.EndsWith(".wav") Then Return True
        If Filename.EndsWith(".wma") Then Return True
        If Filename.EndsWith(".xls") Then Return True
        If Filename.EndsWith(".xlsx") Then Return True
        If Filename.EndsWith(".xps") Then Return True
        If Filename.EndsWith(".zip") Then Return True
        If Filename.EndsWith(".zipx") Then Return True
        Return False
    End Function

    Private Function SpecialFile(ByVal FileName As String) As Boolean
        ''FileName = FileName.ToLower
        '' ' --- Don't try to alter the source for this program!!! ---
        ''If ContainsIgnoreCase(FileName, My.Application.Info.AssemblyName) Then
        ''    If FileName.EndsWith(".vb") Then Return False
        ''    If FileName.EndsWith("\settings.designer.vb") Then Return False
        ''    If FileName.EndsWith("\settings.settings") Then Return False
        ''    If FileName.EndsWith("\app.config") Then Return False
        ''End If
        '' ' --- Other files not to mess with ---
        ''If ContainsIgnoreCase(FileName, "IDRISMakeLib") OrElse ContainsIgnoreCase(FileName, "IDRIS_IDE") Then
        ''    If FileName.EndsWith("\settings.designer.vb") Then Return False
        ''    If FileName.EndsWith("\settings.settings") Then Return False
        ''    If FileName.EndsWith("\app.config") Then Return False
        ''End If
        '' ' --- All filenames which contain environment info ---
        ''If FileName.EndsWith(".vbproj") Then Return True
        ''If FileName.EndsWith("\app.config") Then Return True
        ''If FileName.EndsWith("\settings.designer.vb") Then Return True
        ''If FileName.EndsWith("\settings.settings") Then Return True
        ''If FileName.EndsWith("\assemblyinfo.vb") Then Return True
        Return False
    End Function

    Private Function ProjectFile(ByVal Filename As String) As Boolean
        Filename = Filename.ToLower
        ' --- Code Signing files ---
        If Filename.EndsWith(".pfx") Then Return True
        ' --- VB.NET Project Files ---
        If Filename.EndsWith(".sln") Then Return True
        If Filename.EndsWith(".vbproj") Then Return True
        If Filename.EndsWith("\app.config") Then Return True
        If Filename.Contains("\my project\") Then Return True
        ' --- C# project files ---
        If Filename.EndsWith(".csproj") Then Return True
        If Filename.Contains("\properties\") Then Return True
        ' --- SQL project files ---
        If Filename.EndsWith(".sqlproj") Then Return True
        ' --- VB6 project files ---
        If Filename.EndsWith(".vbp") Then Return True
        Return False
    End Function

    Private Sub ButtonShowDiffs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonShowDiffs.Click
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            ShowDifferences()
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            ShowDifferences()
        ElseIf TabControlMain.SelectedTab.Name = "TabPageFrom" Then
            ViewFromOnlyFile()
        ElseIf TabControlMain.SelectedTab.Name = "TabPageTo" Then
            ViewToOnlyFile()
        End If
    End Sub

    Private Sub ShowDifferences()
        Dim MyFormDiff As New FormDiff
        Dim TempFilename As String = ""
        Dim TempListBox As ListBox = Nothing
        ' ----------------------------------
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            TempListBox = ListBoxDiff
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            TempListBox = ListBoxProjDiff
        End If
        If TempListBox.SelectedItems.Count <> 1 Then Exit Sub
        Try
            TempFilename = CType(TempListBox.SelectedItem, System.String)
            If String.IsNullOrEmpty(TempFilename) Then Exit Sub
            If KnownBinaryFile(TempFilename) Then Exit Sub
            If My.Settings.UseExternalCompare AndAlso Not String.IsNullOrWhiteSpace(My.Settings.ExternalCompareApp) Then
                Try
                    Dim CmdLine As String = My.Settings.ExternalCompareApp + " """ + ComboFromDir.Text + "\" + TempFilename + """ """ + ComboToDir.Text + "\" + TempFilename + """"
                    Shell(CmdLine, AppWinStyle.MaximizedFocus, False)
                Catch ex As Exception
                    MessageBox.Show("Error running external compare application: " + My.Settings.ExternalCompareApp +
                                    vbCrLf + vbCrLf +
                                    ex.Message, "External Compare Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try
                Exit Sub
            End If
            Cursor.Current = Cursors.WaitCursor
            MyFormDiff.TextBoxDiff.Clear()
            MyFormDiff.Text = TempFilename
            With MyFileCompare
                .Clear()
                .ResetFlags()
                .FirstDiffOnly = False
                .TabsToSpaces = My.Settings.OptionIgnoreSpaces
                .SquishSpaces = My.Settings.OptionIgnoreSpaces
                .SquishLines = My.Settings.OptionIgnoreSpaces
                .TrimBlanks = My.Settings.OptionIgnoreSpaces
                .IgnoreVersionNumbers = My.Settings.OptionIgnoreVersions
                .DoCompare(ComboFromDir.Text + "\" + TempFilename, ComboToDir.Text + "\" + TempFilename)
                MyFormDiff.TextBoxDiff.Text = .Results
            End With
            MyFormDiff.TextBoxDiff.SelectionStart = 0
            Cursor.Current = Cursors.Default
            MyFormDiff.ShowDialog()
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MessageBox.Show("Error encountered:" + vbCrLf + vbCrLf + ex.Message)
        End Try
    End Sub

    Private Sub ViewFromOnlyFile()
        If ListBoxFrom.SelectedItems.Count <> 1 Then Exit Sub
        Dim MyFormView As New FormDiff
        Try
            Dim TempFilename As String = CType(ListBoxFrom.SelectedItem, System.String)
            If String.IsNullOrEmpty(TempFilename) Then Exit Sub
            If TempFilename.EndsWith("\...") Then Exit Sub
            If KnownBinaryFile(TempFilename) Then Exit Sub
            MyFormView.TextBoxDiff.Clear()
            MyFormView.Text = TempFilename
            Dim CurrEncoding As Encoding = GetFileEncoding(ComboFromDir.Text + "\" + TempFilename)
            MyFormView.TextBoxDiff.Text = File.ReadAllText(ComboFromDir.Text + "\" + TempFilename, CurrEncoding)
            MyFormView.TextBoxDiff.SelectionStart = 0
            MyFormView.ShowDialog()
        Catch ex As Exception
            ' --- do nothing ---
        End Try
    End Sub

    Private Sub ViewToOnlyFile()
        If ListBoxTo.SelectedItems.Count <> 1 Then Exit Sub
        Dim MyFormView As New FormDiff
        Try
            Dim TempFilename As String = CType(ListBoxTo.SelectedItem, System.String)
            If String.IsNullOrEmpty(TempFilename) Then Exit Sub
            If TempFilename.EndsWith("\...") Then Exit Sub
            If KnownBinaryFile(TempFilename) Then Exit Sub
            MyFormView.TextBoxDiff.Clear()
            MyFormView.Text = TempFilename
            Dim CurrEncoding As Encoding = GetFileEncoding(ComboToDir.Text + "\" + TempFilename)
            MyFormView.TextBoxDiff.Text = File.ReadAllText(ComboToDir.Text + "\" + TempFilename, CurrEncoding)
            MyFormView.TextBoxDiff.SelectionStart = 0
            MyFormView.ShowDialog()
        Catch ex As Exception
            ' --- do nothing ---
        End Try
    End Sub

    Private Sub TabControlMain_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlMain.SelectedIndexChanged
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ListBoxDiff_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxDiff.SelectedValueChanged
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ListBoxProjDiff_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxProjDiff.SelectedValueChanged
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ListBoxFrom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxFrom.SelectedIndexChanged
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ListBoxTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxTo.SelectedIndexChanged
        MarkEnabled(CurrEnabled)
    End Sub

    Private Sub ButtonCopyFromTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCopyFromTo.Click
        Dim TempListBox As ListBox
        Dim LastOverwriteAnswer As OverwriteResult = OverwriteResult.Unknown
        Dim LastReadonlyAnswer As OverwriteResult = OverwriteResult.Unknown
        Dim CopyAnswer As OverwriteResult
        Dim CopiedCount As Integer = 0
        Dim CopiedIndexes As New List(Of Integer)
        ' ---------------------------------------
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            TempListBox = ListBoxDiff
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            TempListBox = ListBoxProjDiff
        ElseIf TabControlMain.SelectedTab.Name = "TabPageFrom" Then
            TempListBox = ListBoxFrom
        Else
            Exit Sub
        End If
        If TempListBox.SelectedItems.Count < 1 Then Exit Sub
        For Each CurrIndex As Integer In TempListBox.SelectedIndices
            Dim TempFilename As String = CType(TempListBox.Items(CurrIndex), System.String)
            If String.IsNullOrEmpty(TempFilename) Then Continue For
            If TempFilename.EndsWith("\...") Then Continue For
            Do
                CopyAnswer = DoCopyFile(ComboFromDir.Text + "\", ComboToDir.Text + "\", TempFilename, LastOverwriteAnswer, LastReadonlyAnswer)
            Loop Until CopyAnswer <> OverwriteResult.Retry
            If CopyAnswer = OverwriteResult.Abort Then Exit For
            If CopyAnswer = OverwriteResult.IgnoreAll Then
                LastReadonlyAnswer = CopyAnswer
            End If
            If CopyAnswer = OverwriteResult.YesToAll OrElse CopyAnswer = OverwriteResult.NoToAll Then
                LastOverwriteAnswer = CopyAnswer
            End If
            If CopyAnswer = OverwriteResult.Ignore OrElse CopyAnswer = OverwriteResult.IgnoreAll Then
                Continue For
            End If
            If CopyAnswer = OverwriteResult.Yes OrElse CopyAnswer = OverwriteResult.YesToAll Then
                CopiedCount += 1
                CopiedIndexes.Add(CurrIndex)
                StatusLabelCounts.Text = "Files Copied: " + CopiedCount.ToString
                Application.DoEvents()
            End If
        Next
        If CopiedCount > 0 Then
            CopiedIndexes.Sort()
            CopiedIndexes.Reverse()
            For Each CurrIndex As Integer In CopiedIndexes
                TempListBox.Items.RemoveAt(CurrIndex)
            Next
        End If
        StatusLabelCounts.Text = "Files Copied: " + CopiedCount.ToString + " - Copy Completed"
    End Sub

    Private Sub ButtonCopyToFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCopyToFrom.Click
        Dim TempListBox As ListBox
        Dim LastOverwriteAnswer As OverwriteResult = OverwriteResult.Unknown
        Dim LastReadonlyAnswer As OverwriteResult = OverwriteResult.Unknown
        Dim CopyAnswer As OverwriteResult
        Dim CopiedCount As Integer = 0
        Dim CopiedIndexes As New List(Of Integer)
        ' ---------------------------------------
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            TempListBox = ListBoxDiff
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            TempListBox = ListBoxProjDiff
        ElseIf TabControlMain.SelectedTab.Name = "TabPageTo" Then
            TempListBox = ListBoxTo
        Else
            Exit Sub
        End If
        If TempListBox.SelectedItems.Count < 1 Then Exit Sub
        For Each CurrIndex As Integer In TempListBox.SelectedIndices
            Dim TempFilename As String = CType(TempListBox.Items(CurrIndex), System.String)
            If String.IsNullOrEmpty(TempFilename) Then Continue For
            If TempFilename.EndsWith("\...") Then Continue For
            Do
                CopyAnswer = DoCopyFile(ComboToDir.Text + "\", ComboFromDir.Text + "\", TempFilename, LastOverwriteAnswer, LastReadonlyAnswer)
            Loop Until CopyAnswer <> OverwriteResult.Retry
            If CopyAnswer = OverwriteResult.Abort Then Exit For
            If CopyAnswer = OverwriteResult.IgnoreAll Then
                LastReadonlyAnswer = CopyAnswer
            End If
            If CopyAnswer = OverwriteResult.YesToAll OrElse CopyAnswer = OverwriteResult.NoToAll Then
                LastOverwriteAnswer = CopyAnswer
            End If
            If CopyAnswer = OverwriteResult.Ignore OrElse CopyAnswer = OverwriteResult.IgnoreAll Then
                Continue For
            End If
            If CopyAnswer = OverwriteResult.Yes OrElse CopyAnswer = OverwriteResult.YesToAll Then
                CopiedCount += 1
                CopiedIndexes.Add(CurrIndex)
                StatusLabelCounts.Text = "Files Copied: " + CopiedCount.ToString
                Application.DoEvents()
            End If
        Next
        If CopiedCount > 0 Then
            CopiedIndexes.Sort()
            CopiedIndexes.Reverse()
            For Each CurrIndex As Integer In CopiedIndexes
                TempListBox.Items.RemoveAt(CurrIndex)
            Next
        End If
        StatusLabelCounts.Text = "Files Copied: " + CopiedCount.ToString + " - Copy Completed"
    End Sub

    Private Sub ButtonDeleteFromOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDeleteFromOnly.Click
        Dim TempListBox As ListBox
        Dim Answer As DialogResult
        Dim SingleMultiWording As String
        Dim DeletedCount As Integer = 0
        Dim DeletedIndexes As New List(Of Integer)
        ' ----------------------------------------
        If TabControlMain.SelectedTab.Name <> "TabPageFrom" Then Exit Sub
        TempListBox = ListBoxFrom
        If TempListBox.SelectedItems.Count < 1 Then Exit Sub
        If TempListBox.SelectedItems.Count = 1 Then
            SingleMultiWording = "this file"
        Else
            SingleMultiWording = "these files"
        End If
        Answer = MessageBox.Show("Do you really want to delete " + SingleMultiWording + "?", "Delete Files", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Answer = DialogResult.No Then Exit Sub
        For Each CurrIndex As Integer In TempListBox.SelectedIndices
            Dim TempFilename As String = CType(TempListBox.Items(CurrIndex), System.String)
            If String.IsNullOrEmpty(TempFilename) Then Continue For
            If TempFilename.EndsWith("\...") Then Continue For
            Do While (File.GetAttributes(ComboFromDir.Text + "\" + TempFilename) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly
                Answer = MessageBox.Show("File is Read-Only: " + ComboFromDir.Text + "\" + TempFilename, Me.Text, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                If Answer = DialogResult.Abort Then
                    Exit Sub
                End If
                If Answer = DialogResult.Ignore Then
                    Continue For
                End If
            Loop
            Try
                File.Delete(ComboFromDir.Text + "\" + TempFilename)
                DeletedCount += 1
                DeletedIndexes.Add(CurrIndex)
                Dim TempPath As String = ComboFromDir.Text + "\" + TempFilename
                Do
                    TempPath = TempPath.Substring(0, TempPath.LastIndexOf("\"c)) ' Get just the path
                    Try
                        If Directory.GetFiles(TempPath).Count = 0 AndAlso Directory.GetDirectories(TempPath).Count = 0 Then
                            Directory.Delete(TempPath)
                        Else
                            Exit Do
                        End If
                    Catch ex As Exception
                        Exit Do
                    End Try
                Loop While Not String.IsNullOrEmpty(TempPath)
            Catch ex As Exception
                Answer = MessageBox.Show("Error deleting file: " + ComboFromDir.Text + "\" + TempFilename, Me.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                If Answer = DialogResult.Cancel Then Exit Sub
            End Try
        Next
        If DeletedCount > 0 Then
            DeletedIndexes.Sort()
            DeletedIndexes.Reverse()
            For Each CurrIndex As Integer In DeletedIndexes
                TempListBox.Items.RemoveAt(CurrIndex)
            Next
        End If
        StatusLabelCounts.Text = "Files Deleted: " + DeletedCount.ToString + " - Delete Completed"
    End Sub

    Private Sub ButtonDeleteToOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDeleteToOnly.Click
        Dim TempListBox As ListBox
        Dim Answer As DialogResult
        Dim SingleMultiWording As String
        Dim DeletedCount As Integer = 0
        Dim DeletedIndexes As New List(Of Integer)
        ' ----------------------------------------
        If TabControlMain.SelectedTab.Name <> "TabPageTo" Then Exit Sub
        TempListBox = ListBoxTo
        If TempListBox.SelectedItems.Count < 1 Then Exit Sub
        If TempListBox.SelectedItems.Count = 1 Then
            SingleMultiWording = "this file"
        Else
            SingleMultiWording = "these files"
        End If
        Answer = MessageBox.Show("Do you really want to delete " + SingleMultiWording + "?", "Delete Files", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Answer = DialogResult.No Then Exit Sub
        For Each CurrIndex As Integer In TempListBox.SelectedIndices
            Dim TempFilename As String = CType(TempListBox.Items(CurrIndex), System.String)
            If String.IsNullOrEmpty(TempFilename) Then Continue For
            If TempFilename.EndsWith("\...") Then Continue For
            Do While (File.GetAttributes(ComboToDir.Text + "\" + TempFilename) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly
                Answer = MessageBox.Show("File is Read-Only: " + ComboToDir.Text + "\" + TempFilename, Me.Text, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                If Answer = DialogResult.Abort Then
                    Exit Sub
                End If
                If Answer = DialogResult.Ignore Then
                    Continue For
                End If
            Loop
            Try
                File.Delete(ComboToDir.Text + "\" + TempFilename)
                DeletedCount += 1
                DeletedIndexes.Add(CurrIndex)
                Dim TempPath As String = ComboToDir.Text + "\" + TempFilename
                Do
                    TempPath = TempPath.Substring(0, TempPath.LastIndexOf("\"c)) ' Get just the path
                    Try
                        If Directory.GetFiles(TempPath).Count = 0 AndAlso Directory.GetDirectories(TempPath).Count = 0 Then
                            Directory.Delete(TempPath)
                        Else
                            Exit Do
                        End If
                    Catch ex As Exception
                        Exit Do
                    End Try
                Loop While Not String.IsNullOrEmpty(TempPath)
            Catch ex As Exception
                Answer = MessageBox.Show("Error deleting file: " + ComboToDir.Text + "\" + TempFilename, Me.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                If Answer = DialogResult.Cancel Then Exit Sub
            End Try
        Next
        If DeletedCount > 0 Then
            DeletedIndexes.Sort()
            DeletedIndexes.Reverse()
            For Each CurrIndex As Integer In DeletedIndexes
                TempListBox.Items.RemoveAt(CurrIndex)
            Next
        End If
        StatusLabelCounts.Text = "Files Deleted: " + DeletedCount.ToString + " - Delete Completed"
    End Sub

    Private Function DoCopyFile(ByVal FromDir As String,
                                ByVal ToDir As String,
                                ByVal Filename As String,
                                ByVal LastOverwriteAnswer As OverwriteResult,
                                ByVal LastReadonlyAnswer As OverwriteResult) As OverwriteResult
        Dim ReadOnlyAnswer As OverwriteResult
        Dim OverwriteAnswer As OverwriteResult = OverwriteResult.Yes
        ' ----------------------------------------------------------
        ' --- Check if the target directory exists ---
        Dim DirPath As String = ToDir + Filename ' Filename might contain some path info
        DirPath = DirPath.Substring(0, DirPath.LastIndexOf("\"))
        Try
            If Not Directory.Exists(DirPath) Then
                Directory.CreateDirectory(DirPath)
            End If
        Catch ex As Exception
            MessageBox.Show("Error creating directory: " + DirPath, "Error Creating Directory", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return OverwriteResult.Abort
        End Try
        If File.Exists(ToDir + Filename) Then
            ' --- Check if file is read-only ---
            Do While (File.GetAttributes(ToDir + Filename) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly
                If LastReadonlyAnswer = OverwriteResult.IgnoreAll Then
                    Return OverwriteResult.IgnoreAll
                End If
                Dim TempFormTargetReadonly As New FormTargetReadonly
                TempFormTargetReadonly.LabelFilename.Text = ToDir + Filename
                TempFormTargetReadonly.ShowDialog()
                ReadOnlyAnswer = TempFormTargetReadonly.Result
                If ReadOnlyAnswer <> OverwriteResult.Retry Then
                    Return ReadOnlyAnswer
                End If
            Loop
            ' --- Check if the target file is newer ---
            Dim FromDate As DateTime = File.GetLastWriteTime(FromDir + Filename)
            Dim ToDate As DateTime = File.GetLastWriteTime(ToDir + Filename)
            If FromDate < ToDate Then
                If LastOverwriteAnswer = OverwriteResult.NoToAll Then
                    Return OverwriteResult.NoToAll
                End If
                If LastOverwriteAnswer <> OverwriteResult.YesToAll Then
                    Dim TempFormTargetNewer As New FormTargetNewer
                    TempFormTargetNewer.LabelFilename.Text = ToDir + Filename
                    TempFormTargetNewer.ShowDialog()
                    OverwriteAnswer = TempFormTargetNewer.Result
                    If OverwriteAnswer <> OverwriteResult.Yes AndAlso OverwriteAnswer <> OverwriteResult.YesToAll Then
                        Return OverwriteAnswer
                    End If
                End If
            End If
        End If
        ' --- Copy the file ---
        Try
            File.Copy(FromDir + Filename, ToDir + Filename, True)
        Catch ex As Exception
            MessageBox.Show("Error copying file: " + Filename, "Error copying file", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return OverwriteResult.Abort
        End Try
        Try
            If (File.GetAttributes(ToDir + Filename) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                File.SetAttributes(ToDir + Filename, File.GetAttributes(ToDir + Filename) And Not FileAttributes.ReadOnly)
            End If
        Catch ex As Exception
            MessageBox.Show("Error marking file read-write: " + Filename, "Error marking file", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return OverwriteResult.Abort
        End Try
        Return OverwriteAnswer
    End Function

    Private Sub StartCompare()
        ' --- Check for invalid selections ---
        If ComboFromDir.Text = "" OrElse ComboToDir.Text = "" Then Exit Sub
        If ComboFromDir.Text = ComboToDir.Text Then Exit Sub
        ' --- Compare files ---
        Try
            DoCancel = True
            ListBoxDiff.Items.Clear()
            ListBoxProjDiff.Items.Clear()
            ListBoxFrom.Items.Clear()
            ListBoxTo.Items.Clear()
            MarkEnabled(False)
            ButtonCompare.Visible = False
            ButtonCancel.Visible = True
            ToolStripMenuItemCompare.Enabled = False
            ButtonCancel.Focus()
            Application.DoEvents()
            DirCount = 0
            FileCount = 0
            DiffCount = 0
            ShowStatus()
            Dim FromDir As New DirectoryInfo(ComboFromDir.Text)
            Dim ToDir As New DirectoryInfo(ComboToDir.Text)
            DoCancel = False
            ListFilesFrom(FromDir)
            ListFilesTo(ToDir)
            If DoCancel Then
                StatusLabelCounts.Text += " - Canceled"
            Else
                StatusLabelCounts.Text += " - Done"
            End If
        Catch ex As Exception
            DoCancel = True
            MessageBox.Show(ex.Message, "Error comparing", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DoCancel = True
            MarkEnabled(True)
            ButtonCancel.Visible = False
            ButtonCompare.Visible = True
            ToolStripMenuItemCompare.Enabled = True
        End Try
    End Sub

    ''Private Function FixFile(ByVal FromDir As String, ByVal Pathname As String, ByVal Filename As String) As Boolean
    ''    Dim Lines() As String
    ''    Dim CurrLine As String
    ''    Dim Changed As Boolean = False
    ''    Dim LineChanged As Boolean = False
    ''    Dim RemoveNothings As Boolean = False
    ''    Dim PlatformTarget As Boolean = False
    ''    Dim SourceEnv As Char
    ''    Dim TargetEnv As Char
    ''    Dim FromDirU As String = FromDir.ToUpper
    ''    Dim PathnameU As String = Pathname.ToUpper
    ''    Dim FilenameU As String = Filename.ToUpper
    ''    Dim FindItem As String = ""
    ''    Dim ReplaceItem As String = ""
    ''    Dim CurrEncoding As Encoding
    ''    ' ----------------------------------------
    ''    ' --- Note: All settings are in uppercase! ---
    ''    ''If FromDirU.Contains("_FINANCE") Then ' Must check first
    ''    ''    SourceEnv = "F"c ' Finance
    ''    ''ElseIf FromDirU.Contains("\DROPBOX\") OrElse FromDirU.Contains("\SKYDRIVE\") OrElse FromDirU.Contains("\ONEDRIVE\") Then
    ''    ''    SourceEnv = "D"c ' dropbox/skydrive/onedrive
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.LocalDrive) OrElse FromDirU.StartsWith(My.Settings.LocalPath.Replace("*", GetUserNameAdj.ToUpper)) Then
    ''    ''    SourceEnv = "L"c ' local
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.TestDrive) OrElse FromDirU.StartsWith(My.Settings.TestPath) Then
    ''    ''    SourceEnv = "T"c ' test
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.AcceptDrive) OrElse FromDirU.StartsWith(My.Settings.AcceptPath) Then
    ''    ''    SourceEnv = "A"c ' acceptance
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.AcceptDNDrive) OrElse FromDirU.StartsWith(My.Settings.AcceptDNPath) Then
    ''    ''    SourceEnv = "X"c ' DN acceptance
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.ProdDrive) OrElse FromDirU.StartsWith(My.Settings.ProdPath) Then
    ''    ''    SourceEnv = "P"c ' production
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.ProdDrive2) OrElse FromDirU.StartsWith(My.Settings.ProdPath2) Then
    ''    ''    SourceEnv = "V"c ' production sourcesafe
    ''    ''ElseIf FromDirU.StartsWith(My.Settings.ReportDrive) OrElse FromDirU.StartsWith(My.Settings.ReportPath) Then
    ''    ''    SourceEnv = "R"c ' Report
    ''    ''Else
    ''    ''    SourceEnv = "C"c ' PC
    ''    ''End If
    ''    ''If PathnameU.Contains("_FINANCE") Then ' Must check first
    ''    ''    TargetEnv = "F"c ' Finance
    ''    ''ElseIf PathnameU.Contains("\DROPBOX\") OrElse PathnameU.Contains("\SKYDRIVE\") OrElse PathnameU.Contains("\ONEDRIVE\") Then
    ''    ''    TargetEnv = "D"c ' dropbox/skydrive/onedrive
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.LocalDrive) OrElse PathnameU.StartsWith(My.Settings.LocalPath.Replace("*", GetUserNameAdj.ToUpper)) Then
    ''    ''    TargetEnv = "L"c ' local
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.TestDrive) OrElse PathnameU.StartsWith(My.Settings.TestPath) Then
    ''    ''    TargetEnv = "T"c ' test
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.AcceptDrive) OrElse PathnameU.StartsWith(My.Settings.AcceptPath) Then
    ''    ''    TargetEnv = "A"c ' acceptance
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.AcceptDNDrive) OrElse PathnameU.StartsWith(My.Settings.AcceptDNPath) Then
    ''    ''    TargetEnv = "X"c ' DN acceptance
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.ProdDrive) OrElse PathnameU.StartsWith(My.Settings.ProdPath) Then
    ''    ''    TargetEnv = "P"c ' production
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.ProdDrive2) OrElse PathnameU.StartsWith(My.Settings.ProdPath2) Then
    ''    ''    TargetEnv = "V"c ' production sourcesafe
    ''    ''ElseIf PathnameU.StartsWith(My.Settings.ReportDrive) OrElse PathnameU.StartsWith(My.Settings.ReportPath) Then
    ''    ''    TargetEnv = "R"c ' Report
    ''    ''Else
    ''    ''    TargetEnv = "C"c ' PC
    ''    ''End If
    ''    ' --- Read lines from file ---
    ''    Try
    ''        CurrEncoding = GetFileEncoding(Pathname + Filename)
    ''        Lines = File.ReadAllLines(Pathname + Filename, CurrEncoding)
    ''    Catch ex As Exception
    ''        Return False
    ''    End Try
    ''    For CurrLineNum As Integer = 0 To Lines.GetUpperBound(0)
    ''        CurrLine = Lines(CurrLineNum)
    ''        If CurrLine Is Nothing Then Continue For
    ''        LineChanged = False
    ''        ' --- Check for lines that should not be changed ---
    ''        If MyFileCompare.ExcludeLine(CurrLine) Then Continue For
    ''        ' --- Check for security key hashcodes ---
    ''        If TargetEnv = "C"c Then ReplaceItem = My.Settings.LocalSecKey.ToUpper
    ''        If TargetEnv = "D"c Then ReplaceItem = My.Settings.LocalSecKey.ToUpper
    ''        If TargetEnv = "L"c Then ReplaceItem = My.Settings.LocalSecKey.ToUpper
    ''        If TargetEnv = "T"c Then ReplaceItem = My.Settings.TestSecKey.ToUpper
    ''        If TargetEnv = "A"c Then ReplaceItem = My.Settings.AcceptSecKey.ToUpper
    ''        If TargetEnv = "X"c Then ReplaceItem = My.Settings.AcceptSecKey.ToUpper ' Same as Accept
    ''        If TargetEnv = "F"c Then ReplaceItem = My.Settings.FinanceSecKey.ToUpper
    ''        If TargetEnv = "P"c Then ReplaceItem = My.Settings.ProdSecKey.ToUpper
    ''        If TargetEnv = "V"c Then ReplaceItem = My.Settings.ProdSecKey.ToUpper
    ''        If TargetEnv = "R"c Then ReplaceItem = My.Settings.ProdSecKey.ToUpper ' Same as Prod
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.LocalSecKey) Then FindItem = My.Settings.LocalSecKey
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.TestSecKey) Then FindItem = My.Settings.TestSecKey
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.AcceptSecKey) Then FindItem = My.Settings.AcceptSecKey
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.FinanceSecKey) Then FindItem = My.Settings.FinanceSecKey
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.ProdSecKey) Then FindItem = My.Settings.ProdSecKey
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for public key tokens ---
    ''        If TargetEnv = "C"c Then ReplaceItem = My.Settings.LocalPublicKeyToken.ToUpper
    ''        If TargetEnv = "D"c Then ReplaceItem = My.Settings.LocalPublicKeyToken.ToUpper
    ''        If TargetEnv = "L"c Then ReplaceItem = My.Settings.LocalPublicKeyToken.ToUpper
    ''        If TargetEnv = "T"c Then ReplaceItem = My.Settings.TestPublicKeyToken.ToUpper
    ''        If TargetEnv = "A"c Then ReplaceItem = My.Settings.AcceptPublicKeyToken.ToUpper
    ''        If TargetEnv = "X"c Then ReplaceItem = My.Settings.AcceptPublicKeyToken.ToUpper ' Same as Acceptance
    ''        If TargetEnv = "F"c Then ReplaceItem = My.Settings.FinancePublicKeyToken.ToUpper
    ''        If TargetEnv = "P"c Then ReplaceItem = My.Settings.ProdPublicKeyToken.ToUpper
    ''        If TargetEnv = "V"c Then ReplaceItem = My.Settings.ProdPublicKeyToken.ToUpper
    ''        If TargetEnv = "R"c Then ReplaceItem = My.Settings.ProdPublicKeyToken.ToUpper ' Same as Prod
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.LocalPublicKeyToken) Then FindItem = My.Settings.LocalPublicKeyToken
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.TestPublicKeyToken) Then FindItem = My.Settings.TestPublicKeyToken
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.AcceptPublicKeyToken) Then FindItem = My.Settings.AcceptPublicKeyToken
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.FinancePublicKeyToken) Then FindItem = My.Settings.FinancePublicKeyToken
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.ProdPublicKeyToken) Then FindItem = My.Settings.ProdPublicKeyToken
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for path names ---
    ''        If TargetEnv = "C"c Then ReplaceItem = My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower
    ''        If TargetEnv = "D"c Then ReplaceItem = "#PATH#"
    ''        If TargetEnv = "L"c Then ReplaceItem = My.Settings.LocalPath.Replace("*", GetUserNameAdj)
    ''        If TargetEnv = "T"c Then ReplaceItem = My.Settings.TestPath
    ''        If TargetEnv = "A"c Then ReplaceItem = My.Settings.AcceptPath
    ''        If TargetEnv = "X"c Then ReplaceItem = My.Settings.AcceptDNPath
    ''        If TargetEnv = "F"c Then ReplaceItem = My.Settings.FinancePath
    ''        If TargetEnv = "P"c Then ReplaceItem = My.Settings.ProdPath
    ''        If TargetEnv = "V"c Then ReplaceItem = My.Settings.ProdPath2
    ''        If TargetEnv = "R"c Then ReplaceItem = My.Settings.ReportPath
    ''        If ReplaceItem.StartsWith("\\") Then
    ''            ReplaceItem = ReplaceItem.ToLower
    ''        End If
    ''        If ContainsIgnoreCase(CurrLine, "<PublishUrl>") AndAlso ContainsIgnoreCase(ReplaceItem, "$") Then
    ''            ReplaceItem = ReplaceItem.Replace("$", "%24")
    ''        End If
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, "<PublishUrl>") OrElse
    ''            ContainsIgnoreCase(CurrLine, "<value") OrElse
    ''            ContainsIgnoreCase(CurrLine, "DefaultSettingValueAttribute") Then
    ''            If ContainsIgnoreCase(CurrLine, My.Settings.FinancePath) Then
    ''                FindItem = My.Settings.FinancePath
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower) Then
    ''                FindItem = My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower.Replace("$", "%24")) Then
    ''                FindItem = My.Settings.PCPath.Replace("LOCALHOST", My.Computer.Name).ToLower.Replace("$", "%24")
    ''            ElseIf ContainsIgnoreCase(CurrLine, "#PATH#") Then
    ''                FindItem = "#PATH#"
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.LocalPath.Replace("*", GetUserNameAdj)) Then
    ''                FindItem = My.Settings.LocalPath.Replace("*", GetUserNameAdj)
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.TestPath) Then
    ''                FindItem = My.Settings.TestPath
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.AcceptPath) Then
    ''                FindItem = My.Settings.AcceptPath
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.AcceptDNPath) Then
    ''                FindItem = My.Settings.AcceptDNPath
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.ProdPath) Then
    ''                FindItem = My.Settings.ProdPath
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.ProdPath2) Then
    ''                FindItem = My.Settings.ProdPath2
    ''            ElseIf ContainsIgnoreCase(CurrLine, My.Settings.ReportPath) Then
    ''                FindItem = My.Settings.ReportPath
    ''            End If
    ''            If ComboApplication.SelectedIndex >= 0 Then
    ''                If ComboApplication.Items(ComboApplication.SelectedIndex).ToString.ToUpper = "ARENA" Then
    ''                    If SourceEnv = "F"c AndAlso TargetEnv <> "F"c Then
    ''                        FindItem += "Finance\Arena\"
    ''                        ReplaceItem += "Arena\"
    ''                    ElseIf TargetEnv = "F"c Then
    ''                        FindItem += "Arena\"
    ''                        ReplaceItem += "Finance\Arena\"
    ''                    End If
    ''                End If
    ''            End If
    ''        End If
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        ElseIf ContainsIgnoreCase(CurrLine, "<InstallFrom>") Then
    ''            If (ReplaceItem.StartsWith("\\") OrElse ReplaceItem = "#PATH#") AndAlso ContainsIgnoreCase(CurrLine, ">Disk<") Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, ">Disk<", ">Unc<")
    ''                LineChanged = True
    ''            ElseIf Not (ReplaceItem.StartsWith("\\") OrElse ReplaceItem = "#PATH#") AndAlso ContainsIgnoreCase(CurrLine, ">Unc<") Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, ">Unc<", ">Disk<")
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for drive names ---
    ''        If TargetEnv = "C"c Then ReplaceItem = My.Settings.PCDrive.ToUpper
    ''        If TargetEnv = "D"c Then ReplaceItem = "#DRIVE#"
    ''        If TargetEnv = "L"c Then ReplaceItem = My.Settings.LocalDrive.ToUpper
    ''        If TargetEnv = "T"c Then ReplaceItem = My.Settings.TestDrive.ToUpper
    ''        If TargetEnv = "A"c Then ReplaceItem = My.Settings.AcceptDrive.ToUpper
    ''        If TargetEnv = "X"c Then ReplaceItem = My.Settings.AcceptDNDrive.ToUpper
    ''        If TargetEnv = "F"c Then ReplaceItem = My.Settings.FinanceDrive.ToUpper
    ''        If TargetEnv = "P"c Then ReplaceItem = My.Settings.ProdDrive.ToUpper
    ''        If TargetEnv = "V"c Then ReplaceItem = My.Settings.ProdDrive2.ToUpper
    ''        If TargetEnv = "R"c Then ReplaceItem = My.Settings.ReportDrive.ToUpper
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.PCDrive) Then FindItem = My.Settings.PCDrive
    ''        If ContainsIgnoreCase(CurrLine, "#DRIVE#") Then FindItem = "#DRIVE#"
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.LocalDrive) Then FindItem = My.Settings.LocalDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.TestDrive) Then FindItem = My.Settings.TestDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.AcceptDrive) Then FindItem = My.Settings.AcceptDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.AcceptDNDrive) Then FindItem = My.Settings.AcceptDNDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.FinanceDrive) Then FindItem = My.Settings.FinanceDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.ProdDrive) Then FindItem = My.Settings.ProdDrive
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.ProdDrive2) Then FindItem = My.Settings.ProdDrive2
    ''        If ContainsIgnoreCase(CurrLine, My.Settings.ReportDrive) Then FindItem = My.Settings.ReportDrive
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for environment names with a hyphen ---
    ''        If TargetEnv = "C"c Then ReplaceItem = " - PC"
    ''        If TargetEnv = "D"c Then ReplaceItem = " - Local"
    ''        If TargetEnv = "L"c Then ReplaceItem = " - Local"
    ''        If TargetEnv = "T"c Then ReplaceItem = " - Test"
    ''        If TargetEnv = "A"c Then ReplaceItem = " - Accept"
    ''        If TargetEnv = "X"c Then ReplaceItem = " - AcceptDN"
    ''        If TargetEnv = "F"c Then ReplaceItem = " - Finance"
    ''        If TargetEnv = "P"c Then ReplaceItem = ""
    ''        If TargetEnv = "V"c Then ReplaceItem = ""
    ''        If TargetEnv = "R"c Then ReplaceItem = " - Report"
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, " - PC") Then FindItem = " - PC"
    ''        If ContainsIgnoreCase(CurrLine, " - Local") Then FindItem = " - Local"
    ''        If ContainsIgnoreCase(CurrLine, " - Test") Then FindItem = " - Test"
    ''        If ContainsIgnoreCase(CurrLine, " - AcceptDN") Then
    ''            FindItem = " - AcceptDN"
    ''        ElseIf ContainsIgnoreCase(CurrLine, " - Accept") Then
    ''            FindItem = " - Accept"
    ''        End If
    ''        If ContainsIgnoreCase(CurrLine, " - Finance") Then FindItem = " - Finance"
    ''        If ContainsIgnoreCase(CurrLine, " - Prod") Then FindItem = " - Prod"
    ''        If ContainsIgnoreCase(CurrLine, " - Report") Then FindItem = " - Report"
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for environment names in ProductName ---
    ''        If TargetEnv = "C"c Then ReplaceItem = " - PC</ProductName>"
    ''        If TargetEnv = "D"c Then ReplaceItem = " - Local</ProductName>"
    ''        If TargetEnv = "L"c Then ReplaceItem = " - Local</ProductName>"
    ''        If TargetEnv = "T"c Then ReplaceItem = " - Test</ProductName>"
    ''        If TargetEnv = "A"c Then ReplaceItem = " - Accept</ProductName>"
    ''        If TargetEnv = "X"c Then ReplaceItem = " - AcceptDN</ProductName>"
    ''        If TargetEnv = "F"c Then ReplaceItem = " - Finance</ProductName>"
    ''        If TargetEnv = "P"c Then ReplaceItem = "</ProductName>"
    ''        If TargetEnv = "V"c Then ReplaceItem = "</ProductName>"
    ''        If TargetEnv = "R"c Then ReplaceItem = " - Report</ProductName>"
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, " - PC</ProductName>") Then FindItem = " - PC</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - Local</ProductName>") Then FindItem = " - Local</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - Test</ProductName>") Then FindItem = " - Test</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - Accept</ProductName>") Then FindItem = " - Accept</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - AcceptDN</ProductName>") Then FindItem = " - AcceptDN</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - Finance</ProductName>") Then FindItem = " - Finance</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, " - Report</ProductName>") Then FindItem = " - Report</ProductName>"
    ''        If ContainsIgnoreCase(CurrLine, "</ProductName>") AndAlso
    ''            Not ContainsIgnoreCase(CurrLine, " - ") AndAlso
    ''            Not ContainsIgnoreCase(CurrLine, "Microsoft") AndAlso
    ''            Not ContainsIgnoreCase(CurrLine, "Report Viewer") AndAlso
    ''            Not ContainsIgnoreCase(CurrLine, "Windows Installer") AndAlso
    ''            Not ContainsIgnoreCase(CurrLine, "Framework") Then
    ''            FindItem = "</ProductName>"
    ''        End If
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for environment names in xml ---
    ''        If TargetEnv = "C"c Then ReplaceItem = ">PC<"
    ''        If TargetEnv = "D"c Then ReplaceItem = ">Local<"
    ''        If TargetEnv = "L"c Then ReplaceItem = ">Local<"
    ''        If TargetEnv = "T"c Then ReplaceItem = ">Test<"
    ''        If TargetEnv = "A"c Then ReplaceItem = ">Accept<"
    ''        If TargetEnv = "X"c Then ReplaceItem = ">AcceptDN<"
    ''        If TargetEnv = "F"c Then ReplaceItem = ">Finance<"
    ''        If TargetEnv = "P"c Then ReplaceItem = ">Prod<"
    ''        If TargetEnv = "V"c Then ReplaceItem = ">Prod<"
    ''        If TargetEnv = "R"c Then ReplaceItem = ">Report<"
    ''        FindItem = ""
    ''        If ContainsIgnoreCase(CurrLine, ">PC<") Then FindItem = ">PC<"
    ''        If ContainsIgnoreCase(CurrLine, ">Local<") Then FindItem = ">Local<"
    ''        If ContainsIgnoreCase(CurrLine, ">Test<") Then FindItem = ">Test<"
    ''        If ContainsIgnoreCase(CurrLine, ">Accept<") Then FindItem = ">Accept<"
    ''        If ContainsIgnoreCase(CurrLine, ">AcceptDN<") Then FindItem = ">AcceptDN<"
    ''        If ContainsIgnoreCase(CurrLine, ">Finance<") Then FindItem = ">Finance<"
    ''        If ContainsIgnoreCase(CurrLine, ">Prod<") Then FindItem = ">Prod<"
    ''        If ContainsIgnoreCase(CurrLine, ">Report<") Then FindItem = ">Report<"
    ''        If FindItem <> "" AndAlso FindItem <> ReplaceItem Then
    ''            If ContainsIgnoreCase(CurrLine, FindItem) Then
    ''                CurrLine = ReplaceIgnoreCase(CurrLine, FindItem, ReplaceItem)
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for Reference Includes ---
    ''        If CurrLine IsNot Nothing AndAlso CurrLine.Trim.StartsWith("<Reference Include=""") AndAlso
    ''            CurrLine.IndexOf("/>") < 0 AndAlso
    ''            CurrLine.IndexOf("""Microsoft.") < 0 Then
    ''            If CurrLine.IndexOf(","c) >= 0 AndAlso ContainsIgnoreCase(CurrLine, "PublicKeyToken=") Then
    ''                CurrLine = CurrLine.Substring(0, CurrLine.IndexOf(","c)) + """>"
    ''                LineChanged = True
    ''            End If
    ''        End If
    ''        ' --- Check for DebugType = pdb-only/full lines ---
    ''        If CurrLine IsNot Nothing AndAlso CurrLine.Trim = "<DebugType>pdbonly</DebugType>" Then
    ''            CurrLine = CurrLine.Replace("pdbonly", "None")
    ''            LineChanged = True
    ''        End If
    ''        If CurrLine IsNot Nothing AndAlso CurrLine.Trim = "<DebugType>full</DebugType>" Then
    ''            CurrLine = CurrLine.Replace("full", "Full")
    ''            LineChanged = True
    ''        End If
    ''        ' --- Check for SpecifcVersion "False" lines ---
    ''        If CurrLine IsNot Nothing AndAlso CurrLine.Trim = "<SpecificVersion>False</SpecificVersion>" Then
    ''            CurrLine = Nothing
    ''            LineChanged = True
    ''            RemoveNothings = True
    ''        End If
    ''        ' --- Check for incorrect Option lines ---
    ''        If CurrLine IsNot Nothing Then
    ''            If CurrLine.Trim = "<OptionStrict>Off</OptionStrict>" Then
    ''                CurrLine = CurrLine.Replace(">Off<", ">On<")
    ''                LineChanged = True
    ''            End If
    ''            If CurrLine.Trim = "<OptionInfer>On</OptionInfer>" Then
    ''                CurrLine = CurrLine.Replace(">On<", ">Off<")
    ''                LineChanged = True
    ''            End If
    ''            ' --- Check for PlatformTarget = AnyCPU lines ---
    ''            If CurrLine.Trim = "<PlatformTarget>AnyCPU</PlatformTarget>" Then
    ''                CurrLine = CurrLine.Replace("AnyCPU", "x86")
    ''                LineChanged = True
    ''            End If
    ''            If CurrLine.Trim = "<PlatformTarget>x64</PlatformTarget>" Then
    ''                CurrLine = CurrLine.Replace("x64", "x86")
    ''                LineChanged = True
    ''            End If
    ''            If CurrLine.Trim.StartsWith("<PlatformTarget>") Then
    ''                PlatformTarget = True
    ''            End If
    ''        End If
    ''        ' --- Check if any changes made ---
    ''        If LineChanged Then
    ''            Lines(CurrLineNum) = CurrLine
    ''            Changed = True
    ''        End If
    ''    Next
    ''    ' --- Check if there were any changes made ---
    ''    If Changed Then
    ''        ' --- Check if any lines were removed by setting them to Nothing ---
    ''        If RemoveNothings Then
    ''            Dim NewLines As New List(Of String)
    ''            For Each CurrLine In Lines
    ''                If CurrLine IsNot Nothing Then
    ''                    NewLines.Add(CurrLine)
    ''                End If
    ''            Next
    ''            Lines = NewLines.ToArray
    ''            RemoveNothings = False
    ''        End If
    ''        ' --- Write out the result ---
    ''        Try
    ''            File.WriteAllLines(Pathname + Filename, Lines, CurrEncoding)
    ''        Catch ex As Exception
    ''            Return False
    ''        End Try
    ''    End If
    ''    Return True
    ''End Function

    Private Sub FillComboApplication()
        ComboApplication.Items.Clear()
        AppPathList.Clear()
        Dim TempList() As String = My.Settings.AppPathList.Replace(vbCr, "").Replace(vbLf, "").Split(";"c)
        For Each TempItem As String In TempList
            If TempItem <> "" AndAlso TempItem.Contains("="c) Then
                ComboApplication.Items.Add(TempItem.Substring(0, TempItem.IndexOf("="c)))
                AppPathList.Add(TempItem)
            End If
        Next
        AppPathExtra.Clear()
        Dim TempExtras() As String = My.Settings.AppPathExtra.Replace(vbCr, "").Replace(vbLf, "").Split(";"c)
        For Each TempItem As String In TempExtras
            If TempItem <> "" AndAlso TempItem.Contains("="c) Then
                Dim Found As Boolean = False
                For Each AppName As String In ComboApplication.Items
                    If AppName = TempItem.Substring(0, TempItem.IndexOf("="c)) Then
                        Found = True
                        Exit For
                    End If
                Next
                If Not Found Then
                    ComboApplication.Items.Add(TempItem.Substring(0, TempItem.IndexOf("="c)))
                End If
                AppPathExtra.Add(TempItem)
            End If
        Next
    End Sub

    Private Sub ComboApplication_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboApplication.SelectedIndexChanged
        ComboFromDir.Enabled = False
        ComboToDir.Enabled = False
        ComboFromDir.Text = ""
        ComboToDir.Text = ""
        ComboFromDir.Items.Clear()
        ComboToDir.Items.Clear()
        StatusLabelCounts.Text = ""
        Dim CurrApp As String = ComboApplication.Text
        If CurrApp = "" Then Exit Sub
        ComboFromDir.Enabled = True
        ComboToDir.Enabled = True
        For Each TempAppLine As String In AppPathList
            If TempAppLine.ToUpper.StartsWith(CurrApp.ToUpper + "=") Then
                TempAppLine = TempAppLine.Substring((CurrApp + "=").Length)
                Dim TempPaths() As String = TempAppLine.Split("|"c)
                For Each CurrPath As String In TempPaths
                    AddComboFromToItems(CurrPath)
                Next
                Exit For
            End If
        Next
        For Each TempAppLine As String In AppPathExtra
            If TempAppLine.ToUpper.StartsWith(CurrApp.ToUpper + "=") Then
                TempAppLine = TempAppLine.Substring((CurrApp + "=").Length)
                Dim TempPaths() As String = TempAppLine.Split("|"c)
                For Each CurrPath As String In TempPaths
                    AddComboFromToItems(CurrPath)
                Next
            End If
        Next
        TabControlMain.SelectedIndex = 0
        ListBoxDiff.Items.Clear()
        ListBoxProjDiff.Items.Clear()
        ListBoxFrom.Items.Clear()
        ListBoxTo.Items.Clear()
        FillMyFileCompareSettings()
    End Sub

    Private Sub ComboApplication_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboApplication.TextChanged
        StatusLabelCounts.Text = ""
    End Sub

    Private Sub ComboFromDir_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboFromDir.SelectedIndexChanged
        StatusLabelCounts.Text = ""
        If String.IsNullOrWhiteSpace(ComboFromDir.Text) OrElse String.IsNullOrWhiteSpace(ComboToDir.Text) Then
            ButtonCompare.Enabled = False
        Else
            ButtonCompare.Enabled = (ComboFromDir.Text <> ComboToDir.Text)
        End If
    End Sub

    Private Sub ComboFromDir_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboFromDir.TextChanged
        StatusLabelCounts.Text = ""
        If String.IsNullOrWhiteSpace(ComboFromDir.Text) OrElse String.IsNullOrWhiteSpace(ComboToDir.Text) Then
            ButtonCompare.Enabled = False
        Else
            ButtonCompare.Enabled = (ComboFromDir.Text <> ComboToDir.Text)
        End If
    End Sub

    Private Sub ComboToDir_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboToDir.SelectedIndexChanged
        StatusLabelCounts.Text = ""
        If String.IsNullOrWhiteSpace(ComboFromDir.Text) OrElse String.IsNullOrWhiteSpace(ComboToDir.Text) Then
            ButtonCompare.Enabled = False
        Else
            ButtonCompare.Enabled = (ComboFromDir.Text <> ComboToDir.Text)
        End If
    End Sub

    Private Sub ComboToDir_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboToDir.TextChanged
        StatusLabelCounts.Text = ""
        If String.IsNullOrWhiteSpace(ComboFromDir.Text) OrElse String.IsNullOrWhiteSpace(ComboToDir.Text) Then
            ButtonCompare.Enabled = False
        Else
            ButtonCompare.Enabled = (ComboFromDir.Text <> ComboToDir.Text)
        End If
    End Sub

    Private Sub AddComboFromToItems(ByVal Value As String)
        Value = Value.Replace("*", GetUserNameAdj)
        If Directory.Exists(Value) Then
            If Not ComboFromDir.Items.Contains(Value) Then
                ComboFromDir.Items.Add(Value)
            End If
            If Not ComboToDir.Items.Contains(Value) Then
                ComboToDir.Items.Add(Value)
            End If
        End If
    End Sub

    Private Sub AddAppPathExtra(ByVal AppName As String, ByVal AppPath As String)
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If String.IsNullOrEmpty(AppName) Then Exit Sub
        If String.IsNullOrEmpty(AppPath) Then Exit Sub
        If Not Directory.Exists(AppPath) Then
            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Path Not Found: " + AppPath)
        End If
        For Each TempAppLine As String In AppPathList
            If TempAppLine.ToUpper.StartsWith(AppName.ToUpper + "=") Then
                TempAppLine = TempAppLine.Substring((AppName + "=").Length)
                Dim TempPaths() As String = TempAppLine.Split("|"c)
                For Each CurrPath As String In TempPaths
                    If CurrPath.ToUpper = AppPath.ToUpper Then
                        Exit Sub ' found
                    End If
                Next
                Exit For
            End If
        Next
        For TempIndex As Integer = 0 To AppPathExtra.Count - 1
            Dim TempAppLine As String = AppPathExtra.Item(TempIndex)
            If TempAppLine.ToUpper.StartsWith(AppName.ToUpper + "=") Then
                TempAppLine = TempAppLine.Substring((AppName + "=").Length)
                Dim TempPaths() As String = TempAppLine.Split("|"c)
                For Each CurrPath As String In TempPaths
                    If CurrPath.ToUpper = AppPath.ToUpper Then
                        Exit Sub ' found
                    End If
                Next
                ' --- Path Not Found ---
                AppPathExtra.Item(TempIndex) += "|" + AppPath
                SaveAppPathExtra()
                AddComboFromToItems(AppPath)
                Exit Sub
            End If
        Next
        ' --- Application not found in AppPathExtra ---
        AppPathExtra.Add(AppName + "=" + AppPath)
        SaveAppPathExtra()
        AddComboFromToItems(AppPath)
    End Sub

    Private Sub SaveAppPathExtra()
        Dim Result As New StringBuilder
        For Each TempAppLine As String In AppPathExtra
            Result.Append(TempAppLine)
            Result.Append(";")
            Result.Append(vbCrLf)
        Next
        My.Settings.AppPathExtra = Result.ToString
        My.Settings.Save()
    End Sub

    Private Sub UsernameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsernameToolStripMenuItem.Click
        Dim TempUserName As String = My.Settings.AltUserName
        If My.Settings.AltUserName = "" Then TempUserName = GetUserName()
        TempUserName = InputBox("Enter User Name for comparisons", "User Name Option", TempUserName).ToLower
        If TempUserName = "" Then TempUserName = GetUserName()
        My.Settings.AltUserName = TempUserName
        My.Settings.Save()
    End Sub

    Private Sub IgnoreSpacesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IgnoreSpacesToolStripMenuItem.Click
        IgnoreSpacesToolStripMenuItem.Checked = Not IgnoreSpacesToolStripMenuItem.Checked
        My.Settings.OptionIgnoreSpaces = IgnoreSpacesToolStripMenuItem.Checked
        My.Settings.Save()
    End Sub

    Private Sub IgnoreVersionsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles IgnoreVersionsToolStripMenuItem.Click
        IgnoreVersionsToolStripMenuItem.Checked = Not IgnoreVersionsToolStripMenuItem.Checked
        My.Settings.OptionIgnoreVersions = IgnoreVersionsToolStripMenuItem.Checked
        My.Settings.Save()
    End Sub

    Private Sub ToolStripMenuItemQuickCompareBinary_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemQuickCompareBinary.Click
        ToolStripMenuItemQuickCompareBinary.Checked = Not ToolStripMenuItemQuickCompareBinary.Checked
        My.Settings.QuickCompareBinary = ToolStripMenuItemQuickCompareBinary.Checked
        My.Settings.Save()
    End Sub

    Private Sub ButtonSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSelectAll.Click
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            SelectAll(ListBoxDiff)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            SelectAll(ListBoxProjDiff)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageFrom" Then
            SelectAll(ListBoxFrom)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageTo" Then
            SelectAll(ListBoxTo)
        End If
    End Sub

    Private Sub SelectAll(ByVal CurrListBox As ListBox)
        With CurrListBox
            .ClearSelected()
            For TempIndex As Integer = 0 To .Items.Count - 1
                .SelectedItems.Add(.Items(TempIndex))
            Next
        End With
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        If TabControlMain.SelectedTab.Name = "TabPageDiff" Then
            SelectAll(ListBoxDiff)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageProjDiff" Then
            SelectAll(ListBoxProjDiff)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageFrom" Then
            SelectAll(ListBoxFrom)
        ElseIf TabControlMain.SelectedTab.Name = "TabPageTo" Then
            SelectAll(ListBoxTo)
        End If
    End Sub

    Private Function GetUserNameAdj() As String
        If My.Settings.AltUserName <> "" Then
            Return My.Settings.AltUserName.ToLower
        Else
            Return GetUserName()
        End If
    End Function

    Private Sub AddAppToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddAppToolStripMenuItem.Click
        Dim NewApp As String
        NewApp = InputBox("Enter new application name: ", My.Application.Info.AssemblyName, "")
        If NewApp = "" Then Exit Sub
        Dim Found As Boolean = False
        For Each AppName As String In ComboApplication.Items
            If AppName = NewApp Then
                MessageBox.Show("Application already in list", My.Application.Info.AssemblyName, MessageBoxButtons.OK)
                Exit Sub
            End If
        Next
        ComboApplication.Items.Add(NewApp)
    End Sub

    Private Function FileEOLErrors(ByVal FileName As String) As Boolean
        Dim CurrByte As Integer = 0
        Dim LastByte As Integer = 0
        Dim FoundEOLError As Boolean = False
        Dim CurrFS As New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
        ' ------------------------------------------------------------------------------------
        Try
            CurrByte = CurrFS.ReadByte
            Do While CurrByte >= 0
                ' --- Check for binary file ---
                If FileCompareClass.BinaryChar(CurrByte) Then
                    CurrFS.Close()
                    Return False
                End If
                ' --- Check for CRCR or LFLF patterns ---
                If LastByte = CurrByte AndAlso (LastByte = 10 OrElse LastByte = 13) Then
                    FoundEOLError = True
                End If
                ' --- Read next byte ---
                LastByte = CurrByte
                CurrByte = CurrFS.ReadByte
            Loop
        Catch ex As Exception
        End Try
        CurrFS.Close()
        Return FoundEOLError
    End Function

    Private Function FixFileEOL(ByVal FileName As String) As Boolean
        Dim Lines() As String
        Dim CurrEncoding As Encoding
        ' --------------------------
        ' --- Read lines from file, which uses any old EOL it finds ---
        Try
            CurrEncoding = GetFileEncoding(FileName)
            Lines = File.ReadAllLines(FileName, CurrEncoding)
        Catch ex As Exception
            Return False
        End Try
        ' --- Write out the result, which fixes the EOL errors ---
        Try
            File.WriteAllLines(FileName, Lines, CurrEncoding)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Sub ExternalCompareProgramToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExternalCompareProgramToolStripMenuItem.Click
        Dim TempFormExternalCompare As New FormExternalCompare
        TempFormExternalCompare.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItemCompare_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemCompare.Click
        InitCompare()
    End Sub

    Private Sub ToolStripMenuItemIncludeTestProj_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemIncludeTestProj.Click
        ToolStripMenuItemIncludeTestProj.Checked = Not ToolStripMenuItemIncludeTestProj.Checked
        My.Settings.IncludeTestProjects = ToolStripMenuItemIncludeTestProj.Checked
        My.Settings.Save()
    End Sub

    Private Sub ToolStripMenuItemIgnoreMissingDirectoryContents_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemIgnoreMissingDirectoryContents.Click
        ToolStripMenuItemIgnoreMissingDirectoryContents.Checked = Not ToolStripMenuItemIgnoreMissingDirectoryContents.Checked
        My.Settings.IgnoreMissingDirectories = ToolStripMenuItemIgnoreMissingDirectoryContents.Checked
        My.Settings.Save()
    End Sub

    Private Sub ExcludeFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcludeFilesToolStripMenuItem.Click
        Dim Result As String
        Result = InputBox("Enter file extensions as "".abc;.xyz;"", Space to erase:", "Exclude Files", My.Settings.ExcludeFileList)
        If String.IsNullOrEmpty(Result) Then
            Result = My.Settings.ExcludeFileList
        End If
        My.Settings.ExcludeFileList = Result.Trim
        My.Settings.Save()
    End Sub

End Class
