' -------------------------------
' --- ModMain.vb - 01/29/2016 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' 01/29/2016 - SBakker
'            - Made VS 2015 "14.0" the only version allowed for compiling.
' 12/17/2015 - SBakker
'            - Added VS 2015 "14.0" version information.
' 10/21/2014 - SBakker
'            - Removed all versions of VS prior to 2014. Only the latest one can be used to compile.
' 09/17/2014 - SBakker
'            - Unset the Archive bit after updating Versions if it wasn't set before.
' 05/23/2014 - SBakker
'            - Check for datetime changes to "\App.Config", "\My Project\Settings.Designer.vb", and
'              "\My Project\Settings.settings".
' 03/24/2014 - SBakker
'            - Added "/s /e" to xcopy command in PublishAll.bat. We are now using resource files below
'              the Bin directory.
' 03/14/2014 - SBakker
'            - Expanded GetFileEncoding() to search the entire file for valid UTF-8 encoding sequences
'              within the file but without a leading BOM. This is a valid and common UTF-8 standard.
'              Also check if the result is Nothing, meaning the file is a Binary file and unreadable.
'              The current method used here internally wasn't working very reliably.
' 02/14/2014 - SBakker
'            - Changed to use "bin\x86\Release" for publishing purposes.
' 01/24/2014 - SBakker
'            - Allow "\\computername\" in the PublishPath, which is replaced with "\\%COMPUTERNAME%\"
'              in the PublishAll.bat file. Only necessary if programs are going to be published
'              locally on more than one computer.
' 11/05/2013 - SBakker
'            - Added ":eof" to end of BuildAll.bat in case it doesn't find a Visual Studio devenv.exe.
' 11/04/2013 - SBakker
'            - Added support for VS 2013.
' 10/25/2013 - SBakker
'            - Also allow "/nopub", because that's what I tried to use the first time.
' 10/23/2013 - SBakker
'            - Added a variable BuildPublish, along with command line options "/npub", "/nopublish",
'              "/pub", and "/publish". For now the variable defaults to True, but if we start putting
'              programs into a local user directory, it should be changed to False. PublishAll.bat
'              would then become an xcopy batch file to the correct directory and not changed here.
'            - Look for newest VS.NET version first.
'            - Switch to Arena versions of ConfigInfo, DataConn, and Utilities.
' 02/25/2013 - SBakker
'            - Remove the logfile lines which are " 1 up-to-date".
' 02/20/2013 - SBakker
'            - Switch to using VS 2010 first, then VS 2012, when setting BuildProg.
' 01/01/2013 - SBakker
'            - Check for Bin\Security\BuildAllConfig.txt and Bin\Security\PublishAllConfig.txt files
'              so that various versions of Visual Studio can be used with this one program without
'              needing changes.
' 09/12/2012 - SBakker
'            - Hardcoding in the version numbers to the BuildAll.bat file, instead of using
'              the current version number.
' 07/16/2012 - SBakker
'            - Added a "BackFillChange" flag to handle the case where a lower-level DLL is
'              changed, but doesn't bump up to the version number of higher-level projects.
'              Usually due to the last section of the version number not bumping up enough.
' 04/13/2012 - SBakker
'            - Added "/np", "/nopause" options, so that BuildAll.bat and PublishAll.bat
'              don't have a "PAUSE" in them, so can be called from other batch files.
' 02/13/2012 - SBakker
'            - Fix PublishPath if it has "%24" instead of "$". This is how UNC paths are
'              stored.
' 11/02/2011 - SBakker
'            - Ignore warning messages, such as "target is a local path".
' 09/02/2011 - SBakker
'            - Delete all additional files that go along with the EXE in the "bin\" dir.
' 06/30/2011 - SBakker
'            - Delete the published EXE files from "bin\", to prevent them from being run
'              directly. They should only be run as Click-Once files after publishing.
' 05/27/2011 - SBakker
'            - Added ClearVersions ("/c" or "/clear") option to reset all version numbers to
'              their lowest possible values, based on the dates of their source files.
' 02/22/2011 - SBakker
'            - Check if absolute or relative paths. Only add CurrFileVersion.DirectoryName
'              if relative.
'            - Ignore "C:\Windows" and "C:\Program Files" paths for "<Reference Include=".
' 02/15/2011 - SBakker
'            - Changed FINDSTR command to only display "Build: 0" messages, meaning the
'              compile failed.
' 01/20/2011 - SBakker
'            - Update PublisherName if it is incorrect.
' 01/13/2011 - SBakker
'            - Added check to see if CurrLine is Nothing. File.ReadAllLines() can return a
'              Nothing line in some cases, usually at the end of the file.
'            - Load the CompanyName from "...\Bin\Security\CompanyName.txt". Build the
'              CopyrightName from the CompanyName. Default CompanyName and CopyrightName
'              to the ones in this project if the file isn't found.
' 01/07/2011 - SBakker
'            - Update the CompanyName and the Copyright year if either of them is incorrect.
' 12/27/2010 - SBakker
'            - Added a parameter "/test" ("/t") to control the building of BuildTest.Bat.
'            - Added short forms for all of the parameters.
' 11/30/2010 - SBakker
'            - Added "/FORCE" parameter, to force the update of the version number for every
'              project. This is useful when projects are changed from the outside, yet don't
'              trigger a normal version update.
' 11/10/2010 - SBakker
'            - Added a CrLf after "pause", so that it won't appear as a difference.
'            - Added BumpedVersionUp flag and loop, so any changes will increment the new
'              version number, and then go back to the beginning.
' 10/22/2010 - SBakker
'            - Added AssemblyDateTime property, for comparing against file dates if there is
'              no file version (such as an embedded resource file).
' 07/28/2010 - SBakker
'            - Added "<EmbeddedResource Include=" and "<Content Include=" files for checking
'              if they are later than the project version number. Added URL decoding for any
'              filenames which had been URL encoded to protect special characters.
' 07/22/2010 - SBakker
'            - Remove extra information after "<Reference Include=" project names. It isn't
'              needed in VS 2010, and may be out of date or incorrect. Also remove the
'              <SpecificVersion> "False" lines, and squish out removed lines when saving.
' 07/09/2010 - SBakker
'            - Fixed so it doesn't increment revision unless it's necessary.
' 07/08/2010 - SBakker
'            - Handle "<Reference Include=" better. It will now find references
'              that have a HintPath and get the file's version from the DLL.
'              It can also handle references which don't have any attributes in
'              the "Reference Include" line, although this is rare.
' 07/02/2010 - SBakker
'            - Update the Revision number (last number in Version) if a file's
'              datetime is later than the project file's.
' 06/09/2010 - SBakker
'            - Pause before deleting log file. Then you can press Ctrl-C to break
'              out and read the log file.
' 06/03/2010 - SBakker
'            - Added My.Settings for VisualStudioVersion and FrameworkVersion.
' 06/02/2010 - SBakker
'            - Added support for Visual Studio 2010, if installed.
'            - Added automatic building of a BuildTest.bat for test projects.
' 05/24/2010 - SBakker
'            - Find dates of included files and use them to update version number.
'            - Renamed all the routines to make them more understandable.
' 01/25/2010 - SBakker
'            - Update "<Reference Include" versions with NewVersion, not Version.
' 01/15/2010 - SBakker
'            - Update "<Reference Include" version numbers in project files.
' 12/30/2009 - SBakker
'            - Display FullName, not FileName, for project file when saving.
' 12/28/2009 - SBakker
'            - Added the ability for this to create the PublishAll.bat file.
' 12/25/2009 - SBakker
'            - Created this program to update all version numbers in projects and
'              assemblies to be the maximum of itself or any referenced projects.
'            - Create the BuildAll.bat file to make sure it contains all projects.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FileUtils
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Web
Imports UpdateVersions.FileVersion

Module ModMain

    Private VersionList As New List(Of FileVersion)
    Private ForceUpdate As Boolean = False
    Private ClearVersions As Boolean = False
    Private NoPause As Boolean = False
    Private TodaysVersion As String = Today.Year.ToString + "." + Today.Month.ToString + "." + Today.Day.ToString + ".0"
    Private CompanyName As String = My.Application.Info.CompanyName
    Private CopyrightName As String = "Copyright © " + CompanyName + " " + Year(Today).ToString
    Private BuildAllConfig As String = ""
    Private PublishAllConfig As String = ""

    Public Sub Main()
        Dim StartDir As String = ""
        Dim AltDirs As String = ""
        Dim AltDirList() As String = Nothing
        Dim BuildDebug As Boolean = False
        Dim BuildTest As Boolean = False
        Dim BuildPublish As Boolean = True
        Dim BackFillChange As Boolean = False
        Dim TempFilename As String = ""
        Dim TempFileLines() As String = Nothing
        Dim CurrEncoding As Encoding = Nothing
        ' -------------------------------------
        ' --- Read through command line arguments ---
        For Each Arg As String In My.Application.CommandLineArgs
            If Arg.StartsWith("/") Then
                ' --- Handle options ---
                If Arg.ToLower.StartsWith("/a=") OrElse Arg.ToLower.StartsWith("/alt=") Then
                    If AltDirs <> "" Then AltDirs += vbTab
                    AltDirs += Arg.Substring(Arg.IndexOf("="c) + 1)
                ElseIf Arg.ToLower = "/d" OrElse Arg.ToLower = "/debug" Then
                    BuildDebug = True
                ElseIf Arg.ToLower = "/t" OrElse Arg.ToLower = "/test" Then
                    BuildTest = True
                ElseIf Arg.ToLower = "/f" OrElse Arg.ToLower = "/force" Then
                    ForceUpdate = True
                ElseIf Arg.ToLower = "/c" OrElse Arg.ToLower = "/clear" Then
                    ClearVersions = True
                ElseIf Arg.ToLower = "/np" OrElse Arg.ToLower = "/nopause" Then
                    NoPause = True
                ElseIf Arg.ToLower = "/npub" OrElse Arg.ToLower = "/nopub" OrElse Arg.ToLower = "/nopublish" Then
                    BuildPublish = False
                ElseIf Arg.ToLower = "/pub" OrElse Arg.ToLower = "/publish" Then
                    BuildPublish = True
                Else
                    Console.WriteLine("Error: Unknown parameter: " + Arg)
                    Exit Sub
                End If
            Else
                If StartDir = "" Then StartDir = Arg
            End If
        Next
        ' --- Check parameters ---
        If StartDir = "" Then
            Console.WriteLine("Error: Missing Starting Directory parameter")
            Exit Sub
        End If
        If Not Directory.Exists(StartDir) Then
            Console.WriteLine("Error: Starting Directory not found: " + StartDir)
            Exit Sub
        End If
        ' --- See if CompanyName.txt file exists ---
        TempFilename = StartDir + "\Bin\Security\CompanyName.txt"
        Try
            If File.Exists(TempFilename) Then
                CurrEncoding = GetFileEncoding(TempFilename)
                If CurrEncoding IsNot Nothing Then
                    TempFileLines = File.ReadAllLines(TempFilename, CurrEncoding)
                    CompanyName = TempFileLines(0)
                    CopyrightName = "Copyright © " + CompanyName + " " + Year(Today).ToString
                End If
            End If
        Catch ex As Exception
            ' --- Couldn't get CompanyName from a file ---
        End Try
        ' --- See if BuildAllConfig.txt file exists ---
        TempFilename = StartDir + "\Bin\Security\BuildAllConfig.txt"
        Try
            If File.Exists(TempFilename) Then
                CurrEncoding = GetFileEncoding(TempFilename)
                If CurrEncoding IsNot Nothing Then
                    TempFileLines = File.ReadAllLines(TempFilename, GetFileEncoding(TempFilename))
                    BuildAllConfig = TempFileLines(0)
                End If
            End If
        Catch ex As Exception
            ' --- Couldn't get BuildAllConfig name from a file ---
        End Try
        ' --- See if PublishAllConfig.txt file exists ---
        TempFilename = StartDir + "\Bin\Security\PublishAllConfig.txt"
        Try
            If File.Exists(TempFilename) Then
                CurrEncoding = GetFileEncoding(TempFilename)
                If CurrEncoding IsNot Nothing Then
                    TempFileLines = File.ReadAllLines(TempFilename, GetFileEncoding(TempFilename))
                    PublishAllConfig = TempFileLines(0)
                End If
            End If
        Catch ex As Exception
            ' --- Couldn't get PublishAllConfig from a file ---
        End Try
        ' --- Find all project files in subdirectories ---
        Console.WriteLine()
        Console.WriteLine("Finding project version numbers...")
        Console.WriteLine()
        Dim TopDir As New DirectoryInfo(StartDir)
        If Not FindProjectVersions(TopDir, False) Then
            Exit Sub
        End If
        ' --- Find any projects in alternate directories ---
        If AltDirs <> "" Then
            AltDirList = AltDirs.Split(CChar(vbTab))
            For Each AltDirName As String In AltDirList
                Dim AltDirDI As New DirectoryInfo(AltDirName)
                FindProjectVersions(AltDirDI, True)
            Next
        End If
        ' --- Find included file versions ---
        Console.WriteLine()
        Console.WriteLine("Finding included file version numbers...")
        Console.WriteLine()
        For CurrIndex As Integer = 0 To VersionList.Count - 1
            FindIncludedVersions(CurrIndex)
        Next
        ' --- Check referenced version numbers ---
        Console.WriteLine()
        Console.WriteLine("Finding referenced version numbers...")
        Console.WriteLine()
        Do
            BackFillChange = False
            For CurrIndex As Integer = 0 To VersionList.Count - 1
                VersionList(CurrIndex).Updated = False
            Next
            For CurrIndex As Integer = 0 To VersionList.Count - 1
                If FindReferencedVersions(CurrIndex) Then
                    BackFillChange = True
                End If
            Next
        Loop Until Not BackFillChange
        ' --- Save new version numbers ---
        Console.WriteLine()
        Console.WriteLine("Saving version numbers...")
        Console.WriteLine()
        Dim ChangedAnything As Boolean = False
        For CurrIndex As Integer = 0 To VersionList.Count - 1
            ChangedAnything = ChangedAnything Or SaveVersions(CurrIndex)
        Next
        ' --- Create the "BuildAll.bat" file ---
        CreateBuildAll(StartDir)
        ' --- Create the "BuildDebug.bat" file ---
        If BuildDebug Then
            CreateBuildDebug(StartDir)
        End If
        ' --- Create the "BuildTest.bat" file ---
        If BuildTest Then
            CreateBuildTest(StartDir)
        End If
        ' --- Create the "PublishAll.bat" file ---
        If BuildPublish Then
            CreatePublishAll(StartDir)
        End If
        ' --- Done ---
        If ChangedAnything Then Console.WriteLine()
        Console.WriteLine("*** Done ***")
        Console.WriteLine()
    End Sub

    Private Function FindProjectVersions(ByVal ThisDir As DirectoryInfo, ByVal IsAltProject As Boolean) As Boolean
        Dim Lines() As String
        Dim Version As String = ""
        Dim TempVer As String = ""
        Dim AssemblyFilePath As String = ""
        Dim IsExecutable As Boolean
        Dim PublishPath As String = ""
        Dim CurrEncoding As Encoding
        ' ---------------------------------
        For Each TempFile As FileInfo In ThisDir.GetFiles
            If Not TempFile.FullName.EndsWith(".vbproj") Then Continue For
            Dim TempFilePath As String = TempFile.FullName.Substring(0, TempFile.FullName.LastIndexOf("\"c) + 1)
            Console.WriteLine("Checking " + TempFile.FullName.Substring(TempFilePath.Length).Replace(".vbproj", "") + "...")
            ' --- Check if this project is a WinExe project ---
            CurrEncoding = GetFileEncoding(TempFile.FullName)
            If CurrEncoding Is Nothing Then Continue For ' Binary file
            Lines = File.ReadAllLines(TempFile.FullName, CurrEncoding)
            IsExecutable = False
            For Each CurrLine As String In Lines
                If CurrLine Is Nothing Then Continue For
                If CurrLine.IndexOf("<OutputType>WinExe</OutputType>") >= 0 Then
                    IsExecutable = True
                End If
                If CurrLine.IndexOf("<PublishUrl>") >= 0 AndAlso CurrLine.IndexOf("</PublishUrl>") >= 0 Then
                    PublishPath = CurrLine.Substring(CurrLine.IndexOf("<PublishUrl>") + Len("<PublishUrl>"))
                    PublishPath = Left(PublishPath, Len(PublishPath) - Len("</PublishUrl>"))
                    If PublishPath.Contains("%24") Then PublishPath = PublishPath.Replace("%24", "$")
                    If PublishPath = "publish\" Then PublishPath = ""
                End If
            Next
            ' --- Get the version number from the assembly ---
            AssemblyFilePath = TempFilePath + "My Project\AssemblyInfo.vb"
            If Not File.Exists(AssemblyFilePath) Then Continue For
            CurrEncoding = GetFileEncoding(AssemblyFilePath)
            If CurrEncoding Is Nothing Then Continue For ' Binary file
            Lines = File.ReadAllLines(AssemblyFilePath, CurrEncoding)
            Version = ""
            For Each CurrLine As String In Lines
                If CurrLine Is Nothing Then Continue For
                If CurrLine.StartsWith("<Assembly: AssemblyVersion(""") Then
                    TempVer = CurrLine.Replace("<Assembly: AssemblyVersion(""", "").Replace(""")>", "").Trim
                    If Version = "" Then
                        Version = TempVer
                    ElseIf Version <> TempVer Then
                        Console.WriteLine("Error: Mismatched versions: " + TempFile.FullName)
                        Return False
                    End If
                End If
                If CurrLine.StartsWith("<Assembly: AssemblyFileVersion(""") Then
                    TempVer = CurrLine.Replace("<Assembly: AssemblyFileVersion(""", "").Replace(""")>", "").Trim
                    If Version = "" Then
                        Version = TempVer
                    ElseIf Version <> TempVer Then
                        Console.WriteLine("Error: Mismatched versions: " + TempFile.FullName)
                        Return False
                    End If
                End If
            Next
            ' --- Check if Version found ---
            If Version = "" Then
                Console.WriteLine("Error: Version not found: " + TempFile.FullName)
                Return False
            End If
            ' --- Save Filename and Version ---
            Dim TempFileVersion As New FileVersion
            With TempFileVersion
                .FullName = TempFile.FullName
                .DirectoryName = TempFile.DirectoryName
                .AssemblyFullName = AssemblyFilePath
                .FileName = TempFile.FullName.Substring(TempFilePath.Length).Replace(".vbproj", "")
                .Version = Version
                If ForceUpdate Then
                    If FileVersion.ExpandedVersion(.Version) < FileVersion.ExpandedVersion(TodaysVersion) Then
                        .NewVersion = TodaysVersion
                    Else
                        .NewVersion = Version
                        .IncNewRevision()
                    End If
                ElseIf ClearVersions Then
                    .NewVersion = "0.0.0.0"
                Else
                    .NewVersion = Version
                End If
                .IsExecutable = IsExecutable
                .PublishPath = PublishPath.Replace("\\computername\", "\\%COMPUTERNAME%\")
                .IsAltProject = IsAltProject
                .AssemblyDateTime = File.GetLastWriteTime(AssemblyFilePath)
            End With
            VersionList.Add(TempFileVersion)
        Next
        For Each TempDir As DirectoryInfo In ThisDir.GetDirectories
            FindProjectVersions(TempDir, IsAltProject)
        Next
        Return True
    End Function

    Private Sub FindIncludedVersions(ByVal CurrIndex As Integer)
        Dim Lines() As String
        Dim CurrFileDate As DateTime
        Dim CurrFileVersion As FileVersion
        Dim IncludeFileName As String
        Dim IncludeFileDate As DateTime
        Dim IncludeFileVersion As String
        Dim CurrEncoding As Encoding
        ' --------------------------------
        CurrFileVersion = VersionList(CurrIndex)
        Console.WriteLine("Checking " + CurrFileVersion.FileName + "...")
        CurrEncoding = GetFileEncoding(CurrFileVersion.FullName)
        If CurrEncoding Is Nothing Then Exit Sub
        Lines = File.ReadAllLines(CurrFileVersion.FullName, CurrEncoding)
        CurrFileDate = File.GetLastWriteTime(CurrFileVersion.AssemblyFullName)
        For Each CurrLine As String In Lines
            If CurrLine Is Nothing Then Continue For
            ' --- Check for dates on source programs in the same directory as the project ---
            If CurrLine.Trim.StartsWith("<Compile Include=""") AndAlso CurrLine.IndexOf("My Project\") < 0 Then
                ' --- Get the name of the included source file ---
                IncludeFileName = CurrLine.Trim.Replace("<Compile Include=""", "")
                IncludeFileName = IncludeFileName.Substring(0, IncludeFileName.IndexOf(""""))
                ' --- Check if absolute path or relative path ---
                If (IncludeFileName.Length < 2) OrElse (Not IncludeFileName.StartsWith("\\") AndAlso Not IncludeFileName.Substring(1, 1) = ":") Then
                    IncludeFileName = CurrFileVersion.DirectoryName + "\" + IncludeFileName
                End If
                If Not File.Exists(IncludeFileName) Then
                    ' --- May be a filename with embedded URL encoded characters ---
                    IncludeFileName = HttpUtility.UrlDecode(IncludeFileName)
                End If
                If File.Exists(IncludeFileName) Then
                    IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
                    IncludeFileVersion = IncludeFileDate.Year.ToString + "." + IncludeFileDate.Month.ToString + "." + IncludeFileDate.Day.ToString + ".0"
                    If ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) < ExpandedMajorMinorBuild(IncludeFileVersion) Then
                        CurrFileVersion.NewVersion = IncludeFileVersion
                    ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                           ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) Then
                        If CurrFileDate < IncludeFileDate Then
                            CurrFileVersion.IncNewRevision()
                        End If
                    End If
                End If
            End If
        Next
        VersionList(CurrIndex) = CurrFileVersion
    End Sub

    Private Function FindReferencedVersions(ByVal CurrIndex As Integer) As Boolean
        Dim Lines() As String
        Dim SubFileName As String
        Dim CurrFileVersion As FileVersion
        Dim SubFileVersion As FileVersion
        Dim Found As Boolean
        Dim FindHintPath As Boolean = False
        Dim ReferencedFile As String
        Dim ReferencedVersion As FileVersionInfo
        Dim IncludeFileName As String
        Dim IncludeFileVersion As String
        Dim IncludeFileDate As Date
        Dim BackfillChange As Boolean = False
        Dim CurrEncoding As Encoding
        ' --------------------------------------
        CurrFileVersion = VersionList(CurrIndex)
        If CurrFileVersion.Updated Then Return False
        Console.WriteLine("Checking " + CurrFileVersion.FileName + "...")
        CurrEncoding = GetFileEncoding(CurrFileVersion.FullName)
        If CurrEncoding Is Nothing Then Return False
        Lines = File.ReadAllLines(CurrFileVersion.FullName, CurrEncoding)
        For Each CurrLine As String In Lines
            If CurrLine Is Nothing Then Continue For
            ' --- Check for version numbers of reference projects ---
            If CurrLine.Trim.StartsWith("<Reference Include=""") AndAlso CurrLine.IndexOf("/>") < 0 Then
                FindHintPath = False
                SubFileName = CurrLine.Trim.Replace("<Reference Include=""", "")
                If CurrLine.IndexOf(","c) >= 0 Then
                    SubFileName = SubFileName.Substring(0, SubFileName.IndexOf(","))
                Else
                    SubFileName = SubFileName.Substring(0, SubFileName.IndexOf(""""))
                End If
                Found = False
                For SubIndex As Integer = 0 To VersionList.Count - 1
                    SubFileVersion = VersionList(SubIndex)
                    If SubFileVersion.FileName = SubFileName Then
                        Found = True
                        If FindReferencedVersions(SubIndex) Then
                            BackfillChange = True
                        End If
                        SubFileVersion = VersionList(SubIndex)
                        If CurrFileVersion.NewVersionLessThan(SubFileVersion.NewVersion) Then
                            CurrFileVersion.NewVersion = SubFileVersion.NewVersion
                        End If
                        If CurrFileVersion.Level < SubFileVersion.Level + 1 Then
                            CurrFileVersion.Level = SubFileVersion.Level + 1
                        End If
                        If SubFileVersion.NewVersionLessThan(CurrFileVersion.NewVersion) Then
                            If CurrFileVersion.NewVersion = CurrFileVersion.Version AndAlso SubFileVersion.NewVersion <> SubFileVersion.Version Then
                                CurrFileVersion.IncNewRevision()
                                SubFileVersion.NewVersion = CurrFileVersion.NewVersion
                                VersionList(SubIndex) = SubFileVersion
                                BackfillChange = True
                            End If
                        End If
                        Exit For
                    End If
                Next
                If Not Found AndAlso SubFileName <> "" Then
                    FindHintPath = True
                End If
            ElseIf CurrLine.Trim.StartsWith("<HintPath>") AndAlso FindHintPath Then
                ReferencedFile = CurrLine.Trim.Replace("<HintPath>", "").Replace("</HintPath>", "")
                ' --- Check if absolute path or relative path ---
                If (ReferencedFile.Length < 2) OrElse (Not ReferencedFile.StartsWith("\\") AndAlso Not ReferencedFile.Substring(1, 1) = ":") Then
                    ReferencedFile = CurrFileVersion.DirectoryName + "\" + ReferencedFile
                End If
                If ReferencedFile.ToUpper.StartsWith("C:\WINDOWS") OrElse ReferencedFile.ToUpper.StartsWith("C:\PROGRAM FILES") Then
                    FindHintPath = False
                    Continue For
                End If
                If File.Exists(ReferencedFile) Then
                    ReferencedVersion = FileVersionInfo.GetVersionInfo(ReferencedFile)
                    If CurrFileVersion.NewVersionLessThan(ReferencedVersion.FileVersion) Then
                        CurrFileVersion.NewVersion = ReferencedVersion.FileVersion
                    End If
                Else
                    Console.WriteLine("Error! Cannot locate file: " + ReferencedFile + " for project " + CurrFileVersion.AssemblyFullName)
                End If
            ElseIf CurrLine.Trim.StartsWith("</Reference>") Then
                FindHintPath = False
            End If
            ' --- Check for EmbeddedResource Includes ---
            If CurrLine.Trim.StartsWith("<EmbeddedResource Include=""") AndAlso CurrLine.IndexOf(""" />") >= 0 Then
                IncludeFileName = CurrLine.Trim.Replace("<EmbeddedResource Include=""", "").Replace(""" />", "")
                ' --- Check if absolute path or relative path ---
                If (IncludeFileName.Length < 2) OrElse (Not IncludeFileName.StartsWith("\\") AndAlso Not IncludeFileName.Substring(1, 1) = ":") Then
                    IncludeFileName = CurrFileVersion.DirectoryName + "\" + IncludeFileName
                End If
                If Not File.Exists(IncludeFileName) Then
                    ' --- May be a filename with embedded URL encoded characters ---
                    IncludeFileName = HttpUtility.UrlDecode(IncludeFileName)
                End If
                If File.Exists(IncludeFileName) Then
                    IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
                    IncludeFileVersion = IncludeFileDate.ToString("yyyy.M.d.0")
                    If CurrFileVersion.NewVersionLessThan(IncludeFileVersion) Then
                        CurrFileVersion.NewVersion = IncludeFileVersion
                    ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                           ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) AndAlso _
                           CurrFileVersion.AssemblyDateTime < IncludeFileDate Then
                        CurrFileVersion.IncNewRevision()
                    End If
                Else
                    Console.WriteLine("Error! Cannot locate file: " + IncludeFileName + " for project " + CurrFileVersion.AssemblyFullName)
                End If
            End If
            ' --- Check for Content Includes ---
            If CurrLine.Trim.StartsWith("<Content Include=""") AndAlso CurrLine.IndexOf(""" />") >= 0 Then
                IncludeFileName = CurrLine.Trim.Replace("<Content Include=""", "").Replace(""" />", "")
                ' --- Check if absolute path or relative path ---
                If (IncludeFileName.Length < 2) OrElse (Not IncludeFileName.StartsWith("\\") AndAlso Not IncludeFileName.Substring(1, 1) = ":") Then
                    IncludeFileName = CurrFileVersion.DirectoryName + "\" + IncludeFileName
                End If
                If Not File.Exists(IncludeFileName) Then
                    ' --- May be a filename with embedded URL encoded characters ---
                    IncludeFileName = HttpUtility.UrlDecode(IncludeFileName)
                End If
                If File.Exists(IncludeFileName) Then
                    IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
                    IncludeFileVersion = IncludeFileDate.ToString("yyyy.M.d.0")
                    If CurrFileVersion.NewVersionLessThan(IncludeFileVersion) Then
                        CurrFileVersion.NewVersion = IncludeFileVersion
                    ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                           ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) AndAlso _
                           CurrFileVersion.AssemblyDateTime < IncludeFileDate Then
                        CurrFileVersion.IncNewRevision()
                    End If
                Else
                    Console.WriteLine("Error! Cannot locate file: " + IncludeFileName + " for project " + CurrFileVersion.AssemblyFullName)
                End If
            End If
        Next
        ' --- Also check specific files to see if they have changed ---
        IncludeFileName = CurrFileVersion.DirectoryName + "\App.Config"
        If File.Exists(IncludeFileName) Then
            IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
            IncludeFileVersion = IncludeFileDate.ToString("yyyy.M.d.0")
            If CurrFileVersion.NewVersionLessThan(IncludeFileVersion) Then
                CurrFileVersion.NewVersion = IncludeFileVersion
            ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                   ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) AndAlso _
                   CurrFileVersion.AssemblyDateTime < IncludeFileDate Then
                CurrFileVersion.IncNewRevision()
            End If
        End If
        IncludeFileName = CurrFileVersion.DirectoryName + "\My Project\Settings.Designer.vb"
        If File.Exists(IncludeFileName) Then
            IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
            IncludeFileVersion = IncludeFileDate.ToString("yyyy.M.d.0")
            If CurrFileVersion.NewVersionLessThan(IncludeFileVersion) Then
                CurrFileVersion.NewVersion = IncludeFileVersion
            ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                   ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) AndAlso _
                   CurrFileVersion.AssemblyDateTime < IncludeFileDate Then
                CurrFileVersion.IncNewRevision()
            End If
        End If
        IncludeFileName = CurrFileVersion.DirectoryName + "\My Project\Settings.settings"
        If File.Exists(IncludeFileName) Then
            IncludeFileDate = File.GetLastWriteTime(IncludeFileName)
            IncludeFileVersion = IncludeFileDate.ToString("yyyy.M.d.0")
            If CurrFileVersion.NewVersionLessThan(IncludeFileVersion) Then
                CurrFileVersion.NewVersion = IncludeFileVersion
            ElseIf ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(IncludeFileVersion) AndAlso _
                   ExpandedMajorMinorBuild(CurrFileVersion.NewVersion) = ExpandedMajorMinorBuild(CurrFileVersion.Version) AndAlso _
                   CurrFileVersion.AssemblyDateTime < IncludeFileDate Then
                CurrFileVersion.IncNewRevision()
            End If
        End If
        CurrFileVersion.Updated = True
        VersionList(CurrIndex) = CurrFileVersion
        Return BackfillChange
    End Function

    Private Function SaveVersions(ByVal CurrIndex As Integer) As Boolean
        Dim Lines() As String
        Dim CurrLine As String
        Dim CurrEncoding As Encoding
        Dim Changed As Boolean = False
        Dim ChangedAnything As Boolean = False
        Dim RemoveNothings As Boolean = False
        Dim BumpedVersionUp As Boolean = False
        Dim CurrFileVersion As FileVersion
        ' ------------------------------------
        CurrFileVersion = VersionList(CurrIndex)
        ' --- Get current encoding so it is written out correctly ---
        CurrEncoding = GetFileEncoding(CurrFileVersion.FullName)
        If CurrEncoding Is Nothing Then Return False
        ' --- Fix project file ---
        Lines = File.ReadAllLines(CurrFileVersion.FullName, CurrEncoding)
        Changed = False
        Do
            BumpedVersionUp = False
            For TempIndex As Integer = 0 To Lines.GetUpperBound(0)
                CurrLine = Lines(TempIndex)
                If CurrLine Is Nothing Then Continue For
                If CurrLine.StartsWith("    <ApplicationRevision>") AndAlso _
                    CurrLine <> "    <ApplicationRevision>" + CurrFileVersion.NewVersionRevision + "</ApplicationRevision>" Then
                    Changed = True
                    If CurrFileVersion.Version = CurrFileVersion.NewVersion Then
                        CurrFileVersion.IncNewRevision()
                        BumpedVersionUp = True
                    End If
                    Lines(TempIndex) = "    <ApplicationRevision>" + CurrFileVersion.NewVersionRevision + "</ApplicationRevision>"
                End If
                If CurrLine.StartsWith("    <ApplicationVersion>") AndAlso _
                    CurrLine <> "    <ApplicationVersion>" + CurrFileVersion.NewVersion + "</ApplicationVersion>" Then
                    Changed = True
                    If CurrFileVersion.Version = CurrFileVersion.NewVersion Then
                        CurrFileVersion.IncNewRevision()
                        BumpedVersionUp = True
                    End If
                    Lines(TempIndex) = "    <ApplicationVersion>" + CurrFileVersion.NewVersion + "</ApplicationVersion>"
                End If
                If CurrLine.StartsWith("    <MinimumRequiredVersion>") AndAlso _
                    CurrLine <> "    <MinimumRequiredVersion>" + CurrFileVersion.NewVersion + "</MinimumRequiredVersion>" Then
                    Changed = True
                    If CurrFileVersion.Version = CurrFileVersion.NewVersion Then
                        CurrFileVersion.IncNewRevision()
                        BumpedVersionUp = True
                    End If
                    Lines(TempIndex) = "    <MinimumRequiredVersion>" + CurrFileVersion.NewVersion + "</MinimumRequiredVersion>"
                End If
                ' --- Check for Reference Includes ---
                If CurrLine.Trim.StartsWith("<Reference Include=""") AndAlso CurrLine.IndexOf("/>") < 0 Then
                    If CurrLine.IndexOf(","c) >= 0 AndAlso CurrLine.Contains("PublicKeyToken=") Then
                        Lines(TempIndex) = CurrLine.Substring(0, CurrLine.IndexOf(","c)) + """>"
                        Changed = True
                        If CurrFileVersion.Version = CurrFileVersion.NewVersion Then
                            CurrFileVersion.IncNewRevision()
                            BumpedVersionUp = True
                        End If
                    End If
                End If
                ' --- Check for SpecifcVersion "False" lines ---
                If CurrLine.Trim = "<SpecificVersion>False</SpecificVersion>" Then
                    Lines(TempIndex) = Nothing
                    Changed = True
                    RemoveNothings = True
                    If CurrFileVersion.Version = CurrFileVersion.NewVersion Then
                        CurrFileVersion.IncNewRevision()
                        BumpedVersionUp = True
                    End If
                End If
                ' --- Check for wrong PublisherName ---
                If CurrLine.StartsWith("    <PublisherName>") AndAlso CurrLine.EndsWith("</PublisherName>") AndAlso _
                    CurrLine <> "    <PublisherName>" + CompanyName + "</PublisherName>" Then
                    Lines(TempIndex) = "    <PublisherName>" + CompanyName + "</PublisherName>"
                    Changed = True
                End If
            Next
        Loop Until Not BumpedVersionUp
        If Changed Then
            ' --- Check if any lines were removed by setting them to Nothing ---
            If RemoveNothings Then
                Dim NewLines As New List(Of String)
                For Each CurrLine In Lines
                    If CurrLine IsNot Nothing Then
                        NewLines.Add(CurrLine)
                    End If
                Next
                Lines = NewLines.ToArray
                RemoveNothings = False
            End If
            ' --- Write out the result ---
            Try
                ChangedAnything = True
                Dim SaveAttributes As FileAttributes = File.GetAttributes(CurrFileVersion.FullName)
                File.WriteAllLines(CurrFileVersion.FullName, Lines, CurrEncoding)
                ' --- Unset the Archive bit if it wasn't set before ---
                If (SaveAttributes And FileAttributes.Archive) <> FileAttributes.Archive Then
                    File.SetAttributes(CurrFileVersion.FullName, SaveAttributes)
                End If
                Console.WriteLine(CurrFileVersion.FullName + " changed from " + CurrFileVersion.Version + " to " + CurrFileVersion.NewVersion)
            Catch ex As Exception
                Console.WriteLine("Error: Unable to update file: " + CurrFileVersion.FileName)
            End Try
        End If
        ' --- Get current encoding so it is written out correctly ---
        CurrEncoding = GetFileEncoding(CurrFileVersion.AssemblyFullName)
        If CurrEncoding Is Nothing Then Return False
        ' --- Fix assembly file ---
        Lines = File.ReadAllLines(CurrFileVersion.AssemblyFullName, CurrEncoding)
        Changed = False
        For TempIndex As Integer = 0 To Lines.GetUpperBound(0)
            CurrLine = Lines(TempIndex)
            If CurrLine Is Nothing Then Continue For
            If CurrLine.StartsWith("<Assembly: AssemblyVersion(""") AndAlso _
                CurrLine <> "<Assembly: AssemblyVersion(""" + CurrFileVersion.NewVersion + """)> " Then
                Lines(TempIndex) = "<Assembly: AssemblyVersion(""" + CurrFileVersion.NewVersion + """)> "
                Changed = True
            End If
            If CurrLine.StartsWith("<Assembly: AssemblyFileVersion(""") AndAlso _
                CurrLine <> "<Assembly: AssemblyFileVersion(""" + CurrFileVersion.NewVersion + """)> " Then
                Lines(TempIndex) = "<Assembly: AssemblyFileVersion(""" + CurrFileVersion.NewVersion + """)> "
                Changed = True
            End If
            If CurrLine.StartsWith("<Assembly: AssemblyCompany(""") AndAlso _
                CurrLine <> "<Assembly: AssemblyCompany(""" + CompanyName + """)> " Then
                Lines(TempIndex) = "<Assembly: AssemblyCompany(""" + CompanyName + """)> "
                Changed = True
            End If
            If CurrLine.StartsWith("<Assembly: AssemblyCopyright(""") AndAlso _
                CurrLine <> "<Assembly: AssemblyCopyright(""" + Left(CopyrightName, CopyrightName.Length - 4) + CurrFileVersion.NewVersion.Substring(0, 4) + """)> " Then
                Lines(TempIndex) = "<Assembly: AssemblyCopyright(""" + Left(CopyrightName, CopyrightName.Length - 4) + CurrFileVersion.NewVersion.Substring(0, 4) + """)> "
                Changed = True
            End If
        Next
        If Changed Then
            Try
                ChangedAnything = True
                Dim SaveAttributes As FileAttributes = File.GetAttributes(CurrFileVersion.AssemblyFullName)
                File.WriteAllLines(CurrFileVersion.AssemblyFullName, Lines, CurrEncoding)
                ' --- Unset the Archive bit if it wasn't set before ---
                If (SaveAttributes And FileAttributes.Archive) <> FileAttributes.Archive Then
                    File.SetAttributes(CurrFileVersion.AssemblyFullName, SaveAttributes)
                End If
                Console.WriteLine(CurrFileVersion.AssemblyFullName + " changed from " + CurrFileVersion.Version + " to " + CurrFileVersion.NewVersion)
            Catch ex As Exception
                Console.WriteLine("Error: Unable to update file: " + CurrFileVersion.AssemblyFullName)
            End Try
        End If
        Return ChangedAnything
    End Function

    Private Sub CreateBuildAll(ByVal StartDir As String)
        ' --- Create the "BuildAll.bat" file ---
        Dim FoundFile As Boolean = False
        Dim BuildAll As New StringBuilder
        Dim CurrFileVersion As FileVersion
        Dim Found As Boolean
        Dim CurrLevel As Integer = 0
        AddBuildHeader(BuildAll, "BuildAll")
        Do
            Found = False
            For CurrIndex As Integer = 0 To VersionList.Count - 1
                CurrFileVersion = VersionList(CurrIndex)
                With CurrFileVersion
                    If .Level = CurrLevel AndAlso Not .IsAltProject Then
                        ' --- Make sure the solution file is only one level down ---
                        If File.Exists(StartDir + "\" + .FileName + "\" + .FileName + ".sln") Then
                            If Not Found Then
                                BuildAll.AppendLine("")
                                BuildAll.AppendLine("REM")
                                BuildAll.AppendLine("REM --- Level " + .Level.ToString + " ---")
                                Found = True
                            End If
                            BuildAll.AppendLine("%buildprog% """ + .FileName + "\" + .FileName + ".sln"" /Out %logfile% /Build Release")
                            FoundFile = True
                        End If
                    End If
                End With
            Next
            CurrLevel += 1
        Loop While Found
        AddBuildFooter(BuildAll)
        If FoundFile Then
            Try
                File.WriteAllText(StartDir + "\BuildAll.bat", BuildAll.ToString)
            Catch ex As Exception
                Console.WriteLine("Error: Unable to create BuildAll.bat")
            End Try
        End If
    End Sub

    Private Sub CreateBuildDebug(ByVal StartDir As String)
        ' --- Create the "BuildDebug.bat" file ---
        Dim FoundFile As Boolean = False
        Dim BuildDebug As New StringBuilder
        Dim CurrFileVersion As FileVersion
        Dim Found As Boolean
        Dim CurrLevel As Integer = 0
        AddBuildHeader(BuildDebug, "BuildDebug")
        Do
            Found = False
            For CurrIndex As Integer = 0 To VersionList.Count - 1
                CurrFileVersion = VersionList(CurrIndex)
                With CurrFileVersion
                    If .Level = CurrLevel AndAlso Not .IsAltProject Then
                        ' --- Make sure the solution file is only one level down ---
                        If File.Exists(StartDir + "\" + .FileName + "\" + .FileName + ".sln") Then
                            If Not Found Then
                                BuildDebug.AppendLine("")
                                BuildDebug.AppendLine("REM")
                                BuildDebug.AppendLine("REM --- Level " + .Level.ToString + " ---")
                                Found = True
                            End If
                            BuildDebug.AppendLine("%buildprog% """ + .FileName + "\" + .FileName + ".sln"" /Out %logfile% /Build Debug")
                            FoundFile = True
                        End If
                    End If
                End With
            Next
            CurrLevel += 1
        Loop While Found
        AddBuildFooter(BuildDebug)
        If FoundFile Then
            Try
                File.WriteAllText(StartDir + "\BuildDebug.bat", BuildDebug.ToString)
            Catch ex As Exception
                Console.WriteLine("Error: Unable to create BuildDebug.bat")
            End Try
        End If
    End Sub

    Private Sub CreateBuildTest(ByVal StartDir As String)
        ' --- Create the "BuildTest.bat" file ---
        Dim FoundFile As Boolean = False
        Dim BuildTest As New StringBuilder
        Dim TopDir As New DirectoryInfo(StartDir)
        AddBuildHeader(BuildTest, "BuildTest")
        BuildTest.AppendLine()
        Dim CurrFileVersion As FileVersion
        For CurrIndex As Integer = 0 To VersionList.Count - 1
            CurrFileVersion = VersionList(CurrIndex)
            With CurrFileVersion
                If Not .IsAltProject Then
                    ' --- Make sure the solution file is more than one level down ---
                    If Not File.Exists(StartDir + "\" + .FileName + "\" + .FileName + ".sln") Then
                        BuildTest.AppendLine("%buildprog% """ + .DirectoryName.Replace(TopDir.FullName + "\", "") + ".sln"" /Out %logfile% /Build Release")
                        FoundFile = True
                    End If
                End If
            End With
        Next
        AddBuildFooter(BuildTest)
        If FoundFile Then
            Try
                File.WriteAllText(StartDir + "\BuildTest.bat", BuildTest.ToString)
            Catch ex As Exception
                Console.WriteLine("Error: Unable to create BuildTest.bat")
            End Try
        End If
    End Sub

    Private Sub CreatePublishAll(ByVal StartDir As String)
        ' --- Create the "PublishAll.bat" file ---
        Dim PublishAll As New StringBuilder
        With PublishAll
            If Not String.IsNullOrWhiteSpace(PublishAllConfig) Then
                .AppendLine(PublishAllConfig)
            Else
                .AppendLine("set buildprog=""C:\WINDOWS\Microsoft.NET\Framework\" + My.Settings.FrameworkVersion + "\msbuild.exe""")
            End If
        End With
        Dim CurrFileVersion As FileVersion
        Dim Found As Boolean
        Dim DelExeList As String = ""
        Do
            Found = False
            For CurrIndex As Integer = 0 To VersionList.Count - 1
                CurrFileVersion = VersionList(CurrIndex)
                With CurrFileVersion
                    ' --- Make sure the solution file is only one level down ---
                    If File.Exists(StartDir + "\" + .FileName + "\" + .FileName + ".sln") Then
                        If .IsExecutable And .PublishPath <> "" Then
                            PublishAll.AppendLine("")
                            ' --- Ignore warning messages, such as "target is a local path" ---
                            PublishAll.AppendLine("%buildprog% /p:Configuration=Release /t:Publish /v:minimal """ + .FileName + "\" + .FileName + ".sln"" | find /v "": warning """)
                            PublishAll.AppendLine("xcopy /d /s /y /s /e """ + .FileName + "\" + .FileName + "\bin\x86\Release\app.publish\*.*"" """ + .PublishPath + """")
                            DelExeList += "del bin\" + .FileName + ".*" + vbCrLf
                        End If
                    End If
                End With
            Next
        Loop While Found
        With PublishAll
            .AppendLine("")
            If Not String.IsNullOrWhiteSpace(DelExeList) Then
                .AppendLine(DelExeList)
            End If
            If Not NoPause Then
                .AppendLine("pause")
            End If
        End With
        Try
            File.WriteAllText(StartDir + "\PublishAll.bat", PublishAll.ToString)
        Catch ex As Exception
            Console.WriteLine("Error: Unable to create PublishAll.bat")
        End Try
    End Sub

    Private Sub AddBuildHeader(ByRef SB As StringBuilder, ByVal LogFileName As String)
        With SB
            .AppendLine("@echo off")
            If Not String.IsNullOrWhiteSpace(BuildAllConfig) Then
                .AppendLine(BuildAllConfig)
            Else
                .AppendLine("set buildprog=""C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe""")
                .AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files\Microsoft Visual Studio 11.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\devenv.exe""")
                ''.AppendLine("if not exist %buildprog% set buildprog=""C:\Program Files\Microsoft Visual Studio 10.0\Common7\IDE\devenv.exe""")
            End If
            .AppendLine("if not exist %buildprog% (")
            .AppendLine("echo Visual Studio compiler not found:")
            .AppendLine("echo %buildprog%")
            .AppendLine("echo.")
            If Not NoPause Then
                .AppendLine("pause")
            End If
            .AppendLine("goto :eof")
            .AppendLine(")")
            .AppendLine("")
            .AppendLine("set logfile=""" + LogFileName + ".log""")
            .AppendLine("attrib -r %logfile% >nul 2>nul")
            .AppendLine("del %logfile% >nul 2>nul")
            .AppendLine("")
            .AppendLine("@echo on")
        End With
    End Sub

    Private Sub AddBuildFooter(ByRef SB As StringBuilder)
        With SB
            .AppendLine("")
            .AppendLine("REM")
            .AppendLine("REM --- Show log file ---")
            .AppendLine("findstr /r /c:"" Build[ :][s ][t0]"" %logfile% | findstr /v /c:"" 1 up-to-date""")
            .AppendLine("")
            If Not NoPause Then
                .AppendLine("pause")
            End If
            .AppendLine("")
            .AppendLine("@del %logfile% >nul 2>nul")
            .AppendLine("")
            .AppendLine(":eof")
        End With
    End Sub

End Module
