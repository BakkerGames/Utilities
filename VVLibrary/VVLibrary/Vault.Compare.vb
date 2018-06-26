' -------------------------------------
' --- Vault.Compare.vb - 06/26/2018 ---
' -------------------------------------

' ----------------------------------------------------------------------------------------------------
' 06/26/2018 - SBakker
'            - Added IncludeSpecificationList to hold "IncludeDir" and "IncludeExt" specifications.
' 03/04/2017 - SBakker
'            - Added filename to error messages.
' 06/30/2016 - SBakker
'            - Perform better file checking, to determine if source is newer or vault is newer.
' 05/28/2016 - SBakker
'            - Added more graceful error handling.
' 05/23/2016 - SBakker
'            - Switched from old method with "_current" files to new method where everything is a
'              history file.
' ----------------------------------------------------------------------------------------------------

Imports System.IO

Partial Public Class Vault

    Public Event FileNotInVault(ByVal Filename As String)
    Public Event FileOnlyInVault(ByVal Filename As String)
    Public Event SourceFileIsNewer(ByVal Filename As String)
    Public Event VaultFileIsNewer(ByVal Filename As String)

    Public Sub CompareAll()
        If String.IsNullOrEmpty(SourcePath) Then
            Throw New ArgumentNullException(SourcePath)
        End If
        If String.IsNullOrEmpty(VaultPath) Then
            Throw New ArgumentNullException(VaultPath)
        End If
        ' --- Create ignore specification list ---
        BuildIgnoreSpecificationList()
        ' --- Backup starting at the top ---
        CompareAllRecursive(SourcePath, VaultPath)
    End Sub

    Private Sub CompareAllRecursive(ByVal CurrSourcePath As String, ByVal CurrVaultPath As String)
        Dim BaseFilename As String
        Dim BaseExtension As String
        Dim ModifiedDate As Date
        Dim HistoryDirectory As String
        Dim HistoryFilename As String
        ' ----------------------------
        If Not Directory.Exists(CurrVaultPath) Then
            RaiseEvent FileNotInVault(CurrSourcePath)
            Exit Sub
        End If
        ' --- Check all source files ---
        Dim SourceDirInfo As DirectoryInfo = New DirectoryInfo(CurrSourcePath)
        For Each CurrFileInfo As FileInfo In SourceDirInfo.GetFiles
            Try
                ' --- Ignore hidden files ---
                If (CurrFileInfo.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Continue For
                End If
                ' --- Ignore special files ---
                If CurrFileInfo.Name.StartsWith(".") Then
                    Continue For
                End If
                ' --- Ignore files from ignore list ---
                Dim IgnoreFlag As Boolean = False
                For Each CurrSpec As String In IgnoreSpecificationList
                    If CurrFileInfo.Name.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                        IgnoreFlag = True
                        Exit For
                    End If
                Next
                ' --- Include files previously ignored from include list ---
                If IgnoreFlag Then
                    For Each CurrSpec As String In IncludeSpecificationList
                        If CurrFileInfo.Name.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                            IgnoreFlag = False
                            Exit For
                        End If
                    Next
                End If
                If IgnoreFlag Then Continue For
                ' --- Extract info for later use ---
                BaseFilename = CurrFileInfo.Name
                BaseExtension = CurrFileInfo.Extension
                ' --- Check if directory exists ---
                HistoryDirectory = $"{CurrVaultPath}\{BaseFilename}"
                If Not Directory.Exists(HistoryDirectory) Then
                    RaiseEvent FileNotInVault(CurrFileInfo.FullName)
                    Continue For
                End If
                ' --- Get the filename of the history file ---
                ModifiedDate = File.GetLastWriteTimeUtc(CurrFileInfo.FullName)
                HistoryFilename = $"{HistoryDirectory}\{ModifiedDate.ToString(DateTimePattern)}{BaseExtension}"
                ' --- Check if file exists in history ---
                If FilenameIsLatest(HistoryFilename, HistoryDirectory) Then
                    Continue For
                End If
                If FileMatchesLatest(CurrFileInfo, HistoryDirectory) Then
                    Continue For
                End If
                ' --- Files are different ---
                If FilenameNewerThanLatest(HistoryFilename, HistoryDirectory) Then
                    RaiseEvent SourceFileIsNewer(CurrFileInfo.FullName)
                Else
                    RaiseEvent VaultFileIsNewer(CurrFileInfo.FullName)
                End If
            Catch ex As Exception
                RaiseEvent AccessError($"{ex.Message} - {CurrSourcePath}\{CurrFileInfo.Name}")
            End Try
        Next
        ' --- Check for deleted files in the Vault ---
        Dim VaultDirInfo As DirectoryInfo = New DirectoryInfo(CurrVaultPath)
        For Each CurrDirInfo As DirectoryInfo In VaultDirInfo.GetDirectories
            Dim TempSourceItem As String = $"{CurrSourcePath}\{CurrDirInfo.Name}"
            If Not Directory.Exists(TempSourceItem) AndAlso
               Not File.Exists(TempSourceItem) Then
                RaiseEvent FileOnlyInVault(CurrDirInfo.FullName)
            End If
        Next
        ' --- Now, recurse into all subdirectories ---
        For Each CurrDirInfo As DirectoryInfo In SourceDirInfo.GetDirectories
            ' --- Ignore hidden directories ---
            If (CurrDirInfo.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Continue For
            End If
            ' --- Ignore special directories ---
            If CurrDirInfo.Name.StartsWith("."c) Then
                Continue For
            End If
            ' --- Ignore directories from ignore list ---
            Dim IgnoreFlag As Boolean = False
            For Each CurrSpec As String In IgnoreSpecificationList
                If CurrDirInfo.FullName.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                    IgnoreFlag = True
                    Exit For
                End If
            Next
            ' --- Include directories previously ignored from include list ---
            If IgnoreFlag Then
                For Each CurrSpec As String In IncludeSpecificationList
                    If CurrDirInfo.Name.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                        IgnoreFlag = False
                        Exit For
                    End If
                Next
            End If
            If IgnoreFlag Then Continue For
            ' --- Recursively backup each subdirectory ---
            CompareAllRecursive(CurrDirInfo.FullName, $"{CurrVaultPath}\{CurrDirInfo.Name}")
        Next
    End Sub

End Class
