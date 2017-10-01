' ------------------------------------
' --- Vault.Backup.vb - 09/30/2017 ---
' ------------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/30/2017 - SBakker
'            - Don't bother to SetAttributes to readonly.
' 09/07/2017 - SBakker
'            - Don't ignore ".gitignore" files.
' 08/31/2017 - SBakker
'            - Due to file datetime getting reset on Git rollbacks, changed comparisons to only check
'              if file exists at all. No longer makes duplicates with different datetimes. No longer
'              possible to reconstruct "current" from vault with any accuracy, as that is meaningless.
' 03/04/2017 - SBakker
'            - Added filename to error messages.
' 05/28/2016 - SBakker
'            - Added more graceful error handling.
' 05/23/2016 - SBakker
'            - Switched from old method with "_current" files to new method where everything is a
'              history file.
' ----------------------------------------------------------------------------------------------------

Imports System.IO

Partial Public Class Vault

    Public Event FileVaulted(ByVal Filename As String)

    Public Sub BackupSingle(ByVal Filename As String)
        Dim FullSourceFilename As String
        Dim BaseFilename As String
        Dim BaseExtension As String
        Dim ModifiedDate As Date
        Dim HistoryDirectory As String
        Dim HistoryFilename As String
        ' ----------------------------
        If String.IsNullOrEmpty(SourcePath) Then
            Throw New ArgumentNullException(SourcePath)
        End If
        If String.IsNullOrEmpty(VaultPath) Then
            Throw New ArgumentNullException(VaultPath)
        End If
        FullSourceFilename = $"{SourcePath}\{Filename}"
        Try
            If Not File.Exists(FullSourceFilename) Then
                Throw New SystemException($"File not found: {FullSourceFilename}")
            End If
            Dim CurrFileInfo As New FileInfo(FullSourceFilename)
            ' --- Ignore hidden files ---
            If (CurrFileInfo.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                Throw New SystemException($"File is hidden, cannot be vaulted: {FullSourceFilename}")
            End If
            ' --- Ignore special files ---
            If CurrFileInfo.Name.StartsWith(".") Then
                Throw New SystemException($"File is special, cannot be vaulted: {FullSourceFilename}")
            End If
            ' --- Ignore files from ignore list ---
            Dim IgnoreFlag As Boolean = False
            For Each CurrSpec As String In IgnoreSpecificationList
                If CurrFileInfo.Name.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                    IgnoreFlag = True
                    Exit For
                End If
            Next
            If IgnoreFlag Then
                Throw New SystemException($"File is to be ignored, cannot be vaulted: {FullSourceFilename}")
            End If
            ' --- Extract info for later use ---
            BaseFilename = CurrFileInfo.Name
            BaseExtension = CurrFileInfo.Extension
            ' --- Check if directory exists ---
            HistoryDirectory = $"{VaultPath}\{Filename}"
            If Not Directory.Exists(HistoryDirectory) Then
                Directory.CreateDirectory(HistoryDirectory)
            End If
            ' --- Get the filename of the history file ---
            ModifiedDate = File.GetLastWriteTimeUtc(CurrFileInfo.FullName)
            HistoryFilename = $"{HistoryDirectory}\{ModifiedDate.ToString(DateTimePattern)}{BaseExtension}"
            ' --- Check if file exists in history ---
            If File.Exists(HistoryFilename) Then
                Exit Sub
            End If
            If FileFoundInHistory(CurrFileInfo, HistoryDirectory) Then
                Exit Sub
            End If
            ' --- Now, copy the source file to history file ---
            File.Copy(CurrFileInfo.FullName, HistoryFilename)
            ' --- Mark the history file as ReadOnly but nothing else, not even Archive ---
            ''File.SetAttributes(HistoryFilename, FileAttributes.ReadOnly)
            ' --- Display the filename ---
            RaiseEvent FileVaulted(CurrFileInfo.FullName)
        Catch ex As Exception
            RaiseEvent AccessError($"{ex.Message} - {FullSourceFilename}")
        End Try
    End Sub

    Public Sub BackupAll()
        If String.IsNullOrEmpty(SourcePath) Then
            Throw New ArgumentNullException(SourcePath)
        End If
        If String.IsNullOrEmpty(VaultPath) Then
            Throw New ArgumentNullException(VaultPath)
        End If
        ' --- Create ignore specification list ---
        BuildIgnoreSpecificationList()
        ' --- Backup starting at the top ---
        BackupAllRecursive(SourcePath, VaultPath)
    End Sub

    Private Sub BackupAllRecursive(ByVal CurrSourcePath As String, ByVal CurrVaultPath As String)
        Dim BaseFilename As String
        Dim BaseExtension As String
        Dim ModifiedDate As Date
        Dim HistoryDirectory As String
        Dim HistoryFilename As String
        ' ----------------------------
        If Not Directory.Exists(CurrVaultPath) Then
            ' --- Create vault directory ---
            Directory.CreateDirectory(CurrVaultPath)
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
                    Dim DoIgnore As Boolean = True
                    If CurrFileInfo.Name.Equals(".gitignore") Then DoIgnore = False
                    If DoIgnore Then Continue For
                End If
                ' --- Ignore files from ignore list ---
                Dim IgnoreFlag As Boolean = False
                For Each CurrSpec As String In IgnoreSpecificationList
                    If CurrFileInfo.Name.EndsWith(CurrSpec, StringComparison.InvariantCultureIgnoreCase) Then
                        IgnoreFlag = True
                        Exit For
                    End If
                Next
                If IgnoreFlag Then Continue For
                ' --- Extract info for later use ---
                BaseFilename = CurrFileInfo.Name
                BaseExtension = CurrFileInfo.Extension
                ' --- Check if directory exists ---
                HistoryDirectory = $"{CurrVaultPath}\{BaseFilename}"
                If Not Directory.Exists(HistoryDirectory) Then
                    Directory.CreateDirectory(HistoryDirectory)
                End If
                ' --- Get the filename of the history file ---
                ModifiedDate = File.GetLastWriteTimeUtc(CurrFileInfo.FullName)
                HistoryFilename = $"{HistoryDirectory}\{ModifiedDate.ToString(DateTimePattern)}{BaseExtension}"
                ' --- Check if file exists in history ---
                If File.Exists(HistoryFilename) Then
                    Continue For
                End If
                If FileFoundInHistory(CurrFileInfo, HistoryDirectory) Then
                    Continue For
                End If
                ' --- Now, copy the source file to history file ---
                File.Copy(CurrFileInfo.FullName, HistoryFilename)
                ' --- Mark the history file as ReadOnly but nothing else, not even Archive ---
                ''File.SetAttributes(HistoryFilename, FileAttributes.ReadOnly)
                ' --- Display the filename ---
                RaiseEvent FileVaulted(CurrFileInfo.FullName)
            Catch ex As Exception
                RaiseEvent AccessError($"{ex.Message} - {CurrSourcePath} - {CurrVaultPath}")
            End Try
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
            If IgnoreFlag Then Continue For
            ' --- Recursively backup each subdirectory ---
            BackupAllRecursive(CurrDirInfo.FullName, $"{CurrVaultPath}\{CurrDirInfo.Name}")
        Next
    End Sub

End Class
