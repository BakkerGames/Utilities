' ------------------------------------
' --- Vault.Common.vb - 09/30/2017 ---
' ------------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/30/2017 - SBakker
'            - Ignore errors if can't write out MD5 file. Probably due to file path length problem.
'            - Don't bother to SetAttributes to readonly.
' 09/29/2017 - SBakker
'            - Added even better historical MD5 checking, so no MD5 hash will get calculated more than
'              once ever.
' 09/28/2017 - SBakker
'            - Updated to Arena.Common.JSON.
'            - Changed FileFoundInHistory() to save MD5 hashes on history files as separate text files.
'              This quickens future searches by skipping MD5 calcs.
' 02/22/2017 - SBakker
'            - Removed support for .vvignore file. Use .vvconfig file instead now.
' 02/11/2017 - SBakker
'            - Adding support for .vvconfig file.
' 06/30/2016 - SBakker
'            - Added FilenameIsLastest(), FilenameNewerThanLatest(), and FileMatchesLatest().
' 05/28/2016 - SBakker
'            - Added more graceful error handling.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports Arena.Common.JSON

Partial Public Class Vault

    Public Event AccessError(ByVal Message As String)

    Private Sub BuildIgnoreSpecificationList()
        ' --- Check for ignore specifications ---
        IgnoreSpecificationList = New List(Of String)
        If vvConfig IsNot Nothing Then
            Dim vvIgnoreDir As JArray = CType(vvConfig.GetValueOrNull("IgnoreDir"), JArray)
            For Each CurrLine As String In vvIgnoreDir
                If String.IsNullOrEmpty(CurrLine) Then
                    Continue For
                End If
                IgnoreSpecificationList.Add($"\{CurrLine}") ' Need leading slash for dirs
            Next
            Dim vvIgnoreExt As JArray = CType(vvConfig.GetValueOrNull("IgnoreExt"), JArray)
            For Each CurrLine As String In vvIgnoreExt
                If String.IsNullOrEmpty(CurrLine) Then
                    Continue For
                End If
                IgnoreSpecificationList.Add(CurrLine)
            Next
        End If
    End Sub

    Private Function FileFoundInHistory(ByVal SourceFileInfo As FileInfo,
                                        ByVal HistoryDirectory As String) As Boolean
        ' --- This routine finds if the source file aready exists in the history directory using MD5 hashes.
        ' --- It saves the MD5 hash as a new filename with the file contents being the hashed filename.
        ' --- This lets it skip recalculating MD5 hashes in the future for the same files and just ask if
        ' --- there is a filename which equals the MD5 hash string. That's much faster!
        Dim SourceMD5 As String = MD5Utilities.CalcMD5(SourceFileInfo.FullName)
        Dim SourceMD5Filename As String = $"{HistoryDirectory}\{SourceMD5}"
        ' --- Look for MD5 filename from an earlier search ---
        Try
            If File.Exists(SourceMD5Filename) Then
                Return True
            End If
        Catch ex As Exception
            ' --- Can't read MD5 filename, just ignore ---
        End Try
        ' --- Get lists of MD5 files and vault files ---
        Dim MD5FilenameList As New List(Of String)
        Dim VaultFilenameList As New List(Of String)
        Dim HistoryDirInfo As DirectoryInfo = New DirectoryInfo(HistoryDirectory)
        For Each TempHistFileInfo As FileInfo In HistoryDirInfo.GetFiles
            ' --- Handle MD5 filenames ---
            If TempHistFileInfo.Name.Length = 32 AndAlso Not TempHistFileInfo.Name.Contains("_") Then
                MD5FilenameList.Add(TempHistFileInfo.Name)
            ElseIf SourceFileInfo.Length = TempHistFileInfo.Length Then
                ' --- Only interested in files with matching lengths ---
                VaultFilenameList.Add(TempHistFileInfo.Name)
            End If
        Next
        ' --- Remove all files where MD5 was already calculated ---
        If VaultFilenameList.Count > 0 Then
            For Each MD5Filename As String In MD5FilenameList
                Dim TempFilename As String = File.ReadAllText($"{HistoryDirectory}\{MD5Filename}")
                For TempIndex As Integer = 0 To VaultFilenameList.Count - 1
                    If VaultFilenameList(TempIndex) = TempFilename Then
                        ' --- This file's MD5 has already been calculated ---
                        VaultFilenameList.RemoveAt(TempIndex)
                        Exit For
                    End If
                Next
            Next
        End If
        ' --- Now calculate and check MD5 for the remaining files ---
        For Each VaultFilename As String In VaultFilenameList
            Dim HistoryMD5 As String = MD5Utilities.CalcMD5($"{HistoryDirectory}\{VaultFilename}")
            Dim HistoryMD5Filename As String = $"{HistoryDirectory}\{HistoryMD5}"
            Try
                File.WriteAllText(HistoryMD5Filename, VaultFilename)
                ''File.SetAttributes(HistoryMD5Filename, FileAttributes.ReadOnly)
            Catch ex As Exception
                ' --- Can't write out MD5 filename, just ignore ---
            End Try
            If SourceMD5 = HistoryMD5 Then
                Return True
            End If
        Next
        ' --- Not found ---
        Return False
    End Function

    Private Function FilenameIsLatest(ByVal HistoryFilename As String,
                                      ByVal HistoryDirectory As String) As Boolean
        Dim LatestFullFilename As String = ""
        ' -----------------------------------
        Dim HistoryDirInfo As DirectoryInfo = New DirectoryInfo(HistoryDirectory)
        For Each TempHistFileInfo As FileInfo In HistoryDirInfo.GetFiles
            If LatestFullFilename < TempHistFileInfo.FullName Then
                LatestFullFilename = TempHistFileInfo.FullName
            End If
            If HistoryFilename < LatestFullFilename Then ' Found one later
                Return False
            End If
        Next
        If String.IsNullOrEmpty(LatestFullFilename) Then
            Return False
        End If
        If HistoryFilename = LatestFullFilename Then
            Return True
        End If
        Return False
    End Function

    Private Function FilenameNewerThanLatest(ByVal HistoryFilename As String,
                                             ByVal HistoryDirectory As String) As Boolean
        Dim LatestFullFilename As String = ""
        ' -----------------------------------
        Dim HistoryDirInfo As DirectoryInfo = New DirectoryInfo(HistoryDirectory)
        For Each TempHistFileInfo As FileInfo In HistoryDirInfo.GetFiles
            If LatestFullFilename < TempHistFileInfo.FullName Then
                LatestFullFilename = TempHistFileInfo.FullName
            End If
            If HistoryFilename < LatestFullFilename Then ' Found one later
                Return False
            End If
        Next
        If String.IsNullOrEmpty(LatestFullFilename) Then
            Return False
        End If
        If HistoryFilename > LatestFullFilename Then
            Return True
        End If
        Return False
    End Function

    Private Function FileMatchesLatest(ByVal SourceFileInfo As FileInfo,
                                       ByVal HistoryDirectory As String) As Boolean
        Dim SourceMD5 As String = Nothing
        Dim LatestFileLength As Long = 0
        Dim LatestFullFilename As String = ""
        ' -----------------------------------
        Dim HistoryDirInfo As DirectoryInfo = New DirectoryInfo(HistoryDirectory)
        For Each TempHistFileInfo As FileInfo In HistoryDirInfo.GetFiles
            If LatestFullFilename < TempHistFileInfo.FullName Then
                LatestFullFilename = TempHistFileInfo.FullName
                LatestFileLength = TempHistFileInfo.Length
            End If
        Next
        If String.IsNullOrEmpty(LatestFullFilename) Then
            Return False
        End If
        ' --- Check lengths first ---
        If SourceFileInfo.Length <> LatestFileLength Then
            Return False
        End If
        ' --- See if the source and target MD5 match ---
        SourceMD5 = MD5Utilities.CalcMD5(SourceFileInfo.FullName)
        If SourceMD5 = MD5Utilities.CalcMD5(LatestFullFilename) Then
            Return True
        End If
        Return False
    End Function

End Class
