' ------------------------------------
' --- Vault.Common.vb - 09/28/2017 ---
' ------------------------------------

' ----------------------------------------------------------------------------------------------------
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
        Dim SourceMD5 As String = MD5Utilities.CalcMD5(SourceFileInfo.FullName)
        Dim SourceMD5Filename As String = $"{HistoryDirectory}\{SourceMD5}"
        ' --- Look for MD5 filename from an earlier search ---
        If File.Exists(SourceMD5Filename) Then
            Return True
        End If
        Dim HistoryDirInfo As DirectoryInfo = New DirectoryInfo(HistoryDirectory)
        For Each TempHistFileInfo As FileInfo In HistoryDirInfo.GetFiles
            ' --- Ignore MD5 filenames ---
            If TempHistFileInfo.Name.Length = 32 AndAlso Not TempHistFileInfo.Name.Contains("_") Then
                Continue For
            End If
            ' --- Check lengths first ---
            If SourceFileInfo.Length <> TempHistFileInfo.Length Then
                Continue For
            End If
            ' --- See if the source and target MD5 match ---
            Dim HistoryMD5 As String = MD5Utilities.CalcMD5(TempHistFileInfo.FullName)
            Dim HistoryMD5Filename As String = $"{HistoryDirectory}\{HistoryMD5}"
            If Not File.Exists(HistoryMD5Filename) Then
                ' --- Save this MD5 filename for next time ---
                File.WriteAllText(HistoryMD5Filename, TempHistFileInfo.Name)
                File.SetAttributes(HistoryMD5Filename, FileAttributes.ReadOnly)
            End If
            If SourceMD5 = HistoryMD5 Then
                Return True
            End If
        Next
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
