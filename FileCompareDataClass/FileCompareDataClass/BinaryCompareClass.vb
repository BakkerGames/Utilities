' ------------------------------------------
' --- BinaryCompareClass.vb - 03/09/2016 ---
' ------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 03/09/2016 - SBakker
'            - Added logic to compare files using Length, LastWriteTimeUtc, and then MD5.
' 04/22/2015 - SBakker
'            - Changed to use MD5 hash for file comparisions.
' 03/13/2014 - SBakker
'            - Created straight Binary Compare Class to just see if two files exactly match. Jumps out
'              at first difference.
' ----------------------------------------------------------------------------------------------------

Imports MD5ClassLibrary.MD5Utilities
Imports System.IO

Public Class BinaryCompareClass

    Public Shared Function BinaryFilesMatch(ByVal Filename1 As String, ByVal Filename2 As String) As Boolean
        If String.IsNullOrWhiteSpace(Filename1) Then Return False
        If String.IsNullOrWhiteSpace(Filename2) Then Return False
        If Not File.Exists(Filename1) Then Return False
        If Not File.Exists(Filename2) Then Return False
        ' --- Check files using FileInfo and MD5 ---
        Dim FileOneInfo As New FileInfo(Filename1)
        Dim FileTwoInfo As New FileInfo(Filename2)
        Return FilesMatch(FileOneInfo, FileTwoInfo)
    End Function

    Private Shared Function FilesMatch(ByVal SourceFileInfo As FileInfo, ByVal TargetFileInfo As FileInfo) As Boolean
        Dim SourceMD5 As String
        Dim TargetMD5 As String
        ' ---------------------
        If SourceFileInfo.Length <> TargetFileInfo.Length Then
            Return False
        End If
        If SourceFileInfo.LastWriteTimeUtc = TargetFileInfo.LastWriteTimeUtc Then
            Return True
        End If
        Try
            SourceMD5 = CalcMD5(SourceFileInfo.FullName)
        Catch ex As Exception
            Console.WriteLine("Error accessing file: " + SourceFileInfo.FullName)
            Console.WriteLine()
            Return True ' Ignore file access errors
        End Try
        Try
            TargetMD5 = CalcMD5(TargetFileInfo.FullName)
        Catch ex As Exception
            Console.WriteLine("Error accessing file: " + TargetFileInfo.FullName)
            Console.WriteLine()
            Return True ' Ignore file access errors
        End Try
        If SourceMD5 = TargetMD5 Then
            Return True
        End If
        Return False
    End Function

End Class
