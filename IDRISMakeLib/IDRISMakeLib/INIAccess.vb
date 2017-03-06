' --- INIAccess - 10/31/2006 ---

Imports System.IO

Module INIAccess

    Public Function GetINIValue(ByVal Filename As String, ByVal Section As String, ByVal ItemName As String) As String
        ' --- This will trim all strings and values during comparisons.
        ' --- Many text editors will unexpectedly trim the ends of lines
        ' --- during editing, so it is not advised to have values which
        ' --- end with significant spaces.
        Dim CurrLine As String
        Dim InSection As Boolean = False
        Dim Result As String = ""
        ' ------------------------------
        Try
            If Not File.Exists(Filename) Then
                Return Result
            End If
        Catch ex As Exception
            Return Result
        End Try
        Dim sr As New StreamReader(Filename)
        Do While Not sr.EndOfStream
            CurrLine = sr.ReadLine().Trim
            If CurrLine.StartsWith("[") AndAlso CurrLine.EndsWith("]") Then
                InSection = String.Equals(CurrLine.Substring(1, CurrLine.Length - 2).Trim, _
                                          Section.Trim, StringComparison.OrdinalIgnoreCase)
            ElseIf InSection And CurrLine <> "" Then
                If String.Equals(CurrLine.Substring(0, ItemName.Length), ItemName, _
                                 StringComparison.OrdinalIgnoreCase) Then
                    CurrLine = CurrLine.Substring(ItemName.Length).Trim
                    If CurrLine.StartsWith("=") Then
                        Result = CurrLine.Substring(1).Trim
                        Exit Do
                    End If
                End If
            End If
        Loop
        sr.Close()
        Return Result
    End Function

End Module
