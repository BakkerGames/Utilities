' ------------------------------------
' --- MD5Utilities.vb - 01/12/2015 ---
' ------------------------------------

' ----------------------------------------------------------------------------------------------------
' 01/12/2015 - SBakker
'            - Created MD5Utilities, so CalcMD5 isn't duplicated elsewhere.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Public Class MD5Utilities

    Public Shared Function CalcMD5(ByVal FileName As String) As String
        Dim MD5Result As String = ""
        Dim MD5Hasher As MD5 = MD5.Create
        Dim fs As FileStream = File.OpenRead(FileName)
        Dim Result() As Byte = MD5Hasher.ComputeHash(fs)
        fs.Close()
        Dim HexResult As New StringBuilder
        For Each b As Byte In Result
            HexResult.Append(b.ToString("x2"))
        Next
        MD5Result = HexResult.ToString
        Return MD5Result
    End Function

End Class
