' ---------------------------------
' --- FileUtils.vb - 03/14/2014 ---
' ---------------------------------

' ----------------------------------------------------------------------------------------------------
' 03/14/2014 - SBakker
'            - Expanded GetFileEncoding() to search the entire file for valid UTF-8 encoding sequences
'              within the file but without a leading BOM. This is a valid and common UTF-8 standard,
'              to not have a BOM. Also checks for CanBeASCII if all chars are 1-126, and CanBe1252 if
'              no invalid bytes are found. Returns Nothing if the file is binary or empty or missing.
'              Will also stop searching if the file is determined to be Binary.
'              NOTE: Does not handle Unicode values over 0xFFFF in UTF-8 - will consider them Binary.
'              NOTE: Considers 0 and 127 to not be valid ASCII or 1252 bytes, so are probably Binary.
' 01/07/2011 - SBakker
'            - Moved GetFileEncoding() into Arena_Utilities so it can be used everywhere.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text

Public Class FileUtils

    ''' <summary>
    ''' Determine the Text Encoding for the specified file. Returns Nothing if Binary.
    ''' </summary>
    Public Shared Function GetFileEncoding(ByVal Filename As String) As Encoding
        ' NOTE: Encoding.GetEncoding(1252) = Windows-1252, aka CP-1252 or ASCII-8, used for Batch files or plain text files.
        Dim Result As Encoding
        Dim CurrFS As System.IO.FileStream
        Dim BOM As Byte() = New Byte(3) {} ' Byte Order Mark
        ' --------------------------------------------------
        Result = Nothing
        ' --- Try to determine the encoding from the file ---
        Try
            CurrFS = New System.IO.FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.Read)
            If CurrFS.Length = 0 Then
                Return Nothing
            End If
            Try
                Dim CurrByte As Integer = CurrFS.ReadByte
                Dim CurrIndex As Integer = 0
                Do While CurrByte >= 0 AndAlso CurrIndex <= 3
                    BOM(CurrIndex) = CByte(CurrByte)
                    CurrByte = CurrFS.ReadByte
                    CurrIndex += 1
                Loop
                Do While CurrIndex <= 3
                    BOM(CurrIndex) = 0
                    CurrIndex += 1
                Loop
                If BOM(0) = &HEF AndAlso BOM(1) = &HBB AndAlso BOM(2) = &HBF Then
                    Result = Encoding.UTF8
                ElseIf BOM(0) = Asc("<"c) AndAlso BOM(1) = Asc("?"c) AndAlso BOM(2) = Asc("x"c) AndAlso BOM(3) = Asc("m"c) Then ' xml file
                    Result = Encoding.UTF8
                ElseIf BOM(0) = &HFF AndAlso BOM(1) = &HFE Then
                    Result = Encoding.Unicode ' aka UTF-16
                ElseIf BOM(0) = &HFE AndAlso BOM(1) = &HFF Then
                    Result = Encoding.BigEndianUnicode ' aka UTF-16BE
                ElseIf BOM(0) = 0 AndAlso BOM(1) = 0 AndAlso BOM(2) = &HFE AndAlso BOM(3) = &HFF Then
                    Result = Encoding.UTF32
                End If
                CurrFS.Close()
            Catch ex As Exception
                ' --- Can't read BOM, file may be empty or too short ---
                CurrFS.Close()
            End Try
            ' --- Check for UTF-8 without BOM, ASCII, or CodePage 1252 ---
            If Result Is Nothing Then
                ' --- Reopen the file and look for valid embedded UTF-8 sequences ---
                CurrFS = New System.IO.FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim ContinuationBytes As Integer = 0
                Dim CanBeASCII As Boolean = True
                Dim CanBeUTF8 As Boolean = True
                Dim CanBe1252 As Boolean = True
                Dim CurrByte As Integer = CurrFS.ReadByte ' has to hold -1
                ' --- Read entire file, byte by byte, until an invalid sequence is found ---
                Do While CurrByte >= 0 AndAlso (CanBe1252 OrElse CanBeASCII OrElse CanBeUTF8)
                    If CanBeASCII Then
                        If CurrByte = 0 OrElse CurrByte >= 127 Then
                            CanBeASCII = False
                        End If
                    End If
                    If CanBe1252 Then
                        If CurrByte = 0 OrElse
                            CurrByte = 127 OrElse
                            CurrByte = 129 OrElse
                            CurrByte = 141 OrElse
                            CurrByte = 143 OrElse
                            CurrByte = 144 OrElse
                            CurrByte = 157 Then
                            CanBe1252 = False
                        End If
                    End If
                    If CanBeUTF8 Then
                        If CurrByte = 192 OrElse CurrByte = 193 OrElse CurrByte >= 240 Then
                            CanBeUTF8 = False
                        End If
                        If ContinuationBytes = 0 Then
                            If CurrByte >= 128 AndAlso CurrByte <= 191 Then
                                CanBeUTF8 = False
                            End If
                            ' --- Only handle 2-byte and 3-byte sequences. Longer sequences could be invalid and are unused in Unicode. ---
                            If CurrByte >= 194 AndAlso CurrByte <= 223 Then
                                ContinuationBytes = 1 ' 2-byte character
                            ElseIf CurrByte >= 224 AndAlso CurrByte <= 239 Then
                                ContinuationBytes = 2 ' 3-byte character
                            End If
                        Else
                            If CurrByte < 128 OrElse CurrByte > 191 Then
                                CanBeUTF8 = False
                            Else
                                ContinuationBytes -= 1
                            End If
                        End If
                    End If
                    CurrByte = CurrFS.ReadByte
                Loop
                CurrFS.Close()
                ' --- Check results from reading entire file ---
                If CanBeUTF8 AndAlso ContinuationBytes <> 0 Then
                    CanBeUTF8 = False
                End If
                If CanBeASCII Then
                    Result = Encoding.ASCII
                ElseIf CanBeUTF8 Then
                    Result = Encoding.UTF8
                ElseIf CanBe1252 Then
                    Result = Encoding.GetEncoding(1252)
                Else
                    Result = Nothing ' Binary
                End If
            End If
        Catch ex As Exception
            ' --- File not found ---
            Result = Nothing
        End Try
        Return Result
    End Function

End Class
