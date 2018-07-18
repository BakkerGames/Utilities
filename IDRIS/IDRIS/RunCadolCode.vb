' ------------------------------------
' --- RunCadolCode.vb - 07/06/2018 ---
' ------------------------------------

Imports System.IO

Module RunCadolCode

    Private CodePath As String = "D:\IDRIS\LOCAL\PROGRAMS\DEVICE00\"

    Public CodeVolName As String = ""
    Public CodeLibName As String = ""
    Public CodeProgNum As Integer = 0
    Public CodeLineNum As Integer = 0

    Public LibraryProgs(255) As List(Of String)

    Public ExitIDRIS As Boolean = False

    Public CallStack As New List(Of CallStackItem)

    Public Sub LoadLibrary(ByVal VolName As String, ByVal LibName As String)
        Dim TempProg As Integer
        ' ---------------------
        If String.IsNullOrWhiteSpace(VolName) Then Exit Sub
        If String.IsNullOrWhiteSpace(LibName) Then Exit Sub
        ' --- Check if a new library ---
        If CodeVolName <> VolName.ToUpper.Replace("/", "_") OrElse
            CodeLibName <> LibName.ToUpper.Replace("/", "_") Then
            CodeVolName = VolName.ToUpper.Replace("/", "_")
            CodeLibName = LibName.ToUpper.Replace("/", "_")
            For TempProg = 0 To 255
                LibraryProgs(TempProg) = Nothing
            Next
            For Each CurrFileName As String In Directory.EnumerateFiles(CodePath + CodeVolName + "\" + CodeLibName, "*.cvp")
                Try
                    Dim BaseFileName As String = CurrFileName.Substring(CurrFileName.LastIndexOf("\") + 1)
                    TempProg = CInt(BaseFileName.Substring(0, 3))
                    Dim TempLines As String() = File.ReadAllLines(CurrFileName)
                    LibraryProgs(TempProg) = New List(Of String)
                    For Each TempLine As String In TempLines
                        LibraryProgs(TempProg).Add(TempLine)
                    Next
                Catch ex As Exception
                    MessageBox.Show("Invalid file found: " + CurrFileName)
                    Exit Sub
                End Try
            Next
        End If
        CodeProgNum = 0
        CodeLineNum = 0
    End Sub

    Public Sub RunCode()
        Do
            If CodeProgNum < 0 OrElse CodeProgNum > 255 Then
                MessageBox.Show("Invalid Program Number: " + CodeProgNum.ToString)
                Exit Sub
            End If
            If CodeLineNum < 0 Then
                MessageBox.Show("Invalid Line Number: " + CodeLineNum.ToString)
                Exit Sub
            End If
            If CodeLineNum >= LibraryProgs(CodeProgNum).Count Then
                MessageBox.Show("Line Number past end of program: " + CodeLineNum.ToString)
                Exit Sub
            End If
            Dim CurrLine As String = LibraryProgs(CodeProgNum).Item(CodeLineNum)
            CodeLineNum += 1
            ' --- Parse the line of code ---
            ParseCode(CurrLine)
            Application.DoEvents()
        Loop Until ExitIDRIS
    End Sub

End Module
