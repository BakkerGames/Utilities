' -------------------------------
' --- ModMain.vb - 10/24/2012 ---
' -------------------------------

Imports System.IO
Imports System.Text

Module ModMain

    Private Sub ShowSyntax()
        Console.WriteLine()
        Console.WriteLine("Syntax: " + My.Application.Info.ProductName + " <FromFilename> <ToFilename>")
    End Sub

    Public Sub Main()

        Dim CurrArg As String
        Dim CurrLine As String
        Dim FromFileName As String = Nothing
        Dim ToFilename As String = Nothing
        ' ----------------------------------

        ' --- Check Command-line arguments ---
        For i As Integer = 0 To CmdLineArgs.Count - 1
            CurrArg = CmdLineArgs.Arg(i)
            If CurrArg.StartsWith("/") OrElse CurrArg.StartsWith("-") Then
                ShowSyntax()
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(FromFileName) Then
                FromFileName = CurrArg
            ElseIf String.IsNullOrWhiteSpace(ToFilename) Then
                ToFilename = CurrArg
            Else
                ShowSyntax()
                Exit Sub
            End If
        Next

        Dim sw As New StreamWriter(ToFilename, True, Encoding.ASCII)

        Using sr As New StreamReader(FromFileName)
            While Not sr.EndOfStream
                CurrLine = sr.ReadLine
                ' --- Make sure the line starts with a 2-digit record type ---
                If CurrLine.Length < 2 Then Continue While
                If CurrLine.Substring(0, 1) < "0" Then Continue While
                If CurrLine.Substring(0, 1) > "9" Then Continue While
                If CurrLine.Substring(1, 1) < "0" Then Continue While
                If CurrLine.Substring(1, 1) > "9" Then Continue While
                sw.WriteLine(CurrLine)
            End While
            sw.Close()
        End Using

        sw.Close()

    End Sub

End Module
