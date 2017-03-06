' --------------------------------
' --- AppUtils.vb - 03/06/2014 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 03/06/2014 - SBakker
'            - Added CommandLineArguments as a function so the arguments to the current program can
'              be handled as a single string. Used by Bootstrap to pass along any parameters.
' ----------------------------------------------------------------------------------------------------

Imports System.Text

Public Class AppUtils

    ''' <summary>
    ''' Return all command line arguments sent to the current program in a single, unaltered string.
    ''' </summary>
    Public Shared Function CommandLineArguments() As String
        Dim MyCommandLine As String = Environment.CommandLine
        ' ---------------------------------------------------
        Try
            If MyCommandLine.StartsWith("""") Then
                Return MyCommandLine.Substring(MyCommandLine.IndexOf("""", 1) + 1).Trim
            Else
                Return MyCommandLine.Substring(MyCommandLine.IndexOf(" ")).Trim
            End If
        Catch
            Return ""
        End Try
    End Function

End Class
