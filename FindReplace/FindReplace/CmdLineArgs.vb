' -----------------------------------
' --- CmdLineArgs.vb - 11/18/2010 ---
' -----------------------------------

' ------------------------------------------------------------------------------------------
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 04/17/2009 - SBakker
'            - Building a CmdLineArgs class. The built-in ones don't handle
'              arguments the way I need.
' ------------------------------------------------------------------------------------------

Public Class CmdLineArgs

    ' --- The first argument is going to be the name of the application.
    ' --- This will become Arg(-1), which is basically hidden unless needed.

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private Shared _DidSplit As Boolean = False
    Private Shared _Args As New List(Of String)

    Public Shared ReadOnly Property Count() As Integer
        Get
            If Not _DidSplit Then
                DoSplit()
                _DidSplit = True
            End If
            Return _Args.Count - 1
        End Get
    End Property

    Public Shared ReadOnly Property Arg(ByVal ArgNum As Integer) As String
        Get
            Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
            If Not _DidSplit Then
                DoSplit()
                _DidSplit = True
            End If
            If ArgNum < -1 OrElse ArgNum >= _Args.Count - 1 Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Argument number out of range - " + ArgNum.ToString)
            End If
            Return _Args(ArgNum + 1)
        End Get
    End Property

    Private Shared Sub DoSplit()
        Dim InArg As Boolean = False
        Dim InQuote As Boolean = False
        Dim CurrArg As String = ""
        Dim CurrChar As Char = "~"c
        Dim LastChar As Char = "~"c
        ' ----------------------------
        For i As Integer = 0 To Environment.CommandLine.Length - 1
            LastChar = CurrChar
            CurrChar = Environment.CommandLine(i)
            If Not InArg Then
                If CurrChar = """"c Then
                    InArg = True
                    InQuote = True
                    CurrChar = "~"c
                ElseIf CurrChar <> " "c Then
                    InArg = True
                    CurrArg = CurrChar
                End If
            Else
                If CurrChar = """"c And LastChar = """"c Then
                    CurrArg += CurrChar
                    CurrChar = "~"c
                ElseIf CurrChar = """"c Then
                    ' --- check for double-quote ---
                    If i >= Environment.CommandLine.Length - 1 OrElse Environment.CommandLine(i + 1) <> """"c Then
                        _Args.Add(CurrArg)
                        InArg = False
                        InQuote = False
                        CurrArg = ""
                        CurrChar = "~"c
                    End If
                ElseIf CurrChar = " "c AndAlso Not InQuote Then
                    If CurrArg <> "" Then
                        _Args.Add(CurrArg)
                    End If
                    InArg = False
                    CurrArg = ""
                Else
                    CurrArg += CurrChar
                End If
            End If
        Next
        If InArg And CurrArg <> "" Then
            _Args.Add(CurrArg)
        End If
        _DidSplit = True
    End Sub

End Class
