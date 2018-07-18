' --------------------------------------
' --- ParseCadolCode.vb - 07/06/2018 ---
' --------------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/19/2016 - SBakker
'            - Working on ENTERALPHA, ENTERNUM, ESC.
' 07/27/2016 - SBakker
'            - Working on attributes.
' 09/17/2013 - SBakker
'            - Added additional error information.
' ----------------------------------------------------------------------------------------------------

Imports System.Text

Module ParseCadolCode

    Public Tokens() As String
    Public CurrTokenNum As Integer = 0

    Public Sub ParseCode(ByVal CurrLine As String)
        If String.IsNullOrWhiteSpace(CurrLine) Then Exit Sub
        Tokens = CurrLine.Split(CChar(vbTab))
        CurrTokenNum = 0
        If IsNumeric(Tokens(CurrTokenNum)) Then
            CurrTokenNum += 1
            If Tokens.Count = 1 Then Exit Sub ' Only has a line number
        End If
        Try
            ParseCommand()
        Catch ex As Exception
            ExitIDRIS = True
            Throw New SystemException("Fatal error processing line: " + CurrLine + vbCrLf + ex.Message)
        End Try
        Application.DoEvents()
    End Sub

    Public Sub ParseCommand()

        Dim CurrNumericValue As Long = 0
        ' ------------------------------

        Dim CurrCommand As String = Tokens(CurrTokenNum)

        ' ------------------------------------
        ' --- Handle Single-Token Commands ---
        ' ------------------------------------

        If CurrTokenNum = Tokens.Count - 1 Then

            Select Case CurrCommand

                Case "CLEAR"
                    FormMain.ClearScreen()
                    Exit Sub

                Case "ESC"
                    DoEscape()
                    Exit Sub

                Case "EXITRUNTIME"
                    ExitIDRIS = True
                    Exit Sub

                Case "GRAPHOFF"
                    DoGraphOff()
                    Exit Sub

                Case "GRAPHON"
                    DoGraphOn()
                    Exit Sub

                Case "HOME"
                    FormMain.ResetCursorPos()
                    Exit Sub

                Case "KFREE"
                    DoKeyboardUnlock()
                    Exit Sub

                Case "KLOCK"
                    DoKeyboardLock()
                    Exit Sub

                Case "LOCK"
                    DoLock()
                    Exit Sub

                Case "NL"
                    FormMain.DoNewLine()
                    Exit Sub

                Case "NOP"
                    ' --- Do nothing! ---
                    Exit Sub

                Case "PRINTOFF"
                    DoPrintOff()
                    Exit Sub

                Case "PRINTON"
                    DoPrintOn()
                    Exit Sub

                Case "RESETSCREEN"
                    FormMain.ResetScreen()
                    Exit Sub

                Case "RETURN"
                    DoReturn()
                    Exit Sub

                Case "UNLOCK"
                    DoUnlock()
                    Exit Sub

                Case "ZERO"
                    HandleZero()
                    Exit Sub

            End Select

            Throw New SystemException($"Unhandled Command: {CurrCommand}")

        End If

        ' -----------------------------------
        ' --- Handle Multi-Token Commands ---
        ' -----------------------------------

        Select Case CurrCommand

            Case "NL"
                CurrTokenNum += 1
                CurrNumericValue = GetNumericValue(CurrTokenNum)
                For LoopNum As Long = 1 To CurrNumericValue
                    FormMain.DoNewLine()
                Next
                Exit Sub

            Case "GOTO"
                CurrTokenNum += 1
                CodeLineNum = CInt(GetNumericValue(CurrTokenNum))
                Exit Sub

            Case "LOAD"
                CurrTokenNum += 1
                CodeProgNum = CInt(GetNumericValue(CurrTokenNum))
                CodeLineNum = 0
                Exit Sub

            Case "GOS"
                CurrTokenNum += 1
                Dim TempCallStack As New CallStackItem
                With TempCallStack
                    .ProgNum = CodeProgNum
                    .LineNum = CodeLineNum
                End With
                CallStack.Add(TempCallStack)
                CodeLineNum = CInt(GetNumericValue(CurrTokenNum))
                Exit Sub

            Case "GOSUB"
                CurrTokenNum += 1
                Dim TempCallStack As New CallStackItem
                With TempCallStack
                    .ProgNum = CodeProgNum
                    .LineNum = CodeLineNum
                End With
                CallStack.Add(TempCallStack)
                CodeProgNum = CInt(GetNumericValue(CurrTokenNum))
                CodeLineNum = 0
                Exit Sub

            Case "DISPLAY"
                CurrTokenNum += 1
                FormMain.DrawString(GetStringValue(CurrTokenNum))
                Exit Sub

            Case "ATT"
                CurrTokenNum += 1
                With FormMain
                    Dim SaveAttrib As Byte = CByte(GetNumericValue(CurrTokenNum))
                    ' --- Set attribute ---
                    .Screen(.CursorX, .CursorY) = AttributeChar
                    .Attrib(.CursorX, .CursorY) = SaveAttrib
                    ' --- Propogate attribute ---
                    Dim CurrX As Integer = .CursorX + 1
                    Dim CurrY As Integer = .CursorY
                    If CurrX > ScreenWidth Then
                        CurrX = 0
                        CurrY += 1
                    End If
                    Do While CurrY <= ScreenHeight
                        If .Screen(CurrX, CurrY) = AttributeChar Then
                            Exit Do
                        End If
                        .Attrib(CurrX, CurrY) = SaveAttrib
                        CurrX += 1
                        If CurrX > ScreenWidth Then
                            CurrX = 0
                            CurrY += 1
                        End If
                    Loop
                    .DrawChar(AttributeChar, True)
                End With
                Exit Sub

            Case "ENTERALPHA"
                CurrTokenNum += 1
                Dim Result As New StringBuilder
                Dim CurrChar As Char
                Do
                    CurrChar = GetKeyboardChar()
                    If CurrChar >= " "c AndAlso CurrChar <= "~"c Then
                        If FormMain.CursorX < ScreenWidth Then
                            Result.Append(CurrChar)
                            FormMain.DrawChar(CurrChar)
                        End If
                    ElseIf CurrChar = Chr(8) Then ' Backspace
                        If Result.Length > 0 AndAlso FormMain.CursorX > 0 Then
                            Result.Length -= 1
                            FormMain.CursorX -= 1
                            FormMain.DrawChar(" "c)
                        End If
                    End If
                Loop Until CurrChar = vbCr
                Exit Sub

            Case "ENTERNUM"
                CurrTokenNum += 1
                Dim Result As New StringBuilder
                Dim CurrChar As Char
                Do
                    CurrChar = GetKeyboardChar()
                    If (CurrChar >= "0"c AndAlso CurrChar <= "9"c) OrElse (CurrChar = "."c) OrElse (CurrChar = "-"c) Then
                        If FormMain.CursorX < ScreenWidth Then
                            Result.Append(CurrChar)
                            FormMain.DrawChar(CurrChar)
                        End If
                    ElseIf CurrChar = Chr(8) Then ' Backspace
                        If Result.Length > 0 AndAlso FormMain.CursorX > 0 Then
                            Result.Length -= 1
                            FormMain.CursorX -= 1
                            FormMain.DrawChar(" "c)
                        End If
                    End If
                Loop Until CurrChar = vbCr
                Exit Sub

            Case "CLOSE"
                CurrTokenNum += 1
                CurrNumericValue = GetNumericValue(CurrTokenNum)
                'TODO: ### close the file ###
                Exit Sub

        End Select

        ' ------------------
        ' --- Assignment ---
        ' ------------------

        If CurrCommand = "LET" Then ' Skip LET
            CurrTokenNum += 1
        End If

        If IsNumericTarget(CurrTokenNum) Then
            'TODO: ### Handle numeric assignment ###
            Exit Sub
        End If

        ' ----------
        ' --- If ---
        ' ----------

        If CurrCommand = "IF" Then
            'TODO: ### Handle IF ###
            Exit Sub
        End If

        If CurrCommand = "THEN" Then
            'TODO: ### Handle THEN ###
            Exit Sub
        End If

        If CurrCommand = "ELSE" Then
            'TODO: ### Handle ELSE ###
            Exit Sub
        End If

        If CurrCommand = "ENDIF" Then
            'TODO: ### Handle ENDIF ###
            Exit Sub
        End If

        ' -------------
        ' --- Error ---
        ' -------------

        Throw New SystemException("Unhandled Command: " + CurrCommand)

    End Sub

    Private Function IsNumericTarget(ByRef currTokenNum As Integer) As Boolean
        If Tokens(currTokenNum).StartsWith("N") OrElse
                Tokens(currTokenNum).StartsWith("F") OrElse
                Tokens(currTokenNum).StartsWith("G") Then
            If Tokens(currTokenNum).Length = 1 Then Return True
            If IsNumeric(Tokens(currTokenNum).Substring(1)) Then
                Dim OffsetNum As Integer = CInt(Tokens(currTokenNum).Substring(1))
                If OffsetNum < 1 Then Return False
                If OffsetNum > 99 Then Return False
                If OffsetNum > 9 AndAlso Tokens(currTokenNum).StartsWith("G") Then Return False
            End If
        End If
        Select Case Tokens(currTokenNum)
            Case "RP" : Return True
            Case "RP2" : Return True
            Case "IRP" : Return True
            Case "IRP2" : Return True
            Case "ZP" : Return True
            Case "ZP2" : Return True
            Case "IZP" : Return True
            Case "IZP2" : Return True
            Case "XP" : Return True
            Case "XP2" : Return True
            Case "IXP" : Return True
            Case "IXP2" : Return True
            Case "YP" : Return True
            Case "YP2" : Return True
            Case "IYP" : Return True
            Case "IYP2" : Return True
            Case "WP" : Return True
            Case "WP2" : Return True
            Case "IWP" : Return True
            Case "IWP2" : Return True
            Case "SP" : Return True
            Case "SP2" : Return True
            Case "ISP" : Return True
            Case "ISP2" : Return True
            Case "TP" : Return True
            Case "TP2" : Return True
            Case "ITP" : Return True
            Case "ITP2" : Return True
            Case "UP" : Return True
            Case "UP2" : Return True
            Case "IUP" : Return True
            Case "IUP2" : Return True
            Case "VP" : Return True
            Case "VP2" : Return True
            Case "IVP" : Return True
            Case "IVP2" : Return True
            Case "LIB" : Return True
            Case "PROG" : Return True
            Case "PRIVG" : Return True
            Case "CHAR" : Return True
            Case "LENGTH" : Return True
            Case "STATUS" : Return True
            Case "ESCVAL" : Return True
            Case "CANVAL" : Return True
            Case "LOCKVAL" : Return True
            Case "TCHAN" : Return True
            Case "TERM" : Return True
            Case "LANG" : Return True
            Case "PRTNUM" : Return True
            Case "TFA" : Return True
            Case "VOL" : Return True
            Case "PVOL" : Return True
            Case "REQVOL" : Return True
            Case "USER" : Return True
            Case "ORIG" : Return True
            Case "OPER" : Return True
        End Select
        If Tokens(currTokenNum) = "N" AndAlso Tokens(currTokenNum + 1) = "[" Then
            Dim saveTokenNum As Integer = currTokenNum
            saveTokenNum += 2
            GetNumericValue(saveTokenNum)
            If Tokens(saveTokenNum) = "]" Then
                saveTokenNum += 1
                If Tokens(saveTokenNum) = "=" Then Return True
            End If
        End If
        'TODO: ### buffers? ###
        Return False
    End Function

    Private Function GetNumericValue(ByRef CurrTokenNum As Integer) As Long

        Dim CurrValue As Long = 0
        Dim TempValue As Long = 0
        Dim CurrOperator As String = ""
        Dim UnaryMinus As Boolean = False
        ' -------------------------------

        ' --- Check if missing tokens needed in the expression ---
        If CurrTokenNum >= Tokens.Count Then
            Throw New SystemException("Numeric expression underflow")
        End If

        ' --- Check for negative sign in front of expression ---
        If Tokens(CurrTokenNum) = "-" Then
            UnaryMinus = True
            CurrTokenNum += 1
        End If

        ' --- Check for beginning of expression ---
        If IsNumeric(Tokens(CurrTokenNum)) Then
            CurrValue = CLng(Tokens(CurrTokenNum))
            CurrTokenNum += 1
        ElseIf Tokens(CurrTokenNum) = "(" Then ' Handle nesting parenthesis
            CurrTokenNum += 1
            CurrValue = GetNumericValue(CurrTokenNum)
        Else
            CurrValue = GetNumericMemoryValue(CurrTokenNum)
        End If

        ' --- Negate value as necessary ---
        If UnaryMinus Then
            CurrValue = -CurrValue
            UnaryMinus = False
        End If

        ' --- Check for end of line ---
        If CurrTokenNum = Tokens.Count Then
            Return CurrValue
        End If

        ' --- Get operator between expressions ---
        CurrOperator = Tokens(CurrTokenNum)
        CurrTokenNum += 1

        Select Case CurrOperator
            Case ")" ' --- Ending parenthesis ---
                Return CurrValue
            Case "]" ' --- Ending bracket ---
                Return CurrValue
            Case "+"
                CurrValue = CurrValue + GetNumericValue(CurrTokenNum)
            Case "-"
                CurrValue = CurrValue - GetNumericValue(CurrTokenNum)
            Case "*"
                CurrValue = CurrValue * GetNumericValue(CurrTokenNum)
            Case "/"
                TempValue = GetNumericValue(CurrTokenNum)
                REMVAL = CurrValue - (TempValue * (CurrValue \ TempValue))
                CurrValue = CurrValue \ TempValue
            Case Else
                Throw New SystemException("Numeric operator not recognized: " + CurrOperator)
        End Select

        Return CurrValue

    End Function

    Private Function GetNumericMemoryValue(ByRef currTokenNum As Integer) As Long
        Select Case Tokens(currTokenNum)
            Case "RP" : Return GetNumeric(MemPos_RP, 1)
            Case "RP2" : Return GetNumeric(MemPos_RP2, 1)
            Case "IRP" : Return GetNumeric(MemPos_IRP, 1)
            Case "IRP2" : Return GetNumeric(MemPos_IRP2, 1)
            Case "ZP" : Return GetNumeric(MemPos_ZP, 1)
            Case "ZP2" : Return GetNumeric(MemPos_ZP2, 1)
            Case "IZP" : Return GetNumeric(MemPos_IZP, 1)
            Case "IZP2" : Return GetNumeric(MemPos_IZP2, 1)
            Case "XP" : Return GetNumeric(MemPos_XP, 1)
            Case "XP2" : Return GetNumeric(MemPos_XP2, 1)
            Case "IXP" : Return GetNumeric(MemPos_IXP, 1)
            Case "IXP2" : Return GetNumeric(MemPos_IXP2, 1)
            Case "YP" : Return GetNumeric(MemPos_YP, 1)
            Case "YP2" : Return GetNumeric(MemPos_YP2, 1)
            Case "IYP" : Return GetNumeric(MemPos_IYP, 1)
            Case "IYP2" : Return GetNumeric(MemPos_IYP2, 1)
            Case "WP" : Return GetNumeric(MemPos_WP, 1)
            Case "WP2" : Return GetNumeric(MemPos_WP2, 1)
            Case "IWP" : Return GetNumeric(MemPos_IWP, 1)
            Case "IWP2" : Return GetNumeric(MemPos_IWP2, 1)
            Case "SP" : Return GetNumeric(MemPos_SP, 1)
            Case "SP2" : Return GetNumeric(MemPos_SP2, 1)
            Case "ISP" : Return GetNumeric(MemPos_ISP, 1)
            Case "ISP2" : Return GetNumeric(MemPos_ISP2, 1)
            Case "TP" : Return GetNumeric(MemPos_TP, 1)
            Case "TP2" : Return GetNumeric(MemPos_TP2, 1)
            Case "ITP" : Return GetNumeric(MemPos_ITP, 1)
            Case "ITP2" : Return GetNumeric(MemPos_ITP2, 1)
            Case "UP" : Return GetNumeric(MemPos_UP, 1)
            Case "UP2" : Return GetNumeric(MemPos_UP2, 1)
            Case "IUP" : Return GetNumeric(MemPos_IUP, 1)
            Case "IUP2" : Return GetNumeric(MemPos_IUP2, 1)
            Case "VP" : Return GetNumeric(MemPos_VP, 1)
            Case "VP2" : Return GetNumeric(MemPos_VP2, 1)
            Case "IVP" : Return GetNumeric(MemPos_IVP, 1)
            Case "IVP2" : Return GetNumeric(MemPos_IVP2, 1)
            Case "LIB" : Return GetNumeric(MemPos_Lib, 1)
            Case "PROG" : Return GetNumeric(MemPos_Prog, 1)
            Case "PRIVG" : Return GetNumeric(MemPos_Privg, 1)
            Case "CHAR" : Return GetNumeric(MemPos_Char, 1)
            Case "LENGTH" : Return GetNumeric(MemPos_Length, 1)
            Case "STATUS" : Return GetNumeric(MemPos_Status, 1)
            Case "ESCVAL" : Return GetNumeric(MemPos_EscVal, 1)
            Case "CANVAL" : Return GetNumeric(MemPos_CanVal, 1)
            Case "LOCKVAL" : Return GetNumeric(MemPos_LockVal, 1)
            Case "TCHAN" : Return GetNumeric(MemPos_TChan, 1)
            Case "TERM" : Return GetNumeric(MemPos_Term, 1)
            Case "LANG" : Return GetNumeric(MemPos_Lang, 1)
            Case "PRTNUM" : Return GetNumeric(MemPos_PrtNum, 1)
            Case "TFA" : Return GetNumeric(MemPos_TFA, 1)
            Case "VOL" : Return GetNumeric(MemPos_Vol, 1)
            Case "PVOL" : Return GetNumeric(MemPos_PVol, 1)
            Case "REQVOL" : Return GetNumeric(MemPos_ReqVol, 1)
            Case "USER" : Return GetNumeric(MemPos_User, 2)
            Case "ORIG" : Return GetNumeric(MemPos_Orig, 2)
            Case "OPER" : Return GetNumeric(MemPos_Oper, 2)
        End Select
        Dim OffsetNum As Integer = 0
        If Tokens(currTokenNum).StartsWith("N") OrElse
           Tokens(currTokenNum).StartsWith("F") OrElse
           Tokens(currTokenNum).StartsWith("G") Then
            If IsNumeric(Tokens(currTokenNum).Substring(1)) Then
                OffsetNum = CInt(Tokens(currTokenNum).Substring(1))
                If OffsetNum < 0 OrElse OffsetNum > 99 Then
                    Throw New SystemException($"Invalid N/F/G Offset {OffsetNum}")
                End If
                If OffsetNum > 9 AndAlso Tokens(currTokenNum).StartsWith("G") Then
                    Throw New SystemException($"Invalid G Offset {OffsetNum}")
                End If
            End If
        End If
        If Tokens(currTokenNum).StartsWith("N") Then
            currTokenNum += 1
            Return N(OffsetNum)
        End If
        If Tokens(currTokenNum).StartsWith("F") Then
            currTokenNum += 1
            Return F(OffsetNum)
        End If
        If Tokens(currTokenNum).StartsWith("G") Then
            currTokenNum += 1
            'TODO: ### This needs to be more "global" !!! ###
            Return GetNumeric(MemPos_G + OffsetNum, 1)
        End If
        Throw New NotImplementedException()
    End Function

    Private Function GetAlphaMemoryValue(currTokenNum As Integer) As String
        Throw New NotImplementedException()
    End Function

    Private Function GetStringValue(ByRef CurrTokenNum As Integer) As String

        Dim CurrValue As String = ""
        ' --------------------------

        ' --- Check if missing tokens needed in the expression ---
        If CurrTokenNum >= Tokens.Count Then
            Throw New SystemException("String expression underflow")
        End If

        CurrValue = Tokens(CurrTokenNum)

        If CurrValue.StartsWith("""") AndAlso CurrValue.EndsWith("""") Then
            CurrValue = CurrValue.Substring(1, CurrValue.Length - 2)
            CurrTokenNum += 1
        ElseIf CurrValue.StartsWith("'") AndAlso CurrValue.EndsWith("'") Then
            CurrValue = CurrValue.Substring(1, CurrValue.Length - 2)
            CurrTokenNum += 1
        ElseIf CurrValue.StartsWith("%") AndAlso CurrValue.EndsWith("%") Then
            CurrValue = CurrValue.Substring(1, CurrValue.Length - 2)
            CurrTokenNum += 1
        ElseIf CurrValue.StartsWith("$") AndAlso CurrValue.EndsWith("$") Then
            CurrValue = CurrValue.Substring(1, CurrValue.Length - 2)
            CurrTokenNum += 1
        Else
            Throw New SystemException("String expression not recognized: " + CurrValue)
        End If

        Return CurrValue

    End Function

End Module
