' --------------------------------
' --- Commands.vb - 10/26/2016 ---
' --------------------------------

Module Commands

    ' ------------------
    ' --- Assignment ---
    ' ------------------

    Public Sub HandleZero()
        For Offset As Integer = 0 To 20
            N(Offset) = 0
        Next
    End Sub

    ' ---------------------
    ' --- Device Output ---
    ' ---------------------

    Public Sub DoDisplay()
        If Tokens(CurrTokenNum).StartsWith("""") OrElse
            Tokens(CurrTokenNum).StartsWith("'") OrElse
            Tokens(CurrTokenNum).StartsWith("%") OrElse
            Tokens(CurrTokenNum).StartsWith("$") Then
            If Tokens(CurrTokenNum).EndsWith(Tokens(CurrTokenNum).Substring(0, 1)) Then
                DoDisplayString(Tokens(CurrTokenNum).Substring(1, Tokens(CurrTokenNum).Length - 2))
                Exit Sub
            End If
            Throw New SystemException("Unterminated display string! " + Tokens(CurrTokenNum))
        End If
        'TODO: ### handle rest of display types ###
    End Sub

    Public Sub DoDisplayString(ByVal Value As String)
        ' --- this handles literal strings with no other processing ---
        If Value <> "" Then
            ' --- handle display ---
            If MEMTF(MemPos_PrintOn) Then
                'TODO: --- print to printer ---
                'CheckFormFeed()
                'Print #PrinterFileNum, Value;
                'LET_MEMTF(MemPos_PageHasData, True)
                'LET_MEMTF(MemPos_LineHasData, True)
            Else
                ' --- display to screen ---
                FormMain.DrawString(Value)
            End If
        End If
        ' --- save number of chars displayed ---
        MEM(MemPos_Char) = CByte(Value.Length)
    End Sub

    Public Function DoEnterNumeric(ByVal EnterFormat As String, ByVal Target As String) As Boolean
        'TODO: ### add logic ###
        Return False
    End Function

    Public Function DoEnterAlpha(ByVal EnterFormat As String, ByVal Target As String) As Boolean
        'TODO: ### add logic ###
        Return False
    End Function

End Module
