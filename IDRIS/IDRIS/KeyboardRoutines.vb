' ----------------------------------------
' --- KeyboardRoutines.vb - 01/25/2016 ---
' ----------------------------------------

Module KeyboardRoutines

    Public Sub AddKeyboardChar(ByVal CurrChar As Char)
        KeyboardQueue.Enqueue(CurrChar)
    End Sub

    Public Function GetKeyboardChar() As Char
        Do While KeyboardQueue.Count = 0
            Application.DoEvents()
        Loop
        Return KeyboardQueue.Dequeue
    End Function

    Public Function PeekKeyboardChar() As Char
        If KeyboardQueue.Count > 0 Then
            Return KeyboardQueue.Peek
        End If
        Return Nothing ' Check if this is OK
    End Function

    Public Function KeyboardStackCount() As Integer
        Return KeyboardQueue.Count
    End Function

    Public Sub ClearKeyboardStack()
        KeyboardQueue.Clear()
    End Sub

End Module
