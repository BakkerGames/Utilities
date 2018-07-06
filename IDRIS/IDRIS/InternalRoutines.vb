' ----------------------------------------
' --- InternalRoutines.vb - 10/21/2016 ---
' ----------------------------------------

' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Module InternalRoutines

    Friend Sub DoEscape()
        'TODO: ### fill in code here to use WHEN and/or clear stack ###
        CodeProgNum = 0
        CodeLineNum = 0
        ExitIDRIS = True
    End Sub

    Friend Sub DoCancel()
        Throw New NotImplementedException
    End Sub

    Friend Sub DoLock()
        ' --- Next attempt to read will lock a record ---
        LockFlag = True
    End Sub

    Friend Sub DoUnlock()
        ' --- Check if record is currently locked ---
        If HasLockedRec Then
            'TODO: ### On Error Resume Next
            'TODO: ### If rsLockedRec.State = adStateOpen Then
            'TODO: ### rsLockedRec.CancelUpdate()
            'TODO: ### rsLockedRec.Close()
            'TODO: ### End If
            'TODO: ### On Error GoTo 0
            'TODO: ### ReleaseAppLock LockedResource
            LockedResource = ""
            LockedFileNum = -1
            LockedRecNum = -1
            LockedRecLen = -1
            'TODO: ### LockedCadolXref = Nothing
            HasLockedRec = False
        End If
        ' --- Clear the lock flag ---
        LockFlag = False
    End Sub

    Friend Sub DoKeyboardLock()
        KeyboardLocked = True
    End Sub

    Friend Sub DoKeyboardUnlock()
        KeyboardLocked = False
    End Sub

    Friend Sub DoGraphOff()
        GraphicsCharFlag = False
    End Sub

    Friend Sub DoGraphOn()
        GraphicsCharFlag = True
    End Sub

    Friend Sub DoReturn()
        If CallStack.Count = 0 Then
            Throw New SystemException("Gosub Stack Underflow!")
        End If
        Dim TempCallStack As CallStackItem = CallStack.Last
        CallStack.RemoveAt(CallStack.Count - 1)
        With TempCallStack
            CodeProgNum = .ProgNum
            CodeLineNum = .LineNum
        End With
    End Sub

    Friend Sub DoPrintOff()
        Throw New NotImplementedException
    End Sub

    Friend Sub DoPrintOn()
        Throw New NotImplementedException
    End Sub

End Module
