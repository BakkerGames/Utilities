' ------------------------------------------
' --- FormTargetReadonly.vb - 05/19/2011 ---
' ------------------------------------------

' ------------------------------------------------------------------------------------------
' 05/19/2011 - SBakker
'            - Added FormTargetReadonly to handle responses like IgnoreAll.
' ------------------------------------------------------------------------------------------

Public Class FormTargetReadonly

    Public Property Result As OverwriteResult = OverwriteResult.Unknown

    Private Sub FormTargetReadonly_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ButtonRetry.Focus()
    End Sub

    Private Sub ButtonAbort_Click(sender As System.Object, e As System.EventArgs) Handles ButtonAbort.Click
        Result = OverwriteResult.Abort
        Me.Close()
    End Sub

    Private Sub ButtonRetry_Click(sender As System.Object, e As System.EventArgs) Handles ButtonRetry.Click
        Result = OverwriteResult.Retry
        Me.Close()
    End Sub

    Private Sub ButtonIgnore_Click(sender As System.Object, e As System.EventArgs) Handles ButtonIgnore.Click
        Result = OverwriteResult.Ignore
        Me.Close()
    End Sub

    Private Sub ButtonIgnoreAll_Click(sender As System.Object, e As System.EventArgs) Handles ButtonIgnoreAll.Click
        Result = OverwriteResult.IgnoreAll
        Me.Close()
    End Sub

End Class