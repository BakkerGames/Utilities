' ---------------------------------------
' --- FormTargetNewer.vb - 05/19/2011 ---
' ---------------------------------------

' ------------------------------------------------------------------------------------------
' 05/19/2011 - SBakker
'            - Moved Enum OverwriteResult into FormMain.vb.
'            - Put message into header of form, and only display the filename in the form.
' 04/07/2011 - SBakker
'            - Arranged the buttons in the Windows order for the same screen.
'            - Make "No" the default and set focus on it when loading the form.
' 03/18/2011 - SBakker
'            - Added FormTargetNewer to handle responses like NoToAll and YesToAll.
' ------------------------------------------------------------------------------------------

Public Class FormTargetNewer

    Public Property Result As OverwriteResult = OverwriteResult.Unknown

    Private Sub FormTargetNewer_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ButtonNo.Focus()
    End Sub

    Private Sub ButtonNo_Click(sender As System.Object, e As System.EventArgs) Handles ButtonNo.Click
        Result = OverwriteResult.No
        Me.Close()
    End Sub

    Private Sub ButtonNoToAll_Click(sender As System.Object, e As System.EventArgs) Handles ButtonNoToAll.Click
        Result = OverwriteResult.NoToAll
        Me.Close()
    End Sub

    Private Sub ButtonYes_Click(sender As System.Object, e As System.EventArgs) Handles ButtonYes.Click
        Result = OverwriteResult.Yes
        Me.Close()
    End Sub

    Private Sub ButtonYesToAll_Click(sender As System.Object, e As System.EventArgs) Handles ButtonYesToAll.Click
        Result = OverwriteResult.YesToAll
        Me.Close()
    End Sub

    Private Sub ButtonAbort_Click(sender As System.Object, e As System.EventArgs) Handles ButtonAbort.Click
        Result = OverwriteResult.Abort
        Me.Close()
    End Sub

End Class