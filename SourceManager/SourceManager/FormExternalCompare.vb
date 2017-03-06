' -------------------------------------------
' --- FormExternalCompare.vb - 01/12/2012 ---
' -------------------------------------------

' ------------------------------------------------------------------------------------------
' 01/12/2012 - SBakker
'            - Added FormExternalCompare and the new settings UseExternalCompare and
'              ExternalCompareApp. This lets an outside program be used to compare the files
'              instead of the built-in file compare.
' ------------------------------------------------------------------------------------------

Public Class FormExternalCompare

    Private Sub FormExternalCompare_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CheckBoxUseExternalCompare.Checked = My.Settings.UseExternalCompare
        TextBoxExternalCompareApp.Text = My.Settings.ExternalCompareApp
    End Sub

    Private Sub ButtonSave_Click(sender As System.Object, e As System.EventArgs) Handles ButtonSave.Click
        My.Settings.UseExternalCompare = CheckBoxUseExternalCompare.Checked
        My.Settings.ExternalCompareApp = TextBoxExternalCompareApp.Text
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub ButtonCancel_Click(sender As System.Object, e As System.EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

End Class