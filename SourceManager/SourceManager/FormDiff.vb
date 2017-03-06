' --------------------------------
' --- FormDiff.vb - 08/23/2011 ---
' --------------------------------

' ------------------------------------------------------------------------------------------
' 08/23/2011 - SBakker
'            - Handled ESC to close the form, in two way just in case.
'            - Save if the form is maximized or not, and properly reload AFTER it has been
'              shown, not during creation.
' ------------------------------------------------------------------------------------------

Public Class FormDiff

    Private FormLoaded As Boolean = False

    Private Sub FormDiff_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If My.Settings.FormDiffMaximized Then
            Me.WindowState = FormWindowState.Maximized
        End If
        For LoopNum As Integer = 0 To ComboBoxFont.Items.Count - 1
            If CStr(ComboBoxFont.Items(LoopNum)) = My.Settings.DiffFont Then
                ComboBoxFont.SelectedIndex = LoopNum
                Exit For
            End If
        Next
        Me.Focus()
        FormLoaded = True
    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    Private Sub ComboBoxFont_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxFont.SelectedIndexChanged
        Select Case CStr(ComboBoxFont.Items(ComboBoxFont.SelectedIndex))
            Case "Proportional"
                TextBoxDiff.Font = New Font("Microsoft Sans Serif", 8)
                My.Settings.DiffFont = CStr(ComboBoxFont.Items(ComboBoxFont.SelectedIndex))
                My.Settings.Save()
            Case "Fixed"
                TextBoxDiff.Font = New Font("Courier New", 8)
                My.Settings.DiffFont = CStr(ComboBoxFont.Items(ComboBoxFont.SelectedIndex))
                My.Settings.Save()
        End Select
    End Sub

    Private Sub FormDiff_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If Not FormLoaded Then Exit Sub
        My.Settings.FormDiffMaximized = (Me.WindowState = FormWindowState.Maximized)
        My.Settings.Save()
    End Sub

    Private Sub TextBoxDiff_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBoxDiff.KeyDown
        If e.KeyCode = Keys.Escape Then
            ButtonCancel.PerformClick()
        End If
    End Sub

End Class
