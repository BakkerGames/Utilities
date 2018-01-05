' ---------------------------------
' --- IDRISFile.vb - 01/05/2018 ---
' ---------------------------------

' ------------------------------------------------------------------------------------------
' 01/05/2018 - SBakker
'            - Added SelectionStart.
' 02/03/2011 - SBakker
'            - Started working on IDRIS_IDE program in VB.NET.
' ------------------------------------------------------------------------------------------

Public Class IDRISFile

    Public FullPath As String = ""
    Public FileName As String = ""
    Public Changed As Boolean = False

    Public FileText As String = ""
    Public SelectionStart As Integer = 0

End Class
