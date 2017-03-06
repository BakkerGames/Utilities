' ----------------------------------------------
' --- IDRIS_DataFormat.Part1.vb - 06/16/2014 ---
' ----------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 06/16/2014 - SBakker
'            - Added partial class to hold GetAllByTableName().
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.StringUtils

Partial Public Class IDRIS_DataFormat

    Public Shared Function GetAllByTableName(ByVal TableName As String) As List(Of IDRIS_DataFormat)
        Return IDRIS_DataFormat.GetAllBySQL(BaseQuery + _
                                            FirstConj + " [TableName] = '" + StringToSQL(TableName) + "'" + _
                                            " ORDER BY [FieldNumber]")
    End Function

End Class
