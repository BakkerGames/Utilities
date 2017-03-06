' --------------------------------------------------
' --- IDRIS_Cadol_Sql_Xref.Part1.vb - 06/16/2014 ---
' --------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 06/16/2014 - SBakker
'            - Added partial class to hold GetByTableName().
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.StringUtils

Partial Public Class IDRIS_Cadol_Sql_Xref

    Public Shared Function GetByTableName(ByVal TableName As String) As IDRIS_Cadol_Sql_Xref
        Return IDRIS_Cadol_Sql_Xref.GetBySQL(BaseQuery + _
                                             FirstConj + " [SQLTableName] = '" + StringToSQL(TableName) + "'")
    End Function

End Class
