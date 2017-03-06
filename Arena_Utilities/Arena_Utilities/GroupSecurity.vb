' -------------------------------------
' --- GroupSecurity.vb - 07/22/2016 ---
' -------------------------------------

' ------------------------------------------------------------------------------------------
' 07/22/2016 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrEmpty().
' 09/17/2013 - SBakker
'            - Added additional error information.
' 02/25/2013 - SBakker - URD 11941
'            - Made common location to store GroupClientList information for use by all
'              data and search classes to verify that the data read is permitted.
'            - Added InGroupClientList() for checking single Clients within programs.
' ------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------
' --- Store Group Client List information, for use in every data class and search class. ---
' --- Started by storing a string value in the GroupClientList property.                 ---
' ------------------------------------------------------------------------------------------

Public Class GroupSecurity

#Region " Constants "

    Private Const GroupClientListNotFilledError As String = "GroupClientList not yet filled"

#End Region

    ' --- Used to determine if the list has been filled yet (even if blank) ---
    ''' <summary>Is set to True by assigning GroupClientList.</summary>
    Public Shared IsGroupClientListFilled As Boolean = False

    ' --- Used to check if the Group Client List has any items, rather than checking the string itself ---
    Private Shared _HasGroupClientList As Boolean = False
    ''' <summary>Set to True or False by assigning GroupClientList</summary>
    Public Shared ReadOnly Property HasGroupClientList As Boolean
        Get
            ' --- Put error checking here so it doesn't have to be duplicated in outside routines ---
            Static FuncName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            If Not IsGroupClientListFilled Then
                Throw New SystemException(FuncName + vbCrLf + GroupClientListNotFilledError)
            End If
            Return _HasGroupClientList
        End Get
    End Property

    Private Shared _GroupClientList As String = ""
    ''' <summary>
    ''' GroupClientList must have the format "0001,0002,0003".
    ''' </summary>
    Public Shared Property GroupClientList As String
        Get
            ' --- Put error checking here so it doesn't have to be duplicated in outside routines ---
            Static FuncName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            If Not IsGroupClientListFilled Then
                Throw New SystemException(FuncName + vbCrLf + GroupClientListNotFilledError)
            End If
            Return _GroupClientList
        End Get
        Set(value As String)
            Static FuncName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            Try
                If String.IsNullOrEmpty(value) Then
                    _GroupClientList = ""
                    _GroupClientListSQL = ""
                    _GroupClientListSplit = Nothing
                    _HasGroupClientList = False
                Else
                    _GroupClientList = value
                    _GroupClientListSQL = "('" + value.Replace(",", "','") + "')"
                    ' --- Convert String array to List(Of String) ---
                    _GroupClientListSplit = New List(Of String)
                    Dim TempList() As String = _GroupClientList.Split(","c)
                    For Each TempClient As String In TempList
                        If TempClient.Length <> 4 OrElse TempClient.Trim.Length <> 4 Then
                            Throw New SystemException(FuncName + vbCrLf + "Invalid Client Number: """ + TempClient + """")
                        End If
                        _GroupClientListSplit.Add(TempClient)
                    Next
                    _HasGroupClientList = True
                End If
                IsGroupClientListFilled = True
            Catch ex As Exception
                Throw New SystemException(FuncName + vbCrLf + "Error filling GroupClientList" + vbCrLf + ex.Message)
            End Try
        End Set
    End Property

    Private Shared _GroupClientListSQL As String = ""
    ''' <summary>
    ''' GroupClientListSQL has the format "('0001','0002','0003')". Must be used with "IN".
    ''' </summary>
    Public Shared ReadOnly Property GroupClientListSQL As String
        Get
            ' --- Put error checking here so it doesn't have to be duplicated in outside routines ---
            Static FuncName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            If Not IsGroupClientListFilled Then
                Throw New SystemException(FuncName + vbCrLf + GroupClientListNotFilledError)
            End If
            Return _GroupClientListSQL
        End Get
    End Property

    Private Shared _GroupClientListSplit As List(Of String) = Nothing
    ''' <summary>
    ''' Split list of Group Clients, to use in "For Each".
    ''' </summary>
    Public Shared ReadOnly Property GroupClientListSplit As List(Of String)
        Get
            ' --- Put error checking here so it doesn't have to be duplicated in outside routines ---
            Static FuncName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            If Not IsGroupClientListFilled Then
                Throw New SystemException(FuncName + vbCrLf + GroupClientListNotFilledError)
            End If
            Return _GroupClientListSplit
        End Get
    End Property

    ''' <summary>
    ''' Validates a single Client against the Group Client List.
    ''' </summary>
    Public Shared Function InGroupClientList(ByVal ClientNumber As String) As Boolean
        If Not HasGroupClientList Then Return True
        For Each CurrClient As String In _GroupClientListSplit
            If ClientNumber = CurrClient Then Return True
        Next
        Return False
    End Function

End Class
