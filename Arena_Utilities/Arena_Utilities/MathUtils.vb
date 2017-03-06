' ---------------------------------
' --- MathUtils.vb - 07/07/2014 ---
' ---------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/07/2015 - SBakker
'            - Added Round6() function. Found place in LTDClaimCalcs that uses it.
' 05/08/2014 - SBakker
'            - Added RoundX() for three, four and five decimal places. These are all possible decimal
'              field sizes in Arena.
' 03/18/2014 - SBakker
'            - Added functions Round0 and Round2 to handle rounding consistently and simply.
' ----------------------------------------------------------------------------------------------------

Public Class MathUtils

    ''' <summary>
    ''' Rounds a value to 0 decimal places and returns an integer
    ''' </summary>
    Public Shared Function Round0(ByVal Value As Decimal) As Integer
        Return CInt(Math.Round(Value, 0, MidpointRounding.AwayFromZero))
    End Function

    ''' <summary>
    ''' Rounds a value to 2 decimal places
    ''' </summary>
    Public Shared Function Round2(ByVal Value As Decimal) As Decimal
        Return Math.Round(Value, 2, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Rounds a value to 3 decimal places
    ''' </summary>
    Public Shared Function Round3(ByVal Value As Decimal) As Decimal
        Return Math.Round(Value, 3, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Rounds a value to 4 decimal places
    ''' </summary>
    Public Shared Function Round4(ByVal Value As Decimal) As Decimal
        Return Math.Round(Value, 4, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Rounds a value to 5 decimal places
    ''' </summary>
    Public Shared Function Round5(ByVal Value As Decimal) As Decimal
        Return Math.Round(Value, 5, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Rounds a value to 6 decimal places
    ''' </summary>
    Public Shared Function Round6(ByVal Value As Decimal) As Decimal
        Return Math.Round(Value, 6, MidpointRounding.AwayFromZero)
    End Function

End Class
