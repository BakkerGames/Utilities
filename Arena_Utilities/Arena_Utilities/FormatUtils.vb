' -----------------------------------
' --- FormatUtils.vb - 07/22/2016 ---
' -----------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/22/2016 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrEmpty().
' 08/21/2015 - SBakker
'            - Added OLS Claim Number routines.
' 09/30/2014 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrWhiteSpace().
' 09/09/2014 - SBakker
'            - Added UnformatNumberInt(), for returning Integers from string fields. Throws error if
'              convert fails.
' 08/07/2014 - SBakker
'            - Added UnformatNumber() and UnformatNumberNullable() to change a string to a decimal.
'              The first returns 0 for "", the second returns Nothing for "". Both will throw errors
'              if the convert fails.
' 05/08/2014 - SBakker
'            - Added FormatNumberX() and FormatCurrencyX() for zero, three, four, and five decimal
'              places. These are all possible decimal field sizes in Arena.
' 03/31/2014 - SBakker
'            - Added UnformatPhoneToNumber(), because for some reason CLSTDxx L14 has the VALUE field
'              filled with the numeric version of the Phone Number. Return as a decimal value - too
'              large for an integer.
' 03/19/2014 - SBakker
'            - Added function PlusIfPositive(), for adding "+" in front of some CLSTDxx amount fields.
' 03/12/2014 - SBakker
'            - Added UnformatSSNToNumber(), because for some reason CLSTDxx L11 has the VALUE field
'              filled with the numeric version of the SSN.
' 03/11/2014 - SBakker
'            - Added an Integer form of ValidateBankRTCode, and allow "0" to be valid in the String
'              form.
' 02/13/2014 - SBakker
'            - Added FormatNumber2() function to consistently format all 2-decimal place values.
'            - Added comments and cleaned up FromYN/FromYNBlank
' 09/17/2013 - SBakker
'            - Added additional error information.
' 07/15/2013 - SBakker
'            - Added function ToYNNull to handle saving values to a nullable string field.
'              FromYNBlank already handles filling Nullable(Of Boolean) fields.
' 01/31/2013 - SBakker
'            - Added functions ToYNBlank and FromYNBlank.
' 06/22/2012 - SBakker
'            - Allow LTD Claim Numbers to have leading zeros added for each section of the
'              claim number, so "101 95 1" is valid.
' 11/23/2011 - SBakker
'            - Cleaned up First Two Digits logic in ValidateBankRTCode to read exactly like
'              the rule, so it's easier to understand.
' 03/31/2011 - SBakker
'            - Added ValidateBankRTCode to be able to check a Bank Routine/Transit number.
' 02/25/2011 - SBakker
'            - Added Boolean Function Routines "ToYN" and "FromYN", to convert between the
'              boolean True/False and fields than hold "Y" or "N".
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 04/05/2010 - SBakker
'            - Removed adding of leading zeros to Social Security Number and Zip
'              Codes, at the request of Claims. Also removed adding of leading
'              zeros from TaxID, for the same reason, even though not requested.
'            - Changed "AddLastLeadingZeros" to "AddLeadingZeros" in the routine
'              "FormatBySection". Must always specify this parameter, to avoid
'              assuming the wrong value.
'            - Added a thrown error if AddLeadingZeros is False, yet the length
'              of any section is incorrect.
' 03/09/2010 - SBakker
'            - Added ValidatePostalCode, to check postal codes against the
'              country's postal pattern. The calling program would allow blank
'              first, if applicable.
' 08/10/2009 - SBakker
'            - Added new option "AddLastLeadingZeros" to "FormatBySection". Then
'              partial values can have a wildcard added as expected. Example is
'              UnformatSSNPartial("0057") would return "005-7", not "005-07".
' 04/28/2009 - SBakker
'            - Allow LTD Claim numbers to end in "...0000". Apparently we have
'              some that do.
' 03/03/2009 - SBakker
'            - Added UnformatSSN function for Kim. Is the same as FormatSSN.
' 02/17/2009 - SBakker
'            - Added "Unformat...Partial" functions. They will return partial
'              strings, used for searching, say when only Client or Client+Year
'              is entered. They also work with the full strings too.
' 02/13/2009 - SBakker
'            - Added FormatUtils to handle formatting and validating of any
'              standard fields. It also contains some Unformatting functions,
'              to return a value packed with leading zeros, but no separators.
' ----------------------------------------------------------------------------------------------------

Imports System.Text
Imports System.Text.RegularExpressions

Public Class FormatUtils

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

#Region " LTD Claim Number Routines "

    Public Shared Function FormatLTDClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {3, 2, 4}
            Dim MinValues() As Integer = {1, 0, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid LTD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatLTDClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {3, 2, 4}
            Dim MinValues() As Integer = {1, 0, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid LTD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatLTDClaimNumPartial(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {3, 2, 4}
            Dim MinValues() As Integer = {1, 0, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {3, 2}
            Dim MinValues() As Integer = {1, 0}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {3}
            Dim MinValues() As Integer = {1}
            Return FormatBySection(Value, 1, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid LTD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " STD Claim Number Routines "

    Public Shared Function FormatSTDClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 1998, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid STD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatSTDClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 1998, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid STD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatSTDClaimNumPartial(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 1998, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {4, 4}
            Dim MinValues() As Integer = {1, 1998}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {4}
            Dim MinValues() As Integer = {1}
            Return FormatBySection(Value, 1, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid STD Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " OLS Claim Number Routines "

    Public Shared Function FormatOLSClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 2015, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid OLS Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatOLSClaimNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 2015, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid OLS Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatOLSClaimNumPartial(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 4, 6}
            Dim MinValues() As Integer = {1, 2015, 1}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {4, 4}
            Dim MinValues() As Integer = {1, 2015}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "", True)
        Catch ex As Exception
            ' --- No error here ---
        End Try
        Try
            Dim SectionLen() As Integer = {4}
            Dim MinValues() As Integer = {1}
            Return FormatBySection(Value, 1, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid OLS Claim Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " Cession Number Routines "

    Public Shared Function FormatCessionNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 7, 4}
            Dim MinValues() As Integer = {1, 1, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Cession Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatCessionNum(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 7, 4}
            Dim MinValues() As Integer = {1, 1, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Cession Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatCessionNumPartial(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {4, 7, 4}
            Dim MinValues() As Integer = {1, 1, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "", True)
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {4, 7}
            Dim MinValues() As Integer = {1, 1}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "", True)
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {4}
            Dim MinValues() As Integer = {1}
            Return FormatBySection(Value, 1, SectionLen, MinValues, "", True)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Cession Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " Social Security Number Routines "

    Public Shared Function FormatSSN(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {3, 2, 4}
            Dim MinValues() As Integer = {0, 0, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", False)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Social Security Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function UnformatSSN(ByVal Value As String) As String
        ' --- This is just included for consistency ---
        Return FormatSSN(Value)
    End Function

    Public Shared Function UnformatSSNToNumber(ByVal Value As String) As Integer
        If String.IsNullOrEmpty(Value) OrElse Value.Length <> 11 Then
            Return 0
        End If
        Dim SSNMatch As Match = Regex.Match(Value, "\d\d\d-\d\d-\d\d\d\d")
        If SSNMatch.Length <> 11 Then
            Return 0
        End If
        Return CInt(Value.Replace("-", ""))
    End Function

    Public Shared Function UnformatSSNPartial(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {3, 2, 4}
            Dim MinValues() As Integer = {0, 0, 0}
            Return FormatBySection(Value, 3, SectionLen, MinValues, "-", False)
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {3, 2}
            Dim MinValues() As Integer = {0, 0}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "-", False)
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {3}
            Dim MinValues() As Integer = {0}
            Return FormatBySection(Value, 1, SectionLen, MinValues, "-", False)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Social Security Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " Tax ID Number Routines "

    Public Shared Function FormatTaxIDNumber(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {2, 7}
            Dim MinValues() As Integer = {1, 1}
            Return FormatBySection(Value, 2, SectionLen, MinValues, "-", False)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Tax ID Number: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " Zip/Postal Code Routines "

    Public Shared Function FormatUSZipCode(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {5, 4}
            Dim MinValues() As Integer = {1, 0}
            Dim Result As String = FormatBySection(Value, 2, SectionLen, MinValues, "-", False)
            If Result.EndsWith("-0000") Then
                Result = Left(Result, 5)
            End If
            Return Result
        Catch ex As Exception
            Try
                Dim SectionLen() As Integer = {5}
                Dim MinValues() As Integer = {0}
                Return FormatBySection(Value, 1, SectionLen, MinValues, "", False)
            Catch ex2 As Exception
                Throw New SystemException(FuncName + vbCrLf + "Invalid US Zip Code: """ + Value + """")
            End Try
        End Try
    End Function

    Public Shared Function FormatCanadianPostalCode(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the value ---
        Try
            Dim Result As New StringBuilder
            Dim CharCount As Integer = 0
            For Each c As Char In Value.ToUpper
                Select Case c
                    Case "A"c To "Z"c
                        CharCount += 1
                        If CharCount <> 1 AndAlso CharCount <> 3 AndAlso CharCount <> 5 Then
                            Throw New Exception ' caught below
                        End If
                        Result.Append(c)
                    Case "0"c To "9"c
                        CharCount += 1
                        If CharCount <> 2 AndAlso CharCount <> 4 AndAlso CharCount <> 6 Then
                            Throw New Exception ' caught below
                        End If
                        If CharCount = 4 Then
                            Result.Append(" ")
                        End If
                        Result.Append(c)
                End Select
            Next
            If CharCount <> 6 Then
                Throw New Exception ' caught below
            End If
            Return Result.ToString
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Canadian Postal Code: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

    Public Shared Function ValidatePostalCode(ByVal Value As String, ByVal Pattern As String) As Boolean
        ' --- Some countries have no postal pattern, so all are valid ---
        If String.IsNullOrEmpty(Pattern) Then Return True
        ' --- Get list of each possible pattern, separated by the pipe character ---
        Dim Patterns() As String = Pattern.ToUpper.Split("|"c)
        For Each CurrPattern As String In Patterns
            If Value.Length <> CurrPattern.Length Then Continue For
            For CurrIndex As Integer = 0 To CurrPattern.Length - 1
                Select Case CurrPattern.Chars(CurrIndex)
                    Case "#"c
                        If Value.Chars(CurrIndex) < "0"c OrElse Value.Chars(CurrIndex) > "9"c Then Return False
                    Case "A"c
                        If Value.Chars(CurrIndex) < "A"c OrElse Value.Chars(CurrIndex) > "Z"c Then Return False
                    Case Else
                        If Value.Chars(CurrIndex) <> CurrPattern.Chars(CurrIndex) Then Return False
                End Select
            Next
            Return True
        Next
        Return False
    End Function

#End Region

#Region " Bank Routing/Transit Codes "

    Public Shared Function ValidateBankRTCode(ByVal Value As Integer) As Boolean
        If Value = 0 Then Return True
        If Value < 0 Or Value > 999999999 Then Return False
        Return ValidateBankRTCode(Right("000000000" + Value.ToString, 9))
    End Function

    Public Shared Function ValidateBankRTCode(ByVal Value As String) As Boolean
        If String.IsNullOrEmpty(Value) Then Return True
        If Value = "0" Then Return True
        If Value.Length <> 9 Then Return False
        If Not IsNumeric(Value) Then Return False
        ' --- The first two digits must be in the ranges 00 through 12, 21 through 32, 61 through 72, or 80. ---
        Try
            Dim FirstTwoDigits As Integer = CInt(Value.Substring(0, 2))
            Dim FirstTwoOK As Boolean = False
            If FirstTwoDigits >= 0 AndAlso FirstTwoDigits <= 12 Then FirstTwoOK = True
            If FirstTwoDigits >= 21 AndAlso FirstTwoDigits <= 32 Then FirstTwoOK = True
            If FirstTwoDigits >= 61 AndAlso FirstTwoDigits <= 72 Then FirstTwoOK = True
            If FirstTwoDigits = 80 Then FirstTwoOK = True
            If Not FirstTwoOK Then Return False
        Catch ex As Exception
            Return False
        End Try
        ' --- Validate that Checksum mod 10 = 0 ---
        Dim CheckSum As Integer = 0
        Dim CurrDigit As Integer = 0
        For CharNum As Integer = 0 To Value.Length - 1
            If Not Char.IsDigit(Value(CharNum)) Then Return False
            CurrDigit = CInt(CStr(Value(CharNum)))
            Select Case CharNum Mod 3
                Case 0
                    CheckSum += CurrDigit * 3
                Case 1
                    CheckSum += CurrDigit * 7
                Case 2
                    CheckSum += CurrDigit
            End Select
        Next
        Return ((CheckSum Mod 10) = 0)
    End Function

#End Region

#Region " Phone Number Routines "

    Public Shared Function UnformatPhoneToNumber(ByVal Value As String) As Decimal
        If String.IsNullOrEmpty(Value) OrElse Value.Length <> 12 Then
            Return 0
        End If
        Dim SSNMatch As Match = Regex.Match(Value, "\d\d\d-\d\d\d-\d\d\d\d")
        If SSNMatch.Length <> 12 Then
            Return 0
        End If
        Return CDec(Value.Replace("-", ""))
    End Function

#End Region

#Region " FormatBySection Routines "

    Public Shared Function FormatBySection(ByVal Value As String, _
                                           ByVal NumSections As Integer, _
                                           ByVal SectionLen() As Integer, _
                                           ByVal MinValues() As Integer, _
                                           ByVal Separator As String, _
                                           ByVal AddLeadingZeros As Boolean) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim Result As New StringBuilder
        Dim Section(NumSections - 1) As String
        Dim CurrSection As Integer = -1
        Dim InSection As Boolean = False
        Dim CharCount As Integer = 0
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Parse the string into sections ---
        For Each c As Char In Value
            ' --- Check for valid chars ---
            If c >= "0"c AndAlso c <= "9"c Then
                ' --- Jump to next section ---
                If Not InSection Then
                    CurrSection += 1
                    InSection = True
                    CharCount = 0
                End If
                CharCount += 1
                ' --- If past end of section, move to next section ---
                If CurrSection < NumSections AndAlso CharCount > SectionLen(CurrSection) Then
                    CurrSection += 1
                    CharCount = 1
                End If
                ' --- Add value to current section. Will throw an error if too many sections. ---
                Section(CurrSection) += c
            Else
                InSection = False
            End If
        Next
        Try
            ' --- Make sure there were exactly enough sections ---
            If CurrSection <> NumSections - 1 Then
                Throw New Exception ' caught below
            End If
            ' --- Build and return the result ---
            For TempSection As Integer = 0 To NumSections - 1
                ' --- Check against MinValues ---
                If CInt(Section(TempSection)) < MinValues(TempSection) Then
                    Throw New Exception ' caught below
                End If
                ' --- Add Separator ---
                If TempSection > 0 Then
                    Result.Append(Separator)
                End If
                ' --- Add leading zeros ---
                If Section(TempSection).Length < SectionLen(TempSection) Then
                    If AddLeadingZeros Then
                        Result.Append(StrDup(SectionLen(TempSection) - Section(TempSection).Length, "0"c)) ' add leading zeros
                    Else
                        Throw New Exception ' caught below
                    End If
                End If
                ' --- Add actual info ---
                Result.Append(Section(TempSection))
            Next
            Return Result.ToString
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Error during format: """ + Value + """" + vbCrLf + ex.Message)
        End Try
    End Function

#End Region

#Region " Format Boolean Routines "

    Public Shared Function ToYN(ByVal Value As Boolean) As String
        If Value Then
            Return "Y"
        Else
            Return "N"
        End If
    End Function

    ''' <summary>
    ''' "Y" = True, otherwise False
    ''' </summary>
    Public Shared Function FromYN(ByVal Value As String) As Boolean
        Return (Value.ToUpper.StartsWith("Y"))
    End Function

    Public Shared Function ToYNBlank(ByVal Value As Nullable(Of Boolean)) As String
        If Not Value.HasValue Then
            Return ""
        ElseIf Value.Value Then
            Return "Y"
        Else
            Return "N"
        End If
    End Function

    ''' <summary>
    ''' "Y" = True, "N" = False, otherwise Nothing
    ''' </summary>
    Public Shared Function FromYNBlank(ByVal Value As String) As Nullable(Of Boolean)
        If Value.ToUpper.StartsWith("Y") Then
            Return True
        ElseIf Value.ToUpper.StartsWith("N") Then
            Return False
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function ToYNNull(ByVal Value As Nullable(Of Boolean)) As String
        If Not Value.HasValue Then
            Return Nothing
        ElseIf Value.Value Then
            Return "Y"
        Else
            Return "N"
        End If
    End Function

#End Region

#Region " Format Number Routines "

    ' --- Zero ---

    ''' <summary>
    ''' Formats a value with zero decimal places, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatNumber0(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatNumber0(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a value with zero decimal places.
    ''' </summary>
    Public Shared Function FormatNumber0(ByVal Value As Decimal) As String
        Return FormatNumber(Value, 0, TriState.True, TriState.False, TriState.False)
    End Function

    ' --- Two ---

    ''' <summary>
    ''' Formats a value with two decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatNumber2(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatNumber2(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a value with two decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatNumber2(ByVal Value As Decimal) As String
        Return FormatNumber(Value, 2, TriState.True, TriState.False, TriState.False)
    End Function

    ' --- Three ---

    ''' <summary>
    ''' Formats a value with three decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatNumber3(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatNumber3(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a value with three decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatNumber3(ByVal Value As Decimal) As String
        Return FormatNumber(Value, 3, TriState.True, TriState.False, TriState.False)
    End Function

    ' --- Four ---

    ''' <summary>
    ''' Formats a value with four decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatNumber4(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatNumber4(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a value with four decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatNumber4(ByVal Value As Decimal) As String
        Return FormatNumber(Value, 4, TriState.True, TriState.False, TriState.False)
    End Function

    ' --- Five ---

    ''' <summary>
    ''' Formats a value with five decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatNumber5(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatNumber5(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a value with five decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatNumber5(ByVal Value As Decimal) As String
        Return FormatNumber(Value, 5, TriState.True, TriState.False, TriState.False)
    End Function

    ' --- UnformatNumber ---

    ''' <summary>
    ''' Returns a decimal from a string. Empty string = 0. Throws an error for invalid strings.
    ''' </summary>
    Public Shared Function UnformatNumber(ByVal Value As String) As Decimal
        If String.IsNullOrEmpty(Value) Then
            Return 0
        End If
        Return CDec(Value.Trim)
    End Function

    ''' <summary>
    ''' Returns a decimal from a string. Empty string = Nothing. Throws an error for invalid strings.
    ''' </summary>
    Public Shared Function UnformatNumberNullable(ByVal Value As String) As Decimal?
        If String.IsNullOrEmpty(Value) Then
            Return Nothing
        End If
        Return CDec(Value.Trim)
    End Function

    ''' <summary>
    ''' Returns an integer from a string. Empty string = 0. Throws an error for invalid strings.
    ''' </summary>
    Public Shared Function UnformatNumberInt(ByVal Value As String) As Integer
        If String.IsNullOrEmpty(Value) Then
            Return 0
        End If
        Return CInt(Value.Trim)
    End Function

#End Region

#Region " Format Currency Routines "

    ' --- Zero ---

    ''' <summary>
    ''' Formats a dollar value with zero decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatCurrency0(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatCurrency0(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a dollar value with zero decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatCurrency0(ByVal Value As Decimal) As String
        Return FormatCurrency(Value, 0, TriState.True, TriState.False, TriState.True)
    End Function

    ' --- Two ---

    ''' <summary>
    ''' Formats a dollar value with two decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatCurrency2(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatCurrency2(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a dollar value with two decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatCurrency2(ByVal Value As Decimal) As String
        Return FormatCurrency(Value, 2, TriState.True, TriState.False, TriState.True)
    End Function

    ' --- Three ---

    ''' <summary>
    ''' Formats a dollar value with three decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatCurrency3(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatCurrency3(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a dollar value with three decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatCurrency3(ByVal Value As Decimal) As String
        Return FormatCurrency(Value, 3, TriState.True, TriState.False, TriState.True)
    End Function

    ' --- Four ---

    ''' <summary>
    ''' Formats a dollar value with four decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatCurrency4(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatCurrency4(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a dollar value with four decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatCurrency4(ByVal Value As Decimal) As String
        Return FormatCurrency(Value, 4, TriState.True, TriState.False, TriState.True)
    End Function

    ' --- Five ---

    ''' <summary>
    ''' Formats a dollar value with five decimal places and a leading zero, or blank if Nothing.
    ''' </summary>
    Public Shared Function FormatCurrency5(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then
            Return ""
        End If
        Return FormatCurrency5(Value.Value)
    End Function

    ''' <summary>
    ''' Formats a dollar value with five decimal places and a leading zero.
    ''' </summary>
    Public Shared Function FormatCurrency5(ByVal Value As Decimal) As String
        Return FormatCurrency(Value, 5, TriState.True, TriState.False, TriState.True)
    End Function

#End Region

#Region " PlusIfPositive Routines "

    ''' <summary>
    ''' Returns a "+" if the Value is greater than zero, or "" otherwise.
    ''' </summary>
    Public Shared Function PlusIfPositive(ByVal Value As Nullable(Of Decimal)) As String
        If Not Value.HasValue Then Return ""
        Return PlusIfPositive(Value.Value)
    End Function

    ''' <summary>
    ''' Returns a "+" if the Value is greater than zero, or "" otherwise.
    ''' </summary>
    Public Shared Function PlusIfPositive(ByVal Value As Decimal) As String
        If Value > 0 Then Return "+"
        Return ""
    End Function

#End Region

End Class
