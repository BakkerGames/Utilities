' ---------------------------------------
' --- PseudoVariables.vb - 10/26/2016 ---
' ---------------------------------------

Module PseudoVariables

#Region " N "

    Public Property N(ByVal Offset As Int64) As Int64
        Get
            Select Case Offset
                Case 0 : Return N0
                Case 1 : Return N1
                Case 2 : Return N2
                Case 3 : Return N3
                Case 4 : Return N4
                Case 5 : Return N5
                Case 6 : Return N6
                Case 7 : Return N7
                Case 8 : Return N8
                Case 9 : Return N9
                Case 10 : Return N10
                Case 11 : Return N11
                Case 12 : Return N12
                Case 13 : Return N13
                Case 14 : Return N14
                Case 15 : Return N15
                Case 16 : Return N16
                Case 17 : Return N17
                Case 18 : Return N18
                Case 19 : Return N19
                Case 20 : Return N20
                Case 21 : Return N21
                Case 22 : Return N22
                Case 23 : Return N23
                Case 24 : Return N24
                Case 25 : Return N25
                Case 26 : Return N26
                Case 27 : Return N27
                Case 28 : Return N28
                Case 29 : Return N29
                Case 30 : Return N30
                Case 31 : Return N31
                Case 32 : Return N32
                Case 33 : Return N33
                Case 34 : Return N34
                Case 35 : Return N35
                Case 36 : Return N36
                Case 37 : Return N37
                Case 38 : Return N38
                Case 39 : Return N39
                Case 40 : Return N40
                Case 41 : Return N41
                Case 42 : Return N42
                Case 43 : Return N43
                Case 44 : Return N44
                Case 45 : Return N45
                Case 46 : Return N46
                Case 47 : Return N47
                Case 48 : Return N48
                Case 49 : Return N49
                Case 50 : Return N50
                Case 51 : Return N51
                Case 52 : Return N52
                Case 53 : Return N53
                Case 54 : Return N54
                Case 55 : Return N55
                Case 56 : Return N56
                Case 57 : Return N57
                Case 58 : Return N58
                Case 59 : Return N59
                Case 60 : Return N60
                Case 61 : Return N61
                Case 62 : Return N62
                Case 63 : Return N63
                Case 64 : Return N64
                Case 65 : Return N65
                Case 66 : Return N66
                Case 67 : Return N67
                Case 68 : Return N68
                Case 69 : Return N69
                Case 70 : Return N70
                Case 71 : Return N71
                Case 72 : Return N72
                Case 73 : Return N73
                Case 74 : Return N74
                Case 75 : Return N75
                Case 76 : Return N76
                Case 77 : Return N77
                Case 78 : Return N78
                Case 79 : Return N79
                Case 80 : Return N80
                Case 81 : Return N81
                Case 82 : Return N82
                Case 83 : Return N83
                Case 84 : Return N84
                Case 85 : Return N85
                Case 86 : Return N86
                Case 87 : Return N87
                Case 88 : Return N88
                Case 89 : Return N89
                Case 90 : Return N90
                Case 91 : Return N91
                Case 92 : Return N92
                Case 93 : Return N93
                Case 94 : Return N94
                Case 95 : Return N95
                Case 96 : Return N96
                Case 97 : Return N97
                Case 98 : Return N98
                Case 99 : Return N99
                Case Else
                    Throw New System.Exception("N: Invalid Offset: " + Offset.ToString)
            End Select
        End Get
        Set(value As Int64)
            Select Case Offset
                Case 0 : N0 = value
                Case 1 : N1 = value
                Case 2 : N2 = value
                Case 3 : N3 = value
                Case 4 : N4 = value
                Case 5 : N5 = value
                Case 6 : N6 = value
                Case 7 : N7 = value
                Case 8 : N8 = value
                Case 9 : N9 = value
                Case 10 : N10 = value
                Case 11 : N11 = value
                Case 12 : N12 = value
                Case 13 : N13 = value
                Case 14 : N14 = value
                Case 15 : N15 = value
                Case 16 : N16 = value
                Case 17 : N17 = value
                Case 18 : N18 = value
                Case 19 : N19 = value
                Case 20 : N20 = value
                Case 21 : N21 = value
                Case 22 : N22 = value
                Case 23 : N23 = value
                Case 24 : N24 = value
                Case 25 : N25 = value
                Case 26 : N26 = value
                Case 27 : N27 = value
                Case 28 : N28 = value
                Case 29 : N29 = value
                Case 30 : N30 = value
                Case 31 : N31 = value
                Case 32 : N32 = value
                Case 33 : N33 = value
                Case 34 : N34 = value
                Case 35 : N35 = value
                Case 36 : N36 = value
                Case 37 : N37 = value
                Case 38 : N38 = value
                Case 39 : N39 = value
                Case 40 : N40 = value
                Case 41 : N41 = value
                Case 42 : N42 = value
                Case 43 : N43 = value
                Case 44 : N44 = value
                Case 45 : N45 = value
                Case 46 : N46 = value
                Case 47 : N47 = value
                Case 48 : N48 = value
                Case 49 : N49 = value
                Case 50 : N50 = value
                Case 51 : N51 = value
                Case 52 : N52 = value
                Case 53 : N53 = value
                Case 54 : N54 = value
                Case 55 : N55 = value
                Case 56 : N56 = value
                Case 57 : N57 = value
                Case 58 : N58 = value
                Case 59 : N59 = value
                Case 60 : N60 = value
                Case 61 : N61 = value
                Case 62 : N62 = value
                Case 63 : N63 = value
                Case 64 : N64 = value
                Case 65 : N65 = value
                Case 66 : N66 = value
                Case 67 : N67 = value
                Case 68 : N68 = value
                Case 69 : N69 = value
                Case 70 : N70 = value
                Case 71 : N71 = value
                Case 72 : N72 = value
                Case 73 : N73 = value
                Case 74 : N74 = value
                Case 75 : N75 = value
                Case 76 : N76 = value
                Case 77 : N77 = value
                Case 78 : N78 = value
                Case 79 : N79 = value
                Case 80 : N80 = value
                Case 81 : N81 = value
                Case 82 : N82 = value
                Case 83 : N83 = value
                Case 84 : N84 = value
                Case 85 : N85 = value
                Case 86 : N86 = value
                Case 87 : N87 = value
                Case 88 : N88 = value
                Case 89 : N89 = value
                Case 90 : N90 = value
                Case 91 : N91 = value
                Case 92 : N92 = value
                Case 93 : N93 = value
                Case 94 : N94 = value
                Case 95 : N95 = value
                Case 96 : N96 = value
                Case 97 : N97 = value
                Case 98 : N98 = value
                Case 99 : N99 = value
                Case Else
                    Throw New System.Exception("N: Invalid Offset: " + Offset.ToString)
            End Select
        End Set
    End Property

#End Region

#Region " F "

    Public Property F(ByVal Offset As Int64) As Int64
        Get
            Select Case Offset
                Case MinFLow To MaxFLow
                    Return MEM(MemPos_F + CInt(Offset))
                Case MinFHigh To MaxFHigh
                    Return MEM(MemPos_F10 + CInt(Offset))
                Case Else
                    Throw New System.Exception("F: Invalid Offset: " + Offset.ToString)
            End Select
        End Get
        Set(value As Int64)
            Select Case Offset
                Case MinFLow To MaxFLow
                    LetByte(MemPos_F + CInt(Offset), value)
                Case MinFHigh To MaxFHigh
                    LetByte(MemPos_F10 + CInt(Offset), value)
                Case Else
                    Throw New System.Exception("F: Invalid Offset: " + Offset.ToString)
            End Select
        End Set
    End Property

#End Region

#Region " G "

    Public Property G(ByVal Offset As Integer) As Int64
        Get
            Return MEM(MemPos_G + Offset)
        End Get
        Set(value As Int64)
            LetByte(MemPos_G + Offset, value)
        End Set
    End Property

#End Region

#Region " General ALPHA Registers "

    Public Property KEY As String
        Get
            Return GetAlpha(MemPos_Key)
        End Get
        Set(value As String)
            LetAlpha(MemPos_Key, value)
        End Set
    End Property

    Public Property KEY(ByVal Offset As Integer) As String
        Get
            Return GetAlpha(MemPos_Key + Offset)
        End Get
        Set(value As String)
            LetAlpha(MemPos_Key + Offset, value)
        End Set
    End Property

    Public Property DATEVAL As String
        Get
            Return GetAlpha(MemPos_DateVal)
        End Get
        Set(value As String)
            LetAlpha(MemPos_DateVal, value)
        End Set
    End Property

    Public Property DATEVAL(ByVal Offset As Integer) As String
        Get
            Return GetAlpha(MemPos_DateVal + Offset)
        End Get
        Set(value As String)
            LetAlpha(MemPos_DateVal + Offset, value)
        End Set
    End Property

#End Region

End Module
