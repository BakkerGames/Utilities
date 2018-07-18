' ---------------------------------------
' --- MemoryPositions.vb - 07/20/2011 ---
' ---------------------------------------

' ------------------------------------------------------------------------------
' 01/20/2006 - Added N64-N99, A3-A9 to E3-E9, F10-F99.
' ------------------------------------------------------------------------------

Module MemoryPositions

    ' --------------------------------------------------------
    ' --- Memory Map (numbers are in decimal):             ---
    ' --- Page      0: System Variables                    ---
    ' --- Pages  1- 2: Numeric Variables                   ---
    ' --- Pages  2- 3: Alpha Variables A, B, C (Start=234) ---
    ' --- Pages  4- 8: Buffers R, Z, X, Y, W               ---
    ' --- Pages  9-10: File Table (32 entries)             ---
    ' --- Pages 11-14: Buffers S, T, U, V                  ---
    ' --- Page     15: Alpha Variables D, E                ---
    ' --- Pages 16-47: Track Buffer                        ---
    ' --- Page     48: Device Table (32 entries)           ---
    ' --- Pages 48-50: Volume Table (64 entries, Start=32) ---
    ' --- Pages 51-52: TFA Table (32 entries)              ---
    ' --- Pages 53-54: Channel Table (128 entries)         ---
    ' --- Page     55: High Numeric Variables N64-N99      ---
    ' --- Pages 56-58: High Alpha Vars A3-A9 to E3-E9      ---
    ' --------------------------------------------------------

    ' --------------------------------------------------------
    ' --- File table record layout (16 bytes):             ---
    ' ---    Ofs 0: Open (True/False)                      ---
    ' ---    Ofs 1: Volume Number                          ---
    ' ---    Ofs 2: Type (2=Data, 5=Directory, 6=Pseudo)   ---
    ' ---    Ofs 3: Name (a8)                              ---
    ' --------------------------------------------------------
    ' --- Device table record layout (1 byte):             ---
    ' ---    Ofs 0: Open (True/False)                      ---
    ' --------------------------------------------------------
    ' --- Volume table record layout (10 bytes):           ---
    ' ---    Ofs 0: Open (True/False)                      ---
    ' ---    Ofs 1: Device Number                          ---
    ' ---    Ofs 2: Name (a8)                              ---
    ' --------------------------------------------------------
    ' --- TFA table record layout (10 bytes):              ---
    ' ---    Ofs 0: Open (True/False)                      ---
    ' ---    Ofs 1: Volume Number                          ---
    ' ---    Ofs 2: Name (a8)                              ---
    ' --------------------------------------------------------
    ' --- Channel table record layout (3 bytes):           ---
    ' ---    Ofs 0: Open (True/False)                      ---
    ' ---    Ofs 1: TFA Number                             ---
    ' ---    Ofs 2: Channel Locked (True/False)            ---
    ' ---    Channel path is stored in ChannelPaths().     ---
    ' --------------------------------------------------------

    ' --- Table Starting Positions ---

    Public Const MemPos_FileTable As Integer = 9 * 256
    Public Const MemPos_TrackBuffer As Integer = 16 * 256
    Public Const MemPos_DevTable As Integer = 48 * 256
    Public Const MemPos_VolTable As Integer = MemPos_DevTable + 32
    Public Const MemPos_TFATable As Integer = 51 * 256
    Public Const MemPos_ChanTable As Integer = 53 * 256
    Public Const MemPos_NumHigh As Integer = 55 * 256
    Public Const MemPos_HighAlpha As Integer = 56 * 256
    Public Const TotalMemSize As Integer = 59 * 256

    ' --- Table Record Sizes ---

    Public Const FileEntrySize As Integer = 16
    Public Const DevEntrySize As Integer = 1
    Public Const VolEntrySize As Integer = 10
    Public Const TFAEntrySize As Integer = 10
    Public Const ChanEntrySize As Integer = 3

    ' --- Highest Table Entry numbers ---

    Public Const MaxFile As Integer = 31
    Public Const MaxDevice As Integer = 31
    Public Const MaxVolume As Integer = 63
    Public Const MaxTFA As Integer = 31
    Public Const MaxChannel As Integer = 127

    ' --- Buffer Pointers ---

    Public Const MemPos_RP As Integer = 0                    ' x00
    Public Const MemPos_RP2 As Integer = 1                   ' x01
    Public Const MemPos_IRP As Integer = 2                   ' x02
    Public Const MemPos_IRP2 As Integer = 3                  ' x03
    Public Const MemPos_ZP As Integer = 4                    ' x04
    Public Const MemPos_ZP2 As Integer = 5                   ' x05
    Public Const MemPos_IZP As Integer = 6                   ' x06
    Public Const MemPos_IZP2 As Integer = 7                  ' x07
    Public Const MemPos_XP As Integer = 8                    ' x08
    Public Const MemPos_XP2 As Integer = 9                   ' x09
    Public Const MemPos_IXP As Integer = 10                  ' x0A
    Public Const MemPos_IXP2 As Integer = 11                 ' x0B
    Public Const MemPos_YP As Integer = 12                   ' x0C
    Public Const MemPos_YP2 As Integer = 13                  ' x0D
    Public Const MemPos_IYP As Integer = 14                  ' x0E
    Public Const MemPos_IYP2 As Integer = 15                 ' x0F
    Public Const MemPos_WP As Integer = 16                   ' x10
    Public Const MemPos_WP2 As Integer = 17                  ' x11
    Public Const MemPos_IWP As Integer = 18                  ' x12
    Public Const MemPos_IWP2 As Integer = 19                 ' x13

    Public Const MemPos_SP As Integer = 112                  ' x70
    Public Const MemPos_SP2 As Integer = 113                 ' x71
    Public Const MemPos_ISP As Integer = 114                 ' x72
    Public Const MemPos_ISP2 As Integer = 115                ' x73
    Public Const MemPos_TP As Integer = 116                  ' x74
    Public Const MemPos_TP2 As Integer = 117                 ' x75
    Public Const MemPos_ITP As Integer = 118                 ' x76
    Public Const MemPos_ITP2 As Integer = 119                ' x77
    Public Const MemPos_UP As Integer = 120                  ' x78
    Public Const MemPos_UP2 As Integer = 121                 ' x79
    Public Const MemPos_IUP As Integer = 122                 ' x7A
    Public Const MemPos_IUP2 As Integer = 123                ' x7B
    Public Const MemPos_VP As Integer = 124                  ' x7C
    Public Const MemPos_VP2 As Integer = 125                 ' x7D
    Public Const MemPos_IVP As Integer = 126                 ' x7E
    Public Const MemPos_IVP2 As Integer = 127                ' x7F

    ' --- Byte System Variables ---

    Public Const MemPos_Lib As Integer = 20                  ' x14
    Public Const MemPos_Prog As Integer = 21                 ' x15-x16
    Public Const MemPos_Privg As Integer = 46                ' x2E
    Public Const MemPos_Char As Integer = 48                 ' x30
    Public Const MemPos_Length As Integer = 49               ' x31
    Public Const MemPos_Status As Integer = 50               ' x32
    Public Const MemPos_EscVal As Integer = 51               ' x33
    Public Const MemPos_CanVal As Integer = 52               ' x34
    Public Const MemPos_LockVal As Integer = 53              ' x35
    Public Const MemPos_TChan As Integer = 54                ' x36
    Public Const MemPos_Term As Integer = 73                 ' x49
    Public Const MemPos_Lang As Integer = 76                 ' x4C
    Public Const MemPos_PrtNum As Integer = 77               ' x4D
    Public Const MemPos_TFA As Integer = 89                  ' x59
    Public Const MemPos_Vol As Integer = 90                  ' x5A
    Public Const MemPos_PVol As Integer = 91                 ' x5B
    Public Const MemPos_ReqVol As Integer = 103              ' x67

    ' --- 2-Byte System Variables ---

    Public Const MemPos_User As Integer = 102                ' x66-x67
    Public Const MemPos_Orig As Integer = 104                ' x68-x69
    Public Const MemPos_Oper As Integer = 106                ' x6A-x6B

    ' --- Read-Only Byte System Variables ---

    Public Const MemPos_MachType As Integer = 109            ' x6D as integer = Constant 20 in IDRIS
    Public Const MemPos_SysRel As Integer = 110              ' x6E as integer = Constant 5 in IDRIS
    Public Const MemPos_SysRev As Integer = 111              ' x6F as integer = Constant 0 in IDRIS

    ' --- Flag registers ---

    Public Const MemPos_F As Integer = 92                    ' x5C
    Public Const MemPos_F1 As Integer = 93                   ' x5D
    Public Const MemPos_F2 As Integer = 94                   ' x5E
    Public Const MemPos_F3 As Integer = 95                   ' x5F
    Public Const MemPos_F4 As Integer = 96                   ' x60
    Public Const MemPos_F5 As Integer = 97                   ' x61
    Public Const MemPos_F6 As Integer = 98                   ' x62
    Public Const MemPos_F7 As Integer = 99                   ' x63
    Public Const MemPos_F8 As Integer = 100                  ' x64
    Public Const MemPos_F9 As Integer = 101                  ' x65
    Public Const MinFLow As Integer = 0
    Public Const MaxFLow As Integer = 9

    Public Const MemPos_F10 As Integer = 160                 ' xA0
    Public Const MemPos_F11 As Integer = 161                 ' xA1
    Public Const MemPos_F12 As Integer = 162                 ' xA2
    Public Const MemPos_F13 As Integer = 163                 ' xA3
    Public Const MemPos_F14 As Integer = 164                 ' xA4
    Public Const MemPos_F15 As Integer = 165                 ' xA5
    Public Const MemPos_F16 As Integer = 166                 ' xA6
    Public Const MemPos_F17 As Integer = 167                 ' xA7
    Public Const MemPos_F18 As Integer = 168                 ' xA8
    Public Const MemPos_F19 As Integer = 169                 ' xA9
    Public Const MemPos_F20 As Integer = 170                 ' xAA
    Public Const MemPos_F21 As Integer = 171                 ' xAB
    Public Const MemPos_F22 As Integer = 172                 ' xAC
    Public Const MemPos_F23 As Integer = 173                 ' xAD
    Public Const MemPos_F24 As Integer = 174                 ' xAE
    Public Const MemPos_F25 As Integer = 175                 ' xAF
    Public Const MemPos_F26 As Integer = 176                 ' xB0
    Public Const MemPos_F27 As Integer = 177                 ' xB1
    Public Const MemPos_F28 As Integer = 178                 ' xB2
    Public Const MemPos_F29 As Integer = 179                 ' xB3
    Public Const MemPos_F30 As Integer = 180                 ' xB4
    Public Const MemPos_F31 As Integer = 181                 ' xB5
    Public Const MemPos_F32 As Integer = 182                 ' xB6
    Public Const MemPos_F33 As Integer = 183                 ' xB7
    Public Const MemPos_F34 As Integer = 184                 ' xB8
    Public Const MemPos_F35 As Integer = 185                 ' xB9
    Public Const MemPos_F36 As Integer = 186                 ' xBA
    Public Const MemPos_F37 As Integer = 187                 ' xBB
    Public Const MemPos_F38 As Integer = 188                 ' xBC
    Public Const MemPos_F39 As Integer = 189                 ' xBD
    Public Const MemPos_F40 As Integer = 190                 ' xBE
    Public Const MemPos_F41 As Integer = 191                 ' xBF
    Public Const MemPos_F42 As Integer = 192                 ' xC0
    Public Const MemPos_F43 As Integer = 193                 ' xC1
    Public Const MemPos_F44 As Integer = 194                 ' xC2
    Public Const MemPos_F45 As Integer = 195                 ' xC3
    Public Const MemPos_F46 As Integer = 196                 ' xC4
    Public Const MemPos_F47 As Integer = 197                 ' xC5
    Public Const MemPos_F48 As Integer = 198                 ' xC6
    Public Const MemPos_F49 As Integer = 199                 ' xC7
    Public Const MemPos_F50 As Integer = 200                 ' xC8
    Public Const MemPos_F51 As Integer = 201                 ' xC9
    Public Const MemPos_F52 As Integer = 202                 ' xCA
    Public Const MemPos_F53 As Integer = 203                 ' xCB
    Public Const MemPos_F54 As Integer = 204                 ' xCC
    Public Const MemPos_F55 As Integer = 205                 ' xCD
    Public Const MemPos_F56 As Integer = 206                 ' xCE
    Public Const MemPos_F57 As Integer = 207                 ' xCF
    Public Const MemPos_F58 As Integer = 208                 ' xD0
    Public Const MemPos_F59 As Integer = 209                 ' xD1
    Public Const MemPos_F60 As Integer = 210                 ' xD2
    Public Const MemPos_F61 As Integer = 211                 ' xD3
    Public Const MemPos_F62 As Integer = 212                 ' xD4
    Public Const MemPos_F63 As Integer = 213                 ' xD5
    Public Const MemPos_F64 As Integer = 214                 ' xD6
    Public Const MemPos_F65 As Integer = 215                 ' xD7
    Public Const MemPos_F66 As Integer = 216                 ' xD8
    Public Const MemPos_F67 As Integer = 217                 ' xD9
    Public Const MemPos_F68 As Integer = 218                 ' xDA
    Public Const MemPos_F69 As Integer = 219                 ' xDB
    Public Const MemPos_F70 As Integer = 220                 ' xDC
    Public Const MemPos_F71 As Integer = 221                 ' xDD
    Public Const MemPos_F72 As Integer = 222                 ' xDE
    Public Const MemPos_F73 As Integer = 223                 ' xDF
    Public Const MemPos_F74 As Integer = 224                 ' xE0
    Public Const MemPos_F75 As Integer = 225                 ' xE1
    Public Const MemPos_F76 As Integer = 226                 ' xE2
    Public Const MemPos_F77 As Integer = 227                 ' xE3
    Public Const MemPos_F78 As Integer = 228                 ' xE4
    Public Const MemPos_F79 As Integer = 229                 ' xE5
    Public Const MemPos_F80 As Integer = 230                 ' xE6
    Public Const MemPos_F81 As Integer = 231                 ' xE7
    Public Const MemPos_F82 As Integer = 232                 ' xE8
    Public Const MemPos_F83 As Integer = 233                 ' xE9
    Public Const MemPos_F84 As Integer = 234                 ' xEA
    Public Const MemPos_F85 As Integer = 235                 ' xEB
    Public Const MemPos_F86 As Integer = 236                 ' xEC
    Public Const MemPos_F87 As Integer = 237                 ' xED
    Public Const MemPos_F88 As Integer = 238                 ' xEE
    Public Const MemPos_F89 As Integer = 239                 ' xEF
    Public Const MemPos_F90 As Integer = 240                 ' xF0
    Public Const MemPos_F91 As Integer = 241                 ' xF1
    Public Const MemPos_F92 As Integer = 242                 ' xF2
    Public Const MemPos_F93 As Integer = 243                 ' xF3
    Public Const MemPos_F94 As Integer = 244                 ' xF4
    Public Const MemPos_F95 As Integer = 245                 ' xF5
    Public Const MemPos_F96 As Integer = 246                 ' xF6
    Public Const MemPos_F97 As Integer = 247                 ' xF7
    Public Const MemPos_F98 As Integer = 248                 ' xF8
    Public Const MemPos_F99 As Integer = 249                 ' xF9
    Public Const MinFHigh As Integer = 10
    Public Const MaxFHigh As Integer = 99

    ' --- Local copy of Global Flag registers ---

    Public Const MemPos_G As Integer = 128                   ' x80-x81
    Public Const MemPos_G1 As Integer = 130                  ' x82
    Public Const MemPos_G2 As Integer = 131                  ' x83
    Public Const MemPos_G3 As Integer = 132                  ' x84
    Public Const MemPos_G4 As Integer = 133                  ' x85
    Public Const MemPos_G5 As Integer = 134                  ' x86
    Public Const MemPos_G6 As Integer = 135                  ' x87
    Public Const MemPos_G7 As Integer = 136                  ' x88
    Public Const MemPos_G8 As Integer = 137                  ' x89
    Public Const MemPos_G9 As Integer = 138                  ' x8A

    ' --- New Internal IDRIS Variables ---

    Public Const MemPos_SortState As Integer = 139           ' x8B - New in IDRIS
    Public Const MemPos_PrintDev As Integer = 140            ' x8C - New in IDRIS
    Public Const MemPos_PrintOn As Integer = 141             ' x8D - New in IDRIS
    Public Const MemPos_Background As Integer = 142          ' x8E - New in IDRIS
    Public Const MemPos_TBAlloc As Integer = 143             ' x8F - New in IDRIS
    Public Const MemPos_FFPending As Integer = 144           ' x90 - New in IDRIS
    Public Const MemPos_PageHasData As Integer = 145         ' x91 - New in IDRIS
    Public Const MemPos_LineHasData As Integer = 146         ' x92 - New in IDRIS
    Public Const MemPos_ILMFlag As Integer = 147             ' x93 - New in IDRIS
    Public Const MemPos_LocalEdit As Integer = 148           ' x94 - New in IDRIS
    Public Const MemPos_ScriptRunFlag As Integer = 149       ' x95 - New in IDRIS
    Public Const MemPos_ScriptWriteFlag As Integer = 150     ' x96 - New in IDRIS

    ' --- Numeric 6-byte Variables ---

    Public Const NumSlotSize As Integer = 6

    Public Const MemPos_N As Integer = 256                   ' x0100
    Public Const MinNLow As Integer = 0
    Public Const MaxNLow As Integer = 63

    Public Const MemPos_Rec As Integer = MemPos_N + ((MaxNLow + 1) * NumSlotSize)
    Public Const MemPos_RemVal As Integer = MemPos_N + ((MaxNLow + 2) * NumSlotSize)

    Public Const MemPos_N64 As Integer = 55 * 256            ' x3700
    Public Const MinNHigh As Integer = 64
    Public Const MaxNHigh As Integer = 99

    ' --- Low Alpha Varables ---

    Public Const MemPos_DateVal As Integer = 746             ' x02EA
    Public Const MemPos_Key As Integer = MemPos_DateVal + 18 ' x02FC
    Public Const MemPos_A As Integer = MemPos_Key + 20       ' x0310
    Public Const MemPos_A1 As Integer = MemPos_A + 40        ' x0338
    Public Const MemPos_A2 As Integer = MemPos_A + 60        ' x034C
    Public Const MemPos_B As Integer = MemPos_A + 80         ' x0360
    Public Const MemPos_B1 As Integer = MemPos_B + 40        ' x0388
    Public Const MemPos_B2 As Integer = MemPos_B + 60        ' x039C
    Public Const MemPos_C As Integer = MemPos_B + 80         ' x03B0
    Public Const MemPos_C1 As Integer = MemPos_C + 40        ' x03D8
    Public Const MemPos_C2 As Integer = MemPos_C + 60        ' x03EC

    ' --- Medium Alpha Varables ---

    Public Const MemPos_D As Integer = 3840                  ' x0F00
    Public Const MemPos_D1 As Integer = MemPos_D + 40        ' x0F28
    Public Const MemPos_D2 As Integer = MemPos_D + 60        ' x0F3C
    Public Const MemPos_E As Integer = MemPos_D + 80         ' x0F50
    Public Const MemPos_E1 As Integer = MemPos_E + 40        ' x0F78
    Public Const MemPos_E2 As Integer = MemPos_E + 60        ' x0F8C

    ' --- High Alpha Varables ---

    Public Const MemPos_A3 As Integer = MemPos_HighAlpha     ' x3800
    Public Const MemPos_A4 As Integer = MemPos_A3 + 20       ' x3816
    Public Const MemPos_A5 As Integer = MemPos_A4 + 20       ' x382C
    Public Const MemPos_A6 As Integer = MemPos_A5 + 20       ' x3842
    Public Const MemPos_A7 As Integer = MemPos_A6 + 20       ' x3858
    Public Const MemPos_A8 As Integer = MemPos_A7 + 20       ' x386E
    Public Const MemPos_A9 As Integer = MemPos_A8 + 20       ' x3884

    Public Const MemPos_B3 As Integer = MemPos_A9 + 20       ' x389A
    Public Const MemPos_B4 As Integer = MemPos_B3 + 20       ' x38B0
    Public Const MemPos_B5 As Integer = MemPos_B4 + 20       ' x38C6
    Public Const MemPos_B6 As Integer = MemPos_B5 + 20       ' x38DC
    Public Const MemPos_B7 As Integer = MemPos_B6 + 20       ' x38F2
    Public Const MemPos_B8 As Integer = MemPos_B7 + 20       ' x3908
    Public Const MemPos_B9 As Integer = MemPos_B8 + 20       ' x391E

    Public Const MemPos_C3 As Integer = MemPos_B9 + 20       ' x3934
    Public Const MemPos_C4 As Integer = MemPos_C3 + 20       ' x394A
    Public Const MemPos_C5 As Integer = MemPos_C4 + 20       ' x3960
    Public Const MemPos_C6 As Integer = MemPos_C5 + 20       ' x3976
    Public Const MemPos_C7 As Integer = MemPos_C6 + 20       ' x398C
    Public Const MemPos_C8 As Integer = MemPos_C7 + 20       ' x39A2
    Public Const MemPos_C9 As Integer = MemPos_C8 + 20       ' x39B8

    Public Const MemPos_D3 As Integer = MemPos_C9 + 20       ' x39CE
    Public Const MemPos_D4 As Integer = MemPos_D3 + 20       ' x39E4
    Public Const MemPos_D5 As Integer = MemPos_D4 + 20       ' x39FA
    Public Const MemPos_D6 As Integer = MemPos_D5 + 20       ' x3A10
    Public Const MemPos_D7 As Integer = MemPos_D6 + 20       ' x3A26
    Public Const MemPos_D8 As Integer = MemPos_D7 + 20       ' x3A3C
    Public Const MemPos_D9 As Integer = MemPos_D8 + 20       ' x3A52

    Public Const MemPos_E3 As Integer = MemPos_D9 + 20       ' x3A68
    Public Const MemPos_E4 As Integer = MemPos_E3 + 20       ' x3A7E
    Public Const MemPos_E5 As Integer = MemPos_E4 + 20       ' x3A94
    Public Const MemPos_E6 As Integer = MemPos_E5 + 20       ' x3AAA
    Public Const MemPos_E7 As Integer = MemPos_E6 + 20       ' x3AC0
    Public Const MemPos_E8 As Integer = MemPos_E7 + 20       ' x3AD6
    Public Const MemPos_E9 As Integer = MemPos_E8 + 20       ' x3AEC

    ' --- Buffers ---

    Public Const MemPos_R As Integer = 4 * 256               ' x0400
    Public Const MemPos_Z As Integer = 5 * 256               ' x0500
    Public Const MemPos_X As Integer = 6 * 256               ' x0600
    Public Const MemPos_Y As Integer = 7 * 256               ' x0700
    Public Const MemPos_W As Integer = 8 * 256               ' x0800
    Public Const MemPos_S As Integer = 11 * 256              ' x0B00
    Public Const MemPos_T As Integer = 12 * 256              ' x0C00
    Public Const MemPos_U As Integer = 13 * 256              ' x0D00
    Public Const MemPos_V As Integer = 14 * 256              ' x0E00

End Module
