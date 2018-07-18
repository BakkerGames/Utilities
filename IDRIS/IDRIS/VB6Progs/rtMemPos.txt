Attribute VB_Name = "rtMemPos"
' -----------------------------
' --- rtMemPos - 01/20/2006 ---
' -----------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 01/20/2006 - Added N64-N99, A3-A9 to E3-E9, F10-F99.
' ------------------------------------------------------------------------------

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

Public Const MemPos_FileTable = 9 * 256
Public Const MemPos_TrackBuffer = 16 * 256
Public Const MemPos_DevTable = 48 * 256
Public Const MemPos_VolTable = MemPos_DevTable + 32
Public Const MemPos_TFATable = 51 * 256
Public Const MemPos_ChanTable = 53 * 256
Public Const MemPos_NumHigh = 55 * 256
Public Const MemPos_HighAlpha = 56 * 256
Public Const TotalMemSize = 59 * 256

' --- Table Record Sizes ---

Public Const FileEntrySize As Long = 16
Public Const DevEntrySize As Long = 1
Public Const VolEntrySize As Long = 10
Public Const TFAEntrySize As Long = 10
Public Const ChanEntrySize As Long = 3

' --- Highest Table Entry numbers ---

Public Const MaxFile As Long = 31
Public Const MaxDevice As Long = 31
Public Const MaxVolume As Long = 63
Public Const MaxTFA As Long = 31
Public Const MaxChannel As Long = 127

' --- Buffer Pointers ---

Public Const MemPos_RP = 0                    ' x00
Public Const MemPos_RP2 = 1                   ' x01
Public Const MemPos_IRP = 2                   ' x02
Public Const MemPos_IRP2 = 3                  ' x03
Public Const MemPos_ZP = 4                    ' x04
Public Const MemPos_ZP2 = 5                   ' x05
Public Const MemPos_IZP = 6                   ' x06
Public Const MemPos_IZP2 = 7                  ' x07
Public Const MemPos_XP = 8                    ' x08
Public Const MemPos_XP2 = 9                   ' x09
Public Const MemPos_IXP = 10                  ' x0A
Public Const MemPos_IXP2 = 11                 ' x0B
Public Const MemPos_YP = 12                   ' x0C
Public Const MemPos_YP2 = 13                  ' x0D
Public Const MemPos_IYP = 14                  ' x0E
Public Const MemPos_IYP2 = 15                 ' x0F
Public Const MemPos_WP = 16                   ' x10
Public Const MemPos_WP2 = 17                  ' x11
Public Const MemPos_IWP = 18                  ' x12
Public Const MemPos_IWP2 = 19                 ' x13

Public Const MemPos_SP = 112                  ' x70
Public Const MemPos_SP2 = 113                 ' x71
Public Const MemPos_ISP = 114                 ' x72
Public Const MemPos_ISP2 = 115                ' x73
Public Const MemPos_TP = 116                  ' x74
Public Const MemPos_TP2 = 117                 ' x75
Public Const MemPos_ITP = 118                 ' x76
Public Const MemPos_ITP2 = 119                ' x77
Public Const MemPos_UP = 120                  ' x78
Public Const MemPos_UP2 = 121                 ' x79
Public Const MemPos_IUP = 122                 ' x7A
Public Const MemPos_IUP2 = 123                ' x7B
Public Const MemPos_VP = 124                  ' x7C
Public Const MemPos_VP2 = 125                 ' x7D
Public Const MemPos_IVP = 126                 ' x7E
Public Const MemPos_IVP2 = 127                ' x7F

   ' --- Byte System Variables ---

Public Const MemPos_Lib = 20                  ' x14
Public Const MemPos_Prog = 21                 ' x15-x16
Public Const MemPos_Privg = 46                ' x2E
Public Const MemPos_Char = 48                 ' x30
Public Const MemPos_Length = 49               ' x31
Public Const MemPos_Status = 50               ' x32
Public Const MemPos_EscVal = 51               ' x33
Public Const MemPos_CanVal = 52               ' x34
Public Const MemPos_LockVal = 53              ' x35
Public Const MemPos_TChan = 54                ' x36
Public Const MemPos_Term = 73                 ' x49
Public Const MemPos_Lang = 76                 ' x4C
Public Const MemPos_PrtNum = 77               ' x4D
Public Const MemPos_TFA = 89                  ' x59
Public Const MemPos_Vol = 90                  ' x5A
Public Const MemPos_PVol = 91                 ' x5B
Public Const MemPos_ReqVol = 103              ' x67

' --- 2-Byte System Variables ---

Public Const MemPos_User = 102                ' x66-x67
Public Const MemPos_Orig = 104                ' x68-x69
Public Const MemPos_Oper = 106                ' x6A-x6B

' --- Read-Only Byte System Variables ---

Public Const MemPos_MachType = 109            ' x6D = Constant 20 in IDRIS
Public Const MemPos_SysRel = 110              ' x6E = Constant 5 in IDRIS
Public Const MemPos_SysRev = 111              ' x6F = Constant 0 in IDRIS

' --- Flag registers ---

Public Const MemPos_F = 92                    ' x5C
Public Const MemPos_F1 = 93                   ' x5D
Public Const MemPos_F2 = 94                   ' x5E
Public Const MemPos_F3 = 95                   ' x5F
Public Const MemPos_F4 = 96                   ' x60
Public Const MemPos_F5 = 97                   ' x61
Public Const MemPos_F6 = 98                   ' x62
Public Const MemPos_F7 = 99                   ' x63
Public Const MemPos_F8 = 100                  ' x64
Public Const MemPos_F9 = 101                  ' x65
Public Const MinFLow = 0
Public Const MaxFLow = 9

Public Const MemPos_F10 = 160                 ' xA0
Public Const MemPos_F11 = 161                 ' xA1
Public Const MemPos_F12 = 162                 ' xA2
Public Const MemPos_F13 = 163                 ' xA3
Public Const MemPos_F14 = 164                 ' xA4
Public Const MemPos_F15 = 165                 ' xA5
Public Const MemPos_F16 = 166                 ' xA6
Public Const MemPos_F17 = 167                 ' xA7
Public Const MemPos_F18 = 168                 ' xA8
Public Const MemPos_F19 = 169                 ' xA9
Public Const MemPos_F20 = 170                 ' xAA
Public Const MemPos_F21 = 171                 ' xAB
Public Const MemPos_F22 = 172                 ' xAC
Public Const MemPos_F23 = 173                 ' xAD
Public Const MemPos_F24 = 174                 ' xAE
Public Const MemPos_F25 = 175                 ' xAF
Public Const MemPos_F26 = 176                 ' xB0
Public Const MemPos_F27 = 177                 ' xB1
Public Const MemPos_F28 = 178                 ' xB2
Public Const MemPos_F29 = 179                 ' xB3
Public Const MemPos_F30 = 180                 ' xB4
Public Const MemPos_F31 = 181                 ' xB5
Public Const MemPos_F32 = 182                 ' xB6
Public Const MemPos_F33 = 183                 ' xB7
Public Const MemPos_F34 = 184                 ' xB8
Public Const MemPos_F35 = 185                 ' xB9
Public Const MemPos_F36 = 186                 ' xBA
Public Const MemPos_F37 = 187                 ' xBB
Public Const MemPos_F38 = 188                 ' xBC
Public Const MemPos_F39 = 189                 ' xBD
Public Const MemPos_F40 = 190                 ' xBE
Public Const MemPos_F41 = 191                 ' xBF
Public Const MemPos_F42 = 192                 ' xC0
Public Const MemPos_F43 = 193                 ' xC1
Public Const MemPos_F44 = 194                 ' xC2
Public Const MemPos_F45 = 195                 ' xC3
Public Const MemPos_F46 = 196                 ' xC4
Public Const MemPos_F47 = 197                 ' xC5
Public Const MemPos_F48 = 198                 ' xC6
Public Const MemPos_F49 = 199                 ' xC7
Public Const MemPos_F50 = 200                 ' xC8
Public Const MemPos_F51 = 201                 ' xC9
Public Const MemPos_F52 = 202                 ' xCA
Public Const MemPos_F53 = 203                 ' xCB
Public Const MemPos_F54 = 204                 ' xCC
Public Const MemPos_F55 = 205                 ' xCD
Public Const MemPos_F56 = 206                 ' xCE
Public Const MemPos_F57 = 207                 ' xCF
Public Const MemPos_F58 = 208                 ' xD0
Public Const MemPos_F59 = 209                 ' xD1
Public Const MemPos_F60 = 210                 ' xD2
Public Const MemPos_F61 = 211                 ' xD3
Public Const MemPos_F62 = 212                 ' xD4
Public Const MemPos_F63 = 213                 ' xD5
Public Const MemPos_F64 = 214                 ' xD6
Public Const MemPos_F65 = 215                 ' xD7
Public Const MemPos_F66 = 216                 ' xD8
Public Const MemPos_F67 = 217                 ' xD9
Public Const MemPos_F68 = 218                 ' xDA
Public Const MemPos_F69 = 219                 ' xDB
Public Const MemPos_F70 = 220                 ' xDC
Public Const MemPos_F71 = 221                 ' xDD
Public Const MemPos_F72 = 222                 ' xDE
Public Const MemPos_F73 = 223                 ' xDF
Public Const MemPos_F74 = 224                 ' xE0
Public Const MemPos_F75 = 225                 ' xE1
Public Const MemPos_F76 = 226                 ' xE2
Public Const MemPos_F77 = 227                 ' xE3
Public Const MemPos_F78 = 228                 ' xE4
Public Const MemPos_F79 = 229                 ' xE5
Public Const MemPos_F80 = 230                 ' xE6
Public Const MemPos_F81 = 231                 ' xE7
Public Const MemPos_F82 = 232                 ' xE8
Public Const MemPos_F83 = 233                 ' xE9
Public Const MemPos_F84 = 234                 ' xEA
Public Const MemPos_F85 = 235                 ' xEB
Public Const MemPos_F86 = 236                 ' xEC
Public Const MemPos_F87 = 237                 ' xED
Public Const MemPos_F88 = 238                 ' xEE
Public Const MemPos_F89 = 239                 ' xEF
Public Const MemPos_F90 = 240                 ' xF0
Public Const MemPos_F91 = 241                 ' xF1
Public Const MemPos_F92 = 242                 ' xF2
Public Const MemPos_F93 = 243                 ' xF3
Public Const MemPos_F94 = 244                 ' xF4
Public Const MemPos_F95 = 245                 ' xF5
Public Const MemPos_F96 = 246                 ' xF6
Public Const MemPos_F97 = 247                 ' xF7
Public Const MemPos_F98 = 248                 ' xF8
Public Const MemPos_F99 = 249                 ' xF9
Public Const MinFHigh = 10
Public Const MaxFHigh = 99

' --- Local copy of Global Flag registers ---

Public Const MemPos_G = 128                   ' x80-x81
Public Const MemPos_G1 = 130                  ' x82
Public Const MemPos_G2 = 131                  ' x83
Public Const MemPos_G3 = 132                  ' x84
Public Const MemPos_G4 = 133                  ' x85
Public Const MemPos_G5 = 134                  ' x86
Public Const MemPos_G6 = 135                  ' x87
Public Const MemPos_G7 = 136                  ' x88
Public Const MemPos_G8 = 137                  ' x89
Public Const MemPos_G9 = 138                  ' x8A

' --- New Internal IDRIS Variables ---

Public Const MemPos_SortState = 139           ' x8B - New in IDRIS
Public Const MemPos_PrintDev = 140            ' x8C - New in IDRIS
Public Const MemPos_PrintOn = 141             ' x8D - New in IDRIS
Public Const MemPos_Background = 142          ' x8E - New in IDRIS
Public Const MemPos_TBAlloc = 143             ' x8F - New in IDRIS
Public Const MemPos_FFPending = 144           ' x90 - New in IDRIS
Public Const MemPos_PageHasData = 145         ' x91 - New in IDRIS
Public Const MemPos_LineHasData = 146         ' x92 - New in IDRIS
Public Const MemPos_ILMFlag = 147             ' x93 - New in IDRIS
Public Const MemPos_LocalEdit = 148           ' x94 - New in IDRIS
Public Const MemPos_ScriptRunFlag = 149       ' x95 - New in IDRIS
Public Const MemPos_ScriptWriteFlag = 150     ' x96 - New in IDRIS

' --- Numeric 6-byte Variables ---

Public Const NumSlotSize = 6

Public Const MemPos_N = 256                   ' x0100
Public Const MinNLow = 0
Public Const MaxNLow = 63

Public Const MemPos_Rec = MemPos_N + ((MaxNLow + 1) * NumSlotSize)
Public Const MemPos_RemVal = MemPos_N + ((MaxNLow + 2) * NumSlotSize)

Public Const MemPos_N64 = 55 * 256            ' x3700
Public Const MinNHigh = 64
Public Const MaxNHigh = 99

' --- Low Alpha Varables ---

Public Const MemPos_DateVal = 746             ' x02EA
Public Const MemPos_Key = MemPos_DateVal + 18 ' x02FC
Public Const MemPos_A = MemPos_Key + 20       ' x0310
Public Const MemPos_A1 = MemPos_A + 40        ' x0338
Public Const MemPos_A2 = MemPos_A + 60        ' x034C
Public Const MemPos_B = MemPos_A + 80         ' x0360
Public Const MemPos_B1 = MemPos_B + 40        ' x0388
Public Const MemPos_B2 = MemPos_B + 60        ' x039C
Public Const MemPos_C = MemPos_B + 80         ' x03B0
Public Const MemPos_C1 = MemPos_C + 40        ' x03D8
Public Const MemPos_C2 = MemPos_C + 60        ' x03EC

' --- Medium Alpha Varables ---

Public Const MemPos_D = 3840                  ' x0F00
Public Const MemPos_D1 = MemPos_D + 40        ' x0F28
Public Const MemPos_D2 = MemPos_D + 60        ' x0F3C
Public Const MemPos_E = MemPos_D + 80         ' x0F50
Public Const MemPos_E1 = MemPos_E + 40        ' x0F78
Public Const MemPos_E2 = MemPos_E + 60        ' x0F8C

' --- High Alpha Varables ---

Public Const MemPos_A3 = MemPos_HighAlpha     ' x3800
Public Const MemPos_A4 = MemPos_A3 + 20       ' x3816
Public Const MemPos_A5 = MemPos_A4 + 20       ' x382C
Public Const MemPos_A6 = MemPos_A5 + 20       ' x3842
Public Const MemPos_A7 = MemPos_A6 + 20       ' x3858
Public Const MemPos_A8 = MemPos_A7 + 20       ' x386E
Public Const MemPos_A9 = MemPos_A8 + 20       ' x3884

Public Const MemPos_B3 = MemPos_A9 + 20       ' x389A
Public Const MemPos_B4 = MemPos_B3 + 20       ' x38B0
Public Const MemPos_B5 = MemPos_B4 + 20       ' x38C6
Public Const MemPos_B6 = MemPos_B5 + 20       ' x38DC
Public Const MemPos_B7 = MemPos_B6 + 20       ' x38F2
Public Const MemPos_B8 = MemPos_B7 + 20       ' x3908
Public Const MemPos_B9 = MemPos_B8 + 20       ' x391E

Public Const MemPos_C3 = MemPos_B9 + 20       ' x3934
Public Const MemPos_C4 = MemPos_C3 + 20       ' x394A
Public Const MemPos_C5 = MemPos_C4 + 20       ' x3960
Public Const MemPos_C6 = MemPos_C5 + 20       ' x3976
Public Const MemPos_C7 = MemPos_C6 + 20       ' x398C
Public Const MemPos_C8 = MemPos_C7 + 20       ' x39A2
Public Const MemPos_C9 = MemPos_C8 + 20       ' x39B8

Public Const MemPos_D3 = MemPos_C9 + 20       ' x39CE
Public Const MemPos_D4 = MemPos_D3 + 20       ' x39E4
Public Const MemPos_D5 = MemPos_D4 + 20       ' x39FA
Public Const MemPos_D6 = MemPos_D5 + 20       ' x3A10
Public Const MemPos_D7 = MemPos_D6 + 20       ' x3A26
Public Const MemPos_D8 = MemPos_D7 + 20       ' x3A3C
Public Const MemPos_D9 = MemPos_D8 + 20       ' x3A52

Public Const MemPos_E3 = MemPos_D9 + 20       ' x3A68
Public Const MemPos_E4 = MemPos_E3 + 20       ' x3A7E
Public Const MemPos_E5 = MemPos_E4 + 20       ' x3A94
Public Const MemPos_E6 = MemPos_E5 + 20       ' x3AAA
Public Const MemPos_E7 = MemPos_E6 + 20       ' x3AC0
Public Const MemPos_E8 = MemPos_E7 + 20       ' x3AD6
Public Const MemPos_E9 = MemPos_E8 + 20       ' x3AEC

' --- Buffers ---

Public Const MemPos_R = 4 * 256               ' x0400
Public Const MemPos_Z = 5 * 256               ' x0500
Public Const MemPos_X = 6 * 256               ' x0600
Public Const MemPos_Y = 7 * 256               ' x0700
Public Const MemPos_W = 8 * 256               ' x0800
Public Const MemPos_S = 11 * 256              ' x0B00
Public Const MemPos_T = 12 * 256              ' x0C00
Public Const MemPos_U = 13 * 256              ' x0D00
Public Const MemPos_V = 14 * 256              ' x0E00
