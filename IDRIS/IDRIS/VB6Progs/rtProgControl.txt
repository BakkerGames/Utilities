Attribute VB_Name = "rtProgControl"
' ----------------------------------
' --- rtProgControl - 10/13/2008 ---
' ----------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 10/13/2008 - SBAKKER - URD 11164
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
' 01/30/2006 - Removed "DebugMessage '*** SWITCHING...'" messages. The problem
'              these were tracking has been corrected and is not needed anymore.
' ------------------------------------------------------------------------------

Public Sub ProgControl()
   Dim lngProgNum As Long
   Dim lngJumpNum As Long
   Dim objItem As rtStackEntry
   ' -------------------------
   Do While Not EXITING
      CheckDoEvents
      If GosubStack.Count = 0 Then
         ExecuteEscape
         If EXITING Then Exit Sub
         lngProgNum = 0
         lngJumpNum = 0
         GoTo SelectProg
      End If
      ' --- get next gosub entry ---
      Set objItem = GosubStack.Item(GosubStack.Count)
      GosubStack.Remove GosubStack.Count
      ' --- ignore "WHEN" entries on gosub stack, clear variables ---
      If objItem.ItemType = WHEN_CANCEL_TYPEVAL Then
         LET_CANVAL 0 ' normal cancel processing
         GoTo NextGosubEntry
      End If
      If objItem.ItemType = WHEN_ESCAPE_TYPEVAL Then
         LET_ESCVAL 0 ' normal escape processing
         GoTo NextGosubEntry
      End If
      If objItem.ItemType = WHEN_ERROR_TYPEVAL Then
         GoTo NextGosubEntry
      End If
      ' --- check if jumping to another library ---
      If objItem.DevNum <> CurrDevNum Or _
            UCase$(objItem.VolName) <> UCase$(CurrVolName) Or _
            UCase$(objItem.LibName) <> UCase$(CurrLibName) Then
         If Not MEMTF(MemPos_Background) Then
            SpawnTarget = objItem.ToString ' save for later
            ' --- must set value before sending message ---
            SWITCHING = True
            ' --- begin switching protocol ---
            DebugMessage "Switching to: " & SpawnTarget
            SendToServer "APPLICATION" & vbTab & "SWITCHING" & vbTab & SpawnTarget
            DoEvents
         End If
         MUSTEXIT = True
         EXITING = True
         DebugMessage "Exiting from ProgControl"
         Exit Sub
      End If
GetProgNum:
      ' --- get program and jump point ---
      lngProgNum = objItem.ProgNum
      lngJumpNum = objItem.JumpNum
      Set objItem = Nothing
      ' --- check for program 0, jumppoint 0 ---
      If lngProgNum = 0 And lngJumpNum = 0 Then
         ExecuteEscape
         If EXITING Then Exit Sub
      End If
SelectProg:
      ' --- convert to proper range of values ---
      If lngProgNum > 255 Then
         lngProgNum = ModPos(lngProgNum, 256)
      End If
      ' --- store in PROG variable ---
      LET_PROG lngProgNum
      CurrJumpPoint = lngJumpNum
      ' --- show incoming messages ---
      If DebugFlag And DebugFlagLevel > 1 Then
         DebugMessage "CALL: PROG=" & Trim$(Str$(lngProgNum)) & ", JUMP=" & Trim$(Str$(lngJumpNum))
      End If
      ' --- select routine ---
      Select Case lngProgNum
         ' --- normal library routines ---
         Case 0: Prog_000 lngJumpNum
         Case 1: Prog_001 lngJumpNum
         Case 2: Prog_002 lngJumpNum
         Case 3: Prog_003 lngJumpNum
         Case 4: Prog_004 lngJumpNum
         Case 5: Prog_005 lngJumpNum
         Case 6: Prog_006 lngJumpNum
         Case 7: Prog_007 lngJumpNum
         Case 8: Prog_008 lngJumpNum
         Case 9: Prog_009 lngJumpNum
         Case 10: Prog_010 lngJumpNum
         Case 11: Prog_011 lngJumpNum
         Case 12: Prog_012 lngJumpNum
         Case 13: Prog_013 lngJumpNum
         Case 14: Prog_014 lngJumpNum
         Case 15: Prog_015 lngJumpNum
         Case 16: Prog_016 lngJumpNum
         Case 17: Prog_017 lngJumpNum
         Case 18: Prog_018 lngJumpNum
         Case 19: Prog_019 lngJumpNum
         Case 20: Prog_020 lngJumpNum
         Case 21: Prog_021 lngJumpNum
         Case 22: Prog_022 lngJumpNum
         Case 23: Prog_023 lngJumpNum
         Case 24: Prog_024 lngJumpNum
         Case 25: Prog_025 lngJumpNum
         Case 26: Prog_026 lngJumpNum
         Case 27: Prog_027 lngJumpNum
         Case 28: Prog_028 lngJumpNum
         Case 29: Prog_029 lngJumpNum
         Case 30: Prog_030 lngJumpNum
         Case 31: Prog_031 lngJumpNum
         Case 32: Prog_032 lngJumpNum
         Case 33: Prog_033 lngJumpNum
         Case 34: Prog_034 lngJumpNum
         Case 35: Prog_035 lngJumpNum
         Case 36: Prog_036 lngJumpNum
         Case 37: Prog_037 lngJumpNum
         Case 38: Prog_038 lngJumpNum
         Case 39: Prog_039 lngJumpNum
         Case 40: Prog_040 lngJumpNum
         Case 41: Prog_041 lngJumpNum
         Case 42: Prog_042 lngJumpNum
         Case 43: Prog_043 lngJumpNum
         Case 44: Prog_044 lngJumpNum
         Case 45: Prog_045 lngJumpNum
         Case 46: Prog_046 lngJumpNum
         Case 47: Prog_047 lngJumpNum
         Case 48: Prog_048 lngJumpNum
         Case 49: Prog_049 lngJumpNum
         Case 50: Prog_050 lngJumpNum
         Case 51: Prog_051 lngJumpNum
         Case 52: Prog_052 lngJumpNum
         Case 53: Prog_053 lngJumpNum
         Case 54: Prog_054 lngJumpNum
         Case 55: Prog_055 lngJumpNum
         Case 56: Prog_056 lngJumpNum
         Case 57: Prog_057 lngJumpNum
         Case 58: Prog_058 lngJumpNum
         Case 59: Prog_059 lngJumpNum
         Case 60: Prog_060 lngJumpNum
         Case 61: Prog_061 lngJumpNum
         Case 62: Prog_062 lngJumpNum
         Case 63: Prog_063 lngJumpNum
         Case 64: Prog_064 lngJumpNum
         Case 65: Prog_065 lngJumpNum
         Case 66: Prog_066 lngJumpNum
         Case 67: Prog_067 lngJumpNum
         Case 68: Prog_068 lngJumpNum
         Case 69: Prog_069 lngJumpNum
         Case 70: Prog_070 lngJumpNum
         Case 71: Prog_071 lngJumpNum
         Case 72: Prog_072 lngJumpNum
         Case 73: Prog_073 lngJumpNum
         Case 74: Prog_074 lngJumpNum
         Case 75: Prog_075 lngJumpNum
         Case 76: Prog_076 lngJumpNum
         Case 77: Prog_077 lngJumpNum
         Case 78: Prog_078 lngJumpNum
         Case 79: Prog_079 lngJumpNum
         Case 80: Prog_080 lngJumpNum
         Case 81: Prog_081 lngJumpNum
         Case 82: Prog_082 lngJumpNum
         Case 83: Prog_083 lngJumpNum
         Case 84: Prog_084 lngJumpNum
         Case 85: Prog_085 lngJumpNum
         Case 86: Prog_086 lngJumpNum
         Case 87: Prog_087 lngJumpNum
         Case 88: Prog_088 lngJumpNum
         Case 89: Prog_089 lngJumpNum
         Case 90: Prog_090 lngJumpNum
         Case 91: Prog_091 lngJumpNum
         Case 92: Prog_092 lngJumpNum
         Case 93: Prog_093 lngJumpNum
         Case 94: Prog_094 lngJumpNum
         Case 95: Prog_095 lngJumpNum
         Case 96: Prog_096 lngJumpNum
         Case 97: Prog_097 lngJumpNum
         Case 98: Prog_098 lngJumpNum
         Case 99: Prog_099 lngJumpNum
         Case 100: Prog_100 lngJumpNum
         Case 101: Prog_101 lngJumpNum
         Case 102: Prog_102 lngJumpNum
         Case 103: Prog_103 lngJumpNum
         Case 104: Prog_104 lngJumpNum
         Case 105: Prog_105 lngJumpNum
         Case 106: Prog_106 lngJumpNum
         Case 107: Prog_107 lngJumpNum
         Case 108: Prog_108 lngJumpNum
         Case 109: Prog_109 lngJumpNum
         Case 110: Prog_110 lngJumpNum
         Case 111: Prog_111 lngJumpNum
         Case 112: Prog_112 lngJumpNum
         Case 113: Prog_113 lngJumpNum
         Case 114: Prog_114 lngJumpNum
         Case 115: Prog_115 lngJumpNum
         Case 116: Prog_116 lngJumpNum
         Case 117: Prog_117 lngJumpNum
         Case 118: Prog_118 lngJumpNum
         Case 119: Prog_119 lngJumpNum
         Case 120: Prog_120 lngJumpNum
         Case 121: Prog_121 lngJumpNum
         Case 122: Prog_122 lngJumpNum
         Case 123: Prog_123 lngJumpNum
         Case 124: Prog_124 lngJumpNum
         Case 125: Prog_125 lngJumpNum
         Case 126: Prog_126 lngJumpNum
         Case 127: Prog_127 lngJumpNum
         Case 128: Prog_128 lngJumpNum
         Case 129: Prog_129 lngJumpNum
         Case 130: Prog_130 lngJumpNum
         Case 131: Prog_131 lngJumpNum
         Case 132: Prog_132 lngJumpNum
         Case 133: Prog_133 lngJumpNum
         Case 134: Prog_134 lngJumpNum
         Case 135: Prog_135 lngJumpNum
         Case 136: Prog_136 lngJumpNum
         Case 137: Prog_137 lngJumpNum
         Case 138: Prog_138 lngJumpNum
         Case 139: Prog_139 lngJumpNum
         Case 140: Prog_140 lngJumpNum
         Case 141: Prog_141 lngJumpNum
         Case 142: Prog_142 lngJumpNum
         Case 143: Prog_143 lngJumpNum
         Case 144: Prog_144 lngJumpNum
         Case 145: Prog_145 lngJumpNum
         Case 146: Prog_146 lngJumpNum
         Case 147: Prog_147 lngJumpNum
         Case 148: Prog_148 lngJumpNum
         Case 149: Prog_149 lngJumpNum
         Case 150: Prog_150 lngJumpNum
         Case 151: Prog_151 lngJumpNum
         Case 152: Prog_152 lngJumpNum
         Case 153: Prog_153 lngJumpNum
         Case 154: Prog_154 lngJumpNum
         Case 155: Prog_155 lngJumpNum
         Case 156: Prog_156 lngJumpNum
         Case 157: Prog_157 lngJumpNum
         Case 158: Prog_158 lngJumpNum
         Case 159: Prog_159 lngJumpNum
         Case 160: Prog_160 lngJumpNum
         Case 161: Prog_161 lngJumpNum
         Case 162: Prog_162 lngJumpNum
         Case 163: Prog_163 lngJumpNum
         Case 164: Prog_164 lngJumpNum
         Case 165: Prog_165 lngJumpNum
         Case 166: Prog_166 lngJumpNum
         Case 167: Prog_167 lngJumpNum
         Case 168: Prog_168 lngJumpNum
         Case 169: Prog_169 lngJumpNum
         Case 170: Prog_170 lngJumpNum
         Case 171: Prog_171 lngJumpNum
         Case 172: Prog_172 lngJumpNum
         Case 173: Prog_173 lngJumpNum
         Case 174: Prog_174 lngJumpNum
         Case 175: Prog_175 lngJumpNum
         Case 176: Prog_176 lngJumpNum
         Case 177: Prog_177 lngJumpNum
         Case 178: Prog_178 lngJumpNum
         Case 179: Prog_179 lngJumpNum
         Case 180: Prog_180 lngJumpNum
         Case 181: Prog_181 lngJumpNum
         Case 182: Prog_182 lngJumpNum
         Case 183: Prog_183 lngJumpNum
         Case 184: Prog_184 lngJumpNum
         Case 185: Prog_185 lngJumpNum
         Case 186: Prog_186 lngJumpNum
         Case 187: Prog_187 lngJumpNum
         Case 188: Prog_188 lngJumpNum
         Case 189: Prog_189 lngJumpNum
         Case 190: Prog_190 lngJumpNum
         Case 191: Prog_191 lngJumpNum
         Case 192: Prog_192 lngJumpNum
         Case 193: Prog_193 lngJumpNum
         Case 194: Prog_194 lngJumpNum
         Case 195: Prog_195 lngJumpNum
         Case 196: Prog_196 lngJumpNum
         Case 197: Prog_197 lngJumpNum
         Case 198: Prog_198 lngJumpNum
         Case 199: Prog_199 lngJumpNum
         Case 200: Prog_200 lngJumpNum
         Case 201: Prog_201 lngJumpNum
         Case 202: Prog_202 lngJumpNum
         Case 203: Prog_203 lngJumpNum
         Case 204: Prog_204 lngJumpNum
         Case 205: Prog_205 lngJumpNum
         Case 206: Prog_206 lngJumpNum
         Case 207: Prog_207 lngJumpNum
         Case 208: Prog_208 lngJumpNum
         Case 209: Prog_209 lngJumpNum
         Case 210: Prog_210 lngJumpNum
         Case 211: Prog_211 lngJumpNum
         Case 212: Prog_212 lngJumpNum
         Case 213: Prog_213 lngJumpNum
         Case 214: Prog_214 lngJumpNum
         Case 215: Prog_215 lngJumpNum
         Case 216: Prog_216 lngJumpNum
         Case 217: Prog_217 lngJumpNum
         Case 218: Prog_218 lngJumpNum
         Case 219: Prog_219 lngJumpNum
         Case 220: Prog_220 lngJumpNum
         Case 221: Prog_221 lngJumpNum
         Case 222: Prog_222 lngJumpNum
         Case 223: Prog_223 lngJumpNum
         Case 224: Prog_224 lngJumpNum
         Case 225: Prog_225 lngJumpNum
         Case 226: Prog_226 lngJumpNum
         Case 227: Prog_227 lngJumpNum
         Case 228: Prog_228 lngJumpNum
         Case 229: Prog_229 lngJumpNum
         Case 230: Prog_230 lngJumpNum
         Case 231: Prog_231 lngJumpNum
         Case 232: Prog_232 lngJumpNum
         Case 233: Prog_233 lngJumpNum
         Case 234: Prog_234 lngJumpNum
         Case 235: Prog_235 lngJumpNum
         Case 236: Prog_236 lngJumpNum
         Case 237: Prog_237 lngJumpNum
         Case 238: Prog_238 lngJumpNum
         Case 239: Prog_239 lngJumpNum
         Case 240: Prog_240 lngJumpNum
         Case 241: Prog_241 lngJumpNum
         Case 242: Prog_242 lngJumpNum
         Case 243: Prog_243 lngJumpNum
         Case 244: Prog_244 lngJumpNum
         Case 245: Prog_245 lngJumpNum
         Case 246: Prog_246 lngJumpNum
         Case 247: Prog_247 lngJumpNum
         Case 248: Prog_248 lngJumpNum
         Case 249: Prog_249 lngJumpNum
         Case 250: Prog_250 lngJumpNum
         Case 251: Prog_251 lngJumpNum
         Case 252: Prog_252 lngJumpNum
         Case 253: Prog_253 lngJumpNum
         Case 254: Prog_254 lngJumpNum
         Case 255: Prog_255 lngJumpNum
         ' ------------------------------------------------
         ' --- negative /IDRISYS and /USERLIB routines  ---
         ' --- only implementented ones are listed here ---
         ' ------------------------------------------------
         ' --- /USERLIB routines ---
         Case -256: SysProg_000 lngJumpNum
         Case -255: SysProg_001 lngJumpNum
         ' --- /IDRISYS routines ---
         Case -251: SysProg_005 lngJumpNum
         Case -247: SysProg_009 lngJumpNum
         Case -246: SysProg_010 lngJumpNum
         Case -200: SysProg_056 lngJumpNum
         Case -165: SysProg_091 lngJumpNum
         Case -144: SysProg_112 lngJumpNum
         Case -115: SysProg_141 lngJumpNum
         Case -114: SysProg_142 lngJumpNum
         Case -100: SysProg_156 lngJumpNum
         Case -96: SysProg_160 lngJumpNum
         Case -95: SysProg_161 lngJumpNum
         Case -71: SysProg_185 lngJumpNum
         Case -70: SysProg_186 lngJumpNum
         Case -69: SysProg_187 lngJumpNum
         Case -61: SysProg_195 lngJumpNum
         Case -53: SysProg_203 lngJumpNum
         Case -52: SysProg_204 lngJumpNum
         Case -51: SysProg_205 lngJumpNum
         Case -44: SysProg_212 lngJumpNum
         Case -34: SysProg_222 lngJumpNum
         Case -32: SysProg_224 lngJumpNum
         Case -31: SysProg_225 lngJumpNum
         Case -30: SysProg_226 lngJumpNum
         ' --- more /USERLIB routines ---
         Case -26: SysProg_230 lngJumpNum
         Case -25: SysProg_231 lngJumpNum
         Case -24: SysProg_232 lngJumpNum
         Case -23: SysProg_233 lngJumpNum
         Case -22: SysProg_234 lngJumpNum
         Case -21: SysProg_235 lngJumpNum
         Case -20: SysProg_236 lngJumpNum
         Case -17: SysProg_239 lngJumpNum
         Case -15: SysProg_241 lngJumpNum
         Case -14: SysProg_242 lngJumpNum
         Case -13: SysProg_243 lngJumpNum
         Case -12: SysProg_244 lngJumpNum
         Case -11: SysProg_245 lngJumpNum
         ' --- more /IDRISYS routines ---
         Case -8: SysProg_248 lngJumpNum
         Case -7: SysProg_249 lngJumpNum
         Case -6: SysProg_250 lngJumpNum
         Case -5: SysProg_251 lngJumpNum
         Case -2: SysProg_254 lngJumpNum
         ' --- unknown program number ---
         Case Else
            FATALERROR "UNKNOWN PROGRAM NUMBER: " & Trim$(Str$(lngProgNum))
      End Select
NextGosubEntry:
   Loop
End Sub
