VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtKbdBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------
' --- rtKbdBuffer - 09/16/2008 ---
' --------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 09/16/2008 - SBAKKER - URD 11164
'            - Made changes recommended by CodeAdvisor.
' ------------------------------------------------------------------------------

' --------------------------
' --- internal variables ---
' --------------------------

Private Chars As Collection

' ----------------------
' --- class routines ---
' ----------------------

Private Sub Class_Initialize()
   Set Chars = New Collection
End Sub

Private Sub Class_Terminate()
   Do While Chars.Count > 0
      Chars.Remove 1
   Loop
   Set Chars = Nothing
End Sub

' -----------------------
' --- public routines ---
' -----------------------

Public Sub Add(ByVal Value As Byte)
   Chars.Add Value
End Sub

Public Sub Clear()
   Do While Chars.Count > 0
      Chars.Remove 1
   Loop
End Sub

Public Sub ClearNum(ByVal Number As Integer)
   Dim LoopNum As Integer
   ' --------------------
   For LoopNum = 1 To Number
      If Chars.Count = 0 Then
         ThrowError "rtKbdBuffer.ClearNum", "Keyboard buffer underflow"
         Exit Sub
      End If
      Chars.Remove 1
   Next LoopNum
End Sub

Public Function Count() As Long
   Count = Chars.Count
End Function

Public Function GetChar() As Byte
   If Chars.Count = 0 Then
      ThrowError "rtKbdBuffer.GetChar", "Keyboard buffer underflow"
      Exit Function
   End If
   GetChar = CByte(Chars.Item(1))
   Chars.Remove 1
End Function

Public Function Peek() As Byte
   ' --- peek may be used even if no chars, returns 255 ---
   If Chars.Count = 0 Then
      Peek = 255
   Else
      Peek = CByte(Chars.Item(1))
   End If
End Function

Public Function ToString() As String
   Dim lngLoop As Long
   Dim strResult As String
   ' ---------------------
   strResult = ""
   For lngLoop = 1 To Chars.Count
      strResult = strResult & Right$("00" & Hex$(Chars.Item(lngLoop)), 2)
   Next lngLoop
   ToString = strResult
End Function

Public Sub FromString(ByVal Value As String)
   Dim bTemp As Byte
   Dim lngLoop As Long
   ' -----------------
   Do While Chars.Count > 0
      Chars.Remove 1
   Loop
   For lngLoop = 1 To Len(Value) Step 2
      bTemp = Val("&H" & Mid$(Value, lngLoop, 2))
      Chars.Add bTemp
   Next
End Sub
