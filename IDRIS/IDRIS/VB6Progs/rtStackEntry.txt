VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtStackEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------
' --- rtStackEntry - 02/08/2006 ---
' ---------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 02/08/2006 - Added error checking to FromString.
' ------------------------------------------------------------------------------

Public ItemType As Long
Public DevNum As Long
Public VolName As String
Public LibName As String
Public ProgNum As Long
Public JumpNum As Long

Public Function ToString() As String
   Dim strResult As String
   ' ---------------------
   strResult = ""
   Select Case ItemType
      Case WHEN_CANCEL_TYPEVAL
         strResult = "C:"
      Case WHEN_ESCAPE_TYPEVAL
         strResult = "E:"
      Case WHEN_ERROR_TYPEVAL
         strResult = "R:"
   End Select
   strResult = strResult & _
               Trim$(Str$(DevNum)) & ":" & _
               VolName & ":" & _
               LibName & ":" & _
               Trim$(Str$(ProgNum)) & ":" & _
               Trim$(Str$(JumpNum))
   ToString = strResult
End Function

Public Sub FromString(ByVal Value As String)
   Dim Tokens() As String
   ' --------------------
   ItemType = 0
   DevNum = 0
   VolName = ""
   LibName = ""
   ProgNum = 0
   JumpNum = 0
   Tokens = Split(Value, ":")
   If UBound(Tokens) < 3 Then Exit Sub
   If Tokens(0) = "C" Or Tokens(0) = "E" Or Tokens(0) = "R" Then
      If UBound(Tokens) < 4 Then Exit Sub
      Select Case Tokens(0)
         Case "C"
            ItemType = WHEN_CANCEL_TYPEVAL
         Case "E"
            ItemType = WHEN_ESCAPE_TYPEVAL
         Case "R"
            ItemType = WHEN_ERROR_TYPEVAL
      End Select
      DevNum = Val(Tokens(1))
      VolName = Tokens(2)
      LibName = Tokens(3)
      ProgNum = Val(Tokens(4))
      If UBound(Tokens) >= 5 Then
         JumpNum = Val(Tokens(5))
      End If
   Else
      ItemType = GOSUB_TYPEVAL
      DevNum = Val(Tokens(0))
      VolName = Tokens(1)
      LibName = Tokens(2)
      ProgNum = Val(Tokens(3))
      If UBound(Tokens) >= 4 Then
         JumpNum = Val(Tokens(4))
      End If
   End If
End Sub
