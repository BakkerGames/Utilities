VERSION 5.00
Begin VB.Form rtDebugLog 
   Caption         =   "Debug Log"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4455
   ScaleMode       =   0  'User
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8235
   End
End
Attribute VB_Name = "rtDebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------
' --- frmDebugLog - 10/03/2005 ---
' --------------------------------

Option Explicit

' -----------------------------------------------------------------------------
' 10/03/2005 - Changed SendMessage to type-safe SendMessageBynum.
' -----------------------------------------------------------------------------

Private Sub Form_Load()
   rtDebugLogLoaded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      Me.Hide
      Cancel = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   rtDebugLogLoaded = False
End Sub

Private Sub Form_Resize()
   If Me.WindowState <> vbMinimized Then
      txtLog.Width = Me.ScaleWidth
      txtLog.Height = Me.ScaleHeight
   End If
End Sub

Public Sub AddMessage(ByVal Value As String)
   Dim lngPos As Long
   ' ----------------
   If Not Me.Visible Then
      Me.Show
   End If
   SendMessageBynum txtLog.hWnd, WM_SETREDRAW, False, 0
   If (Len(Value) + 2 < 10000) And (Len(txtLog.Text) + Len(Value) + 2 > 10000) Then
      Do While Len(txtLog.Text) + Len(Value) + 2 > 10000
         lngPos = InStr(txtLog.Text, vbCrLf)
         txtLog.Text = Mid$(txtLog.Text, lngPos + 2)
      Loop
   End If
   txtLog.Text = txtLog.Text & Format(Now, "yyyy.mm.dd hh:mm:ss") & " - " & Value & vbCrLf
   SendMessageBynum txtLog.hWnd, WM_SETREDRAW, True, 0
   txtLog.SelStart = Len(txtLog.Text)
End Sub
