VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtCadolXref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------
' --- rtCadolXref - 02/28/2005 ---
' --------------------------------

Option Explicit

Public SQLTableName As String
Public CadolKey As String
Public CadolByte As Long
Public CadolLength As Long
Public CadolValue As Long
Public Multiple As Boolean
Public DataFormats As Collection
