VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtCadolFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------
' --- rtCadolFile - 10/22/2004 ---
' --------------------------------

Option Explicit

Public Device As Long
Public Volume As String
Public AdjVolume As String
Public FileName As String
Public CadolXrefs As Collection
Public LastKey As String
Public LastRec As Long
Public RecSet As ADODB.Recordset
