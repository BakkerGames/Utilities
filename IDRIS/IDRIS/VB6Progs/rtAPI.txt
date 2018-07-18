Attribute VB_Name = "rtAPI"
' ---------------------------
' --- modAPI - 10/03/2005 ---
' ---------------------------

' ----------------------------------------------
' --- This contains useful Windows API calls ---
' ----------------------------------------------

' -----------------------------------------------------------------------------
' 10/03/2005 - Changed SendMessage to type-safe SendMessageBynum.
' -----------------------------------------------------------------------------

Option Explicit

' ------------------------------------------------
' --- Sleep command - put this thread to sleep ---
' ------------------------------------------------

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ------------------------
' --- Process commands ---
' ------------------------

Private Const SYNCHRONIZE As Long = &H100000
Private Const INFINITE As Long = &HFFFFFFFF ' infinite timeout

Private Declare Function OpenProcess Lib "kernel32" _
                                    (ByVal dwDesiredAccess As Long, _
                                     ByVal bInheritHandle As Long, _
                                     ByVal dwProcessId As Long) As Long
                                                     
Private Declare Function WaitForSingleObject Lib "kernel32" _
                                    (ByVal hHandle As Long, _
                                     ByVal dwMilliseconds As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" _
                                    (ByVal hObject As Long) As Long

' ----------------------
' --- Shell commands ---
' ----------------------

Public Const sw_ShowNormal As Long = 1
Public Const sw_ShowMaximized As Long = 3
Public Const sw_ShowDefault As Long = 10

Public Declare Function ShellExecute Lib "shell32.dll" _
                        Alias "ShellExecuteA" _
                       (ByVal hWnd As Long, _
                        ByVal lpOperation As String, _
                        ByVal lpFile As String, _
                        ByVal lpParameters As String, _
                        ByVal lpDirectory As String, _
                        ByVal nShowCmd As Long) _
                        As Long

' -----------------------
' --- File operations ---
' -----------------------

Private Const FILE_BEGIN = 0
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1
Private Const GENERIC_WRITE = &H40000000
Private Const MAX_PATH = 1024

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
                                   (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As Any, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" _
                                       (ByVal hFile As Long, _
                                        ByVal lDistanceToMove As Long, _
                                        lpDistanceToMoveHigh As Long, _
                                        ByVal dwMoveMethod As Long) As Long
                                        
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function GetTempFileName Lib "kernel32" _
                                   Alias "GetTempFileNameA" _
                                   (ByVal lpszPath As String, _
                                    ByVal lpPrefixString As String, _
                                    ByVal wUnique As Long, _
                                    ByVal lpTempFileName As String) As Long
' ------------------------
' --- Message routines ---
' ------------------------

Public Const WM_SETREDRAW = &HB

Public Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" _
               (ByVal hWnd As Long, _
               ByVal wMsg As Long, _
               ByVal wParam As Long, _
               lParam As Long) _
               As Long

' ------------------------
' --- Printer routines ---
' ------------------------

Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" _
                                     (ByVal pName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" _
                                     (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      ByVal pDefault As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" _
                                     (ByVal hPrinter As Long, _
                                      ByVal Level As Long, _
                                      pDocInfo As DOCINFO) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" _
                                     (ByVal hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" _
                                     (ByVal hPrinter As Long, _
                                      pBuf As Any, _
                                      ByVal cdBuf As Long, _
                                      pcWritten As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" _
                                     (ByVal hPrinter As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" _
                                     (ByVal hPrinter As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" _
                                     (ByVal hPrinter As Long) As Long

' ---------------------------
' --- Wrapper subroutines ---
' ---------------------------

Public Sub WaitForTerminate(ByVal PID As Long)
   Dim PHnd As Long
   ' --------------
   PHnd = OpenProcess(SYNCHRONIZE, 0, PID)
   If PHnd <> 0 Then
      WaitForSingleObject PHnd, INFINITE
      CloseHandle PHnd
   End If
End Sub

Public Sub SetFileSize(ByVal FileName As String, ByVal newSize As Long)
   
   ' Extend or trim a file to a given length.
   ' If the file is extended, the added bytes are undefined
    
   Dim fileHandle As Long
   
   ' open the file, get the handle
   fileHandle = CreateFile(FileName, GENERIC_WRITE, 0&, ByVal 0&, OPEN_EXISTING, 0&, 0&)
   
   ' raise error if not found
   If fileHandle = INVALID_HANDLE_VALUE Then
      Err.Raise 53     ' This is "file not found"
   End If
   
   ' move the file pointer to new position, raise error if fails
   If SetFilePointer(fileHandle, newSize, 0&, FILE_BEGIN) = -1 Then
      CloseHandle fileHandle
      Err.Raise 5     ' this is "illegal function call"
   End If
   
   ' attempt to set the end of file, raise error
   If SetEndOfFile(fileHandle) = 0 Then
      CloseHandle fileHandle
      Err.Raise 5     ' this is "illegal function call"
   End If
   
   ' close the file and exit
   CloseHandle fileHandle
   
End Sub

Public Function GetTempFile(ByVal Path As String, ByVal Prefix As String) As String
   Dim lngResult As Long
   Dim strResult As String
   ' ---------------------
   strResult = Space(MAX_PATH)
   lngResult = GetTempFileName(Path, Prefix, 0, strResult)
   If lngResult = 0 Then
      GetTempFile = "" ' error
   Else
      GetTempFile = Left$(strResult, InStr(strResult, Chr$(0)) - 1)
   End If
End Function
