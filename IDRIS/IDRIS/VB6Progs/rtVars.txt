Attribute VB_Name = "rtVars"
' ---------------------------
' --- rtVars - 10/01/2008 ---
' ---------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 10/01/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
' 02/06/2006 - Added Debug variables.
' 01/20/2006 - Added N64-N99 variables.
' 01/18/2006 - Added AltCommonFilePath variable to hold an alternate (and
'              hopefully existing) directory if the CommonFilePath isn't found.
' ------------------------------------------------------------------------------

Public Const APP_NAME = "IDRIS"

' --- CustomWindow Result ---

Public CustomWindowResult As String
Public CustomWindowProcessDone As Boolean

' --- flags to change operation of runtime ---

Public Const CompressMultiFF As Boolean = True ' true to prevent extra blank pages
Public Const AllowLocalEdit As Boolean = True  ' true to allow edit in Client program
Public Const ReadKeyCacheSize As Long = 100    ' used to fine-tune performance

' --- memory array ---

Public MEM(TotalMemSize - 1) As Byte

' --- numeric variables ---

Public N As Currency
Public N1 As Currency
Public N2 As Currency
Public N3 As Currency
Public N4 As Currency
Public N5 As Currency
Public N6 As Currency
Public N7 As Currency
Public N8 As Currency
Public N9 As Currency
Public N10 As Currency
Public N11 As Currency
Public N12 As Currency
Public N13 As Currency
Public N14 As Currency
Public N15 As Currency
Public N16 As Currency
Public N17 As Currency
Public N18 As Currency
Public N19 As Currency
Public N20 As Currency
Public N21 As Currency
Public N22 As Currency
Public N23 As Currency
Public N24 As Currency
Public N25 As Currency
Public N26 As Currency
Public N27 As Currency
Public N28 As Currency
Public N29 As Currency
Public N30 As Currency
Public N31 As Currency
Public N32 As Currency
Public N33 As Currency
Public N34 As Currency
Public N35 As Currency
Public N36 As Currency
Public N37 As Currency
Public N38 As Currency
Public N39 As Currency
Public N40 As Currency
Public N41 As Currency
Public N42 As Currency
Public N43 As Currency
Public N44 As Currency
Public N45 As Currency
Public N46 As Currency
Public N47 As Currency
Public N48 As Currency
Public N49 As Currency
Public N50 As Currency
Public N51 As Currency
Public N52 As Currency
Public N53 As Currency
Public N54 As Currency
Public N55 As Currency
Public N56 As Currency
Public N57 As Currency
Public N58 As Currency
Public N59 As Currency
Public N60 As Currency
Public N61 As Currency
Public N62 As Currency
Public N63 As Currency

Public REC As Long

Public REMVAL As Currency

' --- high numeric variables ---

Public N64 As Currency
Public N65 As Currency
Public N66 As Currency
Public N67 As Currency
Public N68 As Currency
Public N69 As Currency
Public N70 As Currency
Public N71 As Currency
Public N72 As Currency
Public N73 As Currency
Public N74 As Currency
Public N75 As Currency
Public N76 As Currency
Public N77 As Currency
Public N78 As Currency
Public N79 As Currency
Public N80 As Currency
Public N81 As Currency
Public N82 As Currency
Public N83 As Currency
Public N84 As Currency
Public N85 As Currency
Public N86 As Currency
Public N87 As Currency
Public N88 As Currency
Public N89 As Currency
Public N90 As Currency
Public N91 As Currency
Public N92 As Currency
Public N93 As Currency
Public N94 As Currency
Public N95 As Currency
Public N96 As Currency
Public N97 As Currency
Public N98 As Currency
Public N99 As Currency

' --- special variables used inside complicated commands ---

Public EXITING As Boolean ' exiting entire application
Public MUSTEXIT As Boolean ' used in ESC and CAN commands
Public SWITCHING As Boolean ' when exiting this runtime
Public WAITTOEXIT As Boolean ' when exiting all runtimes
Public ERRORTHROWN As Boolean ' prevent multiple errors

Public NUMERIC_RESULT As Currency
Public ALPHA_RESULT As String
Public UPDATE_VALUE As Currency
Public FREEZE_LENGTH As Boolean

' --- Cadol Constants ---

Public Const FALSEVAL = 0
Public Const TRUEVAL = 1

Public Const GOSUB_TYPEVAL = 0
Public Const WHEN_CANCEL_TYPEVAL = 1
Public Const WHEN_ESCAPE_TYPEVAL = 2
Public Const WHEN_ERROR_TYPEVAL = 3

' --- Must match IDRISClient!! ---
' --- special data embedded in the keyboard stream ---
Public Const SpecChar_LocalEditOn As Byte = 1
Public Const SpecChar_LocalEditOff As Byte = 2
Public Const SpecChar_ScriptRunOn As Byte = 3
Public Const SpecChar_ScriptRunOff As Byte = 4
Public Const SpecChar_ScriptWriteOn As Byte = 5
Public Const SpecChar_ScriptWriteOff As Byte = 6

' --- Internal resources ---

Public LoginID As String
Public MachineName As String

Public CurrDevNum As Long
Public CurrVolName As String
Public CurrLibName As String
Public CurrJumpPoint As Long
Public CurrIP As String

Public GosubStack As Collection
Public KBuff As rtKbdBuffer

Public DebugFlag As Boolean
Public DebugFlagLevel As Long
Public DebugLogFile As Long

Public Last_DoEvents As Single
Public CheckDoEventsCount As Long

Public InBreakMode As Boolean
Public BreakFilename As String
Public LastILCode As String
Public ProgILCode As String
Public DebugOneStep As Boolean
Public DebugStepOver As Boolean
Public Breakpoints As Collection
Public Watchpoints As Collection

Public ShortClientCompare As String
Public LongClientCompare As String

' --- Form loaded flags ---

Public rtFormMainLoaded As Boolean
Public rtDebugLogLoaded As Boolean

' --- Printing variables ---

Public PrinterFileNum As Long
Public PrinterFileName As String
Public PrinterType As String
Public PrinterDeviceName As String
Public PrinterParameters As String
Public PrinterHandle As Long
Public PrinterJobNum As Long

' --- Sorting Variables ---

' --- if MAXMEMTAGS is > 999 change SaveMemory and LoadMemory using 3 chars ---
Public Const MAXMEMTAGS As Long = 255 ' max tags it can sort in memory

Public SortFileNum As Long
Public SortFileName As String
Public SortTagSize As Long
Public SortLineCount As Long
Public FetchLineCount As Long

Public SortTags(MAXMEMTAGS) As String ' includes record number (z14)
Public SortIndex(MAXMEMTAGS) As Long

' --- Channel Variables ---

Public ChannelPaths(MaxChannel) As String
Public ChannelFileNums(MaxChannel) As Long

' --- SQL Connection variables ---

Public cnSQL As ADODB.Connection

' --- Server connection variables ---

Public HostIP As String
Public PortVal As Long
Public BackgroundIP As String
Public BackgroundPort As Long

Public ClientList As String
Public ReadOnly As Boolean

Public ReadyToRun As Boolean
Public PendingInput As String
Public PendingOutput As String

Public SpawnTarget As String

Public ServerSendComplete As Boolean

' --- File access variables ---

Public LockFlag As Boolean
Public HasLockedRec As Boolean
Public LockedFileNum As Long
Public LockedSQLTable As String
Public LockedRecNum As Long
Public LockedRecLen As Long
Public LockedResource As String
Public LockedCadolXref As rtCadolXref
Public rsLockedRec As ADODB.Recordset
Public rsLockedResult As ADODB.Recordset
Public SQLSubQuery As String
Public SQLSubQueryFile As String
Public Files(MaxFile) As rtCadolFile
Public UpdateRecField As Boolean
Public rsRecord As ADODB.Recordset

' --- Global Registers ---

Public rsGRegs As ADODB.Recordset

' --- External variables from INI file ---

Public IniFilename As String
Public EnvName As String
Public BinPath As String
Public FileServerPath As String
Public LibraryPath As String
Public TempPath As String
Public UDLFilename As String
Public ErrorFilename As String
Public CommonFilePath As String
Public AltCommonFilePath As String
