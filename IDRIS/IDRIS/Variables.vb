' ------------------------------
' --- Variables - 07/27/2016 ---
' ------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/27/2016 - SBakker
'            - Moved Screen constants here.
' ----------------------------------------------------------------------------------------------------

Module Variables

    Public Const APP_NAME As String = "IDRIS"

    ' --- CustomWindow Result ---

    Public CustomWindowResult As String
    Public CustomWindowProcessDone As Boolean

    ' --- flags to change operation of runtime ---

    Public Const CompressMultiFF As Boolean = True ' true to prevent extra blank pages
    Public Const AllowLocalEdit As Boolean = True  ' true to allow edit in Client program
    Public Const ReadKeyCacheSize As Int64 = 100    ' used to fine-tune performance

    ' --- memory array ---

    Public MEM(TotalMemSize - 1) As Byte

    ' --- numeric variables ---

    Public N0 As Int64
    Public N1 As Int64
    Public N2 As Int64
    Public N3 As Int64
    Public N4 As Int64
    Public N5 As Int64
    Public N6 As Int64
    Public N7 As Int64
    Public N8 As Int64
    Public N9 As Int64
    Public N10 As Int64
    Public N11 As Int64
    Public N12 As Int64
    Public N13 As Int64
    Public N14 As Int64
    Public N15 As Int64
    Public N16 As Int64
    Public N17 As Int64
    Public N18 As Int64
    Public N19 As Int64
    Public N20 As Int64
    Public N21 As Int64
    Public N22 As Int64
    Public N23 As Int64
    Public N24 As Int64
    Public N25 As Int64
    Public N26 As Int64
    Public N27 As Int64
    Public N28 As Int64
    Public N29 As Int64
    Public N30 As Int64
    Public N31 As Int64
    Public N32 As Int64
    Public N33 As Int64
    Public N34 As Int64
    Public N35 As Int64
    Public N36 As Int64
    Public N37 As Int64
    Public N38 As Int64
    Public N39 As Int64
    Public N40 As Int64
    Public N41 As Int64
    Public N42 As Int64
    Public N43 As Int64
    Public N44 As Int64
    Public N45 As Int64
    Public N46 As Int64
    Public N47 As Int64
    Public N48 As Int64
    Public N49 As Int64
    Public N50 As Int64
    Public N51 As Int64
    Public N52 As Int64
    Public N53 As Int64
    Public N54 As Int64
    Public N55 As Int64
    Public N56 As Int64
    Public N57 As Int64
    Public N58 As Int64
    Public N59 As Int64
    Public N60 As Int64
    Public N61 As Int64
    Public N62 As Int64
    Public N63 As Int64

    Public REC As Int64

    Public REMVAL As Int64

    ' --- high numeric variables ---

    Public N64 As Int64
    Public N65 As Int64
    Public N66 As Int64
    Public N67 As Int64
    Public N68 As Int64
    Public N69 As Int64
    Public N70 As Int64
    Public N71 As Int64
    Public N72 As Int64
    Public N73 As Int64
    Public N74 As Int64
    Public N75 As Int64
    Public N76 As Int64
    Public N77 As Int64
    Public N78 As Int64
    Public N79 As Int64
    Public N80 As Int64
    Public N81 As Int64
    Public N82 As Int64
    Public N83 As Int64
    Public N84 As Int64
    Public N85 As Int64
    Public N86 As Int64
    Public N87 As Int64
    Public N88 As Int64
    Public N89 As Int64
    Public N90 As Int64
    Public N91 As Int64
    Public N92 As Int64
    Public N93 As Int64
    Public N94 As Int64
    Public N95 As Int64
    Public N96 As Int64
    Public N97 As Int64
    Public N98 As Int64
    Public N99 As Int64

    ' --- special variables used inside complicated commands ---

    Public EXITING As Boolean ' exiting entire application
    Public MUSTEXIT As Boolean ' used in ESC and CAN commands
    Public SWITCHING As Boolean ' when exiting this runtime
    Public WAITTOEXIT As Boolean ' when exiting all runtimes
    Public ERRORTHROWN As Boolean ' prevent multiple errors

    Public NUMERIC_RESULT As Int64
    Public ALPHA_RESULT As String
    Public UPDATE_VALUE As Int64
    Public FREEZE_LENGTH As Boolean

    ' --- Cadol Constants ---

    Public Const FALSEVAL As Integer = 0
    Public Const TRUEVAL As Integer = 1

    Public Const GOSUB_TYPEVAL As Integer = 0
    Public Const WHEN_CANCEL_TYPEVAL As Integer = 1
    Public Const WHEN_ESCAPE_TYPEVAL As Integer = 2
    Public Const WHEN_ERROR_TYPEVAL As Integer = 3

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

    Public CurrDevNum As Int64
    Public CurrVolName As String
    Public CurrLibName As String
    Public CurrJumpPoint As Int64
    Public CurrIP As String

    Public GosubStack As New Stack(Of CallStackItem)
    Public KeyboardQueue As New Queue(Of Char)
    Public KeyboardLocked As Boolean = False

    Public DebugFlag As Boolean
    Public DebugFlagLevel As Int64
    Public DebugLogFile As Int64

    Public Last_DoEvents As Single
    Public CheckDoEventsCount As Int64

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

    Public GraphicsCharFlag As Boolean

    ' --- Form loaded flags ---

    Public rtFormMainLoaded As Boolean
    Public rtDebugLogLoaded As Boolean

    ' --- Printing variables ---

    Public PrinterFileNum As Int64
    Public PrinterFileName As String
    Public PrinterType As String
    Public PrinterDeviceName As String
    Public PrinterParameters As String
    Public PrinterHandle As Int64
    Public PrinterJobNum As Int64

    ' --- Sorting Variables ---

    ' --- if MAXMEMTAGS is > 999 change SaveMemory and LoadMemory using 3 chars ---
    Public Const MAXMEMTAGS As Int64 = 255 ' max tags it can sort in memory

    Public SortFileNum As Int64
    Public SortFileName As String
    Public SortTagSize As Int64
    Public SortLineCount As Int64
    Public FetchLineCount As Int64

    Public SortTags(MAXMEMTAGS) As String ' includes record number (z14)
    Public SortIndex(MAXMEMTAGS) As Int64

    ' --- Channel Variables ---

    'TODO: ### Public ChannelPaths(MaxChannel) As String
    'TODO: ### Public ChannelFileNums(MaxChannel) as Int64

    ' --- SQL Connection variables ---

    'TODO: ### Public cnSQL As ADODB.Connection

    ' --- Server connection variables ---

    Public HostIP As String
    Public PortVal As Int64
    Public BackgroundIP As String
    Public BackgroundPort As Int64

    Public ClientList As String
    'TODO: ### Public ReadOnly As Boolean

    Public ReadyToRun As Boolean
    Public PendingInput As String
    Public PendingOutput As String

    Public SpawnTarget As String

    Public ServerSendComplete As Boolean

    ' --- File access variables ---

    Public LockFlag As Boolean
    Public HasLockedRec As Boolean
    Public LockedFileNum As Int64
    Public LockedSQLTable As String
    Public LockedRecNum As Int64
    Public LockedRecLen As Int64
    Public LockedResource As String
    'TODO: ### Public LockedCadolXref As rtCadolXref
    'TODO: ### Public rsLockedRec As ADODB.Recordset
    'TODO: ### Public rsLockedResult As ADODB.Recordset
    Public SQLSubQuery As String
    Public SQLSubQueryFile As String
    'TODO: ### Public Files(MaxFile) As rtCadolFile
    Public UpdateRecField As Boolean
    'TODO: ### Public rsRecord As ADODB.Recordset

    ' --- Global Registers ---

    'TODO: ### Public rsGRegs As ADODB.Recordset

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

    ' --- Screen Constants ---

    Public Const ScreenWidth As Integer = 79
    Public Const ScreenHeight As Integer = 23
    Public Const ScreenWidthP1 As Integer = ScreenWidth + 1
    Public Const ScreenHeightP1 As Integer = ScreenHeight + 1
    Public Const AttributeChar As Byte = 0

End Module
