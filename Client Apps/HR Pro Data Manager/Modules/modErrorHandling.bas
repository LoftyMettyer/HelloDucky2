Attribute VB_Name = "modErrorHandling"
'Option Explicit
'
'' Define your custom errors here.  Be sure to use numbers
'' greater than 512, to avoid conflicts with OLE error numbers.
'Public Const MyObjectError1 = 1000
'Public Const MyObjectError2 = 1010
'Public Const MyObjectErrorN = 1234
'Public Const MyUnhandledError = 9999
'
'
'Private Const SYSEXC_MAXIMUM_PARAMETERS = 15
'
'Private Const CASIZE = 14
'
'Private Type CONTEXT
'  Dbls(0 To 66) As Double
'  Longs(0 To 6) As Long
'End Type
'
'Private Type SYSEXC_RECORD
'    ExceptionCode As Long
'    ExceptionFlags As Long
'    pExceptionRecord As Long
'    ExceptionAddress As Long
'    NumberParameters As Long
'    ExceptionInformation(SYSEXC_MAXIMUM_PARAMETERS) As Long
'End Type
'
'Private Type SYSEXC_POINTERS
'    pExceptionRecord As SYSEXC_RECORD
'    ContextRecord As CONTEXT
'End Type
'
'Public Enum ENUM_ERRMAP
'    ERRMAP_FIRST = vbObjectError + 4096             'vbObjectError = $H80040000 = -2147221504
'    ERRMAP_RESERVED_FIRST = ERRMAP_FIRST            'Errors reserved for HuntERR and UJ apps.
'       ERR_SYSEXCEPTION                             'System exception like access violation
'    ERRMAP_RESERVED_LAST = ERRMAP_RESERVED_FIRST + 100
'    ERRMAP_EXC_FIRST = ERRMAP_RESERVED_LAST + 1     'Exceptions - reraised by ErrorIn
'        EXC_GENERAL = ERRMAP_EXC_FIRST
'    ERRMAP_EXC_LAST = ERRMAP_EXC_FIRST + 1000
'    ERRMAP_APP_FIRST
'        ERR_GENERAL = ERRMAP_APP_FIRST              ' = vbObjectError + 4096+1+1000+1
'        'Application errors here
'End Enum
'
'Private Enum ENUM_SYSEXC
'    SYSEXC_ACCESS_VIOLATION = &HC0000005
'    SYSEXC_DATATYPE_MISALIGNMENT = &H80000002
'    SYSEXC_BREAKPOINT = &H80000003
'    SYSEXC_SINGLE_STEP = &H80000004
'    SYSEXC_ARRAY_BOUNDS_EXCEEDED = &HC000008C
'    SYSEXC_FLT_DENORMAL_OPERAND = &HC000008D
'    SYSEXC_FLT_DIVIDE_BY_ZERO = &HC000008E
'    SYSEXC_FLT_INEXACT_RESULT = &HC000008F
'    SYSEXC_FLT_INVALID_OPERATION = &HC0000090
'    SYSEXC_FLT_OVERFLOW = &HC0000091
'    SYSEXC_FLT_STACK_CHECK = &HC0000092
'    SYSEXC_FLT_UNDERFLOW = &HC0000093
'    SYSEXC_INT_DIVIDE_BY_ZERO = &HC0000094
'    SYSEXC_INT_OVERFLOW = &HC0000095
'    SYSEXC_PRIVILEGED_INSTRUCTION = &HC0000096
'    SYSEXC_IN_PAGE_ERROR = &HC0000006
'    SYSEXC_ILLEGAL_INSTRUCTION = &HC000001D
'    SYSEXC_NONCONTINUABLE_EXCEPTION = &HC0000025
'    SYSEXC_STACK_OVERFLOW = &HC00000FD
'    SYSEXC_INVALID_DISPOSITION = &HC0000026
'    SYSEXC_GUARD_PAGE_VIOLATION = &H80000001
'    SYSEXC_INVALID_HANDLE = &HC0000008
'    SYSEXC_CONTROL_C_EXIT = &HC000013A
'End Enum
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'    (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Sub CopyExceptionRecord Lib "kernel32" Alias "RtlMoveMemory" (pDest As SYSEXC_RECORD, ByVal LPSYSEXC_RECORD As Long, ByVal lngBytes As Long)
'Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" _
'    (ByVal lpTopLevelExceptionFilter As Long) As Long
'
'Private Function GetErrorTextFromResource(ErrorNum As Long) _
'          As String
'      Dim strMsg As String
'
'
'      ' this function will retrieve an error description from a resource
'      ' file (.RES).  The ErrorNum is the index of the string
'      ' in the resource file.  Called by RaiseError
'
'
'      On Error GoTo GetErrorTextFromResourceError
'
'
'      ' get the string from a resource file
'      GetErrorTextFromResource = LoadResString(ErrorNum)
'
'
'      Exit Function
'
'
'GetErrorTextFromResourceError:
'
'
'      If Err.Number <> 0 Then
'            GetErrorTextFromResource = "An unknown error has occurred!"
'      End If
'
'
'End Function
'
'
'Public Sub RaiseError(ErrorNumber As Long, Source As String)
'      Dim strErrorText As String
'
'
'      'there are a number of methods for retrieving the error
'      'message.  The following method uses a resource file to
'      'retrieve strings indexed by the error number you are
'      'raising.
'      strErrorText = GetErrorTextFromResource(ErrorNumber)
'
'
'      'raise an error back to the client
'      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText
'
'
'End Sub
'
'Private Function SysExcHandler(ByRef ExcPtrs As SYSEXC_POINTERS) As Long
'
'  On Error GoTo ErrorTrap
'
'  Dim ExcRec As SYSEXC_RECORD, strExc As String
'
'  ExcRec = ExcPtrs.pExceptionRecord
'  Do Until ExcRec.pExceptionRecord = 0
'    CopyExceptionRecord ExcRec, ExcRec.pExceptionRecord, Len(ExcRec)
'  Loop
'
'  strExc = GetExcAsText(ExcRec.ExceptionCode)
'  Err.Raise ERR_SYSEXCEPTION, Err.Source, "(&H" & Hex$(ExcRec.ExceptionCode) & ") " & strExc
'
'ErrorTrap:
'  COAMsgBox Err.Description
'  End
'
'End Function
'
'Private Function GetExcAsText(ByVal ExcNum As Long) As String
'  Select Case ExcNum
'    Case SYSEXC_ACCESS_VIOLATION:          GetExcAsText = "Access violation"
'    Case SYSEXC_DATATYPE_MISALIGNMENT:     GetExcAsText = "Datatype misalignment"
'    Case SYSEXC_BREAKPOINT:                GetExcAsText = "Breakpoint"
'    Case SYSEXC_SINGLE_STEP:               GetExcAsText = "Single step"
'    Case SYSEXC_ARRAY_BOUNDS_EXCEEDED:     GetExcAsText = "Array bounds exceeded"
'    Case SYSEXC_FLT_DENORMAL_OPERAND:      GetExcAsText = "Float Denormal Operand"
'    Case SYSEXC_FLT_DIVIDE_BY_ZERO:        GetExcAsText = "Divide By Zero"
'    Case SYSEXC_FLT_INEXACT_RESULT:        GetExcAsText = "Floating Point Inexact Result"
'    Case SYSEXC_FLT_INVALID_OPERATION:     GetExcAsText = "Invalid Operation"
'    Case SYSEXC_FLT_OVERFLOW:              GetExcAsText = "Float Overflow"
'    Case SYSEXC_FLT_STACK_CHECK:           GetExcAsText = "Float Stack Check"
'    Case SYSEXC_FLT_UNDERFLOW:             GetExcAsText = "Float Underflow"
'    Case SYSEXC_INT_DIVIDE_BY_ZERO:        GetExcAsText = "Integer Divide By Zero"
'    Case SYSEXC_INT_OVERFLOW:              GetExcAsText = "Integer Overflow"
'    Case SYSEXC_PRIVILEGED_INSTRUCTION:    GetExcAsText = "Privileged Instruction"
'    Case SYSEXC_IN_PAGE_ERROR:             GetExcAsText = "In Page Error"
'    Case SYSEXC_ILLEGAL_INSTRUCTION:       GetExcAsText = "Illegal Instruction"
'    Case SYSEXC_NONCONTINUABLE_EXCEPTION:  GetExcAsText = "Non Continuable Exception"
'    Case SYSEXC_STACK_OVERFLOW:            GetExcAsText = "Stack Overflow"
'    Case SYSEXC_INVALID_DISPOSITION:       GetExcAsText = "Invalid Disposition"
'    Case SYSEXC_GUARD_PAGE_VIOLATION:      GetExcAsText = "Guard Page Violation"
'    Case SYSEXC_INVALID_HANDLE:            GetExcAsText = "Invalid Handle"
'    Case SYSEXC_CONTROL_C_EXIT:            GetExcAsText = "Control-C Exit"
'  End Select
'End Function
'
'' Turn on the exception violation handler
'Public Sub SetExceptionHandler()
'  Call SetUnhandledExceptionFilter(0)
'  Call SetUnhandledExceptionFilter(AddressOf SysExcHandler)
'End Sub
