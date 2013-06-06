VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSetDfltPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'A bulk of this code is based on the MSDN's Article ID Q167735
'entitled "Setting Printer to Item in the Printers Collection
'Fails."
'
'"SYMPTOMS
'Attempting to set the default printer to an object variable has
'no effect. For instance, given a system with more than one
'printer installed, the following code will not change the
'default printer:
'
'   Private Sub Form_Load()
'       Dim Prt As Printer
'       For Each Prt In Printers
'      If Not Prt Is Printer Then
'            Set Printer = Prt
'         Exit For
'      End If
'       Next
'
'      Printer.Print "Hi, Mom"
'      Printer.EndDoc
'   End Sub
'
'The expected behavior is that the document should print to the
'first non-default printer found in the printers collection. The
'actual behavior is that the document prints to the original
'default printer." - Source: Microsoft's MSDN Article ID#Q167735
'
'I modified this code from it original and wrapped in a class for
'the purpose of storing the original printer configuration during
'class initialization and reseting it back during
'termination if it was modified.
'***************************************************************

Option Explicit

'Retrieves the string associated with the specified key in
'the given section of the WIN.INI file
Private Declare Function GetProfileString Lib "kernel32" _
    Alias "GetProfileStringA" _
    (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long

'Copies a string into the specified section of the WIN.INI file
Private Declare Function WriteProfileString Lib "kernel32" _
    Alias "WriteProfileStringA" _
    (ByVal lpszSection As String, _
    ByVal lpszKeyName As String, _
    ByVal lpszString As String) As Long

'Sends a message to the window (via hwnd) and does not return
'until the window procedure has processed the message.
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As String) As Long

Private Const HWND_BROADCAST = &HFFFF   'Used to send messages to all top-level windows in the system by
                                        'specifying HWND_BROADCAST as the first parameter to the SendMessage
                                        
Private Const WM_WININICHANGE = &H1A    'The WM_WININICHANGE message is obsolete. It is included for
                                        'compatibility with earlier versions of the system. New
                                        'applications should use the WM_SETTINGCHANGE message.

'Data structure contains operating system version information
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Returns information that a program can use to identify the operating system
Private Declare Function GetVersionExA Lib "kernel32" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

'Function retrieves a handle identifying the specified printer or print server
Private Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" _
    (ByVal pPrinterName As String, _
    phPrinter As Long, _
    pDefault As PRINTER_DEFAULTS) As Long

'Function sets the data for a specified printer or sets the state of the specified
'printer by pausing printing, resuming printing, or clearing all print jobs
Private Declare Function SetPrinter Lib "winspool.drv" _
    Alias "SetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    pPrinter As Any, _
    ByVal Command As Long) As Long

'Function retrieves information about a specified printer
Private Declare Function GetPrinter Lib "winspool.drv" _
    Alias "GetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    pPrinter As Any, _
    ByVal cbBuf As Long, _
    pcbNeeded As Long) As Long

'Function copies a string to a buffer
Private Declare Function lstrcpy Lib "kernel32" _
    Alias "lstrcpyA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As Any) As Long

'Function closes the specified printer object
Private Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

'Function returns the calling thread's last-error code value
Private Declare Function GetLastError Lib "kernel32" () As Long

'Constants for DEVMODE structure
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

'Constants for DesiredAccess member of PRINTER_DEFAULTS
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

'Constant that goes into PRINTER_INFO_5 Attributes member
'to set it as default
Private Const PRINTER_ATTRIBUTE_DEFAULT = 4

'Data structure contains information about the device initialization
'and environment of a printer
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long        'Windows 95 only
    dmICMIntent As Long        'Windows 95 only
    dmMediaType As Long        'Windows 95 only
    dmDitherType As Long       'Windows 95 only
    dmReserved1 As Long        'Windows 95 only
    dmReserved2 As Long        'Windows 95 only
End Type

'Data structure specifies detailed printer information.
Private Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

'Data structure specifies the default data type, environment,
'initialization data, and access rights for a printer.
Private Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As DEVMODE
    DesiredAccess As Long
End Type

'Member variables
Private m_sCurrPrinterDevName As String
Private m_sPrevPrinterDevName As String
Private m_sPrevPrinterDriver As String
Private m_sPrevPrinterPort As String

Private Function PtrCtoVbString(Add As Long) As String
    'Because Microsoft Visual Basic does not support a pointer data type,
    'you cannot directly receive a pointer (such as a LPSTR) as the return
    'value from a Windows API or DLL function.

    'You can work around this by receiving the return value as a long
    'integer data type. Then use the lstrcpy Windows API function to copy
    'the returned string into a Visual Basic string.
    'Source - Article ID: Q78304
    
    Dim sTemp As String * 512, X As Long
    
    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Function SetDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal PrinterPort As String) As Boolean
    Dim DeviceLine As String
    Dim r As Long
    Dim L As Long
    
    DeviceLine = DeviceName & "," & DriverName & "," & PrinterPort
    'Store the new printer information in the [WINDOWS] section of
    'the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    
    If r Then
        'Cause all applications to reload the INI file:
        L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
        SetDefaultPrinter = True
        m_sCurrPrinterDevName = DeviceName
    Else
        SetDefaultPrinter = False
    End If
End Function

Private Function Win95SetDefaultPrinter(ByRef DeviceName As String) As Boolean
    Dim Handle As Long          'handle to printer
    Dim pd As PRINTER_DEFAULTS
    Dim X As Long
    Dim need As Long            'bytes needed
    Dim pi5 As PRINTER_INFO_5   'your PRINTER_INFO structure
    Dim LastError As Long

    ' none - exit
    If DeviceName = "" Then
        Win95SetDefaultPrinter = False
        Exit Function
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS

    'Get a handle to the printer
    X = OpenPrinter(DeviceName, Handle, pd)
    'failed the open
    If X = False Then
        Win95SetDefaultPrinter = False
        Exit Function
    End If

    'Make an initial call to GetPrinter, requesting Level 5
    '(PRINTER_INFO_5) information, to determine how many bytes
    'you need
    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    'don't want to check GetLastError here - it's supposed to fail
    'with a 122 - ERROR_INSUFFICIENT_BUFFER
    'redim t as large as you need...
    ReDim t((need \ 4)) As Long

    'and call GetPrinter for keepers this time
    X = GetPrinter(Handle, 5, t(0), need, need)
    'failed the GetPrinter
    If X = False Then
        Win95SetDefaultPrinter = False
        Exit Function
    End If

    'Set the members of the pi5 structure for use with SetPrinter.
    'PtrCtoVbString copies the memory pointed at by the two string
    'pointers contained in the t() array into a Visual Basic string.
    'The other three elements are just DWORDS (long integers) and
    'don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    'This is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

    'Call SetPrinter to set it
    X = SetPrinter(Handle, 5, pi5, 0)
    'failed the SetPrinter
    If X = False Then
        Win95SetDefaultPrinter = False
        Exit Function
    End If

    ' and close the handle
    Call ClosePrinter(Handle)
    m_sCurrPrinterDevName = DeviceName
    Win95SetDefaultPrinter = True
End Function

Private Sub GetDriverAndPort(ByVal Buffer As String, ByRef DriverName As String, ByRef PrinterPort As String)
    Dim iDriver As Integer
    Dim iPort As Integer
    
    DriverName = ""
    PrinterPort = ""

    'The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then
        'Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        'The port name is the second entry after the driver name
        'separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            'Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
        End If
    End If
End Sub

Private Function WinNTSetDefaultPrinter(ByRef DeviceName As String) As Boolean
    Dim Buffer As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
    
    If DeviceName <> "" Then
        'Get the printer information for the currently selected
        'printer in the list. The information is taken from the
        'WIN.INI file.
        Buffer = Space(1024)
        r = GetProfileString("PrinterPorts", DeviceName, "", Buffer, Len(Buffer))
        
        'Parse the driver name and port name out of the buffer
        Call GetDriverAndPort(Buffer, DriverName, PrinterPort)

        If DriverName <> "" And PrinterPort <> "" Then
            WinNTSetDefaultPrinter = SetDefaultPrinter(DeviceName, DriverName, PrinterPort)
        Else
            WinNTSetDefaultPrinter = False
        End If
    End If
End Function

Function SetPrinterAsDefault(ByVal DeviceName As String) As Boolean
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    
    'TM20020912 Fault 1432 - Detect the difference between newer versions of Windows and then
    'either use the WinNT or Win95 function, depending which is more similar to the OS version.
    
    'If its not currently set as the default then set it...
    If m_sCurrPrinterDevName <> DeviceName Then
        
        'Windows NT 3.1...
        If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 And osinfo.dwBuildNumber = 1057 And osinfo.dwPlatformId = 2 Then
          SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)
        
        'Windows 95...
        ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwPlatformId = 1 Then  'And osinfo.dwBuildNumber = 67109814
          SetPrinterAsDefault = Win95SetDefaultPrinter(DeviceName)
        
        'Windows NT 4.0...
        ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwBuildNumber = 1381 And osinfo.dwPlatformId = 2 Then
          SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)
        
        'Windows 98...
        ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 10 Then
          SetPrinterAsDefault = Win95SetDefaultPrinter(DeviceName)
          
        'Windows Me...
        ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 90 Then
          SetPrinterAsDefault = Win95SetDefaultPrinter(DeviceName)
          
        'Windows 2000...
        ElseIf osinfo.dwMajorVersion = 5 And osinfo.dwMinorVersion = 0 Then
          SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)
          
        'Windows XP...
        ElseIf osinfo.dwMajorVersion = 5 And osinfo.dwMinorVersion = 1 Then
          SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)
          
        'Windows .NET Server...
        ElseIf osinfo.dwMajorVersion = 5 And osinfo.dwMinorVersion = 2 Then
          SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)
          
        End If
    Else
        SetPrinterAsDefault = True
    End If
End Function

Private Sub Class_Initialize()
    Dim Buffer As String
    Dim r As Long
    
    Buffer = Space(8192)
    r = GetProfileString("windows", "Device", "", Buffer, Len(Buffer))
    If r Then
        'Remove the wasted space
        Buffer = Mid(Buffer, 1, r)
        'Store the current default printer before we change it
        m_sPrevPrinterDevName = Mid(Buffer, 1, InStr(Buffer, ",") - 1)
        m_sPrevPrinterDriver = Mid(Buffer, InStr(Buffer, ",") + 1, InStrRev(Buffer, ",") - InStr(Buffer, ",") - 1)
        m_sPrevPrinterPort = Mid(Buffer, InStrRev(Buffer, ",") + 1)
    Else
        m_sPrevPrinterDevName = ""
        m_sPrevPrinterDriver = ""
        m_sPrevPrinterDevName = ""
    End If
    m_sCurrPrinterDevName = m_sPrevPrinterDevName
End Sub

Private Sub Class_Terminate()
  
  On Error Resume Next
  
  'Set it back before we leave...
 If gblnResetPrinterDefaultBack = True Then
  Call SetPrinterAsDefault(m_sPrevPrinterDevName)
  gblnResetPrinterDefaultBack = False
End If

End Sub


