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

Private Const HWND_BROADCAST = &HFFFF
Private Const WM_WININICHANGE = &H1A

' constants for DEVMODE structure
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Private Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

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
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Private Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Private Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Private Declare Function GetProfileString Lib "kernel32" _
Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hWnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As String) As Long

Private Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long

Private Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long

Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub Win95SetDefaultPrinter(strPrinterName As String)
    Dim Handle As Long          'handle to printer
    Dim pd As PRINTER_DEFAULTS
    Dim x As Long
    Dim need As Long            ' bytes needed
    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
    'Dim LastError As Long

    ' determine which printer was selected
    ' none - exit
    If strPrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    x = OpenPrinter(strPrinterName, Handle, pd)
    ' failed the open
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    x = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

       ' call SetPrinter to set it
       x = SetPrinter(Handle, 5, pi5, 0)

       If x = False Then   ' SetPrinter failed
           COAMsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
           Exit Sub
       Else
           If Printer.DeviceName <> strPrinterName Then
           ' Make sure Printer object is set to the new printer
              SelectPrinter strPrinterName
           End If
       End If

    ' and close the handle
    ClosePrinter (Handle)
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub WinNTSetDefaultPrinter(strPrinterName As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    'Dim PrinterName As String
    Dim r As Long
    If strPrinterName <> vbNullString Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        Buffer = Space(1024)
        r = GetProfileString("PrinterPorts", strPrinterName, "", _
            Buffer, Len(Buffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

        If DriverName <> "" And PrinterPort <> "" Then
            SetDefaultPrinter strPrinterName, DriverName, PrinterPort
            If Printer.DeviceName <> strPrinterName Then
            ' Make sure Printer object is set to the new printer
               SelectPrinter strPrinterName
            End If
        End If

    End If
End Sub

Public Function SetPrinterAsDefault(strPrinterName As String) As Boolean

  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer

  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(osinfo)

  If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    Call Win95SetDefaultPrinter(strPrinterName)
  Else
  ' This assumes that future versions of Windows use the NT method
    Call WinNTSetDefaultPrinter(strPrinterName)
  End If

  'MH20030922 Fault 6124 (Q257688)
  Printer.TrackDefault = True

End Function

Private Sub SelectPrinter(NewPrinter As String)

  Dim Prt As Printer

  For Each Prt In Printers
    If LCase(Trim(Prt.DeviceName)) = LCase(Trim(NewPrinter)) Then
      Set Printer = Prt
      'MH20030922 Fault 6124 (Q257688)
      Printer.TrackDefault = True
      Exit For
    End If
  Next

End Sub
