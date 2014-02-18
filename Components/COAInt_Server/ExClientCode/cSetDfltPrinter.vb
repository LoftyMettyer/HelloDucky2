Option Strict Off
Option Explicit On
Friend Class cSetDfltPrinter
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


	'Retrieves the string associated with the specified key in
	'the given section of the WIN.INI file
	Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

	'Copies a string into the specified section of the WIN.INI file
	Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer

	'Sends a message to the window (via hwnd) and does not return
	'until the window procedure has processed the message.
	Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As String) As Integer

	Private Const HWND_BROADCAST As Integer = &HFFFF 'Used to send messages to all top-level windows in the system by
	'specifying HWND_BROADCAST as the first parameter to the SendMessage

	Private Const WM_WININICHANGE As Integer = &H1A	'The WM_WININICHANGE message is obsolete. It is included for
	'compatibility with earlier versions of the system. New
	'applications should use the WM_SETTINGCHANGE message.

	'Data structure contains operating system version information
	Private Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=128)> Public szCSDVersion() As Char
	End Structure

	'Returns information that a program can use to identify the operating system
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionExA Lib "kernel32" (ByRef lpVersionInformation As OSVERSIONINFO) As Short

	'Function retrieves a handle identifying the specified printer or print server
	'UPGRADE_WARNING: Structure PRINTER_DEFAULTS may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, ByRef phPrinter As Integer, ByRef pDefault As PRINTER_DEFAULTS) As Integer

	'Function sets the data for a specified printer or sets the state of the specified
	'printer by pausing printing, resuming printing, or clearing all print jobs
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Integer, ByVal Level As Integer, ByRef pPrinter As Object, ByVal Command_Renamed As Integer) As Integer

	'Function retrieves information about a specified printer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Integer, ByVal Level As Integer, ByRef pPrinter As Long, ByVal cbBuf As Integer, ByRef pcbNeeded As Integer) As Integer

	'Function copies a string to a buffer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Integer

	'Function closes the specified printer object
	Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Integer) As Integer

	'Function returns the calling thread's last-error code value
	Private Declare Function GetLastError Lib "kernel32" () As Integer

	'Constants for DEVMODE structure
	Private Const CCHDEVICENAME As Short = 32
	Private Const CCHFORMNAME As Short = 32

	'Constants for DesiredAccess member of PRINTER_DEFAULTS
	Private Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
	Private Const PRINTER_ACCESS_ADMINISTER As Integer = &H4
	Private Const PRINTER_ACCESS_USE As Integer = &H8
	Private Const PRINTER_ALL_ACCESS As Boolean = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

	'Constant that goes into PRINTER_INFO_5 Attributes member
	'to set it as default
	Private Const PRINTER_ATTRIBUTE_DEFAULT As Short = 4

	'Data structure contains information about the device initialization
	'and environment of a printer
	Private Structure DEVMODE
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHDEVICENAME), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=CCHDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHFORMNAME), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=CCHFORMNAME)> Public dmFormName() As Char
		Dim dmLogPixels As Short
		Dim dmBitsPerPel As Integer
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
		Dim dmICMMethod As Integer 'Windows 95 only
		Dim dmICMIntent As Integer 'Windows 95 only
		Dim dmMediaType As Integer 'Windows 95 only
		Dim dmDitherType As Integer	'Windows 95 only
		Dim dmReserved1 As Integer 'Windows 95 only
		Dim dmReserved2 As Integer 'Windows 95 only
	End Structure

	'Data structure specifies detailed printer information.
	Private Structure PRINTER_INFO_5
		Dim pPrinterName As String
		Dim pPortName As String
		Dim Attributes As Integer
		Dim DeviceNotSelectedTimeout As Integer
		Dim TransmissionRetryTimeout As Integer
	End Structure

	'Data structure specifies the default data type, environment,
	'initialization data, and access rights for a printer.
	Private Structure PRINTER_DEFAULTS
		Dim pDatatype As Integer
		Dim pDevMode As DEVMODE
		Dim DesiredAccess As Integer
	End Structure

	'Member variables
	Private m_sCurrPrinterDevName As String
	Private m_sPrevPrinterDevName As String
	Private m_sPrevPrinterDriver As String
	Private m_sPrevPrinterPort As String

	Private Function PtrCtoVbString(ByRef Add As Integer) As String
		'Because Microsoft Visual Basic does not support a pointer data type,
		'you cannot directly receive a pointer (such as a LPSTR) as the return
		'value from a Windows API or DLL function.

		'You can work around this by receiving the return value as a long
		'integer data type. Then use the lstrcpy Windows API function to copy
		'the returned string into a Visual Basic string.
		'Source - Article ID: Q78304

		Dim sTemp As New String(" ", 512)
		Dim X As Integer

		X = lstrcpy(sTemp, Add)
		If (InStr(1, sTemp, Chr(0)) = 0) Then
			PtrCtoVbString = ""
		Else
			PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
		End If
	End Function

	Private Function SetDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal PrinterPort As String) As Boolean
		Dim DeviceLine As String
		Dim r As Integer
		Dim L As Integer

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
		Dim Handle As Integer	'handle to printer
		Dim pd As PRINTER_DEFAULTS
		Dim X As Integer
		Dim need As Integer	'bytes needed
		Dim pi5 As PRINTER_INFO_5	'your PRINTER_INFO structure
		' none - exit
		If DeviceName = "" Then
			Win95SetDefaultPrinter = False
			Exit Function
		End If

		' set the PRINTER_DEFAULTS members
		pd.pDatatype = 0
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
		X = GetPrinter(Handle, 5, 0, 0, need)
		'don't want to check GetLastError here - it's supposed to fail
		'with a 122 - ERROR_INSUFFICIENT_BUFFER
		'redim t as large as you need...
		Dim t(need \ 4) As Integer

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
		'UPGRADE_WARNING: Couldn't resolve default property of object pi5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		Dim iDriver As Short
		Dim iPort As Short

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
		Dim r As Integer

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
		Dim retvalue As Short

		osinfo.dwOSVersionInfoSize = 148
		osinfo.szCSDVersion = Space(128)
		retvalue = GetVersionExA(osinfo)

		'TM20020912 Fault 1432 - Detect the difference between newer versions of Windows and then
		'either use the WinNT or Win95 function, depending which is more similar to the OS version.

		'If its not currently set as the default then set it...
		If m_sCurrPrinterDevName <> DeviceName Then

			'Windows NT 3.1...
			If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 And osinfo.dwBuildNumber = 1057 And osinfo.dwPlatformId = 2 Then
				SetPrinterAsDefault = WinNTSetDefaultPrinter(DeviceName)

				'Windows 95...
			ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwPlatformId = 1 Then	 'And osinfo.dwBuildNumber = 67109814
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

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim Buffer As String
		Dim r As Integer

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
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()

		'  On Error Resume Next
		'
		'  'Set it back before we leave...
		' If gblnResetPrinterDefaultBack = True Then
		'  Call SetPrinterAsDefault(m_sPrevPrinterDevName)
		'  gblnResetPrinterDefaultBack = False
		'End If

	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class