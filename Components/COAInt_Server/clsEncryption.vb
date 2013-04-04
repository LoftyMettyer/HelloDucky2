Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsEncryption_NET.clsEncryption")> Public Class clsEncryption
	Private InitTrue As Boolean
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMem Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private byteArray() As Byte
	Private hiByte As Integer
	Private hiBound As Integer
	Private AddTbl(255, 255) As Byte
	Private XTbl(255, 255) As Byte
	Private LsTbl(255, 255) As Byte
	Private RsTbl(255, 255) As Byte
	
	Private Sub InitTbl()
		If InitTrue = True Then Exit Sub
		Dim i As Short
		Dim j As Short
		Dim k As Short
		For i = 0 To 255
			For j = 0 To 255
				XTbl(i, j) = CByte(i Xor j)
				AddTbl(i, j) = CByte((i + j) Mod 255)
			Next j
		Next i
		InitTrue = True
	End Sub
	Private Sub Append(ByRef StringData As String, Optional ByRef Length As Integer = 0)
		Dim DataLength As Integer
		If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
		If DataLength + hiByte > hiBound Then
			hiBound = hiBound + 1024
			ReDim Preserve byteArray(hiBound)
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CopyMem(VarPtr(byteArray(hiByte)), StringData, DataLength)
		hiByte = hiByte + DataLength
	End Sub
	Private Function DeHex(ByRef Data As String) As String
		Dim iCount As Double
		Reset_Renamed()
		For iCount = 1 To Len(Data) Step 2
			Append(Chr(Val("&H" & Mid(Data, iCount, 2))))
		Next 
		DeHex = GData
		Reset_Renamed()
	End Function
	Public Function EnHex(ByRef Data As String) As String
		Dim iCount As Double
		Dim sTemp As String
		Reset_Renamed()
		For iCount = 1 To Len(Data)
			sTemp = Hex(Asc(Mid(Data, iCount, 1)))
			If Len(sTemp) < 2 Then sTemp = "0" & sTemp
			Append(sTemp)
		Next 
		EnHex = GData
		Reset_Renamed()
	End Function
	Private Function FileExist(ByRef FileName As String) As Boolean
		On Error GoTo errorhandler
		'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
		GoSub begin
		
errorhandler: 
		FileExist = False
		Exit Function
		
begin: 
		Call FileLen(FileName)
		FileExist = True
	End Function
	Private ReadOnly Property GData() As String
		Get
			Dim StringData As String
			StringData = Space(hiByte)
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMem(StringData, VarPtr(byteArray(0)), hiByte)
			GData = StringData
		End Get
	End Property
	Public Function EncryptFile(ByRef InFile As String, ByRef OutFile As String, ByRef Overwrite As Boolean, Optional ByRef Key As String = "") As Boolean
		On Error GoTo errorhandler
		'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
		GoSub begin
		
errorhandler: 
		EncryptFile = False
		Exit Function
		
begin: 
		If FileExist(InFile) = False Then
			EncryptFile = False
			Exit Function
		End If
		If FileExist(OutFile) = True And Overwrite = False Then
			EncryptFile = False
			Exit Function
		End If
		Dim FileO As Short
		Dim Buffer() As Byte
		Dim bKey() As Byte
		Dim bOut() As Byte
		FileO = FreeFile
		FileOpen(FileO, InFile, OpenMode.Binary)
		ReDim Buffer(LOF(FileO))
		Buffer(LOF(1)) = 32
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FileO, Buffer)
		FileClose(FileO)
		
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		bKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Key, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object EncryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bOut = EncryptByte(Buffer, bKey)
		If FileExist(OutFile) = True Then Kill(OutFile)
		FileO = FreeFile
		FileOpen(FileO, OutFile, OpenMode.Binary)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(FileO, bOut)
		FileClose(FileO)
		EncryptFile = True
	End Function
	Public Function EncryptString(ByRef Text As String, Optional ByRef Key As String = "", Optional ByRef OutputInHex As Boolean = False) As String
		Dim byteArray() As Byte
		Dim bKey() As Byte
		Dim bOut() As Byte
		Text = Text & " "
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		byteArray = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Text, vbFromUnicode))
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		bKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Key, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object EncryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bOut = EncryptByte(byteArray, bKey)
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		EncryptString = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bOut), vbUnicode)
		If OutputInHex = True Then EncryptString = EnHex(EncryptString)
	End Function
	Public Function DecryptString(ByRef Text As String, Optional ByRef Key As String = "", Optional ByRef IsTextInHex As Boolean = False) As String
		Dim byteArray() As Byte
		Dim bKey() As Byte
		Dim bOut() As Byte
		If IsTextInHex = True Then Text = DeHex(Text)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		byteArray = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Text, vbFromUnicode))
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		bKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Key, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object DecryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bOut = DecryptByte(byteArray, bKey)
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		DecryptString = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bOut), vbUnicode)
	End Function
	Public Function DecryptFile(ByRef InFile As String, ByRef OutFile As String, ByRef Overwrite As Boolean, Optional ByRef Key As String = "") As Boolean
		On Error GoTo errorhandler
		'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
		GoSub begin
		
errorhandler: 
		DecryptFile = False
		Exit Function
		
begin: 
		If FileExist(InFile) = False Then
			DecryptFile = False
			Exit Function
		End If
		If FileExist(OutFile) = True And Overwrite = False Then
			DecryptFile = False
			Exit Function
		End If
		Dim FileO As Short
		Dim Buffer() As Byte
		Dim bKey() As Byte
		Dim bOut() As Byte
		FileO = FreeFile
		FileOpen(FileO, InFile, OpenMode.Binary)
		ReDim Buffer(LOF(FileO) - 1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FileO, Buffer)
		FileClose(FileO)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		bKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Key, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object DecryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bOut = DecryptByte(Buffer, bKey)
		If FileExist(OutFile) = True Then Kill(OutFile)
		FileO = FreeFile
		FileOpen(FileO, OutFile, OpenMode.Binary)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(FileO, bOut)
		FileClose(FileO)
		DecryptFile = True
	End Function
	'UPGRADE_NOTE: Reset was upgraded to Reset_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Reset_Renamed()
		hiByte = 0
		hiBound = 1024
		ReDim byteArray(hiBound)
	End Sub
	Public Function EncryptByte(ByRef ds() As Byte, ByRef pass() As Byte) As Object
		Call InitTbl()
		Dim tmp2() As Byte
		Dim p As Short
		Dim i As Integer
		Dim Bound As Short
		ReDim tmp2((UBound(ds)) + 4)
		Randomize()
		tmp2(0) = Int((Rnd() * 254) + 1)
		tmp2(1) = Int((Rnd() * 254) + 1)
		tmp2(2) = Int((Rnd() * 254) + 1)
		tmp2(3) = Int((Rnd() * 254) + 1)
		tmp2(4) = Int((Rnd() * 254) + 1)
		
		Call CopyMem(tmp2(5), ds(0), UBound(ds))
		ReDim ds(UBound(tmp2))
		ds = VB6.CopyArray(tmp2)
		ReDim tmp2(0)
		Bound = (UBound(pass) - 1)
		p = 0
		
		For i = 0 To UBound(ds) - 1
			If p = Bound Then p = 0
			ds(i) = XTbl(ds(i), AddTbl(ds(i + 1), pass(p)))
			ds(i + 1) = XTbl(ds(i), ds(i + 1))
			ds(i) = XTbl(ds(i), AddTbl(ds(i + 1), pass(p + 1)))
			p = p + 1
		Next i
		
		EncryptByte = VB6.CopyArray(ds)
	End Function
	Public Function DecryptByte(ByRef ds() As Byte, ByRef pass() As Byte) As Object
		Call InitTbl()
		Dim tmp2() As Byte
		Dim p As Integer
		Dim i As Integer
		Dim Bound As Short
		Bound = (UBound(pass) - 1)
		p = (UBound(ds)) Mod (UBound(pass) - 1)
		For i = (UBound(ds)) To 1 Step -1
			If p = 0 Then p = Bound
			ds(i - 1) = XTbl(ds(i - 1), AddTbl(ds(i), pass(p)))
			ds(i) = XTbl(ds(i - 1), ds(i))
			ds(i - 1) = XTbl(ds(i - 1), AddTbl(ds(i), pass(p - 1)))
			p = p - 1
		Next i
		tmp2 = VB6.CopyArray(ds)
		ReDim ds((UBound(tmp2)) - 4)
		Call CopyMem(ds(0), tmp2(5), UBound(ds))
		ReDim Preserve ds(UBound(ds) - 1)
		DecryptByte = VB6.CopyArray(ds)
	End Function
	Private Function LShift(ByVal ds As Byte, ByVal n As Byte) As Object
		Dim Lsbyte As Byte
		Dim i As Byte
		n = n Mod 8
		For i = 0 To n
			Lsbyte = 128 * CShort(ds And 1)
			Lsbyte = Lsbyte + (CShort(ds And 254) / 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object LShift. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LShift = Lsbyte
		Next i
	End Function
	Private Function RShift(ByVal ds As Byte, ByVal n As Byte) As Object
		Dim Rsbyte As Byte
		Dim i As Byte
		n = n Mod 8
		For i = 0 To n
			Rsbyte = (CShort(ds And 128) / 128)
			Rsbyte = Rsbyte + (CShort(ds And 127) * 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object RShift. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RShift = Rsbyte
		Next i
	End Function
End Class