Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsCrypt_NET.clsCrypt")> Public Class clsCrypt
	
	Private mfInitTrue As Boolean
	Private mabytArray() As Byte
	Private mlngHiByte As Integer
	Private mlngHiBound As Integer
	Private mabytAddTable(255, 255) As Byte
	Private mabytXTable(255, 255) As Byte
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
  Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Long, ByVal Length As Integer)

  ' TODO - Values that are not supported in .NET upgrade - hardcode values for now.
  Private Const vbFromUnicode = 128
  Private Const vbUnicode = 64

	Public Function EncryptString(ByRef psText As String, Optional ByRef psKey As String = "", Optional ByRef pbOutputInHex As Boolean = False) As String
		
		Dim abytArray() As Byte
		Dim abytKey() As Byte
		Dim abytOut() As Byte
		
		psText = psText & " "
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
    abytArray = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(psText, vbFromUnicode))
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		abytKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(psKey, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object EncryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		abytOut = EncryptByte(abytArray, abytKey)
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		EncryptString = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(abytOut), vbUnicode)
		
		If pbOutputInHex = True Then EncryptString = EnHex(EncryptString)
		
	End Function
	
	Public Function EnHex(ByRef psData As String) As String
		Dim dblCount As Double
		Dim sTemp As String
		
		Reset()
		
		For dblCount = 1 To Len(psData)
			sTemp = Hex(Asc(Mid(psData, dblCount, 1)))
			If Len(sTemp) < 2 Then sTemp = "0" & sTemp
			Append(sTemp)
		Next 
		
		EnHex = GData
		
		Reset()
		
	End Function
	
	Public Function EncryptByte(ByRef pabytText() As Byte, ByRef pabytKey() As Byte) As Object
		Dim abytTemp() As Byte
		Dim iTemp As Short
		Dim iLoop As Integer
		Dim iBound As Short
		
		Call InitTbl()
		
		ReDim abytTemp((UBound(pabytText)) + 4)
		Randomize()
		abytTemp(0) = Int((Rnd() * 254) + 1)
		abytTemp(1) = Int((Rnd() * 254) + 1)
		abytTemp(2) = Int((Rnd() * 254) + 1)
		abytTemp(3) = Int((Rnd() * 254) + 1)
		abytTemp(4) = Int((Rnd() * 254) + 1)
		
		Call CopyMem(abytTemp(5), pabytText(0), UBound(pabytText))
		
		ReDim pabytText(UBound(abytTemp))
		pabytText = VB6.CopyArray(abytTemp)
		ReDim abytTemp(0)
		iBound = (UBound(pabytKey) - 1)
		iTemp = 0
		
		For iLoop = 0 To UBound(pabytText) - 1
			If iTemp = iBound Then iTemp = 0
			pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp)))
			pabytText(iLoop + 1) = mabytXTable(pabytText(iLoop), pabytText(iLoop + 1))
			pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp + 1)))
			iTemp = iTemp + 1
		Next iLoop
		
		EncryptByte = VB6.CopyArray(pabytText)
		
	End Function
	
	
	
	Private ReadOnly Property GData() As String
		Get
			Dim sStringData As String
			
			sStringData = Space(mlngHiByte)
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMem(sStringData, VarPtr(mabytArray(0)), mlngHiByte)
			GData = sStringData
			
		End Get
	End Property
	
	
	Private Sub Append(ByRef psStringData As String, Optional ByRef plngLength As Integer = 0)
		Dim lngDataLength As Integer
		
		If plngLength > 0 Then
			lngDataLength = plngLength
		Else
			lngDataLength = Len(psStringData)
		End If
		
		If lngDataLength + mlngHiByte > mlngHiBound Then
			mlngHiBound = mlngHiBound + 1024
			ReDim Preserve mabytArray(mlngHiBound)
		End If
		
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CopyMem(VarPtr(mabytArray(mlngHiByte)), psStringData, lngDataLength)
		mlngHiByte = mlngHiByte + lngDataLength
		
	End Sub
	
	Private Sub InitTbl()
		Dim i As Short
		Dim j As Short
		Dim k As Short
		
		If mfInitTrue = True Then Exit Sub
		
		For i = 0 To 255
			For j = 0 To 255
				mabytXTable(i, j) = CByte(i Xor j)
				mabytAddTable(i, j) = CByte((i + j) Mod 255)
			Next j
		Next i
		
		mfInitTrue = True
		
	End Sub
	
	
	
	Public Function EncryptQueryString(ByRef plngInstanceID As Integer, ByRef plngStepID As Integer, ByRef psUser As String, ByRef psPassword As String, ByRef psServerName As String, ByRef psDatabaseName As String, ByRef mADOCon As ADODB.Connection) As String
		
		On Error GoTo ErrorTrap
		
		Dim skey As String
		Dim sEncryptedString As String
		Dim sSourceString As String
		Dim sServerName As String
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		
		Const ENCRYPTIONKEY As String = "jmltn"
		
		' Get the server name - gsServerName may be '.'
		' which screws up the Workflow queryString if the web site is not
		' on the same server as the SQL database.
		'sSQL = "SELECT @@SERVERNAME AS [serverName]"
		sSQL = "SELECT SERVERPROPERTY('servername') AS [serverName]"
		rsTemp = New ADODB.Recordset
		rsTemp.Open(sSQL, mADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		With rsTemp
			If Not (.EOF And .BOF) Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rsTemp.Fields("ServerName").Value) Then
					sServerName = psServerName
				Else
					sServerName = rsTemp.Fields("ServerName").Value
				End If
			Else
				sServerName = psServerName
			End If
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		skey = ENCRYPTIONKEY
		sSourceString = CStr(plngInstanceID) & vbTab & CStr(plngStepID) & vbTab & psUser & vbTab & psPassword & vbTab & sServerName & vbTab & psDatabaseName
		
		sEncryptedString = EncryptString(sSourceString, skey, True)
		sEncryptedString = CompactString(sEncryptedString)
		
TidyUpAndExit: 
		EncryptQueryString = sEncryptedString
		Exit Function
		
ErrorTrap: 
		sEncryptedString = ""
		Resume TidyUpAndExit
		
	End Function
	Public Function CompactString(ByRef psSourceString As String) As String
		' Compact the encrypted string.
		' psSourceString is a string of the hexadecimal values of the Ascii codes for each character in the encrypted string.
		' In this string each character in the encrypted string is represented as 2 hex digits.
		' As it's a string of hex characters all characters are in the range 0-9, A-F
		' Valid hypertext link characters are 0-9, A-Z, a-z and some others (we'll be using $ and @).
		' Take advantage of this by implementing our own base64 encoding as follows:
		Dim sCompactedString As String
		Dim sSubString As String
		Dim sModifiedSourceString As String
		Dim iValue As Short
		Dim iTemp As Short
		Dim sNewString As String
		
		sCompactedString = ""
		sModifiedSourceString = psSourceString
		Do While Len(sModifiedSourceString) > 0
			' Read the hex characters in chunks of 3 (ie. possible values 0 - 4095)
			' This chunk of 3 Hex characters can then be translated into 2 base64 characters (ie. still have possible values 0 - 4095)
			' Woohoo! We've reduced the length of the encrypted string by about one third!
			sNewString = ""
			sSubString = Left(sModifiedSourceString & "000", 3)
			sModifiedSourceString = Mid(sModifiedSourceString, 4)
			iValue = Val("&H" & sSubString)
			
			' Use our own base64 digit set.
			' Base64 digit values 0-9 are represented as 0-9
			' Base64 digit values 10-35 are represented as A-Z
			' Base64 digit values 36-61 are represented as a-z
			' Base64 digit value 62 is represented as $
			' Base64 digit value 63 is represented as @
			
			iTemp = iValue Mod 64
			If iTemp = 63 Then
				sNewString = "@"
			ElseIf iTemp = 62 Then 
				sNewString = "$"
			ElseIf iTemp >= 36 Then 
				sNewString = Chr(iTemp + 61)
			ElseIf iTemp >= 10 Then 
				sNewString = Chr(iTemp + 55)
			Else
				sNewString = Chr(iTemp + 48)
			End If
			
			iTemp = (iValue - iTemp) / 64
			
			If iTemp = 63 Then
				sNewString = "@" & sNewString
			ElseIf iTemp = 62 Then 
				sNewString = "$" & sNewString
			ElseIf iTemp >= 36 Then 
				sNewString = Chr(iTemp + 61) & sNewString
			ElseIf iTemp >= 10 Then 
				sNewString = Chr(iTemp + 55) & sNewString
			Else
				sNewString = Chr(iTemp + 48) & sNewString
			End If
			
			sCompactedString = sCompactedString & sNewString
		Loop 
		
		' Append the number of characters to ignore, to the compacted string
		CompactString = sCompactedString & CStr((3 - (Len(psSourceString) Mod 3)) Mod 3)
		
	End Function
	
	
	Public Function DecompactString(ByVal psSourceString As String) As String
		Dim sModifiedSourceString As String
		Dim sDecompactedString As String
		Dim sSubString As String
		Dim sNewString As String
		Dim iTemp As Short
		Dim iTempTotal As Short
		Dim iAscCode As Short
		Dim iTrailingCharsToIgnore As Short
		
		On Error GoTo DecompactString_ERROR
		
		sDecompactedString = ""
		iTrailingCharsToIgnore = CShort(Right(psSourceString, 1))
		sModifiedSourceString = Left(psSourceString, Len(psSourceString) - 1)
		
		Do While Len(sModifiedSourceString) > 0
			sSubString = Left(sModifiedSourceString & "00", 2)
			sModifiedSourceString = Mid(sModifiedSourceString, 3)
			iTempTotal = 0
			
			iAscCode = Asc(Mid(sSubString, 2, 1))
			If iAscCode = 64 Then ' @
				iTemp = 63
			ElseIf iAscCode = 36 Then  ' $
				iTemp = 62
			ElseIf iAscCode >= 97 Then  ' a-z
				iTemp = iAscCode - 61
			ElseIf iAscCode >= 65 Then  'A-Z
				iTemp = iAscCode - 55
			Else ' 0-9
				iTemp = iAscCode - 48
			End If
			iTempTotal = iTempTotal + iTemp
			
			iAscCode = Asc(Left(sSubString, 1))
			If iAscCode = 64 Then ' @
				iTemp = 63
			ElseIf iAscCode = 36 Then  ' $
				iTemp = 62
			ElseIf iAscCode >= 97 Then  ' a-z
				iTemp = iAscCode - 61
			ElseIf iAscCode >= 65 Then  'A-Z
				iTemp = iAscCode - 55
			Else ' 0-9
				iTemp = iAscCode - 48
			End If
			iTempTotal = iTempTotal + (iTemp * 64)
			
			sNewString = Right("000" & Hex(iTempTotal), 3)
			sDecompactedString = sDecompactedString & sNewString
		Loop 
		
		sDecompactedString = Left(sDecompactedString, Len(sDecompactedString) - iTrailingCharsToIgnore)
		DecompactString = sDecompactedString
		
		Exit Function
		
DecompactString_ERROR: 
		DecompactString = ""
		
	End Function
	
	
	Public Function DecryptString(ByVal psText As String, Optional ByVal psKey As String = "", Optional ByVal pfIsTextInHex As Boolean = True) As String
		
		Dim abytArray() As Byte
		Dim abytKey() As Byte
		Dim abytOut() As Byte
		
		On Error GoTo DecryptString_ERROR
		
		Const ENCRYPTIONKEY As String = "jmltn"
		
		If Len(psKey) = 0 Then
			psKey = ENCRYPTIONKEY
		End If
		
		If pfIsTextInHex = True Then
			psText = DeHex(psText)
		End If
		
		DecryptString = psText
		'abytArray() = System.Text.Encoding.Default.GetBytes(psText)
		'abytKey = System.Text.Encoding.Default.GetBytes(psKey)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		abytArray = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(psText, vbFromUnicode))
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		abytKey = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(psKey, vbFromUnicode))
		'UPGRADE_WARNING: Couldn't resolve default property of object DecryptByte(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		abytOut = DecryptByte(abytArray, abytKey)
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		DecryptString = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(abytOut), vbUnicode) 'System.Text.Encoding.Default.GetString(abytOut)
		
		Exit Function
		
DecryptString_ERROR: 
		DecryptString = ""
		
	End Function
	
	Private Function DeHex(ByVal psData As String) As String
		
		ResetByteArray()
		Dim iCount As Short
		
		For iCount = 1 To Len(psData) Step 2
			Append((Chr(CShort("&H" & Mid(psData, iCount, 2)))))
		Next 
		
		DeHex = GData()
		
		ResetByteArray()
		
	End Function
	
	Private Sub ResetByteArray()
		
		mlngHiByte = 0
		mlngHiBound = 1024
		ReDim mabytArray(mlngHiBound)
	End Sub
	
	Public Function DecryptByte(ByRef pabytDs() As Byte, ByRef pabytPass() As Byte) As Object
		
		Dim abytTemp() As Byte
		Dim iBound As Short
		Dim iLoop As Short
		Call InitTbl()
		
		iBound = (UBound(pabytPass) - 1)
		Dim iTemp As Short
		iTemp = (UBound(pabytDs)) Mod (UBound(pabytPass) - 1)
		
		For iLoop = (UBound(pabytDs)) To 1 Step -1
			If iTemp = 0 Then iTemp = iBound
			pabytDs(iLoop - 1) = mabytXTable(pabytDs(iLoop - 1), mabytAddTable(pabytDs(iLoop), pabytPass(iTemp)))
			pabytDs(iLoop) = mabytXTable(pabytDs(iLoop - 1), pabytDs(iLoop))
			pabytDs(iLoop - 1) = mabytXTable(pabytDs(iLoop - 1), mabytAddTable(pabytDs(iLoop), pabytPass(iTemp - 1)))
			iTemp = iTemp - 1
		Next 
		
		abytTemp = VB6.CopyArray(pabytDs)
		
		ReDim pabytDs((UBound(abytTemp)) - 4)
		
		' copy index 5 to ubound -4 to pabytds.
		For iLoop = 0 To UBound(abytTemp) - 5
			pabytDs(iLoop) = abytTemp(iLoop + 5)
		Next 
		
		
		'System.Text.Encoding.Default.GetBytes(System.Text.Encoding.Default.GetString(abytTemp, 5, UBound(abytTemp) - 4)).CopyTo(pabytDs, 0)
		
		ReDim Preserve pabytDs(UBound(pabytDs) - 1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object DecryptByte. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DecryptByte = VB6.CopyArray(pabytDs)
	End Function
End Class