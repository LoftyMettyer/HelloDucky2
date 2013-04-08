Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Utilities_NET.Utilities")> Public Class Utilities
	
	Private mclsData As clsDataAccess
	
	Private mstrCommonDialogFormatsWord As String
	Private mstrCommonDialogFormatsExcel As String
	Private mintDefaultIndexWord As Short
	Private mintDefaultIndexExcel As Short
	Private mstrOfficeSaveAsValues As String
	
	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String
	
	Public Function VerifyCheckString(ByRef psString As String, ByRef psCheckSum As String) As Boolean
		' Verify a checksum of the given string
		On Error GoTo ErrorTrap
		
		Dim sCheckString As String
		Dim lngIndex As Integer
		Dim iCharCode As Short
		Dim lngTotalChar As Integer
		Dim aiKeys() As Short
		Dim fModify As Boolean
		Dim iKeyIndex As Short
		
		Const KEYCOUNT As Short = 10
		Const PREFIXKEYCOUNT As Short = 5
		
		ReDim aiKeys(KEYCOUNT)
		
		lngTotalChar = 0
		iKeyIndex = 1
		sCheckString = ""
		
		' Get the key digits.
		fModify = True
		For lngIndex = 1 To UBound(aiKeys)
			If lngIndex <= PREFIXKEYCOUNT Then
				aiKeys(lngIndex) = CShort(Mid(psCheckSum, lngIndex, 1))
			Else
				aiKeys(lngIndex) = CShort(Mid(psCheckSum, Len(psCheckSum) - KEYCOUNT + lngIndex, 1))
			End If
			
			If fModify Then
				aiKeys(lngIndex) = 9 - aiKeys(lngIndex)
			End If
			
			fModify = Not fModify
		Next lngIndex
		
		' Do the checksum of the given string
		For lngIndex = 1 To Len(psString)
			iCharCode = Asc(Mid(psString, lngIndex, 1))
			
			iCharCode = iCharCode + (aiKeys(iKeyIndex) * (iKeyIndex + 1))
			
			lngTotalChar = lngTotalChar + iCharCode
			
			iKeyIndex = iKeyIndex + 1
			If iKeyIndex > UBound(aiKeys) Then
				iKeyIndex = 1
			End If
		Next lngIndex
		
		sCheckString = Mid(psCheckSum, 1, PREFIXKEYCOUNT) & CStr(lngTotalChar) & Mid(psCheckSum, Len(psCheckSum) - PREFIXKEYCOUNT + 1)
		
TidyUpAndExit: 
		VerifyCheckString = (sCheckString = psCheckSum)
		
		Exit Function
		
ErrorTrap: 
		'Open "c:\temp\debug.txt" For Append As #99
		'Print #99, "ERROR"
		'Close #99
		Resume TidyUpAndExit
		
	End Function
	
	
	
	Public Function UDFFunctions(ByRef pbCreate As Object) As Object
		
		On Error GoTo UDFFunctions_ERROR
		
		Dim iCount As Short
		Dim strDropCode As String
		Dim strFunctionName As String
		Dim sUDFCode As String
		Dim datData As clsDataAccess
		Dim iStart As Short
		Dim iEnd As Short
		Dim strFunctionNumber As String
		
		Const FUNCTIONPREFIX As String = "udf_ASRSys_"
		
		If gbEnableUDFFunctions Then
			
			For iCount = 1 To UBound(mastrUDFsRequired)
				
				'JPD 20060110 Fault 10509
				'iStart = Len("CREATE FUNCTION udf_ASRSys_") + 1
				iStart = InStr(mastrUDFsRequired(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
				iEnd = InStr(1, Mid(mastrUDFsRequired(iCount), 1, 1000), "(@Per")
				strFunctionNumber = Mid(mastrUDFsRequired(iCount), iStart, iEnd - iStart)
				strFunctionName = FUNCTIONPREFIX & strFunctionNumber
				
				'Drop existing function (could exist if the expression is used more than once in a report)
				strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
				
				mclsData.ExecuteSql(strDropCode)
				
				' Create the new function
				'UPGRADE_WARNING: Couldn't resolve default property of object pbCreate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If pbCreate Then
					sUDFCode = mastrUDFsRequired(iCount)
					mclsData.ExecuteSql(sUDFCode)
				End If
				
			Next iCount
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		UDFFunctions = True
		Exit Function
		
UDFFunctions_ERROR: 
		'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		UDFFunctions = False
		
	End Function
	
	Public Function FormatEventDuration(ByRef lngSeconds As Integer) As String
		
		Dim strHours As String
		Dim strMins As String
		Dim strSeconds As String
		Dim dblRemainder As Double
		
		Const TIME_SEPARATOR As String = ":"
		
		If Not (lngSeconds < 0) Then
			strHours = CStr(Fix(lngSeconds / 3600))
			strHours = New String("0", 2 - Len(strHours)) & strHours
			dblRemainder = CDbl(lngSeconds Mod 3600)
			
			strMins = CStr(Fix(dblRemainder / 60))
			strMins = New String("0", 2 - Len(strMins)) & strMins
			'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dblRemainder = CDbl(dblRemainder Mod 60)
			
			strSeconds = CStr(Fix(dblRemainder))
			strSeconds = New String("0", 2 - Len(strSeconds)) & strSeconds
			
			FormatEventDuration = strHours & TIME_SEPARATOR & strMins & TIME_SEPARATOR & strSeconds
		Else
			FormatEventDuration = ""
		End If
		
	End Function
	
	Public Function GetFilteredIDs(ByRef plngExprID As Integer, ByRef pavPromptedValues As Object) As String
		' Return a string describing the record IDs from the given table
		' that satisfy the given criteria.
		Dim sIDSQL As String
    Dim avPrompts(,) As Object
		Dim iDataType As Short
		Dim lngComponentID As Integer
		Dim iLoop As Short
		
		ReDim avPrompts(1, 0)
		
		If IsArray(pavPromptedValues) Then
			ReDim avPrompts(1, UBound(pavPromptedValues, 2))
			
			For iLoop = 0 To UBound(pavPromptedValues, 2)
				' Get the prompt data type.
				'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))
					
					'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPrompts(0, iLoop) = lngComponentID
					
					' NB. Locale to server conversions are done on the client.
					Select Case iDataType
						Case 2
							' Numeric.
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
						Case 3
							' avPrompts.
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
						Case 4
							' Date.
							' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
							' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
							' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
							' THINGS UP.
							'avPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
						Case Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
					End Select
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPrompts(0, iLoop) = 0
				End If
			Next iLoop
		End If
		
		Dim blnOK As Boolean
		
		blnOK = True
		
		ReDim mastrUDFsRequired(0)
		
		blnOK = datGeneral.FilteredIDs(plngExprID, sIDSQL, avPrompts)
		
		' Generate any UDFs that are used in this filter
		If blnOK And gbEnableUDFFunctions Then
			datGeneral.FilterUDFs(plngExprID, mastrUDFsRequired)
		End If
		
		GetFilteredIDs = sIDSQL
		
	End Function
	
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
				CreateASRDev_SysProtects(gADOCon)
			Else
				gADOCon = Value
			End If
			
			SetupTablesCollection()
			
		End Set
	End Property
	
	
	
	Public ReadOnly Property OfficeGetDefaultIndexWord() As Short
		Get
			OfficeGetDefaultIndexWord = mintDefaultIndexWord
		End Get
	End Property
	
	Public ReadOnly Property OfficeGetDefaultIndexExcel() As Short
		Get
			OfficeGetDefaultIndexExcel = mintDefaultIndexExcel
		End Get
	End Property
	
	Public ReadOnly Property OfficeGetCommonDialogFormatsWord() As String
		Get
			OfficeGetCommonDialogFormatsWord = mstrCommonDialogFormatsWord
		End Get
	End Property
	
	Public ReadOnly Property OfficeGetCommonDialogFormatsExcel() As String
		Get
			OfficeGetCommonDialogFormatsExcel = mstrCommonDialogFormatsExcel
		End Get
	End Property
	
	Public ReadOnly Property OfficeGetSaveAsValues() As String
		Get
			OfficeGetSaveAsValues = mstrOfficeSaveAsValues
		End Get
	End Property
	
	Public Function GetPictures(ByRef plngScreenID As Integer, ByRef psTempPath As String) As Object
    Dim avPictures(,) As Object
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		Dim sFileName As String
		
		ReDim avPictures(2, 0)
		
		sSQL = "SELECT DISTINCT ASRSysControls.pictureID, ASRSysPictures.name" & " FROM ASRSysControls" & " INNER JOIN ASRSysPictures ON ASRSysControls.pictureID = ASRSysPictures.pictureID" & " WHERE screenID = " & Trim(Str(plngScreenID)) & " AND controlType = " & Trim(Str(Declarations.ControlTypes.ctlImage))
		
		rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		With rsTemp
			Do While Not .EOF
				sFileName = LoadScreenControlPicture(.Fields("PictureID").Value, psTempPath, .Fields("Name").Value)
				
				If Len(sFileName) > 0 Then
					ReDim Preserve avPictures(2, UBound(avPictures, 2) + 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object avPictures(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPictures(1, UBound(avPictures, 2)) = .Fields("PictureID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object avPictures(2, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPictures(2, UBound(avPictures, 2)) = Mid(sFileName, InStrRev(sFileName, "\") + 1)
				End If
				
				.MoveNext()
			Loop 
			
			.Close()
		End With
		
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetPictures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetPictures = VB6.CopyArray(avPictures)
		
	End Function
	
	Public Function GetBackgroundPicture(ByRef psTempPath As String) As Object
		
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		Dim sFileName As String
		Dim lngPictureID As Short
		
		sSQL = "SELECT DISTINCT ASRSysSystemSettings.settingValue " & "FROM ASRSysSystemSettings " & "WHERE ASRSysSystemSettings.section = 'desktopsetting' " & "     AND  ASRSysSystemSettings.settingKey = 'bitmapid'"
		rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		With rsTemp
			If Not (.BOF And .EOF) Then
				lngPictureID = .Fields("settingValue").Value
			End If
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		sSQL = vbNullString
		sSQL = "SELECT DISTINCT ASRSysPictures.pictureID, ASRSysPictures.name" & " FROM ASRSysPictures " & " WHERE ASRSysPictures.pictureID = " & CStr(lngPictureID)
		rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		With rsTemp
			If Not (.BOF And .EOF) Then
				sFileName = LoadScreenControlPicture(.Fields("PictureID").Value, psTempPath, .Fields("Name").Value)
			End If
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		GetBackgroundPicture = Mid(sFileName, InStrRev(sFileName, "\") + 1)
		
	End Function
	
	Public Function GetBackgroundPosition() As Short
		
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		Dim intBGPos As Short
		
		sSQL = "SELECT DISTINCT ASRSysSystemSettings.settingValue " & "FROM ASRSysSystemSettings " & "WHERE ASRSysSystemSettings.section = 'desktopsetting' " & "     AND  ASRSysSystemSettings.settingKey = 'bitmaplocation'"
		rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		With rsTemp
			If Not (.BOF And .EOF) Then
				intBGPos = .Fields("settingValue").Value
			End If
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		GetBackgroundPosition = intBGPos
		
	End Function
	
  Private Function LoadScreenControlPicture(ByVal plngPictureID As Integer, ByVal psTempPath As String, ByVal psName As String) As String
    ' Read the given picture from the database.
    On Error GoTo ErrorTrap

    Dim iFragments As Short
    Dim iTempFile As Short
    Dim lngOffset As Integer
    Dim lngPictureSize As Integer
    Dim sTempName As String
    Dim sPictureFile As String
    Dim bytChunks() As Byte
    Dim rsPictures As ADODB.Recordset
    Dim sSQL As String

    Const conChunkSize As Short = 2 ^ 14

    sPictureFile = ""

    sSQL = "SELECT picture" & " FROM ASRSysPictures" & " WHERE pictureID = " & Trim(Str(plngPictureID))
    rsPictures = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    With rsPictures
      If Not (.BOF And .EOF) Then
        .MoveFirst()

        'sTempName = GetFileName(psTempPath, psName)
        sTempName = psTempPath & "\" & psName

        iTempFile = 1
        FileOpen(iTempFile, sTempName, OpenMode.Binary, OpenAccess.Write)

        lngPictureSize = .Fields("Picture").ActualSize
        iFragments = lngPictureSize Mod conChunkSize

        ReDim bytChunks(iFragments)

        Do While lngOffset < lngPictureSize
          'UPGRADE_WARNING: Couldn't resolve default property of object rsPictures!Picture.GetChunk(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          bytChunks = .Fields("Picture").GetChunk(conChunkSize)
          lngOffset = lngOffset + conChunkSize
          'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
          FilePut(iTempFile, bytChunks)
        Loop

        FileClose(iTempFile)

        sPictureFile = sTempName
      End If

      .Close()
    End With
    'UPGRADE_NOTE: Object rsPictures may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsPictures = Nothing

TidyUpAndExit:

    LoadScreenControlPicture = sPictureFile
    Exit Function

ErrorTrap:
    sPictureFile = Err.Description
    Resume TidyUpAndExit

  End Function
	
	
	
	
	
	
	
	Private Function GetFileName(ByRef psPath As String, ByRef psName As String) As String
		On Error GoTo ErrorTrap
		
		Dim iIndex As Short
		Dim sFullFileName As String
		Dim sTempFileName As String
		
		sFullFileName = ""
		
		iIndex = 0
		sTempFileName = "_" & psName
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Do While Dir(psPath & "\" & sTempFileName) <> vbNullString
			iIndex = iIndex + 1
			sTempFileName = "_" & Trim(Str(iIndex)) & psName
		Loop 
		
		sFullFileName = psPath & "\" & sTempFileName
		
TidyUpAndExit: 
		GetFileName = sFullFileName
		Exit Function
		
ErrorTrap: 
		sFullFileName = ""
		Resume TidyUpAndExit
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mclsData = New clsDataAccess
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsData = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Sub OfficeInitialise(ByRef intWordVer As Short, ByRef intExcelVer As Short)
		
		OfficeGetCommonDialogFormats("Word", intWordVer)
		OfficeGetCommonDialogFormats("Excel", intExcelVer)
		
		mstrOfficeSaveAsValues = OfficeGetSaveAsValuesForDestin("Word", intWordVer) & "|" & OfficeGetSaveAsValuesForDestin("WordTemplate", intWordVer) & "|" & OfficeGetSaveAsValuesForDestin("Excel", intExcelVer) & "|" & OfficeGetSaveAsValuesForDestin("ExcelTemplate", intExcelVer)
		
	End Sub
	
	
	Private Function OfficeGetSaveAsValuesForDestin(ByRef strDestin As String, ByRef intOfficeVersion As Short) As String
		
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		Dim strFormatField As String
		Dim strOutput As String
		
		On Error GoTo LocalErr
		
		strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007")
		
		
		strSQL = "SELECT * " & "FROM   ASRSysFileFormats " & "WHERE  Destination = '" & Replace(strDestin, "'", "''") & "' " & "  AND  NOT " & strFormatField & " IS NULL " & "ORDER BY ID"
		rsTemp = datGeneral.GetRecords(strSQL)
		
		Do While Not rsTemp.EOF
			If strOutput <> vbNullString Then
				strOutput = strOutput & "|"
			End If
			
			strOutput = strOutput & rsTemp.Fields("Extension").Value & "|" & rsTemp.Fields(strFormatField).Value
			
			rsTemp.MoveNext()
		Loop 
		
		OfficeGetSaveAsValuesForDestin = strOutput
		
		
LocalErr: 
		If Not rsTemp Is Nothing Then
			If rsTemp.State <> ADODB.ObjectStateEnum.adStateClosed Then
				rsTemp.Close()
			End If
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If
		
	End Function
	
	
	Private Function OfficeGetCommonDialogFormats(ByRef strDestin As String, ByRef intOfficeVersion As Short) As String
		
		Dim objSettings As clsSettings
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		Dim strOutput As String
		Dim strFormatField As String
		Dim intCount As Short
		Dim intDefaultFormat As Short
		Dim intDefaultFormatIndex As Short
		
		On Error GoTo LocalErr
		
		objSettings = New clsSettings
		'UPGRADE_WARNING: Couldn't resolve default property of object objSettings.GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		intDefaultFormat = objSettings.GetUserSetting("Output", strDestin & "Format", -1)
		'UPGRADE_NOTE: Object objSettings may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objSettings = Nothing
		
		
		strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007")
		
		strSQL = "SELECT * " & "FROM   ASRSysFileFormats " & "WHERE  Destination = '" & Replace(strDestin, "'", "''") & "' " & "  AND  NOT " & strFormatField & " IS NULL " & "ORDER BY ID"
		rsTemp = datGeneral.GetRecords(strSQL)
		
		intCount = 1
		intDefaultFormatIndex = -1
		strOutput = vbNullString
		Do While Not rsTemp.EOF
			If strOutput <> vbNullString Then
				strOutput = strOutput & "|"
			End If
			
			If rsTemp.Fields(strFormatField).Value = intDefaultFormat Then
				intDefaultFormatIndex = intCount
				
			ElseIf rsTemp.Fields("Default").Value = True Then 
				If intDefaultFormatIndex = -1 Then
					intDefaultFormatIndex = intCount
				End If
				
			End If
			
			strOutput = strOutput & rsTemp.Fields("Description").Value & "|*." & rsTemp.Fields("Extension").Value
			'mlngFileFormats.Add rsTemp.Fields("ID").Value & "|" & rsTemp.Fields("Extension").Value
			
			intCount = intCount + 1
			rsTemp.MoveNext()
		Loop 
		
		
		If strDestin = "Word" Then
			mstrCommonDialogFormatsWord = strOutput
			mintDefaultIndexWord = intDefaultFormatIndex
		Else
			mstrCommonDialogFormatsExcel = strOutput
			mintDefaultIndexExcel = intDefaultFormatIndex
		End If
		
		
LocalErr: 
		If Not rsTemp Is Nothing Then
			If rsTemp.State <> ADODB.ObjectStateEnum.adStateClosed Then
				rsTemp.Close()
			End If
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If
		
	End Function
End Class