Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsOLE_NET.clsOLE")> Public Class clsOLE
	
	Private Declare Function WNetGetConnection Lib "mpr.dll"  Alias "WNetGetConnectionA"(ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
	
	Private mstrTempLocationPhysical As String
	Private mstrTempLocationUNC As String
	Private mstrTempFileToDelete As String
	
	Private miOLEType As Short
	Private mstrDisplayFileName As String
	Private mstrFileName As String
	Private mstrPath As String
	Private mstrUNC As String
	Private mstrDocumentSize As String
	Private mstrFileVersion As String
	Private mstrFileSize As String
	Private mstrFileCreateDate As String
	Private mstrFileModifyDate As String
	Private mstrFileDateAccessed As String
	
	Private mobjFileSystem As Scripting.FileSystemObject
	
	Private mbUseEncryption As Boolean
	Private mbUseFileSecurity As Boolean
	Private mstrCurrentSessionKey As String
	Private mstrCurrentUser As String
	Private mstrProcessUser As String
	
	Private mobjStream As ADODB.Stream
	Private mstrDummyConnectionString As String
	
	' Holds the names of the OLE files for this record session
	Private mastrOLEFilesInThisSession() As String
	
	' Do we use encryption?
	Public WriteOnly Property UseEncryption() As Boolean
		Set(ByVal Value As Boolean)
			mbUseEncryption = Value
		End Set
	End Property
	
	' Do we use file security?
	Public WriteOnly Property UseFileSecurity() As Boolean
		Set(ByVal Value As Boolean)
			mbUseFileSecurity = Value
		End Set
	End Property
	
	' The current session key (used for encryption purposes)
	Public WriteOnly Property CurrentSessionKey() As String
		Set(ByVal Value As String)
			mstrCurrentSessionKey = Value
		End Set
	End Property
	
	' The current user (used for security purposes)
	Public WriteOnly Property CurrentUser() As String
		Set(ByVal Value As String)
			mstrCurrentUser = Value
		End Set
	End Property
	
	' Path in which temporary documents are to be created (physical directory on the server)
	Public WriteOnly Property TempLocationPhysical() As String
		Set(ByVal Value As String)
			mstrTempLocationPhysical = Value
		End Set
	End Property
	
	' The current UNC of the asp page being run
	Public WriteOnly Property TempLocationUNC() As String
		Set(ByVal Value As String)
			mstrTempLocationUNC = Value
		End Set
	End Property
	
	
	Public Property OLEType() As Short
		Get
			OLEType = miOLEType
		End Get
		Set(ByVal Value As Short)
			miOLEType = Value
		End Set
	End Property
	
	
	Public Property FileName() As String
		Get
			' If linked file return proper link
			If miOLEType = 2 Then
				FileName = mstrTempLocationUNC & GetFileNameOnly(mstrFileName)
			Else
				FileName = mstrFileName
			End If
			
		End Get
		Set(ByVal Value As String)
			mstrFileName = Value
		End Set
	End Property
	
	
	Public Property DisplayFilename() As String
		Get
			DisplayFilename = mstrDisplayFileName
		End Get
		Set(ByVal Value As String)
			mstrDisplayFileName = Value
		End Set
	End Property
	
	Public ReadOnly Property DocumentModifyDate() As String
		Get
			DocumentModifyDate = mstrFileModifyDate
		End Get
	End Property
	
	' Returns the size of the document in a nice formatted method
	Public ReadOnly Property DocumentSize() As String
		Get
			Select Case Len(mstrDocumentSize)
				Case Is < 5
					DocumentSize = mstrDocumentSize & " bytes"
					
				Case Is < 7
					DocumentSize = Mid(mstrDocumentSize, 1, Len(mstrDocumentSize) - 3) & "KB"
					
				Case 7
					DocumentSize = Mid(mstrDocumentSize, 1, 1) & "." & Mid(mstrDocumentSize, 2, 2) & "MB"
					
				Case Is < 10
					DocumentSize = Mid(mstrDocumentSize, 1, Len(mstrDocumentSize) - 6) & "MB"
					
			End Select
			
		End Get
	End Property
	
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
			Else
				gADOCon = Value
			End If
			
			' We need to create a dummy connection string as the one passed in from the Intranet doesn't read the image data type.
			' I don't know why...
			'  mstrDummyConnectionString = Replace(gADOCon.ConnectionString, "MSDASQL", "SQLOLEDB") & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=32767;Use Encryption for Data=False;Tag with column collation when possible=False"
			'  mstrDummyConnectionString = Replace(mstrDummyConnectionString, "Mode=ReadWrite;", "")
			'  mstrDummyConnectionString = Replace(mstrDummyConnectionString, "APP=HR Pro Intranet;", "APP=HR Pro Intranet Embedding;")
			'  mstrDummyConnectionString = Replace(mstrDummyConnectionString, "APP=HR Pro Self-service Intranet;", "APP=HR Pro Intranet Embedding;")
			
			'Changed for Native Client
			mstrDummyConnectionString = gADOCon.ConnectionString & ";Pooling=False;DataTypeCompatibility=80"
			mstrDummyConnectionString = Replace(mstrDummyConnectionString, "Application Name=OpenHR Intranet;", "Application Name=OpenHR Intranet Embedding;")
			mstrDummyConnectionString = Replace(mstrDummyConnectionString, "Application Name=OpenHR Self-service Intranet;", "Application Name=OpenHR Intranet Embedding;")
			
		End Set
	End Property
	
	Public WriteOnly Property OLEFileSize() As String
		Set(ByVal Value As String)
			mstrFileSize = Value
		End Set
	End Property
	
	Public WriteOnly Property OLEModifiedDate() As String
		Set(ByVal Value As String)
			mstrFileModifyDate = Value
		End Set
	End Property
	
	
	Public Function CreateOLEDocument(ByRef plngRecordID As Object, ByRef plngColumnID As Integer, ByRef pstrRealSource As String) As String
		
		On Error GoTo ErrorTrap
		
		Dim bOK As Boolean
		Dim sSQL As String
		Dim objDummyConnection As ADODB.Connection
		Dim rsDocument As ADODB.Recordset
		
		Dim strTempFile As String
		Dim strProperties As String
		Dim strColumnName As String
		Dim objTextStream As Scripting.TextStream
		
		bOK = True
		
		objDummyConnection = New ADODB.Connection
		'UPGRADE_NOTE: Object mobjStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjStream = Nothing
		
		' Open a temporary connection string to stream the data
		objDummyConnection.Open(mstrDummyConnectionString)
		
		' New record - thus no stream will exist
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If plngRecordID = 0 Then
			CreateOLEDocument = ""
			GoTo TidyUpAndExit
		End If
		
		strColumnName = datGeneral.GetColumnName(CInt(plngColumnID))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID
		
		rsDocument = New ADODB.Recordset
		rsDocument.Open(sSQL, objDummyConnection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		With rsDocument
			.MoveFirst()
			
			If mobjStream Is Nothing Then
				mobjStream = New ADODB.Stream
			End If
			
			If mobjStream.State <> ADODB.ObjectStateEnum.adStateOpen Then
				mobjStream.Open()
				mobjStream.Type = ADODB.StreamTypeEnum.adTypeBinary
			End If
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsDocument.Fields(strColumnName).Value) Then
				mobjStream.Write(rsDocument.Fields(strColumnName).Value)
			Else
				CreateOLEDocument = ""
				miOLEType = 3
				mstrDisplayFileName = ""
				GoTo TidyUpAndExit
			End If
			
			If mobjStream.Size > 0 Then
				strTempFile = GetTmpFName
				mobjStream.SaveToFile(strTempFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				objTextStream = mobjFileSystem.OpenTextFile(strTempFile, Scripting.IOMode.ForReading)
				strProperties = Trim(objTextStream.Read(400))
				
				miOLEType = Val(Mid(strProperties, 9, 2))
				mstrDisplayFileName = Trim(GetFileNameOnly(Mid(strProperties, 11, 70)))
				mstrFileName = IIf(miOLEType = 2, GetTmpFName, mstrDisplayFileName)
				mstrPath = Trim(Mid(strProperties, 81, 210))
				mstrUNC = Trim(Mid(strProperties, 291, 60))
				mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
				mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
				mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))
				
				objTextStream.Close()
				
				' Generate the file if it's not linked
				If miOLEType = 2 Then
					mstrFileName = GenerateDocumentFromStream
					mstrTempFileToDelete = mstrFileName
				Else
					mstrFileName = mstrUNC & mstrPath & "\" & mstrFileName
					mstrTempFileToDelete = ""
				End If
				
			End If
			
		End With
		
TidyUpAndExit: 
		If Not rsDocument.State = ADODB.ObjectStateEnum.adStateClosed Then
			rsDocument.Close()
		End If
		
		If Not objDummyConnection.State = ADODB.ObjectStateEnum.adStateClosed Then
			objDummyConnection.Close()
		End If
		
ExitFunction: 
		'UPGRADE_NOTE: Object rsDocument may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsDocument = Nothing
		'UPGRADE_NOTE: Object objDummyConnection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDummyConnection = Nothing
		'UPGRADE_NOTE: Object objTextStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTextStream = Nothing
		CreateOLEDocument = mstrFileName
		
		Exit Function
		
ErrorTrap: 
		mstrFileName = ""
		mstrDisplayFileName = ""
		ProgramError("CreateOLEDocument", Err, Erl())
		
		Resume ExitFunction
		
	End Function
	
	Public Function CloseStream() As Boolean
		mobjStream.Close()
	End Function
	
	' Save the document part out to a file
	Private Function GenerateDocumentFromStream() As String
		On Error GoTo generate_error
		
		Dim objEncryption As clsEncryption
		Dim objDocumentStream As ADODB.Stream
		Dim strGeneratedFile As String
		Dim strAccessibleFile As String
		
		' Set file paths
		If miOLEType = 2 And mbUseEncryption Then
			strGeneratedFile = GetTmpFName()
			strAccessibleFile = mstrTempLocationPhysical & GetFileNameOnly(mstrFileName)
		Else
			strGeneratedFile = mstrTempLocationPhysical & GetFileNameOnly(mstrFileName)
			strAccessibleFile = strGeneratedFile
		End If
		
		' Set the stream state
		If mobjStream.State = ADODB.ObjectStateEnum.adStateClosed Then
			mobjStream.Open()
			mobjStream.Type = ADODB.StreamTypeEnum.adTypeBinary
		End If
		
		' Setup new document stream
		objDocumentStream = New ADODB.Stream
		objDocumentStream.Type = ADODB.StreamTypeEnum.adTypeBinary
		objDocumentStream.Open()
		
		' Copy out the document part of the stream
		mobjStream.Position = 400
		mobjStream.CopyTo(objDocumentStream, mobjStream.Size - 400)
		objDocumentStream.SaveToFile(strGeneratedFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
		
		' Encrypt file if necessary
		If miOLEType = 2 And mbUseEncryption Then
			objEncryption = New clsEncryption
			objEncryption.EncryptFile(strGeneratedFile, strAccessibleFile, True, "230678" & mstrCurrentSessionKey)
		End If
		
		' Set file security if necessary
		If miOLEType = 2 And mbUseFileSecurity Then
			SetAccess(mstrCurrentUser, strAccessibleFile, GENERIC_READ Or GENERIC_EXECUTE Or GENERIC_WRITE Or DELETE)
			SetAccess(mstrProcessUser, strAccessibleFile, GENERIC_READ Or GENERIC_EXECUTE Or GENERIC_WRITE Or DELETE)
			SetAccess("Everyone", strAccessibleFile, 0)
		End If
		
		' Return the accessible document
		GenerateDocumentFromStream = strAccessibleFile
		
		objDocumentStream.Close()
		
		Kill(strGeneratedFile)
		
generate_error_handler: 
		'UPGRADE_NOTE: Object objDocumentStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDocumentStream = Nothing
		'UPGRADE_NOTE: Object objEncryption may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objEncryption = Nothing
		
		Exit Function
		
generate_error: 
		
		ProgramError("GenerateDocumentFromStream", Err, Erl())
		Resume generate_error_handler
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		'Set datData = New clsGeneral
		mobjStream = New ADODB.Stream
		mobjFileSystem = New Scripting.FileSystemObject
		
		miOLEType = 3
		mstrFileName = ""
		
		mstrCurrentSessionKey = ""
		mstrCurrentUser = ""
		mstrProcessUser = Environ("USERNAME")
		mbUseEncryption = False
		mbUseFileSecurity = False
		
		ReDim mastrOLEFilesInThisSession(0)
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	' Commit the file back to the database
	Public Function SaveStream(ByRef plngRecordID As Integer, ByRef plngColumnID As Integer, ByRef pstrRealSource As String, ByRef pbReadOLEDirect As Boolean) As Boolean
		
		On Error GoTo ErrorTrap
		
		Dim bOK As Boolean
		Dim cmADO As ADODB.Command
		Dim pmADO As ADODB.Parameter
		
		Dim objFile As ADODB.Stream
		
		Dim objEncryption As clsEncryption
		Dim objPropertiesStream As Scripting.TextStream
		Dim strTempFilePath As String
		Dim strTempFileName As String
		Dim strFileName As String
		Dim strUNC As String
		Dim strPath As String
		Dim strOLEType As String
		Dim strEmbedFileName As String
		Dim bUpdateField As Boolean
		Dim test As Boolean
		
		bOK = True
		bUpdateField = False
		
		' Is there a file attached?
		If mstrFileName <> "" Then
			
			If mbUseEncryption Then
				strTempFilePath = Space(1024)
				Call GetTempPath(1024, strTempFilePath)
				strEmbedFileName = Left(strTempFilePath, InStr(strTempFilePath, Chr(0)) - 1) & GetFileNameOnly(mstrFileName)
				objEncryption = New clsEncryption
				test = objEncryption.DecryptFile(mstrFileName, strEmbedFileName, True, "230678" & mstrCurrentSessionKey)
				'UPGRADE_NOTE: Object objEncryption may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objEncryption = Nothing
			Else
				If miOLEType = 3 Then
					strEmbedFileName = mstrFileName
				Else
					strEmbedFileName = IIf(pbReadOLEDirect, mstrFileName, mstrTempLocationPhysical & GetFileNameOnly(mstrFileName))
				End If
			End If
			
			' Save the document information to file to read in.
			strTempFileName = GetTmpFName
			
			strOLEType = Trim(Str(miOLEType))
			
			
			' Create a textfile of properties from the passed in file
			If miOLEType = 2 Then
				strUNC = GetUNCOnly(strEmbedFileName)
				strPath = GetPathOnly(strEmbedFileName, True)
				strFileName = mstrDisplayFileName
			Else
				strUNC = Trim(GetUNCOnly(strEmbedFileName))
				strPath = GetPathOnly(strEmbedFileName, True)
				strFileName = mstrDisplayFileName
				strPath = Replace(strPath, strUNC, "")
			End If
			
			objPropertiesStream = mobjFileSystem.OpenTextFile(strTempFileName, Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
			objPropertiesStream.Write("<<V002>>") ' Structure version info
			objPropertiesStream.Write(strOLEType & Space(2 - Len(strOLEType)))
			objPropertiesStream.Write(strFileName & Space(70 - Len(strFileName)))
			objPropertiesStream.Write(strPath & Space(210 - Len(strPath)))
			objPropertiesStream.Write(strUNC & Space(60 - Len(strUNC)))
			objPropertiesStream.Write(mstrFileSize & Space(10 - Len(mstrFileSize)))
			objPropertiesStream.Write(Space(20))
			objPropertiesStream.Write(mstrFileModifyDate & Space(20 - Len(mstrFileModifyDate)))
			objPropertiesStream.Close()
			
			If mobjStream.State = ADODB.ObjectStateEnum.adStateOpen Then
				mobjStream.Close()
				mobjStream.Open()
			Else
				mobjStream.Open()
				mobjStream.Type = ADODB.StreamTypeEnum.adTypeBinary
			End If
			
			' Load the properties header
			mobjStream.LoadFromFile(strTempFileName)
			mobjStream.Position = mobjStream.Size
			
			' If embedded file tack onto the end of the stream
			If miOLEType = 2 Then ' Embedded
				mobjStream.Position = mobjStream.Size
				objFile = New ADODB.Stream
				objFile.Open()
				objFile.Type = ADODB.StreamTypeEnum.adTypeBinary
				objFile.LoadFromFile(strEmbedFileName)
				mobjStream.Write(objFile.Read)
				objFile.Close()
				'UPGRADE_NOTE: Object objFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objFile = Nothing
			End If
			
			mobjFileSystem.DeleteFile(strTempFileName, True)
			
			' Delete the existing file
			If Not pbReadOLEDirect And miOLEType = 2 Then
				Kill(mstrFileName)
				Kill(strEmbedFileName)
			End If
			
			' Flag the update to occur
			bUpdateField = True
			
		End If
		
		' Fling the stream into the database. Use a stored procedure because we may be accessing the view with a UDF attached.
		cmADO = New ADODB.Command
		With cmADO
			
			.CommandText = "spASRUpdateOLEField_" & plngColumnID
			.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
			.CommandTimeout = 0
			.ActiveConnection = gADOCon
			
			pmADO = .CreateParameter("currentID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
			.Parameters.Append(pmADO)
			pmADO.Value = plngRecordID
			
			pmADO = .CreateParameter("UploadFile", ADODB.DataTypeEnum.adLongVarBinary, ADODB.ParameterDirectionEnum.adParamInput, -1)
			.Parameters.Append(pmADO)
			
			If mstrFileName <> "" Then
				If mobjStream.State = ADODB.ObjectStateEnum.adStateClosed Then
					mobjStream.Open()
				End If
				
				If mobjStream.Size > 0 Then
					mobjStream.Position = 0
					'UPGRADE_WARNING: Couldn't resolve default property of object mobjStream.Read. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pmADO.Value = mobjStream.Read
				Else
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					pmADO.Value = System.DBNull.Value
				End If
				
			Else
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				pmADO.Value = System.DBNull.Value
			End If
			
		End With
		
		cmADO.Execute()
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object pmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pmADO = Nothing
		'UPGRADE_NOTE: Object cmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmADO = Nothing
		
		SaveStream = bOK
		Exit Function
		
ErrorTrap: 
		bOK = False
		ProgramError("SaveStream", Err, Erl())
		
		Resume TidyUpAndExit
		
	End Function
	
	' Extracts just the filename from a path
	Private Function GetFileNameOnly(ByRef pstrFilePath As String) As String
		
		On Error GoTo ErrorTrap
		
		Dim astrPath() As String
		astrPath = Split(pstrFilePath, "\")
		GetFileNameOnly = Trim(astrPath(UBound(astrPath)))
		Exit Function
		
ErrorTrap: 
		ProgramError("GetFileNameOnly", Err, Erl())
		GetFileNameOnly = ""
		
	End Function
	
	Private Function GetUNCOnly(ByVal pstrFileName As String) As String
		
		On Error GoTo GetUNCPath_Err
		
		Dim strMsg As String
		Dim lngReturn As Integer
		Dim strLocalName As String
		Dim strRemoteName As String
		Dim lngRemoteName As Integer
		Dim strUNCPath As String
		
		strLocalName = GetDriveOnly(pstrFileName)
		strRemoteName = New String(Chr(32), 255)
		lngRemoteName = Len(strRemoteName)
		
		'Attempt to grab UNC
		lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)
		
		If lngReturn = 0 Then
			GetUNCOnly = Trim(Replace(strRemoteName, Chr(0), ""))
			
			' UNC passed in
		ElseIf lngReturn = 1200 Then 
			GetUNCOnly = GetUNCFromPath(pstrFileName)
			
			' Local path
		ElseIf lngReturn = 2250 Then 
			GetUNCOnly = GetDriveOnly(pstrFileName)
		Else
			GetUNCOnly = Trim(strLocalName)
		End If
		
GetUNCPath_End: 
		Exit Function
		
GetUNCPath_Err: 
		GetUNCOnly = Trim(strLocalName)
		ProgramError("GetUNCOnly", Err, Erl())
		
		Resume GetUNCPath_End
	End Function
	
	' Extracts the path from a given filename (with a final "\" at the end")
	Public Function GetPathOnly(ByRef pstrFilePath As String, ByRef pbStripDriveLetter As Boolean) As String
		On Error GoTo path_error
		
		Dim l As Short
		Dim tempchar As String
		Dim strPath As String
		
		l = Len(pstrFilePath)
		
		While l > 0
			tempchar = Mid(pstrFilePath, l, 1)
			If tempchar = "\" Then
				strPath = Mid(pstrFilePath, 1, l - 1)
				
				' Strip off drive letter
				If pbStripDriveLetter And Mid(strPath, 2, 1) = ":" Then
					strPath = Mid(strPath, 3, Len(strPath))
				End If
				
				GetPathOnly = strPath
				
				Exit Function
			End If
			l = l - 1
		End While
		
		Exit Function
		
path_error: 
		ProgramError("GetPathOnly", Err, Erl())
		
	End Function
	
	Private Function GetDriveOnly(ByVal pstrFileName As String) As String
		On Error GoTo getdrive_error
		
		If Mid(pstrFileName, 2, 1) = ":" Then
			GetDriveOnly = Mid(pstrFileName, 1, 1) & ":"
		Else
			GetDriveOnly = ""
		End If
		
		Exit Function
		
getdrive_error: 
		ProgramError("getDriveOnly", Err, Erl())
		
	End Function
	
	' Extracts the path from a given filename
	Public Function GetWebsiteName(ByRef pstrUNC As String) As String
		On Error GoTo Website_error
		
		Dim l As Short
		Dim tempchar As String
		Dim strPath As String
		
		l = Len(pstrUNC)
		
		While l > 0
			tempchar = Mid(pstrUNC, l, 1)
			If tempchar = "/" Then
				strPath = Mid(pstrUNC, 1, l - 1)
				
				GetWebsiteName = strPath
				
				Exit Function
			End If
			l = l - 1
		End While
		
		Exit Function
		
Website_error: 
		ProgramError("GetWebsiteName", Err, Erl())
		
	End Function
	
	Public Function UNCAndPath() As String
		
		If miOLEType = 3 Then
			UNCAndPath = GetUNCOnly(mstrFileName) & GetPathOnly(mstrPath, True)
		Else
			UNCAndPath = ""
		End If
		
		
	End Function
	
	' Extracts the unc from a given path
	Private Function GetUNCFromPath(ByRef pstrFilePath As String) As String
		On Error GoTo getUNC_error
		
		Dim l As Short
		Dim tempchar As String
		Dim strPath As String
		
		' Is file path passed as a unc or drive map
		If InStr(1, pstrFilePath, "\\", CompareMethod.Text) > 0 Then
			strPath = Left(pstrFilePath, InStr(3, pstrFilePath, "\", CompareMethod.Text))
			GetUNCFromPath = Left(pstrFilePath, InStr(Len(strPath) + 1, pstrFilePath, "\", CompareMethod.Text) - 1)
		ElseIf InStr(1, Left(pstrFilePath, 2), ":", CompareMethod.Text) Then 
			l = Len(pstrFilePath)
			
			While l > 0
				tempchar = Mid(pstrFilePath, l, 1)
				If tempchar = "\" Then
					strPath = Mid(pstrFilePath, 1, l - 1)
					
					GetUNCFromPath = strPath
					Exit Function
				End If
				l = l - 1
			End While
			
		End If
		
		Exit Function
		
getUNC_error: 
		ProgramError("GetUNCFromPath", Err, Erl())
		
	End Function
	
	' Called from VBScript to populate the buttons
	Public Function GetPropertiesFromStream(ByRef plngRecordID As Object, ByRef plngColumnID As Object, ByRef pstrRealSource As String) As String
		
		On Error GoTo ErrorTrap
		
		Dim bOK As Boolean
		
		Dim objDocumentStream As ADODB.Stream
		Dim objStream As ADODB.Stream
		Dim rsDocument As ADODB.Recordset
		Dim objDummyConnection As ADODB.Connection
		
		Dim strAccessibleFile As String
		Dim bIsPhoto As Boolean
		Dim strTempFile As String
		Dim strProperties As String
		Dim objTextStream As Scripting.TextStream
		
		Dim sSQL As String
		Dim strColumnName As String
		
		Dim iOLEType As Short
		Dim strFileName As String
		Dim strPath As String
		Dim strUNC As String
		
		' New record - thus no stream will exist
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If plngRecordID = 0 Then
			GetPropertiesFromStream = ""
			Exit Function
		End If
		
		objStream = New ADODB.Stream
		objStream.Open()
		objStream.Type = ADODB.StreamTypeEnum.adTypeBinary
		
		strTempFile = GetTmpFName
		bOK = True
		'UPGRADE_WARNING: Couldn't resolve default property of object plngColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bIsPhoto = datGeneral.IsPhotoDataType(CInt(plngColumnID))
		
		' Open a temporary connection string to stream the data
		objDummyConnection = New ADODB.Connection
		objDummyConnection.Open(mstrDummyConnectionString)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object plngColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strColumnName = datGeneral.GetColumnName(CInt(plngColumnID))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID
		
		rsDocument = New ADODB.Recordset
		rsDocument.Open(sSQL, objDummyConnection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		With rsDocument
			.MoveFirst()
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsDocument.Fields(strColumnName).Value) Then
				objStream.Write(rsDocument.Fields(strColumnName).Value)
			Else
				GetPropertiesFromStream = ""
				GoTo TidyUpAndExit
			End If
			
			If objStream.Size > 0 Then
				objStream.SaveToFile(strTempFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				objTextStream = mobjFileSystem.OpenTextFile(strTempFile, Scripting.IOMode.ForReading)
				strProperties = Trim(objTextStream.Read(400))
				
				iOLEType = Val(Mid(strProperties, 9, 2))
				strFileName = Trim(GetFileNameOnly(Mid(strProperties, 11, 70)))
				strPath = Trim(Mid(strProperties, 81, 210))
				strUNC = Trim(Mid(strProperties, 291, 60))
				objTextStream.Close()
				
				' Generate the properties display tag
				If bIsPhoto Then
					'If it's an embedded photo we'll need to generate it
					If iOLEType = 2 Then
						'strAccessibleFile = mstrTempLocationPhysical & GetFileNameOnly(strFileName)
						strAccessibleFile = mstrTempLocationPhysical & VB6.Format(Today, "yymmdd") & GetFileNameOnly(GetTmpFName)
						objDocumentStream = New ADODB.Stream
						objDocumentStream.Type = ADODB.StreamTypeEnum.adTypeBinary
						objDocumentStream.Open()
						objStream.Position = 400
						objStream.CopyTo(objDocumentStream, objStream.Size - 400)
						objDocumentStream.SaveToFile(strAccessibleFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
						GetPropertiesFromStream = strAccessibleFile & "::EMBEDDED_OLE_DOCUMENT::"
						
						ReDim Preserve mastrOLEFilesInThisSession(UBound(mastrOLEFilesInThisSession) + 1)
						mastrOLEFilesInThisSession(UBound(mastrOLEFilesInThisSession)) = strAccessibleFile
						
					Else
						GetPropertiesFromStream = strUNC & strPath & "\" & strFileName & "::LINKED_OLE_DOCUMENT::"
					End If
				Else
					If iOLEType = 2 Then
						GetPropertiesFromStream = strFileName & "::EMBEDDED_OLE_DOCUMENT::"
					Else
						GetPropertiesFromStream = strUNC & strPath & "\" & strFileName & "::LINKED_OLE_DOCUMENT::"
					End If
					
				End If
				
			End If
			
			Kill(strTempFile)
			
		End With
		
TidyUpAndExit: 
		If Not rsDocument.State = ADODB.ObjectStateEnum.adStateClosed Then
			rsDocument.Close()
		End If
		
		If Not objDummyConnection.State = ADODB.ObjectStateEnum.adStateClosed Then
			objDummyConnection.Close()
		End If
		
ExitFunction: 
		'UPGRADE_NOTE: Object objStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objStream = Nothing
		'UPGRADE_NOTE: Object objDocumentStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDocumentStream = Nothing
		'UPGRADE_NOTE: Object objDummyConnection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDummyConnection = Nothing
		'UPGRADE_NOTE: Object rsDocument may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsDocument = Nothing
		Exit Function
		
ErrorTrap: 
		GetPropertiesFromStream = ""
		ProgramError("GetPropertiesFromStream", Err, Erl())
		Resume TidyUpAndExit
		
	End Function
	
	' Delete the temporary file
	Public Sub DeleteTempFile()
		
		On Error GoTo ErrorTrap
		
		If mobjFileSystem.FileExists(mstrTempFileToDelete) Then
			Kill(mstrTempFileToDelete)
		End If
		
TidyUpAndExit: 
		Exit Sub
		
ErrorTrap: 
		ProgramError("DeleteTempFile", Err, Erl())
		
	End Sub
	
	' Delete the temporary files created in this session
	Public Sub CleanupOLEFiles()
		
		On Error GoTo ErrorTrap
		
		Dim iCount As Short
		
		For iCount = LBound(mastrOLEFilesInThisSession) To UBound(mastrOLEFilesInThisSession)
			If mobjFileSystem.FileExists(mastrOLEFilesInThisSession(iCount)) Then
				Kill(mastrOLEFilesInThisSession(iCount))
			End If
		Next iCount
		
TidyUpAndExit: 
		ReDim mastrOLEFilesInThisSession(0)
		Exit Sub
		
ErrorTrap: 
		ProgramError("CleanupOLEFiles", Err, Erl())
		Resume TidyUpAndExit
		
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		' Erase any files that have got left over from this session
		CleanupOLEFiles()
		
		'UPGRADE_NOTE: Object mobjFileSystem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjFileSystem = Nothing
		'UPGRADE_NOTE: Object mobjStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjStream = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class