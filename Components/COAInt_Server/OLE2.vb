' This class temporarily replaces clsOLE whilst we decide on what to do with this functionality. Only the properties that are called externally have had code stumps
' created. Original code is held in clsOLE which is not inclued in the project file as it errors and needs rework to handle encryption and different file access methods.
' Possibly suggest rewriting the class rather than upgrading.

Imports System.IO
Imports System.Text

Public Class Ole
	Private _mstrTempLocationPhysical As String
	' Holds the names of the OLE files for this record session
	Private _mastrOleFilesInThisSession() As String
	Private mstrTempLocationUNC As String

	Private _miOLEType As Short
	Private _mstrDisplayFileName As String
	Private _mstrFileName As String
	Private _mstrPath As String
	Private _mstrUnc As String
	Private _mstrDocumentSize As String
	Private _mstrFileSize As String
	Private _mstrFileCreateDate As String
	Private _mstrFileModifyDate As String
	Private _mstrDummyConnectionString As String
	Private _mbUseEncryption As Boolean
	Private _mobjStream As ADODB.Stream
	Private _mfileToEmbed As Byte()

	' Do we use encryption?
	Public WriteOnly Property UseEncryption() As Boolean
		Set(ByVal value As Boolean)
			_mbUseEncryption = value
		End Set
	End Property

	' Do we use file security?
	Public WriteOnly Property UseFileSecurity() As Boolean
		Set(ByVal value As Boolean)
		End Set
	End Property

	' The current session key (used for encryption purposes)
	Public WriteOnly Property CurrentSessionKey() As String
		Set(ByVal value As String)
		End Set
	End Property

	' The current user (used for security purposes)
	Public WriteOnly Property CurrentUser() As String
		Set(ByVal value As String)
		End Set
	End Property

	' The current UNC of the asp page being run
	Public WriteOnly Property TempLocationUNC() As String
		Set(ByVal value As String)
			mstrTempLocationUNC = value
		End Set
	End Property

	Public WriteOnly Property OLEFileSize() As String
		Set(ByVal value As String)
			_mstrFileSize = value
		End Set
	End Property

	Public WriteOnly Property OLEModifiedDate() As String
		Set(ByVal value As String)
			Dim localDate As DateTime
			Try
				localDate = DateTime.Parse(value).ToShortDateString()

			Catch ex As Exception
				localDate = ""
			End Try

			_mstrFileModifyDate = localDate

		End Set
	End Property

	Public Property OLEType() As Short
		Get
			OLEType = _miOLEType
		End Get
		Set(ByVal value As Short)
			_miOLEType = value
		End Set
	End Property

	' Path in which temporary documents are to be created (physical directory on the server)
	Public WriteOnly Property TempLocationPhysical() As String
		Set(ByVal value As String)
			_mstrTempLocationPhysical = value
		End Set
	End Property

	Public WriteOnly Property Connection() As ADODB.Connection
		Set(ByVal value As ADODB.Connection)
			gADOCon = value

			_mstrDummyConnectionString = gADOCon.ConnectionString & ";Pooling=False;DataTypeCompatibility=80"
			_mstrDummyConnectionString = Replace(_mstrDummyConnectionString, "Application Name=OpenHR Intranet;", "Application Name=OpenHR Intranet Embedding;")
			_mstrDummyConnectionString = Replace(_mstrDummyConnectionString, "Application Name=OpenHR Self-service Intranet;", "Application Name=OpenHR Intranet Embedding;")

		End Set
	End Property

	Public Sub CleanupOleFiles()
	End Sub

	Public Function CreateOLEDocument(ByRef plngRecordID As Object, ByRef plngColumnID As Object, ByRef pstrRealSource As Object) As Byte()

		On Error GoTo ErrorTrap

		Dim sSQL As String
		Dim objDummyConnection As ADODB.Connection
		Dim rsDocument As ADODB.Recordset

		Dim strTempFile As String
		Dim strProperties As String
		Dim strColumnName As String
		'Dim objTextStream As Scripting.TextStream
		Dim objTextStream As FileStream

		Dim abtImage As Byte()
		ReDim abtImage(0)
		Dim responseFile As Byte()

		objDummyConnection = New ADODB.Connection
		'UPGRADE_NOTE: Object mobjStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		_mobjStream = Nothing

		' Open a temporary connection string to stream the data
		objDummyConnection.Open(_mstrDummyConnectionString)

		' New record - thus no stream will exist
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If plngRecordID = 0 Then
			'CreateOLEDocument = ""
			GoTo TidyUpAndExit
		End If

		strColumnName = datGeneral.GetColumnName(CInt(plngColumnID))

		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID

		rsDocument = New ADODB.Recordset
		rsDocument.Open(sSQL, objDummyConnection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

		With rsDocument
			.MoveFirst()

			If _mobjStream Is Nothing Then
				_mobjStream = New ADODB.Stream
			End If

			If _mobjStream.State <> ADODB.ObjectStateEnum.adStateOpen Then
				_mobjStream.Open()
				_mobjStream.Type = ADODB.StreamTypeEnum.adTypeBinary
			End If

			abtImage = CType(rsDocument.Fields(strColumnName).Value, Byte())
			Dim fFileOK = (abtImage.GetLength(0) > 0)

			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDBNull(rsDocument.Fields(strColumnName).Value) Then
				_mobjStream.Write(rsDocument.Fields(strColumnName).Value)
			Else
				'CreateOLEDocument = ""
				_miOLEType = 3
				_mstrDisplayFileName = ""
				GoTo TidyUpAndExit
			End If

			If _mobjStream.Size > 0 Then
				strTempFile = GetTmpFName()
				_mobjStream.SaveToFile(strTempFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)

				'objTextStream = mobjFileSystem.OpenTextFile(strTempFile, Scripting.IOMode.ForReading)
				'strProperties = Trim(objTextStream.Read(400))				
				objTextStream = File.OpenRead(strTempFile)
				Dim b As Byte() = New Byte(399) {}
				' strProperties = Trim(objTextStream.Read(400))				
				Dim temp As New UTF8Encoding(True)
				objTextStream.Read(b, 0, b.Length)
				strProperties &= temp.GetString(b)

				responseFile = New Byte((objTextStream.Length - 1) - 400) {}
				objTextStream.Read(responseFile, 0, responseFile.Length)

				'Dim outputFile As Byte() = New Byte(responseFile.Length - 400) {}
				'Array.Copy(responseFile, 400, outputFile, 0, responseFile.Length - 400)


				_miOLEType = Val(Mid(strProperties, 9, 2))
				_mstrDisplayFileName = Trim(Path.GetFileName(Mid(strProperties, 11, 70)))
				_mstrFileName = IIf(_miOLEType = 2, GetTmpFName, _mstrDisplayFileName)
				_mstrPath = Trim(Mid(strProperties, 81, 210))
				_mstrUnc = Trim(Mid(strProperties, 291, 60))
				_mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
				_mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
				_mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))

				objTextStream.Close()

				' Generate the file if it's not linked
				If _miOLEType = 2 Then
					' TODO: content stream to client - no holding area.
					' mstrFileName = GenerateDocumentFromStream
				Else
					_mstrFileName = _mstrUnc & _mstrPath & "\" & _mstrFileName
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
		' CreateOLEDocument = mstrFileName

		Return responseFile

		Exit Function

ErrorTrap:
		_mstrFileName = ""
		_mstrDisplayFileName = ""
		ProgramError("CreateOLEDocument", Err, Erl())

		Resume ExitFunction

	End Function

	Public Function CloseStream() As Boolean
		_mobjStream.Close()
	End Function

	Public Function GetPropertiesFromStream(ByRef plngRecordID As Object, ByRef plngColumnID As Object, ByRef pstrRealSource As String) As String

		On Error GoTo ErrorTrap

		Dim objDocumentStream As ADODB.Stream
		Dim objStream As ADODB.Stream
		Dim rsDocument As ADODB.Recordset
		Dim objDummyConnection As ADODB.Connection

		Dim strAccessibleFile As String
		Dim bIsPhoto As Boolean
		Dim strTempFile As String
		Dim strProperties As String = ""
		'Dim objTextStream As Scripting.TextStream

		Dim sSQL As String
		Dim strColumnName As String

		'Dim iOleType As Short
		'Dim strFileName As String
		'Dim strPath As String
		'Dim strUnc As String
		'Dim strDocumentSize As String
		'Dim strFileCreateDate As String
		'Dim strFileModifyDate As String

		' New record - thus no stream will exist
		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If plngRecordID = 0 Then
			GetPropertiesFromStream = ""
			Exit Function
		End If

		objStream = New ADODB.Stream
		objStream.Open()
		objStream.Type = ADODB.StreamTypeEnum.adTypeBinary

		strTempFile = GetTmpFName()
		'UPGRADE_WARNING: Couldn't resolve default property of object plngColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bIsPhoto = datGeneral.IsPhotoDataType(CInt(plngColumnID))

		' Open a temporary connection string to stream the data
		objDummyConnection = New ADODB.Connection
		objDummyConnection.Open(_mstrDummyConnectionString)

		'UPGRADE_WARNING: Couldn't resolve default property of object plngColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strColumnName = datGeneral.GetColumnName(CInt(plngColumnID))

		'UPGRADE_WARNING: Couldn't resolve default property of object plngRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID

		rsDocument = New ADODB.Recordset
		rsDocument.Open(sSQL, objDummyConnection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

		With rsDocument
			.MoveFirst()

			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDBNull(rsDocument.Fields(strColumnName).Value) Then
				objStream.Write(rsDocument.Fields(strColumnName).Value)
			Else
				GetPropertiesFromStream = ""
				GoTo TidyUpAndExit
			End If

			If objStream.Size > 0 Then
				objStream.SaveToFile(strTempFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				' objTextStream = mobjFileSystem.OpenTextFile(strTempFile, Scripting.IOMode.ForReading)
				Dim objTextStream = File.OpenRead(strTempFile)
				Dim b As Byte() = New Byte(400) {}
				' strProperties = Trim(objTextStream.Read(400))				
				Dim temp As New UTF8Encoding(True)
				objTextStream.Read(b, 0, b.Length)
				strProperties &= temp.GetString(b)

				_miOLEType = Val(Mid(strProperties, 9, 2))
				_mstrDisplayFileName = Trim(Path.GetFileName(Mid(strProperties, 11, 70)))
				_mstrFileName = Trim(Path.GetFileName(Mid(strProperties, 11, 70)))
				_mstrPath = Trim(Mid(strProperties, 81, 210))
				_mstrUnc = Trim(Mid(strProperties, 291, 60))
				_mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
				_mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
				_mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))
				objTextStream.Close()

				' Generate the properties display tag
				If bIsPhoto Then
					'If it's an embedded photo we'll need to generate it
					If _miOLEType = 2 Then
						'strAccessibleFile = mstrTempLocationPhysical & GetFileNameOnly(strFileName)
						strAccessibleFile = _mstrTempLocationPhysical & VB6.Format(Today, "yymmdd") & Path.GetFileName(GetTmpFName)
						objDocumentStream = New ADODB.Stream
						objDocumentStream.Type = ADODB.StreamTypeEnum.adTypeBinary
						objDocumentStream.Open()
						objStream.Position = 400
						objStream.CopyTo(objDocumentStream, objStream.Size - 400)
						objDocumentStream.SaveToFile(strAccessibleFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
						GetPropertiesFromStream = strAccessibleFile & "::EMBEDDED_OLE_DOCUMENT::"

						ReDim Preserve _mastrOleFilesInThisSession(UBound(_mastrOleFilesInThisSession) + 1)
						_mastrOleFilesInThisSession(UBound(_mastrOleFilesInThisSession)) = strAccessibleFile

					Else
						GetPropertiesFromStream = _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::"
					End If
				Else
					If _miOLEType = 2 Then
						GetPropertiesFromStream = _mstrFileName & "::EMBEDDED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
					Else
						GetPropertiesFromStream = _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
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

		'		'UPGRADE_NOTE: Object objStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'		objStream = Nothing
		'		'UPGRADE_NOTE: Object objDocumentStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'		objDocumentStream = Nothing
		'		'UPGRADE_NOTE: Object objDummyConnection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'		objDummyConnection = Nothing
		'		'UPGRADE_NOTE: Object rsDocument may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'		rsDocument = Nothing

		Exit Function

ErrorTrap:
		GetPropertiesFromStream = ""
		ProgramError("GetPropertiesFromStream", Err, Erl())
		Resume TidyUpAndExit



		Return GetPropertiesFromStream
	End Function

	Public Property FileName() As String
		Get
			' If linked file return proper link
			If _miOLEType = 2 Then
				FileName = mstrTempLocationUNC & Path.GetFileName(_mstrFileName)
			Else
				FileName = _mstrFileName
			End If

		End Get
		Set(ByVal Value As String)
			_mstrFileName = Value
		End Set
	End Property

	Public Property DisplayFilename() As String
		Get
			DisplayFilename = _mstrDisplayFileName
		End Get
		Set(ByVal Value As String)
			_mstrDisplayFileName = Value
		End Set
	End Property


	' Returns the size of the document in a nice formatted method
	Public ReadOnly Property DocumentSize() As String
		Get
			Select Case Len(_mstrDocumentSize)
				Case Is < 5
					DocumentSize = _mstrDocumentSize & " bytes"

				Case Is < 7
					DocumentSize = Mid(_mstrDocumentSize, 1, Len(_mstrDocumentSize) - 3) & "KB"

				Case 7
					DocumentSize = Mid(_mstrDocumentSize, 1, 1) & "." & Mid(_mstrDocumentSize, 2, 2) & "MB"

				Case Is < 10
					DocumentSize = Mid(_mstrDocumentSize, 1, Len(_mstrDocumentSize) - 6) & "MB"

			End Select

		End Get
	End Property

	Public ReadOnly Property DocumentModifyDate() As String
		Get
			DocumentModifyDate = _mstrFileModifyDate
		End Get
	End Property

	Public Sub New()
		'Set datData = New clsGeneral
		_mobjStream = New ADODB.Stream
		' mobjFileSystem = New Scripting.FileSystemObject

		_miOLEType = 3
		_mstrFileName = ""
		_mstrPath = ""

		Environ("USERNAME")
		_mbUseEncryption = False

		ReDim _mastrOleFilesInThisSession(0)
	End Sub

	' Commit the file back to the database
	Public Function SaveStream(ByRef plngRecordID As Integer, ByRef plngColumnID As Integer, ByRef pstrRealSource As String, ByRef pbReadOLEDirect As Boolean, ByVal buffer As Byte()) As Boolean

		On Error GoTo ErrorTrap

		Dim bOK As Boolean
		Dim cmADO As ADODB.Command
		Dim pmADO As ADODB.Parameter

		Dim objFile As ADODB.Stream

		'Dim objEncryption As clsEncryption
		'Dim objPropertiesStream As Scripting.TextStream
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
		If _mstrFileName <> "" Then

			strOLEType = Trim(Str(_miOLEType))

			Dim sb As New StringBuilder
			sb.Append("<<V002>>")
			sb.Append(strOLEType & Space(2 - Len(strOLEType)))
			sb.Append(GetFileNameOnly(_mstrFileName) & Space(70 - Len(GetFileNameOnly(_mstrFileName))))
			sb.Append(GetPathOnly(_mstrFileName, True) & Space(210 - Len(GetPathOnly(_mstrFileName, True))))
			sb.Append(GetUNCFromPath(_mstrFileName) & Space(60 - Len(GetUNCFromPath(_mstrFileName))))
			sb.Append(_mstrFileSize & Space(10 - Len(_mstrFileSize)))
			sb.Append(Space(20))
			sb.Append(_mstrFileModifyDate & Space(20 - Len(_mstrFileModifyDate)))

			Dim utf8 As Encoding = Encoding.UTF8
			Dim header As Byte() = utf8.GetBytes(sb.ToString())
			ReDim _mfileToEmbed((header.Length + buffer.Length) - 1)

			header.CopyTo(_mfileToEmbed, 0)

			' If embedded file tack onto the end of the stream
			If _miOLEType = 2 Then	' Embedded
				buffer.CopyTo(_mfileToEmbed, header.Length)
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

			If _mstrFileName <> "" Then
				If _mfileToEmbed.Length > 0 Then
					pmADO.Value = _mfileToEmbed
				Else
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
		lngReturn = 0	'  WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

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

		If _miOLEType = 3 Then
			UNCAndPath = GetUNCOnly(_mstrFileName) & GetPathOnly(_mstrPath, True)
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
			GetUNCFromPath = Left(pstrFilePath, InStr(pstrFilePath, "\", CompareMethod.Text) - 1)
		End If

		Exit Function

getUNC_error:
		ProgramError("GetUNCFromPath", Err, Erl())

	End Function


End Class

