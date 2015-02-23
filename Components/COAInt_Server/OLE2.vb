Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Text
Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient

Public Class Ole
	Inherits BaseForDMI

	Private _mstrTempLocationPhysical As String
	' Holds the names of the OLE files for this record session
	Private _mastrOleFilesInThisSession() As String
	Private mstrTempLocationUNC As String

	Private _msOLEVersionType As String
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

	Private _misPhoto As Boolean

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


	Public Property IsPhoto() As Boolean
		Get
			IsPhoto = _misPhoto
		End Get
		Set(value As Boolean)
			_misPhoto = value
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

	Public Sub CleanupOleFiles()
	End Sub

	Public Function CreateOLEDocument(ByRef plngRecordID As Long, ByRef plngColumnID As Long, ByRef pstrRealSource As String) As Byte()

		Dim sSQL As String
		Dim rsDocument As DataRow

		Dim strTempFile As String
		Dim strProperties As String = ""
		Dim strColumnName As String
		Dim objTextStream As FileStream

		Dim abtImage As Byte()
		Dim responseFile As Byte()
		Dim fFileOK As Boolean
		Dim sErrorMessage As String

		Try

			' New record - thus no stream will exist
			If plngRecordID = 0 Then
				Return Nothing
			End If

			strColumnName = GetColumnName(plngColumnID)

			sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID
			rsDocument = DB.GetDataTable(sSQL).Rows(0)

			If Not IsDBNull(rsDocument(strColumnName)) Then
				abtImage = CType(rsDocument(strColumnName), Byte())

				_msOLEVersionType = Encoding.UTF8.GetString(rsDocument(strColumnName), 0, 8)

				If Not _msOLEVersionType = "<<V002>>" Then
					sErrorMessage = String.Format("Incorrect header version for column {0} in GetPropertiesFromStream ", strColumnName)
					ProgramError(sErrorMessage, Err, Erl())
					Return Nothing
				End If

				fFileOK = (abtImage.GetLength(0) > 0)
				strTempFile = Path.GetTempFileName()

				Dim b As Byte() = rsDocument(strColumnName)

				Dim fs = New FileStream(strTempFile, FileMode.Create)
				fs.Write(b, 0, b.Length)
				fs.Close()

				b = New Byte(399) {}
				objTextStream = File.OpenRead(strTempFile)
				Dim temp As New UTF8Encoding(True)
				objTextStream.Read(b, 0, b.Length)
				strProperties &= temp.GetString(b)

				responseFile = New Byte((objTextStream.Length - 1) - 400) {}
				objTextStream.Read(responseFile, 0, responseFile.Length)

				_miOLEType = Val(Mid(strProperties, 9, 2))
				_mstrDisplayFileName = Trim(Path.GetFileName(Mid(strProperties, 11, 70)))
				_mstrFileName = IIf(_miOLEType = 2, Path.GetTempFileName(), _mstrDisplayFileName)
				_mstrPath = Trim(Mid(strProperties, 81, 210))
				_mstrUnc = Trim(Mid(strProperties, 291, 60))
				_mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
				_mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
				_mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))

				objTextStream.Close()

				' Generate the file if it's not linked
				If _miOLEType = 2 Then
					' TODO: content stream to client - no holding area. - No need for this??? hyperlinks take care of it???
					' mstrFileName = GenerateDocumentFromStream
				Else
					If _mstrPath.Length > 0 AndAlso _mstrPath.Substring(0, 2) = "\\" Then
						_mstrFileName = _mstrPath & "\" & _mstrFileName
					Else
						_mstrFileName = _mstrUnc & _mstrPath & "\" & _mstrFileName
					End If

				End If

			End If

		Catch ex As Exception
			_mstrFileName = ""
			_mstrDisplayFileName = ""
			ProgramError("GetPropertiesFromStream", Err, Erl())

		Finally

		End Try

		Return responseFile

	End Function

	Public Function GetPropertiesFromStream(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String) As String

		Dim rsDocument As DataRow
		Dim strTempFile As String
		Dim sSQL As String
		Dim strColumnName As String
		Dim sErrorMessage As String

		If plngRecordID = 0 Then
			Return ""
		End If

		strTempFile = Path.GetTempFileName()
		_misPhoto = IsPhotoDataType(plngColumnID)

		strColumnName = GetColumnName(plngColumnID)
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID

		rsDocument = DB.GetDataTable(sSQL).Rows(0)

		Try

			If Not IsDBNull(rsDocument(strColumnName)) Then

				_msOLEVersionType = Encoding.UTF8.GetString(rsDocument(strColumnName), 0, 8)

				If Not _msOLEVersionType = "<<V002>>" Then
					sErrorMessage = String.Format("Incorrect header version for column {0} in GetPropertiesFromStream ", strColumnName)
					ProgramError(sErrorMessage, Err, Erl())
					Return ""
				Else
					_miOLEType = Val(Encoding.UTF8.GetString(rsDocument(strColumnName), 8, 2))
					_mstrDisplayFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(rsDocument(strColumnName), 10, 70)))
					_mstrFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(rsDocument(strColumnName), 10, 70)))
					_mstrPath = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 80, 210))
					_mstrUnc = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 290, 60))
					_mstrDocumentSize = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 350, 10))
					_mstrFileCreateDate = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 360, 20))
					_mstrFileModifyDate = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 380, 20))
				End If

			Else
				Return ""
			End If

			If _miOLEType = 2 Then
				GetPropertiesFromStream = _mstrFileName & "::EMBEDDED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate & vbTab & _misPhoto.ToString()
			Else
				If _mstrPath.Length > 0 AndAlso _mstrPath.Substring(0, 2) = "\\" Then
					GetPropertiesFromStream = _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate & vbTab & _misPhoto.ToString()
				Else
					GetPropertiesFromStream = _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate & vbTab & _misPhoto.ToString()
				End If

			End If

		Catch ex As Exception
			GetPropertiesFromStream = ""
			ProgramError("GetPropertiesFromStream", Err, Erl())

		Finally

		End Try

		Return GetPropertiesFromStream

	End Function


	Public Function ExtractPhotoToBase64(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String) As String

		Dim rsDocument As DataRow

		Dim sSQL As String
		Dim strColumnName As String

		If plngRecordID = 0 Then
			ExtractPhotoToBase64 = ""
			Exit Function
		End If

		Dim bIsPhoto = IsPhotoDataType(plngColumnID)

		strColumnName = GetColumnName(plngColumnID)
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID

		rsDocument = DB.GetDataTable(sSQL).Rows(0)

		Try

			If Not IsDBNull(rsDocument(strColumnName)) Then
				_miOLEType = Val(Encoding.UTF8.GetString(rsDocument(strColumnName), 8, 2))
				_mstrDisplayFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(rsDocument(strColumnName), 10, 70)))
				_mstrFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(rsDocument(strColumnName), 10, 70)))
				_mstrPath = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 80, 210))
				_mstrUnc = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 290, 60))
				_mstrDocumentSize = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 350, 10))
				_mstrFileCreateDate = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 360, 20))
				_mstrFileModifyDate = Trim(Encoding.UTF8.GetString(rsDocument(strColumnName), 380, 20))

			Else
				Return ""
			End If

			If _miOLEType = 2 Then
				'Dim base64String As String					
				Dim abtImage = CType(rsDocument(strColumnName), Byte())
				Dim binaryData As Byte() = New Byte(abtImage.Length - 400) {}
				Try

					Buffer.BlockCopy(abtImage, 400, binaryData, 0, abtImage.Length - 400)

					ExtractPhotoToBase64 = Convert.ToBase64String(binaryData, 0, binaryData.Length)


				Catch exp As ArgumentNullException
					Console.WriteLine("Binary data array is null.")

				End Try
			Else
				If _mstrPath.Length > 0 AndAlso _mstrPath.Substring(0, 2) = "\\" Then
					ExtractPhotoToBase64 = _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
				Else
					ExtractPhotoToBase64 = _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
				End If

			End If


		Catch ex As Exception
			ProgramError("ExtractPhotoToBase64", Err, Erl())
			Return ""

		Finally

		End Try

		Return ExtractPhotoToBase64

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
				Case Else
					DocumentSize = ""
			End Select
		End Get
	End Property

	Public ReadOnly Property DocumentModifyDate() As String
		Get
			DocumentModifyDate = _mstrFileModifyDate
		End Get
	End Property

	Public Sub New()

		_miOLEType = 3
		_mstrFileName = ""
		_mstrPath = ""

		Environ("USERNAME")
		_mbUseEncryption = False

		ReDim _mastrOleFilesInThisSession(0)
	End Sub

	' Commit the file back to the database
	Public Function SaveStream(ByRef plngRecordID As Integer, ByRef plngColumnID As Integer, ByRef pstrRealSource As String, ByRef pbReadOLEDirect As Boolean, ByVal buffer As Byte()) As String

		Dim strErrMessage As String = ""
		Dim strOLEType As String
		Dim bUpdateField As Boolean = False
		Dim mfileToEmbed As Byte()

		Try

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

				ReDim mfileToEmbed((header.Length) - 1)

				header.CopyTo(mfileToEmbed, 0)

				' If embedded file tack onto the end of the stream
				If _miOLEType = 2 Then	' Embedded
					ReDim Preserve mfileToEmbed((header.Length + buffer.Length) - 1)
					buffer.CopyTo(mfileToEmbed, header.Length)
				End If

				' Flag the update to occur
				bUpdateField = True

			End If

			Dim prmCurrentID As New SqlParameter("piID", SqlDbType.Int)
			prmCurrentID.Value = plngRecordID

			Dim prmBlob As New SqlParameter("pimgUploadFile", SqlDbType.VarBinary)

			If _mstrFileName <> "" Then
				If mfileToEmbed.Length > 0 Then
					prmBlob.Value = mfileToEmbed
				Else
					prmBlob.Value = DBNull.Value
				End If
			Else
				prmBlob.Value = DBNull.Value
			End If

			DB.ExecuteSP("spASRUpdateOLEField_" & plngColumnID, prmCurrentID, prmBlob)

		Catch ex As Exception
			' ProgramError("SaveStream", Err, Erl())
			strErrMessage = ex.Message
		End Try

		Return strErrMessage

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

