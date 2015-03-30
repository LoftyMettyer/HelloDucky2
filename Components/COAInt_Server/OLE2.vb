Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Text
Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Enums

Public Class Ole
	Inherits BaseForDMI

	Private _msOLEVersionType As String
	Private _miOLEType As OLEType
	Private _mstrDisplayFileName As String
	Private _mstrFileName As String
	Private _mstrPath As String
	Private _mstrUnc As String
	Private _mstrDocumentSize As String
	Private _mstrFileSize As String
	Private _mstrFileCreateDate As String
	Private _mstrFileModifyDate As String

	Private _misPhoto As Boolean

	Public WriteOnly Property OLEFileSize() As String
		Set(ByVal value As String)
			_mstrFileSize = value
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

	Public Property OLEType() As OLEType
		Get
			OLEType = _miOLEType
		End Get
		Set(ByVal value As OLEType)
			_miOLEType = value
		End Set
	End Property

	Public Function CreateOLEDocument(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String) As Byte()

		Dim sSQL As String
		Dim rsDocument As DataRow

		Dim strTempFile As String
		Dim strProperties As String = ""
		Dim strColumnName As String
		Dim objTextStream As FileStream

		Dim responseFile As Byte()

		Try

			' New record - thus no stream will exist
			If plngRecordID = 0 Then
				Return Nothing
			End If

			strColumnName = GetColumnName(plngColumnID)

			sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID
			rsDocument = DB.GetDataTable(sSQL).Rows(0)

			If Not IsDBNull(rsDocument(strColumnName)) Then
				_msOLEVersionType = Encoding.UTF8.GetString(CType(rsDocument(strColumnName), Byte()), 0, 8)

				If Not _msOLEVersionType = "<<V002>>" Then
					Throw New Exception(String.Format("Incorrect header version for column {0} in GetPropertiesFromStream ", strColumnName))
				End If

				strTempFile = Path.GetTempFileName()

				Dim b As Byte() = CType(rsDocument(strColumnName), Byte())

				Dim fs = New FileStream(strTempFile, FileMode.Create)
				fs.Write(b, 0, b.Length)
				fs.Close()

				b = New Byte(399) {}
				objTextStream = File.OpenRead(strTempFile)
				Dim temp As New UTF8Encoding(True)
				objTextStream.Read(b, 0, b.Length)
				strProperties &= temp.GetString(b)

				responseFile = New Byte(CInt((objTextStream.Length - 1) - 400)) {}
				objTextStream.Read(responseFile, 0, responseFile.Length)

				_miOLEType = CType(Mid(strProperties, 9, 2), OLEType)
				_mstrDisplayFileName = Trim(Path.GetFileName(Mid(strProperties, 11, 70)))
				_mstrFileName = IIf(_miOLEType = OLEType.Embedded, Path.GetTempFileName(), _mstrDisplayFileName).ToString()
				_mstrPath = Trim(Mid(strProperties, 81, 210))
				_mstrUnc = Trim(Mid(strProperties, 291, 60))
				_mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
				_mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
				_mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))

				objTextStream.Close()

				' Generate the file if it's not linked
				If Not _miOLEType = OLEType.Embedded Then
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
			Throw

		Finally

		End Try

		Return responseFile

	End Function

	Public Function GetPropertiesFromStream(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String) As String

		Dim rsDocument As DataRow
		Dim sSQL As String
		Dim strColumnName As String

		If plngRecordID = 0 Then
			Return ""
		End If

		_misPhoto = IsPhotoDataType(plngColumnID)

		strColumnName = GetColumnName(plngColumnID)
		sSQL = "SELECT " & strColumnName & " FROM " & pstrRealSource & " WHERE ID=" & plngRecordID

		rsDocument = DB.GetDataTable(sSQL).Rows(0)

		Try

			If Not IsDBNull(rsDocument(strColumnName)) Then

				_msOLEVersionType = Encoding.UTF8.GetString(rsDocument(strColumnName), 0, 8)

				If Not _msOLEVersionType = "<<V002>>" Then
					Throw New Exception(String.Format("Incorrect header version for column {0} in GetPropertiesFromStream ", strColumnName))
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
			Throw

		Finally

		End Try

		Return GetPropertiesFromStream

	End Function

	Public Function ExtractPhotoToBase64(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String) As String

		Dim rsDocument As DataRow

		Dim sSQL As String
		Dim strColumnName As String
		Dim sExtracted As String = ""

		If plngRecordID = 0 Then
			Return ""
		End If

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

					sExtracted = ExtractPhotoToBase64
				Catch exp As ArgumentNullException
					Console.WriteLine("Binary data array is null.")

				End Try
			Else
				If _mstrPath.Length > 0 AndAlso _mstrPath.Substring(0, 2) = "\\" Then
					sExtracted = _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
				Else
					sExtracted = _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & _mstrDocumentSize & vbTab & _mstrFileCreateDate & vbTab & _mstrFileModifyDate
				End If

			End If

		Catch ex As Exception
			Throw

		Finally

		End Try

		Return sExtracted

	End Function

	Public Property FileName() As String
		Get
			' If linked file return proper link
			If _miOLEType = OLEType.Embedded Then
				FileName = Path.GetFileName(_mstrFileName)
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

		_miOLEType = OLEType.Linked
		_mstrFileName = ""
		_mstrPath = ""

	End Sub

	' Commit the file back to the database
	Public Function SaveStream(plngRecordID As Integer, plngColumnID As Integer, buffer As Byte()) As String

		Dim strErrMessage As String = ""
		Dim strOLEType As String
		Dim mfileToEmbed As Byte()

		Try

			' Is there a file attached?
			If _mstrFileName <> "" Then

				strOLEType = Trim(Str(_miOLEType))

				Dim sb As New StringBuilder
				sb.Append("<<V002>>")
				sb.Append(strOLEType & Space(2 - Len(strOLEType)))
				sb.Append(Path.GetFileName(_mstrFileName).PadRight(70))
				If _miOLEType = OLEType.Embedded Then
					' Pad the folder and root variables for embedded files.
					sb.Append(Space(210))
					sb.Append(Space(60))
				Else
					sb.Append(_mstrFileName.GetDirectoryNameOnly().PadRight(210))
					sb.Append(Path.GetPathRoot(_mstrFileName).PadRight(60))
				End If

				sb.Append(_mstrFileSize & Space(10 - Len(_mstrFileSize)))
				sb.Append(Space(20))
				sb.Append(_mstrFileModifyDate & Space(20 - Len(_mstrFileModifyDate)))

				Dim utf8 As Encoding = Encoding.UTF8
				Dim header As Byte() = utf8.GetBytes(sb.ToString())

				ReDim mfileToEmbed((header.Length) - 1)

				header.CopyTo(mfileToEmbed, 0)

				' If embedded file tack onto the end of the stream
				If _miOLEType = OLEType.Embedded Then
					ReDim Preserve mfileToEmbed((header.Length + buffer.Length) - 1)
					buffer.CopyTo(mfileToEmbed, header.Length)
				End If

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
			strErrMessage = ex.Message.RemoveSensitive()
		End Try

		Return strErrMessage

	End Function

	Public Function UNCAndPath() As String

		If _miOLEType = OLEType.Linked And _mstrFileName.Length > 0 Then
			Return Path.GetDirectoryName(_mstrFileName)
		Else
			Return ""
		End If

	End Function


End Class

