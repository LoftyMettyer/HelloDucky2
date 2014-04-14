Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Partial Public Class Functions

	Public Class UniqueCode
		Public Code As String
		Public Value As Integer
		Public LastRecordID As Integer
	End Class

	Public Shared Numbers As New ArrayList

	Private Shared IsInstanciated As Boolean = False
	
	Private Shared Sub LoadUniqueCodes()

		Using conn As New SqlConnection("context connection=true")
			conn.Open()

			' If no numbers in process get next number from database
			Dim cmd As New SqlCommand("SELECT [CodePrefix], [MaxCodeSuffix] FROM dbo.[tbsys_uniquecodes]", conn)

			Dim dr As SqlDataReader = cmd.ExecuteReader()
			Dim uniqueCode As Functions.UniqueCode

			Do While dr.Read()
				uniqueCode = New Functions.UniqueCode
				uniqueCode.Code = dr.Item(0).ToString.ToUpper
				uniqueCode.Value = CInt(dr.Item(1).ToString)
				uniqueCode.LastRecordID = 0
				Numbers.Add(uniqueCode)
			Loop
		End Using

	End Sub

	<Microsoft.SqlServer.Server.SqlProcedure(Name:="spstat_flushuniquecode")> _
	Public Shared Sub FlushUniqueCode()

		Dim cmd As SqlCommand
		Dim sUpdate As String
		Dim objReader As SqlDataReader
		Dim UniqueCode As Functions.UniqueCode

		Try
			If Not IsInstanciated Then
				LoadUniqueCodes()
				IsInstanciated = True
			Else
				Using conn As New SqlConnection("context connection=true")
					conn.Open()
					' Flush current values to database
					For Each UniqueCode In Numbers
						UniqueCode.LastRecordID = 0
						sUpdate = String.Format("IF EXISTS(SELECT * FROM dbo.[tbsys_uniquecodes] WHERE CodePrefix = '{1}') " & _
						  "UPDATE dbo.[tbsys_uniquecodes] SET MaxCodeSuffix = {0} WHERE CodePrefix = '{1}' AND MaxCodeSuffix <> {0} " & _
						  "ELSE INSERT dbo.[tbsys_uniquecodes] (codeprefix, maxcodesuffix) VALUES ('{1}', {0})" _
						  , CInt(UniqueCode.Value), UniqueCode.Code)
						cmd = New SqlCommand(sUpdate, conn)
						cmd.ExecuteNonQuery()
					Next
				End Using
			End If

		Catch ex As Exception
			'   strMessage = ex.Message

		End Try

	End Sub

	<Microsoft.SqlServer.Server.SqlFunction(Name:="udfstat_getuniquecode", DataAccess:=DataAccessKind.Read)> _
	Public Shared Function GetUniqueCode(ByVal Prefix As String, ByVal RootValue As Long, ByVal RecordID As Integer) As SqlTypes.SqlString

		Dim sUniqueCode As String = ""
		Dim UniqueCode As UniqueCode
		Dim bFound As Boolean = False

		Try
			If Not IsInstanciated Then
				LoadUniqueCodes()
				IsInstanciated = True
			End If

			For Each UniqueCode In Numbers
				If UniqueCode.Code = Prefix.ToUpper Then
					bFound = True
					If Not UniqueCode.LastRecordID = RecordID Then
						UniqueCode.Value = UniqueCode.Value + 1
					End If
					UniqueCode.LastRecordID = RecordID
					sUniqueCode = UniqueCode.Value.ToString
					Exit For
				End If
			Next

			If Not bFound Then
				UniqueCode = New Functions.UniqueCode
				UniqueCode.Code = Prefix.ToUpper
				UniqueCode.Value = CInt(RootValue)
				UniqueCode.LastRecordID = RecordID
				Numbers.Add(UniqueCode)
				sUniqueCode = UniqueCode.Value.ToString
			End If

		Catch ex As Exception
			sUniqueCode = ex.Message
		End Try

		Return sUniqueCode

	End Function

End Class
