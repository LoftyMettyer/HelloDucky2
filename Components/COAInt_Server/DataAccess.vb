Option Strict On
Option Explicit On

Imports ADODB

Friend Class clsDataAccess
	
	Public Function OpenRecordset(ByRef sSQL As String, ByRef CursorType As CursorTypeEnum, ByRef LockType As LockTypeEnum _
		, Optional ByVal iCursorLocation As CursorLocationEnum = CursorLocationEnum.adUseServer) As Recordset

		' Open a recordset from the given SQL query, with the given recordset properties.
		Dim rsTemp As Recordset
		Dim iOldCursorLocation As CursorLocationEnum

		Try

			iOldCursorLocation = gADOCon.CursorLocation

			rsTemp = New Recordset
			gADOCon.CursorLocation = iCursorLocation

			rsTemp.Open(sSQL, gADOCon, CursorType, LockType, CommandTypeEnum.adCmdText)

		Catch ex As Exception
			Return Nothing

		Finally
			If (iOldCursorLocation = CursorLocationEnum.adUseClient) Or (iOldCursorLocation = CursorLocationEnum.adUseServer) Then
				gADOCon.CursorLocation = iOldCursorLocation
			Else
				gADOCon.CursorLocation = CursorLocationEnum.adUseServer
			End If

		End Try

		Return rsTemp

	End Function

	Public Function OpenPersistentRecordset(ByRef sSQL As String, ByRef CursorType As CursorTypeEnum, ByRef LockType As LockTypeEnum) As Recordset
		' Open a recordset from the given SQL query, with the given recordset properties.
		Dim rsTemp As Recordset

		rsTemp = New Recordset

		rsTemp.let_ActiveConnection(gADOCon)
		rsTemp.Properties("Preserve On Commit").Value = True
		rsTemp.Properties("Preserve On Abort").Value = True
		rsTemp.Open(sSQL, , CursorType, LockType, CommandTypeEnum.adCmdText)

		OpenPersistentRecordset = rsTemp

	End Function


	Public Sub ExecuteSql(ByRef sSQL As String)
		' Execute the given SQL statement.
		gADOCon.Execute(sSQL, , CommandTypeEnum.adCmdText)

	End Sub

	Public Function ExecuteSqlReturnAffected(ByRef sSQL As String) As Object
		' Execute the given SQL statement, and return the number of rows affected.
		Dim lngAffected As Object

		gADOCon.Execute(sSQL, lngAffected, CommandTypeEnum.adCmdText)
		ExecuteSqlReturnAffected = lngAffected

	End Function
End Class