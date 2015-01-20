Option Explicit On
Option Strict Off

Imports System.Web.Mvc
Imports System.IO
Imports System.Web
Imports System.Data.SqlClient
Imports System.Collections.ObjectModel
Imports System.Web.Script.Serialization
Imports DMI.NET.Classes
Imports Newtonsoft.Json
Imports HR.Intranet.Server

Namespace Controllers
	Public Class GenericController
		Inherits Controller

		<HttpGet>
		Public Function GetLookupFindRecords2(piTableID As Integer, piOrderID As Integer, piLookupColumnID As Integer, psFilterValue As String, piCallingColumnID As Integer, piFirstRecPos As Integer) As String
			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class

			Dim rstLookup As DataTable
			Dim _prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = piFirstRecPos}
			Dim _prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmLookupColumnGridPosition = New SqlParameter("piLookupColumnGridNumber", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

			Dim sThousandColumns As String = ""
			Dim sErrorDescription As String = ""
			Try
				sThousandColumns = Get1000SeparatorFindColumns(piTableID, 0, piOrderID)
			Catch ex As Exception				
			End Try

			rstLookup = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords2" _
											, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = piTableID} _
											, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = 0} _
											, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = piOrderID} _
											, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = piLookupColumnID} _
											, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 10000} _
											, _prmIsFirstPage _
											, _prmIsLastPage _
											, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = ""} _
											, _prmColumnType _
											, _prmColumnSize _
											, _prmColumnDecimals _
											, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = "LOAD"} _
											, _prmTotalRecCount _
											, _prmFirstRecPos _
											, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = 0} _
											, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = psFilterValue} _
											, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = piCallingColumnID} _
											, _prmLookupColumnGridPosition _
											, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = "False"})


			If rstLookup Is Nothing Or rstLookup.Rows.Count = 0 Then
				Return "{""total"":1,""page"":1,""records"":0,""rows"":"""",""colmodel"":""""}"
			Else
			
				Dim colModel As List(Of Object) = jqGridColModel.CreateColModel(rstLookup, sThousandColumns)
				
				Return "{""total"":1,""page"":1,""records"":" & rstLookup.Rows.Count & ",""rows"":" & JsonConvert.SerializeObject(rstLookup) & ", ""colmodel"":" & JsonConvert.SerializeObject(colModel) & "}"
			End If

		End Function

		<HttpGet>
		Public Function GetLookupFindRecords(piLookupColumnID As Integer, psFilterValue As String, piCallingColumnID As Integer, piFirstRecPos As Integer) As String
			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class

			Dim rstLookup As DataTable
			Dim _prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim _prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = piFirstRecPos}
			Dim _prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim _prmLookupColumnGridPosition = New SqlParameter("piLookupColumnGridNumber", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim sThousandColumns As String = ""
			Dim sErrorDescription As String = ""

			Dim prmThousandColumns = New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Try
				objDataAccess.ExecuteSP("spASRIntGetLookupFindColumnInfo", _
										New SqlParameter("@piLookupColumnID", SqlDbType.Int) With {.Value = piLookupColumnID}, _
										prmThousandColumns _
				)
				sThousandColumns = prmThousandColumns.Value.ToString()
			Catch ex As Exception			
			End Try

			rstLookup = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords" _
											, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = piLookupColumnID} _
											, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 10000} _
											, _prmIsFirstPage _
											, _prmIsLastPage _
											, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = ""} _
											, _prmColumnType _
											, _prmColumnSize _
											, _prmColumnDecimals _
											, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = "LOAD"} _
											, _prmTotalRecCount _
											, _prmFirstRecPos _
											, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = 0} _
											, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = psFilterValue} _
											, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = piCallingColumnID} _
											, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = "False"})

			If rstLookup Is Nothing Or rstLookup.Rows.Count = 0 Then
				Return "{""total"":1,""page"":1,""records"":0,""rows"":"""",""colmodel"":""""}"
			Else
				Dim colModel As List(Of Object) = jqGridColModel.CreateColModel(rstLookup, sThousandColumns)

				Return "{""total"":1,""page"":1,""records"":" & rstLookup.Rows.Count & ",""rows"":" &
					JsonConvert.SerializeObject(rstLookup) & ", ""colmodel"":" & JsonConvert.SerializeObject(colModel) & "}"
			End If

		End Function
	End Class
End Namespace