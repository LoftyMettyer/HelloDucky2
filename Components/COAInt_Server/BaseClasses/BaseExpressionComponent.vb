Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Namespace BaseClasses
	Public Class BaseExpressionComponent
		Inherits BaseForDMI

		Protected ReadOnly Login As LoginInfo
		Protected General As New clsGeneral
		Protected DB As New clsDataAccess
		Protected AccessLog As AccessLog

		Protected Functions As ICollection(Of Metadata.Function)
		Protected Operators As ICollection(Of Metadata.Operator)

		Public Sub New(Value As SessionInfo)
			SessionInfo = Value
			Login = Value.LoginInfo
			DB = New clsDataAccess(Value.LoginInfo)
			General = New clsGeneral(Value.LoginInfo)
			AccessLog = New AccessLog(Value.LoginInfo)
			Functions = Value.Functions
			Operators = Value.Operators
			RunDate = Now
		End Sub

		Protected RunDate As DateTime

		' keep a manual record of allocated IDs in case users in SYS MGR have created expressions but not yet saved changes
		Protected Function GetUniqueID(strSetting As String, strTable As String, strColumn As String) As Integer

			Dim prmSettingKey = New SqlParameter("settingkey", SqlDbType.VarChar, 50)
			prmSettingKey.Value = strSetting

			Dim prmSettingValue = New SqlParameter("settingvalue", SqlDbType.Int)
			prmSettingValue.Direction = ParameterDirection.Output

			DB.ExecuteSP("spASRIntGetUniqueExpressionID", prmSettingKey, prmSettingValue)

			Return CInt(prmSettingValue.Value)

		End Function

		Public Overridable Function PrintComponent(piLevel As Integer) As Boolean
			Return False
		End Function

#Region "From modExpression"

		Protected Function ExprDeleted(lngExprID As Integer) As Boolean

			Dim rsExprTemp As DataTable
			Dim sSQL As String
			Dim bFound As Boolean

			sSQL = String.Format("SELECT ExprID FROM ASRSysExpressions WHERE ExprID = {0}", lngExprID)
			rsExprTemp = DB.GetDataTable(sSQL)

			bFound = rsExprTemp.Rows.Count > 0

			Return Not bFound

		End Function

		Public Function HasExpressionComponent(plngExprIDBeingSearched As Integer, plngExprIDSearchedFor As Integer) As Boolean

			Dim rsExprComp As DataTable
			Dim rsExpr As DataTable
			Dim sSQL As String
			Dim lngSubExprID As Integer

			HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)

			Try


				If Not HasExpressionComponent Then
					sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
					rsExprComp = DB.GetDataTable(sSQL)

					With rsExprComp
						For Each objRow As DataRow In .Rows

							Select Case CType(objRow("Type"), ExpressionComponentTypes)
								Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									lngSubExprID = CInt(IIf(IsDBNull(objRow("CalculationID")), 0, objRow("CalculationID")))

									If lngSubExprID > 0 Then
										HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
									End If

								Case ExpressionComponentTypes.giCOMPONENT_FILTER
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									lngSubExprID = CInt(IIf(IsDBNull(objRow("FilterID")), 0, objRow("FilterID")))

									If lngSubExprID > 0 Then
										HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
									End If

								Case ExpressionComponentTypes.giCOMPONENT_FIELD
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									lngSubExprID = CInt(IIf(IsDBNull(objRow("FieldSelectionFilter")), 0, objRow("FieldSelectionFilter")))

									If lngSubExprID > 0 Then
										HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
									End If

								Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
									sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(objRow("ComponentID"))
									rsExpr = DB.GetDataTable(sSQL)
									For Each objFunctionRow As DataRow In rsExpr.Rows

										HasExpressionComponent = HasExpressionComponent(CInt(objFunctionRow("ExprID")), plngExprIDSearchedFor)

										If HasExpressionComponent Then
											Exit For
										End If

									Next

							End Select

							If HasExpressionComponent Then
								Exit For
							End If

						Next
					End With

				End If


			Catch ex As Exception


			End Try

		End Function

#End Region

	End Class
End Namespace