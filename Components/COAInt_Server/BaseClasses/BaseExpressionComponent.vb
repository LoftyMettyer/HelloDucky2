Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Namespace BaseClasses
	Public Class BaseExpressionComponent

		Protected ReadOnly Login As LoginInfo
		Protected General As New clsGeneral
		Protected DB As New clsDataAccess
		Protected AccessLog As AccessLog

		Public Sub New(ByVal Value As LoginInfo)
			Login = Value
			DB = New clsDataAccess(Login)
			General = New clsGeneral(Login)
			AccessLog = New AccessLog(Login)
		End Sub

		' keep a manual record of allocated IDs in case users in SYS MGR have created expressions but not yet saved changes
		Protected Function GetUniqueID(ByRef strSetting As String, ByRef strTable As String, ByRef strColumn As String) As Integer

			Dim prmSettingKey = New SqlParameter("settingkey", SqlDbType.VarChar, 50)
			prmSettingKey.Value = strSetting

			Dim prmSettingValue = New SqlParameter("settingvalue", SqlDbType.Int)
			prmSettingValue.Direction = ParameterDirection.Output

			DB.ExecuteSP("spASRIntGetUniqueExpressionID", prmSettingKey, prmSettingValue)

			Return CInt(prmSettingValue.Value)

		End Function


#Region "From modExpression"

		Protected Function ExprDeleted(ByVal lngExprID As Integer) As Boolean

			Dim rsExprTemp As DataTable
			Dim sSQL As String
			Dim bFound As Boolean

			sSQL = String.Format("SELECT ExprID FROM ASRSysExpressions WHERE ExprID = {0}", lngExprID)
			rsExprTemp = DB.GetDataTable(sSQL)

			bFound = rsExprTemp.Rows.Count > 0

			Return Not bFound

		End Function

		Public Function GetPickListField(ByVal lngPicklistID As Integer, ByVal sField As String) As Object

			Dim sSQL As String
			Dim rsExpr As DataTable

			On Error GoTo ErrorTrap

			sSQL = "SELECT * FROM ASRSysPickListName WHERE PickListID = " & lngPicklistID
			rsExpr = DB.GetDataTable(sSQL)

			With rsExpr
				If .Rows.Count > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object GetPickListField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetPickListField = .Rows(0)(sField)
				End If
			End With


TidyUpAndExit:
			'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsExpr = Nothing
			Exit Function

ErrorTrap:
			'NO MSGBOX ON THE SERVER ! - MsgBox "Error retrieving field value from database.", vbOKOnly + vbCritical, App.Title
			Resume TidyUpAndExit

		End Function

		Public Function HasExpressionComponent(ByVal plngExprIDBeingSearched As Integer, ByVal plngExprIDSearchedFor As Integer) As Boolean
			'JPD 20040507 Fault 8600
			On Error GoTo ErrorTrap

			Dim rsExprComp As DataTable
			Dim rsExpr As DataTable
			Dim sSQL As String
			Dim lngSubExprID As Integer

			HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)

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
								'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								rsExpr = Nothing
						End Select

						If HasExpressionComponent Then
							Exit For
						End If

					Next
				End With

			End If

TidyUpAndExit:
			'UPGRADE_NOTE: Object rsExprComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsExprComp = Nothing

			Exit Function

ErrorTrap:
			Resume TidyUpAndExit

		End Function

#End Region

	End Class
End Namespace