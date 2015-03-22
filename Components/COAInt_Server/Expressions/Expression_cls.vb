Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions

	Public Class Expression
		Inherits BaseExpressionComponent

		Private mobjBaseExpr As clsExprExpression

		Private mlngBaseTableID As Integer
		Private mlngExpressionID As Integer
		Private miType As Short
		Private miReturnType As ExpressionValueTypes
		Private mvarPrompts(,) As Object
		Private mastrUDFsRequired() As String

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
		End Sub


		Public Function Initialise(plngBaseTableID As Integer, plngExpressionID As Integer, piType As Short, piReturnType As ExpressionValueTypes) As Boolean

			mlngBaseTableID = plngBaseTableID
			mlngExpressionID = plngExpressionID
			miType = piType
			miReturnType = piReturnType
			Return True

		End Function

		Public Function TestFilterCode(ByRef psFilterCode As String) As Integer

			Dim rstTemp As DataTable

			Try
				rstTemp = DB.GetDataTable(psFilterCode)
				Return CInt(rstTemp.Rows(0)(0))

			Catch ex As Exception
				Return -1

			End Try

		End Function

		Public Function RuntimeFilterCode() As String
			Dim strSQL As String = ""
			Dim strFilterCode As String
			Dim fOK As Boolean

			If mastrUDFsRequired Is Nothing Then
				ReDim mastrUDFsRequired(0)
			End If

			fOK = mobjBaseExpr.RuntimeFilterCode(strFilterCode, True, mastrUDFsRequired, False, mvarPrompts)

			If fOK And gcoTablePrivileges.Item((mobjBaseExpr.BaseTableName)).AllowSelect Then
				strSQL = "SELECT COUNT(ID) FROM " & gcoTablePrivileges.Item((mobjBaseExpr.BaseTableName)).RealSource & " WHERE ID IN (" & strFilterCode & ")"
			End If

			Return strSQL

		End Function

		Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

			' Purpose : This function calls the individual functions that
			'           generate the components of the main SQL string.
			Dim iLoop As Short
			Dim iDataType As Short
			Dim lngComponentID As Integer

			ReDim mvarPrompts(1, 0)

			Try

				If IsArray(pavPromptedValues) Then
					ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

					For iLoop = 0 To UBound(pavPromptedValues, 2)

						' Get the prompt data type.
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 11))
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

							'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarPrompts(0, iLoop) = lngComponentID

							' NB. Locale to server conversions are done on the client.
							Select Case iDataType
								Case 2
									' Numeric.
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
								Case 3
									' Logic.
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
								Case 4
									' Date.
									' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
									' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
									' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
									' THINGS UP.
									'mvarPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
								Case Else
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
							End Select
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarPrompts(0, iLoop) = 0
						End If
					Next iLoop
				End If

			Catch ex As Exception
				Return False

			End Try

			Return True

		End Function

		Public Function SaveExpression(ByRef psName As String, ByRef psUserName As String, ByRef psAccess As String, ByRef psDescription As String) As Boolean

			' Save the expression.
			Dim lngOriginalExprID As Integer

			lngOriginalExprID = mobjBaseExpr.ExpressionID

			' Remove leading and training tabs from the description.
			Do While Left(psDescription, 1) = vbTab
				psDescription = Mid(psDescription, 2)
			Loop
			Do While Right(psDescription, 1) = vbTab
				psDescription = Left(psDescription, Len(psDescription) - 1)
			Loop

			mobjBaseExpr.Name = psName
			mobjBaseExpr.Owner = psUserName
			mobjBaseExpr.Access = psAccess
			mobjBaseExpr.Description = psDescription

			SaveExpression = mobjBaseExpr.WriteExpression

			If SaveExpression Then
				If lngOriginalExprID = 0 Then
					Select Case mobjBaseExpr.ExpressionType
						Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
							AccessLog.UtilCreated(UtilityType.utlCalculation, (mobjBaseExpr.ExpressionID))
						Case ExpressionTypes.giEXPR_RUNTIMEFILTER
							AccessLog.UtilCreated(UtilityType.utlFilter, (mobjBaseExpr.ExpressionID))
					End Select
				Else
					Select Case mobjBaseExpr.ExpressionType
						Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
							AccessLog.UtilUpdateLastSaved(UtilityType.utlCalculation, (mobjBaseExpr.ExpressionID))
						Case ExpressionTypes.giEXPR_RUNTIMEFILTER
							AccessLog.UtilUpdateLastSaved(UtilityType.utlFilter, (mobjBaseExpr.ExpressionID))
					End Select
				End If
			End If

		End Function

		Public Function SetExpressionDefinition(ByRef psComponentString1 As String, ByRef psComponentString2 As String, ByRef psComponentString3 As String, ByRef psComponentString4 As String, ByRef psComponentString5 As String, ByRef psNames As String) As Boolean

			' Construct the expression from the given definition strings.
			Dim iIndex As Integer
			Dim iIndex2 As Integer
			Dim iParameterIndex As Integer
			Dim sDefn As String
			Dim sCompType As String
			Dim sParameter As String

			Dim sNodeKey As String
			Dim lngCompID As Integer
			Dim iType As ExpressionComponentTypes
			Dim lngFieldColumnID As Integer
			Dim iFieldPassBy As Short
			Dim lngFieldSelectionTableID As Integer
			Dim iFieldSelectionRecord As Short
			Dim lngFieldSelectionLine As Integer
			Dim lngFieldSelectionOrderID As Integer
			Dim lngFieldSelectionFilter As Integer
			Dim lngFunctionID As Integer
			Dim lngCalculationID As Integer
			Dim lngOperatorID As Integer
			Dim iValueType As ExpressionValueTypes
			Dim sValueCharacter As String
			Dim dblValueNumeric As Double
			Dim fValueLogic As Boolean
			Dim sValueDate As String
			Dim dtValueDate As Date
			Dim sPromptDescription As String
			Dim sPromptMask As String
			Dim iPromptSize As Short
			Dim iPromptDecimals As Short
			Dim iFunctionReturnType As Short
			Dim lngLookupTableID As Integer
			Dim lngLookupColumnID As Integer
			Dim lngFilterID As Integer
			Dim iPromptDateType As Short
			Dim lngFieldTableID As Integer
			Dim sYear As String
			Dim sMonth As String
			Dim sDay As String
			Dim sName As String

			Dim objExpr As clsExprExpression
			Dim objComponent As clsExprComponent
			Dim avSubExpressions(,) As Object

			Dim objParameter As clsExprComponent
			Dim iNextIndex As Short
			Dim iCount As Integer
			Dim sParentNodeKey As String

			Try

				' Loop through each component in the definition.
				sDefn = psComponentString1 & psComponentString2 & psComponentString3 & psComponentString4 & psComponentString5
				sCompType = "U"

				If Not mobjBaseExpr Is Nothing Then
					'UPGRADE_NOTE: Object mobjBaseExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					mobjBaseExpr = Nothing
				End If
				mobjBaseExpr = New clsExprExpression(SessionInfo)
				mobjBaseExpr.Initialise(mlngBaseTableID, mlngExpressionID, miType, miReturnType)
				mobjBaseExpr.ExpandedNode = True

				' Create an array of sub-expressions.
				' Column 1 = sub-expression node key.
				' Column 2 = sub-expression's parent component node key.
				' Column 3 = sub-expression's expression object.
				ReDim avSubExpressions(3, 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(1, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				avSubExpressions(1, 1) = "ROOT"
				'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(2, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				avSubExpressions(2, 1) = ""
				avSubExpressions(3, 1) = mobjBaseExpr

				Do While Len(sDefn) > 0
					iIndex = InStr(sDefn, vbTab)

					If iIndex > 0 Then
						sParameter = Left(sDefn, iIndex - 1)
						sDefn = Mid(sDefn, iIndex + 1)
					Else
						sParameter = sDefn
						sDefn = ""
					End If

					If sCompType = "U" Then
						' Reading a new component.
						If sParameter = "ROOT" Then
							sCompType = "C"
						Else
							If Left(sParameter, 1) = "C" Then
								sCompType = "E"
							Else
								sCompType = "C"
							End If
						End If

						sParentNodeKey = sParameter
						iParameterIndex = 1
					Else
						If sCompType = "E" Then
							' Currently reading an expression.
							Select Case iParameterIndex
								Case 1
									sNodeKey = sParameter
								Case 15
									sCompType = "U"

									' Put this sub-expression in our array of sub-expressions.
									For iCount = 1 To UBound(avSubExpressions, 2)
										'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(2, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (avSubExpressions(2, iCount) = sParentNodeKey) And (avSubExpressions(1, iCount) = "") Then
											'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avSubExpressions(1, iCount) = sNodeKey

											iIndex2 = InStr(psNames, vbTab)
											If iIndex2 > 0 Then
												sName = Left(psNames, iIndex2 - 1)
												psNames = Mid(psNames, iIndex2 + 1)
											Else
												sName = psNames
												psNames = ""
											End If

											'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avSubExpressions(3, iCount).Name = sName
											Exit For
										End If
									Next iCount
							End Select
						Else
							' Currently reading a component.
							Select Case iParameterIndex
								Case 1
									sNodeKey = sParameter
									lngCompID = CInt(Mid(sNodeKey, 2))
								Case 4
									If Len(sParameter) > 0 Then
										iType = CShort(sParameter)
									Else
										iType = 0
									End If
								Case 5
									If Len(sParameter) > 0 Then
										lngFieldColumnID = CInt(sParameter)
									Else
										lngFieldColumnID = 0
									End If
								Case 6
									If Len(sParameter) > 0 Then
										iFieldPassBy = CShort(sParameter)
									Else
										iFieldPassBy = 0
									End If
								Case 7
									If Len(sParameter) > 0 Then
										lngFieldSelectionTableID = CInt(sParameter)
									Else
										lngFieldSelectionTableID = 0
									End If
								Case 8
									If Len(sParameter) > 0 Then
										iFieldSelectionRecord = CShort(sParameter)
									Else
										iFieldSelectionRecord = 0
									End If
								Case 9
									If Len(sParameter) > 0 Then
										lngFieldSelectionLine = CInt(sParameter)
									Else
										lngFieldSelectionLine = 0
									End If
								Case 10
									If Len(sParameter) > 0 Then
										lngFieldSelectionOrderID = CInt(sParameter)
									Else
										lngFieldSelectionOrderID = 0
									End If
								Case 11
									If Len(sParameter) > 0 Then
										lngFieldSelectionFilter = CInt(sParameter)
									Else
										lngFieldSelectionFilter = 0
									End If
								Case 12
									If Len(sParameter) > 0 Then
										lngFunctionID = CInt(sParameter)
									Else
										lngFunctionID = 0
									End If
								Case 13
									If Len(sParameter) > 0 Then
										lngCalculationID = CInt(sParameter)
									Else
										lngCalculationID = 0
									End If
								Case 14
									If Len(sParameter) > 0 Then
										lngOperatorID = CInt(sParameter)
									Else
										lngOperatorID = 0
									End If
								Case 15
									If Len(sParameter) > 0 Then
										iValueType = CShort(sParameter)
									Else
										iValueType = 0
									End If
								Case 16
									sValueCharacter = sParameter
								Case 17
									If Len(sParameter) > 0 Then
										dblValueNumeric = CDbl(sParameter)
									Else
										dblValueNumeric = 0
									End If
								Case 18
									fValueLogic = (sParameter = "1")
								Case 19
									' Date coming through in SQL Server mm/dd/yyyy format.
									' Need to convert it to a date value.
									sValueDate = sParameter
									If Len(sParameter) > 0 And sParameter <> "null" Then
										sMonth = Left(sValueDate, 2)
										sDay = Mid(sValueDate, 4, 2)
										sYear = Mid(sValueDate, 7)
										dtValueDate = DateSerial(CInt(sYear), CInt(sMonth), CInt(sDay))
									End If
								Case 20
									sPromptDescription = sParameter
								Case 21
									sPromptMask = sParameter
								Case 22
									If Len(sParameter) > 0 Then
										iPromptSize = CShort(sParameter)
									Else
										iPromptSize = 0
									End If
								Case 23
									If Len(sParameter) > 0 Then
										iPromptDecimals = CShort(sParameter)
									Else
										iPromptDecimals = 0
									End If
								Case 24
									If Len(sParameter) > 0 Then
										iFunctionReturnType = CShort(sParameter)
									Else
										iFunctionReturnType = 0
									End If
								Case 25
									If Len(sParameter) > 0 Then
										lngLookupTableID = CInt(sParameter)
									Else
										lngLookupTableID = 0
									End If
								Case 26
									If Len(sParameter) > 0 Then
										lngLookupColumnID = CInt(sParameter)
									Else
										lngLookupColumnID = 0
									End If
								Case 27
									If Len(sParameter) > 0 Then
										lngFilterID = CInt(sParameter)
									Else
										lngFilterID = 0
									End If
								Case 29
									If Len(sParameter) > 0 Then
										iPromptDateType = CShort(sParameter)
									Else
										iPromptDateType = 0
									End If
								Case 31
									If Len(sParameter) > 0 Then
										lngFieldTableID = CInt(sParameter)
									Else
										lngFieldTableID = 0
									End If
								Case 33
									sCompType = "U"

									' Find the component's parent expression.
									For iCount = 1 To UBound(avSubExpressions, 2)
										'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (avSubExpressions(1, iCount) = sParentNodeKey) Then
											objExpr = avSubExpressions(3, iCount)
											Exit For
										End If
									Next iCount

									objComponent = objExpr.AddComponent
									objComponent.ComponentType = iType

									With objComponent.Component
										Select Case iType
											Case ExpressionComponentTypes.giCOMPONENT_FIELD
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.TableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.TableID = lngFieldTableID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ColumnID = lngFieldColumnID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.FieldPassType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.FieldPassType = iFieldPassBy
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.SelectionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.SelectionType = iFieldSelectionRecord
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.SelectionLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.SelectionLine = lngFieldSelectionLine
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.SelectionOrderID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.SelectionOrderID = lngFieldSelectionOrderID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.SelectionFilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.SelectionFilterID = lngFieldSelectionFilter

											Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.FunctionID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.FunctionID = lngFunctionID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												For Each objParameter In .Parameters
													iNextIndex = UBound(avSubExpressions, 2) + 1
													ReDim Preserve avSubExpressions(3, iNextIndex)
													'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avSubExpressions(1, iNextIndex) = ""
													'UPGRADE_WARNING: Couldn't resolve default property of object avSubExpressions(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avSubExpressions(2, iNextIndex) = sNodeKey
													avSubExpressions(3, iNextIndex) = objParameter.Component
												Next objParameter

											Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.CalculationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.CalculationID = lngCalculationID

											Case ExpressionComponentTypes.giCOMPONENT_VALUE
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ReturnType = iValueType
												Select Case iValueType
													Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = sValueCharacter
													Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = dblValueNumeric
													Case ExpressionValueTypes.giEXPRVALUE_LOGIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = fValueLogic
													Case ExpressionValueTypes.giEXPRVALUE_DATE
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = dtValueDate
												End Select

											Case ExpressionComponentTypes.giCOMPONENT_OPERATOR
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.OperatorID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.OperatorID = lngOperatorID

											Case ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.TableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.TableID = lngLookupTableID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ColumnID = lngLookupColumnID
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ReturnType = iValueType

												Select Case iValueType
													Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = sValueCharacter
													Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = dblValueNumeric
													Case ExpressionValueTypes.giEXPRVALUE_LOGIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = fValueLogic
													Case ExpressionValueTypes.giEXPRVALUE_DATE
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.Value = dtValueDate
												End Select

											Case ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.Prompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.Prompt = sPromptDescription
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ValueType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ValueType = iValueType
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ReturnSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ReturnSize = iPromptSize
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ReturnDecimals. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ReturnDecimals = iPromptDecimals
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.ValueFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.ValueFormat = sPromptMask
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.DefaultDateType = iPromptDateType

												Select Case iValueType
													Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.DefaultValue = sValueCharacter
													Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.DefaultValue = dblValueNumeric
													Case ExpressionValueTypes.giEXPRVALUE_LOGIC
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.DefaultValue = fValueLogic
													Case ExpressionValueTypes.giEXPRVALUE_DATE
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.DefaultValue = dtValueDate
													Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
														'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														.DefaultValue = sValueCharacter
												End Select
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.LookupColumn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.LookupColumn = lngFieldColumnID

											Case ExpressionComponentTypes.giCOMPONENT_CUSTOMCALC
												' Not required.

											Case ExpressionComponentTypes.giCOMPONENT_EXPRESSION
												' Sub-expressions are handled via the Function component class.

											Case ExpressionComponentTypes.giCOMPONENT_FILTER
												' Load information for filters
												'UPGRADE_WARNING: Couldn't resolve default property of object objComponent.Component.FilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.FilterID = lngFilterID
										End Select
									End With

									objComponent.ComponentID = lngCompID
									iIndex2 = InStr(psNames, vbTab)
									If iIndex2 > 0 Then
										psNames = Mid(psNames, iIndex2 + 1)
									Else
										psNames = ""
									End If

									If iType = 7 Then

									End If
							End Select
						End If

						iParameterIndex = iParameterIndex + 1
					End If
				Loop


			Catch ex As Exception
				Return False

			End Try

			Return True

		End Function

		Public Function ReturnType() As ExpressionValueTypes
			Return mobjBaseExpr.ReturnType
		End Function

		Public Function ExpressionID() As Integer
			Return mobjBaseExpr.ExpressionID
		End Function

		Public Function ExistingExpressionReturnType(plngExprID As Integer) As ExpressionValueTypes
			Dim objExpression As clsExprExpression

			' Instantiate the calculation expression.
			objExpression = New clsExprExpression(SessionInfo)

			With objExpression
				' Construct the calculation expression.
				.ExpressionID = plngExprID
				.ConstructExpression()
				.ValidateExpression(False)
				Return .ReturnType
			End With

			'UPGRADE_NOTE: Object objExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objExpression = Nothing
		End Function

		Public Function ValidateExpression() As ExprValidationCodes

			ValidateExpression = mobjBaseExpr.ValidateExpression(True)

			If (ValidateExpression = ExprValidationCodes.giEXPRVALIDATION_NOERRORS) And (mlngExpressionID > 0) Then

				If mobjBaseExpr.ContainsExpression(mlngExpressionID) Then
					Return ExprValidationCodes.giEXPRVALIDATION_CYCLIC
				End If
			End If

		End Function

		Public Function ValidityMessage(piValidityCode As ExprValidationCodes) As String
			Return mobjBaseExpr.ValidityMessage(piValidityCode)
		End Function

		Public Function UDFFunctions(ByRef pbCreate As Boolean) As Boolean
			Return General.UDFFunctions(mastrUDFsRequired, pbCreate)
		End Function

	End Class
End Namespace