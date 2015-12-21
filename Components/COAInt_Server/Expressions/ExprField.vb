Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Namespace Expressions
	Friend Class clsExprField
		Inherits BaseExpressionComponent

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
			miFieldPassType = FieldPassTypes.ByValue
		End Sub

		' Component definition variables.
		Private mlngTableID As Integer
		Private mlngColumnID As Integer
		Private miFieldPassType As FieldPassTypes
		Private miSelectionType As FieldSelectionTypes
		Private mlngSelectionLine As Integer
		Private mlngSelOrderID As Integer
		Private mlngSelFilterID As Integer

		' Class handling variables.
		Private mobjBaseComponent As clsExprComponent

		Private mstrUDFRuntimeCode As String = ""

		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			' Retrun TRUE if the current expression (or any of its sub expressions)
			' contains the given expression. This ensures no cyclic expressions get created.
			'JPD 20040507 Fault 8600

			Dim bContainsExpression = False

			Try

				If mlngSelFilterID > 0 Then
					' Check if the calc component IS the one we're checking for.
					bContainsExpression = (plngExprID = mlngSelFilterID)

					If Not bContainsExpression Then
						' The calc component IS NOT the one we're checking for.
						' Check if it contains the one we're looking for.
						bContainsExpression = HasExpressionComponent(mlngSelFilterID, plngExprID)
					End If
				End If

			Catch ex As Exception
				Return True

			End Try

			Return bContainsExpression

		End Function

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_FIELD}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int), _
						New SqlParameter("operatorID", SqlDbType.Int), _
						New SqlParameter("valueType", SqlDbType.TinyInt), _
						New SqlParameter("valueCharacter", SqlDbType.VarChar, 255), _
						New SqlParameter("valueNumeric", SqlDbType.Float), _
						New SqlParameter("valueLogic", SqlDbType.Bit), _
						New SqlParameter("valueDate", SqlDbType.DateTime), _
						New SqlParameter("LookupTableID", SqlDbType.Int), _
						New SqlParameter("LookupColumnID", SqlDbType.Int), _
						New SqlParameter("fieldTableID", SqlDbType.Int) With {.Value = mlngTableID}, _
						New SqlParameter("fieldColumnID", SqlDbType.Int) With {.Value = mlngColumnID}, _
						New SqlParameter("fieldPassBy", SqlDbType.TinyInt) With {.Value = miFieldPassType}, _
						New SqlParameter("fieldSelectionRecord", SqlDbType.TinyInt) With {.Value = miSelectionType}, _
						New SqlParameter("fieldSelectionLine", SqlDbType.Int) With {.Value = mlngSelectionLine}, _
						New SqlParameter("fieldSelectionOrderID", SqlDbType.Int) With {.Value = mlngSelOrderID}, _
						New SqlParameter("fieldSelectionFilter", SqlDbType.Int) With {.Value = mlngSelFilterID}, _
						New SqlParameter("promptDescription", SqlDbType.VarChar, 255), _
						New SqlParameter("promptSize", SqlDbType.SmallInt), _
						New SqlParameter("promptDecimals", SqlDbType.SmallInt), _
						New SqlParameter("promptMask", SqlDbType.VarChar, 255), _
						New SqlParameter("promptDateType", SqlDbType.Int))

				Return True

			Catch ex As Exception
				Return False

			End Try

		End Function

		Public Function CopyComponent() As Object
			' Copies the selected component.
			' When editting a component we actually copy the component first
			' and edit the copy. If the changes are confirmed then the copy
			' replaces the original. If the changes are cancelled then the
			' copy is discarded.
			Dim objFieldCopy As New clsExprField(SessionInfo)

			' Copy the component's basic properties.
			With objFieldCopy
				.ColumnID = mlngColumnID
				.FieldPassType = miFieldPassType
				.SelectionLine = mlngSelectionLine
				.SelectionOrderID = mlngSelOrderID
				.SelectionType = miSelectionType
				.SelectionFilterID = mlngSelFilterID
				.TableID = mlngTableID
			End With

			CopyComponent = objFieldCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objFieldCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objFieldCopy = Nothing

		End Function

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_FIELD
			End Get
		End Property

		Public Property SelectionFilterID() As Integer
			Get
				' Return the Selection Filter property value.
				SelectionFilterID = mlngSelFilterID

			End Get
			Set(ByVal Value As Integer)
				' Set the Selection Filter property value.
				mlngSelFilterID = Value

			End Set
		End Property

		Public ReadOnly Property ComponentDescription() As String
			Get
				' Return a description of the field component.

				Dim fOK As Boolean
				Dim fChildField As Boolean
				Dim sSQL As String
				Dim sTableName As String
				Dim sColumnName As String
				Dim sSelectionType As String = ""
				Dim rsInfo As DataTable

				Try

					' Get the column and table name.
					sSQL = "SELECT ASRSysColumns.columnName, ASRSysTables.tableName FROM ASRSysColumns INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID WHERE ASRSysColumns.columnID = " & Trim(Str(mlngColumnID))
					rsInfo = DB.GetDataTable(sSQL)
					With rsInfo
						fOK = (.Rows.Count > 0)

						If fOK Then
							sColumnName = .Rows(0)("ColumnName").ToString()
							sTableName = .Rows(0)("TableName").ToString()
						Else
							sColumnName = "<unknown>"
							sTableName = "<unknown>"
						End If

					End With
					'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsInfo = Nothing

					If fOK Then
						' Add the selection type description if required.
						If (miFieldPassType = FieldPassTypes.ByValue) Then
							' Only give the full description if the field is in a child table of the
							' expression's parent table.

							sSQL = "SELECT * FROM ASRSysRelations WHERE parentID = " & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & " AND childID = " & Trim(Str(mlngTableID))
							rsInfo = DB.GetDataTable(sSQL)

							fChildField = (rsInfo.Rows.Count > 0)

							'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsInfo = Nothing

							If fChildField Then
								Select Case miSelectionType
									Case FieldSelectionTypes.giSELECT_FIRSTRECORD
										sSelectionType = "(first record"
									Case FieldSelectionTypes.giSELECT_LASTRECORD
										sSelectionType = "(last record"
									Case FieldSelectionTypes.giSELECT_SPECIFICRECORD
										sSelectionType = "(line " & Trim(Str(mlngSelectionLine))
									Case FieldSelectionTypes.giSELECT_RECORDTOTAL
										sSelectionType = "(total"
									Case FieldSelectionTypes.giSELECT_RECORDCOUNT
										sSelectionType = "(record count"
									Case Else
										sSelectionType = "(<unknown>"
								End Select

								If mlngSelOrderID > 0 Then
									' Get the order name.
									sSQL = "SELECT name FROM ASRSysOrders WHERE orderID = " & Trim(Str(mlngSelOrderID))
									rsInfo = DB.GetDataTable(sSQL)

									With rsInfo
										If (.Rows.Count > 0) Then
											sSelectionType = sSelectionType & ", order by '" & .Rows(0)("Name").ToString() & "'"
										End If
									End With
									'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									rsInfo = Nothing
								End If

								If mlngSelFilterID > 0 Then
									' Get the filter name.
									sSQL = "SELECT name FROM ASRSysExpressions WHERE exprID = " & Trim(Str(mlngSelFilterID))
									rsInfo = DB.GetDataTable(sSQL)
									With rsInfo
										If (.Rows.Count > 0) Then
											sSelectionType = sSelectionType & ", filter by '" & .Rows(0)("Name").ToString() & "'"
										End If
									End With
									'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									rsInfo = Nothing
								End If

								sSelectionType = sSelectionType & ")"
							End If
						Else
							sSelectionType = " (by reference)"
						End If
					End If

				Catch ex As Exception
					sTableName = "<unknown>"
					sColumnName = "<unknown>"
					sSelectionType = "<unknown>"

				End Try

				Return sTableName & " : " & sColumnName & " " & sSelectionType

			End Get
		End Property

		Public ReadOnly Property ReturnType() As ExpressionValueTypes
			Get

				Dim fOK As Boolean
				Dim iType As ExpressionValueTypes

				fOK = True

				' If the component returns the record count then
				' the return type must be numeric; otherwise the
				' return type is determined by the field type.
				If miSelectionType = FieldSelectionTypes.giSELECT_RECORDCOUNT Then
					iType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				Else

					' Determine the field's type by creating an
					' instance of the column class, and instructing
					' it to read its own details (including type).
					Select Case Columns.GetById(mlngColumnID).DataType
						Case ColumnDataType.sqlNumeric, ColumnDataType.sqlInteger
							iType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
						Case ColumnDataType.sqlDate
							iType = ExpressionValueTypes.giEXPRVALUE_DATE
						Case ColumnDataType.sqlVarChar, ColumnDataType.sqlLongVarChar
							iType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
						Case ColumnDataType.sqlBoolean
							iType = ExpressionValueTypes.giEXPRVALUE_LOGIC
						Case ColumnDataType.sqlOle
							iType = ExpressionValueTypes.giEXPRVALUE_OLE
						Case ColumnDataType.sqlVarBinary
							iType = ExpressionValueTypes.giEXPRVALUE_PHOTO
						Case Else
							fOK = False
					End Select

					If fOK Then
						If miFieldPassType = FieldPassTypes.ByReference Then
							iType = iType + giEXPRVALUE_BYREF_OFFSET
						End If
					End If
				End If

				If fOK Then
					Return iType
				Else
					Return ExpressionValueTypes.giEXPRVALUE_UNDEFINED
				End If

			End Get
		End Property

		Public Property SelectionOrderID() As Integer
			Get
				' Return the Selection Order property value.
				SelectionOrderID = mlngSelOrderID

			End Get
			Set(ByVal Value As Integer)
				' Set the Selection Order property value.
				mlngSelOrderID = Value

			End Set
		End Property

		Public Property SelectionType() As FieldSelectionTypes
			Get
				' Return the selection type.
				SelectionType = miSelectionType

			End Get
			Set(ByVal Value As FieldSelectionTypes)
				miSelectionType = Value
			End Set
		End Property

		Public Property SelectionLine() As Integer
			Get
				' Return the record slection line property.
				SelectionLine = mlngSelectionLine

			End Get
			Set(ByVal Value As Integer)
				' Set the record slection line property.
				mlngSelectionLine = Value

			End Set
		End Property

		Public Property FieldPassType() As FieldPassTypes
			Get
				' Return the field pass type property.
				FieldPassType = miFieldPassType

			End Get
			Set(ByVal Value As FieldPassTypes)
				' Set the field pass type property.
				miFieldPassType = Value

			End Set
		End Property

		Public Property ColumnID() As Integer
			Get
				' Return the column id property.
				ColumnID = mlngColumnID

			End Get
			Set(ByVal Value As Integer)
				' Set the column id property.
				mlngColumnID = Value

			End Set
		End Property


		Public Property BaseComponent() As clsExprComponent
			Get
				' Return the component's base component object.
				BaseComponent = mobjBaseComponent

			End Get
			Set(ByVal Value As clsExprComponent)
				' Set the component's base component object property.
				mobjBaseComponent = Value

			End Set
		End Property


		Public Property TableID() As Integer
			Get
				' Return the table id property.
				TableID = mlngTableID

			End Get
			Set(ByVal Value As Integer)
				' Set the table id property.
				mlngTableID = Value

			End Set
		End Property

		'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Private Sub Class_Initialize_Renamed()
			' Initialise properties.

		End Sub


		Private Function GenerateCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, pfApplyPermissions As Boolean, pfValidating As Boolean _
																 , ByRef pavPromptedValues As Object, ByRef psUDFs() As String _
																 , Optional ByRef plngFixedExprID As Integer = 0 _
																 , Optional ByRef psFixedSQLCode As String = "") As Boolean

			Dim fOK As Boolean
			Dim fFound As Boolean
			Dim fColumnOK As Boolean
			Dim fParentField As Boolean
			Dim fNewSourceTable As Boolean
			Dim iLoop As Integer
			Dim iNextIndex As Integer
			Dim sSQL As String
			Dim sCode As String
			Dim sOtherTableName As String
			Dim sOrderCode As String
			Dim sFilterCode As String = ""
			Dim sColumnCode As String = ""
			Dim rsInfo As DataTable
			Dim asViews() As String
			Dim avOrderJoinTables(,) As Object
			Dim objFilterExpr As clsExprExpression
			Dim objOrderTableView As TablePrivilege
			Dim objTableView As TablePrivilege
			Dim objOrderColumns As CColumnPrivileges
			Dim objView As TablePrivilege
			Dim objViewColumns As CColumnPrivileges
			Dim objBaseTable As TablePrivilege
			Dim objBaseColumns As CColumnPrivileges
			Dim objBaseColumn As ColumnPrivilege
			Dim strUDFReturnType As String = ""

			sCode = ""
			fOK = True

			If (miFieldPassType = FieldPassTypes.ByReference) Then
				sCode = Trim(Str(mlngColumnID))

				' JDM - 26/02/04 - Fault 8152 - Add this field to list of tables needed to be joined
				If mobjBaseComponent.ParentExpression.BaseTableID <> mlngTableID Then

					fNewSourceTable = True
					For iLoop = 1 To UBound(palngSourceTables, 2)
						'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (palngSourceTables(1, iLoop) = 0) And (palngSourceTables(2, iLoop) = mlngTableID) Then
							fNewSourceTable = False
							Exit For
						End If
					Next iLoop

					If fNewSourceTable Then
						iNextIndex = UBound(palngSourceTables, 2) + 1
						ReDim Preserve palngSourceTables(2, iNextIndex)
						'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						palngSourceTables(1, iNextIndex) = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						palngSourceTables(2, iNextIndex) = mlngTableID
					End If

				End If

			Else
				' Get the table and columns names.
				objBaseTable = gcoTablePrivileges.FindTableID(mlngTableID)
				objBaseColumns = GetColumnPrivileges((objBaseTable.TableName))
				objBaseColumn = objBaseColumns.FindColumnID(mlngColumnID)

				If mobjBaseComponent.ParentExpression.BaseTableID = mlngTableID Or mobjBaseComponent.ParentExpression.SecondTableID = mlngTableID Then
					' The field is in the expression's base table.
					If Not pfApplyPermissions Then
						sCode = objBaseTable.TableName & "." & objBaseColumn.ColumnName
					Else
						fColumnOK = objBaseColumn.AllowSelect

						If fColumnOK Then
							sCode = objBaseTable.RealSource & "." & objBaseColumn.ColumnName
						Else
							fOK = (objBaseTable.TableType = TableTypes.tabTopLevel)

							If fOK Then
								fOK = False
								' The column cannot be read from the table directly. Try the views on the table.
								ReDim asViews(0)
								For Each objView In gcoTablePrivileges.Collection
									If (objView.TableID = mlngTableID) And (Not objView.IsTable) And (objView.AllowSelect) Then

										objViewColumns = GetColumnPrivileges((objView.ViewName))

										If objViewColumns.IsValid((objBaseColumn.ColumnName)) Then
											If objViewColumns.Item((objBaseColumn.ColumnName)).AllowSelect Then
												' Add the view info to an array to be put into the column list or order code below.
												iNextIndex = UBound(asViews) + 1
												ReDim Preserve asViews(iNextIndex)
												asViews(iNextIndex) = objView.ViewName

												fNewSourceTable = True
												For iLoop = 1 To UBound(palngSourceTables, 2)
													'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													If (palngSourceTables(1, iLoop) = 1) And (palngSourceTables(2, iLoop) = objView.ViewID) Then
														fNewSourceTable = False
														Exit For
													End If
												Next iLoop

												If fNewSourceTable Then
													iNextIndex = UBound(palngSourceTables, 2) + 1
													ReDim Preserve palngSourceTables(2, iNextIndex)
													'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													palngSourceTables(1, iNextIndex) = 1
													'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													palngSourceTables(2, iNextIndex) = objView.ViewID
												End If
											End If
										End If

										'UPGRADE_NOTE: Object objViewColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
										objViewColumns = Nothing
									End If
								Next objView
								'UPGRADE_NOTE: Object objView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objView = Nothing

								fOK = (UBound(asViews) > 0)

								If fOK Then
									For iNextIndex = 1 To UBound(asViews)
										If iNextIndex = 1 Then
											sColumnCode = vbNewLine & "CASE " & vbNewLine
										End If

										sColumnCode = sColumnCode & "WHEN NOT " & asViews(iNextIndex) & "." & objBaseColumn.ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & objBaseColumn.ColumnName & vbNewLine
									Next iNextIndex

									If Len(sColumnCode) > 0 Then
										sColumnCode = sColumnCode & "ELSE NULL" & vbNewLine & "END"

										sCode = sColumnCode
									End If
								End If
							End If
						End If
					End If

				Else

					' Check if the table is a child or parent of the expression's base table.
					fParentField = Relations.Exists(Function(n) n.ParentID = mlngTableID And n.ChildID = mobjBaseComponent.ParentExpression.BaseTableID)

					If fParentField Then
						' The field is from a parent table of the expression's base table.
						If Not pfApplyPermissions Then
							sCode = objBaseTable.TableName & "." & objBaseColumn.ColumnName

							fNewSourceTable = True
							For iLoop = 1 To UBound(palngSourceTables, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (palngSourceTables(1, iLoop) = 0) And (palngSourceTables(2, iLoop) = objBaseTable.TableID) Then
									fNewSourceTable = False
									Exit For
								End If
							Next iLoop

							If fNewSourceTable Then
								iNextIndex = UBound(palngSourceTables, 2) + 1
								ReDim Preserve palngSourceTables(2, iNextIndex)
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								palngSourceTables(1, iNextIndex) = 0
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								palngSourceTables(2, iNextIndex) = objBaseTable.TableID
							End If
						Else
							fColumnOK = objBaseColumn.AllowSelect

							If fColumnOK Then
								sCode = objBaseTable.RealSource & "." & objBaseColumn.ColumnName

								fNewSourceTable = True
								For iLoop = 1 To UBound(palngSourceTables, 2)
									'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (palngSourceTables(1, iLoop) = 0) And (palngSourceTables(2, iLoop) = objBaseTable.TableID) Then
										fNewSourceTable = False
										Exit For
									End If
								Next iLoop

								If fNewSourceTable Then
									iNextIndex = UBound(palngSourceTables, 2) + 1
									ReDim Preserve palngSourceTables(2, iNextIndex)
									'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									palngSourceTables(1, iNextIndex) = 0
									'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									palngSourceTables(2, iNextIndex) = objBaseTable.TableID
								End If
							Else
								fOK = (objBaseTable.TableType = TableTypes.tabTopLevel)

								If fOK Then
									' The column cannot be read from the table directly. Try the views on the table.
									ReDim asViews(0)
									For Each objView In gcoTablePrivileges.Collection
										If (objView.TableID = mlngTableID) And (Not objView.IsTable) And (objView.AllowSelect) Then

											objViewColumns = GetColumnPrivileges((objView.ViewName))

											If objViewColumns.IsValid((objBaseColumn.ColumnName)) Then
												If objViewColumns.Item((objBaseColumn.ColumnName)).AllowSelect Then
													' Add the view info to an array to be put into the column list or order code below.
													iNextIndex = UBound(asViews) + 1
													ReDim Preserve asViews(iNextIndex)
													asViews(iNextIndex) = objView.ViewName

													fNewSourceTable = True
													For iLoop = 1 To UBound(palngSourceTables, 2)
														'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														If (palngSourceTables(1, iLoop) = 1) And (palngSourceTables(2, iLoop) = objView.ViewID) Then
															fNewSourceTable = False
															Exit For
														End If
													Next iLoop

													If fNewSourceTable Then
														iNextIndex = UBound(palngSourceTables, 2) + 1
														ReDim Preserve palngSourceTables(2, iNextIndex)
														'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														palngSourceTables(1, iNextIndex) = 1
														'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														palngSourceTables(2, iNextIndex) = objView.ViewID
													End If
												End If
											End If

											'UPGRADE_NOTE: Object objViewColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
											objViewColumns = Nothing
										End If
									Next objView
									'UPGRADE_NOTE: Object objView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									objView = Nothing

									fOK = (UBound(asViews) > 0)

									If fOK Then
										For iNextIndex = 1 To UBound(asViews)
											If iNextIndex = 1 Then
												sColumnCode = vbNewLine & "CASE " & vbNewLine
											End If

											sColumnCode = sColumnCode & "WHEN NOT " & asViews(iNextIndex) & "." & objBaseColumn.ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & objBaseColumn.ColumnName & vbNewLine
										Next iNextIndex

										If Len(sColumnCode) > 0 Then
											sColumnCode = sColumnCode & "ELSE NULL" & vbNewLine & "END"

											sCode = sColumnCode
										End If
									End If
								End If
							End If
						End If

					Else

						' The field is from a child table of the expression's base table.
						sCode = "(" & vbNewLine

						' Construct the order code if required.
						' Create an array of tables that need to be joined to make the order valid.
						' Column 1 = table ID.
						' Column 2 = table name.
						sOrderCode = ""
						ReDim avOrderJoinTables(2, 0)
						If mlngSelOrderID > 0 Then

							sSQL = String.Format("SELECT c.columnName, c.columnID, c.tableID, t.tableName, oi.ascending " & _
								"FROM ASRSysOrderItems oi " & _
								"JOIN ASRSysColumns c ON oi.columnID = c.columnID " & _
								"JOIN ASRSysTables t ON t.tableID = c.tableID " & _
								"WHERE orderID = {0} AND type = 'O' AND c.columnID = oi.columnID " & _
								"AND c.tableID = t.tableID " & _
								"AND c.dataType <> -4 AND c.datatype <> -3 " & _
								"ORDER BY oi.sequence", mlngSelOrderID)

							rsInfo = DB.GetDataTable(sSQL)
							With rsInfo
								For Each objRow As DataRow In .Rows

									If Not pfApplyPermissions Then
										' Construct the order code. Remember that if we are selecting the last record,
										' we must reverse the ASC/DESC options.
										sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "").ToString() & objRow("TableName").ToString() & "." & objRow("ColumnName").ToString() & IIf(miSelectionType = FieldSelectionTypes.giSELECT_LASTRECORD, IIf(objRow("Ascending"), " DESC", ""), IIf(objRow("Ascending"), "", " DESC"))

										If (objRow("TableID") <> mlngTableID) And ((objRow("TableID") <> mobjBaseComponent.ParentExpression.BaseTableID)) Then

											' Check if the table has already been added to the array of tables used in the order.
											fFound = False
											For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
												If avOrderJoinTables(1, iNextIndex) = objRow("TableID") And (avOrderJoinTables(2, iNextIndex) = objRow("TableName")) Then

													fFound = True
													Exit For
												End If
											Next iNextIndex

											If Not fFound Then
												iNextIndex = UBound(avOrderJoinTables, 2) + 1
												ReDim Preserve avOrderJoinTables(2, iNextIndex)
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avOrderJoinTables(1, iNextIndex) = objRow("TableID")
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avOrderJoinTables(2, iNextIndex) = objRow("TableName")
											End If
										End If
									Else
										' Get the permission object for the table.
										objOrderTableView = gcoTablePrivileges.Item(objRow("TableName"))
										objOrderColumns = GetColumnPrivileges((objOrderTableView.TableName))

										fColumnOK = objOrderColumns.Item(objRow("ColumnName")).AllowSelect

										If fColumnOK Then
											' Column can be read directly from the table.
											sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & objOrderTableView.RealSource & "." & objRow("ColumnName") & IIf(miSelectionType = FieldSelectionTypes.giSELECT_LASTRECORD, IIf(objRow("Ascending"), " DESC", ""), IIf(objRow("Ascending"), "", " DESC"))

											If (objRow("TableID") <> mlngTableID) And ((objRow("TableID") <> mobjBaseComponent.ParentExpression.BaseTableID)) Then

												' Check if the table has already been added to the array of tables used in the order.
												fFound = False
												For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
													If avOrderJoinTables(1, iNextIndex) = objRow("TableID") And (avOrderJoinTables(2, iNextIndex) = objRow("TableName")) Then

														fFound = True
														Exit For
													End If
												Next iNextIndex

												If Not fFound Then
													iNextIndex = UBound(avOrderJoinTables, 2) + 1
													ReDim Preserve avOrderJoinTables(2, iNextIndex)
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avOrderJoinTables(1, iNextIndex) = objRow("TableID")
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avOrderJoinTables(2, iNextIndex) = objRow("TableName")
												End If
											End If
										Else
											fOK = (objOrderTableView.TableType = TableTypes.tabTopLevel)

											If fOK Then
												' The column cannot be read from the table directly. Try the views on the table.
												ReDim asViews(0)
												For Each objView In gcoTablePrivileges.Collection
													If (objView.TableID = objRow("TableID")) And (Not objView.IsTable) And (objView.AllowSelect) Then

														objViewColumns = GetColumnPrivileges((objView.ViewName))

														If objViewColumns.IsValid(objRow("ColumnName")) Then
															If objViewColumns.Item(objRow("ColumnName")).AllowSelect Then
																' Add the view info to an array to be put into the column list or order code below.
																iNextIndex = UBound(asViews) + 1
																ReDim Preserve asViews(iNextIndex)
																asViews(iNextIndex) = objView.ViewName

																' Add the view to the Join code.
																' Check if the view has already been added to the join code.
																fFound = False
																For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
																	'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																	If avOrderJoinTables(1, iNextIndex) = objRow("TableID") And avOrderJoinTables(2, iNextIndex) = objView.ViewName Then
																		fFound = True
																		Exit For
																	End If
																Next iNextIndex

																If Not fFound Then
																	' The view has not yet been added to the join code, so add it to the array and the join code.
																	iNextIndex = UBound(avOrderJoinTables, 2) + 1
																	ReDim Preserve avOrderJoinTables(2, iNextIndex)
																	'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																	avOrderJoinTables(1, iNextIndex) = objRow("TableID")
																	'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																	avOrderJoinTables(2, iNextIndex) = objView.ViewName
																End If
															End If
														End If

														'UPGRADE_NOTE: Object objViewColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
														objViewColumns = Nothing
													End If
												Next objView
												'UPGRADE_NOTE: Object objView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
												objView = Nothing

												fOK = (UBound(asViews) > 0)

												If fOK Then
													For iNextIndex = 1 To UBound(asViews)
														If iNextIndex = 1 Then
															sColumnCode = vbNewLine & "CASE " & vbNewLine
														End If

														sColumnCode = sColumnCode & "WHEN NOT " & asViews(iNextIndex) & "." & objRow("ColumnName") & " IS NULL THEN " & asViews(iNextIndex) & "." & objRow("ColumnName") & vbNewLine
													Next iNextIndex

													If Len(sColumnCode) > 0 Then
														sColumnCode = sColumnCode & "ELSE NULL" & vbNewLine & "END" & IIf(miSelectionType = FieldSelectionTypes.giSELECT_LASTRECORD, IIf(objRow("Ascending"), " DESC", ""), IIf(objRow("Ascending"), "", " DESC"))

														sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & sColumnCode
													End If
												End If
											End If
										End If

										'UPGRADE_NOTE: Object objOrderTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
										objOrderTableView = Nothing
										'UPGRADE_NOTE: Object objOrderColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
										objOrderColumns = Nothing
									End If

								Next

							End With
							'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsInfo = Nothing
						End If

						If fOK Then
							' Create the filter code if required.
							sFilterCode = ""
							If mlngSelFilterID > 0 Then

								If mlngSelFilterID = plngFixedExprID Then
									sFilterCode = psFixedSQLCode
								Else
									objFilterExpr = New clsExprExpression(SessionInfo)
									objFilterExpr.ExpressionID = mlngSelFilterID
									objFilterExpr.ConstructExpression()
									fOK = objFilterExpr.RuntimeFilterCode(sFilterCode, pfApplyPermissions, psUDFs, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
									'UPGRADE_NOTE: Object objFilterExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									objFilterExpr = Nothing
								End If
							End If
						End If

						If fOK Then
							Select Case miSelectionType
								Case FieldSelectionTypes.giSELECT_FIRSTRECORD, FieldSelectionTypes.giSELECT_LASTRECORD
									' First and Last record selection uses the same code here.
									' The difference is made when creating the 'order by' code above.
									If Not pfApplyPermissions Then
										sCode = sCode & "SELECT TOP 1 " & objBaseTable.TableName & "." & objBaseColumn.ColumnName & vbNewLine & "FROM " & objBaseTable.TableName & vbNewLine

										' Add the JOIN code for the order.
										For iLoop = 1 To UBound(avOrderJoinTables, 2)
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											sCode = sCode & "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & objBaseTable.TableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
										Next iLoop

										sCode = sCode & "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & objBaseTable.TableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

										' Add the filter code as required.
										If Len(sFilterCode) > 0 Then
											sCode = sCode & "AND " & objBaseTable.TableName & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
										End If

										' Add the order code as required.
										If mlngSelOrderID > 0 Then
											sCode = sCode & "ORDER BY " & sOrderCode & vbNewLine
										End If
									Else
										fOK = objBaseColumn.AllowSelect

										If fOK Then
											sCode = sCode & "SELECT TOP 1 " & objBaseTable.RealSource & "." & objBaseColumn.ColumnName & vbNewLine & "FROM " & objBaseTable.RealSource & vbNewLine

											' Add the JOIN code for the order.
											For iLoop = 1 To UBound(avOrderJoinTables, 2)
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												objOrderTableView = gcoTablePrivileges.FindTableID(CInt(avOrderJoinTables(1, iLoop)))
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If objOrderTableView.TableName = avOrderJoinTables(2, iLoop) Then
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													sCode = sCode & "LEFT OUTER JOIN " & objOrderTableView.RealSource & " ON " & objBaseTable.RealSource & ".id_" & avOrderJoinTables(1, iLoop) & " = " & objOrderTableView.RealSource & ".id" & vbNewLine
												Else
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													sCode = sCode & "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & objBaseTable.RealSource & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
												End If
											Next iLoop

											sOtherTableName = mobjBaseComponent.ParentExpression.BaseTableName
											objTableView = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName)
											If objTableView.TableType = TableTypes.tabChild Then
												sOtherTableName = objTableView.RealSource
											End If
											'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
											objTableView = Nothing

											sCode = sCode & "WHERE " & sOtherTableName & ".id = " & objBaseTable.RealSource & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

											' Add the filter code as required.
											If Len(sFilterCode) > 0 Then
												sCode = sCode & "AND " & objBaseTable.RealSource & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
											End If

											' Add the order code as required.
											If mlngSelOrderID > 0 Then
												sCode = sCode & "ORDER BY " & sOrderCode & vbNewLine
											End If
										End If
									End If

								Case FieldSelectionTypes.giSELECT_RECORDCOUNT
									' No need to add the order code as it makes no differnt when selecting the record count.
									If Not pfApplyPermissions Then
										sCode = sCode & "SELECT COUNT(" & objBaseTable.TableName & ".id)" & vbNewLine & "FROM " & objBaseTable.TableName & vbNewLine & "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & objBaseTable.TableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

										' Add the filter code as required.
										If Len(sFilterCode) > 0 Then
											sCode = sCode & "AND " & objBaseTable.TableName & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
										End If
									Else
										fOK = objBaseColumn.AllowSelect

										If fOK Then
											sOtherTableName = mobjBaseComponent.ParentExpression.BaseTableName
											objTableView = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName)
											If objTableView.TableType = TableTypes.tabChild Then
												sOtherTableName = objTableView.RealSource
											End If
											'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
											objTableView = Nothing

											sCode = sCode & "SELECT COUNT(" & objBaseTable.RealSource & ".id)" & vbNewLine & "FROM " & objBaseTable.RealSource & vbNewLine & "WHERE " & sOtherTableName & ".id = " & objBaseTable.RealSource & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

											' Add the filter code as required.
											If Len(sFilterCode) > 0 Then
												sCode = sCode & "AND " & objBaseTable.RealSource & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
											End If
										End If
									End If

								Case FieldSelectionTypes.giSELECT_RECORDTOTAL
									' No need to add the order code as it makes no differnt when selecting the record total.
									If Not pfApplyPermissions Then
										sCode = sCode & "SELECT SUM(" & objBaseTable.TableName & "." & objBaseColumn.ColumnName & ")" & vbNewLine & "FROM " & objBaseTable.TableName & vbNewLine & "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & objBaseTable.TableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

										' Add the filter code as required.
										If Len(sFilterCode) > 0 Then
											sCode = sCode & "AND " & objBaseTable.TableName & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
										End If
									Else
										fOK = objBaseColumn.AllowSelect

										If fOK Then
											sOtherTableName = mobjBaseComponent.ParentExpression.BaseTableName
											objTableView = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName)
											If objTableView.TableType = TableTypes.tabChild Then
												sOtherTableName = objTableView.RealSource
											End If
											'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
											objTableView = Nothing

											sCode = sCode & "SELECT SUM(" & objBaseTable.RealSource & "." & objBaseColumn.ColumnName & ")" & vbNewLine & "FROM " & objBaseTable.RealSource & vbNewLine & "WHERE " & sOtherTableName & ".id = " & objBaseTable.RealSource & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

											' Add the filter code as required.
											If Len(sFilterCode) > 0 Then
												sCode = sCode & "AND " & objBaseTable.RealSource & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
											End If
										End If
									End If

								Case FieldSelectionTypes.giSELECT_SPECIFICRECORD
									' Specific for runtime filters.

									Select Case ReturnType
										Case ExpressionValueTypes.giEXPRVALUE_DATE
											strUDFReturnType = "datetime"

										Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
											strUDFReturnType = "varchar(MAX)"

										Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
											strUDFReturnType = "float"

										Case ExpressionValueTypes.giEXPRVALUE_LOGIC
											strUDFReturnType = "bit"

									End Select

									' Create the udf code for this field
									mstrUDFRuntimeCode = "CREATE FUNCTION [" & Login.Username & "].udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(@PersonnelID float)" & vbNewLine & "RETURNS " & strUDFReturnType & vbNewLine & "AS" & vbNewLine & "BEGIN" & vbNewLine & "   DECLARE @Result " & strUDFReturnType & vbNewLine & "   DECLARE GetRecord CURSOR SCROLL FOR "

									If Not pfApplyPermissions Then
										mstrUDFRuntimeCode = mstrUDFRuntimeCode & "SELECT " & objBaseTable.TableName & "." & objBaseColumn.ColumnName & vbNewLine & "FROM " & objBaseTable.TableName & vbNewLine

										' Add the JOIN code for the order.
										For iLoop = 1 To UBound(avOrderJoinTables, 2)
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											mstrUDFRuntimeCode = mstrUDFRuntimeCode & "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & objBaseTable.TableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
										Next iLoop

										mstrUDFRuntimeCode = mstrUDFRuntimeCode & "WHERE @PersonnelID = " & objBaseTable.TableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

										' Add the filter code as required.
										If Len(sFilterCode) > 0 Then
											mstrUDFRuntimeCode = mstrUDFRuntimeCode & "AND " & objBaseTable.TableName & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
										End If

										' Add the order code as required.
										If mlngSelOrderID > 0 Then
											mstrUDFRuntimeCode = mstrUDFRuntimeCode & "ORDER BY " & sOrderCode & vbNewLine
										End If
									Else
										fOK = objBaseColumn.AllowSelect

										If fOK Then
											mstrUDFRuntimeCode = mstrUDFRuntimeCode & "SELECT " & objBaseTable.RealSource & "." & objBaseColumn.ColumnName & vbNewLine & "FROM " & objBaseTable.RealSource & vbNewLine

											' Add the JOIN code for the order.
											For iLoop = 1 To UBound(avOrderJoinTables, 2)
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												objOrderTableView = gcoTablePrivileges.FindTableID(CInt(avOrderJoinTables(1, iLoop)))
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If objOrderTableView.TableName = avOrderJoinTables(2, iLoop) Then
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													mstrUDFRuntimeCode = mstrUDFRuntimeCode & "LEFT OUTER JOIN " & objOrderTableView.RealSource & " ON " & objBaseTable.RealSource & ".id_" & avOrderJoinTables(1, iLoop) & " = " & objOrderTableView.RealSource & ".id" & vbNewLine
												Else
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													mstrUDFRuntimeCode = mstrUDFRuntimeCode & "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & objBaseTable.RealSource & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
												End If
											Next iLoop

											sOtherTableName = mobjBaseComponent.ParentExpression.BaseTableName
											objTableView = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName)
											If objTableView.TableType = TableTypes.tabChild Then
												sOtherTableName = objTableView.RealSource
											End If
											'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
											objTableView = Nothing

											mstrUDFRuntimeCode = mstrUDFRuntimeCode & "WHERE @PersonnelID = " & objBaseTable.RealSource & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine

											' Add the filter code as required.
											If Len(sFilterCode) > 0 Then
												mstrUDFRuntimeCode = mstrUDFRuntimeCode & "AND " & objBaseTable.RealSource & ".id IN" & vbNewLine & "(" & vbNewLine & sFilterCode & ")" & vbNewLine
											End If

											' Add the order code as required.
											If mlngSelOrderID > 0 Then
												mstrUDFRuntimeCode = mstrUDFRuntimeCode & "ORDER BY " & sOrderCode & vbNewLine
											End If
										End If
									End If

									' Finish off udf code
									mstrUDFRuntimeCode = mstrUDFRuntimeCode & "OPEN GetRecord" & vbNewLine & "FETCH ABSOLUTE " & Trim(Str(mlngSelectionLine)) & " FROM GetRecord INTO @Result" & vbNewLine & "CLOSE GetRecord" & vbNewLine & "DEALLOCATE GetRecord" & vbNewLine & "RETURN @Result" & vbNewLine & "END"

									sCode = sCode & " [" & Login.Username & "].udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(" & mobjBaseComponent.ParentExpression.BaseTableName & ".id)"

									fOK = True

								Case Else
									' Unrecognised child record selection option.
									fOK = False
							End Select

							sCode = sCode & ")"

							' Add the table name to the list of source tables if it is not already there.
							fNewSourceTable = True
							For iLoop = 1 To UBound(palngSourceTables, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (palngSourceTables(1, iLoop) = 0) And (palngSourceTables(2, iLoop) = objBaseTable.TableID) Then
									fNewSourceTable = False
									Exit For
								End If
							Next iLoop

							If fNewSourceTable Then
								iNextIndex = UBound(palngSourceTables, 2) + 1
								ReDim Preserve palngSourceTables(2, iNextIndex)
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								palngSourceTables(1, iNextIndex) = 0
								'UPGRADE_WARNING: Couldn't resolve default property of object palngSourceTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								palngSourceTables(2, iNextIndex) = objBaseTable.TableID
							End If
						End If
					End If
				End If

				'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objBaseTable = Nothing
				'UPGRADE_NOTE: Object objBaseColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objBaseColumn = Nothing
				'UPGRADE_NOTE: Object objBaseColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objBaseColumns = Nothing

				' If the return type is a date, then convert the datetime value
				' to a varchar, and then back to a datetime. This gets rid of the time part
				' of the datetime value, which may cause errors when comparing datetime values.
				If ReturnType = ExpressionValueTypes.giEXPRVALUE_DATE Then
					sCode = "convert(" & vbNewLine & "datetime, " & vbNewLine & "convert(" & vbNewLine & "varchar(20), " & vbNewLine & sCode & "," & vbNewLine & "101)" & vbNewLine & ")"
				End If

				If ReturnType = ExpressionValueTypes.giEXPRVALUE_NUMERIC Then
					sCode = "convert(" & vbNewLine & "float, " & vbNewLine & sCode & vbNewLine & ")"
				End If

				' JDM - 19/12/01 - Fault 3299 - Problems concatenating strings
				If ReturnType = ExpressionValueTypes.giEXPRVALUE_CHARACTER Then
					sCode = "IsNull((" & sCode & "),'')"
				End If

			End If


			If mstrUDFRuntimeCode.Length > 0 Then
				ReDim Preserve psUDFs(psUDFs.Length)
				psUDFs(psUDFs.Length - 1) = mstrUDFRuntimeCode
			End If

			psRuntimeCode = sCode

			Return fOK

		End Function

		Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
																, pfValidating As Boolean, ByRef pavPromptedValues As Object _
																, ByRef psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			RuntimeCode = GenerateCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)

		End Function

	End Class
End Namespace