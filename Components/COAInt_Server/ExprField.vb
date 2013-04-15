Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class clsExprField
	
	' Component definition variables.
	Private mlngTableID As Integer
	Private mlngColumnID As Integer
	Private miFieldPassType As Short
	Private miSelectionType As modExpression.FieldSelectionTypes
	Private mlngSelectionLine As Integer
	Private mlngSelOrderID As Integer
	Private mlngSelFilterID As Integer
	
	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent
	
	Private mstrUDFRuntimeCode As String
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		ContainsExpression = False
		
		If mlngSelFilterID > 0 Then
			' Check if the calc component IS the one we're checking for.
			ContainsExpression = (plngExprID = mlngSelFilterID)
			
			If Not ContainsExpression Then
				' The calc component IS NOT the one we're checking for.
				' Check if it contains the one we're looking for.
				ContainsExpression = HasExpressionComponent(mlngSelFilterID, plngExprID)
			End If
		End If
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	
	
	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		Dim Printer As New Printer
		' Print the component definition to the printer object.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		' Position the printing.
		With Printer
			.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
			.CurrentY = .CurrentY + giPRINT_YSPACE
			Printer.Print(ComponentDescription)
		End With
		
TidyUpAndExit: 
		PrintComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function WriteComponent() As Object
		' Write the component definition to the component recordset.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		
		fOK = True
		
		sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type," & " fieldTableID, fieldColumnID, fieldPassBy, fieldSelectionRecord," & " fieldSelectionLine, fieldSelectionOrderID, fieldSelectionFilter, valueLogic)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD)) & "," & " " & Trim(Str(mlngTableID)) & "," & " " & Trim(Str(mlngColumnID)) & "," & " " & Trim(Str(miFieldPassType)) & "," & " " & Trim(Str(miSelectionType)) & "," & " " & Trim(Str(mlngSelectionLine)) & "," & " " & Trim(Str(mlngSelOrderID)) & "," & " " & Trim(Str(mlngSelFilterID)) & "," & " 0)"
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
TidyUpAndExit: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function CopyComponent() As Object
		' Copies the selected component.
		' When editting a component we actually copy the component first
		' and edit the copy. If the changes are confirmed then the copy
		' replaces the original. If the changes are cancelled then the
		' copy is discarded.
		Dim objFieldCopy As New clsExprField
		
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
	
	Public ReadOnly Property ComponentType() As Short
		Get
			' Return the component type.
			ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD
			
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
			On Error GoTo ErrorTrap
			
			Dim fOK As Boolean
			Dim fChildField As Boolean
			Dim sSQL As String
			Dim sTableName As String
			Dim sColumnName As String
      Dim sSelectionType As String = ""
			Dim rsInfo As ADODB.Recordset
			
			' Get the column and table name.
			sSQL = "SELECT ASRSysColumns.columnName, ASRSysTables.tableName" & " FROM ASRSysColumns" & " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & " WHERE ASRSysColumns.columnID = " & Trim(Str(mlngColumnID))
			rsInfo = datGeneral.GetRecords(sSQL)
			With rsInfo
				fOK = Not (.EOF And .BOF)
				
				If fOK Then
					sColumnName = .Fields("ColumnName").Value
					sTableName = .Fields("TableName").Value
				Else
					sColumnName = "<unknown>"
					sTableName = "<unknown>"
				End If
				
				.Close()
			End With
			'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsInfo = Nothing
			
			If fOK Then
				' Add the selection type description if required.
				If (miFieldPassType = modExpression.FieldPassTypes.giPASSBY_VALUE) Then
					' Only give the full description if the field is in a child table of the
					' expression's parent table.
					
					sSQL = "SELECT *" & " FROM ASRSysRelations" & " WHERE parentID = " & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & " AND childID = " & Trim(Str(mlngTableID))
					rsInfo = datGeneral.GetRecords(sSQL)
					With rsInfo
						fChildField = Not (.EOF And .BOF)
						
						.Close()
					End With
					'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsInfo = Nothing
					
					If fChildField Then
						Select Case miSelectionType
							Case modExpression.FieldSelectionTypes.giSELECT_FIRSTRECORD
								sSelectionType = "(first record"
							Case modExpression.FieldSelectionTypes.giSELECT_LASTRECORD
								sSelectionType = "(last record"
							Case modExpression.FieldSelectionTypes.giSELECT_SPECIFICRECORD
								sSelectionType = "(line " & Trim(Str(mlngSelectionLine))
							Case modExpression.FieldSelectionTypes.giSELECT_RECORDTOTAL
								sSelectionType = "(total"
							Case modExpression.FieldSelectionTypes.giSELECT_RECORDCOUNT
								sSelectionType = "(record count"
							Case Else
								sSelectionType = "(<unknown>"
						End Select
						
						If mlngSelOrderID > 0 Then
							' Get the order name.
							sSQL = "SELECT name" & " FROM ASRSysOrders" & " WHERE orderID = " & Trim(Str(mlngSelOrderID))
							rsInfo = datGeneral.GetRecords(sSQL)
							With rsInfo
								If Not (.BOF And .EOF) Then
									sSelectionType = sSelectionType & ", order by '" & .Fields("Name").Value & "'"
								End If
								
								.Close()
							End With
							'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsInfo = Nothing
						End If
						
						If mlngSelFilterID > 0 Then
							' Get the filter name.
							sSQL = "SELECT name" & " FROM ASRSysExpressions" & " WHERE exprID = " & Trim(Str(mlngSelFilterID))
							rsInfo = datGeneral.GetRecords(sSQL)
							With rsInfo
								If Not (.BOF And .EOF) Then
									sSelectionType = sSelectionType & ", filter by '" & .Fields("Name").Value & "'"
								End If
								
								.Close()
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
			
TidyUpAndExit: 
			' Return the component description (to be displayed in the expression treeview).
			ComponentDescription = sTableName & " : " & sColumnName & " " & sSelectionType
			Exit Property
			
ErrorTrap: 
			sTableName = "<unknown>"
			sColumnName = "<unknown>"
			sSelectionType = "<unknown>"
			fOK = False
			Resume TidyUpAndExit
			
		End Get
	End Property
	
	Public ReadOnly Property ReturnType() As Short
		Get
			' Return the calculation's return type.
			On Error GoTo ErrorTrap
			
			Dim fOK As Boolean
			Dim iType As modExpression.ExpressionValueTypes
			Dim sSQL As String
			Dim rsColumn As ADODB.Recordset
			
			fOK = True
			
			' If the component returns the record count then
			' the return type must be numeric; otherwise the
			' return type is determined by the field type.
			If miSelectionType = modExpression.FieldSelectionTypes.giSELECT_RECORDCOUNT Then
				iType = modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
			Else
				' Determine the field's type by creating an
				' instance of the column class, and instructing
				' it to read its own details (including type).
				sSQL = "SELECT dataType" & " FROM ASRSysColumns" & " WHERE columnID = " & Trim(Str(mlngColumnID))
				rsColumn = datGeneral.GetRecords(sSQL)
				With rsColumn
					
					fOK = Not (.EOF And .BOF)
					
					If fOK Then
						Select Case .Fields("DataType").Value
							Case Declarations.SQLDataType.sqlNumeric, Declarations.SQLDataType.sqlInteger
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							Case Declarations.SQLDataType.sqlDate
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_DATE
							Case Declarations.SQLDataType.sqlVarChar, Declarations.SQLDataType.sqlLongVarChar
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							Case Declarations.SQLDataType.sqlBoolean
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							Case Declarations.SQLDataType.sqlOle
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_OLE
							Case Declarations.SQLDataType.sqlVarBinary
								iType = modExpression.ExpressionValueTypes.giEXPRVALUE_PHOTO
							Case Else
								fOK = False
						End Select
						
						If fOK Then
							If miFieldPassType = modExpression.FieldPassTypes.giPASSBY_REFERENCE Then
								iType = iType + giEXPRVALUE_BYREF_OFFSET
							End If
						End If
					End If
					
					.Close()
				End With
				'UPGRADE_NOTE: Object rsColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsColumn = Nothing
			End If
			
TidyUpAndExit: 
			If fOK Then
				ReturnType = iType
			Else
				ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED
			End If
			Exit Property
			
ErrorTrap: 
			fOK = False
			Resume TidyUpAndExit
			
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
	
	
	Public Property SelectionType() As Short
		Get
			' Return the selection type.
			SelectionType = miSelectionType
			
		End Get
		Set(ByVal Value As Short)
			' Set the selection type.
			'  If mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER Then
			'    miSelectionType = giSELECT_FIRSTRECORD
			'  Else
			miSelectionType = Value
			'  End If
			
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
	
	
	Public Property FieldPassType() As Short
		Get
			' Return the field pass type property.
			FieldPassType = miFieldPassType
			
		End Get
		Set(ByVal Value As Short)
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
		miFieldPassType = modExpression.FieldPassTypes.giPASSBY_VALUE
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Function GenerateCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, ByRef pfUDFCode As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim fOK As Boolean
		Dim fFound As Boolean
		Dim fColumnOK As Boolean
		Dim fParentField As Boolean
		Dim fNewSourceTable As Boolean
		Dim iLoop As Short
		Dim iNextIndex As Short
		Dim sSQL As String
		Dim sCode As String
		Dim sOtherTableName As String
		Dim sOrderCode As String
    Dim sFilterCode As String = ""
    Dim sColumnCode As String = ""
		Dim rsInfo As ADODB.Recordset
		Dim asViews() As String
    Dim avOrderJoinTables(,) As Object
		Dim objFilterExpr As clsExprExpression
		Dim objOrderTableView As CTablePrivilege
		Dim objTableView As CTablePrivilege
		Dim objOrderColumns As CColumnPrivileges
		Dim objView As CTablePrivilege
		Dim objViewColumns As CColumnPrivileges
		Dim objBaseTable As CTablePrivilege
		Dim objBaseColumns As CColumnPrivileges
		Dim objBaseColumn As CColumnPrivilege
    Dim strUDFReturnType As String = ""
		
		sCode = ""
		fOK = True
		
		If (miFieldPassType = modExpression.FieldPassTypes.giPASSBY_REFERENCE) Then
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
			
			If mobjBaseComponent.ParentExpression.BaseTableID = mlngTableID Then
				' The field is in the expression's base table.
				If Not pfApplyPermissions Then
					sCode = objBaseTable.TableName & "." & objBaseColumn.ColumnName
				Else
					fColumnOK = objBaseColumn.AllowSelect
					
					If fColumnOK Then
						sCode = objBaseTable.RealSource & "." & objBaseColumn.ColumnName
					Else
						fOK = (objBaseTable.TableType = Declarations.TableTypes.tabTopLevel)
						
						If fOK Then
							fOK = False
							' The column cannot be read from the table directly. Try the views on the table.
							ReDim asViews(0)
							For	Each objView In gcoTablePrivileges.Collection
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
				sSQL = "SELECT *" & " FROM ASRSysRelations" & " WHERE parentID = " & Trim(Str(mlngTableID)) & " AND childID = " & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID))
				rsInfo = datGeneral.GetRecords(sSQL)
				With rsInfo
					fParentField = Not (.EOF And .BOF)
					.Close()
				End With
				'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsInfo = Nothing
				
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
							fOK = (objBaseTable.TableType = Declarations.TableTypes.tabTopLevel)
							
							If fOK Then
								' The column cannot be read from the table directly. Try the views on the table.
								ReDim asViews(0)
								For	Each objView In gcoTablePrivileges.Collection
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
						sSQL = "SELECT ASRSysColumns.columnName, ASRSysColumns.columnID, ASRSysColumns.tableID, ASRSysTables.tableName, ASRSysOrderItems.ascending" & " FROM ASRSysOrderItems" & " JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID" & " JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID" & " WHERE orderID = " & Trim(Str(mlngSelOrderID)) & " AND type = 'O'" & " AND ASRSysColumns.columnID = ASRSysOrderItems.columnID" & " AND ASRSysColumns.tableID = ASRSysTables.tableID" & " ORDER BY sequence"
						
						rsInfo = datGeneral.GetRecords(sSQL)
						With rsInfo
							
							Do While Not .EOF
								If Not pfApplyPermissions Then
									' Construct the order code. Remember that if we are selecting the last record,
									' we must reverse the ASC/DESC options.
									sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & .Fields("TableName").Value & "." & .Fields("ColumnName").Value & IIf(miSelectionType = modExpression.FieldSelectionTypes.giSELECT_LASTRECORD, IIf(.Fields("Ascending").Value, " DESC", ""), IIf(.Fields("Ascending").Value, "", " DESC"))
									
									If (.Fields("TableID").Value <> mlngTableID) And ((.Fields("TableID").Value <> mobjBaseComponent.ParentExpression.BaseTableID) Or pfUDFCode) Then
										
										' Check if the table has already been added to the array of tables used in the order.
										fFound = False
										For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
											If avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value And (avOrderJoinTables(2, iNextIndex) = .Fields("TableName").Value) Then
												
												fFound = True
												Exit For
											End If
										Next iNextIndex
										
										If Not fFound Then
											iNextIndex = UBound(avOrderJoinTables, 2) + 1
											ReDim Preserve avOrderJoinTables(2, iNextIndex)
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value
											'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avOrderJoinTables(2, iNextIndex) = .Fields("TableName").Value
										End If
									End If
								Else
									' Get the permission object for the table.
									objOrderTableView = gcoTablePrivileges.Item(.Fields("TableName"))
									objOrderColumns = GetColumnPrivileges((objOrderTableView.TableName))
									
									fColumnOK = objOrderColumns.Item(.Fields("ColumnName")).AllowSelect
									
									If fColumnOK Then
										' Column can be read directly from the table.
										sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & objOrderTableView.RealSource & "." & .Fields("ColumnName").Value & IIf(miSelectionType = modExpression.FieldSelectionTypes.giSELECT_LASTRECORD, IIf(.Fields("Ascending").Value, " DESC", ""), IIf(.Fields("Ascending").Value, "", " DESC"))
										
										If (.Fields("TableID").Value <> mlngTableID) And ((.Fields("TableID").Value <> mobjBaseComponent.ParentExpression.BaseTableID) Or pfUDFCode) Then
											
											' Check if the table has already been added to the array of tables used in the order.
											fFound = False
											For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
												If avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value And (avOrderJoinTables(2, iNextIndex) = .Fields("TableName").Value) Then
													
													fFound = True
													Exit For
												End If
											Next iNextIndex
											
											If Not fFound Then
												iNextIndex = UBound(avOrderJoinTables, 2) + 1
												ReDim Preserve avOrderJoinTables(2, iNextIndex)
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value
												'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avOrderJoinTables(2, iNextIndex) = .Fields("TableName").Value
											End If
										End If
									Else
										fOK = (objOrderTableView.TableType = Declarations.TableTypes.tabTopLevel)
										
										If fOK Then
											' The column cannot be read from the table directly. Try the views on the table.
											ReDim asViews(0)
											For	Each objView In gcoTablePrivileges.Collection
												If (objView.TableID = .Fields("TableID").Value) And (Not objView.IsTable) And (objView.AllowSelect) Then
													
													objViewColumns = GetColumnPrivileges((objView.ViewName))
													
													If objViewColumns.IsValid(.Fields("ColumnName")) Then
														If objViewColumns.Item(.Fields("ColumnName")).AllowSelect Then
															' Add the view info to an array to be put into the column list or order code below.
															iNextIndex = UBound(asViews) + 1
															ReDim Preserve asViews(iNextIndex)
															asViews(iNextIndex) = objView.ViewName
															
															' Add the view to the Join code.
															' Check if the view has already been added to the join code.
															fFound = False
															For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
																'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																If avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value And avOrderJoinTables(2, iNextIndex) = objView.ViewName Then
																	fFound = True
																	Exit For
																End If
															Next iNextIndex
															
															If Not fFound Then
																' The view has not yet been added to the join code, so add it to the array and the join code.
																iNextIndex = UBound(avOrderJoinTables, 2) + 1
																ReDim Preserve avOrderJoinTables(2, iNextIndex)
																'UPGRADE_WARNING: Couldn't resolve default property of object avOrderJoinTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																avOrderJoinTables(1, iNextIndex) = .Fields("TableID").Value
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
													
													sColumnCode = sColumnCode & "WHEN NOT " & asViews(iNextIndex) & "." & .Fields("ColumnName").Value & " IS NULL THEN " & asViews(iNextIndex) & "." & .Fields("ColumnName").Value & vbNewLine
												Next iNextIndex
												
												If Len(sColumnCode) > 0 Then
													sColumnCode = sColumnCode & "ELSE NULL" & vbNewLine & "END" & IIf(miSelectionType = modExpression.FieldSelectionTypes.giSELECT_LASTRECORD, IIf(.Fields("Ascending").Value, " DESC", ""), IIf(.Fields("Ascending").Value, "", " DESC"))
													
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
								
								.MoveNext()
							Loop 
							
							.Close()
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
								objFilterExpr = New clsExprExpression
								objFilterExpr.ExpressionID = mlngSelFilterID
								objFilterExpr.ConstructExpression()
								fOK = objFilterExpr.RuntimeFilterCode(sFilterCode, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
								'UPGRADE_NOTE: Object objFilterExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objFilterExpr = Nothing
							End If
						End If
					End If
					
					If fOK Then
						Select Case miSelectionType
							Case modExpression.FieldSelectionTypes.giSELECT_FIRSTRECORD, modExpression.FieldSelectionTypes.giSELECT_LASTRECORD
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
										If objTableView.TableType = Declarations.TableTypes.tabChild Then
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
								
							Case modExpression.FieldSelectionTypes.giSELECT_RECORDCOUNT
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
										If objTableView.TableType = Declarations.TableTypes.tabChild Then
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
								
							Case modExpression.FieldSelectionTypes.giSELECT_RECORDTOTAL
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
										If objTableView.TableType = Declarations.TableTypes.tabChild Then
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
								
							Case modExpression.FieldSelectionTypes.giSELECT_SPECIFICRECORD
								' Specific for runtime filters.
								
								Select Case ReturnType
									Case modExpression.ExpressionValueTypes.giEXPRVALUE_DATE
										strUDFReturnType = "datetime"
										
									Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
										strUDFReturnType = "varchar(MAX)"
										
									Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
										strUDFReturnType = "float"
										
									Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
										strUDFReturnType = "bit"
										
								End Select
								
								' Create the udf code for this field
								mstrUDFRuntimeCode = "CREATE FUNCTION [" & gsUsername & "].udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(@PersonnelID float)" & vbNewLine & "RETURNS " & strUDFReturnType & vbNewLine & "AS" & vbNewLine & "BEGIN" & vbNewLine & "   DECLARE @Result " & strUDFReturnType & vbNewLine & "   DECLARE GetRecord CURSOR SCROLL FOR "
								
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
										If objTableView.TableType = Declarations.TableTypes.tabChild Then
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
								
								sCode = sCode & " [" & gsUsername & "].udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(" & mobjBaseComponent.ParentExpression.BaseTableName & ".id)"
								
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
			If ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_DATE Then
				sCode = "convert(" & vbNewLine & "datetime, " & vbNewLine & "convert(" & vbNewLine & "varchar(20), " & vbNewLine & sCode & "," & vbNewLine & "101)" & vbNewLine & ")"
			End If
			
			If ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC Then
				sCode = "convert(" & vbNewLine & "float, " & vbNewLine & sCode & vbNewLine & ")"
			End If
			
			' JDM - 19/12/01 - Fault 3299 - Problems concatenating strings
			If ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER Then
				sCode = "IsNull((" & sCode & "),'')"
			End If
			
		End If
		
TidyUpAndExit: 
		If fOK Then
			psRuntimeCode = IIf(pfUDFCode, mstrUDFRuntimeCode, sCode)
		Else
			psRuntimeCode = ""
		End If
		
		GenerateCode = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		RuntimeCode = GenerateCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, False, plngFixedExprID, psFixedSQLCode)
		
	End Function
	
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
    Dim strUDFCode As String = ""
		
		UDFCode = GenerateCode(strUDFCode, palngSourceTables, pfApplyPermissions, pfValidating, "", True, plngFixedExprID, psFixedSQLCode)
		
		If Len(strUDFCode) > 0 Then
			ReDim Preserve psRuntimeCode(UBound(psRuntimeCode) + 1)
			psRuntimeCode(UBound(psRuntimeCode)) = strUDFCode
		End If
		
	End Function
End Class