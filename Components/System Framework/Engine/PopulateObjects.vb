Option Strict Off

Imports SystemFramework.Enums
Imports SystemFramework.Enums.Errors

<HideModuleName()>
Public Module PopulateObjects

  Private _componentbase As DataSet
  Private _componentfunction As DataSet

  Public Sub PopulateThings()
    PopulateTables()
    PopulateTableRelations()
    PopulateTableColumns()
    PopulateTableOrders()
    PopulateTableOrderItems()
    Dim sw As New Stopwatch
    sw.Start()
    PopulateTableExpressions()

    Console.WriteLine("PTE:" & sw.ElapsedMilliseconds)

    PopulateTableViews()
    PopulateTableViewItems()
		PopulateTableValidations()
		PopulateTableTriggerCode()
    PopulateTableRecordDescriptions()
    PopulateTableMasks()
		PopulateFusionMessages()

    _componentbase = Nothing
    _componentfunction = Nothing
  End Sub

  Public Sub PopulateSystemSettings()

    SystemSettings.Clear()

    Dim ds As DataSet = CommitDb.ExecStoredProcedure("spadmin_getsystemsettings", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows
      Dim setting As New Setting
      setting.Module = row("section").ToString
      setting.Parameter = row.Item("settingkey").ToString
      setting.Value = row.Item("settingvalue").ToString

      SystemSettings.Add(setting)
    Next

  End Sub

  Public Sub PopulateModuleSettings()

    ModuleSetup.Clear()

    Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getmodulesetup", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim setting As New Setting
      setting.Module = row.Item("modulekey").ToString
      setting.Parameter = row.Item("parameterkey").ToString

      If Not row.Item("value").ToString = "" Then
        Select Case row.Item("subtype").ToString
          Case 1
            setting.Table = Tables.GetById(row.Item("value").ToString)
          Case Else
            setting.Value = row.Item("value").ToString
        End Select

        ModuleSetup.Add(setting)
      End If
    Next

  End Sub

  Public Sub PopulateSystemThings()

    Operators.Clear()
    Functions.Clear()

    Dim ds As DataSet = CommitDb.ExecStoredProcedure("spadmin_getcomponentcode", Nothing)
    For Each row As DataRow In ds.Tables(0).Rows

      Dim codeLibrary As New CodeLibrary
      codeLibrary.Id = row.Item("id").ToString
      codeLibrary.Name = row.Item("name").ToString
      codeLibrary.Code = row.Item("code").ToString
      codeLibrary.PreCode = row.Item("precode").ToString
      codeLibrary.AfterCode = row.Item("aftercode").ToString
      codeLibrary.ReturnType = row.Item("returntype").ToString
      codeLibrary.OperatorType = row.Item("operatortype").ToString
      codeLibrary.RecordIdRequired = row.Item("recordidrequired").ToString
      codeLibrary.CalculatePostAudit = row.Item("calculatepostaudit").ToString
      codeLibrary.IsUniqueCode = row.Item("isuniquecode").ToString
      codeLibrary.IsTimeDependant = row.Item("istimedependant").ToString
      codeLibrary.IsGetFieldFromDb = row.Item("isgetfieldfromdb").ToString
      codeLibrary.CaseCount = row.Item("casecount").ToString
      codeLibrary.MakeTypeSafe = row.Item("maketypesafe").ToString
      codeLibrary.OvernightOnly = row.Item("overnightonly").ToString
      codeLibrary.Tuning.Rating = row.Item("performancerating").ToString
      codeLibrary.DependsOnBankHoliday = row.Item("dependsonbankholiday").ToString
      codeLibrary.Dependancies = GetCodeLibraryDependancies(codeLibrary)

      If CBool(row.Item("isoperator").ToString) Then
        Operators.Add(codeLibrary)
      Else
        Functions.Add(codeLibrary)
      End If

    Next

  End Sub

  Private Function GetCodeLibraryDependancies(ByVal codeLibrary As CodeLibrary) As ICollection(Of Setting)

    Dim params As New Connectivity.Parameters
    params.Add("@componentid", codeLibrary.Id)
    Dim ds As DataSet = CommitDb.ExecStoredProcedure("spadmin_getcomponentcodedependancies", params)

    Dim dependancies As New Collection(Of Setting)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim setting As New Setting
      setting.SettingType = row.Item("type").ToString
      setting.Module = row.Item("parameterkey").ToString
      setting.Parameter = row.Item("modulekey").ToString
      setting.Value = row.Item("value").ToString
      setting.Code = row.Item("code").ToString

      dependancies.Add(setting)
    Next

    Return dependancies

  End Function

  Private Sub PopulateTables()

    Tables.Clear()

    Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_gettables", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim table As New Table
      table.Id = row.Item("id").ToString
      table.TableType = row.Item("tabletype").ToString
      table.Name = row.Item("name").ToString
      table.SchemaName = "dbo"
      table.IsRemoteView = row.Item("isremoteview").ToString
      table.AuditInsert = row.Item("auditinsert").ToString
      table.AuditDelete = row.Item("auditdelete").ToString
      table.DefaultEmailId = row.Item("defaultemailid").ToString
			table.DefaultOrderId = row.Item("defaultorderid").ToString
			table.InsertTriggerDisabled = row.Item("InsertTriggerDisabled").ToString
			table.UpdateTriggerDisabled = row.Item("UpdateTriggerDisabled").ToString
			table.DeleteTriggerDisabled = row.Item("DeleteTriggerDisabled").ToString

      table.State = row.Item("state")

      Tables.Add(table)
    Next

  End Sub

  Private Sub PopulateTableColumns()

    Try

      Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getcolumns", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Tables.GetById(row.Item("tableid").ToString)

        Dim column As New Column
        column.Table = table
        column.Id = row.Item("id").ToString
        column.Name = row.Item("name").ToString
        column.SchemaName = "dbo"
        column.Description = row.Item("description").ToString
        column.State = row.Item("state")
        column.DefaultCalcId = NullSafe(row, "defaultcalcid", 0).ToString
        column.DefaultValue = row.Item("defaultvalue").ToString
        column.CalcId = row.Item("calcid").ToString
        column.DataType = row.Item("datatype")
        column.Size = row.Item("size").ToString
        column.Decimals = row.Item("decimals").ToString
        column.Audit = row.Item("audit")
        column.Mandatory = row.Item("mandatory")
        column.Multiline = row.Item("multiline")
        column.IsReadOnly = row.Item("isreadonly")
        column.CaseType = row.Item("case").ToString
        column.CalculateIfEmpty = row.Item("calculateifempty")
        column.TrimType = NullSafe(row, "trimming", 0).ToString
        column.Alignment = row.Item("alignment").ToString

        table.Columns.Add(column)
      Next

    Catch ex As Exception
      ErrorLog.Add(Section.LoadingData, "PopulateTableColumns", Severity.Error, "Table or column not found", "Error in system metatdata - table or column not found")

    End Try

  End Sub

  Private Sub PopulateTableOrders()

    Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getorders", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim table As Table = Tables.GetById(row.Item("tableid").ToString)

      Dim tableOrder = New TableOrder
      tableOrder.Id = row.Item("orderid").ToString
      tableOrder.Name = row.Item("name").ToString
      tableOrder.Table = table

      table.TableOrders.Add(tableOrder)
    Next

  End Sub

  Private Sub PopulateTableOrderItems()

    Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getorderitems", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim orderId As Integer = CInt(row.Item("orderid"))
      Dim tableOrder As TableOrder = Tables.SelectMany(Function(t) t.TableOrders).Where(Function(o) o.Id = orderId).FirstOrDefault

      Dim orderItem As New TableOrderItem

      If Not tableOrder Is Nothing Then

        orderItem.Id = row.Item("orderid").ToString
        orderItem.ColumnType = row.Item("type")
        orderItem.Sequence = row.Item("sequence")
        orderItem.Ascending = row.Item("ascending")

        orderItem.TableOrder = tableOrder
        orderItem.Column = tableOrder.Table.Columns.GetById(CInt(row.Item("columnid")))

        tableOrder.Items.Add(orderItem)
      Else
        ErrorLog.Add(Section.LoadingData, "OrderItems", Severity.Warning, "Order not found", "order " & CStr(orderId) & " not found")

      End If
    Next

  End Sub

  Private Sub PopulateTableValidations()

    Dim ds = MetadataDb.ExecStoredProcedure("spadmin_getvalidations", Nothing)

    For Each row As DataRow In ds.Tables(0).Rows

      Dim table As Table = Tables.GetById(row.Item("tableid"))

      Dim validation As New Validation
      validation.ValidationType = row.Item("validationtype").ToString
      validation.Column = table.Columns.GetById(row.Item("columnid").ToString)

      table.Validations.Add(validation)
    Next

  End Sub

	Private Sub PopulateTableTriggerCode()

		Dim ds = MetadataDb.ExecStoredProcedure("spadmin_gettriggercode", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(row.Item("tableid"))

			Dim trigger As New TriggerCode
			trigger.Name = row.Item("Name").ToString
			trigger.CodePosition = CType(row.Item("CodePosition"), TriggerCodePosition)
			trigger.Content = row.Item("Content").ToString

			table.CodeTriggers.Add(trigger)
		Next

	End Sub

	Private Sub PopulateTableRelations()

		Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getrelations", Nothing)

		Dim table As Table
		Dim relation As Relation

		For Each row As DataRow In ds.Tables(0).Rows

			'add the relationshop from the parents perspective
			table = Tables.GetById(CInt(row.Item("parentid")))

			relation = New Relation
			relation.RelationshipType = RelationshipType.Child
			relation.ParentId = row.Item("parentid").ToString
			relation.ChildId = row.Item("childid").ToString
			relation.Name = row.Item("childname").ToString

			table.Relations.Add(relation)

			'add the relationshop from the childs perspective
			table = Tables.GetById(CInt(row.Item("childid")))

			relation = New Relation
			relation.RelationshipType = RelationshipType.Parent
			relation.ParentId = row.Item("parentid").ToString
			relation.ChildId = row.Item("childid").ToString
			relation.Name = row.Item("parentname").ToString

			table.Relations.Add(relation)
		Next

	End Sub

	Private Sub PopulateTableExpressions()

		Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getexpressions", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

			Dim expression As New Expression
			expression.Id = row.Item("id").ToString
			expression.SubType = ComponentTypes.Expression
			expression.TableId = CInt(row.Item("tableid"))
			expression.Name = row.Item("name").ToString
			expression.ExpressionType = row.Item("type").ToString
			expression.SchemaName = "dbo"
			expression.Description = row.Item("description").ToString
			expression.State = row.Item("state")
			expression.ReturnType = row.Item("returntype")
			expression.Size = row.Item("size")
			expression.Decimals = row.Item("decimals")
			expression.BaseTable = table
			expression.BaseExpression = expression
			expression.Level = 1

			expression.Components = LoadComponents(expression, ComponentTypes.Expression)

			table.Expressions.Add(expression)

			' Copy of all the expressions (sometimes calcs get their table references screwed up in the System Manager after being table copied.
			Expressions.Add(expression)

		Next

	End Sub

	Private Sub PopulateTableViews()

		Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getviews", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

			Dim view As New View
			view.Id = row.Item("id").ToString
			view.Name = row.Item("name").ToString
			view.Description = row.Item("description").ToString

			view.Table = table
			view.Filter = table.Expressions.GetById(CInt(row.Item("filterid")))

			table.Views.Add(view)
		Next

	End Sub

	Private Sub PopulateTableViewItems()

		Dim objDataset As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getviewitems", Nothing)

		For Each row As DataRow In objDataset.Tables(0).Rows

			Dim viewId As Integer = CInt(row.Item("viewid"))
			Dim view As View = Tables.SelectMany(Function(t) t.Views).Where(Function(o) o.Id = viewId).FirstOrDefault

			Dim column As Column = view.Table.Columns.GetById(CInt(row.Item("columnid")))
			If column IsNot Nothing Then
				view.Columns.Add(column)
			End If
		Next

	End Sub

	Private Sub PopulateTableRecordDescriptions()

		Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getdescriptions", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

			Dim description As New RecordDescription
			description.Id = row.Item("id").ToString
			description.Name = row.Item("name").ToString
			description.SchemaName = "dbo"
			description.Description = row.Item("description").ToString
			description.State = row.Item("state")
			description.ReturnType = row.Item("returntype")
			description.Size = row.Item("size")
			description.Decimals = row.Item("decimals")
			description.BaseTable = table
			description.BaseExpression = description

			description.Components = LoadComponents(description, ComponentTypes.Expression)

			table.RecordDescription = description
		Next

	End Sub

	Private Function LoadComponents(ByVal expression As Component, ByVal type As ComponentTypes) As ICollection(Of Component)

		If _componentfunction Is Nothing Then
			_componentfunction = MetadataDb.ExecStoredProcedure("spadmin_getcomponent_function", Nothing)
		End If

		If _componentbase Is Nothing Then
			_componentbase = MetadataDb.ExecStoredProcedure("spadmin_getcomponent_base", Nothing)
		End If

		Dim rows As DataRow()

		Select Case type
			Case ComponentTypes.Function
				rows = _componentfunction.Tables(0).Select("ExpressionID = " & expression.Id)
			Case ComponentTypes.Calculation
				rows = _componentbase.Tables(0).Select("ExpressionID = " & expression.CalculationId)
			Case Else
				rows = _componentbase.Tables(0).Select("ExpressionID = " & expression.Id)
		End Select

		Dim collection As New Collection(Of Component)

		For Each row As DataRow In rows

			Dim component As New Component
			component.Parent = expression
			component.Id = row.Item("componentid").ToString
			component.SubType = row.Item("subtype")
			component.Name = row.Item("name")
			component.ReturnType = row.Item("returntype")
			component.FunctionId = row.Item("functionid").ToString
			component.OperatorId = row.Item("operatorid").ToString
			component.TableId = row.Item("tableid").ToString
			component.ColumnId = row.Item("columnid").ToString
			component.ChildRowDetails.RowSelection = row.Item("columnaggregiatetype").ToString
			component.ChildRowDetails.RowNumber = row.Item("specificline").ToString
			component.ChildRowDetails.FilterId = row.Item("columnfilterid").ToString
			component.ChildRowDetails.OrderId = row.Item("columnorderid").ToString
			component.IsColumnByReference = row.Item("iscolumnbyreference").ToString
			component.CalculationId = row.Item("calculationid").ToString
			component.FilterId = row.Item("filterid").ToString
			component.ValueType = row.Item("valuetype").ToString
			component.Level = expression.Level + 1

			Select Case component.ValueType
				Case ComponentValueTypes.Date
					component.ValueDate = row.Item("valuedate")
				Case ComponentValueTypes.Logic
					component.ValueLogic = row.Item("valuelogic").ToString
				Case ComponentValueTypes.Numeric
					component.ValueNumeric = row.Item("valuenumeric").ToString
				Case ComponentValueTypes.String
					component.ValueString = row.Item("valuestring").ToString
			End Select

			component.LookupTableId = row.Item("lookuptableid").ToString
			component.LookupColumnId = row.Item("lookupcolumnid").ToString

			component.BaseExpression = expression.BaseExpression

			Select Case component.SubType
				Case ComponentTypes.Function
					component.Components = LoadComponents(component, ComponentTypes.Function)
				Case ComponentTypes.Expression, ComponentTypes.Calculation
					component.Components = LoadComponents(component, component.SubType)
			End Select

			collection.Add(component)
		Next

		Return collection

	End Function

	Private Sub PopulateFusionMessages()

		Dim ds As DataSet = CommitDb.ExecStoredProcedure("fusion.spGetMessageDefinitions", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

			If Not table Is Nothing Then
				Dim message As New FusionMessage
				message.Id = row.Item("id").ToString
				message.Name = row.Item("name").ToString
				message.StopDeletion = row.Item("stopdeletion").ToString
				message.ByPassValidation = row.Item("bypassvalidation").ToString

				table.FusionMessages.Add(message)
			End If
		Next

	End Sub

	Private Sub PopulateTableMasks()

		Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getmasks", Nothing)

		For Each row As DataRow In ds.Tables(0).Rows

			Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

			Dim mask As New Mask
			mask.Id = row.Item("id").ToString
			mask.Name = row.Item("name").ToString
			mask.AssociatedColumn = table.Columns(0)							' No need for masks to be aware of calling column!
			mask.SchemaName = "dbo"
			mask.Description = row.Item("description").ToString
			mask.State = row.Item("state")
			mask.ReturnType = row.Item("returntype")
			mask.Size = row.Item("size")
			mask.Decimals = row.Item("decimals")
			mask.BaseTable = table
			mask.BaseExpression = mask

			mask.Components = LoadComponents(mask, ComponentTypes.Expression)

			table.Masks.Add(mask)
		Next

	End Sub

	Private Function NullSafe(ByVal row As DataRow, ByVal columnName As String, ByVal defaultValue As Object) As Object

		If row.IsNull(columnName) Then
			Return defaultValue
		Else
			Return row.Item(columnName)
		End If

	End Function

End Module
