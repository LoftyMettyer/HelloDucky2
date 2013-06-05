﻿Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects2

    Public Sub PopulateThings2()
      PopulateTables()
      PopulateColumns()
      PopulateTableOrders()
      PopulateTableOrderItems()
      PopulateTableValidations()
      PopulateTableRelations()
      PopulateTableExpressions()
      PopulateTableViews()
      PopulateTableViewItems()
      PopulateTableRecordDescriptions()
      PopulateTableMasks()
    End Sub

    Public Sub PopulateTables()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As New Table
        table.ID = row.Item("id").ToString
        table.TableType = row.Item("tabletype").ToString
        table.Name = row.Item("name").ToString
        table.SchemaName = "dbo"
        table.IsRemoteView = row.Item("isremoteview")
        table.AuditInsert = row.Item("auditinsert").ToString
        table.AuditDelete = row.Item("auditdelete").ToString
        table.DefaultEmailID = row.Item("defaultemailid").ToString
        table.DefaultOrderID = row.Item("defaultorderid").ToString
        table.State = row.Item("state")
        table.Root = table

        table.Objects.Parent = table
        table.Objects.Root = table.Root

        Globals.Things.Add(table)
      Next

      ' Objects with no table attached
      Dim emptyTable = New Table With {.ID = 0, .Name = "System Objects"}
      emptyTable.Objects.Parent = emptyTable
      emptyTable.Objects.Root = emptyTable.Root
      Globals.Things.Add(emptyTable)

    End Sub

    Public Sub PopulateColumns()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim column As New Column
        column.ID = row.Item("id").ToString
        column.Name = row.Item("name").ToString
        column.SchemaName = "dbo"
        column.Description = row.Item("description").ToString
        column.State = row.Item("state")

        column.DefaultCalcID = NullSafe(row, "defaultcalcid", 0).ToString
        column.DefaultValue = row.Item("defaultvalue").ToString
        column.CalcID = row.Item("calcid").ToString
        column.DataType = row.Item("datatype")
        column.Size = row.Item("size")
        column.Decimals = row.Item("decimals")
        column.Audit = row.Item("audit")
        column.Mandatory = row.Item("mandatory")
        column.Multiline = row.Item("multiline")
        column.IsReadOnly = row.Item("isreadonly")
        column.CaseType = row.Item("case").ToString
        column.CalculateIfEmpty = row.Item("calculateifempty")
        column.TrimType = NullSafe(row, "trimming", 0).ToString
        column.Alignment = row.Item("alignment").ToString
        column.Table = table
        column.Parent = table

        table.Objects.Add(column)
      Next

    End Sub

    Public Sub PopulateTableOrders()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorders2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim tableOrder = New TableOrder
        tableOrder.ID = row.Item("orderid").ToString
        tableOrder.Name = row.Item("name").ToString
        tableOrder.SubType = row.Item("type").ToString
        tableOrder.Parent = table

        table.Objects.Add(tableOrder)
      Next

    End Sub

    Public Sub PopulateTableOrderItems()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorderitems2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim orderId As String = row.Item("orderid").ToString
        Dim tableOrder As TableOrder = Globals.Things.OfType(Of Table).SelectMany(Function(t) t.Objects.OfType(Of TableOrder)()).FirstOrDefault(Function(o) o.ID = orderId)

        Dim orderItem As New TableOrderItem
        orderItem.ID = row.Item("orderid").ToString
        orderItem.ColumnType = row.Item("type")
        orderItem.Sequence = row.Item("sequence")
        orderItem.Ascending = row.Item("ascending")
        orderItem.Column = tableOrder.Parent.GetObject(Type.Column, row.Item("columnid").ToString)

        tableOrder.Objects.Add(orderItem)
      Next

    End Sub

    Public Sub PopulateTableValidations()

      Dim ds = Globals.MetadataDB.ExecStoredProcedure("spadmin_getvalidations2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim validation As New Validation
        validation.ValidationType = row.Item("validationtype").ToString
        validation.Column = CType(table, Table).Column(row.Item("columnid").ToString)

        table.Objects.Add(validation)
      Next

    End Sub

    Public Sub PopulateTableRelations()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getrelations2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        'TODO previous sp returns where parent or where child
        Dim table As Table = Nothing

        Dim relation As New Relation
        relation.RelationshipType = row.Item("relationship").ToString


        'Select Case objRelation.RelationshipType
        '  Case ScriptDB.RelationshipType.Parent
        '    objRelation.Parent = Table.Objects.Parent

        '  Case ScriptDB.RelationshipType.Child
        '    objRelation.Parent = Table

        'End Select

        relation.Parent = table
        relation.ParentID = row.Item("parentid").ToString
        relation.ChildID = row.Item("childid").ToString
        relation.Name = row.Item("name").ToString

        'table.Objects.Add(relation)
      Next

    End Sub

    Public Sub PopulateTableExpressions()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions2", New Connectivity.Parameters)

      For Each row As DataRow In objDataset.Tables(0).Rows

        Dim table As Things.Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim expression As New Expression
        expression.ID = row.Item("id").ToString
        expression.Parent = table
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

        expression.Objects = Things.LoadComponents2(expression, ScriptDB.ComponentTypes.Expression)

        table.Objects.Add(expression)
      Next

    End Sub

    Public Sub PopulateTableViews()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviews2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim view As New View
        view.ID = row.Item("id").ToString
        view.Name = row.Item("name").ToString
        view.Description = row.Item("description").ToString
        view.Parent = table
        view.Filter = table.GetObject(Type.Expression, row.Item("filterid").ToString)

        table.Objects.Add(view)
      Next

    End Sub

    Public Sub PopulateTableViewItems()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviewitems2", New Connectivity.Parameters)

      For Each row As DataRow In objDataset.Tables(0).Rows

        Dim viewId As String = row.Item("viewid").ToString
        Dim view As View = Globals.Things.OfType(Of Table).SelectMany(Function(t) t.Objects.OfType(Of View)()).FirstOrDefault(Function(o) o.ID = viewId)

        Dim column As Column = view.Parent.GetObject(Type.Column, row.Item("columnid").ToString)
        If column IsNot Nothing Then
          view.Objects.Add(column)
        End If
      Next

    End Sub

    Public Sub PopulateTableRecordDescriptions()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getdescriptions2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim description As New RecordDescription
        description.ID = row.Item("id").ToString
        description.Parent = table
        description.Name = row.Item("name").ToString
        description.SchemaName = "dbo"
        description.Description = row.Item("description").ToString
        description.State = row.Item("state")
        description.ReturnType = row.Item("returntype")
        description.Size = row.Item("size")
        description.Decimals = row.Item("decimals")
        description.BaseTable = table
        description.BaseExpression = description

        description.Objects = LoadComponents2(description, ScriptDB.ComponentTypes.Expression)

        table.Objects.Add(description)
      Next

    End Sub

    Public Function LoadComponents2(ByVal expression As Component, ByVal Type As ScriptDB.ComponentTypes) As Things.Collections.Generic

      Dim collection As New Things.Collections.Generic

      collection.Parent = expression

      Dim objDataset As DataSet
      Dim objParameters As New Connectivity.Parameters

      Select Case Type
        Case ScriptDB.ComponentTypes.Function
          objParameters.Add("@expressionid", expression.ID)
          objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_function", objParameters)

        Case ScriptDB.ComponentTypes.Calculation
          objParameters.Add("@expressionid", expression.CalculationID)
          'objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_calculation", objParameters)
          objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

          '   Case ScriptDB.ComponentTypes.Expression
          '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

        Case Else
          objParameters.Add("@expressionid", expression.ID)
          objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)
      End Select

      For Each row As DataRow In objDataset.Tables(0).Rows
        Dim component As New Things.Component
        component.ID = row.Item("componentid").ToString
        component.SubType = row.Item("subtype")
        component.Name = row.Item("name")
        component.ReturnType = row.Item("returntype")
        component.FunctionID = row.Item("functionid").ToString
        component.OperatorID = row.Item("operatorid").ToString
        component.TableID = row.Item("tableid").ToString
        component.ColumnID = row.Item("columnid").ToString
        component.ChildRowDetails.RowSelection = row.Item("columnaggregiatetype").ToString
        component.ChildRowDetails.RowNumber = row.Item("specificline").ToString
        component.ChildRowDetails.FilterID = row.Item("columnfilterid").ToString
        component.ChildRowDetails.OrderID = row.Item("columnorderid").ToString
        component.IsColumnByReference = row.Item("iscolumnbyreference").ToString
        component.CalculationID = row.Item("calculationid").ToString
        component.ValueType = row.Item("valuetype").ToString

        Select Case component.ValueType
          Case ScriptDB.ComponentValueTypes.Date
            component.ValueDate = row.Item("valuedate")
          Case ScriptDB.ComponentValueTypes.Logic
            component.ValueLogic = row.Item("valuelogic").ToString
          Case ScriptDB.ComponentValueTypes.Numeric
            component.ValueNumeric = row.Item("valuenumeric").ToString
          Case ScriptDB.ComponentValueTypes.String
            component.ValueString = row.Item("valuestring").ToString
        End Select

        component.LookupTableID = row.Item("lookuptableid").ToString
        component.LookupColumnID = row.Item("lookupcolumnid").ToString

        component.Root = expression.Root
        component.BaseExpression = expression.BaseExpression

        Select Case component.SubType
          Case ScriptDB.ComponentTypes.Function
            component.Objects = Things.LoadComponents2(component, ScriptDB.ComponentTypes.Function)
          Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
            component.Objects = Things.LoadComponents2(component, component.SubType)
        End Select

        collection.Add(component)
      Next

      Return collection

    End Function

    Public Sub PopulateTableMasks()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmasks2", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim mask As New Things.Mask
        mask.ID = row.Item("id").ToString
        mask.Parent = table
        mask.Name = row.Item("name").ToString
        mask.AssociatedColumn = table.Column(row.Item("columnid").ToString)
        mask.SchemaName = "dbo"
        mask.Description = row.Item("description").ToString
        mask.State = row.Item("state")
        mask.ReturnType = row.Item("returntype")
        mask.Size = row.Item("size")
        mask.Decimals = row.Item("decimals")
        mask.BaseTable = table
        mask.BaseExpression = mask

        mask.Objects = Things.LoadComponents2(mask, ScriptDB.ComponentTypes.Expression)

        table.Objects.Add(mask)
      Next

    End Sub

    'TODO use or not? still used by HCM is that different to my one?
    Public Sub PopulateExpressions()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Dim objExpression As Things.Expression
      Dim objTable As Things.Table

      objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions", objParameters)

      For Each objRow In objDataset.Tables(0).Rows

        objExpression = New Things.Expression
        objExpression.ID = objRow.Item("id").ToString

        objExpression.Name = objRow.Item("name").ToString
        objExpression.ExpressionType = objRow.Item("type").ToString
        objExpression.SchemaName = "dbo"
        objExpression.Description = objRow.Item("description").ToString
        objExpression.State = objRow.Item("state")
        objExpression.ReturnType = objRow.Item("returntype")
        objExpression.Size = objRow.Item("size")
        objExpression.Decimals = objRow.Item("decimals")
        objExpression.BaseExpression = objExpression

        objTable = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString)
        objExpression.BaseTable = objTable
        objExpression.Parent = objTable

        objTable.Objects.Add(objExpression)

      Next

    End Sub

    Public Sub PopulateScreens()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getscreens", New Connectivity.Parameters)

      For Each row In objDataset.Tables(0).Rows

        Dim table As Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim screen As Screen = New Things.Screen
        screen.ID = row.Item("id").ToString
        screen.Name = row.Item("name").ToString
        screen.SchemaName = "dbo"
        screen.Description = row.Item("description").ToString
        screen.State = row.Item("state")
        screen.Table = table
        screen.Parent = table

        table.Objects.Add(screen)
      Next

    End Sub

    Public Sub PopulateWorkflows()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getworkflows", New Connectivity.Parameters)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Things.GetObject(Type.Table, row.Item("tableid").ToString)

        Dim workflow As New Things.Workflow
        workflow.ID = row.Item("id").ToString
        workflow.Name = row.Item("name").ToString
        workflow.SchemaName = "dbo"
        workflow.Description = row.Item("description").ToString
        workflow.InitiationType = row.Item("initiationType").ToString
        workflow.Enabled = row.Item("enabled").ToString
        workflow.State = row.Item("state")
        workflow.Table = table

        table.Objects.Add(workflow)
      Next

    End Sub

    Private Function NullSafe(ByVal ObjectData As System.Data.DataRow, ByVal ColumnName As String, ByRef DefaultValue As Object)

      If ObjectData.IsNull(ColumnName) Then
        Return DefaultValue
      Else
        Return ObjectData.Item(ColumnName)
      End If

    End Function

  End Module

End Namespace
