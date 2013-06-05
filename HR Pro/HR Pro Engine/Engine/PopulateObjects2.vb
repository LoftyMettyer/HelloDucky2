Option Strict Off

Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects2

    Public Sub PopulateThings()
      componentfunction = Nothing
      componentbase = Nothing
      Globals.Tables.Clear()
      Globals.Workflows.Clear()
      PopulateTables()
      PopulateTableRelations()
      PopulateColumns()
      PopulateTableOrders()
      PopulateTableOrderItems()
      PopulateTableExpressions()
      PopulateTableViews()
      PopulateTableViewItems()
      PopulateTableValidations()
      PopulateTableRecordDescriptions()
      PopulateTableMasks()
    End Sub

    Public Sub PopulateTables()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", Nothing)

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

        Globals.Tables.Add(table)
      Next

    End Sub

    Public Sub PopulateColumns()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim column As New Column
        column.Table = table

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

        table.Columns.Add(column)
      Next

    End Sub

    Public Sub PopulateTableOrders()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorders2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim tableOrder = New TableOrder
        tableOrder.ID = row.Item("orderid").ToString
        tableOrder.Name = row.Item("name").ToString
        tableOrder.SubType = row.Item("type").ToString
        tableOrder.Table = table

        table.TableOrders.Add(tableOrder)
      Next

    End Sub

    Public Sub PopulateTableOrderItems()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorderitems2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim orderId As Integer = CInt(row.Item("orderid"))
        Dim tableOrder As TableOrder = Globals.Tables.SelectMany(Function(t) t.TableOrders).Where(Function(o) o.ID = orderId).FirstOrDefault

        Dim orderItem As New TableOrderItem
        orderItem.ID = row.Item("orderid").ToString
        orderItem.ColumnType = row.Item("type")
        orderItem.Sequence = row.Item("sequence")
        orderItem.Ascending = row.Item("ascending")

        orderItem.TableOrder = tableOrder
        orderItem.Column = tableOrder.Table.Columns.GetById(row.Item("columnid").ToString)

        tableOrder.TableOrderItems.Add(orderItem)
      Next

    End Sub

    Public Sub PopulateTableValidations()

      Dim ds = Globals.MetadataDB.ExecStoredProcedure("spadmin_getvalidations2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim validation As New Validation
        validation.ValidationType = row.Item("validationtype").ToString
        validation.Column = table.Columns.GetById(row.Item("columnid").ToString)

        table.Validations.Add(validation)
      Next

    End Sub

    Public Sub PopulateTableRelations()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getrelations2", Nothing)

      Dim table As Table
      Dim relation As Relation

      For Each row As DataRow In ds.Tables(0).Rows

        'add the relationshop from the parents perspective
        table = Globals.Tables.GetById(row.Item("parentid").ToString)

        relation = New Relation
        relation.RelationshipType = ScriptDB.RelationshipType.Child
        relation.ParentID = row.Item("parentid").ToString
        relation.ChildID = row.Item("childid").ToString
        relation.Name = row.Item("childname").ToString

        table.Relations.Add(relation)

        'add the relationshop from the childs perspective
        table = Globals.Tables.GetById(row.Item("childid").ToString)

        relation = New Relation
        relation.RelationshipType = ScriptDB.RelationshipType.Parent
        relation.ParentID = row.Item("parentid").ToString
        relation.ChildID = row.Item("childid").ToString
        relation.Name = row.Item("parentname").ToString

        table.Relations.Add(relation)
      Next

    End Sub

    Public Sub PopulateTableExpressions()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim expression As New Expression
        expression.ID = row.Item("id").ToString
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

        expression.Components = Things.LoadComponents(expression, ScriptDB.ComponentTypes.Expression)

        table.Expressions.Add(expression)
      Next

    End Sub

    Public Sub PopulateTableViews()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviews2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Things.Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim view As New View
        view.ID = row.Item("id").ToString
        view.Name = row.Item("name").ToString
        view.Description = row.Item("description").ToString

        view.Table = table
        view.Filter = table.Expressions.GetById(row.Item("filterid").ToString)

        table.Views.Add(view)
      Next

    End Sub

    Public Sub PopulateTableViewItems()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviewitems2", Nothing)

      For Each row As DataRow In objDataset.Tables(0).Rows

        Dim viewId As Integer = CInt(row.Item("viewid"))
        Dim view As View = Globals.Tables.SelectMany(Function(t) t.Views).Where(Function(o) o.ID = viewId).FirstOrDefault

        Dim column As Column = view.Table.Columns.GetById(row.Item("columnid").ToString)
        If column IsNot Nothing Then
          view.Columns.Add(column)
        End If
      Next

    End Sub

    Public Sub PopulateTableRecordDescriptions()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getdescriptions2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim description As New RecordDescription
        description.ID = row.Item("id").ToString
        description.Name = row.Item("name").ToString
        description.SchemaName = "dbo"
        description.Description = row.Item("description").ToString
        description.State = row.Item("state")
        description.ReturnType = row.Item("returntype")
        description.Size = row.Item("size")
        description.Decimals = row.Item("decimals")
        description.BaseTable = table
        description.BaseExpression = description

        description.Components = LoadComponents(description, ScriptDB.ComponentTypes.Expression)

        table.RecordDescription = description
      Next

    End Sub

    Private componentfunction As DataSet
    Private componentbase As DataSet

    Public Function LoadComponents(ByVal expression As Component, ByVal Type As ScriptDB.ComponentTypes) As List(Of Component)

      If componentfunction Is Nothing Then
        componentfunction = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_function2", Nothing)
      End If

      If componentbase Is Nothing Then
        componentbase = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base2", Nothing)
      End If

      Dim rows As DataRow()

      Select Case Type
        Case ScriptDB.ComponentTypes.Function
          rows = componentfunction.Tables(0).Select("ExpressionID = " & expression.ID)
        Case ScriptDB.ComponentTypes.Calculation
          rows = componentbase.Tables(0).Select("ExpressionID = " & expression.CalculationID)
        Case Else
          rows = componentbase.Tables(0).Select("ExpressionID = " & expression.ID)
      End Select

      Dim collection As New List(Of Component)

      For Each row As DataRow In rows

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

        component.BaseExpression = expression.BaseExpression

        Select Case component.SubType
          Case ScriptDB.ComponentTypes.Function
            component.Components = Things.LoadComponents(component, ScriptDB.ComponentTypes.Function)
          Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
            component.Components = Things.LoadComponents(component, component.SubType)
        End Select

        collection.Add(component)
      Next

      Return collection

    End Function

    Public Sub PopulateTableMasks()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmasks2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim mask As New Things.Mask
        mask.ID = row.Item("id").ToString
        mask.Name = row.Item("name").ToString
        mask.AssociatedColumn = table.Columns.GetById(row.Item("columnid").ToString)
        mask.SchemaName = "dbo"
        mask.Description = row.Item("description").ToString
        mask.State = row.Item("state")
        mask.ReturnType = row.Item("returntype")
        mask.Size = row.Item("size")
        mask.Decimals = row.Item("decimals")
        mask.BaseTable = table
        mask.BaseExpression = mask

        mask.Components = Things.LoadComponents(mask, ScriptDB.ComponentTypes.Expression)

        table.Masks.Add(mask)
      Next

    End Sub

    Public Sub PopulateScreens()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getscreens", Nothing)

      For Each row In objDataset.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim screen As Screen = New Things.Screen
        screen.ID = row.Item("id").ToString
        screen.Name = row.Item("name").ToString
        screen.SchemaName = "dbo"
        screen.Description = row.Item("description").ToString
        screen.State = row.Item("state")
        screen.Table = table

        table.Screens.Add(screen)
      Next

    End Sub

    Public Sub PopulateWorkflows()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getworkflows", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim workflow As New Things.Workflow
        workflow.ID = row.Item("id").ToString
        workflow.Name = row.Item("name").ToString
        workflow.SchemaName = "dbo"
        workflow.Description = row.Item("description").ToString
        workflow.InitiationType = row.Item("initiationType").ToString
        workflow.Enabled = row.Item("enabled").ToString
        workflow.State = row.Item("state")
        workflow.Table = table

        table.Workflows.Add(workflow)
      Next

    End Sub

    Private Function NullSafe(ByVal ObjectData As System.Data.DataRow, ByVal ColumnName As String, ByVal DefaultValue As Object) As Object

      If ObjectData.IsNull(ColumnName) Then
        Return DefaultValue
      Else
        Return ObjectData.Item(ColumnName)
      End If

    End Function

  End Module

End Namespace
