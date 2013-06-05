Option Strict Off

Namespace Things

  <HideModuleName()>
  Public Module PopulateObjects

    Public Sub PopulateThings()
      componentfunction = Nothing
      componentbase = Nothing
      Globals.Tables.Clear()
      Globals.Workflows.Clear()

      PopulateTables()
      PopulateTableRelations()
      PopulateTableColumns()
      PopulateTableOrders()
      PopulateTableOrderItems()
      PopulateTableExpressions()
      PopulateTableViews()
      PopulateTableViewItems()
      PopulateTableValidations()
      PopulateTableRecordDescriptions()
      PopulateTableMasks()
    End Sub

    Public Sub PopulateSystemSettings()

      Globals.SystemSettings.Clear()

      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getsystemsettings", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows
        Dim setting As New Setting
        setting.Module = row("section").ToString
        setting.Parameter = row.Item("settingkey").ToString
        setting.Value = row.Item("settingvalue").ToString

        Globals.SystemSettings.Add(setting)
      Next

    End Sub

    Public Sub PopulateModuleSettings()

      Globals.ModuleSetup.Clear()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmodulesetup", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim setting As New Setting
        setting.Module = row.Item("modulekey").ToString
        setting.Parameter = row.Item("parameterkey").ToString

        If Not row.Item("value").ToString = "" Then
          Select Case row.Item("subtype").ToString
            Case 1
              setting.Table = Globals.Tables.GetById(row.Item("value").ToString)
            Case Else
              setting.Value = row.Item("value").ToString
          End Select

          Globals.ModuleSetup.Add(setting)
        End If
      Next

    End Sub

    Public Sub PopulateSystemThings()

      Globals.Operators.Clear()
      Globals.Functions.Clear()

      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcode", Nothing)
      For Each row As DataRow In ds.Tables(0).Rows

        Dim codeLibrary As New CodeLibrary
        codeLibrary.ID = row.Item("id").ToString
        codeLibrary.Name = row.Item("name").ToString
        codeLibrary.Code = row.Item("code").ToString
        codeLibrary.PreCode = row.Item("precode").ToString
        codeLibrary.AfterCode = row.Item("aftercode").ToString
        codeLibrary.ReturnType = row.Item("returntype").ToString
        codeLibrary.OperatorType = row.Item("operatortype").ToString
        codeLibrary.RecordIDRequired = row.Item("recordidrequired").ToString
        codeLibrary.CalculatePostAudit = row.Item("calculatepostaudit").ToString
        codeLibrary.IsUniqueCode = row.Item("isuniquecode").ToString
        codeLibrary.IsTimeDependant = row.Item("istimedependant").ToString
        codeLibrary.IsGetFieldFromDB = row.Item("isgetfieldfromdb").ToString
        codeLibrary.CaseCount = row.Item("casecount").ToString
        codeLibrary.MakeTypeSafe = row.Item("maketypesafe").ToString
        codeLibrary.OvernightOnly = row.Item("overnightonly").ToString
        codeLibrary.Tuning.Rating = row.Item("performancerating").ToString
        codeLibrary.DependsOnBankHoliday = row.Item("dependsonbankholiday").ToString
        codeLibrary.Dependancies = GetCodeLibraryDependancies(codeLibrary)

        If CBool(row.Item("isoperator").ToString) Then
          Globals.Operators.Add(codeLibrary)
        Else
          Globals.Functions.Add(codeLibrary)
        End If

      Next

    End Sub

    Public Function GetCodeLibraryDependancies(ByVal codeLibrary As CodeLibrary) As ICollection(Of Setting)

      Dim params As New Connectivity.Parameters
      params.Add("@componentid", codeLibrary.ID)
      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcodedependancies", params)

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

    Public Sub PopulateTables()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As New Table
        table.ID = row.Item("id").ToString
        table.TableType = row.Item("tabletype").ToString
        table.Name = row.Item("name").ToString
        table.SchemaName = "dbo"
        table.IsRemoteView = row.Item("isremoteview").ToString
        table.AuditInsert = row.Item("auditinsert").ToString
        table.AuditDelete = row.Item("auditdelete").ToString
        table.DefaultEmailID = row.Item("defaultemailid").ToString
        table.DefaultOrderID = row.Item("defaultorderid").ToString
        table.State = row.Item("state")

        Globals.Tables.Add(table)
      Next

    End Sub

    Public Sub PopulateTableColumns()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns", Nothing)

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

    End Sub

    Public Sub PopulateTableOrders()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorders", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim tableOrder = New TableOrder
        tableOrder.ID = row.Item("orderid").ToString
        tableOrder.Name = row.Item("name").ToString
        tableOrder.Table = table

        table.TableOrders.Add(tableOrder)
      Next

    End Sub

    Public Sub PopulateTableOrderItems()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorderitems", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim orderId As Integer = CInt(row.Item("orderid"))
        Dim tableOrder As TableOrder = Globals.Tables.SelectMany(Function(t) t.TableOrders).Where(Function(o) o.ID = orderId).FirstOrDefault

        Dim orderItem As New TableOrderItem
        orderItem.ID = row.Item("orderid").ToString
        orderItem.ColumnType = row.Item("type")
        orderItem.Sequence = row.Item("sequence")
        orderItem.Ascending = row.Item("ascending")

        orderItem.TableOrder = tableOrder
        orderItem.Column = tableOrder.Table.Columns.GetById(CInt(row.Item("columnid")))

        tableOrder.Items.Add(orderItem)
      Next

    End Sub

    Public Sub PopulateTableValidations()

      Dim ds = Globals.MetadataDB.ExecStoredProcedure("spadmin_getvalidations", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid"))

        Dim validation As New Validation
        validation.ValidationType = row.Item("validationtype").ToString
        validation.Column = table.Columns.GetById(row.Item("columnid").ToString)

        table.Validations.Add(validation)
      Next

    End Sub

    Public Sub PopulateTableRelations()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getrelations", Nothing)

      Dim table As Table
      Dim relation As Relation

      For Each row As DataRow In ds.Tables(0).Rows

        'add the relationshop from the parents perspective
        table = Globals.Tables.GetById(CInt(row.Item("parentid")))

        relation = New Relation
        relation.RelationshipType = RelationshipType.Child
        relation.ParentID = row.Item("parentid").ToString
        relation.ChildID = row.Item("childid").ToString
        relation.Name = row.Item("childname").ToString

        table.Relations.Add(relation)

        'add the relationshop from the childs perspective
        table = Globals.Tables.GetById(CInt(row.Item("childid")))

        relation = New Relation
        relation.RelationshipType = RelationshipType.Parent
        relation.ParentID = row.Item("parentid").ToString
        relation.ChildID = row.Item("childid").ToString
        relation.Name = row.Item("parentname").ToString

        table.Relations.Add(relation)
      Next

    End Sub

    Public Sub PopulateTableExpressions()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

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

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviews", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

        Dim view As New View
        view.ID = row.Item("id").ToString
        view.Name = row.Item("name").ToString
        view.Description = row.Item("description").ToString

        view.Table = table
        view.Filter = table.Expressions.GetById(CInt(row.Item("filterid")))

        table.Views.Add(view)
      Next

    End Sub

    Public Sub PopulateTableViewItems()

      Dim objDataset As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviewitems", Nothing)

      For Each row As DataRow In objDataset.Tables(0).Rows

        Dim viewId As Integer = CInt(row.Item("viewid"))
        Dim view As View = Globals.Tables.SelectMany(Function(t) t.Views).Where(Function(o) o.ID = viewId).FirstOrDefault

        Dim column As Column = view.Table.Columns.GetById(CInt(row.Item("columnid")))
        If column IsNot Nothing Then
          view.Columns.Add(column)
        End If
      Next

    End Sub

    Public Sub PopulateTableRecordDescriptions()

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getdescriptions", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

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

    Public Function LoadComponents(ByVal expression As Component, ByVal type As ScriptDB.ComponentTypes) As ICollection(Of Component)

      If componentfunction Is Nothing Then
        componentfunction = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_function", Nothing)
      End If

      If componentbase Is Nothing Then
        componentbase = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", Nothing)
      End If

      Dim rows As DataRow()

      Select Case type
        Case ScriptDB.ComponentTypes.Function
          rows = componentfunction.Tables(0).Select("ExpressionID = " & expression.ID)
        Case ScriptDB.ComponentTypes.Calculation
          rows = componentbase.Tables(0).Select("ExpressionID = " & expression.CalculationID)
        Case Else
          rows = componentbase.Tables(0).Select("ExpressionID = " & expression.ID)
      End Select

      Dim collection As New Collection(Of Component)

      For Each row As DataRow In rows

        Dim component As New Component
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

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmasks", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

        Dim mask As New Mask
        mask.ID = row.Item("id").ToString
        mask.Name = row.Item("name").ToString
        mask.AssociatedColumn = table.Columns.GetById(CInt(row.Item("columnid")))
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

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

        Dim screen As Screen = New Screen
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

        Dim table As Table = Globals.Tables.GetById(CInt(row.Item("tableid")))

        Dim workflow As New Workflow
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

#Region "Workflow"

    'Public Function LoadWorkflowElements(byval objWorkflow As Workflow) As Collection

    '  Dim objObjects As New Collection
    '  Dim objElement As WorkflowElement
    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow
    '  Dim objParameters As New Connectivity.Parameters

    '    ' Populate element
    '    objParameters.Add("@workflowid", objWorkflow.ID)
    '    objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelements", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objElement = New WorkflowElement
    '      objElement.ID = objRow.Item("elementid")
    '      objElement.SubType = objRow.Item("type")
    '      objElement.Caption = objRow.Item("caption")
    '      objElement.ConnectionPairID = objRow.Item("ConnectionPairID")
    '      objElement.LeftCoord = objRow.Item("LeftCoord")
    '      objElement.TopCoord = objRow.Item("TopCoord")
    '      objElement.DecisionCaptionType = objRow.Item("DecisionCaptionType")
    '      objElement.Identifier = objRow.Item("Identifier")
    '      objElement.TrueFlowIdentifier = objRow.Item("TrueFlowIdentifier")
    '      objElement.DataAction = objRow.Item("DataAction")
    '      objElement.DataTableID = objRow.Item("DataTableID")
    '      objElement.DataRecord = objRow.Item("DataRecord")
    '      objElement.EmailID = objRow.Item("EmailID")
    '      objElement.EmailRecord = objRow.Item("EmailRecord")
    '      objElement.WebFormBGColor = objRow.Item("WebFormBGColor")
    '      objElement.WebFormBGImageID = objRow.Item("WebFormBGImageID")
    '      objElement.WebFormBGImageLocation = objRow.Item("WebFormBGImageLocation")
    '      objElement.WebFormDefaultFontName = objRow.Item("WebFormDefaultFontName")
    '      objElement.WebFormDefaultFontSize = objRow.Item("WebFormDefaultFontSize")
    '      objElement.WebFormDefaultFontBold = objRow.Item("WebFormDefaultFontBold")
    '      objElement.WebFormDefaultFontItalic = objRow.Item("WebFormDefaultFontItalic")
    '      objElement.WebFormDefaultFontStrikeThru = objRow.Item("WebFormDefaultFontStrikeThru")
    '      objElement.WebFormDefaultFontUnderline = objRow.Item("WebFormDefaultFontUnderline")
    '      objElement.WebFormWidth = objRow.Item("WebFormWidth")
    '      objElement.RecSelWebFormIdentifier = objRow.Item("RecSelWebFormIdentifier")
    '      objElement.RecSelIdentifier = objRow.Item("RecSelIdentifier")
    '      objElement.SecondaryDataRecord = objRow.Item("SecondaryDataRecord")
    '      objElement.SecondaryRecSelWebFormIdentifier = objRow.Item("SecondaryRecSelWebFormIdentifier")
    '      objElement.SecondaryRecSelIdentifier = objRow.Item("SecondaryRecSelIdentifier")
    '      objElement.EmailSubject = objRow.Item("EmailSubject")
    '      objElement.TimeoutFrequency = objRow.Item("TimeoutFrequency")
    '      objElement.TimeoutPeriod = objRow.Item("TimeoutPeriod")
    '      objElement.DataRecordTable = objRow.Item("DataRecordTable")
    '      objElement.SecondaryDataRecordTable = objRow.Item("SecondaryDataRecordTable")
    '      objElement.TrueFlowType = objRow.Item("TrueFlowType")
    '      objElement.TrueFlowExprID = objRow.Item("TrueFlowExprID")
    '      objElement.DescriptionExprID = objRow.Item("DescriptionExprID")
    '      objElement.WebFormFGColor = objRow.Item("WebFormFGColor")
    '      objElement.DescHasWorkflowName = objRow.Item("DescHasWorkflowName")
    '      objElement.DescHasElementCaption = objRow.Item("DescHasElementCaption")
    '      objElement.EmailCCID = objRow.Item("EmailCCID")
    '      objElement.TimeoutExcludeWeekend = objRow.Item("TimeoutExcludeWeekend")
    '      objElement.CompletionMessageType = objRow.Item("CompletionMessageType")
    '      objElement.CompletionMessage = objRow.Item("CompletionMessage")
    '      objElement.SavedForLaterMessageType = objRow.Item("SavedForLaterMessageType")
    '      objElement.SavedForLaterMessage = objRow.Item("SavedForLaterMessage")
    '      objElement.FollowOnFormsMessageType = objRow.Item("FollowOnFormsMessageType")
    '      objElement.FollowOnFormsMessage = objRow.Item("FollowOnFormsMessage")
    '      objElement.Attachment_Type = objRow.Item("Attachment_Type")
    '      objElement.Attachment_File = objRow.Item("Attachment_File")
    '      objElement.Attachment_WFElementIdentifier = objRow.Item("Attachment_WFElementIdentifier")
    '      objElement.Attachment_WFValueIdentifier = objRow.Item("Attachment_WFValueIdentifier")
    '      objElement.Attachment_DBColumnID = objRow.Item("Attachment_DBColumnID")
    '      objElement.Attachment_DBRecord = objRow.Item("Attachment_DBRecord")
    '      objElement.Attachment_DBElement = objRow.Item("Attachment_DBElement")
    '      objElement.Attachment_DBValue = objRow.Item("Attachment_DBValue")

    '      objElement.Objects = Things.LoadWorkflowElementDetails(objElement)

    '      objObjects.Add(objElement)

    '    Next

    '  LoadWorkflowElements = objObjects


    'End Function

    'Public Function LoadWorkflowElementDetails(ByVal objWorkflowElement As WorkflowElement) As Collection

    '  Dim objObjects As New Collection
    '  Dim objElementColumn As WorkflowElementColumn
    '  '   Dim objElementItem As WorkflowElementItem

    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow
    '  Dim objParameters As New Connectivity.Parameters

    '    ' Populate element
    '    objParameters.Add("@elementid", objWorkflowElement.ID)
    '    objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelementcolumns", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objElementColumn = New WorkflowElementColumn
    '      objElementColumn.ID = objRow.Item("elementid")
    '      objElementColumn.ColumnID = objRow.Item("columnid")

    '      objElementColumn.ValueType = objRow.Item("valuetype")
    '      objElementColumn.Value = objRow.Item("value")
    '      objElementColumn.WFFormIdentifier = objRow.Item("wfformidentifier")
    '      objElementColumn.WFValueIdentifier = objRow.Item("wfvalueidentifier")
    '      objElementColumn.DBColumnID = objRow.Item("dbcolumnid")
    '      objElementColumn.DBRecord = objRow.Item("dbrecord")
    '      objElementColumn.CalcID = objRow.Item("calcid")

    '      objObjects.Add(objElementColumn)

    '    Next

    'End Function

#End Region

    Private Function NullSafe(ByVal row As DataRow, ByVal columnName As String, ByVal defaultValue As Object) As Object

      If row.IsNull(columnName) Then
        Return defaultValue
      Else
        Return row.Item(columnName)
      End If

    End Function

  End Module

End Namespace
