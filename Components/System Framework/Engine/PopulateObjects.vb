Option Strict Off

Namespace Things

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

    Public Function GetCodeLibraryDependancies(ByVal codeLibrary As CodeLibrary) As ICollection(Of Setting)

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

    Public Sub PopulateTables()

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
        table.State = row.Item("state")

        Tables.Add(table)
      Next

    End Sub

    Public Sub PopulateTableColumns()

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
        ErrorLog.Add(ErrorHandler.Section.LoadingData, "PopulateTableColumns", ErrorHandler.Severity.Error, "Table or column not found", "Error in system metatdata - table or column not found")

      End Try

    End Sub

    Public Sub PopulateTableOrders()

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

    Public Sub PopulateTableOrderItems()

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
          ErrorLog.Add(ErrorHandler.Section.LoadingData, "OrderItems", ErrorHandler.Severity.Warning, "Order not found", "order " & CStr(orderId) & " not found")

        End If
      Next

    End Sub

    Public Sub PopulateTableValidations()

      Dim ds = MetadataDb.ExecStoredProcedure("spadmin_getvalidations", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Tables.GetById(row.Item("tableid"))

        Dim validation As New Validation
        validation.ValidationType = row.Item("validationtype").ToString
        validation.Column = table.Columns.GetById(row.Item("columnid").ToString)

        table.Validations.Add(validation)
      Next

    End Sub

    Public Sub PopulateTableRelations()

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

    Public Sub PopulateTableExpressions()

      Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getexpressions", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

        Dim expression As New Expression
        expression.Id = row.Item("id").ToString
        expression.SubType = ScriptDB.ComponentTypes.Expression
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

        expression.Components = LoadComponents(expression, ScriptDB.ComponentTypes.Expression)

        table.Expressions.Add(expression)

        ' Copy of all the expressions (sometimes calcs get their table references screwed up in the System Manager after being table copied.
        Expressions.Add(expression)

      Next

    End Sub

    Public Sub PopulateTableViews()

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

    Public Sub PopulateTableViewItems()

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

    Public Sub PopulateTableRecordDescriptions()

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

        description.Components = LoadComponents(description, ScriptDB.ComponentTypes.Expression)

        table.RecordDescription = description
      Next

    End Sub

    Public Function LoadComponents(ByVal expression As Component, ByVal type As ScriptDB.ComponentTypes) As ICollection(Of Component)

      If _componentfunction Is Nothing Then
        _componentfunction = MetadataDb.ExecStoredProcedure("spadmin_getcomponent_function", Nothing)
      End If

      If _componentbase Is Nothing Then
        _componentbase = MetadataDb.ExecStoredProcedure("spadmin_getcomponent_base", Nothing)
      End If

      Dim rows As DataRow()

      Select Case type
        Case ScriptDB.ComponentTypes.Function
          rows = _componentfunction.Tables(0).Select("ExpressionID = " & expression.Id)
        Case ScriptDB.ComponentTypes.Calculation
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
          Case ScriptDB.ComponentValueTypes.Date
            component.ValueDate = row.Item("valuedate")
          Case ScriptDB.ComponentValueTypes.Logic
            component.ValueLogic = row.Item("valuelogic").ToString
          Case ScriptDB.ComponentValueTypes.Numeric
            component.ValueNumeric = row.Item("valuenumeric").ToString
          Case ScriptDB.ComponentValueTypes.String
            component.ValueString = row.Item("valuestring").ToString
        End Select

        component.LookupTableId = row.Item("lookuptableid").ToString
        component.LookupColumnId = row.Item("lookupcolumnid").ToString

        component.BaseExpression = expression.BaseExpression

        Select Case component.SubType
          Case ScriptDB.ComponentTypes.Function
            component.Components = LoadComponents(component, ScriptDB.ComponentTypes.Function)
          Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
            component.Components = LoadComponents(component, component.SubType)
        End Select

        collection.Add(component)
      Next

      Return collection

    End Function

    Public Sub PopulateFusionMessages()

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

    Public Sub PopulateTableMasks()

      Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getmasks", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

        Dim mask As New Mask
        mask.Id = row.Item("id").ToString
        mask.Name = row.Item("name").ToString
        mask.AssociatedColumn = table.Columns(0)              ' No need for masks to be aware of calling column!
        mask.SchemaName = "dbo"
        mask.Description = row.Item("description").ToString
        mask.State = row.Item("state")
        mask.ReturnType = row.Item("returntype")
        mask.Size = row.Item("size")
        mask.Decimals = row.Item("decimals")
        mask.BaseTable = table
        mask.BaseExpression = mask

        mask.Components = LoadComponents(mask, ScriptDB.ComponentTypes.Expression)

        table.Masks.Add(mask)
      Next

    End Sub

    Public Sub PopulateScreens()

      Dim objDataset As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getscreens", Nothing)

      For Each row In objDataset.Tables(0).Rows

        Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

        Dim screen As Screen = New Screen
        screen.Id = row.Item("id").ToString
        screen.Name = row.Item("name").ToString
        screen.SchemaName = "dbo"
        screen.Description = row.Item("description").ToString
        screen.State = row.Item("state")
        screen.Table = table

        table.Screens.Add(screen)
      Next

    End Sub

    Public Sub PopulateWorkflows()

      Dim ds As DataSet = MetadataDb.ExecStoredProcedure("spadmin_getworkflows", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Tables.GetById(CInt(row.Item("tableid")))

        Dim workflow As New Workflow
        workflow.Id = row.Item("id").ToString
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
