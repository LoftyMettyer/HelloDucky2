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
        setting.Module = row.Item("section").ToString
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
        setting.SubType = row.Item("subtype").ToString

        If Not row.Item("value").ToString = "" Then
          Select Case setting.SubType
            Case Type.Table
              setting.Table = Globals.Tables.GetById(row.Item("value").ToString)
            Case Type.Column
              setting.Value = row.Item("value").ToString
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

        If row.Item("isoperator") Then
          Globals.Operators.Add(codeLibrary)
        Else
          Globals.Functions.Add(codeLibrary)
        End If

      Next

    End Sub

    Public Function GetCodeLibraryDependancies(ByVal codeLibrary As CodeLibrary) As List(Of Setting)

      Dim params As New Connectivity.Parameters
      params.Add("@componentid", codeLibrary.ID)
      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcodedependancies", params)

      Dim dependancies As New List(Of Setting)

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
        table.IsRemoteView = row.Item("isremoteview")
        table.AuditInsert = row.Item("auditinsert").ToString
        table.AuditDelete = row.Item("auditdelete").ToString
        table.DefaultEmailID = row.Item("defaultemailid").ToString
        table.DefaultOrderID = row.Item("defaultorderid").ToString
        table.State = row.Item("state")

        Globals.Tables.Add(table)
      Next

    End Sub

    Public Sub PopulateTableColumns()

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

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

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

        tableOrder.Items.Add(orderItem)
      Next

    End Sub

    Public Sub PopulateTableValidations()

      Dim ds = Globals.MetadataDB.ExecStoredProcedure("spadmin_getvalidations2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

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

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

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

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

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

      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmasks2", Nothing)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

        Dim mask As New Mask
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

        Dim table As Table = Globals.Tables.GetById(row.Item("tableid").ToString)

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

    '  objObjects.Parent = objWorkflow
    '  objObjects.Root = objWorkflow.Root

    '  Try

    '    ' Populate element
    '    objParameters.Add("@workflowid", objWorkflow.ID)
    '    objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelements", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objElement = New WorkflowElement
    '      objElement.ID = objRow.Item("elementid").ToString
    '      objElement.SubType = objRow.Item("type")
    '      objElement.Caption = objRow.Item("caption")
    '      objElement.ConnectionPairID = objRow.Item("ConnectionPairID")
    '      objElement.LeftCoord = objRow.Item("LeftCoord")
    '      objElement.TopCoord = objRow.Item("TopCoord")
    '      objElement.DecisionCaptionType = objRow.Item("DecisionCaptionType").ToString
    '      objElement.Identifier = objRow.Item("Identifier").ToString
    '      objElement.TrueFlowIdentifier = objRow.Item("TrueFlowIdentifier").ToString
    '      objElement.DataAction = objRow.Item("DataAction").ToString
    '      objElement.DataTableID = objRow.Item("DataTableID").ToString
    '      objElement.DataRecord = objRow.Item("DataRecord").ToString
    '      objElement.EmailID = objRow.Item("EmailID").ToString
    '      objElement.EmailRecord = objRow.Item("EmailRecord").ToString
    '      objElement.WebFormBGColor = objRow.Item("WebFormBGColor").ToString
    '      objElement.WebFormBGImageID = objRow.Item("WebFormBGImageID").ToString
    '      objElement.WebFormBGImageLocation = objRow.Item("WebFormBGImageLocation").ToString
    '      objElement.WebFormDefaultFontName = objRow.Item("WebFormDefaultFontName").ToString
    '      objElement.WebFormDefaultFontSize = objRow.Item("WebFormDefaultFontSize").ToString
    '      objElement.WebFormDefaultFontBold = objRow.Item("WebFormDefaultFontBold").ToString
    '      objElement.WebFormDefaultFontItalic = objRow.Item("WebFormDefaultFontItalic").ToString
    '      objElement.WebFormDefaultFontStrikeThru = objRow.Item("WebFormDefaultFontStrikeThru").ToString
    '      objElement.WebFormDefaultFontUnderline = objRow.Item("WebFormDefaultFontUnderline").ToString
    '      objElement.WebFormWidth = objRow.Item("WebFormWidth").ToString
    '      objElement.RecSelWebFormIdentifier = objRow.Item("RecSelWebFormIdentifier").ToString
    '      objElement.RecSelIdentifier = objRow.Item("RecSelIdentifier").ToString
    '      objElement.SecondaryDataRecord = objRow.Item("SecondaryDataRecord").ToString
    '      objElement.SecondaryRecSelWebFormIdentifier = objRow.Item("SecondaryRecSelWebFormIdentifier").ToString
    '      objElement.SecondaryRecSelIdentifier = objRow.Item("SecondaryRecSelIdentifier").ToString
    '      objElement.EmailSubject = objRow.Item("EmailSubject").ToString
    '      objElement.TimeoutFrequency = objRow.Item("TimeoutFrequency").ToString
    '      objElement.TimeoutPeriod = objRow.Item("TimeoutPeriod").ToString
    '      objElement.DataRecordTable = objRow.Item("DataRecordTable").ToString
    '      objElement.SecondaryDataRecordTable = objRow.Item("SecondaryDataRecordTable").ToString
    '      objElement.TrueFlowType = objRow.Item("TrueFlowType").ToString
    '      objElement.TrueFlowExprID = objRow.Item("TrueFlowExprID").ToString
    '      objElement.DescriptionExprID = objRow.Item("DescriptionExprID").ToString
    '      objElement.WebFormFGColor = objRow.Item("WebFormFGColor").ToString
    '      objElement.DescHasWorkflowName = objRow.Item("DescHasWorkflowName").ToString
    '      objElement.DescHasElementCaption = objRow.Item("DescHasElementCaption").ToString
    '      objElement.EmailCCID = objRow.Item("EmailCCID").ToString
    '      objElement.TimeoutExcludeWeekend = objRow.Item("TimeoutExcludeWeekend").ToString
    '      objElement.CompletionMessageType = objRow.Item("CompletionMessageType").ToString
    '      objElement.CompletionMessage = objRow.Item("CompletionMessage").ToString
    '      objElement.SavedForLaterMessageType = objRow.Item("SavedForLaterMessageType").ToString
    '      objElement.SavedForLaterMessage = objRow.Item("SavedForLaterMessage").ToString
    '      objElement.FollowOnFormsMessageType = objRow.Item("FollowOnFormsMessageType").ToString
    '      objElement.FollowOnFormsMessage = objRow.Item("FollowOnFormsMessage").ToString
    '      objElement.Attachment_Type = objRow.Item("Attachment_Type").ToString
    '      objElement.Attachment_File = objRow.Item("Attachment_File").ToString
    '      objElement.Attachment_WFElementIdentifier = objRow.Item("Attachment_WFElementIdentifier").ToString
    '      objElement.Attachment_WFValueIdentifier = objRow.Item("Attachment_WFValueIdentifier").ToString
    '      objElement.Attachment_DBColumnID = objRow.Item("Attachment_DBColumnID").ToString
    '      objElement.Attachment_DBRecord = objRow.Item("Attachment_DBRecord").ToString
    '      objElement.Attachment_DBElement = objRow.Item("Attachment_DBElement").ToString
    '      objElement.Attachment_DBValue = objRow.Item("Attachment_DBValue").ToString

    '      objElement.Objects = Things.LoadWorkflowElementDetails(objElement)

    '      objObjects.Add(objElement)

    '    Next


    '  Catch ex As Exception
    '    Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.LoadingData, String.Empty, HRProEngine.ErrorHandler.Severity.Error, ex.Message, String.Empty)

    '  End Try

    '  LoadWorkflowElements = objObjects


    'End Function

    'Public Function LoadWorkflowElementDetails(ByVal objWorkflowElement As WorkflowElement) As Collection

    '  Dim objObjects As New Collection
    '  Dim objElementColumn As WorkflowElementColumn
    '  '   Dim objElementItem As WorkflowElementItem

    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow
    '  Dim objParameters As New Connectivity.Parameters

    '  objObjects.Parent = objWorkflowElement
    '  objObjects.Root = objWorkflowElement.Root

    '  Try

    '    ' Populate element
    '    objParameters.Add("@elementid", objWorkflowElement.ID)
    '    objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelementcolumns", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objElementColumn = New WorkflowElementColumn
    '      objElementColumn.ID = objRow.Item("elementid").ToString
    '      objElementColumn.ColumnID = objRow.Item("columnid").ToString

    '      objElementColumn.ValueType = objRow.Item("valuetype").ToString
    '      objElementColumn.Value = objRow.Item("value").ToString
    '      objElementColumn.WFFormIdentifier = objRow.Item("wfformidentifier").ToString
    '      objElementColumn.WFValueIdentifier = objRow.Item("wfvalueidentifier").ToString
    '      objElementColumn.DBColumnID = objRow.Item("dbcolumnid").ToString
    '      objElementColumn.DBRecord = objRow.Item("dbrecord").ToString
    '      objElementColumn.CalcID = objRow.Item("calcid").ToString

    '      objObjects.Add(objElementColumn)

    '    Next


    '  Catch ex As Exception
    '    Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.LoadingData, String.Empty, HRProEngine.ErrorHandler.Severity.Error, ex.Message, String.Empty)

    '  Finally

    '  End Try

    'End Function

#End Region

    Private Function NullSafe(ByVal row As System.Data.DataRow, ByVal columnName As String, ByVal defaultValue As Object) As Object

      If row.IsNull(columnName) Then
        Return defaultValue
      Else
        Return row.Item(columnName)
      End If

    End Function

  End Module

End Namespace
