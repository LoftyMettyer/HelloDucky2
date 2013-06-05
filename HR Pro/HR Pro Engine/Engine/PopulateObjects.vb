Option Strict Off

Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects


    Public Sub PopulateSystemSettings()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim objSetting As Things.Setting


      Try

        ' Clear existing objects
        Globals.SystemSettings.Clear()

        ' Populate module setup
        objDataset = Globals.CommitDB.ExecStoredProcedure("spadmin_getsystemsettings", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objSetting = New Things.Setting
          objSetting.Module = objRow.Item("section").ToString
          objSetting.Parameter = objRow.Item("settingkey").ToString
          objSetting.Value = objRow.Item("settingvalue").ToString

          Globals.SystemSettings.Add(objSetting)
        Next

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, String.Empty, SystemFramework.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      End Try

    End Sub

    Public Sub PopulateModuleSettings()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim objSetting As Things.Setting


      Try

        ' Clear existing objects
        Globals.ModuleSetup.Clear()

        ' Populate module setup
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmodulesetup", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objSetting = New Things.Setting
          objSetting.Module = objRow.Item("modulekey").ToString
          objSetting.Parameter = objRow.Item("parameterkey").ToString
          objSetting.SubType = objRow.Item("subtype").ToString

          If Not objRow.Item("value").ToString = "" Then
            Select Case objSetting.SubType
              Case Type.Table
                objSetting.Table = Globals.Tables.GetById(objRow.Item("value").ToString)
              Case Type.Column
                objSetting.Value = objRow.Item("value").ToString
                'If objRow.Item("tableid").ToString > 0 Then
                'objSetting.Column = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString).Objects(Things.Type.Column).GetObject(Type.Column, objRow.Item("columnid").ToString)
                'End If
              Case Else
                objSetting.Value = objRow.Item("value").ToString

            End Select

            Globals.ModuleSetup.Add(objSetting)
          End If
        Next



      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, String.Empty, SystemFramework.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      End Try

    End Sub

    Public Function PopulateCodeLibraryDependancies(ByVal objFunction As Things.CodeLibrary) As List(Of Setting)

      Dim params As New Connectivity.Parameters
      params.Add("@componentid", CInt(objFunction.ID))
      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcodedependancies", params)

      Dim dependancies As New List(Of Setting)

      For Each row As DataRow In ds.Tables(0).Rows

        Dim setting As New Things.Setting
        setting.SettingType = row.Item("type").ToString
        setting.Module = row.Item("parameterkey").ToString
        setting.Parameter = row.Item("modulekey").ToString
        setting.Value = row.Item("value").ToString
        setting.Code = row.Item("code").ToString

        dependancies.Add(setting)
      Next

      Return dependancies

    End Function

    Public Sub PopulateSystemThings()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim objCodeLibrary As Things.CodeLibrary

      Try

        ' Clear existing objects
        Globals.Operators.Clear()
        Globals.Functions.Clear()
        Globals.ModuleSetup.Clear()

        ' Populate functions
        objDataset = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcode", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objCodeLibrary = New Things.CodeLibrary
          objCodeLibrary.ID = objRow.Item("id").ToString
          objCodeLibrary.Name = objRow.Item("name").ToString
          objCodeLibrary.Code = objRow.Item("code").ToString
          objCodeLibrary.PreCode = objRow.Item("precode").ToString
          objCodeLibrary.AfterCode = objRow.Item("aftercode").ToString
          objCodeLibrary.ReturnType = objRow.Item("returntype").ToString
          objCodeLibrary.OperatorType = objRow.Item("operatortype").ToString
          objCodeLibrary.RecordIDRequired = objRow.Item("recordidrequired").ToString
          objCodeLibrary.CalculatePostAudit = objRow.Item("calculatepostaudit").ToString
          objCodeLibrary.IsUniqueCode = objRow.Item("isuniquecode").ToString
          objCodeLibrary.IsTimeDependant = objRow.Item("istimedependant").ToString
          objCodeLibrary.IsGetFieldFromDB = objRow.Item("isgetfieldfromdb").ToString
          objCodeLibrary.CaseCount = objRow.Item("casecount").ToString
          objCodeLibrary.MakeTypeSafe = objRow.Item("maketypesafe").ToString
          objCodeLibrary.OvernightOnly = objRow.Item("overnightonly").ToString
          objCodeLibrary.Tuning.Rating = objRow.Item("performancerating").ToString
          objCodeLibrary.DependsOnBankHoliday = objRow.Item("dependsonbankholiday").ToString
          objCodeLibrary.Dependancies = PopulateCodeLibraryDependancies(objCodeLibrary)

          If objRow.Item("isoperator") Then
            Globals.Operators.Add(objCodeLibrary)
          Else
            Globals.Functions.Add(objCodeLibrary)
          End If

        Next

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, String.Empty, SystemFramework.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      Finally
        objDataset = Nothing

      End Try

    End Sub

    'Public Sub PopulateThings()

    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow

    '  Dim objTable As Things.Table

    '  Dim objParameters As New Connectivity.Parameters

    '  ' Clear existing objects
    '  Globals.Tables.Clear()
    '  Globals.Workflows.Clear()
    '  'CommitDB.Open()

    '  'objParameters.Add("@type", CInt(iType))
    '  objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", objParameters)

    '  '   ProgressInfo.TotalSteps2 = objDataset.Tables(0).Rows.Count
    '  For Each objRow In objDataset.Tables(0).Rows

    '    objTable = New Things.Table
    '    objTable.ID = objRow.Item("id").ToString
    '    objTable.TableType = objRow.Item("tabletype").ToString
    '    objTable.Name = objRow.Item("name").ToString
    '    objTable.SchemaName = "dbo"
    '    objTable.IsRemoteView = objRow.Item("isremoteview")
    '    objTable.AuditInsert = objRow.Item("auditinsert").ToString
    '    objTable.AuditDelete = objRow.Item("auditdelete").ToString
    '    objTable.DefaultEmailID = objRow.Item("defaultemailid").ToString
    '    objTable.DefaultOrderID = objRow.Item("defaultorderid").ToString
    '    objTable.State = objRow.Item("state")

    '    ' needs putting back in when I figure out how to put an IF statement in the Access storedprocs (queries). Otherwise will need to split the code

    '    ' Get all child objects for this table
    '    Things.PopulateTable(objTable)
    '    objTable.Root = objTable

    '    Globals.Tables.Add(objTable)

    '  Next

    '  objDataset = Nothing
    '  objRow = Nothing

    '  '     ProgressInfo.NextStep1()

    'End Sub

    'Public Function LoadWorkflowElementDetails(ByRef objWorkflowElement As Things.WorkflowElement) As Things.Collection

    '  Dim objObjects As New Things.Collection
    '  Dim objElementColumn As Things.WorkflowElementColumn
    '  '   Dim objElementItem As Things.WorkflowElementItem

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

    '      objElementColumn = New Things.WorkflowElementColumn
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


    '  ' Load element columns



    '  ' Load element items

    'End Function

    'Public Function LoadWorkflowElements(ByRef objWorkflow As Things.Workflow) As Things.Collection

    '  Dim objObjects As New Things.Collection
    '  Dim objElement As Things.WorkflowElement
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

    '      objElement = New Things.WorkflowElement
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

    'Public Function LoadComponents(ByRef objExpression As Things.Component, ByVal Type As ScriptDB.ComponentTypes) As Things.Collections.Generic

    '  Dim objObjects As New Things.Collections.Generic
    '  Dim objComponent As Things.Component
    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow
    '  Dim objParameters As New Connectivity.Parameters

    '  objObjects.Parent = objExpression

    '  '      objObjects.root = objExpression.Root


    '  'Debug.Assert(CInt(objExpression.ID) <> 45953)

    '  Try

    '    ' Populate components

    '    '        objParameters.Add("@componenttype", CInt([Type]))

    '    '   Debug.Assert(Not objExpression.ID = 41839)
    '    '    Debug.Assert(Not objExpression.ID = 41814, False)

    '    '   Debug.Assert(objExpression.Name <> "srp")


    '    Select Case Type
    '      Case ScriptDB.ComponentTypes.Function
    '        objParameters.Add("@expressionid", objExpression.ID)
    '        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_function", objParameters)

    '      Case ScriptDB.ComponentTypes.Calculation
    '        objParameters.Add("@expressionid", objExpression.CalculationID)
    '        'objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_calculation", objParameters)
    '        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

    '        '   Case ScriptDB.ComponentTypes.Expression
    '        '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

    '      Case Else
    '        objParameters.Add("@expressionid", objExpression.ID)
    '        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)
    '    End Select

    '    For Each objRow In objDataset.Tables(0).Rows
    '      objComponent = New Things.Component
    '      objComponent.ID = objRow.Item("componentid").ToString
    '      objComponent.SubType = objRow.Item("subtype")
    '      objComponent.Name = objRow.Item("name")
    '      objComponent.ReturnType = objRow.Item("returntype")
    '      objComponent.FunctionID = objRow.Item("functionid").ToString
    '      objComponent.OperatorID = objRow.Item("operatorid").ToString
    '      objComponent.TableID = objRow.Item("tableid").ToString
    '      objComponent.ColumnID = objRow.Item("columnid").ToString
    '      objComponent.ChildRowDetails.RowSelection = objRow.Item("columnaggregiatetype").ToString
    '      objComponent.ChildRowDetails.RowNumber = objRow.Item("specificline").ToString
    '      objComponent.ChildRowDetails.FilterID = objRow.Item("columnfilterid").ToString
    '      objComponent.ChildRowDetails.OrderID = objRow.Item("columnorderid").ToString
    '      objComponent.IsColumnByReference = objRow.Item("iscolumnbyreference").ToString
    '      objComponent.CalculationID = objRow.Item("calculationid").ToString
    '      objComponent.ValueType = objRow.Item("valuetype").ToString

    '      Select Case objComponent.ValueType
    '        Case ScriptDB.ComponentValueTypes.Date
    '          objComponent.ValueDate = objRow.Item("valuedate")
    '        Case ScriptDB.ComponentValueTypes.Logic
    '          objComponent.ValueLogic = objRow.Item("valuelogic").ToString
    '        Case ScriptDB.ComponentValueTypes.Numeric
    '          objComponent.ValueNumeric = objRow.Item("valuenumeric").ToString
    '        Case ScriptDB.ComponentValueTypes.String
    '          objComponent.ValueString = objRow.Item("valuestring").ToString
    '      End Select

    '      objComponent.LookupTableID = objRow.Item("lookuptableid").ToString
    '      objComponent.LookupColumnID = objRow.Item("lookupcolumnid").ToString

    '      objComponent.Root = objExpression.Root
    '      objComponent.BaseExpression = objExpression.BaseExpression

    '      Select Case objComponent.SubType

    '        Case ScriptDB.ComponentTypes.Function
    '          objComponent.Objects = Things.LoadComponents(objComponent, ScriptDB.ComponentTypes.Function)
    '        Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
    '          objComponent.Objects = Things.LoadComponents(objComponent, objComponent.SubType)

    '      End Select

    '      objObjects.Add(objComponent)
    '    Next


    '  Catch ex As Exception
    '    Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, String.Empty, SystemFramework.ErrorHandler.Severity.Error, ex.Message, String.Empty)

    '  Finally
    '    objDataset = Nothing
    '    objRow = Nothing

    '  End Try

    '  LoadComponents = objObjects


    'End Function

    'Public Sub PopulateTable(ByVal table As Things.Table)

    '  Dim objDataset As DataSet
    '  Dim objRow As DataRow
    '  Dim objParameters As New Connectivity.Parameters

    '  Dim objColumn As Things.Column
    '  Dim objRelation As Things.Relation
    '  Dim objExpression As Things.Expression
    '  Dim objView As Things.View
    '  Dim objValidation As Things.Validation
    '  Dim objTableOrder As Things.TableOrder
    '  Dim objDescription As Things.RecordDescription
    '  Dim objMask As Things.Mask

    '  'table.Objects.Parent = table
    '  'table.Objects.Root = table.Root

    '  Try

    '    ' Populate relations
    '    objParameters.Add("@tableid", table.ID)
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getrelations", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows
    '      objRelation = New Things.Relation
    '      objRelation.RelationshipType = objRow.Item("relationship").ToString

    '      'Select Case objRelation.RelationshipType
    '      '  Case ScriptDB.RelationshipType.Parent
    '      '    objRelation.Parent = Table.Objects.Parent

    '      '  Case ScriptDB.RelationshipType.Child
    '      '    objRelation.Parent = Table

    '      'End Select

    '      objRelation.Parent = table
    '      objRelation.ParentID = objRow.Item("parentid").ToString
    '      objRelation.ChildID = objRow.Item("childid").ToString
    '      objRelation.Name = objRow.Item("name").ToString
    '      table.Relations.Add(objRelation)
    '    Next

    '    ' Populate columns
    '    objParameters.Clear()
    '    objParameters.Add("@parentid", table.ID)
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objColumn = New Things.Column
    '      objColumn.ID = objRow.Item("id").ToString
    '      objColumn.Parent = table
    '      objColumn.Name = objRow.Item("name").ToString
    '      objColumn.SchemaName = "dbo"
    '      objColumn.Description = objRow.Item("description").ToString
    '      objColumn.Table = table
    '      objColumn.State = objRow.Item("state")

    '      objColumn.DefaultCalcID = NullSafe(objRow, "defaultcalcid", 0).ToString
    '      objColumn.DefaultValue = objRow.Item("defaultvalue").ToString
    '      objColumn.CalcID = objRow.Item("calcid").ToString
    '      objColumn.DataType = objRow.Item("datatype")
    '      objColumn.Size = objRow.Item("size")
    '      objColumn.Decimals = objRow.Item("decimals")
    '      objColumn.Audit = objRow.Item("audit")
    '      objColumn.Mandatory = objRow.Item("mandatory")
    '      objColumn.Multiline = objRow.Item("multiline")
    '      objColumn.IsReadOnly = objRow.Item("isreadonly")
    '      objColumn.CaseType = objRow.Item("case").ToString
    '      objColumn.CalculateIfEmpty = objRow.Item("calculateifempty")
    '      objColumn.TrimType = NullSafe(objRow, "trimming", 0).ToString
    '      objColumn.Alignment = objRow.Item("alignment").ToString
    '      objColumn.UniqueType = objRow.Item("uniquechecktype").ToString

    '      table.Columns.Add(objColumn)
    '    Next

    '    ' Orders
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorders", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows
    '      objTableOrder = New Things.TableOrder
    '      objTableOrder.ID = objRow.Item("orderid").ToString
    '      objTableOrder.Parent = table
    '      objTableOrder.Name = objRow.Item("name").ToString
    '      objTableOrder.SubType = objRow.Item("type").ToString
    '      Things.PopulateOrderItems(objTableOrder)
    '      table.TableOrders.Add(objTableOrder)
    '    Next
    '    objDataset.Dispose()
    '    objDataset = Nothing

    '    ' Populate expressions
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objExpression = New Things.Expression
    '      objExpression.ID = objRow.Item("id").ToString
    '      objExpression.Parent = table
    '      objExpression.Name = objRow.Item("name").ToString
    '      objExpression.ExpressionType = objRow.Item("type").ToString
    '      objExpression.SchemaName = "dbo"
    '      objExpression.Description = objRow.Item("description").ToString
    '      objExpression.State = objRow.Item("state")
    '      objExpression.ReturnType = objRow.Item("returntype")
    '      objExpression.Size = objRow.Item("size")
    '      objExpression.Decimals = objRow.Item("decimals")
    '      objExpression.BaseTable = table
    '      objExpression.BaseExpression = objExpression

    '      'Get all child objects for this expression
    '      objExpression.Objects = Things.LoadComponents(objExpression, ScriptDB.ComponentTypes.Expression)
    '      table.Expressions.Add(objExpression)
    '    Next

    '    ' Views
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviews", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows
    '      objView = New Things.View
    '      objView.ID = objRow.Item("id").ToString
    '      objView.Parent = table
    '      objView.Name = objRow.Item("name").ToString
    '      objView.Description = objRow.Item("description").ToString
    '      objView.Filter = table.Expressions.GetById(objRow.Item("filterid").ToString)
    '      Things.PopulateViewItems(objView)
    '      table.Views.Add(objView)
    '    Next

    '    ' Validations
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getvalidations", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows
    '      objValidation = New Things.Validation
    '      objValidation.ValidationType = objRow.Item("validationtype").ToString
    '      objValidation.Column = table.Columns.GetById(objRow.Item("columnid").ToString)
    '      table.Validations.Add(objValidation)
    '    Next

    '    ' Record Description
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getdescriptions", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows

    '      objDescription = New Things.RecordDescription
    '      objDescription.ID = objRow.Item("id").ToString
    '      objDescription.Parent = table
    '      objDescription.Name = objRow.Item("name").ToString
    '      objDescription.SchemaName = "dbo"
    '      objDescription.Description = objRow.Item("description").ToString
    '      objDescription.State = objRow.Item("state")
    '      objDescription.ReturnType = objRow.Item("returntype")
    '      objDescription.Size = objRow.Item("size")
    '      objDescription.Decimals = objRow.Item("decimals")
    '      objDescription.BaseTable = table
    '      objDescription.BaseExpression = objDescription

    '      'Get all child objects for this expression
    '      objDescription.Objects = Things.LoadComponents(objDescription, ScriptDB.ComponentTypes.Expression)

    '      table.RecordDescription = objDescription
    '    Next

    '    ' Masks
    '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getmasks", objParameters)
    '    For Each objRow In objDataset.Tables(0).Rows
    '      objMask = New Things.Mask
    '      objMask.ID = objRow.Item("id").ToString
    '      objMask.Parent = table
    '      objMask.Name = objRow.Item("name").ToString
    '      objMask.AssociatedColumn = table.Columns.GetById(objRow.Item("columnid").ToString)
    '      objMask.SchemaName = "dbo"
    '      objMask.Description = objRow.Item("description").ToString
    '      objMask.State = objRow.Item("state")
    '      objMask.ReturnType = objRow.Item("returntype")
    '      objMask.Size = objRow.Item("size")
    '      objMask.Decimals = objRow.Item("decimals")
    '      objMask.BaseTable = table
    '      objMask.BaseExpression = objMask

    '      'Get all child objects for this expression
    '      objMask.Objects = Things.LoadComponents(objMask, ScriptDB.ComponentTypes.Expression)

    '      table.Masks.Add(objMask)
    '    Next



    '  Catch ex As Exception

    '  Finally
    '    objDataset = Nothing
    '    objRow = Nothing

    '  End Try

    'End Sub

    Public Sub PopulateViewItems(ByRef objView As Things.View)

      Dim objColumn As Things.Column

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Try
        objParameters.Add("@viewid", objView.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviewitems", objParameters)

        For Each objRow In objDataset.Tables(0).Rows
          objColumn = CType(objView.Parent, Table).Columns.GetById(objRow.Item("columnid").ToString)
          If Not objColumn Is Nothing Then
            objView.Columns.Add(objColumn)
          End If
        Next

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing
      End Try

    End Sub

    Public Sub PopulateOrderItems(ByRef objOrder As Things.TableOrder)

      Dim objObjects As New Things.Collections.Generic
      Dim objOrderItem As Things.TableOrderItem
      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      'TODO check
      objObjects.Parent = objOrder

      Try

        objParameters.Add("@tableid", objOrder.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorderitems", objParameters)

        For Each objRow In objDataset.Tables(0).Rows
          objOrderItem = New Things.TableOrderItem
          objOrderItem.ID = objRow.Item("orderid").ToString
          objOrderItem.ColumnType = objRow.Item("type")
          objOrderItem.Sequence = objRow.Item("sequence")
          objOrderItem.Ascending = objRow.Item("ascending")
          objOrderItem.Column = CType(objOrder.Parent, Table).Columns.GetById(objRow.Item("columnid").ToString)

          objOrder.TableOrderItems.Add(objOrderItem)
        Next


      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try


    End Sub

    Public Function PopulateUtilities(ByVal Type As Things.Type) As Boolean

      Dim bOK As Boolean = True

      Try
        Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spweb_getglobals", New Connectivity.Parameters)

        For Each row As DataRow In ds.Tables(0).Rows

          Dim modify As New Things.GlobalModify
          modify.Name = row("name").ToString
          modify.Description = row("description").ToString
          PopulateDataModifyItems(modify)

          'TODO: ADD BACK IN
          ' Globals.Things.Add(objGlobalModify)
        Next

      Catch ex As Exception
        bOK = False
      End Try

      Return bOK

    End Function

    Public Sub PopulateDataModifyItems(ByRef objDataModify As Things.GlobalModify)

      Dim params As New Connectivity.Parameters
      params.Add("@id", objDataModify.ID)
      Dim ds As DataSet = Globals.MetadataDB.ExecStoredProcedure("spweb_getglobal", params)

      For Each row As DataRow In ds.Tables(1).Rows

        Dim objItem As New Things.GlobalModifyItem
        objItem.CalculationID = row.Item("calcid").ToString
        objItem.Value = row.Item("value").ToString

        objDataModify.GlobalModifyItems.Add(objItem)
      Next

    End Sub

    Private Function NullSafe(ByRef ObjectData As System.Data.DataRow, ByVal ColumnName As String, ByVal DefaultValue As Object) As Object

      If ObjectData.IsNull(ColumnName) Then
        Return DefaultValue
      Else
        Return ObjectData.Item(ColumnName)
      End If

    End Function

  End Module

End Namespace
