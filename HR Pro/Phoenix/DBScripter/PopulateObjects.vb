Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects

    Public Sub PopulateModuleSettings(ByRef ProgressInfo As HCMProgressBar)

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim objSetting As Things.Setting

      Try

        ' Clear existing objects
        Globals.ModuleSetup.Clear()

        objDataset = CommitDB.ExecStoredProcedure("spadmin_getmodulesettings", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objSetting = New Things.Setting
          objSetting.SubType = objRow.Item("type").ToString
          objSetting.Module = objRow.Item("module").ToString
          objSetting.Parameter = objRow.Item("parameter").ToString
          objSetting.Value = objRow.Item("value").ToString
          Select Case objSetting.SubType
            Case Type.Table
              objSetting.Table = Globals.Things.GetObject(Type.Table, objRow.Item("value").ToString)
            Case Type.Column
              If objRow.Item("tableid").ToString > 0 Then
                objSetting.Column = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString).Objects(Things.Type.Column).GetObject(Type.Column, objRow.Item("columnid").ToString)
              End If
          End Select

          Globals.ModuleSetup.Add(objSetting)

        Next

      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      End Try

    End Sub

    Public Function PopulateCodeLibraryDependancies(ByRef objFunction As Things.CodeLibrary) As Things.Collection

      Dim objDependancies As New Things.Collection
      Dim objSetting As Things.Setting
      Dim objParameters As New Connectivity.Parameters
      Dim objDataset As DataSet
      Dim objRow As DataRow

      Try
        objDependancies.Clear()
        objParameters.Clear()
        objParameters.Add("@componentid", CInt(objFunction.ID))
        objDataset = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcodedependancies", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objSetting = New Things.Setting
          objSetting.Module = objRow.Item("parameterkey").ToString
          objSetting.Parameter = objRow.Item("modulekey").ToString
          objSetting.Value = objRow.Item("value").ToString
          objSetting.SubType = objRow.Item("settingtype").ToString
          objDependancies.Add(objSetting)
        Next

      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      End Try

      Return objDependancies

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
          objCodeLibrary.SplitIntoCase = objRow.Item("splitintocase")
          objCodeLibrary.AppendWildcard = objRow.Item("appendwildcard")
          objCodeLibrary.AfterCode = objRow.Item("aftercode").ToString
          objCodeLibrary.ReturnType = objRow.Item("returntype").ToString
          objCodeLibrary.OperatorType = objRow.Item("operatortype").ToString
          objCodeLibrary.BypassValidation = objRow.Item("bypassvalidation").ToString
          objCodeLibrary.Dependancies = PopulateCodeLibraryDependancies(objCodeLibrary)

          If objRow.Item("isoperator") Then
            Globals.Operators.Add(objCodeLibrary)
          Else
            Globals.Functions.Add(objCodeLibrary)
          End If

        Next

      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, vbNullString)

      End Try

    End Sub

    Public Sub PopulateThings() 'ByRef ProgressInfo As HCMProgressBar)

      Dim objDataset As DataSet
      Dim objRow As DataRow

      Dim objTable As Things.Table
      Dim objDescription As Things.RecordDescription

      Dim objParameters As New Connectivity.Parameters

      ' Clear existing objects
      Globals.Things.Clear()
      Globals.Workflows.Clear()
      'CommitDB.Open()

      'objParameters.Add("@type", CInt(iType))
      objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", objParameters)

      '   ProgressInfo.TotalSteps2 = objDataset.Tables(0).Rows.Count
      For Each objRow In objDataset.Tables(0).Rows

        objTable = New Things.Table
        objTable.ID = objRow.Item("id").ToString
        objTable.TableType = objRow.Item("tabletype").ToString
        objTable.Name = objRow.Item("name").ToString
        objTable.SchemaName = "dbo"
        objTable.IsRemoteView = objRow.Item("isremoteview")
        objTable.AuditInsert = objRow.Item("auditinsert").ToString
        objTable.AuditDelete = objRow.Item("auditdelete").ToString
        objTable.DefaultEmailID = objRow.Item("defaultemailid").ToString
        objTable.DefaultOrderID = objRow.Item("defaultorderid").ToString
        objTable.State = objRow.Item("state")

        ' needs putting back in when I figure out how to put an IF statement in the Access storedprocs (queries). Otherwise will need to split the code

        '' Record Description
        'objDescription = New Things.RecordDescription
        'objDescription.ID = objRow.Item("recorddescriptionid").ToString
        'objDescription.ExpressionType = ScriptDB.ExpressionType.RecordDescription
        'objDescription.BaseTable = objTable
        'objDescription.Objects = Things.LoadComponents(objDescription, ScriptDB.ComponentTypes.Expression)
        'objTable.RecordDescription = objDescription

        ' Get all child objects for this table
        Things.PopulateTable(objTable)
        objTable.Root = objTable

        Globals.Things.Add(objTable)
        ''      ProgressInfo.NextStep2()

      Next

      objDataset = Nothing
      objRow = Nothing

      '     ProgressInfo.NextStep1()

    End Sub

    Public Function LoadWorkflowElementDetails(ByRef objWorkflowElement As Things.WorkflowElement) As Things.Collection

      Dim objObjects As New Things.Collection
      Dim objElementColumn As Things.WorkflowElementColumn
      '   Dim objElementItem As Things.WorkflowElementItem

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      objObjects.Parent = objWorkflowElement
      objObjects.Root = objWorkflowElement.Root

      Try

        ' Populate element
        objParameters.Add("@elementid", objWorkflowElement.ID)
        objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelementcolumns", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objElementColumn = New Things.WorkflowElementColumn
          objElementColumn.ID = objRow.Item("elementid").ToString
          objElementColumn.ColumnID = objRow.Item("columnid").ToString

          objElementColumn.ValueType = objRow.Item("valuetype").ToString
          objElementColumn.Value = objRow.Item("value").ToString
          objElementColumn.WFFormIdentifier = objRow.Item("wfformidentifier").ToString
          objElementColumn.WFValueIdentifier = objRow.Item("wfvalueidentifier").ToString
          objElementColumn.DBColumnID = objRow.Item("dbcolumnid").ToString
          objElementColumn.DBRecord = objRow.Item("dbrecord").ToString
          objElementColumn.CalcID = objRow.Item("calcid").ToString

          objObjects.Add(objElementColumn)

        Next


      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, String.Empty)

      Finally

      End Try


      ' Load element columns



      ' Load element items

    End Function

    Public Function LoadWorkflowElements(ByRef objWorkflow As Things.Workflow) As Things.Collection

      Dim objObjects As New Things.Collection
      Dim objElement As Things.WorkflowElement
      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      objObjects.Parent = objWorkflow
      objObjects.Root = objWorkflow.Root

      Try

        ' Populate element
        objParameters.Add("@workflowid", objWorkflow.ID)
        objDataset = CommitDB.ExecStoredProcedure("spadmin_getworkflowelements", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objElement = New Things.WorkflowElement
          objElement.ID = objRow.Item("elementid").ToString
          objElement.SubType = objRow.Item("type")
          objElement.Caption = objRow.Item("caption")
          objElement.ConnectionPairID = objRow.Item("ConnectionPairID")
          objElement.LeftCoord = objRow.Item("LeftCoord")
          objElement.TopCoord = objRow.Item("TopCoord")
          objElement.DecisionCaptionType = objRow.Item("DecisionCaptionType").ToString
          objElement.Identifier = objRow.Item("Identifier").ToString
          objElement.TrueFlowIdentifier = objRow.Item("TrueFlowIdentifier").ToString
          objElement.DataAction = objRow.Item("DataAction").ToString
          objElement.DataTableID = objRow.Item("DataTableID").ToString
          objElement.DataRecord = objRow.Item("DataRecord").ToString
          objElement.EmailID = objRow.Item("EmailID").ToString
          objElement.EmailRecord = objRow.Item("EmailRecord").ToString
          objElement.WebFormBGColor = objRow.Item("WebFormBGColor").ToString
          objElement.WebFormBGImageID = objRow.Item("WebFormBGImageID").ToString
          objElement.WebFormBGImageLocation = objRow.Item("WebFormBGImageLocation").ToString
          objElement.WebFormDefaultFontName = objRow.Item("WebFormDefaultFontName").ToString
          objElement.WebFormDefaultFontSize = objRow.Item("WebFormDefaultFontSize").ToString
          objElement.WebFormDefaultFontBold = objRow.Item("WebFormDefaultFontBold").ToString
          objElement.WebFormDefaultFontItalic = objRow.Item("WebFormDefaultFontItalic").ToString
          objElement.WebFormDefaultFontStrikeThru = objRow.Item("WebFormDefaultFontStrikeThru").ToString
          objElement.WebFormDefaultFontUnderline = objRow.Item("WebFormDefaultFontUnderline").ToString
          objElement.WebFormWidth = objRow.Item("WebFormWidth").ToString
          objElement.RecSelWebFormIdentifier = objRow.Item("RecSelWebFormIdentifier").ToString
          objElement.RecSelIdentifier = objRow.Item("RecSelIdentifier").ToString
          objElement.SecondaryDataRecord = objRow.Item("SecondaryDataRecord").ToString
          objElement.SecondaryRecSelWebFormIdentifier = objRow.Item("SecondaryRecSelWebFormIdentifier").ToString
          objElement.SecondaryRecSelIdentifier = objRow.Item("SecondaryRecSelIdentifier").ToString
          objElement.EmailSubject = objRow.Item("EmailSubject").ToString
          objElement.TimeoutFrequency = objRow.Item("TimeoutFrequency").ToString
          objElement.TimeoutPeriod = objRow.Item("TimeoutPeriod").ToString
          objElement.DataRecordTable = objRow.Item("DataRecordTable").ToString
          objElement.SecondaryDataRecordTable = objRow.Item("SecondaryDataRecordTable").ToString
          objElement.TrueFlowType = objRow.Item("TrueFlowType").ToString
          objElement.TrueFlowExprID = objRow.Item("TrueFlowExprID").ToString
          objElement.DescriptionExprID = objRow.Item("DescriptionExprID").ToString
          objElement.WebFormFGColor = objRow.Item("WebFormFGColor").ToString
          objElement.DescHasWorkflowName = objRow.Item("DescHasWorkflowName").ToString
          objElement.DescHasElementCaption = objRow.Item("DescHasElementCaption").ToString
          objElement.EmailCCID = objRow.Item("EmailCCID").ToString
          objElement.TimeoutExcludeWeekend = objRow.Item("TimeoutExcludeWeekend").ToString
          objElement.CompletionMessageType = objRow.Item("CompletionMessageType").ToString
          objElement.CompletionMessage = objRow.Item("CompletionMessage").ToString
          objElement.SavedForLaterMessageType = objRow.Item("SavedForLaterMessageType").ToString
          objElement.SavedForLaterMessage = objRow.Item("SavedForLaterMessage").ToString
          objElement.FollowOnFormsMessageType = objRow.Item("FollowOnFormsMessageType").ToString
          objElement.FollowOnFormsMessage = objRow.Item("FollowOnFormsMessage").ToString
          objElement.Attachment_Type = objRow.Item("Attachment_Type").ToString
          objElement.Attachment_File = objRow.Item("Attachment_File").ToString
          objElement.Attachment_WFElementIdentifier = objRow.Item("Attachment_WFElementIdentifier").ToString
          objElement.Attachment_WFValueIdentifier = objRow.Item("Attachment_WFValueIdentifier").ToString
          objElement.Attachment_DBColumnID = objRow.Item("Attachment_DBColumnID").ToString
          objElement.Attachment_DBRecord = objRow.Item("Attachment_DBRecord").ToString
          objElement.Attachment_DBElement = objRow.Item("Attachment_DBElement").ToString
          objElement.Attachment_DBValue = objRow.Item("Attachment_DBValue").ToString

          objElement.Objects = Things.LoadWorkflowElementDetails(objElement)

          objObjects.Add(objElement)

        Next


      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, String.Empty)

      End Try

      LoadWorkflowElements = objObjects


    End Function

    Public Function LoadComponents(ByRef objExpression As Things.Component, ByVal Type As ScriptDB.ComponentTypes) As Things.Collection

      Dim objObjects As New Things.Collection
      Dim objComponent As Things.Component
      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      objObjects.Parent = objExpression

      '      objObjects.root = objExpression.Root

      Try

        ' Populate components

        '        objParameters.Add("@componenttype", CInt([Type]))

        '   Debug.Assert(Not objExpression.ID = 41839)
        '    Debug.Assert(Not objExpression.ID = 41814, False)


        Select Case Type
          Case ScriptDB.ComponentTypes.Function
            objParameters.Add("@expressionid", objExpression.ID)
            objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_function", objParameters)

          Case ScriptDB.ComponentTypes.Calculation
            objParameters.Add("@expressionid", objExpression.CalculationID)
            'objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_calculation", objParameters)
            objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

            '   Case ScriptDB.ComponentTypes.Expression
            '    objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)

          Case Else
            objParameters.Add("@expressionid", objExpression.ID)
            objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcomponent_base", objParameters)
        End Select

        For Each objRow In objDataset.Tables(0).Rows
          objComponent = New Things.Component
          objComponent.ID = objRow.Item("componentid").ToString
          objComponent.SubType = objRow.Item("subtype")
          objComponent.Name = objRow.Item("name")
          objComponent.ReturnType = objRow.Item("returntype")
          objComponent.FunctionID = objRow.Item("functionid").ToString
          objComponent.OperatorID = objRow.Item("operatorid").ToString
          objComponent.TableID = objRow.Item("tableid").ToString
          objComponent.ColumnID = objRow.Item("columnid").ToString
          objComponent.ColumnAggregiateType = objRow.Item("columnaggregiatetype").ToString
          objComponent.IsColumnByReference = objRow.Item("iscolumnbyreference").ToString
          objComponent.ColumnFilterID = objRow.Item("columnfilterid").ToString
          objComponent.ColumnOrderID = objRow.Item("columnorderid").ToString
          objComponent.CalculationID = objRow.Item("calculationid").ToString
          objComponent.ValueType = objRow.Item("valuetype").ToString

          Select Case objComponent.ValueType
            Case ScriptDB.ComponentValueTypes.Date
              objComponent.ValueDate = objRow.Item("valuedate")
            Case ScriptDB.ComponentValueTypes.Logic
              objComponent.ValueLogic = objRow.Item("valuelogic").ToString
            Case ScriptDB.ComponentValueTypes.Numeric
              objComponent.ValueNumeric = objRow.Item("valuenumeric").ToString
            Case ScriptDB.ComponentValueTypes.String
              objComponent.ValueString = objRow.Item("valuestring").ToString
          End Select

          objComponent.LookupTableID = objRow.Item("lookuptableid").ToString
          objComponent.LookupColumnID = objRow.Item("lookupcolumnid").ToString

          objComponent.Root = objExpression.Root
          objComponent.BaseExpression = objExpression.BaseExpression

          Select Case objComponent.SubType

            Case ScriptDB.ComponentTypes.Function
              objComponent.Objects = Things.LoadComponents(objComponent, ScriptDB.ComponentTypes.Function)
            Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
              objComponent.Objects = Things.LoadComponents(objComponent, objComponent.SubType)

          End Select


          ' Calculations are attached after all the expressions have been loaded - otherwise we have horrible recursion problems


          ' If objComponent.SubType = ScriptDB.ComponentTypes.Calculation Then
          '            objComponent.Objects = Things.LoadComponents(objComponent, ScriptDB.ComponentTypes.Calculation)

          '            objComponent.Objects = Globals.Things.GetObject(Enums.Type.Expression, objComponent.CalculationID)
          ' objComponent.CalculationID = 

          '' Isn't already loaded - do it now
          'If objComponent.Objects Is Nothing Then

          '  'Globals.Things(

          '  ' the base table of this expression
          '  'objBaseTable = Globals.Things.GetObject(Enums.Type.Table, (CType(objExpression, Things.Expression).BaseTableID))

          '  objBaseTable = CType(objExpression, Things.Expression).BaseTable

          '  objComponent.Objects = Things.LoadComponents(objComponent, ScriptDB.ComponentTypes.Calculation)

          '  objCalculation = New Things.Expression
          '  objCalculation.ID = objRow.Item("id").ToString
          '  objExpression.Name = objRow.Item("name").ToString
          '  objExpression.Description = objRow.Item("description").ToString
          '  objExpression.Parent = objObjects.Parent
          '  objExpression.DataType = objRow.Item("datatype")
          '  objExpression.Size = objRow.Item("size")
          '  objExpression.Decimals = objRow.Item("decimals")
          '  objExpression.Root = objExpression

          '  ' Get all child objects for this table
          '  objExpression.Objects = Things.LoadComponents(objExpression, ScriptDB.ComponentTypes.Expression)
          '  objObjects.Add(objExpression)

          'End If

          'End If

          objObjects.Add(objComponent)
        Next


      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, String.Empty, DBScripter.ErrorHandler.Severity.Error, ex.Message, String.Empty)

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try

      LoadComponents = objObjects


    End Function

    ' TODO? - Get all the child objects in one hit - means we only hit the db once for every table.
    Public Sub PopulateTable(ByRef Table As Things.Base)

      ', ByVal iObjectType As ObjectType

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      '    Dim objObjects As New Things.Collection
      Dim objColumn As Things.Column
      Dim objRelation As Things.Relation
      Dim objExpression As Things.Expression
      Dim objView As Things.View
      '    Dim objWorkflow As Things.Workflow
      '     Dim objScreen As Things.Screen
      '      Dim objDiaryLink As Things.DiaryLink
      Dim objTableOrder As Things.TableOrder

      Table.Objects.Parent = Table
      Table.Objects.Root = Table.Root

      Try

        ' Populate relations
        objParameters.Add("@tableid", Table.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getrelations", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objRelation = New Things.Relation
          objRelation.Parent = Table
          objRelation.RelationshipType = objRow.Item("relationship").ToString
          objRelation.ParentID = objRow.Item("parentid").ToString
          objRelation.ChildID = objRow.Item("childid").ToString
          objRelation.Name = objRow.Item("name").ToString
          Table.Objects.Add(objRelation)
        Next

        ' Populate columns
        objParameters.Clear()
        objParameters.Add("@parentid", Table.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objColumn = New Things.Column
          objColumn.ID = objRow.Item("id").ToString
          objColumn.Parent = Table
          objColumn.Name = objRow.Item("name").ToString
          objColumn.SchemaName = "dbo"
          objColumn.Description = objRow.Item("description").ToString
          objColumn.Table = Table
          objColumn.State = objRow.Item("state")

          objColumn.CalcID = objRow.Item("calcid").ToString
          objColumn.DataType = objRow.Item("datatype")
          objColumn.Size = objRow.Item("size")
          objColumn.Decimals = objRow.Item("decimals")
          objColumn.Audit = objRow.Item("audit")
          objColumn.Mandatory = objRow.Item("mandatory")
          objColumn.Multiline = objRow.Item("multiline")
          objColumn.IsReadOnly = objRow.Item("isreadonly")
          objColumn.CaseType = objRow.Item("case").ToString
          objColumn.CalculateIfEmpty = objRow.Item("calculateifempty")
          objColumn.TrimType = objRow.Item("trimming").ToString
          objColumn.Alignment = objRow.Item("alignment").ToString

          Table.Objects.Add(objColumn)
        Next


        ' Orders
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getorders", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objTableOrder = New Things.TableOrder
          objTableOrder.ID = objRow.Item("orderid").ToString
          objTableOrder.Parent = Table
          objTableOrder.Name = objRow.Item("name").ToString
          objTableOrder.SubType = objRow.Item("type").ToString
          Things.PopulateOrderItems(objTableOrder)
          Table.Objects.Add(objTableOrder)
        Next
        objDataset.Dispose()
        objDataset = Nothing

        ' Populate expressions
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getexpressions", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objExpression = New Things.Expression
          objExpression.ID = objRow.Item("id").ToString
          objExpression.Parent = Table
          objExpression.Name = objRow.Item("name").ToString
          objExpression.SchemaName = "dbo"
          objExpression.Description = objRow.Item("description").ToString
          objExpression.State = objRow.Item("state")
          objExpression.ReturnType = objRow.Item("returntype")
          objExpression.Size = objRow.Item("size")
          objExpression.Decimals = objRow.Item("decimals")
          objExpression.BaseTable = Table
          objExpression.BaseExpression = objExpression

          'Get all child objects for this expression
          objExpression.Objects = Things.LoadComponents(objExpression, ScriptDB.ComponentTypes.Expression)
          Table.Objects.Add(objExpression)
        Next

        ' Views
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviews", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objView = New Things.View
          objView.ID = objRow.Item("id").ToString
          objView.Parent = Table
          objView.Name = objRow.Item("name").ToString
          objView.Description = objRow.Item("description").ToString
          objView.Filter = Table.GetObject(Things.Type.Expression, objRow.Item("filterid").ToString)
          Things.PopulateViewItems(objView)
          Table.Objects.Add(objView)
        Next


      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try

    End Sub

    Public Sub PopulateViewItems(ByRef objView As Things.View)

      Dim objColumn As Things.Column

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Try
        objParameters.Add("@viewid", objView.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getviewitems", objParameters)

        For Each objRow In objDataset.Tables(0).Rows
          objColumn = objView.Parent.GetObject(Things.Type.Column, objRow.Item("columnid").ToString)
          If Not objColumn Is Nothing Then
            objView.Objects.Add(objColumn)
          End If
        Next

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing
      End Try

    End Sub


    Public Sub PopulateOrderItems(ByRef objOrder As Things.TableOrder)

      Dim objObjects As New Things.Collection
      Dim objOrderItem As Things.TableOrderItem
      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

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
          objOrderItem.Column = objOrder.Parent.GetObject(Things.Type.Column, objRow.Item("columnid").ToString)

          objOrder.Objects.Add(objOrderItem)
        Next


      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try


    End Sub

    Public Function PopulateUtilities(ByRef Type As Things.Type) As Boolean

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim bOK As Boolean = True

      Dim objGlobalModify As Things.GlobalModify

      Try

        ' Populate relations
        ' objParameters.Add("@type", CInt(Type))
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spweb_getglobals", objParameters)
        For Each objRow In objDataset.Tables(0).Rows
          objGlobalModify = New Things.GlobalModify
          objGlobalModify.Name = objRow("name").ToString
          objGlobalModify.Description = objRow("description").ToString
          '          objGlobalModify.u = objRow("subtype")
          PopulateDataModifyItems(objGlobalModify)


          Globals.Things.Add(objGlobalModify)
        Next

      Catch ex As Exception
        bOK = False
      End Try

      Return bOK

    End Function

    Public Sub PopulateDataModifyItems(ByRef objDataModify As Things.GlobalModify)

      Dim objObjects As New Things.Collection
      Dim objItem As Things.GlobalModifyItem
      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      objObjects.Parent = objDataModify

      Try

        objParameters.Add("@id", objDataModify.ID)
        objDataset = Globals.MetadataDB.ExecStoredProcedure("spweb_getglobal", objParameters)
        For Each objRow In objDataset.Tables(1).Rows
          objItem = New Things.GlobalModifyItem
          objItem.CalculationID = objRow.Item("calcid").ToString
          objItem.Value = objRow.Item("value").ToString
          objDataModify.Objects.Add(objItem)
        Next

      Catch ex As Exception

      End Try

    End Sub


  End Module

End Namespace
