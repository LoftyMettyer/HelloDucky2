Option Strict Off

Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects

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

      ' Populate functions
      Dim ds As DataSet = Globals.CommitDB.ExecStoredProcedure("spadmin_getcomponentcode", Nothing)
      For Each row As DataRow In ds.Tables(0).Rows

        Dim codeLibrary As New Things.CodeLibrary
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

    Public Function GetCodeLibraryDependancies(ByVal codeLibrary As Things.CodeLibrary) As List(Of Setting)

      Dim params As New Connectivity.Parameters
      params.Add("@componentid", CInt(codeLibrary.ID))
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

#Region "Workflow"

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
