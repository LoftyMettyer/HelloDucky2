Namespace Things

  <HideModuleName()> _
  Public Module PopulateObjects2

    Public Sub PopulateTables()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters
      Dim objTable As Things.Table

      Try

        Globals.Things.Clear()
        Globals.Workflows.Clear()

        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_gettables", objParameters)
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
          objTable.Root = objTable

          Globals.Things.Add(objTable)
        Next

        ' Objects with no table attached
        objTable = New Things.Table
        objTable.ID = 0
        objTable.Name = "System Objects"
        Globals.Things.Add(objTable)

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing


      End Try





    End Sub

    Public Sub PopulateExpressions()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Dim objExpression As Things.Expression
      Dim objTable As Things.Table

      Try
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

      Catch ex As Exception

      End Try

    End Sub

    Public Sub PopulateColumns()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Dim objColumn As Things.Column
      Dim objTable As Things.Table

      Try

        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getcolumns", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objColumn = New Things.Column
          objColumn.ID = objRow.Item("id").ToString
          objColumn.Name = objRow.Item("name").ToString
          objColumn.SchemaName = "dbo"
          objColumn.Description = objRow.Item("description").ToString
          objColumn.State = objRow.Item("state")

          objColumn.DefaultCalcID = NullSafe(objRow, "defaultcalcid", 0).ToString
          objColumn.DefaultValue = objRow.Item("defaultvalue").ToString
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
          objColumn.TrimType = NullSafe(objRow, "trimming", 0).ToString
          objColumn.Alignment = objRow.Item("alignment").ToString

          ' Attach to table
          objTable = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString)
          objColumn.Table = objTable
          objColumn.Parent = objTable

          objTable.Objects.Add(objColumn)
        Next

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try

    End Sub

    Public Sub PopulateScreens()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Dim objScreen As Things.Screen
      Dim objTable As Things.Table

      Try

        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getscreens", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objScreen = New Things.Screen
          objScreen.ID = objRow.Item("id").ToString
          objScreen.Name = objRow.Item("name").ToString
          objScreen.SchemaName = "dbo"
          objScreen.Description = objRow.Item("description").ToString
          objScreen.State = objRow.Item("state")

          ' Attach to table
          objTable = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString)
          objScreen.Table = objTable
          objScreen.Parent = objTable

          objTable.Objects.Add(objScreen)
        Next

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try

    End Sub

    Public Sub PopulateWorkflows()

      Dim objDataset As DataSet
      Dim objRow As DataRow
      Dim objParameters As New Connectivity.Parameters

      Dim objWorkflow As Things.Workflow
      Dim objTable As Things.Table

      Try

        objDataset = Globals.MetadataDB.ExecStoredProcedure("spadmin_getworkflows", objParameters)
        For Each objRow In objDataset.Tables(0).Rows

          objWorkflow = New Things.Workflow
          objWorkflow.ID = objRow.Item("id").ToString
          objWorkflow.Name = objRow.Item("name").ToString
          objWorkflow.SchemaName = "dbo"
          objWorkflow.Description = objRow.Item("description").ToString
          objWorkflow.InitiationType = objRow.Item("initiationType").ToString
          objWorkflow.Enabled = objRow.Item("enabled").ToString
          objWorkflow.State = objRow.Item("state")

          ' Attach to table
          objTable = Globals.Things.GetObject(Type.Table, objRow.Item("tableid").ToString)
          objWorkflow.Table = objTable

          objTable.Objects.Add(objWorkflow)
        Next

      Catch ex As Exception

      Finally
        objDataset = Nothing
        objRow = Nothing

      End Try

    End Sub


    Private Function NullSafe(ByRef ObjectData As System.Data.DataRow, ByRef ColumnName As String, ByRef DefaultValue As Object)

      If ObjectData.IsNull(ColumnName) Then
        Return DefaultValue
      Else
        Return ObjectData.Item(ColumnName)
      End If

    End Function



  End Module

End Namespace
