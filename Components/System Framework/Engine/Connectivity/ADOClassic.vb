
Namespace Connectivity
  Public Class AdoClassic
    Implements IConnection

    Public DB As OleDb.OleDbConnection
    Public NativeObject As ADODB.Connection

#Region "IConnection interface"

    Public Sub Close() Implements ComInterfaces.IConnection.Close
      DB.Close()
      'NativeObject.Close()
    End Sub

    Public Sub BeginTrans()
      NativeObject.BeginTrans()
    End Sub

    Public Sub RollbackTrans()
      NativeObject.RollbackTrans()
    End Sub

    Public Sub CommitTrans()
      NativeObject.CommitTrans()
    End Sub

    Public Function ExecSql(ByVal sql As String) As DataSet

      Dim command As New ADODB.Command
      command.CommandType = ADODB.CommandTypeEnum.adCmdText
      command.CommandText = sql
      command.ActiveConnection = NativeObject

      Dim rs As ADODB.Recordset = command.Execute

      Dim da As New OleDb.OleDbDataAdapter
      Dim ds As New DataSet

      da.Fill(ds, rs, "mytable")

      Return ds

    End Function

    Public Function ExecStoredProcedure(ByVal queryName As String, ByVal parms As Parameters) As DataSet Implements IConnection.ExecStoredProcedure

      Dim objAdapter As New OleDb.OleDbDataAdapter
      Dim objCommand As New ADODB.Command
      Dim rsDataset As ADODB.Recordset
      Dim dsDataSet As New DataSet

      Try

        With objCommand
          .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
          .CommandText = queryName
          .ActiveConnection = NativeObject

          If parms IsNot Nothing Then

            For Each objParameter In parms

              Select Case objParameter.DbType
                Case Connectivity.DbType.Integer
                  .Parameters(objParameter.Name).Value = CInt(objParameter.Value.ToString)
                Case Connectivity.DbType.String
                  .Parameters(objParameter.Name).Value = objParameter.Value.ToString
              End Select
            Next
          End If

          rsDataset = .Execute
        End With

        ' Convert recordset to ADO.NET dataset
        objAdapter.Fill(dsDataSet, rsDataset, "mytable")

        Return dsDataSet

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, "ExecuteQuery", SystemFramework.ErrorHandler.Severity.Error, ex.Message, ex.InnerException.ToString)
        Return Nothing

      End Try

    End Function

    Public Property Login As Login Implements ComInterfaces.IConnection.Login
      Get
        Return Nothing
      End Get
      Set(ByVal value As Login)

      End Set
    End Property

    Public Sub Open() Implements ComInterfaces.IConnection.Open
    End Sub

    Public Function ScriptStatement(ByVal statement As String, ByRef isCritical As Boolean) As Boolean Implements ComInterfaces.IConnection.ScriptStatement

      Dim bOk As Boolean = True
      Dim iSeverity As ErrorHandler.Severity

      Try
        NativeObject.Execute(statement)

      Catch ex As Exception
        iSeverity = CType(IIf(isCritical = True, ErrorHandler.Severity.Error, ErrorHandler.Severity.Warning), ErrorHandler.Severity)
        ErrorLog.Add(ErrorHandler.Section.General, "ADOClassic.ScriptStatement", iSeverity, ex.Message, statement)
        bOk = False
      End Try

      Return bOk

    End Function

#End Region

  End Class
End Namespace
