Imports System.Data.SqlClient

Namespace Connectivity
  Public Class ADOClassic
    Implements COMInterfaces.IConnection

    Public DB As OleDb.OleDbConnection
    Public NativeObject As ADODB.Connection

#Region "IConnection interface"

    Public Sub Close() Implements COMInterfaces.IConnection.Close
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

    Public Function ExecStoredProcedure(ByVal ProcedureName As String, ByVal Parms As Parameters) As System.Data.DataSet Implements COMInterfaces.IConnection.ExecStoredProcedure

      Dim objAdapter As New OleDb.OleDbDataAdapter
      Dim sqlParm As New ADODB.Parameter
      Dim objCommand As New ADODB.Command
      Dim rsDataset As ADODB.Recordset
      Dim dsDataSet As New DataSet

      Try

        With objCommand
          .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
          .CommandText = ProcedureName
          .ActiveConnection = NativeObject

          If Parms IsNot Nothing Then

            For Each objParameter In Parms

              Select Case objParameter.DBType
                Case Connectivity.DBType.Integer
                  .Parameters(objParameter.Name).Value = CInt(objParameter.Value.ToString)
                Case Connectivity.DBType.String
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

    Public Property Login As Login Implements COMInterfaces.IConnection.Login
      Get
        Return Nothing
      End Get
      Set(ByVal value As Login)

      End Set
    End Property

    Public Sub Open() Implements COMInterfaces.IConnection.Open
    End Sub

    Public Function ScriptStatement(ByVal statement As String, ByRef IsCritical As Boolean) As Boolean Implements COMInterfaces.IConnection.ScriptStatement

      Dim bOK As Boolean = True
      Dim iSeverity As ErrorHandler.Severity

      Try
        NativeObject.Execute(statement)

      Catch ex As Exception
        iSeverity = CType(IIf(IsCritical = True, ErrorHandler.Severity.Error, ErrorHandler.Severity.Warning), ErrorHandler.Severity)
        Globals.ErrorLog.Add(ErrorHandler.Section.General, "ADOClassic.ScriptStatement", iSeverity, ex.Message, statement)
        bOK = False
      End Try

      Return bOK

    End Function

#End Region

  End Class
End Namespace
