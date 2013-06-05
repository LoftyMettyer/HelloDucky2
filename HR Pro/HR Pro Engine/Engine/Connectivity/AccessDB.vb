Namespace Connectivity
  Public Class AccessDB
    Implements IConnection

    Public DB As OleDb.OleDbConnection
    Public NativeObject As DAO.Database

    Public Function ExecuteQuery(ByVal QueryName As String, ByVal Parms As Connectivity.Parameters) As System.Data.DataSet Implements IConnection.ExecStoredProcedure

      Dim objAdapter As New OleDb.OleDbDataAdapter
      Dim sqlParms As OleDb.OleDbParameterCollection
      Dim sqlParm As OleDb.OleDbParameter
      Dim objCommand As New OleDb.OleDbCommand

      Dim dsDataSet As New DataSet

      Try

        With objCommand
          .CommandType = CommandType.StoredProcedure
          .CommandText = QueryName
          .Connection = DB
          '       .Connection.Open()

          ' Clear any previous parameters from the Command object
          Call .Parameters.Clear()

          If Parms IsNot Nothing Then

            ' Convert passed in parameter array to sql parameters
            sqlParms = objCommand.Parameters
            For Each objParameter In Parms

              Select Case objParameter.DBType
                Case Connectivity.DBType.Integer
                  sqlParm = sqlParms.AddWithValue(objParameter.Name, CInt(objParameter.Value))

                Case Connectivity.DBType.String
                  sqlParm = sqlParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)

                Case Connectivity.DBType.GUID
                  If objParameter.Value Is Nothing OrElse CType(objParameter.Value, Guid) = Guid.Empty Then
                    sqlParm = sqlParms.AddWithValue(objParameter.Name, DBNull.Value)
                  Else
                    sqlParm = sqlParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)
                  End If

              End Select

              .Connection.Close()

            Next
          End If
        End With

        objAdapter.SelectCommand = objCommand
        objAdapter.Fill(dsDataSet)

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, "ExecuteQuery", SystemFramework.ErrorHandler.Severity.Error, ex.Message, ex.InnerException.ToString)
        Return Nothing

      Finally
        objAdapter = Nothing
        sqlParms = Nothing
        sqlParm = Nothing
        objCommand = Nothing

      End Try

      Return dsDataSet

    End Function

    Public Sub Close() Implements IConnection.Close
      DB.Close()
      NativeObject.Close()

      DB.Dispose()

      DB = Nothing
      NativeObject = Nothing

    End Sub

    Public Property Login As Login Implements IConnection.Login
      Get
        Return Nothing
      End Get
      Set(ByVal value As Login)

      End Set
    End Property

    Public Sub Open() Implements IConnection.Open
    End Sub

    Public Function ScriptStatement(ByVal statement As String) As Boolean Implements IConnection.ScriptStatement
      Return False
    End Function

  End Class
End Namespace
