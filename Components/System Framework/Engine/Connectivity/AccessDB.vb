Namespace Connectivity
  Public Class AccessDb
    Implements IConnection

    Public Db As OleDb.OleDbConnection
    Public NativeObject As DAO.Database

    Public Function ExecuteQuery(ByVal queryName As String, ByVal parms As Parameters) As DataSet Implements IConnection.ExecStoredProcedure

      Dim objAdapter As New OleDb.OleDbDataAdapter
      Dim sqlParms As OleDb.OleDbParameterCollection
      Dim objCommand As New OleDb.OleDbCommand

      Dim dsDataSet As New DataSet

      Try

        With objCommand
          .CommandType = CommandType.StoredProcedure
          .CommandText = queryName
          .Connection = Db
          '       .Connection.Open()

          ' Clear any previous parameters from the Command object
          Call .Parameters.Clear()

          If parms IsNot Nothing Then

            ' Convert passed in parameter array to sql parameters
            sqlParms = objCommand.Parameters
            For Each objParameter In parms

              Select Case objParameter.DBType
                Case DBType.Integer
                  sqlParms.AddWithValue(objParameter.Name, CInt(objParameter.Value))

                Case DBType.String
                  sqlParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)

                Case DBType.GUID
                  If objParameter.Value Is Nothing OrElse CType(objParameter.Value, Guid) = Guid.Empty Then
                    sqlParms.AddWithValue(objParameter.Name, DBNull.Value)
                  Else
                    sqlParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)
                  End If

              End Select

              .Connection.Close()

            Next
          End If
        End With

        objAdapter.SelectCommand = objCommand
        objAdapter.Fill(dsDataSet)

      Catch ex As Exception
        ErrorLog.Add(ErrorHandler.Section.LoadingData, "ExecuteQuery", ErrorHandler.Severity.Error, ex.Message, ex.InnerException.ToString)
        Return Nothing

      Finally

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

    Public Function ScriptStatement(ByVal statement As String, ByRef isCritical As Boolean) As Boolean Implements IConnection.ScriptStatement
      Return False
    End Function

  End Class
End Namespace
