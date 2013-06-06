Imports System.Data.SqlClient

Namespace Connectivity

  Public Class Sql
    Implements IConnection

    Private _mobjLogin As Login
    Private _mConn As New SqlConnection

    Public Property Login As Login Implements IConnection.Login
      Get
        Return _mobjLogin
      End Get
      Set(ByVal value As Login)
        _mobjLogin = value
      End Set
    End Property

    Public Sub Open() Implements IConnection.Open

      Dim sConnection As String = vbNullString

      Try

        If Login.UseContext Then
          _mConn = New SqlConnection("context connection=true")
        Else

          sConnection = String.Format("Initial Catalog={0}; Server={1};" _
                              & "User ID={2}; Password={3}; APP={4};" _
                              , Login.Database, Login.Server _
                              , Login.UserName, Login.Password _
                              , "ScriptDB")
          _mConn = New SqlConnection(sConnection)

        End If

        _mConn.Open()

      Catch ex As Exception
        ErrorLog.Add(ErrorHandler.Section.LoadingData, String.Empty, ErrorHandler.Severity.Error, ex.Message, sConnection)

      End Try

    End Sub

    Public Function ExecStoredProcedure(ByVal queryName As String, ByVal parms As Parameters) As DataSet Implements IConnection.ExecStoredProcedure

      Dim sqlParms As SqlParameterCollection
      Dim objParameter As Parameter
      Dim cmdSqlCommand As New SqlCommand
      Dim adpAdapter As New SqlDataAdapter
      Dim dsDataSet As New DataSet

      ' Configure the SqlCommand object
      With cmdSqlCommand
        .CommandType = CommandType.StoredProcedure      'Set type to StoredProcedure
        .CommandText = queryName                    'Specify stored procedure to run
        .Connection = _mConn

        ' Clear any previous parameters from the Command object
        Call .Parameters.Clear()

        ' Convert passed in parameter array to sql parameters
        sqlParms = cmdSqlCommand.Parameters
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

        Next

      End With

      ' Configure Adapter to use newly created command object and fill the dataset.
      Try
        adpAdapter.SelectCommand = cmdSqlCommand
        adpAdapter.Fill(dsDataSet)

      Catch ex As Exception

      End Try

      Return dsDataSet
    End Function

    Public Sub Close() Implements IConnection.Close
      _mConn.Close()
    End Sub

    Public Function ScriptStatement(ByVal statement As String, ByRef isCritical As Boolean) As Boolean Implements IConnection.ScriptStatement

      Dim objCommand As New SqlCommand
      Dim bOk As Boolean

      objCommand.CommandText = statement

      If Login.UseContext Then
        Microsoft.SqlServer.Server.SqlContext.Pipe.ExecuteAndSend(objCommand)
      Else
        objCommand.Connection = _mConn
        objCommand.ExecuteNonQuery()
      End If

      bOk = True

      Return bOk

    End Function

  End Class


End Namespace
