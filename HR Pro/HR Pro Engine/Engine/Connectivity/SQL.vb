Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Namespace Connectivity

  Public Class SQL
    Implements IConnection

    Private mobjLogin As Connectivity.Login
    Private mConn As New SqlConnection

    Public Property Login As Connectivity.Login Implements IConnection.Login
      Get
        Return mobjLogin
      End Get
      Set(ByVal value As Connectivity.Login)
        mobjLogin = value
      End Set
    End Property

    Public Sub Open() Implements IConnection.Open

      Dim sConnection As String = vbNullString

      Try

        If Login.UseContext Then
          mConn = New SqlConnection("context connection=true")
        Else

          sConnection = String.Format("Initial Catalog={0}; Server={1};" _
                              & "User ID={2}; Password={3}; APP={4};" _
                              , Login.Database, Login.Server _
                              , Login.UserName, Login.Password _
                              , "ScriptDB")
          mConn = New SqlConnection(sConnection)

        End If

        mConn.Open()

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.LoadingData, String.Empty, SystemFramework.ErrorHandler.Severity.Error, ex.Message, sConnection)

      End Try

    End Sub

    Public Function ExecStoredProcedure(ByVal ProcedureName As String, ByVal Parms As Connectivity.Parameters) As System.Data.DataSet Implements IConnection.ExecStoredProcedure

      Dim SQLParms As SqlParameterCollection
      Dim objParameter As Connectivity.Parameter
      Dim sqlParm As SqlParameter
      Dim _cmdSQLCommand As New SqlCommand
      Dim _adpAdapter As New SqlClient.SqlDataAdapter
      Dim dsDataSet As New DataSet
      Dim bOK As Boolean

      ' Configure the SqlCommand object
      With _cmdSQLCommand
        .CommandType = CommandType.StoredProcedure      'Set type to StoredProcedure
        .CommandText = ProcedureName                    'Specify stored procedure to run
        .Connection = mConn

        ' Clear any previous parameters from the Command object
        Call .Parameters.Clear()

        ' Convert passed in parameter array to sql parameters
        SQLParms = _cmdSQLCommand.Parameters
        For Each objParameter In Parms

          Select Case objParameter.DBType
            Case Connectivity.DBType.Integer
              sqlParm = SQLParms.AddWithValue(objParameter.Name, CInt(objParameter.Value))

            Case Connectivity.DBType.String
              sqlParm = SQLParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)

            Case Connectivity.DBType.GUID
              If objParameter.Value Is Nothing OrElse CType(objParameter.Value, Guid) = Guid.Empty Then
                sqlParm = SQLParms.AddWithValue(objParameter.Name, DBNull.Value)
              Else
                sqlParm = SQLParms.AddWithValue(objParameter.Name, objParameter.Value.ToString)
              End If

          End Select

        Next

      End With

      ' Configure Adapter to use newly created command object and fill the dataset.
      Try
        _adpAdapter.SelectCommand = _cmdSQLCommand
        _adpAdapter.Fill(dsDataSet)

      Catch ex As Exception
        bOK = False

      End Try

      Return dsDataSet
    End Function

    Public Sub Close() Implements IConnection.Close
      mConn.Close()
    End Sub

    Public Function ScriptStatement(ByVal Statement As String) As Boolean Implements IConnection.ScriptStatement

      Dim objCommand As New SqlCommand
      Dim bOK As Boolean

      objCommand.CommandText = Statement

      'Try

      If Login.UseContext Then
        Microsoft.SqlServer.Server.SqlContext.Pipe.ExecuteAndSend(objCommand)
      Else
        objCommand.Connection = mConn
        objCommand.ExecuteNonQuery()
      End If

      bOK = True

      'Catch ex As Exception
      '  Globals.ErrorLog.Add(HCM.ErrorHandler.Section.General, String.Empty, HCM.ErrorHandler.Severity.Error, ex.Message, Statement)
      '  bOK = False

      'Finally
      '  objCommand.Dispose()

      'End Try

      Return bOK

    End Function

  End Class


End Namespace
