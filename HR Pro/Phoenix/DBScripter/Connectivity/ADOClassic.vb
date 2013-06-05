﻿Imports System.Data.SqlClient

Namespace Connectivity
  Public Class ADOClassic
    Implements iConnection

    Public DB As OleDb.OleDbConnection
    Public NativeObject As ADODB.Connection

#Region "IConnection interface"

    Public Sub Close() Implements Interfaces.iConnection.Close
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

    Public Function ExecStoredProcedure(ByVal ProcedureName As String, ByRef Parms As Parameters) As System.Data.DataSet Implements Interfaces.iConnection.ExecStoredProcedure

      Dim objAdapter As New OleDb.OleDbDataAdapter
      '      Dim sqlParms As OleDb.OleDbParameterCollection
      Dim sqlParm As New ADODB.Parameter
      Dim objCommand As New ADODB.Command
      Dim rsDataset As ADODB.Recordset

      Dim dsDataSet As New DataSet

      Try

        With objCommand
          .CommandType = CommandType.StoredProcedure
          .CommandText = ProcedureName
          .ActiveConnection = NativeObject

          ' Convert passed in parameter array to sql parameters
          '          sqlParms = objCommand.Parameters
          For Each objParameter In Parms

            Select Case objParameter.DBType
              Case Connectivity.DBType.Integer
                '    sqlParm = .CreateParameter(objParameter.Name, ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 0, CInt(objParameter.Value.ToString))
                .Parameters(objParameter.Name).Value = CInt(objParameter.Value.ToString)

              Case Connectivity.DBType.String
                '                sqlParm = .CreateParameter(objParameter.Name, ADODB.DataTypeEnum.adLongVarChar, ADODB.ParameterDirectionEnum.adParamInput, 0, objParameter.Value.ToString)
                .Parameters(objParameter.Name).Value = objParameter.Value.ToString

                'Case Connectivity.DBType.GUID
                '  If objParameter.Value = System.Guid.Empty Then
                '    sqlParm = .CreateParameter(objParameter.Name, ADODB.DataTypeEnum.adGUID, ADODB.ParameterDirectionEnum.adParamInput, 0, DBNull.Value)
                '  Else
                '    sqlParm = .CreateParameter(objParameter.Name, ADODB.DataTypeEnum.adGUID, ADODB.ParameterDirectionEnum.adParamInput, 0, objParameter.Value.ToString)
                '  End If

            End Select

            '          .Parameters.Append(sqlParm)

          Next

          rsDataset = .Execute

        End With


        ' Convert recordset to ADO.NET dataset
        objAdapter.Fill(dsDataSet, rsDataset, "mytable")

        ' objAdapter.Fill(rsDataset, dsDataSet)

        '        objAdapter.SelectCommand = objCommand
        '       objAdapter.Fill(dsDataSet)
        objAdapter = Nothing
        objCommand = Nothing

        Return dsDataSet

      Catch ex As Exception
        Globals.ErrorLog.Add(DBScripter.ErrorHandler.Section.LoadingData, "ExecuteQuery", DBScripter.ErrorHandler.Severity.Error, ex.Message, ex.InnerException.ToString)
        Return Nothing

      End Try

    End Function

    Public Property Login As Login Implements Interfaces.iConnection.Login
      Get

      End Get
      Set(ByVal value As Login)

      End Set
    End Property

    Public Sub Open() Implements Interfaces.iConnection.Open
    End Sub

    Public Function ScriptStatement(ByVal Statement As String) As Boolean Implements Interfaces.iConnection.ScriptStatement

      Dim bOK As Boolean = True

      Try
        NativeObject.Execute(Statement)

      Catch ex As Exception
        Globals.ErrorLog.Add(ErrorHandler.Section.General, "ADOClassic.ScriptStatement", 1, ex.Message, Statement)
        bOK = False
      End Try

      Return bOK

    End Function

#End Region

  End Class
End Namespace
