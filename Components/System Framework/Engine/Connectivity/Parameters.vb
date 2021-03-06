﻿Imports System.Runtime.InteropServices
Imports SystemFramework.Enums

Namespace Connectivity

  <ClassInterface(ClassInterfaceType.None)>
  Public Class Parameters
    Inherits Collection(Of Parameter)

    Public Overloads Sub Add(ByVal name As String, ByVal value As Integer)

      Dim param As New Parameter

      param.Name = [name]
      param.DbType = Connection.DbType.Integer
      param.Value = [value]
      Items.Add(param)

    End Sub

    Public Overloads Sub Add(ByVal name As String, ByVal value As String)

      Dim param As New Parameter

      param.Name = name
      param.DbType = Connection.DbType.String

      If value Is Nothing Then
        value = ""
      End If

      param.Value = value
      Items.Add(param)

    End Sub

  End Class

End Namespace

