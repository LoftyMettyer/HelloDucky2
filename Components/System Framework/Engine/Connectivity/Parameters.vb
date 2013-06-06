Imports System.Runtime.InteropServices

Namespace Connectivity

  <ClassInterface(ClassInterfaceType.None)>
  Public Class Parameters
    Inherits Collection(Of Connectivity.Parameter)

    Public Overloads Sub Add(ByVal name As String, ByVal value As Integer)

      Dim param As New Connectivity.Parameter

      param.Name = [name]
      param.DBType = DBType.Integer
      param.Value = [value]
      Me.Items.Add(param)

    End Sub

    Public Overloads Sub Add(ByVal name As String, ByVal value As String)

      Dim param As New Connectivity.Parameter

      param.Name = name
      param.DBType = DBType.String

      If value Is Nothing Then
        value = ""
      End If

      param.Value = value
      Me.Items.Add(param)

    End Sub

  End Class

End Namespace

