Imports System.Runtime.InteropServices

Namespace Connectivity

  <ClassInterface(ClassInterfaceType.None)> _
    Public Class Parameters
    Inherits System.ComponentModel.BindingList(Of Connectivity.Parameter)

    Public Shadows Sub Add(ByRef [Name] As String, ByRef [Value] As HCMGuid)

      Dim objParameter As New Connectivity.Parameter

      objParameter.Name = [Name]
      objParameter.DBType = DBType.Integer
      objParameter.Value = [Value]
      Me.Items.Add(objParameter)

    End Sub

    Public Shadows Sub Add(ByRef [Name] As String, ByRef [Value] As Integer)

      Add([Name], Value.ToString)

    End Sub

    Public Shadows Sub Add(ByRef [Name] As String, ByRef [Value] As String)

      Dim objParameter As New Connectivity.Parameter

      objParameter.Name = [Name]
      objParameter.DBType = DBType.String

      If [Value] Is Nothing Then
        [Value] = ""
      End If

      objParameter.Value = [Value]
      Me.Items.Add(objParameter)

    End Sub

  End Class

End Namespace

