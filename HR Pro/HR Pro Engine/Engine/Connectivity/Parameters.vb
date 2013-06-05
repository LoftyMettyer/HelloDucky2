Imports System.Runtime.InteropServices

Namespace Connectivity

  <ClassInterface(ClassInterfaceType.None)> _
    Public Class Parameters
    Inherits System.ComponentModel.BindingList(Of Connectivity.Parameter)

    Public Shadows Sub Add(ByVal [Name] As String, ByVal [Value] As Integer)

      Dim objParameter As New Connectivity.Parameter

      objParameter.Name = [Name]
      objParameter.DBType = DBType.Integer
      objParameter.Value = [Value]
      Me.Items.Add(objParameter)

    End Sub

    Public Shadows Sub Add(ByVal [Name] As String, ByVal [Value] As String)

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

