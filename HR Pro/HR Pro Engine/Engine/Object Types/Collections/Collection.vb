Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Things.Collections

  <Serializable()> _
  Public Class Generic
    Inherits Things.Collections.BaseCollection
    Implements ICollection_Objects

    'TODO: WANT TO REMOVE
    Public Function Setting(ByVal [Module] As String, ByVal [Parameter] As String) As Things.Setting Implements ICollection_Objects.Setting

      Dim objSetting As New Things.Setting

      For Each objChild As Things.Base In MyBase.Items
        If objChild.Type = Type.Setting Then
          objSetting = CType(objChild, Things.Setting)
          If objSetting.Module.ToLower = [Module].ToLower And objSetting.Parameter.ToLower = Parameter.ToLower Then
            Return objSetting
          End If
        End If
      Next

      Return New Things.Setting

    End Function

    'TODO: WANT TO REMOVE
    Public Function Table(ByVal ID As Integer) As Things.Table Implements ICollection_Objects.Table

      For Each objChild As Things.Base In MyBase.Items
        If objChild.ID = ID And objChild.Type = Type.Table Then
          Return CType(objChild, Things.Table)
        End If
      Next

      Return Nothing

    End Function

  End Class

End Namespace
