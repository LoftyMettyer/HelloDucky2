Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Things.Collections

  <Serializable()> _
  Public Class Generic
    Inherits Things.Collections.BaseCollection
    Implements iCollection_Objects

    Public Function Setting(ByVal [Module] As String, ByVal [Parameter] As String) As Things.Setting Implements iCollection_Objects.Setting

      Dim objChild As Things.Base
      Dim objSetting As New Things.Setting

      For Each objChild In MyBase.Items
        If objChild.Type = Type.Setting Then
          objSetting = CType(objChild, Things.Setting)
          If objSetting.Module.ToLower = [Module].ToLower And objSetting.Parameter.ToLower = Parameter.ToLower Then
            Return objSetting
          End If
        End If
      Next

      Return New Things.Setting

    End Function

    Public Function Table(ByRef ID As HCMGuid) As Things.Table Implements iCollection_Objects.Table

      For Each objChild As Things.Base In MyBase.Items
        If objChild.ID = ID And objChild.Type = Type.Table Then
          Return CType(objChild, Things.Table)
        End If
      Next

      Return Nothing

    End Function

  End Class

End Namespace
