Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Things

  <Serializable()> _
  Public Class Collection
    Inherits Things.BaseCollection
    Implements iObjectCollection

    Public Function Setting(ByVal [Module] As String, ByVal [Parameter] As String) As Things.Setting Implements iObjectCollection.Setting

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

    Public Function Table(ByRef ID As HCMGuid) As Things.Table Implements iObjectCollection.Table

      Dim objChild As Things.Base

      For Each objChild In MyBase.Items
        If objChild.ID = ID And objChild.Type = Type.Table Then
          Return objChild
        End If
      Next

      Return Nothing

    End Function

  End Class

End Namespace
