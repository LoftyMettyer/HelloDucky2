Namespace Things
  Public Class Setting
    Inherits Things.Base

    Public [Module] As String
    Public Parameter As String
    Public ParameterType As String
    Public Table As Things.Table
    Public Column As Things.Column
    Public Value As String
    Public Code As String
    Public SettingType As Enums.SettingType

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Setting
      End Get
    End Property

  End Class
End Namespace
