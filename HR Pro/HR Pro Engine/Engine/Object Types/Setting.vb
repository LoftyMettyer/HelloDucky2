Namespace Things
  Public Class Setting
    Inherits Things.Base

    Public Property [Module] As String
    Public Property Parameter As String
    Public Property ParameterType As String
    Public Property Table As Things.Table
    Public Property Column As Things.Column
    Public Property Value As String
    Public Property Code As String
    Public Property SettingType As Enums.SettingType

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Setting
      End Get
    End Property

  End Class
End Namespace
