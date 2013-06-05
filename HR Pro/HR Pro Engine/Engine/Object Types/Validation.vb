Namespace Things
  Public Class Validation
    Inherits Things.Base

    Public ValidationType As Things.ValidationType
    Public Column As Things.Column
    Public Severity As Things.ValidationSeverity = ValidationSeverity.Error

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Validation
      End Get
    End Property

    Public ReadOnly Property Message As String
      Get

        Dim sMessage As String = vbNullString

        Select Case Me.ValidationType
          Case Enums.ValidationType.Mandatory
            sMessage = String.Format("{0} is mandatory.", Column.Name)

          Case Enums.ValidationType.Duplicate

          Case Enums.ValidationType.UniqueInTable, Enums.ValidationType.UniqueInSiblings
            sMessage = String.Format("{0} is not unique.", Column.Name)

          Case Else
            sMessage = vbNullString

        End Select

        Return sMessage

      End Get
    End Property


  End Class
End Namespace
