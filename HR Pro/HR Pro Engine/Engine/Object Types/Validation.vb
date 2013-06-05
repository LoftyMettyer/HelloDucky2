Namespace Things
  Public Class Validation
    Inherits Base

    Public Property ValidationType As ValidationType
    Public Property Column As Column
    Public Property Severity As ValidationSeverity = ValidationSeverity.Error

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
