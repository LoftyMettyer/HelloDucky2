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
          Case ValidationType.Mandatory
            sMessage = String.Format("{0} is mandatory.", Column.Name)

          Case ValidationType.Duplicate

          Case ValidationType.UniqueInTable, ValidationType.UniqueInSiblings
            sMessage = String.Format("{0} is not unique.", Column.Name)

          Case Else
            sMessage = vbNullString

        End Select

        Return sMessage

      End Get
    End Property


  End Class
End Namespace
