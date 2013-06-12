Imports System.Runtime.InteropServices
Imports SystemFramework.Enums.Errors
Imports SystemFramework.Structures

Namespace Collections

  <ClassInterface(ClassInterfaceType.None)>
  Public Class Errors
    Inherits Collection(Of [Error])
    Implements IErrors

    Private _isCatastrophic As Boolean

    Public Overloads Sub Add(ByVal section As Section, ByVal objectName As String, ByVal severity As Severity, ByVal message As String, ByVal detail As String)

      Dim thisItem As [Error]

      thisItem.Section = section
      thisItem.ObjectName = objectName
      thisItem.Severity = severity
      thisItem.Message = message
      thisItem.Detail = detail
      thisItem.User = Login.UserName

      If Not Items.Any(Function(e) e.Section = thisItem.Section AndAlso
            e.ObjectName = thisItem.ObjectName AndAlso
            e.Section = thisItem.Section AndAlso
            e.Message = thisItem.Message AndAlso
            e.Detail = thisItem.Detail AndAlso
            e.User = thisItem.User) Then
        Add(thisItem)
      End If

    End Sub

    Public Function DetailedReport() As String

      'For Each thisItem As [Error] In Items
      '  message += String.Format("{1}{0}{1}{2}{1}", thisItem.Message, vbNewLine, thisItem.Detail)
      'Next

      Return Items.Aggregate(vbNullString, Function(current, thisItem) current + String.Format("{1}{0}{1}{2}{1}", thisItem.Message, vbNewLine, thisItem.Detail))

    End Function

    Public Sub OutputToFile(ByVal fileName As String) Implements IErrors.OutputToFile

      IO.File.Delete(fileName)
      Dim objWriter As IO.StreamWriter = IO.File.AppendText(fileName)

      For Each objError As Structures.Error In Items
        Dim message As String = String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)
        objWriter.Write(message)
      Next

      objWriter.Close()

    End Sub

    Public ReadOnly Property IsCatastrophic As Boolean Implements IErrors.IsCatastrophic
      Get
        Return _isCatastrophic
      End Get
    End Property

    Public Sub Show() Implements IErrors.Show

      Using frm As New Forms.ErrorLog
        frm.ShowDialog()
        _isCatastrophic = frm.Abort
      End Using

    End Sub

    Public ReadOnly Property ErrorCount As Integer Implements IErrors.ErrorCount
      Get
        Return Items.Count
      End Get
    End Property

  End Class

End Namespace
