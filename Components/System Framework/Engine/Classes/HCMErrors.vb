Imports System.Runtime.InteropServices

Namespace ErrorHandler

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
      thisItem.DateTime = Now

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

    Public Function QuickReport() As String

      Dim message As String = vbNullString

      For Each thisItem As [Error] In Items

        Select Case thisItem.Severity
          Case Severity.Error
            message += String.Format("{0} - {1}", thisItem.ObjectName, thisItem.Message)

          Case Severity.Warning
            'sMessage = sMessage & String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)

        End Select

      Next

      Return message

    End Function

    Public Sub OutputToFile(ByVal fileName As String) Implements IErrors.OutputToFile

      IO.File.Delete(fileName)
      Dim objWriter As IO.StreamWriter = IO.File.AppendText(fileName)

      For Each objError As ErrorHandler.Error In Items
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

  Public Structure [Error]
    Public Id As Guid
    Public Section As Section
		Public ObjectName As String
    Public Severity As Severity
		Public Message As String
		Public Detail As String
		Public DateTime As Date
		Public User As String
		Public ErrorNumber As Long
    Public ErrorArticleId As Long


  End Structure

  Public Enum Severity
    [Error] = 0
    Warning = 1
  End Enum

  Public Enum Section
    General = 0
    LoadingData = 1
    UdFs = 2
    Triggers = 3
    Views = 4
    TableAndColumns = 5
  End Enum

End Namespace
