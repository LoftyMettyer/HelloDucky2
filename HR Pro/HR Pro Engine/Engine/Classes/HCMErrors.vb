Imports System.Runtime.InteropServices

Namespace ErrorHandler

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class Errors
    Inherits System.ComponentModel.BindingList(Of ErrorHandler.Error)
    Implements COMInterfaces.IErrors

    Private _isCatastrophic As Boolean

    Public Overloads Sub Add(ByVal Section As ErrorHandler.Section, ByVal ObjectName As String, ByVal Severity As SystemFramework.ErrorHandler.Severity, ByVal Message As String, ByVal Detail As String)

      Dim item As ErrorHandler.Error

      item.Section = Section
      item.ObjectName = ObjectName
      item.Severity = Severity
      item.Message = Message
      item.Detail = Detail
      item.User = Globals.Login.UserName
      item.DateTime = Now

      MyBase.Add(item)

    End Sub

    Public Function DetailedReport() As String

      Dim message As String = vbNullString

      For Each item As ErrorHandler.Error In Me.Items
        message += String.Format("{1}{0}{1}{2}{1}", item.Message, vbNewLine, item.Detail)
      Next

      Return message

    End Function

    Public Function QuickReport() As String

      Dim message As String = vbNullString

      For Each item As ErrorHandler.Error In Me.Items

        Select Case item.Severity
          Case Severity.Error
            message = message & String.Format("{0} - {1}", item.ObjectName, item.Message)

          Case Severity.Warning
            'sMessage = sMessage & String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)

        End Select

      Next

      Return message

    End Function

    Public Sub OutputToFile(ByVal FileName As String) Implements COMInterfaces.IErrors.OutputToFile

      System.IO.File.Delete(FileName)
      Dim objWriter As System.IO.StreamWriter = System.IO.File.AppendText(FileName)

      For Each objError As ErrorHandler.Error In Me.Items
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
    Public Section As ErrorHandler.Section
    Public ObjectName As String
    Public Severity As ErrorHandler.Severity
    Public Message As String
    Public Detail As String
    Public DateTime As Date
    Public User As String
  End Structure

  Public Enum Severity
    [Error] = 0
    Warning = 1
  End Enum

  Public Enum Section
    General = 0
    LoadingData = 1
    UDFs = 2
    Triggers = 3
    Views = 4
    TableAndColumns = 5
  End Enum

End Namespace
