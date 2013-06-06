﻿Imports System.Runtime.InteropServices

Namespace ErrorHandler

  <ClassInterface(ClassInterfaceType.None)>
  Public Class Errors
    Inherits Collection(Of [Error])
    Implements IErrors

    Private _isCatastrophic As Boolean

    Public Overloads Sub Add(ByVal section As Section, ByVal objectName As String, ByVal severity As Severity, ByVal message As String, ByVal detail As String)

      Dim item As [Error]

      item.Section = section
      item.ObjectName = objectName
      item.Severity = severity
      item.Message = message
      item.Detail = Detail
      item.User = Globals.Login.UserName
      item.DateTime = Now

      If Not Items.Any(Function(e) e.Section = item.Section AndAlso
            e.ObjectName = item.ObjectName AndAlso
            e.Section = item.Section AndAlso
            e.Message = item.Message AndAlso
            e.Detail = item.Detail AndAlso
            e.User = item.User) Then
        Add(item)
      End If

    End Sub

    Public Function DetailedReport() As String

      Dim message As String = vbNullString

      For Each item As [Error] In Me.Items
        message += String.Format("{1}{0}{1}{2}{1}", item.Message, vbNewLine, item.Detail)
      Next

      Return message

    End Function

    Public Function QuickReport() As String

      Dim message As String = vbNullString

      For Each item As [Error] In Items

        Select Case item.Severity
          Case Severity.Error
            message += String.Format("{0} - {1}", item.ObjectName, item.Message)

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
