Option Strict On

Imports System.Runtime.InteropServices
'Imports System.IO

Namespace ErrorHandler

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class Errors
    Inherits System.ComponentModel.BindingList(Of ErrorHandler.Error)
    Implements iErrors

    Public Shadows Sub Add(ByVal Section As Phoenix.ErrorHandler.Section, ByVal ObjectName As String, ByVal Severity As Phoenix.ErrorHandler.Severity, ByVal Message As String, ByVal Detail As String)

      Dim objError As Phoenix.ErrorHandler.Error

      objError.Section = Section
      objError.ObjectName = ObjectName
      objError.Severity = Severity
      objError.Message = Message
      objError.Detail = Detail
      objError.User = "A user"
      objError.DateTime = Now

      MyBase.Add(objError)

    End Sub

    Public Sub OutputToFile(ByRef FileName As String) Implements Interfaces.iErrors.OutputToFile

      Dim objWriter As System.IO.StreamWriter
      Dim objError As Phoenix.ErrorHandler.Error
      Dim sMessage As String

      objWriter = System.IO.File.AppendText(FileName)

      For Each objError In Me.Items

        sMessage = String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)
        objWriter.Write(sMessage)

      Next

      objWriter.Close()

    End Sub

  End Class

  Public Structure [Error]
    Public Section As Phoenix.ErrorHandler.Section
    Public ObjectName As String
    Public Severity As Phoenix.ErrorHandler.Severity
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
