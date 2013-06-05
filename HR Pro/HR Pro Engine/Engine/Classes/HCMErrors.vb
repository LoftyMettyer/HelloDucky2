Option Strict On

Imports System.Runtime.InteropServices

Namespace ErrorHandler

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class Errors
    Inherits System.ComponentModel.BindingList(Of ErrorHandler.Error)
    Implements COMInterfaces.IErrors

    Private mbIsCatastrophic As Boolean

    Public Shadows Sub Add(ByVal Section As ErrorHandler.Section, ByVal ObjectName As String, ByVal Severity As SystemFramework.ErrorHandler.Severity, ByVal Message As String, ByVal Detail As String)

      Dim objError As SystemFramework.ErrorHandler.Error

      objError.Section = Section
      objError.ObjectName = ObjectName
      objError.Severity = Severity
      objError.Message = Message
      objError.Detail = Detail
      objError.User = Globals.Login.UserName
      objError.DateTime = Now

      MyBase.Add(objError)

    End Sub

    Public Function DetailedReport() As String

      Dim sMessage As String = vbNullString

      Try

        For Each objError In Me.Items
          sMessage = sMessage & String.Format("{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)
        Next

      Catch ex As Exception

      End Try

      Return sMessage

    End Function

    Public Function QuickReport() As String

      Dim sMessage As String = vbNullString

      Try

        For Each objError In Me.Items

          Select Case objError.Severity
            Case Severity.Error
              sMessage = sMessage & String.Format("{0} - {1}", objError.ObjectName, objError.Message)

            Case Severity.Warning
              'sMessage = sMessage & String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)

          End Select

        Next

      Catch ex As Exception

      End Try

      Return sMessage

    End Function

    Public Sub OutputToFile(ByRef FileName As String) Implements COMInterfaces.IErrors.OutputToFile

      Dim objWriter As System.IO.StreamWriter
      Dim objError As SystemFramework.ErrorHandler.Error
      Dim sMessage As String

      Try

        System.IO.File.Delete(FileName)
        objWriter = System.IO.File.AppendText(FileName)

        For Each objError In Me.Items
          sMessage = String.Format("{1}{1}{1}{1}{0}{1}{2}{1}", objError.Message, vbNewLine, objError.Detail)
          objWriter.Write(sMessage)
        Next

        objWriter.Close()

      Catch ex As Exception

      End Try

    End Sub

    Public ReadOnly Property IsCatastrophic As Boolean Implements IErrors.IsCatastrophic
      Get
        Return mbIsCatastrophic
      End Get
    End Property

    Public Sub Show() Implements IErrors.Show

      Dim frmErrorLog As New Forms.ErrorLog

      Try
        frmErrorLog.ShowDialog()
        mbIsCatastrophic = frmErrorLog.Abort

      Catch ex As Exception

      Finally
        frmErrorLog = Nothing

      End Try

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
