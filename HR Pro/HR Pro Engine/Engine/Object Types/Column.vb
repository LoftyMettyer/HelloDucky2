﻿Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Column
    Inherits Base

    Public Property Table As Table

    Public Property CalcID As Integer
    Public Property Calculation As Expression

    Public Property DataType As ScriptDB.ColumnTypes
    Public Property Size As Integer
    Public Property Decimals As Integer

    ' Options
    Public Property Audit As Boolean
    Public Property Multiline As Boolean
    Public Property CalculateIfEmpty As Boolean
    Public Property IsReadOnly As Boolean

    ' Formatting
    Public Property CaseType As Enums.CaseType
    Public Property TrimType As Enums.TrimType
    Public Property Alignment As Enums.AlignType
    Public Property Mandatory As Boolean
    Public Property OLEType As ScriptDB.OLEType

    Public Property UniqueType As Enums.UniqueCheckScope

    Public Property DefaultCalcID As Integer
    Public Property DefaultCalculation As Expression
    Public Property DefaultValue As String

    Public ReadOnly Property DataTypeSize As String
      Get

        Select Case Me.DataType
          Case ScriptDB.ColumnTypes.Text
            If Me.Multiline Or Me.Size > 8000 Then
              Return "MAX"
            Else
              Return CStr(Me.Size)
            End If

          Case ScriptDB.ColumnTypes.Numeric
            Return (Me.Size + Me.Decimals).ToString

          Case ScriptDB.ColumnTypes.Logic
            Return "1"

          Case ScriptDB.ColumnTypes.Date
            Return "20"

          Case Else
            Return CStr(Me.Size)

        End Select
      End Get
    End Property

    ' Declaration syntax for a column type
    Public ReadOnly Property DataTypeSyntax As String
      Get

        Dim sqlType As String = String.Empty

        Select Case Me.DataType
          Case ScriptDB.ColumnTypes.Text
            If Me.Multiline Or Me.Size > 8000 Then
              sqlType = "varchar(MAX)"
            Else
              sqlType = String.Format("varchar({0})", Me.Size)
            End If

          Case ScriptDB.ColumnTypes.Integer
            sqlType = String.Format("integer")

          Case ScriptDB.ColumnTypes.Numeric
            sqlType = String.Format("numeric({0},{1})", Me.Size, Me.Decimals)

          Case ScriptDB.ColumnTypes.Date
            sqlType = "datetime"

          Case ScriptDB.ColumnTypes.Logic
            sqlType = "bit"

          Case ScriptDB.ColumnTypes.WorkingPattern
            sqlType = "varchar(14)"

          Case ScriptDB.ColumnTypes.Link
            sqlType = "varchar(255)"

          Case ScriptDB.ColumnTypes.Photograph
            sqlType = "varchar(255)"

          Case ScriptDB.ColumnTypes.Binary
            sqlType = "varbinary(MAX)"

        End Select

        Return sqlType

      End Get

    End Property

    Public ReadOnly Property HasDefaultValue As Boolean
      Get

        If Me.DefaultCalcID > 0 Then
          Return True
        Else
          Select Case Me.DataType
            Case ScriptDB.ColumnTypes.Text
              Return Me.DefaultValue.Length > 0
            Case ScriptDB.ColumnTypes.Numeric, ScriptDB.ColumnTypes.Integer
              Return CInt(Me.DefaultValue) <> 0
            Case ScriptDB.ColumnTypes.Logic
              Return CBool(Me.DefaultValue) <> True
            Case Else
              Return False
          End Select
        End If

      End Get
    End Property


    Public ReadOnly Property IsCalculated As Boolean
      Get
        Return CalcID > 0
      End Get
    End Property

    Public Function ApplyFormatting() As String
      Return ApplyFormatting(String.Empty)
    End Function

    Public Function ApplyFormatting(ByVal prefix As String) As String

      Dim format As String = Me.Name

      If Not prefix = String.Empty Then
        format = String.Format("[{0}].[{1}]", prefix, format)
      End If

      If Me.DataType = ScriptDB.ColumnTypes.Text Then

        ' Case
        Select Case Me.CaseType
          Case Enums.CaseType.Lower
            format = String.Format("LOWER({0})", format)
          Case Enums.CaseType.Proper
            format = String.Format("dbo.udfsys_propercase({0})", format)
          Case Enums.CaseType.Upper
            format = String.Format("UPPER({0})", format)
        End Select

        ' Trim type
        Select Case Me.TrimType
          Case Enums.TrimType.Both
            format = String.Format("LTRIM(RTRIM({0}))", format)
          Case Enums.TrimType.Left
            format = String.Format("LTRIM({0})", format)
          Case Enums.TrimType.Right
            format = String.Format("RTRIM({0})", format)
        End Select

      End If

      Return format
    End Function

    Public ReadOnly Property SafeReturnType As String
      Get

        Select Case CInt(Me.DataType)
          Case ScriptDB.ColumnTypes.Text, ScriptDB.ColumnTypes.WorkingPattern, ScriptDB.ColumnTypes.Link
            Return "''"
          Case ScriptDB.ColumnTypes.Integer, ScriptDB.ColumnTypes.Numeric, ScriptDB.ColumnTypes.Logic
            Return "0"
          Case ScriptDB.ColumnTypes.Date
            Return "NULL"
          Case Else
            Return "0"
        End Select
      End Get
    End Property

  End Class
End Namespace

