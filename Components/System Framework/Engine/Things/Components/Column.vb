﻿Imports System.Runtime.InteropServices
Imports SystemFramework.Enums

<ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
Public Class Column
  Inherits Base

  Public Property Table As Table

  Public Property CalcId As Integer
  Public Property Calculation As Expression

  Public Property DataType As ColumnTypes
  Public Property Size As Integer
  Public Property Decimals As Integer

  ' Options
  Public Property Audit As Boolean
  Public Property Multiline As Boolean
  Public Property CalculateIfEmpty As Boolean
  Public Property IsReadOnly As Boolean

  ' Formatting
  Public Property CaseType As CaseType
  Public Property TrimType As TrimType
  Public Property Alignment As AlignType
  Public Property Mandatory As Boolean

  Public Property UniqueType As UniqueCheckScope

  Public Property DefaultCalcId As Integer
  Public Property DefaultCalculation As Expression
  Public Property DefaultValue As String

  ' Declaration syntax for a column type
  Public ReadOnly Property DataTypeSyntax As String
    Get

      Dim sqlType As String = String.Empty

      Select Case DataType
        Case ColumnTypes.Text
          If Multiline Or Size > 8000 Then
            sqlType = "varchar(MAX)"
          Else
            sqlType = String.Format("varchar({0})", Size)
          End If

        Case ColumnTypes.Integer
          sqlType = String.Format("integer")

        Case ColumnTypes.Numeric
          sqlType = String.Format("numeric({0},{1})", Size, Decimals)

        Case ColumnTypes.Date
          sqlType = "datetime"

        Case ColumnTypes.Logic
          sqlType = "bit"

        Case ColumnTypes.WorkingPattern
          sqlType = "varchar(14)"

        Case ColumnTypes.Link
          sqlType = "varchar(255)"

        Case ColumnTypes.Photograph
          sqlType = "varchar(255)"

        Case ColumnTypes.Binary
          sqlType = "varbinary(MAX)"

      End Select

      Return sqlType

    End Get

  End Property

  Public ReadOnly Property IsCalculated As Boolean
    Get
      Return CalcId > 0
    End Get
  End Property

  Public Function ApplyFormatting(ByVal prefix As String) As String

    Dim format As String = Name

    If Not prefix = String.Empty Then
      format = String.Format("[{0}].[{1}]", prefix, format)
    End If

    If DataType = ColumnTypes.Text Then

      ' Case
      Select Case CaseType
        Case CaseType.Lower
          format = String.Format("LOWER({0})", format)
        Case CaseType.Proper
          format = String.Format("dbo.udfsys_propercase({0})", format)
        Case CaseType.Upper
          format = String.Format("UPPER({0})", format)
      End Select

      ' Trim type
      Select Case TrimType
        Case TrimType.Both
          format = String.Format("LTRIM(RTRIM({0}))", format)
        Case TrimType.Left
          format = String.Format("LTRIM({0})", format)
        Case TrimType.Right
          format = String.Format("RTRIM({0})", format)
      End Select

    End If

    Return format
  End Function

  Public ReadOnly Property SafeReturnType As String
    Get

      Select Case CInt(DataType)
        Case ColumnTypes.Text, ColumnTypes.WorkingPattern, ColumnTypes.Link
          Return "''"
        Case ColumnTypes.Integer, ColumnTypes.Numeric, ColumnTypes.Logic
          Return "0"
        Case ColumnTypes.Date
          Return "NULL"
        Case Else
          Return "0"
      End Select
    End Get
  End Property

  Public ReadOnly Property NullCheckValue As String
    Get

      Select Case CInt(DataType)
        Case ColumnTypes.Text, ColumnTypes.WorkingPattern, ColumnTypes.Link
          Return "''"
        Case ColumnTypes.Integer, ColumnTypes.Numeric, ColumnTypes.Logic
          Return "0"
        Case ColumnTypes.Date
          Return "'1899-12-31'"
        Case Else
          Return "0"
      End Select
    End Get
  End Property




  Public ReadOnly Property ComponentReturnType As ComponentValueTypes
    Get

      Select Case CInt(DataType)

        Case ColumnTypes.Text, ColumnTypes.WorkingPattern, ColumnTypes.Link
          Return ComponentValueTypes.String

        Case ColumnTypes.Integer, ColumnTypes.Numeric
          Return ComponentValueTypes.Numeric

        Case ColumnTypes.Logic
          Return ComponentValueTypes.Logic

        Case ColumnTypes.Date
          Return ComponentValueTypes.Date

        Case Else
          Return ComponentValueTypes.Numeric
      End Select

    End Get
  End Property

End Class

