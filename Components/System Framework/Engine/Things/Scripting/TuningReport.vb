﻿Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)>
Public Class TuningReport
  Implements IErrors

  Public ReadOnly Expressions As ICollection(Of Column)

  Public Sub New()
    Expressions = New Collection(Of Column)
  End Sub

  Public Sub OutputToFile(ByVal fileName As String) Implements IErrors.OutputToFile

    Dim objWriter As IO.StreamWriter
    Dim objThing As Base
    Dim objTable As Table
    Dim objColumn As Column
    Dim sMessage As String

    IO.File.Delete(fileName)
    objWriter = IO.File.AppendText(fileName)

    objWriter.Write(String.Format("{0}{0}{1}{0}EXPRESSION USAGE{0}{1}{0}{0}", vbNewLine, "-------------------"))

    For Each objTable In Tables
      For Each objColumn In objTable.Columns
        If objColumn.IsCalculated And objColumn.State <> DataRowState.Deleted Then
          sMessage = String.Format("({0}) {1}.{2}     | Expression = {3} - ({4})", objColumn.Tuning.Usage.ToString.PadLeft(3) _
                , objColumn.Table.Name, objColumn.Name, objColumn.Calculation.Name _
                , If(objColumn.Calculation.IsComplex, "Complex ", "Simple")) & vbNewLine
        Else
          sMessage = String.Format("({0}) {1}.{2}", objColumn.Tuning.Usage.ToString.PadLeft(3) _
                , objColumn.Table.Name, objColumn.Name) & vbNewLine
        End If

        objWriter.Write(sMessage)
      Next
    Next


    'For Each objThing In Expressions
    '  objColumn = CType(objThing, Column)
    '  sMessage = String.Format("({0}) {1}.{2}{3}", objColumn.Tuning.Usage.ToString.PadLeft(3) _
    '        , objColumn.Table.Name, objThing.Name _
    '        , If(objColumn.Calculation.IsComplex, " (COMPLEX) ", "")) & vbNewLine
    '  objWriter.Write(sMessage)
    'Next

    objWriter.Write(String.Format("{0}{0}{1}{0}FUNCTION USAGE{0}{1}{0}{0}", vbNewLine, "-------------------"))
    For Each objThing In Functions
      sMessage = String.Format("({0}) - {1}", objThing.Tuning.Usage.ToString.PadLeft(5) _
            , objThing.Name) & vbNewLine
      objWriter.Write(sMessage)
    Next

    objWriter.Close()

  End Sub

  Public ReadOnly Property ErrorCount As Integer Implements IErrors.ErrorCount
    Get
      Return 0
    End Get
  End Property

  Public ReadOnly Property IsCatastrophic As Boolean Implements IErrors.IsCatastrophic
    Get
      Return False
    End Get
  End Property

  Public Sub Show() Implements IErrors.Show

  End Sub
End Class
