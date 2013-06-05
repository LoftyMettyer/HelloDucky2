Imports System.Runtime.InteropServices
Imports SystemFramework.Things

Namespace Tuning

  <ClassInterface(ClassInterfaceType.None)>
  Public Class Report
    Implements COMInterfaces.IErrors

    Public Expressions As ICollection(Of Column)

    Public Sub New()
      Expressions = New Collection(Of Column)
    End Sub

    Public Sub OutputToFile(ByVal FileName As String) Implements COMInterfaces.IErrors.OutputToFile

      Dim objWriter As System.IO.StreamWriter
      Dim objThing As Base
      Dim objTable As Table
      Dim objColumn As Column
      Dim sMessage As String

      System.IO.File.Delete(FileName)
      objWriter = System.IO.File.AppendText(FileName)

      objWriter.Write(String.Format("{0}{0}{1}{0}EXPRESSION USAGE{0}{1}{0}{0}", vbNewLine, "-------------------"))

      For Each objTable In Globals.Tables
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
      For Each objThing In Globals.Functions
        sMessage = String.Format("({0}) - {1}", objThing.Tuning.Usage.ToString.PadLeft(5) _
              , objThing.Name) & vbNewLine
        objWriter.Write(sMessage)
      Next

      objWriter.Close()

    End Sub

    Public ReadOnly Property ErrorCount As Integer Implements COMInterfaces.IErrors.ErrorCount
      Get
        Return 0
      End Get
    End Property

    Public ReadOnly Property IsCatastrophic As Boolean Implements COMInterfaces.IErrors.IsCatastrophic
      Get
        Return False
      End Get
    End Property

    Public Sub Show() Implements COMInterfaces.IErrors.Show

    End Sub
  End Class

End Namespace
