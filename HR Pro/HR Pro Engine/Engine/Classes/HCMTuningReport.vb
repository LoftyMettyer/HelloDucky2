Option Strict On

Imports System.Runtime.InteropServices

Namespace Tuning

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class Report
    Implements COMInterfaces.iErrors

    Public Expressions As New Things.Collections.Generic

    Public Sub OutputToFile(ByRef FileName As String) Implements COMInterfaces.iErrors.OutputToFile

      Dim objWriter As System.IO.StreamWriter
      Dim objThing As Things.Base
      Dim objTable As Things.Table
      Dim objColumn As Things.Column
      Dim sMessage As String

      System.IO.File.Delete(FileName)
      objWriter = System.IO.File.AppendText(FileName)

      objWriter.Write(String.Format("{0}{0}{1}{0}EXPRESSION USAGE{0}{1}{0}{0}", vbNewLine, "-------------------"))

      For Each objTable In Globals.Things
        For Each objColumn In objTable.Objects(Things.Type.Column)
          If objColumn.IsCalculated Then
            sMessage = String.Format("({0}) {1}.{2}     | Expression = {3} - ({4})", objColumn.Tuning.Usage.ToString.PadLeft(3) _
                  , objColumn.Table.Name, objColumn.Name, objColumn.Calculation.Name _
                  , IIf(objColumn.Calculation.IsComplex, "Complex ", "Simple")) & vbNewLine
          Else
            sMessage = String.Format("({0}) {1}.{2}", objColumn.Tuning.Usage.ToString.PadLeft(3) _
                  , objColumn.Table.Name, objColumn.Name) & vbNewLine
          End If

          objWriter.Write(sMessage)
        Next
      Next


      'For Each objThing In Expressions
      '  objColumn = CType(objThing, Things.Column)
      '  sMessage = String.Format("({0}) {1}.{2}{3}", objColumn.Tuning.Usage.ToString.PadLeft(3) _
      '        , objColumn.Table.Name, objThing.Name _
      '        , IIf(objColumn.Calculation.IsComplex, " (COMPLEX) ", "")) & vbNewLine
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

    Public ReadOnly Property ErrorCount As Integer Implements COMInterfaces.iErrors.ErrorCount
      Get
        Return 0
      End Get
    End Property

    Public ReadOnly Property IsCatastrophic As Boolean Implements COMInterfaces.iErrors.IsCatastrophic
      Get
        Return False
      End Get
    End Property

    Public Sub Show() Implements COMInterfaces.iErrors.Show

    End Sub
  End Class

End Namespace
