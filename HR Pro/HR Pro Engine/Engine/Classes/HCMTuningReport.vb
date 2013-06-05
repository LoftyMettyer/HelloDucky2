Option Strict On

Imports System.Runtime.InteropServices

Namespace Tuning

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class Report

    Public Expressions As New Things.Collection

    Public Sub OutputToFile(ByRef FileName As String) ' Implements Interfaces.iErrors.OutputToFile

      Dim objWriter As System.IO.StreamWriter
      Dim objThing As Things.Base
      Dim objColumn As Things.Column
      Dim sMessage As String

      System.IO.File.Delete(FileName)
      objWriter = System.IO.File.AppendText(FileName)

      For Each objThing In Expressions
        objColumn = CType(objThing, Things.Column)
        sMessage = String.Format("({0}) {1}.{2}{3}", objColumn.Tuning.Usage.ToString.PadLeft(3) _
              , objColumn.Table.Name, objThing.Name _
              , IIf(objColumn.Calculation.IsComplex, " (COMPLEX) ", "")) & vbNewLine
        objWriter.Write(sMessage)
      Next

      objWriter.Close()

    End Sub

  End Class

End Namespace
