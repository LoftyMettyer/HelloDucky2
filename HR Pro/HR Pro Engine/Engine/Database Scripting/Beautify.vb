Namespace ScriptDB

  Public Module Beautify

    Public Sub Cleanwhitespace(ByRef Input As String)

      ' Put correct indentation
      Input.Replace(vbNewLine, vbNewLine & Space(4))

      ' Remove blank lines
      Input.Replace(vbNewLine & vbNewLine, vbNewLine)


    End Sub

    Public Sub FormatDeclarations(ByRef Input As String)
    End Sub

    Public Function MakeSingleLine(ByVal Input As String) As String

      Dim sReturn As String = Input

      sReturn.Replace(vbNewLine, "")
      sReturn.Replace(" ", "")

      Return sReturn

    End Function

  End Module


End Namespace
