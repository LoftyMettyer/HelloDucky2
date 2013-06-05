Namespace ScriptDB

  <HideModuleName()> _
  Public Module StringFunctions

    Public Sub Beautify(ByRef Input As String)

      ' Put correct indentation
      Input.Replace(vbNewLine, vbNewLine & Space(4))

      ' Remove blank lines
      Input.Replace(vbNewLine & vbNewLine, vbNewLine)


    End Sub

    Public Sub FormatDeclarations(ByRef Input As String)



    End Sub

  End Module


End Namespace
