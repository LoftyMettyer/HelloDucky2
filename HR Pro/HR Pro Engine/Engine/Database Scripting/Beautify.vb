Namespace ScriptDB

  Public Module Beautify

    Public Function CleanWhitespace(ByVal value As String) As String

      value = value.Replace(vbNewLine, vbNewLine & Space(4))
      value = value.Replace(vbNewLine & vbNewLine, vbNewLine)

      Return value

    End Function

    Public Function MakeSingleLine(ByVal value As String) As String

      value = Replace(value, Chr(13), " ")
      value = Replace(value, Chr(10), "")
      value = Replace(value, vbTab, " ")

      Return value

    End Function

  End Module

End Namespace
