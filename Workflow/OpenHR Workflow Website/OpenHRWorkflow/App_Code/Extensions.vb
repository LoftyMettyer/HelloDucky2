Imports System.Runtime.CompilerServices
Imports System.Data
Imports Utilities

Public Module Extensions

  <Extension()> _
  Public Sub Apply(value As FontInfo, dataReader As IDataReader, Optional namePrefix As String = Nothing)

    With value
      .Name = NullSafeString(dataReader(namePrefix & "FontName"))
      .Size = ToPointFontUnit(NullSafeInteger(dataReader(namePrefix & "FontSize")))
      .Bold = NullSafeBoolean(dataReader(namePrefix & "FontBold"))
      .Italic = NullSafeBoolean(dataReader(namePrefix & "FontItalic"))
      .Strikeout = NullSafeBoolean(dataReader(namePrefix & "FontStrikeThru"))
      .Underline = NullSafeBoolean(dataReader(namePrefix & "FontUnderline"))
    End With

  End Sub

  <Extension()> _
  Public Sub ApplyFont(value As CssStyleCollection, dataReader As IDataReader)

    Dim decoration As String

    decoration = If(NullSafeBoolean(dataReader("FontStrikeThru")), " line-through", "") & _
                 If(NullSafeBoolean(dataReader("FontUnderline")), " underline", "")

    If decoration.Length = 0 Then
      decoration = "none"
    End If

    value.Add("font-family", NullSafeString(dataReader("FontName")).ToString)
    value.Add("font-size", ToPoint(NullSafeInteger(dataReader("FontSize"))).ToString & "pt")
    value.Add("font-weight", If(NullSafeBoolean(dataReader("FontBold")), "bold", "normal"))
    value.Add("font-style", If(NullSafeBoolean(dataReader("FontItalic")), "italic", "normal"))
    value.Add("text-decoration", decoration)

  End Sub

End Module
