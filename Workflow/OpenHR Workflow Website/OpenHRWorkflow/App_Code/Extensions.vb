Imports System.Runtime.CompilerServices
Imports System.Data
Imports System.Drawing
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
  Public Sub ApplyFont(value As WebControl, dataReader As IDataReader, Optional namePrefix As String = Nothing)
    value.Font.Apply(dataReader, namePrefix)
  End Sub

  <Extension()> _
  Public Sub ApplyFont(value As CssStyleCollection, dataReader As IDataReader)

    Dim css As String = ""

    If NullSafeBoolean(dataReader("FontItalic")) Then css += "italic "
    If NullSafeBoolean(dataReader("FontBold")) Then css += "bold "
    css += ToPoint(NullSafeInteger(dataReader("FontSize"))).ToString & "pt " & NullSafeString(dataReader("FontName"))

    value.Add("font", css)

    Dim decoration As String = If(NullSafeBoolean(dataReader("FontStrikeThru")), "line-through ", "") & _
                               If(NullSafeBoolean(dataReader("FontUnderline")), "underline ", "")

    If decoration.Length > 0 Then
      value.Add("text-decoration", decoration.TrimEnd)
    End If

  End Sub

  Public Function GetFontCss(dataReader As IDataReader) As String

    Dim css As String = "font: "

    If NullSafeBoolean(dataReader("FontItalic")) Then css += "italic "
    If NullSafeBoolean(dataReader("FontBold")) Then css += "bold "
    css += ToPoint(NullSafeInteger(dataReader("FontSize"))).ToString & "pt " & NullSafeString(dataReader("FontName")) & ";"

    Dim decoration As String = If(NullSafeBoolean(dataReader("FontStrikeThru")), "line-through ", "") & _
                               If(NullSafeBoolean(dataReader("FontUnderline")), "underline ", "")

    If decoration.Length > 0 Then
      css += " text-decoration: " & decoration.TrimEnd & ";"
    End If

    Return css

  End Function

  <Extension()> _
  Public Sub ApplyLocation(value As CssStyleCollection, dataReader As IDataReader)

    value("position") = "absolute"
    value("top") = Unit.Pixel(NullSafeInteger(dataReader("TopCoord"))).ToString
    value("left") = Unit.Pixel(NullSafeInteger(dataReader("LeftCoord"))).ToString

  End Sub

  <Extension()> _
  Public Sub ApplyLocation(value As WebControl, dataReader As IDataReader)
    value.Style.ApplyLocation(dataReader)
  End Sub

  <Extension()> _
  Public Sub ApplySize(value As CssStyleCollection, dataReader As IDataReader, Optional widthAdjustment As Integer = 0, Optional heightAdjustment As Integer = 0)

    value("Height") = Unit.Pixel(NullSafeInteger(dataReader("Height")) + heightAdjustment).ToString
    value("Width") = Unit.Pixel(NullSafeInteger(dataReader("Width")) + widthAdjustment).ToString

  End Sub

  <Extension()> _
  Public Sub ApplySize(value As WebControl, dataReader As IDataReader, Optional widthAdjustment As Integer = 0, Optional heightAdjustment As Integer = 0)

    value.Height = Unit.Pixel(NullSafeInteger(dataReader("Height")) + heightAdjustment)
    value.Width = Unit.Pixel(NullSafeInteger(dataReader("Width")) + widthAdjustment)

  End Sub

  <Extension()> _
  Public Sub ApplyColor(value As WebControl, dataReader As IDataReader, Optional canBeTranparent As Boolean = False)

    value.ForeColor = General.GetColour(AdjustedForeColor(NullSafeInteger(dataReader("ForeColor"))))

    If canBeTranparent AndAlso NullSafeInteger(dataReader("BackStyle")) = 0 Then
      value.BackColor = Color.Transparent
    Else
      value.BackColor = General.GetColour(AdjustedBackColor(NullSafeInteger(dataReader("BackColor"))))
    End If

  End Sub

  Private Function AdjustedForeColor(color As Integer) As Integer
    'TODO PG NOW
    Select Case color
      Case 6697779 '#333366
        Return 3355443 '#333333
      Case Else
        Return color
    End Select
    Return color
  End Function

  Private Function AdjustedBackColor(color As Integer) As Integer
    'TODO PG NOW
    Select Case color
      Case 15988214
        Return 16777215 '#FFFFFF
      Case Else
        Return color
    End Select
    Return color
  End Function

  <Extension()> _
  Public Sub ApplyColor(value As CssStyleCollection, dataReader As IDataReader, Optional canBeTransparent As Boolean = False)

    value("color") = General.GetHtmlColour(AdjustedForeColor(NullSafeInteger(dataReader("ForeColor"))))

    If canBeTransparent AndAlso NullSafeInteger(dataReader("BackStyle")) = 0 Then
      value("background-color") = "transparent"
    Else
      value("background-color") = General.GetHtmlColour(AdjustedBackColor(NullSafeInteger(dataReader("BackColor"))))
    End If

  End Sub

  Public Function GetColorCss(datareader As IDataReader, Optional canBeTransparent As Boolean = False) As String

    Dim css As String = "color: " & General.GetHtmlColour(AdjustedForeColor(NullSafeInteger(datareader("ForeColor")))) & ";"

    If canBeTransparent AndAlso NullSafeInteger(datareader("BackStyle")) = 0 Then
      css += " background-color: transparent;"
    Else
      css += " background-color: " & General.GetHtmlColour(AdjustedBackColor(NullSafeInteger(datareader("BackColor")))) & ";"
    End If

    Return css

  End Function

  <Extension()> _
  Public Sub ApplyBorder(value As WebControl, adjustSize As Boolean, Optional adjustSizeAmount As Integer = -4)

    value.BorderStyle = BorderStyle.Solid
    value.BorderColor = ColorTranslator.FromHtml("#999")
    value.BorderWidth = Unit.Pixel(1)

    If adjustSize Then
      value.Width = Unit.Pixel(CInt(value.Width.Value) + adjustSizeAmount)
      value.Height = Unit.Pixel(CInt(value.Height.Value) + adjustSizeAmount)
    End If

  End Sub

End Module
