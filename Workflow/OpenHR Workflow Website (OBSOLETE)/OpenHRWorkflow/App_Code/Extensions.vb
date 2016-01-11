Imports System.Runtime.CompilerServices
Imports System.Data
Imports System.Drawing
Imports Utilities

Public Module Extensions
   <Extension()>
   Public Sub Apply(value As FontInfo, formItem As FormItem, Optional namePrefix As String = Nothing)
      'TODO PG NOW namePrefix code
      With value
         .Name = formItem.FontName
         .Size = ToPointFontUnit(formItem.FontSize)
         .Bold = formItem.FontBold
         .Italic = formItem.FontItalic
         .Strikeout = formItem.FontStrikeThru
         .Underline = formItem.FontUnderline
      End With
   End Sub

   <Extension()>
   Public Sub ApplyFont(value As WebControl, formItem As FormItem, Optional namePrefix As String = Nothing)
      value.Font.Apply(formItem, namePrefix)
   End Sub

   <Extension()>
   Public Sub ApplyFont(value As CssStyleCollection, formItem As FormItem)

      Dim css As String = ""

      If formItem.FontItalic Then css += "italic "
      If formItem.FontBold Then css += "bold "
      css += ToPoint(formItem.FontSize).ToString & "pt " & formItem.FontName

      value.Add("font", css)

      Dim decoration As String = If(formItem.FontStrikeThru, "line-through ", "") &
                                 If(formItem.FontUnderline, "underline ", "")

      If decoration.Length > 0 Then
         value.Add("text-decoration", decoration.TrimEnd)
      End If
   End Sub

   Public Function GetFontCss(formItem As FormItem) As String

      Dim css As String = "font: "

      If formItem.FontItalic Then css += "italic "
      If formItem.FontBold Then css += "bold "
      css += ToPoint(formItem.FontSize).ToString & "pt " & formItem.FontName & ";"

      Dim decoration As String = If(formItem.FontStrikeThru, "line-through ", "") &
                                 If(formItem.FontUnderline, "underline ", "")

      If decoration.Length > 0 Then
         css += " text-decoration: " & decoration.TrimEnd & ";"
      End If

      Return css
   End Function

   <Extension()>
   Public Sub ApplyLocation(value As CssStyleCollection, formItem As FormItem)

      value("position") = "absolute"
      value("top") = Unit.Pixel(formItem.Top).ToString
      value("left") = Unit.Pixel(formItem.Left).ToString
   End Sub

   <Extension()>
   Public Sub ApplyLocation(value As WebControl, formItem As FormItem)
      value.Style.ApplyLocation(formItem)
   End Sub

   <Extension()>
   Public Sub ApplySize(value As CssStyleCollection, formItem As FormItem, Optional widthAdjustment As Integer = 0,
                        Optional heightAdjustment As Integer = 0)

      value("Height") = Unit.Pixel(formItem.Height + heightAdjustment).ToString
      value("Width") = Unit.Pixel(formItem.Width + widthAdjustment).ToString
   End Sub

   <Extension()>
   Public Sub ApplySize(value As WebControl, formItem As FormItem, Optional widthAdjustment As Integer = 0,
                        Optional heightAdjustment As Integer = 0)

      value.Height = Unit.Pixel(formItem.Height + heightAdjustment)
      value.Width = Unit.Pixel(formItem.Width + widthAdjustment)
   End Sub

   <Extension()>
   Public Sub ApplyColor(value As WebControl, formItem As FormItem, Optional canBeTranparent As Boolean = False)

      value.ForeColor = General.GetColour(AdjustedForeColor(formItem.ForeColor))

      If canBeTranparent AndAlso formItem.BackStyle = 0 Then
         value.BackColor = Color.Transparent
      Else
         value.BackColor = General.GetColour(AdjustedBackColor(formItem.BackColor))
      End If
   End Sub

   'TODO default colors
   Private Function AdjustedForeColor(color As Integer) As Integer
      Select Case color
         'Case 6697779 '#333366
         '   Return 3355443 '#333333
         Case Else
            Return color
      End Select
      Return color
   End Function

   Private Function AdjustedBackColor(color As Integer) As Integer
      Select Case color
         'Case 15988214
         '   Return 16777215 '#FFFFFF
         Case Else
            Return color
      End Select
      Return color
   End Function

   <Extension()>
   Public Sub ApplyColor(value As CssStyleCollection, formItem As FormItem, Optional canBeTransparent As Boolean = False)

      value("color") = General.GetHtmlColour(AdjustedForeColor(formItem.ForeColor))

      If canBeTransparent AndAlso formItem.BackStyle = 0 Then
         value("background-color") = "transparent"
      Else
         value("background-color") = General.GetHtmlColour(AdjustedBackColor(formItem.BackColor))
      End If
   End Sub

   Public Function GetColorCss(formItem As FormItem, Optional canBeTransparent As Boolean = False) As String

      Dim css As String = "color: " & General.GetHtmlColour(AdjustedForeColor(formItem.ForeColor)) & ";"

      If canBeTransparent AndAlso formItem.BackStyle = 0 Then
         css += " background-color: transparent;"
      Else
         css += " background-color: " & General.GetHtmlColour(AdjustedBackColor(formItem.BackColor)) & ";"
      End If

      Return css
   End Function

   <Extension()>
   Public Sub ApplyBorder(value As WebControl, adjustSize As Boolean, Optional adjustSizeAmount As Integer = - 4)

      value.BorderStyle = BorderStyle.Solid
      value.BorderColor = ColorTranslator.FromHtml("#999")
      value.BorderWidth = 1

      If adjustSize Then
         value.Width = CInt(value.Width.Value) + adjustSizeAmount
         value.Height = CInt(value.Height.Value) + adjustSizeAmount
      End If
   End Sub
End Module
