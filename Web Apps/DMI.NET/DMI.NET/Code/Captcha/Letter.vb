Option Strict On
Option Explicit On

Imports System.Drawing

Namespace Code.Captcha

  Public Class Letter
    Private ReadOnly ValidFonts As String() = {"Comic Sans MS", "Courier New", "Georgia", "Impact", "Times New Roman", "Tahoma", "Verdana"}
    Public Sub New(c As Char)
      Dim rnd As New Random()

      Font = New Font(ValidFonts(rnd.[Next](ValidFonts.Count() - 1)), rnd.[Next](20) + 20 _
                , CType(iif(rnd.NextDouble() > 0.5, FontStyle.Bold, FontStyle.Regular), FontStyle) _
                  Or CType(iif(rnd.NextDouble() > 0.5, FontStyle.Italic, FontStyle.Regular), FontStyle) _
                , GraphicsUnit.Point)

      letter = c
    End Sub

    Public Property Font As Font
    Public Property Letter As Char
    Public Property Space As Integer

    Public ReadOnly Property LetterSize() As Size
      Get
        Dim Bmp = New Bitmap(1, 1)
        Dim Grph = Graphics.FromImage(Bmp)
        Return Grph.MeasureString(letter.ToString(), font).ToSize()
      End Get
    End Property

  End Class

End NameSpace