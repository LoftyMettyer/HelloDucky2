Option Strict On
Option Explicit On

Imports System.Drawing

Namespace Code.Captcha
  Friend Class CaptchaImage

 		Const HMargin As Integer = 5
		Const VMargin As Integer = 3
    Const ValidChars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

    Friend Shared Function CreateBitmap(encode as string) As Bitmap
    
		  Dim letter As New List(Of Letter)()
		  Dim TotalWidth As Integer = 0
		  Dim MaxHeight As Integer = 0
		  For Each c As Char In encode
			  Dim ltr = New Letter(c)
			  letter.Add(ltr)
			  Dim space As Integer = (New Random()).[Next](5) + 1
			  ltr.space = space
			  Threading.Thread.Sleep(10)
			  TotalWidth += ltr.LetterSize.Width + space
			  If MaxHeight < ltr.LetterSize.Height Then
				  MaxHeight = ltr.LetterSize.Height
			  End If
			  Threading.Thread.Sleep(10)
		  Next

		  Dim bmp As New Bitmap(TotalWidth + HMargin, MaxHeight + VMargin)
		  Dim Grph = Graphics.FromImage(bmp)
		  Grph.FillRectangle(New SolidBrush(Color.Lavender), 0, 0, bmp.Width, bmp.Height)
	    MergeBackground(bmp)
		  Grph.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
		  Grph.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
		  Dim xPos As Integer = HMargin
		  For Each ltr As Letter In letter
			  Grph.DrawString(ltr.letter.ToString(), ltr.font, New SolidBrush(Color.Navy), xPos, VMargin)
			  xPos += ltr.LetterSize.Width + ltr.space
		  Next

      Return bmp

    End Function

    Private Shared Sub MergeBackground(ByRef bmp As Bitmap)

      Dim grp = Graphics.FromImage(bmp)
      Dim r = new Random()

      dim filename = HttpContext.Current.Server.MapPath("~/content/images/captcha/")
      filename &= String.Format("captcha{0}.jpg", r.Next(8))

      Dim background As Image = Image.FromFile(filename)
      grp.DrawImage(background, New Rectangle(0, 0, bmp.Width, bmp.Height))

    End Sub

    Friend Shared Function RandomString(size As Integer) As String

      Dim res As New StringBuilder()
      Dim rnd As New Random()
      While 0 < Math.Max(Threading.Interlocked.Decrement(size), size + 1)
	      res.Append(ValidChars(rnd.[Next](ValidChars.Length)))
      End While
      Return res.ToString()

    End Function

  End Class




End NameSpace