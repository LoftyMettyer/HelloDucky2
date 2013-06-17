Imports Utilities
Imports System.Data.SqlClient

Public Class Forms

  Public Shared Sub LoadControlData(page As Page, formId As Integer)

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("SELECT mfe.*, p.Name AS PictureName FROM tbsys_mobileformelements mfe LEFT JOIN ASRSysPictures p ON mfe.PictureID = p.PictureID WHERE form = " & formId, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()

        Dim control As Control = page.Master.FindControl("mainCPH").FindControl(CStr(dr("Name")))
        If control Is Nothing Then control = page.Master.FindControl("footerCPH").FindControl(CStr(dr("Name")))
        Dim general As New General

        Select Case CInt(dr("Type"))

          Case 0 ' Button

            With CType(control, ImageButton)
              .ImageUrl = "~/" & Picture.LoadPicture(NullSafeInteger(dr("PictureID")))
              .Font.Name = NullSafeString(dr("FontName"))
              .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
              .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
              .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
            End With

            ' Footer text
            If NullSafeString(dr("Caption")).Length > 0 Then

              With CType(control.Parent.FindControl(CStr(dr("Name")) & "_label"), HtmlGenericControl)
                .InnerText = NullSafeString(dr("caption"))
                'TODO this can all be css
                .Style("word-wrap") = "break-word"
                .Style("overflow") = "auto"
                .Style.Add("background-color", "Transparent")
                .Style.Add("font-family", "Verdana")
                .Style.Add("font-size", "6pt")
                .Style.Add("font-weight", "normal")
                .Style.Add("font-style", "normal")
              End With
            End If

          Case 2 ' Label

            With CType(control, HtmlGenericControl)
              .InnerText = NullSafeString(dr("caption"))
              .Style("word-wrap") = "break-word"
              .Style.Add("color", general.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
              .Style.Add("font-family", NullSafeString(dr("FontName")))
              .Style.Add("font-size", NullSafeString(dr("FontSize")) & "pt")
              .Style.Add("font-weight", If(NullSafeBoolean(dr("FontBold")), "bold", "normal"))
              .Style.Add("font-style", If(NullSafeBoolean(dr("FontItalic")), "italic", "normal"))
            End With

          Case 3 ' Input value - character

            With CType(control, HtmlInputText)
              .Style.Add("color", general.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
              .Style.Add("font-family", NullSafeString(dr("FontName")))
              .Style.Add("font-size", NullSafeString(dr("FontSize")) & "pt")
              .Style.Add("font-weight", If(NullSafeBoolean(dr("FontBold")), "bold", "normal"))
              .Style.Add("font-style", If(NullSafeBoolean(dr("FontItalic")), "italic", "normal"))
            End With

        End Select
      End While

    End Using

  End Sub

End Class
