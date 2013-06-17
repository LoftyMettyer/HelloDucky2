Imports Utilities
Imports System.Data.SqlClient

Public Class Forms

  Public Shared Sub LoadControlData(page As Page, formId As Integer)

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformelements WHERE form = " & formId, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()

        Dim control As Control = page.Master.FindControl("mainCPH").FindControl(CStr(dr("Name")))
        If control Is Nothing Then control = page.Master.FindControl("footerCPH").FindControl(CStr(dr("Name")))
        Dim general As New General

        Select Case CInt(dr("Type"))

          Case 0 ' Button

            CType(control, Image).ImageUrl = Picture.GetUrl(NullSafeInteger(dr("PictureID")))

            ' Footer text
            If NullSafeString(dr("Caption")).Length > 0 Then
              CType(control.Parent.FindControl(CStr(dr("Name")) & "_label"), Label).Text = NullSafeString(dr("caption"))
            End If

          Case 2 ' Label

            With CType(control, Label)
              .Text = NullSafeString(dr("caption"))
              .Style("word-wrap") = "break-word" 'TODO move to css
              .Style.Add("color", general.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
              .Style.Add("font-family", NullSafeString(dr("FontName")))
              .Style.Add("font-size", NullSafeString(dr("FontSize")) & "pt")
              .Style.Add("font-weight", If(NullSafeBoolean(dr("FontBold")), "bold", "normal"))
              .Style.Add("font-style", If(NullSafeBoolean(dr("FontItalic")), "italic", "normal"))
            End With

          Case 3 ' Input value - character

            With CType(control, TextBox)
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
