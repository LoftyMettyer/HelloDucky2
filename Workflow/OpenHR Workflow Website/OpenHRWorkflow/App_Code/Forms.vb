Imports Utilities
Imports System.Data.SqlClient

Public Class Forms
  Public Shared Sub RedirectIfNotLicensed()

    Dim db As New Database
    If Not db.IsMobileLicensed() Then
      HttpContext.Current.Session("message") = "You are not licensed for the OpenHR Mobile module. Please contact your Advanced Business Solutions Account Manager for details"
      HttpContext.Current.Response.Redirect("~/Message.aspx")
    End If

  End Sub

  Public Shared Sub RedirectIfDbLocked()

    Dim db As New Database
    If db.IsSystemLocked() Then
      HttpContext.Current.Session("message") = "The system is currently being modified. Please retry again shortly."
      HttpContext.Current.Response.Redirect("~/Message.aspx")
    End If

  End Sub

  Public Shared Sub RedirectToNotConfigured()

    Dim message As String = ""

    If Configuration.WorkflowUrl.Length = 0 Then message += "Workflow url is not defined, "
    If Configuration.MobileKey.Length = 0 Then message += "Mobile key is not defined, "

    If message.Length > 0 Then
      HttpContext.Current.Session("message") = "The system is not configured correctly, " & message.TrimEnd(","c, " "c) & ". Please contact your system administrator."
      HttpContext.Current.Response.Redirect("~/Message.aspx")
    End If

  End Sub

  Public Shared Sub RedirectToHomeIfAuthentcated()

    'Go to the home page if already logged in
    If HttpContext.Current.Request.IsAuthenticated Then
      HttpContext.Current.Response.Redirect("~/Home.aspx")
    End If

  End Sub

  Public Shared Sub LoadControlData(page As Page, formId As Integer)

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformelements WHERE form = " & formId, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()

        Dim control As Control = page.Master.FindControl("mainCPH").FindControl(CStr(dr("Name")))
        If control Is Nothing Then control = page.Master.FindControl("footerCPH").FindControl(CStr(dr("Name")))

        Select Case CInt(dr("Type"))

          Case 0 ' Button

            CType(control.Controls(0), Image).ImageUrl = Picture.GetUrl(NullSafeInteger(dr("PictureID")))
            CType(control.Controls(1), Label).Text = NullSafeString(dr("caption"))

          Case 2 ' Label

            With CType(control, Label)
              .Text = NullSafeString(dr("caption"))
              .Font.Name = NullSafeString(dr("FontName"))
              .Font.Size = New FontUnit(NullSafeSingle(dr("FontSize")))
              .Font.Bold = NullSafeBoolean(dr("FontBold"))
              .Font.Italic = NullSafeBoolean(dr("FontItalic"))
              .Font.Underline = NullSafeBoolean(dr("FontUnderline"))
              .Font.Strikeout = NullSafeBoolean(dr("FontStrikeout"))

              .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
              .Style("word-wrap") = "break-word"
            End With

          Case 3 ' Input value - character

            With CType(control, TextBox)
              .Font.Name = NullSafeString(dr("FontName"))
              .Font.Size = New FontUnit(NullSafeSingle(dr("FontSize")))
              .Font.Bold = NullSafeBoolean(dr("FontBold"))
              .Font.Italic = NullSafeBoolean(dr("FontItalic"))
              .Font.Underline = NullSafeBoolean(dr("FontUnderline"))
              .Font.Strikeout = NullSafeBoolean(dr("FontStrikeout"))

              .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
            End With

        End Select
      End While

    End Using

  End Sub

End Class

Public Class FontSetting
  Public Name As String
  Public Size As Single
  Public Bold As Boolean
  Public Italic As Boolean
  Public Underline As Boolean
  Public Strikeout As Boolean
End Class