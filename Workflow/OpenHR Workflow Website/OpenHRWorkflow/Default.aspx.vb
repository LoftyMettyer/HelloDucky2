Option Strict On

Imports System
Imports System.Data
Imports System.Globalization
Imports System.Threading
Imports System.Drawing
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports Utilities
Imports System.Data.SqlClient
Imports System.Transactions
Imports System.Reflection

Public Class [Default]
   Inherits System.Web.UI.Page

   Private _config As Config
   Private _instanceID As Integer
   Private _elementID As Integer
   Private _server As String
   Private _database As String
   Private _user As String
   Private _password As String
   Private _db As Database
   Private _form As WorkflowForm
   Private _imageCount As Integer
   Private _autoFocusControl As String

   Private Const TabStripHeight As Integer = 21
   Private Const FormInputPrefix As String = "FI_"
   Private Const AssemblyName As String = "OPENHRWORKFLOW"
   Private Const DefaultTitle As String = "OpenHR Workflow"

   Dim sMessage As String = ""

   Protected Sub Page_Init(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Init

      Dim sTemp As String
      Dim sQueryString As String
      Dim objCrypt As New Crypt
      Dim sKeyParameter As String = ""
      Dim sSiblingForms As String

      'Decript the url (use the rawUrl rather than the querystring itself, some characters are ignored in the querystring)
      Try
         sTemp = Server.UrlDecode(Request.RawUrl)
         Dim iTemp As Integer = sTemp.IndexOf("?")

         If iTemp >= 0 Then
            sQueryString = sTemp.Substring(iTemp + 1)
         Else
            ' NPG20120326 Fault HRPRO-2128
            Response.Redirect("~/Account/Login.aspx", False)
            Return
            'TODO this still goes through Page_Load
         End If

         ' Try the newer encryption first
         Try
            ' Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
            Dim sCultureName As String = Thread.CurrentThread.CurrentCulture.Name

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-GB")

            sTemp = objCrypt.DecompactString(sQueryString)
            sTemp = objCrypt.DecryptString(sTemp, "", True)

            ' Reset the culture to be the one used by the client. Fault HRPRO-1404
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(sCultureName)
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(sCultureName)

            ' Extract the required parameters from the decrypted queryString.
            Dim values = sTemp.Split(vbTab(0))

            _instanceID = CInt(values(0))
            _elementID = CInt(values(1))
            _user = values(2)
            _password = values(3)
            _server = values(4)
            _database = values(5)
            If values.Count > 6 Then sKeyParameter = values(6)

         Catch ex As Exception
            ' Older encryption method
            sQueryString = objCrypt.ProcessDecryptString(sQueryString)
            sTemp = objCrypt.DecryptString(sQueryString, "", False)

            ' Extract the required parameters from the decrypted queryString.
            If _instanceID = 0 Then
               _instanceID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
            End If
            sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

            If _elementID = 0 Then
               _elementID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
            End If
            sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

            _user = Left(sTemp, InStr(sTemp, vbTab) - 1)
            sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

            _password = Left(sTemp, InStr(sTemp, vbTab) - 1)
            sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

            _server = Left(sTemp, InStr(sTemp, vbTab) - 1)
            sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

            _database = Mid(sTemp, InStr(sTemp, vbTab) + 1)

         End Try
      Catch ex As Exception
         sMessage = "Invalid query string."
      End Try

      'Set up the database class
      'TODO can we trust this is fixed see db in later code
      _db = New Database(GetConnectionString)

      'Activating mobile security. I've hijacked the _instanceID and populated it with the User ID that is to be activated.
      If sMessage.Length = 0 And Not IsPostBack And _elementID = - 2 And _instanceID > 0 Then

         sMessage = _db.ActivateUser(_instanceID)

         If sMessage.Length = 0 Then
            sMessage = "You have been successfully activated"
         End If
      End If

      If sMessage.Length = 0 And Not IsPostBack And _instanceID < 0 And _elementID = - 1 Then

         ' Externally initiated Workflow.
         Dim workflowID = - _instanceID

         Dim result As InstantiateWorkflowResult

         If sKeyParameter.Length > 0 Then
            'Instantiate from mobile
            result = _db.InstantiateWorkflow(workflowID, sKeyParameter)
         Else
            result = _db.InstantiateWorkflow(workflowID)
         End If

         'TODO redirect to url for now already existing workflow
         'TODO what is follow on forms for

         _instanceID = result.InstanceId

         If result.Message.Length = 0 Then
            If result.FormElements.Length = 0 Then
               sMessage = "Workflow initiated successfully."
            Else
               Dim followOnForms =
                      result.FormElements.Split(New String() {vbTab}, StringSplitOptions.RemoveEmptyEntries).ToList

               _elementID = CInt(followOnForms(0))
               followOnForms.RemoveAt(0)

               sSiblingForms = String.Join(vbTab,
                                           followOnForms.Select(
                                              Function(f) _db.GetWorkflowQueryString(_instanceID, CInt(f))))

               'TODO dont need this if we're redirecting and processing a new url
               'TODO how to deal with subling forms
            End If
         Else
            sMessage = "Error:<BR><BR>" & result.Message
         End If

      End If
   End Sub

   Protected Sub Page_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load

      Dim sAssemblyName As String = ""
      Dim sWebSiteVersion As String = ""

      Dim sSiblingForms As String = ""

      Try
         Response.CacheControl = "no-cache"
         Response.AddHeader("Pragma", "no-cache")
         Response.Expires = - 1

         _config = New Config()
         _config.Initialise(Server.MapPath("themes/ThemeHex.xml"))

      Catch ex As Exception
      End Try

      'Set the page title
      Dim pageTitle As String = DefaultTitle

      'TODO PG NOW can this ever be zero length see other code
      With Assembly.GetExecutingAssembly.GetName.Version
         sWebSiteVersion = String.Format("{0}.{1}.{2}", .Major, .Minor, .Build)
      End With

      sAssemblyName = Assembly.GetExecutingAssembly.GetName.Name.ToUpper

      If sAssemblyName = AssemblyName Then
         'Compiled version of the web site, so perform version checks.
         pageTitle += " - v" & sWebSiteVersion
      Else
         'Development version of the web site
         pageTitle += " (development)"
      End If

      Page.Title = pageTitle

      'Set the page culture
      SetCulture()

      If IsPostBack Then
         _instanceID = CInt(ViewState("InstanceID"))
         _elementID = CInt(ViewState("ElementID"))
         _user = ViewState("User").ToString
         _password = ViewState("Pwd").ToString
         _server = ViewState("Server").ToString
         _database = ViewState("Database").ToString
      End If

      Dim db As New Database(GetConnectionString)

      'check to see if the database is locked
      If sMessage.Length = 0 And Not IsPostBack Then

         If db.IsSystemLocked() Then
            sMessage = "Database locked.<BR><BR>Contact your system administrator."
         End If
      End If

      'check if the database and website versions match.
      If Not IsPostBack Then

         Dim dbVersion As String = db.GetSetting("database", "version", False)

         If sAssemblyName = AssemblyName Then
            ' Complied version of the web site, so perform version checks.
            If sWebSiteVersion.Length > 0 Then
               ' Just get the major and minor parts of the 4 part version.
               sWebSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major & "." &
                                 Assembly.GetExecutingAssembly.GetName.Version.Minor
            End If

            If dbVersion <> sWebSiteVersion Or sWebSiteVersion.Length = 0 Then
               ' Version mismatch.
               If dbVersion.Length = 0 Then
                  dbVersion = "&lt;unknown&gt;"
               End If
               If sWebSiteVersion.Length = 0 Then
                  sWebSiteVersion = "&lt;unknown&gt;"
               End If

               sMessage = "The Workflow website version (" & sWebSiteVersion & ")" &
                          " is incompatible with the database version (" & dbVersion & ")." &
                          "<BR><BR>Contact your system administrator."
            End If
         End If
      End If

      If sMessage.Length = 0 Then

         Try
            ScriptManager.GetCurrent(Page).AsyncPostBackTimeout = _config.SubmissionTimeout

            If sMessage.Length = 0 Then
               ' Remember the useful parameters for use in postbacks.

               ViewState("InstanceID") = _instanceID
               ViewState("ElementID") = _elementID
               ViewState("User") = _user
               ViewState("Pwd") = _password
               ViewState("Server") = _server
               ViewState("Database") = _database

               'FileUpload.apsx and FileDownload.aspx require these variables
               Session("User") = _user
               Session("Pwd") = _password
               Session("Server") = _server
               Session("Database") = _database
               Session("ElementID") = _elementID
               Session("InstanceID") = _instanceID

               ' Get the selected tab number for this workflow, if any...
               If Not IsPostBack Then
                  hdnDefaultPageNo.Value = db.GetWorkflowCurrentTab(_instanceID).ToString
               End If

               _form = _db.GetWorkflowForm(_instanceID, _elementID)

               Dim dr As SqlDataReader = _form.Items

               Dim scriptString As String = "function pageLoad() {"

               Dim tabPages() As Panel
               ReDim Preserve tabPages(0)

               While dr.Read And sMessage.Length = 0

                  Dim tabPage As Integer = NullSafeInteger(dr("pageno"))

                  ' Create the tab for this control. Do this first in case the tabstrip control hasn't been read yet,
                  ' and the tabs haven't been generated.
                  Try
                     Dim strTemp As String = tabPages(tabPage).ID.ToString
                     ' OK, if the id exists, the div has already been created. Do nothing.
                  Catch ex As Exception
                     ' Create the new div, give it a unique id then we can refer to that when it's reused in the next loop.
                     ' store the id in the array for reference. NB 21 is the itemtype for a page Tab
                     If tabPage > tabPages.GetUpperBound(0) Then ReDim Preserve tabPages(tabPage)

                     tabPages(tabPage) = New Panel
                     tabPages(tabPage).ID = FormInputPrefix & tabPage.ToString & "_21_PageTab"
                     tabPages(tabPage).Style.Add("position", "absolute")

                     ' Add this tab to the web form
                     pnlInputDiv.Controls.Add(tabPages(tabPage))
                  End Try

                  ' Generate the unique ID for this control and process it onto the form.
                  Dim sID As String = FormInputPrefix & NullSafeString(dr("id")) & "_" & NullSafeString(dr("ItemType")) &
                                      "_"

                  Select Case NullSafeInteger(dr("ItemType"))

                     Case 0 ' Button
                        Dim control = New HtmlInputButton
                        With control
                           .ID = sID
                           .Style.ApplyLocation(dr)
                           .Style.ApplySize(dr)
                           .Style.ApplyFont(dr)

                           .Attributes.Add("TabIndex", NullSafeInteger(dr("tabIndex")).ToString)
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           ' If the button has no caption, we treat as a transparent button.
                           ' This is so we can emulate picture buttons with very little code changes.
                           If NullSafeString(dr("caption")) = vbNullString Then
                              .Style.Add("filter", "alpha(opacity=0)")
                              .Style.Add("opacity", "0")
                           End If

                           ' stops the mobiles displaying buttons with over-rounded corners...
                           If IsMobileBrowser() OrElse IsMacSafari() Then
                              .Style.Add("-webkit-appearance", "none")
                              .Style.Add("background-color", "#E6E6E6")
                              .Style.Add("border", "solid 1px #CCC")
                              .Style.Add("border-radius", "4px")
                           End If

                           If _
                              NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso
                              NullSafeInteger(dr("BackColor")) <> - 2147483633 Then
                              .Style.Add("background-color",
                                         General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
                              .Style.Add("border", "1px solid #CCC")
                              .Style.Add("border-radius", "4px")
                           End If

                           If NullSafeInteger(dr("ForeColor")) <> 6697779 Then
                              .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))).ToString)
                           End If

                           .Style.Add("padding", "0px")
                           .Style.Add("white-space", "normal")

                           .Value = NullSafeString(dr("caption"))

                           .Style.Add("z-index", "2")

                           .Attributes.Add("onclick", "try{setPostbackMode(1);}catch(e){};")
                        End With

                        tabPages(tabPage).Controls.Add(control)

                        AddHandler control.ServerClick, AddressOf ButtonClick

                     Case 1 ' Database value

                        Dim control = New Label
                        With control
                           .ApplyLocation(dr)
                           .ApplySize(dr)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)

                           If NullSafeBoolean(dr("PictureBorder")) Then
                              .ApplyBorder(True)
                           End If

                           .Style("word-wrap") = "break-word"
                           .Style("overflow") = "auto"

                           Select Case NullSafeInteger(dr("sourceItemType"))
                              Case - 7 ' Logic
                                 If NullSafeString(dr("value")) = String.Empty Then
                                    .Text = "&lt;undefined&gt;"
                                 ElseIf NullSafeString(dr("value")) = "1" Then
                                    .Text = Boolean.TrueString
                                 Else
                                    .Text = Boolean.FalseString
                                 End If

                              Case 2, 4 ' Numeric, Integer
                                 If IsDBNull(dr("value")) Then
                                    .Text = "&lt;undefined&gt;"
                                 Else
                                    Dim value = CStr(dr("value")).Replace(".",
                                                                          Thread.CurrentThread.CurrentCulture.
                                                                            NumberFormat.NumberDecimalSeparator)
                                    If value.Chars(0) = "-" Then
                                       value = value.Substring(1) & "-"
                                    End If
                                    .Text = value
                                 End If

                              Case 11 ' Date
                                 If NullSafeString(dr("value")) = String.Empty Then
                                    .Text = "&lt;undefined&gt;"
                                 ElseIf CStr(dr("value")).Trim.Length = 0 Then
                                    .Text = "&lt;undefined&gt;"
                                 Else
                                    .Text = General.ConvertSqlDateToLocale(NullSafeString(dr("value")))
                                 End If
                              Case Else 'Text
                                 .Text = NullSafeString(dr("value"))
                           End Select

                        End With

                        tabPages(tabPage).Controls.Add(control)

                     Case 2 ' Label
                        Dim control = New Label
                        With control
                           .ApplyLocation(dr)
                           .ApplySize(dr, 0, 1)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)

                           If NullSafeBoolean(dr("PictureBorder")) Then
                              .ApplyBorder(True)
                           End If

                           ' NPG20120305 Fault HRPRO-1967 reverted by PBG20120419 Fault HRPRO-2157
                           '.Style("word-wrap") = "break-word"
                           .Style("overflow") = "auto"

                           If NullSafeInteger(dr("captionType")) = 3 Then 'calculated caption
                              .Text = NullSafeString(dr("value"))
                           Else
                              .Text = NullSafeString(dr("caption"))
                           End If
                        End With

                        tabPages(tabPage).Controls.Add(control)

                     Case 3 ' Input value - character
                        Dim control = New TextBox
                        With control
                           .ID = sID
                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           .ApplyLocation(dr)
                           .ApplySize(dr, - 1, - 1)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr)
                           .ApplyBorder(True)

                           If NullSafeBoolean(dr("PasswordType")) Then
                              .TextMode = TextBoxMode.Password
                           Else
                              .TextMode = TextBoxMode.MultiLine
                              .Wrap = True
                              .Style("overflow") = "auto"
                              .Style("word-wrap") = "break-word"
                              .Style("resize") = "none"
                           End If
                           .Style("padding") = "1px"

                           .Text = NullSafeString(dr("value"))

                           .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                           If NullSafeInteger(dr("inputSize")) > 0 Then
                              .Attributes("maxlength") = NullSafeString(dr("inputSize"))
                           End If

                           If IsMobileBrowser() Then
                              .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                           End If

                        End With

                        tabPages(tabPage).Controls.Add(control)

                     Case 4 ' Workflow value

                        Dim control = New Label
                        With control
                           .ApplyLocation(dr)
                           .ApplySize(dr)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)

                           If NullSafeBoolean(dr("PictureBorder")) Then
                              .ApplyBorder(True)
                           End If

                           .Style("word-wrap") = "break-word"
                           .Style("overflow") = "auto"

                           Select Case NullSafeInteger(dr("sourceItemType"))
                              Case 6 ' Logic
                                 If NullSafeString(dr("value")) = String.Empty Then
                                    .Text = "&lt;undefined&gt;"
                                 ElseIf NullSafeString(dr("value")) = "1" Then
                                    .Text = Boolean.TrueString
                                 Else
                                    .Text = Boolean.FalseString
                                 End If

                              Case 5 ' Number
                                 If NullSafeString(dr("value")) = String.Empty Then
                                    .Text = String.Empty
                                 Else
                                    Dim value = NullSafeString(dr("value")).Replace(".",
                                                                                    Thread.CurrentThread.CurrentCulture.
                                                                                      NumberFormat.
                                                                                      NumberDecimalSeparator)
                                    If value.Length > 0 AndAlso value.Chars(0) = "-" Then
                                       value = value.Substring(1) & "-"
                                    End If
                                    .Text = value
                                 End If

                              Case 7 ' Date
                                 If IsDBNull(dr("value")) Then
                                    .Text = "&lt;undefined&gt;"
                                 ElseIf CStr(dr("value")).Trim.ToUpper = "NULL" Then
                                    .Text = "&lt;undefined&gt;"
                                 Else
                                    .Text = General.ConvertSqlDateToLocale(NullSafeString(dr("value")))
                                 End If
                              Case Else 'Text
                                 .Text = NullSafeString(dr("value"))
                           End Select

                        End With

                        tabPages(tabPage).Controls.Add(control)

                     Case 5 ' Input value - numeric

                        Dim control = New TextBox
                        With control
                           .ID = sID
                           .CssClass = "numeric"

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           .ApplyLocation(dr)
                           .ApplySize(dr, - 1, - 1)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)
                           .ApplyBorder(True)
                           .Style("padding") = "1px"

                           'add attributes that denote the min & max values, number of decimal places is also implied
                           Dim max =
                                  New String("9"c,
                                             NullSafeInteger(dr("inputSize")) - NullSafeInteger(dr("inputDecimals"))) &
                                  If _
                                     (NullSafeInteger(dr("inputDecimals")) > 0,
                                      "." & New String("9"c, NullSafeInteger(dr("inputDecimals"))), "")

                           .Attributes.Add("data-numeric", String.Format("{{vMin: '-{0}', vMax: '{0}'}}", max))

                           'set the control value
                           Dim value As Single
                           If NullSafeString(dr("value")) <> "" Then
                              value = CSng(NullSafeString(dr("value")).Replace(".",
                                                                               Thread.CurrentThread.CurrentCulture.
                                                                                 NumberFormat.NumberDecimalSeparator))
                           End If
                           .Text =
                              value.ToString("N" & NullSafeInteger(dr("inputDecimals"))).Replace(
                                 Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator, "")

                           .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                           If IsMobileBrowser() Then
                              .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                           End If

                        End With
                        tabPages(tabPage).Controls.Add(control)

                     Case 6 ' Input value - logic

                        Dim checkBox = New CheckBox
                        With checkBox
                           .ID = sID
                           .ApplyLocation(dr)
                           .ApplySize(dr)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           .CssClass = If(NullSafeInteger(dr("alignment")) = 0, "checkbox left", "checkbox right")
                           If IsAndroidBrowser() Then .CssClass += " android"
                           .Style("line-height") = NullSafeInteger(dr("Height")).ToString & "px"

                           .Text = NullSafeString(dr("caption"))
                           .Checked = (NullSafeString(dr("value")).ToLower = "true")

                           If IsMobileBrowser() Then
                              .Attributes("onclick") = "FilterMobileLookup('" & sID & "');"
                           End If
                        End With

                        tabPages(tabPage).Controls.Add(checkBox)

                     Case 7 ' Input value - date

                        Dim control = New TextBox
                        With control
                           .ID = sID
                           .CssClass = "date"

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           .Style.ApplyFont(dr)
                           .ApplyColor(dr, True)

                           If GetBrowserFamily() = "IOS" Then
                              'use html5 date
                              .Attributes.Add("type", "date")
                              'ios sizing fix
                              .ApplySize(dr, - 10, - 3)
                              'ios requires date in yyyy-mm-dd format
                              .Text =
                                 If _
                                    (NullSafeString(dr("value")) = "", "",
                                     DateTime.ParseExact(NullSafeString(dr("value")), "MM/dd/yyyy", Nothing).ToString(
                                        "yyyy-MM-dd"))
                           Else
                              .CssClass += " withPicker"
                              .ApplySize(dr, - 1, - 1)
                              .ApplyBorder(True)
                              .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"
                              .Text = General.ConvertSqlDateToLocale(NullSafeString(dr("value")))
                              If IsMobileBrowser() Then
                                 'stop keyboard popping up on mobiles
                                 .ReadOnly = True
                              End If
                           End If

                           If IsMobileBrowser() Then
                              .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                           End If
                        End With

                        Dim panel As New Panel
                        panel.Controls.Add(control)
                        panel.ApplyLocation(dr)

                        tabPages(tabPage).Controls.Add(panel)

                     Case 8 ' Frame

                        Dim top = NullSafeInteger(dr("TopCoord"))
                        Dim left = NullSafeInteger(dr("LeftCoord"))
                        Dim width = NullSafeInteger(dr("Width"))
                        Dim height = NullSafeInteger(dr("Height"))
                        Dim fontAdjustment = CInt(CInt(dr("FontSize"))*0.8)

                        width -= 2
                        height -= 2

                        If NullSafeString(dr("caption")).Trim.Length = 0 Then
                           top += fontAdjustment
                           height -= fontAdjustment
                        End If

                        Dim html = "<fieldset style='" &
                                   " position: absolute;" &
                                   " top: " & top & "px;" &
                                   " left: " & left & "px;" &
                                   " width: " & width & "px;" &
                                   " height: " & height & "px;" &
                                   " " & GetFontCss(dr) &
                                   " " & GetColorCss(dr) &
                                   " border: 1px solid #999;" &
                                   " '>"

                        If NullSafeString(dr("caption")).Trim.Length > 0 Then
                           html += String.Format("<legend>{0}</legend>", NullSafeString(dr("caption"))) & vbCrLf
                        End If

                        html += "</fieldset>" & vbCrLf

                        tabPages(tabPage).Controls.Add(New LiteralControl(html))

                     Case 9 ' Line

                        Dim html As String

                        Select Case NullSafeInteger(dr("Orientation"))
                           Case 0
                              ' Vertical
                              html = "<div style='position: absolute;" &
                                     " left: " & NullSafeString(dr("LeftCoord")) & "px;" &
                                     " top: " & NullSafeString(dr("TopCoord")) & "px;" &
                                     " height: " & NullSafeString(dr("Height")) & "px;" &
                                     " width: 0px;" &
                                     " border-left: 1px solid " &
                                     General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) & ";'" &
                                     "></div>"
                           Case Else
                              ' Horizontal
                              html = "<img style='position: absolute;" &
                                     " left: " & NullSafeString(dr("LeftCoord")) & "px;" &
                                     " top: " & NullSafeString(dr("TopCoord")) & "px;" &
                                     " height: 0px;" &
                                     " width: " & NullSafeString(dr("Width")) & "px;" &
                                     " border-top: 1px solid " & General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) &
                                     ";'" &
                                     "></div>"
                        End Select

                        tabPages(tabPage).Controls.Add(New LiteralControl(html))

                     Case 10 ' Image

                        Dim control = New WebControls.Image

                        With control
                           .ApplyLocation(dr)
                           .ApplySize(dr)

                           If NullSafeBoolean(dr("PictureBorder")) Then
                              .ApplyBorder(True, - 2)
                           End If

                           Dim imageUrl As String = LoadPicture(NullSafeInteger(dr("pictureID")), sMessage)
                           If sMessage.Length > 0 Then
                              Exit While
                           End If
                           .ImageUrl = imageUrl
                        End With

                        tabPages(tabPage).Controls.Add(control)

                     Case 11 ' Record Selection Grid
                        ' NPG20110501 Fault HR PRO 1414
                        ' We're using the ASP.NET standard gridview control now. To replicate the old infragistics
                        ' grid we'll put the Gridview control within a DIV to enable scroll bars and fix the height&width, 
                        ' but also put a header DIV above the grid which contains copies of the column headers. This is 
                        ' to simulate fixing the headers when the grid is scrolled. We use this table to allow 
                        ' clickable sorting and resizable column widths.
                        '
                        ' =========================================================
                        ' Grids are now created using the clsRecordSelector class.
                        ' =========================================================

                        Dim recordSelector = New RecordSelector
                        With recordSelector

                           .CssClass = "recordSelector"
                           .Style.Add("Position", "Absolute")
                           .Attributes.CssStyle("LEFT") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                           .Attributes.CssStyle("TOP") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                           .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString

                           ' Don't use .height - it causes large row heights if the grid isn't filled.
                           ' Use .ControlHeight instead - custom property.
                           .ControlHeight = NullSafeInteger(dr("Height"))

                           .Width = NullSafeInteger(dr("Width"))

                           .BorderColor = Color.Black
                           .BorderStyle = BorderStyle.Solid
                           .BorderWidth = 1

                           .Style.Add("border-bottom-width", "2px")

                           .ID = sID & "Grid"
                           .AllowPaging = True
                           .AllowSorting = True
                           '.EnableSortingAndPagingCallbacks = True

                           ' Androids currently can't scroll internal divs, so fix 
                           ' pagesize of record selector to height of control.
                           If GetBrowserFamily() = "ANDROID" Then
                              Dim piRowHeight As Double = (CInt(NullSafeString(dr("FontSize"))) - 8) + 21
                              .PageSize =
                                 Math.Min(CInt(Math.Truncate((CInt(NullSafeInteger(dr("Height")) - 42)/piRowHeight))),
                                          _config.LookupRowsRange)
                              .RowStyle.Height = Unit.Pixel(CInt(piRowHeight))
                           Else
                              .PageSize = _config.LookupRowsRange
                           End If

                           .IsLookup = False
                           ' EnableViewState must be on. Mucks up the grid data otherwise. Should be reviewed
                           ' if performance is silly, but while paging is enabled it shouldn't be too bad.
                           .EnableViewState = True
                           .IsEmpty = False
                           .EmptyDataText = "no records to display"

                           ' Header Row
                           .ColumnHeaders = NullSafeBoolean(dr("ColumnHeaders"))
                           .HeadFontSize = NullSafeSingle(dr("HeadFontSize"))
                           .HeadLines = NullSafeInteger(dr("Headlines"))

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           Dim backColor As Integer = NullSafeInteger(dr("BackColor"))

                           If backColor = 16777215 AndAlso NullSafeInteger(dr("BackColorEven")) = 15988214 Then
                              backColor = NullSafeInteger(dr("BackColorEven"))
                           End If

                           .BackColor = General.GetColour(backColor)
                           .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                           .HeaderStyle.BackColor = General.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                           .HeaderStyle.BorderColor = General.GetColour(10720408)
                           .HeaderStyle.BorderStyle = BorderStyle.Double
                           .HeaderStyle.BorderWidth = Unit.Pixel(0)

                           .HeaderStyle.Font.Apply(dr, "Head")

                           .HeaderStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                           .HeaderStyle.Wrap = False
                           .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                           .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                           ' PagerStyle settings
                           .PagerStyle.BackColor = General.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                           .PagerStyle.BorderColor = General.GetColour(10720408)
                           .PagerStyle.BorderStyle = BorderStyle.Solid
                           .PagerStyle.BorderWidth = Unit.Pixel(0)

                           .PagerStyle.Font.Apply(dr, "Head")

                           .PagerStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                           .PagerStyle.Wrap = False
                           .PagerStyle.VerticalAlign = VerticalAlign.Middle
                           .PagerStyle.HorizontalAlign = HorizontalAlign.Center

                           .Font.Apply(dr)

                           If NullSafeInteger(dr("ForeColorEven")) <> NullSafeInteger(dr("ForeColor")) Then
                              .RowStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColorEven")))
                           End If

                           If NullSafeInteger(dr("BackColorEven")) <> backColor Then
                              .RowStyle.BackColor = General.GetColour(NullSafeInteger(dr("BackColorEven")))
                           End If

                           If NullSafeInteger(dr("ForeColorOdd")) <> NullSafeInteger(dr("ForeColor")) Then
                              .AlternatingRowStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColorOdd")))
                           End If

                           If NullSafeInteger(dr("BackColorOdd")) <> NullSafeInteger(dr("BackColorEven")) Then
                              .AlternatingRowStyle.BackColor = General.GetColour(NullSafeInteger(dr("BackColorOdd")))
                           End If

                           If IsDBNull(dr("ForeColorHighlight")) Then
                              .SelectedRowStyle.ForeColor = SystemColors.HighlightText
                           Else
                              .SelectedRowStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColorHighlight")))
                           End If
                           If IsDBNull(dr("BackColorHighlight")) Then
                              .SelectedRowStyle.BackColor = SystemColors.Highlight
                           Else
                              .SelectedRowStyle.BackColor = General.GetColour(NullSafeInteger(dr("BackColorHighlight")))
                           End If

                        End With

                        ' ==================================================
                        ' Add the Paging Grid View to the holding panel now.
                        ' Before the databind, or you'll get errors!
                        ' ==================================================
                        tabPages(tabPage).Controls.Add(recordSelector)

                        If Not IsPostBack Then

                           Dim result = db.GetWorkflowGridItems(NullSafeInteger(dr("id")), _instanceID)

                           Session(sID & "DATA") = result.Data

                           recordSelector.DataKeyNames = New String() {"ID"}

                           If result.Data.Rows.Count > 0 Then
                              recordSelector.IsEmpty = False
                              recordSelector.DataSource = result.Data
                              recordSelector.DataBind()
                           Else
                              recordSelector.IsEmpty = True
                              ShowNoResultFound(result.Data, recordSelector)
                           End If

                           'set the default value
                           If NullSafeInteger(dr("value")) <> 0 Then

                              Dim colIndex As Integer = result.Data.Columns.IndexOf("ID")

                              For r = 0 To result.Data.Rows.Count - 1
                                 If result.Data.Rows(r).Item(colIndex).ToString = CStr(dr("value")) Then
                                    ' set selected page index and row number
                                    recordSelector.PageIndex = CInt(r\recordSelector.PageSize)
                                    recordSelector.SelectedIndex = CInt(r Mod recordSelector.PageSize)
                                    recordSelector.DataBind()
                                    Exit For
                                 End If
                              Next
                           End If

                           If recordSelector.SelectedIndex = - 1 AndAlso recordSelector.Rows.Count > 0 Then
                              recordSelector.SelectedIndex = 0
                           End If

                           If Not result.Ok Then
                              sMessage =
                                 "Error loading web form. Web Form record selector item record has been deleted or not selected."
                              Exit While
                           End If
                        Else
                           ' If a postback, check for empty datagrid and set empty row message
                           If recordSelector.IsEmpty Then
                              ShowNoResultFound(TryCast(HttpContext.Current.Session(sID & "DATA"), DataTable),
                                                recordSelector)
                           End If
                        End If

                        ' ============================================================
                        ' Hidden field is used to store scroll position of the grid.
                        ' ============================================================
                        tabPages(tabPage).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})


                     Case 14 ' lookup  Inputs
                        If Not IsMobileBrowser() Then

                           ' ============================================================
                           ' Create a textbox as the main control
                           ' ============================================================
                           Dim textBox = New TextBox

                           With textBox
                              .ID = sID & "TextBox"
                              .ApplyLocation(dr)
                              .ApplySize(dr, - 1, - 1)
                              .Style.ApplyFont(dr)
                              .ApplyColor(dr)
                              .ApplyBorder(True)

                              .TabIndex = NullSafeShort(dr("tabIndex"))
                              UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID & "TextBox")

                              .ReadOnly = True
                              .Style.Add("padding", "1px")
                              .Style.Add("background-image", "url('images/downarrow.gif')")
                              .Style.Add("background-position", "right top")
                              .Style.Add("background-repeat", "no-repeat")
                              .Style.Add("background-origin", "content-box")
                              .Style.Add("background-size", "17px 100%")
                           End With

                           tabPages(tabPage).Controls.Add(textBox)

                           ' ============================================================
                           ' Create the Lookup table grid, as per normal record selectors.
                           ' This will be hidden on page_load, and displayed when the 
                           ' DropdownList above is clicked. The magic is brought together
                           ' by the AJAX DropDownExtender control below.
                           ' ============================================================
                           Dim recordSelector = New RecordSelector

                           With recordSelector
                              .ID = sID & "Grid"
                              .IsLookup = True
                              .EnableViewState = True
                              ' Must be set to True
                              .IsEmpty = False
                              .EmptyDataText = "no records to display"
                              .AllowPaging = True
                              .AllowSorting = True
                              '.EnableSortingAndPagingCallbacks = True
                              .PageSize = _config.LookupRowsRange
                              .ShowFooter = False

                              .CssClass = "recordSelector"
                              .Style.Add("Position", "Absolute")
                              .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                              .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                              .Attributes.CssStyle("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                              .Attributes.CssStyle("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                              .Attributes.CssStyle("width") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString

                              ' Don't set the height of this control. Must use the ControlHeight property
                              ' to stop the grid's rows from autosizing.
                              .ControlHeight = NullSafeInteger(dr("Height"))
                              .Width = Unit.Pixel(NullSafeInteger(dr("Width")))

                              ' Header Row - fixed for lookups.
                              .ColumnHeaders = True
                              .HeadFontSize = NullSafeSingle(dr("FontSize"))
                              .HeadLines = 1

                              .ApplyFont(dr)
                              .ApplyColor(dr)
                              .ApplyBorder(False)

                              .SelectedRowStyle.ForeColor = General.GetColour(2774907)
                              .SelectedRowStyle.BackColor = General.GetColour(10480637)

                              ' HEADER formatting
                              .HeaderStyle.BackColor = General.GetColour(16248553)
                              .HeaderStyle.BorderColor = General.GetColour(10720408)
                              .HeaderStyle.BorderStyle = BorderStyle.Solid
                              .HeaderStyle.BorderWidth = Unit.Pixel(0)

                              .HeaderStyle.Font.Apply(dr)
                              .HeaderStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                              .HeaderStyle.Wrap = False
                              .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                              .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                              .PagerStyle.Font.Apply(dr)
                              .PagerStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                              .PagerStyle.Wrap = False
                              .PagerStyle.VerticalAlign = VerticalAlign.Middle
                              .PagerStyle.HorizontalAlign = HorizontalAlign.Center
                              .PagerStyle.BorderWidth = Unit.Pixel(0)
                           End With

                           Dim filterSql = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")),
                                                           NullSafeInteger(dr("lookupFilterColumnDataType")),
                                                           NullSafeInteger(dr("LookupFilterOperator")),
                                                           FormInputPrefix & NullSafeString(dr("lookupFilterValueID")) &
                                                           "_" & NullSafeString(dr("lookupFilterValueType")) & "_")


                           ' ==========================================================
                           ' Hidden Field to store any lookup filter code
                           ' ==========================================================
                           If (filterSql.Length > 0) Then
                              tabPages(tabPage).Controls.Add(New HiddenField _
                                                               With {.ID = "lookup" & sID, .Value = filterSql})
                           End If

                           tabPages(tabPage).Controls.Add(recordSelector)

                           If Not IsPostBack Then

                              'get the data
                              Dim result = db.GetWorkflowItemValues(CInt(NullSafeString(dr("id"))), _instanceID)

                              'insert a blank row
                              result.Data.Rows.InsertAt(result.Data.NewRow(), 0)

                              'bind data to grid
                              recordSelector.IsEmpty = (result.Data.Rows.Count - 1 = 0)
                              recordSelector.DataSource = result.Data
                              recordSelector.DataBind()

                              'store the data its needed for paging, sorting
                              Session(sID & "DATA") = result.Data

                              'store info its needed later
                              textBox.Attributes.Add("LookupColumnIndex", result.LookupColumnIndex.ToString)
                              textBox.Attributes.Add("DataType",
                                                     result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

                              'set the default value
                              textBox.Text = result.DefaultValue

                              For i As Integer = 0 To recordSelector.Rows.Count - 1
                                 If i > recordSelector.PageSize Then Exit For
                                 ' don't bother if on other pages
                                 If recordSelector.Rows(i).Cells(result.LookupColumnIndex).Text = result.DefaultValue _
                                    Then
                                    recordSelector.SelectedIndex = i
                                    Exit For
                                 End If
                              Next
                           End If

                           ' =============================================================================
                           ' AJAX DropDownExtender (DDE) Control
                           ' This simply links up the DropDownList and the Lookup Grid to make a dropdown.
                           ' =============================================================================
                           Dim dde As New AjaxControlToolkit.DropDownExtender

                           With dde
                              .DropArrowImageUrl = "~/Images/Blank.gif"
                              .DropArrowBackColor = Color.Transparent
                              .HighlightBackColor = textBox.BackColor
                              .HighlightBorderColor = textBox.BorderColor

                              ' Careful with the case here, use 'dde' in JavaScript:
                              .ID = sID & "DDE"
                              .BehaviorID = sID & "dde"
                              .DropDownControlID = sID
                              .Enabled = True
                              .TargetControlID = sID & "TextBox"
                              ' Client-side handler.
                              If (filterSql.Length > 0) Then
                                 .OnClientPopup = "InitializeLookup"
                                 ' can't pass the ID of the control, so use ._id in JS.
                              End If
                           End With

                           tabPages(tabPage).Controls.Add(dde)

                           ' =================================================================
                           ' Attach a JavaScript functino to the 'add_shown' method of this
                           ' DropDownExtender. Used to check if popup is bigger than the
                           ' parent form, and resize the parent form if necessary
                           ' =================================================================
                           scriptString += "var bhvDdl=$find('" & dde.BehaviorID.ToString & "');"
                           scriptString += "try {bhvDdl.add_shown(ResizeComboForForm);} catch (e) {}"

                           ' ====================================================
                           ' hidden field to store scroll position (not required?)
                           ' ====================================================
                           tabPages(tabPage).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})

                           ' ====================================================
                           ' hidden field to hold any filter SQL code
                           ' ====================================================
                           tabPages(tabPage).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

                           ' ============================================================
                           ' Hidden Button for JS to call which fires filter click event. 
                           ' ============================================================
                           Dim button = New Button
                           With button
                              .ID = sID & "refresh"
                              .Style.Add("display", "none")
                              .Text = .ID
                           End With

                           AddHandler button.Click, AddressOf SetLookupFilter

                           tabPages(tabPage).Controls.Add(button)

                        Else
                           ' ================================================================================================================
                           ' Mobile Browser - convert lookup data to a standard dropdown.
                           ' ================================================================================================================
                           Dim control As New DropDownList

                           With control
                              .ID = sID
                              .ApplyLocation(dr)
                              .ApplySize(dr, - 1, - 1)
                              .Style.ApplyFont(dr)
                              .ApplyColor(dr)
                              If Not IsMobileBrowser() Then .ApplyBorder(False)
                              .Style.Add("padding", "1px")

                              .TabIndex = NullSafeShort(dr("tabIndex"))
                              UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                              .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")

                              tabPages(tabPage).Controls.Add(control)

                              Dim filterSql = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")),
                                                              NullSafeInteger(dr("lookupFilterColumnDataType")),
                                                              NullSafeInteger(dr("LookupFilterOperator")),
                                                              FormInputPrefix &
                                                              NullSafeString(dr("lookupFilterValueID")) & "_" &
                                                              NullSafeString(dr("lookupFilterValueType")) & "_")

                              If (filterSql.Length > 0) Then
                                 tabPages(tabPage).Controls.Add(New HiddenField _
                                                                  With {.ID = "lookup" & sID, .Value = filterSql})
                              End If

                              If Not IsPostBack Then

                                 'get the data
                                 Dim result = db.GetWorkflowItemValues(CInt(NullSafeString(dr("id"))), _instanceID)

                                 'insert a blank row
                                 result.Data.Rows.InsertAt(result.Data.NewRow(), 0)

                                 'bind to the data
                                 .DataTextField = result.Data.Columns(result.LookupColumnIndex).ColumnName

                                 If result.Data.Columns(result.LookupColumnIndex).DataType Is GetType(DateTime) Then
                                    .DataTextFormatString = "{0:d}"
                                 End If
                                 control.DataSource = result.Data
                                 control.DataBind()

                                 'store the data its needed for paging, sorting
                                 Session(sID & "DATA") = result.Data

                                 'store info its needed later
                                 .Attributes.Add("LookupColumnIndex", result.LookupColumnIndex.ToString)
                                 .Attributes.Add("DataType",
                                                 result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

                                 'set the default and selected value
                                 Dim item As ListItem = control.Items.FindByValue(result.DefaultValue)
                                 If item IsNot Nothing Then
                                    control.SelectedValue = item.Value
                                 Else
                                    'The selected value is not in the list, so add it after the blank row
                                    control.Items.Insert(1, result.DefaultValue)
                                    control.SelectedIndex = 1
                                 End If
                              End If

                           End With

                           ' hidden field to hold any filter SQL code
                           tabPages(tabPage).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

                           ' Hidden Button for JS to call which fires filter click event. 
                           Dim button = New Button
                           With button
                              .ID = sID & "refresh"
                              .Style.Add("display", "none")
                           End With

                           AddHandler button.Click, AddressOf SetLookupFilter

                           tabPages(tabPage).Controls.Add(button)
                        End If

                     Case 13 ' Dropdown (13) Inputs

                        Dim control As New DropDownList

                        With control
                           .ID = sID
                           .ApplyLocation(dr)
                           .ApplySize(dr, - 1, - 1)
                           .Style.ApplyFont(dr)
                           .ApplyColor(dr)
                           If Not IsMobileBrowser() Then .ApplyBorder(False)
                           .Style.Add("padding", "1px")

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           If IsMobileBrowser() Then
                              .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                           End If

                           tabPages(tabPage).Controls.Add(control)

                           Dim filterSql = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")),
                                                           NullSafeInteger(dr("lookupFilterColumnDataType")),
                                                           NullSafeInteger(dr("LookupFilterOperator")),
                                                           FormInputPrefix & NullSafeString(dr("lookupFilterValueID")) &
                                                           "_" & NullSafeString(dr("lookupFilterValueType")) & "_")

                           If filterSql.Length > 0 Then
                              tabPages(tabPage).Controls.Add(New HiddenField _
                                                               With {.ID = "lookup" & sID, .Value = filterSql})
                           End If

                           If Not IsPostBack Then
                              'get the data
                              Dim result = db.GetWorkflowItemValues(CInt(NullSafeString(dr("id"))), _instanceID)

                              'insert a blank row
                              result.Data.Rows.InsertAt(result.Data.NewRow(), 0)

                              'bind data to grid
                              For Each column As DataColumn In result.Data.Columns
                                 If Not column.ColumnName.StartsWith("ASRSys") Then
                                    .DataTextField = column.ColumnName
                                 End If
                              Next
                              .DataSource = result.Data
                              .DataBind()

                              'store info its needed later
                              .Attributes.Add("LookupColumnIndex", result.LookupColumnIndex.ToString)
                              .Attributes.Add("DataType",
                                              result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

                              'set the default value
                              Dim item As ListItem = control.Items.FindByValue(result.DefaultValue)
                              If item IsNot Nothing Then
                                 .SelectedValue = item.Value
                              End If

                           End If

                        End With

                     Case 15 ' OptionGroup

                        Dim top = NullSafeInteger(dr("TopCoord"))
                        Dim left = NullSafeInteger(dr("LeftCoord"))
                        Dim width = NullSafeInteger(dr("Width"))
                        Dim height = NullSafeInteger(dr("Height"))
                        Dim fontAdjustment = CInt(CInt(dr("FontSize"))*0.8)
                        Dim borderCss As String

                        Dim radioTop As Int32

                        If Not NullSafeBoolean(dr("PictureBorder")) Then
                           borderCss = "border-style: none;"
                           radioTop = 2
                        Else
                           borderCss = "border: 1px solid #999;"
                           width -= 2
                           height -= 2

                           If NullSafeString(dr("caption")).Trim.Length = 0 Then
                              top += fontAdjustment
                              height -= fontAdjustment
                           End If

                           radioTop = 19 + CInt((NullSafeInteger(dr("FontSize")) - 8)*1.375)

                           If IsAndroidBrowser() AndAlso NullSafeInteger(dr("Orientation")) = 0 Then
                              radioTop -= 5
                           End If
                        End If

                        Dim html = "<fieldset style='" &
                                   " position: absolute; " &
                                   " top: " & top & "px; " &
                                   " left: " & left & "px; " &
                                   " width: " & width & "px; " &
                                   " height: " & height & "px; " &
                                   " " & GetFontCss(dr) &
                                   " " & GetColorCss(dr) &
                                   " " & borderCss &
                                   " '>"

                        If NullSafeBoolean(dr("PictureBorder")) And (NullSafeString(dr("caption")).Trim.Length > 0) Then
                           html += String.Format("<legend>{0}</legend>", NullSafeString(dr("caption"))) & vbCrLf
                        End If

                        html += "</fieldset>" & vbCrLf

                        tabPages(tabPage).Controls.Add(New LiteralControl(html))

                        Dim radioList As New RadioButtonList
                        With radioList
                           .ID = sID
                           .Style.ApplyFont(dr)
                           .CssClass = "radioList"
                           If IsAndroidBrowser() Then .CssClass += " android"

                           .TabIndex = NullSafeShort(dr("tabIndex"))
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID & "_0")

                           .RepeatDirection =
                              If _
                                 (NullSafeInteger(dr("Orientation")) = 0, RepeatDirection.Vertical,
                                  RepeatDirection.Horizontal)

                           .Style("position") = "absolute"
                           .Style("top") = Unit.Pixel(radioTop + NullSafeInteger(dr("TopCoord"))).ToString
                           .Style("left") = Unit.Pixel(9 + NullSafeInteger(dr("LeftCoord"))).ToString
                           .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 12)
                        End With

                        tabPages(tabPage).Controls.Add(radioList)

                        If Not IsPostBack Then

                           'get the data
                           Dim result = db.GetWorkflowItemValues(CInt(NullSafeString(dr("id"))), _instanceID)

                           'bind to the data
                           radioList.DataTextField = result.Data.Columns(0).ColumnName
                           radioList.DataSource = result.Data
                           radioList.DataBind()

                           'set the default value
                           radioList.SelectedValue = result.DefaultValue

                           If radioList.SelectedIndex = - 1 Then
                              radioList.SelectedIndex = 0
                           End If

                        End If

                        If IsMobileBrowser() Then
                           For Each item As ListItem In radioList.Items
                              item.Attributes.Add("onchange", "FilterMobileLookup('" & sID & "');")
                           Next
                        End If

                     Case 17 ' Input value - file upload

                        Dim control = New HtmlInputButton
                        With control
                           .ID = sID
                           .Style.ApplyLocation(dr)
                           .Style.ApplySize(dr)
                           .Style.ApplyFont(dr)

                           .Attributes.Add("TabIndex", NullSafeInteger(dr("tabIndex")).ToString)
                           UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                           ' stops the mobiles displaying buttons with over-rounded corners...
                           If IsMobileBrowser() OrElse IsMacSafari() Then
                              .Style.Add("-webkit-appearance", "none")
                              .Style.Add("background-color", "#E6E6E6")
                              .Style.Add("border", "solid 1px #CCC")
                              .Style.Add("border-radius", "4px")
                           End If

                           If _
                              NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso
                              NullSafeInteger(dr("BackColor")) <> - 2147483633 Then
                              .Style.Add("background-color",
                                         General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
                              .Style.Add("border", "solid 1px #CCC")
                              .Style.Add("border-radius", "4px")
                           End If

                           If NullSafeInteger(dr("ForeColor")) <> 6697779 Then
                              .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))).ToString)
                           End If

                           .Style.Add("padding", "0px")
                           .Style.Add("white-space", "normal")

                           .Value = NullSafeString(dr("caption"))

                           Dim crypt As New Crypt,
                               sEncodedID As String = crypt.SimpleEncrypt(NullSafeString(dr("id")).ToString,
                                                                          Session.SessionID)

                           If Not IsMobileBrowser() Then
                              .Attributes.Add("onclick",
                                              "try{showFileUpload(true, '" & sEncodedID &
                                              "', document.getElementById('file" & sID & "').value);}catch(e){};")
                           Else
                              .Attributes.Add("onclick",
                                              "try{alert('Your browser does not support file upload.');}catch(e){};")
                           End If
                        End With

                        tabPages(tabPage).Controls.Add(control)

                        tabPages(tabPage).Controls.Add(New HiddenField _
                                                         With {.ID = "file" & sID, .Value = NullSafeString(dr("value"))})

                     Case 19, 20 ' DB File or WF File

                        Dim crypt As New Crypt,
                            sEncodedID As String = crypt.SimpleEncrypt(NullSafeString(dr("id")).ToString,
                                                                       Session.SessionID)

                        Dim html = "<span id='" & sID & "' tabindex=" & NullSafeInteger(dr("tabIndex")).ToString &
                                   " style='position: absolute; display:inline-block; word-wrap:break-word; overflow:auto;" &
                                   " top: " & NullSafeString(dr("TopCoord")) & "px;" &
                                   " left: " & NullSafeString(dr("LeftCoord")) & "px;" &
                                   " height:" & NullSafeString(dr("Height")) & "px;" &
                                   " width:" & NullSafeInteger(dr("Width")) & "px;" &
                                   " " & GetFontCss(dr) &
                                   " " & GetColorCss(dr) &
                                   "'" &
                                   " onclick='FileDownload_Click(""" & sEncodedID & """);'" &
                                   " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" &
                                   " >" &
                                   HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) &
                                   "</span>"

                        UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                        tabPages(tabPage).Controls.Add(New LiteralControl(html))

                     Case 21 ' Tab Strip

                        'split out the tab names to calculate number of tabs - may not have loaded all tabs yet, so can't count them.
                        Dim arrTabCaptions As String() = NullSafeString(dr("Caption")).Split(New Char() {";"c})

                        pnlTabsDiv.Style("width") = CStr(dr("Width")) & "px"
                        pnlTabsDiv.Style("height") = CStr(dr("Height")) & "px"
                        pnlTabsDiv.Style("left") = CStr(dr("LeftCoord")) & "px"
                        pnlTabsDiv.Style("top") = CStr(dr("TopCoord")) & "px"

                        Dim ctlTabsDiv As New Panel
                        ctlTabsDiv.ID = "TabsDiv"
                        ctlTabsDiv.Style.Add("height", TabStripHeight & "px")
                        ctlTabsDiv.Style.Add("position", "relative")
                        ctlTabsDiv.Style.Add("z-index", "1")

                        If IsMobileBrowser() And Not IsAndroidBrowser() Then
                           ctlTabsDiv.Style.Add("overflow-x", "auto")
                        Else
                           ' for non-mobile browsers we display arrows to scroll the tab bar left and right.
                           ctlTabsDiv.Style.Add("overflow", "hidden")
                           ctlTabsDiv.Style.Add("margin-right", "51px")

                           ' Nav arrows for non-mobile browsers
                           Dim ctlFormTabArrows As New Panel
                           With ctlFormTabArrows
                              .Style.Add("position", "absolute")
                              .Style.Add("top", "0px")
                              .Style.Add("right", "0px")
                              .Style.Add("width", "48px")
                              .Style.Add("z-index", "1")
                              .BackColor = Color.White
                              .BorderColor = Color.Black
                              .BorderWidth = 1
                           End With

                           ' Left scroll arrow
                           Dim image = New WebControls.Image
                           With image
                              .Style.Add("width", "24px")
                              .Style.Add("height", TabStripHeight - 2 & "px")
                              .ImageUrl = "~/Images/tab-prev.gif"
                              .Style.Add("margin", "0px")
                              .Style.Add("padding", "0px")
                              .Attributes.Add("onclick",
                                              "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft - 20;")
                           End With
                           ctlFormTabArrows.Controls.Add(image)

                           ' Right scroll arrow
                           image = New WebControls.Image
                           With image
                              .Style.Add("width", "24px")
                              .Style.Add("height", TabStripHeight - 2 & "px")
                              .ImageUrl = "~/Images/tab-next.gif"
                              .Style.Add("margin", "0px")
                              .Style.Add("padding", "0px")
                              .Attributes.Add("onclick",
                                              "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft + 20;")
                           End With
                           ctlFormTabArrows.Controls.Add(image)

                           pnlTabsDiv.Controls.Add(ctlFormTabArrows)
                        End If

                        ' generate the tabs.
                        Dim ctlTabsTable As New Table
                        ctlTabsTable.CellSpacing = 0
                        ' ctlTabsTable.Style.Add("margin-top", "2px")
                        Dim trPager As TableRow = New TableRow()
                        trPager.Height = Unit.Pixel(TabStripHeight - 1)
                        ' to prevent vertical scrollbar
                        trPager.Style.Add("white-space", "nowrap")

                        Dim iTabNo As Integer = 1
                        ' add a cell for each tab
                        For Each sTabCaption In arrTabCaptions
                           If sTabCaption.Trim.Length > 0 Then
                              Dim tcTabCell As TableCell = New TableCell

                              With tcTabCell
                                 .ID = FormInputPrefix & iTabNo.ToString & "_21_Panel"
                                 .BorderColor = Color.Black
                                 .Style.Add("padding-left", "5px")
                                 .Style.Add("padding-right", "5px")
                                 .Style.Add("border-radius", "5px 5px 0px 0px")
                                 .Style.Add("width", "50px")
                                 .BorderWidth = 1
                                 .BorderStyle = BorderStyle.Solid
                                 .BackColor = Color.White

                                 ' label the button...
                                 Dim label = New Label
                                 label.Font.Name = "Verdana"
                                 label.Font.Size = New FontUnit(11, UnitType.Pixel)
                                 label.Text = sTabCaption.ToString

                                 .Controls.Add(label)

                                 ' Tab Clicking/mouseover
                                 .Attributes.Add("onclick", "SetCurrentTab(" & iTabNo.ToString & ");")
                                 .Attributes.Add("onmouseover", "this.style.cursor='pointer';")
                                 .Attributes.Add("onmouseout", "this.style.cursor='';")
                              End With

                              trPager.Cells.Add(tcTabCell)

                              ' NPG20120321 Fault HRPRO-2113
                              ' Rather than put the controls div inside the relevant tab page (issues with referencing the AJAX controls on postback), 
                              ' we move the controls div into the form by the top and left of the tabstrip, if it exists

                              If iTabNo > 0 Then ' Tab 0 is the base page.

                                 ' create any MISSING tabs...
                                 Try
                                    Dim strTemp As String = tabPages(iTabNo).ID.ToString
                                    ' OK, if the id exists, the div has already been created. Do nothing.
                                 Catch ex As Exception
                                    ' Otherwise create the div
                                    ' Create the new div, give it a unique id then we can refer to that when it's reused in the next loop.
                                    ' store the id in the array for reference. NB 21 is the itemtype for a page Tab
                                    If iTabNo > tabPages.GetUpperBound(0) Then ReDim Preserve tabPages(iTabNo)

                                    tabPages(iTabNo) = New Panel
                                    tabPages(iTabNo).ID = FormInputPrefix & iTabNo.ToString & "_21_PageTab"
                                    tabPages(iTabNo).Style.Add("position", "absolute")

                                    ' Add this tab to the web form
                                    pnlInputDiv.Controls.Add(tabPages(iTabNo))
                                 End Try

                                 ' Move all tabs to their relative position within the tab frame.
                                 Try
                                    tabPages(iTabNo).Style.Add("top",
                                                               NullSafeInteger(dr("TopCoord")) + TabStripHeight & "px")
                                    tabPages(iTabNo).Style.Add("left", NullSafeInteger(dr("LeftCoord")) & "px")

                                    ' Hide all tabs but the first.
                                    tabPages(iTabNo).Style.Add("display", "none")
                                 Catch ex As Exception

                                 End Try
                              End If

                              iTabNo += 1
                              ' keep tabs on the number of tabs hehehe :P
                           End If
                        Next

                        'add row to table
                        ctlTabsTable.Rows.Add(trPager)

                        'add table to div
                        ctlTabsDiv.Controls.Add(ctlTabsTable)
                        pnlTabsDiv.Controls.AddAt(0, ctlTabsDiv)

                  End Select
               End While

               dr.Close()
               _form.Connection.Close()

               If (Not ClientScript.IsStartupScriptRegistered("Startup")) Then
                  ' Form the script to be registered at client side.
                  scriptString += "}"
                  ClientScript.RegisterStartupScript(ClientScript.GetType, "Startup", scriptString, True)
               End If

               If sMessage.Length = 0 Then

                  If _form.ErrorMessage <> "" Then
                     sMessage = _form.ErrorMessage
                  End If

                  If _form.BackImage > 0 Then
                     Dim image As String = LoadPicture(_form.BackImage, sMessage)
                     If sMessage.Length = 0 Then
                        divInput.Style("background-image") = image
                        divInput.Style("background-repeat") = General.BackgroundRepeat(_form.BackImageLocation)
                        divInput.Style("background-position") = General.BackgroundPosition(_form.BackImageLocation)
                     End If
                  End If

                  If _form.BackColour > 0 Then
                     divInput.Style("background-color") = General.GetHtmlColour(_form.BackColour)
                  End If

                  pnlInputDiv.Style("width") = _form.Width.ToString & "px"
                  pnlInputDiv.Style("height") = _form.Height.ToString & "px"
                  pnlInputDiv.Style("left") = "-2px"

                  'TODO PG dont need can get size in jQuery from pnlInputDiv css
                  hdnFormWidth.Value = _form.Width.ToString
                  hdnFormHeight.Value = _form.Height.ToString

                  hdnSiblingForms.Value = sSiblingForms.ToString
               End If

            End If

            ' Resize the mobile 'viewport' to fit the webform
            AddHeaderTags(_form.Width)

         Catch ex As Exception
            sMessage = "Error loading web form controls:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" &
                       "Contact your system administrator."
         End Try

      End If

      If sMessage.Length > 0 Then

         If IsPostBack Then
            bulletErrors.Items.Clear()
            bulletWarnings.Items.Clear()

            hdnErrorMessage.Value = sMessage
            hdnFollowOnForms.Value = ""
            SetSubmissionMessage(sMessage & "<BR><BR>Click", "here", "to close this form.")
         Else
            Session("message") = sMessage
            Response.Redirect("Message.aspx")
         End If
      End If
   End Sub

   Private Function SetSubmissionMessage(message As String) As Boolean
      Dim m1 = "", m2 = "", m3 = ""
      Dim result As Boolean = General.SplitMessage(message, m1, m2, m3)
      If result Then SetSubmissionMessage(m1, m2, m3)
      Return result
   End Function

   Private Sub SetSubmissionMessage(message1 As String, message2 As String, message3 As String)
      hdnSubmissionMessage_1.Value = message1.Trim
      hdnSubmissionMessage_2.Value = message2.Trim
      hdnSubmissionMessage_3.Value = message3.Trim
      hdnNoSubmissionMessage.Value = If(message1.Length + message2.Length + message3.Length = 0, "1", "0")
   End Sub

   Private Sub GetControls(controlCollection As ControlCollection, result As ICollection(Of Control),
                           Optional predicate As Func(Of Control, Boolean) = Nothing)

      For Each c As Control In controlCollection
         If predicate Is Nothing OrElse predicate(c) Then
            result.Add(c)
         End If
         If c.HasControls Then
            GetControls(c.Controls, result, predicate)
         End If
      Next
   End Sub

   Public Sub ButtonClick(ByVal sender As System.Object, ByVal e As EventArgs)

      Dim db As New Database(GetConnectionString)
      Dim valueString As String = ""
      Dim sMessage As String = ""

      Try
         ' Read the web form item values & build up a string of the form input values.
         ' This is a tab delimited string of itemIDs and values.
         Dim controlList As New List(Of Control)
         GetControls(Page.Controls, controlList, Function(c) c.ClientID.StartsWith(FormInputPrefix) AndAlso
                                                             (c.ClientID.EndsWith("_") OrElse
                                                              c.ClientID.EndsWith("TextBox") OrElse
                                                              c.ClientID.EndsWith("Grid")))

         For Each ctlFormInput As Control In controlList

            Dim sIDString As String = ctlFormInput.ID.Substring(Len(FormInputPrefix))
            Dim iTemp As Int16 = CShort(sIDString.IndexOf("_"))
            Dim sTemp As String = sIDString.Substring(iTemp + 1)
            sIDString = sIDString.Substring(0, iTemp) & vbTab
            iTemp = CShort(sTemp.IndexOf("_"))
            Dim sType As String = sTemp.Substring(0, iTemp)
            Dim value As String

            Select Case CShort(sType)

               Case 0 ' Button

                  Dim btn As HtmlInputButton = DirectCast(sender, HtmlInputButton)

                  If (ctlFormInput.ID = btn.ID) Then
                     hdnLastButtonClicked.Value = btn.ID
                     valueString += sIDString & "1" & vbTab
                  ElseIf (TypeOf ctlFormInput Is HtmlInputButton) Then
                     valueString += sIDString & "0" & vbTab
                  End If

               Case 3 ' Character Input

                  If TypeOf ctlFormInput Is TextBox Then
                     value = DirectCast(ctlFormInput, TextBox).Text.Replace(vbTab, " ")
                     valueString += sIDString & value & vbTab
                  End If

               Case 5 ' Numeric Input

                  If TypeOf ctlFormInput Is TextBox Then
                     Dim control = DirectCast(ctlFormInput, TextBox)
                     value =
                        If _
                           (CSng(control.Text) = CSng(0), "0",
                            control.Text.Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator,
                                                 "."))
                     valueString += sIDString & value & vbTab
                  End If

               Case 6 ' Logic Input

                  If TypeOf ctlFormInput Is CheckBox Then
                     value = If(DirectCast(ctlFormInput, CheckBox).Checked, "1", "0")
                     valueString += sIDString & value & vbTab
                  End If

               Case 7 ' Date Input

                  If TypeOf ctlFormInput Is TextBox Then
                     Dim control = DirectCast(ctlFormInput, TextBox)
                     value = If(control.Text.Trim = "", "null", DateTime.Parse(control.Text).ToString("MM/dd/yyyy"))
                     valueString += sIDString & value & vbTab
                  End If

               Case 11 ' Grid (RecordSelector) Input
                  If TypeOf ctlFormInput Is RecordSelector Then
                     Dim control = DirectCast(ctlFormInput, RecordSelector)
                     value = If(control.SelectedValue IsNot Nothing, CStr(control.SelectedValue), "0")
                     valueString += sIDString & value & vbTab
                  End If

               Case 13 ' Dropdown Input

                  If TypeOf ctlFormInput Is DropDownList Then
                     value = DirectCast(ctlFormInput, DropDownList).Text
                     valueString += sIDString & value & vbTab
                  End If

               Case 14 ' Lookup Input

                  If Not IsMobileBrowser() Then

                     If TypeOf ctlFormInput Is TextBox Then
                        Dim control = DirectCast(ctlFormInput, TextBox)

                        If control.Attributes("DataType") = "System.DateTime" Then
                           value = If(control.Text = "", "null", General.ConvertLocaleDateToSql(control.Text))
                        ElseIf _
                           control.Attributes("DataType") = "System.Decimal" Or
                           control.Attributes("DataType") = "System.Int32" Then
                           value =
                              If _
                                 (control.Text = "", "",
                                  control.Text.Replace(
                                     Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, "."))
                        Else
                           value = control.Text
                        End If

                        valueString += sIDString & value & vbTab
                     End If
                  Else
                     ' Mobile Browser - it's a Dropdown List.
                     If TypeOf ctlFormInput Is DropDownList Then
                        value = DirectCast(ctlFormInput, DropDownList).Text
                        valueString += sIDString & value & vbTab
                     End If

                  End If

               Case 15 ' OptionGroup Input

                  If TypeOf ctlFormInput Is RadioButtonList Then
                     value = DirectCast(ctlFormInput, RadioButtonList).SelectedValue
                     valueString += sIDString & value & vbTab
                  End If

               Case 17 ' FileUpload

                  If TypeOf ctlFormInput Is HtmlInputButton Then
                     value = DirectCast(pnlInput.FindControl("file" & ctlFormInput.ID), HiddenField).Value
                     valueString += sIDString & value & vbTab
                  End If

            End Select

         Next

      Catch ex As Exception
         sMessage = "Error reading web form item values:<BR><BR>" & ex.Message
      End Try

      If sMessage.Length = 0 Then

         ' Validate the web form entry.
         errorMessagePanel.Font.Name = "Verdana"
         errorMessagePanel.Font.Size = _config.ValidationMessageFontSize
         errorMessagePanel.ForeColor = General.GetColour(6697779)

         bulletErrors.Items.Clear()
         bulletWarnings.Items.Clear()

         Dim result = db.WorkflowValidateWebForm(_elementID, _instanceID, valueString)

         result.Errors.ForEach(Sub(f) bulletErrors.Items.Add(f))

         If hdnOverrideWarnings.Value <> "1" Then
            result.Warnings.ForEach(Sub(f) bulletWarnings.Items.Add(f))
         End If

         hdnCount_Errors.Value = CStr(bulletErrors.Items.Count)
         hdnCount_Warnings.Value = CStr(bulletWarnings.Items.Count)
         hdnOverrideWarnings.Value = "0"

         lblErrors.Text = If(bulletErrors.Items.Count > 0,
                             "Unable to submit this form due to the following error" &
                             If(bulletErrors.Items.Count = 1, "", "s") & ":",
                             "")

         lblWarnings.Text = If(bulletWarnings.Items.Count > 0,
                               If(bulletErrors.Items.Count > 0, "And the following warning" &
                                                                If(bulletWarnings.Items.Count = 1, "", "s") & ":",
                                  "Submitting this form raises the following warning" &
                                  If(bulletWarnings.Items.Count = 1, "", "s") & ":"),
                               "")

         overrideWarning.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)

         ' Submit the webform
         If bulletWarnings.Items.Count = 0 And bulletErrors.Items.Count = 0 Then

            Try
               'TODO NOW PG why transactionscope???
               Dim submit As SubmitWebFormResult
               Using (New TransactionScope(TransactionScopeOption.Suppress))
                  submit = db.WorkflowSubmitWebForm(_elementID, _instanceID, valueString,
                                                    NullSafeInteger(hdnDefaultPageNo.Value))
               End Using

               hdnFollowOnForms.Value = ""

               If submit.SavedForLater Then
                  Select Case _form.SavedForLaterMessageType
                     Case 1 ' Custom
                        If Not SetSubmissionMessage(_form.SavedForLaterMessage) Then
                           SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here",
                                                "to close this form.")
                        End If
                     Case 2 ' None
                        SetSubmissionMessage("", "", "")
                     Case Else 'System default
                        SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here",
                                             "to close this form.")
                  End Select

               ElseIf submit.FormElements.Length = 0 Then
                  Select Case _form.CompletionMessageType
                     Case 1 ' Custom
                        If Not SetSubmissionMessage(_form.CompletionMessage) Then
                           SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to close this form.")
                        End If
                     Case 2 ' None
                        SetSubmissionMessage("", "", "")
                     Case Else 'System default
                        SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to close this form.")
                  End Select
               Else
                  Dim followOnForms As String() = submit.FormElements.
                         Split(New String() {vbTab}, StringSplitOptions.RemoveEmptyEntries).
                         Select(Function(f) db.GetWorkflowQueryString(_instanceID, CInt(f))).
                         ToArray()

                  hdnFollowOnForms.Value = String.Join(vbTab, followOnForms)

                  Select Case _form.FollowOnFormsMessageType
                     Case 1 ' Custom
                        If Not SetSubmissionMessage(_form.FollowOnFormsMessage) Then
                           SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here",
                                                "to complete the follow-on Workflow form" &
                                                If(followOnForms.Count = 1, "", "s") & ".")
                        End If
                     Case 2 ' None
                        SetSubmissionMessage("", "", "")
                     Case Else 'System default
                        SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here",
                                             "to complete the follow-on Workflow form" &
                                             If(followOnForms.Count = 1, "", "s") & ".")
                  End Select
               End If

            Catch ex As Exception
               sMessage = "Error submitting the web form:<BR><BR>" & ex.Message
            End Try

         End If

      End If

      If sMessage.Length > 0 Then
         bulletErrors.Items.Clear()
         bulletWarnings.Items.Clear()

         hdnErrorMessage.Value = sMessage
         hdnFollowOnForms.Value = ""
         SetSubmissionMessage(sMessage & "<BR><BR>Click", "here", "to close this form.")
      End If
   End Sub

   Private _minTabIndex As Short = - 1

   Private Sub UpdateAutoFocusControl(tabIndex As Short, focusId As String)
      If _minTabIndex < 0 Or tabIndex < _minTabIndex Then
         _autoFocusControl = focusId
         _minTabIndex = tabIndex
      End If
   End Sub

   Public Function LocaleDateFormat() As String
      Return Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern.ToUpper
   End Function

   Public Function LocaleDateFormatjQuery() As String
      'jQuery date formats are different to .NET's
      Return LocaleDateFormat.ToLower.Replace("yyyy", "yy")
   End Function

   Public Function LocaleDecimal() As String
      Return Thread.CurrentThread.CurrentUICulture.NumberFormat.NumberDecimalSeparator
   End Function

   Public Function AndroidLayerBug() As Boolean
      Return IsAndroidBrowser()
   End Function

   Public Function IsMobileBrowser() As Boolean
      Return Utilities.IsMobileBrowser()
   End Function

   Public Function AutoFocusControl() As String
      Return _autoFocusControl
   End Function

   Public Function ColourThemeHex() As String
      Return _config.ColourThemeHex
   End Function

   Private Function GetConnectionString() As String
      Dim connectionString = "Application Name=OpenHR Workflow;Data Source=" & _server & ";Initial Catalog=" & _database &
                             ";Integrated Security=false;User ID=" & _user & ";Password=" & _password & ";Pooling=false"
      Return connectionString
   End Function

   Private Function LoadPicture(ByVal piPictureID As Int32, ByRef psErrorMessage As String) As String

      Dim conn As SqlConnection
      Dim cmdSelect As SqlCommand
      Dim dr As SqlDataReader
      Dim sImageFileName As String
      Dim sImageFilePath As String
      Dim sTempName As String
      Dim fs As IO.FileStream
      Dim bw As IO.BinaryWriter
      Const iBufferSize As Integer = 100
      Dim outByte(iBufferSize - 1) As Byte
      Dim retVal As Long
      Dim startIndex As Long
      Dim sExtension As String = ""
      Dim iIndex As Integer
      Dim sName As String

      Try
         _imageCount += 1

         psErrorMessage = ""
         sImageFileName = ""
         sImageFilePath = Server.MapPath("pictures")

         conn = New SqlConnection(GetConnectionString)
         conn.Open()

         cmdSelect = New SqlCommand
         cmdSelect.CommandText = "spASRGetPicture"
         cmdSelect.Connection = conn
         cmdSelect.CommandType = CommandType.StoredProcedure
         cmdSelect.CommandTimeout = _config.SubmissionTimeoutInSeconds

         cmdSelect.Parameters.Add("@piPictureID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmdSelect.Parameters("@piPictureID").Value = piPictureID

         Try
            dr = cmdSelect.ExecuteReader(CommandBehavior.SequentialAccess)

            Do While dr.Read
               sName = NullSafeString(dr("name"))
               iIndex = sName.LastIndexOf(".")
               If iIndex >= 0 Then
                  sExtension = sName.Substring(iIndex)
               End If

               sImageFileName = Session.SessionID().ToString &
                                "_" & _imageCount.ToString &
                                "_" & Date.Now.Ticks.ToString &
                                sExtension
               sTempName = sImageFilePath & "\" & sImageFileName

               ' Create a file to hold the output.
               fs = New IO.FileStream(sTempName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
               bw = New IO.BinaryWriter(fs)

               ' Reset the starting byte for a new BLOB.
               startIndex = 0

               ' Read bytes into outbyte() and retain the number of bytes returned.
               retVal = dr.GetBytes(1, startIndex, outByte, 0, iBufferSize)

               ' Continue reading and writing while there are bytes beyond the size of the buffer.
               Do While retVal = iBufferSize
                  bw.Write(outByte)
                  bw.Flush()

                  ' Reposition the start index to the end of the last buffer and fill the buffer.
                  startIndex += iBufferSize
                  retVal = dr.GetBytes(1, startIndex, outByte, 0, iBufferSize)
               Loop

               ' Write the remaining buffer.
               bw.Write(outByte)
               bw.Flush()

               ' Close the output file.
               bw.Close()
               fs.Close()
            Loop

            dr.Close()
            cmdSelect.Dispose()

            ' Ensure URL encoding doesn't stuff up the picture name, so encode the % character as %25.
            LoadPicture = "pictures/" & sImageFileName

         Catch ex As Exception
            LoadPicture = ""
            psErrorMessage = ex.Message

         Finally
            conn.Close()
            conn.Dispose()
         End Try
      Catch ex As Exception
         LoadPicture = ""
         psErrorMessage = ex.Message
      End Try
   End Function

   Private Function LookupFilterSQL(ByVal psColumnName As String, ByVal piColumnDataType As Integer,
                                    ByVal piOperatorID As Integer, ByVal psValue As String) As String

      Dim filterSql As String = ""

      Try
         If (psColumnName.Length > 0) And (piOperatorID > 0) And (psValue.Length > 0) Then

            Select Case piColumnDataType
               Case SqlDataType.Boolean
                  Select Case piOperatorID
                     Case FilterOperators.giFILTEROP_EQUALS
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab
                     Case FilterOperators.giFILTEROP_NOTEQUALTO
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab
                  End Select

               Case SqlDataType.Numeric, SqlDataType.Integer
                  Select Case piOperatorID
                     Case FilterOperators.giFILTEROP_EQUALS
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab

                     Case FilterOperators.giFILTEROP_NOTEQUALTO
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab

                     Case FilterOperators.giFILTEROP_ISATMOST
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) <= " & vbTab

                     Case FilterOperators.giFILTEROP_ISATLEAST
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) >= " & vbTab

                     Case FilterOperators.giFILTEROP_ISMORETHAN
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) > " & vbTab

                     Case FilterOperators.giFILTEROP_ISLESSTHAN
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], 0) < " & vbTab
                  End Select

               Case SqlDataType.Date
                  Select Case piOperatorID
                     Case FilterOperators.giFILTEROP_ON
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

                     Case FilterOperators.giFILTEROP_NOTON
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

                     Case FilterOperators.giFILTEROP_ONORBEFORE
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "LEN(ISNULL([ASRSysLookupFilterValue], '')) = 0 OR (LEN('" & vbTab &
                                    "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') <= '" & vbTab & "')"

                     Case FilterOperators.giFILTEROP_ONORAFTER
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab &
                                    "') = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') >= '" &
                                    vbTab & "')"

                     Case FilterOperators.giFILTEROP_AFTER
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "(LEN('" & vbTab &
                                    "') = 0 AND LEN(ISNULL([ASRSysLookupFilterValue], '')) > 0) OR (LEN('" & vbTab &
                                    "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') > '" & vbTab & "')"

                     Case FilterOperators.giFILTEROP_BEFORE
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab &
                                    "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') < '" & vbTab & "'"
                  End Select

               Case SqlDataType.VarChar, SqlDataType.VarBinary, SqlDataType.LongVarChar
                  Select Case piOperatorID
                     Case FilterOperators.giFILTEROP_IS
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

                     Case FilterOperators.giFILTEROP_ISNOT
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

                     Case FilterOperators.giFILTEROP_CONTAINS
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab &
                                    "ISNULL([ASRSysLookupFilterValue], '') LIKE '%" & vbTab & "%'"

                     Case FilterOperators.giFILTEROP_DOESNOTCONTAIN
                        filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab &
                                    "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') NOT LIKE '%" & vbTab & "%'"
                  End Select
            End Select
         End If

      Catch ex As Exception
      End Try

      Return filterSql
   End Function

   Private Sub ShowNoResultFound(ByVal source As DataTable, ByVal gv As RecordSelector)

      source.Clear()
      source.Rows.Add(source.NewRow())
      ' create a new blank row to the DataTable
      ' Bind the DataTable which contain a blank row to the GridView
      gv.DataSource = source
      gv.DataBind()
      ' Get the total number of columns in the GridView to know what the Column Span should be
      Dim columnsCount As Integer = gv.Columns.Count
      gv.Rows(0).Cells.Clear()
      ' clear all the cells in the row
      gv.Rows(0).Cells.Add(New TableCell())
      'add a new blank cell
      gv.Rows(0).Cells(0).ColumnSpan = columnsCount
      'set the column span to the new added cell

      'You can set the styles here
      gv.Rows(0).Cells(0).HorizontalAlign = HorizontalAlign.Center
      'set No Results found to the new added cell
      gv.Rows(0).Cells(0).Text = gv.EmptyDataText

      gv.SelectedIndex = - 1
   End Sub

   Protected Sub BtnDoFilterClick(sender As Object, e As EventArgs) Handles btnDoFilter.Click

      For Each value As String In hdnMobileLookupFilter.Value.Split(CChar(vbTab))
         SetLookupFilter(Nothing, Nothing, value)
      Next
   End Sub

   Sub SetLookupFilter(ByVal sender As Object, ByVal e As EventArgs, Optional lookupID As String = "")

      If sender IsNot Nothing Then
         ' get button's ID
         lookupID = DirectCast(sender, Button).ID
      End If

      If lookupID.Length = 0 Then Return

      ' Create a datatable from the data in the session variable
      Dim dataTable As DataTable = TryCast(HttpContext.Current.Session(lookupID.Replace("refresh", "DATA")), DataTable)

      ' get the filter sql
      Dim hiddenField As HiddenField = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "filterSQL")),
                                               HiddenField)

      Dim filterSql As String = hiddenField.Value

      If TypeOf (pnlInputDiv.FindControl(lookupID.Replace("refresh", ""))) Is DropDownList Then

         ' This is a dropdownlist style lookup (mobiles only)
         Dim dropdown As DropDownList = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "")), DropDownList)

         ' Store the current value, so we can re-add it after filtering.
         Dim strCurrentSelection As String = dropdown.Text

         ' Filter the table now.
         FilterDataTable(dataTable, filterSql)

         ' insert the previously selected item
         Dim objDataRow As DataRow = dataTable.NewRow()
         objDataRow(0) = strCurrentSelection
         dataTable.Rows.InsertAt(objDataRow, 0)

         ' Rebind the new datatable
         dropdown.DataSource = dataTable
         dropdown.DataBind()

         ' Insert empty row at top of list
         objDataRow = dataTable.NewRow()
         dataTable.Rows.InsertAt(objDataRow, 0)

         ' reset filter.
         hiddenField.Value = ""
      Else
         ' This is a normal grid lookup (not Mobile)
         FilterDataTable(dataTable, filterSql)

         Dim gridView As RecordSelector = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "Grid")),
                                                  RecordSelector)

         gridView.filterSQL = filterSql.ToString
         gridView.DataSource = dataTable
         gridView.DataBind()
      End If

      ' reset filter.
      hiddenField.Value = ""
   End Sub

   Private Sub FilterDataTable(ByRef dataTable As DataTable, ByVal filterSql As String)
      If dataTable IsNot Nothing Then
         Dim dataView As New DataView(dataTable)
         dataView.RowFilter = filterSql

         dataTable = dataView.ToTable()

         If dataTable.Rows.Count < 2 Then
            ' create a blank row to display.
            Dim objDataRow As DataRow
            objDataRow = dataTable.NewRow()
            dataTable.Rows.InsertAt(objDataRow, 0)
         End If
      End If
   End Sub

   Private Sub AddHeaderTags(ByVal lngViewportWidth As Long)

      ' Create the following timeout meta tag programatically for all browsers
      '    <meta http-equiv="refresh" content="5; URL=timeout.aspx" />
      Dim meta As New HtmlMeta()
      meta.HttpEquiv = "refresh"
      meta.Content = (Session.Timeout*60).ToString & "; URL=timeout.aspx"

      Page.Header.Controls.Add(meta)

      ' for Mobiles only, set the viewport and 'home page' icons
      If IsMobileBrowser() Then
         meta = New HtmlMeta()
         meta.Name = "viewport"
         meta.Content = "width=" & lngViewportWidth & ", user-scalable=yes"
         Page.Header.Controls.Add(meta)

         Dim link As New HtmlLink()
         link.Attributes("rel") = "apple-touch-icon"
         link.Href = "favicon.ico"
         Page.Header.Controls.Add(link)
      End If
   End Sub

   Private Sub SetCulture()

      Dim culture As String

      If Request.UserLanguages IsNot Nothing Then
         culture = Request.UserLanguages(0)
      ElseIf Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") IsNot Nothing Then
         culture = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
      Else
         culture = ConfigurationManager.AppSettings("defaultculture")
      End If

      If culture.ToLower = "en-us" Then culture = "en-GB"

      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(culture)
      Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(culture)
   End Sub
End Class
