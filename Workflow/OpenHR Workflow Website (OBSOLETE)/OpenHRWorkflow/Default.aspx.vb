Option Strict On

Imports System.Data
Imports System.Threading
Imports System.Globalization
Imports System.Reflection
Imports System.Drawing
Imports AjaxControlToolkit
Imports System.Transactions
Imports System.Data.SqlClient
Imports System.IO

Public Class [Default]
   Inherits Page

   Private _config As Config
   Private _url As WorkflowUrl
   Private _form As WorkflowForm
   Private _db As Database
   Private _imageCount As Integer
   Private _siblingForms As String = ""
   Private _minTabIndex As Short?
   Private _autoFocusControl As String

   Private Const TabStripHeight As Integer = 21
   Private Const FormInputPrefix As String = "FI_"
   Private Const DefaultTitle As String = "OpenHR Workflow"
   Private Const CompiledAssemblyName As String = "OPENHRWORKFLOW"

   Protected Sub Page_PreInit(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.PreInit

      Dim message As String = ""

      'The script manager calls this page to get it combined js files, if the calls is from there ignore it
      If Request.QueryString.Count > 1 Then
         Return
      End If

      'Page requested with no workflow details, just redirect to the login page
      If Request.QueryString.Count = 0 Then
         Response.Redirect("~/Account/Login.aspx", True)
      End If

      'Extract the workflow details from the url (use the rawUrl rather than queryString)
      'as some characters are ignored in the queryString
      Dim query = Server.UrlDecode(Request.RawUrl)
      query = query.Substring(query.IndexOf("?") + 1)

      _url = New WorkflowUrl

      Try
         'Try the latest encryption method
         'Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
         Dim currentCulture = Thread.CurrentThread.CurrentCulture

         Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
         Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-GB")

         Dim crypt As New Crypt
         query = crypt.DecompactString(query)
         query = crypt.DecryptString(query, "", True)

         'Reset the culture to be the one used by the client. Fault HRPRO-1404
         Thread.CurrentThread.CurrentCulture = currentCulture
         Thread.CurrentThread.CurrentUICulture = currentCulture

         'Extract the required parameters from the decrypted queryString.
         Dim values = query.Split(vbTab(0))

         _url.InstanceID = CInt(values(0))
         _url.ElementID = CInt(values(1))
         _url.User = values(2)
         _url.Password = values(3)
         _url.Server = values(4)
         _url.Database = values(5)
         If values.Count > 6 Then _url.UserName = values(6)

      Catch ex As Exception
         'Try the older encryption method
         Try
            Dim crypt As New Crypt
            query = crypt.ProcessDecryptString(query)
            query = crypt.DecryptString(query, "", False)

            Dim values = query.Split(vbTab(0))

            If _url.InstanceID = 0 Then _url.InstanceID = CInt(values(0))
            If _url.ElementID = 0 Then _url.ElementID = CInt(values(1))
            _url.User = values(2)
            _url.Password = values(3)
            _url.Server = values(4)
            _url.Database = values(5)
         Catch exx As Exception
            message = "Invalid query string."
         End Try
      End Try

      _db = New Database(GetConnectionString)

      'TODO PG prob have to check db lock & version before doing below

      'Code below is only for initial page load not postbacks
      If Not IsPostBack Then

         'Activating mobile security. I've hijacked the _instanceID and populated it with the User ID that is to be activated.
         If message.Length = 0 And _url.ElementID = -2 And _url.InstanceID > 0 Then

            message = _db.ActivateUser(_url.InstanceID)

            If message.Length = 0 Then
               message = "You have been successfully activated."
            End If
         End If

         'Initiate the workflow if thats whats required
         If message.Length = 0 And _url.InstanceID < 0 And _url.ElementID = -1 Then

            Dim result As InstantiateWorkflowResult = _db.InstantiateWorkflow(-_url.InstanceID, _url.UserName)

            If result.Message <> "" Then
               message = "Error:<BR><BR>" & result.Message
            Else
               If result.FormElements.Length = 0 Then
                  message = "Workflow initiated successfully."
               Else
                  'The first form element is this workflow and any others are sibling forms (that need to be opened at the same time)
                  Dim forms = result.FormElements.Split(New String() {vbTab}, StringSplitOptions.RemoveEmptyEntries).ToList

                  _url.InstanceID = result.InstanceId
                  _url.ElementID = CInt(forms(0))
                  forms.RemoveAt(0)

                  _siblingForms = String.Join(vbTab, forms.Select(Function(f) _db.GetWorkflowQueryString(_url.InstanceID, CInt(f))))

                  Dim crypt As New Crypt
                  Dim newUrl = crypt.EncryptQueryString(_url.InstanceID, _url.ElementID, _url.User, _url.Password, _url.Server, _url.Database, "", "")
                  Response.Redirect("~/Default.aspx?" & newUrl, True)
               End If
            End If
         End If

      End If

      If message.Length > 0 Then
         Session("message") = message
         Response.Redirect("Message.aspx")
      End If
   End Sub

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

      Dim message As String = ""
      Dim assemblyName As String = ""
      Dim webSiteVersion As String = ""

      Try
         Response.CacheControl = "no-cache"
         Response.AddHeader("Pragma", "no-cache")
         Response.Expires = -1

         _config = New Config()
         _config.Initialise(Server.MapPath("themes/ThemeHex.xml"))

      Catch ex As Exception
      End Try

      'Set the page title
      Dim pageTitle As String = DefaultTitle

      'TODO PG NOW can this ever be zero length see other code
      With Assembly.GetExecutingAssembly.GetName.Version
         webSiteVersion = String.Format("{0}.{1}.{2}", .Major, .Minor, .Build)
      End With

      assemblyName = Assembly.GetExecutingAssembly.GetName.Name.ToUpper

      If assemblyName = CompiledAssemblyName Then
         'Compiled version of the web site, so perform version checks.
         pageTitle += " - v" & webSiteVersion
      Else
         'Development version of the web site
         pageTitle += " (development)"
      End If

      Page.Title = pageTitle

      'Set the page culture
      SetPageCulture()

      'check to see if the database is locked
      If message.Length = 0 And Not IsPostBack Then

         If _db.IsSystemLocked() Then
            message = "Database locked.<BR><BR>Contact your system administrator."
         End If
      End If

      'check if the database and website versions match.
      If message.Length = 0 And Not IsPostBack Then

         Dim dbVersion As String = _db.GetSetting("database", "version", False)

         If assemblyName = CompiledAssemblyName Then
            ' Complied version of the web site, so perform version checks.
            If webSiteVersion.Length > 0 Then
               ' Just get the major and minor parts of the 4 part version.
               webSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major & "." & Assembly.GetExecutingAssembly.GetName.Version.Minor
            End If

            If dbVersion <> webSiteVersion Or webSiteVersion.Length = 0 Then
               ' Version mismatch.
               If dbVersion.Length = 0 Then
                  dbVersion = "&lt;unknown&gt;"
               End If
               If webSiteVersion.Length = 0 Then
                  webSiteVersion = "&lt;unknown&gt;"
               End If

               message = "The Workflow website version (" & webSiteVersion & ")" & " is incompatible with the database version (" & dbVersion & ")." & "<BR><BR>Contact your system administrator."
            End If
         End If
      End If

      If message.Length = 0 Then

         Try
            'FileUpload.apsx and FileDownload.aspx require the url details
            Session("workflowUrl") = _url

            ' Get the selected tab number for this workflow, if any...
            If Not IsPostBack Then
               hdnDefaultPageNo.Value = _db.GetWorkflowCurrentTab(_url.InstanceID).ToString
            End If

            'Get the worklfow form details
            _form = _db.GetWorkflowForm(_url.InstanceID, _url.ElementID)

            Dim script As String = ""
            message = CreateControls(_form, script)

            ScriptManager.GetCurrent(Page).AsyncPostBackTimeout = _config.SubmissionTimeout

            If (Not ClientScript.IsStartupScriptRegistered("Startup")) Then
               ' Form the script to be registered at client side.
               ClientScript.RegisterStartupScript(ClientScript.GetType, "Startup", "function pageLoad() {" & script & "}", True)
            End If

            If message.Length = 0 Then

               If _form.ErrorMessage <> "" Then
                  message = _form.ErrorMessage
               End If

               If _form.BackImage > 0 Then
                  Dim image As String = LoadPicture(_form.BackImage, message)
                  If message.Length = 0 Then
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

               hdnSiblingForms.Value = _siblingForms.ToString
            End If

            ' Resize the mobile 'viewport' to fit the webform
            AddHeaderTags(_form.Width)

         Catch ex As Exception
            message = "Error loading web form controls:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
         End Try

      End If

      If message.Length > 0 Then

         If IsPostBack Then
            bulletErrors.Items.Clear()
            bulletWarnings.Items.Clear()

            hdnErrorMessage.Value = message
            hdnFollowOnForms.Value = ""
            SetSubmissionMessage(message & "<BR><BR>Click", "here", "to close this form.")
         Else
            Session("message") = message
            Response.Redirect("Message.aspx")
         End If
      End If
   End Sub

   Private Function CreateControls(workflowForm As WorkflowForm, ByRef script As String) As String

      Dim message As String = ""

      'If there is a tab control move it to the beginning of the item list so that its created before other controls
      Dim tabItem As FormItem = _form.Items.FirstOrDefault(Function(fi) fi.ItemType = 21)
      If tabItem IsNot Nothing Then
         _form.Items.Remove(tabItem)
         _form.Items.Insert(0, tabItem)
      End If

      'Add the main form page that control will be added to
      Dim tabPages As New List(Of Panel)
      Dim tabPage = New Panel
      tabPage.ID = FormInputPrefix & (0).ToString & "_21_PageTab"
      tabPage.Style.Add("position", "absolute")
      tabPages.Add(tabPage)
      pnlInputDiv.Controls.Add(tabPage)

      'Create each of the controls for the form
      For Each formItem As FormItem In workflowForm.Items

         'Generate the unique ID for this control and process it onto the form.
         Dim sID As String = FormInputPrefix & formItem.Id & "_" & formItem.ItemType & "_"

         Select Case formItem.ItemType

            Case 0 ' Button
               Dim control = New HtmlInputButton
               With control
                  .ID = sID
                  .Style.ApplyLocation(formItem)
                  .Style.ApplySize(formItem)
                  .Style.ApplyFont(formItem)

                  .Attributes.Add("TabIndex", formItem.TabIndex.ToString)
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  ' If the button has no caption, we treat as a transparent button.
                  ' This is so we can emulate picture buttons with very little code changes.
                  If formItem.Caption = Nothing Then
                     .Style.Add("filter", "alpha(opacity=0)")
                     .Style.Add("opacity", "0")
                  End If

                  ' stops the mobiles displaying buttons with over-rounded corners...
                  If IsMobileBrowser() OrElse Utilities.IsMacSafari() Then
                     .Style.Add("-webkit-appearance", "none")
                     .Style.Add("background-color", "#E6E6E6")
                     .Style.Add("border", "solid 1px #CCC")
                     .Style.Add("border-radius", "4px")
                  End If

                  If formItem.BackColor <> 16249587 AndAlso formItem.BackColor <> -2147483633 Then
                     .Style.Add("background-color", General.GetHtmlColour(formItem.BackColor).ToString)
                     .Style.Add("border", "1px solid #CCC")
                     .Style.Add("border-radius", "4px")
                  End If

                  If formItem.ForeColor <> 6697779 Then
                     .Style.Add("color", General.GetHtmlColour(formItem.ForeColor).ToString)
                  End If

                  .Style.Add("padding", "0px")
                  .Style.Add("white-space", "normal")

                  .Value = formItem.Caption

                  .Style.Add("z-index", "2")

                  .Attributes.Add("onclick", "try{setPostbackMode(1);}catch(e){};")
               End With

               tabPages(formItem.PageNo).Controls.Add(control)

               AddHandler control.ServerClick, AddressOf ButtonClick

            Case 1 ' Database value

               Dim control = New Label
               With control
                  .ApplyLocation(formItem)
                  .ApplySize(formItem)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)

                  If formItem.PictureBorder Then
                     .ApplyBorder(True)
                  End If

                  .Style("word-wrap") = "break-word"
                  .Style("overflow") = "auto"

                  Select Case formItem.SourceItemType
                     Case -7 ' Logic
                        If formItem.Value = Nothing Then
                           .Text = "&lt;undefined&gt;"
                        ElseIf formItem.Value = "1" Then
                           .Text = Boolean.TrueString
                        Else
                           .Text = Boolean.FalseString
                        End If

                     Case 2, 4 ' Numeric, Integer
                        If formItem.Value = Nothing Then
                           .Text = "&lt;undefined&gt;"
                        Else
                           Dim value = formItem.Value.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                           If value.Chars(0) = "-" Then
                              value = value.Substring(1) & "-"
                           End If
                           .Text = value
                        End If

                     Case 11 ' Date
                        If formItem.Value = Nothing OrElse formItem.Value.Trim.Length = 0 Then
                           .Text = "&lt;undefined&gt;"
                        Else
                           .Text = General.ConvertSqlDateToLocale(formItem.Value)
                        End If
                     Case Else 'Text
                        .Text = formItem.Value
                  End Select

               End With

               tabPages(formItem.PageNo).Controls.Add(control)

            Case 2 ' Label
               Dim control = New Label
               With control
                  .ApplyLocation(formItem)
                  .ApplySize(formItem, 0, 1)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)

                  If formItem.PictureBorder Then
                     .ApplyBorder(True)
                  End If

                  ' NPG20120305 Fault HRPRO-1967 reverted by PBG20120419 Fault HRPRO-2157
                  '.Style("word-wrap") = "break-word"
                  .Style("overflow") = "auto"

                  If formItem.CaptionType = 3 Then 'calculated caption
                     .Text = formItem.Value
                  Else
                     .Text = formItem.Caption
                  End If
               End With

               tabPages(formItem.PageNo).Controls.Add(control)

            Case 3 ' Input value - character
               Dim control = New TextBox
               With control
                  .ID = sID
                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  .ApplyLocation(formItem)
                  .ApplySize(formItem, -1, -1)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem)
                  .ApplyBorder(True)

                  If formItem.PasswordType Then
                     .TextMode = TextBoxMode.Password
                  Else
                     .TextMode = TextBoxMode.MultiLine
                     .Wrap = True
                     .Style("overflow") = "auto"
                     .Style("word-wrap") = "break-word"
                     .Style("resize") = "none"
                  End If
                  .Style("padding") = "1px"

                  .Text = formItem.Value

                  .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                  If formItem.InputSize > 0 Then
                     .Attributes("maxlength") = formItem.InputSize.ToString
                  End If

                  If IsMobileBrowser() Then
                     .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                  End If

               End With

               tabPages(formItem.PageNo).Controls.Add(control)

            Case 4 ' Workflow value

               Dim control = New Label
               With control
                  .ApplyLocation(formItem)
                  .ApplySize(formItem)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)

                  If formItem.PictureBorder Then
                     .ApplyBorder(True)
                  End If

                  .Style("word-wrap") = "break-word"
                  .Style("overflow") = "auto"

                  Select Case formItem.SourceItemType
                     Case 6 ' Logic
                        If formItem.Value = Nothing Then
                           .Text = "&lt;undefined&gt;"
                        ElseIf formItem.Value = "1" Then
                           .Text = Boolean.TrueString
                        Else
                           .Text = Boolean.FalseString
                        End If

                     Case 5 ' Number
                        If formItem.Value = Nothing Then
                           .Text = String.Empty
                        Else
                           Dim value = formItem.Value.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                           If value.Length > 0 AndAlso value.Chars(0) = "-" Then
                              value = value.Substring(1) & "-"
                           End If
                           .Text = value
                        End If

                     Case 7 ' Date
                        If formItem.Value = Nothing OrElse formItem.Value.Trim.ToLower = "null" Then
                           .Text = "&lt;undefined&gt;"
                        Else
                           .Text = General.ConvertSqlDateToLocale(formItem.Value)
                        End If
                     Case Else 'Text
                        .Text = formItem.Value
                  End Select

               End With

               tabPages(formItem.PageNo).Controls.Add(control)

            Case 5 ' Input value - numeric

               Dim control = New TextBox
               With control
                  .ID = sID
                  .CssClass = "numeric"

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  .ApplyLocation(formItem)
                  .ApplySize(formItem, -1, -1)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)
                  .ApplyBorder(True)
                  .Style("padding") = "1px"

                  'add attributes that denote the min & max values, number of decimal places is also implied
                  Dim max = New String("9"c, formItem.InputSize - formItem.InputDecimals) &
                            If(formItem.InputDecimals > 0, "." & New String("9"c, formItem.InputDecimals), "")

                  .Attributes.Add("data-numeric", String.Format("{{vMin: '-{0}', vMax: '{0}'}}", max))

                  'set the control value
                  Dim value As Single
                  If Not formItem.Value = Nothing Then
                     value = CSng(formItem.Value.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator))
                  End If
                  .Text = value.ToString("N" & formItem.InputDecimals).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator, "")

                  .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                  If IsMobileBrowser() Then
                     .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                  End If

               End With
               tabPages(formItem.PageNo).Controls.Add(control)

            Case 6 ' Input value - logic

               Dim checkBox = New CheckBox
               With checkBox
                  .ID = sID
                  .ApplyLocation(formItem)
                  .ApplySize(formItem)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  .CssClass = If(formItem.Alignment = 0, "checkbox left", "checkbox right")
                  If Utilities.IsAndroidBrowser() Then .CssClass += " android"
                  .Style("line-height") = formItem.Height.ToString & "px"

                  .Text = formItem.Caption
                  .Checked = (formItem.Value.ToLower = "true")

                  If IsMobileBrowser() Then
                     .Attributes("onclick") = "FilterMobileLookup('" & sID & "');"
                  End If
               End With

               tabPages(formItem.PageNo).Controls.Add(checkBox)

            Case 7 ' Input value - date

               Dim control = New TextBox
               With control
                  .ID = sID
                  .CssClass = "date"

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem, True)

                  If Utilities.GetBrowserFamily() = "IOS" Then
                     'use html5 date
                     .Attributes.Add("type", "date")
                     'ios sizing fix
                     .ApplySize(formItem, -10, -3)
                     'ios requires date in yyyy-mm-dd format
                     .Text = If(formItem.Value = Nothing, "", DateTime.ParseExact(formItem.Value, "MM/dd/yyyy", Nothing).ToString("yyyy-MM-dd"))
                  Else
                     .CssClass += " withPicker"
                     .ApplySize(formItem, -1, -1)
                     .ApplyBorder(True)
                     .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"
                     .Text = General.ConvertSqlDateToLocale(formItem.Value)
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
               panel.ApplyLocation(formItem)

               tabPages(formItem.PageNo).Controls.Add(panel)

            Case 8 ' Frame

               Dim top = formItem.Top, left = formItem.Left, width = formItem.Width - 2, height = formItem.Height - 2
               Dim fontAdjustment = CInt(formItem.FontSize * 0.8)

               If formItem.Caption.Trim.Length = 0 Then
                  top += fontAdjustment
                  height -= fontAdjustment
               End If

               Dim html = String.Format("<fieldset style='position:absolute; top:{0}px; left:{1}px; width:{2}px; height:{3}px; {4} {5} border:1px solid #999;'>",
                                        top, left, width, height, GetFontCss(formItem), GetColorCss(formItem))

               If formItem.Caption.Trim.Length > 0 Then
                  html += String.Format("<legend>{0}</legend>", formItem.Caption) & vbCrLf
               End If

               html += "</fieldset>" & vbCrLf

               tabPages(formItem.PageNo).Controls.Add(New LiteralControl(html))

            Case 9 ' Line

               Dim html As String

               Select Case formItem.Orientation
                  Case 0 ' Vertical
                     html = String.Format("<div style='position:absolute; left:{0}px; top:{1}px; height:{2}px; width:0px; border-left: 1px solid {3};'></div>",
                                         formItem.Left, formItem.Top, formItem.Height, General.GetHtmlColour(formItem.BackColor))
                  Case Else ' Horizontal
                     html = String.Format("<div style='position:absolute; left:{0}px; top:{1}px; height:0px; width:{2}px; border-left: 1px solid {3};'></div>",
                                         formItem.Left, formItem.Top, formItem.Width, General.GetHtmlColour(formItem.BackColor))
               End Select

               tabPages(formItem.PageNo).Controls.Add(New LiteralControl(html))

            Case 10 ' Image

               Dim control = New WebControls.Image

               With control
                  .ApplyLocation(formItem)
                  .ApplySize(formItem)

                  If formItem.PictureBorder Then
                     .ApplyBorder(True, -2)
                  End If

                  Dim imageUrl As String = LoadPicture(formItem.PictureID, message)
                  If message.Length > 0 Then
                     Exit For
                  End If
                  .ImageUrl = imageUrl
               End With

               tabPages(formItem.PageNo).Controls.Add(control)

            Case 11 ' Record Selection Grid
               ' NPG20110501 Fault HR PRO 1414
               ' We're using the ASP.NET standard gridview control now. To replicate the old infragistics
               ' grid we'll put the Gridview control within a DIV to enable scroll bars and fix the height&width, 
               ' but also put a header DIV above the grid which contains copies of the column headers. This is 
               ' to simulate fixing the headers when the grid is scrolled. We use this table to allow 
               ' clickable sorting and resizable column widths.

               ' Grids are now created using the clsRecordSelector class.
               Dim recordSelector = New RecordSelector
               With recordSelector

                  .CssClass = "recordSelector"
                  .Style.Add("Position", "Absolute")
                  .Attributes.CssStyle("LEFT") = Unit.Pixel(formItem.Left).ToString
                  .Attributes.CssStyle("TOP") = Unit.Pixel(formItem.Top).ToString
                  .Attributes.CssStyle("WIDTH") = Unit.Pixel(formItem.Width).ToString

                  ' Don't use .height - it causes large row heights if the grid isn't filled.
                  ' Use .ControlHeight instead - custom property.
                  .ControlHeight = formItem.Height
                  .Width = formItem.Width

                  'TODO PG changing this color makes no difference must be set in the recordSelector class
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
                  If Utilities.GetBrowserFamily() = "ANDROID" Then
                     Dim piRowHeight As Double = (formItem.FontSize - 8) + 21
                     .PageSize = Math.Min(CInt(Math.Truncate((CInt(formItem.Height - 42) / piRowHeight))), _config.LookupRowsRange)
                     .RowStyle.Height = CInt(piRowHeight)
                  Else
                     .PageSize = _config.LookupRowsRange
                  End If

                  .IsLookup = False
                  ' EnableViewState must be on. Mucks up the grid data otherwise. Should be reviewed
                  ' if performance is silly, but while paging is enabled it shouldn't be too bad.
                  .EnableViewState = True
                  .IsEmpty = False
                  .EmptyDataText = "no records to display"
                  .ShowHeaderWhenEmpty = True

                  ' Header Row
                  .ColumnHeaders = formItem.ColumnHeaders
                  .HeadFontSize = CSng(formItem.HeadFontSize)
                  .HeadLines = formItem.HeadLines

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  Dim backColor As Integer = formItem.BackColor

                  If backColor = 16777215 AndAlso formItem.BackColorEven = 15988214 Then
                     backColor = formItem.BackColorEven
                  End If

                  .BackColor = General.GetColour(backColor)
                  .ForeColor = General.GetColour(formItem.ForeColor)

                  .HeaderStyle.BackColor = General.GetColour(formItem.HeaderBackColor)
                  .HeaderStyle.BorderColor = General.GetColour(10720408)
                  .HeaderStyle.BorderStyle = BorderStyle.Double
                  .HeaderStyle.BorderWidth = 0

                  .HeaderStyle.Font.Apply(formItem, "Head")

                  .HeaderStyle.ForeColor = General.GetColour(formItem.ForeColor)
                  .HeaderStyle.Wrap = False
                  .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                  .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                  ' PagerStyle settings
                  .PagerStyle.BackColor = General.GetColour(formItem.HeaderBackColor)
                  .PagerStyle.BorderColor = General.GetColour(10720408)
                  .PagerStyle.BorderStyle = BorderStyle.Solid
                  .PagerStyle.BorderWidth = 0
                  .PagerStyle.Font.Apply(formItem, "Head")
                  .PagerStyle.ForeColor = General.GetColour(formItem.ForeColor)
                  .PagerStyle.Wrap = False
                  .PagerStyle.VerticalAlign = VerticalAlign.Middle
                  .PagerStyle.HorizontalAlign = HorizontalAlign.Center

                  .Font.Apply(formItem)

                  If formItem.ForeColorEven <> formItem.ForeColor Then
                     .RowStyle.ForeColor = General.GetColour(formItem.ForeColorEven)
                  End If

                  If formItem.BackColorEven <> backColor Then
                     .RowStyle.BackColor = General.GetColour(formItem.BackColorEven)
                  End If

                  If formItem.ForeColorOdd <> formItem.ForeColor Then
                     .AlternatingRowStyle.ForeColor = General.GetColour(formItem.ForeColorOdd)
                  End If

                  If formItem.BackColorOdd <> formItem.BackColorEven Then
                     .AlternatingRowStyle.BackColor = General.GetColour(formItem.BackColorOdd)
                  End If

                  If Not formItem.ForeColorHighlight.HasValue Then
                     .SelectedRowStyle.ForeColor = SystemColors.HighlightText
                  Else
                     .SelectedRowStyle.ForeColor = General.GetColour(formItem.ForeColorHighlight.Value)
                  End If
                  If Not formItem.BackColorHighlight.HasValue Then
                     .SelectedRowStyle.BackColor = SystemColors.Highlight
                  Else
                     .SelectedRowStyle.BackColor = General.GetColour(formItem.BackColorHighlight.Value)
                  End If

               End With

               ' ==================================================
               ' Add the Paging Grid View to the holding panel now.
               ' Before the databind, or you'll get errors!
               ' ==================================================
               tabPages(formItem.PageNo).Controls.Add(recordSelector)

               If Not IsPostBack Then

                  Dim result = _db.GetWorkflowGridItems(formItem.Id, _url.InstanceID)

                  Session(sID & "DATA") = result.Data

                  recordSelector.DataKeyNames = New String() {"ID"}

                  recordSelector.IsEmpty = (result.Data.Rows.Count = 0)
                  recordSelector.DataSource = result.Data
                  recordSelector.DataBind()

                  'set the default value
                  If Utilities.NullSafeInteger(formItem.Value) <> 0 Then

                     Dim colIndex As Integer = result.Data.Columns.IndexOf("ID")

                     For r = 0 To result.Data.Rows.Count - 1
                        If result.Data.Rows(r).Item(colIndex).ToString = formItem.Value Then
                           ' set selected page index and row number
                           recordSelector.PageIndex = CInt(r \ recordSelector.PageSize)
                           recordSelector.SelectedIndex = CInt(r Mod recordSelector.PageSize)
                           recordSelector.DataBind()
                           Exit For
                        End If
                     Next
                  End If

                  If recordSelector.SelectedIndex = -1 AndAlso recordSelector.Rows.Count > 0 Then
                     recordSelector.SelectedIndex = 0
                  End If

                  If Not result.Ok Then
                     message = "Error loading web form. Web Form record selector item record has been deleted or not selected."
                     Exit For
                  End If
               End If

               ' Hidden field is used to store scroll position of the grid.
               tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})

            Case 14 ' lookup  Inputs
               If Not IsMobileBrowser() Then

                  ' ============================================================
                  ' Create a textbox as the main control
                  ' ============================================================
                  Dim textBox = New TextBox

                  With textBox
                     .ID = sID & "TextBox"
                     .ApplyLocation(formItem)
                     .ApplySize(formItem, -1, -1)
                     .Style.ApplyFont(formItem)
                     .ApplyColor(formItem)
                     .ApplyBorder(True)

                     .TabIndex = formItem.TabIndex
                     UpdateAutoFocusControl(formItem.TabIndex, sID & "TextBox")

                     .ReadOnly = True
                     .Style.Add("padding", "1px")
                     .Style.Add("background-image", "url('images/downarrow.gif')")
                     .Style.Add("background-position", "right top")
                     .Style.Add("background-repeat", "no-repeat")
                     .Style.Add("background-origin", "content-box")
                     .Style.Add("background-size", "17px 100%")
                  End With

                  tabPages(formItem.PageNo).Controls.Add(textBox)

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
                     .Style("top") = Unit.Pixel(formItem.Top).ToString
                     .Style("left") = Unit.Pixel(formItem.Left).ToString

                     .Attributes.CssStyle("left") = Unit.Pixel(formItem.Left).ToString
                     .Attributes.CssStyle("top") = Unit.Pixel(formItem.Top).ToString
                     .Attributes.CssStyle("width") = Unit.Pixel(formItem.Width).ToString

                     ' Don't set the height of this control. Must use the ControlHeight property
                     ' to stop the grid's rows from autosizing.
                     .ControlHeight = formItem.Height
                     .Width = formItem.Width

                     ' Header Row - fixed for lookups.
                     .ColumnHeaders = True
                     .HeadFontSize = CSng(formItem.FontSize)
                     .HeadLines = 1

                     .ApplyFont(formItem)
                     .ApplyColor(formItem)
                     .ApplyBorder(False)

                     .SelectedRowStyle.ForeColor = General.GetColour(2774907)
                     .SelectedRowStyle.BackColor = General.GetColour(10480637)

                     ' HEADER formatting
                     .HeaderStyle.BackColor = General.GetColour(16248553)
                     .HeaderStyle.BorderColor = General.GetColour(10720408)
                     .HeaderStyle.BorderStyle = BorderStyle.Solid
                     .HeaderStyle.BorderWidth = 0

                     .HeaderStyle.Font.Apply(formItem)
                     .HeaderStyle.ForeColor = General.GetColour(formItem.ForeColor)
                     .HeaderStyle.Wrap = False
                     .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                     .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                     .PagerStyle.Font.Apply(formItem)
                     .PagerStyle.ForeColor = General.GetColour(formItem.ForeColor)
                     .PagerStyle.Wrap = False
                     .PagerStyle.VerticalAlign = VerticalAlign.Middle
                     .PagerStyle.HorizontalAlign = HorizontalAlign.Center
                     .PagerStyle.BorderWidth = 0
                  End With

                  Dim filterSql = LookupFilterSQL(formItem.LookupFilterColumnName,
                                                  formItem.LookupFilterColumnDataType,
                                                  formItem.LookupFilterOperator,
                                                  FormInputPrefix &
                                                  formItem.LookupFilterValueID &
                                                  "_" & formItem.LookupFilterValueType & "_")

                  ' Hidden Field to store any lookup filter code
                  If (filterSql.Length > 0) Then
                     tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = filterSql})
                  End If

                  tabPages(formItem.PageNo).Controls.Add(recordSelector)

                  If Not IsPostBack Then

                     'get the data
                     Dim result = _db.GetWorkflowItemValues(formItem.Id, _url.InstanceID)

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
                     textBox.Attributes.Add("DataType", result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

                     'set the default value
                     textBox.Text = result.DefaultValue

                     For i As Integer = 0 To recordSelector.Rows.Count - 1
                        If i > recordSelector.PageSize Then Exit For
                        ' don't bother if on other pages
                        If recordSelector.Rows(i).Cells(result.LookupColumnIndex).Text = result.DefaultValue Then
                           recordSelector.SelectedIndex = i
                           Exit For
                        End If
                     Next
                  End If

                  ' AJAX DropDownExtender (DDE) Control - this simply links up the DropDownList and the Lookup Grid to make a dropdown.
                  Dim dde As New DropDownExtender

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

                  tabPages(formItem.PageNo).Controls.Add(dde)

                  ' Attach a JavaScript functino to the 'add_shown' method of this
                  ' DropDownExtender. Used to check if popup is bigger than the
                  ' parent form, and resize the parent form if necessary
                  script += "var bhvDdl=$find('" & dde.BehaviorID.ToString & "');"
                  script += "try {bhvDdl.add_shown(ResizeComboForForm);} catch (e) {}"

                  ' hidden field to store scroll position (not required?)
                  tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})

                  ' hidden field to hold any filter SQL code
                  tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

                  ' Hidden Button for JS to call which fires filter click event. 
                  Dim button = New Button
                  With button
                     .ID = sID & "refresh"
                     .Style.Add("display", "none")
                     .Text = .ID
                  End With

                  AddHandler button.Click, AddressOf SetLookupFilter

                  tabPages(formItem.PageNo).Controls.Add(button)
               Else
                  ' ================================================================================================================
                  ' Mobile Browser - convert lookup data to a standard dropdown.
                  ' ================================================================================================================
                  Dim control As New DropDownList

                  With control
                     .ID = sID
                     .ApplyLocation(formItem)
                     .ApplySize(formItem, -1, -1)
                     .Style.ApplyFont(formItem)
                     .ApplyColor(formItem)
                     If Not IsMobileBrowser() Then .ApplyBorder(False)
                     .Style.Add("padding", "1px")

                     .TabIndex = formItem.TabIndex
                     UpdateAutoFocusControl(formItem.TabIndex, sID)

                     .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")

                     tabPages(formItem.PageNo).Controls.Add(control)

                     Dim filterSql = LookupFilterSQL(formItem.LookupFilterColumnName,
                                                     formItem.LookupFilterColumnDataType,
                                                     formItem.LookupFilterOperator,
                                                     FormInputPrefix &
                                                     formItem.LookupFilterValueID & "_" &
                                                     formItem.LookupFilterValueType & "_")

                     If (filterSql.Length > 0) Then
                        tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = filterSql})
                     End If

                     If Not IsPostBack Then

                        'get the data
                        Dim result = _db.GetWorkflowItemValues(formItem.Id, _url.InstanceID)

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
                        .Attributes.Add("DataType", result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

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
                  tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

                  ' Hidden Button for JS to call which fires filter click event. 
                  Dim button = New Button
                  With button
                     .ID = sID & "refresh"
                     .Style.Add("display", "none")
                  End With

                  AddHandler button.Click, AddressOf SetLookupFilter

                  tabPages(formItem.PageNo).Controls.Add(button)
               End If

            Case 13 ' Dropdown (13) Inputs

               Dim control As New DropDownList

               With control
                  .ID = sID
                  .ApplyLocation(formItem)
                  .ApplySize(formItem, -1, -1)
                  .Style.ApplyFont(formItem)
                  .ApplyColor(formItem)
                  If Not IsMobileBrowser() Then .ApplyBorder(False)
                  .Style.Add("padding", "1px")

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  If IsMobileBrowser() Then
                     .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                  End If

                  tabPages(formItem.PageNo).Controls.Add(control)

                  Dim filterSql = LookupFilterSQL(formItem.LookupFilterColumnName,
                                                  formItem.LookupFilterColumnDataType,
                                                  formItem.LookupFilterOperator,
                                                  FormInputPrefix &
                                                  formItem.LookupFilterValueID &
                                                  "_" & formItem.LookupFilterValueType & "_")

                  If filterSql.Length > 0 Then
                     tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = filterSql})
                  End If

                  If Not IsPostBack Then
                     'get the data
                     Dim result = _db.GetWorkflowItemValues(formItem.Id, _url.InstanceID)

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
                     .Attributes.Add("DataType", result.Data.Columns(result.LookupColumnIndex).DataType.ToString)

                     'set the default value
                     Dim item As ListItem = control.Items.FindByValue(result.DefaultValue)
                     If item IsNot Nothing Then
                        .SelectedValue = item.Value
                     End If

                  End If

               End With

            Case 15 ' OptionGroup

               Dim top = formItem.Top, left = formItem.Left, width = formItem.Width, height = formItem.Height
               Dim fontAdjustment = CInt(formItem.FontSize * 0.8)
               Dim borderCss As String

               Dim radioTop As Int32

               If Not formItem.PictureBorder Then
                  borderCss = "border-style: none;"
                  radioTop = 2
               Else
                  borderCss = "border: 1px solid #999;"
                  width -= 2
                  height -= 2

                  If formItem.Caption.Trim.Length = 0 Then
                     top += fontAdjustment
                     height -= fontAdjustment
                  End If

                  radioTop = 19 + CInt((formItem.FontSize - 8) * 1.375)

                  If Utilities.IsAndroidBrowser() AndAlso formItem.Orientation = 0 Then
                     radioTop -= 5
                  End If
               End If

               Dim html = String.Format("<fieldset style='position:absolute; top:{0}px; left:{1}px; width:{2}px; height:{3}px; {4} {5} {6}'>",
                                        top, left, width, height, GetFontCss(formItem), GetColorCss(formItem), borderCss)

               If formItem.PictureBorder And formItem.Caption.Trim.Length > 0 Then
                  html += String.Format("<legend>{0}</legend>", formItem.Caption) & vbCrLf
               End If

               html += "</fieldset>" & vbCrLf

               tabPages(formItem.PageNo).Controls.Add(New LiteralControl(html))

               Dim radioList As New RadioButtonList
               With radioList
                  .ID = sID
                  .Style.ApplyFont(formItem)
                  .CssClass = "radioList"
                  If Utilities.IsAndroidBrowser() Then .CssClass += " android"

                  .TabIndex = formItem.TabIndex
                  UpdateAutoFocusControl(formItem.TabIndex, sID & "_0")

                  .RepeatDirection = If(formItem.Orientation = 0, WebControls.RepeatDirection.Vertical, WebControls.RepeatDirection.Horizontal)

                  .Style("position") = "absolute"
                  .Style("top") = Unit.Pixel(radioTop + formItem.Top).ToString
                  .Style("left") = Unit.Pixel(9 + formItem.Left).ToString
                  .Width() = formItem.Width - 12
               End With

               tabPages(formItem.PageNo).Controls.Add(radioList)

               If Not IsPostBack Then

                  'get the data
                  Dim result = _db.GetWorkflowItemValues(formItem.Id, _url.InstanceID)

                  'bind to the data
                  radioList.DataTextField = result.Data.Columns(0).ColumnName
                  radioList.DataSource = result.Data
                  radioList.DataBind()

                  'set the default value
                  radioList.SelectedValue = result.DefaultValue

                  If radioList.SelectedIndex = -1 Then
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
                  .Style.ApplyLocation(formItem)
                  .Style.ApplySize(formItem)
                  .Style.ApplyFont(formItem)

                  .Attributes.Add("TabIndex", formItem.TabIndex.ToString)
                  UpdateAutoFocusControl(formItem.TabIndex, sID)

                  ' stops the mobiles displaying buttons with over-rounded corners...
                  If IsMobileBrowser() OrElse Utilities.IsMacSafari() Then
                     .Style.Add("-webkit-appearance", "none")
                     .Style.Add("background-color", "#E6E6E6")
                     .Style.Add("border", "solid 1px #CCC")
                     .Style.Add("border-radius", "4px")
                  End If

                  If formItem.BackColor <> 16249587 AndAlso formItem.BackColor <> -2147483633 Then
                     .Style.Add("background-color", General.GetHtmlColour(formItem.BackColor).ToString)
                     .Style.Add("border", "solid 1px #CCC")
                     .Style.Add("border-radius", "4px")
                  End If

                  If formItem.ForeColor <> 6697779 Then
                     .Style.Add("color", General.GetHtmlColour(formItem.ForeColor).ToString)
                  End If

                  .Style.Add("padding", "0px")
                  .Style.Add("white-space", "normal")

                  .Value = formItem.Caption

                  Dim crypt As New Crypt,
                      sEncodedID As String = crypt.SimpleEncrypt(formItem.Id.ToString, Session.SessionID)

                  If Not IsMobileBrowser() Then
                     .Attributes.Add("onclick", "try{showFileUpload(true, '" & sEncodedID & "', document.getElementById('file" & sID & "').value);}catch(e){};")
                  Else
                     .Attributes.Add("onclick", "try{alert('Your browser does not support file upload.');}catch(e){};")
                  End If
               End With

               tabPages(formItem.PageNo).Controls.Add(control)

               tabPages(formItem.PageNo).Controls.Add(New HiddenField With {.ID = "file" & sID, .Value = formItem.Value})

            Case 19, 20 ' DB File or WF File

               Dim crypt As New Crypt, sEncodedID As String = crypt.SimpleEncrypt(formItem.Id.ToString, Session.SessionID)

               Dim html = "<span id='{0}' tabindex='{1}' style='position:absolute; display:inline-block; word-wrap:break-word; overflow:auto; " &
                           "top:{2}px; left:{3}px; width:{4}px; height:{5}px; {6} {7}' " &
                           "onclick='FileDownload_Click(""{8}"");' onkeypress='FileDownload_KeyPress(""{8}"");'>{9}</span>"

               html = String.Format(html, sID, formItem.TabIndex, formItem.Top, formItem.Left, formItem.Width, formItem.Height,
                                    GetFontCss(formItem), GetColorCss(formItem), sEncodedID, HttpUtility.HtmlEncode(formItem.Caption))

               UpdateAutoFocusControl(formItem.TabIndex, sID)

               tabPages(formItem.PageNo).Controls.Add(New LiteralControl(html))

            Case 21 ' Tab Strip

               'split out the tab names to calculate number of tabs - may not have loaded all tabs yet, so can't count them.
               Dim arrTabCaptions As String() = formItem.Caption.Split(New Char() {";"c})

               pnlTabsDiv.Style("width") = formItem.Width & "px"
               pnlTabsDiv.Style("height") = formItem.Height & "px"
               pnlTabsDiv.Style("left") = formItem.Left & "px"
               pnlTabsDiv.Style("top") = formItem.Top & "px"

               Dim ctlTabsDiv As New Panel
               ctlTabsDiv.ID = "TabsDiv"
               ctlTabsDiv.Style.Add("height", TabStripHeight & "px")
               ctlTabsDiv.Style.Add("position", "relative")
               ctlTabsDiv.Style.Add("z-index", "1")

               If IsMobileBrowser() And Not Utilities.IsAndroidBrowser() Then
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
                     .Attributes.Add("onclick", "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft - 20;")
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
                     .Attributes.Add("onclick", "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft + 20;")
                  End With
                  ctlFormTabArrows.Controls.Add(image)

                  pnlTabsDiv.Controls.Add(ctlFormTabArrows)
               End If

               ' generate the tabs.
               Dim ctlTabsTable As New Table
               ctlTabsTable.CellSpacing = 0
               ' ctlTabsTable.Style.Add("margin-top", "2px")
               Dim trPager As TableRow = New TableRow()
               trPager.Height = TabStripHeight - 1
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

                     ' Create the tab pages
                     tabPage = New Panel
                     tabPage.ID = FormInputPrefix & iTabNo.ToString & "_21_PageTab"
                     tabPage.Style.Add("position", "absolute")
                     tabPage.Style.Add("top", (formItem.Top + TabStripHeight) & "px")
                     tabPage.Style.Add("left", formItem.Left & "px")
                     If iTabNo > 1 Then
                        tabPage.Style.Add("display", "none")
                     End If
                     ' Add this tab to the web form
                     tabPages.Add(tabPage)
                     pnlInputDiv.Controls.Add(tabPage)

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

         If message <> "" Then Exit For
      Next

      Return message
   End Function

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

   Private Sub GetControls(controlCollection As ControlCollection, result As ICollection(Of Control), Optional predicate As Func(Of Control, Boolean) = Nothing)

      For Each c As Control In controlCollection
         If predicate Is Nothing OrElse predicate(c) Then
            result.Add(c)
         End If
         If c.HasControls Then
            GetControls(c.Controls, result, predicate)
         End If
      Next
   End Sub

   Public Sub ButtonClick(ByVal sender As Object, ByVal e As EventArgs)

      Dim valueString As String = ""
      Dim message As String = ""

      Try
         ' Read the web form item values & build up a string of the form input values.
         ' This is a tab delimited string of itemIDs and values.
         Dim controlList As New List(Of Control)
         GetControls(Page.Controls, controlList, Function(c) c.ClientID.StartsWith(FormInputPrefix) AndAlso
                                 (c.ClientID.EndsWith("_") OrElse c.ClientID.EndsWith("TextBox") OrElse c.ClientID.EndsWith("Grid")))

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
                     value = If (CSng(control.Text) = CSng(0), "0",control.Text.Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator,"."))
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
                     value = If (control.Text.Trim = "", "null",DateTime.Parse(control.Text).ToString("MM/dd/yyyy"))
                     valueString += sIDString & value & vbTab
                  End If

               Case 11 ' Grid (RecordSelector) Input
                  If TypeOf ctlFormInput Is RecordSelector Then
                     Dim control = DirectCast(ctlFormInput, RecordSelector)
                     value =If(control.SelectedValue IsNot Nothing, CStr(control.SelectedValue), "0")
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
                           value =If(control.Text = "", "null", General.ConvertLocaleDateToSql(control.Text))
                        ElseIf control.Attributes("DataType") = "System.Decimal" Or control.Attributes("DataType") = "System.Int32" Then
                           value = If (control.Text = "", "", control.Text.Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, "."))
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
         message = "Error reading web form item values:<BR><BR>" & ex.Message
      End Try

      If message.Length = 0 Then

         ' Validate the web form entry.
         errorMessagePanel.Font.Name = "Verdana"
         errorMessagePanel.Font.Size = _config.ValidationMessageFontSize
         errorMessagePanel.ForeColor = General.GetColour(6697779)

         bulletErrors.Items.Clear()
         bulletWarnings.Items.Clear()

         Dim result = _db.WorkflowValidateWebForm(_url.ElementID, _url.InstanceID, valueString)

         result.Errors.ForEach(Sub(f) bulletErrors.Items.Add(f))

         If hdnOverrideWarnings.Value <> "1" Then
            result.Warnings.ForEach(Sub(f) bulletWarnings.Items.Add(f))
         End If

         hdnCount_Errors.Value = CStr(bulletErrors.Items.Count)
         hdnCount_Warnings.Value = CStr(bulletWarnings.Items.Count)
         hdnOverrideWarnings.Value = "0"

         lblErrors.Text = If(bulletErrors.Items.Count > 0, "Unable to submit this form due to the following error" & If(bulletErrors.Items.Count = 1, "", "s") & ":","")

         lblWarnings.Text = If(bulletWarnings.Items.Count > 0,
                            If(bulletErrors.Items.Count > 0, "And the following warning" & If(bulletWarnings.Items.Count = 1, "", "s") & ":",
                                  "Submitting this form raises the following warning" & If(bulletWarnings.Items.Count = 1, "", "s") & ":"), "")

         overrideWarning.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)

         ' Submit the webform
         If bulletWarnings.Items.Count = 0 And bulletErrors.Items.Count = 0 Then

            Try
               'TODO NOW PG why transactionscope???
               Dim submit As SubmitWebFormResult
               Using (New TransactionScope(TransactionScopeOption.Suppress))
                  submit = _db.WorkflowSubmitWebForm(_url.ElementID, _url.InstanceID, valueString, Utilities.NullSafeInteger(hdnDefaultPageNo.Value))
               End Using

               hdnFollowOnForms.Value = ""

               If submit.SavedForLater Then
                  Select Case _form.SavedForLaterMessageType
                     Case 1 ' Custom
                        If Not SetSubmissionMessage(_form.SavedForLaterMessage) Then
                           SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here", "to close this form.")
                        End If
                     Case 2 ' None
                        SetSubmissionMessage("", "", "")
                     Case Else 'System default
                        SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here", "to close this form.")
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
                         Select(Function(f) _db.GetWorkflowQueryString(_url.InstanceID, CInt(f))).
                         ToArray()

                  hdnFollowOnForms.Value = String.Join(vbTab, followOnForms)

                  Select Case _form.FollowOnFormsMessageType
                     Case 1 ' Custom
                        If Not SetSubmissionMessage(_form.FollowOnFormsMessage) Then
                           SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to complete the follow-on Workflow form" & If(followOnForms.Count = 1, "", "s") & ".")
                        End If
                     Case 2 ' None
                        SetSubmissionMessage("", "", "")
                     Case Else 'System default
                        SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to complete the follow-on Workflow form" & If(followOnForms.Count = 1, "", "s") & ".")
                  End Select
               End If

            Catch ex As Exception
               message = "Error submitting the web form:<BR><BR>" & ex.Message
            End Try
         End If

      End If

      If message.Length > 0 Then
         bulletErrors.Items.Clear()
         bulletWarnings.Items.Clear()

         hdnErrorMessage.Value = message
         hdnFollowOnForms.Value = ""
         SetSubmissionMessage(message & "<BR><BR>Click", "here", "to close this form.")
      End If
   End Sub

   Private Sub UpdateAutoFocusControl(tabIndex As Short, focusId As String)
      If Not _minTabIndex.HasValue OrElse tabIndex < _minTabIndex.Value Then
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
      Return Utilities.IsAndroidBrowser()
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
      Return "Application Name=OpenHR Workflow;Data Source=" & _url.Server & ";Initial Catalog=" & _url.Database & ";Integrated Security=false;User ID=" & _url.User & ";Password=" & _url.Password & ";Pooling=false"
   End Function

   Private Function LoadPicture(ByVal piPictureID As Int32, ByRef psErrorMessage As String) As String

      Dim conn As SqlConnection
      Dim cmdSelect As SqlCommand
      Dim dr As SqlDataReader
      Dim sImageFileName As String
      Dim sImageFilePath As String
      Dim sTempName As String
      Dim fs As FileStream
      Dim bw As BinaryWriter
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
               sName = Utilities.NullSafeString(dr("name"))
               iIndex = sName.LastIndexOf(".")
               If iIndex >= 0 Then
                  sExtension = sName.Substring(iIndex)
               End If

               sImageFileName = Session.SessionID().ToString & "_" & _imageCount.ToString & "_" & Date.Now.Ticks.ToString & sExtension
               sTempName = sImageFilePath & "\" & sImageFileName

               ' Create a file to hold the output.
               fs = New FileStream(sTempName, FileMode.OpenOrCreate, FileAccess.Write)
               bw = New BinaryWriter(fs)

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

   Private Function LookupFilterSQL(ByVal columnName As String, ByVal columnDataType As Integer, ByVal operatorID As Integer, ByVal value As String) As String

      If Not (columnName.Length > 0 And operatorID > 0 And value.Length > 0) Then
         Return ""
      End If

      Select Case columnDataType
         Case SqlDataType.Boolean
            Select Case operatorID
               Case FilterOperators.giFILTEROP_EQUALS
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab
               Case FilterOperators.giFILTEROP_NOTEQUALTO
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab
            End Select
         Case SqlDataType.Numeric, SqlDataType.Integer
            Select Case operatorID
               Case FilterOperators.giFILTEROP_EQUALS
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab

               Case FilterOperators.giFILTEROP_NOTEQUALTO
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab

               Case FilterOperators.giFILTEROP_ISATMOST
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <= " & vbTab

               Case FilterOperators.giFILTEROP_ISATLEAST
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) >= " & vbTab

               Case FilterOperators.giFILTEROP_ISMORETHAN
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) > " & vbTab

               Case FilterOperators.giFILTEROP_ISLESSTHAN
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) < " & vbTab
            End Select

         Case SqlDataType.Date
            Select Case operatorID
               Case FilterOperators.giFILTEROP_ON
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

               Case FilterOperators.giFILTEROP_NOTON
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

               Case FilterOperators.giFILTEROP_ONORBEFORE
                  Return columnDataType.ToString & vbTab & value & vbTab & "LEN(ISNULL([ASRSysLookupFilterValue], '')) = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') <= '" & vbTab & "')"

               Case FilterOperators.giFILTEROP_ONORAFTER
                  Return columnDataType.ToString & vbTab & value & vbTab & "LEN('" & vbTab & "') = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') >= '" & vbTab & "')"

               Case FilterOperators.giFILTEROP_AFTER
                  Return columnDataType.ToString & vbTab & value & vbTab & "(LEN('" & vbTab & "') = 0 AND LEN(ISNULL([ASRSysLookupFilterValue], '')) > 0) OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') > '" & vbTab & "')"

               Case FilterOperators.giFILTEROP_BEFORE
                  Return columnDataType.ToString & vbTab & value & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') < '" & vbTab & "'"
            End Select

         Case SqlDataType.VarChar, SqlDataType.VarBinary, SqlDataType.LongVarChar
            Select Case operatorID
               Case FilterOperators.giFILTEROP_IS
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

               Case FilterOperators.giFILTEROP_ISNOT
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

               Case FilterOperators.giFILTEROP_CONTAINS
                  Return columnDataType.ToString & vbTab & value & vbTab & "ISNULL([ASRSysLookupFilterValue], '') LIKE '%" & vbTab & "%'"

               Case FilterOperators.giFILTEROP_DOESNOTCONTAIN
                  Return columnDataType.ToString & vbTab & value & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') NOT LIKE '%" & vbTab & "%'"
            End Select
      End Select

      Return ""

   End Function

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
      Dim hiddenField As HiddenField = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "filterSQL")), HiddenField)

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

         Dim gridView As RecordSelector = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "Grid")), RecordSelector)

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
            dataTable.Rows.InsertAt(dataTable.NewRow(), 0)
         End If
      End If
   End Sub

   Private Sub AddHeaderTags(ByVal viewportWidth As Long)

      ' Create the following timeout meta tag programatically for all browsers
      ' <meta http-equiv="refresh" content="5; URL=timeout.aspx" />
      Dim meta As New HtmlMeta()
      meta.HttpEquiv = "refresh"
      meta.Content = (Session.Timeout * 60).ToString & "; URL=timeout.aspx"

      Page.Header.Controls.Add(meta)

      ' for Mobiles only, set the viewport and 'home page' icons
      If IsMobileBrowser() Then
         meta = New HtmlMeta()
         meta.Name = "viewport"
         meta.Content = "width=" & viewportWidth & ", user-scalable=yes"
         Page.Header.Controls.Add(meta)

         Dim link As New HtmlLink()
         link.Attributes("rel") = "apple-touch-icon"
         link.Href = "favicon.ico"
         Page.Header.Controls.Add(link)
      End If
   End Sub

   Private Sub SetPageCulture()

      Dim cult As String

      If Request.UserLanguages IsNot Nothing Then
         cult = Request.UserLanguages(0)
      ElseIf Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") IsNot Nothing Then
         cult = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
      Else
         cult = ConfigurationManager.AppSettings("defaultculture")
      End If

      If cult.ToLower = "en-us" Then cult = "en-GB"

      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(cult)
      Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(cult)
   End Sub
End Class
