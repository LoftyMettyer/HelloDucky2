Option Strict On

Imports System
Imports System.Data
Imports System.Globalization
Imports System.Threading
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports Utilities
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Transactions


Public Class _Default
    Inherits System.Web.UI.Page

    Private miInstanceID As Integer
    Private miElementID As Integer
    Private msServer As String
    Private msDatabase As String
    Private miImageCount As Int16
    Private msUser As String
    Private msPwd As String
    Private mobjConfig As New Config
    Private msRefreshLiteralsCode As String
    Private miCompletionMessageType As Integer
    Private msCompletionMessage As String
    Private miSavedForLaterMessageType As Integer
    Private msSavedForLaterMessage As String
    Private miFollowOnFormsMessageType As Integer
    Private msFollowOnFormsMessage As String
    Private miSubmissionTimeoutInSeconds As Int32
    Private msLastSelectedValue As String
    Private msForeColorHighlight As String
    Private msBackColorHighlight As String
    Private m_iLookupColumnIndex As Integer

    Private Const FORMINPUTPREFIX As String = "forminput_"
    Private Const ASSEMBLYNAME As String = "HRPROWORKFLOW"
    Private Const ROWHEIGHTFONTRATIO As Single = 2.5
    Private Const MAXDROPDOWNROWS As Int16 = 6

    Private Enum SQLDataType
        sqlUnknown = 0      ' ?
        sqlOle = -4         ' OLE columns
        sqlBoolean = -7     ' Logic columns
        sqlNumeric = 2      ' Numeric columns
        sqlInteger = 4      ' Integer columns
        sqlDate = 11        ' Date columns
        sqlVarChar = 12     ' Character columns
        sqlVarBinary = -3   ' Photo columns
        sqlLongVarChar = -1 ' Working Pattern columns
    End Enum

    Private Enum FilterOperators
        giFILTEROP_UNDEFINED = 0
        giFILTEROP_EQUALS = 1
        giFILTEROP_NOTEQUALTO = 2
        giFILTEROP_ISATMOST = 3
        giFILTEROP_ISATLEAST = 4
        giFILTEROP_ISMORETHAN = 5
        giFILTEROP_ISLESSTHAN = 6
        giFILTEROP_ON = 7
        giFILTEROP_NOTON = 8
        giFILTEROP_AFTER = 9
        giFILTEROP_BEFORE = 10
        giFILTEROP_ONORAFTER = 11
        giFILTEROP_ONORBEFORE = 12
        giFILTEROP_CONTAINS = 13
        giFILTEROP_IS = 14
        giFILTEROP_DOESNOTCONTAIN = 15
        giFILTEROP_ISNOT = 16
    End Enum

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()

        ScriptManager.GetCurrent(Page).AsyncPostBackTimeout = SubmissionTimeout()








    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ctlForm_Date As Infragistics.WebUI.WebSchedule.WebDateChooser
        Dim ctlForm_Button As Infragistics.WebUI.WebDataInput.WebImageButton
        Dim ctlForm_InputButton As Button
        Dim ctlForm_Label As Label
        Dim ctlForm_TextInput As TextBox
        Dim ctlForm_CheckBox As LiteralControl
        Dim ctlForm_CheckBoxReal As CheckBox
        Dim ctlForm_Dropdown As System.Web.UI.WebControls.DropDownList  'Infragistics.WebUI.WebCombo.WebCombo
        Dim ctlForm_Image As System.Web.UI.WebControls.Image
        Dim ctlForm_NumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
        'Dim ctlForm_RecordSelectionGrid As Infragistics.WebUI.UltraWebGrid.UltraWebGrid
        Dim ctlForm_PagingGridView As RecordSelector
        Dim ctlForm_Frame As LiteralControl
        Dim ctlForm_Line As LiteralControl
        Dim ctlForm_OptionGroup As LiteralControl
        Dim ctlForm_OptionGroupReal As TextBox
        Dim ctlForm_HiddenField As HiddenField
        Dim ctlForm_Literal As LiteralControl
        Dim ctlForm_DDE As AjaxControlToolkit.DropDownExtender
        Dim ctlForm_UpdatePanel As System.Web.UI.UpdatePanel
        Dim sBackgroundImage As String
        Dim sBackgroundRepeat As String
        Dim sBackgroundPosition As String
        Dim iBackgroundColour As Integer
        Dim sBackgroundColourHex As String
        Dim iBackgroundImagePosition As Integer
        Dim sAssemblyName As String
        Dim sWebSiteVersion As String
        Dim sMessage As String
        Dim sAcceptLanguage As String
        Dim sQueryString As String
        Dim objCrypt As New Crypt
        Dim objGeneral As New General
        Dim blnLocked As Boolean
        Dim strConn As String
        Dim conn As System.Data.SqlClient.SqlConnection
        Dim cmdCheck As System.Data.SqlClient.SqlCommand
        Dim cmdSelect As System.Data.SqlClient.SqlCommand
        Dim cmdInitiate As System.Data.SqlClient.SqlCommand
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim iTemp As Integer
        Dim sTemp As String = String.Empty
        Dim sTemp2 As String
        Dim sTemp3 As String
        Dim sDBVersion As String
        Dim sID As String
        Dim sImageFileName As String
        Dim sBackColour As String
        Dim objNumberFormatInfo As NumberFormatInfo
        Dim dtDate As Date
        Dim iYear As Int16
        Dim iMonth As Int16
        Dim iDay As Int16
        Dim objGridColumn As System.Data.DataColumn 'Infragistics.WebUI.UltraWebGrid.UltraGridColumn
        ' Dim objGridCell As Infragistics.WebUI.UltraWebGrid.UltraGridCell
        Dim iHeaderHeight As Int32
        Dim iTempHeight As Int32
        Dim iTempWidth As Int32
        Dim connGrid As System.Data.SqlClient.SqlConnection
        Dim drGrid As System.Data.SqlClient.SqlDataReader
        Dim cmdGrid As System.Data.SqlClient.SqlCommand
        Dim cmdQS As System.Data.SqlClient.SqlCommand
        Dim iMinTabIndex As Integer
        Dim sDefaultValue As String
        Dim fRecordOK As Boolean
        Dim iGridTopPadding As Integer
        Dim iRowHeight As Integer
        Dim iDropHeight As Integer
        Dim iYOffset As Integer
        Dim sDefaultFocusControl As String
        Dim ctlDefaultFocusControl As Control
        Dim fChecked As Boolean
        Dim ctlFormCheckBox As CheckBox
        Dim iWorkflowID As Integer
        Dim sFormElements As String
        Dim arrFollowOnForms() As String
        Dim iFollowOnFormCount As Integer
        Dim iIndex As Integer
        Dim sStep As String
        Dim arrQueryStrings() As String
        Dim sSiblingForms As String
        Dim iFormHeight As Integer
        Dim iFormWidth As Integer
        Dim sTitle As String
        Dim sMessage1 As String
        Dim sMessage2 As String
        Dim sMessage3 As String
        Dim sDecoration As String
        Dim sEncodedID As String
        Dim iMaxLength As Integer
        Dim sFilterSQL As String
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim objDataRow As System.Data.DataRow
        Dim iItemType As Integer

        Const sDEFAULTTITLE As String = "HR Pro Workflow"
        Const IMAGEBORDERWIDTH As Integer = 2

        sAssemblyName = ""
        sWebSiteVersion = ""
        sMessage = ""
        sMessage1 = ""
        sMessage2 = ""
        sMessage3 = ""
        sQueryString = ""
        miImageCount = 0
        sDefaultFocusControl = ""
        iMinTabIndex = -1
        msRefreshLiteralsCode = ""
        ReDim arrQueryStrings(0)
        sSiblingForms = ""
        sTitle = sDEFAULTTITLE



        Try
            mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))
            miSubmissionTimeoutInSeconds = mobjConfig.SubmissionTimeoutInSeconds

            Response.CacheControl = "no-cache"
            Response.AddHeader("Pragma", "no-cache")
            Response.Expires = -1

            If Not IsPostBack Then
                Session.Clear()
                Session("TimeoutSecs") = Session.Timeout * 60
            End If
        Catch ex As Exception
        End Try

        Try
            sAssemblyName = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToUpper

            sWebSiteVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
             & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
             & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

            If sAssemblyName = ASSEMBLYNAME Then
                ' Compiled version of the web site, so perform version checks.
                If sWebSiteVersion.Length = 0 Then
                    sTitle = sDEFAULTTITLE & " (unknown version)"
                Else
                    sTitle = sDEFAULTTITLE & " - v" & sWebSiteVersion
                End If
            Else
                ' Development version of the web site, so do NOT perform version checks.
                sTitle = sDEFAULTTITLE & " (development)"
            End If
        Catch ex As Exception
            sTitle = sDEFAULTTITLE
        End Try
        Page.Title = sTitle

        Try
            If Request.UserLanguages IsNot Nothing Then
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(Request.UserLanguages(0))
                Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(Request.UserLanguages(0))
            Else
                If Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") IsNot Nothing Then
                    sAcceptLanguage = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
                Else
                    ' Cannot read the client culture from the request. 
                    ' Use the default culture from the config file.
                    sAcceptLanguage = System.Configuration.ConfigurationManager.AppSettings("defaultculture")
                End If

                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(sAcceptLanguage)
                Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(sAcceptLanguage)
            End If
        Catch ex As Exception
            sMessage = "Error reading the client culture:<BR><BR>" & ex.Message
        End Try

        If sMessage.Length = 0 Then
            If IsPostBack Then
                miInstanceID = CInt(Session("InstanceID"))
                miElementID = CInt(Session("ElementID"))
                msUser = Session("User").ToString
                msPwd = Session("Pwd").ToString
                msServer = Session("Server").ToString
                msDatabase = Session("Database").ToString

                miInstanceID = CInt(Me.ViewState("InstanceID"))
                miElementID = CInt(Me.ViewState("ElementID"))
                msUser = Me.ViewState("User").ToString
                msPwd = Me.ViewState("Pwd").ToString
                msServer = Me.ViewState("Server").ToString
                msDatabase = Me.ViewState("Database").ToString

            Else
                Try
                    ' Read and decrypt the queryString.
                    ' Use the rawURL rather than the QueryString itself, as some of the 
                    ' encryption characters are ignored in the QueryString.
                    miElementID = 0
                    miInstanceID = 0

                    sTemp = Request.RawUrl.ToString
                    iTemp = sTemp.IndexOf("?")

                    If iTemp >= 0 Then
                        sQueryString = sTemp.Substring(iTemp + 1)
                    End If

                    ' Try the newer encryption first
                    Try
                        ' Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
                        Dim sCultureName As String
                        sCultureName = Thread.CurrentThread.CurrentCulture.Name

                        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-gb")
                        Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-gb")

                        sTemp = objCrypt.DecompactString(sQueryString)
                        sTemp = objCrypt.DecryptString(sTemp, "", True)

                        ' Reset the culture to be the one used by the client. Fault HRPRO-1404
                        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(sCultureName)
                        Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(sCultureName)

                        ' Extract the required parameters from the decrypted queryString.
                        miInstanceID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        miElementID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msUser = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msPwd = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msServer = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msDatabase = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                    Catch ex As Exception
                        ' Older encryption method
                        sQueryString = objCrypt.ProcessDecryptString(sQueryString)
                        sTemp = objCrypt.DecryptString(sQueryString, "", False)

                        ' Extract the required parameters from the decrypted queryString.
                        If miInstanceID = 0 Then
                            miInstanceID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
                        End If
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        If miElementID = 0 Then
                            miElementID = CInt(Left(sTemp, InStr(sTemp, vbTab) - 1))
                        End If
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msUser = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msPwd = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msServer = Left(sTemp, InStr(sTemp, vbTab) - 1)
                        sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                        msDatabase = Mid(sTemp, InStr(sTemp, vbTab) + 1)
                    End Try
                Catch theError As Exception
                    sMessage = "Invalid query string."
                End Try
            End If
        End If

        If sMessage.Length = 0 Then
            Try ' conn creation 
                strConn = "Application Name=HR Pro Workflow;Data Source=" & msServer & ";Initial Catalog=" & msDatabase & ";Integrated Security=false;User ID=" & msUser & ";Password=" & msPwd & ";Pooling=false"
                conn = New SqlClient.SqlConnection(strConn)
                conn.Open()
                Try
                    If (sMessage.Length = 0) _
                     And (Not IsPostBack) Then

                        ' Check if the database is locked.
                        cmdCheck = New SqlClient.SqlCommand
                        cmdCheck.CommandText = "sp_ASRLockCheck"
                        cmdCheck.Connection = conn
                        cmdCheck.CommandType = CommandType.StoredProcedure
                        cmdCheck.CommandTimeout = miSubmissionTimeoutInSeconds

                        dr = cmdCheck.ExecuteReader()

                        blnLocked = False
                        While dr.Read
                            If NullSafeInteger(dr("priority")) <> 3 Then
                                ' Not a read-only lock.
                                blnLocked = True
                                sMessage = "Database locked.<BR><BR>Contact your system administrator."
                                Exit While
                            End If
                        End While

                        dr.Close()
                        cmdCheck.Dispose()
                    End If

                    If (sMessage.Length = 0) _
                     And (Not IsPostBack) Then

                        ' Check if the database and website versions match.
                        cmdCheck = New SqlClient.SqlCommand
                        cmdCheck.CommandText = "spASRGetSetting"
                        cmdCheck.Connection = conn
                        cmdCheck.CommandType = CommandType.StoredProcedure
                        cmdCheck.CommandTimeout = miSubmissionTimeoutInSeconds

                        cmdCheck.Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                        cmdCheck.Parameters("@psSection").Value = "database"

                        cmdCheck.Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                        cmdCheck.Parameters("@psKey").Value = "version"

                        cmdCheck.Parameters.Add("@psDefault", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                        cmdCheck.Parameters("@psDefault").Value = ""

                        cmdCheck.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
                        cmdCheck.Parameters("@pfUserSetting").Value = False

                        cmdCheck.Parameters.Add("@psResult", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                        cmdCheck.ExecuteNonQuery()

                        sDBVersion = CStr(cmdCheck.Parameters("@psResult").Value)

                        If sAssemblyName = ASSEMBLYNAME Then
                            ' Complied version of the web site, so perform version checks.
                            If sWebSiteVersion.Length > 0 Then
                                ' Just get the major and minor parts of the 4 part version.
                                sWebSiteVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major & _
                                 "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor
                            End If

                            If (sDBVersion <> sWebSiteVersion) _
                             Or (sWebSiteVersion.Length = 0) Then
                                ' Version mismatch.
                                If sDBVersion.Length = 0 Then
                                    sDBVersion = "&lt;unknown&gt;"
                                End If
                                If sWebSiteVersion.Length = 0 Then
                                    sWebSiteVersion = "&lt;unknown&gt;"
                                End If

                                sMessage = "The Workflow website version (" & sWebSiteVersion & ")" & _
                                 " is incompatible with the database version (" & sDBVersion & ")." & _
                                 "<BR><BR>Contact your system administrator."
                            End If
                        End If

                        cmdCheck.Dispose()
                    End If

                    If (sMessage.Length = 0) _
                     And (miInstanceID < 0) _
                     And (miElementID = -1) _
                     And (Not IsPostBack) Then

                        ' Externally initiated Workflow.
                        iWorkflowID = -miInstanceID

                        cmdInitiate = New SqlClient.SqlCommand
                        cmdInitiate.CommandText = "spASRInstantiateWorkflow"
                        cmdInitiate.Connection = conn
                        cmdInitiate.CommandType = CommandType.StoredProcedure
                        cmdInitiate.CommandTimeout = miSubmissionTimeoutInSeconds

                        cmdInitiate.Parameters.Add("@piWorkflowID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdInitiate.Parameters("@piWorkflowID").Value = iWorkflowID

                        cmdInitiate.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdInitiate.Parameters.Add("@psFormElements", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        cmdInitiate.Parameters.Add("@psMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                        cmdInitiate.ExecuteNonQuery()

                        miInstanceID = NullSafeInteger(cmdInitiate.Parameters("@piInstanceID").Value)
                        sFormElements = CStr(cmdInitiate.Parameters("@psFormElements").Value())
                        sMessage = NullSafeString(cmdInitiate.Parameters("@psMessage").Value)

                        cmdInitiate.Dispose()

                        If sMessage.Length = 0 Then
                            If sFormElements.Length = 0 Then
                                sMessage = "Workflow initiated successfully."
                            Else
                                arrFollowOnForms = sFormElements.Split(CChar(vbTab))
                                iFollowOnFormCount = arrFollowOnForms.GetUpperBound(0)

                                For iIndex = 0 To iFollowOnFormCount - 1
                                    sStep = arrFollowOnForms(iIndex)

                                    If iIndex = 0 Then
                                        miElementID = CInt(sStep)
                                    Else
                                        cmdQS = New SqlClient.SqlCommand("spASRGetWorkflowQueryString", conn)
                                        cmdQS.CommandType = CommandType.StoredProcedure
                                        cmdQS.CommandTimeout = miSubmissionTimeoutInSeconds

                                        cmdQS.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                        cmdQS.Parameters("@piInstanceID").Value = miInstanceID

                                        cmdQS.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                                        cmdQS.Parameters("@piElementID").Value = CLng(sStep)

                                        cmdQS.Parameters.Add("@psQueryString", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                        cmdQS.ExecuteNonQuery()

                                        sQueryString = CStr(cmdQS.Parameters("@psQueryString").Value())

                                        ReDim Preserve arrQueryStrings(arrQueryStrings.GetUpperBound(0) + 1)
                                        arrQueryStrings(arrQueryStrings.GetUpperBound(0)) = sQueryString

                                        cmdQS.Dispose()
                                    End If
                                Next iIndex

                                sSiblingForms = Join(arrQueryStrings, vbTab)
                            End If
                        Else
                            sMessage = "Error:<BR><BR>" & sMessage
                        End If
                    End If

                    If sMessage.Length = 0 Then
                        ' Remember the useful parameters for use in postbacks.

                        Me.ViewState("InstanceID") = miInstanceID
                        Me.ViewState("ElementID") = miElementID
                        Me.ViewState("User") = msUser
                        Me.ViewState("Pwd") = msPwd
                        Me.ViewState("Server") = msServer
                        Me.ViewState("Database") = msDatabase

                        Session("User") = msUser
                        Session("Pwd") = msPwd
                        Session("Server") = msServer
                        Session("Database") = msDatabase
                        Session("ElementID") = miElementID
                        Session("InstanceID") = miInstanceID

                        cmdSelect = New SqlClient.SqlCommand
                        cmdSelect.CommandText = "spASRGetWorkflowFormItems"
                        cmdSelect.Connection = conn
                        cmdSelect.CommandType = CommandType.StoredProcedure
                        cmdSelect.CommandTimeout = miSubmissionTimeoutInSeconds

                        cmdSelect.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdSelect.Parameters("@piInstanceID").Value = miInstanceID

                        cmdSelect.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdSelect.Parameters("@piElementID").Value = miElementID

                        cmdSelect.Parameters.Add("@psErrorMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piBackColour", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piBackImage", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piBackImageLocation", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piWidth", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piHeight", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piCompletionMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@psCompletionMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piSavedForLaterMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@psSavedForLaterMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@piFollowOnFormsMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
                        cmdSelect.Parameters.Add("@psFollowOnFormsMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output

                        dr = cmdSelect.ExecuteReader

                        Dim scriptString As String = "function pageLoad() {"

                        While (dr.Read) And (sMessage.Length = 0)
                            sID = FORMINPUTPREFIX & NullSafeString(dr("id")) & "_" & NullSafeString(dr("ItemType")) & "_"
                            sEncodedID = objCrypt.SimpleEncrypt(NullSafeString(dr("id")).ToString, Session.SessionID)

                            Select NullSafeInteger(dr("ItemType"))
                                Case 0 ' Button
                                    ctlForm_Button = New Infragistics.WebUI.WebDataInput.WebImageButton
                                    With ctlForm_Button
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        ' .StylePreset = Infragistics.WebUI.WebDataInput.ButtonStylePreset.Mac

                                        .Appearance.Style.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .Appearance.Style.BorderStyle = BorderStyle.Solid
                                        .Appearance.Style.BorderWidth = 1
                                        .Appearance.InnerBorder.StyleTop = BorderStyle.None
                                        .Appearance.Style.BorderColor = objGeneral.GetColour(9999523)
                                        .Appearance.Style.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        .FocusAppearance.Style.BorderColor = objGeneral.GetColour(562943)
                                        .FocusAppearance.Style.BackColor = objGeneral.GetColour(12775933)
                                        .HoverAppearance.Style.BorderColor = objGeneral.GetColour(562943)

                                        .Text = NullSafeString(dr("caption"))
                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                                        .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                                        .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                                        .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                                        .Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        .Height = Unit.Pixel(NullSafeInteger(dr("Height")) - 7)

                                        .ClientSideEvents.Click = "try{setPostbackMode(1);}catch(e){};"
                                    End With

                                    ' pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Button)
                                    pnlInputDiv.Controls.Add(ctlForm_Button)

                                    AddHandler ctlForm_Button.Click, AddressOf Me.ButtonClick

                                Case 1 ' Database value
                                    If (NullSafeInteger(dr("sourceItemType")) = -7) _
                                     Or (NullSafeInteger(dr("sourceItemType")) = 2) _
                                     Or (NullSafeInteger(dr("sourceItemType")) = 4) _
                                     Or (NullSafeInteger(dr("sourceItemType")) = 11) Then
                                        ' -7 = Logic
                                        ' 2, 4	= Numeric, Integer
                                        ' 11= Date
                                        ctlForm_Label = New Label
                                        With ctlForm_Label
                                            .Style("position") = "absolute"
                                            .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                            .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                            .Style("word-wrap") = "break-word"
                                            .Style("overflow") = "auto"
                                            .Style("text-align") = "left"

                                            Select Case NullSafeInteger(dr("sourceItemType"))
                                                Case -7 ' Logic
                                                    If NullSafeString(dr("value")) = String.Empty Then
                                                        .Text = "&lt;undefined&gt;"
                                                    ElseIf NullSafeString(dr("value")) = "1" Then
                                                        .Text = Boolean.TrueString
                                                    Else
                                                        .Text = Boolean.FalseString
                                                    End If

                                                Case 2, 4   ' Numeric, Integer
                                                    If IsDBNull(dr("value")) Then
                                                        sTemp = "&lt;undefined&gt;"
                                                    Else
                                                        sTemp = CStr(dr("value")).Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                                                    End If
                                                    If sTemp.Chars(0) = "-" Then
                                                        sTemp = sTemp.Substring(1) & "-"
                                                    End If
                                                    .Text = sTemp
                                                    .Style("direction") = "rtl"

                                                Case 11 ' Date
                                                    If NullSafeString(dr("value")) = String.Empty Then
                                                        .Text = "&lt;undefined&gt;"
                                                    ElseIf CStr(dr("value")).Trim.Length = 0 Then
                                                        .Text = "&lt;undefined&gt;"
                                                    Else
                                                        .Text = objGeneral.ConvertSQLDateToLocale(NullSafeString(dr("value")))
                                                    End If
                                            End Select

                                            If NullSafeInteger(dr("BackStyle")) = 0 Then
                                                .BackColor = System.Drawing.Color.Transparent
                                            Else
                                                .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                            End If

                                            .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                            .Font.Name = NullSafeString(dr("FontName"))
                                            .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                            .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                                            .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                                            .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                                            .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                                            iTempHeight = NullSafeInteger(dr("Height"))
                                            iTempWidth = NullSafeInteger(dr("Width"))

                                            If NullSafeBoolean(dr("PictureBorder")) Then
                                                .BorderStyle = BorderStyle.Solid
                                                .BorderColor = objGeneral.GetColour(5730458)
                                                .BorderWidth = Unit.Pixel(1)

                                                iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                                                iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                                            End If

                                            .Height() = Unit.Pixel(iTempHeight)
                                            .Width() = Unit.Pixel(iTempWidth)

                                        End With

                                        'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Label)
                                        pnlInputDiv.Controls.Add(ctlForm_Label)

                                    Else
                                        ' Text
                                        ctlForm_TextInput = New TextBox
                                        With ctlForm_TextInput
                                            .TabIndex = -1

                                            .Style("position") = "absolute"
                                            .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                            .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                            .Style("word-wrap") = "break-word"
                                            .Style("overflow") = "auto"
                                            .Style("text-align") = "left"
                                            .TextMode = TextBoxMode.MultiLine
                                            .Wrap = True
                                            .ReadOnly = True

                                            .Text = NullSafeString(dr("value"))

                                            If NullSafeInteger(dr("BackStyle")) = 0 Then
                                                .BackColor = System.Drawing.Color.Transparent
                                            Else
                                                .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                            End If
                                            .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                            .Font.Name = NullSafeString(dr("FontName"))
                                            .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                            .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                            .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                            .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                            .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                            iTempHeight = NullSafeInteger(dr("Height"))
                                            iTempWidth = NullSafeInteger(dr("Width"))

                                            If NullSafeBoolean(dr("PictureBorder")) Then
                                                .BorderStyle = BorderStyle.Solid
                                                .BorderColor = objGeneral.GetColour(5730458)
                                                .BorderWidth = Unit.Pixel(1)

                                                iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                                                iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                                            Else
                                                .BorderStyle = BorderStyle.None
                                            End If

                                            .Height() = Unit.Pixel(iTempHeight)
                                            .Width() = Unit.Pixel(iTempWidth)

                                        End With

                                        'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_TextInput)
                                        pnlInputDiv.Controls.Add(ctlForm_TextInput)
                                    End If

                                Case 2 ' Label
                                    ctlForm_Label = New Label
                                    With ctlForm_Label
                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        .Style("word-wrap") = "break-word"
                                        .Style("overflow") = "auto"
                                        .Style("text-align") = "left"

                                        If NullSafeInteger(dr("captionType")) = 3 Then
                                            ' Calculated caption
                                            .Text = NullSafeString(dr("value"))
                                        Else
                                            .Text = NullSafeString(dr("caption"))
                                        End If

                                        If NullSafeInteger(dr("BackStyle")) = 0 Then
                                            .BackColor = System.Drawing.Color.Transparent
                                        Else
                                            .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        End If
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                                        .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                                        .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                                        .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                                        iTempHeight = NullSafeInteger(dr("Height"))
                                        iTempWidth = NullSafeInteger(dr("Width"))

                                        If NullSafeBoolean(dr("PictureBorder")) Then
                                            .BorderStyle = BorderStyle.Solid
                                            .BorderColor = objGeneral.GetColour(5730458)
                                            .BorderWidth = Unit.Pixel(1)

                                            iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                                            iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                                        End If

                                        .Height() = Unit.Pixel(iTempHeight)
                                        .Width() = Unit.Pixel(iTempWidth)
                                    End With

                                    'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Label)
                                    pnlInputDiv.Controls.Add(ctlForm_Label)


                                Case 3 ' Input value - character
                                    ctlForm_TextInput = New TextBox
                                    With ctlForm_TextInput
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                            ctlDefaultFocusControl = ctlForm_TextInput
                                        End If

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        .Style("word-wrap") = "break-word"
                                        .Style("overflow") = "auto"

                                        If NullSafeBoolean(dr("PasswordType")) Then
                                            .TextMode = TextBoxMode.Password
                                        Else
                                            .TextMode = TextBoxMode.MultiLine
                                        End If
                                        .Wrap = True

                                        .Text = NullSafeString(dr("value"))

                                        .BorderStyle = BorderStyle.Solid
                                        .BorderWidth = Unit.Pixel(1)
                                        .BorderColor = objGeneral.GetColour(5730458)

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                        .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 6)
                                        .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 6)

                                        .Attributes("onfocus") = "try{" & sID & ".select();activateControl();}catch(e){};"
                                        .Attributes("onkeydown") = "try{checkMaxLength(" & NullSafeString(dr("inputSize")) & ");}catch(e){}"
                                        .Attributes("onpaste") = "try{checkMaxLength(" & NullSafeString(dr("inputSize")) & ");}catch(e){}"

                                    End With

                                    'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_TextInput)
                                    pnlInputDiv.Controls.Add(ctlForm_TextInput)

                                Case 4 ' Workflow value
                                    If (NullSafeInteger(dr("sourceItemType")) = 6) _
                                     Or (NullSafeInteger(dr("sourceItemType")) = 5) _
                                     Or (NullSafeInteger(dr("sourceItemType")) = 7) Then
                                        ' 6 = Logic
                                        ' 5 = Number
                                        ' 7 = Date

                                        ctlForm_Label = New Label
                                        With ctlForm_Label
                                            .Style("position") = "absolute"
                                            .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                            .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                            .Style("word-wrap") = "break-word"
                                            .Style("overflow") = "auto"
                                            .Style("text-align") = "left"

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
                                                        sTemp = String.Empty
                                                    Else
                                                        sTemp = NullSafeString(dr("value")).Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                                                    End If

                                                    If sTemp.Length > 0 AndAlso sTemp.Chars(0) = "-" Then
                                                        sTemp = sTemp.Substring(1) & "-"
                                                    End If
                                                    .Text = sTemp
                                                    .Style("direction") = "rtl"

                                                Case 7 ' Date
                                                    If IsDBNull(dr("value")) Then
                                                        .Text = "&lt;undefined&gt;"
                                                    ElseIf CStr(dr("value")).Trim.ToUpper = "NULL" Then
                                                        .Text = "&lt;undefined&gt;"
                                                    Else
                                                        .Text = objGeneral.ConvertSQLDateToLocale(NullSafeString(dr("value")))
                                                    End If
                                            End Select

                                            If NullSafeInteger(dr("BackStyle")) = 0 Then
                                                .BackColor = System.Drawing.Color.Transparent
                                            Else
                                                .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                            End If
                                            .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                            .Font.Name = NullSafeString(dr("FontName"))
                                            .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                            .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                            .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                            .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                            .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                            iTempHeight = NullSafeInteger(dr("Height"))
                                            iTempWidth = NullSafeInteger(dr("Width"))

                                            If NullSafeBoolean(dr("PictureBorder")) Then
                                                .BorderStyle = BorderStyle.Solid
                                                .BorderColor = objGeneral.GetColour(5730458)
                                                .BorderWidth = Unit.Pixel(1)

                                                iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                                                iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                                            End If

                                            .Height() = Unit.Pixel(iTempHeight)
                                            .Width() = Unit.Pixel(iTempWidth)

                                        End With

                                        'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Label)
                                        pnlInputDiv.Controls.Add(ctlForm_Label)
                                    Else
                                        ' Text
                                        ctlForm_TextInput = New TextBox
                                        With ctlForm_TextInput
                                            .TabIndex = -1

                                            .Style("position") = "absolute"
                                            .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                            .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                            .Style("word-wrap") = "break-word"
                                            .Style("overflow") = "auto"
                                            .Style("text-align") = "left"
                                            .TextMode = TextBoxMode.MultiLine
                                            .Wrap = True
                                            .ReadOnly = True

                                            .Text = NullSafeString(dr("value"))

                                            If NullSafeInteger(dr("BackStyle")) = 0 Then
                                                .BackColor = System.Drawing.Color.Transparent
                                            Else
                                                .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                            End If
                                            .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                            .Font.Name = NullSafeString(dr("FontName"))
                                            .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                            .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                            .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                            .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                            .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                            iTempHeight = NullSafeInteger(dr("Height"))
                                            iTempWidth = NullSafeInteger(dr("Width"))

                                            If NullSafeBoolean(dr("PictureBorder")) Then
                                                .BorderStyle = BorderStyle.Solid
                                                .BorderColor = objGeneral.GetColour(5730458)
                                                .BorderWidth = Unit.Pixel(1)

                                                iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                                                iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                                            Else
                                                .BorderStyle = BorderStyle.None
                                            End If

                                            .Height() = Unit.Pixel(iTempHeight)
                                            .Width() = Unit.Pixel(iTempWidth)
                                        End With

                                        'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_TextInput)
                                        pnlInputDiv.Controls.Add(ctlForm_TextInput)
                                    End If

                                Case 5 ' Input value - numeric
                                    ctlForm_NumericInput = New Infragistics.WebUI.WebDataInput.WebNumericEdit
                                    With ctlForm_NumericInput
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = ""
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                            ctlDefaultFocusControl = ctlForm_NumericInput
                                        End If

                                        .SelectionOnFocus = Infragistics.WebUI.WebDataInput.SelectionOnFocus.SelectAll

                                        objNumberFormatInfo = DirectCast(Thread.CurrentThread.CurrentCulture.NumberFormat.Clone, NumberFormatInfo)
                                        objNumberFormatInfo.NumberDecimalDigits = NullSafeInteger(dr("inputDecimals"))
                                        objNumberFormatInfo.NumberGroupSeparator = ""
                                        .NumberFormat = objNumberFormatInfo

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        iMaxLength = NullSafeInteger(dr("inputSize")) + 1   ' Add 1 for the minus sign.
                                        If NullSafeInteger(dr("inputDecimals")) > 0 Then
                                            iMaxLength = iMaxLength + 1 ' Add 1 for the decimal point if required.
                                        End If
                                        .MaxLength = iMaxLength
                                        .MinDecimalPlaces = DirectCast(NullSafeInteger(dr("inputDecimals")), Infragistics.WebUI.WebDataInput.MinDecimalPlaces)
                                        .MaxValue = (10 ^ (NullSafeInteger(dr("inputSize")) - NullSafeInteger(dr("inputDecimals")))) - 1 + (1 - (1 / (10 ^ NullSafeInteger(dr("inputDecimals")))))
                                        .MinValue = (-1 * .MaxValue)
                                        .DataMode = Infragistics.WebUI.WebDataInput.NumericDataMode.Decimal

                                        .Text = NullSafeString(dr("value"))

                                        .Nullable = False

                                        .BorderColor = objGeneral.GetColour(5730458)
                                        .BorderStyle = BorderStyle.Solid
                                        .BorderWidth = Unit.Pixel(1)

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                        .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 6)
                                        .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 6)

                                        .Attributes("onfocus") = "try{" & sID & ".select();activateControl();}catch(e){};"

                                        .ClientSideEvents.KeyPress = "WebNumericEditValidation_KeyPress"
                                        .ClientSideEvents.KeyDown = "WebNumericEditValidation_KeyDown"
                                        .Attributes("onpaste") = "try{WebNumericEditValidation_Paste(this, event, '" & sID & "');}catch(e){};"

                                    End With

                                    ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_NumericInput)
                                    pnlInputDiv.Controls.Add(ctlForm_NumericInput)

                                Case 6 ' Input value - logic
                                    ' NB. We use a table with a label and checkbox in, instead of just a checkbox
                                    ' for formatting purposes.
                                    ctlForm_CheckBoxReal = New CheckBox
                                    With ctlForm_CheckBoxReal
                                        .Height = Unit.Parse("0")
                                        .Width = Unit.Parse("0")
                                        .TabIndex = 0
                                        .Style("visibility") = "hidden"
                                        .Checked = (NullSafeString(dr("value")).ToUpper = "TRUE")
                                        .ID = sID
                                    End With
                                    ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_CheckBoxReal)
                                    pnlInputDiv.Controls.Add(ctlForm_CheckBoxReal)

                                    msRefreshLiteralsCode = msRefreshLiteralsCode & vbNewLine & _
                                     vbTab & vbTab & "try" & vbNewLine & _
                                     vbTab & vbTab & "{" & vbNewLine & _
                                     vbTab & vbTab & vbTab & "frmMain.chk" & sID & ".checked = frmMain." & sID & ".checked;" & vbNewLine & _
                                     vbTab & vbTab & "}" & vbNewLine & _
                                     vbTab & vbTab & "catch(e) {}"

                                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                                        sBackColour = "Transparent"
                                    Else
                                        sBackColour = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColor")))
                                    End If

                                    sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                                     CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                                    If sTemp2.Length = 0 Then
                                        sTemp2 = " none"
                                    End If

                                    sTemp = "<TABLE BORDER='0' CELLSPACING='0' CELLPADDING='0'" & _
                                     " WIDTH=" & NullSafeString(dr("Width")) & _
                                     " style='DISPLAY: inline-block; POSITION: absolute; TEXT-ALIGN: Left;" & _
                                     " TOP: " & NullSafeString(dr("TopCoord")) & "px; " & " LEFT: " & NullSafeString(dr("LeftCoord")) & "px; " & _
                                     " WIDTH: " & NullSafeString(dr("Width")) & "px; " & " HEIGHT: " & NullSafeString(dr("Height")) & "px; " & _
                                     " BACKGROUND-COLOR: " & sBackColour & "; " & _
                                     " COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                                     " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                     " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                     " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                     " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                     " TEXT-DECORATION:" & sTemp2 & "'>" & vbCrLf & _
                                     "<TR>" & vbCrLf

                                    If IsPostBack Then
                                        If pnlInput.FindControl(sID) Is Nothing Then
                                            fChecked = True
                                        Else
                                            ctlFormCheckBox = DirectCast(pnlInput.FindControl(sID), CheckBox)
                                            fChecked = ctlFormCheckBox.Checked
                                        End If

                                        If NullSafeInteger(dr("alignment")) = 0 Then
                                            sTemp = sTemp & _
                                             "<TD width='1px'><input type='checkbox'" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onclick=""" & sID & ".checked = checked;""" & _
                                             " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                                             CStr(IIf(fChecked, " CHECKED", "")) & _
                                             " style='height:14px;width:14px;'" & _
                                             " tabIndex='-1'" & _
                                             " id='chk" & sID & "'" & _
                                             " name='chk" & sID & "'></TD>" & vbCrLf & _
                                             "<TD width='4px'></TD><TD><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                                             " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             ">&nbsp;&nbsp;" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf
                                        Else
                                            sTemp = sTemp & _
                                             "<TD><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                                             " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf & _
                                             "<TD width='1px'><input type='checkbox'" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onclick=""" & sID & ".checked = checked;""" & _
                                             " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                                             CStr(IIf(fChecked, " CHECKED", "")) & _
                                             " style='height:14px;width:14px;'" & _
                                             " tabIndex='-1'" & _
                                             " id='chk" & sID & "'" & _
                                             " name='chk" & sID & "'></TD>" & vbCrLf
                                        End If
                                    Else
                                        If NullSafeInteger(dr("alignment")) = 0 Then
                                            sTemp = sTemp & _
                                             "<TD width='1px'><input type='checkbox'" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onclick=""" & sID & ".checked = checked;""" & _
                                             " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                                             CStr(IIf(UCase(NullSafeString(dr("value"))) = "TRUE", " CHECKED", "")) & _
                                             " style='height:14px;width:14px;'" & _
                                             " tabIndex='-1'" & _
                                             " id='chk" & sID & "'" & _
                                             " name='chk" & sID & "'></TD>" & vbCrLf & _
                                             "<TD width='4px'></TD><TD><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                                             " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             ">&nbsp;&nbsp;" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf
                                        Else
                                            sTemp = sTemp & _
                                             "<TD><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                                             " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf & _
                                             "<TD width='1px'><input type='checkbox'" & _
                                             " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                                             " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                                             " onclick=""" & sID & ".checked = checked;""" & _
                                             " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                                             CStr(IIf(NullSafeString(dr("value")).ToUpper = "TRUE", " CHECKED", "")) & _
                                             " style='height:14px;width:14px;'" & _
                                             " tabIndex='-1'" & _
                                             " id='chk" & sID & "'" & _
                                             " name='chk" & sID & "'></TD>" & vbCrLf
                                        End If
                                    End If

                                    sTemp = sTemp & _
                                     "</TR>" & vbCrLf & _
                                     "</TABLE>"

                                    ctlForm_CheckBox = New LiteralControl(sTemp)
                                    ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_CheckBox)
                                    pnlInputDiv.Controls.Add(ctlForm_CheckBox)

                                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                        sDefaultFocusControl = "chk" & sID
                                        iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                    End If

                                Case 7 ' Input value - date
                                    ctlForm_Date = New Infragistics.WebUI.WebSchedule.WebDateChooser
                                    With ctlForm_Date
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                            ctlDefaultFocusControl = ctlForm_Date
                                        End If

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        .CalendarLayout.FooterFormat = "Today: {0:d}"
                                        .CalendarLayout.FirstDayOfWeek = WebControls.FirstDayOfWeek.Sunday
                                        .CalendarLayout.ShowTitle = False

                                        .CalendarLayout.DayStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.DayStyle.Font.Name = "Verdana"
                                        .CalendarLayout.DayStyle.ForeColor = objGeneral.GetColour(6697779)
                                        .CalendarLayout.DayStyle.BackColor = objGeneral.GetColour(15988214)

                                        .CalendarLayout.FooterStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.FooterStyle.Font.Name = "Verdana"
                                        .CalendarLayout.FooterStyle.ForeColor = objGeneral.GetColour(6697779)
                                        .CalendarLayout.FooterStyle.BackColor = objGeneral.GetColour(16248553)

                                        .CalendarLayout.SelectedDayStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.SelectedDayStyle.Font.Name = "Verdana"
                                        .CalendarLayout.SelectedDayStyle.Font.Bold = True
                                        .CalendarLayout.SelectedDayStyle.Font.Underline = True
                                        .CalendarLayout.SelectedDayStyle.ForeColor = objGeneral.GetColour(2774907)
                                        .CalendarLayout.SelectedDayStyle.BackColor = objGeneral.GetColour(10480637)

                                        .CalendarLayout.OtherMonthDayStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.OtherMonthDayStyle.Font.Name = "Verdana"
                                        .CalendarLayout.OtherMonthDayStyle.ForeColor = objGeneral.GetColour(11375765)

                                        .CalendarLayout.NextPrevStyle.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
                                        .CalendarLayout.NextPrevStyle.BackColor = objGeneral.GetColour(16248553)
                                        .CalendarLayout.NextPrevStyle.ForeColor = objGeneral.GetColour(6697779)

                                        .CalendarLayout.CalendarStyle.Width = Unit.Pixel(152)
                                        .CalendarLayout.CalendarStyle.Height = Unit.Pixel(80)
                                        .CalendarLayout.CalendarStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.CalendarStyle.Font.Name = "Verdana"
                                        .CalendarLayout.CalendarStyle.BackColor = System.Drawing.Color.White

                                        .CalendarLayout.WeekendDayStyle.BackColor = objGeneral.GetColour(15004669)

                                        .CalendarLayout.TodayDayStyle.ForeColor = objGeneral.GetColour(2774907)
                                        .CalendarLayout.TodayDayStyle.BackColor = objGeneral.GetColour(10480637)

                                        .CalendarLayout.DropDownStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.DropDownStyle.Font.Name = "Verdana"
                                        .CalendarLayout.DropDownStyle.BorderStyle = BorderStyle.Solid
                                        .CalendarLayout.DropDownStyle.BorderColor = objGeneral.GetColour(10720408)

                                        .CalendarLayout.DayHeaderStyle.BackColor = objGeneral.GetColour(16248553)
                                        .CalendarLayout.DayHeaderStyle.ForeColor = objGeneral.GetColour(6697779)
                                        .CalendarLayout.DayHeaderStyle.Font.Size = FontUnit.Parse(CStr(8))
                                        .CalendarLayout.DayHeaderStyle.Font.Name = "Verdana"
                                        .CalendarLayout.DayHeaderStyle.Font.Bold = True

                                        .CalendarLayout.TitleStyle.ForeColor = objGeneral.GetColour(6697779)
                                        .CalendarLayout.TitleStyle.BackColor = objGeneral.GetColour(16248553)
                                        .NullDateLabel = ""

                                        If (Not IsDBNull(dr("value"))) Then
                                            If CStr(dr("value")).Length > 0 Then
                                                iYear = CShort(NullSafeString(dr("value")).Substring(6, 4))
                                                iMonth = CShort(NullSafeString(dr("value")).Substring(0, 2))
                                                iDay = CShort(NullSafeString(dr("value")).Substring(3, 2))

                                                dtDate = DateSerial(iYear, iMonth, iDay)
                                                .Value = dtDate
                                            End If
                                        End If

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        .BorderColor = objGeneral.GetColour(5730458)

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                        .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                                        .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 2)

                                        .ClientSideEvents.EditKeyDown = "dateControlKeyPress"
                                        .ClientSideEvents.TextChanged = "dateControlTextChanged"
                                        .ClientSideEvents.BeforeDropDown = "dateControlBeforeDropDown"

                                    End With

                                    ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Date)
                                    pnlInputDiv.Controls.Add(ctlForm_Date)

                                Case 8 ' Frame
                                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                                        sBackColour = "Transparent"
                                    Else
                                        sBackColour = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColor")))
                                    End If

                                    sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                                     CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                                    If sTemp2.Length = 0 Then
                                        sTemp2 = " none"
                                    End If

                                    Dim fieldsetTopCoord As Int32 = _
                                     CInt((NullSafeInteger(dr("TopCoord")) + (NullSafeSingle(dr("FontSize")) * 2.5 / 3)))
                                    Dim fieldsetLeftCoord As Int32 = NullSafeInteger(dr("LeftCoord"))
                                    Dim fieldsetWidth As Int32 = NullSafeInteger(dr("Width")) - 4
                                    Dim fieldsetHeight As Int32 = _
                                     CInt((NullSafeInteger(dr("Height")) - 1 _
                                     - CInt(IIf(NullSafeString(dr("caption")).Trim.Length > 0, 0, 2)) _
                                     - (NullSafeSingle(dr("FontSize")) * 2.5 / 3)))

                                    sTemp = "<fieldset style='TOP: " & fieldsetTopCoord.ToString & "px; " & _
                                    " LEFT: " & fieldsetLeftCoord.ToString & "px; " & _
                                    " WIDTH: " & fieldsetWidth.ToString & "px; " & _
                                    " HEIGHT: " & fieldsetHeight.ToString & "px; " & _
                                    " BACKGROUND-COLOR: " & sBackColour & "; " & _
                                    " BORDER-STYLE: solid; " & _
                                    " BORDER-COLOR: #9894a3; " & _
                                    " BORDER-WIDTH: 1px; " & _
                                    " COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                                    " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                    " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                    " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                    " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                    " TEXT-DECORATION:" & sTemp2 & ";" & _
                                    " POSITION: absolute;padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin: 0px; padding-top: 0px;'>"

                                    If NullSafeString(dr("caption")).Trim.Length > 0 Then
                                        Dim legendTop As Int32 = CInt((NullSafeSingle(dr("FontSize")) * -11 / 10))

                                        sTemp = sTemp & _
                                        "<legend" & _
                                        " style='top: " & legendTop.ToString & _
                                        "px; COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                                        " padding-right: 2px; padding-left: 2px; padding-bottom: 0px; margin-left: 5px; padding-top: 0px; position: relative;' align='Left'>" & _
                                        NullSafeString(dr("caption")) & _
                                        "</legend>"
                                    End If

                                    sTemp = sTemp & _
                                    "</fieldset>" & vbCrLf

                                    ctlForm_Frame = New LiteralControl(sTemp)

                                    ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Frame)
                                    pnlInputDiv.Controls.Add(ctlForm_Frame)

                                Case 9 ' Line
                                    Select Case NullSafeInteger(dr("Orientation"))
                                        Case 0
                                            ' Vertical
                                            sTemp = "<IMG style='POSITION: absolute;" & _
                                             " LEFT: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                                             " TOP: " & NullSafeString(dr("TopCoord")) & "px;" & _
                                             " HEIGHT:" & NullSafeString(dr("Height")) & "px;" & _
                                             " WIDTH:0px;" & _
                                             " BORDER-TOP-STYLE:none;" & _
                                             " BORDER-RIGHT-STYLE:none;" & _
                                             " BORDER-BOTTOM-STYLE:none;" & _
                                             " BORDER-LEFT-COLOR:" & objGeneral.GetHTMLColour(NullSafeInteger(dr("Backcolor"))) & ";" & _
                                             " BORDER-LEFT-STYLE:solid;" & _
                                             " BORDER-LEFT-WIDTH:1px'/>"
                                        Case 1
                                            ' Horizontal
                                            sTemp = "<IMG style='POSITION: absolute;" & _
                                             " LEFT: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                                             " TOP: " & NullSafeString(dr("TopCoord")) & "px;" & _
                                             " HEIGHT:0px;" & _
                                             " WIDTH:" & NullSafeString(dr("Width")) & "px;" & _
                                             " BORDER-LEFT-STYLE:none;" & _
                                             " BORDER-RIGHT-STYLE:none;" & _
                                             " BORDER-BOTTOM-STYLE:none;" & _
                                             " BORDER-TOP-COLOR:" & objGeneral.GetHTMLColour(NullSafeInteger(dr("Backcolor"))) & ";" & _
                                             " BORDER-TOP-STYLE:solid;" & _
                                             " BORDER-TOP-WIDTH:1px'/>"
                                    End Select

                                    ctlForm_Line = New LiteralControl(sTemp)

                                    'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Line)
                                    pnlInputDiv.Controls.Add(ctlForm_Line)

                                Case 10 ' Image
                                    ctlForm_Image = New System.Web.UI.WebControls.Image

                                    With ctlForm_Image
                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        sImageFileName = LoadPicture(NullSafeInteger(dr("pictureID")), sMessage)
                                        If sMessage.Length > 0 Then
                                            Exit While
                                        End If
                                        .ImageUrl = sImageFileName

                                        iTempHeight = NullSafeInteger(dr("Height"))
                                        iTempWidth = NullSafeInteger(dr("Width"))

                                        If NullSafeBoolean(dr("PictureBorder")) Then
                                            .BorderStyle = BorderStyle.Solid
                                            .BorderColor = objGeneral.GetColour(10720408)
                                            .BorderWidth = 1

                                            iTempHeight = iTempHeight - 2
                                            iTempWidth = iTempWidth - 2
                                        End If

                                        .Height() = Unit.Pixel(iTempHeight)
                                        .Width() = Unit.Pixel(iTempWidth)

                                    End With

                                    'pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Image)
                                    pnlInputDiv.Controls.Add(ctlForm_Image)

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

                                    ctlForm_PagingGridView = New RecordSelector
                                    With ctlForm_PagingGridView

                                        .Style.Add("Position", "Absolute")
                                        .Attributes.CssStyle("LEFT") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        .Attributes.CssStyle("TOP") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString

                                        ' Don't use .height - it causes large row heights if the grid isn't filled.
                                        ' Use .ControlHeight instead - custom property.
                                        '.Height = NullSafeInteger(dr("Height"))
                                        ' .Attributes.Add("Height", NullSafeInteger(dr("Height")).ToString)
                                        '.Style.Add("height", Unit.Pixel(NullSafeInteger(dr("Height"))).ToString)
                                        .ControlHeight = NullSafeInteger(dr("Height"))

                                        .Width = NullSafeInteger(dr("Width"))

                                        .BorderColor = Color.Black
                                        .BorderStyle = BorderStyle.Solid
                                        .BorderWidth = 1

                                        .Style.Add("border-bottom-width", "2px")
                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("HeaderBackColor")))

                                        .ID = sID & "Grid"
                                        .AllowPaging = True
                                        .AllowSorting = True
                                        '.EnableSortingAndPagingCallbacks = True
                                        .PageSize = mobjConfig.LookupRowsRange
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

                                        ' set top row as default item
                                        .SelectedIndex = 0

                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        ''    ' .DisplayLayout.ClientSideEvents.ColumnHeaderClickHandler = "activateGridPostback"

                                        ''    '.DisplayLayout.AllowSortingDefault = Infragistics.WebUI.UltraWebGrid.AllowSorting.Yes
                                        ''    '.DisplayLayout.HeaderClickActionDefault = Infragistics.WebUI.UltraWebGrid.HeaderClickAction.SortMulti
                                        ''    '.DisplayLayout.SelectTypeRowDefault = Infragistics.WebUI.UltraWebGrid.SelectType.Single
                                        ''    '.DisplayLayout.TableLayout = Infragistics.WebUI.UltraWebGrid.TableLayout.Fixed
                                        ''    '.DisplayLayout.StationaryMargins = Infragistics.WebUI.UltraWebGrid.StationaryMargins.Header
                                        ''    '.DisplayLayout.RowStyleDefault.Cursor = Infragistics.WebUI.Shared.Cursors.Default

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))

                                        ''    '.BorderColor = objGeneral.GetColour(10720408)
                                        ''    '.BorderStyle = BorderStyle.Solid
                                        ''    '.BorderWidth = Unit.Pixel(1)

                                        ''    '.DisplayLayout.AllowColSizingDefault = Infragistics.WebUI.UltraWebGrid.AllowSizing.Free
                                        ''    '.DisplayLayout.CellClickActionDefault = Infragistics.WebUI.UltraWebGrid.CellClickAction.RowSelect
                                        ''    '.DisplayLayout.ColHeadersVisibleDefault = DirectCast(IIf(NullSafeBoolean(dr("ColumnHeaders")) And (NullSafeInteger(dr("Headlines")) > 0), _
                                        ''    ' Infragistics.WebUI.UltraWebGrid.ShowMarginInfo.Yes, _
                                        ''    ' Infragistics.WebUI.UltraWebGrid.ShowMarginInfo.No), Infragistics.WebUI.UltraWebGrid.ShowMarginInfo)
                                        ''    '.DisplayLayout.GridLinesDefault = Infragistics.WebUI.UltraWebGrid.UltraGridLines.Both

                                        .HeaderStyle.BackColor = objGeneral.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                                        .HeaderStyle.BorderColor = objGeneral.GetColour(10720408)
                                        .HeaderStyle.BorderStyle = BorderStyle.Double

                                        '.HeaderStyle.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthTop = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        '.HeaderStyle.BorderDetails.WidthRight = Unit.Pixel(1)
                                        .HeaderStyle.BorderWidth = Unit.Pixel(0)
                                        .HeaderStyle.Font.Name = NullSafeString(dr("HeadFontName"))
                                        .HeaderStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("HeadFontSize")))
                                        .HeaderStyle.Font.Bold = NullSafeBoolean(dr("HeadFontBold"))
                                        .HeaderStyle.Font.Italic = NullSafeBoolean(dr("HeadFontItalic"))
                                        .HeaderStyle.Font.Strikeout = NullSafeBoolean(dr("HeadFontStrikeThru"))
                                        .HeaderStyle.Font.Underline = NullSafeBoolean(dr("HeadFontUnderline"))
                                        .HeaderStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        '.HeaderStyle.Padding.Top = Unit.Pixel(iGridTopPadding)
                                        '.HeaderStyle.Padding.Bottom = Unit.Pixel(0)
                                        '.HeaderStyle.Padding.Left = Unit.Pixel(2)
                                        '.HeaderStyle.Padding.Right = Unit.Pixel(2)
                                        .HeaderStyle.Wrap = False
                                        .HeaderStyle.Height = Unit.Pixel(iHeaderHeight)
                                        .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                                        .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                                        ' PagerStyle settings
                                        .PagerStyle.BackColor = objGeneral.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                                        .PagerStyle.BorderColor = objGeneral.GetColour(10720408)
                                        .PagerStyle.BorderStyle = BorderStyle.Solid

                                        '.HeaderStyle.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthTop = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        '.HeaderStyle.BorderDetails.WidthRight = Unit.Pixel(1)

                                        .PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
                                        .PagerSettings.FirstPageImageUrl = "Images/page-first.gif"
                                        .PagerSettings.LastPageImageUrl = "Images/page-last.gif"
                                        .PagerSettings.NextPageImageUrl = "Images/page-next.gif"
                                        .PagerSettings.PreviousPageImageUrl = "Images/page-prev.gif"

                                        .PagerStyle.BorderWidth = Unit.Pixel(0)
                                        .PagerStyle.Font.Name = NullSafeString(dr("HeadFontName"))
                                        .PagerStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("HeadFontSize")))
                                        .PagerStyle.Font.Bold = NullSafeBoolean(dr("HeadFontBold"))
                                        .PagerStyle.Font.Italic = NullSafeBoolean(dr("HeadFontItalic"))
                                        .PagerStyle.Font.Strikeout = NullSafeBoolean(dr("HeadFontStrikeThru"))
                                        .PagerStyle.Font.Underline = NullSafeBoolean(dr("HeadFontUnderline"))
                                        .PagerStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        '.HeaderStyle.Padding.Top = Unit.Pixel(iGridTopPadding)
                                        '.HeaderStyle.Padding.Bottom = Unit.Pixel(0)
                                        '.HeaderStyle.Padding.Left = Unit.Pixel(2)
                                        '.HeaderStyle.Padding.Right = Unit.Pixel(2)
                                        .PagerStyle.Wrap = False
                                        ''.PagerStyle.Height = Unit.Pixel(iHeaderHeight)
                                        .PagerStyle.VerticalAlign = VerticalAlign.Middle
                                        .PagerStyle.HorizontalAlign = HorizontalAlign.Center

                                        ' ROW formatting
                                        .AlternatingRowStyle.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColorOdd")))
                                        .AlternatingRowStyle.Font.Name = NullSafeString(dr("FontName"))
                                        .AlternatingRowStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .AlternatingRowStyle.Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .AlternatingRowStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .AlternatingRowStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .AlternatingRowStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .AlternatingRowStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColorOdd")))
                                        '.DisplayLayout.RowAlternateStyleDefault.Padding.Left = Unit.Pixel(3)
                                        '.DisplayLayout.RowAlternateStyleDefault.Padding.Right = Unit.Pixel(3)
                                        '.DisplayLayout.RowAlternateStyleDefault.Padding.Top = Unit.Pixel(0)
                                        '.DisplayLayout.RowAlternateStyleDefault.Padding.Bottom = Unit.Pixel(1)
                                        .AlternatingRowStyle.VerticalAlign = VerticalAlign.Middle

                                        ''    '.DisplayLayout.RowSelectorsDefault = Infragistics.WebUI.UltraWebGrid.RowSelectors.No

                                        .RowStyle.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColorEven")))
                                        .RowStyle.BorderColor = objGeneral.GetColour(10720408)
                                        .RowStyle.BorderStyle = BorderStyle.Solid
                                        .RowStyle.BorderWidth = Unit.Pixel(0)
                                        '.RowStyle.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        '.RowStyle.BorderDetails.WidthTop = Unit.Pixel(0)
                                        '.RowStyle.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        '.RowStyle.BorderDetails.WidthRight = Unit.Pixel(1)
                                        .RowStyle.Font.Name = NullSafeString(dr("FontName"))
                                        .RowStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .RowStyle.Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .RowStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .RowStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .RowStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .RowStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColorEven")))
                                        '.RowStyle.Padding.Left = Unit.Pixel(3)
                                        '.RowStyle.Padding.Right = Unit.Pixel(3)
                                        '.RowStyle.Padding.Top = Unit.Pixel(0)
                                        '.RowStyle.Padding.Bottom = Unit.Pixel(1)
                                        .RowStyle.VerticalAlign = VerticalAlign.Middle

                                        iRowHeight = 21
                                        '.DisplayLayout.RowHeightDefault = Unit.Pixel(iRowHeight)

                                        ' NPG20110802 Fault HRPRO-1685, remove rowheight to autofit.
                                        '.RowStyle.Height = Unit.Pixel(iRowHeight)

                                        If IsDBNull(dr("ForeColorHighlight")) Then
                                            '.DisplayLayout.SelectedRowStyleDefault.ForeColor = System.Drawing.SystemColors.HighlightText
                                            .SelectedRowStyle.ForeColor = System.Drawing.SystemColors.HighlightText
                                            msForeColorHighlight = System.Drawing.SystemColors.HighlightText.ToString
                                        Else
                                            '.DisplayLayout.SelectedRowStyleDefault.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColorHighlight")))
                                            .SelectedRowStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColorHighlight")))
                                            msForeColorHighlight = objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColorHighlight")))
                                        End If
                                        If IsDBNull(dr("BackColorHighlight")) Then
                                            ' .DisplayLayout.SelectedRowStyleDefault.BackColor = System.Drawing.SystemColors.Highlight
                                            .SelectedRowStyle.BackColor = System.Drawing.SystemColors.Highlight
                                            msBackColorHighlight = System.Drawing.SystemColors.Highlight.ToString
                                        Else
                                            ' .DisplayLayout.SelectedRowStyleDefault.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColorHighlight")))
                                            .SelectedRowStyle.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColorHighlight")))
                                            msBackColorHighlight = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColorHighlight")))
                                        End If

                                        ''    '.DisplayLayout.ActivationObject.BorderColor = objGeneral.GetColour(10720408)
                                        ''    '.DisplayLayout.ActivationObject.BorderStyle = BorderStyle.Solid
                                        ''    '.DisplayLayout.ActivationObject.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        ''    '.DisplayLayout.ActivationObject.BorderDetails.WidthTop = Unit.Pixel(1)
                                        ''    '.DisplayLayout.ActivationObject.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        ''    '.DisplayLayout.ActivationObject.BorderDetails.WidthRight = Unit.Pixel(1)

                                        ''    'iTempHeight = NullSafeInteger(dr("Height")) - iHeaderHeight
                                        ''    'iTempHeight = CInt(IIf(iTempHeight < 0, 1, iTempHeight))
                                        ''    '.Height = Unit.Pixel(iTempHeight)
                                        ''    '.Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        ''    '.Style.Add("Position", "static")

                                        ''    ' select row 1
                                        ''    ' .SelectedIndex = 1
                                        ''    ' .SelectedRowStyle.BackColor = System.Drawing.Color.LightCyan

                                        ''    ' Add client side events to the rows as they're generated
                                        ''    AddHandler .RowDataBound, AddressOf Me.GridRowDataBound

                                    End With

                                    ' ==================================================
                                    ' Add the Paging Grid View to the holding panel now.
                                    ' Before the databind, or you'll get errors!
                                    ' ==================================================
                                    pnlInputDiv.Controls.Add(ctlForm_PagingGridView)


                                    If (Not IsPostBack) Then
                                        connGrid = New SqlClient.SqlConnection(strConn)
                                        connGrid.Open()

                                        Try
                                            cmdGrid = New SqlClient.SqlCommand
                                            cmdGrid.CommandText = "spASRGetWorkflowGridItems"
                                            cmdGrid.Connection = connGrid
                                            cmdGrid.CommandType = CommandType.StoredProcedure
                                            cmdGrid.CommandTimeout = miSubmissionTimeoutInSeconds

                                            cmdGrid.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                            cmdGrid.Parameters("@piInstanceID").Value = miInstanceID

                                            cmdGrid.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
                                            cmdGrid.Parameters("@piElementItemID").Value = NullSafeString(dr("ID"))

                                            cmdGrid.Parameters.Add("@pfOK", SqlDbType.Bit).Direction = ParameterDirection.Output

                                            'drGrid = cmdGrid.ExecuteReader()
                                            da = New SqlDataAdapter(cmdGrid)
                                            dt = New DataTable()

                                            ' Fill the datatable with data from the datadapter.
                                            da.Fill(dt)

                                            Session(sID & "DATA") = dt

                                            ' NOTE: Do the dataBind() after adding to the panel
                                            ' otherwise you get an error.
                                            ' ctlForm_PagingGridView.DataKeyNames = New String() {"ID"}

                                            If dt.Rows.Count > 0 Then
                                                ''                                              Stop
                                                ctlForm_PagingGridView.IsEmpty = False
                                                ctlForm_PagingGridView.DataSource = dt
                                                ctlForm_PagingGridView.DataBind()
                                            Else
                                                ''                                                Stop
                                                ctlForm_PagingGridView.IsEmpty = True
                                                ShowNoResultFound(dt, ctlForm_PagingGridView)
                                            End If

                                            'drGrid.Close()
                                            'drGrid = Nothing

                                            fRecordOK = CBool(cmdGrid.Parameters("@pfOK").Value)
                                            If Not fRecordOK Then
                                                sMessage = "Error loading web form. Web Form record selector item record has been deleted or not selected."
                                                Exit While
                                            End If

                                            cmdGrid.Dispose()
                                            cmdGrid = Nothing

                                            ' '' Add the newly created column header row to the column header table

                                            ''rowHeaderRow.BackColor = objGeneral.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                                            ''rowHeaderRow.BorderColor = objGeneral.GetColour(10720408)
                                            ''rowHeaderRow.BorderStyle = BorderStyle.Solid
                                            ''rowHeaderRow.BorderWidth = Unit.Pixel(0)
                                            ''rowHeaderRow.Font.Name = NullSafeString(dr("HeadFontName"))
                                            ''rowHeaderRow.Font.Size = FontUnit.Parse(NullSafeString(dr("HeadFontSize")))
                                            ''rowHeaderRow.Font.Bold = NullSafeBoolean(dr("HeadFontBold"))
                                            ''rowHeaderRow.Font.Italic = NullSafeBoolean(dr("HeadFontItalic"))
                                            ''rowHeaderRow.Font.Strikeout = NullSafeBoolean(dr("HeadFontStrikeThru"))
                                            ''rowHeaderRow.Font.Underline = NullSafeBoolean(dr("HeadFontUnderline"))
                                            ''rowHeaderRow.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                            ' '' rowHeaderRow.Wrap = False
                                            ''rowHeaderRow.Height = Unit.Pixel(iHeaderHeight)
                                            ''rowHeaderRow.VerticalAlign = VerticalAlign.Middle
                                            ''rowHeaderRow.HorizontalAlign = HorizontalAlign.Center

                                            ''ctlForm_Table.Rows.Add(rowHeaderRow)

                                            ''If Not (NullSafeBoolean(dr("ColumnHeaders")) And (NullSafeInteger(dr("Headlines")) > 0)) Then
                                            ''    ctlForm_Table.Style.Add("display", "none")
                                            ''End If

                                            '' ''NPG20100611 Fault HRPRO 873
                                            '' ''Deduct 1 pixel from the iGridWidth variable to prevent horizontal scrollbar.
                                            ' ''iGridWidth = NullSafeInteger(dr("Width")) - 1

                                            ' ''Adjust available width for the vertical scrollbar.
                                            ''iGapBetweenBorderAndText = (CInt(NullSafeSingle(dr("FontSize")) + 6) \ 4)
                                            ''iEffectiveRowHeight = CInt(NullSafeSingle(dr("FontSize"))) _
                                            '' + 1 _
                                            '' + (2 * iGapBetweenBorderAndText)

                                            ''If (ctlForm_RecordSelectionGrid.Rows.Count * iEffectiveRowHeight > iTempHeight) Then
                                            ''    iGridWidth = iGridWidth - 16
                                            ''    ctlForm_Table.Width = iGridWidth
                                            ''    ctlForm_RecordSelectionGrid.Width = iGridWidth
                                            ''    ctlForm_RecordSelectionGrid.Attributes.CssStyle("WIDTH") = iGridWidth.ToString
                                            ''End If

                                            ' ''NPG20100611 Fault HRPRO 873
                                            ' ''If the grid column widths are left to default it mucks up, so calculate every time.
                                            ' ''   I've also excluded the iGRIDBORDERWIDTH value when calculating the widths as this left a big gap.

                                            '' ''If iGridWidth > (iVisibleColumnCount * .DisplayLayout.ColWidthDefault.Value) _
                                            '' '' And (iVisibleColumnCount > 0) Then

                                            '' ''    iLoop = 0
                                            '' ''    iWidthUsed = 0
                                            '' ''    iGridWidth = iGridWidth - 2

                                            '' ''    For Each objGridColumn In .Columns
                                            '' ''        If objGridColumn.Hidden Then
                                            '' ''            objGridColumn.Width = Unit.Pixel(0)
                                            '' ''        Else
                                            '' ''            iLoop = iLoop + 1
                                            '' ''            If iLoop < iVisibleColumnCount Then
                                            '' ''                objGridColumn.Width = Unit.Pixel(CInt(iGridWidth / iVisibleColumnCount) - iGRIDBORDERWIDTH)

                                            '' ''                Dim iminWidth As Int32 = CInt(iGridWidth / iVisibleColumnCount)

                                            '' ''                If iminWidth < 100 Then iminWidth = 100

                                            '' ''                objGridColumn.Width = Unit.Pixel(iminWidth)

                                            '' ''                iWidthUsed = iWidthUsed + CInt(objGridColumn.Width.Value) + iGRIDBORDERWIDTH
                                            '' ''                iWidthUsed = iWidthUsed + CInt(objGridColumn.Width.Value)
                                            '' ''            Else
                                            '' ''                objGridColumn.Width = Unit.Pixel(iGridWidth - iWidthUsed - iGRIDBORDERWIDTH)
                                            '' ''                objGridColumn.Width = Unit.Pixel(iGridWidth - iWidthUsed)
                                            '' ''            End If
                                            '' ''        End If
                                            '' ''    Next objGridColumn
                                            '' ''    objGridColumn = Nothing
                                            '' ''End If

                                            ' '' Select the first row (if available).
                                            ' ''If ctlForm_RecordSelectionGrid.Rows.Count > 0 Then
                                            ' ''    If CStr(dr("value")).Length > 0 Then

                                            ' ''        'objGridCell = .Columns(iIDColumnIndex).Find(NullSafeString(dr("value")))

                                            ' ''        'If Not objGridCell Is Nothing Then
                                            ' ''        '    .Rows(objGridCell.Row.Index).Selected = True
                                            ' ''        '    .Rows(objGridCell.Row.Index).Activated = True
                                            ' ''        'Else
                                            ' ''        '    .Rows(0).Selected = True
                                            ' ''        '    .Rows(0).Activated = True
                                            ' ''        'End If
                                            ' ''    Else
                                            ' ''        '.Rows(0).Selected = True
                                            ' ''        '.Rows(0).Activated = True
                                            ' ''    End If
                                            ' ''End If

                                        Catch ex As Exception
                                            sMessage = "Error loading web form grid values:<BR><BR>" & _
                                             ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                                             "Contact your system administrator."
                                            Exit While

                                        Finally
                                            connGrid.Close()
                                            connGrid.Dispose()
                                        End Try
                                    Else
                                        ' If not a postback, check for empty datagrid and set empty row message
                                        Dim dtSource As DataTable = TryCast(HttpContext.Current.Session(sID & "DATA"), DataTable)

                                        If dtSource.Rows.Count = 1 Then
                                            ShowNoResultFound(dtSource, ctlForm_PagingGridView)
                                        End If
                                    End If

                                    ' ============================================================
                                    ' Hidden field is used to store scroll position of the grid.
                                    ' ============================================================
                                    ctlForm_HiddenField = New System.Web.UI.WebControls.HiddenField
                                    With ctlForm_HiddenField
                                        .ID = sID & "scrollpos"
                                    End With

                                    pnlInputDiv.Controls.Add(ctlForm_HiddenField)


                                Case 14 ' lookup  Inputs
                                    ' ctlForm_Dropdown = New Infragistics.WebUI.WebCombo.WebCombo()


                                    ' ==============================================================
                                    ' Lookups. So, these are created on the fly as:


                                    ctlForm_UpdatePanel = New System.Web.UI.UpdatePanel

                                    ' iGridWidth = NullSafeInteger(800)
                                    'Dim iGridHeight As Integer = NullSafeInteger(300)
                                    'Dim iColWidth As Integer = 100


                                    ' ============================================================
                                    ' Create a dropdown list as the main control
                                    ' ============================================================
                                    ctlForm_Dropdown = New DropDownList

                                    With ctlForm_Dropdown
                                        .ID = sID & "TextBox"
                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString
                                        .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                                        .Attributes.CssStyle("HEIGHT") = Unit.Pixel(NullSafeInteger(dr("Height"))).ToString
                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        .BorderColor = objGeneral.GetColour(5730458)
                                        .BorderStyle = BorderStyle.Solid
                                        .BorderWidth = Unit.Pixel(1)
                                        ' create a blank entry and select it. This prevents the dropdown
                                        ' becoming long by default. Need to hide it completely really, but
                                        ' under the kosh at the mo.
                                        .Items.Add("")
                                        .SelectedIndex = 0
                                    End With

                                    pnlInputDiv.Controls.Add(ctlForm_Dropdown)


                                    ' ============================================================
                                    ' Create the Lookup table grid, as per normal record selectors.
                                    ' This will be hidden on page_load, and displayed when the 
                                    ' DropdownList above is clicked. The magic is brought together
                                    ' by the AJAX DropDownExtender control below.
                                    ' ============================================================
                                    ctlForm_PagingGridView = New RecordSelector

                                    With ctlForm_PagingGridView
                                        .ID = sID & "Grid"
                                        .IsLookup = True
                                        .EnableViewState = True ' Must be set to True
                                        .IsEmpty = False
                                        .EmptyDataText = "no records to display"
                                        .AllowPaging = True
                                        .AllowSorting = True
                                        '.EnableSortingAndPagingCallbacks = True
                                        .PageSize = mobjConfig.LookupRowsRange
                                        .ShowFooter = False

                                        .Style.Add("Position", "Absolute")
                                        .Attributes.CssStyle("LEFT") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                                        .Attributes.CssStyle("TOP") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString

                                        ' Don't set the height of this control. Must use the ControlHeight property
                                        ' to stop the grid's rows from autosizing.
                                        ' .Height = Unit.Pixel(NullSafeInteger(dr("Height")))
                                        ' .Attributes.Add("Height", Unit.Pixel(NullSafeInteger(dr("Height"))).ToString)
                                        ' .Style.Add("height", Unit.Pixel(NullSafeInteger(dr("Height"))).ToString)
                                        .ControlHeight = NullSafeInteger(dr("Height"))
                                        .Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        ' Header Row - fixed for lookups.
                                        .ColumnHeaders = True
                                        .HeadFontSize = NullSafeSingle(dr("FontSize"))
                                        .HeadLines = 1

                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        ''.SelForeColor = System.Drawing.SystemColors.HighlightText
                                        ''.SelBackColor = System.Drawing.SystemColors.Highlight

                                        msBackColorHighlight = System.Drawing.SystemColors.Highlight.ToString

                                        .BorderColor = objGeneral.GetColour(5730458)
                                        .BorderStyle = BorderStyle.Solid
                                        .BorderWidth = Unit.Pixel(1)

                                        ''.DropDownLayout.FrameStyle.BorderColor = objGeneral.GetColour(10720408)
                                        ''.DropDownLayout.FrameStyle.BorderStyle = BorderStyle.Solid
                                        ''.DropDownLayout.FrameStyle.BorderWidth = Unit.Pixel(1)
                                        ''.DropDownLayout.FrameStyle.BackColor = objGeneral.GetColour(16248040)

                                        .RowStyle.Font.Name = NullSafeString(dr("FontName"))
                                        .RowStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .RowStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .RowStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .RowStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .RowStyle.BackColor = objGeneral.GetColour(15988214)
                                        .RowStyle.ForeColor = objGeneral.GetColour(6697779)

                                        iRowHeight = 21
                                        '.DisplayLayout.RowHeightDefault = Unit.Pixel(iRowHeight)

                                        ' NPG20110802 Fault HRPRO-1685, remove rowheight to autofit.
                                        ' .RowStyle.Height = Unit.Pixel(iRowHeight)

                                        ''.ExpandEffects.Type = Infragistics.WebUI.WebCombo.ExpandEffectType.Slide

                                        ''.DropDownLayout.RowStyle.Padding.Left = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Right = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Top = Unit.Pixel(0)
                                        ''.DropDownLayout.RowStyle.Padding.Bottom = Unit.Pixel(1)
                                        .RowStyle.VerticalAlign = VerticalAlign.Middle

                                        .SelectedRowStyle.ForeColor = objGeneral.GetColour(2774907)
                                        .SelectedRowStyle.BackColor = objGeneral.GetColour(10480637)

                                        ''.DropDownLayout.BorderCollapse = Infragistics.WebUI.UltraWebGrid.BorderCollapse.Collapse

                                        ''.DropDownLayout.TableLayout = Infragistics.WebUI.UltraWebGrid.TableLayout.Fixed

                                        ''.DropDownLayout.AllowSorting = Infragistics.WebUI.UltraWebGrid.AllowSorting.Yes
                                        ''.DropDownLayout.HeaderClickAction = Infragistics.WebUI.UltraWebGrid.HeaderClickAction.SortMulti
                                        ''.DropDownLayout.StationaryMargins = Infragistics.WebUI.UltraWebGrid.StationaryMargins.Header
                                        ''.DropDownLayout.RowStyle.Cursor = Infragistics.WebUI.Shared.Cursors.Default

                                        ''.DropDownLayout.AllowColSizing = Infragistics.WebUI.UltraWebGrid.AllowSizing.Free
                                        ''.DropDownLayout.GridLines = Infragistics.WebUI.UltraWebGrid.UltraGridLines.Both

                                        ' HEADER formatting
                                        'iGridTopPadding = CInt(NullSafeSingle(dr("FontSize")) / 8)
                                        'iHeaderHeight = CInt(((NullSafeSingle(dr("FontSize")) + iGridTopPadding) * 2) _
                                        ' - 2 _
                                        ' - (NullSafeSingle(dr("FontSize")) * 2 * (iGridTopPadding - 1) / 4))

                                        .HeaderStyle.BackColor = objGeneral.GetColour(16248553)
                                        .HeaderStyle.BorderColor = objGeneral.GetColour(10720408)
                                        .HeaderStyle.BorderStyle = BorderStyle.Solid
                                        .HeaderStyle.BorderWidth = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthTop = Unit.Pixel(0)
                                        '.HeaderStyle.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        '.HeaderStyle.BorderDetails.WidthRight = Unit.Pixel(1)

                                        .HeaderStyle.Font.Name = NullSafeString(dr("FontName"))
                                        .HeaderStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .HeaderStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .HeaderStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .HeaderStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .HeaderStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        '.HeaderStyle.Padding.Top = Unit.Pixel(iGridTopPadding)
                                        '.HeaderStyle.Padding.Bottom = Unit.Pixel(0)
                                        '.HeaderStyle.Padding.Left = Unit.Pixel(2)
                                        '.HeaderStyle.Padding.Right = Unit.Pixel(2)
                                        .HeaderStyle.Wrap = False
                                        .HeaderStyle.Height = Unit.Pixel(iHeaderHeight)
                                        .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                                        .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                                        ' ROW formatting
                                        '.RowStyle.Padding.Left = Unit.Pixel(3)
                                        '.RowStyle.Padding.Left = Unit.Pixel(3)
                                        '.RowStyle.Padding.Right = Unit.Pixel(3)
                                        '.RowStyle.Padding.Top = Unit.Pixel(0)
                                        '.RowStyle.Padding.Bottom = Unit.Pixel(1)
                                        .RowStyle.VerticalAlign = VerticalAlign.Middle

                                        '.RowSelectors = Infragistics.WebUI.UltraWebGrid.RowSelectors.No
                                        '.RowsRange = mobjConfig.LookupRowsRange

                                        .PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
                                        .PagerSettings.FirstPageImageUrl = "Images/page-first.gif"
                                        .PagerSettings.LastPageImageUrl = "Images/page-last.gif"
                                        .PagerSettings.NextPageImageUrl = "Images/page-next.gif"
                                        .PagerSettings.PreviousPageImageUrl = "Images/page-prev.gif"

                                        .PagerStyle.BorderWidth = Unit.Pixel(0)
                                        .PagerStyle.Font.Name = NullSafeString(dr("FontName"))
                                        .PagerStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .PagerStyle.Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .PagerStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .PagerStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .PagerStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        .PagerStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        .PagerStyle.Wrap = False
                                        .PagerStyle.VerticalAlign = VerticalAlign.Middle
                                        .PagerStyle.HorizontalAlign = HorizontalAlign.Center
                                    End With

                                    sFilterSQL = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")), _
                                            NullSafeInteger(dr("lookupFilterColumnDataType")), _
                                            NullSafeInteger(dr("LookupFilterOperator")), _
                                            FORMINPUTPREFIX & NullSafeString(dr("lookupFilterValueID")) & "_" & NullSafeString(dr("lookupFilterValueType")) & "_")


                                    ' ==========================================================
                                    ' Hidden Field to store any lookup filter code
                                    ' ==========================================================
                                    If (sFilterSQL.Length > 0) Then
                                        ctlForm_HiddenField = New HiddenField
                                        With ctlForm_HiddenField
                                            .ID = "lookup" & sID
                                            .Value = sFilterSQL
                                        End With
                                        pnlInputDiv.Controls.Add(ctlForm_HiddenField)
                                    End If


                                    ' The following client side handlers are now part of the web page itself, beginRequest and endRequest handlers
                                    ' InitializeDataSource & DataBound methods need to be defined because we have EnableXmlHTTP enabled.
                                    ' EnableXmlHTTP is required becaue we're using client-side filtering using 'selectwhere'
                                    'AddHandler ctlForm_Dropdown.DataBound, AddressOf Me.LookupDataBound
                                    'AddHandler ctlForm_Dropdown.Load, AddressOf Me.InitializeLookupData

                                    ' The following is set on the DropDownExtender Control below.
                                    'AddHandler ctlForm_Dropdown.InitializeDataSource, AddressOf Me.InitializeLookupData

                                    pnlInputDiv.Controls.Add(ctlForm_PagingGridView)


                                    If (Not IsPostBack) Then
                                        connGrid = New SqlClient.SqlConnection(strConn)
                                        connGrid.Open()
                                        Try
                                            cmdGrid = New SqlClient.SqlCommand
                                            cmdGrid.CommandText = "spASRGetWorkflowItemValues"
                                            cmdGrid.Connection = connGrid
                                            cmdGrid.CommandType = CommandType.StoredProcedure
                                            cmdGrid.CommandTimeout = miSubmissionTimeoutInSeconds

                                            cmdGrid.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
                                            cmdGrid.Parameters("@piElementItemID").Value = CInt(NullSafeString(dr("id")))

                                            cmdGrid.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                            cmdGrid.Parameters("@piInstanceID").Value = miInstanceID

                                            cmdGrid.Parameters.Add("@piLookupColumnIndex", SqlDbType.Int).Direction = ParameterDirection.Output
                                            cmdGrid.Parameters.Add("@piItemType", SqlDbType.Int).Direction = ParameterDirection.Output
                                            cmdGrid.Parameters.Add("@psDefaultValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                            da = New SqlDataAdapter(cmdGrid)
                                            dt = New DataTable()

                                            ' Fill the datatable with data from the datadapter.
                                            da.Fill(dt)

                                            Session(sID & "DATA") = dt

                                            '' Create a blank row at the top of the dropdown grid.
                                            objDataRow = dt.NewRow()
                                            dt.Rows.InsertAt(objDataRow, 0)

                                            'drGrid = cmdGrid.ExecuteReader()

                                            'drGrid.Close()
                                            'drGrid = Nothing

                                            m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)

                                            iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                                            ctlForm_PagingGridView.Attributes.Remove("LookupColumnIndex")
                                            ctlForm_PagingGridView.Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                                            ctlForm_PagingGridView.Attributes.Remove("DefaultValue")
                                            ctlForm_PagingGridView.Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                                            ' Datatype - 
                                            ' curSelDataType = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

                                            'dt.Columns(CInt(ctlForm_Dropdown.Attributes("LookupColumnIndex"))).DataType
                                            ctlForm_Dropdown.Attributes.Remove("DataType")
                                            ctlForm_Dropdown.Attributes.Add("DataType", NullSafeString(dt.Columns(CInt(ctlForm_Dropdown.Attributes("LookupColumnIndex"))).DataType.ToString))

                                            ' Yup we store the data to a session variable. This is so we can sort/filter 
                                            ' it and stillreset if necessary without running the SP again
                                            ctlForm_PagingGridView.DataSource = Session(sID & "DATA")
                                            ctlForm_PagingGridView.DataBind()

                                            cmdGrid.Dispose()
                                            cmdGrid = Nothing

                                            '' Add the newly created column header row to the column header table

                                            'rowHeaderRow.BackColor = objGeneral.GetColour(16248553)
                                            'rowHeaderRow.BorderColor = objGeneral.GetColour(10720408)
                                            'rowHeaderRow.BorderStyle = BorderStyle.Solid
                                            'rowHeaderRow.BorderWidth = Unit.Pixel(0)
                                            'rowHeaderRow.Font.Name = NullSafeString(dr("FontName"))
                                            'rowHeaderRow.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                            'rowHeaderRow.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                            'rowHeaderRow.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                            'rowHeaderRow.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                            'rowHeaderRow.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                            ''rowHeaderRow.Wrap = False
                                            'rowHeaderRow.Height = Unit.Pixel(iHeaderHeight)
                                            'rowHeaderRow.VerticalAlign = VerticalAlign.Middle
                                            'rowHeaderRow.HorizontalAlign = HorizontalAlign.Center

                                            'ctlForm_Table.Rows.Add(rowHeaderRow)

                                            ' ''NPG20100611 Fault HRPRO 873
                                            ' ''Deduct 1 pixel from the iGridWidth variable to prevent horizontal scrollbar.
                                            ''iGridWidth = NullSafeInteger(dr("Width")) - 1

                                            ''Adjust available width for the vertical scrollbar.
                                            'iGapBetweenBorderAndText = (CInt(NullSafeSingle(dr("FontSize")) + 6) \ 4)
                                            'iEffectiveRowHeight = CInt(NullSafeSingle(dr("FontSize"))) _
                                            ' + 1 _
                                            ' + (2 * iGapBetweenBorderAndText)

                                            'iTempHeight = NullSafeInteger(ctlForm_GridContainer.Height.Value)
                                            'iTempHeight = CInt(IIf(iTempHeight < 0, 1, iTempHeight))

                                            'iGridWidth = iVisibleColumnCount * iColWidth
                                            'iGridWidth = CInt(IIf(iGridWidth < 0, 1, iGridWidth))
                                            '' minimum width should be the width of the dropdown control itself
                                            'iGridWidth = CInt(IIf(iGridWidth < NullSafeInteger(dr("Width")), NullSafeInteger(dr("Width")), iGridWidth))

                                            '' maximum width???

                                            'If (ctlForm_RecordSelectionGrid.Rows.Count * iEffectiveRowHeight > iTempHeight) Then
                                            '    ' There's a scrollbar
                                            '    ctlForm_GridContainer.Width = iGridWidth - 16
                                            '    iGridWidth = iGridWidth - 32
                                            '    ctlForm_Table.Width = iGridWidth
                                            '    ctlForm_RecordSelectionGrid.Width = iGridWidth
                                            '    ctlForm_RecordSelectionGrid.Attributes.CssStyle("WIDTH") = iGridWidth.ToString
                                            'Else
                                            '    ctlForm_GridContainer.Width = iGridWidth
                                            '    iGridWidth = iGridWidth - 1
                                            '    ctlForm_Table.Width = iGridWidth
                                            '    ctlForm_RecordSelectionGrid.Width = iGridWidth
                                            '    ctlForm_RecordSelectionGrid.Attributes.CssStyle("WIDTH") = iGridWidth.ToString
                                            'End If

                                            '' Set the size of the grid as per old DropDown setting...
                                            '' iRowHeight = CInt(ctlForm_TextInput.Height.Value) - 6
                                            'iRowHeight = CInt(ctlForm_Dropdown.Height.Value) - 6
                                            'iRowHeight = CInt(IIf(iRowHeight < 22, 22, iRowHeight))
                                            'iDropHeight = (iRowHeight * CInt(IIf(ctlForm_RecordSelectionGrid.Rows.Count > MAXDROPDOWNROWS, MAXDROPDOWNROWS, ctlForm_RecordSelectionGrid.Rows.Count))) + 1

                                        Catch ex As Exception

                                            sMessage = "Error loading lookup values:<BR><BR>" & _
                                             ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                                             "Contact your system administrator."
                                            Exit While

                                        Finally
                                            connGrid.Close()
                                            connGrid.Dispose()
                                        End Try

                                        ' ==================================================
                                        ' Set the dropdownList to the default value.
                                        ' ==================================================
                                        If ctlForm_PagingGridView.Attributes("DefaultValue").ToString.Length > 0 Then
                                            ctlForm_Dropdown.Items.Add(ctlForm_PagingGridView.Attributes("DefaultValue").ToString)
                                            ctlForm_Dropdown.SelectedIndex = 1
                                        End If

                                        For jncount As Integer = 0 To ctlForm_PagingGridView.Rows.Count - 1
                                            If jncount > ctlForm_PagingGridView.PageSize Then Exit For ' don't bother if on other pages
                                            If ctlForm_PagingGridView.Rows(jncount).Cells(m_iLookupColumnIndex).Text = ctlForm_PagingGridView.Attributes("DefaultValue").ToString Then
                                                ctlForm_PagingGridView.SelectedIndex = jncount
                                                Exit For
                                            End If

                                        Next
                                    End If




                                    ' =============================================================================
                                    ' AJAX DropDownExtender (DDE) Control
                                    ' This simply links up the DropDownList and the Lookup Grid to make a dropdown.
                                    ' =============================================================================
                                    ctlForm_DDE = New AjaxControlToolkit.DropDownExtender

                                    With ctlForm_DDE
                                        .DropArrowBackColor = Color.Transparent
                                        .DropArrowWidth = Unit.Pixel(20)
                                        .HighlightBackColor = Color.Transparent
                                        ' Careful with the case here, use 'dde' in JavaScript:
                                        .ID = sID & "DDE"
                                        .BehaviorID = sID & "dde"
                                        .DropDownControlID = sID
                                        .Enabled = True
                                        .TargetControlID = sID & "TextBox"
                                        ' Client-side handler.
                                        .OnClientPopup = "InitializeLookup"     ' can't pass the ID of the control, so use ._id in JS.
                                    End With

                                    pnlInputDiv.Controls.Add(ctlForm_DDE)

                                    ' =================================================================
                                    ' Attach a JavaScript functino to the 'add_shown' method of this
                                    ' DropDownExtender. Used to check if popup is bigger than the
                                    ' parent form, and resize the parent form if necessary
                                    ' =================================================================
                                    scriptString += "var bhvDdl=$find('" & ctlForm_DDE.BehaviorID.ToString & "');"
                                    scriptString += "try {bhvDdl .add_shown(ResizeComboForForm);} catch (e) {}"


                                    ' ====================================================
                                    ' hidden field to store scroll position (not required?)
                                    ' ====================================================
                                    ctlForm_HiddenField = New System.Web.UI.WebControls.HiddenField
                                    With ctlForm_HiddenField
                                        .ID = sID & "scrollpos"
                                    End With

                                    pnlInputDiv.Controls.Add(ctlForm_HiddenField)

                                    ' ====================================================
                                    ' hidden field to hold any filter SQL code
                                    ' ====================================================
                                    ctlForm_HiddenField = New System.Web.UI.WebControls.HiddenField
                                    With ctlForm_HiddenField
                                        .ID = sID & "filterSQL"
                                    End With

                                    pnlInputDiv.Controls.Add(ctlForm_HiddenField)

                                    ' ============================================================
                                    ' Hidden Button for JS to call which fires filter click event. 
                                    ' ============================================================
                                    ctlForm_InputButton = New Button

                                    With ctlForm_InputButton
                                        .ID = sID & "refresh"
                                        .Style.Add("display", "none")
                                    End With

                                    AddHandler ctlForm_InputButton.Click, AddressOf SetLookupFilter

                                    pnlInputDiv.Controls.Add(ctlForm_InputButton)






                                Case 13 ' Dropdown (13) Inputs

                                    ctlForm_Dropdown = New System.Web.UI.WebControls.DropDownList ' Infragistics.WebUI.WebCombo.WebCombo()

                                    With ctlForm_Dropdown
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        ' EnableXmlHTTP is required becaue we're using client-side filtering using 'selectwhere'
                                        ''.EnableXmlHTTP = True
                                        ''.ClientSideEvents.BeforeDropDown = "InitializeLookup"
                                        ''.ClientSideEvents.EditKeyDown = "dropdownControlKeyPress"
                                        ''.ClientSideEvents.AfterDropDown = "ResizeFormForCombo"

                                        '.DropDownLayout.XmlLoadOnDemandType = Infragistics.WebUI.WebCombo.ComboLoadOnDemandType.Accumulative

                                        'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Dropdown)
                                        pnlInputDiv.Controls.Add(ctlForm_Dropdown)

                                        sFilterSQL = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")), _
                                                NullSafeInteger(dr("lookupFilterColumnDataType")), _
                                                NullSafeInteger(dr("LookupFilterOperator")), _
                                                FORMINPUTPREFIX & NullSafeString(dr("lookupFilterValueID")) & "_" & NullSafeString(dr("lookupFilterValueType")) & "_")

                                        If (sFilterSQL.Length > 0) Then
                                            ctlForm_HiddenField = New HiddenField
                                            With ctlForm_HiddenField
                                                .ID = "lookup" & sID
                                                .Value = sFilterSQL
                                            End With
                                            'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_HiddenField)
                                            pnlInputDiv.Controls.Add(ctlForm_HiddenField)
                                        End If

                                        ' InitializeDataSource & DataBound methods need to be defined because we have EnableXmlHTTP enabled.
                                        ' EnableXmlHTTP is required becaue we're using client-side filtering using 'selectwhere'
                                        'AddHandler ctlForm_Dropdown.InitializeDataSource, AddressOf Me.InitializeLookupData

                                        'AddHandler ctlForm_Dropdown.Load, AddressOf Me.InitializeLookupData

                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(dr("FontBold"))
                                        .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                                        .BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        ''.SelForeColor = System.Drawing.SystemColors.HighlightText
                                        ''.SelBackColor = System.Drawing.SystemColors.Highlight
                                        ' .BorderColor = objGeneral.GetColour(5730458)

                                        ''.DropDownLayout.FrameStyle.BorderColor = objGeneral.GetColour(10720408)
                                        ''.DropDownLayout.FrameStyle.BorderStyle = BorderStyle.Solid
                                        ''.DropDownLayout.FrameStyle.BorderWidth = Unit.Pixel(1)
                                        ''.DropDownLayout.FrameStyle.BackColor = objGeneral.GetColour(16248040)

                                        ''.DropDownLayout.RowStyle.Font.Name = NullSafeString(dr("FontName"))
                                        ''.DropDownLayout.RowStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        ''.DropDownLayout.RowStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        ''.DropDownLayout.RowStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        ''.DropDownLayout.RowStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        ''.DropDownLayout.RowStyle.BackColor = objGeneral.GetColour(15988214)
                                        ''.DropDownLayout.RowStyle.ForeColor = objGeneral.GetColour(6697779)

                                        ''.ExpandEffects.Type = Infragistics.WebUI.WebCombo.ExpandEffectType.Slide

                                        ''.DropDownLayout.RowStyle.Padding.Left = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Right = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Top = Unit.Pixel(0)
                                        ''.DropDownLayout.RowStyle.Padding.Bottom = Unit.Pixel(1)
                                        ''.DropDownLayout.RowStyle.VerticalAlign = VerticalAlign.Middle

                                        ''.DropDownLayout.SelectedRowStyle.ForeColor = objGeneral.GetColour(2774907)
                                        ''.DropDownLayout.SelectedRowStyle.BackColor = objGeneral.GetColour(10480637)

                                        ''.DropDownLayout.BorderCollapse = Infragistics.WebUI.UltraWebGrid.BorderCollapse.Collapse

                                        ''.DropDownLayout.TableLayout = Infragistics.WebUI.UltraWebGrid.TableLayout.Fixed

                                        .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                                        .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 2)

                                        ''.DropDownLayout.AllowSorting = Infragistics.WebUI.UltraWebGrid.AllowSorting.Yes
                                        ''.DropDownLayout.HeaderClickAction = Infragistics.WebUI.UltraWebGrid.HeaderClickAction.SortMulti
                                        ''.DropDownLayout.StationaryMargins = Infragistics.WebUI.UltraWebGrid.StationaryMargins.Header
                                        ''.DropDownLayout.RowStyle.Cursor = Infragistics.WebUI.Shared.Cursors.Default

                                        ''.DropDownLayout.AllowColSizing = Infragistics.WebUI.UltraWebGrid.AllowSizing.Free

                                        ''.DropDownLayout.GridLines = Infragistics.WebUI.UltraWebGrid.UltraGridLines.Both

                                        ' HEADER formatting
                                        iGridTopPadding = CInt(NullSafeSingle(dr("FontSize")) / 8)
                                        iHeaderHeight = CInt(((NullSafeSingle(dr("FontSize")) + iGridTopPadding) * 2) _
                                         - 2 _
                                         - (NullSafeSingle(dr("FontSize")) * 2 * (iGridTopPadding - 1) / 4))

                                        ''.DropDownLayout.HeaderStyle.BackColor = objGeneral.GetColour(16248553)
                                        ''.DropDownLayout.HeaderStyle.BorderColor = objGeneral.GetColour(10720408)
                                        ''.DropDownLayout.HeaderStyle.BorderStyle = BorderStyle.Solid
                                        ''.DropDownLayout.HeaderStyle.BorderDetails.WidthLeft = Unit.Pixel(0)
                                        ''.DropDownLayout.HeaderStyle.BorderDetails.WidthTop = Unit.Pixel(0)
                                        ''.DropDownLayout.HeaderStyle.BorderDetails.WidthBottom = Unit.Pixel(1)
                                        ''.DropDownLayout.HeaderStyle.BorderDetails.WidthRight = Unit.Pixel(1)

                                        ''.DropDownLayout.HeaderStyle.Font.Name = NullSafeString(dr("FontName"))
                                        ''.DropDownLayout.HeaderStyle.Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        ''.DropDownLayout.HeaderStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                                        ''.DropDownLayout.HeaderStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                                        ''.DropDownLayout.HeaderStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                                        ''.DropDownLayout.HeaderStyle.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        ''.DropDownLayout.HeaderStyle.Padding.Top = Unit.Pixel(iGridTopPadding)
                                        ''.DropDownLayout.HeaderStyle.Padding.Bottom = Unit.Pixel(0)
                                        ''.DropDownLayout.HeaderStyle.Padding.Left = Unit.Pixel(2)
                                        ''.DropDownLayout.HeaderStyle.Padding.Right = Unit.Pixel(2)
                                        ''.DropDownLayout.HeaderStyle.Wrap = False
                                        ''.DropDownLayout.HeaderStyle.Height = Unit.Pixel(iHeaderHeight)
                                        ''.DropDownLayout.HeaderStyle.VerticalAlign = VerticalAlign.Middle
                                        ''.DropDownLayout.HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                                        ' ROW formatting
                                        ''.DropDownLayout.RowStyle.Padding.Left = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Left = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Right = Unit.Pixel(3)
                                        ''.DropDownLayout.RowStyle.Padding.Top = Unit.Pixel(0)
                                        ''.DropDownLayout.RowStyle.Padding.Bottom = Unit.Pixel(1)
                                        ''.DropDownLayout.RowStyle.VerticalAlign = VerticalAlign.Middle

                                        ''.DropDownLayout.RowSelectors = Infragistics.WebUI.UltraWebGrid.RowSelectors.No
                                        ''.DropDownLayout.RowsRange = mobjConfig.LookupRowsRange


                                        If (Not IsPostBack) Then
                                            connGrid = New SqlClient.SqlConnection(strConn)
                                            connGrid.Open()

                                            Try

                                                cmdGrid = New SqlClient.SqlCommand
                                                cmdGrid.CommandText = "spASRGetWorkflowItemValues"
                                                cmdGrid.Connection = connGrid
                                                cmdGrid.CommandType = CommandType.StoredProcedure
                                                cmdGrid.CommandTimeout = miSubmissionTimeoutInSeconds

                                                cmdGrid.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
                                                cmdGrid.Parameters("@piElementItemID").Value = CInt(NullSafeString(dr("id")))

                                                cmdGrid.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                                cmdGrid.Parameters("@piInstanceID").Value = miInstanceID

                                                cmdGrid.Parameters.Add("@piLookupColumnIndex", SqlDbType.Int).Direction = ParameterDirection.Output
                                                cmdGrid.Parameters.Add("@piItemType", SqlDbType.Int).Direction = ParameterDirection.Output
                                                cmdGrid.Parameters.Add("@psDefaultValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                                da = New SqlDataAdapter(cmdGrid)
                                                dt = New DataTable()

                                                ' Create a blank row at the top of the dropdown grid.
                                                objDataRow = dt.NewRow()
                                                dt.Rows.InsertAt(objDataRow, 0)

                                                ' Fill the datatable with data from the datadapter.
                                                da.Fill(dt)
                                                Session(sID & "_DATA") = dt

                                                '' Format the column(s)
                                                For Each objGridColumn In dt.Columns
                                                    If objGridColumn.ColumnName.StartsWith("ASRSys") Then
                                                        'objGridColumn..Hidden = True
                                                    Else
                                                        'If iLookupColumnIndex = objGridColumn.Index Then
                                                        .DataTextField = objGridColumn.ColumnName.ToString

                                                        'If objGridColumn.DataType = "System.Decimal" _
                                                        'Or objGridColumn.DataType = "System.Int32" Then

                                                        '    .ClientSideEvents.AfterSelectChange = "ChangeLookup"
                                                        'End If

                                                        '.Attributes.Remove("DataType")
                                                        '.Attributes.Add("DataType", objGridColumn.DataType)

                                                        'If iType = 13 Then
                                                        '    If .Rows.Count > MAXDROPDOWNROWS Then
                                                        '        ' Vertical scrollbar will be visible. Adjust the column width
                                                        '        .DropDownLayout.ColWidthDefault = Unit.Pixel(CInt(.DropDownLayout.DropdownWidth.Value) - 20)
                                                        '    Else
                                                        '        ' Vertical scrollbar will NOT be visible.
                                                        '        .DropDownLayout.ColWidthDefault = Unit.Pixel(CInt(.DropDownLayout.DropdownWidth.Value) - 5)
                                                        '    End If
                                                        'End If

                                                        'End If
                                                        'objGridColumn.AllowNull = False

                                                        'If objGridColumn.DataType = "System.DateTime" Then
                                                        '    objGridColumn.Format = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern

                                                        'ElseIf objGridColumn.DataType = "System.Boolean" Then
                                                        '    objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Center

                                                        'ElseIf objGridColumn.DataType = "System.Decimal" _
                                                        '    Or objGridColumn.DataType = "System.Int32" Then

                                                        '    objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Right
                                                        'End If
                                                    End If
                                                Next objGridColumn

                                                ctlForm_Dropdown.DataSource = dt

                                                m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)
                                                iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                                                .Attributes.Remove("LookupColumnIndex")
                                                .Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                                                .Attributes.Remove("DefaultValue")
                                                .Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                                                ' AddHandler ctlForm_Dropdown.DataBound, AddressOf Me.LookupDataBound

                                                ctlForm_Dropdown.DataBind()

                                                cmdGrid.Dispose()
                                                cmdGrid = Nothing



                                                ' Only show headers for lookups, not dropdown lists
                                                If iItemType = 14 Then
                                                    ''.DropDownLayout.ColHeadersVisible = Infragistics.WebUI.UltraWebGrid.ShowMarginInfo.Yes
                                                Else
                                                    ''.DropDownLayout.ColHeadersVisible = Infragistics.WebUI.UltraWebGrid.ShowMarginInfo.No
                                                End If

                                                iRowHeight = CInt(.Height.Value) - 6
                                                iRowHeight = CInt(IIf(iRowHeight < 22, 22, iRowHeight))
                                                iDropHeight = (iRowHeight * CInt(IIf(dt.Rows.Count > MAXDROPDOWNROWS, MAXDROPDOWNROWS, dt.Rows.Count))) + 1
                                                ''.DropDownLayout.DropdownHeight = Unit.Pixel(iDropHeight)

                                            Catch ex As Exception
                                                sMessage = "Error loading lookup values:<BR><BR>" & _
                                                 ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                                                 "Contact your system administrator."
                                                Exit While

                                            Finally
                                                connGrid.Close()
                                                connGrid.Dispose()
                                            End Try

                                        End If


                                        ' Set dropdown width to fit the columns displayed.
                                        ''If NullSafeInteger(dr("ItemType")) = 14 Then
                                        ''    .DropDownLayout.DropdownWidth = System.Web.UI.WebControls.Unit.Empty
                                        ''Else
                                        ''    .DropDownLayout.DropdownWidth = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        ''End If


                                    End With

                                Case 15 ' OptionGroup
                                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                                        sBackColour = "Transparent"
                                    Else
                                        sBackColour = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColor")))
                                    End If

                                    sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                                     CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                                    If sTemp2.Length = 0 Then
                                        sTemp2 = " none"
                                    End If

                                    sTemp3 = ""

                                    Dim fieldsetTop As Int32

                                    If Not NullSafeBoolean(dr("PictureBorder")) Then
                                        fieldsetTop = NullSafeInteger(dr("TopCoord"))

                                        sTemp3 = " BORDER-STYLE: none;"
                                        sTemp = "<fieldset style='z-index: 0; TOP: " & fieldsetTop.ToString & "px; " & _
                                         " LEFT: " & (NullSafeInteger(dr("LeftCoord")) - 1).ToString & "px; " & _
                                         " WIDTH: " & (NullSafeInteger(dr("Width")) - 1).ToString & "px; " & _
                                         " HEIGHT: " & (NullSafeInteger(dr("Height")) + 1).ToString & "px; " & _
                                         " BACKGROUND-COLOR: " & sBackColour & "; " & _
                                         " COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                                         " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                         " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                         " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                         " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                         " TEXT-DECORATION:" & sTemp2 & ";" & sTemp3 & _
                                         " POSITION: absolute;'>"

                                        iYOffset = CInt(NullSafeSingle(dr("FontSize")) / 2)
                                    Else
                                        fieldsetTop = _
                                     CInt((NullSafeInteger(dr("TopCoord")) + (NullSafeSingle(dr("FontSize")) * 2.5 / 3)))
                                        Dim fieldsetLeft As Int32 = NullSafeInteger(dr("LeftCoord"))
                                        Dim fieldsetWidth As Int32 = NullSafeInteger(dr("Width")) - 2
                                        Dim fieldsetHeight As Int32 = _
                                         CInt((NullSafeInteger(dr("Height")) - 1 - (NullSafeSingle(dr("FontSize")) * 2.5 / 3)))

                                        sTemp = "<fieldset style='z-index: 0; TOP: " & fieldsetTop.ToString & "px; " & _
                                         " LEFT: " & fieldsetLeft.ToString & "px; " & _
                                         " WIDTH: " & fieldsetWidth.ToString & "px; " & _
                                         " HEIGHT: " & fieldsetHeight.ToString & "px; " & _
                                         " BACKGROUND-COLOR: " & sBackColour & "; " & _
                                         " BORDER-STYLE: solid; " & _
                                         " BORDER-COLOR: #9894a3; " & _
                                         " BORDER-WIDTH: 1px; " & _
                                         " COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                                         " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                         " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                         " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                         " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                         " TEXT-DECORATION:" & sTemp2 & ";" & sTemp3 & _
                                         " POSITION: absolute;padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin: 0px; padding-top: 0px;'>"

                                        iYOffset = CInt(2 - NullSafeSingle(dr("FontSize")) - (2 * (NullSafeSingle(dr("FontSize")) / 4) - 2))
                                    End If

                                    If NullSafeBoolean(dr("PictureBorder")) And (NullSafeString(dr("caption")).Trim.Length > 0) Then
                                        Dim legendTop As Int32 = CInt((NullSafeSingle(dr("FontSize")) * -11 / 10))

                                        sTemp = sTemp & _
                                        "<legend" & _
                                        " style='top: " & legendTop.ToString & "px;" & _
                                        " COLOR: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                                        " padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin-left: 5px; padding-top: " & CInt(NullSafeSingle(dr("FontSize")) / 4).ToString & "px; position: relative;' align='Left'>" & _
                                        NullSafeString(dr("caption")) & _
                                        "</legend>"
                                    End If

                                    connGrid = New SqlClient.SqlConnection(strConn)
                                    connGrid.Open()

                                    Try
                                        cmdGrid = New SqlClient.SqlCommand
                                        cmdGrid.CommandText = "spASRGetWorkflowItemValues"
                                        cmdGrid.Connection = connGrid
                                        cmdGrid.CommandType = CommandType.StoredProcedure
                                        cmdGrid.CommandTimeout = miSubmissionTimeoutInSeconds

                                        cmdGrid.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
                                        cmdGrid.Parameters("@piElementItemID").Value = NullSafeString(dr("ID"))

                                        cmdGrid.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                        cmdGrid.Parameters("@piInstanceID").Value = miInstanceID

                                        cmdGrid.Parameters.Add("@piLookupColumnIndex", SqlDbType.Int).Direction = ParameterDirection.Output
                                        cmdGrid.Parameters.Add("@piItemType", SqlDbType.Int).Direction = ParameterDirection.Output
                                        cmdGrid.Parameters.Add("@psDefaultValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                        drGrid = cmdGrid.ExecuteReader

                                        Dim graphic As Graphics = Graphics.FromImage(New Bitmap(1, 1, Imaging.PixelFormat.Format32bppArgb))
                                        Dim style As FontStyle = _
                                         CType(IIf(NullSafeBoolean(dr("FontBold")), FontStyle.Bold, FontStyle.Regular), FontStyle) Or _
                                         CType(IIf(NullSafeBoolean(dr("FontItalic")), FontStyle.Italic, FontStyle.Regular), FontStyle)

                                        Dim font As Font = New Font(dr("FontName").ToString(), Convert.ToInt32(dr("FontSize")), style)

                                        Dim stringSize As SizeF = New SizeF
                                        Dim lastLeft As Double = CInt(((NullSafeSingle(dr("FontSize"))) * 5 / 4)) - 1
                                        Dim spacer As Single = graphic.MeasureString("WW", font).Width

                                        iTemp = 0
                                        sDefaultValue = ""
                                        While drGrid.Read
                                            Select Case NullSafeInteger(dr("Orientation"))
                                                Case 0 ' Vertical

                                                    Dim spanTop As Int32 = _
                                                        CInt((NullSafeInteger(dr("FontSize")) * 1.25) + 1) + _
                                                        CInt(iTemp * CInt((NullSafeInteger(dr("FontSize")) * 1.5) + 4)) - _
                                                        CInt(IIf(NullSafeBoolean(dr("PictureBorder")), 0, 10))

                                                    Dim spanLeft As Int32 = CInt(((NullSafeSingle(dr("FontSize"))) * 5 / 4) - 1)

                                                    sTemp = sTemp & _
                                                     "<span tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                                                     " style=""z-index: 0;" & _
                                                     " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                                     " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                                     " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                                     " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                                     " TEXT-DECORATION:" & sTemp2 & ";" & _
                                                     " left: " & spanLeft.ToString & "px; position: absolute; top: " & spanTop.ToString & "px"">" & _
                                                     " <input id=""opt" & sID & "_" & iTemp.ToString & """ type=""radio""" & _
                                                     " name=""opt" & sID & """ value=""" & drGrid(0).ToString & """" & _
                                                     " onfocus = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='#ff9608'; activateControl();}catch(e){};""" & _
                                                     " onblur = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='';}catch(e){};""" & _
                                                     " onclick = """ & sID & ".value=opt" & sID & "[" & iTemp.ToString & "].value;"""
                                                Case 1 ' Horizontal
                                                    stringSize = graphic.MeasureString(drGrid(0).ToString(), font)
                                                    Dim spanTop As Int32 = CInt((NullSafeInteger(dr("FontSize")) * 1.25) + 1) - _
                                                        CInt(IIf(NullSafeBoolean(dr("PictureBorder")), 0, 10))

                                                    sTemp = sTemp & _
                                                     "<span tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                                                     " style=""z-index: 0;" & _
                                                     " FONT-FAMILY: " & NullSafeString(dr("FontName")) & "; " & _
                                                     " FONT-SIZE: " & NullSafeString(dr("FontSize")) & "pt; " & _
                                                     " FONT-WEIGHT: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                                                     " FONT-STYLE: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                                                     " TEXT-DECORATION:" & sTemp2 & ";" & _
                                                     " left: " & lastLeft & "px; position: absolute; top: " & spanTop.ToString & "px"">" & _
                                                     " <input id=""opt" & sID & "_" & iTemp.ToString & """ type=""radio""" & _
                                                     " name=""opt" & sID & """ value=""" & drGrid(0).ToString & """" & _
                                                     " onfocus = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='#ff9608'; activateControl();}catch(e){};""" & _
                                                     " onblur = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='';}catch(e){};""" & _
                                                     " onclick = """ & sID & ".value=opt" & sID & "[" & iTemp.ToString & "].value;"""

                                                    lastLeft += (stringSize.Width + (font.Size * 2) + 28)
                                            End Select

                                            If iTemp = 0 Or CInt(drGrid.GetValue(1)) = 1 Then
                                                sTemp = sTemp & _
                                                 " Checked=""checked"""
                                                sDefaultValue = drGrid(0).ToString
                                            End If

                                            sTemp = sTemp & _
                                             "/>" & _
                                             " <label style=""position: absolute; left:20px; top:" & (10 - (0.9 * NullSafeInteger(dr("FontSize")))).ToString & "px"" id=""forOpt" & sID & "_" & iTemp.ToString & """ for=""opt" & sID & "_" & iTemp.ToString & """ tabindex=""-1""" _
                                             & " onmouseover = ""try{this.style.color='#ff9608'; }catch(e){};""" _
                                             & " onmouseout = ""try{this.style.color='';}catch(e){};""" _
                                             & ">" _
                                             & drGrid(0).ToString _
                                             & "</label>" & _
                                             " </span>"

                                            msRefreshLiteralsCode = msRefreshLiteralsCode & vbNewLine & _
                                             vbTab & vbTab & "try" & vbNewLine & _
                                             vbTab & vbTab & "{" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "if (frmMain.opt" & sID & "_" & iTemp.ToString & ".value == frmMain." & sID & ".value)" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "{" & vbNewLine & _
                                             vbTab & vbTab & vbTab & vbTab & "frmMain.opt" & sID & "_" & iTemp.ToString & ".checked = 'checked';" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "}" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "else" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "{" & vbNewLine & _
                                             vbTab & vbTab & vbTab & vbTab & "frmMain.opt" & sID & "_" & iTemp.ToString & ".checked = '';" & vbNewLine & _
                                             vbTab & vbTab & vbTab & "}" & vbNewLine & _
                                             vbTab & vbTab & "}" & vbNewLine & _
                                            vbTab & vbTab & "catch(e) {}"

                                            iTemp = iTemp + 1
                                        End While

                                        drGrid.Close()
                                        drGrid = Nothing
                                        cmdGrid.Dispose()
                                        cmdGrid = Nothing

                                        sTemp = sTemp & _
                                         "</fieldset>" & vbCrLf

                                        ctlForm_OptionGroup = New LiteralControl(sTemp)

                                        ' pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_OptionGroup)
                                        pnlInputDiv.Controls.Add(ctlForm_OptionGroup)

                                        ctlForm_OptionGroupReal = New TextBox
                                        With ctlForm_OptionGroupReal
                                            .Height = Unit.Parse("0")
                                            .Width = Unit.Parse("0")
                                            .TabIndex = 0
                                            .Style("visibility") = "hidden"
                                            .Text = sDefaultValue
                                            .ID = sID
                                        End With

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = "opt" & sID & "_0"
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_OptionGroupReal)
                                        pnlInputDiv.Controls.Add(ctlForm_OptionGroupReal)

                                    Catch ex As Exception
                                        sMessage = "Error loading web form option group values:<BR><BR>" & _
                                        ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                                        "Contact your system administrator."
                                        Exit While

                                    Finally
                                        connGrid.Close()
                                        connGrid.Dispose()
                                    End Try

                                Case 17 ' Input value - file upload
                                    ctlForm_Button = New Infragistics.WebUI.WebDataInput.WebImageButton
                                    With ctlForm_Button
                                        .ID = sID
                                        .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                                        If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                            sDefaultFocusControl = sID
                                            iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                        End If

                                        .Style("position") = "absolute"
                                        .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                                        .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                                        .Appearance.Style.BackColor = objGeneral.GetColour(NullSafeInteger(dr("BackColor")))
                                        .Appearance.Style.BorderStyle = BorderStyle.Solid
                                        .Appearance.Style.BorderWidth = 1
                                        .Appearance.InnerBorder.StyleTop = BorderStyle.None
                                        .Appearance.Style.BorderColor = objGeneral.GetColour(9999523)
                                        .Appearance.Style.ForeColor = objGeneral.GetColour(NullSafeInteger(dr("ForeColor")))
                                        .FocusAppearance.Style.BorderColor = objGeneral.GetColour(562943)
                                        .FocusAppearance.Style.BackColor = objGeneral.GetColour(12775933)
                                        .HoverAppearance.Style.BorderColor = objGeneral.GetColour(562943)


                                        .Text = NullSafeString(dr("caption"))
                                        .Font.Name = NullSafeString(dr("FontName"))
                                        .Font.Size = FontUnit.Parse(NullSafeString(dr("FontSize")))
                                        .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                                        .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                                        .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                                        .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                                        .Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                                        .Height = Unit.Pixel(NullSafeInteger(dr("Height")) - 7)

                                        .ClientSideEvents.Click = "try{showFileUpload(true, '" & sEncodedID & "', document.getElementById('file" & sID & "').value);}catch(e){};"

                                        AddHandler ctlForm_Button.Click, AddressOf Me.DisableControls
                                    End With

                                    'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Button)
                                    pnlInputDiv.Controls.Add(ctlForm_Button)

                                    ctlForm_HiddenField = New HiddenField
                                    With ctlForm_HiddenField
                                        .ID = "file" & sID
                                        .Value = NullSafeString(dr("value"))
                                    End With
                                    'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_HiddenField)
                                    pnlInputDiv.Controls.Add(ctlForm_HiddenField)

                                Case 19 ' DB File
                                    sDecoration = ""
                                    If NullSafeBoolean(dr("FontUnderline")) Then
                                        sDecoration = sDecoration & " underline"
                                    End If
                                    If NullSafeBoolean(dr("FontStrikeThru")) Then
                                        sDecoration = sDecoration & " line-through"
                                    End If
                                    If sDecoration.Length = 0 Then
                                        sDecoration = "none"
                                    End If

                                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                                        sBackColour = "Transparent"
                                    Else
                                        sBackColour = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColor")))
                                    End If

                                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                        sDefaultFocusControl = sID
                                        iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                    End If

                                    sTemp = "<span id='" & sID & "' tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                                     " style='POSITION: absolute; display:inline-block; word-wrap:break-word; overflow:auto; text-align:left;" & _
                                     " TOP: " & NullSafeString(dr("TopCoord")) & "px;" & _
                                     " LEFT: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                                     " HEIGHT:" & NullSafeString(dr("Height")) & "px;" & _
                                     " WIDTH:" & NullSafeInteger(dr("Width")) & "px;" & _
                                     " font-family:" & NullSafeString(dr("FontName")) & ";" & _
                                     " font-size:" & NullSafeString(dr("FontSize")).ToString & "pt;" & _
                                     " font-weight:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontBold"))), "bold;", "normal;").ToString & _
                                     " font-style:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontItalic"))), "italic;", "normal;").ToString & _
                                     " text-decoration:" & sDecoration & ";" & _
                                     " background-color: " & sBackColour & "; " & _
                                     " color: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                                     "' onclick='FileDownload_Click(""" & sEncodedID & """);'" & _
                                     " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" & _
                                     " onmouseover=""this.style.cursor='pointer';this.style.color='#ff9608';""" & _
                                     " onmouseout=""this.style.cursor='';this.style.color='" & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "';""" & _
                                     " onfocus=""this.style.color='#ff9608';""" & _
                                     " onblur=""this.style.color='" & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "';"">" & _
                                     HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) & _
                                     "</span>"

                                    ctlForm_Literal = New LiteralControl(sTemp)

                                    'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Literal)
                                    pnlInputDiv.Controls.Add(ctlForm_Literal)

                                Case 20 ' WF File
                                    sDecoration = ""
                                    If NullSafeBoolean(dr("FontUnderline")) Then
                                        sDecoration = sDecoration & " underline"
                                    End If
                                    If NullSafeBoolean(dr("FontStrikeThru")) Then
                                        sDecoration = sDecoration & " line-through"
                                    End If
                                    If sDecoration.Length = 0 Then
                                        sDecoration = "none"
                                    End If

                                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                                        sBackColour = "Transparent"
                                    Else
                                        sBackColour = objGeneral.GetHTMLColour(NullSafeInteger(dr("BackColor")))
                                    End If

                                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                                        sDefaultFocusControl = sID
                                        iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                                    End If

                                    sTemp = "<span id='" & sID & "' tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                                     " style='POSITION: absolute; display:inline-block; word-wrap:break-word; overflow:auto; text-align:left;" & _
                                     " TOP: " & NullSafeString(dr("TopCoord")) & "px;" & _
                                     " LEFT: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                                     " HEIGHT:" & NullSafeString(dr("Height")) & "px;" & _
                                     " WIDTH:" & NullSafeInteger(dr("Width")) & "px;" & _
                                     " font-family:" & NullSafeString(dr("FontName")) & ";" & _
                                     " font-size:" & NullSafeString(dr("FontSize")).ToString & "pt;" & _
                                     " font-weight:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontBold"))), "bold;", "normal;").ToString & _
                                     " font-style:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontItalic"))), "italic;", "normal;").ToString & _
                                     " text-decoration:" & sDecoration & ";" & _
                                     " background-color: " & sBackColour & "; " & _
                                     " color: " & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                                     "' onclick='FileDownload_Click(""" & sEncodedID & """);'" & _
                                     " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" & _
                                     " onmouseover=""this.style.cursor='pointer';this.style.color='#ff9608';""" & _
                                     " onmouseout=""this.style.cursor='';this.style.color='" & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "';""" & _
                                     " onfocus=""this.style.color='#ff9608';""" & _
                                     " onblur=""this.style.color='" & objGeneral.GetHTMLColour(NullSafeInteger(dr("ForeColor"))) & "';"">" & _
                                     HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) & _
                                     "</span>"

                                    ctlForm_Literal = New LiteralControl(sTemp)

                                    'pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Literal)
                                    pnlInputDiv.Controls.Add(ctlForm_Literal)

                            End Select
                        End While

        dr.Close()


                        If (Not ClientScript.IsStartupScriptRegistered("Startup")) Then
                            ' Form the script to be registered at client side.

                            scriptString += "}"
                            ClientScript.RegisterStartupScript(ClientScript.GetType, "Startup", scriptString, True)

                        End If

        If sMessage.Length = 0 Then
            If CStr(cmdSelect.Parameters("@psErrorMessage").Value) <> "" Then
                sMessage = CStr(cmdSelect.Parameters("@psErrorMessage").Value)
            Else
                sBackgroundImage = ""
                sBackgroundRepeat = ""
                sBackgroundPosition = ""
                If CInt(cmdSelect.Parameters("@piBackImage").Value) > 0 Then
                    sBackgroundImage = LoadPicture(CInt(cmdSelect.Parameters("@piBackImage").Value), sMessage)
                    If sMessage.Length = 0 Then
                        pnlInputDiv.Style("Background-image") = sBackgroundImage
                        ' pnlInput.BackImageUrl = sBackgroundImage

                    End If
                    If sMessage.Length = 0 Then
                        sBackgroundImage = "url('" & sBackgroundImage & "')"
                    End If

                    iBackgroundImagePosition = CInt(cmdSelect.Parameters("@piBackImageLocation").Value())
                    sBackgroundRepeat = objGeneral.BackgroundRepeat(CShort(iBackgroundImagePosition))
                    sBackgroundPosition = objGeneral.BackgroundPosition(CShort(iBackgroundImagePosition))
                End If
                'pnlInput.Style("background-repeat") = sBackgroundRepeat
                'pnlInput.Style("background-position") = sBackgroundPosition
                pnlInputDiv.Style("background-repeat") = sBackgroundRepeat
                pnlInputDiv.Style("background-position") = sBackgroundPosition



                sBackgroundColourHex = ""
                If Not IsDBNull(cmdSelect.Parameters("@piBackColour").Value) Then
                    iBackgroundColour = CInt(cmdSelect.Parameters("@piBackColour").Value())
                    sBackgroundColourHex = objGeneral.GetHTMLColour(iBackgroundColour).ToString()
                    pnlInputDiv.Style("Background-color") = objGeneral.GetHTMLColour(NullSafeInteger(iBackgroundColour))  ' objGeneral.GetColour(iBackgroundColour)
                End If

                iFormWidth = CInt(cmdSelect.Parameters("@piWidth").Value)
                iFormHeight = CInt(cmdSelect.Parameters("@piHeight").Value)
                pnlInputDiv.Style("width") = iFormWidth.ToString & "px"
                pnlInputDiv.Style("height") = iFormHeight.ToString & "px"

                hdnFormHeight.Value = iFormHeight.ToString
                hdnFormWidth.Value = iFormWidth.ToString
                hdnFormBackColourHex.Value = sBackgroundColourHex
                hdnFormBackImage.Value = sBackgroundImage
                hdnFormBackRepeat.Value = sBackgroundRepeat
                hdnFormBackPosition.Value = sBackgroundPosition

                hdnColourThemeHex.Value = mobjConfig.ColourThemeHex().ToString
                hdnSiblingForms.Value = sSiblingForms.ToString

                miCompletionMessageType = NullSafeInteger(cmdSelect.Parameters("@piCompletionMessageType").Value)
                msCompletionMessage = NullSafeString(cmdSelect.Parameters("@psCompletionMessage").Value)
                miSavedForLaterMessageType = NullSafeInteger(cmdSelect.Parameters("@piSavedForLaterMessageType").Value)
                msSavedForLaterMessage = NullSafeString(cmdSelect.Parameters("@psSavedForLaterMessage").Value)
                miFollowOnFormsMessageType = NullSafeInteger(cmdSelect.Parameters("@piFollowOnFormsMessageType").Value)
                msFollowOnFormsMessage = NullSafeString(cmdSelect.Parameters("@psFollowOnFormsMessage").Value)

                'Creates a new async trigger
                Dim trigger As New AsyncPostBackTrigger()

                'Sets the event name of the control
                'trigger.EventName = "goSubmit"
                'Adds the trigger to the UpdatePanels' triggers collection

                'Sets the control that will trigger a post-back on the UpdatePanel

                trigger.ControlID = "btnSubmit"
                pnlInput.Triggers.Add(trigger)

                pnlInput.UpdateMode = UpdatePanelUpdateMode.Conditional
                'pnlInput.ChildrenAsTriggers = False
                ' These two are now created clientside as PageRequestManager functions...
                'pnlInput.ClientSideEvents.RefreshRequest = "goSubmit();"

                'pnlInput.ClientSideEvents.RefreshComplete = "showMessage();"

                ' This is now created in the Page_Init event.
                ' pnlInput.ClientSideEvents.InitializePanel = "WARP_SetTimeout();"

                If sDefaultFocusControl.Length > 0 Then
                    frmMain.DefaultFocus = sDefaultFocusControl
                    hdnFirstControl.Value = sDefaultFocusControl
                Else
                    If Not ctlDefaultFocusControl Is Nothing Then
                        ctlDefaultFocusControl.Focus()
                    End If
                End If
            End If
        End If

        cmdSelect.Dispose()
                    End If

                Catch ex As Exception
                    sMessage = "Error loading web form controls:<BR><BR>" & _
                     ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                     "Contact your system administrator."
                Finally
                    conn.Close()
                    conn.Dispose()
                End Try

            Catch ex As Exception   ' conn creation 
                sMessage = "Error creating SQL connection:<BR><BR>" & _
                ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
                "Contact your system administrator."
            End Try
        End If

        If sMessage.Length > 0 Then
            Session("message") = sMessage

            If IsPostBack Then
                bulletErrors.Items.Clear()
                bulletWarnings.Items.Clear()

                hdnErrorMessage.Value = sMessage

                sMessage1 = sMessage & "<BR><BR>Click "
                sMessage2 = "here"
                sMessage3 = " to close this form."

                hdnSubmissionMessage_1.Value = Replace(sMessage1, " ", "&nbsp;")
                hdnSubmissionMessage_2.Value = Replace(sMessage2, " ", "&nbsp;")
                hdnSubmissionMessage_3.Value = Replace(sMessage3, " ", "&nbsp;")
                hdnNoSubmissionMessage.Value = CStr(IIf((sMessage1.Length = 0) And (sMessage2.Length = 0) And (sMessage3.Length = 0), "1", "0"))
                hdnFollowOnForms.Value = ""
            Else
                Response.Redirect("Message.aspx")
            End If
        End If




        Dim ctlFormInput As Control
        '   Dim sID As String
        Dim sIDString As String
        '    Dim iTemp As Int16
        '    Dim sTemp As String
        Dim iType As Int16
        Dim sType As String
        Dim ctlFormDropdown As System.Web.UI.WebControls.DropDownList   'Infragistics.WebUI.WebCombo.WebCombo
        '    Dim objGridColumn As Infragistics.WebUI.UltraWebGrid.UltraGridColumn
        '    Dim iLookupColumnIndex As Int16
        Dim ctlFormRecordSelectionGrid As System.Web.UI.WebControls.GridView ' Infragistics.WebUI.UltraWebGrid.UltraWebGrid

        ' Set the DataSource properties of each combo (dropdown list or lookup), 
        ' and then bind the DataSource to the combo.
        ' The DataSource property is set to the DataTable created on initial Page_Init,
        ' and stored in a sesion variable for each combo.

        If (Not IsPostBack) Then
            For Each ctlFormInput In pnlInput.ContentTemplateContainer.Controls
                sID = ctlFormInput.ID

                If (Left(sID, Len(FORMINPUTPREFIX)) = FORMINPUTPREFIX) Then
                    sIDString = sID.Substring(Len(FORMINPUTPREFIX))

                    iTemp = CShort(sIDString.IndexOf("_"))
                    sTemp = sIDString.Substring(iTemp + 1)
                    sIDString = sIDString.Substring(0, iTemp) & vbTab

                    iTemp = CShort(sTemp.IndexOf("_"))
                    sType = sTemp.Substring(0, iTemp)
                    iType = CShort(sType)

                    'If (iType = 11) Then
                    '    ' 11 = Record Selector
                    '    If (TypeOf ctlFormInput Is Infragistics.WebUI.UltraWebGrid.UltraWebGrid) Then
                    '        ctlFormRecordSelectionGrid = DirectCast(ctlFormInput, Infragistics.WebUI.UltraWebGrid.UltraWebGrid)

                    '        With ctlFormRecordSelectionGrid
                    '            'iLookupColumnIndex = CShort(.Attributes.Item("LookupColumnIndex"))
                    '            .DataSource = Session(sID & "_DATA")
                    '            .DataBind()

                    '            '' Format the column(s)
                    '            'For Each objGridColumn In .Columns
                    '            '    If objGridColumn.BaseColumnName.StartsWith("ASRSys") Then
                    '            '        objGridColumn.Hidden = True
                    '            '    Else
                    '            '        If iLookupColumnIndex = objGridColumn.Index Then
                    '            '            .DataTextField = objGridColumn.BaseColumnName

                    '            '            If objGridColumn.DataType = "System.Decimal" _
                    '            '            Or objGridColumn.DataType = "System.Int32" Then

                    '            '                .ClientSideEvents.AfterSelectChange = "ChangeLookup"
                    '            '            End If

                    '            '            .Attributes.Remove("DataType")
                    '            '            .Attributes.Add("DataType", objGridColumn.DataType)
                    '            '        End If
                    '            '        objGridColumn.AllowNull = False

                    '            '        If objGridColumn.DataType = "System.DateTime" Then
                    '            '            objGridColumn.Format = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern

                    '            '        ElseIf objGridColumn.DataType = "System.Boolean" Then
                    '            '            objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Center

                    '            '        ElseIf objGridColumn.DataType = "System.Decimal" _
                    '            '            Or objGridColumn.DataType = "System.Int32" Then

                    '            '            objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Right
                    '            '        End If
                    '            '    End If
                    '            'Next objGridColumn
                    '        End With


                    '    End If
                    'End If

                    'If ((iType = 13) Or (iType = 14)) Then
                    '  ' 13 = Dropdown Input
                    '  ' 14 = Lookup Input
                    '              If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then '  Infragistics.WebUI.WebCombo.WebCombo) Then
                    '                  ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)

                    '                  With ctlFormDropdown
                    '                      iLookupColumnIndex = CShort(.Attributes.Item("LookupColumnIndex"))
                    '                      .DataSource = Session(sID & "_DATA")
                    '                      .DataBind()

                    '                      ' Format the column(s)
                    '                      'For Each objGridColumn In .Columns
                    '                      '    If objGridColumn.BaseColumnName.StartsWith("ASRSys") Then
                    '                      '        objGridColumn.Hidden = True
                    '                      '    Else
                    '                      '        If iLookupColumnIndex = objGridColumn.Index Then
                    '                      '            .DataTextField = objGridColumn.BaseColumnName

                    '                      '            If objGridColumn.DataType = "System.Decimal" _
                    '                      '            Or objGridColumn.DataType = "System.Int32" Then

                    '                      '                .ClientSideEvents.AfterSelectChange = "ChangeLookup"
                    '                      '            End If

                    '                      '            .Attributes.Remove("DataType")
                    '                      '            .Attributes.Add("DataType", objGridColumn.DataType)

                    '                      '            If iType = 13 Then
                    '                      '                If .Rows.Count > MAXDROPDOWNROWS Then
                    '                      '                    ' Vertical scrollbar will be visible. Adjust the column width
                    '                      '                    .DropDownLayout.ColWidthDefault = Unit.Pixel(CInt(.DropDownLayout.DropdownWidth.Value) - 20)
                    '                      '                Else
                    '                      '                    ' Vertical scrollbar will NOT be visible.
                    '                      '                    .DropDownLayout.ColWidthDefault = Unit.Pixel(CInt(.DropDownLayout.DropdownWidth.Value) - 5)
                    '                      '                End If
                    '                      '            End If

                    '                      '        End If
                    '                      '        objGridColumn.AllowNull = False

                    '                      '        If objGridColumn.DataType = "System.DateTime" Then
                    '                      '            objGridColumn.Format = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern

                    '                      '        ElseIf objGridColumn.DataType = "System.Boolean" Then
                    '                      '            objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Center

                    '                      '        ElseIf objGridColumn.DataType = "System.Decimal" _
                    '                      '            Or objGridColumn.DataType = "System.Int32" Then

                    '                      '            objGridColumn.CellStyle.HorizontalAlign = HorizontalAlign.Right
                    '                      '        End If
                    '                      '    End If
                    '                      'Next objGridColumn
                    '                  End With
                    '              End If
                    'End If
                End If
            Next ctlFormInput
        End If


    End Sub


    Public Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) '    Infragistics.WebUI.WebDataInput.ButtonEventArgs)

        Dim objGeneral As New General
        Dim strConn As String
        Dim conn As System.Data.SqlClient.SqlConnection
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim cmdValidate As System.Data.SqlClient.SqlCommand
        Dim cmdUpdate As System.Data.SqlClient.SqlCommand
        Dim cmdQS As System.Data.SqlClient.SqlCommand
        Dim sFormInput1 As String
        Dim sFormValidation1 As String
        Dim ctlFormInput As Control
        Dim sID As String
        Dim ctlFormCheckBox As CheckBox
        Dim ctlFormTextInput As TextBox
        Dim ctlFormDateInput As Infragistics.WebUI.WebSchedule.WebDateChooser
        Dim ctlFormNumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
        ' Dim ctlFormRecordSelectionGrid As Infragistics.WebUI.UltraWebGrid.UltraWebGrid
        Dim ctlForm_PagingGridView As RecordSelector
        Dim ctlFormDropdown As System.Web.UI.WebControls.DropDownList ' Infragistics.WebUI.WebCombo.WebCombo
        Dim ctlForm_HiddenField As HiddenField
        Dim sMessage As String
        Dim sIDString As String
        Dim sDateValueString As String
        Dim sNumValueString As String
        Dim iTemp As Int16
        Dim sTemp As String
        Dim iType As Int16
        Dim sType As String
        Dim sRecordID As String
        ' Dim objGridColumn As Infragistics.WebUI.UltraWebGrid.UltraGridColumn
        Dim sFormElements As String
        Dim arrFollowOnForms() As String
        Dim fSavedForLater As Boolean
        Dim sMessage1 As String
        Dim sMessage2 As String
        Dim sMessage3 As String
        Dim iFollowOnFormCount As Integer
        Dim iIndex As Integer
        Dim sStep As String
        Dim sQueryString As String
        Dim arrQueryStrings() As String
        Dim sFollowOnForms As String
        Dim sColumnCaption As String

        sMessage = ""
        fSavedForLater = False
        iFollowOnFormCount = 0
        sMessage1 = ""
        sMessage2 = ""
        sMessage3 = ""
        sFormInput1 = ""
        sFormValidation1 = ""
        sFollowOnForms = ""
        ReDim arrQueryStrings(0)
        Try ' Read the web form item values.
            ' Build up a string of the form input values.
            ' This is a tab delimited string of itemIDs and values.

            For Each ctlFormInput In pnlInputDiv.Controls   ' pnlInput.ContentTemplateContainer.Controls
                sID = ctlFormInput.ID

                If (Left(sID, Len(FORMINPUTPREFIX)) = FORMINPUTPREFIX) Then
                    sIDString = sID.Substring(Len(FORMINPUTPREFIX))
                    iTemp = CShort(sIDString.IndexOf("_"))
                    sTemp = sIDString.Substring(iTemp + 1)
                    sIDString = sIDString.Substring(0, iTemp) & vbTab

                    iTemp = CShort(sTemp.IndexOf("_"))
                    sType = sTemp.Substring(0, iTemp)
                    iType = CShort(sType)

                    Select Case iType
                        Case 0 ' Button
                            Dim btn As Infragistics.WebUI.WebDataInput.WebImageButton = _
                             DirectCast(sender, Infragistics.WebUI.WebDataInput.WebImageButton)

                            If (ctlFormInput.ID = btn.ID) Then
                                hdnLastButtonClicked.Value = btn.ID
                                sFormInput1 = sFormInput1 & sIDString & "1" & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & "1" & vbTab
                            Else
                                If (TypeOf ctlFormInput Is Infragistics.WebUI.WebDataInput.WebImageButton) Then
                                    sFormInput1 = sFormInput1 & sIDString & "0" & vbTab
                                    sFormValidation1 = sFormValidation1 & sIDString & "0" & vbTab
                                End If
                            End If

                        Case 3 ' Character Input

                            If (TypeOf ctlFormInput Is TextBox) Then
                                ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
                                sFormInput1 = sFormInput1 & sIDString & Replace(ctlFormTextInput.Text, vbTab, " ") & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & Replace(ctlFormTextInput.Text, vbTab, " ") & vbTab
                            End If

                        Case 5 ' Numeric Input
                            If (TypeOf ctlFormInput Is Infragistics.WebUI.WebDataInput.WebNumericEdit) Then
                                ctlFormNumericInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebDataInput.WebNumericEdit)
                                sNumValueString = CStr(IIf(IsDBNull(ctlFormNumericInput.Value), "0", CStr(ctlFormNumericInput.Value).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")))
                                sFormInput1 = sFormInput1 & sIDString & sNumValueString & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & sNumValueString & vbTab
                            End If

                        Case 6 ' Logic Input
                            If (TypeOf ctlFormInput Is CheckBox) Then
                                ctlFormCheckBox = DirectCast(ctlFormInput, CheckBox)
                                'sFormInput1 = sFormInput1 & sIDString & IIf(UCase(ctlFormCheckBox.Checked) = "TRUE", "1", "0") & vbTab
                                sFormInput1 = sFormInput1 & sIDString & CStr(IIf(ctlFormCheckBox.Checked, "1", "0")) & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & CStr(IIf(ctlFormCheckBox.Checked, "1", "0")) & vbTab
                            End If

                        Case 7 ' Date Input
                            If (TypeOf ctlFormInput Is Infragistics.WebUI.WebSchedule.WebDateChooser) Then
                                ctlFormDateInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebSchedule.WebDateChooser)

                                If (ctlFormDateInput.Text = ctlFormDateInput.NullDateLabel) Then
                                    sDateValueString = "null"
                                Else
                                    sDateValueString = Format(ctlFormDateInput.Value, "MM/dd/yyyy")
                                End If

                                sFormInput1 = sFormInput1 & sIDString & sDateValueString & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & sDateValueString & vbTab
                            End If

                            'Case 11 ' Grid (RecordSelector) Input
                            '    If (TypeOf ctlFormInput Is Infragistics.WebUI.UltraWebGrid.UltraWebGrid) Then
                            '        ctlFormRecordSelectionGrid = DirectCast(ctlFormInput, Infragistics.WebUI.UltraWebGrid.UltraWebGrid)
                            '        sRecordID = "0"

                            '        If ctlFormRecordSelectionGrid.DisplayLayout.SelectedRows.Count > 0 Then
                            '            For Each objGridColumn In ctlFormRecordSelectionGrid.Columns
                            '                sColumnCaption = UCase(objGridColumn.Header.Caption)

                            '                If (sColumnCaption = "ID") Then
                            '                    sRecordID = ctlFormRecordSelectionGrid.DisplayLayout.SelectedRows(0).GetCellText(objGridColumn)
                            '                    Exit For
                            '                End If
                            '            Next objGridColumn
                            '        End If

                            '        sFormInput1 = sFormInput1 & sIDString & sRecordID & vbTab
                            '        sFormValidation1 = sFormValidation1 & sIDString & sRecordID & vbTab
                            '    End If


                        Case 11 ' Grid (RecordSelector) Input
                            If (TypeOf ctlFormInput Is RecordSelector) Then    ' Infragistics.WebUI.UltraWebGrid.UltraWebGrid) Then
                                ctlForm_PagingGridView = DirectCast(ctlFormInput, RecordSelector)
                                sRecordID = "0"

                                If Not ctlForm_PagingGridView.IsEmpty And ctlForm_PagingGridView.SelectedIndex >= 0 Then
                                    For iColCount As Integer = 0 To ctlForm_PagingGridView.HeaderRow.Cells.Count - 1
                                        sColumnCaption = UCase(ctlForm_PagingGridView.HeaderRow.Cells(iColCount).Text)

                                        If (sColumnCaption = "ID") Then
                                            sRecordID = ctlForm_PagingGridView.SelectedRow.Cells(iColCount).Text
                                            Exit For
                                        End If
                                    Next iColCount
                                End If

                                sFormInput1 = sFormInput1 & sIDString & sRecordID & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & sRecordID & vbTab
                            End If



                        Case 13 ' Dropdown Input
                            If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                                ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)

                                sTemp = ctlFormDropdown.Text ' .DisplayValue
                                sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
                            End If



                            'Case 14 ' Lookup Input
                            '    If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                            '        ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)

                            '        sTemp = ctlFormDropdown.Text ' .DisplayValue

                            '        If ctlFormDropdown.Attributes("DataType") = "System.DateTime" Then
                            '            If sTemp Is Nothing Then
                            '                sTemp = "null"
                            '            Else
                            '                If (sTemp.Length = 0) Then
                            '                    sTemp = "null"
                            '                Else
                            '                    sTemp = objGeneral.ConvertLocaleDateToSQL(sTemp)
                            '                End If
                            '            End If
                            '        ElseIf ctlFormDropdown.Attributes("DataType") = "System.Decimal" _
                            '         Or ctlFormDropdown.Attributes("DataType") = "System.Int32" Then

                            '            If sTemp Is Nothing Then
                            '                sTemp = ""
                            '            Else
                            '                sTemp = CStr(IIf(sTemp.Length = 0, "", CStr(sTemp).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")))
                            '            End If

                            '        End If

                            '        sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
                            '        sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
                            '    End If


                        Case 14 ' Lookup Input

                            If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                                ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)

                                sTemp = ctlFormDropdown.Text ' .DisplayValue

                                If ctlFormDropdown.Attributes("DataType") = "System.DateTime" Then
                                    If sTemp Is Nothing Then
                                        sTemp = "null"
                                    Else
                                        If (sTemp.Length = 0) Then
                                            sTemp = "null"
                                        Else
                                            sTemp = objGeneral.ConvertLocaleDateToSQL(sTemp)
                                        End If
                                    End If
                                ElseIf ctlFormDropdown.Attributes("DataType") = "System.Decimal" _
                                 Or ctlFormDropdown.Attributes("DataType") = "System.Int32" Then

                                    If sTemp Is Nothing Then
                                        sTemp = ""
                                    Else
                                        sTemp = CStr(IIf(sTemp.Length = 0, "", CStr(sTemp).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")))
                                    End If

                                End If

                                sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
                            End If


                        Case 15 ' OptionGroup Input
                            If (TypeOf ctlFormInput Is TextBox) Then
                                ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
                                sFormInput1 = sFormInput1 & sIDString & ctlFormTextInput.Text & vbTab
                                sFormValidation1 = sFormValidation1 & sIDString & ctlFormTextInput.Text & vbTab
                            End If

                        Case 17 ' FileUpload
                            If (TypeOf ctlFormInput Is Infragistics.WebUI.WebDataInput.WebImageButton) Then

                                If pnlInput.FindControl("file" & sID) Is Nothing Then
                                    sFormValidation1 = sFormValidation1 & sIDString & "0" & vbTab
                                    sFormInput1 = sFormInput1 & sIDString & "0" & vbTab
                                Else
                                    ctlForm_HiddenField = DirectCast(pnlInput.FindControl("file" & sID), HiddenField)
                                    sFormValidation1 = sFormValidation1 & sIDString & ctlForm_HiddenField.Value.ToString & vbTab
                                    sFormInput1 = sFormInput1 & sIDString & ctlForm_HiddenField.Value.ToString & vbTab
                                End If
                            End If

                    End Select
                End If
            Next ctlFormInput

        Catch ex As Exception
            sMessage = "Error reading web form item values:<BR><BR>" & ex.Message
        End Try

        If sMessage.Length = 0 Then
            Try ' Open the database connection
                strConn = "Application Name=HR Pro Workflow;Data Source=" & msServer & ";Initial Catalog=" & msDatabase & ";Integrated Security=false;User ID=" & msUser & ";Password=" & msPwd & ";Pooling=false"
                conn = New SqlClient.SqlConnection(strConn)
                conn.Open()

                Try ' Validate the web form entry.
                    lblErrors.Font.Size = mobjConfig.ValidationMessageFontSize
                    lblErrors.ForeColor = objGeneral.GetColour(6697779)

                    lblWarnings.Font.Size = mobjConfig.ValidationMessageFontSize
                    lblWarnings.ForeColor = objGeneral.GetColour(6697779)
                    lblWarningsPrompt_1.Font.Size = mobjConfig.ValidationMessageFontSize
                    lblWarningsPrompt_1.ForeColor = objGeneral.GetColour(6697779)
                    lblWarningsPrompt_2.Font.Size = mobjConfig.ValidationMessageFontSize
                    lblWarningsPrompt_3.Font.Size = mobjConfig.ValidationMessageFontSize
                    lblWarningsPrompt_3.ForeColor = objGeneral.GetColour(6697779)

                    bulletErrors.Font.Size = mobjConfig.ValidationMessageFontSize
                    bulletErrors.ForeColor = objGeneral.GetColour(6697779)

                    bulletWarnings.Font.Size = mobjConfig.ValidationMessageFontSize
                    bulletWarnings.ForeColor = objGeneral.GetColour(6697779)

                    bulletErrors.Items.Clear()
                    bulletWarnings.Items.Clear()

                    cmdValidate = New SqlClient.SqlCommand("spASRSysWorkflowWebFormValidation", conn)
                    cmdValidate.CommandType = CommandType.StoredProcedure
                    cmdValidate.CommandTimeout = miSubmissionTimeoutInSeconds

                    cmdValidate.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                    cmdValidate.Parameters("@piInstanceID").Value = miInstanceID

                    cmdValidate.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                    cmdValidate.Parameters("@piElementID").Value = miElementID

                    cmdValidate.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
                    cmdValidate.Parameters("@psFormInput1").Value = sFormValidation1

                    dr = cmdValidate.ExecuteReader

                    While dr.Read
                        If NullSafeInteger(dr("failureType")) = 0 Then
                            bulletErrors.Items.Add(NullSafeString(dr("Message")))
                        ElseIf CDbl(hdnOverrideWarnings.Value) <> 1 Then
                            bulletWarnings.Items.Add(NullSafeString(dr("Message")))
                        End If
                    End While

                    dr.Close()
                    cmdValidate.Dispose()

                    hdnCount_Errors.Value = CStr(bulletErrors.Items.Count)
                    hdnCount_Warnings.Value = CStr(bulletWarnings.Items.Count)
                    hdnOverrideWarnings.Value = CStr(0)

                    lblErrors.Text = CStr(IIf(bulletErrors.Items.Count > 0, _
                     "Unable to submit this form due to the following error" & _
                     CStr(IIf(bulletErrors.Items.Count = 1, "", "s")) & ":", _
                     ""))

                    lblWarnings.Text = CStr(IIf(bulletWarnings.Items.Count > 0, _
                     CStr(IIf(bulletErrors.Items.Count > 0, _
                     "And the following warning" & _
                    CStr(IIf(bulletWarnings.Items.Count = 1, "", "s")) & ":", _
                     "Submitting this form raises the following warning" & _
                    CStr(IIf(bulletWarnings.Items.Count = 1, "", "s")) & ":")), _
                     ""))

                    lblWarningsPrompt_1.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)
                    lblWarningsPrompt_2.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)
                    lblWarningsPrompt_3.Text = "to ignore " & _
                     CStr(IIf(bulletWarnings.Items.Count = 1, "this warning", "these warnings")) & " and submit the form."
                    lblWarningsPrompt_3.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)

                Catch ex As Exception
                    sMessage = "Error validating the web form:<BR><BR>" & ex.Message
                End Try

                ' Submit the webform
                If (sMessage.Length = 0) _
                 And (bulletWarnings.Items.Count = 0) _
                 And (bulletErrors.Items.Count = 0) Then

                    Using (New TransactionScope(TransactionScopeOption.Suppress))
                        Try
                            cmdUpdate = New SqlClient.SqlCommand("spASRSubmitWorkflowStep", conn)
                            cmdUpdate.CommandType = CommandType.StoredProcedure
                            cmdUpdate.CommandTimeout = miSubmissionTimeoutInSeconds

                            cmdUpdate.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                            cmdUpdate.Parameters("@piInstanceID").Value = miInstanceID

                            cmdUpdate.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                            cmdUpdate.Parameters("@piElementID").Value = miElementID

                            cmdUpdate.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
                            cmdUpdate.Parameters("@psFormInput1").Value = sFormInput1

                            cmdUpdate.Parameters.Add("@psFormElements", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
                            cmdUpdate.Parameters.Add("@pfSavedForLater", SqlDbType.Bit).Direction = ParameterDirection.Output

                            cmdUpdate.ExecuteNonQuery()

                            sFormElements = CStr(cmdUpdate.Parameters("@psFormElements").Value())
                            fSavedForLater = CBool(cmdUpdate.Parameters("@pfSavedForLater").Value())

                            cmdUpdate.Dispose()

                            If fSavedForLater Then
                                Select Case miSavedForLaterMessageType
                                    Case 1 ' Custom
                                        If Not objGeneral.SplitMessage(msSavedForLaterMessage, sMessage1, sMessage2, sMessage3) Then
                                            sMessage1 = "Workflow step saved for later.<BR><BR>Click "
                                            sMessage2 = "here"
                                            sMessage3 = " to close this form."
                                        End If
                                    Case 2 ' None
                                        sMessage1 = ""
                                        sMessage2 = ""
                                        sMessage3 = ""
                                    Case Else   'System default
                                        sMessage1 = "Workflow step saved for later.<BR><BR>Click "
                                        sMessage2 = "here"
                                        sMessage3 = " to close this form."
                                End Select

                            ElseIf sFormElements.Length = 0 Then
                                Select Case miCompletionMessageType
                                    Case 1 ' Custom
                                        If Not objGeneral.SplitMessage(msCompletionMessage, sMessage1, sMessage2, sMessage3) Then
                                            sMessage1 = "Workflow step completed.<BR><BR>Click "
                                            sMessage2 = "here"
                                            sMessage3 = " to close this form."
                                        End If
                                    Case 2 ' None
                                        sMessage1 = ""
                                        sMessage2 = ""
                                        sMessage3 = ""
                                    Case Else   'System default
                                        sMessage1 = "Workflow step completed.<BR><BR>Click "
                                        sMessage2 = "here"
                                        sMessage3 = " to close this form."
                                End Select
                            Else
                                arrFollowOnForms = sFormElements.Split(CChar(vbTab))
                                iFollowOnFormCount = arrFollowOnForms.GetUpperBound(0)

                                For iIndex = 0 To iFollowOnFormCount - 1
                                    sStep = arrFollowOnForms(iIndex)

                                    cmdQS = New SqlClient.SqlCommand("spASRGetWorkflowQueryString", conn)
                                    cmdQS.CommandType = CommandType.StoredProcedure
                                    cmdQS.CommandTimeout = miSubmissionTimeoutInSeconds

                                    cmdQS.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                                    cmdQS.Parameters("@piInstanceID").Value = miInstanceID

                                    cmdQS.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                                    cmdQS.Parameters("@piElementID").Value = CLng(sStep)

                                    cmdQS.Parameters.Add("@psQueryString", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                    cmdQS.ExecuteNonQuery()

                                    sQueryString = CStr(cmdQS.Parameters("@psQueryString").Value())

                                    ReDim Preserve arrQueryStrings(arrQueryStrings.GetUpperBound(0) + 1)
                                    arrQueryStrings(arrQueryStrings.GetUpperBound(0)) = sQueryString

                                    cmdQS.Dispose()
                                Next iIndex

                                sFollowOnForms = Join(arrQueryStrings, vbTab)

                                Select Case miFollowOnFormsMessageType
                                    Case 1 ' Custom
                                        If Not objGeneral.SplitMessage(msFollowOnFormsMessage, sMessage1, sMessage2, sMessage3) Then
                                            sMessage1 = "Workflow step completed.<BR><BR>Click "
                                            sMessage2 = "here"
                                            sMessage3 = " to complete the follow-on Workflow form" & _
                                             CStr(IIf(iFollowOnFormCount = 1, "", "s")) & "."
                                        End If
                                    Case 2 ' None
                                        sMessage1 = ""
                                        sMessage2 = ""
                                        sMessage3 = ""
                                    Case Else   'System default
                                        sMessage1 = "Workflow step completed.<BR><BR>Click "
                                        sMessage2 = "here"
                                        sMessage3 = " to complete the follow-on Workflow form" & _
                                         CStr(IIf(iFollowOnFormCount = 1, "", "s")) & "."
                                End Select

                            End If

                            sMessage1 = NullSafeString(sMessage1)
                            sMessage2 = NullSafeString(sMessage2)
                            sMessage3 = NullSafeString(sMessage3)

                            hdnSubmissionMessage_1.Value = Replace(sMessage1, " ", "&nbsp;")
                            hdnSubmissionMessage_2.Value = Replace(sMessage2, " ", "&nbsp;")
                            hdnSubmissionMessage_3.Value = Replace(sMessage3, " ", "&nbsp;")
                            hdnNoSubmissionMessage.Value = CStr(IIf((sMessage1.Length = 0) And (sMessage2.Length = 0) And (sMessage3.Length = 0), "1", "0"))
                            hdnFollowOnForms.Value = sFollowOnForms

                            If hdnNoSubmissionMessage.Value <> "1" Then
                                EnableDisableControls(False)
                            End If

                        Catch ex As Exception
                            sMessage = "Error submitting the web form:<BR><BR>" & ex.Message
                        End Try
                    End Using
                End If

                conn.Close()
                conn.Dispose()

            Catch ex As Exception
                sMessage = "Error connecting to the database:<BR><BR>" & ex.Message
            End Try
        End If

        If sMessage.Length > 0 Then
            bulletErrors.Items.Clear()
            bulletWarnings.Items.Clear()

            hdnErrorMessage.Value = sMessage

            sMessage1 = sMessage & "<BR><BR>Click "
            sMessage2 = "here"
            sMessage3 = " to close this form."

            hdnSubmissionMessage_1.Value = Replace(sMessage1, " ", "&nbsp;")
            hdnSubmissionMessage_2.Value = Replace(sMessage2, " ", "&nbsp;")
            hdnSubmissionMessage_3.Value = Replace(sMessage3, " ", "&nbsp;")
            hdnNoSubmissionMessage.Value = CStr(IIf((sMessage1.Length = 0) And (sMessage2.Length = 0) And (sMessage3.Length = 0), "1", "0"))
            hdnFollowOnForms.Value = ""
            EnableDisableControls(False)
        End If

    End Sub
    Public Sub DisableControls(ByVal sender As System.Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs)
        EnableDisableControls(False)
    End Sub
    Public Sub EnableControls(ByVal sender As System.Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs)
        EnableDisableControls(True)
    End Sub
    Private Sub EnableDisableControls(ByVal pfEnabled As Boolean)

        Dim ctlFormInput As Control
        Dim sID As String
        Dim ctlFormCheckBox As CheckBox
        Dim ctlFormTextInput As TextBox
        Dim ctlFormButtonInput As Infragistics.WebUI.WebDataInput.WebImageButton
        Dim ctlFormDateInput As Infragistics.WebUI.WebSchedule.WebDateChooser
        Dim ctlFormNumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
        Dim ctlFormRecordSelectionGrid As System.Web.UI.WebControls.GridView ' Infragistics.WebUI.UltraWebGrid.UltraWebGrid
        Dim ctlFormDropdown As System.Web.UI.WebControls.DropDownList ' Infragistics.WebUI.WebCombo.WebCombo
        Dim ctlForm_DDE As AjaxControlToolkit.DropDownExtender
        Dim sMessage As String
        Dim sIDString As String
        Dim iTemp As Int16
        Dim sTemp As String
        Dim iType As Int16
        Dim sType As String
        Dim sMessage1 As String
        Dim sMessage2 As String
        Dim sMessage3 As String

        sMessage = ""
        sMessage1 = ""
        sMessage2 = ""
        sMessage3 = ""

        Try ' Disable all controls.
            For Each ctlFormInput In pnlInput.ContentTemplateContainer.Controls
                sID = ctlFormInput.ID

                If (Left(sID, Len(FORMINPUTPREFIX)) = FORMINPUTPREFIX) Then
                    sIDString = sID.Substring(Len(FORMINPUTPREFIX))

                    iTemp = CShort(sIDString.IndexOf("_"))
                    sTemp = sIDString.Substring(iTemp + 1)
                    sIDString = sIDString.Substring(0, iTemp) & vbTab

                    iTemp = CShort(sTemp.IndexOf("_"))
                    sType = sTemp.Substring(0, iTemp)
                    iType = CShort(sType)

                    Select Case iType
                        Case 0 ' Button
                            ctlFormButtonInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebDataInput.WebImageButton)
                            ctlFormButtonInput.Enabled = pfEnabled

                        Case 1 ' Database value
                        Case 2 ' Label

                        Case 3 ' Character Input
                            If (TypeOf ctlFormInput Is TextBox) Then
                                ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
                                ctlFormTextInput.Enabled = pfEnabled
                            End If

                        Case 4 ' Workflow value

                        Case 5 ' Numeric Input
                            If (TypeOf ctlFormInput Is Infragistics.WebUI.WebDataInput.WebNumericEdit) Then
                                ctlFormNumericInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebDataInput.WebNumericEdit)
                                ctlFormNumericInput.Enabled = pfEnabled
                            End If

                        Case 6 ' Logic Input
                            If (TypeOf ctlFormInput Is CheckBox) Then
                                ctlFormCheckBox = CType(ctlFormInput, CheckBox)
                                ctlFormCheckBox.Enabled = pfEnabled
                            End If

                        Case 7 ' Date Input
                            If (TypeOf ctlFormInput Is Infragistics.WebUI.WebSchedule.WebDateChooser) Then
                                ctlFormDateInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebSchedule.WebDateChooser)
                                ctlFormDateInput.Enabled = pfEnabled
                            End If

                        Case 8 ' Frame
                        Case 9 ' Line
                        Case 10 ' Image

                        Case 11 ' Grid (RecordSelector) Input                            
                            If (TypeOf ctlFormInput Is System.Web.UI.WebControls.GridView) Then 'Infragistics.WebUI.UltraWebGrid.UltraWebGrid) Then
                                ctlFormRecordSelectionGrid = DirectCast(ctlFormInput, System.Web.UI.WebControls.GridView) ' Infragistics.WebUI.UltraWebGrid.UltraWebGrid)
                                ctlFormRecordSelectionGrid.Enabled = pfEnabled
                                If pfEnabled Then
                                    'ctlFormRecordSelectionGrid.DisplayLayout.ReadOnly = Infragistics.WebUI.UltraWebGrid.ReadOnly.NotSet
                                Else
                                    'ctlFormRecordSelectionGrid.DisplayLayout.ReadOnly = Infragistics.WebUI.UltraWebGrid.ReadOnly.LevelZero
                                End If
                            End If

                        Case 13 ' Dropdown Input
                            If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                                ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)
                                ctlFormDropdown.Enabled = pfEnabled
                            End If

                        Case 14 ' Lookup Input
                            If (TypeOf ctlFormInput Is AjaxControlToolkit.DropDownExtender) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                                ctlForm_DDE = DirectCast(ctlFormInput, AjaxControlToolkit.DropDownExtender) 'Infragistics.WebUI.WebCombo.WebCombo)
                                ctlForm_DDE.Enabled = pfEnabled
                            End If

                        Case 15 ' OptionGroup Input
                            If (TypeOf ctlFormInput Is TextBox) Then
                                ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
                                ctlFormTextInput.Enabled = pfEnabled
                            End If

                        Case 17 ' Input value - file upload
                            ctlFormButtonInput = DirectCast(ctlFormInput, Infragistics.WebUI.WebDataInput.WebImageButton)
                            ctlFormButtonInput.Enabled = pfEnabled

                    End Select
                End If
            Next ctlFormInput

        Catch ex As Exception
            If pfEnabled Then
                sMessage = "Error enabling web form items:<BR><BR>" & ex.Message
            Else
                sMessage = "Error disabling web form items:<BR><BR>" & ex.Message
            End If
        End Try

        If sMessage.Length > 0 Then
            bulletErrors.Items.Clear()
            bulletWarnings.Items.Clear()

            hdnErrorMessage.Value = sMessage

            sMessage1 = sMessage & "<BR><BR>Click "
            sMessage2 = "here"
            sMessage3 = " to close this form."

            hdnSubmissionMessage_1.Value = Replace(sMessage1, " ", "&nbsp;")
            hdnSubmissionMessage_2.Value = Replace(sMessage2, " ", "&nbsp;")
            hdnSubmissionMessage_3.Value = Replace(sMessage3, " ", "&nbsp;")
            hdnNoSubmissionMessage.Value = CStr(IIf((sMessage1.Length = 0) And (sMessage2.Length = 0) And (sMessage3.Length = 0), "1", "0"))
            hdnFollowOnForms.Value = ""
        End If
    End Sub
    Public Function LocaleDateFormat() As String
        LocaleDateFormat = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern.ToUpper
    End Function
    Public Function LocaleDecimal() As String
        LocaleDecimal = Thread.CurrentThread.CurrentUICulture.NumberFormat.NumberDecimalSeparator
    End Function
    Public Function ColourThemeHex() As String
        ColourThemeHex = mobjConfig.ColourThemeHex
    End Function
    Public Function ColourThemeFolder() As String
        ColourThemeFolder = mobjConfig.ColourThemeFolder
    End Function
    'Public Function SubmissionTimeout() As String
    '    SubmissionTimeout = mobjConfig.SubmissionTimeout.ToString
    'End Function
    Public Function SubmissionTimeout() As Int32
        SubmissionTimeout = mobjConfig.SubmissionTimeout
    End Function


    'Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs) Handles btnSubmit.Click
    '    ButtonClick(sender, e)
    'End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        ButtonClick(sender, e)
    End Sub

    Private Function LoadPicture(ByVal piPictureID As Int32, _
     ByRef psErrorMessage As String) As String

        Dim strConn As String
        Dim conn As System.Data.SqlClient.SqlConnection
        Dim cmdSelect As System.Data.SqlClient.SqlCommand
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim sImageFileName As String
        Dim sImageFilePath As String
        Dim sTempName As String
        Dim fs As System.IO.FileStream
        Dim bw As System.IO.BinaryWriter
        Dim iBufferSize As Integer = 100
        Dim outByte(iBufferSize - 1) As Byte
        Dim retVal As Long
        Dim startIndex As Long = 0
        Dim sExtension As String = ""
        Dim iIndex As Integer
        Dim sName As String

        Try
            miImageCount = CShort(miImageCount + 1)

            psErrorMessage = ""
            LoadPicture = ""
            sImageFileName = ""
            sImageFilePath = Server.MapPath("pictures")

            strConn = "Application Name=HR Pro Workflow;Data Source=" & msServer & ";Initial Catalog=" & msDatabase & ";Integrated Security=false;User ID=" & msUser & ";Password=" & msPwd & ";Pooling=false"
            conn = New SqlClient.SqlConnection(strConn)
            conn.Open()

            cmdSelect = New SqlClient.SqlCommand
            cmdSelect.CommandText = "spASRGetPicture"
            cmdSelect.Connection = conn
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.CommandTimeout = miSubmissionTimeoutInSeconds

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

                    sImageFileName = Session.SessionID().ToString & _
                     "_" & miImageCount.ToString & _
                     "_" & Date.Now.Ticks.ToString & _
                     sExtension
                    sTempName = sImageFilePath & "\" & sImageFileName

                    ' Create a file to hold the output.
                    fs = New System.IO.FileStream(sTempName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
                    bw = New System.IO.BinaryWriter(fs)

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



    Public Function RefreshLiteralsCode() As String
        Dim sb As New StringBuilder

        Try
            sb.AppendLine("function refreshLiterals()")
            sb.AppendLine("{")
            sb.AppendLine(vbTab & " try")
            sb.AppendLine(vbTab & "{")
            sb.AppendLine(msRefreshLiteralsCode)
            sb.AppendLine(vbTab & "}")
            sb.AppendLine(vbTab & "catch(e) {}")
            sb.AppendLine(vbTab & "}")
            Return sb.ToString

        Catch ex As Exception
            Return ""
        End Try

    End Function









    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        Dim cs As ClientScriptManager

        cs = Page.ClientScript
        If Not cs.IsClientScriptBlockRegistered("RefreshLiteralsCode") Then
            cs.RegisterClientScriptBlock(Me.GetType, "RefreshLiteralsCode", RefreshLiteralsCode, True)
        End If
    End Sub


    'Protected Sub btnReEnableControls_Click(ByVal sender As Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs) Handles btnReEnableControls.Click
    '    EnableDisableControls(True)

    'End Sub

    Protected Sub btnReEnableControls_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReEnableControls.Click
        EnableDisableControls(True)
    End Sub


    'Public Sub InitializeLookupData(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DropDownList) ' Infragistics.WebUI.WebCombo.WebComboEventArgs)
    '    'Dim sID As String

    '    ' ''Dim objCombo As Infragistics.WebUI.WebCombo.WebCombo = _
    '    ' ''    DirectCast(sender, Infragistics.WebUI.WebCombo.WebCombo)

    '    'Dim objCombo As System.Web.UI.WebControls.DropDownList = DirectCast(sender, System.Web.UI.WebControls.DropDownList)

    '    'Try
    '    '    With objCombo
    '    '        sID = .ID
    '    '        .DataSource = Session(sID & "_DATA")

    '    '        If IsPostBack Then
    '    '            ''.Attributes("DefaultValue") = .Text
    '    '        End If

    '    '    End With
    '    'Catch ex As Exception
    '    'End Try

    'End Sub

    'Public Sub LookupDataBound(ByVal sender As Object, ByVal e As System.EventArgs)


    '    Dim objCombo As System.Web.UI.WebControls.DropDownList = _
    '        DirectCast(sender, System.Web.UI.WebControls.DropDownList)

    '    Try
    '        objCombo.Text = objCombo.Attributes("DefaultValue")

    '    Catch ex As Exception
    '        ' ???handle exception
    '    Finally
    '    End Try
    'End Sub

    'Public Sub DDEDataBound(ByVal sender As Object, ByVal e As System.EventArgs)



    '    Dim objCombo As AjaxControlToolkit.DropDownExtender = _
    '                DirectCast(sender, AjaxControlToolkit.DropDownExtender)

    '    Try
    '        ''objCombo.Text = objCombo.Attributes("DefaultValue")


    '    Catch ex As Exception
    '        ' ???handle exception
    '    Finally
    '    End Try
    'End Sub

    Private Function LookupFilterSQL(ByVal psColumnName As String, ByVal piColumnDataType As Integer, ByVal piOperatorID As Integer, ByVal psValue As String) As String
        Dim sLookupFilterSQL As String = ""

        Try

            If (psColumnName.Length > 0) _
                And (piOperatorID > 0) _
                And (psValue.Length > 0) Then

                Select Case piColumnDataType
                    Case SQLDataType.sqlBoolean
                        Select Case piOperatorID
                            Case FilterOperators.giFILTEROP_EQUALS
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab
                            Case FilterOperators.giFILTEROP_NOTEQUALTO
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab
                        End Select

                    Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
                        Select Case piOperatorID
                            Case FilterOperators.giFILTEROP_EQUALS
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab

                            Case FilterOperators.giFILTEROP_NOTEQUALTO
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab

                            Case FilterOperators.giFILTEROP_ISATMOST
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <= " & vbTab

                            Case FilterOperators.giFILTEROP_ISATLEAST
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) >= " & vbTab

                            Case FilterOperators.giFILTEROP_ISMORETHAN
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) > " & vbTab

                            Case FilterOperators.giFILTEROP_ISLESSTHAN
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) < " & vbTab
                        End Select

                    Case SQLDataType.sqlDate
                        Select Case piOperatorID
                            Case FilterOperators.giFILTEROP_ON
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

                            Case FilterOperators.giFILTEROP_NOTON
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

                            Case FilterOperators.giFILTEROP_ONORBEFORE
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN(ISNULL([ASRSysLookupFilterValue], '')) = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') <= '" & vbTab & "')"

                            Case FilterOperators.giFILTEROP_ONORAFTER
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') >= '" & vbTab & "')"

                            Case FilterOperators.giFILTEROP_AFTER
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "(LEN('" & vbTab & "') = 0 AND LEN(ISNULL([ASRSysLookupFilterValue], '')) > 0) OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') > '" & vbTab & "')"

                            Case FilterOperators.giFILTEROP_BEFORE
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') < '" & vbTab & "'"
                        End Select

                    Case SQLDataType.sqlVarChar, SQLDataType.sqlVarBinary, SQLDataType.sqlLongVarChar
                        Select Case piOperatorID
                            Case FilterOperators.giFILTEROP_IS
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

                            Case FilterOperators.giFILTEROP_ISNOT
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

                            Case FilterOperators.giFILTEROP_CONTAINS
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') LIKE '%" & vbTab & "%'"

                            Case FilterOperators.giFILTEROP_DOESNOTCONTAIN
                                sLookupFilterSQL = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') NOT LIKE '%" & vbTab & "%'"
                        End Select
                    Case Else
                End Select
            End If

        Catch ex As Exception
        End Try


        LookupFilterSQL = sLookupFilterSQL

    End Function

    'Sub GridRowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
    '    ' As each row is added to the grid's HTML table, do the following: 
    '    'check the item being bound is actually a DataRow, if it is, 
    '    'wire up the required html events and attach the relevant JavaScripts

    '    Dim iIDCol As Integer = 0
    '    Dim sRowID As String = "0"
    '    Dim grdGrid As System.Web.UI.WebControls.GridView
    '    Dim mydte As DateTime
    '    Dim sColumnCaption As String

    '    grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

    '    If e.Row.RowType = DataControlRowType.DataRow Then

    '        e.Row.Style.Add("overflow", "hidden")
    '        e.Row.Style.Add("cursor", "pointer")



    '        ' loop through the columns of this row. Hide ID columns
    '        For iColCount As Integer = 0 To e.Row.Cells.Count - 1
    '            sColumnCaption = UCase(grdGrid.HeaderRow.Cells(iColCount).Text)

    '            If sColumnCaption = "ID" Or (Left(sColumnCaption, 3) = "ID_" And Val(Mid(sColumnCaption, 4)) > 0) Then
    '                iIDCol = iColCount  ' store ID column number to assign to the javascript click event.
    '                e.Row.Cells(iColCount).Style.Add("display", "none")
    '                If sColumnCaption = "ID" Then
    '                    ' Background colour to black.
    '                    ' Javascript can see this and use it to recognise the ID column. Couldn't think of 
    '                    ' a better way, but sure there is one.
    '                    e.Row.Cells(iColCount).BackColor = Color.Black
    '                End If
    '            End If

    '            ' add overflow hidden and nowrap, stops the cells wrapping text or overflowing into adjacent cols.
    '            e.Row.Cells(iColCount).Style.Add("overflow", "hidden")
    '            e.Row.Cells(iColCount).Style.Add("white-space", "nowrap")

    '            ' this sets minimum width, not max.
    '            e.Row.Cells(iColCount).Width = Unit.Pixel(100)



    '            ' Format the cells according to DataType
    '            Dim curSelDataType As String = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

    '            If curSelDataType.Contains("INT") _
    '                Or curSelDataType.Contains("DECIMAL") _
    '                Or curSelDataType.Contains("SINGLE") _
    '                Or curSelDataType.Contains("DOUBLE") _
    '                Then curSelDataType = "Integer"
    '            If curSelDataType.Contains("DATETIME") Then curSelDataType = "DateTime"
    '            If curSelDataType.Contains("BOOLEAN") Then curSelDataType = "Boolean"

    '            Try
    '                Select Case curSelDataType
    '                    Case "DateTime"
    '                        ' Is the cell a date? 
    '                        mydte = DateTime.Parse(e.Row.Cells(iColCount).Text.ToString())
    '                        e.Row.Cells(iColCount).Text = mydte.ToShortDateString()
    '                    Case "Integer"
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "right")
    '                    Case "Boolean"
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "center")
    '                    Case Else   ' String
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "left")
    '                End Select
    '            Catch
    '                ' um...
    '            End Try
    '        Next

    '        ' Add some javascript events to each row for highlight functionality
    '        ' Not required, wasn't current functionality, but here you go if it should ever be 'needed'...
    '        'e.Row.Attributes("onmouseover") = "oldgridSelectedColor = this.style.backgroundColor;this.style.backgroundColor = '" & msBackColorHighlight.Trim & "';"
    '        'e.Row.Attributes("onmouseout") = "this.style.backgroundColor = oldgridSelectedColor;this.style.textDecoration = 'none';"
    '        'e.Row.Style.Add("cursor", "pointer")

    '        ' Add the javascript event to each row for click functionality
    '        If grdGrid.ID.Contains("Grid") Then
    '            e.Row.Attributes.Add("onclick", "changeRow('" & grdGrid.ID.ToString & "', '" & m_iRowIdx.ToString & "', '" & msBackColorHighlight.Trim & "', '" & iIDCol.ToString & "');oldgridSelectedColor = this.style.backgroundColor;")
    '        Else    ' dropdown?
    '            'e.Row.Attributes.Add("onclick", "changeDDERow('" & g.ID.ToString & "', '" & m_iRowIdx.ToString & "', 'TextBox1');")
    '        End If

    '    ElseIf e.Row.RowType = DataControlRowType.Pager Then
    '        ' This enables postback for the grid

    '        Dim pagerTable As Table = DirectCast(e.Row.Cells(0).Controls(0), Table)
    '        Dim pagerRow As TableRow = pagerTable.Rows(0)

    '        ' Dim pagerSettings As PagerSettings = DirectCast(sender, GridView).PagerSettings
    '        ' Dim cellsCount As Integer = pagerRow.Cells.Count

    '        For iColCount As Integer = 0 To pagerRow.Cells.Count - 1
    '            pagerRow.Cells(iColCount).Attributes.Add("onclick", "try{setPostbackMode(2);}catch(e){};")
    '        Next


    '    End If

    '    m_iRowIdx += 1

    'End Sub


    'Sub DDE_GridRowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)

    '    ' For Lookup Grids only - DropdownExtender

    '    ' As each row is added to the grid's HTML table, do the following: 
    '    'check the item being bound is actually a DataRow, if it is, 
    '    'wire up the required html events and attach the relevant JavaScripts

    '    Dim iIDCol As Integer = 0
    '    Dim sRowID As String = "0"
    '    Dim grdGrid As System.Web.UI.WebControls.GridView
    '    Dim mydte As DateTime
    '    Dim sColumnCaption As String

    '    grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

    '    If e.Row.RowType = DataControlRowType.DataRow Then

    '        e.Row.Style.Add("overflow", "hidden")
    '        e.Row.Style.Add("cursor", "pointer")

    '        ' loop through the columns of this row. Hide ID columns
    '        For iColCount As Integer = 0 To e.Row.Cells.Count - 1
    '            sColumnCaption = UCase(grdGrid.HeaderRow.Cells(iColCount).Text)

    '            If sColumnCaption.StartsWith("ASRSYS") Then
    '                iIDCol = iColCount  ' store ID column number to assign to the javascript click event.
    '                e.Row.Cells(iColCount).Style.Add("display", "none")
    '                If sColumnCaption.StartsWith("ASRSYS") Then
    '                    ' Background colour to black.
    '                    ' Javascript can see this and use it to recognise the ID column. Couldn't think of 
    '                    ' a better way, but sure there is one.
    '                    e.Row.Cells(iColCount).BackColor = Color.Black
    '                End If
    '            End If

    '            ' add overflow hidden and nowrap, stops the cells wrapping text or overflowing into adjacent cols.
    '            e.Row.Cells(iColCount).Style.Add("overflow", "hidden")
    '            e.Row.Cells(iColCount).Style.Add("white-space", "nowrap")

    '            ' this sets minimum width, not max.
    '            e.Row.Cells(iColCount).Width = Unit.Pixel(100)

    '            ' Format the cells according to DataType
    '            Dim curSelDataType As String = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

    '            If curSelDataType.Contains("INT") _
    '                Or curSelDataType.Contains("DECIMAL") _
    '                Or curSelDataType.Contains("SINGLE") _
    '                Or curSelDataType.Contains("DOUBLE") _
    '                Then curSelDataType = "Integer"
    '            If curSelDataType.Contains("DATETIME") Then curSelDataType = "DateTime"
    '            If curSelDataType.Contains("BOOLEAN") Then curSelDataType = "Boolean"

    '            Try
    '                Select Case curSelDataType
    '                    Case "DateTime"
    '                        ' Is the cell a date? 
    '                        mydte = DateTime.Parse(e.Row.Cells(iColCount).Text.ToString())
    '                        e.Row.Cells(iColCount).Text = mydte.ToShortDateString()
    '                    Case "Integer"
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "right")
    '                    Case "Boolean"
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "center")
    '                    Case Else   ' String
    '                        e.Row.Cells(iColCount).Style.Add("text-align", "left")
    '                End Select
    '            Catch
    '                ' um...
    '            End Try
    '        Next

    '        ' Add some javascript events to each row for highlight functionality
    '        ' Not required, wasn't current functionality, but here you go if it should ever be 'needed'...
    '        'e.Row.Attributes("onmouseover") = "oldgridSelectedColor = this.style.backgroundColor;this.style.backgroundColor = '" & msBackColorHighlight.Trim & "';"
    '        'e.Row.Attributes("onmouseout") = "this.style.backgroundColor = oldgridSelectedColor;this.style.textDecoration = 'none';"
    '        'e.Row.Style.Add("cursor", "pointer")

    '        ' Add the javascript event to each row for click functionality

    '        e.Row.Attributes.Add("onclick", "changeDDERow('" & grdGrid.ID.ToString & "', '" & m_iRowIdx.ToString & "', '" & msBackColorHighlight.Trim & "', '" & m_iLookupColumnIndex & "');oldgridSelectedColor = this.style.backgroundColor;")


    '    ElseIf e.Row.RowType = DataControlRowType.Header Then

    '        ' Get the lookupfiltervalue column number, if applicable and store to a tag.
    '        For iColCount As Integer = 0 To e.Row.Cells.Count - 1
    '            sColumnCaption = UCase(e.Row.Cells(iColCount).Text)

    '            If sColumnCaption.ToUpper = "ASRSYSLOOKUPFILTERVALUE" Then
    '                grdGrid.Attributes.Remove("LookupFilterColumn")
    '                grdGrid.Attributes.Add("LookupFilterColumn", iColCount.ToString)
    '            End If
    '        Next

    '    ElseIf e.Row.RowType = DataControlRowType.Pager Then
    '        ' This enables postback for the grid

    '        Dim pagerTable As Table = DirectCast(e.Row.Cells(0).Controls(0), Table)
    '        Dim pagerRow As TableRow = pagerTable.Rows(0)

    '        ' Dim pagerSettings As PagerSettings = DirectCast(sender, GridView).PagerSettings
    '        ' Dim cellsCount As Integer = pagerRow.Cells.Count

    '        For iColCount As Integer = 0 To pagerRow.Cells.Count - 1
    '            pagerRow.Cells(iColCount).Attributes.Add("onclick", "try{setPostbackMode(2);}catch(e){};")
    '        Next

    '    End If
    '    m_iRowIdx += 1
    'End Sub

    'Protected Sub PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs)

    '    Dim g As System.Web.UI.WebControls.GridView
    '    Dim iIDCol As Integer = 0
    '    g = CType(sender, System.Web.UI.WebControls.GridView)

    '    g.PageIndex = e.NewPageIndex
    '    g.DataBind()
    'End Sub


    '    <System.Web.Services.WebMethod()> _
    'Public Shared Function SetGridFilter(ByVal sGridID As String, ByVal sGridFilter As String) As Boolean
    '        'On server side we can do anything, like we can access the Session.
    '        'We can do database access operation. Without postback.

    '        Dim dataTable As DataTable
    '        dataTable = TryCast(HttpContext.Current.Session(sGridID.Replace("DDE", "DATA")), DataTable)

    '        If dataTable IsNot Nothing Then
    '            Dim dataView As New DataView(dataTable)
    '            dataView.RowFilter = sGridFilter   ' e.g. "ISNULL([ASRSysLookupFilterValue], '') = 'HERTFORDSHIRE'"
    '            dataTable = dataView.ToTable()
    '            HttpContext.Current.Session(sGridID.Replace("DDE", "DATA")) = dataTable
    '        End If

    '        Try
    '            Return True
    '        Catch ex As Exception
    '            Throw ex
    '        End Try

    '    End Function




    Private Sub ShowNoResultFound(ByVal source As DataTable, ByVal gv As RecordSelector)

        source.Clear()
        source.Rows.Add(source.NewRow())
        '' create a new blank row to the DataTable
        '' Bind the DataTable which contain a blank row to the GridView
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
        'gv.Rows(0).Cells(0).ForeColor = System.Drawing.Color.Red
        'gv.Rows(0).Cells(0).Font.Bold = True
        'set No Results found to the new added cell
        gv.Rows(0).Cells(0).Text = gv.EmptyDataText

        gv.SelectedIndex = -1


    End Sub

    Sub SetLookupFilter(ByVal sender As Object, ByVal e As System.EventArgs)

        ' get button's ID
        Dim btnSender As Button
        btnSender = DirectCast(sender, Button)

        ' Create a datatable from the data in the session variable
        Dim dataTable As DataTable
        dataTable = TryCast(HttpContext.Current.Session(btnSender.ID.Replace("refresh", "DATA")), DataTable)

        ' get the filter sql
        Dim hiddenField As HiddenField
        hiddenField = TryCast(pnlInputDiv.FindControl(btnSender.ID.Replace("refresh", "filterSQL")), HiddenField)

        Dim filterSQL As String = hiddenField.Value

        If dataTable IsNot Nothing Then
            Dim dataView As New DataView(dataTable)
            dataView.RowFilter = filterSQL    '   "ISNULL([ASRSysLookupFilterValue], '') = 'HERTFORDSHIRE'"

            dataTable = dataView.ToTable()
            ' HttpContext.Current.Session(btnSender.ID.Replace("refresh", "DATA")) = dataTable

            If dataTable.Rows.Count = 0 Then
                ' create a blank row to display.
                Dim objDataRow As System.Data.DataRow
                objDataRow = dataTable.NewRow()
                dataTable.Rows.InsertAt(objDataRow, 0)
            End If


        End If


        Dim gridView As RecordSelector 'GridView
        gridView = TryCast(pnlInputDiv.FindControl(btnSender.ID.Replace("refresh", "Grid")), RecordSelector)

        gridView.filterSQL = filterSQL.ToString

        gridView.DataSource = dataTable
        gridView.DataBind()


        ' reset filter.
        hiddenField.Value = ""


    End Sub

End Class
