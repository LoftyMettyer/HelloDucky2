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
  Private miCompletionMessageType As Integer
  Private msCompletionMessage As String
  Private miSavedForLaterMessageType As Integer
  Private msSavedForLaterMessage As String
  Private miFollowOnFormsMessageType As Integer
  Private msFollowOnFormsMessage As String
  Private miSubmissionTimeoutInSeconds As Int32
  Private m_iLookupColumnIndex As Integer
  Private iPageNo As Integer = 0
  Private _autoFocusControl As String

  Private Const FORMINPUTPREFIX As String = "FI_"
  Private Const ASSEMBLYNAME As String = "OPENHRWORKFLOW"
  Private Const DEFAULTTITLE As String = "OpenHR Workflow"
  Private Const miTabStripHeight As Integer = 21

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

    ScriptManager.GetCurrent(Page).AsyncPostBackTimeout = mobjConfig.SubmissionTimeout

  End Sub

#End Region

  Private Sub Page_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load

    Dim ctlForm_Dropdown As DropDownList
    Dim ctlForm_Image As WebControls.Image
    Dim ctlForm_PagingGridView As RecordSelector

    Dim ctlForm_PageTab() As Panel
    Dim sAssemblyName As String
    Dim sWebSiteVersion As String
    Dim sMessage As String
    Dim sQueryString As String
    Dim objCrypt As New Crypt
    Dim conn As SqlConnection
    Dim cmdCheck As SqlCommand
    Dim cmdSelect As SqlCommand
    Dim cmdInitiate As SqlCommand
    Dim cmdActivate As SqlCommand
    Dim dr As SqlDataReader
    Dim sTemp As String = String.Empty
    Dim sDBVersion As String
    Dim sID As String
    Dim connGrid As SqlConnection
    Dim drGrid As SqlDataReader
    Dim cmdGrid As SqlCommand
    Dim cmdQS As SqlCommand
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
    Dim sEncodedID As String
    Dim sFilterSQL As String
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim objDataRow As DataRow
    Dim iItemType As Integer
    Dim iCurrentPageTab As Integer

    ' MOBILE - start
    Dim sKeyParameter As String = ""
    Dim sPwdParameter As String = ""
    ' MOBILE - end

    sAssemblyName = ""
    sWebSiteVersion = ""
    sMessage = ""
    sQueryString = ""
    miImageCount = 0
    ReDim arrQueryStrings(0)
    sSiblingForms = ""

    Try
      mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

      miSubmissionTimeoutInSeconds = mobjConfig.SubmissionTimeoutInSeconds

      Response.CacheControl = "no-cache"
      Response.AddHeader("Pragma", "no-cache")
      Response.Expires = -1

      'HRPRO-2197 removed session clearing
      'If Not IsPostBack And Not IsMobileBrowser() Then
      '  Session.Clear()
      'End If
    Catch ex As Exception
    End Try

    Dim sTitle As String
    Try
      sAssemblyName = Assembly.GetExecutingAssembly.GetName.Name.ToUpper

      sWebSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

      If sAssemblyName = ASSEMBLYNAME Then
        ' Compiled version of the web site, so perform version checks.
        If sWebSiteVersion.Length = 0 Then
          sTitle = DEFAULTTITLE & " (unknown version)"
        Else
          sTitle = DEFAULTTITLE & " - v" & sWebSiteVersion
        End If
      Else
        ' Development version of the web site, so do NOT perform version checks.
        sTitle = DEFAULTTITLE & " (development)"
      End If
    Catch ex As Exception
      sTitle = DEFAULTTITLE
    End Try
    Page.Title = sTitle

    Try
      Dim cultureString As String

      If Request.UserLanguages IsNot Nothing Then
        cultureString = Request.UserLanguages(0)
      ElseIf Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") IsNot Nothing Then
        cultureString = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
      Else
        cultureString = ConfigurationManager.AppSettings("defaultculture")
      End If

      If cultureString.ToLower = "en-us" Then cultureString = "en-GB"

      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(cultureString)
      Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(cultureString)

    Catch ex As Exception
      sMessage = "Error reading the client culture:<BR><BR>" & ex.Message
    End Try

    If sMessage.Length = 0 Then
      If IsPostBack Then

        miInstanceID = CInt(ViewState("InstanceID"))
        miElementID = CInt(ViewState("ElementID"))
        msUser = ViewState("User").ToString
        msPwd = ViewState("Pwd").ToString
        msServer = ViewState("Server").ToString
        msDatabase = ViewState("Database").ToString

      Else
        Try
          ' Read and decrypt the queryString.
          ' Use the rawURL rather than the QueryString itself, as some of the 
          ' encryption characters are ignored in the QueryString.
          miElementID = 0
          miInstanceID = 0

          ' NPG20120201 - Fault HRPRO-1828
          ' Request.RawUrl replaces symbols with % codes, e.g. $=%40.
          sTemp = Server.UrlDecode(Request.RawUrl.ToString)
          Dim iTemp As Integer = sTemp.IndexOf("?")

          If iTemp >= 0 Then
            sQueryString = sTemp.Substring(iTemp + 1)
          Else
            ' NPG20120326 Fault HRPRO-2128
            Response.Redirect("~/Account/Login.aspx", False)
            Return
          End If

          ' Try the newer encryption first
          Try
            ' Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
            Dim sCultureName As String
            sCultureName = Thread.CurrentThread.CurrentCulture.Name

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-GB")

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

            ' MOBILE - start
            sKeyParameter = ""
            sPwdParameter = ""

            'msDatabase = Mid(sTemp, InStr(sTemp, vbTab) + 1)
            If InStr(sTemp, vbTab) > 0 Then
              msDatabase = Left(sTemp, InStr(sTemp, vbTab) - 1)

              ' See if there are any extra parameters used for record identification
              Try
                sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                sKeyParameter = Left(sTemp, InStr(sTemp, vbTab) - 1)
                sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                sPwdParameter = Mid(sTemp, InStr(sTemp, vbTab) + 1)

              Catch ex As Exception
                sKeyParameter = ""
                sPwdParameter = ""
              End Try
            Else
              msDatabase = Mid(sTemp, InStr(sTemp, vbTab) + 1)
            End If
            ' MOBILE - end


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

    ' - Mobile START - 
    ' This bit is simply for activating Mobile Security.
    ' NPG20111215 - I've hijacked the miInstanceID and populated it with the 
    ' User ID that is to be activated.
    If (sMessage.Length = 0) _
     And (miElementID = -2) _
     And (miInstanceID > 0) _
     And (Not IsPostBack) Then
      Try ' conn creation 
        ' update tbsysMobile_Logins, and copy the 'newpassword' string to the 'password' field using 'userid' from miInstanceID
        ' Establish Connection
        Dim myConnection As New SqlConnection(GetConnectionString)
        myConnection.Open()

        cmdActivate = New SqlCommand
        cmdActivate.CommandText = "spASRSysMobileActivateUser"
        cmdActivate.Connection = myConnection
        cmdActivate.CommandType = CommandType.StoredProcedure

        cmdActivate.Parameters.Add("@piRecordID", SqlDbType.Int).Direction = ParameterDirection.Input
        cmdActivate.Parameters("@piRecordID").Value = NullSafeInteger(miInstanceID)

        cmdActivate.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdActivate.ExecuteNonQuery()

        sMessage = CStr(cmdActivate.Parameters("@psMessage").Value())

        cmdActivate.Dispose()
        ' set message to something to skip all the normal workflow stuff.
        If sMessage.Length = 0 Then
          sMessage = "You have been successfully activated"
        End If

      Catch ex As Exception
        sMessage = "Unable to activate user."
      End Try
      ' done. smessage populated, so should skip the rest of default.aspx.
    End If
    ' - Mobile END -

    If sMessage.Length = 0 Then
      Try
        conn = New SqlConnection(GetConnectionString)
        conn.Open()
        Try
          If (sMessage.Length = 0) And (Not IsPostBack) Then

            ' Check if the database is locked.
            cmdCheck = New SqlCommand
            cmdCheck.CommandText = "sp_ASRLockCheck"
            cmdCheck.Connection = conn
            cmdCheck.CommandType = CommandType.StoredProcedure
            cmdCheck.CommandTimeout = miSubmissionTimeoutInSeconds

            dr = cmdCheck.ExecuteReader()

            While dr.Read
              If NullSafeInteger(dr("priority")) <> 3 Then
                ' Not a read-only lock.
                sMessage = "Database locked.<BR><BR>Contact your system administrator."
                Exit While
              End If
            End While

            dr.Close()
            cmdCheck.Dispose()
          End If

          If sMessage.Length = 0 And Not IsPostBack Then

            ' Check if the database and website versions match.
            cmdCheck = New SqlCommand
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
                sWebSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major & _
                 "." & Assembly.GetExecutingAssembly.GetName.Version.Minor
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

                sMessage = "The Workflow website version (" & sWebSiteVersion & ")" & " is incompatible with the database version (" & sDBVersion & ")." & "<BR><BR>Contact your system administrator."
              End If
            End If

            cmdCheck.Dispose()
          End If

          If (sMessage.Length = 0) And (miInstanceID < 0) And (miElementID = -1) And (Not IsPostBack) Then

            ' Externally initiated Workflow.
            iWorkflowID = -miInstanceID

            cmdInitiate = New SqlCommand

            ' MOBILE - start
            If Len(sKeyParameter) > 0 Then
              'sPWDParameter
              cmdInitiate.CommandText = "spASRMobileInstantiateWorkflow"
            Else
              cmdInitiate.CommandText = "spASRInstantiateWorkflow"
            End If
            ' MOBILE - end

            cmdInitiate.Connection = conn
            cmdInitiate.CommandType = CommandType.StoredProcedure
            cmdInitiate.CommandTimeout = miSubmissionTimeoutInSeconds

            cmdInitiate.Parameters.Add("@piWorkflowID", SqlDbType.Int).Direction = ParameterDirection.Input
            cmdInitiate.Parameters("@piWorkflowID").Value = iWorkflowID

            ' MOBILE - start
            If Len(sKeyParameter) > 0 Then
              cmdInitiate.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
              cmdInitiate.Parameters("@psKeyParameter").Value = sKeyParameter

              cmdInitiate.Parameters.Add("@psPWDParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
              cmdInitiate.Parameters("@psPWDParameter").Value = sPwdParameter
            End If
            ' MOBILE - end

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
                    cmdQS = New SqlCommand("spASRGetWorkflowQueryString", conn)
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

            ViewState("InstanceID") = miInstanceID
            ViewState("ElementID") = miElementID
            ViewState("User") = msUser
            ViewState("Pwd") = msPwd
            ViewState("Server") = msServer
            ViewState("Database") = msDatabase

            'FileUpload.apsx and FileDownload.aspx require these variables
            Session("User") = msUser
            Session("Pwd") = msPwd
            Session("Server") = msServer
            Session("Database") = msDatabase
            Session("ElementID") = miElementID
            Session("InstanceID") = miInstanceID

            ' Get the selected tab number for this workflow, if any...
            If Not IsPostBack Then
              Try
                cmdSelect = New SqlCommand("SELECT [pageno] FROM [dbo].[ASRSysWorkflowInstances] WHERE [ID] = " & NullSafeInteger(miInstanceID).ToString, conn)
                dr = cmdSelect.ExecuteReader()

                While dr.Read()
                  ' store the tab
                  iPageNo = NullSafeInteger(dr("pageno"))
                End While

                dr.Close()
                cmdSelect.Dispose()

              Catch ex As Exception
                iPageNo = 0
              End Try

              hdnDefaultPageNo.Value = iPageNo.ToString
            End If

            cmdSelect = New SqlCommand
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

            ReDim Preserve ctlForm_PageTab(0)
            While (dr.Read) And (sMessage.Length = 0)

              iCurrentPageTab = NullSafeInteger(dr("pageno"))

              ' Create the tab for this control. Do this first in case the tabstrip control hasn't been read yet,
              ' and the tabs haven't been generated.
              Try
                Dim strTemp As String = ctlForm_PageTab(iCurrentPageTab).ID.ToString
                ' OK, if the id exists, the div has already been created. Do nothing.
              Catch ex As Exception
                ' Otherwise create the div
                ' Create the new div, give it a unique id then we can refer to that when it's reused in the next loop.
                ' store the id in the array for reference. NB 21 is the itemtype for a page Tab
                If iCurrentPageTab > ctlForm_PageTab.GetUpperBound(0) Then ReDim Preserve ctlForm_PageTab(iCurrentPageTab)

                ctlForm_PageTab(iCurrentPageTab) = New Panel
                ctlForm_PageTab(iCurrentPageTab).ID = FORMINPUTPREFIX & iCurrentPageTab.ToString & "_21_PageTab"
                ctlForm_PageTab(iCurrentPageTab).Style.Add("position", "absolute")

                ' Add this tab to the web form
                pnlInputDiv.Controls.Add(ctlForm_PageTab(iCurrentPageTab))
              End Try

              ' Generate the unique ID for this control and process it onto the form.
              sID = FORMINPUTPREFIX & NullSafeString(dr("id")) & "_" & NullSafeString(dr("ItemType")) & "_"
              sEncodedID = objCrypt.SimpleEncrypt(NullSafeString(dr("id")).ToString, Session.SessionID)

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

                    If NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso NullSafeInteger(dr("BackColor")) <> -2147483633 Then
                      .Style.Add("background-color", General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

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

                    '.Style("word-wrap") = "break-word"
                    ' NPG20120305 Fault HRPRO-1967 reverted by PBG20120419 Fault HRPRO-2157
                    .Style("overflow") = "auto"

                    If NullSafeInteger(dr("captionType")) = 3 Then
                      ' Calculated caption
                      .Text = NullSafeString(dr("value"))
                    Else
                      .Text = NullSafeString(dr("caption"))
                    End If
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

                Case 3 ' Input value - character
                  Dim control = New TextBox
                  With control
                    .ID = sID
                    .TabIndex = NullSafeShort(dr("tabIndex"))
                    UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                    .ApplyLocation(dr)
                    .ApplySize(dr, -1, -1)
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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

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
                          sTemp = String.Empty
                        Else
                          sTemp = NullSafeString(dr("value")).Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                        End If

                        If sTemp.Length > 0 AndAlso sTemp.Chars(0) = "-" Then
                          sTemp = sTemp.Substring(1) & "-"
                        End If
                        .Text = sTemp

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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

                Case 5 ' Input value - numeric

                  Dim control = New TextBox
                  With control
                    .ID = sID
                    .CssClass = "numeric"

                    .TabIndex = NullSafeShort(dr("tabIndex"))
                    UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                    .ApplyLocation(dr)
                    .ApplySize(dr, -1, -1)
                    .Style.ApplyFont(dr)
                    .ApplyColor(dr, True)
                    .ApplyBorder(True)
                    .Style("padding") = "1px"

                    'add attributes that denote the min & max values, number of decimal places is also implied
                    Dim max = New String("9"c, NullSafeInteger(dr("inputSize")) - NullSafeInteger(dr("inputDecimals"))) & _
                      If(NullSafeInteger(dr("inputDecimals")) > 0, "." & New String("9"c, NullSafeInteger(dr("inputDecimals"))), "")

                    .Attributes.Add("data-numeric", String.Format("{{vMin: '-{0}', vMax: '{0}'}}", max))

                    'set the control value
                    Dim value As Single
                    If NullSafeString(dr("value")) <> "" Then
                      value = CSng(NullSafeString(dr("value")).Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator))
                    End If
                    .Text = value.ToString("N" & NullSafeInteger(dr("inputDecimals"))).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator, "")

                    .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                    If IsMobileBrowser() Then
                      .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                    End If

                  End With
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(checkBox)

                Case 7 ' Input value - date

                  If GetBrowserFamily() = "IOS" Then
                    ' Use the built in date barrel control.
                    ' HTML 5 only, and even then some browsers don't work properly. Yes YOU, android!
                    Dim control = New HtmlInputText
                    Dim hdnValue As String = ""

                    With control
                      .ID = sID
                      .Style.ApplyLocation(dr)
                      .Style.ApplySize(dr, -10, -3)
                      .Style.ApplyFont(dr)
                      .Style.ApplyColor(dr)

                      .Attributes.Add("TabIndex", NullSafeInteger(dr("tabIndex")).ToString)
                      UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                      .Attributes.Add("type", "date")
                      .Attributes.Add("onblur", "document.getElementById('" & sID & "Value').value = this.value;")

                      If Not IsPostBack Then

                        If (Not IsDBNull(dr("value"))) Then
                          If CStr(dr("value")).Length > 0 Then
                            Dim sDateString As String

                            Dim iYear = CShort(NullSafeString(dr("value")).Substring(6, 4))
                            sDateString = iYear.ToString & "-"

                            Dim iMonth = CShort(NullSafeString(dr("value")).Substring(0, 2))
                            If iMonth < 10 Then
                              sDateString &= "0" & iMonth.ToString & "-"
                            Else
                              sDateString &= iMonth.ToString & "-"
                            End If

                            Dim iDay = CShort(NullSafeString(dr("value")).Substring(3, 2))
                            If iDay < 10 Then
                              sDateString &= "0" & iDay.ToString
                            Else
                              sDateString &= iDay.ToString
                            End If

                            hdnValue = sDateString
                            .Value = hdnValue

                          End If
                        End If
                      Else
                        ' retrieve value from hidden field
                        .Value = Request.Form(sID & "Value").ToString
                      End If

                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

                    ' Yippee, can't find a way of storing the value to a server visible variable. 
                    ' So, use a hidden value.
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = sID & "Value", .Value = hdnValue})

                  Else
                    'TODO merge with IOS date above, once AjaxToolkit removed cos it cant postback input[type=date], .ApplySize(dr, -10, -3) for IOS, no .ApplyBorder(), date must be yyyy-mm-dd, remove the get value code from ButtonClick
                    Dim control = New TextBox
                    With control
                      .ID = sID
                      .CssClass = "date"

                      .TabIndex = NullSafeShort(dr("tabIndex"))
                      UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                      .ApplySize(dr, -1, -1)
                      .Style.ApplyFont(dr)
                      .ApplyColor(dr, True)
                      .ApplyBorder(True)

                      .Text = General.ConvertSqlDateToLocale(NullSafeString(dr("value")))

                      .Attributes("onfocus") = "try{" & sID & ".select();}catch(e){};"

                      If IsMobileBrowser() Then
                        .ReadOnly = True
                        .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                      End If
                    End With

                    Dim panel As New Panel
                    panel.Controls.Add(control)
                    panel.ApplyLocation(dr)

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(panel)
                  End If

                Case 8 ' Frame

                  Dim top = NullSafeInteger(dr("TopCoord"))
                  Dim left = NullSafeInteger(dr("LeftCoord"))
                  Dim width = NullSafeInteger(dr("Width"))
                  Dim height = NullSafeInteger(dr("Height"))
                  Dim fontAdjustment = CInt(CInt(dr("FontSize")) * 0.8)

                  width -= 2
                  height -= 2

                  If NullSafeString(dr("caption")).Trim.Length = 0 Then
                    top += fontAdjustment
                    height -= fontAdjustment
                  End If

                  sTemp = "<fieldset style='" & _
                 " position: absolute;" & _
                 " top: " & top & "px;" & _
                 " left: " & left & "px;" & _
                 " width: " & width & "px;" & _
                 " height: " & height & "px;" & _
                 " " & GetFontCss(dr) & _
                 " " & GetColorCss(dr) & _
                 " border: 1px solid #999;" & _
                 " '>"

                  If NullSafeString(dr("caption")).Trim.Length > 0 Then
                    sTemp += String.Format("<legend>{0}</legend>", NullSafeString(dr("caption"))) & vbCrLf
                  End If

                  sTemp += "</fieldset>" & vbCrLf

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New LiteralControl(sTemp))

                Case 9 ' Line
                  Select Case NullSafeInteger(dr("Orientation"))
                    Case 0
                      ' Vertical
                      sTemp = "<div style='position: absolute;" & _
                       " left: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                       " top: " & NullSafeString(dr("TopCoord")) & "px;" & _
                       " height: " & NullSafeString(dr("Height")) & "px;" & _
                       " width: 0px;" & _
                       " border-left: 1px solid " & General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) & ";'" & _
                       "></div>"
                    Case 1
                      ' Horizontal
                      sTemp = "<img style='position: absolute;" & _
                       " left: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                       " top: " & NullSafeString(dr("TopCoord")) & "px;" & _
                       " height: 0px;" & _
                       " width: " & NullSafeString(dr("Width")) & "px;" & _
                       " border-top: 1px solid " & General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) & ";'" & _
                       "></div>"
                  End Select

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New LiteralControl(sTemp))

                Case 10 ' Image
                  ctlForm_Image = New WebControls.Image

                  With ctlForm_Image
                    .ApplyLocation(dr)
                    .ApplySize(dr)

                    If NullSafeBoolean(dr("PictureBorder")) Then
                      .ApplyBorder(True, -2)
                    End If

                    Dim imageUrl As String = LoadPicture(NullSafeInteger(dr("pictureID")), sMessage)
                    If sMessage.Length > 0 Then
                      Exit While
                    End If
                    .ImageUrl = imageUrl
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Image)

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
                      .PageSize = Math.Min(CInt(Math.Truncate((CInt(NullSafeInteger(dr("Height")) - 42) / piRowHeight))), mobjConfig.LookupRowsRange)
                      .RowStyle.Height = Unit.Pixel(CInt(piRowHeight))
                    Else
                      .PageSize = mobjConfig.LookupRowsRange
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
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_PagingGridView)


                  If (Not IsPostBack) Then
                    connGrid = New SqlConnection(GetConnectionString)
                    connGrid.Open()

                    Try
                      cmdGrid = New SqlCommand
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
                        ctlForm_PagingGridView.IsEmpty = False
                        ctlForm_PagingGridView.DataSource = dt
                        ctlForm_PagingGridView.DataBind()
                      Else
                        ctlForm_PagingGridView.IsEmpty = True
                        ShowNoResultFound(dt, ctlForm_PagingGridView)
                      End If

                      ' ------------------------------------------------
                      ' Set default/first row
                      ' ------------------------------------------------
                      If ctlForm_PagingGridView.Rows.Count > 0 Then
                        If CStr(dr("value")).Length > 0 And CStr(dr("value")) <> "0" Then
                          Dim iIndexColumnNumber As Integer = dt.Columns.IndexOf("ID")
                          Dim iRowNumber As Long = 0

                          For Each rRow As DataRow In dt.Rows
                            If rRow.Item(iIndexColumnNumber).ToString = CStr(dr("value")) Then

                              ' set selected page index
                              Dim iCurrentPage As Long = iRowNumber \ ctlForm_PagingGridView.PageSize
                              ctlForm_PagingGridView.PageIndex = CInt(iCurrentPage)

                              ' set row number
                              Dim iCurrentRow As Long = iRowNumber Mod ctlForm_PagingGridView.PageSize
                              ctlForm_PagingGridView.SelectedIndex = CInt(iCurrentRow)

                              ctlForm_PagingGridView.DataBind()
                              Exit For

                            End If

                            iRowNumber += 1
                          Next
                        Else
                          ' set top row as default item
                          ctlForm_PagingGridView.SelectedIndex = 0
                        End If
                      End If

                      Dim recordOk = CBool(cmdGrid.Parameters("@pfOK").Value)
                      If Not recordOk Then
                        sMessage = "Error loading web form. Web Form record selector item record has been deleted or not selected."
                        Exit While
                      End If

                      cmdGrid.Dispose()

                    Catch ex As Exception
                      sMessage = "Error loading web form grid values:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
                      Exit While

                    Finally
                      connGrid.Close()
                      connGrid.Dispose()
                    End Try
                  Else
                    ' If a postback, check for empty datagrid and set empty row message
                    Dim dtSource As DataTable = TryCast(HttpContext.Current.Session(sID & "DATA"), DataTable)

                    If ctlForm_PagingGridView.IsEmpty Then
                      ShowNoResultFound(dtSource, ctlForm_PagingGridView)
                    End If
                  End If

                  ' ============================================================
                  ' Hidden field is used to store scroll position of the grid.
                  ' ============================================================
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})


                Case 14 ' lookup  Inputs
                  If Not IsMobileBrowser() Then

                    ' ============================================================
                    ' Create a textbox as the main control
                    ' ============================================================
                    Dim textBox = New TextBox

                    With textBox
                      .ID = sID & "TextBox"
                      .ApplyLocation(dr)
                      .ApplySize(dr, -1, -1)
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

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(textBox)

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

                    sFilterSQL = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")), _
                            NullSafeInteger(dr("lookupFilterColumnDataType")), _
                            NullSafeInteger(dr("LookupFilterOperator")), _
                            FORMINPUTPREFIX & NullSafeString(dr("lookupFilterValueID")) & "_" & NullSafeString(dr("lookupFilterValueType")) & "_")


                    ' ==========================================================
                    ' Hidden Field to store any lookup filter code
                    ' ==========================================================
                    If (sFilterSQL.Length > 0) Then
                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = sFilterSQL})
                    End If

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_PagingGridView)


                    If (Not IsPostBack) Then
                      connGrid = New SqlConnection(GetConnectionString)
                      connGrid.Open()
                      Try
                        cmdGrid = New SqlCommand
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

                        ' Create a blank row at the top of the dropdown grid.
                        objDataRow = dt.NewRow()
                        dt.Rows.InsertAt(objDataRow, 0)

                        m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)

                        iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                        textBox.Attributes.Remove("LookupColumnIndex")
                        textBox.Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                        textBox.Attributes.Remove("DefaultValue")
                        textBox.Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                        textBox.Attributes.Remove("DataType")
                        textBox.Attributes.Add("DataType", NullSafeString(dt.Columns(CInt(textBox.Attributes("LookupColumnIndex"))).DataType.ToString))

                        ctlForm_PagingGridView.DataSource = dt
                        ctlForm_PagingGridView.DataBind()

                        ctlForm_PagingGridView.IsEmpty = (dt.Rows.Count - 1 = 0)

                        cmdGrid.Dispose()

                      Catch ex As Exception

                        sMessage = "Error loading lookup values:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
                        Exit While

                      Finally
                        connGrid.Close()
                        connGrid.Dispose()
                      End Try

                      ' ==================================================
                      ' Set the dropdownList to the default value.
                      ' ==================================================
                      If textBox.Attributes("DefaultValue").ToString.Length > 0 Then
                        textBox.Text = textBox.Attributes("DefaultValue").ToString
                      End If

                      For jncount As Integer = 0 To ctlForm_PagingGridView.Rows.Count - 1
                        If jncount > ctlForm_PagingGridView.PageSize Then Exit For ' don't bother if on other pages
                        If ctlForm_PagingGridView.Rows(jncount).Cells(m_iLookupColumnIndex).Text = textBox.Attributes("DefaultValue").ToString Then
                          ctlForm_PagingGridView.SelectedIndex = jncount
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
                      If (sFilterSQL.Length > 0) Then
                        .OnClientPopup = "InitializeLookup"     ' can't pass the ID of the control, so use ._id in JS.
                      End If
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(dde)

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
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = sID & "scrollpos"})

                    ' ====================================================
                    ' hidden field to hold any filter SQL code
                    ' ====================================================
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

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

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(button)

                  Else
                    ' ================================================================================================================
                    ' Mobile Browser - convert lookup data to a standard dropdown.
                    ' ================================================================================================================
                    ctlForm_Dropdown = New DropDownList

                    With ctlForm_Dropdown
                      .ID = sID
                      .ApplyLocation(dr)
                      .ApplySize(dr, -1, -1)
                      .Style.ApplyFont(dr)
                      .ApplyColor(dr)
                      If Not IsMobileBrowser() Then .ApplyBorder(False)
                      .Style.Add("padding", "1px")

                      .TabIndex = NullSafeShort(dr("tabIndex"))
                      UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                      .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")

                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Dropdown)

                      sFilterSQL = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")), _
                              NullSafeInteger(dr("lookupFilterColumnDataType")), _
                              NullSafeInteger(dr("LookupFilterOperator")), _
                              FORMINPUTPREFIX & NullSafeString(dr("lookupFilterValueID")) & "_" & NullSafeString(dr("lookupFilterValueType")) & "_")

                      If (sFilterSQL.Length > 0) Then
                        ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = sFilterSQL})
                      End If

                      If (Not IsPostBack) Then
                        connGrid = New SqlConnection(GetConnectionString)
                        connGrid.Open()

                        Try

                          cmdGrid = New SqlCommand
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
                          Session(sID & "DATA") = dt

                          ctlForm_Dropdown.DataSource = dt

                          m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)
                          iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                          If dt.Columns(m_iLookupColumnIndex).DataType Is GetType(DateTime) Then
                            .DataTextFormatString = "{0:d}"
                          End If
                          .DataTextField = dt.Columns(m_iLookupColumnIndex).ColumnName.ToString

                          .Attributes.Remove("LookupColumnIndex")
                          .Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                          .Attributes.Remove("DefaultValue")
                          .Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                          ctlForm_Dropdown.DataBind()

                          cmdGrid.Dispose()

                        Catch ex As Exception
                          sMessage = "Error loading lookup values:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
                          Exit While

                        Finally
                          connGrid.Close()
                          connGrid.Dispose()
                        End Try

                        ' ==================================================
                        ' Set the dropdownList to the default value.
                        ' ==================================================

                        Dim listItem As ListItem = ctlForm_Dropdown.Items.FindByValue(ctlForm_Dropdown.Attributes("DefaultValue").ToString)
                        If listItem IsNot Nothing Then
                          ctlForm_Dropdown.SelectedValue = listItem.Value
                        Else
                          'The selected value is not in the list, so add it after the blank row
                          ctlForm_Dropdown.Items.Insert(1, ctlForm_Dropdown.Attributes("DefaultValue").ToString)
                          ctlForm_Dropdown.SelectedIndex = 1
                        End If
                      End If

                    End With

                    ' ====================================================
                    ' hidden field to hold any filter SQL code
                    ' ====================================================
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = sID & "filterSQL"})

                    ' ============================================================
                    ' Hidden Button for JS to call which fires filter click event. 
                    ' ============================================================
                    Dim button = New Button
                    With button
                      .ID = sID & "refresh"
                      .Style.Add("display", "none")
                    End With

                    AddHandler button.Click, AddressOf SetLookupFilter

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(button)

                  End If

                Case 13 ' Dropdown (13) Inputs

                  ctlForm_Dropdown = New DropDownList

                  With ctlForm_Dropdown
                    .ID = sID
                    .ApplyLocation(dr)
                    .ApplySize(dr, -1, -1)
                    .Style.ApplyFont(dr)
                    .ApplyColor(dr)
                    If Not IsMobileBrowser() Then .ApplyBorder(False)
                    .Style.Add("padding", "1px")

                    .TabIndex = NullSafeShort(dr("tabIndex"))
                    UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                    If IsMobileBrowser() Then
                      .Attributes.Add("onchange", "FilterMobileLookup('" & .ID & "');")
                    End If

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Dropdown)

                    sFilterSQL = LookupFilterSQL(NullSafeString(dr("lookupFilterColumnName")), _
                            NullSafeInteger(dr("lookupFilterColumnDataType")), _
                            NullSafeInteger(dr("LookupFilterOperator")), _
                            FORMINPUTPREFIX & NullSafeString(dr("lookupFilterValueID")) & "_" & NullSafeString(dr("lookupFilterValueType")) & "_")

                    If sFilterSQL.Length > 0 Then
                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = "lookup" & sID, .Value = sFilterSQL})
                    End If

                    If (Not IsPostBack) Then
                      connGrid = New SqlConnection(GetConnectionString)
                      connGrid.Open()

                      Try

                        cmdGrid = New SqlCommand
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

                        ' Format the column(s)
                        For Each column As DataColumn In dt.Columns
                          If Not column.ColumnName.StartsWith("ASRSys") Then
                            .DataTextField = column.ColumnName.ToString
                          End If
                        Next

                        ctlForm_Dropdown.DataSource = dt

                        m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)
                        iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                        .Attributes.Remove("LookupColumnIndex")
                        .Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                        .Attributes.Remove("DefaultValue")
                        .Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                        ctlForm_Dropdown.DataBind()

                        cmdGrid.Dispose()

                      Catch ex As Exception
                        sMessage = "Error loading lookup values:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
                        Exit While

                      Finally
                        connGrid.Close()
                        connGrid.Dispose()
                      End Try

                      ' ==================================================
                      ' Set the dropdownList to the default value.
                      ' ==================================================

                      Dim listItem As ListItem = ctlForm_Dropdown.Items.FindByValue(ctlForm_Dropdown.Attributes("DefaultValue").ToString)
                      If listItem IsNot Nothing Then
                        ctlForm_Dropdown.SelectedValue = listItem.Value
                      End If

                    End If

                  End With

                Case 15 ' OptionGroup

                  Dim top = NullSafeInteger(dr("TopCoord"))
                  Dim left = NullSafeInteger(dr("LeftCoord"))
                  Dim width = NullSafeInteger(dr("Width"))
                  Dim height = NullSafeInteger(dr("Height"))
                  Dim fontAdjustment = CInt(CInt(dr("FontSize")) * 0.8)
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

                    radioTop = 19 + CInt((NullSafeInteger(dr("FontSize")) - 8) * 1.375)

                    If IsAndroidBrowser() AndAlso NullSafeInteger(dr("Orientation")) = 0 Then
                      radioTop -= 5
                    End If
                  End If

                  sTemp = "<fieldset style='" & _
                   " position: absolute; " & _
                   " top: " & top & "px; " & _
                   " left: " & left & "px; " & _
                   " width: " & width & "px; " & _
                   " height: " & height & "px; " & _
                   " " & GetFontCss(dr) & _
                   " " & GetColorCss(dr) & _
                   " " & borderCss & _
                   " '>"

                  If NullSafeBoolean(dr("PictureBorder")) And (NullSafeString(dr("caption")).Trim.Length > 0) Then
                    sTemp += String.Format("<legend>{0}</legend>", NullSafeString(dr("caption"))) & vbCrLf
                  End If

                  sTemp += "</fieldset>" & vbCrLf

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New LiteralControl(sTemp))

                  Dim radioList As New RadioButtonList
                  With radioList
                    .ID = sID
                    .Style.ApplyFont(dr)
                    .CssClass = "radioList"
                    If IsAndroidBrowser() Then .CssClass += " android"

                    .TabIndex = NullSafeShort(dr("tabIndex"))
                    UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID & "_0")

                    .RepeatDirection = If(NullSafeInteger(dr("Orientation")) = 0, RepeatDirection.Vertical, RepeatDirection.Horizontal)

                    .Style("position") = "absolute"
                    .Style("top") = Unit.Pixel(radioTop + NullSafeInteger(dr("TopCoord"))).ToString
                    .Style("left") = Unit.Pixel(9 + NullSafeInteger(dr("LeftCoord"))).ToString
                    .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 12)
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(radioList)

                  If Not IsPostBack Then

                    connGrid = New SqlConnection(GetConnectionString)
                    connGrid.Open()
                    Try
                      cmdGrid = New SqlCommand
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

                      While drGrid.Read
                        radioList.Items.Add(New ListItem() With { _
                                            .Text = drGrid(0).ToString, _
                                            .Value = drGrid(0).ToString, _
                                            .Selected = (CInt(drGrid.GetValue(1)) = 1) _
                                          })
                      End While

                      If radioList.SelectedIndex = -1 Then
                        radioList.SelectedIndex = 0
                      End If

                      drGrid.Close()
                      cmdGrid.Dispose()

                    Catch ex As Exception
                      sMessage = "Error loading web form option group values:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
                      Exit While

                    Finally
                      connGrid.Close()
                      connGrid.Dispose()
                    End Try

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

                    If NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso NullSafeInteger(dr("BackColor")) <> -2147483633 Then
                      .Style.Add("background-color", General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
                      .Style.Add("border", "solid 1px #CCC")
                      .Style.Add("border-radius", "4px")
                    End If

                    If NullSafeInteger(dr("ForeColor")) <> 6697779 Then
                      .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))).ToString)
                    End If

                    .Style.Add("padding", "0px")
                    .Style.Add("white-space", "normal")

                    .Value = NullSafeString(dr("caption"))

                    If Not IsMobileBrowser() Then
                      .Attributes.Add("onclick", "try{showFileUpload(true, '" & sEncodedID & "', document.getElementById('file" & sID & "').value);}catch(e){};")
                    Else
                      .Attributes.Add("onclick", "try{alert('Your browser does not support file upload.');}catch(e){};")
                    End If
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(control)

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New HiddenField With {.ID = "file" & sID, .Value = NullSafeString(dr("value"))})

                Case 19, 20 ' DB File or WF File

                  sTemp = "<span id='" & sID & "' tabindex=" & NullSafeInteger(dr("tabIndex")).ToString & _
                   " style='position: absolute; display:inline-block; word-wrap:break-word; overflow:auto;" & _
                   " top: " & NullSafeString(dr("TopCoord")) & "px;" & _
                   " left: " & NullSafeString(dr("LeftCoord")) & "px;" & _
                   " height:" & NullSafeString(dr("Height")) & "px;" & _
                   " width:" & NullSafeInteger(dr("Width")) & "px;" & _
                   " " & GetFontCss(dr) & _
                   " " & GetColorCss(dr) & _
                   "'" & _
                   " onclick='FileDownload_Click(""" & sEncodedID & """);'" & _
                   " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" & _
                   " >" & _
                   HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) & _
                   "</span>"

                  UpdateAutoFocusControl(NullSafeShort(dr("tabIndex")), sID)

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(New LiteralControl(sTemp))

                Case 21   ' Tab Strip

                  'split out the tab names to calculate number of tabs - may not have loaded all tabs yet, so can't count them.
                  sTemp = NullSafeString(dr("Caption"))   '"Page 1;Page 2;"
                  Dim arrTabCaptions As String() = sTemp.Split(New Char() {";"c})

                  pnlTabsDiv.Style("width") = CStr(dr("Width")) & "px"
                  pnlTabsDiv.Style("height") = CStr(dr("Height")) & "px"
                  pnlTabsDiv.Style("left") = CStr(dr("LeftCoord")) & "px"
                  pnlTabsDiv.Style("top") = CStr(dr("TopCoord")) & "px"

                  Dim ctlTabsDiv As New Panel
                  ctlTabsDiv.ID = "TabsDiv"
                  ctlTabsDiv.Style.Add("height", miTabStripHeight & "px")
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
                    ctlForm_Image = New WebControls.Image
                    With ctlForm_Image
                      .Style.Add("width", "24px")
                      .Style.Add("height", miTabStripHeight - 2 & "px")
                      .ImageUrl = "~/Images/tab-prev.gif"
                      .Style.Add("margin", "0px")
                      .Style.Add("padding", "0px")
                      .Attributes.Add("onclick", "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft - 20;")
                    End With
                    ctlFormTabArrows.Controls.Add(ctlForm_Image)

                    ' Right scroll arrow
                    ctlForm_Image = New WebControls.Image
                    With ctlForm_Image
                      .Style.Add("width", "24px")
                      .Style.Add("height", miTabStripHeight - 2 & "px")
                      .ImageUrl = "~/Images/tab-next.gif"
                      .Style.Add("margin", "0px")
                      .Style.Add("padding", "0px")
                      .Attributes.Add("onclick", "var TabDiv = document.getElementById('TabsDiv');TabDiv.scrollLeft = TabDiv.scrollLeft + 20;")
                    End With
                    ctlFormTabArrows.Controls.Add(ctlForm_Image)

                    pnlTabsDiv.Controls.Add(ctlFormTabArrows)

                  End If

                  ' generate the tabs.
                  Dim ctlTabsTable As New Table
                  ctlTabsTable.CellSpacing = 0
                  ' ctlTabsTable.Style.Add("margin-top", "2px")
                  Dim trPager As TableRow = New TableRow()
                  trPager.Height = Unit.Pixel(miTabStripHeight - 1) ' to prevent vertical scrollbar
                  trPager.Style.Add("white-space", "nowrap")

                  Dim iTabNo As Integer = 1
                  ' add a cell for each tab
                  For Each sTabCaption In arrTabCaptions
                    If sTabCaption.Trim.Length > 0 Then
                      Dim tcTabCell As TableCell = New TableCell

                      With tcTabCell
                        .ID = FORMINPUTPREFIX & iTabNo.ToString & "_21_Panel"
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

                      If iTabNo > 0 Then  ' Tab 0 is the base page.

                        ' create any MISSING tabs...
                        Try
                          Dim strTemp As String = ctlForm_PageTab(iTabNo).ID.ToString
                          ' OK, if the id exists, the div has already been created. Do nothing.
                        Catch ex As Exception
                          ' Otherwise create the div
                          ' Create the new div, give it a unique id then we can refer to that when it's reused in the next loop.
                          ' store the id in the array for reference. NB 21 is the itemtype for a page Tab
                          If iTabNo > ctlForm_PageTab.GetUpperBound(0) Then ReDim Preserve ctlForm_PageTab(iTabNo)

                          ctlForm_PageTab(iTabNo) = New Panel
                          ctlForm_PageTab(iTabNo).ID = FORMINPUTPREFIX & iTabNo.ToString & "_21_PageTab"
                          ctlForm_PageTab(iTabNo).Style.Add("position", "absolute")

                          ' Add this tab to the web form
                          pnlInputDiv.Controls.Add(ctlForm_PageTab(iTabNo))
                        End Try

                        ' Move all tabs to their relative position within the tab frame.
                        Try
                          ctlForm_PageTab(iTabNo).Style.Add("top", NullSafeInteger(dr("TopCoord")) + miTabStripHeight & "px")
                          ctlForm_PageTab(iTabNo).Style.Add("left", NullSafeInteger(dr("LeftCoord")) & "px")

                          ' Hide all tabs but the first.
                          ctlForm_PageTab(iTabNo).Style.Add("display", "none")
                        Catch ex As Exception

                        End Try
                      End If

                      iTabNo += 1 ' keep tabs on the number of tabs hehehe :P
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

            If (Not ClientScript.IsStartupScriptRegistered("Startup")) Then
              ' Form the script to be registered at client side.
              scriptString += "}"
              ClientScript.RegisterStartupScript(ClientScript.GetType, "Startup", scriptString, True)
            End If

            If sMessage.Length = 0 Then
              If CStr(cmdSelect.Parameters("@psErrorMessage").Value) <> "" Then
                sMessage = CStr(cmdSelect.Parameters("@psErrorMessage").Value)
              Else

                If CInt(cmdSelect.Parameters("@piBackImage").Value) > 0 Then
                  Dim image As String = LoadPicture(CInt(cmdSelect.Parameters("@piBackImage").Value), sMessage)
                  If sMessage.Length = 0 Then
                    divInput.Style("background-image") = image
                    divInput.Style("background-repeat") = General.BackgroundRepeat(CShort(cmdSelect.Parameters("@piBackImageLocation").Value))
                    divInput.Style("background-position") = General.BackgroundPosition(CShort(cmdSelect.Parameters("@piBackImageLocation").Value))
                  End If
                End If

                If Not IsDBNull(cmdSelect.Parameters("@piBackColour").Value) Then
                  divInput.Style("background-color") = General.GetHtmlColour(CInt(cmdSelect.Parameters("@piBackColour").Value))
                End If

                iFormWidth = CInt(cmdSelect.Parameters("@piWidth").Value)
                iFormHeight = CInt(cmdSelect.Parameters("@piHeight").Value)

                pnlInputDiv.Style("width") = iFormWidth.ToString & "px"
                pnlInputDiv.Style("height") = iFormHeight.ToString & "px"
                pnlInputDiv.Style("left") = "-2px"

                hdnFormHeight.Value = iFormHeight.ToString
                hdnFormWidth.Value = iFormWidth.ToString

                hdnSiblingForms.Value = sSiblingForms.ToString

                miCompletionMessageType = NullSafeInteger(cmdSelect.Parameters("@piCompletionMessageType").Value)
                msCompletionMessage = NullSafeString(cmdSelect.Parameters("@psCompletionMessage").Value)
                miSavedForLaterMessageType = NullSafeInteger(cmdSelect.Parameters("@piSavedForLaterMessageType").Value)
                msSavedForLaterMessage = NullSafeString(cmdSelect.Parameters("@psSavedForLaterMessage").Value)
                miFollowOnFormsMessageType = NullSafeInteger(cmdSelect.Parameters("@piFollowOnFormsMessageType").Value)
                msFollowOnFormsMessage = NullSafeString(cmdSelect.Parameters("@psFollowOnFormsMessage").Value)
              End If
            End If

            cmdSelect.Dispose()

          End If

          ' Resize the mobile 'viewport' to fit the webform
          AddHeaderTags(iFormWidth)

        Catch ex As Exception
          sMessage = "Error loading web form controls:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
        Finally
          conn.Close()
          conn.Dispose()
        End Try

      Catch ex As Exception   ' conn creation 
        sMessage = "Error creating SQL connection:<BR><BR>" & ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & "Contact your system administrator."
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

  Public Sub ButtonClick(ByVal sender As System.Object, ByVal e As EventArgs)

    Dim conn As SqlConnection
    Dim dr As SqlDataReader
    Dim cmdValidate As SqlCommand
    Dim cmdUpdate As SqlCommand
    Dim cmdQs As SqlCommand
    Dim valueString As String
    Dim ctlFormInput As Control
    Dim sID As String
    Dim sIDString As String
    Dim iTemp As Int16
    Dim sTemp As String
    Dim iType As Int16
    Dim sType As String
    Dim sFormElements As String
    Dim arrFollowOnForms() As String
    Dim fSavedForLater As Boolean
    Dim sMessage As String
    Dim iFollowOnFormCount As Integer
    Dim iIndex As Integer
    Dim sStep As String
    Dim arrQueryStrings() As String
    Dim sFollowOnForms As String
    Dim value As String

    sMessage = ""
    valueString = ""
    sFollowOnForms = ""
    ReDim arrQueryStrings(0)

    Try
      ' Read the web form item values & build up a string of the form input values.
      ' This is a tab delimited string of itemIDs and values.
      Dim controlList As New List(Of Control)
      GetControls(Page.Controls, controlList, Function(c) c.ClientID.StartsWith(FORMINPUTPREFIX) AndAlso _
                                                (c.ClientID.EndsWith("_") OrElse c.ClientID.EndsWith("TextBox") OrElse c.ClientID.EndsWith("Grid")))

      For Each ctlFormInput In controlList

        sID = ctlFormInput.ID
        sIDString = sID.Substring(Len(FORMINPUTPREFIX))
        iTemp = CShort(sIDString.IndexOf("_"))
        sTemp = sIDString.Substring(iTemp + 1)
        sIDString = sIDString.Substring(0, iTemp) & vbTab

        iTemp = CShort(sTemp.IndexOf("_"))
        sType = sTemp.Substring(0, iTemp)
        iType = CShort(sType)

        Select Case iType

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
              value = If(CSng(control.Text) = CSng(0), "0", control.Text.Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, "."))
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

            If TypeOf ctlFormInput Is HtmlInputText Then
              'an HTML5 compliant mobile device?
              Dim control = DirectCast(pnlInput.FindControl(sID & "Value"), HiddenField)
              value = If(control.Value = "", "null", Format(DateTime.Parse(control.Value), "MM/dd/yyyy"))
              valueString += sIDString & value & vbTab
            End If

          Case 11 ' Grid (RecordSelector) Input
            If TypeOf ctlFormInput Is RecordSelector Then
              Dim control = DirectCast(ctlFormInput, RecordSelector)

              value = "0"
              If Not control.IsEmpty And control.SelectedIndex >= 0 Then
                For iColCount As Integer = 0 To control.HeaderRow.Cells.Count - 1
                  If (control.HeaderRow.Cells(iColCount).Text.ToLower() = "id") Then
                    value = control.SelectedRow.Cells(iColCount).Text
                    Exit For
                  End If
                Next
              End If

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

                sTemp = control.Text

                If control.Attributes("DataType") = "System.DateTime" Then
                  If sTemp Is Nothing Then
                    sTemp = "null"
                  Else
                    If (sTemp.Length = 0) Then
                      sTemp = "null"
                    Else
                      sTemp = General.ConvertLocaleDateToSql(sTemp)
                    End If
                  End If
                ElseIf control.Attributes("DataType") = "System.Decimal" Or control.Attributes("DataType") = "System.Int32" Then

                  If sTemp Is Nothing Then
                    sTemp = ""
                  Else
                    sTemp = If(sTemp.Length = 0, "", sTemp.Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, "."))
                  End If

                End If

                valueString += sIDString & sTemp & vbTab
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
              value = DirectCast(pnlInput.FindControl("file" & sID), HiddenField).Value
              valueString += sIDString & value & vbTab
            End If

        End Select

      Next

    Catch ex As Exception
      sMessage = "Error reading web form item values:<BR><BR>" & ex.Message
    End Try

    If sMessage.Length = 0 Then
      Try
        conn = New SqlConnection(GetConnectionString)
        conn.Open()

        Try ' Validate the web form entry.
          errorMessagePanel.Font.Name = "Verdana"
          errorMessagePanel.Font.Size = mobjConfig.ValidationMessageFontSize
          errorMessagePanel.ForeColor = General.GetColour(6697779)

          bulletErrors.Items.Clear()
          bulletWarnings.Items.Clear()

          cmdValidate = New SqlCommand("spASRSysWorkflowWebFormValidation", conn)
          cmdValidate.CommandType = CommandType.StoredProcedure
          cmdValidate.CommandTimeout = miSubmissionTimeoutInSeconds

          cmdValidate.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
          cmdValidate.Parameters("@piInstanceID").Value = miInstanceID

          cmdValidate.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
          cmdValidate.Parameters("@piElementID").Value = miElementID

          cmdValidate.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
          cmdValidate.Parameters("@psFormInput1").Value = valueString

          dr = cmdValidate.ExecuteReader

          While dr.Read
            If NullSafeInteger(dr("failureType")) = 0 Then
              bulletErrors.Items.Add(NullSafeString(dr("Message")))
            ElseIf hdnOverrideWarnings.Value <> "1" Then
              bulletWarnings.Items.Add(NullSafeString(dr("Message")))
            End If
          End While

          dr.Close()
          cmdValidate.Dispose()

          hdnCount_Errors.Value = CStr(bulletErrors.Items.Count)
          hdnCount_Warnings.Value = CStr(bulletWarnings.Items.Count)
          hdnOverrideWarnings.Value = "0"

          lblErrors.Text = If(bulletErrors.Items.Count > 0, _
            "Unable to submit this form due to the following error" & _
            If(bulletErrors.Items.Count = 1, "", "s") & ":", _
            "")

          lblWarnings.Text = If(bulletWarnings.Items.Count > 0, _
            If(bulletErrors.Items.Count > 0, "And the following warning" & _
            If(bulletWarnings.Items.Count = 1, "", "s") & ":", "Submitting this form raises the following warning" & _
            If(bulletWarnings.Items.Count = 1, "", "s") & ":"), _
           "")

          overrideWarning.Visible = (bulletWarnings.Items.Count > 0 And bulletErrors.Items.Count = 0)

        Catch ex As Exception
          sMessage = "Error validating the web form:<BR><BR>" & ex.Message
        End Try

        ' Submit the webform
        If (sMessage.Length = 0) And (bulletWarnings.Items.Count = 0) And (bulletErrors.Items.Count = 0) Then

          Using (New TransactionScope(TransactionScopeOption.Suppress))
            Try
              ' Get the currently selected tab...
              iPageNo = NullSafeInteger(hdnDefaultPageNo.Value)

              cmdUpdate = New SqlCommand("spASRSubmitWorkflowStep", conn)
              cmdUpdate.CommandType = CommandType.StoredProcedure
              cmdUpdate.CommandTimeout = miSubmissionTimeoutInSeconds

              cmdUpdate.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
              cmdUpdate.Parameters("@piInstanceID").Value = miInstanceID

              cmdUpdate.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
              cmdUpdate.Parameters("@piElementID").Value = miElementID

              cmdUpdate.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
              cmdUpdate.Parameters("@psFormInput1").Value = valueString

              cmdUpdate.Parameters.Add("@psFormElements", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
              cmdUpdate.Parameters.Add("@pfSavedForLater", SqlDbType.Bit).Direction = ParameterDirection.Output

              cmdUpdate.Parameters.Add("@piPageNo", SqlDbType.Int).Direction = ParameterDirection.Input
              cmdUpdate.Parameters("@piPageNo").Value = iPageNo

              cmdUpdate.ExecuteNonQuery()

              sFormElements = CStr(cmdUpdate.Parameters("@psFormElements").Value())
              fSavedForLater = CBool(cmdUpdate.Parameters("@pfSavedForLater").Value())

              cmdUpdate.Dispose()

              If fSavedForLater Then
                Select Case miSavedForLaterMessageType
                  Case 1 ' Custom
                    If Not SetSubmissionMessage(msSavedForLaterMessage) Then
                      SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here", "to close this form.")
                    End If
                  Case 2 ' None
                    SetSubmissionMessage("", "", "")
                  Case Else   'System default
                    SetSubmissionMessage("Workflow step saved for later.<BR><BR>Click", "here", "to close this form.")
                End Select

              ElseIf sFormElements.Length = 0 Then
                Select Case miCompletionMessageType
                  Case 1 ' Custom
                    If Not SetSubmissionMessage(msCompletionMessage) Then
                      SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to close this form.")
                    End If
                  Case 2 ' None
                    SetSubmissionMessage("", "", "")
                  Case Else   'System default
                    SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to close this form.")
                End Select
              Else
                arrFollowOnForms = sFormElements.Split(CChar(vbTab))
                iFollowOnFormCount = arrFollowOnForms.GetUpperBound(0)

                For iIndex = 0 To iFollowOnFormCount - 1
                  sStep = arrFollowOnForms(iIndex)

                  cmdQs = New SqlCommand("spASRGetWorkflowQueryString", conn)
                  cmdQs.CommandType = CommandType.StoredProcedure
                  cmdQs.CommandTimeout = miSubmissionTimeoutInSeconds

                  cmdQs.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdQs.Parameters("@piInstanceID").Value = miInstanceID

                  cmdQs.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdQs.Parameters("@piElementID").Value = CLng(sStep)

                  cmdQs.Parameters.Add("@psQueryString", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                  cmdQs.ExecuteNonQuery()

                  Dim sQueryString As String = CStr(cmdQs.Parameters("@psQueryString").Value())

                  ReDim Preserve arrQueryStrings(arrQueryStrings.GetUpperBound(0) + 1)
                  arrQueryStrings(arrQueryStrings.GetUpperBound(0)) = sQueryString

                  cmdQs.Dispose()
                Next iIndex

                sFollowOnForms = Join(arrQueryStrings, vbTab)

                Select Case miFollowOnFormsMessageType
                  Case 1 ' Custom
                    If Not SetSubmissionMessage(msFollowOnFormsMessage) Then
                      SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to complete the follow-on Workflow form" & If(iFollowOnFormCount = 1, "", "s") & ".")
                    End If
                  Case 2 ' None
                    SetSubmissionMessage("", "", "")
                  Case Else   'System default
                    SetSubmissionMessage("Workflow step completed.<BR><BR>Click", "here", "to complete the follow-on Workflow form" & If(iFollowOnFormCount = 1, "", "s") & ".")
                End Select

              End If

              hdnFollowOnForms.Value = sFollowOnForms

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
      hdnFollowOnForms.Value = ""
      SetSubmissionMessage(sMessage & "<BR><BR>Click", "here", "to close this form.")
    End If

  End Sub

  Private _minTabIndex As Short = -1
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
    Return mobjConfig.ColourThemeHex
  End Function

  Private Function GetConnectionString() As String
    Dim connectionString = "Application Name=OpenHR Workflow;Data Source=" & msServer & ";Initial Catalog=" & msDatabase & ";Integrated Security=false;User ID=" & msUser & ";Password=" & msPwd & ";Pooling=false"
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
      miImageCount = CShort(miImageCount + 1)

      psErrorMessage = ""
      sImageFileName = ""
      sImageFilePath = Server.MapPath("pictures")

      conn = New SqlConnection(GetConnectionString)
      conn.Open()

      cmdSelect = New SqlCommand
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

  Private Function LookupFilterSQL(ByVal psColumnName As String, ByVal piColumnDataType As Integer, ByVal piOperatorID As Integer, ByVal psValue As String) As String

    Dim filterSql As String = ""

    Try
      If (psColumnName.Length > 0) And (piOperatorID > 0) And (psValue.Length > 0) Then

        Select Case piColumnDataType
          Case SQLDataType.sqlBoolean
            Select Case piOperatorID
              Case FilterOperators.giFILTEROP_EQUALS
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab
              Case FilterOperators.giFILTEROP_NOTEQUALTO
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab
            End Select

          Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
            Select Case piOperatorID
              Case FilterOperators.giFILTEROP_EQUALS
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) = " & vbTab

              Case FilterOperators.giFILTEROP_NOTEQUALTO
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <> " & vbTab

              Case FilterOperators.giFILTEROP_ISATMOST
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) <= " & vbTab

              Case FilterOperators.giFILTEROP_ISATLEAST
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) >= " & vbTab

              Case FilterOperators.giFILTEROP_ISMORETHAN
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) > " & vbTab

              Case FilterOperators.giFILTEROP_ISLESSTHAN
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], 0) < " & vbTab
            End Select

          Case SQLDataType.sqlDate
            Select Case piOperatorID
              Case FilterOperators.giFILTEROP_ON
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

              Case FilterOperators.giFILTEROP_NOTON
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

              Case FilterOperators.giFILTEROP_ONORBEFORE
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN(ISNULL([ASRSysLookupFilterValue], '')) = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') <= '" & vbTab & "')"

              Case FilterOperators.giFILTEROP_ONORAFTER
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') = 0 OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') >= '" & vbTab & "')"

              Case FilterOperators.giFILTEROP_AFTER
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "(LEN('" & vbTab & "') = 0 AND LEN(ISNULL([ASRSysLookupFilterValue], '')) > 0) OR (LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') > '" & vbTab & "')"

              Case FilterOperators.giFILTEROP_BEFORE
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') < '" & vbTab & "'"
            End Select

          Case SQLDataType.sqlVarChar, SQLDataType.sqlVarBinary, SQLDataType.sqlLongVarChar
            Select Case piOperatorID
              Case FilterOperators.giFILTEROP_IS
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') = '" & vbTab & "'"

              Case FilterOperators.giFILTEROP_ISNOT
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') <> '" & vbTab & "'"

              Case FilterOperators.giFILTEROP_CONTAINS
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "ISNULL([ASRSysLookupFilterValue], '') LIKE '%" & vbTab & "%'"

              Case FilterOperators.giFILTEROP_DOESNOTCONTAIN
                filterSql = piColumnDataType.ToString & vbTab & psValue & vbTab & "LEN('" & vbTab & "') > 0 AND ISNULL([ASRSysLookupFilterValue], '') NOT LIKE '%" & vbTab & "%'"
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

    gv.SelectedIndex = -1

  End Sub

  Protected Sub BtnDoFilterClick(sender As Object, e As EventArgs) Handles btnDoFilter.Click

    Dim arrLookups() As String = hdnMobileLookupFilter.Value.Split(CChar(vbTab))

    For Each value As String In arrLookups
      SetLookupFilter(Nothing, Nothing, value)
    Next
  End Sub

  Sub SetLookupFilter(ByVal sender As Object, ByVal e As EventArgs, Optional lookupID As String = "")

    If Not (sender Is Nothing) Then
      ' get button's ID
      Dim btnSender As Button
      btnSender = DirectCast(sender, Button)

      lookupID = btnSender.ID
    End If

    If lookupID.Length = 0 Then Return

    ' Create a datatable from the data in the session variable
    Dim dataTable As DataTable
    dataTable = TryCast(HttpContext.Current.Session(lookupID.Replace("refresh", "DATA")), DataTable)

    ' get the filter sql
    Dim hiddenField As HiddenField
    hiddenField = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "filterSQL")), HiddenField)

    Dim filterSql As String = hiddenField.Value

    If TypeOf (pnlInputDiv.FindControl(lookupID.Replace("refresh", ""))) Is DropDownList Then

      ' This is a dropdownlist style lookup (mobiles only)
      Dim dropdown As DropDownList
      dropdown = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "")), DropDownList)

      ' Store the current value, so we can re-add it after filtering.
      Dim strCurrentSelection As String = dropdown.Text

      ' Filter the table now.
      FilterDataTable(dataTable, filterSql)

      ' insert the previously selected item
      Dim objDataRow As DataRow
      objDataRow = dataTable.NewRow()
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

      Dim gridView As RecordSelector
      gridView = TryCast(pnlInputDiv.FindControl(lookupID.Replace("refresh", "Grid")), RecordSelector)

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
    meta.Content = (Session.Timeout * 60).ToString & "; URL=timeout.aspx"

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

End Class
