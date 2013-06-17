Option Strict On

Imports System
Imports System.Data
Imports System.Globalization
Imports System.Threading
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports Utilities
Imports System.Data.SqlClient
Imports System.Transactions
Imports System.Reflection
Imports System.Linq

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
  Private m_iLookupColumnIndex As Integer
  Private iPageNo As Integer = 0

  Private Const FORMINPUTPREFIX As String = "forminput_"
  Private Const ASSEMBLYNAME As String = "OPENHRWORKFLOW"
  Private Const MAXDROPDOWNROWS As Int16 = 6
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

    ScriptManager.GetCurrent(Page).AsyncPostBackTimeout = SubmissionTimeout()

  End Sub

#End Region

  Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    Dim ctlForm_Date As Infragistics.WebUI.WebSchedule.WebDateChooser
    Dim ctlForm_InputButton As Button
    Dim ctlForm_HTMLInputButton As HtmlInputButton
    Dim ctlForm_Label As Label
    Dim ctlForm_TextInput As TextBox
    Dim ctlForm_CheckBox As LiteralControl
    Dim ctlForm_CheckBoxReal As CheckBox
    Dim ctlForm_Dropdown As System.Web.UI.WebControls.DropDownList
    Dim ctlForm_Image As WebControls.Image
    Dim ctlForm_NumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Dim ctlForm_PagingGridView As RecordSelector
    Dim ctlForm_Frame As LiteralControl
    Dim ctlForm_Line As LiteralControl
    Dim ctlForm_OptionGroup As LiteralControl
    Dim ctlForm_OptionGroupReal As TextBox
    Dim ctlForm_HiddenField As HiddenField
    Dim ctlForm_Literal As LiteralControl
    Dim ctlForm_UpdatePanel As System.Web.UI.UpdatePanel
    Dim ctlForm_PageTab() As Panel
    Dim ctlForm_HTMLInputText As HtmlInputText
    Dim sBackgroundImage As String
    Dim sBackgroundRepeat As String
    Dim sBackgroundPosition As String
    Dim iBackgroundColour As Integer
    Dim sBackgroundColourHex As String
    Dim iBackgroundImagePosition As Integer
    Dim sAssemblyName As String
    Dim sWebSiteVersion As String
    Dim sMessage As String
    Dim sQueryString As String
    Dim objCrypt As New Crypt
    Dim blnLocked As Boolean
    Dim conn As SqlConnection
    Dim cmdCheck As SqlCommand
    Dim cmdSelect As SqlCommand
    Dim cmdInitiate As SqlCommand
    Dim cmdActivate As System.Data.SqlClient.SqlCommand
    Dim dr As SqlDataReader
    Dim iTemp As Integer
    Dim sTemp As String = String.Empty
    Dim sTemp2 As String
    Dim sDBVersion As String
    Dim sID As String
    Dim sImageFileName As String
    Dim sBackColour As String
    Dim objNumberFormatInfo As NumberFormatInfo
    Dim dtDate As Date
    Dim iYear As Int16
    Dim iMonth As Int16
    Dim iDay As Int16
    Dim objGridColumn As DataColumn
    Dim iHeaderHeight As Int32
    Dim iTempHeight As Int32
    Dim iTempWidth As Int32
    Dim connGrid As SqlConnection
    Dim drGrid As SqlDataReader
    Dim cmdGrid As SqlCommand
    Dim cmdQS As SqlCommand
    Dim iMinTabIndex As Integer
    Dim sDefaultValue As String
    Dim fRecordOK As Boolean
    Dim iGridTopPadding As Integer
    Dim iRowHeight As Integer
    Dim iDropHeight As Integer
    Dim sDefaultFocusControl As String
    Dim ctlDefaultFocusControl As New Control
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
    Dim objDataRow As DataRow
    Dim iItemType As Integer
    Dim iPageTabCount As Integer
    Dim iCurrentPageTab As Integer

    ' MOBILE - start
    Dim sKeyParameter As String = ""
    Dim sPWDParameter As String = ""
    ' MOBILE - end

    Const sDEFAULTTITLE As String = "OpenHR Workflow"
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
    iPageTabCount = 0

    Try
      mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

      miSubmissionTimeoutInSeconds = mobjConfig.SubmissionTimeoutInSeconds

      Response.CacheControl = "no-cache"
      Response.AddHeader("Pragma", "no-cache")
      Response.Expires = -1

      If Not IsPostBack And Not isMobileBrowser() Then
        Session.Clear()
      End If
    Catch ex As Exception
    End Try

    Try
      sAssemblyName = Assembly.GetExecutingAssembly.GetName.Name.ToUpper

      sWebSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

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
      Dim cultureString As String

      If Request.UserLanguages IsNot Nothing Then
        cultureString = Request.UserLanguages(0)
      ElseIf Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") IsNot Nothing Then
        cultureString = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
      Else
        cultureString = System.Configuration.ConfigurationManager.AppSettings("defaultculture")
      End If

      If cultureString.ToLower = "en-us" Then cultureString = "en-GB"

      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(cultureString)
      Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(cultureString)

    Catch ex As Exception
      sMessage = "Error reading the client culture:<BR><BR>" & ex.Message
    End Try

    If sMessage.Length = 0 Then
      If IsPostBack Then

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

          ' NPG20120201 - Fault HRPRO-1828
          ' Request.RawUrl replaces symbols with % codes, e.g. $=%40.
          sTemp = Server.UrlDecode(Request.RawUrl.ToString)
          iTemp = sTemp.IndexOf("?")

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
            sPWDParameter = ""

            'msDatabase = Mid(sTemp, InStr(sTemp, vbTab) + 1)
            If InStr(sTemp, vbTab) > 0 Then
              msDatabase = Left(sTemp, InStr(sTemp, vbTab) - 1)

              ' See if there are any extra parameters used for record identification
              Try
                sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                sKeyParameter = Left(sTemp, InStr(sTemp, vbTab) - 1)
                sTemp = Mid(sTemp, InStr(sTemp, vbTab) + 1)

                sPWDParameter = Mid(sTemp, InStr(sTemp, vbTab) + 1)

              Catch ex As Exception
                sKeyParameter = ""
                sPWDParameter = ""
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
        Dim myConnection As New SqlClient.SqlConnection(GetConnectionString)
        myConnection.Open()

        cmdActivate = New SqlClient.SqlCommand
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
          If (sMessage.Length = 0) _
           And (Not IsPostBack) Then

            ' Check if the database is locked.
            cmdCheck = New SqlCommand
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
              cmdInitiate.Parameters("@psPWDParameter").Value = sPWDParameter
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
                'If iCurrentPageTab = 0 Then
                pnlInputDiv.Controls.Add(ctlForm_PageTab(iCurrentPageTab))
                'Else
                'ctlForm_PageTab(iCurrentPageTab).Style.Add("display", "none")
                'pnlTabsDiv.Controls.Add(ctlForm_PageTab(iCurrentPageTab))
                'End If
              End Try



              ' Generate the unique ID for this control and process it onto the form.
              sID = FORMINPUTPREFIX & NullSafeString(dr("id")) & "_" & NullSafeString(dr("ItemType")) & "_"
              sEncodedID = objCrypt.SimpleEncrypt(NullSafeString(dr("id")).ToString, Session.SessionID)

              Select Case NullSafeInteger(dr("ItemType"))

                Case 0 ' Button
                  ctlForm_HTMLInputButton = New HtmlInputButton
                  With ctlForm_HTMLInputButton
                    .ID = sID
                    .Attributes.Add("TabIndex", CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString)

                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                      sDefaultFocusControl = sID
                      iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                    End If

                    .Style("position") = "absolute"
                    .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                    .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                    ' If the button has no caption, we treat as a transparent button.
                    ' This is so we can emulate picture buttons with very little code changes.
                    If NullSafeString(dr("caption")) = vbNullString Then
                      .Style.Add("filter", "alpha(opacity=0)")
                      .Style.Add("opacity", "0")
                    End If

                    ' stops the mobiles displaying buttons with over-rounded corners...
                    If IsMobileBrowser() Then
                      .Style.Add("-webkit-appearance", "none")
                      .Style.Add("background-color", "#CCCCCC")
                      .Style.Add("border", "solid 1px #C0C0C0")
                      .Style.Add("border-radius", "0px")
                    End If

                    If NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso NullSafeInteger(dr("BackColor")) <> -2147483633 Then
                      .Style.Add("background-color", General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
                      .Style.Add("border", "solid 1px " & General.GetHtmlColour(9999523).ToString)
                    End If

                    If NullSafeInteger(dr("ForeColor")) <> 6697779 Then
                      .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))).ToString)
                    End If

                    .Style.Add("padding", "0px")
                    .Style.Add("white-space", "normal")

                    .Value = NullSafeString(dr("caption"))

                    sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                       CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                    If sTemp2.Length = 0 Then
                      sTemp2 = " none"
                    End If

                    .Style.Add("Font-family", NullSafeString(dr("FontName")).ToString)
                    .Style.Add("Font-Size", ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt")
                    .Style.Add("Font-weight", CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")))
                    .Style.Add("FontStyle", CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")))
                    .Style.Add("Text-Decoration", sTemp2)

                    .Style.Add("Width", Unit.Pixel(NullSafeInteger(dr("Width"))).ToString)
                    .Style.Add("Height", Unit.Pixel(NullSafeInteger(dr("Height"))).ToString)

                    .Attributes.Add("onclick", "try{setPostbackMode(1);}catch(e){};")
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HTMLInputButton)

                  AddHandler ctlForm_HTMLInputButton.ServerClick, AddressOf Me.ButtonClick

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
                            .Text = General.ConvertSQLDateToLocale(NullSafeString(dr("value")))
                          End If
                      End Select

                      If NullSafeInteger(dr("BackStyle")) = 0 Then
                        .BackColor = Color.Transparent
                      Else
                        .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      End If

                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                      .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                      .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                      .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                      iTempHeight = NullSafeInteger(dr("Height"))
                      iTempWidth = NullSafeInteger(dr("Width"))

                      If NullSafeBoolean(dr("PictureBorder")) Then
                        .BorderStyle = BorderStyle.Solid
                        .BorderColor = General.GetColour(5730458)
                        .BorderWidth = Unit.Pixel(1)

                        iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                        iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                      End If

                      .Height() = Unit.Pixel(iTempHeight)
                      .Width() = Unit.Pixel(iTempWidth)

                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Label)

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
                        .BackColor = Color.Transparent
                      Else
                        .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      End If
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      iTempHeight = NullSafeInteger(dr("Height"))
                      iTempWidth = NullSafeInteger(dr("Width"))

                      If NullSafeBoolean(dr("PictureBorder")) Then
                        .BorderStyle = BorderStyle.Solid
                        .BorderColor = General.GetColour(5730458)
                        .BorderWidth = Unit.Pixel(1)

                        iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                        iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                      Else
                        .BorderStyle = BorderStyle.None
                      End If

                      .Height() = Unit.Pixel(iTempHeight)
                      .Width() = Unit.Pixel(iTempWidth)
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_TextInput)
                  End If

                Case 2 ' Label
                  ctlForm_Label = New Label
                  With ctlForm_Label
                    .Style("position") = "absolute"
                    .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                    .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                    .Style("word-wrap") = "break-word"
                    ' NPG20120305 Fault HRPRO-1967 reverted by PBG20120419 Fault HRPRO-2157
                    .Style("overflow") = "auto"
                    .Style("text-align") = "left"

                    If NullSafeInteger(dr("captionType")) = 3 Then
                      ' Calculated caption
                      .Text = NullSafeString(dr("value"))
                    Else
                      .Text = NullSafeString(dr("caption"))
                    End If

                    If NullSafeInteger(dr("BackStyle")) = 0 Then
                      .BackColor = Color.Transparent
                    Else
                      .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                    End If
                    .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                    .Font.Name = NullSafeString(dr("FontName"))
                    .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                    .Font.Bold = NullSafeBoolean(NullSafeBoolean(dr("FontBold")))
                    .Font.Italic = NullSafeBoolean(NullSafeBoolean(dr("FontItalic")))
                    .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(dr("FontStrikeThru")))
                    .Font.Underline = NullSafeBoolean(NullSafeBoolean(dr("FontUnderline")))

                    iTempHeight = NullSafeInteger(dr("Height"))
                    iTempWidth = NullSafeInteger(dr("Width"))

                    If NullSafeBoolean(dr("PictureBorder")) Then
                      .BorderStyle = BorderStyle.Solid
                      .BorderColor = General.GetColour(5730458)
                      .BorderWidth = Unit.Pixel(1)

                      iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                      iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                    End If

                    ' NPG20120305 Fault HRPRO-1967 reverted by PBG20120419 Fault HRPRO-2157
                    .Height() = Unit.Pixel(iTempHeight)
                    .Width() = Unit.Pixel(iTempWidth)
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Label)

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

                    If NullSafeBoolean(dr("PasswordType")) Then
                      .TextMode = TextBoxMode.Password
                    Else
                      .TextMode = TextBoxMode.MultiLine
                      .Wrap = True
                      .Style("overflow") = "auto"
                      .Style("word-wrap") = "break-word"
                      .Style("resize") = "none"
                    End If

                    .Text = NullSafeString(dr("value"))

                    .BorderStyle = BorderStyle.Solid
                    .BorderWidth = Unit.Pixel(1)
                    .BorderColor = General.GetColour(5730458)

                    .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                    .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                    .Font.Name = NullSafeString(dr("FontName"))
                    .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                    .Font.Bold = NullSafeBoolean(dr("FontBold"))
                    .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                    .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                    .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                    .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 6)
                    .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 6)

                    .Attributes("onfocus") = "try{" & sID & ".select();activateControl();}catch(e){};"
                    .Attributes("onkeydown") = "try{checkMaxLength(" & NullSafeString(dr("inputSize")) & ");}catch(e){}"
                    .Attributes("onpaste") = "try{checkMaxLength(" & NullSafeString(dr("inputSize")) & ");}catch(e){}"

                    If IsMobileBrowser() Then .Attributes.Add("onchange", "FilterMobileLookup('" & .ID.ToString & "');")

                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_TextInput)

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
                            .Text = General.ConvertSQLDateToLocale(NullSafeString(dr("value")))
                          End If
                      End Select

                      If NullSafeInteger(dr("BackStyle")) = 0 Then
                        .BackColor = Color.Transparent
                      Else
                        .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      End If
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      iTempHeight = NullSafeInteger(dr("Height"))
                      iTempWidth = NullSafeInteger(dr("Width"))

                      If NullSafeBoolean(dr("PictureBorder")) Then
                        .BorderStyle = BorderStyle.Solid
                        .BorderColor = General.GetColour(5730458)
                        .BorderWidth = Unit.Pixel(1)

                        iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                        iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                      End If

                      .Height() = Unit.Pixel(iTempHeight)
                      .Width() = Unit.Pixel(iTempWidth)

                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Label)
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
                        .BackColor = Color.Transparent
                      Else
                        .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      End If
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      iTempHeight = NullSafeInteger(dr("Height"))
                      iTempWidth = NullSafeInteger(dr("Width"))

                      If NullSafeBoolean(dr("PictureBorder")) Then
                        .BorderStyle = BorderStyle.Solid
                        .BorderColor = General.GetColour(5730458)
                        .BorderWidth = Unit.Pixel(1)

                        iTempHeight = iTempHeight - (2 * IMAGEBORDERWIDTH)
                        iTempWidth = iTempWidth - (2 * IMAGEBORDERWIDTH)
                      Else
                        .BorderStyle = BorderStyle.None
                      End If

                      .Height() = Unit.Pixel(iTempHeight)
                      .Width() = Unit.Pixel(iTempWidth)
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_TextInput)
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

                    .BorderColor = General.GetColour(5730458)
                    .BorderStyle = BorderStyle.Solid
                    .BorderWidth = Unit.Pixel(1)

                    .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                    .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                    .Font.Name = NullSafeString(dr("FontName"))
                    .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
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

                    If IsMobileBrowser() Then .ClientSideEvents.TextChanged = "FilterMobileLookup('" & .ID.ToString & "');"

                  End With

                  ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_NumericInput)
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_NumericInput)

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
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_CheckBoxReal)

                  msRefreshLiteralsCode = msRefreshLiteralsCode & vbNewLine & _
                   vbTab & vbTab & "try" & vbNewLine & _
                   vbTab & vbTab & "{" & vbNewLine & _
                   vbTab & vbTab & vbTab & "frmMain.chk" & sID & ".checked = frmMain." & sID & ".checked;" & vbNewLine & _
                   vbTab & vbTab & "}" & vbNewLine & _
                   vbTab & vbTab & "catch(e) {}"

                  If NullSafeInteger(dr("BackStyle")) = 0 Then
                    sBackColour = "Transparent"
                  Else
                    sBackColour = General.GetHtmlColour(NullSafeInteger(dr("BackColor")))
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
                   " COLOR: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                   " font-family: " & NullSafeString(dr("FontName")) & "; " & _
                   " font-size: " & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt; " & _
                   " font-weight: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                   " font-style: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                   " text-decoration:" & sTemp2 & "'>" & vbCrLf & _
                   "<TR style='background-color:red;'>" & vbCrLf

                  'android checkboxes need to be moved up
                  Dim androidFix As String = String.Empty

                  If IsAndroidBrowser() Then
                    androidFix = "position:relative;top:-4px;"
                  End If

                  If IsPostBack Then
                    If pnlInput.FindControl(sID) Is Nothing Then
                      fChecked = True
                    Else
                      ctlFormCheckBox = DirectCast(pnlInput.FindControl(sID), CheckBox)
                      fChecked = ctlFormCheckBox.Checked
                    End If

                    If NullSafeInteger(dr("alignment")) = 0 Then
                      sTemp = sTemp & _
                       "<TD><input type='checkbox'" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onclick=""" & sID & ".checked = checked;""" & _
                       CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", "")) & _
                       " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                       CStr(IIf(fChecked, " CHECKED", "")) & _
                       " style='height:14px;width:14px;margin:0px;" & androidFix & "'" & _
                       " tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                       " id='chk" & sID & "'" & _
                       " name='chk" & sID & "'></TD>" & vbCrLf & _
                       "</TD><TD width='100%'><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='-1'" & _
                       " style='padding-left: 3px;'" & _
                       " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf
                    Else
                      sTemp = sTemp & _
                       "<TD width='100%'><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                       " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf & _
                       "<TD><input type='checkbox'" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onclick=""" & sID & ".checked = checked;""" & _
                       CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", "")) & _
                       " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                       CStr(IIf(fChecked, " CHECKED", "")) & _
                       " style='height:14px;width:14px;margin:0px;" & androidFix & "'" & _
                       " tabIndex='-1'" & _
                       " id='chk" & sID & "'" & _
                       " name='chk" & sID & "'></TD>" & vbCrLf
                    End If
                  Else
                    If NullSafeInteger(dr("alignment")) = 0 Then
                      sTemp = sTemp & _
                       "<TD><input type='checkbox'" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onclick=""" & sID & ".checked = checked;""" & _
                       CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", "")) & _
                       " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                       CStr(IIf(UCase(NullSafeString(dr("value"))) = "TRUE", " CHECKED", "")) & _
                       " style='height:14px;width:14px;margin:0px;" & androidFix & "'" & _
                       " tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                       " id='chk" & sID & "'" & _
                       " name='chk" & sID & "'></TD>" & vbCrLf & _
                       "</TD><TD width='100%'><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='-1'" & _
                       " style='padding-left: 3px;'" & _
                       " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf
                    Else
                      sTemp = sTemp & _
                       "<TD width='100%'><LABEL ID='forChk" & sID & "' FOR='chk" & sID & "' tabIndex='" & NullSafeInteger(dr("tabIndex")) + 1 & "'" & _
                       " onkeypress = ""try{if(window.event.keyCode == 32){chk" & sID & ".click()};}catch(e){}""" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onfocus = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onblur = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       ">" & NullSafeString(dr("caption")) & "</LABEL></TD>" & vbCrLf & _
                       "<TD><input type='checkbox'" & _
                       " onmouseover = ""try{forChk" & sID & ".style.color='#ff9608'; }catch(e){};""" & _
                       " onmouseout = ""try{forChk" & sID & ".style.color='';}catch(e){};""" & _
                       " onclick=""" & sID & ".checked = checked;""" & _
                       CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", "")) & _
                       " onfocus=""try{" & sID & ".select();activateControl();}catch(e){};""" & _
                       CStr(IIf(NullSafeString(dr("value")).ToUpper = "TRUE", " CHECKED", "")) & _
                       " style='height:14px;width:14px;margin:0px;" & androidFix & "'" & _
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
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_CheckBox)

                  If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                    sDefaultFocusControl = "chk" & sID
                    iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                  End If

                Case 7 ' Input value - date
                  Dim HDNValue As String = ""

                  If getBrowserFamily() = "IOS" Then
                    ' Use the built in date barrel control.
                    ' HTML 5 only, and even then some browsers don't work properly. Yes YOU, android!
                    ctlForm_HTMLInputText = New HtmlInputText

                    With ctlForm_HTMLInputText
                      .ID = sID

                      .Attributes.Add("type", "date")
                      .Attributes.Add("TabIndex", CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString)
                      .Attributes.Add("onblur", "document.getElementById('" & sID & "Value').value = this.value;")

                      If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                        sDefaultFocusControl = sID
                        iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                        ctlDefaultFocusControl = ctlForm_HTMLInputText
                      End If

                      .Style("position") = "absolute"
                      .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                      .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                      .Style("margin") = "0px"
                      .Style("height") = Unit.Pixel(NullSafeInteger(dr("Height")) - 3).ToString
                      If Not IsPostBack Then

                        If (Not IsDBNull(dr("value"))) Then
                          If CStr(dr("value")).Length > 0 Then
                            Dim sDateString As String

                            iYear = CShort(NullSafeString(dr("value")).Substring(6, 4))
                            sDateString = iYear.ToString & "-"

                            iMonth = CShort(NullSafeString(dr("value")).Substring(0, 2))
                            If iMonth < 10 Then
                              sDateString &= "0" & iMonth.ToString & "-"
                            Else
                              sDateString &= iMonth.ToString & "-"
                            End If

                            iDay = CShort(NullSafeString(dr("value")).Substring(3, 2))
                            If iDay < 10 Then
                              sDateString &= "0" & iDay.ToString & "-"
                            Else
                              sDateString &= iDay.ToString
                            End If

                            HDNValue = sDateString
                            .Value = HDNValue

                          End If
                        End If
                      Else
                        ' retrieve value from hidden field
                        Dim tmpDateValue As String = Request.Form(sID & "Value").ToString
                        .Value = tmpDateValue
                      End If

                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HTMLInputText)

                    ' Yippee, can't find a way of storing the value to a server visible variable. 
                    ' So, use a hidden value.
                    ctlForm_HiddenField = New HiddenField

                    With ctlForm_HiddenField
                      .ID = sID & "Value"
                      .Value = HDNValue
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)

                  Else
                    ' Use the infragistics control.
                    ctlForm_Date = New Infragistics.WebUI.WebSchedule.WebDateChooser
                    With ctlForm_Date
                      .ID = sID
                      .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                      ' Mobiles sometimes show keyboards when dateboxes are clicked.
                      ' These can overlap the calendar control, so suppress it.
                      If IsMobileBrowser() Then
                        .Editable = False
                        .Attributes.Add("onclick", "showCalendar('" & .ClientID.ToString & "');")
                      End If

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

                      Dim fontUnit = New FontUnit(11, UnitType.Pixel)

                      .CalendarLayout.DayStyle.Font.Size = fontUnit
                      .CalendarLayout.DayStyle.Font.Name = "Verdana"
                      .CalendarLayout.DayStyle.ForeColor = General.GetColour(6697779)
                      .CalendarLayout.DayStyle.BackColor = General.GetColour(15988214)

                      .CalendarLayout.FooterStyle.Font.Size = fontUnit
                      .CalendarLayout.FooterStyle.Font.Name = "Verdana"
                      .CalendarLayout.FooterStyle.ForeColor = General.GetColour(6697779)
                      .CalendarLayout.FooterStyle.BackColor = General.GetColour(16248553)

                      .CalendarLayout.SelectedDayStyle.Font.Size = fontUnit
                      .CalendarLayout.SelectedDayStyle.Font.Name = "Verdana"
                      .CalendarLayout.SelectedDayStyle.Font.Bold = True
                      .CalendarLayout.SelectedDayStyle.Font.Underline = True
                      .CalendarLayout.SelectedDayStyle.ForeColor = General.GetColour(2774907)
                      .CalendarLayout.SelectedDayStyle.BackColor = General.GetColour(10480637)

                      .CalendarLayout.OtherMonthDayStyle.Font.Size = fontUnit
                      .CalendarLayout.OtherMonthDayStyle.Font.Name = "Verdana"
                      .CalendarLayout.OtherMonthDayStyle.ForeColor = General.GetColour(11375765)

                      .CalendarLayout.NextPrevStyle.ForeColor = SystemColors.InactiveCaptionText
                      .CalendarLayout.NextPrevStyle.BackColor = General.GetColour(16248553)
                      .CalendarLayout.NextPrevStyle.ForeColor = General.GetColour(6697779)

                      .CalendarLayout.CalendarStyle.Width = Unit.Pixel(152)
                      .CalendarLayout.CalendarStyle.Height = Unit.Pixel(80)
                      .CalendarLayout.CalendarStyle.Font.Size = fontUnit
                      .CalendarLayout.CalendarStyle.Font.Name = "Verdana"
                      .CalendarLayout.CalendarStyle.BackColor = Color.White

                      .CalendarLayout.WeekendDayStyle.BackColor = General.GetColour(15004669)

                      .CalendarLayout.TodayDayStyle.ForeColor = General.GetColour(2774907)
                      .CalendarLayout.TodayDayStyle.BackColor = General.GetColour(10480637)

                      .CalendarLayout.DropDownStyle.Font.Size = fontUnit
                      .CalendarLayout.DropDownStyle.Font.Name = "Verdana"
                      .CalendarLayout.DropDownStyle.BorderStyle = BorderStyle.Solid
                      .CalendarLayout.DropDownStyle.BorderColor = General.GetColour(10720408)

                      .CalendarLayout.DayHeaderStyle.BackColor = General.GetColour(16248553)
                      .CalendarLayout.DayHeaderStyle.ForeColor = General.GetColour(6697779)
                      .CalendarLayout.DayHeaderStyle.Font.Size = fontUnit
                      .CalendarLayout.DayHeaderStyle.Font.Name = "Verdana"
                      .CalendarLayout.DayHeaderStyle.Font.Bold = True

                      .CalendarLayout.TitleStyle.ForeColor = General.GetColour(6697779)
                      .CalendarLayout.TitleStyle.BackColor = General.GetColour(16248553)
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

                      .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                      .BorderColor = General.GetColour(5730458)

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      If IsMobileBrowser() Then
                        .DropButton.ImageUrl1 = "~/Images/Calendar16.png"
                        .DropButton.ImageUrl2 = "~/Images/Calendar16.png"
                        .DropButton.ImageUrlHover = "~/Images/Calendar16.png"
                      Else
                        .DropButton.ImageUrl1 = "~/Images/downarrow.gif"
                        .DropButton.ImageUrl2 = "~/Images/downarrow.gif"
                        .DropButton.ImageUrlHover = "~/Images/downarrow-hover.gif"
                      End If
                      .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                      .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 2)

                      .ClientSideEvents.EditKeyDown = "dateControlKeyPress"
                      .ClientSideEvents.TextChanged = "dateControlTextChanged"
                      .ClientSideEvents.BeforeDropDown = "dateControlBeforeDropDown"

                      If IsMobileBrowser() Then .ClientSideEvents.AfterCloseUp = "FilterMobileLookup('" & sID.ToString & "');"
                    End With

                    pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_Date)
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Date)

                  End If


                Case 8 ' Frame
                  If NullSafeInteger(dr("BackStyle")) = 0 Then
                    sBackColour = "Transparent"
                  Else
                    sBackColour = General.GetHtmlColour(NullSafeInteger(dr("BackColor")))
                  End If

                  sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                   CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                  If sTemp2.Length = 0 Then
                    sTemp2 = " none"
                  End If

                  Dim top = NullSafeInteger(dr("TopCoord"))
                  Dim left = NullSafeInteger(dr("LeftCoord"))
                  Dim width = NullSafeInteger(dr("Width"))
                  Dim height = NullSafeInteger(dr("Height"))
                  Dim fontAdjustment = CInt(CInt(dr("FontSize")) * 0.8)
                  Dim borderCss As String = "BORDER-STYLE: solid; BORDER-COLOR: #9894a3; BORDER-WIDTH: 1px;"

                  width -= 2
                  height -= 2

                  If NullSafeString(dr("caption")).Trim.Length = 0 Then
                    top += fontAdjustment
                    height -= fontAdjustment
                  End If

                  sTemp = "<fieldset style='z-index: 0; " & _
                 " TOP: " & top & "px; " & _
                 " LEFT: " & left & "px; " & _
                 " WIDTH: " & width & "px; " & _
                 " HEIGHT: " & height & "px; " & _
                 " BACKGROUND-COLOR: " & sBackColour & "; " & _
                 " COLOR: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                 " font-family: " & NullSafeString(dr("FontName")) & "; " & _
                 " font-size: " & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt; " & _
                 " font-weight: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                 " font-style: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                 " text-decoration:" & sTemp2 & ";" & _
                 " " & borderCss & _
                 " POSITION: absolute;'>"

                  If NullSafeString(dr("caption")).Trim.Length > 0 Then

                    sTemp += "<legend style='color: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                    " margin-left: 5px; margin-right:5px; padding-left: 2px; padding-right: 2px;' align='Left'>" & _
                    NullSafeString(dr("caption")) & _
                    "</legend>"
                  End If

                  sTemp = sTemp & _
                  "</fieldset>" & vbCrLf

                  ctlForm_Frame = New LiteralControl(sTemp)

                  ' pnlInput.contenttemplatecontainer.Controls.Add(ctlForm_Frame)
                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Frame)

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
                       " BORDER-LEFT-COLOR:" & General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) & ";" & _
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
                       " BORDER-TOP-COLOR:" & General.GetHtmlColour(NullSafeInteger(dr("Backcolor"))) & ";" & _
                       " BORDER-TOP-STYLE:solid;" & _
                       " BORDER-TOP-WIDTH:1px'/>"
                  End Select

                  ctlForm_Line = New LiteralControl(sTemp)

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Line)

                Case 10 ' Image
                  ctlForm_Image = New WebControls.Image

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
                      .BorderColor = General.GetColour(10720408)
                      .BorderWidth = 1

                      iTempHeight = iTempHeight - 2
                      iTempWidth = iTempWidth - 2
                    End If

                    .Height() = Unit.Pixel(iTempHeight)
                    .Width() = Unit.Pixel(iTempWidth)

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
                    If getBrowserFamily() = "ANDROID" Then
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

                    .TabIndex = CShort(NullSafeInteger(dr("tabIndex")) + 1)

                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                      sDefaultFocusControl = sID
                      iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                    End If

                    .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                    .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                    .HeaderStyle.BackColor = General.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                    .HeaderStyle.BorderColor = General.GetColour(10720408)
                    .HeaderStyle.BorderStyle = BorderStyle.Double
                    .HeaderStyle.BorderWidth = Unit.Pixel(0)
                    .HeaderStyle.Font.Name = NullSafeString(dr("HeadFontName"))
                    .HeaderStyle.Font.Size = ToPointFontUnit(NullSafeInteger(dr("HeadFontSize")))
                    .HeaderStyle.Font.Bold = NullSafeBoolean(dr("HeadFontBold"))
                    .HeaderStyle.Font.Italic = NullSafeBoolean(dr("HeadFontItalic"))
                    .HeaderStyle.Font.Strikeout = NullSafeBoolean(dr("HeadFontStrikeThru"))
                    .HeaderStyle.Font.Underline = NullSafeBoolean(dr("HeadFontUnderline"))
                    .HeaderStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                    .HeaderStyle.Wrap = False
                    .HeaderStyle.Height = Unit.Pixel(iHeaderHeight)
                    .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                    .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                    ' PagerStyle settings
                    .PagerStyle.BackColor = General.GetColour(NullSafeInteger(dr("HeaderBackColor")))
                    .PagerStyle.BorderColor = General.GetColour(10720408)
                    .PagerStyle.BorderStyle = BorderStyle.Solid
                    .PagerStyle.BorderWidth = Unit.Pixel(0)
                    .PagerStyle.Font.Name = NullSafeString(dr("HeadFontName"))
                    .PagerStyle.Font.Size = ToPointFontUnit(NullSafeInteger(dr("HeadFontSize")))
                    .PagerStyle.Font.Bold = NullSafeBoolean(dr("HeadFontBold"))
                    .PagerStyle.Font.Italic = NullSafeBoolean(dr("HeadFontItalic"))
                    .PagerStyle.Font.Strikeout = NullSafeBoolean(dr("HeadFontStrikeThru"))
                    .PagerStyle.Font.Underline = NullSafeBoolean(dr("HeadFontUnderline"))
                    .PagerStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                    .PagerStyle.Wrap = False
                    .PagerStyle.VerticalAlign = VerticalAlign.Middle
                    .PagerStyle.HorizontalAlign = HorizontalAlign.Center

                    .Font.Name = NullSafeString(dr("FontName"))
                    .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                    .Font.Bold = NullSafeBoolean(dr("FontBold"))
                    .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                    .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                    .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                    ' ROW formatting
                    .AlternatingRowStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColorOdd")))
                    .AlternatingRowStyle.BackColor = General.GetColour(NullSafeInteger(dr("BackColorOdd")))

                    .RowStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColorEven")))
                    .RowStyle.BackColor = General.GetColour(NullSafeInteger(dr("BackColorEven")))

                    iRowHeight = 21

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
                        ''                                              
                        ctlForm_PagingGridView.IsEmpty = False
                        ctlForm_PagingGridView.DataSource = dt
                        ctlForm_PagingGridView.DataBind()

                      Else
                        ''                                                
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

                      fRecordOK = CBool(cmdGrid.Parameters("@pfOK").Value)
                      If Not fRecordOK Then
                        sMessage = "Error loading web form. Web Form record selector item record has been deleted or not selected."
                        Exit While
                      End If

                      cmdGrid.Dispose()

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

                    'If dtSource.Rows.Count = 0 Then
                    If ctlForm_PagingGridView.IsEmpty Then
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

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)


                Case 14 ' lookup  Inputs
                  If Not IsMobileBrowser() Then

                    ctlForm_UpdatePanel = New System.Web.UI.UpdatePanel

                    ' ============================================================
                    ' Create a textbox as the main control
                    ' ============================================================
                    ctlForm_TextInput = New TextBox

                    Dim backgroundColor As Color = General.GetColour(NullSafeInteger(dr("BackColor")))

                    With ctlForm_TextInput
                      .ID = sID & "TextBox"
                      .Style("position") = "absolute"
                      .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                      .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                      .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width")) - (2 * IMAGEBORDERWIDTH)).ToString
                      .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - (2 * IMAGEBORDERWIDTH))
                      .Attributes.CssStyle("HEIGHT") = Unit.Pixel(NullSafeInteger(dr("Height")) - (2 * IMAGEBORDERWIDTH)).ToString
                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                      .Width = Unit.Pixel(NullSafeInteger(dr("Width")))
                      .BackColor = backgroundColor
                      '.BackColor = Color.Transparent
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                      .BorderColor = General.GetColour(5730458)
                      .BorderStyle = BorderStyle.Solid
                      .BorderWidth = Unit.Pixel(1)
                      .ReadOnly = True
                      .Style.Add("z-index", "0")
                      .Style.Add("background-image", "url('images/downarrow.gif');")
                      .Style.Add("background-position", "right;")
                      .Style.Add("background-repeat", "no-repeat;")

                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_TextInput)

                    ' add a little dropdown to make the textbox look like a dropdown.                                    
                    ctlForm_Image = New WebControls.Image
                    Dim itmpDropDownWidth As Integer = 17

                    With ctlForm_Image
                      .ImageUrl = "Images/downarrow.gif"
                      .Style("position") = "absolute"
                      .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord")) + 1).ToString
                      .Style("left") = Unit.Pixel((NullSafeInteger(dr("LeftCoord")) + NullSafeInteger(dr("Width"))) - itmpDropDownWidth).ToString
                      .Attributes.CssStyle("WIDTH") = Unit.Pixel(itmpDropDownWidth).ToString
                      .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - (IMAGEBORDERWIDTH))
                      .Attributes.CssStyle("HEIGHT") = Unit.Pixel(NullSafeInteger(dr("Height")) - (IMAGEBORDERWIDTH)).ToString
                      .Style.Add("z-index", "0")
                    End With

                    'ctlForm_PageTab(iCurrentPageTab).controls.add(ctlForm_Image)

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
                      .Attributes.CssStyle("LEFT") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString
                      .Attributes.CssStyle("TOP") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                      .Attributes.CssStyle("WIDTH") = Unit.Pixel(NullSafeInteger(dr("Width"))).ToString

                      ' Don't set the height of this control. Must use the ControlHeight property
                      ' to stop the grid's rows from autosizing.
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
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .BorderColor = General.GetColour(5730458)
                      .BorderStyle = BorderStyle.Solid
                      .BorderWidth = Unit.Pixel(1)

                      .RowStyle.Font.Name = NullSafeString(dr("FontName"))
                      .RowStyle.Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .RowStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .RowStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .RowStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                      .RowStyle.BackColor = General.GetColour(15988214)
                      .RowStyle.ForeColor = General.GetColour(6697779)

                      .RowStyle.BorderColor = General.GetColour(10720408)
                      .RowStyle.BorderStyle = BorderStyle.Solid
                      .RowStyle.BorderWidth = Unit.Pixel(1)

                      iRowHeight = 21

                      .RowStyle.VerticalAlign = VerticalAlign.Middle

                      .SelectedRowStyle.ForeColor = General.GetColour(2774907)
                      .SelectedRowStyle.BackColor = General.GetColour(10480637)

                      ' HEADER formatting
                      .HeaderStyle.BackColor = General.GetColour(16248553)
                      .HeaderStyle.BorderColor = General.GetColour(10720408)
                      .HeaderStyle.BorderStyle = BorderStyle.Solid
                      .HeaderStyle.BorderWidth = Unit.Pixel(0)
                      .HeaderStyle.Font.Name = NullSafeString(dr("FontName"))
                      .HeaderStyle.Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .HeaderStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .HeaderStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .HeaderStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                      .HeaderStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
                      .HeaderStyle.Wrap = False
                      .HeaderStyle.Height = Unit.Pixel(iHeaderHeight)
                      .HeaderStyle.VerticalAlign = VerticalAlign.Middle
                      .HeaderStyle.HorizontalAlign = HorizontalAlign.Center

                      ' ROW formatting
                      .RowStyle.VerticalAlign = VerticalAlign.Middle
                      .PagerStyle.BorderWidth = Unit.Pixel(0)
                      .PagerStyle.Font.Name = NullSafeString(dr("FontName"))
                      .PagerStyle.Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .PagerStyle.Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .PagerStyle.Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .PagerStyle.Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .PagerStyle.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
                      .PagerStyle.ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))
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
                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)
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

                        '' Create a blank row at the top of the dropdown grid.
                        objDataRow = dt.NewRow()
                        dt.Rows.InsertAt(objDataRow, 0)

                        m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)

                        iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                        ctlForm_TextInput.Attributes.Remove("LookupColumnIndex")
                        ctlForm_TextInput.Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                        ctlForm_TextInput.Attributes.Remove("DefaultValue")
                        ctlForm_TextInput.Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                        ctlForm_TextInput.Attributes.Remove("DataType")
                        ctlForm_TextInput.Attributes.Add("DataType", NullSafeString(dt.Columns(CInt(ctlForm_TextInput.Attributes("LookupColumnIndex"))).DataType.ToString))

                        ' Yup we store the data to a session variable. This is so we can sort/filter 
                        ' it and stillreset if necessary without running the SP again
                        ctlForm_PagingGridView.DataSource = Session(sID & "DATA")
                        ctlForm_PagingGridView.DataBind()

                        cmdGrid.Dispose()

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
                      If ctlForm_TextInput.Attributes("DefaultValue").ToString.Length > 0 Then
                        ctlForm_TextInput.Text = ctlForm_TextInput.Attributes("DefaultValue").ToString
                      End If

                      For jncount As Integer = 0 To ctlForm_PagingGridView.Rows.Count - 1
                        If jncount > ctlForm_PagingGridView.PageSize Then Exit For ' don't bother if on other pages
                        If ctlForm_PagingGridView.Rows(jncount).Cells(m_iLookupColumnIndex).Text = ctlForm_TextInput.Attributes("DefaultValue").ToString Then
                          ctlForm_PagingGridView.SelectedIndex = jncount
                          Exit For
                        End If

                      Next
                    End If

                    ' =============================================================================
                    ' AJAX DropDownExtender (DDE) Control
                    ' This simply links up the DropDownList and the Lookup Grid to make a dropdown.
                    ' =============================================================================
                    Dim ctlForm_DDE As New AjaxControlToolkit.DropDownExtender

                    With ctlForm_DDE
                      .DropArrowBackColor = Color.Transparent
                      .DropArrowWidth = Unit.Pixel(20)
                      .HighlightBackColor = backgroundColor
                      '.HighlightBackColor = Color.Transparent

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

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_DDE)

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

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)

                    ' ====================================================
                    ' hidden field to hold any filter SQL code
                    ' ====================================================
                    ctlForm_HiddenField = New System.Web.UI.WebControls.HiddenField
                    With ctlForm_HiddenField
                      .ID = sID & "filterSQL"
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)

                    ' ============================================================
                    ' Hidden Button for JS to call which fires filter click event. 
                    ' ============================================================
                    ctlForm_InputButton = New Button

                    With ctlForm_InputButton
                      .ID = sID & "refresh"
                      .Style.Add("display", "none")
                      .Text = .ID
                    End With

                    AddHandler ctlForm_InputButton.Click, AddressOf SetLookupFilter

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_InputButton)

                  Else
                    ' ================================================================================================================
                    ' ================================================================================================================
                    ' ================================================================================================================
                    ' Mobile Browser - convert lookup data to a standard dropdown.
                    ' ================================================================================================================
                    ' ================================================================================================================
                    ' ================================================================================================================
                    ctlForm_Dropdown = New System.Web.UI.WebControls.DropDownList

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

                      If IsMobileBrowser() Then .Attributes.Add("onchange", "FilterMobileLookup('" & .ID.ToString & "');")

                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Dropdown)

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
                        ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)
                      End If

                      .Font.Name = NullSafeString(dr("FontName"))
                      .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                      .Font.Bold = NullSafeBoolean(dr("FontBold"))
                      .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                      .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                      .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                      .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                      .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                      .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                      .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 2)

                      ' HEADER formatting
                      iGridTopPadding = CInt(NullSafeSingle(dr("FontSize")) / 8)
                      iHeaderHeight = CInt(((NullSafeSingle(dr("FontSize")) + iGridTopPadding) * 2) _
                       - 2 _
                       - (NullSafeSingle(dr("FontSize")) * 2 * (iGridTopPadding - 1) / 4))

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

                          iRowHeight = CInt(.Height.Value) - 6
                          iRowHeight = CInt(IIf(iRowHeight < 22, 22, iRowHeight))
                          iDropHeight = (iRowHeight * CInt(IIf(dt.Rows.Count > MAXDROPDOWNROWS, MAXDROPDOWNROWS, dt.Rows.Count))) + 1

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
                    ctlForm_HiddenField = New System.Web.UI.WebControls.HiddenField
                    With ctlForm_HiddenField
                      .ID = sID & "filterSQL"
                    End With

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)

                    ' ============================================================
                    ' Hidden Button for JS to call which fires filter click event. 
                    ' ============================================================
                    ctlForm_InputButton = New Button

                    With ctlForm_InputButton
                      .ID = sID & "refresh"
                      .Style.Add("display", "none")
                    End With

                    AddHandler ctlForm_InputButton.Click, AddressOf SetLookupFilter

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_InputButton)

                  End If

                Case 13 ' Dropdown (13) Inputs

                  ctlForm_Dropdown = New System.Web.UI.WebControls.DropDownList

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

                    If IsMobileBrowser() Then .Attributes.Add("onchange", "FilterMobileLookup('" & .ID.ToString & "');")

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Dropdown)

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
                      ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)
                    End If

                    .Font.Name = NullSafeString(dr("FontName"))
                    .Font.Size = ToPointFontUnit(NullSafeInteger(dr("FontSize")))
                    .Font.Bold = NullSafeBoolean(dr("FontBold"))
                    .Font.Italic = NullSafeBoolean(dr("FontItalic"))
                    .Font.Strikeout = NullSafeBoolean(dr("FontStrikeThru"))
                    .Font.Underline = NullSafeBoolean(dr("FontUnderline"))

                    .BackColor = General.GetColour(NullSafeInteger(dr("BackColor")))
                    .ForeColor = General.GetColour(NullSafeInteger(dr("ForeColor")))

                    .Height() = Unit.Pixel(NullSafeInteger(dr("Height")) - 2)
                    .Width() = Unit.Pixel(NullSafeInteger(dr("Width")) - 2)

                    ' HEADER formatting
                    iGridTopPadding = CInt(NullSafeSingle(dr("FontSize")) / 8)
                    iHeaderHeight = CInt(((NullSafeSingle(dr("FontSize")) + iGridTopPadding) * 2) _
                     - 2 _
                     - (NullSafeSingle(dr("FontSize")) * 2 * (iGridTopPadding - 1) / 4))

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
                        Session(sID & "_DATA") = dt

                        ' Format the column(s)
                        For Each objGridColumn In dt.Columns
                          If objGridColumn.ColumnName.StartsWith("ASRSys") Then
                          Else
                            .DataTextField = objGridColumn.ColumnName.ToString
                          End If
                        Next objGridColumn

                        ctlForm_Dropdown.DataSource = dt

                        m_iLookupColumnIndex = NullSafeInteger(cmdGrid.Parameters("@piLookupColumnIndex").Value)
                        iItemType = NullSafeInteger(cmdGrid.Parameters("@piItemType").Value)

                        .Attributes.Remove("LookupColumnIndex")
                        .Attributes.Add("LookupColumnIndex", m_iLookupColumnIndex.ToString)

                        .Attributes.Remove("DefaultValue")
                        .Attributes.Add("DefaultValue", NullSafeString(cmdGrid.Parameters("@psDefaultValue").Value))

                        ctlForm_Dropdown.DataBind()

                        cmdGrid.Dispose()

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

                      ' ==================================================
                      ' Set the dropdownList to the default value.
                      ' ==================================================

                      Dim listItem As ListItem = ctlForm_Dropdown.Items.FindByValue(ctlForm_Dropdown.Attributes("DefaultValue").ToString)
                      If listItem IsNot Nothing Then
                        ctlForm_Dropdown.SelectedValue = listItem.Value
                      End If


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
                    sBackColour = General.GetHtmlColour(NullSafeInteger(dr("BackColor")))
                  End If

                  sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                           CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                  If sTemp2.Length = 0 Then
                    sTemp2 = " none"
                  End If

                  Dim top = NullSafeInteger(dr("TopCoord"))
                  Dim left = NullSafeInteger(dr("LeftCoord"))
                  Dim width = NullSafeInteger(dr("Width"))
                  Dim height = NullSafeInteger(dr("Height"))
                  Dim fontAdjustment = CInt(CInt(dr("FontSize")) * 0.8)
                  Dim startAdjustment As Integer = 0
                  Dim borderCss As String = String.Empty

                  If Not NullSafeBoolean(dr("PictureBorder")) Then
                    borderCss = "BORDER-STYLE: none;"
                  Else
                    borderCss = "BORDER-STYLE: solid; BORDER-COLOR: #9894a3; BORDER-WIDTH: 1px;"
                    width -= 2
                    height -= 2

                    If NullSafeString(dr("caption")).Trim.Length > 0 Then
                      startAdjustment = fontAdjustment

                      If BrowserRequiresFieldsetAdjustment() Then
                        startAdjustment -= 10
                      End If
                    Else
                      top += fontAdjustment
                      height -= fontAdjustment
                    End If
                  End If

                  sTemp = "<fieldset style='z-index: 0; " & _
                   " TOP: " & top & "px; " & _
                   " LEFT: " & left & "px; " & _
                   " WIDTH: " & width & "px; " & _
                   " HEIGHT: " & height & "px; " & _
                   " BACKGROUND-COLOR: " & sBackColour & "; " & _
                   " COLOR: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                   " font-family: " & NullSafeString(dr("FontName")) & "; " & _
                   " font-size: " & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt; " & _
                   " font-weight: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                   " font-style: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                   " text-decoration:" & sTemp2 & ";" & _
                   " " & borderCss & _
                   " POSITION: absolute;'>"

                  If NullSafeBoolean(dr("PictureBorder")) And (NullSafeString(dr("caption")).Trim.Length > 0) Then

                    sTemp += "<legend style='color: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & ";" & _
                    " margin-left: 5px; margin-right:5px; padding-left: 2px; padding-right: 2px;' align='Left'>" & _
                    NullSafeString(dr("caption")) & _
                    "</legend>"
                  End If

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
                              CInt(IIf(NullSafeBoolean(dr("PictureBorder")), 0, 10)) + _
                              startAdjustment

                          Dim spanLeft As Int32 = CInt(((NullSafeSingle(dr("FontSize"))) * 5 / 4) - 1)

                          sTemp = sTemp & _
                           "<span tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                           " style=""z-index: 0;" & _
                           " font-family: " & NullSafeString(dr("FontName")) & "; " & _
                           " font-size: " & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt; " & _
                           " font-weight: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                           " font-style: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                           " text-decoration:" & sTemp2 & ";" & _
                           " left: " & spanLeft.ToString & "px; position: absolute; top: " & spanTop.ToString & "px"">" & _
                           " <input id=""opt" & sID & "_" & iTemp.ToString & """ type=""radio""" & _
                           " style=""margin: 0px; padding: 3px;""" & _
                           " name=""opt" & sID & """ value=""" & drGrid(0).ToString & """" & _
                           " onfocus = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='#ff9608'; activateControl();}catch(e){};""" & _
                           " onblur = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='';}catch(e){};""" & _
                           " onclick = """ & sID & ".value=opt" & sID & "[" & iTemp.ToString & "].value;""" & _
                                                  CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", ""))

                        Case 1 ' Horizontal
                          stringSize = graphic.MeasureString(drGrid(0).ToString(), font)
                          Dim spanTop As Int32 = CInt((NullSafeInteger(dr("FontSize")) * 1.25) + 1) - _
                              CInt(IIf(NullSafeBoolean(dr("PictureBorder")), 0, 10)) + _
                              startAdjustment

                          sTemp = sTemp & _
                           "<span tabindex=" & CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString & _
                           " style=""z-index: 0;" & _
                           " font-family: " & NullSafeString(dr("FontName")) & "; " & _
                           " font-size: " & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt; " & _
                           " font-weight: " & CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")) & ";" & _
                           " font-style: " & CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")) & ";" & _
                           " text-decoration:" & sTemp2 & ";" & _
                           " left: " & lastLeft & "px; position: absolute; top: " & spanTop.ToString & "px"">" & _
                           " <input id=""opt" & sID & "_" & iTemp.ToString & """ type=""radio""" & _
                           " style=""margin: 0px; padding: 3px;""" & _
                           " name=""opt" & sID & """ value=""" & drGrid(0).ToString & """" & _
                           " onfocus = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='#ff9608'; activateControl();}catch(e){};""" & _
                           " onblur = ""try{forOpt" & sID & "_" & iTemp.ToString & ".style.color='';}catch(e){};""" & _
                           " onclick = """ & sID & ".value=opt" & sID & "[" & iTemp.ToString & "].value;""" & _
                                            CStr(IIf(IsMobileBrowser, " FilterMobileLookup('" & sID.ToString & "');""", ""))

                          lastLeft += (stringSize.Width + (font.Size * 2) + 28)
                      End Select

                      If iTemp = 0 Or CInt(drGrid.GetValue(1)) = 1 Then
                        sTemp = sTemp & _
                         " Checked=""checked"""
                        sDefaultValue = drGrid(0).ToString
                      End If

                      sTemp = sTemp & _
                      "/>" & _
                      " <label id=""forOpt" & sID & "_" & iTemp.ToString & """ for=""opt" & sID & "_" & iTemp.ToString & """ tabindex=""-1""" _
                      & " style=""position: relative; top: -2px;""" _
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
                    cmdGrid.Dispose()

                    sTemp = sTemp & _
                     "</fieldset>" & vbCrLf

                    ctlForm_OptionGroup = New LiteralControl(sTemp)

                    ' pnlInput.ContentTemplateContainer.Controls.Add(ctlForm_OptionGroup)
                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_OptionGroup)

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

                    ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_OptionGroupReal)

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

                  ctlForm_HTMLInputButton = New HtmlInputButton
                  With ctlForm_HTMLInputButton
                    .ID = sID
                    .Attributes.Add("TabIndex", CShort(NullSafeInteger(dr("tabIndex")) + 1).ToString)

                    If (iMinTabIndex < 0) Or (NullSafeInteger(dr("tabIndex")) < iMinTabIndex) Then
                      sDefaultFocusControl = sID
                      iMinTabIndex = NullSafeInteger(dr("tabIndex"))
                    End If

                    .Style("position") = "absolute"
                    .Style("top") = Unit.Pixel(NullSafeInteger(dr("TopCoord"))).ToString
                    .Style("left") = Unit.Pixel(NullSafeInteger(dr("LeftCoord"))).ToString

                    ' stops the mobiles displaying buttons with over-rounded corners...
                    If IsMobileBrowser() Then
                      .Style.Add("-webkit-appearance", "none")
                      .Style.Add("background-color", "#CCCCCC")
                      .Style.Add("border", "solid 1px #C0C0C0")
                      .Style.Add("border-radius", "0px")
                    End If

                    If NullSafeInteger(dr("BackColor")) <> 16249587 AndAlso NullSafeInteger(dr("BackColor")) <> -2147483633 Then
                      .Style.Add("background-color", General.GetHtmlColour(NullSafeInteger(dr("BackColor"))).ToString)
                      .Style.Add("border", "solid 1px " & General.GetHtmlColour(9999523).ToString)
                    End If

                    If NullSafeInteger(dr("ForeColor")) <> 6697779 Then
                      .Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))).ToString)
                    End If

                    .Style.Add("padding", "0px")
                    .Style.Add("white-space", "normal")

                    .Value = NullSafeString(dr("caption"))

                    sTemp2 = CStr(IIf(NullSafeBoolean(dr("FontStrikeThru")), " line-through", "")) & _
                       CStr(IIf(NullSafeBoolean(dr("FontUnderline")), " underline", ""))

                    If sTemp2.Length = 0 Then
                      sTemp2 = " none"
                    End If

                    .Style.Add("Font-family", NullSafeString(dr("FontName")).ToString)
                    .Style.Add("Font-Size", ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt")
                    .Style.Add("Font-weight", CStr(IIf(NullSafeBoolean(dr("FontBold")), "bold", "normal")))
                    .Style.Add("FontStyle", CStr(IIf(NullSafeBoolean(dr("FontItalic")), "italic", "normal")))
                    .Style.Add("Text-Decoration", sTemp2)

                    .Style.Add("Width", Unit.Pixel(NullSafeInteger(dr("Width"))).ToString)
                    .Style.Add("Height", Unit.Pixel(NullSafeInteger(dr("Height"))).ToString)

                    If Not IsMobileBrowser() Then
                      .Attributes.Add("onclick", "try{showFileUpload(true, '" & sEncodedID & "', document.getElementById('file" & sID & "').value);}catch(e){};")

                      AddHandler ctlForm_HTMLInputButton.ServerClick, AddressOf Me.DisableControls
                    Else
                      .Attributes.Add("onclick", "try{alert('Your browser does not support file upload.');}catch(e){};")
                    End If
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HTMLInputButton)

                  ctlForm_HiddenField = New HiddenField
                  With ctlForm_HiddenField
                    .ID = "file" & sID
                    .Value = NullSafeString(dr("value"))
                  End With

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_HiddenField)
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
                    sBackColour = General.GetHtmlColour(NullSafeInteger(dr("BackColor")))
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
                   " font-size:" & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt;" & _
                   " font-weight:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontBold"))), "bold;", "normal;").ToString & _
                   " font-style:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontItalic"))), "italic;", "normal;").ToString & _
                   " text-decoration:" & sDecoration & ";" & _
                   " background-color: " & sBackColour & "; " & _
                   " color: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                   "' onclick='FileDownload_Click(""" & sEncodedID & """);'" & _
                   " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" & _
                   " onmouseover=""this.style.cursor='pointer';this.style.color='#ff9608';""" & _
                   " onmouseout=""this.style.cursor='';this.style.color='" & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "';""" & _
                   " onfocus=""this.style.color='#ff9608';""" & _
                   " onblur=""this.style.color='" & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "';"">" & _
                   HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) & _
                   "</span>"

                  ctlForm_Literal = New LiteralControl(sTemp)

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Literal)

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
                    sBackColour = General.GetHtmlColour(NullSafeInteger(dr("BackColor")))
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
                   " font-size:" & ToPoint(NullSafeInteger(dr("FontSize"))).ToString & "pt;" & _
                   " font-weight:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontBold"))), "bold;", "normal;").ToString & _
                   " font-style:" & IIf(NullSafeBoolean(NullSafeBoolean(dr("FontItalic"))), "italic;", "normal;").ToString & _
                   " text-decoration:" & sDecoration & ";" & _
                   " background-color: " & sBackColour & "; " & _
                   " color: " & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "; " & _
                   "' onclick='FileDownload_Click(""" & sEncodedID & """);'" & _
                   " onkeypress='FileDownload_KeyPress(""" & sEncodedID & """);'" & _
                   " onmouseover=""this.style.cursor='pointer';this.style.color='#ff9608';""" & _
                   " onmouseout=""this.style.cursor='';this.style.color='" & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "';""" & _
                   " onfocus=""this.style.color='#ff9608';""" & _
                   " onblur=""this.style.color='" & General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))) & "';"">" & _
                   HttpUtility.HtmlEncode(NullSafeString(dr("caption"))) & _
                   "</span>"

                  ctlForm_Literal = New LiteralControl(sTemp)

                  ctlForm_PageTab(iCurrentPageTab).Controls.Add(ctlForm_Literal)

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

                  If IsMobileBrowser() Then
                    ctlTabsDiv.Style.Add("overflow-x", "auto")
                  Else
                    ' for non-mobile browsers we display arrows to scroll the tab bar left and right.
                    ctlTabsDiv.Style.Add("overflow", "hidden")
                    ctlTabsDiv.Style.Add("margin-right", "51px")

                    ' Nav arrows for non-mobile browsers
                    Dim ctlForm_TabArrows As New Panel
                    With ctlForm_TabArrows
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
                    ctlForm_TabArrows.Controls.Add(ctlForm_Image)

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
                    ctlForm_TabArrows.Controls.Add(ctlForm_Image)

                    pnlTabsDiv.Controls.Add(ctlForm_TabArrows)

                  End If

                  ' generate the tabs.
                  Dim ctlTabsTable As New Table
                  ctlTabsTable.CellSpacing = 0
                  ' ctlTabsTable.Style.Add("margin-top", "2px")
                  Dim trPager As TableRow = New TableRow()
                  trPager.Height = Unit.Pixel(miTabStripHeight - 1) ' to prevent vertical scrollbar
                  trPager.Style.Add("white-space", "nowrap")

                  Dim tcTabCell As New TableCell
                  'Dim ctlForm_Label As New Label

                  Dim iTabNo As Integer = 1
                  ' add a cell for each tab
                  For Each sTabCaption In arrTabCaptions
                    If sTabCaption.Trim.Length > 0 Then
                      tcTabCell = New TableCell

                      With tcTabCell
                        .ID = "forminput_" & iTabNo.ToString & "_21_Panel"
                        .BorderColor = Color.Black
                        .Style.Add("padding-left", "5px")
                        .Style.Add("padding-right", "5px")
                        .Style.Add("border-radius", "5px 5px 0px 0px")
                        .Style.Add("width", "50px")
                        .BorderWidth = 1
                        .BorderStyle = BorderStyle.Solid
                        .BackColor = Color.White

                        ' label the button...
                        ctlForm_Label = New Label
                        ctlForm_Label.Font.Name = "Verdana"
                        ctlForm_Label.Font.Size = New FontUnit(11, UnitType.Pixel)
                        ctlForm_Label.Text = sTabCaption.ToString

                        .Controls.Add(ctlForm_Label)

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
                          iTempHeight = miTabStripHeight + NullSafeInteger(dr("TopCoord"))
                          iTempWidth = NullSafeInteger(dr("LeftCoord"))

                          ctlForm_PageTab(iTabNo).Style.Add("top", iTempHeight & "px")
                          ctlForm_PageTab(iTabNo).Style.Add("left", iTempWidth & "px")

                          ' Hide all tabs but the first.
                          ctlForm_PageTab(iTabNo).Style.Add("display", "none")
                        Catch ex As Exception
                          Beep()
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
                sBackgroundImage = ""
                sBackgroundRepeat = ""
                sBackgroundPosition = ""
                If CInt(cmdSelect.Parameters("@piBackImage").Value) > 0 Then
                  sBackgroundImage = LoadPicture(CInt(cmdSelect.Parameters("@piBackImage").Value), sMessage)
                  If sMessage.Length = 0 Then
                    divInput.Style("Background-image") = sBackgroundImage
                  End If
                  If sMessage.Length = 0 Then
                    sBackgroundImage = "url('" & sBackgroundImage & "')"
                  End If

                  iBackgroundImagePosition = CInt(cmdSelect.Parameters("@piBackImageLocation").Value())
                  sBackgroundRepeat = General.BackgroundRepeat(CShort(iBackgroundImagePosition))
                  sBackgroundPosition = General.BackgroundPosition(CShort(iBackgroundImagePosition))

                End If
                divInput.Style("background-repeat") = sBackgroundRepeat
                divInput.Style("background-position") = sBackgroundPosition

                sBackgroundColourHex = ""
                If Not IsDBNull(cmdSelect.Parameters("@piBackColour").Value) Then
                  iBackgroundColour = CInt(cmdSelect.Parameters("@piBackColour").Value())
                  sBackgroundColourHex = General.GetHtmlColour(iBackgroundColour).ToString()

                  divInput.Style("Background-color") = General.GetHtmlColour(NullSafeInteger(iBackgroundColour))
                End If

                iFormWidth = CInt(cmdSelect.Parameters("@piWidth").Value)
                iFormHeight = CInt(cmdSelect.Parameters("@piHeight").Value)

                pnlInputDiv.Style("width") = iFormWidth.ToString & "px"
                pnlInputDiv.Style("height") = iFormHeight.ToString & "px"
                pnlInputDiv.Style("left") = "-2px"

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

          ' Resize the mobile 'viewport' to fit the webform - mobiles only
          'If isMobileBrowser() Then AddHeaderTags(iFormWidth)
          AddHeaderTags(iFormWidth)

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

  End Sub

  Public Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)

    Dim conn As SqlConnection
    Dim dr As SqlDataReader
    Dim cmdValidate As SqlCommand
    Dim cmdUpdate As SqlCommand
    Dim cmdQS As SqlCommand
    Dim sFormInput1 As String
    Dim sFormValidation1 As String
    Dim ctlFormInput As Control
    Dim ctlFormCheckBox As CheckBox
    Dim ctlFormTextInput As TextBox
    Dim ctlFormDateInput As Infragistics.WebUI.WebSchedule.WebDateChooser
    Dim ctlFormNumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Dim ctlForm_PagingGridView As RecordSelector
    Dim ctlFormDropdown As DropDownList
    Dim ctlForm_HiddenField As HiddenField


    Dim sID As String
    Dim sIDString As String
    Dim sDateValueString As String
    Dim sNumValueString As String
    Dim iTemp As Int16
    Dim sTemp As String
    Dim iType As Int16
    Dim sType As String
    Dim sRecordID As String
    Dim sFormElements As String
    Dim arrFollowOnForms() As String
    Dim fSavedForLater As Boolean
    Dim sMessage As String
    Dim sMessage1 As String
    Dim sMessage2 As String
    Dim sMessage3 As String
    Dim iFollowOnFormCount As Integer
    Dim iIndex As Integer
    Dim sStep As String
    Dim arrQueryStrings() As String
    Dim sFollowOnForms As String
    Dim sColumnCaption As String

    sMessage = ""
    sMessage1 = ""
    sMessage2 = ""
    sMessage3 = ""
    sFormInput1 = ""
    sFormValidation1 = ""
    sFollowOnForms = ""
    ReDim arrQueryStrings(0)

    Try
      ' Read the web form item values & build up a string of the form input values.
      ' This is a tab delimited string of itemIDs and values.

      Dim controlList = pnlInputDiv.Controls.Cast(Of Control)() _
            .Union(pnlTabsDiv.Controls.Cast(Of Control)) _
            .Where(Function(c) c.ClientID.EndsWith("_PageTab")) _
            .SelectMany(Function(c) c.Controls.Cast(Of Control)()) _
            .Where(Function(c) c.ClientID.StartsWith(FORMINPUTPREFIX))

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
              sFormInput1 = sFormInput1 & sIDString & "1" & vbTab
              sFormValidation1 = sFormValidation1 & sIDString & "1" & vbTab
            Else
              If (TypeOf ctlFormInput Is HtmlInputButton) Then
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

            ' Is this an HTML5 compliant mobile device?
            If (TypeOf ctlFormInput Is HtmlInputText) Then

              If pnlInput.FindControl(sID & "Value") Is Nothing Then
                sDateValueString = "null"
              Else
                ctlForm_HiddenField = DirectCast(pnlInput.FindControl(sID & "Value"), HiddenField)
                If ctlForm_HiddenField.Value.ToString = "" Then
                  sDateValueString = "null"
                Else
                  sDateValueString = Format(DateTime.Parse(ctlForm_HiddenField.Value), "MM/dd/yyyy")
                  'sDateValueString = General.ConvertLocaleDateToSQL(ctlForm_HiddenField.Value)
                End If
              End If

              'ctlFormHTMLInputText = DirectCast(ctlFormInput, HtmlInputText)

              'If (ctlFormHTMLInputText.Value.ToString = vbNullString) Or (ctlFormHTMLInputText.Value.ToString = "  /  /") Then
              '  sDateValueString = "null"
              'Else
              '  sDateValueString = General.ConvertLocaleDateToSQL(ctlFormHTMLInputText.Value)
              'End If

              sFormInput1 = sFormInput1 & sIDString & sDateValueString & vbTab
              sFormValidation1 = sFormValidation1 & sIDString & sDateValueString & vbTab
            End If

          Case 11 ' Grid (RecordSelector) Input
            If (TypeOf ctlFormInput Is RecordSelector) Then
              ctlForm_PagingGridView = DirectCast(ctlFormInput, RecordSelector)
              sRecordID = "0"

              If Not ctlForm_PagingGridView.IsEmpty And ctlForm_PagingGridView.SelectedIndex >= 0 Then
                For iColCount As Integer = 0 To ctlForm_PagingGridView.HeaderRow.Cells.Count - 1
                  sColumnCaption = UCase(ctlForm_PagingGridView.HeaderRow.Cells(iColCount).Text)

                  If (sColumnCaption = "ID") Then
                    sRecordID = ctlForm_PagingGridView.SelectedRow.Cells(iColCount).Text
                    Exit For
                  End If
                Next
              End If

              sFormInput1 = sFormInput1 & sIDString & sRecordID & vbTab
              sFormValidation1 = sFormValidation1 & sIDString & sRecordID & vbTab
            End If

          Case 13 ' Dropdown Input
            If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then
              ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList)

              sTemp = ctlFormDropdown.Text ' .DisplayValue
              sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
              sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
            End If

          Case 14 ' Lookup Input
            If Not IsMobileBrowser() Then


              If (TypeOf ctlFormInput Is System.Web.UI.WebControls.TextBox) Then
                ctlFormTextInput = DirectCast(ctlFormInput, System.Web.UI.WebControls.TextBox)

                sTemp = ctlFormTextInput.Text

                If ctlFormTextInput.Attributes("DataType") = "System.DateTime" Then
                  If sTemp Is Nothing Then
                    sTemp = "null"
                  Else
                    If (sTemp.Length = 0) Then
                      sTemp = "null"
                    Else
                      sTemp = General.ConvertLocaleDateToSQL(sTemp)
                    End If
                  End If
                ElseIf ctlFormTextInput.Attributes("DataType") = "System.Decimal" _
                 Or ctlFormTextInput.Attributes("DataType") = "System.Int32" Then

                  If sTemp Is Nothing Then
                    sTemp = ""
                  Else
                    sTemp = CStr(IIf(sTemp.Length = 0, "", CStr(sTemp).Replace(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")))
                  End If

                End If

                sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
                sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
              End If
            Else
              ' Mobile Browser - it's a Dropdown List.
              If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)

                sTemp = ctlFormDropdown.Text ' .DisplayValue
                sFormInput1 = sFormInput1 & sIDString & sTemp & vbTab
                sFormValidation1 = sFormValidation1 & sIDString & sTemp & vbTab
              End If

            End If

          Case 15 ' OptionGroup Input
            If (TypeOf ctlFormInput Is TextBox) Then
              ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
              sFormInput1 = sFormInput1 & sIDString & ctlFormTextInput.Text & vbTab
              sFormValidation1 = sFormValidation1 & sIDString & ctlFormTextInput.Text & vbTab
            End If

          Case 17 ' FileUpload
            ' If (TypeOf ctlFormInput Is Infragistics.WebUI.WebDataInput.WebImageButton) Then
            If (TypeOf ctlFormInput Is HtmlInputButton) Then

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

      Next

    Catch ex As Exception
      sMessage = "Error reading web form item values:<BR><BR>" & ex.Message
    End Try

    If sMessage.Length = 0 Then
      Try
        conn = New SqlConnection(GetConnectionString)
        conn.Open()

        Try ' Validate the web form entry.
          lblErrors.Font.Size = mobjConfig.ValidationMessageFontSize
          lblErrors.ForeColor = General.GetColour(6697779)

          lblWarnings.Font.Size = mobjConfig.ValidationMessageFontSize
          lblWarnings.ForeColor = General.GetColour(6697779)
          lblWarningsPrompt_1.Font.Size = mobjConfig.ValidationMessageFontSize
          lblWarningsPrompt_1.ForeColor = General.GetColour(6697779)
          lblWarningsPrompt_2.Font.Size = mobjConfig.ValidationMessageFontSize
          lblWarningsPrompt_3.Font.Size = mobjConfig.ValidationMessageFontSize
          lblWarningsPrompt_3.ForeColor = General.GetColour(6697779)

          bulletErrors.Font.Size = mobjConfig.ValidationMessageFontSize
          bulletErrors.ForeColor = General.GetColour(6697779)

          bulletWarnings.Font.Size = mobjConfig.ValidationMessageFontSize
          bulletWarnings.ForeColor = General.GetColour(6697779)

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
              cmdUpdate.Parameters("@psFormInput1").Value = sFormInput1

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
                    If Not General.SplitMessage(msSavedForLaterMessage, sMessage1, sMessage2, sMessage3) Then
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
                    If Not General.SplitMessage(msCompletionMessage, sMessage1, sMessage2, sMessage3) Then
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

                  cmdQS = New SqlCommand("spASRGetWorkflowQueryString", conn)
                  cmdQS.CommandType = CommandType.StoredProcedure
                  cmdQS.CommandTimeout = miSubmissionTimeoutInSeconds

                  cmdQS.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdQS.Parameters("@piInstanceID").Value = miInstanceID

                  cmdQS.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdQS.Parameters("@piElementID").Value = CLng(sStep)

                  cmdQS.Parameters.Add("@psQueryString", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                  cmdQS.ExecuteNonQuery()

                  Dim sQueryString As String = CStr(cmdQS.Parameters("@psQueryString").Value())

                  ReDim Preserve arrQueryStrings(arrQueryStrings.GetUpperBound(0) + 1)
                  arrQueryStrings(arrQueryStrings.GetUpperBound(0)) = sQueryString

                  cmdQS.Dispose()
                Next iIndex

                sFollowOnForms = Join(arrQueryStrings, vbTab)

                Select Case miFollowOnFormsMessageType
                  Case 1 ' Custom
                    If Not General.SplitMessage(msFollowOnFormsMessage, sMessage1, sMessage2, sMessage3) Then
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

  Public Sub DisableControls(ByVal sender As System.Object, e As System.EventArgs)
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
    Dim ctlFormHTMLInputButton As HtmlInputButton
    Dim ctlFormDateInput As Infragistics.WebUI.WebSchedule.WebDateChooser
    Dim ctlFormNumericInput As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Dim ctlFormRecordSelectionGrid As System.Web.UI.WebControls.GridView
    Dim ctlFormDropdown As System.Web.UI.WebControls.DropDownList
    Dim sIDString As String
    Dim iTemp As Int16
    Dim sTemp As String
    Dim iType As Int16
    Dim sType As String
    Dim sMessage As String = ""
    Dim sMessage1 As String = ""
    Dim sMessage2 As String = ""
    Dim sMessage3 As String = ""

    Try ' Disable all controls.
      Dim controlList = pnlInputDiv.Controls.Cast(Of Control)() _
            .Union(pnlTabsDiv.Controls.Cast(Of Control)) _
            .Where(Function(c) c.ClientID.EndsWith("_PageTab")) _
            .SelectMany(Function(c) c.Controls.Cast(Of Control)()) _
            .Where(Function(c) c.ClientID.StartsWith(FORMINPUTPREFIX))

      For Each ctlFormInput In controlList
        'For Each ctlFormInput In pnlInput.ContentTemplateContainer.Controls

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
              ctlFormHTMLInputButton = DirectCast(ctlFormInput, HtmlInputButton)
              'ctlFormHTMLInputButton.Style.Remove("disabled")
              'If Not pfEnabled Then ctlFormHTMLInputButton.Style.Add("disabled", "disabled")

              ctlFormHTMLInputButton.Attributes.Remove("disabled")
              If Not pfEnabled Then ctlFormHTMLInputButton.Attributes.Add("disabled", "disabled")

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
              End If

            Case 13 ' Dropdown Input
              If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList) 'Infragistics.WebUI.WebCombo.WebCombo)
                ctlFormDropdown.Enabled = pfEnabled
              End If

            Case 14 ' Lookup Input
              If Not IsMobileBrowser() Then

                If (TypeOf ctlFormInput Is AjaxControlToolkit.DropDownExtender) Then 'Infragistics.WebUI.WebCombo.WebCombo) Then
                  DirectCast(ctlFormInput, AjaxControlToolkit.DropDownExtender).Enabled = pfEnabled
                End If
              Else
                ' Mobile Browser
                If (TypeOf ctlFormInput Is System.Web.UI.WebControls.DropDownList) Then
                  ctlFormDropdown = DirectCast(ctlFormInput, System.Web.UI.WebControls.DropDownList)
                  ctlFormDropdown.Enabled = pfEnabled
                End If
              End If

            Case 15 ' OptionGroup Input
              If (TypeOf ctlFormInput Is TextBox) Then
                ctlFormTextInput = DirectCast(ctlFormInput, TextBox)
                ctlFormTextInput.Enabled = pfEnabled
              End If

            Case 17 ' Input value - file upload
              ctlFormHTMLInputButton = DirectCast(ctlFormInput, HtmlInputButton)
              ctlFormHTMLInputButton.Style.Remove("disabled")
              If pfEnabled Then ctlFormHTMLInputButton.Style.Add("disabled", "disabled")

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

  Public Function SubmissionTimeout() As Int32
    SubmissionTimeout = mobjConfig.SubmissionTimeout
  End Function

  Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
    ButtonClick(sender, e)
  End Sub

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
    Dim fs As System.IO.FileStream
    Dim bw As System.IO.BinaryWriter
    Dim iBufferSize As Integer = 100
    Dim outByte(iBufferSize - 1) As Byte
    Dim retVal As Long
    Dim startIndex As Long
    Dim sExtension As String = ""
    Dim iIndex As Integer
    Dim sName As String

    Try
      miImageCount = CShort(miImageCount + 1)

      psErrorMessage = ""
      LoadPicture = ""
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

  Protected Sub btnReEnableControls_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReEnableControls.Click
    EnableDisableControls(True)
  End Sub

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
    'gv.Rows(0).Cells(0).ForeColor = Color.Red
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

    If TypeOf (pnlInputDiv.FindControl(btnSender.ID.Replace("refresh", ""))) Is DropDownList Then

      ' This is a dropdownlist style lookup (mobiles only)
      Dim dropdown As DropDownList
      dropdown = TryCast(pnlInputDiv.FindControl(btnSender.ID.Replace("refresh", "")), DropDownList)

      ' Store the current value, so we can re-add it after filtering.
      Dim strCurrentSelection As String = dropdown.Text

      ' Filter the table now.
      filterDataTable(dataTable, filterSQL)

      ' insert the previously selected item
      Dim objDataRow As DataRow
      objDataRow = dataTable.NewRow()
      'objDataRow.ItemArray = dataTable.Rows(dropdown.SelectedIndex).ItemArray
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

      filterDataTable(dataTable, filterSQL)

      Dim gridView As RecordSelector 'GridView
      gridView = TryCast(pnlInputDiv.FindControl(btnSender.ID.Replace("refresh", "Grid")), RecordSelector)

      gridView.filterSQL = filterSQL.ToString

      gridView.DataSource = dataTable
      gridView.DataBind()
    End If

    ' reset filter.
    hiddenField.Value = ""


  End Sub

  Private Sub filterDataTable(ByRef dataTable As DataTable, ByVal filterSQL As String)
    If dataTable IsNot Nothing Then
      Dim dataView As New DataView(dataTable)
      dataView.RowFilter = filterSQL

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
    If isMobileBrowser() Then
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

  Private Function getBrowserFamily() As String

    Dim ua As String = Request.UserAgent.ToUpper

    If ua.Contains("MSIE") Then
      Return "MSIE"
    ElseIf ua.Contains("IPHONE") OrElse ua.Contains("IPAD") Then
      Return "IOS"
    ElseIf ua.Contains("ANDROID") Then
      Return "ANDROID"
    ElseIf ua.Contains("BLACKBERRY") Then
      Return "BLACKBERRY"
    Else
      Return "UNKNOWN"
    End If
  End Function

End Class
