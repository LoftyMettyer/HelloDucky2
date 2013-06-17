﻿Imports System.Data
Imports System.Linq
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.IO

Public Class Database
   Private ReadOnly _connectionString As String
   Private ReadOnly _timeout As Integer

   Public Sub New(connectionString As String)
      _connectionString = connectionString
      _timeout = App.Config.SubmissionTimeoutInSeconds
   End Sub

   Public Shared Function CreateConnectionString(server As String, database As String, user As String, password As String) As String
      Return "Application Name=OpenHR Workflow;Data Source=" & server & ";Initial Catalog=" & database & ";Integrated Security=false;User ID=" & user & ";Password=" & password & ";Pooling=true"
   End Function

   Public Function IsSystemLocked() As Boolean

      Using conn As New SqlConnection(_connectionString)
         conn.Open()
         ' Check if the database is locked.
         Dim cmd = New SqlCommand("sp_ASRLockCheck", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         Dim dr = cmd.ExecuteReader()

         While dr.Read
            ' Not a read-only lock.
            If NullSafeInteger(dr("priority")) <> 3 Then Return True
         End While

         Return False
      End Using
   End Function

   Public Function IsMobileLicensed() As Boolean

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRIntActivateModule", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@sModule", SqlDbType.VarChar, 50).Direction = ParameterDirection.Input
         cmd.Parameters("@sModule").Value = "MOBILE"

         cmd.Parameters.Add("@bLicensed", SqlDbType.Bit).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return NullSafeBoolean(cmd.Parameters("@bLicensed").Value())
      End Using
   End Function

   Public Function CheckLoginDetails(userName As String) As CheckLoginResult

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileCheckLogin", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psKeyParameter").Value = userName

         cmd.Parameters.Add("@piUserGroupID", SqlDbType.Int).Direction = ParameterDirection.Output

         cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Dim result As CheckLoginResult
         result.InvalidReason = NullSafeString(cmd.Parameters("@psMessage").Value())
         result.UserGroupID = NullSafeInteger(cmd.Parameters("@piUserGroupID").Value())
         result.Valid = (result.InvalidReason = Nothing)
         Return result
      End Using
   End Function

   Public Function GetPendingStepCount(userName As String) As Integer

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileCheckPendingWorkflowSteps", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psKeyParameter").Value = userName

         Dim dr As SqlDataReader = cmd.ExecuteReader

         Dim count As Integer
         While dr.Read
            count += 1
         End While
         Return count
      End Using
   End Function

   Public Function GetUserID(email As String) As Integer

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileGetUserIDFromEmail", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psEmail").Value = email

         cmd.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return NullSafeInteger(cmd.Parameters("@piUserID").Value())

      End Using
   End Function

   Public Function Register(email As String) As String

      'Check the email address relates to a user
      Dim userID = GetUserID(email)

      If userID = 0 Then
         Return "No records exist with the given email address."
      End If

      Dim crypt As New Crypt
      Dim encryptedString As String = crypt.EncryptQueryString((userID), -2,
                                                               App.Config.Login,
                                                               App.Config.Password,
                                                               App.Config.Server,
                                                               App.Config.Database,
                                                               "",
                                                               "")

      Dim activationUrl As String = App.Config.WorkflowUrl & "?" & encryptedString

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileRegistration", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psEmailAddress").Value = email

         cmd.Parameters.Add("@psActivationURL", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psActivationURL").Value = activationUrl

         cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return CStr(cmd.Parameters("@psMessage").Value())

      End Using
   End Function

   Public Function ForgotLogin(email As String) As String

      'Check the email address relates to a user
      Dim userID = GetUserID(email)

      If userID = 0 Then
         Return "No records exist with the given email address."
      End If

      'Send it all to sql to validate and email out
      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileForgotLogin", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psEmailAddress").Value = email

         cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return CStr(cmd.Parameters("@psMessage").Value())

      End Using
   End Function

   Public Function GetLoginCount(userName As String) As Integer

      'Does not include being logged into the mobile site
      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetCurrentUsersCountOnServer", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

         cmd.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psLoginName").Value = userName

         cmd.ExecuteNonQuery()

         Return CInt(cmd.Parameters("@iLoginCount").Value)

      End Using
   End Function

   Public Function ActivateUser(userId As Integer) As String

      ' update tbsysMobile_Logins, and copy the 'newpassword' string to the 'password' field

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileActivateUser", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piRecordID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piRecordID").Value = userId

         cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return CStr(cmd.Parameters("@psMessage").Value())

      End Using
   End Function

   Public Function GetWorkflowForm(instanceId As Integer, elementId As Integer) As WorkflowForm

      Dim result As New WorkflowForm With {.Items = New List(Of FormItem)}

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetWorkflowFormItems", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementID").Value = elementId

         cmd.Parameters.Add("@psErrorMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piBackColour", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piBackImage", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piBackImageLocation", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piWidth", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piHeight", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piCompletionMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psCompletionMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piSavedForLaterMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psSavedForLaterMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piFollowOnFormsMessageType", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psFollowOnFormsMessage", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output

         Dim dr = cmd.ExecuteReader

         Do While dr.Read()
            Dim item As New FormItem
            With item
               .Id = NullSafeInteger(dr("Id"))
               .Value = NullSafeString(dr("Value"))
               .ItemType = NullSafeInteger(dr("ItemType"))
               .Caption = NullSafeString(dr("Caption"))
               .InputSize = NullSafeInteger(dr("InputSize"))
               .InputDecimals = NullSafeInteger(dr("InputDecimals"))
               .Left = NullSafeInteger(dr("LeftCoord"))
               .Top = NullSafeInteger(dr("TopCoord"))
               .Width = NullSafeInteger(dr("Width"))
               .Height = NullSafeInteger(dr("Height"))
               .TabIndex = NullSafeShort(dr("TabIndex"))
               .PageNo = NullSafeInteger(dr("PageNo"))
               .PictureID = NullSafeInteger(dr("PictureID"))
               .PictureBorder = NullSafeBoolean(dr("PictureBorder"))
               .FontName = NullSafeString(dr("FontName"))
               .FontSize = NullSafeInteger(dr("FontSize"))
               .FontBold = NullSafeBoolean(dr("FontBold"))
               .FontItalic = NullSafeBoolean(dr("FontItalic"))
               .FontUnderline = NullSafeBoolean(dr("FontUnderline"))
               .FontStrikeThru = NullSafeBoolean(dr("FontStrikeThru"))
               .ForeColor = NullSafeInteger(dr("ForeColor"))
               .BackStyle = NullSafeInteger(dr("BackStyle"))
               .BackColor = NullSafeInteger(dr("BackColor"))
               .LookupFilterColumnName = NullSafeString(dr("LookupFilterColumnName"))
               .LookupFilterColumnDataType = NullSafeInteger(dr("LookupFilterColumnDataType"))
               .LookupFilterOperator = NullSafeInteger(dr("LookupFilterOperator"))
               .LookupFilterValueID = NullSafeString(dr("LookupFilterValueID"))
               .LookupFilterValueType = NullSafeString(dr("LookupFilterValueType"))
               .ColumnHeaders = NullSafeBoolean(dr("ColumnHeaders"))
               .HeadFontSize = NullSafeInteger(dr("HeadFontSize"))
               .HeadLines = NullSafeInteger(dr("Headlines"))
               .HeaderBackColor = NullSafeInteger(dr("HeaderBackColor"))
               .ForeColorEven = NullSafeInteger(dr("ForeColorEven"))
               .ForeColorOdd = NullSafeInteger(dr("ForeColorOdd"))
               .BackColorEven = NullSafeInteger(dr("BackColorEven"))
               .BackColorOdd = NullSafeInteger(dr("BackColorOdd"))
               If Not IsDBNull(dr("ForeColorHighlight")) Then .ForeColorHighlight = NullSafeInteger(dr("ForeColorHighlight"))
               If Not IsDBNull(dr("BackColorHighlight")) Then .BackColorHighlight = NullSafeInteger(dr("BackColorHighlight"))
               .SourceItemType = NullSafeInteger(dr("SourceItemType"))
               .CaptionType = NullSafeInteger(dr("CaptionType"))
               .PasswordType = NullSafeBoolean(dr("PasswordType"))
               .Orientation = NullSafeInteger(dr("Orientation"))
               .Alignment = NullSafeInteger(dr("Alignment"))
            End With
            result.Items.Add(item)
         Loop
         dr.Close()

         result.ErrorMessage = NullSafeString(cmd.Parameters("@psErrorMessage").Value)
         result.BackColour = NullSafeInteger(cmd.Parameters("@piBackColour").Value())
         result.BackImage = NullSafeInteger(cmd.Parameters("@piBackImage").Value())
         result.BackImageLocation = NullSafeInteger(cmd.Parameters("@piBackImageLocation").Value())
         result.Width = NullSafeInteger(cmd.Parameters("@piWidth").Value())
         result.Height = NullSafeInteger(cmd.Parameters("@piHeight").Value())
         result.CompletionMessageType = NullSafeInteger(cmd.Parameters("@piCompletionMessageType").Value())
         result.CompletionMessage = NullSafeString(cmd.Parameters("@psCompletionMessage").Value())
         result.SavedForLaterMessageType = NullSafeInteger(cmd.Parameters("@piSavedForLaterMessageType").Value())
         result.SavedForLaterMessage = NullSafeString(cmd.Parameters("@psSavedForLaterMessage").Value())
         result.FollowOnFormsMessageType = NullSafeInteger(cmd.Parameters("@piFollowOnFormsMessageType").Value())
         result.FollowOnFormsMessage = NullSafeString(cmd.Parameters("@psFollowOnFormsMessage").Value())

         Return result
      End Using
   End Function

   Public Function InstantiateWorkflow(workflowId As Integer, Optional keyParameter As String = "") As InstantiateWorkflowResult

      Dim result As New InstantiateWorkflowResult

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand()
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Connection = conn
         cmd.CommandTimeout = _timeout

         If Not keyParameter = Nothing Then
            cmd.CommandText = "spASRMobileInstantiateWorkflow"
         Else
            cmd.CommandText = "spASRInstantiateWorkflow"
         End If

         cmd.Parameters.Add("@piWorkflowID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piWorkflowID").Value = workflowId

         If Len(keyParameter) > 0 Then
            cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@psKeyParameter").Value = keyParameter

            cmd.Parameters.Add("@psPWDParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@psPWDParameter").Value = ""
         End If

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psFormElements", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         result.InstanceId = NullSafeInteger(cmd.Parameters("@piInstanceID").Value)
         result.FormElements = NullSafeString(cmd.Parameters("@psFormElements").Value)
         result.Message = NullSafeString(cmd.Parameters("@psMessage").Value)

         Return result
      End Using
   End Function

   Public Function GetWorkflowQueryString(instanceId As Integer, [step] As Integer) As String

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetWorkflowQueryString", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementID").Value = [step]

         cmd.Parameters.Add("@psQueryString", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return CStr(cmd.Parameters("@psQueryString").Value())
      End Using
   End Function

   Public Function GetWorkflowItemValues(elementItemID As Integer, instanceId As Integer) As WorkflowItemValuesResult

      Dim result As New WorkflowItemValuesResult

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetWorkflowItemValues", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementItemID").Value = elementItemID

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piLookupColumnIndex", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@piItemType", SqlDbType.Int).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@psDefaultValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

         Dim adapter As New SqlDataAdapter(cmd)
         result.Data = New DataTable()
         adapter.Fill(result.Data)

         result.LookupColumnIndex = NullSafeInteger(cmd.Parameters("@piLookupColumnIndex").Value)
         result.DefaultValue = NullSafeString(cmd.Parameters("@psDefaultValue").Value)

         Return result
      End Using
   End Function

   Public Function GetWorkflowGridItems(elementItemID As Integer, instanceId As Integer) As WorkflowGridItemsResult

      Dim result As New WorkflowGridItemsResult

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetWorkflowGridItems", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementItemID").Value = elementItemID

         cmd.Parameters.Add("@pfOK", SqlDbType.Bit).Direction = ParameterDirection.Output

         Dim adapter As New SqlDataAdapter(cmd)
         result.Data = New DataTable()
         adapter.Fill(result.Data)

         result.Ok = CBool(cmd.Parameters("@pfOK").Value)

         Return result
      End Using
   End Function

   Public Function WorkflowValidateWebForm(elementItemID As Integer, instanceId As Integer, values As String) As ValidateWebFormResult

      Dim result As New ValidateWebFormResult With {.Warnings = New List(Of String), .Errors = New List(Of String)}

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysWorkflowWebFormValidation", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementID").Value = elementItemID

         cmd.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psFormInput1").Value = values

         Dim dr As SqlDataReader = cmd.ExecuteReader

         While dr.Read
            If NullSafeInteger(dr("failureType")) = 0 Then
               result.Errors.Add(NullSafeString(dr("Message")))
            Else
               result.Warnings.Add(NullSafeString(dr("Message")))
            End If
         End While
         dr.Close()

         Return result
      End Using
   End Function

   Public Function WorkflowSubmitWebForm(elementItemID As Integer, instanceId As Integer, values As String,
                                         page As Integer) As SubmitWebFormResult

      Dim result As New SubmitWebFormResult

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSubmitWorkflowStep", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piInstanceID").Value = instanceId

         cmd.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piElementID").Value = elementItemID

         cmd.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psFormInput1").Value = values

         cmd.Parameters.Add("@psFormElements", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
         cmd.Parameters.Add("@pfSavedForLater", SqlDbType.Bit).Direction = ParameterDirection.Output

         cmd.Parameters.Add("@piPageNo", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piPageNo").Value = page

         cmd.ExecuteNonQuery()

         result.FormElements = CStr(cmd.Parameters("@psFormElements").Value())
         result.SavedForLater = CBool(cmd.Parameters("@pfSavedForLater").Value())

         Return result
      End Using
   End Function

   Public Function GetWorkflowCurrentTab(instanceId As Integer) As Integer

      Dim tabPage As Integer

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("SELECT [pageno] FROM [dbo].[ASRSysWorkflowInstances] WHERE [ID] = " & instanceId.ToString, conn)
         Dim dr As SqlDataReader = cmd.ExecuteReader()

         If dr.Read() Then
            tabPage = NullSafeInteger(dr("pageno"))
         End If
         dr.Close()

      End Using

      Return tabPage
   End Function

   Public Function GetSetting(section As String, key As String, userSetting As Boolean) As String

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetSetting", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
         cmd.Parameters("@psSection").Value = section

         cmd.Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
         cmd.Parameters("@psKey").Value = key

         cmd.Parameters.Add("@psDefault", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
         cmd.Parameters("@psDefault").Value = ""

         cmd.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
         cmd.Parameters("@pfUserSetting").Value = userSetting

         cmd.Parameters.Add("@psResult", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

         cmd.ExecuteNonQuery()

         Return CStr(cmd.Parameters("@psResult").Value)

      End Using
   End Function

   Public Function ChangePassword(userName As String, currentPassword As String, newPassword As String) As String

      If GetLoginCount(userName) > 0 Then
         Return "Could not change your password. You are logged into the system using another application."
      End If

      ' Attempt to change the password on the SQL Server.
      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("sp_password", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@old").Value = currentPassword

         cmd.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@new").Value = newPassword

         cmd.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@loginame").Value = userName

         Try
            cmd.ExecuteNonQuery()
         Catch ex As SqlException
            If ex.Number = 15151 Then
               Return "Current password is incorrect."
            Else
               Return ex.Message
            End If
         End Try
      End Using

      ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobilePasswordOK", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@sCurrentUser", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@sCurrentUser").Value = userName

         cmd.ExecuteNonQuery()
      End Using

      Return String.Empty
   End Function

   Public Function CanRunWorkflows(userGroupID As Integer) As Boolean

      Using conn As New SqlConnection(_connectionString)

         ' get the run permissions for workflow for this user group.
         Dim sql As String = "SELECT  [i].[itemKey], [p].[permitted]" &
                             " FROM [ASRSysGroupPermissions] p" &
                             " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" &
                             " WHERE [p].[itemID] IN (" &
                             " SELECT [itemID] FROM [ASRSysPermissionItems]	" &
                             " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " &
                             " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " &
                             userGroupID.ToString & ")"

         conn.Open()
         Dim cmd As New SqlCommand(sql, conn)
         Dim dr As SqlDataReader = cmd.ExecuteReader()

         While dr.Read()
            Select Case CStr(dr("itemKey"))
               Case "RUN"
                  Return CBool(dr("permitted"))
            End Select
         End While

         Return False
      End Using
   End Function

   Public Function GetWorkflowList(userGroupID As Integer) As List(Of WorkflowLink)

      Dim list As New List(Of WorkflowLink)

      Using conn As New SqlConnection(_connectionString)

         Dim sql As String = "SELECT w.Id, w.Name, w.PictureID" &
                             " FROM tbsys_mobilegroupworkflows gw" &
                             " INNER JOIN tbsys_workflows w on gw.WorkflowID = w.ID" &
                             " WHERE gw.UserGroupID = " & userGroupID & " AND w.enabled = 1 ORDER BY gw.Pos ASC"

         conn.Open()
         Dim cmd As New SqlCommand(sql, conn)
         Dim dr As SqlDataReader = cmd.ExecuteReader()

         While dr.Read()
            list.Add(
               New WorkflowLink() With _
                       {.ID = NullSafeInteger(dr("ID")), _
                       .Name = NullSafeString(dr("Name")), _
                       .PictureID = NullSafeInteger(dr("PictureID")) _
                       })
         End While

         Return list
      End Using
   End Function

   Public Function GetPendingStepList(userName As String) As List(Of WorkflowStepLink)

      Dim list As New List(Of WorkflowStepLink)

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRSysMobileCheckPendingWorkflowSteps", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
         cmd.Parameters("@psKeyParameter").Value = userName

         Dim dr As SqlDataReader = cmd.ExecuteReader

         While dr.Read()

            Dim desc As String = CStr(dr("description"))

            If desc.StartsWith(CStr(dr("name")).Trim() & " -") Then
               desc = desc.Substring(CStr(dr("name")).Trim().Length + 2).Trim()
            End If

            list.Add(New WorkflowStepLink() With _
                       {.Url = NullSafeString(dr("Url")),
                       .Name = NullSafeString(dr("Name")),
                       .Desc = desc,
                       .PictureID = NullSafeInteger(dr("PictureID"))
                       })
         End While

         Return (From x In list Order By x.Name, x.Desc).ToList()
      End Using
   End Function

   Public Function GetPicture(id As Integer) As Picture

      Using conn As New SqlConnection(_connectionString)
         conn.Open()

         Dim cmd As New SqlCommand("spASRGetPicture", conn)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.CommandTimeout = _timeout

         cmd.Parameters.Add("@piPictureID", SqlDbType.Int).Direction = ParameterDirection.Input
         cmd.Parameters("@piPictureID").Value = id

         Using reader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)

            Const bufferSize As Integer = 200

            Dim writer As BinaryWriter, stream As New MemoryStream, buffer(bufferSize - 1) As Byte, bytesRead As Long, index As Long

            If reader.Read Then

               Dim picture As New Picture
               picture.Id = id
               picture.Name = NullSafeString(reader("Name"))

               ' Create a file to hold the output.
               writer = New BinaryWriter(stream)

               ' Read bytes into buffer() and retain the number of bytes returned.
               bytesRead = reader.GetBytes(1, index, buffer, 0, bufferSize)

               ' Continue reading and writing while there are bytes beyond the size of the buffer.
               Do While bytesRead = bufferSize
                  writer.Write(buffer)
                  writer.Flush()

                  ' Reposition the start index to the end of the last buffer and fill the buffer.
                  index += bufferSize
                  bytesRead = reader.GetBytes(1, index, buffer, 0, bufferSize)
               Loop

               ' Write the remaining buffer.
               writer.Write(buffer, 0, CInt(bytesRead) - 1)
               writer.Flush()
               writer.Close()

               picture.Image = stream.ToArray()
               stream.Close()

               Return picture
            End If
         End Using
      End Using

      Return Nothing

   End Function

End Class

Public Structure CheckLoginResult
   Public Valid As Boolean
   Public InvalidReason As String
   Public UserGroupID As Integer
End Structure

Public Class InstantiateWorkflowResult
   Public InstanceId As Integer
   Public FormElements As String
   Public Message As String
End Class

Public Class WorkflowItemValuesResult
   Public Data As DataTable
   Public LookupColumnIndex As Integer
   Public DefaultValue As String
End Class

Public Class WorkflowGridItemsResult
   Public Data As DataTable
   Public Ok As Boolean
End Class

Public Class ValidateWebFormResult
   Public Warnings As List(Of String)
   Public Errors As List(Of String)
End Class

Public Class SubmitWebFormResult
   Public FormElements As String
   Public SavedForLater As Boolean
End Class

Public Class WorkflowLink
   Public ID As Integer
   Public Name As String
   Public PictureID As Integer
End Class

Public Class WorkflowStepLink
   Public Url As String
   Public Name As String
   Public Desc As String
   Public PictureID As Integer
End Class

Public Class Picture
   Public Id As Integer
   Public Name As String
   Public Image As Byte()
End Class

Public Class WorkflowForm
   Public ErrorMessage As String
   Public BackColour As Integer
   Public BackImage As Integer
   Public BackImageLocation As Integer
   Public Width As Integer
   Public Height As Integer
   Public CompletionMessageType As Integer
   Public CompletionMessage As String
   Public SavedForLaterMessageType As Integer
   Public SavedForLaterMessage As String
   Public FollowOnFormsMessageType As Integer
   Public FollowOnFormsMessage As String
   Public Items As List(Of FormItem)
End Class

Public Class FormItem
   Public Id As Integer
   Public Value As String
   Public ItemType As Integer
   Public Caption As String
   Public InputSize As Integer
   Public InputDecimals As Integer
   Public Left As Integer
   Public Top As Integer
   Public Width As Integer
   Public Height As Integer
   Public TabIndex As Short
   Public PageNo As Integer
   Public PictureID As Integer
   Public PictureBorder As Boolean
   Public FontName As String
   Public FontSize As Integer
   Public FontBold As Boolean
   Public FontItalic As Boolean
   Public FontUnderline As Boolean
   Public FontStrikeThru As Boolean
   Public ForeColor As Integer
   Public BackStyle As Integer
   Public BackColor As Integer
   Public LookupFilterColumnName As String
   Public LookupFilterColumnDataType As Integer
   Public LookupFilterOperator As Integer
   Public LookupFilterValueID As String
   Public LookupFilterValueType As String
   Public ColumnHeaders As Boolean
   Public HeadFontSize As Integer
   Public HeadLines As Integer
   Public HeaderBackColor As Integer
   Public ForeColorEven As Integer
   Public ForeColorOdd As Integer
   Public BackColorEven As Integer
   Public BackColorOdd As Integer
   Public ForeColorHighlight As Integer?
   Public BackColorHighlight As Integer?
   Public SourceItemType As Integer
   Public CaptionType As Integer
   Public PasswordType As Boolean
   Public Orientation As Integer
   Public Alignment As Integer
End Class
