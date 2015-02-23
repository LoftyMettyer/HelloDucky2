Option Strict Off
Option Explicit On

Imports System.Globalization
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Metadata

Public Class clsDiary
	Inherits BaseForDMI

	Private Const mintVIEWBYDAY As Short = 0
	Private Const mintVIEWBYWEEK As Short = 1
	Private Const mintVIEWBYMONTH As Short = 2
	Private Const mintVIEWBYLIST As Short = 3

	Private Const mstrDATEDAY As String = "ddd dd"
	Private Const mstrDATESQL As String = "MM/dd/yyyy"
	Private Const mstrDATEMONTHYEAR As String = "mmm yyyy"
	Private Const mstrDATEMEDIUM As String = "ddd d mmm yyyy"
	Private Const mstrDATELONG As String = "Long Date" 'Control Panel, Regional Settings

	Private Const mstrSQLOrderByClause As String = " ORDER BY EventDate, DiaryEventsID"

	Private rsTables As New DataTable
	Private rsSingleRecord As New DataTable
	Private mlngRecordCount As Long

	Private mdtSelectedDate As Date
	Private mlngSelectedID As Integer

	Private mstrDateRangeDesc As String
	Private mintCurrentView As Short
	Private mblnEventOwner As Boolean
	Private mblnWriteAccess As Boolean
	Private mblnAlarmAccess As Boolean
	Private mblnFormattingGrid As Boolean
	Private mblnPrinting As Boolean
	Private mstrSystemEvents As String
	Private mlngColumnID As Integer

	Private mintFilterEventType As Short
	Private mintFilterAlarmStatus As Short
	Private mintFilterPastPresent As Short
	Private mblnFilterOnlyMine As Boolean

	Private mblnViewingAlarmedEvents As Boolean

	Public Property ViewingAlarms() As Boolean
		Get
			ViewingAlarms = mblnViewingAlarmedEvents
		End Get
		Set(ByVal Value As Boolean)
			mblnViewingAlarmedEvents = Value
		End Set
	End Property

	Public ReadOnly Property CurrentView() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object CurrentView. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CurrentView = mintCurrentView
		End Get
	End Property

	Public ReadOnly Property AllowSystemEvents() As Boolean
		Get
			Return True
		End Get
	End Property


	Public Property FormattingGrid() As Boolean
		Get
			FormattingGrid = mblnFormattingGrid
		End Get
		Set(ByVal Value As Boolean)
			mblnFormattingGrid = Value
		End Set
	End Property


	Public Property Printing() As Boolean
		Get
			Printing = mblnPrinting
		End Get
		Set(ByVal Value As Boolean)
			mblnPrinting = Value
		End Set
	End Property

	Public ReadOnly Property EventOwner() As Boolean
		Get
			EventOwner = mblnEventOwner
		End Get
	End Property

	Public ReadOnly Property WriteAccess() As Boolean
		Get
			WriteAccess = mblnWriteAccess
		End Get
	End Property

	Public ReadOnly Property AlarmAccess() As Boolean
		Get
			AlarmAccess = mblnAlarmAccess
		End Get
	End Property

	Public ReadOnly Property GetDateRangeDesc() As String
		Get
			GetDateRangeDesc = mstrDateRangeDesc
		End Get
	End Property


	Public Property DiaryEventID() As Integer
		Get
			DiaryEventID = mlngSelectedID
		End Get
		Set(ByVal Value As Integer)

			Dim strOwner As String
			Dim strUser As String
			Dim strAccess As String

			mlngSelectedID = Value

			If Value > 0 Then
				rsSingleRecord = GetCurrentRecord()
				If rsSingleRecord.Rows.Count = 0 Then
					' COAMsgBox "This diary event has been deleted by another user.", vbCritical, "Diary"
					mlngSelectedID = 0
					RefreshDiaryData()
				End If
			End If


			If mlngSelectedID > 0 Then

				Dim rowSingleRecord = rsSingleRecord.Rows(0)

				mdtSelectedDate = rowSingleRecord("EventDate")

				strOwner = LCase(Trim(rowSingleRecord("UserName").ToString()))
				strUser = LCase(_login.Username)
				strAccess = rowSingleRecord("Access").ToString()
				mlngColumnID = CShort(rowSingleRecord("ColumnID"))

				mblnEventOwner = (strOwner = strUser)

				'Write Access  if eventowner or not read only and not system event
				mblnWriteAccess = ((mblnEventOwner Or strAccess = "RW") And mlngColumnID = 0)

				'Alarm Access if you have write access or its a system event in the past
				mblnAlarmAccess = (mblnWriteAccess Or (mlngColumnID > 0 And CDate(rowSingleRecord("EventDate")) < Now))
			Else
				mblnEventOwner = False
				mblnWriteAccess = False
				mblnAlarmAccess = False

			End If

		End Set
	End Property



	Public Property DateSelected() As Date
		Get
			DateSelected = mdtSelectedDate
		End Get
		Set(ByVal Value As Date)
			mdtSelectedDate = Value
		End Set
	End Property

	Public Property FilterEventType() As Short
		Get
			FilterEventType = mintFilterEventType
		End Get
		Set(ByVal Value As Short)
			'0 - Both System and Manual Events
			'1 - System Events
			'2 - Manual Events
			'3 - Manual Events where current user is owner
			If Value < 2 And Not AllowSystemEvents Then
				Value = 2
			End If
			mintFilterEventType = Value
		End Set
	End Property

	Public Property FilterAlarmStatus() As Short
		Get
			FilterAlarmStatus = mintFilterAlarmStatus
		End Get
		Set(ByVal Value As Short)
			'0 - Both Alarmed and Non-Alarmed
			'1 - Alarmed Only
			'2 - Non-Alarmed Only
			mintFilterAlarmStatus = Value
		End Set
	End Property

	Public Property FilterPastPresent() As Short
		Get
			FilterPastPresent = mintFilterPastPresent
		End Get
		Set(ByVal Value As Short)
			'0 - All
			'1 - Only Past Events
			'2 - Only Future Events
			mintFilterPastPresent = Value
		End Set
	End Property

	Public Property FilterOnlyMine() As Boolean
		Get
			FilterOnlyMine = mblnFilterOnlyMine
		End Get
		Set(ByVal Value As Boolean)
			mblnFilterOnlyMine = Value
		End Set
	End Property

	Private Function SQLWhereCurrentEvent() As String
		SQLWhereCurrentEvent = "WHERE DiaryEventsID = " & CStr(mlngSelectedID)
	End Function

	Public Function SQLCurrentDateTime() As String
		SQLCurrentDateTime = Replace(VB6.Format(Now, "MM/dd/yyyy hh"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & ":" & VB6.Format(Now, "nn")
	End Function

	Private Function SQLFilter() As String

		Dim strManualEvents As String

		strManualEvents = "ColumnID = 0 AND (UserName = '" & _login.Username & "'" & IIf(mblnFilterOnlyMine, ")", " OR Access <> 'HD')")

		Select Case FilterEventType
			Case 0 'All Manual and System events
				SQLFilter = "(" & strManualEvents & IIf(mstrSystemEvents <> vbNullString, " OR  " & mstrSystemEvents, "") & ")"

			Case 1 'System events
				SQLFilter = mstrSystemEvents

			Case 2 'Manual events
				SQLFilter = strManualEvents

		End Select

		If mintFilterPastPresent > 0 Then
			SQLFilter = SQLFilter & " AND DateDiff(n, '" & SQLCurrentDateTime() & "', EventDate) " & IIf(mintFilterPastPresent = 1, "< 0", ">= 0")
		End If

		If mintFilterAlarmStatus > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SQLFilter = SQLFilter & Choose(mintFilterAlarmStatus, " AND Alarm = 1", " AND Alarm = 0")
		End If

	End Function

	Public Sub ChangeView(ByRef intNewViewMode As Short)

		'  Dim blnListView As Boolean

		mintCurrentView = intNewViewMode

		'  blnListView = (intNewViewMode = mintVIEWBYLIST)
		'
		'  With frmDiary.SSActiveToolBars1
		'
		'    'Do not allow previous, next or print
		'    'for list view but do allow up and down
		'    .Redraw = False
		'    .Tools(1).Enabled = Not blnListView   'Previous
		'    .Tools(2).Enabled = Not blnListView   'Next
		'    .Tools(10).Enabled = Not blnListView  'Print
		'
		'    'Only Allow 'Alarm' for day view
		'    .Tools(14).Enabled = Not blnListView       'Alarm
		'    .Redraw = True
		'
		'  End With

	End Sub

	Public Sub RefreshDiaryData()

		Dim dtStartDate As Date
		Dim dtEndDate As Date

		Call GetDateRange(dtStartDate, dtEndDate)	'This sets up these two variables
		Call GetDiaryData(False, dtStartDate, dtEndDate)

	End Sub

	Private Sub GetDateRange(ByRef dtStartDate As Date, ByRef dtEndDate As Date)

		'  With frmDiary
		'
		'    Select Case mintCurrentView
		'    Case mintVIEWBYDAY
		'Range Start: Today
		'Range End:   Today
		dtStartDate = mdtSelectedDate
		dtEndDate = mdtSelectedDate

		mstrDateRangeDesc = VB6.Format(mdtSelectedDate, mstrDATELONG)
		' frmDiary.lblSelectedDay = mstrDateRangeDesc

		'    Case mintVIEWBYWEEK
		'      'Range Start: Monday just passed
		'      'Range End:   One week minus one day
		'      intOffSet = (Weekday(mdtSelectedDate, vbMonday) - 1) * -1
		'      dtStartDate = DateAdd("d", intOffSet, mdtSelectedDate)
		'      dtEndDate = DateAdd("d", 6, dtStartDate)
		'
		'      strStartMonth = Format(dtStartDate, mstrDATEMONTHYEAR)
		'      strEndMonth = Format(dtEndDate, mstrDATEMONTHYEAR)
		'
		'      frmDiary.lblWeekNo = "Week " & .mvwViewbyMonth.Week
		'      If strStartMonth = strEndMonth Then
		'        frmDiary.lblSelectedWeek = strStartMonth
		'      Else
		'        frmDiary.lblSelectedWeek = strStartMonth & " - " & strEndMonth
		'      End If
		'
		'      mstrDateRangeDesc = frmDiary.lblWeekNo & " (" & _
		''                          Format(dtStartDate, DateFormat) & " - " & _
		''                          Format(dtEndDate, DateFormat) & ")"
		'
		'    Case mintVIEWBYMONTH
		'      'Range Start: First Day of current month
		'      'Range End:   Six months later minus one day
		'      dtStartDate = frmDiary.mvwViewbyMonth.VisibleDays(1)
		'      Do While Day(dtStartDate) > 1
		'        dtStartDate = DateAdd("d", 1, dtStartDate)
		'      Loop
		'      dtEndDate = DateAdd("d", -1, DateAdd("m", 6, dtStartDate))
		'
		'      mstrDateRangeDesc = Format(dtStartDate, DateFormat) & " - " & _
		''                          Format(dtEndDate, DateFormat)
		'
		'    Case mintVIEWBYLIST
		'      dtStartDate = frmDiary.mvwViewbyMonth.MinDate
		'      dtEndDate = frmDiary.mvwViewbyMonth.MaxDate
		'      mstrDateRangeDesc = "(All Dates)"
		'
		'    End Select

		'End With

	End Sub


	Public Function GetRecordCount() As Long

		GetDiaryData(True)
		Return mlngRecordCount

	End Function

	Public Function GetDiaryData(ByRef blnOnlyCountFilterMatch As Boolean, Optional ByRef dtStartDate As Date = #12:00:00 AM#, Optional ByRef dtEndDate As Date = #12:00:00 AM#) As DataTable

		Dim strSQL As String
		Dim strSQLSelect As String
		Dim strSQLFrom As String
		Dim strSQLWhere As String
		Dim strSQLOrderBy As String

		Dim blnClearedFilter As Boolean

		'*******************************
		'* SELECT AND ORDER BY CLAUSES *
		'*******************************
		If mintCurrentView = mintVIEWBYMONTH And Not mblnPrinting Then
			'If month view then only check a maximum of one item per date
			'(but need to retrieve all the data if printing!)
			strSQLSelect = "SELECT DISTINCT convert(datetime,convert(varchar(10),EventDate,101)) as 'EventDate'"
			strSQLOrderBy = "ORDER BY EventDate"

		Else
			strSQLSelect = "SELECT DiaryEventsID, Alarm, EventDate, EventTitle, Convert(varchar,EventDate,112) as FindDate "
			strSQLOrderBy = "ORDER BY EventTitle"	' mstrSQLOrderByClause

		End If
		strSQLFrom = "FROM ASRSysDiaryEvents "


		'Check record count (before applying any date range !)
		Do
			blnClearedFilter = False

			strSQL = "SELECT COUNT(*) " & strSQLFrom & "WHERE " & SQLFilter()

			rsTables = DB.GetDataTable(strSQL, CommandType.Text)
			mlngRecordCount = CLng(rsTables.Rows(0)(0))

		Loop While blnClearedFilter = True


		'If not only counting records in filter then retreive actual data
		If Not blnOnlyCountFilterMatch Then

			strSQLWhere = "WHERE EventDate BETWEEN " & "'" & Replace(VB6.Format(dtStartDate, mstrDATESQL), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & " 00:00' AND " & "'" & Replace(VB6.Format(dtEndDate, mstrDATESQL), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & " 23:59' AND "

			strSQL = strSQLSelect & strSQLFrom & strSQLWhere & SQLFilter() & strSQLOrderBy
			Return DB.GetDataTable(strSQL, CommandType.Text)
		End If

	End Function

	Public Function GetCurrentRecord() As DataTable
		Return DB.GetDataTable("SELECT ASRSysDiaryEvents.*, CONVERT(integer,ASRSysDiaryEvents.TimeStamp) AS intTimeStamp FROM ASRSysDiaryEvents " & SQLWhereCurrentEvent(), CommandType.Text)
	End Function

	Public Function FilterText() As String

		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(mintFilterAlarmStatus + 1, , Alarmed , Non-Alarmed ). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FilterText = Choose(mintFilterPastPresent + 1, "", "Past ", "Current and Future ") & Choose(mintFilterAlarmStatus + 1, "", "Alarmed ", "Non-Alarmed ")

		Select Case FilterEventType
			Case 0
				FilterText = FilterText & "System Events and Manual Events" & IIf(mblnFilterOnlyMine, " where owner is '" & _login.Username & "'", "")
			Case 1
				FilterText = FilterText & "System Diary Events"
			Case 2
				FilterText = FilterText & "Manual Diary Events" & IIf(mblnFilterOnlyMine, " where owner is '" & _login.Username & "'", "")
		End Select

	End Function

	Public Function GetAlarmCount(ByRef intAlarmSetting As Short) As Long

		Dim rsAlarmCount As DataTable
		Dim strSQL As String
		Dim dtStartDate As Date
		Dim dtEndDate As Date
		Dim strEventDateMatch As String

		If mintCurrentView <> mintVIEWBYLIST Then
			Call GetDateRange(dtStartDate, dtEndDate)	'This sets up these two variables
			strEventDateMatch = "(EventDate BETWEEN " & "'" & Replace(VB6.Format(dtStartDate, mstrDATESQL), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & " 00:00' AND " & "'" & Replace(VB6.Format(dtEndDate, mstrDATESQL), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & " 23:59') AND "
		End If

		strSQL = "SELECT COUNT(DiaryEventsID) AS recCount FROM ASRSysDiaryEvents WHERE " & strEventDateMatch & "Alarm = '" & CStr(intAlarmSetting) & "' AND " & SQLFilter()

		'Set rsAlarmCount = datGeneral.GetRecords(strSQL)
		rsAlarmCount = DB.GetDataTable(strSQL, CommandType.Text)

		Return CLng(rsAlarmCount.Rows(0)("recCount"))

	End Function


	Public Function ShowAlarmedEvents(intFilterPastPresent As Short, intViewMode As Short) As Boolean

		Try

			Me.ViewingAlarms = True	'This need to be done here in case
			'you don't view alarms (filter will be saved!)

			'Set filter to all alarmed events in the past
			Me.FilterEventType = 0 'All
			Me.FilterAlarmStatus = 1 'Alarmed
			Me.FilterPastPresent = intFilterPastPresent
			'Me.DateSelected = Now
			mintCurrentView = intViewMode

			If GetRecordCount() > 0 Then

				'Set filter to all alarmed events in the past
				'(These seem to be getting reset in the exe so ensure that they are set)
				Me.FilterEventType = 0 'All
				Me.FilterPastPresent = intFilterPastPresent
				Me.FilterAlarmStatus = 1 'Alarmed
				Me.DateSelected = Now
				mintCurrentView = intViewMode

				'This is only if you a diary event pops up during the day
				If mintCurrentView = mintVIEWBYDAY Then
					Me.FilterPastPresent = 0 'All (incase the minute has ticked over).
				End If

			End If

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Public Sub CheckAccessToSystemEvents()

		Dim rsTemp As DataTable
		Dim strSQL As String

		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim strColList As String

		strSQL = "SELECT DISTINCT t.TableID, t.TableName, c.ColumnID, c.ColumnName FROM ASRSysDiaryLinks d JOIN ASRSysColumns c ON d.ColumnID = c.ColumnID JOIN ASRSysTables t ON c.TableID = t.TableID"
		rsTemp = DB.GetDataTable(strSQL, CommandType.Text)

		If rsTemp.Rows.Count = 0 Then
			'Can't find any links !!!
			mstrSystemEvents = "ColumnID = -1" 'MH20061010 Fault 11566
			Exit Sub
		End If

		For Each objTableView In gcoTablePrivileges.Collection

			strColList = vbNullString
			If (objTableView.AllowSelect) Then

				For Each objRow As DataRow In rsTemp.Rows

					'Loop thru all of the views for this table where the user has select access
					If (objTableView.TableID = CInt(objRow("TableID"))) Then
						objColumnPrivileges = gcolColumnPrivilegesCollection.Item(IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName))


						If objColumnPrivileges.IsValid(objRow("ColumnName").ToString()) Then
							If objColumnPrivileges.Item(objRow("ColumnName").ToString()).AllowSelect Then
								strColList = strColList & IIf(strColList <> vbNullString, ", ", "") & objRow("ColumnID").ToString()
							End If
						End If

					End If

				Next

				If strColList <> vbNullString Then
					mstrSystemEvents = mstrSystemEvents & IIf(mstrSystemEvents <> vbNullString, " OR " & vbCrLf, "") & "(RowID IN (SELECT ID FROM " & objTableView.RealSource & ") AND (ColumnID IN (" & strColList & ")))"
				End If

			End If
		Next objTableView

		If mstrSystemEvents <> vbNullString Then
			mstrSystemEvents = "(" & mstrSystemEvents & ")"
		End If

	End Sub


End Class