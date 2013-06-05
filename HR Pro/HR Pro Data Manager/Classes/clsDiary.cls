VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'THIS SQL WILL TAKE THE EVENTTIME AND INCORPORATE IT INTO THE EVENTDATE
'----------------------------------------------------------------------
'select convert(datetime,convert(varchar(10),eventdate,111)+' '+eventtime)
'From asrsysdiaryevents
'order by eventdate, eventtime


Option Explicit

Private Const mintVIEWBYDAY As Integer = 0
Private Const mintVIEWBYWEEK As Integer = 1
Private Const mintVIEWBYMONTH As Integer = 2
Private Const mintVIEWBYLIST As Integer = 3

Private Const mstrDATEDAY = "ddd dd"
Private Const mstrDATESQL = "mm/dd/yyyy"
Private Const mstrDATEMONTHYEAR = "mmm yyyy"
Private Const mstrDATEMEDIUM = "ddd d mmm yyyy"
Private Const mstrDATELONG = "Long Date"    'Control Panel, Regional Settings

Private Const mstrSQLOrderByClause As String = _
    " ORDER BY EventDate, DiaryEventsID"

Private rsTables As New ADODB.Recordset
Private rsSingleRecord As New ADODB.Recordset
Private datData As clsDataAccess
Private mlngRecordCount As Long

Private mdtSelectedDate As Date
Private mlngSelectedID As Long

Private mstrDateRangeDesc As String
Private mintCurrentView As Integer
Private mblnEventOwner As Boolean
Private mblnWriteAccess As Boolean
Private mblnAlarmAccess As Boolean
Private mblnFormattingGrid As Boolean
Private mblnPrinting As Boolean
Private mstrSystemEvents As String
Private mobjClipBoard As Collection
Private mlngColumnID As Long

Private mintFilterEventType As Integer
Private mintFilterAlarmStatus As Integer
Private mintFilterPastPresent As Integer
Private mblnFilterOnlyMine As Boolean

Private mlngIconId As Picture

Private mblnViewingAlarmedEvents As Boolean

Public Property Get ViewingAlarms() As Boolean
  ViewingAlarms = mblnViewingAlarmedEvents
End Property
Public Property Let ViewingAlarms(ByVal blnNewValue As Boolean)
  mblnViewingAlarmedEvents = blnNewValue
End Property

Private Function SQLWhereCurrentEvent() As String
  SQLWhereCurrentEvent = "WHERE DiaryEventsID = " & CStr(mlngSelectedID)
End Function

Public Property Get CurrentView() As Variant
  CurrentView = mintCurrentView
End Property

Public Property Get AllowSystemEvents() As Boolean
  AllowSystemEvents = datGeneral.SystemPermission("DIARY", "SYSTEMEVENTS")
End Property

Public Function SQLCurrentDateTime() As String
  SQLCurrentDateTime = Replace(Format(Now, "mm/dd/yyyy hh"), UI.GetSystemDateSeparator, "/") & ":" & Format(Now, "nn")
End Function


'Public Property Get RecordCount() As Long
'  RecordCount = mlngRecordCount
'End Property

Private Function SQLFilter() As String
  
  Dim strManualEvents As String
  
  strManualEvents = _
    "ColumnID = 0 AND (UserName = '" & datGeneral.UserNameForSQL & "'" & _
    IIf(mblnFilterOnlyMine, ")", " OR Access <> 'HD')")
  
  Select Case FilterEventType
  Case 0    'All Manual and System events
    SQLFilter = "(" & strManualEvents & _
             IIf(mstrSystemEvents <> vbNullString, " OR  " & mstrSystemEvents, "") & ")"

  Case 1    'System events
    SQLFilter = mstrSystemEvents
    
  Case 2    'Manual events
    SQLFilter = strManualEvents

  End Select

  If mintFilterPastPresent > 0 Then
    SQLFilter = SQLFilter & _
      " AND DateDiff(n, '" & SQLCurrentDateTime & "', EventDate) " & _
      IIf(mintFilterPastPresent = 1, "< 0", ">= 0")
  End If

  If mintFilterAlarmStatus > 0 Then
    SQLFilter = SQLFilter & _
      Choose(mintFilterAlarmStatus, " AND Alarm = 1", " AND Alarm = 0")
  End If

End Function

Public Property Get FormattingGrid() As Boolean
  FormattingGrid = mblnFormattingGrid
End Property

Public Property Let FormattingGrid(ByVal blnNewValue As Boolean)
  mblnFormattingGrid = blnNewValue
End Property

Public Property Get Printing() As Boolean
  Printing = mblnPrinting
End Property

Public Property Let Printing(ByVal blnNewValue As Boolean)
  mblnPrinting = blnNewValue
End Property

Public Property Get EventOwner() As Boolean
  EventOwner = mblnEventOwner
End Property

Public Property Get WriteAccess() As Boolean
  WriteAccess = mblnWriteAccess
End Property

Public Property Get AlarmAccess() As Boolean
  AlarmAccess = mblnAlarmAccess
End Property

Public Property Get GetDateRangeDesc() As String
  GetDateRangeDesc = mstrDateRangeDesc
End Property

Public Property Get DiaryEventID() As Long
  DiaryEventID = mlngSelectedID
End Property

Public Property Let DiaryEventID(ByVal vNewID As Long)
  
  Dim strOwner As String
  Dim strUser As String
  Dim strAccess As String
  Dim blnEnabled As String
  
  mlngSelectedID = vNewID
  
  If vNewID > 0 Then
    Set rsSingleRecord = GetCurrentRecord
    If rsSingleRecord.BOF And rsSingleRecord.EOF Then
      COAMsgBox "This diary event has been deleted by another user.", vbCritical, "Diary"
      mlngSelectedID = 0
      gobjDiary.RefreshDiaryData
    End If
  End If

  
  If mlngSelectedID > 0 Then
    
    mdtSelectedDate = rsSingleRecord.Fields("EventDate").Value

    strOwner = LCase(Trim$(rsSingleRecord.Fields("UserName").Value))
    strUser = LCase(gsUserName)
    strAccess = rsSingleRecord.Fields("Access").Value
    mlngColumnID = CInt(rsSingleRecord.Fields("ColumnID").Value)
    
    mblnEventOwner = (strOwner = strUser)
    
    'Write Access  if eventowner or not read only and not system event
    mblnWriteAccess = ((mblnEventOwner Or strAccess = "RW") And mlngColumnID = 0)
    
    'Alarm Access if you have write access or its a system event in the past
    mblnAlarmAccess = (mblnWriteAccess Or _
                      (mlngColumnID > 0 And rsSingleRecord.Fields("EventDate").Value < Now))
  Else
    mblnEventOwner = False
    mblnWriteAccess = False
    mblnAlarmAccess = False

  End If
  
  
  '02/08/2001 MH Fault 2066
  'With frmDiary.SSActiveToolBars1
  '  '.Tools(8).Enabled = True  'mblnWriteAccess     'Delete
  '  .Tools(8).Enabled = (mlngRecordCount > 0)     'Delete
  '  .Tools(21).Enabled = (mlngSelectedID > 0 And lngColumnID = 0)  'Duplicate
  '  .Tools(11).Enabled = (mlngSelectedID > 0)    'Edit
  'End With
  'With frmDiary.DiaryToolBar
  '  .Item("Delete").Enabled = (mlngRecordCount > 0)
  '
  '  blnEnabled = (mlngSelectedID > 0 And mlngRecordCount > 0 And mintCurrentView <> mintVIEWBYMONTH)
  '
  '  .Item("Edit").Enabled = blnEnabled
  '  .Item("Repeat").Enabled = blnEnabled
  '  .Item("Cut").Enabled = blnEnabled
  '  .Item("Copy").Enabled = blnEnabled
  '
  '  .Item("Paste").Enabled = (mintCurrentView <> mintVIEWBYMONTH And Not (mobjClipBoard Is Nothing))
  '
  'End With
  EnableTools

End Property

Public Sub ChangeView(intNewViewMode As Integer)
  
'  Dim blnListView As Boolean
  
  mintCurrentView = intNewViewMode
  EnableTools

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

Public Property Get DateSelected() As Date
  DateSelected = mdtSelectedDate
End Property


Public Property Let DateSelected(ByVal vNewDate As Date)
  
  Dim intCount As Integer
  Dim dtStartDate As Date
  Dim dtEndDate As Date

  mdtSelectedDate = vNewDate
  
  With frmDiary
    
    .mvwViewbyMonth.Value = vNewDate
    
    If mintCurrentView = mintVIEWBYWEEK Then
      For intCount = 0 To 6
        If intCount = Weekday(mdtSelectedDate, vbMonday) - 1 Then
          .lblDayTitle(intCount).ForeColor = vbRed
        Else
          .lblDayTitle(intCount).ForeColor = vbBlack
        End If
      Next
    End If

    EnableTools
'    Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
'
'    'Check if enable "previous"
'    .SSActiveToolBars1.Tools(1).Enabled = (DateDiff("d", dtStartDate, .mvwViewbyMonth.MinDate) < 0)
'
'    'Check if enable "next"
'    .SSActiveToolBars1.Tools(2).Enabled = (DateDiff("d", dtEndDate, .mvwViewbyMonth.MaxDate) > 0)
  
  End With

End Property


Public Sub MoveDate(intMovement As Integer)

  Dim dtNewDate As Date
  
  'This redraw will be turned on again after
  'the data has been refreshed
  frmDiary.AutoRedraw = False
  
  Select Case mintCurrentView
  Case mintVIEWBYDAY
    dtNewDate = DateAdd("d", intMovement, mdtSelectedDate)
  Case mintVIEWBYWEEK
    dtNewDate = DateAdd("ww", intMovement, mdtSelectedDate)
  Case mintVIEWBYMONTH
    'Check to see the scroll rate specified on the month control
    intMovement = intMovement * Val(frmDiary.mvwViewbyMonth.ScrollRate)
    dtNewDate = DateAdd("m", intMovement, mdtSelectedDate)
  End Select

  'Setting the diary date can can cause an error if date is outside
  'monthview.mindate and monthview.maxdate
  On Local Error GoTo LocalErr
  Me.DateSelected = dtNewDate
  Me.RefreshDiaryData

Exit Sub

LocalErr:
  With frmDiary.mvwViewbyMonth
    If DateDiff("d", .MaxDate, dtNewDate) > 0 Then
      Me.DateSelected = .MaxDate
    
    ElseIf DateDiff("d", .MinDate, dtNewDate) < 0 Then
      Me.DateSelected = .MinDate
    
    Else
      COAMsgBox "Error Changing Date" & vbCrLf & "(" & Err.Description & ")", vbCritical
      Me.DateSelected = Now
    
    End If

  End With

  Resume Next

End Sub


Public Sub RefreshDiaryData()
  
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  
  Screen.MousePointer = vbHourglass
  frmDiary.AutoRedraw = False
  DoEvents

  Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
  Call GetDiaryData(False, dtStartDate, dtEndDate)
  Call PopulateDiaryData(dtStartDate, dtEndDate)
  
  frmDiary.AutoRedraw = True
  Screen.MousePointer = vbDefault

End Sub


Private Sub GetDateRange(dtStartDate As Date, dtEndDate As Date)

  Dim strStartMonth As String
  Dim strEndMonth As String
  Dim intOffSet As Integer

  With frmDiary

    Select Case mintCurrentView
    Case mintVIEWBYDAY
      'Range Start: Today
      'Range End:   Today
      dtStartDate = mdtSelectedDate
      dtEndDate = mdtSelectedDate

      mstrDateRangeDesc = Format(mdtSelectedDate, mstrDATELONG)
      frmDiary.lblSelectedDay = mstrDateRangeDesc
    
    Case mintVIEWBYWEEK
      'Range Start: Monday just passed
      'Range End:   One week minus one day
      intOffSet = (Weekday(mdtSelectedDate, vbMonday) - 1) * -1
      dtStartDate = DateAdd("d", intOffSet, mdtSelectedDate)
      dtEndDate = DateAdd("d", 6, dtStartDate)

      strStartMonth = Format(dtStartDate, mstrDATEMONTHYEAR)
      strEndMonth = Format(dtEndDate, mstrDATEMONTHYEAR)

      frmDiary.lblWeekNo = "Week " & .mvwViewbyMonth.Week
      If strStartMonth = strEndMonth Then
        frmDiary.lblSelectedWeek = strStartMonth
      Else
        frmDiary.lblSelectedWeek = strStartMonth & " - " & strEndMonth
      End If
  
      mstrDateRangeDesc = frmDiary.lblWeekNo & " (" & _
                          Format(dtStartDate, DateFormat) & " - " & _
                          Format(dtEndDate, DateFormat) & ")"
    
    Case mintVIEWBYMONTH
      'Range Start: First Day of current month
      'Range End:   Six months later minus one day
      dtStartDate = frmDiary.mvwViewbyMonth.VisibleDays(1)
      Do While Day(dtStartDate) > 1
        dtStartDate = DateAdd("d", 1, dtStartDate)
      Loop
      dtEndDate = DateAdd("d", -1, DateAdd("m", 6, dtStartDate))

      mstrDateRangeDesc = Format(dtStartDate, DateFormat) & " - " & _
                          Format(dtEndDate, DateFormat)

    Case mintVIEWBYLIST
      dtStartDate = frmDiary.mvwViewbyMonth.MinDate
      dtEndDate = frmDiary.mvwViewbyMonth.MaxDate
      mstrDateRangeDesc = "(All Dates)"
  
    End Select

  End With

End Sub


Public Function GetRecordCount() As Long

  Dim dtStartDate As Date
  Dim dtEndDate As Date

  Screen.MousePointer = vbHourglass
  Call GetDiaryData(True)
  Screen.MousePointer = vbDefault

  GetRecordCount = mlngRecordCount

End Function

Private Sub GetDiaryData(blnOnlyCountFilterMatch As Boolean, Optional dtStartDate As Date, Optional dtEndDate As Date)
  
  Dim strSQL As String
  Dim strSQLSelect As String
  Dim strSQLFrom As String
  Dim strSQLWhere As String
  Dim strSQLOrderBy As String
  
  Dim blnListView As Boolean
  Dim blnClearedFilter As Boolean
  
  blnListView = (mintCurrentView = mintVIEWBYLIST)
  
  
  '*******************************
  '* SELECT AND ORDER BY CLAUSES *
  '*******************************
  If mintCurrentView = mintVIEWBYMONTH And Not mblnPrinting Then
    'If month view then only check a maximum of one item per date
    '(but need to retrieve all the data if printing!)
    'strSQLSelect = "SELECT DISTINCT EventDate "
    strSQLSelect = "SELECT DISTINCT convert(datetime,convert(varchar(10),EventDate,101)) as 'EventDate'"
    strSQLOrderBy = "ORDER BY EventDate"
  
  Else
    strSQLSelect = "SELECT DiaryEventsID, Alarm, EventDate, EventTitle, " & _
                   "Convert(varchar,EventDate,112) as FindDate "
    strSQLOrderBy = mstrSQLOrderByClause
  
  End If
  strSQLFrom = "FROM ASRSysDiaryEvents "


  'Check record count (before applying any date range !)
  Do
    blnClearedFilter = False
    
    strSQL = "SELECT COUNT(*) " & strSQLFrom & "WHERE " & SQLFilter
    Set rsTables = datData.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)
    mlngRecordCount = rsTables(0)

    'MH20020701 Fault 4088
    'If mlngRecordCount = 0 And Not blnOnlyCountFilterMatch And _
    '  (mintFilterEventType > 0 Or mintFilterPastPresent > 0 Or mintFilterAlarmStatus > 0 Or mblnFilterOnlyMine)) Then
    
    If mlngRecordCount = 0 And Not blnOnlyCountFilterMatch And _
      ((FilterEventType > 0 And AllowSystemEvents) Or _
       (mintFilterPastPresent > 0 Or mintFilterAlarmStatus > 0 Or mblnFilterOnlyMine)) Then
      'Filter applied and no records match (and not only counting match for alarm!)

        If gobjDiary.ViewingAlarms = True Then
          COAMsgBox "All alarms on past events have now been cleared.", vbInformation + vbOKOnly, "Diary"
          If frmDiaryDetail.Visible Then frmDiaryDetail.Hide
          frmDiary.Hide
        Else
          FilterEventType = 0
          FilterPastPresent = 0
          FilterAlarmStatus = 0
          FilterOnlyMine = False

          COAMsgBox "No records match the current filter." & vbCrLf & _
                 "No filter is applied.", vbInformation + vbOKOnly, "Diary Filter"
          frmDiary.Caption = Me.FilterText
          blnClearedFilter = True
        End If

    End If

  Loop While blnClearedFilter = True


  'If not only counting records in filter then retreive actual data
  If Not blnOnlyCountFilterMatch Then

    strSQLWhere = "WHERE EventDate BETWEEN " & _
                  "'" & Replace(Format(dtStartDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 00:00' AND " & _
                  "'" & Replace(Format(dtEndDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 23:59' AND "

    strSQL = strSQLSelect & strSQLFrom & strSQLWhere & SQLFilter & strSQLOrderBy
    Set rsTables = datData.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)
  End If

End Sub


Private Sub PopulateDiaryData(dtStartDate As Date, dtEndDate As Date)
  
  Dim blnNoRecordsInView As Boolean
  
  Select Case mintCurrentView
  Case mintVIEWBYDAY
    Call PopulateByDayView
  Case mintVIEWBYWEEK
    Call PopulateByWeekView(dtStartDate)
  Case mintVIEWBYMONTH
    Call PopulateByMonthView(dtStartDate, dtEndDate)
  Case mintVIEWBYLIST
    Call PopulateByListView
  End Select
  
  EnableTools

End Sub


Private Sub Class_Initialize()
  Set datData = New HRProDataMgr.clsDataAccess
  mdtSelectedDate = Now

  CheckAccessToSystemEvents

End Sub


Private Sub PopulateByDayView()

  Dim strOutput As String
  Dim intCount As Integer

  With rsTables

    frmDiary.grdViewByDay.Redraw = False
    frmDiary.grdViewByDay.RemoveAll
    If Not .EOF Then
      .MoveFirst
    
      Do While Not .EOF
      
        strOutput = .Fields("DiaryEventsID").Value & vbTab & _
                    .Fields("Alarm").Value & vbTab & _
                    Format(.Fields("EventDate").Value, "hh:nn") & vbTab & _
                    .Fields("EventTitle").Value
        frmDiary.grdViewByDay.AddItem strOutput
      
        If .Fields(0).Value = mlngSelectedID Then
          With frmDiary.grdViewByDay
            .Redraw = True
            .SelBookmarks.Add .GetBookmark(.Rows - 1)
            .Redraw = False
          End With
        End If

        .MoveNext
      Loop
    End If
    
  End With

  With frmDiary.grdViewByDay
  
    .Redraw = True
    If .Rows > 17 Then
      .ScrollBars = ssScrollBarsVertical
      '.Columns("EventTitle").Width = 6770
      .Columns("EventTitle").Width = .Width - 1085
    Else
      .ScrollBars = ssScrollBarsNone
      '.Columns("EventTitle").Width = 7005
      .Columns("EventTitle").Width = .Width - 840
    End If
  
    .Columns("EventTime").Width = 600
  
  End With

End Sub


Private Sub PopulateByWeekView(dtStartDate As Date)

  Dim dtTempDate As Date
  Dim intDay As Integer
  Dim strOutput As String
  Dim intCount As Integer

  With rsTables

    If Not .EOF Then .MoveFirst

    For intDay = 0 To 6

      dtTempDate = Format(DateAdd("d", intDay, dtStartDate), DateFormat)

      With frmDiary.lblDayTitle(intDay)
        .Caption = Format(dtTempDate, mstrDATEDAY)
        .Tag = dtTempDate
        If Format(dtTempDate, mstrDATESQL) = Format(mdtSelectedDate, mstrDATESQL) Then
          .ForeColor = vbRed
        Else
          .ForeColor = vbBlack
        End If
      End With

      frmDiary.grdViewByWeek(intDay).Redraw = False
      frmDiary.grdViewByWeek(intDay).RemoveAll
      If Not .EOF Then
        Do While Format(.Fields("EventDate").Value, DateFormat) = dtTempDate

          strOutput = .Fields("DiaryEventsID").Value & vbTab & _
                      .Fields("Alarm").Value & vbTab & _
                      Format(.Fields("EventDate").Value, "hh:nn") & vbTab & _
                      .Fields("EventTitle").Value
          frmDiary.grdViewByWeek(intDay).AddItem strOutput
          
          If .Fields(0).Value = mlngSelectedID Then
            With frmDiary.grdViewByWeek(intDay)
              .Redraw = True
              .SelBookmarks.Add .GetBookmark(.Rows - 1)
              .Redraw = False
            End With
          End If

          .MoveNext
          If .EOF Then Exit Do
        Loop
      End If
      
      frmDiary.grdViewByWeek(intDay).Redraw = True
    
      With frmDiary.grdViewByWeek(intDay)
        If .Rows > 4 Then
          .ScrollBars = ssScrollBarsVertical
          .Columns("EventTitle").Width = 2625 '.Width - 1070    '2275
        Else
          .ScrollBars = ssScrollBarsNone
          .Columns("EventTitle").Width = 2850 '.Width - 845 '2500
        End If
        
        .Columns("EventTime").Width = 600
        
      End With
      
    Next

  End With

End Sub


Private Sub PopulateByMonthView(dtStartDate As Date, dtEndDate As Date)
On Error GoTo DateError
  Dim dtTempDate As Date
  Dim intDay As Integer
  Dim blnFound As Boolean
  Dim strDateFormat As String

  strDateFormat = DateFormat

  With frmDiary.mvwViewbyMonth
    
    'DoEvents
    'frmDiary.fraViewType(mintVIEWBYMONTH).ClipControls = False
    
    '.Visible = False
    '.Value = dtStartDate
    '.Value = dtEndDate
    '.Value = Me.DateSelected

    With rsTables

      If Not .EOF Then
        .MoveFirst
      End If

      For intDay = 0 To DateDiff("d", dtStartDate, dtEndDate)
        dtTempDate = DateAdd("d", intDay, dtStartDate)

        If Not .EOF Then
          blnFound = Format(.Fields("EventDate").Value, strDateFormat) = _
                     Format(dtTempDate, strDateFormat)
          If blnFound Then .MoveNext
        Else
          blnFound = False
        End If

        frmDiary.mvwViewbyMonth.DayBold(dtTempDate) = blnFound
      Next

    End With

    'frmDiary.fraViewType(mintVIEWBYMONTH).ClipControls = True
    'frmDiary.fraViewType(mintVIEWBYMONTH).Visible = False
  End With

DateError:
  'NHRD22072003 Fault 6294
  'If Err.Number = 35773 Then
  '  COAMsgBox "Date ranges entered are not valid." & vbCrLf & "The dates must be within the months shown.", vbOKOnly + vbExclamation
  'End If
End Sub


Private Sub PopulateByListView()

  Dim blnExactMatch As Boolean

  With frmDiary.grdViewByList
    
    .Visible = False
    .Rebind
    .Rows = mlngRecordCount

    If rsTables.BOF And rsTables.EOF Then
      .Visible = True
      .ScrollBars = ssScrollBarsNone
      Exit Sub
    End If

    'Check if we can find the exact record
    blnExactMatch = False
    If mlngSelectedID > 0 Then
      rsTables.MoveFirst
      rsTables.Find "DiaryEventsID = " & CStr(mlngSelectedID)
      blnExactMatch = Not (rsTables.EOF)
    End If

    'Check if we can find the exact date
    If Not blnExactMatch Then
      rsTables.MoveFirst
      rsTables.Find "FindDate >= " & Format(mdtSelectedDate, "yyyymmdd")
      If Not rsTables.EOF Then
        blnExactMatch = (Format(rsTables.Fields("EventDate").Value, "yyyymmdd") = Format(mdtSelectedDate, "yyyymmdd"))
      End If
    End If

'MH20060914 Fault 11416
'    If blnExactMatch Then
'      DiaryEventID = rsTables.Fields("DiaryEventsID").Value
'      .SelBookmarks.Add rsTables.Bookmark
'    Else
'      rsTables.MovePrevious
'      If rsTables.BOF Then
'        rsTables.MoveFirst
'      End If
'    End If
    If Not blnExactMatch Then
      rsTables.MovePrevious
      If rsTables.BOF Then
        rsTables.MoveFirst
      End If
    End If

    '.Visible = True

    If mlngRecordCount <= 19 Then
      'Less than 20 records so show all
      .MoveFirst
      .ScrollBars = ssScrollBarsNone
      '.Columns("Title").Width = 6000
      .Columns("Title").Width = .Width - 1845

    Else

      If rsTables.EOF Or (rsTables.Bookmark >= mlngRecordCount - 18) Then
        .FirstRow = mlngRecordCount - 17
      'MH20001201 This is well dodgey... the first twenty rows
      'in the grid don't need an offset but the others do !
      '(I hate this bloody grid !)
      ElseIf rsTables.Bookmark <= 20 Then
        .FirstRow = rsTables.Bookmark
      Else
        .FirstRow = rsTables.Bookmark + 1
      End If

      .ScrollBars = ssScrollBarsVertical
      '.Columns("Title").Width = 5765
      .Columns("Title").Width = .Width - 2090

    End If

    DiaryEventID = .Columns("DiaryEventID").Value
    .SelBookmarks.Add .Bookmark

    .Visible = True

  End With

End Sub


Public Function GetCurrentRecord() As ADODB.Recordset
  Screen.MousePointer = vbHourglass
  'Set GetCurrentRecord = datGeneral.GetRecords( _
    "SELECT ASRSysDiaryEvents.*, " & _
    "CONVERT(integer,ASRSysDiaryEvents.TimeStamp) AS intTimeStamp " & _
    "FROM ASRSysDiaryEvents " & SQLWhereCurrentEvent)
  Set GetCurrentRecord = datGeneral.GetReadOnlyRecords( _
    "SELECT ASRSysDiaryEvents.*, " & _
    "CONVERT(integer,ASRSysDiaryEvents.TimeStamp) AS intTimeStamp " & _
    "FROM ASRSysDiaryEvents " & SQLWhereCurrentEvent)
  Screen.MousePointer = vbDefault
End Function


Public Sub PutRecord(strTitle As String, dtDate As Date, _
                     strTime As String, strNotes As String, _
                     intAlarm As Integer, ByVal strAccess As String, _
                     lngCopiedFrom As Long)

  Dim strSQL As String
  Dim strSQLTitle As String
  Dim strSQLNotes As String
  Dim strSQLDate As String
  
  'This sub will check to see if you are currently updating a
  'record by checking the DiaryEventId.  If this variable is
  'blank then a new record will be created.

  strSQLTitle = Replace(Trim$(strTitle), "'", "''")
  strSQLNotes = Replace(strNotes, "'", "''")
  strSQLDate = Replace(Format(dtDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " " & Replace(Left(strTime, 2) & ":" & Right(strTime, 2), "_", "0")

  If mlngSelectedID > 0 Then
    strSQL = "UPDATE ASRSysDiaryEvents SET " & _
             "EventTitle = '" & strSQLTitle & "', " & _
             "EventDate = '" & strSQLDate & "', " & _
             "EventNotes = '" & strSQLNotes & "', " & _
             "Alarm = '" & CStr(intAlarm) & "', " & _
             "Access = '" & strAccess & "' " & _
             SQLWhereCurrentEvent

  Else
    strSQL = "INSERT ASRSysDiaryEvents (" & _
                "TableID, LinkID, ColumnID, RowID, " & _
                "EventTitle, " & _
                "EventDate, " & _
                "EventNotes, " & _
                "UserName, " & _
                "Alarm, " & _
                "Access, " & _
                "CopiedFromID) "
    strSQL = strSQL & "VALUES( " & _
                "0, 0, 0, 0, " & _
                "'" & strSQLTitle & "', " & _
                "'" & strSQLDate & "', " & _
                "'" & strSQLNotes & "', " & _
                "'" & datGeneral.UserNameForSQL & "', " & _
                CStr(intAlarm) & ", " & _
                "'" & strAccess & "', " & _
                CStr(lngCopiedFrom) & ")"
  End If
  
  Call ExecuteSql(strSQL)

End Sub

Public Property Get FilterEventType() As Integer
  FilterEventType = mintFilterEventType
End Property

Public Property Let FilterEventType(ByVal intNewValue As Integer)
  '0 - Both System and Manual Events
  '1 - System Events
  '2 - Manual Events
  '3 - Manual Events where current user is owner
  If intNewValue < 2 And Not AllowSystemEvents Then
    intNewValue = 2
  End If
  mintFilterEventType = intNewValue
End Property

Public Property Get FilterAlarmStatus() As Integer
  FilterAlarmStatus = mintFilterAlarmStatus
End Property
Public Property Let FilterAlarmStatus(ByVal intNewValue As Integer)
  '0 - Both Alarmed and Non-Alarmed
  '1 - Alarmed Only
  '2 - Non-Alarmed Only
  mintFilterAlarmStatus = intNewValue
End Property

Public Property Get FilterPastPresent() As Integer
  FilterPastPresent = mintFilterPastPresent
End Property
Public Property Let FilterPastPresent(ByVal intNewValue As Integer)
  '0 - All
  '1 - Only Past Events
  '2 - Only Future Events
  mintFilterPastPresent = intNewValue
End Property

Public Property Get FilterOnlyMine() As Boolean
  FilterOnlyMine = mblnFilterOnlyMine
End Property

Public Property Let FilterOnlyMine(ByVal blnNewValue As Boolean)
  mblnFilterOnlyMine = blnNewValue
End Property

Public Function FilterText() As String

  FilterText = _
      Choose(mintFilterPastPresent + 1, "", "Past ", "Current and Future ") _
      & _
      Choose(mintFilterAlarmStatus + 1, "", "Alarmed ", "Non-Alarmed ") _


  Select Case FilterEventType
  Case 0
    FilterText = FilterText & _
        "System Events and Manual Events" & _
        IIf(mblnFilterOnlyMine, " where owner is '" & gsUserName & "'", "")
  Case 1
    FilterText = FilterText & _
        "System Diary Events"
  Case 2
    FilterText = FilterText & _
        "Manual Diary Events" & _
        IIf(mblnFilterOnlyMine, " where owner is '" & gsUserName & "'", "")
  End Select


  'Check if ClearFilter button should be enabled
  'frmDiary.SSActiveToolBars1.Tools(19).Enabled = _
      (mintFilterPastPresent > 0 Or _
       mintFilterAlarmStatus > 0 Or _
       mintFilterEventType > 0)
  EnableTools

End Function


Private Sub ExecuteSql(strSQL As String)
    
  Dim intCurrentMousePointer As Integer
  
  intCurrentMousePointer = Screen.MousePointer
  Screen.MousePointer = vbHourglass
  
  DoEvents
  gADOCon.Execute strSQL, , adCmdText
  
  Screen.MousePointer = intCurrentMousePointer

End Sub


Public Function GetAlarmCount(intAlarmSetting As Integer) As Long

  Dim rsAlarmCount As ADODB.Recordset
  Dim strSQL As String
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  Dim strEventDateMatch As String
  
  Screen.MousePointer = vbHourglass

  If mintCurrentView <> mintVIEWBYLIST Then
    Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
    strEventDateMatch = _
           "(EventDate BETWEEN " & _
           "'" & Replace(Format(dtStartDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 00:00' AND " & _
           "'" & Replace(Format(dtEndDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 23:59')" & " AND "
  End If

  strSQL = "SELECT COUNT(DiaryEventsID) AS recCount " & _
      "FROM ASRSysDiaryEvents WHERE " & _
      strEventDateMatch & _
      "Alarm = '" & CStr(intAlarmSetting) & "' AND " & _
      SQLFilter
      
  'Set rsAlarmCount = datGeneral.GetRecords(strSQL)
  Set rsAlarmCount = datGeneral.GetReadOnlyRecords(strSQL)

  GetAlarmCount = rsAlarmCount!recCount
  
  rsAlarmCount.Close
  Set rsAlarmCount = Nothing

  Screen.MousePointer = vbDefault

End Function


'Public Sub GetNextAlarmTime(lngNextAlarmTime As Long)
'
'  Dim strSQL As String
'  Dim strSQLEventDate As String
'  Dim rsTemp As ADODB.Recordset
'
'  Me.FilterEventType = 2      'All Manual
'  Me.FilterAlarmStatus = 1    'Alarmed
'  Me.FilterPastPresent = 2    'Current and Future
'
'  strSQL = "SELECT TOP 1 * " & _
'           "FROM AsrSysDiaryEvents " & _
'           "WHERE " & SQLFilter & mstrSQLOrderByClause
'
'  'Set rsTemp = datGeneral.GetRecords(strSQL)
'  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
'
'  lngNextAlarmTime = -1
'  If Not rsTemp.EOF Then
'
'    If Format(rsTemp.Fields("EventDate").Value, mstrDATESQL) = Format(Now, mstrDATESQL) Then
'      'This will get the number of seconds past midnight that
'      'the alarm should occur (for comparison to TIMER)
'      lngNextAlarmTime = _
'        ConvertTimeToMins(Format(rsTemp.Fields("EventDate").Value, "hh:nn")) * 60
'    End If
'
'  End If
'
'  rsTemp.Close
'  Set rsTemp = Nothing
'
'End Sub


Public Function ShowAlarmedEvents(intFilterPastPresent As Integer, intViewMode As Integer) As Boolean

  Dim dtStartDate As Date
  Dim dtEndDate As Date

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  Dim lngErrorNumber As Long

  On Local Error GoTo LocalErr

  lngErrorNumber = 0
  Me.ViewingAlarms = True   'This need to be done here in case
                                  'you don't view alarms (filter will be saved!)

  'Set filter to all alarmed events in the past
  Me.FilterEventType = 0      'All
  Me.FilterAlarmStatus = 1    'Alarmed
  Me.FilterPastPresent = intFilterPastPresent
  'Me.DateSelected = Now
  mintCurrentView = intViewMode
  
  
  If GetRecordCount > 0 Then
    'NHRD07102004 Fault 4436 Adds a Bell resource icon to the System Tray

    RemoveSysTrayIcon
    AddSysTrayIcon "HR Pro Alarmed Diary Event"

    'strMBText = "There are alarmed events " & _
                IIf(intFilterPastPresent = 1, "prior to", "for the") & _
                " current date and time" & vbCr & _
                "Would you like to view these now?"
    'intMBButtons = vbYesNo + vbQuestion
    'strMBTitle = "Alarmed Diary Events"
    'intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    intMBResponse = frmDiaryAlert.ShowAlert(intFilterPastPresent <> 1)
    lngErrorNumber = frmDiaryAlert.ErrorNumber
    Set frmDiaryAlert = Nothing

    'Removes the Bell resource icon from the System Tray
    If lngErrorNumber = 0 Then
      RemoveSysTrayIcon
    End If


    If intMBResponse = vbYes Then

      'Set filter to all alarmed events in the past
      '(These seem to be getting reset in the exe so ensure that they are set)
      Me.FilterEventType = 0      'All
      Me.FilterPastPresent = intFilterPastPresent
      Me.FilterAlarmStatus = 1    'Alarmed
      Me.DateSelected = Now
      mintCurrentView = intViewMode

      'This is only if you a diary event pops up during the day
      If mintCurrentView = mintVIEWBYDAY Then
        Me.FilterPastPresent = 0    'All (incase the minute has ticked over).
      End If

      frmDiary.Initialise
      frmDiary.Show vbModal
      Unload frmDiary
      Set frmDiary = Nothing

    End If
  End If

  ShowAlarmedEvents = (lngErrorNumber = 0)

Exit Function

LocalErr:
  ShowAlarmedEvents = False

End Function


Public Sub SwitchAlarms(strSetAlarm As String)

  'Update alarm flag where:
  'The old alarm flag doesn't not match the new alarm flag
  'The event date is the same as the current date in the diary
  'Its a system event in the past OR
  '    it a manual event and the user has write access to this event

  Dim strSQL As String
  Dim strEventDateMatch As String
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  Dim lngRecs As Long

  Screen.MousePointer = vbHourglass
  
  If mintCurrentView <> mintVIEWBYLIST Then
    Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
    strEventDateMatch = _
           "(EventDate BETWEEN " & _
           "'" & Replace(Format(dtStartDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 00:00' AND " & _
           "'" & Replace(Format(dtEndDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 23:59')" & " AND "
  End If
  
  strSQL = "UPDATE ASRSysDiaryEvents " & _
           "SET Alarm = " & strSetAlarm & _
           " WHERE Alarm <> " & strSetAlarm & " AND " & _
           strEventDateMatch & _
           "((ColumnID > 0 AND (EventDate < '" & Replace(DiaryFormat(Now, mstrDATESQL & " hh:nn"), UI.GetSystemDateSeparator, "/") & "')) OR" & _
           " (ColumnID = 0 AND (UserName = '" & datGeneral.UserNameForSQL & "' OR Access = 'RW'))) AND " & _
           SQLFilter
  Screen.MousePointer = vbHourglass
  lngRecs = datData.ExecuteSqlReturnAffected(strSQL)
  Screen.MousePointer = vbDefault

  'MH20001019 Fault 1068
  'Warning message if you have been unable to change any of the alarms !
  If lngRecs = 0 Then
    COAMsgBox "You do not have access to the alarms on any of these events", vbInformation, "Diary Alarms"
  End If

End Sub


Private Sub Class_Terminate()
  
  If rsTables.State = adStateOpen Then
    rsTables.Close
  End If
  Set rsTables = Nothing
  Set frmDiary = Nothing
  Set mobjClipBoard = Nothing

End Sub

Public Function GetPrintData(piDateRange As HRProDataMgr.ReturnPrintDateType, pdStartManualDate As Date _
    , pdEndManualDate As Date, plngStartExpressionID As Long, plngEndExpressionID As Long) As String


  Dim dtStartDate As Date
  Dim dtEndDate As Date
  Dim vStartDate As Variant
  Dim vEndDate As Variant
  Dim bShowWeekNumber As Boolean
  Dim strDateRange As String
  
  Screen.MousePointer = vbHourglass
  frmDiary.AutoRedraw = False
  DoEvents

  Select Case piDateRange
    Case RETURN_DEFAULT
      Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
      bShowWeekNumber = True
    
      strDateRange = vbNullString

    Case RETURN_MANUAL
      dtStartDate = pdStartManualDate
      dtEndDate = pdEndManualDate
      bShowWeekNumber = False
    
      strDateRange = " (" & Format(dtStartDate, DateFormat) & " - " & _
                            Format(dtEndDate, DateFormat) & ")"
    Case RETURN_CALCULATION
      dtStartDate = datGeneral.GetValueForRecordIndependantCalc(plngStartExpressionID)
      dtEndDate = datGeneral.GetValueForRecordIndependantCalc(plngEndExpressionID)
      bShowWeekNumber = False
  
      strDateRange = " (" & Format(dtStartDate, DateFormat) & " - " & _
                            Format(dtEndDate, DateFormat) & ")"
  End Select

  Call GetDiaryData(False, dtStartDate, dtEndDate)
  'Call PopulateDiaryData(dtStartDate, dtEndDate)

  ' Set the date range (so the heading text works)
  'mstrDateRangeDesc = IIf(bShowWeekNumber, frmDiary.lblWeekNo, "") & " (" & _
                      Format(dtStartDate, DateFormat) & " - " & _
                      Format(dtEndDate, DateFormat) & ")"
  
  
  frmDiary.AutoRedraw = True
  With frmDiary.grdPrint
    .Rebind
    .SelBookmarks.Add .GetBookmark(0)
  End With
  Screen.MousePointer = vbDefault

  GetPrintData = strDateRange

End Function


Public Sub ListUnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String

  sDateFormat = DateFormat

  ' This is required As ADODB.Recordset not set when this sub is first run
  If rsTables Is Nothing Then Exit Sub
  If rsTables.State = adStateClosed Then Exit Sub

  If (rsTables.BOF And rsTables.EOF) Then
    Exit Sub
  End If

  If (mintCurrentView = mintVIEWBYMONTH And Not mblnPrinting) _
      Or mblnFormattingGrid Then
    Exit Sub
  End If
  
  'If StartLocation < 0 Then 'Or StartLocation >= frmDiary.grdViewByList.Rows Then
  If StartLocation < 0 Or StartLocation > rsTables.RecordCount Then
    Exit Sub
  End If


'  'MH20010308 Fault 1423
'  'The fix below caused another error (i.e. fixing 780 caused 1423)
'  'As fault 780 really related to the find window and was not recreated
'  'in the diary I have removed the fix for 780 from here !
'
'''  'MH20001026 Fault 780 (Was on find window but put fix in here too)
'''  'Setting focus to the grid stops the order combo box changing
'''  'when using the fancy (tarty) mouse wheel.  This stops the
'''  'invalid bookmark error
'''  With frmDiary.grdViewByList
'''    If .Visible And .Enabled Then
'''      .SetFocus
'''    End If
'''  End With
  
  
  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not rsTables.EOF Then
        rsTables.MoveLast
      End If
    Else
      If Not rsTables.BOF Then
        rsTables.MoveFirst
      End If
    End If
  Else
    rsTables.Bookmark = StartLocation
    If ReadPriorRows Then
      rsTables.MovePrevious
    Else
      rsTables.MoveNext
    End If
  End If

  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If rsTables.BOF Or rsTables.EOF Then
      'With frmDiary.grdViewByList
      '  .Rows = .FirstRow + .Row
      'End With
      Exit For
    End If

    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        RowBuf.Value(iRowIndex, 0) = rsTables.Fields("DiaryEventsID").Value
        RowBuf.Value(iRowIndex, 1) = rsTables.Fields("Alarm").Value
        RowBuf.Value(iRowIndex, 2) = Format(rsTables.Fields("EventDate").Value, sDateFormat & "  hh:nn")
        RowBuf.Value(iRowIndex, 3) = rsTables.Fields("EventTitle").Value
        
        RowBuf.Bookmark(iRowIndex) = rsTables.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = rsTables.Bookmark
      Case 2
        RowBuf.Value(iRowIndex, 0) = rsTables(0)
        RowBuf.Bookmark(iRowIndex) = rsTables.Bookmark
      Case 3
    End Select

    If ReadPriorRows Then
      rsTables.MovePrevious
    Else
      rsTables.MoveNext
    End If

    iRowsRead = iRowsRead + 1
  Next iRowIndex

  RowBuf.RowCount = iRowsRead
  
End Sub


'Private Sub CheckAccessToSystemEvents()
'
'  Dim rsTemp As ADODB.Recordset
'  Dim strSQL As String
'
'  Dim objTableView As CTablePrivilege
'  Dim objColumnPrivileges As CColumnPrivileges
'  Dim strWhere As String
'
'  strSQL = "SELECT DISTINCT ASRSysTables.TableID, ASRSysTables.TableName, " & _
'           "                ASRSysColumns.ColumnID, ASRSysColumns.ColumnName " & vbCrLf & _
'           "FROM ASRSysDiaryLinks " & vbCrLf & _
'           "JOIN ASRSysColumns ON ASRSysDiaryLinks.ColumnID = ASRSysColumns.ColumnID " & vbCrLf & _
'           "JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID"
'  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
'
'  If rsTemp.BOF And rsTemp.EOF Then
'    'Can't find any links !!!
'    'mstrSystemEvents = ""
'    mstrSystemEvents = "ColumnID = -1"    'MH20061010 Fault 11566
'    Exit Sub
'  End If
'
'  rsTemp.MoveFirst
'
'  Do While Not rsTemp.EOF
'
'    Set objColumnPrivileges = gcolColumnPrivilegesCollection.Item(rsTemp!TableName)
'
'    If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
'      mstrSystemEvents = mstrSystemEvents & _
'        IIf(mstrSystemEvents <> vbNullString, " OR ", "") & _
'        "(ColumnID = " & CStr(rsTemp!ColumnID) & ")" & vbCrLf
'
'    Else
'
'      strWhere = vbNullString
'
'      For Each objTableView In gcoTablePrivileges.Collection
'
'        'Loop thru all of the views for this table where the user has select access
'        If (objTableView.IsTable = False) And _
'           (objTableView.TableID = rsTemp!TableID) And _
'           (objTableView.AllowSelect) Then
'
'          Set objColumnPrivileges = gcolColumnPrivilegesCollection.Item(objTableView.ViewName)
'
'          If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
'            If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
'
'              strWhere = strWhere & _
'                IIf(strWhere <> vbNullString, " OR ", "") & _
'                "RowID IN (SELECT ID FROM " & objTableView.ViewName & ")"
'
'            End If
'          End If
'
'        End If
'      Next
'
'      If strWhere <> vbNullString Then
'        mstrSystemEvents = mstrSystemEvents & _
'          IIf(mstrSystemEvents <> vbNullString, " OR ", "") & _
'          "(ColumnID = " & CStr(rsTemp!ColumnID) & " AND (" & strWhere & "))" & vbCrLf
'      End If
'
'    End If
'
'    rsTemp.MoveNext
'  Loop
'
'  If mstrSystemEvents <> vbNullString Then
'    mstrSystemEvents = "(" & mstrSystemEvents & ")"
'  End If
'
'End Sub


Private Sub CheckAccessToSystemEvents()

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim strColList As String

  strSQL = "SELECT DISTINCT ASRSysTables.TableID, ASRSysTables.TableName, " & _
           "                ASRSysColumns.ColumnID, ASRSysColumns.ColumnName " & vbCrLf & _
           "FROM ASRSysDiaryLinks " & vbCrLf & _
           "JOIN ASRSysColumns ON ASRSysDiaryLinks.ColumnID = ASRSysColumns.ColumnID " & vbCrLf & _
           "JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID"
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    'Can't find any links !!!
    'mstrSystemEvents = ""
    mstrSystemEvents = "ColumnID = -1"    'MH20061010 Fault 11566
    Exit Sub
  End If
  
  
  For Each objTableView In gcoTablePrivileges.Collection

    strColList = vbNullString
    If (objTableView.AllowSelect) Then
    
      rsTemp.MoveFirst
      Do While Not rsTemp.EOF

        'Loop thru all of the views for this table where the user has select access
        If (objTableView.TableID = rsTemp!TableID) Then
          Set objColumnPrivileges = gcolColumnPrivilegesCollection.Item(IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName))
  
          If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
            If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
              strColList = strColList & IIf(strColList <> vbNullString, ", ", "") & CStr(rsTemp!ColumnID)
            End If
          End If
  
        End If
    
        rsTemp.MoveNext
      Loop
    
      If strColList <> vbNullString Then
        mstrSystemEvents = mstrSystemEvents & _
          IIf(mstrSystemEvents <> vbNullString, " OR " & vbCrLf, "") & _
          "(RowID IN (SELECT ID FROM " & objTableView.RealSource & ") AND (ColumnID IN (" & strColList & ")))"
      End If
    
    End If
  Next


  If mstrSystemEvents <> vbNullString Then
    mstrSystemEvents = "(" & mstrSystemEvents & ")"
  End If

End Sub


Private Sub EnableTools()

  Dim dtStartDate As Date
  Dim dtEndDate As Date
  Dim blnEnabled As Boolean

  With frmDiary.DiaryToolBar

    If mintCurrentView <> mintVIEWBYLIST Then
      
      If mintCurrentView = mintVIEWBYMONTH Then
        'Just get start and end of month, if in month view
        dtStartDate = DateAdd("d", (Day(mdtSelectedDate) * -1) + 1, mdtSelectedDate)
        dtEndDate = DateAdd("d", 1, DateAdd("m", 1, dtStartDate))
      Else
        Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
      End If

      .Item("Previous").Enabled = (DateDiff("d", dtStartDate, frmDiary.mvwViewbyMonth.MinDate) < 0)
      .Item("Next").Enabled = (DateDiff("d", dtEndDate, frmDiary.mvwViewbyMonth.MaxDate) > 0)

    Else
      .Item("Previous").Enabled = False
      .Item("Next").Enabled = False
    
    End If

    blnEnabled = (mlngSelectedID > 0 And mlngRecordCount > 0 And mintCurrentView <> mintVIEWBYMONTH)
    
    .Item("Repeat").Enabled = (blnEnabled And mlngColumnID = 0)
    .Item("Edit").Enabled = blnEnabled
    .Item("Delete").Enabled = (mlngRecordCount > 0)

    .Item("Filter").Enabled = (gobjDiary.ViewingAlarms = False)
    
    'MH20020701 Fault 4088
    '.Item("ClearFilter").Enabled = (frmDiary.ViewingAlarms = False And _
          (mintFilterPastPresent > 0 Or mintFilterAlarmStatus > 0 Or mintFilterEventType > 0))
    .Item("ClearFilter").Enabled = (gobjDiary.ViewingAlarms = False And _
          (mintFilterPastPresent > 0 Or mintFilterAlarmStatus > 0 Or _
          (FilterEventType > 0 And AllowSystemEvents)))

    .Item("Cut").Enabled = (blnEnabled And mlngColumnID = 0)
    .Item("Copy").Enabled = (blnEnabled And mlngColumnID = 0)
    .Item("Paste").Enabled = (mintCurrentView <> mintVIEWBYMONTH And Not (mobjClipBoard Is Nothing))
    
'    .Item("Print").Enabled = (mintCurrentView <> mintVIEWBYMONTH)

    If rsTables.State <> adStateClosed Then
      .Item("Alarm").Enabled = Not (rsTables.BOF And rsTables.EOF)
    End If

  End With

End Sub

Public Sub DeleteCurrentEntry()

  Dim strSQL As String
  Dim strEventTitle As String
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  Dim strSelectedIDs As String
  Dim lngRecs As Long             '23/07/2001 MH Fault 2131

  strSelectedIDs = GetSelectedIDs

  With frmSelection
    .Icon = frmDiary.Icon
    .HelpContextID = frmDiary.HelpContextID
    .OptionCount = 4
    .Source = "Diary"
    
    If Trim(strSelectedIDs) = vbNullString Then
      'No items are selected so disable these options
      .optSelection(0).Enabled = False
      .optSelection(3).Enabled = False
    
      
      'ND26022002 Fault 3397
      '.optSelection(1) will need to be enabled when there are entries in the current view
      'but non are selected i.e. Trim(strSelectedIDs) = vbNullString
      .optSelection(1).Enabled = rsTables.RecordCount > 0
      'old line - .optSelection(1).Enabled = (mlngRecordCount > 0)
    End If

    .Show vbModal
  
    'Answer
    '0 = Highlighted only
    '1 = All in current view (and filter!)
    '2 = All that we have access to
    '3 = Highlighted and copies
    
    If .Answer <> -1 Then

      Screen.MousePointer = vbHourglass

      'This is all the events that the current user has access to
      '(So when answer = 2 then don't add to this where clause!)
      strSQL = "DELETE FROM ASRSysDiaryEvents WHERE ColumnID = 0 " & _
               "AND (UserName = '" & datGeneral.UserNameForSQL & "' OR Access = 'RW') "

      If .Answer = 1 Then       'All in current view (and filter!)
        Call GetDateRange(dtStartDate, dtEndDate)     'This sets up these two variables
        strSQL = strSQL & _
          "AND EventDate BETWEEN " & _
          "'" & Replace(Format(dtStartDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 00:00' AND " & _
          "'" & Replace(Format(dtEndDate, mstrDATESQL), UI.GetSystemDateSeparator, "/") & " 23:59' AND " & _
          SQLFilter

      Else
      
        Select Case .Answer
        Case 0    'Highlighted
          strSQL = strSQL & _
            "AND DiaryEventsID IN (" & strSelectedIDs & ")"

        Case 3    'Highlighted and copies
          'MH20030326 Fault 5208
          'strSQL = strSQL & _
             "AND (DiaryEventsID IN (" & strSelectedIDs & ")" & _
             " OR DiaryEventsID IN (SELECT CopiedFromID FROM ASRSysDiaryEvents WHERE DiaryEventsID IN (" & strSelectedIDs & "))" & _
             " OR CopiedFromID IN (" & strSelectedIDs & ")" & _
             " OR CopiedFromID IN (SELECT CopiedFromID FROM ASRSysDiaryEvents WHERE DiaryEventsID IN (" & strSelectedIDs & ")))"
          strSQL = strSQL & _
             "AND (DiaryEventsID IN (" & strSelectedIDs & ")" & _
             " OR DiaryEventsID IN (SELECT CopiedFromID FROM ASRSysDiaryEvents WHERE DiaryEventsID IN (" & strSelectedIDs & "))" & _
             " OR CopiedFromID IN (" & strSelectedIDs & ")" & _
             " OR (CopiedFromID IN (SELECT CopiedFromID FROM ASRSysDiaryEvents WHERE DiaryEventsID IN (" & strSelectedIDs & "))) AND CopiedFromID > 0)"

        End Select

      End If
      
      
      '23/07/2001 MH Fault 2131
      'Call ExecuteSql(strSQL)
      'Warning message if you have been unable to delete any of the events
      lngRecs = datData.ExecuteSqlReturnAffected(strSQL)
      If lngRecs = 0 Then
        COAMsgBox "You do not have permission to delete " & _
               IIf(InStr(strSelectedIDs, ",") > 1, "all of these events.", "this event."), _
               vbExclamation, "Diary Delete"
      End If

      Let DiaryEventID = 0

    End If

    Unload frmSelection
    Set frmSelection = Nothing

  End With

  Screen.MousePointer = vbDefault

End Sub


Public Sub CutEntries(blnDelete As Boolean)

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim blnReadOnlyEvents As Boolean
  Dim objTempDiaryEvent As clsDiaryEvent
  Dim strSelectedIDs As String

  strSelectedIDs = GetSelectedIDs
  If strSelectedIDs = vbNullString Then
    Exit Sub
  End If

  Set mobjClipBoard = Nothing 'Not sure if required??
  Set mobjClipBoard = New Collection


  strSQL = "SELECT * FROM ASRSysDiaryEvents " & _
           "WHERE DiaryEventsID IN (" & strSelectedIDs & ")"
  'Set rsTemp = datGeneral.GetRecords(strSQL)
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  blnReadOnlyEvents = False
  Do While Not rsTemp.EOF

    If rsTemp!ColumnID > 0 Then
      blnReadOnlyEvents = True
    Else

      Set objTempDiaryEvent = New clsDiaryEvent

      With objTempDiaryEvent
        .EventTitle = rsTemp!EventTitle
        .EventDate = rsTemp!EventDate
        .EventTime = DiaryFormat(rsTemp!EventDate, "hh:nn")
        .EventNotes = rsTemp!EventNotes
        .UserName = rsTemp!UserName
        .Alarm = IIf(rsTemp!Alarm, True, False)
        .Access = rsTemp!Access
        .CopiedFromID = rsTemp!diaryeventsid
      End With

      If rsTemp!Access <> "RW" And Trim(UCase(rsTemp!UserName)) <> Trim(UCase(gsUserName)) Then
        blnReadOnlyEvents = True
      End If
      
      mobjClipBoard.Add objTempDiaryEvent

      Set objTempDiaryEvent = Nothing

    End If

    rsTemp.MoveNext
  Loop

  rsTemp.Close
  Set rsTemp = Nothing


  If blnDelete Then
    
    If blnReadOnlyEvents Then
        COAMsgBox "You do not have permission to delete " & _
               IIf(InStr(strSelectedIDs, ",") > 1, "all of these events.", "this event."), _
               vbExclamation, "Diary Delete"
    End If

    strSQL = "DELETE FROM ASRSysDiaryEvents " & _
             "WHERE ColumnID = 0 AND DiaryEventsID IN (" & strSelectedIDs & ") AND " & _
             "(Lower(UserName) = '" & LCase(datGeneral.UserNameForSQL) & "' OR Access = 'RW') AND " & _
             SQLFilter
    ExecuteSql strSQL

    DiaryEventID = 0

  End If

End Sub


Public Sub PasteEntries()

  Dim objTempDiaryEvent As clsDiaryEvent

  For Each objTempDiaryEvent In mobjClipBoard

    With objTempDiaryEvent
      Select Case mintCurrentView
      Case mintVIEWBYLIST
        DiaryEventID = 0
        PutRecord .EventTitle, .EventDate, .EventTime, .EventNotes, .Alarm, .Access, .CopiedFromID
      
      Case mintVIEWBYDAY, mintVIEWBYWEEK
        DiaryEventID = 0
        PutRecord .EventTitle, mdtSelectedDate, .EventTime, .EventNotes, .Alarm, .Access, .CopiedFromID

      Case mintVIEWBYMONTH
        COAMsgBox "Unable to paste into the month view", vbExclamation, "Diary"

      End Select
    End With

  Next

End Sub


Private Function GetSelectedIDs() As String

  Dim intCount As Integer

  Select Case mintCurrentView
  Case mintVIEWBYDAY
    GetSelectedIDs = GetSelectedIDsFromGrid(frmDiary.grdViewByDay)
  
  Case mintVIEWBYWEEK
    GetSelectedIDs = vbNullString
    For intCount = frmDiary.grdViewByWeek.LBound To frmDiary.grdViewByWeek.UBound
      GetSelectedIDs = GetSelectedIDs & _
          GetSelectedIDsFromGrid(frmDiary.grdViewByWeek(intCount))
    Next
  
  Case mintVIEWBYLIST
    GetSelectedIDs = GetSelectedIDsFromGrid(frmDiary.grdViewByList)
  
  Case mintVIEWBYMONTH
    GetSelectedIDs = vbNullString
  
  End Select
  
End Function

Private Function GetSelectedIDsFromGrid(ByRef grdTemp As SSDataWidgets_B.SSDBGrid) As String

  Dim strResult As String
  Dim lngCount As Long

  strResult = vbNullString
  
  With grdTemp
    For lngCount = 0 To .SelBookmarks.Count - 1
      strResult = strResult & _
        IIf(strResult <> vbNullString, ", ", vbNullString) & _
        CStr(Val(.Columns(0).CellValue(.SelBookmarks.Item(lngCount))))
    Next
  End With

  GetSelectedIDsFromGrid = strResult

End Function


Public Function CheckAlarmedEvents(tmrDiary As Timer, strLastAlarmCheck As String) As String

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String
  Dim strCurrentDate As String
  Dim blnAlarmedEvents As Boolean

  On Local Error GoTo LocalErr

  gobjErrorStack.Disable


  'Ignore the first time as user would be prompted at logon
  strCurrentDate = SQLCurrentDateTime
  If strLastAlarmCheck = vbNullString Then
    strLastAlarmCheck = strCurrentDate
    Exit Function
  End If


  'Check if there are any alarmed events
  strSQL = "SELECT TOP 1 * FROM AsrSysDiaryEvents " & _
           "WHERE Alarm = 1 " & _
           "AND DateDiff(n, '" & strCurrentDate & "', EventDate) <= 0 " & _
           "AND DateDiff(n, '" & strLastAlarmCheck & "', EventDate) > 0" & _
           mstrSQLOrderByClause
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  blnAlarmedEvents = Not (rsTemp.BOF And rsTemp.EOF)
  rsTemp.Close
  Set rsTemp = Nothing


  'Show events if required
  If blnAlarmedEvents Then
    If gobjDiary.ShowAlarmedEvents(2, 0) Then
      'Make sure that the timer fires as the minute ticks over
      tmrDiary.Interval = (60 - (Int(Timer) Mod 60)) * 1000
      strLastAlarmCheck = strCurrentDate
    End If
  Else
    'Make sure that the timer fires as the minute ticks over
    tmrDiary.Interval = (60 - (Int(Timer) Mod 60)) * 1000
    strLastAlarmCheck = strCurrentDate
  End If

  gobjErrorStack.Enable

  CheckAlarmedEvents = strLastAlarmCheck

Exit Function

LocalErr:
  tmrDiary.Enabled = False
  gobjErrorStack.Enable
  COAMsgBox "An error occurred checking for alarmed diary events.  No further checks will be made for alarmed events until HR Pro is restarted." & _
        IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", ""), vbExclamation, "Alarmed Diary Events"

End Function
