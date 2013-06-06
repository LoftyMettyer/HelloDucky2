Attribute VB_Name = "modAbsenceCalendar"
Option Explicit

'AE20071108 Fault #12547
Const TWIPS = 10

' Box (Label) Constants
Const MONTH_BOXWIDTH = 990
Const MONTH_BOXHEIGHT = 550
Const MONTH_BOXSTARTX = 120
Const MONTH_BOXSTARTY = 460
Const DAY_BOXWIDTH = 220
Const DAY_BOXHEIGHT = 200
Const DAY_BOXSTARTX = MONTH_BOXSTARTX + MONTH_BOXWIDTH
Const DAY_BOXSTARTY = 200
Const CALDATES_BOXWIDTH = 220
Const CALDATES_BOXHEIGHT = 200
Const CALDATES_BOXSTARTX = MONTH_BOXSTARTX + MONTH_BOXWIDTH
Const CALDATES_BOXSTARTY = MONTH_BOXSTARTY
Const CAL_BOXWIDTH = 220
Const CAL_BOXHEIGHT = 200
Const CAL_BOXSTARTX = MONTH_BOXSTARTX + MONTH_BOXWIDTH
Const CAL_BOXSTARTY = MONTH_BOXSTARTY + CAL_BOXWIDTH - TWIPS

'AE20071108 Fault #12547
Const CAL_FONTSIZE_SMALL = 5.5
Const CAL_FONTSIZE_NORMAL = 6.75

' Global Information
Public gdtmStartMonth As Date ' Holds which month the calendar starts on

Public dtmAbsStartDate As Date
Public strAbsStartSession As String
Public dtmAbsEndDate As Date
Public strAbsEndSession As String
Public strAbsType As String
Public strAbsCalendarCode As String
Public strAbsCode As String
Public strAbsWPattern As String
'Public strCurrentWorkingPattern As String
'Public strCurrentRegion As String
'Public gsRegionInfoSource As String

Public Const glngColour_Weekend = 13405581
Public Const glngColour_BankHoliday = 13405581
Public Const glngColour_Default = 16761024

Public Function DrawMonths() As Boolean
  
  ' Draws the months on the form
  
  On Error GoTo DrawMonths_ERROR
    
  Dim count_y As Integer
 
  
  ' Set the information for the month that is already there
  frmAbsenceCalendar.lblMonth(1).Caption = Space(24 - Len(Format(DateAdd("m", 0, gdtmStartMonth), "mmmm"))) & UCase(Format(DateAdd("m", 0, gdtmStartMonth), "mmmm"))
  frmAbsenceCalendar.lblMonth(1).Tag = Month(DateAdd("m", 0, gdtmStartMonth))
  
  ' Create the other 11 months
  For count_y = 2 To 12
    If frmAbsenceCalendar.lblMonth.Count < count_y Then Load frmAbsenceCalendar.lblMonth(count_y)
    frmAbsenceCalendar.lblMonth(count_y).Top = MONTH_BOXSTARTY + (MONTH_BOXHEIGHT * (count_y - 1))
    frmAbsenceCalendar.lblMonth(count_y).Visible = True
    frmAbsenceCalendar.lblMonth(count_y).Width = MONTH_BOXWIDTH
    frmAbsenceCalendar.lblMonth(count_y).Height = MONTH_BOXHEIGHT
    frmAbsenceCalendar.lblMonth(count_y).Caption = Space(24 - Len(Format(DateAdd("m", count_y - 1, gdtmStartMonth), "mmmm"))) & UCase(Format(DateAdd("m", count_y - 1, gdtmStartMonth), "mmmm"))
    frmAbsenceCalendar.lblMonth(count_y).Tag = Month(DateAdd("m", count_y - 1, gdtmStartMonth))
  Next count_y

  DrawMonths = True
  Exit Function
    
DrawMonths_ERROR:
    
  ' Theres been an error
  DrawMonths = False
  MsgBox "An error has occurred whilst drawing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function DrawDays() As Boolean
  
  ' Draws the days on the form (M, T, W, T, F, S, S)
  
  On Error GoTo DrawDays_ERROR
  
  Dim count_x As Integer
  
  For count_x = 2 To 37
    Load frmAbsenceCalendar.lblDay(count_x)
    frmAbsenceCalendar.lblDay(count_x).Visible = True
    frmAbsenceCalendar.lblDay(count_x).Width = DAY_BOXWIDTH
    frmAbsenceCalendar.lblDay(count_x).Height = DAY_BOXHEIGHT
    frmAbsenceCalendar.lblDay(count_x).Left = (DAY_BOXSTARTX + (DAY_BOXWIDTH * (count_x - 1))) - (TWIPS * (count_x) - 1)
    frmAbsenceCalendar.lblDay(count_x).Caption = UCase(Left(Format(DateAdd("d", count_x - 1, giWeekdayStart), "ddd"), 1))
    frmAbsenceCalendar.lblDay(count_x).Tag = IIf(frmAbsenceCalendar.lblDay(count_x - 1).Tag = 7, 1, frmAbsenceCalendar.lblDay(count_x - 1).Tag + 1)
  Next count_x

  DrawDays = True
  Exit Function
    
DrawDays_ERROR:
    
  ' Theres been an error
  DrawDays = False
  MsgBox "An error has occurred whilst drawing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function DrawCalDates() As Boolean
  
  ' Draws the Date boxes on the form
  
  On Error GoTo DrawCalDates_ERROR
  
  Dim count_x As Integer
  Dim count_y As Integer
  Dim lngIndex As Long
  
  For count_x = 1 To 37
    For count_y = 1 To 12
      lngIndex = (count_y * 100) + count_x
      Load frmAbsenceCalendar.lblCalDates(lngIndex)

      With frmAbsenceCalendar.lblCalDates.Item(lngIndex)
        .Visible = True
        .Move (CALDATES_BOXSTARTX + (CALDATES_BOXWIDTH * (count_x - 1))) - (TWIPS * (count_x) - 1), _
              frmAbsenceCalendar.lblMonth(count_y).Top
        .Tag = lngIndex
      End With

    Next count_y
  Next count_x

  DrawCalDates = True
  Exit Function
  
DrawCalDates_ERROR:

  ' Theres been an error
  DrawCalDates = False
  MsgBox "An error has occurred whilst drawing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function DrawCal() As Boolean
  
  On Error GoTo DrawCal_ERROR
  
  ' Draws the individual calendar day boxes (2 per day, one for AM, one for PM)
  
  Dim count_x As Integer, count_y As Integer, count_month As Integer, Count As Integer
  Dim intIndex As Integer, intLeft As Integer, intTop As Integer

  For count_month = 1 To 12
    intTop = frmAbsenceCalendar.lblMonth(count_month).Top

    For count_x = 0 To 36
      intLeft = (CAL_BOXSTARTX + ((CAL_BOXWIDTH - TWIPS) * count_x)) - TWIPS
      
      For Count = 1 To 2
        
        intIndex = (count_month * 100) + (count_x * 2) + Count
        Load frmAbsenceCalendar.lblCal(intIndex)
        'frmAbsenceCalendar.lblCal(intIndex).Visible = True
        'frmAbsenceCalendar.lblCal(intIndex).Top = frmAbsenceCalendar.lblMonth(count_month).Top + ((CAL_BOXHEIGHT - 15) * Count)
        'frmAbsenceCalendar.lblCal(intIndex).Left = intLeft
        With frmAbsenceCalendar.lblCal(intIndex)
          .Move intLeft, intTop + ((CAL_BOXHEIGHT - TWIPS) * Count)
          .Visible = True
          .Width = CAL_BOXWIDTH
          
          If count_month = 12 And Count = 2 Then
            .Height = CAL_BOXHEIGHT - (TWIPS * 2)
          Else
            .Height = CAL_BOXHEIGHT
          End If
        End With

      Next
    Next
  Next

  DrawCal = True
  Exit Function
  
DrawCal_ERROR:
  
  ' Theres been an error
  DrawCal = False
  MsgBox "An error has occurred whilst drawing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function DrawLines() As Boolean
  
  ' Draws the lines over the labels
  
  On Error GoTo DrawLines_ERROR
  
  Dim count_vert As Integer, count_hori As Integer
  
  For count_vert = 1 To 37
    Load frmAbsenceCalendar.linVertical(count_vert)
    frmAbsenceCalendar.linVertical(count_vert).Visible = True
    frmAbsenceCalendar.linVertical(count_vert).X1 = MONTH_BOXSTARTX + MONTH_BOXWIDTH + ((CAL_BOXWIDTH - TWIPS) * count_vert) - TWIPS
    frmAbsenceCalendar.linVertical(count_vert).X2 = MONTH_BOXSTARTX + MONTH_BOXWIDTH + ((CAL_BOXWIDTH - TWIPS) * count_vert) - TWIPS
    frmAbsenceCalendar.linVertical(count_vert).Y1 = frmAbsenceCalendar.lblMonth(1).Top
    frmAbsenceCalendar.linVertical(count_vert).Y2 = frmAbsenceCalendar.lblMonth(12).Top + MONTH_BOXHEIGHT
    frmAbsenceCalendar.linVertical(count_vert).ZOrder 0
  Next count_vert
  
  Load frmAbsenceCalendar.linHorizontal(13)
  frmAbsenceCalendar.linHorizontal(13).Visible = True
  frmAbsenceCalendar.linHorizontal(13).X1 = CAL_BOXSTARTX - TWIPS
  frmAbsenceCalendar.linHorizontal(13).X2 = CAL_BOXSTARTX - TWIPS + (37 * (CAL_BOXWIDTH - TWIPS))
  frmAbsenceCalendar.linHorizontal(13).Y1 = frmAbsenceCalendar.lblMonth(12).Top + MONTH_BOXHEIGHT - TWIPS
  frmAbsenceCalendar.linHorizontal(13).Y2 = frmAbsenceCalendar.lblMonth(12).Top + MONTH_BOXHEIGHT - TWIPS
  frmAbsenceCalendar.linHorizontal(13).BorderWidth = 1
  frmAbsenceCalendar.linHorizontal(13).ZOrder 0
      
  DrawLines = True
  Exit Function
  
DrawLines_ERROR:
  
  ' Theres been an error
  DrawLines = False
  MsgBox "An error has occurred whilst drawing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"
  
End Function

Public Function NumberOfDaysInMonth(dtInput As Date)
  
  'Return the number of days in the month
 
  Dim dtNextMonth As Date
  
  dtNextMonth = DateAdd("m", 1, dtInput)
  NumberOfDaysInMonth = Day(DateAdd("d", Day(dtNextMonth) * -1, dtNextMonth))

End Function

Public Function WeekDayMonthStart(dtInput As Date)
  
  'Pass a full date into this function and it will return the
  'vb constant for the day of the week that month started
  
  WeekDayMonthStart = Weekday(DateAdd("d", (Day(dtInput) - 1) * -1, dtInput))

End Function

Public Function ClearAll() As Boolean

  ' This function clears the caldates and cal boxes, ready to display a new year
  
  On Error GoTo ClearAll_ERROR
  
  Dim clearcount As Integer, ctl As Control, intDay As Integer
  
  ' CalDates
  For Each ctl In frmAbsenceCalendar
    Select Case ctl.Name
    Case "lblCal"
      ctl.Caption = ""
      ctl.Tag = ""
      ctl.BackColor = frmAbsenceCalendar.lblCal(9999).BackColor '&HFFC0C0
    Case "lblCalDates"
      ctl.Caption = ""
      ctl.Tag = ""
      ctl.BackColor = frmAbsenceCalendar.lblCalDates(0).BackColor '&HFF8080
    End Select
  Next ctl


  ClearAll = True
  Exit Function
  
ClearAll_ERROR:
  
  ' Theres been an error
  ClearAll = False
  MsgBox "An error has occurred whilst clearing the calendar. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function GetYearLayout() As Boolean
  
  ' This function clears the calendar and displays a new year

  On Error GoTo GetYearLayout_ERROR

  Dim intDaysInMonth As Integer, intCount As Integer, intDayOfFirst As Integer
  Dim c As Integer, temphandle As Integer
  Dim tempYear As Integer, tempflag As Boolean
  Dim intOneToTwelve As Integer
  
  ' Wipe everything first
  ClearAll
  
  If Month(gdtmStartMonth) = 1 Then
  
    ' Loop through the months
    For intCount = 1 To 12
          
      ' Find the first day of each month
      'JPD 20030828 Fault 2012
      'intDayOfFirst = (intCount * 100) + Weekday("01/" & Str(intCount) & "/" + frmAbsenceCalendar.lblCurrentYear.Caption, giWeekdayStart) - 1
      intDayOfFirst = (intCount * 100) + Weekday(ConvertSQLDateToLocale(Right("0" & Trim(Str(intCount)), 2) & "/01/" & frmAbsenceCalendar.lblCurrentYear.Caption), giWeekdayStart) - 1
                  
      'store the first day of the month in the tag of the month label
      'so can use this when populating the cal boxes with data
      frmAbsenceCalendar.lblMonth(intCount).Tag = intDayOfFirst + 1
                  
      ' Set the captions of Caldates ( from 1 to no. days in the month)
      'JPD 20030828 Fault 2012
      'For c = 1 To NumberOfDaysInMonth(CDate("01/" & Str(intCount) & "/" + frmAbsenceCalendar.lblCurrentYear.Caption))
      For c = 1 To NumberOfDaysInMonth(CDate(ConvertSQLDateToLocale(Right("0" & Trim(Str(intCount)), 2) & "/01/" & frmAbsenceCalendar.lblCurrentYear.Caption)))
        frmAbsenceCalendar.lblCalDates(intDayOfFirst + c).Caption = c
      Next c
        
    Next intCount
  
  Else
  
    ' Loop through the months
    intOneToTwelve = 0
    For intCount = Month(gdtmStartMonth) To Month(gdtmStartMonth) + 12
      intOneToTwelve = intOneToTwelve + 1
      ' Find the first day of each month
      
      If intOneToTwelve > 12 Then Exit For
      
      If intCount > 12 Then
        intCount = intCount - 12
        tempflag = True
        tempYear = frmAbsenceCalendar.lblCurrentYear.Caption + 1
      End If
        
      'JPD 20030828 Fault 2012
      'intDayOfFirst = (intonetotwelve * 100) + Weekday("01/" & Str(intCount) & "/" + Str(IIf(tempflag = True, tempYear, frmAbsenceCalendar.lblCurrentYear.Caption)), vbMonday) - 1
      intDayOfFirst = (intOneToTwelve * 100) + Weekday(ConvertSQLDateToLocale(Right("0" & Trim(Str(intCount)), 2) & "/01/" & Trim(Str(IIf(tempflag = True, tempYear, frmAbsenceCalendar.lblCurrentYear.Caption)))), giWeekdayStart) - 1
                  
      'store the first day of the month in the tag of the month label
      'so can use this when populating the cal boxes with data
      frmAbsenceCalendar.lblMonth(intOneToTwelve).Tag = intDayOfFirst + 1
                  
      ' Set the captions of Caldates ( from 1 to no. days in the month)
      
      'JPD 20030828 Fault 2012
      'For c = 1 To NumberOfDaysInMonth(CDate("01/" & Str(intCount) & "/" + Str(IIf(tempflag = True, tempYear, frmAbsenceCalendar.lblCurrentYear.Caption))))
      For c = 1 To NumberOfDaysInMonth(CDate(ConvertSQLDateToLocale(Right("0" & Trim(Str(intCount)), 2) & "/01/" & Trim(Str(IIf(tempflag = True, tempYear, frmAbsenceCalendar.lblCurrentYear.Caption))))))
        frmAbsenceCalendar.lblCalDates(intDayOfFirst + c).Caption = c
      Next c
    
    Next intCount
    
  End If
  
  GetYearLayout = True
  Exit Function

GetYearLayout_ERROR:
  ' Theres been an error
  GetYearLayout = False
  MsgBox "An error has occurred whilst retrieving the year layout. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

Public Function GetCalIndex(dtmDate As Date, booSession As Boolean) As Integer

  ' This function returns the index value of the cal label for the specified date.
  '
  ' INPUTS :
  '
  ' dtmDate     - the date we want to find the index for
  ' booSession  - False for AM, True for PM
  '
  ' OUTPUT :
  '
  ' GetCalIndex - the index (integer)
  '
  
  Dim intFirstDayIndex As Integer, temphundred As Integer, Temp As Integer
  Dim diff As Integer, intDay As Integer
  
  If dtmDate < frmAbsenceCalendar.lblFirstDate Or dtmDate > frmAbsenceCalendar.lblLastDate Then
    GetCalIndex = 9999
    Exit Function
  End If
    
  ' Allow for the offset if the calendar does not start in january * IS THERE AN ERROR IN HERE ???
  If Month(dtmDate) < Month(gdtmStartMonth) Then
    intFirstDayIndex = frmAbsenceCalendar.lblMonth(Month(dtmDate) + ((12 - Month(gdtmStartMonth)) + 1)).Tag
  Else
    intFirstDayIndex = frmAbsenceCalendar.lblMonth(Month(dtmDate) - (Month(gdtmStartMonth) - 1)).Tag
  End If
  
  ' Find the index of the first day of the relevant month in the correct column
  temphundred = CInt(intFirstDayIndex / 100)
  Temp = intFirstDayIndex
  intFirstDayIndex = ((Temp * 2) - (temphundred * 100)) - 1
  intDay = Day(dtmDate)
  
  ' Determine the index depending on whether session is am or pm
  If Not booSession Then
    'am
    diff = intDay - 1
    GetCalIndex = intFirstDayIndex + (2 * diff)
  Else
    'pm
    diff = intDay - 1
    GetCalIndex = (intFirstDayIndex + (2 * diff)) + 1
  End If

End Function

Public Function FillDifferentMonths() As Boolean
  
  ' This function fills in the cal boxes for all dates in the months covered between the
  ' start and end dates of the absence
  
  Dim intNumberOfDifferentCovered As Integer, intCount As Integer
  Dim tempdate As Date, tempdate2 As Date
  Dim intStart As Integer, intEnd As Integer
  
  FillDifferentMonths = True
  
  ' How many different months are coverered between the start and end date ?
  intNumberOfDifferentCovered = CInt(DateDiff("m", dtmAbsStartDate, dtmAbsEndDate))
  
  ' Loop through each month
  For intCount = 0 To intNumberOfDifferentCovered
    
    ' Work out the indexes of the cal boxes to fill for each month
    If intCount = 0 Then
      intStart = GetCalIndex(DateAdd("m", intCount, dtmAbsStartDate), IIf(strAbsStartSession = "AM", False, True))
      
      'JPD 20030828 Fault 2012
      'intEnd = GetCalIndex(NumberOfDaysInMonth(DateAdd("m", intCount, dtmAbsStartDate)) & "/" & Month(DateAdd("m", intCount, dtmAbsStartDate)) & "/" & Year(DateAdd("m", intCount, dtmAbsStartDate)), True)
      intEnd = GetCalIndex(ConvertSQLDateToLocale(Right("0" & Trim(Str(Month(DateAdd("m", intCount, dtmAbsStartDate)))), 2) & "/" & NumberOfDaysInMonth(DateAdd("m", intCount, dtmAbsStartDate)) & "/" & Year(DateAdd("m", intCount, dtmAbsStartDate))), True)
    Else
      tempdate = DateAdd("m", intCount, dtmAbsStartDate)
      'tempdate = "01" & (Right(Str(tempdate), 6))
      
      'JPD 20030828 Fault 2012
      'If Len(CStr(dtmAbsStartDate)) = 8 Then
      '  tempdate = "01" & (Right(Str(tempdate), 6))
      'Else
      '  tempdate = "01" & (Right(Str(tempdate), 8))
      'End If
      tempdate = DateAdd("d", 1 - Day(tempdate), tempdate)
      
      intStart = GetCalIndex(tempdate, False)
      
      tempdate2 = DateAdd("m", intCount, dtmAbsStartDate)
      'tempdate2 = Str(NumberOfDaysInMonth(DateAdd("m", intCount, dtmAbsStartDate))) & (Right(Str(tempdate2), 6))
      'JPD 20030828 Fault 2012
      'tempdate2 = Str(NumberOfDaysInMonth(DateAdd("m", intCount, dtmAbsStartDate))) & (Right(Str(tempdate2), IIf(Len(CStr(tempdate2)) = 8, 6, 8)))
      tempdate2 = DateAdd("m", 1, tempdate)
      tempdate2 = DateAdd("d", -1 * Day(tempdate2), tempdate2)
      
      If CDate(dtmAbsEndDate) <= tempdate2 Then
        intEnd = GetCalIndex(dtmAbsEndDate, IIf(strAbsEndSession = "AM", False, True))
      Else
        intEnd = GetCalIndex(tempdate2, True)
      End If
    End If
    
    ' Fill the cal boxes
    If Not FillCalBoxes(intStart, intEnd) Then FillDifferentMonths = False
  
  Next intCount
  
End Function

Public Function FillSameMonths() As Boolean

  ' This function fills the cal boxes between two dates in the same month
  
  Dim intStart As Integer, intEnd As Integer
  
  FillSameMonths = True
  intStart = GetCalIndex(dtmAbsStartDate, IIf(strAbsStartSession = "AM", False, True))
  intEnd = GetCalIndex(dtmAbsEndDate, IIf(strAbsEndSession = "AM", False, True))
  If Not FillCalBoxes(intStart, intEnd) Then FillSameMonths = False

End Function

Public Function FillCalBoxes(intStart As Integer, intEnd As Integer) As Boolean

  ' This function actually fills the cal boxes between the indexes specified
  ' according to the options selected by the user.
  
  On Error GoTo Error_FillCalBoxes
  
  Dim Count As Integer
  Dim dtmNextChangeDate As Date
  Dim dtmCurrentDate As Date
  Dim strCurrentSession As String
  Dim rstHistoricWPatterns As Recordset
  
  Dim blnIsBankHoliday As Boolean
  Dim blnIsWeekend As Boolean
  Dim blnHasEvent As Boolean
  Dim blnIsWorkingDay As Boolean
  Dim strColour As String
  Dim strCaption As String
  
  ' Loop through the indexes as specified.
  For Count = intStart To intEnd Step 1
    
    ' Set current date variable
    dtmCurrentDate = GetCalDay(Count)
    strCurrentSession = IIf(Count Mod 2 = 0, "PM", "AM")
    
    If frmAbsenceCalendar.WPsEnabled Then
    
      ' If we are using historic WPattern, ensure we use the right WPattern for each day of absence
      If gwptWorkingPatternType = wptHistoricWPattern Then
        'Only bother doing this guff if the date is after the next change of WPattern date
        If (dtmCurrentDate >= dtmNextChangeDate) And dtmCurrentDate <> ConvertSQLDateToLocale("12/31/9999") Then
  
          ' Get the wpattern for the start of the absence period
          Set rstHistoricWPatterns = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " AS 'Date', " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & _
                                                    "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & _
                                                   "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & frmAbsenceCalendar.PersonnelRecordID & " " & _
                                                   "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " <= '" & Replace(Format(dtmCurrentDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                   "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")
          If rstHistoricWPatterns.BOF And rstHistoricWPatterns.EOF Then
            'strAbsWPattern = ""
            dtmNextChangeDate = ConvertSQLDateToLocale("12/31/9999")
          Else
            strAbsWPattern = IIf(IsNull(rstHistoricWPatterns.Fields("WP").Value), "              ", rstHistoricWPatterns.Fields("WP").Value)
  
            ' JDM - Fault reported by Virgin. Causes problems with days if pattern is now 14 characters long
            strAbsWPattern = strAbsWPattern & Space(14 - Len(strAbsWPattern))
  
            ' Now get the date of next change
            Set rstHistoricWPatterns = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " AS 'Date', " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & _
                                                      "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & _
                                                     "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & frmAbsenceCalendar.PersonnelRecordID & " " & _
                                                     "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " > '" & Replace(Format(dtmCurrentDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                     "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")
            If rstHistoricWPatterns.EOF Then
              dtmNextChangeDate = ConvertSQLDateToLocale("12/31/9999")
            Else
              dtmNextChangeDate = rstHistoricWPatterns.Fields("Date").Value
            End If
          End If
          Set rstHistoricWPatterns = Nothing
        End If
      Else
        strAbsWPattern = frmAbsenceCalendar.ASRWorkingPattern1.Value
      End If
    Else
      'WPs DISABLED
      strAbsWPattern = "SSMMTTWWTTFFSS"
    End If
    
    With frmAbsenceCalendar
    
      blnIsBankHoliday = frmAbsenceCalendar.AbsCal_IsDayABankHoliday(Count)
      blnIsWeekend = ((Weekday(dtmCurrentDate) = vbSaturday) Or (Weekday(dtmCurrentDate) = vbSunday))
      strColour = GetColour(strAbsType)
      strCaption = Replace(strAbsCalendarCode, "&", "&&")
      blnHasEvent = True
      blnIsWorkingDay = AbsCal_DoTheyWorkOnThisDay(strAbsWPattern, dtmCurrentDate, strCurrentSession)
 
      If .lblCal(Count).Tag = "HAS_EVENT" Then
        .lblCal(Count).BackColor = vbWhite
        .lblCal(Count).ForeColor = vbBlack
        If .ShowCaptions Then
          .lblCal(Count).Caption = "."
        Else
          .lblCal(Count).Caption = vbNullString
        End If
      Else
      
        If blnHasEvent And (Not blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
          If .WorkingDaysOnly Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          Else
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          End If
        End If
      
        If blnHasEvent And (blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
          If .WorkingDaysOnly And .ShowWeekends Then
            'Weekend
            .lblCal(Count).BackColor = glngColour_Weekend
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf .WorkingDaysOnly And (Not .ShowWeekends) Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          Else
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          End If
        End If
        
        If blnHasEvent And (blnIsWeekend) And (blnIsBankHoliday) And (Not blnIsWorkingDay) Then
          If .IncludeBankHolidays Then
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          ElseIf (Not .IncludeBankHolidays) And .ShowBankHolidays Then
            'Bank Holiday
            .lblCal(Count).BackColor = glngColour_BankHoliday
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf (Not .IncludeBankHolidays) And .ShowWeekends Then
            'Weekend
            .lblCal(Count).BackColor = glngColour_Weekend
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf (Not .IncludeBankHolidays) And (Not .ShowBankHolidays) Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf (Not .IncludeBankHolidays) And (Not .ShowWeekends) Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf .WorkingDaysOnly And .ShowBankHolidays Then
            'Weekend
            .lblCal(Count).BackColor = glngColour_Weekend
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf .WorkingDaysOnly And (Not .ShowWeekends) Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf .WorkingDaysOnly And .ShowWeekends Then
            'Weekend
            .lblCal(Count).BackColor = glngColour_Weekend
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          ElseIf .WorkingDaysOnly And (Not .ShowBankHolidays) Then
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          
          Else
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          End If
        End If
        
        If blnHasEvent And (Not blnIsBankHoliday) And (blnIsWorkingDay) Then
          'Colour & Caption
          .lblCal(Count).BackColor = strColour
          .lblCal(Count).ForeColor = vbBlack
          If .ShowCaptions Then
            .lblCal(Count).Caption = strCaption
          Else
            .lblCal(Count).Caption = vbNullString
          End If
          'Set event flag, used for setting multiple events.
          .lblCal(Count).Tag = "HAS_EVENT"
        End If
        
        If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (Not blnIsWorkingDay) Then
          If .IncludeBankHolidays Then
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          ElseIf .ShowBankHolidays Then
            'Bank Holiday
            .lblCal(Count).BackColor = glngColour_BankHoliday
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          Else
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          End If
        End If
        
        If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (blnIsWorkingDay) Then
          If .IncludeBankHolidays Then
            'Colour & Caption
            .lblCal(Count).BackColor = strColour
            .lblCal(Count).ForeColor = vbBlack
            If .ShowCaptions Then
              .lblCal(Count).Caption = strCaption
            Else
              .lblCal(Count).Caption = vbNullString
            End If
            'Set event flag, used for setting multiple events.
            .lblCal(Count).Tag = "HAS_EVENT"
          ElseIf .ShowBankHolidays Then
            'Bank Holiday
            .lblCal(Count).BackColor = glngColour_BankHoliday
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          Else
            'Default
            .lblCal(Count).BackColor = glngColour_Default
            .lblCal(Count).ForeColor = vbBlack
            .lblCal(Count).Caption = vbNullString
          End If
        End If
      
      End If
      
      'AE20071108 Fault #12547
      If Len(.lblCal(Count).Caption) > 1 Then
        .lblCal(Count).FontSize = CAL_FONTSIZE_SMALL
      Else
        .lblCal(Count).FontSize = CAL_FONTSIZE_NORMAL
      End If
      .lblCal(Count).Alignment = vbCenter
    End With
      
  Next Count
  
  FillCalBoxes = True
  
  Exit Function
  
Error_FillCalBoxes:
  FillCalBoxes = False

End Function

Public Function GetCalDay(intIndex As Integer) As Date

  ' This function returns the day value of the cal label for the specified index.
  '
  ' INPUTS :
  '
  ' intIndex    - the index we want to find the day for
  '
  ' OUTPUT :
  '
  ' GetCalDay   - the day (integer)
  '
  
  Dim intGridRow As Integer, intMonth As Integer, intStartIndex As Integer
  Dim inttempindex As Integer

'# find the month for the cal box clicked on by the user !
  
  ' what row have we clicked on ?
  intGridRow = (intIndex \ 100)
  
  ' whats the start index of the month ?
  intStartIndex = frmAbsenceCalendar.lblMonth(intGridRow).Tag
  
  ' Allow for the offset if the calendar does not start in january
  intMonth = intGridRow + Month(gdtmStartMonth) - 1
  If intMonth > 12 Then
  intMonth = intMonth - 12
  End If

'# intMonth now contains the month

inttempindex = IIf(CInt(intIndex Mod 2) = 0, intIndex, intIndex + 1)
inttempindex = (inttempindex - (intGridRow * 100)) / 2
inttempindex = inttempindex + (intGridRow * 100)

'# inttempindex now contains the day

If frmAbsenceCalendar.lblCalDates(inttempindex).Caption <> "" Then
  If intGridRow <= intMonth Then
    
    'GetCalDay = CDate(frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & CStr(intMonth) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption)
    
    'JPD 20030828 Fault 2012
    'If Left(Format("01/31/00", DateFormat), 2) = "31" Then
    '  GetCalDay = CDate(frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & CStr(intMonth) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption)
    'Else
    '  GetCalDay = CDate(CStr(intMonth) & "/" & frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & frmAbsenceCalendar.lblCurrentYear.Caption)
    'End If
    GetCalDay = CDate(ConvertSQLDateToLocale(Right("0" & Trim(CStr(intMonth)), 2) & "/" & Right("0" & Trim(frmAbsenceCalendar.lblCalDates(inttempindex).Caption), 2) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption))
  
  Else
    
    'GetCalDay = CDate(frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & CStr(intMonth) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption + 1)
  
    'JPD 20030828 Fault 2012
    'If Left(Format("01/31/00", DateFormat), 2) = "31" Then
    '  GetCalDay = CDate(frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & CStr(intMonth) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption + 1)
    'Else
    '  GetCalDay = CDate(CStr(intMonth) & "/" & frmAbsenceCalendar.lblCalDates(inttempindex).Caption & "/" & frmAbsenceCalendar.lblCurrentYear.Caption + 1)
    'End If
    GetCalDay = CDate(ConvertSQLDateToLocale(Right("0" & Trim(CStr(intMonth)), 2) & "/" & Right("0" & Trim(frmAbsenceCalendar.lblCalDates(inttempindex).Caption), 2) & "/" & frmAbsenceCalendar.lblCurrentYear.Caption + 1))
  
  End If
Else
  'JPD 20030828 Fault 2012
  'GetCalDay = CDate("31/12/9999")
  GetCalDay = CDate(ConvertSQLDateToLocale("12/31/9999"))

End If

End Function

Private Function GetColour(strType As String) As OLE_COLOR

  ' This function returns the colour for the specified absence type.
  ' Derived from the key. If it cannot be found, then it defaults to
  ' The colour for 'Other' which is Black
  
  Dim ctl As Control
  
  For Each ctl In frmAbsenceCalendar.lblColourKey_Type
    If ctl.Index <> 9999 Then
      If UCase(RTrim(ctl.Caption)) = UCase(RTrim(strType)) Then
        GetColour = frmAbsenceCalendar.lblColourKey_Colour(ctl.Index).BackColor
        Exit Function
      End If
    End If
  Next ctl
  
  GetColour = vbBlack
  
End Function

Public Sub AbsCal_GetFirstAndLastViewedDates()

  Dim sDateFormat
  
  sDateFormat = LCase(DateFormat)
  
  If InStr(sDateFormat, "yyyy") Then
    sDateFormat = Replace(sDateFormat, "yyyy", frmAbsenceCalendar.lblCurrentYear.Caption)
  Else
    sDateFormat = Replace(sDateFormat, "yy", frmAbsenceCalendar.lblCurrentYear.Caption)
  End If
  
  sDateFormat = Replace(sDateFormat, "mm", frmAbsenceCalendar.cboStartMonth.ListIndex + 1)
  
  sDateFormat = Replace(sDateFormat, "dd", "01")
    
  'First Date Label
  frmAbsenceCalendar.lblFirstDate.Caption = sDateFormat
  'frmAbsenceCalendar.lblFirstDate.Caption = "01/" & (Month(gdtmStartMonth)) & "/" & (frmAbsenceCalendar.lblCurrentYear.Caption)

  'Last Date Label
  'JPD 20030828 Fault 2012
  sDateFormat = LCase(DateFormat)
  
  If InStr(sDateFormat, "yyyy") Then
    sDateFormat = Replace(sDateFormat, "yyyy", frmAbsenceCalendar.lblCurrentYear.Caption + IIf(Month(gdtmStartMonth) = 1, 0, 1))
  Else
    sDateFormat = Replace(sDateFormat, "yy", frmAbsenceCalendar.lblCurrentYear.Caption + IIf(Month(gdtmStartMonth) = 1, 0, 1))
  End If

  sDateFormat = Replace(sDateFormat, "mm", 12 - IIf(Month(gdtmStartMonth) = 1, 0, (12 - Month(gdtmStartMonth) + 1)))
  
  Select Case Month(gdtmStartMonth)
    Case 1:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 2:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 3:
      If (frmAbsenceCalendar.lblCurrentYear.Caption + 1) Mod 4 = 0 Then
        sDateFormat = Replace(sDateFormat, "dd", "29")
      Else
        sDateFormat = Replace(sDateFormat, "dd", "28")
      End If
    Case 4:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 5:
      sDateFormat = Replace(sDateFormat, "dd", "30")
    Case 6:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 7:
      sDateFormat = Replace(sDateFormat, "dd", "30")
    Case 8:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 9:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 10:
      sDateFormat = Replace(sDateFormat, "dd", "30")
    Case 11:
      sDateFormat = Replace(sDateFormat, "dd", "31")
    Case 12:
      sDateFormat = Replace(sDateFormat, "dd", "30")
  End Select
  
  frmAbsenceCalendar.lblLastDate.Caption = sDateFormat

End Sub

Public Function AbsCal_DoTheyWorkOnThisDay(strWorkingPattern As String, pdtDate As Date, strperiod As String) As Boolean
  
  ' Inputs  - 1 to 7 depending on the weekday 1 = sunday etc, "AM" or "PM"
  ' Outputs - True/False
  
  Dim intWeekDay As String

  intWeekDay = Weekday(pdtDate, vbSunday)
  
  Select Case UCase(strperiod)
  Case "AM"
    If Mid(strWorkingPattern, (intWeekDay * 2) - 1, 1) = " " Then
      AbsCal_DoTheyWorkOnThisDay = False
    Else
      AbsCal_DoTheyWorkOnThisDay = True
    End If
  Case "PM"
    If Mid(strWorkingPattern, (intWeekDay * 2), 1) = " " Then
      AbsCal_DoTheyWorkOnThisDay = False
    Else
      AbsCal_DoTheyWorkOnThisDay = True
    End If
  End Select
  
End Function





