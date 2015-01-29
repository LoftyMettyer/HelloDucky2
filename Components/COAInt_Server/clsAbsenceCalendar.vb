Option Strict Off
Option Explicit On

Imports System.Globalization
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports System.Web
Imports VB = Microsoft.VisualBasic
Imports System.Text
Imports HR.Intranet.Server.Classes
Imports System.Collections.Generic
Imports System.Linq

Public Class AbsenceCalendar
	Inherits BaseForDMI

	Private Const CELLSIZE As Integer = 17
	Private Const FULL_WP As String = "SSMMTTWWTTFFSS"

	Private mstrClientDateFormat As String

	Private mstrRealSource As String
	Private mlngPersonnelRecordID As Long
	Private mrstAbsenceRecords As DataTable
	Private mbColourKeyLoaded As Boolean

	Private mdStartDate As Date
	Private mdLeavingDate As Date
	Private mstrRegion As String
	Private mstrWorkingPattern As String

	'Private mdtmWorkingPatternDate As Date ' Next change date of the working pattern

	Private mstrHexColour_OptionBoxes As String

	Private mstrBlankSpace As String

	Private mdAbsStartDate As Date
	Private mstrAbsStartSession As String
	Private mdAbsEndDate As Date
	Private mstrAbsEndSession As String
	Private mstrAbsType As String
	Private mstrAbsCalendarCode As String
	Private mdblAbsDuration As Double
	Private mstrAbsReason As String

	Private mbDisplay_ShowBankHolidays As Boolean
	Private mbDisplay_ShowWeekends As Boolean
	Private mbDisplay_ShowCaptions As Boolean
	Private mbDisplay_IncludeBankHolidays As Boolean
	Private mbDisplay_IncludeWorkingDaysOnly As Boolean

	Public mdCalendarStartDate As Date
	Public mdCalendarEndDate As Date

	Private miAbsenceRecordsFound As Integer

	Private miStrAbsenceTypes As Integer
	Dim mastrAbsenceTypes(,) As String ' Store the absence types (redefined later as ???,3 so as to auto clear it)
	'0 = Contains the colour
	'1 = Contains the text
	'2 = Contains the code
	'3 = Contains the caption
	'4 = Contains the calendar code
	'5 = Contains the type code

	Dim mavAbsences(733) As AbsenceBreakdownDate	' Stores each of the absence cells

	Public mavWorkingPatternChanges As New List(Of WorkingPatternChange)

	' 0 = Contains the date of change
	' 1 = Contains the working pattern

	'***************************************************************************************
	Private mblnDisableWPs As Boolean

	Private mblnDisableRegions As Boolean
	Private mblnFailReport As Boolean

	Private mstrSQLSelect_RegInfoRegion As String

	Private mstrSQLSelect_PersonnelStaticRegion As String
	Private mstrSQLSelect_PersonnelStaticWP As String

	Private mstrSQLSelect_AbsenceStartDate As String
	Private mstrSQLSelect_AbsenceStartSession As String
	Private mstrSQLSelect_AbsenceEndDate As String
	Private mstrSQLSelect_AbsenceEndSession As String
	Private mstrSQLSelect_AbsenceType As String
	Private mstrSQLSelect_AbsenceReason As String
	Private mstrSQLSelect_AbsenceDuration As String

	Private mstrSQLSelect_AbsenceTypeCode As String
	Private mstrSQLSelect_AbsenceTypeCalCode As String

	Private mstrSQLSelect_PersonnelStartDate As String
	Private mstrSQLSelect_PersonnelLeavingDate As String

	Private mvarTableViews(,) As Object
	Private mobjTableView As TablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges

	Private mstrAbsenceTableRealSource As String

	Private mstrErrorMSG As String

	Public ReadOnly Property DisableRegions() As Boolean
		Get
			DisableRegions = mblnDisableRegions
		End Get
	End Property

	Public ReadOnly Property DisableWPs() As Boolean
		Get
			DisableWPs = mblnDisableWPs
		End Get
	End Property

	Public ReadOnly Property ErrorMSG() As String
		Get
			ErrorMSG = mstrErrorMSG
		End Get
	End Property

	Public ReadOnly Property ReportFailed() As Boolean
		Get
			ReportFailed = mblnFailReport
		End Get
	End Property

	' Used by the ASP to calculate the default start month of the absence calendar
	Public Property StartMonth() As Integer
		Get
			StartMonth = Month(mdCalendarStartDate)
		End Get
		Set(ByVal Value As Integer)

			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object piStartMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mdCalendarStartDate = DateSerial(Year(mdCalendarStartDate), Value, 1)
				mdCalendarEndDate = DateTime.FromOADate(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mdCalendarStartDate).ToOADate - DateTime.FromOADate(0.5).ToOADate)

			Else
				mdCalendarStartDate = DateSerial(Year(mdCalendarStartDate), AbsenceModule.giAbsenceCalStartMonth, 1)
				mdCalendarEndDate = DateTime.FromOADate(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mdCalendarStartDate).ToOADate - DateTime.FromOADate(0.5).ToOADate)

			End If

		End Set
	End Property

	' Used by the ASP to calculate the default start year of the absence calendar
	Public Property StartYear() As Integer
		Get
			StartYear = Year(mdCalendarStartDate)
		End Get
		Set(ByVal Value As Integer)

			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object piStartYear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mdCalendarStartDate = DateSerial(Value, Month(mdCalendarStartDate), 1)
				mdCalendarEndDate = DateTime.FromOADate(DateAdd(DateInterval.Year, 1, mdCalendarStartDate).ToOADate - DateTime.FromOADate(0.5).ToOADate)
			End If

		End Set
	End Property

	Public WriteOnly Property RecordID() As Long
		Set(ByVal Value As Long)

			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object piRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mlngPersonnelRecordID = Value
			End If

		End Set
	End Property

	Public WriteOnly Property ClientDateFormat() As String
		Set(ByVal value As String)
			' Clients date format passed in from the asp page
			mstrClientDateFormat = value
		End Set
	End Property

	' How many absence records were found
	Public ReadOnly Property AbsenceRecordCount() As Integer
		Get
			Return miAbsenceRecordsFound
		End Get
	End Property

	Public WriteOnly Property RealSource() As String
		Set(ByVal value As String)

			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Not IsNothing(value) Then
				mstrRealSource = value
			End If

		End Set
	End Property

	Public WriteOnly Property ShowWeekends() As String
		Set(ByVal Value As String)
			' Are the weekends to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_ShowWeekends = IIf(Value = "highlighted", True, IIf(IsNothing(Value), AbsenceModule.gfAbsenceCalWeekendShading, False))
		End Set
	End Property

	Public WriteOnly Property ShowCaptions() As String
		Set(ByVal Value As String)
			' Are the captions to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_ShowCaptions = IIf(Value = "show", True, IIf(IsNothing(Value), AbsenceModule.gfAbsenceCalShowCaptions, False))
		End Set
	End Property

	Public WriteOnly Property ShowBankHolidays() As String
		Set(ByVal Value As String)
			' Are the bank holidays to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_ShowBankHolidays = IIf(Value = "highlighted", True, IIf(IsNothing(Value), AbsenceModule.gfAbsenceCalBHolShading, False))
		End Set
	End Property

	Public WriteOnly Property IncludeBankHolidays() As String
		Set(ByVal Value As String)
			' Are the bank holidays to be included (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_IncludeBankHolidays = IIf(Value = "included", True, IIf(IsNothing(Value), AbsenceModule.gfAbsenceCalBHolInclude, False))
		End Set
	End Property

	Public WriteOnly Property IncludeWorkingDaysOnly() As String
		Set(ByVal Value As String)
			' Are the working days only to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_IncludeWorkingDaysOnly = IIf(Value = "included", True, IIf(IsNothing(Value), AbsenceModule.gfAbsenceCalIncludeWorkingDaysOnly, False))
		End Set
	End Property

	' Used by the ASP to calculate the whether we have access to the absence table
	Public ReadOnly Property AbsenceTableAccess() As Boolean
		Get
			Return mblnFailReport
		End Get
	End Property

	' Used by the ASP to calculate the whether we have access to the working pattern table
	Public ReadOnly Property WPTableAccess() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object WPTableAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			WPTableAccess = If(mblnDisableWPs, "0", "1")
		End Get
	End Property

	' Converts RGB value into a hex code for IExplorer
	Private Function GetHexColour(iRed As Integer, iGreen As Integer, iBlue As Integer) As String
		Return "#" & Right("0" & Hex(iRed), 2) & Right("0" & Hex(iGreen), 2) & Right("0" & Hex(iBlue), 2)
	End Function

	' Load the defaults
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		ReDim mvarTableViews(3, 0)

		mstrHexColour_OptionBoxes = "ThreeDFace"

		' A blank cell
		mstrBlankSpace = "<TD HEIGHT=" & CELLSIZE & " WIDTH=" & CELLSIZE & ">&nbsp;</TD>"

	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	Public Function HTML_SelectedStartMonthCombo(ByVal piStartMonth As Integer) As String

		'Build month selection dropdown combo
		Dim iCount As Integer
		Dim strHtml As String

		'UPGRADE_WARNING: Couldn't resolve default property of object piStartMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		piStartMonth = If(IsNumeric(piStartMonth), piStartMonth, AbsenceModule.giAbsenceCalStartMonth)

		'strHTML = "<SELECT id=cboStartMonth style=""HEIGHT: 22px; WIDTH: 150px"" onchange=""return cboStartMonth_onchange()"">" & vbNewLine
		strHtml = "<SELECT id=cboStartMonth onchange=""return cboStartMonth_onchange()"">" & vbNewLine

		For iCount = 1 To 12

			'UPGRADE_WARNING: Couldn't resolve default property of object piStartMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If iCount = piStartMonth Then
				strHtml = strHtml & "<OPTION selected value=" & Trim(Str(iCount)) & ">" & StrConv(MonthName(iCount), VbStrConv.ProperCase) & vbNewLine
			Else
				strHtml = strHtml & "<OPTION value=" & Trim(Str(iCount)) & ">" & StrConv(MonthName(iCount), VbStrConv.ProperCase) & vbNewLine
			End If

		Next iCount

		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_SelectedStartMonthCombo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_SelectedStartMonthCombo = strHtml & "  </SELECT>"

	End Function

	Private Function HTML_Calendar_Heading() As String

		'Build month selection dropdown combo
		Dim strHtml As String
		Dim iStartNumber As Integer
		Dim iCount As Integer
		Dim dTempDate As Date

		dTempDate = DateSerial(Year(mdCalendarStartDate), Month(mdCalendarStartDate), 1)
		iStartNumber = Weekday(dTempDate, FirstDayOfWeek.Sunday)

		strHtml = "<TR>" & mstrBlankSpace & vbNewLine & "<TD><TABLE style=""HEIGHT: 100%; WIDTH: 100%"" align=left class='invisible' cellPadding=0 cellSpacing=0 width=""100%"" height=""100%""> " & vbNewLine & "<TBODY align=center style=""FONT-SIZE: x-small"">"
		strHtml = strHtml & "<TR>"

		' Before first day of month
		dTempDate = DateTime.FromOADate(DateTime.FromOADate(dTempDate.ToOADate - iStartNumber).ToOADate + 1)
		For iCount = 0 To 36
			strHtml = strHtml & "<TD class=""smallfont"" ALIGN=center NOWRAP WIDTH=" & CELLSIZE & " HEIGHT=" & CELLSIZE & ">" & UCase(Left(VB6.Format(dTempDate, "ddd", FirstDayOfWeek.Sunday), 1)) & "</TD>" & vbNewLine
			dTempDate = DateTime.FromOADate(dTempDate.ToOADate + 1)
		Next iCount

		strHtml = strHtml & "</TR>"

		Return strHtml & "</TBODY></TABLE>" & vbNewLine & "</TD>" & vbNewLine & "</TR>" & vbNewLine

	End Function

	Private Function HTML_Month(piMonthNumber As Integer, piYear As Integer) As String

		Dim iCount As Integer
		Dim strHtml As String
		Dim strHtmlDays As String
		Dim strHtmlDaysStart As String
		Dim iStartNumber As Integer
		Dim iEndNumber As Integer
		Dim dTempDate As Date
		Dim iIndexAM As Integer
		Dim iIndexPM As Integer
		Dim strHtmlCellString As String

		strHtml = "<SPAN id=Month" & LTrim(Str(piMonthNumber)) & ">" & vbNewLine & "<TR>" _
			& "<TD class='smallfont'>&nbsp;" & MonthName(piMonthNumber) & "&nbsp;</TD>" _
			& "<TD> <TABLE class='invisible' cellPadding=0 cellSpacing=0 width=""100%"" height=""100%"">"

		' Calculate month parameters
		dTempDate = New DateTime(piYear, piMonthNumber, 1)
		iStartNumber = Weekday(dTempDate, FirstDayOfWeek.Sunday) - 1
		iEndNumber = iStartNumber + DateTime.DaysInMonth(dTempDate.Year, dTempDate.Month)

		' Draw the day numbers
		strHtml = strHtml & "<TR>"
		For iCount = 0 To 36
			If iCount >= iEndNumber Or iCount < iStartNumber Then
				strHtmlDays = "<TD class=""calendarheader_nonday"" NOWRAP width=" & CELLSIZE & " height=" & CELLSIZE & " align=center>&nbsp;</TD>" & vbNewLine
			Else
				strHtmlDays = "<TD class=""calendarheader_day"" NOWRAP width=" & CELLSIZE & " height=" & CELLSIZE & " align=center>" & Str(iCount + 1 - iStartNumber) & "</TD>" & vbNewLine
			End If

			strHtml = strHtml & strHtmlDays & vbNewLine
		Next iCount
		strHtml = strHtml & "</TR>" & vbNewLine

		' Draw the spaces for the absence types
		strHtml = strHtml & "<TR>" & vbNewLine
		For iCount = 0 To 36
			strHtmlDaysStart = "<TD><TABLE style=""HEIGHT: 100%; WIDTH: 100%"" align=left class=""calendarcell"" cellPadding=0 cellSpacing=0 width=""100%""> " & vbNewLine & "<TBODY style=""FONT-SIZE: xx-small"">"

			If iCount >= iEndNumber Or iCount < iStartNumber Then
				strHtmlDays = "<TR><TD name=DateID_9999 id=DateID_9999 class=""calendar_nonday"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD></TR>" & "<TR><TD name=DateID_9999 id=DateID_9999 class=""calendar_nonday"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD></TR>"
			Else
				dTempDate = New Date(piYear, piMonthNumber, iCount + 1 - iStartNumber)

				iIndexAM = GetCalIndex(dTempDate, False)
				iIndexPM = GetCalIndex(dTempDate, True)

				' Is a weekend
				mavAbsences(iIndexAM).IsWeekend = (Weekday(dTempDate, FirstDayOfWeek.Monday) > 5)
				mavAbsences(iIndexPM).IsWeekend = (Weekday(dTempDate, FirstDayOfWeek.Monday) > 5)

				'------------------------------------------------
				'AM
				'------------------------------------------------
				If (dTempDate < mdStartDate) Or (dTempDate > mdLeavingDate And Not mdLeavingDate = DateTime.FromOADate(0)) Then
					strHtmlCellString = "<TD name=DateID_" & iIndexAM.ToString & " id=DateID_" & iIndexAM.ToString & "class=""calendar_nonday"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD>" & vbNewLine
				Else
					'Build the cell string

					' Build onclick event code
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mavAbsences(iIndexAM).ContainsData Then
						strHtmlCellString = "<TD style='font-size: " & IIf(Len(mavAbsences(iIndexAM).Caption) < 2, "8", "6") & "pt;background-color:" & mavAbsences(iIndexAM).DisplayColor & "' name=DateID_" _
								& iIndexAM.ToString & " id=DateID_" & iIndexAM.ToString & " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP " _
								& " onclick=""ShowDetails('" & VB6.Format(mavAbsences(iIndexAM).StartDate, mstrClientDateFormat) & "','" _
								& mavAbsences(iIndexAM).StartSession & "','" & VB6.Format(mavAbsences(iIndexAM).EndDate, mstrClientDateFormat) & "','" & mavAbsences(iIndexAM).EndSession & "','" & mavAbsences(iIndexAM).Duration & "','" _
								& Replace(mastrAbsenceTypes(mavAbsences(iIndexAM).Type, 0), "'", "") & "','" & Replace(mastrAbsenceTypes(mavAbsences(iIndexAM).Type, 5), "'", "") & "','" _
								& Replace(mastrAbsenceTypes(mavAbsences(iIndexAM).Type, 4), "'", "") & "','" & HttpUtility.HtmlEncode(Left(mavAbsences(iIndexAM).Reason, 100)) & "','" & mavAbsences(iIndexAM).Region & "','" _
								& mavAbsences(iIndexAM).WorkingPattern & "'," & IIf(mavAbsences(iIndexAM).IsWorkingDay, "true", "false") & ")"">" & "<FONT SIZE='1'>" & mavAbsences(iIndexAM).Caption & "</FONT></TD>" & vbNewLine
					Else
						strHtmlCellString = "<TD name=DateID_" & iIndexAM.ToString & " id=DateID_" & iIndexAM.ToString & " class=""calendar_day"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" _
								& CELLSIZE & " NOWRAP>&nbsp;</TD>" & vbNewLine
					End If

				End If

				' Add current cell to the table
				strHtmlDays = "<TR>" & strHtmlCellString & "</TR>"


				'------------------------------------------------
				'PM
				'------------------------------------------------
				If (dTempDate < mdStartDate) Or (dTempDate > mdLeavingDate And Not mdLeavingDate = DateTime.FromOADate(0)) Then
					strHtmlCellString = "<TD name=DateID_" & iIndexPM.ToString & " id=DateID_" & iIndexPM.ToString & " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP></TD>" & vbNewLine
				Else

					' Build onclick event code
					If mavAbsences(iIndexPM).ContainsData Then
						strHtmlCellString = "<TD style='font-size: " & IIf(Len(mavAbsences(iIndexPM).Caption) < 2, "8", "6") & "pt;background-color:" & mavAbsences(iIndexPM).DisplayColor & "' name=DateID_" & LTrim(Str(iIndexPM)) & " id=DateID_" & LTrim(Str(CDbl(iIndexPM))) _
								& " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP" _
								& " onclick=""ShowDetails('" & VB6.Format(mavAbsences(iIndexPM).StartDate, mstrClientDateFormat) & "','" & mavAbsences(iIndexPM).StartSession & "','" _
								& VB6.Format(mavAbsences(iIndexPM).EndDate, mstrClientDateFormat) & "','" & mavAbsences(iIndexPM).EndSession & "','" & mavAbsences(iIndexPM).Duration _
								& "','" & Replace(mastrAbsenceTypes(mavAbsences(iIndexPM).Type, 0), "'", "") & "','" & Replace(mastrAbsenceTypes(mavAbsences(iIndexPM).Type, 5), "'", "") _
								& "','" & Replace(mastrAbsenceTypes(mavAbsences(iIndexPM).Type, 4), "'", "") & "','" & HttpUtility.HtmlEncode(Left(mavAbsences(iIndexPM).Reason, 100)) & "','" _
								& mavAbsences(iIndexPM).WorkingPattern & "'," & IIf(mavAbsences(iIndexPM).IsWorkingDay, "true", "false") & ")"">" & "<FONT SIZE='1'>" & mavAbsences(iIndexPM).Caption & "</FONT></TD>" & vbNewLine
					Else
						strHtmlCellString = "<TD name=DateID_" & LTrim(Str(iIndexPM)) & " id=DateID_" & LTrim(Str(iIndexPM)) & " class=""calendar_day"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD>"
					End If

				End If

				' Create the cell for this day session
				strHtmlDays = strHtmlDays & "<TR>" & strHtmlCellString & "</TR>"

			End If

			' Add current cell to the table
			strHtml = strHtml & strHtmlDaysStart & strHtmlDays & "</TBODY></TABLE></TD>" & vbNewLine

		Next iCount
		strHtml = strHtml & "</TR>"

		' Finish off this month HTML code
		Return strHtml & "   </TABLE>" & vbNewLine & "</TD>" & vbNewLine & "</TR>" & "</SPAN>"

	End Function

	Private Function GetAbsenceRecordSet() As Integer

		Dim sSQL As String

		On Error GoTo GetAbsenceRecordSet_ERROR

		' Get Recordset Containing Absence info for the current employee
		sSQL = "SELECT " & mstrSQLSelect_AbsenceStartDate & " as 'StartDate', " & vbNewLine & mstrSQLSelect_AbsenceStartSession & " as 'StartSession', " & vbNewLine

		If Not IsDBNull(mdLeavingDate) Then
			sSQL = sSQL & "isnull(" & mstrSQLSelect_AbsenceEndDate & ",'" & Replace(VB6.Format(mdLeavingDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "') as 'EndDate', " & vbNewLine
		Else
			sSQL = sSQL & "isnull(" & mstrSQLSelect_AbsenceEndDate & ",'" & Replace(VB6.Format(Now, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "') as 'EndDate', " & vbNewLine
		End If

		sSQL = sSQL & mstrSQLSelect_AbsenceEndSession & " as 'EndSession', " & vbNewLine & mstrSQLSelect_AbsenceType & " as 'Type', " & vbNewLine & mstrSQLSelect_AbsenceTypeCalCode & " as 'CalendarCode', " & vbNewLine & mstrSQLSelect_AbsenceTypeCode & " as 'Code', " & vbNewLine & mstrSQLSelect_AbsenceReason & " as 'Reason', " & vbNewLine & mstrSQLSelect_AbsenceDuration & " as 'Duration' " & vbNewLine

		sSQL = sSQL & "FROM " & mstrAbsenceTableRealSource & vbNewLine
		sSQL = sSQL & "           INNER JOIN " & AbsenceModule.gsAbsenceTypeTableName & vbNewLine
		sSQL = sSQL & "           ON " & mstrAbsenceTableRealSource & "." & AbsenceModule.gsAbsenceTypeColumnName & " = " & AbsenceModule.gsAbsenceTypeTableName & "." & AbsenceModule.gsAbsenceTypeTypeColumnName & vbNewLine

		sSQL = sSQL & "WHERE " & mstrAbsenceTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & vbNewLine
		sSQL = sSQL & " AND (" & mstrSQLSelect_AbsenceStartDate & " IS NOT NULL) " & vbNewLine
		sSQL = sSQL & "ORDER BY 'StartDate' ASC"

		mrstAbsenceRecords = DB.GetDataTable(sSQL)

		' Set amount of absence records found
		Return mrstAbsenceRecords.Rows.Count

GetAbsenceRecordSet_ERROR:

		'MsgBox "Error retrieving the Absence recordset." & vbNewLine & Err.Description, vbExclamation + vbOKOnly, App.Title
		'UPGRADE_NOTE: Object mrstAbsenceRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mrstAbsenceRecords = Nothing
		Return 0

	End Function

	Public Sub Initialise()

		Dim fOK As Boolean

		fOK = True

		' Read the necessary settings for the calendar to work
		mbColourKeyLoaded = False

		' Check the Module Setup and Data Permissions for the Absence Calendar Specific columns
		If fOK Then
			fOK = CheckPermission_AbsCalSpecifics()
			If Not fOK Then
				Exit Sub
			End If
		End If

		' Check the Module Setup and Data Permissions for the Regional/Bank Holiday columns
		CheckPermission_RegionInfo()

		' Check the Module Setup and Data Permissions for the Working Pattern columns
		CheckPermission_WPInfo()

		' Set the start day to 1
		mdCalendarStartDate = DateSerial(Year(mdCalendarStartDate), Month(mdCalendarStartDate), 1)

		' Populate the absence collection
		For iCount = 0 To mavAbsences.Count - 1
			mavAbsences(iCount) = New AbsenceBreakdownDate
		Next

		' Only load the records from the DB once
		GetPersonnelRecordSet()

		GetWorkingPatterns()

		'GetRegions
		miAbsenceRecordsFound = GetAbsenceRecordSet()

		LoadColourKey()

		' Default start and end dates
		mdCalendarStartDate = DateSerial(Year(Now), AbsenceModule.giAbsenceCalStartMonth, 1)
		mdCalendarEndDate = DateTime.FromOADate(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mdCalendarStartDate).ToOADate - DateTime.FromOADate(0.5).ToOADate)

	End Sub

	' Loads the absence types
	Private Function LoadColourKey() As Boolean

		' Have colour already been loaded?
		If mbColourKeyLoaded Then
			Return True
		End If

		On Error GoTo errLoadColourKey

		Dim rstColourKey As DataTable
		Dim strColourKeySQL As String
		Dim intCounter As Integer
		Dim strHexColour As String

		strColourKeySQL = "SELECT DISTINCT " & AbsenceModule.gsAbsenceTypeTypeColumnName & " AS Type, " & AbsenceModule.gsAbsenceTypeCalCodeColumnName & " AS CalCode," & AbsenceModule.gsAbsenceTypeCodeColumnName & " AS TypeCode" & " FROM " & AbsenceModule.gsAbsenceTypeTableName & " ORDER BY " & AbsenceModule.gsAbsenceTypeTypeColumnName
		rstColourKey = DB.GetDataTable(strColourKeySQL)

		If rstColourKey.Rows.Count = 0 Then
			Return False
		End If

		'ReDim Preserve mastrAbsenceTypes(rstColourKey.RecordCount + 1, 5)
		ReDim Preserve mastrAbsenceTypes(20, 5)

		intCounter = 0

		For Each objRow As DataRow In rstColourKey.Rows

			If intCounter <= 18 Then

				' Set the colour box caption and show the label
				mastrAbsenceTypes(intCounter, 0) = objRow(0).ToString()

				Select Case intCounter Mod 5
					Case 0
						strHexColour = GetHexColour(255, 192, 192)
					Case 1
						strHexColour = GetHexColour(192, 255, 192)
					Case 2
						strHexColour = GetHexColour(255, 255, 192)
					Case 3
						strHexColour = GetHexColour(255, 224, 192)
					Case 4
						strHexColour = GetHexColour(192, 255, 255)
				End Select

				mastrAbsenceTypes(intCounter, 1) = strHexColour
				mastrAbsenceTypes(intCounter, 2) = CStr(intCounter)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(intCounter, 3) = UCase(Left(IIf(IsDBNull(objRow("CalCode")), "", objRow("CalCode")), 2))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(intCounter, 4) = Replace(IIf(IsDBNull(objRow("CalCode")), "", objRow("CalCode")), "'", "")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(intCounter, 5) = Replace(IIf(IsDBNull(objRow("TypeCode")), "", objRow("TypeCode")), "'", "")

			End If

			intCounter = intCounter + 1
		Next

		' Now add the 'Other' box (if needed)
		If intCounter > 17 Then
			intCounter = IIf(intCounter > 17, 18, intCounter)
			mastrAbsenceTypes(intCounter, 0) = "Other"
			mastrAbsenceTypes(intCounter, 1) = "black"
			mastrAbsenceTypes(intCounter, 2) = LTrim(Str(intCounter))
			mastrAbsenceTypes(intCounter, 3) = "&nbsp"
			intCounter = intCounter + 1

			ReDim Preserve mastrAbsenceTypes(20, 5)

		End If

		' Now add the multiple box
		mastrAbsenceTypes(intCounter, 0) = "Multiple"
		mastrAbsenceTypes(intCounter, 1) = "white"
		mastrAbsenceTypes(intCounter, 2) = LTrim(Str(intCounter))
		mastrAbsenceTypes(intCounter, 3) = "."

		' If we are here, then notify calling procedure of success and exit
		LoadColourKey = True
		mbColourKeyLoaded = True
		miStrAbsenceTypes = intCounter

		Exit Function

errLoadColourKey:

		LoadColourKey = False

	End Function

	Public Function HTML_LoadColourKey() As String

		' Load the colour key variables
		If Not LoadColourKey() Then
			Exit Function
		End If

		Dim intCounter As Integer
		Dim strKeyText As String
		Dim strKeyColour As String
		Dim strKeyCode As String
		Dim strKeyCaption As String
		Dim bSecondColumn As Boolean

		Dim strHTML As String
		Dim strHTML_KeyType As String

		strHTML = vbNullString

		' Build start of table
		strHTML = strHTML & "<TABLE class='outline' cellPadding=0 cellSpacing=0 width=250>" & vbNewLine
		strHTML = strHTML & "<TR>" & vbNewLine
		strHTML = strHTML & "   <TD style='width: 50px;'>" & vbNewLine

		bSecondColumn = False

		For intCounter = 0 To miStrAbsenceTypes	'UBound(mastrAbsenceTypes, 1) - 1

			' Position the colour box control depending on its index
			If intCounter >= 10 And Not bSecondColumn Then
				bSecondColumn = True
				strHTML = strHTML & "   </TD>" & vbNewLine
				strHTML = strHTML & "   <TD style='width: 50px;'>" & vbNewLine
			End If

			' Set the colour box caption and show the label
			strKeyText = IIf(Len(mastrAbsenceTypes(intCounter, 0)) = 0, "&nbsp", mastrAbsenceTypes(intCounter, 0))
			strKeyColour = mastrAbsenceTypes(intCounter, 1)
			strKeyCode = mastrAbsenceTypes(intCounter, 2)
			strKeyCaption = mastrAbsenceTypes(intCounter, 3)

			' Generate HTML code for this key
			strHTML_KeyType = "<TABLE class='invisible' cellPadding=0 cellSpacing=2>" & vbNewLine _
			& " <TR>" & vbNewLine _
			& "   <TD width=" & CELLSIZE & ">" & vbNewLine _
			& "   </TD>" & vbNewLine _
			& "   <TD style='font-size: " & IIf(Len(strKeyCaption) < 2, "8", "6") & "pt;' ID=KEY_" & intCounter & " NAME=KEY_" & intCounter & " class='bordered' height=" & CELLSIZE & " width=" & CELLSIZE & " align=center valign=middle NOWRAP bgColor=""" & strKeyColour & """>" & vbNewLine _
			& IIf(Trim(strKeyCaption) = vbNullString, "&nbsp", strKeyCaption) & vbNewLine _
			& "   </TD>" & vbNewLine _
			& "   <TD>&nbsp;" & vbNewLine _
			& strKeyText & vbNewLine _
			& "   </TD>" & vbNewLine _
			& "</TR>" & vbNewLine _
			& "</TABLE>" & vbNewLine

			' Add current key to key table
			strHTML = strHTML & strHTML_KeyType

		Next intCounter

		' Pad extra blank absence types to make the entries in list line up correctly
		For intCounter = miStrAbsenceTypes + 1 To 19 'UBound(mastrAbsenceTypes, 1) + 1 To 20

			' Position the colour box control depending on its index
			If intCounter >= 10 And Not bSecondColumn Then
				bSecondColumn = True
				strHTML = strHTML & "</TD><TD bordercolor=" & mstrHexColour_OptionBoxes & ">"
			End If

			' JDM - 11/10/02 - Duplicating the ID of the last absence type
			strKeyCode = "junkpointer" & LTrim(Str(intCounter))

			'UPGRADE_WARNING: Couldn't resolve default property of object strHTML_KeyType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strHTML_KeyType = "<TABLE class=""invisible"" cellPadding=0 cellSpacing=2>" & "<TR>" & "<TD width=" & CELLSIZE & "></TD>" & "<TD width=10%>&nbsp&nbsp&nbsp&nbsp&nbsp</TD>" & "<TD></TD>" & "</TR>" & "</TABLE>"

			' Add current key to key table
			'UPGRADE_WARNING: Couldn't resolve default property of object strHTML_KeyType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strHTML = strHTML & strHTML_KeyType

		Next intCounter

		' Finish off the table text
		strHTML = strHTML & "</TD></TR></TABLE>"

		' If we are here, then notify calling procedure of success and exit
		Return strHTML

	End Function

	Public Function HTML_Calendar() As String

		Dim strHtml As String
		Dim iMonth As Integer
		Dim iYear As Integer
		Dim iCount As Integer

		' Base HTML for the table
		strHtml = "<table id=MainGrid border=0 cellPadding=0 cellSpacing=0 width=""100%""" & ">" & "<tbody>"

		' Calculate the bank holidays
		FillGridWithData()

		If Not mblnDisableRegions Then
			GenerateRegionData()
		End If

		' Add day names (MTWTFSS)
		strHtml = strHtml & HTML_Calendar_Heading()

		' HTML main code
		For iCount = 0 To 11
			iMonth = mdCalendarStartDate.AddMonths(iCount).Month
			iYear = mdCalendarStartDate.AddMonths(iCount).Year
			strHtml = strHtml & HTML_Month(iMonth, iYear)
		Next iCount

		' Finish off the table text
		strHtml = strHtml & "</tbody></table>"

		' Return HTML code for the main calendar
		Return strHtml

	End Function

	Private Sub FillGridWithData()

		Try

			' Load the colour key variables
			If Not LoadColourKey() Then
				Exit Sub
			End If

			Dim intStart As Integer
			Dim intEnd As Integer

			' If there are no absence records for the current employee then skip
			' this bit (but still show the form)
			If mrstAbsenceRecords.Rows.Count = 0 Then
				Exit Sub
			End If

			With mrstAbsenceRecords

				For Each objRow As DataRow In .Rows

					' Load each absence record data into variables
					' (has to be done because start/end dates may be modified by code to fill grid correctly)

					' JDM - Kak-Handed way of sorting out American settings on different versions of IIS
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("StartDate")) Then
						mdAbsStartDate = Now
					Else
						mdAbsStartDate = CDate(objRow("StartDate"))
					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("EndDate")) Then
						mdAbsEndDate = Now
					Else
						mdAbsEndDate = CDate(objRow("EndDate"))
					End If


					mstrAbsStartSession = objRow("StartSession").ToString.ToUpper()
					mstrAbsEndSession = objRow("EndSession").ToString.ToUpper()

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mstrAbsType = IIf(IsDBNull(objRow("Type")), "", objRow("Type"))

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mstrAbsCalendarCode = IIf(IsDBNull(objRow("CalendarCode")), "", objRow("CalendarCode"))
					mdblAbsDuration = CDbl(objRow("Duration"))
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mstrAbsReason = IIf(IsDBNull(objRow("Reason")), "", objRow("Reason").ToString)

					If mdAbsStartDate <= mdCalendarEndDate And mdAbsEndDate >= mdCalendarStartDate Then
						intStart = GetCalIndex(mdAbsStartDate, mstrAbsStartSession = "PM")
						intEnd = GetCalIndex(mdAbsEndDate, mstrAbsEndSession = "PM")

						FillCalBoxes(intStart, intEnd)
					End If

				Next
			End With

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Private Function GetPersonnelRecordSet() As Boolean

		On Error GoTo PersonnelERROR

		Dim lngCount As Integer
		Dim sSQL As String
		Dim prstPersonnelData As DataTable
		Dim strAbsWPattern As String

		' Botch as we have a lot of rubbish code that does not handle nulls at all.
		mdStartDate = DateTime.FromOADate(0)
		mdLeavingDate = DateTime.FromOADate(0)

		If Not mblnFailReport Then
			sSQL = vbNullString
			sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStartDate & " AS 'StartDate', " & vbNewLine
			sSQL = sSQL & "      " & mstrSQLSelect_PersonnelLeavingDate & " AS 'LeavingDate' " & vbNewLine
			sSQL = sSQL & "FROM " & PersonnelModule.gsPersonnelTableName & vbNewLine
			For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
				'<Personnel CODE>
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarTableViews(0, lngCount) = PersonnelModule.glngPersonnelTableID Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & "     ON  " & PersonnelModule.gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
				End If
			Next lngCount
			sSQL = sSQL & "WHERE " & PersonnelModule.gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID

			' Get the start and leaving date
			prstPersonnelData = DB.GetDataTable(sSQL)

			If prstPersonnelData.Rows.Count > 0 Then
				If Not IsDBNull(prstPersonnelData.Rows(0)("StartDate")) Then
					mdStartDate = CDate(prstPersonnelData.Rows(0)("StartDate"))
				End If

				If Not IsDBNull(prstPersonnelData.Rows(0)("LeavingDate")) Then
					mdLeavingDate = CDate(prstPersonnelData.Rows(0)("LeavingDate"))
				End If

			End If
		Else
			GoTo PersonnelERROR
		End If

		If Not mblnDisableRegions Then
			' Get the employees current region
			If PersonnelModule.grtRegionType = RegionType.rtStaticRegion Then
				' Its a static region, get it from personnel
				sSQL = "SELECT " & mstrSQLSelect_PersonnelStaticRegion & "  AS 'Region'  " & vbNewLine
				sSQL = sSQL & "FROM " & PersonnelModule.gsPersonnelTableName & vbNewLine
				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<Personnel CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = PersonnelModule.glngPersonnelTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     ON  " & PersonnelModule.gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount
				sSQL = sSQL & "WHERE " & PersonnelModule.gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID
				prstPersonnelData = DB.GetDataTable(sSQL)
			Else
				' Its a historic region, so get topmost from the history
				prstPersonnelData = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & " ORDER BY " & PersonnelModule.gsPersonnelHRegionDateColumnName & " DESC")
			End If

			If prstPersonnelData.Rows.Count > 0 Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrRegion = Replace(IIf(IsDBNull(prstPersonnelData.Rows(0)("Region")), "", IIf(prstPersonnelData.Rows(0)("Region") = "", "", prstPersonnelData.Rows(0)("Region"))), "&", "&&")
			Else
				mstrRegion = "&lt;None&gt;"
			End If
		Else
			'Regions DISABLED
			mstrRegion = vbNullString
		End If

		If Not mblnDisableWPs Then
			' Get the employees current working pattern
			If PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern Then
				' Its a static working pattern, get it from personnel
				sSQL = vbNullString
				sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStaticWP & "  AS 'WP'  " & vbNewLine
				sSQL = sSQL & "FROM " & PersonnelModule.gsPersonnelTableName & vbNewLine
				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<Personnel CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = PersonnelModule.glngPersonnelTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     ON  " & PersonnelModule.gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount
				sSQL = sSQL & "WHERE " & PersonnelModule.gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID
				prstPersonnelData = DB.GetDataTable(sSQL)

			Else
				' Its a historic working pattern, so get topmost from the history
				prstPersonnelData = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & "FROM " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & "AND " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " <= '" _
																									& Replace(VB6.Format(Now, "MM/dd/yy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & "ORDER BY " & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " DESC")
			End If

			If prstPersonnelData.Rows.Count > 0 Then
				mstrWorkingPattern = prstPersonnelData.Rows(0)("WP").ToString
			Else
				mstrWorkingPattern = Space(14)
			End If

		Else
			'WPs DISABLED
			strAbsWPattern = "SSMMTTWWTTFFSS"

		End If

		'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstPersonnelData = Nothing
		Return True

PersonnelERROR:

		Return False

	End Function

	Public Function HTML_EmployeeInformation() As String

		Dim strHtml As String

		strHtml = vbNullString

		' Region Info
		If Not mblnDisableRegions Then
			strHtml = strHtml & "<TR bordercolor=" & mstrHexColour_OptionBoxes & ">" & vbNewLine _
				& "   <TD nowrap>&nbsp;Region :</TD>" & vbNewLine _
				& "   <TD>" & mstrRegion & "</TD>" & vbNewLine _
				& "</TR>" & vbNewLine
		End If

		' Start Date Info
		strHtml = strHtml & "<TR bordercolor=" & mstrHexColour_OptionBoxes & ">" & vbNewLine _
			& "   <TD nowrap>&nbsp;Start Date :</TD>" & vbNewLine _
			& "   <TD>" & IIf(mdStartDate = DateTime.FromOADate(0), "&lt;None&gt;", VB6.Format(mdStartDate, mstrClientDateFormat)) & "</TD>" & vbNewLine _
			& "</TR>" & vbNewLine

		' Leaving Date Info
		strHtml = strHtml & "<TR bordercolor=" & mstrHexColour_OptionBoxes & ">" & vbNewLine _
			& "   <TD nowrap>&nbsp;Leaving Date :</TD>" & vbNewLine _
			& "   <TD>" & IIf(mdLeavingDate = DateTime.FromOADate(0), "&lt;None&gt;", VB6.Format(mdLeavingDate, mstrClientDateFormat)) & "</TD>" & vbNewLine _
			& "</TR>" & vbNewLine

		If Not mblnDisableWPs Then
			' Working Pattern Info
			strHtml &= "<TR bordercolor=" & mstrHexColour_OptionBoxes & "><TD nowrap>&nbsp;&nbsp;&nbsp;</TD><TD></TD></TR><tr><td>Current Working Pattern :</td><td>" _
				& HTML_WorkingPattern(mstrWorkingPattern) & "</td></tr>" & vbNewLine
		End If

		Return strHtml

	End Function

	Public Function HTML_ToggleDisplay() As String

		Dim strHtml As String

		Dim iCount As Integer
		Dim strColour As String
		Dim dTempDate As Date

		Dim blnIsBankHoliday As Boolean
		Dim blnIsWeekend As Boolean
		Dim blnHasEvent As Boolean
		Dim blnIsWorkingDay As Boolean
		Dim strHtmlRefresh As New StringBuilder

		Dim strCaption As String
		Dim strKeyName As String

		' Create function header strings
		strHtml = "<script type=""text/javascript"">" & vbNewLine

		strHtmlRefresh.Append("function refreshDateSpecifics() {" & vbNewLine &
														"refreshToggleValues();" & vbNewLine)

		' Create option strings
		For iCount = 0 To UBound(mavAbsences, 1)

			dTempDate = GetCalDay(iCount)
			Dim sDateObjectName = String.Format("$('#DateID_{0}')[0]", iCount.ToString.TrimStart)

			If (dTempDate <= mdLeavingDate Or mdLeavingDate = DateTime.FromOADate(0)) And dTempDate >= mdStartDate And (dTempDate <= mdCalendarEndDate And dTempDate >= mdCalendarStartDate) Then

				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsBankHoliday = mavAbsences(iCount).IsBankHoliday
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsWeekend = mavAbsences(iCount).IsWeekend
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strColour = mavAbsences(iCount).DisplayColor
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCaption = HttpUtility.HtmlEncode(mavAbsences(iCount).Caption)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnHasEvent = mavAbsences(iCount).ContainsData
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsWorkingDay = mavAbsences(iCount).IsWorkingDay

				If (Not blnIsWeekend) And (Not blnHasEvent) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("{0}.className = 'calendar_day';" & vbNewLine, sDateObjectName)
				End If

				If blnIsWeekend And (Not blnHasEvent) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtShowWeekends.value == 'highlighted') {{" & vbNewLine &
						"   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   }}" & vbNewLine &
						"else " & vbNewLine & "   {{" & vbNewLine &
						"   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day'; }}" & vbNewLine, sDateObjectName)
				End If


				'Has an event therefore deal with the Caption
				If blnHasEvent And (Not blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included') {{" & vbNewLine & _
						"  {0}.innerHTML = '';" & vbNewLine & _
						"  {0}.className = 'calendar_day';" & vbNewLine & _
						"  {0}.style.backgroundColor = ''; }}" & vbNewLine & _
						"else {{" & vbNewLine & _
						"   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine & _
						"     {0}.innerHTML = '{1}'; }}" & vbNewLine & _
						"   else {{" & vbNewLine & _
						"     {0}.innerHTML = ''; }}" & vbNewLine & _
						"  {0}.style.backgroundColor = '{2}'; }}" & vbNewLine, sDateObjectName, strCaption, strColour)

				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowWeekends.value == 'highlighted') {{" & vbNewLine &
						"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine &
						"else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowWeekends.value == 'unhighlighted') {{" & vbNewLine &
						"   {0}.className = 'calendar_day';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine &
						"else {{" & vbNewLine &
						"   {0}.style.backgroundColor = '{1}';" & vbNewLine &
						"   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
						"     {0}.innerHTML = '{2}'; }}" & vbNewLine &
						"   else {{" & vbNewLine &
						"     {0}.innerHTML = ''; }}" & vbNewLine &
						"   }}" & vbNewLine, sDateObjectName, strColour, strCaption)

				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (blnIsWeekend) And (blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtIncludeBankHolidays.value == 'included') {{" & vbNewLine &
						"   {0}.style.backgroundColor = '" & strColour & "';" & vbNewLine &
						"   if (frmChangeDetails.txtShowCaptions.value == 'show')  {{" & vbNewLine &
						"     {0}.innerHTML = '{1}'; }}" & vbNewLine &
						"   else {{" & vbNewLine &
						"     {0}.innerHTML = ''; }}" & vbNewLine & "   }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeBankHolidays.value == 'unincluded' && frmChangeDetails.txtShowWeekends.value == ""highlighted"") {{" & vbNewLine &
					"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeBankHolidays.value == 'unincluded' && frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") {{" & vbNewLine &
					"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeBankHolidays.value == 'unincluded' && frmChangeDetails.txtShowWeekends.value == ""unhighlighted"")  {{" & vbNewLine &
					"   {0}.className = 'calendar_day';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeBankHolidays.value == 'unincluded' && frmChangeDetails.txtShowBankHolidays.value == ""unhighlighted"") {{" & vbNewLine &
					"   {0}.className = 'calendar_day';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") {{" & vbNewLine &
					"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowWeekends.value == ""unhighlighted"") {{" & vbNewLine &
					"   {0}.className = 'calendar_day';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowWeekends.value == ""highlighted"") " & vbNewLine & "   {{" & vbNewLine &
					"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == 'included' && frmChangeDetails.txtShowBankHolidays.value == ""unhighlighted"")  {{" & vbNewLine &
					"   {0}.className = 'calendar_day';" & vbNewLine &
					"   {0}.style.backgroundColor = '';" & vbNewLine &
					"   {0}.innerHTML = ''; }}" & vbNewLine &
					"else {{" & vbNewLine &
					"   {0}.style.backgroundColor = '{2}';" & vbNewLine &
					"   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
					"     {0}.innerHTML = '{1}'; }}" & vbNewLine &
					"   else {{" & vbNewLine &
					"     {0}.innerHTML = ''; }}" & vbNewLine & "   }}" & vbNewLine, sDateObjectName, strCaption, strColour)
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (Not blnIsBankHoliday) And (blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("{0}.style.backgroundColor = '{1}';" & vbNewLine &
						"if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
						"  {0}.innerHTML = '{2}'; }}" & vbNewLine &
						"else {{" & vbNewLine &
						"  {0}.innerHTML = ''; }}" & vbNewLine, sDateObjectName, strColour, strCaption)
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtIncludeBankHolidays.value == 'included') {{" & vbNewLine &
						"   {0}.style.backgroundColor = '{1}';" & vbNewLine &
						"   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
						"    {0}.innerHTML = '{2}'; }}" & vbNewLine &
						"   else {{" & vbNewLine &
						"     {0}.innerHTML = ''; }}" & vbNewLine & "   }}" & vbNewLine &
						"else if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") {{" & vbNewLine &
						"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine &
						"else  {{" & vbNewLine &
						"   {0}.className = 'calendar_day';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine, sDateObjectName, strColour, strCaption)
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtIncludeBankHolidays.value == 'included') {{" & vbNewLine &
						"   {0}.style.backgroundColor = '{1}';" & vbNewLine &
						"   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
						"    {0}.innerHTML = '{2}'; }}" & vbNewLine &
						"   else {{" & vbNewLine &
						"     {0}.innerHTML = ''; }}" & vbNewLine & "   }}" & vbNewLine &
						"else if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") {{" & vbNewLine &
						"   {0}.className = 'calendar_nonworkingday';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine &
						"else {{" & vbNewLine &
						"   {0}.className = 'calendar_day';" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.innerHTML = ''; }}" & vbNewLine, sDateObjectName, strColour, strCaption)
				End If

				If (Not blnHasEvent) And blnIsBankHoliday And (Not blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.AppendFormat("if (frmChangeDetails.txtShowBankHolidays.value == 'highlighted') {{" & vbNewLine & _
						"  {0}.style.backgroundColor = '';" & vbNewLine &
						"  {0}.className = 'calendar_nonworkingday';" & vbNewLine & "   }}" & vbNewLine &
						"else {{" & vbNewLine & _
						"  {0}.style.backgroundColor = '';" & vbNewLine & _
						"  {0}.className = 'calendar_day'; }}" & vbNewLine, sDateObjectName)
				End If

				If (Not blnHasEvent) And blnIsBankHoliday And (blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh.Append(String.Format("if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") {{" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.className = 'calendar_nonworkingday'; }}" & vbNewLine &
						"else {{" & vbNewLine &
						"   {0}.style.backgroundColor = '';" & vbNewLine &
						"   {0}.className = 'calendar_day'; }}" & vbNewLine, sDateObjectName))
				End If
			End If
		Next iCount

		For iCount = 0 To miStrAbsenceTypes Step 1
			strKeyName = "KEY_" & LTrim(Str(iCount))
			strCaption = IIf(Trim(mastrAbsenceTypes(iCount, 3)) = vbNullString, "&nbsp", HttpUtility.HtmlEncode(mastrAbsenceTypes(iCount, 3)))
			strHtmlRefresh.AppendFormat("   if (frmChangeDetails.txtShowCaptions.value == 'show') {{" & vbNewLine &
				"     {0}.innerHTML = '{1}'; }}" & vbNewLine &
				"   else" & vbNewLine & "     {{" & vbNewLine &
				"     {0}.innerHTML = '&nbsp'; }}" & vbNewLine,
				strKeyName, strCaption)
		Next iCount

		strHtmlRefresh.Append(" }" & vbNewLine)

		' Concatenate functions into HTML string
		Return strHtml & strHtmlRefresh.ToString & vbNewLine & "</script>" & vbNewLine

	End Function

	Private Function FillCalBoxes(intStart As Integer, intEnd As Integer) As Boolean

		' This function actually fills the cal boxes between the indexes specified
		' according to the options selected by the user.

		Try

			Dim dtmCurrentDate As Date
			Dim strColour As String
			Dim objCurrentWP As WorkingPatternChange
			Dim objNextWP As WorkingPatternChange

			'Scroll forward in list to correct start working pattern for absence.
			dtmCurrentDate = GetCalDay(intStart)

			' Get correct start working pattern for absence
			objCurrentWP = mavWorkingPatternChanges.OrderBy(Function(m) m.ChangeDate).Where(Function(m) m.ChangeDate <= dtmCurrentDate).LastOrDefault
			objNextWP = mavWorkingPatternChanges.OrderBy(Function(m) m.ChangeDate).Where(Function(m) m.ChangeDate > objCurrentWP.ChangeDate).FirstOrDefault

			' Loop through the indexes as specified.
			For iCount = intStart To intEnd

				' Set current date variable
				dtmCurrentDate = GetCalDay(iCount)

				'Calculate the working pattern for this day
				If objNextWP IsNot Nothing Then
					If dtmCurrentDate >= objNextWP.ChangeDate Then
						objCurrentWP = objNextWP
						objNextWP = mavWorkingPatternChanges.OrderBy(Function(m) m.ChangeDate).Where(Function(m) m.ChangeDate > objCurrentWP.ChangeDate).FirstOrDefault
					End If
				End If

				' Mark this day as having an absence
				If Not mavAbsences(iCount).ContainsData Then
					strColour = GetColour(mstrAbsType)
					mavAbsences(iCount).ContainsData = True
					mavAbsences(iCount).Type = GetAbsenceCode(mstrAbsType)
					mavAbsences(iCount).Caption = mstrAbsCalendarCode
				Else
					strColour = GetColour("Multiple")
					mavAbsences(iCount).Type = GetAbsenceCode("Multiple")
					mavAbsences(iCount).Caption = "."
				End If

				' Store the details for this day
				mavAbsences(iCount).IsWorkingDay = AbsCal_DoTheyWorkOnThisDay(Weekday(dtmCurrentDate, FirstDayOfWeek.Sunday), IIf(iCount Mod 2 = 0, "AM", "PM"), objCurrentWP.WorkingPattern)
				mavAbsences(iCount).DisplayColor = strColour
				mavAbsences(iCount).Reason = Replace(mstrAbsReason, "'", "")
				mavAbsences(iCount).WorkingPattern = objCurrentWP.WorkingPattern
				mavAbsences(iCount).Duration = mdblAbsDuration
				mavAbsences(iCount).StartDate = mdAbsStartDate
				mavAbsences(iCount).StartSession = mstrAbsStartSession
				mavAbsences(iCount).EndDate = mdAbsEndDate
				mavAbsences(iCount).EndSession = mstrAbsEndSession

			Next iCount

			Return True

		Catch ex As Exception
			Throw

		End Try

	End Function

	Private Function GetColour(strType As String) As String

		' This function returns the colour for the specified absence type.
		' Derived from the key. If it cannot be found, then it defaults to
		' The colour for 'Other' which is Black

		Dim iCount As Integer
		Dim strColourString As String

		' Default
		strColourString = "black"

		For iCount = 0 To miStrAbsenceTypes	'UBound(mastrAbsenceTypes, 1) - 1

			If UCase(Trim(mastrAbsenceTypes(iCount, 0))) = UCase(Trim(strType)) Then
				strColourString = mastrAbsenceTypes(iCount, 1)
				Exit For
			End If

		Next iCount

		Return strColourString

	End Function

	Private Function GetAbsenceCode(strType As String) As String

		' This function returns the colour for the specified absence type.
		' Derived from the key. If it cannot be found, then it defaults to
		' The colour for 'Other' which is Black

		Dim iCount As Integer

		GetAbsenceCode = Trim(Str(miStrAbsenceTypes))	' Id for other (if nothing is found)
		For iCount = 0 To miStrAbsenceTypes	'UBound(mastrAbsenceTypes, 1) - 1

			'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If UCase(Trim(mastrAbsenceTypes(iCount, 0))) = UCase(Trim(strType)) Then
				GetAbsenceCode = Replace(mastrAbsenceTypes(iCount, 2), "'", "")
				Exit Function
			End If

		Next iCount

	End Function

	Public Function HTML_ForwardBackYear() As String

		Dim strHtml As String

		strHtml = "<TR>" & vbNewLine _
				& "   <TD valign=middle align=center colspan=2>" & vbNewLine _
				& "     <INPUT id=""cmdPreviousYear"" name=""cmdPreviousYear"" type=""button"" class=""btn"" value=""<<""" & vbNewLine _
				& "         onclick=""return cmdPreviousYear_onclick()""" & vbNewLine _
				& "         onmouseover = ""try{button_onMouseOver(this);}catch(e){}""" & vbNewLine _
				& "         onmouseout = ""try{button_onMouseOut(this);}catch(e){}""" & vbNewLine _
				& "         onfocus = ""try{button_onFocus(this);}catch(e){}""" & vbNewLine _
				& "         onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbNewLine

		' Different display if the calendar scrolls over a year
		If Year(mdCalendarStartDate) = Year(mdCalendarEndDate) Then
			strHtml = strHtml & LTrim(Str(Year(mdCalendarStartDate)))
		Else
			strHtml = strHtml & LTrim(Str(Year(mdCalendarStartDate))) & " - " & LTrim(Str(Year(mdCalendarEndDate)))
		End If

		strHtml = strHtml & "     <INPUT id=""cmdNextYear"" name=""cmdNextYear"" type=""button"" class=""btn"" value="">>""" & vbNewLine _
				& "         onclick=""return cmdNextYear_onclick()""" & vbNewLine _
				& "         onmouseover = ""try{button_onMouseOver(this);}catch(e){}""" & vbNewLine _
				& "         onmouseout = ""try{button_onMouseOut(this);}catch(e){}""" & vbNewLine _
				& "         onfocus = ""try{button_onFocus(this);}catch(e){}""" & vbNewLine _
				& "         onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbNewLine _
				& "  </TD>" & vbNewLine _
				& "</TR>" & vbNewLine

		Return strHtml

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

		'GetCalDay = mdCalendarStartDate + ((intIndex - 1) / 2)
		Return DateTime.FromOADate(mdCalendarStartDate.ToOADate + ((intIndex) / 2))

	End Function

	Public Function AbsCal_DoTheyWorkOnThisDay(intDay As Integer, strperiod As String, workingPattern As String) As Boolean

		Dim bFound As Boolean = True

		' Inputs  - 1 to 7 depending on the weekday 1 = sunday etc, "AM" or "PM"
		' Outputs - True/False
		Select Case strperiod.ToUpper
			Case "AM"
				If (Mid(workingPattern, (intDay * 2) - 1, 1) = " ") Or (Mid(workingPattern, (intDay * 2) - 1, 1) = "") Then
					bFound = False
				End If
			Case "PM"
				If (Mid(workingPattern, intDay * 2, 1) = " ") Or (Mid(workingPattern, intDay * 2, 1) = "") Then
					bFound = False
				End If
		End Select

		Return bFound

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

		' Determine the index depending on whether session is am or pm
		If Not booSession Then
			'am
			GetCalIndex = ((dtmDate.ToOADate - mdCalendarStartDate.ToOADate) * 2)	'intFirstDayIndex + (2 * diff)
		Else
			'pm
			GetCalIndex = ((dtmDate.ToOADate - mdCalendarStartDate.ToOADate) * 2) + 1	'(intFirstDayIndex + (2 * diff)) + 1
		End If

		' Only allow dates on this year to get processed
		If GetCalIndex < 0 Then
			GetCalIndex = 0
		End If

		If GetCalIndex > UBound(mavAbsences) Then
			GetCalIndex = UBound(mavAbsences) - If(booSession, 0, 1)
		End If

	End Function

	Public Function HTML_WorkingPattern(pstrWorkingPattern As String) As String

		Dim strHtml As String
		Dim iCount As Integer

		pstrWorkingPattern = pstrWorkingPattern & Space(14 - Len(pstrWorkingPattern))

		strHtml = "<table class='invisible' cellspacing=0 cellpadding=0 frame=0>" & vbNewLine

		' Row 1 contains day names
		strHtml = strHtml & "<tr align=middle>" & "<td>&nbsp;</td><td>" & UCase(Left(VB6.Format(1, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(2, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(3, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(4, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(5, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(6, "ddd"), 1)) & "</td>" & "<td>" & UCase(Left(VB6.Format(7, "ddd"), 1)) & "</td></tr>" & vbNewLine

		' Row two contains the AM fields
		strHtml = strHtml & "<tr><td>AM</td>"

		For iCount = 1 To 13 Step 2
			If Not Mid(pstrWorkingPattern, iCount, 1) = " " Then
				strHtml = strHtml & "<td><input id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" checked disabled></td>"
			Else
				strHtml = strHtml & "<td><input id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled></td>"
			End If
		Next iCount
		strHtml = strHtml & "</tr>"


		' Row three contains the PM fields
		strHtml = strHtml & "<tr><td>PM</td>"
		For iCount = 2 To 14 Step 2
			strHtml = strHtml & "<td><input id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"""
			If Not Mid(pstrWorkingPattern, iCount, 1) = " " Then
				strHtml = strHtml & " Checked"
			End If
			strHtml = strHtml & " disabled></td>"
		Next iCount
		strHtml = strHtml & "</tr></table>"

		Return strHtml
	End Function

	Public Function HTML_DisplayOptions() As String

		'Build the display options HTML
		Dim strHtml As String

		' Show include bank holidays option
		strHtml = "<tr><td colSpan=""2"">" & "<input id=""chkIncludeBankHolidays"" name=""chkIncludeBankHolidays"" type=""checkbox"" tabindex=-1 " & "onclick=""return refreshDateSpecifics()""" & "onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & "onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""" & IIf(mbDisplay_IncludeBankHolidays And (Not mblnDisableRegions), " CHECKED ", "") & IIf(mblnDisableRegions, " DISABLED='disabled' ", "") & ">" & "<label for=""chkIncludeBankHolidays"" Class=""checkbox" & IIf(mblnDisableRegions, " checkboxdisabled", "") & """ TabIndex = 0" & "    onkeypress = ""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & "    onmouseover = ""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & "    onmouseout = ""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & "    onfocus = ""try{checkboxLabel_onFocus(this);}catch(e){}""" & "    onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & "&nbsp;Include Bank Holidays" & "</label></td></tr>"

		' Show include working days only option
		strHtml = strHtml & "<tr><td colSpan=""2"">" & "<input id=""chkIncludeWorkingDaysOnly"" name=""chkIncludeWorkingDaysOnly"" type=""checkbox"" tabindex=-1 " & "onclick=""return refreshDateSpecifics()""" & "onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & "onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""" & IIf(mbDisplay_IncludeWorkingDaysOnly And (Not mblnDisableWPs), " CHECKED ", "") & IIf(mblnDisableWPs, " DISABLED='disabled' ", "") & ">" & "<label for=""chkIncludeWorkingDaysOnly"" Class=""checkbox" & IIf(mblnDisableWPs, " checkboxdisabled", "") & """ TabIndex = 0" & "    onkeypress = ""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & "    onmouseover = ""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & "    onmouseout = ""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & "    onfocus = ""try{checkboxLabel_onFocus(this);}catch(e){}""" & "    onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & "&nbsp;Working Days Only" & "</label></td></tr>"

		' Show show bank holidays option
		strHtml = strHtml & "<tr><td colSpan=""2"">" & "<input id=""chkShowBankHolidays"" name=""chkShowBankHolidays"" type=""checkbox"" tabindex=-1 " & "onclick=""return refreshDateSpecifics()""" & "onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & "onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""" & IIf(mbDisplay_ShowBankHolidays And (Not mblnDisableRegions), " CHECKED ", "") & IIf(mblnDisableRegions, " DISABLED='disabled' ", "") & ">" & "<label for=""chkShowBankHolidays"" Class=""checkbox" & IIf(mblnDisableRegions, " checkboxdisabled", "") & """ TabIndex = 0" & "    onkeypress = ""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & "    onmouseover = ""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & "    onmouseout = ""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & "    onfocus = ""try{checkboxLabel_onFocus(this);}catch(e){}""" & "    onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & "&nbsp;Show Bank Holidays" & "</label></td></tr>"

		' Show show calendar captions option
		strHtml = strHtml & "<tr><td colSpan=""2"">" & "<input id=""chkShowCaptions"" name=""chkShowCaptions"" type=""checkbox"" tabindex=-1 " & "onclick=""return refreshDateSpecifics()""" & "onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & "onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""" & IIf(mbDisplay_ShowCaptions, " CHECKED ", "") & ">" & "<label for=""chkShowCaptions"" Class=""checkbox"" TabIndex = 0" & "    onkeypress = ""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & "    onmouseover = ""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & "    onmouseout = ""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & "    onfocus = ""try{checkboxLabel_onFocus(this);}catch(e){}""" & "    onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & "&nbsp;Show Calendar Captions" & "</label></td></tr>"

		' Show show weekends option
		strHtml = strHtml & "<tr><td colSpan=""2"">" & "<input id=""chkShowWeekends"" name=""chkShowWeekends"" type=""checkbox"" tabindex=-1 " & "onclick=""return refreshDateSpecifics()""" & "onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & "onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""" & IIf(mbDisplay_ShowWeekends, " CHECKED ", "") & ">" & "<label for=""chkShowWeekends"" Class=""checkbox"" TabIndex = 0" & "    onkeypress = ""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & "    onmouseover = ""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & "    onmouseout = ""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & "    onfocus = ""try{checkboxLabel_onFocus(this);}catch(e){}""" & "    onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & "&nbsp;Show Weekends" & "</label></td></tr>"

		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_DisplayOptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Return strHtml

	End Function

	Private Sub GetWorkingPatterns()

		Dim rstHistoricWPatterns As DataTable
		Dim sSQL As String
		Dim strAbsWPattern As String

		Try

			' Define a blank working pattern array
			mavWorkingPatternChanges.Add(
				New WorkingPatternChange With {
					.ChangeDate = DateTime.MinValue,
					.WorkingPattern = Space(14)})

			If Not mblnDisableWPs Then
				' If we are using historic WPattern, ensure we use the right WPattern for each day of absence
				If PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptHistoricWPattern Then

					sSQL = String.Format("SELECT  [{1}] AS [Date], [{2}] AS [WP] FROM {0} WHERE ID_{3} = {4}" &
											"ORDER BY [{1}] ASC",
											PersonnelModule.gsPersonnelHWorkingPatternTableRealSource,
											PersonnelModule.gsPersonnelHWorkingPatternDateColumnName,
											PersonnelModule.gsPersonnelHWorkingPatternColumnName,
											PersonnelModule.glngPersonnelTableID,
											mlngPersonnelRecordID)

					' Get the working patterns for the absence period
					rstHistoricWPatterns = DB.GetDataTable(sSQL)

					For Each objRow As DataRow In rstHistoricWPatterns.Rows

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						strAbsWPattern = IIf(IsDBNull(objRow("WP").ToString()), Space(14), objRow("WP").ToString())
						strAbsWPattern &= Space(14 - Len(strAbsWPattern))

						mavWorkingPatternChanges.Add(
							New WorkingPatternChange With {
								.ChangeDate = IIf(IsDBNull(objRow("Date")), DateTime.MinValue, CDate(objRow("Date"))),
								.WorkingPattern = strAbsWPattern})

					Next

				Else

					' Its a static working pattern, get it from personnel
					sSQL = "SELECT " & mstrSQLSelect_PersonnelStaticWP & "  AS 'WP'  " & vbNewLine
					sSQL = sSQL & "FROM " & PersonnelModule.gsPersonnelTableName & vbNewLine
					For iCount = 0 To UBound(mvarTableViews, 2) Step 1
						'<Personnel CODE>
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarTableViews(0, iCount) = PersonnelModule.glngPersonnelTableID Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, iCount) & vbNewLine
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "     ON  " & PersonnelModule.gsPersonnelTableName & ".ID = " & mvarTableViews(3, iCount) & ".ID" & vbNewLine
						End If
					Next iCount
					sSQL = sSQL & "WHERE " & PersonnelModule.gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID

					rstHistoricWPatterns = DB.GetDataTable(sSQL)

					' Stuff the working pattern into array
					If rstHistoricWPatterns.Rows.Count > 0 Then
						mavWorkingPatternChanges(0).ChangeDate = DateTime.MinValue
						mavWorkingPatternChanges(0).WorkingPattern = Left(IIf(IsDBNull(rstHistoricWPatterns.Rows(0)("WP")), Space(14), rstHistoricWPatterns.Rows(0)("WP")) & Space(14), 14)
					End If

				End If
			Else

				mavWorkingPatternChanges(0).ChangeDate = DateTime.MinValue
				mavWorkingPatternChanges(0).WorkingPattern = FULL_WP

			End If

		Catch ex As Exception
			Throw

		End Try


	End Sub

	Private Sub GenerateRegionData()

		Dim intTemp As Integer
		Dim bNewRegionFound As Boolean
		Dim strRegionAtCurrentDate As String
		Dim dtmNextChangeDate As Date
		Dim intCount As Integer
		Dim rstBankHolRegion As DataTable
		Dim dtmCurrentDate As Date
		Dim sSQL As String
		Dim lngCount As Integer

		Dim dtDummyDate As Date = DateTime.MaxValue

		bNewRegionFound = False

		If Not mblnDisableRegions Then
			' If we are using historic region, find the region change dates
			If PersonnelModule.grtRegionType = RegionType.rtHistoricRegion Then

				' Get the first region for this employee within this calendar year
				rstBankHolRegion = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " AS 'Date', " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " <= '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & PersonnelModule.gsPersonnelHRegionDateColumnName & " DESC")

				' Was there a region at the start of the calendar
				If rstBankHolRegion.Rows.Count = 0 Then
					strRegionAtCurrentDate = ""
				Else
					strRegionAtCurrentDate = rstBankHolRegion.Rows(0)("Region").ToString()
					bNewRegionFound = True
				End If

				' Get the second region for this employee within this calendar year
				rstBankHolRegion = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " AS 'Date', " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " > '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & PersonnelModule.gsPersonnelHRegionDateColumnName & " ASC")

				' Was there a region at the start of the calendar
				If rstBankHolRegion.Rows.Count = 0 Then
					dtmNextChangeDate = dtDummyDate
					' dtmNextChangeDate = dtDummyDate
				Else
					dtmNextChangeDate = CDate(rstBankHolRegion.Rows(0)("Date"))
				End If


				For intCount = LBound(mavAbsences, 1) To UBound(mavAbsences, 1) Step 2

					' Get the date of the current index
					dtmCurrentDate = GetCalDay(intCount)

					' Only refer to the region table if the current date is a region change date
					If (dtmCurrentDate >= dtmNextChangeDate) And (dtmCurrentDate <> dtDummyDate) Then


						'JDM - 11/09/01 - Fault 2820 - Bank hols not showing for year starting with working pattern.
						' Find the employees region for this date
						rstBankHolRegion = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " AS 'Date', " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " >= '" & VB6.Format(dtmNextChangeDate, "MM/dd/yyyy") & "' " & "ORDER BY " & PersonnelModule.gsPersonnelHRegionDateColumnName & " ASC")

						If rstBankHolRegion.Rows.Count = 0 Then

							' No regions found for this user
							dtmNextChangeDate = dtDummyDate

						Else

							strRegionAtCurrentDate = rstBankHolRegion.Rows(0)("Region").ToString()
							bNewRegionFound = True

							' Now get the next change date
							rstBankHolRegion = DB.GetDataTable("SELECT TOP 1 " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " AS 'Date', " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & "ID_" & PersonnelModule.glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " > '" & VB6.Format(rstBankHolRegion.Rows(0)("Date"), "MM/dd/yyyy") & "' " & "ORDER BY " & PersonnelModule.gsPersonnelHRegionDateColumnName & " ASC")
							If rstBankHolRegion.Rows.Count = 0 Then
								dtmNextChangeDate = dtDummyDate
							Else
								dtmNextChangeDate = CDate(rstBankHolRegion.Rows(0)("Date"))
							End If

						End If

					End If

					' Define the region for this period
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavAbsences(intCount).Region = Replace(strRegionAtCurrentDate, "'", "''")
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount + 1, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavAbsences(intCount + 1).Region = Replace(strRegionAtCurrentDate, "'", "''")

					' If current region has changed
					If bNewRegionFound Then

						If BankHolidayModule.gfBankHolidaysEnabled Then

							' Get bank holidays for this region
							' DONE
							sSQL = vbNullString
							sSQL = sSQL & "SELECT " & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " AS 'Date' " & vbNewLine
							sSQL = sSQL & "FROM " & BankHolidayModule.gsBHolTableRealSource & " " & vbNewLine

							sSQL = sSQL & "WHERE " & BankHolidayModule.gsBHolTableRealSource & ".ID_" & BankHolidayModule.glngBHolRegionTableID & " = " & vbNewLine
							sSQL = sSQL & "        (SELECT " & BankHolidayModule.gsBHolRegionTableName & ".ID " & vbNewLine
							sSQL = sSQL & "         FROM " & BankHolidayModule.gsBHolRegionTableName & vbNewLine
							For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
								'<REGIONAL CODE>
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarTableViews(0, lngCount) = BankHolidayModule.glngBHolRegionTableID Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sSQL = sSQL & "           ON  " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
								End If
							Next lngCount
							sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbNewLine

							sSQL = sSQL & " AND " & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= '" & Replace(VB6.Format(dtmCurrentDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & vbNewLine
							sSQL = sSQL & " AND " & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " <= '" & Replace(VB6.Format(DateTime.FromOADate(dtmNextChangeDate.ToOADate - 1), "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & vbNewLine
							sSQL = sSQL & "ORDER BY " & BankHolidayModule.gsBHolDateColumnName & " ASC"
							rstBankHolRegion = DB.GetDataTable(sSQL)

							' Cycle through the recordset checking for the current day

							For Each objRow As DataRow In rstBankHolRegion.Rows

								intTemp = GetCalIndex(CDate(objRow("Date")), False)

								'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mavAbsences(intTemp).IsBankHoliday = True
								'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp + 1, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mavAbsences(intTemp + 1).IsBankHoliday = True

							Next

						End If

						' Flag this region has had it's bank holidays drawn
						bNewRegionFound = False

					End If

				Next intCount

			Else

				If BankHolidayModule.gfBankHolidaysEnabled Then

					' We are using a static region so just use the employees current region
					strRegionAtCurrentDate = mstrRegion
					' DONE
					sSQL = vbNullString
					sSQL = sSQL & "SELECT " & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " AS 'Date' " & vbNewLine
					sSQL = sSQL & "FROM " & BankHolidayModule.gsBHolTableRealSource & " " & vbNewLine
					sSQL = sSQL & "WHERE " & BankHolidayModule.gsBHolTableRealSource & ".ID_" & BankHolidayModule.glngBHolRegionTableID & " = " & vbNewLine
					sSQL = sSQL & "        (SELECT " & BankHolidayModule.gsBHolRegionTableName & ".ID " & vbNewLine
					sSQL = sSQL & "         FROM " & BankHolidayModule.gsBHolRegionTableName & vbNewLine
					For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
						'<REGIONAL CODE>
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarTableViews(0, lngCount) = BankHolidayModule.glngBHolRegionTableID Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "           ON  " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
						End If
					Next lngCount
					sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbNewLine
					sSQL = sSQL & "ORDER BY " & BankHolidayModule.gsBHolDateColumnName & " ASC" & vbNewLine

					rstBankHolRegion = DB.GetDataTable(sSQL)

					' Cycle through the recordset checking for the current day
					For Each objRow As DataRow In rstBankHolRegion.Rows

						intTemp = GetCalIndex(CDate(objRow("Date")), False)

						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intTemp).IsBankHoliday = True
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp + 1, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intTemp + 1).IsBankHoliday = True

					Next

					' Define the region for this period
					For intCount = LBound(mavAbsences, 1) To UBound(mavAbsences, 1) Step 2

						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intCount).Region = Replace(strRegionAtCurrentDate, "'", "''")
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount + 1, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intCount + 1).Region = Replace(strRegionAtCurrentDate, "'", "''")

					Next intCount

				End If

			End If
		End If 'Not mblnDisableRegions

	End Sub

	Private Sub CheckPermission_RegionInfo()

		Dim strTableColumn As String

		'Check the  Bank Holiday Region Table - Region Table
		'           Bank Holiday Region Table - Region Column
		'           Bank Holidays Table - Bank Holiday Table
		'           Bank Holidays Table - Date Column
		'           Bank Holidays Table - Descripiton Column
		'...Bank Holiday module setup information.
		'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
		If BankHolidayModule.gsBHolRegionTableName = "" Or BankHolidayModule.gsBHolRegionColumnName = "" Or BankHolidayModule.gsBHolTableName = "" Or BankHolidayModule.gsBHolDateColumnName = "" Or BankHolidayModule.gsBHolDescriptionColumnName = "" Then

			GoTo DisableRegions
		End If

		'Check the  Career Change Region - Static Region Column
		'           Career Change Region - Historic Region Table
		'           Career Change Region - Historic Region Column
		'           Career Change Region - Historic Region Effective Date Column
		'...Personnel - Career Change module setup information.
		'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
		If PersonnelModule.gsPersonnelRegionColumnName = "" Then
			If PersonnelModule.gsPersonnelHRegionTableName = "" Or PersonnelModule.gsPersonnelHRegionColumnName = "" Or PersonnelModule.gsPersonnelHRegionDateColumnName = "" Then

				GoTo DisableRegions
			End If
		End If




		'*******************************************************************
		' All Region module information is setup                           *
		' Now check the permissions on the Region module setup information *
		'*******************************************************************
		'Bank Holiday Region Table - Region Table (Regional Information)
		'Bank Holiday Region Table - Region Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolRegionTableID, BankHolidayModule.gsBHolRegionTableName, BankHolidayModule.gsBHolRegionColumnName, strTableColumn) Then
			mstrSQLSelect_RegInfoRegion = strTableColumn
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Bank Holidays Table - Bank Holiday Table (Region History)
		'Bank Holidays Table - Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolTableID, BankHolidayModule.gsBHolTableName, BankHolidayModule.gsBHolDateColumnName, strTableColumn) Then
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Bank Holidays Table - Bank Holiday Table (Region History)
		'Bank Holidays Table - Descripiton Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolTableID, BankHolidayModule.gsBHolTableName, BankHolidayModule.gsBHolDescriptionColumnName, strTableColumn) Then
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




		'*******************************************************************
		' Permission granted on all Region module information.             *
		' Now check the permissions on the                                 *
		' Personnel Career Change Region module setup information          *
		'*******************************************************************
		'Check Career Change Region access
		If PersonnelModule.gsPersonnelRegionColumnName <> "" Then
			'Personnel Table
			'Career Change Region - Static Region Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelRegionColumnName, strTableColumn) Then
				mstrSQLSelect_PersonnelStaticRegion = strTableColumn
				strTableColumn = vbNullString
			Else
				GoTo DisableRegions
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		Else
			'Career Change Region - Historic Region Table
			'Career Change Region - Historic Region Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(PersonnelModule.glngPersonnelHRegionTableID, PersonnelModule.gsPersonnelHRegionTableName, PersonnelModule.gsPersonnelHRegionColumnName, strTableColumn) Then
				strTableColumn = vbNullString
			Else
				GoTo DisableRegions
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

			'Career Change Region - Historic Region Table
			'Career Change Region - Historic Region Effective Date Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(PersonnelModule.glngPersonnelHRegionTableID, PersonnelModule.gsPersonnelHRegionTableName, PersonnelModule.gsPersonnelHRegionDateColumnName, strTableColumn) Then
				strTableColumn = vbNullString
			Else
				GoTo DisableRegions
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		End If

TidyUpAndExit:
		Exit Sub

DisableRegions:
		mblnDisableRegions = True
		ShowBankHolidays = False
		IncludeBankHolidays = False
		GoTo TidyUpAndExit

	End Sub

	Private Sub CheckPermission_WPInfo()

		Dim objColumn As CColumnPrivileges
		Dim pblnColumnOK As Boolean
		Dim strTableColumn As String

		'Check the  Career Change Working Pattern - Static Working Pattern Column
		'           Career Change Working Pattern - Historic Working Pattern Table
		'           Career Change Working Pattern - Historic Working Pattern Column
		'           Career Change Working Pattern - Historic Working Pattern Effective Date Column
		'...Personnel - Career Change module setup information.
		'If any are blank then we need to allow the report to run, but disable the Working Dys Display Option.
		If PersonnelModule.gsPersonnelWorkingPatternColumnName = "" Then
			If PersonnelModule.gsPersonnelHWorkingPatternTableName = "" Or PersonnelModule.gsPersonnelHWorkingPatternColumnName = "" Or PersonnelModule.gsPersonnelHWorkingPatternDateColumnName = "" Then

				GoTo DisableWPs
			End If
		End If

		'****************************************************************************
		' All Working Pattern module information is setup                           *
		' Now check the permissions on the Working Pattern module setup information *
		'****************************************************************************
		'Check Career Change Working Pattern access
		If PersonnelModule.gsPersonnelWorkingPatternColumnName <> "" Then
			'Career Change Working Pattern - Static Working Pattern Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelWorkingPatternColumnName, strTableColumn) Then
				mstrSQLSelect_PersonnelStaticWP = strTableColumn
				strTableColumn = vbNullString
			Else
				GoTo DisableWPs
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		Else
			'Career Change Working Pattern - Historic Working Pattern Table
			objColumn = GetColumnPrivileges(PersonnelModule.gsPersonnelHWorkingPatternTableName)

			'Career Change Working Pattern - Historic Working Pattern Column
			pblnColumnOK = objColumn.IsValid(PersonnelModule.gsPersonnelHWorkingPatternColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(PersonnelModule.gsPersonnelHWorkingPatternColumnName).AllowSelect
			End If
			If pblnColumnOK = False Then
				GoTo DisableWPs
			End If

			'Career Change Working Pattern - Historic Working Pattern Effective Date Column
			pblnColumnOK = objColumn.IsValid(PersonnelModule.gsPersonnelHWorkingPatternDateColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(PersonnelModule.gsPersonnelHWorkingPatternDateColumnName).AllowSelect
			End If
			If pblnColumnOK = False Then
				GoTo DisableWPs
			End If

		End If

TidyUpAndExit:
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objColumn = Nothing
		Exit Sub

DisableWPs:
		mblnDisableWPs = True
		IncludeWorkingDaysOnly = False
		GoTo TidyUpAndExit

	End Sub

	Private Function CheckPermission_AbsCalSpecifics() As Boolean

		Dim strTableColumn As String
		Dim strModulePermErrorMSG As String

		strModulePermErrorMSG = vbNullString

		'Check Module Setup on each of the module columns.
		'
		'                       II
		'                       II
		'                       II
		'                       II
		'                    \  II  /
		'                     \ II /
		'                      \II/
		'                       \/

		'Check the Absence Table
		'          Absence Table - Start Date Column
		'          Absence Table - Start Session Column
		'          Absence Table - End Date Column
		'          Absence Table - End Session Column
		'          Absence Table - Absence Type Column
		'          Absence Table - Absence Reason Column
		'          Absence Table - Absence Duration Column
		'...Absence module setup information.
		'If any are blank then we need to fail the Absence Calendar report.
		If AbsenceModule.gsAbsenceTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Table' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceStartDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Date Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceStartSessionColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Session Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceEndDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'End Date Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceEndSessionColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'End Session Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceTypeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceReasonColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Reason Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceDurationColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Duration Column' in the Absence module setup must be defined." & vbNewLine
		End If


		'Check the Absence Type Table
		'          Absence Type Table - Absence Type Column
		'          Absence Type Table - Absence Code Column
		'          Absence Type Table - Calendar Code Column
		'...Absence module setup information.
		'If any are blank then we need to fail the Absence Calendar report.
		If AbsenceModule.gsAbsenceTypeTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Table' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceTypeTypeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceTypeCodeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Code Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If AbsenceModule.gsAbsenceTypeCalCodeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Calendar Code Column' in the Absence module setup must be defined." & vbNewLine
		End If


		'Check the Personnel Table
		'          Personnel Table - Start Date Column
		'          Personnel Table - Leaving Date Column
		'...Personnel module setup information.
		'If any are blank then we need to fail the Absence Calendar report.
		If PersonnelModule.gsPersonnelTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Personnel Table' in the Personnel module setup must be defined." & vbNewLine
		End If
		If PersonnelModule.gsPersonnelStartDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Date Column' in the Personnel module setup must be defined." & vbNewLine
		End If
		If PersonnelModule.gsPersonnelLeavingDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Leaving Date Column' in the Personnel module setup must be defined." & vbNewLine
		End If

		If Len(strModulePermErrorMSG) > 0 Then
			strModulePermErrorMSG = strModulePermErrorMSG & vbNewLine
		End If

		If Len(strModulePermErrorMSG) > 0 Then
			strModulePermErrorMSG = "The Absence Calendar failed for the following reasons: " & vbNewLine & vbNewLine & strModulePermErrorMSG
			mstrErrorMSG = strModulePermErrorMSG
			'MsgBox strModulePermErrorMSG, vbOKOnly + vbExclamation, "Absence Calendar"
			GoTo FailReport
		End If

		'Check Permissions on each of these columns and set the select string for each.
		'
		'                                     II
		'                                     II
		'                                     II
		'                                     II
		'                                  \  II  /
		'                                   \ II /
		'                                    \II/
		'                                     \/

		'Absence Specifics
		'Absence Table - Start Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceStartDateColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceStartDate = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - Start Date Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - Start Session Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceStartSessionColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceStartSession = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - Start Session Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - End Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceEndDateColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceEndDate = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - End Date Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - End Session Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceEndSessionColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceEndSession = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - End Session Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - Absence Type Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceTypeColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceType = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - Absence Type Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - Absence Reason Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceReasonColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceReason = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - Absence Reason Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Table - Absence Duration Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTableID, AbsenceModule.gsAbsenceTableName, AbsenceModule.gsAbsenceDurationColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceDuration = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Table - Absence Duration Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Absence Type Specifics
		'Absence Type Table - Absence Type Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTypeTableID, AbsenceModule.gsAbsenceTypeTableName, AbsenceModule.gsAbsenceTypeTypeColumnName, strTableColumn) Then
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Type Table - Absence Type Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Type Table - Absence Code Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTypeTableID, AbsenceModule.gsAbsenceTypeTableName, AbsenceModule.gsAbsenceTypeCodeColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceTypeCode = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Type Table - Absence Code Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Type Table - Calendar Code Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(AbsenceModule.glngAbsenceTypeTableID, AbsenceModule.gsAbsenceTypeTableName, AbsenceModule.gsAbsenceTypeCalCodeColumnName, strTableColumn) Then
			mstrSQLSelect_AbsenceTypeCalCode = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Type Table - Calendar Code Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Personnel Specifics
		'Personnel Table - Start Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelStartDateColumnName, strTableColumn) Then
			mstrSQLSelect_PersonnelStartDate = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Personnel Table - Start Date Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Personnel Table - Leaving Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelLeavingDateColumnName, strTableColumn) Then
			mstrSQLSelect_PersonnelLeavingDate = strTableColumn
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Personnel Table - Leaving Date Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		If Len(strModulePermErrorMSG) > 0 Then
			strModulePermErrorMSG = "The Absence Calendar failed for the following reasons: " & vbNewLine & vbNewLine & strModulePermErrorMSG
			mstrErrorMSG = strModulePermErrorMSG
			'MsgBox strModulePermErrorMSG, vbOKOnly + vbExclamation, "Absence Calendar"
			GoTo FailReport
		End If

		CheckPermission_AbsCalSpecifics = True

TidyUpAndExit:
		Exit Function

FailReport:
		mblnFailReport = True
		CheckPermission_AbsCalSpecifics = False
		GoTo TidyUpAndExit

	End Function

	Private Function CheckPermission_Columns(plngTableID As Integer, pstrTableName As String, pstrColumnName As String, ByRef strSQLRef As String) As Boolean

		'This function checks if the current user has read(select) permissions
		'on this column. If the user only has access through views then the
		'relevent views are added to the mvarTableViews() array which in turn
		'are used to create the join part of the query.

		Dim lngTempTableID As Integer
		Dim strTempTableName As String
		Dim strTempColumnName As String
		Dim blnColumnOK As Boolean
		Dim blnFound As Boolean
		Dim blnNoSelect As Boolean
		Dim strSource As String
		Dim intNextIndex As Integer
		Dim blnOK As Boolean
		Dim strTable As String
		Dim strColumn As String

		Dim pintNextIndex As Integer

		' Set flags with their starting values
		blnOK = True
		blnNoSelect = False

		' Load the temp variables
		lngTempTableID = plngTableID
		strTempTableName = pstrTableName
		strTempColumnName = pstrColumnName

		' Check permission on that column
		mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
		mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

		blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

		If blnColumnOK Then
			blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
		End If

		If blnColumnOK Then
			' this column can be read direct from the tbl/view or from a parent table
			strTable = mstrRealSource
			strColumn = strTempColumnName

			If (plngTableID = AbsenceModule.glngAbsenceTableID) And (mstrAbsenceTableRealSource = vbNullString) Then
				mstrAbsenceTableRealSource = strTable
			End If

			'    ' If the table isnt the base table (or its realsource) then
			'    ' Check if it has already been added to the array. If not, add it.
			'    If lngTempTableID <> mlngCalendarReportsBaseTable Then
			blnFound = False
			For intNextIndex = 1 To UBound(mvarTableViews, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarTableViews(1, intNextIndex) = 0 And mvarTableViews(2, intNextIndex) = lngTempTableID Then
					blnFound = True
					Exit For
				End If
			Next intNextIndex

			If Not blnFound Then
				intNextIndex = UBound(mvarTableViews, 2) + 1
				ReDim Preserve mvarTableViews(3, intNextIndex)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarTableViews(1, intNextIndex) = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarTableViews(2, intNextIndex) = lngTempTableID
			End If
			'    End If

			strSQLRef = strTable & "." & strColumn
		Else

			' this column cannot be read direct. If its from a parent, try parent views
			' Loop thru the views on the table, seeing if any have read permis for the column

			Dim mstrViews(0) As String
			For Each mobjTableView1 As TablePrivilege In gcoTablePrivileges.Collection
				If (Not mobjTableView1.IsTable) And (mobjTableView1.TableID = lngTempTableID) And (mobjTableView1.AllowSelect) Then

					strSource = mobjTableView1.ViewName
					mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

					' Get the column permission for the view
					mobjColumnPrivileges = GetColumnPrivileges(strSource)

					' If we can see the column from this view
					If mobjColumnPrivileges.IsValid(strTempColumnName) Then
						If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

							ReDim Preserve mstrViews(UBound(mstrViews) + 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mstrViews(UBound(mstrViews)) = mobjTableView1.ViewName

							' Check if view has already been added to the array
							blnFound = False
							For intNextIndex = 0 To UBound(mvarTableViews, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarTableViews(1, intNextIndex) = 1 And mvarTableViews(2, intNextIndex) = mobjTableView1.ViewID Then
									blnFound = True
									Exit For
								End If
							Next intNextIndex

							If Not blnFound Then
								' View hasnt yet been added, so add it !
								intNextIndex = UBound(mvarTableViews, 2) + 1
								ReDim Preserve mvarTableViews(3, intNextIndex)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(0, intNextIndex) = mobjTableView1.TableID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(1, intNextIndex) = 1
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(2, intNextIndex) = mobjTableView1.ViewID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(3, intNextIndex) = mobjTableView1.ViewName
							End If

						End If
					End If
				End If

			Next
			'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjTableView = Nothing

			' Does the user have select permission thru ANY views ?
			If UBound(mstrViews) = 0 Then
				blnNoSelect = True
			Else
				strSQLRef = ""
				For pintNextIndex = 1 To UBound(mstrViews)
					If pintNextIndex = 1 Then
						strSQLRef = "CASE"
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(pintNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQLRef = strSQLRef & " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & strTempColumnName
				Next pintNextIndex

				If Len(strSQLRef) > 0 Then
					strSQLRef = strSQLRef & " ELSE NULL" & " END "
				End If

			End If

			' If we cant see a column, then get outta here
			If blnNoSelect Then
				strSQLRef = vbNullString
				CheckPermission_Columns = False
				Exit Function
			End If

			If Not blnOK Then
				strSQLRef = vbNullString
				CheckPermission_Columns = False
				Exit Function
			End If

		End If

		Return True

	End Function

End Class