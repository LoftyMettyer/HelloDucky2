Option Strict Off
Option Explicit On

Imports System.Globalization
Imports ADODB
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports VB = Microsoft.VisualBasic
Public Class AbsenceCalendar

	Private Const CELLSIZE As Integer = 17
	Private Const FULL_WP As String = "SSMMTTWWTTFFSS"

	Private mstrClientDateFormat As String

	Private mstrRealSource As String
	Private mlngPersonnelRecordID As Long
	Private mrstAbsenceRecords As ADODB.Recordset
	Private mbColourKeyLoaded As Boolean

	Private mdStartDate As Date
	Private mdLeavingDate As Date
	Private mstrRegion As String
	Private mstrWorkingPattern As String

	Private miWorkingPatternArray As Integer ' Location in the Working Pattern array
	Private mdtmWorkingPatternDate As Date ' Next change date of the working pattern

	Private mstrHexColour_OptionBoxes As String

	Private mstrBlankSpace As String

	Private mdAbsStartDate As Date
	Private mstrAbsStartSession As String
	Private mdAbsEndDate As Date
	Private mstrAbsEndSession As String
	Private mstrAbsType As String
	Private mstrAbsCalendarCode As String
	Private mstrAbsWPattern As String
	Private mlngAbsDuration As Double
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

	Dim mavAbsences(,) As Object ' Stores each of the absence cells (redefined later as 733,6 so as to auto clear it)
	'0 = Contains data (true / false)
	'1 = Weekend (true / false)
	'2 = Caption
	'3 = Is a bank holiday (true / false)
	'4 = Is a working day (true / false)
	'5 = Display Colour
	'6 = Absence Type(s) for this day
	'7 = Reason
	'8 = Working Pattern
	'9 = Duration
	'10 = Start date of absence
	'11 = Start session of absence
	'12 = End date of absence
	'13 = End session of absence
	'14 = Region

	Dim mavWorkingPatternChanges(,) As Object	' Stores the working pattern changes
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
				mdCalendarStartDate = DateSerial(Year(mdCalendarStartDate), giAbsenceCalStartMonth, 1)
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

	Public WriteOnly Property Username() As String
		Set(ByVal value As String)
			gsUsername = value
		End Set
	End Property

	Public WriteOnly Property Connection() As ADODB.Connection
		Set(ByVal value As ADODB.Connection)
			gADOCon = value
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
			mbDisplay_ShowWeekends = IIf(Value = "highlighted", True, IIf(IsNothing(Value), gfAbsenceCalWeekendShading, False))
		End Set
	End Property

	Public WriteOnly Property ShowCaptions() As String
		Set(ByVal Value As String)
			' Are the captions to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_ShowCaptions = IIf(Value = "show", True, IIf(IsNothing(Value), gfAbsenceCalShowCaptions, False))
		End Set
	End Property

	Public WriteOnly Property ShowBankHolidays() As String
		Set(ByVal Value As String)
			' Are the bank holidays to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_ShowBankHolidays = IIf(Value = "highlighted", True, IIf(IsNothing(Value), gfAbsenceCalBHolShading, False))
		End Set
	End Property

	Public WriteOnly Property IncludeBankHolidays() As String
		Set(ByVal Value As String)
			' Are the bank holidays to be included (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_IncludeBankHolidays = IIf(Value = "included", True, IIf(IsNothing(Value), gfAbsenceCalBHolInclude, False))
		End Set
	End Property

	Public WriteOnly Property IncludeWorkingDaysOnly() As String
		Set(ByVal Value As String)
			' Are the working days only to be shown (if parameter is empty read the default DB value)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbDisplay_IncludeWorkingDaysOnly = IIf(Value = "included", True, IIf(IsNothing(Value), gfAbsenceCalIncludeWorkingDaysOnly, False))
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
	Private Function GetHexColour(ByRef iRed As Integer, ByRef iGreen As Integer, ByRef iBlue As Integer) As String
		Return "#" & Right("0" & Hex(iRed), 2) & Right("0" & Hex(iGreen), 2) & Right("0" & Hex(iBlue), 2)
	End Function

	' Load the defaults
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		ReDim mvarTableViews(3, 0)
		ReDim mavWorkingPatternChanges(1, 0)

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
		piStartMonth = If(IsNumeric(piStartMonth), piStartMonth, giAbsenceCalStartMonth)

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

	Private Function HTML_Month(ByRef piMonthNumber As Integer, ByRef piYear As Integer) As String

		Dim iCount As Integer
		Dim strHtml As String
		Dim strHtmlDays As String
		Dim strHtmlDaysStart As String
		Dim iStartNumber As Integer
		Dim iEndNumber As Integer
		Dim dTempDate As Date
		Dim iIndexAM As String
		Dim iIndexPM As String
		Dim strHtmlCellString As String

		strHtml = "<SPAN id=Month" & LTrim(Str(piMonthNumber)) & ">" & vbNewLine & "<TR>" _
			& "<TD class='smallfont'>&nbsp;" & MonthName(piMonthNumber) & "&nbsp;</TD>" _
			& "<TD> <TABLE class='invisible' cellPadding=0 cellSpacing=0 width=""100%"" height=""100%"">"

		' Calculate month parameters

		' JDM - 28/11/2002 - Fault 4772 - Problem if date inputted is in MMDDYY (stupid yank format)
		dTempDate = DateSerial(piYear, piMonthNumber, 1)

		iStartNumber = Weekday(dTempDate, FirstDayOfWeek.Sunday) - 1
		'UPGRADE_WARNING: Couldn't resolve default property of object NumberOfDaysInMonth(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iEndNumber = iStartNumber + NumberOfDaysInMonth(dTempDate)

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
				dTempDate = DateSerial(piYear, piMonthNumber, iCount + 1 - iStartNumber)
				iIndexAM = CStr(GetCalIndex(dTempDate, False))
				iIndexPM = CStr(GetCalIndex(dTempDate, True))

				' Is a weekend
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(CInt(iIndexAM), 1) = (Weekday(dTempDate, FirstDayOfWeek.Monday) > 5)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(CInt(iIndexPM), 1) = (Weekday(dTempDate, FirstDayOfWeek.Monday) > 5)

				'------------------------------------------------
				'AM
				'------------------------------------------------
				If (dTempDate < mdStartDate) Or (dTempDate > mdLeavingDate And Not mdLeavingDate = DateTime.FromOADate(0)) Then
					strHtmlCellString = "<TD name=DateID_" & LTrim(Str(CDbl(iIndexAM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexAM))) & "class=""calendar_nonday"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD>" & vbNewLine
				Else
					'Build the cell string

					' Build onclick event code
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mavAbsences(CInt(iIndexAM), 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strHtmlCellString = "<TD style='font-size: " & IIf(Len(mavAbsences(CInt(iIndexAM), 2)) < 2, "8", "6") & "pt;background-color:" & mavAbsences(CInt(iIndexAM), 5) & "' name=DateID_" _
								& LTrim(Str(CDbl(iIndexAM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexAM))) & " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP " _
								& " onclick=""ShowDetails('" & VB6.Format(mavAbsences(CInt(iIndexAM), 10), mstrClientDateFormat) & "','" _
								& mavAbsences(CInt(iIndexAM), 11) & "','" & VB6.Format(mavAbsences(CInt(iIndexAM), 12), mstrClientDateFormat) & "','" & mavAbsences(CInt(iIndexAM), 13) & "','" & mavAbsences(CInt(iIndexAM), 9) & "','" _
								& Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexAM), 6), 0), "'", "") & "','" & Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexAM), 6), 5), "'", "") & "','" _
								& Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexAM), 6), 4), "'", "") & "','" & HTMLEncode(Left(mavAbsences(CInt(iIndexAM), 7), 100)) & "','" & mavAbsences(CInt(iIndexAM), 14) & "','" _
								& mavAbsences(CInt(iIndexAM), 8) & "')"">" & "<FONT SIZE='1'>" & mavAbsences(CInt(iIndexAM), 2) & "</FONT>" & "</TD>" & vbNewLine
					Else
						strHtmlCellString = "<TD name=DateID_" & LTrim(Str(CDbl(iIndexAM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexAM))) & " class=""calendar_day"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" _
								& CELLSIZE & " NOWRAP>&nbsp;</TD>" & vbNewLine
					End If

				End If

				' Add current cell to the table
				'strHTML_Days = "<TR" & strHTMLColourCode & " " & strHTMLMouseCode & strHTMLOnClickCode & " >" _
				'& strHTMLCellString & "</TR>"
				strHtmlDays = "<TR>" & strHtmlCellString & "</TR>"


				'------------------------------------------------
				'PM
				'------------------------------------------------
				If (dTempDate < mdStartDate) Or (dTempDate > mdLeavingDate And Not mdLeavingDate = DateTime.FromOADate(0)) Then
					strHtmlCellString = "<TD name=DateID_" & LTrim(Str(CDbl(iIndexPM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexPM))) & " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP></TD>" & vbNewLine
				Else

					' Build onclick event code
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mavAbsences(CInt(iIndexPM), 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexAM, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iIndexPM, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strHtmlCellString = "<TD style='font-size: " & IIf(Len(mavAbsences(CInt(iIndexPM), 2)) < 2, "8", "6") & "pt;background-color:" & mavAbsences(CInt(iIndexPM), 5) & "' name=DateID_" & LTrim(Str(CDbl(iIndexPM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexPM))) _
								& " HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP" _
								& " onclick=""ShowDetails('" & VB6.Format(mavAbsences(CInt(iIndexPM), 10), mstrClientDateFormat) & "','" & mavAbsences(CInt(iIndexPM), 11) & "','" _
								& VB6.Format(mavAbsences(CInt(iIndexPM), 12), mstrClientDateFormat) & "','" & mavAbsences(CInt(iIndexPM), 13) & "','" & mavAbsences(CInt(iIndexPM), 9) _
								& "','" & Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexPM), 6), 0), "'", "") & "','" & Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexPM), 6), 5), "'", "") _
								& "','" & Replace(mastrAbsenceTypes(mavAbsences(CInt(iIndexPM), 6), 4), "'", "") & "','" & HTMLEncode(Left(mavAbsences(CInt(iIndexPM), 7), 100)) & "','" _
								& mavAbsences(CInt(iIndexPM), 14) & "','" & mavAbsences(CInt(iIndexPM), 8) & "')"">" & "<FONT SIZE='1'>" & mavAbsences(CInt(iIndexAM), 2) & "</FONT>" & "</TD>" & vbNewLine
					Else
						strHtmlCellString = "<TD name=DateID_" & LTrim(Str(CDbl(iIndexPM))) & " id=DateID_" & LTrim(Str(CDbl(iIndexPM))) & " class=""calendar_day"" HEIGHT=" & CELLSIZE & " VALIGN=middle ALIGN=center WIDTH=" & CELLSIZE & " NOWRAP>&nbsp;</TD>"
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

	Private Function NumberOfDaysInMonth(ByRef dtInput As Date) As Integer
		Return DateTime.DaysInMonth(Year(dtInput), Month(dtInput))
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
		sSQL = sSQL & "           INNER JOIN " & gsAbsenceTypeTableName & vbNewLine
		sSQL = sSQL & "           ON " & mstrAbsenceTableRealSource & "." & gsAbsenceTypeColumnName & " = " & gsAbsenceTypeTableName & "." & gsAbsenceTypeTypeColumnName & vbNewLine

		sSQL = sSQL & "WHERE " & mstrAbsenceTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & vbNewLine
		sSQL = sSQL & " AND (" & mstrSQLSelect_AbsenceStartDate & " IS NOT NULL) " & vbNewLine
		sSQL = sSQL & "ORDER BY 'StartDate' ASC"

		mrstAbsenceRecords = dataAccess.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

		' Set amount of absence records found
		Return mrstAbsenceRecords.RecordCount

GetAbsenceRecordSet_ERROR:

		'MsgBox "Error retrieving the Absence recordset." & vbNewLine & Err.Description, vbExclamation + vbOKOnly, App.Title
		'UPGRADE_NOTE: Object mrstAbsenceRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mrstAbsenceRecords = Nothing
		Return 0

	End Function

	Public Sub Initialise()

		Dim objTableView As TablePrivilege
		Dim fOK As Boolean

		fOK = True

		' Read the necessary settings for the calendar to work
		SetupTablesCollection()
		ReadPersonnelParameters()
		ReadAbsenceParameters()
		ReadBankHolidayParameters()

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

		' structure of absence types deocumented in declaration section
		ReDim mavAbsences(733, 14)

		' Only load the records from the DB once
		GetPersonnelRecordSet()

		GetWorkingPatterns()

		'GetRegions
		miAbsenceRecordsFound = GetAbsenceRecordSet()

		LoadColourKey()

		' Default start and end dates
		mdCalendarStartDate = DateSerial(Year(Now), giAbsenceCalStartMonth, 1)
		mdCalendarEndDate = DateTime.FromOADate(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mdCalendarStartDate).ToOADate - DateTime.FromOADate(0.5).ToOADate)

	End Sub

	' Loads the absence types
	Private Function LoadColourKey() As Boolean

		' Have colour already been loaded?
		If mbColourKeyLoaded Then
			LoadColourKey = True
		End If

		On Error GoTo errLoadColourKey

		Dim rstColourKey As ADODB.Recordset
		Dim strColourKeySQL As String
		Dim intCounter As Integer
		Dim strHexColour As String

		strColourKeySQL = "SELECT DISTINCT " & gsAbsenceTypeTypeColumnName & " AS Type, " & gsAbsenceTypeCalCodeColumnName & " AS CalCode," & gsAbsenceTypeCodeColumnName & " AS TypeCode" & " FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName
		rstColourKey = datGeneral.GetRecords(strColourKeySQL)

		If rstColourKey.BOF And rstColourKey.EOF Then
			'MsgBox "You have no absence types defined in your Absence Type table", vbExclamation + vbOKOnly, "Absence Calendar"
			LoadColourKey = False
			Exit Function
		End If

		'ReDim Preserve mastrAbsenceTypes(rstColourKey.RecordCount + 1, 5)
		ReDim Preserve mastrAbsenceTypes(20, 5)

		intCounter = 0
		Do Until rstColourKey.EOF

			If intCounter <= 18 Then

				' Set the colour box caption and show the label
				mastrAbsenceTypes(intCounter, 0) = rstColourKey.Fields(0).Value

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
				mastrAbsenceTypes(intCounter, 3) = UCase(Left(IIf(IsDBNull(rstColourKey.Fields("CalCode").Value), "", rstColourKey.Fields("CalCode").Value), 2))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(intCounter, 4) = Replace(IIf(IsDBNull(rstColourKey.Fields("CalCode").Value), "", rstColourKey.Fields("CalCode").Value), "'", "")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(intCounter, 5) = Replace(IIf(IsDBNull(rstColourKey.Fields("TypeCode").Value), "", rstColourKey.Fields("TypeCode").Value), "'", "")

			End If

			intCounter = intCounter + 1
			rstColourKey.MoveNext()

		Loop

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
		strHTML = strHTML & "   <TD>"

		bSecondColumn = False

		For intCounter = 0 To miStrAbsenceTypes	'UBound(mastrAbsenceTypes, 1) - 1

			' Position the colour box control depending on its index
			If intCounter >= 10 And Not bSecondColumn Then
				bSecondColumn = True
				strHTML = strHTML & "   </TD>" & vbNewLine
				strHTML = strHTML & "   <TD>" & vbNewLine
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
		For iCount = 1 To 12
			iMonth = Month(DateAdd(Microsoft.VisualBasic.DateInterval.Month, iCount - 1, mdCalendarStartDate))
			iYear = Year(DateAdd(Microsoft.VisualBasic.DateInterval.Month, iCount - 1, mdCalendarStartDate))

			strHtml = strHtml & HTML_Month(iMonth, iYear)
		Next iCount

		' Finish off the table text
		strHtml = strHtml & "</tbody></table>"

		' Return HTML code for the main calendar
		Return strHtml

	End Function

	Private Sub FillGridWithData()

		' Populate the grid with data
		On Error Resume Next

		' Load the colour key variables
		If Not LoadColourKey() Then
			Exit Sub
		End If

		Dim intStart As Integer
		Dim intEnd As Integer

		' If there are no absence records for the current employee then skip
		' this bit (but still show the form)
		If mrstAbsenceRecords.BOF And mrstAbsenceRecords.EOF Then
			Exit Sub
		End If

		mstrAbsWPattern = ""

		With mrstAbsenceRecords
			.MoveFirst()

			' Loop through the absence recordset
			Do Until .EOF
				' Load each absence record data into variables
				' (has to be done because start/end dates may be modified by code to fill grid correctly)

				' JDM - Kak-Handed way of sorting out American settings on different versions of IIS
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDBNull(.Fields("StartDate").Value) Then
					mdAbsStartDate = DateSerial(Year(Now), Month(Now), Now.Day)
				Else
					mdAbsStartDate = DateSerial(Year(.Fields("StartDate").Value), Month(.Fields("StartDate").Value), CDate(.Fields("StartDate").Value).Day)
				End If

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDBNull(.Fields("EndDate").Value) Then
					mdAbsEndDate = DateSerial(Year(Now), Month(Now), Now.Day)
				Else
					mdAbsEndDate = DateSerial(Year(.Fields("EndDate").Value), Month(.Fields("EndDate").Value), CDate(.Fields("EndDate").Value).Day)
				End If

				mstrAbsStartSession = CStr(.Fields("StartSession").Value).ToUpper()
				mstrAbsEndSession = CStr(.Fields("EndSession").Value).ToUpper()

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrAbsType = IIf(IsDBNull(.Fields("Type").Value), "", .Fields("Type").Value)

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrAbsCalendarCode = IIf(IsDBNull(.Fields("CalendarCode").Value), "", .Fields("CalendarCode").Value)
				mlngAbsDuration = .Fields("Duration").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrAbsReason = IIf(IsDBNull(.Fields("Reason").Value), "", .Fields("Reason").Value)

				If mdAbsStartDate <= mdCalendarEndDate And mdAbsEndDate >= mdCalendarStartDate Then
					intStart = GetCalIndex(mdAbsStartDate, mstrAbsStartSession = "PM")
					intEnd = GetCalIndex(mdAbsEndDate, mstrAbsEndSession = "PM")

					FillCalBoxes(intStart, intEnd)
				End If

				.MoveNext()
			Loop
		End With

	End Sub

	Private Function GetPersonnelRecordSet() As Boolean

		On Error GoTo PersonnelERROR

		Dim lngCount As Integer
		Dim sSQL As String
		Dim prstPersonnelData As ADODB.Recordset

		' Botch as we have a lot of rubbish code that does not handle nulls at all.
		mdStartDate = DateTime.FromOADate(0)
		mdLeavingDate = DateTime.FromOADate(0)

		If Not mblnFailReport Then
			sSQL = vbNullString
			sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStartDate & " AS 'StartDate', " & vbNewLine
			sSQL = sSQL & "      " & mstrSQLSelect_PersonnelLeavingDate & " AS 'LeavingDate' " & vbNewLine
			sSQL = sSQL & "FROM " & gsPersonnelTableName & vbNewLine
			For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
				'<Personnel CODE>
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
				End If
			Next lngCount
			sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID

			' Get the start and leaving date
			prstPersonnelData = datGeneral.GetRecords(sSQL)

			If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdStartDate = IIf(IsDBNull(prstPersonnelData.Fields("StartDate").Value), mdStartDate, VB6.Format(prstPersonnelData.Fields("StartDate").Value, DateFormat))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdLeavingDate = IIf(IsDBNull(prstPersonnelData.Fields("LeavingDate").Value), mdLeavingDate, VB6.Format(prstPersonnelData.Fields("LeavingDate").Value, DateFormat))
			End If
			prstPersonnelData.Close()
			'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstPersonnelData = Nothing
		Else
			GoTo PersonnelERROR
		End If

		If Not mblnDisableRegions Then
			' Get the employees current region
			If grtRegionType = RegionType.rtStaticRegion Then
				' Its a static region, get it from personnel
				sSQL = "SELECT " & mstrSQLSelect_PersonnelStaticRegion & "  AS 'Region'  " & vbNewLine
				sSQL = sSQL & "FROM " & gsPersonnelTableName & vbNewLine
				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<Personnel CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount
				sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID
				prstPersonnelData = datGeneral.GetRecords(sSQL)
			Else
				' Its a historic region, so get topmost from the history
				prstPersonnelData = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & gsPersonnelHRegionTableRealSource & " " & "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC")
			End If

			If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrRegion = Replace(IIf(IsDBNull(prstPersonnelData.Fields("Region").Value), "", IIf(prstPersonnelData.Fields("Region").Value = "", "", prstPersonnelData.Fields("Region").Value)), "&", "&&")
			Else
				mstrRegion = "&lt;None&gt;"
			End If
			prstPersonnelData.Close()
			'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstPersonnelData = Nothing
		Else
			'Regions DISABLED
			mstrRegion = vbNullString
		End If

		If Not mblnDisableWPs Then
			' Get the employees current working pattern
			If modPersonnelSpecifics.gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern Then
				' Its a static working pattern, get it from personnel
				sSQL = vbNullString
				sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStaticWP & "  AS 'WP'  " & vbNewLine
				sSQL = sSQL & "FROM " & gsPersonnelTableName & vbNewLine
				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<Personnel CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount
				sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID
				prstPersonnelData = datGeneral.GetRecords(sSQL)

			Else
				' Its a historic working pattern, so get topmost from the history
				prstPersonnelData = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " <= '" _
																									& Replace(VB6.Format(Now, "MM/dd/yy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")
			End If

			If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
				mstrWorkingPattern = prstPersonnelData.Fields("WP").Value
			Else
				mstrWorkingPattern = Space(14)
			End If

			prstPersonnelData.Close()
			'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstPersonnelData = Nothing

		Else
			'WPs DISABLED
			mstrAbsWPattern = "SSMMTTWWTTFFSS"

		End If

		'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstPersonnelData = Nothing
		GetPersonnelRecordSet = True
		Exit Function

PersonnelERROR:

		'UPGRADE_NOTE: Object prstPersonnelData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstPersonnelData = Nothing
		GetPersonnelRecordSet = False
		'MsgBox "Error whilst retrieving the personnel information." & vbNewLine & Err.Description, vbExclamation + vbOKOnly, App.Title
		Exit Function

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
			strHtml = strHtml & "<TR bordercolor=" & mstrHexColour_OptionBoxes & "><TD nowrap>&nbsp;Working Pattern :&nbsp;&nbsp;</TD><TD>" & HTML_WorkingPattern(mstrWorkingPattern) & "</TD></TR>" & vbNewLine
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
		Dim strHtmlRefresh As String
		Dim strCaption As String

		' Create function header strings
		strHtml = "<script type=""text/javascript"">" & vbNewLine

		strHtmlRefresh = vbNullString
		strHtmlRefresh = strHtmlRefresh & "function refreshDateSpecifics() {" & vbNewLine
		strHtmlRefresh = strHtmlRefresh & " refreshToggleValues();" & vbNewLine

		' Create option strings
		For iCount = 0 To UBound(mavAbsences, 1)

			dTempDate = GetCalDay(iCount)

			If (dTempDate <= mdLeavingDate Or mdLeavingDate = DateTime.FromOADate(0)) And dTempDate >= mdStartDate And (dTempDate <= mdCalendarEndDate And dTempDate >= mdCalendarStartDate) Then

				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsBankHoliday = mavAbsences(iCount, 3)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsWeekend = mavAbsences(iCount, 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strColour = mavAbsences(iCount, 5)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCaption = mavAbsences(iCount, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnHasEvent = mavAbsences(iCount, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(iCount, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnIsWorkingDay = mavAbsences(iCount, 4)

				If (Not blnIsWeekend) And (Not blnHasEvent) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine
				End If

				If blnIsWeekend And (Not blnHasEvent) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtShowWeekends.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   }" & vbNewLine
				End If


				'Has an event therefore deal with the Caption
				If blnHasEvent And (Not blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   }" & vbNewLine
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (blnIsWeekend) And (Not blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowWeekends.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowWeekends.value == ""unhighlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   }" & vbNewLine
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (blnIsWeekend) And (blnIsBankHoliday) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtIncludeBankHolidays.value == ""included"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   }" & vbNewLine
					strHtmlRefresh = strHtmlRefresh & "else if (frmChangeDetails.txtIncludeBankHolidays.value == ""unincluded"" && frmChangeDetails.txtShowWeekends.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtIncludeBankHolidays.value == ""unincluded"" && frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine
					strHtmlRefresh = strHtmlRefresh & "else if (frmChangeDetails.txtIncludeBankHolidays.value == ""unincluded"" && frmChangeDetails.txtShowWeekends.value == ""unhighlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtIncludeBankHolidays.value == ""unincluded"" && frmChangeDetails.txtShowBankHolidays.value == ""unhighlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine
					strHtmlRefresh = strHtmlRefresh & "else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowWeekends.value == ""unhighlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine
					strHtmlRefresh = strHtmlRefresh & "else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowWeekends.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"" && frmChangeDetails.txtShowBankHolidays.value == ""unhighlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   }" & vbNewLine
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And (Not blnIsBankHoliday) And (blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtIncludeBankHolidays.value == ""included"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine
				End If

				'Has an event therefore deal with the Caption
				If blnHasEvent And blnIsBankHoliday And (Not blnIsWeekend) And (blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtIncludeBankHolidays.value == ""included"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """ & strColour & """;" & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = """ & strCaption & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "     }" & vbNewLine & "   }" & vbNewLine & "else if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".innerText = '';" & vbNewLine & "   }" & vbNewLine
				End If

				If (Not blnHasEvent) And blnIsBankHoliday And (Not blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   }" & vbNewLine
				End If

				If (Not blnHasEvent) And blnIsBankHoliday And (blnIsWeekend) And (Not blnIsWorkingDay) Then
					strHtmlRefresh = strHtmlRefresh & "if (frmChangeDetails.txtShowBankHolidays.value == ""highlighted"") " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_nonworkingday';" & vbNewLine & "   }" & vbNewLine & "else " & vbNewLine & "   {" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".style.backgroundColor = """";" & vbNewLine & "   DateID_" & LTrim(Str(iCount)) & ".className = 'calendar_day';" & vbNewLine & "   }" & vbNewLine
				End If
			End If
		Next iCount

		For iCount = 0 To miStrAbsenceTypes Step 1
			strHtmlRefresh = strHtmlRefresh & vbNewLine & vbNewLine & "   if (frmChangeDetails.txtShowCaptions.value == 'show') " & vbNewLine & "     {" & vbNewLine & "     KEY_" & LTrim(Str(iCount)) & ".innerHTML = """ & IIf(Trim(mastrAbsenceTypes(iCount, 3)) = vbNullString, "&nbsp", mastrAbsenceTypes(iCount, 3)) & """;" & vbNewLine & "     }" & vbNewLine & "   else" & vbNewLine & "     {" & vbNewLine & "     KEY_" & LTrim(Str(iCount)) & ".innerHTML = '&nbsp';" & vbNewLine & "     }" & vbNewLine
		Next iCount

		strHtmlRefresh = strHtmlRefresh & " }" & vbNewLine

		' Concatenate functions into HTML string
		Return strHtml & strHtmlRefresh & vbNewLine & "</script>" & vbNewLine

	End Function

	Private Function FillCalBoxes(ByRef intStart As Integer, ByRef intEnd As Integer) As Boolean

		' This function actually fills the cal boxes between the indexes specified
		' according to the options selected by the user.

		On Error GoTo Error_FillCalBoxes

		Dim iCount As Integer
		Dim dtmCurrentDate As Date
		Dim strColour As String
		Dim iArrayCount As Integer

		'Scroll forward in list to correct start working pattern for absence.
		dtmCurrentDate = GetCalDay(intStart)
		miWorkingPatternArray = 0
		For iArrayCount = 0 To UBound(mavWorkingPatternChanges, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, miWorkingPatternArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If dtmCurrentDate > mavWorkingPatternChanges(0, miWorkingPatternArray) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, iArrayCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mstrAbsWPattern = mavWorkingPatternChanges(1, iArrayCount)
				miWorkingPatternArray = miWorkingPatternArray + 1

				If miWorkingPatternArray > UBound(mavWorkingPatternChanges, 2) Then
					miWorkingPatternArray = miWorkingPatternArray - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, miWorkingPatternArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdtmWorkingPatternDate = mavWorkingPatternChanges(0, miWorkingPatternArray)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, miWorkingPatternArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdtmWorkingPatternDate = mavWorkingPatternChanges(0, miWorkingPatternArray)
				End If
			End If
		Next iArrayCount

		' Loop through the indexes as specified.
		For iCount = intStart To intEnd

			' Set current date variable
			dtmCurrentDate = GetCalDay(iCount)


			'Calculate the working pattern for this day
			If dtmCurrentDate >= mdtmWorkingPatternDate Then

				'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, miWorkingPatternArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mstrAbsWPattern = mavWorkingPatternChanges(1, miWorkingPatternArray)
				miWorkingPatternArray = miWorkingPatternArray + 1

				If miWorkingPatternArray <= UBound(mavWorkingPatternChanges, 2) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, miWorkingPatternArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdtmWorkingPatternDate = mavWorkingPatternChanges(0, miWorkingPatternArray)
				Else
					mdtmWorkingPatternDate = CDate("31/12/9999")
				End If

			End If

			' Mark this day as having an absence
			If Not mavAbsences(iCount, 0) Then
				strColour = GetColour(mstrAbsType)
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(iCount, 0) = True
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(iCount, 6) = GetAbsenceCode(mstrAbsType) ' Absence type for this day
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(iCount, 2) = mstrAbsCalendarCode
			Else
				strColour = GetColour("Multiple")
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(iCount, 6) = GetAbsenceCode("Multiple")	' Absence type for this day
				'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mavAbsences(iCount, 2) = "."
			End If

			' Is this day a working day
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 4) = AbsCal_DoTheyWorkOnThisDay(Weekday(dtmCurrentDate, FirstDayOfWeek.Sunday), IIf(iCount Mod 2 = 0, "AM", "PM"))

			' Store the details for this day
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 5) = strColour ' Absence display colour
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 7) = Replace(mstrAbsReason, "'", "") ' Absence reason
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 8) = mstrAbsWPattern ' Working pattern
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 9) = LTrim(CStr(mlngAbsDuration))	' Duration
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 10) = mdAbsStartDate 'Format(mdAbsStartDate, DateFormat) ' Start date of absence
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 11) = mstrAbsStartSession	' Start session of absence
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 12) = mdAbsEndDate 'Format(mdAbsEndDate, DateFormat)   ' End date of absence
			'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(Count, 13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavAbsences(iCount, 13) = mstrAbsEndSession	' End session of absence
			'    mavAbsences(Count, 14) = Replace(mstrAbsRegion, "'", "''")    ' Region

		Next iCount

		FillCalBoxes = True
		Exit Function

Error_FillCalBoxes:

		FillCalBoxes = False

	End Function

	Private Function GetColour(ByRef strType As String) As String

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

		GetColour = strColourString

	End Function


	Public Function HTML_HighlightAbsenceTypes() As String

		'Build a function for highlighting the current absence type
		Dim strHtml As String

		' Create function header strings
		strHtml = "<script type=""text/javascript"">" & vbNewLine & "function HighlightAbsenceTypes(pstrAbsenceType, pbWorkingDay){" & vbNewLine

		strHtml = strHtml & "if (pbWorkingDay == true and frmChangeDetails.txtIncludeWorkingDaysOnly.value == ""included"")" & "{" & "opener.document.getElementById(pstrAbsenceType).style.visibility=""hidden""}" & vbNewLine


		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_HighlightAbsenceTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_HighlightAbsenceTypes = strHtml & vbNewLine & "}" & vbNewLine & "</script>" & vbNewLine

	End Function

	Public Function HTML_UnHighlightAbsenceTypes() As String

		'Build a function for highlighting the current absence type
		Dim strHtml As String

		' Create function header strings
		strHtml = "<acript type""=javascript"">" & vbNewLine & "function UnHighlightAbsenceTypes(pstrAbsenceType, pbWorkingDay){" & vbNewLine

		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_UnHighlightAbsenceTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_UnHighlightAbsenceTypes = strHtml & vbNewLine & "}" & vbNewLine & "</script>" & vbNewLine

	End Function

	Private Function GetAbsenceCode(ByRef strType As String) As String

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

	Public Function WeekDayMonthStart(ByRef dtInput As Date) As Object

		'Pass a full date into this function and it will return the
		'vb constant for the day of the week that month started

		WeekDayMonthStart = Weekday(DateAdd(Microsoft.VisualBasic.DateInterval.Day, (VB.Day(dtInput) - 1) * -1, dtInput))

	End Function

	Public Function GetCalDay(ByRef intIndex As Integer) As Date

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
		GetCalDay = DateTime.FromOADate(mdCalendarStartDate.ToOADate + ((intIndex) / 2))

	End Function

	Public Function AbsCal_DoTheyWorkOnThisDay(ByVal intDay As Integer, ByVal strperiod As String) As Boolean

		Dim bFound As Boolean = True

		' Inputs  - 1 to 7 depending on the weekday 1 = sunday etc, "AM" or "PM"
		' Outputs - True/False
		Select Case UCase(strperiod)
			Case "AM"
				If (Mid(mstrAbsWPattern, (intDay * 2) - 1, 1) = " ") Or (Mid(mstrAbsWPattern, (intDay * 2) - 1, 1) = "") Then
					bFound = False
				End If
			Case "PM"
				If (Mid(mstrAbsWPattern, intDay * 2, 1) = " ") Or (Mid(mstrAbsWPattern, intDay * 2, 1) = "") Then
					bFound = False
				End If
		End Select

		Return bFound

	End Function

	Public Function GetCalIndex(ByVal dtmDate As Date, ByVal booSession As Boolean) As Integer

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

	Public Function HTML_WorkingPattern(ByRef pstrWorkingPattern As String) As String

		Dim strHtml As String
		Dim iCount As Integer

		pstrWorkingPattern = pstrWorkingPattern & Space(14 - Len(pstrWorkingPattern))

		strHtml = "<table class='invisible' cellspacing=0 cellpadding=0 frame=0>" & vbNewLine

		' Row 1 contains day names
		strHtml = strHtml & "<TR align=middle>" & "<td>&nbsp;</td><TD>" & UCase(Left(VB6.Format(1, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(2, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(3, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(4, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(5, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(6, "ddd"), 1)) & "</TD>" & "<TD>" & UCase(Left(VB6.Format(7, "ddd"), 1)) & "</TD></TR>" & vbNewLine

		' Row two contains the AM fields
		strHtml = strHtml & "<TR><td>AM</td>"

		For iCount = 1 To 13 Step 2
			If Not Mid(pstrWorkingPattern, iCount, 1) = " " Then
				strHtml = strHtml & "<TD><INPUT id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" checked disabled></TD>"
			Else
				strHtml = strHtml & "<TD><INPUT id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled></TD>"
			End If
		Next iCount
		strHtml = strHtml & "</TR>"


		' Row three contains the PM fields
		strHtml = strHtml & "<TR><td>PM</td>"
		For iCount = 2 To 14 Step 2
			strHtml = strHtml & "<TD><INPUT id=checkbox1 name=checkbox1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"""
			If Not Mid(pstrWorkingPattern, iCount, 1) = " " Then
				strHtml = strHtml & " Checked"
			End If
			strHtml = strHtml & " disabled></TD>"
		Next iCount
		strHtml = strHtml & "</TR></TABLE>"

		HTML_WorkingPattern = strHtml
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

		Dim iCount As Integer
		Dim rstHistoricWPatterns As ADODB.Recordset
		Dim sSQL As String
		Dim lngCount As Integer

		' Define a blank working pattern array
		ReDim mavWorkingPatternChanges(1, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mavWorkingPatternChanges(0, 0) = CDate("01/01/1899")
		'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mavWorkingPatternChanges(1, 0) = Space(14)

		If Not mblnDisableWPs Then
			' If we are using historic WPattern, ensure we use the right WPattern for each day of absence
			If gwptWorkingPatternType = WorkingPatternType.wptHistoricWPattern Then

				' Get the wpattern for the start of the absence period
				rstHistoricWPatterns = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " AS 'Date', " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " <= '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")

				If Not (rstHistoricWPatterns.BOF And rstHistoricWPatterns.EOF) Then

					' Start working pattern for this employee
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavWorkingPatternChanges(0, 0) = rstHistoricWPatterns.Fields("Date").Value

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mstrAbsWPattern = IIf(IsDBNull(rstHistoricWPatterns.Fields("WP").Value), Space(14), rstHistoricWPatterns.Fields("WP").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavWorkingPatternChanges(1, 0) = mstrAbsWPattern

				End If

				' Now get the rest of the working patterns
				Dim sSQLWorkingPatterns As String = String.Format("SELECT " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " AS 'Date', " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " > '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " ASC")
				rstHistoricWPatterns = dataAccess.OpenRecordset(sSQLWorkingPatterns, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

				If Not (rstHistoricWPatterns.EOF And rstHistoricWPatterns.BOF) Then

					' Size the array for the amount of working patterns this employee has.
					ReDim Preserve mavWorkingPatternChanges(1, rstHistoricWPatterns.RecordCount)

					rstHistoricWPatterns.MoveFirst()

					' Load all the working patterns into array
					For iCount = 1 To rstHistoricWPatterns.RecordCount

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						mstrAbsWPattern = IIf(IsDBNull(rstHistoricWPatterns.Fields("WP").Value), Space(14), rstHistoricWPatterns.Fields("WP").Value)
						mstrAbsWPattern = mstrAbsWPattern & Space(14 - Len(mstrAbsWPattern))

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavWorkingPatternChanges(0, iCount) = IIf(IsDBNull(rstHistoricWPatterns.Fields("Date").Value), CDate("01/01/1899"), rstHistoricWPatterns.Fields("Date").Value)
						'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavWorkingPatternChanges(1, iCount) = mstrAbsWPattern

						' Go to next record
						rstHistoricWPatterns.MoveNext()

					Next iCount

					'Else

					' Size the array for the amount of working patterns this employee has.
					'ReDim Preserve mavWorkingPatternChanges(1, 1)

					'mavWorkingPatternChanges(0, 1) = CDate("31/12/9999")
					'mavWorkingPatternChanges(1, 1) = Space(14)

				End If

				'UPGRADE_NOTE: Object rstHistoricWPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstHistoricWPatterns = Nothing

			Else

				' Its a static working pattern, get it from personnel
				sSQL = vbNullString
				sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStaticWP & "  AS 'WP'  " & vbNewLine
				sSQL = sSQL & "FROM " & gsPersonnelTableName & vbNewLine
				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<Personnel CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount
				sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelRecordID

				rstHistoricWPatterns = datGeneral.GetRecords(sSQL)

				' Stuff the working pattern into array
				If Not (rstHistoricWPatterns.EOF And rstHistoricWPatterns.BOF) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavWorkingPatternChanges(0, 0) = CDate("01/01/1899")
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavWorkingPatternChanges(1, 0) = Left(IIf(IsDBNull(rstHistoricWPatterns.Fields("WP").Value), Space(14), rstHistoricWPatterns.Fields("WP").Value) & Space(14), 14)
				End If

			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavWorkingPatternChanges(0, 0) = CDate("01/01/1899")
			'UPGRADE_WARNING: Couldn't resolve default property of object mavWorkingPatternChanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mavWorkingPatternChanges(1, 0) = FULL_WP
		End If

	End Sub

	Private Sub GenerateRegionData()

		Dim intTemp As Integer
		Dim bNewRegionFound As Boolean
		Dim strRegionAtCurrentDate As String
		Dim dtmNextChangeDate As Date
		Dim intCount As Integer
		Dim rstBankHolRegion As ADODB.Recordset
		Dim dtmCurrentDate As Date
		Dim sSQL As String
		Dim lngCount As Integer

		bNewRegionFound = False

		If Not mblnDisableRegions Then
			' If we are using historic region, find the region change dates
			If grtRegionType = RegionType.rtHistoricRegion Then

				' Get the first region for this employee within this calendar year
				rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & gsPersonnelHRegionTableRealSource & " " & "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " <= '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC")

				' Was there a region at the start of the calendar
				If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
					strRegionAtCurrentDate = ""
				Else
					strRegionAtCurrentDate = rstBankHolRegion.Fields("Region").Value
					bNewRegionFound = True
				End If

				' Get the second region for this employee within this calendar year
				rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & gsPersonnelHRegionTableRealSource & " " & "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " > '" & VB6.Format(mdCalendarStartDate, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")

				' Was there a region at the start of the calendar
				If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
					dtmNextChangeDate = CDate("31/12/9999")
				Else
					dtmNextChangeDate = rstBankHolRegion.Fields("Date").Value
				End If


				For intCount = LBound(mavAbsences, 1) To UBound(mavAbsences, 1) Step 2

					' Get the date of the current index
					dtmCurrentDate = GetCalDay(intCount)

					' Only refer to the region table if the current date is a region change date
					If (dtmCurrentDate >= dtmNextChangeDate) And (dtmCurrentDate <> CDate("31/12/9999")) Then


						'JDM - 11/09/01 - Fault 2820 - Bank hols not showing for year starting with working pattern.
						' Find the employees region for this date
						rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & gsPersonnelHRegionTableRealSource & " " & "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " >= '" & VB6.Format(dtmNextChangeDate, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")

						If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then

							' No regions found for this user
							dtmNextChangeDate = CDate("31/12/9999")

						Else

							strRegionAtCurrentDate = rstBankHolRegion.Fields("Region").Value
							bNewRegionFound = True

							' Now get the next change date
							rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & "FROM " & gsPersonnelHRegionTableRealSource & " " & "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelRecordID & " " & "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " > '" & VB6.Format(rstBankHolRegion.Fields("Date").Value, "MM/dd/yyyy") & "' " & "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")
							If rstBankHolRegion.EOF Then
								dtmNextChangeDate = CDate("31/12/9999")
							Else
								dtmNextChangeDate = rstBankHolRegion.Fields("Date").Value
							End If

						End If

					End If

					' Define the region for this period
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavAbsences(intCount, 14) = Replace(strRegionAtCurrentDate, "'", "''")
					'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount + 1, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mavAbsences(intCount + 1, 14) = Replace(strRegionAtCurrentDate, "'", "''")

					' If current region has changed
					If bNewRegionFound Then

						If gfBankHolidaysEnabled Then

							' Get bank holidays for this region
							' DONE
							sSQL = vbNullString
							sSQL = sSQL & "SELECT " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " AS 'Date' " & vbNewLine
							sSQL = sSQL & "FROM " & gsBHolTableRealSource & " " & vbNewLine

							sSQL = sSQL & "WHERE " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & " = " & vbNewLine
							sSQL = sSQL & "        (SELECT " & gsBHolRegionTableName & ".ID " & vbNewLine
							sSQL = sSQL & "         FROM " & gsBHolRegionTableName & vbNewLine
							For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
								'<REGIONAL CODE>
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sSQL = sSQL & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
								End If
							Next lngCount
							sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbNewLine

							sSQL = sSQL & " AND " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(VB6.Format(dtmCurrentDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & vbNewLine
							sSQL = sSQL & " AND " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(VB6.Format(DateTime.FromOADate(dtmNextChangeDate.ToOADate - 1), "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "' " & vbNewLine
							sSQL = sSQL & "ORDER BY " & gsBHolDateColumnName & " ASC"
							rstBankHolRegion = datGeneral.GetRecords(sSQL)

							' Cycle through the recordset checking for the current day
							If Not (rstBankHolRegion.BOF And rstBankHolRegion.EOF) Then

								rstBankHolRegion.MoveFirst()
								Do Until rstBankHolRegion.EOF
									intTemp = GetCalIndex(CDate(rstBankHolRegion.Fields("Date").Value), False)

									'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mavAbsences(intTemp, 3) = True
									'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp + 1, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mavAbsences(intTemp + 1, 3) = True

									rstBankHolRegion.MoveNext()
								Loop
							End If

						End If

						' Flag this region has had it's bank holidays drawn
						bNewRegionFound = False

					End If

				Next intCount

			Else

				If gfBankHolidaysEnabled Then

					' We are using a static region so just use the employees current region
					strRegionAtCurrentDate = mstrRegion
					' DONE
					sSQL = vbNullString
					sSQL = sSQL & "SELECT " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " AS 'Date' " & vbNewLine
					sSQL = sSQL & "FROM " & gsBHolTableRealSource & " " & vbNewLine
					sSQL = sSQL & "WHERE " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & " = " & vbNewLine
					sSQL = sSQL & "        (SELECT " & gsBHolRegionTableName & ".ID " & vbNewLine
					sSQL = sSQL & "         FROM " & gsBHolRegionTableName & vbNewLine
					For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
						'<REGIONAL CODE>
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
						End If
					Next lngCount
					sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbNewLine
					sSQL = sSQL & "ORDER BY " & gsBHolDateColumnName & " ASC" & vbNewLine

					rstBankHolRegion = datGeneral.GetRecords(sSQL)

					' Cycle through the recordset checking for the current day
					If Not (rstBankHolRegion.BOF And rstBankHolRegion.EOF) Then
						rstBankHolRegion.MoveFirst()
						Do Until rstBankHolRegion.EOF

							intTemp = GetCalIndex(CDate(rstBankHolRegion.Fields("Date").Value), False)

							'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavAbsences(intTemp, 3) = True
							'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intTemp + 1, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavAbsences(intTemp + 1, 3) = True

							rstBankHolRegion.MoveNext()
						Loop
					End If

					' Define the region for this period
					For intCount = LBound(mavAbsences, 1) To UBound(mavAbsences, 1) Step 2

						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intCount, 14) = Replace(strRegionAtCurrentDate, "'", "''")
						'UPGRADE_WARNING: Couldn't resolve default property of object mavAbsences(intCount + 1, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mavAbsences(intCount + 1, 14) = Replace(strRegionAtCurrentDate, "'", "''")

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
		If gsBHolRegionTableName = "" Or gsBHolRegionColumnName = "" Or gsBHolTableName = "" Or gsBHolDateColumnName = "" Or gsBHolDescriptionColumnName = "" Then

			GoTo DisableRegions
		End If

		'Check the  Career Change Region - Static Region Column
		'           Career Change Region - Historic Region Table
		'           Career Change Region - Historic Region Column
		'           Career Change Region - Historic Region Effective Date Column
		'...Personnel - Career Change module setup information.
		'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
		If gsPersonnelRegionColumnName = "" Then
			If gsPersonnelHRegionTableName = "" Or gsPersonnelHRegionColumnName = "" Or gsPersonnelHRegionDateColumnName = "" Then

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
		If CheckPermission_Columns(glngBHolRegionTableID, gsBHolRegionTableName, gsBHolRegionColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, gsBHolDateColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, gsBHolDescriptionColumnName, strTableColumn) Then
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
		If gsPersonnelRegionColumnName <> "" Then
			'Personnel Table
			'Career Change Region - Static Region Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelRegionColumnName, strTableColumn) Then
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
			If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, gsPersonnelHRegionColumnName, strTableColumn) Then
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
			If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, gsPersonnelHRegionDateColumnName, strTableColumn) Then
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
		If gsPersonnelWorkingPatternColumnName = "" Then
			If gsPersonnelHWorkingPatternTableName = "" Or gsPersonnelHWorkingPatternColumnName = "" Or gsPersonnelHWorkingPatternDateColumnName = "" Then

				GoTo DisableWPs
			End If
		End If

		'****************************************************************************
		' All Working Pattern module information is setup                           *
		' Now check the permissions on the Working Pattern module setup information *
		'****************************************************************************
		'Check Career Change Working Pattern access
		If gsPersonnelWorkingPatternColumnName <> "" Then
			'Career Change Working Pattern - Static Working Pattern Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelWorkingPatternColumnName, strTableColumn) Then
				mstrSQLSelect_PersonnelStaticWP = strTableColumn
				strTableColumn = vbNullString
			Else
				GoTo DisableWPs
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		Else
			'Career Change Working Pattern - Historic Working Pattern Table
			objColumn = GetColumnPrivileges(gsPersonnelHWorkingPatternTableName)

			'Career Change Working Pattern - Historic Working Pattern Column
			pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternColumnName).AllowSelect
			End If
			If pblnColumnOK = False Then
				GoTo DisableWPs
			End If

			'Career Change Working Pattern - Historic Working Pattern Effective Date Column
			pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternDateColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternDateColumnName).AllowSelect
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
		If gsAbsenceTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Table' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceStartDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Date Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceStartSessionColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Session Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceEndDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'End Date Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceEndSessionColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'End Session Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceTypeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceReasonColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Reason Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceDurationColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Duration Column' in the Absence module setup must be defined." & vbNewLine
		End If


		'Check the Absence Type Table
		'          Absence Type Table - Absence Type Column
		'          Absence Type Table - Absence Code Column
		'          Absence Type Table - Calendar Code Column
		'...Absence module setup information.
		'If any are blank then we need to fail the Absence Calendar report.
		If gsAbsenceTypeTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Table' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceTypeTypeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Type Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceTypeCodeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Absence Code Column' in the Absence module setup must be defined." & vbNewLine
		End If
		If gsAbsenceTypeCalCodeColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Calendar Code Column' in the Absence module setup must be defined." & vbNewLine
		End If


		'Check the Personnel Table
		'          Personnel Table - Start Date Column
		'          Personnel Table - Leaving Date Column
		'...Personnel module setup information.
		'If any are blank then we need to fail the Absence Calendar report.
		If gsPersonnelTableName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Personnel Table' in the Personnel module setup must be defined." & vbNewLine
		End If
		If gsPersonnelStartDateColumnName = "" Then
			strModulePermErrorMSG = strModulePermErrorMSG & "The 'Start Date Column' in the Personnel module setup must be defined." & vbNewLine
		End If
		If gsPersonnelLeavingDateColumnName = "" Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceStartDateColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceStartSessionColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceEndDateColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceEndSessionColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceTypeColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceReasonColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, gsAbsenceDurationColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, gsAbsenceTypeTypeColumnName, strTableColumn) Then
			strTableColumn = vbNullString
		Else
			strModulePermErrorMSG = strModulePermErrorMSG & "Permission Denied on 'Absence Type Table - Absence Type Column'" & vbNewLine
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		'Absence Type Table - Absence Code Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, gsAbsenceTypeCodeColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, gsAbsenceTypeCalCodeColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelStartDateColumnName, strTableColumn) Then
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
		If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelLeavingDateColumnName, strTableColumn) Then
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

	Private Function CheckPermission_Columns(ByRef plngTableID As Integer, ByRef pstrTableName As String, ByRef pstrColumnName As String, ByRef strSQLRef As String) As Boolean

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

			If (plngTableID = glngAbsenceTableID) And (mstrAbsenceTableRealSource = vbNullString) Then
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

			Dim mstrViews(0) As Object
			For Each mobjTableView In gcoTablePrivileges.Collection
				If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTempTableID) And (mobjTableView.AllowSelect) Then

					strSource = mobjTableView.ViewName
					mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

					' Get the column permission for the view
					mobjColumnPrivileges = GetColumnPrivileges(strSource)

					' If we can see the column from this view
					If mobjColumnPrivileges.IsValid(strTempColumnName) Then
						If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

							ReDim Preserve mstrViews(UBound(mstrViews) + 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

							' Check if view has already been added to the array
							blnFound = False
							For intNextIndex = 0 To UBound(mvarTableViews, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarTableViews(1, intNextIndex) = 1 And mvarTableViews(2, intNextIndex) = mobjTableView.ViewID Then
									blnFound = True
									Exit For
								End If
							Next intNextIndex

							If Not blnFound Then
								' View hasnt yet been added, so add it !
								intNextIndex = UBound(mvarTableViews, 2) + 1
								ReDim Preserve mvarTableViews(3, intNextIndex)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(0, intNextIndex) = mobjTableView.TableID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(1, intNextIndex) = 1
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(2, intNextIndex) = mobjTableView.ViewID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(3, intNextIndex) = mobjTableView.ViewName
							End If

						End If
					End If
				End If

			Next mobjTableView
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

		CheckPermission_Columns = True

	End Function

End Class