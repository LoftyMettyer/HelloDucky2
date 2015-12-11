Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports System.Web
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq

Public Class CrossTab
	Inherits BaseReport

	Private mstrSQLSelect As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String

	Private mlngCrossTabID As Integer
	Private mblnNoRecords As Boolean

	Private fOK As Boolean
	Private mstrStatusMessage As String

	Private mlngCrossTabType As CrossTabType
	Public mstrTempTableName As String

	Private mstrBaseTable As String
	Private mlngBaseTableID As Integer
	Private rsCrossTabData As DataTable
	Private mblnIntersection As Boolean
	Private mblnPageBreak As Boolean
	Private mblnShowAllPagesTogether As Boolean
	Private mdReportStartDate As DateTime?
	Private mdReportEndDate As DateTime?

	Private mblnShowPercentage As Boolean
	Private mblnPercentageofPage As Boolean
	Private mbUse1000Separator As Boolean
	Private mblnSuppressZeros As Boolean
	Private mlngRecordDescExprID As Integer
	Private mstrPicklistFilter As String
	Private mstrPicklistFilterName As String
	Private mblnChkPicklistFilter As Boolean

	Private mblnOutputScreen As Boolean
	Private mblnOutputPrinter As Boolean
	Private mstrOutputPrinterName As String
	Private mblnOutputSave As Boolean
	Private mlngOutputSaveExisting As Integer
	Private mblnOutputEmail As Boolean
	Private mlngOutputEmailID As Integer
	Private mstrOutputEmailName As String
	Private mstrOutputEmailSubject As String
	Private mstrOutputEmailAttachAs As String
	Private mstrOutputPivotArray() As String

	Private Const HOR As Short = 0 'Horizontal
	Private Const VER As Short = 1 'Vertical
	Private Const PGB As Short = 2 'Page Break
	Private Const INS As Short = 3 'Intersection

	Private Const TYPECOUNT As Short = 0
	Private Const TYPEAVERAGE As Short = 1
	Private Const TYPEMAXIMUM As Short = 2
	Private Const TYPEMINIMUM As Short = 3
	Private Const TYPETOTAL As Short = 4

	Private mvarHeadings(2) As Object
	Private mvarSearches(2) As Object

	Private mdblHorTotal(,,) As Double
	Private mdblVerTotal(,,) As Double
	Private mdblPgbTotal(,,) As Double
	Private mdblPageTotal(,) As Double
	Private mdblGrandTotal() As Double

	Private mdblDataArray(,,,) As Double
	Private mstrOutput() As String

	Private mlngIntersectionDecimals As Integer
	Private mstrIntersectionMask As String
	Private mdblPercentageFactor As Double

	Private mstrType() As String 'e.g. mstrtype(TYPETOTAL) = for example: "Total"
	Private mlngColID(3) As Integer
	Private mstrColName(3) As String 'e.g. mstrColName(INS) = "Salary" (the name of the intersection field)
	Private mlngColDataType(3) As String
	Private mstrFormat(3) As String
	Private mdblMin(2) As Double
	Private mdblMax(2) As Double
	Private mdblStep(2) As Double

	Private mlngType As Integer

	Private msAbsenceBreakdownTypes As String

	Private mvarPrompts(,) As Object

	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String

	Private _descriptionsAsArray As ArrayList
	Private _cellColoursAsArray As ArrayList
	Private _axisLabelsAsArray As ArrayList
	Public ReadOnly Property AxisLabelsAsArray As ArrayList
		Get
			Return _axisLabelsAsArray
		End Get
	End Property

	Public Enum enumNineBoxDescriptionOrColour
		Description
		Colour
	End Enum


	Public PivotData As DataTable

	Public WriteOnly Property CustomReportID() As Integer
		Set(ByVal Value As Integer)
			mlngCrossTabID = Value
		End Set
	End Property

	Public WriteOnly Property FailedMessage() As String
		Set(ByVal value As String)
			Logs.AddDetailEntry(value)
		End Set
	End Property

	Public ReadOnly Property NoRecords() As Boolean
		Get
			Return mblnNoRecords
		End Get
	End Property

	Public WriteOnly Property CrossTabID() As Integer
		Set(ByVal Value As Integer)
			mlngCrossTabID = Value
		End Set
	End Property

	Public ReadOnly Property ErrorString() As String
		Get
			ErrorString = mstrStatusMessage
		End Get
	End Property

	Public Property EventLogID() As Integer
		Get
			EventLogID = Logs.EventLogID
		End Get
		Set(ByVal value As Integer)
			Logs.EventLogID = value
		End Set
	End Property

	Public Property IntersectionType() As Integer
		Get
			IntersectionType = mlngType
		End Get
		Set(ByVal value As Integer)
			mlngType = value
		End Set
	End Property

	Public Property ShowPercentage() As Boolean
		Get
			ShowPercentage = mblnShowPercentage
		End Get
		Set(ByVal value As Boolean)
			mblnShowPercentage = value
		End Set
	End Property

	Public Property PercentageOfPage() As Boolean
		Get
			PercentageOfPage = mblnPercentageofPage
		End Get
		Set(ByVal value As Boolean)
			mblnPercentageofPage = value
		End Set
	End Property

	Public Property SuppressZeros() As Boolean
		Get
			SuppressZeros = mblnSuppressZeros
		End Get
		Set(ByVal value As Boolean)
			mblnSuppressZeros = value
		End Set
	End Property

	Public ReadOnly Property OutputArrayData(lngIndex As Integer) As String
		Get
			Return mstrOutput(lngIndex)
		End Get
	End Property

	Public ReadOnly Property OutputArrayDataUBound() As Integer
		Get
			Return UBound(mstrOutput)
		End Get
	End Property

	Public ReadOnly Property CrossTabName() As String
		Get

			Dim strOutput As String = Name

			If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
				strOutput &= String.Format(" ({0} -> {1})", CDate(mdReportStartDate).ToString(LocaleDateFormat), CDate(mdReportEndDate).ToString(LocaleDateFormat))
			End If

			If mblnChkPicklistFilter Then
				strOutput = strOutput & mstrPicklistFilterName
			End If

			Return strOutput

		End Get
	End Property

	Public ReadOnly Property ColumnHeading(lngArray As Integer, lngIndex As Integer) As String
		Get
			'ColumnHeading = Replace(mvarHeadings(lngArray)(CLng(lngIndex)), Chr(34), Chr(34) & " + String.fromCharCode(34) + " & Chr(34))
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ColumnHeading = mvarHeadings(lngArray)(lngIndex)
		End Get
	End Property

	Public ReadOnly Property ColumnHeadingUbound(lngIndex As Integer) As Integer
		Get
			If Not mvarHeadings(lngIndex) Is Nothing Then
				ColumnHeadingUbound = UBound(mvarHeadings(lngIndex))
			Else
				ColumnHeadingUbound = 0
			End If
		End Get
	End Property

	Public ReadOnly Property PageBreakColumn() As Boolean
		Get
			PageBreakColumn = mblnPageBreak
		End Get
	End Property

	Public ReadOnly Property PageBreakColumnName() As String
		Get
			PageBreakColumnName = IIf(mblnPageBreak, Replace(mstrColName(PGB), "_", " "), "<None>").ToString()
		End Get
	End Property

	Public ReadOnly Property IntersectionColumn() As Boolean
		Get
			IntersectionColumn = mblnIntersection
		End Get
	End Property

	Public ReadOnly Property IntersectionColumnName() As String
		Get
			IntersectionColumnName = IIf(mblnIntersection, Replace(mstrColName(INS), "_", " "), "<None>").ToString()
		End Get
	End Property

	Public ReadOnly Property HorizontalColumnName() As String
		Get

			If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
				HorizontalColumnName = "Day"
			Else
				HorizontalColumnName = Replace(mstrColName(HOR), "_", " ")
			End If

		End Get
	End Property

	Public ReadOnly Property VerticalColumnName() As String
		Get
			VerticalColumnName = Replace(mstrColName(VER), "_", " ")
		End Get
	End Property

	Public ReadOnly Property BaseTableName() As String
		Get
			BaseTableName = mstrBaseTable
		End Get
	End Property

	Public ReadOnly Property RecordDescExprID() As Integer
		Get
			RecordDescExprID = mlngRecordDescExprID
		End Get
	End Property

	' What type of cross tab are we running as
	Public ReadOnly Property CrossTabType() As CrossTabType
		Get
			Return mlngCrossTabType
		End Get
	End Property

	Public ReadOnly Property OutputScreen() As Boolean
		Get
			OutputScreen = mblnOutputScreen
		End Get
	End Property

	Public ReadOnly Property OutputPrinter() As Boolean
		Get
			OutputPrinter = mblnOutputPrinter
		End Get
	End Property

	Public ReadOnly Property OutputPrinterName() As String
		Get
			OutputPrinterName = mstrOutputPrinterName
		End Get
	End Property

	Public ReadOnly Property OutputSave() As Boolean
		Get
			OutputSave = mblnOutputSave
		End Get
	End Property

	Public ReadOnly Property OutputSaveExisting() As Integer
		Get
			OutputSaveExisting = mlngOutputSaveExisting
		End Get
	End Property

	Public ReadOnly Property OutputEmail() As Boolean
		Get
			OutputEmail = mblnOutputEmail
		End Get
	End Property

	Public ReadOnly Property OutputEmailID() As Integer
		Get
			OutputEmailID = mlngOutputEmailID
		End Get
	End Property

	Public ReadOnly Property OutputEmailGroupName() As String
		Get
			OutputEmailGroupName = mstrOutputEmailName
		End Get
	End Property

	Public ReadOnly Property OutputEmailSubject() As String
		Get
			OutputEmailSubject = mstrOutputEmailSubject
		End Get
	End Property

	Public ReadOnly Property OutputEmailAttachAs() As String
		Get
			OutputEmailAttachAs = mstrOutputEmailAttachAs
		End Get
	End Property

	Public ReadOnly Property IntersectionDecimals() As Integer
		Get
			IntersectionDecimals = mlngIntersectionDecimals
		End Get
	End Property

	Public ReadOnly Property Heading(lngCol As Integer, lngRow As Integer) As String
		Get
			'Heading = Replace(mvarHeadings(lngCol)(lngRow), Chr(34), Chr(34) & " + String.fromCharCode(34) + " & Chr(34))
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Return mvarHeadings(lngCol)(lngRow)
		End Get
	End Property

	Public ReadOnly Property DataArrayCols() As Integer
		Get
			Return UBound(Split(mstrOutput(0), vbTab))
		End Get
	End Property

	Public ReadOnly Property DataArrayRows() As Integer
		Get
			Return UBound(mstrOutput)
		End Get
	End Property

	Public ReadOnly Property DataArray(lngCol As Integer, lngRow As Integer) As String
		Get
			Return Split(mstrOutput(lngRow), vbTab)(lngCol)
		End Get
	End Property

	Public Property Use1000Separator() As Boolean
		Get
			Use1000Separator = mbUse1000Separator
		End Get
		Set(ByVal value As Boolean)
			mbUse1000Separator = value
		End Set
	End Property

	Public ReadOnly Property OutputPivotArrayData(lngIndex As Integer) As String
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Return mstrOutputPivotArray(lngIndex)
		End Get
	End Property

	' 9-Box Grid Reports
	Property XAxisLabel() As String
	Property XAxisSubLabel1() As String
	Property XAxisSubLabel2() As String
	Property XAxisSubLabel3() As String
	Property YAxisLabel() As String
	Property YAxisSubLabel1() As String
	Property YAxisSubLabel2() As String
	Property YAxisSubLabel3() As String
	Property Description1() As String
	Property ColorDesc1() As String
	Property Description2() As String
	Property ColorDesc2() As String
	Property Description3() As String
	Property ColorDesc3() As String
	Property Description4() As String
	Property ColorDesc4() As String
	Property Description5() As String
	Property ColorDesc5() As String
	Property Description6() As String
	Property ColorDesc6() As String
	Property Description7() As String
	Property ColorDesc7() As String
	Property Description8() As String
	Property ColorDesc8() As String
	Property Description9() As String
	Property ColorDesc9() As String


	Public Function EventLogAddHeader() As Integer

		' JDM - 05/12/02 - Fault 4840 - Wrong report type in event log
		If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
			Logs.AddHeader(EventLog_Type.eltStandardReport, "Absence Breakdown")
		ElseIf mlngCrossTabType = CrossTabType.ctt9GridBox Then
			Logs.AddHeader(EventLog_Type.elt9GridBox, Name)
		Else
			Logs.AddHeader(EventLog_Type.eltCrossTab, Name)
		End If

		Return Logs.EventLogID
	End Function

	Public Sub EventLogChangeHeaderStatus(lngStatus As EventLog_Status)
		Logs.ChangeHeaderStatus(lngStatus)
	End Sub

	Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.
		Dim iLoop As Integer
		Dim iDataType As Short
		Dim lngComponentID As Integer

		Try

			ReDim mvarPrompts(1, 0)

			If IsArray(pavPromptedValues) Then
				ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

				For iLoop = 0 To UBound(pavPromptedValues, 2)
					' Get the prompt data type.
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarPrompts(0, iLoop) = lngComponentID

						' NB. Locale to server conversions are done on the client.
						Select Case iDataType
							Case 2
								' Numeric.
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
							Case 3
								' Logic.
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
							Case 4
								' Date.
								' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
								' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
								' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
								' THINGS UP.
								'mvarPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
							Case Else
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
						End Select
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarPrompts(0, iLoop) = 0
					End If
				Next iLoop
			End If

		Catch ex As Exception
			mstrStatusMessage = "Error whilst setting prompted values. " & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Public Function RetreiveDefinition() As Boolean

		Dim rsCrossTabDefinition As DataTable
		Dim strSQL As String

		ReDim mastrUDFsRequired(0)

		strSQL = "SELECT ASRSysCrossTab.*, 'TableName' = ASRSysTables.TableName, 'RecordDescExprID' = ASRSysTables.RecordDescExprID, 'IntersectionColName' = ASRSysColumns.ColumnName, " & "'IntersectionDecimals' = ASRSysColumns.Decimals " & "FROM ASRSysCrossTab " & "JOIN ASRSysTables ON ASRSysCrossTab.TableID = ASRSysTables.TableID " & "LEFT OUTER JOIN ASRSysColumns ON ASRSysCrossTab.IntersectionColID = ASRSysColumns.ColumnID " & "WHERE CrossTabID = " & CStr(mlngCrossTabID)

		Try

			rsCrossTabDefinition = DB.GetDataTable(strSQL)
			If rsCrossTabDefinition.Rows.Count = 0 Then
				'UPGRADE_NOTE: Object rsCrossTabDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsCrossTabDefinition = Nothing
				mstrStatusMessage = "This definition has been deleted by another user"
				RetreiveDefinition = False
				Exit Function
			End If

			Dim objRow = rsCrossTabDefinition.Rows(0)
			With rsCrossTabDefinition

				If LCase(CType(objRow("Username"), String)) <> LCase(_login.Username) And CurrentUserAccess(UtilityType.utlCrossTab, mlngCrossTabID) = ACCESS_HIDDEN Then
					'UPGRADE_NOTE: Object rsCrossTabDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsCrossTabDefinition = Nothing
					mstrStatusMessage = "This definition has been made hidden by another user."
					RetreiveDefinition = False
					Exit Function
				End If

				'Set the CrossTab type
				mlngCrossTabType = CInt(objRow("CrossTabType"))

				mlngBaseTableID = CInt(objRow("TableID"))
				mstrBaseTable = objRow("TableName").ToString()
				mlngRecordDescExprID = CInt(objRow("RecordDescExprID"))
				Name = objRow("Name").ToString()
				mblnChkPicklistFilter = CBool(objRow("PrintFilterHeader"))

				mblnShowPercentage = CBool(objRow("Percentage"))
				mblnPercentageofPage = CBool(objRow("PercentageOfPage"))
				mblnSuppressZeros = CBool(objRow("SuppressZeros"))
				mbUse1000Separator = CBool(objRow("ThousandSeparators"))

				OutputPreview = CBool(objRow("OutputPreview"))
				OutputFormat = CType(objRow("OutputFormat"), OutputFormats)
				mblnOutputScreen = CBool(objRow("OutputScreen"))
				mblnOutputPrinter = CBool(objRow("OutputPrinter"))
				mstrOutputPrinterName = objRow("OutputPrinterName").ToString()
				mblnOutputSave = CBool(objRow("OutputSave"))
				mlngOutputSaveExisting = CInt(objRow("OutputSaveExisting"))
				mblnOutputEmail = CBool(objRow("OutputEmail"))
				mlngOutputEmailID = CInt(objRow("OutputEmailAddr"))
				mstrOutputEmailName = GetEmailGroupName(CInt(objRow("OutputEmailAddr")))
				mstrOutputEmailSubject = objRow("OutputEmailSubject").ToString()
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrOutputEmailAttachAs = IIf(IsDBNull(objRow("OutputEmailAttachAs")), vbNullString, objRow("OutputEmailAttachAs"))
				OutputFilename = objRow("OutputFilename").ToString()

				mlngColID(HOR) = CInt(objRow("HorizontalColID"))
				mdblMin(HOR) = Val(objRow("HorizontalStart"))
				mdblMax(HOR) = Val(objRow("HorizontalStop"))
				mdblStep(HOR) = Val(objRow("HorizontalStep"))

				If CrossTabType = CrossTabType.ctt9GridBox Then
					mdblStep(HOR) = Math.Round((mdblMax(HOR) - mdblMin(HOR)) / 3, GetDecimalsSize(mlngColID(HOR)) + 2)
				End If

				mstrFormat(HOR) = GetFormat(mlngColID(HOR))
				mstrColName(HOR) = GetColumnName(mlngColID(HOR))
				mlngColDataType(HOR) = CStr(GetDataType(mlngBaseTableID, mlngColID(HOR)))
				mlngColID(VER) = CInt(objRow("VerticalColID"))
				mdblMin(VER) = Val(objRow("VerticalStart"))
				mdblMax(VER) = Val(objRow("VerticalStop"))
				mdblStep(VER) = Val(objRow("VerticalStep"))

				If CrossTabType = CrossTabType.ctt9GridBox Then
					mdblStep(VER) = Math.Round((mdblMax(VER) - mdblMin(VER)) / 3, GetDecimalsSize(mlngColID(VER)) + 2)
				End If

				mstrColName(VER) = GetColumnName(mlngColID(VER))
				mlngColDataType(VER) = CStr(GetDataType(mlngBaseTableID, mlngColID(VER)))
				mstrFormat(VER) = GetFormat(mlngColID(VER))

				mlngColID(PGB) = CInt(objRow("PageBreakColID"))
				mblnPageBreak = (mlngColID(PGB) > 0)
				If mblnPageBreak Then
					mstrColName(PGB) = GetColumnName(mlngColID(PGB))
					mlngColDataType(PGB) = CStr(GetDataType(mlngBaseTableID, mlngColID(PGB)))
					mstrFormat(PGB) = GetFormat(mlngColID(PGB))
					mdblMin(PGB) = Val(objRow("PageBreakStart"))
					mdblMax(PGB) = Val(objRow("PageBreakStop"))
					mdblStep(PGB) = Val(objRow("PageBreakStep"))
				End If

				mblnIntersection = (CInt(objRow("IntersectionColID")) > 0)
				If mblnIntersection Then
					mlngType = CInt(objRow("IntersectionType"))
					mlngColID(INS) = CInt(objRow("IntersectionColID"))
					mstrColName(INS) = objRow("IntersectionColName").ToString()
					mlngIntersectionDecimals = CInt(objRow("IntersectionDecimals"))
					mstrIntersectionMask = New String("#", 20) & "0"
					If CInt(objRow("IntersectionDecimals")) > 0 Then
						mstrIntersectionMask = mstrIntersectionMask & "." & New String("0", CInt(objRow("IntersectionDecimals")))
					End If
				Else
					mlngType = 0
				End If

				' 9-Box Grid Reports
				XAxisLabel = NullSafeString(objRow("XAxisLabel"))
				XAxisSubLabel1 = NullSafeString(objRow("XAxisSubLabel1"))
				XAxisSubLabel2 = NullSafeString(objRow("XAxisSubLabel2"))
				XAxisSubLabel3 = NullSafeString(objRow("XAxisSubLabel3"))
				YAxisLabel = NullSafeString(objRow("YAxisLabel"))
				YAxisSubLabel1 = NullSafeString(objRow("YAxisSubLabel1"))
				YAxisSubLabel2 = NullSafeString(objRow("YAxisSubLabel2"))
				YAxisSubLabel3 = NullSafeString(objRow("YAxisSubLabel3"))
				Description1 = NullSafeString(objRow("Description1"))
				ColorDesc1 = NullSafeString(objRow("ColorDesc1"))
				Description2 = NullSafeString(objRow("Description2"))
				ColorDesc2 = NullSafeString(objRow("ColorDesc2"))
				Description3 = NullSafeString(objRow("Description3"))
				ColorDesc3 = NullSafeString(objRow("ColorDesc3"))
				Description4 = NullSafeString(objRow("Description4"))
				ColorDesc4 = NullSafeString(objRow("ColorDesc4"))
				Description5 = NullSafeString(objRow("Description5"))
				ColorDesc5 = NullSafeString(objRow("ColorDesc5"))
				Description6 = NullSafeString(objRow("Description6"))
				ColorDesc6 = NullSafeString(objRow("ColorDesc6"))
				Description7 = NullSafeString(objRow("Description7"))
				ColorDesc7 = NullSafeString(objRow("ColorDesc7"))
				Description8 = NullSafeString(objRow("Description8"))
				ColorDesc8 = NullSafeString(objRow("ColorDesc8"))
				Description9 = NullSafeString(objRow("Description9"))
				ColorDesc9 = NullSafeString(objRow("ColorDesc9"))

				fOK = IsRecordSelectionValid(CInt(objRow("PickListID")), CInt(objRow("FilterID")))
				If fOK = False Then
					mblnNoRecords = True
					Exit Function
				End If

				mstrPicklistFilter = GetPicklistFilterSelect(CInt(objRow("PickListID")), CInt(objRow("FilterID")))
				If fOK = False Then
					mblnNoRecords = True
					Exit Function
				End If

			End With


		Catch ex As Exception
			mstrStatusMessage = "Error reading Cross Tab definition"
			mblnNoRecords = True
			fOK = False

		End Try

		Return fOK

	End Function

	Private Function IsRecordSelectionValid(lngPicklistID As Integer, lngFilterID As Integer) As Boolean

		Dim iResult As RecordSelectionValidityCodes
		Dim fCurrentUserIsSysSecMgr As Boolean

		fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

		' Filter
		If lngFilterID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionType.Filter, lngFilterID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrStatusMessage = "The base table filter used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not fCurrentUserIsSysSecMgr Then
						mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
					End If
			End Select
		ElseIf lngPicklistID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionType.Picklist, lngPicklistID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrStatusMessage = "The base table picklist used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not fCurrentUserIsSysSecMgr Then
						mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
					End If
			End Select
		End If

		IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)

	End Function

	Private Function GetPicklistFilterSelect(lngPicklistID As Integer, lngFilterID As Integer) As String

		Dim rsTemp As DataTable

		If lngPicklistID > 0 Then

			mstrStatusMessage = IsPicklistValid(lngPicklistID)
			If mstrStatusMessage <> vbNullString Then
				fOK = False
				Return False
			End If

			'Get List of IDs from Picklist
			rsTemp = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & lngPicklistID)
			fOK = (rsTemp.Rows.Count > 0)

			If Not fOK Then
				mstrStatusMessage = "The base table picklist contains no records."
			Else
				GetPicklistFilterSelect = vbNullString
				For Each objRow As DataRow In rsTemp.Rows
					GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "").ToString() & objRow(0).ToString()
				Next
			End If

			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing

			'MH20020704 Fault 4022
			rsTemp = DB.GetDataTable("SELECT name from ASRSysPicklistName WHERE PicklistID = " & CStr(lngPicklistID))
			mstrPicklistFilterName = " (Base Table Picklist : " & rsTemp.Rows(0)("Name").ToString() & ")"
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing

		ElseIf lngFilterID > 0 Then

			mstrStatusMessage = IsFilterExist(lngFilterID)
			If mstrStatusMessage <> vbNullString Then
				'mblnInvalidPicklistFilter = True
				fOK = False
				Return False
			End If

			'Get list of IDs from Filter
			fOK = FilteredIDs(lngFilterID, GetPicklistFilterSelect, mastrUDFsRequired, mvarPrompts)

			If Not fOK Then
				' Permission denied on something in the filter.
				mstrStatusMessage = "You do not have permission to use the '" & General.GetFilterName(lngFilterID) & "' filter."
			End If

			'MH20020704 Fault 4022
			rsTemp = DB.GetDataTable("SELECT Name from ASRSysExpressions WHERE ExprID = " & CStr(lngFilterID))
			mstrPicklistFilterName = " (Base Table Filter : " & rsTemp.Rows(0)("Name").ToString() & ")"

			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing

		Else
			mstrPicklistFilterName = " (No Picklist Or Filter Selected)"

		End If

	End Function

	Private Function GetFormat(lngColumnID As Integer) As String

		Dim objColumn = Columns.GetById(lngColumnID)
		Dim sFormat As String = ""

		Select Case objColumn.DataType
			Case ColumnDataType.sqlNumeric
				sFormat = New String("#", objColumn.Size - 1) & "0"
				If objColumn.Decimals > 0 Then
					sFormat = sFormat & "." & New String("0", objColumn.Decimals)
				End If

			Case ColumnDataType.sqlInteger
				sFormat = New String("#", 9) & "0"

		End Select

		Return sFormat

	End Function

	Public Function CreateTempTable() As Boolean

		Dim strColumn(,) As String
		Dim strSQL As String
		Dim lngMax As Integer

		Try

			lngMax = 2
			ReDim strColumn(2, lngMax)

			strColumn(1, 0) = "ID"
			strColumn(2, 0) = "ID"

			strColumn(1, 1) = mstrColName(HOR)
			strColumn(2, 1) = "Hor"

			strColumn(1, 2) = mstrColName(VER)
			strColumn(2, 2) = "Ver"

			If mblnPageBreak Then
				lngMax = lngMax + 1
				ReDim Preserve strColumn(2, lngMax)

				strColumn(1, lngMax) = mstrColName(PGB)
				strColumn(2, lngMax) = "Pgb"
			End If

			If mblnIntersection Then
				lngMax = lngMax + 1
				ReDim Preserve strColumn(2, lngMax)

				strColumn(1, lngMax) = mstrColName(INS)
				strColumn(2, lngMax) = "Ins"
			End If

			If mlngCrossTabType = CrossTabType.cttAbsenceBreakdown Then

				lngMax = lngMax + 7
				ReDim Preserve strColumn(2, lngMax)

				strColumn(1, lngMax) = AbsenceModule.gsAbsenceDurationColumnName
				strColumn(2, lngMax) = "Value"

				strColumn(1, lngMax - 4) = AbsenceModule.gsAbsenceStartDateColumnName
				strColumn(2, lngMax - 4) = "Start_Date"

				strColumn(1, lngMax - 3) = AbsenceModule.gsAbsenceStartSessionColumnName
				strColumn(2, lngMax - 3) = "Start_Session"

				strColumn(1, lngMax - 2) = AbsenceModule.gsAbsenceEndDateColumnName
				strColumn(2, lngMax - 2) = "End_Date"

				strColumn(1, lngMax - 1) = AbsenceModule.gsAbsenceEndSessionColumnName
				strColumn(2, lngMax - 1) = "End_Session"

				strColumn(1, lngMax - 5) = "ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID))
				strColumn(2, lngMax - 5) = "Personnel_ID"

				strColumn(1, lngMax - 6) = AbsenceModule.gsAbsenceDurationColumnName ' Used to hold the day number (1=Mon, 2=Tues etc.)
				strColumn(2, lngMax - 6) = "Day_Number"
			End If

			fOK = True
			GetSQL2(strColumn)
			If fOK = False Then
				mblnNoRecords = True
				Return False
			End If

			mstrTempTableName = General.UniqueSQLObjectName("ASRSysTempCrossTab", 3)
			mstrSQLSelect = mstrSQLSelect & ", " & "space(255) as 'RecDesc' INTO " & mstrTempTableName

			strSQL = "SELECT " & mstrSQLSelect & vbNewLine & " FROM " & mstrSQLFrom & vbNewLine & mstrSQLJoin & vbNewLine & mstrSQLWhere

			'MH20010327 Seems that it might be moving on pass this line of code too
			'quickly so I've tried returning the number of rows effected to make
			'sure that it completes fully
			DB.ExecuteSql(strSQL)

			rsCrossTabData = DB.GetFromSP("spASRIntGetCrosstabReportData" _
				, New SqlParameter("@tablename", SqlDbType.Char, 30) With {.Value = mstrTempTableName} _
				, New SqlParameter("@recordDescid", SqlDbType.Int) With {.Value = mlngRecordDescExprID} _
				, New SqlParameter("@isAbsenceBreakdown", SqlDbType.Int) With {.Value = (mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown)} _
				, New SqlParameter("@startDate", SqlDbType.DateTime) With {.IsNullable = True, .Value = mdReportStartDate} _
				, New SqlParameter("@endDate", SqlDbType.DateTime) With {.IsNullable = True, .Value = mdReportEndDate})

			If rsCrossTabData.Rows.Count = 0 Then
				mstrStatusMessage = "No records meet the selection criteria."
				Logs.AddDetailEntry("Completed successfully. " & mstrStatusMessage)
				Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
				mblnNoRecords = True
				fOK = False
			End If

		Catch ex As Exception
			mstrStatusMessage = ex.Message
			Return False
		End Try

		Return fOK

	End Function

	Private Sub GetSQL2(ByRef strCol(,) As String)

		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim sSource As String
		Dim lngCount As Integer
		Dim fColumnOK As Boolean
		Dim alngTableViews(,) As Integer
		Dim iNextIndex As Integer
		Dim fFound As Boolean

		Dim sCaseStatement As String
		Dim strSelectedRecords As String
		Dim sWhereIDs As String
		Dim strColumn As String
		Dim blnCharColumn As Boolean


		Try

			fOK = True
			ReDim alngTableViews(2, 0)

			mstrSQLFrom = gcoTablePrivileges.Item(mstrBaseTable).RealSource
			mstrSQLSelect = vbNullString
			mstrSQLJoin = vbNullString
			Dim asViews(0) As Object

			blnCharColumn = (Val(mlngColDataType(lngCount)) = ColumnDataType.sqlVarChar)


			For lngCount = 0 To UBound(strCol, 2)

				objColumnPrivileges = GetColumnPrivileges(mstrBaseTable)
				fColumnOK = objColumnPrivileges.IsValid(strCol(1, lngCount))
				If fColumnOK Then
					fColumnOK = objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect

					If fColumnOK Then
						fColumnOK = gcoTablePrivileges.Item(mstrBaseTable).AllowSelect
					End If

				End If

				'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objColumnPrivileges = Nothing

				If lngCount <= UBound(mlngColDataType) Then
					blnCharColumn = (Val(mlngColDataType(lngCount)) = ColumnDataType.sqlVarChar)
				End If

				If fColumnOK Then
					' The column can be read from the base table/view, or directly from a parent table.
					' Add the column to the column list.

					If strSelectedRecords = vbNullString And mstrPicklistFilter <> vbNullString Then

						If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
							strSelectedRecords = mstrSQLFrom & ".ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID)) & " IN (" & mstrPicklistFilter & ")"
						Else
							strSelectedRecords = mstrSQLFrom & ".ID IN (" & mstrPicklistFilter & ")"
						End If

					End If

					strColumn = mstrSQLFrom & "." & strCol(1, lngCount)
					If blnCharColumn Then
						strColumn = FormatSQLColumn(strColumn)
					End If

					mstrSQLSelect = mstrSQLSelect & IIf(Len(mstrSQLSelect) > 0, ", ", "").ToString() & strColumn & " AS '" & strCol(2, lngCount) & "'"

				Else

					ReDim asViews(0)
					For Each objTableView In gcoTablePrivileges.Collection

						'Loop thru all of the views for this table where the user has select access
						If (Not objTableView.IsTable) And (objTableView.TableID = mlngBaseTableID) And (objTableView.AllowSelect) Then

							sSource = objTableView.ViewName

							' Get the column permission for the view.
							objColumnPrivileges = GetColumnPrivileges(sSource)

							If objColumnPrivileges.IsValid(strCol(1, lngCount)) Then
								If objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect Then
									' Add the view info to an array to be put into the column list or order code below.
									iNextIndex = UBound(asViews) + 1
									ReDim Preserve asViews(iNextIndex)
									'UPGRADE_WARNING: Couldn't resolve default property of object asViews(iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									asViews(iNextIndex) = sSource


									'=== This is the join code section ===
									' Add the view to the Join code.
									' Check if the view has already been added to the join code.
									fFound = False
									For iNextIndex = 1 To UBound(alngTableViews, 2)
										If alngTableViews(2, iNextIndex) = objTableView.ViewID Then
											fFound = True
											Exit For
										End If
									Next iNextIndex

									If Not fFound Then
										' The view has not yet been added to the join code, so add it to the array and the join code.
										' (also include the picklist info)

										iNextIndex = UBound(alngTableViews, 2) + 1
										ReDim Preserve alngTableViews(2, iNextIndex)
										alngTableViews(1, iNextIndex) = 1
										alngTableViews(2, iNextIndex) = objTableView.ViewID

										mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"

										sWhereIDs = sWhereIDs & IIf(sWhereIDs <> vbNullString, " OR ", vbNullString) & mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"

										'If mstrPicklistFilter <> vbNullString Then
										strSelectedRecords = strSelectedRecords & IIf(strSelectedRecords <> vbNullString, " OR ", vbNullString) & "(" & IIf(mstrPicklistFilter <> vbNullString, sSource & ".ID IN (" & mstrPicklistFilter & ") AND ", vbNullString) & sSource & ".ID > 0)"
										'End If

									End If
								End If
								'=== End of Join Code ===


								'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objColumnPrivileges = Nothing
							End If

						End If
					Next objTableView
					'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objTableView = Nothing

					' The current user does have permission to 'read' the column through a/some view(s) on the
					' table.
					If UBound(asViews) = 0 Then
						fOK = False
						'MH20010716 Fault 2497
						'If its the ID column they they don't have any access to the table.
						'mstrStatusMessage = "You do not have permission to see the column '" & strCol(1, lngCount) & "' " & _
						'"either directly or through any views." & vbNewLine
						mstrStatusMessage = "You do not have permission to see the " & IIf(strCol(1, lngCount) = "ID", "table '" & mstrBaseTable, "column '" & strCol(1, lngCount)) & "' either directly or through any views." & vbNewLine
						Exit Sub
					Else

						sCaseStatement = ""
						For iNextIndex = 1 To UBound(asViews)
							'UPGRADE_WARNING: Couldn't resolve default property of object asViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCaseStatement = sCaseStatement & IIf(sCaseStatement <> "", vbCrLf & " , ", "").ToString() & asViews(iNextIndex).ToString() & "." & strCol(1, lngCount)
						Next iNextIndex

						If Len(sCaseStatement) > 0 Then
							strColumn = "COALESCE(" & sCaseStatement & ", NULL)"

							If blnCharColumn Then
								strColumn = FormatSQLColumn(strColumn)
							End If

							mstrSQLSelect = mstrSQLSelect & IIf(Len(mstrSQLSelect) > 0, ", ", "").ToString() & vbCrLf & strColumn & "AS '" & strCol(2, lngCount) & "'"
						End If

					End If
				End If
			Next

			If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown And Not msAbsenceBreakdownTypes = vbNullString Then
				mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "(UPPER(" & AbsenceModule.gsAbsenceTypeColumnName & ") IN " & msAbsenceBreakdownTypes & ")"
			End If

			If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
				Dim sReportStartDate As String = CDate(mdReportStartDate).ToString(SQLDateFormat)
				Dim sReportEndDate As String = CDate(mdReportEndDate).ToString(SQLDateFormat)
				mstrSQLWhere &= IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "( " & AbsenceModule.gsAbsenceStartDateColumnName & " <= CONVERT(datetime, '" & sReportEndDate & "'))" & "And (" & AbsenceModule.gsAbsenceEndDateColumnName & " >= CONVERT(datetime, '" & sReportStartDate & "') OR " & AbsenceModule.gsAbsenceEndDateColumnName & " IS NULL)"
			End If

			If strSelectedRecords <> vbNullString Then
				mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "(" & strSelectedRecords & ")"
			End If

		Catch ex As Exception
			mstrStatusMessage = "Error retrieving data"
			fOK = False

		End Try

	End Sub

	Public Function GetHeadingsAndSearches() As Boolean

		Dim strHeading() As String
		Dim strSearch() As String
		Dim lngLoop As Integer

		Try

			For lngLoop = 0 To 2

				ReDim strHeading(0)
				ReDim strSearch(0)

				If lngLoop = 2 And mblnPageBreak = False Then
					'When no page break field is specified
					strHeading(0) = "<None>"

				Else
					GetHeadingsAndSearchesForColumns(lngLoop, strHeading, strSearch)

				End If

				'Store each array in an array of variants (an array in an array!)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarHeadings(lngLoop) = VB6.CopyArray(strHeading)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSearches(lngLoop) = VB6.CopyArray(strSearch)

			Next

		Catch ex As Exception
			mstrStatusMessage = "Error building headings and search arrays"
			Throw

		End Try

		Return True

	End Function

	Private Sub GetHeadingsAndSearchesForColumns(lngLoop As Integer, ByRef strHeading() As String, ByRef strSearch() As String)

		Dim strFieldValue As String
		Dim lngCount As Integer
		Dim dblGroup As Double
		Dim dblGroupMax As Double
		Dim dblUnit As Double
		Dim strColumnName As String
		Dim strWhereEmpty As String
		Dim strOrder As String
		Dim sColumnNames() As String

		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strColumnName = Choose(lngLoop + 1, "Hor", "Ver", "Pgb")

		If mdblMin(lngLoop) = 0 And mdblMax(lngLoop) = 0 Then

			lngCount = 0

			strWhereEmpty = strColumnName & " IS NULL"
			If mlngColDataType(lngLoop) <> CStr(ColumnDataType.sqlNumeric) And mlngColDataType(lngLoop) <> CStr(ColumnDataType.sqlInteger) And mlngColDataType(lngLoop) <> CStr(ColumnDataType.sqlBoolean) Then
				strWhereEmpty = strWhereEmpty & " OR Trim(" & strColumnName & ") = ''"
			End If

			' Don't put in empty clauses if we're running an absence breakdown
			If mlngCrossTabType <> Enums.CrossTabType.cttAbsenceBreakdown Then
				ReDim Preserve strHeading(lngCount)
				ReDim Preserve strSearch(lngCount)
				strHeading(lngCount) = "<Empty>"
				strSearch(lngCount) = strWhereEmpty
				lngCount += 1
			End If

			If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown And strColumnName = "Hor" Then
				sColumnNames = {FormatSQLColumn(strColumnName), "Day_Number", "DisplayOrder"}
				strOrder = "DisplayOrder"
			Else
				sColumnNames = {FormatSQLColumn(strColumnName)}
				strOrder = strColumnName
			End If

			For Each objRow As DataRow In rsCrossTabData.DefaultView.ToTable(True, sColumnNames).Select("", strOrder)

				'MH20010213 Had to make this change so that working pattern would work

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				strFieldValue = IIf(IsDBNull(objRow(0)), vbNullString, objRow(0))

				If Trim(strFieldValue) <> vbNullString Then
					ReDim Preserve strHeading(lngCount)
					ReDim Preserve strSearch(lngCount)

					Select Case mlngColDataType(lngLoop)
						Case CStr(ColumnDataType.sqlDate)
							strHeading(lngCount) = VB6.Format(objRow(0), LocaleDateFormat)
							strSearch(lngCount) = strColumnName & " = '" & CDate(objRow(0)).ToString(LocaleDateFormat) & "'"

						Case CStr(ColumnDataType.sqlBoolean)
							strHeading(lngCount) = IIf(objRow(0), "True", "False")
							strSearch(lngCount) = strColumnName & " = " & IIf(objRow(0), "1", "0")

						Case CStr(ColumnDataType.sqlNumeric), CStr(ColumnDataType.sqlInteger)
							strHeading(lngCount) = ConvertNumberForDisplay(VB6.Format(objRow(0), mstrFormat(lngLoop)))
							strSearch(lngCount) = strColumnName & " = " & ConvertNumberForSQL(objRow(0))

						Case Else
							strHeading(lngCount) = HttpUtility.HtmlEncode(objRow(0).ToString)
							strSearch(lngCount) = FormatSQLColumn(strColumnName) & " = '" & Replace(strFieldValue, "'", "''") & "'"

					End Select

					lngCount += 1

				End If

			Next

		Else

			ReDim Preserve strHeading(1)
			ReDim Preserve strSearch(1)

			'First element of range for null values...
			strHeading(0) = "<Empty>"
			strSearch(0) = strColumnName & " IS NULL"

			'Second element of range for those less than minimum value of range...
			strHeading(1) = "< " & ConvertNumberForDisplay(VB6.Format(mdblMin(lngLoop), mstrFormat(lngLoop)))
			strSearch(1) = strColumnName & " < " & ConvertNumberForSQL(CStr(mdblMin(lngLoop)))

			dblUnit = GetSmallestUnit(lngLoop)

			lngCount = 2

			If mlngCrossTabType = CrossTabType.ctt9GridBox Then

				Dim stepValue = Math.Round((mdblMax(lngLoop) - mdblMin(lngLoop)) / 3, 3)
				Dim lastValue = mdblMin(lngLoop)

				For iLoop2 = 0 To 2

					ReDim Preserve strHeading(lngCount)
					ReDim Preserve strSearch(lngCount)

					Dim dblFrom = lastValue
					Dim dblTo As Double

					dblTo = Math.Min(Math.Round(lastValue + stepValue, 2), mdblMax(lngLoop))

					strHeading(lngCount) = String.Format("{0} - {1}", dblFrom.ToString(CultureInfo.InvariantCulture), dblTo.ToString(CultureInfo.InvariantCulture))
					strSearch(lngCount) = String.Format("{0} >= {1} AND {0} <= {2}", strColumnName, dblFrom.ToString(CultureInfo.InvariantCulture), dblTo.ToString(CultureInfo.InvariantCulture))

					lastValue = dblTo + 0.01
					lngCount += 1

				Next

				' Because the Nine Box Grid is upside down swap rows 3 and 5 around
				If lngLoop = 1 Then
					Dim swapNineBox As String
					swapNineBox = strHeading(2)
					strHeading(2) = strHeading(4)
					strHeading(4) = swapNineBox

					swapNineBox = strSearch(2)
					strSearch(2) = strSearch(4)
					strSearch(4) = swapNineBox
				End If

			Else

				If mdblStep(lngLoop) = 0 Then
					mstrStatusMessage = "Step value for " & strColumnName & " column cannot be zero"
					fOK = False
					Exit Sub
				End If

				For dblGroup = mdblMin(lngLoop) To mdblMax(lngLoop) Step mdblStep(lngLoop)
					ReDim Preserve strHeading(lngCount)
					ReDim Preserve strSearch(lngCount)
					dblGroupMax = dblGroup + mdblStep(lngLoop) - dblUnit
					strHeading(lngCount) = ConvertNumberForDisplay(VB6.Format(dblGroup, mstrFormat(lngLoop))) & IIf(dblGroupMax <> dblGroup, " - " & ConvertNumberForDisplay(Format(dblGroupMax, mstrFormat(lngLoop))), "")
					strSearch(lngCount) = String.Format("{0} >= {1} AND {0} <= {2}", strColumnName, ConvertNumberForSQL(CStr(dblGroup)), ConvertNumberForSQL(CStr(dblGroupMax)))
					lngCount += 1
				Next

			End If

			ReDim Preserve strHeading(lngCount)
			ReDim Preserve strSearch(lngCount)
			'Last element of range for those more than maximum value of range...
			strHeading(lngCount) = "> " & ConvertNumberForDisplay(VB6.Format(dblGroup - dblUnit, mstrFormat(lngLoop)))
			strSearch(lngCount) = strColumnName & " > " & ConvertNumberForSQL(CStr(dblGroup - dblUnit))

		End If

	End Sub

	Private Function GetSmallestUnit(lngLoop As Integer) As Double

		'e.g. mstrFormat(lngLoop) = 0.0,   GetSmallestUnit = 0.1
		'     mstrFormat(lngLoop) = 0.000, GetSmallestUnit = 0.001
		'     mstrFormat(lngLoop) = #0,    GetSmallestUnit = 1
		'     mstrFormat(lngLoop) = #0.00, GetSmallestUnit = 0.01

		Dim strTemp As String
		Dim intFound As Integer

		intFound = InStr(mstrFormat(lngLoop), ".")
		If intFound > 0 Then
			strTemp = Mid(mstrFormat(lngLoop), intFound, Len(mstrFormat(lngLoop)) - intFound) & "1"
			GetSmallestUnit = CDbl(ConvertNumberForDisplay(strTemp))
		Else
			GetSmallestUnit = 1
		End If

		'	Return 1.0

	End Function

	Public Function IntersectionTypeValue(index As Integer) As String
		Return mstrType(index)
	End Function

	Public Function BuildTypeArray() As Boolean

		If mblnIntersection Then

			ReDim mstrType(4)
			mstrType(TYPECOUNT) = "Count"
			mstrType(TYPEAVERAGE) = "Average"
			mstrType(TYPEMAXIMUM) = "Maximum"
			mstrType(TYPEMINIMUM) = "Minimum"
			mstrType(TYPETOTAL) = "Total"

		Else

			ReDim mstrType(0)
			mstrType(TYPECOUNT) = "Count"

		End If

		Return True


	End Function

	Public Function BuildDataArrays() As Boolean

		Dim strTempValue As String
		Dim dblThisIntersectionVal As Double

		Dim lngCol As Integer
		Dim lngRow As Integer
		Dim lngPage As Integer
		Dim lngNumCols As Integer
		Dim lngNumRows As Integer
		Dim lngNumPages As Integer

		Try

			lngNumCols = UBound(mvarHeadings(0))
			lngNumRows = UBound(mvarHeadings(1))
			lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

			ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4)
			ReDim mdblHorTotal(lngNumCols, lngNumPages, 4)
			ReDim mdblVerTotal(lngNumRows, lngNumPages, 4)
			ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4)	'+1 for totals !
			ReDim mdblPageTotal(lngNumPages, 4)
			ReDim mdblGrandTotal(4)

			For Each objRow In rsCrossTabData.Rows

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				strTempValue = IIf(Not IsDBNull(objRow("HOR")), objRow("HOR"), vbNullString)
				lngCol = GetGroupNumber(strTempValue, HOR)

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				strTempValue = IIf(Not IsDBNull(objRow("VER")), objRow("VER"), vbNullString)
				lngRow = GetGroupNumber(strTempValue, VER)

				If mblnPageBreak Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					strTempValue = IIf(Not IsDBNull(objRow("PGB")), objRow("PGB"), vbNullString)
					lngPage = GetGroupNumber(strTempValue, PGB)
				Else
					lngPage = 0
				End If

				'Count
				mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) + 1
				mdblHorTotal(lngCol, lngPage, TYPECOUNT) = mdblHorTotal(lngCol, lngPage, TYPECOUNT) + 1
				mdblVerTotal(lngRow, lngPage, TYPECOUNT) = mdblVerTotal(lngRow, lngPage, TYPECOUNT) + 1
				mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = mdblPgbTotal(lngCol, lngRow, TYPECOUNT) + 1
				mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) + 1
				mdblPageTotal(lngPage, TYPECOUNT) = mdblPageTotal(lngPage, TYPECOUNT) + 1
				mdblGrandTotal(TYPECOUNT) = mdblGrandTotal(TYPECOUNT) + 1

				'If mblnIntersection And IsNull(objRow(.Fields.Count - 1)) = False Then
				If mblnIntersection Then

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("INS")) Then
						dblThisIntersectionVal = 0
					Else
						dblThisIntersectionVal = Val(ConvertNumberForSQL(objRow("INS")))
					End If


					'Total
					mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
					mdblHorTotal(lngCol, lngPage, TYPETOTAL) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) + dblThisIntersectionVal
					mdblVerTotal(lngRow, lngPage, TYPETOTAL) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
					mdblPgbTotal(lngCol, lngRow, TYPETOTAL) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) + dblThisIntersectionVal
					mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) + dblThisIntersectionVal
					mdblPageTotal(lngPage, TYPETOTAL) = mdblPageTotal(lngPage, TYPETOTAL) + dblThisIntersectionVal
					mdblGrandTotal(TYPETOTAL) = mdblGrandTotal(TYPETOTAL) + dblThisIntersectionVal

					'Average
					mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) / mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT)
					mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) / mdblHorTotal(lngCol, lngPage, TYPECOUNT)
					mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) / mdblVerTotal(lngRow, lngPage, TYPECOUNT)
					mdblPgbTotal(lngCol, lngRow, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) / mdblPgbTotal(lngCol, lngRow, TYPECOUNT)
					mdblPgbTotal(lngCol, lngNumRows + 1, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) / mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT)
					mdblPageTotal(lngPage, TYPEAVERAGE) = mdblPageTotal(lngPage, TYPETOTAL) / mdblPageTotal(lngPage, TYPECOUNT)
					mdblGrandTotal(TYPEAVERAGE) = mdblGrandTotal(TYPETOTAL) / mdblGrandTotal(TYPECOUNT)

					'Minimum
					If dblThisIntersectionVal < mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
						mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
						If dblThisIntersectionVal < mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
							mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal < mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
							mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
							mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
							mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal < mdblPageTotal(lngPage, TYPEMINIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
							mdblPageTotal(lngPage, TYPEMINIMUM) = dblThisIntersectionVal
							If dblThisIntersectionVal < mdblGrandTotal(TYPEMINIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
								mdblGrandTotal(TYPEMINIMUM) = dblThisIntersectionVal
							End If
						End If
					End If

					'Maximum
					If dblThisIntersectionVal > mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
						mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
						If dblThisIntersectionVal > mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
							mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal > mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
							mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
							mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
							mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) = dblThisIntersectionVal
						End If
						If dblThisIntersectionVal > mdblPageTotal(lngPage, TYPEMAXIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
							mdblPageTotal(lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
							If dblThisIntersectionVal > mdblGrandTotal(TYPEMAXIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
								mdblGrandTotal(TYPEMAXIMUM) = dblThisIntersectionVal
							End If
						End If
					End If
				End If

			Next

			Dim swapNineBox As Double
			If mlngCrossTabType = CrossTabType.ctt9GridBox Then
				For lngPage = 0 To ColumnHeadingUbound(2)
					For lngCol = 0 To UBound(mvarHeadings(HOR))
						swapNineBox = mdblDataArray(lngCol, 2, lngPage, TYPECOUNT)
						mdblDataArray(lngCol, 2, lngPage, TYPECOUNT) = mdblDataArray(lngCol, 4, lngPage, TYPECOUNT)
						mdblDataArray(lngCol, 4, lngPage, TYPECOUNT) = swapNineBox
					Next
				Next
			End If


			'Build an array with the nine box grid cells descriptions so we can map cells to their descriptions
			_descriptionsAsArray = New ArrayList
			_descriptionsAsArray.Add(Description1)
			_descriptionsAsArray.Add(Description2)
			_descriptionsAsArray.Add(Description3)
			_descriptionsAsArray.Add(Description4)
			_descriptionsAsArray.Add(Description5)
			_descriptionsAsArray.Add(Description6)
			_descriptionsAsArray.Add(Description7)
			_descriptionsAsArray.Add(Description8)
			_descriptionsAsArray.Add(Description9)

			'Build an array with the nine box grid cells colours so we can map cells to their colours
			_cellColoursAsArray = New ArrayList
			_cellColoursAsArray.Add(ColorDesc1)
			_cellColoursAsArray.Add(ColorDesc2)
			_cellColoursAsArray.Add(ColorDesc3)
			_cellColoursAsArray.Add(ColorDesc4)
			_cellColoursAsArray.Add(ColorDesc5)
			_cellColoursAsArray.Add(ColorDesc6)
			_cellColoursAsArray.Add(ColorDesc7)
			_cellColoursAsArray.Add(ColorDesc8)
			_cellColoursAsArray.Add(ColorDesc9)

			'Build an array with the axis labels
			_axisLabelsAsArray = New ArrayList
			_axisLabelsAsArray.Add(XAxisLabel)
			_axisLabelsAsArray.Add(XAxisSubLabel1)
			_axisLabelsAsArray.Add(XAxisSubLabel2)
			_axisLabelsAsArray.Add(XAxisSubLabel3)
			_axisLabelsAsArray.Add(YAxisLabel)
			_axisLabelsAsArray.Add(YAxisSubLabel1)
			_axisLabelsAsArray.Add(YAxisSubLabel2)
			_axisLabelsAsArray.Add(YAxisSubLabel3)

		Catch ex As Exception
			mstrStatusMessage = "Error processing data"
			Return False

		End Try

		Return True

	End Function

	Public Function GetGroupNumber(strValue As String, Index As Integer) As Integer

		Dim dblValue As Double
		Dim lngCount As Integer
		Dim dblLoop As Double

		Try

			'Non range column
			If mdblMin(Index) = 0 And mdblMax(Index) = 0 Then

				For lngCount = 0 To UBound(mvarHeadings(Index))

					Select Case mlngColDataType(Index)
						Case CStr(ColumnDataType.sqlDate)

							If Not strValue Is Nothing Then
								If mvarHeadings(Index)(lngCount) = CDate(strValue).ToString(LocaleDateFormat) Then
									Return lngCount
								End If
							End If

						Case CStr(ColumnDataType.sqlNumeric), CStr(ColumnDataType.sqlInteger)
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If UCase(mvarHeadings(Index)(lngCount)) = ConvertNumberForDisplay(VB6.Format(strValue, mstrFormat(Index))) Then
								Return lngCount
							End If

						Case CStr(ColumnDataType.sqlBoolean)
							Select Case LCase(strValue)
								Case ""
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If mvarHeadings(Index)(lngCount) = "<Empty>" Then
										Return lngCount
									End If
								Case "false", "0"
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If mvarHeadings(Index)(lngCount) = "False" Then
										'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										Return lngCount
									End If
								Case Else
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If mvarHeadings(Index)(lngCount) = "True" Then
										Return lngCount
									End If
							End Select

						Case Else

							'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Not strValue Is Nothing Then
								If LCase(mvarHeadings(Index)(lngCount)).ToString().Trim() = LCase(HttpUtility.HtmlEncode(strValue.Trim())) Then
									Return lngCount
								End If
							End If
					End Select

				Next

			Else 'Numeric ranges...

				lngCount = 1
				dblValue = Val(strValue)

				If mlngCrossTabType = CrossTabType.ctt9GridBox Then

					Dim stepValue = Math.Round((mdblMax(Index) - mdblMin(Index)) / 3, 3)
					Dim lastValue2 = mdblMin(Index)

					For iLoop2 = 0 To 2

						Dim dblFrom = lastValue2
						Dim dblTo As Double

						dblTo = Math.Min(Math.Round(lastValue2 + stepValue, 2), mdblMax(Index))

						lastValue2 = dblTo + 0.01
						lngCount += 1

						If dblValue >= dblFrom AndAlso dblValue <= dblTo Then
							Return lngCount
						End If

					Next

				Else


					If strValue = vbNullString Then
						Return 0
					ElseIf dblValue < mdblMin(Index) AndAlso Not mlngCrossTabType = CrossTabType.ctt9GridBox Then
						Return 1
					End If


					For dblLoop = mdblMin(Index) To mdblMax(Index) Step mdblStep(Index)
						lngCount += 1

						Dim dblMin = Math.Min(dblLoop, dblLoop + mdblStep(Index))
						Dim dblMax = Math.Max(dblLoop, dblLoop + mdblStep(Index))

						If dblValue >= dblMin AndAlso dblValue < dblMax Then
							Return lngCount
						End If
					Next
					Return lngCount + 1
				End If

			End If

		Catch ex As Exception
			Return 0

		End Try

		Return 0

	End Function

	Public Sub BuildOutputStrings(lngSinglePage As Integer)

		Const strDelim As String = vbTab
		Dim strTempDelim As String

		Dim lngNumCols As Integer
		Dim lngNumRows As Integer
		Dim lngNumPages As Integer

		Dim lngCol As Integer
		Dim lngRow As Integer
		Dim lngPage As Integer
		Dim lngTYPE As Integer

		Dim sngAverage As Single
		Dim iAverageColumn As Integer

		Try

			lngNumCols = UBound(mvarHeadings(HOR))
			lngNumRows = UBound(mvarHeadings(VER))
			lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)), 0)
			iAverageColumn = lngNumCols - 1

			' JDM - 22/06/01 - Fault 2476 - Display totals instead
			If mlngCrossTabType <> Enums.CrossTabType.cttAbsenceBreakdown Then
				lngTYPE = mlngType
			Else
				lngTYPE = TYPETOTAL
			End If

			'mdblPercentageFactor will be used in FORMATCELL, if required
			GetPercentageFactor(lngSinglePage, lngTYPE)

			ReDim mstrOutput(lngNumRows + 2)

			'Add First Column details (Vertical headings)
			mstrOutput(0) = strDelim & mstrOutput(0)
			For lngRow = 0 To lngNumRows
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mstrOutput(lngRow + 1) = Trim(mvarHeadings(VER)(lngRow)) & strDelim & mstrOutput(lngRow + 1)
			Next
			mstrOutput(lngNumRows + 2) = IIf(mlngCrossTabType = Enums.CrossTabType.cttNormal, mstrType(mlngType), "Total") & strDelim & mstrOutput(lngNumRows + 2)

			If mblnShowAllPagesTogether Then

				'Now add the main row data
				For lngPage = 0 To lngNumPages
					For lngCol = 0 To lngNumCols

						strTempDelim = IIf(lngCol < lngNumCols Or lngPage < lngNumPages, strDelim, "")

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mstrOutput(0) = mstrOutput(0) & Trim(mvarHeadings(0)(lngCol)) & strTempDelim


						For lngRow = 0 To lngNumRows
							mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & FormatCell(mdblDataArray(lngCol, lngRow, lngPage, lngTYPE), lngCol) & strTempDelim
						Next

						mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(mdblHorTotal(lngCol, lngPage, lngTYPE), lngCol) & strTempDelim

					Next
				Next


				If mblnPageBreak Then
					For lngCol = 0 To lngNumCols
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mstrOutput(0) = mstrOutput(0) & strDelim & Trim(mvarHeadings(0)(lngCol))

						For lngRow = 0 To lngNumRows + 1
							mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & FormatCell(mdblPgbTotal(lngCol, lngRow, lngTYPE), lngCol)
						Next
					Next
				End If

			Else
				'Now add the main row data
				For lngCol = 0 To lngNumCols

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mstrOutput(0) = mstrOutput(0) & Trim(mvarHeadings(0)(lngCol)) & IIf(lngCol <> lngNumCols, strDelim, "")
					For lngRow = 0 To lngNumRows
						mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & FormatCell(mdblDataArray(lngCol, lngRow, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
					Next

					' JDM - 10/09/2003 - Fault 7048 - Make the average column not total up.
					If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown And lngCol = iAverageColumn Then
						sngAverage = mdblHorTotal(lngCol - 1, lngSinglePage, TYPETOTAL) / mdblHorTotal(lngCol, lngSinglePage, TYPECOUNT)
						mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(sngAverage) & IIf(lngCol <> lngNumCols, strDelim, "")
					Else
						mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(mdblHorTotal(lngCol, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
					End If

				Next

				'Add the last column details (Vertical totals)
				If mlngCrossTabType = Enums.CrossTabType.cttNormal Then
					mstrOutput(0) = mstrOutput(0) & strDelim & mstrType(mlngType)
					For lngRow = 0 To lngNumRows
						mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & FormatCell(mdblVerTotal(lngRow, lngSinglePage, lngTYPE))
					Next
					mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & strDelim & FormatCell(mdblPageTotal(lngSinglePage, lngTYPE))
				End If
			End If

			Exit Sub

		Catch ex As Exception
			mstrStatusMessage = "Error building output strings (" & ex.Message.RemoveSensitive() & ")"
			fOK = False

		End Try

	End Sub

	Private Function FormatCell(dblCellValue As Double, Optional lngHOR As Integer = 0) As String

		Dim strMask As String = ""

		Try

			If dblCellValue <> 0 Or mblnSuppressZeros = False Then

				If mlngCrossTabType <> Enums.CrossTabType.cttNormal Then

					' 1000 seperators
					If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then
						strMask = IIf(mbUse1000Separator, "#,", "#") & "0.00"
					Else
						strMask = IIf(mbUse1000Separator, "#,", "#") & "0"

						If lngHOR = 2 Then
							strMask = New String("#", 20) & "0.00%"
						ElseIf lngHOR = 0 And mlngCrossTabType = Enums.CrossTabType.cttTurnover Then
							strMask = New String("#", 20) & "0.0"
						End If
					End If

				Else

					' 1000 seperators
					strMask = IIf(mbUse1000Separator, "#,0", "#0")

					If mblnShowPercentage Then
						'If percentage
						dblCellValue = dblCellValue * mdblPercentageFactor
						strMask = strMask & ".00%"

					ElseIf mlngType > 0 Then
						'if not count then
						'value should be displayed as per field definition

						If mlngIntersectionDecimals > 0 Then
							strMask = strMask & "." & New String("0", mlngIntersectionDecimals)
						End If

					End If

				End If

				Return dblCellValue.ToString(strMask)
			Else
				Return ""
			End If


		Catch ex As Exception
			mstrStatusMessage = "Error formatting data"
			fOK = False
			Return ""

		End Try

	End Function

	Private Sub GetPercentageFactor(lngPage As Integer, mlngType As Integer)

		'mdblPercentageFactor will be used in FORMATCELL, if required
		mdblPercentageFactor = 0
		If mblnShowPercentage Then
			If mblnPercentageofPage Then
				If mdblPageTotal(lngPage, mlngType) > 0 Then
					mdblPercentageFactor = 1 / mdblPageTotal(lngPage, mlngType)
				End If
			Else
				If mdblGrandTotal(mlngType) > 0 Then
					mdblPercentageFactor = 1 / mdblGrandTotal(mlngType)
				End If
			End If
		End If

	End Sub

	Public Sub BuildBreakdownStrings(lngHOR As Integer, lngVER As Integer, lngPGB As Integer)

		Dim strOutput As String
		Dim strOrder As String
		Dim strWhere As String
		Dim lngCount As Integer

		Try

			strWhere = vbNullString
			If lngHOR <= UBound(mvarSearches(HOR)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strWhere = IIf(strWhere = vbNullString, " ", strWhere & " AND ").ToString() & "(" & mvarSearches(HOR)(lngHOR) & ")"
			End If

			If lngVER <= UBound(mvarSearches(VER)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strWhere = IIf(strWhere = vbNullString, " ", strWhere & " AND ").ToString() & "(" & mvarSearches(VER)(lngVER) & ")"
			End If

			If mblnPageBreak Then
				If lngPGB <= UBound(mvarSearches(PGB)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strWhere = IIf(strWhere = vbNullString, " ", strWhere & " AND ").ToString() & "(" & mvarSearches(PGB)(lngPGB) & ")"
				End If
			End If

			Select Case mlngType
				Case TYPEMINIMUM : strOrder = "Ins, "
				Case TYPEMAXIMUM : strOrder = "Ins DESC, "
			End Select
			strOrder &= "RecDesc"

			ReDim mstrOutput(0)
			lngCount = 0

			For Each objRow As DataRow In rsCrossTabData.Select(strWhere, strOrder)

				strOutput = objRow("RecDesc")

				' Build output string for absence breakdown
				If mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown Then

					strOutput = strOutput & vbTab

					' Add absence start date
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("Start_Date")) Then
						strOutput = strOutput & vbTab
					Else
						strOutput = strOutput & CDate(objRow("Start_Date")).ToString(LocaleDateFormat) & vbTab
					End If

					' Add absence end date
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("End_Date")) Then
						strOutput = strOutput & vbTab
					Else
						strOutput = strOutput & CDate(objRow("End_Date")).ToString(LocaleDateFormat) & vbTab
					End If

					' Add occurences
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("Value")) Then
						strOutput = strOutput & vbTab
					Else
						'MH20040128 Fault 7995 - Round average to 2 decimal places
						strOutput = strOutput & Format(objRow("Value"), "0.00") & vbTab
					End If

				End If

				If mblnIntersection Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDBNull(objRow("Ins")) Then
						strOutput = strOutput & vbTab & VB6.Format(objRow("Ins"), mstrIntersectionMask)
					End If
				End If

				lngCount += 1
				ReDim Preserve mstrOutput(lngCount)
				mstrOutput(lngCount) = FormatString(strOutput)

			Next

		Catch ex As Exception
			mstrStatusMessage = "Error reading breakdown : " & ex.Message
			Throw
		End Try

	End Sub

	Public Function AbsenceBreakdownRunStoredProcedure() As Boolean

		Try

			For Each objRow As DataRow In From objRow1 As DataRow In rsCrossTabData.Rows Where CInt(objRow1("Day_Number")) < 8
				objRow("HOR") = WeekdayName(CInt(objRow("Day_Number")), False, FirstDayOfWeek.Monday)
			Next

		Catch ex As Exception
			mstrStatusMessage = "Error converting absence data"
			Return False

		End Try

		Return True

	End Function

	Public Function AbsenceBreakdownBuildDataArrays() As Boolean

		Dim strTempValue As String

		Dim lngCol As Integer
		Dim lngRow As Integer
		Dim lngPage As Integer
		Dim lngNumCols As Integer
		Dim lngNumRows As Integer
		Dim lngNumPages As Integer

		Try

			lngNumCols = UBound(mvarHeadings(0))
			lngNumRows = UBound(mvarHeadings(1))
			lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

			ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4)
			ReDim mdblHorTotal(lngNumCols, lngNumPages, 4)
			ReDim mdblVerTotal(lngNumRows, lngNumPages, 4)
			ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4)	'+1 for totals !
			ReDim mdblPageTotal(lngNumPages, 4)
			ReDim mdblGrandTotal(4)

			' Because the stored procedure has run we need to requery the recordset

			If rsCrossTabData.Rows.Count = 0 Then
				Return False
			End If

			For Each objRow As DataRow In rsCrossTabData.Rows

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				strTempValue = IIf(Not IsDBNull(objRow("HOR")), objRow("HOR").ToString().Trim(), vbNullString)
				lngCol = GetGroupNumber(strTempValue, HOR)

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				strTempValue = IIf(Not IsDBNull(objRow("VER")), objRow("VER").ToString().Trim(), vbNullString)
				lngRow = GetGroupNumber(strTempValue, VER)

				'Count
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) = mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 143)
				mdblHorTotal(lngCol, 0, TYPECOUNT) = mdblHorTotal(lngCol, 0, TYPECOUNT) + 1
				mdblVerTotal(lngRow, 0, TYPECOUNT) = mdblVerTotal(lngRow, 0, TYPECOUNT) + 1

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) = mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 143)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblHorTotal(lngCol, 0, TYPETOTAL) = mdblHorTotal(lngCol, 0, TYPETOTAL) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 0)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblVerTotal(lngRow, 0, TYPETOTAL) = mdblVerTotal(lngRow, 0, TYPETOTAL) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 0)

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 143)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 0)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsDBNull(objRow("VALUE")), objRow("VALUE"), 0)

			Next

		Catch ex As Exception
			mstrStatusMessage = "Error processing data"
			Return False

		End Try

		Return True

	End Function

	Public Function AbsenceBreakdownGetHeadingsAndSearches() As Boolean

		Dim strHeading() As String
		Dim strSearch() As String
		Dim lngLoop As Integer

		Try

			For lngLoop = 0 To 2

				ReDim strHeading(0)
				ReDim strSearch(0)

				If lngLoop = 2 And mblnPageBreak = False Then
					'When no page break field is specified
					strHeading(0) = "<None>"
				Else
					GetHeadingsAndSearchesForColumns(lngLoop, strHeading, strSearch)
				End If

				'Store each array in an array of variants (an array in an array!)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarHeadings(lngLoop) = VB6.CopyArray(strHeading)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSearches(lngLoop) = VB6.CopyArray(strSearch)

			Next

		Catch ex As Exception
			mstrStatusMessage = "Error building headings and search arrays"
			Return False

		End Try

		Return True

	End Function

	Public Function AbsenceBreakdownRetreiveDefinition(pdtStartDate As DateTime, pdtEndDate As DateTime, plngHorColID As Long, plngVerColID As Long _
																										 , plngPicklistID As Integer, plngFilterID As Integer, plngPersonnelID As Integer, pstrIncludedTypes As String) As Boolean

		Dim lngHorColID As Integer
		Dim lngVerColID As Integer
		Dim lngPicklistID As Integer
		Dim lngFilterID As Integer
		Dim lngPersonnelID As Integer
		Dim strIncludedTypes As String

		ReDim mastrUDFsRequired(0)

		' Define this cross tab as an absence breakdown
		mlngCrossTabType = Enums.CrossTabType.cttAbsenceBreakdown

		mdReportStartDate = pdtStartDate
		mdReportEndDate = pdtEndDate

		' Initialse the ok variable
		fOK = True

		' Convert variants into correct types
		lngHorColID = IIf(IsNumeric(plngHorColID), plngHorColID, 0)
		lngVerColID = IIf(IsNumeric(plngVerColID), plngVerColID, 0)
		lngPicklistID = IIf(IsNumeric(plngPicklistID), plngPicklistID, 0)
		lngFilterID = IIf(IsNumeric(plngFilterID), plngFilterID, 0)
		lngPersonnelID = IIf(IsNumeric(plngPersonnelID), plngPersonnelID, 0)

		' Force the inputted string to be formatted correctly
		'UPGRADE_WARNING: Couldn't resolve default property of object pstrIncludedTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pstrIncludedTypes = Trim(pstrIncludedTypes)
		'UPGRADE_WARNING: Couldn't resolve default property of object pstrIncludedTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strIncludedTypes = Replace(pstrIncludedTypes, "'", "''")
		strIncludedTypes = "'" & Replace(strIncludedTypes, ",", "','")
		strIncludedTypes = Mid(strIncludedTypes, 1, Len(strIncludedTypes) - 2)

		Name = "Absence Breakdown"

		'MH20040129 Fault 7857
		If DateDiff(DateInterval.Day, pdtStartDate, pdtEndDate) < 0 Then
			mstrStatusMessage = "The report end date is before the report start date."
			fOK = False
			mblnNoRecords = True
			Exit Function
		End If


		mlngBaseTableID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))
		mstrBaseTable = GetTableName(mlngBaseTableID)

		mlngRecordDescExprID = CInt(General.GetRecDescExprID(mlngBaseTableID))

		' Add the different reason types
		msAbsenceBreakdownTypes = "(" & IIf(Len(strIncludedTypes) = 0, "''", strIncludedTypes) & ")"

		' Load the appropraite records
		If lngPersonnelID > 0 Then
			mstrPicklistFilter = CStr(lngPersonnelID)
		Else
			mstrPicklistFilter = GetPicklistFilterSelect(lngPicklistID, lngFilterID)
		End If

		If fOK = False Then
			mblnNoRecords = True
			Exit Function
		End If

		mlngColID(HOR) = lngHorColID
		mstrColName(HOR) = GetColumnName(lngHorColID)
		mlngColDataType(HOR) = CStr(GetDataType(mlngBaseTableID, lngHorColID))
		mstrFormat(HOR) = GetFormat(mlngColID(HOR))

		mlngColID(VER) = lngVerColID
		mstrColName(VER) = GetColumnName(lngVerColID)
		mlngColDataType(VER) = CStr(GetDataType(mlngBaseTableID, lngVerColID))
		mstrFormat(VER) = GetFormat(mlngColID(VER))

		mlngIntersectionDecimals = 2
		mblnIntersection = False
		mblnShowAllPagesTogether = False

		Return True

	End Function

	Public Function SetAbsenceBreakDownDisplayOptions(pbShowBasePicklistFilter As Boolean) As Boolean

		' Set Report Display Options
		mblnChkPicklistFilter = pbShowBasePicklistFilter
		Return True

	End Function

	Public Function UDFFunctions(pbCreate As Boolean) As Boolean
		Return General.UDFFunctions(mastrUDFsRequired, pbCreate)
	End Function

	Private Function FormatSQLColumn(sColumn As String) As String

		Dim sReturnValue As String

		sReturnValue = sColumn
		'sReturnValue = "left(rtrim(" & sReturnValue & "), 100)"
		'sReturnValue = "replace(" & sReturnValue & ",char(9),'')"
		'sReturnValue = "replace(" & sReturnValue & ",char(10),'')"
		'sReturnValue = "replace(" & sReturnValue & ",char(13),'')"

		Return sReturnValue

	End Function

	Private Function FormatString(sHeading As String) As String

		Dim sReturnValue As String

		sReturnValue = sHeading
		'sReturnValue = Replace(sReturnValue, vbTab, "")
		sReturnValue = Replace(sReturnValue, Chr(10), "")
		sReturnValue = Replace(sReturnValue, Chr(13), "")
		sReturnValue = Replace(sReturnValue, "'", "&apos;")

		Return sReturnValue

	End Function

	Public Function ClearUp() As Boolean

		Try
			If mlngCrossTabType = CrossTabType.ctt9GridBox Then
				AccessLog.UtilUpdateLastRun(UtilityType.utlNineBoxGrid, mlngCrossTabID)
			Else
				AccessLog.UtilUpdateLastRun(UtilityType.utlCrossTab, mlngCrossTabID)
			End If

			' Delete the temptable if exists
			General.DropUniqueSQLObject(mstrTempTableName, 3)

			Return True

		Catch ex As Exception
			Throw

		End Try

	End Function

	Public Function CreatePivotDataset() As Boolean

		Dim strSQL As String

		Try

			strSQL = "SELECT HOR as 'Horizontal', VER as 'Vertical'" & IIf(PageBreakColumn, ", PGB as 'Page Break'", vbNullString) _
				& ", RecDesc as 'Record Description'" & IIf(IntersectionColumn, ", Ins as 'Intersection'", vbNullString) _
				& IIf(CrossTabType = CrossTabType.cttAbsenceBreakdown, ", Value as 'Duration'", vbNullString) _
				& " FROM " & mstrTempTableName

			If mlngCrossTabType = CrossTabType.cttAbsenceBreakdown Then
				strSQL = strSQL & " WHERE [IsSummary] = 0"

			ElseIf PageBreakColumn Then
				strSQL = strSQL & " ORDER BY PGB"
			End If

			PivotData = DB.GetDataTable(strSQL)

		Catch ex As Exception
			Return False
		End Try

		Return True
	End Function

	'****************************************************************
	' NullSafeString
	'****************************************************************
	Public Function NullSafeString(ByVal arg As Object, _
	Optional ByVal returnIfEmpty As String = "") As String

		Dim returnValue As String

		If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
			OrElse (arg Is String.Empty) Then
			returnValue = returnIfEmpty
		Else
			Try
				returnValue = CStr(arg)
			Catch
				returnValue = returnIfEmpty
			End Try

		End If

		Return returnValue

	End Function

	'****************************************************************
	' NullSafeInteger
	'****************************************************************
	Public Function NullSafeInteger(ByVal arg As Object, _
	Optional ByVal returnIfEmpty As Integer = 0) As Integer

		Dim returnValue As Integer

		If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
			OrElse (arg Is String.Empty) Then
			returnValue = returnIfEmpty
		Else
			Try
				returnValue = CInt(arg)
			Catch
				returnValue = returnIfEmpty
			End Try

		End If

		Return returnValue

	End Function

	Public Function ReturnDescriptionOrColourForNineBoxGridCell(col As Integer, row As Integer, descriptionOrColour As enumNineBoxDescriptionOrColour) As String
		If row = 0 Then
			If descriptionOrColour = enumNineBoxDescriptionOrColour.Description Then
				Return _descriptionsAsArray(col)
			Else
				Return _cellColoursAsArray(col)
			End If
		ElseIf row = 1 Then
			If descriptionOrColour = enumNineBoxDescriptionOrColour.Description Then
				Return _descriptionsAsArray(col + 3)
			Else
				Return _cellColoursAsArray(col + 3)
			End If
		Else 'row = 2
			If descriptionOrColour = enumNineBoxDescriptionOrColour.Description Then
				Return _descriptionsAsArray(col + 6)
			Else
				Return _cellColoursAsArray(col + 6)
			End If
		End If
	End Function
End Class