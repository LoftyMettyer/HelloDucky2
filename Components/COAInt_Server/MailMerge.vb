Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Public Class MailMerge
	Inherits BaseForDMI

	Private mrsMailMergeColumns As DataTable
	Private mrsMergeData As DataTable
	Private mlngMailMergeID As Integer
	Private mblnNoRecords As Boolean

	Private mlngSuccessCount As Integer
	Private mlngFailCount As Integer

	Private fOK As Boolean
	Private mstrStatusMessage As String

	'Merge Definition Variables
	Private mstrDefName As String
	Private mstrDefBaseTable As String
	Private mlngDefBaseTableID As Integer
	Private mlngDefPickListID As Integer
	Private mlngDefFilterID As Integer

	Private mstrDefTemplateFile As String
	Private mblnDefPauseBeforeMerge As Boolean
	Private mblnDefSuppressBlankLines As Boolean

	Private mstrDefEMailSubject As String
	Private mlngDefEmailAddrCalc As Integer
	Private mblnDefEMailAttachment As Boolean
	Private mstrDefAttachmentName As String

	Private mintDefOutputFormat As MailMergeOutputTypes
	Private mblnDefOutputScreen As Boolean
	Private mblnDefOutputPrinter As Boolean
	Private mstrDefOutputPrinterName As String
	Private mblnDefOutputSave As Boolean
	Private mstrDefOutputFileName As String

	Private mlngDocManMapID As Integer
	Private mblnDocManManualHeader As Boolean

	Private mintType() As SQLDataType
	Private mlngSize() As Integer
	Private mintDecimals() As Integer


	'SQL Variables
	Private mstrSQLSelect As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrder As String

	Private mlngTableViews(,) As Integer
	Private mstrWhereIDs As String

	'	Private mstrOutputArray_Data() As Object
	Private mvarPrompts(,) As Object
	Private mstrClientDateFormat As String

	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String

	Private mlngSingleRecordID As Integer

	' Modify this after we convert the actual SQL code to pull a datatable back directly
	Public ReadOnly Property MergeData As DataTable
		Get
			Return mrsMergeData
		End Get
	End Property

	Public ReadOnly Property Columns As DataTable
		Get
			Return mrsMailMergeColumns
		End Get
	End Property

	Public ReadOnly Property DefName() As String
		Get
			Return mstrDefName
		End Get
	End Property

	Public ReadOnly Property DefTemplateFile() As String
		Get
			Return mstrDefTemplateFile
		End Get
	End Property

	Public ReadOnly Property DefPauseBeforeMerge() As Boolean
		Get
			Return mblnDefPauseBeforeMerge
		End Get
	End Property

	Public ReadOnly Property DefSuppressBlankLines() As String
		Get
			Return CStr(mblnDefSuppressBlankLines)
		End Get
	End Property

	Public ReadOnly Property DefEMailSubject() As String
		Get
			Return mstrDefEMailSubject
		End Get
	End Property

	Public ReadOnly Property DefEmailAddrCalc() As Integer
		Get
			Return mlngDefEmailAddrCalc
		End Get
	End Property

	Public ReadOnly Property DefEMailAttachment() As Boolean
		Get
			Return mblnDefEMailAttachment
		End Get
	End Property

	Public ReadOnly Property DefAttachmentName() As String
		Get
			Return mstrDefAttachmentName
		End Get
	End Property

	Public ReadOnly Property DefOutputFormat() As MailMergeOutputTypes
		Get
			Return mintDefOutputFormat
		End Get
	End Property

	Public ReadOnly Property DefOutputScreen() As Boolean
		Get
			Return False ' mblnDefOutputScreen
		End Get
	End Property

	Public ReadOnly Property DefOutputPrinter() As Boolean
		Get
			Return mblnDefOutputPrinter
		End Get
	End Property

	Public ReadOnly Property DefOutputPrinterName() As String
		Get
			Return mstrDefOutputPrinterName
		End Get
	End Property

	Public ReadOnly Property DefOutputSave() As Boolean
		Get
			Return mblnDefOutputSave
		End Get
	End Property

	Public ReadOnly Property DefOutputFileName() As String
		Get
			Return mstrDefOutputFileName
		End Get
	End Property

	Public ReadOnly Property DefDocManMapID() As Integer
		Get
			Return mlngDocManMapID
		End Get
	End Property

	Public ReadOnly Property DefDocManManualHeader() As Boolean
		Get
			Return mblnDocManManualHeader
		End Get
	End Property

	Public WriteOnly Property CustomReportID() As Integer
		Set(ByVal Value As Integer)
			mlngMailMergeID = Value
		End Set
	End Property

	Public WriteOnly Property FailedMessage() As String
		Set(ByVal Value As String)
			Logs.AddDetailEntry(Value)
		End Set
	End Property

	Public WriteOnly Property ClientDateFormat() As String
		Set(ByVal Value As String)
			mstrClientDateFormat = Value
		End Set
	End Property

	Public ReadOnly Property NoRecords() As Boolean
		Get
			NoRecords = mblnNoRecords
		End Get
	End Property

	Public WriteOnly Property MailMergeID() As Integer
		Set(ByVal Value As Integer)
			mlngMailMergeID = Value
		End Set
	End Property

	Public ReadOnly Property ErrorString() As String
		Get
			ErrorString = mstrStatusMessage
		End Get
	End Property

	Public WriteOnly Property SingleRecordID() As Integer
		Set(ByVal Value As Integer)
			mlngSingleRecordID = Value
		End Set
	End Property

	Public WriteOnly Property EventLogID() As Integer
		Set(ByVal Value As Integer)
			Logs.EventLogID = Value
		End Set
	End Property

	Public Property SuccessCount() As Integer
		Get
			SuccessCount = mlngSuccessCount
		End Get
		Set(ByVal Value As Integer)
			mlngSuccessCount = Value
		End Set
	End Property

	Public Property FailCount() As Integer
		Get
			FailCount = mlngFailCount
		End Get
		Set(ByVal Value As Integer)
			mlngFailCount = Value
		End Set
	End Property

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		ReDim mlngTableViews(2, 0)
		fOK = True

	End Sub

	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	Public Function EventLogAddHeader() As Integer
		Logs.AddHeader(EventLog_Type.eltMailMerge, mstrDefName)
		EventLogAddHeader = Logs.EventLogID
	End Function

	Public Function SQLGetMergeData() As Boolean

		Dim strSQL As String

		Try

			strSQL = "SELECT " & mstrSQLSelect & " FROM " & mstrSQLFrom & mstrSQLJoin & vbNewLine & mstrSQLWhere & vbNewLine & mstrSQLOrder & vbNewLine
			mrsMergeData = DB.GetDataTable(strSQL)

			If mrsMergeData.Rows.Count = 0 Then
				mstrStatusMessage = "No records meet selection criteria"
				mblnNoRecords = True
				fOK = False
			Else
				mblnNoRecords = False
			End If

			Return fOK

		Catch ex As Exception
			fOK = False
			mstrStatusMessage = "Error retrieving merge data " & ex.Message
			Return fOK

		End Try

	End Function

	Public Function SQLGetMergeDefinition() As Boolean

		On Error GoTo LocalErr

		Dim rsMailMergeDefinition As DataTable
		Dim strSQL As String
		Dim objRow As DataRow

		strSQL = "SELECT ASRSysMailMergeName.*, ASRSysTables.TableName as TableName, ASRSysTables.RecordDescExprID as RecordDescExprID FROM ASRSysMailMergeName JOIN ASRSYSTables ON (ASRSysTables.TableID = ASRSysMailMergeName.TableID) WHERE MailMergeID = " & mlngMailMergeID & " "
		rsMailMergeDefinition = DB.GetDataTable(strSQL)

		If rsMailMergeDefinition.Rows.Count = 0 Then
			'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsMailMergeDefinition = Nothing
			mstrStatusMessage = "This definition has been deleted by another user."
			fOK = False
			Return fOK
		End If

		objRow = rsMailMergeDefinition.Rows(0)

		If LCase(CType(objRow("Username"), String)) <> LCase(gsUsername) And General.CurrentUserAccess(UtilityType.utlMailMerge, mlngMailMergeID) = ACCESS_HIDDEN Then
			'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsMailMergeDefinition = Nothing
			mstrStatusMessage = "This definition has been made hidden by another user."
			fOK = False
			GoTo TidyAndExit
		End If

		mstrDefName = objRow("Name").ToString()
		mlngDefBaseTableID = CInt(objRow("TableID"))
		mstrDefBaseTable = objRow("TableName").ToString()
		mlngDefPickListID = CInt(objRow("PickListID"))
		mlngDefFilterID = CInt(objRow("FilterID"))

		mstrDefTemplateFile = objRow("TemplateFileName").ToString()
		mblnDefSuppressBlankLines = CBool(objRow("SuppressBlanks"))
		mblnDefPauseBeforeMerge = CBool(objRow("PauseBeforeMerge"))

		mlngDefEmailAddrCalc = 0
		'mstrEmailAddr = vbNullString

		mintDefOutputFormat = CType(objRow("OutputFormat"), MailMergeOutputTypes)
		Select Case mintDefOutputFormat
			Case MailMergeOutputTypes.WordDocument
				mblnDefOutputScreen = CBool(objRow("OutputScreen"))
				mblnDefOutputPrinter = CBool(objRow("OutputPrinter"))
				mstrDefOutputPrinterName = objRow("OutputPrinterName").ToString()
				mblnDefOutputSave = CBool(objRow("OutputSave"))
				mstrDefOutputFileName = objRow("OutputFilename").ToString()

			Case MailMergeOutputTypes.IndividualEmail
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IIf(IsDBNull(objRow("EmailAddrID")), 0, objRow("EmailAddrID")) = 0 Then
					mstrStatusMessage = "No destination email address selected"
					'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsMailMergeDefinition = Nothing
					fOK = False
					GoTo TidyAndExit
				End If

				mstrDefEMailSubject = objRow("EmailSubject").ToString()
				mlngDefEmailAddrCalc = CInt(objRow("EmailAddrID"))
				mblnDefEMailAttachment = CBool(objRow("EMailAsAttachment"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrDefAttachmentName = IIf(IsDBNull(objRow("EmailAttachmentName")), "", objRow("EmailAttachmentName")).ToString()

			Case MailMergeOutputTypes.DocumentManagement
				mblnDefOutputPrinter = True
				mblnDefOutputScreen = CBool(objRow("OutputScreen"))
				mstrDefOutputPrinterName = objRow("OutputPrinterName").ToString()
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngDocManMapID = CInt(IIf(IsDBNull(objRow("DocumentMapID")), 0, objRow("DocumentMapID")))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mblnDocManManualHeader = CBool(IIf(IsDBNull(objRow("ManualDocManHeader")), 0, objRow("ManualDocManHeader")))

		End Select


		fOK = IsRecordSelectionValid()

		If fOK Then
			AccessLog.UtilUpdateLastRun(UtilityType.utlMailMerge, mlngMailMergeID)
		End If

TidyAndExit:
		'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsMailMergeDefinition = Nothing
		Return fOK

LocalErr:
		mstrStatusMessage = "Error reading Mail Merge definition"
		fOK = False
		Return fOK

	End Function

	Private Function GetPicklistFilterSelect() As String

		Dim rsTemp As DataTable

		On Error GoTo LocalErr

		If mlngSingleRecordID > 0 Then
			GetPicklistFilterSelect = CStr(mlngSingleRecordID)

		ElseIf mlngDefPickListID > 0 Then

			mstrStatusMessage = IsPicklistValid(mlngDefPickListID)
			If mstrStatusMessage <> vbNullString Then
				fOK = False
				Exit Function
			End If

			'Get List of IDs from Picklist
			rsTemp = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & mlngDefPickListID)
			fOK = (rsTemp.Rows.Count > 0)

			If Not fOK Then
				mstrStatusMessage = "The base table picklist contains no records."
			Else
				GetPicklistFilterSelect = vbNullString

				For Each objRow As DataRow In rsTemp.Rows
					GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "").ToString() & objRow(0).ToString()

				Next
			End If


		ElseIf mlngDefFilterID > 0 Then

			mstrStatusMessage = IsFilterValid(mlngDefFilterID)
			If mstrStatusMessage <> vbNullString Then
				fOK = False
				Exit Function
			End If

			'Get list of IDs from Filter
			fOK = General.FilteredIDs(mlngDefFilterID, GetPicklistFilterSelect, mastrUDFsRequired, mvarPrompts)

			If Not fOK Then
				' Permission denied on something in the filter.
				mstrStatusMessage = "You do not have permission to use the '" & General.GetFilterName(mlngDefFilterID) & "' filter."
			End If

		End If

		Exit Function

LocalErr:
		mstrStatusMessage = "Error processing picklist"
		fOK = False

	End Function


	Public Function SQLGetMergeColumns() As Boolean

		Dim strSQL As String
		Dim strSQL1 As String
		Dim strSQL2 As String

		On Error GoTo LocalErr

		'Merge Column Types
		'
		' "C" is a column which has been selected by the user
		' "E" is an express which has been selected by the user
		' "X" is a system column required by the merge
		'     (currently only used for the email field where required)

		strSQL1 = "SELECT 'ColExp'   = 'Col',  'ColExpId' = ASRSysColumns.ColumnID, 'TableID'  = ASRSysTables.TableID, 'Table'    = ASRSysTables.Tablename, 'Name'     = ASRSysColumns.ColumnName, 'Type'     = ASRSysColumns.DataType, 'Size'     = ASRSysMailMergeColumns.Size, 'Decimals' = ASRSysMailMergeColumns.Decimals FROM ASRSysMailMergeColumns JOIN ASRSysColumns ON (ASRSysColumns.ColumnID = ASRSysMailMergeColumns.ColumnID) JOIN ASRSysTables ON (ASRSysTables.TableID = ASRSysColumns.TableID) WHERE ASRSysMailMergeColumns.Type = 'C' AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "
		strSQL2 = "SELECT 'ColExp'   = 'Exp', 'ColExpId' = ASRSysExpressions.ExprID, 'TableID'  = 0, 'Table'    = 'Calculation_', 'Name'     = ASRSysExpressions.Name, 'Type' = ASRSysExpressions.ReturnType, 'Size'     = ASRSysMailMergeColumns.Size, 'Decimals' = ASRSysMailMergeColumns.Decimals FROM ASRSysMailMergeColumns LEFT OUTER JOIN ASRSysExpressions ON (ASRSysExpressions.ExprID = ASRSysMailMergeColumns.ColumnID) WHERE ASRSysMailMergeColumns.Type = 'E' AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "
		strSQL = strSQL1 & vbNewLine & "UNION" & vbNewLine & vbNewLine & strSQL2 & vbNewLine & "ORDER BY 'ColExp', 'Table', 'Name'"
		mrsMailMergeColumns = DB.GetDataTable(strSQL)

		Return fOK

LocalErr:
		mstrStatusMessage = "Error reading calculation/column definitions."
		fOK = False
		SQLGetMergeColumns = fOK

	End Function

	Public Function SQLCodeCreate() As Boolean

		Dim strPicklistFilterSelect As String
		Dim objExpr As clsExprExpression
		Dim intIndex As Integer
		Dim objTableView As TablePrivilege

		fOK = True

		On Error GoTo LocalErr

		mstrSQLSelect = vbNullString
		mstrSQLFrom = vbNullString
		mstrSQLJoin = vbNullString
		mstrSQLWhere = vbNullString
		mstrSQLOrder = vbNullString

		ReDim mastrUDFsRequired(0)

		ReDim mlngTableViews(2, 0)

		intIndex = 0
		ReDim mintType(intIndex)
		ReDim mlngSize(intIndex)

		' JPD20030219 Fault 5070
		' Check the user has permission to read the base table.
		fOK = False
		For Each objTableView In gcoTablePrivileges.Collection
			If (objTableView.TableID = mlngDefBaseTableID) And (objTableView.AllowSelect) Then
				fOK = True
				Exit For
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing

		If Not fOK Then
			mstrStatusMessage = "You do not have permission to read the base table either directly or through any views."
			Exit Function
		End If

		mstrSQLFrom = gcoTablePrivileges.Item(mstrDefBaseTable).RealSource
		mstrSQLSelect = mstrSQLFrom & ".ID, '' AS [?Receipt]"

		With mrsMailMergeColumns
			fOK = (.Rows.Count > 0)
			If fOK Then

				For Each objRow As DataRow In .Rows

					'01/08/2001 MH Fault 2125
					intIndex = intIndex + 1
					ReDim Preserve mintType(intIndex)
					ReDim Preserve mlngSize(intIndex)
					ReDim Preserve mintDecimals(intIndex)

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mlngSize(intIndex) = CInt(IIf(IsDBNull(objRow("Size")), 0, objRow("Size")))
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mintDecimals(intIndex) = CInt(IIf(IsDBNull(objRow("Decimals")), 0, objRow("Decimals")))

					Select Case objRow("ColExp")
						Case "Col"
							Call SQLAddColumn(mstrSQLSelect, CInt(objRow("TableID")), objRow("Table").ToString(), objRow("Name").ToString(), objRow("Table").ToString() & "_" & objRow("Name").ToString())
							mintType(intIndex) = CType(objRow("Type"), SQLDataType)

						Case "Exp"
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(objRow("Name")) Then
								'MH20011127
								'mstrStatusMessage = _
								'"This definition contains one or more calculation(s) which" & vbNewLine & _
								'"have been deleted by another user."
								mstrStatusMessage = "This definition contains one or more calculation(s) which " & "have been deleted by another user."
								fOK = False
								Exit Function

							ElseIf IsCalcValid(CInt(objRow("ColExpID"))) <> vbNullString Then
								'MH20011127
								'mstrStatusMessage = "You cannot run this Global definition as it contains one or more calculation(s) which have been deleted or made hidden by another user." & vbNewLine & _
								'"Please re-visit your definition to remove the hidden calculations." & vbNewLine
								mstrStatusMessage = "You cannot run this Mail Merge definition as it contains one or more calculation(s) which have been deleted or made hidden by another user. " & "Please re-visit your definition to remove the hidden calculations."
								fOK = False
								Exit Function

							Else
								SQLAddCalculation(CInt(objRow("ColExpID")), objRow("Table").ToString() & objRow("Name").ToString())

								objExpr = NewExpression()
								objExpr.ExpressionID = CInt(objRow("ColExpID"))
								objExpr.ConstructExpression()
								objExpr.ValidateExpression(True)

								Select Case objExpr.ReturnType
									Case ExpressionValueTypes.giEXPRVALUE_DATE, ExpressionValueTypes.giEXPRVALUE_BYREF_DATE
										mintType(intIndex) = SQLDataType.sqlDate
									Case ExpressionValueTypes.giEXPRVALUE_NUMERIC, ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
										mintType(intIndex) = SQLDataType.sqlNumeric
									Case ExpressionValueTypes.giEXPRVALUE_LOGIC, ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC
										mintType(intIndex) = SQLDataType.sqlBoolean
									Case Else
										mintType(intIndex) = 0
								End Select

								'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objExpr = Nothing

							End If

					End Select

					If fOK = False Then
						Exit Function
					End If

				Next

			End If
		End With

		If mstrWhereIDs <> vbNullString Then
			mstrSQLWhere = "(" & mstrWhereIDs & ")" & IIf(mstrSQLWhere <> vbNullString, " OR ", "").ToString() & mstrSQLWhere
		End If

		strPicklistFilterSelect = GetPicklistFilterSelect()
		If fOK = False Then
			Exit Function
		End If
		If strPicklistFilterSelect <> vbNullString Then
			mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", "").ToString() & mstrSQLFrom & ".ID IN (" & strPicklistFilterSelect & ")"
		End If

		If mstrSQLWhere <> vbNullString Then
			mstrSQLWhere = " WHERE " & mstrSQLWhere
		End If

		'UPGRADE_NOTE: Object mrsMailMergeColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mrsMailMergeColumns = Nothing

		Call SQLOrderByClause()
		Return fOK

LocalErr:
		mstrStatusMessage = "Error processing calculation/column definitions."
		Return False

	End Function

	Private Sub SQLAddColumn(ByRef sColumnList As String, lngTableID As Integer, sTableName As String, sColumnName As String, strColCode As String)

		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim fColumnOK As Boolean
		Dim sSource As String
		Dim fFound As Boolean
		Dim iNextIndex As Integer

		Dim sRealSource As String
		Dim sCaseStatement As String
		Dim sWhereColumn As String
		Dim sBaseIDColumn As String

		Dim asViews() As String

		On Error GoTo LocalErr


		objColumnPrivileges = GetColumnPrivileges(sTableName)
		fColumnOK = objColumnPrivileges.IsValid(sColumnName)
		If fColumnOK Then
			fColumnOK = objColumnPrivileges.Item(sColumnName).AllowSelect
		End If

		'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objColumnPrivileges = Nothing

		If fColumnOK Then
			' The column can be read from the base table/view, or directly from a parent table.
			' Add the column to the column list.

			sRealSource = gcoTablePrivileges.Item(sTableName).RealSource

			sColumnList = sColumnList & IIf(sColumnList <> vbNullString, ", ", "").ToString() & sRealSource & "." & sColumnName

			If strColCode <> vbNullString Then
				sColumnList = sColumnList & " AS " & "'" & strColCode & "'"
			End If

			If sTableName <> mstrDefBaseTable Then

				fFound = False
				For iNextIndex = 1 To UBound(mlngTableViews, 2)
					If mlngTableViews(1, iNextIndex) = 0 And mlngTableViews(2, iNextIndex) = lngTableID Then
						fFound = True
						Exit For
					End If
				Next iNextIndex

				' if this column is not from the base table then it must be from a parent
				' table, therefore include it in the join code
				If Not fFound Then
					iNextIndex = UBound(mlngTableViews, 2) + 1
					ReDim Preserve mlngTableViews(2, iNextIndex)
					mlngTableViews(1, iNextIndex) = 0
					mlngTableViews(2, iNextIndex) = lngTableID


					' The table has not yet been added to the join code, and it is
					' not the base table so add it to the array and the join code.
					mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & sRealSource & " ON " & mstrSQLFrom & ".ID_" & CStr(lngTableID) & " = " & sRealSource & ".ID"
				End If

			End If

		Else

			ReDim asViews(0)
			For Each objTableView In gcoTablePrivileges.Collection

				'Loop thru all of the views for this table where the user has select access
				If (Not objTableView.IsTable) And (objTableView.TableID = lngTableID) And (objTableView.AllowSelect) Then

					sSource = objTableView.ViewName

					' Get the column permission for the view.
					objColumnPrivileges = GetColumnPrivileges(sSource)

					If objColumnPrivileges.IsValid(sColumnName) Then
						If objColumnPrivileges.Item(sColumnName).AllowSelect Then
							' Add the view info to an array to be put into the column list or order code below.
							iNextIndex = UBound(asViews) + 1
							ReDim Preserve asViews(iNextIndex)
							asViews(iNextIndex) = sSource

							'=== This is the join code section ===
							' Add the view to the Join code.
							' Check if the view has already been added to the join code.
							fFound = False
							For iNextIndex = 1 To UBound(mlngTableViews, 2)
								If mlngTableViews(2, iNextIndex) = objTableView.ViewID Then
									fFound = True
									Exit For
								End If
							Next iNextIndex

							If Not fFound Then
								' The view has not yet been added to the join code, so add it to the array and the join code.

								iNextIndex = UBound(mlngTableViews, 2) + 1
								ReDim Preserve mlngTableViews(2, iNextIndex)
								mlngTableViews(1, iNextIndex) = 1
								mlngTableViews(2, iNextIndex) = objTableView.ViewID

								If objTableView.TableID = mlngDefBaseTableID Then
									sBaseIDColumn = mstrSQLFrom & ".ID"
								Else
									sBaseIDColumn = mstrSQLFrom & ".ID_" & CStr(objTableView.TableID)
								End If

								mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & sBaseIDColumn & " = " & sSource & ".ID"

								mstrWhereIDs = mstrWhereIDs & IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString).ToString() & sBaseIDColumn & " IN (SELECT ID FROM " & sSource & ")" & " OR (ISNULL(" & sBaseIDColumn & ", 0) = 0)"

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
				mstrStatusMessage = "You do not have permission to see the column '" & sColumnName & "' " & "either directly or through any views."
				fOK = False
				Exit Sub

			Else
				' Add the column to the column list.
				sCaseStatement = "CASE"
				sWhereColumn = vbNullString
				For iNextIndex = 1 To UBound(asViews)
					sCaseStatement = sCaseStatement & " WHEN NOT " & asViews(iNextIndex) & "." & sColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & sColumnName & vbNewLine
				Next iNextIndex

				If Len(sCaseStatement) > 0 Then
					sCaseStatement = sCaseStatement & " ELSE NULL END"

					If strColCode <> vbNullString Then
						sCaseStatement = sCaseStatement & " AS " & "'" & strColCode & "'"
					End If

					sColumnList = sColumnList & IIf(Len(sColumnList) > 0, ", ", "").ToString() & vbNewLine & sCaseStatement

					If sWhereColumn <> vbNullString Then
						mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", vbNullString).ToString() & "((" & sWhereColumn & "))"
					End If

				End If
			End If
		End If

		Exit Sub

LocalErr:
		mstrStatusMessage = "Error building SQL Statement"
		fOK = False

	End Sub

	Private Sub SQLOrderByClause()

		Dim rsTemp As DataTable
		Dim strSQL As String

		Try

			strSQL = "SELECT ASRSysTables.TableID, ASRSysTables.TableName, ASRSysColumns.ColumnID, ASRSysColumns.ColumnName, ASRSysMailMergeColumns.SortOrder FROM ASRSysMailMergeColumns JOIN ASRSysColumns ON (ASRSysMailMergeColumns.ColumnID = ASRSysColumns.ColumnID) JOIN ASRSysTables ON (ASRSysColumns.TableID = ASRSysTables.TableID) WHERE ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " AND SortOrderSequence > 0 ORDER BY SortOrderSequence"
			rsTemp = DB.GetDataTable(strSQL)

			With rsTemp
				For Each objRow As DataRow In .Rows

					SQLAddColumn(mstrSQLOrder, CInt(objRow("TableID")), objRow("TableName").ToString(), objRow("ColumnName").ToString(), vbNullString)

					mstrSQLOrder = mstrSQLOrder & IIf(Left(objRow("SortOrder").ToString(), 1) = "A", " ASC", " DESC").ToString()

					If fOK = False Then
						Exit Sub
					End If
				Next
			End With

			If mstrSQLOrder <> vbNullString Then
				mstrSQLOrder = " ORDER BY " & mstrSQLOrder
			End If

		Catch ex As Exception
			mstrStatusMessage = "Error building 'Order By' clause"
			fOK = False

		End Try

	End Sub

	Private Sub SQLAddCalculation(lngExpID As Integer, strColCode As String)

		Dim lngCalcViews(,) As Integer
		Dim objCalcExpr As clsExprExpression
		Dim intCount As Integer
		Dim blnFound As Boolean
		Dim intNextIndex As Integer
		Dim sCalcCode As String
		Dim sSource As String
		Dim lngTestTableID As Integer
		Dim objTableView As TablePrivilege

		ReDim lngCalcViews(2, 0)
		objCalcExpr = NewExpression()
		fOK = objCalcExpr.Initialise(mlngDefBaseTableID, lngExpID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
		If fOK Then
			fOK = objCalcExpr.RuntimeCalculationCode(lngCalcViews, sCalcCode, mastrUDFsRequired, True, False, mvarPrompts)
		End If

		If fOK = False Then
			'mstrStatusMessage = "You do not have permission to run a calculation contained in this mail merge."
			mstrStatusMessage = "You do not have permission to use the '" & Trim(objCalcExpr.Name) & "' calculation."
			'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalcExpr = Nothing
			Exit Sub
		End If
		'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objCalcExpr = Nothing


		mstrSQLSelect = mstrSQLSelect & IIf(mstrSQLSelect <> vbNullString, ", ", vbNullString).ToString() & sCalcCode & " AS '" & strColCode & "'"


		' Add the required views to the JOIN code.
		For intCount = 1 To UBound(lngCalcViews, 2)
			If lngCalcViews(1, intCount) = 1 Then
				' Check if view has already been added to the array
				blnFound = False
				For intNextIndex = 1 To UBound(mlngTableViews, 2)
					If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
						blnFound = True
						Exit For
					End If
				Next intNextIndex

				If Not blnFound Then
					' View hasnt yet been added, so add it !
					intNextIndex = UBound(mlngTableViews, 2) + 1
					ReDim Preserve mlngTableViews(2, intNextIndex)
					mlngTableViews(1, intNextIndex) = 1
					mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount)

					lngTestTableID = lngCalcViews(2, intCount)

					objTableView = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount))
					sSource = objTableView.RealSource

					'TM20020904 Fault 4364 - depending on whether the table that is about to
					'                        joined is a Parent or Child denotes which ID
					'                        columns are used to establish the join.
					If General.IsAParentOf((objTableView.TableID), mlngDefBaseTableID) Then
						'Table/View is parent of Base Table.
						mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID_" & CStr(objTableView.TableID) & " = " & sSource & ".ID"
					End If

				End If

			ElseIf lngCalcViews(1, intCount) = 0 Then
				' Check if table has already been added to the array
				blnFound = False
				For intNextIndex = 1 To UBound(mlngTableViews, 2)
					If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
						blnFound = True
						Exit For
					End If
				Next intNextIndex

				'TM20020827 Fault 4339
				'Don't add the table id to the array if it is the base table id.
				If Not blnFound Then
					blnFound = (lngCalcViews(2, intCount) = mlngDefBaseTableID)
				End If

				If Not blnFound Then
					' Table hasnt yet been added, so add it !
					intNextIndex = UBound(mlngTableViews, 2) + 1
					ReDim Preserve mlngTableViews(2, intNextIndex)
					mlngTableViews(1, intNextIndex) = 0
					mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount)

					lngTestTableID = lngCalcViews(2, intCount)

					objTableView = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount))
					'sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
					sSource = objTableView.RealSource

					'TM20020904 Fault 4364 - depending on whether the table that is about to
					'                        joined is a Parent or Child denotes which ID
					'                        columns are used to establish the join.
					If General.IsAParentOf((objTableView.TableID), mlngDefBaseTableID) Then
						'Table/View is parent of Base Table.
						mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID_" & lngCalcViews(2, intCount) & " = " & sSource & ".ID"
					End If

				End If

			End If
		Next

	End Sub

	Public Sub EventLogChangeHeaderStatus(lngStatus As EventLog_Status)
		Logs.ChangeHeaderStatus(lngStatus, mlngSuccessCount, mlngFailCount)
	End Sub

	Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop As Short
		Dim iDataType As Short
		Dim lngComponentID As Integer

		fOK = True

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

		Return fOK

ErrorTrap:
		mstrStatusMessage = "Error whilst setting prompted values. " & Err.Description
		Return False

	End Function

	Private Function IsRecordSelectionValid() As Boolean

		Dim sSQL As String
		Dim rsTemp As DataTable
		Dim iResult As RecordSelectionValidityCodes
		Dim fCurrentUserIsSysSecMgr As Boolean

		fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

		' Filter
		If mlngDefFilterID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_FILTER, mlngDefFilterID)
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
		ElseIf mlngDefPickListID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_PICKLIST, mlngDefPickListID)
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

		'******* Check calculations for hidden/deleted elements *******
		If Len(mstrStatusMessage) = 0 Then
			sSQL = "SELECT * FROM ASRSysMailMergeColumns WHERE MailMergeID = " & mlngMailMergeID & " AND LOWER(Type) = 'e' "

			rsTemp = DB.GetDataTable(sSQL)
			With rsTemp
				If .Rows.Count > 0 Then

					For Each objRow As DataRow In .Rows

						iResult = ValidateCalculation(CInt(objRow("ColumnID")))
						Select Case iResult
							Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
								mstrStatusMessage = "A calculation used in this definition has been deleted by another user."
							Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
								mstrStatusMessage = "A calculation used in this definition is invalid."
							Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
								If Not fCurrentUserIsSysSecMgr Then
									mstrStatusMessage = "A calculation used in this definition has been made hidden by another user."
								End If
						End Select

						If Len(mstrStatusMessage) > 0 Then
							Exit For
						End If

					Next
				End If
			End With

			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If

		IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)

	End Function

	Public Function UDFFunctions(ByRef pbCreate As Boolean) As Boolean
		Return General.UDFFunctions(mastrUDFsRequired, pbCreate)
	End Function

End Class