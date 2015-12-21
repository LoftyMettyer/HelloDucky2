Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Expressions
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.ReportOutput
Imports HR.Intranet.Server.Enums

Public Class MatchReportRun
  	Inherits BaseReport

  public Data As New List(of String)

  'LOFTY -- SECOND HACK FOR DISPLAY ASPX
  Public Property ReportCaption as String
  Public Property DisplayColumns as New Collection(Of DisplayColumn)

 	Public Property ReportDataTable As New DataTable

  'LOFTY -- QUICK HACK - NEEDS SORTING
  private glngPostTableID As Integer = 2
  private glngPersonnelTableID As Integer = 1
  Private gstrPostTableName as String = ""
  Private gstrJobTitleColumnName as String = ""
  Private gsPersonnelTableName as String = ""
  Private gsPersonnelJobTitleColumnName as String = ""
  Private gsPersonnelGradeColumnName as String = ""
  Private gstrGradeTableName as String = ""
  Private gstrPostGradeColumnName as String = ""
  Private gstrNumLevelColumnName as String = ""
  Private gstrGradeColumnName as String = ""
  Private gsPersonnelEmployeeNumberColumnName as string
  Private gsPersonnelManagerStaffNoColumnName as string

  private gblnReportPackMode as Boolean = False

  'LOFTY END HACK

	Private mblnUserCancelled As Boolean
	Private mstrTempTableName As String
	
	Private mcolColDetails As Collection
	Private mcolRelations As Collection
  Private frmBreakDown As frmMatchRunBreakDown
	Private mlngTableViews(,) As Integer
	Private mstrExcelFormats() As String
	Private fOK As Boolean
	'Private gblnBatchMode As Boolean
	Private mstrErrorMessage As String
	Private mblnNoRecords As Boolean
	Private mbDefinitionOwner As Boolean
	Private alngSourceTables(,) As Integer
	
	Private mlngMatchReportID As Integer
	Private mlngMatchReportType As MatchReportType
	Private mstrRecordSelectionName As String
	Private mstrDescription As String
	Private mlngNumRecords As Integer
	Private mblnEqualGrade As Boolean
	Private mblnReportingStructure As Boolean
	
	Private mlngScoreMode As Integer
	Private mblnScoreCheck As Boolean
	Private mlngScoreLimit As Integer
	
	Private mlngTable1ID As Integer
	Private mstrTable1Name As String
	Private mstrTable1RealSource As String
	Private mlngTable1RecDescExprID As Integer
	Private mlngTable1AllRecords As Integer
	Private mlngTable1PickListID As Integer
	Private mlngTable1FilterID As Integer
	Private mstrTable1Where As String
	
	Private mlngTable2ID As Integer
	Private mstrTable2Name As String
	Private mstrTable2RealSource As String
	Private mlngTable2RecDescExprID As Integer
	Private mlngTable2AllRecords As Integer
	Private mlngTable2PickListID As Integer
	Private mlngTable2FilterID As Integer
	Private mstrTable2Where As String
	
	Private mstrSQL As String
	Private mstrSQLWhere As String
	Private mstrSQLGroupBy As String
	Private mstrSQLOrderBy As String
	Private mstrSQLMatchScore As String
	
	Private mcolSQLSelect As Collection
	Private mcolSQLJoin As Collection
	Private mcolSQLWhere As Collection
	Private mcolSQLOrderBy As Collection
	Private mcolSQLMatchScore As Collection
	'Private mcolSQLOrderBy As Collection
	
	Private mblnPreviewOnScreen As Boolean
	
	'New Default Output Variables
	Private mlngOutputFormat As Integer
	Private mblnOutputScreen As Boolean
	Private mblnOutputPrinter As Boolean
	Private mstrOutputPrinterName As String
	Private mblnOutputSave As Boolean
	Private mlngOutputSaveExisting As Integer
	'Private mlngOutputSaveFormat As Long ' May need in future
	Private mblnOutputEmail As Boolean
	Private mlngOutputEmailAddr As Integer
	Private mstrOutputEmailSubject As String
	Private mstrOutputEmailAttachAs As String
	'Private mlngOutputEmailFileFormat As Long ' May need in future
	Private mstrOutputFileName As String
	Private mblnChkPicklistFilter As Boolean 'might not need
	Private mstrOutputTitlePage As String
	Private mstrOutputReportPackTitle As String
	Private mstrOutputOverrideFilter As String
	Private mblnOutputTOC As Boolean
	Private mlngOverrideFilterID As Integer
	Private mblnOutputRetainPivotOrChart As Boolean
	'Private mblnOutputRetainCharts As Boolean
	
	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String
	Private mvarPrompts(,) As Object
	
	Public WriteOnly Property MatchReportID() As Integer
		Set(ByVal Value As Integer)
			mlngMatchReportID = Value
		End Set
	End Property
	
	Public WriteOnly Property MatchReportType_Renamed() As MatchReportType
		Set(ByVal Value As MatchReportType)		
			mlngMatchReportType = Value			
		End Set
	End Property
	
	Public ReadOnly Property ErrorString() As String
		Get
			ErrorString = mstrErrorMessage
		End Get
	End Property
	
	Public ReadOnly Property UserCancelled() As Boolean
		Get
			UserCancelled = mblnUserCancelled
		End Get
	End Property
	
	Public ReadOnly Property NoRecords() As Boolean
		Get
			NoRecords = mblnNoRecords
		End Get
	End Property
	
	
	Public ReadOnly Property PreviewOnScreen() As Boolean
		Get
			PreviewOnScreen = ((fOK And mblnPreviewOnScreen) And Not mblnNoRecords)
		End Get
	End Property
	
	Public Function RunMatchReport(optional plngTableID As Integer = 0, Optional plngRecordID As Integer = 0) As List(Of String)
		Dim strUtilityName As String
		
    Try

		  fOK = True
		
		  If frmBreakDown Is Nothing Then
			  frmBreakDown = New frmMatchRunBreakDown
		  End If
		
		  If fOK Then fOK = GetMatchReportDefinition
		
		  strUtilityName = "Match Report"
		
		  Select Case mlngMatchReportType
			  Case MatchReportType.mrtNormal
				  Logs.AddHeader(EventLog_Type.eltMatchReport, Name)
			  Case MatchReportType.mrtSucession
				  Logs.AddHeader(EventLog_Type.eltSuccessionPlanning, Name)
				  strUtilityName = "Succession Planning Report"
			  Case MatchReportType.mrtCareer
				  Logs.AddHeader(EventLog_Type.eltCareerProgression, Name)
				  strUtilityName = "Career Progression Report"
		  End Select
		
		  If fOK Then fOK = GetDetailsRecordsets
		  If fOK Then fOK = GetRelationRecordsets
		  If fOK Then fOK = CheckModuleSetupPermissions
		  If fOK Then fOK = GetDataRecordset(plngTableID, plngRecordID)

      ReportDataTable.Columns.Add("talentchart", GetType(String))
      DisplayColumns.Add(New DisplayColumn() With {.Name = "talentchart"})
		
		  'UPGRADE_WARNING: Couldn't resolve default property of object InitialiseFormBreakdown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		  If fOK Then fOK = InitialiseFormBreakdown
      If fOK Then fOK = PopulateGridMain
		
		  RemoveTemporarySQLObjects()
		
		  If fOK Then
			  If Not mblnPreviewOnScreen Then
				  fOK = OutputReport(False)
			  End If
		  End If
		

		  AccessLog.UtilUpdateLastRun(UtilityType.utlMatchReport, mlngMatchReportID)
		  mblnUserCancelled = (InStr(LCase(mstrErrorMessage), "cancelled by user") > 0)
		
		  If mblnNoRecords Then
			  Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
			  Logs.AddDetailEntry(mstrErrorMessage)
			  mstrErrorMessage = "Completed successfully." & vbCrLf & mstrErrorMessage
			  fOK = True
		  ElseIf fOK Then 
			  Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
			  mstrErrorMessage = "Completed successfully."
		  ElseIf mblnUserCancelled Then 
			  Logs.ChangeHeaderStatus(EventLog_Status.elsCancelled)
			  mstrErrorMessage = "Cancelled by user."
		  Else
			  'Only details records for failures !
			  Logs.AddDetailEntry(mstrErrorMessage)
			  Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			  mstrErrorMessage = "Failed." & vbCrLf & vbCrLf & mstrErrorMessage
		  End If

		  Return Data

	  Catch ex As Exception
        fOK = False
        mstrErrorMessage = "Error whilst running this definition." & vbCrLf & ex.Message
        Return Nothing

    End Try
		
	End Function
	
	Private Function GetDetailsRecordsets() As Boolean
				
		Dim objColumn As DisplayColumn
		Dim rsMatchReportsDetails As DataTable
		Dim strTempSQL As String
		Dim intTemp As Short
		
		Try

		  strTempSQL = "SELECT ASRSysMatchReportDetails.*, " & "       ASRSysTables.TableID, ASRSysTables.TableName," & "       ASRSysColumns.ColumnName, ASRSysColumns.DataType " & "FROM ASRSysMatchReportDetails " & "LEFT OUTER JOIN   ASRSysColumns on ASRSysMatchReportDetails.colexprid = ASRSysColumns.columnid" & "                  and ASRSysMatchReportDetails.ColType = 'C' " & "LEFT OUTER JOIN   ASRSysTables on ASRSysColumns.TableID = ASRSysTables.TableID " & "WHERE  MatchReportID = " & CStr(mlngMatchReportID) & " " & "ORDER BY [ColSequence]"
		
		  rsMatchReportsDetails = General.GetReadOnlyRecords(strTempSQL)
		
		  mcolColDetails = New Collection
		
		  objColumn = New DisplayColumn
		  objColumn.ColType = "C"
		  objColumn.TableID = mlngTable1ID
		  objColumn.TableName = mstrTable1Name
		  objColumn.Name = "ID"
		  objColumn.Hidden = True
		  objColumn.Heading = "ID1"
		  mcolColDetails.Add(objColumn, "ID1")
		
		  If mlngTable2ID > 0 Then
			  objColumn = New DisplayColumn
			  objColumn.ColType = "C"
			  objColumn.TableID = mlngTable2ID
			  objColumn.TableName = mstrTable2Name
			  objColumn.Name = "ID"
			  objColumn.Hidden = True
			  objColumn.Heading = "ID2"
			  mcolColDetails.Add(objColumn, "ID2")
		  End If
	
		
		  Dim objExpr As clsExprExpression
		  With rsMatchReportsDetails
			  If .Rows.Count = 0 Then
				  GetDetailsRecordsets = False
				  mstrErrorMessage = "No columns found in the specified definition." & vbCrLf & "Please remove this definition and create a new one."
				  Exit Function
			  End If
			
			  intTemp = 0
			  for each objRow as DataRow in .Rows
				  intTemp = intTemp + 1
				
				  ReDim Preserve mstrExcelFormats(intTemp)
				
				  objColumn = New DisplayColumn

				  objColumn.ColType = objRow("ColType").ToString
				  objColumn.ID = objRow("ColExprID")
				  objColumn.Size = objRow("ColSize")
				  objColumn.Decimals = objRow("ColDecs")
				  objColumn.Heading = objRow("ColHeading").ToString()
				  objColumn.Sequence = objRow("ColSequence").ToString()
				  objColumn.SortSeq = objRow("SortOrderSeq").ToString()
				  objColumn.SortDir = objRow("SortOrderDirection").ToString()
				  objColumn.Use1000Separator = DoesColumnUseSeparators(CInt(objRow("ColExprID")))
				
				  If objColumn.ColType = "C" Then
					  objColumn.TableID = CInt(objRow("TableID"))
					  objColumn.TableName = objRow("TableName").ToString()
					  objColumn.Name = objRow("ColumnName").ToString()
					  objColumn.DataType = CInt(objRow("DataType"))
					
					  Select Case CInt(objRow("DataType"))
						  Case ColumnDataType.sqlNumeric, ColumnDataType.sqlInteger
							
							  If objColumn.Decimals > 0 Then
								  If objColumn.Decimals > 127 Then
									  mstrExcelFormats(intTemp) = "0." & New String("0", 127)
								  Else
									  mstrExcelFormats(intTemp) = "0." & New String("0", objColumn.Size)
								  End If
							  Else
								  If objColumn.Size > 0 Then
									  mstrExcelFormats(intTemp) = "0"
								  Else
									  mstrExcelFormats(intTemp) = "General"
								  End If
							  End If
							
						  Case ColumnDataType.sqlDate
							  mstrExcelFormats(intTemp) = DateFormat
						  Case Else
							  mstrExcelFormats(intTemp) = "@"
					  End Select
					
				  Else
					  'MH20010307
					  objExpr = New clsExprExpression(SessionInfo)
					
					  objExpr.ExpressionID = CInt(objRow("ColExprID"))
					  objExpr.ConstructExpression()
					
					  'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
					  ' when all we need is the expression return type (time saving measure).
					  objExpr.ValidateExpression(True) 'MH20010508
					
					  objColumn.TableID = objExpr.BaseTableID
					  objColumn.TableName = objExpr.BaseTableName
					  objColumn.DataType = ColumnDataType.sqlNumeric
									
					  Select Case objExpr.ReturnType
						  Case ExpressionValueTypes.giEXPRVALUE_NUMERIC, ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
							  If objColumn.Decimals > 0 Then
								  If objColumn.Decimals > 127 Then
									  mstrExcelFormats(intTemp) = "0." & New String("0", 127)
								  Else
									  mstrExcelFormats(intTemp) = "0." & New String("0", objColumn.Decimals)
								  End If
							  Else
								  If objColumn.Size > 0 Then
									  mstrExcelFormats(intTemp) = "0"
								  Else
									  mstrExcelFormats(intTemp) = "General"
								  End If
							  End If
							
						  Case ExpressionValueTypes.giEXPRVALUE_DATE, ExpressionValueTypes.giEXPRVALUE_BYREF_DATE
							  mstrExcelFormats(intTemp) = DateFormat
						  Case Else
							  mstrExcelFormats(intTemp) = "@"
					  End Select
					
					  'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					  objExpr = Nothing
					
				  End If
				
				  mcolColDetails.Add(objColumn, objColumn.ColType & CStr(objColumn.ID))
				

			  Next 

		  End With

		Catch ex As Exception
		  mstrErrorMessage = "Error whilst retrieving the details recordsets'." & vbCrLf & ex.Message
      Return False

		End Try
		
		Return True
		
	End Function
	
	Private Function GetRelationRecordsets() As Boolean
			
		Dim objRelation As clsMatchRelation
		Dim objColumn As DisplayColumn
		Dim objBreakdownCols As Collection
		Dim rsMatchReportsDetails As DataTable
		Dim rsMatchBreakdownColumns As DataTable
		Dim strTempSQL As String
				
    Try

		strTempSQL = "SELECT ASRSysMatchReportTables.*," & "       a.TableName as Table1Name, b.TableName as Table2Name " & "FROM   ASRSysMatchReportTables " & "JOIN   ASRSysTables a on ASRSysMatchReportTables.Table1ID = a.TableID " & "LEFT OUTER JOIN   ASRSysTables b on ASRSysMatchReportTables.Table2ID = b.TableID " & "WHERE  MatchReportID = " & CStr(mlngMatchReportID) & "ORDER BY ASRSysMatchReportTables.MatchRelationID"
		
		rsMatchReportsDetails = General.GetReadOnlyRecords(strTempSQL)
		If rsMatchReportsDetails.Rows.Count = 0 Then
			mstrErrorMessage = "Cannot load the table relation information for this definition."
			Return False
		End If
			
		mcolRelations = New Collection()
		ReDim mlngTableViews(2, 0)
		
		With rsMatchReportsDetails
			for each objRow as DataRow in rsMatchReportsDetails.Rows
				
				objRelation = New clsMatchRelation
				
				objRelation.Table1ID = CInt(objRow("Table1ID"))
				objRelation.Table1Name = objRow("Table1Name").ToString()
							
				If Not TablePermission( CInt(objRow("Table1ID"))) Then
					mstrErrorMessage = "You do not have permission to read the '" & objRow("Table1Name").ToString() & "' table either directly or through any views."
					GetRelationRecordsets = False
					Exit Function
				End If
				
				If  CInt(objRow("Table2ID")) > 0 Then
					If Not TablePermission( CInt(objRow("Table2ID"))) Then
						mstrErrorMessage = "You do not have permission to read the '" & objRow("Table2Name").ToString() & "' table either directly or through any views."
						GetRelationRecordsets = False
						Exit Function
					End If
				End If
				
				
				objRelation.Table1RealSource = gcoTablePrivileges.Item((objRelation.Table1Name)).RealSource
				AddToJoinArray(0, CInt(objRow("Table1ID")))
				
				objRelation.Table2ID =  CInt(objRow("Table2ID"))
				If objRelation.Table2ID > 0 Then
					objRelation.Table2Name = objRow("Table2Name").ToString()
					objRelation.Table2RealSource = gcoTablePrivileges.Item((objRelation.Table2Name)).RealSource
					AddToJoinArray(0,  CInt(objRow("Table2ID")))
				End If
				
				objRelation.RequiredExprID =  CInt(objRow("RequiredExprID"))
				objRelation.PreferredExprID =  CInt(objRow("PreferredExprID"))
				objRelation.MatchScoreID =  CInt(objRow("MatchScoreExprID"))
											
				strTempSQL = "SELECT   ASRSysMatchReportBreakdown.*, " & "         ASRSysTables.TableID, ASRSysTables.TableName, " & "         ASRSysColumns.ColumnName, ASRSysColumns.DataType " & "FROM     ASRSysMatchReportBreakdown " & "JOIN     ASRSysMatchReportTables " & "ON       ASRSysMatchReportBreakdown.MatchRelationID = ASRSysMatchReportTables.MatchRelationID " & "LEFT OUTER JOIN ASRSysColumns " & "ON       ASRSysMatchReportBreakdown.ColExprID = ASRSysColumns.ColumnID AND ASRSysMatchReportBreakdown.ColType = 'C' " & "LEFT OUTER JOIN ASRSysTables " & "ON       ASRSysColumns.TableID = ASRSysTables.TableID " & "WHERE    ASRSysMatchReportBreakdown.MatchReportID = " & CStr(mlngMatchReportID) & " " & "AND      ASRSysMatchReportTables.Table1ID = " & CStr(objRelation.Table1ID) & " " & "ORDER BY ColSequence"			
				rsMatchBreakdownColumns = General.GetReadOnlyRecords(strTempSQL)
				
				objBreakdownCols = New Collection
				
				With rsMatchBreakdownColumns
					for each objBreakdownRow as DataRow In .Rows
						objColumn = New DisplayColumn
						
						objColumn.ColType = objBreakdownRow("ColType").ToString()
						objColumn.ID = cint(objBreakdownRow("ColExprID"))
						objColumn.Size = cint(objBreakdownRow("ColSize"))
						objColumn.Decimals = cint(objBreakdownRow("ColDecs"))
						objColumn.Heading = objBreakdownRow("ColHeading").ToString()
						'objcolumn.Sequence = !ColSequence
						'objColumn.SortSeq = !SortOrderSeq
						'objColumn.SortDir = !SortOrderDirection
						objColumn.Use1000Separator = DoesColumnUseSeparators(cint(objBreakdownRow("ColExprID")))
						
						If objColumn.ColType = "C" Then
							objColumn.DataType = cint(objBreakdownRow("DataType"))
							objColumn.TableID = cint(objBreakdownRow("TableID"))
							objColumn.TableName = objBreakdownRow("TableName").ToString()
							objColumn.Name = objBreakdownRow("ColumnName").ToString()
						Else
							objColumn.DataType = ColumnDataType.sqlNumeric
						End If
						
						objBreakdownCols.Add(objColumn, objColumn.ColType & CStr(objColumn.ID))
						
					next 
				End With
				
				objRelation.BreakdownColumns = objBreakdownCols
				mcolRelations.Add(objRelation, "T" & CStr(objRelation.Table1ID))
				
				'UPGRADE_NOTE: Object objBreakdownCols may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objBreakdownCols = Nothing
				'UPGRADE_NOTE: Object objRelation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objRelation = Nothing
			
			Next 
		End With

    Catch ex As Exception
		  mstrErrorMessage = "Error whilst retrieving the relation recordsets" & vbCrLf & ex.Message
      Return False

    End Try
		
		Return True
	
		
	End Function
	
	Private Function GetDataRecordset(plngTableID As Integer, plngRecordID As Integer) As Boolean
		
		Dim rsTemp As DataTable
		Dim strReportingStructure As String
				
    Try

		  fOK = True
		
		  mstrTable1RealSource = gcoTablePrivileges.Item(mstrTable1Name).RealSource
		  If mlngTable2ID > 0 Then
			  mstrTable2RealSource = gcoTablePrivileges.Item(mstrTable2Name).RealSource
		  End If
		  ReDim alngSourceTables(2, 0)
		
		  If fOK Then fOK = GenerateSQLWhere(plngTableID, plngRecordID)
		  If fOK Then fOK = GenerateSQLMatchScore
		  If fOK Then fOK = GenerateSQLSelect
		  If fOK Then fOK = GenerateSQLJoin
		  If fOK Then fOK = GenerateSQLOrderBy
		  If fOK Then fOK = General.UDFFunctions(mastrUDFsRequired, True)
		
		  If fOK Then
			  mstrTempTableName = GetTempTable
		  End If
		
		  If fOK = False Then
			  Exit Function
		  End If
		
		  If mlngMatchReportType = MatchReportType.mrtNormal Then
			  'MH20050104 Fault 9550
			  '    mstrSQL = "SELECT ID FROM " & mstrTable1Name &
			  mstrSQL = "SELECT ID FROM " & mstrTable1RealSource & IIf(mstrTable1Where <> vbNullString, " WHERE " & mstrTable1Where, vbNullString)
		  Else
			  'MH20050104 Fault 9550
			  '    mstrSQL = "SELECT ID FROM " & mstrTable2Name &
			  mstrSQL = "SELECT ID FROM " & mstrTable2RealSource & IIf(mstrTable2Where <> vbNullString, " WHERE " & mstrTable2Where, vbNullString)
		  End If
		
		  rsTemp = General.GetReadOnlyRecords(mstrSQL)
			
		  for each objRow as DataRow in rsTemp.Rows
			
			  If fOK Then
				
				  'Reporting Structure
				  If mlngMatchReportType <> MatchReportType.mrtNormal Then
					  strReportingStructure = GetReportingStructure(cint(objRow(0)))
				  End If
				
				
				  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLJoin(T0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLSelect(T0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				  mstrSQL = "INSERT INTO [" & _login.Username & "].[" & mstrTempTableName & "]" & " SELECT " & IIf(mlngNumRecords > 0, "TOP " & CStr(mlngNumRecords) & " ", vbNullString) & mcolSQLSelect.Item("T0") & vbCrLf & " FROM " & mstrTable1RealSource & vbCrLf & mcolSQLJoin.Item("T0")
				
				  Select Case mlngMatchReportType
					  Case MatchReportType.mrtNormal
						  mstrSQL = mstrSQL & " WHERE " & mstrTable1RealSource & ".ID = " & objRow(0).ToString() & vbCrLf
					  Case MatchReportType.mrtSucession
						  mstrSQL = mstrSQL & " WHERE " & mstrTable1RealSource & ".ID = " & GetJobTableID(CInt(objRow(0))) & vbCrLf
					  Case MatchReportType.mrtCareer
						  mstrSQL = mstrSQL & " WHERE " & mstrTable2RealSource & ".ID = " & objRow(0).ToString() & vbCrLf
				  End Select
				
				  If mstrSQLWhere <> vbNullString Then
					  mstrSQL = mstrSQL & " AND " & mstrSQLWhere & vbCrLf
				  End If
				
				  If strReportingStructure <> vbNullString Then
					  mstrSQL = mstrSQL & " AND " & strReportingStructure & vbCrLf
				  End If
				
				
				  If mlngMatchReportType = MatchReportType.mrtNormal Then
					  mstrSQL = mstrSQL & mstrSQLGroupBy & vbCrLf & IIf(mstrTable2Where <> vbNullString, " HAVING " & mstrTable2Where & vbCrLf, "")
					
					  If mblnScoreCheck Then
						  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLMatchScore(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  mstrSQL = mstrSQL & IIf(mstrTable2Where <> vbNullString, " AND ", " HAVING ") & mcolSQLMatchScore.Item("T0") & IIf(mlngScoreMode = 0, " >= ", " <= ") & CStr(mlngScoreLimit) & vbCrLf
					  End If
					
				  Else
					  mstrSQL = mstrSQL & mstrSQLGroupBy & vbCrLf & IIf(mstrTable1Where <> vbNullString, " HAVING " & mstrTable1Where & vbCrLf, "")
					
					  If mblnScoreCheck Then
						  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLMatchScore(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  mstrSQL = mstrSQL & IIf(mstrTable1Where <> vbNullString, " AND ", " HAVING ") & mcolSQLMatchScore.Item("T0") & IIf(mlngScoreMode = 0, " >= ", " <= ") & CStr(mlngScoreLimit) & vbCrLf
					  End If
					
				  End If
				
				  'MH20030606
				  'Still need order in case we are doing TOP X records
				  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLMatchScore(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				  mstrSQL = mstrSQL & " ORDER BY " & mcolSQLMatchScore.Item("T0") & IIf(mlngScoreMode = 0, " DESC", vbNullString)
				
				  DB.ExecuteSql(mstrSQL)
				
			  End If
		
			  If mstrErrorMessage <> vbNullString Then
				
				  'MH20060103 Bodge fix to ignore warning about nulls...
				  If mstrErrorMessage <> "Warning: Null value is eliminated by an aggregate or other SET operation." Then
					  GetDataRecordset = False
					  Exit Function
				  End If
				
			  End If
					
		  Next 
		
		  fOK = General.UDFFunctions(mastrUDFsRequired, False)
		
		  Return fOK

  Catch ex As Exception
        mstrErrorMessage = "Error retrieving data" & vbCrLf & ex.Message
        Return False
  End Try
    		
		
	End Function

	Public Function GetRecordsetBreakdown(lngTableID As Integer, lngRecord1ID As Integer, lngRecord2ID As Integer) As string
		
		Dim objRelation As clsMatchRelation
    Dim sBreakdownSQL as String
		
		objRelation = mcolRelations.Item("T" & CStr(lngTableID))
			
		'MH20030909
		'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLSelect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBreakdownSQL = "SELECT " & mcolSQLSelect.Item("T" & CStr(lngTableID)) & vbCrLf & "FROM " & objRelation.Table1RealSource & vbCrLf
		
		If lngTableID = mlngTable1ID Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLJoin(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sBreakdownSQL = sBreakdownSQL & mcolSQLJoin.Item("T" & CStr(lngTableID)) & " WHERE " & objRelation.Table1RealSource & ".ID = " & CStr(lngRecord1ID)
			If objRelation.Table2ID > 0 Then
				sBreakdownSQL = sBreakdownSQL & " AND " & objRelation.Table2RealSource & ".ID = " & CStr(lngRecord2ID)
			End If
			
			
		Else
			If objRelation.Table2ID > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLJoin(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sBreakdownSQL = sBreakdownSQL & mcolSQLJoin.Item("T" & CStr(lngTableID)) & " AND " & objRelation.Table1RealSource & ".ID_" & CStr(mlngTable1ID) & " = " & CStr(lngRecord1ID) & " AND " & objRelation.Table2RealSource & ".ID_" & CStr(mlngTable2ID) & " = " & CStr(lngRecord2ID)
			End If
			
			sBreakdownSQL = sBreakdownSQL & " WHERE " & objRelation.Table1RealSource & ".ID_" & CStr(mlngTable1ID) & " = " & CStr(lngRecord1ID)
			If frmBreakDown.chkAllRecords.Checked = True And mlngTable2ID > 0 Then
				sBreakdownSQL = sBreakdownSQL & " OR " & objRelation.Table2RealSource & ".ID_" & CStr(mlngTable2ID) & " = " & CStr(lngRecord2ID)
			End If

    End If

		'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLOrderBy(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBreakdownSQL = sBreakdownSQL & " ORDER BY " & mcolSQLOrderBy.Item("T" & CStr(lngTableID))
		
		return sBreakdownSQL
		
	End Function
	
	Private Function GenerateSQLMatchScore() As Boolean
		
		Dim objColumn As Column
		Dim mobjColumnPrivileges As CColumnPrivileges
		Dim mobjTableView As TablePrivilege
		Dim objRelation As clsMatchRelation
		
		Dim blnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim iLoop1 As Short
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean
		Dim lngScoreExpr As Integer
		Dim strOutput As String
		
		Dim pintLoop As Short
		Dim pstrColumnCode As String
		Dim pstrColumnCount As String
		Dim pstrSource As String
		Dim pintNextIndex As Short
		Dim strRealSource1 As String
		Dim strRealSource2 As String
		
		Dim sFilterCode As String
		Dim sCalcCode As String
		Dim objCalcExpr As clsExprExpression
		Dim lngCount As Integer
		
		
		pstrColumnCode = vbNullString
		pstrColumnCount = vbNullString
		lngCount = 0
		mcolSQLMatchScore = New Collection
		
		For	Each objRelation In mcolRelations
			
			sCalcCode = vbNullString
			
			If objRelation.MatchScoreID > 0 Then
				
				objCalcExpr = New clsExprExpression(SessionInfo)

				
				blnOK = objCalcExpr.Initialise(objRelation.Table1ID, (objRelation.MatchScoreID), ExpressionTypes.giEXPR_MATCHSCOREEXPRESSION, ExpressionValueTypes.giEXPRVALUE_LOGIC, objRelation.Table2ID)
        blnOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.MatchScoreID, ExpressionTypes.giEXPR_MATCHSCOREEXPRESSION, ExpressionValueTypes.giEXPRVALUE_LOGIC)



        If blnOK Then
					blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, mastrUDFsRequired, True, False, mvarPrompts)
				End If
				
				If blnOK Then
					' Add the required views to the JOIN code.
					For iLoop1 = 1 To UBound(alngSourceTables, 2)
						AddToJoinArray(alngSourceTables(1, iLoop1), alngSourceTables(2, iLoop1))
					Next iLoop1
				Else
					' Permission denied on something in the calculation.
					mstrErrorMessage = "You do not have permission to use the match score expression."
					GenerateSQLMatchScore = False
					Exit Function
				End If
				'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objCalcExpr = Nothing
				
				
				'      sCalcCode = GetCalcCode(objRelation.Table1ID, objRelation.Table2ID, objRelation.MatchScoreID, giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_LOGIC)
				'      If sCalcCode = vbNullString Then
				'        ' Permission denied on something in the calculation.
				'        mstrErrorMessage = "You do not have permission to use a match score calculation."
				'        GenerateSQLMatchScore = False
				'        Exit Function
				'      End If
						
				
				If sCalcCode <> vbNullString Then
					sCalcCode = "isnull(" & sCalcCode & ",0)"
				End If
				
			Else
				'If objRelation.Table2ID <> mlngTable2ID Then
				'sCalcCode = "case when " & objRelation.Table2RealSource & ".ID > 0 then 100 else 0 end"
				sCalcCode = "100"
				'End If
				
			End If
			
			If sCalcCode <> vbNullString Then
				
				objCalcExpr = New clsExprExpression(SessionInfo)
				
				If objRelation.PreferredExprID > 0 Then
					
					'If objRelation.Table2ID > 0 Then
					'  blnOK = objCalcExpr.Initialise(objRelation.Table2ID, objRelation.PreferredExprID, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table1ID)
					'Else
					blnOK = objCalcExpr.Initialise((objRelation.Table1ID), (objRelation.PreferredExprID), ExpressionTypes.giEXPR_MATCHJOINEXPRESSION, ExpressionValueTypes.giEXPRVALUE_LOGIC, objRelation.Table2ID)
					'End If
					
					If blnOK Then
						blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sFilterCode, mastrUDFsRequired, True, False, mvarPrompts)
					End If
					
					If blnOK Then
						' Add the required views to the JOIN code.
						For iLoop1 = 1 To UBound(alngSourceTables, 2)
							AddToJoinArray(alngSourceTables(1, iLoop1), alngSourceTables(2, iLoop1))
						Next iLoop1
					Else
						' Permission denied on something in the calculation.
						mstrErrorMessage = "You do not have permission to use the required match expression."
						GenerateSQLMatchScore = False
						Exit Function
					End If
					
					If sFilterCode <> vbNullString Then
						'sCalcCode = " CASE WHEN " & sFilterCode & " THEN " & sCalcCode & " ELSE 0 END "
						sCalcCode = " CASE WHEN " & sFilterCode & " = 1 THEN " & sCalcCode & " ELSE 0 END "
						'Stop
					End If
					'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objCalcExpr = Nothing
					
				End If
				
				'End If
				
				
				strRealSource1 = objRelation.Table1RealSource
				
				If objRelation.Table2ID > 0 Then
					strRealSource2 = objRelation.Table2RealSource
					
					sCalcCode = "case when " & strRealSource1 & ".ID > 0 and " & strRealSource2 & ".ID > 0 then " & sCalcCode & " else 0 end"
				End If
				
				If objRelation.Table1ID = mlngTable1ID Then
					pstrColumnCount = pstrColumnCount & IIf(pstrColumnCount <> vbNullString, "+", "") & "1"
					
					pstrColumnCode = pstrColumnCode & IIf(pstrColumnCode <> vbNullString, "+", "") & "max(cast(" & sCalcCode & " as float))"
					
				ElseIf objRelation.Table2ID = 0 And mcolRelations.Count() = 1 Then 
					pstrColumnCount = pstrColumnCount & IIf(pstrColumnCount <> vbNullString, "+", "") & "1"
					
					pstrColumnCode = pstrColumnCode & IIf(pstrColumnCode <> vbNullString, "+", "") & "cast(sum(" & sCalcCode & ") as float)"
					
				Else
					pstrColumnCount = pstrColumnCount & IIf(pstrColumnCount <> vbNullString, "+", "") & "count(distinct " & strRealSource1 & ".ID)"
					
					'pstrColumnCode = pstrColumnCode & _
					'IIf(pstrColumnCode <> vbNullString, "+", "") & _
					'"cast((sum(" & sCalcCode & ") * count(distinct " & strRealSource1 & ".ID) / " & _
					'"case when sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) = 0 then 1 else sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) end) as float)"
					pstrColumnCode = pstrColumnCode & IIf(pstrColumnCode <> vbNullString, "+", "") & "cast((sum(" & sCalcCode & ") * count(distinct " & strRealSource1 & ".ID) / " & "case when sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) = 0 then 1 else cast(sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) as float) end) as float)"
					
				End If
				
				mcolSQLMatchScore.Add(sCalcCode, "T" & CStr(objRelation.Table1ID))
				
			End If
			
		Next objRelation
		
		If pstrColumnCount <> "1" And mlngTable2ID > 0 Then
			strOutput = "((" & pstrColumnCode & ") / " & "case when " & pstrColumnCount & " = 0 then 1 else " & pstrColumnCount & " end)"
		Else
			strOutput = pstrColumnCode
		End If
		mcolSQLMatchScore.Add(strOutput, "T0")
		
		GenerateSQLMatchScore = True
		
	End Function
		
	Private Function GenerateSQLSelect() As Boolean
		
		Dim objRelation As clsMatchRelation
		
		'UPGRADE_NOTE: Object mcolSQLSelect may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolSQLSelect = Nothing
		mcolSQLSelect = New Collection
		
		'UPGRADE_NOTE: Object mcolSQLOrderBy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolSQLOrderBy = Nothing
		mcolSQLOrderBy = New Collection
		
		mstrErrorMessage = vbNullString
		GenerateSQLSelect = False
		
		mstrSQLGroupBy = vbNullString
		
		GetSelectStatement(mcolColDetails, 0, "")
		If mstrErrorMessage <> vbNullString Then
			Exit Function
		End If
		
		
		For	Each objRelation In mcolRelations
			GetSelectStatement((objRelation.BreakdownColumns), (objRelation.Table1ID), (objRelation.Table1RealSource))
			If mstrErrorMessage <> vbNullString Then
				Exit Function
			End If
			
			'mcolSQLSelect.Add strOutput, "T" & CStr(objRelation.Table1ID)
		Next objRelation
		
		GenerateSQLSelect = True
		
	End Function
		
	Private Function GetSelectStatement(ByRef colColumns As Collection, ByRef lngTableID As Integer, ByRef strTable1RealSource As String) As String
		
		Dim objColumn As DisplayColumn
		Dim mobjColumnPrivileges As CColumnPrivileges
		Dim mobjTableView As TablePrivilege
		
		Dim blnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim pblnNoSelect As Boolean
		Dim pstrColumnCode As String
		Dim pstrSource As String
		Dim strRealSource As String
		Dim blnBooleanColumn As Boolean
		
		Dim strSQLSelect As String
		Dim strSQLOrderBy As String
		Dim strOrderColumn As String
		
		strSQLOrderBy = vbNullString
		If strTable1RealSource <> vbNullString Then
			strSQLOrderBy = "case when " & strTable1RealSource & ".ID is null then 1 else 0 end"
		End If
		
		
		' Set flags with their starting values
		blnOK = True
		pblnNoSelect = False

    Try
		
		  strSQLSelect = vbNullString
		  pstrColumnCode = vbNullString
		
		  For	Each objColumn In colColumns

        if Not objColumn.Name = "ID" then
		      DisplayColumns.Add(objColumn)
          ReportDataTable.Columns.Add( string.format("{0}_{1}", objColumn.TableName, objColumn.Name), GetType(String))
        End If
			
			  ' If its a COLUMN then...
			  If objColumn.ColType = "C" Then
				
				  ' Check permission on that column
				  mobjColumnPrivileges = GetColumnPrivileges((objColumn.TableName))
				  blnBooleanColumn = mobjColumnPrivileges.Item((objColumn.Name)).DataType = ColumnDataType.sqlBoolean
				
				  pblnColumnOK = gcoTablePrivileges.Item((objColumn.TableName)).AllowSelect
				
				  'MH20040422 Fault 8267
				  'If pblnColumnOK Then
				  If pblnColumnOK Or objColumn.Name = "ID" Then
					  strRealSource = gcoTablePrivileges.Item((objColumn.TableName)).RealSource
					  pblnColumnOK = mobjColumnPrivileges.IsValid((objColumn.Name))
					  If pblnColumnOK Then
						  pblnColumnOK = mobjColumnPrivileges.Item((objColumn.Name)).AllowSelect
					  End If
				  End If
				
				  If pblnColumnOK Then
					  pstrColumnCode = strRealSource & "." & Trim(objColumn.Name)
					
					  AddToJoinArray(0, (objColumn.TableID))
				  Else
					
					  ' this column cannot be read direct. If its from a parent, try parent views
					  ' Loop thru the views on the table, seeing if any have read permis for the column
					
					  pblnNoSelect = True
					  Dim mstrViews(0) As Object
					  pstrColumnCode = vbNullString
					
					  For	Each mobjTableView In gcoTablePrivileges.Collection
						  If (Not mobjTableView.IsTable) And (mobjTableView.TableID = objColumn.TableID) And (mobjTableView.AllowSelect) Then
							
							  pstrSource = mobjTableView.ViewName
							  strRealSource = gcoTablePrivileges.Item(pstrSource).RealSource
							
							  ' Get the column permission for the view
							  mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
							
							  ' If we can see the column from this view
							  If mobjColumnPrivileges.IsValid((objColumn.Name)) Then
								  If mobjColumnPrivileges.Item((objColumn.Name)).AllowSelect Then
									  pstrColumnCode = pstrColumnCode & " WHEN NOT " & pstrSource & "." & objColumn.Name & " IS NULL THEN " & pstrSource & "." & objColumn.Name
									  AddToJoinArray(1, (mobjTableView.ViewID))
								  End If
							  End If
						  End If
						
					  Next mobjTableView
					
					  'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					  mobjTableView = Nothing
					
					  ' Does the user have select permission thru ANY views ?
					  ' If we cant see a column, then get outta here
					  If pstrColumnCode = vbNullString Then
						  strSQLSelect = vbNullString
						  mstrErrorMessage = "You do not have permission to see the column '" & objColumn.Name & "' either directly or through any views."
              Return False
						
					  Else
						  pstrColumnCode = "CASE" & pstrColumnCode & " ELSE NULL END"
						
					  End If
					
					  If Not blnOK Then
						  strSQLSelect = vbNullString
						  Exit Function
					  End If
					
				  End If
				
				
				  'MH20040422 Fault 8285
				  'If mobjColumnPrivileges.Item(objColumn.ColumnName).DataType = sqlBoolean Then
				  If blnBooleanColumn Then
					  pstrColumnCode = "(case when " & pstrColumnCode & " = 1 then 'Y' else 'N' end)"
				  End If
				
				  If lngTableID = 0 Then
					  mstrSQLGroupBy = mstrSQLGroupBy & IIf(mstrSQLGroupBy <> vbNullString, ", ", "") & pstrColumnCode
				  End If
				
				  strOrderColumn = pstrColumnCode
				
				  'pstrColumnCode = pstrColumnCode & " AS '" & objColumn.TableName & objColumn.ColumnName & "'"
				  pstrColumnCode = pstrColumnCode & " AS [" & Replace(objColumn.Heading, "'", "''") & "]"
			  Else
				  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLMatchScore(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				  pstrColumnCode = vbCrLf & mcolSQLMatchScore.Item("T" & CStr(lngTableID))
				  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLMatchScore(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				  strOrderColumn = mcolSQLMatchScore.Item("T" & CStr(lngTableID))
				  pstrColumnCode = pstrColumnCode & " AS [Match_Score]"
				
			  End If
			
			  If mlngMatchReportType <> MatchReportType.mrtNormal Then
				  objColumn.SQL = pstrColumnCode
			  End If
			
			  strSQLSelect = strSQLSelect & IIf(strSQLSelect <> vbNullString, ", ", "") & pstrColumnCode
			
			  strSQLOrderBy = strSQLOrderBy & IIf(strSQLOrderBy <> vbNullString, ", ", "") & strOrderColumn
			
		  Next objColumn
		
		  If lngTableID = 0 And mstrSQLGroupBy <> vbNullString Then
			  mstrSQLGroupBy = "GROUP BY " & mstrSQLGroupBy & vbCrLf
		  End If
		
		  mcolSQLSelect.Add(strSQLSelect, "T" & CStr(lngTableID))
		  mcolSQLOrderBy.Add(strSQLOrderBy, "T" & CStr(lngTableID))
		
		  Return True

    Catch ex As Exception
		  fOK = False
		  mstrErrorMessage = "Error whilst generating SQL Select statement." & vbCrLf & ex.Message
      Return False

    End Try
		
	End Function
	
	Private Function GenerateSQLJoin() As Boolean
			
		Dim pobjTableView As TablePrivilege
		Dim objRelation As clsMatchRelation
		Dim objCalcExpr As clsExprExpression
		
		Dim strOutputMain As String
		Dim strOutputBaseBreakdown As String
		Dim strOutputChildBreakdown As String
		Dim strOutputGrade As String
		
		Dim pintLoop As Short
		Dim pintLoop1 As Short
		Dim sCalcCode As String
		Dim sTemp As String = ""
		Dim strRealSource As String
		Dim blnFound As Boolean
		
		Dim strESelect As String
		Dim strEJoin As String
		Dim strPSelect As String
		Dim strPJoin As String
		Dim strViewIDs As String
		Dim strArray() As String
		Dim lngIndex As Integer
		
		Dim blnChildOf1 As Boolean
		Dim blnChildOf2 As Boolean
		
		
		mcolSQLJoin = New Collection
		strOutputMain = vbNullString
		
    Try

		  If mlngTable2ID > 0 Then
			  strOutputBaseBreakdown = "CROSS JOIN " & mstrTable2RealSource
		  End If
				
		  If mlngMatchReportType <> MatchReportType.mrtNormal Then
			  strRealSource = gcoTablePrivileges.Item(gstrGradeTableName).RealSource
			
			  GetSelectAndJoinForColumn(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelGradeColumnName, strESelect, strEJoin, strViewIDs)
			  If strESelect = vbNullString Then
				  mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelGradeColumnName & "' either directly or through any views."
				  GenerateSQLJoin = False
				  Exit Function
			  End If
			
			  strArray = Split(strViewIDs, " ")
			  For lngIndex = 1 To UBound(strArray)
				  AddToJoinArray(1, CInt(strArray(lngIndex)))
			  Next 
			
			
			  GetSelectAndJoinForColumn(glngPostTableID, gstrPostTableName, gstrPostGradeColumnName, strPSelect, strPJoin, strViewIDs)
			  If strPSelect = vbNullString Then
				  mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrPostGradeColumnName & "' either directly or through any views."
				  GenerateSQLJoin = False
				  Exit Function
			  End If
			
			  strArray = Split(strViewIDs, " ")
			  For lngIndex = 1 To UBound(strArray)
				  AddToJoinArray(1, CInt(strArray(lngIndex)))
			  Next 
			
			  strOutputGrade = " LEFT OUTER JOIN " & strRealSource & " ASRSys" & gsPersonnelTableName & gstrGradeTableName & " ON (" & strESelect & ") = " & "ASRSys" & gsPersonnelTableName & gstrGradeTableName & "." & gstrGradeColumnName & vbCrLf & " LEFT OUTER JOIN " & strRealSource & " ASRSys" & gstrPostTableName & gstrGradeTableName & " ON (" & strPSelect & ") = " & "ASRSys" & gstrPostTableName & gstrGradeTableName & "." & gstrGradeColumnName & vbCrLf
			
		  End If
		
		
		  For pintLoop = 1 To UBound(mlngTableViews, 2)
			
			  ' Get the table/view object from the id stored in the array
			  If mlngTableViews(1, pintLoop) = 0 Then
				  pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
			  Else
				  pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
			  End If
			
			
			  If pobjTableView.TableID = mlngTable1ID Then
				
				  strOutputBaseBreakdown = strOutputBaseBreakdown & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrTable1RealSource & ".ID = " & pobjTableView.RealSource & ".ID" & vbCrLf
				
			  ElseIf pobjTableView.TableID = mlngTable2ID Then 
				
				  strOutputBaseBreakdown = strOutputBaseBreakdown & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID" & vbCrLf
				
			  Else
				
				  blnChildOf1 = IsAChildOf((pobjTableView.TableID), mlngTable1ID)
				  blnChildOf2 = IsAChildOf((pobjTableView.TableID), mlngTable2ID)
				
				  If blnChildOf1 And blnChildOf2 Then
					  mstrErrorMessage = "Cannot use the '" & pobjTableView.TableName & "' table as it is a child table of both the '" & mstrTable1Name & "' and the '" & mstrTable2Name & "' tables."
					  GenerateSQLJoin = False
					  Exit Function
					
				  ElseIf blnChildOf1 Then 
					
					  strOutputMain = strOutputMain & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrTable1RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable1ID) & vbCrLf
					
				  ElseIf blnChildOf2 Then 
					
					  blnFound = False
					  For	Each objRelation In mcolRelations
						  If objRelation.Table2ID = pobjTableView.TableID Then
							  blnFound = True
							  Exit For
						  End If
					  Next objRelation
					
					  sCalcCode = vbNullString
					
					  If blnFound Then
						
						  If objRelation.RequiredExprID > 0 Then
							  If objRelation.Table2ID > 0 Then
								  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLWhere(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  sCalcCode = sCalcCode & mcolSQLWhere.Item("T" & CStr(objRelation.Table2ID)) & " = 1 "
							  Else
								  'UPGRADE_WARNING: Couldn't resolve default property of object mcolSQLWhere(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  sCalcCode = sCalcCode & mcolSQLWhere.Item("T" & CStr(objRelation.Table1ID)) & " = 1 "
							  End If
							
						  End If
												
						  If objRelation.RequiredExprID > 0 Then
							  strOutputMain = strOutputMain & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & "(" & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable2ID) & vbCrLf & IIf(sCalcCode <> vbNullString, " AND " & sCalcCode, "") & ")" & vbCrLf
						  End If
						
						  If objRelation.PreferredExprID > 0 Then
							  objCalcExpr = New clsExprExpression(SessionInfo)
							  fOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.PreferredExprID, ExpressionTypes.giEXPR_MATCHJOINEXPRESSION, ExpressionValueTypes.giEXPRVALUE_LOGIC, objRelation.Table2ID)
							  If fOK Then
                  fOK = objCalcExpr.RuntimeFilterCode(sTemp, True, mastrUDFsRequired, False, mvarPrompts)
							  End If
							
							  If fOK Then
								  For pintLoop1 = 1 To UBound(alngSourceTables, 2)
									  AddToJoinArray(Val(CStr(alngSourceTables(1, pintLoop1))), Val(CStr(alngSourceTables(2, pintLoop1))))
								  Next 
							  Else
								  mstrErrorMessage = "You do not have permission to use the preferred match expression."
								  Exit Function
							  End If
							  'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							  objCalcExpr = Nothing
							
							  sCalcCode = sCalcCode & IIf(sCalcCode <> vbNullString, " AND ", vbNullString) & sTemp & " = 1 "
							
						  End If
											
						  If objRelation.RequiredExprID = 0 Then
							  strOutputMain = strOutputMain & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & "(" & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable2ID) & vbCrLf & IIf(sCalcCode <> vbNullString, " AND " & sCalcCode, "") & ")" & vbCrLf
						  End If
						
						  If sCalcCode <> vbNullString Then
							  If objRelation.Table1ID <> mlngTable1ID Then
								  'MH20030909
								  strOutputChildBreakdown = "FULL OUTER JOIN " & objRelation.Table2RealSource & " ON " & sCalcCode
								  mcolSQLJoin.Add(strOutputChildBreakdown, "T" & CStr(objRelation.Table1ID))
								  'mcolSQLJoin.Add sCalcCode, "T" & CStr(objRelation.Table1ID)
							  End If
						  End If
						
					  End If
					
				  End If
			  End If
			
		  Next 
		
		
		  mcolSQLJoin.Add(strOutputBaseBreakdown & strOutputMain & strOutputGrade, "T0")
		  mcolSQLJoin.Add(strOutputBaseBreakdown & strOutputGrade, "T" & CStr(mlngTable1ID))		
		  Return True

    Catch ex As Exception
		  mstrErrorMessage = "Error in GenerateSQLJoin." & vbCrLf & ex.Message
      Return False

    End Try	
		
	End Function
	
	Private Function GenerateSQLWhere(plngTableID As Integer, plngRecordID As Integer) As Boolean
		
		Dim objRelation As clsMatchRelation
		Dim objCalcExpr As clsExprExpression
		Dim strPicklistFilterSelect As String
		Dim sCalcCode As String
		Dim pintLoop1 As Integer
		Dim strReportingStructure As String
		
		Dim lngTable1RecordID As Integer
		Dim lngTable2RecordID As Integer
		
		'UPGRADE_NOTE: Object mcolSQLWhere may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolSQLWhere = Nothing
		mcolSQLWhere = New Collection
		
		
		mstrTable1Where = vbNullString
		mstrTable2Where = vbNullString
		mstrSQLWhere = vbNullString
		
		
		'Single Record
		If plngRecordID > 0 Then
			If mlngMatchReportType = MatchReportType.mrtSucession Then
				If mlngTable1ID = glngPostTableID Then
					lngTable1RecordID = GetJobTableID(plngRecordID)
				ElseIf mlngTable2ID = glngPostTableID Then 
					lngTable2RecordID = GetJobTableID(plngRecordID)
				End If
			Else
				If mlngTable1ID = plngTableID Then
					lngTable1RecordID = plngRecordID
				Else
					lngTable2RecordID = plngRecordID
				End If
			End If
		End If
		
		
		strPicklistFilterSelect = GetPicklistFilterSelect(lngTable1RecordID, mlngTable1PickListID, mlngTable1FilterID)
		If fOK = False Then
			Exit Function
		End If
		If strPicklistFilterSelect <> vbNullString Then
			mstrTable1Where = mstrTable1Where & IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & mstrTable1RealSource & ".ID IN (" & strPicklistFilterSelect & ")"
		End If
		
		
		strPicklistFilterSelect = GetPicklistFilterSelect(lngTable2RecordID, mlngTable2PickListID, mlngTable2FilterID)
		If fOK = False Then
			Exit Function
		End If
		If strPicklistFilterSelect <> vbNullString Then
			mstrTable2Where = mstrTable2Where & IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & mstrTable2RealSource & ".ID IN (" & strPicklistFilterSelect & ")"
		End If
		
		
		For	Each objRelation In mcolRelations
			
			If objRelation.RequiredExprID > 0 Then
				objCalcExpr = New clsExprExpression(SessionInfo)
				'MH20030918 Fault 7005
				'      If objRelation.Table2ID > 0 Then
				'        fOK = objCalcExpr.Initialise(objRelation.Table2ID, objRelation.RequiredExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table1ID)
				'      Else
				'        fOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.RequiredExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC, 0)
				'      End If
				fOK = objCalcExpr.Initialise((objRelation.Table1ID), (objRelation.RequiredExprID), ExpressionTypes.giEXPR_MATCHWHEREEXPRESSION, ExpressionValueTypes.giEXPRVALUE_LOGIC, objRelation.Table2ID)
				
				If fOK Then
					fOK = objCalcExpr.RuntimeFilterCode(sCalcCode, True, mastrUDFsRequired, False, mvarPrompts)
				End If
				
				If fOK Then
					For pintLoop1 = 1 To UBound(alngSourceTables, 2)
						AddToJoinArray(Val(CStr(alngSourceTables(1, pintLoop1))), Val(CStr(alngSourceTables(2, pintLoop1))))
					Next 
				Else
					'mstrErrorMessage = objCalcExpr.ErrorMessage
					mstrErrorMessage = "You do not have permission to use the required match expression."
					Exit Function
				End If
				
				'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objCalcExpr = Nothing				
				
				If objRelation.Table1ID <> mlngTable1ID And objRelation.Table2ID > 0 Then
					'If mlngMatchReportType = mrtNormal Then
					'  mstrTable2Where = mstrTable2Where & _
					''    IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & _
					''    "count(distinct " & objRelation.Table1RealSource & ".ID) = " & _
					''    "count(distinct " & objRelation.Table2RealSource & ".ID)"
					'Else
					'  mstrTable1Where = mstrTable1Where & _
					''    IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & _
					''    "count(distinct " & objRelation.Table1RealSource & ".ID) = " & _
					''    "count(distinct " & objRelation.Table2RealSource & ".ID)"
					'End If
					If mlngMatchReportType = MatchReportType.mrtNormal Then
						mstrTable2Where = mstrTable2Where & IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & "count(" & objRelation.Table1RealSource & ".ID) = " & "count(" & objRelation.Table2RealSource & ".ID)"
					Else
						mstrTable1Where = mstrTable1Where & IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & "count(" & objRelation.Table1RealSource & ".ID) = " & "count(" & objRelation.Table2RealSource & ".ID)"
					End If
				Else
					mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", vbNullString) & "(" & sCalcCode & ") = 1 "
				End If
				
				If objRelation.Table2ID > 0 Then
					mcolSQLWhere.Add(sCalcCode, "T" & CStr(objRelation.Table2ID))
				Else
					mcolSQLWhere.Add(sCalcCode, "T" & CStr(objRelation.Table1ID))
				End If
				
			End If
			
		Next objRelation
		
		GenerateSQLWhere = True
		
	End Function
	
	Private Function GetMatchReportDefinition() As Boolean
			
		Dim rsTemp_Definition As DataTable
		Dim strSQL As String
		
    Try
   	
		strSQL = "SELECT ASRSYSMatchReportName.*, a.TableName as 'Table1Name', a.RecordDescExprID as 'Table1RecDescExprID', b.TableName as 'Table2Name', " & "b.RecordDescExprID as 'Table2RecDescExprID' " & "FROM ASRSYSMatchReportName " & "JOIN ASRSysTables a ON ASRSysMatchReportName.Table1ID = a.TableID " & "LEFT OUTER JOIN ASRSysTables b ON ASRSysMatchReportName.Table2ID = b.TableID " & "WHERE MatchReportID = " & CStr(mlngMatchReportID)
		
		rsTemp_Definition = DB.GetDataTable(strSQL)
	    
		With rsTemp_Definition
			
			If .Rows.Count = 0 Then
				GetMatchReportDefinition = False
				mstrErrorMessage = "Could not find specified definition !"
				Exit Function
			End If
			
      dim objRow = .Rows(0)

			Name = objRow("Name").ToString()
			mstrDescription = objRow("Description").ToString()
			mlngNumRecords = CInt(objRow("NumRecords"))
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mlngScoreMode = IIf(IsDbNull(objRow("ScoreMode")), 0, CInt(objRow("ScoreMode")))
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mblnScoreCheck = IIf(IsDbNull(objRow("ScoreCheck")), False, CInt(objRow("ScoreCheck")))
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mlngScoreLimit = IIf(IsDbNull(objRow("ScoreLimit")), 0,CInt( objRow("ScoreLimit")))
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mblnEqualGrade = IIf(IsDbNull(objRow("EqualGrade")), False, CInt(objRow("EqualGrade")))
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mblnReportingStructure = IIf(IsDbNull(objRow("ReportingStructure")), 0, CInt(objRow("ReportingStructure")))
			
			mlngTable1ID = CInt(objRow("Table1ID"))
			mstrTable1Name = objRow("Table1Name").ToString()
			mlngTable1RecDescExprID = CInt(objRow("Table1RecDescExprID"))
			mlngTable1AllRecords = CInt(objRow("Table1AllRecords"))
			mlngTable1PickListID = CInt(objRow("Table1Picklist"))
			mlngTable1FilterID = CInt(objRow("Table1Filter"))
			' Override filter and/or picklist if in Report pack mode
			If mlngTable1ID = glngPersonnelTableID And gblnReportPackMode Then
				mlngTable1FilterID = mlngOverrideFilterID
				mlngTable1PickListID = 0
			End If
			
			If Not TablePermission(CInt(objRow("Table1ID"))) Then
				mstrErrorMessage = "You do not have permission to read the '" & objRow("Table1Name").ToString() & "' table either directly or through any views."
				GetMatchReportDefinition = False
				Exit Function
			End If
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(objRow("PrintFilterHeader")) Then
				If objRow("PrintFilterHeader") Then
					If mlngTable1PickListID > 0 Then
						mstrRecordSelectionName = " (Base Table picklist: " & General.GetPicklistName(mlngTable1PickListID) & ")"
					ElseIf mlngTable1FilterID > 0 Then 
						mstrRecordSelectionName = " (Base Table filter: " & General.GetFilterName(mlngTable1FilterID) & ")"
					End If
				End If
			End If
			
			mlngTable2ID = CInt(objRow("Table2ID"))
			If mlngTable2ID > 0 Then
				mstrTable2Name = objRow("Table2Name").ToString()
				mlngTable2RecDescExprID = CInt(objRow("Table2RecDescExprID"))
				mlngTable2AllRecords = CInt(objRow("Table2AllRecords"))
				mlngTable2PickListID = CInt(objRow("Table2Picklist"))
				mlngTable2FilterID = CInt(objRow("Table2Filter"))
				
				If mlngTable2ID = glngPersonnelTableID And gblnReportPackMode Then
					mlngTable2FilterID = mlngOverrideFilterID
					mlngTable2PickListID = 0
				End If
				
				If Not TablePermission(CInt(objRow("Table2ID"))) Then
					mstrErrorMessage = "You do not have permission to read the '" & objRow("Table2Name").ToString() & "' table either directly or through any views."
					GetMatchReportDefinition = False
					Exit Function
				End If
			End If
			
			mbDefinitionOwner = (LCase(Trim(_login.Username)) = LCase(Trim(objRow("UserName").ToString())))
			
			'Change Output Options to Report Pack owning these Jobs if in Report Pack mode
      Dim lblnReportPackMode = false
			mblnPreviewOnScreen = IIf(lblnReportPackMode, mblnPreviewOnScreen, cbool(objRow("OutputPreview")))
			mblnOutputScreen = IIf(lblnReportPackMode, mblnOutputScreen, CBool(objRow("OutputScreen")))
			mlngOutputFormat = IIf(lblnReportPackMode, mlngOutputFormat, CInt(objRow("OutputFormat")))
			mblnOutputPrinter = IIf(lblnReportPackMode, mblnOutputPrinter, CBool(objRow("OutputPrinter")))
			mstrOutputPrinterName = IIf(lblnReportPackMode, mstrOutputPrinterName, objRow("OutputPrinterName").ToString())
			mblnOutputSave = IIf(lblnReportPackMode, mblnOutputSave, CBool(objRow("OutputSave")))
			mlngOutputSaveExisting = IIf(lblnReportPackMode, mlngOutputSaveExisting, CInt(objRow("OutputSaveExisting")))
			mblnOutputEmail = IIf(lblnReportPackMode, mblnOutputEmail, CBool(objRow("OutputEmail")))
			mlngOutputEmailAddr = IIf(lblnReportPackMode, mlngOutputEmailAddr, CInt(objRow("OutputEmailAddr")))
			mstrOutputEmailSubject = IIf(lblnReportPackMode, mstrOutputEmailSubject, objRow("OutputEmailSubject").ToString())
			mstrOutputEmailAttachAs = IIf(lblnReportPackMode, mstrOutputEmailAttachAs, objRow("OutputEmailAttachAs").ToString())
			mstrOutputFileName = IIf(lblnReportPackMode, mstrOutputFileName, objRow("OutputFilename").ToString())
			mlngOverrideFilterID = IIf(lblnReportPackMode, mlngOverrideFilterID, 0)
			mblnOutputRetainPivotOrChart = IIf(lblnReportPackMode, mblnOutputRetainPivotOrChart, 0)
			
			mblnPreviewOnScreen = (mblnPreviewOnScreen Or (mlngOutputFormat = OutputFormats.DataOnly And mblnOutputScreen))
			
		End With
		
		If frmBreakDown Is Nothing Then
			frmBreakDown = New frmMatchRunBreakDown
		End If
		frmBreakDown.lblTable1Name.Text = mstrTable1Name
		frmBreakDown.lblTable2Name.Text = mstrTable2Name
		
		Return IsRecordSelectionValid

    Catch ex As Exception
		  mstrErrorMessage = "Error whilst reading the definition !" & vbCrLf & ex.Message
      Return False

    End Try
		
	End Function
		
	Private Function IsRecordSelectionValid() As Boolean
		Dim iResult As RecordSelectionValidityCodes
		
		' Base Table First
		If mlngTable1FilterID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngTable1FilterID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrErrorMessage = "The base table filter used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrErrorMessage = "The base table filter used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not _login.IsSystemOrSecurityAdmin Then
						mstrErrorMessage = "The base table filter used in this definition has been made hidden by another user."
					End If
			End Select
		ElseIf mlngTable1PickListID > 0 Then 
			iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngTable1PickListID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrErrorMessage = "The base table picklist used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrErrorMessage = "The base table picklist used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not _login.IsSystemOrSecurityAdmin Then
						mstrErrorMessage = "The base table picklist used in this definition has been made hidden by another user."
					End If
			End Select
		End If
		
		If Len(mstrErrorMessage) = 0 Then
			' Criteria Table
			If mlngTable2FilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngTable2FilterID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorMessage = "The match table filter used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorMessage = "The match table filter used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not _login.IsSystemOrSecurityAdmin Then
							mstrErrorMessage = "The match table filter used in this definition has been made hidden by another user."
						End If
				End Select
			ElseIf mlngTable2PickListID > 0 Then 
				iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngTable2PickListID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorMessage = "The match table picklist used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorMessage = "The match table picklist used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not _login.IsSystemOrSecurityAdmin Then
							mstrErrorMessage = "The match table picklist used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If
		
		IsRecordSelectionValid = (Len(mstrErrorMessage) = 0)
		
	End Function
	
	Private Sub AddToJoinArray(lngTYPE As Integer, lngTableID As Integer)
		
		Dim lngIndex As Short
		
		If lngTYPE = 0 Then 'Table
			If lngTableID = mlngTable1ID Or lngTableID = mlngTable2ID Then
				Exit Sub
			End If
		End If
		
		For lngIndex = 1 To UBound(mlngTableViews, 2)
			If mlngTableViews(1, lngIndex) = lngTYPE And mlngTableViews(2, lngIndex) = lngTableID Then
				Exit Sub
			End If
		Next 
		
		If lngTableID = 0 Then
			Stop
		End If
		
		'Only get here if not already in array
		lngIndex = UBound(mlngTableViews, 2) + 1
		ReDim Preserve mlngTableViews(2, lngIndex)
		mlngTableViews(1, lngIndex) = lngTYPE
		mlngTableViews(2, lngIndex) = lngTableID
		
	End Sub
	
	
	Private Function GenerateSQLOrderBy() As Boolean
		
		Dim objColumn As DisplayColumn
		Dim lngIndex As Integer
		
		mstrSQLOrderBy = " ORDER BY "
		For lngIndex = 1 To mcolColDetails.Count()
			
			For	Each objColumn In mcolColDetails
				If objColumn.SortSeq = lngIndex Then
					mstrSQLOrderBy = mstrSQLOrderBy & IIf(lngIndex > 1, ", ", "") & "[" & objColumn.Heading & "]" & IIf(objColumn.SortDir = "D", " DESC", "")
				End If
			Next objColumn
			
		Next 
		
		GenerateSQLOrderBy = True
		
	End Function
		
	Public Function PopulateGridBreakdown(lngTableID As Integer) As Boolean
		
		Dim objRelation As clsMatchRelation
		
		objRelation = mcolRelations.Item("T" & CStr(lngTableID))
		
		If frmBreakDown Is Nothing Then
			frmBreakDown = New frmMatchRunBreakDown
		End If
		
		PopulateGridBreakdown = False
		If PopulateGrid((objRelation.BreakdownColumns), True) Then
			PopulateGridBreakdown = True
		End If
	
		
		frmBreakDown.chkAllRecords.Enabled = (objRelation.Table1ID <> mlngTable1ID)
		If objRelation.Table1ID = mlngTable1ID Then
			frmBreakDown.chkAllRecords.Checked = False
		End If
		
	
		'UPGRADE_NOTE: Object objRelation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objRelation = Nothing
		
	End Function
	
	
	Private Function PopulateGridMain() As Boolean
		
		PopulateGridMain = False		
		mstrSQL = "SELECT DISTINCT * FROM [" & _login.Username & "].[" & mstrTempTableName & "]" & vbCrLf & "WHERE not (ID1 is null) " & mstrSQLOrderBy
			
		If PopulateGrid(mcolColDetails, False) Then		
			PopulateGridMain = True				
		End If
		
	End Function
	

	Private Function PopulateGrid(ByRef colColumns As Collection, ByRef blnBreakdownOutput As Boolean) As Boolean
		
		Dim objColumn As Column
		Dim rsMatchReportsData As DataTable
		Dim strOutput As String
		Dim vData As Object
		Dim vDataTemp As Object
		Dim lngIndex As Integer
		Dim iCount As Short
		Dim iCount2 As Short
		
		Dim aryAddString As ArrayList
		
    Try

		  rsMatchReportsData = DB.GetDataTable(mstrSQL)
			
		  If rsMatchReportsData.Rows.Count = 0 Then
			  mstrErrorMessage = "No records meet selection criteria."
			  mblnNoRecords = True
			  fOK = False
			  Return False
		  End If
				
	
		  With rsMatchReportsData
								
			  If .Rows.Count > 0 Then
				  for each objRow as DataRow in .Rows
					
					  strOutput = vbNullString
     				aryAddString = New ArrayList()

					  'strOutput = .Fields(0).Value & vbTab & .Fields(1).Value
					  For lngIndex = 0 To .Columns.Count - 1
						
						  objColumn = colColumns.Item(lngIndex + 1)
						  'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  vData = IIf(IsDbNull(objRow(lngIndex)), vbNullString, objRow(lngIndex).ToString())
						  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  vData = Replace(vData, vbCr, "")
						  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  vData = Replace(vData, vbTab, "")
						
						  If objColumn.IsNumeric Then
							
							  If objColumn.Decimals > 0 Then
								  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  vData = VB6.Format(vData, "0." & New String("0", objColumn.Decimals))
							  Else
								  If objColumn.Size > 0 Then
									  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  If vData = "0" Then
										  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  vData = VB6.Format(vData, "0")
									  Else
										  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  vData = VB6.Format(vData, "#")
									  End If
								  End If
							  End If
							
							  If objColumn.Use1000Separator Then
								  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  vDataTemp = vData
								  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  vData = ""
								  iCount2 = 1
								  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  If InStr(1, vDataTemp, ".") > 0 Then
									  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  For iCount = InStr(1, vDataTemp, ".") - 1 To 1 Step -1
										  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  vData = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(vDataTemp, iCount, 1) & vData
										  iCount2 = iCount2 + 1
									  Next iCount
									  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  vData = vData & "." & Right(vDataTemp, Len(vDataTemp) - InStr(1, vDataTemp, "."))

								  Else
									  For iCount = Len(vDataTemp) To 1 Step -1
										  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  'UPGRADE_WARNING: Couldn't resolve default property of object vDataTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  vData = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(vDataTemp, iCount, 1) & vData
										  iCount2 = iCount2 + 1
									  Next iCount
								  End If
							  End If
							
						  End If
						
						  ' If its a date column, format it as dateformat
						  If objColumn.DataType = ColumnDataType.sqlDate Then
							  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							  vData = VB6.Format(vData, DateFormat)
						  End If
						
						  If objColumn.Size > 0 Then 'Size restriction
							  If objColumn.Decimals > 0 Then
								  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								  If InStr(vData, ".") > objColumn.Size Then
									  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  vData = Left(vData, objColumn.Size) & Mid(vData, InStr(vData, "."))
								  End If
								
							  Else
								  If Len(vData) > objColumn.Size Then
									  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  vData = Left(vData, objColumn.Size)
								  End If
								
							  End If
						  End If
						
						  'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						  strOutput = strOutput & IIf(lngIndex > 0, vbTab, "") & vData					
              aryAddString.Add(vData)
						
					  Next 
					
					  '          If mlngMatchReportType <> mrtNormal Then
					  '            strOutput = strOutput & _
					  ''                IIf(lngIndex > 0, vbTab, "") & "..."
					  '          End If
					
					  If Not blnBreakdownOutput Then
						  'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						  If Not IsDbNull(objRow(0)) And Not IsDbNull(objRow(1)) Then
							  frmBreakDown.AddToCrossRef(cint(objRow(0)), cint(objRow(1)))
						  End If
					  End If


            ' LOFTY - HACK ALERT! -- Breakdown for this record
            If 1 = 1 Then
              Dim childTableId = mcolRelations(1).Table1ID
              'cint(objRow(0)), cint(objRow(1))
              dim breakdownSQL = GetRecordsetBreakdown(childTableId, cint(objRow(0)), cint(objRow(1)))

              Dim breakdownData As DataTable
              Dim breakdownValue as String = ""
              breakdownData = DB.GetDataTable(breakdownSQL)
              With breakdownData
                for Each objBreakdown as DataRow in .Rows

'                  breakdownValue = breakdownValue & String.Format("{0} {1}", objBreakdown(0),  objBreakdown(.Columns.Count - 1)) & vbNewLine
'                  breakdownValue = breakdownValue & String.Format("{{""Competency"":""{0}"" {1}}}", objBreakdown(0),  objBreakdown(.Columns.Count - 1)) & vbNewLine
                  Dim competency As String = IIf(IsDBNull(objBreakdown("Competency")), "", objBreakdown("Competency")).ToString()
                  Dim minimumScore As Double = CDbl(IIf(IsDBNull(objBreakdown("MinScore")), 0, objBreakdown("MinScore")))
                  Dim preferredScore As Double = CDbl(IIf(IsDBNull(objBreakdown("PrefScore")), 0, objBreakdown("PrefScore")))
                  Dim actualScore As Double = CDbl(IIf(IsDBNull(objBreakdown("ActualScore")), 0, objBreakdown("ActualScore")))

                  breakdownValue &= IIf(Len(breakdownValue) > 0, ",", "") & _
                    String.Format("{{""Competency"":""{0}"", ""MinScore"":{1}, ""PrefScore"":{2}, ""ActualScore"":{3}}}" _
                    ,competency, minimumScore, preferredScore, actualScore)

                Next

                breakdownValue = IIf(Len(breakdownValue) > 0, "[" & breakdownValue & "]" , "") 
              End With

              ' Hack it into the last column
              strOutput = strOutput & IIf(lngIndex > 0, vbTab, "") & breakdownValue
              aryAddString.Add("")
              aryAddString.Add("")
              aryAddString.Add(breakdownValue)
            End If

            Data.Add(strOutput)
            AddItemToReportData(aryAddString)
								
				  Next 
			  End If
			
		  End With			
			
		  Return True	

      Catch ex As Exception
        Return False

    End Try		

  End Function
	
	Private Function GetPicklistFilterSelect(lngSingleID As Integer, lngPicklistID As Integer, lngFilterID As Integer) As String
		
		Dim rsTemp As DataTable
			
    Try
		
		  If lngSingleID > 0 Then
			  GetPicklistFilterSelect = CStr(lngSingleID)
			
		  ElseIf lngPicklistID > 0 Then 
			
			  mstrErrorMessage = IsPicklistValid(lngPicklistID)
			  If mstrErrorMessage <> vbNullString Then
				  fOK = False
          Exit Function
			  End If			
			
			  'Get List of IDs from Picklist
			  rsTemp = General.GetReadOnlyRecords("EXEC sp_ASRGetPickListRecords " & CStr(lngPicklistID))
			  fOK = rsTemp.Rows.Count > 0
			
			  If Not fOK Then
				  mstrErrorMessage = "The base table picklist contains no records."
			  Else
				  GetPicklistFilterSelect = vbNullString
				  For each objRow as DataRow in rsTemp.Rows
					  GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & objRow(0).ToString()
          Next
			  End If
					
		  ElseIf lngFilterID > 0 Then 
			
			  mstrErrorMessage = IsFilterValid(lngFilterID)
			  If mstrErrorMessage <> vbNullString Then
				  fOK = False
				  Return fOK
			  End If
			
			  'Get list of IDs from Filter
			  fOK = FilteredIDs(lngFilterID, GetPicklistFilterSelect, mastrUDFsRequired, mvarPrompts)
					
			  If Not fOK Then
				  ' Permission denied on something in the filter.
				  mstrErrorMessage = "You do not have permission to use the '" & General.GetFilterName(lngFilterID) & "' filter."
			  End If
			
		  End If
		
    Catch ex As Exception
		  mstrErrorMessage = "Error processing record selection"
		  fOK = False

    End Try
		
	End Function
	
	Private Function InitialiseFormBreakdown() As Boolean
		
		Dim objRelation As clsMatchRelation
		
		If frmBreakDown Is Nothing Then
			frmBreakDown = New frmMatchRunBreakDown
		End If

		
		With frmBreakDown
			.Loading = True
			
			.ParentForm_Renamed = Me
			.lblTable1Name.Text = mstrTable1Name & " :"
			.Table1RecDescExprID = mlngTable1RecDescExprID
				
			
			With .cboRelation
				.Items.Clear()
				For	Each objRelation In mcolRelations
					.Items.Add(objRelation.Table1Name)
				Next objRelation
				If .Items.Count > 0 Then
					.SelectedIndex = 0
				End If
			End With
			
			.Loading = False
			
		End With
		
		Return True
		
	End Function

	Private Function OutputReport(blnPrompt As Boolean) As Boolean
		
		Dim objOutput As clsOutputRun
		Dim objColumn As DisplayColumn
		
		objOutput = New clsOutputRun
			
		'UPGRADE_WARNING: Couldn't resolve default property of object objOutput.SetOptions(blnPrompt, mlngOutputFormat, mblnOutputScreen, mblnOutputPrinter, mstrOutputPrinterName, mblnOutputSave, mlngOutputSaveExisting, mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, mstrOutputEmailAttachAs, mstrOutputFileName, False, mblnPreviewOnScreen, mstrOutputTitlePage, mstrOutputReportPackTitle, mstrOutputOverrideFilter, mblnOutputTOC, mblnOutputCoverSheet, mlngOverrideFilterID, mblnOutputRetainPivotOrChart). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If objOutput.SetOptions(blnPrompt, mlngOutputFormat, mblnOutputScreen, mblnOutputPrinter, mstrOutputPrinterName, mblnOutputSave, mlngOutputSaveExisting, mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, mstrOutputEmailAttachAs, mstrOutputFileName) Then
			
			objOutput.PageTitles = False
					
			objOutput.SizeColumnsIndependently = True
			If objOutput.GetFile Then
				objOutput.AddPage(Name & mstrRecordSelectionName, mstrTable1Name)
				
				For	Each objColumn In mcolColDetails
					'Ignore hidden columns
					If objColumn.Heading <> vbNullString And objColumn.Hidden = False Then
						objOutput.AddColumn((objColumn.Heading), (objColumn.DataType), (objColumn.Decimals), objColumn.Use1000Separator)
					End If
				Next objColumn

				' Implement differently. Client side pdf output?
				'objOutput.DataGrid(grdOutput)
							
				objOutput.Complete()
				
			End If
			
			mstrErrorMessage = objOutput.ErrorMessage
			fOK = (mstrErrorMessage = vbNullString)
			
		Else
			mstrErrorMessage = objOutput.ErrorMessage
			fOK = (mstrErrorMessage = vbNullString)
			
		End If
		
	
		'UPGRADE_NOTE: Object objOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objOutput = Nothing
		
		OutputReport = fOK
		
	End Function
	
	
	Private Function GetReportingStructure(ByRef lngSingleRecord As Integer) As String
		
		Dim rsTemp As DataTable
		Dim strSQL As String
		Dim strSQL1 As String
		Dim strSQL2 As String
		Dim strResult As String
		Dim strLastResult As String
		
		Dim strESelect As String
		Dim strEJoin As String
		Dim strMSelect As String
		Dim strMJoin As String
		Dim strViewIDs As String
		
		If mblnReportingStructure Then
			
			strViewIDs = vbNullString
			GetSelectAndJoinForColumn(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelEmployeeNumberColumnName, strESelect, strEJoin, strViewIDs)
			GetSelectAndJoinForColumn(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelManagerStaffNoColumnName, strMSelect, strMJoin, strViewIDs)
			
			
			
			strSQL = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & " FROM " & gsPersonnelTableName & strEJoin & " WHERE " & gsPersonnelTableName & ".ID = " & CStr(lngSingleRecord)
			
			If mlngMatchReportType = MatchReportType.mrtSucession Then
				strSQL1 = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & " FROM " & gsPersonnelTableName & strEJoin & " WHERE " & strMSelect & " IN ("
				strSQL2 = ")"
			Else
				strSQL1 = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & " FROM " & gsPersonnelTableName & strEJoin & " WHERE " & strESelect & " IN (" & "SELECT " & strMSelect & " FROM " & gsPersonnelTableName & strEJoin & strMJoin & " WHERE " & strESelect & " IN ("
				strSQL2 = "))"
			End If
			
			
			strResult = "0"
			Do 
				rsTemp = DB.GetDataTable(strSQL)
				
				strLastResult = vbNullString
        for each objRow as DataRow in rsTemp.Rows

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(objRow(1)) Then
						If Trim(objRow(1).ToString()) <> vbNullString Then
								
							If mlngMatchReportType = MatchReportType.mrtSucession Then
								strResult = strResult & IIf(strResult <> vbNullString, ", ", "") & objRow(0).ToString()
							Else
								strResult = strResult & IIf(strResult <> vbNullString, ", ", "") & CStr(GetJobTableID((CInt(objRow(0)))))
							End If
								
							strLastResult = strLastResult & IIf(strLastResult <> vbNullString, ", ", "") & "'" & objRow(1).ToString() & "'"
								
						End If
					End If

				Next
				
				strSQL = strSQL1 & strLastResult & strSQL2
				
			Loop While strLastResult <> vbNullString
			
			
			If strResult <> vbNullString Then
				If mlngMatchReportType = MatchReportType.mrtSucession Then
					strResult = IIf(mlngTable1ID = glngPersonnelTableID, mstrTable1RealSource, mstrTable2RealSource) & ".ID IN (" & strResult & ")"
				Else
					strResult = IIf(mlngTable1ID = glngPostTableID, mstrTable1RealSource, mstrTable2RealSource) & ".ID IN (" & strResult & ")"
				End If
			End If
			
		End If
		
		strResult = strResult & IIf(strResult <> vbNullString, " AND ", vbNullString) & "ASRSys" & gsPersonnelTableName & gstrGradeTableName & "." & gstrNumLevelColumnName & " <" & IIf(mblnEqualGrade, "=", "") & " " & "ASRSys" & gstrPostTableName & gstrGradeTableName & "." & gstrNumLevelColumnName
		
		GetReportingStructure = strResult
		
	End Function
	
	
	Private Function GetJobTableID(ByRef lngRecordID As Integer) As Integer
		
		Dim mobjColumnPrivileges As CColumnPrivileges
		
		Dim rsTemp As DataTable
		Dim strSQL As String
		Dim strJoin As String
		
		'If Not gcoTablePrivileges.Item(gsPersonnelTableName).AllowSelect Then
		'  mstrErrorMessage = "Unable to run this report as you do not have access to the " & gsPersonnelTableName & " Table"
		'  Exit Function
		'End If
		
		'If Not gcoTablePrivileges.Item(gstrPostTableName).AllowSelect Then
		'  mstrErrorMessage = "Unable to run this report as you do not have access to the " & gstrPostTableName & " Table"
		'  Exit Function
		'End If
		
		
		strSQL = GetSQLForColumn(glngPostTableID, gstrPostTableName, gstrJobTitleColumnName, 1) & " = (" & GetSQLForColumn(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelJobTitleColumnName, 2) & " = " & CStr(lngRecordID) & ")"
		rsTemp = DB.GetDataTable(strSQL)
		
		If rsTemp.Rows.Count > 0 Then
			GetJobTableID = CInt(rsTemp.Rows(0)("ID"))
		Else
			GetJobTableID = 0
		End If	
		
	End Function
	
	
	Private Function GetSQLForColumn(ByRef lngTableID As Integer, ByRef strTable As String, ByRef strColumn As String, ByRef intMode As Short) As String
		
		Dim strSelect As String
		Dim strJoin As String
		
		GetSelectAndJoinForColumn(lngTableID, strTable, strColumn, strSelect, strJoin, vbNullString)
		
		If strSelect = vbNullString Then
			mstrErrorMessage = vbCrLf & vbCrLf & "You do not have permission to see the column '" & strColumn & "'" & vbCrLf & "either directly or through any views."
		Else
			If intMode = 1 Then
				GetSQLForColumn = "SELECT " & strTable & ".ID FROM " & strTable & strJoin & " WHERE " & strSelect
			Else
				GetSQLForColumn = "SELECT " & strSelect & " FROM " & strTable & strJoin & " WHERE " & strTable & ".ID"
			End If
		End If
		
	End Function
	
	
	Private Sub GetSelectAndJoinForColumn(ByRef lngTableID As Integer, ByRef strTable As String, ByRef strColumn As String, ByRef strSelect As String, ByRef strJoin As String, ByRef strViewIDs As String)
		
		Dim mobjColumnPrivileges As CColumnPrivileges
		Dim mobjTableView As TablePrivilege
		Dim pblnColumnOK As Boolean
		
		mobjColumnPrivileges = GetColumnPrivileges(strTable)
		
		pblnColumnOK = mobjColumnPrivileges.IsValid(strColumn)
		If pblnColumnOK Then
			pblnColumnOK = mobjColumnPrivileges.Item(strColumn).AllowSelect
		End If
		
		If pblnColumnOK Then
			strSelect = strTable & "." & strColumn
			strJoin = vbNullString
		Else
			strSelect = vbNullString
			strJoin = vbNullString
			For	Each mobjTableView In gcoTablePrivileges.Collection
				If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTableID) And (mobjTableView.AllowSelect) Then
					
					mobjColumnPrivileges = GetColumnPrivileges((mobjTableView.ViewName))
					If mobjColumnPrivileges.IsValid(strColumn) Then
						If mobjColumnPrivileges.Item(strColumn).AllowSelect Then
							
							strSelect = strSelect & " WHEN NOT " & mobjTableView.ViewName & "." & strColumn & " IS NULL THEN " & mobjTableView.ViewName & "." & strColumn
							
							If InStr(strViewIDs, CStr(mobjTableView.ViewID)) = 0 Then
								strJoin = strJoin & " LEFT OUTER JOIN " & mobjTableView.ViewName & " ON " & mobjTableView.TableName & ".ID = " & mobjTableView.ViewName & ".ID" & vbCrLf
								strViewIDs = strViewIDs & " " & CStr(mobjTableView.ViewID)
							End If
							
						End If
					End If
				End If
				
			Next mobjTableView
			
			If strSelect <> vbNullString Then
				strSelect = "CASE" & strSelect & " ELSE NULL END"
			End If
			
		End If
		
	End Sub
	
	Private Function GetTempTable() As String
		
		Dim objColumn As DisplayColumn
		Dim strTempTable As String
		Dim strError As String 
		Dim strSQL As String
		Dim lngSize As Integer
		
		For	Each objColumn In mcolColDetails
			strSQL = strSQL & IIf(strSQL <> vbNullString, ", ", vbNullString) & vbCrLf
			strSQL = strSQL & "[" & objColumn.Heading & "]"
			
			Select Case objColumn.DataType

				Case ColumnDataType.sqlVarChar, ColumnDataType.sqlLongVarChar 'sqlLongVarChar = Working Pattern
					lngSize = objColumn.Size
					strSQL = strSQL & "[varchar] (" & IIf(lngSize = VARCHAR_MAX_Size, "MAX", lngSize) & ")"
				Case ColumnDataType.sqlBoolean
					strSQL = strSQL & "[varchar] (1)"
				Case ColumnDataType.sqlDate
					strSQL = strSQL & "[datetime]"
				Case ColumnDataType.sqlNumeric, ColumnDataType.sqlInteger
					strSQL = strSQL & "[float]"
				Case Else
					strSQL = strSQL & "[int]"
			End Select
			
			strSQL = strSQL & " NULL"
		Next objColumn
		
		strTempTable = General.UniqueSQLObjectName("ASRSysTempMatchReport", 3)
		strSQL = "CREATE TABLE [" & _login.Username & "].[" & strTempTable & "]" & " (" & strSQL & ")"
		
		DB.ExecuteSql(strSQL)
		mstrErrorMessage = strError
		fOK = (mstrErrorMessage = vbNullString)
		
		GetTempTable = strTempTable
		
	End Function
	
	Private Sub RemoveTemporarySQLObjects()
		General.DropUniqueSQLObject(mstrTempTableName, 3)		
	End Sub
	
	Private Function TablePermission(lngTableID As Integer) As Boolean
		
		Dim objTableView As TablePrivilege
		Dim blnFound As Boolean
		
		blnFound = False
		For	Each objTableView In gcoTablePrivileges.Collection
			If (objTableView.TableID = lngTableID) And (objTableView.AllowSelect) Then
				blnFound = True
				Exit For
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing
		
		TablePermission = blnFound
		
	End Function
	
	
	Private Function HasColumnPermission(lngTableID As Integer, strTable As String, strColumn As String) As Boolean
		
		Dim mobjColumnPrivileges As CColumnPrivileges
		Dim mobjTableView As TablePrivilege
		Dim pblnColumnOK As Boolean
		
		mobjColumnPrivileges = GetColumnPrivileges(strTable)
		
		pblnColumnOK = mobjColumnPrivileges.IsValid(strColumn)
		If pblnColumnOK Then
			pblnColumnOK = mobjColumnPrivileges.Item(strColumn).AllowSelect
		End If
			
		If Not pblnColumnOK Then
			
			For	Each mobjTableView In gcoTablePrivileges.Collection
				If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTableID) And (mobjTableView.AllowSelect) Then
					
					mobjColumnPrivileges = GetColumnPrivileges((mobjTableView.ViewName))
					If mobjColumnPrivileges.IsValid(strColumn) Then
						If mobjColumnPrivileges.Item(strColumn).AllowSelect Then
							pblnColumnOK = True
							Exit For
						End If
					End If
					
				End If
			Next mobjTableView
			
		End If
		
		Return pblnColumnOK
		
	End Function
	
	
	Private Function CheckModuleSetupPermissions() As Boolean
		
		If mlngMatchReportType <> MatchReportType.mrtNormal Then
			
			CheckModuleSetupPermissions = False
			
			If Not HasColumnPermission(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelGradeColumnName) Then
				mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelGradeColumnName & "' either directly or through any views."
				Exit Function
			End If
			If Not HasColumnPermission(glngPostTableID, gstrPostTableName, gstrPostGradeColumnName) Then
				mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrPostGradeColumnName & "' either directly or through any views."
				Exit Function
			End If
			If Not HasColumnPermission(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelJobTitleColumnName) Then
				mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelJobTitleColumnName & "' either directly or through any views."
				Exit Function
			End If
			If Not HasColumnPermission(glngPostTableID, gstrPostTableName, gstrJobTitleColumnName) Then
				mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrJobTitleColumnName & "' either directly or through any views."
				Exit Function
			End If
			
		End If
		
		CheckModuleSetupPermissions = True
		
	End Function
	Public Sub SetOutputParameters(lngOutputFormat As Integer, blnOutputScreen As Boolean, blnOutputPrinter As Boolean, strOutputPrinterName As String _
                                 , blnOutputSave As Boolean, lngOutputSaveExisting As Integer, blnOutputEmail As Boolean, lngOutputEmailAddr As Integer _
                                 , strOutputEmailSubject As String, strOutputEmailAttachAs As String, strOutputFilename As String _
                                 , blnPreviewOnScreen As Boolean, blnChkPicklistFilter As Boolean, Optional strOutputTitlePage As String = "" _
                                 , Optional strOutputReportPackTitle As String = "", Optional strOutputOverrideFilter As String = "" _
                                 , Optional blnOutputTOC As Boolean = False, Optional blnOutputCoverSheet As Boolean = False _
                                 , Optional lngOverrideFilterID As Integer = 0, Optional blnOutputRetainPivotOrChart As Boolean = False)
		
		mlngOutputFormat = lngOutputFormat
		mblnOutputScreen = blnOutputScreen
		mblnOutputPrinter = blnOutputPrinter
		mstrOutputPrinterName = strOutputPrinterName
		mblnOutputSave = blnOutputSave
		mlngOutputSaveExisting = lngOutputSaveExisting
		mblnOutputEmail = blnOutputEmail
		mlngOutputEmailAddr = lngOutputEmailAddr
		mstrOutputEmailSubject = strOutputEmailSubject
		mstrOutputEmailAttachAs = strOutputEmailAttachAs
		mstrOutputFileName = strOutputFilename
		mblnChkPicklistFilter = blnChkPicklistFilter
		mblnPreviewOnScreen = (blnPreviewOnScreen Or (mlngOutputFormat = OutputFormats.DataOnly And mblnOutputScreen))
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mstrOutputTitlePage = IIf(IsNothing(strOutputTitlePage), ExpressionValueTypes.giEXPRVALUE_CHARACTER, strOutputTitlePage)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mstrOutputReportPackTitle = IIf(IsNothing(strOutputReportPackTitle), ExpressionValueTypes.giEXPRVALUE_CHARACTER, strOutputReportPackTitle)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mstrOutputOverrideFilter = IIf(IsNothing(strOutputOverrideFilter), ExpressionValueTypes.giEXPRVALUE_CHARACTER, strOutputOverrideFilter)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mblnOutputTOC = IIf(IsNothing(blnOutputTOC), ExpressionValueTypes.giEXPRVALUE_CHARACTER, blnOutputTOC)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mlngOverrideFilterID = IIf(IsNothing(lngOverrideFilterID), ExpressionValueTypes.giEXPRVALUE_CHARACTER, lngOverrideFilterID)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		mblnOutputRetainPivotOrChart = IIf(IsNothing(blnOutputRetainPivotOrChart), ExpressionValueTypes.giEXPRVALUE_CHARACTER, blnOutputRetainPivotOrChart)
		'mblnOutputRetainCharts = IIf(IsMissing(blnOutputRetainCharts), giEXPRVALUE_CHARACTER, blnOutputRetainCharts)
	End Sub


  ' Move to base class because its sort of shared?

      Public Function SetPromptedValues(pavPromptedValues As Object) As Boolean

        ' Purpose : This function calls the individual functions that
        '           generate the components of the main SQL string.
        Dim iLoop As Short
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
            Logs.AddDetailEntry( "Error whilst setting prompted values. " & ex.Message.RemoveSensitive())
            Return False

        End Try

        Return True

    End Function

  	Private Function AddItemToReportData(data As IEnumerable) As Boolean

		Dim dr As DataRow
		Dim iColumn As Integer = 0

		dr = ReportDataTable.Rows.Add()

		If Not data Is Nothing Then
			For Each objData In data
				dr(iColumn) = objData
				iColumn += 1
			Next
		End If

		Return True

	End Function

End Class
