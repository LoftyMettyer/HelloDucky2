Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports HR.Intranet.Server.BaseClasses

Public Class frmMatchRunBreakDown
  Inherits BaseReport

  public Property ParentForm_Renamed as Object
  Public Property  lblTable1Name() As new TextBox
  Public Property lblTable2Name() as New TextBox
  Public Property  Table1RecDescExprID() As Integer
  Public Property Loading As Boolean
  Public Property  cboRelation() As new ListBox
  public chkAllRecords As New CheckBox

  Private mcolRecDesc1 As Collection
	Private mcolRecDesc2 As Collection
  Private mlngCrossRefArray(,) As Integer

  Private mlngTable1RecDescExprID As Integer
	Private mlngTable2RecDescExprID As Integer

  public sub ShowBreakdown(lngRecord1ID As long, lngRecord2ID As long, mlngMatchReportType As long)
  End sub

	Public Sub AddToCrossRef(ByRef lngID1 As Integer, ByRef lngID2 As Integer)
		
		'Dim strRecDesc As String
		'Dim lngIndex As Integer
		
  '  Try
	
		'  lngIndex = UBound(mlngCrossRefArray, 2) + 1
		'  ReDim Preserve mlngCrossRefArray(1, lngIndex)
		'  mlngCrossRefArray(0, lngIndex) = lngID1
		'  mlngCrossRefArray(1, lngIndex) = lngID2
			
		'  strRecDesc = GetRecordDesc(mlngTable1RecDescExprID, lngID1)
		'  mcolRecDesc1.Add(strRecDesc, "ID" & CStr(lngID1))
		'  'If lngIndex = 1 Or mlngTable2RecDescExprID = 0 Then
		'	 ' cboTable1.Items.Add(New VB6.ListBoxItem(strRecDesc, lngID1))
		'  'End If
		
		'  If lngID2 > 0 Then
		'	  strRecDesc = GetRecordDesc(mlngTable2RecDescExprID, lngID2)
		'	  mcolRecDesc2.Add(strRecDesc, "ID" & CStr(lngID2))
		'  End If

		'Catch ex As Exception


  '  End Try	
		

	End Sub

  	Private Function GetRecordDesc(lngRecDescExprID As Integer, lngRecordID As Integer) As String

        Dim prmRecordDesc = New SqlParameter("psRecDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

        Dim unKnownVar1 As Integer = 0
        Dim unKnownVar2 As Integer = 0
        Dim unKnownVar3 As Integer = 0

        DB.ExecuteSP("sp_ASRIntGetRecordDescription" _
                            , New SqlParameter("piTableID", SqlDbType.Int) With {.Value = unKnownVar1} _
                            , New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = lngRecordID} _
                            , New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = unKnownVar2} _
                            , New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = unKnownVar3} _
                            , prmRecordDesc)

        Return prmRecordDesc.Value.ToString

'    ' 
'    Return TRUE if the user has been granted the given permission.
'		Dim cmADO As ADODB.Command
'		Dim pmADO As ADODB.Parameter
		
'		On Error GoTo LocalErr
		
'		If lngRecDescExprID < 1 Then
'			'GetRecordDesc = "Record Description Undefined"
'			GetRecordDesc = vbNullString
'			Exit Function
'		End If
		
		
'		' Check if the user can create New instances of the given category.
'		cmADO = New ADODB.Command
'		With cmADO
'			.CommandText = "dbo.sp_ASRExpr_" & lngRecDescExprID
'			.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
'			.CommandTimeout = 0
'			.ActiveConnection = gADOCon
			
'			pmADO = .CreateParameter("Result", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, VARCHAR_MAX_Size)
'			.Parameters.Append(pmADO)
			
'			pmADO = .CreateParameter("RecordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
'			.Parameters.Append(pmADO)
'			pmADO.Value = lngRecordID
			
'			cmADO.Execute()
			
'			GetRecordDesc = .Parameters(0).Value
			
'		End With
'		'UPGRADE_NOTE: Object cmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
'		cmADO = Nothing
		
'		Exit Function
		
'LocalErr: 
'		'COAMsgBox "Error reading record description" & vbCr & _
'		'"(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(lngRecDescExprID)
'		'fOK = False
		
	End Function

End Class
