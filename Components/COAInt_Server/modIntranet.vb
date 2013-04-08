Option Strict Off
Option Explicit On
Module modIntranet
	
	Public Sub CreateASRDev_SysProtects(ByRef pConn As ADODB.Connection)
		
		Dim cmdCreateCache As ADODB.Command
		
		' Create the cached system tables on the server - Don;t do it in a stored procedure because the #temp will then only be visible to that stored procedure
		cmdCreateCache = New ADODB.Command
		cmdCreateCache.CommandText = "DECLARE @iUserGroupID  integer, " & vbNewLine & " @sUserGroupName   sysname, " & vbNewLine & " @sActualLoginName varchar(250) " & vbNewLine & "-- Get the current user's group ID. " & vbNewLine & "EXEC spASRIntGetActualUserDetails " & vbNewLine & " @sActualLoginName OUTPUT, " & vbNewLine & " @sUserGroupName OUTPUT, " & vbNewLine & " @iUserGroupID OUTPUT " & vbNewLine & "-- Create the SysProtects cache table " & vbNewLine & "IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL" & vbNewLine & " DROP TABLE #SysProtects " & vbNewLine & "CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) " & vbNewLine & " INSERT #SysProtects " & vbNewLine & " SELECT ID, Action, Columns, ProtectType " & vbNewLine & "       FROM sysprotects " & vbNewLine & "       WHERE uid = @iUserGroupID" & vbNewLine & "CREATE INDEX [IDX_ID] ON #SysProtects(ID)"
		cmdCreateCache.ActiveConnection = pConn
		cmdCreateCache.Execute()
		'UPGRADE_NOTE: Object cmdCreateCache may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCreateCache = Nothing
		
	End Sub
	
	Public Function UDFFunctions(ByRef pastrUDFFunctions() As String, ByRef pbCreate As Boolean) As Boolean
		
		On Error GoTo UDFFunctions_ERROR
		
		Dim iCount As Short
		Dim strDropCode As String
		Dim strFunctionName As String
		Dim sUDFCode As String
		Dim datData As clsDataAccess
		Dim iStart As Short
		Dim iEnd As Short
		Dim strFunctionNumber As String
		
		Const FUNCTIONPREFIX As String = "udf_ASRSys_"
		
		datData = New clsDataAccess
		
		If gbEnableUDFFunctions Then
			
			For iCount = 1 To UBound(pastrUDFFunctions)
				
				'JPD 20060110 Fault 10509
				'iStart = Len("CREATE FUNCTION udf_ASRSys_") + 1
				iStart = InStr(pastrUDFFunctions(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
				iEnd = InStr(1, Mid(pastrUDFFunctions(iCount), 1, 1000), "(@Pers")
				strFunctionNumber = Mid(pastrUDFFunctions(iCount), iStart, iEnd - iStart)
				strFunctionName = FUNCTIONPREFIX & strFunctionNumber
				
				'Drop existing function (could exist if the expression is used more than once in a report)
				strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
				datData.ExecuteSql(strDropCode)
				
				
				' Create the new function
				If pbCreate Then
					sUDFCode = pastrUDFFunctions(iCount)
					datData.ExecuteSql(sUDFCode)
				End If
				
			Next iCount
		End If
		
		UDFFunctions = True
		Exit Function
		
UDFFunctions_ERROR: 
		UDFFunctions = False
		
	End Function
	
	'UPGRADE_WARNING: Sub Main in a DLL won't get called. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A90BF69E-29C2-4F6F-9E44-92CFC7FAA399"'
	Public Sub Main()
		
		'Instantiate User Interface class
		UI = New clsUI
		
		' Are we in debug mode
		ASRDEVELOPMENT = Not vbCompiled
		
		' Enable UDF functions (this needs to be altered to be SQL Server specific)
		gbEnableUDFFunctions = True
		
	End Sub
	
	Public Function vbCompiled() As Boolean
		
		'Much better (and clever-er) !
		On Error Resume Next
		Err.Clear()
		Debug.Print(1 / 0)
		vbCompiled = (Err.Number = 0)
		
	End Function
	
	
  Public Function GetEmailGroupName(ByVal lngGroupID As Integer) As String

    Dim datData As clsDataAccess
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo LocalErr

    GetEmailGroupName = vbNullString

    datData = New clsDataAccess

    strSQL = "SELECT Name FROM ASRSysEmailGroupName " & "WHERE EmailGroupID = " & CStr(lngGroupID)
    rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    With rsTemp
      If Not .BOF And Not .EOF Then
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If Not IsDBNull(rsTemp.Fields("Name").Value) Then
          GetEmailGroupName = rsTemp.Fields("Name").Value
        End If
      End If
    End With

    rsTemp.Close()

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    datData = Nothing

    Exit Function

LocalErr:

  End Function
	
	' Encode an string so that it can be displayed correctly
	' inside the browser.
	'
	' Same effect as the Server.HTMLEncode method in ASP
	Public Function HTMLEncode(ByVal Text As String) As String
		Dim i As Short
		Dim acode As Short
		Dim repl As String
		
		HTMLEncode = Text
		
		For i = Len(HTMLEncode) To 1 Step -1
			acode = Asc(Mid(HTMLEncode, i, 1))
			Select Case acode
				Case 32
					repl = "&nbsp;"
				Case 34
					repl = "&quot;"
				Case 38
					repl = "&amp;"
				Case 60
					repl = "&lt;"
				Case 62
					repl = "&gt;"
				Case 32 To 127
					' don't touch alphanumeric chars
				Case Else
					repl = "&#" & CStr(acode) & ";"
			End Select
			If Len(repl) Then
				HTMLEncode = Left(HTMLEncode, i - 1) & repl & Mid(HTMLEncode, i + 1)
				repl = ""
			End If
		Next 
	End Function
End Module