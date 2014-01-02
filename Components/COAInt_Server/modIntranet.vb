Option Strict Off
Option Explicit On
Module modIntranet

Private mobjGeneral As New clsGeneral

	Public Function UDFFunctions(ByRef pastrUDFFunctions() As String, ByRef pbCreate As Boolean) As Boolean
		Return mobjGeneral.UDFFunctions(pastrUDFFunctions, pbCreate)
	End Function
	
	Public Sub Initialise()




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
    Dim i As Integer
    Dim acode As Integer
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