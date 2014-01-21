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

		Dim rsTemp As DataTable
		Dim strSQL As String

		Try

			strSQL = "SELECT Name FROM ASRSysEmailGroupName " & "WHERE EmailGroupID = " & CStr(lngGroupID)
			rsTemp = dataAccess.GetDataTable(strSQL, CommandType.Text)

			For Each objRow In rsTemp.Rows
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDBNull(objRow("Name")) Then
					Return objRow("Name")
				End If
			Next

		Catch ex As Exception
			Throw

		End Try

		Return ""

	End Function


	Friend Function DecToBin(ByRef DeciValue As Integer, Optional ByRef NoOfBits As Short = 8) As String

		Dim i As Short 'make sure there are enough bits to contain the number
		Do While DeciValue > (2 ^ NoOfBits) - 1
			NoOfBits = NoOfBits + 8
		Loop
		DecToBin = vbNullString
		'build the string
		For i = 0 To (NoOfBits - 1)
			DecToBin = CStr(CShort(DeciValue And 2 ^ i) / 2 ^ i) & DecToBin
		Next i
	End Function


End Module