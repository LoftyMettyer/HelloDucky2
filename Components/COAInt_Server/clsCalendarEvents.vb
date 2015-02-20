Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses

Friend Class clsCalendarEvents
	Inherits BaseForDMI
	Implements IEnumerable

	' Hold local collection
	Private mCol As Collection

	Public Function Add(ByVal pstrKey As String, ByVal pstrName As String, Optional ByVal plngTableID As Integer = 0, Optional ByVal pstrTableName As String = "" _
											, Optional ByVal plngFilterID As Integer = 0, Optional ByVal plngStartDateID As Integer = 0, Optional ByVal pstrStartDateName As String = "" _
											, Optional ByVal plngStartSessionID As Integer = 0, Optional ByVal pstrStartSessionName As String = "", Optional ByVal plngEndDateID As Integer = 0 _
											, Optional ByVal pstrEndDateName As String = "", Optional ByVal plngEndSessionID As Integer = 0, Optional ByVal pstrEndSessionName As String = "" _
											, Optional ByVal plngDurationID As Integer = 0, Optional ByVal pstrDurationName As String = "", Optional ByVal pintLegendType As Short = 0 _
											, Optional ByVal pstrLegendCharacter As String = "", Optional ByVal plngLegendTableID As Integer = 0, Optional ByVal pstrLegendTableName As String = "" _
											, Optional ByVal plngLegendColumnID As Integer = 0, Optional ByVal pstrLegendColumnName As String = "", Optional ByVal plngLegendCodeID As Integer = 0 _
											, Optional ByVal pstrLegendCodeName As String = "", Optional ByVal plngLegendEventTypeID As Integer = 0, Optional ByVal pstrLegendEventTypeName As String = "" _
											, Optional ByVal plngDesc1ID As Integer = 0, Optional ByRef pstrDesc1Name As String = "", Optional ByVal plngDesc2ID As Integer = 0 _
											, Optional ByVal pstrDesc2Name As String = "", Optional ByVal pstrBaseDescription As String = "", Optional ByVal pstrRegion As String = "" _
											, Optional ByVal pstrWorkingPattern As String = "", Optional ByVal pstrDesc1Value As String = "", Optional ByVal pstrDesc2Value As String = "") As clsCalendarEvent

		' Add a new object to the collection
		Dim objNewMember As clsCalendarEvent
		objNewMember = New clsCalendarEvent

		With objNewMember
			.Key = pstrKey
			.Name = pstrName

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngTableID) Then .TableID = plngTableID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrTableName) Then .TableName = pstrTableName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngFilterID) Then .FilterID = plngFilterID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngStartDateID) Then .StartDateID = plngStartDateID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrStartDateName) Then .StartDateName = pstrStartDateName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngStartSessionID) Then .StartSessionID = plngStartSessionID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrStartSessionName) Then .StartSessionName = pstrStartSessionName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngEndDateID) Then .EndDateID = plngEndDateID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrEndDateName) Then .EndDateName = pstrEndDateName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngEndSessionID) Then .EndSessionID = plngEndSessionID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrEndSessionName) Then .EndSessionName = pstrEndSessionName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngDurationID) Then .DurationID = plngDurationID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrDurationName) Then .DurationName = pstrDurationName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pintLegendType) Then .LegendType = pintLegendType
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrLegendCharacter) Then .LegendCharacter = pstrLegendCharacter
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngLegendTableID) Then .LegendTableID = plngLegendTableID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrLegendTableName) Then .LegendTableName = pstrLegendTableName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngLegendColumnID) Then .LegendColumnID = plngLegendColumnID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrLegendColumnName) Then .LegendColumnName = pstrLegendColumnName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngLegendCodeID) Then .LegendCodeID = plngLegendCodeID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrLegendCodeName) Then .LegendCodeName = pstrLegendCodeName
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(plngLegendEventTypeID) Then .LegendEventTypeID = plngLegendEventTypeID
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrLegendEventTypeName) Then .LegendEventTypeName = pstrLegendEventTypeName

			If plngDesc1ID > 0 Then
				.Description1ID = plngDesc1ID
				.Description1_TableID = GetColumnTable(plngDesc1ID)
				.Description1_TableName = GetColumnTableName(plngDesc1ID)
				.Description1_ColumnName = GetColumnName(plngDesc1ID)
			End If
			If Not IsNothing(pstrDesc1Name) Then .Description1Name = pstrDesc1Name

			If plngDesc2ID > 0 Then
				.Description2ID = plngDesc2ID
				.Description2_TableID = GetColumnTable(plngDesc2ID)
				.Description2_TableName = GetColumnTableName(plngDesc2ID)
				.Description2_ColumnName = GetColumnName(plngDesc2ID)
			End If

			If Not IsNothing(pstrDesc2Name) Then .Description2Name = pstrDesc2Name

			'*************************************
			'optional event data used for calendar report breakdown
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrBaseDescription) Then .BaseDescription = pstrBaseDescription
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrRegion) Then .Region = pstrRegion
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrWorkingPattern) Then .WorkingPattern = pstrWorkingPattern
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrDesc1Value) Then .Desc1Value = pstrDesc1Value
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(pstrDesc2Value) Then .Desc2Value = pstrDesc2Value
			'*************************************

		End With

		mCol.Add(objNewMember, pstrKey)

		Add = objNewMember

		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing

	End Function

	Public ReadOnly Property Item(ByVal pstrEventName As String) As clsCalendarEvent
		Get
			' Provide a reference to a specific item in the collection

			Item = mCol.Item(pstrEventName)
		End Get
	End Property

	Public Property Collection() As Collection
		Get
			Collection = mCol
		End Get
		Set(ByVal Value As Collection)
			mCol = Value
		End Set
	End Property

	Public ReadOnly Property Count() As Integer
		Get
			' Provide number of objects in the collection
			Count = mCol.Count()
		End Get
	End Property

	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mCol.GetEnumerator
	End Function

	Public Sub Remove(ByRef pstrName As String)
		' Remove a specific object from the collection
		mCol.Remove(pstrName)
	End Sub

	Public Sub RemoveAll()
		' Remove all object from the collection
		Do While mCol.Count() > 1
			mCol.Remove(mCol.Count())
		Loop
	End Sub

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mCol = New Collection
	End Sub

	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
	End Sub

	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

End Class