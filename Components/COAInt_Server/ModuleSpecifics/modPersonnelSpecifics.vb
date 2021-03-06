Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Namespace ModuleSpecifics

	Friend Class modPersonnelSpecifics
		Inherits BaseModuleSpecific

		Public Sub New(value As SessionInfo)
			MyBase.New(value)
		End Sub

		' Module parameters.
		Public gfPersonnelEnabled As Boolean
		Public grtRegionType As RegionType
		Public gwptWorkingPatternType As WorkingPatternType

		Public glngPersonnelTableID As Integer
		Public gsPersonnelTableName As String
		Private mvar_lngPersonnelEmployeeNumberID As Integer
		Public gsPersonnelEmployeeNumberColumnName As String
		Private mvar_lngPersonnelSurnameID As Integer
		Public gsPersonnelSurnameColumnName As String
		Private mvar_lngPersonnelForenameID As Integer
		Public gsPersonnelForenameColumnName As String

		'Private glngPersonnelStartDateID As Long
		Public glngPersonnelStartDateID As Integer

		Public gsPersonnelStartDateColumnName As String
		Private mvar_lngPersonnelLeavingDateID As Integer
		Public gsPersonnelLeavingDateColumnName As String
		Private mvar_lngPersonnelFullPartTimeID As Integer
		Public gsPersonnelFullPartTimeColumnName As String
		Private mvar_lngPersonnelEmailID As Integer
		Public gsPersonnelEmailColumnName As String
		Private mvar_lngPersonnelDepartmentID As Integer
		Public gsPersonnelDepartmentColumnName As String

		' Static Region
		Private mvar_lngPersonnelRegionID As Integer
		Public gsPersonnelRegionColumnName As String
		' Historic Region
		Private mvar_lngPersonnelHRegionTableID As Integer
		Public glngPersonnelHRegionTableID As Integer
		Public gsPersonnelHRegionTableName As String
		Private mvar_lngPersonnelHRegionFieldID As Integer
		Public gsPersonnelHRegionColumnName As String
		Private mvar_lngPersonnelHRegionDateID As Integer
		Public gsPersonnelHRegionDateColumnName As String
		Public gsPersonnelHRegionTableRealSource As String

		' Static Working Pattern
		Private mvar_lngPersonnelWorkingPatternID As Integer
		Public gsPersonnelWorkingPatternColumnName As String
		' Historic Working Pattern
		Private mvar_lngPersonnelHWorkingPatternTableID As Integer
		Public gsPersonnelHWorkingPatternTableName As String
		Private mvar_lngPersonnelHWorkingPatternFieldID As Integer
		Public gsPersonnelHWorkingPatternColumnName As String
		Private mvar_lngPersonnelHWorkingPatternDateID As Integer
		Public gsPersonnelHWorkingPatternDateColumnName As String
		Public gsPersonnelHWorkingPatternTableRealSource As String

		Public glngHierarchyTableID As Integer
		Public gsHierarchyTableName As String

		Public Sub ReadPersonnelParameters()

			Dim objTable As TablePrivilege

			' Read the Personnel module parameters from the database.
			glngPersonnelTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
			If glngPersonnelTableID > 0 Then
				gsPersonnelTableName = _tables.GetById(glngPersonnelTableID).Name
			Else
				gsPersonnelTableName = ""
			End If

			mvar_lngPersonnelEmployeeNumberID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER))
			If mvar_lngPersonnelEmployeeNumberID > 0 Then
				gsPersonnelEmployeeNumberColumnName = _columns.GetById(mvar_lngPersonnelEmployeeNumberID).Name
			Else
				gsPersonnelEmployeeNumberColumnName = ""
			End If

			mvar_lngPersonnelSurnameID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
			If mvar_lngPersonnelSurnameID > 0 Then
				gsPersonnelSurnameColumnName = _columns.GetById(mvar_lngPersonnelSurnameID).Name
			Else
				gsPersonnelSurnameColumnName = ""
			End If

			mvar_lngPersonnelForenameID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
			If mvar_lngPersonnelForenameID > 0 Then
				gsPersonnelForenameColumnName = _columns.GetById(mvar_lngPersonnelForenameID).Name
			Else
				gsPersonnelForenameColumnName = ""
			End If

			glngPersonnelStartDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE))
			If glngPersonnelStartDateID > 0 Then
				gsPersonnelStartDateColumnName = _columns.GetById(glngPersonnelStartDateID).Name
			Else
				gsPersonnelStartDateColumnName = ""
			End If

			mvar_lngPersonnelLeavingDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE))
			If mvar_lngPersonnelLeavingDateID > 0 Then
				gsPersonnelLeavingDateColumnName = _columns.GetById(mvar_lngPersonnelLeavingDateID).Name
			Else
				gsPersonnelLeavingDateColumnName = ""
			End If

			mvar_lngPersonnelFullPartTimeID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FULLPARTTIME))
			If mvar_lngPersonnelFullPartTimeID > 0 Then
				gsPersonnelFullPartTimeColumnName = _columns.GetById(mvar_lngPersonnelFullPartTimeID).Name
			Else
				gsPersonnelFullPartTimeColumnName = ""
			End If

			mvar_lngPersonnelEmailID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMAIL))
			If mvar_lngPersonnelEmailID > 0 Then
				gsPersonnelEmailColumnName = _columns.GetById(mvar_lngPersonnelEmailID).Name
			Else
				gsPersonnelEmailColumnName = ""
			End If

			mvar_lngPersonnelDepartmentID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT))
			If mvar_lngPersonnelDepartmentID > 0 Then
				gsPersonnelDepartmentColumnName = _columns.GetById(mvar_lngPersonnelDepartmentID).Name
			Else
				gsPersonnelDepartmentColumnName = ""
			End If

			' Static Region
			mvar_lngPersonnelRegionID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION))
			If mvar_lngPersonnelRegionID > 0 Then
				gsPersonnelRegionColumnName = _columns.GetById(mvar_lngPersonnelRegionID).Name
				grtRegionType = RegionType.rtStaticRegion
			Else
				gsPersonnelRegionColumnName = ""
				grtRegionType = RegionType.rtNotDefined
			End If

			' Historic Region
			mvar_lngPersonnelHRegionTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE))
			glngPersonnelHRegionTableID = mvar_lngPersonnelHRegionTableID
			If mvar_lngPersonnelHRegionTableID > 0 Then
				gsPersonnelHRegionTableName = _tables.GetById(mvar_lngPersonnelHRegionTableID).Name

				' Get the realsource into a variable too
				objTable = _tablePrivileges.FindTableID(mvar_lngPersonnelHRegionTableID)
				gsPersonnelHRegionTableRealSource = objTable.RealSource
				grtRegionType = RegionType.rtHistoricRegion
			Else
				gsPersonnelHRegionTableName = ""
				If grtRegionType <> RegionType.rtStaticRegion Then grtRegionType = RegionType.rtNotDefined
			End If

			mvar_lngPersonnelHRegionFieldID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD))
			If mvar_lngPersonnelHRegionFieldID > 0 Then
				gsPersonnelHRegionColumnName = _columns.GetById(mvar_lngPersonnelHRegionFieldID).Name
			Else
				gsPersonnelHRegionColumnName = ""
				If grtRegionType <> RegionType.rtStaticRegion Then grtRegionType = RegionType.rtNotDefined
			End If

			mvar_lngPersonnelHRegionDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE))
			If mvar_lngPersonnelHRegionDateID > 0 Then
				gsPersonnelHRegionDateColumnName = _columns.GetById(mvar_lngPersonnelHRegionDateID).Name
			Else
				gsPersonnelHRegionDateColumnName = ""
				If grtRegionType <> RegionType.rtStaticRegion Then grtRegionType = RegionType.rtNotDefined
			End If

			' Static Working Pattern
			mvar_lngPersonnelWorkingPatternID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN))
			If mvar_lngPersonnelWorkingPatternID > 0 Then
				gsPersonnelWorkingPatternColumnName = _columns.GetById(mvar_lngPersonnelWorkingPatternID).Name
				gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern
			Else
				gsPersonnelWorkingPatternColumnName = ""
				gwptWorkingPatternType = WorkingPatternType.wptnotDefined
			End If

			' Historic Working Pattern
			mvar_lngPersonnelHWorkingPatternTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE))
			If mvar_lngPersonnelHWorkingPatternTableID > 0 Then
				gsPersonnelHWorkingPatternTableName = _tables.GetById(mvar_lngPersonnelHWorkingPatternTableID).Name

				' Get the realsource into a variable too
				objTable = _tablePrivileges.FindTableID(mvar_lngPersonnelHWorkingPatternTableID)
				gsPersonnelHWorkingPatternTableRealSource = objTable.RealSource
				gwptWorkingPatternType = WorkingPatternType.wptHistoricWPattern
			Else
				gsPersonnelHWorkingPatternTableName = ""
				If gwptWorkingPatternType <> WorkingPatternType.wptStaticWPattern Then gwptWorkingPatternType = WorkingPatternType.wptnotDefined
			End If

			mvar_lngPersonnelHWorkingPatternFieldID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD))
			If mvar_lngPersonnelHWorkingPatternFieldID > 0 Then
				gsPersonnelHWorkingPatternColumnName = _columns.GetById(mvar_lngPersonnelHWorkingPatternFieldID).Name
			Else
				gsPersonnelHWorkingPatternColumnName = ""
				If gwptWorkingPatternType <> WorkingPatternType.wptStaticWPattern Then gwptWorkingPatternType = WorkingPatternType.wptnotDefined
			End If

			mvar_lngPersonnelHWorkingPatternDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE))
			If mvar_lngPersonnelHWorkingPatternDateID > 0 Then
				gsPersonnelHWorkingPatternDateColumnName = _columns.GetById(mvar_lngPersonnelHWorkingPatternDateID).Name
			Else
				gsPersonnelHWorkingPatternDateColumnName = ""
				If gwptWorkingPatternType <> WorkingPatternType.wptStaticWPattern Then gwptWorkingPatternType = WorkingPatternType.wptnotDefined
			End If

			' Read the Personnel module parameters from the database.
			glngHierarchyTableID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE))
			If glngHierarchyTableID > 0 Then
				gsHierarchyTableName = _tables.GetById(glngHierarchyTableID).Name
			Else
				gsHierarchyTableName = ""
			End If

			'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTable = Nothing

		End Sub

		Public Function ValidatePersonnelParameters() As Boolean

			' Validate the configuration of the Personnel module parameters

			Dim fValid As Boolean

			' Check that the Personnel module is installed.
			fValid = gfPersonnelEnabled

			' -----------------------------------------------
			If fValid Then
				' Check the Personnel Table ID is valid.
				fValid = (glngPersonnelTableID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Personnel table is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Employee Number ID is valid.
				fValid = (mvar_lngPersonnelEmployeeNumberID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Employee Number column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Surname ID is valid.
				fValid = (mvar_lngPersonnelSurnameID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Surname column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Forename ID is valid.
				fValid = (mvar_lngPersonnelForenameID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Forename column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the StartDate ID is valid.
				fValid = (glngPersonnelStartDateID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Start Date column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Leaving Date ID is valid.
				fValid = (mvar_lngPersonnelLeavingDateID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Leaving Date column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the FullPartTime ID is valid.
				fValid = (mvar_lngPersonnelFullPartTimeID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Full/Part Time column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Email ID is valid.
				fValid = (mvar_lngPersonnelEmailID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Email column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Department Date ID is valid.
				fValid = (mvar_lngPersonnelDepartmentID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Department column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Working Pattern Date ID is valid.
				fValid = (mvar_lngPersonnelWorkingPatternID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Working Pattern column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			If fValid Then
				' Check the Region ID is valid.
				fValid = (mvar_lngPersonnelRegionID > 0)
				If Not fValid Then
					'NO MSGBOX ON THE SERVER ! - MsgBox "The Personnel module is not properly configured." & vbNewLine & _
					'"The Region column is not defined.", vbOKOnly, App.ProductName
				End If
			End If

			'
			'  If fValid Then
			'    ' Get the column privileges for the Course table.
			'    Set objCourseColumnPrivileges = GetColumnPrivileges(gsCourseTableName)
			'
			'    ' Check that the user has permission to see the Course Title column.
			'    fValid = objCourseColumnPrivileges.Item(gsCourseTitleColumnName).AllowSelect
			'    If Not fValid Then
			'      MsgBox "You do not have permission to see the defined Course Title column.", vbOKOnly, App.ProductName
			'    End If
			'  End If
			'
			'
			'  If fValid And (Len(gsCourseCancelledByColumnName) > 0) Then
			'    ' Check the Course Cancelled By column can be updated if it is defined.
			'    fValid = objCourseColumnPrivileges.Item(gsCourseCancelledByColumnName).AllowUpdate
			'    If Not fValid Then
			'      MsgBox "You do not have permission to update the defined Course Cancelled By column.", vbOKOnly, App.ProductName
			'    End If
			'  End If
			'
			'
			'
			'  ' Return the validation value.
			ValidatePersonnelParameters = fValid

		End Function
	End Class
End Namespace