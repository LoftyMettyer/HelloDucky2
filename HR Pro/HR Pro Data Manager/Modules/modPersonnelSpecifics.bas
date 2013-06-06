Attribute VB_Name = "modPersonnelSpecifics"
Option Explicit

' Module parameters.
Public gfPersonnelEnabled As Boolean
Public grtRegionType As RegionType
Public gwptWorkingPatternType As WorkingPatternType

Public Enum RegionType
  rtNotDefined = 0
  rtStaticRegion = 1
  rtHistoricRegion = 2
End Enum

Public Enum WorkingPatternType
  wptnotDefined = 0
  wptStaticWPattern = 1
  wptHistoricWPattern = 2
End Enum


' Module constants.
Public Const gsMODULEKEY_PERSONNEL = "MODULE_PERSONNEL"
Public Const gsPARAMETERKEY_PERSONNELTABLE = "Param_TablePersonnel"
Public Const gsPARAMETERKEY_EMPLOYEENUMBER = "Param_FieldsEmployeeNumber"
Public Const gsPARAMETERKEY_FORENAME = "Param_FieldsForename"
Public Const gsPARAMETERKEY_SURNAME = "Param_FieldsSurname"
Public Const gsPARAMETERKEY_STARTDATE = "Param_FieldsStartDate"
Public Const gsPARAMETERKEY_LEAVINGDATE = "Param_FieldsLeavingDate"
Public Const gsPARAMETERKEY_FULLPARTTIME = "Param_FieldsFullPartTime"
'Public Const gsPARAMETERKEY_EMAIL = "Param_FieldsEmail"
Public Const gsPARAMETERKEY_DEPARTMENT = "Param_FieldsDepartment"
Public Const gsPARAMETERKEY_GRADE = "Param_FieldsGrade"
Public Const gsPARAMETERKEY_MANAGERSTAFFNO = "Param_FieldsManagerStaffNo"
Public Const gsPARAMETERKEY_JOBTITLE = "Param_FieldsJobTitle"
Public Const gsPARAMETERKEY_LOGINNAME = "Param_FieldsLoginName"
Public Const gsPARAMETERKEY_SECONDLOGINNAME = "Param_FieldsSecondLoginName"

'Region Constants - The following key is used for static region field
Public Const gsPARAMETERKEY_REGION = "Param_FieldsRegion"
'Region Constants - The following keys are used for historical region fields
Public Const gsPARAMETERKEY_HREGIONTABLE = "Param_FieldsHRegionTable"
Public Const gsPARAMETERKEY_HREGIONFIELD = "Param_FieldsHRegion"
Public Const gsPARAMETERKEY_HREGIONDATE = "Param_FieldsHRegionDate"

'WP Constants - The following key is used for static WP field
Public Const gsPARAMETERKEY_WORKINGPATTERN = "Param_FieldsWorkingPattern"
'WP Constants - The following keys are used for historical WP fields
Public Const gsPARAMETERKEY_HWORKINGPATTERNTABLE = "Param_FieldsHWorkingPatternTable"
Public Const gsPARAMETERKEY_HWORKINGPATTERNFIELD = "Param_FieldsHWorkingPattern"
Public Const gsPARAMETERKEY_HWORKINGPATTERNDATE = "Param_FieldsHWorkingPatternDate"

' HIERARCHY MODULE CONSTANTS
Public Const gsMODULEKEY_HIERARCHY = "MODULE_HIERARCHY"
Public Const gsPARAMETERKEY_HIERARCHYTABLE = "Param_TableHierarchy"
Public Const gsPARAMETERKEY_IDENTIFIER = "Param_FieldIdentifier"
Public Const gsPARAMETERKEY_REPORTSTO = "Param_FieldReportsTo"
Public Const gsPARAMETERKEY_POSTALLOCATIONTABLE = "Param_TablePostAllocation"
Public Const gsPARAMETERKEY_POSTALLOCSTARTDATE = "Param_FieldStartDate"
Public Const gsPARAMETERKEY_POSTALLOCENDDATE = "Param_FieldEndDate"

Public glngPersonnelTableID As Long
Public gsPersonnelTableName As String
Private mvar_lngPersonnelEmployeeNumberID As Long
Public gsPersonnelEmployeeNumberColumnName As String
Private mvar_lngPersonnelSurnameID As Long
Public gsPersonnelSurnameColumnName As String
Private mvar_lngPersonnelForenameID As Long
Public gsPersonnelForenameColumnName As String

'Private glngPersonnelStartDateID As Long
Public glngPersonnelStartDateID As Long

Public gsPersonnelStartDateColumnName As String
Private mvar_lngPersonnelLeavingDateID As Long
Public gsPersonnelLeavingDateColumnName As String
Private mvar_lngPersonnelFullPartTimeID As Long
Public gsPersonnelFullPartTimeColumnName As String
'Private mvar_lngPersonnelEmailID As Long
'Public gsPersonnelEmailColumnName As String
Private mvar_lngPersonnelDepartmentID As Long
Public gsPersonnelDepartmentColumnName As String

Private mvar_lngPersonnelGradeID As Long
Public gsPersonnelGradeColumnName As String
Private mvar_lngPersonnelManagerStaffNoID As Long
Public gsPersonnelManagerStaffNoColumnName As String
Private mvar_lngPersonnelJobTitleID As Long
Public gsPersonnelJobTitleColumnName As String


' Static Region
Private mvar_lngPersonnelRegionID As Long
Public gsPersonnelRegionColumnName As String
' Historic Region
Private mvar_lngPersonnelHRegionTableID As Long
Public glngPersonnelHRegionTableID As Long
Public gsPersonnelHRegionTableName As String
Private mvar_lngPersonnelHRegionFieldID As Long
Public gsPersonnelHRegionColumnName As String
Private mvar_lngPersonnelHRegionDateID As Long
Public gsPersonnelHRegionDateColumnName As String
Public gsPersonnelHRegionTableRealSource As String

' Static Working Pattern
Private mvar_lngPersonnelWorkingPatternID As Long
Public gsPersonnelWorkingPatternColumnName As String
' Historic Working Pattern
Private mvar_lngPersonnelHWorkingPatternTableID As Long
Public gsPersonnelHWorkingPatternTableName As String
Private mvar_lngPersonnelHWorkingPatternFieldID As Long
Public gsPersonnelHWorkingPatternColumnName As String
Private mvar_lngPersonnelHWorkingPatternDateID As Long
Public gsPersonnelHWorkingPatternDateColumnName As String
Public gsPersonnelHWorkingPatternTableRealSource As String

Public glngHierarchyTableID As Long
Public gsHierarchyTableName As String
Public glngLoginColumnID As Long
Public glngSecondLoginColumnID As Long
Public glngPostAllocationTableID As Long
Public glngReportsToColumnID As Long
Public glngIdentifyingColumnID As Long

Public Function HierarchyFunctionConfigured(plngFunctionID As Long) As Boolean
  ' Return a boolean value showing if the module parameters are correctly configured
  ' for the given Hierarchy function.
  Dim fValid As Boolean
  Dim fPostBasedSystem As Boolean
  
  fValid = False
  
  If Not gbEnableUDFFunctions Then
    HierarchyFunctionConfigured = False
    Exit Function
  End If

  fPostBasedSystem = (glngPersonnelTableID <> glngHierarchyTableID)

  Select Case plngFunctionID
    Case 67, 71 'HIER_FN_HasPersonnelSubordinate, _
      HIER_FN_IsPersonnelSubordinateOf
      fValid = (glngIdentifyingColumnID > 0) And _
        (glngReportsToColumnID > 0) And _
        ((Not fPostBasedSystem) Or (glngPersonnelTableID > 0)) And _
        ((Not fPostBasedSystem) Or (glngPostAllocationTableID > 0))

    Case 68, 72  'HIER_FN_HasPersonnelSubordinateUser, _
      HIER_FN_IsPersonnelSubordinateOfUser
      fValid = (glngIdentifyingColumnID > 0) And _
        (glngReportsToColumnID > 0) And _
        (glngPersonnelTableID > 0) And _
        (glngLoginColumnID > 0) And _
        ((Not fPostBasedSystem) Or (glngPostAllocationTableID > 0))

    Case 66, 70 'HIER_FN_HasPostSubordinateUser, _
      HIER_FN_IsPostSubordinateOfUser
      fValid = (glngIdentifyingColumnID > 0) And _
        (glngReportsToColumnID > 0) And _
        (glngPersonnelTableID > 0) And _
        (glngLoginColumnID > 0) And _
        (fPostBasedSystem) And _
        (glngPostAllocationTableID > 0)

    Case 65, 69 'HIER_FN_HasPostSubordinate, _
      HIER_FN_IsPostSubordinateOf
      fValid = (glngIdentifyingColumnID > 0) And _
        (glngReportsToColumnID > 0) And _
        (fPostBasedSystem)
  End Select
  
  HierarchyFunctionConfigured = fValid
    
End Function


Public Function IdentifyingColumnDataType() As SQLDataType
  Dim lngIdentifyingColumnID As Long
  Dim datGeneral As DataMgr.clsGeneral

  Set datGeneral = New DataMgr.clsGeneral

  lngIdentifyingColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER))
  
  If lngIdentifyingColumnID = 0 Then
    IdentifyingColumnDataType = sqlUnknown
  Else
    IdentifyingColumnDataType = datGeneral.GetColumnDataType(lngIdentifyingColumnID)
  End If

  Set datGeneral = Nothing

End Function



Public Sub ReadPersonnelParameters()
  
  Dim objTable As CTablePrivilege
  
  ' Read the Personnel module parameters from the database.
  glngPersonnelTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
  If glngPersonnelTableID > 0 Then
    gsPersonnelTableName = datGeneral.GetTableName(glngPersonnelTableID)
  Else
    gsPersonnelTableName = ""
  End If

  mvar_lngPersonnelEmployeeNumberID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER))
  If mvar_lngPersonnelEmployeeNumberID > 0 Then
    gsPersonnelEmployeeNumberColumnName = datGeneral.GetColumnName(mvar_lngPersonnelEmployeeNumberID)
  Else
    gsPersonnelEmployeeNumberColumnName = ""
  End If

  mvar_lngPersonnelSurnameID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
  If mvar_lngPersonnelSurnameID > 0 Then
    gsPersonnelSurnameColumnName = datGeneral.GetColumnName(mvar_lngPersonnelSurnameID)
  Else
    gsPersonnelSurnameColumnName = ""
  End If

  mvar_lngPersonnelForenameID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
  If mvar_lngPersonnelForenameID > 0 Then
    gsPersonnelForenameColumnName = datGeneral.GetColumnName(mvar_lngPersonnelForenameID)
  Else
    gsPersonnelForenameColumnName = ""
  End If

  glngPersonnelStartDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE))
  If glngPersonnelStartDateID > 0 Then
    gsPersonnelStartDateColumnName = datGeneral.GetColumnName(glngPersonnelStartDateID)
  Else
    gsPersonnelStartDateColumnName = ""
  End If

  mvar_lngPersonnelLeavingDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE))
  If mvar_lngPersonnelLeavingDateID > 0 Then
    gsPersonnelLeavingDateColumnName = datGeneral.GetColumnName(mvar_lngPersonnelLeavingDateID)
  Else
    gsPersonnelLeavingDateColumnName = ""
  End If

  mvar_lngPersonnelFullPartTimeID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FULLPARTTIME))
  If mvar_lngPersonnelFullPartTimeID > 0 Then
    gsPersonnelFullPartTimeColumnName = datGeneral.GetColumnName(mvar_lngPersonnelFullPartTimeID)
  Else
    gsPersonnelFullPartTimeColumnName = ""
  End If

  'mvar_lngPersonnelEmailID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMAIL))
  'If mvar_lngPersonnelEmailID > 0 Then
  '  gsPersonnelEmailColumnName = datGeneral.GetColumnName(mvar_lngPersonnelEmailID)
  'Else
  '  gsPersonnelEmailColumnName = ""
  'End If

  mvar_lngPersonnelDepartmentID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT))
  If mvar_lngPersonnelDepartmentID > 0 Then
    gsPersonnelDepartmentColumnName = datGeneral.GetColumnName(mvar_lngPersonnelDepartmentID)
  Else
    gsPersonnelDepartmentColumnName = ""
  End If
  
  mvar_lngPersonnelGradeID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_GRADE))
  If mvar_lngPersonnelGradeID > 0 Then
    gsPersonnelGradeColumnName = datGeneral.GetColumnName(mvar_lngPersonnelGradeID)
  Else
    gsPersonnelGradeColumnName = ""
  End If
  
  mvar_lngPersonnelManagerStaffNoID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_MANAGERSTAFFNO))
  If mvar_lngPersonnelManagerStaffNoID > 0 Then
    gsPersonnelManagerStaffNoColumnName = datGeneral.GetColumnName(mvar_lngPersonnelManagerStaffNoID)
  Else
    gsPersonnelManagerStaffNoColumnName = ""
  End If
  
  mvar_lngPersonnelJobTitleID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_JOBTITLE))
  If mvar_lngPersonnelJobTitleID > 0 Then
    gsPersonnelJobTitleColumnName = datGeneral.GetColumnName(mvar_lngPersonnelJobTitleID)
  Else
    gsPersonnelJobTitleColumnName = ""
  End If


  ' Static Region
  mvar_lngPersonnelRegionID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION))
  If mvar_lngPersonnelRegionID > 0 Then
    gsPersonnelRegionColumnName = datGeneral.GetColumnName(mvar_lngPersonnelRegionID)
    grtRegionType = rtStaticRegion
  Else
    gsPersonnelRegionColumnName = ""
    grtRegionType = rtNotDefined
  End If
  
  ' Historic Region
  mvar_lngPersonnelHRegionTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE))
  glngPersonnelHRegionTableID = mvar_lngPersonnelHRegionTableID
  If mvar_lngPersonnelHRegionTableID > 0 Then
    gsPersonnelHRegionTableName = datGeneral.GetTableName(mvar_lngPersonnelHRegionTableID)
    ' Get the realsource into a variable too
    Set objTable = gcoTablePrivileges.FindTableID(mvar_lngPersonnelHRegionTableID)
    gsPersonnelHRegionTableRealSource = objTable.RealSource
    grtRegionType = rtHistoricRegion
  Else
    gsPersonnelHRegionTableName = ""
    If grtRegionType <> rtStaticRegion Then grtRegionType = rtNotDefined
  End If

  mvar_lngPersonnelHRegionFieldID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD))
  If mvar_lngPersonnelHRegionFieldID > 0 Then
    gsPersonnelHRegionColumnName = datGeneral.GetColumnName(mvar_lngPersonnelHRegionFieldID)
  Else
    gsPersonnelHRegionColumnName = ""
    If grtRegionType <> rtStaticRegion Then grtRegionType = rtNotDefined
  End If
  
  mvar_lngPersonnelHRegionDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE))
  If mvar_lngPersonnelHRegionDateID > 0 Then
    gsPersonnelHRegionDateColumnName = datGeneral.GetColumnName(mvar_lngPersonnelHRegionDateID)
  Else
    gsPersonnelHRegionDateColumnName = ""
    If grtRegionType <> rtStaticRegion Then grtRegionType = rtNotDefined
  End If
  
  ' Static Working Pattern
  mvar_lngPersonnelWorkingPatternID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN))
  If mvar_lngPersonnelWorkingPatternID > 0 Then
    gsPersonnelWorkingPatternColumnName = datGeneral.GetColumnName(mvar_lngPersonnelWorkingPatternID)
    gwptWorkingPatternType = wptStaticWPattern
  Else
    gsPersonnelWorkingPatternColumnName = ""
    gwptWorkingPatternType = wptnotDefined
  End If
  
  ' Historic Working Pattern
  mvar_lngPersonnelHWorkingPatternTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE))
  If mvar_lngPersonnelHWorkingPatternTableID > 0 Then
    gsPersonnelHWorkingPatternTableName = datGeneral.GetTableName(mvar_lngPersonnelHWorkingPatternTableID)
    ' Get the realsource into a variable too
    Set objTable = gcoTablePrivileges.FindTableID(mvar_lngPersonnelHWorkingPatternTableID)
    gsPersonnelHWorkingPatternTableRealSource = objTable.RealSource
    gwptWorkingPatternType = wptHistoricWPattern
  Else
    gsPersonnelHWorkingPatternTableName = ""
    If gwptWorkingPatternType <> wptStaticWPattern Then gwptWorkingPatternType = wptnotDefined
  End If

  mvar_lngPersonnelHWorkingPatternFieldID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD))
  If mvar_lngPersonnelHWorkingPatternFieldID > 0 Then
    gsPersonnelHWorkingPatternColumnName = datGeneral.GetColumnName(mvar_lngPersonnelHWorkingPatternFieldID)
  Else
    gsPersonnelHWorkingPatternColumnName = ""
    If gwptWorkingPatternType <> wptStaticWPattern Then gwptWorkingPatternType = wptnotDefined
  End If
  
  mvar_lngPersonnelHWorkingPatternDateID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE))
  If mvar_lngPersonnelHWorkingPatternDateID > 0 Then
    gsPersonnelHWorkingPatternDateColumnName = datGeneral.GetColumnName(mvar_lngPersonnelHWorkingPatternDateID)
  Else
    gsPersonnelHWorkingPatternDateColumnName = ""
    If gwptWorkingPatternType <> wptStaticWPattern Then gwptWorkingPatternType = wptnotDefined
  End If
  
  glngHierarchyTableID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE))
  If glngHierarchyTableID > 0 Then
    gsHierarchyTableName = datGeneral.GetTableName(glngHierarchyTableID)
  Else
    gsHierarchyTableName = ""
  End If
  
  glngIdentifyingColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER))
  glngReportsToColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_REPORTSTO))
  glngPostAllocationTableID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCATIONTABLE))
  glngLoginColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME))
  glngSecondLoginColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECONDLOGINNAME))
  
  If (glngLoginColumnID = 0) And (glngSecondLoginColumnID > 0) Then
    glngLoginColumnID = glngSecondLoginColumnID
    glngSecondLoginColumnID = 0
  End If
  
  Set objTable = Nothing

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
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
        "The Personnel table is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Employee Number ID is valid.
    fValid = (mvar_lngPersonnelEmployeeNumberID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Employee Number column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Surname ID is valid.
    fValid = (mvar_lngPersonnelSurnameID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Surname column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Forename ID is valid.
    fValid = (mvar_lngPersonnelForenameID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Forename column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the StartDate ID is valid.
    fValid = (glngPersonnelStartDateID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Start Date column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Leaving Date ID is valid.
    fValid = (mvar_lngPersonnelLeavingDateID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Leaving Date column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the FullPartTime ID is valid.
    fValid = (mvar_lngPersonnelFullPartTimeID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Full/Part Time column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  'If fValid Then
  '  ' Check the Email ID is valid.
  '  fValid = (mvar_lngPersonnelEmailID > 0)
  '  If Not fValid Then
  '    COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
  '       "The Email column is not defined.", vbOKOnly, App.ProductName
  '  End If
  'End If

  If fValid Then
    ' Check the Department Date ID is valid.
    fValid = (mvar_lngPersonnelDepartmentID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Department column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Working Pattern Date ID is valid.
    fValid = (mvar_lngPersonnelWorkingPatternID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Working Pattern column is not defined.", vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Check the Region ID is valid.
    fValid = (mvar_lngPersonnelRegionID > 0)
    If Not fValid Then
      COAMsgBox "The Personnel module is not properly configured." & vbCrLf & _
         "The Region column is not defined.", vbOKOnly, App.ProductName
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
'      COAMsgBox "You do not have permission to see the defined Course Title column.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'
'  If fValid And (Len(gsCourseCancelledByColumnName) > 0) Then
'    ' Check the Course Cancelled By column can be updated if it is defined.
'    fValid = objCourseColumnPrivileges.Item(gsCourseCancelledByColumnName).AllowUpdate
'    If Not fValid Then
'      COAMsgBox "You do not have permission to update the defined Course Cancelled By column.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'
'
'  ' Return the validation value.
  ValidatePersonnelParameters = fValid

End Function


