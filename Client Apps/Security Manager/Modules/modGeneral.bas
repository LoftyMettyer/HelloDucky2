Attribute VB_Name = "modGeneral"
Public Const gsMODULEKEY_PERSONNEL = "MODULE_PERSONNEL"
Public Const gsPARAMETERKEY_PERSONNELTABLE = "Param_TablePersonnel"
Public Const gsPARAMETERKEY_FORENAME = "Param_FieldsForename"
Public Const gsPARAMETERKEY_SURNAME = "Param_FieldsSurname"
Public Const gsPARAMETERKEY_LOGINNAME = "Param_FieldsLoginName"
Public Const gsPARAMETERKEY_SECONDLOGINNAME = "Param_FieldsSecondLoginName"
Public Const gsPARAMETERKEY_LEAVINGDATE = "Param_FieldsLeavingDate"

' HIERARCHY MODULE CONSTANTS
Public Const gsMODULEKEY_HIERARCHY = "MODULE_HIERARCHY"
Public Const gsPARAMETERKEY_HIERARCHYTABLE = "Param_TableHierarchy"
Public Const gsPARAMETERKEY_IDENTIFIER = "Param_FieldIdentifier"
Public Const gsPARAMETERKEY_REPORTSTO = "Param_FieldReportsTo"
Public Const gsPARAMETERKEY_POSTALLOCATIONTABLE = "Param_TablePostAllocation"

Public glngPersonnelTableID As Long
Public glngForenameColumnID As Long
Public glngSurnameColumnID As Long
Public glngLoginColumnID As Long
Public glngSecondLoginColumnID As Long
Public glngHierarchyTableID As Long
Public glngPostAllocationTableID As Long
Public glngReportsToColumnID As Long
Public glngIdentifyingColumnID As Long

Public Enum CreateUserMode
  iUSERCREATE_SQLLOGIN = 0
  iUSERCREATE_WINDOWSMANUAL = 1
  iUSERCREATE_WINDOWSAUTO = 2
End Enum

Public Enum CreateUserStatus
  iSUCCESS = 0
  iFAILED_USERNAMEISBLANK = 1
  iFAILED_USERNAMEISRESERVED = 2
  iFAILED_USERNAMEISKEYWORD = 3
  iFAILED_USERNAMEISUSED = 4
  iWARNING_LOGINEXISTS = 5
  iFAILED_ILLEGALPASSWORD = 6
  iFAILED_USERNAMEISNUMERIC = 7
  iFAILED_PASSWORDNOTMINIMUM = 8
  iFAILED_USERNAMEGREATERTHANLOGINSIZE = 9
  iFAILED_USERNAMEISTOOLONG = 10
  iFAILED_NTACCOUNTNOTEXIST = 11
  iFAILED_PASSWORDNOTCOMPLEX = 12
  iSUCCESS_USERALREADYADDED = 13
End Enum

Public Enum LoginType
  iUSERTYPE_SQLLOGIN = 1
  iUSERTYPE_TRUSTEDUSER = 2
  iUSERTYPE_TRUSTEDGROUP = 3
  ' NPG20090206 Fault 11931
  iUSERTYPE_ORPHANUSER = 4
  iUSERTYPE_ORPHANGROUP = 5
End Enum

Public Enum enum_Module
  modPersonnel = 1
  modRecruitment = 2 ^ 1
  modAbsence = 2 ^ 2
  modTraining = 2 ^ 3
  modIntranet = 2 ^ 4
  modAFD = 2 ^ 5
  modFullSysMgr = 2 ^ 6
  modCMG = 2 ^ 7
  modQAddress = 2 ^ 8
  modAccord = 2 ^ 9
  modWorkflow = 2 ^ 10
  modVersionOne = 2 ^ 11
  modMobile = 2 ^ 12
  modFusion = 2 ^ 13
  modXMLExport = 2 ^ 14
  mod3rdPartyTables = 2 ^ 15
  modNineBoxGrid = 2 ^ 16
  modEditableGrids = 2 ^ 17
  modCustomisationPowerPack = 2 ^ 18
  modTalentReports = 2 ^ 19
End Enum

Public Enum SQLServerAuthenticationType
  iWINDOWSONLY = 1
  iMIXEDMODE = 2
End Enum

Public glngPageNum As Long
Public gstrPrintGroupName As String

Public Const giMAXIMUMUSERNAMELENGTH = 50

Public giWindowState As FormWindowStateConstants
Public glngWindowLeft As Long
Public glngWindowTop As Long
Public glngWindowHeight As Long
Public glngWindowWidth As Long

Public gobjLicence As New clsLicence

Public Enum ArraySortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Public Sub DisplayApplication()
  'JPD 20030908 Fault 5756
  
  'JPD 20030917 Fault 6991
  If (glngWindowLeft = 0) And _
    (glngWindowTop = 0) And _
    (glngWindowWidth = 0) And _
    (glngWindowHeight = 0) Then
    
    Exit Sub
  End If
  
  frmMain.WindowState = IIf(giWindowState = vbMinimized, vbNormal, giWindowState)
  If frmMain.WindowState = vbNormal Then
    frmMain.Left = glngWindowLeft
    frmMain.Top = glngWindowTop
    frmMain.Width = glngWindowWidth
    frmMain.Height = glngWindowHeight
  End If

End Sub

Public Function UniqueSQLObjectName(strPrefix As String, intType As Integer) As String
  
  Dim cmdUniqOBJ As New ADODB.Command
  Dim pmADO As ADODB.Parameter

  With cmdUniqOBJ
    .CommandText = "sp_ASRUniqueObjectName"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Unique", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Prefix", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = strPrefix

    Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = intType

    .Execute

    UniqueSQLObjectName = IIf(IsNull(.Parameters(0).Value), "", .Parameters(0).Value)
  End With
  Set cmdUniqOBJ = Nothing


End Function
Public Function DropUniqueSQLObject(sSQLObjectName As String, iType As Integer) As Boolean

  On Error GoTo ErrorTrap

  Dim cmdUniqOBJ As New ADODB.Command
  Dim pmADO As ADODB.Parameter
 
  If Len(sSQLObjectName) > 0 Then

    With cmdUniqOBJ
      .CommandText = "sp_ASRDropUniqueObject"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("Unique", adVarChar, adParamInput, 255)
      .Parameters.Append pmADO
      pmADO.Value = sSQLObjectName
  
      Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = iType
  
      .Execute
  
    End With
    
  End If
    
  DropUniqueSQLObject = True
   
    
TidyUpAndExit:
  Set cmdUniqOBJ = Nothing
  Exit Function
ErrorTrap:
  DropUniqueSQLObject = False
  GoTo TidyUpAndExit
  
End Function

Public Sub LoadTableCombo(cboTemp As ComboBox, Optional strSQL As String)
  
  Dim rsTemp As New ADODB.Recordset

  If strSQL = vbNullString Then
    strSQL = "SELECT TableID, TableName FROM ASRSysTables " & _
             " ORDER BY TableName"
  End If
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  With cboTemp
    .Clear
    
    Do While Not rsTemp.EOF
      .AddItem rsTemp!TableName
      .ItemData(.NewIndex) = rsTemp!TableID
      rsTemp.MoveNext
    Loop

    If .ListCount > 0 Then
      SetComboItem cboTemp, glngPersonnelTableID
      If .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End If
    
  End With

  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Public Sub SetComboItem(cboCombo As ComboBox, lItem As Long)

  Dim lCount As Long
  
  With cboCombo
    For lCount = 1 To .ListCount
      If .ItemData(lCount - 1) = lItem Then
        .ListIndex = lCount - 1
        Exit For
      End If
    Next
  End With

End Sub

Public Function SetupModuleParameters()

  'Load the settings for personnel records
  glngPersonnelTableID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
  
  glngLoginColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME))
  glngSecondLoginColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECONDLOGINNAME))
  
  If (glngLoginColumnID = 0) And (glngSecondLoginColumnID > 0) Then
    glngLoginColumnID = glngSecondLoginColumnID
    glngSecondLoginColumnID = 0
  End If
  
  glngForenameColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
  glngSurnameColumnID = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
  glngHierarchyTableID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE))
  
  glngPostAllocationTableID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCATIONTABLE))
  glngReportsToColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_REPORTSTO))
  glngIdentifyingColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER))
  
End Function

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

    Case 66, 70  'HIER_FN_HasPostSubordinateUser, _
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

  lngIdentifyingColumnID = Val(GetModuleParameter(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER))
  
  If lngIdentifyingColumnID = 0 Then
    IdentifyingColumnDataType = sqlUnknown
  Else
    IdentifyingColumnDataType = GetColumnDataType(lngIdentifyingColumnID)
  End If

End Function




Public Function GetModuleParameter(psModuleKey As String, psParameterKey As String) As String
  ' Return the value of the given module parameter.
  Dim sSQL As String
  Dim rsModule As New ADODB.Recordset
        
  sSQL = "SELECT parameterValue" & _
    " FROM ASRSysModuleSetup" & _
    " WHERE moduleKey = '" & psModuleKey & "'" & _
    " AND parameterKey = '" & psParameterKey & "'"
  rsModule.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
  If Not (rsModule.BOF And rsModule.EOF) Then
    If IsNull(rsModule!parameterValue) Then
      GetModuleParameter = vbNullString
    Else
      GetModuleParameter = rsModule!parameterValue
    End If
  Else
    GetModuleParameter = vbNullString
  End If
  rsModule.Close
  
  Set rsModule = Nothing

End Function

Public Function ConvertNumberForDisplay(ByVal strInput As String) As String
  'Get a number in the correct format for display
  '(e.g. on french systems replace decimal point for a decimal comma)
  ConvertNumberForDisplay = Replace(strInput, ".", UI.GetSystemDecimalSeparator)
End Function

Public Sub FormatTDBNumberControl(objInput As Object)

  If TypeOf objInput Is TDBNumber6Ctl.TDBNumber Then
    objInput.Separator = "x"
    objInput.DecimalPoint = UI.GetSystemDecimalSeparator
    objInput.Separator = UI.GetSystemThousandSeparator
  End If

End Sub


Public Function CalcIsReadOnly(lExprID As Long) As Boolean

  On Error GoTo ErrTrap
  Dim rsTemp As New ADODB.Recordset
  
  rsTemp.Open "SELECT * FROM ASRSysExpressions WHERE ExprID = " & CStr(lExprID), gADOCon, adOpenForwardOnly, adLockReadOnly
  CalcIsReadOnly = (rsTemp!Access <> ACCESS_READWRITE And _
                    UCase(rsTemp!UserName) <> UCase(gsUserName))

ExitPoint:
  Set rsTemp = Nothing
  Exit Function
  
ErrTrap:
  MsgBox "Error determining Access rights of selected calc:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume ExitPoint
End Function

Public Function GetColumnTableName(plngColumnID As Long) As String
  
  ' Return the table id of the given column.
  Dim sSQL As String
  Dim rsData As New ADODB.Recordset

  sSQL = "SELECT tableName" & _
    " FROM ASRSysColumns " & _
    " JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID" & _
    " WHERE columnID = " & Trim(Str(plngColumnID))
  rsData.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnTableName = rsData.Fields(0).Value
  Else
    GetColumnTableName = ""
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function IsChildOfTable(plngParentTableID As Long, plngChildTableID As Long) As Boolean
  'Checks if the passed child table is a child of the passed parent table.
  Dim sSQL As String
  Dim rsTemp As New ADODB.Recordset

  sSQL = "SELECT COUNT(*) AS [result]" & _
    " FROM ASRSysRelations" & _
    " WHERE parentID = " & CStr(plngParentTableID) & _
    "   AND childID = " & CStr(plngChildTableID)
    
  rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsTemp
    IsChildOfTable = (!Result > 0)
    
    .Close
  End With
  Set rsTemp = Nothing

End Function


Public Function GetColumnSize(plngColumnID As Long) As Long
  
  ' Return the size of the given column.
  Dim sSQL As String
  Dim rsData As New ADODB.Recordset

  sSQL = "SELECT size" & _
    " FROM ASRSysColumns " & _
    " WHERE columnID = " & Trim(Str(plngColumnID))
  rsData.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnSize = IIf(IsNull(rsData!Size), 0, rsData!Size)
  Else
    GetColumnSize = 0
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function GetColumnDataType(plngColumnID As Long) As SQLDataType
  
  ' Return the size of the given column.
  Dim sSQL As String
  Dim rsData As New ADODB.Recordset

  sSQL = "SELECT dataType" & _
    " FROM ASRSysColumns " & _
    " WHERE columnID = " & Trim(Str(plngColumnID))
  rsData.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnDataType = IIf(IsNull(rsData!DataType), sqlUnknown, rsData!DataType)
  Else
    GetColumnDataType = sqlUnknown
  End If

  rsData.Close
  Set rsData = Nothing

End Function


Public Function CalculateBottomOfPage() As Long
With Printer
    CalculateBottomOfPage = .ScaleHeight - (giPRINT_YINDENT)
End With
End Function
Public Function CheckEndOfPage2(Optional mlngBottom As Long, Optional fReset As Boolean) As Boolean
  If Printer.CurrentY > mlngBottom Then
    Call FooterText2
    Printer.NewPage
    
    If fReset Then glngPageNum = 0
    
    Printer.CurrentY = giPRINT_YINDENT
    Printer.CurrentX = giPRINT_XINDENT
    
    ' Flag that page has changed
    CheckEndOfPage2 = True
  End If
End Function
Public Sub ForceEndOfPage()
  Call FooterText2
  Printer.NewPage
    
  Printer.CurrentY = giPRINT_YINDENT
  Printer.CurrentX = giPRINT_XINDENT
    
End Sub

Public Function FooterText2()
  
  Dim strPageNum As String
  
  glngPageNum = glngPageNum + 1
  strPageNum = "Page " & CStr(glngPageNum)

  Printer.FontSize = 8
  Printer.Print " "
  Printer.FontBold = False
  
  Printer.CurrentX = giPRINT_XINDENT
  Printer.Print "Printed on " & Format(Now, DateFormat) & _
                " at " & Format(Now, "hh:nn") & " by " & gsUserName;
  
  Printer.CurrentX = (Printer.ScaleWidth - giPRINT_XINDENT) - Printer.TextWidth(strPageNum)
  Printer.Print strPageNum

  Printer.FontSize = 10

End Function

Public Function SetStringLength(ByVal psInputString As String, piLength As Integer) As String

  ' Sets an input field to the desired length
  ' i.e. trims or pads out with spaces

  If Len(psInputString) > piLength Then
    SetStringLength = Left(psInputString, piLength)
  Else
    SetStringLength = psInputString & Space(piLength - Len(psInputString))
  End If

End Function

Public Function IsInArray(pastrInArray() As String, pstrSearchString As String) As Boolean
   
  Dim iCount As Integer
   
  IsInArray = False
  For iCount = LBound(pastrInArray) To UBound(pastrInArray)
    If LCase(pastrInArray(iCount)) = LCase(pstrSearchString) Then
      IsInArray = True
      Exit For
    End If
  Next iCount

End Function

Public Function GetUsersInWindowsGroup(pstrGroupName As String) As String()

  On Error GoTo ErrorTrap

  Dim astrUsers() As String
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  
  sSQL = "SELECT * FROM OpenRowSet (NetGroupGetMembers, '" & pstrGroupName & "')"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  ReDim astrUsers(0)
  Do While Not rsGroups.EOF
    ReDim Preserve astrUsers(UBound(astrUsers) + 1)
    astrUsers(UBound(astrUsers)) = Trim(rsGroups!Domain) & "\" & Trim(rsGroups!Name)
    rsGroups.MoveNext
  Loop

  rsGroups.Close
  
TidyUpAndExit:
  Set rsGroups = Nothing
  GetUsersInWindowsGroup = astrUsers
  Exit Function

ErrorTrap:
  ReDim Preserve astrUsers(0)
  GoTo TidyUpAndExit

End Function

' Sorts the array in the specified order
Public Sub SortArray(ByRef vArray As Variant, Optional ByVal sortOrder As ArraySortOrder = SortAscending)
   Dim i          As Long
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim vTemp      As Variant
   Dim distance   As Long
   Dim bSortOrder As Boolean
   
   If Not IsArray(vArray) Then Exit Sub
   
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   bSortOrder = IIf(sortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1
   
   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3
      For i = distance + iLBound To iUBound
         vTemp = vArray(i)
         j = i
         Do While (vArray(j - distance) > vTemp) Xor bSortOrder
            vArray(j) = vArray(j - distance)
            j = j - distance
            If j - distance < iLBound Then Exit Do
         Loop
         vArray(j) = vTemp
      Next i
   Loop Until distance = 1
End Sub

Public Function CheckPasswordComplexity(ByVal strUserName As String, ByVal pstrPassword As String) As Boolean

  Dim bCheck As Boolean
  Dim iValids As Integer
  
  Dim bLower As Boolean
  Dim bUpper As Boolean
  Dim bNumeric As Boolean
  Dim bFunny As Boolean
  Dim bUserNameInPart As Boolean
  Dim strUserNamePart As String
  
  Dim iCount As Integer
  Dim iCurrentChar As String
  
  bCheck = (giDomainComplexity > 0)

  If bCheck Then
    
    ' No part* of the username in the password (* part being 3 or more consecutive characters)
    For iCount = 1 To Len(strUserName) - 2
      strUserNamePart = Mid$(strUserName, iCount, 3)
      If InStrB(1, pstrPassword, strUserNamePart, vbBinaryCompare) Then
        bUserNameInPart = True
      End If
    Next iCount
    
    ' Has the right sort of character!
    For iCount = 1 To Len(pstrPassword)
      iCurrentChar = AscW(Mid$(pstrPassword, iCount, 1))
    
      ' Check for a uppercase
      If iCurrentChar >= 65 And iCurrentChar <= 90 Then
        iValids = IIf(bUpper = False, iValids + 1, iValids)
        bUpper = True
      
      ' Check for a lowercase
      ElseIf iCurrentChar >= 97 And iCurrentChar <= 122 Then
        iValids = IIf(bLower = False, iValids + 1, iValids)
        bLower = True
      
      ' Check for a number (0-9)
      ElseIf iCurrentChar >= 48 And iCurrentChar <= 57 Then
        iValids = IIf(bNumeric = False, iValids + 1, iValids)
        bNumeric = True
    
      ' Check for a Non–alphanumeric (For example: !, $, #, or %)
      ElseIf bFunny = False Then
        iValids = iValids + 1
        bFunny = True
      End If
    
    Next iCount
     
    ' Have we done enough
    CheckPasswordComplexity = (iValids > 2 And Len(pstrPassword) > 6 And bUserNameInPart = False)
  
  Else
    CheckPasswordComplexity = True
  End If

End Function

