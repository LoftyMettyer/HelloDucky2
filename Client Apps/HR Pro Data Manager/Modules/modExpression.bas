Attribute VB_Name = "modExpression"
Option Explicit

' Expression Type constants
' NB. If you modify this enum, you'll need to do the same to the mathcing enums in:
'     System Manager - Application
'     Security Manager - modExpression
'     Intranet Server DLL - modExpression
Public Enum ExpressionTypes
  giEXPR_UNKNOWNTYPE = 0
  giEXPR_COLUMNCALCULATION = 1
  giEXPR_GOTFOCUS = 2             ' Not used.
  giEXPR_RECORDVALIDATION = 3
  giEXPR_DEFAULTVALUE = 4         ' Not used.
  giEXPR_STATICFILTER = 5
  giEXPR_PAGEBREAK = 6            ' Not used.
  giEXPR_ORDER = 7                ' Not used.
  giEXPR_RECORDDESCRIPTION = 8
  giEXPR_VIEWFILTER = 9
  giEXPR_RUNTIMECALCULATION = 10
  giEXPR_RUNTIMEFILTER = 11
  giEXPR_EMAIL = 12               ' System Manager Only
  giEXPR_LINKFILTER = 13          ' System Manager Only
  giEXPR_UTILRUNTIMEFILTER = 14   'Import filter
  giEXPR_MATCHJOINEXPRESSION = 15
  giEXPR_MATCHSCOREEXPRESSION = 16
  giEXPR_MATCHWHEREEXPRESSION = 17
  giEXPR_RECORDINDEPENDANTCALC = 18
  giEXPR_OUTLOOKFOLDER = 19         'System Manager Only
  giEXPR_OUTLOOKSUBJECT = 20        'System Manager Only
  giEXPR_WORKFLOWCALCULATION = 21   'System Manager Only
  giEXPR_WORKFLOWSTATICFILTER = 22  'System Manager Only
  giEXPR_WORKFLOWRUNTIMEFILTER = 23 'System Manager Only
End Enum

' Expression Value types
' NB. If you modify this enum, you'll need to do the same to the mathcing enums in:
'     System Manager - Application
'     Security Manager - modExpression
'     Intranet Server DLL - modExpression
Public Enum ExpressionComponentTypes
  giCOMPONENT_FIELD = 1
  giCOMPONENT_FUNCTION = 2
  giCOMPONENT_CALCULATION = 3
  giCOMPONENT_VALUE = 4
  giCOMPONENT_OPERATOR = 5
  giCOMPONENT_TABLEVALUE = 6
  giCOMPONENT_PROMPTEDVALUE = 7
  giCOMPONENT_CUSTOMCALC = 8  ' Not used.
  giCOMPONENT_EXPRESSION = 9
  giCOMPONENT_FILTER = 10
  giCOMPONENT_WORKFLOWVALUE = 11
  giCOMPONENT_WORKFLOWFIELD = 12
End Enum

Public Enum FieldSelectionTypes
  giSELECT_FIRSTRECORD = 1
  giSELECT_LASTRECORD = 2
  giSELECT_SPECIFICRECORD = 3
  giSELECT_RECORDTOTAL = 4
  giSELECT_RECORDCOUNT = 5
End Enum

Public Enum FieldPassTypes
  giPASSBY_VALUE = 1
  giPASSBY_REFERENCE = 2
End Enum

Public Const giPRINT_XINDENT = 1000
Public Const giPRINT_YINDENT = 1000
Public Const giPRINT_XSPACE = 500
Public Const giPRINT_YSPACE = 100

Public Enum ExprValidationCodes
  giEXPRVALIDATION_NOERRORS = 0
  giEXPRVALIDATION_MISSINGOPERAND = 1
  giEXPRVALIDATION_SYNTAXERROR = 2
  giEXPRVALIDATION_EXPRTYPEMISMATCH = 3
  giEXPRVALIDATION_UNKNOWNERROR = 4
  giEXPRVALIDATION_OPERANDTYPEMISMATCH = 5
  giEXPRVALIDATION_PARAMETERTYPEMISMATCH = 6
  giEXPRVALIDATION_NOCOMPONENTS = 7
  giEXPRVALIDATION_PARAMETERSYNTAXERROR = 8
  giEXPRVALIDATION_PARAMETERNOCOMPONENTS = 9
  giEXPRVALIDATION_FILTEREVALUATION = 10
  giEXPRVALIDATION_SQLERROR = 11          ' JPD20020419 Fault 3687
  giEXPRVALIDATION_ASSOCSQLERROR = 12     ' JPD20020419 Fault 3687
End Enum

Public Enum ExpressionValueTypes
  giEXPRVALUE_UNDEFINED = 0
  giEXPRVALUE_CHARACTER = 1
  giEXPRVALUE_NUMERIC = 2
  giEXPRVALUE_LOGIC = 3
  giEXPRVALUE_DATE = 4
  giEXPRVALUE_TABLEVALUE = 5
  giEXPRVALUE_OLE = 6
  giEXPRVALUE_PHOTO = 7
  
  giEXPRVALUE_BYREF_UNDEFINED = 100
  giEXPRVALUE_BYREF_CHARACTER = 101
  giEXPRVALUE_BYREF_NUMERIC = 102
  giEXPRVALUE_BYREF_LOGIC = 103
  giEXPRVALUE_BYREF_DATE = 104
  giEXPRVALUE_BYREF_TABLEVALUE = 105 ' Not used.
  giEXPRVALUE_BYREF_OLE = 106 ' Not used.
  giEXPRVALUE_BYREF_PHOTO = 107 ' Not used.
End Enum
Public Const giEXPRVALUE_BYREF_OFFSET = 100

Public Const gsDUMMY_CHARACTER = "ASRDUMMYCHARVALUE"
Public Const gsDUMMY_NUMERIC = 1
Public Const gsDUMMY_LOGIC = True
Public Const gsDUMMY_DATE = #1/1/1998#
Public Const gsDUMMY_BYREF_CHARACTER = sqlVarChar & vbTab & "a"
Public Const gsDUMMY_BYREF_NUMERIC = sqlNumeric & vbTab & "1"
Public Const gsDUMMY_BYREF_LOGIC = sqlBoolean & vbTab & "0"
Public Const gsDUMMY_BYREF_DATE = sqlDate & vbTab & "1/1/1998"

' Order object constants.
Public Enum OrderTypes
  giORDERTYPE_STATIC = 0
  giORDERTYPE_DYNAMIC = 1
End Enum

' Parameter Type constants.
Public Const gsPARAMETERTYPE_ORDERID = "PType_OrderID"

'View Expression constants
Public Enum ExpressionColour
  EXPRESSIONBUILDER_COLOUROFF = 1
  EXPRESSIONBUILDER_COLOURON = 2
  EXPRESSIONBUILDER_COLOURLASTSAVE = 3
End Enum

Public Enum ExpressionSaveView
  EXPRESSIONBUILDER_NODESMINIMIZE = 1
  EXPRESSIONBUILDER_NODESEXPAND = 2
  EXPRESSIONBUILDER_NODESLASTSAVE = 3
  EXPRESSIONBUILDER_NODESTOPLEVEL = 4
End Enum

Public gvLastPromptedValue As Variant

Public Function HasExpressionComponent(plngExprIDBeingSearched As Long, plngExprIDSearchedFor As Long) As Boolean
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap

  Dim rsExprComp As ADODB.Recordset
  Dim rsExpr As ADODB.Recordset
  Dim fHasExpr As Boolean
  Dim sSQL As String
  Dim lngSubExprID As Long
  
  HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)
  
  If Not HasExpressionComponent Then
    sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
    Set rsExprComp = datGeneral.GetRecords(sSQL)
    
    With rsExprComp
      Do Until .EOF
        Select Case !Type
          Case giCOMPONENT_CALCULATION
            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
      
          Case giCOMPONENT_FILTER
            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
      
          Case giCOMPONENT_FIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
        
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
            Set rsExpr = datGeneral.GetRecords(sSQL)
            Do Until rsExpr.EOF
              HasExpressionComponent = HasExpressionComponent(rsExpr!ExprID, plngExprIDSearchedFor)
              
              If HasExpressionComponent Then
                Exit Do
              End If
              
              rsExpr.MoveNext
            Loop
            rsExpr.Close
            Set rsExpr = Nothing
        End Select
        
        If HasExpressionComponent Then
          Exit Do
        End If
        
        .MoveNext
      Loop
    End With
  
    rsExprComp.Close
  End If
  
TidyUpAndExit:
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Function




Public Function HasFunctionComponent(ByRef plngExprID As Long, ByRef plngFunctionID As Long) As Boolean
  
  On Error GoTo ErrorTrap

  Dim rsExpr As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT dbo.udf_ASRHasFunctionComponent (" & plngExprID & "," & plngFunctionID & ")"
  
  Set rsExpr = New ADODB.Recordset
  rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  HasFunctionComponent = rsExpr.Fields(0).Value
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Exit Function
  
ErrorTrap:
  HasFunctionComponent = False
  Resume TidyUpAndExit

End Function

Public Function HasHiddenComponents(lngExprID As Long) As Boolean

'********************************************************************************
' HasHiddenComponents - Loops through the passed expression searching for       *
'                       hidden expressions (calcs/filters).                     *
'                       Note: This function calls itself and drills down the    *
'                       expression checking for hidden calcs & filters, then    *
'                       works its way up the expressions/components.            *
'                                                                               *
' 'TM20010802 Fault 2617                                                        *
'********************************************************************************

  Dim rsExpr As ADODB.Recordset
  Dim rsExprComp As ADODB.Recordset
  Dim lngCalcFilterID As Long
  Dim bHasHiddenComp As Boolean
  Dim sStartAccess As String
  Dim sSQL As String

  On Error GoTo ErrorTrap
  
'''  sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
'''  Set rsExpr = datGeneral.GetRecords(sSQL)
 
  sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & lngExprID
  Set rsExprComp = datGeneral.GetRecords(sSQL)
  
  bHasHiddenComp = False
  
  With rsExprComp
    Do Until .EOF
      Select Case !Type
        Case giCOMPONENT_CALCULATION
          lngCalcFilterID = IIf(IsNull(!CalculationID), 0, !CalculationID)
      
          If lngCalcFilterID > 0 Then
            If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
              bHasHiddenComp = True
              'TM20011003
              'Need this function to just find out if there are any hidden components,
              'it was also setting the access of the functions and therefore changing
              'time stamp.
              'SetExprAccess lngCalcFilterID, "HD"
            End If
          End If
      
        Case giCOMPONENT_FILTER
          lngCalcFilterID = IIf(IsNull(!FilterID), 0, !FilterID)
      
          If lngCalcFilterID > 0 Then
            If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
              bHasHiddenComp = True
              'TM20011003
              'Need this function to just find out if there are any hidden components,
              'it was also setting the access of the functions and therefore changing
              'time stamp.
              'SetExprAccess lngCalcFilterID, "HD"
            End If
          End If
      
        Case giCOMPONENT_FIELD
          lngCalcFilterID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
      
          If lngCalcFilterID > 0 Then
            If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
              bHasHiddenComp = True
              'TM20011003
              'Need this function to just find out if there are any hidden components,
              'it was also setting the access of the functions and therefore changing
              'time stamp.
              'SetExprAccess lngCalcFilterID, "HD"
            End If
          End If
      
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
            Set rsExpr = datGeneral.GetRecords(sSQL)
            Do Until rsExpr.EOF
              If HasHiddenComponents(rsExpr!ExprID) Or GetExprField(rsExpr!ExprID, "Access") = ACCESS_HIDDEN Then
                bHasHiddenComp = True
                Exit Do
              End If
              
              rsExpr.MoveNext
            Loop
            rsExpr.Close
            Set rsExpr = Nothing
        End Select
      
        If bHasHiddenComp Then
          Exit Do
        End If
      
      .MoveNext
    Loop
  End With

  'TM20011003
  'Need this function to just find out if there are any hidden components,
  'it was also setting the access of the functions and therefore changing
  'time stamp.
'  If bHasHiddenComp Then SetExprAccess lngExprID, "HD"
  HasHiddenComponents = bHasHiddenComp

  rsExprComp.Close
'''  rsExpr.Close
  
TidyUpAndExit:
'''  Set rsExpr = Nothing
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  HasHiddenComponents = False
  Resume TidyUpAndExit

End Function


Public Function GetExprField(lngExprID As Long, sField As String) As Variant

  Dim sSQL As String
  Dim rsExpr As ADODB.Recordset
  
  On Error GoTo ErrorTrap
  
  sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
  
  Set rsExpr = datGeneral.GetRecords(sSQL)
  
  With rsExpr
    If .RecordCount > 0 Then
      GetExprField = .Fields(sField).Value
    End If
  End With
  
  rsExpr.Close
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error retrieving field value from database.", vbOKOnly + vbCritical, App.Title
  Resume TidyUpAndExit
    
End Function

Public Function GetPickListField(lngPicklistID As Long, sField As String) As Variant

  Dim sSQL As String
  Dim rsExpr As ADODB.Recordset
  
  On Error GoTo ErrorTrap
  
  sSQL = "SELECT * FROM ASRSysPickListName WHERE PickListID = " & lngPicklistID
  
  Set rsExpr = datGeneral.GetRecords(sSQL)
  
  With rsExpr
    If .RecordCount > 0 Then
      GetPickListField = .Fields(sField).Value
    End If
  End With
  
  rsExpr.Close
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error retrieving field value from database.", vbOKOnly + vbCritical, App.Title
  Resume TidyUpAndExit
    
End Function

Public Function SetExprAccess(lngExprID As Long, sAccessType As String) As Boolean

  Dim datData As clsDataAccess
  Dim sSQL As String

  On Error GoTo ErrorTrap

  Set datData = New clsDataAccess
  
  sSQL = "UPDATE ASRSysExpressions " & _
           "SET Access = '" & sAccessType & "' " & _
           "WHERE ExprID = " & lngExprID
           
  datData.ExecuteSql sSQL

TidyUpAndExit:
  Set datData = Nothing
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error setting Expression Access.", vbOKOnly + vbCritical, App.Title
  Resume TidyUpAndExit
  
End Function

Public Function isOwnerOfParent(objComp As clsExprComponent) As Boolean
  
'********************************************************************************
' isOwnerOfParent - checks if the current user is the owner of the current      *
'                   expression.
'********************************************************************************
  
  On Error GoTo ErrorTrap
  
  With objComp
    isOwnerOfParent = (LCase(Trim(gsUserName)) = LCase(Trim(.ParentExpression.Owner)))
  End With

TidyUpAndExit:

  Exit Function

ErrorTrap:
  isOwnerOfParent = False
  COAMsgBox "Error checking owner of expression.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit
  
End Function

Public Function isOwnerOfComp(objComp As clsExprComponent) As Boolean

'********************************************************************************
' isOwnerOfParent - checks if the current user is the owner of the current      *
'                   component.
'********************************************************************************

  Dim lngExprID As Long

  On Error GoTo ErrorTrap
  
  With objComp
    If (.ComponentType = giCOMPONENT_CALCULATION) _
        Or (.ComponentType = giCOMPONENT_FILTER) _
        Or (.ComponentType = giCOMPONENT_FIELD) Then
      
      Select Case .ComponentType
        Case giCOMPONENT_CALCULATION
          lngExprID = .Component.CalculationID
        Case giCOMPONENT_FILTER
          lngExprID = .Component.FilterID
        Case giCOMPONENT_FIELD
          lngExprID = .Component.SelectionFilterID
      End Select
      
      If lngExprID > 0 Then
        isOwnerOfComp = (LCase(Trim(gsUserName)) = LCase(Trim(GetExprField(lngExprID, "Username"))))
      Else
        isOwnerOfComp = True
      End If
    End If
  End With

TidyUpAndExit:

  Exit Function

ErrorTrap:
  isOwnerOfComp = False
  COAMsgBox "Error checking owner of component.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit
  
End Function

Public Function isCompHidden(objComp As clsExprComponent) As Boolean

'********************************************************************************
' isCompHidden - Checks if the current is hidden or has hidden components.      *
'********************************************************************************

  Dim blnHidden As Boolean

  On Error GoTo ErrorTrap
  
  blnHidden = False
  
  With objComp
    If (.ComponentType = giCOMPONENT_CALCULATION) _
        Or (.ComponentType = giCOMPONENT_FILTER) _
        Or (.ComponentType = giCOMPONENT_FIELD) Then
      
      Select Case .ComponentType
        Case giCOMPONENT_CALCULATION
          blnHidden = (HasHiddenComponents(.Component.CalculationID) _
                      Or (GetExprField(.Component.CalculationID, "Access") = ACCESS_HIDDEN))
        Case giCOMPONENT_FILTER
          blnHidden = (HasHiddenComponents(.Component.FilterID) _
                        Or (GetExprField(.Component.FilterID, "Access") = ACCESS_HIDDEN))
        Case giCOMPONENT_FIELD
          If (.Component.SelectionFilterID > 0) Then
            blnHidden = (HasHiddenComponents(.Component.SelectionFilterID) _
                        Or (GetExprField(.Component.SelectionFilterID, "Access") = ACCESS_HIDDEN))
          End If
      End Select
      
    End If
  End With

  isCompHidden = blnHidden

TidyUpAndExit:

  Exit Function

ErrorTrap:
  isCompHidden = True
  COAMsgBox "Error checking access of component.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit
  
End Function

Public Function isCompDeleted(objComp As clsExprComponent) As Boolean

'********************************************************************************
' isCompDeleted - Checks if the current component still exists in the db.       *
'********************************************************************************

  Dim blnDeleted As Boolean
  Dim rsTemp As ADODB.Recordset
  Dim lngCompID As Long
  Dim sSQL As String

  On Error GoTo ErrorTrap
  
  blnDeleted = False
  
  With objComp
    If (.ComponentType = giCOMPONENT_CALCULATION) _
        Or (.ComponentType = giCOMPONENT_FILTER) _
        Or (.ComponentType = giCOMPONENT_FIELD) Then
      
      Select Case .ComponentType
        Case giCOMPONENT_CALCULATION
          lngCompID = .Component.CalculationID
        Case giCOMPONENT_FILTER
          lngCompID = .Component.FilterID
        Case giCOMPONENT_FIELD
          lngCompID = .Component.SelectionFilterID
      End Select
      
      If lngCompID > 0 Then
        sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngCompID
        
        Set rsTemp = datGeneral.GetRecords(sSQL)
           
        blnDeleted = (rsTemp.BOF And rsTemp.EOF)
        
        Set rsTemp = Nothing
      End If
    End If
  End With

  isCompDeleted = blnDeleted
  
TidyUpAndExit:

  Exit Function

ErrorTrap:
  isCompDeleted = True
  COAMsgBox "Error checking if component has been deleted.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit
    
End Function

Public Function ValidComponent(objComp As clsExprComponent, _
                                bShowMessages As Boolean) As Integer
                                
'********************************************************************************
' ValidComponent - Checks whether the component hasd been deleted or made hidden*
'                  by another user.                                             *
'********************************************************************************
                                
  Dim bIsDeleted As Boolean
  Dim bIsHidden As Boolean
  Dim bIsExprOwner As Boolean
  Dim bIsCompOwner As Boolean

  On Error GoTo ErrorTrap
  
  ValidComponent = 0
  
  'Check if the selected component has been deleted.
  bIsDeleted = isCompDeleted(objComp)
  If bIsDeleted Then
    'the current expression no longer exists.
    If bShowMessages Then
      COAMsgBox "The selected component has been deleted by another user. " _
              , vbExclamation + vbOKOnly, App.Title
    End If
    ValidComponent = 3
    Exit Function
  End If
 
  'Check if the component has been made hidden since being in the expression
  'component screen.
  bIsHidden = isCompHidden(objComp)
  If bIsHidden Then
    'selected component is hidden.
    bIsExprOwner = isOwnerOfParent(objComp)
    bIsCompOwner = isOwnerOfComp(objComp)
    
    If bIsExprOwner Then
      'current user is the owner of the current expression.
      If bIsCompOwner Then
        'current user is the owner of the current component.
        If objComp.ParentExpression.Access <> ACCESS_HIDDEN Then
          'the current expression is not already hidden.
          If bShowMessages Then
            COAMsgBox "The selected component is hidden, " & _
                    "the expression will now be made hidden." _
                    , vbInformation + vbOKOnly, App.Title
          End If
          objComp.ParentExpression.Access = ACCESS_HIDDEN
          ValidComponent = 1
          Exit Function
        End If
      Else
        'current user is the owner of the expression but NOT the owner
        'of the selected (hidden) component.
        If bShowMessages Then
          COAMsgBox "The selected component is owned by another user and " & vbCrLf & _
                  "has been made hidden. " _
                  , vbExclamation + vbOKOnly, App.Title
        End If
        ValidComponent = 4
        Exit Function
      End If
    Else
      'current user is not the owner of the current expression.
      If bShowMessages Then
        COAMsgBox "The selected component is hidden and cannot be added to " & _
                "another user's expression." _
                , vbExclamation + vbOKOnly, App.Title
      End If
      ValidComponent = 2
      Exit Function
    End If
  End If

TidyUpAndExit:

  Exit Function

ErrorTrap:
  ValidComponent = -1
  COAMsgBox "Error checking validity of component.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit

End Function

Public Function ValidateExpr(objExpr As clsExprExpression, bShowMessages As Boolean) As Integer

'********************************************************************************
' ValidateExpr - Checks whether the expression contains components which have   *
'                been deleted or made hidden by another user.                   *
'********************************************************************************

  Dim objComp As clsExprComponent
  Dim iValidationCode As Integer
  Dim iTempCode As Integer
  Dim sMessage As String

  On Error GoTo ErrorTrap
  
  iTempCode = 0
  iValidationCode = 0
  
  'loop through each component in the expression retrieving a code that idicates
  'the state of the components.
  For Each objComp In objExpr.Components
    With objComp
      iTempCode = ValidComponent(objComp, False)
      'need to take the highest code from the expression's components.
      '
      'Codes:
      '
      '4 -->  Expression is owned by current user but it contains hidden
      '       components owned by another user.
      '3 -->  Expression is owned by current user has deleted components
      '       in the definition.
      '2 -->  Expression is NOT owned by current user but it contains hidden
      '       components, therefore should now be hidden to all but owner.
      '1 -->  Expression is owned by current user but it contains hidden
      '       components owned by the current user and is not already hidden,
      '       therefore the expression should made hidden.
      '0 -->  Expression has the correct access defined for the components
      '       within the definition, no message required!
      
      If iTempCode > iValidationCode Then
        iValidationCode = iTempCode
      End If
    End With
  Next objComp
  
  Set objComp = Nothing
  iTempCode = 0
  
  '****************************

  If bShowMessages Then
  
    'Decode iValidationCode using the above return codes.
    Select Case iValidationCode
    Case 0:
      'All ok no message required
    Case 1:
      sMessage = "The selected " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " contains hidden components. " & vbCrLf & _
                  "The " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " will now be made hidden."
    Case 2:
      sMessage = "The selected " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " contains hidden components and is " & _
                  "owned by '" & objExpr.Owner & "'. " & vbCrLf & _
                  "The " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " will now be made hidden."
    Case 3:
      sMessage = "The selected " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " contains deleted components. " & vbCrLf & _
                  "These components will " & _
                  "now be removed from the " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & "."
    Case 4:
      sMessage = "The selected " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & " contains hidden components which " & _
                  "are owned by another user. " & vbCrLf & "These components will " & _
                  "now be removed from the " & LCase(ExpressionTypeName(objExpr.ExpressionType)) & "."
    End Select
  
    If iValidationCode > 0 Then
      COAMsgBox sMessage, vbExclamation + vbOKOnly, App.Title
    End If
    
  End If
  
  ValidateExpr = iValidationCode

TidyUpAndExit:

  Exit Function

ErrorTrap:
  ValidateExpr = -1
  COAMsgBox "Error checking validity of expression.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit

End Function

Public Function RemoveUnowned_HDComps(objExpr As clsExprExpression) As Boolean

'********************************************************************************
' RemoveUnowned_HDComps - Removes the any invalid components.                   *
'********************************************************************************
  
  Dim objComp As clsExprComponent
  
  On Error GoTo ErrorTrap
  
  For Each objComp In objExpr.Components
    With objComp
      If (isCompHidden(objComp) And (Not isOwnerOfComp(objComp))) Or isCompDeleted(objComp) Then
        objComp.ComponentID = 0
        objExpr.DeleteComponent objComp
      End If
    End With
  Next objComp

  RemoveUnowned_HDComps = True
  
TidyUpAndExit:
  
  Exit Function

ErrorTrap:
  RemoveUnowned_HDComps = False
  COAMsgBox "Error removing hidden components.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit

End Function

Public Function ExpressionTypeName(piType As Integer) As String
  ' Return the description of the expression type.
  Select Case piType
    Case giEXPR_COLUMNCALCULATION
      ExpressionTypeName = "Column Calculation"
    
    Case giEXPR_GOTFOCUS ' NOT USED.
      ExpressionTypeName = "Field Entry Validation Clause"
    
    Case giEXPR_RECORDVALIDATION
      ExpressionTypeName = "Field Validation"
      
    Case giEXPR_DEFAULTVALUE ' NOT USED.
      ExpressionTypeName = "Default Value"
      
    Case giEXPR_STATICFILTER
      ExpressionTypeName = "Filter"
      
    Case giEXPR_PAGEBREAK ' NOT USED.
      ExpressionTypeName = "Page Break"
  
    Case giEXPR_ORDER ' NOT USED.
      ExpressionTypeName = "Order"
  
    Case giEXPR_RECORDDESCRIPTION
      ExpressionTypeName = "Record Description"
  
    Case giEXPR_VIEWFILTER
      ExpressionTypeName = "View Filter"
  
    Case giEXPR_RUNTIMECALCULATION
      ExpressionTypeName = "Runtime Calculation"
  
    Case giEXPR_RUNTIMEFILTER
      ExpressionTypeName = "Filter"
  
    Case giEXPR_UTILRUNTIMEFILTER
      ExpressionTypeName = "Import File Filter"
  
    Case giEXPR_MATCHJOINEXPRESSION, giEXPR_MATCHWHEREEXPRESSION
      ExpressionTypeName = "Match Relation Expression"
  
    Case giEXPR_MATCHSCOREEXPRESSION
      ExpressionTypeName = "Match Score Calculation"
  
    Case Else
      ExpressionTypeName = "Expression"
  End Select

End Function





Public Function UniqueColumnValue(sTableName As String, sColumnName As String) As Long
  On Error GoTo ErrorTrap

  Dim lngUniqueValue As Long
  Dim sSQL As String
  Dim rsUniqueValue As Recordset
  
  ' Create a record set with a unique value for the given table and column.
  sSQL = "SELECT MAX(" & sColumnName & ") + 1 AS newValue" & _
    " FROM " & sTableName
  Set rsUniqueValue = datGeneral.GetRecords(sSQL)
  With rsUniqueValue
    If IsNull(!NewValue) Then
      lngUniqueValue = 1
    Else
      lngUniqueValue = !NewValue
    End If
    
    .Close
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsUniqueValue = Nothing
  'Return the unique column value.
  UniqueColumnValue = lngUniqueValue
  Exit Function
  
ErrorTrap:
  lngUniqueValue = 1
  Resume TidyUpAndExit

End Function


Public Function ValidNameChar(ByVal piAsciiCode As Integer, ByVal piPosition As Integer) As Integer
  ' Validate the characters used to create table and column names.
  On Error GoTo ErrorTrap
  
  If piAsciiCode = Asc(" ") Then
    ' Substitute underscores for spaces.
    If piPosition <> 0 Then
      piAsciiCode = Asc("_")
    Else
      piAsciiCode = 0
    End If
  Else
    ' Allow only pure alpha-numerics and underscores.
    ' Do not allow numerics in the first chracter position.
'    If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or _
'      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9") And piPosition <> 0) Or _
'      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
'      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
'      piAsciiCode = 0
'    End If
'  End If
  
  ' RH 15/08/2000 - BUG...we should be able to start filter/calcs with a number char
    If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or _
      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9")) Or _
      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
      piAsciiCode = 0
    End If
  End If
  
  ValidNameChar = piAsciiCode
  Exit Function
  
ErrorTrap:
  ValidNameChar = 0
  Err = False
  
End Function







Public Function ValidateOperatorParameters(plngOperatorID As Long, piResultType As ExpressionValueTypes, _
  piParam1Type As Integer, piParam2Type As Integer) As Boolean
  ' Validate the given operator with the given parameters.
  ' Return the result type in the piResultType parameter.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Validate the parameter types for the given operator.
  Select Case plngOperatorID
    Case 1 ' PLUS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 2 ' MINUS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 3 ' TIMES BY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 4 ' DIVIDED BY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 5 ' AND
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC

    Case 6 ' OR
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC

    Case 7 ' IS EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 8 ' IS NOT EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 9 ' IS LESS THAN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 10 ' IS GREATER THAN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC
    
    Case 11 ' IS LESS THAN OR EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC
    
    Case 12 ' IS GREATER THAN OR EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 13 ' NOT
      fOK = (piParam1Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC
    
    Case 14 ' IS CONTAINED WITHIN
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_LOGIC

    Case 15 ' TO THE POWER OF
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 16 ' MODULAS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 17 ' CONCATENATED WITH
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER
    
    Case Else ' Unknown operator
      fOK = False
  End Select
  
TidyUpAndExit:
  If Not fOK Then
    piResultType = giEXPR_UNKNOWNTYPE
  End If
  
  ValidateOperatorParameters = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function ValidateFunctionParameters(plngFunctionID As Variant, piResultType As ExpressionValueTypes, _
  Optional piParam1Type As Integer, Optional piParam2Type As Integer, _
  Optional piParam3Type As Integer, Optional piParam4Type As Integer, _
  Optional piParam5Type As Integer, Optional piParam6Type As Integer) As Boolean
  ' Validate the given function with the given parameters.
  ' Return the result type in the piResultType parameter.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  
  fOK = True

  ' Get the parameter types.
  piParam1Type = IIf(IsMissing(piParam1Type), giEXPRVALUE_UNDEFINED, piParam1Type)
  piParam2Type = IIf(IsMissing(piParam2Type), giEXPRVALUE_UNDEFINED, piParam2Type)
  piParam3Type = IIf(IsMissing(piParam3Type), giEXPRVALUE_UNDEFINED, piParam3Type)
  piParam4Type = IIf(IsMissing(piParam4Type), giEXPRVALUE_UNDEFINED, piParam4Type)
  piParam5Type = IIf(IsMissing(piParam5Type), giEXPRVALUE_UNDEFINED, piParam5Type)
  piParam6Type = IIf(IsMissing(piParam6Type), giEXPRVALUE_UNDEFINED, piParam6Type)

  ' Validate the parameter types for the given function.
  Select Case plngFunctionID
    Case 1 ' SYSTEM DATE
      fOK = True
      piResultType = giEXPRVALUE_DATE

    Case 2 ' CONVERT TO UPPERCASE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 3 ' CONVERT NUMERIC TO STRING
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 4 ' IF, THEN, ELSE
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = piParam3Type)
      piResultType = piParam2Type

    Case 5 ' REMOVE LEADING AND TRAINING SPACES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 6 ' EXTRACT CHARACTERS FROM THE LEFT
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 7 ' LENGTH OF STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 8 ' CONVERT TO LOWERCASE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 9 ' MAXIMUM
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 10 ' MINIMUM
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 11 ' SEARCH FOR CHARACTER STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 12 ' CAPITALIZE INITIALS
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 13 'EXTRACT CHARACTERS FROM THE RIGHT
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 14 ' EXTRACT PART OF A STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC) And _
        (piParam3Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 15 ' SYSTEM TIME
      fOK = True
      piResultType = giEXPRVALUE_CHARACTER

    Case 16 ' IS FIELD EMPTY
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) Or _
        (piParam1Type = giEXPRVALUE_NUMERIC) Or _
        (piParam1Type = giEXPRVALUE_LOGIC) Or _
        (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 17 ' CURRENT USER
      fOK = True
      piResultType = giEXPRVALUE_CHARACTER

    Case 18 ' WHOLE YEARS UNTIL CURRENT DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 19 ' REMAINING MONTHS SINCE WHOLE YEARS
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 20 ' INITIALS FROM FORENAMES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 21 ' FIRST NAME FROM FORENAMES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 22 ' WEEKDAYS FROM START AND END DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 23 ' ADD MONTHS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 24 ' ADD YEARS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 25 ' CONVERT CHARACTER TO NUMERIC
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 26 ' WHOLE MONTHS BETWEEN TWO DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 27 ' PARENTHESESES
      fOK = True
      piResultType = piParam1Type

    Case 28 ' DAY OF THE WEEK
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 29 ' NUMBER OF WORKING DAYS PER WEEK
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 30 ' ABSENCE DURATION
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_CHARACTER) And _
        (piParam3Type = giEXPRVALUE_DATE) And _
        (piParam4Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 31 ' ROUND DOWN TO NEAREST WHOLE NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 32 ' YEAR OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 33 ' MONTH OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 34 ' DAY OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 35 ' NICE DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER

    Case 36 ' NICE TIME
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 37 ' ROUND DATE TO START OF NEAREST MONTH
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    Case 38 ' IS BETWEEN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE) And (piParam3Type = giEXPRVALUE_DATE)) Or _
       ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC) And (piParam3Type = giEXPRVALUE_NUMERIC))
      piResultType = giEXPRVALUE_LOGIC

    Case 39 ' SERVICE YEARS
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 40 ' SERVICE MONTHS
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 41 ' STATUTORY REDUNDANCY PAY
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_DATE) And _
        (piParam4Type = giEXPRVALUE_NUMERIC) And _
        (piParam5Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 42 ' GET FIELD FROM DATABASE RECORD
      fOK = (((piParam1Type = giEXPRVALUE_BYREF_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_DATE) And (piParam2Type = giEXPRVALUE_DATE))) And _
        ((piParam3Type = giEXPRVALUE_BYREF_CHARACTER) Or _
        (piParam3Type = giEXPRVALUE_BYREF_NUMERIC) Or _
        (piParam3Type = giEXPRVALUE_BYREF_LOGIC) Or _
        (piParam3Type = giEXPRVALUE_BYREF_DATE))
      piResultType = (piParam3Type - giEXPRVALUE_BYREF_OFFSET)

    Case 43 ' GET UNIQUE CODE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 44 ' ADD DAYS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 45 ' DAYS BETWEEN TWO DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC
            
    Case 46 ' WORKING DAYS BETWEEN TWO DATES (INC BHOLS)
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 47 ' ABSENCE BETWEEN TWO DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 48 ' ROUND UP TO NEAREST WHOLE NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 49 ' ROUND TO NEAREST NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 51 ' CONVERT CURRENCY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_CHARACTER) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 52 ' Field Last Changed Date
      fOK = (piParam1Type = giEXPRVALUE_BYREF_CHARACTER Or _
        piParam1Type = giEXPRVALUE_BYREF_NUMERIC Or _
        piParam1Type = giEXPRVALUE_BYREF_LOGIC Or _
        piParam1Type = giEXPRVALUE_BYREF_DATE)
      piResultType = giEXPRVALUE_DATE
    
    Case 53 ' Field changed between two dates
      fOK = ((piParam1Type = giEXPRVALUE_BYREF_CHARACTER Or _
        piParam1Type = giEXPRVALUE_BYREF_NUMERIC Or _
        piParam1Type = giEXPRVALUE_BYREF_LOGIC Or _
        piParam1Type = giEXPRVALUE_BYREF_DATE)) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC
    
    Case 54 'Whole months between two dates
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
            (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC
    
    ' JPD20021121 Fault 3177
    Case 55 ' First Day of Month
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 56 ' Last Day of Month
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 57 ' First Day of Year
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 58 ' Last Day of Year
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021129 Fault 4337
    Case 59 ' NAME OF MONTH
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER
    
    ' JPD20021129 Fault 4337
    Case 60 ' NAME OF DAY
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER
    
    ' JPD20021129 Fault 3606
    Case 61 ' IS FIELD POPULATED
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) Or _
        (piParam1Type = giEXPRVALUE_NUMERIC) Or _
        (piParam1Type = giEXPRVALUE_LOGIC) Or _
        (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 65 'IS POST SUBORDINATE OF
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 66 'IS POST SUBORDINATE OF USER
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 67 'IS PERSONNEL SUBORDINATE OF
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 68 'IS PERSONNEL SUBORDINATE OF USER
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 69 'HAS POST SUBORDINATE
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 70 'HAS POST SUBORDINATE USER
      fOK = True
      '  fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 71 'HAS PERSONNEL SUBORDINATE
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 72 'HAS PERSONNEL SUBORDINATE USER
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC
    
    Case 73 'BRADFORD FACTOR
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case Else ' Unknown function
      fOK = False
      
  End Select

TidyUpAndExit:
  If Not fOK Then
    piResultType = giEXPR_UNKNOWNTYPE
  End If

  ValidateFunctionParameters = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ExprDeleted(lngExprID) As Boolean

  Dim rsExprTemp As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
  
  Set rsExprTemp = datGeneral.GetRecords(sSQL)
  
  With rsExprTemp
    If .BOF And .EOF Then ExprDeleted = True
    .Close
  End With
  
  sSQL = vbNullString
  Set rsExprTemp = Nothing
  
End Function

Public Function CreateRuntimeUDFFunctions() As Boolean
  CreateRuntimeUDFFunctions = True
End Function

Public Function DeleteRuntimeUDFFunctions() As Boolean
  DeleteRuntimeUDFFunctions = True
End Function
