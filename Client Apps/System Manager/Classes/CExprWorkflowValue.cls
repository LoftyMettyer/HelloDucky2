VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprWorkflowValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum WorkflowElementProperties
  WORKFLOWELEMENTPROP_ITEMVALUE = 0
  WORKFLOWELEMENTPROP_COMPETIONCOUNT = 1
  WORKFLOWELEMENTPROP_FAILURECOUNT = 2
  WORKFLOWELEMENTPROP_TIMEOUTCOUNT = 3
  WORKFLOWELEMENTPROP_MESSAGE = 4
End Enum

' Component definition variables.
Private msElementIdentifier As String
Private msItemIdentifier As String
Private miElementProperty As WorkflowElementProperties

' Class handling variables.
Private mobjBaseComponent As CExprComponent



Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  ' Do nothing for table values.
  
End Sub




Public Function ElementPropertyDescription(piElementProperty As WorkflowElementProperties) As String
  Dim sDescription As String
  
  Select Case piElementProperty
    Case WORKFLOWELEMENTPROP_ITEMVALUE
      sDescription = "<Item Value>"
    Case WORKFLOWELEMENTPROP_COMPETIONCOUNT
      sDescription = "<Completion Count>"
    Case WORKFLOWELEMENTPROP_FAILURECOUNT
      sDescription = "<Failure Count>"
    Case WORKFLOWELEMENTPROP_TIMEOUTCOUNT
      sDescription = "<Timeout Count>"
    Case WORKFLOWELEMENTPROP_MESSAGE
      sDescription = "<Message>"
    Case Else
      sDescription = "<unknown>"
  End Select
  
  ElementPropertyDescription = sDescription
  
End Function

Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  ' Do nothing for workflow values.

End Sub





Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ContainsExpression = False
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function






Public Property Get BaseComponent() As CExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property



Public Property Set BaseComponent(ByVal pobjNewValue As CExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  ComponentContainsColumn = False
  
End Function

Public Property Get ComponentDescription() As String
  ' Return the component description.
  Dim sDesc As String
  Dim sElementDesc As String
  Dim sItemDesc As String
  Dim aWFPrecedingElements() As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  
  sDesc = ""
  sElementDesc = ""
  sItemDesc = ""
  
  If Len(Trim(msElementIdentifier)) > 0 Then
    aWFPrecedingElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements

    For iLoop = 2 To UBound(aWFPrecedingElements)
      If UCase(Trim(aWFPrecedingElements(iLoop).Identifier)) = UCase(Trim(msElementIdentifier)) Then
        sElementDesc = msElementIdentifier
        
        Select Case miElementProperty
          Case WORKFLOWELEMENTPROP_COMPETIONCOUNT, _
            WORKFLOWELEMENTPROP_FAILURECOUNT, _
            WORKFLOWELEMENTPROP_MESSAGE, _
            WORKFLOWELEMENTPROP_TIMEOUTCOUNT

            sItemDesc = ElementPropertyDescription(miElementProperty)

          Case Else
            Set wfTemp = aWFPrecedingElements(iLoop)
            asItems = wfTemp.Items
        
            For iLoop2 = 1 To UBound(asItems, 2)
              If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(msItemIdentifier)) Then
                sItemDesc = msItemIdentifier
          
                Exit For
              End If
            Next iLoop2
            Set wfTemp = Nothing
        End Select
        
        Exit For
      End If
    Next iLoop
  End If
  
  If Len(sElementDesc) = 0 Then
    sDesc = "<unknown>"
  Else
    If Len(sItemDesc) = 0 Then
      sDesc = sElementDesc & "." & "<unknown>"
    Else
      sDesc = sElementDesc & "." & sItemDesc
    End If
  End If
  
  ComponentDescription = sDesc

End Property


Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type.
  ComponentType = giCOMPONENT_WORKFLOWVALUE
  
End Property

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objValueCopy As New CExprWorkflowValue
  
  ' Copy the component's basic properties.
  With objValueCopy
    .WorkflowElement = msElementIdentifier
    .WorkflowItem = msItemIdentifier
    .WorkflowElementProperty = miElementProperty
  End With
  
  Set CopyComponent = objValueCopy
   
  ' Disassociate object variables.
  Set objValueCopy = Nothing
   
End Function

Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objValueClone As New CExprWorkflowValue
  
  fOK = True
  
  ' Copy the component's basic properties.
  With objValueClone
    .WorkflowElement = msElementIdentifier
    .WorkflowItem = msItemIdentifier
    .WorkflowElementProperty = miElementProperty
  End With
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objValueClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objValueClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function ConstructComponent() As Boolean
  ' Read the Value component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean

  fReadOk = False
  
  With recCompEdit
    ' Read the Workflow Value component record
    msElementIdentifier = IIf(IsNull(!WorkflowElement), "", !WorkflowElement)
    msItemIdentifier = IIf(IsNull(!WorkflowItem), "", !WorkflowItem)
    miElementProperty = IIf(IsNull(!WorkflowElementProperty), WORKFLOWELEMENTPROP_ITEMVALUE, !WorkflowElementProperty)
  End With
   
  fReadOk = True

TidyUpAndExit:
  ConstructComponent = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Public Function ReturnType_FromDatabase() As ExpressionValueTypes
  ' Return a string describing the component in a stored procedure.
  ' ie. Code that declares the component variable, and assigns the required value to it.
  Dim rsTemp As dao.Recordset
  Dim rsTemp2 As dao.Recordset
  Dim lngElementID As Long
  Dim iItemType As WorkflowWebFormItemTypes
  Dim lngColumnID As Long
  Dim iType As ExpressionValueTypes
  Dim sSQL As String
  
  lngElementID = 0
  lngColumnID = 0
  iItemType = giWFFORMITEM_UNKNOWN
  
  If (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION _
    Or mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER _
    Or mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) _
    And (mobjBaseComponent.ParentExpression.UtilityID > 0) Then
    
    ' Get the Element ID.
    sSQL = "SELECT TOP 1 tmpWorkflowElements.ID" & _
      " FROM tmpWorkflowElements" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
      "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'"
  
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsTemp
      If Not (.BOF And .EOF) Then
        lngElementID = !id
      End If
      
      .Close
    End With
    Set rsTemp = Nothing
    
    Select Case miElementProperty
      Case WORKFLOWELEMENTPROP_COMPETIONCOUNT, _
        WORKFLOWELEMENTPROP_FAILURECOUNT, _
        WORKFLOWELEMENTPROP_TIMEOUTCOUNT
        
        iType = giEXPRVALUE_NUMERIC

      Case WORKFLOWELEMENTPROP_MESSAGE
        iType = giEXPRVALUE_CHARACTER
      
      Case Else
        ' Get the Item type.
        sSQL = "SELECT TOP 1 tmpWorkflowElementItems.itemType," & _
          "   tmpWorkflowElementItems.lookupColumnID" & _
          " FROM tmpWorkflowElementItems" & _
          " WHERE tmpWorkflowElementItems.elementID = " & CStr(lngElementID) & _
          "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElementItems.identifier))) = '" & Replace(UCase(Trim(msItemIdentifier)), "'", "''") & "'"
      
        Set rsTemp2 = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        With rsTemp2
          If Not (.BOF And .EOF) Then
            iItemType = !ItemType
            lngColumnID = !LookupColumnID
          End If
          
          .Close
        End With
        Set rsTemp = Nothing
    
        ' Get the expression data type of the item
        Select Case iItemType
          Case giWFFORMITEM_BUTTON
            iType = giEXPRVALUE_LOGIC
    
          Case giWFFORMITEM_INPUTVALUE_CHAR
            iType = giEXPRVALUE_CHARACTER
    
          Case giWFFORMITEM_INPUTVALUE_NUMERIC
            iType = giEXPRVALUE_NUMERIC
    
          Case giWFFORMITEM_INPUTVALUE_LOGIC
            iType = giEXPRVALUE_LOGIC
    
          Case giWFFORMITEM_INPUTVALUE_DATE
            iType = giEXPRVALUE_DATE
    
          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
            iType = giEXPRVALUE_CHARACTER
    
          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            Select Case GetColumnDataType(lngColumnID)
              Case dtLONGVARCHAR
                iType = giEXPRVALUE_CHARACTER
              Case dtNUMERIC
                iType = giEXPRVALUE_NUMERIC
              Case dtINTEGER
                iType = giEXPRVALUE_NUMERIC
              Case dtTIMESTAMP
                iType = giEXPRVALUE_DATE
              Case dtVARCHAR
                iType = giEXPRVALUE_CHARACTER
            End Select
    
          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            iType = giEXPRVALUE_CHARACTER
        End Select
    End Select
  End If
  
  ReturnType_FromDatabase = iType
  
End Function

Public Property Get WorkflowElement() As String
  ' Return the WorkflowElement property.
  WorkflowElement = msElementIdentifier
  
End Property

Public Property Get WorkflowItem() As String
  ' Return the Workflow Item property.
  WorkflowItem = msItemIdentifier
  
End Property


Public Property Let WorkflowElement(ByVal psNewValue As String)
  ' Set the Workflow Element property.
  msElementIdentifier = psNewValue

End Property





Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the Workflow value's return type.
  Dim iReturnType As ExpressionValueTypes
  Dim asItems() As String
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFAllElements() As VB.Control
  Dim wfTemp As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sSQL As String
  Dim rsInfo As dao.Recordset
  
  iReturnType = giEXPRVALUE_UNDEFINED
  aWFPrecedingElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements
  aWFAllElements = mobjBaseComponent.ParentExpression.AllWorkflowElements

  If Len(msElementIdentifier) > 0 Then
    
    Select Case miElementProperty
      Case WORKFLOWELEMENTPROP_COMPETIONCOUNT, _
        WORKFLOWELEMENTPROP_FAILURECOUNT, _
        WORKFLOWELEMENTPROP_TIMEOUTCOUNT
        
        iReturnType = giEXPRVALUE_NUMERIC

      Case WORKFLOWELEMENTPROP_MESSAGE
        iReturnType = giEXPRVALUE_CHARACTER
      
      Case Else
        If Len(msItemIdentifier) > 0 Then
          If UBound(aWFAllElements) < 0 Then
            sSQL = "SELECT tmpWorkflowElementItems.itemType," & _
              "   tmpWorkflowElementItems.lookupColumnID" & _
              " FROM tmpWorkflowElementItems" & _
              " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
              " WHERE UCASE(LTRIM(RTRIM(tmpWorkflowElementItems.identifier))) = '" & UCase(Trim(msItemIdentifier)) & "'" & _
              "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & UCase(Trim(msElementIdentifier)) & "'" & _
              "   AND tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID)
      
            Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          
            If Not (rsInfo.BOF And rsInfo.EOF) Then
              Select Case rsInfo!ItemType
                Case giWFFORMITEM_BUTTON
                  iReturnType = giEXPRVALUE_LOGIC
        
                Case giWFFORMITEM_INPUTVALUE_CHAR
                  iReturnType = giEXPRVALUE_CHARACTER
        
                Case giWFFORMITEM_INPUTVALUE_NUMERIC
                  iReturnType = giEXPRVALUE_NUMERIC
        
                Case giWFFORMITEM_INPUTVALUE_LOGIC
                  iReturnType = giEXPRVALUE_LOGIC
        
                Case giWFFORMITEM_INPUTVALUE_DATE
                  iReturnType = giEXPRVALUE_DATE
        
                Case giWFFORMITEM_INPUTVALUE_DROPDOWN
                  iReturnType = giEXPRVALUE_CHARACTER
        
                Case giWFFORMITEM_INPUTVALUE_LOOKUP
                  Select Case GetColumnDataType(rsInfo!LookupColumnID)
                    Case dtLONGVARCHAR
                      iReturnType = giEXPRVALUE_CHARACTER
                    Case dtNUMERIC
                      iReturnType = giEXPRVALUE_NUMERIC
                    Case dtINTEGER
                      iReturnType = giEXPRVALUE_NUMERIC
                    Case dtTIMESTAMP
                      iReturnType = giEXPRVALUE_DATE
                    Case dtVARCHAR
                      iReturnType = giEXPRVALUE_CHARACTER
                  End Select
        
                Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                  iReturnType = giEXPRVALUE_CHARACTER
        
              End Select
            End If
          
            rsInfo.Close
            Set rsInfo = Nothing
          Else
            For iLoop = 1 To UBound(aWFPrecedingElements)
              If UCase(Trim(aWFPrecedingElements(iLoop).Identifier)) = UCase(Trim(msElementIdentifier)) Then
                Set wfTemp = aWFPrecedingElements(iLoop)
                asItems = wfTemp.Items
                
                For iLoop2 = 1 To UBound(asItems, 2)
                  If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(msItemIdentifier)) Then
                    Select Case asItems(2, iLoop2)
                      Case giWFFORMITEM_BUTTON
                        iReturnType = giEXPRVALUE_LOGIC
        
                      Case giWFFORMITEM_INPUTVALUE_CHAR
                        iReturnType = giEXPRVALUE_CHARACTER
        
                      Case giWFFORMITEM_INPUTVALUE_NUMERIC
                        iReturnType = giEXPRVALUE_NUMERIC
        
                      Case giWFFORMITEM_INPUTVALUE_LOGIC
                        iReturnType = giEXPRVALUE_LOGIC
        
                      Case giWFFORMITEM_INPUTVALUE_DATE
                        iReturnType = giEXPRVALUE_DATE
        
                      Case giWFFORMITEM_INPUTVALUE_DROPDOWN
                        iReturnType = giEXPRVALUE_CHARACTER
        
                      Case giWFFORMITEM_INPUTVALUE_LOOKUP
                        Select Case GetColumnDataType(CLng(asItems(49, iLoop2)))
                          Case dtLONGVARCHAR
                            iReturnType = giEXPRVALUE_CHARACTER
                          Case dtNUMERIC
                            iReturnType = giEXPRVALUE_NUMERIC
                          Case dtINTEGER
                            iReturnType = giEXPRVALUE_NUMERIC
                          Case dtTIMESTAMP
                            iReturnType = giEXPRVALUE_DATE
                          Case dtVARCHAR
                            iReturnType = giEXPRVALUE_CHARACTER
                        End Select
        
                      Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                        iReturnType = giEXPRVALUE_CHARACTER
        
                    End Select
        
                    Exit For
                  End If
                Next iLoop2
                
                If iReturnType <> giEXPRVALUE_UNDEFINED Then
                  Exit For
                End If
              End If
            Next iLoop
          End If
        End If
    End Select
  End If
  
  ReturnType = iReturnType

End Property

Public Property Let WorkflowItem(ByVal psNewValue As String)
  ' Set the Workflow Item property.
  msItemIdentifier = psNewValue

End Property






Private Sub Class_Initialize()
  ' Initialize properties.
  msElementIdentifier = ""
  msItemIdentifier = ""
  miElementProperty = WORKFLOWELEMENTPROP_ITEMVALUE
  
End Sub











Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  With recCompEdit
    .AddNew
    
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_WORKFLOWVALUE
    !WorkflowElement = msElementIdentifier
    !WorkflowItem = msItemIdentifier
    !WorkflowElementProperty = miElementProperty
    
    .Update
  End With

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function




Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean)
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated Value of the component.
   
End Function


Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  ' Return a string describing the component in a stored procedure.
  ' ie. Code that declares the component variable, and assigns the required value to it.
  Dim sVarName As String
  Dim sProcCode As String
  
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  
  Select Case miElementProperty
    Case WORKFLOWELEMENTPROP_COMPETIONCOUNT
      sProcCode = _
        "    DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
        "    SELECT " & sVarName & " = convert(float, isnull(ASRSysWorkflowInstanceSteps.completionCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID"

    Case WORKFLOWELEMENTPROP_FAILURECOUNT
      sProcCode = _
        "    DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
        "    SELECT " & sVarName & " = convert(float, isnull(ASRSysWorkflowInstanceSteps.failedCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID"

    Case WORKFLOWELEMENTPROP_TIMEOUTCOUNT
      sProcCode = _
        "    DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
        "    SELECT " & sVarName & " = convert(float, isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID"

    Case WORKFLOWELEMENTPROP_MESSAGE
      sProcCode = _
        "    DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine & _
        "    SELECT " & sVarName & " = isnull(ASRSysWorkflowInstanceSteps.message, '')" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID"
    
    Case Else
      Select Case ReturnType_FromDatabase
        Case giEXPRVALUE_CHARACTER
          sProcCode = _
            "    DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine & _
            "    SELECT " & sVarName & " = CASE" & vbNewLine & _
            "        WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN isnull(ASRSysWorkflowInstanceValues.tempValue, '')" & vbNewLine & _
            "        ELSE isnull(ASRSysWorkflowInstanceValues.value, '')" & vbNewLine & _
            "    END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID"
    
        Case giEXPRVALUE_NUMERIC
          sProcCode = _
            "    DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
            "    SET @sValue = null;" & vbNewLine & vbNewLine & _
            "    SELECT @sValue = CASE" & vbNewLine & _
            "        WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN isnull(ASRSysWorkflowInstanceValues.tempValue, 0)" & vbNewLine & _
            "        ELSE isnull(ASRSysWorkflowInstanceValues.value, 0)" & vbNewLine & _
            "    END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & vbNewLine & _
            "    SET " & sVarName & " = convert(float, @sValue)"
    
        Case giEXPRVALUE_LOGIC
          sProcCode = _
            "    DECLARE " & sVarName & " bit" & vbNewLine & vbNewLine & _
            "    SET @sValue = null;" & vbNewLine & vbNewLine & _
            "    SELECT @sValue = CASE" & vbNewLine & _
            "        WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN isnull(ASRSysWorkflowInstanceValues.tempValue, 0)" & vbNewLine & _
            "        ELSE isnull(ASRSysWorkflowInstanceValues.value, 0)" & vbNewLine & _
            "    END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & vbNewLine & _
            "    SET " & sVarName & " = convert(bit, @sValue)"
    
        Case giEXPRVALUE_DATE
          sProcCode = _
            "    DECLARE " & sVarName & " datetime" & vbNewLine & vbNewLine & _
            "    SET @sValue = null;" & vbNewLine & vbNewLine & _
            "    SELECT @sValue = CASE" & vbNewLine & _
            "        WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN ASRSysWorkflowInstanceValues.tempValue" & vbNewLine & _
            "        ELSE ASRSysWorkflowInstanceValues.value" & vbNewLine & _
            "    END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & _
            "        AND len(isnull(CASE WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN ASRSysWorkflowInstanceValues.tempValue ELSE ASRSysWorkflowInstanceValues.value END, '')) > 0" & vbNewLine & _
            "        AND UPPER(LTRIM(RTRIM(isnull(CASE" & vbNewLine & _
            "                WHEN ASRSysWorkflowInstanceValues.elementID = @piTempElement THEN ASRSysWorkflowInstanceValues.tempValue" & vbNewLine & _
            "                ELSE ASRSysWorkflowInstanceValues.value" & vbNewLine & _
            "            END, 'null')))) <> 'NULL'" & vbNewLine & vbNewLine & _
            "    SET " & sVarName & " = convert(datetime, @sValue)"
        Case Else
          sProcCode = ""
      End Select
  End Select
  
  StoredProcedureCode = sProcCode

End Function

Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sCode As String
  
  fOK = True
  sCode = ""
  
  Select Case miElementProperty
    Case WORKFLOWELEMENTPROP_COMPETIONCOUNT
      sCode = _
        "    (SELECT convert(float, isnull(ASRSysWorkflowInstanceSteps.completionCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID)"
    
    Case WORKFLOWELEMENTPROP_FAILURECOUNT
      sCode = _
        "    (SELECT convert(float, isnull(ASRSysWorkflowInstanceSteps.failedCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID)"

    Case WORKFLOWELEMENTPROP_TIMEOUTCOUNT
      sCode = _
        "    (SELECT convert(float, isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0))" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID)"

    Case WORKFLOWELEMENTPROP_MESSAGE
      sCode = _
        "    (SELECT isnull(ASRSysWorkflowInstanceSteps.message, '')" & vbNewLine & _
        "    FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
        "        AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID)"
    
    Case Else
      Select Case ReturnType_FromDatabase
        Case giEXPRVALUE_CHARACTER
          sCode = _
            "    (SELECT isnull(ASRSysWorkflowInstanceValues.value, '')" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID)"
    
        Case giEXPRVALUE_NUMERIC
          sCode = _
            "    (SELECT" & vbNewLine & _
            "        CASE" & vbNewLine & _
            "            WHEN isnumeric(ASRSysWorkflowInstanceValues.value) = 1 THEN convert(float, isnull(ASRSysWorkflowInstanceValues.value, 0))" & vbNewLine & _
            "            ELSE 0" & vbNewLine & _
            "        END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID)"
    
        Case giEXPRVALUE_LOGIC
          sCode = _
            "    (SELECT" & vbNewLine & _
            "        CASE" & vbNewLine & _
            "            WHEN isnumeric(ASRSysWorkflowInstanceValues.value) = 1 THEN convert(bit, isnull(ASRSysWorkflowInstanceValues.value, 0))" & vbNewLine & _
            "            ELSE 0" & vbNewLine & _
            "        END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID)"
    
        Case giEXPRVALUE_DATE
          sCode = _
            "    (SELECT " & vbNewLine & _
            "        CASE" & vbNewLine & _
            "            WHEN isdate(ASRSysWorkflowInstanceValues.value) = 1 THEN convert(datetime, isnull(ASRSysWorkflowInstanceValues.value, 0))" & vbNewLine & _
            "            ELSE 0" & vbNewLine & _
            "        END" & vbNewLine & _
            "    FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowElements.identifier = '" & Replace(msElementIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(msItemIdentifier, "'", "''") & "'" & vbNewLine & _
            "        AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & _
            "        AND len(isnull(ASRSysWorkflowInstanceValues.value, '')) > 0" & vbNewLine & _
            "        AND UPPER(LTRIM(RTRIM(isnull(ASRSysWorkflowInstanceValues.value, '')))) <> 'NULL')"
    
        Case Else
          sCode = ""
      End Select
    
  End Select
  
TidyUpAndExit:
  If fOK Then
    psRuntimeCode = sCode
  Else
    psRuntimeCode = ""
  End If
  
  RuntimeCode = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  glngBottom = CalculateBottomOfPage
  
  ' Position the printing.
  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  'Printer.Print ComponentDescription
  Printer.Print "Workflow Value : " & ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & ComponentDescription & vbNewLine
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function






Public Property Get WorkflowElementProperty() As WorkflowElementProperties
  WorkflowElementProperty = miElementProperty

End Property

Public Property Let WorkflowElementProperty(ByVal piNewValue As WorkflowElementProperties)
  miElementProperty = piNewValue
  
End Property

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function
