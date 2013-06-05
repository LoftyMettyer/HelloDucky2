VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{A48C54F8-25F4-4F50-9112-A9A3B0DBAD63}#1.0#0"; "COA_Label.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Object = "{98B2556E-F719-4726-9028-5F2EAB345800}#1.0#0"; "COASD_Checkbox.ocx"
Object = "{3EBC9263-7DE3-4E87-8721-81ACE59CD84E}#1.1#0"; "COASD_Combo.ocx"
Object = "{3CCEDCBE-4766-494F-84C9-95993D77BD56}#1.0#0"; "COASD_Command.ocx"
Object = "{FFAE31F9-C18D-4C20-AAF7-74C1356185D9}#1.0#0"; "COASD_Frame.ocx"
Object = "{5F165695-EDF2-40E1-BD8E-8D2E6325BDCF}#1.0#0"; "COASD_Image.ocx"
Object = "{CE18FF03-F3BF-4C4F-81DC-192ED1E1B91F}#1.0#0"; "COASD_OptionGroup.ocx"
Object = "{58F88252-94BB-43CE-9EF9-C971F73B93D4}#1.0#0"; "COASD_Selection.ocx"
Object = "{714061F3-25A6-4821-B196-7D15DCCDE00E}#1.0#0"; "COASD_SelectionBox.ocx"
Object = "{63212438-5384-4CC0-B836-A2C015CCBF9B}#1.0#0"; "COAWF_WebForm.ocx"
Object = "{BD3A90B9-91E4-40D5-A504-C6DFB4380BBC}#1.0#0"; "COASD_Grid.ocx"
Begin VB.Form frmWorkflowWFDesigner 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Workflow Web Form Designer"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5071
   Icon            =   "frmWorkflowWFDesigner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin COASDCommand.COASD_Command ASRDummyFileUpload 
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "File Upload"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      ForeColor       =   -2147483630
      WFItemType      =   0
   End
   Begin COAWFWebForm.COAWF_Webform wfDummyElement 
      Height          =   795
      Left            =   4200
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1402
      Caption         =   "Web Form"
   End
   Begin COASDOptionGroup.COASD_OptionGroup ASRDummyOptions 
      Height          =   630
      Index           =   0
      Left            =   2040
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2381
      _ExtentY        =   1111
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin COASDGrid.COASD_Grid ASRDummyGrid 
      Height          =   1035
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1826
   End
   Begin COASDSelection.COASD_Selection ASRSelectionMarkers 
      Height          =   795
      Index           =   0
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1402
   End
   Begin COASDSelectionBox.COASD_SelectionBox asrBoxMovementMarker 
      Height          =   510
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   900
      BorderColor     =   -2147483640
   End
   Begin COASDSelectionBox.COASD_SelectionBox asrboxMultiSelection 
      Height          =   570
      Left            =   4320
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BorderColor     =   -2147483640
      BorderStyle     =   3
   End
   Begin COALine.COA_Line ASRDummyLine 
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   4320
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   53
   End
   Begin COASDFrame.COASD_Frame asrDummyFrame 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1931
   End
   Begin COALabel.COA_Label asrDummyLabel 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "asrDummyLabel"
      FontSize        =   8.25
   End
   Begin COASDCheckbox.COASD_Checkbox asrDummyCheckBox 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
   End
   Begin COASDCombo.COASD_Combo asrDummyCombo 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
   End
   Begin COALabel.COA_Label asrDummyTextBox 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin COASDImage.COASD_Image asrDummyImage 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
   End
   Begin COASDCommand.COASD_Command btnWorkflow 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      ForeColor       =   -2147483643
   End
   Begin VB.Label lblBlankDesigner 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<Drag Web Form items onto the designer>"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   3705
   End
   Begin ActiveBarLibraryCtl.ActiveBar abWebForm 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmWorkflowWFDesigner.frx":000C
   End
End
Attribute VB_Name = "frmWorkflowWFDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constants.
Const giMAXTABS = 50
Const giSTANDARDMOVEMENT = 15
Const giMOVEMENTMARKERWIDTH = 20
Const gLngDFLTSCREENHEIGHT = 4900
Const gLngDFLTSCREENWIDTH = 7100
Const gLngAUTOFORMATLABELCOLUMN = 300
Const gLngAUTOFORMATYOFFSET = 100
Const gLngAUTOFORMATYSTART = 300

Private Const MIN_FORM_HEIGHT = 2700
Private Const MIN_FORM_WIDTH = 5520

' Properties.
Private mfAlignToGrid As Boolean
Private giGridX As Long
Private giGridY As Long
Private gLngScreenID As Long
Private mlngPictureLocation As Long
Private mlngPictureID As Long
Private msWFIdentifier As String
Private mfReadOnly As Boolean
Private mlngTimeoutFrequency As Long
Private miTimeoutPeriod As TimeoutPeriod
Private mfTimeoutExcludeWeekend As Boolean
Private mlngDescriptionExprID As Long
Private mfDescriptionHasWorkflowName As Boolean
Private mfDescriptionHasElementCaption As Boolean
Private masValidations() As String

Private miCompletionMessageType As MessageType
Private msCompletionMessage As String
Private miSavedForLaterMessageType As MessageType
Private msSavedForLaterMessage As String
Private miFollowOnFormsMessageType As MessageType
Private msFollowOnFormsMessage As String

' Globals.
Private gfLoading As Boolean
Private gfMultiSelecting As Boolean
Private gLngMultiSelectionXStart As Long
Private gLngMultiSelectionYStart As Long
Private gfStretchDown As Boolean
Private gfStretchUp As Boolean
Private gfStretchRight As Boolean
Private gfStretchLeft As Boolean
Private gfMoveSelection As Boolean
Private gLngOldX As Long
Private gLngOldY As Long
Private gfMouseDown As Boolean
Private gfExitToWorkflowDesigner As Boolean
Private gfActivating As Boolean
Private giLastActionFlag As UndoActionFlags
Private giUndo_ControlIndex As Integer
Private giUndo_ControlAutoLabelIndex As Integer
Private gsUndo_ControlType As String
Private gsUndo_ControlAutoLabelType As String
Private gavUndo_PastedControls() As Variant
Private gactlUndo_DeletedControls() As VB.Control
Private gactlClipboardControls() As VB.Control
Private gbAutoSendFrameToBack As Boolean

Private mlngLastX As Long
Private mlngLastY As Long
Private mlngXOffset As Long
Private mlngYOffset As Long

Private mlngLastFormWidth As Long
Private mlngLastFormheight As Long

Private mbKeyStretching As Boolean
Private mbKeyMoving As Boolean

Private mlngMouseX As Long
Private mlngMouseY As Long

Private mfrmCallingForm As Form
Private mwfElement As COAWF_Webform
Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfForcedChanged As Boolean
Private mfLoading As Boolean

Private maWFPrecedingElements() As VB.Control
Private maWFAllElements() As VB.Control

Private mavIdentifierLog() As Variant

Private mlngBaseTableID As Long
Private miInitiationType As WorkflowInitiationTypes
Private mlngWorkflowID As Long

Private maobjOriginalExpressions() As CExpression
Private mfExpressionsChanged As Boolean

Private Function ControlIsUsed(pctlControl As VB.Control, _
  Optional pavMessages As Variant) As Boolean
  ' Return True if the given controls is used and connot be deleted.
  
  Dim ctlControl As VB.Control
  Dim fControlIsUsed As Boolean

  fControlIsUsed = False
  
  '----------
  ' Check if any lookup items use the given control for filtering.
  '----------
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      With ctlControl
        If CLng(.WFItemType) = giWFFORMITEM_INPUTVALUE_LOOKUP Then
          If .LookupFilterValue = pctlControl.WFIdentifier Then

            ReDim Preserve pavMessages(3, UBound(pavMessages, 2) + 1)
            pavMessages(1, UBound(pavMessages, 2)) = GetWebFormItemTypeName(.WFItemType) & " (" & .WFIdentifier & ")"
            pavMessages(2, UBound(pavMessages, 2)) = "Lookup filter value"
            pavMessages(3, UBound(pavMessages, 2)) = .TabIndex
          
            fControlIsUsed = True
          End If
        End If
      End With
    End If
  Next ctlControl
  Set ctlControl = Nothing

  ControlIsUsed = fControlIsUsed
  
End Function

Public Function CurrentElementDefinition() As COAWF_Webform
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = SaveWebFormProperties(wfDummyElement)
  
TidyUpAndExit:
  If fOK Then
    Set CurrentElementDefinition = wfDummyElement
  End If

  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub RefreshExpressionNames()
  ' Refresh any labels that display their calculation name.
  Dim ctlControl As VB.Control
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim lngExprID As Long
  Dim sCaption As String
  Dim fCalcDefault As Boolean
  
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      With ctlControl
        iWFItemType = CLng(.WFItemType)

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTION) _
          And WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTIONTYPE) _
          And WebFormItemHasProperty(iWFItemType, WFITEMPROP_CALCULATION) Then
          
          If .CaptionType = giWFDATAVALUE_CALC Then
            lngExprID = .CalculationID
            sCaption = GetExpressionName(lngExprID)

            If Len(Trim(sCaption)) = 0 Then
              sCaption = "<Calculated>"
              .CalculationID = 0
            Else
              sCaption = "<" & sCaption & ">"
            End If

            .Caption = sCaption
          End If
        End If
        
        fCalcDefault = WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUE_EXPRID)
        If fCalcDefault _
          And WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
          
          fCalcDefault = (.DefaultValueType = giWFDATAVALUE_CALC)
        End If
        If fCalcDefault Then
          If (iWFItemType <> giWFFORMITEM_INPUTVALUE_LOGIC) _
            And (iWFItemType <> giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Then
            
            lngExprID = .CalculationID
            sCaption = GetExpressionName(lngExprID)

            If Len(Trim(sCaption)) = 0 Then
              sCaption = "<Calculated>"
              .CalculationID = 0
            Else
              sCaption = "<" & sCaption & ">"
            End If

            .Caption = sCaption
          End If
        End If
      End With
    End If
  Next ctlControl
  ' Disassociate object variables.
  Set ctlControl = Nothing

End Sub


Public Sub ShowPropertiesForm(Optional pvarShowWebFormProps As Variant)
  Dim frmTimeout As frmWorkflowTimeout
  Dim iSelectedControlCount As Integer
  Dim iItemType As WorkflowWebFormItemTypes
  Dim ctlControl As VB.Control
  Dim ctlSelectedControl As VB.Control
  Dim ctlMarker As COASD_Selection
  Dim fShowWebFormProperties As Boolean
  
  fShowWebFormProperties = False
  If Not IsMissing(pvarShowWebFormProps) Then
    fShowWebFormProperties = CBool(pvarShowWebFormProps)
  End If
  
  ' Determine the type of the selected item.
  If fShowWebFormProperties Then
    iSelectedControlCount = 0
  Else
    iSelectedControlCount = SelectedControlsCount
  End If
  
  If iSelectedControlCount > 1 Then
    Exit Sub
  ElseIf iSelectedControlCount = 1 Then
    For Each ctlControl In Me.Controls
      If IsWebFormControl(ctlControl) Then
        If ctlControl.Selected Then
          Set ctlSelectedControl = ctlControl
          Exit For
        End If
      End If
    Next ctlControl
    ' Disassociate object variables.
    Set ctlControl = Nothing
  End If

  ' Show the properties form
  Set frmTimeout = New frmWorkflowTimeout
  With frmTimeout
    .Initialise ctlSelectedControl, Me
    .Changed = False
    .Show vbModal
  End With

  UnLoad frmTimeout
  Set frmTimeout = Nothing

  If Not IsChanged Then
    IsChanged = WorkflowExpressionsChanged
  End If
  
  RefreshExpressionNames
  
  ' Refresh the selection markers
  For Each ctlMarker In ASRSelectionMarkers
    With ctlMarker
      If .Visible Then
        .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
        .RefreshSelectionMarkers True
      End If
    End With
  Next ctlMarker
  Set ctlMarker = Nothing
  
  ' Refresh the properties screen.
  Set frmWorkflowWFItemProps.CurrentWebForm = Me
  frmWorkflowWFItemProps.RefreshProperties

End Sub

Public Property Set CallingForm(pfrmForm As Form)
  Set mfrmCallingForm = pfrmForm
  mfReadOnly = pfrmForm.ReadOnly
  miInitiationType = pfrmForm.InitiationType
  mlngWorkflowID = pfrmForm.WorkflowID

  If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
    mlngBaseTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  Else
    mlngBaseTableID = pfrmForm.BaseTable
  End If
  
End Property
Public Property Get BaseTable() As Long
  BaseTable = mlngBaseTableID
  
End Property

Public Property Let BaseTable(ByVal plngNewValue As Long)
  mlngBaseTableID = plngNewValue

End Property
Public Property Get InitiationType() As WorkflowInitiationTypes
  InitiationType = miInitiationType
  
End Property

Public Property Let InitiationType(ByVal piNewValue As WorkflowInitiationTypes)
  miInitiationType = piNewValue

End Property



Public Property Get CallingForm() As Form
  Set CallingForm = mfrmCallingForm
End Property

Public Property Set Element(pwfElement As COAWF_Webform)

  Set mwfElement = pwfElement
  
  ReDim maWFPrecedingElements(1)
  Set maWFPrecedingElements(1) = mwfElement
  mfrmCallingForm.PrecedingElements mwfElement, maWFPrecedingElements
  
  ReDim maWFAllElements(0)
  mfrmCallingForm.AllElements maWFAllElements
  
  mfExpressionsChanged = False
  RememberOriginalExpressions

  LoadWebForm
  
End Property

Private Function WorkflowExpressionsChanged() As Boolean
  ' Return TRUE if any of the Workflow expressions have been modified or created.
  Dim fExpressionsChanged As Boolean
  Dim sSQL As String
  Dim rsTemp As dao.Recordset
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim objExpression As CExpression
  Dim alngCurrentExpressions() As Long

  If Not mfExpressionsChanged Then
    ReDim alngCurrentExpressions(0)

    ' Check each 'live' expression (ie. the one currently to be saved)
    sSQL = "SELECT tmpExpressions.exprID, tmpExpressions.lastSave" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
      "   AND tmpExpressions.deleted = FALSE" & _
      "   AND tmpExpressions.parentComponentID = 0"

    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    Do While Not rsTemp.EOF
      ' Check if the lastSave value of the 'live' expression is more recent than the one read when this definition was loaded.
      fFound = False

      ReDim Preserve alngCurrentExpressions(UBound(alngCurrentExpressions) + 1)
      alngCurrentExpressions(UBound(alngCurrentExpressions)) = rsTemp!ExprID

      For iLoop = 1 To UBound(maobjOriginalExpressions)
        Set objExpression = maobjOriginalExpressions(iLoop)

        If rsTemp!ExprID = objExpression.ExpressionID Then
          fExpressionsChanged = (rsTemp!LastSave <> objExpression.LastSave)
          fFound = True
          Set objExpression = Nothing
          Exit For
        End If

        Set objExpression = Nothing
      Next iLoop

      If Not fFound Then
        fExpressionsChanged = True
      End If

      If fExpressionsChanged Then
        Exit Do
      End If

      rsTemp.MoveNext
    Loop

    rsTemp.Close
    Set rsTemp = Nothing

    If Not fExpressionsChanged Then
      ' All 'live' expressions are unchanged from the original values, but have any original ones been deleted?
      For iLoop = 1 To UBound(maobjOriginalExpressions)
        Set objExpression = maobjOriginalExpressions(iLoop)
        fFound = False

        For iLoop2 = 1 To UBound(alngCurrentExpressions)
          If objExpression.ExpressionID = alngCurrentExpressions(iLoop2) Then
            fFound = True
            Exit For
          End If
        Next iLoop2

        Set objExpression = Nothing

        If Not fFound Then
          ' Original expression no longer 'live', so must have been deleted.
          fExpressionsChanged = True
          Exit For
        End If
      Next iLoop
    End If

    mfExpressionsChanged = fExpressionsChanged
  End If

  WorkflowExpressionsChanged = mfExpressionsChanged
  
End Function


Private Sub RememberOriginalExpressions()
  ' Read all of the Workflows original expressions.
  Dim sSQL As String
  Dim rsTemp As dao.Recordset
  Dim objExpression As CExpression

  ReDim maobjOriginalExpressions(0)

  sSQL = "SELECT tmpExpressions.exprID" & _
    " FROM tmpExpressions" & _
    " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
    "   AND tmpExpressions.deleted = FALSE" & _
    "   AND tmpExpressions.parentComponentID = 0"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  Do While Not rsTemp.EOF
    Set objExpression = New CExpression
    objExpression.ExpressionID = rsTemp!ExprID
    objExpression.ConstructExpression

    ReDim Preserve maobjOriginalExpressions(UBound(maobjOriginalExpressions) + 1)
    Set maobjOriginalExpressions(UBound(maobjOriginalExpressions)) = objExpression
    Set objExpression = Nothing

    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
    
End Sub


Private Sub RestoreOriginalExpressions()
  ' Restore the Workflows original expressions.
  Dim sSQL As String
  Dim objExpression As CExpression
  Dim iLoop As Integer
  Dim fChanged As Boolean
  Dim sOriginalExprIDs As String
  Dim rsTemp As dao.Recordset

  sOriginalExprIDs = "0"

  For iLoop = 1 To UBound(maobjOriginalExpressions)
    Set objExpression = maobjOriginalExpressions(iLoop)

    sSQL = "UPDATE tmpExpressions" & _
      " SET deleted = FALSE" & _
      " WHERE exprID = " & CStr(objExpression.ExpressionID)
    daoDb.Execute sSQL, dbFailOnError

    sOriginalExprIDs = sOriginalExprIDs & "," & CStr(objExpression.ExpressionID)
    fChanged = objExpression.IsChanged
    objExpression.EvaluatedReturnType = objExpression.ReturnType

    objExpression.WriteExpression_Transaction

    ' Changed flag will be set to true when restoring the original definition
    ' regardless of the original value. Manually restore the original value now if required.
    If fChanged Then
      sSQL = "UPDATE tmpExpressions" & _
        " SET lastSave = " & Replace(Format(objExpression.LastSave, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & _
        " WHERE exprID = " & CStr(objExpression.ExpressionID)
      daoDb.Execute sSQL, dbFailOnError
    Else
      sSQL = "UPDATE tmpExpressions" & _
        " SET changed = FALSE," & _
        "   lastSave = " & Replace(Format(objExpression.LastSave, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & _
        " WHERE exprID = " & CStr(objExpression.ExpressionID)
      daoDb.Execute sSQL, dbFailOnError
    End If

    Set objExpression = Nothing
  Next iLoop

  ' Mark any 'live' expressions that were newly created as deleted.
  sSQL = "SELECT tmpExpressions.exprID" & _
    " FROM tmpExpressions" & _
    " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
    "   AND exprID NOT IN (" & sOriginalExprIDs & ")" & _
    "   AND tmpExpressions.deleted = FALSE" & _
    "   AND tmpExpressions.parentComponentID = 0"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  Do While Not rsTemp.EOF
    Set objExpression = New CExpression
    objExpression.ExpressionID = rsTemp!ExprID
    objExpression.DeleteExpression
    Set objExpression = Nothing

    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
    
End Sub


Public Function IsUniqueIdentifier(psIdentifier As String, _
  Optional pvControlToIgnore As VB.Control) As Boolean
  
  Dim ctlControl As VB.Control
  Dim fIgnoreControl As Boolean
  
  For Each ctlControl In Me.Controls
    fIgnoreControl = False
    
    If Not IsMissing(pvControlToIgnore) Then
      If ctlControl Is pvControlToIgnore Then
        fIgnoreControl = True
      End If
    End If
    
    If (Not fIgnoreControl) And IsWebFormControl(ctlControl) Then
      With ctlControl
        If WebFormItemHasProperty(.WFItemType, WFITEMPROP_WFIDENTIFIER) Then
          If UCase(Trim(.WFIdentifier)) = UCase(Trim(psIdentifier)) Then
            IsUniqueIdentifier = False
            Exit Function
          End If
        End If
      End With
    End If
  Next ctlControl
  
  IsUniqueIdentifier = True
  
End Function

Public Function IsUniqueElementIdentifier(psIdentifier As String) As Boolean
  
  IsUniqueElementIdentifier = mfrmCallingForm.UniqueIdentifier(psIdentifier, mwfElement.ControlIndex)
  
End Function


Public Property Get Element() As COAWF_Webform
  Set Element = mwfElement

End Property

Public Sub PrecedingElements(paWFPrecedingElements As Variant)
  paWFPrecedingElements = maWFPrecedingElements

End Sub

Public Sub AllElements(paWFAllElements As Variant)
  paWFAllElements = maWFAllElements
End Sub

Private Function UniqueCaption(pctlItem As VB.Control) As String
  On Error GoTo ErrorTrap
  
  Dim sCaption As String
  Dim sTemp As String
  Dim ctlTemp As VB.Control
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim iItemType As WorkflowWebFormItemTypes
  
  iItemType = pctlItem.WFItemType
  
  Select Case iItemType
    Case giWFFORMITEM_BUTTON
      sCaption = "Button"
      
    Case giWFFORMITEM_IMAGE
      sCaption = "Image"
    
    Case giWFFORMITEM_INPUTVALUE_CHAR
      sCaption = "InputChar"
    
    Case giWFFORMITEM_INPUTVALUE_NUMERIC
      sCaption = "InputNumeric"

    Case giWFFORMITEM_INPUTVALUE_LOGIC
      sCaption = "InputLogic"
    
    Case giWFFORMITEM_INPUTVALUE_DROPDOWN
      sCaption = "InputDropdown"
    
    Case giWFFORMITEM_INPUTVALUE_LOOKUP
      sCaption = "InputLookup"
      
    Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      sCaption = "InputOptionGroup"

    Case giWFFORMITEM_INPUTVALUE_DATE
      sCaption = "InputDate"
    
    Case giWFFORMITEM_INPUTVALUE_GRID
      sCaption = "RecordSelector"

    Case giWFFORMITEM_LABEL
      sCaption = "Label"

    Case giWFFORMITEM_FRAME
      sCaption = "Frame"

    Case giWFFORMITEM_LINE
      sCaption = "Line"

    Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
      sCaption = "InputFileUpload"

    Case Else
      sCaption = ""
  End Select

  If Len(sCaption) > 0 Then
    fFound = True
    iLoop = 0
      
    Do While fFound
      fFound = False
      iLoop = iLoop + 1
      sTemp = sCaption & CStr(iLoop)
      
      For Each ctlTemp In Me.Controls
        If IsWebFormControl(ctlTemp) Then
          If (ctlTemp.WFItemType = iItemType) _
            And (Not ctlTemp Is pctlItem) Then
                
            If WebFormItemHasProperty(iItemType, WFITEMPROP_CAPTION) Then
              If ctlTemp.Caption = sTemp Then
                fFound = True
                Exit For
              End If
            End If
            
            If WebFormItemHasProperty(iItemType, WFITEMPROP_WFIDENTIFIER) Then
              If ctlTemp.WFIdentifier = sTemp Then
                fFound = True
                Exit For
              End If
            End If
          End If
        End If
      Next ctlTemp
      Set ctlTemp = Nothing
      
      If Not fFound Then
        sCaption = sTemp
      End If
    Loop
  End If
  
TidyUpAndExit:
  UniqueCaption = sCaption
  Exit Function

ErrorTrap:
  sCaption = ""
  GoTo TidyUpAndExit
  
End Function

Public Function UpdateIdentifiers(pfElement As Boolean, _
  psOldIdentifier As String, _
  psNewIdentifier As String, _
  plngOldParameter As Long, _
  plngNewParameter As Long)

  ' plngOldParameter/plngNewParameter refer to tableIDs if we're dealing with recordSelectors.
  ' plngOldParameter/plngNewParameter refer to column data types if we're dealing with lookups.
  
  Dim iLoop As Integer
  Dim fElementIdentifierChanged As Boolean
  Dim fElementTableChanged As Boolean
  Dim fFound As Boolean
  Dim frmUsage As frmUsage
  Dim asMessages() As String
  Dim sSQL As String
  Dim rsTemp As dao.Recordset
  Dim objExpr As CExpression
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim sExprType As String
  Dim sExprName As String
  Dim sComponentType As String
  Dim alngValidTables() As Long
  Dim asItems() As String
  Dim ctlControl As VB.Control
  Dim iSQLDataType As SQLDataType
  Dim fItemOK As Boolean
  
  ' Clear the array of validation messages
  ' Column 0 = The message
  ReDim asMessages(0)
  
  '----------
  ' Update any lookup items that used the old identifier for filtering.
  '----------
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      With ctlControl
        If CLng(.WFItemType) = giWFFORMITEM_INPUTVALUE_LOOKUP Then
          If .LookupFilterValue = psOldIdentifier Then
          
            ' Check if the datatype has changed, invalidating it's use in a lookup filter.
            If (plngOldParameter <> plngNewParameter) Then
              iSQLDataType = GetColumnDataType(.LookupFilterColumn)
              fItemOK = True
      
              Select Case iSQLDataType
                Case dtVARCHAR, dtlongvarchar
                  fItemOK = (plngNewParameter = dtVARCHAR) _
                    Or (plngNewParameter = dtlongvarchar)
                Case dtTIMESTAMP
                    fItemOK = (plngNewParameter = dtTIMESTAMP)
                Case dtinteger, dtNUMERIC
                  fItemOK = (plngNewParameter = dtinteger) _
                    Or (plngNewParameter = dtNUMERIC)
                Case Else
                  fItemOK = False
              End Select
              
              If Not fItemOK Then
                ReDim Preserve asMessages(UBound(asMessages) + 1)
                asMessages(UBound(asMessages)) = _
                  GetWebFormItemTypeName(.WFItemType) & " (" & .WFIdentifier & ") : " & _
                  "Invalid lookup filter value selected"
              End If
            End If
            
            .LookupFilterValue = psNewIdentifier
          End If
        End If
      End With
    End If
  Next ctlControl
  Set ctlControl = Nothing
        
  
  '----------
  ' Update any expressions that used the old identifier or the old table.
  '----------
  If (UCase(Trim(psOldIdentifier)) <> UCase(Trim(psNewIdentifier))) _
    Or (plngOldParameter <> plngNewParameter) Then
    
    fElementIdentifierChanged = pfElement _
      And (UCase(Trim(psOldIdentifier)) <> UCase(Trim(psNewIdentifier)))
    fElementTableChanged = (plngOldParameter <> plngNewParameter)
    
    ' Update the identifiers in any of this Workflow's expressions
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.type, tmpComponents.workflowItem, tmpComponents.workflowRecordTableID" & _
      " FROM tmpComponents" & _
      " INNER JOIN tmpExpressions ON tmpComponents.exprID = tmpExpressions.exprID" & _
      " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
      "   AND (tmpExpressions.type = " & CStr(giEXPR_WORKFLOWCALCULATION) & _
      "     OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
      "     OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER) & ")"
      
    If pfElement Then
      sSQL = sSQL & _
      "   AND ucase(ltrim(rtrim(tmpComponents.workflowElement))) = '" & Replace(UCase(Trim(psOldIdentifier)), "'", "''") & "'"
    Else
      sSQL = sSQL & _
      "   AND ucase(ltrim(rtrim(tmpComponents.workflowElement))) = '" & Replace(UCase(Trim(WFIdentifier)), "'", "''") & "'" & _
      "   AND ucase(ltrim(rtrim(tmpComponents.workflowItem))) = '" & Replace(UCase(Trim(psOldIdentifier)), "'", "''") & "'"
    End If
    
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    Do While Not (rsTemp.BOF Or rsTemp.EOF)
      sComponentType = ComponentTypeName(rsTemp!Type)

      If pfElement Then
        Set objComp = New CExprComponent
        objComp.ComponentID = rsTemp!ComponentID
        lngExprID = objComp.RootExpressionID
        Set objComp = Nothing

        sSQL = "UPDATE tmpExpressions" & _
          " SET tmpExpressions.changed = TRUE" & _
          " WHERE tmpExpressions.exprID = " & CStr(lngExprID)
        daoDb.Execute sSQL, dbFailOnError

        sSQL = "UPDATE tmpComponents" & _
          " SET tmpComponents.workflowElement = '" & Replace(psNewIdentifier, "'", "''") & "'" & _
          " WHERE tmpComponents.componentID = " & CStr(rsTemp!ComponentID)
        daoDb.Execute sSQL, dbFailOnError
      Else
        Set objComp = New CExprComponent
        objComp.ComponentID = rsTemp!ComponentID
        lngExprID = objComp.RootExpressionID
        Set objComp = Nothing

        sSQL = "UPDATE tmpExpressions" & _
          " SET tmpExpressions.changed = TRUE" & _
          " WHERE tmpExpressions.exprID = " & CStr(lngExprID)
        daoDb.Execute sSQL, dbFailOnError

        sSQL = "UPDATE tmpComponents" & _
          " SET tmpComponents.WorkflowItem = '" & Replace(psNewIdentifier, "'", "''") & "'" & _
          " WHERE tmpComponents.componentID = " & CStr(rsTemp!ComponentID)
        daoDb.Execute sSQL, dbFailOnError

        ' Check if the recordSelector table is still valid.
        If (plngOldParameter <> plngNewParameter) Then
          ReDim alngValidTables(0)
          TableAscendants plngNewParameter, alngValidTables
          
          fFound = False

          For iLoop = 1 To UBound(alngValidTables)
            If (alngValidTables(iLoop) = rsTemp!WorkflowRecordTableID) Then
              fFound = True
              Exit For
            End If
          Next iLoop

          If Not fFound Then
            ' Get the expression name and type description.
            Set objExpr = New CExpression
            objExpr.ExpressionID = lngExprID

            If objExpr.ReadExpressionDetails Then
              sExprName = objExpr.Name
              sExprType = objExpr.ExpressionTypeName

              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = _
                sExprType & " (" & sExprName & ") : " & _
                "Invalid " & sComponentType & IIf(rsTemp!Type = giCOMPONENT_WORKFLOWVALUE, " value", " record selector") & " selected"
            End If
          End If
        End If
      End If
      
      rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
  If UBound(asMessages) > 0 Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList

    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList asMessages(iLoop)
    Next iLoop

    Screen.MousePointer = vbNormal

    frmUsage.Width = (3 * Screen.Width / 4)

    frmUsage.ShowMessage "Workflow '" & Trim(mfrmCallingForm.WorkflowName) & "'", "The following Expressions/Web Form items made reference this web form" & IIf(pfElement, "", " item") & ", and will need reviewing:", _
      UsageCheckObject.Workflow, _
      USAGEBUTTONS_PRINT + USAGEBUTTONS_OK, "validation"

    UnLoad frmUsage
    Set frmUsage = Nothing
  End If
  
End Function

Private Sub ValidateIdentifiers(pctlControl As VB.Control, _
  pasMessages As Variant)
  
  On Error GoTo ErrorTrap
  
  Dim fValidElementIdentifier As Boolean
  Dim fValidItemIdentifier As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim lngColumnID As Long
  Dim lngTableID As Long
  Dim asItems() As String
  Dim iGoodItems As Integer
  Dim sMsg As String
  Dim sSubMessage1 As String
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  
  iWFItemType = CInt(pctlControl.WFItemType)
  
  fValidElementIdentifier = True
  fValidItemIdentifier = True
    
  Select Case iWFItemType
    '------------------------------------------------------------
    ' Database Value - only need to validate identifiers if the DBValue is for an 'identified' record.
    '------------------------------------------------------------
    Case giWFFORMITEM_DBVALUE, _
      giWFFORMITEM_DBFILE
      
      If (pctlControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD) Then
  
        fValidElementIdentifier = (Len(Trim(pctlControl.WFWorkflowForm)) > 0)
        lngColumnID = pctlControl.ColumnID
        lngTableID = GetTableIDFromColumnID(lngColumnID)
        sSubMessage1 = " (" & GetColumnName(lngColumnID) & ")"
  
        If fValidElementIdentifier Then
          fValidElementIdentifier = False
          
          For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
            Set wfTemp = maWFPrecedingElements(iLoop)
        
            If UCase(Trim(wfTemp.Identifier)) = UCase(Trim(pctlControl.WFWorkflowForm)) Then
              
              If wfTemp.ElementType = elem_WebForm Then
                fValidItemIdentifier = False
                
                iGoodItems = 0
                
                asItems = wfTemp.Items
    
                For iLoop2 = 1 To UBound(asItems, 2)
                  If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                    
                    ReDim alngValidTables(0)
                    TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                    
                    If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pctlControl.WFWorkflowValue)) Then
                      fFound = False
                      For lngLoop = 1 To UBound(alngValidTables)
                        If lngTableID = alngValidTables(lngLoop) Then
                          fFound = True
                          Exit For
                        End If
                      Next lngLoop
                      If fFound Then
                        fValidItemIdentifier = True
                      End If
                    End If
                  
                    For lngLoop = 1 To UBound(alngValidTables)
                      If lngTableID = alngValidTables(lngLoop) Then
                        iGoodItems = iGoodItems + 1
                        Exit For
                      End If
                    Next lngLoop
                  End If
                Next iLoop2
              
                ' If there are no recSels that match the required table then the element is invalid too.
                fValidElementIdentifier = (iGoodItems > 0)
              
              ElseIf wfTemp.ElementType = elem_StoredData Then
                ReDim alngValidTables(0)
                TableAscendants wfTemp.DataTableID, alngValidTables
                
                'JPD 20061227
                'If wfTemp.DataAction = DATAACTION_DELETE Then
                '  ' Cannot do anything with a Deleted record, but can use its ascendants.
                '  ' Remove the table itself from the array of valid tables.
                '  alngValidTables(1) = 0
                'End If
                
                fFound = False
                For lngLoop = 1 To UBound(alngValidTables)
                  If lngTableID = alngValidTables(lngLoop) Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop
                fValidElementIdentifier = fFound
              End If
                
              Exit For
            End If
          
            Set wfTemp = Nothing
          Next iLoop
        End If
        
        If Not fValidElementIdentifier Then
          ReDim Preserve pasMessages(UBound(pasMessages) + 1)
          pasMessages(UBound(pasMessages)) = "Database Value" & sSubMessage1 & " - Invalid element identifier"
        ElseIf Not fValidItemIdentifier Then
          ReDim Preserve pasMessages(UBound(pasMessages) + 1)
          pasMessages(UBound(pasMessages)) = "Database Value" & sSubMessage1 & " - Invalid record selector"
        End If
      End If
  
    '------------------------------------------------------------
    ' Record Selector - only need to validate identifiers if the records are related to an 'identified' record.
    '------------------------------------------------------------
    Case giWFFORMITEM_INPUTVALUE_GRID
      If (pctlControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD) Then
        fValidElementIdentifier = (Len(Trim(pctlControl.WFWorkflowForm)) > 0)
        lngTableID = pctlControl.TableID
        
        If fValidElementIdentifier Then
          fValidElementIdentifier = False
          
          For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
            Set wfTemp = maWFPrecedingElements(iLoop)
  
            If UCase(Trim(wfTemp.Identifier)) = UCase(Trim(pctlControl.WFWorkflowForm)) Then
              If wfTemp.ElementType = elem_WebForm Then
                fValidElementIdentifier = True
  
                fValidItemIdentifier = (Len(Trim(pctlControl.WFWorkflowValue)) > 0)
                If fValidItemIdentifier Then
                  fValidItemIdentifier = False
  
                  asItems = wfTemp.Items
  
                  For iLoop2 = 1 To UBound(asItems, 2)
                    If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) _
                      And UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pctlControl.WFWorkflowValue)) Then
  
                      ReDim alngValidTables(0)
                      TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                      
                      fFound = False
                      For lngLoop = 1 To UBound(alngValidTables)
                        If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                          fFound = True
                          Exit For
                        End If
                      Next lngLoop
                      fValidItemIdentifier = fFound
                      
                      Exit For
                    End If
                  Next iLoop2
                End If
              ElseIf wfTemp.ElementType = elem_StoredData Then
                ReDim alngValidTables(0)
                TableAscendants wfTemp.DataTableID, alngValidTables
                
                'JPD 20061227
                If wfTemp.DataAction = DATAACTION_DELETE Then
                  ' Cannot do anything with a Deleted record, but can use its ascendants.
                  ' Remove the table itself from the array of valid tables.
                  alngValidTables(1) = 0
                End If
                
                fFound = False
                For lngLoop = 1 To UBound(alngValidTables)
                  If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop
                fValidElementIdentifier = fFound
              End If
                
              Exit For
            End If
            
            Set wfTemp = Nothing
          Next iLoop
        End If
  
        If Not fValidElementIdentifier Then
          ReDim Preserve pasMessages(UBound(pasMessages) + 1)
          pasMessages(UBound(pasMessages)) = "Record Selector Input (" & pctlControl.WFIdentifier & ") - Invalid element identifier"
        ElseIf Not fValidItemIdentifier Then
          ReDim Preserve pasMessages(UBound(pasMessages) + 1)
          pasMessages(UBound(pasMessages)) = "Record Selector Input (" & pctlControl.WFIdentifier & ") - Invalid record selector"
        End If
      End If
    
    '------------------------------------------------------------
    ' Workflow Value
    '------------------------------------------------------------
    Case giWFFORMITEM_WFVALUE, _
      giWFFORMITEM_WFFILE
      
      fValidElementIdentifier = (Len(Trim(pctlControl.WFWorkflowForm)) > 0)
      sSubMessage1 = " (" & pctlControl.WFWorkflowForm & ")"

      If fValidElementIdentifier Then
        fValidElementIdentifier = False
        
        For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
          Set wfTemp = maWFPrecedingElements(iLoop)
      
          If UCase(Trim(wfTemp.Identifier)) = UCase(Trim(pctlControl.WFWorkflowForm)) Then
            If wfTemp.ElementType = elem_WebForm Then
              fValidElementIdentifier = True
              fValidItemIdentifier = (Len(Trim(pctlControl.WFWorkflowValue)) > 0)
                  
              If fValidItemIdentifier Then
                fValidItemIdentifier = False

                asItems = wfTemp.Items

                For iLoop2 = 1 To UBound(asItems, 2)
                  If ((asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_CHAR) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DATE) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
                    And UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pctlControl.WFWorkflowValue)) Then

                    fValidItemIdentifier = True
                    Exit For
                  End If
                Next iLoop2
              End If
            End If

            Exit For
          End If
          
          Set wfTemp = Nothing
        Next iLoop
      End If
      
      If Not fValidElementIdentifier Then
        ReDim Preserve pasMessages(UBound(pasMessages) + 1)
        pasMessages(UBound(pasMessages)) = "Workflow Value" & sSubMessage1 & " - Invalid web form identifier"
      ElseIf Not fValidItemIdentifier Then
        ReDim Preserve pasMessages(UBound(pasMessages) + 1)
        pasMessages(UBound(pasMessages)) = "Workflow Value" & sSubMessage1 & " - Invalid value selector"
      End If

  End Select

  'JPD 20070723 Fault 12405
  'If Not fValidElementIdentifier Then
  '  pctlControl.WFWorkflowForm = ""
  '  pctlControl.WFWorkflowValue = ""
  '
  '  mfForcedChanged = True
  'ElseIf Not fValidItemIdentifier Then
  '  pctlControl.WFWorkflowValue = ""
  '
  '  mfForcedChanged = True
  'End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Function ValidateWebForm() As Boolean
  ' Save the definition of each instance of each type of screen control to the database.
  On Error GoTo ErrorTrap
  
  Dim ctlControl As VB.Control
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim iButtonCount As Integer
  Dim asIdentifiers() As String
  Dim asDuplicateIdentifiers() As String
  Dim asMessages() As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim fFound As Boolean
  Dim frmUsage As frmUsage
  Dim fContinue As Boolean
  Dim asItems() As String
  Dim fDoCheck As Boolean
  
  iButtonCount = 0
  ReDim asIdentifiers(0)
  ReDim asDuplicateIdentifiers(0)
  ReDim asMessages(0)
  
  If (Len(Trim(msWFIdentifier)) = 0) Then
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "No identifier"
  Else
    If Not IsUniqueElementIdentifier(msWFIdentifier) Then
      ReDim Preserve asMessages(UBound(asMessages) + 1)
      asMessages(UBound(asMessages)) = "Non-unique identifier"
    End If
  End If
        
  ' Validate each form item.
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then

      With ctlControl
        iWFItemType = CLng(.WFItemType)

        ' Identifier must be unique within the form
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_WFIDENTIFIER) Then
          ReDim Preserve asIdentifiers(UBound(asIdentifiers) + 1)
          asIdentifiers(UBound(asIdentifiers)) = .WFIdentifier
        End If

        Select Case iWFItemType
          Case giWFFORMITEM_BUTTON
            ' Count the number of Submit buttons on the webform (must be at least one).
            If (.Behaviour = WORKFLOWBUTTONACTION_SUBMIT) _
              Or (.Behaviour = WORKFLOWBUTTONACTION_CANCEL) Then
                        
              iButtonCount = iButtonCount + 1
            End If

          Case giWFFORMITEM_LABEL
            ' Ensure a calculation is selected if required.
            If .CaptionType = giWFDATAVALUE_CALC _
              And .CalculationID = 0 Then
            
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = "Label - <Calculated> - No calculation selected."
            End If
            
          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            ' Check OptionGroups have item values.
            If .NoOptions Then
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no control values."
            End If
            
          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
            ' Check Dropdowns have item values.
            If Len(Trim(.ControlValueList)) = 0 Then
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no control values."
            End If
            
          Case giWFFORMITEM_INPUTVALUE_GRID
            ' Check valid table, record, etc. selected
            If .TableID = 0 Then
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no table defined."
            End If
            
          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            ' Check valid table, column selected
            If .LookupTableID = 0 Then
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no lookup table defined."
            Else
              If .LookupColumnID = 0 Then
                ReDim Preserve asMessages(UBound(asMessages) + 1)
                asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no lookup column defined."
              End If
            End If
        
        End Select
      
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
          fDoCheck = True

          If WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
            fDoCheck = (.DefaultValueType = giWFDATAVALUE_CALC)
          End If
          
          If fDoCheck _
            And .CalculationID = 0 Then

            ReDim Preserve asMessages(UBound(asMessages) + 1)
            asMessages(UBound(asMessages)) = "Input value - " & .WFIdentifier & " has no default value calculation selected."
          End If
        End If
        
      End With
    
      ' Validate element and recSel identifiers
      ValidateIdentifiers ctlControl, asMessages
    End If
  Next ctlControl

  ' There cannot be items with the same identifier.
  For iLoop = 1 To UBound(asIdentifiers) - 1
    For iLoop2 = (iLoop + 1) To UBound(asIdentifiers)
      If UCase(Trim(asIdentifiers(iLoop))) = UCase(Trim(asIdentifiers(iLoop2))) Then
        ' Duplicate found. Have already noticed it?
        fFound = False
        For iLoop3 = 1 To UBound(asDuplicateIdentifiers)
          If UCase(Trim(asIdentifiers(iLoop))) = UCase(Trim(asDuplicateIdentifiers(iLoop3))) Then
            fFound = True
            Exit For
          End If
        Next iLoop3
        
        If Not fFound Then
          ReDim Preserve asDuplicateIdentifiers(UBound(asDuplicateIdentifiers) + 1)
          asDuplicateIdentifiers(UBound(asDuplicateIdentifiers)) = asIdentifiers(iLoop)
        End If
      End If
    Next iLoop2
  Next iLoop
  
  ' There must be at least 1 button.
  If iButtonCount = 0 Then
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "The web form must have at least 1 Submit button."
    'JPD 20060719 Fault 11334
    '  ElseIf iButtonCount > 2 Then
    '    ReDim Preserve asMessages(UBound(asMessages) + 1)
    '    asMessages(UBound(asMessages)) = "The web form can have at most 2 buttons."
  End If
  
  ' There can't be duplicate identifiers.
  For iLoop = 1 To UBound(asDuplicateIdentifiers)
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "There is more than 1 item with the identifier '" & asDuplicateIdentifiers(iLoop) & "'."
  Next iLoop
  
  ' Display the validity failures to the user.
  fContinue = (UBound(asMessages) = 0)
  
  If Not fContinue Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList
      
    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList (asMessages(iLoop))
    Next iLoop
    
    Screen.MousePointer = vbNormal
    frmUsage.ShowMessage "Workflow", "The Web Form definition is invalid for the reasons listed below." & _
      vbCrLf & "Do you wish to continue?", UsageCheckObject.Workflow, _
      USAGEBUTTONS_YES + USAGEBUTTONS_NO + USAGEBUTTONS_PRINT, "validation"
    
    fContinue = (frmUsage.Choice = vbYes)
    
    UnLoad frmUsage
    Set frmUsage = Nothing
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  ValidateWebForm = fContinue
  Exit Function

ErrorTrap:
  fContinue = True
  Resume TidyUpAndExit

End Function

Public Property Let Validations(ByVal pavNewValue As Variant)
  masValidations = pavNewValue

End Property

Public Property Get Validations() As Variant
  Validations = masValidations

End Property

Private Sub abWebForm_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)
  EditMenu pTool.Name
End Sub

Private Sub abWebForm_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True
End Sub

Private Sub asrDummyCheckBox_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub asrDummyCombo_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub ASRDummyFileUpload_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub ASRDummyFileUpload_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop ASRDummyFileUpload(Index), Source, X, Y

End Sub


Private Sub ASRDummyFileUpload_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift

End Sub

Private Sub ASRDummyFileUpload_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown ASRDummyFileUpload(Index), Button, Shift, X, Y

End Sub

Private Sub ASRDummyFileUpload_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove ASRDummyFileUpload(Index), Button, X, Y

End Sub

Private Sub ASRDummyFileUpload_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp ASRDummyFileUpload(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyFrame_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub asrDummyFrame_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRDummyGrid_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub ASRDummyGrid_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop ASRDummyGrid(Index), Source, X, Y
End Sub

Private Sub ASRDummyGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown ASRDummyGrid(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyGrid_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove ASRDummyGrid(Index), Button, X, Y
End Sub

Private Sub ASRDummyGrid_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp ASRDummyGrid(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyCheckBox_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyCheckBox(Index), Source, X, Y
End Sub

Private Sub asrDummyCheckBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyCheckBox(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyCheckBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyCheckBox(Index), Button, X, Y
End Sub

Private Sub asrDummyCheckBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyCheckBox(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyCombo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyCombo(Index), Source, X, Y
End Sub

Private Sub asrDummyCombo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyCombo(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyCombo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyCombo(Index), Button, X, Y
End Sub

Private Sub asrDummyCombo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyCombo(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyFrame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyFrame(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyFrame(Index), Button, X, Y
End Sub

Private Sub asrDummyFrame_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyFrame(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyImage_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub asrDummyImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyImage(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyImage(Index), Button, X, Y
End Sub

Private Sub asrDummyImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyImage(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyLabel_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub asrDummyLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyLabel(Index), Source, X, Y
End Sub

Private Sub asrDummyLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyLabel(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyLabel(Index), Button, X, Y
End Sub

Private Sub asrDummyLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyLabel(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyLine_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub ASRDummyOptions_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub ASRDummyOptions_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop ASRDummyOptions(Index), Source, X, Y
End Sub

Private Sub ASRDummyOptions_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRDummyOptions_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown ASRDummyOptions(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove ASRDummyOptions(Index), Button, X, Y
End Sub

Private Sub ASRDummyOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp ASRDummyOptions(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyTextBox_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub btnWorkflow_DblClick(Index As Integer)
  ShowPropertiesForm

End Sub

Private Sub btnWorkflow_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop btnWorkflow(Index), Source, X, Y
End Sub

Private Sub btnWorkflow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown btnWorkflow(Index), Button, Shift, X, Y
End Sub

Private Sub btnWorkflow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove btnWorkflow(Index), Button, X, Y
End Sub

Private Sub btnWorkflow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp btnWorkflow(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyTextBox_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyTextBox(Index), Source, X, Y
End Sub

Private Sub asrDummyTextBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown asrDummyTextBox(Index), Button, Shift, X, Y
End Sub

Private Sub asrDummyTextBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove asrDummyTextBox(Index), Button, X, Y
End Sub

Private Sub asrDummyTextBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp asrDummyTextBox(Index), Button, Shift, X, Y
End Sub

Private Sub ASRSelectionMarkers_Stretch(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  Dim lngHeight As Long
  Dim lngWidth As Long
  Dim lngTop As Long
  Dim lngLeft As Long
  Dim bCanStretch As Boolean
  Dim iGridSize As Integer
    
  'UI.LockWindow Me.hWnd
  On Error GoTo CannotStretch
  
  iGridSize = 2
  
  If Not mfReadOnly Then
    For iCount = 1 To ASRSelectionMarkers.Count - 1
      With ASRSelectionMarkers(iCount)
                    
        If .Visible Then
        
          ' Default sizes for the stretch
          lngTop = .AttachedObject.Top
          lngHeight = .AttachedObject.Height
          lngLeft = .AttachedObject.Left
          lngWidth = .AttachedObject.Width
          bCanStretch = False
                
          Select Case Direction
        
            ' Stretch North West
            Case "TopLeft"
              bCanStretch = (Not .HasLockedHeight) Or (Not .HasLockedWidth)
              
              lngTop = IIf(Not .HasLockedHeight And (.Original_Height - Y > .AttachedObject.MinimumHeight), .Original_Top + Y, lngTop)
              lngLeft = IIf(Not .HasLockedWidth And (.Original_Width - X > .AttachedObject.MinimumWidth), .Original_Left + X, lngLeft)
              lngWidth = IIf(Not .HasLockedWidth And (.Original_Width - X > .AttachedObject.MinimumWidth), .Original_Width - X, lngWidth)
              lngHeight = IIf(Not .HasLockedHeight And (.Original_Height - Y > .AttachedObject.MinimumHeight), .Original_Height - Y, lngHeight)
                
            ' Stretch North
            Case "TopCentre"
              bCanStretch = (.Original_Height - Y > .AttachedObject.MinimumHeight) And (Not .HasLockedHeight)
              
              lngTop = .Original_Top + Y
              lngHeight = .Original_Height - Y
  
            ' Stretch North East
            Case "TopRight"
              bCanStretch = (Not .HasLockedHeight) Or (Not .HasLockedWidth)

              lngTop = IIf(Not .HasLockedHeight And (.Original_Height - Y > .AttachedObject.MinimumHeight), .Original_Top + Y, lngTop)
              lngWidth = IIf(Not .HasLockedWidth And (.Original_Width + X > .AttachedObject.MinimumWidth), .Original_Width + X, lngWidth)
              lngHeight = IIf(Not .HasLockedHeight And (.Original_Height - Y > .AttachedObject.MinimumHeight), .Original_Height - Y, lngHeight)
              
            Case "CentreLeft"
              bCanStretch = (.Original_Width - X > .AttachedObject.MinimumWidth And Not .HasLockedWidth)

              lngLeft = .Original_Left + X
              lngWidth = .Original_Width - X
            
            Case "CentreRight"
              bCanStretch = (.Original_Width + X > .AttachedObject.MinimumWidth) And (Not .HasLockedWidth)
              
              lngWidth = .Original_Width + X

            Case "BottomLeft"
              bCanStretch = IIf(IsWithin(lngWidth, .AttachedObject.Width, iGridSize) And IsWithin(lngHeight, .AttachedObject.Height, iGridSize), False, True)

              lngLeft = IIf(Not .HasLockedWidth And (.Original_Width - X > .AttachedObject.MinimumWidth), .Original_Left + X, lngLeft)
              lngWidth = IIf(Not .HasLockedWidth And (.Original_Width - X > .AttachedObject.MinimumWidth), .Original_Width - X, lngWidth)
              lngHeight = IIf(Not .HasLockedHeight And (.Original_Height + Y > .AttachedObject.MinimumHeight), .Original_Height + Y, lngHeight)
              
            Case "BottomCentre"
              bCanStretch = IIf(IsWithin(lngHeight, .AttachedObject.Height, iGridSize), True, Not .HasLockedHeight)
              
              lngHeight = .Original_Height + Y
  
            Case "BottomRight"
              bCanStretch = IIf(IsWithin(lngWidth, .AttachedObject.Width, iGridSize) And IsWithin(lngHeight, .AttachedObject.Height, iGridSize), False, True)
        
              lngWidth = .Original_Width + X
              lngHeight = .Original_Height + Y

          End Select
                  
          ' Only move the control if it is stretchable
          If bCanStretch Then
            .AttachedObject.Move lngLeft, lngTop, lngWidth, lngHeight
          End If
        
          ' If the controls height behaviour is set to full change it to fixed
          If WebFormItemHasProperty(.AttachedObject.WFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
            If (lngHeight <> .Original_Height) Then
              .AttachedObject.HeightBehaviour = behaveFixed
            End If
          End If

          If WebFormItemHasProperty(.AttachedObject.WFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
            If (lngWidth <> .Original_Width) Then
              .AttachedObject.WidthBehaviour = behaveFixed
            End If
          End If
        End If
          
      End With
    Next iCount
  End If
  
  'UI.UnlockWindow
  Exit Sub
  
CannotStretch:
  Exit Sub

End Sub

Private Sub ASRSelectionMarkers_StretchEnd(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  
  If Not mfReadOnly Then
    For iCount = 1 To ASRSelectionMarkers.Count - 1
      With ASRSelectionMarkers(iCount)
        If .Visible Then
          .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
          .RefreshSelectionMarkers True
        End If
      End With
    Next iCount
   
    Set frmWorkflowWFItemProps.CurrentWebForm = Me
    frmWorkflowWFItemProps.RefreshProperties
    
    ' Flag screen as having changed
    IsChanged = True
  End If
  
End Sub

Private Sub ASRSelectionMarkers_StretchStart(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  
  If Not mfReadOnly Then
    For iCount = 1 To ASRSelectionMarkers.Count - 1
      With ASRSelectionMarkers(iCount)
        If .Visible Then
          .SaveOriginalSizes
          .ShowSelectionMarkers False
        End If
      End With
    Next iCount
  
    ' Store original x,y coordinates
    mlngXOffset = X
    mlngYOffset = Y
  End If
  
End Sub

Private Sub Form_Activate()

  ' Ensure the screen designer form is at the front of the display.
  On Error GoTo ErrorTrap
  
  Me.ZOrder 0
  
  ' Refresh the properties screen.
  Set frmWorkflowWFItemProps.CurrentWebForm = Me
  frmWorkflowWFItemProps.RefreshProperties

  ' Refresh the menu/toolbar display.
  frmSysMgr.RefreshMenu

  RefreshBlankDesignLabel
  
  gfActivating = True

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error activating Workflow Web Form Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub

Private Function DropControl(pVarPageContainer As Variant, pCtlSource As Control, pSngX As Single, pSngY As Single) As Boolean
  
  ' Drop the given control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As WorkflowWebFormItemTypes
  Dim lngColumnID As Long
  Dim sCaption As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim objMisc As New Misc
  Dim ctlControl As VB.Control
  Dim tmpValue As String
  Dim tmpID As String
  Dim sAutoLabel As String
  Dim fTableOK As Boolean
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim alngValidTables() As Long
  Dim fDBValueDefaulted As Boolean
  Dim lngTableID As Long
  Dim iColumnDataType As DataTypes
  Dim fIsFile As Boolean
  Dim sElementIdentifier As String
  Dim sItemIdentifier As String
  
  ' Deselect all existing controls.
  fOK = DeselectAllControls
  
  If fOK Then
  
    ' Check that a column or standard control is being dropped
    If (pCtlSource Is frmWorkflowWFToolbox.trvColumns) Or _
      (pCtlSource Is frmWorkflowWFToolbox.trvStandardControls) Or _
      (pCtlSource Is frmWorkflowWFToolbox.trvWorkflowValue) Then
        
      ' If we are dropping a column control ...
      If pCtlSource Is frmWorkflowWFToolbox.trvColumns Then
        
        'Find the definition for the column being dropped
        With frmWorkflowWFToolbox.trvColumns.SelectedItem
          lngColumnID = val(Mid(.key, 2))
        End With
          
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", lngColumnID
          
          If Not .NoMatch Then
            lngTableID = recColEdit.Fields("TableId")
            iColumnDataType = .Fields("dataType")
            
            ' Add the required control type.
            If (iColumnDataType = dtLONGVARBINARY) _
              Or (iColumnDataType = dtVARBINARY) Then
            
              iControlType = giWFFORMITEM_DBFILE
            Else
              iControlType = giWFFORMITEM_DBVALUE
            End If
            
            Set ctlControl = AddControl(iControlType)
            fOK = Not (ctlControl Is Nothing)
        
            'Check that a new control was added successfully
            If fOK Then
  
              ' Set the last action flag and enable the Undo menu option.
              If Me.abWebForm.Tools("ID_AutoLabel").Checked = True Then
                giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL
              Else
                giLastActionFlag = giACTION_DROPCONTROL
              End If
                
              giUndo_ControlIndex = ctlControl.Index
              gsUndo_ControlType = ctlControl.Name
            
              Set ctlControl.Container = pVarPageContainer
              ctlControl.Left = AlignX(CLng(pSngX))
              ctlControl.Top = AlignY(CLng(pSngY))
              ctlControl.ColumnID = .Fields("columnID")
              
              ' Give the control a tooltip.
              sColumnName = .Fields("columnName")
              With recTabEdit
                .Index = "idxTableID"
                .Seek "=", lngTableID
                    
                If Not .NoMatch Then
                  sTableName = .Fields("tableName")
                  ctlControl.ToolTipText = "<" & sTableName & "." & sColumnName & ">"
                End If
              End With
              
              ' Initialise the new control's font and forecolour.
              If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
                Set objFont = New StdFont
                objFont.Name = Me.Font.Name
                objFont.Size = Me.Font.Size
                objFont.Bold = Me.Font.Bold
                objFont.Italic = Me.Font.Italic
                
                If (iColumnDataType = dtLONGVARBINARY) _
                  Or (iColumnDataType = dtVARBINARY) Then
                  
                  objFont.Underline = True
                End If
                
                Set ctlControl.Font = objFont
                Set objFont = Nothing
              End If
              
              If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
                ctlControl.ForeColor = Me.ForeColor
              End If
              
              If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
                ctlControl.BackColor = Me.BackColor
              End If
              
              If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION) Then
                ctlControl.Caption = "<" & objMisc.StrReplace(.Fields("columnName"), "_", " ", False) & ">" & vbNullString
              End If
              
              If WebFormControl_HasText(iControlType) Then
                ctlControl.Caption = ctlControl.ToolTipText
              End If
              
              If (lngTableID = mlngBaseTableID) And (miInitiationType <> WORKFLOWINITIATIONTYPE_EXTERNAL) Then
                ctlControl.WFDatabaseRecord = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED, _
                  giWFRECSEL_TRIGGEREDRECORD, _
                  giWFRECSEL_INITIATOR)
              Else
                'JPD 20070329 Fault 12040
                fDBValueDefaulted = False
                
                If (miInitiationType <> WORKFLOWINITIATIONTYPE_EXTERNAL) Then
                  ReDim alngValidTables(0)
                  TableAscendants mlngBaseTableID, alngValidTables
                  
                  For lngLoop = 1 To UBound(alngValidTables)
                    If (lngTableID = alngValidTables(lngLoop)) Then
                      ctlControl.WFDatabaseRecord = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED, _
                        giWFRECSEL_TRIGGEREDRECORD, _
                        giWFRECSEL_INITIATOR)
                      fDBValueDefaulted = True
                      
                      Exit For
                    End If
                  Next lngLoop
                End If
                
                If Not fDBValueDefaulted Then
                  ctlControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD

                  If UBound(maWFPrecedingElements) > 1 Then
                    fTableOK = False
                  
                    For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
                      Set wfTemp = maWFPrecedingElements(iLoop)
                  
                      If wfTemp.ElementType = elem_WebForm Then
                        ' Add  an item to the combo for each grid in the preceding web form.
                        asItems = wfTemp.Items
                        For iLoop2 = 1 To UBound(asItems, 2)
                          If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                            ReDim alngValidTables(0)
                            TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                  
                            For lngLoop = 1 To UBound(alngValidTables)
                              If (lngTableID = alngValidTables(lngLoop)) Then
                                fDBValueDefaulted = True
                                ctlControl.WFWorkflowForm = wfTemp.Identifier
                                ctlControl.WFWorkflowValue = asItems(9, iLoop2)
                                Exit For
                              End If
                            Next lngLoop
                  
                            If fDBValueDefaulted Then
                              Exit For
                            End If
                          End If
                        Next iLoop2
                      ElseIf wfTemp.ElementType = elem_StoredData Then
                        ReDim alngValidTables(0)
                        TableAscendants wfTemp.DataTableID, alngValidTables
                  
                        'JPD 20061227 DBValues can now be from DELETE StoredData elements
                        For lngLoop = 1 To UBound(alngValidTables)
                          If (lngTableID = alngValidTables(lngLoop)) Then
                            fDBValueDefaulted = True
                            ctlControl.WFWorkflowForm = wfTemp.Identifier
                            
                            Exit For
                          End If
                        Next lngLoop
                      End If
                  
                      Set wfTemp = Nothing
                      
                      If fDBValueDefaulted Then
                        Exit For
                      End If
                    Next iLoop
                  End If
                End If
              End If

              ' Default the control's propertes.
              fOK = AutoSizeControl(ctlControl)
                
              If fOK Then
                ctlControl.Selected = True
                fOK = SelectControl(ctlControl)
              End If
              
            End If
            
            If Me.abWebForm.Tools("ID_AutoLabel").Checked = True Then
              AutoLabel pVarPageContainer, pSngX, pSngY, sColumnName
            End If
            
            If fOK Then
              'The ActiveBar control does not have the visible property, so to avoid err
              'we only check the visible property of other controls.
              If ctlControl.Name <> "abWebForm" Then
                ctlControl.Visible = True
                ctlControl.ZOrder 0
              End If
            End If
            
            Set ctlControl = Nothing
          
          End If
        End With
        
      ' If we are dropping a standard control ...
      ElseIf pCtlSource Is frmWorkflowWFToolbox.trvStandardControls Then
        sAutoLabel = ""
        iControlType = giWFFORMITEM_UNKNOWN
        
        ' Add the new control to the screen.
        Select Case frmWorkflowWFToolbox.trvStandardControls.SelectedItem.key
        
          Case "BUTTON"
            iControlType = giWFFORMITEM_BUTTON

          Case "IMAGECTRL"
            iControlType = giWFFORMITEM_IMAGE
            
          Case "INPUT_CHARACTER"
            iControlType = giWFFORMITEM_INPUTVALUE_CHAR

          Case "INPUT_DATE"
            iControlType = giWFFORMITEM_INPUTVALUE_DATE
          
          Case "INPUT_DROPDOWN"
            iControlType = giWFFORMITEM_INPUTVALUE_DROPDOWN
          
          Case "INPUT_LOGIC"
            iControlType = giWFFORMITEM_INPUTVALUE_LOGIC
            
          Case "INPUT_LOOKUP"
            iControlType = giWFFORMITEM_INPUTVALUE_LOOKUP
            
          Case "INPUT_NUMERIC"
            iControlType = giWFFORMITEM_INPUTVALUE_NUMERIC
          
          Case "INPUT_GRID"
            iControlType = giWFFORMITEM_INPUTVALUE_GRID
            
          Case "INPUT_OPTIONGROUP"
            iControlType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP

          Case "LABELCTRL"
            iControlType = giWFFORMITEM_LABEL
            
          Case "FRAMECTRL"
            iControlType = giWFFORMITEM_FRAME

          Case "LINECTRL"
            iControlType = giWFFORMITEM_LINE
        
          Case "INPUT_FILEUPLOAD"
            iControlType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD
          
        End Select
      
        If iControlType <> giWFFORMITEM_UNKNOWN Then
          Set ctlControl = AddControl(iControlType)
          sCaption = UniqueCaption(ctlControl)
      
          If WebFormControl_HasAutoLabel(iControlType) Then
            sAutoLabel = sCaption
          End If
        End If
        
        fOK = Not (ctlControl Is Nothing)
        
        'Check that a new control was added successfully
        If fOK Then
  
          With ctlControl

            ' Set the last action flag and enable the Undo menu option.
            If (Me.abWebForm.Tools("ID_AutoLabel").Checked) And (Len(sAutoLabel) > 0) Then
              giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL
            Else
              giLastActionFlag = giACTION_DROPCONTROL
            End If
            giUndo_ControlIndex = .Index
            gsUndo_ControlType = .Name
          
            Set .Container = pVarPageContainer
            .Left = AlignX(CLng(pSngX))
            .Top = AlignY(CLng(pSngY))
            
            ' Setting this to 0 reset the defaulted table for recordSelectors!
            '.ColumnID = 0
            
            iControlType = .WFItemType
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_WFIDENTIFIER) Then
              .WFIdentifier = sCaption
            End If
            
            ' Initialise the new control's font and forecolour.
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
              Set objFont = New StdFont
              objFont.Name = Me.Font.Name
              objFont.Size = Me.Font.Size
              objFont.Bold = Me.Font.Bold
              objFont.Italic = Me.Font.Italic
              Set .Font = objFont
              Set objFont = Nothing
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADFONT) Then
              Set objFont = New StdFont
              objFont.Name = Me.Font.Name
              objFont.Size = Me.Font.Size
              objFont.Bold = Me.Font.Bold
              objFont.Italic = Me.Font.Italic
              Set ctlControl.HeadFont = objFont
              Set objFont = Nothing
            End If
              
            If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
              .WFInputSize = 8
              .WFInputDecimals = 0
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
              .WFInputSize = 50
              .WFInputDecimals = 0
            End If
            
            If iControlType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
              .WFInputSize = WORKFLOWWEBFORM_MAXSIZE_FILEUPLOAD
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
              .ForeColor = Me.ForeColor
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
              Select Case iControlType
                Case giWFFORMITEM_BUTTON, _
                  giWFFORMITEM_INPUTVALUE_FILEUPLOAD
                  .BackColor = 16249587
                
                Case giWFFORMITEM_LINE
                  .BackColor = 10172547
                
                Case giWFFORMITEM_INPUTVALUE_CHAR, _
                  giWFFORMITEM_INPUTVALUE_DATE, _
                  giWFFORMITEM_INPUTVALUE_NUMERIC, _
                  giWFFORMITEM_INPUTVALUE_DROPDOWN, _
                  giWFFORMITEM_INPUTVALUE_LOOKUP
                  ctlControl.BackColor = 15988214
                
                Case giWFFORMITEM_INPUTVALUE_GRID
                  ctlControl.BackColor = 16777215
                
                Case Else
                  ctlControl.BackColor = Me.BackColor
                End Select
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADERBACKCOLOR) Then
              ctlControl.HeaderBackColor = 16248553
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOREVEN) Then
              ctlControl.ForeColorEven = 6697779
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORODD) Then
              ctlControl.ForeColorOdd = 6697779
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOREVEN) Then
              ctlControl.BackColorEven = 15988214
            End If
                        
            If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORODD) Then
              ctlControl.BackColorOdd = 15988214
            End If
                        
            If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
              ctlControl.BackColorHighlight = 10480637
            End If
                        
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
              ctlControl.ForeColorHighlight = 2774907
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION) Then
              .Caption = sCaption
            End If
            
            ' Default the control's propertes.
            fOK = AutoSizeControl(ctlControl)
            
            If fOK Then
              ctlControl.Selected = True
              fOK = SelectControl(ctlControl)
            End If
            
            If fOK Then
              If (Me.abWebForm.Tools("ID_AutoLabel").Checked) And (Len(sAutoLabel) > 0) Then
                AutoLabel pVarPageContainer, pSngX, pSngY, sAutoLabel
              End If
              
              .Visible = True
              
              ' Put frame at the back
              If iControlType = giWFFORMITEM_FRAME And gbAutoSendFrameToBack Then
                .ZOrder 1
              Else
                .ZOrder 0
              End If
            
            End If
          End With
        End If
        
        ' Disassociate object variables.
        Set ctlControl = Nothing
        
      ' If we are dropping a workflow web form value ...
      ElseIf pCtlSource Is frmWorkflowWFToolbox.trvWorkflowValue Then
        fIsFile = False
        
        With frmWorkflowWFToolbox.trvWorkflowValue.SelectedItem
          tmpValue = "<" & .Parent.Text
          tmpValue = tmpValue & " : " & .Text & ">"
          sElementIdentifier = .Parent.Text
          sItemIdentifier = .Text
        End With
        
        For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
          Set wfTemp = maWFPrecedingElements(iLoop)

          If UCase(Trim(wfTemp.Identifier)) = UCase(Trim(sElementIdentifier)) Then
            If wfTemp.ElementType = elem_WebForm Then
              asItems = wfTemp.Items

              For iLoop2 = 1 To UBound(asItems, 2)
                If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(sItemIdentifier)) Then
                  fIsFile = (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)
                  Exit For
                End If
              Next iLoop2
            End If
            
            Exit For
          End If

          Set wfTemp = Nothing
        Next iLoop
        
        If fIsFile Then
          Set ctlControl = AddControl(giWFFORMITEM_WFFILE)
        Else
          Set ctlControl = AddControl(giWFFORMITEM_WFVALUE)
        End If
        fOK = Not (ctlControl Is Nothing)
        
        'Check that a new control was added successfully
        If fOK Then
          With ctlControl
            .WFWorkflowForm = sElementIdentifier
            .WFWorkflowValue = sItemIdentifier

            ' Set the last action flag and enable the Undo menu option.
            If Me.abWebForm.Tools("ID_AutoLabel").Checked = True Then
              giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL
            Else
              giLastActionFlag = giACTION_DROPCONTROL
            End If
            giUndo_ControlIndex = .Index
            gsUndo_ControlType = .Name
          
            Set .Container = pVarPageContainer
            .Left = AlignX(CLng(pSngX))
            .Top = AlignY(CLng(pSngY))
            .ColumnID = 0
            
            ' Initialise the new control's font and forecolour.
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
              Set objFont = New StdFont
              objFont.Name = Me.Font.Name
              objFont.Size = Me.Font.Size
              objFont.Bold = Me.Font.Bold
              objFont.Italic = Me.Font.Italic
              
              If fIsFile Then
                objFont.Underline = True
              End If
              
              Set .Font = objFont
              Set objFont = Nothing
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
              .ForeColor = Me.ForeColor
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
              ctlControl.BackColor = Me.BackColor
            End If
            
            If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION) Then
              .Caption = tmpValue
              
              If fIsFile Then
                .ToolTipText = tmpValue
              End If
            End If
            
            ' Default the control's propertes.
            fOK = AutoSizeControl(ctlControl)
              
            If fOK Then
              ctlControl.Selected = True
              fOK = SelectControl(ctlControl)
            End If
            
            If Me.abWebForm.Tools("ID_AutoLabel").Checked = True Then
              AutoLabel pVarPageContainer, pSngX, pSngY, ctlControl.WFWorkflowValue
            End If
            
            If fOK Then
              .Visible = True
              
              ' Put frame at the back
              If iControlType = giWFFORMITEM_FRAME And gbAutoSendFrameToBack Then
                .ZOrder 1
              Else
                .ZOrder 0
              End If
              
            End If
          End With
        End If
        
        ' Disassociate object variables.
        Set ctlControl = Nothing
        
      End If
    
      ' Set focus on the screen designer form.
      Me.SetFocus
  
    End If
  End If
    
  If fOK Then
    ' Mark the screen as having changed.
    mfChanged = True
    frmSysMgr.RefreshMenu
  
    ' Refresh the properties screen.
    Set frmWorkflowWFItemProps.CurrentWebForm = Me
    frmWorkflowWFItemProps.RefreshProperties
  End If
  
  RefreshBlankDesignLabel

TidyUpAndExit:
  ' Disassociate object variables.
  Set objMisc = Nothing
  Set objFont = Nothing
  Set ctlControl = Nothing
  ' Return the success/failure value.
  DropControl = fOK
  Exit Function

ErrorTrap:
  ' Flag the error.
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub Form_DblClick()
  ShowPropertiesForm
  
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

  ' Drop a control onto the screen.
  On Error GoTo ErrorTrap
  
  If CurrentContainer Is Me Then
    If Not DropControl(Me, Source, X, Y) Then
      MsgBox "Unable to drop the control." & vbCr & vbCr & _
        Err.Description, vbExclamation + vbOKOnly, App.ProductName
    End If
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_GotFocus()
  ' Refresh the properties screen.
  Set frmWorkflowWFItemProps.CurrentWebForm = Me
  frmWorkflowWFItemProps.RefreshProperties
End Sub

Private Sub Form_Initialize()

  ' Initialise global variables.
  On Error GoTo ErrorTrap

  gfMultiSelecting = False
  gfExitToWorkflowDesigner = False
  
  gbAutoSendFrameToBack = True
  
  ' Initialise properties.
  mfAlignToGrid = True
  giGridX = 40
  giGridY = 40
  
  ASRSelectionMarkers(0).Visible = False
  
  ' Hide the dummy control array controls.
  With asrDummyLabel(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With asrDummyTextBox(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With asrDummyImage(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With asrDummyFrame(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With asrDummyCombo(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With asrDummyCheckBox(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With ASRDummyGrid(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With btnWorkflow(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  With ASRDummyFileUpload(0)
    .Left = -.Width
    .Top = -.Height
    .Visible = False
  End With
  
  ' Disable the 'undo' menu option until we have somethig to undo.
  giLastActionFlag = giACTION_NOACTION

  RefreshBlankDesignLabel
  
  ' Initialise gloabl arrays.
  ReDim gactlUndo_DeletedControls(0)
  ReDim gactlClipboardControls(0)
  ReDim gavUndo_PastedControls(2, 0)

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error initialising Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim iCount As Integer
  
  ' Complete stretching of selected controls
  If mbKeyStretching Then
    ASRSelectionMarkers_StretchEnd 0, "", 0, 0, 0, 0
  End If

  ' Complete moving of selected controls
  If mbKeyMoving Then
     For iCount = 1 To ASRSelectionMarkers.Count - 1
      With ASRSelectionMarkers(iCount)
        If .Visible Then
          .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize
          .ShowSelectionMarkers True
        End If
      End With
     Next iCount
  End If

  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyUp(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

  RefreshBlankDesignLabel
  
End Sub

Private Sub Form_Load()

  On Error GoTo ErrorTrap
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  ReDim mavIdentifierLog(6, 0)
  
  RefreshBlankDesignLabel

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error loading Web Form Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub

Public Sub RefreshBlankDesignLabel()
  Me.lblBlankDesigner.Visible = (WebFormControlsCount < 1)
  Me.lblBlankDesigner.ZOrder 0
  Me.lblBlankDesigner.Left = (Me.ScaleWidth - lblBlankDesigner.Width) / 2
  Me.lblBlankDesigner.Top = (Me.ScaleHeight - lblBlankDesigner.Height) / 2
End Sub

Public Property Get IsChanged() As Boolean
  ' Return the 'Web Form changed' flag.
  IsChanged = mfChanged
End Property

Public Property Let IsChanged(pfNewValue As Boolean)
  ' Set the 'Web Form changed' flag.
  mfChanged = pfNewValue
  ' Menu may be dependent on the status of the Web Form.
  frmSysMgr.RefreshMenu
End Property

Public Property Get GridX() As Long
  ' Return the horizontal grid interval.
  GridX = giGridX
End Property

Public Property Let GridX(plngGridSize As Long)
  ' Set the horizontal grid interval.
  giGridX = plngGridSize
End Property

Public Property Get GridY() As Long
  ' Return the vertical grid interval.
  GridY = giGridY
End Property

Public Property Let GridY(plngGridSize As Long)
  ' Set the vertical grid interval.
  giGridY = plngGridSize
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Process key strokes.
  On Error GoTo ErrorTrap
  
  Dim sngXMove As Single
  Dim sngYMove As Single
  Dim sngXStretch As Single
  Dim sngYStretch As Single
  Dim strDirection As String
  Dim iCount As Integer
  Dim bHandled As Boolean
  
  bHandled = False
  mbKeyMoving = False
  mbKeyStretching = False

  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
  
  'TODO - Right Click Menu
  ' F4 needs to bring up properties dialog
  If KeyCode = vbKeyF4 Then
    EditMenu "ID_ObjectProperties"
    bHandled = True
  End If

  If (Not mfReadOnly) Then
    ' DELETE key deletes any selected controls.
    ' If there are no selected controls then the current tab page is deleted.
    ' If there are no selected controls and no tab pages then nothing happens.
    If Not bHandled Then
      If KeyCode = vbKeyDelete Then
        If SelectedControlsCount > 0 Then
          If Not DeleteSelectedControls Then
            MsgBox "Unable to delete controls." & vbCr & vbCr & _
              Err.Description, vbExclamation + vbOKOnly, App.ProductName
          Else
            bHandled = True
          End If
        End If
      End If
    End If
    
    ' SHIFT and ARROW keys stretch any selected controls.
    If Not bHandled Then
      If ((Shift And vbShiftMask) > 0) Then
      
        ' Determine which stretch is being done.
        Select Case KeyCode
          Case vbKeyLeft
            strDirection = "CentreRight"
            sngXStretch = -giSTANDARDMOVEMENT
          Case vbKeyRight
            strDirection = "CentreRight"
            sngXStretch = giSTANDARDMOVEMENT
          Case vbKeyUp
            strDirection = "BottomCentre"
            sngYStretch = -giSTANDARDMOVEMENT
          Case vbKeyDown
            strDirection = "BottomCentre"
            sngYStretch = giSTANDARDMOVEMENT
        End Select
      
        ' Stretch the selected controls if required.
        If (sngXStretch <> 0) Or (sngYStretch <> 0) Then
          ASRSelectionMarkers_StretchStart 0, strDirection, 0, 0, sngXStretch, sngYStretch
          ASRSelectionMarkers_Stretch 0, strDirection, 0, 0, sngXStretch, sngYStretch
          mbKeyStretching = True
        End If
      End If
  
      ' CTRL and ARROW keys move the selected controls.
      If ((Shift And vbCtrlMask) > 0) Then
    
        mbKeyMoving = True
    
        sngXMove = 0
        sngYMove = 0
        
        ' Determine which movement is being made.
        Select Case KeyCode
          Case vbKeyLeft
            sngXMove = -giSTANDARDMOVEMENT
          Case vbKeyRight
            sngXMove = giSTANDARDMOVEMENT
          Case vbKeyUp
            sngYMove = -giSTANDARDMOVEMENT
          Case vbKeyDown
            sngYMove = giSTANDARDMOVEMENT
    
        End Select
  
        ' Flag the selected selction markers to be moved
        If (sngXMove <> 0) Or (sngYMove <> 0) Then
          For iCount = 1 To ASRSelectionMarkers.Count - 1
            ASRSelectionMarkers(iCount).ShowSelectionMarkers False
          Next iCount
        
          WebFormControl_KeyMove sngXMove, sngYMove
        End If
     
      End If
  
      ' CTRL-keys
      If ((Shift And vbCtrlMask) > 0) Then
      
        Select Case KeyCode
          Case vbKeyZ
            EditMenu "ID_Undo"
          Case vbKeyX
            EditMenu "ID_Cut"
          Case vbKeyC
            EditMenu "ID_Copy"
          Case vbKeyV
            EditMenu "ID_Paste"
          Case vbKeyA
            EditMenu "ID_ScreenSelectAll"
          
        End Select
              
        bHandled = True
      End If
    End If
  End If
  
  If Not bHandled Then
    bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  End If

  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If
  
  RefreshBlankDesignLabel
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Public Property Set DefaultFont(pobjNewValue As Object)
  ' Set the Web Form's default font.
  Set Me.Font = pobjNewValue
End Property

Public Property Get DefaultFont() As Object
  ' Return the Web Form's default font.
  Set DefaultFont = Me.Font
End Property

Public Sub EditMenu(ByVal psMenuOption As String)
  ' Process the menu options.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim lngPictureID As Long
  Dim sFileName As String
  
  Select Case psMenuOption
    ' Undo the last deletion, cut or addition of a control.
    Case "ID_Undo"
      If Not UndoLastAction Then
        MsgBox "Unable to undo the last action." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Cut the selected controls.
    Case "ID_Cut"
      If Not CutSelectedControls Then
        MsgBox "Unable to cut controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Copy the selected control.
    Case "ID_Copy"
      If Not CopySelectedControls Then
        MsgBox "Unable to copy controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Paste the control from the clipboard.
    Case "ID_Paste"
      If Not PasteControls Then
        MsgBox "Unable to paste controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Delete the selected control.
    Case "ID_ScreenObjectDelete"
      ' If there are no selected controls then the current tab page is deleted.
      ' If there are no selected controls and no tab pages then nothing happens.
      If SelectedControlsCount > 0 Then
        If Not DeleteSelectedControls Then
          MsgBox "Unable to delete controls." & vbCr & vbCr & _
            Err.Description, vbExclamation + vbOKOnly, App.ProductName
        End If
      End If
    
    ' Select all controls on the current Web Form.
    Case "ID_ScreenSelectAll"
      If Not SelectAllControls Then
        MsgBox "Unable to select controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Save the current Web Form.
    Case "ID_mnuWFSave"
      SaveWebForm
    
    ' Display the object properties grid.
    Case "ID_ObjectProperties"
      If Not frmWorkflowWFItemProps.Visible Then
        frmWorkflowWFItemProps.Show
      Else
        frmWorkflowWFItemProps.WindowState = vbNormal
        frmWorkflowWFItemProps.ZOrder 0
      End If
            
    ' Display the object properties screen.
    Case "ID_ObjectPropertiesScreen"
      ShowPropertiesForm
      
    ' Display the WebForm properties screen.
    Case "ID_WebFormPropertiesScreen"
      ShowPropertiesForm True
  
    ' Display the Toolbox window.
    Case "ID_Toolbox"
      If Not frmWorkflowWFToolbox.Visible Then
        frmWorkflowWFToolbox.Show
      Else
        frmWorkflowWFToolbox.WindowState = vbNormal
        frmWorkflowWFToolbox.ZOrder 0
      End If
     
    Case "ID_AutoLabel"
      If mblnAutoLabelling = True Then Exit Sub
      mblnAutoLabelling = True
      'Set the checked property of the AutoLabel button.
      Me.abWebForm.Tools("ID_AutoLabel").Checked = Not Me.abWebForm.Tools("ID_AutoLabel").Checked
      frmSysMgr.tbMain.Tools("ID_AutoLabel").Checked = Me.abWebForm.Tools("ID_AutoLabel").Checked
      mblnAutoLabelling = False
     
    ' Bring selected controls to front
    Case "ID_BringToFront"
      BringSelectedControlsToFront
    
    ' Send selected controls to back
    Case "ID_SendToBack"
      SendSelectedControlsToBack

    ' Left align selected controls
    Case "ID_ScreenControlAlignLeft"
      LeftAlignSelectedControls
    
    ' Centre align selected controls
    Case "ID_ScreenControlAlignCentre"
      CentreAlignSelectedControls

    ' Right align selected controls
    Case "ID_ScreenControlAlignRight"
      RightAlignSelectedControls
     
    ' Top align selected controls
    Case "ID_ScreenControlAlignTop"
      TopAlignSelectedControls
     
    ' Middle align selected controls
    Case "ID_ScreenControlAlignMiddle"
      MiddleAlignSelectedControls
     
    ' Bottom align selected controls
    Case "ID_ScreenControlAlignBottom"
      BottomAlignSelectedControls
    
    ' Call the pop-up that allows the user to define the object
    ' tab order for the current screen.
    Case "ID_ObjectOrder"
      Set frmWorkflowWFTabOrder.CurrentScreen = Me
      frmWorkflowWFTabOrder.Show vbModal
      Set frmWorkflowWFTabOrder = Nothing
    
  End Select
  
  RefreshBlankDesignLabel
  
  Exit Sub
  
ErrorTrap:

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim VarPageContainer As Variant

  ' Used to work out where to paste controls
  mlngMouseX = X
  mlngMouseY = Y

  ' Only handle left button presses here.
  If Button <> vbLeftButton Then
    Exit Sub
  End If
  
  ' Deselect all screen controls unless the SHIFT or CTRL keys are pressed.
  If ((Shift And vbShiftMask) = 0) And ((Shift And vbCtrlMask) = 0) Then
    fOK = DeselectAllControls
  Else
    fOK = True
  End If
  
  If fOK Then
    ' Start the multi-selection frame.
    gfMultiSelecting = True
    gLngMultiSelectionXStart = X
    gLngMultiSelectionYStart = Y
      
    Set VarPageContainer = CurrentContainer
    
    ' Position and display the multi-selection box.
    With asrboxMultiSelection
      .Left = gLngMultiSelectionXStart
      .Top = gLngMultiSelectionYStart
      .Width = 0
      .Height = 0
      Set .Container = VarPageContainer
      .Visible = True
      .ZOrder 0
    End With
  
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set VarPageContainer = Nothing
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  ' Position and size the multi-selection lines as required.
  On Error GoTo ErrorTrap
  
  Dim lngTop As Long
  Dim lngLeft As Long
  Dim lngRight As Long
  Dim lngBottom As Long
  Dim lngRightLimit As Long
  Dim lngBottomLimit As Long
  
  If gfMultiSelecting Then

    ' Calculate the cordinates of the multi-selection area.
    If X < gLngMultiSelectionXStart Then
      lngLeft = X
      lngRight = gLngMultiSelectionXStart
    Else
      lngLeft = gLngMultiSelectionXStart
      lngRight = X
    End If
      
    If Y < gLngMultiSelectionYStart Then
      lngTop = Y
      lngBottom = gLngMultiSelectionYStart
    Else
      lngTop = gLngMultiSelectionYStart
      lngBottom = Y
    End If

    lngRightLimit = Me.Width - (2 * XFrame) - XBorder
    lngBottomLimit = Me.Height - (2 * YFrame) - CaptionHeight - (4 * YBorder)
    
    If lngLeft < 0 Then lngLeft = 0
    If lngRight > lngRightLimit Then lngRight = lngRightLimit
    If lngTop < 0 Then lngTop = 0
    If lngBottom > lngBottomLimit Then lngBottom = lngBottomLimit
    
    ' Size and position the multi-selection box.
    With asrboxMultiSelection
      .Left = lngLeft
      .Top = lngTop
      .Width = lngRight - lngLeft
      .Height = lngBottom - lngTop
    End With

    Me.Refresh

  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select control that lie within the multi-selection area.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fInVerticalBand As Boolean
  Dim fInHorizontalBand As Boolean
  Dim sngSelectionTop As Single
  Dim sngSelectionLeft As Single
  Dim sngSelectionRight As Single
  Dim sngSelectionBottom As Single
  Dim lngXMouse As Long
  Dim lngYMouse As Long
  Dim ctlControl As VB.Control
  Dim VarPageContainer As Variant
  Dim bInSelectionBand As Boolean
  
  Select Case Button
    
    ' Handle left button presses.
    Case vbLeftButton
     
      ' Put up an hourglass
      Screen.MousePointer = vbHourglass
   
      ' End the multi-selection.
      gfMultiSelecting = False
    
      
      Set VarPageContainer = asrboxMultiSelection.Container
      
      ' Hide the multi-selection box and move it onto the form.
      ' NB. This is done so that it is not left on any tabpage containers, thus
      ' making it impossible to unload the tab pages.
      With asrboxMultiSelection
        sngSelectionTop = .Top
        sngSelectionBottom = .Top + .Height
        sngSelectionLeft = .Left
        sngSelectionRight = .Left + .Width
        Set .Container = Me
        .Visible = False
      End With
      
      ' Lock the window refresh.
'      UI.LockWindow Me.hWnd
      
      ' Select thr highlighted controls
      For Each ctlControl In Me.Controls
        'The ActiveBar control does mot have the visible property, so to avoid err
        'we only check the visible property of other controls.
        If ctlControl.Name <> "abWebForm" Then
          If ctlControl.Visible Then

            If IsWebFormControl(ctlControl) Then
              With ctlControl

                fInVerticalBand = (.Left < sngSelectionRight) And (.Left + .Width > sngSelectionLeft)
                fInHorizontalBand = (.Top < sngSelectionBottom) And (.Top + .Height > sngSelectionTop)


                ' Only include the frame if the rubber band crosses a line (i.e. skip if only controls within frame are selected)
                If ctlControl.Name = "asrDummyFrame" Then

                  'If band is entiterly within selection band dont select the frame
                  bInSelectionBand = IIf((sngSelectionLeft > .Left) And (sngSelectionRight < .Left + .Width) _
                    And (sngSelectionTop > .Top) And (sngSelectionBottom < .Top + .Height), False, fInVerticalBand And fInHorizontalBand)

                Else
                  bInSelectionBand = fInVerticalBand And fInHorizontalBand
                End If

                If bInSelectionBand Then 'fInHorizontalBand And fInVerticalBand Then

                  ' Holding down control now deselects controls
                  If ((Shift And vbCtrlMask) = 2) And .Selected Then
                    DeselectControl ctlControl
                    .Selected = False
                  Else
                    .Selected = True
                    SelectControl ctlControl
                  End If

                End If

              End With
            End If

          End If
        End If

      Next ctlControl
      
      ' Disassociate object variables.
      Set ctlControl = Nothing
      Set VarPageContainer = Nothing
  
      ' Unlock the window refresh.
      'UI.UnlockWindow
      
      ' Mark the screen as having changed.
      frmSysMgr.RefreshMenu
      
      ' Refresh the properties screen.
      Set frmWorkflowWFItemProps.CurrentWebForm = Me
      frmWorkflowWFItemProps.RefreshProperties
 
      ' Handle right button presses.
      Case vbRightButton
        UI.GetMousePos lngXMouse, lngYMouse
        frmSysMgr.tbMain.Bands("ID_mnuWebFormEdit").TrackPopup -1, -1
        
  End Select
  
TidyUpAndExit:
  
  ' Close the progress bar
  gobjProgress.CloseProgress
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set VarPageContainer = Nothing
  
  ' Reset the screen mousepointer.
  Screen.MousePointer = vbNormal
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the screen has been modified then prompt the user
  ' whether or not to save the changes.
  On Error GoTo ErrorTrap
  
  frmWorkflowWFItemProps.ApplyCurrentProperty
  
  If mfReadOnly Then mfChanged = False
  
  If mfChanged Then
    Select Case MsgBox("Apply web form changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
        mblnDisplayScrOpen = False
      Case vbYes
        Cancel = (Not SaveWebForm())
        If Cancel = True Then mblnDisplayScrOpen = False Else mblnDisplayScrOpen = True
      Case vbNo
        mfChanged = False
        mblnDisplayScrOpen = True
    
        ' Restore the original expression definitions
        RestoreOriginalExpressions
    End Select
  End If

  ' Set the flag that determines whether we need to display the screen manager
  ' after the screen designer is unloaded.
  gfExitToWorkflowDesigner = (UnloadMode = vbFormControlMenu) And mblnDisplayScrOpen
  If Not mfChanged Then gfExitToWorkflowDesigner = True

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
End Sub

Private Sub Form_Resize()

  ' Resize the form.
  On Error GoTo ErrorTrap

  ' Only perform the resize method if the form is not minimized.
  If Me.WindowState <> vbMinimized Then
    
    If Not mfLoading Then
      Call MoveAndPersistControls
    End If
    
    mlngLastFormWidth = Me.Width
    mlngLastFormheight = Me.Height
  End If
  
  RefreshBlankDesignLabel
  
  If Not mfLoading Then mfChanged = True
  
  'This is required so that the window state menu is refreshed.
  'However it makes everything flash so I'd like to change it.
  frmSysMgr.RefreshMenu

  frmWorkflowWFItemProps.RefreshProperties

  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error resizing Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Sub

Public Function XFrame() As Double
  ' Return the width of a control frame.
  XFrame = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
End Function

Public Function YFrame() As Double
  ' Return the height of a control frame.
  YFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
End Function
Private Function XBorder() As Double
  ' Return the width of a control border.
  XBorder = UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX
End Function

Private Function YBorder() As Double
  ' Return the height of a control border.
  YBorder = UI.GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY
End Function
Private Function CaptionHeight() As Double
  ' Return the height of a form's caption bar.
  CaptionHeight = UI.GetSystemMetrics(SM_CYSMCAPTION) * Screen.TwipsPerPixelY
End Function

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Private Sub Form_Unload(Cancel As Integer)
  
  On Error GoTo ErrorTrap
  
  Dim iForm As Integer
  Dim awfElements() As COAWF_Webform
  
  UnLoad frmWorkflowWFItemProps
  UnLoad frmWorkflowWFToolbox
  
  ' Display the Workflow Designer form if we are not exiting the system.
  If gfExitToWorkflowDesigner Then
    For iForm = 0 To Forms.Count - 1 Step 1
      If Forms(iForm).Name = "frmWorkflowDesigner" Then
        
        If mfChanged Then
          Forms(iForm).IsChanged = True
        ElseIf Not Forms(iForm).IsChanged Then
          Forms(iForm).IsChanged = Forms(iForm).WorkflowExpressionsChanged
        End If
        
        ReDim awfElements(1)
        Set awfElements(1) = mwfElement
        Forms(iForm).UpdateIdentifiers mwfElement, awfElements, mavIdentifierLog
        Forms(iForm).Show
      End If
    Next iForm

  End If

  Unhook Me.hWnd

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub asrDummyFrame_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyFrame(Index), Source, X, Y
End Sub

Private Sub asrDummyImage_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop asrDummyImage(Index), Source, X, Y
End Sub

Private Function DeleteSelectedControls(Optional pbIsCutting As Boolean) As Boolean
  ' Delete the selected controls.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim avWebFormControls() As Variant
  Dim ctlControl As VB.Control
  Dim iSelectedControls As Integer
  Dim fDeleteOK As Boolean
  Dim avMessages() As Variant
  Dim iLoop As Integer
  Dim frmUsage As frmUsage
  Dim iItemIndex As Integer
  
  ' How many controls do we have
  iSelectedControls = SelectedControlsCount

 ' Open a progress bar
  With gobjProgress
    .Caption = "Web Form Designer"
    .Bar1Value = 0
    .Bar1MaxValue = iSelectedControls
    .Bar1Caption = IIf(pbIsCutting, "Cutting", "Deleting") & " Web Form Items..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  fOK = True
  
  CurrentContainer.SetFocus
  
  ' Do nothing if there are no selected controls.
  If iSelectedControls > 0 Then
    fDeleteOK = True
    ReDim avMessages(3, 0)
    
    ' Construct an array of the selected controls.
    ReDim avWebFormControls(0)
    For Each ctlControl In Me.Controls
      If IsWebFormControl(ctlControl) Then
        If ctlControl.Selected Then
          iIndex = UBound(avWebFormControls) + 1
          ReDim Preserve avWebFormControls(iIndex)
          Set avWebFormControls(iIndex) = ctlControl
          
          If ControlIsUsed(ctlControl, avMessages) Then
            fDeleteOK = False
          End If
        End If
      End If
    Next ctlControl
    
    ' Disassociate object variables.
    Set ctlControl = Nothing
  
    If Not fDeleteOK Then
      If UBound(avMessages, 2) > 0 Then
        Set frmUsage = New frmUsage
        frmUsage.ResetList

        For iLoop = 1 To UBound(avMessages, 2)
          frmUsage.AddToList avMessages(1, iLoop) & " - " & avMessages(2, iLoop), avMessages(3, iLoop)
        Next iLoop
    
        ' Close progress bar
        gobjProgress.CloseProgress

        Screen.MousePointer = vbNormal
    
        frmUsage.Width = (3 * Screen.Width / 4)
    
        frmUsage.ShowMessage "Web Form '" & Trim(msWFIdentifier) & "'", _
          "The selected item(s) cannot be deleted as they are used as follows:", _
          UsageCheckObject.Workflow, _
          USAGEBUTTONS_PRINT + USAGEBUTTONS_OK + USAGEBUTTONS_SELECT
    
        If frmUsage.Choice = vbRetry Then
          ' Highlight the item 'selected' in the usage check form.
          DeselectAllControls

          If frmUsage.Selection >= 0 Then
            iItemIndex = CInt(frmUsage.Selection)

            If iItemIndex > 0 Then
              For Each ctlControl In Me.Controls
                If IsWebFormControl(ctlControl) Then
                  If ctlControl.TabIndex = iItemIndex Then
                    ctlControl.Selected = True
                    SelectControl ctlControl
                    Exit For
                  End If
                End If
              Next ctlControl
              
              Set ctlControl = Nothing
              
'  '            mcolwfElements(CStr(iElementIndex).HighLighted = True
'              SelectElement mcolwfElements(CStr(iElementIndex))
'
'              'JPD 20061129 Fault 11533 - Ensure the selected element is visible.
'              MoveToItem mcolwfElements(CStr(iElementIndex))
'
'              ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
'              miSelectionOrder(UBound(miSelectionOrder)) = iElementIndex
'
'              RefreshMenu
            End If
          End If
        End If
    
        UnLoad frmUsage
        Set frmUsage = Nothing
      End If
    Else
      ' Clear the array of deleted controls.
      For iIndex = 1 To UBound(gactlUndo_DeletedControls)
        Set ctlControl = gactlUndo_DeletedControls(iIndex)
        UnLoad ctlControl
        ' Disassociate object variables.
        Set ctlControl = Nothing
      Next iIndex
      ReDim gactlUndo_DeletedControls(0)
    
      ' Move all selected screen controls from the screen into the array of deleted controls.
      For iIndex = 1 To UBound(avWebFormControls)
             
        Set ctlControl = avWebFormControls(iIndex)
  
        iIndex2 = UBound(gactlUndo_DeletedControls) + 1
        ReDim Preserve gactlUndo_DeletedControls(iIndex2)
        Set gactlUndo_DeletedControls(iIndex2) = ctlControl
  
        With ctlControl
          If ctlControl.Tag > 0 Then
            
            ' Hide the selection markers
            ASRSelectionMarkers(ctlControl.Tag).Visible = False
            ASRSelectionMarkers(ctlControl.Tag).AttachedObject = gactlUndo_DeletedControls(iIndex2)
            
            ' Unload the control's selection markers.
            gobjProgress.UpdateProgress
            
            If Not fOK Then
              Exit For
            End If
          End If
    
          '.Tag = 0
          .Visible = False
          .Selected = False
        End With
        
        ' Disassociate object variables.
        Set ctlControl = Nothing
      Next iIndex
  
      ' Mark the screen as having changed.
      mfChanged = True
      
      If fOK Then
        ' Set the last action flag and enable the Undo menu option.
        giLastActionFlag = giACTION_DELETECONTROLS
        frmSysMgr.RefreshMenu
      End If
    End If
  End If
  
TidyUpAndExit:
  
  ' Close progress bar
  gobjProgress.CloseProgress
  
  frmWorkflowWFItemProps.RefreshProperties
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  ' Return the success/failure value.
  DeleteSelectedControls = fOK
  Exit Function
  
ErrorTrap:
  ' Flag the error.
  fOK = False
  Resume TidyUpAndExit
  
End Function

' Return a count of the number of selected controls.
Public Function SelectedControlsCount() As Integer
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim iSelectedControls As Integer
  
  ' Initialize the count.
  iSelectedControls = 0
  
  ' Count the number of custom screen controls that are selected.
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    iSelectedControls = iSelectedControls + IIf(ASRSelectionMarkers(iCount).Visible, 1, 0)
  Next iCount
        
TidyUpAndExit:
  SelectedControlsCount = iSelectedControls
  Exit Function
  
ErrorTrap:
  iSelectedControls = 0
  Resume TidyUpAndExit
  
End Function

Public Function WebFormControlsCount(Optional pbShowThisPageOnly As Boolean) As Integer
  ' Return a count of the number of selected controls.
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim ctlControl As VB.Control
  
  ' Initialize the count.
  iCount = 0
  
  ' Count the number of custom screen controls that are selected.
  For Each ctlControl In Me.Controls
    'The ActiveBar control does mot have the visible property, so to avoid err
    'we only check the visible property of other controls.
    If ctlControl.Name <> "abWebForm" Then
      If ctlControl.Visible Or Not pbShowThisPageOnly Then
        If IsWebFormControl(ctlControl) Then
          iCount = iCount + 1
        End If
      End If
    End If
  Next ctlControl
        
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  WebFormControlsCount = iCount
  Exit Function
  
ErrorTrap:
  iCount = 0
  Resume TidyUpAndExit

End Function

Public Function ClipboardControlsCount() As Integer
  ' Return a count of the number of controls in the clipboard control.
  ClipboardControlsCount = UBound(gactlClipboardControls)
End Function

Private Function ReadColumnControlValues(plngColumnID As Long) As Variant
  ' Return an array of control values for the given column.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim avValues As Variant
  Dim asResults() As String
  Dim sSQL As String
  Dim rsControlValues As dao.Recordset
  
  ' Pull the column control values from the database.
  sSQL = "SELECT value" & _
    " FROM tmpControlValues" & _
    " WHERE columnID = " & plngColumnID & _
    " ORDER BY sequence"
  Set rsControlValues = daoDb.OpenRecordset(sSQL, dbOpenSnapshot, dbReadOnly)
  
  ' Load the control values into an array
  'avValues = rsControlValues.GetRows(100)
  avValues = rsControlValues.GetRows(rsControlValues.RecordCount)

  ' Copy the required values from the 2-dimensional variant array, into
  ' a 1-dimensional string array.
  ReDim asResults(UBound(avValues, 2))
  For iLoop = LBound(avValues, 2) To UBound(avValues, 2)
    asResults(iLoop) = CStr(avValues(0, iLoop))
  Next iLoop

TidyUpAndExit:
  rsControlValues.Close
  Set rsControlValues = Nothing
  ReadColumnControlValues = asResults
  Exit Function
  
ErrorTrap:
  MsgBox "Error reading column control values." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  ReDim asResults(0)
  Resume TidyUpAndExit
  
End Function

Public Property Get AlignToGrid() As Boolean
  ' Return the value of the 'align to grid' property.
  AlignToGrid = mfAlignToGrid
End Property

Public Property Let AlignToGrid(ByVal pfAlignToGrid As Boolean)
  ' Set the value of the 'align to grid' property.
  mfAlignToGrid = pfAlignToGrid
End Property

Private Function CopySelectedControls() As Boolean
  ' Copy the selected controls to the clipboard array.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iControlType As WorkflowWebFormItemTypes
  Dim ctlSourceControl As VB.Control
  Dim ctlCopiedControl As VB.Control
  
  ' Do nothing if no controls are selected.
  If SelectedControlsCount = 0 Then
    CopySelectedControls = True
    Exit Function
  End If
  
  ' Clear the array of copied controls.
  For iIndex = 1 To UBound(gactlClipboardControls)
    Set ctlCopiedControl = gactlClipboardControls(iIndex)
    UnLoad ctlCopiedControl
  Next iIndex
  ReDim gactlClipboardControls(0)
  ' Disassociate object variables.
  Set ctlCopiedControl = Nothing
  
  ' Create a copy of each selected control in the array.
  For Each ctlSourceControl In Me.Controls
    If IsWebFormControl(ctlSourceControl) Then
      If ctlSourceControl.Selected Then
      
        iControlType = WebFormControl_Type(ctlSourceControl)
        
        ' Create a new instance of the required control type.
        Set ctlCopiedControl = AddControl(iControlType)
        
        fOK = Not (ctlCopiedControl Is Nothing)
        
        If fOK Then
          ' Copy the properties from the selected control to the new control.
          fOK = CopyControlProperties(ctlSourceControl, ctlCopiedControl, False)
          
          iIndex = UBound(gactlClipboardControls) + 1
          ReDim Preserve gactlClipboardControls(iIndex)
          Set gactlClipboardControls(iIndex) = ctlCopiedControl
        Else
          Exit For
        End If
        
        Set ctlCopiedControl = Nothing
      
      End If
    End If
  Next ctlSourceControl

TidyUpAndExit:
  If Not fOK Then
    ' Clear the array of copied controls.
    For iIndex = 1 To UBound(gactlClipboardControls)
      Set ctlCopiedControl = gactlClipboardControls(iIndex)
      UnLoad ctlCopiedControl
    Next iIndex
    ReDim gactlClipboardControls(0)
  End If
  ' Disassociate object variables.
  Set ctlSourceControl = Nothing
  Set ctlCopiedControl = Nothing
  CopySelectedControls = fOK
  frmSysMgr.RefreshMenu
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CopyControlProperties(pCtlSource As VB.Control, _
  pCtlDestination As VB.Control, _
  pfPasting As Boolean) As Boolean
  ' Copy the properties from the pCtlSource control to the pCtlDestination control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As WorkflowWebFormItemTypes
  Dim sFileName As String
  Dim sIdentifierRoot As String
  Dim sIdentifier As String
  Dim iMaxSuffix As Integer
  Dim iSuffix As Integer
  Dim fIdentifierContainsIndex As Boolean
  Dim ctlControl As VB.Control
  Dim fRootIdentifierFound As Boolean
  Dim asItemValues() As String

  ' Get the given control's type.
  iControlType = WebFormControl_Type(pCtlSource)

  With pCtlDestination
    ' Copy the source control's properties to the destination control.
    If WebFormItemHasProperty(iControlType, WFITEMPROP_WFIDENTIFIER) Then
      sIdentifier = pCtlSource.WFIdentifier
      
      If pfPasting Then
        sIdentifier = "Copy of " & sIdentifier

        iMaxSuffix = 0
        fRootIdentifierFound = False
              
        For Each ctlControl In Me.Controls
          'The ActiveBar control does mot have the visible property, so to avoid err
          'we only check the visible property of other controls.
          If ctlControl.Name <> "abWebForm" Then
            If ctlControl.Visible Then
              If IsWebFormControl(ctlControl) Then
                If WebFormItemHasProperty(CLng(ctlControl.WFItemType), WFITEMPROP_WFIDENTIFIER) Then
                  If UCase(Left(ctlControl.WFIdentifier, Len(sIdentifier))) = UCase(sIdentifier) Then
                    If UCase(ctlControl.WFIdentifier) = UCase(sIdentifier) Then
                      fRootIdentifierFound = True
                    End If
                  
                    iSuffix = val(Mid(ctlControl.WFIdentifier, Len(sIdentifier) + 1))
                  
                    If iSuffix > iMaxSuffix Then
                      iMaxSuffix = iSuffix
                    End If
                  
                  End If
                End If
              End If
            End If
          End If
        Next ctlControl

        If fRootIdentifierFound Then
          sIdentifier = sIdentifier & " " & CStr(iMaxSuffix + 1)
        End If
      End If
    
      .WFIdentifier = sIdentifier
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
      .WFDefaultCharValue = pCtlSource.WFDefaultCharValue
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_DATE) Then
      .WFDefaultValueDateString = pCtlSource.WFDefaultValueDateString
      .Caption = .WFDefaultValueDateString
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_LOGIC) Then
      .WFDefaultValue = pCtlSource.WFDefaultValue
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
      .WFDefaultNumericValue = pCtlSource.WFDefaultNumericValue
      .Caption = .WFDefaultNumericValue
    End If

    If WebFormItemHasProperty(iControlType, WFITEMPROP_SIZE) Then
      .WFInputSize = pCtlSource.WFInputSize
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DECIMALS) Then
      .WFInputDecimals = pCtlSource.WFInputDecimals
    End If

    If WebFormItemHasProperty(iControlType, WFITEMPROP_ALIGNMENT) Then
      .Alignment = pCtlSource.Alignment
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
      .BackColor = pCtlSource.BackColor
    End If
     
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKSTYLE) Then
      .BackStyle = pCtlSource.BackStyle
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BORDERSTYLE) Then
      .BorderStyle = pCtlSource.BorderStyle
    End If
        
    If (WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION)) Or _
      WebFormControl_HasText(iControlType) Then
      .Caption = pCtlSource.Caption
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
      Set .Font = pCtlSource.Font
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
      .ForeColor = pCtlSource.ForeColor
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOREVEN) Then
      .BackColorEven = pCtlSource.BackColorEven
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORODD) Then
      .BackColorOdd = pCtlSource.BackColorOdd
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
      .BackColorHighlight = pCtlSource.BackColorHighlight
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_COLUMNHEADERS) Then
      .ColumnHeaders = pCtlSource.ColumnHeaders
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADERBACKCOLOR) Then
      .HeaderBackColor = pCtlSource.HeaderBackColor
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOREVEN) Then
      .ForeColorEven = pCtlSource.ForeColorEven
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORODD) Then
      .ForeColorOdd = pCtlSource.ForeColorOdd
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
      .ForeColorHighlight = pCtlSource.ForeColorHighlight
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADFONT) Then
      Set .HeadFont = pCtlSource.HeadFont
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADLINES) Then
      .HeadLines = pCtlSource.HeadLines
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_TABLEID) Then
      .TableID = pCtlSource.TableID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DBRECORD) Then
      .WFDatabaseRecord = pCtlSource.WFDatabaseRecord
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_RECSELTYPE) Then
      .WFDatabaseRecord = pCtlSource.WFDatabaseRecord
    End If
        
    If WebFormItemHasProperty(iControlType, WFITEMPROP_ELEMENTIDENTIFIER) _
      Or (iControlType = giWFFORMITEM_WFVALUE) _
      Or (iControlType = giWFFORMITEM_WFFILE) Then
      
      .WFWorkflowForm = pCtlSource.WFWorkflowForm
      .WFWorkflowValue = pCtlSource.WFWorkflowValue
    End If
      
    If WebFormItemHasProperty(iControlType, WFITEMPROP_PICTURE) Then
      .PictureID = pCtlSource.PictureID
      If .PictureID > 0 Then
        recPictEdit.Index = "idxID"
        recPictEdit.Seek "=", .PictureID
                    
        If Not recPictEdit.NoMatch Then
          sFileName = ReadPicture
          .Picture = sFileName
          Kill sFileName
        End If
      Else
        .Picture = "No picture"
      End If
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_CONTROLVALUELIST) Then
      If iControlType = giWFFORMITEM_INPUTVALUE_DROPDOWN Then
        'Dropdown List
        .ControlValueList = pCtlSource.ControlValueList
      Else
        'Option Group
        asItemValues = Split(pCtlSource.ControlValueList, vbTab)
        .NoOptions = (Len(pCtlSource.ControlValueList) = 0)
        .SetOptions asItemValues
      End If
    End If

    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_LOOKUP) Or _
      WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_LIST) Then
      
      .DefaultStringValue = pCtlSource.DefaultStringValue
      
      Select Case iControlType
        Case giWFFORMITEM_INPUTVALUE_DROPDOWN, giWFFORMITEM_INPUTVALUE_LOOKUP
          'Dropdown List / Lookup
          .Caption = .DefaultStringValue
        
        Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
          'Option Group
          .SelectOption (.DefaultStringValue)
          
      End Select
      
    End If

    If WebFormItemHasProperty(iControlType, WFITEMPROP_LOOKUPTABLEID) Then
      .LookupTableID = pCtlSource.LookupTableID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_LOOKUPCOLUMNID) Then
      .LookupColumnID = pCtlSource.LookupColumnID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_SUBMITTYPE) Then
      .Behaviour = pCtlSource.Behaviour
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_MANDATORY) Then
      .Mandatory = pCtlSource.Mandatory
    End If
            
    If WebFormItemHasProperty(iControlType, WFITEMPROP_RECORDTABLEID) Then
      .WFRecordTableID = pCtlSource.WFRecordTableID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_ORIENTATION) Then
      .Alignment = pCtlSource.Alignment
    End If

    If WebFormItemHasProperty(iControlType, WFITEMPROP_RECORDORDER) Then
      .WFRecordOrderID = pCtlSource.WFRecordOrderID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_RECORDFILTER) Then
      .WFRecordFilterID = pCtlSource.WFRecordFilterID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_CALCULATION) _
      Or WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
      .CalculationID = pCtlSource.CalculationID
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTIONTYPE) Then
      .CaptionType = pCtlSource.CaptionType
    End If
            
    If WebFormItemHasProperty(iControlType, WFITEMPROP_DEFAULTVALUETYPE) Then
      .DefaultValueType = pCtlSource.DefaultValueType
    End If
                
    If WebFormItemHasProperty(iControlType, WFITEMPROP_VERTICALOFFSET) Then
      .VerticalOffsetBehaviour = pCtlSource.VerticalOffsetBehaviour
      .VerticalOffset = pCtlSource.VerticalOffset
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_HORIZONTALOFFSET) Then
      .HorizontalOffsetBehaviour = pCtlSource.HorizontalOffsetBehaviour
      .HorizontalOffset = pCtlSource.HorizontalOffset
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
      .HeightBehaviour = pCtlSource.HeightBehaviour
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_WIDTHBEHAVIOUR) Then
      .WidthBehaviour = pCtlSource.WidthBehaviour
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_PASSWORDTYPE) Then
      .PasswordType = pCtlSource.PasswordType
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_FILEEXTENSIONS) Then
      .WFFileExtensions = pCtlSource.WFFileExtensions
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_LOOKUPFILTERCOLUMN) Then
      .LookupFilterColumn = pCtlSource.LookupFilterColumn
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_LOOKUPFILTEROPERATOR) Then
      .LookupFilterOperator = pCtlSource.LookupFilterOperator
    End If
    
    If WebFormItemHasProperty(iControlType, WFITEMPROP_LOOKUPFILTERVALUE) Then
      .LookupFilterValue = pCtlSource.LookupFilterValue
    End If
    
    ' Copy the source control's position and dimension's to the destination control.
    .ColumnID = pCtlSource.ColumnID
    .Top = pCtlSource.Top
    .Left = pCtlSource.Left
    .Height = pCtlSource.Height
    .Width = pCtlSource.Width
  
    .ToolTipText = pCtlSource.ToolTipText

    ' Force the value of some of the destination control's properties.
    .Selected = False
    .Visible = False
    
  End With
  
  fOK = True
  
TidyUpAndExit:
  CopyControlProperties = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function PasteControls() As Boolean

  ' Paste the controls from the clipboard onto the current page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim iControlType As WorkflowWebFormItemTypes
  Dim lngXOffset As Long
  Dim lngYOffset As Long
  Dim ctlControl As VB.Control
  Dim ctlNewControl As VB.Control
  Dim VarPageContainer As Variant
  
  ' Do nothing if there's nothing in the clipboard.
  If ClipboardControlsCount = 0 Then
    PasteControls = True
    Exit Function
  End If
  
 ' Open a progress bar
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = ClipboardControlsCount
    .Bar1Caption = "Pasting Controls..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  ' Lock the forms refreshing.
  UI.LockWindow Me.hWnd
  
  ' Get the current page container.
  Set VarPageContainer = CurrentContainer
  
  ' Get the offset for the new positions of the controls.
  lngXOffset = VarPageContainer.Width
  lngYOffset = VarPageContainer.Height
  
  For iIndex = 1 To UBound(gactlClipboardControls)
    Set ctlControl = gactlClipboardControls(iIndex)
    With ctlControl
      If .Left < lngXOffset Then
        lngXOffset = .Left
      End If
      If .Top < lngYOffset Then
        lngYOffset = .Top
      End If
    End With
  Next iIndex
  
  Set ctlControl = Nothing
  
  ' Deselect all existing controls.
  fOK = DeselectAllControls
  
  If fOK Then
  
    ReDim gavUndo_PastedControls(2, 0)
  
    ' Drop each control from the clipboard onto the current page.
    For iIndex = 1 To UBound(gactlClipboardControls)
    
      Set ctlControl = gactlClipboardControls(iIndex)
      
      ' Add the required control type.
      iControlType = WebFormControl_Type(ctlControl)
      Set ctlNewControl = AddControl(iControlType)
      
      fOK = Not (ctlNewControl Is Nothing)
      If fOK Then
      
        fOK = CopyControlProperties(ctlControl, ctlNewControl, True)
    
        If fOK Then
          With ctlNewControl
            Set .Container = VarPageContainer
            .Left = AlignX(.Left - lngXOffset)
            .Top = AlignY(.Top - lngYOffset)
            
            ctlNewControl.Selected = True
            fOK = SelectControl(ctlNewControl)
            
            If fOK Then
              .Visible = True
            
              iIndex2 = UBound(gavUndo_PastedControls, 2) + 1
              ReDim Preserve gavUndo_PastedControls(2, iIndex2)
              gavUndo_PastedControls(1, iIndex2) = .Name
              gavUndo_PastedControls(2, iIndex2) = .Index
            End If
          End With
        End If
      End If
      
      If Not fOK Then
        Exit For
      End If
      
      'Update the progress bar
      gobjProgress.UpdateProgress
    Next iIndex
  End If

  If fOK Then
    ' Mark the screen as having changed.
    mfChanged = True

    ' Set the last action flag and enable the Undo menu option.
    giLastActionFlag = giACTION_PASTECONTROLS
    frmSysMgr.RefreshMenu
  
    ' Refresh the properties screen.
    Set frmWorkflowWFItemProps.CurrentWebForm = Me
    frmWorkflowWFItemProps.RefreshProperties
  End If

TidyUpAndExit:
  ' Unlock the forms refreshing.
  UI.UnlockWindow
  
  'Close the progress bar
  gobjProgress.CloseProgress
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set VarPageContainer = Nothing
  PasteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function DeleteControl(pctlControl As VB.Control) As Boolean
  ' Delete the given screen control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  
  fOK = True
  
  ' Get the index of the given control.
  iIndex = val(pctlControl.Tag)
  
  ' Do not delete the control array dummy (index = 0).
  If pctlControl.Index = 0 Then
    DeleteControl = True
    Exit Function
  End If
  
  ' Hide the selection markers
  If Not pctlControl.Tag = "" Then
    ASRSelectionMarkers(pctlControl.Tag).Visible = False
  End If
  
  ' Unload the screen control.
  UnLoad pctlControl
  
  If iIndex > 0 Then
    ' Unload the control's selection markers.
    fOK = True
  End If
        
TidyUpAndExit:
  DeleteControl = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AddControl(piElementType As WorkflowWebFormItemTypes) As VB.Control
  
  On Error GoTo ErrorTrap

  Select Case piElementType
    Case giWFFORMITEM_BUTTON
      Load btnWorkflow(btnWorkflow.UBound + 1)
      Set AddControl = btnWorkflow(btnWorkflow.UBound)
      With AddControl
        .BackColor = vbButtonFace
        .ForeColor = vbButtonText
      End With
      
    Case giWFFORMITEM_DBVALUE, _
      giWFFORMITEM_DBFILE
      
      Load asrDummyLabel(asrDummyLabel.UBound + 1)
      Set AddControl = asrDummyLabel(asrDummyLabel.UBound)
      With AddControl
        .BorderStyle = vbBSNone
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
        .ForeColor = vbWindowText
      End With
      
    Case giWFFORMITEM_INPUTVALUE_CHAR
      Load asrDummyTextBox(asrDummyTextBox.UBound + 1)
      Set AddControl = asrDummyTextBox(asrDummyTextBox.UBound)
      With AddControl
        .ForeColor = vbWindowText
      End With
        
    Case giWFFORMITEM_INPUTVALUE_DATE
      Load asrDummyCombo(asrDummyCombo.UBound + 1)
      Set AddControl = asrDummyCombo(asrDummyCombo.UBound)
      With AddControl
        .ForeColor = vbWindowText
      End With
    
    Case giWFFORMITEM_INPUTVALUE_DROPDOWN
      Load asrDummyCombo(asrDummyCombo.UBound + 1)
      Set AddControl = asrDummyCombo(asrDummyCombo.UBound)
      With AddControl
        .ForeColor = vbWindowText
      End With

    Case giWFFORMITEM_INPUTVALUE_LOGIC
      Load asrDummyCheckBox(asrDummyCheckBox.UBound + 1)
      Set AddControl = asrDummyCheckBox(asrDummyCheckBox.UBound)
      With AddControl
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
        .ForeColor = vbWindowText
      End With
    
    Case giWFFORMITEM_INPUTVALUE_LOOKUP
      Load asrDummyCombo(asrDummyCombo.UBound + 1)
      Set AddControl = asrDummyCombo(asrDummyCombo.UBound)
      With AddControl
        .ForeColor = vbWindowText
      End With

    Case giWFFORMITEM_INPUTVALUE_NUMERIC
      Load asrDummyTextBox(asrDummyTextBox.UBound + 1)
      Set AddControl = asrDummyTextBox(asrDummyTextBox.UBound)
      With AddControl
        .Alignment = vbRightJustify
        .Caption = " " & .WFDefaultNumericValue
        .ForeColor = vbWindowText
      End With
    
    Case giWFFORMITEM_INPUTVALUE_GRID
      Load ASRDummyGrid(ASRDummyGrid.UBound + 1)
      Set AddControl = ASRDummyGrid(ASRDummyGrid.UBound)
      With AddControl
        .TableID = mlngBaseTableID
        .WFDatabaseRecord = giWFRECSEL_ALL
        .ForeColor = vbWindowText
      End With
      
    Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      Load ASRDummyOptions(ASRDummyOptions.UBound + 1)
      Set AddControl = ASRDummyOptions(ASRDummyOptions.UBound)
      With AddControl
        .ClearOptions
        .NoOptions = True
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
        .ForeColor = vbWindowText
      End With
      
    Case giWFFORMITEM_LABEL
      Load asrDummyLabel(asrDummyLabel.UBound + 1)
      Set AddControl = asrDummyLabel(asrDummyLabel.UBound)
      With AddControl
        .BorderStyle = vbBSNone
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
        .ForeColor = vbWindowText
      End With
      
    Case giWFFORMITEM_WFVALUE, _
      giWFFORMITEM_WFFILE
      
      Load asrDummyLabel(asrDummyLabel.UBound + 1)
      Set AddControl = asrDummyLabel(asrDummyLabel.UBound)
      With AddControl
        .BorderStyle = vbBSNone
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
        .ForeColor = vbWindowText
      End With
    
    Case giWFFORMITEM_FRAME
      Load asrDummyFrame(asrDummyFrame.UBound + 1)
      Set AddControl = asrDummyFrame(asrDummyFrame.UBound)
      With AddControl
        .BackColor = Me.BackColor
        .BackStyle = BACKSTYLE_TRANSPARENT
      End With
    
    Case giWFFORMITEM_LINE
      Load ASRDummyLine(ASRDummyLine.UBound + 1)
      Set AddControl = ASRDummyLine(ASRDummyLine.UBound)
      
    Case giWFFORMITEM_IMAGE
      Load asrDummyImage(asrDummyImage.UBound + 1)
      Set AddControl = asrDummyImage(asrDummyImage.UBound)
    
    Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
      Load ASRDummyFileUpload(ASRDummyFileUpload.UBound + 1)
      Set AddControl = ASRDummyFileUpload(ASRDummyFileUpload.UBound)
      With AddControl
        .BackColor = vbButtonFace
        .ForeColor = vbButtonText
      End With
    
    Case Else
      Load asrDummyLabel(asrDummyLabel.UBound + 1)
      Set AddControl = asrDummyLabel(asrDummyLabel.UBound)
      With AddControl
        .BorderStyle = vbBSNone
        .BackColor = Me.BackColor
      End With
      
  End Select

  AddControl.WFItemType = piElementType

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Unable to load control type " & Trim(Str(piElementType)) & "." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, Application.Name
  Set AddControl = Nothing
  Resume TidyUpAndExit

End Function


Private Function CurrentContainer() As Variant
  ' Return the current page container.
  Set CurrentContainer = Me
End Function

Private Function CutSelectedControls() As Boolean
  ' Cut the selected controls.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Copy the selected controls to the clipboard.
  fOK = CopySelectedControls
  
  If fOK Then
    ' Delete the selected controls.
    fOK = DeleteSelectedControls(True)
  End If
  
  If fOK Then
    ' Set the last action flag and enable the Undo menu option.
    giLastActionFlag = giACTION_CUTCONTROLS
    frmSysMgr.RefreshMenu
  End If

TidyUpAndExit:
  CutSelectedControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AlignX(pLngX As Long) As Long
  ' Return the given X coordinate aligned to the X grid if required.
  If mfAlignToGrid Then
    AlignX = pLngX - (pLngX Mod giGridX)
  Else
    AlignX = pLngX
  End If
End Function
Private Function AlignY(pLngY As Long) As Long
  ' Return the given Y coordinate aligned to the Y grid if required.
  If mfAlignToGrid Then
    AlignY = pLngY - (pLngY Mod giGridY)
  Else
    AlignY = pLngY
  End If
End Function

Private Function SaveWebForm() As Boolean

  ' Save the web form to the current workflow element object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  Screen.MousePointer = vbHourglass
  
  fOK = ValidateWebForm

  If fOK Then
    With gobjProgress
      .Caption = "Web Form Designer"
      .Bar1Value = 0
      .Bar1MaxValue = 100
      .Bar1Caption = "Saving Web Form Design..."
      .Cancel = False
      .Time = False
      .OpenProgress
    End With
  
    fOK = GetControlLevel(Me.hWnd)
  
    If fOK Then
      mavIdentifierLog(3, 0) = Me.WFIdentifier
      mavIdentifierLog(4, 0) = False
      
      ' Save the Element Properties (Identifier, Background Colour/Image)
      fOK = SaveWebFormProperties(mwfElement)
    End If
  End If
  
ExitSaveWebForm:
  If fOK Then
    
    mfChanged = False
    Application.Changed = True
    mfrmCallingForm.IsChanged = True
    frmSysMgr.RefreshMenu
  End If
  
  gobjProgress.CloseProgress
  Screen.MousePointer = vbNormal
  SaveWebForm = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitSaveWebForm
  
End Function



Private Function SaveWebFormProperties(pwfElement As COAWF_Webform) As Boolean

  ' Save the web form to the current workflow element object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  fOK = True

  ' Save the Element Properties (Identifier, Background Colour/Image)
  pwfElement.Identifier = Me.WFIdentifier
  pwfElement.Caption = Me.Caption
  pwfElement.WebFormFGColor = Me.ForeColor
  pwfElement.WebFormBGColor = Me.BackColor
  pwfElement.WebFormBGImageID = Me.PictureID
  pwfElement.WebFormBGImageLocation = Me.PictureLocation
  ' AE20080509 Fault #13161
  'pwfElement.WebFormDefaultFont = Me.Font
  Set pwfElement.WebFormDefaultFont = Me.Font
  pwfElement.WebFormWidth = TwipsToPixels(Me.ScaleWidth)
  pwfElement.WebFormHeight = TwipsToPixels(Me.ScaleHeight)

  pwfElement.WebFormTimeoutFrequency = Me.TimeoutFrequency
  pwfElement.WebFormTimeoutPeriod = Me.TimeoutPeriod
  pwfElement.WebFormTimeoutExcludeWeekend = Me.TimeoutExcludeWeekend
  pwfElement.DescriptionExprID = Me.DescriptionExprID
  pwfElement.DescriptionHasWorkflowName = Me.DescriptionHasWorkflowName
  pwfElement.DescriptionHasElementCaption = Me.DescriptionHasElementCaption

  pwfElement.WFCompletionMessageType = Me.WFCompletionMessageType
  pwfElement.WFCompletionMessage = Me.WFCompletionMessage
  pwfElement.WFSavedForLaterMessageType = Me.WFSavedForLaterMessageType
  pwfElement.WFSavedForLaterMessage = Me.WFSavedForLaterMessage
  pwfElement.WFFollowOnFormsMessageType = Me.WFFollowOnFormsMessageType
  pwfElement.WFFollowOnFormsMessage = Me.WFFollowOnFormsMessage
  
  pwfElement.Validations = Me.Validations

  fOK = SaveWebFormItems(pwfElement)

ExitSaveWebFormProperties:
  SaveWebFormProperties = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitSaveWebFormProperties
  
End Function

Public Property Let WFIdentifier(NEW_VALUE As String)
  msWFIdentifier = NEW_VALUE
End Property

Public Property Get WFIdentifier() As String
  WFIdentifier = msWFIdentifier
End Property

Private Function SelectAllControls() As Boolean

  ' Select all controls on the web form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ctlControl As VB.Control
  Dim VarPageContainer As Variant
  
  fOK = True
  
  ' Put up an hourglass
  Screen.MousePointer = vbHourglass
   
  ' Get the current page container.
  Set VarPageContainer = CurrentContainer
  
  ' Select all the controls on this container
  For Each ctlControl In Me.Controls
    'The ActiveBar control does mot have the visible property, so to avoid err
    'we only check the visible property of other controls.
    If ctlControl.Name <> "abWebForm" Then
      If ctlControl.Visible Then
        If IsWebFormControl(ctlControl) Then
          If ctlControl.Container Is VarPageContainer Then
            ctlControl.Selected = True
            SelectControl ctlControl
          End If
        End If
      End If
    End If
  Next ctlControl
  
  'Refresh the menu
  frmSysMgr.RefreshMenu
  
TidyUpAndExit:

  'Reset the mousepointer
  Screen.MousePointer = vbNormal
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SelectAllControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoLastAction() As Boolean
  ' Undo the last action.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  
  Select Case giLastActionFlag
    ' Undo the previous control Drop.
    Case giACTION_DROPCONTROL, giACTION_DROPCONTROLAUTOLABEL
      If Not UndoDropControl Then
        MsgBox "Unable to undo Drop Control." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Cut.
    Case giACTION_CUTCONTROLS
      If Not UndoCutControls Then
        MsgBox "Unable to undo Cut Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Paste.
    Case giACTION_PASTECONTROLS
      If Not UndoPasteControls Then
        MsgBox "Unable to undo Paste Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Delete.
    Case giACTION_DELETECONTROLS
      If Not UndoDeleteControls Then
        MsgBox "Unable to undo Delete Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
  End Select
  
  ' Clear the last action flag.
  giLastActionFlag = giACTION_NOACTION
  
  ' Disable the Undo button on the menubar.
  frmSysMgr.RefreshMenu
  
  ' Refresh the properties screen.
  Set frmWorkflowWFItemProps.CurrentWebForm = Me
  frmWorkflowWFItemProps.RefreshProperties

  fOK = True
  
TidyUpAndExit:
  UndoLastAction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function AutoSizeControl(pctlControl As VB.Control) As Boolean
  ' Initialise the given control's properties.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iType As Long
  Dim iDigits As Integer
  Dim iMinLength As Integer
  Dim iMaxLength As Integer
  Dim lngColumnID As Long
  Dim lngMinWidth As Long
  Dim lngMinHeight As Long
  Dim iExtraWidth As Integer
  Dim sngWidth As Single
  Dim iLoop As Integer
  Dim fLiteral As Boolean
  Dim sMask As String

  lngColumnID = pctlControl.ColumnID
  iType = WebFormControl_Type(pctlControl)
    
  ' If we are initialising a column control.
  If lngColumnID >= 0 Then
      
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", lngColumnID
          
      If Not .NoMatch Then
            
        ' Set the width of the new control.
        Select Case iType
          Case giWFFORMITEM_DBVALUE
            ' If the column has size, then set the control
            ' width to column size * average character width.
            If .Fields("size") > 0 Then
              If .Fields("datatype") = dtVARCHAR Then
                If .Fields("Multiline") Then
                  pctlControl.Width = Me.TextWidth(String(8000, "W")) + (2 * XFrame)
                Else
                  If Len(.Fields("Mask")) > 0 Then
                    fLiteral = False
                    sMask = .Fields("Mask")
                    
                    For iLoop = 1 To Len(sMask)
                      If fLiteral Then
                        sngWidth = sngWidth + Me.TextWidth(String(1, Mid(sMask, iLoop, 1)))
                        fLiteral = False
                      Else
                        Select Case Mid(sMask, iLoop, 1)
                          Case "A"
                            sngWidth = sngWidth + Me.TextWidth(String(1, "W"))
                          Case "a"
                            sngWidth = sngWidth + Me.TextWidth(String(1, "w"))
                          Case "9"
                            sngWidth = sngWidth + Me.TextWidth(String(1, "8"))
                          Case "#"
                            sngWidth = sngWidth + Me.TextWidth(String(1, "8"))
                          Case "B"
                            sngWidth = sngWidth + Me.TextWidth(String(1, "0"))
                          Case "\"
                            fLiteral = True
                          Case Else
                            sngWidth = sngWidth + Me.TextWidth(String(1, Mid(sMask, iLoop, 1)))
                        End Select
                      End If
                    Next iLoop
                  
                    pctlControl.Width = sngWidth + (2 * XFrame)
                  Else
                    pctlControl.Width = Default_ColumnWidth_Textbox(.Fields("size").value)
                  End If
                End If
              Else
                pctlControl.Width = Default_ColumnWidth_Numeric(.Fields("size").value, .Fields("decimals").value, .Fields("Use1000Separator").value)
              End If
            End If
        End Select
      End If
    End With
  End If
            
  Select Case iType
    ' Set the control to have the minimum width and height for labels.
    Case giWFFORMITEM_LABEL
      lngMinWidth = Me.TextWidth(pctlControl.Caption)
      lngMinWidth = IIf(lngMinWidth < 255, 255, lngMinWidth)
      pctlControl.Width = lngMinWidth
      lngMinHeight = Me.TextHeight(pctlControl.Caption)
      lngMinHeight = IIf(lngMinHeight < 195, 195, lngMinHeight)
      pctlControl.Height = lngMinHeight
                
    ' Set the control to have the minimum height for textboxes.
    ' Do not set width.
    Case giWFFORMITEM_INPUTVALUE_CHAR
      pctlControl.Height = pctlControl.MinimumHeight
                
    ' Set the control to have the minimum width and height for check boxes.
    Case giWFFORMITEM_INPUTVALUE_LOGIC
      'lngMinWidth = 360 + Me.TextWidth("W" & pctlControl.Caption)
      'pctlControl.Width = lngMinWidth
      'lngMinHeight = UI.GetCharHeight(Me.hDC)
      'If lngMinHeight < 285 Then lngMinHeight = 285
      pctlControl.Height = pctlControl.MinimumHeight
    
    Case giWFFORMITEM_LINE
      pctlControl.Length = 1000
      pctlControl.Height = 1000
      
    Case giWFFORMITEM_INPUTVALUE_GRID
      pctlControl.Width = IIf((Me.ScaleWidth - pctlControl.Left - 500) < 3000, 3000, (Me.ScaleWidth - pctlControl.Left - 500))
      pctlControl.Height = IIf((Me.ScaleHeight - pctlControl.Top - 500) < 2000, 2000, (Me.ScaleHeight - pctlControl.Top - 500))
      
  End Select
          
  ' Ensure the control does not extend past the right-hand edge
  ' of the parent container.
  With pctlControl
    If .Left + .Width > .Container.Width Then
      .Width = .Container.Width - .Left
    End If
  End With
  
  fOK = True
  
TidyUpAndExit:
  AutoSizeControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get UndoAction() As UndoActionFlags
  ' Return the key that identifies the alast action that can be 'undone'.
  UndoAction = giLastActionFlag
End Property


Private Function UndoDropControl() As Boolean
  ' Delete the last control that was dropped on the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL Then
    fOK = DeleteControl(asrDummyLabel(giUndo_ControlAutoLabelIndex))
  End If
  
  Select Case gsUndo_ControlType
    Case "asrDummyLabel"
      fOK = DeleteControl(asrDummyLabel(giUndo_ControlIndex))
    Case "asrDummyTextBox"
      fOK = DeleteControl(asrDummyTextBox(giUndo_ControlIndex))
    Case "asrDummyImage"
      fOK = DeleteControl(asrDummyImage(giUndo_ControlIndex))
    Case "asrDummyFrame"
      fOK = DeleteControl(asrDummyFrame(giUndo_ControlIndex))
    Case "asrDummyCombo"
      fOK = DeleteControl(asrDummyCombo(giUndo_ControlIndex))
    Case "asrDummyCheckBox"
      fOK = DeleteControl(asrDummyCheckBox(giUndo_ControlIndex))
    Case "ASRDummyGrid"
      fOK = DeleteControl(ASRDummyGrid(giUndo_ControlIndex))
    Case "ASRDummyLine"
      fOK = DeleteControl(ASRDummyLine(giUndo_ControlIndex))
    Case "btnWorkflow"
      fOK = DeleteControl(btnWorkflow(giUndo_ControlIndex))
    Case "ASRDummyOptions"
      fOK = DeleteControl(ASRDummyOptions(giUndo_ControlIndex))
    Case "ASRDummyFileUpload"
      fOK = DeleteControl(ASRDummyFileUpload(giUndo_ControlIndex))
  End Select

TidyUpAndExit:
  UndoDropControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoPasteControls() As Boolean
  ' Delete the last controls that were pasted on the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  
  fOK = True
  
  ' Delete the pasted controls.
  For iIndex = 1 To UBound(gavUndo_PastedControls, 2)
  
    iIndex2 = gavUndo_PastedControls(2, iIndex)
    
    Select Case gavUndo_PastedControls(1, iIndex)
      Case "asrDummyLabel"
        fOK = DeleteControl(asrDummyLabel(iIndex2))
      Case "asrDummyTextBox"
        fOK = DeleteControl(asrDummyTextBox(iIndex2))
      Case "asrDummyImage"
        fOK = DeleteControl(asrDummyImage(iIndex2))
      Case "asrDummyFrame"
        fOK = DeleteControl(asrDummyFrame(iIndex2))
      Case "asrDummyCombo"
        fOK = DeleteControl(asrDummyCombo(iIndex2))
      Case "asrDummyCheckBox"
        fOK = DeleteControl(asrDummyCheckBox(iIndex2))
      Case "asrDummyGrid"
        fOK = DeleteControl(ASRDummyGrid(iIndex2))
      Case "ASRDummyLine"
        fOK = DeleteControl(ASRDummyLine(iIndex2))
      Case "btnWorkflow"
        fOK = DeleteControl(btnWorkflow(iIndex2))
      Case "asrDummyFileUpload"
        fOK = DeleteControl(ASRDummyFileUpload(iIndex2))
    End Select
  Next iIndex
  
TidyUpAndExit:
  UndoPasteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoCutControls() As Boolean
  ' Paste the cut controls back onto their original page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = UndoDeleteControls

TidyUpAndExit:
  UndoCutControls = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function UndoDeleteControls() As Boolean

  ' Paste the deleted controls onto their original page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim ctlNewControl As VB.Control

 ' Open a progress bar
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = UBound(gactlUndo_DeletedControls)
    .Bar1Caption = "Undoing Screen Control Deletion..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  ' Restore the deleted controls to their original positions.
  For iIndex = 1 To UBound(gactlUndo_DeletedControls)

    Set ctlNewControl = gactlUndo_DeletedControls(iIndex)
    ctlNewControl.Visible = True
    fOK = SelectControl(ctlNewControl)
    
    ' Disassociate object variables.
'    Set ctlNewControl = Nothing
    
    Set gactlUndo_DeletedControls(iIndex) = Nothing
  
    If Not fOK Then
      Exit For
    End If

    'Update the progress bar
    gobjProgress.UpdateProgress

  Next iIndex

  ' Clear the array of deleted controls.
  ReDim gactlUndo_DeletedControls(0)

TidyUpAndExit:

  'Close the progress bar
  gobjProgress.CloseProgress

  UndoDeleteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function GetControlLevel(pLngHWnd As Long) As Boolean
  ' Determine the control level of each screen control. Set the 'controlLevel' property
  ' of the screen controls with the determined value.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iCounter As Integer
  Dim lngChildHWnd As Long
  Dim actlWebFormControls() As VB.Control
  Dim ctlControl As VB.Control
  
  ' Create an array of the screen control's.
  ReDim actlWebFormControls(0)
    
  ' Construct an array of the screen controls.
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      iIndex = UBound(actlWebFormControls) + 1
      ReDim Preserve actlWebFormControls(iIndex)
      Set actlWebFormControls(iIndex) = ctlControl
    End If
  Next ctlControl
    
  ' Disassociate object variables.
  Set ctlControl = Nothing
  
  iCounter = 1
  
  ' Get the hWnd of the first child window of the given page.
  lngChildHWnd = UI.GetChildWindowHWnd(pLngHWnd, GW_CHILD)
    
  ' Find all the child windows of the screen designer.
  Do While lngChildHWnd <> 0
    ' Check if the child window is a screen control.
    For iLoop = 1 To UBound(actlWebFormControls)
      Set ctlControl = actlWebFormControls(iLoop)
      If lngChildHWnd = ctlControl.hWnd Then
        ctlControl.ControlLevel = iCounter
        iCounter = iCounter + 1
        Exit For
      End If
      Set ctlControl = Nothing
    Next iLoop
    
    ' Get the hWnd of the next child window of the screen designer.
    lngChildHWnd = UI.GetChildWindowHWnd(lngChildHWnd, GW_HWNDNEXT)
  Loop

  fOK = True
  
TidyUpAndExit:
  GetControlLevel = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function SetControlLevel() As Boolean
  ' Set the correct z-order for each control.
  ' The controlLevel property of each control will determine the z-order of each control, but
  ' we now need to actually set that z-order value.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLevel As Integer
  Dim iMaxLevel As Integer
  Dim ctlControl As VB.Control
  
  ' Initialise the array of control information.
  iMaxLevel = 0
  
  ' Find the highest control level.
  For Each ctlControl In Me.Controls
    With ctlControl
      If IsWebFormControl(ctlControl) Then
        If ctlControl.ControlLevel > iMaxLevel Then iMaxLevel = ctlControl.ControlLevel
      End If
    End With
  Next ctlControl
  ' Disassociate object variables.
  Set ctlControl = Nothing
  
  ' Set the z-order for each control.
  For iLevel = iMaxLevel To 0 Step -1
    For Each ctlControl In Me.Controls
      If IsWebFormControl(ctlControl) Then
        If ctlControl.ControlLevel = iLevel Then
          ctlControl.ZOrder 0
        End If
      End If
    Next ctlControl
  Next iLevel
  
  fOK = True
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SetControlLevel = fOK
  Exit Function
  
ErrorTrap:
  MsgBox "Error setting control level." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function IsWebFormControl(pctlControl As VB.Control) As Boolean
  ' Return true if the given control is a screen control.
  On Error GoTo ErrorTrap
  
  Dim fIsWebFormControl As Boolean
  Dim iIndex As Integer
  Dim sName As String
  
  sName = pctlControl.Name
  fIsWebFormControl = False
  
  If sName = "asrDummyLabel" Or _
    sName = "asrDummyTextBox" Or _
    sName = "asrDummyPhoto" Or _
    sName = "asrDummyOLEContents" Or _
    sName = "asrDummyImage" Or _
    sName = "asrDummyFrame" Or _
    sName = "asrDummyCombo" Or _
    sName = "asrDummySpinner" Or _
    sName = "asrDummyCheckBox" Or _
    sName = "asrDummyLink" Or _
    sName = "ASRCustomDummyWP" Or _
    sName = "ASRDummyLine" Or _
    sName = "ASRDummyOptions" Or _
    sName = "ASRDummyGrid" Or _
    sName = "btnWorkflow" Or _
    sName = "ASRDummyFileUpload" Then
    
    ' Do not bother with the dummy screen controls.
    If (pctlControl.Index > 0) Then
    
      fIsWebFormControl = True
      
      ' Do not bother with controls in the deleted array.
      For iIndex = 1 To UBound(gactlUndo_DeletedControls)
        If pctlControl Is gactlUndo_DeletedControls(iIndex) Then
          fIsWebFormControl = False
          Exit For
        End If
      Next iIndex
    
      If fIsWebFormControl Then
        ' Do not bother with controls in the clipboard array.
        For iIndex = 1 To UBound(gactlClipboardControls)
          If pctlControl Is gactlClipboardControls(iIndex) Then
            fIsWebFormControl = False
            Exit For
          End If
        Next iIndex
      End If
      
    End If
  End If
  
TidyUpAndExit:
  IsWebFormControl = fIsWebFormControl
  Exit Function

ErrorTrap:
  fIsWebFormControl = False
  Resume TidyUpAndExit
  
End Function

Public Function WebFormControl_IsTabStop(piControlType As WorkflowWebFormItemTypes) As Boolean
  ' Return true if the given control has a Caption property.
  WebFormControl_IsTabStop = _
    (piControlType = giWFFORMITEM_BUTTON) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_CHAR) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_DATE) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_LOGIC) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_NUMERIC) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_GRID) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_DROPDOWN) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_LOOKUP) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) Or _
    (piControlType = giWFFORMITEM_DBFILE) Or _
    (piControlType = giWFFORMITEM_WFFILE)

End Function

Public Function WebFormControl_HasAutoLabel(piControlType As WorkflowWebFormItemTypes) As Boolean
  ' Return true if the given control gets a label created if auto-label is enabled.
  WebFormControl_HasAutoLabel = _
    (piControlType = giWFFORMITEM_INPUTVALUE_CHAR) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_DATE) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_DROPDOWN) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_LOOKUP) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_NUMERIC) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_GRID)

End Function

Public Function WebFormControl_HasText(piControlType As WorkflowWebFormItemTypes) As Boolean
  ' Return true if the given control has a Caption property.
  WebFormControl_HasText = _
    (piControlType = giWFFORMITEM_BUTTON) Or _
    (piControlType = giWFFORMITEM_DBVALUE) Or _
    (piControlType = giWFFORMITEM_FRAME) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_CHAR) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_LOGIC) Or _
    (piControlType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Or _
    (piControlType = giWFFORMITEM_LABEL) Or _
    (piControlType = giWFFORMITEM_WFVALUE) Or _
    (piControlType = giWFFORMITEM_DBFILE) Or _
    (piControlType = giWFFORMITEM_WFFILE)
End Function


Public Function WebFormControl_Type(pctlControl As VB.Control) As WorkflowWebFormItemTypes

  ' Return the control type of the given control.
  Select Case pctlControl.WFItemType
    Case 0
      WebFormControl_Type = giWFFORMITEM_BUTTON
    Case 1
      WebFormControl_Type = giWFFORMITEM_DBVALUE
    Case 2
      WebFormControl_Type = giWFFORMITEM_LABEL
    Case 3
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_CHAR
    Case 4
      WebFormControl_Type = giWFFORMITEM_WFVALUE
    Case 5
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_NUMERIC
    Case 6
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_LOGIC
    Case 7
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_DATE
    Case 8
      WebFormControl_Type = giWFFORMITEM_FRAME
    Case 9
      WebFormControl_Type = giWFFORMITEM_LINE
    Case 10
      WebFormControl_Type = giWFFORMITEM_IMAGE
    Case 11
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_GRID
    Case 12
      WebFormControl_Type = giWFFORMITEM_FORMATCODE
    Case 13
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_DROPDOWN
    Case 14
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_LOOKUP
    Case 15
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_OPTIONGROUP
    Case 17
      WebFormControl_Type = giWFFORMITEM_INPUTVALUE_FILEUPLOAD
    Case 19
      WebFormControl_Type = giWFFORMITEM_DBFILE
    Case 20
      WebFormControl_Type = giWFFORMITEM_WFFILE
    Case Else
      WebFormControl_Type = giWFFORMITEM_UNKNOWN
  End Select
  
End Function

Private Function SaveWebFormItems(pwfElement As COAWF_Webform) As Boolean

  ' Save the definition of each instance of each type of screen control to the database.
  On Error GoTo ErrorTrap
  
  Dim fSaveOK As Boolean
  Dim ctlControl As VB.Control
  Dim asItems() As String
  Dim iNewIndex As Integer
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim objFont As StdFont
  Dim iLoop As Integer
  Dim iSQLDataType As SQLDataType
  Dim sDescription As String
  Dim fDoingRealElement As Boolean
  
  fSaveOK = True
  fDoingRealElement = (pwfElement Is mwfElement)
  
  ReDim asItems(0)
  ReDim asItems(WFITEMPROPERTYCOUNT, 0)
  
  If fDoingRealElement Then
    For iLoop = 1 To UBound(mavIdentifierLog, 2)
      mavIdentifierLog(4, iLoop) = True
    Next iLoop
  End If
  
  ' Save each screen control.
  For Each ctlControl In Me.Controls
    If fSaveOK And IsWebFormControl(ctlControl) Then
      iNewIndex = UBound(asItems, 2) + 1
      ReDim Preserve asItems(WFITEMPROPERTYCOUNT, iNewIndex)
      
      With ctlControl
        iWFItemType = CLng(.WFItemType)
        
        ' Save all the properties.
        
        'Description
        Select Case iWFItemType
          Case giWFFORMITEM_BUTTON
            sDescription = "Button - '" & IIf(Len(.Caption) > 0, Replace(.Caption, "&&", "&"), vbNullString) & "'"
          Case giWFFORMITEM_DBVALUE, _
            giWFFORMITEM_DBFILE
            
            sDescription = "Database value - " & GetColumnName(.ColumnID)
          Case giWFFORMITEM_LABEL
            sDescription = "Label - '" & IIf(Len(.Caption) > 0, Replace(.Caption, "&&", "&"), vbNullString) & "'"
          Case giWFFORMITEM_INPUTVALUE_CHAR, _
            giWFFORMITEM_INPUTVALUE_DROPDOWN, _
            giWFFORMITEM_INPUTVALUE_LOGIC, _
            giWFFORMITEM_INPUTVALUE_LOOKUP, _
            giWFFORMITEM_INPUTVALUE_DATE, _
            giWFFORMITEM_INPUTVALUE_NUMERIC, _
            giWFFORMITEM_INPUTVALUE_OPTIONGROUP, _
            giWFFORMITEM_INPUTVALUE_GRID, _
            giWFFORMITEM_INPUTVALUE_FILEUPLOAD
            
            sDescription = "Input value - " & .WFIdentifier
          Case giWFFORMITEM_WFVALUE, _
            giWFFORMITEM_WFFILE
            
            sDescription = "Workflow value - " & .WFWorkflowForm & "." & .WFWorkflowValue
          Case giWFFORMITEM_FORMATCODE
            sDescription = "Formatting - " & FormatDescription(IIf(Len(.Caption) > 0, Replace(.Caption, "&&", "&"), vbNullString))
          Case Else
            sDescription = ""
        End Select
        asItems(1, iNewIndex) = sDescription
        
        If fDoingRealElement _
          And WebFormItemHasProperty(iWFItemType, WFITEMPROP_WFIDENTIFIER) Then
          
          For iLoop = 1 To UBound(mavIdentifierLog, 2)
            If mavIdentifierLog(1, iLoop) Is ctlControl Then
              mavIdentifierLog(3, iLoop) = .WFIdentifier
              mavIdentifierLog(4, iLoop) = False
              Exit For
            End If
          Next iLoop
        End If
        
        'Item Type
        asItems(2, iNewIndex) = iWFItemType
        
        'Caption
        If (WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTION)) Or _
          (WebFormControl_HasText(iWFItemType)) Then
          asItems(3, iNewIndex) = IIf(Len(.Caption) > 0, Replace(.Caption, "&&", "&"), vbNullString)
        Else
          asItems(3, iNewIndex) = ""
        End If

        If iWFItemType = giWFFORMITEM_INPUTVALUE_CHAR Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_CHARACTER
          'Input Size
          asItems(7, iNewIndex) = ctlControl.WFInputSize
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.WFDefaultCharValue
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_DATE Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_DATE
          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.WFDefaultValueDateString
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_CHARACTER
          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.DefaultStringValue
        
          'Control Values List (Tab delimited string)
          asItems(47, iNewIndex) = ctlControl.ControlValueList
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_LOOKUP Then
          'Input Return Type
          iSQLDataType = GetColumnDataType(.LookupColumnID)
          Select Case iSQLDataType
          Case dtVARCHAR, dtlongvarchar
            asItems(6, iNewIndex) = giEXPRVALUE_CHARACTER
          Case dtTIMESTAMP
            asItems(6, iNewIndex) = giEXPRVALUE_DATE
          Case dtLONGVARBINARY
            asItems(6, iNewIndex) = giEXPRVALUE_OLE
          Case dtVARBINARY
            asItems(6, iNewIndex) = giEXPRVALUE_PHOTO
          Case dtinteger, dtNUMERIC
            asItems(6, iNewIndex) = giEXPRVALUE_NUMERIC
          Case dtBIT
            asItems(6, iNewIndex) = giEXPRVALUE_LOGIC
          Case Else
            asItems(6, iNewIndex) = giEXPRVALUE_UNDEFINED
          End Select

          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.DefaultStringValue
          'Lookup Table ID
          asItems(48, iNewIndex) = .LookupTableID
          'Lookup Column ID
          asItems(49, iNewIndex) = .LookupColumnID
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_CHARACTER
          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.DefaultStringValue
          
          If ctlControl.NoOptions Then
            asItems(47, iNewIndex) = ""
          Else
            'Control Values List (Tab delimited string)
            asItems(47, iNewIndex) = ctlControl.ControlValueList
          End If
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_LOGIC Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_LOGIC
          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.WFDefaultValue
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_GRID Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_UNDEFINED
          'Input Size
          asItems(7, iNewIndex) = 0
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = 0
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_NUMERIC Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_NUMERIC
          'Input Size
          asItems(7, iNewIndex) = ctlControl.WFInputSize
          'Input Decimals
          asItems(8, iNewIndex) = ctlControl.WFInputDecimals
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ctlControl.WFDefaultNumericValue
          
        ElseIf iWFItemType = giWFFORMITEM_BUTTON Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_UNDEFINED
          'Input Size
          asItems(7, iNewIndex) = 10
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ""

        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_UNDEFINED
          'Input Size
          asItems(7, iNewIndex) = ctlControl.WFInputSize
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = .WFIdentifier
          'Input Default
          asItems(10, iNewIndex) = ""
        
          'File Extensions List (Tab delimited string)
          asItems(66, iNewIndex) = ctlControl.WFFileExtensions
        Else
          'Input Return Type
          asItems(6, iNewIndex) = giEXPRVALUE_UNDEFINED
          'Input Size
          asItems(7, iNewIndex) = 10
          'Input Decimals
          asItems(8, iNewIndex) = 0
          'Input Identifier
          asItems(9, iNewIndex) = ""
          'Input Default
          asItems(10, iNewIndex) = ""
        
        End If
        
        If ((iWFItemType = giWFFORMITEM_DBVALUE) _
          Or (iWFItemType = giWFFORMITEM_DBFILE)) Then
          
          'DB Column ID
          asItems(4, iNewIndex) = .ColumnID
          'DB Record
          asItems(5, iNewIndex) = .WFDatabaseRecord
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_GRID Then
          'DB Column ID
          asItems(4, iNewIndex) = 0
          'DB Record
          asItems(5, iNewIndex) = .WFDatabaseRecord
        Else
          'DB Column ID
          asItems(4, iNewIndex) = 0
          'DB Record
          asItems(5, iNewIndex) = 0
        End If
        
        If ((iWFItemType = giWFFORMITEM_WFVALUE) _
          Or (iWFItemType = giWFFORMITEM_WFFILE)) Then
          'WF Form Identifier
          asItems(11, iNewIndex) = .WFWorkflowForm
          'WF Value Indentifier
          asItems(12, iNewIndex) = .WFWorkflowValue
          
        ElseIf (iWFItemType = giWFFORMITEM_INPUTVALUE_GRID) _
          Or (iWFItemType = giWFFORMITEM_DBVALUE) _
          Or (iWFItemType = giWFFORMITEM_DBFILE) Then
          
          If (.WFDatabaseRecord = 1) Then
            'WF Form Identifier
            asItems(11, iNewIndex) = .WFWorkflowForm
            'WF Value Indentifier
            asItems(12, iNewIndex) = .WFWorkflowValue
          Else
            'WF Form Identifier
            asItems(11, iNewIndex) = ""
            'WF Value Indentifier
            asItems(12, iNewIndex) = ""
          End If
        
        Else
          'WF Form Identifier
          asItems(11, iNewIndex) = ""
          'WF Value Indentifier
          asItems(12, iNewIndex) = ""
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_TABLEID) Then
          asItems(44, iNewIndex) = .TableID
        
          If fDoingRealElement Then
            For iLoop = 1 To UBound(mavIdentifierLog, 2)
              If mavIdentifierLog(1, iLoop) Is ctlControl Then
                mavIdentifierLog(6, iLoop) = .TableID
                Exit For
              End If
            Next iLoop
          End If
        End If
                
        'Dimensions, Coords., Fonts & Colours
        asItems(13, iNewIndex) = TwipsToPixels(.Left)
        asItems(14, iNewIndex) = TwipsToPixels(.Top)
        asItems(15, iNewIndex) = TwipsToPixels(.Width)
        asItems(16, iNewIndex) = TwipsToPixels(.Height)
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLOR) Then
          asItems(17, iNewIndex) = .BackColor
        Else
          asItems(17, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLOR) Then
          asItems(18, iNewIndex) = .ForeColor
        Else
          asItems(18, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FONT) Then
          Set objFont = .Font
          asItems(19, iNewIndex) = objFont.Name
          asItems(20, iNewIndex) = objFont.Size
          asItems(21, iNewIndex) = objFont.Bold
          asItems(22, iNewIndex) = objFont.Italic
          asItems(23, iNewIndex) = objFont.Strikethrough
          asItems(24, iNewIndex) = objFont.Underline
          Set objFont = Nothing
        Else
          asItems(19, iNewIndex) = "Verdana"
          asItems(20, iNewIndex) = "8"
          asItems(21, iNewIndex) = False
          asItems(22, iNewIndex) = False
          asItems(23, iNewIndex) = False
          asItems(24, iNewIndex) = False
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_PICTURE) Then
          asItems(25, iNewIndex) = .PictureID
        Else
          asItems(25, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BORDERSTYLE) Then
          asItems(26, iNewIndex) = CStr(.BorderStyle)
        Else
          asItems(26, iNewIndex) = CStr(vbBSNone)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_ALIGNMENT) Then
          asItems(27, iNewIndex) = .Alignment
        Else
          asItems(27, iNewIndex) = 0
        End If
        
        asItems(28, iNewIndex) = .ControlLevel
        
        If WebFormControl_IsTabStop(iWFItemType) Then
          asItems(29, iNewIndex) = .TabIndex
        Else
          asItems(29, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKSTYLE) Then
          asItems(30, iNewIndex) = .BackStyle
        Else
          asItems(30, iNewIndex) = -1
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLOREVEN) Then
          asItems(31, iNewIndex) = .BackColorEven
        Else
          asItems(31, iNewIndex) = 0
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLORODD) Then
          asItems(32, iNewIndex) = .BackColorOdd
        Else
          asItems(32, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_COLUMNHEADERS) Then
          asItems(33, iNewIndex) = .ColumnHeaders
        Else
          asItems(33, iNewIndex) = False
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLOREVEN) Then
          asItems(34, iNewIndex) = .ForeColorEven
        Else
          asItems(34, iNewIndex) = 0
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLORODD) Then
          asItems(35, iNewIndex) = .ForeColorOdd
        Else
          asItems(35, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADERBACKCOLOR) Then
          asItems(36, iNewIndex) = .HeaderBackColor
        Else
          asItems(36, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADFONT) Then
          Set objFont = .HeadFont
          asItems(37, iNewIndex) = objFont.Name
          asItems(38, iNewIndex) = objFont.Size
          asItems(39, iNewIndex) = objFont.Bold
          asItems(40, iNewIndex) = objFont.Italic
          asItems(41, iNewIndex) = objFont.Strikethrough
          asItems(42, iNewIndex) = objFont.Underline
          Set objFont = Nothing
        Else
          asItems(37, iNewIndex) = "Verdana"
          asItems(38, iNewIndex) = "8"
          asItems(39, iNewIndex) = False
          asItems(40, iNewIndex) = False
          asItems(41, iNewIndex) = False
          asItems(42, iNewIndex) = False
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADLINES) Then
          asItems(43, iNewIndex) = .HeadLines
        Else
          asItems(43, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
          asItems(45, iNewIndex) = .ForeColorHighlight
        Else
          asItems(45, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
          asItems(46, iNewIndex) = .BackColorHighlight
        Else
          asItems(46, iNewIndex) = 0
        End If
      
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDTABLEID) Then
          asItems(50, iNewIndex) = .WFRecordTableID
        Else
          asItems(50, iNewIndex) = 0
        End If
      
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_ORIENTATION) Then
          asItems(51, iNewIndex) = .Alignment
        Else
          asItems(51, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDORDER) Then
          asItems(52, iNewIndex) = .WFRecordOrderID
        Else
          asItems(52, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDFILTER) Then
          asItems(53, iNewIndex) = .WFRecordFilterID
        Else
          asItems(53, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_SUBMITTYPE) Then
          asItems(54, iNewIndex) = .Behaviour
        Else
          asItems(54, iNewIndex) = WORKFLOWBUTTONACTION_SUBMIT
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_MANDATORY) Then
          asItems(55, iNewIndex) = .Mandatory
        Else
          asItems(55, iNewIndex) = False
        End If
      
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CALCULATION) _
          Or WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
          
          asItems(56, iNewIndex) = .CalculationID
        Else
          asItems(56, iNewIndex) = 0
        End If

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTIONTYPE) Then
          asItems(57, iNewIndex) = .CaptionType
        Else
          asItems(57, iNewIndex) = giWFDATAVALUE_FIXED
        End If

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
          asItems(58, iNewIndex) = .DefaultValueType
        Else
          asItems(58, iNewIndex) = giWFDATAVALUE_FIXED
        End If

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_VERTICALOFFSET) Then
          asItems(59, iNewIndex) = .VerticalOffsetBehaviour
          asItems(61, iNewIndex) = .VerticalOffset
        Else
          asItems(59, iNewIndex) = 0
          asItems(61, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HORIZONTALOFFSET) Then
          asItems(60, iNewIndex) = .HorizontalOffsetBehaviour
          asItems(62, iNewIndex) = .HorizontalOffset
        Else
          asItems(60, iNewIndex) = 0
          asItems(62, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
          asItems(61, iNewIndex) = .HeightBehaviour
        Else
          asItems(61, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
          asItems(62, iNewIndex) = .WidthBehaviour
        Else
          asItems(62, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_PASSWORDTYPE) Then
          asItems(65, iNewIndex) = .PasswordType
        Else
          asItems(65, iNewIndex) = False
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_LOOKUPFILTERCOLUMN) Then
          asItems(67, iNewIndex) = .LookupFilterColumn
        Else
          asItems(67, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_LOOKUPFILTEROPERATOR) Then
          asItems(68, iNewIndex) = .LookupFilterOperator
        Else
          asItems(68, iNewIndex) = 0
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_LOOKUPFILTERVALUE) Then
          asItems(69, iNewIndex) = .LookupFilterValue
        Else
          asItems(69, iNewIndex) = ""
        End If
      End With
      
    End If
  Next ctlControl
  
  If fDoingRealElement Then
    For iLoop = 1 To UBound(mavIdentifierLog, 2)
      If mavIdentifierLog(4, iLoop) Then
        mavIdentifierLog(3, iLoop) = ""
      End If
    Next iLoop
  End If
  
  pwfElement.Items = asItems
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SaveWebFormItems = fSaveOK
  Exit Function
  
ErrorTrap:
  fSaveOK = False
  Resume TidyUpAndExit
  
End Function

Private Function DeselectAllControls(Optional pctlException As VB.Control) As Boolean
  
  Dim iCount As Integer
  Dim ctlControl As Control

  ' Hide all the selection markers
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    ASRSelectionMarkers(iCount).Visible = False
  Next iCount
  
  ' Deselect the controls
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      With ctlControl
        ctlControl.Selected = False
      End With
    End If
  Next ctlControl
  
  DeselectAllControls = True
  
End Function

Private Function WebFormControl_DragDrop(pctlControl As VB.Control, pCtlSource As Control, pSngX As Single, pSngY As Single) As Boolean
  
  ' Drop a control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  With pctlControl
    fOK = DropControl(.Container, pCtlSource, pSngX + .Left, pSngY + .Top)
  End With
  
TidyUpAndExit:
  If Not fOK Then
    MsgBox "Unable to drop the control." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  WebFormControl_DragDrop = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function WebFormControl_MouseMove(pctlControl As VB.Control, pButton As Integer, pSngX As Single, pSngY As Single) As Boolean
  
  ' Move the control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim lngNewX As Long
  Dim lngNewY As Long
  
  If mfReadOnly Then
    WebFormControl_MouseMove = True
    Exit Function
  End If
  
  ' Remove the original offset of the mouse cursor
  pSngX = pSngX - mlngXOffset
  pSngY = pSngY - mlngYOffset
  
  fOK = True
  
  ' Only run if the mouse pointer has moved significantly
  If (mlngLastX > pSngX + giGridX) Or (mlngLastX < pSngX - giGridX) _
      Or (mlngLastY > pSngY + giGridY) Or (mlngLastY < pSngY - giGridY) Then
 
    ' Move the selected controls if the left button key is down, and the control is selected
    If pButton = vbLeftButton And pctlControl.Selected Then
    
      For iCount = 1 To ASRSelectionMarkers.Count - 1
        With ASRSelectionMarkers(iCount)
          If .Visible Then
            .ShowSelectionMarkers False
          
            lngNewX = AlignX(pSngX + .AttachedObject.Left)
            lngNewY = AlignX(pSngY + .AttachedObject.Top)
            .AttachedObject.Move lngNewX, lngNewY
          End If
        End With
      Next iCount
    
      gfMoveSelection = True

    End If
      
  End If

TidyUpAndExit:
  WebFormControl_MouseMove = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function WebFormControl_MouseUp(pctlControl As VB.Control, piButton As Integer, piShift As Integer, X As Single, Y As Single) As Boolean
  
  ' Actually move the selected controls to the positions of their movement frames.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim lngXMouse As Long
  Dim lngYMouse As Long
  Dim ctlControl As VB.Control
  Dim avWebFormControls() As Variant
  Dim iCount As Integer
    
  fOK = True

  Select Case piButton
    
    ' Handle left button presses.
    Case vbLeftButton

      ' Deselect all OTHER screen controls if the CTRL or SHIFT keys are not pressed,
      ' and if we do not already have the control selected as part of a multiple selection.
      If Not gfMoveSelection Then
      
        If ((piShift And vbShiftMask) = 0) And ((piShift And vbCtrlMask) = 0) Then
          DeselectAllControls
        End If
          
        ' Toggle this control if the shift/ctrl key is pressed
        If ((piShift And vbShiftMask) <> 0) Or ((piShift And vbCtrlMask) <> 0) Then
          pctlControl.Selected = Not pctlControl.Selected
          'Debug.Print pctlControl.Selected
        Else
          DeselectAllControls
          pctlControl.Selected = True
        End If
        
        ' JDM - 20/08/02 - Fault 4309 - Holding down control now selects/deselects controls
        If pctlControl.Selected Then
          SelectControl pctlControl
        Else
          DeselectControl pctlControl
        End If
        
      Else

        ' End placementing of all selected objects
        For iCount = 1 To ASRSelectionMarkers.Count - 1
          With ASRSelectionMarkers(iCount)
            If .Visible Then
              .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize
            End If
          End With
        Next iCount
        
        ' Flag screen as having changed
        IsChanged = True
        
      End If

      ' Show all selected selection markers
      For iCount = 1 To ASRSelectionMarkers.Count - 1
        ASRSelectionMarkers(iCount).ShowSelectionMarkers True
      Next iCount
      
      ' Refresh the properties screen.
      frmSysMgr.RefreshMenu
      Set frmWorkflowWFItemProps.CurrentWebForm = Me
      frmWorkflowWFItemProps.RefreshProperties
      
    ' Handle right button presses.
    Case vbRightButton
      UI.GetMousePos lngXMouse, lngYMouse
      'frmSysMgr.tbMain.PopupMenu "ID_mnuWebFormEdit", ssPopupMenuLeftAlign, lngXMouse, lngYMouse
      frmSysMgr.tbMain.Bands("ID_mnuWebFormEdit").TrackPopup -1, -1
  End Select

  gfMoveSelection = False

TidyUpAndExit:

  ' Stop moving the control.
  gfMoveSelection = False

  ' Disassociate object variables.
  Set ctlControl = Nothing
  UI.UnlockWindow
  WebFormControl_MouseUp = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get ReadOnly() As Boolean
  ReadOnly = mfReadOnly
  
End Property


Private Function AutoLabel(pVarPageContainer As Variant, pSngX As Single, pSngY As Single, sCaption As String) As Boolean
  
  ' Drop the given control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As WorkflowWebFormItemTypes
  Dim lngColumnID As Long
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim objMisc As New Misc
  Dim ctlControl As VB.Control
  
  fOK = True
  
  If fOK Then
  
    iControlType = giWFFORMITEM_LABEL
    Set ctlControl = AddControl(iControlType)
              
    fOK = Not (ctlControl Is Nothing)
          
    'Check that a new control was added successfully
    If fOK Then
  
      With ctlControl

        Set .Container = pVarPageContainer
        .Left = AlignX((CLng(pSngX) - Me.TextWidth(sCaption + Space(5))))
        If .Left < 0 Then
          .Left = CLng(pSngX)
          .Top = AlignY((CLng(pSngY) - (Me.TextHeight(sCaption) + 20)))
        Else
          .Top = AlignY(CLng(pSngY))
        End If
        
        .ColumnID = 0
        
        .WFIdentifier = "Label" & ctlControl.Index
        
        ' Initialise the new control's font and forecolour.
        If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
          Set objFont = New StdFont
          objFont.Name = Me.Font.Name
          objFont.Size = Me.Font.Size
          objFont.Bold = Me.Font.Bold
          objFont.Italic = Me.Font.Italic
          Set .Font = objFont
          Set objFont = Nothing
        End If
            
        If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
          ' AE20080609 Fault #13080
          '.ForeColor = Me.ForeColor
          .ForeColor = Me.ForeColor
        End If
        
        If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
          .BackColor = Me.BackColor
        End If
        
        If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION) Then
          .Caption = Replace(sCaption, "_", " ") & ":"
        End If
            
        ' Default the control's propertes.
        fOK = AutoSizeControl(ctlControl)
              
        If fOK Then
          fOK = SelectControl(ctlControl)
        End If
            
        If fOK Then
          .Visible = True
          .ZOrder 0
        End If
      
        If giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL Then
          giUndo_ControlAutoLabelIndex = .Index
          gsUndo_ControlAutoLabelType = .Name
        Else
          giUndo_ControlAutoLabelIndex = .Index
          gsUndo_ControlAutoLabelType = ""
        End If
      
      End With
      
    End If
          
    ' Disassociate object variables.
    Set ctlControl = Nothing
          
  End If
    
  ' Set focus on the screen designer form.
  Me.SetFocus
  
    
  If fOK Then
    ' Mark the screen as having changed.
    mfChanged = True
    frmSysMgr.RefreshMenu
  
    ' Refresh the properties screen.
    Set frmWorkflowWFItemProps.CurrentWebForm = Me
    frmWorkflowWFItemProps.RefreshProperties
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objMisc = Nothing
  Set objFont = Nothing
  Set ctlControl = Nothing
  ' Return the success/failure value.
  AutoLabel = fOK
  Exit Function

ErrorTrap:
  ' Flag the error.
  fOK = False
  MsgBox "Could not automatically add a label for this control." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit
  
End Function

Private Sub ASRDummyLine_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  WebFormControl_DragDrop ASRDummyLine(Index), Source, X, Y
End Sub

Private Sub ASRDummyLine_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  WebFormControl_MouseDown ASRDummyLine(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  WebFormControl_MouseMove ASRDummyLine(Index), Button, X, Y
End Sub

Private Sub ASRDummyLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  WebFormControl_MouseUp ASRDummyLine(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyImage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyCheckBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyTextBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyCombo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyLine_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyLabel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub btnWorkflow_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub WebFormControl_MouseDown(pctlControl As VB.Control, piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)

  Dim iCount As Integer

  ' Only handle left button presses here.
  If piButton <> vbLeftButton Then
    Exit Sub
  End If

  mlngXOffset = pSngX
  mlngYOffset = pSngY

  ' Flag the selected selction markers to be moved
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    ASRSelectionMarkers(iCount).ShowSelectionMarkers False
  Next iCount

End Sub

Public Function SelectControl(pctlControl As VB.Control) As Boolean
  
  Dim iIndex As Integer
  Dim iCount As Integer
  Dim objMarkers As Object
  
  ' Have selection markers for this control already been created
  If pctlControl.Selected Then
  
    If pctlControl.Tag = "" Then
    
      iIndex = ASRSelectionMarkers.Count
      Load ASRSelectionMarkers(iIndex)
      
      With ASRSelectionMarkers(iIndex)
        Set .Container = pctlControl.Container
        .WFDesigner = True
        .AttachedObject = pctlControl
        .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
        .RefreshSelectionMarkers True
        .ZOrder 0
        .Visible = True
      End With
    
      pctlControl.Tag = iIndex
    
    Else
      With ASRSelectionMarkers(pctlControl.Tag)
        ' Ensure the selection markers are in the same container
        ' as the control - this can get out of synch sometimes.
        Set .Container = pctlControl.Container
        .ZOrder 0
        .Visible = True
      End With
    End If
  End If
  
  SelectControl = True
  
End Function

Public Function LoadWebFormItems() As Boolean
  
  ' Load controls onto the selected tab page.
  On Error GoTo ErrorTrap
    
  Dim fLoadOk As Boolean
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim iDisplayType As Integer
  Dim lngTableID As Long
  Dim lngPictureID As Long
  Dim sFileName As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim iNextIndex As Integer
  Dim iRecordCount As Integer
  Dim iCount As Integer
  Dim asItems() As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim frmUsage As frmUsage
  Dim asMessages() As String
  Dim asItemValues() As String
  Dim avTabIndexes() As Variant
  Dim sCaption As String
  Dim lngExprID As Long
  
  iNextIndex = 1
  fLoadOk = True
  ReDim asMessages(0)
 
  ReDim avTabIndexes(1, 0)
  
  Screen.MousePointer = vbHourglass
  
  ' Load the screen controls if everything is okay so far.
  If fLoadOk Then
    
    ' Log the original identifiers of the controls
    ' Column 1 = the control
    ' Column 2 = original identifier
    ' Column 3 = current identifier (defaulted to original value, updated in SaveWebFormItems)
    ' Column 4 = deleted flag
    ' Column 5 = original recordSelector table
    ' Column 6 = current recordSelector table (defaulted to original value, updated in SaveWebFormItems)
    ' NB. Row 0 is for the form itself.
    ReDim mavIdentifierLog(6, 0)
    mavIdentifierLog(2, 0) = mwfElement.Identifier
    mavIdentifierLog(3, 0) = mwfElement.Identifier
    mavIdentifierLog(4, 0) = False
    mavIdentifierLog(5, 0) = 0
    mavIdentifierLog(6, 0) = 0
    
    asItems = mwfElement.Items
    
    For iLoop = 1 To UBound(asItems, 2) Step 1
    
      ' Get the control's type.
      iWFItemType = CInt(asItems(2, iLoop))
  
      ' Create the new control.
      Set ctlControl = AddControl(iWFItemType)

      If Not ctlControl Is Nothing Then             ' Indent 05 - start
        
        ' Set the Web form Item Identifier for this control.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_WFIDENTIFIER) Then
          'JPD 20061010 Fault 11355
          'ctlControl.WFIdentifier = asItems(1, iLoop)
          ctlControl.WFIdentifier = asItems(9, iLoop)
        
          ReDim Preserve mavIdentifierLog(6, UBound(mavIdentifierLog, 2) + 1)
          Set mavIdentifierLog(1, UBound(mavIdentifierLog, 2)) = ctlControl
          mavIdentifierLog(2, UBound(mavIdentifierLog, 2)) = asItems(9, iLoop)
          mavIdentifierLog(3, UBound(mavIdentifierLog, 2)) = asItems(9, iLoop)
          mavIdentifierLog(4, UBound(mavIdentifierLog, 2)) = True
          mavIdentifierLog(5, UBound(mavIdentifierLog, 2)) = 0
          mavIdentifierLog(6, UBound(mavIdentifierLog, 2)) = 0
        End If
        
        ' Set the item type of the control
        ctlControl.WFItemType = CInt(asItems(2, iLoop))
        
        ' Set the controls caption.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTION) Then
          sCaption = asItems(3, iLoop)

          If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTIONTYPE) Then
            If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CALCULATION) Then
              If asItems(57, iLoop) = giWFDATAVALUE_CALC Then
                lngExprID = asItems(56, iLoop)
                sCaption = GetExpressionName(lngExprID)
                
                If Len(Trim(sCaption)) = 0 Then
                  sCaption = "<Calculated>"
                Else
                  sCaption = "<" & sCaption & ">"
                End If
              End If
            End If
          End If
          
          ctlControl.Caption = Replace(sCaption, "&", "&&")
        End If
                    
        If ((iWFItemType = giWFFORMITEM_DBVALUE) _
          Or (iWFItemType = giWFFORMITEM_DBFILE)) Then
          
          ' Set the control's column ID.
          ctlControl.ColumnID = IIf(asItems(4, iLoop) = "", 0, asItems(4, iLoop))
          ctlControl.WFDatabaseRecord = IIf(asItems(5, iLoop) = "", 0, asItems(5, iLoop))
          If (ctlControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD) Then
            ctlControl.WFWorkflowForm = asItems(11, iLoop)
            ctlControl.WFWorkflowValue = asItems(12, iLoop)
          End If

          ctlControl.ToolTipText = "<" & GetColumnName(ctlControl.ColumnID, False) & ">"
          
          If (iWFItemType = giWFFORMITEM_DBVALUE) Then
            ctlControl.Caption = ctlControl.ToolTipText
          End If
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_GRID Then
          ctlControl.WFDatabaseRecord = IIf(asItems(5, iLoop) = "", 0, asItems(5, iLoop))
          If (ctlControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD) Then
            ctlControl.WFWorkflowForm = asItems(11, iLoop)
            ctlControl.WFWorkflowValue = asItems(12, iLoop)
          End If
        End If
        
        If iWFItemType = giWFFORMITEM_INPUTVALUE_CHAR Then
          'Input Size
          ctlControl.WFInputSize = asItems(7, iLoop)
          'Input Decimals
          ctlControl.WFInputDecimals = 0
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.WFDefaultCharValue = CStr(asItems(10, iLoop))
          ctlControl.Caption = " " & ctlControl.WFDefaultCharValue
          ctlControl.PasswordType = asItems(65, iLoop)
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_DATE Then
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.WFDefaultValueDateString = asItems(10, iLoop)
          ctlControl.Caption = " " & ctlControl.WFDefaultValueDateString
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN Then
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.DefaultStringValue = asItems(10, iLoop)
          ' Tab delimited Control Value List
          ctlControl.ControlValueList = asItems(47, iLoop)
          ctlControl.Caption = asItems(10, iLoop)
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_LOGIC Then
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.WFDefaultValue = CBool(IIf(asItems(10, iLoop) = "", False, asItems(10, iLoop)))
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_LOOKUP Then
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.DefaultStringValue = asItems(10, iLoop)
          ctlControl.Caption = asItems(10, iLoop)
          ctlControl.LookupTableID = asItems(48, iLoop)
          ctlControl.LookupColumnID = asItems(49, iLoop)
          ctlControl.LookupFilterColumn = asItems(67, iLoop)
          ctlControl.LookupFilterOperator = asItems(68, iLoop)
          ctlControl.LookupFilterValue = asItems(69, iLoop)
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_NUMERIC Then
          'Input Size
          ctlControl.WFInputSize = asItems(7, iLoop)
          'Input Decimals
          ctlControl.WFInputDecimals = asItems(8, iLoop)
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          'Input Default
          ctlControl.WFDefaultNumericValue = asItems(10, iLoop)
          ctlControl.Caption = " " & ctlControl.WFDefaultNumericValue
          
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
          'Input Identifier
          ctlControl.WFIdentifier = asItems(9, iLoop)
          ctlControl.Caption = Replace(asItems(3, iLoop), "&", "&&")
          If asItems(47, iLoop) = vbNullString Then
            ctlControl.DefaultStringValue = vbNullString
            ctlControl.NoOptions = True
          Else
            asItemValues = Split(asItems(47, iLoop), vbTab)
            ctlControl.SetOptions asItemValues
            ctlControl.DefaultStringValue = asItems(10, iLoop)
            ctlControl.SelectOption (asItems(10, iLoop))
            ctlControl.NoOptions = False
          End If
        
        ElseIf iWFItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
          ctlControl.WFInputSize = asItems(7, iLoop)
          ctlControl.WFFileExtensions = asItems(66, iLoop)
        End If
        
        If ((iWFItemType = giWFFORMITEM_WFVALUE) _
          Or (iWFItemType = giWFFORMITEM_WFFILE)) Then
          ctlControl.WFWorkflowForm = asItems(11, iLoop)
          ctlControl.WFWorkflowValue = asItems(12, iLoop)
          
          If (iWFItemType = giWFFORMITEM_WFVALUE) Then
            ctlControl.Caption = "<" & ctlControl.WFWorkflowForm & " : " & ctlControl.WFWorkflowValue & ">"
          Else
            ctlControl.ToolTipText = "<" & ctlControl.WFWorkflowForm & " : " & ctlControl.WFWorkflowValue & ">"
          End If
        End If
        
        'AE20080115
        'Moved up here because the orientation needs to be set before
        'the height and width
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_ORIENTATION) Then
          ctlControl.Alignment = asItems(51, iLoop)
        End If
        
        ' Set the control's location.
        ctlControl.Left = PixelsToTwips(CLng(asItems(13, iLoop)))
        ctlControl.Top = PixelsToTwips(CLng(asItems(14, iLoop)))
        
        ' Set the control's dimensions.
        ctlControl.Width = PixelsToTwips(CLng(asItems(15, iLoop)))
        ctlControl.Height = PixelsToTwips(CLng(asItems(16, iLoop)))
       
       ' Set the control's size/behaviour properties
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_VERTICALOFFSET) Then
          ctlControl.VerticalOffsetBehaviour = asItems(59, iLoop)
          ctlControl.VerticalOffset = asItems(61, iLoop)
          
          If ctlControl.VerticalOffset = 0 Then
            If ctlControl.VerticalOffsetBehaviour = offsetTop Then
              ctlControl.VerticalOffset = ctlControl.Top
            Else
              ctlControl.VerticalOffset = PixelsToTwips(mwfElement.WebFormHeight) - ctlControl.Top - ctlControl.Height
            End If
          End If
        End If
                
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HORIZONTALOFFSET) Then
          ctlControl.HorizontalOffsetBehaviour = asItems(60, iLoop)
          ctlControl.HorizontalOffset = asItems(62, iLoop)
          
          If ctlControl.HorizontalOffset = 0 Then
            If ctlControl.HorizontalOffsetBehaviour = offsetLeft Then
              ctlControl.HorizontalOffset = ctlControl.Left
            Else
              ctlControl.HorizontalOffset = PixelsToTwips(mwfElement.WebFormWidth) - ctlControl.Left - ctlControl.Width
            End If
          End If
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
          ctlControl.HeightBehaviour = asItems(61, iLoop)
          
          If ctlControl.HeightBehaviour <> behaveFixed Then
            ctlControl.Top = 0
            ctlControl.Height = PixelsToTwips(mwfElement.WebFormHeight)
          End If
        End If
                
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
          ctlControl.WidthBehaviour = asItems(62, iLoop)
          
          If ctlControl.WidthBehaviour <> behaveFixed Then
            ctlControl.Left = 0
            ctlControl.Width = PixelsToTwips(mwfElement.WebFormWidth)
          End If
        End If
        
        ' Set the BackColor and ForeColor properties.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLOR) Then
          ctlControl.BackColor = asItems(17, iLoop)
        End If

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLOR) Then
          ctlControl.ForeColor = asItems(18, iLoop)
        End If

        ' Font properties.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FONT) Then
          Set objFont = New StdFont
          objFont.Name = asItems(19, iLoop)
          objFont.Size = asItems(20, iLoop)
          objFont.Bold = asItems(21, iLoop)
          objFont.Italic = asItems(22, iLoop)
          objFont.Strikethrough = asItems(23, iLoop)
          objFont.Underline = asItems(24, iLoop)
          Set ctlControl.Font = objFont
          Set objFont = Nothing
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_PICTURE) Then
          ctlControl.PictureID = IIf(asItems(25, iLoop) = "", 0, asItems(25, iLoop))
          
          If ctlControl.PictureID > 0 Then
            recPictEdit.Index = "idxID"
            recPictEdit.Seek "=", ctlControl.PictureID
                
            If Not recPictEdit.NoMatch Then
              sFileName = ReadPicture
              ctlControl.Picture = sFileName

              Kill sFileName
            End If
          End If
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BORDERSTYLE) Then
          ctlControl.BorderStyle = IIf(asItems(26, iLoop) = "1", vbFixedSingle, vbBSNone)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_ALIGNMENT) Then
          ctlControl.Alignment = asItems(27, iLoop)
        End If
        
        ctlControl.ControlLevel = asItems(28, iLoop)
        If WebFormControl_IsTabStop(iWFItemType) Then
          ReDim Preserve avTabIndexes(1, UBound(avTabIndexes, 2) + 1)
          Set avTabIndexes(0, UBound(avTabIndexes, 2)) = ctlControl
          avTabIndexes(1, UBound(avTabIndexes, 2)) = asItems(29, iLoop)
        End If
  
        ' Set the BackStyle properties.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKSTYLE) Then
          ctlControl.BackStyle = asItems(30, iLoop)
        End If
  
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLOREVEN) Then
          ctlControl.BackColorEven = asItems(31, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLORODD) Then
          ctlControl.BackColorOdd = asItems(32, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_COLUMNHEADERS) Then
          ctlControl.ColumnHeaders = asItems(33, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLOREVEN) Then
          ctlControl.ForeColorEven = asItems(34, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLORODD) Then
          ctlControl.ForeColorOdd = asItems(35, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADERBACKCOLOR) Then
          ctlControl.HeaderBackColor = asItems(36, iLoop)
        End If
        
         ' HeadFont properties.
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADFONT) Then
          Set objFont = New StdFont
          objFont.Name = asItems(37, iLoop)
          objFont.Size = asItems(38, iLoop)
          objFont.Bold = asItems(39, iLoop)
          objFont.Italic = asItems(40, iLoop)
          objFont.Strikethrough = asItems(41, iLoop)
          objFont.Underline = asItems(42, iLoop)
          Set ctlControl.HeadFont = objFont
          Set objFont = Nothing
        End If

        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEADLINES) Then
          ctlControl.HeadLines = asItems(43, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_TABLEID) Then
          ctlControl.TableID = asItems(44, iLoop)
        
          mavIdentifierLog(5, UBound(mavIdentifierLog, 2)) = asItems(44, iLoop)
          mavIdentifierLog(6, UBound(mavIdentifierLog, 2)) = asItems(44, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
          ctlControl.ForeColorHighlight = asItems(45, iLoop)
        End If
        
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
          ctlControl.BackColorHighlight = asItems(46, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDTABLEID) Then
          ctlControl.WFRecordTableID = asItems(50, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDORDER) Then
          ctlControl.WFRecordOrderID = asItems(52, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_RECORDFILTER) Then
          ctlControl.WFRecordFilterID = asItems(53, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_SUBMITTYPE) Then
          ctlControl.Behaviour = asItems(54, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_MANDATORY) Then
          ctlControl.Mandatory = asItems(55, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CALCULATION) _
          Or WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
          ctlControl.CalculationID = asItems(56, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_CAPTIONTYPE) Then
          ctlControl.CaptionType = asItems(57, iLoop)
        End If
       
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
          ctlControl.DefaultValueType = asItems(58, iLoop)
        End If
       
        ValidateIdentifiers ctlControl, asMessages
        
        ctlControl.Visible = True
        
        ' Disassociate object variables.
        Set ctlControl = Nothing
      End If
    Next iLoop

    RefreshExpressionNames

    For iLoop = 1 To UBound(avTabIndexes, 2)
      iNextIndex = 0
      iLoop3 = 0
      
      For iLoop2 = 1 To UBound(avTabIndexes, 2)
        If (iNextIndex = 0) Or ((avTabIndexes(1, iLoop2) > 0) And (avTabIndexes(1, iLoop2) <= iNextIndex)) Then
          iNextIndex = avTabIndexes(1, iLoop2)
          iLoop3 = iLoop2
        End If
      Next iLoop2
      
      avTabIndexes(0, iLoop3).TabIndex = iLoop
      avTabIndexes(1, iLoop3) = 0
    Next iLoop
    
    If UBound(asMessages) > 0 Then
      Set frmUsage = New frmUsage
      frmUsage.ResetList
        
      For iLoop = 1 To UBound(asMessages)
        frmUsage.AddToList (asMessages(iLoop))
      Next iLoop
    
      Screen.MousePointer = vbNormal
      frmUsage.ShowMessage "Workflow '" & Trim(mfrmCallingForm.WorkflowName) & "'", "The following web form items are invalid, and will need reviewing:", _
        UsageCheckObject.Workflow, _
        USAGEBUTTONS_PRINT + USAGEBUTTONS_OK, "validation"
      
      UnLoad frmUsage
      Set frmUsage = Nothing
    End If
    
    fLoadOk = SetControlLevel
    
  End If

TidyUpAndExit:

  ' Unlock the window refreshing.
  UI.UnlockWindow
    
  ' Reset the screen mousepointer.
  Screen.MousePointer = vbNormal
  
  LoadWebFormItems = fLoadOk
  Exit Function
  
ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading Web Form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Function LoadWebForm() As Boolean
  
  ' Load controls onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fLoadOk As Boolean
  Dim iPageNo As Integer
  Dim iCtrlType As Integer
  Dim iDisplayType As Integer
  Dim lngTableID As Long
  Dim lngPictureID As Long
  Dim sFileName As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim iNextIndex As Integer
  Dim iRecordCount As Integer
  Dim iTabPages As Integer
  Dim iCount As Integer
  Dim tmpWFElement As COAWF_Webform
  
  iNextIndex = 1
  fLoadOk = True
  mfForcedChanged = False
  
  Screen.MousePointer = vbHourglass
  
  ' Load the screen properties.
  If fLoadOk Then
    
    ' Lock the screen refeshing.
    UI.LockWindow Me.hWnd
    
    mfChanged = False
    
    ' Set the Web From designer caption.
    If Trim(mwfElement.Identifier) = vbNullString Then
      Me.WFIdentifier = mwfElement.Caption
    Else
      Me.WFIdentifier = mwfElement.Identifier
    End If
    
    Me.Caption = mwfElement.Caption
    
'    If (PixelsToTwips(mwfElement.WebFormWidth) < MIN_FORM_WIDTH) Or (PixelsToTwips(mwfElement.WebFormHeight) < MIN_FORM_HEIGHT) Then
'      Me.BackColor = vbWhite
'    Else
      Me.BackColor = mwfElement.WebFormBGColor
'    End If
    
    Me.ForeColor = mwfElement.WebFormFGColor
    Set Me.Font = mwfElement.WebFormDefaultFont
    
    ' Set the web form icon.
    lngPictureID = IIf(IsNull(mwfElement.WebFormBGImageID), 0, mwfElement.WebFormBGImageID)
    If lngPictureID > 0 Then
      recPictEdit.Index = "idxID"
      recPictEdit.Seek "=", lngPictureID
      If Not recPictEdit.NoMatch Then
        sFileName = ReadPicture
        Me.PictureID = mwfElement.WebFormBGImageID
        Me.PictureLocation = mwfElement.WebFormBGImageLocation
        Me.Picture = LoadPicture(sFileName)
        Kill sFileName
      End If
    End If
  
    Me.TimeoutFrequency = mwfElement.WebFormTimeoutFrequency
    Me.TimeoutPeriod = mwfElement.WebFormTimeoutPeriod
    Me.TimeoutExcludeWeekend = mwfElement.WebFormTimeoutExcludeWeekend
    Me.DescriptionExprID = mwfElement.DescriptionExprID
    Me.DescriptionHasWorkflowName = mwfElement.DescriptionHasWorkflowName
    Me.DescriptionHasElementCaption = mwfElement.DescriptionHasElementCaption
  
    Me.WFCompletionMessageType = mwfElement.WFCompletionMessageType
    Me.WFCompletionMessage = mwfElement.WFCompletionMessage
    Me.WFSavedForLaterMessageType = mwfElement.WFSavedForLaterMessageType
    Me.WFSavedForLaterMessage = mwfElement.WFSavedForLaterMessage
    Me.WFFollowOnFormsMessageType = mwfElement.WFFollowOnFormsMessageType
    Me.WFFollowOnFormsMessage = mwfElement.WFFollowOnFormsMessage
  
    Me.Validations = mwfElement.Validations
  End If
       
  ' Load the web form items (controls)
  LoadWebFormItems

  IsChanged = False
  
TidyUpAndExit:

  ' Unlock the window refreshing.
  UI.UnlockWindow
  
  ' Position, Resize and Move Designer
  ' Resize form.
  If (PixelsToTwips(mwfElement.WebFormWidth) < MIN_FORM_WIDTH) Then
    Me.Width = (Screen.Width / 2) - 720
    mlngLastFormWidth = Me.Width
  Else
    Me.Width = PixelsToTwips(mwfElement.WebFormWidth) + (Me.Width - Me.ScaleWidth)
    mlngLastFormWidth = Me.Width
  End If
  
  DoEvents
  
  If (PixelsToTwips(mwfElement.WebFormHeight) < MIN_FORM_HEIGHT) Then
    Me.Height = Forms(0).ScaleHeight / 2
    mlngLastFormheight = Me.Height
  Else
    Me.Height = PixelsToTwips(mwfElement.WebFormHeight) + (Me.Height - Me.ScaleHeight)
    mlngLastFormheight = Me.Height
  End If
  
  ' Position the form.
  If Me.Height > Forms(0).ScaleHeight Then
    Me.Top = Forms(0).ScaleHeight / 4
  Else
    Me.Top = (Forms(0).ScaleHeight - Me.Height) / 2
  End If
  
  If Me.Width > (Forms(0).ScaleWidth - frmWorkflowWFToolbox.Width - frmWorkflowWFItemProps.Width) Then
    Me.Left = frmWorkflowWFToolbox.Width + 360
  Else
    Me.Left = frmWorkflowWFToolbox.Width + _
                ((frmWorkflowWFItemProps.Left - frmWorkflowWFToolbox.Width - Me.Width) / 2)
  End If

  ' Reset the screen moousepointer.
  Screen.MousePointer = vbNormal
  mfChanged = mfForcedChanged
  
  LoadWebForm = fLoadOk
  Exit Function
  
ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading Web Form Element." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Function WebFormControl_KeyMove(pSngX As Single, pSngY As Single) As Boolean
  ' Move the control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  
  fOK = True
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
    
        If .AttachedObject.Selected Then
          .AttachedObject.Move pSngX + .AttachedObject.Left, pSngY + .AttachedObject.Top
        End If
      
      End If
    End With
  Next iCount
  
  ' Flag screen as having changed
  IsChanged = True

TidyUpAndExit:
  WebFormControl_KeyMove = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function SendSelectedControlsToBack()
  ' Scroll through each selected control and send to back
  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .AttachedObject.ZOrder 1
      End If
    End With
  Next iCount

  ' Flag screen as having changed
  IsChanged = True

End Function

Private Function BringSelectedControlsToFront()
' Scroll through each selected control and bring to front
  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .AttachedObject.ZOrder 0
      End If
    End With
  Next iCount

  ' Flag screen as having changed
  IsChanged = True

End Function

Private Function DeselectControl(pctlControl As VB.Control) As Boolean
 
  ' Deselect current control
  ASRSelectionMarkers(pctlControl.Tag).Visible = False
  pctlControl.Selected = False

  DeselectControl = True
  
End Function

Public Function ScreenHasControls() As Boolean

' Does this screen have any user controls on it
  Dim ctlControl As VB.Control
  
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      ScreenHasControls = True
      Exit Function
    End If
  Next ctlControl

End Function

Private Function LeftAlignSelectedControls()

' Left align the selected controls
  Dim iCount As Integer
  Dim lngLeft As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngLeft = .Left
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngLeft
        .AttachedObject.Left = lngLeft + .MarkerSize
        Application.Changed = True
      End If
    End With
  Next iCount
    
  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function

Private Function RightAlignSelectedControls()

' Right align the selected controls
  Dim iCount As Integer
  Dim lngRight As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngRight = .Left + .Width
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngRight - .Width
        .AttachedObject.Left = .Left + .MarkerSize
      End If
    End With
  Next iCount

  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function


Private Function CentreAlignSelectedControls()

' Centre align the selected controls
  Dim iCount As Integer
  Dim lngCentre As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngCentre = .Left + (.Width / 2)
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngCentre - (.Width / 2)
        .AttachedObject.Left = .Left + .MarkerSize
      End If
    End With
  Next iCount

  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function

Private Function TopAlignSelectedControls()

' Top align the selected controls
  Dim iCount As Integer
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngTop
        .AttachedObject.Top = .Top + .MarkerSize
      End If
    End With
  Next iCount

  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function

Private Function MiddleAlignSelectedControls()

' Middle align the selected controls
  Dim iCount As Integer
  Dim lngTop As Long
  Dim lngMiddle As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngMiddle = .Top + (.Height / 2)
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngMiddle - (.Height / 2)
        .AttachedObject.Top = (lngMiddle - (.Height / 2)) + .MarkerSize
      End If
    End With
  Next iCount

  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function

Private Function BottomAlignSelectedControls()

' Bottom align the selected controls
  Dim iCount As Integer
  Dim lngBottom As Long

  'Find out the bottom most control - this is used as the align point
  lngBottom = 0
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top + .Height > lngBottom Then
        lngBottom = .Top + .Height
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngBottom - .Height
        .AttachedObject.Top = (lngBottom - .Height) + .MarkerSize
      End If
    End With
  Next iCount

  ' Mark the screen as having changed.
  mfChanged = True
  frmSysMgr.RefreshMenu

End Function

Public Property Get PictureLocation() As Long
  PictureLocation = mlngPictureLocation
End Property
Public Property Get WFCompletionMessageType() As MessageType
  WFCompletionMessageType = miCompletionMessageType
  
End Property

Public Property Let WFCompletionMessageType(ByVal piNewValue As MessageType)
  miCompletionMessageType = piNewValue

End Property

Public Property Get WFCompletionMessage() As String
  If miCompletionMessageType = MESSAGE_CUSTOM Then
    WFCompletionMessage = msCompletionMessage
  Else
    WFCompletionMessage = ""
  End If
  
End Property

Public Property Let WFCompletionMessage(ByVal psNewValue As String)
  msCompletionMessage = psNewValue

End Property

Public Property Get WFSavedForLaterMessageType() As MessageType
  WFSavedForLaterMessageType = miSavedForLaterMessageType

End Property

Public Property Let WFSavedForLaterMessageType(ByVal piNewValue As MessageType)
  miSavedForLaterMessageType = piNewValue

End Property

Public Property Get WFSavedForLaterMessage() As String
  If miSavedForLaterMessageType = MESSAGE_CUSTOM Then
    WFSavedForLaterMessage = msSavedForLaterMessage
  Else
    WFSavedForLaterMessage = ""
  End If

End Property

Public Property Let WFSavedForLaterMessage(ByVal psNewValue As String)
  msSavedForLaterMessage = psNewValue

End Property

Public Property Get WFFollowOnFormsMessageType() As MessageType
  WFFollowOnFormsMessageType = miFollowOnFormsMessageType

End Property

Public Property Let WFFollowOnFormsMessageType(ByVal piNewValue As MessageType)
  miFollowOnFormsMessageType = piNewValue

End Property

Public Property Get WFFollowOnFormsMessage() As String
  If miFollowOnFormsMessageType = MESSAGE_CUSTOM Then
    WFFollowOnFormsMessage = msFollowOnFormsMessage
  Else
    WFFollowOnFormsMessage = ""
  End If
  

End Property

Public Property Let WFFollowOnFormsMessage(ByVal psNewValue As String)
  msFollowOnFormsMessage = psNewValue

End Property

Public Property Let PictureLocation(plngNewValue As Long)
  mlngPictureLocation = plngNewValue
End Property
Public Property Get PictureID() As Long
  PictureID = mlngPictureID
End Property
Public Property Get TimeoutPeriod() As TimeoutPeriod
  TimeoutPeriod = miTimeoutPeriod
  
End Property

Public Property Get TimeoutFrequency() As Long
  TimeoutFrequency = mlngTimeoutFrequency
  
End Property

Public Property Let PictureID(plngNewValue As Long)
  mlngPictureID = plngNewValue
End Property

Public Property Let TimeoutPeriod(piNewValue As TimeoutPeriod)
  miTimeoutPeriod = piNewValue
  
End Property

Public Property Let TimeoutFrequency(plngNewValue As Long)
  mlngTimeoutFrequency = plngNewValue
  
End Property


Public Property Get Loading() As Boolean
  Loading = mfLoading
End Property
Public Property Let Loading(pbNewValue As Boolean)
  mfLoading = pbNewValue
End Property



Public Property Get DescriptionExprID() As Long
  DescriptionExprID = mlngDescriptionExprID
  
End Property

Public Property Let DescriptionExprID(ByVal plngNewValue As Long)
  mlngDescriptionExprID = plngNewValue
  
End Property


Public Property Get DescriptionHasWorkflowName() As Boolean
  DescriptionHasWorkflowName = mfDescriptionHasWorkflowName
  
End Property

Public Property Let DescriptionHasWorkflowName(ByVal pfNewValue As Boolean)
  mfDescriptionHasWorkflowName = pfNewValue
  
End Property

Public Property Get DescriptionHasElementCaption() As Boolean
  DescriptionHasElementCaption = mfDescriptionHasElementCaption
  
End Property

Public Property Let DescriptionHasElementCaption(ByVal pfNewValue As Boolean)
  mfDescriptionHasElementCaption = pfNewValue
  
End Property

Private Sub MoveAndPersistControls()

  Dim ctlControl As VB.Control
  Dim iWFItemType As WorkflowWebFormItemTypes

  ' Track the change in size to the form
  Dim lngOffsetX As Long: lngOffsetX = mlngLastFormWidth - Me.Width
  Dim lngOffsetY As Long: lngOffsetY = mlngLastFormheight - Me.Height

  Call DeselectAllControls

  ' Make sure all the required controls are selected
  For Each ctlControl In Me.Controls
    If IsWebFormControl(ctlControl) Then
      With ctlControl
        iWFItemType = CLng(.WFItemType)

        ' Select the anchored controls
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_VERTICALOFFSETBEHAVIOUR) _
          Or WebFormItemHasProperty(iWFItemType, WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR) Then

            If ctlControl.VerticalOffsetBehaviour <> offsetTop _
              Or ctlControl.HorizontalOffsetBehaviour <> offsetLeft Then
              
              ctlControl.Selected = True
              SelectControl ctlControl
              
            End If
        End If
        
        ' Select the persisted controls
        If WebFormItemHasProperty(iWFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) _
          Or WebFormItemHasProperty(iWFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then

            If ctlControl.HeightBehaviour <> behaveFixed _
              Or ctlControl.WidthBehaviour <> behaveFixed Then
              
              ctlControl.Selected = True
              SelectControl ctlControl
              
            End If
        End If
      End With
    End If
  Next

  ' Array to keep track of each objects offsets
  Dim arrOffset() As Integer
  ReDim arrOffset(ASRSelectionMarkers.Count - 1, 2)
  
  Dim iCount As Integer: iCount = 0
  
  ' Flag the selected selction markers to be moved
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    ASRSelectionMarkers(iCount).ShowSelectionMarkers False
    
    ' Set the images offset values
    arrOffset(iCount, 0) = lngOffsetX
    arrOffset(iCount, 1) = lngOffsetY
  Next iCount
  
  ' Track the co-ordinates and bools in variables to save time
  Dim lngTop As Long:  Dim lngLeft As Long
  Dim lngHeight As Long:  Dim lngWidth As Long
  Dim bHasVOffset As Boolean:  Dim bHasHOffset As Boolean
  Dim bHasHBehave As Boolean:  Dim bHasWBehave As Boolean

    ' Move the selected controls if the left button key is down, and the control is selected
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If IsWebFormControl(.AttachedObject) Then
        bHasVOffset = WebFormItemHasProperty(CLng(.AttachedObject.WFItemType), WFITEMPROP_VERTICALOFFSETBEHAVIOUR)
        bHasHOffset = WebFormItemHasProperty(CLng(.AttachedObject.WFItemType), WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR)
        bHasHBehave = WebFormItemHasProperty(CLng(.AttachedObject.WFItemType), WFITEMPROP_HEIGHTBEHAVIOUR)
        bHasWBehave = WebFormItemHasProperty(CLng(.AttachedObject.WFItemType), WFITEMPROP_WIDTHBEHAVIOUR)
        
        lngTop = .AttachedObject.Top
        lngLeft = .AttachedObject.Left
        lngHeight = .AttachedObject.Height
        lngWidth = .AttachedObject.Width
        
        If bHasVOffset Or bHasHBehave Then
          
          ' If its locked top then the MoveY offset = 0
          If Not (.AttachedObject.VerticalOffsetBehaviour = offsetTop) Then
            lngTop = lngTop - arrOffset(iCount, 1)
          End If
          
          ' If its locked left then the MoveX offset = 0
          If Not (.AttachedObject.HorizontalOffsetBehaviour = offsetLeft) Then
            lngLeft = lngLeft - arrOffset(iCount, 0)
          End If
          
        End If
        
        If bHasHBehave Or bHasWBehave Then
  
          ' If its height behaviour is set to full then height = form height
          If .AttachedObject.HeightBehaviour = behaveFull And .AttachedObject.WidthBehaviour = behaveFull Then
            lngTop = 0
            lngLeft = 0
            lngHeight = Me.ScaleHeight
            lngWidth = Me.ScaleWidth
          ElseIf .AttachedObject.HeightBehaviour = behaveFull Then
            lngTop = 0
            lngHeight = Me.ScaleHeight
          ElseIf .AttachedObject.WidthBehaviour = behaveFull Then
            lngLeft = 0
            lngWidth = Me.ScaleWidth
          End If
          
        End If
        
        ' Now lets do the move having worked out the dimensions
        .AttachedObject.Move lngLeft, lngTop, lngWidth, lngHeight
        
        ' End placementing of object
        .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize
        
        ' Deselect it and hey presto its done!
        DeselectControl .AttachedObject
        .AttachedObject.Selected = False
      End If
    End With

  Next iCount

  Call DeselectAllControls
  
  ' Flag screen as having changed
  If ASRSelectionMarkers.Count > 0 Then IsChanged = True
  
End Sub

Public Property Get TimeoutExcludeWeekend() As Boolean
  TimeoutExcludeWeekend = mfTimeoutExcludeWeekend
End Property

Public Property Let TimeoutExcludeWeekend(ByVal pfNewValue As Boolean)
  mfTimeoutExcludeWeekend = pfNewValue
End Property

Private Sub lblBlankDesigner_DblClick()
  Form_DblClick
  
End Sub

Private Sub lblBlankDesigner_DragDrop(Source As Control, X As Single, Y As Single)
  Form_DragDrop Source, X + lblBlankDesigner.Left, Y + lblBlankDesigner.Top
End Sub

Private Sub lblBlankDesigner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Form_MouseDown Button, Shift, X + lblBlankDesigner.Left, Y + lblBlankDesigner.Top
  
End Sub

Private Sub lblBlankDesigner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Form_MouseMove Button, Shift, X + lblBlankDesigner.Left, Y + lblBlankDesigner.Top

End Sub


Private Sub lblBlankDesigner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Form_MouseUp Button, Shift, X + lblBlankDesigner.Left, Y + lblBlankDesigner.Top

End Sub


' Default column width following font change to Verdana (Textbox)
Public Function Default_ColumnWidth_Textbox(ByRef plngColumnWidth As Long) As Long
  Default_ColumnWidth_Textbox = CLng(((plngColumnWidth + 1) * 95 + 105) / 10) * 10
End Function

' Default column width following font change to Verdana (Textbox)
Public Function Default_ColumnWidth_Numeric(ByRef plngNumeric As Long, ByRef plngDecimals As Long, ByRef pbSeperators As Boolean) As Long

  Dim lngSeperators As Long
  Dim lngWidth As Long

  lngSeperators = 60 * IIf(pbSeperators, plngNumeric / 3, 0)
  lngWidth = plngNumeric + IIf(plngDecimals > 0, plngDecimals + 1, 0) + 1

  Default_ColumnWidth_Numeric = (plngNumeric * 105) + 120 + 60 + lngSeperators
End Function


