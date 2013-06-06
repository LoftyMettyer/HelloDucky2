VERSION 5.00
Begin VB.Form frmWorkflowElementItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Element Item"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5056
   Icon            =   "frmWorkflowElementItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraItem 
      Caption         =   "File :"
      Height          =   850
      Index           =   18
      Left            =   2160
      TabIndex        =   33
      Top             =   2760
      Width           =   3200
      Begin VB.CommandButton cmdFileFile 
         Caption         =   "..."
         Height          =   315
         Left            =   2300
         TabIndex        =   36
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtFileFile 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label lblFileFile 
         Caption         =   "File :"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Calculation :"
      Height          =   850
      Index           =   16
      Left            =   2160
      TabIndex        =   29
      Top             =   1800
      Width           =   3200
      Begin VB.TextBox txtCalcCalculation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   300
         Width           =   1000
      End
      Begin VB.CommandButton cmdCalcCalculation 
         Caption         =   "..."
         Height          =   315
         Left            =   2300
         TabIndex        =   32
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Label lblCalcCalculation 
         Caption         =   "Calculation :"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Formatting :"
      Height          =   850
      Index           =   12
      Left            =   6000
      TabIndex        =   26
      Top             =   4560
      Width           =   3200
      Begin VB.ComboBox cboFormattingOption 
         Height          =   315
         ItemData        =   "frmWorkflowElementItem.frx":000C
         Left            =   1000
         List            =   "frmWorkflowElementItem.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   300
         Width           =   2000
      End
      Begin VB.Label lblFormattingOption 
         Caption         =   "Option :"
         Height          =   195
         Left            =   195
         TabIndex        =   27
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Workflow Value :"
      Height          =   1200
      Index           =   4
      Left            =   6000
      TabIndex        =   21
      Top             =   3240
      Width           =   3400
      Begin VB.ComboBox cboWFWebForm 
         Height          =   315
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   300
         Width           =   1950
      End
      Begin VB.ComboBox cboWFValue 
         Height          =   315
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   700
         Width           =   1950
      End
      Begin VB.Label lblWFWebForm 
         Caption         =   "Web Form :"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblWFValue 
         Caption         =   "Value :"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   765
         Width           =   720
      End
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   2160
      TabIndex        =   37
      Top             =   240
      Width           =   2600
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   39
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Database Value :"
      Height          =   2900
      Index           =   1
      Left            =   5500
      TabIndex        =   10
      Top             =   100
      Width           =   3700
      Begin VB.ComboBox cboDBValueRecordSelector 
         Height          =   315
         ItemData        =   "frmWorkflowElementItem.frx":0032
         Left            =   1770
         List            =   "frmWorkflowElementItem.frx":0034
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1900
         Width           =   1815
      End
      Begin VB.ComboBox cboDBValueWebForm 
         Height          =   315
         ItemData        =   "frmWorkflowElementItem.frx":0036
         Left            =   1770
         List            =   "frmWorkflowElementItem.frx":0038
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1500
         Width           =   1815
      End
      Begin VB.ComboBox cboDBValueRecord 
         Height          =   315
         ItemData        =   "frmWorkflowElementItem.frx":003A
         Left            =   1770
         List            =   "frmWorkflowElementItem.frx":003C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1100
         Width           =   1815
      End
      Begin VB.ComboBox cboDBValueColumn 
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   700
         Width           =   1815
      End
      Begin VB.ComboBox cboDBValueTable 
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblDBValueRecordSelector 
         Caption         =   "Record Selector :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   1965
         Width           =   1515
      End
      Begin VB.Label lblDBValueWebForm 
         Caption         =   "Element :"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblDBValueRecordID 
         Caption         =   "Record :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label lblDBValueColumn 
         Caption         =   "Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   765
         Width           =   900
      End
      Begin VB.Label lblDBValueTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Label :"
      Height          =   850
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   3200
      Begin VB.TextBox txtLabelCaption 
         Height          =   315
         Left            =   1050
         MaxLength       =   200
         TabIndex        =   9
         Top             =   300
         Width           =   1950
      End
      Begin VB.Label lblLabelCaption 
         Caption         =   "Caption :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame fraItemType 
      Caption         =   "Type :"
      Height          =   3000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   1900
      Begin VB.OptionButton optItemType 
         Caption         =   "F&ile"
         Height          =   315
         Index           =   18
         Left            =   105
         TabIndex        =   6
         Top             =   2050
         Width           =   915
      End
      Begin VB.OptionButton optItemType 
         Caption         =   "C&alculation"
         Height          =   315
         Index           =   16
         Left            =   105
         TabIndex        =   2
         Top             =   650
         Width           =   1410
      End
      Begin VB.OptionButton optItemType 
         Caption         =   "&Formatting"
         Height          =   315
         Index           =   12
         Left            =   105
         TabIndex        =   5
         Top             =   1700
         Width           =   1410
      End
      Begin VB.OptionButton optItemType 
         Caption         =   "&Workflow Value"
         Height          =   315
         Index           =   4
         Left            =   105
         TabIndex        =   4
         Top             =   1350
         Width           =   1770
      End
      Begin VB.OptionButton optItemType 
         Caption         =   "&Label"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   1
         Top             =   300
         Width           =   1020
      End
      Begin VB.OptionButton optItemType 
         Caption         =   "&Database Value"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1000
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmWorkflowElementItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private miItemType As WorkflowWebFormItemTypes
Private miElementType As ElementType
Private msCaption As String
Private mlngDBColumnID As Long
Private miDBRecord As Integer
Private msWFFormIdentifier As String
Private msWFValueIdentifier As String
Private msFormattingOption As String
Private mlngCalculationExprID As Long
Private msAttachmentFile As String

Private msDBWebForm As String
Private msDBRecordSelector As String

' Form handling variables.
Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfLoading As Boolean

Private mlngPersonnelTableID As Long
Private mlngBaseTableID As Long
Private miInitiationType As WorkflowInitiationTypes

Private mfrmCallingForm As Form
Private maWFPrecedingElements() As VB.Control
Private maWFAllElements() As VB.Control

Private mfInitializing As Boolean
Private msInitializeMessage As String
Private mfAttachmentSelection As Boolean

Public Property Let Changed(ByVal pfNewValue As Boolean)
  If Not mfLoading Then
    mfChanged = pfNewValue
    RefreshScreen
  End If
  
End Property

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
End Property

Public Sub Initialize(pfrmCallingForm As Form, _
  piElementType As ElementType, _
  piType As WorkflowEmailItemTypes, _
  psCaption As String, _
  plngDBColumnID As Long, _
  piDBRecord As Integer, _
  psWFForm As String, _
  psWFValue As String, _
  pfCopy As Boolean, _
  psDBRecordWebForm As String, _
  psDBRecordSelector As String, _
  plngCalculationExprID As Long, _
  pfAttachmentSelection As Boolean, _
  psAttachmentFile As String)
  
  mfLoading = True
  mfInitializing = True
  msInitializeMessage = ""
  mfAttachmentSelection = pfAttachmentSelection
  
  Set mfrmCallingForm = pfrmCallingForm
  miElementType = piElementType
  
  ReDim maWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements maWFPrecedingElements
  
  ReDim maWFAllElements(0)
  mfrmCallingForm.CallingForm.AllElements maWFAllElements
  
  FormatTypeFrame
  
  If piType = giWFEMAILITEM_FORMATCODE Then
    msFormattingOption = psCaption
    msCaption = ""
  Else
    msCaption = psCaption
    msFormattingOption = ""
  End If
  
  mlngDBColumnID = plngDBColumnID
  miDBRecord = piDBRecord
  msWFFormIdentifier = psWFForm
  msWFValueIdentifier = psWFValue
  msDBWebForm = psDBRecordWebForm
  msDBRecordSelector = psDBRecordSelector
  mlngBaseTableID = pfrmCallingForm.BaseTable
  miInitiationType = pfrmCallingForm.InitiationType
  mlngCalculationExprID = plngCalculationExprID
  msAttachmentFile = psAttachmentFile
  
  ItemType = piType
  
  mfChanged = ((mlngDBColumnID <> plngDBColumnID) _
    Or (miDBRecord <> piDBRecord) _
    Or (msWFFormIdentifier <> psWFForm) _
    Or (msWFValueIdentifier <> psWFValue) _
    Or (msDBWebForm <> psDBRecordWebForm) _
    Or (msDBRecordSelector <> psDBRecordSelector) _
    Or (mlngCalculationExprID <> plngCalculationExprID))
    
  mfLoading = False
  
  If Len(msInitializeMessage) > 0 Then
    MsgBox msInitializeMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  mfInitializing = False
  
  If pfCopy Then mfChanged = True
  RefreshScreen
    
    
End Sub

Private Sub FormatTypeFrame()
  Dim sngY As Single
  Const YGAP = 350
  
  sngY = 300
  
  If mfAttachmentSelection Then
    optItemType(giWFEMAILITEM_FILEATTACHMENT).Top = sngY
    sngY = sngY + YGAP
  End If
  optItemType(giWFEMAILITEM_FILEATTACHMENT).Visible = (mfAttachmentSelection)
  
  If Not mfAttachmentSelection Then
    optItemType(giWFEMAILITEM_LABEL).Top = sngY
    sngY = sngY + YGAP
  End If
  optItemType(giWFEMAILITEM_LABEL).Visible = (Not mfAttachmentSelection)
  
  If Not mfAttachmentSelection Then
    optItemType(giWFEMAILITEM_CALCULATION).Top = sngY
    sngY = sngY + YGAP
  End If
  optItemType(giWFEMAILITEM_CALCULATION).Visible = (Not mfAttachmentSelection)
  
  optItemType(giWFEMAILITEM_DBVALUE).Top = sngY
  sngY = sngY + YGAP
  
  optItemType(giWFEMAILITEM_WFVALUE).Top = sngY
  sngY = sngY + YGAP
  
  If Not mfAttachmentSelection Then
    optItemType(giWFEMAILITEM_FORMATCODE).Top = sngY
    sngY = sngY + YGAP
  End If
  optItemType(giWFEMAILITEM_FORMATCODE).Visible = (Not mfAttachmentSelection)
  
End Sub


Private Sub DisplayItemFrame()
  Dim fraTemp As Frame
  
  ' Initialize the displayed controls.
  InitializeItemControls
  
  ' Display only the frame that defines the selected component type.
  For Each fraTemp In fraItem
    fraTemp.Visible = (fraTemp.Index = miItemType)
  Next fraTemp
  Set fraTemp = Nothing
  
  RefreshScreen
  
End Sub

Private Sub InitializeItemControls()

  ' Call the required sub-routine to initialze the item
  ' definition controls.
  Select Case miItemType
    Case giWFEMAILITEM_DBVALUE
      InitializeDBValueControls
    
    Case giWFEMAILITEM_LABEL
      InitializeLabelControls
    
    Case giWFEMAILITEM_WFVALUE
      InitializeWFValueControls
      
    Case giWFEMAILITEM_FORMATCODE
      InitializeFormattingControls
      
    Case giWFEMAILITEM_CALCULATION
      InitializeCalculationControls
      
    Case giWFEMAILITEM_FILEATTACHMENT
      InitializeFileAttachmentControls
  End Select
  
End Sub

Private Sub InitializeDBValueControls()
  ' Initialize the Database Value item controls.
  cboDBValueTable_Refresh

  cboDBValueRecord_Refresh
  
End Sub

Private Sub cboFormattingOption_Refresh()
  ' Select the current formattingoption.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim sOption As String

  iIndex = 0
  sOption = msFormattingOption

  ' Loop through the available options.
  For iLoop = 0 To cboFormattingOption.ListCount - 1
    ' Select the current return type if it is in the combo's list.
    Select Case cboFormattingOption.ItemData(iLoop)
      Case 1
        If sOption = "L" Then
          iIndex = iLoop
          Exit For
        End If
  
      Case 2
        If sOption = "N" Then
          iIndex = iLoop
          Exit For
        End If
  
      Case 3
        If sOption = "T" Then
        iIndex = iLoop
        Exit For
      End If
    End Select
  Next iLoop

  cboFormattingOption.ListIndex = iIndex
  
End Sub


Private Sub InitializeFormattingControls()
  ' Initialize the Formatting item controls.
  cboFormattingOption_Refresh
  
End Sub


Private Sub InitializeLabelControls()
  ' Initialize the Label item controls.
  txtLabelCaption.Text = msCaption
End Sub

Private Sub InitializeFileAttachmentControls()
  ' Initialize the FileAttachment item controls.
  txtFileFile.Text = msAttachmentFile
  
End Sub



Private Sub InitializeCalculationControls()
  ' Initialize the Calculation item controls.
  txtCalcCalculation.Text = GetExpressionName(mlngCalculationExprID)
  
End Sub


Private Sub InitializeWFValueControls()
  ' Initialize the Workflow Value item controls.
  cboWFWebForm_Refresh
End Sub

Private Sub RefreshScreen()
  Dim fEnableOK As Boolean
  Dim fEnableDBValueRecord As Boolean
  Dim fEnableDBValueWebForm As Boolean
  Dim fEnableDBValueRecordSelector As Boolean
  Dim wfTemp As VB.Control
  
  fEnableOK = mfChanged
  
  Select Case miItemType
    Case giWFEMAILITEM_DBVALUE
      fEnableDBValueRecord = False
      If cboDBValueRecord.ListIndex >= 0 Then
        fEnableDBValueRecord = (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_UNKNOWN)
      End If
      
      fEnableDBValueWebForm = False
      If cboDBValueRecord.ListIndex >= 0 Then
        If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
          And (cboDBValueWebForm.ListCount > 0) _
          And (cboDBValueWebForm.ListIndex >= 0) Then
        
          fEnableDBValueWebForm = (cboDBValueWebForm.ItemData(cboDBValueWebForm.ListIndex) > 0)
        End If
      End If
      
      fEnableDBValueRecordSelector = False
      If cboDBValueRecord.ListIndex >= 0 Then
        If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
          And (cboDBValueRecordSelector.ListCount > 0) _
          And (cboDBValueRecordSelector.ListIndex >= 0) Then
        
         fEnableDBValueRecordSelector = (cboDBValueRecordSelector.ItemData(cboDBValueRecordSelector.ListIndex) > 0)
          If fEnableDBValueRecordSelector Then
            Set wfTemp = SelectedElement
            If Not wfTemp Is Nothing Then
              fEnableDBValueRecordSelector = (wfTemp.ElementType = elem_WebForm)
            End If
          End If
        End If
      End If
      
      cboDBValueRecord.Enabled = fEnableDBValueRecord
      cboDBValueRecord.BackColor = IIf(fEnableDBValueRecord, vbWindowBackground, vbButtonFace)
      lblDBValueRecordID.Enabled = fEnableDBValueRecord
      
      cboDBValueWebForm.Enabled = fEnableDBValueWebForm
      cboDBValueWebForm.BackColor = IIf(fEnableDBValueWebForm, vbWindowBackground, vbButtonFace)
      lblDBValueWebForm.Enabled = fEnableDBValueWebForm
      
      cboDBValueRecordSelector.Enabled = fEnableDBValueRecordSelector
      cboDBValueRecordSelector.BackColor = IIf(fEnableDBValueRecordSelector, vbWindowBackground, vbButtonFace)
      lblDBValueRecordSelector.Enabled = fEnableDBValueRecordSelector
      
      fEnableOK = fEnableOK And cboDBValueColumn.Enabled And fEnableDBValueRecord
      If fEnableOK Then
        If cboDBValueRecord.ListIndex >= 0 Then
          If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
            fEnableOK = fEnableOK And fEnableDBValueWebForm
            
            If fEnableOK Then
              Set wfTemp = SelectedElement
              If Not wfTemp Is Nothing Then
                If (wfTemp.ElementType = elem_WebForm) Then
                  fEnableOK = fEnableOK And fEnableDBValueRecordSelector
                End If
              End If
            End If
          End If
        End If
      End If
      
    Case giWFEMAILITEM_LABEL
    
    Case giWFEMAILITEM_WFVALUE
      fEnableOK = fEnableOK And cboWFValue.Enabled
  
    Case giWFEMAILITEM_FORMATCODE
  
    Case giWFEMAILITEM_CALCULATION
      fEnableOK = fEnableOK And (mlngCalculationExprID > 0)
  
    Case giWFEMAILITEM_FILEATTACHMENT
      fEnableOK = fEnableOK And (Len(msAttachmentFile) > 0)
  End Select
  
  cmdOk.Enabled = fEnableOK
  
End Sub

Private Function SelectedElement() As VB.Control
  ' Return the element that has been selected for the DBValue record.
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  
  If cboDBValueWebForm.ListIndex >= 0 Then
    For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
      Set wfTemp = maWFPrecedingElements(lngLoop)

      If wfTemp.ControlIndex = cboDBValueWebForm.ItemData(cboDBValueWebForm.ListIndex) Then
        Set SelectedElement = wfTemp
        Exit For
      End If
      
      Set wfTemp = Nothing
    Next lngLoop
  End If
  
End Function

Private Function ValidateItem() As Boolean
  ' Most validation is done using the disabling of the OK button.
  ' Only validations that require some message displayed are done here.
  
  ValidateItem = True
  
End Function

Private Sub cboDBValueColumn_Click()
  If cboDBValueColumn.Enabled Then
    mlngDBColumnID = cboDBValueColumn.ItemData(cboDBValueColumn.ListIndex)
    
    Changed = True
  End If
  
End Sub

Private Sub cboDBValueRecord_Click()
  If cboDBValueColumn.Enabled Then
    miDBRecord = cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex)
    
    cboDBValueWebForm_Refresh
    cboDBValueRecordSelector_Refresh
    
    Changed = True
  End If

End Sub

Private Sub cboDBValueWebForm_Refresh()
  ' Populate the DB Element combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim fWebFormWithSelector As Boolean
  Dim asItems() As String
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngTableID As Long
  
  With cboDBValueWebForm
    ' Clear the current contents of the combo.
    .Clear

    If cboDBValueRecord.ListIndex >= 0 Then
    
      
      lngTableID = -1
      If cboDBValueTable.ListIndex >= 0 Then
        lngTableID = cboDBValueTable.ItemData(cboDBValueTable.ListIndex)
      End If
    
      If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_INITIATOR) _
        And (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_TRIGGEREDRECORD) Then
        
        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
          fWebFormWithSelector = False
          Set wfTemp = maWFPrecedingElements(lngLoop)
  
          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items
  
            For lngLoop2 = 1 To UBound(asItems, 2)
              If (asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
                
                fFound = False
                For lngLoop3 = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop3) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop3
                
                If fFound Then
                  fWebFormWithSelector = True
                  Exit For
                End If
              End If
            Next lngLoop2
            
          ElseIf wfTemp.ElementType = elem_StoredData Then
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables
            
            'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
            'If wfTemp.DataAction = DATAACTION_DELETE Then
            '  ' Cannot do anything with a Deleted record, but can use its ascendants.
            '  ' Remove the table itself from the array of valid tables.
            '  alngValidTables(1) = 0
            'End If
            
            fFound = False
            For lngLoop3 = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop3) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop3
            
            If fFound Then
              fWebFormWithSelector = True
            End If
          End If
  
          If fWebFormWithSelector Then
            .AddItem wfTemp.Identifier
            .ItemData(.NewIndex) = wfTemp.ControlIndex
          End If
  
          Set wfTemp = Nothing
        Next lngLoop
      End If
    End If
    
    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = msDBWebForm Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(msDBWebForm)) > 0) Then
        sMsg = "The previously selected Database Value Element is no longer valid."
  
        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Database Value Element has been selected."
        End If
  
        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If
        
        mfChanged = True
      End If
      
      iIndex = 0
    End If
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      If cboDBValueRecord.ListIndex >= 0 Then
        If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_INITIATOR) _
          And (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_TRIGGEREDRECORD) Then
          .Enabled = False
    
          .AddItem "<no values>"
          .ItemData(.NewIndex) = 0
          .ListIndex = 0
        Else
          msDBWebForm = ""
          msDBRecordSelector = ""
        End If
      End If
    End If
  End With
    
End Sub


Private Sub cboDBValueRecordSelector_Refresh()
  ' Populate the DBValue RecordSelector combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim iElementType As ElementType
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngTableID As Long
  
  With cboDBValueRecordSelector
    ' Clear the current contents of the combo.
    .Clear

    If cboDBValueWebForm.ListIndex >= 0 Then
      
      lngTableID = -1
      If cboDBValueTable.ListIndex >= 0 Then
        lngTableID = cboDBValueTable.ItemData(cboDBValueTable.ListIndex)
      End If
      
      For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
        Set wfTemp = maWFPrecedingElements(lngLoop)

        If wfTemp.ControlIndex = cboDBValueWebForm.ItemData(cboDBValueWebForm.ListIndex) Then
          iElementType = wfTemp.ElementType

          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items
  
            For lngLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
                
                fFound = False
                For lngLoop3 = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop3) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop3
                
                If fFound Then
                  'JPD 20061010 Fault 11355
                  '.AddItem asItems(1, lngLoop2)
                  .AddItem asItems(9, lngLoop2)
                  .ItemData(.NewIndex) = 1
                End If
              End If
            Next lngLoop2
          End If
          
          Exit For
        End If

        Set wfTemp = Nothing
      Next lngLoop
    End If

    If iElementType <> elem_WebForm Then
      msDBRecordSelector = ""
    End If

    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = msDBRecordSelector Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(msDBRecordSelector)) > 0) Then
        sMsg = "The previously selected Database Value Record Selector is no longer valid."
        
        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Database Value Record Selector has been selected."
        End If
        
        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If
        
        mfChanged = True
      End If
      
      iIndex = 0
    End If
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      If cboDBValueRecord.ListIndex >= 0 Then
        If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_INITIATOR) _
          And (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) <> giWFRECSEL_TRIGGEREDRECORD) Then
          
          .Enabled = False

          If iElementType = elem_WebForm Then
            .AddItem "<no values>"
            .ItemData(.NewIndex) = 0
            .ListIndex = 0
          End If
        Else
          msDBRecordSelector = ""
        End If
      End If
    End If
  End With
    
End Sub



Private Sub cboDBValueRecordSelector_Click()
  msDBRecordSelector = cboDBValueRecordSelector.List(cboDBValueRecordSelector.ListIndex)
  Changed = True

End Sub


Private Sub cboDBValueTable_Click()
  ' Populate the field combo with the relevant fields.
  cboDBValueColumn_Refresh

  cboDBValueRecord_Refresh
  cboDBValueWebForm_Refresh
  cboDBValueRecordSelector_Refresh

End Sub

Private Sub cboDBValueWebForm_Click()
  If miDBRecord = giWFRECSEL_IDENTIFIEDRECORD Then
    msDBWebForm = cboDBValueWebForm.List(cboDBValueWebForm.ListIndex)
  Else
    msDBWebForm = ""
  End If
  
  Changed = True
  cboDBValueRecordSelector_Refresh

End Sub


Private Sub cboFormattingOption_Click()
  Select Case cboFormattingOption.ItemData(cboFormattingOption.ListIndex)
    Case 1
      ' Line
      msFormattingOption = "L"
    Case 2
      ' New line
      msFormattingOption = "N"
    Case Else
      ' 3 = Tab
      msFormattingOption = "T"
  End Select
  
  Changed = True

End Sub


Private Sub cboWFValue_Click()
  msWFValueIdentifier = cboWFValue.List(cboWFValue.ListIndex)
  
  Changed = True

End Sub


Private Sub cboWFWebForm_Click()
  ' Populate the field combo with the relevant fields.
  msWFFormIdentifier = cboWFWebForm.List(cboWFWebForm.ListIndex)
  
  Changed = True
  
  cboWFValue_Refresh
  
End Sub


Private Sub cmdCalcCalculation_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long

  lngOriginalID = mlngCalculationExprID

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_CHARACTER
    .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
    .UtilityBaseTable = mfrmCallingForm.CallingForm.BaseTable
    .WorkflowInitiationType = mfrmCallingForm.CallingForm.InitiationType
    .PrecedingWorkflowElements = maWFPrecedingElements
    .AllWorkflowElements = maWFAllElements

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(mfrmCallingForm.CallingForm.ReadOnly) Then
      mlngCalculationExprID = .ExpressionID
    Else
      ' Check in case the original expression has been deleted.
      If Not CheckExpression(mlngCalculationExprID, 0, False) Then
        mlngCalculationExprID = 0
      End If
    End If

    ' Read the selected expression info.
    txtCalcCalculation.Text = GetExpressionName(mlngCalculationExprID)
  End With

  Set objExpr = Nothing

  If lngOriginalID <> mlngCalculationExprID Then
    Changed = True
  End If
  
End Sub

Private Sub cmdCancel_Click()
  ' Set the cancelled flag.
  mfCancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdFileFile_Click()
  Dim frmFileSel As frmEmailLinkAttachmentSel

  Set frmFileSel = New frmEmailLinkAttachmentSel

  If Trim(gstrEmailAttachmentPath) = vbNullString Then
    MsgBox "You will need to set up an email path in configuration prior to adding email attachments", vbExclamation, "Email Link"
    Exit Sub
  End If

  With frmFileSel
    .Show vbModal
    If .Cancelled = False Then
      txtFileFile.Text = .FileName
      msAttachmentFile = txtFileFile.Text
      Changed = True
    End If
  End With

  Set frmFileSel = Nothing

End Sub

Private Sub cmdOK_Click()
  
  If ValidateItem Then
    ' Set the cancelled flag.
    mfCancelled = False
    
    ' Unload the form.
    UnLoad Me
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  
  fraOKCancel.BorderStyle = vbBSNone
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)

  FormatScreen

End Sub

Private Sub FormatScreen()
  ' Position and size controls.
  Dim fraTemp As Frame
  
  Const iXGAP = 200
  Const iYGAP = 200
  Const iXFRAMEGAP = 150
  Const iYFRAMEGAP = 100
  Const iITEMFRAMEWIDTH = 1900
  Const iFRAMEWIDTH = 5700
  Const iFRAMEHEIGHT = 2650
  
  ' Position and size the item type frame.
  With fraItemType
    .Left = iXFRAMEGAP
    .Top = iYFRAMEGAP
    .Width = iITEMFRAMEWIDTH
    .Height = iFRAMEHEIGHT
  End With
  
  ' Position and size the item definition frames.
  For Each fraTemp In fraItem
    With fraTemp
      .Left = fraItemType.Left + iITEMFRAMEWIDTH + iXFRAMEGAP
      .Top = iYFRAMEGAP
      .Width = iFRAMEWIDTH
      .Height = iFRAMEHEIGHT
    End With
  Next fraTemp
  Set fraTemp = Nothing

  ' Format the controls within the frames.
  FormatDBValueFrame
  FormatLabelFrame
  FormatWFValueFrame
  FormatFormattingOptionFrame
  FormatCalculationFrame
  FormatFileFrame

  ' Position and size the OK/Cancel command controls.
  With fraOKCancel
    .Top = iYFRAMEGAP + iYGAP + iFRAMEHEIGHT
    .Left = fraItem(fraItem.LBound).Left + _
      fraItem(fraItem.LBound).Width - .Width
  End With

  ' Size the form.
  Me.Width = fraItem(fraItem.UBound).Left + _
    fraItem(fraItem.UBound).Width + iXFRAMEGAP + _
   (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
  Me.Height = fraOKCancel.Top + fraOKCancel.Height + iXFRAMEGAP + _
    (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
  
End Sub


Private Sub FormatDBValueFrame()
  ' Size and position the Database Value item controls.
  Const iXGAP = 200

  cboDBValueTable.Width = fraItem(giWFEMAILITEM_DBVALUE).Width - cboDBValueTable.Left - iXGAP
  cboDBValueColumn.Width = fraItem(giWFEMAILITEM_DBVALUE).Width - cboDBValueColumn.Left - iXGAP
  cboDBValueRecord.Width = fraItem(giWFEMAILITEM_DBVALUE).Width - cboDBValueRecord.Left - iXGAP
  cboDBValueWebForm.Width = fraItem(giWFEMAILITEM_DBVALUE).Width - cboDBValueWebForm.Left - iXGAP
  cboDBValueRecordSelector.Width = fraItem(giWFEMAILITEM_DBVALUE).Width - cboDBValueRecordSelector.Left - iXGAP

End Sub

Private Sub FormatWFValueFrame()
  ' Size and position the Workflow value item controls.
  Const iXGAP = 200

  cboWFWebForm.Width = fraItem(giWFEMAILITEM_WFVALUE).Width - cboWFWebForm.Left - iXGAP
  cboWFValue.Width = fraItem(giWFEMAILITEM_WFVALUE).Width - cboWFValue.Left - iXGAP
End Sub

Private Sub cboDBValueColumn_Refresh()
  ' Populate the DB Value Column combo and
  ' select the current column if it is still valid.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim fColumnOK As Boolean

  iIndex = -1

  lngTableID = 0
  If cboDBValueTable.ListIndex >= 0 Then
    lngTableID = cboDBValueTable.ItemData(cboDBValueTable.ListIndex)
  End If
  
  ' Clear the current contents of the combo.
  cboDBValueColumn.Clear

  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngTableID

    If Not .NoMatch Then
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If

      ' Add  an item to the combo for each table that has not been deleted.
      Do While Not .EOF
        ' Do not allow the user to select system columns, deleted columns, or
        ' OLE or Photo type columns.
        If (!TableID = lngTableID) And _
          (!Deleted = False) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If mfAttachmentSelection Then
            fColumnOK = ((!DataType = rdTypeLONGVARBINARY) _
              Or (!DataType = rdTypeVARBINARY))
          Else
            fColumnOK = ((!DataType <> rdTypeLONGVARBINARY) _
              And (!DataType <> rdTypeVARBINARY))
          End If
          
          If fColumnOK Then
            cboDBValueColumn.AddItem .Fields("columnName")
            cboDBValueColumn.ItemData(cboDBValueColumn.NewIndex) = .Fields("columnID")
          
            If .Fields("columnID") = mlngDBColumnID Then
              iIndex = cboDBValueColumn.NewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Enable the combo if there are items.
  With cboDBValueColumn
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no columns>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
    End If
  End With
    
End Sub

Private Sub cboWFValue_Refresh()
  ' Populate the WF Value combo and
  ' select the current value if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sMsg As String
  Dim fValueOK As Boolean

  iIndex = -1

  ' Clear the current contents of the combo.
  cboWFValue.Clear

  If cboWFWebForm.Enabled Then
    ' Add  an item to the combo for each input item in the preceding web form.
    Set wfTemp = maWFPrecedingElements(cboWFWebForm.ItemData(cboWFWebForm.ListIndex))

    asItems = wfTemp.Items

    For iLoop = 1 To UBound(asItems, 2)
      If mfAttachmentSelection Then
        fValueOK = (asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)
      Else
        fValueOK = (asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_CHAR Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_DATE Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_LOGIC Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_NUMERIC Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_DROPDOWN Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_LOOKUP Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP)
      End If
      
      If fValueOK Then
        cboWFValue.AddItem asItems(9, iLoop)
      End If
    Next iLoop
  End If

  For iLoop = 0 To cboWFValue.ListCount - 1
    If cboWFValue.List(iLoop) = msWFValueIdentifier Then
      iIndex = iLoop
    End If
  Next iLoop

  If (iIndex < 0) Then
    If (Len(Trim(msWFValueIdentifier)) > 0) Then
      sMsg = "The previously selected Workflow Value identifier is no longer valid."
      
      If cboWFValue.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value identifier has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
        
      mfChanged = True
    End If
    
    iIndex = 0
  End If
  
  ' Enable the combo if there are items.
  With cboWFValue
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no values>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
    End If
  End With

End Sub

Private Sub cboDBValueRecord_Refresh()
  ' Populate the DB Value Record combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim fWebFormWithSelector As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim lngTableID As Long
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  
  lngTableID = -1
  If cboDBValueTable.ListIndex >= 0 Then
    lngTableID = cboDBValueTable.ItemData(cboDBValueTable.ListIndex)
  End If
  
  With cboDBValueRecord
    ' Clear the current contents of the combo.
    .Clear

    For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
      fWebFormWithSelector = False
      Set wfTemp = maWFPrecedingElements(lngLoop)
        
      If wfTemp.ElementType = elem_WebForm Then
        asItems = wfTemp.Items
          
        For lngLoop2 = 1 To UBound(asItems, 2)
          If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
            ReDim alngValidTables(0)
            TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
            
            fFound = False
            For lngLoop3 = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop3) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop3
            
            If fFound Then
              fWebFormWithSelector = True
              Exit For
            End If
          End If
        Next lngLoop2
      ElseIf wfTemp.ElementType = elem_StoredData Then
        ReDim alngValidTables(0)
        TableAscendants wfTemp.DataTableID, alngValidTables
        
        'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
        'If wfTemp.DataAction = DATAACTION_DELETE Then
        '  ' Cannot do anything with a Deleted record, but can use its ascendants.
        '  ' Remove the table itself from the array of valid tables.
        '  alngValidTables(1) = 0
        'End If
      
        fFound = False
        For lngLoop3 = 1 To UBound(alngValidTables)
          If alngValidTables(lngLoop3) = lngTableID Then
            fFound = True
            Exit For
          End If
        Next lngLoop3
        
        If fFound Then
          fWebFormWithSelector = True
        End If
      End If
        
      If fWebFormWithSelector Then
        Exit For
      End If
        
      Set wfTemp = Nothing
    Next lngLoop

    If fWebFormWithSelector Then
      .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
      .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
    End If

    If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
      ReDim alngValidTables(0)
      TableAscendants mlngPersonnelTableID, alngValidTables
      
      fFound = False
      If mlngPersonnelTableID > 0 Then
        For lngLoop3 = 1 To UBound(alngValidTables)
          If alngValidTables(lngLoop3) = lngTableID Then
            fFound = True
            Exit For
          End If
        Next lngLoop3
      End If
      
      If fFound Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
        .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
      End If
    End If
    
    If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
      fFound = False
      If mlngBaseTableID > 0 Then
        ReDim alngValidTables(0)
        TableAscendants mlngBaseTableID, alngValidTables
        
        For lngLoop3 = 1 To UBound(alngValidTables)
          If alngValidTables(lngLoop3) = lngTableID Then
            fFound = True
            Exit For
          End If
        Next lngLoop3
      End If
      
      If fFound Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
        .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
      End If
    End If
    
    iIndex = -1
    iDefaultIndex = 0
    For lngLoop = 0 To .ListCount - 1
      If .ItemData(lngLoop) = miDBRecord Then
        iIndex = lngLoop
        Exit For
      End If
    
      If (.ItemData(lngLoop) = giWFRECSEL_INITIATOR) _
        Or (.ItemData(lngLoop) = giWFRECSEL_TRIGGEREDRECORD) Then
        iDefaultIndex = lngLoop
      End If
    Next lngLoop

    ' Enable the combo if there are items.
    If .ListCount > 0 Then
      .Enabled = True
      
      If iIndex < 0 Then
        iIndex = iDefaultIndex
        mfChanged = True
      End If
      
      .ListIndex = iIndex
    Else
      .AddItem "<no values>"
      .ItemData(.NewIndex) = giWFRECSEL_UNKNOWN
      .ListIndex = 0
    End If
  End With
    
End Sub



Private Sub cboDBValueTable_Refresh()
  ' Populate the DB Value Table combo and
  ' select the current table if it is still valid.
  Dim fTableOK As Boolean
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  
  iIndex = -1
  iDefaultIndex = -1
  
  ' Get the table of the selected column.
  lngTableID = 0
  If mlngDBColumnID > 0 Then
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", mlngDBColumnID

      If Not .NoMatch Then
        lngTableID = !TableID
      End If
    End With
  End If
  
  ' Clear the current contents of the combo.
  cboDBValueTable.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combo for each table that has not been deleted.
    Do While Not .EOF
      fTableOK = False

      If (Not .Fields("deleted")) Then
        cboDBValueTable.AddItem !TableName
        cboDBValueTable.ItemData(cboDBValueTable.NewIndex) = !TableID

        If !TableID = lngTableID Then
          iIndex = cboDBValueTable.NewIndex
        End If

        If ((miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) And (!TableID = mlngPersonnelTableID)) _
          Or ((miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) And (!TableID = mlngBaseTableID)) Then
          iDefaultIndex = cboDBValueTable.NewIndex
        End If
      End If

      .MoveNext
    Loop
  End With

  ' Enable the combo if there are items.
  With cboDBValueTable
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        If iDefaultIndex >= 0 Then
          iIndex = iDefaultIndex
        Else
          iIndex = 0
        End If
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no tables>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
    
      cboDBValueColumn_Refresh
    End If
  End With
    
End Sub

Private Sub cboWFWebForm_Refresh()
  ' Populate the WF WebForm combo and
  ' select the current webform if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim sMsg As String
  
  iIndex = -1

  ' Clear the current contents of the combo.
  cboWFWebForm.Clear

  ' Add  an item to the combo for each preceding web form.
  For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
    If maWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
      cboWFWebForm.AddItem maWFPrecedingElements(iLoop).Identifier
      cboWFWebForm.ItemData(cboWFWebForm.NewIndex) = iLoop
    End If
  Next iLoop

  For iLoop = 0 To cboWFWebForm.ListCount - 1
    If cboWFWebForm.List(iLoop) = msWFFormIdentifier Then
      iIndex = iLoop
    End If
  Next iLoop
  
  If (iIndex < 0) Then
    If (Len(Trim(msWFFormIdentifier)) > 0) Then
      sMsg = "The previously selected Workflow Value Web Form is no longer valid."
  
      If cboWFWebForm.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value Web Form has been selected."
      End If
  
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
      
      mfChanged = True
    End If
    
    iIndex = 0
  End If
  
  ' Enable the combo if there are items.
  With cboWFWebForm
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no preceding web forms>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
      
      cboWFValue_Refresh
    End If
  End With
    
End Sub

Private Sub FormatLabelFrame()
  ' Size and position the Labelitem controls.
  Const iXGAP = 200
  
  txtLabelCaption.Width = fraItem(giWFEMAILITEM_LABEL).Width - txtLabelCaption.Left - iXGAP

End Sub


Private Sub FormatCalculationFrame()
  ' Size and position the Calculation item controls.
  Const iXGAP = 200
  
  txtCalcCalculation.Width = fraItem(giWFEMAILITEM_CALCULATION).Width _
    - txtCalcCalculation.Left _
    - iXGAP _
    - cmdCalcCalculation.Width
  cmdCalcCalculation.Left = txtCalcCalculation.Left _
    + txtCalcCalculation.Width

End Sub



Private Sub FormatFileFrame()
  ' Size and position the File item controls.
  Const iXGAP = 200
  
  txtFileFile.Width = fraItem(giWFEMAILITEM_FILEATTACHMENT).Width _
    - txtFileFile.Left _
    - iXGAP _
    - cmdFileFile.Width
  cmdFileFile.Left = txtFileFile.Left _
    + txtFileFile.Width

End Sub





Private Sub FormatFormattingOptionFrame()
  ' Size and position the Formatting Option controls.
  Const iXGAP = 200
  
  cboFormattingOption.Width = fraItem(giWFEMAILITEM_FORMATCODE).Width - cboFormattingOption.Left - iXGAP

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged And cmdOk.Enabled Then
      iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If iAnswer = vbYes Then
        Call cmdOK_Click
        If Me.Cancelled Then Cancel = 1
      ElseIf iAnswer = vbNo Then
        mfCancelled = True
      ElseIf iAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      mfCancelled = True
    End If
  End If

End Sub

Private Sub optItemType_Click(Index As Integer)
  ' Set the component type property.
  miItemType = Index
  
  DisplayItemFrame

  Changed = True

End Sub



Public Property Get ItemType() As WorkflowWebFormItemTypes
  ItemType = miItemType
  
End Property

Public Property Let ItemType(ByVal piNewValue As WorkflowWebFormItemTypes)
  miItemType = piNewValue
  
  If miItemType = giWFEMAILITEM_LABEL Then
    optItemType_Click (miItemType)
  Else
    optItemType(miItemType).value = True
  End If
  
End Property

Public Property Get ItemCaption() As String
  If ItemType = giWFEMAILITEM_FORMATCODE Then
    ItemCaption = msFormattingOption
  Else
    ItemCaption = msCaption
  End If
  
End Property

Public Property Let ItemCaption(ByVal psNewValue As String)
  If ItemType = giWFEMAILITEM_FORMATCODE Then
    msFormattingOption = psNewValue
  Else
    msCaption = psNewValue
  End If
    
End Property

Private Sub txtLabelCaption_Change()
  Changed = True
  msCaption = txtLabelCaption.Text

End Sub


Private Sub txtLabelCaption_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub



Public Property Get ItemDBColumnID() As Long
  If miItemType = giWFEMAILITEM_DBVALUE Then
    ItemDBColumnID = mlngDBColumnID
  Else
    ItemDBColumnID = 0
  End If
    
End Property

Public Property Let ItemDBColumnID(ByVal plngNewValue As Long)
  mlngDBColumnID = plngNewValue
  
End Property




Public Property Get ItemDBRecord() As Integer
  If miItemType = giWFEMAILITEM_DBVALUE Then
    ItemDBRecord = miDBRecord
  Else
    ItemDBRecord = 0
  End If
  
End Property

Public Property Let ItemDBRecord(ByVal piNewValue As Integer)
  miDBRecord = piNewValue
  
End Property

Public Property Let ItemWFFormIdentifier(ByVal psNewValue As String)
  msWFFormIdentifier = psNewValue
  
End Property

Public Property Let ItemDBRecordSelector(ByVal psNewValue As String)
  msDBRecordSelector = psNewValue
  
End Property


Public Property Let ItemDBWebForm(ByVal psNewValue As String)
  msDBWebForm = psNewValue
  
End Property



Public Property Let ItemWFValueIdentifier(ByVal psNewValue As String)
  msWFValueIdentifier = psNewValue
  
End Property

Public Property Get ItemWFFormIdentifier() As String
  ItemWFFormIdentifier = msWFFormIdentifier
  
End Property

Public Property Get ItemDBWebForm() As String
  If miItemType = giWFEMAILITEM_DBVALUE _
    And miDBRecord = 1 Then
    
    ItemDBWebForm = msDBWebForm
  Else
    ItemDBWebForm = ""
  End If

End Property


Public Property Get ItemDBRecordSelector() As String
  If miItemType = giWFEMAILITEM_DBVALUE _
    And miDBRecord = 1 Then
    
    ItemDBRecordSelector = msDBRecordSelector
  Else
    ItemDBRecordSelector = ""
  End If
  
End Property



Public Property Get ItemWFValueIdentifier() As String
  ItemWFValueIdentifier = msWFValueIdentifier
  
End Property



Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationExprID
  
End Property

Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationExprID = plngNewValue

End Property
Private Function CheckExpression(plngExprID As Long, _
  plngTableID As Long, _
  pfCheckTable As Boolean) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If pfCheckTable And (plngTableID <= 0) Then
    fOK = False
  Else
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", plngExprID, False

      If .NoMatch Then
        fOK = False
      Else
        If pfCheckTable _
          And !TableID <> plngTableID Then
          
          fOK = False
        End If
      End If
    End With
  End If
  
TidyUpAndExit:
  CheckExpression = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function






Public Property Get FileAttachment() As String
  FileAttachment = msAttachmentFile

End Property

Public Property Let FileAttachment(ByVal psNewValue As String)
  msAttachmentFile = psNewValue
  
End Property

