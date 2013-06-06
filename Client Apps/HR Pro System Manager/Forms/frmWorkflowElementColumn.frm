VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmWorkflowElementColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Element Column"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1057
   Icon            =   "frmWorkflowElementColumn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefinition 
      Caption         =   "Calculation :"
      Height          =   850
      Index           =   3
      Left            =   4680
      TabIndex        =   34
      Top             =   4440
      Width           =   3990
      Begin VB.CommandButton cmdCalcCalculation 
         Height          =   315
         Left            =   3480
         Picture         =   "frmWorkflowElementColumn.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtCalcCalculation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label lblCalcCalculation 
         Caption         =   "Calculation :"
         Height          =   195
         Left            =   195
         TabIndex        =   37
         Top             =   360
         Width           =   1290
      End
   End
   Begin VB.Frame fraDefinition 
      Caption         =   "Database Value :"
      Height          =   2400
      Index           =   2
      Left            =   4700
      TabIndex        =   20
      Top             =   1900
      Width           =   4020
      Begin VB.ComboBox cboDBValueTable 
         Height          =   315
         Left            =   1725
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   2085
      End
      Begin VB.ComboBox cboDBValueColumn 
         Height          =   315
         Left            =   1725
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   700
         Width           =   2085
      End
      Begin VB.ComboBox cboDBValueRecord 
         Height          =   315
         ItemData        =   "frmWorkflowElementColumn.frx":015A
         Left            =   1725
         List            =   "frmWorkflowElementColumn.frx":015C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1100
         Width           =   2085
      End
      Begin VB.ComboBox cboDBValueWebForm 
         Height          =   315
         ItemData        =   "frmWorkflowElementColumn.frx":015E
         Left            =   1725
         List            =   "frmWorkflowElementColumn.frx":0160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1500
         Width           =   2085
      End
      Begin VB.ComboBox cboDBValueRecordSelector 
         Height          =   315
         ItemData        =   "frmWorkflowElementColumn.frx":0162
         Left            =   1725
         List            =   "frmWorkflowElementColumn.frx":0164
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1900
         Width           =   2085
      End
      Begin VB.Label lblDBValueTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblDBValueColumn 
         Caption         =   "Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   765
         Width           =   900
      End
      Begin VB.Label lblDBValueRecordID 
         Caption         =   "Record :"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   1155
         Width           =   930
      End
      Begin VB.Label lblDBValueWebForm 
         Caption         =   "Element :"
         Height          =   195
         Left            =   200
         TabIndex        =   27
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblDBValueRecordSelector 
         Caption         =   "Record Selector :"
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   1965
         Width           =   1515
      End
   End
   Begin VB.Frame fraDefinition 
      Caption         =   "Workflow Value :"
      Height          =   1200
      Index           =   1
      Left            =   4700
      TabIndex        =   15
      Top             =   600
      Width           =   4035
      Begin VB.ComboBox cboWebForm 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   2085
      End
      Begin VB.ComboBox cboWebFormValue 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   700
         Width           =   2085
      End
      Begin VB.Label lblWFValue 
         Caption         =   "Value :"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   765
         Width           =   810
      End
      Begin VB.Label lblWFWebForm 
         Caption         =   "Web Form :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame fraDefinition 
      Caption         =   "Fixed Value :"
      Height          =   2800
      Index           =   0
      Left            =   2100
      TabIndex        =   6
      Top             =   565
      Width           =   2400
      Begin VB.Frame fraLogicValues 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   200
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   2000
         Begin VB.OptionButton optLogicValue 
            Caption         =   "&False"
            Height          =   315
            Index           =   1
            Left            =   1000
            TabIndex        =   12
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton optLogicValue 
            Caption         =   "&True"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.TextBox txtTextValue 
         Height          =   315
         Left            =   200
         TabIndex        =   13
         Top             =   1900
         Width           =   2000
      End
      Begin VB.ComboBox cboOptions 
         Height          =   315
         Left            =   200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2300
         Visible         =   0   'False
         Width           =   2000
      End
      Begin TDBNumberCtrl.TDBNumber tdbNumberValue 
         Height          =   315
         Left            =   200
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   2000
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   999999999
         MinValue        =   -999999999
         Value           =   0
         SelStart        =   0
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   ""
         Format          =   "###############"
         DisplayFormat   =   "###############"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "&Caption"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmWorkflowElementColumn.frx":0166
         MousePointer    =   0
      End
      Begin COASpinner.COA_Spinner asrSpinnerValue 
         Height          =   315
         Left            =   195
         TabIndex        =   8
         Top             =   705
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   999999999
         MinimumValue    =   -999999999
         Text            =   "999"
      End
      Begin GTMaskDate.GTMaskDate ASRDateValue 
         Height          =   315
         Left            =   200
         TabIndex        =   9
         Top             =   1100
         Visible         =   0   'False
         Width           =   1500
         _Version        =   65537
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         NullText        =   "__/__/____"
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelect      =   -1  'True
         MaskCentury     =   2
         SpinButtonEnabled=   0   'False
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTips        =   0   'False
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cboColumns 
      Height          =   315
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   3090
   End
   Begin VB.Frame fraType 
      Caption         =   "Value :"
      Height          =   1965
      Left            =   100
      TabIndex        =   2
      Top             =   600
      Width           =   1900
      Begin VB.OptionButton optValueType 
         Caption         =   "C&alculation"
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   38
         Top             =   1500
         Width           =   1500
      End
      Begin VB.OptionButton optValueType 
         Caption         =   "&Database Value"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1100
         Width           =   1680
      End
      Begin VB.OptionButton optValueType 
         Caption         =   "&Fixed Value"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optValueType 
         Caption         =   "&Workflow Value"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   700
         Width           =   1680
      End
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   2040
      TabIndex        =   31
      Top             =   3600
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   33
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Label lblColumn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column :"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "frmWorkflowElementColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form handling variables.
Private mfChanged As Boolean
Private mfForcedChanged As Boolean
Private mfLoading As Boolean
Private mfCancelled As Boolean

Private mfInitializing As Boolean
Private msInitializeMessage As String

Private mlngTableID As Long
Private mlngColumnID As Long
Private miValueType As WorkflowStoredDataValueTypes

Private msValue As String
Private msWFFormIdentifier As String
Private msWFValueIdentifier As String
Private mlngCalculationExprID As Long

Private malngColumnInfo() As Long
Private miColumnDataType As SQLDataType
Private miColumnOLEType As OLEType
Private mfColumnMaxOLESizeEnabled As Boolean
Private miControlType As ControlTypes
Private mblnOptionGroup As Boolean
Private mfrmCallingForm As Form

Private malngColumnsDone() As Long
Private maWFPrecedingElements() As VB.Control
Private maWFAllElements() As VB.Control

Private mlngDBColumnID As Long
Private miDBRecord As WorkflowRecordSelectorTypes

Private mlngPersonnelTableID As Long
Private mlngBaseTableID As Long
Private miInitiationType As WorkflowInitiationTypes

Private Sub RefreshValueControls()
  ' Refresh the controls.
  Dim fraTemp As Frame
  
  If Not mfInitializing Then
    If ((miColumnDataType = sqlOle) _
        Or (miColumnDataType = sqlVarBinary)) Then
      
      If ((miValueType = giWFDATAVALUE_FIXED) _
        Or (miValueType = giWFDATAVALUE_CALC)) Then
        
        ValueType = giWFDATAVALUE_WFVALUE
      End If
      
      If ((miValueType = giWFDATAVALUE_WFVALUE) _
        And (miColumnOLEType <> OLE_EMBEDDED) _
        And (Not mfColumnMaxOLESizeEnabled)) Then
        
        ValueType = giWFDATAVALUE_DBVALUE
      End If
    End If
  End If
    
  optValueType(giWFDATAVALUE_FIXED).Enabled = ((miColumnDataType <> sqlOle) _
    And (miColumnDataType <> sqlVarBinary))
  optValueType(giWFDATAVALUE_WFVALUE).Enabled = ((miColumnDataType <> sqlOle) _
    And (miColumnDataType <> sqlVarBinary)) _
    Or ((miColumnOLEType = OLE_EMBEDDED) And (mfColumnMaxOLESizeEnabled))
  optValueType(giWFDATAVALUE_CALC).Enabled = ((miColumnDataType <> sqlOle) _
    And (miColumnDataType <> sqlVarBinary))
  
  Select Case miValueType
    Case giWFDATAVALUE_FIXED
      RefreshValueControls_Fixed
    
    Case giWFDATAVALUE_WFVALUE
      cboWebForm_Refresh
  
    Case giWFDATAVALUE_DBVALUE
      cboDBValueTable_Refresh
  
    Case giWFDATAVALUE_CALC
      InitializeCalculationControls
  End Select

  ' Display only the frame that defines the selected component type.
  For Each fraTemp In fraDefinition
    fraTemp.Visible = (fraTemp.Index = miValueType)
  Next fraTemp
  Set fraTemp = Nothing

  RefreshScreen
  
End Sub

Public Property Get value() As String
  value = msValue

End Property


Public Property Get ItemWFFormIdentifier() As String
  ItemWFFormIdentifier = msWFFormIdentifier
  
End Property

Public Property Let ItemWFFormIdentifier(ByVal psNewValue As String)
  msWFFormIdentifier = psNewValue
  
End Property


Public Property Get ItemWFValueIdentifier() As String
  ItemWFValueIdentifier = msWFValueIdentifier
  
End Property



Public Property Let ItemWFValueIdentifier(ByVal psNewValue As String)
  msWFValueIdentifier = psNewValue
  
End Property


Public Property Get Changed() As Boolean
  Changed = mfChanged
  
End Property






Public Property Let Changed(ByVal pfNewValue As Boolean)
  If Not mfLoading Then
    mfChanged = pfNewValue
    RefreshScreen
  End If
  
End Property



Private Sub ASRDateValue_Change()
  Changed = True
  
End Sub

Private Sub ASRDateValue_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    ASRDateValue.DateValue = Date
  End If

End Sub


Private Sub ASRDateValue_LostFocus()
  ValidateGTMaskDate ASRDateValue

End Sub


Private Sub asrSpinnerValue_Change()
  Changed = True

End Sub

Private Sub cboColumns_Click()
  
  If cboColumns.ListIndex <> -1 Then
    mlngColumnID = cboColumns.ItemData(cboColumns.ListIndex)
    
    ' AE20080311 Fault #12991
    If miColumnDataType > 0 Then
      If miColumnDataType <> malngColumnInfo(1, cboColumns.ListIndex) Then
        mlngCalculationExprID = 0
      End If
    End If
    
    miColumnDataType = malngColumnInfo(1, cboColumns.ListIndex)
    miControlType = malngColumnInfo(2, cboColumns.ListIndex)
    miColumnOLEType = malngColumnInfo(10, cboColumns.ListIndex)
    mfColumnMaxOLESizeEnabled = (malngColumnInfo(11, cboColumns.ListIndex) = 1)

    RefreshValueControls
  End If

  Changed = True
  
End Sub


Private Sub cboWebForm_Refresh()
  ' Populate the WebForm combo and
  ' select the current webform if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim sMsg As String

  iIndex = -1

  ' Clear the current contents of the combo.
  cboWebForm.Clear

  ' Add  an item to the combo for each preceding web form.
  ' Ignore the first item, as it will be the current web form.
  For iLoop = 2 To UBound(maWFPrecedingElements)
    If maWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
      cboWebForm.AddItem maWFPrecedingElements(iLoop).Identifier
      cboWebForm.ItemData(cboWebForm.NewIndex) = iLoop
    End If
  Next iLoop

  For iLoop = 0 To cboWebForm.ListCount - 1
    If cboWebForm.List(iLoop) = msWFFormIdentifier Then
      iIndex = iLoop
    End If
  Next iLoop

  If (iIndex < 0) Then
    If (Len(Trim(msWFFormIdentifier)) > 0) Then
      sMsg = "The previously selected Workflow Value Web Form is no longer valid."
  
      If cboWebForm.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value Web Form has been selected."
      End If
  
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
      
      mfForcedChanged = True
    End If
    
    iIndex = 0
  End If
  
  ' Enable the combo if there are items.
  With cboWebForm
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

      cboWebFormValue_Refresh
    End If
  End With
    
End Sub

Private Sub cboWebFormValue_Refresh()
  ' Populate the WF Value combo and
  ' select the current value if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim fMatchingDataType As Boolean
  Dim sMsg As String

  iIndex = -1

  ' Clear the current contents of the combo.
  cboWebFormValue.Clear

  If cboWebForm.Enabled Then
    ' Add  an item to the combo for each input item in the preceding web form.
    Set wfTemp = maWFPrecedingElements(cboWebForm.ItemData(cboWebForm.ListIndex))

    asItems = wfTemp.Items

    For iLoop = 1 To UBound(asItems, 2)
      fMatchingDataType = False
      
      If asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_CHAR Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_DATE Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_NUMERIC Or _
          asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_LOGIC Then

        Select Case CInt(asItems(6, iLoop))
          Case giEXPRVALUE_CHARACTER
            fMatchingDataType = (miColumnDataType = sqlVarChar) Or _
              (miColumnDataType = sqlLongVarChar)
          Case giEXPRVALUE_NUMERIC
            fMatchingDataType = (miColumnDataType = sqlNumeric) Or _
              (miColumnDataType = sqlInteger)
          Case giEXPRVALUE_LOGIC
            fMatchingDataType = (miColumnDataType = sqlBoolean)
          Case giEXPRVALUE_DATE
            fMatchingDataType = (miColumnDataType = sqlDate)
        End Select

      ElseIf asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_DROPDOWN Or _
        asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
          
        fMatchingDataType = (miColumnDataType = sqlVarChar) Or _
          (miColumnDataType = sqlLongVarChar)
      
      ElseIf asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_LOOKUP Then
      
        fMatchingDataType = (miColumnDataType = GetColumnDataType(CLng(asItems(49, iLoop))))
        
      ElseIf asItems(2, iLoop) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
      
        fMatchingDataType = ((miColumnDataType = sqlOle) Or _
          (miColumnDataType = sqlVarBinary)) _
          And (miColumnOLEType = OLE_EMBEDDED) _
          And (mfColumnMaxOLESizeEnabled)
      End If
    
      If fMatchingDataType Then
        cboWebFormValue.AddItem asItems(9, iLoop)
      End If
    Next iLoop
  End If

  For iLoop = 0 To cboWebFormValue.ListCount - 1
    If cboWebFormValue.List(iLoop) = msWFValueIdentifier Then
      iIndex = iLoop
    End If
  Next iLoop

  If (iIndex < 0) Then
    If (Len(Trim(msWFValueIdentifier)) > 0) Then
      sMsg = "The previously selected Workflow Value identifier is no longer valid."
      
      If cboWebFormValue.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value identifier has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
      
      mfForcedChanged = True
    End If
    
    iIndex = 0
  End If
  
  ' Enable the combo if there are items.
  With cboWebFormValue
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no values>"
      .ItemData(.NewIndex) = -1
      .ListIndex = 0
    End If
  End With
    
End Sub

Private Sub RefreshValueControls_Fixed()
  ' Display the Value controls that match the selected column's data type.
  Dim lngCount As Long
  Dim lngSize As Long
  Dim lngDecimals As Long
  Dim lngSpinnerMaximum As Long
  Dim lngSpinnerMinimum As Long
  Dim lngSpinnerIncrement As Long
  Dim blnMandatory As Boolean
  Dim bThousandSeparators As Boolean
  Dim sFormat As String
  Dim blnSpinner As Boolean
  Dim blnTextBox As Boolean
  Dim bMultiline As Boolean

'''''  optLogicValue(0).Value = True
'''''  tdbNumberValue.Value = 0
'''''  asrSpinnerValue.Text = "0"
'''''  ASRDateValue.Text = ""
'''''  txtTextValue.Text = ""
  
  cboOptions_Refresh

  ' Read the column's details from the array.
  If cboColumns.ListIndex <> -1 Then
    If miColumnDataType = sqlInteger Then
      'Integers always size 10
      '(don't know why but size is stored as 1) !!!!
      lngSize = 10
    Else
      lngSize = malngColumnInfo(3, cboColumns.ListIndex)
    End If

    lngDecimals = malngColumnInfo(4, cboColumns.ListIndex)
    lngSpinnerMaximum = malngColumnInfo(5, cboColumns.ListIndex)
    lngSpinnerMinimum = malngColumnInfo(6, cboColumns.ListIndex)
    lngSpinnerIncrement = malngColumnInfo(7, cboColumns.ListIndex)
    blnMandatory = malngColumnInfo(8, cboColumns.ListIndex)
    bThousandSeparators = malngColumnInfo(9, cboColumns.ListIndex)
    bMultiline = malngColumnInfo(12, cboColumns.ListIndex)
  End If

  cboOptions.Visible = mblnOptionGroup
  If mblnOptionGroup Then
    fraLogicValues.Visible = False
    tdbNumberValue.Visible = False
    asrSpinnerValue.Visible = False
    ASRDateValue.Visible = False
    txtTextValue.Visible = False

    If (Not blnMandatory) And (miControlType <> giCTRL_OPTIONGROUP) Then
      cboOptions.AddItem ""
    End If
    Exit Sub
  End If

  fraLogicValues.Visible = (miColumnDataType = sqlBoolean)

  tdbNumberValue.Visible = ((miColumnDataType = sqlNumeric) Or _
    ((miColumnDataType = sqlInteger) And (miControlType <> giCTRL_SPINNER)))
  If (miColumnDataType = sqlNumeric Or miColumnDataType = sqlInteger) Then
    ' Loop and create the format mask
    sFormat = "0"
    For lngCount = 2 To (lngSize - lngDecimals)
      If bThousandSeparators = True Then
        sFormat = IIf(lngCount Mod 3 = 0 And (lngCount <> (lngSize - lngDecimals)), ",#", "#") & sFormat
      Else
        sFormat = "#" & sFormat
      End If
    Next lngCount

    If lngDecimals > 0 Then
      sFormat = sFormat & "."
      For lngCount = 1 To lngDecimals
        sFormat = sFormat & "0"
      Next lngCount
    End If

    With tdbNumberValue
      .Format = sFormat
      .DisplayFormat = sFormat
    End With
  End If

  ASRDateValue.Visible = (miColumnDataType = sqlDate)
  If miColumnDataType = sqlDate Then
    'MH20001003 Fault 1048
    'After making the greentree control visible you will
    'be unable to use the arrow keys to scroll though the
    'combo box items.  Setting focus to the form seems to
    'fixed this !
    If Me.Visible Then Me.SetFocus
  End If

  blnSpinner = (miColumnDataType = sqlInteger And miControlType = giCTRL_SPINNER)

  asrSpinnerValue.Visible = blnSpinner
  If (miColumnDataType = sqlInteger) Then
    With asrSpinnerValue
      .MinimumValue = lngSpinnerMinimum
      .MaximumValue = lngSpinnerMaximum
      .Increment = lngSpinnerIncrement
    End With
  End If

  blnTextBox = (miColumnDataType = sqlVarChar) Or _
               (miColumnDataType = sqlLongVarChar)

  txtTextValue.Visible = blnTextBox
  If blnTextBox Then
    If bMultiline Then
      txtTextValue.MaxLength = 0
    Else
      txtTextValue.MaxLength = lngSize
    End If
  End If
  
End Sub


Private Sub cboOptions_Refresh()
  Dim fOptionGroup As Boolean
  Dim iIndex As Integer
  
  iIndex = 0
  
  cboOptions.Clear
      
  With recContValEdit
    .Index = "idxColumnID"
    .Seek ">=", cboColumns.ItemData(cboColumns.ListIndex)
        
    If Not .NoMatch Then
      Do While Not .EOF
        ' If no more control values for this column exit loop
        If !ColumnID <> cboColumns.ItemData(cboColumns.ListIndex) Then
          Exit Do
        End If
  
        cboOptions.AddItem Trim(!value)
  
        If Trim(!value) = msValue Then
          iIndex = cboOptions.NewIndex
        End If
        
        .MoveNext
      Loop
    End If
  End With

  mblnOptionGroup = (cboOptions.ListCount > 0)

  If mblnOptionGroup Then
    cboOptions.ListIndex = iIndex
  End If

End Sub


Private Sub cboDBValueColumn_Click()
  If cboDBValueColumn.ListCount > 0 Then
    mlngDBColumnID = cboDBValueColumn.ItemData(cboDBValueColumn.ListIndex)
    Changed = True
  Else
    mlngDBColumnID = -1
  End If

End Sub


Private Sub cboDBValueRecord_Click()
  If cboDBValueRecord.ListCount > 0 Then
    miDBRecord = cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex)
    Changed = True
  Else
    miDBRecord = giWFRECSEL_UNKNOWN
  End If

  cboDBValueWebForm_Refresh

End Sub


Private Sub cboDBValueRecordSelector_Click()
  If cboDBValueRecordSelector.ListCount > 0 Then
    msWFValueIdentifier = cboDBValueRecordSelector.List(cboDBValueRecordSelector.ListIndex)
    Changed = True
  Else
    msWFValueIdentifier = ""
  End If

End Sub


Private Sub cboDBValueTable_Click()
  ' Populate the field combo with the relevant fields.
  cboDBValueColumn_Refresh
  cboDBValueRecord_Refresh

End Sub

Private Sub cboDBValueWebForm_Click()
  If cboDBValueWebForm.ListCount > 0 Then
    msWFFormIdentifier = cboDBValueWebForm.List(cboDBValueWebForm.ListIndex)
    Changed = True
  Else
    msWFFormIdentifier = ""
  End If
  
  cboDBValueRecordSelector_Refresh

End Sub


Private Sub cboOptions_Click()
  Changed = True
  
End Sub


Private Sub cboWebform_Click()
  msWFFormIdentifier = cboWebForm.List(cboWebForm.ListIndex)
  
  Changed = True
  
  cboWebFormValue_Refresh
  
End Sub


Private Sub cboWebFormValue_Click()
  msWFValueIdentifier = cboWebFormValue.List(cboWebFormValue.ListIndex)
  
  Changed = True

End Sub


Private Sub cmdCancel_Click()
  ' Set the cancelled flag.
  mfCancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
  ' Set the cancelled flag.
  Dim objMisc As Misc
  
  mfCancelled = False

  If optValueType(giWFDATAVALUE_FIXED) Then
    ' The column is to be updated with a fixed value.
    If mblnOptionGroup Then
      msValue = cboOptions.Text
    Else
      Select Case miColumnDataType
        Case sqlBoolean
          msValue = IIf(optLogicValue(0).value, "True", "False")
        Case sqlNumeric, sqlInteger
          If miControlType = giCTRL_SPINNER Then
            msValue = Trim(asrSpinnerValue.Text)
          Else
            msValue = tdbNumberValue.Text
          End If
        Case sqlDate
          If IsDate(ASRDateValue.Text) Then
            Set objMisc = New Misc
            msValue = objMisc.ConvertLocaleDateToSQL(ASRDateValue.Text)
            Set objMisc = Nothing
          Else
            msValue = "Null"
          End If
        Case sqlVarChar, sqlLongVarChar
          msValue = txtTextValue.Text
      End Select
    End If
  End If

  ' Unload the form.
  UnLoad Me

End Sub


Public Sub Initialize(pfrmCallingForm As Form, _
  palngColumnsDone As Variant, _
  plngTableID As Long, _
  plngColumnID As Long, _
  piValueType As WorkflowStoredDataValueTypes, _
  psValue As String, _
  psWFForm As String, _
  psWFValue As String, _
  plngDBColumnID As Long, _
  piDBRecord As WorkflowRecordSelectorTypes, _
  pfNew As Boolean, _
  plngCalculationExprID)

  Dim sDefaultCharacter As String
  Dim dblDefaultNumeric As Double
  Dim fDefaultLogic As Boolean
  Dim dtDefaultDate As Date
  Dim objMisc As Misc
  Dim sDefaultDate As String
  
  mfForcedChanged = False

  mfInitializing = True
  msInitializeMessage = ""

  Set mfrmCallingForm = pfrmCallingForm
  malngColumnsDone = palngColumnsDone
  mlngTableID = plngTableID
  mlngColumnID = plngColumnID
  msValue = psValue
  msWFFormIdentifier = psWFForm
  msWFValueIdentifier = psWFValue
  mlngDBColumnID = plngDBColumnID
  miDBRecord = piDBRecord
  mlngBaseTableID = pfrmCallingForm.BaseTable
  miInitiationType = pfrmCallingForm.InitiationType
  mlngCalculationExprID = plngCalculationExprID

  ReDim maWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements maWFPrecedingElements

  ReDim maWFAllElements(0)
  mfrmCallingForm.CallingForm.AllElements maWFAllElements
  ' Populate the column combo (and assign the current or defaulted value)
  cboColumns_Refresh
  ValueType = piValueType

  If piValueType = giWFDATAVALUE_FIXED Then
    ' Initialise the Fixed Value controls.
    ' NB. DBValue and WFValue controls are initialised as the value type is set.
    sDefaultCharacter = vbNullString
    dblDefaultNumeric = 0
    fDefaultLogic = True
    dtDefaultDate = Date

    Select Case miColumnDataType
      Case sqlBoolean
        fDefaultLogic = (msValue = "True")

      Case sqlNumeric, sqlInteger
        dblDefaultNumeric = Val(msValue)

      Case sqlDate
        Set objMisc = New Misc
        sDefaultDate = IIf(Len(msValue) > 0 And (UCase(msValue) <> "NULL"), objMisc.ConvertSQLDateToLocale(msValue), "")
        Set objMisc = Nothing

      Case sqlVarChar, sqlLongVarChar
        sDefaultCharacter = msValue
    End Select

    txtTextValue.Text = sDefaultCharacter

    If miControlType = giCTRL_SPINNER Then
      asrSpinnerValue.value = dblDefaultNumeric
    Else
      tdbNumberValue.value = dblDefaultNumeric
    End If

    optLogicValue(0).value = fDefaultLogic
    optLogicValue(1).value = Not optLogicValue(0).value

    If Len(sDefaultDate) > 0 Then
      ASRDateValue.Text = sDefaultDate
    Else
      ASRDateValue.Text = ""
    End If
  End If
  
  If Len(msInitializeMessage) > 0 Then
    MsgBox msInitializeMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  mfInitializing = False
  
  mfChanged = pfNew Or mfForcedChanged
  RefreshScreen
    
End Sub



Private Sub cboColumns_Refresh()
  ' Populate the Column combo and
  ' select the current column if it is still valid.
  Dim iIndex As Integer
  Dim fDone As Boolean
  Dim iLoop As Integer
  
  iIndex = -1

  ' Clear the current contents of the combo.
  cboColumns.Clear
  
  ' Create an array of column info
  ' Col 1 = data type
  ' Col 2 = control type
  ' Col 3 = size
  ' Col 4 = decimals
  ' Col 5 = spinner max.
  ' Col 6 = spinner min.
  ' Col 7 = spinner inc.
  ' Col 8 = mandatory
  ' Col 9 = use 1000 sep.
  ' Col 10 = OLE type
  ' Col 11 = Max OLE Size Enabled
  ' Col 12 = Multiline
  ReDim malngColumnInfo(12, 0)
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mlngTableID

    If Not .NoMatch Then
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If

      ' Add  an item to the combo for each column that has not been deleted.
      Do While Not .EOF
        ' Do not allow the user to select system columns, or deleted columns
        If (!TableID = mlngTableID) _
          And (!Deleted = False) _
          And (!ColumnType <> giCOLUMNTYPE_LINK) _
          And (!ColumnType <> giCOLUMNTYPE_SYSTEM) Then

          fDone = False
          For iLoop = 1 To UBound(malngColumnsDone)
            If malngColumnsDone(iLoop) = .Fields("columnID") Then
              fDone = True
              Exit For
            End If
          Next iLoop
          
          If Not fDone Then
            cboColumns.AddItem .Fields("columnName")
            cboColumns.ItemData(cboColumns.NewIndex) = .Fields("columnID")
  
            ReDim Preserve malngColumnInfo(12, cboColumns.NewIndex)
            malngColumnInfo(1, cboColumns.NewIndex) = .Fields("dataType")
            malngColumnInfo(2, cboColumns.NewIndex) = .Fields("controlType")
            malngColumnInfo(3, cboColumns.NewIndex) = IIf(IsNull(.Fields("Size")), 0, .Fields("Size"))
            malngColumnInfo(4, cboColumns.NewIndex) = IIf(IsNull(.Fields("Decimals")), 0, .Fields("Decimals"))
            malngColumnInfo(5, cboColumns.NewIndex) = IIf(IsNull(.Fields("SpinnerMaximum")), 0, .Fields("SpinnerMaximum"))
            malngColumnInfo(6, cboColumns.NewIndex) = IIf(IsNull(.Fields("SpinnerMinimum")), 0, .Fields("SpinnerMinimum"))
            malngColumnInfo(7, cboColumns.NewIndex) = IIf(IsNull(.Fields("SpinnerIncrement")), 0, .Fields("SpinnerIncrement"))
            malngColumnInfo(8, cboColumns.NewIndex) = IIf(IsNull(.Fields("Mandatory")), 0, .Fields("Mandatory"))
            malngColumnInfo(9, cboColumns.NewIndex) = IIf(IsNull(.Fields("Use1000Separator")), 0, .Fields("Use1000Separator"))
            malngColumnInfo(10, cboColumns.NewIndex) = IIf(IsNull(.Fields("OLEType")), 0, .Fields("OLEType"))
            malngColumnInfo(11, cboColumns.NewIndex) = IIf(IsNull(.Fields("MaxOLESizeEnabled")), 0, IIf(.Fields("MaxOLESizeEnabled"), 1, 0))
            malngColumnInfo(12, cboColumns.NewIndex) = IIf(IsNull(.Fields("Multiline")), 0, IIf(.Fields("Multiline"), 1, 0))
            
            If .Fields("columnID") = mlngColumnID Then
              iIndex = cboColumns.NewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Enable the combo if there are items.
  With cboColumns
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    
      miColumnDataType = malngColumnInfo(1, cboColumns.ListIndex)
      miControlType = malngColumnInfo(2, cboColumns.ListIndex)
      miColumnOLEType = malngColumnInfo(10, cboColumns.ListIndex)
      mfColumnMaxOLESizeEnabled = (malngColumnInfo(11, cboColumns.ListIndex) = 1)
    Else
      .Enabled = False

      .AddItem "<no columns>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
    End If
  End With
    
End Sub


Private Sub RefreshScreen()
  Dim fEnableOK As Boolean
  Dim fEnableDBValueRecord As Boolean
  Dim fEnableDBValueWebForm As Boolean
  Dim fEnableDBValueRecordSelector As Boolean
  Dim wfTemp As VB.Control
  
  fEnableOK = mfChanged
  
  If fEnableOK Then
    Select Case miValueType
      Case giWFDATAVALUE_FIXED
      
      Case giWFDATAVALUE_WFVALUE
        fEnableOK = (cboWebFormValue.ListIndex >= 0)
        If fEnableOK Then
          fEnableOK = (cboWebFormValue.ItemData(cboWebFormValue.ListIndex) <> -1)
        End If
    
      Case giWFDATAVALUE_DBVALUE
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

           fEnableDBValueRecordSelector = (cboDBValueRecordSelector.ItemData(cboDBValueRecordSelector.ListIndex) <> giWFRECSEL_INITIATOR) _
            And (cboDBValueRecordSelector.ItemData(cboDBValueRecordSelector.ListIndex) <> giWFRECSEL_TRIGGEREDRECORD)
           
            If fEnableDBValueRecordSelector Then
              Set wfTemp = SelectedElement
              If Not wfTemp Is Nothing Then
                fEnableDBValueRecordSelector = (wfTemp.ElementType = elem_WebForm)
              End If
            End If
          End If
        End If

        cboDBValueRecord.Enabled = fEnableDBValueRecord

        cboDBValueWebForm.Enabled = fEnableDBValueWebForm
        cboDBValueWebForm.BackColor = IIf(fEnableDBValueWebForm, vbWindowBackground, vbButtonFace)
        lblDBValueWebForm.Enabled = fEnableDBValueWebForm

        cboDBValueRecordSelector.Enabled = fEnableDBValueRecordSelector
        cboDBValueRecordSelector.BackColor = IIf(fEnableDBValueRecordSelector, vbWindowBackground, vbButtonFace)
        lblDBValueRecordSelector.Enabled = fEnableDBValueRecordSelector

        fEnableOK = cboDBValueColumn.Enabled And fEnableDBValueRecord
        If fEnableOK Then
          If cboDBValueRecord.ListIndex >= 0 Then
            If (cboDBValueRecord.ItemData(cboDBValueRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
              fEnableOK = fEnableDBValueWebForm

              If fEnableOK Then
                Set wfTemp = SelectedElement
                If Not wfTemp Is Nothing Then
                  If (wfTemp.ElementType = elem_WebForm) Then
                    fEnableOK = fEnableDBValueRecordSelector
                  End If
                End If
              End If
            End If
          End If
        End If
    
    Case giWFDATAVALUE_CALC
      fEnableOK = fEnableOK And (mlngCalculationExprID > 0)
      
    End Select
  End If
  
  fEnableOK = fEnableOK And (cboColumns.Enabled)
  
  cmdOK.Enabled = fEnableOK
  
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


Private Sub Form_Load()
  fraOKCancel.BorderStyle = vbBSNone
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  UI.FormatGTDateControl ASRDateValue
  FormatFrames

End Sub


Private Sub FormatFrames()
  ' Position controls that aren't correctly positioned at design time.
  ' ie. the Straight value controls which share the same position.
  Dim fraTemp As Frame
  
  Const iXGAP = 200
  Const iYGAP = 200
  Const iXFRAMEGAP = 150
  Const iYFRAMEGAP = 100
  Const iITEMFRAMEWIDTH = 1900
  Const iFRAMEWIDTH = 5700
  Const iFRAMEHEIGHT = 2400

  mfLoading = True
  
  ' Position and size the item type frame.
  With fraType
    .Left = iXFRAMEGAP
    .Top = cboColumns.Top + cboColumns.Height + iYFRAMEGAP
    .Width = iITEMFRAMEWIDTH
    .Height = iFRAMEHEIGHT
  End With
  
  ' Position and size the item definition frames.
  For Each fraTemp In fraDefinition
    With fraTemp
      .Left = fraType.Left + iITEMFRAMEWIDTH + iXFRAMEGAP
      .Top = fraType.Top
      .Width = iFRAMEWIDTH
      .Height = iFRAMEHEIGHT
    End With
  Next fraTemp
  Set fraTemp = Nothing

  ' Format the controls within the frames.
  FormatFrame_FixedValue
  FormatFrame_WFValue
  FormatFrame_DBValue
  FormatFrame_CalcValue

  ' Position and size the OK/Cancel command controls.
  With fraOKCancel
    .Top = fraType.Top + fraType.Height + iYGAP
    .Left = fraDefinition(fraDefinition.LBound).Left + _
      fraDefinition(fraDefinition.LBound).Width - .Width
  End With

  ' Size the form.
  Me.Width = fraDefinition(fraDefinition.UBound).Left + _
    fraDefinition(fraDefinition.UBound).Width + iXFRAMEGAP + _
   (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
  Me.Height = fraOKCancel.Top + fraOKCancel.Height + iXFRAMEGAP + _
    (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
  
  mfLoading = False

End Sub


Private Sub FormatFrame_WFValue()
  ' Size and position the Workflow value item controls.
  Const iXGAP = 200

  cboWebForm.Width = fraDefinition(giWFDATAVALUE_WFVALUE).Width - cboWebForm.Left - iXGAP
  cboWebFormValue.Width = fraDefinition(giWFDATAVALUE_WFVALUE).Width - cboWebFormValue.Left - iXGAP
  
End Sub


Private Sub FormatFrame_FixedValue()
  ' Size and position the Workflow value item controls.
  Const iXGAP = 200
  Const iYGAP = 300
  
  txtTextValue.Left = iXGAP
  txtTextValue.Top = iYGAP
  txtTextValue.Width = fraDefinition(giWFDATAVALUE_FIXED).Width - txtTextValue.Left - iXGAP

  fraLogicValues.Left = iXGAP
  fraLogicValues.Top = iYGAP
  fraLogicValues.BackColor = Me.BackColor

  tdbNumberValue.Left = iXGAP
  tdbNumberValue.Top = iYGAP
  UI.FormatTDBNumberControl Me.tdbNumberValue
  tdbNumberValue.Width = fraDefinition(giWFDATAVALUE_FIXED).Width - tdbNumberValue.Left - iXGAP

  asrSpinnerValue.Left = iXGAP
  asrSpinnerValue.Top = iYGAP
  asrSpinnerValue.Width = fraDefinition(giWFDATAVALUE_FIXED).Width - asrSpinnerValue.Left - iXGAP

  ASRDateValue.Left = iXGAP
  ASRDateValue.Top = iYGAP

  cboOptions.Left = iXGAP
  cboOptions.Top = iYGAP
  cboOptions.Width = fraDefinition(giWFDATAVALUE_FIXED).Width - cboOptions.Left - iXGAP

End Sub

Private Sub FormatFrame_CalcValue()
    ' Size and position the Calculation item controls.
  Const iXGAP = 200
  
  txtCalcCalculation.Width = fraDefinition(giWFDATAVALUE_CALC).Width - txtCalcCalculation.Left - iXGAP - cmdCalcCalculation.Width
  cmdCalcCalculation.Left = txtCalcCalculation.Left + txtCalcCalculation.Width

End Sub


Private Sub FormatFrame_DBValue()
  ' Size and position the Database Value item controls.
  Const iXGAP = 200

  cboDBValueTable.Width = fraDefinition(giWFDATAVALUE_DBVALUE).Width - cboDBValueTable.Left - iXGAP
  cboDBValueColumn.Width = fraDefinition(giWFDATAVALUE_DBVALUE).Width - cboDBValueColumn.Left - iXGAP
  cboDBValueRecord.Width = fraDefinition(giWFDATAVALUE_DBVALUE).Width - cboDBValueRecord.Left - iXGAP
  cboDBValueWebForm.Width = fraDefinition(giWFDATAVALUE_DBVALUE).Width - cboDBValueWebForm.Left - iXGAP
  cboDBValueRecordSelector.Width = fraDefinition(giWFDATAVALUE_DBVALUE).Width - cboDBValueRecordSelector.Left - iXGAP

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged And cmdOK.Enabled Then
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


Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
    
End Property



Public Property Get ValueType() As WorkflowStoredDataValueTypes
  ValueType = miValueType

End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
    
End Property


Public Property Let ColumnID(ByVal plngNewValue As Long)
  mlngColumnID = plngNewValue
  
End Property





Public Property Let ValueType(ByVal piNewValue As WorkflowStoredDataValueTypes)
  miValueType = piNewValue
  
  If miValueType = giWFDATAVALUE_FIXED Then
    optValueType_Click (miValueType)
  Else
    optValueType(miValueType).value = True
  End If

End Property

Private Sub optLogicValue_Click(Index As Integer)
  Changed = True
  
End Sub

Private Sub optValueType_Click(Index As Integer)
  ' Set the component type property.
  miValueType = Index
  RefreshValueControls
  Changed = True

End Sub


Private Sub cboDBValueTable_Refresh()
  ' Populate the DB Value Table combo and
  ' select the current table if it is still valid.
  Dim fTableOK As Boolean
  Dim iLoop As Integer
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
      End If

      .MoveNext
    Loop
  End With

  ' Enable the combo if there are items.
  With cboDBValueTable
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = lngTableID Then
        iIndex = iLoop
      End If
      
      If ((miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) And (.ItemData(iLoop) = mlngPersonnelTableID)) Or _
        ((miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) And (.ItemData(iLoop) = mlngBaseTableID)) Then
        
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
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
      cboDBValueRecord_Refresh
    End If
  End With
    
End Sub
Private Sub cboDBValueColumn_Refresh()
  ' Populate the DB Value Column combo and
  ' select the current column if it is still valid.
  Dim iIndex As Integer
  Dim lngTableID As Long
  Dim fMatchingDataType As Boolean
  Dim iLoop As Integer
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
        If (!TableID = lngTableID) And _
          (!Deleted = False) And _
          (!DataType = miColumnDataType) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) Then

          fColumnOK = ((miColumnDataType <> sqlOle) _
            And (miColumnDataType <> sqlVarBinary)) _
            Or (miColumnOLEType = !OLEType)
            
          If fColumnOK _
            And ((miColumnDataType = sqlOle) Or (miColumnDataType = sqlVarBinary)) _
            And (miColumnOLEType = OLE_EMBEDDED) Then
          
            fColumnOK = mfColumnMaxOLESizeEnabled Or (Not !MaxOLESizeEnabled)
          End If
          
          If fColumnOK Then
            cboDBValueColumn.AddItem .Fields("columnName")
            cboDBValueColumn.ItemData(cboDBValueColumn.NewIndex) = .Fields("columnID")
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Enable the combo if there are items.
  With cboDBValueColumn
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = mlngDBColumnID Then
        iIndex = iLoop
      End If
    Next iLoop
    
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


Public Property Get ItemDBColumnID() As Long
  If miValueType = giWFDATAVALUE_DBVALUE Then
    ItemDBColumnID = mlngDBColumnID
  Else
    ItemDBColumnID = 0
  End If
    
End Property


Public Property Let ItemDBColumnID(ByVal plngNewValue As Long)
  mlngDBColumnID = plngNewValue
  
End Property





Public Property Get ItemDBRecord() As WorkflowRecordSelectorTypes
  If miValueType = giWFDATAVALUE_DBVALUE Then
    ItemDBRecord = miDBRecord
  Else
    ItemDBRecord = 0
  End If
  
End Property


Public Property Let ItemDBRecord(ByVal piNewValue As WorkflowRecordSelectorTypes)
  miDBRecord = piNewValue
  
End Property


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
  
  ' Get an array of the valid table IDs (base table and it's descendants)
  ReDim alngValidTables(0)

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
            
            fFound = False
            
            ' Get an array of the valid table IDs (base table and it's ascendants)
            ReDim alngValidTables(0)
            TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
            
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
        fFound = False
        
        ReDim alngValidTables(0)
        TableAscendants wfTemp.DataTableID, alngValidTables
  
        'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
        'If wfTemp.DataAction = DATAACTION_DELETE Then
        '  ' Cannot do anything with a Deleted record, but can use its ascendants.
        '  ' Remove the table itself from the array of valid tables.
        '  alngValidTables(1) = 0
        'End If
        
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
      fFound = False
      If mlngPersonnelTableID > 0 Then
        ReDim alngValidTables(0)
        TableAscendants mlngPersonnelTableID, alngValidTables
        
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
      End If

      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no values>"
      .ItemData(.NewIndex) = giWFRECSEL_UNKNOWN
      .ListIndex = 0
    
      cboDBValueWebForm_Refresh
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
      
                ' Get an array of the valid table IDs (base table and it's descendants)
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

    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = msWFValueIdentifier Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(msWFValueIdentifier)) > 0) Then
        sMsg = "The previously selected Database Value Record Selector is no longer valid."
        
        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Database Value Record Selector has been selected."
        End If
        
        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If
        
        mfForcedChanged = True
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
          msWFValueIdentifier = ""
        End If
      End If
    End If
  End With
    
  RefreshScreen
  
End Sub

Private Sub InitializeCalculationControls()
  ' Initialize the Calculation item controls.
  txtCalcCalculation.Text = GetExpressionName(mlngCalculationExprID)
  
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
                ' Get an array of the valid table IDs (base table and it's descendants)
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
            ' Get an array of the valid table IDs (base table and it's descendants)
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
      If .List(lngLoop) = msWFFormIdentifier Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(msWFFormIdentifier)) > 0) Then
        sMsg = "The previously selected Database Value Element is no longer valid."

        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Database Value Element has been selected."
        End If

        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If

        mfForcedChanged = True
      End If

      iIndex = 0
    End If

    If .ListCount > 0 Then
      .Enabled = False
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
          msWFFormIdentifier = ""
          msWFValueIdentifier = ""
        End If
      End If
    
      cboDBValueRecordSelector_Refresh
    End If
  End With
    
End Sub

Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationExprID
  
End Property

Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationExprID = plngNewValue

End Property

Private Sub tdbNumberValue_Change()
  Changed = True

End Sub

Private Sub txtTextValue_Change()
  Changed = True

End Sub

Private Sub txtTextValue_GotFocus()
  With txtTextValue
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Private Sub txtTextValue_KeyPress(KeyAscii As Integer)
  If miColumnDataType = sqlInteger Then
    If KeyAscii > 31 Then
      If InStr("1234567890.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Beep
      End If
    End If
  End If

End Sub

Private Sub cmdCalcCalculation_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long

  lngOriginalID = mlngCalculationExprID

  ' Instantiate an expression object.
  Set objExpr = New CExpression
    
  With objExpr
    ' Set the properties of the expression object.
'    .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_DATE
    Select Case miColumnDataType
      Case dtVARCHAR
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_CHARACTER
      Case dtTIMESTAMP
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_DATE
      Case dtLONGVARBINARY
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_OLE
      Case dtVARBINARY
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_PHOTO
      Case dtINTEGER
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_NUMERIC
      Case dtBIT
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_LOGIC
      Case dtNUMERIC
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_NUMERIC
      Case dtLONGVARCHAR
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_CHARACTER
      Case Else
        .Initialise 0, mlngCalculationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_UNDEFINED
    End Select
    
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

