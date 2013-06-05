VERSION 5.00
Begin VB.Form frmWorkflowFindUsage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Usage"
   ClientHeight    =   3495
   ClientLeft      =   75
   ClientTop       =   4890
   ClientWidth     =   5190
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5065
   Icon            =   "frmWorkflowFindUsage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Caption         =   "Criteria :"
      Height          =   2800
      Left            =   100
      TabIndex        =   0
      Top             =   120
      Width           =   5000
      Begin VB.ComboBox cboItem 
         Height          =   315
         ItemData        =   "frmWorkflowFindUsage.frx":000C
         Left            =   2000
         List            =   "frmWorkflowFindUsage.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2300
         Width           =   2800
      End
      Begin VB.ComboBox cboWebform 
         Height          =   315
         ItemData        =   "frmWorkflowFindUsage.frx":0010
         Left            =   2000
         List            =   "frmWorkflowFindUsage.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1900
         Width           =   2800
      End
      Begin VB.ComboBox cboElement 
         Height          =   315
         ItemData        =   "frmWorkflowFindUsage.frx":0014
         Left            =   2000
         List            =   "frmWorkflowFindUsage.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1100
         Width           =   2800
      End
      Begin VB.OptionButton optCriteria 
         Caption         =   "&Web Form Item"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1500
         Width           =   2265
      End
      Begin VB.OptionButton optCriteria 
         Caption         =   "&Element (Web Form / Stored Data)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   700
         Width           =   3705
      End
      Begin VB.OptionButton optCriteria 
         Caption         =   "&Initiator / Triggered"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   2625
      End
      Begin VB.Label lblItem 
         Caption         =   "Item :"
         Height          =   255
         Left            =   825
         TabIndex        =   8
         Top             =   2355
         Width           =   780
      End
      Begin VB.Label lblElement 
         Caption         =   "Element :"
         Height          =   255
         Left            =   825
         TabIndex        =   3
         Top             =   1125
         Width           =   915
      End
      Begin VB.Label lblWebform 
         Caption         =   "Web Form :"
         Height          =   255
         Left            =   825
         TabIndex        =   6
         Top             =   1965
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3900
      TabIndex        =   11
      Top             =   3020
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2600
      TabIndex        =   10
      Top             =   3020
      Width           =   1200
   End
End
Attribute VB_Name = "frmWorkflowFindUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mwfElements() As VB.Control
'Private masElementIdentifiers() As String
'Private masWebformIdentifiers() As String
Private miChoice As Integer
Private miSelection As WorkflowFindUsageOption
Private msElement As String
Private msItem As String
Private miInitiationType As WorkflowInitiationTypes
Public Property Get Choice() As Integer
  Choice = miChoice
End Property

Private Sub FormatScreen()
  Dim sngCurrentY As Single
  Const Y_GAP = 400
  Const Y_BORDERGAP = 100
  Const Y_FORMGAP = 600
  
  sngCurrentY = 300
  
  With optCriteria(0)
    .Visible = (miInitiationType <> WORKFLOWINITIATIONTYPE_EXTERNAL)
    .Caption = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL, "&Initiator", "&Triggered")
    
    If miInitiationType <> WORKFLOWINITIATIONTYPE_EXTERNAL Then
      .Top = sngCurrentY
      sngCurrentY = sngCurrentY + Y_GAP
    End If
  End With
  
  With optCriteria(1)
    .Top = sngCurrentY
    sngCurrentY = sngCurrentY + Y_GAP
  End With
  
  With cboElement
    .Top = sngCurrentY
    lblElement.Top = .Top + ((.Height - lblElement.Height) / 2)
    sngCurrentY = sngCurrentY + Y_GAP
  End With
    
  With optCriteria(2)
    .Top = sngCurrentY
    sngCurrentY = sngCurrentY + Y_GAP
  End With
  
  With cboWebform
    .Top = sngCurrentY
    lblWebform.Top = .Top + ((.Height - lblWebform.Height) / 2)
    sngCurrentY = sngCurrentY + Y_GAP
  End With
    
  With cboItem
    .Top = sngCurrentY
    lblItem.Top = .Top + ((.Height - lblItem.Height) / 2)
    sngCurrentY = sngCurrentY + Y_GAP
  End With
  
  fraOptions.Height = sngCurrentY + Y_BORDERGAP
  
  With cmdOk
    .Top = fraOptions.Top + fraOptions.Height + Y_BORDERGAP
    cmdCancel.Top = .Top
  
    Me.Height = .Top + .Height + Y_FORMGAP
  End With
End Sub

Public Property Get Selection() As WorkflowFindUsageOption
  Selection = miSelection
End Property

Private Sub cboElement_Click()
  msElement = cboElement.Text
  
'  If cboElement.ListIndex <> -1 Then
'    msElement = masElementIdentifiers(cboElement.ItemData(cboElement.ListIndex))
'  End If
End Sub

Private Sub cboItem_Click()
  msItem = cboItem.Text
End Sub

Private Sub cboWebform_Click()
  
  Dim i As Integer
  Dim j As Integer
  Dim asItems() As String
  
  msElement = cboWebform.Text
'  If cboWebform.ListIndex = -1 Then Exit Sub
'  msElement = masWebformIdentifiers(cboWebform.ItemData(cboWebform.ListIndex))

  cboItem.Clear
  
  For i = 1 To UBound(mwfElements)
    If mwfElements(i).Visible Then
      If mwfElements(i).ElementType = elem_WebForm Then
        If mwfElements(i).Identifier = msElement Then
          asItems = mwfElements(i).Items
          
          For j = 1 To UBound(asItems, 2)
            
            Select Case asItems(2, j)
              Case giWFFORMITEM_BUTTON, _
                giWFFORMITEM_INPUTVALUE_CHAR, _
                giWFFORMITEM_INPUTVALUE_NUMERIC, _
                giWFFORMITEM_INPUTVALUE_LOGIC, _
                giWFFORMITEM_INPUTVALUE_DATE, _
                giWFFORMITEM_INPUTVALUE_GRID, _
                giWFFORMITEM_INPUTVALUE_DROPDOWN, _
                giWFFORMITEM_INPUTVALUE_LOOKUP, _
                giWFFORMITEM_INPUTVALUE_OPTIONGROUP, _
                giWFFORMITEM_INPUTVALUE_FILEUPLOAD
              
                  cboItem.AddItem asItems(9, j)
                  cboItem.ItemData(cboItem.NewIndex) = j
                  
            End Select
          Next j
          
        End If
      End If
    End If
  Next i
  
  ' AE20080227 Fault #12957
  ' AE20080221 Fault #12917
'  If cboItem.ListCount > 0 Then
'    cboItem.ListIndex = 0
'  Else
  
  If cboItem.ListCount <= 0 Then
    cboItem.AddItem "<None>"
  End If
  cboItem.ListIndex = 0
    
End Sub

Private Sub cmdCancel_Click()
  miChoice = vbCancel
  UnLoad Me
End Sub

Private Sub cmdOK_Click()
  miChoice = vbOK
  UnLoad Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode <> vbFormCode Then
    miChoice = vbCancel
  End If
  
End Sub

Public Sub Initialise(pwfElements As Object, _
  Optional piType As WorkflowFindUsageOption, _
  Optional psElement As String, _
  Optional psItem As String, _
  Optional piInitiationType As WorkflowInitiationTypes)
 
  Dim ctlWFElement As VB.Control
  ReDim mwfElements(0)
  
  miInitiationType = piInitiationType
  If (piType = wfNone) Then
    piType = wfRecSelType
  End If
  
  If (piInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL) And (piType = wfRecSelType) Then
    piType = wfElement
  End If
  
  For Each ctlWFElement In pwfElements
    ReDim Preserve mwfElements(UBound(mwfElements) + 1)
    Set mwfElements(UBound(mwfElements)) = ctlWFElement
  Next ctlWFElement
    
  FormatScreen
  PopulateCombos
  PopulateDefaults piType, psElement, psItem
  
End Sub

Private Sub PopulateCombos()

  Dim i As Integer
  
  ReDim masElementIdentifiers(0)
  ReDim masWebformIdentifiers(0)
  ReDim masItemIdentifiers(0)
  
  cboElement.Clear
  
  For i = 1 To UBound(mwfElements)
    If mwfElements(i).Visible Then
      If mwfElements(i).ElementType = elem_WebForm Then
      
        cboElement.AddItem mwfElements(i).Identifier
        cboElement.ItemData(cboElement.NewIndex) = i
'        ReDim Preserve masElementIdentifiers(UBound(masElementIdentifiers) + 1)
'        masElementIdentifiers(UBound(masElementIdentifiers)) = mwfElements(i).Identifier
'        cboElement.ItemData(cboElement.NewIndex) = UBound(masElementIdentifiers)
        
        cboWebform.AddItem mwfElements(i).Identifier
        cboWebform.ItemData(cboWebform.NewIndex) = i
'        ReDim Preserve masWebformIdentifiers(UBound(masWebformIdentifiers) + 1)
'        masWebformIdentifiers(UBound(masWebformIdentifiers)) = mwfElements(i).Identifier
'        cboWebform.ItemData(cboWebform.NewIndex) = UBound(masWebformIdentifiers)
      End If
      
      If mwfElements(i).ElementType = elem_StoredData Then
        
        cboElement.AddItem mwfElements(i).Identifier
        cboElement.ItemData(cboElement.NewIndex) = i
'        ReDim Preserve masElementIdentifiers(UBound(masElementIdentifiers) + 1)
'        masElementIdentifiers(UBound(masElementIdentifiers)) = mwfElements(i).Identifier
'        cboElement.ItemData(cboElement.NewIndex) = UBound(masElementIdentifiers)
      End If
    End If
  Next
  
  If cboElement.ListCount <= 0 Then
    cboElement.AddItem "<None>"
  End If
  
  If cboWebform.ListCount <= 0 Then
    cboWebform.AddItem "<None>"
  End If
 
End Sub



Private Sub SelectOption(pwfOption As WorkflowFindUsageOption)
  miSelection = pwfOption
  
  Select Case miSelection
    Case wfRecSelType
      lblElement.Enabled = False
      cboElement.Enabled = False
      cboElement.BackColor = vbButtonFace ' AE20080221 Fault #12914
      lblWebform.Enabled = False
      cboWebform.Enabled = False
      cboWebform.BackColor = vbButtonFace ' AE20080221 Fault #12914
      lblItem.Enabled = False
      cboItem.Enabled = False
      cboItem.BackColor = vbButtonFace ' AE20080221 Fault #12914
      
      cmdOk.Enabled = True
      
    Case wfElement
      If cboElement.ListCount > 0 Then
        lblElement.Enabled = True
        cboElement.Enabled = True
        cboElement.BackColor = vbWindowBackground ' AE20080221 Fault #12914
        lblWebform.Enabled = False
        cboWebform.Enabled = False
        cboWebform.BackColor = vbButtonFace ' AE20080221 Fault #12914
        lblItem.Enabled = False
        cboItem.Enabled = False
        cboItem.BackColor = vbButtonFace ' AE20080221 Fault #12914
      End If
      
      cboWebform.ListIndex = -1
      cboItem.ListIndex = -1
      
      If cboElement.ListCount > 0 Then
        cboElement.ListIndex = 0
      End If
      
      ' AE20080317 Fault #13038
      'cmdOK.Enabled = Not (Trim(cboElement.Text) = vbNullString)
      cmdOk.Enabled = Not (Trim(cboElement.Text) = "<None>")
            
    Case wfWebFormItem
      lblElement.Enabled = False
      cboElement.Enabled = False
      cboElement.BackColor = vbButtonFace ' AE20080221 Fault #12914
      lblWebform.Enabled = True
      cboWebform.Enabled = True
      cboWebform.BackColor = vbWindowBackground ' AE20080221 Fault #12914
      lblItem.Enabled = True
      cboItem.Enabled = True
      cboItem.BackColor = vbWindowBackground ' AE20080221 Fault #12914
      
      cboElement.ListIndex = -1
      
      If cboWebform.ListCount > 0 Then
        cboWebform.ListIndex = 0
      End If
            
      ' AE20080221 Fault #12917
'      If cboItem.ListCount > 0 Then
'        cboItem.ListIndex = 0
'      End If
            
      ' AE20080317 Fault #13038
      'cmdOK.Enabled = Not ((Trim(cboWebform.Text) = vbNullString) And (Trim(cboItem.Text) = vbNullString))
      cmdOk.Enabled = Not ((Trim(cboWebform.Text) = "<None>") And (Trim(cboItem.Text) = "<None>"))

  End Select
  
End Sub

Public Property Get Element() As String
  Element = msElement
End Property

Public Property Get Item() As String
  Item = msItem
End Property

Private Sub PopulateDefaults(piType As WorkflowFindUsageOption, psElement As String, psItem As String)

  Dim iIndex As Integer

  Select Case piType
    Case wfRecSelType
      optCriteria(0).value = True
      
    Case wfElement
      optCriteria(1).value = True
      
      If cboElement.ListCount > 0 Then
        For iIndex = 0 To cboElement.ListCount - 1
          If cboElement.List(iIndex) = psElement Then
            cboElement.ListIndex = iIndex
            Exit For
          End If
        Next iIndex
      End If
      
    Case wfWebFormItem
      optCriteria(2).value = True
      
      If cboWebform.ListCount > 0 Then
          For iIndex = 0 To cboWebform.ListCount - 1
            If cboWebform.List(iIndex) = psElement Then
              cboWebform.ListIndex = iIndex
              Exit For
            End If
          Next iIndex
      
        If cboItem.ListCount > 0 Then
          For iIndex = 0 To cboItem.ListCount - 1
            If cboItem.List(iIndex) = psItem Then
              cboItem.ListIndex = iIndex
              Exit For
            End If
          Next iIndex
        End If
      End If
  
  End Select
  
End Sub

Private Sub optCriteria_Click(Index As Integer)
  Select Case Index
    Case 0
      SelectOption wfRecSelType
    Case 1
      SelectOption wfElement
    Case 2
      SelectOption wfWebFormItem
  End Select

End Sub


