VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "coa_spinner.ocx"
Begin VB.Form frmWorkflowWFTabOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Order"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5073
   Icon            =   "frmWorkflowWFTabOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Page :"
      Height          =   825
      Left            =   150
      TabIndex        =   5
      Top             =   100
      Width           =   5000
      Begin VB.ComboBox cboPage 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   3300
      End
      Begin COASpinner.COA_Spinner asrPage 
         Height          =   315
         Left            =   200
         TabIndex        =   6
         Top             =   300
         Width           =   1000
         _ExtentX        =   1773
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
         MaximumValue    =   99999
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5280
      TabIndex        =   3
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5280
      TabIndex        =   4
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   400
      Left            =   5280
      TabIndex        =   1
      Top             =   1100
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   400
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1100
      Visible         =   0   'False
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   5292
      SortKey         =   2
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Name"
         Object.Tag             =   ""
         Text            =   "Control"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "Type"
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Order"
         Object.Tag             =   ""
         Text            =   "Order"
         Object.Width           =   1059
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   28
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":000C
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":055E
            Key             =   "IMG_FILEDOWNLOAD"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":0AB0
            Key             =   "IMG_UNKNOWN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1002
            Key             =   "IMG_FILEUPLOAD"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1554
            Key             =   "IMG_WEBFORM"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1DF8
            Key             =   "IMG_BUTTON"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":234A
            Key             =   "IMG_GRID"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":289C
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":2DEE
            Key             =   "IMG_COMBOBOX"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":3340
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":3892
            Key             =   "IMG_LINE"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":3DE4
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4336
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4888
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4DDA
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":532C
            Key             =   "IMG_CHECKBOX"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":587E
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":5DD0
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":6322
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":6874
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":6DC6
            Key             =   "IMG_PROPERTIES"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":7318
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":786A
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":7DBC
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":830E
            Key             =   "IMG_TEXTBOX"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":8860
            Key             =   "IMG_TOOLBOX"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":8DB2
            Key             =   "IMG_WORKFLOW"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWorkflowWFTabOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfLoading As Boolean
Private mfrmWebForm As frmWorkflowWFDesigner
Private mobjDragItem As ComctlLib.ListItem

Public Property Get Cancelled() As Boolean
  ' Return the form's 'Cancelled' property.
  Cancelled = mfCancelled
  
End Property

Public Property Get CurrentScreen() As frmWorkflowWFDesigner
  ' Return the form's 'CurrentScreen' property.
  Set CurrentScreen = mfrmWebForm
End Property
Public Property Set CurrentScreen(pFrmScrDes As frmWorkflowWFDesigner)
  ' Return the form's 'CurrentScreen' property.
  Set mfrmWebForm = pFrmScrDes
End Property

Public Property Get Loading() As Boolean
  ' Return the 'Loading' property.
  Loading = mfLoading
 
End Property

Private Sub asrPage_Change()
  ' Display the required page information.
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim iListNo As Integer

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  If Not mfLoading Then
    ' Get the new page number.
    iPageNo = asrPage.value
    
    If cboPage.ListIndex <> (iPageNo) Then cboPage.ListIndex = (iPageNo)
    
    If (iPageNo >= ListView1.LBound) And (iPageNo <= ListView1.UBound) Then
      For iListNo = ListView1.LBound To ListView1.UBound
        If iListNo = iPageNo Then
          ListView1(iListNo).Visible = True
          ListView1(iListNo).ZOrder 0
        Else
          ListView1(iListNo).Visible = False
        End If
      Next iListNo
    
      If ListView1(iPageNo).ListItems.Count > 0 Then
        ListView1(iPageNo).SelectedItem = ListView1(iPageNo).ListItems(1)
      End If
    End If
  End If

  iListCount = Me.ListView1(iPageNo).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(iPageNo).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Sub cboPage_Click()
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer

  If Not mfLoading Then
    With cboPage
      If .ListCount > 0 And .ListIndex >= 0 Then
        iPageNo = .ListIndex
        If asrPage.value <> iPageNo Then asrPage.value = iPageNo
      End If
    End With
  End If
  
  iListCount = Me.ListView1(iPageNo).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(iPageNo).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Sub cmdCancel_Click()
  On Error GoTo ErrorTrap
  
  mfCancelled = True
  UnLoad Me

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdDown_Click()
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim iNewPos As Integer
  Dim objItem As ComctlLib.ListItem
  
  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  iPageNo = val(asrPage.Text)
  
  If (iPageNo >= ListView1.LBound) And (iPageNo <= ListView1.UBound) Then
    
    Set objItem = ListView1(iPageNo).SelectedItem
    
    If Not objItem Is Nothing Then
      
      iNewPos = objItem.Index
      
      If iNewPos < ListView1(iPageNo).ListItems.Count Then
        iNewPos = iNewPos + 1
        ChangeItemPosition ListView1(iPageNo), objItem, iNewPos
      End If
    End If
  End If

  iListCount = Me.ListView1(iPageNo).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(iPageNo).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

TidyUpAndExit:
  Set objItem = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdOK_Click()
  
  ' Save the changes to the screen controls.
  On Error GoTo ErrorTrap
  
  Dim iIndex As Integer
  Dim iListNo As Integer
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim iNextIndex As Integer
  Dim sKey As String
  Dim ctlControl As VB.Control
  Dim objItem As ComctlLib.ListItem
  
  iNextIndex = 1
  
  For iListNo = ListView1.LBound To ListView1.UBound
    
    ' For each tabstop control on the current tab page.
    For Each objItem In ListView1(iListNo).ListItems
      
      sKey = Mid(objItem.key, 3)
      
      iWFItemType = val(Left(sKey, InStr(1, sKey, "_") - 1))
      sKey = Mid(sKey, InStr(1, sKey, "_") + 1)
      iIndex = val(sKey)
      
      ' Get the control object
      Select Case iWFItemType
        Case WorkflowWebFormItemTypes.giWFFORMITEM_BUTTON
          Set ctlControl = mfrmWebForm.btnWorkflow(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_CHAR
          Set ctlControl = mfrmWebForm.asrDummyTextBox(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DATE
          Set ctlControl = mfrmWebForm.asrDummyCombo(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOGIC
          Set ctlControl = mfrmWebForm.asrDummyCheckBox(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_NUMERIC
          Set ctlControl = mfrmWebForm.asrDummyTextBox(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_GRID
          Set ctlControl = mfrmWebForm.ASRDummyGrid(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DROPDOWN
          Set ctlControl = mfrmWebForm.asrDummyCombo(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOOKUP
          Set ctlControl = mfrmWebForm.asrDummyCombo(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_OPTIONGROUP
          Set ctlControl = mfrmWebForm.ASRDummyOptions(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_FILEUPLOAD
          Set ctlControl = mfrmWebForm.ASRDummyFileUpload(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_DBFILE
          Set ctlControl = mfrmWebForm.asrDummyLabel(iIndex)
        Case WorkflowWebFormItemTypes.giWFFORMITEM_WFFILE
          Set ctlControl = mfrmWebForm.asrDummyLabel(iIndex)
        Case Else
          Set ctlControl = Nothing
      End Select
      
      ' Set the tab index of the control.
      If Not ctlControl Is Nothing Then
        ctlControl.TabIndex = iNextIndex
        iNextIndex = iNextIndex + 1
      End If
      
    Next objItem
  Next iListNo
  
  mfrmWebForm.IsChanged = True
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set objItem = Nothing
  UnLoad Me
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub cmdUp_Click()
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim iNewPos As Integer
  Dim objItem As ComctlLib.ListItem

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  iPageNo = val(asrPage.Text)
  
  If iPageNo >= ListView1.LBound And iPageNo <= ListView1.UBound Then
  
    Set objItem = ListView1(iPageNo).SelectedItem
    
    If Not objItem Is Nothing Then
    
      iNewPos = objItem.Index
      
      If iNewPos > 1 Then
        iNewPos = iNewPos - 1
        ChangeItemPosition ListView1(iPageNo), objItem, iNewPos
      End If
      
    End If
  End If

  iListCount = Me.ListView1(iPageNo).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(iPageNo).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

TidyUpAndExit:
  Set objItem = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorTrap
  
  Dim iPageCount As Integer
  Dim iCurrentPageNo As Integer
  Dim objTab As MSComctlLib.Tab

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
       
  If mfLoading Then
    If Not mfrmWebForm Is Nothing Then
    
      ' Add an item to the combo for form itself.
      cboPage.AddItem "Page Form"
      cboPage.ItemData(cboPage.NewIndex) = 0
      
      Load ListView1(1)
      GetPageControls 0
    
      iCurrentPageNo = mfrmWebForm.PageNo
      iPageCount = mfrmWebForm.tabPages.Tabs.Count

      ' Are there tab pages
      If iPageCount > 0 Then

        ' Add items to the combo for each tab page.
        For Each objTab In mfrmWebForm.tabPages.Tabs
          With objTab
            cboPage.AddItem .Caption
            cboPage.ItemData(cboPage.NewIndex) = val(.Tag)

            If .Index > ListView1.UBound Then
              Load ListView1(.Index)
            End If

            GetPageControls .Index
          End With
        Next objTab

         ' Disassociate object variables.
         Set objTab = Nothing

         With asrPage
           .MinimumValue = 0
           .MaximumValue = iPageCount
         End With

         cboPage.ListIndex = iCurrentPageNo

         If ListView1(iCurrentPageNo).ListItems.Count > 0 Then
           ListView1(iCurrentPageNo).SelectedItem = ListView1(iCurrentPageNo).ListItems(1)
         End If
      Else
        cboPage.ListIndex = 0
      End If

      ' If the number of pages on the form is one then the combobox is disabled.
      If iPageCount < 1 Then cboPage.Enabled = False
    
      With asrPage
        .value = iCurrentPageNo
        .Enabled = (cboPage.ListCount > 1)
      End With
    
      ListView1(iCurrentPageNo).Visible = True
      ListView1(iCurrentPageNo).ZOrder 0
      
    End If
      
    mfLoading = False
    
  End If

  ' Following line included to highlight the selected item.
  If Not Me.ListView1(iCurrentPageNo).SelectedItem Is Nothing Then
    Me.ListView1(iCurrentPageNo).SelectedItem.Selected = True
  End If
  
  iListCount = Me.ListView1(iCurrentPageNo).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(iCurrentPageNo).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

TidyUpAndExit:
  Set objTab = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Sub
Private Function DisplayBackgroundControl()
' Add an item to the combo for form itself.
'cboPage.AddItem "Background Controls"
'GetPageControls 0
'asrPage.MinimumValue = 0
'asrPage.MaximumValue = 0
'cboPage.ListIndex = 0
End Function
        
Private Sub Form_Initialize()
  mfLoading = True
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
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    mfCancelled = True
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call ClipCursorClear(0&)
End Sub

Private Sub ListView1_Click(Index As Integer)

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  iListCount = Me.ListView1(Index).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(Index).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

End Sub

Private Sub ListView1_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
  On Error GoTo ErrorTrap
  
  Dim iNewPos As Integer
  
  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
 
  Call ClipCursorClear(0&)
  If ListView1(Index).DropHighlight Is Nothing Or mobjDragItem Is ListView1(Index).DropHighlight Then
    ListView1(Index).Drag vbCancel
  Else
    iNewPos = ListView1(Index).DropHighlight.Index
    ChangeItemPosition ListView1(Index), mobjDragItem, iNewPos
  End If

  iListCount = Me.ListView1(Index).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(Index).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

  Set ListView1(Index).DropHighlight = Nothing
  Set mobjDragItem = Nothing
    
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ListView1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)

  On Error GoTo ErrorTrap
  
  Dim objItem As ComctlLib.ListItem
  
  If Not mobjDragItem Is Nothing Then
    'Get the item at the mouse's coordinates.
    Set objItem = ListView1(Index).HitTest(x, y)
    
    'Check if the item at the mouse's coordinates is a control.
    If Not objItem Is Nothing Then
      If Left(objItem.key, 1) = "C" Then
        objItem.EnsureVisible
      Else
        Set objItem = Nothing
      End If
    End If
    
    'Set the DropHighlight node
    Set ListView1(Index).DropHighlight = objItem
  End If

TidyUpAndExit:
  Set objItem = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ListView1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  iListCount = Me.ListView1(Index).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(Index).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

End Sub

Private Sub ListView1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
 
  iListCount = Me.ListView1(Index).ListItems.Count
  
  ' Validation for when there are no tab controls in the page.
  If iListCount > 0 Then
    iSelectedIndex = Me.ListView1(Index).SelectedItem.Index
  Else
    iSelectedIndex = 0
  End If
  
  ' Refreshes the enabled status of the cmdUp and cmdDown controls.
  RefreshUpDown iListCount, iSelectedIndex

End Sub

Private Sub ListView1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorTrap
  
  Dim objItem As ComctlLib.ListItem
  
  If Button = vbLeftButton Then
    'Get the item at the mouse position
    Set objItem = ListView1(Index).HitTest(x, y)
    
    If Not objItem Is Nothing Then
      'Check if this item is a control
      If Left(objItem.key, 1) = "C" Then
        'If this item is not the selected item, make it
        If Not objItem Is ListView1(Index).SelectedItem Then
          Set ListView1(Index).SelectedItem = objItem
        End If
        Set mobjDragItem = objItem
      End If
    End If
  End If
  
TidyUpAndExit:
  Set objItem = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub ListView1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorTrap
  
  Dim typRect As Rect
  
  If Not mobjDragItem Is Nothing Then
    Call GetWindowRect(ListView1(Index).hWnd, typRect)
    Call ClipCursorRect(typRect)
    
    'Begin drag
    'NHRD17042012 Jira HRPRO-2092
    ListView1(Index).DragIcon = mobjDragItem.CreateDragImage
    ListView1(Index).Drag vbBeginDrag
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub RefreshUpDown(iCount As Integer, iSelIndex As Integer)

  ' Refreshes the enabled status of the up and down buttons,

  If iCount <= 1 Then
    cmdUp.Enabled = False
    cmdDown.Enabled = False
  ElseIf iSelIndex = 1 Then
    cmdUp.Enabled = False
    cmdDown.Enabled = True
  ElseIf iSelIndex = iCount Then
    cmdUp.Enabled = True
    cmdDown.Enabled = False
  Else
    cmdUp.Enabled = True
    cmdDown.Enabled = True
  End If

End Sub

Private Sub ListView1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  ListView1(Index).Drag vbEndDrag
  Set mobjDragItem = Nothing
  Call ClipCursorClear(0&)

End Sub

Private Function GetPageControls(piPageNo As Integer) As Boolean
  ' Populate the listview with the controls on the given page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iWFItemType As WorkflowWebFormItemTypes
  Dim iLoop As Integer
  Dim sName As String
  Dim sIcon As String
  Dim sImageKey As String
  Dim ctlControl As VB.Control
  Dim objItem As ComctlLib.ListItem
  
  UI.LockWindow Me.hWnd
   
  ' Clear out the list view.
  With ListView1(piPageNo)
    .ListItems.Clear
    .Sorted = False
  End With
  
  ' For each control in the screen ...
  For Each ctlControl In mfrmWebForm.Controls
  
    If mfrmWebForm.IsWebFormControl(ctlControl) Then
    
      iWFItemType = ctlControl.WFItemType
      
      If (mfrmWebForm.GetControlPageNo(ctlControl) = piPageNo) And _
        (mfrmWebForm.WebFormControl_IsTabStop(iWFItemType)) Then
      
        If (iWFItemType = giWFFORMITEM_DBFILE) _
          Or (iWFItemType = giWFFORMITEM_WFFILE) Then
        
          sName = ctlControl.Caption
        Else
          sName = ctlControl.WFIdentifier
        End If
        
        ' Get the imageList icon for the given control.
        sIcon = "IMG_UNKNOWN"
        sImageKey = ImageKey(iWFItemType)
        For iLoop = 1 To ImageList1.ListImages.Count
          If ImageList1.ListImages(iLoop).key = sImageKey Then
            sIcon = sImageKey
            Exit For
          End If
        Next iLoop
          
        ' Add the current control to the list view.
        Set objItem = ListView1(piPageNo).ListItems.Add(, "C_" & Trim(Str(iWFItemType)) & "_" & _
                                                        Trim(Str(ctlControl.Index)), _
                                                        sName, sIcon, sIcon)
        With objItem
          .SubItems(1) = ControlTypeName(iWFItemType)
          .SubItems(2) = Right(Space(6) & ctlControl.TabIndex, 6)
        End With
        
      End If
    End If
  Next ctlControl
  
  ListView1(piPageNo).Sorted = True
  For Each objItem In ListView1(piPageNo).ListItems
    objItem.SubItems(2) = Right(Space(6) & objItem.Index, 6)
  Next
    
  fOK = True
  
TidyUpAndExit:
  UI.UnlockWindow
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set objItem = Nothing
  GetPageControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ChangeItemPosition(pListView As ComctlLib.ListView, pListItem As ComctlLib.ListItem, piNewPosition As Integer) As Boolean
  ' Change the selected item's position in the listview.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iPos As Integer
  Dim iOldPos As Integer
  Dim objItem As ComctlLib.ListItem
  
  iOldPos = pListItem.Index
  
  UI.LockWindow pListView.Parent.hWnd
  pListView.Sorted = False
  
  For Each objItem In pListView.ListItems
    iPos = objItem.Index
    If piNewPosition < iOldPos Then
      If iPos >= piNewPosition And iPos < iOldPos Then
        iPos = iPos + 1
      End If
    ElseIf piNewPosition > iOldPos Then
      If iPos <= piNewPosition And iPos > iOldPos Then
        iPos = iPos - 1
      End If
    End If
    objItem.SubItems(2) = Right(Space(10) & iPos, 10)
  Next
  ' Disassocate object variables.
  Set objItem = Nothing
  
  pListItem.SubItems(2) = Right(Space(10) & piNewPosition, 10)
  pListView.Sorted = True
  
  pListItem.EnsureVisible
  
  fOK = True
  
TidyUpAndExit:
  UI.UnlockWindow
  ' Disassocate object variables.
  Set objItem = Nothing
  ChangeItemPosition = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ImageKey(piWFItemType As WorkflowWebFormItemTypes) As String
  ' Return the image key for the given control type.
  On Error GoTo ErrorTrap
  
  Select Case piWFItemType
    Case WorkflowWebFormItemTypes.giWFFORMITEM_BUTTON
      ImageKey = "IMG_BUTTON"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_CHAR
      ImageKey = "IMG_TEXTBOX"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DATE
      ImageKey = "IMG_DATE"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOGIC
      ImageKey = "IMG_CHECKBOX"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_NUMERIC
      ImageKey = "IMG_NUMERIC"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_GRID
      ImageKey = "IMG_GRID"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DROPDOWN
      ImageKey = "IMG_COMBOBOX"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOOKUP
      ImageKey = "IMG_LOOKUP"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      ImageKey = "IMG_RADIO"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_FILEUPLOAD
      ImageKey = "IMG_FILEUPLOAD"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_DBFILE
      ImageKey = "IMG_FILEDOWNLOAD"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_WFFILE
      ImageKey = "IMG_FILEDOWNLOAD"
    Case Else
      ImageKey = "IMG_UNKNOWN"
  End Select
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  ImageKey = "IMG_UNKNOWN"
  Resume TidyUpAndExit
  
End Function

Private Function ControlTypeName(piWFItemType As WorkflowWebFormItemTypes) As String

  ' Return the control type name for the given control type.
  Select Case piWFItemType
    Case WorkflowWebFormItemTypes.giWFFORMITEM_BUTTON
      ControlTypeName = "Button"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_CHAR
      ControlTypeName = "Character"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DATE
      ControlTypeName = "Date"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOGIC
      ControlTypeName = "Logic"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_NUMERIC
      ControlTypeName = "Numeric"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_GRID
      ControlTypeName = "Record Selector"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_DROPDOWN
      ControlTypeName = "Dropdown List"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_LOOKUP
      ControlTypeName = "Lookup"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      ControlTypeName = "Option Group"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_INPUTVALUE_FILEUPLOAD
      ControlTypeName = "File Upload"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_DBFILE
      ControlTypeName = "File Link"
    Case WorkflowWebFormItemTypes.giWFFORMITEM_WFFILE
      ControlTypeName = "File Link"
    Case Else
      ControlTypeName = "Unknown"
  End Select
  
End Function


Private Function CurrentPageContainer(x As Single, y As Single) As Variant
  ' Return the current page container.
  Dim bSelectTab As Boolean
  
  bSelectTab = False
  With mfrmWebForm
    If .tabPages.Tabs.Count > 0 And .tabPages.Selected Then
      If x > .tabPages.ClientLeft And x < .tabPages.ClientLeft + .tabPages.ClientWidth _
        And y > .tabPages.ClientTop And y < .tabPages.ClientTop + .tabPages.ClientHeight Then
          bSelectTab = True
      End If
    End If
    
    If bSelectTab Then
      Set CurrentPageContainer = .objTabContainer(.tabPages.SelectedItem.Tag)
    Else
      Set CurrentPageContainer = Me
    End If
  End With
End Function

