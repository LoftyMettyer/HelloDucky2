VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmWorkflowWFTabOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Order"
   ClientHeight    =   3225
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5280
      TabIndex        =   4
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   400
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   400
      Left            =   5280
      TabIndex        =   2
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4995
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
         Object.Width           =   4057
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
         Object.Width           =   882
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   1290
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
            Picture         =   "frmWorkflowWFTabOrder.frx":03FF
            Key             =   "IMG_FILEDOWNLOAD"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":07D0
            Key             =   "IMG_UNKNOWN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":0BF3
            Key             =   "IMG_FILEUPLOAD"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":0FC2
            Key             =   "IMG_WEBFORM"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":16BC
            Key             =   "IMG_BUTTON"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1A83
            Key             =   "IMG_GRID"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":1E3D
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":2209
            Key             =   "IMG_COMBOBOX"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":25DB
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":29F5
            Key             =   "IMG_LINE"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":2D5C
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":317A
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":3543
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":38DC
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":3CBE
            Key             =   "IMG_CHECKBOX"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4094
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4466
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4870
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":4C58
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":503F
            Key             =   "IMG_PROPERTIES"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":5418
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":5811
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":5BC1
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":5F72
            Key             =   "IMG_TEXTBOX"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":637C
            Key             =   "IMG_TOOLBOX"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFTabOrder.frx":6755
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

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As Rect) As Long
Private Declare Function ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

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
  
  iPageNo = 0
  
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
      
      iWFItemType = Val(Left(sKey, InStr(1, sKey, "_") - 1))
      sKey = Mid(sKey, InStr(1, sKey, "_") + 1)
      iIndex = Val(sKey)
      
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
  
  iPageNo = 0
  
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
  Dim objTab As ComctlLib.Tab

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  If mfLoading Then
    If Not mfrmWebForm Is Nothing Then
    
      iCurrentPageNo = 0
      iPageCount = 1
      GetPageControls 0
      
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

Private Sub Form_Initialize()
  mfLoading = True
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

Private Sub ListView1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
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

Private Sub ListView1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

  On Error GoTo ErrorTrap
  
  Dim objItem As ComctlLib.ListItem
  
  If Not mobjDragItem Is Nothing Then
    'Get the item at the mouse's coordinates.
    Set objItem = ListView1(Index).HitTest(X, Y)
    
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

Private Sub ListView1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  
  Dim objItem As ComctlLib.ListItem
  
  If Button = vbLeftButton Then
    'Get the item at the mouse position
    Set objItem = ListView1(Index).HitTest(X, Y)
    
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

Private Sub ListView1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  
  Dim typRect As Rect
  
  If Not mobjDragItem Is Nothing Then
    Call GetWindowRect(ListView1(Index).hWnd, typRect)
    Call ClipCursorRect(typRect)
    
    'Begin drag
    ListView1(Index).DragIcon = ImageList1.ListImages(mobjDragItem.SmallIcon).Picture
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

Private Sub ListView1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
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
      
      If (mfrmWebForm.WebFormControl_IsTabStop(iWFItemType)) Then
      
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
