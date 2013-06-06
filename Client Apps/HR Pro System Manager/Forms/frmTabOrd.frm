VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmTabOrd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Order"
   ClientHeight    =   4245
   ClientLeft      =   1335
   ClientTop       =   3045
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5034
   Icon            =   "frmTabOrd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Index           =   0
      Left            =   150
      TabIndex        =   2
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
   Begin VB.Frame Frame1 
      Caption         =   "Page :"
      Height          =   825
      Left            =   150
      TabIndex        =   7
      Top             =   100
      Width           =   5000
      Begin COASpinner.COA_Spinner asrPage 
         Height          =   315
         Left            =   200
         TabIndex        =   0
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
      Begin VB.ComboBox cboPage 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3300
      End
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   400
      Left            =   5300
      Picture         =   "frmTabOrd.frx":000C
      TabIndex        =   4
      Top             =   1605
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   400
      Left            =   5300
      TabIndex        =   3
      Top             =   1100
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5300
      TabIndex        =   6
      Top             =   3700
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5300
      TabIndex        =   5
      Top             =   3200
      Width           =   1200
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5300
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":0286
            Key             =   "IMG_CHECK"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":07D8
            Key             =   "IMG_NAVIGATION"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":0B2A
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":107C
            Key             =   "IMG_COMBO"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":15CE
            Key             =   "IMG_TEXT"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":1B20
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":2072
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":25C4
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":2B16
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":3068
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":35BA
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":3B0C
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":405E
            Key             =   "IMG_UNKNOWN"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTabOrd.frx":45B0
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTabOrd"
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

Private gfCancelled As Boolean
Private gfLoading As Boolean
Private gFrmScreen As frmScrDesigner2
Private objDragItem As ComctlLib.ListItem

Public Property Get Cancelled() As Boolean
  ' Return the form's 'Cancelled' property.
  Cancelled = gfCancelled
  
End Property

Public Property Get CurrentScreen() As frmScrDesigner2
  ' Return the form's 'CurrentScreen' property.
  Set CurrentScreen = gFrmScreen

End Property

Public Property Set CurrentScreen(pFrmScrDes As frmScrDesigner2)
  ' Return the form's 'CurrentScreen' property.
  Set gFrmScreen = pFrmScrDes
  
End Property

Public Property Get Loading() As Boolean
  ' Return the 'Loading' property.
  Loading = gfLoading
  
End Property


Private Sub asrPage_Change()
  ' Display the required page information.
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim iListNo As Integer

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  If Not gfLoading Then
  
    ' Get the new page number.
    iPageNo = val(asrPage.Text)
    
    If cboPage.ListIndex <> (iPageNo - 1) Then cboPage.ListIndex = (iPageNo - 1)
    
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

  If Not gfLoading Then
    With cboPage
      If .ListCount > 0 And .ListIndex >= 0 Then
        iPageNo = .ListIndex + 1
        If (val(asrPage.Text) <> iPageNo) Then asrPage.Text = Trim(Str(iPageNo))
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
  
  gfCancelled = True
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
  Dim iControlType As Long
  Dim iNextIndex As Integer
  Dim sKey As String
  Dim ctlControl As VB.Control
  Dim objItem As ComctlLib.ListItem
  
  gFrmScreen.tabPages.TabIndex = 0
  iNextIndex = 1
  
  ' For each tab page.
  For iListNo = ListView1.LBound To ListView1.UBound
    
    ' For each tabstop control on the current tab page.
    For Each objItem In ListView1(iListNo).ListItems
      
      sKey = Mid(objItem.key, 3)
      
      iControlType = val(Left(sKey, InStr(1, sKey, "_") - 1))
      sKey = Mid(sKey, InStr(1, sKey, "_") + 1)
      iIndex = val(sKey)
      
      ' Get the control object
      Select Case iControlType
        Case giCTRL_CHECKBOX
          Set ctlControl = gFrmScreen.asrDummyCheckBox(iIndex)
        Case giCTRL_COMBOBOX
          Set ctlControl = gFrmScreen.asrDummyCombo(iIndex)
        Case giCTRL_IMAGE
          Set ctlControl = gFrmScreen.asrDummyImage(iIndex)
        Case giCTRL_PHOTO
          Set ctlControl = gFrmScreen.asrDummyPhoto(iIndex)
        Case giCTRL_LINK
          Set ctlControl = gFrmScreen.asrDummyLink(iIndex)
        Case giCTRL_OLE
          Set ctlControl = gFrmScreen.asrDummyOLEContents(iIndex)
        Case giCTRL_OPTIONGROUP
          Set ctlControl = gFrmScreen.ASRDummyOptions(iIndex)
        Case giCTRL_SPINNER
          Set ctlControl = gFrmScreen.asrDummySpinner(iIndex)
        Case giCTRL_TEXTBOX
          Set ctlControl = gFrmScreen.asrDummyTextBox(iIndex)
        Case giCTRL_LABEL
          Set ctlControl = gFrmScreen.asrDummyLabel(iIndex)
        Case giCTRL_WORKINGPATTERN
          Set ctlControl = gFrmScreen.ASRCustomDummyWP(iIndex)
        Case giCTRL_NAVIGATION
          Set ctlControl = gFrmScreen.ASRDummyNavigation(iIndex)
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
  
  gFrmScreen.IsChanged = True
  
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
  Dim objTab As ComctlLib.Tab

  ' Variables used for passing count and selected index to RefreshUpDown function.
  Dim iListCount As Integer
  Dim iSelectedIndex As Integer
  
  If gfLoading Then
    If Not gFrmScreen Is Nothing Then
    
      iCurrentPageNo = gFrmScreen.PageNo
      iPageCount = gFrmScreen.tabPages.Tabs.Count
    
      If iPageCount > 0 Then
        ' Add items to the combo for each tab page.
        For Each objTab In gFrmScreen.tabPages.Tabs
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
          .MinimumValue = 1
          .MaximumValue = iPageCount
        End With
      
        cboPage.ListIndex = iCurrentPageNo - 1
        
        If ListView1(iCurrentPageNo).ListItems.Count > 0 Then
          ListView1(iCurrentPageNo).SelectedItem = ListView1(iCurrentPageNo).ListItems(1)
        End If

      Else
        ' Add an item to the combo for form itself.
        cboPage.AddItem ""
        cboPage.ItemData(cboPage.NewIndex) = 0
          
        GetPageControls 0
      
        asrPage.MinimumValue = 0
        asrPage.MaximumValue = 0
      
        cboPage.ListIndex = 0
      End If
      
      ' If the number of pages on the form is one then the combobox is disabled.
      If iPageCount <= 1 Then cboPage.Enabled = False

      With asrPage
        .Text = Trim(Str(iCurrentPageNo))
        .Enabled = (cboPage.ListCount > 1)
      End With
      
      ListView1(iCurrentPageNo).Visible = True
      ListView1(iCurrentPageNo).ZOrder 0
      
    End If
    
    gfLoading = False
  
  End If

  ' Following line included to highlight the selected item.
  Me.ListView1(iCurrentPageNo).SelectedItem.Selected = True
  
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
  gfLoading = True

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
    gfCancelled = True
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
  If ListView1(Index).DropHighlight Is Nothing Or objDragItem Is ListView1(Index).DropHighlight Then
    ListView1(Index).Drag vbCancel
  Else
    iNewPos = ListView1(Index).DropHighlight.Index
    ChangeItemPosition ListView1(Index), objDragItem, iNewPos
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
  Set objDragItem = Nothing
    
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ListView1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
  On Error GoTo ErrorTrap
  
  Dim objItem As ComctlLib.ListItem
  
  If Not objDragItem Is Nothing Then
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
        Set objDragItem = objItem
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
  
  If Not objDragItem Is Nothing Then
    Call GetWindowRect(ListView1(Index).hWnd, typRect)
    Call ClipCursorRect(typRect)
    
    'Begin drag
    ListView1(Index).DragIcon = objDragItem.CreateDragImage
    'ListView1(Index).DragIcon = Me.Icon

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
  Set objDragItem = Nothing
  Call ClipCursorClear(0&)

End Sub

Private Function GetPageControls(piPageNo As Integer) As Boolean
  ' Populate the listview with the controls on the given page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As Long
  Dim iLoop As Integer
  Dim sName As String
  Dim sIcon As String
  Dim sImageKey As String
  Dim ctlControl As VB.Control
  Dim objItem As ComctlLib.ListItem
  Dim bAdd As Boolean
  
  UI.LockWindow Me.hWnd
    
  'JDM - Fault 4808 - 27/11/02 - Load controls off pages not yet loaded
  gFrmScreen.LoadTabPage piPageNo
    
  ' Clear out the list view.
  With ListView1(piPageNo)
    .ListItems.Clear
    .Sorted = False
  End With
  
  ' For each control in the screen ...
  For Each ctlControl In gFrmScreen.Controls
  
    If gFrmScreen.IsScreenControl(ctlControl) Then
    
      iControlType = gFrmScreen.ScreenControl_Type(ctlControl)
      
      If (gFrmScreen.GetControlPageNo(ctlControl) = piPageNo) And _
        (gFrmScreen.ScreenControl_IsTabStop(iControlType)) Then
      
        bAdd = True
      
        If iControlType = giCTRL_NAVIGATION Then
          If (ctlControl.DisplayType = NavigationDisplayType.Hidden Or ctlControl.DisplayType = NavigationDisplayType.Browser) Then
            bAdd = False
          Else
            sName = ctlControl.Caption
          End If
        Else
          sName = ctlControl.ToolTipText
        End If
  
        ' Get the imageList icon for the given control.
        If bAdd Then
          sIcon = "IMG_UNKNOWN"
          sImageKey = ImageKey(iControlType)
          For iLoop = 1 To ImageList1.ListImages.Count
            If ImageList1.ListImages(iLoop).key = sImageKey Then
              sIcon = sImageKey
              Exit For
            End If
          Next iLoop
            
          ' Add the current control to the list view.
          Set objItem = ListView1(piPageNo).ListItems.Add(, "C_" & _
            Trim(Str(iControlType)) & "_" & _
            Trim(Str(ctlControl.Index)), _
            sName, sIcon, sIcon)
          With objItem
            .SubItems(1) = ControlTypeName(iControlType)
            .SubItems(2) = Right(Space(6) & ctlControl.TabIndex, 6)
          End With
        End If
        
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
    objItem.SubItems(2) = Right(Space(6) & iPos, 6)
  Next
  ' Disassocate object variables.
  Set objItem = Nothing
  
  pListItem.SubItems(2) = Right(Space(6) & piNewPosition, 6)
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

Private Function ImageKey(piType As Integer) As String
  ' Return the image key for the given control type.
  On Error GoTo ErrorTrap
  
  Select Case piType
    Case giCTRL_CHECKBOX
      ImageKey = "IMG_CHECK"
    Case giCTRL_COMBOBOX
      ImageKey = "IMG_COMBO"
    Case giCTRL_IMAGE
      ImageKey = "IMG_IMAGE"
    Case giCTRL_OLE
      ImageKey = "IMG_OLE"
    Case giCTRL_OPTIONGROUP
      ImageKey = "IMG_RADIO"
    Case giCTRL_SPINNER
      ImageKey = "IMG_SPINNER"
    Case giCTRL_TEXTBOX
      ImageKey = "IMG_TEXT"
    Case giCTRL_LABEL
      ImageKey = "IMG_LABEL"
    Case giCTRL_FRAME
      ImageKey = "IMG_FRAME"
    Case giCTRL_PHOTO
      ImageKey = "IMG_PHOTO"
    Case giCTRL_LINK
      ImageKey = "IMG_LINK"
    Case giCTRL_WORKINGPATTERN
      ImageKey = "IMG_WORKINGPATTERN"
    Case giCTRL_NAVIGATION
      ImageKey = "IMG_NAVIGATION"
    Case Else
      ImageKey = "IMG_UNKNOWN"
  End Select
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  ImageKey = "IMG_UNKNOWN"
  Resume TidyUpAndExit
  
End Function
Private Function ControlTypeName(piType As Integer) As String
  ' Return the control type name for the given control type.
  Select Case piType
    Case giCTRL_CHECKBOX
      ControlTypeName = "Check Box"
    Case giCTRL_COMBOBOX
      ControlTypeName = "Dropdown List"
    Case giCTRL_IMAGE
      ControlTypeName = "Image"
    Case giCTRL_OLE
      ControlTypeName = "OLE"
    Case giCTRL_OPTIONGROUP
      ControlTypeName = "Option Group"
    Case giCTRL_SPINNER
      ControlTypeName = "Spinner"
    Case giCTRL_TEXTBOX
      ControlTypeName = "Text Box"
    Case giCTRL_LABEL
      ControlTypeName = "Label"
    Case giCTRL_FRAME
      ControlTypeName = "Frame"
    Case giCTRL_PHOTO
      ControlTypeName = "Photo"
    Case giCTRL_LINK
      ControlTypeName = "Link Button"
    Case giCTRL_WORKINGPATTERN
      ControlTypeName = "Working Pattern"
    Case giCTRL_NAVIGATION
      ControlTypeName = "Navigation Control"
    Case Else
      ControlTypeName = "Unknown"
  End Select
  
End Function

