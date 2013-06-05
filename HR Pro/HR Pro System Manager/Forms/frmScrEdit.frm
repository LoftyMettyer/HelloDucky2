VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmScrEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   4065
   ClientLeft      =   2595
   ClientTop       =   1515
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5028
   Icon            =   "frmScrEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabScreenProperties 
      Height          =   3350
      Left            =   100
      TabIndex        =   14
      Top             =   100
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5900
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Defini&tion"
      TabPicture(0)   =   "frmScrEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinitionPage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&History Screens"
      TabPicture(1)   =   "frmScrEdit.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraHistoryScreensPage"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraHistoryScreensPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2850
         Left            =   -74950
         TabIndex        =   15
         Top             =   325
         Width           =   4900
         Begin VB.CommandButton cmdDeselectAll 
            Caption         =   "D&eselect All"
            Height          =   400
            Left            =   3600
            TabIndex        =   20
            Top             =   1450
            Width           =   1200
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Select &All"
            Height          =   400
            Left            =   3600
            TabIndex        =   19
            Top             =   850
            Width           =   1200
         End
         Begin VB.CommandButton cmdSelectDeselect 
            Caption         =   "&Sel/Deselect"
            Height          =   400
            Left            =   3600
            TabIndex        =   18
            Top             =   250
            Width           =   1200
         End
         Begin VB.Frame fraHistoryScreens 
            Caption         =   "History Screens :"
            Height          =   2700
            Left            =   150
            TabIndex        =   16
            Top             =   100
            Width           =   3300
            Begin VB.ListBox listHistoryScreens 
               Height          =   2310
               Left            =   150
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   17
               Top             =   250
               Width           =   3000
            End
         End
      End
      Begin VB.Frame fraDefinitionPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   2850
         Left            =   50
         TabIndex        =   0
         Top             =   325
         Width           =   4900
         Begin VB.CheckBox chkSSIntranet 
            Caption         =   "&Self-service Intranet Screen"
            Height          =   315
            Left            =   200
            TabIndex        =   7
            Top             =   1400
            Width           =   2985
         End
         Begin VB.CheckBox chkQuickEntry 
            Caption         =   "&Quick Access Screen"
            Height          =   315
            Left            =   200
            TabIndex        =   6
            Top             =   1000
            Width           =   2385
         End
         Begin VB.Frame fraIcon 
            Caption         =   "Icon :"
            Height          =   1000
            Left            =   210
            TabIndex        =   8
            Top             =   1800
            Width           =   4550
            Begin VB.CommandButton cmdIconClear 
               Caption         =   "O"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Wingdings 2"
                  Size            =   20.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3420
               MaskColor       =   &H000000FF&
               TabIndex        =   11
               ToolTipText     =   "Clear Path"
               Top             =   300
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.TextBox txtIcon 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   200
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "txtIcon"
               Top             =   300
               Width           =   2910
            End
            Begin VB.CommandButton cmdIcon 
               Caption         =   "..."
               Height          =   315
               Left            =   3105
               TabIndex        =   10
               Top             =   300
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.Image imgIcon 
               Height          =   510
               Left            =   3850
               Top             =   300
               Width           =   510
            End
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1000
            MaxLength       =   255
            TabIndex        =   2
            Text            =   "txtName"
            Top             =   200
            Width           =   3500
         End
         Begin VB.TextBox txtOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1000
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtOrder"
            Top             =   600
            Width           =   3185
         End
         Begin VB.CommandButton cmdOrder 
            Caption         =   "..."
            Height          =   315
            Left            =   4185
            TabIndex        =   5
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   200
            TabIndex        =   1
            Top             =   260
            Width           =   510
         End
         Begin VB.Label lblOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order :"
            Height          =   195
            Left            =   200
            TabIndex        =   3
            Top             =   660
            Width           =   525
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3900
      TabIndex        =   13
      Top             =   3550
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2640
      TabIndex        =   12
      Top             =   3550
      Width           =   1200
   End
End
Attribute VB_Name = "frmScrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables to hold property values
Private blnCancelled As Boolean
Private blnLoading As Boolean
Private lngOrderId As Long
Private glngPictureID As Long
Private lngScreenID As Long
Private lngTableID As Long
Private gfLockedTable As Boolean
Private miTableType As TableTypes

Public Property Get Cancelled() As Boolean
  Cancelled = blnCancelled
End Property

Public Property Get Loading() As Boolean
  Loading = blnLoading
End Property

Public Property Get OrderID() As Long
  
  ' Return the order id property.
  OrderID = lngOrderId

End Property

Public Property Let OrderID(pLngNewID As Long)
  
  ' Set the order id property.
  lngOrderId = pLngNewID

End Property

Public Property Get PictureID() As Long
  ' Return the selected picture's ID.
  PictureID = glngPictureID
  
End Property

Public Property Get ScreenID() As Long
  ScreenID = lngScreenID
End Property

Public Property Let ScreenID(pLngNewID As Long)
  lngScreenID = pLngNewID
End Property

Public Property Get TableID() As Long
  
  ' Return the screen parent table property.
  TableID = lngTableID

End Property

Public Property Let TableID(pLngNewID As Long)

  ' Set the screen parent table property.
  lngTableID = pLngNewID
  
  
  miTableType = iTabParent
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", lngTableID
    
    If Not recTabEdit.NoMatch Then
      miTableType = recTabEdit!TableType
    End If
  End With
  
End Property

Public Property Let LockedTableId(pLngNewID As Long)

  ' Set the locked parent table property.
  If pLngNewID > 0 Then
    gfLockedTable = True
    TableID = pLngNewID
  Else
    gfLockedTable = False
    TableID = 0
  End If
  
End Property


Private Sub cboTables_Click()

'  If cboTables.ListIndex >= 0 Then
'
'    If cboTables.ItemData(cboTables.ListIndex) <> TableID Then
'
'      TableID = cboTables.ItemData(cboTables.ListIndex)
'      OrderID = 0
'      txtOrder.Text = vbNullString
'
'      ' Refresh the history screens listbox.
'      listHistoryScreens_Refresh
'    End If
'
'  End If

End Sub

Private Sub chkQuickEntry_Click()
  RefreshCurrentTab
  
End Sub

Private Sub chkSSIntranet_Click()
  RefreshCurrentTab
End Sub

Private Sub cmdCancel_Click()
  blnCancelled = True
  
  UnLoad Me
End Sub

Private Sub cmdDeselectAll_Click()
  Dim iIndex As Integer
  
  ' Deselect all available history screens.
  For iIndex = 0 To listHistoryScreens.ListCount - 1
    listHistoryScreens.Selected(iIndex) = False
  Next iIndex

  'Go to the top one
  listHistoryScreens.ListIndex = 0
  cmdDeselectAll.Enabled = False
End Sub

Private Sub cmdIcon_Click()
  ' Display the icon selection form.

  frmPictSel.SelectedPicture = glngPictureID
  frmPictSel.PictureType = vbPicTypeIcon

  frmPictSel.Show vbModal
  
  If frmPictSel.SelectedPicture > 0 Then
    glngPictureID = frmPictSel.SelectedPicture
    imgIcon_Refresh
  End If
  
  Set frmPictSel = Nothing
  
End Sub

Private Sub cmdIconClear_Click()
  glngPictureID = frmPictSel.SelectedPicture
  imgIcon_Refresh
End Sub

Private Sub cmdOK_Click()
  ' Validate and save the changes.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim bHasLink As Boolean
  Dim frmForm As Form
  Dim frmForm2 As Form
  Dim fFound As Boolean
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  
  ' Check that a table has been selected.
  fOK = (TableID > 0)
  If Not fOK Then
    MsgBox "No primary table has been specified!", vbOKOnly + vbExclamation, Application.Name
'    If cboTables.Enabled Then
'      cboTables.SetFocus
'    End If
  End If

  ' Check that a screen name has been entered.
  If fOK Then
    fOK = (Len(Trim(txtName.Text)) > 0)
    If Not fOK Then
      MsgBox "Invalid screen name!", vbOKOnly + vbExclamation, Application.Name
      txtName.SetFocus
    End If
  End If
  
  ' Check that the screen name entered is unique.
  If fOK Then
    With recScrEdit
      .Index = "idxName"
      .Seek "=", Trim(txtName.Text), False
      If Not .NoMatch Then
        ' JPD 19/02/2001 - Changed the idxName index to be non-unique.
        Do While (Not .EOF) And fOK
          If (!Name <> Trim(txtName.Text)) Or (!Deleted) Then
            Exit Do
          End If
        
          fOK = (!ScreenID = ScreenID)
          If Not fOK Then
            MsgBox "A screen named '" & Trim(txtName.Text) & "' already exists!", vbOKOnly + vbExclamation, Application.Name
            txtName.SetFocus
          End If
        
          .MoveNext
        Loop
      End If
    End With
  End If
    
  'TM20020219 Fault 3522 - If Quick Entry is chosen then check that the table has
  'at least one column.
  If fOK Then
    If chkQuickEntry.value = vbChecked Then
      bHasLink = False
      With recColEdit
        .Index = "idxTableID"
        .Seek "=", TableID
        If Not .NoMatch Then
        
          Do While (Not .EOF)
            If (!TableID <> TableID) Then
              Exit Do
            End If
            If (!columntype = giCOLUMNTYPE_LINK) Then
              bHasLink = True
            End If
            .MoveNext
          Loop

        End If
      End With
      
      If Not bHasLink Then
        MsgBox "The screen cannot be made a Quick Access screen because the table has no link columns defined.", vbOKOnly + vbExclamation, Application.Name
        chkQuickEntry.SetFocus
        fOK = False
      End If
    End If
  End If
  
  ' If SSIntranet then check that the screen is not already a view screen.
  ' No need to do this check for new screens.
  If fOK And (lngScreenID > 0) Then
    If chkSSIntranet.value = vbChecked Then
      sSQL = "SELECT COUNT(*) AS result" & _
        " FROM tmpViewScreens" & _
        " WHERE screenID = " & CStr(lngScreenID) & _
        " AND deleted = FALSE"
      
      Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If rsTemp!result > 0 Then
        MsgBox "The screen cannot be made a Self-service Intranet screen because it is already defined as a view screen.", _
          vbOKOnly + vbExclamation, Application.Name
        If chkSSIntranet.Enabled Then
          chkSSIntranet.SetFocus
        End If
        fOK = False
      End If
      
      rsTemp.Close
      Set rsTemp = Nothing
    End If
  End If
  
  ' If SSIntranet then check that the no invalid controls are on the screen.
  ' No need to do this check for new screens.
  If fOK And (lngScreenID > 0) Then
    If chkSSIntranet.value = vbChecked Then
      ' See if the screen's designer is open.
      fFound = False
      For Each frmForm In Forms
        With frmForm
          If .Name = "frmScrDesigner2" Then
            If .ScreenID = lngScreenID Then
              fFound = True
              
              ' Check if the screen designer form has any controls on it that are not permitted on SSInt screens.
              If .HasNonSSIntranetControls Then
                MsgBox "The screen cannot be made a Self-service Intranet screen because it contains Image controls.", _
                  vbOKOnly + vbExclamation, Application.Name
                If chkSSIntranet.Enabled Then
                  chkSSIntranet.SetFocus
                End If
                fOK = False
              End If
              
              Exit For
            End If
          End If
        End With
      Next frmForm
      Set frmForm = Nothing
      
'      If fOK And (Not fFound) Then
'        ' Screen Designer not there. We must be just editing the properties of an existing screen.
'        ' Check for tab strip.
'        sSQL = "SELECT COUNT(*) AS result" & _
'          " FROM tmpPageCaptions" & _
'          " WHERE screenID = " & CStr(lngScreenID)
'
'        Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'        If rsTemp!result > 0 Then
'          MsgBox "The screen cannot be made a Self-service Intranet screen because it contains Image controls.", _
'            vbOKOnly + vbExclamation, Application.Name
'          If chkSSIntranet.Enabled Then
'            chkSSIntranet.SetFocus
'          End If
'
'          fOK = False
'        End If
'
'        rsTemp.Close
'        Set rsTemp = Nothing
'      End If
    
      If fOK And (Not fFound) Then
        ' Screen Designer not there. We must be just editing the properties of an existing screen.
        ' Check for OLE, photo, link, and image controls.
'        sSQL = "SELECT COUNT(*) AS result" & _
          " FROM tmpControls" & _
          " WHERE screenID = " & CStr(lngScreenID) & _
          " AND ((controlType = " & CStr(giCTRL_IMAGE) & ")" & _
          "   OR (controlType = " & CStr(giCTRL_OLE) & ")" & _
          "   OR (controlType = " & CStr(giCTRL_PHOTO) & "))"
        
        sSQL = "SELECT COUNT(*) AS result" & _
          " FROM tmpControls" & _
          " WHERE screenID = " & CStr(lngScreenID) & _
          " AND controlType = " & CStr(giCTRL_IMAGE)
        
        Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        If rsTemp!result > 0 Then
          MsgBox "The screen cannot be made a Self-service Intranet screen because it contains Image controls.", _
            vbOKOnly + vbExclamation, Application.Name
          If chkSSIntranet.Enabled Then
            chkSSIntranet.SetFocus
          End If
          fOK = False
        End If
        
        rsTemp.Close
        Set rsTemp = Nothing
      End If
    Else
      'JPD 20040624 Fault 8837
      ' Do not let the screen be changed from being a Self-service Intranet screen if it
      ' already use as a link in the Self-service Intranet module.
      sSQL = "SELECT COUNT(*) AS result" & _
        " FROM tmpSSIntranetLinks" & _
        " WHERE screenID = " & CStr(lngScreenID)
      
      Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If rsTemp!result > 0 Then
        MsgBox "The screen cannot be changed from being a Self-service Intranet screen because it is already associated with a Self-service Intranet link.", _
          vbOKOnly + vbExclamation, Application.Name
        If chkSSIntranet.Enabled Then
          chkSSIntranet.value = vbChecked
          chkSSIntranet.SetFocus
        End If
        
        fOK = False
      End If
      
      rsTemp.Close
      Set rsTemp = Nothing
    End If
  End If
  
  If fOK Then
    fOK = SaveChanges
  End If
    
  If fOK Then
    blnCancelled = False
    UnLoad Me
    
    ' Update the associated Screen Deisgner toolbox (if one exists)
    ' The permitted controls may have changed due to the setting/resetting
    ' of the Self-service Intranet checkbox.
    For Each frmForm In Forms
      If frmForm.Name = "frmScrDesigner2" Then
        If frmForm.ScreenID = lngScreenID Then
          For Each frmForm2 In Forms
            If frmForm2.Name = "frmToolbox" Then
              If frmForm2.CurrentScreen Is frmForm Then
                frmForm2.RefreshControls
              End If
            End If
          Next frmForm2
          Set frmForm2 = Nothing
        End If
      End If
    Next frmForm
    Set frmForm = Nothing
  End If
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdOrder_Click()
  Dim objOrder As Order
  
  Set objOrder = New Order
  objOrder.OrderID = Me.OrderID
  objOrder.TableID = Me.TableID
  objOrder.OrderType = giORDERTYPE_DYNAMIC
  If objOrder.SelectOrder Then
    Me.OrderID = objOrder.OrderID
    txtOrder.Text = objOrder.OrderName
  End If
  Set objOrder = Nothing
  
End Sub

Private Sub cmdSelectAll_Click()

  Dim iIndex As Integer
  
  ' Select all available history screens.
  For iIndex = 0 To listHistoryScreens.ListCount - 1
    listHistoryScreens.Selected(iIndex) = True
  Next iIndex
  
  'Go to the top one
  listHistoryScreens.ListIndex = 0
  cmdSelectAll.Enabled = False
  
End Sub

Private Sub cmdSelectDeselect_Click()
    
  ' Toggle the selection state of the current history screen.
  With listHistoryScreens
    .Selected(.ListIndex) = Not .Selected(.ListIndex)
  End With
  
  RefreshHistoryScreensTab
  
End Sub

Private Sub Form_Activate()

  Dim fQuickEntry As Boolean
  Dim fSSIntranet As Boolean
  
  If Loading Then
    
    SSTabScreenProperties.Tab = 0
  
    If ScreenID > 0 Then
      
      With recScrEdit
        
        .Index = "idxScreenID"
        .Seek "=", ScreenID
        
        If Not .NoMatch Then
          
          TableID = .Fields("tableID")
          OrderID = .Fields("orderID")
          glngPictureID = IIf(IsNull(.Fields("pictureID")), 0, .Fields("pictureID"))
          txtName.Text = Trim(.Fields("name"))
          Me.Caption = "Properties - " & Trim(.Fields("name"))
          
          fQuickEntry = IIf(IsNull(.Fields("quickEntry")), False, .Fields("quickEntry"))
          chkQuickEntry.value = IIf(fQuickEntry, vbChecked, vbUnchecked)
          
          fSSIntranet = IIf(IsNull(.Fields("SSIntranet")), False, .Fields("SSIntranet"))
          chkSSIntranet.value = IIf(fSSIntranet, vbChecked, vbUnchecked)
        End If
      
      End With
      
      'cboTables.Enabled = False
    
    Else
      'cboTables.Enabled = Not gfLockedTable
      Me.Caption = "Properties - New Screen"
    End If
        
    ' Populate the combos, textboxes and listboxes.
    cboTables_Refresh
    txtOrder_Refresh
    imgIcon_Refresh
    listHistoryScreens_Refresh
    
    ' Initially set the current tab page to first tab page
    SSTabScreenProperties.Tab = 0
    'Refresh current tab page
    RefreshCurrentTab
    RefreshHistoryScreensTab
    'Set focus to column name textbox
    If txtName.Enabled Then
      txtName.SetFocus
    End If

    blnLoading = False
  
  End If

End Sub

Private Sub Form_Initialize()
  blnLoading = True
  blnCancelled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    UnLoad Me
End Select
  
End Sub

Private Sub Form_Load()
  
  Dim objControl As VB.Control
  
  For Each objControl In Me.Controls
    If TypeOf objControl Is TextBox Then
      objControl.Text = vbNullString
    End If
  Next
  Set objControl = Nothing
  
  ' Ensure the frames on each of the tab pages have the same
  ' background colour as the tab pages themselves.
  fraDefinitionPage.BackColor = SSTabScreenProperties.BackColor
  fraHistoryScreensPage.BackColor = SSTabScreenProperties.BackColor
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

End Sub

Private Function txtOrder_Refresh() As Boolean
  Dim objOrder As Order
  
  If OrderID > 0 Then
    Set objOrder = New Order
    objOrder.OrderID = OrderID
    If objOrder.ConstructOrder Then
      txtOrder.Text = objOrder.OrderName
    End If
    Set objOrder = Nothing
  End If
End Function

Private Function imgIcon_Refresh() As Boolean
  Dim strFileName As String
  
  If PictureID > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", PictureID
      If Not .NoMatch Then
        strFileName = ReadPicture
        Set imgIcon.Picture = LoadPicture(strFileName)
        Kill strFileName
        
        txtIcon.Text = .Fields("Name")
      End If
    End With
  Else
    Set imgIcon.Picture = LoadPicture(vbNullString)
    txtIcon.Text = vbNullString
  End If
End Function

Private Function cboTables_Refresh() As Boolean

'  With recTabEdit
'
'    .Index = "idxName"
'
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While Not .EOF
'
'      If Not .Fields("deleted") Then
'        cboTables.AddItem .Fields("tableName")
'        cboTables.ItemData(cboTables.NewIndex) = .Fields("tableID")
'
'        If TableID = .Fields("tableID") Then
'          cboTables.ListIndex = cboTables.NewIndex
'        End If
'
'      End If
'
'      .MoveNext
'
'    Loop
'
'  End With
'
'  With cboTables
'
'    If .ListCount > 0 Then
'      If .ListIndex < 0 Then
'        .ListIndex = 0
'      End If
'    End If
'
'  End With

End Function

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub imgIcon_DblClick()
  cmdIcon_Click
End Sub


Private Sub listHistoryScreens_Click()

  ' JDM - 21/08/02 - Fault 4284 - Not refreshing buttons
  RefreshHistoryScreensTab

End Sub

Private Sub listHistoryScreens_ItemCheck(Item As Integer)

  RefreshHistoryScreensTab

End Sub


Private Sub listHistoryScreens_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    UnLoad Me
  End If
End Sub

Private Sub SSTabScreenProperties_Click(PreviousTab As Integer)
  Dim fDefinitionPage As Boolean
  Dim fHistoryScreensPage As Boolean
  
  ' Determine which tab is selected.
  fDefinitionPage = (SSTabScreenProperties.Tab = 0)
  fHistoryScreensPage = (SSTabScreenProperties.Tab = 1)
  
  ' Enable, and make visuble the selected tab.
  fraDefinitionPage.Enabled = fDefinitionPage
  fraDefinitionPage.Visible = fDefinitionPage
  
  fraHistoryScreensPage.Enabled = fHistoryScreensPage
  fraHistoryScreensPage.Visible = fHistoryScreensPage

  ' Refresh the current tab page
  RefreshCurrentTab
  
  cmdDeselectAll.Enabled = AnySelected

End Sub
Private Sub RefreshDefinitionTab()

  If chkQuickEntry.value = vbChecked Then
    chkSSIntranet.value = vbUnchecked
  End If
  
  If chkSSIntranet.value = vbChecked Then
    chkQuickEntry.value = vbUnchecked
  End If
  
  'JPD 20030909 Fault 6210
  If (Not Application.SelfServiceIntranetModule) Then
    If chkSSIntranet.value Then
      MsgBox "This screen cannot be a Self-service Intranet screen as you are not licenced to use the Intranet module.", _
        vbOKOnly + vbExclamation, Application.Name
          
      chkSSIntranet.value = vbUnchecked
    End If
  End If
  
  'NPG20080421 Fault 12982
  cmdOrder.Enabled = (chkSSIntranet.value = vbUnchecked)
  
  'JPD 20030909 Fault 6547
  If chkSSIntranet.value = vbChecked Then
      'NPG20080124 Fault 12870
      'Me.OrderID = 0
      'txtOrder.Text = ""
      
      'NPG20080421 Fault 12982
    If IsParentScreen() = True Then
      Me.OrderID = 0
      txtOrder.Text = ""
      cmdOrder.Enabled = False
    Else
      cmdOrder.Enabled = True
    End If
    
    glngPictureID = 0
    imgIcon_Refresh
  End If
  'NPG20080124 Fault 12870
  ' cmdOrder.Enabled = (chkSSIntranet.Value = vbUnchecked)
  cmdIcon.Enabled = (chkSSIntranet.value = vbUnchecked)
  cmdIconClear.Enabled = (chkSSIntranet.value = vbUnchecked And Trim(txtIcon.Text) <> vbNullString)
  
  'JPD 20050111 Fault 9697
  chkSSIntranet.Enabled = (chkQuickEntry.value = vbUnchecked) _
    And (Application.SelfServiceIntranetModule)
    
  chkQuickEntry.Enabled = (chkSSIntranet.value = vbUnchecked)
  
End Sub

Private Sub RefreshHistoryScreensTab()
  Dim iLoop As Integer
  
  ' Enable the history screens listbox only if there are items.
  With listHistoryScreens
    If .ListCount > 0 Then
      .Enabled = (chkSSIntranet.value = vbUnchecked)
    
      If .Selected(.ListIndex) Then
        cmdSelectDeselect.Caption = "&Deselect"
      Else
        cmdSelectDeselect.Caption = "&Select"
      End If
    Else
      .Enabled = False
    End If
    
    cmdSelectDeselect.Enabled = .Enabled
    cmdSelectAll.Enabled = .Enabled
    cmdDeselectAll.Enabled = .Enabled
  End With

  'JPD 20030909 Fault 6547
  listHistoryScreens.Enabled = (chkSSIntranet.value = vbUnchecked)
  If (chkSSIntranet.value = vbChecked) Then
    For iLoop = 0 To listHistoryScreens.ListCount - 1
      listHistoryScreens.Selected(iLoop) = False
    Next iLoop
  End If
  
End Sub

Private Sub txtIcon_Change()
  cmdIconClear.Enabled = (chkSSIntranet.value = vbUnchecked And Trim(txtIcon.Text) <> vbNullString)
End Sub

Private Sub txtIcon_GotFocus()
  cmdIcon.SetFocus
End Sub

Private Sub txtOrder_GotFocus()
  cmdOrder.SetFocus
End Sub


Private Sub listHistoryScreens_Refresh()
  Dim sSQL As String
  
  ' Trap errors.
  On Error GoTo ErrorTrap
  
  ' Clear the current contents of the listbox.
  listHistoryScreens.Clear
  
  UI.LockWindow Me.hWnd
    
  recScrEdit.Index = "idxTableID"
  recHistScrEdit.Index = "idxParentHistory"
  recRelEdit.Index = "idxParentID"
    
  If Not (recRelEdit.BOF And recRelEdit.EOF) Then
    recRelEdit.MoveFirst
  End If
    
  Do While Not recRelEdit.EOF
    
    ' Add items to the listbox for each screen of each child of the
    ' current screens parent table.
    If recRelEdit!parentID = TableID Then
      
      recScrEdit.Seek "=", recRelEdit!childID

      If Not recScrEdit.NoMatch Then
    
        Do While Not recScrEdit.EOF
      
          'If no more items for this order exit loop
          If recScrEdit!TableID <> recRelEdit!childID Then
            Exit Do
          End If
          
          ' TM110701 - Fault 540. Do not populate the list with Quick Entry screens.
          If Not recScrEdit!Deleted And Not recScrEdit!QuickEntry Then
            listHistoryScreens.AddItem recScrEdit!Name
            listHistoryScreens.ItemData(listHistoryScreens.NewIndex) = recScrEdit!ScreenID
          
            ' Select the new item if is already a history screen of the current screen.
            recHistScrEdit.Seek "=", ScreenID, recScrEdit!ScreenID
            If Not recHistScrEdit.NoMatch Then
              listHistoryScreens.Selected(listHistoryScreens.NewIndex) = True
            End If
          End If
          
          recScrEdit.MoveNext
        
        Loop
      
      End If
        
    End If
      
    recRelEdit.MoveNext
      
  Loop
    
Exit_listHistoryScreens_Refresh:
  
  UI.UnlockWindow
  listHistoryScreens.Refresh
  
  Exit Sub
  
ErrorTrap:
  
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  
  Resume Exit_listHistoryScreens_Refresh

End Sub
Private Sub RefreshCurrentTab()
  
  'Refresh the controls on the active tab page
  Select Case SSTabScreenProperties.Tab
  
    Case 0 ' Definition tab.
      RefreshDefinitionTab
      
    Case 1 ' History Screens tab.
      RefreshHistoryScreensTab
      If listHistoryScreens.ListCount > 0 Then listHistoryScreens.ListIndex = 0
  End Select
  
  Me.Refresh
  
End Sub



Private Function SaveChanges() As Boolean
  ' Save the changes.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  
  fOK = True
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  With recScrEdit
    If ScreenID > 0 Then
      ' Modify an existing screen record if it exists.
      .Index = "idxScreenID"
      .Seek "=", ScreenID
      
      If Not .NoMatch Then
        .Edit
        !Name = Trim(txtName.Text)
        !OrderID = OrderID
        !PictureID = PictureID
        !QuickEntry = (chkQuickEntry.value = vbChecked)
        !SSIntranet = (chkSSIntranet.value = vbChecked)
        !Changed = True
        .Update
      End If
    Else
      ' Create and initialise a new screen record if required.
      Me.ScreenID = Database.UniqueColumnValue("tmpScreens", "screenID")
      .AddNew
      !ScreenID = Me.ScreenID
      !TableID = TableID
      !OrderID = OrderID
      !Name = Trim(txtName.Text)
      !PictureID = PictureID
      !QuickEntry = (chkQuickEntry.value = vbChecked)
      !SSIntranet = (chkSSIntranet.value = vbChecked)
      !GridX = 40
      !GridY = 40
      !AlignToGrid = True
    
      !dfltForeColour = vbBlack
      !dfltFontName = "Verdana"
      !dfltFontSize = 8
      !dfltFontBold = False
      !dfltFontItalic = False
      
      !Height = 4900
      !Width = 7100
      !New = True
      !Changed = False
      !Deleted = False
      .Update
    End If
  End With
  
  ' Delete all existing components for this expression from the database.
  daoDb.Execute "DELETE FROM tmpHistoryScreens WHERE parentScreenID=" & ScreenID, dbFailOnError
  
  ' Save the history screen associations.
  For iIndex = 0 To listHistoryScreens.ListCount - 1
    If listHistoryScreens.Selected(iIndex) Then
      With recHistScrEdit
        .AddNew
        !id = Database.UniqueColumnValue("tmpHistoryScreens", "ID")
        !parentScreenID = ScreenID
        !historyScreenID = listHistoryScreens.ItemData(iIndex)
        .Update
      End With
    End If
  Next iIndex
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  SaveChanges = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Private Function AnySelected() As Boolean
Dim iIndex As Integer
  ' Save the history screen associations.
  For iIndex = 0 To listHistoryScreens.ListCount - 1
    If listHistoryScreens.Selected(iIndex) Then
      AnySelected = True
      Exit For
    End If
  Next iIndex
End Function


Private Function IsParentScreen() As Boolean
'NPG20080421 Fault 12982
Dim sSQL As String
Dim rsTemp As DAO.Recordset
If chkSSIntranet.value = vbChecked Then
      sSQL = "SELECT Count(*) as Result" & _
        " FROM tmpTables" & _
        " WHERE TableID = " & CStr(TableID) & _
        " AND TableType = 1"
      
      Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If rsTemp!result > 0 Then
        IsParentScreen = True
      Else
        IsParentScreen = False
      End If
      
      rsTemp.Close
      Set rsTemp = Nothing
    End If
End Function



