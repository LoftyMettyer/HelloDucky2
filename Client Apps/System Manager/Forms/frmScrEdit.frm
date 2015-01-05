VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScrEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   6855
   ClientLeft      =   2595
   ClientTop       =   1515
   ClientWidth     =   9630
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
   ScaleHeight     =   6855
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDefinition 
      Caption         =   "Definition : "
      Height          =   2310
      Left            =   45
      TabIndex        =   18
      Top             =   90
      Width           =   9510
      Begin VB.TextBox txtDefaultScreenFont 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5985
         TabIndex        =   28
         Top             =   1080
         Width           =   2760
      End
      Begin VB.CommandButton cmdDefaultScreenFont 
         Caption         =   "..."
         Height          =   315
         Left            =   8745
         TabIndex        =   27
         ToolTipText     =   "Select Path"
         Top             =   1080
         Width           =   315
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   3255
      End
      Begin VB.TextBox txtDescription 
         Height          =   1000
         Left            =   1320
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
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
         Left            =   9060
         MaskColor       =   &H000000FF&
         TabIndex        =   7
         ToolTipText     =   "Clear Path"
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txtIcon 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5985
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "txtIcon"
         Top             =   675
         Width           =   2760
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "..."
         Height          =   315
         Left            =   8745
         TabIndex        =   6
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CheckBox chkSSIntranet 
         Caption         =   "Self-service &Intranet Screen"
         Height          =   315
         Left            =   5985
         TabIndex        =   9
         Top             =   1815
         Width           =   2985
      End
      Begin VB.CheckBox chkQuickEntry 
         Caption         =   "&Quick Access Screen"
         Height          =   315
         Left            =   5985
         TabIndex        =   8
         Top             =   1500
         Width           =   2385
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1335
         MaxLength       =   255
         TabIndex        =   2
         Text            =   "txtName"
         Top             =   270
         Width           =   3255
      End
      Begin VB.TextBox txtOrder 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5985
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "txtOrder"
         Top             =   270
         Width           =   2760
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "..."
         Height          =   315
         Left            =   8745
         TabIndex        =   5
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Label lblDefaultScreenFont 
         Caption         =   "Default Font :"
         Height          =   195
         Left            =   4725
         TabIndex        =   26
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label lblCategory 
         Caption         =   "Category :"
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon :"
         Height          =   195
         Left            =   4725
         TabIndex        =   24
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   1140
         Width           =   1170
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   315
         Width           =   870
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Left            =   4725
         TabIndex        =   21
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.Frame fraHistoryScreens 
      Caption         =   "History Screens :"
      Height          =   3690
      Left            =   90
      TabIndex        =   17
      Top             =   2520
      Width           =   9465
      Begin VB.CheckBox chkGroupByCategory 
         Caption         =   "&Group screens by category"
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   3240
         Width           =   2805
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move &Down"
         Enabled         =   0   'False
         Height          =   420
         Left            =   7965
         TabIndex        =   15
         Top             =   2610
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move &Up"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7965
         TabIndex        =   14
         Top             =   2070
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdDeselectAll 
         Caption         =   "D&eselect All"
         Height          =   400
         Left            =   7965
         TabIndex        =   13
         Top             =   1470
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   400
         Left            =   7965
         TabIndex        =   12
         Top             =   915
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelectDeselect 
         Caption         =   "&Sel/Deselect"
         Height          =   400
         Left            =   7965
         TabIndex        =   11
         Top             =   360
         Width           =   1200
      End
      Begin VB.ListBox listHistoryScreens 
         Height          =   2760
         Left            =   150
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   345
         Width           =   7680
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8355
      TabIndex        =   1
      Top             =   6315
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7095
      TabIndex        =   0
      Top             =   6315
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   105
      Top             =   6285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
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
Private miTableType As enum_TableTypes

Private mobjDefaultFont As New StdFont
Private mlngDefaultScreenForeColor As Long

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

Private Sub cmdCancel_Click()
  blnCancelled = True
  
  UnLoad Me
End Sub

Private Sub cmdDefaultScreenFont_Click()

  On Error GoTo ErrorTrap
  
  With comDlgBox
    .FontName = mobjDefaultFont.Name
    .FontSize = mobjDefaultFont.Size
    .FontBold = mobjDefaultFont.Bold
    .FontItalic = mobjDefaultFont.Italic
    .FontUnderline = mobjDefaultFont.Underline
    .FontStrikethru = mobjDefaultFont.Strikethrough
    .Color = mlngDefaultScreenForeColor
       
    .Flags = cdlCFScreenFonts Or cdlCFEffects
    .ShowFont
      
    If mobjDefaultFont.Name <> .FontName _
      Or mobjDefaultFont.Size <> .FontSize _
      Or mobjDefaultFont.Bold <> .FontBold _
      Or mobjDefaultFont.Italic <> .FontItalic _
      Or mobjDefaultFont.Underline <> .FontUnderline _
      Or mobjDefaultFont.Strikethrough <> .FontStrikethru _
      Or txtDefaultScreenFont.ForeColor <> .Color Then
      
      mobjDefaultFont.Name = .FontName
      mobjDefaultFont.Size = .FontSize
      mobjDefaultFont.Bold = .FontBold
      mobjDefaultFont.Italic = .FontItalic
      mobjDefaultFont.Underline = .FontUnderline
      mobjDefaultFont.Strikethrough = .FontStrikethru
      mlngDefaultScreenForeColor = .Color
            
    End If
  End With

  txtDefaultFont_Refresh

ErrorTrap:
  Err = False

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
  End If
  
  imgIcon_Refresh
  
  Set frmPictSel = Nothing
  
End Sub

Private Sub cmdIconClear_Click()
  glngPictureID = frmPictSel.SelectedPicture
  imgIcon_Refresh
End Sub

Private Sub cmdMoveDown_Click()

  ' Move the selected Summary Field UP one position.
  Dim sCurrentItemText As String
  Dim iCurrentItemData As Integer
  Dim iListboxIndex As Integer
  
  iListboxIndex = listHistoryScreens.ListIndex
  
  If iListboxIndex > 0 Then
    With listHistoryScreens
      If (.ListIndex >= 0) And _
        (.ListIndex < (.ListCount - 1)) Then
      
        ' Swap the current Summary Field item with the one above it, keeping it selected.
        sCurrentItemText = .List(.ListIndex)
        iCurrentItemData = .ItemData(.ListIndex)
      
        .List(.ListIndex) = .List(.ListIndex + 1)
        .ItemData(.ListIndex) = .ItemData(.ListIndex + 1)
        
        .List(.ListIndex + 1) = sCurrentItemText
        .ItemData(.ListIndex + 1) = iCurrentItemData
    
'        .ListIndex = .ListIndex + 1
      End If
    End With
  End If

End Sub

Private Sub cmdMoveUp_Click()

  ' Move the selected Summary Field UP one position.
  Dim sCurrentItemText As String
  Dim iCurrentItemData As Integer
  Dim iListboxIndex As Integer
  
  iListboxIndex = listHistoryScreens.ListIndex
  
  If iListboxIndex > 0 Then
    With listHistoryScreens
      If .ListIndex > 0 Then
      
        ' Swap the current Summary Field item with the one above it, keeping it selected.
        sCurrentItemText = .List(.ListIndex)
        iCurrentItemData = .ItemData(.ListIndex)
      
        .List(.ListIndex) = .List(.ListIndex - 1)
        .ItemData(.ListIndex) = .ItemData(.ListIndex - 1)
        
        .List(.ListIndex - 1) = sCurrentItemText
        .ItemData(.ListIndex - 1) = iCurrentItemData
        
    '    .ListIndex = .ListIndex - 1
      End If
    End With
  End If
  
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
      
    
      If fOK And (Not fFound) Then
        ' Screen Designer not there. We must be just editing the properties of an existing screen.
        ' Check for OLE, photo, link, and image controls.
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
  Dim bGrouped As Boolean
  
  If Loading Then
 
    If ScreenID > 0 Then
      
      With recScrEdit
        
        .Index = "idxScreenID"
        .Seek "=", ScreenID
        
        If Not .NoMatch Then
          
          TableID = .Fields("tableID")
          OrderID = .Fields("orderID")
          glngPictureID = IIf(IsNull(.Fields("pictureID")), 0, .Fields("pictureID"))
          txtName.Text = Trim(.Fields("name"))
          txtDescription.Text = IIf(IsNull(.Fields("description").value), "", .Fields("description").value)
          
          GetObjectCategories cboCategory, utlScreen, ScreenID
          SetComboItem cboCategory, IIf(IsNull(.Fields("category").value), 0, .Fields("category").value)
          
          Me.Caption = "Properties - " & Trim(.Fields("name"))
          
          fQuickEntry = IIf(IsNull(.Fields("quickEntry")), False, .Fields("quickEntry"))
          chkQuickEntry.value = IIf(fQuickEntry, vbChecked, vbUnchecked)
          
          fSSIntranet = IIf(IsNull(.Fields("SSIntranet")), False, .Fields("SSIntranet"))
          chkSSIntranet.value = IIf(fSSIntranet, vbChecked, vbUnchecked)
          
          bGrouped = IIf(IsNull(.Fields("groupscreens").value), False, .Fields("groupscreens").value)
          chkGroupByCategory.value = IIf(bGrouped, vbChecked, vbUnchecked)
                    
          mobjDefaultFont.Name = .Fields("DfltFontName")
          mobjDefaultFont.Size = .Fields("DfltFontSize")
          mobjDefaultFont.Bold = .Fields("DfltFontBold")
          mobjDefaultFont.Italic = .Fields("DfltFontItalic")
          mlngDefaultScreenForeColor = .Fields("DfltForecolour")
          
        End If
      
      End With
      
    Else
      
      GetObjectCategories cboCategory, utlScreen, 0
      cboCategory.ListIndex = 0
      Me.Caption = "Properties - New Screen"
      
      Set mobjDefaultFont = gobjDefaultScreenFont
      mlngDefaultScreenForeColor = glngDefaultScreenForeColor
           
    End If
        
    ' Populate the combos, textboxes and listboxes.
    txtOrder_Refresh
    listHistoryScreens_Refresh
    imgIcon_Refresh
    txtDefaultFont_Refresh
    
    chkGroupByCategory.Enabled = (listHistoryScreens.ListCount > 0)
    
    'Refresh current tab page
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
   
End Sub

Private Function txtDefaultFont_Refresh() As Boolean

  txtDefaultScreenFont.Text = GetFontDescription(mobjDefaultFont)
  txtDefaultScreenFont.ForeColor = mlngDefaultScreenForeColor

End Function

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

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub listHistoryScreens_Click()
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

Private Sub RefreshHistoryScreensTab()
  Dim iLoop As Integer
  Dim bIndented As Boolean
  
  ' Enable the history screens listbox only if there are items.
  With listHistoryScreens
    If .ListCount > 0 Then
      .Enabled = (chkSSIntranet.value = vbUnchecked)
    
      If .ItemData(.ListIndex) Then
        
      Else
      
      End If
    
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

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtOrder_GotFocus()
  cmdOrder.SetFocus
End Sub

Private Sub listHistoryScreens_Refresh()
  Dim sSQL As String
  Dim objRecords As DAO.Recordset
  
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
    
  ' Select the top item
  If listHistoryScreens.ListCount > 0 Then
    listHistoryScreens.ListIndex = 0
  Else
    listHistoryScreens.Enabled = False
  End If
        
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
        .Fields("description").value = txtDescription.Text
        !OrderID = OrderID
        !PictureID = PictureID
        !QuickEntry = (chkQuickEntry.value = vbChecked)
        !SSIntranet = (chkSSIntranet.value = vbChecked)
        .Fields("groupscreens").value = (chkGroupByCategory.value = vbChecked)
        .Fields("category").value = GetComboItem(cboCategory)
        .Fields("DfltFontName").value = mobjDefaultFont.Name
        .Fields("DfltFontSize").value = mobjDefaultFont.Size
        .Fields("DfltFontBold").value = mobjDefaultFont.Bold
        .Fields("DfltFontItalic").value = mobjDefaultFont.Italic
        .Fields("DfltForeColour").value = txtDefaultScreenFont.ForeColor
        
        !Changed = True
        .Update
      End If
    Else
      ' Create and initialise a new screen record if required.
      Me.ScreenID = UniqueColumnValue("tmpScreens", "screenID")
      .AddNew
      !ScreenID = Me.ScreenID
      !TableID = TableID
      !OrderID = OrderID
      !Name = Trim(txtName.Text)
      .Fields("description").value = txtDescription.Text
      .Fields("category").value = GetComboItem(cboCategory)
      !PictureID = PictureID
      !QuickEntry = (chkQuickEntry.value = vbChecked)
      !SSIntranet = (chkSSIntranet.value = vbChecked)
      !GridX = 40
      !GridY = 40
      !AlignToGrid = True
    
      !dfltForeColour = vbBlack
      !dfltFontName = mobjDefaultFont.Name
      !dfltFontSize = mobjDefaultFont.Size
      !dfltFontBold = IIf(mobjDefaultFont.Bold, 1, 0)
      !dfltFontItalic = IIf(mobjDefaultFont.Italic, 1, 0)
      !dfltForeColour = txtDefaultScreenFont.ForeColor
            
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
        !ID = UniqueColumnValue("tmpHistoryScreens", "ID")
        !parentScreenID = ScreenID
        !historyScreenID = listHistoryScreens.ItemData(iIndex)
        .Fields("order").value = iIndex
        
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

Private Function imgIcon_Refresh() As Boolean
  Dim strFileName As String
  
  If PictureID > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", PictureID
      If Not .NoMatch Then
        txtIcon.Text = .Fields("Name")
      End If
    End With
  Else
    txtIcon.Text = vbNullString
  End If
End Function

