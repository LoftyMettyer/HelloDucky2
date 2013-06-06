VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelection 
   Caption         =   "Selection"
   ClientHeight    =   4950
   ClientLeft      =   660
   ClientTop       =   1785
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5032
   Icon            =   "frmSelection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000F&
      Height          =   1000
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3800
      Width           =   3315
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   3525
      Top             =   2805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   400
      Left            =   3525
      TabIndex        =   4
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Cop&y..."
      Height          =   400
      Left            =   3525
      TabIndex        =   2
      Top             =   1100
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   400
      Left            =   3525
      TabIndex        =   5
      Top             =   3400
      Width           =   1200
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "N&one"
      Height          =   400
      Left            =   3525
      TabIndex        =   6
      Top             =   3900
      Width           =   1200
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New..."
      Height          =   400
      Left            =   3525
      TabIndex        =   0
      Top             =   100
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3525
      TabIndex        =   7
      Top             =   4400
      Width           =   1200
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Edit..."
      Height          =   400
      Left            =   3525
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   400
      Left            =   3525
      TabIndex        =   3
      Top             =   1600
      Width           =   1200
   End
   Begin VB.ListBox clbItems 
      Height          =   3660
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   100
      Visible         =   0   'False
      Width           =   3315
   End
   Begin ComctlLib.ListView lstItems 
      Height          =   3645
      Left            =   120
      TabIndex        =   9
      Top             =   105
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   6429
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   327682
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Column"
         Object.Tag             =   "Column"
         Text            =   "Column"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SortKey"
         Object.Width           =   0
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar abDefSel 
      Left            =   4140
      Top             =   2520
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
      Bands           =   "frmSelection.frx":000C
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Expression selection variables.
Private mobjExpression As CExpression

' Order selection variables.
Private mobjOrder As Order

'MH20000727 Email selection variables.
Private mobjEmail As clsEmailAddr

'MH20040315 Outlook Folder
Private mobjOutlookFolder As clsOutlookFolder

Private mobjMobile As clsMobile

' Selection variables.
Private mlngAction As Long
Private mlngSelectedID As Long
  
' Form handling variables.
Private mavItemInfo() As Variant

Private Enum ObjectTypes
  OBJECT_UNKNOWN = 0
  OBJECT_EXPRESSION = 1
  OBJECT_ORDER = 2
  OBJECT_EMAIL = 3          'MH20000727
  OBJECT_OUTLOOKFOLDER = 4  'MH20040315
  OBJECT_MOBILEDESIGN = 5
End Enum
  
Private mblnReadOnly As Boolean
Private mblnForcedReadOnly As Boolean
Private mblnSelectMultiple As Boolean
Private mcolSelectedIDs As Collection


Public Property Get SelectMultiple() As Boolean
  SelectMultiple = mblnSelectMultiple
End Property

Public Property Let SelectMultiple(value As Boolean)
  mblnSelectMultiple = value
  lstItems.Visible = Not mblnSelectMultiple
  clbItems.Visible = mblnSelectMultiple
End Property


Public Property Get SelectedIDs() As Collection
  If mcolSelectedIDs Is Nothing Then
    Set mcolSelectedIDs = New Collection
  End If
  Set SelectedIDs = mcolSelectedIDs
End Property

Public Property Let SelectedIDs(value As Collection)
  Set mcolSelectedIDs = value
End Property


Public Property Get Action() As Long

  ' Return the selcted action code.
  Action = mlngAction
  
End Property



Private Sub abDefSel_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  Select Case Tool.Name
    Case "New"
      cmdNew_Click
      
    Case "EditView"
      cmdModify_Click
      
    Case "Copy"
      cmdCopy_Click
      
    Case "Delete"
      cmdDelete_Click
      
    Case "Print"
      cmdPrint_Click

    Case "Select"
      cmdSelect_Click
      
    Case "None"
      cmdDeselect_Click
      
  End Select

End Sub

Private Sub clbItems_Click()
  RefreshControls
End Sub

Private Sub clbItems_DblClick()
  RefreshControls
End Sub

Private Sub clbItems_ItemCheck(Item As Integer)
  RefreshControls
End Sub

Private Sub cmdCancel_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtCancel
  UnLoad Me
  
End Sub

Private Sub cmdCopy_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtCopy
  UnLoad Me
  
End Sub

Private Sub cmdDelete_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtDelete
  UnLoad Me

End Sub

Private Sub cmdDeselect_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtDeselect
  UnLoad Me

End Sub

Private Sub cmdModify_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtEdit
  UnLoad Me

End Sub

Private Sub cmdNew_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtAdd
  UnLoad Me
  
End Sub

Private Sub cmdPrint_Click()
  ' Print the selected object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
'  With comDlgBox
'    .Flags = cdlPDNoSelection Or cdlPDHidePrintToFile Or cdlPDReturnDC
'    .ShowPrinter
'    Printer.Copies = .Copies
'  End With

  DoEvents
  
TidyUpAndExit:
  ' Set the action property and return to the calling form.
  If fOK Then
    mlngAction = edtPrint
    UnLoad Me
  End If
  Exit Sub
  
ErrorTrap:
  ' User pressed cancel.
  fOK = False
  Resume TidyUpAndExit

End Sub


Private Sub cmdSelect_Click()
  ' Set the action property and return to the calling form.
  mlngAction = edtSelect
  UnLoad Me
  
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


  If ObjectType = OBJECT_EXPRESSION Then
    mblnReadOnly = (Application.AccessMode = accSystemReadOnly Or _
                   (Application.AccessMode = accLimited And mobjExpression.ExpressionType <> giEXPR_VIEWFILTER))
  Else
    mblnReadOnly = (Application.AccessMode = accSystemReadOnly)
  End If

  If mblnReadOnly Then
    cmdNew.Enabled = False
    cmdModify.Caption = "&View..."
    cmdCopy.Enabled = False
    cmdDelete.Enabled = False
    cmdSelect.Enabled = False
    cmdDeselect.Enabled = False
  End If

  If mblnForcedReadOnly Then
    If ObjectType = OBJECT_EXPRESSION Then
      cmdModify.Caption = "&View..."
    End If
    cmdSelect.Enabled = False
    cmdDeselect.Enabled = False
  End If
  
End Sub

Public Property Get Order() As Order
  ' Return the selection forms order object.
  Set Order = mobjOrder
  
End Property

Public Property Set Order(ByVal pobjNewValue As Order)
  ' Set the selection form's order object.
  Set mobjOrder = pobjNewValue
  Set mobjExpression = Nothing
  Set mobjEmail = Nothing
  Set mobjOutlookFolder = Nothing
  Set mobjMobile = Nothing
  
  ConfigureScreen
  lstItems_Populate
  
  RefreshControls
  
End Property


Public Property Get Email() As clsEmailAddr
  ' Return the selection forms order object.
  Set Email = mobjEmail
  
End Property

Public Property Set Email(ByVal pobjNewValue As clsEmailAddr)
  ' Set the selection form's order object.
  Set mobjEmail = pobjNewValue
  Set mobjExpression = Nothing
  Set mobjOrder = Nothing
  
  ConfigureScreen
  lstItems_Populate
  
  RefreshControls
  
End Property

Public Property Get MobileDesigner() As clsMobile
  Set MobileDesigner = mobjMobile
End Property

Public Property Set MobileDesigner(ByVal objNewValue As clsMobile)

  Set mobjExpression = Nothing
  Set mobjEmail = Nothing
  Set mobjExpression = Nothing
  Set mobjOrder = Nothing
  Set mobjMobile = objNewValue

  ConfigureScreen
  lstItems_Populate

  RefreshControls

End Property


Public Property Get OutlookFolder() As clsOutlookFolder
  Set OutlookFolder = mobjOutlookFolder
End Property

Public Property Set OutlookFolder(ByVal objNewValue As clsOutlookFolder)
  Set mobjOutlookFolder = objNewValue
  Set mobjEmail = Nothing
  Set mobjExpression = Nothing
  Set mobjOrder = Nothing
  Set mobjMobile = Nothing

  ConfigureScreen
  lstItems_Populate

  RefreshControls

End Property



Private Sub ConfigureScreen()
  ' Display certain screen controls depending on the type of
  ' selection being made.
  
  Const YSMALLGAP = 500
  Const YBIGGAP = 750
  Const YBORDERGAP = 100
  
  Select Case ObjectType
    ' The selection form is displaying the list of expressions.
    Case OBJECT_EXPRESSION
      With mobjExpression
        'JPD 20030911 Fault 6359
        Me.Caption = mobjExpression.ExpressionTypeName & "s"
   
      End With
    
    ' The selection form is displaying the list of orders.
    Case OBJECT_ORDER
      Me.Caption = "Orders"
      txtDesc.Visible = False
      
      cmdSelect.Top = cmdPrint.Top + YBIGGAP
      cmdDeselect.Top = cmdSelect.Top + YSMALLGAP
      cmdCancel.Top = cmdDeselect.Top + YSMALLGAP
      
      lstItems.Height = cmdCancel.Top + cmdCancel.Height
      Me.Height = lstItems.Top + lstItems.Height + UI.CaptionHeight + (2 * UI.YBorder) + (2 * UI.YFrame) + YBORDERGAP
      'Me.HelpContextID = 0
        
    
    'MH20000727 Added Email
    Case OBJECT_EMAIL
      Me.Caption = "Email Addresses"
      txtDesc.Visible = False
  
      cmdSelect.Top = cmdPrint.Top + YBIGGAP
      cmdDeselect.Top = cmdSelect.Top + YSMALLGAP
      cmdCancel.Top = cmdDeselect.Top + YSMALLGAP
      
      lstItems.Height = cmdCancel.Top + cmdCancel.Height
      Me.Height = lstItems.Top + lstItems.Height + UI.CaptionHeight + (2 * UI.YBorder) + (2 * UI.YFrame) + YBORDERGAP
      'Me.HelpContextID = 0
  
      clbItems.Move lstItems.Left, lstItems.Top, lstItems.Width, lstItems.Height
  
    'MH20040315 Added Outlook Folder
    Case OBJECT_OUTLOOKFOLDER
      Me.Caption = "Outlook Folder"
      txtDesc.Visible = False
  
      cmdSelect.Top = cmdPrint.Top + YBIGGAP
      cmdDeselect.Top = cmdSelect.Top + YSMALLGAP
      cmdCancel.Top = cmdDeselect.Top + YSMALLGAP
      
      lstItems.Height = cmdCancel.Top + cmdCancel.Height
      Me.Height = lstItems.Top + lstItems.Height + UI.CaptionHeight + (2 * UI.YBorder) + (2 * UI.YFrame) + YBORDERGAP
      'Me.HelpContextID = 0
    
    ' Mobile Design Groups
    Case OBJECT_MOBILEDESIGN
      Me.Caption = "Mobile Groups"
      txtDesc.Visible = True
      cmdSelect.Visible = False
      cmdDeselect.Visible = False
      cmdCancel.Caption = "OK"
  
  End Select
  
  ' Remove the icon from the caption bar
  RemoveIcon Me
  
End Sub

Private Function ObjectType() As ObjectTypes
  ' Return the current object type.
  If Not mobjExpression Is Nothing Then
    ObjectType = OBJECT_EXPRESSION
    Me.HelpContextID = 5008
    Exit Function
  End If
  
  If Not mobjOrder Is Nothing Then
    ObjectType = OBJECT_ORDER
    Me.HelpContextID = 5019
    Exit Function
  End If
  
  If Not mobjEmail Is Nothing Then
    ObjectType = OBJECT_EMAIL
    Me.HelpContextID = 5015
    Exit Function
  End If
  
  If Not mobjOutlookFolder Is Nothing Then
    ObjectType = OBJECT_OUTLOOKFOLDER
    'Me.HelpContextID = 0
    Exit Function
  End If
  
  If Not mobjMobile Is Nothing Then
    ObjectType = OBJECT_MOBILEDESIGN
    Exit Function
  End If
  
  ObjectType = OBJECT_UNKNOWN
  Me.HelpContextID = 0
  
End Function


Public Property Get Expression() As CExpression
  ' Return the selection forms expression object.
  Set Expression = mobjExpression
  
End Property

Public Property Set Expression(ByVal pobjNewValue As CExpression)
  ' Set the selection form's expression object.
  Set mobjExpression = pobjNewValue
  
  ' Release the selection form's order object, as we are not selecting an order.
  Set mobjOrder = Nothing
  Set mobjEmail = Nothing
  Set mobjOutlookFolder = Nothing
  Set mobjMobile = Nothing
  
  ' Format screen controls as appropriate.
  ConfigureScreen
  
  ' Populate the selection list.
  lstItems_Populate
  
  RefreshControls
  
End Property



Private Function lstItems_Populate() As Boolean
  ' Popluate the listbox with the appropriate items.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim lngCurrentID As Long
    
  fOK = True
  
  ' Clear the listbox.
  lstItems.ListItems.Clear
  clbItems.Clear
  
  ' Clear the array of extra item information.
  ' Column 1 = item id.
  ' Column 2 = item owner.
  ' Column 3 = Item access.
  ' Column 4 = item description.
  ReDim mavItemInfo(4, 0)
  
  Select Case ObjectType
    Case OBJECT_EXPRESSION
      fOK = lstItems_ExpressionsPopulate
      lngCurrentID = mobjExpression.ExpressionID
      
    Case OBJECT_ORDER
      fOK = lstitems_OrdersPopulate
      lngCurrentID = mobjOrder.OrderID
  
    Case OBJECT_EMAIL
      fOK = items_EmailPopulate
      lngCurrentID = mobjEmail.EmailID
    
    Case OBJECT_OUTLOOKFOLDER
      fOK = lstitems_OutlookFolderPopulate
      lngCurrentID = mobjOutlookFolder.FolderID
    
    Case OBJECT_MOBILEDESIGN
      fOK = lstitems_SecurityGroupsPopulate
      lngCurrentID = mobjMobile.MobileID
    
  End Select
  
  If fOK Then
    ' Enable the listbox if there are items.
    With lstItems
      If .ListItems.Count > 0 Then
        iIndex = 0
        
        For iLoop = 1 To .ListItems.Count
          If .ListItems(iLoop).Tag = lngCurrentID Then
            Set .SelectedItem = .ListItems(iLoop)
  
            'TM20020701 Fault 4097 - can only check the access properties if it is an expression.
            If ObjectType = OBJECT_EXPRESSION Then
              ' Show view/edit caption if necessary
              If ((mavItemInfo(3, iLoop) = ACCESS_READONLY) _
                And (Not mavItemInfo(2, iLoop) = gsUserName)) _
                Or mblnForcedReadOnly Then
                      
                cmdModify.Caption = "&View..."
              End If
            Else
              cmdModify.Caption = "&Edit..."
            End If
            
          End If
        Next iLoop
        
      'Else
      '  .Enabled = False
      End If
      .Enabled = True
      CheckListViewColWidth lstItems
    End With
  
    RefreshControls
  End If
  
TidyUpAndExit:
  lstItems_Populate = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function

Private Function lstItems_ExpressionsPopulate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsRecords As DAO.Recordset
  Dim objListItem As ListItem

  fOK = True
  lstItems.ListItems.Clear
  
  ' Define the selection string which determines
  ' what objects are displayed on the selection form.
  sSQL = "SELECT name, exprID, description, UserName, access" & _
    " FROM tmpExpressions" & _
    " WHERE TableID = " & Trim(Str(mobjExpression.BaseTableID)) & _
    " AND type=" & Trim(Str(mobjExpression.ExpressionType)) & _
    " AND (returnType =" & Trim(Str(mobjExpression.ReturnType)) & _
    "   OR type = " & Trim(Str(giEXPR_RUNTIMECALCULATION)) & ")" & _
    " AND (utilityID =" & Trim(Str(mobjExpression.UtilityID)) & _
    "   OR ((type <> " & Trim(Str(giEXPR_WORKFLOWCALCULATION)) & ")" & _
    "     AND (type <> " & Trim(Str(giEXPR_WORKFLOWSTATICFILTER)) & ")" & _
    "     AND (type <> " & Trim(Str(giEXPR_WORKFLOWRUNTIMEFILTER)) & ")))" & _
    " AND parentComponentID = 0" & _
    " AND deleted = False" & _
    " AND (Username ='" & Replace(gsUserName, "'", "''") & "'" & _
    " OR access <> '" & ACCESS_HIDDEN & "')" & _
    " ORDER BY Name"
    '& Trim(Str(giACCESS_HIDDEN)) & ")"
    
  ' Populate the listbox with the required records.
  ' Populate the array of extra item information.
  ' Column 1 = item id.
  ' Column 2 = item owner.
  ' Column 3 = Item access.
  ' Column 4 = item description.
  ' Get the required records.
  Set rsRecords = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsRecords
    Do While Not .EOF

      'lstItems.AddItem !Name
      'lstItems.ItemData(lstItems.NewIndex) = !exprID
      Set objListItem = lstItems.ListItems.Add(, , !Name)
      objListItem.Tag = !ExprID

      iNextIndex = UBound(mavItemInfo, 2) + 1
      ReDim Preserve mavItemInfo(4, iNextIndex)
      mavItemInfo(1, iNextIndex) = !ExprID
      mavItemInfo(2, iNextIndex) = IIf(IsNull(!UserName), gsUserName, !UserName)
      mavItemInfo(3, iNextIndex) = IIf(IsNull(!Access), ACCESS_READWRITE, !Access)
      mavItemInfo(4, iNextIndex) = IIf(IsNull(!Description), "", !Description)

      .MoveNext
    Loop

    .Close
  End With
  Set rsRecords = Nothing

TidyUpAndExit:
  lstItems_ExpressionsPopulate = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function

Private Function lstitems_OrdersPopulate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsRecords As DAO.Recordset
  Dim objListItem As ListItem
    
  fOK = True
  lstItems.ListItems.Clear
  
  ' Define the selection string which determines
  ' what objects are displayed on the selection form.
  sSQL = "SELECT name, orderID " & _
    " FROM tmpOrders " & _
    " WHERE tableID = " & Trim(Str(mobjOrder.TableID)) & _
    " AND type = " & Trim(Str(mobjOrder.OrderType)) & _
    " AND deleted = FALSE" & _
    " ORDER BY name"
  
  ' Populate the listbox with the required records.
  Set rsRecords = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsRecords
    Do While Not .EOF
      
      'lstItems.AddItem !Name
      'lstItems.ItemData(lstItems.NewIndex) = !OrderID
      Set objListItem = lstItems.ListItems.Add(, , !Name)
      objListItem.Tag = !OrderID
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsRecords = Nothing
  
TidyUpAndExit:
  lstitems_OrdersPopulate = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function


Private Function items_EmailPopulate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsRecords As DAO.Recordset
  Dim objListItem As ListItem

  fOK = True
  lstItems.ListItems.Clear

  ' Define the selection string which determines what objects are displayed on the selection form.
  sSQL = "SELECT name, EMailID, Deleted " & _
    " FROM tmpEmailAddresses " & _
    " WHERE (tableID = 0 OR tableID = " & CStr(mobjEmail.TableID) & ")" & _
    " AND deleted = FALSE" & _
    " ORDER BY name"
  
  ' Populate the listbox with the required records.
  Set rsRecords = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsRecords
    Do While Not .EOF
      
      If mblnSelectMultiple Then
        AddItemToListbox clbItems, !Name, !EmailID, Exists(mcolSelectedIDs, CStr(!EmailID))
      Else
        Set objListItem = lstItems.ListItems.Add(, , !Name)
        objListItem.Tag = !EmailID
      End If
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsRecords = Nothing
  
TidyUpAndExit:
  items_EmailPopulate = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function

Private Function Exists(col As Collection, key As String) As Boolean
  On Local Error GoTo LocalErr
  Dim test As Variant
  
  test = col(key)
  Exists = True
  
  Exit Function
  
LocalErr:
  Exists = False
  
End Function


Private Function lstitems_OutlookFolderPopulate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsRecords As DAO.Recordset
  Dim objListItem As ListItem

  fOK = True
  lstItems.ListItems.Clear

  ' Define the selection string which determines
  ' what objects are displayed on the selection form.
  '" WHERE tableID = " & Trim(Str(mobjOutlookFolder.TableID)) & _

  sSQL = "SELECT name, FolderID " & _
    " FROM tmpOutlookFolders " & _
    " WHERE (tableID = 0 OR tableID = " & CStr(mobjOutlookFolder.TableID) & ")" & _
    " AND deleted = FALSE" & _
    " ORDER BY name"
  
  ' Populate the listbox with the required records.
  Set rsRecords = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsRecords
    Do While Not .EOF
      
      'lstItems.AddItem !Name
      'lstItems.ItemData(lstItems.NewIndex) = !OutlookFolderID
      Set objListItem = lstItems.ListItems.Add(, , !Name)
      objListItem.Tag = !FolderID
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsRecords = Nothing
  
TidyUpAndExit:
  lstitems_OutlookFolderPopulate = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit

End Function


Private Function lstitems_SecurityGroupsPopulate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  Dim objListItem As ListItem

  fOK = True
  lstItems.ListItems.Clear

  ' Get the recordset of user groups and their access on this definition.
  sSQL = "SELECT gid, name FROM sysusers" & _
    " WHERE gid = uid AND gid > 0" & _
    "   AND not (name like 'ASRSys%') AND not (name like 'db[_]%')" & _
    " ORDER BY name"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsGroups
    Do While Not .EOF
      Set objListItem = lstItems.ListItems.Add(, , !Name)
      objListItem.Tag = !gid
      .MoveNext
    Loop
      
    .Close
  End With
 
TidyUpAndExit:
  lstitems_SecurityGroupsPopulate = fOK
  Set rsGroups = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit

End Function


Private Sub RefreshControls()
  Dim fSelectionMade As Boolean
  Dim iLoop As Integer
  Dim sDescription As String
  
  'fSelectionMade = (lstItems.SelCount = 1)
  If mblnSelectMultiple Then
    fSelectionMade = (clbItems.SelCount > 0)
  Else
    fSelectionMade = (Not (lstItems.SelectedItem Is Nothing))
  End If

  ' Enable/disable controls depending on the state of other.
  cmdNew.Enabled = Not ObjectType = OBJECT_MOBILEDESIGN
  cmdModify.Enabled = fSelectionMade
  cmdCopy.Enabled = fSelectionMade And Not mblnReadOnly And Not ObjectType = OBJECT_MOBILEDESIGN
  cmdDelete.Enabled = fSelectionMade And Not mblnReadOnly And Not ObjectType = OBJECT_MOBILEDESIGN
  cmdPrint.Enabled = fSelectionMade
  cmdSelect.Enabled = fSelectionMade _
    And (Not mblnReadOnly) _
    And (Not mblnForcedReadOnly) _
    And Not ObjectType = OBJECT_MOBILEDESIGN
  
  If ObjectType = OBJECT_EXPRESSION Then
    ' Refresh the 'description' textbox.
    sDescription = ""
    If fSelectionMade Then
      For iLoop = 1 To UBound(mavItemInfo, 2)
        'If mavItemInfo(1, iLoop) = lstItems.ItemData(lstItems.ListIndex) Then
        If mavItemInfo(1, iLoop) = lstItems.SelectedItem.Tag Then
          sDescription = mavItemInfo(4, iLoop)
          Exit For
        End If
      Next iLoop
    End If
    txtDesc.Text = sDescription
  End If

End Sub





Public Property Get SelectedID() As Long

  ' Return the selected id.
  SelectedID = mlngSelectedID
  
End Property




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim lngIndex As Long
  
  If UnloadMode <> vbFormCode Then
    mlngAction = edtCancel
  End If
  
  mlngSelectedID = -1
  
  If mlngAction > edtCancel And mlngAction <> edtAdd Then
    If mblnSelectMultiple Then
    
      If mlngAction = edtSelect Then
        Set mcolSelectedIDs = New Collection
        For lngIndex = 0 To clbItems.ListCount - 1
          If clbItems.Selected(lngIndex) Then
            mcolSelectedIDs.Add clbItems.ItemData(lngIndex), CStr(clbItems.ItemData(lngIndex))
  
            If mlngSelectedID = -1 Then
              mlngSelectedID = clbItems.ItemData(lngIndex)
            End If
  
          End If
        Next
      Else
        mlngSelectedID = clbItems.ItemData(clbItems.ListIndex)
      End If

    Else
      'If lstItems.ListIndex >= 0 Then
      '  mlngSelectedID = lstItems.ItemData(lstItems.ListIndex)
      If Not (lstItems.SelectedItem Is Nothing) Then
        mlngSelectedID = lstItems.SelectedItem.Tag
      End If
    
    End If
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub lstItems_Click()
  ' Ensure that only the required command controls are enabled.
  RefreshControls
End Sub


Private Sub lstItems_DblClick()
  ' Select the current expression.

  If cmdSelect.Enabled Then
    cmdSelect_Click
  End If

End Sub


Private Sub CheckListViewColWidth(lstvw As ListView)

  Dim objItem As ListItem
  Dim lngMax As Long
  Dim lngLen As Long

  lngMax = 0

  For Each objItem In lstvw.ListItems

    lngLen = TextWidth(objItem.Text)
    If lngMax < lngLen Then
      lngMax = lngLen
    End If

  Next objItem

  lngMax = lngMax + 60
  lstvw.ColumnHeaders(1).Width = lngMax

End Sub

Private Sub lstItems_ItemClick(ByVal Item As ComctlLib.ListItem)

  'TM20020701 Fault 4097 - can only check the access properties if it is an expression.
  If ObjectType = OBJECT_EXPRESSION Then
    If ((mavItemInfo(3, Item.Index) = ACCESS_READONLY) _
      And (Not mavItemInfo(2, Item.Index) = gsUserName)) _
      Or mblnForcedReadOnly Then
      
      cmdModify.Caption = "&View..."
    Else
      cmdModify.Caption = "&Edit..."
    End If
  Else
    cmdModify.Caption = "&Edit..."
  End If
  
End Sub

Private Sub lstItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  If Button = vbRightButton Then
  
    With Me.abDefSel.Bands("bndDefSel")

      ' Enable/disable the required tools.
      .Tools("New").Enabled = Me.cmdNew.Enabled
      .Tools("EditView").Caption = Me.cmdModify.Caption
      .Tools("EditView").Enabled = Me.cmdModify.Enabled
      .Tools("Copy").Enabled = Me.cmdCopy.Enabled
      .Tools("Delete").Enabled = Me.cmdDelete.Enabled
      .Tools("Print").Enabled = Me.cmdPrint.Enabled

      .Tools("Select").Enabled = Me.cmdSelect.Enabled
      .Tools("None").Enabled = Me.cmdDeselect.Enabled

    End With


    abDefSel.Bands("bndDefSel").TrackPopup -1, -1
    
  End If

End Sub


Public Property Let ForcedReadOnly(ByVal pfNewValue As Boolean)
  mblnForcedReadOnly = pfNewValue
  
End Property
