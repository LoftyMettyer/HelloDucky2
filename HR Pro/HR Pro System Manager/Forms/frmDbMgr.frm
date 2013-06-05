VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDbMgr 
   Caption         =   "Database Manager"
   ClientHeight    =   5340
   ClientLeft      =   390
   ClientTop       =   1830
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1013
   Icon            =   "frmDbMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   7080
   Begin ComctlLib.ListView ListView1 
      Height          =   3525
      Left            =   3360
      TabIndex        =   1
      Top             =   315
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   6218
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      OLEDragMode     =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Datatype"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Decimals"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Column type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Default Display Width"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   3150
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   0
      Width           =   150
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   5055
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9393
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   6271
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   556
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
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
      OLEDragMode     =   1
   End
   Begin ActiveBarLibraryCtl.ActiveBar abDbMgr 
      Left            =   3840
      Top             =   4275
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
      Bands           =   "frmDbMgr.frx":000C
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   4095
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":12E88
            Key             =   "IMG_PARENTTABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":136DA
            Key             =   "IMG_CHARACTER"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":13F2C
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1477E
            Key             =   "IMG_LOGIC"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":14FD0
            Key             =   "IMG_RELATION"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":15822
            Key             =   "IMG_RELATIONGROUP"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":16074
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":168C6
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":17118
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1796A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":17CBC
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1850E
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":18D60
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":195B2
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":19E04
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1A656
            Key             =   "IMG_CHILDTABLE"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1AEA8
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1B6FA
            Key             =   "IMG_UNKNOWN"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1BF4C
            Key             =   "IMG_COLUMN"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3075
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   4095
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1C79E
            Key             =   "IMG_DATABASE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1CFF0
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1D542
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1DA94
            Key             =   "IMG_PARENTTABLE"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1DFE6
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1E538
            Key             =   "IMG_CHILDTABLE"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1EA8A
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1EFDC
            Key             =   "IMG_RELATIONGROUP"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1F52E
            Key             =   "IMG_RELATION"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1FA80
            Key             =   "IMG_UNKNOWN"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":1FFD2
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":20524
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":20A76
            Key             =   "IMG_CHARACTER"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":20FC8
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":2151A
            Key             =   "IMG_BUTTON"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":21A6C
            Key             =   "IMG_LOGIC"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":21FBE
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":22510
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":22A62
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":22FB4
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDbMgr.frx":23506
            Key             =   "IMG_RADIO"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Left            =   3375
      TabIndex        =   4
      Top             =   0
      Width           =   3435
   End
End
Attribute VB_Name = "frmDbMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event Activate()
Event Deactivate()
Event UnLoad()

'Instantiate internal classes
Private Misc As New Misc

' Property variables.
Private gfUndoEnabled As Boolean

'Local variables
Public ActiveView As Object
Private gfSplitMoving As Boolean
Private gSngSplitStartX As Single
Private gLngUndoType As Long
Private gaUndoArray() As Variant
Private gfMenuActionKey As Boolean
Private gsTreeViewNodeKey As String
Private lngRecDescExprID As Long

Private fOK As Boolean

Private malngColumnDataWidths() As Long

Private Function ExpressionName(pExprID As Long) As String

  With recExprEdit
    .MoveFirst
    .Index = "idxExprID"
    .Seek ">=", pExprID
    
    If Not .NoMatch Then
      ExpressionName = !Name
    Else
      ExpressionName = "<Unknown>"
    End If
  End With
  
End Function

Private Function PackOutString(sValue As String) As String

'********************************************************************************
' PackOutString - Adds a packer character to the passed string until its length *
'                 is at the specified max length.                               *
'********************************************************************************

  Const lngMaxLength = 4
  Const sPacker = " "

  Dim sNewStr As String
  
  sNewStr = sValue
  
  Do Until Len(sNewStr) >= lngMaxLength
    sNewStr = sPacker & sNewStr
  Loop
  
  PackOutString = sNewStr
  
End Function

Public Property Get MinHeight() As Long
  MinHeight = 1800
End Property

Public Property Get MinWidth() As Long
  MinWidth = 1800
End Property

Private Sub SetColumnSizes()

  Dim iCount As Integer
  Dim objShowColumns As HRProSystemMgr.Properties
 
  Const iEXTRALENGTH = 3
  
  If Treeview1.SelectedItem Is Nothing Then Exit Sub
  
  If (Treeview1.SelectedItem.Tag = giNODE_TABLEGROUP) Then
    Set objShowColumns = gpropShowColumns_DataMgrTable
  ElseIf (Treeview1.SelectedItem.Tag = giNODE_TABLE) Then
    Set objShowColumns = gpropShowColumns_DataMgr
  Else
    Exit Sub
  End If
  
  With Me.ListView1
    For iCount = 1 To .ColumnHeaders.Count Step 1
      If (iCount = 1) Or (CBool(objShowColumns(.ColumnHeaders.Item(iCount).Index)) = True) Then
        If UBound(malngColumnDataWidths) > 0 Then
          .ColumnHeaders(iCount).Width = (IIf(malngColumnDataWidths(iCount) > Len(.ColumnHeaders(iCount).Text), _
            malngColumnDataWidths(iCount), Len(.ColumnHeaders(iCount).Text)) + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC)
        Else
          .ColumnHeaders(iCount).Width = (Len(.ColumnHeaders(iCount).Text) + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC)
        End If
      Else
        .ColumnHeaders(iCount).Width = 0
      End If
    Next iCount
  End With
  
End Sub



Private Sub abDbMgr_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)

'  '==================================================
'  ' Edit menu.
'  '==================================================
'  Select Case pTool.Name
'
'    Case "ID_New"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_Open"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_Delete"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_CopyTable"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_CopyColumn"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_Properties"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_SaveChanges"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.Name
'
'    Case "ID_Print"
'      EditMenu pTool.Name
'
'    Case "ID_CopyClipboard"
'      EditMenu pTool.Name
'
'    '==================================================
'    ' View menu.
'    '==================================================
'    Case "ID_LargeIcons"
'      ' Change the view to display large icons.
'      ChangeView lvwIcon
'
'    Case "ID_SmallIcons"
'      ' Change the view to display small icons.
'      ChangeView lvwSmallIcon
'
'    Case "ID_List"
'      ' Change the view to display a list.
'      ChangeView lvwList
'
'    Case "ID_Details"
'      ' Change the view to display details.
'      ChangeView lvwReport
'
'    Case "ID_CustomiseColumns"
'      ' Change which columns are displayed
'      EditMenu pTool.Name
'
'  End Select
  
  Select Case pTool.Name
  Case "ID_LargeIcons"
    ChangeView lvwIcon
    
  Case "ID_SmallIcons"
    ChangeView lvwSmallIcon
    
  Case "ID_List"
    ChangeView lvwList
    
  Case "ID_Details"
    ChangeView lvwReport
    
  Case Else
    EditMenu pTool.Name
    
  End Select

End Sub

Private Sub abDbMgr_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)

  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub Form_Activate()
  RaiseEvent Activate
  
  SetColumnSizes

End Sub

Private Sub Form_Deactivate()
  RaiseEvent Deactivate
  
End Sub

Private Sub Form_GotFocus()
  ' Set focus on the treeview.
  Treeview1.SetFocus
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyUp(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

  'TM20010920 Fault 2461
  If KeyCode = vbKeyF4 Then
    EditMenu "ID_Properties"
  End If
  
End Sub

Private Sub Form_Load()
  
  ' Set the initial form properties.
  ListView1.View = GetPCSetting(Me.Name, "View", lvwList)
  If ListView1.View = lvwReport Then ListView1.View = lvwSmallIcon
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  
  ' JDM - 06/12/01 - Fault 3258 - Was saving negative values to the registry
  Me.Top = IIf(Me.Top < 0, (Screen.Height - Me.Height) / 2, Me.Top)
  
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  Me.Height = GetPCSetting(Me.Name, "Height", Me.Height)
  Me.Width = GetPCSetting(Me.Name, "Width", Me.Width)
  fraSplit.Left = GetPCSetting(Me.Name, "Split", fraSplit.Left)
  
  If gbMaximizeScreens Then
    Me.WindowState = vbMaximized
  Else
    Me.WindowState = GetPCSetting(Me.Name, "State", Me.WindowState)
  End If
  
  ' Position controls.
  Label1.Left = 0
  Label1.Caption = " All Folders"
  ListView1.Top = Treeview1.Top
  fraSplit.Width = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
  
  ' Format the listview with correct headers
  ListView1_AddColumnHeaders
  SetColumnSizes
  
  ' Populate the treeview.
  InitialiseTreeView
  
  With Treeview1
    ' Populate the listview.
    PopulateListView .Nodes("TABLES")
    .Nodes("TABLES").Selected = True
    .Visible = True
  End With
  
  ' Dodgy fix for the dodgy toolbar.
  With frmSysMgr.tbMain
'    .Redraw = False
'    .Enabled = False
'    .Enabled = True
'    .Redraw = True
  End With
  
  ' Initialize global variables.
  gfUndoEnabled = False
  gfMenuActionKey = False
  gsTreeViewNodeKey = "TABLES"
  
  ChangeView ListView1.View
  
  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME

End Sub

Private Sub Form_Resize()

  Dim lngHeight As Long

  ' If the form is minimized the do nothing.
  
  ' RH 04/08/00 - BUG 51 FIX - Dont resize treeview if SysMgr has bn minimized
  If Me.WindowState <> vbMinimized And frmSysMgr.WindowState <> vbMinimized Then
  'If Me.WindowState <> vbMinimized Then
  
    If Me.WindowState <> vbMaximized Then
      ' Limit the minimum size of the form.
      If Me.Height < Me.MinHeight Then
        Me.Height = Me.MinHeight
      End If
      If Me.Width < Me.MinWidth Then
        Me.Width = Me.MinWidth
      End If
    End If
    
    ' Position the label controls on the form.
    Label1.Top = 0
    Label2.Top = Label1.Top
    Treeview1.Top = Label1.Top + Label1.Height
    ListView1.Top = Treeview1.Top
    
    ' Size the tree and list view controls.
    lngHeight = Me.ScaleHeight - (Treeview1.Top + StatusBar1.Height)
    If lngHeight < 0 Then lngHeight = 0
    Treeview1.Height = lngHeight
    ListView1.Height = lngHeight
  
    ' Position and size the split frame.
    With fraSplit
      .Width = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
      .Top = Label1.Top
      .Height = Label1.Height + Treeview1.Height
      If .Left + .Width > Me.ScaleWidth - 810 Then
        .Left = Me.ScaleWidth - (810 + .Width)
      End If
    End With
    
    ' Call the routine to size the tree and list view controls.
    SplitMove
    
    ' Refresh the form display.
    Me.Refresh
  End If
    
  frmSysMgr.RefreshMenu
  
  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  'Save form size and position to registry
  If Not Me.WindowState = vbMinimized Then
    
    ' JDM - 06/12/01 - Fault 3258 - Only save settings if in normal mode
    If Me.WindowState = vbNormal Then
      SavePCSetting Me.Name, "Top", Me.Top
      SavePCSetting Me.Name, "Left", Me.Left
      SavePCSetting Me.Name, "Height", Me.Height
      SavePCSetting Me.Name, "Width", Me.Width
    End If
    
    SavePCSetting Me.Name, "State", Me.WindowState
    SavePCSetting Me.Name, "Split", fraSplit.Left
    SavePCSetting Me.Name, "View", ListView1.View
  End If
  
  ' Disassociate object variables.
  Set Misc = Nothing
  
  ' Ensure the menu is updated.
  frmSysMgr.RefreshMenu True
  
End Sub

Private Sub fraSplit_MouseDown(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  
  ' Record the split move start position.
  gSngSplitStartX = pSngX
  
  ' Flag that the split is being moved.
  gfSplitMoving = True
  
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' If we are moving the split then move it.
  If gfSplitMoving Then
    fraSplit.Left = fraSplit.Left + (X - gSngSplitStartX)
  End If
  
End Sub

Private Sub fraSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' If the split is being moved then call the routine that resizes the
  ' tree and list views accordingly.
  If gfSplitMoving Then
    SplitMove
  End If
  
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

  'TM20010903 Fault 2180 (Sug)
  
  ' When a ColumnHeader object is clicked, the ListView control is
  ' sorted by the subitems of that column.
  ' Set the SortKey to the Index of the ColumnHeader - 1

  With ListView1
    .SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    .Sorted = True
    .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
  End With
  
End Sub

Private Sub ListView1_DblClick()
  Dim ThisNode As ComctlLib.Node
  
  ' If we have some listview items ...
  If ListView1.ListItems.Count > 0 Then
  
    ' If the select edlistview item has children ...
    If Treeview1.SelectedItem.Children > 0 Then
      
      ' Set the selected item to be the selected item in the treeview.
      Set ThisNode = Treeview1.Nodes(ListView1.SelectedItem.key)
      ThisNode.EnsureVisible
      Treeview1.SelectedItem = ThisNode
      
      ' Populate the listview with the children of the selected item.
      PopulateListView ThisNode
      
      ' Disassociate object variables.
      Set ThisNode = Nothing
    Else
      ' If the selected item does not have children then display its
      ' property page.
'ListView1.Refresh
      EditMenu "ID_Properties"
    End If
  End If
  
End Sub

Private Sub ListView1_GotFocus()
  
  Set ActiveView = ListView1
  frmSysMgr.RefreshMenu

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
  
  Set ActiveView = ListView1
  frmSysMgr.RefreshMenu

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
  
    Case vbKeyInsert
      EditMenu "ID_New"
      ListView1.SetFocus
    
    Case vbKeyReturn
      If (Not ListView1.SelectedItem Is Nothing) And _
      (ListView1_SelectedCount = 1) Then
        If ListView1.SelectedItem.Tag = giNODE_COLUMN Then
          EditMenu "ID_Properties"
        Else
          ListView1_DblClick
        End If
      End If
  
  End Select
  
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
  
  ' If we have just pressed a menu hot-key then do not process
  ' the key press as a jump the next listview item beginning
  ' with that letter.
  If gfMenuActionKey Then
    KeyAscii = 0
    gfMenuActionKey = False
  End If

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
  
  ' Refresh the status bar.
  RefreshStatusBar

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long
  
  ' Display a pop-up menu.
  If Button = vbRightButton Then
  
    frmSysMgr.RefreshMenu
    
    With frmSysMgr.tbMain
      'If .Tools("ID_New").Enabled Then
      If .Tools("ID_New").Enabled Or .Tools("ID_Properties").Enabled Then
        UI.GetMousePos lXMouse, lYMouse
'        .PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
        .Bands("ID_mnuEdit").TrackPopup -1, -1
      End If
    End With
    
  End If
      
  ' Refresh the status bar.
  RefreshStatusBar
  
  ' Refresh the menu.
 ' frmSysMgr.RefreshMenu

End Sub

Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
  
  ' Ensure the specified node is selected.
  Node.Selected = True
  
  ' Populate the listview with the children of the specified node.
  PopulateListView Node
  
  ' Refresh the menu.
  frmSysMgr.RefreshMenu
  
  ' If the specified node is the root node, expand it.
  If Node.key = "DATABASE" Then
    Node.Expanded = True
  End If
  
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
  
  ' Ensure the specified node is selected.
  Node.Selected = True
  
  ' Populate the listview with the children of the specified node.
  PopulateListView Node
  
  ' If the specified node is not the root node refresh the menu.
  If Not Node = "Database" Then
    frmSysMgr.RefreshMenu
  End If
  
End Sub

Private Sub TreeView1_GotFocus()
  Set ActiveView = Treeview1
  frmSysMgr.RefreshMenu
  
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyInsert Then
    EditMenu "ID_New"
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar
  
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)

  ' If we have just pressed menu hot-key then do not process
  ' the key press as a jump the next listview item beginning
  ' with that letter.
  If gfMenuActionKey Or (KeyAscii = vbKeyReturn) Then
    KeyAscii = 0
    gfMenuActionKey = False
  End If

End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long

  ' Pop up a menu if the right mouse button is pressed.
  If Button = vbRightButton Then
  
    ' Ensure the menu is up to date.
    frmSysMgr.RefreshMenu
    
    ' Call the activebar to display the popup menu.
    With frmSysMgr.tbMain
    
      'If .Tools("ID_New").Enabled Then
      If .Tools("ID_New").Enabled Or .Tools("ID_Properties").Enabled Then
        UI.GetMousePos lXMouse, lYMouse
'        .PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
        .Bands("ID_mnuEdit").TrackPopup -1, -1
      End If
      
    End With
  End If
  
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)

  ' Update the treeview tag.
  Treeview1.Tag = Treeview1.SelectedItem.Tag
  
  ' If we are changing node then clear any selections from the listview.
  If Node.key <> gsTreeViewNodeKey Then
    ListView1_ClearSelections
    ListView1.HideColumnHeaders = True
    ListView1.ListItems.Clear
    ListView1.Visible = False
    
  End If

  ' Populate the listview with the children of the specified node.
  PopulateListView Node
  
  ' Refresh the menu.
  frmSysMgr.RefreshMenu
  
End Sub

Private Function CheckRelations(ByVal pObjTable As HRProSystemMgr.Table) As Boolean
  ' Refresh the relations display for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objNode As ComctlLib.Node
  
  fOK = Not pObjTable Is Nothing
  
  If fOK Then
  
    ' If the given table has relations, add it to the relations treeview.
    If pObjTable.HasRelations Then
    
      If Not Misc.IsItemInCollection(Treeview1.Nodes, "R" & pObjTable.TableID) Then
        Set objNode = Treeview1.Nodes.Add("RELATION", _
          tvwChild, "R" & pObjTable.TableID, pObjTable.TableName, "IMG_RELATIONGROUP", "IMG_RELATIONGROUP")
        With objNode
          .Tag = giNODE_RELATION
          .EnsureVisible
          .Selected = True
        End With
      End If
      
    Else
      ' If the given table has no relations, remove it from the relations treeview.
      If Misc.IsItemInCollection(Treeview1.Nodes, "R" & pObjTable.TableID) Then
        Treeview1.Nodes.Remove "R" & pObjTable.TableID
      End If
    End If
    
  End If
  
TidyUpAndExit:
  Set objNode = Nothing
  CheckRelations = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub DB_SaveChanges()

  Dim frmPrompt As frmSaveChangesPrompt
  
  ' Save changes without exiting.
  Set frmPrompt = New frmSaveChangesPrompt
  frmPrompt.Buttons = vbOKCancel
  frmPrompt.Show vbModal
  If frmPrompt.Choice = vbOK Then
    Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
    Me.SetFocus
    frmSysMgr.RefreshMenu
  End If
  Set frmPrompt = Nothing

End Sub

Public Sub EditMenu(psMenuItem As String)
  Dim lngItemType As Long
  Dim frmShowColumns As HRProSystemMgr.frmShowColumns
  
  '# RH 05/04/2000 Fault Fix 127
  Select Case psMenuItem
  
    Case "ID_SaveChanges"
      DB_SaveChanges
      
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      ChangeView 0 'lvwIcon
      Exit Sub
    
    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      ChangeView 1 'lvwSmallIcon
      Exit Sub
    
    Case "ID_List"
      ' Change the view to display a list.
      ChangeView 2 'lvwList
      Exit Sub
    
    Case "ID_Details"
      ' Change the view to display details.
      ChangeView 3 'lvwReport
      Exit Sub
    
    Case "ID_CustomiseColumns"
      ' Customise column view...
      If (Treeview1.SelectedItem.Tag = giNODE_TABLEGROUP) Then
        Set frmShowColumns = New HRProSystemMgr.frmShowColumns
        frmShowColumns.PropertySet = gpropShowColumns_DataMgrTable
        frmShowColumns.Show vbModal
        SetColumnSizes
      ElseIf (Treeview1.SelectedItem.Tag = giNODE_TABLE) Then
        Set frmShowColumns = New HRProSystemMgr.frmShowColumns
        frmShowColumns.PropertySet = gpropShowColumns_DataMgr
        frmShowColumns.Show vbModal
        SetColumnSizes
      End If
      Exit Sub
   
   End Select
  '#
  
  'Get the selected item type
  If Not ActiveView Is Nothing Then
    If ActiveView.SelectedItem Is Nothing Then
      Select Case Treeview1.SelectedItem.Tag
        Case 0 ' Root node.
          lngItemType = giNODE_TABLEGROUP
        Case giNODE_TABLEGROUP
          lngItemType = giNODE_TABLE
        Case giNODE_TABLE
          lngItemType = giNODE_COLUMN
        Case giNODE_RELATIONGROUP
          lngItemType = giNODE_RELATION
      End Select
    Else
      lngItemType = ActiveView.SelectedItem.Tag
    End If
   
    ' Process the menu selection depending on what is currently selected.
    Select Case lngItemType
      ' If the currently selected item is a Table ...
      Case giNODE_TABLEGROUP, giNODE_TABLE
        EditMenuTable psMenuItem
      
      ' If the currently selected item is a Column ...
      Case giNODE_COLUMN
        EditMenuColumn psMenuItem
        
      ' If the currently selected item is the Parent of a Relation ...
      Case giNODE_RELATIONGROUP, giNODE_RELATION
        EditMenuRelation psMenuItem
      
      ' If the currently selected item is a Child of a Relation ...
      Case giNODE_RELATIONCHILD
        EditMenuRelationChild psMenuItem
    End Select
   
    ' Select all columns
    If psMenuItem = "ID_SelectAll" Then
      gfMenuActionKey = True
      ListView1_SelectAll
      ListView1.SetFocus
    End If
    
    'NHRD30072003 Fault 6255 Added the ID_SelectAll check because it was refreshing the view and wiping out selection
    ' Ensure the list view has the correct display for the current treeview selection.
    If (Not Treeview1.SelectedItem Is Nothing) And (psMenuItem <> "ID_SelectAll") Then
      PopulateListView Treeview1.SelectedItem, True
    End If

    RefreshListView

  End If
  
  
End Sub

Private Sub InitialiseTreeView()
  Dim iTableType As Integer
  Dim lngTableID As Long
  Dim sIconKey As String
  Dim objNode As ComctlLib.Node
  
  ' Clear the treeview.
  Treeview1.Nodes.Clear
  
  ' Add the root 'Database' node in the treeview.
  '
  ' DATABASE
  '
  Set objNode = Treeview1.Nodes.Add(, tvwChild, "DATABASE", "Database", "IMG_DATABASE")
  With objNode
    .Tag = 0
    .Expanded = True
    .Sorted = False
  End With

  ' Add the 'Tables' node to the treeview, as a child to the root node.
  '
  ' Database
  '     |
  '     +--TABLES
  '
  Set objNode = Treeview1.Nodes.Add("DATABASE", tvwChild, "TABLES", "Tables", "IMG_TABLE")
  With objNode
    .Tag = giNODE_TABLEGROUP
    .Expanded = True
    .Sorted = True
  End With
    
  ' Loop through the tables in the database and add each one to the treeview
  ' as a child to the 'Tables' node.
  '
  ' Database
  '     |
  '     +--Tables
  '          |
  '          +--<TABLE NODES>
  '
  With recTabEdit
  
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      
      ' If the table is not marked as deleted then add a new node to the treeview.
      If Not .Fields("deleted") Then
        
        iTableType = recTabEdit!TableType
        Select Case iTableType
          Case iTabParent
            sIconKey = "IMG_PARENTTABLE"
          Case iTabChild
            sIconKey = "IMG_CHILDTABLE"
          Case iTabLookup
            sIconKey = "IMG_LOOKUP"
          Case Else
            sIconKey = "IMG_UNKNOWN"
        End Select

        Set objNode = Treeview1.Nodes.Add("TABLES", tvwChild, _
          "T" & .Fields("tableID"), .Fields("tableName"), sIconKey, sIconKey)
        objNode.Tag = giNODE_TABLE
        objNode.Sorted = True
      End If
      
      .MoveNext
    Loop
  End With
    
  ' Add the 'Relationships' node to the treeview, as a child to the root node.
  '
  ' Database
  '     |
  '     +--Tables
  '     |    |
  '     |    +--<table nodes>
  '     |
  '     +--RELATIONSHIPS
  '
  Set objNode = Treeview1.Nodes.Add("DATABASE", tvwChild, "RELATION", "Relationships", "IMG_RELATIONGROUP")
  With objNode
    .Tag = giNODE_RELATIONGROUP
    .Expanded = True
    .Sorted = True
  End With
    
  ' Loop through the relations in the database and add each one to the treeview
  ' as a child to the 'Relations' node.
  '
  ' Database
  '     |
  '     +--Tables
  '     |    |
  '     |    +--<table nodes>
  '     |
  '     +--Relationships
  '          |
  '          +--<RELATION NODES>
  '
  lngTableID = 0
  
  With recRelEdit
  
    .Index = "idxParentID"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If lngTableID <> !parentID Then
      
        lngTableID = !parentID
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngTableID
        
        If Not recTabEdit.NoMatch Then
        
          Set objNode = Treeview1.Nodes.Add("RELATION", tvwChild, _
            "R" & lngTableID, recTabEdit!TableName, "IMG_RELATIONGROUP")
          objNode.Tag = giNODE_RELATION
        End If
        
      End If
      .MoveNext
    Loop
  End With
    
  ' Disassociate object variables.
  Set objNode = Nothing
  
End Sub
Private Sub ListView1_SelectAll()

  Dim iLoop As Integer
  
  ' Loop through the list view items marking each one as selected.
  For iLoop = 1 To ListView1.ListItems.Count
    ListView1.ListItems(iLoop).Selected = True
  Next iLoop

End Sub

Function IsKeyInListView(ByVal key As Variant) As Boolean
  On Error GoTo ErrorTrap
  
  If ListView1.ListItems(key).key = key Then
    IsKeyInListView = True
  End If
  
  Exit Function
  
ErrorTrap:
  IsKeyInListView = False
  If Err.Number <> 35601 Then
    MsgBox "Runtime error '" & Trim(Str(Err.Number)) & "'." & _
      vbCr & vbCr & Err.Description, _
      vbExclamation + vbOKOnly, "Microsoft Visual Basic"
  End If
  Err = False
  
End Function

Private Sub ListView1_AddColumnHeaders()
'********************************************************************************
' ListView1_AddColumnHeaders -  Adds all the required columns to the            *
'                               ColumnHeaders collection. (TM)                  *
'********************************************************************************
  Dim iCount As Integer
  
  With ListView1.ColumnHeaders
    .Clear
    
    .Add , "Name", "Name", , lvwColumnLeft
    .Add , "DataType", "Data Type", , lvwColumnLeft
    .Add , "Size", "Size", , lvwColumnRight
    .Add , "Decimals", "Decimals", , lvwColumnRight
    .Add , "Display", "Display", , lvwColumnRight
    .Add , "ColumnType", "Column Type", , lvwColumnLeft
    .Add , "ControlType", "Control Type", , lvwColumnLeft
    .Add , "ReadOnly", "Read Only", , lvwColumnCenter
    .Add , "Audit", "Audit", , lvwColumnCenter
    .Add , "Multi-line", "Multi-line", , lvwColumnCenter
    .Add , "Case", "Case", , lvwColumnLeft
    .Add , "DefaultValue", "Default Value", , lvwColumnLeft
    .Add , "TextAlignment", "Text Alignment", , lvwColumnLeft
    .Add , "DuplicateCheck", "Duplicate Check", , lvwColumnCenter
    .Add , "Mandatory", "Mandatory", , lvwColumnCenter
    .Add , "UniqueTables", "Unique in Table", , lvwColumnCenter
    .Add , "UniqueSiblings", "Unique in Siblings", , lvwColumnCenter
    .Add , "Mask", "Mask", , lvwColumnCenter
    .Add , "Custom", "Custom Validation", , lvwColumnCenter
    .Add , "DiaryLinks", "Diary Links", , lvwColumnRight
    '.Add , "EmailLinks", "Email Links", , lvwColumnRight
        
    'NHRD23072003 Fault 6207
    If gbAFDEnabled Then
      .Add , "AfdPostcode", "Afd Postcode", , lvwColumnCenter
    End If
    
    If gbQAddressEnabled Then
      .Add , "Quick Address", "Quick Address", , lvwColumnCenter
    End If
    
    
    'NHRD29072003 Fault 6208 Added the ability to show and store Use1000Separator and
    'Trimming column properties
    .Add , "Use1000Separator", "Use 1000 Separator", , lvwColumnCenter
    .Add , "Trimming", "Trimming", , lvwColumnCenter
      
    ReDim malngColumnDataWidths(.Count)
    For iCount = 1 To .Count
      malngColumnDataWidths(iCount) = 0
    Next iCount
  End With
  
End Sub
Private Function AddDetailsToListItem(pListView As ComctlLib.ListView) As Boolean

'********************************************************************************
' AddDetailsToListItem -  Adds all the details for the current passed item in   *
'                         list.                                                 *
'                         Need function required to keep PopulateListView       *
'                         function concise.                                     *
'                         Also made it easier to change the order of the column *
'                         details buy referencing the SubItems by the 'Key'     *
'                         as opposed to the index. (Name always is the first)   *
'                                                                               *
'                         ListView1_AddColumnHeaders() must be called before    *
'                         this function can populate the SubItems.              *
'********************************************************************************

  Dim objItem As ComctlLib.ListItem
  Dim iOrder As Integer
  Dim iTemp As Integer
  
  Dim iDiaryLinks As Integer
  'Dim iEmailLinks As Integer
  Dim objDiaryLink As cDiaryLink
  'Dim objEmailLink As clsEmailLink
  Dim sDefault As String
  Dim objMisc As Misc
  Dim iCount As Integer
  
  iCount = 0
  
  With recColEdit
    '******************************************************************************
    '   Name                                                                      *
    '******************************************************************************
    Set objItem = pListView.ListItems.Add(, "C" & !ColumnID, !ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objItem.Tag = giNODE_COLUMN
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.Text), Len(objItem.Text), malngColumnDataWidths(iCount))
    iOrder = 1
    
    '******************************************************************************
    '   Datatype
    '******************************************************************************
    objItem.SubItems(iOrder) = Database.GetDataDesc(!DataType)
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    '******************************************************************************
    '   Put in correct Icon for datatype
    '     NHRD16112006Fault11560 Added images in ImageLists 1&2 too
    '     There's a little bit more further down too - search on 'Link'
    '******************************************************************************
    Select Case !DataType
      Case 0       ' ?
        objItem.SmallIcon = "IMG_UNKNOWN"
        objItem.Icon = "IMG_UNKNOWN"
      Case -4          ' OLE columns
        objItem.SmallIcon = "IMG_OLE"
        objItem.Icon = "IMG_OLE"
      Case -7      ' Logic columns
        objItem.SmallIcon = "IMG_LOGIC"
        objItem.Icon = "IMG_LOGIC"
      Case 2      ' Numeric columns
        objItem.SmallIcon = "IMG_NUMERIC"
        objItem.Icon = "IMG_NUMERIC"
      Case 4       ' Integer columns
        objItem.SmallIcon = "IMG_NUMERIC"
        objItem.Icon = "IMG_NUMERIC"
      Case 11         ' Date columns
        'NHRD28112006 Fault 11560 Use character bitmap as date bitmap for now
        'objItem.SmallIcon = "IMG_DATE"
        'objItem.Icon = "IMG_DATE"
        objItem.SmallIcon = "IMG_CHARACTER"
        objItem.Icon = "IMG_CHARACTER"
      Case 12      ' Character columns
        objItem.SmallIcon = "IMG_CHARACTER"
        objItem.Icon = "IMG_CHARACTER"
      Case -3    ' Photo columns
        objItem.SmallIcon = "IMG_PHOTO"
        objItem.Icon = "IMG_PHOTO"
      Case -1  ' Working Pattern columns
        objItem.SmallIcon = "IMG_WORKINGPATTERN"
        objItem.Icon = "IMG_WORKINGPATTERN"
    End Select
    
    '******************************************************************************
    '   Size & Decimals
    '******************************************************************************
    If Database.ColumnHasSize(!DataType) Then
    
      If !MultiLine Then
        objItem.SubItems(iOrder) = vbNullString
      Else
        objItem.SubItems(iOrder) = PackOutString(Trim(Str(!Size)))
      End If
      
      iCount = iCount + 1
      malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
      iOrder = iOrder + 1

      If Database.ColumnHasScale(!DataType) Then
        objItem.SubItems(iOrder) = PackOutString(Trim(Str(!Decimals)))
        iCount = iCount + 1
        malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
        iOrder = iOrder + 1
      Else
        objItem.SubItems(iOrder) = ""
        iCount = iCount + 1
        malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
        iOrder = iOrder + 1
      End If
    Else
      objItem.SubItems(iOrder) = ""
      iCount = iCount + 1
      malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
      iOrder = iOrder + 1
      objItem.SubItems(iOrder) = ""
      iCount = iCount + 1
      malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
      iOrder = iOrder + 1
    End If
    
    ' NHRD 29/07/03 not sure what was goiing on here but the additional iOrder + 1 below
    ' sorted out the misaligned data
    'iOrder = iOrder + 1
    
    '******************************************************************************
    '   Default Display Width
    '******************************************************************************
    'TM20011030 Fault 3038
    'Don't show the Default Display Width for Link fields, OLE Columns or Photo Columns.
    If !ColumnType = giCOLUMNTYPE_LINK Or !DataType = sqlOle Or !DataType = sqlVarBinary Then
      objItem.SubItems(iOrder) = ""
    Else
      If !MultiLine Then
        objItem.SubItems(iOrder) = vbNullString
      Else
        objItem.SubItems(iOrder) = PackOutString(Trim(Str(!DefaultDisplayWidth)))
      End If
    End If
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   ColumnType
    '******************************************************************************
    Select Case !ColumnType
      Case giCOLUMNTYPE_SYSTEM
        objItem.SubItems(iOrder) = "System"
      Case giCOLUMNTYPE_DATA
        objItem.SubItems(iOrder) = "Data"
      Case giCOLUMNTYPE_LOOKUP
        objItem.SubItems(iOrder) = "Lookup"
      Case giCOLUMNTYPE_CALCULATED
        objItem.SubItems(iOrder) = "Calculated"
      Case giCOLUMNTYPE_LINK
        objItem.SubItems(iOrder) = "Link"
        'NHRD16112006Fault11560 Added this as when a link column type is
        'displayed it has no dedicated icon
        objItem.SmallIcon = "IMG_LINK"
        objItem.Icon = "IMG_LINK"
    End Select
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Control Type
    '******************************************************************************
    Select Case !ControlType
    Case giCTRL_CHECKBOX
      objItem.SubItems(iOrder) = "Check Box"
        'NHRD16112006Fault11560 Added this as when a Option column type is
        'displayed it has no dedicated icon
        objItem.SmallIcon = "IMG_LOGIC"
        objItem.Icon = "IMG_LOGIC"
    Case giCTRL_COMBOBOX
      objItem.SubItems(iOrder) = "Dropdown List"
    Case giCTRL_OPTIONGROUP
      objItem.SubItems(iOrder) = "Option Group"
        'NHRD16112006Fault11560 Added this as when a Option column type is
        'displayed it has no dedicated icon
        objItem.SmallIcon = "IMG_RADIO"
        objItem.Icon = "IMG_RADIO"
    Case giCTRL_SPINNER
      objItem.SubItems(iOrder) = "Spinner"
    Case giCTRL_TEXTBOX
      objItem.SubItems(iOrder) = "Text Box"
    Case giCTRL_WORKINGPATTERN
      objItem.SubItems(iOrder) = "Working Pattern"
    Case Else
      objItem.SubItems(iOrder) = ""
    End Select
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Read Only
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!ReadOnly, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Audit
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!audit, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Multiline
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!MultiLine, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Text Case
    '******************************************************************************
    Select Case !convertcase
    Case vbUpperCase
      objItem.SubItems(iOrder) = "Upper"
    Case vbProperCase
      objItem.SubItems(iOrder) = "Proper"
    Case vbLowerCase
      objItem.SubItems(iOrder) = "Lower"
    Case Else
      objItem.SubItems(iOrder) = ""
    End Select
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Default Value
    '******************************************************************************
    If Not IsNull(!dfltValueExprID) And !dfltValueExprID <> 0 Then
      objItem.SubItems(iOrder) = "<calc> " & ExpressionName(!dfltValueExprID)
    Else
      'JPD 20041210 Fault 9620
      'objItem.SubItems(iOrder) = !DefaultValue
      If !DataType = dtTIMESTAMP Then
        sDefault = Trim(!DefaultValue)
        If Len(sDefault) = 8 Then
          ' Previous version of HR Pro saved the defult dates in the format mmddyyyy.
          ' If the default is in this format, convert to mm/dd/yyyy format.
          sDefault = Left(sDefault, 2) & "/" & Mid(sDefault, 3, 2) & "/" & Mid(sDefault, 5)
        End If
    
        Set objMisc = New Misc
        sDefault = IIf(Len(sDefault) > 0, objMisc.ConvertSQLDateToLocale(sDefault), "")
        Set objMisc = Nothing
  
        objItem.SubItems(iOrder) = sDefault
      Else
        objItem.SubItems(iOrder) = !DefaultValue
      End If
    End If
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Text Alignment
    '******************************************************************************
    Select Case !Alignment
    Case vbLeftJustify
      objItem.SubItems(iOrder) = "Left"
    Case vbRightJustify
      objItem.SubItems(iOrder) = "Right"
    Case vbCenter
      objItem.SubItems(iOrder) = "Center"
    Case Else
      objItem.SubItems(iOrder) = "Left"
    End Select
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Duplicate Check
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!Duplicate, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Mandatory
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!Mandatory, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Unique check within Table
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!uniqueCheckType = giUNIQUECHECKTYPE_ENTIRE, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Unique check within Child table
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(!uniqueCheckType = giUNIQUECHECKTYPE_SIBLINGSALL, "Y", "N")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Mask Validation
    '******************************************************************************
    objItem.SubItems(iOrder) = IIf(IsNull(!Mask), "N", "Y")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Custom Validation
    '******************************************************************************
    If Not IsNull(!lostFocusExprID) And !lostFocusExprID <> 0 Then
      objItem.SubItems(iOrder) = ExpressionName(!lostFocusExprID)
    Else
      objItem.SubItems(iOrder) = ""
    End If
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1

    '******************************************************************************
    '   Diary Links
    '******************************************************************************
    iDiaryLinks = 0
    With recDiaryEdit
      .Index = "idxColumnID"
      .Seek "=", recColEdit!ColumnID
      
      If Not .NoMatch Then
        Do While Not .EOF
        
          If !ColumnID <> recColEdit!ColumnID Then
            Exit Do
          End If
          
          Set objDiaryLink = New cDiaryLink
          objDiaryLink.DiaryLinkId = !diaryID
          If objDiaryLink.ReadDiaryLink Then
            iDiaryLinks = iDiaryLinks + 1
          End If
          Set objDiaryLink = Nothing
          
          .MoveNext
        Loop
      End If
    End With

    'Diary Links
    objItem.SubItems(iOrder) = IIf(iDiaryLinks > 0, Str(iDiaryLinks), "")
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
  
'    '******************************************************************************
'    '   Email Links
'    '******************************************************************************
'    iEmailLinks = 0
'    With recEmailLinksEdit
'      .Index = "idxColumnID"
'      .Seek "=", recColEdit!ColumnID
'
'      If Not .NoMatch Then
'        Do While Not .EOF
'          If !ColumnID <> recColEdit!ColumnID Then
'            Exit Do
'          End If
'          Set objEmailLink = New clsEmailLink
'          objEmailLink.LinkID = !LinkID
'          If objEmailLink.ReadEmailLink Then
'            iEmailLinks = iEmailLinks + 1
'          End If
'          Set objEmailLink = Nothing
'
'          .MoveNext
'        Loop
'      End If
'    End With
'
'    'Email Links
'    objItem.SubItems(iOrder) = IIf(iEmailLinks > 0, Str(iEmailLinks), "")
'    iCount = iCount + 1
'    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
'    iOrder = iOrder + 1

    '******************************************************************************
    '   Afd Postcode
    '******************************************************************************
    If gbAFDEnabled Then
      objItem.SubItems(iOrder) = IIf(!afdEnabled, "Y", "N")
      iCount = iCount + 1
      malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
      iOrder = iOrder + 1
    End If
    
    '******************************************************************************
    '   Quick Address Postcode
    '******************************************************************************
    If gbQAddressEnabled Then
      objItem.SubItems(iOrder) = IIf(!qaddressEnabled, "Y", "N")
      iCount = iCount + 1
      malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
      iOrder = iOrder + 1
    End If
    
    'NHRD29072003 Fault 6208 Added the ability to show and store Use1000Separator and
    ' Trimming column properties
    '******************************************************************************
    '   Use 1000 Separator
    '******************************************************************************
    If Not IsNull(!Use1000Separator) Then
      objItem.SubItems(iOrder) = IIf(!Use1000Separator, "Y", "N")
    Else
      objItem.SubItems(iOrder) = "N"
    End If
    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
    '******************************************************************************
    '   Trimming
    '******************************************************************************
    If Not IsNull(!Trimming) Then
      Select Case !Trimming
        Case giTRIM_NONE
          objItem.SubItems(iOrder) = "None"
        Case giTRIM_BOTHSIDES
          objItem.SubItems(iOrder) = "Left & Right"
        Case giTRIM_LEFTSIDE
          objItem.SubItems(iOrder) = "Left Only"
        Case giTRIM_RIGHTSIDE
          objItem.SubItems(iOrder) = "Right Only"
        Case Else
          objItem.SubItems(iOrder) = "None"
      End Select
    Else
      objItem.SubItems(iOrder) = "None"
    End If

    iCount = iCount + 1
    malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
    iOrder = iOrder + 1
    
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objItem = Nothing
  
  Exit Function
    
ErrorTrap:
  MsgBox "Error adding list detials to list.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit

End Function


Private Sub ListView1_AddTableColumnHeaders()
'********************************************************************************
' ListView1_AddColumnHeaders -  Adds all the required columns to the            *
'                               ColumnHeaders collection. (TM)                  *
'********************************************************************************
  Dim iCount As Integer
  
  With ListView1.ColumnHeaders
    .Clear
    
    .Add , "Name", "Name", , lvwColumnLeft
    .Add , "Type", "Type", , lvwColumnLeft
    .Add , "PrimaryOrder", "Primary Order", , lvwColumnLeft
    .Add , "RecordDescription", "Record Description", , lvwColumnLeft
    .Add , "DefaultEmail", "Default Email", , lvwColumnLeft
    .Add , "EmailLinks", "Email Links", , lvwColumnRight
    .Add , "CalendarLinks", "Calendar Links", , lvwColumnRight
    
    If IsModuleEnabled(modWorkflow) Then
      .Add , "WorkflowLinks", "Workflow Links", , lvwColumnRight
    End If
    
    ReDim malngColumnDataWidths(.Count)
    For iCount = 1 To .Count
      malngColumnDataWidths(iCount) = 0
    Next iCount
  End With
  
End Sub


Private Sub PopulateListView(pobjNode As ComctlLib.Node, Optional ByVal pfRefresh As Boolean)
  Dim iLoop As Integer
  Dim iNodeIndex As Integer
  Dim iArrayIndex As Integer
  Dim iSelectionCount As Integer
  Dim lngNodeID As Long
  Dim objItem As ComctlLib.ListItem
  Dim objChildNode As ComctlLib.Node
  Dim asSelectedKeys() As String
  Dim lngCount As Long
  Dim iCount As Integer
  Dim iEmailLinks As Integer
  Dim objEmailLink As clsEmailLink
    
  ' Display a description of what is being shown in the listview.
  If Not pobjNode Is pobjNode.Root Then
    Label2.Caption = " " & pobjNode.Parent.Text & " : " & pobjNode.Text
  Else
    Label2.Caption = " " & pobjNode.Text
  End If
  
  ' Update the global variable that records which node is selected in the treeview.
  gsTreeViewNodeKey = pobjNode.key
  
  ' If we are not already displaying the children of the specified node,
  ' or we are forcing a refresh ...
  If pfRefresh Or (ListView1.Tag <> pobjNode.key) Then
    Screen.MousePointer = vbHourglass
    
    ' Remember which listview items are currently selected.
    iSelectionCount = ListView1_SelectedCount
    If iSelectionCount > 0 Then
      ReDim asSelectedKeys(iSelectionCount)
      iArrayIndex = 1
      For iLoop = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(iLoop).Selected Then
          asSelectedKeys(iArrayIndex) = ListView1.ListItems(iLoop).key
          iArrayIndex = iArrayIndex + 1
        End If
      Next iLoop
    End If

    abDbMgr.Tools("ID_CustomiseColumns").Enabled = abDbMgr.Tools("ID_Details").Checked
    frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = abDbMgr.Tools("ID_Details").Checked

    ' Update the listview tag.
    ListView1.Tag = pobjNode.key
    
    ' Clear all items from the listview.
    ListView1.ListItems.Clear

    ReDim malngColumnDataWidths(0)
    
    ' Ensure that we display the column headings when the listview displays details.
    'ListView1.HideColumnHeaders = False
    
    'TM20010917 Fault 2040
    'Hide the list view while being populated, therefore avoiding the messy initial appearance
    'of the listview when in small icon view.
    ListView1.Visible = False
        
    ' If we are populating the listview with Columns (ie. the specified node is
    ' a table) ...
    If pobjNode.Tag = giNODE_TABLE Then
      lngNodeID = Val(Mid(pobjNode.key, 2))
      
      ListView1.HideColumnHeaders = True
      ListView1_AddColumnHeaders
      ListView1.Visible = False
      
      With recColEdit
        .Index = "idxName"
        .Seek ">=", lngNodeID
        
        If Not .NoMatch Then
          Do While Not .EOF
            ' Ignore any columns for tables other than the one used by the specified
            ' node.
            If .Fields("tableID") <> lngNodeID Then
              Exit Do
            End If
              
            ' Ignore deleted and system columns.
            If (Not .Fields("deleted")) And _
              (Not !ColumnType = giCOLUMNTYPE_SYSTEM) Then
    
              ' Add an item to the listview for the column.
              AddDetailsToListItem Me.ListView1
            End If
              
            .MoveNext
          Loop
        End If
      End With
      SetColumnSizes
      ListView1.HideColumnHeaders = False
    
    ' If we are populating the listview with Relations ...
    ElseIf pobjNode.Tag = giNODE_RELATION Then
    
      'Remove the column headers
      ListView1.HideColumnHeaders = True
      
      lngNodeID = Val(Mid(pobjNode.key, 2))
      
      With recRelEdit
        .Index = "idxParentID"
        .Seek ">=", lngNodeID
        
        If Not .NoMatch Then
          Do While Not .EOF
            ' Ignore any relationships that do not concern the given node.
            If !parentID <> lngNodeID Then
              Exit Do
            End If
          
            recTabEdit.Index = "idxTableID"
            recTabEdit.Seek "=", !childID
            
            If Not recTabEdit.NoMatch Then
              ' Add the relation to the listview.
              Set objItem = ListView1.ListItems.Add(, _
                "R" & !childID, recTabEdit!TableName, "IMG_RELATION", "IMG_RELATION")
              objItem.Tag = giNODE_RELATIONCHILD
              
              ' Disassociate object variables.
              Set objItem = Nothing
            End If
          
            .MoveNext
          Loop
        End If
      End With
      
    ' If we are populating the listview with nodes that have children themselves.
    ' ie. if we are listing tables, relation groups, etc.
    ElseIf pobjNode.Children > 0 Then
      ListView1.Sorted = (pobjNode.Tag > 0)
      
      If pobjNode.Tag = giNODE_TABLEGROUP Then
        abDbMgr.Tools("ID_CustomiseColumns").Enabled = False
        frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = False
        ListView1_AddTableColumnHeaders
        ChangeView lvwReport
        fOK = True
      Else
        fOK = False

        ReDim malngColumnDataWidths(1)
        malngColumnDataWidths(1) = 0
      End If
      
      ' Determine the index of the specified node's first child.
      iNodeIndex = pobjNode.Child.FirstSibling.Index
      
      'Copies all the table names from the Treeview to the listview
      Do While iNodeIndex >= 0
        iCount = 0
        
        ' Add items to the listview for each of the specified node's children.
        Set objChildNode = Treeview1.Nodes(iNodeIndex)
        Set objItem = ListView1.ListItems.Add(, objChildNode.key, objChildNode.Text, _
          objChildNode.Image, objChildNode.Image)
          iCount = iCount + 1
          malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objChildNode.Text), Len(objChildNode.Text), malngColumnDataWidths(iCount))
            
        lngNodeID = Val(Mid(objChildNode.key, 2))
        If fOK Then
          Dim iOrder As Integer
          With recTabEdit
            recTabEdit.Index = "idxTableID"
            recTabEdit.Seek "=", lngNodeID
                
            If Not recTabEdit.NoMatch Then
              iOrder = 1
              '************************************
              '   Type                            *
              '************************************
              
              Dim strDisplayString As String
              
              Select Case !TableType
                Case iTabParent
                  strDisplayString = "Parent"
                Case iTabChild
                  strDisplayString = "Child"
                Case iTabLookup
                  strDisplayString = "Lookup"
              End Select

              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))
              
              objItem.SubItems(iOrder) = strDisplayString
              iOrder = iOrder + 1
              '************************************
              '   Primary Order                   *
              '************************************
              'strDisplayString = Val(!defaultorderid)
              With recOrdEdit
                .Index = "idxID"
                .Seek "=", recTabEdit!defaultOrderID
        
                If .NoMatch Then
                  strDisplayString = ""
                Else
                  If !Deleted Then
                    strDisplayString = ""
                  Else
                    strDisplayString = !Name
                  End If
                End If
              End With
              
              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))

              objItem.SubItems(iOrder) = strDisplayString
              iOrder = iOrder + 1
              '************************************
              '   Record Description              *
              '************************************
              lngRecDescExprID = CLng(recTabEdit!RecordDescExprID)
              strDisplayString = GetRecordDescriptionDetails(lngRecDescExprID)
    
              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))

              objItem.SubItems(iOrder) = strDisplayString
              iOrder = iOrder + 1
              '************************************
              '   Default Email ID                *
              '************************************
              With recEmailAddrEdit
                .Index = "idxID"
                .Seek "=", recTabEdit!DefaultEmailID
              
                If .NoMatch Then
                  strDisplayString = ""
                Else
                  If !Deleted Then
                    strDisplayString = ""
                  Else
                  strDisplayString = !Name
                  End If
                End If
              End With
              
              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))

              objItem.SubItems(iOrder) = strDisplayString
              iOrder = iOrder + 1
            
            
              '******************************************************************************
              '   Email Links
              '******************************************************************************
              iEmailLinks = 0
              With recEmailLinksEdit
                .Index = "idxTableID"
                .Seek "=", recTabEdit!TableID

                If Not .NoMatch Then
                  Do While Not .EOF
                    If !TableID <> recTabEdit!TableID Then
                      Exit Do
                    End If
                    If Not !Deleted Then
                      Set objEmailLink = New clsEmailLink
                      objEmailLink.LinkID = !LinkID
                      If objEmailLink.ReadEmailLink Then
                        iEmailLinks = iEmailLinks + 1
                      End If
                      Set objEmailLink = Nothing
                    End If

                    .MoveNext
                  Loop
                End If
              End With

              'Email Links
              objItem.SubItems(iOrder) = IIf(iEmailLinks > 0, CStr(iEmailLinks), "")
              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(objItem.SubItems(iOrder)), Len(objItem.SubItems(iOrder)), malngColumnDataWidths(iCount))
              iOrder = iOrder + 1


              '************************************
              '   Outlook Calendar Links          *
              '************************************
              With recOutlookLinks
                .Index = "idxTableID"
                .Seek "=", lngNodeID
                
                lngCount = 0
                If Not .NoMatch Then
                  Do While !TableID = lngNodeID
                    If Not !Deleted Then
                      lngCount = lngCount + 1
                    End If
                    .MoveNext
                    If .EOF Then
                      Exit Do
                    End If
                  Loop
                End If
              End With
              strDisplayString = IIf(lngCount > 0, CStr(lngCount), vbNullString)

              iCount = iCount + 1
              malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))

              objItem.SubItems(iOrder) = strDisplayString
              iOrder = iOrder + 1
            
              '************************************
              '   Workflow Links                  *
              '************************************
              If IsModuleEnabled(modWorkflow) Then
                With recWorkflowTriggeredLinks
                .Index = "idxTableID"
                .Seek "=", lngNodeID

                lngCount = 0
                If Not .NoMatch Then
                  Do While !TableID = lngNodeID
                    If Not !Deleted Then
                      lngCount = lngCount + 1
                    End If
                    .MoveNext
                    If .EOF Then
                      Exit Do
                    End If
                  Loop
                End If
                End With
                strDisplayString = IIf(lngCount > 0, CStr(lngCount), vbNullString)
  
                iCount = iCount + 1
                malngColumnDataWidths(iCount) = IIf(malngColumnDataWidths(iCount) < Len(strDisplayString), Len(strDisplayString), malngColumnDataWidths(iCount))
  
                objItem.SubItems(iOrder) = strDisplayString
                iOrder = iOrder + 1
              End If
              
            End If
          End With
        End If
        
        objItem.Tag = objChildNode.Tag
        
        ' Disassociate object variables.
        Set objItem = Nothing
        
        ' Stop looping if we have reached the specified node's last child.
        If iNodeIndex = pobjNode.Child.LastSibling.Index Then
          ' Disassociate object variables.
          Set objChildNode = Nothing
          Exit Do
        Else
          iNodeIndex = objChildNode.Next.Index
        End If
        
        ' Disassociate object variables.
        Set objChildNode = Nothing
      Loop
      
      If fOK Then
        SetColumnSizes
        ListView1.HideColumnHeaders = False
      End If
    End If
    
    'JPD 20051006 Fault 10412
    ListView1.Refresh
  End If
 
  ' Reselect any items that were selected before we populated the listview.
  If iSelectionCount > 0 Then
    For iArrayIndex = 1 To UBound(asSelectedKeys)
      ' Only try to select an item if it still appears in the listview.
      If IsKeyInListView(asSelectedKeys(iArrayIndex)) Then
        ListView1.ListItems(asSelectedKeys(iArrayIndex)).Selected = True
      End If
    Next iArrayIndex
    
    ' Ensure the selected item is visible.
    If Not ListView1.SelectedItem Is Nothing Then
      ListView1.SelectedItem.EnsureVisible
    End If
  End If

  ' If no items are selected then try to select the first one.
  If (ListView1_SelectedCount = 0) And (ListView1.ListItems.Count > 0) Then
    ListView1.SelectedItem = ListView1.ListItems(1)
    ListView1.SelectedItem.EnsureVisible
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar

  ' Set the mouse pointer back to normal.
  Screen.MousePointer = vbNormal
  
End Sub
Private Sub RefreshStatusBar()
  Dim iItems As Integer
  Dim iSelections As Integer
  Dim sMessage As String
  Dim sObjectType As String
  
  iItems = ListView1.ListItems.Count
  iSelections = ListView1_SelectedCount
  
  If iItems > 0 Then
    Select Case ListView1.SelectedItem.Tag
      Case giNODE_TABLE
        sObjectType = " table"
      
      Case giNODE_COLUMN
        sObjectType = " column"
      
      Case giNODE_RELATION
        sObjectType = " table"
      
      Case giNODE_RELATIONCHILD
        sObjectType = " relation"
      
      Case Else
        sObjectType = " object"
    End Select
  Else
    sObjectType = " object"
  End If
  
  sMessage = Trim(Str(iItems)) & sObjectType
  If iItems <> 1 Then
    sMessage = sMessage & "s"
  End If
  sMessage = sMessage & ", " & Trim(Str(iSelections)) & " selected."
  
  StatusBar1.Panels(1).Text = sMessage
  
End Sub


Private Sub RemoveNode(psKey As String)

  ' Remove the specified node from the treeview.
  Treeview1.Nodes.Remove psKey
  
  ' Populate the listview with items for the defaulted treeview item.
  PopulateListView Treeview1.SelectedItem, True
  
End Sub

Private Sub SplitMove()
  
  ' Limit the minimum size of the tree and list views.
  If fraSplit.Left < 810 Then
    fraSplit.Left = 810
  ElseIf fraSplit.Left + fraSplit.Width > Me.ScaleWidth - 810 Then
    fraSplit.Left = Me.ScaleWidth - (810 + fraSplit.Width)
  End If
  
  ' Resize the tree view.
  Treeview1.Width = fraSplit.Left - Treeview1.Left
  Label1.Width = Treeview1.Width
  
  ' Resize the listview.
  ListView1.Left = fraSplit.Left + fraSplit.Width
  ListView1.Width = Me.ScaleWidth - ListView1.Left
  Label2.Left = ListView1.Left
  Label2.Width = ListView1.Width
  
  ' Flag that the split move has ended.
  gfSplitMoving = False

End Sub


Public Function ListView1_SelectedCount() As Integer

  Dim iLoop As Integer
  
  ListView1_SelectedCount = 0
  
  ' Loop through the list view items counting how many
  ' are currently selected.
  For iLoop = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(iLoop).Selected = True Then
      ListView1_SelectedCount = ListView1_SelectedCount + 1
    End If
  Next iLoop

End Function

Public Function ListView1_SelectedTag() As Integer

  Dim iLoop As Integer
  
  ListView1_SelectedTag = -1
  
  ' Loop through the list view items counting how many
  ' are currently selected.
  For iLoop = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(iLoop).Selected = True Then
      ListView1_SelectedTag = ListView1.ListItems(iLoop).Tag
      Exit Function
    End If
  Next iLoop
  
  'TM20010917 Fault 2821 & 'TM20010917 Fault 2038
  If ListView1.ListItems.Count > 0 Then
    ListView1_SelectedTag = ListView1.SelectedItem.Tag
  End If

End Function
Private Sub ListView1_ClearSelections()

  Dim iLoop As Integer
  
  ' Loop through the list view items deselecting any currently selected items.
  For iLoop = 1 To ListView1.ListItems.Count
    ListView1.ListItems(iLoop).Selected = False
  Next iLoop

End Sub

Private Sub ColumnDelete()
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  Dim iLoop As Integer
  Dim objTable As HRProSystemMgr.Table
  Dim objColumn As HRProSystemMgr.Column

  fOK = True
  fDeleteAll = False

  ' If we have more than one selection then question the multi-deletion.
  If ListView1_SelectedCount > 1 Then
    
    If MsgBox("Are you sure you want to delete these " & _
      Trim(Str(ListView1_SelectedCount)) & _
      " columns ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
    
      fDeleteAll = True
      fConfirmed = True
    Else
      ListView1.SetFocus
      Exit Sub
    End If
    
  End If
    
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the frmPicMgr form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd
  
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(Treeview1.SelectedItem.key, 2))
  
  ' Loop through all of the listview items to see which ones are selected.
  iLoop = 1
  Do While (iLoop <= ListView1.ListItems.Count) And fOK
    
    ' If the item is selected then delete it.
    If ListView1.ListItems(iLoop).Selected = True Then

      fOK = objTable.ReadTable
      
      If fOK Then
  
        Set objColumn = New HRProSystemMgr.Column
        objColumn.ColumnID = Val(Mid(ListView1.ListItems(iLoop).key, 2))
        objColumn.TableID = objTable.TableID
  
        fOK = objColumn.ReadColumn
        
        If fOK Then
            
          If objColumn.Properties("columnType") = giCOLUMNTYPE_SYSTEM Then
              
            MsgBox "Unable to delete the system column " & _
              objColumn.Properties("columnName") & ".", _
              vbOKOnly + vbExclamation, Application.Name
              
          Else
              
            If Not fDeleteAll Then
              If MsgBox("Are you sure you want to delete the column " & _
                objColumn.Properties("columnName") & _
                " ?", vbYesNo + vbDefaultButton2 + _
                vbQuestion, Application.Name) = vbYes Then
              
                fConfirmed = True
              Else
                fConfirmed = False
              End If
            End If
          
            If fConfirmed Then
              ' Deleted the column from the database.
              fOK = objColumn.DeleteColumn_Transaction
            End If
                  
          End If
        End If
      
        ' Disassociate object variables.
        Set objColumn = Nothing
        
      End If
    End If
  
    iLoop = iLoop + 1
  
  Loop
    
TidyUpAndExit:
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  ' Disassociate object variables.
  Set objTable = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub
Private Sub TableDelete()
  ' Delete the selected tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim aLngTableId() As Long
  Dim objTable As Table
  
  ReDim aLngTableId(0)
  fDeleteAll = False
  fOK = True
  
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd
    
  ' If we have more than one selection then question the multi-deletion.
  If (ActiveView Is ListView1) And (ListView1_SelectedCount > 1) Then
  
    If MsgBox("Are you sure you want to delete these " & _
      Trim(Str(ListView1_SelectedCount)) & _
      " tables ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
      
      fDeleteAll = True
      fConfirmed = True
    Else
      Exit Sub
    End If
    
    ' Read the ids of the tables to be deleted from the listview into an array.
    For iLoop = 1 To ListView1.ListItems.Count
      If ListView1.ListItems(iLoop).Selected = True Then
        iNextIndex = UBound(aLngTableId) + 1
        ReDim Preserve aLngTableId(iNextIndex)
        aLngTableId(iNextIndex) = Val(Mid(ListView1.ListItems(iLoop).key, 2))
      End If
    Next iLoop
  Else
    ' Read the ids of the tables to be deleted from the treeview into an array.
    ReDim aLngTableId(1)
    aLngTableId(1) = Val(Mid(ActiveView.SelectedItem.key, 2))
  End If
    
  Set objTable = New HRProSystemMgr.Table
  
  ' Delete all of the tables in the array..
  For iLoop = 1 To UBound(aLngTableId)
  
    objTable.TableID = aLngTableId(iLoop)
    
    fOK = objTable.ReadTable
    
    If fOK Then
      
      If Not fDeleteAll Then
      
        ' Prompt the user to confirm the deletion.
        If MsgBox("Are you sure you want to delete the table '" & _
          objTable.TableName & "' ?", vbYesNo + vbDefaultButton2 + _
          vbQuestion, Application.Name) = vbYes Then
                          
          fConfirmed = True
        Else
          fConfirmed = False
        End If
      End If
                
      If fConfirmed Then
        ' Delete the table.
        
        fOK = objTable.DeleteTable
        
        If fOK Then
          
          ' Remove table node from treeview
          Treeview1.Nodes.Remove "T" & objTable.TableID
          
          ' Remove table relation node from treeview if it exists.
          If Misc.IsItemInCollection(Treeview1.Nodes, "R" & objTable.TableID) Then
            Treeview1.Nodes.Remove "R" & objTable.TableID
          End If
          
        End If
      End If
    
    End If
    
    If Not fOK Then
      Exit For
    End If
  Next iLoop

  ' Disassociate object variables.
  Set objTable = Nothing

  ' Ensure the relations list is correct.
  If fOK Then
    fOK = CheckAllRelations
  End If
  
TidyUpAndExit:
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub RefreshListView()

  ListView1.Sorted = True
  ListView1.Refresh
    
  RefreshStatusBar
  
  frmSysMgr.RefreshMenu
  
  ListView1.SetFocus

End Sub



Public Property Get UndoEnabled() As Boolean

  ' Return the 'undo enabled' property.
  UndoEnabled = gfUndoEnabled

End Property

Public Property Let UndoEnabled(ByVal pfValue As Boolean)

  ' Set the 'undo enabled' property.
  gfUndoEnabled = pfValue

End Property

Private Function PerformUndo()

End Function

Private Sub EditMenuTable(psMenuItem As String)
  
  Select Case psMenuItem
    
    ' Add a new table
    Case "ID_New"
      gfMenuActionKey = True
      TableAdd
      
    ' Delete a table.
    Case "ID_Delete"
      gfMenuActionKey = True
      TableDelete
      
    ' Update the table properties
    Case "ID_Properties"
      gfMenuActionKey = True
      TableEdit
      
    ' Copy the table.
    Case "ID_CopyDef"  '"ID_CopyTable"
      gfMenuActionKey = True
      TableCopy
      
    ' Print the table definition
    Case "ID_Print"
      gfMenuActionKey = True
      OutputDefintion giNODE_TABLE, giEXPORT_TO_PRINTER
      
    Case "ID_CopyClipboard"
      gfMenuActionKey = True
      OutputDefintion giNODE_TABLE, giEXPORT_TO_CLIPBOARD
      
  End Select

End Sub

Private Sub EditMenuColumn(psMenuItem As String)
    
  Select Case psMenuItem
      
    ' Create a new column
    Case "ID_New"
      gfMenuActionKey = True
      ColumnAdd
            
    ' Delete an existing column
    Case "ID_Delete"
      gfMenuActionKey = True
      ColumnDelete
          
    ' Update the properties of a column
    Case "ID_Properties"
      gfMenuActionKey = True
      ColumnEdit

    ' Copy the currently selected column
    Case "ID_CopyDef"  '"ID_CopyColumn"
      gfMenuActionKey = True
      ColumnCopy
    
    ' Print the definition
    Case "ID_Print"
      gfMenuActionKey = True
      OutputDefintion giNODE_COLUMN, giEXPORT_TO_PRINTER
    
    ' Copy definition to clipboard
    Case "ID_CopyClipboard"
      gfMenuActionKey = True
      OutputDefintion giNODE_COLUMN, giEXPORT_TO_CLIPBOARD
    
  End Select
          
  ' Dodgy fix to avoid locking the dodgy toolbar.
'  With frmSysMgr.tbMain
'    .Redraw = False
'    .Enabled = False
'    .Enabled = True
'    .Redraw = True
'  End With

End Sub

Private Sub EditMenuRelation(psMenuItem As String)
      
  Select Case psMenuItem
        
    ' Create a new relationship
    Case "ID_New"
      gfMenuActionKey = True
      RelationAdd
          
    ' Delete a relationship
    Case "ID_Delete"
      gfMenuActionKey = True
      RelationDelete
      
    ' Update the relationship's properties
    Case "ID_Properties"
      gfMenuActionKey = True
      RelationEdit Val(Mid(ActiveView.SelectedItem.key, 2))
    
    ' Print the definition
    Case "ID_Print"
      gfMenuActionKey = True
      OutputDefintion giNODE_RELATION, giEXPORT_TO_PRINTER
      
  End Select

End Sub

Private Sub ColumnAdd()
  ' Add a new column to the currently selected table.
  Dim fOK As Boolean
  Dim objTable As HRProSystemMgr.Table
  Dim objColumn As HRProSystemMgr.Column
  Dim objItem As ComctlLib.ListItem
  
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(Treeview1.SelectedItem.key, 2))
      
  fOK = objTable.ReadTable
  
  If fOK Then
    Set objColumn = New HRProSystemMgr.Column
    objColumn.TableID = objTable.TableID
    fOK = objColumn.NewColumn
        
    'JDM - 13/08/01 - Fault 2678 - Add an item to the listview for the column, so it can be selected later.
    If fOK Then
      Set objItem = ListView1.ListItems.Add(, "C" & objColumn.ColumnID, "Inserting Column...", "IMG_COLUMN", "IMG_COLUMN")
      ListView1_ClearSelections
      objItem.Selected = True
      objItem.Tag = giNODE_COLUMN
    End If
  End If
  
  ' Disassociate object variables.
  Set objItem = Nothing
  Set objColumn = Nothing
  Set objTable = Nothing

End Sub


Private Sub TableAdd()
  ' Create a new table.
  Dim sImage As String
  Dim objTable As Table
  Dim objNode As ComctlLib.Node
      
  ' Instantiate a new Table object.
  Set objTable = New HRProSystemMgr.Table

  If objTable.NewTable Then
        
    ' Add the new table to TreeView.
    Select Case objTable.TableType
      Case iTabParent
        sImage = "IMG_PARENTTABLE"
      Case iTabChild
        sImage = "IMG_CHILDTABLE"
      Case iTabLookup
        sImage = "IMG_LOOKUP"
      Case Else
        sImage = "IMG_TABLE"
    End Select
    
    Set objNode = Treeview1.Nodes.Add("TABLES", _
      tvwChild, "T" & objTable.TableID, objTable.TableName, _
      sImage, sImage)
    objNode.Tag = giNODE_TABLE
    objNode.Sorted = True
  
    'Make sure new table is visible and selected in the treeview.
    objNode.EnsureVisible
    Treeview1.SelectedItem = objNode
  
    ' Disassociate object variables.
    Set objNode = Nothing
    
    ListView1.SetFocus

  End If

  ' Disassociate object variables.
  Set objTable = Nothing
  
End Sub

Private Sub EditMenuRelationChild(psMenuItem As String)
      
  Select Case psMenuItem
        
    ' Create a new relationship
    Case "ID_New"
      gfMenuActionKey = True
      RelationChildAdd
          
    ' Delete a relationship
    Case "ID_Delete"
      gfMenuActionKey = True
      RelationChildDelete
      
    ' Update the relationship's properties
    Case "ID_Properties"
      gfMenuActionKey = True
      RelationChildEdit
      
  End Select

End Sub

Private Sub ColumnEdit()
  ' Edit the selected column's properties.
  Dim objTable As HRProSystemMgr.Table
  Dim objColumn As HRProSystemMgr.Column

  Screen.MousePointer = vbHourglass
  
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(Treeview1.SelectedItem.key, 2))

  Set objColumn = New HRProSystemMgr.Column
  objColumn.TableID = objTable.TableID
  Set objTable = Nothing
  objColumn.ColumnID = Val(Mid(ListView1.SelectedItem.key, 2))
  
  If objColumn.ReadColumn Then
    If objColumn.Properties("columnType") = giCOLUMNTYPE_SYSTEM Then
      MsgBox "Unable to edit the system column " & _
        objColumn.Properties("columnName") & ".", _
        vbOKOnly + vbExclamation, Application.Name
    Else
      objColumn.EditColumn (False)
    End If
  End If
      
  ' Disassociate object variables.
  Set objColumn = Nothing

End Sub
Private Sub ColumnCopy()
  ' Copy the selected column's properties.
  Dim objTable As HRProSystemMgr.Table
  Dim objColumn As HRProSystemMgr.Column
  Dim objItem As ComctlLib.ListItem
  
  Screen.MousePointer = vbHourglass
  
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(Treeview1.SelectedItem.key, 2))

  Set objColumn = New HRProSystemMgr.Column
  objColumn.TableID = objTable.TableID
  Set objTable = Nothing
  objColumn.ColumnID = Val(Mid(ListView1.SelectedItem.key, 2))
  
  If objColumn.ReadColumn Then

    ' Set object to copy itself
    Set objColumn = objColumn.CloneColumn(True)
    'objColumn.IsChanged = True

    If objColumn.Properties("columnType") = giCOLUMNTYPE_SYSTEM Then
      MsgBox "Unable to edit the system column " & _
        objColumn.Properties("columnName") & ".", _
        vbOKOnly + vbExclamation, Application.Name
    Else
      If objColumn.EditColumn(True) Then
        'JDM - 13/08/01 - Fault 2678 - Add an item to the listview for the column, so it can be selected later.
        Set objItem = ListView1.ListItems.Add(, "C" & objColumn.ColumnID, "Inserting Column...", "IMG_COLUMN", "IMG_COLUMN")
        ListView1_ClearSelections
        objItem.Selected = True
        objItem.Tag = giNODE_COLUMN
      End If
    End If
  End If

  ' Disassociate object variables.
  Set objItem = Nothing
  Set objColumn = Nothing

End Sub
Private Sub TableEdit()
  ' Edit the table's properties.
  Dim objTable As HRProSystemMgr.Table
  Dim objNode As ComctlLib.Node
  Dim sImage As String

  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(ActiveView.SelectedItem.key, 2))
  
  If objTable.EditTable Then
    
    'MH20010131 Fault 1673
    ' Check which image we need for table
    Select Case objTable.TableType
    Case iTabParent
      sImage = "IMG_PARENTTABLE"
    Case iTabChild
      sImage = "IMG_CHILDTABLE"
    Case iTabLookup
      sImage = "IMG_LOOKUP"
    Case Else
      sImage = "IMG_TABLE"
    End Select
    
    'JDM - 22/11/01 - Fault 3187 - Moved update images code around a bit to stop runtime error
    
    If ActiveView Is ListView1 Then
    
      Set objNode = Treeview1.Nodes("T" & objTable.TableID)
      objNode.EnsureVisible
      Treeview1.SelectedItem = objNode
      Treeview1.SelectedItem.Text = objTable.TableName
      Treeview1.SelectedItem.Image = sImage             'MH20010131 Fault 1673
      Treeview1.SelectedItem.SelectedImage = sImage     'MH20010131 Fault 1673

      ' Disassociate object variables.
      Set objNode = Nothing
    Else
      
      ' Update the tree and list views in case the table name has been changed.
      ActiveView.SelectedItem.Text = objTable.TableName
      ActiveView.SelectedItem.Image = sImage             'MH20010131 Fault 1673
      ActiveView.SelectedItem.SelectedImage = sImage     'MH20010131 Fault 1673
    End If
  
    ' Sort the treeview order in case the name of a table has changed.
    Treeview1.Nodes("TABLES").Sorted = True
  
    If Misc.IsItemInCollection(Treeview1.Nodes, "R" & objTable.TableID) Then
      Treeview1.Nodes("R" & objTable.TableID).Text = objTable.TableName
      'JPD 20051128 Fault 10597
      'Treeview1.Nodes("R" & objTable.TableID).Image = sImage             'MH20010131 Fault 1673
      'Treeview1.Nodes("R" & objTable.TableID).SelectedImage = sImage     'MH20010131 Fault 1673
    End If
        
  End If

  ' Disassociate object variables.
  Set objTable = Nothing

End Sub

Private Sub RelationAdd()
  Dim frmRelation As HRProSystemMgr.frmRelate
           
  ' Pop-up the form to take the relationship values.
  Set frmRelation = New HRProSystemMgr.frmRelate
  frmRelation.Show vbModal
  ' Disassociate object variables.
  Set frmRelation = Nothing
 
  ' Refresh the relations display.
  CheckAllRelations

End Sub
Private Sub RelationDelete()
  ' Delete all relations from the selected table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim aLngTableId() As Long
  Dim objTable As HRProSystemMgr.Table
  
  fOK = True
  ReDim aLngTableId(0)
  fDeleteAll = False

  ' If we have more than one selection then question the multi-deletion.
  If (ActiveView Is ListView1) And (ListView1_SelectedCount > 1) Then
  
    If MsgBox("Are you sure you want to delete all relations for these " & _
      Trim(Str(ListView1_SelectedCount)) & _
      " tables ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
      
      fDeleteAll = True
      fConfirmed = True
    Else
      Exit Sub
    End If
    
    ' Read the ids of the tables to be deleted from the listview into an array.
    For iLoop = 1 To ListView1.ListItems.Count
      If ListView1.ListItems(iLoop).Selected = True Then
        iNextIndex = UBound(aLngTableId) + 1
        ReDim Preserve aLngTableId(iNextIndex)
        aLngTableId(iNextIndex) = Val(Mid(ListView1.ListItems(iLoop).key, 2))
      End If
    Next iLoop
  
  Else
    ' Read the ids of the tables to be deleted from the treeview into an array.
    ReDim aLngTableId(1)
    aLngTableId(1) = Val(Mid(ActiveView.SelectedItem.key, 2))
  End If
      
  Set objTable = New HRProSystemMgr.Table
  
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd
  
  ' Delete all of the tables in the array..
  For iLoop = 1 To UBound(aLngTableId)
  
    objTable.TableID = aLngTableId(iLoop)
      
    fOK = objTable.ReadTable
    
    If fOK Then
      
      If Not fDeleteAll Then
        
        ' Prompt the user to confirm the deletion.
        If MsgBox("Are you sure you want to delete all relations from " & _
          objTable.TableName & " ?", _
          vbYesNo + vbDefaultButton2 + vbQuestion, Application.Name) = vbYes Then
                            
          fConfirmed = True
        Else
          fConfirmed = False
        End If
      End If
                  
      If fConfirmed Then
        'Delete all relations for selected table
        fOK = objTable.DeleteAllRelations_Transaction
          
        If fOK Then
          'Check if parent table has any relations remaining
          fOK = CheckRelations(objTable)
        End If
      End If
    
    End If
    
    If Not fOK Then
      Exit For
    End If
  Next iLoop
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objTable = Nothing
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub
Private Sub RelationEdit(psKey As String)
  Dim objTable As HRProSystemMgr.Table
  Dim frmRelation As HRProSystemMgr.frmRelate
      
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = psKey
      
  If objTable.ReadTable Then
        
    Set frmRelation = New HRProSystemMgr.frmRelate
    Set frmRelation.ParentTable = objTable
    frmRelation.Show vbModal
    Set frmRelation = Nothing
  
    CheckAllRelations

  End If
    
  ' Disassociate object variables.
  Set objTable = Nothing
  
End Sub


Private Sub RelationChildAdd()

  ' Adding a relation child is the same as looking at the properties of
  ' a relation parent.
  RelationEdit Val(Mid(Treeview1.SelectedItem.key, 2))
  
End Sub
Private Sub RelationChildDelete()
  ' Delete the selected relationship.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim aRelationId() As Variant
  Dim objTable As Table
  
  fOK = True
  ReDim aRelationId(2, 0)
  fDeleteAll = False

  ' If we have more than one selection then question the multi-deletion.
  If (ListView1_SelectedCount > 1) Then
    If MsgBox("Are you sure you want to delete these " & _
      Trim(Str(ListView1_SelectedCount)) & _
      " relations ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
      
      fDeleteAll = True
      fConfirmed = True
    Else
      Exit Sub
    End If
  End If
  
  ' Read the ids of the tables to be deleted from the listview into an array.
  For iLoop = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(iLoop).Selected = True Then
      iNextIndex = UBound(aRelationId, 2) + 1
      ReDim Preserve aRelationId(2, iNextIndex)
      aRelationId(1, iNextIndex) = Val(Mid(ListView1.ListItems(iLoop).key, 2))
      aRelationId(2, iNextIndex) = ListView1.ListItems(iLoop).Text
    End If
  Next iLoop
           
  Set objTable = New HRProSystemMgr.Table
  objTable.TableID = Val(Mid(Treeview1.SelectedItem.key, 2))
  
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the frmPicMgr form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd
  
  ' Delete all of the tables in the array..
  For iLoop = 1 To UBound(aRelationId, 2)
  
    fOK = objTable.ReadTable
    
    If fOK Then
      
      If Not fDeleteAll Then
        
        ' Prompt the user to confirm the deletion.
        If MsgBox("Are you sure you want to delete the relation between " & _
          objTable.TableName & vbCr & _
          "and " & aRelationId(2, iLoop) & " ?", _
          vbYesNo + vbDefaultButton2 + vbQuestion, Application.Name) = vbYes Then
                            
          fConfirmed = True
        Else
          fConfirmed = False
        End If
      End If
                  
      If fConfirmed Then
        ' Delete the selected relation
        fOK = objTable.DeleteRelation_Transaction(aRelationId(1, iLoop))
          
        If fOK Then
          'Check if parent table has any relations remaining
          fOK = CheckRelations(objTable)
        End If
      End If
    
    End If
          
    If Not fOK Then
      Exit For
    End If
  Next iLoop

TidyUpAndExit:
  ' Disassociate object variables.
  Set objTable = Nothing
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub
Private Sub RelationChildEdit()

  ' Adding a relation child is the same as looking at the properties of
  ' a relation parent.
  RelationEdit Val(Mid(Treeview1.SelectedItem.key, 2))

End Sub


Private Function CheckAllRelations() As Boolean
  ' Refresh the Relations display in the treeview.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim objTable As HRProSystemMgr.Table
  Dim aLngTableIds() As Long
  
  ReDim aLngTableIds(0)
  
  ' Create an array of the existing tables.
  With recTabEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      
      If Not .Fields("deleted") Then
        iIndex = UBound(aLngTableIds) + 1
        ReDim Preserve aLngTableIds(iIndex)
        aLngTableIds(iIndex) = .Fields("tableID")
      End If
      
      .MoveNext
    Loop
  End With
    
  ' Refresh the relations display for each table.
  Set objTable = New HRProSystemMgr.Table
  For iIndex = 1 To UBound(aLngTableIds)
    objTable.TableID = aLngTableIds(iIndex)
    fOK = objTable.ReadTable
    
    If fOK Then
      fOK = CheckRelations(objTable)
    End If
    
    If Not fOK Then
      Exit For
    End If
  Next iIndex
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objTable = Nothing
  CheckAllRelations = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub TableCopy()
  ' Copy the current table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  'NPG20080207 Fault 12874
  ' Dim frmCopy As frmCopyTable
  Dim frmPermissions As frmDefaultPermissions2
  Dim objNewTable As Table
  Dim objSourceTable As Table
  Dim objNode As ComctlLib.Node
  Dim mbGrantRead As Boolean
  Dim mbGrantEdit As Boolean
  Dim mbGrantNew As Boolean
  Dim mbGrantDelete As Boolean
  Dim miTableType As CopySecurityType
  Dim sSQL As String
  Dim rsCheck As dao.Recordset
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim objExpr As CExpression
  Dim asSpecialFunctions() As String
  Dim frmUse As frmUsage
  Dim iLoop As Integer
  
  'NPG20080207 Fault 12874
  ' Ask the user if they want to copy data as well.
  ' Set frmCopy = New frmCopyTable
  Set frmPermissions = New frmDefaultPermissions2
  
  'JPD 20060929 Fault 11462
  ' Cannot copy tables that use the hierarchy functions as they are based on
  ' specific tables defined in module setup.
  ReDim asSpecialFunctions(0)
  sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
    " FROM tmpComponents" & _
    " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
    " AND tmpComponents.functionID IN (65,66,67,68,69,70,71,72,73)"
  Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  Do While Not rsCheck.EOF

    Set objComp = New CExprComponent
    objComp.ComponentID = rsCheck!ComponentID
    lngExprID = objComp.RootExpressionID
    Set objComp = Nothing

    ' Get the expression name and type description.
    Set objExpr = New CExpression
    objExpr.ExpressionID = lngExprID

    If objExpr.ReadExpressionDetails Then
      If objExpr.BaseTableID = Val(Mid(ActiveView.SelectedItem.key, 2)) Then
        ReDim Preserve asSpecialFunctions(UBound(asSpecialFunctions) + 1)
        
        If objExpr.ExpressionType = giEXPR_RUNTIMECALCULATION _
           Or objExpr.ExpressionType = giEXPR_RUNTIMEFILTER _
           Or objExpr.ExpressionType = giEXPR_EMAIL Then
        
          asSpecialFunctions(UBound(asSpecialFunctions)) = GetExpressionUsageDescFromSQL(lngExprID)
        Else
          asSpecialFunctions(UBound(asSpecialFunctions)) = GetExpressionUsageDesc(lngExprID)
        End If
      End If
    End If

    ' Disassociate object variables.
    Set objExpr = Nothing
      
    rsCheck.MoveNext
  Loop
  
  ' Close the recordset.
  rsCheck.Close
  
  If UBound(asSpecialFunctions) > 0 Then
    Set frmUse = New frmUsage
    frmUse.ResetList
    
    For iLoop = 1 To UBound(asSpecialFunctions)
      frmUse.AddToList asSpecialFunctions(iLoop)
    Next iLoop
    
    Screen.MousePointer = vbNormal
    frmUse.ShowMessage GetTableName(Val(Mid(ActiveView.SelectedItem.key, 2))) & " Table", "Unable to copy this table as the following expressions use the Hierarchy functions : ", UsageCheckObject.Table
    UnLoad frmUse
    Set frmUse = Nothing

    Exit Sub
  End If
  
  'What type of table are we copying (set different options of the copytable form)
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", Val(Mid(ActiveView.SelectedItem.key, 2))
    If Not .NoMatch Then
      miTableType = .Fields("TableType")
      'NPG20080207 Fault 12874
      ' frmCopy.SetOptions miTableType, .Fields("New"), .Fields("CopySecurityTableID"), .Fields("TableName"), giQuestion
    End If
  End With
  
  ' Show the copy form
  'NPG20080207 Fault 12874
  ' frmCopy.Show vbModal
  ' fOK = Not frmCopy.Cancelled
  fOK = True
  
  'JDM - 25/06/01 - Fault 551 - Apply permissions to existing security groups
  If fOK Then
    'NPG20080207 Fault 12874
    'If frmCopy.CopySecurity = False Then
      frmPermissions.SetType "copy", miTableType
      frmPermissions.Show vbModal
      If frmPermissions.OkCancel = vbOK Then
        mbGrantRead = frmPermissions.GrantRead
        mbGrantNew = frmPermissions.GrantNew
        mbGrantEdit = frmPermissions.GrantEdit
        mbGrantDelete = frmPermissions.GrantDelete
      Else
        fOK = False
      End If
    'End If
  End If
  
  If fOK Then
  
    '# RH 19/05/00 - Show a progress bar with copytable video whilst the
    '                table is being copied.
    With gobjProgress
      '.AviFile = App.Path & "\videos\copytable.avi"
      .AVI = dbCopyTable
      .MainCaption = "Copy Table"
      .NumberOfBars = 0
      .Cancel = False
      .Time = False
      .Caption = "Copying..."
      .OpenProgress
    End With
    
    ' Instantiate a table object.
    Set objSourceTable = New HRProSystemMgr.Table
    objSourceTable.TableID = Val(Mid(ActiveView.SelectedItem.key, 2))

    ' Get the source table to clone itself.
    
    'NPG20080207 Fault 12874
    ' Set objNewTable = objSourceTable.CloneTable_Transaction(frmCopy.CopyData, frmCopy.CopySecurity, mbGrantRead, mbGrantNew, mbGrantEdit, mbGrantDelete)
    Set objNewTable = objSourceTable.CloneTable_Transaction(frmPermissions.CopyData, frmPermissions.CopyPermissions, mbGrantRead, mbGrantNew, mbGrantEdit, mbGrantDelete)
    
    'JPD 20030829 Fault 5538 - the default permissions part was already commented out.
    ' I've commented out the PermissionsPrompted update as it is now part of the
    ' Table class's CloneTable method (so that it's set for all child tables that
    ' are cloned as their parent is cloned).
    
    ' Pass in the default security permissions for the new table
'''    With recTabEdit
'''      .Index = "idxTableID"
'''      .Seek "=", objNewTable.TableID
'''      If Not .NoMatch Then
'''        .Edit
''''        If Not frmCopy.CopySecurity Then
''''          .Fields("GrantRead") = mbGrantRead
''''          .Fields("GrantNew") = mbGrantNew
''''          .Fields("GrantEdit") = mbGrantEdit
''''          .Fields("GrantDelete") = mbGrantDelete
''''          .Fields("copySecurityTableID") = 0
''''          .Fields("copySecurityTableName") = ""
''''        Else
''''          .Fields("GrantRead") = 0
''''          .Fields("GrantNew") = 0
''''          .Fields("GrantEdit") = 0
''''          .Fields("GrantDelete") = 0
''''          .Fields("copySecurityTableID") = objSourceTable.TableID
''''          .Fields("copySecurityTableName") = objSourceTable.TableName
''''        End If
'''        .Fields("PermissionsPrompted") = True
'''        .Update
'''      End If
'''    End With
    
    ' Tidy up
    Set objSourceTable = Nothing
    gobjProgress.CloseProgress
    
    ' Check that the clone was created.
    fOK = Not objNewTable Is Nothing
    
    If fOK Then
      ' Populate the treeview with the cloned table(s) included..
      InitialiseTreeView

      ' Make sure new table is visible and selected in the treeview.
      Set objNode = Treeview1.Nodes("T" & objNewTable.TableID)
      objNode.EnsureVisible
      Treeview1.SelectedItem = objNode
      Set objNode = Nothing
      
      ListView1.SetFocus
    End If
  
    ' Disassociate object variables.
    Set objNewTable = Nothing
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  
  'NPG20080207 Fault 12874
  Set frmPermissions = Nothing
  Set objNewTable = Nothing
  Set objSourceTable = Nothing
  Set objNode = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Function ChangeView(ViewStyle As ComctlLib.ListViewConstants) As Boolean

  Static InChangeView As Boolean
  On Error GoTo ErrorTrap

  If InChangeView Then Exit Function

  InChangeView = True

'TM20010914 Fault 1753
'As ActiveBar does not support mutual exclusivity on its tools, the following code
'ensures only one of the view options is selected at any one time.

  Me.ListView1.View = ViewStyle
  
  With abDbMgr
    Select Case ViewStyle
      Case lvwIcon
        .Tools("ID_LargeIcons").Checked = True
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = True
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
        frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = False
      
      Case lvwSmallIcon
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = True
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = True
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
        frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = False
      
      Case lvwList
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = True
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = True
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
        frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = False
      
      Case lvwReport
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = True
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = True
        .Tools("ID_CustomiseColumns").Enabled = True
        frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = True
    
    End Select
  End With

  InChangeView = False

  ChangeView = True

  Exit Function

ErrorTrap:
  ChangeView = False
  Err = False

End Function

' Prints a defintion for selected table
Private Sub OutputDefintion(piViewType As HRProSystemMgr.ViewItemTypes _
                            , piOutputType As HRProSystemMgr.OutputDefintionTypes)

  Dim objTable As HRProSystemMgr.Table
  Dim objColumn As HRProSystemMgr.Column
  Dim objItem As ComctlLib.ListItem
  Dim bOK As Boolean
   
  ' Output the appropriate defintion type
  Select Case piViewType
    
    Case giNODE_TABLE
      Set objTable = New HRProSystemMgr.Table
      objTable.TableID = Val(Mid(ActiveView.SelectedItem.key, 2))
      objTable.ReadTable
      bOK = objTable.PrintDefinition(piOutputType)
      
    Case giNODE_COLUMN
      Set objColumn = New HRProSystemMgr.Column
      objColumn.ColumnID = Val(Mid(ActiveView.SelectedItem.key, 2))
      objColumn.ReadColumn
      bOK = objColumn.PrintDefinition(piOutputType)
      
    Case giNODE_RELATION
      
  End Select
  
  ' Disassociate object variables.
  Set objItem = Nothing
  Set objColumn = Nothing
  Set objTable = Nothing

End Sub

Private Function GetRecordDescriptionDetails(lngRecDescExprID As Long) As String
  ' Get the Record Description details.
  Dim sExprName As String

  ' Initialize the default expression name.
  sExprName = ""

  If lngRecDescExprID > 0 Then

    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", lngRecDescExprID, False

      ' Read the expression's name from the recordset.
      If Not .NoMatch Then
        sExprName = !Name
      End If

    End With
  End If

  GetRecordDescriptionDetails = sExprName

End Function
