VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmScrOpen 
   Caption         =   "Screen Manager"
   ClientHeight    =   5670
   ClientLeft      =   315
   ClientTop       =   1665
   ClientWidth     =   5595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5031
   Icon            =   "frmScrOpen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   5595
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Height          =   400
      Index           =   5
      Left            =   4230
      TabIndex        =   7
      Top             =   2925
      Width           =   1290
   End
   Begin VB.Frame fraTable 
      Caption         =   "Table :"
      Height          =   810
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   3990
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Proper&ties..."
      Height          =   400
      Index           =   4
      Left            =   4230
      TabIndex        =   6
      Top             =   2355
      Width           =   1290
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cop&y..."
      Height          =   400
      Index           =   3
      Left            =   4230
      TabIndex        =   4
      Top             =   1245
      Width           =   1290
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Delete"
      Height          =   400
      Index           =   2
      Left            =   4230
      TabIndex        =   5
      Top             =   1800
      Width           =   1290
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Edit..."
      Height          =   400
      Index           =   1
      Left            =   4230
      TabIndex        =   3
      Top             =   690
      Width           =   1290
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&New..."
      Height          =   400
      Index           =   0
      Left            =   4230
      TabIndex        =   2
      Top             =   150
      Width           =   1290
   End
   Begin ComctlLib.StatusBar sbScrOpen 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   5355
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9340
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid ssGrdScreens 
      Height          =   4215
      Left            =   105
      TabIndex        =   1
      Top             =   1035
      Width           =   3990
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      ColumnHeaders   =   0   'False
      Col.Count       =   3
      DefColWidth     =   26458
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3201
      Columns(0).Caption=   "Screen"
      Columns(0).Name =   "Screen"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   26458
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Parent Table"
      Columns(1).Name =   "Parent Table"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   26458
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Screen ID"
      Columns(2).Name =   "Screen ID"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      UseDefaults     =   0   'False
      TabNavigation   =   1
      _ExtentX        =   7038
      _ExtentY        =   7435
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmScrOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event UnLoad()

Private gfLoading As Boolean
Private gLngLockedTableID As Long
Private mstrCurrentTable As String
Private mIamMaximised As Boolean
Private mlngScreenID As Long
Private mstrScreenName As String

Private Const MIN_FORM_HEIGHT = 5000
Private Const MIN_FORM_WIDTH = 6000

Public Property Get ScreenID() As Long
  ' Return the ID of the selected screen.
  If ssGrdScreens.Rows > 0 And ssGrdScreens.Row >= 0 Then
    ScreenID = ssGrdScreens.Columns(2).CellValue(ssGrdScreens.Bookmark)
  Else
    ScreenID = 0
  End If
  
End Property

Private Property Get ScreenName() As String
  ' Return the name of the selected screen.
  If ssGrdScreens.Rows > 0 And ssGrdScreens.Row >= 0 Then
    ScreenName = ssGrdScreens.Columns(0).CellValue(ssGrdScreens.Bookmark)
  Else
    ScreenName = vbNullString
  End If
  
End Property



Private Sub cboTable_Click()
  RefreshScreens
End Sub

Private Sub cmdAction_Click(Index As Integer)

  mstrCurrentTable = cboTable.Text
  
  mIamMaximised = Me.WindowState
  Select Case Index
  
    Case 0: NewScreen
    Case 1: OpenScreen
    Case 2: DeleteScreen
    Case 3: CopyScreen
    Case 4: EditScreen
    Case 5: UnLoad Me
    
  End Select

End Sub

Public Function EditMenu(psMenuOption As String) As Boolean

  On Error GoTo EditMenu_ERROR
  
  mstrCurrentTable = cboTable.Text

  Select Case psMenuOption
  
    Case "ID_New": NewScreen
    Case "ID_Open": OpenScreen
    Case "ID_Delete": DeleteScreen
    Case "ID_CopyDef": CopyScreen
    Case "ID_ScreenProperties": EditScreen
    
  End Select

  EditMenu = True
  Exit Function
  
EditMenu_ERROR:
  
  MsgBox "Error whilst performing toolbar action." & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly
  EditMenu = False
  
End Function


Private Sub Form_Activate()
  
  ' Refresh the list of screens displayed.
  gfLoading = True
  
  LoadTableCombo
  
  RefreshScreens
  
  gfLoading = False
  
End Sub

Private Sub Form_Deactivate()
  
  ' Refresh the menu bar.
  frmSysMgr.RefreshMenu
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Pass any keystrokes onto the toolbar in the frmSysMgr form.
  'frmSysMgr.ActiveBar1.OnKeyDown KeyCode, Shift

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
  
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

End Sub


Private Sub Form_Load()
  
  With ssGrdScreens
    .Columns(0).Width = .Width
  End With
  
  gLngLockedTableID = 0
  
End Sub

Private Sub Form_Resize()
  
  On Error GoTo ErrorTrap
  
  Const XGAP = 150
  Const XGAP_RIGHT = 250
  
  Const YGAP = 100
  Const YGAP_BOTTOM = 650

  Dim pintButtonLeft As Integer
  Dim pintButtonCount As Integer
  Dim lngColumnWidth As Long
  
  frmSysMgr.RefreshMenu
  
  With fraTable
    .Width = Me.Width - XGAP_RIGHT - cmdAction(0).Width - XGAP - .Left
    
    Me.cboTable.Width = .Width - cboTable.Left - XGAP
    Me.ssGrdScreens.Width = .Width
    
    pintButtonLeft = .Left + .Width + (XGAP / 2)
    For pintButtonCount = 0 To 5
      cmdAction(pintButtonCount).Left = pintButtonLeft
    Next pintButtonCount
  End With
  
  
  With ssGrdScreens
    .Height = Me.Height - YGAP_BOTTOM - YGAP - sbScrOpen.Height - .Top
    .Columns(0).Width = .Width
      
    lngColumnWidth = .Width - (UI.GetSystemMetrics(SM_CXBORDER) * 2 * Screen.TwipsPerPixelX)
  
    ' Cater for the display of the vertical scroll bar.
    If .Rows > .VisibleRows Then
      lngColumnWidth = lngColumnWidth - _
        (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
    End If
    .Columns(0).Width = lngColumnWidth
  End With

  cmdAction(5).Top = ssGrdScreens.Top + ssGrdScreens.Height - cmdAction(5).Height
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Exit Sub
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Update the menu.
  frmSysMgr.RefreshMenu True
  
End Sub

'Public Sub EditMenu(ByVal psMenuItem As String)
'  ' Handle the menu options.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  Select Case psMenuItem
'    ' Create a new screen.
'    Case "ID_New"
'      NewScreen
'
'    ' Open the Screen Manager form for the selected screen.
'    Case "ID_Open"
'      OpenScreen
'
'    ' Delete the selected screen.
'    Case "ID_Delete"
'      DeleteScreen
'
'    ' Copy the screen.
'    Case "ID_CopyScreen"
'      CopyScreen
'
'    ' Display the screen properties form for the selected screen.
'    Case "ID_ScreenProperties"
'      EditScreen
'  End Select
'
'TidyUpAndExit:
'  Exit Sub
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Sub

Private Function GetScreenForm(ByVal TableID As Long) As Form
  ' Return the Screen Manager form for the given screen ID, if it exists.
  Dim iLoop As Integer
  
  For iLoop = 1 To Forms.Count - 1
    If Forms(iLoop).Name = "frmScrDesigner2" Then
      If Forms(iLoop).TableID = ScreenID Then
        Set GetScreenForm = Forms(iLoop)
        Exit For
      End If
    End If
  Next iLoop

End Function



Public Sub RefreshScreens()
  ' Populate the grid with screens.
  Dim iSelectedScreen As Integer
  Dim lngScreenID As Long
  Dim lngColumnWidth As Long
  Dim sSQL As String
  Dim frmScreen As frmScrDesigner2
  Dim rsScreens As DAO.Recordset
  
  lngScreenID = ScreenID

  ' Get the list of ALL screens from the database.
  sSQL = "SELECT tmpScreens.screenID, tmpScreens.name, tmpTables.tableName, tmpTables.tableID" & _
    " FROM tmpScreens" & _
    " INNER JOIN tmpTables ON tmpTables.tableID = tmpScreens.tableID" & _
    " WHERE tmpScreens.deleted = FALSE" & _
    " AND tmptables.tablename = '" & cboTable.Text & "'" & _
    " ORDER BY tmpScreens.name"
  Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
'  If rsScreens.BOF And rsScreens.EOF Then
'    frmScrDesigner2.ScreenID = 0
'  End If
  
  ' Populate the Screens grid.
  ssGrdScreens.RemoveAll
  
  iSelectedScreen = 0
  
  Do While Not rsScreens.EOF
  
    ssGrdScreens.AddItem rsScreens!Name & _
      vbTab & rsScreens!TableName & _
      vbTab & rsScreens!ScreenID
        
    If rsScreens!ScreenID = lngScreenID Then
      iSelectedScreen = ssGrdScreens.Rows - 1
    End If
    
    rsScreens.MoveNext
  Loop
  
  ssGrdScreens.Bookmark = iSelectedScreen
  ssGrdScreens.SelBookmarks.Add ssGrdScreens.Bookmark
  ssGrdScreens_Click

  rsScreens.Close
  Set rsScreens = Nothing
  
  ' Ensure the grid columns are correctly sized.
  If gfLoading Then
    With ssGrdScreens
      
      lngColumnWidth = .Width - (UI.GetSystemMetrics(SM_CXBORDER) * 2 * Screen.TwipsPerPixelX)
    
      ' Cater for the display of the vertical scroll bar.
      If .Rows > .VisibleRows Then
        lngColumnWidth = lngColumnWidth - _
          (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
      End If
      .Columns(0).Width = lngColumnWidth
    End With
  End If
  
  Me.ssGrdScreens.ScrollBars = 4
  
  ' Update the status bar.
  sbScrOpen.Panels(1).Text = ssGrdScreens.Rows & " screen" & _
    IIf(ssGrdScreens.Rows = 1, vbNullString, "s")
  
  ' Refresh the menu
  frmSysMgr.RefreshMenu
  
End Sub

'Private Sub SSTBScreen_ToolClick(ByVal pTool As ActiveToolBars.SSTool)
'
'  Select Case pTool.ID
'
'    Case "ID_New"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.ID
'
'    Case "ID_Open"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.ID
'
'    Case "ID_Delete"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.ID
'
'    Case "ID_CopyScreen"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.ID
'
'    Case "ID_ScreenProperties"
'      ' Pass the menu choice onto the active form to process.
'      EditMenu pTool.ID
'
'  End Select
'
'End Sub

Private Sub ssGrdScreens_Click()
'  frmScrDesigner.ScreenID = ScreenID
'frmScrDesigner2.ScreenID = ScreenID
  
  ' Refresh the menu
  frmSysMgr.RefreshMenu

End Sub

Private Sub ssGrdScreens_DblClick()
  
  ' Open the selected screen.
  mstrCurrentTable = cboTable.Text
  OpenScreen

End Sub


Private Sub ssGrdScreens_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass any keystrokes onto the toolbar in the frmSysMgr form.
  'frmSysMgr.ActiveBar1.OnKeyDown KeyCode, Shift
  'KeyCode = 0

  'If enter is pressed, load the currently selected screen.  RH 23/07
  If KeyCode = 13 Then ssGrdScreens_DblClick
  
'  If KeyCode > 64 And KeyCode < 91 Then
'
'    ssGrdScreens.MoveFirst
'    Dim intRowCount As Integer
'    For intRowCount = 0 To ssGrdScreens.Rows - 1
'      If LCase(Left(ssGrdScreens.Columns(0).Text, 1)) = LCase(Chr(KeyCode)) Then
'        ssGrdScreens.SelBookmarks.RemoveAll
'        ssGrdScreens.SelBookmarks.Add ssGrdScreens.Bookmark
'        Exit For
'      End If
'      ssGrdScreens.MoveNext
'    Next
'  End If
  
End Sub

Private Sub ssGrdScreens_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Dim lXMouse As Long
'  Dim lYMouse As Long
'
'  ' Pop-up the menu if the right mouse button is pressed.
'  If Button = vbRightButton Then
'    UI.GetMousePos lXMouse, lYMouse
'    frmSysMgr.RefreshMenu
'    frmSysMgr.tbMain.PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
'  End If

End Sub



Private Function CopyScreen() As Boolean
  ' Copy the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fGoodName As Boolean
  Dim iCounter As Integer
  Dim lngScreenID As Long
  Dim sSQL As String
  Dim sScreenName As String
  Dim rsScreen As DAO.Recordset
  Dim rsControls As DAO.Recordset
  Dim rsPageCaptions As DAO.Recordset
  Dim rsHistoryScreens As DAO.Recordset
  
  fOK = True
      
  ' Show user the system is busy...this operation could take some time...
  Screen.MousePointer = vbHourglass
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  ' Get the selected screen's definitions from the database.
  sSQL = "SELECT *" & _
    " FROM tmpScreens" & _
    " WHERE tmpScreens.screenID = " & Trim(Str(ScreenID))
  Set rsScreen = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          
  With rsScreen
    ' Create a new unique screen name.
    sScreenName = "Copy_of_" & .Fields("Name")
    mstrScreenName = sScreenName
    
    iCounter = 1
    fGoodName = False
    Do While Not fGoodName
      With recScrEdit
        .Index = "idxName"
        .Seek "=", sScreenName, False
        If Not .NoMatch Then
          iCounter = iCounter + 1
          sScreenName = "Copy_" & Trim(Str(iCounter)) & "_of_" & rsScreen.Fields("Name")
        Else
          fGoodName = True
        End If
      End With
    Loop
    
    ' Get a unique ID for the new record.
    lngScreenID = UniqueColumnValue("tmpScreens", "screenID")
          
    ' Add a new record in the database for the copied screen definition.
    recScrEdit.AddNew
          
    recScrEdit!ScreenID = lngScreenID
    recScrEdit!Changed = False
    recScrEdit!New = True
    recScrEdit!Deleted = False
    recScrEdit!Name = sScreenName
    recScrEdit!TableID = .Fields("TableID")
    recScrEdit!OrderID = .Fields("OrderID")
    recScrEdit!Height = .Fields("Height")
    recScrEdit!Width = .Fields("Width")
    recScrEdit!PictureID = .Fields("PictureID")
    recScrEdit!FontName = .Fields("FontName")
    recScrEdit!FontSize = .Fields("FontSize")
    recScrEdit!FontBold = .Fields("FontBold")
    recScrEdit!FontItalic = .Fields("FontItalic")
    recScrEdit!FontStrikethru = .Fields("FontStrikeThru")
    recScrEdit!FontUnderline = .Fields("FontUnderline")
    recScrEdit!GridX = .Fields("GridX")
    recScrEdit!GridY = .Fields("GridY")
    recScrEdit!AlignToGrid = .Fields("AlignToGrid")
    recScrEdit!dfltForeColour = .Fields("DfltForeColour")
    recScrEdit!dfltFontName = .Fields("DfltFontName")
    recScrEdit!dfltFontSize = .Fields("DfltFontSize")
    recScrEdit!dfltFontBold = .Fields("DfltFontBold")
    recScrEdit!dfltFontItalic = .Fields("DfltFontItalic")
    recScrEdit!QuickEntry = .Fields("QuickEntry")
    recScrEdit!SSIntranet = .Fields("SSIntranet")
    
    recScrEdit.Update
        
    ' Copy the screen control definitions.
    sSQL = "SELECT *" & _
      " FROM tmpControls" & _
      " WHERE tmpControls.screenID = " & Trim(Str(.Fields("ScreenID")))
    Set rsControls = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
              
    With rsControls
      ' For each screen control definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied screen control definition.
        recCtrlEdit.AddNew
              
        recCtrlEdit!ScreenID = lngScreenID
        recCtrlEdit!PageNo = .Fields("PageNo")
        recCtrlEdit!ControlLevel = .Fields("ControlLevel")
        recCtrlEdit!TableID = .Fields("TableID")
        recCtrlEdit!ColumnID = .Fields("ColumnID")
        recCtrlEdit!ControlType = .Fields("ControlType")
        recCtrlEdit!ControlIndex = .Fields("ControlIndex")
        recCtrlEdit!TopCoord = .Fields("TopCoord")
        recCtrlEdit!LeftCoord = .Fields("LeftCoord")
        recCtrlEdit!Height = .Fields("Height")
        recCtrlEdit!Width = .Fields("Width")
        recCtrlEdit!Caption = .Fields("Caption")
        recCtrlEdit!BackColor = .Fields("BackColor")
        recCtrlEdit!ForeColor = .Fields("ForeColor")
        recCtrlEdit!FontName = .Fields("FontName")
        recCtrlEdit!FontSize = .Fields("FontSize")
        recCtrlEdit!FontBold = .Fields("FontBold")
        recCtrlEdit!FontItalic = .Fields("FontItalic")
        recCtrlEdit!FontStrikethru = .Fields("FontStrikeThru")
        recCtrlEdit!FontUnderline = .Fields("FontUnderline")
        recCtrlEdit!PictureID = .Fields("PictureID")
        recCtrlEdit!DisplayType = .Fields("DisplayType")
        recCtrlEdit!TabIndex = .Fields("TabIndex")
        recCtrlEdit!BorderStyle = .Fields("BorderStyle")
        recCtrlEdit!Alignment = .Fields("Alignment")
        recCtrlEdit!ReadOnly = .Fields("ReadOnly")      'NPG20071023
        recCtrlEdit!NavigateTo = .Fields("NavigateTo")
        recCtrlEdit!NavigateIn = .Fields("NavigateIn")
        recCtrlEdit!NavigateOnSave = .Fields("NavigateOnSave")

        recCtrlEdit.Update
              
        .MoveNext
      Loop
    End With
    Set rsControls = Nothing
          
    ' Copy the screen page caption definitions.
    sSQL = "SELECT *" & _
      " FROM tmpPageCaptions" & _
      " WHERE tmpPageCaptions.screenID = " & Trim(Str(.Fields("ScreenID")))
    Set rsPageCaptions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
              
    With rsPageCaptions
      ' For each screen page caption definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied screen page caption definition.
        recPageCaptEdit.AddNew
              
        recPageCaptEdit!ScreenID = lngScreenID
        recPageCaptEdit!PageIndexID = .Fields("PageIndexID")
        recPageCaptEdit!Caption = .Fields("Caption")
            
        recPageCaptEdit.Update
              
        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsPageCaptions = Nothing
              
    ' Copy the history screen records.
    sSQL = "SELECT *" & _
      " FROM tmpHistoryScreens" & _
      " WHERE tmpHistoryScreens.parentScreenID = " & Trim(Str(.Fields("ScreenID")))
    Set rsHistoryScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
              
    With rsHistoryScreens
      ' For each history screen definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied history screen definition.
        recHistScrEdit.AddNew
              
        recHistScrEdit!ID = UniqueColumnValue("tmpHistoryScreens", "ID")
        recHistScrEdit!parentScreenID = lngScreenID
        recHistScrEdit!historyScreenID = .Fields("historyScreenID")
            
        recHistScrEdit.Update
              
        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsHistoryScreens = Nothing
                
    .Close
  End With
  ' Disassociate object variables.
  Set rsScreen = Nothing
      
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsScreen = Nothing
  Set rsControls = Nothing
  Set rsPageCaptions = Nothing
  Set rsHistoryScreens = Nothing
  
  ' Show user the system has finished working
  Screen.MousePointer = vbDefault
  
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
    RefreshScreens
    
    'Select Newly Created Screen
    SelectScreen
    
  Else
    daoWS.Rollback
    MsgBox "Unable to copy the screen." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  CopyScreen = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub SelectScreen()
  Dim iCount As Integer
  
  With ssGrdScreens
    .MoveFirst
    For iCount = 0 To .Rows
      If .Columns(0).Text = mstrScreenName Then
        .Bookmark = .AddItemBookmark(ScreenID)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
        Exit For
      End If
      .MoveNext
    Next iCount
  End With
End Sub

Private Function EditScreen() As Boolean
  ' Edit the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmScrDes As frmScrDesigner2

  fOK = (ScreenID > 0)
  
  If fOK Then
    ' Display the screen properties form.
    With frmScrEdit
      .ScreenID = ScreenID
      .Show vbModal
      fOK = Not .Cancelled
    End With
    Set frmScrEdit = Nothing
      
    If fOK Then
      ' The screen name may have been changed so refresh the
      ' screen list, and in the frmScrDeisgner caption if it is open.
      RefreshScreens
  
      ' The screen name may have been changed in the properties window,
      ' so update the caption of the frmScrDesigner form that displays this
      ' screen, if it is open.
      Set frmScrDes = GetScreenForm(ScreenID)
      If Not frmScrDes Is Nothing Then
        frmScrDes.Caption = "Screen Manager - " & ssGrdScreens.Columns(0).CellValue(ssGrdScreens.Bookmark) & vbNullString
      End If
    End If
    
    Me.SetFocus
  End If

TidyUpAndExit:
  EditScreen = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Private Function OpenScreen() As Boolean
  ' Open the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmScrDes As frmScrDesigner2
  
  fOK = (ScreenID > 0)
  
  If fOK Then
  
    'Define a progress bar
    With gobjProgress
      .Caption = "Screen Manager"
      .NumberOfBars = 1
      .Bar1Value = 1
      .Bar1MaxValue = 5
      .Bar1Caption = "Opening Screen..."
      ' NPG Fault 13329
      '.AviFile = ""
      .AVI = dbScreenLoad
      .MainCaption = "Screen Manager"
      .Cancel = False
      .Time = False
      .OpenProgress
    End With
  
    ' Display the Screen Manager form for the selected screen.
    Set frmScrDes = GetScreenForm(ScreenID)
    If frmScrDes Is Nothing Then
      Set frmScrDes = New frmScrDesigner2
      With frmScrDes
        .ScreenID = ScreenID
        .IsNew = False
        
        'Update the progress bar
        gobjProgress.UpdateProgress
       
      End With
    End If
    frmScrDes.Show
    
    'Update the progress bar
    gobjProgress.UpdateProgress

    ' Display the toolbox form.
    With frmToolbox
      Set .CurrentScreen = frmScrDes
      .Show
    End With
    
    'Update the progress bar
    gobjProgress.UpdateProgress
    
    ' Display the screen object properties form.
    With frmScrObjProps
      Set .CurrentScreen = frmScrDes
      .Show
    End With

    'Update the progress bar
    gobjProgress.UpdateProgress
    
    ' Dodgy fix to avoid locking the dodgy toolbar.
'    With frmSysMgr.tbMain
'      .Redraw = False
'      .Enabled = False
'      .Enabled = True
'      .Redraw = True
'    End With
    
    UnLoad Me
    
    ' RH 06/04/2000. Fault 44. Screen Manager prompts when no changes have been made
    DoEvents
    frmScrDes.IsChanged = False
    
  End If

TidyUpAndExit:

  'Close the progress bar
  gobjProgress.CloseProgress
  
  OpenScreen = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Private Function DeleteScreen() As Boolean
  ' Delete the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As DAO.Recordset
  Dim sModuleName As String
  Dim frmUse As frmUsage
  
  fOK = (ScreenID > 0)
  
  If fOK Then
    fOK = GetScreenForm(ScreenID) Is Nothing
    
    If Not fOK Then
      MsgBox "The '" & ScreenName & "' screen cannot be deleted." & vbCr & "It is already open.", _
        vbExclamation + vbOKOnly, Me.Caption
    Else
      Set frmUse = New frmUsage
      frmUse.ResetList

      ' Check if the screen is used.
      sSQL = "SELECT COUNT(*) AS recCount" & _
        " FROM tmpSSIntranetLinks" & _
        " WHERE screenID = " & Trim(CStr(ScreenID))
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If rsInfo!reccount > 0 Then
        'MsgBox "The '" & ScreenName & "' screen cannot be deleted." & vbCr & "It is used in the Self-service Intranet module.", _
          vbExclamation + vbOKOnly, Me.Caption
        frmUse.AddToList ("Self-service Intranet module")
        fOK = False
      End If
      rsInfo.Close
      Set rsInfo = Nothing

      'AE20071127 Fault #12645
      ' Check that the screen is not used in any Module definitions.
      sSQL = "SELECT DISTINCT moduleKey, parameterkey" & _
        " FROM tmpModuleSetup" & _
        " WHERE parameterType = '" & gsPARAMETERTYPE_SCREENID & "'" & _
        " AND parameterValue = '" & Trim(Str(ScreenID)) & "'"
        
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If Not (rsInfo.BOF And rsInfo.EOF) Then
        fOK = False
        Do Until rsInfo.EOF
          Select Case rsInfo!moduleKey
            Case gsMODULEKEY_TRAININGBOOKING
              sModuleName = "Training Booking"
            Case gsMODULEKEY_PERSONNEL
              sModuleName = "Personnel"
            Case gsMODULEKEY_ABSENCE
              sModuleName = "Absence"
            Case gsMODULEKEY_CURRENCY
              sModuleName = "Currency"
            Case gsMODULEKEY_POST
              sModuleName = "Post"
            Case gsMODULEKEY_MATERNITY
              sModuleName = "Maternity"
            Case gsMODULEKEY_SSINTRANET
              sModuleName = "Self Service Intranet"
            Case gsMODULEKEY_HIERARCHY
              sModuleName = "Hierachy"
            Case Else
              sModuleName = "<Unknown>"
          End Select
          frmUse.AddToList (sModuleName & " Configuration")
          rsInfo.MoveNext
        Loop
      End If
      ' Close the recordset.
      rsInfo.Close
      Set rsInfo = Nothing
    End If
    
    If Not fOK Then
      frmUse.ShowMessage ScreenName & " Screen", "The '" & ScreenName & "' screen cannot be deleted as it is used in the following:", UsageCheckObject.Form
    Else
      If MsgBox("Delete screen '" & ScreenName & "', are you sure?", _
        vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then

        fOK = DeleteScreen_Transaction(ScreenID)
        
        If fOK Then
          RefreshScreens
        End If
      End If
    End If
    
    UnLoad frmUse
    Set frmUse = Nothing
  End If

TidyUpAndExit:
  DeleteScreen = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function DeleteScreen_Transaction(plngScreenID As Long) As Boolean
  ' Delete the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  ' Delete the screen's controls from the local database.
  daoDb.Execute "DELETE FROM tmpControls WHERE screenID=" & plngScreenID
  
  ' Delete the screen's history screens from the local database.
  daoDb.Execute "DELETE FROM tmpHistoryScreens" & _
    " WHERE parentScreenID=" & plngScreenID & _
    " OR historyScreenID=" & plngScreenID
        
  ' Delete the screen's page captions from the local database.
  daoDb.Execute "DELETE FROM tmpPageCaptions WHERE tmpPageCaptions.screenID = " & plngScreenID
       
  ' Delete the screen's views from the local database.
  daoDb.Execute "UPDATE tmpViewScreens SET deleted = TRUE WHERE tmpViewScreens.screenID = " & plngScreenID
        
  ' Mark the screen as deleted in the local database.
  With recScrEdit
    .Index = "idxScreenID"
    .Seek "=", plngScreenID
    If Not .NoMatch Then
      .Edit
      
      ' ERROR IS HERE...FOR SOME REASON IT CAUSES A DUP INDEX VALUE SOMEWHERE ?!?
      ' WHEN DELETING A COPY OF A SCREEN - OCCASIONALLY !
      
      .Fields("Deleted") = True
      .Update
    End If
  End With
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteScreen_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox "Error deleting screen..." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Function

Private Function NewScreen() As Boolean
  ' Create a new screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngScreenID As Long
  Dim frmScrDes As frmScrDesigner2
  
  ' Display the screen properties form.
  With frmScrEdit
    .TableID = Me.cboTable.ItemData(Me.cboTable.ListIndex)
    .ScreenID = 0
    .LockedTableId = Me.cboTable.ItemData(Me.cboTable.ListIndex)
    .Show vbModal
    fOK = Not .Cancelled
    lngScreenID = .ScreenID
  End With
  Set frmScrEdit = Nothing
      
  'Define a progress bar
  With gobjProgress
    .Caption = "Screen Manager"
    .AVI = dbScreenLoad
    .NumberOfBars = 1
    .Bar1Value = 1
    .Bar1MaxValue = 4
    .Bar1Caption = "Creating Screen..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
      
  ' If the screen properties were confirmed then display the Screen Manager form.
  If fOK Then
    If lngScreenID > 0 Then
    
      Set frmScrDes = New frmScrDesigner2
      With frmScrDes
        .ScreenID = lngScreenID
        .IsNew = True
        .Show
      End With
          
      'Update the progress bar
      gobjProgress.UpdateProgress
          
      ' Display the toolbox form.
      With frmToolbox
        Set .CurrentScreen = frmScrDes
        .Show
      End With
      
      'Update the progress bar
      gobjProgress.UpdateProgress
      
      ' Display the screen object properties form.
      With frmScrObjProps
        Set .CurrentScreen = frmScrDes
        .Show
      End With

      'Update the progress bar
      gobjProgress.UpdateProgress

    End If
    
    UnLoad Me
    
  End If

TidyUpAndExit:

  'Close the progress bar
  gobjProgress.CloseProgress

  NewScreen = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub LoadTableCombo()
  ' Populate the grid with screens.
  Dim iSelectedScreen As Integer
  Dim lngScreenID As Long
  Dim lngColumnWidth As Long
  Dim sSQL As String
  Dim frmScreen As frmScrDesigner2
  Dim rsScreens As DAO.Recordset
  
  lngScreenID = ScreenID

  ' Get the list of ALL tables from the database.
  sSQL = "SELECT DISTINCT tmpTables.tableName, tmpTables.tableID" & _
    " FROM tmptables" & _
    " WHERE tmptables.deleted = FALSE" & _
    " ORDER BY tmptables.tablename"
  Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
'  If rsScreens.BOF And rsScreens.EOF Then
'    frmScrDesigner2.ScreenID = 0
'  End If
  
  ' If we already have the screen manager open then we can only allow access to
  ' the same primary database, and its associated screens.
  gLngLockedTableID = 0
  
  Do While Not rsScreens.EOF
    
    Set frmScreen = GetScreenForm(rsScreens!TableID)
    
    If Not frmScreen Is Nothing Then
      
      ' Remember which table we are locked to.
      gLngLockedTableID = rsScreens!TableID

      ' Get the updated list of screens from the database.
      sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
        " FROM tmptables" & _
        " WHERE tmptables.tableID = " & rsScreens!TableID & _
        " ORDER BY tmptables.tablename"
    
      Exit Do
    
    End If
  
    rsScreens.MoveNext
  Loop
  
  rsScreens.Close
  
  ' Get the list of VALID screens from the database.
  Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  Me.cboTable.Clear
  
  Do While Not rsScreens.EOF
    cboTable.AddItem rsScreens!TableName
    cboTable.ItemData(cboTable.NewIndex) = rsScreens!TableID
    rsScreens.MoveNext
  Loop
  
  rsScreens.Close
  Set rsScreens = Nothing
  
  If cboTable.ListCount > 0 Then
    If mstrCurrentTable <> "" Then
      SetComboText cboTable, mstrCurrentTable
    Else
      cboTable.ListIndex = 0
    End If
  End If
  
  ' Update the status bar.
  sbScrOpen.Panels(1).Text = ssGrdScreens.Rows & " screen" & _
    IIf(ssGrdScreens.Rows = 1, vbNullString, "s")
  
  ' Refresh the menu
  frmSysMgr.RefreshMenu

End Sub

' Set the combo text specified, if not found set to top
Public Sub SetComboText(cboCombo As ComboBox, sText As String)

    Dim lCount As Long
    Dim bFound As Boolean
    
    bFound = False
    
    With cboCombo
      For lCount = 1 To .ListCount
        If .List(lCount - 1) = sText Then
          .ListIndex = lCount - 1
          bFound = True
          Exit For
        End If
      Next
      
      If Not bFound Then
        .ListIndex = 0
      End If
        
    End With

End Sub

