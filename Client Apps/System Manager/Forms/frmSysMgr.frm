VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "Codejock.SkinFramework.v13.1.0.ocx"
Begin VB.MDIForm frmSysMgr 
   AutoShowChildren=   0   'False
   BackColor       =   &H00F7EEE9&
   Caption         =   "OpenHR - System Manager"
   ClientHeight    =   6915
   ClientLeft      =   2370
   ClientTop       =   2130
   ClientWidth     =   7530
   HelpContextID   =   5066
   Icon            =   "frmSysMgr.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "frmSysMgr.frx":038A
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrKeepAlive 
      Interval        =   6000
      Left            =   4440
      Top             =   3960
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Top             =   6624
      Width           =   7524
      _ExtentX        =   13282
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7530
      TabIndex        =   1
      Top             =   0
      Width           =   7524
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2760
         ScaleHeight     =   285
         ScaleWidth      =   870
         TabIndex        =   2
         Top             =   120
         Width           =   900
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   720
      Top             =   1560
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin ActiveBarLibraryCtl.ActiveBar tbMain 
      Left            =   540
      Top             =   2115
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
      Bands           =   "frmSysMgr.frx":0714
   End
End
Attribute VB_Name = "frmSysMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmDbMgr As SystemMgr.frmDbMgr
Attribute frmDbMgr.VB_VarHelpID = -1
Private WithEvents frmPictMgr As SystemMgr.frmPictMgr
Attribute frmPictMgr.VB_VarHelpID = -1
' JPD 19/3/98 Need to refresh the menu when the Screen Open
' and Manager forms are closed.
Public WithEvents frmScrOpen As SystemMgr.frmScrOpen
Attribute frmScrOpen.VB_VarHelpID = -1
Public WithEvents frmWorkflowOpen As SystemMgr.frmWorkflowOpen
Attribute frmWorkflowOpen.VB_VarHelpID = -1
'RJB 06/08/98 Added View manager
Private WithEvents frmViewMgr As SystemMgr.frmViewMgr
Attribute frmViewMgr.VB_VarHelpID = -1

' Functions to display/tile the background image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal lDC As Long, ByVal hObject As Long) As Long

Public Sub SetBackground(ByRef mbIsLoading As Boolean)

  Dim X, Y, hMemDC, pHeight, pWidth As Long
  Dim pic As StdPicture
  Dim lngPictureID As Long
  Dim sFileName As String

  ' Define the background colour
  Me.BackColor = glngDeskTopColour

  'Place defaults in the images controls
  picHolder.Visible = False
  picWork.Visible = False
  picWork.AutoRedraw = True
  picWork.AutoSize = True
  picWork.BorderStyle = 0

  ' Load the desired picture from the database
  If glngDesktopBitmapID > 0 Then
    recPictEdit.Index = "idxID"
    recPictEdit.Seek "=", glngDesktopBitmapID
    If Not recPictEdit.NoMatch Then
      sFileName = ReadPicture
      picWork.Picture = LoadPicture(sFileName)
      Kill sFileName
    End If
  End If

  'Variables used to set the background tiled image
  Me.Visible = True
  pHeight = picWork.Height
  pWidth = picWork.Width
  Set pic = picWork.Picture
  Set picWork.Picture = Nothing
  picWork.AutoSize = False
  
  hMemDC = CreateCompatibleDC(picWork.hDC)
  SelectObject hMemDC, pic.Handle

  If WindowState <> vbMinimized Then
    picWork.BackColor = Me.BackColor
    picWork.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

    If pWidth > 0 Then

      ' Tiled backdrop
      If glngDesktopBitmapLocation = giLOCATION_TILE Then
        For X = 0 To Me.ScaleWidth Step pWidth
          For Y = 0 To Me.ScaleHeight Step pHeight
            BitBlt picWork.hDC, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
          Next
        Next
      End If

      ' Tiled down the lefthand side
      If glngDesktopBitmapLocation = giLOCATION_LEFTTILE Then
        For Y = 0 To Me.ScaleHeight Step pHeight
          BitBlt picWork.hDC, 0, Y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      ' Tiled down the righthand side
      If glngDesktopBitmapLocation = giLOCATION_RIGHTTILE Then
        For Y = 0 To Me.ScaleHeight Step pHeight
          BitBlt picWork.hDC, (Me.ScaleWidth - pWidth) \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      ' Top left hand side
      If glngDesktopBitmapLocation = giLOCATION_TOPLEFT Then
        BitBlt picWork.hDC, 0, 0, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
      End If

      ' Top right hand side
      If glngDesktopBitmapLocation = giLOCATION_TOPRIGHT Then
        BitBlt picWork.hDC, (Me.ScaleWidth - pWidth) \ Screen.TwipsPerPixelX, 0, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
      End If

      ' Centred on the backdrop
      If glngDesktopBitmapLocation = giLOCATION_CENTRE Then
        X = (ScaleWidth - pWidth) \ 2: X = X \ Screen.TwipsPerPixelX
        Y = (ScaleHeight - pHeight) \ 2: Y = Y \ Screen.TwipsPerPixelY
        BitBlt picWork.hDC, X, Y, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
      End If

      ' Tiled across the top
      If glngDesktopBitmapLocation = giLOCATION_TOPTILE Then
        For X = 0 To Me.ScaleWidth Step pWidth
          BitBlt picWork.hDC, X \ Screen.TwipsPerPixelX, 0, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      'Tiled across the bottom
      If glngDesktopBitmapLocation = giLOCATION_BOTTOMTILE Then
        For X = 0 To Me.ScaleWidth Step pWidth
          BitBlt picWork.hDC, X \ Screen.TwipsPerPixelX, (Me.ScaleHeight - pHeight) \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      Set Picture = picWork.Image
  
    End If
  End If

  ' We have to re-hide the form if we are stiil loading or is messes up the activate routine
  If mbIsLoading = True Then
    Me.Visible = False
  End If

End Sub


Private Sub frmDbMgr_Activate()
  frmSysMgr.RefreshMenu
End Sub

Private Sub frmDbMgr_Deactivate()
  frmSysMgr.RefreshMenu
End Sub

Private Sub frmDbMgr_UnLoad()
  frmSysMgr.RefreshMenu True
End Sub

Private Sub frmPictMgr_Activate()
  RefreshMenu
End Sub

Private Sub frmPictMgr_Deactivate()
  RefreshMenu
End Sub

Private Sub frmPictMgr_UnLoad()
  RefreshMenu True
End Sub

Private Sub frmScrOpen_Unload()
  RefreshMenu True
End Sub


Private Sub frmWorkflowOpen_Unload()
  RefreshMenu True
End Sub

Private Sub frmViewMgr_Activate()
  RefreshMenu
End Sub

Private Sub frmViewMgr_Deactivate()
  RefreshMenu
End Sub

Private Sub frmViewMgr_UnLoad()
  RefreshMenu True
End Sub

Private Sub MDIForm_Activate()

  ' Refresh the menu display.
  frmSysMgr.RefreshMenu
  
  ' Set the new multi-size icons for taskbar, application, and alt-tab
  SetIcon Me.hWnd, "!ABS", True
 
  
End Sub

Private Sub MDIForm_Load()
  Dim objPrinter As Printer

  ' Load the CodeJock Styles
  Call LoadSkin(Me, Me.SkinFramework1)

  With tbMain
    .MenuFontStyle = 1
    .Font.Name = "Verdana"
    .Font.Bold = False
    .Font.Size = 8

    .ControlFont.Name = "Verdana"
    .ControlFont.Bold = False
    .ControlFont.Size = 8

    .ForeColor = 6697779
    .BackColor = 16248553

    .Refresh
  End With
        
  SetCaption
  StatusBar1.SimpleText = gsDatabaseName & _
    " - Version: " & gstrSQLFullVersion

  '# RH 27/07
  '# If we dont want the user to be able to customise the tools/menus then use:
  'tbMain.DisplayContextMenu = False
 
'JPD 20081205 - You can have Printers.Count > 0 but still no valid printers (honestly!)
' So need to have proper error trapping, on top of the Printers.Count check.
On Error GoTo PrinterErrorTrap
  
  If Printers.Count > 0 Then
    Printer.TrackDefault = True
    gstrDefaultPrinterName = Printer.DeviceName
    
    SavePCSetting "Printer", "DeviceName", gstrDefaultPrinterName
    
  '  gstrDefaultPrinterName = GetPCSetting( "Printer", "DeviceName", "")
    For Each objPrinter In Printers
      If LCase(objPrinter.DeviceName) = LCase(gstrDefaultPrinterName) Then
        Set Printer = objPrinter
      End If
    Next
  End If
  
  Exit Sub

PrinterErrorTrap:

End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If ASRDEVELOPMENT And Button = 1 And Shift = 1 Then
    Application.Changed = True
    RefreshMenu
  End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim intFormNo As Integer
  Dim frmPrompt As frmSaveChangesPrompt
  
  
  'MH20071101 On Vista sometimes the close button remains enabled
  '           so stop the form unloading if we are saving changes.
  If gobjProgress.Visible Then
    Cancel = True
    Exit Sub
  End If


  intFormNo = Forms.Count - 1
  Do While intFormNo > 0 And Forms.Count > 1
    If Forms(intFormNo).Name <> "frmScrObjProps" _
      And Forms(intFormNo).Name <> "frmToolbox" _
      And Forms(intFormNo).Name <> "frmWorkflowWFItemProps" _
      And Forms(intFormNo).Name <> "frmWorkflowWFToolbox" Then
      
      UnLoad Forms(intFormNo)
      DoEvents
    End If
    intFormNo = intFormNo - 1
  Loop
  
  If Forms.Count > 1 Then
    Cancel = True
  ElseIf Application.Changed Then
    Set frmPrompt = New frmSaveChangesPrompt
    frmPrompt.Buttons = vbYesNoCancel
    frmPrompt.Show vbModal
    Select Case frmPrompt.Choice
      Case vbCancel
        Cancel = True
        RefreshMenu
      Case vbYes
        Cancel = Not (SaveChanges(frmPrompt.RefreshDatabase))
    End Select
    Set frmPrompt = Nothing
  End If

  'Remove Progress Bar class from memory
  If Cancel = False Then Set gobjProgress = Nothing

End Sub

Private Sub MDIForm_Resize()
  'JPD 20030908 Fault 5756
  If Me.WindowState <> vbMinimized Then
    giWindowState = Me.WindowState

'    If Me.Height < 2000 Then Me.Height = 2000

'    If Me.WindowState = vbNormal Then
'      glngWindowLeft = IIf(Me.Left < (0 - Me.Width), glngWindowLeft, Me.Left)
'      glngWindowTop = IIf(Me.Top < (0 - Me.Height), glngWindowTop, Me.Top)
'      glngWindowHeight = IIf((Me.Top < (0 - Me.Height)) Or (Me.ScaleHeight <= 0), glngWindowHeight, Me.Height)
'      glngWindowWidth = IIf((Me.Left < (0 - Me.Width)) Or (Me.ScaleWidth <= 0), glngWindowWidth, Me.Width)
'    End If
  End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  'JPD 20061130 Fault 11531
  If Not frmWorkflowOpen Is Nothing Then
    UnLoad frmWorkflowOpen
    Set frmWorkflowOpen = Nothing
  End If
  If Not frmScrOpen Is Nothing Then
    UnLoad frmScrOpen
    Set frmScrOpen = Nothing
  End If
  
'Logout
  Logout
  
End Sub

Public Sub RefreshMenu(Optional ByVal pfUnLoad As Boolean)
  
  '# RH 1/9/99 ALL REFRESH CODE HAS BEEN AMENDED SO THE TOOLBARS ON THE
  '#           MDI CHILDREN ARE UPDATED. THE CODE WHICH UPDATES THE TOOLBARS
  '#           ON THE MAIN MDI FORM IS STILL HERE AND STILL RUNS. THIS IS TO
  '#           ALLOW EASY ROLLBACK OF USING MDI FORM TOOLBAR ONLY.
  '#           EVENTUALLY, THE MAIN MDI FORM TOOLBAR AND ITS CODE CAN BE
  '#           DELETED.
  '#
  '#           THE CODE WHICH UPDATED THE MDI CHILD TOOLBARS CAN BE FOUND AT
  '#           THE BEGINNING OF THE INDIVIDUAL REFRESHMENU SUBS, IE.
  '#           REFRESHMENU_DBMGR, REFRESHMENU_SCRMGR ETC.
  '#           THE CODE WHICH UPDATES THE MAIN MDI FORM TOOLBAR FOLLOWS.
  
  Dim iFormCount As Integer
  Dim objTool As ActiveBarLibraryCtl.Tool
  Dim objGroup As ActiveBarLibraryCtl.Tool
  
  ' Get the number of forms that are loaded.
  iFormCount = Forms.Count - IIf(pfUnLoad, 1, 0)
  
'  ClearMenuShortcuts

  ' With the ToolBar control ...
  With tbMain

    ' Disable the redrawing of the toolbar until all changes have been made.
'    .Redraw = False

    ' Make all menus and tools invisible.
    For Each objTool In .Tools
      objTool.Visible = False
    Next
    Set objTool = Nothing

    ' Configure the default menu tools.
    RefreshMenu_Defaults (iFormCount)
    
    If Not (Me.ActiveForm Is Nothing Or iFormCount <= 1) Then

      If TypeOf Me.ActiveForm Is frmDbMgr Then
        RefreshMenu_DBMgr (iFormCount)
      
      ElseIf TypeOf Me.ActiveForm Is frmPictMgr Then
        RefreshMenu_PictMgr (iFormCount)
      
      ElseIf TypeOf Me.ActiveForm Is frmScrOpen Then
        RefreshMenu_ScrMgr (iFormCount)
      
      ElseIf TypeOf Me.ActiveForm Is frmScrDesigner2 Then
        RefreshMenu_ScrDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmScrObjProps Then
        RefreshMenu_ScrDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmToolbox Then
        RefreshMenu_ScrDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowOpen Then
        RefreshMenu_WorkflowMgr (iFormCount)
      
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowDesigner Then
        RefreshMenu_WorkflowMgr (iFormCount)
      
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFDesigner Then
        RefreshMenu_WebFormDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFItemProps Then
        RefreshMenu_WebFormDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFToolbox Then
        RefreshMenu_WebFormDesigner iFormCount
      
      ElseIf TypeOf Me.ActiveForm Is frmViewMgr Then
        RefreshMenu_ViewMgr (iFormCount)
      
      End If
    
    End If
    
    ' Refresh the toolbar control with the changes we've just made.
'    .Redraw = True
    .RecalcLayout

  End With
  
  EnsureSaveButtonsCorrect
    
  CheckForDisabledMenuItems

End Sub

Private Sub EnsureSaveButtonsCorrect()

  Dim objForm As Form
  
  
  'MH20010214 Had a weird run-time error in this sub
  'so just added a bit of error trapping
  On Local Error GoTo LocalErr
  
  For Each objForm In Forms
    If TypeOf objForm Is frmViewMgr Then
      objForm.abViewMgr.Tools("ID_SaveChanges").Enabled = Application.Changed ' tbMain.Tools("ID_SaveChanges").Enabled
    ElseIf TypeOf objForm Is frmPictMgr Then
      objForm.abPictMgrMenu.Tools("ID_SaveChanges").Enabled = Application.Changed ' tbMain.Tools("ID_SaveChanges").Enabled
    ElseIf TypeOf objForm Is frmDbMgr Then
      objForm.abDbMgr.Tools("ID_SaveChanges").Enabled = Application.Changed ' tbMain.Tools("ID_SaveChanges").Enabled
    End If
  Next objForm
  
  Set objForm = Nothing
  
Exit Sub

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
  End If

End Sub


Private Sub CascadeForms()
  Dim lngOffset As Long
  Dim intCount As Integer
  Dim intForm As Integer
  Dim frmActive As Form
  
  lngOffset = (UI.GetSystemMetrics(SM_CYFRAME) + _
    UI.GetSystemMetrics(SM_CYCAPTION)) * Screen.TwipsPerPixelY
  
  Set frmActive = Me.ActiveForm
  intCount = 0
  For intForm = 1 To Forms.Count - 1
    If Forms(intForm).WindowState <> vbMinimized Then
      intCount = intCount + 1
    End If
  Next intForm
  
  If Not frmActive Is Nothing Then
    With frmActive
      If .WindowState <> vbMinimized Then
        intCount = intCount - 1
        .WindowState = vbNormal
        .Top = lngOffset * intCount
        .Left = lngOffset * intCount
        .ZOrder 0
      End If
    End With
  End If
  
  For intForm = Forms.Count - 1 To 1 Step -1
    If Not Forms(intForm) Is frmActive Then
      With Forms(intForm)
        If .WindowState <> vbMinimized Then
          intCount = intCount - 1
          .WindowState = vbNormal
          .Top = lngOffset * intCount - 1
          .Left = lngOffset * intCount - 1
          .ZOrder 1
        End If
      End With
    End If
  Next intForm
  
  Set frmActive = Nothing
End Sub

Private Function ChangeView(ViewStyle As ComctlLib.ListViewConstants) As Boolean
  
  Static InChangeView As Boolean
  Dim objListView As ListView
  
  On Error GoTo ErrorTrap
  
  If InChangeView Then Exit Function
  
  InChangeView = True
  
  Set objListView = Me.ActiveForm.Controls("ListView1")
  
  If ViewStyle = lvwSmallIcon Then
    objListView.View = lvwList
  End If
  objListView.View = ViewStyle
  objListView.Visible = True
  
  ' As ActiveBar does not support mutual exclusivity on its tools, the following code ensures only one of the view options is selected at any one time.
  With tbMain
    Select Case ViewStyle
      Case lvwIcon
        .Tools("ID_LargeIcons").Checked = True
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
      Case lvwSmallIcon
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = True
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
      Case lvwList
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = True
        .Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
      Case lvwReport
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = True
        .Tools("ID_CustomiseColumns").Enabled = True
    End Select
  End With
  
  InChangeView = False
  
  ChangeView = True
  
  CheckForDisabledMenuItems

  Exit Function
  
ErrorTrap:
  ChangeView = False
  Err = False

End Function




Private Sub RefreshMenu_DBMgr(piFormCount As Integer)
  
  '# refresh the toolbar on frmDbMgr first
  
  Dim blnReadonly As Boolean
  Dim bCopyTable As Boolean

  blnReadonly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)
  
  With frmDbMgr.abDbMgr
    
    .Tools("ID_New").Enabled = False
    .Tools("ID_CopyDef").Enabled = False
    .Tools("ID_Delete").Enabled = False
    .Tools("ID_Properties").Enabled = False
    .Tools("ID_Print").Enabled = False
    
    If Not blnReadonly Then
    
      If frmDbMgr.ActiveView Is frmDbMgr.TreeView1 Then
        
        bCopyTable = DoesTableExistInDB(val(Mid(frmDbMgr.TreeView1.SelectedItem.key, 2)))
        
        If frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_RELATIONGROUP Or frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_RELATION Then
          .Tools("ID_New").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtAdd) And (frmDbMgr.TreeView1.Nodes("TABLES").Children > 0)
          .Tools("ID_CopyDef").Enabled = False
        Else
          .Tools("ID_New").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtAdd)
          .Tools("ID_CopyDef").Enabled = bCopyTable
        End If
        
        
        
        .Tools("ID_Delete").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtDelete)
        .Tools("ID_Properties").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtEdit)
        .Tools("ID_Print").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtCopy)
      Else
        .Tools("ID_New").Enabled = (frmDbMgr.ListView1_SelectedTag And edtAdd)
        
        If frmDbMgr.ListView1.ListItems.Count > 0 And frmDbMgr.ListView1_SelectedCount > 0 Then
          .Tools("ID_Delete").Enabled = (frmDbMgr.ListView1_SelectedTag And edtDelete)
          
          If frmDbMgr.ListView1_SelectedCount = 1 Then
            bCopyTable = DoesTableExistInDB(val(Mid(frmDbMgr.ListView1.SelectedItem.key, 2))) And frmDbMgr.ListView1_SelectedTag = giNODE_TABLE
            .Tools("ID_CopyDef").Enabled = bCopyTable Or (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN Or frmDbMgr.ListView1_SelectedTag = giNODE_TABLE)
            .Tools("ID_Properties").Enabled = (frmDbMgr.ListView1.SelectedItem.Tag And edtEdit)
            .Tools("ID_Print").Enabled = (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN Or frmDbMgr.ListView1_SelectedTag = giNODE_TABLE)
          End If
        Else
        
        End If
      End If
    
    End If
      
    .Tools("ID_LargeIcons").Enabled = True
    .Tools("ID_SmallIcons").Enabled = True
    .Tools("ID_List").Enabled = True
    .Tools("ID_Details").Enabled = ((frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLEGROUP) Or (frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLE))
    .Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Enabled And _
      (frmDbMgr.ListView1.View = lvwReport)
    frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Enabled And _
      (frmDbMgr.ListView1.View = lvwReport)
    
    'NHRD10062003 Fault 5822
    If .Tools("ID_Details").Enabled = (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN) Then
      'ChangeView lvwReport
'      .Tools("ID_LargeIcons").Checked = False
'      .Tools("ID_SmallIcons").Checked = False
'      .Tools("ID_List").Checked = False
'      .Tools("ID_Details").Checked = True
      .Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Checked And _
        ((frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLEGROUP) Or (frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLE))

      frmSysMgr.tbMain.Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Checked
    End If
    
    
    
    If (frmDbMgr.ListView1.View = lvwReport) And _
      (Not .Tools("ID_Details").Enabled) Then
      frmDbMgr.ListView1.View = lvwList
      'TM20010917 Fault 2038 & 'TM20010917 Fault 2821
      'Must disable check the list button and uncheck the details button.
      .Tools("ID_List").Checked = True
      .Tools("ID_Details").Checked = False
    End If
    'ChangeView frmDbMgr.ListView1.View
    'ChangeView lvwReport
     ' Display the menu and it's required tools.
'    .Tools("ID_LargeIcons").Visible = True
'    .Tools("ID_SmallIcons").Visible = True
'    .Tools("ID_List").Visible = True
'    .Tools("ID_Details").Visible = True

'  .Redraw = True
  
  End With
  
  ' Refresh the menu (visible) and main toolbar (not visible) control
  
  
  With tbMain
'  .Redraw = False
    '==================================================
    ' Configure the Edit menu.
    '==================================================
    ' Enable/disable the required tools.
    If frmDbMgr.ActiveView Is frmDbMgr.TreeView1 Then
      If frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_RELATIONGROUP Or frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_RELATION Then
        .Tools("ID_New").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtAdd) And (frmDbMgr.TreeView1.Nodes("TABLES").Children > 0) And Not blnReadonly
        .Tools("ID_CopyDef").Enabled = False
        .Tools("ID_CopyDef").Visible = True
      Else
        .Tools("ID_New").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtAdd) And Not blnReadonly
        .Tools("ID_CopyDef").Enabled = bCopyTable And Not blnReadonly
        .Tools("ID_CopyDef").Visible = True
      End If
      .Tools("ID_Delete").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtDelete) And Not blnReadonly
      .Tools("ID_Properties").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag And edtEdit)
      '.Tools("ID_CopyTable").Enabled = bCopyTable
      '.Tools("ID_CopyColumn").Enabled = False
      '.Tools("ID_CopyTable").Visible = True
      '.Tools("ID_CopyColumn").Visible = True
'      .Tools("ID_CopyTable").Visible = (frmDbMgr.TreeView1.SelectedItem.Tag And edtCopy)
'      .Tools("ID_CopyColumn").Visible = False
    Else
      .Tools("ID_New").Enabled = (frmDbMgr.ListView1_SelectedTag And edtAdd) And Not blnReadonly
      If frmDbMgr.ListView1.ListItems.Count > 0 And frmDbMgr.ListView1_SelectedCount > 0 Then
        .Tools("ID_Delete").Enabled = (frmDbMgr.ListView1_SelectedTag And edtDelete) And Not blnReadonly
        If frmDbMgr.ListView1_SelectedCount > 1 Then
          .Tools("ID_Properties").Enabled = False
          '.Tools("ID_CopyTable").Visible = True
          '.Tools("ID_CopyColumn").Visible = True
          '.Tools("ID_CopyTable").Enabled = False
          '.Tools("ID_CopyColumn").Enabled = False
          .Tools("ID_CopyDef").Enabled = False
        Else
          .Tools("ID_Properties").Enabled = (frmDbMgr.ListView1.SelectedItem.Tag And edtEdit)
          '.Tools("ID_CopyTable").Enabled = bCopyTable
'          .Tools("ID_CopyTable").Visible = (frmDbMgr.ListView1.SelectedItem.Tag And edtCopy)

          '.Tools("ID_CopyTable").Visible = True
          '.Tools("ID_CopyColumn").Visible = True

          '.Tools("ID_CopyColumn").Enabled = Not .Tools("ID_CopyTable").Enabled And Not blnReadonly And (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN)
'          .Tools("ID_CopyColumn").Visible = Not .Tools("ID_CopyTable").Visible And (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN)
          .Tools("ID_CopyDef").Enabled = bCopyTable Or (frmDbMgr.ListView1_SelectedTag = giNODE_TABLE) Or (frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN) And Not blnReadonly
        End If
      Else
        .Tools("ID_Delete").Enabled = False
        .Tools("ID_Properties").Enabled = False
        
        '.Tools("ID_CopyTable").Visible = True
        '.Tools("ID_CopyColumn").Visible = True
        .Tools("ID_CopyDef").Enabled = False

      End If
    End If
    .Tools("ID_SelectAll").Enabled = (frmDbMgr.TreeView1.SelectedItem.Tag <> 0) And frmDbMgr.ListView1.ListItems.Count And Not blnReadonly
      
    ' Reassign shortcuts if required.
'    .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'    .Tools("ID_Delete").Shortcut = ssDel
'    .Tools("ID_ScreenSelectAll").Shortcut = ssShortcutNone
'    .Tools("ID_SelectAll").Shortcut = ssCtrlA
'    .Tools("ID_ScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_Properties").Shortcut = ssF4
    
    ' Add the required separators to the menu and tool bars.
'    .Tools("ID_mnuEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuEdit").Menu.Tools("ID_Properties").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_New").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Properties").Index
      
    ' Display the menu and it's required tools.
    .Tools("ID_mnuEdit").Visible = True
    .Tools("ID_New").Visible = True
    .Tools("ID_CopyDef").Visible = True
    .Tools("ID_Delete").Visible = True
    .Tools("ID_SelectAll").Visible = True
    .Tools("ID_Properties").Visible = True
    
    .Tools("ID_PrintDef").Enabled = frmDbMgr.abDbMgr.Tools("ID_Print").Enabled
    .Tools("ID_PrintDef").Visible = True
    .Tools("ID_Print").Visible = True
    .Tools("ID_CopyClipboard").Visible = True
    
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True
    
    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)


    '==================================================
    ' Configure the View menu.
    '==================================================
    ' Enable/disable the required tools.
    .Tools("ID_LargeIcons").Enabled = True
    .Tools("ID_SmallIcons").Enabled = True
    .Tools("ID_List").Enabled = True
    .Tools("ID_Details").Enabled = ((frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLEGROUP) Or (frmDbMgr.TreeView1.SelectedItem.Tag = giNODE_TABLE))
 '(frmDbMgr.ListView1_SelectedTag = giNODE_COLUMN)
    
    If (frmDbMgr.ListView1.View = lvwReport) And _
      (Not .Tools("ID_Details").Enabled) Then
      ''frmDbMgr.ListView1.View = lvwList
      'TM20010917 Fault 2038 & 'TM20010917 Fault 2821
      'Must disable check the list button and uncheck the details button.
      .Tools("ID_List").Checked = True
      .Tools("ID_Details").Checked = False
    End If
    ChangeView frmDbMgr.ListView1.View
    
    ' Add the required separators to the menu and tool bars.
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_LargeIcons").Index
      
    ' Display the menu and it's required tools.
    .Tools("ID_mnuView").Visible = True
    .Tools("ID_LargeIcons").Visible = True
    .Tools("ID_SmallIcons").Visible = True
    .Tools("ID_List").Visible = True
    .Tools("ID_Details").Visible = True
    .Tools("ID_CustomiseColumns").Visible = True
    
     
    '==================================================
    ' Configure the Window menu.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
    
    ' Add the required separators to the menu and tool bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index
      
    ' Display the menu and it's required tools.
    .Tools("ID_mnuWindow").Visible = True
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True

'  .Redraw = True
  End With
  
  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_PictMgr(piFormCount As Integer)
  
  ' With the Toolbar Ctl on frmPictMgr
  With frmPictMgr.abPictMgrMenu
    If frmPictMgr.ListView1.ListItems.Count > 0 And frmPictMgr.ListView1_SelectedCount > 0 Then
      .Tools("ID_Delete").Enabled = True
      If frmPictMgr.ListView1_SelectedCount > 1 Then
        .Tools("ID_Properties").Enabled = False
      Else
        .Tools("ID_Properties").Enabled = True
      End If
    Else
      .Tools("ID_Delete").Enabled = False
      .Tools("ID_Properties").Enabled = False
    End If
    
    .Tools("ID_SaveChanges").Enabled = Application.Changed ' tbMain.Tools("ID_SaveChanges").Enabled
  
  End With
  
  ' With the ToolBar control (hidden)...
  With tbMain
   
    '==================================================
    ' Configure the Edit menu tools.
    '==================================================
    
    ' Enable/disable the required tools.
    .Tools("ID_New").Enabled = True
    .Tools("ID_CopyDef").Visible = False
    .Tools("ID_Delete").Enabled = (frmPictMgr.ListView1_SelectedCount > 0)
    .Tools("ID_SelectAll").Enabled = (frmPictMgr.ListView1.ListItems.Count > 0)
    .Tools("ID_Properties").Enabled = (frmPictMgr.ListView1_SelectedCount = 1)

    ' Reassign shortcuts if required.
'    .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'    .Tools("ID_Delete").Shortcut = ssDel
'    .Tools("ID_ScreenSelectAll").Shortcut = ssShortcutNone
'    .Tools("ID_SelectAll").Shortcut = ssCtrlA
'    .Tools("ID_ScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_Properties").Shortcut = ssF4

    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuEdit").Menu.Tools("ID_Properties").Index
    '.ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_New").Index
    '.ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Properties").Index

    ' Display the required menu and it's tools.
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)

    
    .Tools("ID_mnuEdit").Visible = True
    .Tools("ID_New").Visible = True
    .Tools("ID_CopyDef").Visible = False
    .Tools("ID_Delete").Visible = True
    .Tools("ID_SelectAll").Visible = True
    .Tools("ID_Properties").Visible = True

    '==================================================
    ' Configure the View menu tools.
    '==================================================
    ' Enable/disable the required tools.
    .Tools("ID_LargeIcons").Enabled = True
    .Tools("ID_SmallIcons").Enabled = True
    .Tools("ID_List").Enabled = True
    .Tools("ID_Details").Enabled = True
    .Tools("ID_CustomiseColumns").Visible = True
    ChangeView frmPictMgr.ListView1.View

    ' Add the required separators to the tool and menu bars.
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_LargeIcons").Index

    ' Display the required menu and it's tools.
    .Tools("ID_mnuView").Visible = True
    .Tools("ID_LargeIcons").Visible = True
    .Tools("ID_SmallIcons").Visible = True
    .Tools("ID_List").Visible = True
    .Tools("ID_Details").Visible = True
    .Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Enabled And frmPictMgr.ListView1.View = lvwReport
         
    '==================================================
    ' Configure the Window menu tools.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
      
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuWindow").Visible = True
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True
  
  End With
  
  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_WorkflowMgr(piFormCount As Integer)
'''
'''  ' Refresh the toolbar on the Screen Manager screen.
'''  With frmScrOpen
'''    .cmdAction(0).Enabled = True
'''    .cmdAction(1).Enabled = (frmScrOpen.ScreenID > 0)
'''    .cmdAction(2).Enabled = (frmScrOpen.ScreenID > 0)
'''    .cmdAction(3).Enabled = (frmScrOpen.ScreenID > 0)
'''    .cmdAction(4).Enabled = (frmScrOpen.ScreenID > 0)
'''  End With

  ' With the ToolBar control on the main MDI form (hidden)...
  With tbMain

    '==================================================
    ' Configure the Edit menu tools.
    '==================================================
'''    ' Enable/disable the required tools.
'''    .Tools("ID_New").Enabled = True
'''    .Tools("ID_Open").Enabled = (frmScrOpen.ScreenID > 0)
'''    .Tools("ID_Delete").Enabled = (frmScrOpen.ScreenID > 0)
'''    .Tools("ID_ScreenProperties").Enabled = (frmScrOpen.ScreenID > 0)
'''    .Tools("ID_CopyScreen").Enabled = (frmScrOpen.ScreenID > 0)
'''
'''    ' Reassign shortcuts if required.
''''    .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
''''    .Tools("ID_Delete").Shortcut = ssDel
''''    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssShortcutNone
''''    .Tools("ID_Properties").Shortcut = ssShortcutNone
''''    .Tools("ID_ScreenProperties").Shortcut = ssF4
'''
'''    ' Add the required separators to the tool and menu bars.
''''    .Tools("ID_mnuEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuEdit").Menu.Tools("ID_ScreenProperties").Index
''''    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_New").Index
''''    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_ScreenProperties").Index
'''
'''    ' Display the required menu and it's tools.
'''    .Tools("ID_mnuEdit").Visible = True
'''    .Tools("ID_New").Visible = True
'''    .Tools("ID_Open").Visible = True
'''    .Tools("ID_Delete").Visible = True
'''    .Tools("ID_ScreenProperties").Visible = True
'''    .Tools("ID_CopyScreen").Visible = True

    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)


    '==================================================
    ' Configure the Window menu tools.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True

    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index

    ' Display the required menu and it's tools.
    .Tools("ID_mnuWindow").Visible = False
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True

  End With

  CheckForDisabledMenuItems

End Sub


Private Sub RefreshMenu_ScrMgr(piFormCount As Integer)
  
  ' Refresh the toolbar on the Screen Manager screen.
  With frmScrOpen
    .cmdAction(0).Enabled = True
    .cmdAction(1).Enabled = (frmScrOpen.ScreenID > 0)
    .cmdAction(2).Enabled = (frmScrOpen.ScreenID > 0)
    .cmdAction(3).Enabled = (frmScrOpen.ScreenID > 0)
    .cmdAction(4).Enabled = (frmScrOpen.ScreenID > 0)
  End With
  
  ' With the ToolBar control on the main MDI form (hidden)...
  With tbMain
    
    '==================================================
    ' Configure the Edit menu tools.
    '==================================================
    ' Enable/disable the required tools.
    .Tools("ID_New").Enabled = True
    .Tools("ID_Open").Enabled = (frmScrOpen.ScreenID > 0)
    .Tools("ID_Delete").Enabled = (frmScrOpen.ScreenID > 0)
    .Tools("ID_ScreenProperties").Enabled = (frmScrOpen.ScreenID > 0)
    .Tools("ID_CopyScreen").Enabled = (frmScrOpen.ScreenID > 0)
            
    ' Reassign shortcuts if required.
'    .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'    .Tools("ID_Delete").Shortcut = ssDel
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_Properties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenProperties").Shortcut = ssF4
    
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuEdit").Menu.Tools("ID_ScreenProperties").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_New").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_ScreenProperties").Index
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuEdit").Visible = True
    .Tools("ID_New").Visible = True
    .Tools("ID_CopyDef").Visible = True
    .Tools("ID_Open").Visible = True
    .Tools("ID_Delete").Visible = True
    .Tools("ID_ScreenProperties").Visible = True
    .Tools("ID_CopyScreen").Visible = True
      
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)

      
    '==================================================
    ' Configure the Window menu tools.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
     
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuWindow").Visible = False
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True
  
  End With
  
  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_ScrDesigner(piFormCount As Integer)
  
  Dim objScreen As frmScrDesigner2
  Dim fDesignerActive As Boolean
  Dim fControlsExist As Boolean
  Dim bFormHasControls As Boolean
  Dim bAlignControlsEnabled As Boolean
  Dim bArrangeControlsEnabled As Boolean
  Dim fSSIScreen As Boolean
  
  bAlignControlsEnabled = True
  bArrangeControlsEnabled = True
  fSSIScreen = False
  
  fDesignerActive = (Me.ActiveForm.Name = "frmScrDesigner2")
  
  If fDesignerActive Then
    Set objScreen = Me.ActiveForm
  Else
    Set objScreen = Me.ActiveForm.CurrentScreen
  End If
  
  fSSIScreen = objScreen.IsSSIntranetScreen
  
  ' Refresh the toolbar on the actual screen the user is editing
  With objScreen.abScreen
    
    ' Enable/disable the required tools.
    If objScreen.UndoAction = giACTION_NOACTION Then
      .Tools("ID_Undo").Enabled = False
      .Tools("ID_Undo").ToolTipText = ""
    Else
      .Tools("ID_Undo").Enabled = True
    
      Select Case objScreen.UndoAction
        Case giACTION_DROPTABPAGE
          .Tools("ID_Undo").ToolTipText = "&Undo Add Tab Page"
        Case giACTION_DROPCONTROL
          .Tools("ID_Undo").ToolTipText = "&Undo Add Control"
        Case giACTION_CUTCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Cut"
        Case giACTION_PASTECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Paste"
        Case giACTION_DELETETABPAGE
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Tab Page"
        Case giACTION_DELETECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Control"
        Case giACTION_MOVECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Move"
        Case giACTION_STRETCHCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Stretch"
        Case giACTION_AUTOFORMAT
          .Tools("ID_Undo").ToolTipText = "&Undo AutoFormat"
      End Select
      
    End If
    
    bFormHasControls = objScreen.ScreenHasControls
    fControlsExist = (objScreen.SelectedControlsCount > 0) 'Or (objScreen.tabPages.Tabs.Count > 0)
    .Tools("ID_Cut").Enabled = fControlsExist
    .Tools("ID_Copy").Enabled = fControlsExist
    .Tools("ID_Paste").Enabled = (objScreen.ClipboardControlsCount > 0)
    .Tools("ID_ScreenObjectDelete").Enabled = fControlsExist Or (objScreen.tabPages.Tabs.Count > 0)
    .Tools("ID_ScreenSelectAll").Enabled = bFormHasControls
    .Tools("ID_Save").Enabled = objScreen.IsChanged
            
    ' Reassign shortcuts if required.
'    .Tools("ID_Delete").Shortcut = ssShortcutNone
'    If fDesignerActive Then
'      .Tools("ID_ScreenObjectDelete").Shortcut = ssDel
'      .Tools("ID_Undo").Shortcut = ssCtrlZ
'      .Tools("ID_Cut").Shortcut = ssCtrlX
'      .Tools("ID_Copy").Shortcut = ssCtrlC
'      .Tools("ID_Paste").Shortcut = ssCtrlV
'      .Tools("ID_Save").Shortcut = ssCtrlS
'    Else
'      .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'      .Tools("ID_Undo").Shortcut = ssShortcutNone
'      .Tools("ID_Cut").Shortcut = ssShortcutNone
'      .Tools("ID_Copy").Shortcut = ssShortcutNone
'      .Tools("ID_Paste").Shortcut = ssShortcutNone
'      .Tools("ID_Save").Shortcut = ssShortcutNone
'    End If
      
    ' Display the required menu and it's tools.
    .Tools("ID_Undo").Visible = True
    .Tools("ID_Cut").Visible = True
    .Tools("ID_Copy").Visible = True
    .Tools("ID_Paste").Visible = True
    .Tools("ID_ScreenObjectDelete").Visible = True
    .Tools("ID_Save").Visible = True
      
    .Tools("ID_ScreenDesignerScreenProperties").Enabled = True
    .Tools("ID_ObjectProperties").Enabled = bFormHasControls
    .Tools("ID_Toolbox").Enabled = True
    .Tools("ID_ObjectOrder").Enabled = bFormHasControls
    .Tools("ID_AutoFormat").Enabled = (Not fSSIScreen)
            
    ' Reassign shortcuts if required.
'    .Tools("ID_Properties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssF4
    
    ' Display the required menu and it's tools.
    .Tools("ID_ScreenDesignerScreenProperties").Visible = True
    .Tools("ID_ObjectProperties").Visible = True
    .Tools("ID_Toolbox").Visible = True
    .Tools("ID_ObjectOrder").Visible = True
    .Tools("ID_AutoFormat").Visible = True

End With
  
  ' With the ToolBar control on the main MDI form (hidden)
  
  With tbMain
    
    '==================================================
    ' Configure the Screen Edit menu tools.
    '==================================================
    ' Enable/disable the required tools.
    If objScreen.UndoAction = giACTION_NOACTION Then
      .Tools("ID_Undo").Enabled = False
      .Tools("ID_Undo").ToolTipText = ""
    Else
      .Tools("ID_Undo").Enabled = True
    
      Select Case objScreen.UndoAction
        Case giACTION_DROPTABPAGE
          .Tools("ID_Undo").ToolTipText = "&Undo Add Tab Page"
        Case giACTION_DROPCONTROL
          .Tools("ID_Undo").ToolTipText = "&Undo Add Control"
        Case giACTION_CUTCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Cut"
        Case giACTION_PASTECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Paste"
        Case giACTION_DELETETABPAGE
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Tab Page"
        Case giACTION_DELETECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Control"
        Case giACTION_MOVECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Move"
        Case giACTION_STRETCHCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Stretch"
        Case giACTION_AUTOFORMAT
          .Tools("ID_Undo").ToolTipText = "&Undo AutoFormat"
      End Select
      
    End If
    
    bFormHasControls = objScreen.ScreenHasControls
    fControlsExist = (objScreen.SelectedControlsCount > 0) 'Or (objScreen.tabPages.Tabs.Count > 0)
    .Tools("ID_Cut").Enabled = fControlsExist
    .Tools("ID_Copy").Enabled = fControlsExist
    .Tools("ID_Paste").Enabled = (objScreen.ClipboardControlsCount > 0)
    .Tools("ID_ScreenObjectDelete").Enabled = fControlsExist
'    .Tools("ID_ScreenSelectAll").Enabled = (objScreen.ScreenControlsCount > 0)
    .Tools("ID_ScreenSelectAll").Enabled = bFormHasControls
    .Tools("ID_Save").Enabled = objScreen.IsChanged
            
    ' Reassign shortcuts if required.
'    .Tools("ID_Delete").Shortcut = ssShortcutNone
'    If fDesignerActive Then
'      .Tools("ID_ScreenObjectDelete").Shortcut = ssDel
'      .Tools("ID_Undo").Shortcut = ssCtrlZ
'      .Tools("ID_Cut").Shortcut = ssCtrlX
'      .Tools("ID_Copy").Shortcut = ssCtrlC
'      .Tools("ID_Paste").Shortcut = ssCtrlV
'      .Tools("ID_Save").Shortcut = ssCtrlS
'    Else
'      .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'      .Tools("ID_Undo").Shortcut = ssShortcutNone
'      .Tools("ID_Cut").Shortcut = ssShortcutNone
'      .Tools("ID_Copy").Shortcut = ssShortcutNone
'      .Tools("ID_Paste").Shortcut = ssShortcutNone
'      .Tools("ID_Save").Shortcut = ssShortcutNone
'    End If
'    .Tools("ID_SelectAll").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenSelectAll").Shortcut = ssCtrlA
    
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuScreenEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuScreenEdit").Menu.Tools("ID_Cut").Index
'    .Tools("ID_mnuScreenEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuScreenEdit").Menu.Tools("ID_ScreenSelectAll").Index
'    .Tools("ID_mnuScreenEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuScreenEdit").Menu.Tools("ID_Save").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Undo").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Cut").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Save").Index
      
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuScreenEdit").Visible = True
    .Tools("ID_Undo").Visible = True
    .Tools("ID_Cut").Visible = True
    .Tools("ID_Copy").Visible = True
    .Tools("ID_Paste").Visible = True
    .Tools("ID_ScreenObjectDelete").Visible = True
    .Tools("ID_ScreenSelectAll").Visible = True
    .Tools("ID_Save").Visible = True
      
    '==================================================
    ' Configure the Screen Edit Tools menu.
    '==================================================
    ' Enable/disable the required tools.
    .Tools("ID_ScreenDesignerScreenProperties").Enabled = True
    .Tools("ID_ObjectProperties").Enabled = bFormHasControls
    .Tools("ID_Toolbox").Enabled = True
    .Tools("ID_ObjectOrder").Enabled = bFormHasControls
    .Tools("ID_AutoFormat").Enabled = (Not fSSIScreen)
    .Tools("ID_Options").Enabled = True
            
    ' Reassign shortcuts if required.
'    .Tools("ID_Properties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssF4
    
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuScreenEditTools").Menu.Tools.Add "separator", , .Tools("ID_mnuScreenEditTools").Menu.Tools("ID_AutoFormat").Index
'    .Tools("ID_mnuScreenEditTools").Menu.Tools.Add "separator", , .Tools("ID_mnuScreenEditTools").Menu.Tools("ID_Options").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_ScreenDesignerScreenProperties").Index
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuScreenEditTools").Visible = True
    .Tools("ID_ScreenDesignerScreenProperties").Visible = True
    .Tools("ID_ObjectProperties").Visible = True
    .Tools("ID_Toolbox").Visible = True
    .Tools("ID_ObjectOrder").Visible = True
    .Tools("ID_AutoFormat").Visible = True
    .Tools("ID_Options").Visible = True
      
    .Tools("ID_AutoLabel").Visible = True
    
    '==================================================
    ' Configure the Window menu tools.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
      
    ' Add the required separators to the tool and menu bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuWindow").Visible = True
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True

    ' Align / arrange controls
    With .Bands("ID_mnuScreenEdit")
      .Tools("ID_AlignControls").Visible = bAlignControlsEnabled
      .Tools("ID_BringToFront").Visible = bArrangeControlsEnabled
      .Tools("ID_SendToBack").Visible = bArrangeControlsEnabled
      .Tools("ID_ResurrectAll").Visible = False 'disabled for now
      
      .Tools("ID_BringToFront").Enabled = fControlsExist
      .Tools("ID_SendToBack").Enabled = fControlsExist
      .Tools("ID_ResurrectAll").Enabled = False 'disabled for now
    End With

    With .Bands("ID_mnuAlign")
      .Tools("ID_ScreenControlAlignLeft").Enabled = fControlsExist
      .Tools("ID_ScreenControlAlignRight").Enabled = fControlsExist
      .Tools("ID_ScreenControlAlignCentre").Enabled = fControlsExist
      .Tools("ID_ScreenControlAlignTop").Enabled = fControlsExist
      .Tools("ID_ScreenControlAlignMiddle").Enabled = fControlsExist
      .Tools("ID_ScreenControlAlignBottom").Enabled = fControlsExist
    End With

  End With

  ' Disassociate object variables.
  Set objScreen = Nothing

  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_WebFormDesigner(piFormCount As Integer)
  
  Dim objScreen As frmWorkflowWFDesigner
  Dim fDesignerActive As Boolean
  Dim fControlsExist As Boolean
  Dim bFormHasControls As Boolean
  Dim bAlignControlsEnabled As Boolean
  Dim bArrangeControlsEnabled As Boolean
  
  bAlignControlsEnabled = True
  bArrangeControlsEnabled = True
  
  fDesignerActive = (Me.ActiveForm.Name = "frmWorkflowWFDesigner")
  
  If fDesignerActive Then
    Set objScreen = Me.ActiveForm
  Else
    Set objScreen = Me.ActiveForm.CurrentWebForm
  End If
  
  ' Refresh the toolbar on the actual screen the user is editing
  With objScreen.abWebForm
    
    ' Enable/disable the required tools.
    If objScreen.UndoAction = giACTION_NOACTION Then
      .Tools("ID_Undo").Enabled = False
      .Tools("ID_Undo").ToolTipText = ""
    Else
      .Tools("ID_Undo").Enabled = Not objScreen.ReadOnly
    
      Select Case objScreen.UndoAction
        Case giACTION_DROPCONTROL, giACTION_DROPCONTROLAUTOLABEL
          .Tools("ID_Undo").ToolTipText = "&Undo Add Control"
        Case giACTION_CUTCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Cut"
        Case giACTION_PASTECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Paste"
        Case giACTION_DELETECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Control"
        Case giACTION_MOVECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Move"
        Case giACTION_STRETCHCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Stretch"
      End Select
      
    End If
    
    bFormHasControls = objScreen.ScreenHasControls
    fControlsExist = (objScreen.SelectedControlsCount > 0) 'Or (objScreen.tabPages.Tabs.Count > 0)
    .Tools("ID_Cut").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    .Tools("ID_Copy").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    .Tools("ID_Paste").Enabled = (objScreen.ClipboardControlsCount > 0) And (Not objScreen.ReadOnly)
    .Tools("ID_ScreenObjectDelete").Enabled = (fControlsExist And (Not objScreen.ReadOnly)) Or objScreen.tabPages.Tabs.Count > 0
    .Tools("ID_ScreenSelectAll").Enabled = bFormHasControls And (Not objScreen.ReadOnly)
    .Tools("ID_mnuWFSave").Enabled = objScreen.IsChanged And (Not objScreen.ReadOnly)
            
    ' Display the required menu and it's tools.
    .Tools("ID_Undo").Visible = True
    .Tools("ID_Cut").Visible = True
    .Tools("ID_Copy").Visible = True
    .Tools("ID_Paste").Visible = True
    .Tools("ID_ScreenObjectDelete").Visible = True
    .Tools("ID_mnuWFSave").Visible = True
      
    .Tools("ID_ObjectProperties").Enabled = True
    .Tools("ID_ObjectPropertiesScreen").Enabled = (objScreen.SelectedControlsCount = 1)
    .Tools("ID_WebFormPropertiesScreen").Enabled = True
    .Tools("ID_Toolbox").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_ObjectOrder").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_AutoFormat").Enabled = Not (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)

    ' Display the required menu and it's tools.
    .Tools("ID_ObjectProperties").Visible = True
    .Tools("ID_ObjectPropertiesScreen").Visible = True
    .Tools("ID_WebFormPropertiesScreen").Visible = True
    .Tools("ID_Toolbox").Visible = True
    .Tools("ID_ObjectOrder").Visible = True
    .Tools("ID_AutoFormat").Visible = True

    .Tools("ID_AutoLabel").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_AutoLabel").Visible = True
  End With
  
  ' With the ToolBar control on the main MDI form (hidden)
  With tbMain
    
    '==================================================
    ' Configure the Screen Edit menu tools.
    '==================================================
    ' Enable/disable the required tools.
    If objScreen.UndoAction = giACTION_NOACTION Then
      .Tools("ID_Undo").Enabled = False
      .Tools("ID_Undo").ToolTipText = ""
    Else
      .Tools("ID_Undo").Enabled = (Not objScreen.ReadOnly)
    
      Select Case objScreen.UndoAction
        Case giACTION_DROPCONTROL, giACTION_DROPCONTROLAUTOLABEL
          .Tools("ID_Undo").ToolTipText = "&Undo Add Control"
        Case giACTION_CUTCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Cut"
        Case giACTION_PASTECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Paste"
        Case giACTION_DELETECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Delete Control"
        Case giACTION_MOVECONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Move"
        Case giACTION_STRETCHCONTROLS
          .Tools("ID_Undo").ToolTipText = "&Undo Stretch"
      End Select
      
    End If
    
    bFormHasControls = objScreen.ScreenHasControls
    fControlsExist = (objScreen.SelectedControlsCount > 0) 'Or (objScreen.tabPages.Tabs.Count > 0)
    .Tools("ID_Cut").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    .Tools("ID_Copy").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    .Tools("ID_Paste").Enabled = (objScreen.ClipboardControlsCount > 0) And (Not objScreen.ReadOnly)
    .Tools("ID_ScreenObjectDelete").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    .Tools("ID_ScreenSelectAll").Enabled = bFormHasControls And (Not objScreen.ReadOnly)
    .Tools("ID_mnuWFSave").Enabled = objScreen.IsChanged And (Not objScreen.ReadOnly)
     
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuWebFormEdit").Visible = True
    .Tools("ID_Undo").Visible = True
    .Tools("ID_Cut").Visible = True
    .Tools("ID_Copy").Visible = True
    .Tools("ID_Paste").Visible = True
    .Tools("ID_ScreenObjectDelete").Visible = True
    .Tools("ID_ScreenSelectAll").Visible = True
    .Tools("ID_mnuWFSave").Visible = True
      
    '==================================================
    ' Configure the Screen Edit Tools menu.
    '==================================================
    ' Enable/disable the required tools.
'    .Tools("ID_mnuWFProperties").Enabled = True
    .Tools("ID_ObjectProperties").Enabled = True
    .Tools("ID_ObjectPropertiesScreen").Enabled = (objScreen.SelectedControlsCount = 1)
    .Tools("ID_WebFormPropertiesScreen").Enabled = True
    .Tools("ID_Toolbox").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_ObjectOrder").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_AutoFormat").Enabled = Not (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
            
    ' Display the required menu and it's tools.
'    .Tools("ID_mnuWebFormEditTools").Visible = True
'    .Tools("ID_mnuWFProperties").Visible = True
    .Tools("ID_ObjectProperties").Visible = True
    .Tools("ID_ObjectPropertiesScreen").Visible = True
    .Tools("ID_WebFormPropertiesScreen").Visible = True
    .Tools("ID_Toolbox").Visible = True
    .Tools("ID_ObjectOrder").Visible = True
    .Tools("ID_AutoFormat").Visible = True
    .Tools("ID_Options").Visible = True
      
    .Tools("ID_AutoLabel").Enabled = (Not objScreen.ReadOnly)
    .Tools("ID_AutoLabel").Visible = True
    
    '==================================================
    ' Configure the Window menu tools.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
      
    ' Display the required menu and it's tools.
    .Tools("ID_mnuWindow").Visible = True
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True

    ' Align / arrange controls
    With .Bands("ID_mnuWebFormEdit")
      .Tools("ID_AlignControls").Visible = bAlignControlsEnabled
      .Tools("ID_BringToFront").Visible = bArrangeControlsEnabled
      .Tools("ID_SendToBack").Visible = bArrangeControlsEnabled
      .Tools("ID_ResurrectAll").Visible = False 'disabled for now
      
      .Tools("ID_AlignControls").Enabled = (Not objScreen.ReadOnly)
      .Tools("ID_BringToFront").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_SendToBack").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ResurrectAll").Enabled = False 'disabled for now
    End With

    With .Bands("ID_mnuAlign")
      .Tools("ID_ScreenControlAlignLeft").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ScreenControlAlignRight").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ScreenControlAlignCentre").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ScreenControlAlignTop").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ScreenControlAlignMiddle").Enabled = fControlsExist And (Not objScreen.ReadOnly)
      .Tools("ID_ScreenControlAlignBottom").Enabled = fControlsExist And (Not objScreen.ReadOnly)
    End With

  End With

  ' Disassociate object variables.
  Set objScreen = Nothing

  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_Defaults(piFormCount As Integer)
  ' Refresh the default menu and tool bars for use with all module.
  
  Dim blnReadonly As Boolean
  
  blnReadonly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)
  
  ' With the ToolBar control ...
  With tbMain
    '==================================================
    ' Configure the Module menu.
    '==================================================
    ' Enable/disable the required tools.
    ' Only enable the module menu options if the modules are not already active.
    .Tools("ID_DatMgr").Enabled = (piFormCount <= 1) And Not gbLicenceExpired
    .Tools("ID_PicMgr").Enabled = (piFormCount <= 1 And Not blnReadonly And Not gbLicenceExpired)
    .Tools("ID_ScrMgr").Enabled = (piFormCount <= 1 And Not blnReadonly And Not gbLicenceExpired)
    .Tools("ID_WorkflowMgr").Enabled = (piFormCount <= 1) And Application.WorkflowModule And Not gbLicenceExpired
    .Tools("ID_ViewMgr").Enabled = (piFormCount <= 1) And Not gbLicenceExpired
    .Tools("ID_MobileDesigner").Enabled = (piFormCount <= 1) And Application.MobileModule And Not gbLicenceExpired
    .Tools("ID_ImportDefinitions").Enabled = (piFormCount <= 1) And Not gbLicenceExpired
    
    .Tools("ID_SSIntranet").Enabled = (piFormCount <= 1) And Application.SelfServiceIntranetModule And Not gbLicenceExpired
    .Tools("ID_SaveChanges").Enabled = Application.Changed
             
    ' Enable/disable Module setup menu options if the modules are activated.
    .Tools("ID_TrainingBooking").Enabled = Application.TrainingBookingModule And Not gbLicenceExpired
    .Tools("ID_Personnel").Enabled = Application.PersonnelModule And Not gbLicenceExpired
    .Tools("ID_Maternity").Enabled = Application.PersonnelModule And Not gbLicenceExpired
    .Tools("ID_Post").Enabled = Application.PersonnelModule And Not gbLicenceExpired
    .Tools("ID_Absence").Enabled = Application.AbsenceModule And Not gbLicenceExpired
    .Tools("ID_AccordTransfer").Enabled = IsModuleEnabled(modAccord) And Not gbLicenceExpired
    .Tools("ID_CMG").Enabled = IsModuleEnabled(modCMG) And Not gbLicenceExpired
    .Tools("ID_WorkflowSetup").Enabled = Application.WorkflowModule And Not gbLicenceExpired
    .Tools("ID_MobileSetup").Enabled = Application.MobileModule And Not gbLicenceExpired
    .Tools("ID_ModuleDocument").Enabled = Application.Version1Module And Not gbLicenceExpired
    .Tools("ID_AuditModule").Enabled = Not gbLicenceExpired
    .Tools("ID_BankHoliday").Enabled = Application.PersonnelModule And Not gbLicenceExpired
    .Tools("ID_Currency").Enabled = Not gbLicenceExpired
    .Tools("ID_Configuration").Enabled = Not gbLicenceExpired
    .Tools("ID_CategorySetup").Enabled = Not gbLicenceExpired
    .Tools("ID_LicenceInfo").Enabled = True
    
    ' Display the required menu and it's tools.
    .Tools("ID_mnuModule").Visible = True
    .Tools("ID_mnuConfiguration").Visible = True
    .Tools("ID_mnuAdministration").Visible = True
    
    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)
       
    ' Structure menu
    .Tools("ID_DatMgr").Visible = True
    .Tools("ID_ScrMgr").Visible = True
    .Tools("ID_WorkflowMgr").Visible = True
    .Tools("ID_PicMgr").Visible = True
    .Tools("ID_ViewMgr").Visible = True
    .Tools("ID_SSIntranet").Visible = True
    .Tools("ID_MobileDesigner").Visible = True
    .Tools("ID_ImportDefinitions").Visible = True
    
    ' Configuration menu remove disabled menuitems
    .Tools("ID_TrainingBooking").Visible = True
    .Tools("ID_Personnel").Visible = True
    .Tools("ID_Maternity").Visible = True
    .Tools("ID_Post").Visible = True
    .Tools("ID_Absence").Visible = True
    .Tools("ID_AccordTransfer").Visible = True
    .Tools("ID_CMG").Visible = True
    .Tools("ID_WorkflowSetup").Visible = True
    .Tools("ID_MobileSetup").Visible = True
    .Tools("ID_ModuleDocument").Visible = True
    .Tools("ID_AuditModule").Visible = True
    .Tools("ID_BankHoliday").Visible = True
    .Tools("ID_Currency").Visible = True
    .Tools("ID_Configuration").Visible = True
    .Tools("ID_CategorySetup").Visible = True
    .Tools("ID_LicenceInfo").Visible = True
    
    .Tools("ID_SaveChanges").Visible = True
    .Tools("ID_Logoff").Visible = True
    .Tools("ID_Exit").Visible = True
    
    '==================================================
    ' Configure the Help menu tools.
    '==================================================
    ' Display the required menu and it's tools.
    .Tools("ID_mnuHelp").Visible = True
    .Tools("ID_ContentsandIndex").Visible = True
    .Tools("ID_ViewCurrentUsers").Visible = True
    .Tools("ID_About").Visible = True
    .Tools("ID_VersionInfo").Visible = True
'    .Tools("ID_mnuHelp").Menu.Tools.Add "separator", , .Tools("ID_mnuHelp").Menu.Tools("ID_About").Index

  End With

  CheckForDisabledMenuItems

End Sub

Private Sub RefreshMenu_ViewMgr(piFormCount As Integer)

  ' Refresh the menu and tool bars for use with the View Manager module.
  Dim fEnableNew As Boolean
  Dim fEnableDelete As Boolean
  Dim fEnableProperties As Boolean
  Dim fEnableSelectAll As Boolean
  Dim fEnableViews As Boolean
  Dim fEnableCopyView As Boolean
  
  Dim blnReadonly As Boolean

  blnReadonly = (Application.AccessMode = accSystemReadOnly)

  '==================================================
  ' Configure the Edit menu tools.
  '==================================================
  ' Enable/disable the required tools.
  fEnableNew = False
  fEnableCopyView = False
  fEnableDelete = False
  fEnableProperties = False
  fEnableSelectAll = False
  
' With the Toolbar Ctl on frmViewMgr
  With frmViewMgr
    If .ActiveView Is .trvTables Then
      If Not .trvTables.SelectedItem Is Nothing Then
        'fEnableNew = True
        'fEnableDelete = (.trvTables.SelectedItem.DataKey = "VIEW")
        fEnableNew = Not blnReadonly
        fEnableDelete = (.trvTables.SelectedItem.DataKey = "VIEW") And Not blnReadonly
        fEnableCopyView = (.trvTables.SelectedItem.DataKey = "VIEW") And Not blnReadonly
        fEnableProperties = (.trvTables.SelectedItem.DataKey = "VIEW")
        fEnableSelectAll = (.trvTables.SelectedItem.DataKey = "TABLE") And (.lstViews.ListItems.Count > 0)
      End If
    Else
      If .ActiveView Is .lstViews Then
        'fEnableNew = True
        fEnableNew = Not blnReadonly
        
        If .lstViews.ListItems.Count > 0 And .lstViews_SelectedCount > 0 Then
          'fEnableDelete = True
          fEnableDelete = Not blnReadonly
          fEnableCopyView = Not blnReadonly
          fEnableProperties = (.lstViews_SelectedCount = 1)
        End If
        fEnableSelectAll = .lstViews.ListItems.Count > 0
      End If
    End If
  
    .abViewMgr.Tools("ID_SaveChanges").Enabled = Application.Changed ' tbMain.Tools("ID_SaveChanges").Enabled
    .abViewMgr.Tools("ID_New").Enabled = fEnableNew
    .abViewMgr.Tools("ID_CopyDef").Enabled = fEnableCopyView
    .abViewMgr.Tools("ID_Delete").Enabled = fEnableDelete
    .abViewMgr.Tools("ID_Properties").Enabled = fEnableProperties
  End With
  '==================================================
  ' Configure the View menu.
  '==================================================
  ' Enable/disable the required tools.
  fEnableViews = False
  With frmViewMgr
    If .ActiveView Is .lstViews Then
      fEnableViews = True
    Else
      If .ActiveView Is .trvTables Then
        If Not .trvTables.SelectedItem Is Nothing Then
          fEnableViews = (.trvTables.SelectedItem.DataKey = "TABLE")
        End If
      End If
    End If
    .abViewMgr.Tools("ID_LargeIcons").Enabled = fEnableViews
    .abViewMgr.Tools("ID_SmallIcons").Enabled = fEnableViews
    .abViewMgr.Tools("ID_List").Enabled = fEnableViews
    .abViewMgr.Tools("ID_Details").Enabled = fEnableViews
    .abViewMgr.Tools("ID_CustomiseColumns").Enabled = .abViewMgr.Tools("ID_Details").Enabled And frmViewMgr.lstViews.View = lvwReport
  End With
    
  ' With the main ToolBar control ...
  With tbMain
    
    With frmViewMgr
      If .ActiveView Is .trvTables Then
        If Not .trvTables.SelectedItem Is Nothing Then
          'fEnableNew = True
          'fEnableDelete = (.trvTables.SelectedItem.DataKey = "VIEW")
          fEnableNew = Not blnReadonly
          fEnableDelete = (.trvTables.SelectedItem.DataKey = "VIEW") And Not blnReadonly
          fEnableCopyView = (.trvTables.SelectedItem.DataKey = "VIEW") And Not blnReadonly
          fEnableProperties = (.trvTables.SelectedItem.DataKey = "VIEW")
          fEnableSelectAll = (.trvTables.SelectedItem.DataKey = "TABLE") And (.lstViews.ListItems.Count > 0)
        End If
      Else
        If .ActiveView Is .lstViews Then
          'fEnableNew = True
          fEnableNew = Not blnReadonly
          
          If .lstViews.ListItems.Count > 0 And .lstViews_SelectedCount > 0 Then
            'fEnableDelete = True
            fEnableDelete = Not blnReadonly
            fEnableCopyView = Not blnReadonly
            fEnableProperties = (.lstViews_SelectedCount = 1)
          End If
          fEnableSelectAll = .lstViews.ListItems.Count > 0
        End If
      End If
    End With
    .Tools("ID_New").Enabled = fEnableNew
    .Tools("ID_Delete").Enabled = fEnableDelete
    .Tools("ID_CopyDef").Enabled = fEnableCopyView
    .Tools("ID_Properties").Enabled = fEnableProperties
    .Tools("ID_SelectAll").Enabled = fEnableSelectAll
    
    ' Reassign shortcuts if required.
'    .Tools("ID_ScreenObjectDelete").Shortcut = ssShortcutNone
'    .Tools("ID_Delete").Shortcut = ssDel
'    .Tools("ID_SelectAll").Shortcut = ssCtrlA
'    .Tools("ID_ScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_ScreenDesignerScreenProperties").Shortcut = ssShortcutNone
'    .Tools("ID_Properties").Shortcut = ssF4
                
    ' Add the required separators to the menu and tool bars.
'    .Tools("ID_mnuEdit").Menu.Tools.Add "separator", , .Tools("ID_mnuEdit").Menu.Tools("ID_Properties").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_New").Index
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_Properties").Index
    
    ' Display the menu and it's required tools.
    .Tools("ID_mnuEdit").Visible = True
    .Tools("ID_New").Visible = True
    .Tools("ID_CopyDef").Visible = True
    .Tools("ID_Delete").Visible = True
    .Tools("ID_Properties").Visible = True
    .Tools("ID_SelectAll").Visible = True
     
    .Tools("ID_mnuAdministration").Visible = True
    .Tools("ID_Configuration").Visible = True

    '27/07/2001 MH
    .Tools("ID_SupportMode").Visible = Not IsModuleEnabled(modFullSysMgr)
    .Tools("ID_SupportMode").Enabled = (Application.AccessMode = accLimited)

     
    '==================================================
    ' Configure the View menu.
    '==================================================
    ' Enable/disable the required tools.
    fEnableViews = False
    With frmViewMgr
      If .ActiveView Is .lstViews Then
        fEnableViews = True
      Else
        If .ActiveView Is .trvTables Then
          If Not .trvTables.SelectedItem Is Nothing Then
            fEnableViews = (.trvTables.SelectedItem.DataKey = "TABLE")
          End If
        End If
      End If
    End With
    .Tools("ID_LargeIcons").Enabled = fEnableViews
    .Tools("ID_SmallIcons").Enabled = fEnableViews
    .Tools("ID_List").Enabled = fEnableViews
    .Tools("ID_Details").Enabled = fEnableViews
    .Tools("ID_CustomiseColumns").Visible = True
    .Tools("ID_CustomiseColumns").Enabled = .Tools("ID_Details").Enabled And frmViewMgr.lstViews.View = lvwReport
  
  
    If fEnableViews Then
      With tbMain
        Select Case frmViewMgr.lstViews.View
          Case lvwIcon
            .Tools("ID_LargeIcons").Checked = True
            .Tools("ID_SmallIcons").Checked = False
            .Tools("ID_List").Checked = False
            .Tools("ID_Details").Checked = False
          Case lvwSmallIcon
            .Tools("ID_LargeIcons").Checked = False
            .Tools("ID_SmallIcons").Checked = True
            .Tools("ID_List").Checked = False
            .Tools("ID_Details").Checked = False
          Case lvwList
            .Tools("ID_LargeIcons").Checked = False
            .Tools("ID_SmallIcons").Checked = False
            .Tools("ID_List").Checked = True
            .Tools("ID_Details").Checked = False
          Case lvwReport
            .Tools("ID_LargeIcons").Checked = False
            .Tools("ID_SmallIcons").Checked = False
            .Tools("ID_List").Checked = False
            .Tools("ID_Details").Checked = True
        End Select
      End With
    End If
    
    ' Add the required separators to the menu and tool bars.
'    .ToolBars("Main Toolbar").Tools.Add "separator", , .ToolBars("Main Toolbar").Tools("ID_LargeIcons").Index

    ' Display the menu and it's required tools.
    .Tools("ID_mnuView").Visible = True
    .Tools("ID_LargeIcons").Visible = True
    .Tools("ID_SmallIcons").Visible = True
    .Tools("ID_List").Visible = True
    .Tools("ID_Details").Visible = True
     
    '==================================================
    ' Configure the Window menu.
    '==================================================
    ' Enable/disable the required tools.
    ' Enable the Window menu options depending on the current active form's state.
    If piFormCount > 1 Then
      .Tools("ID_Maximise").Enabled = (Me.ActiveForm.MaxButton And Me.ActiveForm.WindowState <> vbMaximized)
      .Tools("ID_Minimise").Enabled = (Me.ActiveForm.MinButton And Me.ActiveForm.WindowState <> vbMinimized)
      .Tools("ID_Restore").Enabled = Me.ActiveForm.WindowState <> vbNormal
    End If
    .Tools("ID_Close").Enabled = True
    
    ' Add the required separators to the menu and tool bars.
'    .Tools("ID_mnuWindow").Menu.Tools.Add "separator", , .Tools("ID_mnuWindow").Menu.Tools("ID_Maximise").Index
      
    ' Display the menu and it's required tools.
    .Tools("ID_mnuWindow").Visible = True
    .Tools("ID_Cascade").Visible = True
    .Tools("ID_Arrange").Visible = False
    .Tools("ID_Maximise").Visible = True
    .Tools("ID_Minimise").Visible = True
    .Tools("ID_Restore").Visible = True
    .Tools("ID_Close").Visible = True

  End With
    
  CheckForDisabledMenuItems

End Sub


Private Sub ToolClick_DBMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
  
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
          
    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu
      
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
    
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
      
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner
      
    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
      
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
          
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
         
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
       
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
       
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
       
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing

    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing

    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing

    Case "ID_SaveChanges"
      SaveChanges_Click

    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to log off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
    '==================================================
    ' Edit menu.
    '==================================================
    Case "ID_New"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_Open"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Delete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_CopyDef"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_SelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Properties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_CopyColumn"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    Case "ID_Print"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    Case "ID_CopyClipboard"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu

    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu
      

    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
         
    '==================================================
    ' View menu.
    '==================================================
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      'ChangeView lvwIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      'ChangeView lvwSmallIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_List"
      ' Change the view to display a list.
      'ChangeView lvwList
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_Details"
      ' Change the view to display details.
      'ChangeView lvwReport
      objActiveForm.EditMenu pTool.Name
          
    Case "ID_CustomiseColumns"
      ' Customise which columns are displayed.
      objActiveForm.EditMenu pTool.Name
          
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      
      Dim plngHelp As Long
      
      If Not ShowAirHelp(0) Then
        plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to find and view the file " & App.HelpFile & ".", vbExclamation + vbOKOnly, App.EXEName
        End If
      End If

    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault

    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
            
  End Select
  

End Sub

Private Sub ToolClick_PictMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
  
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu
                  
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
        
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
        
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner
        
    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
        
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
             
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
        
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing

    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
        
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
        
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing
 
    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing
 
    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
 
    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing
    
    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      '' Save changes without exiting.
      'Set frmPrompt = New frmSaveChangesPrompt
      'frmPrompt.Buttons = vbOKCancel
      'frmPrompt.Show vbModal
      'If frmPrompt.Choice = vbOK Then
      '  Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
      '  If Not objActiveForm Is Nothing Then
      '    objActiveForm.SetFocus
      '  End If
      '  frmSysMgr.RefreshMenu
      'End If
      'Set frmPrompt = Nothing
      SaveChanges_Click

    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If

    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu

    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu


    '==================================================
    ' Edit menu.
    '==================================================
    Case "ID_New"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_Open"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Delete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_SelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Properties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
         
    '==================================================
    ' View menu.
    '==================================================
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      'ChangeView lvwIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      'ChangeView lvwSmallIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_List"
      ' Change the view to display a list.
      'ChangeView lvwList
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_Details"
      ' Change the view to display details.
      'ChangeView lvwReport
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_CustomiseColumns"
      ' Customise which columns are displayed.
      objActiveForm.EditMenu pTool.Name
      
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.
    
      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, "System Manager"
      End If
         
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault


    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

  End Select
  


End Sub
Private Sub ToolClick_ScrMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
    
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
          
    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu
                  
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
      
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
    
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner
    
    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
    
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
     
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
    
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
        
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
        
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
        
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing
     
    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
     
    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing

    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      '' Save changes without exiting.
      'Set frmPrompt = New frmSaveChangesPrompt
      'frmPrompt.Buttons = vbOKCancel
      'frmPrompt.Show vbModal
      'If frmPrompt.Choice = vbOK Then
      '  Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
      '  If Not objActiveForm Is Nothing Then
      '    objActiveForm.SetFocus
      '  End If
      '  frmSysMgr.RefreshMenu
      'End If
      'Set frmPrompt = Nothing
      SaveChanges_Click

    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If
    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu

    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu
      

    '==================================================
    ' Edit menu.
    '==================================================
    Case "ID_New"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_Open"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Delete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_SelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Properties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    Case "ID_CopyDef"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
         
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.
    
      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, "System Manager"
      End If
    
     
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
      

    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

  End Select

End Sub
Private Sub ToolClick_WorkflowMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)

  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm

  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu

    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu

    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu

    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu

    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing

    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing

    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner

    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions

    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
      
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing

    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing

    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing

    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing

    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing

    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing

    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing
      
    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      '' Save changes without exiting.
      'Set frmPrompt = New frmSaveChangesPrompt
      'frmPrompt.Buttons = vbOKCancel
      'frmPrompt.Show vbModal
      'If frmPrompt.Choice = vbOK Then
      '  Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
      '  If Not objActiveForm Is Nothing Then
      '    objActiveForm.SetFocus
      '  End If
      '  frmSysMgr.RefreshMenu
      'End If
      'Set frmPrompt = Nothing
      SaveChanges_Click

    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If

'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If
    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"

      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu

    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu


'''    '==================================================
'''    ' Edit menu.
'''    '==================================================
'''    Case "ID_New"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_Open"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_Delete"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_SelectAll"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_Properties"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_ScreenProperties"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name
'''
'''    Case "ID_CopyScreen"
'''      ' Pass the menu choice onto the active form to process.
'''      Screen.ActiveForm.EditMenu pTool.Name

    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms

    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons

    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized

    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm

    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.

      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, "System Manager"
      End If

    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass

      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"

      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If

      Screen.MousePointer = vbDefault


    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

  End Select

End Sub

Private Sub ToolClick_ViewMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
    
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
                  
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
      
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
            
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner
            
    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
            
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
                 
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
            
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
        
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
        
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
        
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing

    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing

    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing
      
    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      '' Save changes without exiting.
      'Set frmPrompt = New frmSaveChangesPrompt
      'frmPrompt.Buttons = vbOKCancel
      'frmPrompt.Show vbModal
      'If frmPrompt.Choice = vbOK Then
      '  Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
      '  If Not objActiveForm Is Nothing Then
      '    objActiveForm.SetFocus
      '  End If
      '  frmSysMgr.RefreshMenu
      'End If
      'Set frmPrompt = Nothing
      SaveChanges_Click

    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If

    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu
    
    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu
    
    
    '==================================================
    ' Edit menu.
    '==================================================
    Case "ID_New"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_Open"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Delete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_CopyDef"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_SelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Properties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
     
    '==================================================
    ' View menu.
    '==================================================
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      'ChangeView lvwIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      'ChangeView lvwSmallIcon
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_List"
      ' Change the view to display a list.
      'ChangeView lvwList
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_Details"
      ' Change the view to display details.
      'ChangeView lvwReport
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_CustomiseColumns"
      ' Customise which columns are displayed.
      objActiveForm.EditMenu pTool.Name
      
    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
         
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.
    
      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, Application.Name
      End If
       
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
      
      

    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
  End Select

End Sub
Private Sub ToolClick_SysMgr(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
    
  ' Process the tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Call up the Database Manager
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_ScrMgr"
      ' Call up the Screen Manager
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Call up the Workflow Manager
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Call up the Picture Manager
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
'      With frmSysMgr.tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With
          
    Case "ID_ViewMgr"
      ' Call up the View Manager
      If frmViewMgr Is Nothing Then
        Set frmViewMgr = New SystemMgr.frmViewMgr
      End If
      frmViewMgr.Show
      frmViewMgr.SetFocus
      frmSysMgr.RefreshMenu
            
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
      
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
            
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner

    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions

    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
                 
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
            
            
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
            
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
            
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
            
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing
      
    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing
      
    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
      
    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing

    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing

    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
   
    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      '' Save changes without exiting.
      'Set frmPrompt = New frmSaveChangesPrompt
      'frmPrompt.Buttons = vbOKCancel
      'frmPrompt.Show vbModal
      'If frmPrompt.Choice = vbOK Then
      '  Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
      '  If Not objActiveForm Is Nothing Then
      '    objActiveForm.SetFocus
      '  End If
      '  frmSysMgr.RefreshMenu
      'End If
      'Set frmPrompt = Nothing
      SaveChanges_Click
      
    Case "ID_SaveChangesNew"
      SaveChangesNew_Click
      
    Case "ID_Exit"
      ' Exit from the System Administrator module.
      UnLoad frmSysMgr

    Case "ID_Logoff"
    
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If

    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu

    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu


    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      
      Dim plngHelp As Long
      
      If Not ShowAirHelp(0) Then
        plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to find and view the file " & App.HelpFile & ".", vbExclamation + vbOKOnly, App.EXEName
        End If
      End If
    
    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
    
    
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing


    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
      
      

    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

  End Select
    
End Sub
Private Sub ToolClick_ScrDesigner(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
  
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
          
    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu
                  
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
      
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
            
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner

    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
            
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
                 
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
            
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
        
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
        
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
        
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing
      
    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
      
    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing
    
    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing
      
    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      SaveChanges_Click
          
    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If
     
    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu
     
    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu
    
    
    '==================================================
    ' Screen Edit menu.
    '==================================================
    Case "ID_Undo"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Cut"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Copy"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Paste"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenObjectDelete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenSelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Save"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_BringToFront"
      ' Bring selected controls to front
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_SendToBack"
      ' Send selected controls to back
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenControlAlignLeft"
      'Align controls on the left
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenControlAlignCentre"
      'Align controls in the centre
      objActiveForm.EditMenu pTool.Name
  
    Case "ID_ScreenControlAlignRight"
      'Align controls on the right
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_ScreenControlAlignTop"
      'Align controls at the top
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_ScreenControlAlignMiddle"
      'Align controls in the middle
      objActiveForm.EditMenu pTool.Name
        
    Case "ID_ScreenControlAlignBottom"
      'Align controls at the bottom
      objActiveForm.EditMenu pTool.Name

    '==================================================
    ' Tools menu.
    '==================================================
    Case "ID_ScreenDesignerScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ObjectProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Toolbox"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ObjectOrder"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_AutoFormat"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Options"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_AutoLabel"
      If mblnAutoLabelling = True Then Exit Sub
      mblnAutoLabelling = True
      'TM20011015 Fault 2959
      'Set the checked property of the AutoLabel button.
      Me.tbMain.Tools("ID_AutoLabel").Checked = Not Me.tbMain.Tools("ID_AutoLabel").Checked
      If TypeOf objActiveForm Is frmScrDesigner2 Then
        objActiveForm.abScreen.Tools("ID_AutoLabel").Checked = (Me.tbMain.Tools("ID_AutoLabel").Checked)
      Else
        Dim tmpform As Form
        For Each tmpform In Forms
          If TypeOf tmpform Is frmScrDesigner2 Then
            tmpform.abScreen.Tools("ID_AutoLabel").Checked = (Me.tbMain.Tools("ID_AutoLabel").Checked)
          End If
        Next tmpform
      End If
      
      mblnAutoLabelling = False
      
    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
      
      If mblnDisplayScrOpen = True Then
        
        ' Display the screen manager.
        If frmSysMgr.frmScrOpen Is Nothing Then
          Set frmSysMgr.frmScrOpen = New SystemMgr.frmScrOpen
        End If
        frmSysMgr.frmScrOpen.Show
        frmSysMgr.frmScrOpen.SetFocus
      
      End If
      
      frmSysMgr.RefreshMenu
   
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.
    
      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, Application.Name
      End If
          
    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
      
      
    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

    'JDM - 22/08/02 - Fault 3873 - View users not working
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing

  End Select
  
End Sub

Private Sub ToolClick_WebFormDesigner(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  Dim strVersionFilename As String
  Dim objActiveForm As Object
   
  Set objActiveForm = Me.ActiveForm
  
  ' Process tool click.
  Select Case pTool.Name
    '==================================================
    ' Module menu.
    '==================================================
    Case "ID_DatMgr"
      ' Display the Database Manager.
      If frmDbMgr Is Nothing Then
        Set frmDbMgr = New SystemMgr.frmDbMgr
      End If
      frmDbMgr.Show
      frmDbMgr.SetFocus
      frmSysMgr.RefreshMenu
        
    Case "ID_ScrMgr"
      ' Display the Screen Manager.
      If frmScrOpen Is Nothing Then
        Set frmScrOpen = New SystemMgr.frmScrOpen
      End If
      frmScrOpen.Show
      frmScrOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_WorkflowMgr"
      ' Display the Workflow Manager.
      If frmWorkflowOpen Is Nothing Then
        Set frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
      frmWorkflowOpen.Show
      frmWorkflowOpen.SetFocus
      frmSysMgr.RefreshMenu
    
    Case "ID_PicMgr"
      ' Display the Picture Manager.
      If frmPictMgr Is Nothing Then
        Set frmPictMgr = New SystemMgr.frmPictMgr
      End If
      frmPictMgr.Show
      frmPictMgr.SetFocus
      frmSysMgr.RefreshMenu
          
    Case "ID_ViewMgr"
      ' Display the View Manager.
       If frmViewMgr Is Nothing Then
         Set frmViewMgr = New SystemMgr.frmViewMgr
       End If
       frmViewMgr.Show
       frmViewMgr.SetFocus
       frmSysMgr.RefreshMenu
                  
    Case "ID_TrainingBooking"
      ' Call up the Training Booking Module Setup screen.
      frmTrainingBookingSetup.Show vbModal
      Set frmTrainingBookingSetup = Nothing
      
    Case "ID_SSIntranet"
      ' Call up the Self-service Intranet Module Setup screen.
      frmSSIntranetSetup.Show vbModal
      Set frmSSIntranetSetup = Nothing
            
    ' Edit the mobile definitions
    Case "ID_MobileDesigner"
      EditMobileDesigner
            
    ' Import any definitions
    Case "ID_ImportDefinitions"
      ImportDefinitions
            
    Case "ID_AccordTransfer"
      ' Call up the Payroll Tranfer module setup
      frmAccordPayrollTransfer.Show vbModal
      Set frmAccordPayrollTransfer = Nothing
                 
    Case "ID_CMG"
      ' Call up the CMG/Centrefile module setup
      frmCMGSetup.Show vbModal
      Set frmCMGSetup = Nothing
            
    Case "ID_Personnel"
      ' Call up the Personnel Module Setup screen.
      frmPersonnelSetup.Show vbModal
      Set frmPersonnelSetup = Nothing
        
    Case "ID_Absence"
      ' Call up the Absence Module Setup screen.
      frmAbsenceSetup.Show vbModal
      Set frmAbsenceSetup = Nothing
        
    Case "ID_AuditModule"
      frmAuditSetup.Show vbModal
      Set frmAuditSetup = Nothing
        
    Case "ID_BankHoliday"
      ' Call up the Bank Holiday Setup screen.
      frmBankHolidaySetup.Show vbModal
      Set frmBankHolidaySetup = Nothing

    Case "ID_CategorySetup"
      frmCategorySetup.Show vbModal
      Set frmCategorySetup = Nothing
      
    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
      
    Case "ID_Currency"
      'Call up the Currency setup screen.
      frmCurrencySetup.Show vbModal
      Set frmCurrencySetup = Nothing
    
    Case "ID_Maternity"
      'Call up the Maternity setup screen.
      frmMaternitySetup.Show vbModal
      Set frmMaternitySetup = Nothing

    Case "ID_Post"
      'Call up the Post setup screen.
      frmPostSetup.Show vbModal
      Set frmPostSetup = Nothing

    Case "ID_WorkflowSetup"
      ' Call up the Workflow Module Setup screen.
      frmWorkflowSetup.Show vbModal
      Set frmWorkflowSetup = Nothing

'    Case "ID_MobileSetup"
'      ' Call up the Mobile Module Setup screen.
'      frmMobileSetup.Show vbModal
'      Set frmMobileSetup = Nothing

    Case "ID_ModuleDocument"
      ' Call up the Version 1 Module Setup screen
      frmModuleDocument.Show vbModal
      Set frmModuleDocument = Nothing

    Case "ID_SaveChanges"
      '01/08/2001 MH Fault 2382
      SaveChanges_Click
          
    Case "ID_Exit"
      ' Exit the system.
      UnLoad frmSysMgr

    Case "ID_Logoff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to Log Off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            ' Logoff the system.
            UnLoad frmSysMgr
            ' Close the temporary database.
            If Forms.Count < 1 Then
              If Not daoDb Is Nothing Then
                daoDb.Close
              End If
              Main
            End If
        End If
    
'      ' Logoff the system.
'      UnLoad frmSysMgr
'      ' Close the temporary database.
'      If Forms.Count < 1 Then
'        If Not daoDb Is Nothing Then
'          daoDb.Close
'        End If
'        Main
'      End If
     
    '==================================================
    ' Administration menu.
    '==================================================
    Case "ID_Configuration"
    
      'MsgBox "Config Screen"
      frmConfiguration.Show vbModal
      RefreshMenu
     
    Case "ID_SupportMode"
      frmSupportMode.Show vbModal
      UnLoad frmSupportMode
      Set frmSupportMode = Nothing
      RefreshMenu
    
    
    '==================================================
    ' Screen Edit menu.
    '==================================================
    Case "ID_Undo"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Cut"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Copy"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Paste"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenObjectDelete"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenSelectAll"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_mnuWFSave"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_BringToFront"
      ' Bring selected controls to front
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_SendToBack"
      ' Send selected controls to back
      objActiveForm.EditMenu pTool.Name
      
    Case "ID_ResurrectAll"
      ' Make all controls visible again
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenControlAlignLeft"
      'Align controls on the left
      objActiveForm.EditMenu pTool.Name

    Case "ID_ScreenControlAlignCentre"
      'Align controls in the centre
      objActiveForm.EditMenu pTool.Name
  
    Case "ID_ScreenControlAlignRight"
      'Align controls on the right
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_ScreenControlAlignTop"
      'Align controls at the top
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_ScreenControlAlignMiddle"
      'Align controls in the middle
      objActiveForm.EditMenu pTool.Name
        
    Case "ID_ScreenControlAlignBottom"
      'Align controls at the bottom
      objActiveForm.EditMenu pTool.Name

    '==================================================
    ' Tools menu.
    '==================================================
    Case "ID_ScreenDesignerScreenProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ObjectProperties"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ObjectPropertiesScreen"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_WebFormPropertiesScreen"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name
    
    Case "ID_Toolbox"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_ObjectOrder"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_AutoFormat"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_Options"
      ' Pass the menu choice onto the active form to process.
      objActiveForm.EditMenu pTool.Name

    Case "ID_AutoLabel"
      If mblnAutoLabelling = True Then Exit Sub
      mblnAutoLabelling = True
      'TM20011015 Fault 2959
      'Set the checked property of the AutoLabel button.
      Me.tbMain.Tools("ID_AutoLabel").Checked = Not Me.tbMain.Tools("ID_AutoLabel").Checked
      If TypeOf objActiveForm Is frmScrDesigner2 Then
        objActiveForm.abScreen.Tools("ID_AutoLabel").Checked = (Me.tbMain.Tools("ID_AutoLabel").Checked)
      Else
        Dim tmpform As Form
        For Each tmpform In Forms
          If TypeOf tmpform Is frmScrDesigner2 Then
            tmpform.abScreen.Tools("ID_AutoLabel").Checked = (Me.tbMain.Tools("ID_AutoLabel").Checked)
          End If
        Next tmpform
      End If
      
      mblnAutoLabelling = False
      
    '==================================================
    ' Window menu.
    '==================================================
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms
      
    Case "ID_Arrange"
      ' Arrange window icons
      frmSysMgr.Arrange vbArrangeIcons
        
    Case "ID_Maximise"
      ' Maximise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMaximized
        
    Case "ID_Minimise"
      ' Minimise the current window.
      frmSysMgr.ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      frmSysMgr.ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      UnLoad frmSysMgr.ActiveForm
      
'      If mblnDisplayScrOpen = True Then
'
'        ' Display the screen manager.
'        If frmSysMgr.frmScrOpen Is Nothing Then
'          Set frmSysMgr.frmScrOpen = New SystemMgr.frmScrOpen
'        End If
'        frmSysMgr.frmScrOpen.Show
'        frmSysMgr.frmScrOpen.SetFocus
'
'      End If
      
      frmSysMgr.RefreshMenu
   
    '==================================================
    ' Help menu.
    '==================================================
    Case "ID_ContentsandIndex"
      '' To be done.
    
      Dim plngHelp As Long
      'Call the App.HelpFile function to get the helpfile for current app. e.g.(SYS)
      plngHelp = ShellExecute(0&, vbNullString, gsApplicationPath & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)

      If plngHelp = 0 Then
        MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HRProHelp.chm'.", vbExclamation + vbOKOnly, Application.Name
      End If
     
      
    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = gsApplicationPath & "\OpenHR System Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, Application.Name
      End If
      
      Screen.MousePointer = vbDefault
      
      
    Case "ID_About"
      ' Call up the 'About' screen.
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
'      With tbMain
'        .Redraw = False
'        .Enabled = False
'        .Enabled = True
'        .Redraw = True
'      End With

    'JDM - 22/08/02 - Fault 3873 - View users not working
    Case "ID_ViewCurrentUsers"
      'MH20010524 Will be required for read-only access...
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      UnLoad frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing

  End Select
  
End Sub
Public Sub ClearMenuShortcuts()
  ' Clear all menu shortcuts.
  Dim objTool As ActiveBarLibraryCtl.Tool
    
'  For Each objTool In tbMain.Tools
'    objTool.Shortcut = ssShortcutNone
'  Next
  
  Set objTool = Nothing

End Sub

'Private Sub Timer1_Timer()
'
'  Dim blnSystemLocked As Boolean
'  Dim strLockUser As String
'  Dim strLockType As String
'  Dim strLockDetails As String
'
'  GetLockDetails strLockUser, strLockType
'  blnSystemLocked = (strLockType <> vbNullString And strLockType <> "Lock Read Write")
'
'  If blnSystemLocked Then
'    'If not locked by current app then can we get read only access...
'    strLockDetails = "User :  " & strLockUser & vbCrLf & _
'                     "Date/Time :  " & GetSystemSetting(strLockType, "DateTime", "") & vbCrLf & _
'                     "Machine :  " & GetSystemSetting(strLockType, "Machine", "") & vbCrLf & _
'                     "Type :  " & strLockType
'
'    MsgBox "The database has been locked as follows:" & _
'           vbCrLf & vbCrLf & strLockDetails & vbCrLf & vbCrLf & _
'           "Please log off as soon as possible.", vbExclamation, "Database Locked"
'
'  End If
'
'End Sub

Public Sub SetCaption()

  Me.Caption = Application.Name & " - " & gsDatabaseName & "  " & _
      Choose(Application.AccessMode, "", "[Support Mode]", "[Limited Access]", "[Read Only]")

End Sub

Private Sub tbMain_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)

  ' Get rid of any screen display residue.
  DoEvents

  ' Process the tool click dependent on the active form.
  If Me.ActiveForm Is Nothing Then
    ToolClick_SysMgr pTool
  Else
    'JPD 20050113 Fault 9339
    'If Not Screen.ActiveForm Is Nothing Then
    If Not Me.ActiveForm Is Nothing Then
      If TypeOf Me.ActiveForm Is frmDbMgr Then
        ToolClick_DBMgr pTool
      ElseIf TypeOf Me.ActiveForm Is frmPictMgr Then
        ToolClick_PictMgr pTool
      ElseIf TypeOf Me.ActiveForm Is frmScrOpen Then
        ToolClick_ScrMgr pTool
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowOpen Then
        ToolClick_WorkflowMgr pTool
      ElseIf TypeOf Me.ActiveForm Is frmScrDesigner2 Then
        ToolClick_ScrDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmScrObjProps Then
        ToolClick_ScrDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmToolbox Then
        ToolClick_ScrDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFDesigner Then
        ToolClick_WebFormDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFItemProps Then
        ToolClick_WebFormDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowDesigner Then
        ToolClick_WorkflowMgr pTool
      ElseIf TypeOf Me.ActiveForm Is frmWorkflowWFToolbox Then
        ToolClick_WebFormDesigner pTool
      ElseIf TypeOf Me.ActiveForm Is frmViewMgr Then
        ToolClick_ViewMgr pTool
      End If
    End If
  End If

End Sub

Private Sub tbMain_MenuItemEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  DoEvents
End Sub

Private Sub tbMain_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)

  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub Timer1_Timer()
  ' Poll the server for any messages.
  Dim sSQL As String
  Dim sMessage As String
  Dim rsMessages As New ADODB.Recordset
  
  sMessage = ""
      
  sSQL = "exec sp_ASRGetMessages"
  rsMessages.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  Do While Not rsMessages.EOF
    If Len(sMessage) > 0 Then
      sMessage = sMessage & vbCrLf & vbCrLf & vbCrLf
    End If
    
    sMessage = sMessage & rsMessages.Fields(0).value
    
    rsMessages.MoveNext
  Loop
    
  rsMessages.Close
  Set rsMessages = Nothing

  If Len(sMessage) > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
  End If

  Set rsMessages = Nothing

End Sub

Private Sub SaveChangesNew_Click()

'  SaveChangesNew

End Sub


Private Sub SaveChanges_Click()

  Dim frmPrompt As frmSaveChangesPrompt
  Dim iCount As Integer
  
  '01/08/2001 MH Fault 2382
  For iCount = 0 To (Forms.Count - 1)
    If Forms(iCount).Name = "frmScrDesigner2" Then
      If Forms(iCount).IsChanged Then
        MsgBox "Please save changes within the Screen Designer prior to committing changes to the server.", vbExclamation, App.Title
        Exit Sub
      End If
    End If
    
    If Forms(iCount).Name = "frmWorkflowDesigner" Then
      If Forms(iCount).IsChanged Then
        MsgBox "Please save changes within the Workflow Designer prior to committing changes to the server.", vbExclamation, App.Title
        Exit Sub
      End If
    End If
  Next iCount
  
  ' Save changes
  Set frmPrompt = New frmSaveChangesPrompt
  frmPrompt.Buttons = vbOKCancel
  frmPrompt.Show vbModal
  If frmPrompt.Choice = vbOK Then
    Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
    If Not Me.ActiveForm Is Nothing Then
      Me.ActiveForm.SetFocus
    End If
    frmSysMgr.RefreshMenu
  End If
  Set frmPrompt = Nothing

End Sub

Private Function DoesTableExistInDB(ByRef lngObjectID As Long)

  Dim bFound As Boolean
  Dim rstTables As DAO.Recordset
  Dim sSQL As String
  
  bFound = False
  sSQL = "SELECT tableName, IsCopy" & _
    " FROM tmpTables" & _
    " WHERE tableID = " & lngObjectID & " AND IsCopy <> -1"
  Set rstTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  If Not (rstTables.EOF And rstTables.BOF) Then
    bFound = True
  End If

  rstTables.Close
  Set rstTables = Nothing

  DoesTableExistInDB = bFound

End Function


Private Function CheckForDisabledMenuItems()

  Dim objTool As ActiveBarLibraryCtl.Tool
  Dim strDisabledItems As String
  
  
  '--Use this setting to disable menu items.
  '--e.g. this SQL disables the workflow manager menu item.
  '
  'delete from asrsyssystemsettings
  'where [Section] = 'menu' and [SettingKey] = 'sysmgr'
  'Insert asrsyssystemsettings([Section], [SettingKey], [SettingValue])
  'values('menu', 'sysmgr', '20041')
  
  
  strDisabledItems = GetSystemSetting("Menu", "SysMgr", "")
  For Each objTool In tbMain.Tools
    If InStr("-" & strDisabledItems & "-", "-" & CStr(objTool.ToolID) & "-") > 0 Then
      objTool.Enabled = False
    End If
  Next

End Function

' A dummy procedure to keep the connection alive every 10 seconds
Private Sub tmrKeepAlive_Timer()
 
  On Error GoTo NetworkDown:
 
  Dim sSQL As String
  Dim sMessage As String
  Dim rsMessages As New ADODB.Recordset
  
  sMessage = ""
      
  sSQL = "exec sp_ASRGetMessages"
  rsMessages.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  Do While Not rsMessages.EOF
    If Len(sMessage) > 0 Then
      sMessage = sMessage & vbCrLf & vbCrLf & vbCrLf
    End If
    
    sMessage = sMessage & rsMessages.Fields(0).value
    rsMessages.MoveNext
  Loop
    
  rsMessages.Close
  Set rsMessages = Nothing

  If Len(sMessage) > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
  End If

  Set rsMessages = Nothing
  
  Exit Sub

NetworkDown:

  If Err.Description = "Connection failure" Or InStr(1, Err.Description, "General network error") Then
    AttemptReLogin
  End If

End Sub
