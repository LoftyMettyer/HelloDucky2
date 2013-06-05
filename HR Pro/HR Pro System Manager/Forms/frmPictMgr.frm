VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmPictMgr 
   Caption         =   "Picture Manager"
   ClientHeight    =   3255
   ClientLeft      =   4350
   ClientTop       =   3555
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5022
   Icon            =   "frmPictMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5970
   Begin ComctlLib.ListView ListView1 
      Height          =   2145
      Left            =   165
      TabIndex        =   0
      Top             =   495
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3784
      View            =   1
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "name"
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "type"
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "height"
         Object.Tag             =   ""
         Text            =   "Height"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "width"
         Object.Tag             =   ""
         Text            =   "Width"
         Object.Width           =   1058
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   3720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
      MaxFileSize     =   255
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   2955
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7435
            TextSave        =   ""
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
   Begin ActiveBarLibraryCtl.ActiveBar abPictMgrMenu 
      Left            =   4470
      Top             =   1380
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
      Bands           =   "frmPictMgr.frx":000C
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   3660
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmPictMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event Activate()
Event Deactivate()
Event UnLoad()

'Local variables to hold property values
Private blnChanged As Boolean
Private blnLoading As Boolean

'Local variables
Private gfMenuActionKey As Boolean

Private mfrmUse As frmUsage
 
Private Const MIN_FORM_HEIGHT = 5000
Private Const MIN_FORM_WIDTH = 6000

Public Property Get Changed() As Boolean
  
  ' Return the Changed property.
  Changed = blnChanged

End Property

Public Property Let Changed(pfHasChanged As Boolean)

  ' Set the Changed property.
  blnChanged = pfHasChanged
  
End Property

Public Property Get Loading() As Boolean
  Loading = blnLoading
End Property

Public Property Let Loading(IsLoading As Boolean)
  blnLoading = IsLoading
End Property

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
Private Sub ListView1_ClearSelections()

  Dim iLoop As Integer
  
  ' Loop through the list view items deselecting any currently selected items.
  For iLoop = 1 To ListView1.ListItems.Count
    ListView1.ListItems(iLoop).Selected = False
  Next iLoop

End Sub
Private Sub ListView1_SelectAll()

  Dim iLoop As Integer
  
  ' Loop through the list view items marking each one as selected.
  For iLoop = 1 To ListView1.ListItems.Count
    ListView1.ListItems(iLoop).Selected = True
  Next iLoop

End Sub


Private Sub abPictMgrMenu_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  EditMenu Tool.Name
  
End Sub

Private Sub abPictMgrMenu_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)

  ' Do not let the user modify the layout.
  Cancel = True
 
End Sub

Private Sub Form_Activate()
  Dim fExit As Boolean
  'Dim WaitWindow As WaitMessage.MessageWindow
'  Dim WaitWindow As NewWaitMsg.clsNewWaitMsg
  Dim strKey As String
  Dim strFileName As String
  
  If Me.Loading Then
    
    With recPictEdit
    
      .Index = "idxName"
      
      If Not (.BOF And .EOF) Then
        Screen.MousePointer = vbHourglass
      
        .MoveFirst
      
        'Set WaitWindow = New WaitMessage.MessageWindow
'        Set WaitWindow = New NewWaitMsg.clsNewWaitMsg
'        WaitWindow.Initialise .RecordCount, _
          "Loading pictures, please wait...", Me.Caption, True, True, True, App.Path & "\videos\picture.avi"
      
      'gobjProgress.AviFile = App.Path & "\videos\picture.Avi"
      gobjProgress.AVI = dbPicture
      gobjProgress.MainCaption = "Picture Manager"
      gobjProgress.Caption = "HR Pro - System Manager"
      gobjProgress.NumberOfBars = 1
      gobjProgress.Bar1MaxValue = .RecordCount
      gobjProgress.Bar1Caption = "Loading pictures..."
      gobjProgress.Time = True
      gobjProgress.Cancel = True
      gobjProgress.OpenProgress
      
'        Do While Not .EOF And Not WaitWindow.Cancelled
        Do While Not .EOF And Not gobjProgress.Cancelled
          strKey = "I" & Trim(Str(.Fields("pictureID")))
          strFileName = ReadPicture
          ImageList2.ListImages.Add , strKey, LoadPicture(strFileName)
          ImageList1.ListImages.Add , strKey, LoadPicture(strFileName)
          Kill strFileName
        
'          WaitWindow.UpdateProgress
          gobjProgress.UpdateProgress False
          
          .MoveNext
        Loop
        
        .MoveFirst
      
        'fExit = WaitWindow.Cancelled
        fExit = gobjProgress.Cancelled
      End If
    End With
  
    If fExit Then
      gobjProgress.CloseProgress
      UnLoad Me
      Screen.MousePointer = vbNormal
    Else
      PopulateListView
      Me.Loading = False
      RefreshListView
'      Set WaitWindow = Nothing
      If gobjProgress.Visible = True Then gobjProgress.CloseProgress
      Screen.MousePointer = vbNormal
      Me.SetFocus
    End If
  End If
  
  RaiseEvent Activate
  
  frmSysMgr.RefreshMenu
  
End Sub

Private Sub Form_Deactivate()
  RaiseEvent Deactivate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
  Dim sSection As String
  Dim sAppName As String
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  sSection = Me.Name
  sAppName = App.ProductName
  
  ' Initialize variables and properties.
  Me.Loading = True
  gfMenuActionKey = False
  
  ' Initialize form properties.
  ListView1.View = GetPCSetting(sSection, "View", lvwIcon)
  
  If gbMaximizeScreens Then
    Me.WindowState = vbMaximized
  Else
    Me.WindowState = GetPCSetting(Me.Name, "State", Me.WindowState)
  End If
  
  ChangeView 0
  
End Sub

Private Sub Form_Resize()

  On Local Error Resume Next

  ' Do not resize if the window is minimized.
  If Me.WindowState = vbMinimized Then
    Exit Sub
  End If

  ' Resize the contained controls accordingly.
  ListView1.Move 0, 0, _
    Me.ScaleWidth, _
    Me.ScaleHeight - StatusBar1.Height

  Me.Refresh

  frmSysMgr.RefreshMenu

  ' Get rid of the icon off the form
  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim sSection As String
  Dim sAppName As String
  
  sSection = Me.Name
  sAppName = App.ProductName
  
  If Me.Changed And Me.abPictMgrMenu.Tools("ID_SaveChanges").Enabled = True Then
    Application.Changed = True
  End If
  
  'Save form size and position to registry
  If Me.WindowState = vbNormal Then
    SavePCSetting sSection, "Top", Me.Top
    SavePCSetting sSection, "Left", Me.Left
    SavePCSetting sSection, "Height", Me.Height
    SavePCSetting sSection, "Width", Me.Width
    SavePCSetting sSection, "View", ListView1.View
  End If
  
  SavePCSetting Me.Name, "State", Me.WindowState
 
  'RaiseEvent UnLoad
  frmSysMgr.RefreshMenu True
  
  Unhook Me.hWnd
  
End Sub




Private Sub ListView1_AfterLabelEdit(Cancel As Integer, psNewString As String)
  ' Edit the picture's name.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sKey As String
  
  sKey = ListView1.SelectedItem.key
  
  fOK = EditPicture_Transaction(val(Mid(sKey, 2)), psNewString)
  
  If fOK Then
    ListView1.SelectedItem.Text = psNewString
    RefreshListView
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub ListView1_Click()
   
  frmSysMgr.RefreshMenu

End Sub

Private Sub ListView1_DblClick()
  
  ' Double-click activates the picture editor.
  EditPicture
  
  ListView1.SetFocus
    ' Select the next picture.
  If Not ListView1.SelectedItem Is Nothing Then
    ListView1.ListItems(ListView1.SelectedItem.key).Selected = True
    ListView1.SelectedItem.EnsureVisible
    ListView1.SetFocus
  End If

  RefreshListView

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
  
    Case vbKeyInsert
      NewPicture
      ListView1.SetFocus
    
    Case vbKeySpace
      If Not ListView1.SelectedItem Is Nothing Then
        ListView1.SetFocus
        ListView1.StartLabelEdit
      End If
    
    Case vbKeyReturn
      EditPicture
      ListView1.SetFocus
  
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
  
  ' Refesh the toolbar.
  frmSysMgr.RefreshMenu

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long
  
  ' Ensure an item is selected.
  If ListView1_SelectedCount = 0 Then
    If ListView1.ListItems.Count > 0 Then
      ListView1.ListItems(ListView1.SelectedItem.key).Selected = True
    End If
  End If
  
  ' Display the ActiveBar popup menu when the right button is clicked.
  If Button = vbRightButton Then
    UI.GetMousePos lXMouse, lYMouse
'    frmSysMgr.tbMain.PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
    frmSysMgr.tbMain.Bands("ID_mnuEdit").TrackPopup -1, -1
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar
  
  ' Refresh the menu as we may heve one, or more pictures selected.
  frmSysMgr.RefreshMenu
  
End Sub

Public Sub EditMenu(ByVal MenuItem As String)
  
  Select Case MenuItem
    
    Case "ID_New"
      gfMenuActionKey = True
      NewPicture
    
    Case "ID_Delete"
      gfMenuActionKey = True
      DeletePictures
    
    Case "ID_Properties"
      gfMenuActionKey = True
      EditPicture
      
    Case "ID_SelectAll"
      gfMenuActionKey = True
      ListView1_SelectAll
  
    Case "ID_SaveChanges"
      gfMenuActionKey = True
      Pict_SaveChanges
    
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      ChangeView 0 'lvwIcon

    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      ChangeView 1 'lvwSmallIcon
  
    Case "ID_List"
      ' Change the view to display a list.
      ChangeView 2 'lvwList

    Case "ID_Details"
      ' Change the view to display details.
      ChangeView 3 'lvwReport
  
    Case "ID_CustomiseColumns"
      ' Customise column view...
      Set frmShowColumns = New HRProSystemMgr.frmShowColumns
      frmShowColumns.PropertySet = gpropShowColumns_PictMgr
      frmShowColumns.Show vbModal
      SetColumnSizes
      Exit Sub
  
  End Select
  
End Sub


Private Sub Pict_SaveChanges()

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
      
Private Function EditPicture() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sKey As String
    
  fOK = Not ListView1.SelectedItem Is Nothing
  
  If fOK Then
  
    sKey = ListView1.SelectedItem.key
    
    ' Locate the relevant database record.
    recPictEdit.Index = "idxID"
    recPictEdit.Seek "=", val(Mid(sKey, 2))
    
    If Not recPictEdit.NoMatch Then
      
      With frmPictEdit
        
        ' Get the picture from the associated imagelist control.
        Set .PictureObj = ImageList2.ListImages(sKey).Picture
        .PictureName = recPictEdit.Fields("name")
        
        ' Display the picture edit screen.
        .Show vbModal
        
        ' Write the modifications to the database.
        fOK = Not .Cancelled
        
        If fOK Then
          fOK = EditPicture_Transaction(val(Mid(sKey, 2)), .PictureName)
        End If
        
        If fOK Then
          ' Write the modifications to the listview.
          ListView1.ListItems(sKey).Text = .PictureName
          EditPicture = True
          RefreshListView
        End If
      
      End With
    End If
  End If
  
TidyUpAndExit:
  EditPicture = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub PopulateListView()
  Dim strKey As String
  Dim ThisItem As ComctlLib.ListItem
  Dim strName As String
  Dim iHeight As Integer
  Dim iWidth As Integer
  
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the frmPicMgr form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd
  
  ' Clear all items from the listview.
  ListView1.ListItems.Clear
  
  With recPictEdit
  
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      ' Loop through the picture records, adding an item to the listview for each picture.
      Do While Not .EOF
      
        If Not .Fields("deleted") Then
        
          strKey = "I" & Trim(Str(.Fields("pictureID")))
          strName = Trim(.Fields("name"))
          
          iHeight = Int(Me.ScaleY(ImageList2.ListImages(strKey).Picture.Height, vbHimetric, vbPixels))
          iWidth = Int(Me.ScaleX(ImageList2.ListImages(strKey).Picture.Width, vbHimetric, vbPixels))
          
          Set ThisItem = ListView1.ListItems.Add(, strKey, _
            strName, ImageList2.ListImages(strKey).Index, ImageList1.ListImages(strKey).Index)
          ThisItem.SubItems(1) = Choose(.Fields("PictureType"), "Bitmap", "Metafile", "Icon")
          ThisItem.SubItems(2) = Trim(Str(iHeight))
          ThisItem.SubItems(3) = Trim(Str(iWidth))

          Set ThisItem = Nothing
            
        End If
        
        .MoveNext
      
      Loop
      
      .MoveFirst
    
    End If
    
  End With
  
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  
  ' Ensure the selected item is visible.
  If Not ListView1.SelectedItem Is Nothing Then
    ListView1.ListItems(ListView1.SelectedItem.key).Selected = True
    ListView1.SelectedItem.EnsureVisible
  End If
    
  If Not Loading Then
    RefreshListView
  End If
  
End Sub

Private Function NewPicture() As Boolean
  ' Add a new picture to the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iWidth As Integer
  Dim iHeight As Integer
  Dim lngNewID As Long
  Dim sKey As String
  Dim sFileName As String
  Dim sPictureName As String
  Dim objItem As ComctlLib.ListItem
  Dim sErrorMessage As String
  
  sFileName = GetPictureFile("Open Picture File")
  
  ' If a new picture has been selected then add it to the listview,
  ' and the database.
  fOK = (Len(sFileName) > 0)
  If fOK Then
    sPictureName = JustFileName(sFileName)
    lngNewID = Database.UniqueColumnValue("tmpPictures", "pictureID")
    sKey = "I" & lngNewID
    
    ' Add the new picture to the small imagelist.
    ImageList1.ListImages.Add , sKey, LoadPicture(sFileName)
    ' Add the new picture to the large imagelist.
    ImageList2.ListImages.Add , sKey, LoadPicture(sFileName)
    
    fOK = AddPicture_Transaction(lngNewID, sFileName)
  End If
    
  If fOK Then
    ' Add the new picture to the listview.
    Set objItem = ListView1.ListItems.Add(, sKey, sPictureName, ImageList2.ListImages(sKey).Index, ImageList1.ListImages(sKey).Index)
    objItem.SubItems(1) = Choose(recPictEdit.Fields("PictureType"), "Bitmap", "Metafile", "Icon")
    iHeight = Int(Me.ScaleY(ImageList2.ListImages(sKey).Picture.Height, vbHimetric, vbPixels))
    iWidth = Int(Me.ScaleX(ImageList2.ListImages(sKey).Picture.Width, vbHimetric, vbPixels))
    objItem.SubItems(2) = Trim(Str(iHeight))
    objItem.SubItems(3) = Trim(Str(iWidth))
    Set objItem = Nothing
      
    ' Deselect all other controls.
    ListView1_ClearSelections
    ' Ensure the new picture is selected in the list view
    ListView1.ListItems(sKey).Selected = True
    ListView1.SelectedItem.EnsureVisible
    ListView1.SetFocus
      
    Me.Changed = True
  
    RefreshListView
  End If
  
TidyUpAndExit:
  Set objItem = Nothing
  NewPicture = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  'JPD 20050208 Fault 9787
  If Err.Number = 50003 Then
    sErrorMessage = "Invalid picture"
  Else
    sErrorMessage = Err.Description
  End If
  
  MsgBox "Error adding picture." & vbCrLf & vbCrLf & sErrorMessage, vbOKOnly + vbExclamation, App.Title
  
  Resume TidyUpAndExit

End Function


Public Function AddPicture_Transaction(plngPictureID As Long, psFileName As String) As Boolean
  ' Add the new picture's details to the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iPictureType As Integer
  Dim sPictureName As String
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  iPictureType = ImageList2.ListImages("I" & plngPictureID).Picture.Type
  sPictureName = JustFileName(psFileName)
    
  With recPictEdit
    .AddNew
    .Fields("pictureID") = plngPictureID
    .Fields("pictureType") = iPictureType
    .Fields("name") = sPictureName
    .Fields("new") = True
    .Fields("changed") = False
    .Fields("deleted") = False
  End With

  fOK = WritePicture(psFileName)
        
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  AddPicture_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function EditPicture_Transaction(plngPictureID As Long, psName As String) As Boolean
  ' Write the modified picture's details to the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  With recPictEdit
    .Index = "idxID"
    .Seek "=", plngPictureID
        
    fOK = Not .NoMatch
    
    If fOK Then
      .Edit
      !Name = psName
      !Changed = True
      .Update
    End If
  End With
   
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
    Me.Changed = True
  Else
    daoWS.Rollback
  End If
  EditPicture_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function GetPictureFile(Optional ByVal Title As String) As String
  ' Display a dialogue box for the user to select a bitmap or icon file.
  ' Trap the error caused when the dialogue box is cancelled.
  On Error GoTo ErrorTrap
  
  With comDlgBox
      
    If Len(Trim(Title)) > 0 Then
      .DialogTitle = Trim(Title)
    End If
    .Filter = "Pictures (*.bmp;*.ico;*.jpg;*.gif)|*.bmp;*.ico;*.jpg;*.gif"
    .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNExplorer

    ' Display the dialogue box.
    .ShowOpen
    
    ' Read the font properties of the dialogue box.
    If Len(Trim(.FileName)) > 0 Then
      GetPictureFile = Trim(.FileName)
    Else
      GetPictureFile = vbNullString
    End If
  
  End With
  
  Exit Function
  
ErrorTrap:
  ' User pressed cancel.
  GetPictureFile = vbNullString
  
End Function


Private Sub RefreshListView()

  ListView1.Sorted = True
  ListView1.Refresh
    
  RefreshStatusBar
  
  frmSysMgr.RefreshMenu

End Sub




Private Function DeletePictures() As Boolean
  ' Delete the selected pictures.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  
  fDeleteAll = False
  
  fOK = True
  
  ' If we have more than one selection then question the multi-deletion.
  If ListView1_SelectedCount > 1 Then
  
    If MsgBox("Are you sure you want to delete these " & _
      Trim(Str(ListView1_SelectedCount)) & _
      " pictures ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
    
      fDeleteAll = True
      fConfirmed = True
    Else
      ListView1.SetFocus
      Exit Function
    End If
    
  End If
  
  ' Change the mouse pointer.
  Screen.MousePointer = vbHourglass
  ' Lock the frmPicMgr form to avoid messy screen refresh.
  UI.LockWindow Me.hWnd

  ' Loop through the list view items deleting all of those
  ' currently selected.
  iLoop = 1
  Do While iLoop <= ListView1.ListItems.Count
  
    If ListView1.ListItems(iLoop).Selected = True Then
    
      With recPictEdit
        .Index = "idxID"
        .Seek "=", val(Mid(ListView1.ListItems(iLoop).key, 2))
    
        If Not .NoMatch Then
          If Not fDeleteAll Then
      
            If MsgBox("Are you sure you want to delete " & .Fields("Name") & _
              " ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
          
              fConfirmed = True
            Else
              fConfirmed = False
            End If
            
          End If
          
          ' Remove the picture from the database and the listview..
          If fConfirmed Then
            fOK = DeletePicture_Transaction(val(Mid(ListView1.ListItems(iLoop).key, 2)))
            
            If fOK Then
              ListView1.ListItems.Remove ListView1.ListItems(iLoop).key
              iLoop = iLoop - 1
            End If
          End If
        End If
      End With
    End If
    
    iLoop = iLoop + 1
  Loop
  
  ' Unlock the frmPicMgr form to show the updated listview.
  UI.UnlockWindow
  ' Reset the mousepointer.
  Screen.MousePointer = vbNormal
  
  ' Select the next picture.
  If Not ListView1.SelectedItem Is Nothing Then
    ListView1.ListItems(ListView1.SelectedItem.key).Selected = True
    ListView1.SelectedItem.EnsureVisible
    ListView1.SetFocus
  End If
          
  RefreshListView

TidyUpAndExit:
  DeletePictures = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function PictureIsUsed(glngPictureID As Long) As Boolean
 
  ' Return true if the picture is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Pictures may be used in the following contexts :
  '
  '   As an icon for a screen.
  '   As an image control on a screen.
  '   HR Pro background
  On Error GoTo ErrorTrap
  
  Dim fUsed As Boolean
  Dim sSQL As String
  Dim rsScreens As DAO.Recordset
  Dim sPictureName As String
  Dim sScreenName As String
  Dim sTableName As String
  
  fUsed = False
  
  recPictEdit.Index = "idxID"
  recPictEdit.Seek "=", glngPictureID
      
  If Not recPictEdit.NoMatch Then
    sPictureName = recPictEdit!Name
  Else
    sPictureName = "<unknown>"
  End If
  
  ' Find any screens that use this picture as its icon.
  sSQL = "SELECT DISTINCT tmpScreens.name, tmpScreens.tableID" & _
    " FROM tmpScreens" & _
    " WHERE deleted=FALSE" & _
    " AND tmpScreens.pictureID=" & Trim(Str(glngPictureID))
  Set rsScreens = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsScreens.BOF And rsScreens.EOF) Then
    fUsed = True
    Do Until rsScreens.EOF
        ' Get the names of the screen and its associated table.
      sScreenName = rsScreens.Fields("name")
        
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", rsScreens.Fields("tableId")
        
      If Not recTabEdit.NoMatch Then
        sTableName = recTabEdit!TableName
      Else
        sTableName = "<unknown>"
      End If
      
      mfrmUse.AddToList ("Screen Icon : " & sScreenName & " <" & sTableName & ">")
      rsScreens.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsScreens.Close
  
  ' Check that it is not used as an image control on a screen.
  sSQL = "SELECT DISTINCT tmpScreens.name, tmpScreens.tableID" & _
    " FROM tmpScreens, tmpControls" & _
    " WHERE tmpControls.pictureID=" & Trim(Str(glngPictureID)) & _
    " AND tmpControls.screenID = tmpScreens.screenID"
  Set rsScreens = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsScreens.BOF And rsScreens.EOF) Then
    fUsed = True
    Do Until rsScreens.EOF
      ' Get the names of the screen and its associated table.
      sScreenName = rsScreens.Fields("name")
        
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", rsScreens.Fields("tableId")
        
      If Not recTabEdit.NoMatch Then
        sTableName = recTabEdit!TableName
      Else
        sTableName = "<unknown>"
      End If
      
      mfrmUse.AddToList ("Screen Image : " & sScreenName & " <" & sTableName & ">")
      rsScreens.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsScreens.Close
  
  ' Check that the picture is not used as the HR Pro background.
  If glngPictureID = glngDesktopBitmapID Then
    fUsed = True
    mfrmUse.AddToList ("HR Pro Background Image")
  End If

  
  If Application.WorkflowModule Then
    ' Check that it is not used as a background picture on a workflow web form.
    sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
      "   tmpWorkflowElements.identifier" & _
      " FROM tmpWorkflowElements" & _
      " WHERE tmpWorkflowElements.webFormBGImageID = " & Trim(Str(glngPictureID))
  
    Set rsScreens = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsScreens.BOF And rsScreens.EOF) Then
      Do Until rsScreens.EOF
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", rsScreens.Fields("workflowID")
  
        If Not recWorkflowEdit.NoMatch Then
          If recWorkflowEdit.Fields("deleted").value = False Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsScreens.Fields("identifier") & "' web form background picture>")
          End If
        End If
        
        rsScreens.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsScreens.Close
      
    ' Check that it is not used as a picture on a workflow web form.
    sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
      "   tmpWorkflowElements.identifier" & _
      " FROM tmpWorkflowElementItems" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.id" & _
      " WHERE tmpWorkflowElementItems.pictureID = " & Trim(Str(glngPictureID))
  
    Set rsScreens = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsScreens.BOF And rsScreens.EOF) Then
      Do Until rsScreens.EOF
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", rsScreens.Fields("workflowID")
            
        If Not recWorkflowEdit.NoMatch Then
          If recWorkflowEdit.Fields("deleted").value = False Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsScreens.Fields("identifier") & "' web form picture>")
          End If
        End If
        
        rsScreens.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsScreens.Close
  End If
      
  ' NPG20100427 Fault HRPRO-891
  If Application.SelfServiceIntranetModule Then
    ' Check that it is not used as a separator picture in the SSI.
    sSQL = "SELECT DISTINCT tmpSSIntranetLinks.ID" & _
      " FROM tmpSSIntranetLinks" & _
      " WHERE tmpSSIntranetLinks.PictureID = " & Trim(Str(glngPictureID))
  
    Set rsScreens = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsScreens.BOF And rsScreens.EOF) Then
            fUsed = True
            mfrmUse.AddToList ("Self Service Intranet Module Setup")
    End If
    'Close temporary recordset
    rsScreens.Close
  End If
      
      
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsScreens = Nothing
  PictureIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function





Private Sub RefreshStatusBar()
  Dim iItems As Integer
  Dim iSelections As Integer
  Dim sMessage As String
  
  iItems = ListView1.ListItems.Count
  iSelections = ListView1_SelectedCount
  
  sMessage = Trim(Str(iItems)) & " picture"
  If iItems <> 1 Then
    sMessage = sMessage & "s"
  End If
  sMessage = sMessage & ", " & Trim(Str(iSelections)) & " picture"
  If iSelections <> 1 Then
    sMessage = sMessage & "s"
  End If
  sMessage = sMessage & " selected."

  StatusBar1.Panels(1).Text = sMessage
  
End Sub



Public Function DeletePicture_Transaction(plngPictureID As Long) As Boolean
  ' Delete the given picture.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  With recPictEdit
    .Index = "idxID"
    .Seek "=", plngPictureID
    
'    'TM19012004
'    ' Check that the picture is not being used before deleting it.
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If PictureIsUsed(.Fields("pictureID")) Then
      Screen.MousePointer = vbNormal
      mfrmUse.ShowMessage !Name & " Picture", "The picture cannot be deleted as the picture is used by the following:", UsageCheckObject.Picture
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
    
    If fOK Then
      .Edit
      .Fields("deleted") = True
      .Update
    End If
    
  End With
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
    Me.Changed = True
  Else
    daoWS.Rollback
  End If
  DeletePicture_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ChangeView(ViewStyle As ComctlLib.ListViewConstants) As Boolean

  Static InChangeView As Boolean
  On Error GoTo ErrorTrap

  If InChangeView Then Exit Function

  InChangeView = True

'TM20010914 Fault 1753
'As ActiveBar does not support mutual exclusivity on its tools, the following code
'ensures only one of the view options is selected at any one time.

  Me.ListView1.View = ViewStyle
  With abPictMgrMenu
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
    
    End Select
  End With

  InChangeView = False

  ChangeView = True

  Exit Function

ErrorTrap:
  ChangeView = False
  Err = False

End Function

Private Sub SetColumnSizes()

  Dim iCount As Integer
 
  With Me.ListView1
    For iCount = 2 To .ColumnHeaders.Count Step 1
    
      If gpropShowColumns_PictMgr(.ColumnHeaders.Item(iCount).Text) = True Then
        .ColumnHeaders(iCount).Width = (Len(.ColumnHeaders(iCount).Text) + 1) * UI.GetAvgCharWidth(Me.hDC)
      Else
        .ColumnHeaders(iCount).Width = 0
      End If
    Next iCount
  End With
  
End Sub

