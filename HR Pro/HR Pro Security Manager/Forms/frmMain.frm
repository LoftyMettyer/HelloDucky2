VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "CODEJO~2.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00F7EEE9&
   Caption         =   "OpenHR - Security Manager"
   ClientHeight    =   1710
   ClientLeft      =   2550
   ClientTop       =   2820
   ClientWidth     =   2760
   HelpContextID   =   8001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   90
      Top             =   1185
   End
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   2730
      TabIndex        =   0
      Top             =   0
      Width           =   2760
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   480
         ScaleHeight     =   1170
         ScaleWidth      =   1170
         TabIndex        =   1
         Top             =   195
         Width           =   1200
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2040
      Top             =   960
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin ActiveBarLibraryCtl.ActiveBar abSecurity 
      Left            =   1095
      Top             =   795
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
      Bands           =   "frmMain.frx":058A
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Windows API call used to control textbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Edit Control Messages
Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_UNDO = &H304
Const EM_CANUNDO = &HC6
Const EM_GETMODIFY = &HB8

Private mbLoading As Boolean

Dim mobjSecurity As SecurityGroups

Public gfrmCurrentForm As Form
Private mblnReadOnly As Boolean

' Functions to tile the background image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal lDC As Long, ByVal hObject As Long) As Long
Dim pic As StdPicture, hMemDC As Long, pHeight As Long, pWidth As Long

Private mfLogoffCancelled As Boolean
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
    sFileName = GetPictureFromDatabase(glngDesktopBitmapID)
    picWork.Picture = LoadPicture(sFileName)
    Kill sFileName
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
Private Sub EditPerform(piEditFunction As Integer)
   
If TypeOf Screen.ActiveControl Is TextBox Then
  Call SendMessage(Screen.ActiveControl.hWnd, piEditFunction, 0, 0&)
End If
   
End Sub

Private Sub abSecurity_MenuItemEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  DoEvents
End Sub

Private Sub abSecurity_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub abSecurity_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  Dim i As Integer
  Dim strVersionFilename As String
  Dim sCurrentGroup As String

  ' Get rid of any screen display residue.
  DoEvents

  Select Case Tool.Name
    
    ' Module menu.

    Case "ID_Audit"
      For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmAudit Then
          Forms(i).Visible = True
          Forms(i).SetFocus
          If Forms(i).WindowState = vbMinimized Then Forms(i).WindowState = vbNormal
          Exit Sub
        End If
      Next i
      Screen.MousePointer = vbHourglass
      frmAudit.Show
      Screen.MousePointer = vbDefault

    Case "ID_Group"
      For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmGroupMaint1 Then
          Forms(i).Visible = True
          Forms(i).SetFocus
          Exit Sub
        End If
      Next i
      Set gfrmCurrentForm = New frmGroupMaint1
      gfrmCurrentForm.Show

    Case "ID_SecuritySave"
     
'MH20060905 Fault
'      If Not Me.ActiveForm Is Nothing Then
'        If TypeOf Me.ActiveForm Is frmGroupMaint1 Then
'          Me.ActiveForm.EditMenu "ID_SecuritySave"
'        End If
'      ElseIf Application.Changed Then
'        If MsgBox("Save changes." & vbCrLf & _
'                  "Are you sure ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'          If ApplyChanges Then
'            Application.Changed = False
'            abSecurity.Bands("bndModule").Tools("ID_SecuritySave").Enabled = False
'          End If
'
'        End If
'      End If
      Dim blnGroupMaint1 As Boolean
      
      blnGroupMaint1 = False
      If Not Me.ActiveForm Is Nothing Then
        blnGroupMaint1 = (TypeOf Me.ActiveForm Is frmGroupMaint1)
      End If
      
      If blnGroupMaint1 Then
        Me.ActiveForm.EditMenu "ID_SecuritySave"
      
      ElseIf Application.Changed Then
        If MsgBox("Save changes." & vbCrLf & _
                  "Are you sure ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          If ApplyChanges Then
            Application.Changed = False
            abSecurity.Bands("bndModule").Tools("ID_SecuritySave").Enabled = False
          End If

        End If
      End If


    Case "ID_LogOff"
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If MsgBox("Are you sure you wish to log off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
            'Looks like we want to log off so do the necessary.
            'TM20010920 Fault 2530
            blnIsExiting = True
            mfLogoffCancelled = False
            
            Unload Me
            'RH 12/10/00 - Now call this from frmMain_QueryUnload
            'Call AuditAccess(False, "Security")
            
            ' JPD20030228 Fault 5094
            If Not mfLogoffCancelled Then
              Application.Logout
              Main
            End If
        End If

    Case "ID_Exit"
      'TM20010920 Fault 2530
      blnIsExiting = True
      
      'RH 12/10/00 - Now call this from frmMain_QueryUnload
      'Call AuditAccess(False, "Security")
      Unload Me

    ' Audit Menu.

    Case "ID_AuditOpen"
     ActiveForm.EditMenu Tool.Name

    Case "ID_AuditDelete"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditPrint"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditShowColumns"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditSort"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditSetFilter"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditClearFilter"
      ActiveForm.EditMenu Tool.Name

    Case "ID_AuditRefresh"
      ActiveForm.EditMenu Tool.Name
    
    Case "ID_AuditScheduleTasks"
      ActiveForm.EditMenu Tool.Name

    ' Security Edit Menu.

    Case "ID_SecurityNew"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityAutomaticAdd"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityCopy"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityMove"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityDelete"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityProperties"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecuritySelectAll"
      ActiveForm.EditMenu Tool.Name

    Case "ID_SecurityPrint"
      ActiveForm.EditMenu Tool.Name
      
    Case "ID_UnCheckAll"
      ActiveForm.EditMenu Tool.Name
      
    Case "ID_CheckAll"
      ActiveForm.EditMenu Tool.Name
        
    Case "ID_SecurityResetPassword"
      ActiveForm.EditMenu Tool.Name
        
    Case "ID_FindUser"
      ActiveForm.EditMenu Tool.Name
        
    ' View menu.
    Case "ID_LargeIcons"
      ActiveForm.ChangeView lvwIcon

    Case "ID_SmallIcons"
      ActiveForm.ChangeView lvwSmallIcon

    Case "ID_List"
      ActiveForm.ChangeView lvwList

    Case "ID_Details"
      ActiveForm.ChangeView lvwReport

    ' Tools menu.

    Case "ID_LicenceInfo"
      frmLicence.Show vbModal
      Set frmLicence = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If

    Case "ID_PasswordMaintenance"
      frmPasswordMaintenance.ShowAllUsers = True
      If frmPasswordMaintenance.Initialise Then
        frmPasswordMaintenance.Show vbModal
      End If
      Set frmPasswordMaintenance = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If

    Case "ID_UtilityOwnership"
      If frmUtilityOwnership.Initialise Then
        frmUtilityOwnership.Show vbModal
      End If
    
    Case "ID_SecurityOptions"
      If frmSecurityOptions.Initialise Then
        frmSecurityOptions.Show vbModal
      End If

    ' Window menu.
    
    Case "ID_Cascade"
      ' Cascade the windows.
      CascadeForms

    Case "ID_TileH"
      frmMain.Arrange vbTileHorizontal

    Case "ID_TileV"
      frmMain.Arrange vbTileVertical

    Case "ID_ArrangeIcons"
      ' Arrange window icons
      frmMain.Arrange vbArrangeIcons

    Case "ID_Maximise"
      ' Maximise the current window.
      ActiveForm.WindowState = vbMaximized

    Case "ID_Minimise"
      ' Minimise the current window.
      ActiveForm.WindowState = vbMinimized

    Case "ID_Restore"
      ' Restore the current window.
      ActiveForm.WindowState = vbNormal

    Case "ID_Close"
      ' Close the active module.
      Unload ActiveForm

    '
    ' Help menu.
    '
    Case "ID_ContentsandIndex"
    
      Dim plngHelp As Long
      
      If Not ShowAirHelp(0) Then
        plngHelp = ShellExecute(0&, vbNullString, App.Path & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to find and view the file " & App.HelpFile & ".", vbExclamation + vbOKOnly, App.EXEName
        End If
      End If

    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = App.Path & "\OpenHR Security Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          MsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox "No version information found.", vbExclamation + vbOKOnly, "OpenHR Security Manager"
      End If
      
      Screen.MousePointer = vbNormal


    Case "ID_ViewCurrentUsers"
      frmViewCurrentUsers.Saving = False
      frmViewCurrentUsers.Show vbModal
      Unload frmViewCurrentUsers
      Set frmViewCurrentUsers = Nothing

    Case "ID_About"
      Load frmAbout
      DoEvents     ' Needed to prevent grey square appearing when the 'about' form displays.
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If

  End Select

End Sub

Private Sub MDIForm_Activate()
  RefreshMenu False

  ' NPG20091007 Fault 416
  ' set the new multi-size icons for taskbar, application, and alt-tab
  ' NB Only works when run as the executable
  SetIcon Me.hWnd, "!ABS", True

End Sub

Private Sub MDIForm_Load()
  ' Load the CodeJock Styles
  Call LoadSkin(Me, Me.SkinFramework1)

  mblnReadOnly = (Application.AccessMode <> accFull)
  SetCaption
    
  With abSecurity
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
  
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  blnIsExiting = True
  
  'Prompt the user to apply changes, if changes have been made.
  If Application.Changed Then
'    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Application.Name)
    Select Case MsgBox("Save all changes ?", vbYesNoCancel + vbQuestion, Application.Name)
      Case vbCancel
        Cancel = True
        RefreshMenu False
      Case vbYes
        Cancel = Not ApplyChanges
    End Select
  End If
             
  ' JPD20030228 Fault 5094
  mfLogoffCancelled = Cancel

  'Remove Progress Bar class from memory
  If Cancel = False Then
      
    ' JDM - 07/02/2006 - Fault 10780 - Close all active forms otherwise we get
    '                                  problems unloading the forms after the mdiform has been called
    Do While Forms.Count > 1
      Unload Forms(1)
      DoEvents
    Loop
      
    'MH20060126 Fault 10413
    abSecurity.Tools.RemoveAll
    abSecurity.Bands.RemoveAll
    abSecurity.ReleaseFocus

    Set gobjProgress = Nothing

    If Not gADOCon Is Nothing Then
      If Application.AccessMode = accFull Then
        'rdoCon.BeginTrans
        'UnlockDatabase ("Lock Read Write")
        'rdoCon.CommitTrans
        UnlockDatabase lckReadWrite
      End If
    End If

    Call AuditAccess("Log Out", "Security")

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


Public Sub RefreshMenuAudit(pfUnloadingForm As Boolean)
  ' Configure the Audit specific menu options.
    If TypeOf Screen.ActiveForm Is frmAudit Then
      abSecurity.Bands("mnuSecurity").Tools("ID_mnuAudit").Visible = True
      abSecurity.Bands("bndAudit").Tools("ID_AuditClearFilter").Enabled = Screen.ActiveForm.Filtered
    End If
    
End Sub


Public Sub RefreshMenuGroup(pfUnloadingForm As Boolean)
  ' Configure the Security specific menu options.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim fEnableSelectAll As Boolean
  Dim fEnableViews As Boolean
  Dim fEnablePrintDetails As Boolean
  Dim sRightView As String
    
  fOK = True
        
  With abSecurity.Bands("mnuSecurity")
    .Tools("ID_mnuSecurityEdit").Visible = True
    .Tools("ID_mnuView").Visible = True
    abSecurity.Bands("bndModule").Tools("ID_SecuritySave").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecuritySave").Enabled
    .Refresh
  End With
 
    '==================================================
    ' Configure the Edit menu tools.
    '==================================================
    ' Enable/disable the required tools.
    fEnableSelectAll = False
    fEnablePrintDetails = False
      
    With ActiveForm
  
      If .ActiveView Is .trvConsole Then
        If Not .trvConsole.SelectedItem Is Nothing Then
          Select Case .trvConsole.SelectedItem.DataKey
            Case "GROUPS"
              fEnableSelectAll = (.lvList.ListItems.Count > 0)
              fEnablePrintDetails = (.lvList.ListItems.Count > 0)
            Case "GROUP"
              fEnablePrintDetails = (.lvList.ListItems.Count > 0)
            Case "USERS"
              fEnableSelectAll = (.lvList.ListItems.Count > 0)
              fEnablePrintDetails = True ' (.lvList.ListItems.Count > 0)
            Case "TABLESVIEWS"
              fEnableSelectAll = (.lvList.ListItems.Count > 0)
              fEnablePrintDetails = True ' (.lvList.ListItems.Count > 0)
            Case "TABLE"
            Case "VIEW"
            Case "SYSTEM"
              fEnablePrintDetails = True '(.lvList.ListItems.Count > 0)
          End Select
        
          sRightView = .trvConsole.SelectedItem.DataKey
        End If
      Else
        If .ActiveView Is .lvList Then
          If Not .trvConsole.SelectedItem Is Nothing Then
            Select Case .trvConsole.SelectedItem.DataKey
              Case "GROUPS"
                fEnableSelectAll = (.lvList.ListItems.Count > 0)
                fEnablePrintDetails = (.lvList.ListItems.Count > 0)
              Case "GROUP"
                'fEnablePrintDetails = (.lvList.ListItems.Count > 0)
                fEnablePrintDetails = True '(.lvList.SelectedItem.Text = "User Logins")
              Case "USERS"
                fEnableSelectAll = (.lvList.ListItems.Count > 0)
                fEnablePrintDetails = False '(.lvList.ListItems.Count > 0)
              Case "TABLESVIEWS"
                fEnableSelectAll = (.lvList.ListItems.Count > 0)
            End Select
          End If
          
          sRightView = .trvConsole.SelectedItem.DataKey
          
        End If
      End If
    End With
    
    With abSecurity.Bands("bndEdit_Left")
      
      .Tools("ID_SecurityNew").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityNew").Enabled
      .Tools("ID_SecurityAutomaticAdd").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityAutomaticAdd").Enabled
      .Tools("ID_SecurityDelete").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityDelete").Enabled
      .Tools("ID_SecurityCopy").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityCopy").Enabled
      .Tools("ID_SecurityProperties").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityProperties").Enabled
      
      'NHRD10062003 Fault 4947
      .Tools("ID_SecurityPrint").Enabled = fEnablePrintDetails
      
    End With
    
    With abSecurity.Bands("bndEdit_Right")
      .Tools("ID_SecurityMove").Visible = (sRightView = "USERS")
      .Tools("ID_SecurityMove").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityMove").Enabled
      .Tools("ID_SecurityResetPassword").Visible = (sRightView = "USERS")
      .Tools("ID_SecurityResetPassword").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityResetPassword").Enabled
      .Tools("ID_SecurityAutomaticAdd").Visible = (sRightView = "USERS") And ActiveForm.abGroupMaint.Tools("ID_SecurityAutomaticAdd").Enabled
      .Tools("ID_CheckAll").Visible = ActiveForm.abGroupMaint.Tools("ID_CheckAll").Enabled
      .Tools("ID_UnCheckAll").Visible = ActiveForm.abGroupMaint.Tools("ID_UnCheckAll").Enabled
      .Tools("ID_SecuritySelectAll").Visible = fEnableSelectAll
      .Tools("ID_SecurityProperties").Enabled = ActiveForm.abGroupMaint.Tools("ID_SecurityProperties").Enabled
    End With
  
  
    '==================================================
    ' Configure the View menu.
    '==================================================
    With abSecurity.Bands("bndView")
    .Tools("ID_LargeIcons").Enabled = ActiveForm.abGroupMaint.Tools("ID_LargeIcons").Enabled
    .Tools("ID_SmallIcons").Enabled = ActiveForm.abGroupMaint.Tools("ID_SmallIcons").Enabled
    .Tools("ID_List").Enabled = ActiveForm.abGroupMaint.Tools("ID_List").Enabled
    .Tools("ID_Details").Enabled = ActiveForm.abGroupMaint.Tools("ID_Details").Enabled
      
      .Tools("ID_LargeIcons").Checked = False
      .Tools("ID_SmallIcons").Checked = False
      .Tools("ID_List").Checked = False
      .Tools("ID_Details").Checked = False
      
    If ActiveForm.abGroupMaint.Tools("ID_LargeIcons").Checked = True Then
      .Tools("ID_LargeIcons").Checked = True
    End If
    If ActiveForm.abGroupMaint.Tools("ID_SmallIcons").Checked = True Then
      .Tools("ID_SmallIcons").Checked = True
    End If
    If ActiveForm.abGroupMaint.Tools("ID_List").Checked = True Then
      .Tools("ID_List").Checked = True
    End If
    If ActiveForm.abGroupMaint.Tools("ID_Details").Checked = True Then
      .Tools("ID_Details").Checked = True
    End If
      
  
    .Refresh
    
  End With
    
  abSecurity.RecalcLayout
  abSecurity.Refresh
    
TidyUpAndExit:
  Exit Sub
    
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Public Sub RefreshMenu(pfUnloadingForm As Boolean)
  'Refresh the menu and tool bars.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim iFormCount As Integer
  Dim frmTemp As Form
  
  fOK = True
    
  iFormCount = Forms.Count - IIf(pfUnloadingForm, 1, 0)
  
  abSecurity.Bands("bndTools").Tools("ID_UtilityOwnership").Enabled = Not mblnReadOnly
  
  'MH20060915 Fault 11068
  'abSecurity.Bands("bndTools").Tools("ID_PasswordMaintenance").Enabled = Not mblnReadOnly
  abSecurity.Bands("bndTools").Tools("ID_PasswordMaintenance").Enabled = (gbUserCanManageLogins And Not mblnReadOnly)

  'Configure the Module Menu tools - these are the same whatever screen is active
  
  Set frmTemp = Screen.ActiveForm
  
  ' NPG20081201 Fault 13401
  ' abSecurity.Bands("bndModule").Tools("ID_Audit").Enabled = (iFormCount <= 1) Or (Not TypeOf Screen.ActiveForm Is frmAudit)
  ' abSecurity.Bands("bndModule").Tools("ID_Group").Enabled = (iFormCount <= 1) Or (Not TypeOf Screen.ActiveForm Is frmGroupMaint1)
  If Not frmTemp Is Nothing Then
    abSecurity.Bands("bndModule").Tools("ID_Audit").Enabled = (iFormCount <= 1) Or (Not TypeOf Screen.ActiveForm Is frmAudit)
    abSecurity.Bands("bndModule").Tools("ID_Group").Enabled = (iFormCount <= 1) Or (Not TypeOf Screen.ActiveForm Is frmGroupMaint1)
  Else
    abSecurity.Bands("bndModule").Tools("ID_Audit").Enabled = (iFormCount <= 1)
    abSecurity.Bands("bndModule").Tools("ID_Group").Enabled = (iFormCount <= 1)
  End If
  
  
  'Hide menus as default
  abSecurity.Bands("mnuSecurity").Tools("ID_mnuAudit").Visible = False
  abSecurity.Bands("mnuSecurity").Tools("ID_mnuSecurityEdit").Visible = False
  abSecurity.Bands("mnuSecurity").Tools("ID_mnuView").Visible = False
  
  ' Refresh the menus that are specific to the active child form.
  If iFormCount > 1 And Not frmTemp Is Nothing Then
    If TypeOf Screen.ActiveForm Is frmAudit Then
      RefreshMenuAudit pfUnloadingForm
    ElseIf TypeOf Screen.ActiveForm Is frmGroupMaint1 Then
      RefreshMenuGroup pfUnloadingForm
    End If
  End If

  'MH20060810 Fault 11417
  'If gbShiftSave Then
  
  ' NPG20081201 Fault 13401
  'If Application.Changed Then
    abSecurity.Bands("bndModule").Tools("ID_SecuritySave").Enabled = Application.Changed
  'End If
    
  'RefreshMenuTools
  RefreshMenuWindow pfUnloadingForm
    
  abSecurity.RecalcLayout
  abSecurity.Refresh
      
TidyUpAndExit:
  Exit Sub
    
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Private Sub RefreshMenuWindow(pfUnloadingForm As Boolean)
  ' Configure the Window menu options.
  Dim iFormCount As Integer
  
  iFormCount = Forms.Count - IIf(pfUnloadingForm, 1, 0)
  
  With abSecurity.Bands("bndWindow")
    .Tools("ID_Cascade").Enabled = (iFormCount > 1)
    .Tools("ID_TileV").Enabled = (iFormCount > 1)
    .Tools("ID_TileH").Enabled = (iFormCount > 1)
    .Tools("ID_ArrangeIcons").Enabled = (iFormCount > 1)
    .Tools("ID_Maximise").Enabled = (iFormCount > 1)
    If .Tools("ID_Maximise").Enabled Then
      If Not ActiveForm Is Nothing Then .Tools("ID_Maximise").Enabled = (ActiveForm.WindowState <> vbMaximized)
    End If
    .Tools("ID_Minimise").Enabled = (iFormCount > 1)
    If .Tools("ID_Minimise").Enabled Then
      If Not ActiveForm Is Nothing Then .Tools("ID_Minimise").Enabled = (ActiveForm.WindowState <> vbMinimized)
    End If
    .Tools("ID_Restore").Enabled = (iFormCount > 1)
    If .Tools("ID_Restore").Enabled Then
      If Not ActiveForm Is Nothing Then .Tools("ID_Restore").Enabled = (ActiveForm.WindowState <> vbNormal)
    End If
    .Tools("ID_Close").Enabled = (iFormCount > 1)
  End With

End Sub


Private Sub SetCaption()

  '09/08/2001 MH Fault 2667
  'Me.Caption = "OpenHR - Security Manager"
  Me.Caption = Application.Name & " - " & gsDatabaseName

  Select Case Application.AccessMode
  Case accFull
    If gbShiftSave Then
      Me.Caption = Me.Caption & " [Shift Save]"
    End If
  Case accSupportMode
    Me.Caption = Me.Caption & "  [Support Mode]"
  Case accLimited
    Me.Caption = Me.Caption & "  [Limited Access]"
  Case accSystemReadOnly
    Me.Caption = Me.Caption & "  [Read Only]"
  'case accNone
  End Select

End Sub

Private Sub MDIForm_Resize()
  'JPD 20030908 Fault 5756
  If Me.WindowState <> vbMinimized Then
    giWindowState = Me.WindowState
    
'    If Me.Height < 2000 Then Me.Height = 2000
    
    If Me.WindowState = vbNormal Then
      glngWindowLeft = IIf(Me.Left < (0 - Me.Width), glngWindowLeft, Me.Left)
      glngWindowTop = IIf(Me.Top < (0 - Me.Height), glngWindowTop, Me.Top)
      glngWindowHeight = IIf((Me.Top < (0 - Me.Height)) Or (Me.ScaleHeight <= 0), glngWindowHeight, Me.Height)
      glngWindowWidth = IIf((Me.Left < (0 - Me.Width)) Or (Me.ScaleWidth <= 0), glngWindowWidth, Me.Width)
    End If
  End If

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
    
    sMessage = sMessage & rsMessages.Fields(0).Value
    
    rsMessages.MoveNext
  Loop
    
  rsMessages.Close
  Set rsMessages = Nothing

  If Len(sMessage) > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  
End Sub


