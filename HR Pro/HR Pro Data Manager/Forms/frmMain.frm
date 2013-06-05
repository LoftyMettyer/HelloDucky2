VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "Codejock.SkinFramework.v13.1.0.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00F7EEE9&
   Caption         =   "HR Pro - Data Manager"
   ClientHeight    =   10080
   ClientLeft      =   2550
   ClientTop       =   2820
   ClientWidth     =   12360
   HelpContextID   =   1001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   12330
      TabIndex        =   1
      Top             =   0
      Width           =   12360
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   2445
         ScaleHeight     =   1170
         ScaleWidth      =   1170
         TabIndex        =   2
         Top             =   195
         Width           =   1200
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3570
      Top             =   2460
   End
   Begin MSComDlg.CommonDialog CommonDialogOLD 
      Left            =   2160
      Top             =   1515
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpFile        =   "HRProHelp.chm"
   End
   Begin VB.Timer tmrDiary 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1275
      Top             =   1485
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9825
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16139
            MinWidth        =   1058
            Key             =   "pnlMAIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "FILTERED"
            TextSave        =   "FILTERED"
            Key             =   "pnlFILTER"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Key             =   "pnlCAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Key             =   "pnlNUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "14:58"
            Key             =   "pnlTIME"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3075
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5400
      Top             =   4680
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin ActiveBarLibraryCtl.ActiveBar abMain 
      Left            =   600
      Top             =   1320
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
      Bands           =   "frmMain.frx":038A
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
'
' Edit Control Messages
'
Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_UNDO = &H304
Const EM_CANUNDO = &HC6
Const EM_GETMODIFY = &HB8

Private mbLoading As Boolean
Private mbChanging As Boolean

' Functions to tile the background image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal lDC As Long, ByVal hObject As Long) As Long
Dim pic As StdPicture, hMemDC As Long, pHeight As Long, pWidth As Long

Dim mfMenuDisabled As Boolean

Private mstrLastAlarmCheck As String

'Private mblnLoggingOff As Boolean

Public Sub DisableMenu()
  'JPD 20030905 Fault 5184
  Dim iLoop As Integer
  
  mfMenuDisabled = True

  For iLoop = 0 To abMain.Bands.Item("MainMenu").Tools.Count - 1
    abMain.Bands.Item("MainMenu").Tools.Item(iLoop).Enabled = False
  Next iLoop

  abMain.RecalcLayout
  abMain.ResetHooks
  abMain.Refresh

End Sub

Public Sub EnableMenu(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)
  'JPD 20030905 Fault 5184
  Dim iLoop As Integer
  
  
  'MH20040218 Fault 8080
  'Prevent MDI being inadvertidly loaded again during unload.
  If gcoTablePrivileges Is Nothing Then
    Exit Sub
  End If


  mfMenuDisabled = False

  For iLoop = 0 To abMain.Bands.Item("MainMenu").Tools.Count - 1
    abMain.Bands.Item("MainMenu").Tools.Item(iLoop).Enabled = True
  Next iLoop

  RefreshMainForm pfrmCallingForm, pfUnLoad

End Sub

Private Sub EventLogClick()

  Dim fExit As Boolean
  Dim frmLog As frmEventLog
  
  Screen.MousePointer = vbHourglass
  
  fExit = False
  Set frmLog = New frmEventLog
    
  With frmLog
    .Show vbModal
  End With
  
  Unload frmLog
  Set frmLog = Nothing

End Sub

Private Sub WorkflowLogClick()

  Dim fExit As Boolean
  Dim frmLog As frmWorkflowLog
  
  If GetSystemSetting("workflow", "suspended", "0") = "1" Then
    COAMsgBox "The Workflow Service is currently suspended." & vbCrLf & vbCrLf & "Please contact your system administrator.", vbOKOnly & vbInformation, "Workflow"
  End If
  
  Screen.MousePointer = vbHourglass
  
  fExit = False
  Set frmLog = New frmWorkflowLog
    
  With frmLog
    .Show vbModal
  End With
  
  Unload frmLog
  Set frmLog = Nothing

End Sub


Public Sub Reload()
  'JPD 20040625 Fault 8714
  MDIForm_Load
  
End Sub

Public Sub SetBackground(ByRef mbIsLoading As Boolean)

  Dim x, y, hMemDC, pHeight, pWidth As Long
  Dim pic As StdPicture
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
    
    If Len(sFileName) > 0 Then
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
        For x = 0 To Me.ScaleWidth Step pWidth
          For y = 0 To Me.ScaleHeight Step pHeight
            BitBlt picWork.hDC, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
          Next
        Next
      End If

      ' Tiled down the lefthand side
      If glngDesktopBitmapLocation = giLOCATION_LEFTTILE Then
        For y = 0 To Me.ScaleHeight Step pHeight
          BitBlt picWork.hDC, 0, y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      ' Tiled down the righthand side
      If glngDesktopBitmapLocation = giLOCATION_RIGHTTILE Then
        For y = 0 To Me.ScaleHeight Step pHeight
          BitBlt picWork.hDC, (Me.ScaleWidth - pWidth) \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
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
        x = (ScaleWidth - pWidth) \ 2: x = x \ Screen.TwipsPerPixelX
        y = (ScaleHeight - pHeight) \ 2: y = y \ Screen.TwipsPerPixelY
        BitBlt picWork.hDC, x, y, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
      End If

      ' Tiled across the top
      If glngDesktopBitmapLocation = giLOCATION_TOPTILE Then
        For x = 0 To Me.ScaleWidth Step pWidth
          BitBlt picWork.hDC, x \ Screen.TwipsPerPixelX, 0, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
        Next
      End If

      'Tiled across the bottom
      If glngDesktopBitmapLocation = giLOCATION_BOTTOMTILE Then
        For x = 0 To Me.ScaleWidth Step pWidth
          BitBlt picWork.hDC, x \ Screen.TwipsPerPixelX, (Me.ScaleHeight - pHeight) \ Screen.TwipsPerPixelX, pWidth \ Screen.TwipsPerPixelX, pHeight \ Screen.TwipsPerPixelY, hMemDC, 0, 0, vbSrcCopy
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


Private Sub abMain_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
  ' Populate the Window drop-down menu band with the appropriate options
  ' depending on the current form state, and the forms that exist in the MDI form.
  Dim bNoSeparator As Boolean
  Dim lCount As Long
  Dim frmForm As Form
  Dim objTool As ActiveBarLibraryCtl.Tool
  Dim lWindowCount As Long
  
  ' JPD20020926 Fault 4431 - I apologise for the following lines of code.
  ' SendKeys is a horrible way to manage controls, but there was no other
  ' way I could find to stop the menu band dropping down if someone clicked
  ' on the menu when the find window was still loading. Sorry.
  If Screen.MousePointer = vbHourglass Then
    SendKeys "{ESC}"
    SendKeys "{ESC}"
    Exit Sub
  End If
  
  Select Case Band.Name
    Case "mnuWindow"
      bNoSeparator = False
      With Band.Tools
        ' Clear all options.
        .RemoveAll
        
        ' Add the standard window display options.
        .Insert -1, abMain.Tools("Cascade")
        .Insert -1, abMain.Tools("Arrange")
        .Insert -1, abMain.Tools("Minimise")
        .Insert -1, abMain.Tools("Restore")
        .Insert -1, abMain.Tools("CloseWindow")
        
        ' Add the options for each MDI child form.
        lCount = abMain.Tools.Count
        For Each frmForm In Forms
          If frmForm.Name <> "frmMain" Then
            If frmForm.Visible Then
              lCount = lCount + 1
              lWindowCount = lWindowCount + 1
              Set objTool = Band.Tools.Add(frmForm.hWnd, "WList" & lWindowCount)
              
              ' RH 09/10/00 - BUG, limit the chars in the Window Menu to 100
              'objTool.Caption = "&" & lWindowCount & "  " & frmForm.Caption
              objTool.Caption = "&" & lWindowCount & "  " & Left(frmForm.Caption, 100)
                        
              If Not bNoSeparator Then
                objTool.BeginGroup = True
                bNoSeparator = True
              End If
                        
              If Me.ActiveForm.hWnd = frmForm.hWnd Then
                objTool.Checked = True
              End If
            End If
          End If
        Next
      End With
        
      abMain.RecalcLayout
      
    Case "mnuHistory"
      PopulateHistoryMenu
      
  End Select
          
End Sub


Private Sub abMain_MenuItemEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  DoEvents
End Sub

Private Sub abMain_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub MDIForm_Activate()

  ' Reset the mouse pointer.
  Screen.MousePointer = vbDefault

  ' NPG20091007 Fault HR Pro-416
  ' set the new multi-size icons for taskbar, application, and alt-tab
  SetIcon Me.hWnd, "TASKBAR", True

End Sub

Private Sub MDIForm_Load()
  
  DebugOutput "MDIForm_Load", "LoadSkin"
  
  ' Load the CodeJock Styles
  Call LoadSkin(Me, Me.SkinFramework1)
  
  Dim objDefPrinter As cSetDfltPrinter

  'MH20040218 Fault 8080
  'Prevent MDI being inadvertidly loaded again during unload.
  If gcoTablePrivileges Is Nothing Then
    Exit Sub
  End If

  DebugOutput "MDIForm_Load", "Set abMain"
  
  With abMain
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
  
  abMain.Attach

  Me.Caption = "HR Pro Data Manager - " & gsDatabaseName

  DebugOutput "MDIForm_Load", "EnableTools"
  
  EnableTools
    
  gblnDiaryConstCheck = CBool(GetUserSetting("Diary", "ConstantCheck", True))
  Database.Validation = True

  '******************************************************************************
  ' Set Default printer settings
  
  'TM20020828 Fault 1432
  'TM20020911 Fault 4401
  
  
  'gblnStartupPrinter = (InStr(LCase(Command$), "/printer=false") > 0)
  'If Not gblnStartupPrinter Then
  gblnStartupPrinter = (InStr(LCase(Command$), "/printer=true") > 0)
  If gblnStartupPrinter Then
    'JPD 20081205 - You can have Printers.Count > 0 but still no valid printers (honestly!)
    ' So need to have proper error trapping, on top of the Printers.Count check.
    On Error GoTo PrinterErrorTrap
    
    DebugOutput "MDIForm_Load", "SetPrinterAsDefault"
    
    If Printers.Count > 0 Then
      gstrDefaultPrinterName = Printer.DeviceName
      SavePCSetting "Printer", "DeviceName", gstrDefaultPrinterName
  
      Set objDefPrinter = New cSetDfltPrinter
      objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
      Set objDefPrinter = Nothing
  
    End If
  
  
  
  
PrinterErrorTrap:
  End If
  '******************************************************************************
  
  'Printing options
  gbPrinterPrompt = GetPCSetting("Printer", "Prompt", True)
  gbPrinterConfirm = GetPCSetting("Printer", "Confirm", False)
  
  DebugOutput "MDIForm_Load", "GetScreens"
  
  ' Get the list of screens with which to populate the menu.
  GetScreens

 End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc: This a wrapper function for SendMessage to request the function
'       passed into it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditPerform(EditFunction As Integer)

   If TypeOf Screen.ActiveControl Is TextBox Then
      Call SendMessage(Screen.ActiveControl.hWnd, EditFunction, 0, 0&)
   End If
   
End Sub

Public Sub RefreshEditMenu()
  ' This procedure determines which edit menu options should be enabled.
  With abMain
    ' Check that we do have an active control
    If Not Screen.ActiveControl Is Nothing Then
      If TypeOf Screen.ActiveControl Is TextBox Then
        ' Determine if last edit can be undone
        .Tools("Undo").Enabled = SendMessage(Screen.ActiveControl.hWnd, EM_CANUNDO, 0, 0&)
        ' See if there's anything to cut, copy,
        ' or delete
        .Tools("Cut").Enabled = Screen.ActiveControl.SelLength
        .Tools("Copy").Enabled = Screen.ActiveControl.SelLength
        .Tools("Clear").Enabled = Screen.ActiveControl.SelLength
        ' See if there's anything to paste
        .Tools("Paste").Enabled = Clipboard.GetFormat(vbCFText)
      Else
        ' If active control is not a textbox
        ' then disable all
        .Tools("Undo").Enabled = False
        .Tools("Cut").Enabled = False
        .Tools("Copy").Enabled = False
        .Tools("Clear").Enabled = False
        .Tools("Paste").Enabled = False
      End If
    Else
      .Tools("Undo").Enabled = False
      .Tools("Cut").Enabled = False
      .Tools("Copy").Enabled = False
      .Tools("Clear").Enabled = False
      .Tools("Paste").Enabled = False
    End If
  End With
  
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Unload any remaining forms.
  Dim iFormCount As Integer
  Dim lngFormID As Long

  Do While Forms.Count > 1
    ' Remember how many forms are loaded.
    
    iFormCount = Forms.Count

    ' Try to unload the child form.
    lngFormID = 0
    If Forms(1).Name = "frmRecEdit4" Then
      lngFormID = Forms(1).FormID
    End If
    
    Unload Forms(1)
    DoEvents

    ' If the number of loaded forms has not changed then the
    ' form was not unloaded. So cancel the MDI unload.
    If (iFormCount = Forms.Count) Then
      'MH20040218 Fault 8080
      'Cancel = True
      Cancel = Me.Visible
      Exit Do
    End If
    
    'JPD 20030820 Fault 3047
    If (Forms.Count > 1) And (lngFormID > 0) Then
      If Forms(1).Name = "frmRecEdit4" Then
        If lngFormID = Forms(1).FormID Then
          'MH20040218 Fault 8080
          'Cancel = True
          Cancel = Me.Visible
          Exit Do
        End If
      End If
    End If
  Loop
  
  

  If Forms.Count > 1 Then
  ' If there are child forms still loaded then cancel the MDI unload.
    frmMain.RefreshMainForm frmMain.ActiveForm
    'MH20040218 Fault 8080
    'Cancel = True
    Cancel = Me.Visible
  End If

  If Cancel = False Then
    'Phils PC was not refreshing the screen at logoff
    'and the menu bar was still visible even though
    'the form wasn't !  So had to put in all this stuff
    'and it seemed to correct the problem.  MH20000629
    abMain.Tools.RemoveAll
    abMain.Bands.RemoveAll
    abMain.ReleaseFocus

    Set gcoTablePrivileges = Nothing
    Set gcolColumnPrivilegesCollection = Nothing
    Set gcoLookupValues = Nothing
    Set gcolHistoryScreensCollection = Nothing
    Set gcolSummaryFieldsCollection = Nothing
    Set gcolScreens = Nothing
    Set gcolScreenControls = Nothing
  End If

  'Remove Progress Bar class from memory
  Set gobjProgress = Nothing

  'MH20020410 Fault 3757
  Set gobjDiary = Nothing

  If Cancel = False Then Call AuditAccess(iLOGOFF, "Data")

End Sub

Public Sub abMain_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  ' Perform the selected menu option action.
  Dim lPos As Long
  Dim lForms As Long
  Dim iUsers As Integer
  Dim strVersionFilename As String
  Dim plngHelp As Long
  
  ' JPD20020926 Fault 4431
  If Screen.MousePointer = vbHourglass And Not gbJustRunIt Then Exit Sub
  
  ' Get rid of any screen display residue.
  'DoEvents
  If Tool.Name <> "LogOff" Then
    DoEvents
  End If

  ' Check if the clicked tool is in the window list.
  If Left(Tool.Name, 5) = "WList" Then
    For lForms = 0 To Forms.Count - 1
      If Forms(lForms).hWnd = Tool.ToolID Then
        If Forms(lForms).Visible Then
          If Forms(lForms).Enabled Then
            ' Set focus onto the selected form.
            Forms(lForms).SetFocus
          End If
        End If
            
        Exit For
      End If
    Next
    
    Exit Sub
  End If
  
  Select Case Tool.Name
    ' <DATABASE> menu.
    
    ' <Record Edit Screens> - handled below
    
    ' <Table Screens>       - handled below
    
    ' <Logoff>
    Case "LogOff"
      'If mblnLoggingOff = False Then
      '  mblnLoggingOff = True
        Unload frmSplash
          
        'NHRD 16042002 Fault 3381 Log Off clarification notice.
        'The F3 key was added to the log off option in the active bar designer.
        If COAMsgBox("Are you sure you wish to log off?", vbQuestion + vbYesNo, "Logging Off") = vbYes Then
          'Looks like we want to log off so do the necessary.
          Database.Validation = True
          'RH 12/10/00 - Now call from frmMain_QueryUnload
          'Call AuditAccess(False, "Data")
          Unload Me
          
          If Database.Validation Then
            gbForceLogonScreen = True
            datGeneral.ClearConnection
            Main
          End If
        End If
      '  mblnLoggingOff = False
      'End If

    ' <Exit>
    Case "Exit"
      'RH 12/10/00 - Now call from frmMain_QueryUnload
      'Call AuditAccess(False, "Data")
      Unload Me
      
    ' <RECORD> menu.
    
    ' <New>
    Case "NewRecord"
      ActiveForm.AddNew
    Case "CopyRecord"
      ActiveForm.AddNewCopyOf
    Case "EditRecord"
      ActiveForm.EditRecord
    ' <Save>
    Case "SaveRecord"
      ActiveForm.UpdateWithAVI
    ' <Delete>
    Case "DeleteRecord"
      ActiveForm.DeleteRecord
    ' <First Record>
    Case "FirstRecord"
      ActiveForm.MoveFirst
    ' <Previous Record>
    Case "PreviousRecord"
      ActiveForm.MovePrevious
    ' <Next Record>
    Case "NextRecord"
      ActiveForm.MoveNext
    ' <Last Record>
    Case "LastRecord"
      ActiveForm.MoveLast
    ' <Find>
    Case "FindRecord"
      ActiveForm.Find
    ' <QuickFind>
    Case "QuickFind"
      frmQuickFind.Initialise ActiveForm
    ' <Refresh>
    Case "Refresh"
      ActiveForm.Requery False
    ' <Order>
    Case "Order"
      ActiveForm.SelectOrder
    ' <Filter>
    Case "Filter"
      ActiveForm.SelectFilter
    ' <Clear Filter>
    Case "FilterClear"
      ActiveForm.ClearFilter
    ' <Filter>
    Case "CancelCourse"
      ActiveForm.CancelCourse
    
    ' Payroll resend current record
    Case "ID_Accord_SendAs_Update"
      ActiveForm.SendToAccord False
    
    ' Payroll resend current record
    Case "ID_Accord_SendAs_New"
      ActiveForm.SendToAccord True
    
    ' <Mail Merge>
    Case "MailMergeRec"
      ActiveForm.MailMergeClick

    ' <Envelopes & Labels>
    Case "LabelsRec"
      ActiveForm.LabelsClick
    
    Case "DataTransferRec"
      ActiveForm.DataTransferClick
    Case "Email"
      ActiveForm.EmailClick

    Case "BookCourse"
      ActiveForm.BookCourse
    Case "BulkBooking"
      ActiveForm.BulkBooking
    Case "AddFromWaitingList"
      ActiveForm.AddFromWaitingList
    Case "TransferBooking"
      ActiveForm.TransferBooking
    Case "CancelBooking"
      ActiveForm.CancelBooking

    Case "AbsenceBreakdownRec"
      ActiveForm.AbsenceBreakdownClick
      
    Case "BradfordIndexRec"
      ActiveForm.BradfordFactorClick
      
    Case "AbsenceCalendar"
      ActiveForm.AbsenceCalendarClick

    Case "MatchReportRec"
      ActiveForm.MatchReportClick mrtNormal

    Case "SuccessionRec"
      ActiveForm.MatchReportClick mrtSucession

    Case "CareerRec"
      ActiveForm.MatchReportClick mrtCareer

    ' <Record Profile>
    Case "RecordProfileRec"
      ActiveForm.RecordProfileClick
    
    ' <Calendar Report>
    Case "CalendarReportRec"
      ActiveForm.CalendarReportClick
    
    ' <REPORTS> menu.
    
    ' <Cross Tabulations>
    Case "CrossTab"
      CrossTabClick
    
    ' <Crystal Reports"
    Case "CrystalReports"
      CrystalReportsClick
    
    ' <Custom Reports>
    Case "CustomReports"
      CustomReportsClick
    
    ' <Calendar Report>
    Case "CalendarReport"
      CalendarReportsClick
    
    ' <Match Report>
    Case "MatchReport"
      MatchReportClick mrtNormal
    
    ' <Match Report>
    Case "Career"
      MatchReportClick mrtCareer
    
    ' <Match Report>
    Case "Succession"
      MatchReportClick mrtSucession
    
    ' <Record Profile>
    Case "RecordProfile"
      RecordProfileClick
    
    ' <Mail Merge>
    Case "MailMerge"
      MailMergeClick
      
    ' <Envelopes & Labels>
    Case "mnuLabels"
      LabelsAndEnvelopesClick
      
'    ' <Standard Reports - Absence Calendar>
'    Case "AbsenceCalendar"
'      If Forms.Count > 1 Then
'        If Not ActiveForm.SaveChanges Then
'          Exit Sub
'        End If
'        If ValidateAbsenceParameters Then AbsenceCalendarClick
'      End If
     
    Case "AbsenceBreakdown", _
         "AbsenceBreakdownCfg", _
         "BradfordIndex", _
         "BradfordIndexCfg", _
         "StabilityIndex", _
         "StabilityIndexCfg", _
         "Turnover", _
         "TurnoverCfg"
      If SaveCurrentRecordEditScreen Then
        StandardReportClick Tool.Name
      End If

    'Case "Career"
    '  CareerSuccessionClick mrtCareer

    'Case "Succession"
    '  CareerSuccessionClick mrtSucession


    ' <UTILITIES> menu.
    
    '<Diary>
    Case "Diary"
      gobjDiary.DateSelected = Now
      gobjDiary.ViewingAlarms = False
      frmDiary.Initialise
      frmDiary.Show vbModal
      Set gobjDiary = Nothing

    ' <Data Transfer>
    Case "DataTransfer"
      If SaveCurrentRecordEditScreen Then
        DataTransferClick
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If

    ' <Global Functions - Add>
    Case "GlobalAdd"
      If SaveCurrentRecordEditScreen Then
        GlobalClick glAdd
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If

    ' <Global Functions - Update>
    Case "GlobalUpdate"
      If SaveCurrentRecordEditScreen Then
        GlobalClick glUpdate
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If

    ' <Global Functions - Delete>
    Case "GlobalDelete"
      If SaveCurrentRecordEditScreen Then
        GlobalClick glDelete
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If

    ' <Import>
    Case "Import"
      If SaveCurrentRecordEditScreen Then
        ImportClick
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If

    ' <Export>
    Case "Export"
      ExportClick
    
    ' <Workflow>
    Case "Workflow"
      WorkflowClick
    
    ' Pending Payroll stuff
    Case "ID_Accord_Live"
  
      frmAccordViewTransfers.ConnectionType = ACCORD_LOCAL
      frmAccordViewTransfers.ViewMode = iLIVE_ALL
      frmAccordViewTransfers.Initialise
      frmAccordViewTransfers.Show vbModal
      Set frmAccordViewTransfers = Nothing
    
    ' Archive Payroll stuff
    Case "ID_Accord_Archive"
  
      frmAccordViewTransfers.ConnectionType = ACCORD_LOCAL
      frmAccordViewTransfers.ViewMode = iARCHIVE_ALL

      frmAccordViewTransfers.Initialise
      frmAccordViewTransfers.Show vbModal
      Set frmAccordViewTransfers = Nothing
    
    ' View current transaction in Payroll
    Case "ID_Accord_Current"
      ActiveForm.AccordClick
  
    ' Data dump to Payroll
    Case "ID_Accord_Create"
      With frmAccordExportRecords
        If .Initialise = True Then
          .Show vbModal
        End If
      End With
      Set frmAccordExportRecords = Nothing
  
    Case "BatchJobs"
      If SaveCurrentRecordEditScreen Then
        BatchJobsClick
        ' RH - Note : this sub causes the enabling history menu bug
        RefreshRecordEditScreens
      End If


    ' <Tools> menu.
    Case "Calculations"
      CalculationsClick
      
    Case "PickLists"
      PickListClick
    
    Case "Filters"
      FilterClick
        
    Case "EmailGroups"
      EmailGroupClick

    Case "ID_LabelTemplates"
      LabelTemplatesClick

    Case "ID_DocumentTypes"
      DocumentTypesClick



    ' <WINDOW> menu.
    
    '<Arrange The Minimised Icons>
    Case "Arrange"
      frmMain.Arrange vbArrangeIcons
    
    '<Cascase Open Forms>
    Case "Cascade"
      frmMain.Arrange vbCascade
    
    '<Minimise The Active Form>
    Case "Minimise"
      frmMain.ActiveForm.WindowState = vbMinimized
      frmMain.RefreshMainForm frmMain.ActiveForm

    '<Restore The Minimised window>
    Case "Restore"
      frmMain.ActiveForm.WindowState = vbNormal
      frmMain.RefreshMainForm frmMain.ActiveForm
    
    '<Close The Active Form>
    Case "CloseWindow"
      ' JPD20020926 Fault 4440
      If (TypeOf frmMain.ActiveForm Is frmFind2) Then
        frmMain.ActiveForm.Cancelling = True
      End If
      
      Unload frmMain.ActiveForm
      frmMain.RefreshMainForm frmMain
      
    ' <ADMINISTRATION> menu.
    
    ' <Configuration>
    Case "Configuration"
      
      If frmConfiguration.Initialise(True) Then
        frmConfiguration.Show vbModal
      End If
      Set frmConfiguration = Nothing

      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
    
    
    Case "PC Configuration"
      If frmConfiguration.Initialise(False) Then
        frmConfiguration.Show vbModal
      End If
      Set frmConfiguration = Nothing

      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
    
    
    '<Change Password>
    Case "ChangePassword"
      
      If Database.Validation Then
        'JPD 20020218 Fault 3343 - do this check again when the user clicks okay to apply
        ' the new password, as there may not have been another user with the same login name
        ' at this point, but one might log in after this point but before the new password is committed.
        'TM20020118 Fault 3343 - Check how many sessions the current account has in
        'the system before allowing a password change.
        
        'MH20061017 Fault 11376
        'iUsers = UserSessions(gsUserName)
        iUsers = GetCurrentUsersCountOnServer(gsUserName)
        
        If iUsers < 2 Then
          'If frmChangePassword.Initialise(0, GetMinimumPasswordLength) Then
          If frmChangePassword.Initialise(0, GetSystemSetting("Password", "Minimum Length", 0)) Then
            frmChangePassword.Show vbModal
          End If
          Unload frmChangePassword
          Set frmChangePassword = Nothing
        Else
          COAMsgBox "Cannot change password. This account is currently being used " & _
                  "by " & IIf(iUsers > 2, iUsers & " users", "another user") & " in the system.", vbExclamation + vbOKOnly, App.Title
        End If
      End If

    '<Diary Rebuild>
    Case "DiaryRebuild"
      DiaryRebuild

    '<Diary Delete>
    Case "DiaryDelete"
      frmDiaryDelete.Show vbModal

    '<Diary Alarm Toggle>
    'Case "DiaryToggle"
    '  gblnDiaryConstCheck = (Not gblnDiaryConstCheck)
    '  tmrDiary.Enabled = gblnDiaryConstCheck
    '  Me.abMain.Tools("DiaryToggle").Checked = gblnDiaryConstCheck
    '  'Me.abMain.Bands("bndAdmin_Diary").Tools("DiaryToggle").Caption = IIf(gblnDiaryConstCheck = True, "Disable &Alarms", "Enable &Alarms")
    '  Me.abMain.Bands("mnuAdministration").Tools("DiaryToggle").Caption = IIf(gblnDiaryConstCheck = True, "Disable &Alarms", "Enable &Alarms")
    '  Me.abMain.Refresh

    '<Email Queue>
    Case "EmailQueue"
      'NHRD27102004 Fault 8538
      Screen.MousePointer = vbHourglass
      
      frmEmailQueue.Show vbModal
      Set frmEmailQueue = Nothing

    '<Email Queue>
    Case "OutlookQueue"
      frmOutlookQueue.Show vbModal
      Set frmOutlookQueue = Nothing
    
    '<Event Log>
    Case "EventLog"
      EventLogClick
'      frmEventLog.Show vbModal
'      Set frmEventLog = Nothing
      
    '<Workflow Log>
    Case "WorkflowLog"
      WorkflowLogClick
      
    '<Workflow Pending Steps>
    Case "PendingSteps"
      CheckPendingWorkflowSteps True
      
    '<Workflow OutOfOffice>
    Case "WorkflowOutOfOffice"
      WorkflowOutOfOffice
      
    '<Set Default Printer>
'    Case "SetDefaultPrinter"
'      frmSetDefaultPrinter.Show vbModal
'      Set frmSetDefaultPrinter = Nothing
      
    '<View Current Users>
    Case "ViewCurrentUsers"
      frmViewCurrentUsers.Show vbModal
      Set frmViewCurrentUsers = Nothing
      
    ' <HELP> menu.
    
    ' <Contents>
    
    ' <Search>
    
    ' <About>
    Case "HelpContentsAndIndex"
  
      If Not ShowAirHelp(0) Then
        plngHelp = ShellExecute(0&, vbNullString, App.Path & "\" & App.HelpFile, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          COAMsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to find and view the file " & App.HelpFile & ".", vbExclamation + vbOKOnly, App.EXEName
        End If
      End If
  
    ' <About>
    Case "HelpAbout"
      Screen.MousePointer = vbHourglass
      Load frmAbout
      DoEvents     ' needed to prevent grey sqr appearing when about displays
      frmAbout.Show vbModal
      Set frmAbout = Nothing
      If Not ActiveForm Is Nothing Then
        ActiveForm.SetFocus
      End If
    
    Case "ID_VersionInfo"
      Screen.MousePointer = vbHourglass
      
      strVersionFilename = App.Path & "\HR Pro Data Manager Version Information.htm"
      
      If Len(strVersionFilename) > 0 Then
        plngHelp = ShellExecute(0&, vbNullString, strVersionFilename, vbNullString, vbNullString, vbNormalNoFocus)
        If plngHelp = 0 Then
          COAMsgBox "Error whilst attempting to display version information file.", vbExclamation + vbOKOnly, "HR Pro Data Manager"
        End If
      Else
        COAMsgBox "No version information found.", vbExclamation + vbOKOnly, "HR Pro Data Manager"
      End If
      
      Screen.MousePointer = vbNormal
    
    ' <CMG Recovery>
    Case "ID_CMGRecovery"
      frmCMGRecovery.Show vbModal
      Set frmCMGRecovery = Nothing

    Case "ID_CMGCommit"
      If datGeneral.CMGCommit = True Then
        COAMsgBox "CMG Commit successful", vbOKOnly & vbInformation, "CMG"
      End If
    
    
    ' Version 1 stuff
    Case "ID_PollMode"
      RunPollJob
    
    
    ' Message boxes
    Case "MessageBox"
      COAMsgBox Tool.Text, vbOKOnly
    
    
    ' JPD20021126 Fault 4805
    Case "ID_Print"
      ActiveForm.PrintGrid
       
       
    Case Else
      ' It must be a screen of some kind so decide what type it is.
      Select Case Left(Tool.Name, 2)
        
        Case "QE" ' Quick Entry Screen.
          EditForm_Load Val(Right(Tool.Name, Len(Tool.Name) - 2)), screenQuickEntry
          
        Case "PT" ' Parent screen based on a TABLE.
          EditForm_Load Val(Right(Tool.Name, Len(Tool.Name) - 2)), screenParentTable
        
        Case "PV" ' Parent screen based on a VIEW.
          lPos = InStr(1, Tool.Name, ":")
          EditForm_Load Val(Mid$(Tool.Name, 3, lPos - 1)), screenParentView, Val(Mid$(Tool.Name, lPos + 1, Len(Tool.Name)))
        
        Case "HT" ' History screen based on a TABLE.
          EditForm_Load Val(Right(Tool.Name, Len(Tool.Name) - 2)), screenHistoryTable
        
        Case "HV" ' History screen based on a VIEW.
          lPos = InStr(1, Tool.Name, ":")
          EditForm_Load Val(Mid$(Tool.Name, 3, lPos - 1)), screenHistoryView, Val(Mid$(Tool.Name, lPos + 1, Len(Tool.Name)))
      
        Case "TS" ' Table Screen
          Dim rsTemp As Recordset
          Set rsTemp = datGeneral.GetScreenScreens(Mid$(Tool.Name, 3, Len(Tool.Name)))
          If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
              If rsTemp.RecordCount = 1 Then
                                
                '#rh 27/10  If lookup table is accessed from the main menu, it shouldnt
                '#          really default to a new record.
                'If Not EditForm_Load(rsTemp!ScreenID, screenLookup, , True) Then
                If Not EditForm_Load(rsTemp!ScreenID, screenLookup, , False) Then
                
                End If
              End If
          End If
      End Select
  End Select

  ' RH 07/06/00 - To prevent toolbar locking after anything (event log mainly)
  If (Not Tool.Name = "Exit") And (Not Tool.Name = "LogOff") Then

    ' RH 17/07/00 - Required to prevent history menu becoming active after accessing a
    '               utility that updates data.
    If Not Me.ActiveForm Is Nothing Then
      frmMain.RefreshMainForm Me.ActiveForm
    End If

    With abMain
      .ResetHooks
      .Refresh
    End With

    'MH20060623 Check the diary a.s.a.p...
    If tmrDiary.Enabled Then
      tmrDiary.Interval = 1
      tmrDiary.Enabled = False
      tmrDiary.Enabled = True
    End If

  End If

End Sub



Public Sub RefreshRecordMenu(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmMain.RefreshRecordMenu(pfrmCallingForm,pfUnLoad)", Array(pfrmCallingForm, pfUnLoad)
  ' Refresh the Record menu options.
  
  Dim fNewRecordEnabled As Boolean
  Dim fCopyRecordEnabled As Boolean
  Dim fEditRecordEnabled As Boolean
  Dim fSaveRecordEnabled As Boolean
  Dim fDeleteRecordEnabled As Boolean
  Dim fFirstRecordEnabled As Boolean
  Dim fPreviousRecordEnabled As Boolean
  Dim fNextRecordEnabled As Boolean
  Dim fLastRecordEnabled As Boolean
  Dim fFindRecordEnabled As Boolean
  Dim fRefreshEnabled As Boolean
  Dim fOrderEnabled As Boolean
  Dim fFilterEnabled As Boolean
  Dim fFilterClearEnabled As Boolean
  Dim fCancelCourseEnabled As Boolean
  Dim fCancelCourseVisible As Boolean
  Dim fMailMergeExists As Boolean
  Dim fEnvelopeLabelsExists As Boolean
  Dim fDataTransferExists As Boolean
  Dim fEmailAddrExists As Boolean
  Dim fMatchReportExists As Boolean
  Dim fSuccessionPlanning As Boolean
  Dim fCareerProgression As Boolean
  Dim iScreenType As Integer
  Dim rsRecords As ADODB.Recordset
  Dim vBookMark As Variant
  Dim fQuickFindEnabled As Boolean '- RH 23/08/99
  Dim objTool As Tool
  Dim flag As Boolean
  Dim fAbsenceReportsEnabled As Boolean
  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String
  Dim objTableView As CTablePrivilege
  Dim colColumnPrivileges As CColumnPrivileges
  Dim fAddingNewRecord As Boolean
  Dim fRecordProfileExists As Boolean
  Dim fCalendarReportExists As Boolean
  Dim fCustomReportExists As Boolean
  Dim fGlobalUpdateExists As Boolean
  
  Dim fEnvelopesVisible As Boolean
  Dim fEnvelopesEnabled As Boolean
  
  Dim fBookCourseEnabled As Boolean
  Dim fBookCourseVisible As Boolean
  Dim fBulkBookingEnabled As Boolean
  Dim fBulkBookingVisible As Boolean
  Dim fAddFromWaitingListEnabled As Boolean
  Dim fAddFromWaitingListVisible As Boolean
  Dim fTransferBookingEnabled As Boolean
  Dim fTransferBookingVisible As Boolean
  Dim fCancelBookingEnabled As Boolean
  Dim fCancelBookingVisible As Boolean
  Dim fSelectionMade As Boolean
  Dim fOnlyOneSelectionMade As Boolean
  Dim bAccordResendVisible As Boolean
  Dim bAccordResendEnabled As Boolean
    
  Dim fSomeVisible As Boolean
  Dim fSomeEnabled As Boolean
  Dim fBeginGroup As Boolean
  Dim fBeginGroupDone As Boolean
  Dim strGroupType As String
  
  fNewRecordEnabled = False
  fCopyRecordEnabled = False
  fEditRecordEnabled = False
  fSaveRecordEnabled = False
  fDeleteRecordEnabled = False
  fFirstRecordEnabled = False
  fPreviousRecordEnabled = False
  fNextRecordEnabled = False
  fLastRecordEnabled = False
  fFindRecordEnabled = False
  fRefreshEnabled = False
  fOrderEnabled = False
  fFilterEnabled = False
  fFilterClearEnabled = False
  fCancelCourseEnabled = False
  fCancelCourseVisible = False
  fQuickFindEnabled = False '- RH 23/08/99
  fMailMergeExists = False
  fDataTransferExists = False
  fEmailAddrExists = False
  fMatchReportExists = False
  fSuccessionPlanning = False
  fCareerProgression = False
  fAddingNewRecord = False
  fRecordProfileExists = False
  fCalendarReportExists = False
  fCustomReportExists = False
  fGlobalUpdateExists = False
  
  fBookCourseEnabled = False
  fBookCourseVisible = False
  fBulkBookingEnabled = False
  fBulkBookingVisible = False
  fAddFromWaitingListEnabled = False
  fAddFromWaitingListVisible = False
  fTransferBookingEnabled = False
  fTransferBookingVisible = False
  fCancelBookingEnabled = False
  fCancelBookingVisible = False
  
'  ' Performance - No need to run if form isn't visible. Is there? I guess if you're reading this there probably is! :-(
'  ' Well that's just typical isn't it. It did...
'  If Not pfrmCallingForm.Visible Then
'    With pfrmCallingForm.ActiveBar1
'      .Refresh
'      .RecalcLayout
'    End With
'    Exit Sub
'  End If
  
  fEnvelopesVisible = False
  fEnvelopesEnabled = False
     
  ' Only configure the record menu for record edit and find windows.
  If (Not pfUnLoad) And _
   ((TypeOf pfrmCallingForm Is frmRecEdit4) Or _
    (TypeOf pfrmCallingForm Is frmFind2)) Then

    With pfrmCallingForm
      iScreenType = .ScreenType
      Set rsRecords = .Recordset
    End With
  
    If rsRecords.State <> adStateClosed Then
      ' Check that the current record still exists.
      If TypeOf pfrmCallingForm Is frmRecEdit4 Then
      
        With rsRecords
          ' JPD 20/02/2001 ADO2.6 error when trying to use the absolutePosition property
          ' when the recordset is in adEditAdd mode.
          If .EditMode <> adEditAdd Then
          
            If .AbsolutePosition = adPosUnknown Then
              ' The current record no longer exists. Try to move onto the next record.
              
              COAMsgBox "The current record has been deleted by another user, screen will be refreshed.", vbExclamation, App.ProductName
              
              If (Not .EOF) Then .MoveNext
              
              If .EOF Then
                If Not pfrmCallingForm.RefreshRecordset Then
                  GoTo TidyUpAndExit
                End If
            
                ' There are records in the refreshed recordset. Move to the last record.
                If .EditMode <> adEditAdd Then
                  .MoveLast
                End If
              End If
            
              pfrmCallingForm.UpdateControls
              pfrmCallingForm.UpdateChildren
            End If
          End If
        End With
      End If
    
      ' QuickFind menu option. RH 23/08/99
      fQuickFindEnabled = (iScreenType = screenParentTable) Or _
        (iScreenType = screenParentView) Or _
        (iScreenType = screenLookup)
  
      ' New Record menu option.
      fNewRecordEnabled = pfrmCallingForm.AllowInsert
      
      ' Save Record menu option.
      If TypeOf pfrmCallingForm Is frmRecEdit4 Then
        'TM20020528 Fault 2895 - also need to enable the save button if all the controls have
        'a default value.
        fSaveRecordEnabled = (pfrmCallingForm.AllowUpdate And pfrmCallingForm.Changed) _
          Or (pfrmCallingForm.AllowUpdate _
          And pfrmCallingForm.Recordset.EditMode = adEditAdd _
          And pfrmCallingForm.AllDefaults)
                  
        bAccordResendEnabled = pfrmCallingForm.AllowUpdate
                  
        fDeleteRecordEnabled = pfrmCallingForm.AllowDelete And _
          (rsRecords.EditMode <> adEditAdd)
        
        'MH20010516
        fCopyRecordEnabled = pfrmCallingForm.AllowInsert And _
          (rsRecords.EditMode <> adEditAdd)
          
                'MH20010516
        fCopyRecordEnabled = pfrmCallingForm.AllowInsert And _
          (rsRecords.EditMode <> adEditAdd)
          
          'NHRD16082004 Fault 8773
          Dim objColumn As CColumnPrivilege
          Dim fUniqueExists As Boolean
          fUniqueExists = False
          
          For Each objColumn In pfrmCallingForm.ColumnSelectPrivileges
            If objColumn.UniqueCheck And _
            objColumn.AllowSelect And _
            (objColumn.DataType = sqlNumeric Or _
            objColumn.DataType = sqlInteger Or _
            objColumn.DataType = sqlDate Or _
            objColumn.DataType = sqlVarChar) Then
          
              'fUniqueExists = True
              fUniqueExists = (iScreenType <> screenHistoryTable)
             
              Exit For
            End If
          Next
          
          fQuickFindEnabled = fUniqueExists
      Else
        fSelectionMade = (pfrmCallingForm.ssOleDBGridFindColumns.SelBookmarks.Count > 0) _
          And (pfrmCallingForm.ssOleDBGridFindColumns.Rows > 0)
        fOnlyOneSelectionMade = (pfrmCallingForm.ssOleDBGridFindColumns.SelBookmarks.Count = 1) And (pfrmCallingForm.ssOleDBGridFindColumns.Rows > 0)
        
        fSaveRecordEnabled = False
        fDeleteRecordEnabled = pfrmCallingForm.AllowDelete And _
          fSelectionMade
        fCopyRecordEnabled = pfrmCallingForm.AllowInsert And _
          fOnlyOneSelectionMade
        fEditRecordEnabled = fOnlyOneSelectionMade
      
        fBookCourseEnabled = pfrmCallingForm.CanBookCourse And fOnlyOneSelectionMade
        fBookCourseVisible = pfrmCallingForm.BookCourseVisible
        fBulkBookingEnabled = pfrmCallingForm.CanBulkBooking
        fBulkBookingVisible = pfrmCallingForm.BulkBookingVisible
        fAddFromWaitingListEnabled = pfrmCallingForm.CanAddFromWaitingList
        fAddFromWaitingListVisible = pfrmCallingForm.AddFromWaitingListVisible
        fTransferBookingEnabled = pfrmCallingForm.CanTransferBooking And fOnlyOneSelectionMade
        fTransferBookingVisible = pfrmCallingForm.TransferVisible
        fCancelBookingEnabled = pfrmCallingForm.CanCancelBooking And fOnlyOneSelectionMade
        fCancelBookingVisible = pfrmCallingForm.CancelBookingVisible
      
        fCustomReportExists = pfrmCallingForm.CustomReportExists
        fCalendarReportExists = pfrmCallingForm.CalendarReportExists
        fGlobalUpdateExists = pfrmCallingForm.GlobalUpdateExists
        fDataTransferExists = pfrmCallingForm.DataTransferExists
        fMailMergeExists = pfrmCallingForm.MailMergeExists
      
      End If
      
      If (iScreenType = screenParentTable) Or _
        (iScreenType = screenParentView) Or _
        (iScreenType = screenHistoryTable) Or _
        (iScreenType = screenHistoryView) Or _
        (iScreenType = screenQuickEntry) Or _
        (iScreenType = screenLookup) Then
        
        If Not (rsRecords.BOF And rsRecords.EOF) Then
          'Allow Mail Merge if there are records
  
          'Only check these if not adding new record
          strSQL = "SELECT COUNT(*) FROM ASRSysMailMergeName " & _
                   "WHERE TableID = " & CStr(pfrmCallingForm.TableID) & " AND IsLabel = 0"
          fMailMergeExists = (GetRecCount(strSQL) > 0)
    
          strSQL = "SELECT COUNT(*) FROM ASRSysMailMergeName " & _
                   "WHERE TableID = " & CStr(pfrmCallingForm.TableID) & " AND IsLabel = 1"
          fEnvelopeLabelsExists = (GetRecCount(strSQL) > 0)
    
          strSQL = "SELECT COUNT(*) FROM ASRSysDataTransferName " & _
                   "WHERE FromTableID = " & CStr(pfrmCallingForm.TableID)
          fDataTransferExists = (GetRecCount(strSQL) > 0)
          
          strSQL = "SELECT COUNT(*) FROM ASRSysEmailAddress " & _
                   "WHERE TableID = 0 OR TableID = " & CStr(pfrmCallingForm.TableID)
          fEmailAddrExists = (GetRecCount(strSQL) > 0)

          strSQL = "SELECT COUNT(*) FROM ASRSysMatchReportName " & _
                   "WHERE MatchReportType = 0 " & _
                   "AND " & CStr(pfrmCallingForm.TableID) & " IN (Table1ID, Table2ID)"
          fMatchReportExists = (GetRecCount(strSQL) > 0)

          If gfPersonnelEnabled Then
            strSQL = "SELECT COUNT(*) FROM ASRSysMatchReportName " & _
                     "WHERE MatchReportType = 1 " & _
                     "AND " & CStr(pfrmCallingForm.TableID) & " IN (Table1ID, Table2ID)"
            fSuccessionPlanning = (GetRecCount(strSQL) > 0 And pfrmCallingForm.TableID = glngPersonnelTableID)
  
            strSQL = "SELECT COUNT(*) FROM ASRSysMatchReportName " & _
                     "WHERE MatchReportType = 2 " & _
                     "AND " & CStr(pfrmCallingForm.TableID) & " IN (Table1ID, Table2ID)"
            fCareerProgression = (GetRecCount(strSQL) > 0 And pfrmCallingForm.TableID = glngPersonnelTableID)
          End If


          strSQL = "SELECT COUNT(*) FROM ASRSysRecordProfileName " & _
                   "WHERE baseTable = " & CStr(pfrmCallingForm.TableID)
          fRecordProfileExists = (GetRecCount(strSQL) > 0)

          strSQL = "SELECT COUNT(*) FROM ASRSysCalendarReports " & _
                   "WHERE baseTable = " & CStr(pfrmCallingForm.TableID)
          fCalendarReportExists = (GetRecCount(strSQL) > 0)
  
          With rsRecords
            fAddingNewRecord = (.EditMode = adEditAdd)
            If fAddingNewRecord Then
              ' Adding a new record. Enable the moveFirst, movePrevious options
              ' only if there are some real records.
              fNextRecordEnabled = False
              fLastRecordEnabled = False
              
              If pfrmCallingForm.RequiresLocalCursor Then
                fFirstRecordEnabled = ((rsRecords.RecordCount > 1) Or (rsRecords.EditMode <> adEditAdd))
                fPreviousRecordEnabled = fFirstRecordEnabled
              Else
                Set rsTemp = rsRecords.Clone(adLockReadOnly)
                fFirstRecordEnabled = Not (rsTemp.BOF And rsTemp.EOF)
                fPreviousRecordEnabled = Not (rsTemp.BOF And rsTemp.EOF)
                rsTemp.Close
                Set rsTemp = Nothing
              End If
            Else
              
              If .BOF Then .MoveFirst
              If .EOF Then .MoveLast
              
              vBookMark = .Bookmark
              .MovePrevious
              fFirstRecordEnabled = (Not .BOF)
              fPreviousRecordEnabled = fFirstRecordEnabled
              .Bookmark = vBookMark
              .MoveNext
              fNextRecordEnabled = (Not .EOF)
              fLastRecordEnabled = fNextRecordEnabled
              .Bookmark = vBookMark
            End If
          End With
        End If
        
        ' Find Record menu option.
        fFindRecordEnabled = True
        ' Refresh Record menu option.
        fRefreshEnabled = True
        ' Order menu option.
        fOrderEnabled = True
        ' Filter menu options.
        fFilterEnabled = True
        fFilterClearEnabled = pfrmCallingForm.Filtered
        
        ' Cancel Course menu option.
        'JPD 20041122 Fault 9538
        fCancelCourseVisible = gfTrainingBookingEnabled And _
          (pfrmCallingForm.TableID = glngCourseTableID)
        'fCancelCourseVisible = ((iScreenType = screenParentTable) Or (iScreenType = screenParentView)) And _
          gfTrainingBookingEnabled And _
          (pfrmCallingForm.TableID = glngCourseTableID)
        fCancelCourseEnabled = fCancelCourseVisible And (rsRecords.EditMode <> adEditAdd)
        ' JPD 5/3/01 Added check that enabled the 'Cancel Course' button only if the
        ' course is not already cancelled.
        If fCancelCourseEnabled Then
          fCancelCourseEnabled = False
          
          Set objTableView = pfrmCallingForm.TableView
          Set colColumnPrivileges = pfrmCallingForm.ColumnSelectPrivileges
          
          If colColumnPrivileges.IsValid(gsCourseCancelDateColumnName) Then
            If colColumnPrivileges.Item(gsCourseCancelDateColumnName).AllowSelect Then
              strSQL = "SELECT " & gsCourseCancelDateColumnName & _
                " FROM " & objTableView.RealSource & _
                " WHERE id = " & Trim(Str(pfrmCallingForm.RecordID))
              Set rsTemp = datGeneral.GetRecords(strSQL)
              
              If IsNull(rsTemp.Fields(gsCourseCancelDateColumnName).Value) Then
                fCancelCourseEnabled = True
              End If
              
              rsTemp.Close
              Set rsTemp = Nothing
            End If
          End If
        End If
    
        ' Absence Breakdown Record menu option
        ' JDM - 08/06/01 - Fault 2417 - Left ASRDEVELOPMENT flag in...
        fAbsenceReportsEnabled = _
            ((iScreenType = screenParentTable) Or _
             (iScreenType = screenParentView)) And _
            (pfrmCallingForm.TableID = glngPersonnelTableID) And _
            gfAbsenceEnabled
  
      End If
    
      If TypeOf pfrmCallingForm Is frmRecEdit4 Then
        bAccordResendVisible = datGeneral.SystemPermission("ACCORD", "RESEND") And gbAccordEnabled
        bAccordResendEnabled = IsTableMappedToAccord(pfrmCallingForm.TableView.TableID)
      Else
        bAccordResendVisible = False
        bAccordResendEnabled = False
      End If
    
      ' RH 03/11/00
      ' Should now do this if the calling form is the find window, because
      ' ive added filter/filterclear tools to the find window toolbar
      fFilterEnabled = True
      fFilterClearEnabled = pfrmCallingForm.Filtered
    
      Set rsRecords = Nothing
    End If
  End If
  
  With abMain
    .Tools("NewRecord").Enabled = fNewRecordEnabled
    .Tools("CopyRecord").Enabled = fCopyRecordEnabled
    .Tools("EditRecord").Enabled = fEditRecordEnabled
    .Tools("SaveRecord").Enabled = fSaveRecordEnabled
    .Tools("DeleteRecord").Enabled = fDeleteRecordEnabled
    .Tools("FirstRecord").Enabled = fFirstRecordEnabled
    .Tools("PreviousRecord").Enabled = fPreviousRecordEnabled
    .Tools("NextRecord").Enabled = fNextRecordEnabled
    .Tools("LastRecord").Enabled = fLastRecordEnabled
    .Tools("FindRecord").Enabled = fFindRecordEnabled
    .Tools("Refresh").Enabled = fRefreshEnabled
    .Tools("Order").Enabled = fOrderEnabled
    .Tools("Filter").Enabled = fFilterEnabled
    .Tools("FilterClear").Enabled = fFilterClearEnabled
    .Tools("QuickFind").Enabled = fQuickFindEnabled ' - RH 23/08/99
    ' NPG20090218 Fault 11512
    ' .Tools("QuickFind").Visible = fQuickFindEnabled '8774
    .Tools("QuickFind").Visible = .Tools("QuickFind").Visible And fQuickFindEnabled '8774
   
    ' Payroll resend options
    .Tools("ID_Accord_SendAs_New").Visible = bAccordResendVisible
    .Tools("ID_Accord_SendAs_New").Enabled = bAccordResendEnabled And Not fAddingNewRecord
    .Tools("ID_Accord_SendAs_Update").Visible = bAccordResendVisible
    .Tools("ID_Accord_SendAs_Update").Enabled = bAccordResendEnabled And Not fAddingNewRecord
   
    .Bands("bndRecordUtilities").Tools("MailMergeRec").Visible = fMailMergeExists
    .Bands("bndRecordUtilities").Tools("MailMergeRec").Enabled = MenuEnabled("MAILMERGE") And Not fAddingNewRecord And fSelectionMade
    
    .Bands("bndRecordUtilities").Tools("LabelsRec").Visible = fEnvelopeLabelsExists
    .Bands("bndRecordUtilities").Tools("LabelsRec").Enabled = MenuEnabled("LABELS") And Not fAddingNewRecord
    
    .Bands("bndRecordUtilities").Tools("DataTransferRec").Visible = fDataTransferExists
    .Bands("bndRecordUtilities").Tools("DataTransferRec").Enabled = MenuEnabled("DATATRANSFER") And Not fAddingNewRecord And fSelectionMade
    .Tools("Email").Visible = fEmailAddrExists
    .Tools("Email").Enabled = MenuEnabled("EMAILADDRESSES") And Not fAddingNewRecord
    
    .Bands("bndRecordReports").Tools("MatchReportRec").Visible = fMatchReportExists
    .Bands("bndRecordReports").Tools("MatchReportRec").Enabled = MenuEnabled("MATCHREPORTS") And Not fAddingNewRecord

    .Bands("bndRecordReports").Tools("SuccessionRec").Visible = fSuccessionPlanning
    .Bands("bndRecordReports").Tools("SuccessionRec").Enabled = MenuEnabled("SUCCESSION") And Not fAddingNewRecord

    .Bands("bndRecordReports").Tools("CareerRec").Visible = fCareerProgression
    .Bands("bndRecordReports").Tools("CareerRec").Enabled = MenuEnabled("CAREER") And Not fAddingNewRecord

    .Bands("bndRecordReports").Tools("RecordProfileRec").Visible = fRecordProfileExists
    .Bands("bndRecordReports").Tools("RecordProfileRec").Enabled = MenuEnabled("RECORDPROFILE") And Not fAddingNewRecord

    .Bands("bndRecordReports").Tools("CalendarReportRec").Visible = fCalendarReportExists
    .Bands("bndRecordReports").Tools("CalendarReportRec").Enabled = MenuEnabled("CALENDARREPORTS") And Not fAddingNewRecord And fSelectionMade
    
    'MH20030514 Should be visible but disabled if you don't access to run these reports.
    ' Only enable the absence reports if we are licensed
    .Bands("bndRecordReports").Tools("AbsenceCalendar").Visible = fAbsenceReportsEnabled
    .Bands("bndRecordReports").Tools("AbsenceBreakdownRec").Visible = fAbsenceReportsEnabled
    .Bands("bndRecordReports").Tools("BradfordIndexRec").Visible = fAbsenceReportsEnabled
    .Bands("bndRecordReports").Tools("AbsenceCalendar").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AC") And Not fAddingNewRecord
    .Bands("bndRecordReports").Tools("AbsenceBreakdownRec").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB") And Not fAddingNewRecord
    .Bands("bndRecordReports").Tools("BradfordIndexRec").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF") And Not fAddingNewRecord
        
    ' RH 28/09 - Should have the cancel course icon on the record
    '            menu as well as the recedit toolbar
    .Bands("bndRecordTrainingBooking").Tools("CancelCourse").Visible = fCancelCourseVisible
    .Bands("bndRecordTrainingBooking").Tools("CancelCourse").Enabled = fCancelCourseEnabled
    
    .Bands("bndRecordTrainingBooking").Tools("BookCourse").Visible = fBookCourseVisible
    .Bands("bndRecordTrainingBooking").Tools("BookCourse").Enabled = fBookCourseEnabled
    .Bands("bndRecordTrainingBooking").Tools("BulkBooking").Visible = fBulkBookingVisible
    .Bands("bndRecordTrainingBooking").Tools("BulkBooking").Enabled = fBulkBookingEnabled
    .Bands("bndRecordTrainingBooking").Tools("AddFromWaitingList").Visible = fAddFromWaitingListVisible
    .Bands("bndRecordTrainingBooking").Tools("AddFromWaitingList").Enabled = fAddFromWaitingListEnabled
    .Bands("bndRecordTrainingBooking").Tools("TransferBooking").Visible = fTransferBookingVisible
    .Bands("bndRecordTrainingBooking").Tools("TransferBooking").Enabled = fTransferBookingEnabled
    .Bands("bndRecordTrainingBooking").Tools("CancelBooking").Visible = fCancelBookingVisible
    .Bands("bndRecordTrainingBooking").Tools("CancelBooking").Enabled = fCancelBookingEnabled
    
    ' JPD20021126 Fault 4805
    .Tools("ID_Print").Visible = (Not TypeOf pfrmCallingForm Is frmRecEdit4)
    
    ' Display the sub-menus as required.
    fSomeVisible = False
    fSomeEnabled = False
    For Each objTool In .Bands("bndRecordReports").Tools
      If objTool.Enabled Then fSomeEnabled = True
      If objTool.Visible Then fSomeVisible = True
    Next objTool
    .Tools("mnuRecordReports").Enabled = fSomeEnabled
    .Tools("mnuRecordReports").Visible = fSomeVisible
    
    fSomeVisible = False
    fSomeEnabled = False
    For Each objTool In .Bands("bndRecordUtilities").Tools
      If objTool.Enabled Then fSomeEnabled = True
      If objTool.Visible Then fSomeVisible = True
    Next objTool
    .Tools("mnuRecordUtilities").Enabled = fSomeEnabled
    .Tools("mnuRecordUtilities").Visible = fSomeVisible
    
    fSomeVisible = False
    fSomeEnabled = False
    For Each objTool In .Bands("bndRecordTrainingBooking").Tools
      If objTool.Enabled Then fSomeEnabled = True
      If objTool.Visible Then fSomeVisible = True
    Next objTool
    .Tools("mnuRecordTrainingBooking").Enabled = fSomeEnabled
    .Tools("mnuRecordTrainingBooking").Visible = fSomeVisible

    ' Mark the required items as beginning a group.
    fBeginGroupDone = False
    fBeginGroup = .Tools("Email").Visible And (Not fBeginGroupDone)
    .Bands("mnuRecord").Tools("Email").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Tools("mnuRecordReports").Visible And (Not fBeginGroupDone)
    .Bands("mnuRecord").Tools("mnuRecordReports").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Tools("mnuRecordUtilities").Visible And (Not fBeginGroupDone)
    .Bands("mnuRecord").Tools("mnuRecordUtilities").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Tools("mnuRecordTrainingBooking").Visible And (Not fBeginGroupDone)
    .Bands("mnuRecord").Tools("mnuRecordTrainingBooking").BeginGroup = fBeginGroup
    
    ' Mark the required sub-menu items as beginning a group.
    fBeginGroupDone = False
    fBeginGroup = .Bands("bndRecordReports").Tools("MatchReportRec").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("MatchReportRec").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Bands("bndRecordReports").Tools("SuccessionRec").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("SuccessionRec").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Bands("bndRecordReports").Tools("CareerRec").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("CareerRec").BeginGroup = fBeginGroup
    
    fBeginGroupDone = False
    fBeginGroup = .Bands("bndRecordReports").Tools("AbsenceBreakdownRec").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("AbsenceBreakdownRec").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Bands("bndRecordReports").Tools("AbsenceCalendar").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("AbsenceCalendar").BeginGroup = fBeginGroup
    fBeginGroupDone = fBeginGroupDone Or fBeginGroup
    
    fBeginGroup = .Bands("bndRecordReports").Tools("BradfordIndexRec").Visible And (Not fBeginGroupDone)
    .Bands("bndRecordReports").Tools("BradfordIndexRec").BeginGroup = fBeginGroup

    .Refresh
    .RecalcLayout
  End With
  
  ' Now do the same for the toolbar on the recedit screen
  With pfrmCallingForm.ActiveBar1
    .Tools("NewRecord").Enabled = fNewRecordEnabled
    .Tools("CopyRecord").Enabled = fCopyRecordEnabled
    .Tools("SaveRecord").Enabled = fSaveRecordEnabled
    .Tools("DeleteRecord").Enabled = fDeleteRecordEnabled
    .Tools("FirstRecord").Enabled = fFirstRecordEnabled
    .Tools("PreviousRecord").Enabled = fPreviousRecordEnabled
    .Tools("NextRecord").Enabled = fNextRecordEnabled
    .Tools("LastRecord").Enabled = fLastRecordEnabled
    .Tools("FindRecord").Enabled = fFindRecordEnabled
    .Tools("Refresh").Enabled = fRefreshEnabled
    .Tools("Order").Enabled = fOrderEnabled
    .Tools("Filter").Enabled = fFilterEnabled
    '.Tools("LabelsRec").Enabled = fEnvelopeLabelsExists
    
''''Set the begingroup property for the right tool
'''If fMailMergeExists Then
'''  .Bands(0).Tools("MailMerge").BeginGroup = True
'''ElseIf fDataTransferExists Then
'''  .Bands(0).Tools("DataTransfer").BeginGroup = True
'''ElseIf fEmailAddrExists Then
'''  .Bands(0).Tools("Email").BeginGroup = True
''''ElseIf fAbsenceReportsEnabled Then
''''  .Bands(0).Tools("AbsenceBreakdownRec").BeginGroup = True
'''End If
'''    .Bands(0).Refresh
    
    'JPD 20050113 Fault 8218
    '.Tools("CancelCourse").Visible = fCancelCourseVisible
    .Tools("CancelCourse").Visible = .Tools("CancelCourse").Visible And fCancelCourseVisible
    .Tools("CancelCourse").Enabled = fCancelCourseEnabled
    
    .Tools("QuickFind").Enabled = fQuickFindEnabled ' - RH 23/08/99
    ' NPG20090218 Fault 11512
    ' .Tools("QuickFind").Visible = fQuickFindEnabled  '8774
    .Tools("QuickFind").Visible = .Tools("QuickFind").Visible And fQuickFindEnabled  '8774
    
    If (TypeOf pfrmCallingForm Is frmRecEdit4) Then
      .Tools("FilterClear").Enabled = fFilterClearEnabled
      .Tools("MailMergeRec").Visible = fMailMergeExists And .Tools("MailMergeRec").Visible
      .Tools("MailMergeRec").Enabled = MenuEnabled("MAILMERGE") And Not fAddingNewRecord And fSelectionMade
      .Tools("LabelsRec").Visible = fEnvelopeLabelsExists And .Tools("LabelsRec").Visible
      .Tools("LabelsRec").Enabled = MenuEnabled("LABELS") And Not fAddingNewRecord
      .Tools("DataTransferRec").Visible = fDataTransferExists And .Tools("DataTransferRec").Visible
      .Tools("DataTransferRec").Enabled = MenuEnabled("DATATRANSFER") And Not fAddingNewRecord And fSelectionMade
      .Tools("Email").Visible = fEmailAddrExists And .Tools("Email").Visible
      .Tools("Email").Enabled = MenuEnabled("EMAILADDRESSES") And Not fAddingNewRecord
      .Tools("MatchReportRec").Visible = fMatchReportExists And .Tools("MatchReportRec").Visible
      .Tools("MatchReportRec").Enabled = MenuEnabled("MATCHREPORTS") And Not fAddingNewRecord
      .Tools("SuccessionRec").Visible = fSuccessionPlanning And .Tools("SuccessionRec").Visible
      .Tools("SuccessionRec").Enabled = MenuEnabled("SUCCESSION") And Not fAddingNewRecord
      .Tools("CareerRec").Visible = fCareerProgression And .Tools("CareerRec").Visible
      .Tools("CareerRec").Enabled = MenuEnabled("CAREER") And Not fAddingNewRecord
      .Tools("RecordProfileRec").Visible = fRecordProfileExists And .Tools("RecordProfileRec").Visible
      .Tools("RecordProfileRec").Enabled = MenuEnabled("RECORDPROFILE") And Not fAddingNewRecord
      .Tools("CalendarReportRec").Visible = fCalendarReportExists And .Tools("CalendarReportRec").Visible
      .Tools("CalendarReportRec").Enabled = MenuEnabled("CALENDARREPORTS") And Not fAddingNewRecord And fSelectionMade
      
      ' Only enable the absence reports dropdown if we are licensed
      .Bands(0).Tools("AbsenceCalendar").Visible = fAbsenceReportsEnabled And .Bands(0).Tools("AbsenceCalendar").Visible
      .Bands(0).Tools("AbsenceBreakdownRec").Visible = fAbsenceReportsEnabled And .Bands(0).Tools("AbsenceBreakdownRec").Visible
      .Bands(0).Tools("BradfordFactorRec").Visible = fAbsenceReportsEnabled And .Bands(0).Tools("BradfordFactorRec").Visible
      .Bands(0).Tools("AbsenceCalendar").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AC") And Not fAddingNewRecord
      .Bands(0).Tools("AbsenceBreakdownRec").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB") And Not fAddingNewRecord
      .Bands(0).Tools("BradfordFactorRec").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF") And Not fAddingNewRecord
    
      ' Mark the required items as beginning a group.
      fBeginGroupDone = False
      fBeginGroup = .Tools("CalendarReportRec").Visible And (Not fBeginGroupDone)
      .Tools("CalendarReportRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
      
      fBeginGroup = .Tools("RecordProfileRec").Visible And (Not fBeginGroupDone)
      .Tools("RecordProfileRec").BeginGroup = fBeginGroup
          
      fBeginGroupDone = False
      fBeginGroup = .Tools("MatchReportRec").Visible And (Not fBeginGroupDone)
      .Tools("MatchReportRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
          
      fBeginGroup = .Tools("SuccessionRec").Visible And (Not fBeginGroupDone)
      .Tools("SuccessionRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
          
      fBeginGroup = .Tools("CareerRec").Visible And (Not fBeginGroupDone)
      .Tools("CareerRec").BeginGroup = fBeginGroup
          
      fBeginGroupDone = False
      fBeginGroup = .Tools("AbsenceBreakdownRec").Visible And (Not fBeginGroupDone)
      .Tools("AbsenceBreakdownRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
          
      fBeginGroup = .Tools("AbsenceCalendar").Visible And (Not fBeginGroupDone)
      .Tools("AbsenceCalendar").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
          
      fBeginGroup = .Tools("BradfordFactorRec").Visible And (Not fBeginGroupDone)
      .Tools("BradfordFactorRec").BeginGroup = fBeginGroup
              
      fBeginGroupDone = False
      fBeginGroup = .Tools("LabelsRec").Visible And (Not fBeginGroupDone)
      .Tools("LabelsRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
      
      fBeginGroup = .Tools("MailMergeRec").Visible And (Not fBeginGroupDone)
      .Tools("MailMergeRec").BeginGroup = fBeginGroup
      fBeginGroupDone = fBeginGroupDone Or fBeginGroup
          
      fBeginGroup = .Tools("DataTransferRec").Visible And (Not fBeginGroupDone)
      .Tools("DataTransferRec").BeginGroup = fBeginGroup
        
    ElseIf (TypeOf pfrmCallingForm Is frmFind2) Then
      .Bands(0).Tools("FilterClear").Enabled = fFilterClearEnabled
      .Bands(0).Tools("EditFind").Enabled = fEditRecordEnabled
      .Bands(0).Tools("DeleteFind").Enabled = fDeleteRecordEnabled
      .Bands(0).Tools("BookCourseFind").Enabled = fBookCourseEnabled
      .Bands(0).Tools("TransferFind").Enabled = fTransferBookingEnabled
      .Bands(0).Tools("CancelBookingFind").Enabled = fCancelBookingEnabled
      .Bands(0).Tools("BulkBookingFind").Enabled = fBulkBookingEnabled
      .Bands(0).Tools("AddFromWaitingListFind").Enabled = fAddFromWaitingListEnabled
      .Bands(0).Tools("ID_Print").Enabled = True
      
      .Bands(0).Tools("CustomReports").Visible = fCustomReportExists And (.Bands(0).Tools("CustomReports").Visible)
      .Bands(0).Tools("CustomReports").Enabled = fSelectionMade
      
      .Bands(0).Tools("CalendarReports").Visible = fCalendarReportExists And (.Bands(0).Tools("CalendarReports").Visible)
      .Bands(0).Tools("CalendarReports").Enabled = fSelectionMade
      .Bands(0).Tools("CalendarReports").BeginGroup = Not fCustomReportExists
      
      .Bands(0).Tools("GlobalUpdate").Visible = fGlobalUpdateExists And (.Bands(0).Tools("GlobalUpdate").Visible)
      .Bands(0).Tools("GlobalUpdate").Enabled = fSelectionMade
      '.Bands(0).Tools("GlobalUpdate").BeginGroup = (Not fCustomReportExists And Not fCalendarReportExists)

      .Bands(0).Tools("DataTransfer").Visible = fDataTransferExists And (.Bands(0).Tools("DataTransfer").Visible)
      .Bands(0).Tools("DataTransfer").Enabled = fSelectionMade

      .Bands(0).Tools("MailMerge").Visible = fMailMergeExists And (.Bands(0).Tools("MailMerge").Visible)
      .Bands(0).Tools("MailMerge").Enabled = fSelectionMade

      ' Recalculate the new utility/report group separators
      strGroupType = ""
      For Each objTool In .Bands(0).Tools
        If (objTool.Name = "GlobalUpdate" Or objTool.Name = "DataTransfer" Or objTool.Name = "MailMerge") _
          And (strGroupType = "" Or strGroupType <> "Utilities") Then
            ' set the objtool to begingroup and change the flag
            objTool.BeginGroup = True
            strGroupType = "Utilities"
        End If
            
        If (objTool.Name = "CalendarReports" Or objTool.Name = "CustomReports") _
          And (strGroupType = "" Or strGroupType <> "Reports") Then
            ' set the objtool to begingroup and change the flag
            objTool.BeginGroup = True
            strGroupType = "Reports"
        End If
      Next objTool

    End If

    .Refresh
    .RecalcLayout
  End With
 
TidyUpAndExit:
  gobjErrorStack.PopStack
  
  Exit Sub
ErrorTrap:

  If (Err.Number = 438) Or (Err.Number = 2006) Then
    Resume Next
  Else
    
    ' JDM - 01/05/01 - Fault 2220 - Messes up when history record is deleted by current user.
    '                               Ignoring the fault seems to fix it.
    If (Err.Number = 3021) Then
      GoTo TidyUpAndExit
    Else
      gobjErrorStack.HandleError "(fNewRecordEnabled, fCopyRecordEnabled, fEditRecordEnabled" _
        & ", fSaveRecordEnabled, fDeleteRecordEnabled, fFirstRecordEnabled, fPreviousRecordEnabled" _
        & ", fNextRecordEnabled, fLastRecordEnabled, fFindRecordEnabled, fRefreshEnabled, fOrderEnabled" _
        & ", fFilterEnabled,fFilterClearEnabled, fCancelCourseEnabled, fCancelCourseVisible, fMailMergeExists" _
        & ", fDataTransferExists, fEmailAddrExists, iScreenType, rsRecords, vBookMark, fQuickFindEnabled" _
        & ", objTool, flag, fAbsenceReportsEnabled, rsTemp, strSQL, objTableView, colColumnPrivileges, fAddingNewRecord)" _
      , Array(fNewRecordEnabled, fCopyRecordEnabled, fEditRecordEnabled, fSaveRecordEnabled _
        , fDeleteRecordEnabled, fFirstRecordEnabled, fPreviousRecordEnabled, fNextRecordEnabled _
        , fLastRecordEnabled, fFindRecordEnabled, fRefreshEnabled, fOrderEnabled, fFilterEnabled _
        , fFilterClearEnabled, fCancelCourseEnabled, fCancelCourseVisible, fMailMergeExists _
        , fDataTransferExists, fEmailAddrExists, iScreenType, rsRecords, vBookMark, fQuickFindEnabled _
        , objTool, flag, fAbsenceReportsEnabled, rsTemp, strSQL, objTableView _
        , colColumnPrivileges, fAddingNewRecord)
    End If
  End If

End Sub

Private Sub UpdateStatusBar(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)
  ' Update the status bar.
  On Error GoTo Err_Trap
    
  Dim fFiltered As Boolean
  Dim sMainText As String
  Dim sCaption As String
  
  fFiltered = False
  sMainText = ""
  
  If Not pfUnLoad Then
    With pfrmCallingForm
      fFiltered = .Filtered
  
      Select Case .ScreenType
        Case screenParentTable, screenLookup, screenQuickEntry, screenParentView, screenHistoryTable, screenHistoryView
          sCaption = .StatusCaption
          If .Recordset.EditMode = adEditAdd Then
            sMainText = sCaption & " - Adding New Record."
          Else
            sMainText = sCaption & " - Record " & .Recordset.AbsolutePosition & _
              IIf(Not IsMissing(.RecordCount), " of " & .RecordCount, "") & _
              IIf(fFiltered, " (Filtered)", "")
          End If
        
        Case screenFind, screenHistorySummary
          sMainText = .StatusCaption & IIf(fFiltered, " (Filtered)", "")
      End Select
    End With
  End If
  
  With stbMain
    .Panels("pnlMAIN").Text = sMainText
    .Panels("pnlFILTER").Enabled = fFiltered
  End With
  
  Exit Sub
    
Err_Trap:
  If Err.Number = 3021 Then
    Exit Sub
  End If

End Sub


Public Property Get Loading() As Boolean

    Loading = mbLoading

End Property

Public Property Let Loading(ByVal bLoading As Boolean)

    mbLoading = bLoading

End Property

Public Property Get Changing() As Boolean

    Changing = mbChanging
    
End Property

Public Property Let Changing(ByVal bChange As Boolean)

    mbChanging = bChange

End Property

Public Sub PickListClick()
  ' Display the Picklist definition form.
  Dim fExit As Boolean
  'Dim sSQL As String
  Dim frmPick As frmPicklists
  Dim frmSelection As frmDefSel
    
  Screen.MousePointer = vbHourglass
  fExit = False
  
  Set frmSelection = New frmDefSel
  
  With frmSelection
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
      .EnableRun = False
      .TableComboEnabled = True
      .TableComboVisible = True
      
      If .ShowList(utlPicklist) Then
      
        .CustomShow vbModal
        
        Select Case .Action
          Case edtAdd
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(True, .FromCopy, .TableID) Then
              frmPick.Show vbModal
            End If
            .SelectedID = frmPick.SelectedID
            Unload frmPick
            Set frmPick = Nothing
  
          Case edtEdit
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(False, .FromCopy, .TableID, .SelectedID) Then
              frmPick.Show vbModal
            End If
            If .FromCopy And frmPick.SelectedID > 0 Then
              .SelectedID = frmPick.SelectedID
            End If
            Unload frmPick
            Set frmPick = Nothing
            
          Case edtPrint
            Set frmPick = New frmPicklists
            frmPick.PrintDef .TableID, .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          
          Case 0
            fExit = True
        
        End Select
      End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

  RefreshMainForm Me, False

End Sub


Public Sub AbsenceCalendarClick()

  frmAbsenceCalendar.Initialise
  Unload frmAbsenceCalendar
  Set frmAbsenceCalendar = Nothing
  
End Sub


Public Sub DataTransferClick()
  ' Display the Data Transfer selection form.
  Dim fExit As Boolean
  Dim frmEdit As frmDataTransfer
  Dim frmSelection As frmDefSel
  
  Dim objExecution As clsDataTransferRun
  
  Screen.MousePointer = vbHourglass
    
  Set frmSelection = New frmDefSel
  fExit = False
  
  With frmSelection
    Do While Not fExit
      
      .EnableRun = True
      
      If .ShowList(utlDataTransfer) Then
        .CustomShow vbModal
            
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmDataTransfer
            frmEdit.Initialise True, .FromCopy
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
            
          Case edtEdit
            Set frmEdit = New frmDataTransfer
            frmEdit.Initialise False, .FromCopy, .SelectedID
            If Not frmEdit.Cancelled Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
  
'          Case edtDelete
'            Set frmEdit = New frmDataTransfer
'            frmEdit.Initialise False, .FromCopy, .SelectedID
'            If Not frmEdit.Cancelled Then
'              datGeneral.DeleteRecord "ASRSysDataTransferName", "DataTransferID", .SelectedID
'              datGeneral.DeleteRecord "ASRSysDataTransferColumns", "DataTransferID", .SelectedID
'            End If
'            Unload frmEdit
'            Set frmEdit = Nothing
  
          Case edtSelect
            'FOR NORMAL RUNNING...
            'datGeneral.RunDataTransfer .SelectedID, False
            
            'FOR SILENT MODE...
            'datGeneral.RunDataTransfer .SelectedID, True
            
            Set objExecution = New clsDataTransferRun
            objExecution.ExecuteDataTransfer .SelectedID
            Set objExecution = Nothing
            fExit = gbCloseDefSelAfterRun
          
          Case edtPrint
            Set frmEdit = New frmDataTransfer
            frmEdit.Initialise False, False, .SelectedID, True
            If Not frmEdit.Cancelled Then
              frmEdit.PrintDef .SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing

          Case edtCancel
            fExit = True
        End Select
      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Public Sub FilterClick()
  
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  Dim lngOptions As Long
  
  Set objExpression = New clsExprExpression
    
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
  
  With objExpression
    fOK = .Initialise(0, 0, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    
    If fOK Then
      .SelectExpression False, lngOptions
    End If
  End With
  
  Set objExpression = Nothing

  RefreshMainForm Me, False

End Sub

Public Sub GlobalClick(FormType As GlobalType)
  
  Dim fExit As Boolean
  Dim frmEdit As frmGlobalFunctions
  Dim frmSelection As frmDefSel
  Dim objGlobalRun As clsGlobalRun
  Dim blnOK As Boolean
  Dim lngTYPE As Long

  Screen.MousePointer = vbHourglass
    
  'sType = Choose(FormType, "ADD", "UPDATE", "DELETE")
  lngTYPE = Choose(FormType, UtlGlobalAdd, utlGlobalUpdate, utlGlobalDelete)

  fExit = False
  Set frmSelection = New frmDefSel

  With frmSelection
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
      .EnableRun = True

      If .ShowList(lngTYPE) Then
        .CustomShow vbModal

        Select Case .Action
        Case edtAdd
          Set frmEdit = New frmGlobalFunctions
          If frmEdit.Initialise(True, .FromCopy, FormType) Then
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
          End If
          Unload frmEdit
          Set frmEdit = Nothing
        
        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmEdit = New frmGlobalFunctions
          If frmEdit.Initialise(False, .FromCopy, FormType, .SelectedID) Then
            If Not frmEdit.Cancelled Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
          End If
          Unload frmEdit
          Set frmEdit = Nothing

'        'TM20010808 Fault 2656 - Must validate the definition before allowing the delete.
'        Case edtDelete
'          Set frmEdit = New frmGlobalFunctions
'          frmEdit.Initialise False, False, FormType, .SelectedID
'          If Not frmEdit.Cancelled Then
'            datGeneral.DeleteRecord "ASRSysGlobalFunctions", "FunctionID", .SelectedID
'            datGeneral.DeleteRecord "ASRSysGlobalItems", "FunctionID", .SelectedID
'          End If
'          Unload frmEdit
'          Set frmEdit = Nothing

        Case edtSelect
          
          'Select Case sType
          'Case "ADD", "UPDATE"
            'Set objGlobalAddUpdate = New clsGlobalAddUpdateRun
            ''blnOK = objGlobalAddUpdate.RunGlobalAddUpdate(.SelectedID, False, FormType)
            'blnOK = objGlobalAddUpdate.RunGlobal(.SelectedID, False, FormType)
            'Set objGlobalAddUpdate = Nothing
            Set objGlobalRun = New clsGlobalRun
            'blnOK = objGlobalAddUpdate.RunGlobalAddUpdate(.SelectedID, False, FormType)
            blnOK = objGlobalRun.RunGlobal(.SelectedID, FormType, "")
            Set objGlobalRun = Nothing
            fExit = gbCloseDefSelAfterRun

          'Case "DELETE"
            'Set objGlobalDelete = New clsGlobalDeleteRun
            'blnOK = objGlobalDelete.RunGlobalDelete(.SelectedID, False)
            'Set objGlobalDelete = Nothing
          
          'End Select

        'TM20010808 Fault 2656 - Must validate the definition before allowing the print.
        Case edtPrint
          Set frmEdit = New frmGlobalFunctions
          frmEdit.Initialise False, False, FormType, .SelectedID, True
          If Not frmEdit.Cancelled Then
            frmEdit.PrintDef FormType, .SelectedID
          End If
          Unload frmEdit
          Set frmEdit = Nothing
        
        Case edtCancel
          fExit = True

        End Select
      End If

    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Public Sub ImportClick()

  ' Set reference to the Run class
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmImport
  Dim pobjImport As clsImportRUN
  Dim fExit As Boolean
    
  Screen.MousePointer = vbHourglass
    
  fExit = False
    
  Set frmSelection = New frmDefSel
  
  With frmSelection
    
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
        
      .EnableRun = True
        
      If .ShowList(utlImport) Then
        
        .CustomShow vbModal
            
        Select Case .Action
          Case edtAdd
            
            Set frmEdit = New frmImport
            If frmEdit.Initialise(True, .FromCopy) Then
              frmEdit.Show vbModal
              .SelectedID = frmEdit.SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
                    
          Case edtEdit
            
            Set frmEdit = New frmImport
            
            If frmEdit.Initialise(False, .FromCopy, .SelectedID) Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
                    
          Case edtSelect

            Set pobjImport = New clsImportRUN
            pobjImport.ImportID = .SelectedID
            pobjImport.RunImport
            fExit = gbCloseDefSelAfterRun
            Set pobjImport = Nothing

          Case edtPrint
            Set frmEdit = New frmImport
            frmEdit.PrintDef .SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
            
          Case edtCancel
            fExit = True
        End Select
      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing
  
End Sub


Public Sub WorkflowClick()
  Dim frmSelection As frmDefSel
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fExit As Boolean
  Dim lngInstanceID As Long
  Dim sFormElements As String
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim asForms() As String
  Dim fRunning As Boolean
  Dim strExePath As String
  Dim fIsDLL As Boolean
  Dim sURL As String
  Dim sUser As String
  Dim sPassword As String
  Dim sMessage As String
  
  On Error GoTo ErrorTrap
  
  ' Check the URL has ben defined.
  sURL = WorkflowURL
  If Len(sURL) = 0 Then
    COAMsgBox "No Workflow URL has been configured. Contact your system administrator.", vbInformation + vbOKOnly, "Workflow"
    Exit Sub
  End If
  
  ReadWebLogon sUser, sPassword
  If Len(sUser) = 0 Then
    COAMsgBox "No Workflow Web Logon has been configured. Contact your system administrator.", vbInformation + vbOKOnly, "Workflow"
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass

  fExit = False
  fRunning = False
  strExePath = GetDefaultBrowserApplication(fIsDLL)
  
  Set frmSelection = New frmDefSel

  With frmSelection
    ' Loop until the operation has been cancelled.
    Do While Not fExit
      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlWorkflow) Then
        .CustomShow vbModal

        Select Case .Action
          Case edtSelect
            fRunning = True
            
            Set cmADO = New ADODB.Command
            With cmADO
              .CommandText = "dbo.spASRInstantiateWorkflow"
              .CommandType = adCmdStoredProc
              .CommandTimeout = 0
              Set .ActiveConnection = gADOCon

              Set pmADO = .CreateParameter("WorkflowID", adInteger, adParamInput)
              .Parameters.Append pmADO
              pmADO.Value = frmSelection.SelectedID
    
              Set pmADO = .CreateParameter("InstanceID", adInteger, adParamOutput)
              .Parameters.Append pmADO

              Set pmADO = .CreateParameter("FormElements", adVarChar, adParamOutput, VARCHAR_MAX_Size)
              .Parameters.Append pmADO

              Set pmADO = .CreateParameter("Message", adVarChar, adParamOutput, VARCHAR_MAX_Size)
              .Parameters.Append pmADO

              cmADO.Execute

              lngInstanceID = IIf(IsNull(.Parameters("InstanceID").Value), 0, .Parameters("InstanceID").Value)
              sFormElements = IIf(IsNull(.Parameters("FormElements").Value), vbNullString, .Parameters("FormElements").Value)
              sMessage = IIf(IsNull(.Parameters("Message").Value), vbNullString, .Parameters("Message").Value)
            End With
            Set cmADO = Nothing
            
            If Len(sMessage) > 0 Then
              ' Instantiation failed for some reason. Tell the user why.
              COAMsgBox "Workflow : '" & .SelectedText & "' initiation failed." & vbCrLf & vbCrLf & sMessage, vbExclamation + vbOKOnly, "Workflow"
            Else
              ' Launch the default browser to hit the HR Pro Workflow webservice
              ' passing in the InstanceID and the InstanceStepID
              ' This is done for each form element that needs to be displayed.
              ReDim asForms(0)
              Do While InStr(sFormElements, vbTab) > 0
                iIndex = InStr(sFormElements, vbTab)
                    
                ReDim Preserve asForms(UBound(asForms) + 1)
                asForms(UBound(asForms)) = Left(sFormElements, iIndex - 1)
                
                sFormElements = Mid(sFormElements, iIndex + 1)
              Loop
  
              ' Inform the user that the Workflow has been initiated.
              If UBound(asForms) > 0 Then
                For iLoop = 1 To UBound(asForms)
                  OpenWebForm lngInstanceID, CLng(asForms(iLoop))
                Next iLoop
                
                If Len(Trim(strExePath)) > 1 Then
                  'JPD 20071205 Fault 12680
                  'COAMsgBox "Workflow : '" & .SelectedText & "' initiated successfully." & vbCrLf & vbCrLf & "Please complete the required Workflow forms.", vbInformation + vbOKOnly, "Workflow"
                Else
                  COAMsgBox "Workflow : '" & .SelectedText & "' initiated successfully, but unable to open required Workflow forms." & vbCrLf & vbCrLf & "Please contact your system administrator.", vbExclamation + vbOKOnly, "Workflow"
                End If
              Else
                COAMsgBox "Workflow : '" & .SelectedText & "' initiated successfully.", vbInformation + vbOKOnly, "Workflow"
              End If
  
              Me.SetFocus
  
              fExit = gbCloseDefSelAfterRun
              fRunning = False
            End If
            
          Case edtCancel
            fExit = True
        End Select
      End If
    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

  Exit Sub

ErrorTrap:
  COAMsgBox "Error " & IIf(fRunning, "running Workflow.", "displaying Workflows.") & vbCrLf & _
    Err.Description, _
    vbOKOnly + vbExclamation, Application.Name

End Sub

Public Sub ExportClick()
  
  ' Set reference to the Run class
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmExport
  Dim pobjExport As clsExportRUN
  Dim fExit As Boolean
    
  Screen.MousePointer = vbHourglass
    
  fExit = False
    
  Set frmSelection = New frmDefSel
  
  With frmSelection
    
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
        
      .EnableRun = True
        
      If .ShowList(utlExport) Then
        
        .CustomShow vbModal
            
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmExport
            If frmEdit.Initialise(True, .FromCopy) Then
              frmEdit.Show vbModal
              .SelectedID = frmEdit.SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
                    
          Case edtEdit
            
            Set frmEdit = New frmExport
            If frmEdit.Initialise(False, .FromCopy, .SelectedID) Then
              If Not frmEdit.Cancelled Then
                frmEdit.Show vbModal
                If .FromCopy And frmEdit.SelectedID > 0 Then
                  .SelectedID = frmEdit.SelectedID
                End If
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
                    
'          'TM20010808 Fault 2656 - Must validate the definition before allowing the delete.
'          Case edtDelete
'            Set frmEdit = New frmExport
'            frmEdit.Initialise False, False, .SelectedID
'            If Not frmEdit.Cancelled Then
'              datGeneral.DeleteRecord "AsrSysExportName", "ID", .SelectedID
'              datGeneral.DeleteRecord "AsrSysExportDetails", "ExportID", .SelectedID
'            End If
'            Unload frmEdit
'            Set frmEdit = Nothing
                    
          Case edtSelect

            Set pobjExport = New clsExportRUN
            pobjExport.ExportID = .SelectedID

            'TO RUN AS NORMAL
            pobjExport.RunExport
            fExit = gbCloseDefSelAfterRun
            Set pobjExport = Nothing

          Case edtPrint
            Set frmEdit = New frmExport
            frmEdit.Initialise False, False, .SelectedID, True
            If Not frmEdit.Cancelled Then
              frmEdit.PrintDef .SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
                    
          Case edtCancel
            fExit = True
        End Select
      End If
    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing
  
End Sub


Public Sub BatchJobsClick()
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmBatchJob
    
  Screen.MousePointer = vbHourglass
  
  Set frmSelection = New frmDefSel
  fExit = False
  
  With frmSelection
    Do While Not fExit
      
      .EnableRun = True
      
      If .ShowList(utlBatchJob) Then
        
        .CustomShow vbModal
        DoEvents
        
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmBatchJob
            frmEdit.Initialise True, .FromCopy
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
                      
          Case edtEdit
            Set frmEdit = New frmBatchJob
            If frmEdit.Initialise(False, .FromCopy, .SelectedID) Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
              
          Case edtSelect
            Dim pobjBatchJobRUN As clsBatchJobRUN
            Dim plngEventLogID As Long
            Dim strNotes As String
            Set pobjBatchJobRUN = New clsBatchJobRUN
            strNotes = pobjBatchJobRUN.RunBatchJob(.SelectedID, .SelectedText, plngEventLogID)


'MH20030818 Fault 5673
'            If InStr(UCase(strNotes), "SUCCESSFULLY") And InStr(UCase(strNotes), "FAILED") = 0 Then
'            'If InStr(UCase(strNotes), "SUCCESS") Then
'              COAMsgBox "Batch Job : " & .SelectedText & " Completed successfully.", vbInformation + vbOKOnly, "Batch Jobs"
'            ElseIf InStr(UCase(strNotes), "CANCELLED") Then
'              COAMsgBox "Batch Job : " & .SelectedText & " Cancelled by user.", vbExclamation + vbOKOnly, "Batch Jobs"
'            Else
'              COAMsgBox "Batch Job : " & .SelectedText & " " & vbCrLf & strNotes, vbExclamation + vbOKOnly, "Batch Jobs"
'            End If
            Select Case pobjBatchJobRUN.JobStatus
            Case elsSuccessful
              COAMsgBox "Batch Job : '" & .SelectedText & "' Completed successfully.", vbInformation + vbOKOnly, "Batch Jobs"
            Case elsCancelled
              COAMsgBox "Batch Job : '" & .SelectedText & "' Cancelled by user.", vbExclamation + vbOKOnly, "Batch Jobs"
            Case Else
              COAMsgBox "Batch Job : '" & .SelectedText & "' Failed." & vbCrLf & vbCrLf & strNotes, vbExclamation + vbOKOnly, "Batch Jobs"
            End Select

            Set pobjBatchJobRUN = Nothing
            fExit = gbCloseDefSelAfterRun
            
          Case edtPrint
            Set frmEdit = New frmBatchJob
            frmEdit.PrintDef .SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
            
          Case edtCancel
            fExit = True
  
        End Select
      End If
    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

  Screen.MousePointer = vbNormal

  '# RH090300 To prevent toolbar locking after batch jobs
  With abMain
    .ResetHooks
    .Refresh
  End With
  
End Sub


Public Sub RefreshMainForm(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)
  ' Refresh the menu bar and the status bar.
  Dim iFormCount As Integer
  Dim fRecEditsLeft As Boolean
  Dim iMinWinCount As Integer
  Dim iNormWinCount As Integer
  Dim iVisibleWinCount As Integer
  Dim frmForm As Form
  
  'JPD 20030905 Fault 5184
  If mfMenuDisabled Then
    Exit Sub
  End If

  'MH20050516 Fault 9978
  CheckForNonactiveForms pfrmCallingForm

  'MH20031002 Fault 7083
  If Not pfrmCallingForm Is Nothing Then

    If (TypeOf pfrmCallingForm Is frmRecEdit4) Or _
      (TypeOf pfrmCallingForm Is frmFind2) Then
      
      iFormCount = Forms.Count - IIf(pfUnLoad, 1, 0)
      
      ' Refresh the menu bar.
      With abMain
        ' JPD 6/6/00 Check that the remaining forms are recEdit forms.
        fRecEditsLeft = False
        If iFormCount > 1 Then
          For Each frmForm In Forms
            If TypeOf frmForm Is frmRecEdit4 Then
              If (Not frmForm Is pfrmCallingForm) Or (Not pfUnLoad) Then
                fRecEditsLeft = True
                EnableActiveBar frmForm.ActiveBar1, False
              End If
            End If
          Next frmForm
          Set frmForm = Nothing
        End If
        .Tools("mnuRecord").Visible = fRecEditsLeft
        .Tools("mnuHistory").Visible = fRecEditsLeft
        .Tools("mnuReports").Visible = True
        .Tools("mnuReports").Enabled = True
        .Bands("bndAccord").Tools("ID_Accord_Current").Enabled = fRecEditsLeft And (TypeOf pfrmCallingForm Is frmRecEdit4)

        'MH20060616 Fault 11084
        .Bands("mnuAdministration").Tools("ChangePassword").Enabled = Not fRecEditsLeft

      End With

      'Refresh the Window Menu
      If iFormCount > 1 Then
        
        'Arrange Icons Bit
        For iMinWinCount = 0 To iFormCount - 1
          If (Forms(iMinWinCount).WindowState = vbMinimized) And Forms(iMinWinCount).Visible Then
            abMain.Bands("mnuWindow").Tools("Arrange").Enabled = True
          Else
            abMain.Bands("mnuWindow").Tools("Arrange").Enabled = False
          End If
        Next iMinWinCount
        
        'Cascade Bit
        For iMinWinCount = 0 To iFormCount - 1
          If (Forms(iMinWinCount).WindowState = vbNormal) And Forms(iMinWinCount).Visible Then
            iNormWinCount = iNormWinCount + 1
          End If
        Next iMinWinCount
        
        If iNormWinCount > 1 Then
          abMain.Bands("mnuWindow").Tools("Cascade").Enabled = True
        Else
          abMain.Bands("mnuWindow").Tools("Cascade").Enabled = False
        End If
      Else
        'Arrange Icons Bit
        abMain.Bands("mnuWindow").Tools("Arrange").Enabled = False
      
        'Cascade Bit
        abMain.Bands("mnuWindow").Tools("Cascade").Enabled = False
      End If
            
      iVisibleWinCount = 0
      For iMinWinCount = 0 To iFormCount - 1
        If Forms(iMinWinCount).Visible Then
          iVisibleWinCount = iVisibleWinCount + 1
        End If
      Next iMinWinCount
      
      'Window Menu - Minimise & Restore Options
      If iVisibleWinCount > 1 And Not frmMain.ActiveForm Is Nothing Then
        abMain.Bands("mnuWindow").Tools("Minimise").Enabled = (frmMain.ActiveForm.WindowState = vbNormal)
        abMain.Bands("mnuWindow").Tools("Restore").Enabled = (frmMain.ActiveForm.WindowState = vbMinimized)
      Else
        abMain.Bands("mnuWindow").Tools("Minimise").Enabled = (iVisibleWinCount > 1)
        abMain.Bands("mnuWindow").Tools("Restore").Enabled = (iVisibleWinCount > 1)
      End If
      
      ' Window Menu - Close Window Option
      abMain.Bands("mnuWindow").Tools("CloseWindow").Enabled = (iVisibleWinCount > 1)
      
      ' RH 31/07/00 - Really wierd bug. This seems to fix it.
      abMain.RecalcLayout
      abMain.Refresh
      
      ' Refresh the Edit menu options.
      RefreshEditMenu
        
      ' Refresh the Record menu options.
      RefreshRecordMenu pfrmCallingForm, pfUnLoad
      
      ' Refresh the History menu options.
      RefreshHistoryMenu pfrmCallingForm, pfUnLoad
    
      'MH20030206 Don't think that we need to refresh the reports menu every time?
      '' Refresh the Reports menu options.
      'RefreshReportsMenu pfrmCallingForm, pfUnLoad
      
      ' Refresh the status bar.
      UpdateStatusBar pfrmCallingForm, pfUnLoad
    End If
  
  End If
  
  With abMain
    .RecalcLayout
    .ResetHooks
    .Refresh
  End With

  DoEvents

End Sub
Public Sub RefreshHistoryMenu(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)
  ' Enable/disable the history menu with the appropriate values.
  Dim fHistoryEnabled As Boolean
  Dim objHistoryScreens As clsHistoryScreens

  fHistoryEnabled = Not pfUnLoad
  
  ' Histories only available for top-level or child table screens.
  If fHistoryEnabled Then
    fHistoryEnabled = ((pfrmCallingForm.ScreenType = screenParentTable) Or _
      (pfrmCallingForm.ScreenType = screenParentView) Or _
      (pfrmCallingForm.ScreenType = screenHistoryTable))
  End If
  
  If pfUnLoad Then
    TryUnload pfrmCallingForm
  Else
  
    If pfrmCallingForm.Recordset.State = adStateClosed Then
      TryUnload pfrmCallingForm
    Else
      ' Histories not available for empty recordsets.
      If fHistoryEnabled Then
        fHistoryEnabled = Not (pfrmCallingForm.Recordset.BOF And pfrmCallingForm.Recordset.EOF)
      End If
    
      ' Histories not available when adding new records.
      If fHistoryEnabled Then
        fHistoryEnabled = (pfrmCallingForm.Recordset.EditMode <> adEditAdd)
      End If
  
      ' Histories only enabled if there are history screens for the current record edit screen.
      If fHistoryEnabled Then
        Set objHistoryScreens = GetHistoryScreens(pfrmCallingForm.ScreenID)
        fHistoryEnabled = (objHistoryScreens.Count > 0)
        Set objHistoryScreens = Nothing
      End If
    End If
  End If

  'MH20010517 Fault 2262 Don't enable the history menu too quick
  '           because if the user clicks it, then its a run-time error
  'abMain.Tools("mnuHistory").Enabled = fHistoryEnabled
  abMain.Tools("mnuHistory").Enabled = (fHistoryEnabled And pfrmCallingForm.Visible)
  
End Sub

'Public Sub RefreshReportsMenu(pfrmCallingForm As Form, Optional ByVal pfUnLoad As Boolean)
'
'  ' Enable/disable the items on the Standard Reports menu.
'
'  ' Note that if StdReports is disabled (ie, no RUN permission) then this
'  ' should overide any other conditions on en/disabling the individual
'  ' reports.
'
'  'Dim fStdreportsEnabled As Boolean
'
'  ' Do we have permission to run Standard Reports at all.
'  'fStdreportsEnabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN")
'
'  'TM20011123 Fault 3066 - Only show the reports if the respective modules are enabled.
'  abMain.Tools("AbsenceBreakdown").Visible = gfAbsenceEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB")
'  abMain.Tools("BradfordIndex").Visible = gfAbsenceEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF")
'  abMain.Tools("StabilityIndex").Visible = gfPersonnelEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_SI")
'  abMain.Tools("Turnover").Visible = gfPersonnelEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_TR")
'
'  abMain.Tools("AbsenceBreakdown").Enabled = gfAbsenceEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB")
'  abMain.Tools("BradfordIndex").Enabled = gfAbsenceEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF")
'  abMain.Tools("StabilityIndex").Enabled = gfPersonnelEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_SI")
'  abMain.Tools("Turnover").Enabled = gfPersonnelEnabled And datGeneral.SystemPermission("STANDARDREPORTS", "RUN_TR")
'
'End Sub


Public Sub PopulateHistoryMenu()
  ' Enable/disable the history menu with the appropriate values.
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim lngPos As Long
  Dim lngScreenID As Long
  Dim lngLastScreenID As Long
  Dim sBand As String
  Dim asSubMenus() As String
  Dim objFileTool As ActiveBarLibraryCtl.Tool
  Dim objBandTool As ActiveBarLibraryCtl.Tool
  Dim objHistoryScreens As clsHistoryScreens
  Dim avTablesDone() As Variant
  Dim fFound As Boolean
  Dim fTableDone As Boolean
  Dim fMultiScreen As Boolean

  ' Column 1 = table ID
  ' Column 2 = screen count
  ' Column 3 = sub menu added ?
  ReDim avTablesDone(3, 0)

  Set objHistoryScreens = GetHistoryScreens(Me.ActiveForm.ScreenID)
  
  With abMain
    ' Remove any existing history sub-menus.
    ReDim asSubMenus(0)
    For iLoop = 0 To (.Bands.Count - 1)
      If Left(.Bands(iLoop).Name, 11) = "bndSUBMENU_" Then
        iNextIndex = UBound(asSubMenus) + 1
        ReDim Preserve asSubMenus(iNextIndex)
        asSubMenus(iNextIndex) = .Bands(iLoop).Name
      End If
    Next iLoop
    For iLoop = 1 To UBound(asSubMenus)
      .Bands.Remove asSubMenus(iLoop)
    Next iLoop
    ' Clear the history menu.
    .Bands("mnuHistory").Tools.RemoveAll
    
    For iLoop = 1 To objHistoryScreens.Count
      fFound = False
      For iLoop2 = 1 To UBound(avTablesDone, 2)
        If avTablesDone(1, iLoop2) = objHistoryScreens.Item(iLoop).TableID Then
          fFound = True
          avTablesDone(2, iLoop2) = avTablesDone(2, iLoop2) + 1
        End If
      Next iLoop2
      
      If Not fFound Then
        ReDim Preserve avTablesDone(3, UBound(avTablesDone, 2) + 1)
        avTablesDone(1, UBound(avTablesDone, 2)) = objHistoryScreens.Item(iLoop).TableID
        avTablesDone(2, UBound(avTablesDone, 2)) = 1
        avTablesDone(3, UBound(avTablesDone, 2)) = False
      End If
    Next iLoop

    ' Add menu items for each history screen in the collection.
    For iLoop = 1 To objHistoryScreens.Count
      
      ' Create the history screen menu item (without placing it in the menu).
      If objHistoryScreens.Item(iLoop).ViewID > 0 Then
        Set objFileTool = .Tools.Add(.Tools.Count + 1, "HV" & objHistoryScreens.Item(iLoop).HistoryScreenID & ":" & objHistoryScreens.Item(iLoop).ViewID)
        objFileTool.Caption = Replace(objHistoryScreens.Item(iLoop).HistoryScreenName, "&", "&&") & _
          " (" & RemoveUnderScores(objHistoryScreens.Item(iLoop).ViewName) & " view)..."
      Else
        Set objFileTool = .Tools.Add(.Tools.Count + 1, "HT" & objHistoryScreens.Item(iLoop).HistoryScreenID)
        objFileTool.Caption = Replace(objHistoryScreens.Item(iLoop).HistoryScreenName, "&", "&&") & "..."
      End If
      objFileTool.Style = DDSStandard
      If objHistoryScreens.Item(iLoop).PictureID > 0 Then
        LoadMenuPicture objHistoryScreens.Item(iLoop).PictureID, objFileTool
      Else
        objFileTool.SetPicture 0, LoadResPicture("CHILDTABLE", 0), COL_GREY
      End If
          
      fTableDone = False
      fMultiScreen = False
      For iLoop2 = 1 To UBound(avTablesDone, 2)
        If avTablesDone(1, iLoop2) = objHistoryScreens.Item(iLoop).TableID Then
          fTableDone = avTablesDone(3, iLoop2)
          fMultiScreen = (avTablesDone(2, iLoop2) > 1)
          avTablesDone(3, iLoop2) = True
          Exit For
        End If
      Next iLoop2

      If fTableDone Then
        ' The current screen is for the same table as the last screen added to the menu
        ' which will have created the sub-menu, so just add it to the sub-menu.
        sBand = "bndSUBMENU_" & objHistoryScreens.Item(iLoop).TableName
        iIndex = -1
        For iLoop2 = 0 To (.Bands(sBand).Tools.Count - 1)
          If LCase(.Bands(sBand).Tools(iLoop2).Caption) > LCase(objFileTool.Caption) Then
            iIndex = iLoop2
            Exit For
          End If
        Next iLoop2
        .Bands(sBand).Tools.Insert iIndex, objFileTool
      Else
        If fMultiScreen Then
          ' The current screen is for the same table as the next screen to be added
          ' but is for a different table to the last screen added to the menu
          ' so create a sub-menu, and add this screen to the sub-menu.
          sBand = "bndSUBMENU_" & objHistoryScreens.Item(iLoop).TableName
          .Bands.Add sBand
          .Bands(sBand).Type = DDBTPopup
          .Bands(sBand).Tools.RemoveAll
          
          Set objBandTool = .Tools.Add(.Tools.Count + 1, sBand)
          If objHistoryScreens.Item(iLoop).ViewID > 0 Then
            objBandTool.Caption = objHistoryScreens.Item(iLoop).TableName & _
              " (" & objHistoryScreens.Item(iLoop).ViewName & " view)"
            objBandTool.SetPicture 0, LoadResPicture("VIEW", 0), COL_GREY
          Else
            objBandTool.Caption = objHistoryScreens.Item(iLoop).TableName
            objBandTool.SetPicture 0, LoadResPicture("SCREEN", 0), COL_GREY
          End If
          objBandTool.SubBand = sBand
          
          iIndex = -1
          For iLoop2 = 0 To (.Bands("mnuHistory").Tools.Count - 1)
            If LCase(.Bands("mnuHistory").Tools(iLoop2).Caption) > LCase(objBandTool.Caption) Then
              iIndex = iLoop2
              Exit For
            End If
          Next iLoop2
          .Bands("mnuHistory").Tools.Insert iIndex, objBandTool
          .Bands(sBand).Tools.Insert 0, objFileTool
        Else
          ' The current screen is for a different table/view to the next and last screens
          ' added to the menu so just add this screen to the main menu as normal.
          iIndex = -1
          For iLoop2 = 0 To (.Bands("mnuHistory").Tools.Count - 1)
            'TM20011220 Fault 2670 - need to compare lowercase to lowercase to sort menu items.
            If LCase(.Bands("mnuHistory").Tools(iLoop2).Caption) > LCase(objFileTool.Caption) Then
              iIndex = iLoop2
              Exit For
            End If
          Next iLoop2
          .Bands("mnuHistory").Tools.Insert iIndex, objFileTool
        End If
      End If
    Next iLoop

    ' Position 'beginGroup' lines in the sub-menus.
    For iLoop = 0 To (.Bands("mnuHistory").Tools.Count - 1)
      If Len(.Bands("mnuHistory").Tools(iLoop).SubBand) > 0 Then
        lngLastScreenID = 0
        
        For iLoop2 = 0 To (.Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools.Count - 1)
          If Left(.Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).Name, 2) = "HT" Then
            lngScreenID = Val(Right(.Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).Name, _
              Len(.Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).Name) - 2))
          Else
            lngPos = InStr(1, .Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).Name, ":")
            lngScreenID = Val(Mid$(.Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).Name, 3, lngPos - 1))
          End If
    
          If lngLastScreenID <> lngScreenID Then
            .Bands(.Bands("mnuHistory").Tools(iLoop).SubBand).Tools(iLoop2).BeginGroup = True
          End If
          
          lngLastScreenID = lngScreenID
        Next iLoop2
      End If
    Next iLoop
  
  End With

End Sub


Private Sub MDIForm_Resize()
  'JPD 20030908 Fault 5756
  If Me.WindowState <> vbMinimized Then
    giWindowState = Me.WindowState
    
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
  Dim rsMessages As ADODB.Recordset
    
  If Application.LoggedIn Then
  
    sMessage = ""

    sSQL = "exec dbo.sp_ASRGetMessages"
    Set rsMessages = gobjDataAccess.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsMessages
      Do While Not .EOF
        If Len(sMessage) > 0 Then
          sMessage = sMessage & vbNewLine & vbNewLine & vbNewLine
        End If
        
        sMessage = sMessage & rsMessages.Fields(0).Value
        
        rsMessages.MoveNext
      Loop
    
      .Close
    End With
  
    Set rsMessages = Nothing
  
    If Len(sMessage) > 0 Then
      COAMsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
    End If
  End If

End Sub

Private Sub tmrDiary_Timer()
'
'  'This time will be disabled if the variable
'  'gblnDiaryConstCheck is set to false
'
'  On Local Error GoTo LocalErr
'
'  Static lngNextAlarmTime As Long
'  Static lngLastDiaryCheck As Long
'  Dim lngCurrentTime As Long
'
'  'MH20021120 Fault 4794
'  gobjErrorStack.Disable
'
'  'Get current time (to the nearest whole minute)
'  lngCurrentTime = (Timer \ 60) * 60
'
'  'Check for next event
'  If (lngCurrentTime - lngLastDiaryCheck) >= (glngDiaryIntervalCheck * 60) Or _
'    tmrDiary.Interval = 1 Then
'        lngLastDiaryCheck = lngCurrentTime
'        Call gobjDiary.GetNextAlarmTime(lngNextAlarmTime)
'  End If
'
'  If lngCurrentTime >= lngNextAlarmTime And lngNextAlarmTime > -1 Then
'    '2=Current and Future,0=DayView
'    gobjDiary.ShowAlarmedEvents 2, 0
'    lngLastDiaryCheck = 0   'When next in, check for next alarm
'  End If
'
'  If tmrDiary.Interval < 60000 Then
'    tmrDiary.Interval = (60 - (Int(Timer) Mod 60)) * 1000
'  End If
'
'  'MH20021120 Fault 4794
'  gobjErrorStack.Enable
'
'Exit Sub
'
'LocalErr:
'  On Local Error Resume Next
'  tmrDiary.Enabled = False
'  COAMsgBox "An error occurred checking for alarmed diary events.  No further checks will be made for alarmed events until HR Pro is restarted." & _
'         IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", ""), vbExclamation, "Alarmed Diary Events"
'
  gobjDiary.CheckAlarmedEvents tmrDiary, mstrLastAlarmCheck

End Sub

Public Sub CalendarReportsClick()

  Dim pobjCalendarReports As clsCalendarReportsRUN
  
  'Dim sSQL As String
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmCalendarReport
  
  Screen.MousePointer = vbHourglass
  
  fExit = False
  Set frmSelection = New frmDefSel
    
  With frmSelection
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
      .EnableRun = True
      
      If .ShowList(utlCalendarReport) Then
        
        .CustomShow vbModal
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmCalendarReport
            frmEdit.Initialise True, .FromCopy, , False
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
            Unload frmEdit
            Set frmEdit = Nothing


          Case edtEdit
            Set frmEdit = New frmCalendarReport
            frmEdit.Initialise False, .FromCopy, .SelectedID, False
            If Not frmEdit.Cancelled Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
       
          Case edtSelect
  '        'TO RUN AS NORMAL
            Set pobjCalendarReports = New clsCalendarReportsRUN
            pobjCalendarReports.CalendarReportID = .SelectedID
            pobjCalendarReports.RunCalendarReport ""
            Set pobjCalendarReports = Nothing
            fExit = gbCloseDefSelAfterRun
  
          Case edtPrint
            Set frmEdit = New frmCalendarReport
            frmEdit.Initialise False, False, .SelectedID, True
            If Not frmEdit.Cancelled Then
              frmEdit.PrintDef .SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
  
          Case edtCancel
            fExit = True
              
          End Select
        End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Public Sub RecordProfileClick()

  Dim pobjRecordProfiles As clsRecordProfileRUN
  
  Dim lForms As Long
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmRecordProfile
  
  Screen.MousePointer = vbHourglass
  
  fExit = False
  Set frmSelection = New frmDefSel
    
  With frmSelection
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
      .EnableRun = True
      
      If .ShowList(utlRecordProfile) Then
        
        .Show vbModal
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmRecordProfile
            frmEdit.Initialise True, .FromCopy
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
          
          'TM20010808 Fault 2656 - Must validate, check ownership etc... before allowing the edit/copy.
          Case edtEdit
            Set frmEdit = New frmRecordProfile
            frmEdit.Initialise False, .FromCopy, .SelectedID
            If Not frmEdit.Cancelled Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
        
          Case edtSelect
            Set pobjRecordProfiles = New clsRecordProfileRUN
            pobjRecordProfiles.RecordProfileID = .SelectedID
            pobjRecordProfiles.RunRecordProfile
            Set pobjRecordProfiles = Nothing
            fExit = gbCloseDefSelAfterRun
  
          'TM20010808 Fault 2656 - Must validate, check ownership etc... before allowing the print.
          Case edtPrint
            Set frmEdit = New frmRecordProfile
            frmEdit.Initialise False, False, .SelectedID, True
            If Not frmEdit.Cancelled Then
              frmEdit.PrintDef .SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
  
        Case 0
          fExit = True
              
          End Select
        
        End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

End Sub


Public Sub CustomReportsClick()

  Dim pobjCustomReports As clsCustomReportsRUN
  
  'Dim sSQL As String
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmCustomReports
  
  Screen.MousePointer = vbHourglass
  
  fExit = False
  Set frmSelection = New frmDefSel
    
  With frmSelection
    ' Loop until the picklist operation has been cancelled.
    Do While Not fExit
      .EnableRun = True
      
      If .ShowList(utlCustomReport) Then
        
        .CustomShow vbModal
        Select Case .Action
          Case edtAdd
            Set frmEdit = New frmCustomReports
            frmEdit.Initialise True, .FromCopy
            frmEdit.Show vbModal
            .SelectedID = frmEdit.SelectedID
            Unload frmEdit
            Set frmEdit = Nothing
          
          'TM20010808 Fault 2656 - Must validate, check ownership etc... before allowing the edit/copy.
          Case edtEdit
            Set frmEdit = New frmCustomReports
            frmEdit.Initialise False, .FromCopy, .SelectedID
            If Not frmEdit.Cancelled Then
              frmEdit.Show vbModal
              If .FromCopy And frmEdit.SelectedID > 0 Then
                .SelectedID = frmEdit.SelectedID
              End If
            End If
            Unload frmEdit
            Set frmEdit = Nothing
        
'        'TM20010808 Fault 2656 - Must validate, check ownership etc... before allowing the delete.
'        Case edtDelete
'            Set frmEdit = New frmCustomReports
'            frmEdit.Initialise False, .FromCopy, .SelectedID
'            If Not frmEdit.Cancelled Then
'              datGeneral.DeleteRecord "ASRSysCustomReportsName", "ID", .SelectedID
'              datGeneral.DeleteRecord "ASRSysCustomReportsDetails", "CustomReportID", .SelectedID
'            End If
'            Unload frmEdit
'            Set frmEdit = Nothing
        
          Case edtSelect
  '        'TO RUN AS NORMAL
            Set pobjCustomReports = New clsCustomReportsRUN
            pobjCustomReports.CustomReportID = .SelectedID
            pobjCustomReports.RunCustomReport ("")
            Set pobjCustomReports = Nothing
            fExit = gbCloseDefSelAfterRun
  
          'TM20010808 Fault 2656 - Must validate, check ownership etc... before allowing the print.
          Case edtPrint
            Set frmEdit = New frmCustomReports
            frmEdit.Initialise False, False, .SelectedID, True
            If Not frmEdit.Cancelled Then
              frmEdit.PrintDef .SelectedID
            End If
            Unload frmEdit
            Set frmEdit = Nothing
  
          Case edtCancel
            fExit = True
              
          End Select
        End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

End Sub



Public Sub MatchReportClick(mrtMatchReportType As MatchReportType)
  
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmMatchDef
  Dim frmRun As frmMatchRun
  Dim lngTYPE As UtilityType


  If mrtMatchReportType <> mrtNormal Then
    If Not ValidatePostParameters Then
      Exit Sub
    End If
  End If


  Screen.MousePointer = vbHourglass

  fExit = False
  Set frmSelection = New frmDefSel

  With frmSelection
    
    Do While Not fExit
      .EnableRun = True

      Select Case mrtMatchReportType
      Case mrtNormal: lngTYPE = utlMatchReport
      Case mrtSucession: lngTYPE = utlSuccession
      Case mrtCareer: lngTYPE = utlCareer
      End Select

      If .ShowList(lngTYPE, "MatchReportType = " & CStr(mrtMatchReportType)) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtAdd
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise True, .FromCopy
          frmEdit.Show vbModal
          .SelectedID = frmEdit.SelectedID
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtEdit
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise False, .FromCopy, .SelectedID
          If Not frmEdit.Cancelled Then
            frmEdit.Show vbModal
            If .FromCopy And frmEdit.SelectedID > 0 Then
              .SelectedID = frmEdit.SelectedID
            End If
          End If
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtSelect
          Set frmRun = New frmMatchRun
          frmRun.MatchReportType = mrtMatchReportType
          frmRun.MatchReportID = .SelectedID
          frmRun.RunMatchReport
          If frmRun.PreviewOnScreen Then
            frmRun.Show vbModal
          End If
          Set frmRun = Nothing
          fExit = gbCloseDefSelAfterRun

        Case edtPrint
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise False, False, .SelectedID, True
          If Not frmEdit.Cancelled Then
            frmEdit.PrintDef .SelectedID
          End If
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtCancel
          fExit = True

          End Select
        End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

End Sub


'Public Sub CareerSuccessionClick(lngMatchReportType As MatchReportType)
'
'  Dim frmRun As frmMatchRun
'
'  Set frmRun = New frmMatchRun
'
'  If lngMatchReportType = mrtSucession Then
'    frmRun.SetPostReport mrtSucession, glngSuccessionDef, gblnSuccessionAllowEqual, gblnSuccessionRestrict, gblnSuccessionLevels
'  Else
'    frmRun.SetPostReport mrtCareer, glngCareerDef, gblnCareerAllowEqual, gblnCareerRestrict, gblnCareerLevels
'  End If
'
'  frmRun.RunMatchReport False
'  If frmRun.PreviewOnScreen Then
'    frmRun.Show vbModal
'  End If
'  Set frmRun = Nothing
'
'End Sub

Public Sub CrossTabClick()

  Dim frmDefinition As frmCrossTabDef
  Dim frmExecution As frmCrossTabRun
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  Screen.MousePointer = vbHourglass

  Set frmSelection = New frmDefSel
  blnExit = False
  
  With frmSelection
    Do While Not blnExit
      
      .EnableRun = True
      
      If .ShowList(utlCrossTab) Then
        
        .CustomShow vbModal
        'TM09012004 Fault 4150
        'DoEvents
        
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmCrossTabDef
          If frmDefinition.Initialise(True, .FromCopy) Then
            frmDefinition.Show vbModal
            .SelectedID = frmDefinition.SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmCrossTabDef
          If frmDefinition.Initialise(False, .FromCopy, .SelectedID) Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtSelect
          Set frmExecution = New frmCrossTabRun
          blnOK = frmExecution.ExecuteCrossTab(.SelectedID)
          If frmExecution.PreviewOnScreen Then
            frmExecution.Show vbModal
          End If
          Unload frmExecution
          Set frmExecution = Nothing
          blnExit = gbCloseDefSelAfterRun

        Case edtPrint
          Set frmDefinition = New frmCrossTabDef
          frmDefinition.PrintDef .SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case edtCancel    'Cancel
          blnExit = True

        End Select
      
      End If

    Loop
  
  End With

  Unload frmSelection
  Set frmSelection = Nothing
  
End Sub


Public Sub MailMergeClick()

  Dim frmDefinition As frmMailMerge
  Dim objExecution As clsMailMergeRun
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False
  
  'sSQL = "Select Name, MailMergeID From ASRSysMailMergeName " & _
         "WHERE Username = '" & gsUserName & "' OR Access <> 'HD'"

  Set frmDefinition = New frmMailMerge
  
  With frmSelection
    Do While Not blnExit
      
      .EnableRun = True
      
      If .ShowList(utlMailMerge) Then
       
        .CustomShow vbModal
        
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = False
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = False
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
           
'        'TM20010808 Fault 2656 - Must validate the definition before allowing the delete.
'        Case edtDelete
'          Set frmDefinition = New frmMailMerge
'          frmDefinition.Initialise False, .FromCopy, .SelectedID
'          If Not frmDefinition.Cancelled Then
'            datGeneral.DeleteRecord "ASRSysMailMergeName", "MailMergeID", .SelectedID
'            datGeneral.DeleteRecord "ASRSysMailMergeColumns", "MailMergeID", .SelectedID
'          End If
'          Unload frmDefinition
'          Set frmDefinition = Nothing

        
        Case edtSelect
          Set objExecution = New clsMailMergeRun
          objExecution.ExecuteMailMerge .SelectedID
          Set objExecution = Nothing
          blnExit = gbCloseDefSelAfterRun
          
        'TM20010808 Fault 2656 - Must validate the definition before allowing the print.
        Case edtPrint
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = False
          frmDefinition.Initialise False, False, .SelectedID, True
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case edtCancel
          blnExit = True  'cancel

        End Select
      
      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Public Sub CrystalReportsClick()

End Sub

Public Sub CalculationsClick()

  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(0, 0, giEXPR_RUNTIMECALCULATION, 0) Then
      .SelectExpression False, lngOptions
    End If
  End With

  Set objExpr = Nothing
  
End Sub

'MH20000804 MH Changed to Public so that I can call it from RecEdit4
'Private Sub RefreshRecordEditScreens()
Public Sub RefreshRecordEditScreens()
  ' Refresh the record editing screens.
  ' This is done after utilities that may have updated the data are run.
  Dim frmForm As Form
  
'  Set frmform2 = Me.ActiveForm
  
  ' Loop through the MDI form's child forms.
  For Each frmForm In Forms
    ' Only refresh record editing screens.
    If TypeOf frmForm Is frmRecEdit4 Then
      ' Only refresh the top-level screens, as the refrehment will
      ' cascade down through these to any child screens.
      If (frmForm.ScreenType = screenParentTable) Or _
        (frmForm.ScreenType = screenParentView) Or _
        (frmForm.ScreenType = screenLookup) Then
        frmForm.Requery False
      End If
    End If
  Next frmForm
  Set frmForm = Nothing
  
'  frmMain.RefreshMainForm Me.ActiveForm
  
End Sub


Private Function SaveCurrentRecordEditScreen() As Boolean
  ' Save changes in the current record editing screen.
  
  SaveCurrentRecordEditScreen = True
  
  ' Only refresh record editing screens.
  If Not Me.ActiveForm Is Nothing Then
    If TypeOf Me.ActiveForm Is frmRecEdit4 Then
      SaveCurrentRecordEditScreen = Me.ActiveForm.SaveChanges
    End If
  End If
  
End Function


Private Function GetRecCount(strSQL As String) As Long

  Dim rsTemp As ADODB.Recordset

  GetRecCount = 0

  Set rsTemp = New ADODB.Recordset
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    GetRecCount = Val(rsTemp(0).Value)
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

'Private Function GetMinimumPasswordLength() As Long
'
'  Dim rsTemp As Recordset
'
'  Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT MinimumPasswordLength FROM ASRSysConfig")
'
'  If rsTemp.BOF And rsTemp.EOF Then
'    GetMinimumPasswordLength = 0
'  Else
'    GetMinimumPasswordLength = rsTemp.Fields(0)
'  End If
'
'  Set rsTemp = Nothing
'
'End Function


Public Sub BradfordIndexClick()
  
  Dim pobjBradfordIndex As clsCustomReportsRUN
  
  Set pobjBradfordIndex = New clsCustomReportsRUN
  pobjBradfordIndex.RunBradfordReport ""
  Set pobjBradfordIndex = Nothing

End Sub


Public Sub LabelsAndEnvelopesClick()

  Dim frmDefinition As frmMailMerge
  Dim objExecution As clsMailMergeRun
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False
  
  'sSQL = "Select Name, MailMergeID From ASRSysMailMergeName " & _
         "WHERE Username = '" & gsUserName & "' OR Access <> 'HD'"

  Set frmDefinition = New frmMailMerge
  
  With frmSelection
    Do While Not blnExit
      
      .EnableRun = True
      
      If .ShowList(utlLabel) Then
        
        .CustomShow vbModal
        
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = True
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = True
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
           
        Case edtSelect
          Set objExecution = New clsMailMergeRun
          objExecution.ExecuteMailMerge .SelectedID
          Set objExecution = Nothing
          blnExit = gbCloseDefSelAfterRun

        Case edtPrint
          Set frmDefinition = New frmMailMerge
          frmDefinition.IsLabel = True
          frmDefinition.Initialise False, False, .SelectedID, True
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case edtCancel
          blnExit = True  'cancel

        End Select
      
      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub


Public Sub EmailGroupClick()
  
  Dim frmDefinition As frmEmailDefGroup
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False

  Set frmDefinition = New frmEmailDefGroup
  
  With frmSelection
    Do While Not blnExit
      
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
      .EnableRun = False
      .TableComboVisible = False

      If .ShowList(utlEmailGroup) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing

        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtPrint
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case 0
          blnExit = True  'cancel

        End Select

      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub


Private Sub EnableTools()

  With abMain
    .Tools("mnuRecord").Visible = False
    .Tools("mnuHistory").Visible = False

    .Tools("CustomReports").Enabled = MenuEnabled("CUSTOMREPORTS")
    .Tools("CrossTab").Enabled = MenuEnabled("CROSSTABS")
    .Tools("CalendarReport").Enabled = MenuEnabled("CALENDARREPORTS")
    .Tools("RecordProfile").Enabled = MenuEnabled("RECORDPROFILE")
    '.Tools("CrystalReports").Visible = False

    .Tools("MatchReport").Enabled = MenuEnabled("MATCHREPORTS")
    .Tools("Succession").Visible = gfPersonnelEnabled
    .Tools("Succession").Enabled = MenuEnabled("SUCCESSION")
    .Tools("Career").Visible = gfPersonnelEnabled
    .Tools("Career").Enabled = MenuEnabled("CAREER")


    .Tools("AbsenceBreakdown").Visible = gfAbsenceEnabled
    .Tools("AbsenceBreakdown").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB")
    .Tools("BradfordIndex").Visible = gfAbsenceEnabled
    .Tools("BradfordIndex").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF")
    .Tools("StabilityIndex").Visible = gfPersonnelEnabled
    .Tools("StabilityIndex").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_SI")
    .Tools("Turnover").Visible = gfPersonnelEnabled
    .Tools("Turnover").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_TR")
    
    With .Bands("bndReportConfig")
      .Tools("AbsenceBreakdownCfg").Visible = gfAbsenceEnabled
      .Tools("AbsenceBreakdownCfg").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB")
      .Tools("BradfordIndexCfg").Visible = gfAbsenceEnabled
      .Tools("BradfordIndexCfg").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF")
      .Tools("StabilityIndexCfg").Visible = gfPersonnelEnabled
      .Tools("StabilityIndexCfg").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_SI")
      .Tools("TurnoverCfg").Visible = gfPersonnelEnabled
      .Tools("TurnoverCfg").Enabled = datGeneral.SystemPermission("STANDARDREPORTS", "RUN_TR")
    End With
  
    .Bands("mnuAdministration").Tools("ReportConfiguration").Visible = (gfAbsenceEnabled Or gfPersonnelEnabled)
    .Bands("mnuAdministration").Tools("ID_PollMode").Visible = gbActivateJobServer

    '.Tools("Succession").Visible = gfPersonnelEnabled
    '.Tools("Succession").Enabled = (glngSuccessionDef > 0)
    '.Tools("Career").Visible = gfPersonnelEnabled
    '.Tools("Career").Enabled = (glngCareerDef > 0)

    .Tools("Diary").Enabled = datGeneral.SystemPermission("DIARY", "MANUALEVENTS")
    .Tools("BatchJobs").Enabled = MenuEnabled("BATCHJOBS")
    .Tools("MailMerge").Enabled = MenuEnabled("MAILMERGE")
    .Tools("mnuLabels").Enabled = MenuEnabled("LABELS")
    
    ' Workflow stuff
    .Tools("Workflow").Visible = gbWorkflowEnabled
    .Tools("Workflow").Enabled = gbWorkflowEnabled And MenuEnabled("WORKFLOW")
    .Tools("WorkflowLog").Visible = gbWorkflowEnabled
    .Tools("WorkflowLog").Enabled = gbWorkflowEnabled And _
      (datGeneral.SystemPermission("WORKFLOW", "ADMINISTER") Or _
        datGeneral.SystemPermission("WORKFLOW", "VIEWLOG"))
    .Tools("WorkflowOutOfOffice").Visible = gbWorkflowOutOfOfficeEnabled
    .Tools("WorkflowOutOfOffice").Enabled = gbWorkflowOutOfOfficeEnabled

    .Tools("GlobalAdd").Enabled = MenuEnabled("GLOBALADD")
    .Tools("GlobalUpdate").Enabled = MenuEnabled("GLOBALUPDATE")
    .Tools("GlobalDelete").Enabled = MenuEnabled("GLOBALDELETE")

    .Tools("DataTransfer").Enabled = MenuEnabled("DATATRANSFER")
    .Tools("Import").Enabled = MenuEnabled("IMPORT")
    .Tools("Export").Enabled = MenuEnabled("EXPORT")

    .Tools("Calculations").Enabled = MenuEnabled("CALCULATIONS")
    .Tools("PickLists").Enabled = MenuEnabled("PICKLISTS")
    .Tools("Filters").Enabled = MenuEnabled("FILTERS")
    .Tools("EmailGroups").Enabled = MenuEnabled("EMAILGROUPS")
    .Tools("ID_LabelTemplates").Enabled = MenuEnabled("LABELDEFINITION")
    .Tools("ID_DocumentTypes").Visible = gbVersion1Enabled
    .Tools("ID_DocumentTypes").Enabled = MenuEnabled("VERSION1")

    'JPD 20030912 Fault 6961 & Fault 6962
    '.Tools("EventLog").Enabled = datGeneral.SystemPermission("CONFIGURATION", "USER")
    .Tools("EventLog").Enabled = True
    .Tools("Configuration").Enabled = datGeneral.SystemPermission("CONFIGURATION", "USER")
    .Tools("PC Configuration").Enabled = True
    
    .Tools("WorkflowPopup").Visible = gbWorkflowEnabled
    .Tools("WorkflowPopup").Enabled = gbWorkflowEnabled

    
    'Windows Authentication Stuff
    .Tools("ChangePassword").Enabled = Not gbUseWindowsAuthentication
    
    ' CMG specific stuff
    .Bands("bndHouseKeeping").Tools("ID_CMGCommit").Visible = gbCMGEnabled
    .Bands("bndHouseKeeping").Tools("ID_CMGRecovery").Visible = gbCMGEnabled
    .Bands("bndHouseKeeping").Tools("ID_CMGCommit").Enabled = datGeneral.SystemPermission("CMG", "CMGCOMMIT")
    .Bands("bndHouseKeeping").Tools("ID_CMGRecovery").Enabled = datGeneral.SystemPermission("CMG", "CMGRECOVERY")

    ' Payroll stuff
    .Tools("ID_Accord").Visible = gbAccordEnabled
    .Tools("ID_Accord").Enabled = gbAccordEnabled And datGeneral.SystemPermission("ACCORD", "VIEWTRANSFER")
    .Tools("ID_Accord_Create").Enabled = gbAccordEnabled And datGeneral.SystemPermission("ACCORD", "SENDRECORD")
    .Tools("ID_Accord_Archive").Enabled = gbAccordEnabled And datGeneral.SystemPermission("ACCORD", "VIEWARCHIVE")
    
    .RecalcLayout
  End With

End Sub


Private Function MenuEnabled(strCategory As String) As Boolean
  MenuEnabled = datGeneral.SystemPermission(strCategory, "VIEW") Or _
                datGeneral.SystemPermission(strCategory, "DELETE") Or _
                datGeneral.SystemPermission(strCategory, "RUN") Or _
                gfCurrentUserIsSysSecMgr
End Function


Private Sub StandardReportClick(strToolName As String)

  Dim frmDef As frmConfigurationReports
  Dim bExit As Boolean

  bExit = False
  Set frmDef = New frmConfigurationReports
  
  With frmDef
    .Run = (Right(strToolName, 3) <> "Cfg")

    Select Case strToolName
      Case "AbsenceBreakdown", "AbsenceBreakdownCfg"
        If ValidateAbsenceParameters_BreakdownReport Then
          .ShowControls "Absence Breakdown"
        Else
          bExit = True
        End If
    
      Case "BradfordIndex", "BradfordIndexCfg"
        If ValidateAbsenceParameters_BreakdownReport Then
          .ShowControls "Bradford Factor"
        Else
          bExit = True
        End If
      
      Case "StabilityIndex", "StabilityIndexCfg"
        .ShowControls "Stability"
    
      Case "Turnover", "TurnoverCfg"
        .ShowControls "Turnover"
    End Select

    Do While Not bExit
      .Show vbModal
      bExit = IIf(.Action = rptRun, gbCloseDefSelAfterRun, True)
    Loop
  
  End With

  Unload frmDef
  Set frmDef = Nothing

End Sub

Public Sub LabelTemplatesClick()

  Dim frmDefinition As frmLabelTypeDefinition
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False
   
  With frmSelection
    Do While Not blnExit
      
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
      .EnableRun = False
      .TableComboEnabled = False
      .TableComboVisible = False

      If .ShowList(utlLabelType) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.Initialise True, .FromCopy, , False
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtPrint
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.PrintDefinition .SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case 0
          blnExit = True  'cancel

        End Select
      
      End If
    
    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

  RefreshMainForm Me, False

End Sub


Public Sub TryUnload(frmTemp As Form)

  'Ref sub "CheckForNonactiveForms"
  On Local Error Resume Next
  Unload frmTemp

End Sub

Public Sub CheckForNonactiveForms(pfrmCallingForm As Form)

  Dim frmForm As Form

  'MH20050516 Fault 9978
  'Avoid the error "Unable to unload within this context"
  'if this procedure is called as part of a form resize.
  'Have now added code so that this procedure is called
  'again after the form resize has finished.
  On Local Error Resume Next

  ' Kill off any redundant forms.
  For Each frmForm In Forms
    If (TypeOf frmForm Is frmFind2) Or _
      (TypeOf frmForm Is frmRecEdit4) Then
      If Not frmForm.Recordset Is Nothing Then
        If frmForm.Recordset.State = adStateClosed Then
          If (Not frmForm Is pfrmCallingForm) Then
            Unload frmForm
          End If
        End If
    End If
    End If
  Next frmForm
  Set frmForm = Nothing

End Sub

' Poll the server for any jobs that need executing.
Private Sub RunPollJob()

  On Error GoTo ErrorTrap

  Dim bPreviousCloseDefSelAfterRun As Boolean

  Dim objTool As New ActiveBarLibraryCtl.Tool
  
  Dim sSQL As String
  Dim sJob As String
  Dim rsMessages As ADODB.Recordset
  Dim lngStart As Long
  Dim lngEnd As Long

  bPreviousCloseDefSelAfterRun = gbCloseDefSelAfterRun

  If Application.LoggedIn Then

    gbCloseDefSelAfterRun = True
    gblnBatchMode = True
    gbIsInPollMode = True
    gbJustRunIt = True
  
    With gobjProgress
      .Bar1MaxValue = 1
      .Caption = "Running Jobs..."
      .MainCaption = "Waiting for job..."
      .AVI = dbTable
      .Time = False
      .Cancel = True
      .NumberOfBars = 2
      .OpenProgress
    End With
  
    
    sSQL = "exec dbo.[spASRGetJobs];"
    
    Do While Not gobjProgress.Cancelled
    
      Set rsMessages = gobjDataAccess.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsMessages
        Do While Not .EOF
          sJob = rsMessages.Fields(0).Value
          
          lngStart = InStr(1, sJob, "(")
          lngEnd = InStr(1, sJob, ")")
          If lngStart > 0 And lngEnd > 0 Then
            glngBypassDefsel_ID = Mid(sJob, lngStart + 1, (lngEnd - lngStart) - 1)
            objTool.Name = Mid(sJob, 1, lngStart - 1)
          Else
            objTool.Name = sJob
          End If
          
          abMain_Click objTool
          glngBypassDefsel_ID = 0
          
          gobjProgress.MainCaption = "Waiting for job..."
          
          rsMessages.MoveNext
        Loop
      
        .Close
      End With
    
    Loop
       
    
  End If
    
TidyUpAndExit:
    gobjProgress.CloseProgress
    Set rsMessages = Nothing
    
    gbCloseDefSelAfterRun = bPreviousCloseDefSelAfterRun
    gblnBatchMode = False
    gbIsInPollMode = False
    gbJustRunIt = False
    
    Exit Sub

ErrorTrap:
  GoTo TidyUpAndExit
    
End Sub

Private Sub DocumentTypesClick()

  Dim frmDefinition As frmDocumentMap
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean
  Dim blnOK As Boolean
  Dim strSelectedName As String
  Dim lngSelectedID As Long

  Set frmSelection = New frmDefSel
  blnExit = False
   
  With frmSelection
    Do While Not blnExit
      
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties
      .EnableRun = False
      .TableComboEnabled = False
      .TableComboVisible = False
           
      If .ShowList(utlDocumentMapping) Then
        
        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmDocumentMap
          frmDefinition.Initialise True, .FromCopy, , False
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmDocumentMap
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
           
        Case edtPrint
          Set frmDefinition = New frmDocumentMap
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          frmDefinition.PrintDefinition .SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case 0
          blnExit = True  'cancel

        End Select
           
      End If
    
    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

  RefreshMainForm Me, False
  
End Sub
