VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form frmChangedPlatform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Platform Change Details"
   ClientHeight    =   4995
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5079
   Icon            =   "frmChangedPlatform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraUsage 
      Caption         =   "Details :"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1890
      Begin ComctlLib.ListView lstUsage 
         Height          =   525
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   926
         View            =   3
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
   End
   Begin VB.PictureBox picFormIcon 
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   195
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraButtons 
      Caption         =   "Frame1"
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1200
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   500
         Width           =   1200
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy"
         Height          =   400
         Left            =   0
         TabIndex        =   8
         Top             =   2000
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         Height          =   400
         Left            =   0
         TabIndex        =   2
         Top             =   1000
         Width           =   1200
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   400
         Left            =   0
         TabIndex        =   1
         Top             =   1500
         Width           =   1200
      End
   End
   Begin VB.Label lblUsageMSG 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   2085
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmChangedPlatform.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmChangedPlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_START_WIDTH = 6600
Private Const FORM_START_HEIGHT = 4500

Private Const OFFSET_LEFT = 160
Private Const OFFSET_RIGHT = 160
Private Const OFFSET_TOP = 300
Private Const OFFSET_BOTTOM = 160
Private Const MSG_OFFSET = 240
Private Const BUTTON_OFFSET = 95
Private Const MIN_FORM_HEIGHT = 3375
Private Const MIN_FORM_WIDTH = 6435

Private miChoice As Integer
Private miButtons As Integer

Private mlngMaxTextLength As Long
Private mlngMaxOldValLen As Long
Private mlngMaxNewValLen As Long

Private miMode As ScreenMode

Private Enum ScreenMode
  miMODE_CHANGEDPLATFORM = 0
  miMODE_WORKFLOWURLS = 1
  miMODE_MOBILECREDENTIALS = 2
End Enum

Public Property Get Choice() As Integer
  Choice = miChoice
End Property

Public Function AddToList( _
  pstrAddString As String, _
  Optional pstrOldValue As String, _
  Optional pstrNewValue As String) As Boolean
  
  Dim objItem As ComctlLib.ListItem
  
  ' Add the string to the usage listview
  Const TICKBOXWIDTH = 20
  Const BUFFERWIDTH = 5
  
  With lstUsage
    Set objItem = .ListItems.Add(, , pstrAddString)
    objItem.Tag = pstrAddString
    
    If Not IsMissing(pstrOldValue) Then
      objItem.SubItems(1) = pstrOldValue
    End If
    
    If Not IsMissing(pstrNewValue) Then
      objItem.SubItems(2) = pstrNewValue
    End If
    
    Set objItem = Nothing
    
    If mlngMaxTextLength < TextWidth(pstrAddString) Then
      mlngMaxTextLength = TextWidth(pstrAddString)
    End If
    
    If mlngMaxOldValLen < TextWidth(pstrOldValue) Then
      mlngMaxOldValLen = TextWidth(pstrOldValue)
    End If
    
    If mlngMaxNewValLen < TextWidth(pstrNewValue) Then
      mlngMaxNewValLen = TextWidth(pstrNewValue)
    End If
  End With
  
  AddToList = True
  
End Function

Private Sub FormatButtons()
  ' Format the buttons
  Dim iVisibleButtonCount As Integer
  Dim fVisible As Boolean
  
  Const BUTTONWIDTH = 1200
  Const BUTTONHEIGHT = 400
  
  iVisibleButtonCount = 0
  
  ' COPY
  With cmdCopy
    .Top = 0
    .Left = 0
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1
    
    fVisible = (miButtons And USAGEBUTTONS_COPY)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
  End With
  
  ' PRINT
  With cmdPrint
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1
    
    fVisible = (miButtons And USAGEBUTTONS_PRINT)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
  End With
    
  ' OK
  With cmdOK
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1
    
    fVisible = (miButtons And USAGEBUTTONS_OK)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
    
    If fVisible Then
      .Default = True
    End If
  End With
    
  ' YES
  With cmdYes
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1
    
    fVisible = (miButtons And USAGEBUTTONS_YES)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
  End With
    
  ' NO
  With cmdNo
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1
    
    fVisible = (miButtons And USAGEBUTTONS_NO)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
    
    If fVisible Then
      .Default = True
    End If
  End With
    
  With fraButtons
    .Height = cmdPrint.Height
    .Width = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET)) - BUTTON_OFFSET
  End With

End Sub

Public Sub ResetList()
  lstUsage.ListItems.Clear
  mlngMaxTextLength = 0
  mlngMaxOldValLen = 0
  mlngMaxNewValLen = 0
End Sub

Public Sub ShowMessage(Optional pvMode As Variant)
  Dim sMessage As String
  Dim sFrameCaption As String
     Dim iCount As Integer
  
  ' AE20080317 Fault #13011
'  lblUsageMSG.Caption = "The update script needs to be run on this system for the following reasons." _
'            & vbCrLf & vbCrLf & "Would you like to run the update script now?"
  
  With lstUsage
    If .ListItems.Count > 0 Then
      .ListItems(1).Selected = True
      .ListItems(1).EnsureVisible

      If TextWidth(.ColumnHeaders(2).Text) > mlngMaxOldValLen Then
        .ColumnHeaders(2).Width = TextWidth(.ColumnHeaders(2).Text)
      Else
        .ColumnHeaders(2).Width = (mlngMaxOldValLen + 1)
      End If
      
      If TextWidth(.ColumnHeaders(3).Text) > mlngMaxNewValLen Then
        .ColumnHeaders(3).Width = TextWidth(.ColumnHeaders(3).Text)
      Else
        .ColumnHeaders(3).Width = (mlngMaxNewValLen + 1)
      End If
      
      .ColumnHeaders(1).Width = (mlngMaxTextLength + 1)
       
      .Refresh
    End If
  End With
  
  If IsMissing(pvMode) Then
    miMode = miMODE_CHANGEDPLATFORM
  Else
    miMode = CInt(pvMode)
  End If
  
  Select Case miMode
    Case miMODE_WORKFLOWURLS
      sMessage = "The URLs for the following externally initiated Workflows have changed due to the change in server/database name."
      sFrameCaption = "Externally Initiated Workflows :"
      imgIcon(0).Picture = LoadResPicture("IMG_INFORMATION", 1)
      miButtons = USAGEBUTTONS_COPY + USAGEBUTTONS_PRINT + USAGEBUTTONS_OK
      Me.Caption = "External Workflow URLs"

      With lstUsage.ColumnHeaders
        .Item(1).Text = "Workflow"
        .Item(2).Text = "Old URL"
        .Item(3).Text = "New URL"
      End With

      Me.Width = (3 * Screen.Width / 4)
      Me.Height = (Screen.Height / 2)

    Case miMODE_MOBILECREDENTIALS
      sMessage = "Here are the Mobile Website entries for web.custom.config :"
      sFrameCaption = "Mobile Credentials"
      imgIcon(0).Picture = LoadResPicture("IMG_INFORMATION", 1)
      miButtons = USAGEBUTTONS_COPY + USAGEBUTTONS_PRINT + USAGEBUTTONS_OK
      Me.Caption = "Mobile Credentials"

      With lstUsage.ColumnHeaders
        .Item(1).Text = "Web.custom.config keys"
        ' Hide unrequired columns
        .Item(1).Width = lstUsage.Width
        .Item(2).Width = 0
        .Item(3).Width = 0
        
      End With

      Me.Width = (3 * Screen.Width / 4)
      Me.Height = (Screen.Height / 2)
      
    Case Else
      sMessage = "The update script and System Manager save needs to be run on this system for the reasons detailed below. It is strongly recommended to make a backup of your database before continuing." _
        & vbCrLf & vbCrLf & "Would you like to continue?"
      sFrameCaption = "Details :"
      imgIcon(0).Picture = LoadResPicture("IMG_EXCLAMATION", 1)
      miButtons = USAGEBUTTONS_PRINT + USAGEBUTTONS_YES + USAGEBUTTONS_NO
  
      Me.Width = ((mlngMaxTextLength / 1.5) * 3)
  End Select
  
  lblUsageMSG.Caption = sMessage
  fraUsage.Caption = sFrameCaption
  
  Call FormatButtons

  Me.Show vbModal
  
End Sub

Private Function PrintUsage() As Boolean
  
  Dim objPrintDef As clsPrintDef
  Dim blnOK As Boolean
  Dim sKey As String
  Dim iLoop As Integer
  
  On Error GoTo ErrorTrap
  
  Set objPrintDef = New clsPrintDef

  Screen.MousePointer = vbHourglass
  
  blnOK = (objPrintDef.IsOK)
  
  If blnOK Then
    With objPrintDef
      If .PrintStart(True) Then
        ' AE20080317 Fault #13012
        '.PrintHeader Me.Caption & " Details"
        .PrintHeader Me.Caption
  
        For iLoop = 1 To lstUsage.ListItems.Count
          Select Case miMode
            Case miMODE_WORKFLOWURLS
              Printer.Font.Underline = True
              .PrintBold lstUsage.ListItems(iLoop)
              Printer.Font.Underline = False
              
              .PrintNormal "Old URL: " & lstUsage.ListItems(iLoop).SubItems(1)
              .PrintNormal "New URL: " & lstUsage.ListItems(iLoop).SubItems(2)
              .PrintNormal ""
            
            Case miMODE_MOBILECREDENTIALS
              .PrintNormal lstUsage.ListItems(iLoop)
                
            Case Else
              .PrintNormal lstUsage.ListItems(iLoop) & " : " _
                & lstUsage.ListItems(iLoop).SubItems(1) & " -> " _
                & lstUsage.ListItems(iLoop).SubItems(2)
          End Select
        Next iLoop
        
        .PrintEnd
      End If
    End With
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Unable to print the list." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Sub ResizeForm()
  On Error GoTo ErrorTrap
  
  Dim sngOffset As Single
  
'  If Me.Height < MIN_FORM_HEIGHT Then Me.Height = MIN_FORM_HEIGHT
'  If Me.Width < MIN_FORM_WIDTH Then Me.Width = MIN_FORM_WIDTH
  
  lblUsageMSG.Width = Me.ScaleWidth - imgIcon(0).Width - (3 * MSG_OFFSET)
  lblUsageMSG.Top = 300
  lblUsageMSG.Left = 960
  
  imgIcon(0).Top = MSG_OFFSET
  imgIcon(0).Left = MSG_OFFSET
  imgIcon(0).Width = 480
  imgIcon(0).Height = imgIcon(0).Width
  
  fraButtons.Top = Me.ScaleHeight - BUTTON_OFFSET - fraButtons.Height
  fraButtons.Left = Me.ScaleWidth - OFFSET_RIGHT - fraButtons.Width
  
  sngOffset = lblUsageMSG.Top + lblUsageMSG.Height
  If (sngOffset < imgIcon(0).Top + imgIcon(0).Height) Then
    sngOffset = imgIcon(0).Top + imgIcon(0).Height
  End If
  
  fraUsage.Top = sngOffset + 60
  fraUsage.Left = OFFSET_LEFT
  fraUsage.Width = Me.ScaleWidth - OFFSET_LEFT - OFFSET_RIGHT
  fraUsage.Height = fraButtons.Top - BUTTON_OFFSET - fraUsage.Top
  
  lstUsage.Top = OFFSET_TOP
  lstUsage.Left = OFFSET_LEFT
  lstUsage.Width = fraUsage.Width - OFFSET_LEFT - OFFSET_RIGHT
  lstUsage.Height = fraUsage.Height - OFFSET_TOP - OFFSET_BOTTOM
  
ErrorTrap:

End Sub

Private Sub cmdCopy_Click()
  CopyData

End Sub

Private Function CopyData() As Boolean
  
  Dim iLoop As Integer

  On Error GoTo ErrorTrap

  Clipboard.Clear
  
  Screen.MousePointer = vbHourglass
  
  Clipboard.SetText lblUsageMSG.Caption & vbNewLine & vbNewLine

  For iLoop = 1 To lstUsage.ListItems.Count
    Select Case miMode
      Case miMODE_WORKFLOWURLS
        Clipboard.SetText Clipboard.GetText & _
          "Workflow: " & lstUsage.ListItems(iLoop) & vbNewLine
        Clipboard.SetText Clipboard.GetText & _
          vbTab & "Old URL: " & lstUsage.ListItems(iLoop).SubItems(1) & vbNewLine
        Clipboard.SetText Clipboard.GetText & _
          vbTab & "New URL: " & lstUsage.ListItems(iLoop).SubItems(2) & vbNewLine & vbNewLine
      
      Case miMODE_MOBILECREDENTIALS
          Clipboard.SetText Clipboard.GetText & lstUsage.ListItems(iLoop) & vbNewLine
      
      Case Else
        Clipboard.SetText Clipboard.GetText & _
          lstUsage.ListItems(iLoop) & " : " _
          & lstUsage.ListItems(iLoop).SubItems(1) & " -> " _
          & lstUsage.ListItems(iLoop).SubItems(2)
    End Select
  Next iLoop

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap:
  MsgBox "Unable to copy the list." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function


Private Sub cmdNo_Click()
  miChoice = vbNo
  UnLoad Me
End Sub

Private Sub cmdOK_Click()
  miChoice = vbOK
  UnLoad Me

End Sub

Private Sub cmdPrint_Click()
  PrintUsage
End Sub

Private Sub cmdYes_Click()
  miChoice = vbYes
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
  fraButtons.BorderStyle = vbBSNone
  
  Hook Me.hWnd, FORM_START_WIDTH, FORM_START_HEIGHT, Screen.Width, Screen.Height

  mlngMaxTextLength = 0
  
  lstUsage.View = lvwReport
  lstUsage.HideColumnHeaders = True
    
  With lstUsage.ColumnHeaders
    .Clear
    
    .Add , "Details", "Details", 2000, lvwColumnLeft
    .Add , "OldValue", "Old Value", 2000, lvwColumnLeft
    .Add , "NewValue", "New Value", 2000, lvwColumnLeft
  End With
  
  lstUsage.HideColumnHeaders = False
  lstUsage.Refresh
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    miChoice = vbCancel
  End If
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub
