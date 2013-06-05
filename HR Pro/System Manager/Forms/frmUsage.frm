VERSION 5.00
Begin VB.Form frmUsage 
   Caption         =   "???? Usage"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5065
   Icon            =   "frmUsage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraButtons 
      Caption         =   "Frame1"
      Height          =   3500
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1200
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy"
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   3000
         Width           =   1200
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "&Fix"
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   2500
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   400
         Left            =   0
         TabIndex        =   8
         Top             =   2000
         Width           =   1200
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   400
         Left            =   0
         TabIndex        =   7
         Top             =   1500
         Width           =   1200
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         Height          =   400
         Left            =   0
         TabIndex        =   6
         Top             =   1000
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   500
         Width           =   1200
      End
   End
   Begin VB.Frame fraUsage 
      Caption         =   "Current Usage :"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2250
      Begin VB.ListBox lstUsage 
         Height          =   405
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1680
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmUsage.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblUsageMSG 
      Height          =   585
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OFFSET_LEFT = 160
Private Const OFFSET_RIGHT = 160
Private Const OFFSET_TOP = 300
Private Const OFFSET_BOTTOM = 160
Private Const MSG_OFFSET = 240
Private Const BUTTON_OFFSET = 95
Private Const MIN_FORM_HEIGHT = 4500
Private Const MIN_FORM_WIDTH = 6600

Private miChoice As Integer
Private miButtons As Integer
Private msMode As String
Private mlngSelection As Long

Public Enum UsageButtonOptions
  USAGEBUTTONS_OK = 2 ^ 0
  USAGEBUTTONS_YES = 2 ^ 1
  USAGEBUTTONS_NO = 2 ^ 2
  USAGEBUTTONS_PRINT = 2 ^ 3
  USAGEBUTTONS_SELECT = 2 ^ 4
  USAGEBUTTONS_FIX = 2 ^ 5
  USAGEBUTTONS_COPY = 2 ^ 6
End Enum

Private mlngMaxTextLength As Long

Public Property Get Selection() As Long
  Selection = mlngSelection
  
End Property


Public Property Get Choice() As Integer
  Choice = miChoice
  
End Property

Public Function AddToList(pstrAddString As String, Optional pvItemData As Variant) As Boolean
  ' Add the string to the usage listbox (with the ItemData if provided)
  ' The ItemData is used when the SELECT button is used, to identify the selected row.
  ' See workflow definitions for an example of this is action.
  Const TICKBOXWIDTH = 20
  Const BUFFERWIDTH = 5
  
  With lstUsage
    .AddItem pstrAddString
    
    If Not IsMissing(pvItemData) Then
      .ItemData(.NewIndex) = CLng(pvItemData)
    End If
  
    If mlngMaxTextLength < TextWidth(pstrAddString) Then
      mlngMaxTextLength = TextWidth(pstrAddString)
      
      If ScaleMode = vbTwips Then
        SendMessageLong .hWnd, LB_SETHORIZONTALEXTENT, (mlngMaxTextLength / Screen.TwipsPerPixelX) + BUFFERWIDTH, 0
      Else
        SendMessageLong .hWnd, LB_SETHORIZONTALEXTENT, mlngMaxTextLength + BUFFERWIDTH, 0
      End If
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
  
  ' FIX
  With cmdFix
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1

    fVisible = (miButtons And USAGEBUTTONS_FIX)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
    
    .Enabled = (lstUsage.ListCount > 0)
  End With
    
  ' SELECT
  With cmdSelect
    .Top = 0
    .Left = (iVisibleButtonCount * (BUTTONWIDTH + BUTTON_OFFSET))
    .Height = BUTTONHEIGHT
    .Width = BUTTONWIDTH
    .TabIndex = lstUsage.TabIndex + iVisibleButtonCount + 1

    fVisible = (miButtons And USAGEBUTTONS_SELECT)
    .Visible = fVisible
    iVisibleButtonCount = iVisibleButtonCount + IIf(fVisible, 1, 0)
    
    .Enabled = (lstUsage.ListCount > 0)
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
  Do While (lstUsage.ListCount > 0)
    lstUsage.RemoveItem (0)
  Loop
  mlngMaxTextLength = 0

End Sub


Public Function ShowMessage(pstrCaption As String, _
  pstrMessage As String, _
  pintCheckObject As UsageCheckObject, _
  Optional pvButtons As Variant, _
  Optional pvMode As Variant)
  
  Dim sFrameCaption
  
  mlngSelection = -1
   
  Me.Caption = pstrCaption
  
  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
  
  lblUsageMSG.Caption = pstrMessage
  
  If IsMissing(pvMode) Then
    msMode = "usage"
  Else
    msMode = CStr(pvMode)
  End If
  
  Select Case msMode
    Case "validation"
      sFrameCaption = "Validation exceptions :"
      imgIcon(0).Picture = LoadResPicture("IMG_EXCLAMATION", 1)
    Case "initiator"
      sFrameCaption = "Initiator's Record elements :"
      imgIcon(0).Picture = LoadResPicture("IMG_INFORMATION", 1)
    Case "triggered"
      sFrameCaption = "Triggered Record elements :"
      imgIcon(0).Picture = LoadResPicture("IMG_INFORMATION", 1)
    Case "workflowURLs"
      sFrameCaption = "Externally Initiated Workflows :"
      imgIcon(0).Picture = LoadResPicture("IMG_INFORMATION", 1)
    Case "details"
      sFrameCaption = "Details :"
      imgIcon(0).Picture = LoadResPicture("IMG_EXCLAMATION", 1)
    Case Else
      sFrameCaption = "Current Usage :"
      imgIcon(0).Picture = LoadResPicture("IMG_EXCLAMATION", 1)
  End Select
  fraUsage.Caption = sFrameCaption
  
  If IsMissing(pvButtons) Then
    miButtons = USAGEBUTTONS_OK + USAGEBUTTONS_PRINT
  Else
    miButtons = CInt(pvButtons)
  End If
  FormatButtons

  If lstUsage.ListCount > 0 Then
    lstUsage.Selected(0) = True
  
    If (miButtons And USAGEBUTTONS_SELECT) Then
      cmdSelect.Enabled = (lstUsage.ItemData(lstUsage.ListIndex) >= 0)
    End If
  End If
  
  Me.Show vbModal
  
End Function

Private Function CopyUsage() As Boolean
  
  Dim iLoop As Integer

  On Error GoTo ErrorTrap

  Clipboard.Clear
  
  Screen.MousePointer = vbHourglass
  Clipboard.SetText lblUsageMSG.Caption & vbNewLine
  Clipboard.SetText Clipboard.GetText & String(Len(lblUsageMSG.Caption), "_") & vbNewLine & vbNewLine
  
  For iLoop = 0 To lstUsage.ListCount - 1
    Clipboard.SetText Clipboard.GetText & lstUsage.List(iLoop) & vbNewLine & vbNewLine
  Next iLoop

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap:
  MsgBox "Unable to copy the list." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Function PrintUsage() As Boolean
  
  Dim objPrintDef As clsPrintDef
  Dim blnOK As Boolean
  Dim sKey As String
  Dim iLoop As Integer
  
  On Error GoTo ErrorTrap
  
  Set objPrintDef = New clsPrintDef
  
  blnOK = (objPrintDef.IsOK)
  
  If blnOK Then
    Screen.MousePointer = vbHourglass
    
    With objPrintDef
      If .PrintStart(True) Then
        .PrintHeader Me.Caption & " " & msMode
  
        For iLoop = 0 To lstUsage.ListCount - 1
          .PrintNormal lstUsage.List(iLoop)
        Next iLoop
        
        .PrintEnd
      End If
    End With
  End If

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap:
  MsgBox "Unable to print the list." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function


Private Sub ResizeForm()
  On Error GoTo ErrorTrap
  
  lblUsageMSG.Width = Me.ScaleWidth - imgIcon(0).Width - (3 * MSG_OFFSET)
  lblUsageMSG.Top = 300
  lblUsageMSG.Left = 960
  lblUsageMSG.Height = 600
  
  imgIcon(0).Top = MSG_OFFSET
  imgIcon(0).Left = MSG_OFFSET
  imgIcon(0).Width = 480
  imgIcon(0).Height = imgIcon(0).Width
  
  fraButtons.Top = Me.ScaleHeight - BUTTON_OFFSET - fraButtons.Height
  fraButtons.Left = Me.ScaleWidth - OFFSET_RIGHT - fraButtons.Width
  
  fraUsage.Top = 960
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
  CopyUsage
  
End Sub

Private Sub cmdFix_Click()
  miChoice = vbIgnore
  UnLoad Me

End Sub

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

Private Sub cmdSelect_Click()
  miChoice = vbRetry
  mlngSelection = lstUsage.ItemData(lstUsage.ListIndex)
  UnLoad Me

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
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  RemoveIcon Me
  
  mlngMaxTextLength = 0

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

Private Sub lstUsage_Click()
  If cmdSelect.Visible Then
    cmdSelect.Enabled = (lstUsage.ItemData(lstUsage.ListIndex) >= 0)
  End If
  
End Sub

Private Sub lstUsage_DblClick()
  If cmdSelect.Visible And cmdSelect.Enabled Then
    cmdSelect_Click
  End If
  
End Sub

