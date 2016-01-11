VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000F&
      Height          =   1995
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2070
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
      Begin VB.CommandButton cmdButtons 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.CheckBox chkTheBox 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4000
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "frmMessage.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "frmMessage.frx":044E
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "frmMessage.frx":0D18
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmMessage.frx":115A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'Added as part of suggestion TM20010726 Fault 1607.                             *
'********************************************************************************

Option Explicit

'Message variables
Private msMessage As String
Private msTitle As String
Private msDetails As String
Private msFromErrorMessage As Boolean

'Message box style variables
Private miButtonStyle As Integer
Private miIconStyle As Integer
Private miDefaultButton As Integer
Private miModalType As Integer

Private miSelImage As Integer
Private miDfltButton As Integer

'Checkbox variables
Private mbCheckBox As Boolean
Private miCheckBoxValue As Integer
Private msCheckBoxMessage As String
 
Private miAnswer  As Integer

Private Const lngCharPerLine = 40

Private mlngReturnCode As Long

Private mblnUseCustomButtons As Boolean
Private mintCustomButtonIndex As Integer

Public Function CustomAddButton(pstrButtonName As String, plngReturnCode As Long) As Boolean

  If mintCustomButtonIndex > 0 Then
    Load cmdButtons(mintCustomButtonIndex)
  End If
  
  cmdButtons(mintCustomButtonIndex).Caption = pstrButtonName
  cmdButtons(mintCustomButtonIndex).Tag = plngReturnCode
  mintCustomButtonIndex = mintCustomButtonIndex + 1
  
End Function


Public Property Get DetailsLocked() As Boolean

  DetailsLocked = Me.txtDetails.Locked
  
End Property

Public Property Let DetailsLocked(bLocked As Boolean)

  Me.txtDetails.Locked = bLocked
  
End Property



Private Property Let CheckBoxMessage(sNewValue As String)

  msCheckBoxMessage = sNewValue
  
End Property

Private Property Get CheckBoxMessage() As String

  CheckBoxMessage = msCheckBoxMessage
  
End Property

Private Property Let CheckBoxValue(iNewValue As Integer)

  miCheckBoxValue = iNewValue
  
End Property

Private Property Get CheckBoxValue() As Integer
 
  CheckBoxValue = miCheckBoxValue
  
End Property

Private Property Let CheckBox(bNewValue As Boolean)

  mbCheckBox = bNewValue
  
End Property

Private Property Get CheckBox() As Boolean

  CheckBox = mbCheckBox
  
End Property


Private Property Get ReturnCode() As Long
  ReturnCode = mlngReturnCode
End Property

Private Property Let ReturnCode(plngNewValue As Long)
  mlngReturnCode = plngNewValue
End Property

Private Property Let Title(sNewValue As String)

  msTitle = sNewValue
  
End Property

Private Property Get Title() As String

  Title = msTitle
  
End Property

Private Property Let Message(sNewValue As String)

  msMessage = sNewValue
  
End Property

Private Property Get Message() As String

  Message = msMessage
  
End Property

Private Property Let DfltButton(iNewValue As Integer)

  miDfltButton = iNewValue
  
End Property

Private Property Get DfltButton() As Integer

  DfltButton = miDfltButton
  
End Property

Private Property Let SelectedImage(iNewValue As Integer)
  miSelImage = iNewValue
End Property

Private Property Get SelectedImage() As Integer

  SelectedImage = miSelImage
  
End Property

Private Property Let Answer(iNewValue As Integer)

  miAnswer = iNewValue
  
End Property

Private Property Get Answer() As Integer

  Answer = miAnswer
  
End Property

Private Property Let ButtonStyle(iNewValue As VbMsgBoxStyle)

  miButtonStyle = iNewValue
  
End Property

Private Property Get ButtonStyle() As VbMsgBoxStyle

  ButtonStyle = miButtonStyle
  
End Property

Private Property Let IconStyle(iNewValue As VbMsgBoxStyle)

  miIconStyle = iNewValue
  
End Property

Private Property Get IconStyle() As VbMsgBoxStyle

  IconStyle = miIconStyle
  
End Property

Private Property Let DefaultButton(iNewValue As VbMsgBoxStyle)

  miDefaultButton = iNewValue
  
End Property

Private Property Get DefaultButton() As VbMsgBoxStyle

  DefaultButton = miDefaultButton
  
End Property

Private Property Let ModalType(iNewValue As VbMsgBoxStyle)

  miModalType = iNewValue
  
End Property

Private Property Get ModalType() As VbMsgBoxStyle

  ModalType = miModalType
  
End Property

Private Function DecodeButtonStyle(iButtonStyle As Integer) As VbMsgBoxStyle

'********************************************************************************
' ButtonStyle - Returns required ButtonStyle code.                              *
'********************************************************************************

  Select Case iButtonStyle
    Case 0
      DecodeButtonStyle = vbOKOnly
    Case 1
      DecodeButtonStyle = vbOKCancel
    Case 2
      DecodeButtonStyle = vbAbortRetryIgnore
    Case 3
      DecodeButtonStyle = vbYesNoCancel
    Case 4
      DecodeButtonStyle = vbYesNo
    Case 5
      DecodeButtonStyle = vbRetryCancel
    Case Else
      DecodeButtonStyle = vbOKOnly
  End Select
  
End Function

Private Function DecodeIconStyle(iIconStyle As Integer) As Integer

'********************************************************************************
' IconStyle - Returns required IconStyle code.                                  *
'********************************************************************************

  If iIconStyle < 5 Then
    DecodeIconStyle = 0
    SelectedImage = -1
  ElseIf iIconStyle <= 21 Then
    DecodeIconStyle = vbCritical
    SelectedImage = 0
  ElseIf iIconStyle <= 37 Then
    DecodeIconStyle = vbQuestion
    SelectedImage = 1
  ElseIf iIconStyle <= 53 Then
    DecodeIconStyle = vbExclamation
    SelectedImage = 2
  ElseIf iIconStyle <= 69 Then
    DecodeIconStyle = vbInformation
    SelectedImage = 3
  Else
    DecodeIconStyle = 0
    SelectedImage = -1
  End If

End Function

Private Function DecodeDefaultButton(iDefaultButton As Integer) As VbMsgBoxStyle

'********************************************************************************
' DefaultButton - Returns required DefaultButton code.                          *
'********************************************************************************

  If iDefaultButton < 256 Then
    DecodeDefaultButton = vbDefaultButton1
    DfltButton = 0
  ElseIf iDefaultButton <= 325 Then
    DecodeDefaultButton = vbDefaultButton2
    DfltButton = 1
  ElseIf iDefaultButton <= 581 Then
    DecodeDefaultButton = vbDefaultButton3
    DfltButton = 2
  ElseIf iDefaultButton <= 837 Then
    DecodeDefaultButton = vbDefaultButton4
    DfltButton = 3
  Else
    DecodeDefaultButton = vbDefaultButton1
    DfltButton = 0
  End If

End Function

Private Function DecodeModalType(iModalType As Integer) As VbMsgBoxStyle

'********************************************************************************
' ModalType - Returns required ModalType.                                        *
'********************************************************************************

  If iModalType >= 4094 And iModalType <= 4933 Then
    DecodeModalType = vbSystemModal
  Else
    DecodeModalType = vbApplicationModal
  End If

End Function

Private Sub FormatControls()

'********************************************************************************
' FormatControls -  Positions the enabled/visible elements on the form.         *
'********************************************************************************
  
  Const lngYCONTROLOFFSET = 200
  Const lngXCONTROLOFFSET = 250
  Const lngHeightPerLine = 195
  Const lngButtonWidth = 1200
  Const lngButtonGap = 100
  Const lngButtonHeight = 375
  Const lngFormWidth = 5000
  Const lngMessageWidth = 3500
  
  Dim lngXPos As Long
  Dim lngYPos As Long
  Dim iCount As Long
  
  lngXPos = 0
  lngYPos = 0
  
  With Me
    .Width = lngFormWidth
  End With
  
  'Set Size/Position properties of the Icon if required.
  If IconStyle <> 0 Then
    With Me.imgIcon(SelectedImage)
      .Top = lngYCONTROLOFFSET
      .Left = lngXCONTROLOFFSET
      lngYPos = .Top + .Height
      lngXPos = .Left + .Width
    End With
  End If

  'Set Size/Position properties of the Message.
  With Me.lblMessage
    If IconStyle = 0 Then
      .Left = (Me.Width / 2) - (.Width / 2)
    Else
      .Left = lngXPos + lngXCONTROLOFFSET
    End If
    .Width = lngMessageWidth
    .Top = lngYCONTROLOFFSET
    .Height = lngHeightPerLine * Round((MessageLength(Message) / lngCharPerLine) + 1)
    If .Top + .Height + lngYCONTROLOFFSET > lngYPos Then
      lngYPos = .Top + .Height + lngYCONTROLOFFSET
    End If
  End With
  
  lngXPos = 0
  
  'Set Size/Position properties of the loaded buttons.
  For iCount = 0 To Me.cmdButtons.Count - 1 Step 1
    With Me.cmdButtons(iCount)
      .Left = lngXPos
      .Top = 0
      .Width = lngButtonWidth
      .Height = lngButtonHeight
      lngXPos = .Left + .Width + lngButtonGap
    End With
  Next iCount

  'Set Size/Position properties of the buttons frame.
  With Me.fraButtons
    .Width = lngXPos
    .Height = lngButtonHeight
    .Left = (Me.Width / 2) - (.Width / 2)
    .Top = lngYPos
    lngYPos = .Top + .Height + lngYCONTROLOFFSET
  End With
  
  'Set Size/Position properties of the Checkbox if required.
  If CheckBox Then
    With Me.chkTheBox
      .Left = lngXCONTROLOFFSET
      .Top = lngYPos
      .Height = lngHeightPerLine * Round((Len(CheckBoxMessage) / lngCharPerLine) + 1)
      lngYPos = .Top + .Height
    End With
  End If

  'Details information
  With Me.txtDetails
    .Top = lngYPos
    If .Visible Then
      lngYPos = .Top + .Height
    End If
  End With

  With Me
    '.Height = lngYPos + (2 * lngYCONTROLOFFSET)
    .Height = lngYPos + (.Height - .ScaleHeight)
  End With

End Sub

Private Function MessageLength(sMessage As String) As Integer

'********************************************************************************
' MessageLength - Calculates the number of Carriage Returns in the string and   *
'                 returns the actual length of the string.                      *
'********************************************************************************

  Dim sChar As String
  Dim iLineFeeds As Integer
  Dim i As Integer
  Dim iReturnCount As Integer
  
  iLineFeeds = 0
  iReturnCount = 0
  
  If Len(sMessage) > 0 Then
    For i = 1 To Len(sMessage) Step 1
      sChar = Mid(sMessage, i, 1)
      If Asc(sChar) = 13 Or Asc(sChar) = 10 Then
        iReturnCount = iReturnCount + 1
        If iReturnCount Mod 4 = 0 Then
          iLineFeeds = iLineFeeds + 1
        End If
      Else
        iReturnCount = 0
      End If
      
    Next i
  
    MessageLength = Len(sMessage) + (iLineFeeds * lngCharPerLine)
  Else
    MessageLength = 0
  End If
  
End Function

Public Function MessageBox(sPrompt As String, Buttons As Integer, _
                            sTitle As String, iCheckBoxValue As Integer, _
                            Optional sCheckBoxMessage As String) As VbMsgBoxResult

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmMessageBox.MessageBox(sPrompt,Buttons,sTitle,iCheckBoxValue,sCheckBoxMessage)", Array(sPrompt, Buttons, sTitle, iCheckBoxValue, sCheckBoxMessage)
  
  Dim iButtonsConv As Integer
  Dim fProgressVisible As Boolean

  iButtonsConv = CInt(Buttons)
  
  Message = Left(sPrompt, 1048)
  Title = sTitle
  
  CheckBox = Not IsMissing(iCheckBoxValue)
  
  If CheckBox Then
    CheckBoxValue = iCheckBoxValue
    CheckBoxMessage = Left(sCheckBoxMessage, 255)
  Else
    Me.Height = Me.Height - chkTheBox.Height
  End If
  
  If Not (IsMissing(Buttons)) Then
    DefineControls (iButtonsConv)
  End If

  fProgressVisible = gobjProgress.Visible
  If fProgressVisible Then
    gobjProgress.Visible = False
  End If

  Screen.MousePointer = vbDefault
  Me.Show IIf(ModalType = vbApplicationModal, vbModal, ModalType)
  gobjProgress.Visible = fProgressVisible
 
  If Not IsMissing(iCheckBoxValue) Then
    iCheckBoxValue = CheckBoxValue
  End If
  MessageBox = Answer
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function CustomMessageBox(sPrompt As String, Image As Integer, _
                            Optional sTitle As String, Optional iCheckBoxValue As Variant, _
                            Optional sCheckBoxMessage As String) As Long

  'Uses the buttons that have been added using 'AddButton' function.
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmMessageBox.CustomMessageBox(sPrompt,Image,sTitle,iCheckBoxValue,sCheckBoxMessage)", Array(sPrompt, Image, sTitle, iCheckBoxValue, sCheckBoxMessage)
  
  Dim fProgressVisible As Boolean

  mblnUseCustomButtons = True
  
  Message = Left(sPrompt, 1048)
  Title = sTitle
  
  CheckBox = Not IsMissing(iCheckBoxValue)
  
  If CheckBox Then
    CheckBoxValue = iCheckBoxValue
    CheckBoxMessage = Left(sCheckBoxMessage, 255)
  Else
    Me.Height = Me.Height - chkTheBox.Height
  End If
  
  DefineCustomControls (Image)

  fProgressVisible = gobjProgress.Visible
  If fProgressVisible Then
    gobjProgress.Visible = False
  End If

  Screen.MousePointer = vbDefault
  Me.Show IIf(ModalType = vbApplicationModal, vbModal, ModalType)
  gobjProgress.Visible = fProgressVisible
 
  If Not IsMissing(iCheckBoxValue) Then
    iCheckBoxValue = CheckBoxValue
  End If
  
  CustomMessageBox = ReturnCode
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Private Sub DefineControls(iButtons As Integer)

'********************************************************************************
' DefineControls -  Sets the visible and enabled properties of the relevent     *
'                   controls on the form depending on the message box options   *
'                   defined in the maryStyle() array.                           *
'********************************************************************************

  DecodeStyle iButtons
  
  LoadButtons (ButtonStyle)
  
  ShowButtons
  
  ShowImage
  
  ShowCheckBox
  
  ShowMessage

  ShowDetails

  FormatControls

End Sub

Private Sub DefineCustomControls(Image As Integer)

'**************************************************************************************
' DefineCustomControls -  Sets the visible and enabled properties of the relevent     *
'                         controls on the form depending on the message box options   *
'                         custom                                                      *
'**************************************************************************************

  IconStyle = DecodeIconStyle(Image)
  
  ShowButtons
  
  ShowImage
  
  ShowCheckBox
  
  ShowMessage

  FormatControls

End Sub

Private Sub ShowMessage()

  Me.lblMessage.Caption = Message
  Me.Caption = Title
  
End Sub

Private Sub ShowCheckBox()
  
  With Me.chkTheBox
    If CheckBox Then
      .Enabled = True
      .Visible = True
      .Caption = CheckBoxMessage
      .Value = CheckBoxValue
    Else
      .Visible = False
      .Visible = False
    End If
  End With
  
End Sub

Private Sub ShowImage()

'********************************************************************************
' ShowImage - Enables/Shows the required image. Hides remaining.                *
'********************************************************************************

  Dim iCount As Integer
  
  For iCount = 0 To Me.imgIcon.Count - 1 Step 1
    If SelectedImage = iCount Then
      With Me.imgIcon(iCount)
        .Enabled = True
        .Visible = True
      End With
    Else
      With Me.imgIcon(iCount)
        .Enabled = False
        .Visible = False
      End With
    End If
  Next iCount
  
End Sub

Private Sub ShowButtons()

'********************************************************************************
' ShowButtons - Enables/Shows the required buttons.                             *
'********************************************************************************

  Dim iCount As Integer
  
  For iCount = 0 To Me.cmdButtons.Count - 1 Step 1
    Me.cmdButtons(iCount).Enabled = True
    Me.cmdButtons(iCount).Visible = True
  Next iCount

End Sub

Private Sub LoadButtons(iButtonType As Integer)

'********************************************************************************
' LoadButtons -   Loads the buttons into a control array.                       *
'********************************************************************************

  Dim iCount As Integer
  Dim iButtonCount As Integer
  
  Select Case iButtonType
    Case 0
      iButtonCount = 0
    Case 1, 4, 5
      iButtonCount = 1
    Case 2, 3
      iButtonCount = 2
  End Select

  For iCount = 1 To iButtonCount Step 1
    Load Me.cmdButtons(iCount)
  Next iCount
  
  CaptionButtons (iButtonType)

End Sub

Private Sub CaptionButtons(iButtonType As Integer)

'********************************************************************************
' CaptionButtons -  Sets the caption property of the buttons.                   *
'********************************************************************************

  Dim iCount As Integer
  
  Select Case iButtonType
    Case 0
      Me.cmdButtons(0).Caption = "&OK"
      Me.cmdButtons(0).Tag = "OK"
    Case 1
      Me.cmdButtons(0).Caption = "&OK"
      Me.cmdButtons(0).Tag = "OK"
      Me.cmdButtons(1).Caption = "&Cancel"
      Me.cmdButtons(1).Tag = "CANCEL"
    Case 2
      Me.cmdButtons(0).Caption = "&Abort"
      Me.cmdButtons(0).Tag = "ABORT"
      Me.cmdButtons(1).Caption = "&Retry"
      Me.cmdButtons(1).Tag = "RETRY"
      Me.cmdButtons(2).Caption = "&Ignore"
      Me.cmdButtons(2).Tag = "IGNORE"
    Case 3
      Me.cmdButtons(0).Caption = "&Yes"
      Me.cmdButtons(0).Tag = "YES"
      Me.cmdButtons(1).Caption = "&No"
      Me.cmdButtons(1).Tag = "NO"
      Me.cmdButtons(2).Caption = "&Cancel"
      Me.cmdButtons(2).Tag = "CANCEL"
    Case 4
      Me.cmdButtons(0).Caption = "&Yes"
      Me.cmdButtons(0).Tag = "YES"
      Me.cmdButtons(1).Caption = "&No"
      Me.cmdButtons(1).Tag = "NO"
    Case 5
      Me.cmdButtons(0).Caption = "&Retry"
      Me.cmdButtons(0).Tag = "RETRY"
      Me.cmdButtons(1).Caption = "&Cancel"
      Me.cmdButtons(1).Tag = "CANCEL"
  End Select

End Sub

Private Sub DecodeStyle(i As Integer)

'********************************************************************************
' DecodeStyle - Calculates the selected message box style from the integer      *
'               parameter, for the Button Type, Icon Style, Default Button and  *
'               Modal Type.                                                     *
'********************************************************************************

  If i >= 0 Then
    ' Set the Modal type.
    ModalType = DecodeModalType(i)
    i = i - ModalType
    
    ' Set the Default Button type.
    DefaultButton = DecodeDefaultButton(i)
    i = i - DefaultButton
  
    ' Set the Icon style.
    IconStyle = DecodeIconStyle(i)
    i = i - IconStyle
    
    ' Set the Button style.
    ButtonStyle = DecodeButtonStyle(i)
  Else
    ' Output default Message box style.
    ButtonStyle = vbOKOnly
    IconStyle = False
    DefaultButton = vbDefaultButton1
    ModalType = vbApplicationModal
  End If
  
End Sub

Private Sub DefaultButtonFocus()

'********************************************************************************
' DefaultButtonFocus -  Sets the focus to the default button.                   *
'********************************************************************************

  Me.cmdButtons.Item(DfltButton).SetFocus
  
End Sub

Private Sub chkTheBox_Click()

  CheckBoxValue = Me.chkTheBox.Value

End Sub

Private Sub cmdButtons_Click(Index As Integer)

  Dim mbUnload As Boolean

  mbUnload = True

  If mblnUseCustomButtons Then
    ReturnCode = CLng(Me.cmdButtons(Index).Tag)
    
  Else
  
    With Me.cmdButtons(Index)
      Select Case .Tag
        Case "OK"
          Answer = vbOK
        Case "CANCEL"
          Answer = vbCancel
        Case "ABORT"
          Answer = vbAbort
        Case "RETRY"
          Answer = vbRetry
        Case "IGNORE"
          Answer = vbIgnore
        Case "YES"
          Answer = vbYes
        Case "NO"
          Answer = vbNo
        Case "DETAILS"
          ' Hide/show the details window
          If txtDetails.Visible Then
            txtDetails.Visible = False
            Me.Height = Me.Height - txtDetails.Height - 100
          Else
            txtDetails.Visible = True
            Me.Height = Me.Height + txtDetails.Height + 100
          End If
          
          ' Change the caption on the error box
          cmdButtons(2).Caption = "&Details " & IIf(txtDetails.Visible, "<<", ">>")
          
          Answer = vbOK
          mbUnload = False
          
        ' User has selected to exit the application
        Case "EXIT"
          Answer = vbAbort
          mbUnload = True
          
        Case "IGNORE"
          Answer = vbIgnore
          mbUnload = True
          
      End Select
    End With
  
  End If
  
  If mbUnload Then
    Unload Me
  End If
 
End Sub

Private Sub Form_Activate()

  DefaultButtonFocus
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = Asc("N") Then
    cmdButtons_Click 1
  
  ElseIf KeyCode = Asc("Y") Then
    cmdButtons_Click 0
      
  'MH20020820
  'Print Screen
  ElseIf KeyCode = 44 And FromErrorMessage Then
    If COAMsgBox("Would you like to email details of this error to the helpdesk?", vbQuestion + vbYesNo, "Email Error") = vbYes Then
      frmEmailSel.SendEmail _
        GetSystemSetting("support", "email", "ohrsupport@advancedcomputersoftware.com"), _
        app.Title & " Error", _
        Me.txtDetails, False, True
      Set frmEmailSel = Nothing
    End If

  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    If UnloadMode = vbFormControlMenu Then
      Me.chkTheBox.Value = 0
    End If
    
    If mblnUseCustomButtons Then
      ReturnCode = vbCancel
    Else
      cmdButtons_Click (DfltButton)
    End If
  End If
    
End Sub

Public Function ErrorBox(pstrDetails As String) As Integer

  Dim fProgressVisible As Boolean

  Message = "An unexpected program error has occurred."
  Title = App.ProductName
  details = pstrDetails
  
  FromErrorMessage = True
  
  'TM20020104 Fault 3029 - Lock the details section when COAMsgBox used as error box.
  Me.DetailsLocked = True
  
  DefineControls (vbYesNoCancel + vbCritical)

  'Make the cancel button a display details button
  cmdButtons(2).Caption = "&Details >>"
  cmdButtons(2).Tag = "DETAILS"

  'Make the Yes Button a quit application
  cmdButtons(0).Caption = "E&xit"
  cmdButtons(0).Tag = "EXIT"

  'Make the No Button an ignore option
  cmdButtons(1).Caption = "&Ignore"
  cmdButtons(1).Tag = "IGNORE"
  cmdButtons(1).Enabled = ASRDEVELOPMENT

  ' Get rid of progress bar
  fProgressVisible = gobjProgress.Visible
  If fProgressVisible Then
    gobjProgress.Visible = False
  End If

  Screen.MousePointer = vbDefault
  Me.Show vbModal
  gobjProgress.Visible = fProgressVisible
  ErrorBox = Answer
  
End Function

Private Property Let details(sNewValue As String)

  msDetails = sNewValue
  
End Property

Private Property Get details() As String

  details = msDetails
  
End Property

Private Sub ShowDetails()

  Me.txtDetails.Text = details
  
End Sub

Public Property Get FromErrorMessage() As Boolean
  'NHRD - 11042003 - Added this property for Fault 2969
  FromErrorMessage = msFromErrorMessage
End Property

Public Property Let FromErrorMessage(ByVal vNewValue As Boolean)
  msFromErrorMessage = vNewValue
End Property

