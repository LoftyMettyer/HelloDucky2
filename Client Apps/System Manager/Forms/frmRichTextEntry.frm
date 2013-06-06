VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRichTextEntry 
   Caption         =   "Message"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5081
   Icon            =   "frmRichTextEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectText 
      Caption         =   "&Select Hypertext"
      Height          =   400
      Left            =   100
      TabIndex        =   2
      Top             =   3200
      Width           =   1710
   End
   Begin VB.Frame fraMessage 
      Caption         =   "Message :"
      Height          =   3000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   5000
      Begin RichTextLib.RichTextBox txtMessage 
         Height          =   2500
         Left            =   200
         TabIndex        =   1
         Top             =   300
         Width           =   4600
         _ExtentX        =   8123
         _ExtentY        =   4419
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         MaxLength       =   200
         AutoVerbMenu    =   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmRichTextEntry.frx":000C
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3900
      TabIndex        =   4
      Top             =   3200
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2600
      TabIndex        =   3
      Top             =   3200
      Width           =   1200
   End
End
Attribute VB_Name = "frmRichTextEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OFFSET_TOP = 300
Private Const OFFSET_BOTTOM = 160
Private Const BUTTON_OFFSET = 95
Private Const MIN_FORM_HEIGHT = 4000
Private Const MIN_FORM_WIDTH = 5000

Private miSelStart As Integer
Private miSelLength As Integer

Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfReadOnly As Boolean

Private Function ValidateText() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sPart1 As String
  Dim sPart2 As String
  Dim sPart3 As String
  
  fOK = True
  
  DeconstructMessage txtMessage.TextRTF, _
    sPart1, _
    sPart2, _
    sPart3

  If Len(Replace(Replace(sPart2, vbCr, ""), vbLf, "")) = 0 Then
    MsgBox "Some hypertext must be selected.", vbOKOnly + vbExclamation, Application.Name
    fOK = False
  End If

TidyUpAndExit:
  ValidateText = fOK
  Exit Function
  
ErrorTrap:
  fOK = True
  Resume TidyUpAndExit
  
End Function



Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = mfCancelled
  
End Property

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property


Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fOKToSave As Boolean
  
  fOKToSave = mfChanged And (Not mfReadOnly)
  
  cmdOK.Enabled = fOKToSave

End Sub



Public Property Let Changed(ByVal pfNewValue As Boolean)
  mfChanged = pfNewValue
  RefreshScreen
  
End Property
Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
End Property
Public Sub Initialise(piWhichMessage As WorkflowWebFormMessageType, _
  psMessage As String, _
  pfReadOnly As Boolean)
        
  Dim sNewRTFText As String
  
  Select Case piWhichMessage
    Case WORKFLOWWEBFORMMESSAGE_COMPLETION
      Me.Caption = "Completion Message"
      fraMessage.Caption = "Completion Message :"
      
    Case WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS
      Me.Caption = "Follow On Forms Message"
      fraMessage.Caption = "Follow On Forms Message :"
      
    Case WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER
      Me.Caption = "Save For Later Message"
      fraMessage.Caption = "Save For Later Message :"
      
  End Select

  mfReadOnly = pfReadOnly

  sNewRTFText = "{\rtf1 " & _
    Replace(psMessage, vbCrLf, "\par ") & _
    "\par }"

  txtMessage.TextRTF = sNewRTFText

  If mfReadOnly Then
    ControlsDisableAll Me
  End If

  mfChanged = False
  RefreshScreen
  RemoveIcon Me
  
End Sub

Private Sub ParseMessage(pctlRichTextbox As RichTextBox)
  Dim sSourceText As String
  Dim asText() As String
  Dim iIndex As Integer
  Dim sChar As String
  Dim sNextChar As String
  Dim fDoingSlash As Boolean
  Dim iBracketLevel As Integer
  Dim fIgnoreChar As Boolean
  Dim sRTFCode As String
  Dim sRTFCodeToDo As String
  Dim iTextIndex As Integer
  Dim fLiteral As Boolean
  Dim iSelStart As Integer
  Dim iSelLength As Integer
  Dim sTemp As String
  Dim asDeniedCharacters() As String
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim sDeniedChar As String
  Dim iDeniedCharCount As Integer
  Dim sNewRTFText As String

  iTextIndex = 0
  ReDim asText(2)
  ReDim asDeniedCharacters(0)
  sSourceText = pctlRichTextbox.TextRTF
  fDoingSlash = False
  iBracketLevel = 0
  sRTFCode = ""

  ' Strip RTF tag brackets.
  iIndex = InStr(sSourceText, "{")
  If iIndex > 0 Then
    sSourceText = Mid(sSourceText, iIndex + 1)

    iIndex = InStrRev(sSourceText, "}")
    If iIndex > 0 Then
      sSourceText = Mid(sSourceText, 1, iIndex - 1)
    End If
  End If

  Do While Len(sSourceText) > 0
    sChar = Mid(sSourceText, 1, 1)
    sNextChar = Mid(sSourceText, 2, 1)

    fLiteral = sChar = "\" _
      And ((sNextChar = "\") _
        Or (sNextChar = "{") _
        Or (sNextChar = "}"))

    fIgnoreChar = fDoingSlash Or (iBracketLevel > 0)

    If fDoingSlash Then
      If sChar = " " _
        Or sChar = "{" Then

        fDoingSlash = False
        sRTFCodeToDo = sRTFCode
        sRTFCode = ""
      ElseIf sChar = "\" Then
        sRTFCodeToDo = sRTFCode
        sRTFCode = sChar
      Else
        sRTFCode = sRTFCode & Trim(Replace(Replace(sChar, vbCr, ""), vbLf, ""))
      End If
    End If

    If (iBracketLevel > 0) And sChar = "}" Then
      iBracketLevel = iBracketLevel - 1
      sRTFCodeToDo = sRTFCode
      sRTFCode = ""
    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
        fDoingSlash = True
        sRTFCode = sChar
      ElseIf sChar = "{" Then
        iBracketLevel = iBracketLevel + 1
        sRTFCodeToDo = sRTFCode
        sRTFCode = ""
      ElseIf Not fIgnoreChar Then
'''        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
'''      asText(iTextIndex) = asText(iTextIndex) & sNextChar
    End If

    ' See if we need to interpret the RTF control code.
    If Len(sRTFCodeToDo) > 0 Then
      If ((sRTFCodeToDo = "\ul") And (iTextIndex = 0)) _
        Or ((sRTFCodeToDo = "\ulnone") And (iTextIndex = 1)) Then
        iTextIndex = iTextIndex + 1
      ElseIf (sRTFCodeToDo = "\tab") Or (sRTFCodeToDo = "\cell") Then
        asText(iTextIndex) = asText(iTextIndex) & vbTab
      ElseIf (sRTFCodeToDo = "\row") Then
        asText(iTextIndex) = asText(iTextIndex) & vbNewLine
      ElseIf (Mid(sRTFCodeToDo, 1, 2) = "\'") Then
        fFound = False
        sDeniedChar = Chr(val("&H" & Mid(sRTFCodeToDo, 3)))
        For iLoop = 1 To UBound(asDeniedCharacters)
          If sDeniedChar = asDeniedCharacters(iLoop) Then
            fFound = True
            Exit For
          End If
        Next iLoop
        If Not fFound Then
          ReDim Preserve asDeniedCharacters(UBound(asDeniedCharacters) + 1)
          asDeniedCharacters(UBound(asDeniedCharacters)) = sDeniedChar
        End If
      End If

      sRTFCodeToDo = ""
    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
'''        fDoingSlash = True
'''        sRTFCode = sChar
      ElseIf sChar = "{" Then
'''        iBracketLevel = iBracketLevel + 1
'''        sRTFCodeToDo = sRTFCode
'''        sRTFCode = ""
      ElseIf Not fIgnoreChar Then
        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
      asText(iTextIndex) = asText(iTextIndex) & sNextChar
fDoingSlash = False
    End If

    ' Move forward through the text (jump an extra character if we've just processed a literal.
    If fLiteral Then
      sSourceText = Mid(sSourceText, 3)
    Else
      sSourceText = Mid(sSourceText, 2)
    End If
  Loop

  ' Trim off start CRLF
  If Len(asText(0)) >= 2 Then
    If Mid(asText(0), 1, 1) = vbCr _
      And Mid(asText(0), 2, 1) = vbLf Then

      asText(0) = Mid(asText(0), 3)
    End If
  End If

  ' Trim off end CRLF
  If Len(asText(iTextIndex)) >= 2 Then
    If Mid(asText(iTextIndex), Len(asText(iTextIndex)) - 1, 1) = vbCr _
      And Mid(asText(iTextIndex), Len(asText(iTextIndex)), 1) = vbLf Then

      asText(iTextIndex) = Mid(asText(iTextIndex), 1, Len(asText(iTextIndex)) - 2)
    End If
  End If

  ' Remember the current selection.
  iSelStart = pctlRichTextbox.SelStart
  iSelLength = pctlRichTextbox.SelLength

  ' Count the characters that have NOT made it through our parsing (eg. bullet point characters).
  ' These were all found during the parsing above.
  iDeniedCharCount = 0
  For iIndex = 1 To iSelStart
    For iLoop = 1 To UBound(asDeniedCharacters)
      If Mid(pctlRichTextbox.Text, iIndex, 1) = asDeniedCharacters(iLoop) Then
        iDeniedCharCount = iDeniedCharCount + 1
        Exit For
      End If
    Next iLoop
  Next iIndex

  ' Lock the display to avoid screen flash.
  UI.LockWindow pctlRichTextbox.hWnd
  
  ' Ensure no dodgy formatting remains in the text (eg. from text pasted from Word).
  With pctlRichTextbox
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelBold = .Font.Bold
    .SelItalic = .Font.Italic
    .SelFontName = .Font.Name
    .SelFontSize = .Font.Size
    .SelStrikeThru = .Font.Strikethrough
    .SelColor = vbBlack
  
    ' Remove all text to remove any formatting residue.
    .Text = ""
  End With
  
  ' Apply the processed RTF text.
  sNewRTFText = "{\rtf1 " & _
    Replace(Replace(Replace(Replace(asText(0), "\", "\\"), "{", "\{"), "}", "\}"), vbCrLf, "\par ") & _
    "\ul " & Replace(Replace(Replace(Replace(asText(1), "\", "\\"), "{", "\{"), "}", "\}") & "\ulnone ", vbCrLf, "\par ") & _
    Replace(Replace(Replace(Replace(asText(2), "\", "\\"), "{", "\{"), "}", "\}"), vbCrLf, "\par ") & _
    "\par }"

  pctlRichTextbox.TextRTF = sNewRTFText

  ' Select/position the textpointer.
  pctlRichTextbox.SelStart = iSelStart - iDeniedCharCount
  pctlRichTextbox.SelLength = iSelLength

  ' Unlock the display.
  UI.UnlockWindow

End Sub


Private Sub DeconstructMessage(psRichText As String, _
  ByRef psPart1 As String, _
  ByRef psPart2 As String, _
  ByRef psPart3 As String)
  
  Dim sSourceText As String
  Dim asText() As String
  Dim iIndex As Integer
  Dim sChar As String
  Dim sNextChar As String
  Dim fDoingSlash As Boolean
  Dim iBracketLevel As Integer
  Dim fIgnoreChar As Boolean
  Dim sRTFCode As String
  Dim sRTFCodeToDo As String
  Dim iTextIndex As Integer
  Dim fLiteral As Boolean
  Dim iSelStart As Integer
  Dim iSelLength As Integer
  Dim sTemp As String
  Dim asDeniedCharacters() As String
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim sDeniedChar As String
  Dim iDeniedCharCount As Integer
  Dim sNewRTFText As String

  iTextIndex = 0
  ReDim asText(2)
  ReDim asDeniedCharacters(0)
  sSourceText = psRichText
  fDoingSlash = False
  iBracketLevel = 0
  sRTFCode = ""

  ' Strip RTF tag brackets.
  iIndex = InStr(sSourceText, "{")
  If iIndex > 0 Then
    sSourceText = Mid(sSourceText, iIndex + 1)

    iIndex = InStrRev(sSourceText, "}")
    If iIndex > 0 Then
      sSourceText = Mid(sSourceText, 1, iIndex - 1)
    End If
  End If

  Do While Len(sSourceText) > 0
    sChar = Mid(sSourceText, 1, 1)
    sNextChar = Mid(sSourceText, 2, 1)

    fLiteral = sChar = "\" _
      And ((sNextChar = "\") _
        Or (sNextChar = "{") _
        Or (sNextChar = "}"))

    fIgnoreChar = fDoingSlash Or (iBracketLevel > 0)

    If fDoingSlash Then
      If sChar = " " _
        Or sChar = "{" Then

        fDoingSlash = False
        sRTFCodeToDo = sRTFCode
        sRTFCode = ""
      ElseIf sChar = "\" Then
        sRTFCodeToDo = sRTFCode
        sRTFCode = sChar
      Else
        sRTFCode = sRTFCode & Trim(Replace(Replace(sChar, vbCr, ""), vbLf, ""))
      End If
    End If

    If (iBracketLevel > 0) And sChar = "}" Then
      iBracketLevel = iBracketLevel - 1
      sRTFCodeToDo = sRTFCode
      sRTFCode = ""
    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
        fDoingSlash = True
        sRTFCode = sChar
      ElseIf sChar = "{" Then
        iBracketLevel = iBracketLevel + 1
        sRTFCodeToDo = sRTFCode
        sRTFCode = ""
      ElseIf Not fIgnoreChar Then
'''        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
'''      asText(iTextIndex) = asText(iTextIndex) & sNextChar
    End If

    ' See if we need to interpret the RTF control code.
    If Len(sRTFCodeToDo) > 0 Then
      If ((sRTFCodeToDo = "\ul") And (iTextIndex = 0)) _
        Or ((sRTFCodeToDo = "\ulnone") And (iTextIndex = 1)) Then
        iTextIndex = iTextIndex + 1
      ElseIf (sRTFCodeToDo = "\tab") Or (sRTFCodeToDo = "\cell") Then
        asText(iTextIndex) = asText(iTextIndex) & vbTab
      ElseIf (sRTFCodeToDo = "\row") Then
        asText(iTextIndex) = asText(iTextIndex) & vbNewLine
      ElseIf (Mid(sRTFCodeToDo, 1, 2) = "\'") Then
        fFound = False
        sDeniedChar = Chr(val("&H" & Mid(sRTFCodeToDo, 3)))
        For iLoop = 1 To UBound(asDeniedCharacters)
          If sDeniedChar = asDeniedCharacters(iLoop) Then
            fFound = True
            Exit For
          End If
        Next iLoop
        If Not fFound Then
          ReDim Preserve asDeniedCharacters(UBound(asDeniedCharacters) + 1)
          asDeniedCharacters(UBound(asDeniedCharacters)) = sDeniedChar
        End If
      End If

      sRTFCodeToDo = ""
    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
'''        fDoingSlash = True
'''        sRTFCode = sChar
      ElseIf sChar = "{" Then
'''        iBracketLevel = iBracketLevel + 1
'''        sRTFCodeToDo = sRTFCode
'''        sRTFCode = ""
      ElseIf Not fIgnoreChar Then
        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
      asText(iTextIndex) = asText(iTextIndex) & sNextChar
fDoingSlash = False
    End If

    ' Move forward through the text (jump an extra character if we've just processed a literal.
    If fLiteral Then
      sSourceText = Mid(sSourceText, 3)
    Else
      sSourceText = Mid(sSourceText, 2)
    End If
  Loop

  ' Trim off start CRLF
  If Len(asText(0)) >= 2 Then
    If Mid(asText(0), 1, 1) = vbCr _
      And Mid(asText(0), 2, 1) = vbLf Then

      asText(0) = Mid(asText(0), 3)
    End If
  End If

  ' Trim off end CRLF
  If Len(asText(iTextIndex)) >= 2 Then
    If Mid(asText(iTextIndex), Len(asText(iTextIndex)) - 1, 1) = vbCr _
      And Mid(asText(iTextIndex), Len(asText(iTextIndex)), 1) = vbLf Then

      asText(iTextIndex) = Mid(asText(iTextIndex), 1, Len(asText(iTextIndex)) - 2)
    End If
  End If

  psPart1 = asText(0)
  psPart2 = asText(1)
  psPart3 = asText(2)

End Sub



Private Sub cmdCancel_Click()
  If Changed Then
    Select Case MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title)
      Case vbYes
        cmdOK_Click
        Exit Sub
      Case vbCancel
        Exit Sub
    End Select
  End If

  Cancelled = True
  Me.Hide

End Sub

Private Sub cmdOK_Click()
  Dim fOK As Boolean

  fOK = True

  If Changed Then
    fOK = ValidateText
  End If

  If fOK Then
    ' Flag that the change/deletion has been confirmed.
    mfCancelled = False

    Me.Hide
  End If

End Sub

Private Sub cmdSelectText_Click()
  Dim iStart As Integer
  Dim iLength As Integer
  
  With txtMessage
    If .SelLength > 0 Then
      iStart = .SelStart
      iLength = .SelLength
  
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelUnderline = False
  
      .SelStart = iStart
      .SelLength = iLength
      .SelUnderline = True
      
      miSelStart = iStart
      miSelLength = iLength
    End If
  End With
  
  Changed = True
  
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
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    Cancel = True
  End If

End Sub


Private Sub Form_Resize()
  ResizeForm

End Sub


Private Sub ResizeForm()
  On Error GoTo ErrorTrap
  
  Dim sngOffset As Single
  
  With cmdCancel
    .Top = Me.ScaleHeight - BUTTON_OFFSET - .Height
    .Left = Me.ScaleWidth - .Width - cmdSelectText.Left
  End With
  
  With cmdOK
    .Top = cmdCancel.Top
    .Left = cmdCancel.Left - .Width - 200
  End With
  
  With cmdSelectText
    .Top = cmdCancel.Top
  End With
  
  With fraMessage
    .Width = Me.ScaleWidth - (2 * .Left)
    .Height = cmdCancel.Top - BUTTON_OFFSET - .Top

    txtMessage.Height = .Height - OFFSET_TOP - OFFSET_BOTTOM
  End With
  
  With txtMessage
    .Width = fraMessage.Width - (2 * .Left)
  End With
  
ErrorTrap:

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub txtMessage_Change()
  ParseMessage txtMessage
  Changed = True
  
End Sub



Public Property Get RichText() As String
  Dim sPart1 As String
  Dim sPart2 As String
  Dim sPart3 As String
  Dim sWholeThing As String
  
  DeconstructMessage txtMessage.TextRTF, _
    sPart1, _
    sPart2, _
    sPart3

  sWholeThing = _
    Replace(Replace(Replace(sPart1, "\", "\\"), "{", "\{"), "}", "\}") & _
    "\ul " & _
    Replace(Replace(Replace(sPart2, "\", "\\"), "{", "\{"), "}", "\}") & _
    "\ulnone " & _
    Replace(Replace(Replace(sPart3, "\", "\\"), "{", "\{"), "}", "\}")

  RichText = sWholeThing
  
End Property

Public Property Let RichText(ByVal psNewValue As String)
  Dim sNewRTFText As String
  
  sNewRTFText = "{\rtf1 " & _
    psNewValue & _
    "\par }"

  txtMessage.TextRTF = sNewRTFText

End Property
