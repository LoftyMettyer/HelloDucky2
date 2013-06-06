VERSION 5.00
Begin VB.Form frmColorPalette 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorPalette.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmColorPalette.frx":030A
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'API function & constant declarations
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As udtCHOOSECOLOR) As Long
Private Type udtCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const CC_FULLOPEN = &H2
Private Const CC_ANYCOLOR = &H100

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_HIDE = 0

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Module specific variable declarations
Private Type cpColorInformation
    Clr As OLE_COLOR
    Rct As RECT
    Tip As String
    Caption As String
End Type

Private Clrs(62) As cpColorInformation

Private IsSystemColors As Boolean
Private MouseButId As Integer
Private MouseDownButId As Integer
Private CurClrButId As Integer

Private Const NorClrVal = "&HFFFFFF&HC0C0FF&HC0E0FF&HC0FFFF&HC0FFC0&HFFFFC0&HFFC0C0&HFFC0FF" & _
                          "&HE0E0E0&H8080FF&H80C0FF&H80FFFF&H80FF80&HFFFF80&HFF8080&HFF80FF" & _
                          "&HC0C0C0&H0000FF&H0080FF&H00FFFF&H00FF00&HFFFF00&HFF0000&HFF00FF" & _
                          "&H808080&H0000C0&H0040C0&H00C0C0&H00C000&HC0C000&HC00000&HC000C0" & _
                          "&H404040&H000080&H004080&H008080&H008000&H808000&H800000&H800080" & _
                          "&H000000&H000040&H404080&H004040&H004000&H404000&H400000&H400040"
Private Const SysClrVal = "&H80000000&H80000001&H80000002&H80000003&H80000004&H80000005" & _
                          "&H80000006&H80000007&H80000008&H80000009&H8000000A&H8000000B" & _
                          "&H8000000C&H8000000D&H8000000E&H8000000F&H80000010&H80000011" & _
                          "&H80000012&H80000013&H80000014&H80000015&H80000016&H80000017" & _
                          "&H80000018"
Private Const NorClrTip = ""
Private Const SysClrTip = "Scroll Bars            " & _
                          "Desktop                " & _
                          "Active Title Bar       " & _
                          "Inactive Title Bar     " & _
                          "Menu Bar               " & _
                          "Window Background      " & _
                          "Window Frame           " & _
                          "Menu Text              " & _
                          "Window Text            " & _
                          "Active Title Bar Text  " & _
                          "Active Border          " & _
                          "Inactive Border        " & _
                          "Application Workspace  " & _
                          "Highlight              " & _
                          "Highlight Text         " & _
                          "Button Face            " & _
                          "Button Shadow          " & _
                          "Disabled Text          " & _
                          "Button Text            " & _
                          "Inactive Title Bar Text" & _
                          "Button Highlight       " & _
                          "Button Dark Shadow     " & _
                          "Button Light Shadow    " & _
                          "ToolTip Text           " & _
                          "ToolTip                "
Private Const OtherTip = "Normal Colours    " & _
                         "System Colours    " & _
                         "Show Colour Dialog"

Private pl As Long, Pt As Long

Private Const TipTmr1 = 1
Private Const TipTmr2 = 2
Private IsTmr1Active As Boolean
Private IsTmr2Active As Boolean
Private TipButId As Integer
Private mSelectedColor As OLE_COLOR

Public IsCanceled As Boolean

Private Sub cmdButton_Click(Index As Integer)
  Call DoAction(Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then
'        Call ReleaseCapture
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    
    Dim R As RECT
    
    Me.ScaleMode = vbPixels
    Me.Font.Name = "Verdana"
    
'    Call SetCapture(hwnd)
    
    IsSystemColors = False
    MouseButId = -1
    MouseDownButId = -1
    IsCanceled = True
    
    Call Initialize
       
    Width = (pl + (8 * 16) + 7 + 4) * Screen.TwipsPerPixelX
    Height = (Pt + 4) * Screen.TwipsPerPixelY
    
    Call SetRect(R, 0, 0, ScaleWidth, ScaleHeight)
    Call DrawEdge(hdc, R, BDR_RAISEDINNER, BF_RECT)
    
    Load frmTip

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbLeftButton) Then Exit Sub

    If Not (MouseButId = -1) Then
        If (MouseButId = 58) Or (MouseButId = 59) Or (MouseButId = 60) _
          Or (MouseButId = 61) Or (MouseButId = 62) Then
            Call DrawButton(MouseButId, 1)
        End If

        'Select New
        Call DrawButEdge(MouseButId, 2)

        'Deselect Original
        If (Not MouseDownButId = -1) Then
          Select Case MouseButId
          Case 58 To 62
            Call DrawButEdge(MouseDownButId, 1)
          Case Else
            Call DrawButEdge(MouseDownButId, 0)
          End Select
        End If

        MouseDownButId = MouseButId

        Call ShowTip(False)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim IsMouseOnBut As Boolean
    
'    If Not (MouseButId = -1) Then
'        Exit Sub
'    End If

    For i = 1 To 62
        IsMouseOnBut = (X >= Clrs(i).Rct.Left And Y >= Clrs(i).Rct.Top) And (X <= Clrs(i).Rct.Right And Y <= Clrs(i).Rct.Bottom)
        If IsMouseOnBut Then
            Exit For
        End If
    Next i

    If ((i = MouseDownButId) And IsMouseOnBut) Then
      Exit Sub
    End If
    
    If (Not MouseButId = -1) And (Not MouseButId = i) Then
        Select Case MouseButId
        Case 58 To 62
          Call DrawButEdge(MouseButId, 1)
        Case Else
          Call DrawButEdge(MouseButId, 0)
        End Select
        MouseButId = -1
        Call ShowTip(False)
    End If

    If IsMouseOnBut And (Not MouseButId = i) Then
      MouseButId = i
      
      Call DrawButEdge(MouseButId, 1)
      
      If ShwTip Then
          Call SetTimer(Me.hwnd, CLng(TipTmr1), 500, AddressOf Timer)
          IsTmr1Active = True
      End If
    End If
    
    If Not IsMouseOnBut Then
        If IsTmr1Active Then
            Call KillTimer(Me.hwnd, CLng(TipTmr1))
            IsTmr1Active = False
        End If
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim IsMouseOver As Boolean
    
    If Not (MouseDownButId = -1) Then
        If (MouseDownButId = 58) Or (MouseDownButId = 59) Or (MouseDownButId = 60) _
          Or (MouseDownButId = 61) Or (MouseDownButId = 62) Then
            Call DrawButton(MouseDownButId, 0)
            
        End If
        
        If IsMouseOnBut(MouseDownButId) Then
            Call DoAction(MouseDownButId)
        End If
        
        MouseButId = -1
    End If
    
'    IsMouseOver = X >= 0 And Y >= 0 And X <= ScaleWidth And Y <= ScaleHeight
'    If IsMouseOver Then
'        Call SetCapture(Me.hwnd)
'    End If
End Sub

Private Sub DrawButEdge(ClrId As Integer, EdgeStyle As Integer)
        
    Select Case EdgeStyle
        Case 0: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
        Case 1: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT)
        Case 2: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_SUNKENOUTER, BF_RECT)
    End Select
    
    Refresh
End Sub

Private Sub Initialize()
    Dim i As Integer
    Dim LPos As Long, TPos As Long
    Dim FrmBkClr As Long
    
    pl = 4: Pt = 0
    
    If ShwDef Then
        Call SetRect(Clrs(1).Rct, pl, (Pt + 4), pl + 7 + 16 * 8, (Pt + 4) + 22)
        Pt = (Pt + 4) + 22
    End If
    
    For i = 2 To 49
        LPos = (((i - 2) Mod 8) + pl) + (((i - 2) Mod 8) * 16)
        TPos = (Int((i - 2) / 8) + (Pt + 4)) + (Int((i - 2) / 8) * 16)
        Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 16, TPos + 16)
    Next i
    Pt = (Pt + 4) + (6 * 16) + 5

    If ShwCus Then
        FrmBkClr = Me.ForeColor
        Me.ForeColor = vb3DShadow
        CurrentX = 4: CurrentY = Pt + 2
        Line -(16 * 8 + 4 + 7, CurrentY)
        Me.ForeColor = vb3DHighlight
        CurrentX = 4: CurrentY = Pt + 2 + 1
        Line -(16 * 8 + 4 + 7, CurrentY)
        Me.ForeColor = FrmBkClr
        
        Pt = Pt + 2 + 1
        
        For i = 50 To 57
            LPos = (((i - 50) Mod 8) + 4) + (((i - 50) Mod 8) * 16)
            TPos = (Int((i - 50) / 8) + (Pt + 2)) + (Int((i - 50) / 8) * 16)
            Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 16, TPos + 16)
        Next i
        
        Pt = (Pt + 2) + 16
    End If
    
    If ShwMor Or ShwSys Then
        FrmBkClr = Me.ForeColor
        Me.ForeColor = vb3DShadow
        CurrentX = 4: CurrentY = Pt + 2
        Line -(16 * 8 + 4 + 7, CurrentY)
        Me.ForeColor = vb3DHighlight
        CurrentX = 4: CurrentY = Pt + 2 + 1
        Line -(16 * 8 + 4 + 7, CurrentY)
        Me.ForeColor = FrmBkClr
        
        Pt = Pt + 2 + 1
    End If

    If ShwSys Then
        For i = 58 To 59
            LPos = (((i - 58) Mod 2) * 7 + pl) + (((i - 58) Mod 2) * 64)
            TPos = (Int((i - 58) / 2) + (Pt + 2)) + (Int((i - 58) / 2) * 20)
            Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 64, TPos + 20)
        Next i
        
        Pt = (Pt + 2) + 20
    End If
    
    If ShwMor Then
        Call SetRect(Clrs(60).Rct, pl, (Pt + 2), 4 + 7 + 16 * 8, (Pt + 2) + 20)
        Pt = (Pt + 2) + 20
    End If
    
    ' OK / Cancel buttons
    For i = 61 To 62
      LPos = (((i - 61) Mod 2) * 7 + pl) + (((i - 61) Mod 2) * 64)
      TPos = (Int((i - 61) / 2) + (Pt + 2)) + (Int((i - 61) / 2) * 20)
      Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 64, TPos + 20)
      
      ' Add a new CommandButton
      Load cmdButton(i)
      cmdButton(i).Caption = Clrs(i).Caption
      cmdButton(i).Top = Clrs(i).Rct.Top
      cmdButton(i).Left = Clrs(i).Rct.Left
      cmdButton(i).Height = Clrs(i).Rct.Bottom - Clrs(i).Rct.Top
      cmdButton(i).Width = Clrs(i).Rct.Right - Clrs(i).Rct.Left
      cmdButton(i).Visible = True
    Next i
    
    Pt = (Pt + 2) + 20

    For i = 1 To 62
        Call DrawButton(i, 0)
    Next i

End Sub

Private Sub DrawButton(ButId As Integer, State As Integer)
    Dim Clr As Long, Brsh As Long
    Dim R As RECT
    
    Call OleTranslateColor(Me.BackColor, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hdc, Clrs(ButId).Rct, Brsh)
    Call DeleteObject(Clr)
    Call DeleteObject(Brsh)
    
    Select Case ButId
    Case 1
      ' Default color
      If Not ShwDef Then Exit Sub
      
      Clrs(1).Clr = DefClr
      Clrs(1).Tip = "Default"
      
      Call SetRect(R, Clrs(1).Rct.Left + 3, Clrs(1).Rct.Top + 3, Clrs(1).Rct.Right - 3, Clrs(1).Rct.Bottom - 3)
      Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
      Call SetRect(R, Clrs(1).Rct.Left + 5, Clrs(1).Rct.Top + 5, Clrs(1).Rct.Left + 5 + 12, Clrs(1).Rct.Top + 5 + 12)
      Call OleTranslateColor(Clrs(1).Clr, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
      Call SetRect(R, Clrs(1).Rct.Left + 5 + 12, Clrs(1).Rct.Top + 3, Clrs(1).Rct.Right - 2, Clrs(1).Rct.Bottom - 3)
      Call DrawText(hdc, DefCap, Len(DefCap), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)

    Case 2 To 49
      ' System colors
      If Not IsSystemColors Then
          Clrs(ButId).Clr = CLng(Mid(NorClrVal, (ButId - 2) * 8 + 1, 8))
          Clrs(ButId).Tip = ""
      Else
          If (ButId <= 26) Then
              Clrs(ButId).Clr = CLng(Mid(SysClrVal, (ButId - 2) * 10 + 1, 10))
              Clrs(ButId).Tip = Trim(Mid(SysClrTip, (ButId - 2) * 23 + 1, 23))
          Else
              Clrs(ButId).Clr = &HFFFFFF
              Clrs(ButId).Tip = ""
          End If
      End If
      
      Call SetRect(R, Clrs(ButId).Rct.Left + 2, Clrs(ButId).Rct.Top + 2, Clrs(ButId).Rct.Right - 2, Clrs(ButId).Rct.Bottom - 2)
      Call OleTranslateColor(Clrs(ButId).Clr, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
      Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
    Case 50 To 57
      ' Custom colors
      If Not ShwCus Then Exit Sub
      
      Clrs(ButId).Clr = &HFFFFFF
      Clrs(ButId).Tip = "Custom Color " & Trim(Str(ButId - 49))
      
      If Not (LastSavedCustClr = 0) Then
          If (UBound(CustClrs) >= (ButId - 49)) Then
              Clrs(ButId).Clr = CustClrs(ButId - 49)
          End If
      End If
      
      Call OleTranslateColor(Clrs(ButId).Clr, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call SetRect(R, Clrs(ButId).Rct.Left + 2, Clrs(ButId).Rct.Top + 2, Clrs(ButId).Rct.Right - 2, Clrs(ButId).Rct.Bottom - 2)
      Call FillRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
      Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hdc, R, Brsh)
      Call DeleteObject(Brsh)
      Call DeleteObject(Clr)
      
    Case 58 To 62
      ' Option buttons
      Dim TmpStr As String
      Select Case ButId
          Case 58: TmpStr = "Normal": If Not ShwSys Then Exit Sub
          Case 59: TmpStr = "System": If Not ShwSys Then Exit Sub
          Case 60: TmpStr = MorCap: If Not ShwMor Then Exit Sub
          Case 61: cmdButton(ButId).Caption = "OK": If Not ShwSys Then Exit Sub
          Case 62: cmdButton(ButId).Caption = "Cancel": If Not ShwSys Then Exit Sub
      End Select
      
      If State = 0 Then
          Call SetRect(R, Clrs(ButId).Rct.Left, Clrs(ButId).Rct.Top, Clrs(ButId).Rct.Right, Clrs(ButId).Rct.Bottom)
          Call DrawButEdge(ButId, 1)
      Else
          Call SetRect(R, Clrs(ButId).Rct.Left + 1, Clrs(ButId).Rct.Top + 1, Clrs(ButId).Rct.Right + 1, Clrs(ButId).Rct.Bottom + 1)
      End If

      Call DrawText(hdc, TmpStr, CLng(Len(TmpStr)), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      Clrs(ButId).Tip = Trim(Mid(OtherTip, (ButId - 58) * 17 + 1, 17))
      
    End Select
    
    Refresh
End Sub

Private Sub SelectActiveColorButton()
Dim h As Integer
Dim i As Integer
Dim iMax As Integer
Dim bFound As Boolean

  If (SelectedColor = 0) Then Exit Sub
  
  If (Not ShwSys) Then
    iMax = 1
  Else
    iMax = 2
  End If
  
  For h = 1 To iMax
    For i = 2 To 49
      If Clrs(i).Clr = SelectedColor Then
        Call DrawButEdge(i, 2)
        bFound = True
        Exit For
      End If
    Next i
    
    If (Not bFound) And iMax > 1 Then
      DoAction (59)
    End If
  Next h
  
  MouseDownButId = i
  
  Refresh
End Sub

Private Sub DoAction(ButId As Integer)
    Dim i As Integer
    
    Select Case ButId
        Case 1 To 57
            SelectedColor = Clrs(ButId).Clr
            IsCanceled = False
        Case 58
            If IsSystemColors Then
                IsSystemColors = False
                For i = 2 To 49
                    Call DrawButton(i, 0)
                Next i
            End If
        Case 59
            If Not IsSystemColors Then
                IsSystemColors = True
                For i = 2 To 49
                    Call DrawButton(i, 0)
                Next i
            End If
        Case 60
            SelectedColor = ShowColor
            If Not SelectedColor = -1 Then
                Call SaveCustClr(SelectedColor)
                IsCanceled = False
            Else
                IsCanceled = True
            End If
            Call Form_KeyDown(vbKeyEscape, 0)
        Case 61
          ' OK
          IsCanceled = False
          Call Form_KeyDown(vbKeyEscape, 0)
        Case 62
          ' Cancel
          IsCanceled = True
          Call Form_KeyDown(vbKeyEscape, 0)
    End Select
End Sub

Private Function IsMouseOnBut(ButId As Integer) As Boolean
    Dim Pt As POINTAPI
    
    Call GetCursorPos(Pt)
    Call ScreenToClient(Me.hwnd, Pt)
    IsMouseOnBut = (Pt.X >= Clrs(ButId).Rct.Left And Pt.X <= Clrs(ButId).Rct.Right) And _
                   (Pt.Y >= Clrs(ButId).Rct.Top And Pt.Y <= Clrs(ButId).Rct.Bottom)
End Function

Private Function ShowColor() As Long
    Dim ClrInf As udtCHOOSECOLOR
    Static CustomColors(64) As Byte
    Dim i As Integer
    
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
    
    With ClrInf
        .lStructSize = Len(ClrInf)              'Size of the structure
        .hwndOwner = Me.hwnd                    'Handle of owner window
        .hInstance = App.hInstance              'Instance of application
        .lpCustColors = StrConv(CustomColors, vbUnicode)       'Array of 16 byte values
        .flags = CC_FULLOPEN                    'Flags to open in full mode
    End With
    
    If Not ChooseColor(ClrInf) = 0 Then
        ShowColor = ClrInf.rgbResult
    Else
        ShowColor = -1
    End If
End Function

Private Sub SaveCustClr(ClrVal As OLE_COLOR)
    If (LastSavedCustClr = 0) Then
        ReDim Preserve CustClrs(1) As OLE_COLOR
    Else
        If (UBound(CustClrs) < 8) Then
            ReDim Preserve CustClrs(UBound(CustClrs) + 1) As OLE_COLOR
        End If
    End If
    
    LastSavedCustClr = LastSavedCustClr + 1
    If (LastSavedCustClr > 8) Then LastSavedCustClr = 1
    
    CustClrs(LastSavedCustClr) = ClrVal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = 1 To 60
        Call SetRectEmpty(Clrs(i).Rct)
    Next i
    
    If IsTmr1Active Then
        Call KillTimer(Me.hwnd, CLng(TipTmr1))
        IsTmr1Active = False
    End If
    
    If IsTmr2Active Then
        Call KillTimer(Me.hwnd, CLng(TipTmr2))
        IsTmr2Active = False
    End If

    Unload frmTip
End Sub

Public Sub TipTimer(hwnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)
    Select Case idEvent
        Case 1
            Call ShowTip(True)
            
            Call KillTimer(Me.hwnd, CLng(TipTmr1))
            IsTmr1Active = False
        Case 2
            Call ShowTip(False)
    End Select
End Sub

Private Sub ShowTip(State As Boolean)
    If State Then
        Dim Rct As RECT
        Dim Pt As POINTAPI
        Dim TipTxt As String
        
        'Store the tip text in a variable
        TipTxt = Clrs(MouseButId).Tip
        If TipTxt = "" Then Exit Sub
        
        'Clear Tip Form
        frmTip.Cls
        
        'Draw Tip text and position the Tip Form
        Call GetCursorPos(Pt)
        Call SetRect(Rct, 0, 0, frmTip.ScaleWidth, frmTip.ScaleHeight)
        Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_CALCRECT)
        Call SetRect(Rct, 0, 0, Rct.Right + 8, Rct.Bottom + 6)
        Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
        Call DrawEdge(frmTip.hdc, Rct, BDR_RAISEDINNER, BF_RECT)
        
        If (Pt.X + 2) * Screen.TwipsPerPixelX + Rct.Right * Screen.TwipsPerPixelX > Screen.Width Then
          frmTip.Move Screen.Width - Rct.Right * Screen.TwipsPerPixelX, (Pt.Y + 20) * Screen.TwipsPerPixelY, _
                      Rct.Right * Screen.TwipsPerPixelX, Rct.Bottom * Screen.TwipsPerPixelY
        Else
          frmTip.Move (Pt.X + 2) * Screen.TwipsPerPixelX, (Pt.Y + 20) * Screen.TwipsPerPixelY, _
            Rct.Right * Screen.TwipsPerPixelX, Rct.Bottom * Screen.TwipsPerPixelY
        End If
                
        frmTip.ZOrder
        frmTip.Refresh
        Call ShowWindow(frmTip.hwnd, SW_SHOWNOACTIVATE)
        
        'Set Timer 2 for the duration of tip
        Call SetTimer(Me.hwnd, CLng(TipTmr2), 4000, AddressOf Timer)
        IsTmr2Active = True
    Else
        On Error Resume Next
        
        'Hide Tip Form
        Call ShowWindow(frmTip.hwnd, SW_HIDE)
        
        'Kill Timer 2 if it is active
        If IsTmr2Active Then
            Call KillTimer(Me.hwnd, CLng(TipTmr2))
            IsTmr2Active = False
        End If
        
        'Kill Timer 1 if it is active
        If IsTmr1Active Then
            Call KillTimer(Me.hwnd, CLng(TipTmr1))
            IsTmr1Active = False
        End If
    End If
End Sub

Public Property Get SelectedColor() As OLE_COLOR
  SelectedColor = mSelectedColor
End Property

Public Property Let SelectedColor(ByVal vNewValue As OLE_COLOR)
  mSelectedColor = vNewValue
  Call SelectActiveColorButton
End Property
