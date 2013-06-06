VERSION 5.00
Begin VB.UserControl COA_WorkingPattern 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   LockControls    =   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   1875
   Begin VB.Frame fraWorkingPattern 
      Height          =   930
      Left            =   0
      TabIndex        =   14
      Top             =   -90
      Width           =   1875
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   0
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   1
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   4
         Left            =   585
         TabIndex        =   3
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   3
         Left            =   585
         TabIndex        =   2
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   6
         Left            =   780
         TabIndex        =   5
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   5
         Left            =   780
         TabIndex        =   4
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   8
         Left            =   975
         TabIndex        =   7
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   7
         Left            =   975
         TabIndex        =   6
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   10
         Left            =   1170
         TabIndex        =   9
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   9
         Left            =   1170
         TabIndex        =   8
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   12
         Left            =   1365
         TabIndex        =   11
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   11
         Left            =   1365
         TabIndex        =   10
         Top             =   390
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   14
         Left            =   1560
         TabIndex        =   13
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkDay 
         Height          =   195
         Index           =   13
         Left            =   1560
         TabIndex        =   12
         Top             =   390
         Width           =   195
      End
      Begin VB.Line linFocusLine 
         BorderColor     =   &H80000002&
         Visible         =   0   'False
         X1              =   390
         X2              =   585
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label lblToggle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   75
         TabIndex        =   25
         Top             =   165
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHidden 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblAM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblPM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   585
         Width           =   210
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   21
         Top             =   150
         Width           =   135
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   20
         Top             =   150
         Width           =   120
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   825
         TabIndex        =   19
         Top             =   150
         Width           =   150
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1020
         TabIndex        =   18
         Top             =   150
         Width           =   150
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1245
         TabIndex        =   17
         Top             =   150
         Width           =   105
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1425
         TabIndex        =   16
         Top             =   150
         Width           =   105
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1620
         TabIndex        =   15
         Top             =   150
         Width           =   105
      End
   End
End
Attribute VB_Name = "COA_WorkingPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Constant defining the size of the checkboxes

Const BOXDIMENSION = 195
Public fInResize As Boolean

Public Event Click()

Public Enum DayNumber
  iSUNDAY = 0
  iMONDAY = 1
  iTUESDAY = 2
  iWEDNESDAY = 3
  iTHURSDAY = 4
  iFRIDAY = 5
  iSATURDAY = 6
End Enum

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
    Dim Rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
             INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, _
          IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
              INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

'########################################################################
'# PROPERTIES                                                           #
'########################################################################

Public Property Get Backcolor() As OLE_COLOR

  ' Return the back colour of the control
  
  Backcolor = fraWorkingPattern.Backcolor
  
End Property

Public Property Let Backcolor(ByVal NewColor As OLE_COLOR)

  ' Set the back colour of the individual controls
  
  Dim iCounter As Integer
  
  fraWorkingPattern.Backcolor = NewColor
  
  For iCounter = 1 To 7
    lblDay(iCounter).Backcolor = NewColor
  Next iCounter
  
  lblToggle.Backcolor = NewColor
  lblAM.Backcolor = NewColor
  lblPM.Backcolor = NewColor
  
End Property


Public Property Get BorderStyle() As Integer

  ' Return the borderstyle
  
  BorderStyle = fraWorkingPattern.BorderStyle
  
End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As Integer)

  ' Set the borderstyle
  
  If NewBorderStyle = 0 Then
    fraWorkingPattern.BorderStyle = 0
  Else
    fraWorkingPattern.BorderStyle = 1
  End If
    
End Property


Public Property Get Enabled() As Boolean

  ' Return the controls enabled state
  
  Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal booNewEnabled As Boolean)

  ' Set the controls enabled state
  
  Dim objctl As Control
    
  UserControl.Enabled = booNewEnabled
  
  If Not booNewEnabled Then
    For Each objctl In UserControl.Controls
      If TypeOf objctl Is CheckBox Then
        objctl.Enabled = False
      End If
    Next
  Else
    For Each objctl In UserControl.Controls
      If TypeOf objctl Is CheckBox Then
        objctl.Enabled = True
      End If
    Next
  End If

End Property

Public Property Get hWnd() As Long
  
  hWnd = fraWorkingPattern.hWnd

End Property

Public Property Get Font() As Font

  ' Return the controls font
  
  Set Font = lblHidden.Font
  
End Property

Public Property Set Font(ByVal NewFont As Font)

  ' Set the controls font
  Dim objctl As Control
  Dim intCounter As Integer
  Dim intColumnWidth As Integer
  Dim intDayHeight As Integer
  
  Const iCOLUMNGAP = 0

  Set lblHidden.Font = NewFont
  
  ' Change the Label Font Sizes
  For Each objctl In lblDay
      Set objctl.Font = lblHidden.Font
  Next objctl

  Set lblAM.Font = lblHidden.Font
  Set lblPM.Font = lblHidden.Font
  
  ' Get the column width.
  intColumnWidth = lblHidden.Width
  intDayHeight = lblHidden.Height

  If intColumnWidth < BOXDIMENSION Then
    intColumnWidth = BOXDIMENSION
  End If

  For Each objctl In lblDay
    objctl.Width = intColumnWidth
    objctl.Height = intDayHeight
  Next objctl
  
  'Align the Day Labels
  For intCounter = 1 To 7
    If intCounter = 1 Then
      lblDay(intCounter).Left = Max(lblAM.Width, lblPM.Width) + iCOLUMNGAP + 100
    Else
      lblDay(intCounter).Left = lblDay(intCounter - 1).Left + lblDay(intCounter - 1).Width + iCOLUMNGAP
    End If
  Next intCounter
  
  'Align the AM/PM Labels
  lblAM.Top = lblDay(1).Top + lblDay(1).Height
  lblPM.Top = lblAM.Top + lblAM.Height
  
  'Vertically Align the Checkboxes
  For Each objctl In UserControl.Controls
    If TypeOf objctl Is CheckBox Then
      If objctl.Index Mod 2 <> 0 Then
        objctl.Top = lblAM.Top + ((lblAM.Height / 2) - (195 / 2))
      Else
        objctl.Top = lblPM.Top + ((lblPM.Height / 2) - (195 / 2))
      End If
    End If
  Next objctl
  
  'Horizontally Align the Checkboxes
  For Each objctl In UserControl.Controls
    If TypeOf objctl Is CheckBox Then
      Select Case objctl.Index
        Case 1, 2 ' sun
            objctl.Left = lblDay(1).Left + ((lblDay(1).Width - objctl.Width) / 2)
        Case 3, 4 ' mon
            objctl.Left = lblDay(2).Left + ((lblDay(2).Width - objctl.Width) / 2)
        Case 5, 6 ' tue
            objctl.Left = lblDay(3).Left + ((lblDay(3).Width - objctl.Width) / 2)
        Case 7, 8 ' wed
            objctl.Left = lblDay(4).Left + ((lblDay(4).Width - objctl.Width) / 2)
        Case 9, 10 ' thu
            objctl.Left = lblDay(5).Left + ((lblDay(5).Width - objctl.Width) / 2)
        Case 11, 12 ' fri
            objctl.Left = lblDay(6).Left + ((lblDay(6).Width - objctl.Width) / 2)
        Case 13, 14 ' sat
            objctl.Left = lblDay(7).Left + ((lblDay(7).Width - objctl.Width) / 2)
      End Select
    End If
  Next objctl

  fraWorkingPattern.Width = lblDay(7).Left + lblDay(7).Width + 100
  fraWorkingPattern.Height = (lblPM.Top + lblPM.Height) + 80

  UserControl.Width = fraWorkingPattern.Width
  UserControl.Height = fraWorkingPattern.Height - 90

End Property

Public Property Get Forecolor() As OLE_COLOR

  Forecolor = lblAM.Forecolor
  
End Property

Public Property Let Forecolor(ByVal NewColor As OLE_COLOR)

  Dim iCounter As Integer
  
  For iCounter = 1 To 7
    lblDay(iCounter).Forecolor = NewColor
  Next iCounter
  lblToggle.Forecolor = NewColor
  lblAM.Forecolor = NewColor
  lblPM.Forecolor = NewColor
  
End Property

Public Property Get Value() As String

  Value = CreateCharacterString
  
End Property

Public Property Let Value(ByVal NewString As String)

  Dim iCounter As Integer
  
  If Len(NewString) < 14 Then NewString = NewString & Space(14 - Len(NewString))
  
  For iCounter = 1 To 14
    If Mid(NewString, iCounter, 1) <> " " Then
      Select Case iCounter
        Case 1: chkDay(1).Value = vbChecked
        Case 2: chkDay(2).Value = vbChecked
        Case 3: chkDay(3).Value = vbChecked
        Case 4: chkDay(4).Value = vbChecked
        Case 5: chkDay(5).Value = vbChecked
        Case 6: chkDay(6).Value = vbChecked
        Case 7: chkDay(7).Value = vbChecked
        Case 8: chkDay(8).Value = vbChecked
        Case 9: chkDay(9).Value = vbChecked
        Case 10: chkDay(10).Value = vbChecked
        Case 11: chkDay(11).Value = vbChecked
        Case 12: chkDay(12).Value = vbChecked
        Case 13: chkDay(13).Value = vbChecked
        Case 14: chkDay(14).Value = vbChecked
      End Select
    Else
      Select Case iCounter
        Case 1: chkDay(1).Value = vbUnchecked
        Case 2: chkDay(2).Value = vbUnchecked
        Case 3: chkDay(3).Value = vbUnchecked
        Case 4: chkDay(4).Value = vbUnchecked
        Case 5: chkDay(5).Value = vbUnchecked
        Case 6: chkDay(6).Value = vbUnchecked
        Case 7: chkDay(7).Value = vbUnchecked
        Case 8: chkDay(8).Value = vbUnchecked
        Case 9: chkDay(9).Value = vbUnchecked
        Case 10: chkDay(10).Value = vbUnchecked
        Case 11: chkDay(11).Value = vbUnchecked
        Case 12: chkDay(12).Value = vbUnchecked
        Case 13: chkDay(13).Value = vbUnchecked
        Case 14: chkDay(14).Value = vbUnchecked
      End Select
    End If
  Next iCounter
  
End Property

'########################################################################
'# FUNCTIONS USED BY THE CONTROL WHICH ARE NOT EXPOSED TO ITS CONTAINER #
'########################################################################

Private Function Max(intFirst As Integer, intSecond As Integer) As Integer
  
  If intFirst > intSecond Then
    Max = intFirst
  Else
    Max = intSecond
  End If

End Function

Private Sub chkDay_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub chkDay_GotFocus(Index As Integer)

  linFocusLine.Visible = True
  
  Select Case Index
    Case 1, 3, 5, 7, 9, 11, 13
      linFocusLine.X1 = chkDay(Index).Left + 30
      linFocusLine.X2 = chkDay(Index).Left + BOXDIMENSION - 30
      linFocusLine.Y1 = chkDay(Index).Top - 15
      linFocusLine.Y2 = linFocusLine.Y1
    Case Else
      linFocusLine.X1 = chkDay(Index).Left + 30
      linFocusLine.X2 = chkDay(Index).Left + BOXDIMENSION - 30
      linFocusLine.Y1 = chkDay(Index).Top + BOXDIMENSION + 15
      linFocusLine.Y2 = linFocusLine.Y1
  End Select
    
End Sub

Private Sub chkDay_LostFocus(Index As Integer)

'  If Index = 14 Then
    linFocusLine.Visible = False
'  End If

End Sub

Private Sub lblToggle_Click()

  ' Toggle status of all checkboxes
  
'  Dim intCounter As Integer
'  For intCounter = 1 To 14
'    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
'  Next intCounter

End Sub

Private Sub lblAM_Click()

  ' Select all the AM checkboxes

  Dim intCounter As Integer
  For intCounter = 1 To 13 Step 2
    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
  Next intCounter

End Sub

Private Sub lblPM_Click()

  ' Select all the PM checkboxes

  Dim intCounter As Integer
  For intCounter = 2 To 14 Step 2
    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
  Next intCounter

End Sub

Private Sub lblDay_Click(Index As Integer)

  ' If any of the Day labels are clicked, check the related checkboxes
  
  Select Case Index
    Case 1 ' sun
        chkDay(1).Value = IIf(chkDay(1).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(2).Value = IIf(chkDay(2).Value = vbChecked, vbUnchecked, vbChecked)
    Case 2 ' mon
        chkDay(3).Value = IIf(chkDay(3).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(4).Value = IIf(chkDay(4).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 3
    Case 3 ' tue
        chkDay(5).Value = IIf(chkDay(5).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(6).Value = IIf(chkDay(6).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 5
    Case 4 ' wed
        chkDay(7).Value = IIf(chkDay(7).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(8).Value = IIf(chkDay(8).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 7
    Case 5 ' thu
        chkDay(9).Value = IIf(chkDay(9).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(10).Value = IIf(chkDay(10).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 9
    Case 6 ' fri
        chkDay(11).Value = IIf(chkDay(11).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(12).Value = IIf(chkDay(12).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 11
    Case 7 ' sat
        chkDay(13).Value = IIf(chkDay(13).Value = vbChecked, vbUnchecked, vbChecked)
        chkDay(14).Value = IIf(chkDay(14).Value = vbChecked, vbUnchecked, vbChecked)
        Index = 13
  End Select

  
      linFocusLine.X1 = chkDay(Index).Left + 30
      linFocusLine.X2 = chkDay(Index).Left + BOXDIMENSION - 30
      linFocusLine.Y1 = chkDay(Index).Top - 15
      linFocusLine.Y2 = linFocusLine.Y1
End Sub

Private Function CreateCharacterString() As String

  Dim intCounter As Integer
  
  For intCounter = 1 To 14
    If chkDay(intCounter).Value = vbChecked Then
      Select Case intCounter
        Case 1: CreateCharacterString = CreateCharacterString & Left(WeekdayName(7), 1)
        Case 2: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(7), 1))
        Case 3: CreateCharacterString = CreateCharacterString & Left(WeekdayName(1), 1)
        Case 4: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(1), 1))
        Case 5: CreateCharacterString = CreateCharacterString & Left(WeekdayName(2), 1)
        Case 6: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(2), 1))
        Case 7: CreateCharacterString = CreateCharacterString & Left(WeekdayName(3), 1)
        Case 8: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(3), 1))
        Case 9: CreateCharacterString = CreateCharacterString & Left(WeekdayName(4), 1)
        Case 10: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(4), 1))
        Case 11: CreateCharacterString = CreateCharacterString & Left(WeekdayName(5), 1)
        Case 12: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(5), 1))
        Case 13: CreateCharacterString = CreateCharacterString & Left(WeekdayName(6), 1)
        Case 14: CreateCharacterString = CreateCharacterString & UCase(Left(WeekdayName(6), 1))
'        Case 1: CreateCharacterString = CreateCharacterString & "s"
'        Case 2: CreateCharacterString = CreateCharacterString & "S"
'        Case 3: CreateCharacterString = CreateCharacterString & "m"
'        Case 4: CreateCharacterString = CreateCharacterString & "M"
'        Case 5: CreateCharacterString = CreateCharacterString & "t"
'        Case 6: CreateCharacterString = CreateCharacterString & "T"
'        Case 7: CreateCharacterString = CreateCharacterString & "w"
'        Case 8: CreateCharacterString = CreateCharacterString & "W"
'        Case 9: CreateCharacterString = CreateCharacterString & "t"
'        Case 10: CreateCharacterString = CreateCharacterString & "T"
'        Case 11: CreateCharacterString = CreateCharacterString & "f"
'        Case 12: CreateCharacterString = CreateCharacterString & "F"
'        Case 13: CreateCharacterString = CreateCharacterString & "s"
'        Case 14: CreateCharacterString = CreateCharacterString & "S"
      
      End Select
    Else
      CreateCharacterString = CreateCharacterString & " "
    End If
  Next intCounter

If Len(CreateCharacterString) < 14 Then
  CreateCharacterString = CreateCharacterString & Space(14 - Len(CreateCharacterString))
End If

End Function

Private Sub UserControl_Resize()
  
  On Error Resume Next
  
  If Not fInResize Then
    fInResize = True
    
'    lblDay(1) = UCase(Left(WeekdayName(7), 1))
'    lblDay(2) = UCase(Left(WeekdayName(1), 1))
'    lblDay(3) = UCase(Left(WeekdayName(2), 1))
'    lblDay(4) = UCase(Left(WeekdayName(3), 1))
'    lblDay(5) = UCase(Left(WeekdayName(4), 1))
'    lblDay(6) = UCase(Left(WeekdayName(5), 1))
'    lblDay(7) = UCase(Left(WeekdayName(6), 1))
    
    lblDay(1) = UCase(Left(WeekdayName(1, , vbSunday), 1))
    lblDay(2) = UCase(Left(WeekdayName(2, , vbSunday), 1))
    lblDay(3) = UCase(Left(WeekdayName(3, , vbSunday), 1))
    lblDay(4) = UCase(Left(WeekdayName(4, , vbSunday), 1))
    lblDay(5) = UCase(Left(WeekdayName(5, , vbSunday), 1))
    lblDay(6) = UCase(Left(WeekdayName(6, , vbSunday), 1))
    lblDay(7) = UCase(Left(WeekdayName(7, , vbSunday), 1))
    
    fraWorkingPattern.Width = lblDay(7).Left + lblDay(7).Width + 100
    fraWorkingPattern.Height = (lblPM.Top + lblPM.Height) + 80
    UserControl.Width = fraWorkingPattern.Width
    UserControl.Height = fraWorkingPattern.Height - 90
    fInResize = False
  End If
  
End Sub



Public Property Let StartDay(ByVal pintDayNumber As DayNumber)

  Dim iLeftLabels(7) As Integer
  Dim iLeftCheckboxes(14) As Integer
  Dim iCount As Integer
  
  ' Store original left positions
  For iCount = 1 To 7
    iLeftLabels(iCount) = lblDay(iCount).Left
  Next iCount
  
  For iCount = 1 To 14
    iLeftCheckboxes(iCount) = chkDay(iCount).Left
  Next iCount

  ' Labels
  For iCount = pintDayNumber To 1 Step -1
    lblDay(iCount).Left = iLeftLabels((7 - pintDayNumber) + iCount)
  Next iCount

  For iCount = 7 To pintDayNumber + 1 Step -1
    lblDay(iCount).Left = iLeftLabels(iCount - pintDayNumber)
  Next iCount

  'Checkboxes
  For iCount = (pintDayNumber * 2) To 1 Step -1
    chkDay(iCount).Left = iLeftCheckboxes((14 - (pintDayNumber * 2)) + iCount)
  Next iCount

  For iCount = 14 To (pintDayNumber * 2) + 1 Step -1
    chkDay(iCount).Left = iLeftCheckboxes(iCount - (pintDayNumber * 2))
  Next iCount

End Property
