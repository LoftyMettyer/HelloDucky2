VERSION 5.00
Begin VB.UserControl COASD_WorkingPattern 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   LockControls    =   -1  'True
   ScaleHeight     =   855
   ScaleWidth      =   1890
   Begin VB.Frame fraWorkingPattern 
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   1875
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   14
         Left            =   1560
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   25
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   1365
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   24
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   1170
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   23
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   975
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   22
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   780
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   21
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   20
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   390
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   19
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   1560
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   18
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   11
         Left            =   1365
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   17
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   1170
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   16
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   975
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   15
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   780
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   14
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   13
         Top             =   390
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   390
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   12
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   390
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":0000
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   390
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":024A
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   585
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":0494
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   780
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":06DE
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   975
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":0928
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   10
         Left            =   1170
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":0B72
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   12
         Left            =   1365
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":0DBC
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   14
         Left            =   1560
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":1006
         Top             =   570
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   585
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":1250
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   780
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":149A
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   975
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":16E4
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   9
         Left            =   1170
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":192E
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   11
         Left            =   1365
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":1B78
         Top             =   390
         Width           =   195
      End
      Begin VB.Image imgCheckBox 
         Enabled         =   0   'False
         Height          =   195
         Index           =   13
         Left            =   1560
         MousePointer    =   1  'Arrow
         Picture         =   "COASD_WorkingPattern.ctx":1DC2
         Top             =   390
         Width           =   195
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
         Left            =   1605
         TabIndex        =   11
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
         Left            =   1410
         TabIndex        =   10
         Top             =   150
         Width           =   105
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
         Left            =   1215
         TabIndex        =   9
         Top             =   150
         Width           =   105
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
         TabIndex        =   8
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
         Index           =   3
         Left            =   795
         TabIndex        =   7
         Top             =   150
         Width           =   150
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
         TabIndex        =   6
         Top             =   150
         Width           =   115
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
         Left            =   420
         TabIndex        =   5
         Top             =   150
         Width           =   135
      End
      Begin VB.Label lblPM 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   585
         Width           =   210
      End
      Begin VB.Label lblAM 
         AutoSize        =   -1  'True
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
         TabIndex        =   3
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblHidden 
         AutoSize        =   -1  'True
         Caption         =   "W"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblToggle 
         AutoSize        =   -1  'True
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
         TabIndex        =   1
         Top             =   165
         Visible         =   0   'False
         Width           =   150
      End
   End
End
Attribute VB_Name = "COASD_WorkingPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

' Constants
Const BOXDIMENSION = 195
Const gLngMinHeight = 855
Const gLngMinWidth = 1890

Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private mblnReadOnly As Boolean   'NPG20071022

Private mBackcolour As OLE_COLOR


'########################################################################
'# PROPERTIES                                                           #
'########################################################################

Public Property Get Selected() As Boolean
  Selected = gfSelected
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  gfSelected = pfNewValue
End Property

Public Property Get Backcolor() As OLE_COLOR
  Backcolor = mBackcolour
End Property

Public Property Let Backcolor(ByVal NewColor As OLE_COLOR)

  ' Set the back colour of the individual controls
  
  Dim iCounter As Integer
  
  mBackcolour = NewColor
  
  If Not mblnReadOnly Then
    
    fraWorkingPattern.Backcolor = mBackcolour
    
    For iCounter = 1 To 7
      lblDay(iCounter).Backcolor = mBackcolour
    Next iCounter
    
    lblToggle.Backcolor = mBackcolour
    lblAM.Backcolor = mBackcolour
    lblPM.Backcolor = mBackcolour
    
  End If
  
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

Public Property Get ControlLevel() As Integer
  
  ' Return the control's level in the z-order.
  ControlLevel = giControlLevel
  
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  
  ' Set the control's level in the z-order.
  giControlLevel = piNewValue
  
End Property


Public Property Get ColumnID() As Long
  
  ' Return the control's column ID.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(ByVal pLngNewValue As Long)
  
  ' Set the control's column ID.
  gLngColumnID = pLngNewValue
  
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

Public Property Get Font() As Font

  ' Return the controls font
  
  Set Font = lblDay(1).Font
  
End Property

Public Property Set Font(ByVal pObjNewValue As StdFont)
  ' Set the control's font property.
  Dim objctl As Control
  Dim intCounter As Integer
  Dim intColumnWidth As Integer
  Dim intDayHeight As Integer
  
  Const iCOLUMNGAP = 0
  
  ' Update the sub-controls.
  Set UserControl.Font = pObjNewValue
  
  ' Change the Label Font Sizes
  For Each objctl In UserControl.Controls
    If TypeOf objctl Is Label Then
      Set objctl.Font = pObjNewValue
    End If
  Next objctl

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
  
  'Vertically Align the Checkbox Images
  For Each objctl In UserControl.Controls
    
    If TypeOf objctl Is Image Then
      If CInt(objctl.Index Mod 2) = 0 Then
        objctl.Top = lblAM.Top + ((lblAM.Height / 2) - (195 / 2))
      Else
        objctl.Top = lblPM.Top + ((lblPM.Height / 2) - (195 / 2))
      End If
    End If
  
    If TypeOf objctl Is PictureBox Then
      If CInt(objctl.Index Mod 2) = 0 Then
        objctl.Top = lblAM.Top + ((lblAM.Height / 2) - (195 / 2))
      Else
        objctl.Top = lblPM.Top + ((lblPM.Height / 2) - (195 / 2))
      End If
    End If
  
  Next objctl
  
  'Horizontally Align the Checkbox Images
  
  For Each objctl In UserControl.Controls
    If TypeOf objctl Is Image Then
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
  
    If TypeOf objctl Is PictureBox Then
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

  fraWorkingPattern.Height = (lblPM.Top + lblPM.Height) + 80
  fraWorkingPattern.Width = lblDay(7).Left + lblDay(7).Width + 100

  UserControl.Width = fraWorkingPattern.Width
  UserControl.Height = fraWorkingPattern.Height - 100

  On Error Resume Next
  lblDay(1) = UCase(Left(WeekdayName(7), 1))
  lblDay(2) = UCase(Left(WeekdayName(1), 1))
  lblDay(3) = UCase(Left(WeekdayName(2), 1))
  lblDay(4) = UCase(Left(WeekdayName(3), 1))
  lblDay(5) = UCase(Left(WeekdayName(4), 1))
  lblDay(6) = UCase(Left(WeekdayName(5), 1))
  lblDay(7) = UCase(Left(WeekdayName(6), 1))

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

Public Property Get hWnd() As Long
  
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
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

Private Sub lblToggle_Click()

'  ' Toggle status of all checkboxes
'
'  Dim intCounter As Integer
'  For intCounter = 1 To 14
'    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
'  Next intCounter

End Sub

Private Sub lblAM_Click()

'  ' Select all the AM checkboxes
'
'  Dim intCounter As Integer
'  For intCounter = 1 To 13 Step 2
'    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
'  Next intCounter

End Sub

Private Sub lblPM_Click()

  ' Select all the PM checkboxes

'  Dim intCounter As Integer
'  For intCounter = 2 To 14 Step 2
'    chkDay(intCounter).Value = IIf(chkDay(intCounter).Value = vbChecked, vbUnchecked, vbChecked)
'  Next intCounter

End Sub


Private Function CreateCharacterString() As String
'
'  Dim intCounter As Integer
'
'  For intCounter = 1 To 14
'    If chkDay(intCounter).Value = vbChecked Then
'      Select Case intCounter
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
'      End Select
'    Else
'      CreateCharacterString = CreateCharacterString & " "
'    End If
'  Next intCounter
'
'If Len(CreateCharacterString) < 14 Then
'  CreateCharacterString = CreateCharacterString & Space(14 - Len(CreateCharacterString))
'End If

End Function

Private Sub Picture1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

'##################################

Private Sub UserControl_Initialize()
  'Options = "None"
  On Error Resume Next
  
'  lblDay(1) = UCase(Left(WeekdayName(7), 1))
'  lblDay(2) = UCase(Left(WeekdayName(1), 1))
'  lblDay(3) = UCase(Left(WeekdayName(2), 1))
'  lblDay(4) = UCase(Left(WeekdayName(3), 1))
'  lblDay(5) = UCase(Left(WeekdayName(4), 1))
'  lblDay(6) = UCase(Left(WeekdayName(5), 1))
'  lblDay(7) = UCase(Left(WeekdayName(6), 1))

  lblDay(1) = UCase(Left(WeekdayName(1, , vbSunday), 1))
  lblDay(2) = UCase(Left(WeekdayName(2, , vbSunday), 1))
  lblDay(3) = UCase(Left(WeekdayName(3, , vbSunday), 1))
  lblDay(4) = UCase(Left(WeekdayName(4, , vbSunday), 1))
  lblDay(5) = UCase(Left(WeekdayName(5, , vbSunday), 1))
  lblDay(6) = UCase(Left(WeekdayName(6, , vbSunday), 1))
  lblDay(7) = UCase(Left(WeekdayName(7, , vbSunday), 1))
    
  mBackcolour = vbButtonFace
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub fraWorkingPattern_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub fraWorkingPattern_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub fraWorkingPattern_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub lblDay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblDay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub chkDay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub chkDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub chkDay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'###
Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub
'###
Private Sub lblAM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblAM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblAM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub lblPM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblPM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblPM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Resize()
  
    ' Resize the contained controls as the UserControl is resized.
    
'  With UserControl
'
'    If .Height <> fraWorkingPattern.Height + 80 Then
'      .Height = fraWorkingPattern.Height + 80
'    End If
'
'    If .Width <> fraWorkingPattern.Width + 100 Then
'      .Width = fraWorkingPattern.Width + 100
'    End If
    
    'If .Height < gLngMinHeight Then
    '  .Height = gLngMinHeight
    'End If
    'If .Width < gLngMinWidth Then
    '  .Width = gLngMinWidth
    'End If
     
'    If .Height > (lblPM.Top + lblPM.Height) + 100 Then .Height = (lblPM.Top + lblPM.Height) + 100
'    If .Width > (lblDay(7).Left + lblDay(7).Width + 150) Then .Width = (lblDay(7).Left + lblDay(7).Width + 150)
    
'    End With
    
End Sub


Public Property Get Read_Only() As Boolean
Attribute Read_Only.VB_Description = "Something Nick Said he wanted"
'NPG20071022
  Read_Only = mblnReadOnly

End Property


Public Property Let Read_Only(blnValue As Boolean)
'NPG20071022
  Dim lngIndex As Long

  mblnReadOnly = blnValue

  lblAM.Enabled = Not blnValue
  lblPM.Enabled = Not blnValue
  
  For lngIndex = lblDay.LBound To lblDay.UBound
    lblDay(lngIndex).Enabled = Not blnValue
    lblDay(lngIndex).Backcolor = IIf(mblnReadOnly, vbButtonFace, mBackcolour)
  Next

  For lngIndex = Picture1.LBound To Picture1.UBound
    Picture1(lngIndex).Backcolor = IIf(mblnReadOnly, vbButtonFace, vbWhite)
  Next

  fraWorkingPattern.Backcolor = IIf(mblnReadOnly, vbButtonFace, mBackcolour)
  lblToggle.Backcolor = IIf(mblnReadOnly, vbButtonFace, mBackcolour)
  lblAM.Backcolor = IIf(mblnReadOnly, vbButtonFace, mBackcolour)
  lblPM.Backcolor = IIf(mblnReadOnly, vbButtonFace, mBackcolour)


End Property

