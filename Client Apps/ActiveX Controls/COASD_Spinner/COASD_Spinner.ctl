VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl COASD_Spinner 
   BackColor       =   &H80000005&
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   ScaleHeight     =   1200
   ScaleWidth      =   2595
   Begin MSComCtl2.UpDown upDnSpinner 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   195
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   330
      TabIndex        =   1
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "COASD_Spinner"
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

' Declare Windows API Functions.
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

' System metric constants.
Public Enum SystemMetrics
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CXICON = 11
  SM_CXICONSPACING = 38
  SM_CYICON = 12
  SM_CYICONSPACING = 39
  SM_CYSMCAPTION = 51
  SM_CXVSCROLL = 2
End Enum

' Constant values.
Const gLngMinHeight = 200
Const gLngMinWidth = 200

' Properties.
Private giAlignment As Integer
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private mblnReadOnly As Boolean   'NPG20071022

Private mForecolour As OLE_COLOR
Private mBackcolour As OLE_COLOR

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = gfSelected
  
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  ' Set the Selected property.
  gfSelected = pfNewValue
    
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

Public Property Get Alignment() As Integer
  ' Return the Alignment property.
  Alignment = giAlignment
  
End Property

Public Property Let Alignment(ByVal piNewValue As Integer)
  ' Set the Alignment property if it is valid.
  
  If (piNewValue = vbLeftJustify) Or (piNewValue = vbRightJustify) Then
        
    giAlignment = piNewValue
    
    ' Redraw the control with the new alignment.
    UserControl_Resize
      
  End If
  
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub
Public Property Get Caption() As String
  ' Return the Caption property.
  Caption = lblLabel.Caption
  
End Property

Public Property Let Caption(ByVal psNewValue As String)
  ' Set the Caption property if it has changed.
  lblLabel.Caption = psNewValue
  
End Property

Public Property Get BackColor() As OLE_COLOR
  ' Return the control's background colour property.
  BackColor = UserControl.BackColor
  
End Property

Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's background colour property.
  UserControl.BackColor = pColNewColor
  lblLabel.BackColor = pColNewColor
  mBackcolour = pColNewColor
End Property
Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
End Property


Public Property Get Font() As Font
  ' Return the control's font property.
  Set Font = UserControl.Font
  
End Property

Public Property Set Font(ByVal pObjNewValue As StdFont)
  ' Set the control's font property.
  Dim iLoop As Integer
  
  ' Update the sub-controls.
  Set UserControl.Font = pObjNewValue
  Set lblLabel.Font = pObjNewValue
  
  UserControl_Resize
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ' Return the control's foreground colour property.
  ForeColor = UserControl.ForeColor
  
End Property

Public Property Let ForeColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's foreground colour property.
  UserControl.ForeColor = pColNewColor
  lblLabel.ForeColor = pColNewColor
  mForecolour = pColNewColor
End Property

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long
  
  lngMinHeight = UserControl.TextHeight("X") + (Screen.TwipsPerPixelY * 8)
  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)
  
End Property


Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  Dim lngMinWidth As Long
  
  lngMinWidth = (4 * GetSystemMetricsAPI(SM_CXFRAME) * Screen.TwipsPerPixelX) + _
    upDnSpinner.Width
    
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property



Private Sub lblLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Initialize()
  giAlignment = vbRightJustify
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


Private Sub UserControl_Resize()
  ' Resize the contained controls as the UserControl is resized.
  Dim lngControlWidth As Long
  Dim lngMinHeight As Long
  Dim lngMinWidth As Long
  
  ' Do not let the user make the control too small.
  lngMinHeight = MinimumHeight
  lngMinWidth = MinimumWidth
  
  With UserControl
    If .Width < lngMinWidth Then
      .Width = lngMinWidth
    End If
      
    lngControlWidth = .Width
  End With
  
  ' Resize the dummy spinner sub-controls.
  With lblLabel
    .Top = 0
    .Height = lngMinHeight
    .Width = lngControlWidth - upDnSpinner.Width
  End With
  
  With upDnSpinner
    .Top = 0
    .Height = lngMinHeight
  
    If giAlignment = vbLeftJustify Then
      .Left = 0
      lblLabel.Left = .Width
    Else
      .Left = lblLabel.Width
      lblLabel.Left = 0
    End If
  
  End With

  UserControl.Height = lngMinHeight

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Property Get Read_Only() As Boolean
'NPG20071022
  Read_Only = mblnReadOnly

End Property


Public Property Let Read_Only(blnValue As Boolean)
'NPG20071022
  mblnReadOnly = blnValue

'  AE20080424 Fault #13124
'  With lblLabel
'    .BackColor = IIf(blnValue, vbButtonFace, vbWindowBackground)
'    .ForeColor = IIf(blnValue, vbGrayText, vbWindowText)
'  End With
    
  With lblLabel
    .BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
    .ForeColor = IIf(blnValue, vbGrayText, mForecolour)
  End With
  
End Property
