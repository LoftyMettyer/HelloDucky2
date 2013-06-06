VERSION 5.00
Begin VB.UserControl COASD_Selection 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   810
   ScaleWidth      =   855
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   0
      Left            =   165
      MousePointer    =   8  'Size NW SE
      Tag             =   "TopLeft"
      Top             =   150
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   1
      Left            =   330
      MousePointer    =   7  'Size N S
      Tag             =   "TopCentre"
      Top             =   150
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   2
      Left            =   510
      MousePointer    =   6  'Size NE SW
      Tag             =   "TopRight"
      Top             =   150
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   3
      Left            =   165
      MousePointer    =   9  'Size W E
      Tag             =   "CentreLeft"
      Top             =   315
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   4
      Left            =   510
      MousePointer    =   9  'Size W E
      Tag             =   "CentreRight"
      Top             =   315
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   5
      Left            =   165
      MousePointer    =   6  'Size NE SW
      Tag             =   "BottomLeft"
      Top             =   495
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   6
      Left            =   330
      MousePointer    =   7  'Size N S
      Tag             =   "BottomCentre"
      Top             =   495
      Width           =   105
   End
   Begin VB.Image picSelectionMarker 
      Appearance      =   0  'Flat
      Height          =   105
      Index           =   7
      Left            =   510
      MousePointer    =   8  'Size NW SE
      Tag             =   "BottomRight"
      Top             =   495
      Width           =   105
   End
End
Attribute VB_Name = "COASD_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event Stretch(Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event StretchStart(Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event StretchEnd(Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mlngLastX As Long
Private mlngLastY As Long

Private mlngOriginalHeight As Long
Private mlngOriginalWidth As Long
Private mlngOriginalTop As Long
Private mlngOriginalLeft As Long
Private mbHasLockedWidth As Boolean
Private mbHasLockedHeight As Boolean

Private mbInWFDesigner As Boolean

Private mlngGridSize As Long

Private mobjAttachedControl As Object

Private mbMouseMoveToggle As Boolean

' Returns the size of a marker object
Public Property Get MarkerSize() As Integer
  MarkerSize = picSelectionMarker(0).Height
End Property

' Raise event for this selection marker being held down
Private Sub picSelectionMarker_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  RaiseEvent StretchStart(picSelectionMarker(Index).Tag, Button, Shift, X, Y)

End Sub

' Raise event for this selection marker being moved
Private Sub picSelectionMarker_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If (mlngLastX > X + mlngGridSize) Or (mlngLastX < X - mlngGridSize) _
      Or (mlngLastY > Y + mlngGridSize) Or (mlngLastY < Y - mlngGridSize) Then

      If mbMouseMoveToggle And Button = vbLeftButton Then
        RaiseEvent Stretch(picSelectionMarker(Index).Tag, Button, Shift, X, Y)
      End If
            
      mlngLastX = X
      mlngLastY = Y

  End If

  mbMouseMoveToggle = Not mbMouseMoveToggle

End Sub

' Raise event for this selection marker being released
Private Sub picSelectionMarker_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent StretchEnd(picSelectionMarker(Index).Tag, Button, Shift, X, Y)
End Sub

Public Sub ShowSelectionMarkers(pbShow As Boolean)

  Dim iCount As Integer

  ' Hide the selection markers
  For iCount = 0 To picSelectionMarker.Count - 1
    picSelectionMarker(iCount).Visible = pbShow
  Next iCount

End Sub

Public Sub RefreshSelectionMarkers(pbShow As Boolean)

  Dim iMarkerSize As Integer
  
  iMarkerSize = MarkerSize

  ' Top Left
  picSelectionMarker(0).Top = 0
  picSelectionMarker(0).Left = 0

  ' Top Centre
  picSelectionMarker(1).Top = 0
  picSelectionMarker(1).Left = (UserControl.Width - iMarkerSize) / 2

  ' Top Right
  picSelectionMarker(2).Top = 0
  picSelectionMarker(2).Left = UserControl.Width - iMarkerSize

  ' Centre Left
  picSelectionMarker(3).Top = (UserControl.Height - iMarkerSize) / 2
  picSelectionMarker(3).Left = 0

  ' Centre Right
  picSelectionMarker(4).Top = (UserControl.Height - iMarkerSize) / 2
  picSelectionMarker(4).Left = UserControl.Width - iMarkerSize

  ' Bottom Left
  picSelectionMarker(5).Top = UserControl.Height - iMarkerSize
  picSelectionMarker(5).Left = 0

  ' Bottom Centre
  picSelectionMarker(6).Top = UserControl.Height - iMarkerSize
  picSelectionMarker(6).Left = (UserControl.Width - iMarkerSize) / 2

  ' Bottom Right
  picSelectionMarker(7).Top = UserControl.Height - iMarkerSize
  picSelectionMarker(7).Left = UserControl.Width - iMarkerSize

  ' Hide/Show selection markers
  ShowSelectionMarkers pbShow

End Sub

' Set the attached control
Public Property Let AttachedObject(pctlControl As Object)
  
  Dim sName As String
  
  Set mobjAttachedControl = pctlControl

  sName = LCase(pctlControl.Name)
  
  ' JDM - Fault 4810 - 26/11/02 - Should be able to resize a link button
  'sName = "asrdummylink" Or
  If sName = "asrcustomdummywp" Or _
      sName = "asrdummyoptions" Then
    mbHasLockedHeight = True
    mbHasLockedWidth = True
  End If

  'If sName = "asrdummyline" Or
  If sName = "asrdummycheckbox" Or _
      sName = "asrdummyspinner" Or _
      sName = "asrdummycombo" Then
    mbHasLockedHeight = True
  End If

'  'AE20080109 Lines can now be vertical in WF Designer
'  If sName = "asrdummyline" And Not mbInWFDesigner Then
'    mbHasLockedHeight = True
'  End If
    
End Property

' Access the attached control
Public Property Get AttachedObject() As Object
  Set AttachedObject = mobjAttachedControl
End Property

' Initialse the selection markers
Private Sub UserControl_Initialize()
  mlngGridSize = 10
  mbMouseMoveToggle = True
End Sub

' Returns the size of the pre-stretched control
Public Property Get Original_Height() As Long
  Original_Height = mlngOriginalHeight
End Property

' Returns the size of the pre-stretched control
Public Property Get Original_Width() As Long
  Original_Width = mlngOriginalWidth
End Property

' Returns the size of the pre-stretched control
Public Property Get Original_Top() As Long
  Original_Top = mlngOriginalTop
End Property

' Returns the size of the pre-stretched control
Public Property Get Original_Left() As Long
  Original_Left = mlngOriginalLeft
End Property

' Store the original size of the attached object (for stretching purposes)
Public Sub SaveOriginalSizes()
  mlngOriginalWidth = mobjAttachedControl.Width
  mlngOriginalHeight = mobjAttachedControl.Height
  mlngOriginalTop = mobjAttachedControl.Top
  mlngOriginalLeft = mobjAttachedControl.Left
End Sub

' Does the attached control have a locked height
Public Property Get HasLockedHeight() As Boolean
  
  If LCase(mobjAttachedControl.Name) = "asrdummyline" And mbInWFDesigner Then
    If mobjAttachedControl.Alignment = 1 Then
      ' Horizontal
      mbHasLockedHeight = True
    Else
      ' Vertical
      mbHasLockedHeight = False
    End If
  End If
  
  HasLockedHeight = mbHasLockedHeight
End Property

' Does the attached control have a locked width
Public Property Get HasLockedWidth() As Boolean

  If LCase(mobjAttachedControl.Name) = "asrdummyline" And mbInWFDesigner Then
    If mobjAttachedControl.Alignment = 1 Then
      ' Horizontal
      mbHasLockedWidth = False
    Else
      ' Vertical
      mbHasLockedWidth = True
    End If
  End If
  
  HasLockedWidth = mbHasLockedWidth
End Property

Public Property Get WFDesigner() As Boolean
  WFDesigner = mbInWFDesigner
End Property

Public Property Let WFDesigner(ByVal vNewValue As Boolean)
  mbInWFDesigner = vNewValue
End Property
