VERSION 5.00
Begin VB.UserControl COAWF_Link 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ClipControls    =   0   'False
   DrawWidth       =   2
   FillColor       =   &H80000008&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MaskColor       =   &H000000FF&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   1050
   ScaleWidth      =   1725
   ToolboxBitmap   =   "COAWF_Link.ctx":0000
   Begin VB.Line lnArrowHead 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   720
      X2              =   855
      Y1              =   810
      Y2              =   720
   End
   Begin VB.Line lnArrowHead 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   540
      Y1              =   675
      Y2              =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      Visible         =   0   'False
      X1              =   1000
      X2              =   1000
      Y1              =   200
      Y2              =   400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      Visible         =   0   'False
      X1              =   800
      X2              =   800
      Y1              =   200
      Y2              =   400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      Visible         =   0   'False
      X1              =   600
      X2              =   600
      Y1              =   200
      Y2              =   400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      Visible         =   0   'False
      X1              =   400
      X2              =   400
      Y1              =   200
      Y2              =   400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   200
      X2              =   200
      Y1              =   200
      Y2              =   400
   End
End
Attribute VB_Name = "COAWF_Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

Public Enum ArrowDirection
  arrowDirection_Down = 0
  arrowDirection_Left = 1
  arrowDirection_Right = 2
  arrowDirection_Up = 3
End Enum

Public Enum LineDirection
  lineDirection_down = 0
  lineDirection_Left = 1
  lineDirection_Right = 2
  lineDirection_Up = 3
End Enum

Private msngCurveRadius_1 As Single
Private msngCurveRadius_2 As Single
Private mintLines As Integer

Private miStartDirection As LineDirection
Private miEndDirection As LineDirection
Private mfHighlighted As Boolean

Private msngXOffset As Single
Private msngYOffset As Single

Private msngBorder As Single

Private miStartElementIndex As Integer
Private miEndElementIndex As Integer
Private miStartOutboundFlowCode As Integer
Private mblnCurvedLinks As Boolean

Private Const ARROWHEADWITH = 100
Private Const MINSTARTENDLENGTH = 350
Private Const PIXELWIDTH = 15
Private Const MAXCURVERADIUS = 105
Private Const MINCURVERADIUS = 0

' App Version properties
Private miAppMajor As Integer
Private miAppMinor As Integer
Private miAppRevision As Integer

Public Property Get AppMajor() As Integer
  AppMajor = miAppMajor
End Property

Public Property Let AppMajor(ByVal piNewValue As Integer)
  miAppMajor = piNewValue
  Call SetBackColour
End Property

Public Property Get AppMinor() As Integer
  AppMinor = miAppMinor
End Property

Public Property Let AppMinor(ByVal piNewValue As Integer)
  miAppMinor = piNewValue
  Call SetBackColour
End Property

Public Property Get AppRevision() As Integer
  AppRevision = miAppRevision
End Property

Public Property Let AppRevision(ByVal piNewValue As Integer)
  miAppRevision = piNewValue
  Call SetBackColour
End Property

Private Sub ASRWFLinkArrow1_DblClick()
  ' Pass the DblClick event to the parent form.
  RaiseEvent DblClick

End Sub

Private Sub ASRWFLinkArrow1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Sub picCurve_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Sub UserControl_DblClick()
  ' Pass the DblClick event to the parent form.
  RaiseEvent DblClick

End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  StartDirection = lineDirection_down
  EndDirection = lineDirection_Up
  XOffset = 0
  YOffset = 420
End Sub

Public Property Let CurvedLinks(New_Value As Boolean)
  mblnCurvedLinks = New_Value
  FormatLines
End Property

Public Property Get CurvedLinks() As Boolean
  CurvedLinks = mblnCurvedLinks
End Property

Private Sub Calculate_Radii()
  
  Exit Sub
  If Not mblnCurvedLinks Then
    Exit Sub
  End If
  
  'Calculate the required radius for
  Dim gap As Single
  Dim start As Single
  Dim endY As Single

  Select Case mintLines
    Case 5
      gap = (MINSTARTENDLENGTH + msngBorder) - (msngYOffset - MINSTARTENDLENGTH)
      gap = IIf(gap < 0, (-1 * gap), gap)
      If (gap > 15) And (gap < (2 * MAXCURVERADIUS)) Then
        msngCurveRadius_2 = FixTwips(gap / 2) - PIXELWIDTH
      ElseIf (gap < 15) Then
        msngCurveRadius_2 = 0
      Else
        msngCurveRadius_2 = MAXCURVERADIUS
      End If
       
      gap = msngXOffset - msngBorder
      gap = IIf(gap < 0, (-1 * gap), gap)
      If (gap > 15) And (gap < (2 * MAXCURVERADIUS)) Then
        msngCurveRadius_1 = FixTwips(gap / 2) - PIXELWIDTH
      ElseIf (gap < 15) Then
        msngCurveRadius_1 = 0
      Else
        msngCurveRadius_1 = MAXCURVERADIUS
      End If
     
    Case 4
      gap = (MINSTARTENDLENGTH + msngBorder) - (msngYOffset - MINSTARTENDLENGTH)
      gap = IIf(gap < 0, (-1 * gap), gap)
      If (gap > 15) And (gap < (2 * MAXCURVERADIUS)) Then
        msngCurveRadius_2 = FixTwips(gap / 2) - PIXELWIDTH
      ElseIf (gap < 15) Then
        msngCurveRadius_2 = 0
      Else
        msngCurveRadius_2 = MAXCURVERADIUS
      End If
       
      gap = msngXOffset - msngBorder
      gap = IIf(gap < 0, (-1 * gap), gap)
      If (gap > 15) And (gap < (2 * MAXCURVERADIUS)) Then
        msngCurveRadius_1 = FixTwips(gap / 2) - PIXELWIDTH
      ElseIf (gap < 15) Then
        msngCurveRadius_1 = 0
      Else
        msngCurveRadius_1 = MAXCURVERADIUS
      End If
      
    Case 3
      If StartDirection = lineDirection_down Or StartDirection = lineDirection_Up Then
        
        If (msngXOffset < (2 * PIXELWIDTH)) And (msngXOffset > (-2 * PIXELWIDTH)) Then
          msngCurveRadius_1 = 0
          msngCurveRadius_2 = 0
        
        ElseIf (msngXOffset < (2 * MAXCURVERADIUS)) And (msngXOffset > (-2 * MAXCURVERADIUS)) Then
          msngCurveRadius_1 = FixTwips(msngXOffset / 2)
          msngCurveRadius_2 = 0
        
        Else
          msngCurveRadius_1 = MAXCURVERADIUS
          msngCurveRadius_2 = 0
        End If
      
      Else
        msngCurveRadius_1 = MAXCURVERADIUS
        msngCurveRadius_2 = 0
        
      End If
      
    Case 2
      msngCurveRadius_1 = MAXCURVERADIUS
      msngCurveRadius_2 = 0

    Case 1
      msngCurveRadius_1 = 0
      msngCurveRadius_2 = 0
      
  End Select
  
  If msngCurveRadius_1 < 0 Then msngCurveRadius_1 = (msngCurveRadius_1 * -1)
  If msngCurveRadius_2 < 0 Then msngCurveRadius_2 = (msngCurveRadius_2 * -1)

End Sub

Public Property Let XOffset(ByVal psngNewValue As Single)
  ' Set the XOffset.
  ' ie. the horizontal distance between the start position and the end position.
  msngXOffset = FixTwips(psngNewValue)
  PropertyChanged "XOffset"
  
  FormatLines
  
End Property

Public Property Let YOffset(ByVal psngNewValue As Single)
  ' Set the YOffset.
  ' ie. the vertical distance between the start position and the end position.
  msngYOffset = FixTwips(psngNewValue)
  PropertyChanged "YOffset"
  
  FormatLines
  
End Property

Private Function FixTwips(pValue As Single) As Single
  FixTwips = Round((pValue / 15), 0) * 15
End Function

Private Sub UserControl_Initialize()
  msngBorder = FixTwips(ARROWHEADWITH)
End Sub

Private Sub JoinLines_2Lines()

  ' Join Two lines

  On Error Resume Next

  Dim negX As Boolean
  Dim negY As Boolean
  negX = (msngXOffset < 0)
  negY = (msngYOffset < 0)

  If (miStartDirection = lineDirection_down) And (miEndDirection = lineDirection_Right) _
    Or (miStartDirection = lineDirection_down) And (miEndDirection = lineDirection_Left) Then
    '******************************
    '
    '   |                 |
    '   |       OR        | (negX)
    '   |                 |
    '   o---->       <----o
    '
    '******************************
'    If negX Then
'      picCurve(2).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(2).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(2).Width = msngCurveRadius_1
'      picCurve(2).Height = picCurve(2).Width
'      DrawCurve picCurve(2), curveType_BottomRight
'
'      picCurve(0).Visible = False
'      picCurve(1).Visible = False
'      picCurve(3).Visible = False
'    Else
'      picCurve(3).Left = Line1(1).x1
'      picCurve(3).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(3).Width = msngCurveRadius_1
'      picCurve(3).Height = picCurve(3).Width
'      DrawCurve picCurve(3), curveType_BottomLeft
'
'      picCurve(0).Visible = False
'      picCurve(1).Visible = False
'      picCurve(2).Visible = False
'    End If
    
  ElseIf (miStartDirection = lineDirection_Up) And (miEndDirection = lineDirection_Right) _
    Or (miStartDirection = lineDirection_Up) And (miEndDirection = lineDirection_Left) Then
    '******************************
    '
    '   o---->       <----o
    '   |       OR        |
    '   |                 | (negX)
    '   |                 |
    '
    '******************************
'    If negX Then
'      picCurve(1).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(1).Top = Line1(1).y1
'      picCurve(1).Width = msngCurveRadius_1
'      picCurve(1).Height = picCurve(1).Width
'      DrawCurve picCurve(1), curveType_TopRight
'
'      picCurve(0).Visible = False
'      picCurve(2).Visible = False
'      picCurve(3).Visible = False
'    Else
'      picCurve(0).Left = Line1(0).x2
'      picCurve(0).Top = Line1(1).y1
'      picCurve(0).Width = msngCurveRadius_1
'      picCurve(0).Height = picCurve(0).Width
'      DrawCurve picCurve(0), curveType_TopLeft
'
'      picCurve(1).Visible = False
'      picCurve(2).Visible = False
'      picCurve(3).Visible = False
'    End If
'
  ElseIf (miStartDirection = lineDirection_Right) And (miEndDirection = lineDirection_down) _
    Or (miStartDirection = lineDirection_Right) And (miEndDirection = lineDirection_Up) Then
    '******************************
    '
    '   -----o            ^
    '        |   OR       |
    '        |            | (negY)
    '        v       -----o
    '
    '******************************
'    If negY Then
'      picCurve(2).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(2).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(2).Width = msngCurveRadius_1
'      picCurve(2).Height = picCurve(2).Width
'      DrawCurve picCurve(2), curveType_BottomRight
'
'      picCurve(0).Visible = False
'      picCurve(1).Visible = False
'      picCurve(3).Visible = False
'    Else
'      picCurve(1).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(1).Top = Line1(0).y1
'      picCurve(1).Width = msngCurveRadius_1
'      picCurve(1).Height = picCurve(1).Width
'      DrawCurve picCurve(1), curveType_TopRight
'
'      picCurve(0).Visible = False
'      picCurve(2).Visible = False
'      picCurve(3).Visible = False
'    End If
    
  ElseIf (miStartDirection = lineDirection_Left) And (miEndDirection = lineDirection_down) _
    Or (miStartDirection = lineDirection_Left) And (miEndDirection = lineDirection_Up) Then
    '******************************
    '
    '   o-----       ^
    '   |        OR  |
    '   |            |      (negY)
    '   v            o-----
    '
    '******************************
'    If negY Then
'      picCurve(3).Left = Line1(1).x1
'      picCurve(3).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'      picCurve(3).Width = msngCurveRadius_1
'      picCurve(3).Height = picCurve(3).Width
'      DrawCurve picCurve(3), curveType_BottomLeft
'
'      picCurve(0).Visible = False
'      picCurve(1).Visible = False
'      picCurve(2).Visible = False
'    Else
'      picCurve(0).Left = Line1(1).x1
'      picCurve(0).Top = Line1(0).y2
'      picCurve(0).Width = msngCurveRadius_1
'      picCurve(0).Height = picCurve(0).Width
'      DrawCurve picCurve(0), curveType_TopLeft
'
'      picCurve(1).Visible = False
'      picCurve(2).Visible = False
'      picCurve(3).Visible = False
'    End If

  End If
  
End Sub

Private Sub JoinLines_3Lines()
'
'  ' Join Three Lines
'
'  On Error Resume Next
'
'  Dim negX As Boolean
'  Dim negY As Boolean
'  negX = (msngXOffset < 0)
'  negY = (msngYOffset < 0)
'
'  Select Case miStartDirection
'    Case lineDirection_down
'      Select Case miEndDirection
'        Case lineDirection_down
'          '********************************
'          '
'          '   |     ^       ^     |
'          '   |     |  OR   |     | (negX)
'          '   |     |       |     |
'          '   o-----o       o-----o
'          '
'          '********************************
'          If negX Then
'            picCurve(2).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(3).Left = Line1(2).x1
'            picCurve(3).Top = Line1(2).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(0).Visible = False
'            picCurve(1).Visible = False
'          Else
'            picCurve(3).Left = Line1(0).x2
'            picCurve(3).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(2).Left = Line1(1).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(2).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Visible = False
'            picCurve(1).Visible = False
'          End If
'
'        Case lineDirection_up
'          '*********************************
'          '
'          '   |                    |
'          '   |                    | (negX)
'          '   |                    |
'          '   o-----o   OR   o-----o
'          '         |        |
'          '         |        |
'          '         v        v
'          '
'          '*********************************
'          If negX Then
'            picCurve(2).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Left = Line1(2).x1
'            picCurve(0).Top = Line1(1).y2
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(1).Visible = False
'            picCurve(3).Visible = False
'          Else
'            picCurve(3).Left = Line1(0).x2
'            picCurve(3).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(1).Left = Line1(1).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(1).y2
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(0).Visible = False
'            picCurve(2).Visible = False
'          End If
'
'      End Select
'
'    Case lineDirection_up
'      Select Case miEndDirection
'        Case lineDirection_down
'          '********************************
'          '
'          '         ^       ^
'          '         |       |
'          '         |       |
'          '   o-----o   OR  o-----o
'          '   |                   |
'          '   |                   | (negX)
'          '   |                   |
'          '
'          '********************************
'          If negX Then
'            picCurve(1).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(1).y1
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(3).Left = Line1(2).x1
'            picCurve(3).Top = Line1(2).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(0).Visible = False
'            picCurve(2).Visible = False
'          Else
'            picCurve(0).Left = Line1(0).x2
'            picCurve(0).Top = Line1(1).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(2).Left = Line1(1).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(2).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(1).Visible = False
'            picCurve(3).Visible = False
'          End If
'
'        Case lineDirection_up
'          '********************************
'          '
'          '   o-----o       o-----o
'          '   |     |   OR  |     | (negX)
'          '   |     |       |     |
'          '   |     v       v     |
'          '
'          '********************************
'          If negX Then
'            picCurve(1).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(1).y1
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(0).Left = Line1(2).x1
'            picCurve(0).Top = Line1(1).y2
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(2).Visible = False
'            picCurve(3).Visible = False
'         Else
'            picCurve(0).Left = Line1(0).x2
'            picCurve(0).Top = Line1(1).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(1).Left = Line1(1).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(1).y2
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(2).Visible = False
'            picCurve(3).Visible = False
'          End If
'
'      End Select
'
'    Case lineDirection_right
'      Select Case miEndDirection
'        Case lineDirection_right
'          '*****************************
'          '
'          '   -----o      <----o
'          '        |           |
'          '        |  OR       | (negY)
'          '        |           |
'          '   <----o      -----o
'          '
'          '*****************************
'          If negY Then
'            picCurve(2).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(1).Left = Line1(2).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(2).y1
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(0).Visible = False
'            picCurve(3).Visible = False
'          Else
'            picCurve(1).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(0).y2
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(2).Left = Line1(2).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Visible = False
'            picCurve(3).Visible = False
'          End If
'
'        Case lineDirection_left
'          '**********************************
'          '
'          '   -----o                o---->
'          '        |                |
'          '        |       OR       | (negY)
'          '        |                |
'          '        o---->      -----o
'          '
'          '**********************************
'          If negY Then
'            picCurve(2).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Left = Line1(1).x2
'            picCurve(0).Top = Line1(2).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(1).Visible = False
'            picCurve(3).Visible = False
'          Else
'            picCurve(1).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(0).y2
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(3).Left = Line1(1).x2
'            picCurve(3).Top = Line1(1).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(0).Visible = False
'            picCurve(2).Visible = False
'          End If
'
'      End Select
'
'    Case lineDirection_left
'      Select Case miEndDirection
'        Case lineDirection_right
'          '***********************************
'          '
'          '        o-----       <----o
'          '        |                 |
'          '        |        OR       | (negY)
'          '        |                 |
'          '   <----o                 o-----
'          '
'          '***********************************
'          If negY Then
'            picCurve(3).Left = Line1(1).x1
'            picCurve(3).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(1).Left = Line1(2).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(2).y1
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(0).Visible = False
'            picCurve(2).Visible = False
'          Else
'            picCurve(0).Left = Line1(1).x1
'            picCurve(0).Top = Line1(0).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(2).Left = Line1(2).x1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y2
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(1).Visible = False
'            picCurve(3).Visible = False
'         End If
'
'        Case lineDirection_left
'          '*****************************
'          '
'          '   o-----      o---->
'          '   |           |
'          '   |       OR  |      (negY)
'          '   |           |
'          '   o---->      o-----
'          '
'          '*****************************
'          If negY Then
'            picCurve(3).Left = Line1(1).x1
'            picCurve(3).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(0).Left = Line1(1).x2
'            picCurve(0).Top = Line1(2).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(1).Visible = False
'            picCurve(2).Visible = False
'          Else
'            picCurve(0).Left = Line1(1).x1
'            picCurve(0).Top = Line1(0).y1
'            picCurve(0).Width = msngCurveRadius_1
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(3).Left = Line1(1).x2
'            picCurve(3).Top = Line1(1).y2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(3).Width = msngCurveRadius_1
'            picCurve(3).Height = picCurve(3).Width
'            DrawCurve picCurve(3), curveType_BottomLeft
'
'            picCurve(1).Visible = False
'            picCurve(2).Visible = False
'          End If
'
'      End Select
'  End Select

End Sub

Private Sub JoinLines_4Lines()
'
'  ' Join Four Lines
'
'  On Error Resume Next
'
'  Dim negX As Boolean
'  Dim negY As Boolean
'  negX = (msngXOffset < 0)
'  negY = (msngYOffset < 0)
'
'  Select Case miStartDirection
'    Case lineDirection_down
'      Select Case miEndDirection
'        Case lineDirection_right
'          '******************
'          '
'          '   <----o     |
'          '        |     |
'          '        |     |
'          '        o-----o
'          '
'          '******************
'          picCurve(2).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'          picCurve(2).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'          picCurve(2).Width = msngCurveRadius_1
'          picCurve(2).Height = picCurve(2).Width
'          DrawCurve picCurve(2), curveType_BottomRight
'
'          picCurve(3).Left = Line1(2).x1
'          picCurve(3).Top = Line1(2).y1 - msngCurveRadius_2 + PIXELWIDTH
'          picCurve(3).Width = msngCurveRadius_2
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(1).Left = Line1(3).x1 - msngCurveRadius_2 + PIXELWIDTH
'          picCurve(1).Top = Line1(3).y1
'          picCurve(1).Width = msngCurveRadius_2
'          picCurve(1).Height = picCurve(1).Width
'          DrawCurve picCurve(1), curveType_TopRight
'
'        Case lineDirection_left
'          '******************
'          '
'          '   |     o---->
'          '   |     |
'          '   |     |
'          '   o-----o
'          '
'          '******************
'          picCurve(3).Left = Line1(0).x2
'          picCurve(3).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'          picCurve(3).Width = msngCurveRadius_1
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(2).Left = Line1(1).x2 - msngCurveRadius_2 + PIXELWIDTH
'          picCurve(2).Top = Line1(2).y1 - msngCurveRadius_2 + PIXELWIDTH
'          picCurve(2).Width = msngCurveRadius_1
'          picCurve(2).Height = picCurve(2).Width
'          DrawCurve picCurve(2), curveType_BottomRight
'
'          picCurve(0).Left = Line1(2).x1
'          picCurve(0).Top = Line1(3).y1
'          picCurve(0).Width = msngCurveRadius_2
'          picCurve(0).Height = picCurve(0).Width
'          DrawCurve picCurve(0), curveType_TopLeft
'
'      End Select
'
'    Case lineDirection_up
'      Select Case miEndDirection
'        Case lineDirection_right
'          '******************
'          '
'          '        o-----o
'          '        |     |
'          '        |     |
'          '   <----o     |
'          '
'          '******************
'          picCurve(1).Left = Line1(1).x1
'          picCurve(1).Top = Line1(1).y1
'          picCurve(1).Width = msngCurveRadius_1
'          picCurve(1).Height = picCurve(1).Width
'          DrawCurve picCurve(1), curveType_TopRight
'
'          picCurve(0).Left = Line1(2).x1
'          picCurve(0).Top = Line1(1).y2
'          picCurve(0).Width = msngCurveRadius_2
'          picCurve(0).Height = picCurve(0).Width
'          DrawCurve picCurve(0), curveType_TopLeft
'
'          picCurve(2).Left = Line1(3).x1
'          picCurve(2).Top = Line1(2).y2
'          picCurve(2).Width = msngCurveRadius_2
'          picCurve(2).Height = picCurve(2).Width
'          DrawCurve picCurve(2), curveType_BottomRight
'
'          picCurve(3).Visible = False
'
'        Case lineDirection_left
'          '******************
'          '
'          '   o-----o
'          '   |     |
'          '   |     |
'          '   |     o---->
'          '
'          '******************
'          picCurve(0).Left = Line1(0).x2
'          picCurve(0).Top = Line1(1).y1
'          picCurve(0).Width = msngCurveRadius_1
'          picCurve(0).Height = picCurve(0).Width
'          DrawCurve picCurve(0), curveType_TopLeft
'
'          picCurve(1).Left = Line1(1).x2
'          picCurve(1).Top = Line1(1).y2
'          picCurve(1).Width = msngCurveRadius_2
'          picCurve(1).Height = picCurve(1).Width
'          DrawCurve picCurve(1), curveType_TopRight
'
'          picCurve(3).Left = Line1(2).x2
'          picCurve(3).Top = Line1(2).y2
'          picCurve(3).Width = msngCurveRadius_2
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(2).Visible = False
'
'      End Select
'
'    Case lineDirection_right
'      Select Case miEndDirection
'        Case lineDirection_down
'          '******************
'          '
'          '   -----o
'          '        |     ^
'          '        |     |
'          '        o-----o
'          '
'          '******************
'          picCurve(1).Left = Line1(0).x2
'          picCurve(1).Top = Line1(0).y2
'          picCurve(1).Width = msngCurveRadius_1
'          picCurve(1).Height = picCurve(1).Width
'          DrawCurve picCurve(1), curveType_TopRight
'
'          picCurve(3).Left = Line1(1).x2
'          picCurve(3).Top = Line1(1).y2
'          picCurve(3).Width = msngCurveRadius_2
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(2).Left = Line1(2).x2
'          picCurve(2).Top = Line1(3).y1
'          picCurve(2).Width = msngCurveRadius_2
'          picCurve(2).Height = picCurve(2).Width
'          DrawCurve picCurve(2), curveType_BottomRight
'
'          picCurve(0).Visible = False
'
'        Case lineDirection_up
'          If (negX Or (msngXOffset < MINSTARTENDLENGTH)) And negY Then
'            '******************
'            '
'            '   o-----o
'            '   |     |
'            '   v     |
'            '         |
'            '    -----o
'            '
'            '******************
'            picCurve(2).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(1).Left = Line1(2).x1 - msngCurveRadius_2 + PIXELWIDTH
'            picCurve(1).Top = Line1(2).y1
'            picCurve(1).Width = msngCurveRadius_2
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(0).Left = Line1(3).x2
'            picCurve(0).Top = Line1(2).y2
'            picCurve(0).Width = msngCurveRadius_2
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(3).Visible = False
'
'          ElseIf negX And (Not negY) Then
'            '***************
'            '
'            '     -----o
'            '          |
'            '          |
'            '    o-----o
'            '    |
'            '    |
'            '    v
'            '
'            '***************
'            picCurve(1).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(1).Top = Line1(0).y1
'            picCurve(1).Width = msngCurveRadius_1
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(2).Left = Line1(2).x1 - msngCurveRadius_2 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y2 - msngCurveRadius_2 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_2
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Left = Line1(3).x1
'            picCurve(0).Top = Line1(2).y2
'            picCurve(0).Width = msngCurveRadius_2
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(3).Visible = False
'
'          Else
'            '******************
'            '
'            '        o-----o
'            '        |     |
'            '        |     v
'            '   -----o
'            '
'            '******************
'            picCurve(2).Left = Line1(0).x2 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Top = Line1(1).y1 - msngCurveRadius_1 + PIXELWIDTH
'            picCurve(2).Width = msngCurveRadius_1
'            picCurve(2).Height = picCurve(2).Width
'            DrawCurve picCurve(2), curveType_BottomRight
'
'            picCurve(0).Left = Line1(1).x2
'            picCurve(0).Top = Line1(2).y1
'            picCurve(0).Width = msngCurveRadius_2
'            picCurve(0).Height = picCurve(0).Width
'            DrawCurve picCurve(0), curveType_TopLeft
'
'            picCurve(1).Left = Line1(2).x2 - msngCurveRadius_2 + PIXELWIDTH
'            picCurve(1).Top = Line1(2).y2
'            picCurve(1).Width = msngCurveRadius_2
'            picCurve(1).Height = picCurve(1).Width
'            DrawCurve picCurve(1), curveType_TopRight
'
'            picCurve(3).Visible = False
'
'          End If
'
'      End Select
'
'    Case lineDirection_left
'      Select Case miEndDirection
'        Case lineDirection_down
'          '******************
'          '
'          '         o-----
'          '   ^     |
'          '   |     |
'          '   o-----o
'          '
'          '******************
'          picCurve(0).Left = Line1(1).x1
'          picCurve(0).Top = Line1(0).y1
'          picCurve(0).Width = msngCurveRadius_1
'          picCurve(0).Height = picCurve(0).Width
'          DrawCurve picCurve(0), curveType_TopLeft
'
'          picCurve(2).Left = Line1(2).x1
'          picCurve(2).Top = Line1(1).y2
'          picCurve(2).Width = msngCurveRadius_2
'          picCurve(2).Height = picCurve(2).Width
'          DrawCurve picCurve(2), curveType_BottomRight
'
'          picCurve(3).Left = Line1(3).x1
'          picCurve(3).Top = Line1(3).y1
'          picCurve(3).Width = msngCurveRadius_2
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(1).Visible = False
'
'        Case lineDirection_up
'          '******************
'          '
'          '   o-----o
'          '   |     |
'          '   v     |
'          '         o-----
'          '
'          '******************
'          picCurve(3).Left = Line1(1).x1
'          picCurve(3).Top = Line1(1).y1
'          picCurve(3).Width = msngCurveRadius_1
'          picCurve(3).Height = picCurve(3).Width
'          DrawCurve picCurve(3), curveType_BottomLeft
'
'          picCurve(1).Left = Line1(2).x1
'          picCurve(1).Top = Line1(2).y1
'          picCurve(1).Width = msngCurveRadius_2
'          picCurve(1).Height = picCurve(1).Width
'          DrawCurve picCurve(1), curveType_TopRight
'
'          picCurve(0).Left = Line1(3).x1
'          picCurve(0).Top = Line1(2).y2
'          picCurve(0).Width = msngCurveRadius_2
'          picCurve(0).Height = picCurve(0).Width
'          DrawCurve picCurve(0), curveType_TopLeft
'
'          picCurve(2).Visible = False
'
'      End Select
'  End Select
'
End Sub

Private Sub JoinLines_5Lines()
'
'  ' Join Five Lines
'
'  On Error Resume Next
'
'  Dim negX As Boolean
'  Dim negY As Boolean
'  negX = (msngXOffset < 0)
'  negY = (msngYOffset < 0)
'
'  Select Case miStartDirection
'    Case lineDirection_down
'      If negX And (msngYOffset < (2 * MINSTARTENDLENGTH)) Then
'        '***********************
'        '
'        '    o-----o
'        '    |     |
'        '    v     |     |
'        '          |     |
'        '          o-----o
'        '
'        '***********************
'        picCurve(2).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(2).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(2).Width = msngCurveRadius_1
'        picCurve(2).Height = picCurve(2).Width
'        DrawCurve picCurve(2), curveType_BottomRight
'
'        picCurve(3).Left = Line1(2).x1
'        picCurve(3).Top = Line1(2).y1 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(3).Width = msngCurveRadius_2
'        picCurve(3).Height = picCurve(3).Width
'        DrawCurve picCurve(3), curveType_BottomLeft
'
'        picCurve(1).Left = Line1(3).x1 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(1).Top = Line1(3).y1
'        picCurve(1).Width = msngCurveRadius_2
'        picCurve(1).Height = picCurve(1).Width
'        DrawCurve picCurve(1), curveType_TopRight
'
'        picCurve(0).Left = Line1(4).x1
'        picCurve(0).Top = Line1(3).y2
'        picCurve(0).Width = msngCurveRadius_1
'        picCurve(0).Height = picCurve(0).Width
'        DrawCurve picCurve(0), curveType_TopLeft
'
'      ElseIf (msngYOffset < (2 * MINSTARTENDLENGTH)) Then
'        '***********************
'        '
'        '          o-----o
'        '          |     |
'        '    |     |     v
'        '    |     |
'        '    o-----o
'        '
'        '***********************
'        picCurve(3).Left = Line1(0).x2
'        picCurve(3).Top = Line1(0).y2 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(3).Width = msngCurveRadius_1
'        picCurve(3).Height = picCurve(3).Width
'        DrawCurve picCurve(3), curveType_BottomLeft
'
'        picCurve(2).Left = Line1(1).x2 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Top = Line1(2).y1 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Width = msngCurveRadius_2
'        picCurve(2).Height = picCurve(2).Width
'        DrawCurve picCurve(2), curveType_BottomRight
'
'        picCurve(0).Left = Line1(2).x2
'        picCurve(0).Top = Line1(3).y1
'        picCurve(0).Width = msngCurveRadius_2
'        picCurve(0).Height = picCurve(0).Width
'        DrawCurve picCurve(0), curveType_TopLeft
'
'        picCurve(1).Left = Line1(3).x2 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(1).Top = Line1(3).y2
'        picCurve(1).Width = msngCurveRadius_1
'        picCurve(1).Height = picCurve(1).Width
'        DrawCurve picCurve(1), curveType_TopRight
'
'      End If
'
'    Case lineDirection_up
'      '***********************
'      '
'      '          o-----o
'      '          |     |
'      '    ^     |     |
'      '    |     |
'      '    o-----o
'      '
'      '***********************
'      If negX And negY Then
'        picCurve(1).Left = Line1(1).x1 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(1).Top = Line1(1).y1
'        picCurve(1).Width = msngCurveRadius_1
'        picCurve(1).Height = picCurve(1).Width
'        DrawCurve picCurve(1), curveType_TopRight
'
'        picCurve(0).Left = Line1(2).x1
'        picCurve(0).Top = Line1(1).y2
'        picCurve(0).Width = msngCurveRadius_2
'        picCurve(0).Height = picCurve(0).Width
'        DrawCurve picCurve(0), curveType_TopLeft
'
'        picCurve(2).Left = Line1(3).x1 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Top = Line1(2).y2 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Width = msngCurveRadius_2
'        picCurve(2).Height = picCurve(2).Width
'        DrawCurve picCurve(2), curveType_BottomRight
'
'        picCurve(3).Left = Line1(4).x1
'        picCurve(3).Top = Line1(4).y1 - msngCurveRadius_1 + PIXELWIDTH
'        picCurve(3).Width = msngCurveRadius_1
'        picCurve(3).Height = picCurve(3).Width
'        DrawCurve picCurve(3), curveType_BottomLeft
'
'
'      ElseIf (msngYOffset < (2 * MINSTARTENDLENGTH)) Then
'        '***********************
'        '
'        '    o-----o
'        '    |     |
'        '    |     |     ^
'        '          |     |
'        '          o-----o
'        '
'        '***********************
'        picCurve(0).Left = Line1(0).x2
'        picCurve(0).Top = Line1(1).y1
'        picCurve(0).Width = msngCurveRadius_2
'        picCurve(0).Height = picCurve(0).Width
'        DrawCurve picCurve(0), curveType_TopLeft
'
'        picCurve(1).Left = Line1(1).x2 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(1).Top = Line1(1).y2
'        picCurve(1).Width = msngCurveRadius_2
'        picCurve(1).Height = picCurve(1).Width
'        DrawCurve picCurve(1), curveType_TopRight
'
'        picCurve(3).Left = Line1(2).x2
'        picCurve(3).Top = Line1(2).y2 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(3).Width = msngCurveRadius_2
'        picCurve(3).Height = picCurve(3).Width
'        DrawCurve picCurve(3), curveType_BottomLeft
'
'        picCurve(2).Left = Line1(3).x2 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Top = Line1(4).y1 - msngCurveRadius_2 + PIXELWIDTH
'        picCurve(2).Width = msngCurveRadius_2
'        picCurve(2).Height = picCurve(2).Width
'        DrawCurve picCurve(2), curveType_BottomRight
'
'      End If
'
'    Case lineDirection_right
'
'    Case lineDirection_left
'
'  End Select
'
End Sub

Private Sub FormatLines()

  Dim iLoop As Integer
  Dim linTemp As Line
  Dim afLineRequired() As Boolean
  
  ' Set the line style depending on whether or not the link is highlighted.
  For Each linTemp In Line1
    linTemp.BorderStyle = IIf(mfHighlighted, vbBSDot, vbBSSolid)
  Next linTemp
  Set linTemp = Nothing
  
  ' Work out whow many lines are required.
  ReDim afLineRequired(4)
  For iLoop = 1 To UBound(afLineRequired)
    afLineRequired(iLoop) = False
  Next iLoop
  
  ' Fifth line required?
  afLineRequired(4) = _
    ((miStartDirection = lineDirection_down) And _
        (miEndDirection = lineDirection_Up) And _
        (msngYOffset < (2 * MINSTARTENDLENGTH)) And (msngXOffset <> 0)) Or _
      ((miStartDirection = lineDirection_Up) And _
        (miEndDirection = lineDirection_down) And _
        (msngYOffset > (-2 * MINSTARTENDLENGTH)) And (msngXOffset <> 0)) Or _
      ((miStartDirection = lineDirection_Left) And _
        (miEndDirection = lineDirection_Right) And _
        (msngXOffset > (-2 * MINSTARTENDLENGTH)) And (msngYOffset <> 0)) Or _
      ((miStartDirection = lineDirection_Right) And _
        (miEndDirection = lineDirection_Left) And _
        (msngXOffset < (2 * MINSTARTENDLENGTH)) And (msngYOffset <> 0))
      
  ' Fourth line required?
  afLineRequired(3) = afLineRequired(4) Or _
    ((miStartDirection = lineDirection_Left) And _
      (miEndDirection = lineDirection_down) And _
      ((msngYOffset > (-1 * MINSTARTENDLENGTH)) Or (msngXOffset > (-1 * MINSTARTENDLENGTH)))) Or _
    ((miStartDirection = lineDirection_Left) And _
      (miEndDirection = lineDirection_Up) And _
      ((msngYOffset < MINSTARTENDLENGTH) Or (msngXOffset > (-1 * MINSTARTENDLENGTH)))) Or _
    ((miStartDirection = lineDirection_Right) And _
      (miEndDirection = lineDirection_down) And _
      ((msngYOffset > (-1 * MINSTARTENDLENGTH)) Or (msngXOffset < MINSTARTENDLENGTH))) Or _
    ((miStartDirection = lineDirection_Right) And _
      (miEndDirection = lineDirection_Up) And _
      ((msngYOffset < MINSTARTENDLENGTH) Or (msngXOffset < MINSTARTENDLENGTH))) Or _
    ((miStartDirection = lineDirection_Up) And _
      (miEndDirection = lineDirection_Right) And _
      ((msngYOffset > (-1 * MINSTARTENDLENGTH)) Or (msngXOffset > (-1 * MINSTARTENDLENGTH)))) Or _
    ((miStartDirection = lineDirection_down) And _
      (miEndDirection = lineDirection_Right) And _
      ((msngYOffset < MINSTARTENDLENGTH) Or (msngXOffset > (-1 * MINSTARTENDLENGTH)))) Or _
    ((miStartDirection = lineDirection_Up) And _
      (miEndDirection = lineDirection_Left) And _
      ((msngYOffset > (-1 * MINSTARTENDLENGTH)) Or (msngXOffset < MINSTARTENDLENGTH))) Or _
    ((miStartDirection = lineDirection_down) And _
      (miEndDirection = lineDirection_Left) And _
      ((msngYOffset < MINSTARTENDLENGTH) Or (msngXOffset < MINSTARTENDLENGTH)))

  ' Third line required?
  afLineRequired(2) = afLineRequired(3) Or _
    (miStartDirection = miEndDirection) Or _
    ((miStartDirection = lineDirection_down) And _
      (miEndDirection = lineDirection_Up) And _
      (msngXOffset <> 0)) Or _
    ((miStartDirection = lineDirection_Up) And _
      (miEndDirection = lineDirection_down) And _
      (msngXOffset <> 0)) Or _
    ((miStartDirection = lineDirection_Left) And _
      (miEndDirection = lineDirection_Right) And _
      (msngYOffset <> 0)) Or _
    ((miStartDirection = lineDirection_Right) And _
      (miEndDirection = lineDirection_Left) And _
      (msngYOffset <> 0))

  ' Second line required?
  afLineRequired(1) = afLineRequired(2) Or _
    (msngXOffset <> 0) And (msngYOffset <> 0)
  
'  ' Hide the curves - these are made visible in the DrawCurve routine
'  For iLoop = 0 To picCurve.UBound Step 1
'    picCurve(iLoop).Visible = False
'  Next iLoop
'
  ' Draw the lines, and position the arrow.
  If afLineRequired(4) Then
    ' Five Lines
    mintLines = 5
    Calculate_Radii
    FormatLines_5Lines
  ElseIf afLineRequired(3) Then
    ' Four lines
    mintLines = 4
    Calculate_Radii
    FormatLines_4Lines
  ElseIf afLineRequired(2) Then
    ' Three lines
    mintLines = 3
    Calculate_Radii
    FormatLines_3Lines
  ElseIf afLineRequired(1) Then
    ' Two lines
    mintLines = 2
    Calculate_Radii
    FormatLines_2Lines
  Else
    ' One straight line from start to end
    FormatLines_1Line
  End If
  
  ' Hide the lines as required
  For iLoop = 1 To UBound(afLineRequired)
    Line1(iLoop).Visible = afLineRequired(iLoop)
  Next iLoop
 
End Sub

Private Sub FormatLines_1Line()
  
  ' One straight line from start to end
  
  '***********************************
  '
  '   |                         ^
  '   |     ---->     <----     |
  '   v                         |
  '
  '***********************************
  
  With Line1(0)
    .x1 = IIf(miStartDirection = lineDirection_Left, (-1 * msngXOffset) + msngBorder, msngBorder)
    .x2 = IIf(miStartDirection = lineDirection_Left, msngBorder, msngXOffset + msngBorder)
    .y1 = IIf(miStartDirection = lineDirection_Up, (-1 * msngYOffset) + msngBorder, msngBorder)
    .y2 = IIf(miStartDirection = lineDirection_Up, msngBorder, msngYOffset + msngBorder)
  End With
  
  ' Resize the control.
  With UserControl
    .Width = (IIf(msngXOffset < 0, -1, 1) * msngXOffset) + (2 * msngBorder)
    .Height = (IIf(msngYOffset < 0, -1, 1) * msngYOffset) + (2 * msngBorder)
  End With

  ' Position the arrow
  UserControl.Cls
  AttachArrow miEndDirection, Line1(0)

End Sub

Private Sub FormatLines_2Lines()

  ' Two lines
  
  '******************************
  '
  '   |                 |
  '   |       OR        | (negX)
  '   |                 |
  '   o---->       <----o
  '
  '******************************
  '******************************
  '
  '   o---->       <----o
  '   |       OR        |
  '   |                 | (negX)
  '   |                 |
  '
  '******************************
  '******************************
  '
  '   -----o            ^
  '        |   OR       |
  '        |            | (negY)
  '        v       -----o
  '
  '******************************
  '******************************
  '
  '   o-----       ^
  '   |        OR  |
  '   |            |      (negY)
  '   v            o-----
  '
  '******************************
  
  With Line1(0)
    If (miStartDirection = lineDirection_down) Or (miStartDirection = lineDirection_Up) Then
      .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
      .x2 = .x1
      .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
      .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
    Else
      .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
      .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
      .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
      .y2 = .y1
    End If
  End With

  With Line1(1)
    If (miEndDirection = lineDirection_down) Or (miEndDirection = lineDirection_Up) Then
      .x1 = Line1(0).x2
      .x2 = .x1
      .y1 = Line1(0).y2
      .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
    Else
      .x1 = Line1(0).x2
      .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
      .y1 = Line1(0).y2
      .y2 = .y1
    End If
  End With

  If mblnCurvedLinks Then
    JoinLines_2Lines
  End If
  
  ' Position the arrow.
  AttachArrow miEndDirection, Line1(1)

  ' Resize the control.
  With UserControl
    .Width = (IIf(msngXOffset < 0, -1, 1) * msngXOffset) + (2 * msngBorder) + 20
    .Height = (IIf(msngYOffset < 0, -1, 1) * msngYOffset) + (2 * msngBorder) + 20
  End With

End Sub

Private Sub FormatLines_3Lines()

  ' Three lines
  
  Dim iLoop As Integer
  Dim sngMaxX As Single
  Dim sngMaxY As Single

  Select Case miStartDirection
    Case lineDirection_down
      Select Case miEndDirection
        Case lineDirection_down
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = IIf(msngYOffset < 0, (-1 * msngYOffset), msngYOffset) + msngBorder + MINSTARTENDLENGTH
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          
        Case lineDirection_Up
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = msngBorder
            .y2 = (msngYOffset / 2) + msngBorder
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = msngYOffset + msngBorder
          End With
      End Select
      
    Case lineDirection_Up
      Select Case miEndDirection
        Case lineDirection_down
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = (-1 * msngYOffset) + msngBorder
            .y2 = (-1 * msngYOffset / 2) + msngBorder
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = msngBorder
          End With
        
        Case lineDirection_Up
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset), 0) + msngBorder + MINSTARTENDLENGTH
            .y2 = msngBorder
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < 0, 0, msngYOffset) + msngBorder + MINSTARTENDLENGTH
          End With
      End Select

    Case lineDirection_Right
      Select Case miEndDirection
        Case lineDirection_Right
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = IIf(msngXOffset < 0, (-1 * msngXOffset), msngXOffset) + msngBorder + MINSTARTENDLENGTH
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(1).y2
            .y2 = .y1
          End With

        Case lineDirection_Left
          With Line1(0)
            .x1 = msngBorder
            .x2 = (msngXOffset / 2) + msngBorder
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = msngXOffset + msngBorder
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
      End Select

    Case lineDirection_Left
      Select Case miEndDirection
        Case lineDirection_Right
          With Line1(0)
            .x1 = (-1 * msngXOffset) + msngBorder
            .x2 = (-1 * msngXOffset / 2) + msngBorder
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = msngBorder
            .y1 = Line1(1).y2
            .y2 = .y1
          End With

        Case lineDirection_Left
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset), 0) + msngBorder + MINSTARTENDLENGTH
            .x2 = msngBorder
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < 0, 0, msngXOffset) + msngBorder + MINSTARTENDLENGTH
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
      End Select
  End Select

  If mblnCurvedLinks Then
    JoinLines_3Lines
  End If
  
  ' Resize the control.
  sngMaxX = 0
  sngMaxY = 0
  For iLoop = 0 To 2
    With Line1(iLoop)
      sngMaxX = IIf(sngMaxX < .x1, .x1, sngMaxX)
      sngMaxX = IIf(sngMaxX < .x2, .x2, sngMaxX)
    
      sngMaxY = IIf(sngMaxY < .y1, .y1, sngMaxY)
      sngMaxY = IIf(sngMaxY < .y2, .y2, sngMaxY)
    End With
  Next iLoop
    
  With UserControl
    .Width = sngMaxX + msngBorder + 10
    .Height = sngMaxY + msngBorder + 10
  End With

  ' Position the arrow.
  AttachArrow miEndDirection, Line1(2)

End Sub

Private Sub FormatLines_4Lines()

  ' Four lines
  
  Dim iLoop As Integer
  Dim sngMaxX As Single
  Dim sngMaxY As Single

  Select Case miStartDirection
    Case lineDirection_down
      Select Case miEndDirection
        Case lineDirection_Right
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = IIf(msngYOffset < 0, (msngBorder + (-1 * msngYOffset) + MINSTARTENDLENGTH), _
              (IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder + (msngYOffset / 2), msngBorder + MINSTARTENDLENGTH)))
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder + MINSTARTENDLENGTH, msngBorder + IIf(msngXOffset < 0, 0, msngXOffset) + MINSTARTENDLENGTH)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(2).y2
            .y2 = .y1
          End With
        
        Case lineDirection_Left
          With Line1(0)
            .x1 = IIf(msngXOffset > MINSTARTENDLENGTH, msngBorder, msngBorder + MINSTARTENDLENGTH - msngXOffset)
            .x2 = .x1
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = IIf(msngYOffset < 0, (msngBorder + (-1 * msngYOffset) + MINSTARTENDLENGTH), _
              (IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder + (msngYOffset / 2), msngBorder + MINSTARTENDLENGTH)))
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < MINSTARTENDLENGTH, msngBorder, msngBorder + msngXOffset - MINSTARTENDLENGTH)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = IIf(msngXOffset < MINSTARTENDLENGTH, msngBorder + MINSTARTENDLENGTH, msngXOffset + msngBorder)
            .y1 = Line1(2).y2
            .y2 = .y1
          End With
      End Select

    Case lineDirection_Up
      Select Case miEndDirection
        Case lineDirection_Right
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = .x1
            .y1 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngYOffset, msngBorder + MINSTARTENDLENGTH)
            .y2 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - (msngYOffset / 2), msngBorder)
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < 0, msngBorder + MINSTARTENDLENGTH, msngXOffset + msngBorder + MINSTARTENDLENGTH)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder, msngBorder + MINSTARTENDLENGTH + msngYOffset)
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(2).y2
            .y2 = .y1
          End With

        Case lineDirection_Left
          With Line1(0)
            .x1 = IIf(msngXOffset > MINSTARTENDLENGTH, msngBorder, MINSTARTENDLENGTH - msngXOffset + msngBorder)
            .x2 = .x1
            .y1 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngYOffset, msngBorder + MINSTARTENDLENGTH)
            .y2 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - (msngYOffset / 2), msngBorder)
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = IIf(msngXOffset < MINSTARTENDLENGTH, msngBorder, msngXOffset + msngBorder - MINSTARTENDLENGTH)
            .y1 = Line1(0).y2
            .y2 = .y1
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = .x1
            .y1 = Line1(1).y2
            .y2 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder, msngBorder + msngYOffset + MINSTARTENDLENGTH)
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = IIf(msngXOffset > MINSTARTENDLENGTH, msngBorder + msngXOffset, MINSTARTENDLENGTH + msngBorder)
            .y1 = Line1(2).y2
            .y2 = .y1
          End With
      End Select

    Case lineDirection_Right
      Select Case miEndDirection
        Case lineDirection_down
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = IIf(XOffset < MINSTARTENDLENGTH, .x1 + MINSTARTENDLENGTH, msngBorder + (XOffset / 2))
            .y1 = IIf(msngYOffset < 0, (-1 * msngYOffset) + msngBorder, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < 0, msngBorder + MINSTARTENDLENGTH, msngBorder + MINSTARTENDLENGTH + msngYOffset)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = .x1
            .y1 = Line1(2).y2
            .y2 = IIf(msngYOffset < 0, msngBorder, msngYOffset + msngBorder)
          End With

        Case lineDirection_Up
          With Line1(0)
            .x1 = IIf(msngXOffset < 0, (-1 * msngXOffset) + msngBorder, msngBorder)
            .x2 = IIf(XOffset < MINSTARTENDLENGTH, .x1 + MINSTARTENDLENGTH, msngBorder + (XOffset / 2))
            .y1 = IIf(msngYOffset < MINSTARTENDLENGTH, msngBorder + MINSTARTENDLENGTH - msngYOffset, msngBorder)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset < MINSTARTENDLENGTH, msngBorder, msngBorder + (msngYOffset / 2))
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < 0, msngBorder, msngXOffset + msngBorder)
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = .x1
            .y1 = Line1(2).y2
            .y2 = IIf(msngYOffset < MINSTARTENDLENGTH, msngBorder + MINSTARTENDLENGTH, msngBorder + msngYOffset)
          End With
      End Select

    Case lineDirection_Left
      Select Case miEndDirection
        Case lineDirection_down
          With Line1(0)
            .x1 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngXOffset, MINSTARTENDLENGTH + msngBorder)
            .x2 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder - (msngXOffset / 2), msngBorder)
            .y1 = IIf(msngYOffset >= 0, msngBorder, IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngYOffset, msngBorder + MINSTARTENDLENGTH + msngYOffset))
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset > 0, msngBorder + MINSTARTENDLENGTH + msngYOffset, msngBorder + MINSTARTENDLENGTH)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder, msngBorder + MINSTARTENDLENGTH + msngXOffset)
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = .x1
            .y1 = Line1(2).y2
            .y2 = IIf(msngYOffset > 0, msngBorder + msngYOffset, msngBorder)
          End With

        Case lineDirection_Up
          With Line1(0)
            .x1 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngXOffset, MINSTARTENDLENGTH + msngBorder)
            .x2 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder - (msngXOffset / 2), msngBorder)
            .y1 = IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder, msngBorder + MINSTARTENDLENGTH - msngYOffset)
            .y2 = .y1
          End With
          With Line1(1)
            .x1 = Line1(0).x2
            .x2 = .x1
            .y1 = Line1(0).y2
            .y2 = IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder + msngYOffset - MINSTARTENDLENGTH, msngBorder)
          End With
          With Line1(2)
            .x1 = Line1(1).x2
            .x2 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder, msngBorder + MINSTARTENDLENGTH + msngXOffset)
            .y1 = Line1(1).y2
            .y2 = .y1
          End With
          With Line1(3)
            .x1 = Line1(2).x2
            .x2 = .x1
            .y1 = Line1(2).y2
            .y2 = IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder + msngYOffset, msngBorder + MINSTARTENDLENGTH)
          End With
      End Select
  End Select

  If mblnCurvedLinks Then
    JoinLines_4Lines
  End If
  
  ' Position the arrow.
  AttachArrow miEndDirection, Line1(3)

  ' Resize the control.
  sngMaxX = 0
  sngMaxY = 0
  For iLoop = 0 To 3
    With Line1(iLoop)
      sngMaxX = IIf(sngMaxX < .x1, .x1, sngMaxX)
      sngMaxX = IIf(sngMaxX < .x2, .x2, sngMaxX)
    
      sngMaxY = IIf(sngMaxY < .y1, .y1, sngMaxY)
      sngMaxY = IIf(sngMaxY < .y2, .y2, sngMaxY)
    End With
  Next iLoop
  
  With UserControl
    .Width = sngMaxX + msngBorder + 10
    .Height = sngMaxY + msngBorder + 10
  End With

End Sub

Private Sub FormatLines_5Lines()

  ' Five lines
  
  Dim iLoop As Integer
  Dim sngMaxX As Single
  Dim sngMaxY As Single
 
  Select Case miStartDirection
    Case lineDirection_down
      With Line1(0)
        .x1 = IIf(msngXOffset > 0, msngBorder, msngBorder - msngXOffset)
        .x2 = .x1
        .y1 = IIf(msngYOffset > MINSTARTENDLENGTH, msngBorder, msngBorder + MINSTARTENDLENGTH - msngYOffset)
        .y2 = .y1 + MINSTARTENDLENGTH
      End With
      With Line1(1)
        .x1 = Line1(0).x2
        .x2 = msngBorder + (IIf(msngXOffset < 0, (-1 * msngXOffset), msngXOffset) / 2)
        .y1 = Line1(0).y2
        .y2 = .y1
      End With
      With Line1(2)
        .x1 = Line1(1).x2
        .x2 = .x1
        .y1 = Line1(1).y2
        .y2 = .y1 - ((2 * MINSTARTENDLENGTH) - msngYOffset)
      End With
      With Line1(3)
        .x1 = Line1(2).x2
        .x2 = Line1(0).x1 + msngXOffset
        .y1 = Line1(2).y2
        .y2 = .y1
      End With
      With Line1(4)
        .x1 = Line1(3).x2
        .x2 = .x1
        .y1 = Line1(2).y2
        .y2 = Line1(0).y1 + msngYOffset
      End With
  
    Case lineDirection_Up
      With Line1(0)
        .x1 = IIf(msngXOffset > 0, msngBorder, msngBorder - msngXOffset)
        .x2 = .x1
        .y1 = IIf(msngYOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngYOffset, msngBorder + MINSTARTENDLENGTH)
        .y2 = .y1 - MINSTARTENDLENGTH
      End With
      With Line1(1)
        .x1 = Line1(0).x2
        .x2 = msngBorder + (IIf(msngXOffset < 0, (-1 * msngXOffset), msngXOffset) / 2)
        .y1 = Line1(0).y2
        .y2 = .y1
      End With
      With Line1(2)
        .x1 = Line1(1).x2
        .x2 = .x1
        .y1 = Line1(1).y2
        .y2 = .y1 + ((2 * MINSTARTENDLENGTH) + msngYOffset)
      End With
      With Line1(3)
        .x1 = Line1(2).x2
        .x2 = Line1(0).x1 + msngXOffset
        .y1 = Line1(2).y2
        .y2 = .y1
      End With
      With Line1(4)
        .x1 = Line1(3).x2
        .x2 = .x1
        .y1 = Line1(2).y2
        .y2 = Line1(0).y1 + msngYOffset
      End With

    Case lineDirection_Right
      With Line1(0)
        .x1 = IIf(msngXOffset > MINSTARTENDLENGTH, msngBorder, msngBorder + MINSTARTENDLENGTH - msngXOffset)
        .x2 = .x1 + MINSTARTENDLENGTH
        .y1 = IIf(msngYOffset > 0, msngBorder, msngBorder - msngYOffset)
        .y2 = .y1
      End With
      With Line1(1)
        .x1 = Line1(0).x2
        .x2 = .x1
        .y1 = Line1(0).y2
        .y2 = msngBorder + (IIf(msngYOffset < 0, (-1 * msngYOffset), msngYOffset) / 2)
      End With
      With Line1(2)
        .x1 = Line1(1).x2
        .x2 = .x1 - ((2 * MINSTARTENDLENGTH) - msngXOffset)
        .y1 = Line1(1).y2
        .y2 = .y1
      End With
      With Line1(3)
        .x1 = Line1(2).x2
        .x2 = .x1
        .y1 = Line1(2).y2
        .y2 = Line1(0).y1 + msngYOffset
      End With
      With Line1(4)
        .x1 = Line1(3).x2
        .x2 = Line1(0).x1 + msngXOffset
        .y1 = Line1(3).y2
        .y2 = .y1
      End With

    Case lineDirection_Left
      With Line1(0)
        .x1 = IIf(msngXOffset < (-1 * MINSTARTENDLENGTH), msngBorder - msngXOffset, msngBorder + MINSTARTENDLENGTH)
        .x2 = .x1 - MINSTARTENDLENGTH
        .y1 = IIf(msngYOffset > 0, msngBorder, msngBorder - msngYOffset)
        .y2 = .y1
      End With
      With Line1(1)
        .x1 = Line1(0).x2
        .x2 = .x1
        .y1 = Line1(0).y2
        .y2 = msngBorder + (IIf(msngYOffset < 0, (-1 * msngYOffset), msngYOffset) / 2)
      End With
      With Line1(2)
        .x1 = Line1(1).x2
        .x2 = .x1 + ((2 * MINSTARTENDLENGTH) + msngXOffset)
        .y1 = Line1(1).y2
        .y2 = .y1
      End With
      With Line1(3)
        .x1 = Line1(2).x2
        .x2 = .x1
        .y1 = Line1(2).y2
        .y2 = Line1(0).y1 + msngYOffset
      End With
      With Line1(4)
        .x1 = Line1(2).x2
        .x2 = Line1(0).x1 + msngXOffset
        .y1 = Line1(3).y2
        .y2 = .y1
      End With
  End Select
  
  If mblnCurvedLinks Then
    JoinLines_5Lines
  End If
  
  ' Position the arrow.
  AttachArrow miEndDirection, Line1(4)

  ' Resize the control.
  sngMaxX = 0
  sngMaxY = 0
  For iLoop = 0 To 4
    With Line1(iLoop)
      sngMaxX = IIf(sngMaxX < .x1, .x1, sngMaxX)
      sngMaxX = IIf(sngMaxX < .x2, .x2, sngMaxX)
  
      sngMaxY = IIf(sngMaxY < .y1, .y1, sngMaxY)
      sngMaxY = IIf(sngMaxY < .y2, .y2, sngMaxY)
    End With
  Next iLoop

  With UserControl
    .Width = sngMaxX + msngBorder + 10
    .Height = sngMaxY + msngBorder + 10
  End With

End Sub

Public Property Get Highlighted() As Boolean
  ' Return the 'highlighted' property.
  Highlighted = mfHighlighted
  
End Property

Public Property Let Highlighted(ByVal pfNewValue As Boolean)
  ' Set the 'highlighted' property.
  mfHighlighted = pfNewValue
  PropertyChanged "Highlighted"
  
  FormatLines

End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the 'About' box.
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
    
End Sub

Private Sub ASRWFLinkArrow1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub ASRWFLinkArrow1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub ASRWFLinkArrow1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Public Property Get StartDirection() As LineDirection
  ' Return the current start direction.
  StartDirection = miStartDirection

End Property

Public Property Get EndDirection() As LineDirection
  ' Return the current end direction.
  EndDirection = miEndDirection

End Property

Public Property Let StartDirection(ByVal piNewValue As LineDirection)
  ' Set the current StartDirection.
  miStartDirection = piNewValue
  PropertyChanged "StartDirection"

  FormatLines
  
End Property

Public Property Let EndDirection(ByVal piNewValue As LineDirection)
  ' Set the current EndDirection.
  miEndDirection = piNewValue
  PropertyChanged "EndDirection"

  FormatLines
  
End Property

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  On Error Resume Next

  ' Read the previous set of properties.
  StartDirection = PropBag.ReadProperty("StartDirection", lineDirection_down)
  EndDirection = PropBag.ReadProperty("EndDirection", lineDirection_Up)
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  XOffset = PropBag.ReadProperty("XOffset", 0)
  YOffset = PropBag.ReadProperty("YOffset", 0)

  AppMajor = PropBag.ReadProperty("AppMajor", 3)
  AppMinor = PropBag.ReadProperty("AppMinor", 5)
  AppRevision = PropBag.ReadProperty("AppRevision", 0)
End Sub

Private Sub UserControl_Resize()
  'UserControl.Cls
 ' UserControl.Refresh
  FormatLines
End Sub

'Private Sub UserControl_Resize()
''  picCurve(0).Refresh
''  picCurve(1).Refresh
''  picCurve(2).Refresh
''  picCurve(3).Refresh
'
'UserControl.ForeColor = vbGreen
'UserControl.BackColor = vbRed
'
'  FormatLines
'
'
'
'End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("StartDirection", miStartDirection, lineDirection_down)
  Call PropBag.WriteProperty("EndDirection", miEndDirection, lineDirection_Up)
  Call PropBag.WriteProperty("Highlighted", mfHighlighted, False)
  Call PropBag.WriteProperty("XOffset", msngXOffset, 0)
  Call PropBag.WriteProperty("YOffset", msngYOffset, 0)

  Call PropBag.WriteProperty("AppMajor", miAppMajor, 3)
  Call PropBag.WriteProperty("AppMinor", miAppMinor, 5)
  Call PropBag.WriteProperty("AppRevision", miAppRevision, 0)
End Sub

Public Property Get XOffset() As Single
  ' Return the XOffset.
  ' ie. the horizontal distance between the start position and the end position.
  XOffset = msngXOffset
  
End Property

Public Property Get YOffset() As Single
  ' Return the YOffset.
  ' ie. the vertical distance between the start position and the end position.
  YOffset = msngYOffset
  
End Property

Public Property Get Border() As Single
  Border = msngBorder
  
End Property

Public Property Get StartElementIndex() As Integer
  ' Return the index of the associated start element.
  StartElementIndex = miStartElementIndex

End Property

Public Property Get LineCoordinates() As Variant
  ' Return an array of the coordinates of the lines.
  ' Index 0 = X1
  ' Index 1 = X2
  ' Index 2 = Y1
  ' Index 3 = Y2
  ' NB. Only visible lines are included in the array.
  ' NB. coordinates are 'within' the user control.
  Dim asngCoordinates() As Single
  Dim linTemp As Line
  Dim iCount As Integer
  
  ReDim asngCoordinates(3, 0)
  
  iCount = -1
  For Each linTemp In Line1
    If linTemp.Visible Then
      iCount = iCount + 1
      
      ReDim Preserve asngCoordinates(3, iCount)
      asngCoordinates(0, iCount) = linTemp.x1
      asngCoordinates(1, iCount) = linTemp.x2
      asngCoordinates(2, iCount) = linTemp.y1
      asngCoordinates(3, iCount) = linTemp.y2
    End If
  Next linTemp
  Set linTemp = Nothing
  
  LineCoordinates = asngCoordinates
  
End Property

Public Property Get StartXOffset() As Single
  ' Return the XOffset of the start of the line within the usercontrol.
  StartXOffset = Line1(0).x1

End Property

Public Property Get EndXOffset() As Single
  ' Return the XOffset of the end of the line within the usercontrol.
  Dim iLoop As Integer
  
  For iLoop = 4 To 0 Step -1
    If Line1(iLoop).Visible Then
      EndXOffset = Line1(iLoop).x2
    End If
  Next iLoop
  
End Property

Public Property Get EndYOffset() As Single
  ' Return the YOffset of the end of the line within the usercontrol.
  Dim iLoop As Integer
  
  For iLoop = 4 To 0 Step -1
    If Line1(iLoop).Visible Then
      EndYOffset = Line1(iLoop).y2
    End If
  Next iLoop
  
End Property

Public Property Get StartYOffset() As Single
  ' Return the YOffset of the start of the line within the usercontrol.
  StartYOffset = Line1(0).y1

End Property

Public Property Get EndElementIndex() As Integer
  ' Return the index of the associated end element.
  EndElementIndex = miEndElementIndex

End Property

Public Property Get ArrowPicture() As StdPicture
  ' Return the arrow picture.
'  Set ArrowPicture = ASRWFLinkArrow1.ArrowPicture

End Property

Public Property Get ArrowHorizontalPosition() As Single
  ' Return the arrow's Horizontal position
 ' ArrowHorizontalPosition = ASRWFLinkArrow1.Left
  
End Property

Public Property Get ArrowVerticalPosition() As Single
  ' Return the arrow's Vertical position
 ' ArrowVerticalPosition = ASRWFLinkArrow1.Top
  
End Property

Public Property Let StartElementIndex(ByVal piNewValue As Integer)
  ' Set the index of the associated start element.
  miStartElementIndex = piNewValue

End Property

Public Property Let EndElementIndex(ByVal piNewValue As Integer)
  ' Set the index of the associated end element.
  miEndElementIndex = piNewValue

End Property

Public Property Get StartOutboundFlowCode() As Integer
  ' Return the StartOutboundFlowCode
  StartOutboundFlowCode = miStartOutboundFlowCode
End Property

Public Property Let StartOutboundFlowCode(ByVal piNewValue As Integer)
  ' Set the StartOutboundFlowCode
  miStartOutboundFlowCode = piNewValue
End Property

Private Property Let BackColour(ByVal NewValue As OLE_COLOR)
'  Dim curve As PictureBox
  
'  For Each curve In picCurve
'    curve.BackColor = NewValue
'  Next
  UserControl.BackColor = NewValue
End Property

Private Sub SetBackColour()
  If ((miAppMajor > 3) Or ((miAppMajor = 3) And (miAppMinor > 5))) Then
    BackColour = vbInactiveTitleBar
  Else
    BackColour = vbInactiveTitleBarText
  End If
End Sub

' Attaches and arrow to the given line
Private Sub AttachArrow(ByRef miEndDirection As LineDirection, ByRef TheLine As Line)

  Select Case miEndDirection
    Case lineDirection_down, lineDirection_Right
'      DrawArrow45 TheLine.x2, TheLine.y2, TheLine.x1, TheLine.y1, 100, True
'lnArrowHead(0).

'TheLine.Visible = False
'      lnArrowHead(0).x1 = TheLine.x2 - 100
'      lnArrowHead(0).x2 = TheLine.x2
'      lnArrowHead(0).y1 = TheLine.y2 - 100
'      lnArrowHead(0).y2 = TheLine.y2
''lnArrowHead(0).Visible = False
'lnArrowHead(1).Visible = False
''


    Case lineDirection_Up, lineDirection_Left
   '   DrawArrow45 TheLine.x1, TheLine.y1, TheLine.x2, TheLine.y2, 100, True
      lnArrowHead(0).x1 = TheLine.x2 - 100
      lnArrowHead(0).x2 = TheLine.x2
      lnArrowHead(0).y1 = TheLine.y2 - 100
      lnArrowHead(0).y2 = TheLine.y2
      lnArrowHead(1).x1 = TheLine.x2 + 100
      lnArrowHead(1).x2 = TheLine.x2
      lnArrowHead(1).y1 = TheLine.y2 - 100
      lnArrowHead(1).y2 = TheLine.y2
  End Select




'Set mobjArrowHead = New PictureBox
'mobjArrowHead.Line (300, 300)-(10, 10)
'
''mobjLine.x1 = 10
''mobjLine.x2 = 300
''mobjLine.y1 = 10
''mobjLine.y2 = 300
'UserControl.Controls.Add mobjArrowHead
'


End Sub


' Draw an arrow head at (x2, y2) for a vector from (x1, y1). Draw a tail at (x1, y1). Make the barbs
' and fletchings point at a 45 degree angle away from the shaft.
Private Sub DrawArrow45(ByVal x1 As _
    Single, ByVal y1 As Single, ByVal x2 As Single, ByVal _
    y2 As Single, ByVal length As Single, Optional ByVal _
    draw_head As Boolean = True, Optional ByVal _
    num_fletchings As Integer = 0, Optional ByVal _
    fletching_gap As Single = 60)

    Dim vx As Single
    Dim vy As Single
    Dim dist As Single
    Dim ax As Single
    Dim ay As Single
    Dim barb_scale As Single
    Dim i As Integer
    Dim gap_dx As Single
    Dim gap_dy As Single

    ' Draw the shaft.
'    UserControl.Line (x1, y1)-(x2, y2)

    ' Find the arrow shaft unit vector.
    vx = x2 - x1
    vy = y2 - y1
    dist = Sqr(vx * vx + vy * vy)
    
    If dist = 0 Then Exit Sub
    
    vx = vx / dist
    vy = vy / dist

    ' See if we need to draw multiple fletchings.
    If num_fletchings > 1 Then
        ' Get the fletching spacing vector.
        gap_dx = vx * fletching_gap
        gap_dy = vy * fletching_gap
    End If

    ' Draw the right barb.
    ax = -vy - vx
    ay = vx - vy

    ' Set the proper length.
    ax = ax * length
    ay = ay * length
    
    If draw_head Then
'      UserControl.Line (x2, y2)-Step(ax, ay)
    End If
    
    For i = 0 To num_fletchings - 1
'        UserControl.Line (x1 + i * gap_dx, y1 + i * gap_dy)-Step(ax, ay)
    Next i

    ' Find the left barb.
    ax = (vy - vx) * length
    ay = (-vx - vy) * length
    
    If draw_head Then
'      UserControl.Line (x2, y2)-Step(ax, ay)


'lnArrowHead(1).x1 = x1
'lnArrowHead(1).x2 = x2 - 100
'lnArrowHead(1).y1 = y1
'lnArrowHead(1).y2 = y2 - 100


    
    End If
    
    For i = 0 To num_fletchings - 1
        UserControl.Line (x1 + i * gap_dx, y1 + i * _
            gap_dy)-Step(ax, ay)
    Next i
    
End Sub


