Attribute VB_Name = "modLink"
Option Explicit

Public Enum CurveType
  curveType_TopLeft = 0
  curveType_TopRight = 1
  curveType_BottomRight = 2
  curveType_BottomLeft = 3
End Enum

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long

Public P(0 To 3) As POINTAPI

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut         As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000



Public Sub DrawCurve(ByRef pCanvas As PictureBox, ByVal pCurveType As CurveType, Optional ByVal pColor As Long)

  With pCanvas
    .Visible = True
    .Enabled = True
    .AutoRedraw = True
    .ScaleMode = vbPixels
    .BorderStyle = 0
    .ForeColor = pColor
    .DrawStyle = IIf(pCanvas.Parent.Highlighted, vbBSSolid, vbBSSolid)
    .ZOrder 0
    .ClipControls = True
    .Cls
    
    Select Case pCurveType
      Case CurveType.curveType_TopLeft
        ' Point 1
        P(0).X = 0
        P(0).Y = .ScaleHeight
        ' Point 2
        P(1).X = 0
        P(1).Y = (.ScaleHeight / 2)
        ' Point 3
        P(2).X = (.ScaleWidth / 2)
        P(2).Y = 0
        ' Point 4
        P(3).X = .ScaleWidth
        P(3).Y = 0
      
      Case CurveType.curveType_TopRight
        ' Point 1
        P(0).X = -1
        P(0).Y = 0
        ' Point 2
        P(1).X = (.ScaleWidth / 2)
        P(1).Y = 0
        ' Point 3
        P(2).X = .ScaleWidth
        P(2).Y = (.ScaleHeight / 2)
        ' Point 4
        P(3).X = .ScaleWidth
        P(3).Y = .ScaleHeight + 6
      
      Case CurveType.curveType_BottomRight
        ' Point 1
        P(0).X = .ScaleWidth
        P(0).Y = -10
        ' Point 2
        P(1).X = .ScaleWidth
        P(1).Y = (.ScaleHeight / 2)
        ' Point 3
        P(2).X = (.ScaleWidth / 2)
        P(2).Y = .ScaleHeight
        ' Point 4
        P(3).X = -10
        P(3).Y = .ScaleHeight
      
      Case CurveType.curveType_BottomLeft
        ' Point 1
        P(0).X = .ScaleWidth + 4
        P(0).Y = .ScaleHeight
        ' Point 2
        P(1).X = (.ScaleWidth / 2)
        P(1).Y = .ScaleHeight
        ' Point 3
        P(2).X = 0
        P(2).Y = (.ScaleHeight / 2)
        ' Point 4
        P(3).X = 0
        P(3).Y = -1

    End Select
         
    Call PolyBezier(.hdc, P(0), 4)
  
  End With
  
End Sub


' transparency stuff?


'
'Public Function Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = vbBlack, _
'    Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
'' Return : True if there is no error.
'' hWnd   : hWnd of the window to make transparent
'' Col : Color to make transparent if TrMode=False
'' PcTransp  : 0 Ã  255 >> 0 = transparent  -:- 255 = Opaque
'Dim DisplayStyle As Long
'Dim VoirStyle As Long
'
'    VoirStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
'    If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
'        DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
'        Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
'    End If
'    Transparency = (SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
'
'    If Not Err.Number = 0 Then Err.Clear
'End Function
'
'Public Sub ActiveTransparency(M As Object, d As Boolean, F As Boolean, _
'     T_Transparency As Integer, Optional Color As Long)
'Dim B As Boolean
'        If d And F Then
'        'Makes color (here the background color of the shape) transparent
'        'upon value of T_Transparency
'            B = Transparency(M.hWnd, Color, T_Transparency, False)
'        ElseIf d Then
'            'Makes form, including all components, transparent
'            'upon value of T_Transparency
'            B = Transparency(M.hWnd, 0, T_Transparency, True)
'        Else
'            'Restores the form opaque.
'            B = Transparency(M.hWnd, , 255, True)
'        End If
'End Sub
'


