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

