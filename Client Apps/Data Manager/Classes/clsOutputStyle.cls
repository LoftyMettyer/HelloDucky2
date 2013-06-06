VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngStartCol As Long
Private mlngStartRow As Long
Private mlngEndCol As Long
Private mlngEndRow As Long
Private mblnGridlines As Boolean
Private mblnBold As Boolean
Private mblnUnderLine As Boolean
Private mblnOmitSpacerCol As Boolean
Private mblnOmitSpacerRow As Boolean
Private mlngBackCol As Long
Private mlngForeCol As Long
Private mblnCenterText As Boolean
Private mstrName As String
Private mstrFontName As String
Private mbItalic As Boolean
Private miFontSize As Integer

Public Property Get StartCol() As Long
  StartCol = mlngStartCol
End Property

Public Property Let StartCol(ByVal lngNewValue As Long)
  mlngStartCol = lngNewValue
End Property

Public Property Get StartRow() As Long
  StartRow = mlngStartRow
End Property

Public Property Let StartRow(ByVal lngNewValue As Long)
  mlngStartRow = lngNewValue
End Property

Public Property Get EndCol() As Long
  EndCol = mlngEndCol
End Property

Public Property Let EndCol(ByVal lngNewValue As Long)
  mlngEndCol = lngNewValue
End Property

Public Property Get EndRow() As Long
  EndRow = mlngEndRow
End Property

Public Property Let EndRow(ByVal lngNewValue As Long)
  mlngEndRow = lngNewValue
End Property

Public Property Get Gridlines() As Boolean
  Gridlines = mblnGridlines
End Property

Public Property Let Gridlines(ByVal blnNewValue As Boolean)
  mblnGridlines = blnNewValue
End Property

Public Property Get Bold() As Boolean
  Bold = mblnBold
End Property

Public Property Let Bold(ByVal blnNewValue As Boolean)
  mblnBold = blnNewValue
End Property

Public Property Get OmitSpacerCol() As Boolean
  OmitSpacerCol = mblnOmitSpacerCol
End Property

Public Property Let OmitSpacerCol(ByVal blnNewValue As Boolean)
  mblnOmitSpacerCol = blnNewValue
End Property
Public Property Get OmitSpacerRow() As Boolean
  OmitSpacerRow = mblnOmitSpacerRow
End Property

Public Property Let OmitSpacerRow(ByVal blnNewValue As Boolean)
  mblnOmitSpacerRow = blnNewValue
End Property
Public Property Get Underline() As Boolean
  Underline = mblnUnderLine
End Property

Public Property Let Underline(ByVal blnNewValue As Boolean)
  mblnUnderLine = blnNewValue
End Property

Public Property Get ForeCol() As Long
  ForeCol = mlngForeCol
End Property

Public Property Let ForeCol(ByVal lngNewValue As Long)
  mlngForeCol = lngNewValue
End Property

Public Property Get BackCol() As Long
  BackCol = mlngBackCol
End Property

Public Property Let BackCol(ByVal lngNewValue As Long)
  mlngBackCol = lngNewValue
End Property

Public Property Get CenterText() As Boolean
  CenterText = mblnCenterText
End Property

Public Property Let CenterText(ByVal blnNewValue As Boolean)
  mblnCenterText = blnNewValue
End Property

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(ByVal strNewValue As String)
  mstrName = strNewValue
End Property

Public Property Get Italic() As Boolean
  Italic = mbItalic
End Property

Public Property Let Italic(ByVal bNewValue As Boolean)
  mbItalic = bNewValue
End Property

Public Property Get FontName() As String
  FontName = mstrFontName
End Property

Public Property Let FontName(ByVal strNewValue As String)
  mstrFontName = strNewValue
End Property

Public Property Get FontSize() As Integer
  FontSize = miFontSize
End Property

Public Property Let FontSize(ByVal iNewValue As Integer)
  miFontSize = iNewValue
End Property
