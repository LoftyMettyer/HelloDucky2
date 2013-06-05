VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngScreenID As Long
Private msScreenName As String
Private mlngTableID As Long
Private mlngOrderID As Long
Private mlngHeight As Long
Private mlngWidth As Long
Private mlngPictureID As Long
Private mfQuickEntry As Boolean

Private msFontName As String
Private miFontSize As Integer
Private mfFontBold As Boolean
Private mfFontItalic As Boolean
Private mfFontStrikethru As Boolean
Private mfFontUnderline As Boolean

Private masTabCaptions() As String
Public Property Get ScreenID() As Long
  ScreenID = mlngScreenID
  
End Property

Public Property Let ScreenID(ByVal plngNewValue As Long)
  mlngScreenID = plngNewValue
  
End Property

Public Property Get ScreenName() As String
  ScreenName = msScreenName
  
End Property

Public Property Let ScreenName(ByVal psNewValue As String)
  msScreenName = psNewValue
  
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property

Public Property Let PictureID(ByVal plngNewValue As Long)
  mlngPictureID = plngNewValue
  
End Property
Public Property Get PictureID() As Long
  PictureID = mlngPictureID
  
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue
  
End Property

Public Property Get OrderID() As Long
  OrderID = mlngOrderID
  
End Property

Public Property Let OrderID(ByVal plngNewValue As Long)
  mlngOrderID = plngNewValue
  
End Property

Public Property Get Height() As Long
  Height = mlngHeight
  
End Property

Public Property Let Height(ByVal plngNewValue As Long)
  mlngHeight = plngNewValue
  
End Property

Public Property Get Width() As Long
  Width = mlngWidth
  
End Property

Public Property Let Width(ByVal plngNewValue As Long)
  mlngWidth = plngNewValue
  
End Property

Public Property Get QuickEntry() As Boolean
  QuickEntry = mfQuickEntry
  
End Property

Public Property Let QuickEntry(ByVal pfNewValue As Boolean)
  mfQuickEntry = pfNewValue
  
End Property

Public Property Get FontName() As String
  FontName = msFontName
  
End Property

Public Property Let FontName(ByVal psNewValue As String)
  msFontName = psNewValue
  
End Property

Public Property Get FontSize() As Integer
  FontSize = miFontSize
  
End Property

Public Property Let FontSize(ByVal piNewValue As Integer)
  miFontSize = piNewValue
  
End Property

Public Property Get FontBold() As Boolean
  FontBold = mfFontBold
  
End Property

Public Property Let FontBold(ByVal pfNewValue As Boolean)
  mfFontBold = pfNewValue
  
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = mfFontItalic
  
End Property

Public Property Let FontItalic(ByVal pfNewValue As Boolean)
  mfFontItalic = pfNewValue
  
End Property

Public Property Get FontStrikethru() As Boolean
  FontStrikethru = mfFontStrikethru
  
End Property

Public Property Let FontStrikethru(ByVal pfNewValue As Boolean)
  mfFontStrikethru = pfNewValue
  
End Property

Public Property Get FontUnderline() As Boolean
  FontUnderline = mfFontUnderline
  
End Property

Public Property Let FontUnderline(ByVal pfNewValue As Boolean)
  mfFontUnderline = pfNewValue
  
End Property

Public Property Get TabCaptions() As Variant
  TabCaptions = masTabCaptions
  
End Property

Public Property Let TabCaptions(ByVal pvNewValue As Variant)
  masTabCaptions = pvNewValue
  
End Property

Public Sub ClearTabCaptions()
  ReDim masTabCaptions(0)
  
End Sub
