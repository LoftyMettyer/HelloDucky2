VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(plngScreenID As Long, psScreenName As String, plngTableID As Long, _
  plngOrderID As Long, plngHeight As Long, plngWidth As Long, _
  plngPictureID As Long, pfQuickEntry As Boolean, _
  psFontName As String, piFontSize As Integer, pfFontBold As Boolean, _
  pfFontItalic As Boolean, pfFontStrikeThru As Boolean, pfFontUnderline As Boolean) As clsScreen
  ' Add a new member to the collection of screens.
  Dim objNewMember As New clsScreen
    
  ' Initialise the privileges.
  With objNewMember
    .ScreenID = plngScreenID
    .ScreenName = psScreenName
    .TableID = plngTableID
    .OrderID = plngOrderID
    .Height = plngHeight
    .Width = plngWidth
    .PictureID = plngPictureID
    .QuickEntry = pfQuickEntry
    
    .FontName = psFontName
    .FontSize = piFontSize
    .FontBold = pfFontBold
    .FontItalic = pfFontItalic
    .FontStrikethru = pfFontStrikeThru
    .FontUnderline = pfFontUnderline
    
    ' Initailise the tab caption array.
    .ClearTabCaptions
  End With
  
  mCol.Add objNewMember, Trim(Str(plngScreenID))
    
  Set Add = objNewMember
  Set objNewMember = Nothing
                    
End Function


Public Property Get Count() As Long
  Count = mCol.Count
    
End Property


Public Property Get Item(vIndexKey As Variant) As clsScreen
  Set Item = mCol(vIndexKey)
    
End Property


Public Sub Remove(vIndexKey As Variant)
  mCol.Remove vIndexKey
    
End Sub


Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing

End Sub


