VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinSizeInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Property Get Collection() As Collection
  Set Collection = mCol
End Property

Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
End Property

Public Function Add( _
  lhWnd As Long, MinX As Long, MinY As Long, _
  MaxX As Long, MaxY As Long) As CWinSizeInfo
  
  ' Add a new member to the collection.
  Dim sKey As String
  Dim objNewMember As CWinSizeInfo
  
  Set objNewMember = New CWinSizeInfo
  
  ' Initialise the privileges.
  With objNewMember
    .hWnd = lhWnd
    .MinX = TwipsToPixels(MinX)
    .MinY = TwipsToPixels(MinY)
    .MaxX = TwipsToPixels(MaxX)
    .MaxY = TwipsToPixels(MaxY)
    
    sKey = .hWnd
  End With
  
  mCol.Add objNewMember, sKey
    
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function

Public Property Get Item(ByRef shWndKey As String) As CWinSizeInfo
  Set Item = mCol(shWndKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(shWndKey As String)
    mCol.Remove shWndKey
End Sub

Public Function Exists(ByRef shWndKey As String) As Boolean
  ' Return TRUE if the given key exists in the collection.
  Dim Item As CWinSizeInfo
  
  On Error GoTo err_Exists
  
  Set Item = mCol(shWndKey)
  Exists = True
  Set Item = Nothing
  
  Exit Function
  
err_Exists:
  Exists = False
  
End Function

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
