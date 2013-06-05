VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Cancelled()
Public Event Completed()

Private WithEvents mobjProgress As COA_Progress

Private Sub Class_Initialize()
  If InStr(LCase(Command$), "/progressbar=false") = 0 Then
    Set mobjProgress = New COA_Progress
  End If
End Sub

Private Sub Class_Terminate()
  If Not mobjProgress Is Nothing Then
    Set mobjProgress = Nothing
  End If
End Sub


Private Sub mobjProgress_Cancelled()
  RaiseEvent Cancelled
End Sub

Private Sub mobjProgress_Completed()
  RaiseEvent Completed
End Sub

Public Property Let HidePercentages(pblnHide As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.HidePercentages = pblnHide
  End If
End Property

Public Property Get Cancel() As Boolean
  If Not mobjProgress Is Nothing Then
    Cancel = mobjProgress.Cancel
  End If
End Property

Public Property Let Cancel(pblnCancel As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Cancel = pblnCancel
  End If
End Property

Public Property Get Caption() As String
  If Not mobjProgress Is Nothing Then
    Caption = mobjProgress.Caption
  End If
End Property

Public Property Let Caption(pstrCaption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Caption = pstrCaption
  End If
End Property

Public Property Get Time() As Boolean
  If Not mobjProgress Is Nothing Then
    Time = mobjProgress.Time
  End If
End Property

Public Property Let Time(pblnTime As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Time = pblnTime
  End If
End Property

Public Property Get NumberOfBars() As Integer
  If Not mobjProgress Is Nothing Then
    NumberOfBars = mobjProgress.NumberOfBars
  End If
End Property

Public Property Let NumberOfBars(pintNumberOfBars As Integer)
  If Not mobjProgress Is Nothing Then
    mobjProgress.NumberOfBars = pintNumberOfBars
  End If
End Property

Public Property Get Bar1MaxValue() As Long
  If Not mobjProgress Is Nothing Then
    Bar1MaxValue = mobjProgress.Bar1MaxValue
  End If
End Property

Public Property Let Bar1MaxValue(plngBar1MaxValue As Long)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar1MaxValue = plngBar1MaxValue
  End If
End Property

Public Property Get Bar2MaxValue() As Long
  If Not mobjProgress Is Nothing Then
    Bar2MaxValue = mobjProgress.Bar2MaxValue
  End If
End Property

Public Property Let Bar2MaxValue(plngBar2MaxValue As Long)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar2MaxValue = plngBar2MaxValue
  End If
End Property

Public Property Get Bar1Value() As Long
  If Not mobjProgress Is Nothing Then
    Bar1Value = mobjProgress.Bar1Value
  End If
End Property

Public Property Let Bar1Value(plngBar1Value As Long)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar1Value = plngBar1Value
  End If
End Property

Public Property Get Bar2Value() As Long
  If Not mobjProgress Is Nothing Then
    Bar2Value = mobjProgress.Bar2Value
  End If
End Property

Public Property Let Bar2Value(plngBar2Value As Long)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar2Value = plngBar2Value
  End If
End Property

Public Property Get Bar1Caption() As String
  If Not mobjProgress Is Nothing Then
    Bar1Caption = mobjProgress.Bar1Caption
  End If
End Property

Public Property Let Bar1Caption(pstrBar1Caption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar1Caption = pstrBar1Caption
  End If
End Property

Public Property Get Bar1RecordsCaption() As String
  If Not mobjProgress Is Nothing Then
    Bar1RecordsCaption = mobjProgress.Bar1RecordsCaption
  End If
End Property

Public Property Let Bar1RecordsCaption(pstrBar1RecordsCaption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar1RecordsCaption = pstrBar1RecordsCaption
  End If
End Property

Public Property Get Bar2Caption() As String
  If Not mobjProgress Is Nothing Then
    Bar2Caption = mobjProgress.Bar2Caption
  End If
End Property

Public Property Let Bar2Caption(pstrBar2Caption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar2Caption = pstrBar2Caption
  End If
End Property

Public Property Get Bar2RecordsCaption() As String
  If Not mobjProgress Is Nothing Then
    Bar2RecordsCaption = mobjProgress.Bar2RecordsCaption
  End If
End Property

Public Property Let Bar2RecordsCaption(pstrBar2RecordsCaption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Bar2RecordsCaption = pstrBar2RecordsCaption
  End If
End Property

Public Property Get Visible() As Boolean
  If Not mobjProgress Is Nothing Then
    Visible = mobjProgress.Visible
  End If
End Property

Public Property Let Visible(blnVisible As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.Visible = blnVisible
  End If
End Property

Public Property Get Cancelled() As Boolean
  If Not mobjProgress Is Nothing Then
    Cancelled = mobjProgress.Cancelled
  End If
End Property

Public Function CloseProgress() As Boolean
  If Not mobjProgress Is Nothing Then
    CloseProgress = mobjProgress.CloseProgress
  End If
End Function

Public Function OpenProgress() As Boolean
  If Not mobjProgress Is Nothing Then
    OpenProgress = mobjProgress.OpenProgress
  End If
End Function

Public Sub UpdateProgress(Optional pblnBatchMode As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.UpdateProgress pblnBatchMode
  End If
End Sub

Public Sub UpdateProgress2(Optional pblnBatchMode As Boolean)
  If Not mobjProgress Is Nothing Then
    mobjProgress.UpdateProgress2 pblnBatchMode
  End If
End Sub

Public Sub ResetBar1()
  If Not mobjProgress Is Nothing Then
    mobjProgress.ResetBar1
  End If
End Sub

Public Sub ResetBar2()
  If Not mobjProgress Is Nothing Then
    mobjProgress.ResetBar2
  End If
End Sub

Public Property Get StyleResource() As String
  If Not mobjProgress Is Nothing Then
    StyleResource = gobjProgress.StyleResource
  End If
End Property

Public Property Let StyleResource(ByVal sNewValue As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.StyleResource = sNewValue
  End If
End Property

Public Property Get StyleIni() As String
  If Not mobjProgress Is Nothing Then
    StyleIni = mobjProgress.StyleIni
  End If
End Property

Public Property Let StyleIni(ByVal sNewValue As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.StyleIni = sNewValue
  End If
End Property

Public Property Let MainCaption(strMainCaption As String)
  If Not mobjProgress Is Nothing Then
    mobjProgress.MainCaption = strMainCaption
  End If
End Property

Public Property Get MainCaption() As String
  If Not mobjProgress Is Nothing Then
    MainCaption = mobjProgress.MainCaption
  End If
End Property

Public Property Get AVI() As AVIType
  If Not mobjProgress Is Nothing Then
    AVI = mobjProgress.AVI
  End If
End Property

Public Property Let AVI(ByVal intValue As AVIType)
  If Not mobjProgress Is Nothing Then
    mobjProgress.AVI = intValue
  End If
End Property

Public Sub RefreshCaption(Index As Integer)
  If Not mobjProgress Is Nothing Then
    mobjProgress.RefreshCaption Index
  End If
End Sub
