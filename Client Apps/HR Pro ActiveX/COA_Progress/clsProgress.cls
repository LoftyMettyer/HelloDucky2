VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COA_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

Public Event Cancelled()
Public Event Completed()


Const GAP_AFTER_TIMECANCEL = 135
Const GAP_BEFORE_BUTTONS = 135
Const FRAME_HEIGHT = 1010
Const BUTTON_HEIGHT = 400
Const LABEL_OFFSET = 30
Const LABEL_HEIGHT = 195

Private mintAvi As AVIType
Private mstrStyleResource As String
Private mstrStyleIni As String
Private mblnCancel As Boolean
Private mstrCaption As String
Private mblnTime As Boolean
Private mintNumberOfBars As Integer
Private mlngBar1Value As Long
Private mlngBar1MaxValue As Long
Private mstrBar1Caption As String
Private mlngBar2Value As Long
Private mlngBar2MaxValue As Long
Private mstrBar2Caption As String
Private mstrBar1RecordsCaption As String
Private mstrBar2RecordsCaption As String
Private mLngTimeStarted As Long

Private mblnHidePercentages As Boolean

Private WithEvents fProgress As frmProgress
Attribute fProgress.VB_VarHelpID = -1

Private mblnCancelled As Boolean
Private mintAnswer As Integer
Private mblnAnsered As Boolean

Private mblnVisible As Boolean


Public Enum AVIType
  [dbNone] = 0
  [dbDiary] = 101
  [dbExcel] = 102
  [dbInternet] = 103
  [dbScreenLoad] = 104
  [dbText] = 105
  [dbTransfer] = 106
  [dbVaporise] = 107
  [dbWord] = 108
  [dbAccord] = 109
  [dbSave] = 110
  [dbTable] = 111
  [dbCopyTable] = 112
  [dbPicture] = 113
  [dbScreenAutoLayout] = 114
  [dbWorkflow] = 115
  [dbLoadDomain] = 116
  [dbLoadUsers] = 117
  [dbXferOwnership] = 118
  [dbLogin] = 119
  [dbPicklist] = 120
  [dbSaveRec] = 121
  [dbAutoAdd] = 122
  [dbCrossTab] = 123
  [dbOLESearch] = 124
  [dbEMail] = 125
End Enum

Private Function GetTickCount2() As Long
  On Local Error Resume Next
  GetTickCount2 = Timer
End Function


Private Sub Class_Initialize()
  App.TaskVisible = False
End Sub

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
  Dim Rc          As Long
  Dim rClsId      As udtGUID
  Dim IID         As String
  Dim bIID()      As Byte

  If (riid <> 0) Then
      CopyMemory rClsId, ByVal riid, Len(rClsId)

      bIID = String$(MAX_GUIDLEN, 0)
      Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
      Rc = InStr(1, bIID, vbNullChar) - 1
      IID = Left$(UCase(bIID), Rc)

      Select Case IID
          Case IID_IDispatch
              If ((dwEnabledOptions And dwOptionsSetMask) <> _
           INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                  Err.Raise E_FAIL
                  Exit Sub
              Else
                  If Not m_fSafeForScripting Then
                      Err.Raise E_FAIL
                  End If
                  Exit Sub
              End If

          Case IID_IPersistStorage, IID_IPersistStream, _
        IID_IPersistPropertyBag
              If ((dwEnabledOptions And dwOptionsSetMask) <> _
            INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                  Err.Raise E_FAIL
                  Exit Sub
              Else
                  If Not m_fSafeForInitializing Then
                      Err.Raise E_FAIL
                  End If
                  Exit Sub
              End If

          Case Else
              Err.Raise E_NOINTERFACE
              Exit Sub
      End Select
  End If
    
End Sub

Public Property Let HidePercentages(pblnHide As Boolean)
  
  On Local Error GoTo LocalErr
  
  mblnHidePercentages = pblnHide
  
Exit Property

LocalErr:
  
  mblnHidePercentages = False

End Property

'Public Property Get AviFile() As String
'  AviFile = mstrAviFile
'End Property

'Public Property Let AviFile(pstrAviFile As String)
'
'  On Error GoTo LocalErr
'
'  mstrAviFile = pstrAviFile
'  RefreshProgressBar
'
'  With frmProgress.Animation1
'    .Open mstrAviFile
'    .Visible = True
'  End With
'
'Exit Property
'
'LocalErr:
'  frmProgress.Animation1.Visible = False
'
'End Property

Public Property Get Cancel() As Boolean
  Cancel = mblnCancel
End Property

Public Property Let Cancel(pblnCancel As Boolean)
  mblnCancel = pblnCancel
End Property

Public Property Get Caption() As String
  Caption = mstrCaption
End Property

Public Property Let Caption(pstrCaption As String)
  
  RefreshProgressBar

  mstrCaption = pstrCaption
  fProgress.SetCaption mstrCaption

End Property

Public Property Get Time() As Boolean
  Time = mblnTime
End Property

Public Property Let Time(pblnTime As Boolean)
  mblnTime = pblnTime
End Property

Public Property Get NumberOfBars() As Integer
  NumberOfBars = mintNumberOfBars
End Property

Public Property Let NumberOfBars(pintNumberOfBars As Integer)
  mintNumberOfBars = pintNumberOfBars
End Property

Public Property Get Bar1MaxValue() As Long
  Bar1MaxValue = mlngBar1MaxValue
End Property

Public Property Let Bar1MaxValue(plngBar1MaxValue As Long)
  
  If plngBar1MaxValue > 0 Then
    mlngBar1MaxValue = plngBar1MaxValue
  End If
  
End Property

Public Property Get Bar2MaxValue() As Long
  Bar2MaxValue = mlngBar2MaxValue
End Property

Public Property Let Bar2MaxValue(plngBar2MaxValue As Long)
  
  If plngBar2MaxValue > 0 Then
    mlngBar2MaxValue = plngBar2MaxValue
  End If
  
End Property

Public Property Get Bar1Value() As Long
  Bar1Value = mlngBar1Value
End Property

Public Property Let Bar1Value(plngBar1Value As Long)
  mlngBar1Value = plngBar1Value
End Property

Public Property Get Bar2Value() As Long
  Bar2Value = mlngBar2Value
End Property

Public Property Let Bar2Value(plngBar2Value As Long)
  mlngBar2Value = plngBar2Value
End Property

Public Property Get Bar1Caption() As String
  Bar1Caption = mstrBar1Caption
End Property

Public Property Let Bar1Caption(pstrBar1Caption As String)

  RefreshProgressBar
  mstrBar1Caption = pstrBar1Caption
  RefreshCaption 1

End Property

Public Property Get Bar1RecordsCaption() As String
  
  Bar1RecordsCaption = mstrBar1RecordsCaption

End Property

Public Property Let Bar1RecordsCaption(pstrBar1RecordsCaption As String)
  
  RefreshProgressBar
  mstrBar1RecordsCaption = pstrBar1RecordsCaption
  RefreshCaption 1

End Property

Public Property Get Bar2Caption() As String
  Bar2Caption = mstrBar2Caption
End Property

Public Property Let Bar2Caption(pstrBar2Caption As String)
  
  RefreshProgressBar
  mstrBar2Caption = pstrBar2Caption
  RefreshCaption 2

End Property

Public Property Get Bar2RecordsCaption() As String
  
  Bar2RecordsCaption = mstrBar2RecordsCaption

End Property

Public Property Let Bar2RecordsCaption(pstrBar2RecordsCaption As String)
  
  RefreshProgressBar

  mstrBar2RecordsCaption = pstrBar2RecordsCaption
  RefreshCaption 2

End Property

Public Property Get Visible() As Boolean
  Visible = mblnVisible
End Property

Public Property Let Visible(blnVisible As Boolean)
  mblnVisible = blnVisible
  RefreshProgressBar
End Property


Private Function RefreshProgressBar() As Boolean
  
  Dim intError As Integer


  On Local Error Resume Next
  
  If fProgress Is Nothing Then
    Set fProgress = New frmProgress
    fProgress.StyleResource = mstrStyleResource
    fProgress.StyleIni = mstrStyleIni
  End If
  
    
  If Not fProgress.Visible And mblnVisible Then
    fProgress.Visible = mblnVisible
    DoEvents
    
    intError = 0
    If Not fProgress.FormEnableCloseButton(mblnVisible) Then
      intError = intError + 1
    End If
    
    If Not fProgress.FormSetTopMost Then
      intError = intError + 2
    End If

  Else
    fProgress.Visible = mblnVisible
    DoEvents

  End If
  
  
  Screen.MousePointer = IIf(mblnCancel Or Not mblnVisible, vbNormal, vbHourglass)

End Function

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Function CloseProgress() As Boolean
  
  Visible = False
  
  ' reset cancelled status to default (false)
  mblnCancelled = False
  
  ' reset other properties to default
  'mstrAviFile = ""
  mblnCancel = True
  mblnTime = False
  'mstrCaption = "Progress Bar"
  mintNumberOfBars = 1
  Bar1MaxValue = 100
  Bar1Value = 0
  Bar2MaxValue = 100
  Bar2Value = 0
  Bar1RecordsCaption = ""
  Bar2RecordsCaption = ""
  Bar1Caption = ""
  Bar2Caption = ""
  Caption = ""
  
  ' Bye bye form
  If Not fProgress Is Nothing Then
    Unload fProgress
    Set fProgress = Nothing
  End If
  
  CloseProgress = True
  
End Function


Public Function OpenProgress() As Boolean
  
  Dim lngTop As Long

  
  On Error GoTo LocalErr
  
  ' Instantiate the form if it does not already exist.
  RefreshProgressBar
  
  mLngTimeStarted = GetTickCount2
  
  With fProgress
  
    ' Show or hide the Animation
    .SetAVI CInt(mintAvi)
    
    
    ' Which progress bars to show
    lngTop = 1110
    .fraProgress2.Visible = False
    .fraProgress1.Visible = False
  
    If mintNumberOfBars = 1 Or mintNumberOfBars = 2 Then
      
      If mintNumberOfBars = 2 Then
        .fraProgress2.Visible = True
        .fraProgress2.Top = lngTop
        lngTop = lngTop + FRAME_HEIGHT
      End If
      
      .fraProgress1.Visible = True
      .fraProgress1.Top = lngTop
      lngTop = lngTop + FRAME_HEIGHT
      
    End If
    
    lngTop = lngTop + GAP_BEFORE_BUTTONS
    
    
    
    ' Show time label / cancel button
    .lblTime.Visible = mblnTime
    .cmdCancel.Visible = mblnCancel
    ' NPG20091013 Fault HRPro-312
    .FormEnableCloseButton (False)
    
    ' If the cancel button is invisible, then change the pointer to an hourglass
    If mblnCancel Then Screen.MousePointer = vbNormal Else Screen.MousePointer = vbHourglass
    
    If mblnTime Then
      .lblTime.Visible = True
      .lblTime.Top = lngTop + LABEL_OFFSET
    Else
      .lblTime.Visible = False
    End If
    
    If mblnCancel Then
      .cmdCancel.Visible = True
      .cmdCancel.Top = lngTop
      lngTop = lngTop + BUTTON_HEIGHT + GAP_AFTER_TIMECANCEL
    Else
      .cmdCancel.Value = False
      lngTop = lngTop + (IIf(mblnTime, LABEL_HEIGHT, 0)) + GAP_AFTER_TIMECANCEL
    End If
    
    .WindowState = vbNormal
    
    'MH20021018 Think that you need to allow room for the form border, like this..
    .Height = lngTop + (.Height - .ScaleHeight)
    
    ''' RH: for some reason, lngTop cannot be used to set the height of the form
    ''' so for now, use this sub...
    ''SetFormHeight
    
'    ' Set the other properties
'
'    If Len(mstrBar1RecordsCaption) = 0 Then
'      ' Bar 1 records caption not required so adjust as necessary
'      .picProgress1.Top = .picProgress1.Top - 100
'      .lblProgress1Records.Visible = False
'    End If
'
'    If Len(mstrBar2RecordsCaption) = 0 Then
'      ' Bar 2 records caption not required so adjust as necessary
'      .picProgress2.Top = .picProgress2.Top - 100
'      .lblProgress2Records.Visible = False
'    End If
    
    .SetCaption mstrCaption
    
    .pbrProgress1.Max = 100
    .pbrProgress1.Value = 0
    .pbrProgress2.Max = 100
    .pbrProgress2.Value = 0
    
    RefreshCaption 0
    
    .lblTime.Caption = vbNullString
    
    ' Position the form in the middle of the screen.
    .Top = Int((Screen.Height - .Height) / 2)
    .Left = Int((Screen.Width - .Width) / 2)
    
    Visible = True
    .Show
    
  End With
  
  Visible = True
  
  OpenProgress = True

Exit Function
  
LocalErr:

  If Err.Number = 53 Then
    fProgress.Animation1.Visible = False
    Resume Next
  End If
  
  MsgBox "Warning : An error has occurred." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "HRProProgress"
  OpenProgress = False
  
End Function

Public Sub UpdateProgress(Optional pblnBatchMode As Boolean)
  
  Dim plngNow As Long
  
  Dim plngRemains As Long
  Dim pintHours As Integer, pintMins As Integer, pintSecs As Integer
  Dim pstrTime As String

  RefreshProgressBar

  ' Quick Bodge Test
  If pblnBatchMode Then
    UpdateProgress2
    Exit Sub
  End If
  
  ' Perform another progress step for the first bar.
  
  If mLngTimeStarted = 0 Then
    mLngTimeStarted = GetTickCount2
  End If
  
  ' Increment the number of progress steps made.
  mlngBar1Value = mlngBar1Value + 1
    
  ' Set percentage label RH 09/02/01
  If (Not mblnHidePercentages) And (NumberOfBars > 0) Then fProgress.lblBar1Percent.Caption = Int((mlngBar1Value / mlngBar1MaxValue) * 100) & "%"
  
  ' Check to see if the progress is now complete.
  If mlngBar1Value >= mlngBar1MaxValue Then
    
    Complete
  
  Else
    
    plngNow = GetTickCount2
    
    ' Update the progress bar display.
    fProgress.pbrProgress1.Value = Int((mlngBar1Value / mlngBar1MaxValue) * 100)
  
    ' Update any other properties (such as captions etc)
    RefreshCaption 0
    
    ' Update the 'time remaining' display, if visible
    If mblnTime = True Then
    
      plngRemains = ((plngNow - mLngTimeStarted) / mlngBar1Value) * (mlngBar1MaxValue - mlngBar1Value)
      plngRemains = (plngRemains + 500) \ 1000
      pintHours = plngRemains \ (60 ^ 2)
      pintMins = (plngRemains \ 60) - (pintHours * 60)
      pintSecs = plngRemains Mod 60
      
      pstrTime = "Time remaining :"
      If pintHours > 0 Then
        pstrTime = Space(2) & pstrTime & pintHours & _
          " Hour" & IIf(pintHours <> 1, "s", vbNullString)
      End If
      If pintHours > 0 Or pintMins > 0 Then
        pstrTime = pstrTime & Space(2) & pintMins & _
          " Minute" & IIf(pintMins <> 1, "s", vbNullString)
      End If
      pstrTime = Trim(pstrTime) & Space(2) & pintSecs & _
        " Second" & IIf(pintSecs <> 1, "s", vbNullString)
      
      fProgress.lblTime.Caption = pstrTime
    
    End If
    
  End If
  
  ' Allow the calling program to continue execution.
  DoEvents
  
End Sub

Public Sub UpdateProgress2(Optional pblnBatchMode As Boolean)

  'Only refresh the progress bar once per second
  
  
  ' Quick Bodge Test
  If pblnBatchMode Then
    UpdateProgress
    Exit Sub
  End If
  
  ' Perform another progress step for the second bar.
    
  If mintNumberOfBars < 2 Then Exit Sub
  
  If mLngTimeStarted = 0 Then
    mLngTimeStarted = GetTickCount2
  End If
  
  ' Increment the number of progress steps made.
  mlngBar2Value = mlngBar2Value + 1
  
  ' Update any other properties (such as captions etc)
  RefreshCaption 2

    
  ' Update the progress bar display.
  fProgress.pbrProgress2.Value = Int((mlngBar2Value / mlngBar2MaxValue) * 100)
    
  ' Set percentage label RH 09/02/01
  If (Not mblnHidePercentages) And (NumberOfBars > 0) Then fProgress.lblBar2Percent.Caption = Int((mlngBar2Value / mlngBar2MaxValue) * 100) & "%"
    
  ' Allow the calling program to continue execution.
  DoEvents
  
End Sub

Public Sub ResetBar2()

  RefreshProgressBar

  If mintNumberOfBars < 2 Then Exit Sub

  mlngBar2Value = 0
  fProgress.pbrProgress2.Value = mlngBar2Value
'  mstrBar2Caption = ""
  Bar2Caption = ""
  Bar2RecordsCaption = ""

  DoEvents
  
End Sub

Private Sub Complete()
  
  ' Update the progress form display if the progress is complete.

  If Not fProgress Is Nothing Then
    With fProgress
      '.pbrProgress1.Value = 100
      .pbrProgress1.Value = .pbrProgress1.Max
      If mblnTime = True Then
        .lblTime.Caption = "Completed."
        .lblTime.Refresh
      End If
      If mblnCancel = True Then .cmdCancel.Enabled = False
    End With
    
    ' Allow the calling program to continue execution.
    DoEvents
  End If
  
  ' Raise an event to notify the calling program that progress has completed.
  RaiseEvent Completed
  
End Sub

Private Sub Class_Terminate()

  If Not fProgress Is Nothing Then
    Unload fProgress
    Set fProgress = Nothing
  End If
  
End Sub

Private Sub fProgress_Cancelled()

  RefreshProgressBar

  fProgress.SetCaption "Cancelled - Please Wait..."
  fProgress.cmdCancel.Enabled = False
  DoEvents
  mblnCancelled = True
  RaiseEvent Cancelled
End Sub

Public Property Get StyleResource() As String
  StyleResource = mstrStyleResource
End Property

Public Property Let StyleResource(ByVal sNewValue As String)
  mstrStyleResource = sNewValue
End Property

Public Property Get StyleIni() As String
  StyleIni = mstrStyleIni
End Property

Public Property Let StyleIni(ByVal sNewValue As String)
  mstrStyleIni = sNewValue
End Property

Public Property Let MainCaption(strMainCaption As String)
  fProgress.lblMainCaption.Caption = Replace(strMainCaption, "&", "&&")
End Property

Public Property Get MainCaption() As String
  MainCaption = Replace(fProgress.lblMainCaption.Caption, "&&", "&")
End Property

Public Property Get AVI() As AVIType
  AVI = mintAvi
End Property

Public Property Let AVI(ByVal intValue As AVIType)
  mintAvi = intValue
  RefreshProgressBar
  fProgress.SetAVI CInt(mintAvi)
End Property


Public Sub RefreshCaption(Index As Integer)

  'Index
  '0 = both
  '1 = only 1
  '2 = only 2

  If Index <> 2 Then
    fProgress.lblProgress1.Caption = _
        Replace(mstrBar1Caption, "&", "&&") & vbCrLf & _
        Replace(mstrBar1RecordsCaption, "&", "&&")
  End If
  
  If Index <> 1 Then
    fProgress.lblProgress2.Caption = _
        Replace(mstrBar2Caption, "&", "&&") & vbCrLf & _
        Replace(mstrBar2RecordsCaption, "&", "&&")
  End If

End Sub

