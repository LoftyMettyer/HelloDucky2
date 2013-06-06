VERSION 5.00
Begin VB.UserControl COA_Line 
   CanGetFocus     =   0   'False
   ClientHeight    =   30
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ScaleHeight     =   30
   ScaleWidth      =   1005
   Begin VB.Line linGrey 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   1000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linWhite 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   1005
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "COA_Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Constant values.
Const giMinWidthV = 30
Const giMinHeightV = 100
Const giMinWidthH = 100
Const giMinHeightH = 30

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

' Properties.
Private miAlignment As Integer
Private glColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private miCurrentLength As Integer
Private mbResizing As Boolean
Private miTabIndex As Integer
Private msWFIdentifier As String
Private miWFItemType As Integer

Public Property Let WFIdentifier(New_Value As String)
  msWFIdentifier = New_Value
End Property

Public Property Get WFIdentifier() As String
  WFIdentifier = msWFIdentifier
End Property

Public Property Let WFItemType(New_Value As Integer)
  miWFItemType = New_Value
End Property

Public Property Get WFItemType() As Integer
  WFItemType = miWFItemType
End Property

Public Property Get Length() As Long
  Length = miCurrentLength
End Property

Public Property Let Length(ByVal lNewValue As Long)

  If miAlignment = 1 Then  ' horizontal
    UserControl.Width = lNewValue
  Else ' Vertical
    UserControl.Height = lNewValue
  End If
    
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get Selected() As Boolean
  Selected = gfSelected
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  gfSelected = pfNewValue
End Property

Public Property Get ControlLevel() As Integer
  ControlLevel = giControlLevel
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  giControlLevel = piNewValue
End Property

Public Property Get ColumnID() As Long
  ColumnID = glColumnID
End Property

Public Property Let ColumnID(ByVal pLngNewValue As Long)
  glColumnID = pLngNewValue
End Property

Public Property Get MinimumHeight() As Long
  If miAlignment = 1 Then
    MinimumHeight = giMinHeightH
  Else
    MinimumHeight = giMinHeightV
  End If
End Property

Public Property Get MinimumWidth() As Long
  If miAlignment = 1 Then
    MinimumWidth = giMinWidthH
  Else
    MinimumWidth = giMinWidthV
  End If
End Property

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Resize()

  If mbResizing Then Exit Sub
  
  mbResizing = True
  
  If miAlignment = 1 Then ' Horizontal
    
    UserControl.Height = 30
    
    If UserControl.Width < giMinWidthH Then UserControl.Width = giMinWidthH
    If UserControl.Height <> giMinHeightH Then UserControl.Height = giMinHeightH

    linGrey.X1 = 0
    linGrey.X2 = UserControl.Width
    linGrey.Y1 = 0
    linGrey.Y2 = 0
    
    linWhite.X1 = 0
    linWhite.X2 = UserControl.Width
    linWhite.Y1 = 15
    linWhite.Y2 = 15

    miCurrentLength = UserControl.Width

  Else ' Vertical
     
    UserControl.Width = 30
    If UserControl.Height < giMinHeightV Then UserControl.Height = giMinHeightV
    If UserControl.Width <> giMinWidthV Then UserControl.Width = giMinWidthV
  
    linGrey.X1 = 0
    linGrey.X2 = 0
    linGrey.Y1 = 0
    linGrey.Y2 = UserControl.Height
    
    linWhite.X1 = 15
    linWhite.X2 = 15
    linWhite.Y1 = 0
    linWhite.Y2 = UserControl.Height
  
    miCurrentLength = UserControl.Height
    
  End If
  
  mbResizing = False
  
End Sub

Private Sub UserControl_Initialize()
  miAlignment = 1
  miCurrentLength = 1000
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = linGrey.BorderColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  linGrey.BorderColor = vNewValue
End Property

Public Property Get TabIndex() As Integer
  TabIndex = miTabIndex
End Property

Public Property Let TabIndex(ByVal iNewValue As Integer)
  miTabIndex = iNewValue
End Property

Public Property Get Alignment() As Integer
  Alignment = miAlignment
End Property

Public Property Let Alignment(ByVal vNewValue As Integer)

  If vNewValue = 0 Or vNewValue = 1 Then
    miAlignment = vNewValue
    
    If miAlignment = 1 Then
    
      UserControl.Width = UserControl.Height
    Else
      UserControl.Height = UserControl.Width
    End If
  
  End If
  
End Property


'Private Sub UserControl_Resize()
'
'  If mbResizing Then Exit Sub
'
'  mbResizing = True
'
'  If miAlignment = 1 Then ' Horizontal
'
'    UserControl.Height = 30
'
'    If UserControl.Width < giMinWidthH Then UserControl.Width = giMinWidthH
'    If UserControl.Height <> giMinHeightH Then UserControl.Height = giMinHeightH
'
'    linGrey.X1 = 0
'    linGrey.X2 = UserControl.Width
'    linGrey.Y1 = 0
'    linGrey.Y2 = 0
'
'    linWhite.X1 = 0
'    linWhite.X2 = UserControl.Width
'    linWhite.Y1 = 15
'    linWhite.Y2 = 15
'
'    miCurrentLength = UserControl.Width
'
'  Else ' Vertical
'
'    UserControl.Width = 30
'    If UserControl.Height < giMinHeightV Then UserControl.Height = giMinHeightV
'    If UserControl.Width <> giMinWidthV Then UserControl.Width = giMinWidthV
'
'    linGrey.X1 = 0
'    linGrey.X2 = 0
'    linGrey.Y1 = 0
'    linGrey.Y2 = UserControl.Height
'
'    linWhite.X1 = 15
'    linWhite.X2 = 15
'    linWhite.Y1 = 0
'    linWhite.Y2 = UserControl.Height
'
'    miCurrentLength = UserControl.Height
'
'  End If
'
'  mbResizing = False
'
'End Sub

'Private Sub UserControl_Initialize()
'  miAlignment = 1
'  miCurrentLength = 1000
'End Sub


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

