VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "CODEJO~1.OCX"
Begin VB.UserControl COA_OptionGroup 
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   915
   ScaleWidth      =   1995
   Begin XtremeSuiteControls.GroupBox fraOptGroup 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1905
      _Version        =   851969
      _ExtentX        =   3360
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "Options : "
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton Option1 
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   1635
         _Version        =   851969
         _ExtentX        =   2884
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Option"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "COA_OptionGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

'Declare control events
' Declare public events.
Public Event Click()

'Declare windows API types
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

'Declare windows API function
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long

'Declare local variables
Private InResize As Boolean

Private mvarMaxLength As Integer

' For new property (Alignment)
Private miAlignment As Integer

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

Private Sub Option1_Click(Index As Integer)
  RaiseEvent Click
  
End Sub

Private Sub UserControl_InitProperties()
  On Error Resume Next
  Caption = Extender.Name
  Set Font = Ambient.Font
  ForeColor = Ambient.ForeColor
  BackColor = Ambient.BackColor

End Sub

Private Sub UserControl_Resize()

  Dim Index As Integer
  Dim intHeight As Long
  Dim intWidth As Long
  Dim temp As Integer
  
  Select Case miAlignment
  
    Case 0: 'Vertical
      If Not InResize Then
        InResize = True
        
        If fraOptGroup.BorderStyle = xtpFrameNone Then
          intHeight = UserControl.TextHeight(Caption) * 0.5
          If intHeight < 200 Then intHeight = 200
          intWidth = 0
        Else
          intHeight = UserControl.TextHeight(Caption) * 1.5
          If intHeight < 400 Then intHeight = 400
          intWidth = UserControl.TextWidth(Caption) + 100
        End If
        
        For Index = Option1.LBound To Option1.UBound
          With Option1(Index)
          
            .Width = 350 + UserControl.TextWidth(.Caption)
            .Height = IIf(UserControl.TextHeight(.Caption) < 240, 240, UserControl.TextHeight(.Caption))
              
            If Index = 0 Then
              If fraOptGroup.BorderStyle = xtpFrameNone Then
                .Top = 60
              Else
                .Top = UserControl.TextHeight(.Caption) + 100
              End If
            Else
              .Top = Option1(Index - 1).Top + UserControl.TextHeight(.Caption) + 50
            End If

            .Left = UserControl.TextWidth("W")

            intHeight = intHeight + .Height
            If .Width > intWidth Then intWidth = .Width
          
          End With
        Next Index
        
        intWidth = intWidth + (GetAvgCharWidth(UserControl.hDC) * 2) + 20
        
        With UserControl
          .Height = intHeight
          .Width = intWidth
        End With
        
        fraOptGroup.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        InResize = False
      End If
  
    Case 1: 'Horizonal
      If Not InResize Then
        InResize = True
        
        If fraOptGroup.BorderStyle = xtpFrameNone Then
          intHeight = Option1(Option1.UBound).Height + 90
          If intHeight < 240 Then intHeight = 240
          intWidth = 0
        Else
          intHeight = (Option1(Option1.UBound).Height + 90) * 2
          If intHeight < 400 Then intHeight = 400
          intWidth = UserControl.TextWidth(Caption) + 100
        End If
        
        For Index = Option1.LBound To Option1.UBound
          With Option1(Index)
          
            If fraOptGroup.BorderStyle = xtpFrameNone Then
              .Top = 60
            Else
              .Top = 300
            End If
            .Width = 285 + UserControl.TextWidth(.Caption) + GetAvgCharWidth(UserControl.hDC)
            .Height = IIf(UserControl.TextHeight(.Caption) < 240, 240, UserControl.TextHeight(.Caption))
            
            If Index > 0 Then
              .Left = Option1(Index - 1).Left + Option1(Index - 1).Width + UserControl.TextWidth("WW")
            End If
            
            intWidth = intWidth + .Width
          End With
        Next Index
        
'        intHeight = Option1(Option1.UBound).Height + 90
        
        If Not fraOptGroup.BorderStyle = xtpFrameNone Then
          intWidth = Maximum((Option1(Option1.UBound).Left) + Option1(Option1.UBound).Width + UserControl.TextWidth("WW"), UserControl.TextWidth(Caption) + 400)
        Else
          intWidth = (Option1(Option1.UBound).Left) + Option1(Option1.UBound).Width + UserControl.TextWidth("W")
        End If
        
        With UserControl
          .Height = intHeight
          .Width = intWidth
        End With
  
        fraOptGroup.Width = UserControl.Width
        fraOptGroup.Height = UserControl.Height
        
        InResize = False
      End If
  
  End Select

End Sub

' Returns the maximum of two values
Private Function Maximum(psngValue1 As Long, psngValue2 As Long) As Long
  Maximum = IIf(psngValue1 > psngValue2, psngValue1, psngValue2)
End Function

Public Property Get Alignment() As Integer
  Alignment = miAlignment

End Property

Public Property Let Alignment(ByVal vNewValue As Integer)

  If vNewValue = 0 Or vNewValue = 1 Then
    miAlignment = vNewValue
    UserControl_Resize
  End If
  
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = fraOptGroup.BackColor
  
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  Dim Index As Integer
  
  fraOptGroup.BackColor = NewColor
  
  For Index = Option1.LBound To Option1.UBound
    Option1(Index).BackColor = NewColor
  Next Index
  
End Property

Public Property Get BorderStyle() As Integer
  BorderStyle = IIf(fraOptGroup.BorderStyle = xtpFrameBorder, 1, 0)
End Property

Public Property Let BorderStyle(ByVal NewValue As Integer)

  fraOptGroup.BorderStyle = IIf(NewValue = 1, xtpFrameBorder, xtpFrameNone)
  UserControl_Resize
  
End Property

Public Property Get Caption() As String
  Caption = fraOptGroup.Caption
  
End Property

Public Property Let Caption(ByVal NewCaption As String)
  fraOptGroup.Caption = NewCaption
  
  UserControl_Resize

End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
  
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
  Dim ctlTemp As Control
  
  UserControl.Enabled() = NewEnabled
  If Not NewEnabled Then
    For Each ctlTemp In UserControl.Controls
      If TypeOf ctlTemp Is XtremeSuiteControls.RadioButton Or _
        TypeOf ctlTemp Is XtremeSuiteControls.GroupBox Then
          ctlTemp.Enabled = False
      End If
    Next
  End If
  
End Property

Public Property Get Font() As Font
  Set Font = fraOptGroup.Font
  
End Property

Public Property Set Font(ByVal NewFont As Font)
  Dim Index As Integer
  
  Set UserControl.Font = NewFont
  Set fraOptGroup.Font = UserControl.Font
  
  For Index = Option1.LBound To Option1.UBound
    Set Option1(Index).Font = UserControl.Font
  Next Index
  
  UserControl_Resize

End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = fraOptGroup.ForeColor
  
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
  Dim Index As Integer
  
  fraOptGroup.ForeColor = NewColor
  For Index = Option1.LBound To Option1.UBound
    Option1(Index).ForeColor = NewColor
  Next Index
  
End Property

Public Property Get hWnd() As Long
  hWnd = fraOptGroup.hWnd
  
End Property

Public Property Get MaxLength() As Integer
  MaxLength = mvarMaxLength
  
End Property

Public Property Let MaxLength(intMaxLen As Integer)
  mvarMaxLength = IIf(intMaxLen > 0, intMaxLen, 0)
  
End Property
Public Function SetOptions(ByRef pasOptions As Variant)
  Dim iX As Integer
  Dim iIndex As Integer
  Dim iArrayDim As Integer
  Dim iTemp As Integer
  
  iIndex = 0
  ' Decide if we have a one or 2 dimension array
  iArrayDim = 2
  On Error GoTo err_SetOptions
  
  iTemp = UBound(pasOptions, 2) > 0
  
  If iArrayDim = 2 Then
    For iX = LBound(pasOptions, 2) To UBound(pasOptions, 2)
      If Option1.Count - 1 < iX Then
        Load Option1(iX)
      End If
      
      With Option1(iIndex)
        .Caption = Replace(pasOptions(0, iX), "&", "&&")
        .Visible = True
      End With
      
      iIndex = iIndex + 1
    Next iX
  Else
    For iX = LBound(pasOptions, 1) To UBound(pasOptions, 1)
      If Option1.Count - 1 < iX Then
        Load Option1(iX)
      End If
      
      With Option1(iIndex)
        .Caption = Replace(pasOptions(iX), "&", "&&")
        .Visible = True
      End With
      
      iIndex = iIndex + 1
    Next iX
  End If
  
  UserControl_Resize
  
Exit Function

err_SetOptions:
  If Err.Number = 9 Then
    iArrayDim = 1
    Resume Next
  Else
    Err.Raise Err.Number, "SetOptions", Err.Description
  End If
  
End Function

Public Property Get Text() As String
  Dim Index As Integer
  
  Index = GetSelectedOption
  
  If Index >= 0 Then
    If MaxLength > 0 Then
      Text = Left(Replace(Option1(Index).Caption, "&&", "&"), MaxLength)
    Else
      Text = Replace(Option1(Index).Caption, "&&", "&")
    End If
  Else
    Text = vbNullString
  End If
  
End Property

Public Property Let Text(ByVal NewValue As String)
  Dim Index As Integer

  NewValue = UCase(Trim(NewValue))
  
  For Index = Option1.LBound To Option1.UBound
    'JPD 20050810 Fault 10178
    'If UCase(Left(Option1(Index).Caption, Len(NewValue))) = NewValue Then
    If Replace(UCase(Trim(Option1(Index).Caption)), "&&", "&") = Trim(NewValue) Then
      Option1(Index).Value = True
      Exit For
    Else
      Option1(Index).Value = False
    End If
  Next Index
  
End Property

Public Property Get Value() As Integer
  Value = GetSelectedOption
  
End Property

Public Property Let Value(ByVal NewValue As Integer)
  Dim Index As Integer

  For Index = Option1.LBound To Option1.UBound
    Option1(Index).Value = (Index = NewValue)
  Next Index
  
End Property

Public Sub Refresh()
  UserControl_Resize
  UserControl.Refresh
  
End Sub

Private Function GetSelectedOption() As Integer
  Dim Index As Integer
  Dim Selected As Integer
  
  Selected = -1
  For Index = Option1.LBound To Option1.UBound
    If Option1(Index).Value = True Then
      Selected = Index
      Exit For
    End If
  Next Index

  GetSelectedOption = Selected
  
End Function

Private Function GetAvgCharWidth(ByVal hDC As Long) As Integer
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(hDC, typTxtMetrics)
  GetAvgCharWidth = (typTxtMetrics.tmAveCharWidth * Screen.TwipsPerPixelX)
  
End Function

Public Sub About()
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
    
End Sub





