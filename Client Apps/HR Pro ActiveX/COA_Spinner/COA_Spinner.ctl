VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl COA_Spinner 
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   LockControls    =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   1725
   ToolboxBitmap   =   "COA_Spinner.ctx":0000
   Begin VB.TextBox txtASRSpinner 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Text            =   "0"
      Top             =   200
      Width           =   1000
   End
   Begin ComCtl2.UpDown updnASRSpinner 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   200
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      OrigLeft        =   1260
      OrigTop         =   210
      OrigRight       =   1455
      OrigBottom      =   600
      Max             =   99999
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "COA_Spinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Public events
Public Event Change()
Public Event Click()

' Properties.
Private giSpinnerPosition As Integer

' Globals.
Private gsOldText As String
Private giOldStart As Integer
Private giOldSelLength As Integer
Private gfIgnoringInput As Boolean

' Constants.
Const gLngMinWidth = 600
Const gLngMinHeight = 300

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



Private Sub txtASRSpinner_Change()
  Dim fRestoreOldText As Boolean
  Dim lRealMax As Long
  Dim lRealMin As Long
  
  ' If this method is called as we are restoring an original string
  ' after having invalid data input, then do nothing.
  If gfIgnoringInput Then
    Exit Sub
  End If
  
  ' Determine the real max and min. The properties can be set so that the min is
  ' greater than then max. This results in the up key decreasing the value; the down
  ' key increases it. This is consistent with the normal upDown control operation.
  With updnASRSpinner
    lRealMax = IIf(.Max > .Min, .Max, .Min)
    lRealMin = IIf(.Max > .Min, .Min, .Max)
  End With
  
  ' Initialise local variables.
  fRestoreOldText = False
  
  With txtASRSpinner
  
    ' Restore the original text if the new value does not evaluate
    ' to a numeric.
    If Not IsNumeric(.Text) And _
      .Text <> "" And _
      .Text <> "-" Then
      
      fRestoreOldText = True
    End If
    
    ' Note. A string with a single digit followed by a plus or minus sign
    ' evaluates to a numeric. we want to invalidate these.
    If InStr(.Text, "-") > 1 Or _
      InStr(.Text, "+") > 1 Then
      
      fRestoreOldText = True
    End If
    
    ' Ignore any decimal points or spaces.
    If InStr(.Text, ".") Or _
      InStr(.Text, " ") Then
    
      fRestoreOldText = True
    End If
    
    ' Bound the value to the defined maximum and
    ' minimum values for the spinner.
    If Val(.Text) > lRealMax Then
      gfIgnoringInput = True
      .Text = Trim(Str(lRealMax))
      .SelStart = Len(.Text)
      gfIgnoringInput = False
    End If
    
    If Val(.Text) < lRealMin Then
      gfIgnoringInput = True
      .Text = Trim(Str(lRealMin))
      .SelStart = Len(.Text)
      gfIgnoringInput = False
    End If
    
    ' Restore the original text if we want to ignore the key pressed.
    If fRestoreOldText Then
      gfIgnoringInput = True
      .Text = gsOldText
      .SelStart = giOldStart
      .SelLength = giOldSelLength
      gfIgnoringInput = False
    Else
      ' Call the user-defined change event.
      RaiseEvent Change
    End If
    
    ' Remeber the text, and the selected part in case we need to restore it.
    gsOldText = .Text
    giOldStart = .SelStart
    giOldSelLength = .SelLength
  End With
  
End Sub

Private Sub txtASRSpinner_Click()

  ' Call the user-defined click event.
  RaiseEvent Click

End Sub


Private Sub txtASRSpinner_GotFocus()
    
  gfIgnoringInput = False
  
  With txtASRSpinner
  
    ' Select the whole text.
    .SelStart = 0
    .SelLength = Len(.Text)
  
    ' Remember the original text value and what is selected.
    gsOldText = .Text
    giOldStart = .SelStart
    giOldSelLength = .SelLength

  End With
  
End Sub

Private Sub txtASRSpinner_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim iSign As Integer
  
  ' Determine if the upDown controls are inverted.
  iSign = IIf(updnASRSpinner.Min > updnASRSpinner.Max, -1, 1)
  
  ' If the user presses the up or down keys then simulate
  ' the associated upDown controls events.
  If KeyCode = vbKeyUp Then
    With txtASRSpinner
    
      ' If the shift key is pressed then go straight to the max.
      If (Shift And vbShiftMask) > 0 Then
        .Text = Trim(Str(updnASRSpinner.Max))
      Else
        .Text = Trim(Str(Val(.Text) + (iSign * updnASRSpinner.Increment)))
      End If
    
    End With
  ElseIf KeyCode = vbKeyDown Then
    With txtASRSpinner
    
      ' If the shift key is pressed then go straight to the min.
      If (Shift And vbShiftMask) > 0 Then
        .Text = Trim(Str(updnASRSpinner.Min))
      Else
        .Text = Trim(Str(Val(.Text) - (iSign * updnASRSpinner.Increment)))
      End If
    
    End With
  End If

End Sub

Private Sub txtASRSpinner_KeyPress(KeyAscii As Integer)
  ' Remember what text is selected.
  giOldStart = txtASRSpinner.SelStart
  giOldSelLength = txtASRSpinner.SelLength

End Sub

Private Sub txtASRSpinner_LostFocus()
  ' Update the text to be the string of the value of the text.
  ' This removes any leading zeroes, etc.
  If Not txtASRSpinner.Text = "" Then
    txtASRSpinner.Text = Trim(Str(Val(txtASRSpinner.Text)))
  End If

End Sub

Private Sub updnASRSpinner_DownClick()
  Dim iSign As Integer

  ' Determine if the upDown controls are inverted.
  iSign = IIf(updnASRSpinner.Min > updnASRSpinner.Max, -1, 1)
  
  ' Increment/decrement the spinner text value as required.
  With txtASRSpinner
    .Text = Trim(Str(Val(.Text) - (iSign * updnASRSpinner.Increment)))

    ' Select the whole text.
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

  ' Call the user-defined change event.
  RaiseEvent Click

End Sub


Private Sub updnASRSpinner_UpClick()
  Dim iSign As Integer

  ' Determine if the upDown controls are inverted.
  iSign = IIf(updnASRSpinner.Min > updnASRSpinner.Max, -1, 1)

  With txtASRSpinner
    ' Increment/decrement the text value as required.
    .Text = Trim(Str(Val(.Text) + (iSign * updnASRSpinner.Increment)))

    ' Select the whole text.
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
  ' Call the user-defined change event.
  RaiseEvent Click

End Sub


Private Sub UserControl_InitProperties()
  On Error Resume Next
  
  ' Initialise the properties.
  Set Font = Ambient.Font
  BackColor = vbWhite
  ForeColor = Ambient.ForeColor
  Alignment = vbRightJustify
  SpinnerPosition = vbRightJustify
  
End Sub
Public Property Get BackColor() As OLE_COLOR
  ' Return the backColor property of the text box.
  BackColor = txtASRSpinner.BackColor
  
End Property
Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the 'About' box.
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
    
End Sub

Public Property Let BackColor(ByVal pcolNewColor As OLE_COLOR)
  ' Set the BackColor property of the textbox.
  txtASRSpinner.BackColor = pcolNewColor
  PropertyChanged "BackColor"
  
End Property







Public Property Get Enabled() As Boolean
  ' Return the userControl's enabled property.
  Enabled = UserControl.Enabled
  
End Property

Public Property Let Enabled(ByVal pfNewValue As Boolean)
  ' Set the user control's enabled property.
  txtASRSpinner.Enabled = pfNewValue
  updnASRSpinner.Enabled = pfNewValue
  UserControl.Enabled = pfNewValue
  
  PropertyChanged "Enabled"

End Property

Public Property Get Font() As Font
 ' Return the text box font.
  Set Font = txtASRSpinner.Font

End Property

Public Property Set Font(ByVal pfNewFont As Font)
  ' Set the textbox font.
  Set txtASRSpinner.Font = pfNewFont
  
  PropertyChanged "Font"
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ' Return the foreColor property of the text box.
  ForeColor = txtASRSpinner.ForeColor
  
End Property

Public Property Let ForeColor(ByVal pcolNewColor As OLE_COLOR)
  ' Set the ForeColor property of the textbox.
  txtASRSpinner.ForeColor = pcolNewColor
  PropertyChanged "ForeColor"
  
End Property

Public Property Get MinimumValue() As Long
  ' Return the upDown control's minimum property.
  MinimumValue = updnASRSpinner.Min

End Property

Public Property Let MinimumValue(ByVal plNewValue As Long)
  ' Set the upDown control's min property.
  updnASRSpinner.Min = plNewValue
  PropertyChanged "MinimumValue"

End Property

Public Property Get MaximumValue() As Long
  ' Return the upDown control's maximum property.
  MaximumValue = updnASRSpinner.Max

End Property

Public Property Let MaximumValue(ByVal plNewValue As Long)
  ' Set the upDown control's max property.
  updnASRSpinner.Max = plNewValue
  PropertyChanged "MaximumValue"

End Property

Public Property Get Increment() As Integer
  ' Return the upDown control's increment property.
  Increment = updnASRSpinner.Increment

End Property

Public Property Let Increment(ByVal piNewValue As Integer)
  ' Set the upDown control's increment property.
  updnASRSpinner.Increment = IIf(piNewValue > 0, piNewValue, 1)
  PropertyChanged "Increment"

End Property

Public Property Get SpinnerPosition() As Integer
  ' Return the controls Spinner Position property.
  SpinnerPosition = giSpinnerPosition
  
End Property

Public Property Let SpinnerPosition(ByVal piNewValue As Integer)
  ' Set the control's Spinner Position property.
  If (piNewValue = vbLeftJustify) Or _
    (piNewValue = vbRightJustify) Then
    
    giSpinnerPosition = piNewValue
    PropertyChanged "SpinnerPosition"
    UserControl_Resize
  Else
    MsgBox "Invalid value. " & vbCr & vbCr & _
      "0 = spinner to the left of the textbox." & vbCr & _
      "1 = spinner to the right of the textbox.", vbOKOnly, App.ProductName
  End If

End Property

Public Property Get Alignment() As Integer
  Alignment = txtASRSpinner.Alignment

End Property


Public Property Let Alignment(ByVal piNewValue As Integer)
  ' Set the control's Spinner Position property.
  If (piNewValue = vbLeftJustify) Or _
    (piNewValue = vbRightJustify) Or _
    (piNewValue = vbCenter) Then
    
    txtASRSpinner.Alignment = piNewValue
    
    PropertyChanged "Alignment"
  Else
    MsgBox "Invalid value. " & vbCr & vbCr & _
      "0 = text left aligned." & vbCr & _
      "1 = text right aligned." & vbCr & _
      "2 = text centred.", vbOKOnly, App.ProductName
  End If

End Property



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Read the previous set of properties.
  BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  Enabled = PropBag.ReadProperty("Enabled", True)
  Increment = PropBag.ReadProperty("Increment", 1)
  MaximumValue = PropBag.ReadProperty("MaximumValue", 10)
  MinimumValue = PropBag.ReadProperty("MinimumValue", 0)
  SpinnerPosition = PropBag.ReadProperty("SpinnerPosition", vbRightJustify)
  Alignment = PropBag.ReadProperty("Alignment", vbRightJustify)
  Text = PropBag.ReadProperty("Text", vbNullChar)

End Sub

Private Sub UserControl_Resize()
  Dim lngCtrlHeight As Long
  Dim lngCtrlWidth As Long
  Dim lngUpDnWidth As Long
  
  ' Do not let the user make the control too small.
  With UserControl
    .Width = IIf(.Width < gLngMinWidth, gLngMinWidth, .Width)
    .Height = IIf(.Height < gLngMinHeight, gLngMinHeight, .Height)
    
    lngCtrlWidth = .Width
    lngCtrlHeight = .Height
  End With
  
  lngUpDnWidth = updnASRSpinner.Width
  
  ' Resize the text and upDown controls as our custom
  ' ASRSpinner control is resized. NB. the upDown control
  ' has a fixed width.
  With txtASRSpinner
    .Top = 0
    .Height = lngCtrlHeight
    .Width = lngCtrlWidth - lngUpDnWidth
    .Left = IIf(giSpinnerPosition = vbLeftJustify, lngUpDnWidth, 0)
  End With
  
  With updnASRSpinner
    .Top = 0
    .Height = lngCtrlHeight
    .Left = IIf(giSpinnerPosition = vbLeftJustify, 0, lngCtrlWidth - lngUpDnWidth)
  End With
  
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("BackColor", txtASRSpinner.BackColor, vbWhite)
  Call PropBag.WriteProperty("ForeColor", txtASRSpinner.ForeColor, Ambient.ForeColor)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Increment", updnASRSpinner.Increment, 1)
  Call PropBag.WriteProperty("MaximumValue", updnASRSpinner.Max, 10)
  Call PropBag.WriteProperty("MinimumValue", updnASRSpinner.Min, 0)
  Call PropBag.WriteProperty("SpinnerPosition", giSpinnerPosition, vbRightJustify)
  Call PropBag.WriteProperty("Alignment", txtASRSpinner.Alignment, vbRightJustify)
  Call PropBag.WriteProperty("Text", txtASRSpinner.Text, "")

End Sub




Public Sub Refresh()
  ' Refresh the control's display
  UserControl_Resize
  
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "ASRSpinner value as a string."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "3c"
  ' Return the textbox control's text property.
  Text = txtASRSpinner.Text

End Property

Public Property Let Text(ByVal psNewText As String)
  ' Validate the new value.
  gfIgnoringInput = False
'  txtASRSpinner_Change
  
  ' Set the textbox control's text property.
  txtASRSpinner.Text = psNewText
  
  PropertyChanged "Text"

End Property


Public Property Get Value() As Long
  ' Return the spinner's value.
  Value = Val(txtASRSpinner.Text)

End Property

Public Property Let Value(ByVal pLngNewValue As Long)
  ' Set the spinner's value.
  gfIgnoringInput = False
  
  ' Set the textbox control's text property.
  txtASRSpinner.Text = Trim(Str(pLngNewValue))
  
  PropertyChanged "Value"

End Property
