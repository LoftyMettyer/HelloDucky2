VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.UserControl COA_Lookup 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   2460
   ToolboxBitmap   =   "COA_Lookup.ctx":0000
   Begin XtremeSuiteControls.PushButton cmdDrop 
      Height          =   360
      Left            =   1530
      TabIndex        =   0
      Top             =   0
      Width           =   285
      _Version        =   851969
      _ExtentX        =   503
      _ExtentY        =   644
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      MultiLine       =   0   'False
      DrawFocusRect   =   0   'False
      PushButtonStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtText 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1815
      _Version        =   851969
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   0
   End
End
Attribute VB_Name = "COA_Lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbSizing As Boolean
Private mcoItems As New Collection
Private lhndDrop As Long
Private asLookupItems() As Variant

'Default Property Values:
Const m_def_Caption = ""
Const m_def_Mandatory = 0

'Property Variables:
Dim m_Caption As String
Dim m_Mandatory As Boolean
Dim mbSelect As Boolean
Dim mbInsert As Boolean

'Event Declarations:
Public Event Change() 'MappingInfo=txtText,txtText,-1,Change
Public Event NewEntry()
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeySpace Then
    ShowDropdown 1, 0, 0, 0
    KeyCode = 0
  End If

  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub cmdDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ShowDropdown Button, Shift, X, Y
End Sub

Private Sub ShowDropdown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim lpContainer As RECT
  Dim lpForm As RECT
  Dim lpDesk As RECT
  Dim lhndDesk As Long
  Dim lCount As Long
  Dim lX As Long, lY As Long, lHeight As Long, lWidth As Long
  Dim lIndex As Long
    
  Dim xx As Integer, yy As Integer, totwidth As Long, itmX As ListItem
  
  If Button = vbLeftButton Then
    If Not frmDrop.Visible Then
      RaiseEvent Click
      If Not mbSelect Then
        Unload frmDrop
        txtText.SetFocus
        Exit Sub
      End If

      Screen.MousePointer = vbHourglass

      lhndDrop = frmDrop.hwnd
      lhndDesk = GetDesktopWindow
      
      'Get Desktop co-ordinates
      Call GetWindowRect(lhndDesk, lpDesk)
      
      Call GetClientRect(lhndDrop, lpForm)
      Call GetWindowRect(UserControl.hwnd, lpContainer)
      
      lX = lpContainer.Left
      lY = lpContainer.Bottom
      lWidth = lpForm.Right
      lHeight = lpForm.Bottom
        
      If (lY + lHeight) > lpDesk.Bottom Then
        lY = lpContainer.Top - lHeight
      End If
      
      If (lX + lWidth) > lpDesk.Right Then
        lX = lpContainer.Right - lWidth
      End If
        
      Call MoveWindow(lhndDrop, lX, lY, lWidth, lHeight, 0)
      
      'Set the Form Values
      With frmDrop
        If Mandatory Then
          .cmdClear.Enabled = False
        ElseIf Len(txtText) = 0 Then
          .cmdClear.Enabled = False
        End If
        
        .cmdNew.Enabled = mbInsert
        
'Screen.MousePointer = vbHourglass
        
        ' RH 14/08/00 - Speed the control up...try turning sort off while
        '               loading the listview...then turn it on afterwards
        .lsvList.Sorted = False
        
        'Load Headers
        .lsvList.ListItems.Clear
        .lsvList.ColumnHeaders.Clear
        For xx = 1 To UBound(asLookupItems, 1)
          .lsvList.ColumnHeaders.Add , , asLookupItems(xx, 1)
          .lsvList.ColumnHeaders(xx).Width = TextWidth(.lsvList.ColumnHeaders(xx).Text) + 220
        Next xx

        'Load Data into grid from array
        For yy = 2 To UBound(asLookupItems, 2)
          Set itmX = .lsvList.ListItems.Add()
          For xx = 1 To UBound(asLookupItems, 1)
            If xx = 1 Then
              itmX.Text = asLookupItems(xx, yy)
              If .lsvList.ColumnHeaders(xx).Width < TextWidth(.lsvList.ListItems(yy - 1).Text) + 220 Then _
              .lsvList.ColumnHeaders(xx).Width = TextWidth(.lsvList.ListItems(yy - 1).Text) + 220
              'If the current item is the one in the lookup then flag to select it
              If asLookupItems(xx, yy) = txtText.Text Then lIndex = (yy - 1)
            Else
              If IsNull(asLookupItems(xx, yy)) Then
                itmX.SubItems(xx - 1) = " "
                If .lsvList.ColumnHeaders(xx).Width < TextWidth(" ") + 220 Then _
                  .lsvList.ColumnHeaders(xx).Width = TextWidth(" ") + 220
              Else
                itmX.SubItems(xx - 1) = asLookupItems(xx, yy)
                If .lsvList.ColumnHeaders(xx).Width < TextWidth(asLookupItems(xx, yy)) + 220 Then _
                  .lsvList.ColumnHeaders(xx).Width = TextWidth(asLookupItems(xx, yy)) + 220
              End If
            End If
          Next xx
        Next yy
              
        'If any item has been flagged, select it
        If lIndex > 0 Then .lsvList.ListItems(lIndex).Selected = True

        'Change Width of the grid
        For yy = 1 To .lsvList.ColumnHeaders.Count
          totwidth = totwidth + .lsvList.ColumnHeaders(yy).Width
        Next yy
        
        'Expand the width for margin depending on if vert scroll reqd
        If .lsvList.ListItems.Count > 10 Then
          .lsvList.Width = totwidth + 320
        Else
          .lsvList.Width = totwidth + 90
        End If
        
        'If listview is too wide, limit the width
        If (.lsvList.Width) > 5000 Then .lsvList.Width = 5000
      
        'Size the buttons and the form to suit the grid width
        .cmdCancel.Left = .lsvList.Left + .lsvList.Width + 125
        .cmdClear.Left = .lsvList.Left + .lsvList.Width + 125
        .cmdNew.Left = .lsvList.Left + .lsvList.Width + 125
        .cmdSelect.Left = .lsvList.Left + .lsvList.Width + 125
        .Width = .cmdCancel.Left + 1215 + 100

        ' If no entries, then disable select button !
        If .lsvList.ListItems.Count = 0 Then
          .cmdSelect.Enabled = False
          .cmdSelect.Default = False
          If mbInsert Then
            .cmdNew.Default = True
          Else
            .cmdCancel.Default = True
          End If
        End If
        
        .lsvList.Sorted = True
        
        'Show the form and set focus on the grid
        .Visible = True
        .lsvList.SetFocus

      End With
      
      cmdDrop.Enabled = False
      tmrTimer.Enabled = True
    Else
      Unload frmDrop
      cmdDrop.Enabled = True
      txtText.SetFocus
    End If
  
  End If

  Screen.MousePointer = vbDefault

End Sub

Private Sub tmrTimer_Timer()

  If lhndDrop <> GetActiveWindow Then
    Unload frmDrop
    tmrTimer.Enabled = False
    txtText.SetFocus
  End If

End Sub

Private Sub txtText_GotFocus()

  'If txtText gets focus from user tabbing or clicking with mouse, pass focus
  'on to cmdDrop
  'If not, then change the data first, then pass focus on to cmdDrop
  With frmDrop
    If .Selected Then
      If .Item = "Add New Table Entry" Then
        RaiseEvent NewEntry
      Else
        txtText = .Item
      End If
    End If
    .Selected = False
  End With

  cmdDrop.Enabled = True
  cmdDrop.SetFocus

End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'To prevent the mouse icon from turning into an I-bar, and causing user to
  'think they should be able to type into the control
  Screen.MousePointer = vbArrow
  
End Sub

Private Sub UserControl_Resize()

  Dim lngButtonWidth As Long
  Dim lngCtrlWidth As Long
  Dim lngCtrlHeight As Long


  ' Prevent user from sizing control smaller than the minimum size allowed
  If Not mbSizing Then
    
    ' Do not let the user make the control too small.
    With UserControl
      If (.Width < 390) Then
        .Width = 390
      End If
      
      If (.Height < 315) Then
        .Height = 315
      End If
      
      lngCtrlWidth = .Width
      lngCtrlHeight = .Height
    End With
    
    lngButtonWidth = cmdDrop.Width
    
    ' Resize the text and button controls as our custom control is resized. NB. the button control
    ' has a fixed width.
    With txtText
      .Top = 0
      .Height = lngCtrlHeight
      .Width = (lngCtrlWidth - cmdDrop.Width) + 7
      .Left = 0
    End With
    
    With cmdDrop
      .Top = 0
      .Height = lngCtrlHeight
      .Left = lngCtrlWidth - lngButtonWidth
    End With
  
    If UserControl.Height < txtText.Height Then
      UserControl.Height = txtText.Height
    End If
  End If
      
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  txtText.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Enabled
Public Property Get Enabled() As Boolean
  Enabled = txtText.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
  'Enable/Disable whole control depending upon the Enabled property
  txtText.Enabled() = New_Enabled
  cmdDrop.Enabled = New_Enabled
  PropertyChanged "Enabled"
  UserControl.Enabled = New_Enabled
  If Not New_Enabled Then
    txtText = ""
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Font
Public Property Get Font() As Font
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    'UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = txtText.Font.Bold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtText.Font.Bold = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = txtText.Font.Italic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtText.Font.Italic = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = txtText.Font.Size
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtText.Font.Size = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = txtText.Font.Strikethrough
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtText.Font.Strikethrough = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = txtText.Font.Underline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtText.Font.Underline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtText.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtText.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtText.Font.Bold = PropBag.ReadProperty("FontBold", 0)
    txtText.Font.Italic = PropBag.ReadProperty("FontItalic", 0)
    txtText.Font.Strikethrough = PropBag.ReadProperty("FontStrikethru", 0)
    txtText.Font.Underline = PropBag.ReadProperty("FontUnderline", 0)
    txtText.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set txtText.Font = PropBag.ReadProperty("FontName", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    txtText.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtText.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtText.SelText = PropBag.ReadProperty("SelText", "")
    txtText.Text = PropBag.ReadProperty("Text", "")
    txtText.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtText.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", txtText.Enabled, True)
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtText.Font.Bold, 0)
    Call PropBag.WriteProperty("FontItalic", txtText.Font.Italic, 0)
    Call PropBag.WriteProperty("FontSize", txtText.Font.Size, 0)
    Call PropBag.WriteProperty("FontStrikethru", txtText.Font.Strikethrough, 0)
    Call PropBag.WriteProperty("FontUnderline", txtText.Font.Underline, 0)
    Call PropBag.WriteProperty("ForeColor", txtText.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FontName", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("SelLength", txtText.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtText.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtText.SelText, "")
    Call PropBag.WriteProperty("Text", txtText.Text, "")
    Call PropBag.WriteProperty("ToolTipText", txtText.ToolTipText, "")
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Font
Public Property Get FontName() As Font
    Set FontName = txtText.Font
End Property

Public Property Set FontName(ByVal New_FontName As Font)
    Set txtText.Font = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub txtText_Change()
    On Error Resume Next
    RaiseEvent Change
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtText.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtText.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtText.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtText.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelText
Public Property Get SelText() As String
    SelText = txtText.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtText.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Text
Public Property Get Text() As String
    Text = txtText.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtText.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = txtText.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtText.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Mandatory = m_def_Mandatory
End Sub

Public Sub AddItem(sItem As String)

  mcoItems.Add sItem

End Sub

Public Sub Clear()

  Set mcoItems = Nothing
  Set mcoItems = New Collection

End Sub

Public Property Let AllowSelect(bSelect As Boolean)

  mbSelect = bSelect

End Property

Public Property Let AllowInsert(bInsert As Boolean)

  mbInsert = bInsert

End Property

Public Sub PassArray(vPassedArray As Variant)

  'Receive the array containing items to show in the lookup
  asLookupItems = vPassedArray

End Sub
