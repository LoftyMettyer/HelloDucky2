VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FocusCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const WM_LBUTTONDOWN = &H201

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Declare Windows API functions
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

'Declare local variables
Private gfInitialising As Boolean

'Local variables to hold property values
Private gObjContainer As Object
Private gObjControl As Object
Private gObjTab As Object
Private gObjPage As Object
Private gLngControlType As Long
Private giCheckBoxWidth As Integer
Private gsCheckBoxOffset As String

Public Sub GotFocus(pObjThisControl As Object)
  Dim ControlRect As RECT
  Dim ContainerRect As RECT
  
  'Check if a control already has focus
  If Not Control Is Nothing Then
    'If this control already has focus, then exit
    If Not pObjThisControl Is Nothing Then
      If Control Is pObjThisControl Then
        Exit Sub
      End If
    End If
    
    'Call lost focus for control with focus
    LostFocus
  End If
  
  'Store pointer to the focus control
  Set Control = pObjThisControl
  
  'Get the control type
  ControlType = GetControlType(pObjThisControl)
    
  With frmSysMgr.tbMain
    If ControlType = ctlTab Then
      .Tools("ID_Cut").Enabled = False
      .Tools("ID_Copy").Enabled = False
      .Tools("ID_ScreenObjectDelete").Enabled = pObjThisControl.Tabs.Count > 1
    Else
      .Tools("ID_Cut").Enabled = True
      .Tools("ID_Copy").Enabled = True
      .Tools("ID_ScreenObjectDelete").Enabled = True
    End If
  End With
  
  'If this control is TabStrip then
  'update the properties window and exit.
  If ControlType = ctlTab Then
    SetProperties
    UpdateStatusBar
    Exit Sub
  End If
  
  'Set initialising flag
  gfInitialising = True
  
  'Lock control parent window to prevent display updates
  'UI.LockWindow Control.Parent.hWnd
  
  'If this control has a font property, set
  'the controls parent font to the controls font.
  If GetProperties(ControlType) And propFont Then
    Set Control.Parent.Font = Control.Font
  End If
  
  'Set the sizer box container to the focus control container
  Set Control.Parent.picSizer.Container = Control.Container
  
  'Set initial size and position of sizer box
  Control.Parent.picSizer.Move _
    Control.Left - Control.Parent.XOffSet, _
    Control.Top - Control.Parent.YOffSet, _
    Control.Width + (Control.Parent.XOffSet * 2), _
    Control.Height + (Control.Parent.YOffSet * 2)
  
  'Get position of sizer box and
  'active tab page, relative to entire screen.
  Call GetWindowRect(Control.Parent.picSizer.hWnd, ControlRect)
  Call GetWindowRect(ActivePage.hWnd, ContainerRect)
  
  'Set container for sizer box
  Set Control.Parent.picSizer.Container = ActivePage
  
  'Reset position of sizer box, relative to the entire screen
  With Control.Parent.picSizer
    .Top = (ControlRect.Top - ContainerRect.Top) * Screen.TwipsPerPixelY
    .Left = (ControlRect.Left - ContainerRect.Left) * Screen.TwipsPerPixelX
    .Visible = True
    .ZOrder 0
    .SetFocus
  End With
  
  'Store original container of the focus control
  Set Container = Control.Container
  'Set container of the focus control to the sizer box
  Set Control.Container = Control.Parent.picSizer
  
  'Set position of the focus control within the sizer box
  With Control
    .Top = 0
    .Left = 0
    .Visible = True
    .ZOrder 0
    .Refresh
  End With
  
  'Unlock control parent window
  'UI.UnlockWindow
  
  'Update the properties window
  SetProperties
  
  'Disable the focus control
  Control.Enabled = False
  
  'Update status bar, to show position
  'and size of the focus control
  UpdateStatusBar
  
  'Clear the initialising flag
  gfInitialising = False

End Sub

Public Sub LostFocus()
  Dim ctlThisControl As Control
  Dim ControlRect As RECT, ContainerRect As RECT
  Dim iLoop As Integer
  
  'Check if a control has focus
  If Not Control Is Nothing Then
    With frmScrProps
      For iLoop = .fraPropEdit.LBound To .fraPropEdit.UBound
        .fraPropEdit(iLoop).Visible = False
      Next iLoop
      .Caption = "Properties"
      .grdProperties.Visible = False
    End With
    
    If Control.Name = "TabStrip1" Then
      Set Control = Nothing
      Exit Sub
    End If
    
    'Lock parent form to prevent any updating
    Call UI.LockWindow(Control.Parent.hWnd)
    
    Set Container = ActivePage
  
    'Loop through controls on the active tab page
    For Each ctlThisControl In Control.Parent.Controls
        
      If ctlThisControl.Container Is ActivePage Then
         
        'Check if this control could container the focus control
        If InContainer(Control.Parent.picSizer, ctlThisControl) Then
            
          'Check if the focus control already has a container
          If TypeOf Container Is Frame Then
              
            'Check if this control is contained within the container
            If InContainer(ctlThisControl.Control, Container) Then
                          
              'This control is a better container for the focus control
              Set Container = ctlThisControl.Control
              
            End If
            
          Else
          
              'This control is a possible container for the focus control
              Set Container = ctlThisControl
              
          End If
        End If
      End If
    Next
    
    Set ctlThisControl = Nothing
    
    'Check if the focus control is contained in a frame
    If TypeOf Container Is Frame Then
    
      Call GetWindowRect(Control.Parent.picSizer.hWnd, ControlRect)
      Call GetWindowRect(Container.hWnd, ContainerRect)
      Set Control.Parent.picSizer.Container = Container
      
      'Adjust the focus controls position within it's container
      With Control.Parent.picSizer
        .Top = (ControlRect.Top - ContainerRect.Top) * Screen.TwipsPerPixelY
        .Left = (ControlRect.Left - ContainerRect.Left) * Screen.TwipsPerPixelX
      End With
      
    End If
    
    'Set the focus control container
    Set Control.Container = Container
    
    'Adjust the focus controls position within it's container
    With Control
    
      .Top = .Parent.picSizer.Top + .Parent.YOffSet
      .Left = .Parent.picSizer.Left + .Parent.XOffSet
      .Visible = True
      .ZOrder 0
      .Enabled = True
      
    End With
    
    'Restore the form default font
    Set Control.Parent.Font = Control.Parent.DefaultFont
    
    'Set the sizer box container to the parent form
    Set Control.Parent.picSizer.Container = Control.Parent
    Control.Parent.picSizer.Visible = False
    
    'Clear the status bar
    Control.Parent.StatusBar1.Panels.Clear
    
    'Re-enable updating of tab page
    Call UI.UnlockWindow
    
  End If
  
  If Not Control Is Nothing Then
    frmSysMgr.RefreshMenu
    Set Control = Nothing
  End If
    
  With frmScrProps
  
    ' Reset the property edit textboxes
    For iLoop = .txtPropEdit.LBound To .txtPropEdit.UBound
      .txtPropEdit(iLoop).Text = vbNullString
    Next iLoop
    
    .Caption = "Properties"
    .grdProperties.Visible = False
    
  End With
  
End Sub

Private Function InContainer(pCtlChkControl As Control, pCtlChkContainer As Control) As Boolean
  On Error GoTo ErrorTrap
  
  Dim wndRect1 As RECT
  Dim wndRect2 As RECT

  'Check if container is a frame
  If TypeOf pCtlChkContainer Is Frame Then
  
    'Get position & size of control
    Call GetWindowRect(pCtlChkControl.hWnd, wndRect1)
    'Get position & size of container
    Call GetWindowRect(pCtlChkContainer.hWnd, wndRect2)
    
    'Check if the control fits within the container
    If (wndRect1.Top > wndRect2.Top) _
        And (wndRect1.Bottom < wndRect2.Bottom) _
        And (wndRect1.Left > wndRect2.Left) _
        And (wndRect1.Right < wndRect2.Right) Then
      
      InContainer = True    'Control fits container
    Else
      InContainer = False   'Control does not fit container
    End If
    
  Else
    InContainer = False   'Container is not a frame
  End If
  
  Exit Function
  
ErrorTrap:
  InContainer = False
  Err = False
  
End Function

Public Sub Refresh()
  
  Control.Refresh
  
  With Control.Parent.picSizer
    
    .Height = Control.Height + (Control.Parent.YOffSet * 2)
    .Width = Control.Width + (Control.Parent.XOffSet * 2)
    .ZOrder 0
    .Refresh
  
  End With
  
  Control.Refresh

End Sub

Public Function Resize() As Boolean
  '
  ' Resize the current control to fit the size of the picSizer that borders it.
  '
  
  ' Trap if we cannot resize the control. eg. an option group cannot be resized.
  On Error GoTo ErrorTrap
  
  ' Only resize the control if we are not just initialising the focus control.
  If Not gfInitialising Then
  
    ' Sizer box has been resized, so resize the control
    If Not HasLockedWidth Then
      Control.Width = Control.Parent.picSizer.ScaleWidth
      
      ' Update the properties window Width and Left values.
      SetProperty "Width", Control.Width
      SetProperty "Left", Control.Parent.picSizer.Left + Control.Parent.XOffSet
      
    End If
    
    If Not HasLockedHeight Then
      Control.Height = Control.Parent.picSizer.ScaleHeight
    
      ' Update the properties window Height and Top values.
      SetProperty "Height", Control.Height
      SetProperty "Top", Control.Parent.picSizer.Top + Control.Parent.YOffSet
    
    End If
  End If
  
  Resize = True
 
  Exit Function

ErrorTrap:
  Resize = False
  Err = False
  
End Function

Public Sub SetProperties()
  Dim grdProperties As MSFlexGrid
  Dim lngProps As Long
  
  If Control Is Nothing Then
    Exit Sub
  End If
  
  UI.LockWindow frmScrProps.hWnd

  frmScrProps.Initialising = True
  Set frmScrProps.CurrentScreen = Control.Parent
  
  lngProps = GetProperties(ControlType)
  
  If (ControlType And ctlTab) Then
    
    If Control.Tabs.Count > 0 Then
      frmScrProps.Caption = "Properties - Page " & _
        Control.SelectedItem.Index
    Else
      frmScrProps.Caption = "Properties - Screen"
    End If
  
  Else
    frmScrProps.Caption = "Properties - " & Control.Tag
  End If
  
  Set grdProperties = frmScrProps.grdProperties
  
  'Format properties grid
  With grdProperties
    .Clear
    .RowHeightMin = 17 * Screen.TwipsPerPixelY
    .ColWidth(2) = 0
    .Rows = 1
    .FormatString = "<Property" & Space(22) & "|<Value"
    .Row = 0
    .Col = 0
    .CellFontBold = True
    .Col = 1
    .CellFontBold = True
    
    'Caption property
    If (lngProps And propCaption) Then
      If ControlType = ctlTab Then
        If Control.Tabs.Count > 0 Then
          .AddItem "Caption" & vbTab & Control.SelectedItem.Caption & vbTab & propCaption
          .Row = .Rows - 1
          .RowData(.Row) = propCtlText
        End If
      Else
        .AddItem "Caption" & vbTab & UI.GetCaption(Control) & vbTab & propCaption
        .Row = .Rows - 1
        .RowData(.Row) = propCtlText
      End If
    End If
    
    'BackColor property
    If (lngProps And propBackColor) Then
      .AddItem "Background Colour" & vbTab & "Background Colour" & vbTab & propBackColor
      .Row = .Rows - 1
      .Col = 1
      ' The MSFlexGrid control has a problem when setting the backcolour to Black (&H0)
      If Control.BackColor = vbBlack Then
        .CellBackColor = &H1001001
      Else
        .CellBackColor = Control.BackColor
      End If
      .CellForeColor = UI.GetInverseColor(Control.BackColor)
      .RowData(.Row) = propCtlCommand
    End If
    
    'Border property
    If (lngProps And propBorderStyle) Then
      .AddItem "Border Style" & vbTab & _
        IIf(Control.BorderStyle = 0, "None", "Fixed Single") & _
          vbTab & propBorderStyle
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCombo
    End If

    'Display type property
    If (lngProps And propDisplayType) Then
      .AddItem "Display Type" & vbTab & _
        IIf(Control.DisplayType = 1, "Icon", "Content") & _
        vbTab & propDisplayType
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCombo
    End If

    'Font properties
    If (lngProps And propFont) Then
      .AddItem "Font" & vbTab & Control.Font.Name & ", " & Control.Font.Size & vbTab & propFont
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCommand
      .AddItem "Font Bold" & vbTab & IIf(Control.Font.Bold, "True", "False") & vbTab & propBold
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCheck
      .AddItem "Font Italic" & vbTab & IIf(Control.Font.Italic, "True", "False") & vbTab & propItalic
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCheck
      .AddItem "Font Strikethrough" & vbTab & IIf(Control.Font.Strikethrough, "True", "False") & vbTab & propStrikeThru
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCheck
      .AddItem "Font Underline" & vbTab & IIf(Control.Font.Underline, "True", "False") & vbTab & propUnderline
      .Row = .Rows - 1
      .RowData(.Row) = propCtlCheck
    End If
      
    'ForeColor property
    If (lngProps And propForeColor) Then
      .AddItem "Foreground Colour" & vbTab & "Foreground Colour" & vbTab & propForeColor
      .Row = .Rows - 1
      .Col = 1
      ' The MSFlexGrid control has a problem when setting the backcolour to Black (&H0)
      If Control.ForeColor = vbBlack Then
        .CellBackColor = &H1001001
      Else
        .CellBackColor = Control.ForeColor
      End If
      .CellForeColor = UI.GetInverseColor(Control.ForeColor)
      .RowData(.Row) = propCtlCommand
    End If
    
    'Picture property
    If (lngProps And propPicture) Then
      recPictEdit.Index = "idxID"
      recPictEdit.Seek "=", Val(Control.DataField)
      If recPictEdit.NoMatch Then
        .AddItem "Picture" & vbTab & vbNullString & vbTab & propPicture
      Else
        .AddItem "Picture" & vbTab & recPictEdit!Name & vbTab & propPicture
      End If
      .Row = .Rows - 1
'      .RowData(.Row) = propCtlCheck
      .RowData(.Row) = propCtlCommand
    End If
    
    ' Left property
    If (lngProps And propLeft) Then
      .AddItem "Left" & vbTab & Control.Parent.picSizer.Left + Control.Parent.XOffSet & vbTab & propLeft
      .Row = .Rows - 1
      .RowData(.Row) = propCtlIntText
    End If
    
    ' Top property
    If (lngProps And propTop) Then
      .AddItem "Top" & vbTab & Control.Parent.picSizer.Top + Control.Parent.YOffSet & vbTab & propTop
      .Row = .Rows - 1
      .RowData(.Row) = propCtlIntText
    End If
    
    ' Width property
    If (lngProps And propWidth) Then
      .AddItem "Width" & vbTab & Control.Width & vbTab & propWidth
      .Row = .Rows - 1
      .RowData(.Row) = propCtlIntText
    End If
    
    ' Height property
    If (lngProps And propHeight) Then
      .AddItem "Height" & vbTab & Control.Height & vbTab & propHeight
      .Row = .Rows - 1
      .RowData(.Row) = propCtlIntText
    End If
  
    .ColWidth(1) = .Width - .ColWidth(0)
    .Row = 0
    .Visible = True
  End With
    
  With grdProperties
    If .Rows > 1 Then
      .Row = 1
      .Col = 1
    End If
  End With
  
  Set grdProperties = Nothing
  frmScrProps.Initialising = False
  
  UI.UnlockWindow

End Sub

Public Sub UpdateStatusBar()
  Dim ThisPanel As Panel
  Dim iPanel As Integer

  If Not Control Is Nothing Then
    
    With Control.Parent.StatusBar1
      
      If .Panels.Count < 3 Then
        .Panels.Clear
        
        For iPanel = 1 To 3
          Set ThisPanel = .Panels.Add(iPanel, , , sbrNormal)
          ThisPanel.Alignment = sbrLeft
          ThisPanel.AutoSize = sbrSpring
          Set ThisPanel = Nothing
        Next iPanel
      
      End If
      
      If ControlType = ctlTab Then
            .Panels(1).Text = "Name: " & Control.Tabs(Control.SelectedItem.Key).Caption
      Else
        .Panels(1).Text = "Name: " & Control.Tag
        .Panels(2).Text = "Position  X: " & _
          Trim(Str(Control.Parent.picSizer.Left + Control.Parent.XOffSet)) & _
          "  Y: " & Trim(Str(Control.Parent.picSizer.Top + Control.Parent.YOffSet))
        .Panels(3).Text = "Size  X: " & _
          Trim(Str(Control.Width)) & _
          "  Y: " & Trim(Str(Control.Height))
      End If
    End With
  End If
End Sub

Public Property Get Container() As Object

  ' Return the focussed control''s container
  Set Container = gObjContainer

End Property

Public Property Set Container(ByVal pObjData As Object)

  ' Set the focussed control's container.
  Set gObjContainer = pObjData
  
End Property

Public Property Get Control() As Object

  ' Return the focussed control.
  Set Control = gObjControl
  
End Property

Public Property Set Control(ByVal pObjData As Object)

  ' Set the focussed control.
  Set gObjControl = pObjData
  
  If gObjControl Is Nothing Then
    
    Set gObjPage = Nothing
    Set gObjTab = Nothing
  
  Else
    
    Set gObjTab = Control.Parent.TabStrip1.SelectedItem
    
    If Not gObjTab Is Nothing Then
      Set gObjPage = Control.Parent.picTabPage(gObjTab.Tag)
    Else
      Set gObjPage = Nothing
    End If
  
  End If
  
End Property

Public Property Get ControlType() As ControlTypes

  ' Return the control type.
  ControlType = gLngControlType
  
End Property

Public Property Let ControlType(pLngType As ControlTypes)

  ' Set the control type.
  gLngControlType = pLngType
  
End Property

Public Property Get MaxHeight() As Integer
  '
  ' Return the maximum allowed height for the current control, or zero if
  ' there is no maximum.
  '
  MaxHeight = IIf(HasLockedHeight, MinHeight, 0)
  
End Property

Public Property Get MinHeight() As Integer
  '
  ' Return the minimum allowed height for the current control, or zero if
  ' there is no minimum.
  '
  Dim iMinHeight As Integer

  iMinHeight = 0
  
  Select Case ControlType
    Case ctlLabel
      iMinHeight = Control.Parent.TextHeight(Control.Text)
      iMinHeight = IIf(iMinHeight < 195, 195, iMinHeight)
    
    Case ctlText
'      iMinHeight = Control.Height
    
    Case ctlCombo
      iMinHeight = Control.Height
    
    Case ctlSpin
      iMinHeight = Control.Height
    
    Case ctlCheck
      iMinHeight = UI.GetCharHeight(Control.Parent.hDC)
      If iMinHeight < 285 Then iMinHeight = 285
    
    Case ctlRadio
      iMinHeight = Control.Height
    
    Case ctlOLE
      iMinHeight = 32 * Screen.TwipsPerPixelY
      iMinHeight = (UI.GetSystemMetrics(SM_CYICON) _
        + (UI.GetSystemMetrics(SM_CYICONSPACING) \ 2)) _
        * Screen.TwipsPerPixelY
    
    Case ctlFrame
      iMinHeight = 600
    
    Case ctlImage
      iMinHeight = 360
  End Select
  
  MinHeight = iMinHeight
  
End Property

Public Property Get MaxWidth() As Integer
  '
  ' Return the maximum allowed width for the current control, or zero if
  ' there is no maximum.
  '
  MaxWidth = IIf(HasLockedWidth, MinWidth, 0)

End Property

Public Property Get MinWidth() As Integer
  '
  ' Return the minimum allowed width for the current control, or zero if
  ' there is no minimum.
  '
  Dim iMinWidth As Integer
  
  iMinWidth = 0
  
  Select Case ControlType
    Case ctlLabel
      iMinWidth = Control.Parent.TextWidth(Control.Text)
      iMinWidth = IIf(iMinWidth < 255, 255, iMinWidth)
    
    Case ctlText
      iMinWidth = 255
    
    Case ctlCombo
      iMinWidth = 525
    
    Case ctlSpin
      iMinWidth = 525
    
    Case ctlCheck
      iMinWidth = giCheckBoxWidth + _
        Control.Parent.TextWidth(gsCheckBoxOffset & Control.Caption)
    
    Case ctlRadio
      iMinWidth = Control.Width
    
    Case ctlOLE
      iMinWidth = 32 * Screen.TwipsPerPixelX
      iMinWidth = (UI.GetSystemMetrics(SM_CXICON) _
        + UI.GetSystemMetrics(SM_CXICONSPACING)) _
        * Screen.TwipsPerPixelX
    
    Case ctlFrame
      iMinWidth = 570
    
    Case ctlImage
      iMinWidth = 360
  End Select
  
  MinWidth = iMinWidth
  
End Property

Public Property Get ActivePage() As VB.PictureBox
  
  ' Return the active page.
  Set ActivePage = gObjPage
  
End Property

Public Property Get ActiveTab() As ComctlLib.Tab

  ' Return the active page tab.
  Set ActiveTab = gObjTab

End Property

Private Sub Class_Initialize()

  Set gObjControl = Nothing
  Set gObjContainer = Nothing
  Set gObjTab = Nothing
  Set gObjPage = Nothing
  
  ' Initialise the global variables that define how much space we
  ' need to add onto the textwidth of a checkbox caption, to give
  ' the correct total width of the control.
  giCheckBoxWidth = 200
  gsCheckBoxOffset = "W"
  
End Sub

Private Sub Class_Terminate()

  'Release internal objects
  Set gObjControl = Nothing
  Set gObjContainer = Nothing
  Set gObjTab = Nothing
  Set gObjPage = Nothing

End Sub

Public Sub SetSize()

  Select Case ControlType
    ' Set the focusCtrl to have the minimum width and height for labels.
    Case ctlLabel
      Control.Width = MinWidth
      Control.Height = MinHeight
    
    ' Set the focusCtrl to have the minimum height for textboxes that are not multi-line.
    ' Do not set width.
    Case ctlText
      Control.Height = UI.GetCharHeight(Control.Parent.hDC) + (2 * Control.Parent.YOffSet)
    
    Case ctlCombo
    
    ' Set the focusCtrl to have the minimum height for spinners.
    ' Do not set width.
    Case ctlSpin
      Control.Height = UI.GetCharHeight(Control.Parent.hDC) + (2 * Control.Parent.YOffSet)
    
    ' Set the focusCtrl to have the minimum width and height for check boxes.
    Case ctlCheck
      Control.Width = MinWidth
      Control.Height = MinHeight
    
    Case ctlRadio
    Case ctlOLE
    Case ctlFrame
    Case ctlImage
  End Select

  ' Ensure that the size of the picSizer is corrent
  Refresh
  
End Sub

Public Function HasLockedSize() As Boolean
  '
  ' Return True if the size of the current control is automatically
  'set and cannot be manually altered. eg. Label or option group.
  '
  HasLockedSize = HasLockedWidth And HasLockedHeight
  
End Function
Public Function HasLockedWidth() As Boolean
  '
  ' Return True if the width of the current control is set automatically
  ' and cannot be manually altered.
  '
  HasLockedWidth = False
  
  Select Case ControlType
    Case ctlLabel
      HasLockedWidth = True
    
    Case ctlText
    Case ctlCombo
    Case ctlSpin
    
    Case ctlCheck
      HasLockedWidth = True
    
    Case ctlRadio
      HasLockedWidth = True
  
    Case ctlOLE
    Case ctlFrame
    Case ctlImage
  
  End Select

End Function
Public Function HasLockedHeight()
  '
  ' Return True if the height of the current control is set automatically
  ' and cannot be manually altered.
  '
  HasLockedHeight = False
  
  Select Case ControlType
    Case ctlLabel
      HasLockedHeight = True
      
    Case ctlText
'      HasLockedHeight = True
      
    Case ctlCombo
      HasLockedHeight = True
      
    Case ctlSpin
      HasLockedHeight = True
      
    Case ctlCheck
      HasLockedHeight = True
      
    Case ctlRadio
      HasLockedHeight = True
      
    Case ctlOLE
    Case ctlFrame
    Case ctlImage
  End Select

End Function


Private Sub SetProperty(psProperty As String, pvValue As Variant)
  '
  ' Updates the given property in the frmScrProps grid with
  ' the given value.
  '
  Dim iLoop As Integer
  
  For iLoop = 1 To frmScrProps.grdProperties.Rows
  
    If frmScrProps.grdProperties.TextMatrix(iLoop, 0) = psProperty Then
      frmScrProps.grdProperties.TextMatrix(iLoop, 1) = pvValue
      Exit Sub
    End If
    
  Next iLoop

End Sub
