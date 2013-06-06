VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' VB IDE Safe Subclassing provided by...
'*************************************************************************************************
'* cSelfSub - self-subclassing class template
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'*************************************************************************************************

' While I assisted troubleshooting one of TerriTop's apps, I came up with one solution
' that didn't quite work for him, but he suggested I share it with PSC. So here it is.

' Purpose: Provide custom border colors for some common VB controls. The controls this
'   should work for are those that have static borders. In other words, during normal,
'   click, and enable events, the border does not change. VB Command Buttons are a
'   prime example of borders that do change, therefore, this wouldn't work for them.
'   Not all static border controls have been tested. The ones in this example have been
'   and I could include others by request, but the borders must be static, not dynamic.

' The complex part of this project was trying to determine how borders were drawn on
'   the controls. Some controls have actual border window styles (i.e., WS_Border,
'   WS_EX_ClientEdge, etc) and those can be drawn over during a WM_NCPaint message.
'   But others have their borders drawn on their client area or one of their children's
'   client area. Another exception is the ImageCombo which is an owner-drawn control.


Option Explicit

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const CODE_LEN      As Long = 240                                   'Thunk length in bytes
Private Const WNDPROC_OFF   As Long = &H30                                  'WndProc execution offset
Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
Private Const PAGE_RWX      As Long = &H40                                  'Allocate executable memory
Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Shutdown flag data index
Private Const IDX_HWND      As Long = 2                                     'hWnd data index
Private Const IDX_EBMODE    As Long = 3                                     'EbMode data index
Private Const IDX_CWP       As Long = 4                                     'CallWindowProc data index
Private Const IDX_SWL       As Long = 5                                     'SetWindowsLong data index
Private Const IDX_FREE      As Long = 6                                     'VirtualFree data index
Private Const IDX_ME        As Long = 7                                     'Owner data index
Private Const IDX_WNDPROC   As Long = 8                                     'Original WndProc data index
Private Const IDX_CALLBACK  As Long = 9                                     'zWndProc data index
Private Const IDX_BTABLE    As Long = 10                                    'Before table data index
Private Const IDX_ATABLE    As Long = 11                                    'After table data index
Private Const IDX_EBX       As Long = 14                                    'Data code index

Private z_Base              As Long                                         'Data pointer base
Private z_Funk              As Collection                                   'hWnd/thunk-address collection
Private z_TblEnd            As Long                                         'End of the vTable
Private z_Code(29)          As Currency                                     'Thunk machine-code initialised here

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------


Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_BORDER As Long = &H800000
Private Const WM_NCPAINT As Long = &H85
Private Const WM_PAINT As Long = &HF&
Private Const WM_DRAWITEM As Long = &H2B
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const SM_CYBORDER As Long = 6
Private Const SM_CYEDGE As Long = 46
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_3DDKSHADOW As Long = 21
Private Const COLOR_3DLIGHT As Long = 22
Private Const COLOR_WINDOWFRAME As Long = 6
Private Const LBS_NOINTEGRALHEIGHT As Long = &H100&
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400&
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Settings() As Long
Public Enum Specialcolors 'sneaky way of exposing public "constants" from a class
    bsSysDefault = -2
    bsAutoShade = -1
End Enum
Public Enum BorderStyleOptions
    [_bs_Max] = 3
    [_bsBackColor] = -3
    bsFlat1Color = 0    ' 1 color, 1 pixel border around control
    bsFlat2Color = 1    ' 2 color, 1 pixel border
    bsSunken = 2        ' 4 color, 2 pixel border
    bsRaised = 3        ' 4 color, 2 pixel border
'    bsNone = 4          ' no borders at all -- not coded/not sure it will be
End Enum
Public Enum vbControlType
    ctOther = 0         ' controls that have a true border (non-client area)
    ctComboBox = 1      ' some draw on client and some on non-client
    ctImageCombo = 2    ' owner-drawn by VB
    ctListBox = 3       ' control height can change when changing border styles
    ctTextBox = 4       ' when flat style, VB draws border on client area
End Enum

Private Borders As Collection

Public Sub SetBorder1(ByVal hwnd As Long, _
                    Optional ByVal ctrlType As vbControlType = ctOther, _
                    Optional ByVal Shadow As Long = vbButtonShadow)
  
  SetBorder1 hwnd, bsFlat1Color, ctrlType, Shadow, -1, -1, -1

End Sub

Private Sub SetBorder1(ByVal hwnd As Long, BorderStyle As BorderStyleOptions, _
                    Optional ByVal ctrlType As vbControlType = ctOther, _
                    Optional ByVal Shadow As Long = vbButtonShadow, _
                    Optional ByVal DarkShadow As Long = bsAutoShade, _
                    Optional ByVal Highlight As Long = bsAutoShade, _
                    Optional ByVal LightShadow As Long = bsAutoShade)
    
    ' The control will be subclassed to allow custom borders. Therefore it is
    ' best to call ResetBorder on any control loaded with Controls.Add before
    ' that control is removed if you sent that added control here.
    
    ' ///// Border Styles \\\\\
    ' bsFlat1Color. Solid 1-pixel border, 1 color (i.e., flat).
    '       Uses Shadow only
    ' bsFlat2Color. Left/Top borders are 1 color, right/bottom are another
    '       Uses Shadow & Highlight only
    ' bsSunken. Left/Top outer border are Shadow, Right/Bottom outer are HighLight
    '           Left/Top inner border are DarkShadow, Right/Bottom inner are LightShadow
    ' bsRaised. Left/Top outer border are HighLight, Right/Bottom outer are DarkShadow
    '           Left/Top inner border are LightShadow, Right/Bottom outer are Shadow
    
    ' ///// colors \\\\\ vb system colors can be passed
    ' Shadow: 2nd darkest of 4 color borders; color for a single color border
    ' DarkShadow: the darkest of 4 color borders
    ' LightShadow: 2nd lightest of 4 color borders
    ' Highlight: lightest of 4 color borders
    ' Special values for the above 4 colors
    '   -1 = AutoShade. DarkShadow, LightShadow & Highlight are shades of Shadow
    '           DarkShadow = Shadow darkened to 15% from black
    '           LightShadow = Shadow lightened by 85% of its lightest value (white)
    '           Highlight = Shadow lightened by 100% (or vbWhite)
    '   -2 = System colors: vb3DDKShadow, vbButtonShadow, vb3DLight, vbHighlight respectively
    '   -3 & -4 (Reserved) are used by the class to fake single borders on combo boxes
    
    ' ///// Control Type \\\\\
    ' Some controls have their borders drawn by VB on the control's client area whereas
    '   others are drawn in the non-client area as expected. Think of a form with no
    '   borders but you want borders so you draw it on the form (non-client area).
    '   VB combo boxes are very much like that scenario. Therefore, the control type
    '   needs to be known in advance so the class can handle those special cases.
    '   There are more special cases too & those known are handled herein
    ' ctComboBox: use for comboboxes and drivecombo
    ' ctImageCombo: use for the image combobox
    ' ctListBox: use for listboxes and file listboxes
    ' ctTextBox: use for the textbox control
    ' ctOther: use for other controls like treeview, listview, progressbar, etc
    
    
    ' sanity checks first & don't allow user to pass a reserved color code
    If BorderStyle < bsFlat1Color Or BorderStyle > [_bs_Max] Then Exit Sub
    If DarkShadow = -3 Or DarkShadow = -4 Then DarkShadow = 0
    If LightShadow = -3 Or LightShadow = -4 Then LightShadow = 0
    If Shadow = -3 Or Shadow = -4 Then Shadow = 0
    If Highlight = -3 Or Highlight = -4 Then Highlight = 0
    
    Dim Index As Long, lFlags As Long, isSubclassed As Boolean
    Dim lExStyle As Long, lStyle As Long, lOldStyle As Long
    Dim cHwnd As Long, wRect As RECT, cRect As RECT
    Dim borderSize As Long, borderSizeNew As Long
    ' flags for special case control scenarios
    Dim bIntegralHT As Boolean, bFlatTextBox As Boolean, bHasBorder As Boolean
    
    On Error Resume Next
    If Borders Is Nothing Then          ' first time thru
        Set Borders = New Collection
        Index = 1
        ReDim Settings(1 To 8)
        Borders.Add Index, "h" & hwnd
    Else
        Index = Borders("h" & hwnd)     ' been here before, is this hWnd already ours?
        If Index = 0 Then               ' nope, set it up
            Index = UBound(Settings) + 1
            ReDim Preserve Settings(1 To Index + 7)
            Borders.Add Index, "h" & hwnd
            If Err Then Err.Clear
        Else
            isSubclassed = True         ' yep, use cached settings
        End If
    End If
    
    ' cache colors & style settings
    Settings(Index) = CLng(BorderStyle Or (ctrlType * &H100&))
    Settings(Index + 1) = Shadow
    Settings(Index + 2) = Highlight
    Settings(Index + 3) = DarkShadow
    Settings(Index + 4) = LightShadow
    Settings(Index + 7) = hwnd        ' needed when re-indexing if control is unsubclassed
    If isSubclassed Then
        lStyle = Settings(Index + 5)
        lExStyle = Settings(Index + 6)
    Else
        lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        lStyle = GetWindowLong(hwnd, GWL_STYLE)
        Settings(Index + 5) = lStyle
        Settings(Index + 6) = lExStyle
    End If
    ' setwindowpos flags
    lFlags = SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER
    
    ' special case for combo boxes. We can't remove the borders on these, because
    ' they really aren't borders for the most part; they are drawn on the client
    ' ComboBox.Style=1 is an exception. It has borders come from its child's borders,
    ' but to make things a little bit easier we will treat all combos the same
    If ctrlType = ctComboBox Or ctrlType = ctImageCombo Then
        If BorderStyle = bsFlat2Color Or BorderStyle = bsFlat1Color Then
            ' to fake a 1-pixel border, we will draw the inner level of borders
            ' in the same backcolor as the control. Special flags are set for this.
            ' Note: GetBkColor does not return the control's BackColor property.
            ' These controls are sent a wm_ctlcoloredit or wm_ctlcolorlistbox message
            ' forcing the parent object to provide the bacground color brush
            Settings(Index + 3) = [_bsBackColor]
            Settings(Index + 4) = [_bsBackColor] - 1
            ' normally a 1 pixel border only requires 2 colors and only an outer border
            ' however, we'll tweak so 4 colors are used & also an inner border
            Settings(Index) = CLng(bsSunken Or (ctrlType * &H100&))
            If BorderStyle = bsFlat1Color Then Settings(Index + 2) = Settings(Index + 1)
        End If
    Else
        ' combo styles are not modified, but color tweaked above to appear so
        Select Case BorderStyle
        Case bsFlat1Color, bsFlat2Color
            If ctrlType = ctTextBox Then
                ' another exception: textbox with flat border style draws border on client
                If (lExStyle And WS_EX_CLIENTEDGE) = 0 Then ' flat?
                    lStyle = lStyle And Not WS_BORDER
                    bFlatTextBox = True
                Else
                    lStyle = lStyle Or WS_BORDER
                End If
            Else
                lStyle = lStyle Or WS_BORDER
            End If
            lExStyle = lExStyle And Not WS_EX_CLIENTEDGE
        Case Else
            If ctrlType = ctTextBox Then
                ' another exception: textbox with flat border style draws border on client
                If (Settings(Index + 6) And WS_EX_CLIENTEDGE) = 0 Then ' flat?
                    ' we need 2 borders, one drawn on client & we provide the 2nd one via WS_Border
                    lStyle = lStyle Or WS_BORDER
                    lExStyle = lExStyle And Not WS_EX_CLIENTEDGE
                    bFlatTextBox = True
                Else
                    lStyle = lStyle And Not WS_BORDER
                    lExStyle = lExStyle Or WS_EX_CLIENTEDGE
                End If
            Else
                lStyle = lStyle And Not WS_BORDER
                lExStyle = lExStyle Or WS_EX_CLIENTEDGE
            End If
        End Select
        
        ' now should a user be setting a listbox (IntegralHeight=True) border to
        ' raised/sunken when it was previously set to flat, simply changing the borderstyles
        ' will result in the listbox shrinking in size. We should account for that
        
        ' get current bordersize
        lOldStyle = GetWindowLong(hwnd, GWL_STYLE)
        ' get current bordersize
        If (lOldStyle And WS_BORDER) = WS_BORDER Then
            borderSize = GetSystemMetrics(SM_CYBORDER)
            bHasBorder = True
        End If
        If (GetWindowLong(hwnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then
            borderSize = borderSize + GetSystemMetrics(SM_CYEDGE)
        Else
            bHasBorder = False
        End If
        GetWindowRect hwnd, wRect
        cRect = wRect
        OffsetRect wRect, -wRect.Left, -wRect.Top

        If ctrlType = ctListBox Then
            bIntegralHT = ((lOldStyle And LBS_NOINTEGRALHEIGHT) = 0)
            
            ' get adjusted bordersize
            If (lStyle And WS_BORDER) = WS_BORDER Then borderSizeNew = GetSystemMetrics(SM_CYBORDER)
            If (lExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then borderSizeNew = borderSizeNew + GetSystemMetrics(SM_CYEDGE)
            If bIntegralHT = True Then
                If borderSizeNew > borderSize Then
                    ' adjust height as needed
                    wRect.Bottom = wRect.Bottom + (borderSizeNew - borderSize) * 2
                    lFlags = lFlags And Not SWP_NOSIZE
                End If
            End If
        End If
        If bHasBorder Then
            ' some controls (FileList, Treeview, ListView) have an extra border.
            ' It is same color as the form's back color an is the far outer border.
            ' We will be removing that but to ensure the control looks as though it
            ' stays at the same X,Y coordinates after we remove it, we will shift
            ' the control 1 pixel
            ScreenToClient GetParent(hwnd), cRect
            OffsetRect cRect, 1, 1
            lFlags = lFlags And Not SWP_NOMOVE
        Else
            cRect.Left = 0: cRect.Top = 0
        End If
        SetWindowLong hwnd, GWL_EXSTYLE, lExStyle
        SetWindowLong hwnd, GWL_STYLE, lStyle
    End If
    
    If Not isSubclassed Then    ' haven't subclassed this hWnd yet
        sc_Subclass hwnd
        If ctrlType = ctImageCombo Then
            sc_AddMsg hwnd, WM_DRAWITEM, MSG_AFTER ' ownerdrawn control
            sc_AddMsg hwnd, WM_CTLCOLOREDIT, MSG_BEFORE
        ElseIf ctrlType = ctComboBox Or bFlatTextBox = True Then
            ' draw border on client
            sc_AddMsg hwnd, WM_PAINT, MSG_AFTER
            sc_AddMsg hwnd, WM_CTLCOLOREDIT, MSG_BEFORE
            sc_AddMsg hwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
        Else
            sc_AddMsg hwnd, WM_NCPAINT, MSG_AFTER ' draw border on non-client
        End If
    End If
    
    ' force any border changes now
    SetWindowPos hwnd, 0, cRect.Left, cRect.Top, wRect.Right, wRect.Bottom, lFlags
        
    ' force repaint. SetWindowPos seems not to do this for all controls
    InvalidateRect hwnd, ByVal 0&, True
    
End Sub

Public Sub ReSetBorder(ByVal hwnd As Long, Optional Refresh As Boolean = True)

    ' function removes an custom borders & resets to system defaults
    ' For extra info on comments, see SetBorder. This is basically the
    ' opposite logic that was applied in SetBorder
    
    If Borders Is Nothing Then Exit Sub
    
    Dim Index As Long
    Dim wRect As RECT, cRect As RECT
    Dim lStyle As Long, lStyleEX As Long
    Dim borderSize As Long, borderSizeNew As Long
    Dim lFlags As Long
    Dim ctrlType As Long
    Dim bIntegralHT As Boolean, bHasBorder As Boolean
    
    On Error Resume Next
    Index = Borders("h" & hwnd) ' did we subclass this one?
    If Index = 0 Then
        If Err Then Err.Clear   ' nope, nothing to do
        Exit Sub
    End If
    On Error GoTo 0

    sc_UnSubclass hwnd  ' unsubclass it
    ' these will be the flags used for SetWindowPos
    lFlags = SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
    
    ctrlType = ((Settings(Index) \ &HFF) And &HFF)
    
    ' comboboxes go through un-modified, don't need to run extra code for those
    If Not (ctrlType = ctComboBox Or ctrlType = ctImageCombo) Then
        ' listboxes and comboboxes could shrink in height if there IntegralHeight
        ' property is set to True. This is because leaving the control the same
        ' size but adding extra border pixels to it would reduce the total client
        ' area and thus force the control to resize itself smaller. Account for this.
        ' Additionally, other controls like FileList, Treeview, ListView had
        ' and extra border level we compensated for by shifting its X,Y position.
        ' Re-shift it if needed.
        
        GetWindowRect hwnd, wRect
        cRect = wRect
        OffsetRect wRect, -wRect.Left, -wRect.Top
    
        ' calculate the original bordersize
        If (Settings(Index + 5) And WS_BORDER) = WS_BORDER Then
            borderSizeNew = GetSystemMetrics(SM_CYBORDER)
            bHasBorder = True
        End If
        If (Settings(Index + 6) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then
            borderSizeNew = borderSizeNew + GetSystemMetrics(SM_CYEDGE)
        Else
            bHasBorder = False
        End If
    
        If ctrlType = ctListBox Then
            bIntegralHT = ((Settings(Index + 5) And LBS_NOINTEGRALHEIGHT) = 0)
            
            ' calculate the current vertical bordersize
            lStyle = GetWindowLong(hwnd, GWL_STYLE)
            lStyleEX = GetWindowLong(hwnd, GWL_EXSTYLE)
            If (lStyle And WS_BORDER) = WS_BORDER Then borderSize = GetSystemMetrics(SM_CYBORDER)
            If (lStyleEX And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then borderSize = borderSize + GetSystemMetrics(SM_CYEDGE)
            
            ' now calculate any difference in height
            If bIntegralHT Then ' our control doesn't have IntegralHeight=False
                If borderSizeNew > borderSize Then
                    ' need to resize the control when we re-apply the original borders
                    ' add the extra height needed so the client area can remain same size
                    wRect.Bottom = wRect.Bottom + (borderSizeNew - borderSize) * 2
                    lFlags = lFlags And Not SWP_NOSIZE ' we want to resize the control
                End If
            End If
        End If
        If bHasBorder Then
            ScreenToClient GetParent(hwnd), cRect
            OffsetRect cRect, -1, -1
            lFlags = lFlags And Not SWP_NOMOVE
        Else
            cRect.Left = 0: cRect.Top = 0
        End If
    
    End If
    ' set the new border styles as needed, then force a resize/redraw
    SetWindowLong hwnd, GWL_STYLE, Settings(Index + 5)
    SetWindowLong hwnd, GWL_EXSTYLE, Settings(Index + 6)
    SetWindowPos hwnd, 0, cRect.Left, cRect.Top, wRect.Right, wRect.Bottom, lFlags

    ' now clean up our collection
    If Borders.Count = 1 Then
        Set Borders = Nothing
        Erase Settings
    Else
        If Index + 8 < UBound(Settings) Then
            ' move the last item in our collection over the unsubclassed item
            lFlags = Borders.Item(Borders.Count)
            RtlMoveMemory VarPtr(Settings(Index)), VarPtr(Settings(lFlags)), 32&
            ' now rebuild the collection item
            Borders.Remove "h" & Settings(Index + 7)
            Borders.Add Index, "h" & Settings(Index + 7), "h" & hwnd
        End If
        Borders.Remove "h" & hwnd
        ReDim Preserve Settings(1 To UBound(Settings) - 8)
    End If

    ' force repaint. SetWindowPos seems not to do this for all controls
    If Refresh Then InvalidateRect hwnd, ByVal 0&, True

End Sub

Private Sub RGBToHSL(inColor As Long, Hue As Single, _
                Luminance As Single, Saturation As Single)

' various versions on the net. This one needed some tweaking but will accurately
' return the HSL +/- 2 on a single RGB value compared to the color picker.

' The HSL values are ratios btwn 0:100 & can be applied to any maximum value with a
' minimum value of zero. For example: Luminance is a ratio btwn 0:100 and can be
' multiplied against 240 to return a minimum/maximum Luminance of 0 to 240 similar
' to how the color picker displays it.  It can also be applied against 256, the
' maximum number of shades of light.
    Dim maxRGB As Single, minRGB As Single, hueDelta As Single
    Dim Red As Single, Green As Single, Blue As Single
    Dim bRGB(0 To 3) As Byte
    
    ' get long color into RGB bytes & convert to a ratio
    bRGB(0) = inColor Mod &H100
    bRGB(1) = (inColor \ &H100) Mod &H100
    bRGB(2) = (inColor \ &H10000) Mod &H100
    
    Red = bRGB(0) / 255
    Green = bRGB(1) / 255
    Blue = bRGB(2) / 255
    
    ' simple little routine to get the largest/smallest of the RGB bytes
    If Red > Green Then
        minRGB = Green
        maxRGB = Red
    Else
        minRGB = Red
        maxRGB = Green
    End If
    If Blue > maxRGB Then
        maxRGB = Blue
    Else
        If Blue < minRGB Then minRGB = Blue
    End If

    ' calculate Luminance or Light as some know it
    Luminance = (maxRGB + minRGB) / 2

    ' calculate Saturation
    If maxRGB = minRGB Then
        ' grayscale, only Luminance is used
        Saturation = 0#
        Hue = 0#
    Else
        If Luminance <= 0.5 Then
           Saturation = (maxRGB - minRGB) / (maxRGB + minRGB)
        Else
           Saturation = (maxRGB - minRGB) / (2 - maxRGB - minRGB)
        End If
        ' now for the Hue; much more complicated
        
        ' get the difference between the largest/smallest color range
        hueDelta = maxRGB - minRGB
        
        ' this is a bit complicated, but Hue is calculated here as a relation
        ' to a 360 degree circle.
        
        ' The IF's below determine between what 2 of the 3 primary colors this color falls
        ' in order to calculate the Hue
        If Red = maxRGB Then
           ' falls betweeen green & blue
            Hue = (Green - Blue) / hueDelta
        ElseIf Green = maxRGB Then
           ' falls between blue and red
            Hue = 2# + (Blue - Red) / hueDelta
        Else    'Blue = maxRGB
           ' falls between red & green
            Hue = 4# + (Red - Green) / hueDelta
        End If
        ' convert to degrees
        Hue = Hue * 60
        ' check for reverse rotation & adjust by adding 360
        If Hue < 0# Then Hue = Hue + 360#
        ' now we can convert Hue & Saturation to a simple ratio before returning
        Hue = Hue / 360# * 100#
        Saturation = Saturation * 100#
    End If
    ' convert to proper ratio before returning
    Luminance = Luminance * 100#
    ' done: return result.
End Sub

Private Function HSLToRGB(ByVal Hue As Single, _
        ByVal Luminance As Single, ByVal Saturation As Single) As Long

' various versions on the net. This one needed some tweaking but will accurately
' return the RGB +/- 2 on a single HSL value when compared to the color picker.

' When the parameters are passed as percentages calculated from the RGBtoHSL
' routine, this routine will convert HSL to RGB with 100% accuracy.

' function will return the RGB as a long:
' 1st byte=Red, 2nd byte=Blue , 3rd byte=Green, 4th byte is unused

    Dim Red As Single, Green As Single, Blue As Single
    Dim minRGB As Single, maxRGB As Single
    Dim bRGB(0 To 3) As Byte
    
    ' ensure passed parameters are within 0:100 range
        If Saturation < 0 Then
            Saturation = 0
        ElseIf Saturation > 100 Then
            Saturation = 100
        End If
        If Luminance < 0 Then
            Luminance = 0
        ElseIf Luminance > 100 Then
            Luminance = 100
        End If
        If Hue < 0 Then
            Hue = 0
        ElseIf Hue > 100 Then
            Hue = 100
        End If
    ' now convert ratios to a 0:1 range
    Saturation = Saturation / 100
    Luminance = Luminance / 100

    If Saturation = 0 Then
       ' grayscale
       Red = Luminance
       Green = Luminance
       Blue = Luminance
    
    Else
        Hue = Hue / 100
        ' calculate the least RGB value
        If Luminance <= 0.5 Then
            minRGB = Luminance * (1 - Saturation)
        Else
            minRGB = Luminance - Saturation * (1 - Luminance)
        End If
        ' calculate the greatest RGB value
        maxRGB = 2 * Luminance - minRGB
          
        ' Hue is related to a circle or quadrants in this case
        ' We will use ratio on 360 degree to determine reverse rotation
        Hue = Hue * 360
        ' Any angle calculated > 299.999 is a reverse rotation and the ratio
        ' is actually a negative number. We need to get this negative value.
        ' Note: This section checked by running all possible RGB combinations
        ' thru the code to ensure RGB > HSL > RGB returned original color.
        If Hue >= 300 Then Hue = Hue - 360
        ' now to get the final 0:1 ratio...
        Hue = Hue / 60
        
        ' this is the reverse of the RGBtoHSL routine & is where the magic happens
        If (Hue < 1) Then
           Red = maxRGB
           If (Hue < 0) Then
              Green = minRGB
              Blue = Green - Hue * (maxRGB - minRGB)
           Else
              Blue = minRGB
              Green = Hue * (maxRGB - minRGB) + Blue
           End If
        ElseIf (Hue < 3) Then
           Green = maxRGB
           If (Hue < 2) Then
              Blue = minRGB
              Red = Blue - (Hue - 2) * (maxRGB - minRGB)
           Else
              Red = minRGB
              Blue = (Hue - 2) * (maxRGB - minRGB) + Red
           End If
        Else
           Blue = maxRGB
           If (Hue < 4) Then
              Red = minRGB
              Green = Red - (Hue - 4) * (maxRGB - minRGB)
           Else
              Green = minRGB
              Red = (Hue - 4) * (maxRGB - minRGB) + Green
           End If
    
        End If
             
    End If
    
    ' convert the ratios to RGB & return the result
    bRGB(0) = CByte(Red * 255): bRGB(1) = Green * 255: bRGB(2) = Blue * 255
    HSLToRGB = RGB(bRGB(0), bRGB(1), bRGB(2))

End Function

Private Function ConvertColor(tColor As Long) As Long

' Converts VB color constants to real color values
    If tColor < 0 Then
        ConvertColor = GetSysColor(tColor And &HFF&)
    Else
        ConvertColor = tColor
    End If
End Function


Private Sub Class_Terminate()
  'Terminate all subclassing
  
  Dim I As Integer
  If Not Borders Is Nothing Then
    For I = Borders.Count To 1 Step -1
        ReSetBorder Settings(Borders.Item(I) + 7), False
    Next
  End If
  sc_Terminate
  Set Borders = Nothing
  Erase Settings
End Sub

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long) As Boolean             'Subclass the specified window handle
  Dim nAddr As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError "sc_Subclass", "Invalid window handle"
  End If

  If z_Funk Is Nothing Then                                                 'If first time
    nAddr = zAddressOf(1)                                                   'Get the address of the final private method
    
    If nAddr = 0 Then
      zError "sc_Subclass", "Callback method not found"
    End If
    
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-memory-address collection
    
    z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
    
    z_Base = VarPtr(z_Code(0))                                              'Map the address of z_Code()'s first element to the zData() array
    zData(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                           'Store the EbMode function address in the thunk data
    zData(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                   'Store CallWindowProc function address in the thunk data
    zData(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                    'Store the SetWindowLong function address in the thunk data
    zData(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                    'Store the VirtualFree function address in the thunk data
    zData(IDX_ME) = ObjPtr(Me)                                              'Store my object address in the thunk data
    zData(IDX_CALLBACK) = nAddr                                             'Store the zWndProc address in the thunk data
  End If
  
  z_Base = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                   'Allocate executable memory
  RtlMoveMemory z_Base, VarPtr(z_Code(0)), CODE_LEN                         'Copy the thunk to the allocated memory

  On Error GoTo Catch                                                       'Catch double subclassing
    z_Funk.Add z_Base, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
  On Error GoTo 0

  zData(IDX_EBX) = z_Base                                                   'Patch the data address
  zData(IDX_HWND) = lng_hWnd                                                'Store the window handle in the thunk data
  zData(IDX_BTABLE) = z_Base + CODE_LEN                                     'Store the address of the before table in the thunk data
  zData(IDX_ATABLE) = zData(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)           'Store the address of the after table in the thunk data
  zData(IDX_WNDPROC) = _
                SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_Base + WNDPROC_OFF) 'Set the new WndProc and store the original WndProc in the thunk data
  sc_Subclass = True                                                        'Indicate success
  Exit Function                                                             'Exit

Catch:
  zError "sc_Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim I     As Long
  Dim nAddr As Long

  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubAll", "Subclassing hasn't been started", False
  Else
    With z_Funk
      For I = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        nAddr = .Item(I)                                                    'Map zData() to the hWnd thunk address
        If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already freed itself
          z_Base = nAddr                                                    'Map the thunk memory to the zData() array
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next I                                                                'Next member of the collection
    End With
    
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Public Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "UnSubclass", "Subclassing hasn't been started", False
  Else
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                              'Delete all before messages
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                              'Delete all after messages
    z_Base = zMap_hWnd(lng_hWnd)                                            'Map the thunk memory to the zData() array
    zData(IDX_SHUTDOWN) = -1                                                'Set the shutdown indicator
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
  End If

  If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
  End If

  If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
  sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim I      As Long                                                        'Loop index

  z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
  z_Base = zData(nTable)                                                    'Map zData() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
      Exit Sub
    End If

    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = 0 Then                                                  'If the element is free...
        zData(I) = uMsg                                                     'Use this element
        Exit Sub                                                            'Bail
      ElseIf zData(I) = uMsg Then                                           'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next I                                                                  'Next message table entry

    nCount = I                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim I      As Long                                                        'Loop index

  z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
  z_Base = zData(nTable)                                                    'Map zData() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = uMsg Then                                               'If the message is found...
        zData(I) = 0                                                        'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next I                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  
  MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  If bEnd Then
    End
  End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "z_Base = zMap_hWnd", "Subclassing hasn't been started", True
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    zMap_hWnd = z_Funk("h" & lng_hWnd)                                      'Get the thunk address
    z_Base = zMap_hWnd                                                      'Map zData() to the thunk address
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "z_Base = zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal private method, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(Optional ByVal nOrdinal As Long = 1) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  If z_TblEnd = 0 Then                                                      'First time through...
    RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                              'Get the address of this object instance
  
    If Not zProbe(nAddr + &H1C, I, bSub) Then                               'Probe for a Class method
      If Not zProbe(nAddr + &H6F8, I, bSub) Then                            'Probe for a Form method
        If Not zProbe(nAddr + &H7A4, I, bSub) Then                          'Probe for a UserControl method
          Exit Function                                                     'Bail...
        End If
      End If
    End If
    
    I = I + 4                                                               'Bump to the next entry
    J = I + 1024                                                            'Set a reasonable limit, scan 256 vTable entries
    
    Do While I < J
      RtlMoveMemory VarPtr(nAddr), I, 4                                     'Get the address stored in this vTable entry
      
      If IsBadCodePtr(nAddr) Then                                           'Is the entry an invalid code address?
        z_TblEnd = I                                                        'Cache the vTable end-point
        GoTo Found                                                          'Bad method signature, quit loop
      End If
  
      RtlMoveMemory VarPtr(bVal), nAddr, 1                                  'Get the byte pointed to by the vTable entry
      If bVal <> bSub Then                                                  'If the byte doesn't match the expected value...
        z_TblEnd = I                                                        'Cache the vTable end-point
        GoTo Found                                                          'Bad method signature, quit loop
      End If
      
      I = I + 4                                                             'Next vTable entry
    Loop
    
    Exit Function                                                           'Final method not found
  End If
  
Found:
  RtlMoveMemory VarPtr(zAddressOf), z_TblEnd - (nOrdinal * 4), 4            'Return the specified vTable entry address
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_Base + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_Base + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback: must be private and the last method in the source file-----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'*************************************************************************************************
'* bBefore  - Indicates whether the callback is before or after the original WndProc. Usually you
'*            will know unless the callback for the uMsg value is specified as MSG_BEFORE_AFTER
'*            (both before and after the original WndProc).
'* bHandled - In a before original WndProc callback, setting bHandled to True will prevent the
'*            message being passed to the original WndProc and (if set to do so) the after
'*            original WndProc callback.
'* lReturn  - WndProc return value. Set as per the MSDN documentation for the message value,
'*            and/or, in an after the original WndProc callback, act on the return value as set
'*            by the original WndProc.
'* hWnd     - Window handle.
'* uMsg     - Message value.
'* wParam   - Message related data.
'* lParam   - Message related data.
'*************************************************************************************************
  
    Dim hBrushPen As Long
    Dim wRect As RECT
    Dim hDC As Long, cHwnd As Long, cHDC As Long
    Dim HLS(0 To 3) As Single
    Dim Color As Long
    Dim colIndex As Long
    Dim I As Integer, nrSteps As Integer
    
    colIndex = Borders("h" & hwnd)  ' index into the Settings array
    GetWindowRect hwnd, wRect       ' get window position/size
    OffsetRect wRect, -wRect.Left, -wRect.Top ' offset to 0,0
    
    Select Case uMsg
    
    Case WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX
        ' sent by combo boxes being subclassed
        Dim lgBrush As LOGBRUSH
        If Settings(colIndex + 4) = [_bsBackColor] - 1 Then
            ' 1-pixel border, but 2nd level border is drawn by VB
            ' We will overdraw it with back color. These controls send
            ' parent a request for a colored brush to use for background
            ' Intercept the brush and get its color
            hBrushPen = sc_CallOrigWndProc(hwnd, uMsg, wParam, lParam)
            GetGDIObject hBrushPen, Len(lgBrush), lgBrush
            Settings(colIndex + 3) = lgBrush.lbColor
            ' now draw the 2nd level border
            hDC = GetWindowDC(hwnd)
            InflateRect wRect, -1, -1
            FrameRect hDC, wRect, hBrushPen
            bHandled = True
            lReturn = hBrushPen
        End If
    
    Case Else ' WM_DrawItem, WM_Pain, WM_NCPaint depending on type of control subclassed
        If uMsg = WM_DRAWITEM Then hwnd = FindWindowEx(hwnd, 0, vbNullString, vbNullString)
        '^^ image combo borders are custom drawn on child, get that child window
        hDC = GetWindowDC(hwnd)
        
        If (Settings(colIndex) And &HFF) = bsFlat1Color Then
            ' single color, 1 pixel borders (non-combos). Simply use FrameRect
            If Settings(colIndex + 1) = bsSysDefault Then
                Color = vbRed 'GetSysColor(COLOR_WINDOWFRAME)
            ElseIf Settings(colIndex + 1) = bsAutoShade Then
                Color = vbRed 'GetSysColor(COLOR_WINDOWFRAME)
            Else
                Color = vbRed 'ConvertColor(Settings(colIndex + 1))
            End If
            hBrushPen = CreateSolidBrush(Color)
            FrameRect hDC, wRect, hBrushPen
            DeleteObject hBrushPen
        Else
            ' determine number of passes. 2 for 1-pixel border, 4 for 2-pixels
            If (Settings(colIndex) And &HFF) = bsFlat2Color Then nrSteps = 2 Else nrSteps = 4
            Select Case Settings(colIndex + 1)
                ' get base color (Shadow) and convert it to HSL
                Case bsAutoShade
                    Color = HSLToRGB(HLS(0), 100, HLS(2))
                Case bsSysDefault, [_bsBackColor]
                    Color = GetSysColor(COLOR_BTNSHADOW)
                Case Else
                    Color = ConvertColor(Settings(colIndex + 1))
            End Select
            RGBToHSL Color, HLS(0), HLS(1), HLS(2)
            
            For I = 1 To nrSteps
                ' select appropriate color, create pen & draw
                If (Settings(colIndex) And &HFF) = bsRaised Then
                    ' raised borders use different color scheme
                    Select Case I
                    Case 1
                        Select Case Settings(colIndex + 2)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), 100, HLS(2))
                        Case bsSysDefault, [_bsBackColor]
                            Color = GetSysColor(COLOR_BTNHIGHLIGHT)
                        Case Else
                            Color = ConvertColor(Settings(colIndex + 2))
                        End Select
                    Case 2
                        Select Case Settings(colIndex + 3)
                        Case bsAutoShade
                            If HLS(1) < 20 Then
                                Color = vbBlack
                            Else
                                Color = HSLToRGB(HLS(0), 15, HLS(2))
                            End If
                        Case bsSysDefault
                            Color = GetSysColor(COLOR_3DDKSHADOW)
                        Case [_bsBackColor]
                            Color = GetBkColor(hDC)
                        Case Else
                            Color = ConvertColor(Settings(colIndex + 3))
                        End Select
                    Case 3
                        Select Case Settings(colIndex + 4)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), 85, HLS(2))
                        Case bsSysDefault
                            Color = GetSysColor(COLOR_3DLIGHT)
                        Case [_bsBackColor]
                            Color = GetBkColor(hDC)
                        Case [_bsBackColor] - 1
                            If Settings(colIndex + 3) = [_bsBackColor] Then
                                Color = GetBkColor(hDC)
                            Else
                                Color = Settings(colIndex + 3)
                            End If
                        Case Else
                            Color = ConvertColor(Settings(colIndex + 4))
                        End Select
                    Case 4
                        Select Case Settings(colIndex + 1)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), 100, HLS(2))
                        Case bsSysDefault, [_bsBackColor]
                            Color = GetSysColor(COLOR_BTNSHADOW)
                        Case Else
                            Color = ConvertColor(Settings(colIndex + 1))
                        End Select
                    End Select
                Else                ' sunken borders
                    Select Case I
                    Case 1
                    Case 2
                        Select Case Settings(colIndex + I)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), 100, HLS(2))
                        Case bsSysDefault, [_bsBackColor]
                            Color = GetSysColor(COLOR_BTNHIGHLIGHT)
                        Case Else
                            Color = ConvertColor(Settings(colIndex + I))
                        End Select
                    Case 3
                        Select Case Settings(colIndex + I)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), HLS(1) - 15, HLS(2))
                        Case bsSysDefault
                            Color = GetSysColor(COLOR_3DDKSHADOW)
                        Case [_bsBackColor]
                            Color = GetBkColor(hDC)
                        Case Else
                            Color = ConvertColor(Settings(colIndex + I))
                        End Select
                    Case 4
                        Select Case Settings(colIndex + I)
                        Case bsAutoShade
                            Color = HSLToRGB(HLS(0), 85, HLS(2))
                        Case bsSysDefault
                            Color = GetSysColor(COLOR_3DLIGHT)
                        Case [_bsBackColor]
                            Color = GetBkColor(hDC)
                        Case [_bsBackColor] - 1
                        Case Else
                            Color = ConvertColor(Settings(colIndex + I))
                        End Select
                    End Select
                End If
                
                hBrushPen = SelectObject(hDC, CreatePen(0, 1, Color))
                Select Case I
                Case 1
                    MoveToEx hDC, 0, wRect.Bottom - 2, ByVal 0&
                    LineTo hDC, 0, 0
                    LineTo hDC, wRect.Right - 1, 0
                Case 2
                    LineTo hDC, wRect.Right - 1, 0
                    LineTo hDC, wRect.Right - 1, wRect.Bottom - 1
                    LineTo hDC, -1, wRect.Bottom - 1
                Case 3
                    MoveToEx hDC, 1, wRect.Bottom - 3, ByVal 0&
                    LineTo hDC, 1, 1
                    LineTo hDC, wRect.Right - 2, 1
                Case 4
                    LineTo hDC, wRect.Right - 2, 1
                    LineTo hDC, wRect.Right - 2, wRect.Bottom - 2
                    LineTo hDC, 0, wRect.Bottom - 2
                End Select
                DeleteObject SelectObject(hDC, hBrushPen)
            Next
        End If
    End Select
    ' release DC as needed
    If Not hDC = 0 Then ReleaseDC hwnd, hDC
End Sub


