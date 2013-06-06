Attribute VB_Name = "modWindow"
Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long

Private Declare Function DefWindowProc Lib "user32" Alias _
   "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Private m_WinSizeInfoCol As New CWinSizeInfos

Public Sub Hook(lhWnd As Long, MinX As Long, MinY As Long, Optional MaxX As Long, Optional MaxY As Long)

  'Start subclassing.... if not in development environment
  'otherwise trying to debug can cause IDE crashes
  If Not ASRDEVELOPMENT Then
    'Unhook in case it already exists in the collection
    Unhook lhWnd
    
    lpPrevWndProc = SetWindowLong(lhWnd, GWL_WNDPROC, _
      AddressOf WindowProc)
      
    If MaxX = 0 Then MaxX = Screen.Width
    If MaxY = 0 Then MaxY = Screen.Height
      
    Call m_WinSizeInfoCol.Add(lhWnd, MinX, MinY, MaxX, MaxY)
  End If
End Sub

Public Sub Unhook(lhWnd As Long)
  Dim temp As Long
  
  If m_WinSizeInfoCol.Exists(CStr(lhWnd)) Then
    'Cease subclassing.
    temp = SetWindowLong(lhWnd, GWL_WNDPROC, lpPrevWndProc)
    Call m_WinSizeInfoCol.Remove(CStr(lhWnd))
  End If
End Sub

Private Function WindowProc(ByVal lhWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
  Dim MinMax As MINMAXINFO

  'Check for request for min/max window sizes.
  If uMsg = WM_GETMINMAXINFO Then
    'Retrieve default MinMax settings
    CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
   
    Dim sizeInfo As CWinSizeInfo
    Set sizeInfo = m_WinSizeInfoCol.Item(CStr(lhWnd))
    
    'Specify new minimum size for window.
    MinMax.ptMinTrackSize.X = sizeInfo.MinX
    MinMax.ptMinTrackSize.Y = sizeInfo.MinY

    'Specify new maximum size for window.
    MinMax.ptMaxTrackSize.X = sizeInfo.MaxX
    MinMax.ptMaxTrackSize.Y = sizeInfo.MaxY
    
    'Copy local structure back.
    CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

    WindowProc = DefWindowProc(lhWnd, uMsg, wParam, lParam)
  Else
    WindowProc = CallWindowProc(lpPrevWndProc, lhWnd, uMsg, _
      wParam, lParam)
  End If
End Function

Public Function PixelsToTwips(pixels As Long) As Long
  PixelsToTwips = pixels * Screen.TwipsPerPixelX
End Function

Public Function TwipsToPixels(twips As Long) As Long
  TwipsToPixels = twips / Screen.TwipsPerPixelX
End Function

' Removes the caption bar icon from a form
Public Sub RemoveIcon(ByRef pobjForm As Form)

  pobjForm.Icon = Nothing
  SetWindowLong pobjForm.hWnd, GWL_EXSTYLE, WS_EX_DLGMODALFRAME

End Sub

Public Sub SetBlankIcon(ByRef pobjForm As Form)
  pobjForm.Icon = LoadResPicture("BLANK", vbResIcon)
End Sub

Public Sub SetFormCaption(ByRef pobjForm As Form, ByVal strCaption As String)
  pobjForm.Caption = strCaption
  RemoveIcon pobjForm
End Sub

' Puts the COA colour scheme onto a activebar menu control
Public Sub ApplySkinToActiveBar(ByRef pobjActiveBar As ActiveBarLibraryCtl.ActiveBar)

  With pobjActiveBar
    .MenuFontStyle = 1
    .Font.Name = "Verdana"
    .Font.Bold = False
    .Font.Size = 8

    .ControlFont.Name = "Verdana"
    .ControlFont.Bold = False
    .ControlFont.Size = 8

    .ForeColor = 6697779
    .BackColor = 16248553

    .Refresh
  End With

End Sub


' Textwidth function causes overflow on larger pieces of data. This wrapper should handle it.
Public Function BigTextWidth(ByRef sInString As Variant, ByVal MaximumSize As Long) As Long
  
  On Error GoTo ErrorTrap
  
  Dim lngTextWidth As Long
   
  If Len(sInString) > 500 Then
    lngTextWidth = Printer.TextWidth(Left(sInString, 500)) + _
          BigTextWidth(Right(sInString, Len(sInString) - 500), 0)
  Else
    lngTextWidth = Printer.TextWidth(sInString)
  End If
  
  If MaximumSize > 0 Then
    BigTextWidth = Minimum(lngTextWidth, MaximumSize)
  Else
    BigTextWidth = lngTextWidth
  End If
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  BigTextWidth = Len(sInString) * 100
  GoTo TidyUpAndExit
 
End Function

' Returns the maximum of two values
Public Function Maximum(plngValue1 As Long, plngValue2 As Long) As Long
  Maximum = IIf(plngValue1 > plngValue2, plngValue1, plngValue2)
End Function

' Returns the minimum of two values
Public Function Minimum(plngValue1 As Long, plngValue2 As Long) As Long
  Minimum = IIf(plngValue1 < plngValue2, plngValue1, plngValue2)
End Function

