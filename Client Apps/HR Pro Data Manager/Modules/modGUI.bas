Attribute VB_Name = "modGui"
Option Explicit

'Font enumeration types
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
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
  ntmFlags As Long
  ntmSizeEM As Long
  ntmCellHeight As Long
  ntmAveWidth As Long
End Type

' ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

' tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1

Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

' EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Icon extracting functions
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal Y&, ByVal Flags&) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const ILD_TRANSPARENT = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const LB_SETHORIZONTALEXTENT = &H194

' XP style controls
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()



Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE As Long = (-16)

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
   ByVal FontType As Long, lParam As ListBox) As Long

  Dim FaceName As String
  Dim FullName As String
    
  FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
  EnumFontFamProc = 1

End Function

Public Sub FillComboWithFonts(objControl As ComboBox)

  Dim hDC As Long
  
  objControl.Clear
  hDC = GetDC(objControl.hWnd)
  EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, objControl
  ReleaseDC objControl.hWnd, hDC

End Sub

' Textwidth function causes overflow on larger pieces of data. This wrapper should handle it.
Public Function BigTextWidth(ByRef sInString As Variant, ByVal MaximumSize As Single) As Long
  
  Dim lngTextWidth As Single
  
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
  
End Function





