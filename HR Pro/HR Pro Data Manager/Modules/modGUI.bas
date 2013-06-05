Attribute VB_Name = "modGui"
Option Explicit

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





