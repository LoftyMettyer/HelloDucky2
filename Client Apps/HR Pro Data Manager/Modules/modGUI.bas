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

' Textwidth function causes overflow on larger pieces of data. This wrapper should handle it.
Public Function BigTextWidth(ByRef sInString As Variant, ByVal MaximumSize As Single) As Long
  
  On Error GoTo ErrorTrap
  
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
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  BigTextWidth = Len(sInString) * 100
  GoTo TidyUpAndExit
 
  
End Function





