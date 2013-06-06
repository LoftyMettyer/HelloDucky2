VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function ConvertLocaleDateToSQL(psDate As String) As String

  Dim sLocaleFormat As String
  Dim sSQLFormat As String
  Dim iLocaleIndex As Integer
  Dim iIndex As Integer
  
  If Len(psDate) > 0 Then
    sLocaleFormat = LCase(DateFormat)
    
    iIndex = InStr(sLocaleFormat, "mm")
    If iIndex > 0 Then
      sSQLFormat = Mid(psDate, iIndex, 2) & "/"
    End If

    iIndex = InStr(sLocaleFormat, "dd")
    If iIndex > 0 Then
      sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 2) & "/"
    End If

    iIndex = InStr(sLocaleFormat, "yyyy")
    If iIndex > 0 Then
      sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 4)
    End If

    ConvertLocaleDateToSQL = sSQLFormat
  Else
    ConvertLocaleDateToSQL = ""
  End If
  
End Function
Public Function ValidateDate(psDate As String) As Variant
  ' Validate the given date with the system date format.
  ' Return vbNullString if the given date is invalid
  ' Else return the date value of the given date string.
  Dim sDay As String
  Dim sMonth As String
  Dim sYear As String
  Dim iYear As Integer
  Dim iDaysPerMonth As Integer
  Dim sDateFormat As String
  
  ' Get the system date format.
  sDateFormat = DateFormat
  
  ' Extract the year, month and day parts of the date.
  sYear = StrReplace(Mid(psDate, InStr(1, sDateFormat, "yyyy"), 4), " ", "")
  sMonth = StrReplace(Mid(psDate, InStr(1, sDateFormat, "MM"), 2), " ", "")
  sDay = StrReplace(Mid(psDate, InStr(1, sDateFormat, "dd"), 2), " ", "")

  ' Do not allow any part of the date to be empty.
  If Trim(sYear) = "" Or _
    Trim(sMonth) = "" Or _
    Trim(sDay) = "" Then
    
    ValidateDate = vbNullString
    Exit Function
    
  End If
  
  ' If the user entered less than two digits or less then assume
  ' the current century.
  If Len(Trim(sYear)) <= 2 Then
    ' Get the current century.
    sYear = Str((100 * Int(Year(Now) / 100)) + Val(sYear))
  End If
  
  ' Pad the year with zeroes to make it up to four digits long.
  sYear = Format(sYear, "0000")
  
  ' Validate the month part of the date.
  ' The date is invalid if the month part is empty while the other parts aren't.
  If Val(sMonth) > 12 Or _
    Val(sMonth) < 1 Then
    
    ValidateDate = vbNullString
    Exit Function
  End If
  
  ' Pad the month with zeroes to make it up to two digits long.
  sMonth = Format(sMonth, "00")
  
  ' Validate the day part of the date.
  ' Ensure we have a valid day value for the given month value.
  Select Case Val(sMonth)
    Case 4, 6, 9, 11 ' April, June, September, November (30 days in each)
      iDaysPerMonth = 30
      
    Case 1, 3, 5, 7, 8, 10, 12 'January, March, May, July, August, October, December (31 days in each)
      iDaysPerMonth = 31
    
    Case 2 ' February (28 days usually, 29 days in a leap year)
      iDaysPerMonth = 28
      iYear = Val(sYear)
      ' If the year is evenly divisible by 4 and not by 100
      ' then it is a leap year.
      If (iYear Mod 4 = 0) And _
        (iYear Mod 100 <> 0) Then
        iDaysPerMonth = 29
      Else
        ' If the year is evenly divisible by 4 and 100, then check to
        ' see if the quotient of year divided by 100 is also evenly
        ' divisible by 4. If it is, then this is a leap year.
        If (iYear Mod 4 = 0) And _
          (iYear Mod 100 = 0) And _
          (Int(iYear / 100) Mod 4 = 0) Then
          iDaysPerMonth = 29
        End If
      End If
      
  End Select

  If Val(sDay) < 1 Or Val(sDay) > iDaysPerMonth Then
    ValidateDate = vbNullString
    Exit Function
  End If
    
  ' Pad the day with zeroes to make it up to two digits long.
  sDay = Format(sDay, "00")
  
  ' Return the formatted date string.
  If InStr(1, sDateFormat, "d") < InStr(1, sDateFormat, "M") Then
    ValidateDate = DateValue(sDay & "/" & sMonth & "/" & sYear)
  Else
    ValidateDate = DateValue(sMonth & "/" & sDay & "/" & sYear)
  End If
  
End Function
Public Function DateFormat() As String
  ' Returns the date format.
  ' NB. Windows allows the user to configure totally stupid
  ' date formats (eg. d/M/yyMydy !). This function does not cater
  ' for such stupidity, and simply takes the first occurence of the
  ' 'd', 'M', 'y' characters.
  Dim sSysFormat As String
  Dim sSysDateSeparator As String
  Dim sDateFormat As String
  Dim iLoop As Integer
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  sDateFormat = ""
    
  sSysFormat = UI.GetSystemDateFormat
  sSysDateSeparator = UI.GetSystemDateSeparator
    
  ' Loop through the string picking out the required characters.
  For iLoop = 1 To Len(sSysFormat)
      
    Select Case Mid(sSysFormat, iLoop, 1)
      Case "d"
        If Not fDaysDone Then
          ' Ensure we have two day characters.
          sDateFormat = sDateFormat & "dd"
          fDaysDone = True
        End If
          
      Case "M"
        If Not fMonthsDone Then
          ' Ensure we have two month characters.
          sDateFormat = sDateFormat & "MM"
          fMonthsDone = True
        End If
          
      Case "y"
        If Not fYearsDone Then
          ' Ensure we have four year characters.
          sDateFormat = sDateFormat & "yyyy"
          fYearsDone = True
        End If
          
      Case Else
        sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
    End Select
      
  Next iLoop
    
  ' Ensure that all day, month and year parts of the date
  ' are present in the format.
  If Not fDaysDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "dd"
  End If
    
  If Not fMonthsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "MM"
  End If
    
  If Not fYearsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "yyyy"
  End If
    
  ' Return the date format.
  DateFormat = sDateFormat
  
End Function

Public Function ConvertSQLDateToLocale(psSQLDate As String) As String
  ' Convert the given date string (mm/dd/yyyy) into the locale format.
  ' NB. This function assumes a sensible locale format is used.
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  Dim sLocaleFormat As String
  Dim iLoop As Integer
  Dim sFormattedDate As String
  
  sFormattedDate = ""
  
  ' Get the locale's date format.
  sLocaleFormat = DateFormat
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  
  For iLoop = 1 To Len(sLocaleFormat)
    Select Case UCase(Mid(sLocaleFormat, iLoop, 1))
      Case "D"
        If Not fDaysDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 4, 2)
          fDaysDone = True
        End If
        
      Case "M"
        If Not fMonthsDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 1, 2)
          fMonthsDone = True
        End If
      
      Case "Y"
        If Not fYearsDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 7, 4)
          fYearsDone = True
        End If
      
      Case Else
        sFormattedDate = sFormattedDate & Mid(sLocaleFormat, iLoop, 1)
    End Select
  Next iLoop
  
  ConvertSQLDateToLocale = sFormattedDate
  
End Function



Public Function HiByte(ByVal w As Integer) As Byte
  If w And &H8000& Then
    HiByte = &H80 Or ((w And &H7FFF&) \ &HFF)
  Else
    HiByte = w \ &H100
  End If
End Function

Public Function HiWord(ByVal dw As Long) As Integer
  If dw And &H80000000 Then
    HiWord = (dw \ &HFFFF&) - 1
  Else
    HiWord = dw \ &HFFFF&
  End If
End Function

Public Function LoByte(ByVal w As Integer) As Byte
  LoByte = w And &HFF
End Function

Public Function LoWord(ByVal dw As Long) As Integer
  If dw And &H8000& Then
    LoWord = &H8000& Or (dw And &H7FFF&)
  Else
    LoWord = dw And &HFFFF&
  End If
End Function

Public Function LShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer
  Dim dw As Long
  
  dw = w * (2 ^ c)
  If dw And &H8000& Then
    LShiftWord = CInt(dw And &H7FFF&) Or &H8000
  Else
    LShiftWord = dw And &HFFFF&
  End If
End Function

Public Function RShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer
  Dim dw As Long
  If c = 0 Then
    RShiftWord = w
  Else
    dw = w And &HFFFF&
    dw = dw \ (2 ^ c)
    RShiftWord = dw And &HFFFF&
  End If
End Function

Public Function MakeWord(ByVal bHi As Byte, ByVal bLo As Byte) As Integer
  If bHi And &H80 Then
    MakeWord = (((bHi And &H7F) * &H100) + bLo) Or &H8000
  Else
    MakeWord = (bHi * &H100) + bLo
  End If
End Function

Public Function MakeDWord(ByVal wHi As Integer, ByVal wLo As Integer) As Long
  If wHi And &H8000& Then
    MakeDWord = (((wHi And &H7FFF&) * &H10000) Or (wLo And &HFFFF&)) Or &H8000&
  Else
    MakeDWord = (wHi * &H10000) + wLo
  End If
End Function

Public Function IsItemInCollection(Items As Object, ByVal ItemKey As String) As Boolean
  On Error GoTo ErrorTrap
  
  If Items(ItemKey).Key = ItemKey Then
    IsItemInCollection = True
  End If
  
  Exit Function
  
ErrorTrap:
  IsItemInCollection = False

End Function

Public Function StrReplace(ByVal Source As String, ByVal Search As String, ByVal Replace As String, Optional ByVal CaseSensitive As Boolean) As String
  Dim intStart As Integer, intEnd As String
  Dim strReturn As String
  
  strReturn = vbNullString
  intStart = 1
  intEnd = 0
  
  Do
    intEnd = InStr(intStart, Source, Search, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
    If intEnd > 0 Then
      strReturn = strReturn & Mid(Source, intStart, (intEnd - intStart)) & Replace
      intStart = intEnd + Len(Search)
    Else
      strReturn = strReturn & Mid(Source, intStart)
    End If
  Loop While intEnd > 0 And intStart <= Len(Source)
  
  StrReplace = strReturn
End Function

