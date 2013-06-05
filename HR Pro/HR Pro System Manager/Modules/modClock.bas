Attribute VB_Name = "modClock"
Option Explicit

  Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
  End Type
  
  Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

  Dim T As Long, liFrequency As LARGE_INTEGER, liStart As LARGE_INTEGER, liStop As LARGE_INTEGER
  Dim cuFrequency As Currency, cuStart As Currency, cuStop As Currency

  Dim iCount As Integer

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an ampty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

Public Sub StartClock()

  iCount = QueryPerformanceFrequency(liFrequency)
  cuFrequency = LargeIntToCurrency(liFrequency)

  QueryPerformanceCounter liStart
  cuStart = LargeIntToCurrency(liStart)

End Sub

Public Sub UpdateClock(ByVal Message As String)
  
  QueryPerformanceCounter liStop
  cuStop = LargeIntToCurrency(liStop)

  Debug.Print Message + " - " + CStr((cuStop - cuStart) / cuFrequency) + " seconds"

  QueryPerformanceCounter liStart
  cuStart = LargeIntToCurrency(liStart)

End Sub

