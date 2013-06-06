VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const mlngFileNum As Integer = 99
Const msFileName As String = "\Performance.log"

Private msEventBeingTracked As String
Private mdblLastTick As Double
Private mdblBaseTick As Double
Private mdblIgnoreTicks As Double
Private mdblPauseTick As Double
Private mbPaused As Boolean

Public Sub StartClock(ByRef sMessage As String)
  mdblLastTick = GetTickCount
  mdblBaseTick = mdblLastTick
  msEventBeingTracked = sMessage
End Sub

Public Sub PauseClock()
  mdblPauseTick = GetTickCount
  mbPaused = True
End Sub

Public Sub ContinueClock()

  Dim dblNowTick As Double
  
  dblNowTick = GetTickCount
  mdblIgnoreTicks = mdblIgnoreTicks + (dblNowTick - mdblPauseTick)
  mbPaused = False

End Sub

Public Sub ClearLogFile()
  On Error Resume Next
  Kill App.Path & msFileName
End Sub

Public Sub LogClockEvent(ByRef sArea As String, ByRef sMessage As String)
  
  On Local Error Resume Next

  Dim sOutput As String
  Dim dblTotalElapsed As Double
  Dim dblEventElapsed As Double
  Dim dblNow As Double

  Open App.Path & msFileName For Append As #mlngFileNum

  dblNow = GetTickCount
  dblEventElapsed = dblNow - mdblLastTick
  dblTotalElapsed = dblNow - mdblBaseTick

  sOutput = ClicksToNiceTime(dblTotalElapsed) & " - (" & dblEventElapsed & ") - " _
      & Now & " - " & sArea & " - " & sMessage

  Print #mlngFileNum, sOutput
  Close #mlngFileNum

  mdblLastTick = GetTickCount

End Sub

Private Function ClicksToNiceTime(ByRef dblClicks As Double) As String
  ClicksToNiceTime = Right(String(10, "0") & dblClicks, 10)
End Function

Public Sub LogSummary()

  On Local Error Resume Next
  
  Dim dblNow As Double
  Dim dblTotalElapsed As Double
  Dim sOutput As String
  
  dblNow = GetTickCount

  Open App.Path & msFileName For Append As #mlngFileNum

  If mbPaused Then
    dblTotalElapsed = (mdblPauseTick - mdblBaseTick) - mdblIgnoreTicks
  Else
    dblTotalElapsed = (dblNow - mdblBaseTick) - mdblIgnoreTicks
  End If
  sOutput = "Summary - " & ClicksToNiceTime(dblTotalElapsed) & " - " & msEventBeingTracked

  Print #mlngFileNum, sOutput
  Close #mlngFileNum

End Sub
