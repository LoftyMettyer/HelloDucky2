VERSION 5.00
Begin VB.UserControl COA_CalRepDates 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   690
   ScaleWidth      =   6990
   Begin VB.Line VerticalSeparator 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   5520
      X2              =   5520
      Y1              =   525
      Y2              =   360
   End
   Begin VB.Line HorizontalSeparator 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   5520
      X2              =   5715
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line HorizontalCalendarLine 
      BorderColor     =   &H00000000&
      Index           =   0
      Visible         =   0   'False
      X1              =   1200
      X2              =   1800
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line VerticalCalendarLine 
      BorderColor     =   &H00000000&
      Index           =   0
      Visible         =   0   'False
      X1              =   840
      X2              =   840
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Label lblCalDates 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "COA_CalRepDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngCALDATES_BOXWIDTH As Long
Private mlngCALDATES_BOXHEIGHT As Long
Private mlngCALDATES_BOXSTARTX As Long
Private mlngCALDATES_BOXSTARTY As Long

Event Click(pvarLabel As Variant)

Private mlngLineIndexCount_VER As Long
Private mlngLineIndexCount_HOR As Long

Private mlngBoxSize As Long

Public Property Let BoxSize(NEW_VALUE As Long)
  mlngBoxSize = NEW_VALUE
  Load_Labels
End Property
Public Property Get BoxSize() As Long
  BoxSize = mlngBoxSize
End Property

Private Function Load_CalendarVerticalLines() As Boolean

  Dim count_vert As Integer
  
  For count_vert = 1 To 37 Step 1
    Load VerticalCalendarLine(count_vert)
    VerticalCalendarLine(count_vert).Visible = True
    VerticalCalendarLine(count_vert).X1 = lblCalDates(count_vert * 2).Left
    VerticalCalendarLine(count_vert).X2 = VerticalCalendarLine(count_vert).X1
    VerticalCalendarLine(count_vert).Y1 = -10
    VerticalCalendarLine(count_vert).Y2 = mlngCALDATES_BOXHEIGHT * 2
    VerticalCalendarLine(count_vert).ZOrder 0
  Next count_vert
  
  count_vert = (37 + 1)
  Load VerticalCalendarLine(count_vert)
  VerticalCalendarLine(count_vert).Visible = True
  VerticalCalendarLine(count_vert).X1 = (lblCalDates(((count_vert - 1) * 2)).Left + lblCalDates(((count_vert - 1) * 2)).Width) - 10
  VerticalCalendarLine(count_vert).X2 = VerticalCalendarLine(count_vert).X1
  VerticalCalendarLine(count_vert).Y1 = -10
  VerticalCalendarLine(count_vert).Y2 = (mlngCALDATES_BOXHEIGHT * 2) + 20
  VerticalCalendarLine(count_vert).ZOrder 0

  Load_CalendarVerticalLines = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_CalendarVerticalLines = False
  GoTo TidyUpAndExit
  
End Function

Private Function Load_CalendarHorizontalLines() As Boolean

  On Error GoTo ErrorTrap
  
  Dim intNewIndex As Integer
  
  intNewIndex = HorizontalCalendarLine.UBound + 1
  
  Load HorizontalCalendarLine(intNewIndex)
  With HorizontalCalendarLine(intNewIndex)
    .X1 = -10
    .X2 = (lblCalDates(lblCalDates.UBound).Left + lblCalDates(lblCalDates.UBound).Width) + 20
    .Y1 = (2 * mlngCALDATES_BOXHEIGHT) - 5
    .Y2 = .Y1
    .BorderWidth = 1
    .Visible = True
    .ZOrder 0
  End With
  
  Load_CalendarHorizontalLines = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_CalendarHorizontalLines = False
  GoTo TidyUpAndExit
  
End Function

Public Sub AddEventSeparator(pintIndex As Integer, pstrSession As String, _
                             pblnHasEvent As Boolean, pblnNextHasEvent As Boolean, _
                             pblnNext2HasEvent As Boolean, pblnPrevHasEvent As Boolean)

  'pintIndex is the CalDate index.
  
  Dim intNewIndex_HOR As Integer
  Dim intNewIndex_VER As Integer

  If mlngLineIndexCount_VER < VerticalSeparator.UBound Then
    intNewIndex_VER = mlngLineIndexCount_VER + 1
    mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
  Else
    intNewIndex_VER = VerticalSeparator.UBound + 1
    Load VerticalSeparator(intNewIndex_VER)
    mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
  End If

  If pstrSession = "AM" Then
    If pblnNextHasEvent Then
      If mlngLineIndexCount_HOR < HorizontalSeparator.UBound Then
        intNewIndex_HOR = mlngLineIndexCount_HOR + 1
        mlngLineIndexCount_HOR = mlngLineIndexCount_HOR + 1
      Else
        intNewIndex_HOR = HorizontalSeparator.UBound + 1
        Load HorizontalSeparator(intNewIndex_HOR)
        mlngLineIndexCount_HOR = mlngLineIndexCount_HOR + 1
      End If

      With HorizontalSeparator(intNewIndex_HOR)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 2
        .X1 = lblCalDates(pintIndex).Left
        .X2 = (.X1 + mlngCALDATES_BOXWIDTH - 15)
        .Y1 = (lblCalDates(pintIndex).Top + mlngCALDATES_BOXWIDTH - 15)
        .Y2 = .Y1
        .ZOrder 0
      End With
    End If  'pblnNextHasEvent

    If pblnNext2HasEvent Then
      If mlngLineIndexCount_VER < VerticalSeparator.UBound Then
        intNewIndex_VER = mlngLineIndexCount_VER + 1
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      Else
        intNewIndex_VER = VerticalSeparator.UBound + 1
        Load VerticalSeparator(intNewIndex_VER)
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      End If

      With VerticalSeparator(intNewIndex_VER)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 2
        .X1 = lblCalDates(pintIndex + 2).Left
        .X2 = .X1
        .Y1 = lblCalDates(pintIndex + 2).Top
        .Y2 = (.Y1 + mlngCALDATES_BOXHEIGHT - 15)
        .ZOrder 0
      End With
    End If  'pblnNext2HasEvent

    If pblnPrevHasEvent Then
      If mlngLineIndexCount_VER < VerticalSeparator.UBound Then
        intNewIndex_VER = mlngLineIndexCount_VER + 1
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      Else
        intNewIndex_VER = VerticalSeparator.UBound + 1
        Load VerticalSeparator(intNewIndex_VER)
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      End If

      With VerticalSeparator(intNewIndex_VER)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 2
        .X1 = (lblCalDates(pintIndex + 1).Left + 7)
        .X2 = .X1
        .Y1 = (lblCalDates(pintIndex + 1).Top - 20)
        .Y2 = (.Y1 + mlngCALDATES_BOXHEIGHT + 15)
        .ZOrder 0
      End With
    End If  'pblnPrevHasEvent

  End If  'pstrSession = "AM"

  If pstrSession = "PM" Then
    If pblnPrevHasEvent And pblnNextHasEvent Then
      If mlngLineIndexCount_VER < VerticalSeparator.UBound Then
        intNewIndex_VER = mlngLineIndexCount_VER + 1
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      Else
        intNewIndex_VER = VerticalSeparator.UBound + 1
        Load VerticalSeparator(intNewIndex_VER)
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      End If

      With VerticalSeparator(intNewIndex_VER)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 2
        .X1 = lblCalDates(pintIndex + 1).Left
        .X2 = .X1
        .Y1 = lblCalDates(pintIndex + 1).Top
        .Y2 = (.Y1 + (mlngCALDATES_BOXHEIGHT * 2) - 20)
        .ZOrder 0
      End With

    ElseIf pblnNextHasEvent Or pblnNext2HasEvent Then
      If mlngLineIndexCount_VER < VerticalSeparator.UBound Then
        intNewIndex_VER = mlngLineIndexCount_VER + 1
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      Else
        intNewIndex_VER = VerticalSeparator.UBound + 1
        Load VerticalSeparator(intNewIndex_VER)
        mlngLineIndexCount_VER = mlngLineIndexCount_VER + 1
      End If

      With VerticalSeparator(intNewIndex_VER)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 2
        .X1 = lblCalDates(pintIndex + 2).Left
        .X2 = .X1
        .Y1 = lblCalDates(pintIndex + 2).Top
        .Y2 = .Y1 + (mlngCALDATES_BOXHEIGHT)
        .ZOrder 0
      End With
    End If

  End If  'pstrSession = "PM"
  
End Sub

Public Function HideSeparators() As Boolean

  Dim intIndex As Integer
  
  For intIndex = 1 To HorizontalSeparator.UBound Step 1
    HorizontalSeparator(intIndex).Visible = False
  Next intIndex

  For intIndex = 1 To VerticalSeparator.UBound Step 1
    VerticalSeparator(intIndex).Visible = False
  Next intIndex

End Function

Private Sub lblCalDates_Click(Index As Integer)
  RaiseEvent Click(lblCalDates(Index))
End Sub

Public Sub Load_Labels()
  
  UnloadControls
  
  Dim lngCount_X As Long
  Dim lngCount As Long
  Dim lngNewIndex As Long
  Dim lngLeft As Long
  Dim lngTop As Long

  mlngCALDATES_BOXWIDTH = IIf(mlngBoxSize < 1, 200, mlngBoxSize + 5)
  mlngCALDATES_BOXHEIGHT = IIf(mlngBoxSize < 1, 200, mlngBoxSize + 5)
  mlngCALDATES_BOXSTARTX = 0
  mlngCALDATES_BOXSTARTY = 0

  lblCalDates(0).Width = IIf(mlngBoxSize < 1, 195, mlngBoxSize)
  lblCalDates(0).Height = IIf(mlngBoxSize < 1, 195, mlngBoxSize)
  
  lngTop = 0
  
  For lngCount_X = 0 To 36
    lngLeft = (mlngCALDATES_BOXSTARTX + ((mlngCALDATES_BOXWIDTH - 15) * lngCount_X))
    
    For lngCount = 0 To 1
      lngNewIndex = CLng(lblCalDates().UBound) + 1
      
      Load lblCalDates(lngNewIndex)

      With lblCalDates(lngNewIndex)
        .Move lngLeft, lngTop + ((mlngCALDATES_BOXHEIGHT) * lngCount)
        .Visible = True
      End With
      
    Next lngCount
  Next lngCount_X

  Load_CalendarVerticalLines
  Load_CalendarHorizontalLines
  
  Width = lblCalDates(74).Left + mlngCALDATES_BOXWIDTH
  Height = mlngCALDATES_BOXHEIGHT * 2

End Sub

Private Sub UserControl_Initialize()
  Load_Labels
End Sub

Public Function CalDate(pintIndex As Integer) As Variant
  Set CalDate = lblCalDates(pintIndex)
End Function

Private Function UnloadControls() As Boolean

  Dim iCtlCount As Integer
  Dim objLab As Label
  Dim objLine As Line
  
  ' unload all the calendar session labels
  For Each objLab In lblCalDates
    If objLab.Index > 0 Then
      Unload objLab
    End If
  Next objLab
  
  ' unload all the vertical calendar lines
  For Each objLine In VerticalCalendarLine
    If objLine.Index > 0 Then
      Unload objLine
    End If
  Next objLine
  
  ' unload all the horizontal calendar lines
  For Each objLine In HorizontalCalendarLine
    If objLine.Index > 0 Then
      Unload objLine
    End If
  Next objLine
  
  ' unload all the Vertical Separator lines
  For Each objLine In VerticalSeparator
    If objLine.Index > 0 Then
      Unload objLine
    End If
  Next objLine
  
  ' unload all the Horizontal Separator lines
  For Each objLine In HorizontalSeparator
    If objLine.Index > 0 Then
      Unload objLine
    End If
  Next objLine
  
End Function

