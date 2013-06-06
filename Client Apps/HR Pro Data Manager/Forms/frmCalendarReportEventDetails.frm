VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "COA_WorkingPattern.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmCalendarReportEventDetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1068
   Icon            =   "frmCalendarReportEventDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3670
      TabIndex        =   0
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   5205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4755
      Begin COALine.COA_Line ASRLine 
         Height          =   30
         Index           =   0
         Left            =   300
         Top             =   960
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   53
      End
      Begin COALine.COA_Line ASRLine 
         Height          =   30
         Index           =   2
         Left            =   300
         Top             =   3000
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   53
      End
      Begin COALine.COA_Line ASRLine 
         Height          =   30
         Index           =   1
         Left            =   300
         Top             =   2160
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   53
      End
      Begin COALine.COA_Line ASRLine 
         Height          =   30
         Index           =   3
         Left            =   300
         Top             =   3645
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   53
      End
      Begin COAWorkingPattern.COA_WorkingPattern ASRWorkingPattern1 
         Height          =   765
         Left            =   1995
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4260
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1349
      End
      Begin VB.Label lblEventDesc1 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   16
         Top             =   2355
         Width           =   2600
      End
      Begin VB.Label lblWPatternLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Working Pattern :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   4260
         Width           =   1590
      End
      Begin VB.Label lblRegion 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   21
         Top             =   3855
         Width           =   2600
      End
      Begin VB.Label lblRegionLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Region :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   3855
         Width           =   1245
      End
      Begin VB.Label lblEventDesc2Label 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   2655
         Width           =   1665
      End
      Begin VB.Label lblEventDesc2 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   17
         Top             =   2655
         Width           =   2600
      End
      Begin VB.Label lblEventDesc1Label 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   2355
         Width           =   1665
      End
      Begin VB.Label lblBaseDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1995
         TabIndex        =   14
         Top             =   600
         Width           =   2600
      End
      Begin VB.Label lblBaseDescLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   600
         Width           =   1250
      End
      Begin VB.Label lblEventName 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   12
         Top             =   300
         Width           =   2600
      End
      Begin VB.Label lblEventNameLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Event Name :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   300
         Width           =   1250
      End
      Begin VB.Label lblStartDateLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   1155
         Width           =   1250
      End
      Begin VB.Label lblEndDateLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   200
         TabIndex        =   8
         Top             =   1455
         Width           =   1250
      End
      Begin VB.Label lblCalendarCodeLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar Code :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   3255
         Width           =   1470
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1995
         TabIndex        =   6
         Top             =   1155
         Width           =   270
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1995
         TabIndex        =   5
         Top             =   1455
         Width           =   270
      End
      Begin VB.Label lblCalendarCode 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   4
         Top             =   3255
         Width           =   2600
      End
      Begin VB.Label lblDuration 
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2000
         TabIndex        =   3
         Top             =   1755
         Width           =   2600
      End
      Begin VB.Label lblDurationLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration (Actual) :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   1755
         Width           =   1635
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdDetails 
      Height          =   855
      Left            =   1680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   8415
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   14
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   2
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   14
      Columns(0).Width=   1905
      Columns(0).Caption=   "EventName"
      Columns(0).Name =   "EventName"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2090
      Columns(1).Caption=   "BaseDescription"
      Columns(1).Name =   "BaseDescription"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1296
      Columns(2).Caption=   "StartDate"
      Columns(2).Name =   "StartDate"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1138
      Columns(3).Caption=   "StartSession"
      Columns(3).Name =   "StartSession"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1191
      Columns(4).Caption=   "EndDate"
      Columns(4).Name =   "EndDate"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   979
      Columns(5).Caption=   "EndSession"
      Columns(5).Name =   "EndSession"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   767
      Columns(6).Caption=   "Duration"
      Columns(6).Name =   "Duration"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1323
      Columns(7).Caption=   "EventDescription1Column"
      Columns(7).Name =   "EventDescription1Column"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1270
      Columns(8).Caption=   "EventDescription1Value"
      Columns(8).Name =   "EventDescription1Value"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   1000
      Columns(9).Width=   1244
      Columns(9).Caption=   "EventDescription2Column"
      Columns(9).Name =   "EventDescription2Column"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1111
      Columns(10).Caption=   "EventDescription2Value"
      Columns(10).Name=   "EventDescription2Value"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   1000
      Columns(11).Width=   714
      Columns(11).Caption=   "Legend"
      Columns(11).Name=   "Legend"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1429
      Columns(12).Caption=   "WorkingPattern"
      Columns(12).Name=   "WorkingPattern"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1323
      Columns(13).Caption=   "Region"
      Columns(13).Name=   "Region"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14843
      _ExtentY        =   1508
      _StockProps     =   79
      BackColor       =   12632256
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   1920
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmCalendarReportEventDetails.frx":000C
   End
End
Attribute VB_Name = "frmCalendarReportEventDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private mcolEvents As clsCalendarEvents

Private Const LABELWP_STARTY = 4255

Private mblnShowRegion As Boolean
Private mblnShowWorkingPattern As Boolean

Public Property Let BreakdownCaption(pstrCaption As String)
  Caption = pstrCaption
End Property
Public Function Initialse(pcolEvents As clsCalendarEvents) As Boolean

  Dim fOK As Boolean
  
  fOK = True
  
  Set mcolEvents = pcolEvents
  
  If fOK Then fOK = PopulateGrid

  If fOK Then fOK = ReformatScreen

  Initialse = fOK
  
End Function
Private Function PopulateGrid() As Boolean

  Dim objEvent As clsCalendarEvent
  
  Dim strAddLine As String
  Dim iDecimals As Integer
  
  Dim fOK As Boolean
  
  fOK = True
 
  For Each objEvent In mcolEvents.Collection
    strAddLine = vbNullString
    
    With objEvent
      strAddLine = strAddLine & .Name & vbTab
      strAddLine = strAddLine & .BaseDescription & vbTab
      strAddLine = strAddLine & .StartDateName & vbTab
      strAddLine = strAddLine & .StartSessionName & vbTab
      strAddLine = strAddLine & .EndDateName & vbTab
      strAddLine = strAddLine & .EndSessionName & vbTab
      strAddLine = strAddLine & .DurationName & vbTab
      
      strAddLine = strAddLine & Replace(.Description1Name, "_", " ") & vbTab
      iDecimals = datGeneral.GetDecimalsSize(.Description1ID)
      strAddLine = strAddLine & _
        IIf(datGeneral.DoesColumnUseSeparators(.Description1ID), Format(.Desc1Value, "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "0"), "")), .Desc1Value) & vbTab
      
      strAddLine = strAddLine & Replace(.Description2Name, "_", " ") & vbTab
      iDecimals = datGeneral.GetDecimalsSize(.Description2ID)
      strAddLine = strAddLine & _
        IIf(datGeneral.DoesColumnUseSeparators(.Description2ID), Format(.Desc2Value, "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "0"), "")), .Desc2Value) & vbTab
      
      strAddLine = strAddLine & .LegendCharacter & vbTab
      strAddLine = strAddLine & .WorkingPattern & vbTab
      strAddLine = strAddLine & .Region
    End With
    
    grdDetails.AddItem strAddLine
  Next objEvent

  If grdDetails.Rows > 0 Then
    grdDetails.MoveFirst
    grdDetails.SelBookmarks.RemoveAll
    grdDetails.SelBookmarks.Add grdDetails.Bookmark
    UpdateLabels
    UpdateRecordStatus
  End If
  
  fOK = True
  
TidyUpAndExit:
  PopulateGrid = fOK
  Set objEvent = Nothing
  Exit Function

ErrorTrap:
  PopulateGrid = False
  GoTo TidyUpAndExit

End Function
Private Function ReformatScreen() As Boolean

  On Error GoTo ErrorTrap
  
  ReformatScreen = True
  
  If mblnShowRegion And mblnShowWorkingPattern Then
    lblRegion.Visible = True
    lblRegionLabel.Visible = True
    lblWPatternLabel.Visible = True
    ASRWorkingPattern1.Visible = True
    
    lblWPatternLabel.Top = LABELWP_STARTY
    ASRWorkingPattern1.Top = LABELWP_STARTY

  ElseIf mblnShowRegion Then
    lblRegion.Visible = True
    lblRegionLabel.Visible = True
    lblWPatternLabel.Visible = False
    ASRWorkingPattern1.Visible = False
    
    fraDetails.Height = lblRegion.Top + lblRegion.Height + 100
    
  ElseIf mblnShowWorkingPattern Then
    lblRegion.Visible = False
    lblRegionLabel.Visible = False
    lblWPatternLabel.Visible = True
    ASRWorkingPattern1.Visible = True
    
    lblWPatternLabel.Top = lblRegion.Top
    ASRWorkingPattern1.Top = lblWPatternLabel.Top
    
    fraDetails.Height = ASRWorkingPattern1.Top + ASRWorkingPattern1.Height + 100
    
  Else
    lblRegion.Visible = False
    lblRegionLabel.Visible = False
    lblWPatternLabel.Visible = False
    ASRWorkingPattern1.Visible = False
    
    fraDetails.Height = ASRLine(3).Top - 15
    
  End If
  
  cmdOK.Top = fraDetails.Top + fraDetails.Height + 120
  Me.Height = cmdOK.Top + cmdOK.Height + 120 + 780
  
  ASRWorkingPattern1.Enabled = False
  
  ReformatScreen = True
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  ReformatScreen = False
  GoTo TidyUpAndExit

End Function
Public Property Let ShowRegion(pblnShowRegion As Boolean)
  mblnShowRegion = pblnShowRegion
End Property
Public Property Let ShowWorkingPattern(pblnShowWorkingPattern As Boolean)
  mblnShowWorkingPattern = pblnShowWorkingPattern
End Property
Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  Select Case Tool.Name
  
    Case "First"
    
      grdDetails.MoveFirst
      
    Case "Previous"
    
      grdDetails.MovePrevious
    
    Case "Next"
    
      grdDetails.MoveNext
    
    Case "Last"
    
      grdDetails.MoveLast
  
  End Select
  
  grdDetails.SelBookmarks.RemoveAll
  grdDetails.SelBookmarks.Add grdDetails.Bookmark
  
  UpdateLabels
  UpdateRecordStatus
  
End Sub
Private Sub UpdateLabels()
  
  lblEventName = Replace(grdDetails.Columns("EventName").Text, "&", "&&")
  lblBaseDesc = Replace(grdDetails.Columns("BaseDescription").Text, "&", "&&")
  '------------------------------------------------------------------------------
  'lblStartDate.Caption = IIf(grdDetails.Columns("StartDate").Text = "", "<None>", Format(grdDetails.Columns("StartDate").Text, DateFormat))
  'lblStartSession.Caption = grdDetails.Columns("StartSession").Text
  'lblEndDate.Caption = IIf(grdDetails.Columns("EndDate").Text = "", "<None>", Format(grdDetails.Columns("EndDate").Text, DateFormat))
  'lblEndSession.Caption = grdDetails.Columns("EndSession").Text
  
  lblStartDate.Caption = IIf(grdDetails.Columns("StartDate").Text = "", "<None>", Format(grdDetails.Columns("StartDate").Text, DateFormat)) & _
                         " " & grdDetails.Columns("StartSession").Text
  lblEndDate.Caption = IIf(grdDetails.Columns("EndDate").Text = "", "<None>", Format(grdDetails.Columns("EndDate").Text, DateFormat)) & _
                        " " & grdDetails.Columns("EndSession").Text
  lblDuration.Caption = IIf(IsNull(grdDetails.Columns("Duration").Text), "", grdDetails.Columns("Duration").Text)
  '------------------------------------------------------------------------------
  
  
  If Trim(Replace(grdDetails.Columns("EventDescription1Column").Text, "&", "&&")) = vbNullString Then
    lblEventDesc1Label.Visible = False
    lblEventDesc1.Visible = False
  Else
    lblEventDesc1Label.Visible = True
    lblEventDesc1.Visible = True
    lblEventDesc1Label.Caption = Replace(grdDetails.Columns("EventDescription1Column").Text, "&", "&&") & " : "
    lblEventDesc1.Caption = Replace(grdDetails.Columns("EventDescription1Value").Text, "&", "&&")
  End If
  
  If Trim(Replace(grdDetails.Columns("EventDescription2Column").Text, "&", "&&")) = vbNullString Then
    lblEventDesc2Label.Visible = False
    lblEventDesc2.Visible = False
  Else
    lblEventDesc2Label.Visible = True
    lblEventDesc2.Visible = True
    lblEventDesc2Label.Caption = Replace(grdDetails.Columns("EventDescription2Column").Text, "&", "&&") & " : "
    lblEventDesc2.Caption = Replace(grdDetails.Columns("EventDescription2Value").Text, "&", "&&")
  End If
  
  '------------------------------------------------------------------------------
  lblCalendarCode.Caption = Replace(grdDetails.Columns("Legend").Text, "&", "&&")
  '------------------------------------------------------------------------------
  ASRWorkingPattern1.Value = grdDetails.Columns("WorkingPattern").Text
  lblRegion.Caption = Replace(grdDetails.Columns("Region").Text, "&", "&&")
  
End Sub
Private Sub UpdateRecordStatus()

  With Me.ActiveBar1
    If Me.grdDetails.Rows = 1 Then
      .Bands("bndDetails").Tools("First").Enabled = False
      .Bands("bndDetails").Tools("Previous").Enabled = False
      .Bands("bndDetails").Tools("Next").Enabled = False
      .Bands("bndDetails").Tools("Last").Enabled = False
    ElseIf Me.grdDetails.AddItemRowIndex(Me.grdDetails.Bookmark) = 0 Then
      .Bands("bndDetails").Tools("First").Enabled = False
      .Bands("bndDetails").Tools("Previous").Enabled = False
      .Bands("bndDetails").Tools("Next").Enabled = True
      .Bands("bndDetails").Tools("Last").Enabled = True
    ElseIf Me.grdDetails.AddItemRowIndex(Me.grdDetails.Bookmark) = (Me.grdDetails.Rows - 1) Then
      .Bands("bndDetails").Tools("First").Enabled = True
      .Bands("bndDetails").Tools("Previous").Enabled = True
      .Bands("bndDetails").Tools("Next").Enabled = False
      .Bands("bndDetails").Tools("Last").Enabled = False
    Else
      .Bands("bndDetails").Tools("First").Enabled = True
      .Bands("bndDetails").Tools("Previous").Enabled = True
      .Bands("bndDetails").Tools("Next").Enabled = True
      .Bands("bndDetails").Tools("Last").Enabled = True
    End If
  
    .Bands("bndDetails").Tools("Record").Caption = "Record " & (Me.grdDetails.AddItemRowIndex(Me.grdDetails.Bookmark) + 1) & " of " & Me.grdDetails.Rows
    .Refresh
  
  End With
  
End Sub
Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True

End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



