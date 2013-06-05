VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "COA_WorkingPattern.ocx"
Begin VB.Form frmAbsenceCalendarDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Absence Calendar"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1005
   Icon            =   "frmAbsenceCalendarDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   3645
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   3780
      Begin COAWorkingPattern.COA_WorkingPattern ASRWorkingPattern1 
         Height          =   765
         Left            =   1800
         TabIndex        =   17
         Top             =   2715
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1349
      End
      Begin VB.Label lblDurationLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duration :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   915
         Width           =   1350
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   21
         Top             =   915
         Width           =   240
      End
      Begin VB.Label lblRegionLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   2415
         Width           =   600
      End
      Begin VB.Label lblRegion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   19
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label lblWPatternLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Working Pattern :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   2715
         Width           =   1500
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   15
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label lblTypeCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   14
         Top             =   1515
         Width           =   240
      End
      Begin VB.Label lblCalendarCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   13
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   12
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lblEndSession 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2940
         TabIndex        =   11
         Top             =   615
         Width           =   285
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   10
         Top             =   615
         Width           =   240
      End
      Begin VB.Label lblStartSession 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2940
         TabIndex        =   9
         Top             =   315
         Width           =   285
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1815
         TabIndex        =   8
         Top             =   315
         Width           =   240
      End
      Begin VB.Label lblReasonLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   2115
         Width           =   645
      End
      Begin VB.Label lblTypeCodeLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Code :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   1515
         Width           =   1245
      End
      Begin VB.Label lblCalendarCodeLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar Code :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1815
         Width           =   1440
      End
      Begin VB.Label lblTypeLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label lblEndDateLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label lblStartDateLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2685
      TabIndex        =   0
      Top             =   3810
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid grdDetails 
      Height          =   390
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   1050
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   11
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   11
      Columns(0).Width=   873
      Columns(0).Caption=   "Start Date"
      Columns(0).Name =   "StartDate"
      Columns(0).Alignment=   2
      Columns(0).AllowSizing=   0   'False
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   873
      Columns(1).Caption=   "Start Session"
      Columns(1).Name =   "Start Session"
      Columns(1).Alignment=   2
      Columns(1).AllowSizing=   0   'False
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   873
      Columns(2).Caption=   "End Date"
      Columns(2).Name =   "End Date"
      Columns(2).Alignment=   2
      Columns(2).AllowSizing=   0   'False
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   873
      Columns(3).Caption=   "End Session"
      Columns(3).Name =   "End Session"
      Columns(3).Alignment=   2
      Columns(3).AllowSizing=   0   'False
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   873
      Columns(4).Caption=   "Type"
      Columns(4).Name =   "Type"
      Columns(4).Alignment=   2
      Columns(4).AllowSizing=   0   'False
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   873
      Columns(5).Caption=   "Calendar Code"
      Columns(5).Name =   "Calendar Code"
      Columns(5).Alignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   873
      Columns(6).Caption=   "Type Code"
      Columns(6).Name =   "Type Code"
      Columns(6).Alignment=   2
      Columns(6).AllowSizing=   0   'False
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   873
      Columns(7).Caption=   "Reason"
      Columns(7).Name =   "Reason"
      Columns(7).AllowSizing=   0   'False
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "WPattern"
      Columns(8).Name =   "WPattern"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Region"
      Columns(9).Name =   "Region"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "Duration"
      Columns(10).Name=   "Duration"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   1852
      _ExtentY        =   688
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
      Left            =   2145
      Top             =   3780
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
      Bands           =   "frmAbsenceCalendarDetails.frx":000C
   End
End
Attribute VB_Name = "frmAbsenceCalendarDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnShowRegion As Boolean
Private mblnShowWP As Boolean

Public Function Initialise(rsAbsenceRecords As Recordset, dtmQueryDate As Date, strSession As String, _
                            pbShowRegion As Boolean, pbShowWP As Boolean) As Boolean

  On Error GoTo Initialise_ERROR
  
  Dim rsTemp As Recordset
  Dim strQueryWP As String
  Dim strQueryRegion As String
  
  mblnShowRegion = pbShowRegion
  mblnShowWP = pbShowWP
  lblRegionLabel.Visible = mblnShowRegion
  lblRegion.Visible = lblRegionLabel.Visible
  lblWPatternLabel.Visible = mblnShowWP
  ASRWorkingPattern1.Visible = lblWPatternLabel.Visible
  
  If frmAbsenceCalendar.WPsEnabled Then
    ' Get the working pattern for the date the user clicked on
    If gwptWorkingPatternType = wptHistoricWPattern Then
      Set rsTemp = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " AS 'Date', " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & _
                                                "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & _
                                               "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & frmAbsenceCalendar.mlngPersonnelID & " " & _
                                               "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " <= '" & Replace(Format(dtmQueryDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                               "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")
      
      If rsTemp.BOF And rsTemp.EOF Then
        strQueryWP = Space(14)
      Else
        strQueryWP = rsTemp.Fields("WP").Value
      End If
      
    Else
      strQueryWP = frmAbsenceCalendar.ASRWorkingPattern1.Value
    End If
  End If
  
  If frmAbsenceCalendar.RegionsEnabled Then
    ' Get the region for the date the user clicked on
    If grtRegionType = rtHistoricRegion Then
      Set rsTemp = datGeneral.GetRecords("SELECT " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                         "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                         "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & frmAbsenceCalendar.mlngPersonnelID & " " & _
                                         "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " <= '" & Replace(Format(dtmQueryDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                         "ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC")
    
      If rsTemp.BOF And rsTemp.EOF Then
        strQueryRegion = ""
      Else
        strQueryRegion = rsTemp.Fields("Region").Value
      End If
      
    Else
      strQueryRegion = frmAbsenceCalendar.lblRegion.Caption
    End If
  End If
  
  With rsAbsenceRecords
  
    .MoveFirst
    
    Do Until .EOF
    
      ' Does the date clicked on fall in the middle of an absence record ? If so, add it
      If (.Fields("startdate") < dtmQueryDate) And (dtmQueryDate <= IIf(IsNull(.Fields("enddate")), Date, .Fields("enddate"))) Then
        
        grdDetails.AddItem Format(.Fields("StartDate"), DateFormat) & vbTab & _
                            .Fields("StartSession") & vbTab & _
                            Format(.Fields("EndDate"), DateFormat) & vbTab & _
                            .Fields("EndSession") & vbTab & _
                            .Fields("Type") & vbTab & _
                            .Fields("CalendarCode") & vbTab & _
                            .Fields("Code") & vbTab & _
                            .Fields("Reason") & vbTab & _
                            strQueryWP & vbTab & _
                            strQueryRegion & vbTab & _
                            .Fields("Duration")
        
      ' Does the date clicked on equal the start of an absence record ? If so, check sessions
      ElseIf .Fields("startdate") = dtmQueryDate Then
        
        If UCase(.Fields("startsession")) = "AM" Then
        
        grdDetails.AddItem Format(.Fields("StartDate"), DateFormat) & vbTab & _
                            .Fields("StartSession") & vbTab & _
                            Format(.Fields("EndDate"), DateFormat) & vbTab & _
                            .Fields("EndSession") & vbTab & _
                            .Fields("Type") & vbTab & _
                            .Fields("CalendarCode") & vbTab & _
                            .Fields("Code") & vbTab & _
                            .Fields("Reason") & vbTab & _
                            strQueryWP & vbTab & _
                            strQueryRegion & vbTab & _
                            .Fields("Duration")
        
        End If
        
        If UCase(.Fields("startsession")) = "PM" Then
          If strSession = "PM" Then
            grdDetails.AddItem Format(.Fields("StartDate"), DateFormat) & vbTab & _
                                .Fields("StartSession") & vbTab & _
                                Format(.Fields("EndDate"), DateFormat) & vbTab & _
                                .Fields("EndSession") & vbTab & _
                                .Fields("Type") & vbTab & _
                                .Fields("CalendarCode") & vbTab & _
                                .Fields("Code") & vbTab & _
                                .Fields("Reason") & vbTab & _
                                strQueryWP & vbTab & _
                                strQueryRegion & vbTab & _
                                .Fields("Duration")
  
          End If
        End If
      
      ' Does the date clicked on equal the end of an absence record ? If so, check sessions
      ElseIf .Fields("enddate") = dtmQueryDate Then
      'ElseIf (.Fields("enddate") = dtmQueryDate) Or IsNull(.Fields("enddate")) Then
        
        If UCase(.Fields("endsession")) = "PM" Then
          grdDetails.AddItem Format(.Fields("StartDate"), DateFormat) & vbTab & _
                              .Fields("StartSession") & vbTab & _
                              Format(.Fields("EndDate"), DateFormat) & vbTab & _
                              .Fields("EndSession") & vbTab & _
                              .Fields("Type") & vbTab & _
                              .Fields("CalendarCode") & vbTab & _
                              .Fields("Code") & vbTab & _
                              .Fields("Reason") & vbTab & _
                              strQueryWP & vbTab & _
                              strQueryRegion & vbTab & _
                              .Fields("Duration")
        
        End If
        
        If UCase(.Fields("endsession")) = "AM" Then
          If strSession = "AM" Then
            grdDetails.AddItem Format(.Fields("StartDate"), DateFormat) & vbTab & _
                                .Fields("StartSession") & vbTab & _
                                Format(.Fields("EndDate"), DateFormat) & vbTab & _
                                .Fields("EndSession") & vbTab & _
                                .Fields("Type") & vbTab & _
                                .Fields("CalendarCode") & vbTab & _
                                .Fields("Code") & vbTab & _
                                .Fields("Reason") & vbTab & _
                                strQueryWP & vbTab & _
                                strQueryRegion & vbTab & _
                                .Fields("Duration")
          
          End If
        End If
      End If
          
      .MoveNext
    
    Loop
  
  End With
  
  With grdDetails
  
    If .Rows < 5 Then .Columns(7).Width = .Columns(7).Width + 200
    
    If .Rows = 0 Then
      Initialise = False
      Exit Function
    Else
      Initialise = True
    End If
    
    .MoveFirst
    .SelBookmarks.Add .Bookmark
  
  End With
  
  UpdateLabels
  UpdateRecordStatus
  
  Me.ASRWorkingPattern1.Enabled = False
  Me.ASRWorkingPattern1.ForeColor = vbHighlight ' &H8000000D&
  Me.ASRWorkingPattern1.BorderStyle = 0
  Set rsTemp = Nothing
  Exit Function
  
Initialise_ERROR:

  Initialise = False
  
  COAMsgBox "An error has occurred whilst retrieving data. Please ensure your" & _
  vbCrLf & "Absence module is setup correctly." & vbCrLf & vbCrLf & _
  "If contacting support, please state:" & vbCrLf & Err.Number & _
  " - " & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"

End Function

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
  
  lblStartDate.Caption = IIf(grdDetails.Columns("StartDate").Text = "", "<None>", grdDetails.Columns("StartDate").Text)
  lblStartSession.Caption = grdDetails.Columns("Start Session").Text
  lblEndDate.Caption = IIf(grdDetails.Columns("End Date").Text = "", "<None>", grdDetails.Columns("End Date").Text)
  lblEndSession.Caption = grdDetails.Columns("End Session").Text
  lblType.Caption = Replace(grdDetails.Columns("Type").Text, "&", "&&")
  lblTypeCode.Caption = Replace(grdDetails.Columns("Type Code").Text, "&", "&&")
  lblCalendarCode.Caption = Replace(grdDetails.Columns("Calendar Code").Text, "&", "&&")
  lblReason.Caption = Replace(grdDetails.Columns("Reason").Text, "&", "&&")
  ASRWorkingPattern1.Value = grdDetails.Columns("WPattern").Text
  lblRegion.Caption = Replace(grdDetails.Columns("Region").Text, "&", "&&")
  lblDuration.Caption = IIf(IsNull(grdDetails.Columns("Duration").Text), "", grdDetails.Columns("Duration").Text)
  
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

Private Sub cmdOK_Click()

  Unload Me

End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True

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



