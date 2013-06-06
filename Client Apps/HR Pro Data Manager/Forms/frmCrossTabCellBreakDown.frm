VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCrossTabCellBreakDown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cross Tabs Cell Breakdown"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1016
   Icon            =   "frmCrossTabCellBreakDown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboValue 
      Height          =   315
      Index           =   2
      Left            =   2655
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3345
   End
   Begin VB.ComboBox cboValue 
      Height          =   315
      Index           =   0
      Left            =   2655
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   585
      Width           =   3345
   End
   Begin VB.ComboBox cboValue 
      Height          =   315
      Index           =   1
      Left            =   2655
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   990
      Width           =   3345
   End
   Begin VB.TextBox txtCellValue 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2655
      TabIndex        =   7
      Text            =   "<Selected Cell Value>"
      Top             =   1395
      Width           =   3345
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4815
      TabIndex        =   9
      Top             =   4680
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Height          =   2655
      Left            =   130
      TabIndex        =   8
      Top             =   1890
      Visible         =   0   'False
      Width           =   5865
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      BorderStyle     =   0
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   2
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
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
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Record Description"
      Columns(0).Name =   "Record Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Caption=   "Intersection Value"
      Columns(1).Name =   "Intersection Value"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   10345
      _ExtentY        =   4683
      _StockProps     =   79
      BackColor       =   16777215
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
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Type> :"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6000
      X2              =   6000
      Y1              =   1875
      Y2              =   4550
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   1875
      Y2              =   4550
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   135
      X2              =   6000
      Y1              =   4550
      Y2              =   4550
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   135
      X2              =   6000
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Label lblVerticalFieldName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vertical Field Name> :"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1035
      Width           =   2235
   End
   Begin VB.Label lblHorizontalFieldName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Horizontal Field Name> :"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   630
      Width           =   2430
   End
   Begin VB.Label lblPageBreakFieldName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Page Break Field Name> :"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   2430
   End
End
Attribute VB_Name = "frmCrossTabCellBreakDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParentForm As frmCrossTabRun
Private mlngReportMode As CrossTabType

Public Sub Initialise(frmParent As frmCrossTabRun)
  Set mfrmParentForm = frmParent
  mlngReportMode = cttNormal
End Sub

Private Sub cboValue_Click(Index As Integer)
  If Not (mfrmParentForm Is Nothing) Then
    Call mfrmParentForm.BreakdownComboClick(Index)
  End If
  
  ' JDM - Fault 2433 - Different captions for different day types
  If mlngReportMode = cttAbsenceBreakdown Then

    Select Case cboValue(0).Text
      Case "Total"
        SSDBGrid1.Columns(3).Visible = True
        SSDBGrid1.Columns(3).Caption = "Duration"
        'TM20011203 Fault 3230 - Set the alignment of the Duration column.
        SSDBGrid1.Columns(3).Alignment = ssCaptionAlignmentRight
        SSDBGrid1.Columns(3).CaptionAlignment = ssColCapAlignCenter
        
      Case "<All>"
        SSDBGrid1.Columns(3).Visible = True
        SSDBGrid1.Columns(3).Caption = "Duration"
        'TM20011203 Fault 3230 - Set the alignment of the Duration column.
        SSDBGrid1.Columns(3).Alignment = ssCaptionAlignmentRight
        SSDBGrid1.Columns(3).CaptionAlignment = ssColCapAlignCenter
        
      Case "Count"
        SSDBGrid1.Columns(3).Visible = False
        
      Case Else
        SSDBGrid1.Columns(3).Visible = True
        SSDBGrid1.Columns(3).Caption = cboValue(0).Text & "'s taken"
        'TM20011203 Fault 3230 - Set the alignment of the Duration column.
        SSDBGrid1.Columns(3).Alignment = ssCaptionAlignmentRight
        SSDBGrid1.Columns(3).CaptionAlignment = ssColCapAlignCenter
        
    End Select

  
  End If
  
End Sub

Private Sub cmdOK_Click()
  Me.Hide
End Sub

Private Sub Form_Activate()
  'MSFlexGrid1.SetFocus
  'If SSDBGrid1.Visible Then
  '  SSDBGrid1.SetFocus
  'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Only hide form (not unload) if user clicks 'X'
  If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub SSDBGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
  
'  Dim lngWidth As Long
'  Dim intCount As Integer

'  If Not (mfrmParentForm Is Nothing) Then
'    Call mfrmParentForm.SizeBreakdownColumns
'  End If

'  With SSDBGrid1
'    .Columns(ColIndex).Width = .ResizeWidth
'
'    lngWidth = 0
'    For intCount = 0 To .Cols - 2
'      lngWidth = lngWidth + .Columns(intCount).Width
'    Next
'    .Columns(.Cols - 1).Width = (.Width - lngWidth)
'
'    'lngWidth = 0
'    'For intCount = 0 To .Cols - 2
'    '  lngWidth = lngWidth + .Columns(intCount).Width
'    'Next
'    '.Columns(.Cols - 1).Width = (.Width - lngWidth)
'  End With

End Sub
' A property to specify that this is a drilldown for a standard report
Public Property Let ReportMode(ByVal plNewValue As CrossTabType)

  mlngReportMode = plNewValue

End Property

