VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDefSel 
   Caption         =   "Select"
   ClientHeight    =   6465
   ClientLeft      =   2715
   ClientTop       =   2535
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1027
   Icon            =   "frmDefSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTopButtons 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3410
      Left            =   3240
      TabIndex        =   16
      Top             =   100
      Width           =   1215
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Re&fresh"
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   3000
         Width           =   1200
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Proper&ties..."
         Height          =   400
         Left            =   0
         TabIndex        =   8
         Top             =   2500
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   0
         TabIndex        =   7
         Top             =   2000
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit..."
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   500
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete..."
         Height          =   400
         Left            =   0
         TabIndex        =   6
         Top             =   1500
         Width           =   1200
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Cop&y..."
         Height          =   400
         Left            =   0
         TabIndex        =   5
         Top             =   1000
         Width           =   1200
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New..."
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5240
      Left            =   100
      TabIndex        =   14
      Top             =   100
      Width           =   3015
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Height          =   1080
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4155
         Width           =   3000
      End
      Begin VB.ComboBox cboTables 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.ListBox List2 
         Height          =   2985
         Left            =   0
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin ComctlLib.ListView List1 
         Height          =   3705
         Left            =   0
         TabIndex        =   1
         Top             =   405
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   6535
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Column"
            Object.Tag             =   "Column"
            Text            =   "Column"
            Object.Width           =   52917
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "SortKey"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTables 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table:"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   60
         Width           =   540
      End
   End
   Begin VB.Frame fraBottomButtons 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1360
      Left            =   3240
      TabIndex        =   17
      Top             =   4000
      Width           =   1215
      Begin VB.CommandButton cmdNone 
         Caption         =   "N&one"
         Height          =   400
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Height          =   400
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   1200
      End
   End
   Begin VB.CheckBox chkOnlyMine 
      Caption         =   "On&ly show definitions where owner is 'username'"
      Height          =   405
      Left            =   100
      TabIndex        =   13
      Top             =   5450
      Width           =   4320
   End
   Begin ActiveBarLibraryCtl.ActiveBar abDefSel 
      Left            =   3120
      Top             =   3720
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
      Bands           =   "frmDefSel.frx":000C
   End
End
Attribute VB_Name = "frmDefSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecords As ADODB.Recordset
Private mblnLoading As Boolean
Private mblnBatchPrompt As Boolean
Private mstrEventLogIDs As String
  
Private lngAction As Long
Private mlngOptions As Long
Private msSelectedText As String
Private mlngSelectedID As Long
Private mlngTableID As Long

Private mblnCaptionIsRun As Boolean
Private mblnApplyDefAccess As Boolean
Private mblnApplySystemPermissions As Boolean
Private mblnHideDesc As Boolean
Private mblnTableComboVisible As Boolean
Private mblnScheduledJobs As Boolean

Private msTableName As String
Private msFieldName As String
Private msIDField As String
Private msType As String
Private msTypeCode As String
Private msRecordSource As String
'Private msIcon As String
Private mutlUtilityType As UtilityType
Private msTableIDColumnName As String
Private msAccessTableName As String

Private mbFromCopy As Boolean
'Private mlCopyID As Long

Private mblnHiddenDef As Boolean
Private mblnReadOnlyAccess As Boolean

Private mfEnableNew As Boolean
Private mfEnableView As Boolean
Private mfEnableEdit As Boolean
Private mfEnableDelete As Boolean
Private mfEnableRun As Boolean

Private mblnFirstLoad As Boolean
Private mlngHeight As Long
Private mlngWidth As Long
Private mintOnlyMine As Integer

Private msGeneralCaption As String
Private msSingularCaption As String

Private malngSelectedIDs()

Public Property Get SelectedID() As Long
  SelectedID = mlngSelectedID
End Property

Public Property Let SelectedID(ByVal lngNewValue As Long)
  mlngSelectedID = lngNewValue
End Property

Public Property Get HideDescription() As Boolean
  HideDescription = mblnHideDesc
End Property

Public Property Let HideDescription(ByVal blnNewValue As Boolean)
  mblnHideDesc = blnNewValue
  txtDesc.Visible = Not mblnHideDesc
End Property

Public Property Get HiddenDef() As Boolean
  HiddenDef = mblnHiddenDef
End Property

Public Property Get Action() As Long
  Action = lngAction
End Property

Public Property Let Action(ByVal DefaultValue As Long)
  lngAction = DefaultValue
End Property

Public Property Get Options() As Long
  Options = mlngOptions
End Property

Public Property Let Options(OptionFlags As Long)
  mlngOptions = OptionFlags
End Property

Public Property Get TableComboVisible() As Boolean
  TableComboVisible = mblnTableComboVisible
End Property

Public Property Let TableComboVisible(ByVal blnNewValue As Boolean)
  mblnTableComboVisible = blnNewValue
End Property

Public Property Get TableComboEnabled() As Boolean
  TableComboEnabled = cboTables.Enabled
End Property

Public Property Let TableComboEnabled(ByVal blnNewValue As Boolean)
  With cboTables
    .Enabled = blnNewValue
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property


'Private Sub DrawControls()
'
'  Dim lngButtonLeft As Long
'  Dim lngHighestBottomButton As Long
'  Dim lngOffsetY As Long
'  Const lngGAPY = 100
'
'  lngOffsetY = cboTables.Top
'  cboTables.Visible = mblnTableComboVisible
'  If mblnTableComboVisible Then
'    lngOffsetY = lngOffsetY + cboTables.Height + lngGAPY
'    Call PopulateTables
'  End If
'
'  lngButtonLeft = Me.ScaleWidth - (cmdNew.Width + lngGAPY)
'
'  'List1.Top = lngOffsetY
'  List1.Move lngGAPY, lngOffsetY, lngButtonLeft - (lngGAPY * 2)
'
'
'  ' Display the 'new' command control as required.
'  With cmdNew
'    If (mlngOptions And edtAdd) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'  ' Display the 'edit' command control as required.
'  With cmdEdit
'    If (mlngOptions And edtEdit) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'  ' Display the 'copy' command control as required.
'  With cmdCopy
'    If (mlngOptions And edtCopy) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'  ' Display the 'delete' command control as required.
'  With cmdDelete
'    If (mlngOptions And edtDelete) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'  ' Display the 'print' command control as required.
'  With cmdPrint
'    If (mlngOptions And edtPrint) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'
'  ' Display the 'properties' command control as required.
'  With cmdProperties
'    If (mlngOptions And edtProperties) Then
'      .Visible = True
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    Else
'      .Visible = False
'    End If
'  End With
'
'
'  lngHighestBottomButton = lngOffsetY
'
'  ' Reset the lngOffsetY variable as the following command controls are positioned from the
'  ' bottom upwards.
'  lngOffsetY = List1.Top + List1.Height
'
'  If Not mblnHideDesc Then
'    'txtDesc.Top = lngOffsetY + lngGAPY
'    txtDesc.Move lngGAPY, lngOffsetY + lngGAPY, List1.Width
'    lngOffsetY = txtDesc.Top + txtDesc.Height
'  End If
'
'  ''--- The following lines of code are for the
'  ''--- Only show my definitions check box !
'  If Not mblnApplyDefAccess Then
'    chkOnlyMine.Top = lngOffsetY + lngGAPY
'    'Me.Height = (chkOnlyMine.Top + chkOnlyMine.Height) + lngGAPY + (Me.Height - Me.ScaleHeight)
'  Else
'    chkOnlyMine.Visible = False
'    'Me.Height = lngOffsetY + lngGAPY + (Me.Height - Me.ScaleHeight)
'  End If
'
'  lngOffsetY = lngOffsetY - cmdCancel.Height
'
'  ' Display the 'cancel' command control.
'  With cmdCancel
'    .Visible = True
'    '.Top = lngOffsetY
'    .Move lngButtonLeft, lngOffsetY
'    lngOffsetY = lngOffsetY - .Height - lngGAPY
'  End With
'
'  ' Display the 'none' command control as required.
'  With cmdNone
'    If (mlngOptions And edtDeselect) Then
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY - .Height - lngGAPY
'      .Visible = True
'    Else
'      .Visible = False
'    End If
'  End With
'
'  ' Display the 'select' command control as required.
'  With cmdSelect
'    If (mlngOptions And edtSelect) Then
'      '.Top = lngOffsetY
'      .Move lngButtonLeft, lngOffsetY
'      lngOffsetY = lngOffsetY - .Height - lngGAPY
'      .Visible = True
'      cmdCancel.Caption = "&Cancel"
'    Else
'      .Visible = False
'      cmdCancel.Caption = "&OK"
'    End If
'  End With
'
'  ' JPD 6/6/00 Check if any of the buttons are overlapped.
'  If lngHighestBottomButton > (lngOffsetY + cmdCancel.Height + lngGAPY) Then
'    lngOffsetY = lngHighestBottomButton
'
'    ' Display the 'select' command control as required.
'    With cmdSelect
'      If (mlngOptions And edtSelect) Then
'        .Top = lngOffsetY
'        lngOffsetY = lngOffsetY + .Height + lngGAPY
'      End If
'    End With
'
'    ' Display the 'none' command control as required.
'    With cmdNone
'      If (mlngOptions And edtDeselect) Then
'        .Top = lngOffsetY
'        lngOffsetY = lngOffsetY + .Height + lngGAPY
'      End If
'    End With
'
'    ' Display the 'cancel' command control.
'    With cmdCancel
'      .Top = lngOffsetY
'      lngOffsetY = lngOffsetY + .Height + lngGAPY
'    End With
'
'    ' Resize the form.
'    Me.Height = lngOffsetY + (Me.Height - Me.ScaleHeight)
'
'    ' Resize the listbox, and reposition the description box as required.
'    If Not mblnHideDesc Then
'      txtDesc.Top = lngOffsetY - txtDesc.Height - lngGAPY
'      List1.Height = txtDesc.Top - List1.Top - lngGAPY
'    Else
'      List1.Height = lngOffsetY - List1.Top
'    End If
'  End If
'
'  ' Enable/disable the command controls as required.
'  Refresh_Controls
'
'End Sub

Private Sub abDefSel_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  Select Case Tool.Name
    Case "New"
      cmdNew_Click
      
    Case "EditView"
      cmdEdit_Click
      
    Case "Copy"
      cmdCopy_Click
      
    Case "Delete"
      cmdDelete_Click
      
    Case "Print"
      cmdPrint_Click
      
    Case "Properties"
      cmdProperties_Click
      
    Case "Select"
      cmdSelect_Click
      
    Case "ID_None"
      cmdNone_Click
      
    ' JDM - 07/11/01 - Fault 3088 - Run click does nowt...
    Case "Run"
      cmdSelect_Click
      
  End Select
  
End Sub

Private Sub cboTables_Click()
  
  With cboTables
    If .ListIndex > -1 Then
      If mlngTableID <> .ItemData(.ListIndex) Then
        mlngTableID = .ItemData(.ListIndex)
        Call Populate_List
      End If
    End If
  End With
End Sub

Private Sub chkOnlyMine_Click()

  mintOnlyMine = chkOnlyMine.Value

  If Me.Visible = False Then
    Exit Sub
  End If

  If Not IsEmpty(List1.SelectedItem) And (Not List1.SelectedItem Is Nothing) Then
    mlngSelectedID = Val(List1.SelectedItem.Tag)
    SelectedText = List1.SelectedItem.Text
  Else
    mlngSelectedID = 0
  End If

  Call Populate_List

End Sub

Private Sub cmdCancel_Click()
  ' Cancel the selection form.

  Dim frmOutput As frmEventLog


  If mblnScheduledJobs Then
    If mutlUtilityType = utlBatchJob Then
      'If not batch logon then show event log with only selected jobs
      If Not gblnBatchJobsOnly And mstrEventLogIDs <> vbNullString Then
        Set frmOutput = New frmEventLog
        frmOutput.FilterIDs = mstrEventLogIDs
        frmOutput.Caption = msGeneralCaption & " Event Log"
        
        Me.Hide
        Screen.MousePointer = vbDefault
        frmOutput.Show vbModal
        Set frmOutput = Nothing
      End If
    End If
  End If
  
  lngAction = 0
  Unload Me

End Sub

Private Sub cmdDelete_Click()

'  'Check to see if this is somebody else's read only definition
'  With mrsRecords
'    .MoveFirst
'    .Find msIDField & " = " & CStr(List1.ItemData(List1.ListIndex))
'
'    If Trim$(.Fields("Username").Value) <> gsUserName And _
'             .Fields("Access").Value = "RO" Then
'      COAMsgBox "You do not have access to delete this definition", vbExclamation
'              'StrConv(Trim$(.Fields("Username").Value), vbProperCase)
'      Exit Sub
'    End If
'
'  End With
  
  Dim lngHighLightIndex As Long
  Dim lngSelectedID As Long
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  Dim objExpression As clsExprExpression
  
  lngHighLightIndex = List1.SelectedItem.Index
  lngSelectedID = Val(List1.SelectedItem.Tag)
  
  If CanStillSeeDefinition(lngSelectedID) = False Then
    Exit Sub
  End If
  
  
  'TM20011022 Fault 2946
  FromCopy = False

  lngAction = edtDelete
  
  'TM20010801 Fault 2617
  'If the expression type is Filter or Calculation then we need to check that the
  'expression should not be hidden and not owned by another user.
  If msTypeCode = "CALCULATIONS" Or msTypeCode = "FILTERS" Then
    
    If COAMsgBox("Delete this definition are you sure ?", vbQuestion + vbYesNo, "Delete " & msType) = vbYes Then
      If Not CheckForUseage(msType, lngSelectedID) Then
        Unload Me
      End If
    End If
    
  Else
    ' Ask for user confirmation to delete the utility definition
    If COAMsgBox("Delete this " & LCase(msType) & ", are you sure ?", vbQuestion + vbYesNo, "Delete " & msType) = vbYes Then
      If Not CheckForUseage(msType, lngSelectedID) Then
          
        Select Case msType
        Case "Batch Job"
          datGeneral.DeleteRecord "AsrSysBatchJobDetails", "BatchJobNameID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysBatchJobAccess", "ID", lngSelectedID
          
        Case "Calendar Reports"
          datGeneral.DeleteRecord "ASRSysCalendarReportEvents", "CalendarReportID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysCalendarReportOrder", "CalendarReportID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysCalendarReportAccess", "ID", lngSelectedID
        
        Case "Cross Tab"
          datGeneral.DeleteRecord "ASRSysCrossTabAccess", "ID", lngSelectedID
        
        Case "Custom Reports"
          datGeneral.DeleteRecord "ASRSysCustomReportAccess", "ID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysCustomReportsDetails", "CustomReportID", lngSelectedID
        
        Case "Data Transfer"
          datGeneral.DeleteRecord "ASRSysDataTransferAccess", "ID", lngSelectedID
  
        Case "Export"
          datGeneral.DeleteRecord "AsrSysExportDetails", "ExportID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysExportAccess", "ID", lngSelectedID
                    
        Case "Global Add", "Global Delete", "Global Update"
          datGeneral.DeleteRecord "ASRSysGlobalAccess", "ID", lngSelectedID
  
        Case "Record Profile"
          datGeneral.DeleteRecord "ASRSysRecordProfileDetails", "RecordProfileID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysRecordProfileTables", "RecordProfileID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysRecordProfileAccess", "ID", lngSelectedID
  
        Case "Import"
          datGeneral.DeleteRecord "ASRSysImportDetails", "ImportID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysImportAccess", "ID", lngSelectedID
        
          ' Also need to delete the file filter expression record (if one exists).
          sSQL = "SELECT filterID" & _
            " FROM ASRSysImportName" & _
            " WHERE ID = " & Trim(Str(lngSelectedID))

          Set rsTemp = datGeneral.GetRecords(sSQL)
          If rsTemp!FilterID > 0 Then
            ' Instantiate a new expression object.
            Set objExpression = New clsExprExpression

            With objExpression
              ' Initialise the expression object.
              If .Initialise(0, rsTemp!FilterID, giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
                .DeleteExpression False
              End If
            End With

            Set objExpression = Nothing
          End If

          rsTemp.Close
          Set rsTemp = Nothing
  
        Case "Match Report", "Career Progression", "Succession Planning"
          datGeneral.DeleteRecord "ASRSysMatchReportAccess", "ID", lngSelectedID
          
          'JPD 20040227 Fault 8160
          ' Also need to delete the file filter expression record (if one exists).
          sSQL = "SELECT requiredExprID as [exprID]," & _
            giEXPR_MATCHWHEREEXPRESSION & " AS [exprType]," & _
            giEXPRVALUE_LOGIC & " AS [exprReturnType]" & _
            " FROM ASRSysMatchReportTables" & _
            " WHERE matchReportID = " & Trim(Str(lngSelectedID)) & _
            "   AND requiredExprID > 0" & _
            " UNION" & _
            " SELECT preferredExprID as [exprID]," & _
            giEXPR_MATCHJOINEXPRESSION & " AS [exprType]," & _
            giEXPRVALUE_LOGIC & " AS [exprReturnType]" & _
            " FROM ASRSysMatchReportTables" & _
            " WHERE matchReportID = " & Trim(Str(lngSelectedID)) & _
            "   AND preferredExprID > 0" & _
            " UNION" & _
            " SELECT matchScoreExprID as [exprID]," & _
            giEXPR_MATCHSCOREEXPRESSION & " AS [exprType]," & _
            giEXPRVALUE_NUMERIC & " AS [exprReturnType]" & _
            " FROM ASRSysMatchReportTables" & _
            " WHERE matchReportID = " & Trim(Str(lngSelectedID)) & _
            "   AND matchScoreExprID > 0"

          Set rsTemp = datGeneral.GetRecords(sSQL)
          Do While Not rsTemp.EOF
            If rsTemp!ExprID > 0 Then
              ' Instantiate a new expression object.
              Set objExpression = New clsExprExpression
  
              With objExpression
                ' Initialise the expression object.
                If .Initialise(0, rsTemp!ExprID, rsTemp!exprType, rsTemp!exprReturnType) Then
                  .DeleteExpression False
                End If
              End With
  
              Set objExpression = Nothing
            End If

            rsTemp.MoveNext
          Loop
          
          rsTemp.Close
          Set rsTemp = Nothing
  
        Case "Mail Merge", "Envelopes & Labels"
          datGeneral.DeleteRecord "ASRSysMailMergeColumns", "MailMergeID", lngSelectedID
          datGeneral.DeleteRecord "ASRSysMailMergeAccess", "ID", lngSelectedID
  
        Case "Envelope & Label Template"
          datGeneral.DeleteRecord "ASRSysLabelTypes", "LabelTypeID", lngSelectedID
        
        End Select
  
        datGeneral.DeleteRecord msTableName, msIDField, lngSelectedID
        
        If mutlUtilityType <> -1 Then
          Call DeleteUtilAccessLog(mutlUtilityType, lngSelectedID)
        End If
        
        lngHighLightIndex = List1.SelectedItem.Index
        List1.ListItems.Remove lngHighLightIndex
        If List1.ListItems.Count > 0 Then
          Set List1.SelectedItem = List1.ListItems.Item(IIf(lngHighLightIndex < List1.ListItems.Count, lngHighLightIndex, List1.ListItems.Count))
        End If
  
        'Refresh_Controls
        'CheckListViewColWidth List1
        Populate_List
      End If
    End If
  End If
  
End Sub

Private Sub cmdCopy_Click()
  
  ' same as edit except set FromCopy flag
  'frmMain.Tag = List1.ListIndex
  lngAction = edtEdit
  GetSelected
  FromCopy = True

  If CanStillSeeDefinition(mlngSelectedID) Then
    Unload Me
  End If

End Sub

Private Sub cmdEdit_Click()
  ' Edit the selected item.
  'frmMain.Tag = List1.ListIndex
  lngAction = edtEdit
  GetSelected
  FromCopy = False

  If CanStillSeeDefinition(mlngSelectedID) Then
    Unload Me
  End If

End Sub

Private Sub cmdNew_Click()
  ' Create a new item.
  lngAction = edtAdd
  FromCopy = False
  Unload Me
    
End Sub

Private Sub cmdNone_Click()

  lngAction = edtDeselect
  Unload Me

End Sub

Private Sub cmdPrint_Click()

  ' Select the selected item.
  lngAction = edtPrint
  GetSelected
  
  If CanStillSeeDefinition(mlngSelectedID) Then
    Unload Me
  End If

End Sub

Private Sub cmdProperties_Click()

  'Dim varWhereUsed As Variant
  'Dim intCount As Integer
  'Dim strUsage As String
  
  On Error GoTo Prop_ERROR
  
  lngAction = edtProperties
  GetSelected

  If CanStillSeeDefinition(mlngSelectedID) = False Then
    Exit Sub
  End If
  
  ' RH Show the user we are doing something...checking for usage could take a while
  Screen.MousePointer = vbHourglass
  
  Load frmDefProp

  With frmDefProp

    .Caption = msSingularCaption & " Properties"
    .UtilName = SelectedText
    .PopulateUtil mutlUtilityType, mlngSelectedID

    'With .List1
    '  .Clear
    '   Call datGeneral.CheckForUseage(msType, mlngSelectedID, strUsage)
    '   If strUsage <> vbNullString Then
    '    varWhereUsed = Split(strUsage, vbCrLf)
    '    For intCount = LBound(varWhereUsed) To UBound(varWhereUsed)
    '      If Trim(varWhereUsed(intCount)) <> vbNullString Then
    '        .AddItem Trim(varWhereUsed(intCount))
    '      End If
    '    Next
    '  End If
    'End With

    .CheckForUseage msType, mlngSelectedID
    
    ' RH return the pointer to norma
    Screen.MousePointer = vbDefault
    
    'DynamicallyChangeHelpContextID
    .HelpContextID = DynamicallyChangeHelpContextID
    
    .Show vbModal
  End With
  
TidyUp:

  Unload frmDefProp
  Set frmDefProp = Nothing

  Exit Sub
  
Prop_ERROR:
  
  Screen.MousePointer = vbDefault
  COAMsgBox "Error retrieving properties for this definition." & vbCrLf & "Please contact support stating : " & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Properties"
  Resume TidyUp

End Sub

Private Sub cmdRefresh_Click()
  ' Refresh the list.
  lngAction = edtRefresh
  Unload Me

End Sub

Private Sub cmdSelect_Click()
  
  ' Select the selected item.
  lngAction = edtSelect
  If mblnScheduledJobs Then
    Select Case mutlUtilityType
    Case utlBatchJob
      RunSelectedJobs
      
    Case utlWorkflow
      ReadSelectedIDs
    End Select
  Else
    GetSelected
    If Not CanStillSeeDefinition(mlngSelectedID) Then
      Exit Sub
    End If
  End If

  Unload Me

End Sub



Private Sub Form_Activate()
   
  Screen.MousePointer = vbDefault

  If mblnBatchPrompt Then
    If List2.Visible And List2.Enabled Then
      List2.SetFocus
    End If
  Else
    If List1.Visible And List1.Enabled Then
      List1.SetFocus
    End If
  End If

  'MH20020227 Not Required?
  Refresh_Controls

End Sub


Private Sub Form_Initialize()
  mblnFirstLoad = True 'Read from settings
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim AIRhelp, ShellErrMessage As String
Dim CommandLine As String
Dim plngHelp As Long

Select Case KeyCode
  Case vbKeyF1
    AIRhelp = App.Path & "\HR Pro Data Manager Help\HR Pro Data Manager Help.exe"
    CommandLine = " -csh mapnumber " & CStr(Me.HelpContextID)
    
    plngHelp = ShellExecute(0&, vbNullString, AIRhelp, CommandLine, vbNullString, vbNormalNoFocus)
    
    If plngHelp <> 42 Then
      Select Case plngHelp
        Case 0: ShellErrMessage = "The operating system is out of memory or resources."
        Case 2: ShellErrMessage = "The specified file was not found"
        Case 3: ShellErrMessage = "The specified path was not found."
        Case 5: ShellErrMessage = "Windows 95 only: The operating system denied access to the specified file"
        Case 8: ShellErrMessage = "Windows 95 only: There was not enough memory to complete the operation."
        Case 10: ShellErrMessage = "Wrong Windows version"
        Case 11: ShellErrMessage = "The .EXE file is invalid (non-Win32 .EXE or error in .EXE image)."
        Case 12: ShellErrMessage = "Application was designed for a different operating system"
        Case 13: ShellErrMessage = "Application was designed for MS-DOS 4.0"
        Case 15: ShellErrMessage = "Attempt to load a real-mode program"
        Case 16: ShellErrMessage = "Attempt to load a second instance of an application with non-readonly data segments"
        Case 19: ShellErrMessage = "Attempt to load a compressed application file"
        Case 20: ShellErrMessage = "Dynamic-link library (DLL) file failure"
        Case 26: ShellErrMessage = "A sharing violation occurred."
        Case 27: ShellErrMessage = "The filename association is incomplete or invalid."
        Case 28: ShellErrMessage = "The DDE transaction could not be completed because the request timed out."
        Case 29: ShellErrMessage = "The DDE transaction failed."
        Case 30: ShellErrMessage = "The DDE transaction could not be completed because other DDE transactions were being processed."
        Case 31: ShellErrMessage = "There is no application associated with the given filename extension."
        Case 32: ShellErrMessage = "Windows 95 only: The specified dynamic-link library was not found."
        Case Else
          ShellErrMessage = "Undocumented API Error occurred"
      End Select
    
      If ASRDEVELOPMENT Then
        'More technical response for programmer
        COAMsgBox ShellErrMessage & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HR Pro Data Manager Help.exe'.", vbExclamation + vbOKOnly, "HR Pro Data Manager"
      Else
        'Less technical response for user
        COAMsgBox "Error whilst attempting to display help file." & vbCrLf & vbCrLf & "Please use windows explorer to view the file 'HR Pro Data Manager Help.exe'.", vbExclamation + vbOKOnly, "HR Pro Data Manager"
      End If
    End If
  Case vbKeyF5
    If (mutlUtilityType = utlWorkflow) And (mblnBatchPrompt) Then
      cmdRefresh_Click
    Else
      Populate_List
    End If
  Case vbKeyDelete
    If cmdDelete.Enabled Then cmdDelete_Click
End Select
'  If KeyCode = vbKeyF5 Then
'    If (mutlUtilityType = utlWorkflow) And (mblnBatchPrompt) Then
'      cmdRefresh_Click
'    Else
'      Populate_List
'    End If
'  ElseIf KeyCode = vbKeyDelete Then
'    If cmdDelete.Enabled Then cmdDelete_Click
'  End If

'THIS BIT NEEDS WORK
If plngHelp <> 42 Then
  KeyCode = 112
Else
  KeyCode = 0
End If
      
      

End Sub

Private Sub Form_Load()

  Hook Me.hWnd, Me.Width, Me.Height
  
  mblnLoading = False
  If mlngOptions = 0 Then
    mlngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtSelect + edtPrint + edtProperties
  End If
  
  If Not mblnFirstLoad Then
    Me.Height = mlngHeight
    Me.Width = mlngWidth
  End If

  'mstrEventLogIDs = vbNullString
  SizeControls

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    lngAction = 0
  End If
  If Not FromCopy Then
    GetSelected
  End If
  
  mlngHeight = Me.Height
  mlngWidth = Me.Width

  'RH 29/03/00 - To prevent the lockups of the toolbars after utility usage
  With frmMain.abMain
    .ResetHooks
    .Refresh
  End With
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  'DrawControls
  'ShowControls
  If Me.Visible = True Then
    SizeControls
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unhook Me.hWnd
End Sub

'Private Sub List1_Click()
'  Refresh_Controls
'End Sub

Private Sub List1_ItemClick(ByVal Item As ComctlLib.ListItem)
  Refresh_Controls
End Sub

Private Sub List2_Click()
  Screen.MousePointer = vbHourglass
  Refresh_Controls
  Screen.MousePointer = vbNormal
End Sub

Private Sub List1_DblClick()
  
  ' RH 25/09/00 - BUG 1001
  If Not List1.SelectedItem Is Nothing Then
  
    If (mlngOptions And edtEdit) And cmdSelect.Visible = False Then
      If mfEnableView Then
        lngAction = edtEdit
        GetSelected
        FromCopy = False
        If CanStillSeeDefinition(mlngSelectedID) Then
          Unload Me
        End If
      End If
    ElseIf (mlngOptions And edtSelect) Then
      If mfEnableRun Then
        ' If we are trying to RUN the item, ask confirmation
        If mblnCaptionIsRun Then
          'If COAMsgBox("Are you sure you want to run the '" & List1.SelectedItem.Text & "' " & Me.Caption & " ?", vbYesNo + vbQuestion, "Confirmation...") = vbNo Then
          'NHRD25082004 Fault 7802
          If COAMsgBox("Are you sure you want to run the '" & List1.SelectedItem.Text & "' " & msSingularCaption & " ?", vbYesNo + vbQuestion, "Confirmation...") = vbNo Then
            Exit Sub
          End If
        End If
        ' Select the selected item.
        lngAction = edtSelect
        GetSelected
        If CanStillSeeDefinition(mlngSelectedID) Then
          Unload Me
        End If
      End If
    End If

  End If
  
End Sub

Private Sub List1_GotFocus()
  Refresh_Controls
End Sub

Private Sub Display_Button(Button As VB.CommandButton, ByVal BtnOpt As Long, ByVal x As Long, ByRef y As Long)
  If (Me.Options And BtnOpt) Then
    Button.Move x, y
    Button.Visible = True
    y = y + cmdNew.Height + ((UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY) * 1.5)
  Else
    Button.Visible = False
  End If
  
End Sub



Public Sub Refresh_Controls()

  Dim lngSelected As Long
  Dim sCurrentUserAccess As String
  Dim fNewValue As Boolean
  Dim iCount As Integer
  Dim lngTempIndex As Long
  
  If mblnLoading Then
    Exit Sub
  End If
  
  On Error GoTo LocalErr
  
  lngSelected = 0
  If mblnBatchPrompt Then
    If List2.ListCount > 0 And List2.ListIndex <> -1 Then
      lngSelected = List2.ItemData(List2.ListIndex)
    End If
  Else
    If List1.ListItems.Count > 0 And Not IsEmpty(List1.SelectedItem) Then
      lngSelected = List1.SelectedItem.Tag
    End If
  End If
  
  ' Check if the user has selected the all columns.
  'Only want this part to run if "<All>" is at the top of
  'the list as in Schecule Batch Jobs
  'JPD 20070814 Fault 12430
  If lngSelected = 0 And List2.List(0) = "<All>" Then
    fNewValue = List2.Selected(0)
    UI.LockWindow Me.hWnd
    For iCount = 1 To List2.ListCount - 1
      ' Update all rows in the listbox.
      List2.Selected(iCount) = fNewValue
      'lngColumnID = List2.ItemData(iCount)
    Next iCount
    List2.ListIndex = 0
    UI.UnlockWindow
    cmdProperties.Enabled = False
    cmdSelect.Enabled = (List2.SelCount > 0)
  Else
    
    'MH20050107 Fault 9571
    'Update "<All>" option...
    If mblnBatchPrompt And List2.ListCount > 0 Then
      lngTempIndex = List2.ListIndex
      List2.Selected(0) = (List2.SelCount = List2.ListCount - IIf(List2.Selected(0), 0, 1))
      'NHRD270905 Fault 9612
     ' List2.ListIndex = lngTempIndex
    End If
    
    'NHRD15112004 Fault 9351
    cmdProperties.Enabled = (List2.Text <> "<All>") And (List2.ListCount > 0)
  End If
  
  ' Enable/disable controls as required.
  If lngSelected > 0 Then
    With mrsRecords
      If Not (.BOF And .EOF) Then
        .MoveFirst
        .Find msIDField & " = " & CStr(lngSelected)  'CStr(List1.ItemData(List1.ListIndex))
      
        If Not mblnHideDesc Then
          txtDesc.Text = IIf(IsNull(.Fields("Description").Value), vbNullString, .Fields("Description").Value)
        End If
  
        If mblnApplyDefAccess Then
          If OldAccessUtility(mutlUtilityType) Then
            sCurrentUserAccess = .Fields("Access").Value
          Else
            sCurrentUserAccess = CurrentUserAccess(mutlUtilityType, lngSelected)
          End If
            
          mblnHiddenDef = (sCurrentUserAccess = ACCESS_HIDDEN)
          mblnReadOnlyAccess = (sCurrentUserAccess = ACCESS_READONLY And _
            LCase(Trim$(.Fields("Username").Value)) <> LCase(gsUserName)) And _
            (Not gfCurrentUserIsSysSecMgr)
        End If
      End If
    End With

    cmdNew.Enabled = (cmdNew.Visible And mfEnableNew)
    cmdCopy.Enabled = (cmdCopy.Visible And mfEnableNew)
    cmdEdit.Enabled = (cmdEdit.Visible And mfEnableView)
    cmdDelete.Enabled = (cmdDelete.Visible And mfEnableDelete And Not (mblnReadOnlyAccess))
    If mblnBatchPrompt Then
      cmdSelect.Enabled = (mfEnableRun And List2.SelCount > 0)
    Else
      cmdSelect.Enabled = (cmdSelect.Visible And mfEnableRun)
    End If
    cmdPrint.Enabled = (cmdPrint.Visible And mfEnableView)
    cmdProperties.Enabled = (cmdProperties.Visible And mfEnableView)
  
    If Not mfEnableEdit Or mblnReadOnlyAccess Then
      If mfEnableView Then
        cmdEdit.Caption = "&View..."
        cmdEdit.Enabled = True
      Else
        cmdEdit.Enabled = False
      End If
    Else
        cmdEdit.Caption = "&Edit..."
    End If

    Me.abDefSel.Bands("bndDefSel").Tools("EditView").Caption = cmdEdit.Caption
    Me.abDefSel.Bands("bndDefSel").Tools("Select").Visible = (cmdSelect.Visible And Not mblnCaptionIsRun)
    Me.abDefSel.Bands("bndDefSel").Tools("Run").Visible = (cmdSelect.Visible And mblnCaptionIsRun)
    Me.abDefSel.Bands("bndDefSel").Tools("ID_None").Visible = (cmdNone.Visible)

  Else
    'cmdNew.Enabled = mfEnableNew
    'NHRD15122006 Fault
    cmdNew.Enabled = (cmdNew.Visible And mfEnableNew)
    
    cmdCopy.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    'cmdSelect.Enabled = List2.Text = "<All>"
    If mblnBatchPrompt Then
      cmdSelect.Enabled = (mfEnableRun And List2.SelCount > 0)
    Else
      cmdSelect.Enabled = False
    End If
    cmdPrint.Enabled = False
    
    'NHRD15112004 Fault 9351
    'cmdProperties.Enabled = List2.Text = "<All>"
    cmdProperties.Enabled = (List2.Text <> "<All>") And (List2.ListCount > 0)
    
    cmdEdit.Caption = "&Edit..."
    If mblnHideDesc = False Then
      txtDesc.Text = vbNullString
    End If
  End If
  
  If cmdSelect.Visible Then
    cmdSelect.Default = cmdEdit.Enabled
  ElseIf cmdEdit.Visible Then
    cmdSelect.Default = False
    cmdEdit.Default = cmdEdit.Enabled
  Else
    cmdSelect.Default = False
    cmdNew.Default = mfEnableNew
  End If
  
'If fOK Then
'  ' Check for all columns now being selected or deselected
'  fAllColumns = True
'  For iCount = 1 To List2.ListCount - 1
'    If Not List2.Selected(iCount) Then
'      fAllColumns = False
'      Exit For
'    End If
'  Next iCount
'  ' Update the 'All' row in the listbox.
'  List2.Selected(0) = fAllColumns
'End If

  
  If Not mblnBatchPrompt Then
    If List1.ListItems.Count > 0 Then
      List1.ListItems(List1.SelectedItem.Index).Selected = True   'This highlights the current item!!!!!
      List1.Refresh
    End If
  End If

Exit Sub

LocalErr:
  UI.UnlockWindow
  COAMsgBox Err.Description, vbCritical

End Sub

Public Property Get SelectedText() As String

    SelectedText = msSelectedText

End Property

Public Property Let SelectedText(ByVal sText As String)

    msSelectedText = sText

End Property

Private Sub GetSelected()
  
  If lngAction > 0 And lngAction <> edtAdd Then

' RH BUG 958 - If listbox is empty and none selected then crashed out

'    If Not IsEmpty(List1.SelectedItem) Then
    mlngSelectedID = 0
    If mblnScheduledJobs Then
      If List2.ListIndex >= 0 Then
        mlngSelectedID = List2.ItemData(List2.ListIndex)
        SelectedText = List2.List(List2.ListIndex)
      End If
    ElseIf Not IsEmpty(List1.SelectedItem) And (Not List1.SelectedItem Is Nothing) Then
      mlngSelectedID = Val(List1.SelectedItem.Tag)
      SelectedText = List1.SelectedItem.Text
    End If
  End If

End Sub

'Public Sub RefreshListBox()
'
'    Dim rsList As New Recordset
'
'
'    With List1
'        .Clear
'        Set rsList = datGeneral.GetRecords(msRecordSource)
'        Do While Not rsList.EOF
'            .AddItem rsList.Fields(msFieldName)
'            .ItemData(.NewIndex) = rsList.Fields(msIDField)
'            rsList.MoveNext
'        Loop
'        rsList.Close
'        Set rsList = Nothing
'        If .ListCount > 0 Then
'            .ListIndex = .ListCount - 1
'        End If
'    End With
'
'End Sub

Public Property Let EnableRun(ByVal bEnable As Boolean)
  ' Change the caption on the cmdSelect control as appropriate.
  cmdSelect.Caption = IIf(bEnable, "&Run", "&Select")
  'Me.abDefSel.Bands("bndDefSel").Tools("Select").Caption = cmdSelect.Caption
  mblnCaptionIsRun = bEnable
End Property

Public Property Get FromCopy() As Boolean

    FromCopy = mbFromCopy

End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)

    mbFromCopy = bCopy

End Property


Private Function CheckForUseage(sDefType As String, lItemID As Long) As Boolean
  ' Check if the given record is used.
  Dim sMsg As String
  Dim intCount As Integer

  Load frmDefProp
  
  With frmDefProp
    
    If .CheckForUseage(sDefType, lItemID) Then
        
      With .List1
        sMsg = vbNullString
        For intCount = 0 To .ListCount - 1
          sMsg = sMsg & .List(intCount) & vbCrLf
        Next
    
        'If not an error message then add wording
        If Left$(sMsg, 1) <> "<" Then
          sMsg = "currently being used in:" & vbCrLf & vbCrLf & sMsg
        End If
  
        COAMsgBox "Unable to delete this " & LCase(sDefType) & ", " & sMsg, vbExclamation, Me.Caption
        CheckForUseage = True
      End With
      
    End If
  
  End With
  
  Unload frmDefProp
  Set frmDefProp = Nothing

End Function


Private Sub ApplySystemPermissions()
  ' Enable/disable buttons according to the configured System Permissions.
  ' Initialise the enabled flags.
  
  mfEnableNew = True
  mfEnableView = True
  mfEnableEdit = True
  mfEnableDelete = True
  mfEnableRun = True
  
  If mblnApplySystemPermissions Then
    mfEnableNew = datGeneral.SystemPermission(msTypeCode, "NEW")
    mfEnableDelete = datGeneral.SystemPermission(msTypeCode, "DELETE")
    mfEnableEdit = datGeneral.SystemPermission(msTypeCode, "EDIT")
    mfEnableView = datGeneral.SystemPermission(msTypeCode, "VIEW")

    'If not edit but still have view then change the caption of command button
    If mfEnableEdit = False Then
      If mfEnableView = True Then
        cmdEdit.Caption = "&View..."
        Me.abDefSel.Bands("bndDefSel").Tools("EditView").Caption = cmdEdit.Caption
      End If
    End If
    
    If mblnCaptionIsRun Then
      'JPD 20060922 Fault 11492
      If (mutlUtilityType <> utlWorkflow) Or (Not mblnBatchPrompt) Then
        mfEnableRun = datGeneral.SystemPermission(msTypeCode, "RUN")
      End If
    End If

  End If

End Sub


Private Function Populate_List() As Boolean
Dim fAllColumns As Boolean

  'MH20000302 - Changed this sub from public to private
  
  'MH20000807 - Rather than sort the listview do the sort in the SQL so
  '             that you will always be able to see the selected item
  '             when the list is first shown (Fault 725)
  
  Dim strSQL As String
  'Dim intCount As Integer
  Dim objListItem As ListItem
  Dim objBatchJob As clsBatchJobRUN
  Dim sDescription As String
  Dim lngMax As Long
  Dim lngList2Max As Long
  Dim lngLen As Long
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  Const TICKBOXWIDTH = 20
  
  ' Populate the selection listbox with the information defined in the given parameters.
  On Error GoTo ErrorTrap
  
  If mblnLoading Then
    Exit Function
  End If
  mblnLoading = True
  
  ' Get rid of the icon off the form
  RemoveIcon Me
    
  strSQL = msRecordSource

  If mblnApplyDefAccess Then
    If OldAccessUtility(mutlUtilityType) Then
      If gfCurrentUserIsSysSecMgr Then
        If chkOnlyMine.Value = vbChecked Then
          strSQL = strSQL & _
            IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
            "(" & msTableName & ".userName = '" & datGeneral.UserNameForSQL & "'" & ")"
        End If
      Else
        strSQL = strSQL & _
          IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
          "(" & msTableName & ".userName = '" & datGeneral.UserNameForSQL & "'"
          
        If Not (chkOnlyMine.Value = vbChecked) Then
          strSQL = strSQL & _
            " OR " & msTableName & ".access <> '" & ACCESS_HIDDEN & "'"
        End If
        
        strSQL = strSQL & _
          ")"
      End If
    Else
      If gfCurrentUserIsSysSecMgr Then
        'MH20030521 Fault 5726
        'If chkOnlyMine Then
        If chkOnlyMine.Value = vbChecked Then
          strSQL = strSQL & _
            IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
            "(" & msTableName & ".userName = '" & datGeneral.UserNameForSQL & "'" & ")"
        End If
      Else
        strSQL = strSQL & _
          IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
          "(" & msTableName & ".userName = '" & datGeneral.UserNameForSQL & "'"
        
        'MH20030521 Fault 5726
        'If Not chkOnlyMine Then
        If Not (chkOnlyMine.Value = vbChecked) Then
          strSQL = strSQL & _
            " OR " & msAccessTableName & ".access <> '" & ACCESS_HIDDEN & "'"
        End If
          
        strSQL = strSQL & _
          ")"
      End If
    End If
  End If

  If mlngTableID > 0 Then
    'Changed as there are two tableid columns in match reports...
    'strSQL = strSQL & _
      IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
      msTableIDColumnName & " = " & CStr(mlngTableID)
    strSQL = strSQL & _
      IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
      CStr(mlngTableID) & " IN (" & msTableName & "." & msTableIDColumnName & ")"
  End If

  UI.LockWindow Me.hWnd

  strSQL = strSQL & " ORDER BY " & msTableName & ".name"
  Set mrsRecords = datGeneral.GetReadOnlyRecords(strSQL)

  lngMax = 1000
  lngList2Max = 0
  
  If mblnScheduledJobs Then
    List2.Clear
  Else
    List1.ListItems.Clear
  End If
  
  With mrsRecords
    If Not (.EOF And .BOF) Then
      .MoveFirst
      
      If mblnScheduledJobs Then
        Select Case mutlUtilityType
        Case utlBatchJob
          Set objBatchJob = New clsBatchJobRUN
          'List2.Clear
          Do While Not .EOF
            If objBatchJob.DoesUserHavePermissionForAllJobs(.Fields(msIDField)) Then
              If objBatchJob.CheckBatchNeedsRunning2(.Fields(msIDField), .Fields(msFieldName).Value) = vbNullString Then
                List2.AddItem .Fields(msFieldName).Value
                List2.ItemData(List2.NewIndex) = .Fields(msIDField).Value
                List2.Selected(List2.NewIndex) = True
              
                If lngList2Max < TextWidth(.Fields(msFieldName).Value) Then
                  lngList2Max = TextWidth(.Fields(msFieldName).Value)
                End If
              End If
            End If
            .MoveNext
          Loop
          Set objBatchJob = Nothing
        
        Case utlWorkflow
          Do While Not .EOF
            ' Get the instance step description (if one exists)
            Set cmADO = New ADODB.Command
            With cmADO
              .CommandText = "spASRWorkflowStepDescription"
              .CommandType = adCmdStoredProc
              .CommandTimeout = 0
              Set .ActiveConnection = gADOCon

              Set pmADO = .CreateParameter("InstanceStepID", adInteger, adParamInput)
              .Parameters.Append pmADO
              pmADO.Value = mrsRecords.Fields(msIDField).Value

              Set pmADO = .CreateParameter("Description", adVarChar, adParamOutput, VARCHAR_MAX_Size)
              .Parameters.Append pmADO
            
              Set pmADO = Nothing
            
              .Execute
  
              sDescription = .Parameters("Description").Value
            End With
            Set cmADO = Nothing

            If Len(Trim(sDescription)) = 0 Then
              sDescription = .Fields(msFieldName).Value
            End If
            List2.AddItem sDescription
            List2.ItemData(List2.NewIndex) = .Fields(msIDField).Value
            List2.Selected(List2.NewIndex) = True
            
            If lngList2Max < TextWidth(sDescription) Then
              lngList2Max = TextWidth(sDescription)
            End If
            
            .MoveNext
          Loop
        End Select
      Else
        'List1.ListItems.Clear
        Do While Not .EOF
          Set objListItem = List1.ListItems.Add(, , RemoveUnderScores(.Fields(msFieldName)))
          objListItem.Tag = .Fields(msIDField)
          lngLen = Me.TextWidth(objListItem.Text)
          If lngMax < lngLen Then
            lngMax = lngLen
          End If

          If .Fields(msIDField) = mlngSelectedID Then
            Set List1.SelectedItem = objListItem
          End If

          .MoveNext
        Loop
      End If
    End If
  End With
  
  If List2.ListCount > 0 Then
    ' Add the columns to the grid.
    fAllColumns = True
  End If
  
  fAllColumns = fAllColumns And (List2.ListCount > 0)
  ' Add the 'all columns' column.
  If List2.ListCount > 0 Then
    List2.AddItem "<All>", 0
    List2.ItemData(List2.NewIndex) = 0
    List2.Selected(List2.NewIndex) = fAllColumns
  
    If lngList2Max < TextWidth("<All>") Then
      lngList2Max = TextWidth("<All>")
    End If
  End If
  
  ' See if all the screens are all selected.
  List2.Enabled = (List2.ListCount > 1)
  
  ' Select the first item.
  If List2.Enabled Then
    List2.ListIndex = 0
  End If

  If List1.ListItems.Count > 0 Then
    If IsEmpty(List1.SelectedItem) Then
      Set List1.SelectedItem = List1.ListItems(1)
    End If
  End If
  
  If List2.ListCount > 0 Then
    List2.ListIndex = 0

    If ScaleMode = vbTwips Then
      lngList2Max = lngList2Max / Screen.TwipsPerPixelX  ' if twips change to pixels
    End If
    SendMessageLong List2.hWnd, LB_SETHORIZONTALEXTENT, lngList2Max + TICKBOXWIDTH, 0
  End If
  
  lngMax = lngMax + 60
  List1.ColumnHeaders(1).Width = lngMax
  List1.Refresh

  ApplySystemPermissions

  Populate_List = True
  
Exit_Populate_List:
'  Set mrsRecords = Nothing
  
  'List1.Refresh
  
  mblnLoading = False
  
  'MH20020227
  'Refresh_Controls
  If Me.Visible = True Then
    Refresh_Controls
  End If
  'CheckListViewColWidth List1

  UI.UnlockWindow
  
  Exit Function
  
ErrorTrap:
  Populate_List = False
  COAMsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  If ASRDEVELOPMENT Then
    Stop
  End If
  Err = False
  
  'This resume next causes an infinite loop!
  'Resume Next
  'Resume Exit_Populate_List
  
End Function



Public Function ShowOrders(strSQL As String, lngOrderID As Long) As Boolean
  
  mblnApplyDefAccess = False
  mlngSelectedID = lngOrderID
  
  msRecordSource = strSQL
  msType = "Order"
  msFieldName = "Name"
  msTableIDColumnName = "TableID"
  msTableName = "ASRSysOrders"
  'msIDField = "ASRSysOrders.OrderID"
  msIDField = "OrderID"
  msTypeCode = "ORDER"
  mutlUtilityType = utlOrder
  
  'NHRD04092003 Fault 6273, 5911
  msTypeCode = "ORDERS"
  mblnApplySystemPermissions = Not gfCurrentUserIsSysSecMgr
  
  msGeneralCaption = "Orders"
  msSingularCaption = "Order"
  
  Me.Caption = msGeneralCaption
  
  'Call DrawControls
  ShowControls
  'SizeControls
  ShowOrders = Populate_List

End Function

Private Sub PopulateTables()

  Dim lngTableID As Long

  lngTableID = mlngTableID
  LoadTableCombo2 cboTables

  If lngTableID > 0 Then
    mlngTableID = lngTableID
    SetComboItem cboTables, mlngTableID
  End If
  
'  ' Populate the Tables combo.
'  Dim iIndex As Integer
'  Dim sSQL As String
'  Dim rsTables As ADODB.Recordset
'
'
'  cboTables.Clear
'  iIndex = 0
'
'  sSQL = "SELECT tableName, tableID" & _
'         " FROM ASRSysTables" & _
'         " ORDER BY tableName"
'  Set rsTables = datGeneral.GetRecords(sSQL)
'  With rsTables
'    Do While Not .EOF
'      cboTables.AddItem !TableName
'      cboTables.ItemData(cboTables.NewIndex) = !TableID
'      If !TableID = mlngTableID Then
'        cboTables.ListIndex = cboTables.NewIndex
'      End If
'      .MoveNext
'    Loop
'
'    .Close
'
'    Set rsTables = Nothing
'
'    If cboTables.ListCount < 1 Then
'      TableComboEnabled = False
'    ElseIf cboTables.ListIndex < 0 Then
'      cboTables.ListIndex = 0
'    End If
'
'  End With

End Sub


Private Function CanStillSeeDefinition(lngDefID As Long) As Boolean

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim sCurrentUserAccess As String
  
  'MH20001013 Fault 1055
  'Need to include table name otherwise get Ambiguous column name message !
  If InStr(msRecordSource, " WHERE ") = 0 Then
    strSQL = msRecordSource & " WHERE " & msTableName & "." & msIDField & " = " & CStr(lngDefID)
  Else
    strSQL = msRecordSource & " AND " & msTableName & "." & msIDField & " = " & CStr(lngDefID)
  End If
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  CanStillSeeDefinition = True

  With rsTemp
    If .BOF And .EOF Then
      COAMsgBox "This definition has been deleted by another user", vbExclamation, Me.Caption
      Call Populate_List
      CanStillSeeDefinition = False
      'Exit Sub
    
    ElseIf mblnApplyDefAccess Then
    
      If LCase(Trim$(.Fields("Username").Value)) <> LCase(gsUserName) Then
      
        If Not gfCurrentUserIsSysSecMgr Then
          If OldAccessUtility(mutlUtilityType) Then
            sCurrentUserAccess = .Fields("Access").Value
          Else
            sCurrentUserAccess = CurrentUserAccess(mutlUtilityType, lngDefID)
          End If
          
          If sCurrentUserAccess = ACCESS_HIDDEN Then
            COAMsgBox "This definition has been made hidden by another user", vbExclamation, Me.Caption
            Call Populate_List
            CanStillSeeDefinition = False
          ElseIf sCurrentUserAccess = ACCESS_READONLY And Not mblnReadOnlyAccess Then
            COAMsgBox "This definition is now read only", vbInformation, Me.Caption
            mblnReadOnlyAccess = True
            Call CanStillSeeDefinition(lngDefID)  'Check again after COAMsgBox
      
          ElseIf sCurrentUserAccess = ACCESS_READWRITE And mblnReadOnlyAccess Then
            COAMsgBox "This definition is now read write", vbInformation, Me.Caption
            mblnReadOnlyAccess = False
            Call CanStillSeeDefinition(lngDefID)  'Check again after COAMsgBox
          End If
        End If
      End If
    End If
  End With

End Function


'MH20020527 Fault 3546
'Now try and "CheckListViewColWidth" during Populate_List
'as this should help performance (i.e. not loop though list again!)
'
'Private Sub CheckListViewColWidth(lstvw As ListView)
'
'  Dim objItem As ListItem
'  Dim lngMax As Long
'  Dim lngLen As Long
'
'  lngMax = 0
'
'  For Each objItem In lstvw.ListItems
'
'    lngLen = Me.TextWidth(objItem.Text)
'    If lngMax < lngLen Then
'      lngMax = lngLen
'    End If
'
'  Next objItem
'
'  lngMax = lngMax + 60
'  lstvw.ColumnHeaders(1).Width = lngMax
'  lstvw.Refresh
'
'End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  If Button = vbRightButton Then
  
    With Me.abDefSel.Bands("bndDefSel")

      ' Enable/disable the required tools.
      .Tools("New").Enabled = Me.cmdNew.Enabled
      .Tools("EditView").Enabled = Me.cmdEdit.Enabled
      .Tools("Copy").Enabled = Me.cmdCopy.Enabled
      .Tools("Delete").Enabled = Me.cmdDelete.Enabled
      .Tools("Print").Enabled = Me.cmdPrint.Enabled
      .Tools("Properties").Enabled = Me.cmdProperties.Enabled
      .Tools("Select").Enabled = cmdSelect.Enabled
      .Tools("Run").Enabled = cmdSelect.Enabled
      .Tools("ID_None").Enabled = cmdNone.Enabled

    End With

    abDefSel.Bands("bndDefSel").TrackPopup -1, -1
    
  End If
  
End Sub


Private Sub ShowControls()
  'SizeControls

  Dim lngOffset As Long
  Const lngGap = 100

  On Error GoTo LocalErr

  lngOffset = 0
  
  UI.LockWindow Me.hWnd
  
  'cmdNew (fraTopButtons)
  With cmdNew
    If (mlngOptions And edtAdd) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With
  
  'cmdEdit (fraTopButtons)
  With cmdEdit
    If (mlngOptions And edtEdit) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With
  
  'cmdCopy (fraTopButtons)
  With cmdCopy
    If (mlngOptions And edtCopy) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With
  
  'cmdDelete (fraTopButtons)
  With cmdDelete
    If (mlngOptions And edtDelete) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With
  
  'cmdPrint (fraTopButtons)
  With cmdPrint
    If (mlngOptions And edtPrint) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With
  
  'cmdProperties (fraTopButtons)
  With cmdProperties
    If (mlngOptions And edtProperties) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With

  'cmdRefresh (fraTopButtons)
  With cmdRefresh
    If (mlngOptions And edtRefresh) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With

  fraTopButtons.Height = lngOffset
  
  
  lngOffset = 0
  
  'cmdSelect (fraBottomsButtons)
  With cmdSelect
    If (mlngOptions And edtSelect) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
      cmdCancel.Caption = "&Cancel"
    Else
      .Visible = False
      cmdCancel.Caption = "&OK"
    End If
  End With

  'cmdNone (fraBottomsButtons)
  With cmdNone
    'TM20011217 Fault 3250 - Now using the mblnApplyDefAccess boolean to show the "None" button or not.
    'TM20020520 Fault 3358
    'If Not mblnApplyDefAccess Or (mlngOptions And edtDeselect) Then
    If (mlngOptions And edtDeselect) Then
      .Visible = True
      .Top = lngOffset
      lngOffset = lngOffset + .Height + lngGap
    Else
      .Visible = False
    End If
  End With

  'cmdCancel (fraBottomsButtons) ALWAYS VISIBLE
  With cmdCancel
    .Visible = True
    .Top = lngOffset
    lngOffset = lngOffset + .Height '+ lngGAP
  End With
  
  fraBottomButtons.Height = lngOffset + 10
  
  'MH20020227 Not required?
  '' Enable/disable the command controls as required.
  'Refresh_Controls
  
  
  lblTables.Visible = mblnTableComboVisible
  cboTables.Visible = mblnTableComboVisible
  If mblnTableComboVisible Then
    Call PopulateTables
  End If
  
  txtDesc.Visible = Not mblnHideDesc
  chkOnlyMine.Visible = mblnApplyDefAccess
  chkOnlyMine.Caption = "On&ly show definitions where owner is '" & _
                        StrConv(gsUserName, vbProperCase) & "'"

  'Make frame background same colour as form
  fraMain.BackColor = Me.BackColor
  fraTopButtons.BackColor = Me.BackColor
  fraBottomButtons.BackColor = Me.BackColor

Exit Sub

LocalErr:
  'Only unlock if there is an error otherwise populatelist will unlock window
  UI.UnlockWindow

End Sub


Private Sub SizeControls()

  Dim lngListTop As Long
  Dim lngOffset As Long
  Const lngGap = 100
  Dim blnCheckBoxVisible As Boolean

  'If Me.Width < 4000 Then
    'Me.Width = 4000
    'Exit Sub  'Resize will be called again anyway
  'End If
  
  'If Me.Height < 3000 Then
    'Me.Height = 3000
    'Exit Sub  'Resize will be called again anyway
  'End If
    
  lngOffset = Me.ScaleHeight - (lngGap * 2)
  
  'chkOnlyMine (Outside of frame)
  blnCheckBoxVisible = mblnApplyDefAccess
  If blnCheckBoxVisible Then
    chkOnlyMine.Move fraMain.Left, Me.ScaleHeight - (chkOnlyMine.Height + lngGap), Me.ScaleWidth
    lngOffset = lngOffset - (chkOnlyMine.Height + lngGap)
  End If


  'Move Frames
  fraMain.Move lngGap, lngGap, Me.ScaleWidth - (fraTopButtons.Width + (lngGap * 3)), lngOffset
  fraTopButtons.Move fraMain.Left + fraMain.Width + lngGap, lngGap
  fraBottomButtons.Move fraTopButtons.Left, (fraMain.Top + fraMain.Height) - fraBottomButtons.Height


  'Check for button overlap...
'  If fraTopButtons.Top + fraTopButtons.Height > fraBottomButtons.Top Then
'    fraBottomButtons.Top = fraTopButtons.Top + fraTopButtons.Height
'    Me.Height = (Me.Height - Me.ScaleHeight) + _
'                fraTopButtons.Height + fraBottomButtons.Height + _
'                IIf(blnCheckBoxVisible, chkOnlyMine.Height + lngGap, 0) + (lngGap * 2)
'    Exit Sub  'Resize will be called again anyway
'  End If
     
  'txtDesc (fraMain)
  If Not mblnHideDesc Then
    lngOffset = fraMain.Height - (txtDesc.Height)
    txtDesc.Move 0, lngOffset, fraMain.Width
    lngOffset = lngOffset - lngGap
  Else
    lngOffset = fraMain.Height
  End If

  'cboTables (fraMain)
  lngListTop = 0
  If mblnTableComboVisible Then
    lblTables.Move 0, 60
    cboTables.Move lblTables.Width + lngGap, 0, fraMain.Width - (lblTables.Width + lngGap)
    lngOffset = lngOffset - (cboTables.Height + lngGap)
    lngListTop = cboTables.Height + lngGap
  End If

  'List1 (fraMain)
  List1.Move 0, lngListTop, fraMain.Width + 20, lngOffset
  List2.Move 0, lngListTop, fraMain.Width + 20, lngOffset

End Sub


Public Property Let BatchPrompt(ByVal blnNewValue As Variant)
  mblnBatchPrompt = blnNewValue
  List1.Visible = Not mblnBatchPrompt
  List2.Visible = mblnBatchPrompt
End Property


Public Function RunSelectedJobs() As Boolean

  Dim objBatchJob As clsBatchJobRUN
  Dim blnCancelled As Boolean
  Dim strError As String
  
  Dim strNotes As String
  Dim lngCount As Long
  Dim lngID As Long

  Set objBatchJob = New clsBatchJobRUN
  
  Screen.MousePointer = vbHourglass
  
  For lngCount = 1 To List2.ListCount - 1

    If List2.Selected(lngCount) Then

      lngID = List2.ItemData(lngCount)

      If objBatchJob.DoesUserHavePermissionForAllJobs(lngID) = False Then
        COAMsgBox "The job '" & List2.List(lngCount) & "' will not be run as you do not have permission to run all of the jobs in this batch.", vbInformation, Me.Caption
        List2.Selected(lngCount) = False

      Else
        strError = objBatchJob.CheckBatchNeedsRunning2(lngID, List2.List(lngCount))
        If strError <> vbNullString Then
          COAMsgBox strError, vbInformation, Me.Caption
          List2.Selected(lngCount) = False
        Else
          If objBatchJob.LockJob(lngID, False) = False Then
            COAMsgBox "The job '" & List2.List(lngCount) & "' will not be run as it is already being run by another user.", vbInformation, Me.Caption
            List2.Selected(lngCount) = False
          End If
        End If

      End If

    End If
  Next
  
  Me.Hide

  'NHRD17112004 Fault 9351 Start count from 1 instead of 0
  'For lngCount = 0 To List2.ListCount - 1
  For lngCount = 1 To List2.ListCount - 1
    If List2.Selected(lngCount) Then

      If blnCancelled Then
        objBatchJob.LockJob List2.ItemData(lngCount), True
      
      Else
        strNotes = objBatchJob.RunBatchJob(List2.ItemData(lngCount), List2.List(lngCount), 0)
        If objBatchJob.JobStatus = elsSuccessful Then
          objBatchJob.SetLastCompleteDate List2.ItemData(lngCount), Date
        End If
  
        If objBatchJob.EventLogIDs <> vbNullString Then
          mstrEventLogIDs = mstrEventLogIDs & _
            IIf(mstrEventLogIDs <> vbNullString, ", ", vbNullString) & _
            objBatchJob.EventLogIDs
        End If

        blnCancelled = (objBatchJob.JobStatus = elsCancelled)

      End If

    End If

  Next

  Set objBatchJob = Nothing

  RunSelectedJobs = False

End Function



Public Property Get ListCount() As Long
  ListCount = IIf(mblnScheduledJobs, List2.ListCount, List1.ListItems.Count)
End Property


Public Property Get EventLogIDs() As String
  EventLogIDs = mstrEventLogIDs
End Property

Public Property Let EventLogIDs(ByVal strNewValue As String)
  mstrEventLogIDs = strNewValue
End Property


Public Function ShowList(lngUtilType As UtilityType, Optional msRecordSourceWhere As String, Optional blnScheduledJobs As Boolean) As Boolean

  Dim strExtraWhereClause As String
  'Dim intWhereClauses As Integer

  mblnApplyDefAccess = True
  mblnApplySystemPermissions = Not gfCurrentUserIsSysSecMgr
  strExtraWhereClause = vbNullString
  msRecordSource = vbNullString
  mblnScheduledJobs = blnScheduledJobs
 
  mutlUtilityType = lngUtilType
  msTableIDColumnName = "TableID"

  Select Case lngUtilType
  Case utlBatchJob
    msTypeCode = "BATCHJOBS"
    
    If blnScheduledJobs Then
      msType = "Scheduled Batch Jobs"
      'msRecordSource = "SELECT DISTINCT ASRSysBatchJobName.*" & _
        " FROM ASRSysBatchJobName" & _
        " LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID" & _
        " WHERE (Scheduled = 1) AND (GETDATE() >= StartDate) " & _
        "   AND (RoleToPrompt = '" & gsUserGroup & "')"
      strExtraWhereClause = "(Scheduled = 1) AND (GETDATE() >= StartDate) " & _
                            "AND (GETDATE() <= dateadd(d,1,EndDate) or EndDate is null) " & _
                            "AND (RoleToPrompt = '" & gsUserGroup & "')"
      msGeneralCaption = "Scheduled Batch Jobs"
      msSingularCaption = "Scheduled Batch Job"
      'Dynamically set HelpContextID in order for the Doc To Help jump to work properly
      Me.HelpContextID = 1107
      
    Else
      msType = "Batch Job"
      msGeneralCaption = "Batch Jobs"
      msSingularCaption = "Batch Job"
      
      Me.HelpContextID = 1084
    End If

    msTableName = "ASRSysBatchJobName"
    msIDField = "ID"
    msAccessTableName = "ASRSysBatchJobAccess"


  Case utlCalendarReport
    msTypeCode = "CALENDARREPORTS"
    msType = "Calendar Report"
    msGeneralCaption = "Calendar Reports"
    msSingularCaption = "Calendar Report"
    msTableName = "ASRSYSCalendarReports"
    msIDField = "ID"
    mutlUtilityType = utlCalendarReport
    msTableIDColumnName = "BaseTable"
    msAccessTableName = "ASRSysCalendarReportAccess"
    Me.HelpContextID = 1085

  Case utlCalculation
    msTypeCode = "CALCULATIONS"
    msType = "Calculation"
    msGeneralCaption = "Calculations"
    msSingularCaption = "Calculation"
    msTableName = "ASRSysExpressions"
    msIDField = "ExprID"
    mutlUtilityType = utlCalculation
    Me.HelpContextID = 1086

  Case utlCrossTab
    msTypeCode = "CROSSTABS"
    msType = "Cross Tab"
    msGeneralCaption = "Cross Tabs"
    msSingularCaption = "Cross Tab"
    msTableName = "ASRSysCrossTab"
    msIDField = "CrossTabID"
    mutlUtilityType = utlCrossTab
    msAccessTableName = "ASRSysCrossTabAccess"
    Me.HelpContextID = 1087

  Case utlCustomReport
    msTypeCode = "CUSTOMREPORTS"
    msType = "Custom Report"
    msGeneralCaption = "Custom Reports"
    msSingularCaption = "Custom Report"
    msTableName = "ASRSYSCustomReportsName"
    msIDField = "ID"
    mutlUtilityType = utlCustomReport
    msAccessTableName = "ASRSysCustomReportAccess"
    Me.HelpContextID = 1088

  Case utlDataTransfer
    msTypeCode = "DATATRANSFER"
    msType = "Data Transfer"
    msGeneralCaption = "Data Transfer"
    msSingularCaption = "Data Transfer"
    msTableName = "ASRSysDataTransferName"
    msIDField = "DataTransferID"
    mutlUtilityType = utlDataTransfer
    msTableIDColumnName = "FromTableID"
    msAccessTableName = "ASRSysDataTransferAccess"
    Me.HelpContextID = 1089

  Case utlEmailAddress
    msTypeCode = "EMAILADDRESSES"
    msType = "Email Address"
    msGeneralCaption = "Email Addresses"
    msSingularCaption = "Email Address"
    msTableName = "ASRSysEmailAddress"
    msIDField = "EmailID"
    msRecordSource = "SELECT EmailID, Name FROM " & msTableName & _
                     " WHERE Type = 0"
    mutlUtilityType = utlEmailAddress
    mblnApplyDefAccess = False
    Me.HelpContextID = 1090

  Case utlEmailGroup
    msTypeCode = "EMAILGROUPS"
    msType = "Email Group"
    msGeneralCaption = "Email Groups"
    msSingularCaption = "Email Group"
    msTableName = "ASRSysEmailGroupName"
    msIDField = "EmailGroupID"
    mutlUtilityType = utlEmailGroup
    Me.HelpContextID = 1091

  Case utlExport
    msTypeCode = "EXPORT"
    msType = "Export"
    msGeneralCaption = "Export"
    msSingularCaption = "Export"
    msTableName = "ASRSysExportName"
    msIDField = "ID"
    mutlUtilityType = utlExport
    msAccessTableName = "ASRSysExportAccess"
    Me.HelpContextID = 1092

  Case utlFilter
    msTypeCode = "FILTERS"
    msType = "Filter"
    msGeneralCaption = "Filters"
    msSingularCaption = "Filter"
    msTableName = "ASRSysExpressions"
    msIDField = "ExprID"
    mutlUtilityType = utlFilter
    Me.HelpContextID = 1093
  Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
  
    Select Case lngUtilType
    Case UtlGlobalAdd: msTypeCode = "GLOBALADD"
        Me.HelpContextID = 1094
    
    Case utlGlobalUpdate: msTypeCode = "GLOBALUPDATE"
        Me.HelpContextID = 1096
    
    Case utlGlobalDelete: msTypeCode = "GLOBALDELETE"
        Me.HelpContextID = 1095
    
    End Select
    
    msType = "Global " & StrConv(Mid$(msTypeCode, 7), vbProperCase)
    msGeneralCaption = msType
    msSingularCaption = msType
    msTableName = "ASRSysGlobalFunctions"
    msIDField = "FunctionID"
    msRecordSourceWhere = msTableName & ".Type = '" & Mid$(msTypeCode, 7, 1) & "'"
    msAccessTableName = "ASRSysGlobalAccess"
  
  Case utlImport
    msTypeCode = "IMPORT"
    msType = "Import"
    msGeneralCaption = "Import"
    msSingularCaption = "Import"
    msTableName = "ASRSysImportName"
    msIDField = "ID"
    mutlUtilityType = utlImport
    msAccessTableName = "ASRSysImportAccess"
    Me.HelpContextID = 1097

  Case utlMatchReport
    msTypeCode = "MATCHREPORTS"
    msType = "Match Report"
    mutlUtilityType = utlMatchReport
    msRecordSourceWhere = "ASRSysMatchReportName.matchReportType = 0"
    msGeneralCaption = "Match Reports"
    msSingularCaption = "Match Report"
    msTableName = "ASRSysMatchReportName"
    msIDField = "MatchReportID"
    msTableIDColumnName = "Table1ID, Table2ID"
    msAccessTableName = "ASRSysMatchReportAccess"
    Me.HelpContextID = 1098
  
  Case utlSuccession
    msTypeCode = "SUCCESSION"   '"SUCCESSIONPLANNING"
    msType = "Succession Planning"
    mutlUtilityType = utlSuccession
    msRecordSourceWhere = "ASRSysMatchReportName.matchReportType = 1"
    msGeneralCaption = "Succession Planning"
    msSingularCaption = "Succession Planning"
    msTableName = "ASRSysMatchReportName"
    msIDField = "MatchReportID"
    msTableIDColumnName = "Table1ID, Table2ID"
    msAccessTableName = "ASRSysMatchReportAccess"
    Me.HelpContextID = 1099
  
  Case utlCareer
    msTypeCode = "CAREER"    '"CAREERPROGRESSION"
    msType = "Career Progression"
    mutlUtilityType = utlCareer
    msRecordSourceWhere = "ASRSysMatchReportName.matchReportType = 2"
    msGeneralCaption = "Career Progression"
    msSingularCaption = "Career Progression"
    msTableName = "ASRSysMatchReportName"
    msIDField = "MatchReportID"
    msTableIDColumnName = "Table1ID, Table2ID"
    msAccessTableName = "ASRSysMatchReportAccess"
    Me.HelpContextID = 1100

  Case utlMailMerge
    msTypeCode = "MAILMERGE"
    msType = "Mail Merge"
    msGeneralCaption = "Mail Merge"
    msSingularCaption = "Mail Merge"
    msTableName = "ASRSysMailMergeName"
    msIDField = "MailMergeID"
    mutlUtilityType = utlMailMerge
    msRecordSourceWhere = "ASRSysMailMergeName.IsLabel = 0"
    msAccessTableName = "ASRSysMailMergeAccess"
    Me.HelpContextID = 1101
    
  Case utlLabel
    msTypeCode = "LABELS"
    msType = "Envelopes & Labels"
    msGeneralCaption = "Envelopes & Labels"
    msSingularCaption = "Envelope & Label"
    msTableName = "ASRSysMailMergeName"
    msIDField = "MailMergeID"
    mutlUtilityType = utlLabel
    msRecordSourceWhere = "ASRSysMailMergeName.IsLabel = 1"
    msAccessTableName = "ASRSysMailMergeAccess"
    Me.HelpContextID = 1102

  Case utlLabelType
    msTypeCode = "LABELDEFINITION"
    msType = "Envelope & Label Template"
    msGeneralCaption = "Envelope & Label Templates"
    msSingularCaption = "Envelope & Label Template"
    msTableName = "ASRSysLabelTypes"
    msIDField = "LabelTypeID"
    mutlUtilityType = utlLabelType
    Me.HelpContextID = 1082

  Case utlDocumentMapping
    msTypeCode = "VERSION1"
    msType = "Document Type"
    msGeneralCaption = "Document Types"
    msSingularCaption = "Document Type"
    msTableName = "ASRSysDocumentManagementTypes"
    msIDField = "DocumentMapID"
    mutlUtilityType = utlDocumentMapping
    Me.HelpContextID = 1082

  Case utlPicklist
    msTypeCode = "PICKLISTS"
    msType = "Picklist"
    msGeneralCaption = "Picklists"
    msSingularCaption = "Picklist"
    msTableName = "ASRSysPickListName"
    msIDField = "PicklistID"
    mutlUtilityType = utlPicklist
    Me.HelpContextID = 1104
  
  Case utlRecordProfile
    msTypeCode = "RECORDPROFILE"
    msType = "Record Profile"
    msGeneralCaption = "Record Profile"
    msSingularCaption = "Record Profile"
    msTableName = "ASRSysRecordProfileName"
    msIDField = "RecordProfileID"
    mutlUtilityType = utlRecordProfile
    msTableIDColumnName = "baseTable"
    msAccessTableName = "ASRSysRecordProfileAccess"
    Me.HelpContextID = 1105
  
  Case utlWorkflow
    msTypeCode = "WORKFLOW"
    msIDField = "ID"
    msTableName = "ASRSysWorkflows"
    mblnApplyDefAccess = False
    mutlUtilityType = utlWorkflow
    
    If blnScheduledJobs Then
      msType = "Pending Workflow Steps"
      msRecordSource = "SELECT ASRSysWorkflowInstanceSteps.ID," & _
        "   ASRSysWorkflows.name + ' - ' + ASRSysWorkflowElements.caption AS [name]," & _
        "   '' AS [description]" & _
        " FROM ASRSysWorkflowInstanceSteps" & _
        " INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & _
        " INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID" & _
        " WHERE ASRSysWorkflowInstanceSteps.ID IN (" & mstrEventLogIDs & ")"
    
      msGeneralCaption = "Pending Workflow Steps"
      msSingularCaption = "Pending Workflow Step"
      Me.HelpContextID = 1145
    Else
      msType = "Workflow"
      msGeneralCaption = "Workflow"
      msSingularCaption = "Workflow"
      msRecordSource = "SELECT ID, name, description" & _
        " FROM " & msTableName & _
        " WHERE ASRSysWorkflows.enabled = 1" & _
        "   AND ISNULL(ASRSysWorkflows.initiationType, 0) = 0"
      Me.HelpContextID = 1106
    End If
  
  
  End Select

  Me.Caption = msGeneralCaption
  msFieldName = "Name"
  
  'MH20030916 Removed this fix as it was causing problems with system permissions.
  'Have now made changes in frmConfiguration which fixes fault 6370.
  '''TM11092003 Fault 6370 - put this case in to avoid changing the Resource file Icon names.
  ''Select Case msTypeCode
  ''Case "SUCCESSIONPLANNING": msIcon = "SUCCESSION"
  ''Case "CAREERPROGRESSION": msIcon = "CAREER"
  ''Case Else:  msIcon = msTypeCode
  ''End Select
  'msIcon = msTypeCode
  
  'MH20041111 Fault 9352
  'If msRecordSource = vbNullString Then
    If msRecordSource = vbNullString And OldAccessUtility(mutlUtilityType) Then
      msRecordSource = _
          "SELECT Name, " & _
          IIf(mblnHideDesc, vbNullString, "Description, ") & _
          IIf(mblnApplyDefAccess, "Username, Access, ", vbNullString) & msIDField & _
        " FROM " & msTableName & IIf(strExtraWhereClause <> vbNullString, " WHERE " & strExtraWhereClause, "")
    'Else
    ElseIf msAccessTableName <> vbNullString Then
      
      'MH20040107 Fault 5627 - Ignore ASRSysGroup
      'msRecordSource = _
        "SELECT " & msTableName & ".name," & _
          IIf(mblnHideDesc, vbNullString, msTableName & ".description, ") & _
          IIf(mblnApplyDefAccess, msTableName & ".userName, " & msAccessTableName & ".access, ", vbNullString) & _
          msTableName & "." & msIDField & _
        " FROM " & msTableName & _
        IIf(mblnApplyDefAccess, " INNER JOIN " & msAccessTableName & " ON " & msTableName & "." & msIDField & " = " & msAccessTableName & ".ID" & _
          " INNER JOIN sysusers b ON " & msAccessTableName & ".groupname = b.name" & _
          " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
          "   AND a.Name = current_user)", vbNullString) & _
        IIf(strExtraWhereClause <> vbNullString, " WHERE " & strExtraWhereClause, "")
      msRecordSource = _
        "SELECT " & msTableName & ".name," & _
          IIf(mblnHideDesc, vbNullString, msTableName & ".description, ") & _
          IIf(mblnApplyDefAccess, msTableName & ".userName, " & msAccessTableName & ".access, ", vbNullString) & _
          msTableName & "." & msIDField & _
        " FROM " & msTableName & _
        IIf(mblnApplyDefAccess, " INNER JOIN " & msAccessTableName & " ON " & msTableName & "." & msIDField & " = " & msAccessTableName & ".ID" & _
          " AND " & msAccessTableName & ".groupname = '" & gsUserGroup & "'", vbNullString) & _
        IIf(strExtraWhereClause <> vbNullString, " WHERE " & strExtraWhereClause, "")
    End If
  'End If

  If msRecordSourceWhere <> vbNullString Then
    msRecordSource = msRecordSource & IIf(strExtraWhereClause <> vbNullString, " AND ", " WHERE ") & msRecordSourceWhere
  End If
    
  If mblnFirstLoad Then
    chkOnlyMine.Value = GetUserSetting("DefSel", "OnlyMine " & msTypeCode, 0)
    mblnFirstLoad = False
  Else
    chkOnlyMine.Value = mintOnlyMine
  End If

  ShowControls
  ShowList = Populate_List

End Function
Private Function DynamicallyChangeHelpContextID() As Integer
    
    Select Case mutlUtilityType
        Case utlBatchJob
        DynamicallyChangeHelpContextID = 1084
        
        Case utlCrossTab
        DynamicallyChangeHelpContextID = 1087
        
        Case utlCustomReport
        DynamicallyChangeHelpContextID = 1088
        
        Case utlDataTransfer
        DynamicallyChangeHelpContextID = 1089
        
        Case utlExport
        DynamicallyChangeHelpContextID = 1092
        
        Case UtlGlobalAdd
        DynamicallyChangeHelpContextID = 1094
        
        Case utlGlobalDelete
        DynamicallyChangeHelpContextID = 1095
        
        Case utlGlobalUpdate
        DynamicallyChangeHelpContextID = 1096
        
        Case utlImport
        DynamicallyChangeHelpContextID = 1097
        
        Case utlMailMerge
        DynamicallyChangeHelpContextID = 1101
        
        Case utlPicklist
        DynamicallyChangeHelpContextID = 1104
        
        Case utlFilter
        DynamicallyChangeHelpContextID = 1093
        
        Case utlCalculation
        DynamicallyChangeHelpContextID = 1086
        
        Case utlOrder
        DynamicallyChangeHelpContextID = 1048
        
        Case utlMatchReport
        DynamicallyChangeHelpContextID = 1098
        
        Case utlAbsenceBreakdown
        DynamicallyChangeHelpContextID = 1004
        
        Case utlBradfordFactor
        DynamicallyChangeHelpContextID = 1011
        
        Case utlCalendarReport
        DynamicallyChangeHelpContextID = 1066
        
        Case utlLabel
        DynamicallyChangeHelpContextID = 1102
        
        Case utlLabelType
        DynamicallyChangeHelpContextID = 1102
        
        Case utlRecordProfile
        DynamicallyChangeHelpContextID = 1105
        
        Case utlEmailAddress
        DynamicallyChangeHelpContextID = 1090
        
        Case utlEmailGroup
        DynamicallyChangeHelpContextID = 1091
        
        Case utlSuccession
        DynamicallyChangeHelpContextID = 1099
        
        Case utlCareer
        DynamicallyChangeHelpContextID = 1100
        
        Case utlWorkflow
        DynamicallyChangeHelpContextID = 1106
        
        Case utlWorkFlowPendingSteps
        DynamicallyChangeHelpContextID = 1145
        
        Case utlWorkFlowPendingSteps
        DynamicallyChangeHelpContextID = 1027
        
        
    End Select
End Function

Public Property Get SelectedIDs() As Variant
  ' Return an array of the selected IDs in List2 (ie. scheduled batch jobs, or pending workflow steps)
  ' NB. The ReadSelectedIDs needs to be run before using this property.
  
  SelectedIDs = malngSelectedIDs
  
End Property
Private Sub ReadSelectedIDs()
  ' Read into an array the selected IDs in List2 (ie. scheduled batch jobs, or pending workflow steps)
  Dim lngCount As Long
  
  ReDim malngSelectedIDs(0)
  
  For lngCount = 1 To List2.ListCount - 1
    If List2.Selected(lngCount) Then
      ReDim Preserve malngSelectedIDs(UBound(malngSelectedIDs) + 1)
      
      malngSelectedIDs(UBound(malngSelectedIDs)) = List2.ItemData(lngCount)
    End If
  Next lngCount
  
End Sub

' Show the selection or just force through the specified ID
Public Sub CustomShow(ByVal ShowMode As VBRUN.FormShowConstants)

  If gbJustRunIt Then
    Me.SelectedID = glngBypassDefsel_ID
    Me.Action = edtSelect
  Else
    Me.Show ShowMode
  End If

End Sub




