VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDefSel 
   Caption         =   "Select"
   ClientHeight    =   6045
   ClientLeft      =   2715
   ClientTop       =   2535
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8066
   Icon            =   "frmDefSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTopButtons 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2910
      Left            =   3645
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Proper&ties..."
         Height          =   400
         Left            =   0
         TabIndex        =   13
         Top             =   2505
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   0
         TabIndex        =   12
         Top             =   1995
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit..."
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   495
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete..."
         Height          =   400
         Left            =   0
         TabIndex        =   11
         Top             =   1500
         Width           =   1200
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Cop&y..."
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   1005
         Width           =   1200
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New..."
         Height          =   400
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3465
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Height          =   1080
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4155
         Width           =   3450
      End
      Begin VB.ComboBox cboTables 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   555
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin ComctlLib.ListView List1 
         Height          =   3705
         Left            =   0
         TabIndex        =   4
         Top             =   375
         Width           =   3450
         _ExtentX        =   6085
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
            Object.Width           =   5644
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
         TabIndex        =   5
         Top             =   60
         Width           =   450
      End
   End
   Begin VB.Frame fraBottomButtons 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1360
      Left            =   3645
      TabIndex        =   7
      Top             =   4005
      Width           =   1215
      Begin VB.CommandButton cmdNone 
         Caption         =   "N&one"
         Height          =   400
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   400
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Height          =   400
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   1200
      End
   End
   Begin VB.CheckBox chkOnlyMine 
      Caption         =   "On&ly show definitions where owner is 'username'"
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   5520
      Width           =   4500
   End
   Begin ActiveBarLibraryCtl.ActiveBar abDefSel 
      Left            =   3630
      Top             =   3420
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
  
Private lngAction As Long
Private mlngOptions As Long
Private msSelectedText As String
Private mlngSelectedID As Long
Private mlngTableID As Long

Private mblnCaptionIsRun As Boolean
Private mblnShowOrders As Boolean
Private mblnApplySystemPermissions As Boolean
Private mblnHideDesc As Boolean
Private mblnTableComboVisible As Boolean

Private msTableName As String
Private msFieldName As String
Private msIDField As String
Private msType As String
Private msTypeCode As String
Private msRecordSource As String
Private msIcon As String
Private miExpressionType As ExpressionTypes
Private msTableIDColumnName As String

Private mbFromCopy As Boolean
Private mlCopyID As Long

Private mblnHiddenDef As Boolean
Private mblnReadOnlyAccess As Boolean

Private mfEnableNew As Boolean
Private mfEnableView As Boolean
Private mfEnableEdit As Boolean
Private mfEnableDelete As Boolean
Private mfEnableRun As Boolean
Private mintOnlyMine As Integer

Private msGeneralCaption As String
Private msSingularCaption As String

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
'  If Not mblnShowOrders Then
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
  lngAction = 0
  Unload Me
  
End Sub

Private Sub cmdDelete_Click()

  Dim lngHighLightIndex As Long
  Dim lngSelectedID As Long
  Dim sSQL As String
  Dim rsTemp As Recordset
  Dim objExpression As clsExprExpression
  Dim objDatabase As SecurityMgr.Database
  
  'TM01062004 Fault 8730 need to initialize this object.
  Set objDatabase = New SecurityMgr.Database
  
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
  If msTypeCode = "CALCULATIONS" Or msTypeCode = "FILTERS" _
    Or msTypeCode = "CUSTOMREPORTS" Or msTypeCode = "MAILMERGE" _
    Or msTypeCode = "GLOBALADD" Or msTypeCode = "GLOBALUPDATE" _
    Or msTypeCode = "EXPORT" Or msTypeCode = "GLOBALDELETE" _
    Or msTypeCode = "DATATRANSFER" Then
    
    If MsgBox("Delete this definition are you sure ?", vbQuestion + vbYesNo, "Delete " & msType) = vbYes Then
      If Not CheckForUseage(msType, lngSelectedID) Then
        Unload Me
      End If
    End If
    
  Else
    ' Ask for user confirmation to delete the utility definition
    If MsgBox("Delete this " & LCase(msType) & ", are you sure ?", vbQuestion + vbYesNo, "Delete " & msType) = vbYes Then
      If Not CheckForUseage(msType, lngSelectedID) Then
          
        ' NEWACCESS - needs to be updated as each report/utility is updated for the new access.
        Select Case msType
        Case "Batch Job"
          objDatabase.DeleteRecord "ASRSysBatchJobAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "AsrSysBatchJobDetails", "BatchJobNameID", lngSelectedID
  
        Case "Calendar Reports"
          objDatabase.DeleteRecord "ASRSysCalendarReportAccess", "ID", lngSelectedID
        
        Case "Custom Reports"
          objDatabase.DeleteRecord "ASRSysCustomReportAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "ASRSysCustomReportsDetails", "CustomReportID", lngSelectedID
        
        Case "Cross Tab"
          objDatabase.DeleteRecord "ASRSysCrossTabAccess", "ID", lngSelectedID
  
        Case "Data Transfer"
          objDatabase.DeleteRecord "ASRSysDataTransferAccess", "ID", lngSelectedID
  
        Case "Export"
          objDatabase.DeleteRecord "ASRSysExportAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "AsrSysExportDetails", "ExportID", lngSelectedID
                    
        Case "Global Add", "Global Delete", "Global Update"
          objDatabase.DeleteRecord "ASRSysGlobalAccess", "ID", lngSelectedID
  
        Case "Import"
          objDatabase.DeleteRecord "ASRSysImportAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "ASRSysImportDetails", "ImportID", lngSelectedID
        
        Case "Mail Merge", "Envelopes & Labels"
          objDatabase.DeleteRecord "ASRSysMailMergeAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "ASRSysMailMergeColumns", "MailMergeID", lngSelectedID
        
        Case "Match Report", "Career Progression", "Succession Planning"
          objDatabase.DeleteRecord "ASRSysMatchReportAccess", "ID", lngSelectedID
        
        Case "Record Profile"
          objDatabase.DeleteRecord "ASRSysRecordProfileAccess", "ID", lngSelectedID
          objDatabase.DeleteRecord "ASRSysRecordProfileDetails", "RecordProfileID", lngSelectedID
          objDatabase.DeleteRecord "ASRSysRecordProfileTables", "RecordProfileID", lngSelectedID
           
          ' Also need to delete the file filter expression record (if one exists).
          sSQL = "SELECT filterID" & _
            " FROM ASRSysImportName" & _
            " WHERE ID = " & Trim(Str(lngSelectedID))

          Set rsTemp = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
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
        End Select
  
        objDatabase.DeleteRecord msTableName, msIDField, lngSelectedID
        
        lngHighLightIndex = List1.SelectedItem.Index
        List1.ListItems.Remove lngHighLightIndex
        If List1.ListItems.Count > 0 Then
          Set List1.SelectedItem = List1.ListItems.Item(IIf(lngHighLightIndex < List1.ListItems.Count, lngHighLightIndex, List1.ListItems.Count))
        End If
  
        Refresh_Controls
      End If
    End If
  End If
  
  Set objDatabase = Nothing
  
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
  Dim varWhereUsed As Variant
  Dim intCount As Integer
  Dim strUsage As String
  
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
    .PopulateUtil miExpressionType, mlngSelectedID
    .CheckForUseage msType, mlngSelectedID
    Screen.MousePointer = vbDefault
    .Show vbModal
  End With
     
  SetupModuleParameters
TidyUp:

  Unload frmDefProp
  Set frmDefProp = Nothing

  Exit Sub
  
Prop_ERROR:
  
  Screen.MousePointer = vbDefault
  MsgBox "Error retrieving properties for this definition." & vbCrLf & "Please contact support stating : " & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Properties"
  Resume TidyUp

End Sub

Private Sub cmdSelect_Click()
  
  ' Select the selected item.
  lngAction = edtSelect
  GetSelected
  
  If CanStillSeeDefinition(mlngSelectedID) Then
    Unload Me
  End If
    
End Sub

Private Sub Form_Activate()
 
  Screen.MousePointer = vbDefault

  If List1.Visible And List1.Enabled Then
    List1.SetFocus
  End If

  Refresh_Controls
  
  'JDM - 26/11/01 - fault 3203 - Disable the table menu
  cboTables.Enabled = False

  ' Exterminate the form icon
  RemoveIcon Me

End Sub


Private Sub Form_Initialize()
  mintOnlyMine = -1 'Read from settings
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
  Case KeyCode = vbKeyF5
      Populate_List
  Case KeyCode = vbKeyDelete
      If cmdDelete.Enabled Then
        cmdDelete_Click
       End If
  End Select

End Sub

Private Sub Form_Load()
  
  Hook Me.hWnd, Me.Width, Me.Height
  
  mblnLoading = False
  If mlngOptions = 0 Then
    mlngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtSelect + edtPrint + edtProperties
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    lngAction = 0
  End If
  If Not FromCopy Then
    GetSelected
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

  If Me.Visible = True Then
    SizeControls
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub List1_ItemClick(ByVal Item As ComctlLib.ListItem)
  Refresh_Controls
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
          If MsgBox("Are you sure you want to run the '" & List1.SelectedItem.Text & "' " & Me.Caption & " ?", vbYesNo + vbQuestion, "Confirmation...") = vbNo Then
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

Private Sub Display_Button(Button As VB.CommandButton, ByVal BtnOpt As Long, ByVal X As Long, ByRef Y As Long)
  If (Me.Options And BtnOpt) Then
    Button.Move X, Y
    Button.Visible = True
    Y = Y + cmdNew.Height + ((UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY) * 1.5)
  Else
    Button.Visible = False
  End If
  
End Sub


Public Sub Refresh_Controls()
  
  ' Enable/disable controls as required.
  If List1.ListItems.Count > 0 And Not IsEmpty(List1.SelectedItem) Then
    With mrsRecords
      .MoveFirst
      .Find msIDField & " = " & CStr(List1.SelectedItem.Tag)  'CStr(List1.ItemData(List1.ListIndex))
    
      If mblnHideDesc = False Then
        txtDesc.Text = IIf(IsNull(.Fields("Description").Value), vbNullString, .Fields("Description").Value)
      End If

      If Not mblnShowOrders Then
        mblnHiddenDef = (.Fields("Access").Value = ACCESS_HIDDEN)
        
        'JPD 20030909 Fault 6912
        mblnReadOnlyAccess = (.Fields("Access").Value = ACCESS_READONLY) And _
            (LCase(Trim$(.Fields("Username").Value)) <> LCase(gsUserName)) And _
            (Not gfCurrentUserIsSysSecMgr)
      End If
    End With
  
    cmdNew.Enabled = (cmdNew.Visible And mfEnableNew)
    cmdCopy.Enabled = (cmdCopy.Visible And mfEnableNew)
    cmdEdit.Enabled = (cmdEdit.Visible And mfEnableView)
    cmdDelete.Enabled = (cmdDelete.Visible And mfEnableDelete And Not (mblnReadOnlyAccess))
    cmdSelect.Enabled = (cmdSelect.Visible And mfEnableRun)
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
    
    cmdNew.Enabled = mfEnableNew
    cmdCopy.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSelect.Enabled = False
    cmdPrint.Enabled = False
    cmdProperties.Enabled = False
    
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
  
  CheckListViewColWidth List1
  
  If List1.ListItems.Count > 0 Then
    List1.ListItems(List1.SelectedItem.Index).Selected = True   'This highlights the current item!!!!!
    List1.Refresh
  End If

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
    If Not IsEmpty(List1.SelectedItem) And (Not List1.SelectedItem Is Nothing) Then
      mlngSelectedID = Val(List1.SelectedItem.Tag)
      SelectedText = List1.SelectedItem.Text
    Else
      mlngSelectedID = 0
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

'Private Sub EditFromCopy()
'
'  lngAction = edtEdit
'  mlngSelectedID = mlCopyID
'  SelectedText = List1.List(List1.ListIndex)
'  Unload Me
'
'End Sub

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

        MsgBox "Unable to delete this " & LCase(sDefType) & ", " & sMsg, vbExclamation, Me.Caption
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
    mfEnableNew = SystemPermission(msTypeCode, "NEW")
    mfEnableDelete = SystemPermission(msTypeCode, "DELETE")
    mfEnableEdit = SystemPermission(msTypeCode, "EDIT")
    mfEnableView = SystemPermission(msTypeCode, "VIEW")

    'If not edit but still have view then change the caption of command button
    If mfEnableEdit = False Then
      If mfEnableView = True Then
        cmdEdit.Caption = "&View..."
        Me.abDefSel.Bands("bndDefSel").Tools("EditView").Caption = cmdEdit.Caption
      End If
    End If
    
    If mblnCaptionIsRun Then
      mfEnableRun = SystemPermission(msTypeCode, "RUN")
    End If

  End If

End Sub


Private Function Populate_List() As Boolean
  
  'MH20000302 - Changed this sub from public to private
  
  'MH20000807 - Rather than sort the listview do the sort in the SQL so
  '             that you will always be able to see the selected item
  '             when the list is first shown (Fault 725)
  Dim strSQL As String
  Dim intCount As Integer
  Dim objListItem As ListItem
  
  ' Populate the selection listbox with the information defined in the given parameters.
  On Error GoTo ErrorTrap
  
  If mblnLoading Then
    Exit Function
  End If
  mblnLoading = True
     
  List1.ListItems.Clear
  UI.LockWindow Me.hWnd

  strSQL = msRecordSource

  If Not mblnShowOrders Then
    'JPD 20050812 Fault 10166
    'JPD 20030915 Fault 6966
    strSQL = strSQL & _
      IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
      "(Username = '" & Replace(gsUserName, "'", "''") & "'" & _
      IIf((chkOnlyMine = False), " OR Access <> '" & ACCESS_HIDDEN & "'" & IIf(gfCurrentUserIsSysSecMgr, " OR 1 = 1", vbNullString), vbNullString) & _
      ")"
  End If
  
  If mlngTableID > 0 Then
    strSQL = strSQL & _
      IIf(InStr(strSQL, " WHERE ") = 0, " WHERE ", " AND ") & _
      msTableIDColumnName & " = " & CStr(mlngTableID)
  End If

  strSQL = strSQL & " ORDER BY Name"
  Set mrsRecords = modExpression.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)

  With mrsRecords
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        'List1.AddItem RemoveUnderScores(.Fields(msFieldName))
        'List1.ItemData(List1.NewIndex) = .Fields(msIDField)
        Set objListItem = List1.ListItems.Add(, , RemoveUnderScores(.Fields("Name").Value))
        objListItem.Tag = .Fields(msIDField).Value

        If .Fields(msIDField).Value = mlngSelectedID Then
          'List1.ListIndex = List1.NewIndex
          Set List1.SelectedItem = objListItem
        End If
        
        .MoveNext
      Loop
    End If
  End With

  If List1.ListItems.Count > 0 Then
    If IsEmpty(List1.SelectedItem) Then
      Set List1.SelectedItem = List1.ListItems(1)
    End If
  End If

  
  ApplySystemPermissions
  
  Populate_List = True
  
Exit_Populate_List:
'  Set mrsRecords = Nothing
  
  mblnLoading = False
  UI.UnlockWindow
  'List1.Refresh
  
  Refresh_Controls
  
  Exit Function
  
ErrorTrap:
  Populate_List = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  
  Resume Exit_Populate_List
  
End Function


Public Function ShowList(sCode As String, Optional msRecordSourceWhere As String) As Boolean

  mblnShowOrders = False
  mblnApplySystemPermissions = False
  
  msTypeCode = UCase$(sCode)
  miExpressionType = -1
  msTableIDColumnName = "TableID"

  Select Case msTypeCode
    Case "CALCULATIONS"
      msType = "Calculation"
      msGeneralCaption = "Calculations"
      msSingularCaption = "Calculation"
      msTableName = "ASRSysExpressions"
      msIDField = "ExprID"
      msIcon = "EXPR_CALCULATION"
      'mblnApplySystemPermissions = False
      miExpressionType = giEXPR_RUNTIMECALCULATION
      Me.HelpContextID = 8062
      
    Case "FILTERS"
      msType = "Filter"
      msGeneralCaption = "Filters"
      msSingularCaption = "Filter"
      msTableName = "ASRSysExpressions"
      msIDField = "ExprID"
      msIcon = "EXPR_FILTER"
      miExpressionType = giEXPR_RUNTIMEFILTER
      Me.HelpContextID = 8063
    
  End Select
  
  Me.Caption = msGeneralCaption
  
  msFieldName = "Name"
  msRecordSource = "SELECT Name, Description, Username, Access, " & msIDField & _
         " FROM " & msTableName
         
  If msRecordSourceWhere <> vbNullString Then
    msRecordSource = msRecordSource & " WHERE " & msRecordSourceWhere
  End If
    
  If mintOnlyMine = -1 Then
    chkOnlyMine.Value = GetUserSetting("DefSel", "OnlyMine " & msTypeCode, 0)
  Else
    chkOnlyMine.Value = mintOnlyMine
  End If
  
  'Me.Caption = msType
  'Call DrawControls
  ShowControls
  'SizeControls
  ShowList = Populate_List

End Function

Private Sub PopulateTables()

  Dim lngTableID As Long

  lngTableID = mlngTableID
  LoadTableCombo cboTables

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

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  'MH20001013 Fault 1055
  'Need to include table name otherwise get Ambiguous column name message !
  If InStr(msRecordSource, " WHERE ") = 0 Then
    strSQL = msRecordSource & " WHERE " & msTableName & "." & msIDField & " = " & CStr(lngDefID)
  Else
    strSQL = msRecordSource & " AND " & msTableName & "." & msIDField & " = " & CStr(lngDefID)
  End If
  Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  CanStillSeeDefinition = True

  With rsTemp
    If .BOF And .EOF Then
      MsgBox "This definition has been deleted by another user", vbExclamation, Me.Caption
      Call Populate_List
      CanStillSeeDefinition = False
      'Exit Sub
    
    ElseIf Not mblnShowOrders Then
    
      If LCase(Trim$(!UserName)) <> LCase(gsUserName) Then
      
        If Not gfCurrentUserIsSysSecMgr Then
          If !Access = ACCESS_HIDDEN Then
            MsgBox "This definition has been made hidden by another user", vbExclamation, Me.Caption
            Call Populate_List
            CanStillSeeDefinition = False
      
          ElseIf !Access = ACCESS_READONLY And Not mblnReadOnlyAccess Then
            MsgBox "This definition is now read only", vbInformation, Me.Caption
            mblnReadOnlyAccess = True
            Call CanStillSeeDefinition(lngDefID)  'Check again after msgbox
      
          ElseIf !Access = ACCESS_READWRITE And mblnReadOnlyAccess Then
            MsgBox "This definition is now read write", vbInformation, Me.Caption
            mblnReadOnlyAccess = False
            Call CanStillSeeDefinition(lngDefID)  'Check again after msgbox
  
          End If
        End If
      End If
    End If

  End With

End Function


Private Sub CheckListViewColWidth(lstvw As ListView)

  Dim objItem As ListItem
  Dim lngMax As Long
  Dim lngLen As Long

  lngMax = 0

  For Each objItem In lstvw.ListItems

    lngLen = Me.TextWidth(objItem.Text)
    If lngMax < lngLen Then
      lngMax = lngLen
    End If

  Next objItem

  lngMax = lngMax + 60
  lstvw.ColumnHeaders(1).Width = lngMax
  lstvw.Refresh

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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
    'TM20011217 Fault 3250 & Fault 3358 - Now using the mblnShowOrders boolean to show the "None" button or not.
    If mblnShowOrders Or miExpressionType = giEXPR_RUNTIMEFILTER Then
'    If (mlngOptions And edtDeselect) Then
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
  
  ' Enable/disable the command controls as required.
  Refresh_Controls
  
  
  lblTables.Visible = mblnTableComboVisible
  cboTables.Visible = mblnTableComboVisible
  If mblnTableComboVisible Then
    Call PopulateTables
  End If
  
  txtDesc.Visible = Not mblnHideDesc
  chkOnlyMine.Visible = Not mblnShowOrders
  chkOnlyMine.Caption = "On&ly show definitions where owner is '" & _
                        StrConv(gsUserName, vbProperCase) & "'"

  'Make frame background same colour as form
  fraMain.BackColor = Me.BackColor
  fraTopButtons.BackColor = Me.BackColor
  fraBottomButtons.BackColor = Me.BackColor

LocalErr:
  UI.UnlockWindow

End Sub

Public Function ShowOrders(strSQL As String, lngOrderID As Long) As Boolean
  
  mblnShowOrders = True
  mlngSelectedID = lngOrderID
  
  msRecordSource = strSQL
  msType = "Order"
  msGeneralCaption = "Orders"
  msSingularCaption = "Order"
  msFieldName = "Name"
  msTableIDColumnName = "TableID"
  msTableName = "ASRSysOrders"
  'msIDField = "ASRSysOrders.OrderID"
  msIDField = "OrderID"
  msIcon = "EXPR_ORDER"

  mblnApplySystemPermissions = False
  Me.Caption = msGeneralCaption
  
  'Call DrawControls
  ShowControls
  ShowOrders = Populate_List

End Function

Private Sub SizeControls()

  Dim lngListTop As Long
  Dim lngOffset As Long
  Const lngGap = 100
  Dim blnCheckBoxVisible As Boolean

  lngOffset = Me.ScaleHeight - (lngGap * 2)
  
  'chkOnlyMine (Outside of frame)
  blnCheckBoxVisible = True
  If blnCheckBoxVisible Then
    chkOnlyMine.Move fraMain.Left, Me.ScaleHeight - (chkOnlyMine.Height + lngGap), Me.ScaleWidth
    lngOffset = lngOffset - (chkOnlyMine.Height + lngGap)
  End If

  'Move Frames
  fraMain.Move lngGap, lngGap, Me.ScaleWidth - (fraTopButtons.Width + (lngGap * 3)), lngOffset
  fraTopButtons.Move fraMain.Left + fraMain.Width + lngGap, lngGap
  fraBottomButtons.Move fraTopButtons.Left, (fraMain.Top + fraMain.Height) - fraBottomButtons.Height

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

End Sub


