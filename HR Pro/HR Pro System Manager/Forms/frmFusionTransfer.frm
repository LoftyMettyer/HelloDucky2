VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmFusionTransfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fusion Integration"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5060
   Icon            =   "frmFusionTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNone 
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7575
      TabIndex        =   16
      Top             =   6105
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   6255
      TabIndex        =   8
      Top             =   5385
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7540
      TabIndex        =   9
      Top             =   5385
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   5190
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Definition"
      TabPicture(0)   =   "frmFusionTransfer.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraFusionDefinition"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Settings"
      TabPicture(1)   =   "frmFusionTransfer.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDefaults"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Login Parameters"
      TabPicture(2)   =   "frmFusionTransfer.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraFusionDefinition 
         Caption         =   "Definition : "
         Height          =   4605
         Left            =   -74865
         TabIndex        =   10
         Top             =   405
         Width           =   8340
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Cle&ar"
            Enabled         =   0   'False
            Height          =   400
            Left            =   6930
            TabIndex        =   6
            Top             =   1665
            Width           =   1200
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Height          =   315
            Left            =   7770
            TabIndex        =   3
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5085
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   270
            Width           =   2685
         End
         Begin VB.ComboBox cboFusionType 
            Height          =   315
            ItemData        =   "frmFusionTransfer.frx":0060
            Left            =   945
            List            =   "frmFusionTransfer.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   270
            Width           =   3255
         End
         Begin VB.ComboBox cboFusionTables 
            Height          =   315
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   675
            Width           =   3255
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   6930
            TabIndex        =   5
            Top             =   1170
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdFusionDetails 
            Height          =   3255
            Index           =   0
            Left            =   180
            TabIndex        =   4
            Top             =   1170
            Width           =   6510
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   22
            stylesets.count =   2
            stylesets(0).Name=   "KeyField"
            stylesets(0).BackColor=   14024703
            stylesets(0).HasFont=   -1  'True
            BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(0).Picture=   "frmFusionTransfer.frx":0064
            stylesets(1).Name=   "Mandatory"
            stylesets(1).BackColor=   15400959
            stylesets(1).HasFont=   -1  'True
            BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(1).Picture=   "frmFusionTransfer.frx":0080
            AllowDelete     =   -1  'True
            AllowUpdate     =   0   'False
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   -2147483640
            ForeColorOdd    =   -2147483640
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   22
            Columns(0).Width=   5292
            Columns(0).Caption=   "Fusion Field"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5741
            Columns(1).Caption=   "OpenHR Value"
            Columns(1).Name =   "Display_ToMapType"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "ASRMapType"
            Columns(2).Name =   "ASRMapType"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "ASRTableID"
            Columns(3).Name =   "ASRTableID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "ASRColumnID"
            Columns(4).Name =   "ASRColumnID"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "ASRExprID"
            Columns(5).Name =   "ASRExprID"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "ASRValue"
            Columns(6).Name =   "ASRValue"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "Mandatory"
            Columns(7).Name =   "Mandatory"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "TransferFieldID"
            Columns(8).Name =   "FusionFieldID"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "IsCompanyCode"
            Columns(9).Name =   "IsCompanyCode"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "IsEmployeeCode"
            Columns(10).Name=   "IsEmployeeCode"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "Direction"
            Columns(11).Name=   "Direction"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   3200
            Columns(12).Visible=   0   'False
            Columns(12).Caption=   "IsKeyField"
            Columns(12).Name=   "IsKeyField"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            Columns(13).Width=   3200
            Columns(13).Visible=   0   'False
            Columns(13).Caption=   "AlwaysTransfer"
            Columns(13).Name=   "AlwaysTransfer"
            Columns(13).DataField=   "Column 13"
            Columns(13).DataType=   8
            Columns(13).FieldLen=   256
            Columns(14).Width=   3200
            Columns(14).Visible=   0   'False
            Columns(14).Caption=   "ConvertData"
            Columns(14).Name=   "ConvertData"
            Columns(14).DataField=   "Column 14"
            Columns(14).DataType=   8
            Columns(14).FieldLen=   256
            Columns(15).Width=   3200
            Columns(15).Visible=   0   'False
            Columns(15).Caption=   "IsEmployeeName"
            Columns(15).Name=   "IsEmployeeName"
            Columns(15).DataField=   "Column 15"
            Columns(15).DataType=   8
            Columns(15).FieldLen=   256
            Columns(16).Width=   3200
            Columns(16).Visible=   0   'False
            Columns(16).Caption=   "IsDepartmentCode"
            Columns(16).Name=   "IsDepartmentCode"
            Columns(16).DataField=   "Column 16"
            Columns(16).DataType=   8
            Columns(16).FieldLen=   256
            Columns(17).Width=   3200
            Columns(17).Visible=   0   'False
            Columns(17).Caption=   "IsDepartmentName"
            Columns(17).Name=   "IsDepartmentName"
            Columns(17).DataField=   "Column 17"
            Columns(17).DataType=   8
            Columns(17).FieldLen=   256
            Columns(18).Width=   3200
            Columns(18).Visible=   0   'False
            Columns(18).Caption=   "IsPayrollCode"
            Columns(18).Name=   "IsFusionCode"
            Columns(18).DataField=   "Column 18"
            Columns(18).DataType=   8
            Columns(18).FieldLen=   256
            Columns(19).Width=   3200
            Columns(19).Visible=   0   'False
            Columns(19).Caption=   "Group"
            Columns(19).Name=   "Group"
            Columns(19).DataField=   "Column 19"
            Columns(19).DataType=   8
            Columns(19).FieldLen=   256
            Columns(20).Width=   3200
            Columns(20).Visible=   0   'False
            Columns(20).Caption=   "PreventModify"
            Columns(20).Name=   "PreventModify"
            Columns(20).DataField=   "Column 20"
            Columns(20).DataType=   8
            Columns(20).FieldLen=   256
            Columns(21).Width=   3200
            Columns(21).Visible=   0   'False
            Columns(21).Caption=   "Datatype"
            Columns(21).Name=   "Datatype"
            Columns(21).DataField=   "Column 21"
            Columns(21).DataType=   8
            Columns(21).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   11483
            _ExtentY        =   5741
            _StockProps     =   79
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
         Begin VB.Label lblFilter 
            Caption         =   "Filter : "
            Height          =   240
            Left            =   4410
            TabIndex        =   15
            Top             =   315
            Width           =   555
         End
         Begin VB.Label lblTransferType 
            Caption         =   "Type :"
            Height          =   285
            Left            =   225
            TabIndex        =   12
            Top             =   315
            Width           =   555
         End
         Begin VB.Label lblTransferTable 
            Caption         =   "Table : "
            Height          =   285
            Left            =   225
            TabIndex        =   11
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.Frame fraDefaults 
         Caption         =   "Options : "
         Height          =   840
         Left            =   135
         TabIndex        =   13
         Top             =   495
         Width           =   8340
         Begin VB.CheckBox chkAllowDelete 
            Caption         =   "Allo&w deletion of transfered records"
            Height          =   285
            Left            =   255
            TabIndex        =   7
            Top             =   345
            Width           =   3600
         End
      End
   End
End
Attribute VB_Name = "frmFusionTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Must be public so the details form can change the bookmark of the recordset
Public mrstHeaders As Recordset

Private mbReadOnly As Boolean
Private mbLoading As Boolean
Private mbChanged As Boolean

Private miFusionTypesAmount As Integer

Private mavarFusionBaseTableIDs() As Variant
Private mavarFusionFilterIDs() As Long
Private mstrFusionTypesVisible As String
Private mdblFusionVersions() As Double

Private miStatusForUtilities As FusionTransactionStatus
Private mbAllowDeletions As Boolean

Private mfChanged As Boolean
Private mfCancelled As Boolean


Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  cmdOk.Enabled = True
End Property
Private Sub RefreshButtons()
  
  If Not mbLoading Then
    cmdOk.Enabled = mfChanged And Not mbReadOnly
    cmdEdit.Enabled = (cboFusionTables <> "<None>")
    cmdDelete.Enabled = (cboFusionTables <> "<None>") And (Not mbReadOnly)
    cmdNone.Enabled = (SelectedComboItem(cboFusionTables) > 0) And Not mbReadOnly
    cmdFilter.Enabled = (SelectedComboItem(cboFusionTables) > 0)
  End If

End Sub

Private Sub cboStatusForUtilities_Click()
  Changed = True
End Sub

Private Sub cboFusionTables_Click()

  Dim lngIndex As Long

  If SelectedComboItem(cboFusionTables) <> mavarFusionBaseTableIDs(2, cboFusionType.ListIndex) Then
    
    If mavarFusionBaseTableIDs(2, cboFusionType.ListIndex) > 0 Then
    
      If MsgBox("Changing the base table will reset all the columns for this Fusion type." & vbCrLf _
        & "Are you sure you want to continue?", vbYesNo + vbQuestion, "Fusion Setup") = vbYes Then
        
        PopulateFusionTransferDetails cboFusionType.ListIndex, True
        mavarFusionBaseTableIDs(2, cboFusionType.ListIndex) = SelectedComboItem(cboFusionTables)
        txtFilter.Text = ""
          
      Else
        lngIndex = mavarFusionBaseTableIDs(2, cboFusionType.ListIndex)
        SetComboItem cboFusionTables, lngIndex
      End If
    Else
      mavarFusionBaseTableIDs(2, cboFusionType.ListIndex) = SelectedComboItem(cboFusionTables)
      GoTopOfGrid 0, (cboFusionTables = "<None>")
      cmdEdit.Enabled = (cboFusionTables = "<None>")
    End If
    
    Changed = True
    
  End If
  
  RefreshButtons
  
End Sub

Private Sub cboFusionType_Click()

  Dim iCount As Integer
  Dim iIndex As Integer
  'Set the base table
  SetComboItem cboFusionTables, CLng(mavarFusionBaseTableIDs(2, cboFusionType.ListIndex))
  
  For iCount = grdFusionDetails.LBound To grdFusionDetails.UBound
    grdFusionDetails(iCount).Visible = (cboFusionType.ListIndex = iCount)
    grdFusionDetails.Item(iCount).SelBookmarks.RemoveAll
    GoTopOfGrid CLng(iCount), (cboFusionTables = "<None>")
  Next iCount
  'Set the filter information
  txtFilter.Tag = mavarFusionFilterIDs(cboFusionType.ListIndex)
  txtFilter.Text = GetExpressionName(txtFilter.Tag)
  
  Changed = True
  RefreshButtons
  
End Sub
Private Sub GoTopOfGrid(lngIndex As Long, fClearBookmarks As Boolean)

If fClearBookmarks Then
  grdFusionDetails.Item(lngIndex).SelBookmarks.RemoveAll
Else
  grdFusionDetails.Item(lngIndex).SelBookmarks.Add grdFusionDetails.Item(lngIndex).Bookmark
  grdFusionDetails.Item(lngIndex).Bookmark = grdFusionDetails.Item(lngIndex).SelBookmarks(lngIndex)
End If

End Sub


Private Sub chkAllowDelete_Click()
Changed = True
End Sub

Private Sub chkAllowStatusChange_Click()
Changed = True
End Sub

Private Sub cmdCancel_Click()
  'AE20071119 Fault #12607
'  Dim pintAnswer As Integer
'    If Changed = True And cmdOk.Enabled Then
'      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
'      If pintAnswer = vbYes Then
'        'AE20071108 Fault #12551
'        'Using Me.MousePointer = vbNormal forces the form to be reloaded
'        'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
'        'Me.MousePointer = vbHourglass
'        Screen.MousePointer = vbHourglass
'        cmdOK_Click 'This is just like saving
'        'Me.MousePointer = vbNormal
'        Screen.MousePointer = vbDefault
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo ErrorTrap
  
  DeleteEvent
  GoTopOfGrid 0, False
TidyUpAndExit:
  ''gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  ''gobjErrorStack.HandleError
End Sub

Private Sub cmdEdit_Click()

  Dim lngRow As Long
  Dim frmFusionComponent As New frmFusionComponent
  Dim ctlGrid As SSDBGrid
  Dim strAddString As String
  Dim strMandatory As Boolean
  Dim strFusionFieldID As String
  Dim strMapToDescription As String
  Dim strIsCompanyCode As String
  Dim strIsEmployeeCode As String
  Dim strIsEmployeeName As String
  Dim strIsDepartmentCode As String
  Dim strIsDepartmentName As String
  Dim strIsFusionCode As String
 
  Set ctlGrid = grdFusionDetails(cboFusionType.ListIndex)
  ctlGrid.Bookmark = ctlGrid.SelBookmarks(0)
  lngRow = ctlGrid.AddItemRowIndex(ctlGrid.Bookmark)
  
  With frmFusionComponent
       
    .BaseTableID = GetComboItem(cboFusionTables)
    .Description = ctlGrid.Columns("Description").Text
    .MapType = val(ctlGrid.Columns("ASRMapType").Text)
    .TableID = val(ctlGrid.Columns("ASRTableID").Text)
    .ColumnID = val(ctlGrid.Columns("ASRColumnID").Text)
    .ExprID = val(ctlGrid.Columns("ASRExprID").Text)
    .value = ctlGrid.Columns("ASRValue").Text
    .IsKeyField = ctlGrid.Columns("IsKeyField").Text
    .IsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    .IsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    .IsDepartmentCode = ctlGrid.Columns("IsDepartmentCode").Text
    .IsFusionCode = ctlGrid.Columns("IsFusionCode").Text
    .IsEmployeeName = ctlGrid.Columns("IsEmployeeName").Text
    .IsDepartmentName = ctlGrid.Columns("IsDepartmentName").Text
    .PreventModify = ctlGrid.Columns("PreventModify").Text
    .DataType = ctlGrid.Columns("DataType").Text
    
    '.Direction = ctlGrid.Columns("Direction").Text
    .AlwaysTransferFieldID = ctlGrid.Columns("AlwaysTransfer").Text
    .ConvertData = ctlGrid.Columns("ConvertData").Text
    .NodeKey = ctlGrid.Columns("FusionFieldID").Text
    .FusionTransferID = GetComboItem(cboFusionType)
    
    strIsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    strIsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    strIsEmployeeName = ctlGrid.Columns("IsEmployeeName").Text
    strIsDepartmentCode = ctlGrid.Columns("IsDepartmentCode").Text
    strIsDepartmentName = ctlGrid.Columns("IsDepartmentName").Text
    strIsFusionCode = ctlGrid.Columns("IsFusionCode").Text
    
    strMandatory = ctlGrid.Columns("Mandatory").Text
    strFusionFieldID = ctlGrid.Columns("FusionFieldID").Text
    
    .Show vbModal
    
    If Not .Cancelled Then

      strMapToDescription = MapToDescription(.MapType, .ColumnID, .ExprID, .value)

      strAddString = .Description & vbTab & strMapToDescription _
          & vbTab & CStr(.MapType) & vbTab & .TableID & vbTab & .ColumnID _
          & vbTab & .ExprID & vbTab & .value & vbTab & strMandatory & vbTab & strFusionFieldID _
          & vbTab & strIsCompanyCode & vbTab & strIsEmployeeCode _
          & vbTab & .Direction & vbTab & .IsKeyField & vbTab & .AlwaysTransferFieldID & vbTab & .ConvertData _
          & vbTab & strIsEmployeeName & vbTab & strIsDepartmentCode & vbTab & strIsDepartmentName & vbTab & strIsFusionCode _
          & vbTab & .Group & vbTab & .PreventModify & vbTab & .DataType
          
      ctlGrid.RemoveItem lngRow
      ctlGrid.AddItem strAddString, lngRow
      ctlGrid.SelBookmarks.RemoveAll
      ctlGrid.SelBookmarks.Add ctlGrid.AddItemBookmark(lngRow)
      
      Changed = True
  
    End If
  
  End With

  RefreshButtons

End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Field Selection Filter' expression selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim lngTableID As Long
  Dim lngFilterID As Long
  
  lngTableID = mavarFusionBaseTableIDs(2, cboFusionType.ListIndex)
  lngFilterID = mavarFusionFilterIDs(cboFusionType.ListIndex)
  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise lngTableID, lngFilterID, giEXPR_STATICFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      Changed = True
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
    End If
    mavarFusionFilterIDs(cboFusionType.ListIndex) = txtFilter.Tag
  End With

TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
End Sub

' Clear the current transfer field
Private Sub cmdNone_Click()
  Dim lngRow As Long
  Dim frmComponent As New frmFusionComponent
  Dim ctlGrid As SSDBGrid
  Dim strAddString As String
  Dim strMandatory As Boolean
  Dim strFusionFieldID As Long
  Dim strMapToDescription As String
  Dim strIsCompanyCode As String
  Dim strIsEmployeeCode As String
  Dim strIsEmployeeName As String
  Dim strIsDepartmentCode As String
  Dim strIsDepartmentName As String
  Dim strIsFusionCode As String
  
  Set ctlGrid = grdFusionDetails(cboFusionType.ListIndex)
  ctlGrid.Bookmark = ctlGrid.SelBookmarks(0)
  lngRow = ctlGrid.AddItemRowIndex(ctlGrid.Bookmark)
  
  With frmComponent
    
    .BaseTableID = GetComboItem(cboFusionTables)
    .Description = ctlGrid.Columns("Description").Text
    .IsKeyField = ctlGrid.Columns("IsKeyField").Text
    .IsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    .IsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    .Direction = ctlGrid.Columns("Direction").Text
    .AlwaysTransferFieldID = ctlGrid.Columns("AlwaysTransfer").Text
    .ConvertData = False
    .NodeKey = ctlGrid.Columns("FusionFieldID").Text
    .FusionTransferID = GetComboItem(cboFusionType)
    
    strIsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    strIsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    strIsEmployeeName = ctlGrid.Columns("IsEmployeeName").Text
    strIsDepartmentCode = ctlGrid.Columns("IsDepartmentCode").Text
    strIsDepartmentName = ctlGrid.Columns("IsDepartmentName").Text
    strIsFusionCode = ctlGrid.Columns("IsFusionCode").Text
        
    strMandatory = ctlGrid.Columns("Mandatory").Text
    strFusionFieldID = ctlGrid.Columns("FusionFieldID").Text
    
    strMapToDescription = MapToDescription(.MapType, .ColumnID, .ExprID, .value)

    strAddString = .Description & vbTab & strMapToDescription _
        & vbTab & "" & vbTab & "" & vbTab & "" _
        & vbTab & "" & vbTab & "" & vbTab & strMandatory & vbTab & strFusionFieldID _
        & vbTab & strIsCompanyCode & vbTab & strIsEmployeeCode _
        & vbTab & .Direction & vbTab & .IsKeyField & vbTab & .AlwaysTransferFieldID & vbTab & .ConvertData _
        & vbTab & strIsEmployeeName & vbTab & strIsDepartmentCode & vbTab & strIsDepartmentName & vbTab & strIsFusionCode _
        & vbTab & "0" & vbTab & "0"
        
    ctlGrid.RemoveItem lngRow
    ctlGrid.AddItem strAddString, lngRow
    ctlGrid.SelBookmarks.RemoveAll
    ctlGrid.SelBookmarks.Add ctlGrid.AddItemBookmark(lngRow)
    
    Changed = True
  
  End With

  RefreshButtons

End Sub

Private Sub cmdOK_Click()

  'AE20071119 Fault #12607
  'If ValidateSetup Then
    'SaveChanges
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If

End Sub

Private Function ValidateSetup() As Boolean

  Dim bOK As Boolean
  Dim strMessage As String
  Dim iCount As Integer
  
  bOK = True

  ValidateSetup = bOK
  
End Function

Private Function SaveChanges() As Boolean

  'AE20071119 Fault #12607
  SaveChanges = False
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  Dim iLoop As Integer
  Dim iLoopTypes As Integer
  Dim varBookMark As Variant
  Dim sSQL As String
  Dim iFusionType As Integer

  With recModuleSetup
    .Index = "idxModuleParameter"

' -------------
' Misc Options
' -------------
    
    ' Save delete prohibit
    .Seek "=", gsMODULEKEY_FUSION, gsPARAMETERKEY_FUSION_ALLOWDELETE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_FUSION
      !parameterkey = gsPARAMETERKEY_FUSION_ALLOWDELETE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = chkAllowDelete.value
    .Update
  
  End With

' --------------------------
' TRANSFER DEFINTION OPTIONS
' --------------------------

  ' Store the transfer types
  daoDb.Execute "DELETE FROM tmpFusionTypes WHERE FusionTypeID IN(" & mstrFusionTypesVisible & ")", dbFailOnError

  For iLoop = LBound(mavarFusionBaseTableIDs, 2) To UBound(mavarFusionBaseTableIDs, 2) - 1
    sSQL = "INSERT INTO tmpFusionTypes" & _
      " (IsVisible, FusionType, FusionTypeID, ASRBaseTableID, FilterID)" & _
      " VALUES (1, " & _
      "'" & CStr(mavarFusionBaseTableIDs(0, iLoop)) & "'," & _
      CStr(mavarFusionBaseTableIDs(1, iLoop)) & "," & _
      CStr(mavarFusionBaseTableIDs(2, iLoop)) & "," & _
      CStr(mavarFusionFilterIDs(iLoop)) & ")"

    daoDb.Execute sSQL, dbFailOnError
  Next iLoop

  ' Store the transfer details
  daoDb.Execute "DELETE FROM tmpFusionFieldDefinitions WHERE FusionTypeID IN(" & mstrFusionTypesVisible & ")", dbFailOnError
  For iLoopTypes = 0 To cboFusionType.ListCount - 1
    With grdFusionDetails(iLoopTypes)
      .Redraw = False
      .MoveFirst
      
      iFusionType = cboFusionType.ItemData(iLoopTypes)
      
      For iLoop = 0 To (.Rows - 1)
  
      sSQL = "INSERT INTO tmpFusionFieldDefinitions" & _
        " (NodeKey, FusionTypeID, Mandatory, Description, ASRMapType, ASRTableID, ASRColumnID, ASRExprID, ASRValue, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ConvertData, IsEmployeeName, IsDepartmentCode, IsDepartmentName, IsFusionCode, DataType, PreventModify)" & _
        " VALUES ('" & _
        .Columns("FusionFieldID").value & "' ," & _
        iFusionType & "," & _
        IIf(.Columns("Mandatory").value = True, "1", "0") & "," & _
        "'" & Replace(.Columns("Description").Text, "'", "''") & "'," & _
        IIf(Len(.Columns("ASRMapType").Text) = 0, "null", .Columns("ASRMapType").Text) & "," & _
        IIf(Len(.Columns("ASRTableID").Text) = 0, "null", .Columns("ASRTableID").Text) & "," & _
        IIf(Len(.Columns("ASRColumnID").Text) = 0, "null", .Columns("ASRColumnID").Text) & "," & _
        IIf(Len(.Columns("ASRExprID").Text) = 0, "null", .Columns("ASRExprID").Text) & "," & _
        "'" & Replace(IIf(Len(.Columns("ASRValue").Text) = 0, "", .Columns("ASRValue").Text), "'", "''") & "'," & _
        IIf(.Columns("IsCompanyCode").Text = True, "1", "0") & "," & _
        IIf(.Columns("IsEmployeeCode").Text = True, "1", "0") & ", " & _
        "0 ," & _
        IIf(.Columns("IsKeyField").Text = True, "1", "0") & "," & _
        IIf(.Columns("AlwaysTransfer").Text = True, "1", "0") & "," & _
        IIf(.Columns("ConvertData").Text = True, "1", "0") & "," & _
        IIf(.Columns("IsEmployeeName").Text = True, "1", "0") & "," & _
        IIf(.Columns("IsDepartmentCode").Text = True, "1", "0") & "," & _
        IIf(.Columns("IsDepartmentName").Text = True, "1", "0") & "," & _
        IIf(.Columns("IsFusionCode").Text = True, "1", "0") & ", " & _
        .Columns("DataType").Text & ", " & _
        IIf(.Columns("PreventModify").Text = True, "1", "0") & ")"
      
        daoDb.Execute sSQL, dbFailOnError
        .MoveNext
      Next iLoop
    End With
  Next iLoopTypes

  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass

  Dim iLoop As Integer
  ReDim mavarFusionBaseTableIDs(2, 0)
  ReDim mavarFusionFilterIDs(0)
  ReDim mabvarFusionForceUpdate(0)
  Dim iCount As Integer
  
  mbReadOnly = (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
  mbLoading = True

  ' Read only mode
  ControlsDisableAll Me, Not mbReadOnly
  cmdDelete.Enabled = False
  txtFilter.Enabled = False
  txtFilter.BackColor = vbButtonFace
  cboFusionType.Enabled = True
  cboFusionType.BackColor = vbWhite
  cboFusionType.ForeColor = vbBlack
  For iCount = grdFusionDetails.LBound To grdFusionDetails.UBound
    grdFusionDetails(iCount).Enabled = True
  Next iCount

  ' Don't need this stuff for phase I
  tabOptions.TabVisible(0) = True
  tabOptions.TabVisible(1) = False
  tabOptions.TabVisible(2) = False

  PopulateBaseTables
  ReadParameters
  PopulateFusionTransferTypes
  
  ' Load the transfer types
  For iLoop = 0 To cboFusionType.ListCount - 1
    PopulateFusionTransferDetails iLoop, False
  Next iLoop

  PopulateFields
  EnableDisableTabControls

  mbLoading = False
  RefreshButtons
  
  mfChanged = False
  cmdDelete.Enabled = Not mbReadOnly
  cmdOk.Enabled = False

  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Changed Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        Cancel = (Not SaveChanges)
    End Select
  End If
End Sub

Private Sub grdFusionDetails_Click(Index As Integer)
  Dim iCount As Integer
  
  If grdFusionDetails.Item(0).SelBookmarks.Count > 1 Or cboFusionTables = "<None>" Then
    cmdEdit.Enabled = False
  Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = Not mbReadOnly
  End If

End Sub

Private Sub grdFusionDetails_DblClick(Index As Integer)
  ' Display the properties form for the current transfer definition
  If cmdEdit.Enabled Then
    cmdEdit_Click
  End If
End Sub

Private Sub grdFusionDetails_RowLoaded(Index As Integer, ByVal Bookmark As Variant)

  Dim iCount As Integer
  Dim strType As String

'  If grdFusionDetails(Index).Columns("Mandatory").Text = "True" Then
'    strType = "Mandatory"
'  End If
'
'  If grdFusionDetails(Index).Columns("IsKeyField").Text = "True" Then
'    strType = "Mandatory"
'  End If
'
'  If strType <> "" Then
'    For iCount = 0 To grdFusionDetails(Index).Columns.Count - 1
'      grdFusionDetails(Index).Columns(iCount).CellStyleSet strType
'    Next iCount
'  End If
End Sub

Private Sub tabOptions_Click(PreviousTab As Integer)
  EnableDisableTabControls
End Sub

Private Sub ReadParameters()

  Dim strEncypted As String

  ' Get the configured Personnel table ID and Personnel table view ID.
  With recModuleSetup
    .Index = "idxModuleParameter"
      
' -------------
' MISC OPTIONS
' -------------
  
    ' Get allow deletions
    .Seek "=", gsMODULEKEY_FUSION, gsPARAMETERKEY_FUSION_ALLOWDELETE
    If .NoMatch Then
      .Seek "=", gsMODULEKEY_FUSION, gsPARAMETERKEY_FUSION_ALLOWDELETE
      If .NoMatch Then
        mbAllowDeletions = False
      Else
        mbAllowDeletions = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, !parametervalue)
      End If
    Else
      mbAllowDeletions = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, !parametervalue)
    End If
    
  End With

End Sub

Private Sub PopulateFields()

  ' Options Tab
  chkAllowDelete.value = IIf(mbAllowDeletions = True, vbChecked, vbUnchecked)

End Sub

Private Function SelectedComboItem(cboTemp As ComboBox) As Long
  With cboTemp
    If .ListIndex >= 0 Then
      SelectedComboItem = .ItemData(.ListIndex)
    Else
      SelectedComboItem = 0
    End If
  End With
End Function

Private Sub PopulateFusionTransferTypes()

  Dim rsFusionTypes As DAO.Recordset
  Dim sSQL As String

  sSQL = "SELECT FusionType, FusionTypeID, FilterID, ASRBaseTableID FROM tmpFusionTypes" _
      & " WHERE IsVisible = true" _
      & " ORDER BY FusionTypeID"
  Set rsFusionTypes = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  With rsFusionTypes
    While Not .EOF
      
      mavarFusionBaseTableIDs(0, UBound(mavarFusionBaseTableIDs, 2)) = Trim(!FusionType)
      mavarFusionBaseTableIDs(1, UBound(mavarFusionBaseTableIDs, 2)) = !FusionTypeID
      mavarFusionBaseTableIDs(2, UBound(mavarFusionBaseTableIDs, 2)) = !ASRBaseTableID
      ReDim Preserve mavarFusionBaseTableIDs(2, UBound(mavarFusionBaseTableIDs, 2) + 1)
      
      AddItemToComboBox cboFusionType, !FusionType, !FusionTypeID
      
      ' Filter information
      mavarFusionFilterIDs(UBound(mavarFusionFilterIDs)) = !FilterID
      ReDim Preserve mavarFusionFilterIDs(UBound(mavarFusionFilterIDs) + 1)
           
      ' Remember the visible types
      mstrFusionTypesVisible = mstrFusionTypesVisible & IIf(LenB(mstrFusionTypesVisible) <> 0, ",", "") & Trim(!FusionTypeID)
      
      .MoveNext
    Wend
   
    .Close
  End With
       
  Set rsFusionTypes = Nothing

  ' Set to the top
  cboFusionType.ListIndex = 0

End Sub

Private Sub PopulateBaseTables()
  ' Populate the tables combo.
  
  ' Clear the combo.
  cboFusionTables.Clear
  cboFusionTables.AddItem "<None>"
  
  With recTabEdit
    .Index = "idxName"
    .MoveFirst
    
    Do While Not .EOF

      If Not !Deleted Then
        AddItemToComboBox cboFusionTables, !TableName, !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
End Sub

' Value of map transfer
Private Function MapToDescription(piMapType As SystemMgr.FusionMapType _
  , plngColumnID As Long, plngExprID As Long, pstrValue As String) As String
  
  Select Case piMapType
    Case FUSION_MAPTYPE_COLUMN
      MapToDescription = GetColumnName(plngColumnID)
    
    Case FUSION_MAPTYPE_EXPRESSION
      MapToDescription = GetExpressionName(plngExprID)
    
    Case FUSION_MAPTYPE_VALUE
      MapToDescription = "'" & Trim(pstrValue) & "'"
  
  End Select
  
End Function

Private Sub PopulateFusionTransferDetails(ByVal plngFusionGrid As Long, pbReset As Boolean)

  Dim sSQL As String
  Dim strAddString As String
  Dim strMapToDescription As String
  Dim rsDefinition As DAO.Recordset
  Dim ctlGrid As SSDBGrid
  Dim iFusionTypeID As Integer

  iFusionTypeID = cboFusionType.ItemData(plngFusionGrid)

  ' Unload grid if resetting
  If pbReset Then
    grdFusionDetails(plngFusionGrid).RemoveAll
  Else

    ' Load up a grid for this definition
    If plngFusionGrid > 0 Then
      Load grdFusionDetails(plngFusionGrid)
      grdFusionDetails(plngFusionGrid).RemoveAll
    End If
  End If

  sSQL = "SELECT *" & _
    " FROM tmpFusionFieldDefinitions" & _
    " WHERE FusionTypeID = " & CStr(iFusionTypeID) & _
    " ORDER BY Mandatory, NodeKey"
    
  Set rsDefinition = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsDefinition.EOF
    
    If Not pbReset Then
      If IsNull(rsDefinition!ASRMapType) Then
        strMapToDescription = ""
      Else
        strMapToDescription = MapToDescription(rsDefinition!ASRMapType, rsDefinition!ASRColumnID, rsDefinition!ASRExprID, rsDefinition!ASRValue)
      End If
    Else
      strMapToDescription = ""
    End If
    
    strAddString = rsDefinition!Description & vbTab & strMapToDescription
    
    If pbReset Then
      strAddString = strAddString & vbTab & vbTab & vbTab & vbTab & vbTab
    Else
      strAddString = strAddString _
          & vbTab & rsDefinition!ASRMapType & vbTab & rsDefinition!ASRTableID & vbTab & rsDefinition!ASRColumnID _
          & vbTab & rsDefinition!ASRExprID & vbTab & Trim(rsDefinition!ASRValue)
    End If
               
    strAddString = strAddString _
        & vbTab & rsDefinition!Mandatory & vbTab & rsDefinition!NodeKey _
        & vbTab & rsDefinition!IsCompanyCode & vbTab & rsDefinition!IsEmployeeCode _
        & vbTab & "" _
        & vbTab & rsDefinition!AlwaysTransfer _
        & vbTab & IIf(IsNull(rsDefinition!ConvertData), False, rsDefinition!ConvertData) _
        & vbTab & IIf(IsNull(rsDefinition!IsEmployeeName), False, rsDefinition!IsEmployeeName) _
        & vbTab & IIf(IsNull(rsDefinition!IsDepartmentCode), False, rsDefinition!IsDepartmentCode) _
        & vbTab & IIf(IsNull(rsDefinition!IsDepartmentName), False, rsDefinition!IsDepartmentName) _
        & vbTab & IIf(IsNull(rsDefinition!IsFusionCode), False, rsDefinition!IsFusionCode) _
        & vbTab & 0 & vbTab & 0 & vbTab & 0 _
        & vbTab & IIf(IsNull(rsDefinition!DataType), 0, rsDefinition!DataType)
        
    grdFusionDetails(plngFusionGrid).AddItem strAddString
    rsDefinition.MoveNext
    
  Wend
  GoTopOfGrid plngFusionGrid, (cboFusionTables = "<None>")
  cmdEdit.Enabled = (cboFusionTables = "<None>")
  
  rsDefinition.Close
  Set rsDefinition = Nothing

End Sub

Private Function GetComboItem(cboTemp As ComboBox) As Long
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
End Function

Private Sub EnableDisableTabControls()

  cmdOk.Enabled = Not mbReadOnly And mfChanged
  cmdNone.Enabled = (tabOptions.Tab = 0) And Not mbReadOnly
  cmdEdit.Caption = IIf(mbReadOnly, "&View...", "&Edit...")
  cmdFilter.Enabled = (tabOptions.Tab = 0) And Not mbReadOnly

  fraFusionDefinition.Enabled = (tabOptions.Tab = 0)
  fraDefaults.Enabled = (tabOptions.Tab = 1) And Not mbReadOnly

  RefreshButtons

End Sub

Private Sub DeleteEvent()
  Dim strEventIDs  As String
  Dim plngLoop As Long
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim arrayBookmarks() As Variant
  Dim iFusionType As Integer
  Dim iAnswer As Integer
  On Error GoTo ErrorTrap
  
  iFusionType = cboFusionType.ListIndex
  
  iAnswer = MsgBox("Are you sure you want to clear the selected field(s)?", vbYesNo + vbQuestion, Me.Caption)
  If iAnswer = vbYes Then

    Screen.MousePointer = vbHourglass
    'Workout how many records have been selected
    nTotalSelRows = grdFusionDetails(iFusionType).SelBookmarks.Count
    'Redimension the arrays to the count of the bookmarks
    ReDim arrayBookmarks(nTotalSelRows)
    
    For intCount = 1 To nTotalSelRows
      arrayBookmarks(intCount) = grdFusionDetails(iFusionType).SelBookmarks.Item(intCount - 1)
    Next intCount
    
    For intCount = 1 To nTotalSelRows
      grdFusionDetails(iFusionType).Bookmark = arrayBookmarks(intCount)
      'Clear this bookmarked row
      If Len(strEventIDs) > 0 Then
        strEventIDs = strEventIDs & ","
      End If
      ClearItem (CLng(grdFusionDetails(iFusionType).AddItemRowIndex(grdFusionDetails(iFusionType).Bookmark)))
    Next intCount
    
    grdFusionDetails(iFusionType).SelBookmarks.RemoveAll
    'Go to the top one
    Screen.MousePointer = vbDefault
  End If
  
  'UnLoad frmDeleteSelection
  RefreshButtons

TidyUpAndExit:
  'gobjErrorStack.PopStack
  Exit Sub
  
ErrorTrap:
  'gobjErrorStack.HandleError
  
End Sub

' Clear the current transfer field
Private Sub ClearItem(lngrow2 As Long)
  Dim lngRow As Long
  Dim frmComponent As New frmFusionComponent
  Dim ctlGrid As SSDBGrid
  Dim strAddString As String
  Dim strMandatory As Boolean
  Dim strFusionFieldID As Long
  Dim strMapToDescription As String
  Dim strIsCompanyCode As String
  Dim strIsEmployeeCode As String
  Dim strIsEmployeeName As String
  Dim strIsDepartmentCode As String
  Dim strIsDepartmentName As String
  Dim strIsFusionCode As String
  
  Set ctlGrid = grdFusionDetails(cboFusionType.ListIndex)
'  ctlGrid.Bookmark = ctlGrid.SelBookmarks(0)
'  lngrow = ctlGrid.AddItemRowIndex(ctlGrid.Bookmark)
  
  lngRow = lngrow2
  With frmComponent
    
    .BaseTableID = GetComboItem(cboFusionTables)
    .Description = ctlGrid.Columns("Description").Text
    .IsKeyField = ctlGrid.Columns("IsKeyField").Text
    .IsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    .IsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    .Direction = ctlGrid.Columns("Direction").Text
    .AlwaysTransferFieldID = ctlGrid.Columns("AlwaysTransfer").Text
    .ConvertData = False
    .NodeKey = ctlGrid.Columns("FusionFieldID").Text
    .FusionTransferID = GetComboItem(cboFusionType)
    
    strIsCompanyCode = ctlGrid.Columns("IsCompanyCode").Text
    strIsEmployeeCode = ctlGrid.Columns("IsEmployeeCode").Text
    strIsEmployeeName = ctlGrid.Columns("IsEmployeeName").Text
    strIsDepartmentCode = ctlGrid.Columns("IsDepartmentCode").Text
    strIsDepartmentName = ctlGrid.Columns("IsDepartmentName").Text
    strIsFusionCode = ctlGrid.Columns("IsFusionCode").Text
        
    strMandatory = ctlGrid.Columns("Mandatory").Text
    strFusionFieldID = ctlGrid.Columns("FusionFieldID").Text
    
    strMapToDescription = MapToDescription(.MapType, .ColumnID, .ExprID, .value)

    strAddString = .Description & vbTab & strMapToDescription _
        & vbTab & "" & vbTab & "" & vbTab & "" _
        & vbTab & "" & vbTab & "" & vbTab & strMandatory & vbTab & strFusionFieldID _
        & vbTab & strIsCompanyCode & vbTab & strIsEmployeeCode _
        & vbTab & .Direction & vbTab & .IsKeyField & vbTab & .AlwaysTransferFieldID & vbTab & .ConvertData _
        & vbTab & strIsEmployeeName & vbTab & strIsDepartmentCode & vbTab & strIsDepartmentName & vbTab & strIsFusionCode _
        & vbTab & "0" & vbTab & "0"
        
    ctlGrid.RemoveItem lngRow
    ctlGrid.AddItem strAddString, lngRow
    ctlGrid.SelBookmarks.RemoveAll
    ctlGrid.SelBookmarks.Add ctlGrid.AddItemBookmark(lngRow)
    
    Changed = True
  
  End With

  RefreshButtons

End Sub

