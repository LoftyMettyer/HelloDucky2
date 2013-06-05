VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDataTransfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Transfer Definition"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1026
   Icon            =   "frmDataTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4670
      Left            =   90
      TabIndex        =   28
      Top             =   90
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Definition"
      TabPicture(0)   =   "frmDataTransfer.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinition(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDefinition(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmDataTransfer.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumnDefinition"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraColumnDefinition 
         Enabled         =   0   'False
         Height          =   4140
         Left            =   -74865
         TabIndex        =   20
         Top             =   360
         Width           =   9400
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   7950
            TabIndex        =   25
            Top             =   1935
            Width           =   1200
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   24
            Top             =   1395
            Width           =   1200
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   23
            Top             =   840
            Width           =   1200
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7950
            TabIndex        =   22
            Top             =   315
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   3650
            Left            =   210
            TabIndex        =   21
            Top             =   315
            Width           =   7440
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   9
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
            RowNavigation   =   1
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   9
            Columns(0).Width=   3016
            Columns(0).Caption=   "Source Table"
            Columns(0).Name =   "Source Table"
            Columns(0).CaptionAlignment=   0
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   250
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "SourceTableID"
            Columns(1).Name =   "SourceTableID"
            Columns(1).AllowSizing=   0   'False
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   2963
            Columns(2).Caption=   "Source Column"
            Columns(2).Name =   "Source Column"
            Columns(2).CaptionAlignment=   0
            Columns(2).AllowSizing=   0   'False
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(2).HasForeColor=   -1  'True
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "SourceColumnID"
            Columns(3).Name =   "SourceColumnID"
            Columns(3).AllowSizing=   0   'False
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   3
            Columns(3).FieldLen=   256
            Columns(3).Locked=   -1  'True
            Columns(4).Width=   688
            Columns(4).Name =   "To"
            Columns(4).Alignment=   2
            Columns(4).CaptionAlignment=   2
            Columns(4).AllowSizing=   0   'False
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).Case =   2
            Columns(4).FieldLen=   256
            Columns(4).Locked=   -1  'True
            Columns(4).HasBackColor=   -1  'True
            Columns(4).BackColor=   -2147483633
            Columns(5).Width=   2963
            Columns(5).Caption=   "Destination Table"
            Columns(5).Name =   "Destination Table"
            Columns(5).CaptionAlignment=   0
            Columns(5).AllowSizing=   0   'False
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(5).Locked=   -1  'True
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "DestinationTableID"
            Columns(6).Name =   "DestinationTableID"
            Columns(6).AllowSizing=   0   'False
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   3
            Columns(6).FieldLen=   256
            Columns(6).Locked=   -1  'True
            Columns(7).Width=   3043
            Columns(7).Caption=   "Destination Column"
            Columns(7).Name =   "Destination Column"
            Columns(7).CaptionAlignment=   0
            Columns(7).AllowSizing=   0   'False
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(7).Locked=   -1  'True
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "DestinationColumnID"
            Columns(8).Name =   "DestinationColumnID"
            Columns(8).AllowSizing=   0   'False
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   3
            Columns(8).FieldLen=   256
            Columns(8).Locked=   -1  'True
            TabNavigation   =   1
            _ExtentX        =   13123
            _ExtentY        =   6438
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
      End
      Begin VB.Frame fraDefinition 
         Caption         =   "Data :"
         Height          =   2150
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   2350
         Width           =   9400
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1080
            Width           =   1750
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   705
            Width           =   1750
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   6000
            TabIndex        =   15
            Top             =   1120
            Width           =   1065
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   6000
            TabIndex        =   12
            Top             =   750
            Width           =   1020
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   6000
            TabIndex        =   11
            Top             =   365
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.ComboBox cboFromTable 
            Height          =   315
            ItemData        =   "frmDataTransfer.frx":0044
            Left            =   1620
            List            =   "frmDataTransfer.frx":0046
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   315
            Width           =   3000
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8895
            TabIndex        =   14
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8895
            TabIndex        =   17
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboToTable 
            Height          =   315
            ItemData        =   "frmDataTransfer.frx":0048
            Left            =   1620
            List            =   "frmDataTransfer.frx":004A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1575
            Width           =   3000
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Source :"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Index           =   5
            Left            =   5010
            TabIndex        =   10
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination :"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   1635
            Width           =   915
         End
      End
      Begin VB.Frame fraDefinition 
         Height          =   1950
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   360
         Width           =   9400
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5865
            MaxLength       =   30
            TabIndex        =   5
            Top             =   315
            Width           =   3360
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   1
            Top             =   315
            Width           =   3000
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1620
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   705
            Width           =   3000
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5850
            TabIndex        =   30
            Top             =   720
            Width           =   3405
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            stylesets.count =   2
            stylesets(0).Name=   "SysSecMgr"
            stylesets(0).ForeColor=   -2147483631
            stylesets(0).BackColor=   -2147483633
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
            stylesets(0).Picture=   "frmDataTransfer.frx":004C
            stylesets(1).Name=   "ReadOnly"
            stylesets(1).ForeColor=   -2147483631
            stylesets(1).BackColor=   -2147483633
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
            stylesets(1).Picture=   "frmDataTransfer.frx":0068
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
            SelectTypeRow   =   0
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   2963
            Columns(0).Caption=   "User Group"
            Columns(0).Name =   "GroupName"
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   2566
            Columns(1).Caption=   "Access"
            Columns(1).Name =   "Access"
            Columns(1).AllowSizing=   0   'False
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(1).Style=   3
            Columns(1).Row.Count=   3
            Columns(1).Col.Count=   2
            Columns(1).Row(0).Col(0)=   "Read / Write"
            Columns(1).Row(1).Col(0)=   "Read Only"
            Columns(1).Row(2).Col(0)=   "Hidden"
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "SysSecMgr"
            Columns(2).Name =   "SysSecMgr"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   6006
            _ExtentY        =   1905
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   2
            Left            =   5010
            TabIndex        =   4
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   0
            Top             =   365
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   2
            Top             =   750
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   5010
            TabIndex        =   6
            Top             =   810
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7305
      TabIndex        =   26
      Top             =   4900
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8565
      TabIndex        =   27
      Top             =   4900
      Width           =   1200
   End
End
Attribute VB_Name = "frmDataTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fOK As Boolean
Private msFromTable As String
Private msToTable As String
Private mbCancelled As Boolean

Private mlTransferID As Long
Private mlTimeStamp As Long
Private mbFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mbLoading As Boolean
Private mblnReadOnly As Boolean
Private datData As HRProDataMgr.clsDataAccess
'Private mbChanged As Boolean
Private mblnDefinitionCreator As Boolean

Const giTABSTRIP_DATATRANSFERDEF = 0
Const giTABSTRIP_COLUMNDEF = 1

Private mblnForceHidden As Boolean

Public Property Get SelectedID() As Long
  SelectedID = mlTransferID
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property

Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lTransferID As Long, Optional bPrint As Boolean) As Boolean
    
  Set datData = New HRProDataMgr.clsDataAccess
  
  Screen.MousePointer = vbHourglass
  
  ' Populate the combos.
  LoadCombos
  fOK = True

  EnableAll
   
  If bNew Then
    grdColumns.RemoveAll
    txtName = ""
    txtUserName = gsUserName
    mblnDefinitionCreator = True
    optAllRecords.Value = True
    txtPicklist.Text = ""
    txtFilter.Text = ""
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    cmdFilter.Enabled = False
    cmdPicklist.Enabled = False
    mlTransferID = 0
  
    PopulateAccessGrid
    
    Me.Changed = False
  
  Else
    mlTransferID = lTransferID
    FromCopy = bCopy

    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
  
    PopulateAccessGrid

    RetreiveDefinition
  
    If Me.Cancelled Then Exit Function
    If fOK Then
      Me.Changed = False
    
      If FromCopy Then
        mlTransferID = 0
        Me.Changed = True
      Else
        'Me.Changed = (Not mblnReadOnly)
        Me.Changed = False
      End If
    End If
  End If
  
  CheckIfScrollBarRequired
  Cancelled = False
  SSTab1.Tab = giTABSTRIP_DATATRANSFERDEF
  
  Initialise = fOK
  Screen.MousePointer = vbDefault
  
End Function

Private Sub PopulateAccessGrid()
  ' Populate the access grid.
  Dim rsAccess As ADODB.Recordset
  
  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With
  
  ' Get the recordset of user groups and their access on this definition.
  Set rsAccess = GetUtilityAccessRecords(utlDataTransfer, mlTransferID, mbFromCopy)
  If Not rsAccess Is Nothing Then
    ' Add the user groups and their access on this definition to the access grid.
    With rsAccess
      Do While Not .EOF
        grdAccess.AddItem !Name & vbTab & AccessDescription(!Access) & vbTab & !sysSecMgr
        
        .MoveNext
      Loop
    
      .Close
    End With
  End If
  Set rsAccess = Nothing

End Sub


Private Sub LoadCombos()
  ' Populate the combos.
  Dim rsTables As New Recordset
  Dim sSQL As String

'  sSQL = "Select TableName, TableID From ASRSysTables"
'  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  mbLoading = True
'  With cboFromTable
'    ' Clear the combos.
'    .Clear
'    'cboToTable.Clear
'
'    ' Add the tables to the combos.
'    Do While Not rsTables.EOF
'      .AddItem rsTables!TableName
'      .ItemData(.NewIndex) = rsTables!TableID
'      'cboToTable.AddItem rsTables!TableName
'      'cboToTable.ItemData(cboToTable.NewIndex) = rsTables!TableID
'      rsTables.MoveNext
'    Loop
'
'    If .ListCount > 0 Then
'      .ListIndex = 0
'      'cboToTable.ListIndex = 0
'    End If
'  End With

  
  'Default source to whichever table is first in the list
  'and default destination to personnel, if possible.
  LoadTableCombo cboFromTable
  If cboFromTable.ListCount > 0 Then
    cboFromTable.ListIndex = 0
  End If

  Call PopulateDestination
  'Call PopulateOrderCombo

  mbLoading = False
    
'  rsTables.Close
'  Set rsTables = Nothing

End Sub

Private Sub cboFromTable_DropDown()
  msFromTable = cboFromTable.Text
End Sub

Private Sub cboFromTable_KeyDown(KeyCode As Integer, Shift As Integer)
  msFromTable = cboFromTable.Text
End Sub


Private Sub cboFromTable_Click()
  Call CheckChangeTable(cboFromTable, msFromTable, True)
End Sub

Private Sub cboToTable_Click()
  Call CheckChangeTable(cboToTable, msToTable, False)
End Sub

Private Sub CheckChangeTable(cboTable As ComboBox, strOldValue As String, blnFromCombo)

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  
  If mbLoading Or (cboTable.Text = strOldValue) Then
    Exit Sub
  End If
  
  If grdColumns.Rows > 0 Then
    
    strMBText = "Warning: Changing the base table will result in all table/column " & _
            "specific aspects of this data transfer definition being cleared." & vbCrLf & _
            "Are you sure you wish to continue?"
    intMBButtons = vbQuestion + vbYesNo
    strMBTitle = "Data Transfer"
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
    If intMBResponse <> vbYes Then
      SetComboText cboTable, strOldValue
      Exit Sub
    End If
    
  End If


  grdColumns.RemoveAll
  CheckIfScrollBarRequired
  
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdClearAll.Enabled = False

  If blnFromCombo = True Then
    
    optAllRecords.Value = True
    txtPicklist.Text = ""
    txtPicklist.Tag = ""
    txtFilter.Text = ""
    txtFilter.Tag = ""

    mbLoading = True
    Call PopulateDestination
    mbLoading = False

  End If
  
  Me.Changed = True

End Sub
  
  
'Private Sub PopulateOrderCombo()
'
'  Dim rsTables As Recordset
'  Dim sSQL As String
'
'  sSQL = "SELECT Name,OrderID FROM ASRSysOrders WHERE TableID = " & cboToTable.ItemData(cboToTable.ListIndex)
'  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'  With cboOrder
'    .Clear
'    Do While Not rsTables.EOF
'      .AddItem rsTables!Name
'      .ItemData(.NewIndex) = rsTables!OrderID
'      rsTables.MoveNext
'    Loop
'    If .ListCount > 0 Then
'      .ListIndex = 0
'    End If
'  End With
'
'End Sub


Private Sub cboToTable_DropDown()
  msToTable = cboToTable.Text
End Sub

Private Sub cboToTable_KeyDown(KeyCode As Integer, Shift As Integer)
  msToTable = cboToTable.Text
End Sub


Private Sub cmdCancel_Click()
 
  Dim strSQL As String
  
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  If Me.Changed And Not mblnReadOnly Then
    
    'strMBText = "Data Transfer definition has changed.  Save changes ?"
    strMBText = "You have changed the current definition. Save changes ?"
    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
    strMBTitle = "Data Transfer"
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
    Select Case intMBResponse
    Case vbYes
      If Not SaveDefinition Then
        Exit Sub
      End If
    Case vbCancel
      Exit Sub
    End Select
  End If

  ' Exit without saving the definition.
  Me.Hide
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdClearAll_Click()

  If COAMsgBox("Are you sure you want to clear the column selections ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Me.Changed = True
    grdColumns.RemoveAll
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    CheckIfScrollBarRequired
  End If
  
End Sub

Private Sub cmdDelete_Click()

  Dim lRow As Long

  'MH20011024 Fault 3014
  'Removed confirmation as per request from PJC
  'If COAMsgBox("Are you sure you wish to delete the current row ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Me.Changed = True

    With grdColumns

      'lRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
      'lRow = .AddItemRowIndex(.SelBookmarks.Item(0))
      'grdColumns.RemoveItem lRow

      If .Rows = 1 Then
        .RemoveAll
      Else
        lRow = .AddItemRowIndex(.Bookmark)
        .RemoveItem lRow
        If lRow < .Rows Then
          .Bookmark = .AddItemBookmark(lRow)
        Else
          .Bookmark = .AddItemBookmark(.Rows - 1)
        End If
        .SelBookmarks.Add .Bookmark
      End If

    End With
    
    If grdColumns.Rows = 0 Then
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      cmdClearAll.Enabled = False
    End If
    CheckIfScrollBarRequired
  'End If
  
End Sub

Private Sub cmdEdit_Click()
  
  ' Modify the selected column transfer definition.
  Dim sRow As String
  Dim lRow As Long
  Dim frmEdit As frmDataTransferColumn


  Screen.MousePointer = vbHourglass
  Set frmEdit = New frmDataTransferColumn
  
  With grdColumns
    
    lRow = grdColumns.AddItemRowIndex(.SelBookmarks(0))
    grdColumns.Bookmark = grdColumns.AddItemBookmark(lRow)
    
    frmEdit.ParentForm = Me
    If .Columns(1).Value > 0 Then
      'Has a table id
      frmEdit.Initialise False, cboFromTable.ItemData(cboFromTable.ListIndex), _
        cboToTable.ItemData(cboToTable.ListIndex), .Columns(0).Text, .Columns(2).Text, _
        .Columns(5).Text, .Columns(7).Text
    Else
      'Has no table id, so pass text
      frmEdit.Initialise False, cboFromTable.ItemData(cboFromTable.ListIndex), _
        cboToTable.ItemData(cboToTable.ListIndex), , .Columns(2).Text, .Columns(5).Text, _
        .Columns(7).Text, Mid$(.Columns(0).Text, 2, Len(.Columns(0).Text) - 2)
    End If
  End With
    
  With frmEdit
    .Show vbModal
    
    grdColumns.Bookmark = grdColumns.AddItemBookmark(lRow)
    
    If Not .Cancelled Then
      Me.Changed = True
      
      grdColumns.Columns(1).Text = "0"
      grdColumns.Columns(2).Text = vbNullString
      grdColumns.Columns(3).Text = vbNullString
      
      If .optSystemDate Then
        'sRow = "<System Date>" & vbTab & 0 & vbTab & vbTab
        grdColumns.Columns(0).Text = "<System Date>"
      ElseIf .optText Then
        'sRow = Chr$(34) & .txtOther & Chr$(34) & vbTab & 0 & vbTab & vbTab
        grdColumns.Columns(0).Text = Chr$(34) & .txtOther & Chr$(34)
      Else
        'sRow = .cboFromTable & vbTab & .cboFromTable.ItemData(.cboFromTable.ListIndex) & vbTab
        'sRow = sRow & .cboFromColumn & vbTab & .cboFromColumn.ItemData(.cboFromColumn.ListIndex)
        grdColumns.Columns(0).Text = .cboFromTable
        grdColumns.Columns(1).Text = .cboFromTable.ItemData(.cboFromTable.ListIndex)
        grdColumns.Columns(2).Text = .cboFromColumn
        grdColumns.Columns(3).Text = .cboFromColumn.ItemData(.cboFromColumn.ListIndex)
      End If
      
      'sRow = sRow & vbTab & "TO " & vbTab & .cboToTable & vbTab & .cboToTable.ItemData(.cboToTable.ListIndex)
      'sRow = sRow & vbTab & .cboToColumn & vbTab & .cboToColumn.ItemData(.cboToColumn.ListIndex)
      grdColumns.Columns(4).Text = "TO"
      grdColumns.Columns(5).Text = .cboToTable
      grdColumns.Columns(6).Text = .cboToTable.ItemData(.cboToTable.ListIndex)
      grdColumns.Columns(7).Text = .cboToColumn
      grdColumns.Columns(8).Text = .cboToColumn.ItemData(.cboToColumn.ListIndex)
      
      'With grdColumns
      '  .RemoveItem lRow
      '  .AddItem sRow, lRow
      '  .SelBookmarks.Add .AddItemBookmark(lRow)
      'End With
    End If
  End With
  
  Unload frmEdit
  Set frmEdit = Nothing

End Sub

Private Sub cmdFilter_Click()
  
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim objExpression As clsExprExpression
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    If .Initialise(cboFromTable.ItemData(cboFromTable.ListIndex), Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
    
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) Then
      
        Me.Changed = True
      
        ' Read the selected expression info.
        txtFilter.Text = .Name
        txtFilter.Tag = .ExpressionID
      End If
    
    End If
  End With
  
  Set objExpression = Nothing

  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub cmdNew_Click()
  
  ' Add a new transfer column definition.
  Dim sRow As String
  Dim frmEdit As frmDataTransferColumn
  Dim lRow As Long
  
  lRow = grdColumns.AddItemRowIndex(grdColumns.SelBookmarks(0))
  grdColumns.Bookmark = grdColumns.AddItemBookmark(lRow)
  
  
  Screen.MousePointer = vbHourglass
  
  Set frmEdit = New frmDataTransferColumn
  With frmEdit
    ' Display the column transfer definition form.
    .ParentForm = Me
    .Initialise True, cboFromTable.ItemData(cboFromTable.ListIndex), cboToTable.ItemData(cboToTable.ListIndex)
    .Show vbModal
    
    If Not .Cancelled Then
      ' Add a new row to the grid with the table and column id's in the hidden columns to use later
      ' Note: If a system date or plain text has been used in the from clause,
      ' then the from tableid columnid = 0
      Me.Changed = True
      If .optSystemDate Then
        sRow = "<System Date>" & vbTab & vbTab & vbTab
      ElseIf .optText Then
        sRow = Chr$(34) & .txtOther & Chr$(34) & vbTab & vbTab & vbTab
      Else
        sRow = .cboFromTable & vbTab & .cboFromTable.ItemData(.cboFromTable.ListIndex) & vbTab
        sRow = sRow & .cboFromColumn & vbTab & .cboFromColumn.ItemData(.cboFromColumn.ListIndex)
      End If
      
      sRow = sRow & vbTab & "TO " & vbTab & .cboToTable & vbTab & .cboToTable.ItemData(.cboToTable.ListIndex)
      sRow = sRow & vbTab & .cboToColumn & vbTab & .cboToColumn.ItemData(.cboToColumn.ListIndex)
      
      With grdColumns
        .AddItem sRow
        .MoveLast
        .SelBookmarks.Add .Bookmark
      End With
      
      
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      cmdClearAll.Enabled = True
      
    'Else
    '  grdColumns.Bookmark = grdColumns.AddItemBookmark(lRow)
      
    End If
  End With

  CheckIfScrollBarRequired
  Unload frmEdit
  Set frmEdit = Nothing
  
End Sub

Private Sub cmdOK_Click()
  ' Save the definition and exit.

  If SaveDefinition Then
    Cancelled = True
    Unload Me
  End If

End Sub

Private Sub cmdPicklist_Click()
  ' Display the picklist selection form.
  'Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim frmSelection As frmDefSel
  Dim lFrom As Long
  Dim lTo As Long
  
  Dim rsTemp As Recordset
  Dim blnHiddenPicklist As Boolean
  
  Screen.MousePointer = vbHourglass
  
  ' Construct the SQL string for selecting the picklists on the current table.
  'sSQL = "SELECT name, pickListID FROM ASRSysPickListName"
  'sSQL = sSQL & " WHERE tableID = " & cboFromTable.ItemData(cboFromTable.ListIndex)
  'If IsChildTable(cboFromTable.ItemData(cboFromTable.ListIndex), lParent) Then
  '  sSQL = sSQL & " OR tableID = " & lParent
  'Else
  '  sSQL = "tableID = " & cboFromTable.ItemData(cboFromTable.ListIndex)
  'End If

  Set frmSelection = New frmDefSel
  With frmSelection

    '23/07/2001 MH Fault 2585
    'Only select picklist on "From" table.
    'lFrom = cboFromTable.ItemData(cboFromTable.ListIndex)
    'lTo = cboToTable.ItemData(cboToTable.ListIndex)
    'lParent = GetCommonParent(lFrom, lTo)
    'If lParent = 0 Then
    '  'If not child to child then just get picklist on from table.
    '  .TableID = cboFromTable.ItemData(cboFromTable.ListIndex)
    'Else
    '  .TableID = lParent
    'End If
      
    lFrom = cboFromTable.ItemData(cboFromTable.ListIndex)
    lTo = cboToTable.ItemData(cboToTable.ListIndex)
    .TableID = cboFromTable.ItemData(cboFromTable.ListIndex)

    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(txtPicklist.Tag) > 0 Then
      .SelectedID = Val(txtPicklist.Tag)
    End If
    
    Do While Not fExit
      
      If .ShowList(utlPicklist) Then
          
        .Show vbModal
          
        Select Case .Action
          Case edtAdd
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(True, False, cboFromTable.ItemData(cboFromTable.ListIndex)) Then
              frmPick.Show vbModal
            End If
            frmSelection.SelectedID = frmPick.SelectedID
            Set frmPick = Nothing
          
          Case edtEdit
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(False, .FromCopy, cboFromTable.ItemData(cboFromTable.ListIndex), .SelectedID) Then
              frmPick.Show vbModal
            End If
            If frmSelection.FromCopy And frmPick.SelectedID > 0 Then
              frmSelection.SelectedID = frmPick.SelectedID
            End If
            Set frmPick = Nothing

          'MH20050728 Fault 10232
          Case edtPrint
            Set frmPick = New frmPicklists
            frmPick.PrintDef .TableID, .SelectedID
            Unload frmPick
            Set frmPick = Nothing

          Case edtSelect
            Me.Changed = True
  
            txtPicklist.Text = frmSelection.SelectedText
            txtPicklist.Tag = frmSelection.SelectedID
            txtFilter.Text = ""
            txtFilter.Tag = 0
            fExit = True
      
          Case 0
            If IsPicklistValid(txtPicklist.Tag) <> vbNullString Then
              txtPicklist = "<None>"
              txtPicklist.Tag = 0
            End If
            fExit = True
        
        End Select
      End If
    Loop
  End With
  Set frmSelection = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub Form_Activate()
  Cancelled = False
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = 192
    KeyCode = 0
End Select
  
End Sub

Private Sub Form_Load()
  Cancelled = False
  grdAccess.RowHeight = 239
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancel = True
    Call cmdCancel_Click
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set datData = Nothing
End Sub

Private Sub grdAccess_ComboCloseUp()
  Changed = True
  
  'JPD 20030728 Fault 6486
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) And _
    (Len(grdAccess.Columns("Access").Text) > 0) Then
    ' The 'All Groups' access has changed. Apply the selection to all other groups.
    ForceAccess AccessCode(grdAccess.Columns("Access").Text)
    
    grdAccess.MoveFirst
    grdAccess.Col = 1
  End If

End Sub

Private Sub ForceAccess(Optional pvAccess As Variant)
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    .MoveFirst

    For iLoop = 0 To (.Rows - 1)
      varBookmark = .Bookmark
      
      If iLoop = 0 Then
        .Columns("Access").Text = ""
      Else
        If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
          If mblnForceHidden Then
            .Columns("Access").Text = AccessDescription(ACCESS_HIDDEN)
          Else
            If Not IsMissing(pvAccess) Then
              .Columns("Access").Text = AccessDescription(CStr(pvAccess))
            End If
          End If
        End If
      End If
      
      .MoveNext
    Next iLoop
    
    .MoveFirst
  End With
  
  UI.UnlockWindow

End Sub



Private Sub grdAccess_GotFocus()
  grdAccess.Col = 1

End Sub

Private Sub grdAccess_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  Dim varBkmk As Variant

  'JPD 20030728 Fault 6486
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) Then
    grdAccess.Columns("Access").Text = ""
  End If

  With grdAccess
    varBkmk = .SelBookmarks(0)

    If ((Not mblnDefinitionCreator) Or mblnReadOnly Or mblnForceHidden) Or _
      (.Columns("SysSecMgr").CellText(varBkmk) = "1") Then
      .Columns("Access").Style = ssStyleEdit
    Else
      .Columns("Access").Style = ssStyleComboBox
      .Columns("Access").RemoveAll
      .Columns("Access").AddItem AccessDescription(ACCESS_READWRITE)
      .Columns("Access").AddItem AccessDescription(ACCESS_READONLY)
      .Columns("Access").AddItem AccessDescription(ACCESS_HIDDEN)
    End If
  End With

  If Me.ActiveControl Is grdAccess Then
    grdAccess.Col = 1
  End If

End Sub

Private Sub grdAccess_RowLoaded(ByVal Bookmark As Variant)
  With grdAccess
    If (Not mblnDefinitionCreator) Or mblnReadOnly Or mblnForceHidden Then
      .Columns("GroupName").CellStyleSet "ReadOnly"
      .Columns("Access").CellStyleSet "ReadOnly"
      .ForeColor = vbGrayText
    ElseIf (.Columns("SysSecMgr").CellText(Bookmark) = "1") Then
      .Columns("GroupName").CellStyleSet "SysSecMgr"
      .Columns("Access").CellStyleSet "SysSecMgr"
      .ForeColor = vbWindowText
    Else
      .ForeColor = vbWindowText
    End If
  End With

End Sub


Private Sub grdColumns_Click()
'TM20010821 Fault 2387
'The grid must be enabled to allow the user to scroll through the columns.
'but if the definition is read only disable functionality(click) on grid.
  If Not mblnReadOnly Then
    With grdColumns
      .SelBookmarks.Add .Bookmark
    End With
  End If
End Sub

Private Sub grdColumns_DblClick()
'TM20010821 Fault 2387
'The grid must be enabled to allow the user to scroll through the columns.
'but if the definition is read only disable functionality(double click) on grid.
  If Not mblnReadOnly Then
    If cmdEdit.Enabled Then
      cmdEdit_Click
    Else
      cmdNew_Click
    End If
  End If
End Sub

Private Sub optAllRecords_Click()
  ' Reset the picklist and filter selections.
  
  cmdPicklist.Enabled = False
  txtPicklist.Tag = 0
  txtPicklist.Text = ""
  
  cmdFilter.Enabled = False
  txtFilter.Tag = 0
  txtFilter.Text = ""
  
  ForceDefinitionToBeHiddenIfNeeded

  Me.Changed = True

End Sub

Private Sub optFilter_Click()
  ' Enable the filter selection command button.
  
  cmdPicklist.Enabled = False
  txtPicklist.Tag = 0
  txtPicklist.Text = ""
  
  cmdFilter.Enabled = True
  txtFilter.Tag = 0
  txtFilter.Text = "<None>"
  
  ForceDefinitionToBeHiddenIfNeeded

  Me.Changed = True

End Sub

Private Sub optPicklist_Click()
  ' Enable the picklist selection command button.
  
  cmdPicklist.Enabled = True
  
  cmdFilter.Enabled = False
  txtFilter.Tag = 0
  txtFilter.Text = ""

  If Not mbLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  Me.Changed = True

End Sub

Private Sub RetreiveDefinition()
  ' Initialise the Column Transfer definition grid.
  Dim rsTemp As Recordset
  Dim sSQL As String
  Dim sRow As String
  Dim strRecSelStatus As String
  Dim sMessage As String
  Dim fAlreadyNotified As Boolean

  ' RH BUG 1007 Related ! You DO need to refresh the dest combo, so dont set the
  ' loading flag until after the combos
  
  'mbLoading = True
  
        '  sSQL = "SELECT ASRSysDataTransferName.*, " & _
        '         "       CONVERT(integer,ASRSysDataTransferName.TimeStamp) AS intTimeStamp, " & _
        '         "ASRSysPickListName.Name AS PickListName, " & _
        '         "ASRSysPickListName.Access AS PickListAccess, " & _
        '         "ASRSysExpressions.Name AS FilterName, " & _
        '         "ASRSysExpressions.Access AS FilterAccess " & _
        '         "FROM ASRSysDataTransferName " & _
        '         "LEFT OUTER JOIN ASRSysExpressions " & _
        '         "  ON ASRSysDataTransferName.FilterID = ASRSysExpressions.ExprID " & _
        '         "LEFT OUTER JOIN ASRSysPickListName " & _
        '         "  ON ASRSysDataTransferName.PickListID = ASRSysPickListName.PickListID " & _
        '         "WHERE ASRSysDataTransferName.DataTransferID = " & mlTransferID
        '  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    Screen.MousePointer = vbDefault
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Data Transfer"
    fOK = False
    Exit Sub
  End If

  ' RH 25/09/00 BUG 1007 - shouldnt just set the text property of a style2 combo - the
  '                        text might not exist.
  'cboFromTable = GetItemName(True, rstemp!fromTableID)
  'cboToTable = GetItemName(True, rstemp!toTableID)
  SetComboText cboFromTable, GetItemName(True, rsTemp!fromTableID)
  SetComboText cboToTable, GetItemName(True, rsTemp!toTableID)
  
  mbLoading = True
  
  ' NOTE : this line was commented before the above bug fix
  'SetComboItem cboOrder, rsTemp!OrderID

  txtFilter.Text = ""
  txtPicklist.Text = ""
  cmdFilter.Enabled = False
  cmdPicklist.Enabled = False
    
  
  
  mlTimeStamp = rsTemp!intTimestamp

  'MH20000216
  'Populate description and user name
  txtDesc.Text = IIf(IsNull(rsTemp!Description), vbNullString, rsTemp!Description)

  If mbFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If
  
  
  If rsTemp!FilterID > 0 Then
    optFilter.Value = True
    txtFilter.Tag = rsTemp!FilterID
    txtFilter.Text = rsTemp!FilterName
  ElseIf rsTemp!PicklistID > 0 Then
    optPicklist.Value = True
    txtPicklist.Tag = rsTemp!PicklistID
    txtPicklist = rsTemp!PicklistName
  Else
    optAllRecords.Value = True
  End If
  
  mblnReadOnly = Not datGeneral.SystemPermission("DATATRANSFER", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlDataTransfer, mlTransferID) = ACCESS_READONLY)
  End If
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  rsTemp.Close

  
  'sSQL = "exec sp_ASRGetDataTransferDetails " & mlTransferID
  'Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set rsTemp = GetDataTransfer
    
  
  With grdColumns
    .RemoveAll
    
    Do While Not rsTemp.EOF
      If rsTemp!fromTableID > 0 Then
        sRow = rsTemp!FromTableName & vbTab & _
          rsTemp!fromTableID & vbTab & _
          rsTemp!FromColumnName & vbTab & _
          rsTemp!fromColumnID & vbTab & _
          "TO" & vbTab
      Else
        If rsTemp!fromSysDate Then
          sRow = "<System Date>" & vbTab & _
            0 & vbTab & _
            vbTab & _
            0 & vbTab & _
            "TO" & vbTab
        Else
          sRow = Chr$(34) & rsTemp!fromText & Chr$(34) & vbTab & _
            0 & vbTab & _
            vbTab & _
            0 & vbTab & _
            "TO" & vbTab
        End If
      End If
            
      sRow = sRow & rsTemp!ToTableName & vbTab & _
        rsTemp!toTableID & vbTab & _
        rsTemp!ToColumnName & vbTab & _
        rsTemp!toColumnID
      
      .AddItem sRow
      rsTemp.MoveNext
    Loop
    
    .Enabled = True
    
  End With
    
  rsTemp.Close
  Set rsTemp = Nothing
    
  If grdColumns.Rows > 0 And Not mblnReadOnly Then
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdClearAll.Enabled = True
    
    With grdColumns
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End With
  End If
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Screen.MousePointer = vbDefault
    fOK = False
    Exit Sub
  End If
  
  mbLoading = False

End Sub

Private Function ForceDefinitionToBeHiddenIfNeeded(Optional pvOnlyFatalMessages As Variant) As Boolean
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim lngFilterID As Long
  Dim sRow As String
  Dim iResult As RecordSelectionValidityCodes
  Dim sBigMessage As String
  Dim asDeletedParameters() As String
  Dim asHiddenBySelfParameters() As String
  Dim asHiddenByOtherParameters() As String
  Dim asInvalidParameters() As String
  Dim fChangesRequired As Boolean
  Dim fDefnAlreadyHidden As Boolean
  Dim fNeedToForceHidden As Boolean
  Dim fRemove As Boolean
  Dim strColumnType As String
  Dim lngColumnID As Long
  Dim sCalcName As String
  Dim fOnlyFatalMessages As Boolean
  
  If IsMissing(pvOnlyFatalMessages) Then
    fOnlyFatalMessages = mbLoading
  Else
    fOnlyFatalMessages = CBool(pvOnlyFatalMessages)
  End If
  
  ' Return false if some of the filters/picklists/calcs need to be removed from the definition,
  ' or if the definition needs to be made hidden.
  fChangesRequired = False
  fDefnAlreadyHidden = AllHiddenAccess
  fNeedToForceHidden = False

  ' Dimension arrays to hold details of the filters/picklists that
  ' have been deleted, made hidden or are now invalid.
  ' Column 1 - parameter description
  ReDim asDeletedParameters(0)
  ReDim asHiddenBySelfParameters(0)
  ReDim asHiddenByOtherParameters(0)
  ReDim asInvalidParameters(0)

  ' Check Base Table Picklist
  If (Len(txtPicklist.Tag) > 0) And (Val(txtPicklist.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtPicklist.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The '" & cboFromTable.List(cboFromTable.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtPicklist.Tag = 0
      txtPicklist.Text = "<None>"
    End If
  End If

  ' Base Table Filter
  If Len(txtFilter.Tag) > 0 And Val(txtFilter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtFilter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & cboFromTable.List(cboFromTable.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboFromTable.List(cboFromTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtFilter.Tag = 0
      txtFilter.Text = "<None>"
    End If
  End If

  ' Construct one big message with all of the required error messages.
  sBigMessage = ""

  If UBound(asHiddenBySelfParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      'JPD 20040219 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the " & asHiddenBySelfParameters(1) & " is hidden."
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If (Not mblnForceHidden) And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the " & asHiddenBySelfParameters(1) & " is hidden."
        End If
      Else
        If (Not mbFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden as the " & asHiddenBySelfParameters(1) & " is hidden."
        End If
      End If
    Else
      sBigMessage = "The " & asHiddenBySelfParameters(1) & " will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
    End If
  ElseIf UBound(asHiddenBySelfParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      'JPD 20040308 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the following parameters are hidden :" & vbCrLf
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If (Not mblnForceHidden) And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the following parameters are hidden :" & vbCrLf
        End If
      Else
        If (Not mbFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden as the following parameters are hidden :" & vbCrLf
        End If
      End If
    Else
      sBigMessage = "The following parameters will be removed from this definition as they are hidden and you do not have permission to make this definition hidden :" & vbCrLf
    End If

    If Len(sBigMessage) > 0 Then
      For iLoop = 1 To UBound(asHiddenBySelfParameters)
        sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenBySelfParameters(iLoop)
      Next iLoop
    End If
  End If

  If UBound(asDeletedParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " has been deleted."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " will be removed from this definition as it has been deleted."
    End If
  ElseIf UBound(asDeletedParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been deleted :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been deleted :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asDeletedParameters)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asDeletedParameters(iLoop)
    Next iLoop
  End If

  If UBound(asHiddenByOtherParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " has been made hidden by another user."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " will be removed from this definition as it has been made hidden by another user."
    End If
  ElseIf UBound(asHiddenByOtherParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been made hidden by another user :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been made hidden by another user :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asHiddenByOtherParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenByOtherParameters(iLoop)
    Next iLoop
  End If

  If UBound(asInvalidParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " is invalid."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " will be removed from this definition as it is invalid."
    End If
  ElseIf UBound(asInvalidParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters are invalid :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they are invalid :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asInvalidParameters)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asInvalidParameters(iLoop)
    Next iLoop
  End If

  If Not FormPrint Then
    If mblnForceHidden And (Not fNeedToForceHidden) And (Not fOnlyFatalMessages) Then
      sBigMessage = "This definition no longer has to be hidden." & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        sBigMessage
    End If

    mblnForceHidden = fNeedToForceHidden
    ForceAccess
  End If

  If Len(sBigMessage) > 0 Then
    If FormPrint Then
      sBigMessage = Me.Caption & " print failed. The definition is currently invalid : " & vbCrLf & vbCrLf & sBigMessage
    End If

    COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
  End If

  ForceDefinitionToBeHiddenIfNeeded = (Len(sBigMessage) = 0)

End Function



Private Function AllHiddenAccess() As Boolean
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) <> AccessDescription(ACCESS_HIDDEN) Then
          AllHiddenAccess = False
          Exit Function
        End If
      End If
    Next iLoop
  End With

  AllHiddenAccess = True
  
End Function




Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mbCancelled = bCancel
End Property

Public Property Get FromCopy() As Boolean
  FromCopy = mbFromCopy
End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)
  mbFromCopy = bCopy
End Property

Public Sub AddNew()
  cmdNew_Click
End Sub

Public Sub DeleteRecord()
  cmdDelete_Click
End Sub

Private Function SaveDefinition() As Boolean

  Dim sSQL As String
  Dim lCount As Long
  'Dim lTransferID As Long
  Dim rsTransfer As New Recordset
    
  Dim strSourceTableID As String
  Dim strSourceColumnID As String
  Dim strFreeText As String
  Dim strSystemDate As String
  Dim strDestinTableID As String
  Dim strDestinColumnID As String
    
    
  On Error GoTo Err_Trap
  
  If Not ValidateDefinition Then
    SaveDefinition = False
    Exit Function
  End If
    
  If mlTransferID > 0 Then
    'Editing an existing transfer
    sSQL = "Delete From ASRSysDataTransferColumns Where DataTransferID = " & mlTransferID
    datData.ExecuteSql sSQL
    
    
    '{MH20000216
    ' 1. Include Description column
    ' 2. Ensure that name and description columns can accept ' signs.
    '
    'sSQL = "Update ASRSysDataTransferName Set Name = '" & txtName & "', FromTableID = " & _
    '    cboFromTable.ItemData(cboFromTable.ListIndex)
    sSQL = "Update ASRSysDataTransferName Set " & _
           "Name = '" & Replace(Trim(txtName), "'", "''") & "', " & _
           "Description = '" & Replace(txtDesc, "'", "''") & "', " & _
           "FromTableID = " & _
           cboFromTable.ItemData(cboFromTable.ListIndex)
    'MH20000216}


    If optAllRecords Then
      sSQL = sSQL & ", AllRecords = 1, PickListID = 0, FilterID = 0"
    ElseIf optPicklist Then
      sSQL = sSQL & ", AllRecords = 0, PickListID = " & Val(txtPicklist.Tag) & ", FilterID = 0"
    Else
      sSQL = sSQL & ", AllRecords = 0, PickListID = 0, FilterID = " & Val(txtFilter.Tag)
    End If
    
    sSQL = sSQL & ", ToTableID = " & cboToTable.ItemData(cboToTable.ListIndex)
    'sSQL = sSQL & ", OrderID = " & cboOrder.ItemData(cboOrder.ListIndex)
    
    '{MH20000216
    'Don't update user name as this should remain constant...
    sSQL = sSQL & " Where DataTransferID = " & mlTransferID
            
    datData.ExecuteSql sSQL
    
    sSQL = ""
    With grdColumns
      .Redraw = False
      For lCount = 0 To .Rows - 1
        .Bookmark = .AddItemBookmark(lCount)
        
        strFreeText = "''"
        strSystemDate = "0"
        
        If .Columns(1).Value = 0 Then
          strSourceTableID = "0"
          strSourceColumnID = "0"
          If .Columns(0).Text = "<System Date>" Then
            strSystemDate = "1"
          Else
            strFreeText = "'" & Mid$(.Columns(0).Text, 2, Len(.Columns(0).Text) - 2) & "'"
          End If
        Else
          strSourceTableID = .Columns(1).Value
          strSourceColumnID = .Columns(3).Value
        End If
        
        sSQL = "INSERT INTO ASRSysDataTransferColumns (" & _
                  "dataTransferID, " & _
                  "fromTableID, " & _
                  "fromColumnID, " & _
                  "fromText, " & _
                  "fromSysDate, " & _
                  "toTableID, " & _
                  "toColumnID)" & _
               " VALUES(" & _
                  mlTransferID & ", " & _
                  strSourceTableID & ", " & _
                  strSourceColumnID & ", " & _
                  strFreeText & ", " & _
                  strSystemDate & ", " & _
                  .Columns(6).Value & ", " & _
                  .Columns(8).Value & ")"
        
        datData.ExecuteSql sSQL
      Next
      .Redraw = True
    End With
    
    Call UtilUpdateLastSaved(utlDataTransfer, mlTransferID)
  
  Else
    '{MH20000216
    ' 1. Include Description column
    ' 2. Ensure that name and description columns can accept ' signs.
    '
    'sSQL = "Insert ASRSysDataTransferName Values('" & txtName & "', " & cboFromTable.ItemData(cboFromTable.ListIndex)
    sSQL = "INSERT INTO ASRSysDataTransferName" & _
      " (name, description, fromTableID, allRecords, filterID, picklistID, toTableID, userName)" & _
      " VALUES(" & _
      "'" & Replace(Trim(txtName.Text), "'", "''") & "', " & _
      "'" & Replace(txtDesc.Text, "'", "''") & "', " & _
      cboFromTable.ItemData(cboFromTable.ListIndex)
      
    If optAllRecords Then
      sSQL = sSQL & ", 1, 0, 0"
    ElseIf optPicklist Then
      sSQL = sSQL & ", 0, 0, " & Val(txtPicklist.Tag)
    Else
      sSQL = sSQL & ", 0, " & Val(txtFilter.Tag) & ", 0"
    End If
    sSQL = sSQL & _
           ", " & CStr(cboToTable.ItemData(cboToTable.ListIndex)) & _
           ", '" & datGeneral.UserNameForSQL & "')"
    
           '", " & CStr(cboOrder.ItemData(cboOrder.ListIndex)) & _

    ' RH 04/09/00 - Use insert util def stored procedure
    mlTransferID = InsertDataTransfer(sSQL)

    sSQL = ""
    With grdColumns
      For lCount = 0 To .Rows - 1
        .Bookmark = .AddItemBookmark(lCount)
        If .Columns(1).Value = 0 Then
          sSQL = "INSERT INTO ASRSysDataTransferColumns" & _
            " (dataTransferID, fromTableID, fromColumnID, fromText, fromSysDate, toTableID, toColumnID)" & _
            " VALUES(" & mlTransferID & ", 0, 0"
          If .Columns(0).Text = "<System Date>" Then
            sSQL = sSQL & ", '', 1"
          Else
            sSQL = sSQL & ", '" & Replace(.Columns(0).Text, Chr(34), vbNullString) & "', 0"
          End If
          sSQL = sSQL & ", " & .Columns(6).Value & ", " & .Columns(8).Value & ")"
        Else
          sSQL = "INSERT INTO ASRSysDataTransferColumns" & _
            " (dataTransferID, fromTableID, fromColumnID, fromText, fromSysDate, toTableID, toColumnID)" & _
            " VALUES(" & mlTransferID & ", "
          sSQL = sSQL & .Columns(1).Value & ", " & .Columns(3).Value & ", "
          sSQL = sSQL & "'', 0, " & .Columns(6).Value & ", " & .Columns(8).Value & ")"
      End If
        datData.ExecuteSql sSQL
      Next
    End With
  
    Call UtilCreated(utlDataTransfer, mlTransferID)
  
  End If
  
  SaveAccess
  
  SaveDefinition = True
  
  Exit Function
  
Err_Trap:
  COAMsgBox Err.Description
  SaveDefinition = False

End Function


Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysDataTransferAccess WHERE ID = " & mlTransferID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysDataTransferAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlTransferID & ", sysusers.name," & _
    " CASE" & _
    "   WHEN (SELECT count(*)" & _
    "     FROM ASRSysGroupPermissions" & _
    "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    "     WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
    "       AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
    "   ELSE '" & ACCESS_HIDDEN & "'" & _
    " END" & _
    " FROM sysusers" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    " AND sysusers.name <> 'ASRSysGroup'" & _
    " AND sysusers.uid <> 0)"
  datData.ExecuteSql (sSQL)

  ' Update the new access records with the real access values.
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      sSQL = "IF EXISTS (SELECT * FROM ASRSysDataTransferAccess" & _
        " WHERE ID = " & CStr(mlTransferID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysDataTransferAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlTransferID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub





Private Function InsertDataTransfer(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertTransfer_ERROR

  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean
  
  fSavedOK = True
  
  Set cmADO = New ADODB.Command
  
  With cmADO
    .CommandText = "sp_ASRInsertNewUtility"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
  
    Set .ActiveConnection = gADOCon
              
    Set pmADO = .CreateParameter("newID", adInteger, adParamOutput)
    .Parameters.Append pmADO
            
    Set pmADO = .CreateParameter("insertString", adLongVarChar, adParamInput, -1)
    .Parameters.Append pmADO
    pmADO.Value = pstrSQL
              
    Set pmADO = .CreateParameter("tablename", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = "AsrSysDataTransferName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "DataTransferID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertDataTransfer = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertDataTransfer = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertTransfer_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function



Public Sub EditRecord()

    cmdEdit_Click
    
End Sub

Private Function ValidateDefinition() As Boolean

  Dim lCount As Long
  Dim lFromTableID As Long
  Dim lToTableID As Long
  Dim bError As Boolean
  Dim sMsg As String
  Dim lFromParent As Long
  Dim lToParent As Long
  Dim blnFound As Boolean
  Dim varBookmark As Variant

  Dim strRecSelStatus As String
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean
  Dim strName As String
  
  Dim iCount_Owner As Integer
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobScheduledUserGroups As String
  Dim sHiddenGroups As String
  
  ValidateDefinition = False
  fBatchJobsOK = True
    
  strName = Trim(txtName.Text)
  
  'Dim bPtoP As Boolean        'Parent to parent defined
  'Dim bCtoC As Boolean        'Child to child defined of different parent
  
  
  If Len(strName) = 0 Then
    SSTab1.Tab = 0
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
    
  If optFilter Then
    If Val(txtFilter.Tag) = 0 Then
      SSTab1.Tab = 0
      COAMsgBox "No filter selected.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
  ElseIf optPicklist Then
    If Val(txtPicklist.Tag) = 0 Then
      SSTab1.Tab = 0
      COAMsgBox "No picklist selected.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
  End If

  'Check that at least one column on the destination
  'table has been specified
  blnFound = False
  With grdColumns
    .SelBookmarks.RemoveAll
    For lCount = 0 To .Rows - 1
      varBookmark = .AddItemBookmark(lCount)
      If Val(.Columns(6).CellText(varBookmark)) = cboToTable.ItemData(cboToTable.ListIndex) Then
        blnFound = True
        Exit For
      End If
    Next
  End With

  If blnFound = False Then
    COAMsgBox "No transfer columns specified on the main destination table", vbExclamation, Me.Caption
    ValidateDefinition = False
    Exit Function
  End If
  
  
  'Ensure that user has included all required columns
  If CheckMandatoryColumns = False Then
    Exit Function
  End If
  
  
  'Check if this definition has been changed by another user
  Call UtilityAmended(utlDataTransfer, mlTransferID, mlTimeStamp, blnContinueSave, blnSaveAsNew)
  If blnContinueSave = False Then
    Exit Function
  ElseIf blnSaveAsNew Then
    txtUserName = gsUserName
    
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlTransferID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  
  If ValidateUniqueName(strName, mlTransferID) = False Then
    SSTab1.Tab = 0
    COAMsgBox "A Data Transfer definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
  
If mlTransferID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlDataTransfer, _
      CStr(mlTransferID), _
      txtUserName.Text, _
      iCount_Owner, _
      sBatchJobDetails_Owner, _
      sBatchJobIDs, _
      sBatchJobDetails_NotOwner, _
      fBatchJobsOK, _
      sBatchJobDetails_ScheduledForOtherUsers, _
      sBatchJobScheduledUserGroups, _
      sHiddenGroups

    If (Not fBatchJobsOK) Then
      If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
        COAMsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, "Data Transfer"
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Data Transfer"
      End If

      Screen.MousePointer = vbDefault
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Data Transfer") = vbNo Then
        Screen.MousePointer = vbDefault
        SSTab1.Tab = 0
        Exit Function
      Else
        ' Ok, we are continuing, so lets update all those utils to hidden !
        If Len(Trim(sBatchJobIDs)) > 0 Then
          HideUtilities utlBatchJob, sBatchJobIDs, sHiddenGroups
          Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
        End If
      End If
    End If
  End If
End If
  
  ValidateDefinition = True

End Function

Private Function HiddenGroups() As String
  'Return a TAB delimited string of the user groups to which this definition is hidden.
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim sHiddenGroups As String
  
  sHiddenGroups = ""
  
  With grdAccess
    .Update
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) = AccessDescription(ACCESS_HIDDEN) Then
          sHiddenGroups = sHiddenGroups & .Columns("GroupName").CellText(varBookmark) & vbTab
        End If
      End If
    Next iLoop
  End With

  If Len(sHiddenGroups) > 0 Then
    sHiddenGroups = vbTab & sHiddenGroups
  End If
  
  HiddenGroups = sHiddenGroups
  
End Function





Private Sub SSTab1_Click(PreviousTab As Integer)
  fraDefinition(0).Enabled = (SSTab1.Tab = giTABSTRIP_DATATRANSFERDEF)
  fraDefinition(1).Enabled = fraDefinition(0).Enabled
  
  fraColumnDefinition.Enabled = (SSTab1.Tab = giTABSTRIP_COLUMNDEF)

End Sub

Private Sub txtDesc_Change()
  Me.Changed = True
End Sub

Private Sub txtDesc_GotFocus()
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  cmdOK.Default = False
End Sub

Private Sub txtDesc_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub txtName_Change()
  Me.Changed = True
End Sub

Private Sub txtName_GotFocus()
  UI.txtSelText
End Sub

Public Sub PickListSelected(lPicklistID As Long, sName As String)
  txtPicklist.Text = sName
  txtPicklist.Tag = lPicklistID
End Sub

Public Sub FilterSelected(lFilterID As Long, sName As String)
  txtFilter.Text = sName
  txtFilter.Tag = lFilterID
End Sub

Private Sub EnableAll()
  
  ' Enable all screen controls.
  Dim ctlTemp As Control
    
  For Each ctlTemp In Me.Controls
    ctlTemp.Enabled = True
  Next

  ' Disable some controls as default.
  txtPicklist.Enabled = False
  txtFilter.Enabled = False
  txtUserName.Enabled = False
  
  fraDefinition(0).Enabled = (SSTab1.Tab = giTABSTRIP_DATATRANSFERDEF)
  fraDefinition(1).Enabled = fraDefinition(0).Enabled
  fraColumnDefinition.Enabled = (SSTab1.Tab = giTABSTRIP_COLUMNDEF)
  
End Sub


Private Function ValidateUniqueName(sName As String, lTransferID As Long) As Boolean

  Dim rsName As Recordset
  Dim sSQL As String

  sSQL = "Select * From ASRSysDataTransferName Where Name = '" & Replace(sName, "'", "''") & "' AND DataTransferID <> " & lTransferID

  Set rsName = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ValidateUniqueName = (rsName.EOF And rsName.BOF)

  rsName.Close
  Set rsName = Nothing

End Function

Private Function GetItemName(bTable As Boolean, lItemID As Long) As String

    Dim sSQL As String
    Dim rsItem As Recordset
    
    If bTable Then
        sSQL = "Select TableName From ASRSysTables Where TableID = " & lItemID
    Else
        sSQL = "Select ColumnName From ASRSysColumns Where ColumnID = " & lItemID
     End If
    
                                          Set rsItem = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
        
    If rsItem.BOF And rsItem.EOF Then
      GetItemName = vbNullString
    Else
      GetItemName = rsItem(0)
    End If
    
    rsItem.Close
    Set rsItem = Nothing

End Function

Private Function GetCommonParent(lTableID1 As Long, lTableID2 As Long) As Long

  Dim sSQL As String
  Dim rsTemp As Recordset

  'sSQL = "Select ParentID From ASRSysRelations Where ChildID = " & lTableID
  sSQL = "SELECT ParentID FROM ASRSysRelations " & _
         "WHERE ParentID IN " & _
         "(SELECT ParentID FROM ASRSysRelations " & _
         " WHERE ChildID = " & CStr(lTableID1) & _
         ") AND ChildID = " & CStr(lTableID2)
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  If rsTemp.BOF And rsTemp.EOF Then
    GetCommonParent = 0
  Else
    GetCommonParent = rsTemp(0)
  End If
  rsTemp.Close
  Set rsTemp = Nothing

End Function


Private Function CheckMandatoryColumns() As Boolean
        
  Dim rsColumns As Recordset
  Dim strSQL As String

  Dim lngTableIDs() As Long
  Dim intIndex As Integer
  Dim lngRow As Long
  Dim pvarbookmark As Variant
  Dim lngTableID As Integer
  Dim blnFoundTable As Boolean

  Dim strMandatoryColumns As String
  Dim strTableIDs As String
  Dim strColumnIDs As String

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strTitle As String
  
  strMandatoryColumns = vbNullString
  
  ReDim lngTableIDs(0) As Long
  strTableIDs = vbNullString
  strColumnIDs = vbNullString
  
  With grdColumns
    'AE20071025 Fault #7972
    .Redraw = False
    .Row = 0
    For lngRow = 0 To .Rows - 1

      pvarbookmark = .AddItemBookmark(lngRow)
      .Bookmark = pvarbookmark
      'lngTableID = Val(.Columns(6).CellText(pvarbookmark))
      lngTableID = Val(.Columns(6).Text)
      If lngTableID > 0 Then
      
        strColumnIDs = strColumnIDs & _
          IIf(strColumnIDs <> "", ", ", "") & .Columns(8).Text

        'Loop though all of the columns in the array and
        'check to see if this table is already in the array
        blnFoundTable = False
        For intIndex = 0 To UBound(lngTableIDs)
          blnFoundTable = (lngTableIDs(intIndex) = lngTableID)
          If blnFoundTable Then
            Exit For
          End If
        Next

        'If this table is not in the array then
        'add this table to the array
        If blnFoundTable = False Then
          intIndex = IIf(lngTableIDs(0) = 0, 0, UBound(lngTableIDs) + 1)
          ReDim Preserve lngTableIDs(intIndex) As Long
          lngTableIDs(intIndex) = lngTableID
        
          strTableIDs = strTableIDs & _
            IIf(strTableIDs <> "", ", ", "") & CStr(lngTableID)
        End If
      
      End If

    Next
    'AE20071025 Fault #7972
    .Redraw = True
  End With

  
  'MH20000814
  'Allow save if mandatory ommitted if it has a default value
  'This is to get around the staff number on a applicants to personnel transfer
  
  'MH20000904
  'Allow save if mandatory ommitted and it is a calculated column

  '******************************************************************************
  ' TM20010719 Fault 2242 - ColumnType <> 4 clause added to ignore all linked   *
  ' columns. (It doesn't need to validate the linked columns because this is    *
  ' done using the Vaidate SP.                                                  *
  '******************************************************************************

  If strTableIDs <> vbNullString Then
    strSQL = "SELECT ASRSysTables.TableName, ASRSysColumns.ColumnName " & _
             "FROM ASRSysColumns " & _
             "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
             "WHERE ASRSysColumns.TableID IN (" & strTableIDs & ") " & _
             "  AND ASRSysColumns.ColumnID NOT IN (" & strColumnIDs & ") " & _
             "  AND " & SQLWhereMandatoryColumn & _
             " ORDER BY ASRSysTables.TableName, ASRSysColumns.ColumnName"
             '"  AND Mandatory = '1' " & _
             "  AND Rtrim(DefaultValue) = '' AND Convert(int,dfltValueExprID) = 0 " & _
             "  AND CalcExprID = 0 " & _
             "  AND ColumnType <> 4 "
    Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

    While Not rsColumns.EOF
      strMandatoryColumns = strMandatoryColumns & _
        rsColumns!TableName & "." & rsColumns!ColumnName & vbCrLf
      rsColumns.MoveNext
    Wend

  End If

  CheckMandatoryColumns = (strMandatoryColumns = vbNullString)
  
  If CheckMandatoryColumns = False Then
    strMBText = "Unable to save definition as the following mandatory" & vbCrLf & _
                "columns have not been populated:" & vbCrLf & vbCrLf & _
                strMandatoryColumns & vbCrLf & _
                "Please enter a source to populate these columns."
    intMBButtons = vbExclamation + vbOKOnly
    strTitle = App.ProductName
    COAMsgBox strMBText, intMBButtons, strTitle
  End If

End Function


Private Sub PopulateDestination()

  Dim strSQL As String
  Dim rsTemp As Recordset

  Dim strChild As Long

  strChild = CStr(cboFromTable.ItemData(cboFromTable.ListIndex))

  'Parents of this child
  strSQL = "SELECT ParentID FROM ASRSysRelations WHERE ChildID = " & strChild

  'Children of (parents of this child)
  strSQL = "SELECT ChildID FROM ASRSysRelations " & _
           "WHERE ParentID IN (" & strSQL & ")"

  'Get toplevel tables and (children of (parents of this child))
  strSQL = "SELECT TableID, TableName, DefaultOrderID " & _
           "FROM ASRSysTables " & _
           "WHERE TableID <> " & strChild & _
           " AND (TableType = " & CStr(tabTopLevel) & " OR " & _
           "TableID IN (" & strSQL & "))"

  LoadTableCombo cboToTable, strSQL
'  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
'
'
'  With cboToTable
'    .Clear
'    Do While Not rsTemp.EOF
'      .AddItem rsTemp!TableName
'      .ItemData(.NewIndex) = rsTemp!TableID
'      rsTemp.MoveNext
'    Loop
'    If .ListCount > 0 Then
'      .ListIndex = 0
'    End If
'  End With
'
'
'  rsTemp.Close
'  Set rsTemp = Nothing

End Sub


Private Function GetDefinition() As Recordset

  Dim sSQL As String
  
  sSQL = "SELECT ASRSysDataTransferName.*, " & _
         "       CONVERT(integer,ASRSysDataTransferName.TimeStamp) AS intTimeStamp, " & _
         "ASRSysPickListName.Name AS PickListName, " & _
         "ASRSysPickListName.Access AS PickListAccess, " & _
         "ASRSysExpressions.Name AS FilterName, " & _
         "ASRSysExpressions.Access AS FilterAccess " & _
         "FROM ASRSysDataTransferName " & _
         "LEFT OUTER JOIN ASRSysExpressions " & _
         "  ON ASRSysDataTransferName.FilterID = ASRSysExpressions.ExprID " & _
         "LEFT OUTER JOIN ASRSysPickListName " & _
         "  ON ASRSysDataTransferName.PickListID = ASRSysPickListName.PickListID " & _
         "WHERE ASRSysDataTransferName.DataTransferID = " & mlTransferID
  
  Set GetDefinition = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

End Function


Public Sub PrintDef(lTransferID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsColumns As Recordset
  Dim sSQL As String
  Dim strSource As String
  Dim iLoop As Integer
  Dim varBookmark As Variant

  Set datData = New HRProDataMgr.clsDataAccess
  
  mlTransferID = lTransferID
  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Data Transfer"
    Exit Sub
  End If
  
  
  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Data Transfer : " & rsTemp!Name
        
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal "Owner : " & rsTemp!UserName
        
        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop
          
        ' Data section --------------------------------------------------------
        .PrintNormal
        
        .PrintNormal "Source Table : " & GetItemName(True, rsTemp!fromTableID)
        'MH20000905 Check if picklist/filter has been deleted
        If rsTemp!AllRecords Then
          .PrintNormal "Records : All"
        ElseIf rsTemp!FilterID > 0 Then
          .PrintNormal "Records : Filter '" & IIf(IsNull(rsTemp!FilterName), "<Deleted>", rsTemp!FilterName) & "'"
        Else
          .PrintNormal "Records : Picklist '" & IIf(IsNull(rsTemp!PicklistName), "<Deleted>", rsTemp!PicklistName) & "'"
        End If
        .PrintNormal
        
        .PrintNormal "Destination Table : " & GetItemName(True, rsTemp!toTableID)
        
        '---------
        
        .PrintTitle "Columns"
        
        Set rsColumns = GetDataTransfer
        
        .PrintBold "Source Column / Data" & vbTab & _
                   "Destination Column"
    
        Do While Not rsColumns.EOF
          
          If rsColumns!fromTableID > 0 Then
            strSource = rsColumns!FromTableName & "." & rsColumns!FromColumnName
          ElseIf rsColumns!fromSysDate Then
            strSource = "<System Date>"
          Else
            strSource = Chr$(34) & rsColumns!fromText & Chr$(34)
          End If
    
          '.PrintNonBold rsColumns!ToTableName & "." & rsColumns!ToColumnName & vbTab & _
                        strSource
          .PrintNonBold strSource & vbTab & _
                        rsColumns!ToTableName & "." & rsColumns!ToColumnName
                        
    
          rsColumns.MoveNext
        Loop
        
        'Finish off printing
        .PrintEnd
        .PrintConfirm "Data Transfer : " & rsTemp!Name, "Data Transfer Definition"
        
      End If
    End With
  
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing
  Set datData = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Data Transfer Definition Failed"

End Sub


Private Function GetDataTransfer() As Recordset

  Dim sSQL As String

  sSQL = "SELECT ASRSysTables.tableName AS fromTablename, " & _
         "       ASRSysColumns.columnName AS fromColumnName, " & _
         "       ASRSysTables1.tableName AS toTableName, " & _
         "       ASRSysColumns1.columnName AS toColumnName, " & _
         "       ASRSysDataTransferColumns.fromText, " & _
         "       ASRSysDataTransferColumns.fromSysDate, " & _
         "       ASRSysDataTransferColumns.dataTransferID, " & _
         "       ASRSysDataTransferColumns.fromTableID, " & _
         "       ASRSysDataTransferColumns.fromColumnID, " & _
         "       ASRSysDataTransferColumns.toTableID, " & _
         "       ASRSysDataTransferColumns.toColumnID " & _
         "FROM ASRSysDataTransferColumns " & _
         "LEFT OUTER JOIN ASRSysColumns ASRSysColumns1 " & _
         "ON ASRSysDataTransferColumns.toColumnID = ASRSysColumns1.columnID " & _
         "LEFT OUTER JOIN ASRSysTables ASRSysTables1 " & _
         "ON ASRSysDataTransferColumns.toTableID = ASRSysTables1.tableID " & _
         "LEFT OUTER JOIN ASRSysColumns " & _
         "ON ASRSysDataTransferColumns.fromColumnID = ASRSysColumns.columnID " & _
         "LEFT OUTER JOIN ASRSysTables " & _
         "ON ASRSysDataTransferColumns.fromTableID = ASRSysTables.tableID " & _
         "WHERE ASRSysDataTransferColumns.dataTransferID = " & CStr(mlTransferID) & _
         " ORDER BY toTableName, toColumnName"
  
  Set GetDataTransfer = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

End Function


Private Sub CheckIfScrollBarRequired()

  With grdColumns

    If .Rows > 14 Then
      .ScrollBars = ssScrollBarsVertical
      .Columns("Destination Column").Width = 1710
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns("Destination Column").Width = 1950
    End If

  End With

End Sub

