VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGlobalFunctions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global "
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGlobalFunctions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4560
      Left            =   90
      TabIndex        =   26
      Top             =   90
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   8043
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Definition"
      TabPicture(0)   =   "frmGlobalFunctions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinition(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDefinition(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmGlobalFunctions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumns"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDefinition 
         Height          =   2355
         Index           =   0
         Left            =   135
         TabIndex        =   27
         Top             =   360
         Width           =   9405
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5805
            MaxLength       =   30
            TabIndex        =   4
            Top             =   300
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   1
            Top             =   300
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1530
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1110
            Width           =   3090
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   3090
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1485
            Left            =   5805
            TabIndex        =   5
            Top             =   705
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
            stylesets(0).Picture=   "frmGlobalFunctions.frx":0044
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
            stylesets(1).Picture=   "frmGlobalFunctions.frx":0060
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
            _ExtentY        =   2619
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
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   4950
            TabIndex        =   32
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   31
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   195
            TabIndex        =   30
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   4950
            TabIndex        =   29
            Top             =   765
            Width           =   825
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category :"
            Height          =   240
            Left            =   195
            TabIndex        =   28
            Top             =   765
            Width           =   1005
         End
      End
      Begin VB.Frame fraColumns 
         Enabled         =   0   'False
         Height          =   4050
         Left            =   -74865
         TabIndex        =   18
         Top             =   360
         Width           =   9400
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   23
            Top             =   1935
            Width           =   1200
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7950
            TabIndex        =   20
            Top             =   315
            Width           =   1200
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   21
            Top             =   855
            Width           =   1200
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   22
            Top             =   1395
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   3570
            Left            =   210
            TabIndex        =   19
            Top             =   315
            Width           =   7440
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   7
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
            Columns.Count   =   7
            Columns(0).Width=   6535
            Columns(0).Caption=   "Column"
            Columns(0).Name =   "Column"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   6138
            Columns(1).Caption=   "Value"
            Columns(1).Name =   "Value"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "ColumnID"
            Columns(2).Name =   "ColumnID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "ValueTypeID"
            Columns(3).Name =   "ValueTypeID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "ValueID"
            Columns(4).Name =   "ValueID"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "LookupTableID"
            Columns(5).Name =   "LookupTableID"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "LookupColumnID"
            Columns(6).Name =   "LookupColumnID"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13123
            _ExtentY        =   6297
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
         Height          =   1650
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   2760
         Width           =   9400
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6930
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1080
            Width           =   1965
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6930
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   705
            Width           =   1965
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5865
            TabIndex        =   15
            Top             =   1120
            Width           =   975
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5865
            TabIndex        =   12
            Top             =   750
            Width           =   930
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5865
            TabIndex        =   11
            Top             =   365
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8900
            TabIndex        =   14
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8900
            TabIndex        =   17
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboTables 
            Height          =   315
            ItemData        =   "frmGlobalFunctions.frx":007C
            Left            =   1620
            List            =   "frmGlobalFunctions.frx":007E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   315
            Width           =   3000
         End
         Begin VB.ComboBox cboChild 
            Height          =   315
            ItemData        =   "frmGlobalFunctions.frx":0080
            Left            =   1620
            List            =   "frmGlobalFunctions.frx":0082
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   705
            Width           =   3000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Index           =   5
            Left            =   4920
            TabIndex        =   10
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parent Table :"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   6
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Child Table :"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   8
            Top             =   750
            Width           =   1155
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7350
      TabIndex        =   24
      Top             =   4755
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8595
      TabIndex        =   25
      Top             =   4755
      Width           =   1200
   End
End
Attribute VB_Name = "frmGlobalFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private datGlobal As DataMgr.clsGlobal
Private datData As DataMgr.clsDataAccess
'Private mlngTimeStamp As Long
Private mblnReadOnly As Boolean

Private fOK As Boolean
Private mbLoading As Boolean
Private mlTableID As Long
Private mlPrevTableID As Long
Private msComboText As String
Private mbFromCopy As Boolean
Private mblnFormPrint As Boolean

'Private mbCancelled As Boolean
Private mlFunctionID As Long
Private typGlobal As GlobalType
'Private mDefSel As DefSelScreen
'Private me.Changed As Boolean
Private mlngTimeStamp As Long
Private mblnDefinitionCreator As Boolean

Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean
Private mblnDeletedCalc As Boolean

Private mbNeedsSave As Boolean
Private mblnCancelled As Boolean

Const giTABSTRIP_GLOBALFUNCTIONDEF = 0
Const giTABSTRIP_COLUMNDEF = 1


Private Sub PopulateAccessGrid()
  ' Populate the access grid.
  Dim rsAccess As ADODB.Recordset
  
  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With
  
  ' Get the recordset of user groups and their access on this definition.
  Select Case typGlobal
    Case glAdd
      Set rsAccess = GetUtilityAccessRecords(utlGlobalAdd, mlFunctionID, mbFromCopy)
    Case glUpdate
      Set rsAccess = GetUtilityAccessRecords(utlGlobalUpdate, mlFunctionID, mbFromCopy)
    Case glDelete
      Set rsAccess = GetUtilityAccessRecords(utlGlobalDelete, mlFunctionID, mbFromCopy)
  End Select
  
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

Public Property Get SelectedID() As Long
  SelectedID = mlFunctionID
End Property

Public Property Get DefinitionCreator() As Boolean
  DefinitionCreator = mblnDefinitionCreator
End Property



Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property


'# NB, This is the edited version replacing the listbox with the ssdbgrid.
'#     All original code can be found at the end of this module, commented out
'#     The listview ctl still exists on the form but is invisible.  RH 27/07/99

Public Function Initialise(bNew As Boolean, bFromCopy As Boolean, typeGlobal As GlobalType, _
  Optional lFunctionID As Long, Optional bPrint As Boolean) As Boolean
  ' Initialise the Global Function definition form.

  Screen.MousePointer = vbHourglass
  'Set datGlobal = New DataMgr.clsGlobal

  mlFunctionID = 0
  mbFromCopy = bFromCopy
  typGlobal = typeGlobal
  'JPD 20030911 Fault 6359
  Me.Caption = GetCaption & " Definition"

  mbLoading = True
  fOK = True
  
  FormatForm

  mbLoading = False

  If Not bNew Then
    FormPrint = bPrint
    mlFunctionID = lFunctionID
    
    PopulateAccessGrid

    fOK = RetreiveDefinition

    If fOK Then
    
      Me.Changed = False
  
      If mbFromCopy Then
        mlFunctionID = 0
        Me.Changed = True
      Else
        Me.Changed = ((mblnRecordSelectionInvalid Or mblnDeletedCalc) And Not mblnReadOnly)
      End If
    End If
  Else
    txtUserName = gsUserName
    mblnDefinitionCreator = True
    Call cboTables_Click
    
    GetObjectCategories cboCategory, utlGlobalAdd, 0, cboTables.ItemData(cboTables.ListIndex)
    SetComboItem cboCategory, IIf(glngCurrentCategoryID = -1, 0, glngCurrentCategoryID)
    
    PopulateAccessGrid
    
    Me.Changed = False
  End If

  CheckIfScrollBarRequired
  SSTab1.Tab = giTABSTRIP_GLOBALFUNCTIONDEF
  Screen.MousePointer = vbDefault

  Initialise = fOK

End Function


Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property


Private Sub GetTables(bParent As Boolean)
  ' Populate the Tables combo.
'  Dim rsTables As Recordset
  Dim sSQL As String

  sSQL = vbNullString
  If bParent Then
    'MH20011105 Fault 3097
    'sSQL = "Select TableID, TableName From ASRSysTables Where TableType = 1" & _
           " AND (SELECT COUNT(ChildID) FROM ASRSysRelations WHERE ParentID = ASRSysTables.TableID) > 0"
    sSQL = "SELECT TableID, TableName FROM ASRSysTables " & _
           "WHERE TableID IN (SELECT DISTINCT ParentID FROM ASRSysRelations)"
  Else
    'MH 20010905 Allow Updates and Deletes on all tables (including Lookups)
    sSQL = "Select TableID, TableName From ASRSysTables"
'    Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  Else
'    Set rsTables = datGeneral.GetAllTables
  End If

  LoadTableCombo cboTables, sSQL

'  With cboTables
'    .Clear
'
'    Do While Not rsTables.EOF
'      .AddItem rsTables!TableName
'      .ItemData(.NewIndex) = rsTables!TableID
'      rsTables.MoveNext
'    Loop
'
'    If .ListCount > 0 Then
'      .ListIndex = 0
'    End If
'  End With
'
'  rsTables.Close
'  Set rsTables = Nothing

End Sub

Private Sub GetChildTables(lParentTableID As Long)

  Dim rsTables As Recordset
  Dim sSQL As String
  Dim bOriginalLoading As Boolean

  sSQL = "SELECT ASRSysTables.TableName, ASRSysTables.TableID FROM ASRSysTables INNER JOIN " & _
         "ASRSysRelations ON ASRSysTables.TableID = ASRSysRelations.ChildID Where " & _
         "ASRSysRelations.ParentID = " & lParentTableID
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  bOriginalLoading = mbLoading
  mbLoading = True
  With cboChild
    .Clear

    Do While Not rsTables.EOF
      .AddItem rsTables!TableName
      .ItemData(.NewIndex) = rsTables!TableID
      rsTables.MoveNext
    Loop

    If .ListCount > 0 Then
      .ListIndex = 0
      .Enabled = True
      cmdNew.Enabled = True
      cmdEdit.Enabled = (grdColumns.Rows > 0)
      cmdDelete.Enabled = (grdColumns.Rows > 0)
      cmdClearAll.Enabled = (grdColumns.Rows > 0)
      grdColumns.Enabled = True
    Else
      .Enabled = False
      grdColumns.Enabled = False
      cmdNew.Enabled = True
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      cmdClearAll.Enabled = False
    End If
  End With
  mbLoading = bOriginalLoading

End Sub

Private Sub cboChild_Click()

  If Not mbLoading Then
    If Me.cboChild.Text = msComboText Then Exit Sub

    If grdColumns.Rows > 0 Then
      If COAMsgBox("One or more columns from the '" & msComboText & "' table " & _
                "have been included in the current " & LCase(Me.Caption) & _
                " definition. Changing the child table will remove these " & _
                "columns from the " & LCase(Me.Caption) & " definition." & vbCrLf & _
                "Do you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
        'TM20010910 Fault 2793
        'Commented out the following.
'        mbLoading = True
        cboChild.Text = msComboText
'        mbLoading = False
        Exit Sub
      End If
    End If

    grdColumns.RemoveAll
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    Me.Changed = True
    CheckIfScrollBarRequired
  
    ForceDefinitionToBeHiddenIfNeeded
  End If

End Sub

Private Sub cboChild_DropDown()
  msComboText = cboChild.Text

End Sub

Private Sub cboTables_Click()

  If Not mbLoading Then
    If Me.cboTables.Text = msComboText Then Exit Sub
    
    If grdColumns.Rows > 0 Then
      If COAMsgBox("Warning: Changing the base table will result in all table/column " & _
            "specific aspects of this " & LCase(Me.Caption) & " being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
        'TM20010910 Fault 2793
        'Commented out the following.
'        mbLoading = True
        cboTables.Text = msComboText
'        SetComboText cboTables, msComboText
'        mbLoading = False
        Exit Sub
      End If
    End If

    Me.Changed = True

    grdColumns.RemoveAll
    CheckIfScrollBarRequired
    
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    optAllRecords.Value = True

    txtPicklist.Text = ""
    txtPicklist.Tag = ""
    cmdPicklist.Enabled = False

    txtFilter.Text = ""
    txtFilter.Tag = ""
    cmdFilter.Enabled = False
  
  End If
    
  If typGlobal = glAdd Then
    GetChildTables cboTables.ItemData(cboTables.ListIndex)
  End If
  
  If Not mbLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
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
  
  ' Return false if some of the filters/picklists need to be removed from the definition,
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
          sBigMessage = "The '" & cboTables.List(cboTables.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtPicklist.Tag = 0
      txtPicklist.Text = "<None>"
      mblnRecordSelectionInvalid = True
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
          sBigMessage = "The '" & cboTables.List(cboTables.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTables.List(cboTables.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtFilter.Tag = 0
      txtFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Calcs
  With grdColumns
    If .Rows > 0 Then
      For iLoop = .Rows - 1 To 0 Step -1
        varBookmark = .AddItemBookmark(iLoop)
        strColumnType = .Columns(3).CellValue(varBookmark)
        lngColumnID = .Columns(4).CellValue(varBookmark)
        
        If strColumnType = "4" Then
          fRemove = False
          iResult = ValidateCalculation(lngColumnID)
  
          sCalcName = .Columns(1).CellValue(varBookmark)
          sCalcName = Mid(sCalcName, 2, Len(sCalcName) - 2)

          Select Case iResult
            Case REC_SEL_VALID_HIDDENBYUSER
              ' Calculation hidden by the current user.
              ' Only a problem if the current definition is NOT owned by the current user,
              ' or if the current definition is not already hidden.
              fRemove = (Not mblnDefinitionCreator) And _
                (Not mblnReadOnly) And _
                (Not FormPrint)

              If fRemove Then
                sBigMessage = "The '" & sCalcName & "' calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
                COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
              Else
                fNeedToForceHidden = True
  
                ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
                asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & sCalcName & "' calculation"
              End If

            Case REC_SEL_VALID_DELETED
              ' Calc deleted by another user.
              ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
              asDeletedParameters(UBound(asDeletedParameters)) = "'" & sCalcName & "' calculation"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)

            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                ' Calc hidden by another user.
                ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
                asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & sCalcName & "' calculation"
  
                fRemove = (Not mblnReadOnly) And _
                  (Not FormPrint)
              End If
            Case REC_SEL_VALID_INVALID
              ' Calc invalid.
              ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
              asInvalidParameters(UBound(asInvalidParameters)) = "'" & sCalcName & "' calculation"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)
          End Select

          If fRemove Then
            ' Calc invalid, deleted or hidden by another user. Remove it from this definition.
            If .Rows > 1 Then
              .RemoveItem iLoop
            Else
              .RemoveAll
            End If

            If Not FormPrint Then
              SSTab1.Tab = 1
              .SetFocus
            End If

            mblnRecordSelectionInvalid = True
          End If
        End If
      Next iLoop
    End If
  End With

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



Private Sub cboTables_DropDown()
  msComboText = cboTables.Text

End Sub

Private Sub cmdCancel_Click()
  Dim sMessage As String
  
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  If Me.Changed And Not mblnReadOnly Then
    'TM20010920 Fault 2849
    'The following had been commented out. Removed comments to display title for MSG.
    Select Case typGlobal
    Case glAdd
      sMessage = "Global Add"
    Case glUpdate
      sMessage = "Global Update"
    Case glDelete
      sMessage = "Global Delete"
    Case Else
      sMessage = "Global function"
    End Select
    'sMessage = GetCaption
    'strMBText = sMessage & " definition has changed.  Save changes ?"
    
    strMBText = "You have changed the current definition. Save changes ?"
    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
    strMBTitle = sMessage
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
  
  'Cancelled = True
  Me.Hide
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdClearAll_Click()
  If COAMsgBox("Clear all selected columns, are you sure ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Me.Changed = True
    grdColumns.RemoveAll
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    cmdEdit.Enabled = False
    ForceDefinitionToBeHiddenIfNeeded
    CheckIfScrollBarRequired
  End If

End Sub

Private Sub cmdDelete_Click()
  
  Dim lRow As Long
  
  '08/08/2000 MH Fault 2663 Remove COAMsgBox
  ' Delete the selected column definition.
  'If COAMsgBox("Delete selected column ?", vbQuestion + vbOKCancel, Me.Caption) = vbOK Then
    
    'grdColumns.DeleteSelected
    
    With grdColumns
    
      If .Rows = 1 Then
        .RemoveAll
      Else
        lRow = .AddItemRowIndex(.SelBookmarks(0))
        .RemoveItem lRow
        If lRow < .Rows Then
          .Bookmark = lRow
        Else
          .Bookmark = (.Rows - 1)
        End If
        .SelBookmarks.Add .Bookmark
      End If

    End With

    Me.Changed = True
    If grdColumns.Rows = 0 Then
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      cmdClearAll.Enabled = False
    End If
  'End If

  ForceDefinitionToBeHiddenIfNeeded
  CheckIfScrollBarRequired

End Sub

Private Sub cmdEdit_Click()
  ' Edit the selected new column definition.
  Dim lTableID As Long
  Dim lParentTableID As Long
  Dim lsiItem As ListItem
  Dim frmEdit As frmGlobalFunctionsColumn
'  Dim varBookmark As Variant
  Dim pstrRow As String
  Dim plngRow As Long
  
'  varBookmark = grdColumns.SelBookmarks(0)
'  grdColumns.Bookmark = varBookmark
  
  plngRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
  
  If typGlobal = glAdd Then
    lTableID = cboChild.ItemData(cboChild.ListIndex)
    lParentTableID = cboTables.ItemData(cboTables.ListIndex)
  Else
    lTableID = cboTables.ItemData(cboTables.ListIndex)
    lParentTableID = lTableID
  End If

  Set frmEdit = New frmGlobalFunctionsColumn
  
  With frmEdit
    .ParentForm = Me
    .Icon = Me.Icon
    .Initialise False, _
                lTableID, _
                typGlobal, _
                grdColumns.Columns(2).Value, _
                grdColumns.Columns(3).Value, _
                Val(grdColumns.Columns(4).Value), _
                grdColumns.Columns(1).Value, _
                Val(grdColumns.Columns(5).Value), _
                Val(grdColumns.Columns(6).Value), _
                lParentTableID
    .Show vbModal

    If Not .Cancelled Then
      Me.Changed = True
      
      'TM20020430 Fault 3778 - remove the old row and add the new data into that position in the
      'grid.
      pstrRow = .cboColumns.Text _
                & vbTab & .Value _
                & vbTab & .cboColumns.ItemData(.cboColumns.ListIndex) _
                & vbTab & .ValueType _
                & vbTab & .ValueID _
                & vbTab & .LookupTableID _
                & vbTab & .LookupColumnID
                
      grdColumns.RemoveItem plngRow
      grdColumns.AddItem pstrRow, plngRow

      grdColumns.Bookmark = grdColumns.AddItemBookmark(plngRow)
      grdColumns.SelBookmarks.RemoveAll
      grdColumns.SelBookmarks.Add grdColumns.AddItemBookmark(plngRow)

'      grdColumns.Columns(0).Text = .cboColumns.Text
'      grdColumns.Columns(1).Text = .Value
'      grdColumns.Columns(2).Text = .cboColumns.ItemData(.cboColumns.ListIndex)
'      grdColumns.Columns(3).Text = .ValueType
'      grdColumns.Columns(4).Text = .ValueID
'      grdColumns.Columns(5).Text = .LookupTableID
'      grdColumns.Columns(6).Text = .LookupColumnID
  
      ForceDefinitionToBeHiddenIfNeeded

    End If
  End With
   
  Unload frmEdit
  Set frmEdit = Nothing
  
End Sub

Private Sub cmdNew_Click()
  ' Add a new column definition.
  Dim lsiItem As ListItem
  Dim lParentTableID As Long
  Dim frmEdit As frmGlobalFunctionsColumn
  Dim varBookmark As Variant
  
  varBookmark = grdColumns.Bookmark
  
  Screen.MousePointer = vbHourglass

  If typGlobal = glAdd Then
    mlTableID = cboChild.ItemData(cboChild.ListIndex)
    lParentTableID = cboTables.ItemData(cboTables.ListIndex)
  Else
    mlTableID = cboTables.ItemData(cboTables.ListIndex)
    lParentTableID = mlTableID
  End If

  Set frmEdit = New frmGlobalFunctionsColumn
  
  With frmEdit
    .ParentForm = Me
    .Icon = Me.Icon
    .Initialise True, mlTableID, typGlobal, , , , , , , lParentTableID
    Screen.MousePointer = vbDefault
    
    If .cboColumns.ListCount = 0 Then
      COAMsgBox "All of the available columns have already been populated", vbInformation
      Exit Sub
    
    Else
    
      .Show vbModal

      If Not .Cancelled Then
      
        Me.Changed = True

        grdColumns.AddItem _
          Trim(.cboColumns.Text) & vbTab & _
          .Value & vbTab & _
          .cboColumns.ItemData(.cboColumns.ListIndex) & vbTab & _
          .ValueType & vbTab & _
          .ValueID & vbTab & _
          .LookupTableID & vbTab & _
          .LookupColumnID
        
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        cmdClearAll.Enabled = True
      
        grdColumns.Bookmark = grdColumns.Rows - 1
        grdColumns.SelBookmarks.Add grdColumns.Bookmark
        'cmdOk.Enabled = True

        ForceDefinitionToBeHiddenIfNeeded

      Else
        grdColumns.Bookmark = varBookmark

      End If
    End If

  End With

  Unload frmEdit
  Set frmEdit = Nothing

  CheckIfScrollBarRequired

End Sub

Private Sub cmdOK_Click()
  ' Save and exit.
  If SaveDefinition Then
    Unload Me
  End If

  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPicklist_Click()
  ' Display the picklist selection form.
  Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim frmSelection As frmDefSel
  Dim rsTemp As Recordset
  
  Screen.MousePointer = vbHourglass
  fExit = False

  'sSQL = "SELECT name, pickListID FROM ASRSysPickListName"
  'sSQL = sSQL & " WHERE
  'sSQL = "tableID = " & cboTables.ItemData(cboTables.ListIndex)
  'If cboChild.Visible And cboChild.Enabled Then
  '    sSQL = sSQL & " OR tableID = " & cboChild.ItemData(cboChild.ListIndex)
  'End If

  Set frmSelection = New frmDefSel
  With frmSelection
    .SelectedUtilityType = utlPicklist
    .TableID = cboTables.ItemData(cboTables.ListIndex)
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(txtPicklist.Tag) > 0 Then
      .SelectedID = Val(txtPicklist.Tag)
    End If
      
    ' Loop until a picklist has been selected or cancelled.
    Do While Not fExit

      If .ShowList(utlPicklist) Then
        
        ' Display the selection form.
        .Show vbModal
  
        Select Case .Action
          Case edtAdd
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(True, False, cboTables.ItemData(cboTables.ListIndex)) Then
              frmPick.Show vbModal
            End If
            frmSelection.SelectedID = frmPick.SelectedID
            Set frmPick = Nothing

          Case edtEdit
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(False, .FromCopy, cboTables.ItemData(cboTables.ListIndex), .SelectedID) Then
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
            
            txtPicklist.Text = .SelectedText
            txtPicklist.Tag = .SelectedID
            txtFilter.Text = ""
            txtFilter.Tag = 0

            fExit = True

          Case 0
            fExit = True
  
        End Select
      End If
    Loop
  End With
  
  Set frmSelection = Nothing
  Set frmPick = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub Form_Load()
  
  Set datData = New DataMgr.clsDataAccess
  
  SSTab1.Tab = 0

  grdAccess.RowHeight = 239

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode And Not FormPrint Then
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
  'Set datGlobal = Nothing

  'RH 16/05/00 - To prevent the lockups of the toolbars after utility usage
  '              Shouldnt need cos its in defsel, but lets see if it solves the problem
  '              (Fault 208)
  With frmMain.abMain
    .ResetHooks
    .Refresh
  End With
  
End Sub

Private Function GetCaption() As String
  ' Return the Global Function title.
  Select Case typGlobal
    Case glAdd
      GetCaption = "Global Add"
      Me.HelpContextID = 1041
    
    Case glUpdate
      GetCaption = "Global Update"
      Me.HelpContextID = 1042
    
    Case Else
      GetCaption = "Global Delete"
      Me.HelpContextID = 1043
    
    End Select

End Function

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

Private Sub grdAccess_GotFocus()
  grdAccess.Col = 1

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


Private Sub grdColumns_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
DispPromptMsg = 0
End Sub

Private Sub grdColumns_DblClick()
  If cmdEdit.Enabled Then
    cmdEdit_Click
  Else
    cmdNew_Click
  End If
End Sub


Private Sub optAllRecords_Click()

  cmdFilter.Enabled = False
  txtFilter.Tag = ""
  txtFilter.Text = ""

  cmdPicklist.Enabled = False
  txtPicklist.Tag = ""
  txtPicklist.Text = ""

  ForceDefinitionToBeHiddenIfNeeded

  Me.Changed = True

End Sub

Private Sub optFilter_Click()

  cmdFilter.Enabled = True
  
  cmdPicklist.Enabled = False
  txtPicklist.Tag = 0
  txtPicklist.Text = ""

  txtFilter.Text = "<None>"

  If Not mbLoading Then
    ForceDefinitionToBeHiddenIfNeeded
    Me.Changed = True
  End If
  
End Sub

Private Sub optPicklist_Click()

  cmdPicklist.Enabled = True
  
  cmdFilter.Enabled = False
  txtFilter.Tag = 0
  txtFilter.Text = ""

  txtPicklist.Text = "<None>"

  If Not mbLoading Then
    ForceDefinitionToBeHiddenIfNeeded
    Me.Changed = True
  End If
  
End Sub

Public Property Let FromCopy(bFromCopy As Boolean)

    mbFromCopy = bFromCopy

End Property

Public Property Get FromCopy() As Boolean

    FromCopy = mbFromCopy

End Property

'Public Property Get Cancelled() As Boolean
'
'    Cancelled = mbCancelled
'
'End Property

'Public Property Let Cancelled(ByVal bCancel As Boolean)
'
'    mbCancelled = bCancel
'
'End Property

Public Property Get FormType() As GlobalType

    FormType = typGlobal

End Property

Public Property Let FormType(typForm As GlobalType)

    typGlobal = typForm

End Property

Public Property Get FormTypeName() As String

    Select Case typGlobal
    Case glAdd
      FormTypeName = "Global Add"
    Case glUpdate
      FormTypeName = "Global Update"
    Case glDelete
      FormTypeName = "Global Delete"
    Case Else
      FormTypeName = "Global function"
    End Select

End Property

Public Sub PickListSelected(lPicklistID As Long, sPicklist As String)

    txtPicklist = sPicklist
    txtPicklist.Tag = lPicklistID

End Sub

Private Sub cmdFilter_Click()
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression

  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression

  With objExpression
    ' Initialise the expression object.
    fOK = .Initialise(cboTables.ItemData(cboTables.ListIndex), Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    
    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) Then
  
        txtFilter.Text = IIf(Len(.Name) = 0, "<None>", .Name)
        txtFilter.Tag = .ExpressionID
        
        Changed = True
       
      End If
    End If
  End With

  Set objExpression = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Function SaveDefinition() As Boolean

  Dim sMsg As String
  Dim lChildID As Long
  Dim lFilterID As Long
  Dim lPicklistID As Long
  Dim lCount As Long
  Dim alngColumns() As Long
  Dim iNextIndex As Integer
  Dim lngColumnID As Long
  Dim iLoop As Integer
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean
  Dim sSQL As String
  Dim strName As String
  Dim strRecSelStatus As String
  Dim iUtilityType As UtilityType

  Dim lngColumnType As Long
  Dim strColumnValue As String

  Dim iCount_Owner As Integer
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobScheduledUserGroups As String
  Dim sHiddenGroups As String

  fBatchJobsOK = True

  strName = Trim(txtName.Text)
  fBatchJobsOK = True

  On Error GoTo Err_Trap
  
  'Check if mandatory columns have been ommitted on global adds
  If typGlobal = glAdd Then
    If CheckMandatoryColumns = False Then
      Exit Function
    End If
  End If

  'Check if another user have amended the definition since
  'the current user read it
  Select Case typGlobal
    Case glAdd
      iUtilityType = utlGlobalAdd
    Case glDelete
      iUtilityType = utlGlobalDelete
    Case glUpdate
      iUtilityType = utlGlobalUpdate
  End Select

  Call UtilityAmended(utlGlobalAdd, mlFunctionID, mlngTimeStamp, blnContinueSave, blnSaveAsNew)

  If blnContinueSave = False Then
    SaveDefinition = False
    Exit Function
  ElseIf blnSaveAsNew Then
    txtUserName = gsUserName
    
    mblnDefinitionCreator = True
    mlFunctionID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  sMsg = ""
  
  If Len(strName) > 0 Then
    If mlFunctionID > 0 Then
      If CheckForExistingName(False, strName, Mid$(Me.Caption, 8, 1)) Then
        'sMsg = "There already exists a " & Me.Caption & " definition with this name. Please enter another."
        sMsg = "A " & Me.Caption & " definition called '" & Trim(txtName.Text) & "' already exists."
      End If
    Else
      If CheckForExistingName(True, strName, Mid$(Me.Caption, 8, 1)) Then
        'sMsg = "There already exists a " & Me.Caption & " definition with this name. Please enter another."
        sMsg = "A " & Me.Caption & " definition called '" & Trim(txtName.Text) & "' already exists."
      End If
    End If
  Else
    sMsg = "You must give this definition a name."
  End If
    
  If Len(sMsg) > 0 Then
    SSTab1.Tab = 0
    COAMsgBox sMsg, vbExclamation, Me.Caption
    With txtName
      If .Enabled Then
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
      End If
    End With
    SaveDefinition = False
    Exit Function
  End If
     
  If typGlobal <> glDelete Then
    If grdColumns.Rows = 0 Then
     'NHRD14032012 JIRA HRPRO-1852
      SSTab1.Tab = 1
      COAMsgBox "No columns selected.", vbExclamation, Me.Caption
      SaveDefinition = False
      Exit Function
    End If
  End If
  
  If optPicklist Then
    If Val(txtPicklist.Tag) = 0 Then
      SSTab1.Tab = 0
      COAMsgBox "No picklist selected.", vbExclamation, Me.Caption
      SaveDefinition = False
      Exit Function
    Else
      lPicklistID = txtPicklist.Tag
    End If
  ElseIf optFilter Then
    If Val(txtFilter.Tag) = 0 Then
      SSTab1.Tab = 0
      COAMsgBox "No filter selected.", vbExclamation, Me.Caption
      SaveDefinition = False
      Exit Function
    Else
      lFilterID = txtFilter.Tag
    End If
  End If

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    SaveDefinition = False
    Exit Function
  End If
  
If mlFunctionID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    Select Case typGlobal
    Case glAdd
      CheckCanMakeHiddenInBatchJobs utlGlobalAdd, _
        CStr(mlFunctionID), _
        txtUserName.Text, _
        iCount_Owner, _
        sBatchJobDetails_Owner, _
        sBatchJobIDs, _
        sBatchJobDetails_NotOwner, _
        fBatchJobsOK, _
        sBatchJobDetails_ScheduledForOtherUsers, _
        sBatchJobScheduledUserGroups, _
        sHiddenGroups
    Case glUpdate
      CheckCanMakeHiddenInBatchJobs utlGlobalUpdate, _
        CStr(mlFunctionID), _
        txtUserName.Text, _
        iCount_Owner, _
        sBatchJobDetails_Owner, _
        sBatchJobIDs, _
        sBatchJobDetails_NotOwner, _
        fBatchJobsOK, _
        sBatchJobDetails_ScheduledForOtherUsers, _
        sBatchJobScheduledUserGroups, _
        sHiddenGroups
    Case glDelete
      CheckCanMakeHiddenInBatchJobs utlGlobalDelete, _
        CStr(mlFunctionID), _
        txtUserName.Text, _
        iCount_Owner, _
        sBatchJobDetails_Owner, _
        sBatchJobIDs, _
        sBatchJobDetails_NotOwner, _
        fBatchJobsOK, _
        sBatchJobDetails_ScheduledForOtherUsers, _
        sBatchJobScheduledUserGroups, _
        sHiddenGroups
    End Select
    
    If (Not fBatchJobsOK) Then
      If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
        COAMsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, Me.Caption
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , Me.Caption
      End If

      Screen.MousePointer = vbDefault
      SSTab1.Tab = 0
      SaveDefinition = False
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Screen.MousePointer = vbDefault
        SSTab1.Tab = 0
        SaveDefinition = False
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
  
  If typGlobal = glAdd Then
    lChildID = cboChild.ItemData(cboChild.ListIndex)
  End If

  If mlFunctionID > 0 Then
    AmendFunction strName, txtDesc, cboTables.ItemData(cboTables.ListIndex), _
      optAllRecords.Value, lFilterID, lPicklistID, lChildID
    
    sSQL = "Delete From ASRSysGlobalItems Where FunctionID = " & mlFunctionID
    datData.ExecuteSql sSQL
  
    Call UtilUpdateLastSaved(iUtilityType, mlFunctionID)
  
  Else
    InsertFunction strName, txtDesc, Mid$(Me.Caption, 8, 1), cboTables.ItemData(cboTables.ListIndex), _
      optAllRecords, lFilterID, lPicklistID, lChildID
  
    Call UtilCreated(iUtilityType, mlFunctionID)
    
  End If
  
  SaveAccess
  SaveObjectCategories cboCategory, iUtilityType, mlFunctionID
  

  With grdColumns
    'TM20010807 Fault 2617 - needs to redraw so SaveDefinition can be called before showing the form.
    .Redraw = True
    For lCount = 0 To .Rows - 1
      
      '01/08/2001 MH Fault 2116
      '.Bookmark = lCount
      .Bookmark = .AddItemBookmark(lCount)

      lngColumnType = datGeneral.GetDataType(cboTables.ItemData(cboTables.ListIndex), .Columns(2).Value)
      strColumnValue = .Columns(1).Text
  
      If lngColumnType = sqlNumeric Or lngColumnType = sqlInteger Then
        strColumnValue = Replace(datGeneral.ConvertNumberForSQL(strColumnValue), ",", "")
      End If
      
      InsertFunctionItems .Columns(2).Text, .Columns(3).Text, .Columns(4).Text, _
        Val(.Columns(5).Text), Val(.Columns(6).Text), strColumnValue
    Next
  End With
  
  mbFromCopy = False
  SaveDefinition = True
  Exit Function

Err_Trap:
  COAMsgBox Err.Description
  SaveDefinition = False
  Screen.MousePointer = vbDefault

End Function

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysGlobalAccess WHERE ID = " & mlFunctionID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysGlobalAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlFunctionID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysGlobalAccess" & _
        " WHERE ID = " & CStr(mlFunctionID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysGlobalAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlFunctionID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub




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




Private Function RetreiveDefinition() As Boolean
  
  Dim rsTemp As Recordset
  Dim lsiItem As ListItem
  Dim ctlTemp As Control
  Dim lngDataType As Long
  Dim strRecSelStatus As String
  Dim sMessage As String
  Dim strValue As String
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim iUtilityType As UtilityType
  
  Dim strTemp As String
  
  Dim fAlreadyNotified As Boolean
  
  On Error GoTo LocalErr
  
  mbLoading = True
  fAlreadyNotified = False
  
  Set rsTemp = GetFunctionDetails
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, GetCaption
    fOK = False
    Exit Function
  End If

  mlngTimeStamp = rsTemp!intTimestamp
  
  ' Treat the Global Function as read-only if the user does not have permission to edit them.
  Select Case typGlobal
  Case glAdd
    iUtilityType = utlGlobalAdd
    mblnReadOnly = Not datGeneral.SystemPermission("GLOBALADD", "EDIT")
    If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
      mblnReadOnly = (CurrentUserAccess(utlGlobalAdd, mlFunctionID) = ACCESS_READONLY)
    End If
  Case glDelete
    iUtilityType = utlGlobalDelete
    mblnReadOnly = Not datGeneral.SystemPermission("GLOBALDELETE", "EDIT")
    If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
      mblnReadOnly = (CurrentUserAccess(utlGlobalDelete, mlFunctionID) = ACCESS_READONLY)
    End If
  Case glUpdate
    iUtilityType = utlGlobalUpdate
    mblnReadOnly = Not datGeneral.SystemPermission("GLOBALUPDATE", "EDIT")
    If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
      mblnReadOnly = (CurrentUserAccess(utlGlobalUpdate, mlFunctionID) = ACCESS_READONLY)
    End If
  End Select
        
  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)
  GetObjectCategories cboCategory, iUtilityType, mlFunctionID
  
  If mbFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
    'TM04062004 Fault 8749 - cannot be ReadOnly if is the definition owner.
    mblnReadOnly = False
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(Trim$(rsTemp!userName), vbProperCase)
    mblnDefinitionCreator = (LCase$(Trim$(rsTemp!userName)) = LCase$(gsUserName))
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
  
'TM20020404 Fault 3729 - Uncommented this section of code,
  cboTables.Text = rsTemp!TableName
  mlTableID = rsTemp!TableID

  If typGlobal = glAdd Then
    mlTableID = rsTemp!ChildTableID
    If cboChild.Enabled Then
      'cboChild.Text = rsTemp!childTableName
      SetComboItem cboChild, rsTemp!ChildTableID
    End If
  End If

  If typGlobal <> glDelete Then
    Do While Not rsTemp.EOF
      
      Select Case rsTemp!ValueType
        Case globfuncvaltyp_STRAIGHTVALUE
          ' Straight value.
          ' NB. date values are stored as 'mm/dd/yyy' in the database. We need to reformat to the locale's
          ' format for display.
          'lngDataType = datGlobal.GetDataType(rsTemp!ColumnID)
          lngDataType = GetDataType(rsTemp!ColumnID)
          
          If lngDataType = sqlDate Then
            If IsNull(rsTemp!Value) Then
              
              'grdColumns.AddItem rsTemp!ColumnName & vbTab & ConvertSQLDateToLocale("__/__/____") & vbTab & rsTemp!ColumnID &
              grdColumns.AddItem rsTemp!ColumnName & vbTab & "null" & vbTab & rsTemp!ColumnID & _
              vbTab & rsTemp!ValueType & vbTab & "0" & vbTab & "0" & vbTab & "0"
            Else
              grdColumns.AddItem rsTemp!ColumnName & vbTab & ConvertSQLDateToLocale(rsTemp!Value) & vbTab & rsTemp!ColumnID & _
              vbTab & rsTemp!ValueType & vbTab & "0" & vbTab & "0" & vbTab & "0"
            End If
          Else
            'TM20020910 Fault 4395 - Don't RTrim the value.
'            grdColumns.AddItem rsTemp!ColumnName & vbTab & RTrim(rsTemp!Value) & vbTab & rsTemp!ColumnID & _
'            vbTab & rsTemp!ValueType & vbTab & "0" & vbTab & "0" & vbTab & "0"
            If datGeneral.DoesColumnUseSeparators(rsTemp!ColumnID) Then
  
              ' Thousand separators
              strTemp = Replace(rsTemp!Value, UI.GetSystemThousandSeparator, "")
      
              ' Decimals
              'IIf(iCount <> (InStr(1, rsTemp!Value, ".") - 1), ",", "") &
              strValue = ""
              iCount2 = 1
              If InStr(1, strTemp, ".") > 0 Then
                For iCount = InStr(1, strTemp, ".") - 1 To 1 Step -1
                  strValue = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(strTemp, iCount, 1) & strValue
                  iCount2 = iCount2 + 1
                Next iCount
                strValue = strValue & "." & Right(strTemp, Len(strTemp) - InStr(1, strTemp, "."))
              Else
                For iCount = Len(strTemp) To 1 Step -1
                  strValue = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(strTemp, iCount, 1) & strValue
                  iCount2 = iCount2 + 1
                Next iCount
              End If
            Else
              strValue = rsTemp!Value
            End If

            
            grdColumns.AddItem rsTemp!ColumnName & vbTab & strValue & vbTab & rsTemp!ColumnID & _
            vbTab & rsTemp!ValueType & vbTab & "0" & vbTab & "0" & vbTab & "0"
          End If
            
        Case globfuncvaltyp_LOOKUPTABLE
          ' Value from a lookup table.
          'grdColumns.AddItem rsTemp!ColumnName & vbTab & _
                             datGlobal.GetGlobalTableValue(rsTemp!tableValueID) & vbTab & _
                             rsTemp!ColumnID & vbTab & _
                             rsTemp!ValueType & vbTab & _
                             rsTemp!tableValueID
          'JPD 20041209 Fault 9613
          lngDataType = GetDataType(rsTemp!ColumnID)
          
          If lngDataType = sqlDate Then
            If IsNull(rsTemp!Value) Then
              grdColumns.AddItem rsTemp!ColumnName & vbTab & _
                "null" & vbTab & _
                rsTemp!ColumnID & vbTab & _
                rsTemp!ValueType & vbTab & "0" & vbTab & _
                rsTemp!LookupTableID & vbTab & _
                rsTemp!LookupColumnID
            Else
              grdColumns.AddItem rsTemp!ColumnName & vbTab & _
                ConvertSQLDateToLocale(rsTemp!Value) & vbTab & _
                rsTemp!ColumnID & vbTab & _
                rsTemp!ValueType & vbTab & "0" & vbTab & _
                rsTemp!LookupTableID & vbTab & _
                rsTemp!LookupColumnID
            End If
          Else
            grdColumns.AddItem rsTemp!ColumnName & vbTab & _
              Trim(rsTemp!Value) & vbTab & _
              rsTemp!ColumnID & vbTab & _
              rsTemp!ValueType & vbTab & "0" & vbTab & _
              rsTemp!LookupTableID & vbTab & _
              rsTemp!LookupColumnID
          End If

        Case globfuncvaltyp_FIELD
          ' Field value.
          grdColumns.AddItem _
              rsTemp!ColumnName & vbTab & _
              "<" & GetColumnName(rsTemp!refColumnID) & ">" & vbTab & _
              rsTemp!ColumnID & vbTab & _
              rsTemp!ValueType & vbTab & _
              rsTemp!refColumnID & vbTab & _
              "0" & vbTab & _
              "0" & vbTab & _
              "0"
          
        Case globfuncvaltyp_CALCULATION
          ' Calculated value.
            grdColumns.AddItem _
                rsTemp!ColumnName & vbTab & _
                "<" & rsTemp!exprName & ">" & vbTab & _
                rsTemp!ColumnID & vbTab & _
                rsTemp!ValueType & vbTab & _
                rsTemp!ExprID & vbTab & _
                "0" & vbTab & _
                "0" & vbTab & _
                "0"
'''''          End If

      End Select
      rsTemp.MoveNext
    Loop

  End If
  
  If grdColumns.Rows > 0 Then
    grdColumns.MoveFirst
    grdColumns.SelBookmarks.Add grdColumns.Bookmark
  End If

  If mblnReadOnly And Not mbFromCopy Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdClearAll.Enabled = True
  End If
  grdAccess.Enabled = True

  rsTemp.Close
  Set rsTemp = Nothing

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Cancelled = True
    RetreiveDefinition = False
    Exit Function
  End If
  
  mbLoading = False

  RetreiveDefinition = True

Exit Function

LocalErr:
  COAMsgBox "Error Retrieving Definition" & vbCrLf & "(" & Err.Description & ")", vbExclamation
  RetreiveDefinition = False

End Function

'Private Sub DisableAll()
'  Dim ctlTemp As Control
'
'  For Each ctlTemp In Me.Controls
'    If (Not TypeOf ctlTemp Is Label) And _
'      (Not TypeOf ctlTemp Is Frame) And _
'      (Not TypeOf ctlTemp Is SSTab) Then
'
'      ctlTemp.Enabled = False
'    End If
'  Next ctlTemp
'  Set ctlTemp = Nothing
'
'  cmdCancel.Enabled = True
'
'End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Sub DeleteRecord()
  cmdDelete_Click

End Sub

Public Sub AddNew()

    cmdNew_Click

End Sub

Public Sub DeleteAll()

'  If COAMsgBox("Delete all columns, are you sure ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'    grdColumns.RemoveAll
'    cmdDelete.Enabled = False
'    cmdClearAll.Enabled = False
'    cmdEdit.Enabled = False
'    cmdOk.Enabled = False
'  End If

End Sub

Public Sub EditRecord()

    cmdEdit_Click

End Sub

Private Function Exists(bNew As Boolean, lColumnID As Long) As Boolean

    Dim lCount As Long
    Dim lSelected As Long
    Dim vOldBookmark As Variant
    
    Exists = False
    vOldBookmark = grdColumns.Bookmark
    
    If bNew Then
      For lCount = 0 To grdColumns.Rows - 1
        grdColumns.Bookmark = lCount
        If grdColumns.Columns(2).Text = lColumnID Then
          COAMsgBox "This column already exists in the global function definition.", vbExclamation, Me.Caption
          grdColumns.Bookmark = vOldBookmark
          grdColumns.SelBookmarks.Add grdColumns.Bookmark
          Exists = True
          Exit Function
        End If
      Next
      grdColumns.Bookmark = vOldBookmark
      grdColumns.SelBookmarks.Add grdColumns.Bookmark
    Else
        lSelected = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
        For lCount = 0 To grdColumns.Rows - 1
          grdColumns.Bookmark = lCount
          If lSelected <> lCount Then
            If grdColumns.Columns(2).Text = lColumnID Then
              COAMsgBox "This column already exists in the global function definition.", vbExclamation, Me.Caption
              grdColumns.Bookmark = vOldBookmark
              grdColumns.SelBookmarks.Add grdColumns.Bookmark
              Exists = True
              Exit Function
            End If
          End If
        Next
        grdColumns.Bookmark = vOldBookmark
        grdColumns.SelBookmarks.Add grdColumns.Bookmark
    End If
      
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
  fraDefinition(0).Enabled = (SSTab1.Tab = giTABSTRIP_GLOBALFUNCTIONDEF)
  fraDefinition(1).Enabled = fraDefinition(0).Enabled
  fraColumns.Enabled = (SSTab1.Tab = giTABSTRIP_COLUMNDEF)
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
  ' Select all text.
  UI.txtSelText

End Sub



Private Sub FormatForm()
  ' Format the screen controls depending on the type of Global function being defined.
  Dim fHasChildTable As Boolean
  Dim fHasColumns As Boolean
  Dim lngLeft As Long
  Dim lngTop As Long

  Const GAP = 150

  fHasChildTable = (typGlobal = glAdd)
  fHasColumns = Not (typGlobal = glDelete)

  ' Only display Child Table controls if required.
  lblTitle(1) = IIf(fHasChildTable, "Parent Table :", "Base Table :")
  lblTitle(2).Visible = fHasChildTable
  cboChild.Visible = fHasChildTable
  GetTables fHasChildTable
    
    
  If Not fHasColumns Then
    
    'Frame 0
    With fraDefinition(0)
      Set .Container = Me
      .Move GAP, GAP
    End With
    
    'Frame 1
    With fraDefinition(1)
      Set .Container = Me
      .Move GAP, fraDefinition(0).Top + fraDefinition(0).Height + GAP

      lngLeft = .Left + .Width
      lngTop = .Top + .Height
      Me.ScaleWidth = lngLeft + GAP
      Me.ScaleHeight = lngTop + cmdCancel.Height + GAP

      'Command Cancel
      lngLeft = .Left + .Width - cmdCancel.Width
      cmdCancel.Move lngLeft, lngTop

      'Command OK
      lngLeft = lngLeft - (cmdOK.Width + GAP)
      cmdOK.Move lngLeft, lngTop
    End With

    SSTab1.Visible = False

  End If

End Sub


Private Function CheckMandatoryColumns() As Boolean
        
  Dim rsColumns As Recordset
  Dim strSQL As String
  Dim datData As clsDataAccess

  Dim lngRow As Long
  Dim pvarbookmark As Variant
  Dim strMandatoryColumns As String
  Dim strColumnIDs As String

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strTitle As String
  
  strMandatoryColumns = vbNullString
  strColumnIDs = vbNullString
  
  With grdColumns
    .Row = 0
    For lngRow = 0 To .Rows - 1
      pvarbookmark = .GetBookmark(lngRow)
      strColumnIDs = strColumnIDs & _
        IIf(strColumnIDs <> "", ", ", "") & .Columns(2).CellText(pvarbookmark)
    Next
  End With

  'MH20000904
  'Allow save if mandatory ommitted if it has a default value
  'This is to get around the staff number on a applicants to personnel transfer
  
  'Allow save if mandatory ommitted and it is a calculated column

  '******************************************************************************
  ' TM20010719 Fault 2242 - ColumnType <> 4 clause added to ignore all linked   *
  ' columns. (It doesn't need to validate the linked columns because this is    *
  ' done using the Vaidate SP.                                                  *
  '******************************************************************************

  strSQL = "SELECT ASRSysTables.TableName, ASRSysColumns.ColumnName " & _
           "FROM ASRSysColumns " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
           "WHERE ASRSysColumns.TableID = " & CStr(cboChild.ItemData(cboChild.ListIndex)) & _
           " AND " & SQLWhereMandatoryColumn
           '"  AND (Rtrim(DefaultValue) = '' OR (Rtrim(DefaultValue) = '__/__/____') and DataType = 11)" & _
           "  AND Convert(int,isnull(dfltValueExprID,0)) = 0 " & _
           "  AND CalcExprID = 0 " & _
           "  AND Mandatory = '1' " & _
           "  AND ColumnType <> 4 "

  If grdColumns.Rows > 0 Then
    strSQL = strSQL & _
           "  AND ASRSysColumns.ColumnID NOT IN (" & strColumnIDs & ") "
  End If
           
  strSQL = strSQL & _
           "ORDER BY ASRSysColumns.ColumnName"

  Set datData = New clsDataAccess
  Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  Set datData = Nothing
    
  Do While Not rsColumns.EOF
    strMandatoryColumns = strMandatoryColumns & _
      rsColumns!TableName & "." & rsColumns!ColumnName & vbCrLf
    rsColumns.MoveNext
  Loop
  
  
  CheckMandatoryColumns = (strMandatoryColumns = vbNullString)
  
  If CheckMandatoryColumns = False Then
    strMBText = "Unable to save definition as the following mandatory" & vbCrLf & _
                "columns have not been populated:" & vbCrLf & vbCrLf & _
                strMandatoryColumns & vbCrLf & _
                "Please enter a source to populate these columns."
    intMBButtons = vbExclamation + vbOKOnly
    strTitle = app.ProductName
    'NHRD14032012 JIRA HRPRO-1852
    SSTab1.Tab = 1
    COAMsgBox strMBText, intMBButtons, strTitle
  End If

End Function


Public Sub PrintDef(typeGlobal As GlobalType, lFunctionID As Long)
  
  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsColumns As Recordset
  Dim sSQL As String
  Dim strType As String
  Dim lngDataType As Long
  Dim iLoop As Integer
  Dim strSource As String
  Dim strDestin As String
  Dim varBookmark As Variant
  Dim iUtilityType As UtilityType

  Set datData = New DataMgr.clsDataAccess
  'Set datGlobal = New DataMgr.clsGlobal
  
  typGlobal = typeGlobal
  strType = GetCaption
  
  mlFunctionID = lFunctionID
  Set rsTemp = GetFunctionDetails
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation, GetCaption
    Exit Sub
  End If
  
  Select Case typGlobal
    Case glAdd
      iUtilityType = utlGlobalAdd
    Case glDelete
      iUtilityType = utlGlobalDelete
    Case glUpdate
      iUtilityType = utlGlobalUpdate
  End Select
  
  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader strType & " : " & rsTemp!Name
    
        .PrintNormal "Category : " & GetObjectCategory(iUtilityType, mlFunctionID)
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal "Owner : " & rsTemp!userName
        
        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop
        
        ' Data section --------------------------------------------------------
        .PrintTitle "Data"
        
        .PrintNormal IIf(typGlobal = glAdd, "Parent Table : ", "Base Table : ") & _
                     rsTemp!TableName
        If rsTemp!AllRecords Then
          .PrintNormal "Records : All Records"
        ElseIf rsTemp!FilterID > 0 Then
          '.PrintNormal "Records : Filter '" & IIf(IsNull(rsTemp!FilterName), "<Deleted>", rsTemp!FilterName) & "'"
          .PrintNormal "Records : '" & datGeneral.GetFilterName(rsTemp!FilterID) & "' filter"
        Else
          '.PrintNormal "Records : Picklist '" & IIf(IsNull(rsTemp!PicklistName), "<Deleted>", rsTemp!PicklistName) & "'"
          .PrintNormal "Records : '" & datGeneral.GetPicklistName(rsTemp!PicklistID) & "' picklist"
        End If
        .PrintNormal
        
        
        If typGlobal = glAdd Then
          .PrintNormal "Child Table : " & rsTemp!ChildTableName
          .PrintNormal
        End If
        
        '---------
        
        If typGlobal <> glDelete Then
    
          .PrintTitle "Columns"
          .PrintBold "Destination Column" & vbTab & _
                     "Source Column / Data"
  
          Do While Not rsTemp.EOF
    
            Select Case rsTemp!ValueType
            Case globfuncvaltyp_STRAIGHTVALUE
              ' Straight value.
              ' NB. date values are stored as 'mm/dd/yyy' in the database. We need to reformat to the locale's
              ' format for display.
              lngDataType = GetDataType(rsTemp!ColumnID)
              If lngDataType <> sqlDate Then
                strSource = Chr(34) & Trim(rsTemp!Value) & Chr(34)
              ElseIf IsNull(rsTemp!Value) Then
                'strSource = Chr(34) & ConvertSQLDateToLocale("__/__/____") & Chr(34)
                strSource = "null"
              Else
                strSource = Chr(34) & ConvertSQLDateToLocale(rsTemp!Value) & Chr(34)
              End If
                
            Case globfuncvaltyp_LOOKUPTABLE
                                 'rsDetails!LookupTableID & vbTab & _
                                 'rsDetails!LookupColumnID
              strSource = Chr(34) & Trim(rsTemp!Value) & Chr(34)
    
            Case globfuncvaltyp_FIELD
              ' Field value.
              strSource = rsTemp!TableName & "." & GetColumnName(rsTemp!refColumnID)
    
            Case globfuncvaltyp_CALCULATION
              ' Calculated value.
              strSource = IIf(IsNull(rsTemp!exprName), "<Deleted Calculation>", rsTemp!exprName)
  
            End Select
            
            If typGlobal = glAdd Then
              strDestin = rsTemp!ChildTableName & "." & rsTemp!ColumnName
            Else
              strDestin = rsTemp!TableName & "." & rsTemp!ColumnName
            End If
    
            .PrintNonBold strDestin & vbTab & strSource
            
            rsTemp.MoveNext
          Loop
    
        End If
    
        '---------
    
        .PrintEnd
    
        rsTemp.MoveFirst
        .PrintConfirm strType & " : " & rsTemp!Name, strType & " Definition"
      End If
  
    End With
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing
  'Set datGlobal = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing " & strType & " Definition Failed", vbExclamation

End Sub

Private Function CheckForExistingName(bNew As Boolean, sName As String, sGlobalType As String) As Boolean

    Dim rsName As Recordset
    Dim sSQL As String
    
    sSQL = "Select * From ASRSysGlobalFunctions " & _
           "Where Name = '" & Replace(sName, "'", "''") & "'" & _
           "And Type = '" & sGlobalType & "'"
    
    If Not bNew Then
      sSQL = sSQL & " AND FunctionID <> " & mlFunctionID
    End If
    
    Set rsName = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    If rsName.EOF And rsName.BOF Then
        CheckForExistingName = False
    Else
        CheckForExistingName = True
    End If
    
    rsName.Close
    Set rsName = Nothing

End Function


Private Function InsertFunction(sName As String, sDesc As String, sType As String, lTableID As Long, bAllRecords As Boolean, _
        lFilterID As Long, lPicklistID As Long, Optional lChildID As Long) As Long
        
    Dim sSQL As String
    
    sSQL = "INSERT INTO ASRSysGlobalFunctions" & _
      " (name, description, type, tableID, childTableID, allRecords, filterID, pickListID, userName)" & _
      " VALUES(" & _
      "'" & Replace(sName, "'", "''") & "', " & _
      "'" & Replace(sDesc, "'", "''") & "', " & _
      "'" & sType & "', " & _
      lTableID & ", " & _
      lChildID & ", " & _
      CLng(bAllRecords) & ", " & _
      lFilterID & ", " & _
      lPicklistID & ", " & _
      "'" & datGeneral.UserNameForSQL & "')"
    
    ' RH 04/09/00 - Use the new util def stored procedure
    mlFunctionID = InsertGlobalFunction(sSQL)
            
End Function

Private Function InsertGlobalFunction(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertGlobalFunction_ERROR

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
    pmADO.Value = "AsrSysGlobalFunctions"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "FunctionID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertGlobalFunction = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertGlobalFunction = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertGlobalFunction_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function

Private Function InsertFunctionItems(lColumnID As Long, lValueType As Long, _
    lValueID As Long, lLookupTableID As Long, lLookupColumnID As Long, Optional sValue As String)

    Dim fOK As Boolean
    Dim lngDataType As Long
    Dim sSQL As String
    Dim strTemp As String
    
    fOK = True
    
    sSQL = "INSERT INTO ASRSysGlobalItems (" & _
              "functionID, " & _
              "columnID, " & _
              "valueType, " & _
              "exprID, " & _
              "value, " & _
              "refColumnID, " & _
              "LookupTableID, " & _
              "LookupColumnID) " & _
           "VALUES(" & _
              mlFunctionID & ", " & _
              lColumnID & ", " & _
              lValueType & ", "
    
    Select Case lValueType
      Case globfuncvaltyp_STRAIGHTVALUE
        ' The new column value is a Straight Value.
        ' Get the column data type as we need to format date values for SQL.
        lngDataType = GetDataType(lColumnID)
        
'        If lngDataType = sqlVarChar Then
        If (lngDataType = sqlVarChar) Or _
          (lngDataType = sqlLongVarChar) Then
          strTemp = "'" & Replace(sValue, "'", "''") & "'"
          
        ElseIf lngDataType = sqlDate Then
          ' The give sValue will be the locale's representation of the date.
          ' Convert it to 'mm/dd/yyyy' for SQL.
          If IsDate(sValue) Then
            strTemp = "'" & Replace(Format(CDate(sValue), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
          Else
            strTemp = "null"
          End If
        Else
          ''strTemp = sValue
          'strTemp = "'" & sValue & "'"
          If datGeneral.DoesColumnUseSeparators(lColumnID) Then
            strTemp = "'" & Replace(sValue, UI.GetSystemThousandSeparator, "") & "'"
          Else
            strTemp = "'" & sValue & "'"
          End If
        End If
        
        sSQL = sSQL & "0, " & strTemp & ", 0, 0, 0)"
        
      Case globfuncvaltyp_LOOKUPTABLE
        ' The new column value is pulled from a lookup table.
        'JPD 20041209 Fault 9613
        lngDataType = GetDataType(lColumnID)
        If lngDataType = sqlDate Then
          ' The give sValue will be the locale's representation of the date.
          ' Convert it to 'mm/dd/yyyy' for SQL.
          If IsDate(sValue) Then
            sSQL = sSQL & "0, " & _
              "'" & Replace(Format(CDate(sValue), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "', " & _
              "0, " & _
              lLookupTableID & ", " & _
              lLookupColumnID & ")"
          Else
            sSQL = sSQL & "0, " & _
              "null, " & _
              "0, " & _
              lLookupTableID & ", " & _
              lLookupColumnID & ")"
          End If
        Else
          sSQL = sSQL & "0, " & _
            "'" & Replace(sValue, "'", "''") & "', " & _
            "0, " & _
            lLookupTableID & ", " & _
            lLookupColumnID & ")"
        End If

      Case globfuncvaltyp_FIELD
        ' The new column value is pulled from another field.
        sSQL = sSQL & "0, '', " & lValueID & ", 0, 0)"
        
      Case globfuncvaltyp_CALCULATION
        ' The new column value is calculated.
        sSQL = sSQL & lValueID & ", '', 0, 0, 0)"
      
      Case Else
        fOK = False
    End Select
        
    If fOK Then
      datData.ExecuteSql sSQL
    End If

End Function


Private Function GetFunctionDetails() As Recordset

    Dim sSQL As String
    Dim rsTemp As Recordset
    
    'sSQL = "exec sp_ASRGetGlobalFunction " & lFunctionID
    
    Set GetFunctionDetails = New Recordset
    
    sSQL = "SELECT ASRSysGlobalFunctions.*, " & _
                  "CONVERT(integer,ASRSysGlobalFunctions.TimeStamp) AS intTimeStamp, " & _
                  "ASRSysPickListName.Name AS PickListName, " & _
                  "ASRSysPickListName.Access AS PickListAccess, " & _
                  "ASRSysExpressions1.Name AS FilterName, " & _
                  "ASRSysExpressions1.Access AS FilterAccess, " & _
                  "ASRSysTables1.tableName AS childTableName, " & _
                  "ASRSysTables.tableName, " & _
                  "ASRSysGlobalItems.valueType, " & _
                  "ASRSysGlobalItems.columnID, " & _
                  "ASRSysGlobalItems.exprID, " & _
                  "ASRSysGlobalItems.value, " & _
                  "ASRSysGlobalItems.refColumnID, " & _
                  "ASRSysGlobalItems.LookupTableID, " & _
                  "ASRSysGlobalItems.LookupColumnID, " & _
                  "ASRSysExpressions2.name AS exprName, " & _
                  "ASRSysExpressions2.Access AS exprAccess, " & _
                  "ASRSysColumns.Size AS Size, " & _
                  "ASRSysColumns.Decimals AS Decimals, " & _
                  "ASRSysColumns.ColumnName "

    sSQL = sSQL & _
           "FROM ASRSysGlobalItems " & _
           "INNER JOIN ASRSysColumns ON ASRSysGlobalItems.columnID = ASRSysColumns.columnID " & _
           "RIGHT OUTER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID " & _
           "LEFT OUTER JOIN ASRSysExpressions ASRSysExpressions2 ON ASRSysGlobalItems.exprID = ASRSysExpressions2.exprID " & _
           "LEFT OUTER JOIN ASRSysTables ON ASRSysGlobalFunctions.tableID = ASRSysTables.tableID " & _
           "LEFT OUTER JOIN ASRSysTables ASRSysTables1 ON ASRSysGlobalFunctions.childTableID = ASRSysTables1.tableID " & _
           "LEFT OUTER JOIN ASRSysPickListName ON ASRSysGlobalFunctions.pickListID = ASRSysPickListName.pickListID " & _
           "LEFT OUTER JOIN ASRSysExpressions ASRSysExpressions1 ON ASRSysGlobalFunctions.filterID = ASRSysExpressions1.exprID " & _
           "WHERE ASRSysGlobalFunctions.FunctionID = " & CStr(mlFunctionID) & " " & _
           "ORDER BY ASRSysTables.TableName, ASRSysColumns.ColumnName"

    Set GetFunctionDetails = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

End Function


Private Sub AmendFunction(sName As String, sDesc As String, lTableID As Long, bAllRecords As Boolean, _
    lFilterID As Long, lPicklistID As Long, Optional lChildID As Long)
    
    Dim sSQL As String
    
    sSQL = "Update ASRSysGlobalFunctions Set " & _
           "Name = '" & Replace(sName, "'", "''") & "', " & _
           "Description = '" & Replace(sDesc, "'", "''") & "', " & _
           "TableID = " & lTableID & ", " & _
           "ChildTableID = " & lChildID & ", " & _
           "AllRecords = " & CLng(bAllRecords) & ", " & _
           "FilterID = " & lFilterID & ", " & _
           "PicklistID = " & lPicklistID & " " & _
           "Where FunctionID = " & mlFunctionID
    datData.ExecuteSql sSQL
        
End Sub


Private Function GetColumnName(lColumnID As Long) As String
    
    'Needed to be created as the one in datgeneral needs on tableid
    
    Dim sSQL As String
    Dim rsColumns As Recordset
    
    sSQL = "Select ColumnName From ASRSysColumns Where ColumnID = " & lColumnID
    Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    GetColumnName = rsColumns(0)
    
    rsColumns.Close
    Set rsColumns = Nothing

End Function


Private Function GetDataType(lColumnID As Long) As Long

    'Needed to be created as the one in datgeneral needs on tableid
    
    Dim sSQL As String
    Dim rsTemp As Recordset
    
    sSQL = "Select DataType From ASRSysColumns Where ColumnID = " & lColumnID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    GetDataType = rsTemp(0)
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

Private Sub CheckIfScrollBarRequired()

  With grdColumns

    If .Rows > 12 Then
      .ScrollBars = ssScrollBarsVertical
      .Columns("Value").Width = 3480
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns("Value").Width = 3710
    End If

  End With

End Sub

Private Sub cboCategory_Click()
  Changed = True
End Sub
