VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGroupMaint1 
   BackColor       =   &H80000004&
   Caption         =   "Group Permissions"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   1065
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8013
   Icon            =   "frmGroupMaint1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicVeryGreyCheckBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "frmGroupMaint1.frx":000C
      FillColor       =   &H00404040&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   930
      Picture         =   "frmGroupMaint1.frx":0156
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   4170
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   30
      Left            =   4320
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   11
      Top             =   4815
      Width           =   15
   End
   Begin VB.PictureBox PicGreyCheckbox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "frmGroupMaint1.frx":02A0
      FillColor       =   &H00404040&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1175
      Picture         =   "frmGroupMaint1.frx":03EA
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   4170
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicTickedCheckBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1410
      Picture         =   "frmGroupMaint1.frx":0534
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   4170
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicBlankCheckBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1650
      Picture         =   "frmGroupMaint1.frx":08A6
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   4170
      Visible         =   0   'False
      Width           =   240
   End
   Begin SSActiveTreeView.SSTree sstrvSystemPermissions 
      Height          =   1300
      Left            =   2500
      TabIndex        =   7
      Top             =   360
      Width           =   4000
      _ExtentX        =   7038
      _ExtentY        =   2302
      _Version        =   65538
      LabelEdit       =   1
      LineType        =   0
      Indentation     =   570
      PictureBackgroundUseMask=   0   'False
      HasFont         =   0   'False
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "imgSystemPermissions"
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   2100
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   0
      Width           =   200
   End
   Begin ComctlLib.ListView lvList 
      Height          =   1300
      Left            =   2500
      TabIndex        =   6
      Top             =   3160
      Visible         =   0   'False
      Width           =   4000
      _ExtentX        =   7038
      _ExtentY        =   2275
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "imlLargeIcons"
      SmallIcons      =   "imlSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   0
   End
   Begin SSDataWidgets_B.SSDBGrid ssgrdColumns 
      Height          =   1305
      Left            =   2505
      TabIndex        =   0
      Top             =   1755
      Visible         =   0   'False
      Width           =   6030
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   3
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
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
      Columns(0).Width=   5292
      Columns(0).Caption=   "Column"
      Columns(0).Name =   "Column"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2646
      Columns(1).Caption=   "'Read' permission"
      Columns(1).Name =   "Select"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   11
      Columns(1).FieldLen=   256
      Columns(1).Style=   2
      Columns(1).HasForeColor=   -1  'True
      Columns(2).Width=   2646
      Columns(2).Caption=   "'Edit' permission"
      Columns(2).Name =   "Update"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      TabNavigation   =   1
      _ExtentX        =   10636
      _ExtentY        =   2302
      _StockProps     =   79
      Caption         =   "Column Permissions"
      BackColor       =   -2147483633
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
   Begin SSActiveTreeView.SSTree trvConsole 
      Height          =   2000
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   3519
      _Version        =   65538
      LabelEdit       =   1
      Indentation     =   345
      OLEDropMode     =   1
      AutoSearch      =   0   'False
      HideSelection   =   0   'False
      PictureBackgroundUseMask=   0   'False
      HasFont         =   -1  'True
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "imlSmallIcons"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   1
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   4950
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8969
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSystemPermissions 
      Left            =   180
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupMaint1.frx":0BE8
            Key             =   "SYSIMG_TICK"
            Object.Tag             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupMaint1.frx":0D42
            Key             =   "SYSIMG_GREYNOTICK"
            Object.Tag             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupMaint1.frx":0E9C
            Key             =   "SYSIMG_NOTICK"
            Object.Tag             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupMaint1.frx":0FF6
            Key             =   "SYSIMG_UNKNOWN"
            Object.Tag             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupMaint1.frx":1150
            Key             =   "SYSIMG_GREYTICK"
            Object.Tag             =   "SYSTEM"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlLargeIcons 
      Left            =   1185
      Top             =   2505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":12AA
            Key             =   "CLOSEDFOLDER"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":1AFC
            Key             =   "GROUP"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":234E
            Key             =   "GROUP_ORPHAN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2BA0
            Key             =   "TABLES"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":37F2
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":4044
            Key             =   "TOPLEVELTABLE"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":4896
            Key             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":50E8
            Key             =   "LOOKUPTABLE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":593A
            Key             =   "USER_SQL"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":618C
            Key             =   "USER"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":69DE
            Key             =   "USER_ORPHAN"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":7230
            Key             =   "TABLE"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":7A82
            Key             =   "OPENFOLDER"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":82D4
            Key             =   "CHILDTABLE"
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar abGroupMaint 
      Left            =   1200
      Top             =   4440
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
      Bands           =   "frmGroupMaint1.frx":8B26
   End
   Begin VB.Label lblRightPaneCaption 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Pane Caption"
      Height          =   285
      Left            =   2500
      TabIndex        =   4
      Top             =   0
      Width           =   4000
   End
   Begin VB.Label lblLeftPaneCaption 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Pane Caption"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2000
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   180
      Top             =   2505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":296EE
            Key             =   "SYSTEM"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":29C40
            Key             =   "GROUP"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2A192
            Key             =   "GROUP_ORPHAN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2A6E4
            Key             =   "CLOSEDFOLDER"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2AC36
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2B188
            Key             =   "TOPLEVELTABLE"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2B6DA
            Key             =   "CHILDTABLE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2BC2C
            Key             =   "LOOKUPTABLE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2C17E
            Key             =   "OPENFOLDER"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2C6D0
            Key             =   "USER"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2CC22
            Key             =   "USER_ORPHAN"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGroupMaint1.frx":2D174
            Key             =   "USER_SQL"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGroupMaint1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Drag and drop variables
Private gSngSplitStartX As Single
Private gfSplitMoving As Boolean
Private gctlActiveView As Control
Private gsTreeViewNodeKey As String

Private gfMouseDown As Boolean
Private gobjCurrentNode As SSActiveTreeView.SSNode

Private giListView_ViewGroups As ListViewConstants
Private giListView_ViewCategories As ListViewConstants
Private giListView_ViewUsers As ListViewConstants
Private giListView_ViewTables As ListViewConstants
Private mblnReadOnly As Boolean

'Public gasPrintOptions() As SecurityPrintOptions

Dim mlngFirstColumnWidth As Long

Private mobjSecurity As SecurityGroups

' Array holding the User Defined functions that are needed for this update
Private mvarUDFsRequired() As String

Public Enum groupAction
  GROUPACTION_NEW = 1
  GROUPACTION_EDIT = 2
  GROUPACTION_COPY = 3
End Enum

Private Function CheckTidyPermissionOnChilds(pobjTableView As SecurityTable, _
  psCurrentGroup As String, _
  Optional pavTablesViews As Variant) As String
  ' Find all child tables of the given table/view that will have 'read' permission revoked if the
  ' given table has 'read' permission revoked.
  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim rsChildren As New ADODB.Recordset
  Dim rsParents As New ADODB.Recordset
  Dim objTempTable As SecurityTable
  Dim objTempView As SecurityTable
  Dim fCanReadParent As Boolean
  Dim sMessage As String
  Dim iLoop As Integer
  Dim sSelectedTablesViews As String
  Dim sViewNames As String
  
  sMessage = vbNullString
  
  sSelectedTablesViews = vbTab
  If Not IsMissing(pavTablesViews) Then
    For iLoop = 1 To UBound(pavTablesViews)
      Set objTempTable = pavTablesViews(iLoop)
      sSelectedTablesViews = sSelectedTablesViews & objTempTable.Name & vbTab
      Set objTempTable = Nothing
    Next iLoop
  End If
  
  ' Loop through the children of the given table.
  sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
    " FROM ASRSysRelations" & _
    " INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID" & _
    " WHERE ASRSysRelations.parentID = " & CStr(IIf(pobjTableView.ViewTableID > 0, pobjTableView.ViewTableID, pobjTableView.TableID))
  rsChildren.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsChildren
    Do While (Not .EOF)
      fCanReadParent = False

      sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
        " FROM ASRSysRelations" & _
        " INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID" & _
        " WHERE ASRSysRelations.childID = " & CStr(!TableID)
      rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
      With rsParents
        Do While (Not .EOF) And (Not fCanReadParent)
          Set objTempTable = gObjGroups(psCurrentGroup).Tables(!TableName)

          If (!TableID <> pobjTableView.TableID) And _
            (objTempTable.SelectPrivilege <> giPRIVILEGES_NONEGRANTED) And _
            (InStr(sSelectedTablesViews, vbTab & !TableName & vbTab) = 0) Then
            
            fCanReadParent = True
          Else
            For Each objTempView In gObjGroups(psCurrentGroup).Views
              If (pobjTableView.Name <> objTempView.Name) And _
                (objTempView.ViewTableID = !TableID) And _
                (objTempView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED) And _
                (InStr(sSelectedTablesViews, vbTab & objTempView.Name & vbTab) = 0) Then
                
                fCanReadParent = True
                Exit For
              End If
            Next objTempView
            Set objTempView = Nothing
          End If

          Set objTempTable = Nothing

          .MoveNext
        Loop

        .Close
      End With

      If Not fCanReadParent Then
        ' Child table does NOT have any 'readable' parents. Deny all access to it (and tidy up its own children).
        Set objTempTable = gObjGroups(psCurrentGroup).Tables(!TableName)
        With objTempTable
          If .SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
            'JPD 20050601 Fault 10137
            sMessage = sMessage & .Name & vbNewLine & CheckTidyPermissionOnChilds(objTempTable, psCurrentGroup)
          End If
        End With
        Set objTempTable = Nothing
      End If
      
      .MoveNext
    Loop

    .Close
  End With
  
  Set rsChildren = Nothing
  Set rsParents = Nothing

TidyUpAndExit:
  CheckTidyPermissionOnChilds = sMessage
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Private Sub TidyPermissionOnChilds(pobjTableView As SecurityTable, psCurrentGroup As String)
  ' Remove all permissions on child tables of the given table/view if necessary.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsChildren As New ADODB.Recordset
  Dim rsParents As New ADODB.Recordset
  Dim objTempTable As SecurityTable
  Dim objTempView As SecurityTable
  Dim fCanReadParent As Boolean
  Dim objColumn As SecurityColumn
  
  If pobjTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
    ' No 'read' permission granted. Remove all permissions on child tables IF the child table
    ' has no other parent table/view with 'read' permission granted.
    
    ' Loop through the children of the given table.
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysRelations" & _
      " INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID" & _
      " WHERE ASRSysRelations.parentID = " & CStr(IIf(pobjTableView.ViewTableID > 0, pobjTableView.ViewTableID, pobjTableView.TableID))
    rsChildren.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsChildren
      Do While (Not .EOF)
        fCanReadParent = False
        
        sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
          " FROM ASRSysRelations" & _
          " INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID" & _
          " WHERE ASRSysRelations.childID = " & CStr(!TableID)
        rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        With rsParents
          Do While (Not .EOF) And (Not fCanReadParent)
            Set objTempTable = gObjGroups(psCurrentGroup).Tables(!TableName)
        
            If objTempTable.SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
              fCanReadParent = True
            Else
              For Each objTempView In gObjGroups(psCurrentGroup).Views
                If (objTempView.ViewTableID = !TableID) And (objTempView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED) Then
                  fCanReadParent = True
                  Exit For
                End If
              Next objTempView
              Set objTempView = Nothing
            End If
          
            Set objTempTable = Nothing
          
            .MoveNext
          Loop
          
          .Close
        End With
        
        If Not fCanReadParent Then
          ' Child table does NOT have any 'readable' parents. Deny all access to it (and tidy up its own children).
          Set objTempTable = gObjGroups(psCurrentGroup).Tables(!TableName)
          With objTempTable
            If .SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
              .SelectPrivilege = giPRIVILEGES_NONEGRANTED
              .InsertPrivilege = False
              .DeletePrivilege = False
              .UpdatePrivilege = giPRIVILEGES_NONEGRANTED
              
              For Each objColumn In .Columns
                objColumn.Changed = (objColumn.SelectPrivilege <> False)
                objColumn.SelectPrivilege = False
                objColumn.UpdatePrivilege = False
              Next objColumn
              Set objColumn = Nothing
              
              .Changed = True
              
              TidyPermissionOnChilds objTempTable, psCurrentGroup
            End If
          End With
          Set objTempTable = Nothing
        End If
        
        .MoveNext
      Loop
  
      .Close
    End With
  
  End If
  
TidyUpAndExit:
  Set rsChildren = Nothing
  Set rsParents = Nothing
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Function GetSelectedUsers(aryGroups() As Variant, aryUsers() As Variant)

  'TM20010816 Fault 1808 (SUG)
  'Function added as part of above suggestion.

  'Returns an array containing all the users in the passed Group array.
  
  Dim i As Integer
  Dim sCurrentGroup As String
  Dim objUser As SecurityUser
  Dim iUserCount As Integer
  
  iUserCount = 1
  ReDim aryUsers(0)
  
  For i = 1 To UBound(aryGroups) Step 1
    sCurrentGroup = aryGroups(i)
  
    ' Fill in the users collection information if is has not yet been read.
    If Not gObjGroups(sCurrentGroup).Users_Initialised Then
      InitialiseUsersCollection gObjGroups(sCurrentGroup)
    End If

    For Each objUser In gObjGroups(sCurrentGroup).Users
      If Not objUser.DeleteUser Then
        ReDim Preserve aryUsers(iUserCount)
        aryUsers(iUserCount) = objUser.UserName
        iUserCount = iUserCount + 1
      End If
    Next
  Next i
  
TidyUpAndExit:
  Set objUser = Nothing
  Exit Function

ErrorTrap:
  MsgBox "Error retrieving selected users.", vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Function sstrvSystemPermissions_IsValid(ByRef psKey As String) As Boolean
  ' Return TRUE if the given security table exists in the collection.
  Dim fTest As Boolean
  
  On Error GoTo err_IsValid
  
  fTest = sstrvSystemPermissions.Nodes(psKey).Enabled
  sstrvSystemPermissions_IsValid = True
  
  Exit Function
  
err_IsValid:
  sstrvSystemPermissions_IsValid = False

End Function

Private Sub trvConsole_Initialise()
  ' This procedure will load the treeview with the collections being
  ' used as the nodes.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim nodGroups As SSNode
  Dim nodGroup As SSNode
  Dim nodSystemPermissions As SSNode
  Dim nodUsers As SSNode
  Dim nodTableViews As SSNode
  Dim objGroup As SecurityGroup
  
  ' Update the caption.
  lblLeftPaneCaption.Caption = " User Groups"
  
  ' Add the Groups folder root node.
  Set nodGroups = trvConsole.Nodes.Add(, , "RT", "User Groups", "CLOSEDFOLDER", , "GROUPS")
  With nodGroups
    .ExpandedImage = "OPENFOLDER"
    .Sorted = True
  End With
  
  ' Add each group to the tree with the appropriate child nodes
  For Each objGroup In gObjGroups
    
    ' Add the 'Group' node.
    Set nodGroup = trvConsole.Nodes.Add(nodGroups, tvwChild, "GP_" & objGroup.Name, objGroup.Name, "GROUP", , "GROUP")
    nodGroup.Sorted = False
    
    ' Add the Users Folder node.
    Set nodUsers = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "US_" & nodGroup.Key, "User Logins", "CLOSEDFOLDER", "OPENFOLDER", "USERS")
    With nodUsers
      .ExpandedImage = "OPENFOLDER"
      .Sorted = True
    End With
    
    ' Add the Tables/Views folder node.
    
    'MH20010208 Fault 1825
    'Set nodTableViews = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "TV_" & nodGroup.Key, "Tables / Views", "CLOSEDFOLDER", , "TABLESVIEWS")
    Set nodTableViews = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "TV_" & nodGroup.Key, "Data Permissions", "CLOSEDFOLDER", "OPENFOLDER", "TABLESVIEWS")
    With nodTableViews
      .ExpandedImage = "OPENFOLDER"
      .Sorted = True
    End With
    
    ' Add the System Permissions node.
    Set nodSystemPermissions = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "SY_" & nodGroup.Key, "System Permissions", "SYSTEM", , "SYSTEM")
    nodSystemPermissions.Sorted = True
    
    
    ' Check if the table permissions have been read for the current group.
    If objGroup.Initialised Then
      trvConsole_LoadTablesViews trvConsole.Nodes("TV_GP_" & objGroup.Name)
    End If

  Next objGroup
  Set objGroup = Nothing
  
  nodGroups.Expanded = True
  
  lvList.View = giListView_ViewGroups

TidyUpAndExit:
  ' Disassociate object variables.
  Set nodGroups = Nothing
  Set nodGroup = Nothing
  Set nodSystemPermissions = Nothing
  Set nodUsers = Nothing
  Set nodTableViews = Nothing
  Set objGroup = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub



Private Sub Form_Activate()
  SplitMove
  RefreshSecurityMenu
  
End Sub

Private Sub Form_GotFocus()
  ' Set focus on the treeview.
  trvConsole.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '# RH 26/08/99. To pass shortcut keys thru to the activebar control
  Dim fHandled As Boolean
  
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select

  fHandled = frmMain.abSecurity.OnKeyDown(KeyCode, Shift)

  If fHandled Then
    KeyCode = 0
    Shift = 0
  End If

  ' JPD20030218 Fault 5062
  If (KeyCode <> vbKeyF1) And (KeyCode <> 0) Then
    ' JDM - 19/02/01 - Fault 1869 - Error when pressing CTRL-X on treeview control
    ' For some reason the Sheridan treeview control wants to fire off it own cutn'paste functionality
    ' must trap it here not in it's own keydown event
    If Not ActiveControl Is Nothing Then
      If ActiveControl.Name = "sstrvSystemPermissions" Or ActiveControl.Name = "trvConsole" Then
        KeyCode = 0
        Shift = 0
      End If
    End If

  End If

End Sub

Private Sub Form_Load()
  Dim iWindowState As Integer
  
  mblnReadOnly = (Application.AccessMode <> accFull)
  
  ' Get the last form size and state from the registry.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  Me.Height = GetPCSetting(Me.Name, "Height", Me.Height)
  Me.Width = GetPCSetting(Me.Name, "Width", Me.Width)
  fraSplit.Left = GetPCSetting(Me.Name, "Split", IIf(Me.Width / 3 > 4000, 4000, Me.Width / 3))
    
  giListView_ViewGroups = GetPCSetting(Me.Name, "ViewGroups", lvwIcon)
  giListView_ViewCategories = GetPCSetting(Me.Name, "ViewCategories", lvwIcon)
  giListView_ViewUsers = GetPCSetting(Me.Name, "ViewUsers", lvwIcon)
  giListView_ViewTables = GetPCSetting(Me.Name, "ViewTables", lvwReport)
    
  iWindowState = GetPCSetting(Me.Name, "WindowState", vbNormal)
  Me.WindowState = IIf(iWindowState <> vbMinimized, iWindowState, vbNormal)
  
  gsTreeViewNodeKey = vbNullString

  sstrvSystemPermissions_Initialise
  
  ' Set up the tree view using the groups collection
  trvConsole_Initialise

  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME

End Sub


Private Sub Form_Resize()
  
  ' JDM - 30/07/09 - Problems with codejock and minimising/maximising screens
  If Me.ScaleHeight < 0 Or Me.ScaleWidth < 0 Then
    Exit Sub
  End If
  
  ' Check that we are not less than the minimum size for this form.
  If (Me.WindowState <> vbMinimized) And _
    (frmMain.WindowState <> vbMinimized) Then
  
'    If (Me.WindowState <> vbMaximized) Then
'      With Me
'        If .Width < 4275 Then .Width = 4275
'        If .Height < 4275 Then .Height = 4275
'      End With
'    End If
    
    ' Position the controls on the form.
    lblLeftPaneCaption.Top = 0
    lblRightPaneCaption.Top = 0
    With trvConsole
      .Top = lblLeftPaneCaption.Height

      If (Me.ScaleHeight - (.Top + sbStatus.Height)) > 0 Then
        .Height = Me.ScaleHeight - (.Top + sbStatus.Height)
      End If
      
      lvList.Top = .Top
      lvList.Height = .Height
      
      sstrvSystemPermissions.Top = .Top
      sstrvSystemPermissions.Height = .Height
    
      ssgrdColumns.Top = .Top
      ssgrdColumns.Height = .Height
    End With
    
    ' Position and size the split frame.
    With fraSplit
      .Width = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
      .Top = lblLeftPaneCaption.Top
      .Height = lblLeftPaneCaption.Height + trvConsole.Height
      If .Left + .Width > Me.ScaleWidth - 810 Then
        .Left = Me.ScaleWidth - (810 + .Width)
      End If
    End With
    
    ' Call the routine to size the tree, list view and gridcontrols.
    SplitMove
    
    ' Refresh the form display.
    Me.Refresh
  End If
  
  ' Clear the icon off the caption bar
  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If
  
  frmMain.RefreshMenu False
  
  'NHRD13032003 Fault 1783
  'Forced the various panes within this window to
  'have a background colour of white as highcolor
  'colourschemes didn't take up the transparent iconmask.
  trvConsole.BackColor = vbWhite
  lvList.BackColor = vbWhite
  sstrvSystemPermissions.BackColor = vbWhite
    
End Sub
Private Sub SplitMove()

  'MH20070109 Fault 11791
  If frmMain.WindowState = vbMinimized Then
    Exit Sub
  End If

  If Me.WindowState <> vbMinimized Then
    ' Limit the minimum size of the tree and list views.
    With fraSplit
      If .Left < 810 Then
        .Left = 810
      ElseIf .Left + .Width > Me.ScaleWidth - 2000 Then
        .Left = Me.ScaleWidth - (2000 + .Width)
      End If
    End With
    
    ' Resize the tree view.
    With trvConsole
      .Width = fraSplit.Left - .Left
      lblLeftPaneCaption.Width = .Width
    End With
    
    ' Resize the listview and grid controls.
    With lvList
      .Left = fraSplit.Left + fraSplit.Width
      .Width = Me.ScaleWidth - .Left
    
      sstrvSystemPermissions.Left = .Left
      sstrvSystemPermissions.Width = .Width
      
      ssgrdColumns.Left = .Left
      ssgrdColumns.Width = .Width
      
      lblRightPaneCaption.Left = .Left
      lblRightPaneCaption.Width = .Width
    End With
    
    ' Flag that the split move has ended.
    gfSplitMoving = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Save the form size and state to the registry.
  SavePCSetting Me.Name, "WindowState", Me.WindowState
  
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
    SavePCSetting Me.Name, "Height", Me.Height
    SavePCSetting Me.Name, "Width", Me.Width
  End If
  
  SavePCSetting Me.Name, "Split", fraSplit.Left
  
  SavePCSetting Me.Name, "ViewGroups", giListView_ViewGroups
  SavePCSetting Me.Name, "ViewCategories", giListView_ViewCategories
  SavePCSetting Me.Name, "ViewUsers", giListView_ViewUsers
  SavePCSetting Me.Name, "ViewTables", giListView_ViewTables
  
  frmMain.RefreshMenu True

End Sub

Private Sub fraSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Record the split move start position.
  gSngSplitStartX = X
  
  ' Flag that the split is being moved.
  gfSplitMoving = True

End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' If we are moving the split then move it.
  If gfSplitMoving Then
    fraSplit.Left = fraSplit.Left + (X - gSngSplitStartX)
  End If
  
End Sub


Private Sub fraSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' If the split is being moved then call the routine that resizes the
  ' tree and list views accordingly.
  If gfSplitMoving Then
    SplitMove
  End If

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub ssgrdColumns_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Cancel = mblnReadOnly
    
End Sub

Private Sub ssgrdColumns_BeforeRowColChange(Cancel As Integer)
  Cancel = mblnReadOnly
  'If mblnReadOnly And ssgrdColumns.Col <> 0 Then
  '  ssgrdColumns.Col = 0
  'End If
End Sub

Private Sub ssgrdColumns_Change()
  Dim fNewValue As Boolean
  Dim icol As Integer
  Dim iCount As Integer
  Dim sCurrentGroup As String
  Dim objTableView As SecurityTable
  Dim objParentTableView As SecurityTable
  Dim iOriginalRow As Integer
  Dim vFirstRow As Variant
  Dim iOriginalCol As Integer
  Dim iOriginalSelectPermission As ColumnPrivilegeStates
  Dim rsParents As New ADODB.Recordset
  Dim sSQL As String
  Dim fOrphan As Boolean
  Dim fAllSelectDenied As Boolean
  Dim sMessage As String
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  
  If (Left(trvConsole.SelectedItem.Key, 3) = "TB_") Then
    Set objTableView = gObjGroups(sCurrentGroup).Tables(Mid(trvConsole.SelectedItem.Key, Len("TB_TV_GP_" & sCurrentGroup) + 1))
  Else
    Set objTableView = gObjGroups(sCurrentGroup).Views(Mid(trvConsole.SelectedItem.Key, Len("VW_TV_GP_" & sCurrentGroup) + 1))
  End If
  
  ' JPD20020429 Fault 2036
  If (TypeName(objTableView) = "Nothing") Or _
    (trvConsole.SelectedItem.Key <> gsTreeViewNodeKey) Then
    trvConsole_NodeClick trvConsole.SelectedItem
    Exit Sub
  End If
  
  iOriginalSelectPermission = objTableView.SelectPrivilege
  
  'JPD 20050126 Fault 9748
  ' Do not allow permissions to be granted for child tables whose parent tables have no permissions granted.
  fOrphan = False
  If (objTableView.TableType = tabChild) _
    And (ssgrdColumns.ActiveCell.Value) Then

    fOrphan = True
    
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysRelations" & _
      " INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID" & _
      " WHERE ASRSysRelations.childID = " & CStr(objTableView.TableID)
    rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsParents
      Do While (Not .EOF)
        Set objParentTableView = gObjGroups(sCurrentGroup).Tables(!TableName)
        If objParentTableView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
          fOrphan = False
        End If
        Set objParentTableView = Nothing
        
        If Not fOrphan Then
          Exit Do
        End If
        
        For Each objParentTableView In gObjGroups(sCurrentGroup).Views
          If objParentTableView.ViewTableID = !TableID Then
            If objParentTableView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
              fOrphan = False
            End If
          End If
        Next objParentTableView
        Set objParentTableView = Nothing
        
        If Not fOrphan Then
          Exit Do
        End If
        
        .MoveNext
      Loop

      .Close
    End With
    Set rsParents = Nothing
  End If
  
  If fOrphan Then
    ssgrdColumns.ActiveCell.Value = False
    MsgBox "Permissions cannot be granted for columns in child tables whose parent tables do not have 'read' permission granted.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If
  
  ' Check if all 'read' permissions have been revoked.
  fAllSelectDenied = False
  With ssgrdColumns
    ' Check if the user has selected the all columns
    .Redraw = False

    vFirstRow = .FirstRow
    iOriginalRow = .AddItemRowIndex(.Bookmark)
    iOriginalCol = .Col

    If .AddItemRowIndex(.Bookmark) = 0 Then
      ' The 'All' row has changed.
      fNewValue = .ActiveCell.Value
      
      If .Col = 1 Then
        fAllSelectDenied = Not .ActiveCell.Value
      End If
    Else
      ' An individual column has changed.
      icol = .Col
      .MoveFirst
      .MoveNext

      ' Check all of the table/view columns..
      fAllSelectDenied = True
      For iCount = 1 To .Rows
        If .Columns(1).Value Then
          fAllSelectDenied = False
          Exit For
        End If

        .MoveNext
      Next iCount

      .Col = icol
    End If
    
    ' Ensure the original row is selected.
    ' Make the original top row the current top row.
    .FirstRow = vFirstRow
    .Bookmark = .AddItemBookmark(iOriginalRow)
    .Col = iOriginalCol

    .Redraw = True
  End With

  If fAllSelectDenied Then
    sMessage = CheckTidyPermissionOnChilds(objTableView, sCurrentGroup)
    
    If Len(sMessage) > 0 Then
      sMessage = "Removing 'read' permission will also remove all permissions to the following child tables:" & vbNewLine & vbNewLine & _
        sMessage & vbNewLine & _
        "Do you wish to continue?"
  
      If (MsgBox(sMessage, vbYesNo + vbQuestion, App.Title) <> vbYes) Then
        ssgrdColumns.ActiveCell.Value = True
        Exit Sub
      End If
    End If
  End If
  
  objTableView.Changed = True
  
  With ssgrdColumns
    ' Check if the user has selected the all columns
    .Redraw = False
    
    vFirstRow = .FirstRow
    iOriginalRow = .AddItemRowIndex(.Bookmark)
    iOriginalCol = .Col
    
    If .AddItemRowIndex(.Bookmark) = 0 Then
      ' The 'All' row has changed so modify all other rows.
      fNewValue = .ActiveCell.Value
      icol = .Col
      .MoveFirst
      .MoveNext
      
      ' Update all of the table/view columns..
      For iCount = 1 To .Rows
        If icol = 1 Then
          ' Select privilege has changed.
          .Columns(1).Value = fNewValue
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).SelectPrivilege = fNewValue
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).Changed = True
        Else
          ' Update privilege has changed.
          .Columns(2).Value = fNewValue
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).UpdatePrivilege = fNewValue
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).Changed = True
        End If
        
        .MoveNext
      Next iCount
      
      .Col = icol
    Else
      ' An individual column has changed.
      Select Case .Col
        Case 1
          ' Select privilege has changed.
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).SelectPrivilege = .ActiveCell.Value
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).Changed = True
        Case 2
          ' Update privilege has changed.
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).UpdatePrivilege = .ActiveCell.Value
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).Changed = True
      End Select
    End If
    
    ' Check the column display.
    CheckColumns

    ' Ensure the original row is selected.
    ' Make the original top row the current top row.
    .FirstRow = vFirstRow
    .Bookmark = .AddItemBookmark(iOriginalRow)
    .Col = iOriginalCol

    .Redraw = True
  End With
  
  If objTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED And _
    iOriginalSelectPermission <> giPRIVILEGES_NONEGRANTED Then
  
    TidyPermissionOnChilds objTableView, sCurrentGroup
  End If
  
  ' If the table/view is a top-level table/view then
  ' check if the select permission has changed from All/Some to None, or vice versa.
  ' If so then flag all children of this table as changed. This is done as the
  ' permitted child views on the children need to be recalculated.
    ' Do nothing if the given table.view is a child or lookup table.
  If ((objTableView.TableType <> tabChild) And (objTableView.TableType <> tabLookup)) Then
    If ((iOriginalSelectPermission <> giPRIVILEGES_NONEGRANTED) And (objTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED)) Or _
      ((iOriginalSelectPermission = giPRIVILEGES_NONEGRANTED) And (objTableView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED)) Then
      FlagChildrenChanged objTableView, sCurrentGroup
    End If
  End If
  
  gObjGroups(sCurrentGroup).RequireLogout = True  'MH20010410
  Application.Changed = True
  RefreshSecurityMenu
    
End Sub



Private Sub lvList_DblClick()
  Dim sNewNodeKey As String
  Dim sListItemKey As String
  Dim sGroupName As String
  Dim nodX As SSNode

  On Error GoTo ErrTrap

  ' If we have some listview items ...
  If lvList.ListItems.Count > 0 Then
  
    ' If the selected listview item has children ...
    If trvConsole.SelectedItem.Children > 0 Then
      
      sGroupName = WhichGroup(trvConsole.SelectedItem)
      sListItemKey = lvList.SelectedItem.Key
      sNewNodeKey = vbNullString
      
      ' Set the selected item to be the selected item in the treeview.
      Select Case trvConsole.SelectedItem.DataKey
        Case "GROUPS"
          sNewNodeKey = "GP_" & sListItemKey
          
        Case "GROUP"
          Select Case sListItemKey
            Case "USERS"
              sNewNodeKey = "US_GP_" & sGroupName
            Case "TABLESVIEWS"
              sNewNodeKey = "TV_GP_" & sGroupName
            Case "SYSTEM"
              sNewNodeKey = "SY_GP_" & sGroupName
          End Select
          
        Case "TABLESVIEWS"
          sNewNodeKey = sListItemKey
      End Select
      
      If sNewNodeKey <> vbNullString Then
        Set nodX = trvConsole.Nodes(sNewNodeKey)
        nodX.Expanded = True
        nodX.EnsureVisible
        trvConsole.SelectedItem = nodX
        UpdateRightPane
  
        ' Disassociate object variables.
        Set nodX = Nothing
      End If
    Else
      ' If the selected item does not have children then display its
      ' property page.
      EditMenu "ID_Properties"
    End If
  End If

  'TM20010910 Fault 2805
  'Must set the current active view and then refresh the menu options to enable the correct items.
  Set ActiveView = lvList
  RefreshSecurityMenu

  'NPG20090604 Fault 13696
  ' Any orphaned Windows users?
  If lvList_OrphanCount > 0 And Left(sNewNodeKey, 2) = "US" Then
    If MsgBox("SQL Server Logins do not exist for the user(s) shown in grey." & vbCrLf & _
              "Do you want to remove these users from the database?", _
              vbExclamation + vbYesNo, App.Title) = vbYes Then
      ' Select all orphans in this view...
      lvList_SelectAllOrphans
      ' Call User_Delete to get rid of them...
      User_Delete
      ' Update the view
      UpdateRightPane
      ' Update the menu
      RefreshSecurityMenu
    End If
  End If


  Exit Sub
  
ErrTrap:
  
  
End Sub

Private Sub lvList_GotFocus()
  Set ActiveView = lvList
  RefreshSecurityMenu

End Sub


Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
  ' Refresh the status bar.
  RefreshStatusBar
  
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long

  ' Pop up a menu if the right mouse button is pressed.
  If Button = vbRightButton Then
  
    ' Ensure the menu is up to date.
    RefreshSecurityMenu
    
    ' Call the activebar to display the popup menu.
    UI.GetMousePos lXMouse, lYMouse
    frmMain.abSecurity.Bands("bndEdit_Right").TrackPopup -1, -1
  Else
       
    'JDM - 19/12/01 - Fault 2975 - Delete button not enabling
    RefreshSecurityMenu
  
  End If

  RefreshStatusBar

End Sub


Private Sub ssgrdColumns_Click()
  ' Display a message if the user attempts to revoke 'read' permission on a lookup table.
  
  With ssgrdColumns
    If (.Col = 1) And (.Columns("Select").Locked) Then
      MsgBox "'Read' permission cannot be revoked for any column in a lookup table.", vbInformation + vbOKOnly, App.Title
    End If
  End With

End Sub

Private Sub ssgrdColumns_GotFocus()
  RefreshSecurityMenu
  
  With ssgrdColumns
    If .Col < 0 Then
      .Col = 1
    End If
  
   .Refresh
  End With

End Sub

Public Property Get ActiveView() As Control
  ' Return the active view control.
  Set ActiveView = gctlActiveView
  
End Property

Public Property Set ActiveView(pctlControl As Control)
  ' Set the active view control.
  Set gctlActiveView = pctlControl
  
End Property


Private Sub abGroupMaint_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub abGroupMaint_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  EditMenu Tool.Name
  
End Sub

Private Sub RefreshSecurityMenu()
  ' Refresh the toolbar on this form, and the menu bar on the MDI form.
  ' Configure the Security specific menu options.
  Dim fOK As Boolean
  Dim fEnableNew As Boolean
  Dim fEnableDelete As Boolean
  Dim fEnableProperties As Boolean
  Dim fEnableViews As Boolean
  Dim fEnableAutomaticAdd As Boolean
  Dim fEnableCopy As Boolean
  Dim fEnableMove As Boolean
  Dim fEnablePrintDetails As Boolean
  Dim fEnableChecks As Boolean
  Dim fEnableResetPassword As Boolean
  
  fOK = True
        
  With abGroupMaint.Bands("bndGroupMaint")
    
    '==================================================
    ' Configure the Edit menu tools.
    '==================================================
    ' Enable/disable the required tools.
    fEnableNew = False
    fEnableDelete = False
    fEnableProperties = False
    fEnableAutomaticAdd = False
    fEnableCopy = False
    fEnableMove = False
    fEnableChecks = False
    fEnablePrintDetails = False
    fEnableResetPassword = False
      
    If gctlActiveView Is trvConsole Then
      If Not trvConsole.SelectedItem Is Nothing Then
        Select Case trvConsole.SelectedItem.DataKey
          Case "GROUPS" 'User Groups
            fEnableNew = True
            fEnableChecks = False
            fEnablePrintDetails = True
          Case "GROUP" 'Branch Group
            fEnableCopy = True
            fEnableNew = True
            fEnableDelete = True
            fEnableProperties = Not mblnReadOnly
            fEnableChecks = False
            fEnablePrintDetails = (lvList_SelectedCount = 1)
          Case "USERS" 'User Logins
            fEnableNew = gbUserCanManageLogins Or glngSQLVersion < 9
            fEnableAutomaticAdd = gbUserCanManageLogins
            fEnablePrintDetails = True
            fEnableChecks = False
            fEnablePrintDetails = True
            
          Case "TABLESVIEWS" 'Data Permissions
            fEnableNew = False
            fEnableCopy = False
            fEnableChecks = False
            fEnablePrintDetails = True
            
          Case "TABLE"
            fEnableProperties = True
          Case "VIEW"
            fEnableProperties = True
          Case "SYSTEM" 'System Permissions
            fEnableChecks = True
            fEnablePrintDetails = True
        
        End Select
      End If
    Else
      If gctlActiveView Is lvList Then
        If Not trvConsole.SelectedItem Is Nothing Then
          Select Case trvConsole.SelectedItem.DataKey
            Case "GROUPS"
              fEnableCopy = (lvList_SelectedCount = 1)
              fEnableNew = True
              fEnableDelete = (lvList_SelectedCount > 0)
              fEnableProperties = (lvList_SelectedCount = 1)
              fEnableChecks = False
              fEnablePrintDetails = True
            Case "GROUP"
              fEnableProperties = False 'Not mblnReadOnly    'TM - Bug fix 2520
              fEnableChecks = False
              fEnablePrintDetails = True
            Case "USERS"
              fEnableNew = gbUserCanManageLogins Or glngSQLVersion < 9
              fEnableDelete = (lvList_SelectedCount > 0)
              fEnableAutomaticAdd = gbUserCanManageLogins
              ' NPG20090206 Fault 11931
              ' Enable 'Move' option only if none of the selected items are orphans.
              'fEnableMove = (lvList_SelectedCount > 0)
              fEnableMove = (lvList_SelectedOrphanCount = 0 And lvList_SelectedCount > 0)
              fEnableChecks = False
              fEnablePrintDetails = False
              If (lvList_SelectedCount = 1) Then
                fEnableResetPassword = (InStrB(1, lvList.SelectedItem.Text, "\", vbTextCompare) = 0) And gbUserCanManageLogins
              End If

            Case "TABLESVIEWS"
              fEnableProperties = (lvList_SelectedCount > 0)
              fEnableChecks = False

          End Select
        End If
      End If
    End If
    
   
    
    .Tools("ID_SecurityNew").Enabled = fEnableNew And Not mblnReadOnly
    .Tools("ID_SecurityAutomaticAdd").Enabled = fEnableAutomaticAdd And Not mblnReadOnly
    .Tools("ID_SecurityDelete").Enabled = fEnableDelete And Not mblnReadOnly
    .Tools("ID_SecurityProperties").Enabled = fEnableProperties
    .Tools("ID_SecurityCopy").Enabled = fEnableCopy And Not mblnReadOnly
    .Tools("ID_SecurityMove").Enabled = fEnableMove And Not mblnReadOnly
    .Tools("ID_SecurityResetPassword").Enabled = fEnableResetPassword And Not mblnReadOnly
   
    .Tools("ID_SecuritySave").Enabled = Application.Changed
    .Tools("ID_SecurityPrint").Enabled = fEnablePrintDetails
    
    '==================================================
    ' Configure the View menu.
    '==================================================
    fEnableViews = False
    If gctlActiveView Is lvList Then
      fEnableViews = True
    Else
      If gctlActiveView Is trvConsole Then
        If Not trvConsole.SelectedItem Is Nothing Then
          fEnableViews = (trvConsole.SelectedItem.DataKey = "GROUPS") Or _
          (trvConsole.SelectedItem.DataKey = "GROUP") Or _
          (trvConsole.SelectedItem.DataKey = "USERS") Or _
          (trvConsole.SelectedItem.DataKey = "TABLESVIEWS")
        End If
      End If
    End If
    
    'NHRD10062003 Fault 4947, 5302
    'NHRD12012004 Fault 7137 added "And (Not mblnReadOnly)" to ensure the
    'check buttons are diabled when use has Read Access
    .Tools("ID_CheckAll").Enabled = fEnableChecks And (Not mblnReadOnly)
    .Tools("ID_UnCheckAll").Enabled = fEnableChecks And (Not mblnReadOnly)
    
    ' Refreshmenu some edit menu tools to reflect the button bar
    frmMain.abSecurity.Bands("bndEdit_Right").Tools("ID_CheckAll").Enabled = fEnableChecks And (Not mblnReadOnly)
    frmMain.abSecurity.Bands("bndEdit_Right").Tools("ID_UnCheckAll").Enabled = fEnableChecks And (Not mblnReadOnly)
    frmMain.abSecurity.Bands("bndEdit_Left").Tools("ID_SecurityPrint").Enabled = fEnablePrintDetails
      
    .Tools("ID_LargeIcons").Enabled = fEnableViews
    .Tools("ID_SmallIcons").Enabled = fEnableViews
    .Tools("ID_List").Enabled = fEnableViews
    .Tools("ID_Details").Enabled = fEnableViews And _
      (trvConsole.SelectedItem.DataKey = "TABLESVIEWS")
      
    If (lvList.View = lvwReport) And _
      (Not .Tools("ID_Details").Enabled) Then
      lvList.View = lvwList
    End If
      
    If fEnableViews Then
    
      .Tools("ID_LargeIcons").Checked = False
      .Tools("ID_SmallIcons").Checked = False
      .Tools("ID_List").Checked = False
      .Tools("ID_Details").Checked = False
    
      Select Case lvList.View
        Case lvwIcon
          .Tools("ID_LargeIcons").Checked = True
        Case lvwSmallIcon
          .Tools("ID_SmallIcons").Checked = True
        Case lvwList
          .Tools("ID_List").Checked = True
        Case lvwReport
          .Tools("ID_Details").Checked = True
      End Select
    End If
  End With
  
  frmMain.RefreshMenu False
  
End Sub

Private Sub ssgrdColumns_KeyDown(KeyCode As Integer, Shift As Integer)
  'JPD 20050606 Fault 10149
  ssgrdColumns.SetFocus

End Sub

Private Sub sstrvSystemPermissions_GotFocus()
  RefreshSecurityMenu

End Sub

Private Sub sstrvSystemPermissions_KeyPress(KeyAscii As Integer)
  ' Toggle the selected node's status when the SPACE key is pressed.
  If KeyAscii = vbKeySpace Then
    If Not gobjCurrentNode Is Nothing Then
      ToggleSystemPermission gobjCurrentNode
      Set sstrvSystemPermissions.SelectedItem = gobjCurrentNode
    End If
  End If
  
End Sub


Private Sub sstrvSystemPermissions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Set the flag that shows that the mouse is down.
  gfMouseDown = True

End Sub


Private Sub sstrvSystemPermissions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Set the flag that shows that the mouse is NOT down.
  'gfMouseDown = False

  Dim lXMouse As Long
  Dim lYMouse As Long

  ' Pop up a menu if the right mouse button is pressed.
  If Button = vbRightButton Then
  
    ' Ensure the menu is up to date.
    RefreshSecurityMenu
    
    ' Call the activebar to display the popup menu.
    UI.GetMousePos lXMouse, lYMouse
    frmMain.abSecurity.Bands("bndEdit_Right").TrackPopup -1, -1
  Else
       
    'JDM - 19/12/01 - Fault 2975 - Delete button not enabling
    RefreshSecurityMenu
  
  End If

  RefreshStatusBar

End Sub


Private Sub sstrvSystemPermissions_NodeClick(Node As SSActiveTreeView.SSNode)
  ' Toggle the clicked node's status when clicked using the mouse (not the keys).
  If gfMouseDown Then
    ToggleSystemPermission Node
  End If

  Set gobjCurrentNode = Node

End Sub


Private Sub trvConsole_Collapse(Node As SSActiveTreeView.SSNode)
  ' Update the display for the selected node.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Ensure the specified node is selected.
  Node.Selected = True
  
  ' If we are changing node then clear any selections from the listview.
  If Node.Key <> gsTreeViewNodeKey Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.Key

    ' Update the menu.
    RefreshSecurityMenu
  End If
  
  Select Case Left(Node.Key, 2)
    Case "RT"
      lvList.View = giListView_ViewGroups
    Case "GP"
      lvList.View = giListView_ViewCategories
    Case "US"
      lvList.View = giListView_ViewUsers
    Case "TV"
      lvList.View = giListView_ViewTables
  End Select
  
TidyUpAndExit:
  If Not fOK Then
    MsgBox "Error refreshing display.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub trvConsole_DblClick()
  
  With trvConsole
    If Not .SelectedItem Is Nothing Then
      If .SelectedItem.DataKey = "VIEW" Or .SelectedItem.DataKey = "TABLE" Then
        TableView_Properties
      End If
    End If
  End With
  
End Sub


Private Sub trvConsole_Expand(Node As SSActiveTreeView.SSNode)
  ' Update the display for the selected node.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sCurrentGroup As String

  fOK = True
  
  ' Ensure the specified node is selected.
  Node.Selected = True
      
  ' Initialise the group's collections if required.
  If Left(Node.Key, 2) = "GP" Then
    sCurrentGroup = WhichGroup(Node)
    
    InitialiseGroupCollections sCurrentGroup, True
  End If

  ' If we are changing node then clear any selections from the listview.
  If Node.Key <> gsTreeViewNodeKey Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.Key

    ' Update the menu.
    RefreshSecurityMenu
  End If
  
  Select Case Left(Node.Key, 2)
    Case "RT"
      lvList.View = giListView_ViewGroups
    Case "GP"
      lvList.View = giListView_ViewCategories
    Case "US"
      lvList.View = giListView_ViewUsers
    Case "TV"
      lvList.View = giListView_ViewTables
  End Select
  
TidyUpAndExit:
  If Not fOK Then
    MsgBox "Error refreshing display.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Private Sub trvConsole_GotFocus()
  Set ActiveView = trvConsole
  RefreshSecurityMenu

End Sub

Private Sub trvConsole_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long
  Dim nodX As SSNode

  RefreshSecurityMenu
  
  ' Pop up a menu if the right mouse button is pressed.
  If Button = vbRightButton Then
  
    ' Check that we are over a node
    Set nodX = trvConsole.HitTest(X, Y)
    
    If Not nodX Is Nothing Then
      Set nodX = Nothing
      
      ' Call the activebar to display the popup menu.
      UI.GetMousePos lXMouse, lYMouse
      frmMain.abSecurity.Bands("bndEdit_Left").TrackPopup -1, -1
     
    End If
  End If
 
End Sub

Private Sub trvConsole_NodeClick(Node As SSActiveTreeView.SSNode)
  ' Update the display for the selected node.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
   
  Select Case Left(Node.Key, 2)
    Case "RT"
      lvList.View = giListView_ViewGroups
    Case "GP"
      lvList.View = giListView_ViewCategories
    Case "US"
      lvList.View = giListView_ViewUsers
    Case "TV"
      lvList.View = giListView_ViewTables
  End Select

  ' If we are changing node then clear any selections from the listview.
  If Node.Key <> gsTreeViewNodeKey Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.Key

    ' Update the menu.
    RefreshSecurityMenu
  End If

  ' NPG20090206 Fault 11931
  ' Any orphaned Windows users?
  If lvList_OrphanCount > 0 And Left(Node.Key, 2) = "US" Then
    If MsgBox("SQL Server Logins do not exist for the user(s) shown in grey." & vbCrLf & _
              "Do you want to remove these users from the database?", _
              vbExclamation + vbYesNo, App.Title) = vbYes Then
      ' Select all orphans in this view...
      lvList_SelectAllOrphans
      ' Call User_Delete to get rid of them...
      User_Delete
      ' Update the view
      UpdateRightPane
      ' Update the menu
      RefreshSecurityMenu
    End If
  End If

  ResetHelpContextID

'  'Reset the HelpContextID so that the correct help fpage is called up when
'  'F1 is pressed
'  Select Case trvConsole.SelectedItem.DataKey
'     Case "SYSTEM"
'        Me.HelpContextID = 8049
'
'      Case "GROUPS"
'        Me.HelpContextID = 8050
'
'      Case "GROUP"
'        Me.HelpContextID = 8051
'
'      Case "USERS"
'        Me.HelpContextID = 8052
'
'      Case "TABLESVIEWS"
'        Me.HelpContextID = 8053
'
'  End Select

TidyUpAndExit:
  If Not fOK Then
    MsgBox "Error refreshing display.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Private Sub sstrvSystemPermissions_Refresh()
  ' Load the System Permissions treeview with the selected user group's permissions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim sCurrentGroup As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim objSystemPermission As clsSystemPermission

  fOK = True
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  
  ' Update the right pane caption.
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group - System Permissions"
  
  ' Clear the permissions.
  For Each objNode In sstrvSystemPermissions.Nodes
    If Left$(objNode.Key, 1) = "P" Then
      'objNode.Image = "SYSIMG_NOTICK"
      objNode.Image = IIf(mblnReadOnly, "SYSIMG_GREYNOTICK", "SYSIMG_NOTICK")
    End If
  Next objNode
  
  ' Populate the grid with the permission information.
  For iLoop = 1 To gObjGroups(sCurrentGroup).SystemPermissions.Count
    Set objSystemPermission = gObjGroups(sCurrentGroup).SystemPermissions.Item(iLoop)
    If objSystemPermission.Allowed Then
      'JPD 20040114 Fault 7925
      If sstrvSystemPermissions_IsValid("P_" & objSystemPermission.CategoryKey & "_" & objSystemPermission.ItemKey) Then
        'sstrvSystemPermissions.Nodes("P_" & objSystemPermission.CategoryKey & "_" & objSystemPermission.ItemKey).Image = "SYSIMG_TICK"
        sstrvSystemPermissions.Nodes("P_" & objSystemPermission.CategoryKey & "_" & objSystemPermission.ItemKey).Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
      End If
    End If
    Set objSystemPermission = Nothing
  Next iLoop

  RelateSystemPermissions
  
  RefreshStatusBar

TidyUpAndExit:
  Set objNode = Nothing
  Set objSystemPermission = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Private Sub ssgrdColumns_RefreshTableColumns()
  ' Load the column grid with the column info for the currently selected table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fAllColumnSelect As Boolean
  Dim fAllColumnUpdate As Boolean
  Dim sTableName As String
  Dim sCurrentGroup As String
  Dim objColumn As SecurityColumn
  
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  sTableName = Mid(trvConsole.SelectedItem.Key, Len(trvConsole.SelectedItem.Parent.Key) + 4)
  
  ' Update the right pane caption.
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group - '" & sTableName & "' table columns"
  
  ' Lock the 'read' permission column if the columns are for a lookup table.
  If gObjGroups(sCurrentGroup).Tables(sTableName).TableType = tabLookup Then
    With ssgrdColumns.Columns("Select")
      .Locked = True
      .ForeColor = vbButtonFace
    End With
  Else
    With ssgrdColumns.Columns("Select")
      .Locked = False
      .ForeColor = vbBlack
    End With
  End If

  fAllColumnUpdate = True
  fAllColumnSelect = True
  For Each objColumn In gObjGroups(sCurrentGroup).Tables(sTableName).Columns
    ssgrdColumns.AddItem objColumn.Name & _
      vbTab & objColumn.SelectPrivilege & _
      vbTab & objColumn.UpdatePrivilege
    
    ' Check for all columns being updateable and selectable
    If fAllColumnSelect And Not objColumn.SelectPrivilege Then
      fAllColumnSelect = False
    End If
    If fAllColumnUpdate And Not objColumn.UpdatePrivilege Then
      fAllColumnUpdate = False
    End If
  Next
  
  ' Add an 'All' item to the top of the grid.
  ssgrdColumns.AddItem "(All Columns)" & _
    vbTab & fAllColumnSelect & _
    vbTab & fAllColumnUpdate, 0
  
    'NHRD Fault 4318 Was trying to use these commands to force
    ' the selected item to be highlighted for
    ' Me.trvConsole.Nodes.Item(gsTreeViewNodeKey).Selected = True
    ' Me.trvConsole.Nodes.Item(gsTreeViewNodeKey).BackColor = vbHighlight
    
TidyUpAndExit:
  Set objColumn = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Private Sub CheckColumns()
  Dim fAllChecked As Boolean
  Dim fSelectColumnChanged As Boolean
  Dim fOneChecked As Boolean
  Dim iCount As Integer
  Dim iColumn As Integer
  Dim iCurrentRow As Integer
  Dim iInvalidCount As Integer
  Dim sMessage As String
  Dim sCurrentGroup As String
  Dim objTableView As SecurityTable
  
  fSelectColumnChanged = False

  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
 
  If (Left(trvConsole.SelectedItem.Key, 3) = "TB_") Then
    Set objTableView = gObjGroups(sCurrentGroup).Tables(Mid(trvConsole.SelectedItem.Key, Len("TB_TV_GP_" & sCurrentGroup) + 1))
  Else
    Set objTableView = gObjGroups(sCurrentGroup).Views(Mid(trvConsole.SelectedItem.Key, Len("VW_TV_GP_" & sCurrentGroup) + 1))
  End If
  
  ' Do both the update and select columns
  ' But do the update first as this needs to set the select column
  ' if there is an update on one and no select
  With ssgrdColumns
    For iColumn = 2 To 1 Step -1
      fOneChecked = False
      fAllChecked = True
      
     ' Decide if all the columns are checked
     .MoveFirst
     .MoveNext  ' Move to first row as row 0 is the 'All columns' row.

     For iCount = 1 To .Rows
       If .Columns(iColumn).Value Then
         fOneChecked = True
       Else
         fAllChecked = False
       End If

       ' Ensure that the 'select' column is selected if the 'update' column is selected.
       If (iColumn = 2) And _
        (.Columns(iColumn).Value) And _
        (Not .Columns(1).Value) Then
          fSelectColumnChanged = True

          .Columns(1).Value = True
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).SelectPrivilege = True
          objTableView.Columns(.Columns(0).CellValue(.Bookmark)).Changed = True
       End If

       .MoveNext
     Next iCount

     ' Set the 'All columns' as appropriate.
     .MoveFirst
     .Columns(iColumn).Value = fAllChecked

     ' Update the table/view object as appropriate.
     If fOneChecked Then
      ' Ensure the table/view has the privilege granted.
      If iColumn = 1 Then
        ' Grant the 'select' permission
        objTableView.SelectPrivilege = IIf(fAllChecked, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_SOMEGRANTED)
      ElseIf iColumn = 2 Then
        ' Grant the 'update' permission
        objTableView.UpdatePrivilege = IIf(fAllChecked, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_SOMEGRANTED)
      End If
    Else
      ' Ensure the table/view has the privilege revoked.
      If iColumn = 1 Then
        ' Revoke the 'select' permission
         objTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED
       ElseIf iColumn = 2 Then
        ' Revoke the 'update' permission
         objTableView.UpdatePrivilege = giPRIVILEGES_NONEGRANTED
       End If
     End If
    Next iColumn
  End With

  If fSelectColumnChanged Then
    MsgBox "'Read' permission is automatically granted for all columns which have 'Edit' permission granted.", vbInformation + vbOKOnly, App.Title
  End If
    
  ' Validate the table.view permissions.
  iInvalidCount = 0
  sMessage = vbNullString
  
  With objTableView
    If .SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
      ' If the 'UPDATE' permission is granted, but the 'SELECT' privilege is not,
      ' then inform the user.
      If .UpdatePrivilege <> giPRIVILEGES_NONEGRANTED Then
        .UpdatePrivilege = giPRIVILEGES_NONEGRANTED
        sMessage = "'Edit'"
        iInvalidCount = iInvalidCount + 1
      End If
    
      ' If the 'DELETE' permission is granted, but the 'SELECT' privilege is not,
      ' then inform the user.
      If .DeletePrivilege Then
        .DeletePrivilege = False
        sMessage = sMessage & IIf(iInvalidCount > 0, ", ", vbNullString) & "'Delete'"
        iInvalidCount = iInvalidCount + 1
      End If
    
      ' If the 'INSERT' permission is granted, but the 'SELECT' privilege is not,
      ' then inform the user.
      If .InsertPrivilege Then
        .InsertPrivilege = False
        sMessage = sMessage & IIf(iInvalidCount > 0, ", ", vbNullString) & "'New'"
        iInvalidCount = iInvalidCount + 1
      End If
    
      If iInvalidCount > 0 Then
        sMessage = "Revoking 'Read' permission automatically revokes " & sMessage & _
          " permission" & IIf(iInvalidCount > 1, "s.", ".")
      End If
    End If
    
    If .UpdatePrivilege = giPRIVILEGES_NONEGRANTED Then
      ' If the 'INSERT' permission is granted, but the 'UPDATE' privilege is not,
      ' then inform the user.
      If .InsertPrivilege Then
        .InsertPrivilege = False
        sMessage = sMessage & IIf(Len(sMessage) > 0, vbNewLine, vbNullString) & _
          "Revoking 'Edit' permission automatically revokes 'New' permission."
        iInvalidCount = iInvalidCount + 1
      End If
    End If
  End With
  
  If iInvalidCount > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.Title
  End If
    
  ' Check that all permissions are granted if the user group has permission to
  ' run the System Manager or Security Manager.
  With ssgrdColumns
    .MoveFirst
    If (.Columns(1).Value <> True) Or _
      (.Columns(2).Value <> True) Then

      sMessage = vbNullString
      If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
        sMessage = "System Manager"
      End If

      If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
        If Len(sMessage) > 0 Then
          sMessage = "System & Security Managers"
        Else
          sMessage = "Security Manager"
        End If
      End If

      If Len(sMessage) > 0 Then
        'JPD 20050208 Fault 9790
        MsgBox "Permission to run the " & sMessage & " requires the " & _
          "user group to have full access to all tables and views." & vbNewLine & vbNewLine & _
          "Permission to run the " & sMessage & " will be revoked automatically.", _
          vbOKOnly + vbExclamation, App.ProductName

        gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed = False
        gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed = False
      End If
    End If
  End With
  
End Sub


Public Function User_Delete()
  Dim fDeleteOK As Boolean
  Dim fConfirmed As Boolean
  Dim fDeleteAll As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim iDeletedCount As Integer
  Dim sSQL As String
  Dim sObjectType As String
  Dim sCurrentUserName As String
  Dim sCurrentGroupName As String
  Dim asSelectedUsers() As Variant
  Dim objGroup As SecurityGroup
  Dim rsInfo As New ADODB.Recordset
  
  Dim sMessage As String
  
  ReDim asSelectedUsers(0)
  fDeleteAll = False
  iDeletedCount = 0
  sMessage = vbNullString
  
  ' What is the current group
  sCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  
  ' If we have more than one selection then question the multi-deletion.
  If (lvList_SelectedCount > 1) Then
    ' Read the names of the users to be deleted from the listview into an array.
    For iLoop = 1 To lvList.ListItems.Count
      If lvList.ListItems(iLoop).Selected = True Then
        iNextIndex = UBound(asSelectedUsers) + 1
        ReDim Preserve asSelectedUsers(iNextIndex)
        asSelectedUsers(iNextIndex) = lvList.ListItems(iLoop).Text
      End If
    Next iLoop
  Else
    ' Read the name of the group to be deleted from the treeview into an array.
    ReDim asSelectedUsers(1)
    asSelectedUsers(1) = lvList.SelectedItem.Text
  End If

'********************************************************************************
'TM20010816 Fault 1808 (SUG)

  If lvList_SelectedCount > 1 Then
'    sMessage = "Before the deletion can be actioned, the utility ownership for " & _
'                "all the selected users must be transferred to available users. " & _
'                vbNewLine & "Are you sure you want to delete these " & _
'                Trim(Str(lvList_SelectedCount)) & " users?"
    sMessage = "Are you sure you want to delete these " & _
                Trim(Str(lvList_SelectedCount)) & " users?"
  Else
    sMessage = "Are you sure you want to delete this user?"
  End If

  If MsgBox(sMessage, vbQuestion + vbYesNo, App.ProductName) = vbYes Then
    fDeleteAll = True
    fConfirmed = True
  Else
    Exit Function
  End If

  Dim frmTransfer As New frmTransferOwnership

  'NHRD15082003 Fault 4369 re-do for 30/07/2003 SJH
  With frmTransfer
    If .Initialise(asSelectedUsers) Then
      .Show vbModal
      fConfirmed = Not .Cancelled
    Else
      
'      If (Me.lvList.ListItems.Count < 2) Then
'        'NHRD17092003 Fault 6963
'        fConfirmed = fDeleteAll
'      Else
        fConfirmed = Not .Cancelled ' True
'      End If
    End If
  End With
  Unload frmTransfer
  Set frmTransfer = Nothing

'********************************************************************************

  If fConfirmed Then
    ' Delete all of the users in the array..
    For iLoop = 1 To UBound(asSelectedUsers)
      
      sCurrentUserName = asSelectedUsers(iLoop)
    
      If Not fDeleteAll Then
        
        ' Prompt the user to confirm the deletion.
        If MsgBox("Are you sure you want to delete the user '" & _
          sCurrentUserName & "' ?", vbYesNo + vbDefaultButton2 + _
          vbQuestion, Application.Name) = vbYes Then
                            
          fConfirmed = True
        Else
          fConfirmed = False
        End If
      End If
      
      If fConfirmed Then
      
        ' Check that the user being deleted can be deleted.
        ' ie. is not the logged in user.
        fDeleteOK = Not (UCase$(Trim(sCurrentUserName)) = UCase$(Trim(gsUserName)))
        If Not fDeleteOK Then
          MsgBox "Unable to delete the user '" & sCurrentUserName & "'." & vbNewLine & _
            "You are currently logged in as this user.", _
            vbInformation + vbOKOnly, App.Title
        End If
  
        ' Check that the user being deleted can be deleted.
        ' ie. does not own any database objects.
        If fDeleteOK Then
          sSQL = "SELECT sysobjects.name, sysobjects.xtype" & _
            " FROM sysobjects" & _
            " INNER JOIN sysusers ON sysobjects.uid = sysusers.uid" & _
            " WHERE sysusers.name = '" & Replace(sCurrentUserName, "'", "''") & "'"
            
          'TM20020429 Fault 3690 - No longer need to check for ASRSysTemp* tables
          sSQL = sSQL & "   AND NOT (sysobjects.xtype = 'U' " & _
                        "             AND sysobjects.name LIKE 'ASRSysTemp%') "
        
          rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
          fDeleteOK = (rsInfo.EOF And rsInfo.BOF)
          
          If Not fDeleteOK Then
            Select Case Trim(rsInfo!xtype)
              Case "C"
                sObjectType = "check constraint"
              Case "D"
                sObjectType = "default constraint"
              Case "F"
                sObjectType = "foreign key constraint"
              Case "L"
                sObjectType = "log"
              Case "P"
                sObjectType = "stored procedure"
              Case "PK"
                sObjectType = "primary key constraint"
              Case "RF"
                sObjectType = "replication filter stored procedure"
              Case "S"
                sObjectType = "system table"
              Case "TR"
                sObjectType = "trigger"
              Case "U"
                sObjectType = "user table"
              Case "UQ"
                sObjectType = "unique constraint"
              Case "V"
                sObjectType = "view"
              Case "X"
                sObjectType = "extended stored procedure"
              Case Else
                sObjectType = "object"
            End Select
            
            MsgBox "Unable to delete the user '" & sCurrentUserName & "' as it is the" & vbNewLine & _
              "owner of the '" & rsInfo!Name & "' " & sObjectType & " in the database.", _
              vbInformation + vbOKOnly, App.Title
          End If
          
          rsInfo.Close
          Set rsInfo = Nothing
        End If
  
        If fDeleteOK Then
          Set objGroup = gObjGroups(sCurrentGroupName)
  
          ' See if the user exists in SQL server
          If objGroup.Users(sCurrentUserName).NewUser = True Then
            ' The user has not yet been added to sql server so just remove the item from
            ' the security users collection for this group.
            objGroup.Users.Remove (sCurrentUserName)
          Else
            ' The user does exist in sql server so now see if the user has been moved from another
            ' group
            If objGroup.Users(sCurrentUserName).MovedUserFrom <> vbNullString Then
              ' The user has been moved to this group so set the original groups properties
              ' to indicate that it is now deleted and remove it from this group.
              ' Original group
              'gObjGroups(objGroup.Users(sCurrentUserName).MovedUserFrom).Users(sCurrentUserName).DeleteUser = True
              'gObjGroups(objGroup.Users(sCurrentUserName).MovedUserFrom).Users(sCurrentUserName).MovedUserTo = vbnullstring
              
              With gObjGroups(objGroup.Users(sCurrentUserName).MovedUserFrom).Users(sCurrentUserName)
                .DeleteUser = True
                '.RequireLogout = True   'MH20010410
                .MovedUserTo = vbNullString
              End With
              
              ' Current group
              objGroup.Users.Remove (sCurrentUserName)
            Else
              ' The user has not been moved from another group so set its properties to be deleted.
              objGroup.Users(sCurrentUserName).Changed = True
              objGroup.Users(sCurrentUserName).DeleteUser = True
              'objGroup.Users(sCurrentUserName).RequireLogout = True   'MH20010410
            End If
          End If
          
          Set objGroup = Nothing
          iDeletedCount = iDeletedCount + 1
        
        End If
      End If
    Next iLoop
  End If
  
  If iDeletedCount > 0 Then
    ' Enable the apply button
    Application.Changed = True
  
    UpdateRightPane
  End If
  
End Function

Private Function User_New()
  ' Add a new user to the current group.
  Dim fUserFound As Boolean
  Dim sPassword As String
  Dim sCurrentGroupName As String
  Dim fValuesOK As Boolean
  Dim sSQL As String
  Dim sUserLogin As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim rsRecords As New ADODB.Recordset
  Dim fExit As Boolean
  Dim frmEdit As frmNewUser
  Dim frmPwdEdit As frmPasswordEntry
  Dim iCount As Integer
  Dim astrUserLogins() As String
  Dim astrLoginType() As String
  Dim mbUsersAdded As Boolean
  Dim iLoginType As SecurityMgr.LoginType
  Dim strFoundInGroup As String
  Dim bForcePasswordChange As Boolean
  Dim bCheckPolicy As Boolean
  
  ' Get the name of the currently selected group.
  sCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  
  ' Initialise the user group if necessary.
  If Not gObjGroups(sCurrentGroupName).Users_Initialised Then
    InitialiseUsersCollection gObjGroups(sCurrentGroupName)
  End If
  
  fExit = False
  mbUsersAdded = False
    
  Set frmEdit = New frmNewUser
  With frmEdit
    
    Do While Not fExit
      ' Call the new user form.
      .Show vbModal
    
      ' Check whether the user select OK or cancel.
      If .Cancelled Then
        fExit = True
      Else
        fValuesOK = True
        
        astrUserLogins = Split(.UserLogin, ";")
        astrLoginType = Split(.UserLoginTypes, ";")
        
        For iCount = LBound(astrUserLogins) To UBound(astrUserLogins)

          sUserLogin = astrUserLogins(iCount)
          iLoginType = Val(astrLoginType(iCount))
          
          ' Check that the user name is acceptable.
          ' Check that a user name has been entered.
          If sUserLogin = vbNullString Then
            MsgBox "Please enter a user login.", _
              vbInformation + vbOKOnly, App.Title
            fValuesOK = False
          End If
      
          ' Check if the user login is a reserved login.
          If fValuesOK And (UCase$(sUserLogin) = "SA") Then
            MsgBox "'" & sUserLogin & "' is a reserved system login." & vbNewLine & "Please enter another user login.", _
              vbInformation + vbOKOnly, App.Title
            fValuesOK = False
          End If
      
          ' Check if the user login is an SQL keyword.
          If fValuesOK And Database.IsKeyword(sUserLogin) And iLoginType = iUSERTYPE_SQLLOGIN Then
            MsgBox "'" & sUserLogin & "' is a reserved word." & vbNewLine & "Please enter another user login.", _
              vbInformation + vbOKOnly, App.Title
            fValuesOK = False
          End If
      
          ' Check for it being a group or user name already in use.
          If fValuesOK And IsUserNameInUse(sUserLogin, gObjGroups, strFoundInGroup) Then
          
            If LenB(strFoundInGroup) > 0 Then
              MsgBox "'" & sUserLogin & "' is already in use as a user in the group " & strFoundInGroup & "." & vbNewLine & _
                "Please enter another user name.", _
                vbInformation + vbOKOnly, App.Title
            Else
              MsgBox "'" & sUserLogin & "' is already in use as a group name in the database." & vbNewLine & _
                "Please enter another user name.", _
                vbInformation + vbOKOnly, App.Title
            
            End If
            fValuesOK = False
          End If
      
          If fValuesOK And Len(sUserLogin) > 50 Then
            MsgBox "'" & sUserLogin & "' is longer than 50 characters." & vbNewLine & "Please enter another user login.", _
              vbInformation + vbOKOnly, App.Title
            fValuesOK = False
          End If
      
          If fValuesOK Then
            If iLoginType = iUSERTYPE_SQLLOGIN Then
              ' Prompt the user for a password if we are creating a new login.
              sPassword = vbNullString
              'TM20011113 Fault 3125 - retrieve loginname not just name.
              sSQL = "SELECT loginname " & _
                "FROM master.dbo.syslogins " & _
                "WHERE loginname='" & Replace(sUserLogin, "'", "''") & "'"
              rsRecords.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
              
              If rsRecords.EOF And rsRecords.BOF Then
    
                Set frmPwdEdit = New frmPasswordEntry
    
                With frmPwdEdit
                  .UserName = sUserLogin
                
                  ' Call the new user form.
                  .Show vbModal
    
                  ' Check whether the user select OK or cancel.
                  If Not .Cancelled Then
                    bForcePasswordChange = .ForcePasswordChange
                    sPassword = .Password
                    bCheckPolicy = .CheckPolicy
                  Else
                    sPassword = vbNullString
                    fValuesOK = False
                  End If
                End With
    
                ' Unload the new user form
                Set frmPwdEdit = Nothing
    
'                If sPassword = vbnullstring Then
'                  fValuesOK = False
'                End If
              End If
              rsRecords.Close
              
            Else
              ' Check that the NT account actually exists
              If Not CheckNTAccountExist(sUserLogin) Then
                MsgBox "'" & sUserLogin & "' is not a recognised Windows login.", _
                  vbInformation + vbOKOnly, App.Title
                fValuesOK = False
              End If
            End If

          End If
          
          If fValuesOK Then
            ' Check if the user already exists in another group but is marked as deleted.
            fUserFound = False
            
            For Each objGroup In gObjGroups
              If Not objGroup.Users_Initialised Then
                InitialiseUsersCollection objGroup
              End If
              
              For Each objUser In objGroup.Users
                If UCase$(sUserLogin) = UCase$(objUser.Login) Then
                  ' Mark the user as undeleted and move it to the new group.
                  objUser.DeleteUser = False
                  MoveUserToNewGroup sUserLogin, sCurrentGroupName, objGroup.Name
                  fUserFound = True
                  Exit For
                End If
              Next objUser
              Set objUser = Nothing
              
              If fUserFound Then
                Exit For
              End If
            Next objGroup
            Set objGroup = Nothing
            
            ' Add the user to the group.
            If Not fUserFound Then
              gObjGroups(sCurrentGroupName).Users.Add sUserLogin, True, False, True, _
                vbNullString, vbNullString, sUserLogin, sPassword, , , bForcePasswordChange, iLoginType, bCheckPolicy
                
              'gObjGroups(sCurrentGroupName).Users(sUserLogin).LoginType = iLoginType
            End If
            
            ' Enable the apply button
            Application.Changed = True
            fExit = True
            mbUsersAdded = True

          End If
                    
        Next iCount
      End If
    Loop
  End With
  
  If mbUsersAdded Then
    UpdateRightPane
  End If
  
  ' Unload the new user form
  Set frmEdit = Nothing
  Set rsRecords = Nothing
    
End Function

Public Function Group_Delete()
  
  ' Delete selected groups.
  Dim fDeleteOK As Boolean
  Dim fConfirmed As Boolean
  Dim fDeleteAll As Boolean
  Dim fMovedUsers As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim iDeletedCount As Integer
  Dim sSQL As String
  Dim sCurrentGroupName As String
  Dim asSelectedGroups() As Variant
  Dim rsRoles As New ADODB.Recordset
  Dim objUser As SecurityUser
  Dim sBatchJobs As String
  
  'TM20010816 Fault 1808 (SUG)
  'Need an array of all the selected users eg. all the users in all the selected
  'groups. The array can then be passed to the frmTransfer object.
  Dim asSelectedUsers() As Variant
  Dim sMessage As String
  
  ReDim asSelectedGroups(0)
  fDeleteAll = False
  fMovedUsers = False
  iDeletedCount = 0
  sMessage = vbNullString
  
  ' If we have more than one selection then question the multi-deletion.
  If (ActiveView Is lvList) Then
    If (lvList_SelectedCount > 1) Then
  
      ' Read the names of the groups to be deleted from the listview into an array.
      For iLoop = 1 To lvList.ListItems.Count
        If lvList.ListItems(iLoop).Selected = True Then
          iNextIndex = UBound(asSelectedGroups) + 1
          ReDim Preserve asSelectedGroups(iNextIndex)
          asSelectedGroups(iNextIndex) = lvList.ListItems(iLoop).Text
        End If
      Next iLoop
    Else
      ' Read the name of the group to be deleted from the treeview into an array.
      ReDim asSelectedGroups(1)
      asSelectedGroups(1) = lvList.SelectedItem.Text
    End If
  
  Else
    ' Read the name of the group to be deleted from the treeview into an array.
    ReDim asSelectedGroups(1)
    asSelectedGroups(1) = trvConsole.SelectedItem.Text
  End If
  
  'TM20011107 Fault 3106
  'Have re-arranged the checking so that the transfers cannot happen if the
  'groups cannot ultimately be deleted.
  For iLoop = 1 To UBound(asSelectedGroups)
    
    sCurrentGroupName = asSelectedGroups(iLoop)
        
      ' Check for fixed roles
      sSQL = "sp_helpdbfixedrole"
      rsRoles.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      With rsRoles
        Do While Not .EOF
          If rsRoles.Fields("DbFixedRole").Value = sCurrentGroupName Then
            MsgBox "You cannot delete the system user group '" & sCurrentGroupName & "'.", vbInformation + vbOKOnly, App.Title
            Exit Function
          End If
          .MoveNext
        Loop
        .Close
      End With
      
      ' RH 08/11/00 - BUG 1311 - Do not let a user group be deleted if its used in any batch jobs
      ' TM20011107 Fault 3107 - Order the items in the recordset.
      rsRoles.Open "SELECT AsrSysBatchJobName.Name, ASRSysBatchJobName.username FROM AsrSysBatchJobName WHERE RoleToPrompt = '" & sCurrentGroupName & "' ORDER BY AsrSysBatchJobName.Name", gADOCon, adOpenForwardOnly, adLockReadOnly
      sBatchJobs = vbNullString
      Do While Not rsRoles.EOF
        sBatchJobs = sBatchJobs & SetStringLength(rsRoles!Name, 50) & vbTab & rsRoles!UserName & vbNewLine
        rsRoles.MoveNext
      Loop
      rsRoles.Close
      
      If sBatchJobs <> vbNullString Then
        sBatchJobs = "Cannot delete the user group '" & sCurrentGroupName & "'." & vbNewLine & "It is used as a schedule prompt in the following batch job(s) : " & vbNewLine & vbNewLine & _
                     "Batch Job" & vbTab & vbTab & vbTab & vbTab & "Owner" & vbNewLine & _
                     "======" & vbTab & vbTab & vbTab & "=====" & vbNewLine & vbNewLine & sBatchJobs
        MsgBox sBatchJobs, vbExclamation + vbOKOnly, App.Title
        Exit Function
      End If
      
      If (sCurrentGroupName = "public") Then
        MsgBox "You cannot delete the system user group 'public'.", vbInformation + vbOKOnly, App.Title
        Exit Function
      End If


  Next iLoop

'********************************************************************************
  
  'TM20010910 Fault 2813
  'Find out how many users are in group. If >= 1 then present transfer ownership stuff.
  'If < 1 then just present 'Are you sure? confirmation.
  GetSelectedUsers asSelectedGroups(), asSelectedUsers()

  If UBound(asSelectedUsers) >= 1 Then
'    If lvList_SelectedCount > 1 Then
'      sMessage = "Before the deletion can be actioned, the utility ownership for " & _
'                  "all the selected users must be transferred to available users. " & _
'                  vbNewLine & "Are you sure you want to delete these " & _
'                  Trim(Str(lvList_SelectedCount)) & " user groups?"
'    Else
'      sMessage = "Before the deletion can be actioned, the utility ownership for " & _
'                  "all the selected users must be transferred to available users. " & _
'                  "Are you sure you want to delete this user group?"
'    End If
    
    sMessage = "Are you sure you want to delete " & IIf((UBound(asSelectedGroups) > 1), "these User Groups", "this User Group") & "?"
    If MsgBox(sMessage, vbQuestion + vbYesNo, App.ProductName) = vbYes Then
      fDeleteAll = True
      fConfirmed = True
    Else
      Exit Function
    End If
    
    Dim frmTransfer As New frmTransferOwnership
  
    With frmTransfer
      If .Initialise(asSelectedUsers) Then
        .Show vbModal
      End If
      fConfirmed = Not .Cancelled
    End With
    Unload frmTransfer
    Set frmTransfer = Nothing
  Else
    sMessage = "Are you sure you want to delete " & IIf((UBound(asSelectedGroups) > 1), "these User Groups", "this User Group") & "?"
    If MsgBox(sMessage, vbQuestion + vbYesNo, App.ProductName) = vbYes Then
      fDeleteAll = True
      fConfirmed = True
    Else
      Exit Function
    End If
  End If
  
'********************************************************************************

  ' Delete all of the groups in the array..
  For iLoop = 1 To UBound(asSelectedGroups)
    
    sCurrentGroupName = asSelectedGroups(iLoop)
  
    If fConfirmed Then

      fDeleteOK = True

      If fDeleteOK Then
      
        ' Ensure that the user group is initialised
        If Not gObjGroups(sCurrentGroupName).Users_Initialised Then
          InitialiseUsersCollection gObjGroups(sCurrentGroupName)
        End If
  
        Call DeleteGroup(sCurrentGroupName)
        trvConsole.Nodes.Remove "GP_" & sCurrentGroupName

        iDeletedCount = iDeletedCount + 1
         
      End If
    End If
  Next iLoop
  
  If fMovedUsers Then
    MsgBox "The users from the deleted user group" & IIf(iDeletedCount > 1, "s", vbNullString) & " have been moved to the 'public' user group.", _
      vbInformation + vbOKOnly, App.Title
  End If
  
  If iDeletedCount > 0 Then
    ' Enable the apply button
    Application.Changed = True
  
    ' Select the root node of the console treeview.
    With trvConsole
      .SelectedItem = trvConsole.Nodes("RT")
      .SelectedItem.EnsureVisible
    End With
    UpdateRightPane
  End If
  
  Set rsRoles = Nothing
  
End Function


Private Function Group_New() As Boolean
  ' Create a new group.
  Dim fValuesOK As Boolean
  Dim sGroupName As String
  Dim vAccess As Variant
  
  Load frmNewGroup
  
  With frmNewGroup
    .Initialise GROUPACTION_NEW
    
    Do While .Tag <> "Cancel"
      
      ' Display the new group form.
      .Show vbModal
    
      ' Check whether the user select OK or cancel.
      If .Tag = "OK" Then
      
        sGroupName = Trim(.GroupName)
        fValuesOK = Group_Valid(sGroupName)
        
        If fValuesOK Then
          
          AddGroup gObjGroups, sGroupName
          Call AddGroupToTreeView(sGroupName)
          
          ' Get the configured access values.
          vAccess = .AccessConfiguration
          If IsArray(vAccess) Then
            gObjGroups(sGroupName).AccessConfiguration = vAccess
          Else
            gObjGroups(sGroupName).AccessCopyGroup = vAccess
          End If
          
          ' Enable the apply button
          Application.Changed = True
          .Tag = "Cancel"
        End If
      End If
    Loop
  End With
  
  ' Unload the new group form
  Unload frmNewGroup

End Function

Private Function WhichGroup(ByRef pNodX As SSNode) As String
  ' Calculate which group we belong to based on the given node.

  Select Case pNodX.Level
    Case "2"
      WhichGroup = pNodX.Text
    Case "3"
      WhichGroup = pNodX.Parent.Text
    Case "4"
      WhichGroup = pNodX.Parent.Parent.Text
    Case Else
      WhichGroup = vbNullString
  End Select
  
End Function


Private Function MoveUserToNewGroup(psLoginName As String, psNewGroupName As String, psOldGroupName As String)
  ' Move a user to another user group.
  Dim objOldGroup As SecurityGroup
  Dim objNewGroup As SecurityGroup
  Dim objUser As SecurityUser
  
  ' Ensure that the user group that the user is being moved to is
  ' not the same one that it is currently in.
  If psOldGroupName <> psNewGroupName Then
  
    ' Get the required group and user objects.
    Set objOldGroup = gObjGroups(psOldGroupName)
    Set objNewGroup = gObjGroups(psNewGroupName)
    Set objUser = objOldGroup.Users(psLoginName)
    
    ' Initialise the new group's user collection if required.
    If Not objNewGroup.Users_Initialised Then
      InitialiseUsersCollection objNewGroup
    End If
    
    ' See if the user has just been added and not applied yet.
    If objUser.NewUser Then
      ' Add the user to the new group as a new user.
      objNewGroup.Users.Add objUser.UserName, True, False, True, _
        vbNullString, vbNullString, objUser.Login, objUser.Password, , , , objUser.LoginType
      ' Delete this user from the original group.
      objOldGroup.Users.Remove psLoginName
    Else
      
      ' The user must already exist in SQL so now decide if the user has already
      ' been moved or not
      If objUser.MovedUserFrom <> vbNullString Then
                        
        ' See if the user is being moved back to the original group where it came from.
        If objUser.MovedUserFrom = psNewGroupName Then
            
          ' Remove the user from the original user group.
          objOldGroup.Users.Remove psLoginName
          ' Reset the flags on the new user group.
          With objNewGroup.Users(psLoginName)
            .MovedUserTo = vbNullString
          End With
          
        Else
          ' The user has been moved to another group already so remove them from this
          ' group, add them to the new group and change where they have been
          ' moved to in the original group

          ' Original Group Update
          gObjGroups(objUser.MovedUserFrom).Users(psLoginName).MovedUserTo = psNewGroupName

          ' Remove the user from the original group.
          objOldGroup.Users.Remove psLoginName

          ' Add the user to the new group.
          objNewGroup.Users.Add objUser.UserName, True, False, False, _
            objUser.MovedUserFrom, vbNullString, objUser.Login, objUser.Password

        End If
      Else
      
        ' The user has not already been moved so move it to the new group
        ' and add it to the tree
        objNewGroup.Users.Add objUser.UserName, True, False, False, _
          psOldGroupName, vbNullString, objUser.Login, objUser.Password, , , , objUser.LoginType
          
        ' Change the properties on the moved user to indicate where it has been
        ' moved to.
        objUser.Changed = True
        objUser.MovedUserTo = psNewGroupName
        
      End If
    End If
  
    ' Disassociate object variables.
    Set objOldGroup = Nothing
    Set objNewGroup = Nothing
    Set objUser = Nothing
  
  End If
  
End Function




Private Function trvConsole_LoadTablesViews(pNode As SSNode) As Boolean
  ' Add nodes to the given tree view node for the selected group's collection of tables and views.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sImage As String
  Dim sTableName As String
  Dim sCurrentGroup As String
  Dim objView As SecurityTable
  Dim objTable As SecurityTable
  Dim rsTables As New ADODB.Recordset
  
  fOK = True
  
  ' Get the selected group name.
  sCurrentGroup = WhichGroup(pNode)
  
  ' Add the tables to the tree view.
  For Each objTable In gObjGroups(sCurrentGroup).Tables
    With objTable
      Select Case objTable.TableType
        Case tabParent
          sImage = "TOPLEVELTABLE"
        Case tabChild
          sImage = "CHILDTABLE"
        Case tabLookup
          sImage = "LOOKUPTABLE"
        Case Else
          sImage = "TABLE"
      End Select
      
      trvConsole.Nodes.Add pNode.Key, tvwChild, "TB_" & pNode.Key & Trim$(.Name), Trim$(.Name), sImage, , "TABLE"
    End With
  Next objTable

  ' Add nodes to the given node in the tree view for each view associated with the current user group.
  For Each objView In gObjGroups(sCurrentGroup).Views
    With objView
        
      ' Get the name of the views table.
      sSQL = "SELECT ASRSysTables.tableName" & _
        " FROM ASRSysTables, ASRSysViews" & _
        " WHERE ASRSysViews.viewName = '" & .Name & "'" & _
        " AND ASRSysViews.viewTableID = ASRSysTables.tableID"
      rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not rsTables.EOF And Not rsTables.BOF Then
        sTableName = rsTables.Fields(0).Value
      Else
        sTableName = "<unknown>"
      End If
      rsTables.Close
        
      trvConsole.Nodes.Add pNode.Key, tvwChild, "VW_" & pNode.Key & .Name, sTableName & " - '" & Trim$(.Name) & "' view", "VIEW", , "VIEW"
    End With
  Next

TidyUpAndExit:
  Set objView = Nothing
  Set rsTables = Nothing
  Set objTable = Nothing
  
  trvConsole_LoadTablesViews = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function UpdateRightPane()
  ' This function displays the required control in the right pane.
  Dim fDisplayListView As Boolean
  Dim fDisplayColumnsGrid As Boolean
  Dim fDisplaySystemPermissions As Boolean
  
  fDisplayListView = False
  fDisplayColumnsGrid = False
  fDisplaySystemPermissions = False
  
  If trvConsole.SelectedItem Is Nothing Then
    lblRightPaneCaption.Caption = vbNullString
  Else
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS", "GROUP", "USERS", "TABLESVIEWS"
        fDisplayListView = True
        
      Case "VIEW", "TABLE"
        fDisplayColumnsGrid = True
        
      Case "SYSTEM"
        fDisplaySystemPermissions = True
    End Select
  End If
  
  If fDisplayListView Then
    lvList_Refresh
  End If
  
  If fDisplayColumnsGrid Then
    ssgrdColumns_Refresh
  End If
  
  If fDisplaySystemPermissions Then
    sstrvSystemPermissions_Refresh
  End If
  
  lvList.Visible = fDisplayListView
  ssgrdColumns.Visible = fDisplayColumnsGrid
  sstrvSystemPermissions.Visible = fDisplaySystemPermissions
  
End Function

Private Sub lvList_RefreshTablesViews()
  ' Populate the listview with the users for the selected user group.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sImage As String
  Dim sCurrentGroup As String
  Dim sTableName As String
  Dim ThisItem As ComctlLib.ListItem
  Dim objTable As SecurityTable
  Dim objView As SecurityTable
  Dim rsTables As New ADODB.Recordset

  Const iEXTRALENGTH = 3
  
  fOK = True
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
   
  ' Update the caption.
  
  'MH20010208 Fault 1825
  'lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group -  Tables / Views"
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group -  Data Permissions"
        
  With lvList.ColumnHeaders
    .Add , , "Table/View"
'    .Add , , "'Read' Permission"
'    .Add , , "'New' Permission"
'    .Add , , "'Edit' Permission"
'    .Add , , "'Delete' Permission"
'NHRD16112006Fault????
    .Add , , "'New' Permission"
    .Add , , "'Edit' Permission"
    .Add , , "'Read' Permission"
    .Add , , "'Delete' Permission"
    .Add , , "Related to Parents"
  End With
  
  ' Add the tables to the listview.
  For Each objTable In gObjGroups(sCurrentGroup).Tables
    With objTable
      Select Case objTable.TableType
        Case tabParent
          sImage = "TOPLEVELTABLE"
        Case tabChild
          sImage = "CHILDTABLE"
        Case tabLookup
          sImage = "LOOKUPTABLE"
        Case Else
          sImage = "TABLE"
      End Select
      
      Set ThisItem = lvList.ListItems.Add(, "TB_TV_GP_" & sCurrentGroup & Trim(.Name), Trim(.Name), sImage, sImage)
      
      If ((Len(Trim(.Name)) + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC)) > mlngFirstColumnWidth Then
        mlngFirstColumnWidth = ((Len(Trim(.Name)) + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC))
      End If
      'NHRD17112006Fault 10737, 8525
      ThisItem.SubItems(1) = IIf(.InsertPrivilege, "Full", "None") 'New
      ThisItem.SubItems(2) = PrivilegeDescription(.UpdatePrivilege) 'Edit
      ThisItem.SubItems(3) = PrivilegeDescription(.SelectPrivilege) 'Read
      ThisItem.SubItems(4) = IIf(.DeletePrivilege, "Full", "None") 'Delete
      ThisItem.SubItems(5) = IIf(.ParentCount > 1, IIf(.ParentJoinType = 1, "All", "Any"), vbNullString) 'Related To Parents
      
'      ThisItem.SubItems(1) = PrivilegeDescription(.SelectPrivilege)
'      ThisItem.SubItems(2) = IIf(.InsertPrivilege, "Full", "None")
'      ThisItem.SubItems(3) = PrivilegeDescription(.UpdatePrivilege)
'      ThisItem.SubItems(4) = IIf(.DeletePrivilege, "Full", "None")
'      ThisItem.SubItems(5) = IIf(.ParentCount > 1, IIf(.ParentJoinType = 1, "All", "Any"), vbNullString)
    End With
  Next objTable
          
  ' Add the views to the listview.
  For Each objView In gObjGroups(sCurrentGroup).Views
    With objView
      ' Get the name of the views table.
      sSQL = "SELECT ASRSysTables.tableName" & _
        " FROM ASRSysTables, ASRSysViews" & _
        " WHERE ASRSysViews.viewName = '" & .Name & "'" & _
        " AND ASRSysViews.viewTableID = ASRSysTables.tableID"
      rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not rsTables.EOF And Not rsTables.BOF Then
        sTableName = rsTables!TableName
      Else
        sTableName = "<unknown>"
      End If
      rsTables.Close
      
      Set ThisItem = lvList.ListItems.Add(, "VW_TV_GP_" & sCurrentGroup & Trim(.Name), sTableName & " - '" & Trim(.Name) & "' view", "VIEW", "VIEW")
    
      If ((Len(sTableName & " - '" & Trim(.Name) & "' view") + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC)) > mlngFirstColumnWidth Then
        mlngFirstColumnWidth = ((Len(sTableName & " - '" & Trim(.Name) & "' view") + iEXTRALENGTH) * UI.GetAvgCharWidth(Me.hDC))
      End If
    
      'NHRD17112006Fault 10737, 8525
      ThisItem.SubItems(1) = IIf(.InsertPrivilege, "Full", "None") 'New
      ThisItem.SubItems(2) = PrivilegeDescription(.UpdatePrivilege) 'Edit
      ThisItem.SubItems(3) = PrivilegeDescription(.SelectPrivilege) 'Read
      ThisItem.SubItems(4) = IIf(.DeletePrivilege, "Full", "None") 'Delete
      ThisItem.SubItems(5) = vbNullString

    End With
  Next objView
          
  ' If no items are selected then try to select the first one.
  If (lvList_SelectedCount = 0) And (lvList.ListItems.Count > 0) Then
    lvList.SelectedItem = lvList.ListItems(1)
    lvList.SelectedItem.EnsureVisible
    lvList.ColumnHeaders(1).Width = mlngFirstColumnWidth
  End If
  
TidyUpAndExit:
'  Set WaitWindow = Nothing
  Set rsTables = Nothing
  Set ThisItem = Nothing
  Set objView = Nothing
  Set objTable = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub




Private Sub lvList_Refresh()
  ' Populate the listview with the required data for the selected treeview node.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  fOK = True
  
  ' Clear the list view items.
  lvList.ListItems.Clear
  lvList.ColumnHeaders.Clear

  If Not trvConsole.SelectedItem Is Nothing Then
  
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        lvList.Sorted = True
        lvList_RefreshGroups
        
      Case "GROUP"
        lvList.Sorted = False
        lvList_RefreshGroup
        
      Case "USERS"
        lvList.Sorted = True
        lvList_RefreshUsers
              
      Case "TABLESVIEWS"
        lvList.Sorted = True
        lvList_RefreshTablesViews
    End Select
  
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Private Sub lvList_RefreshGroups()
  ' Populate the listview with the available user groups.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ThisItem As ComctlLib.ListItem
  Dim objGroup As SecurityGroup

  fOK = True
  
  ' Update the caption.
  lblRightPaneCaption.Caption = " User Groups"
        
  With lvList
    
    'JDM - 19/12/01 - Fault 3305 - Icons are scattered all over the place
    .LabelWrap = False
    
    ' Add each user group to the listview.
    For Each objGroup In gObjGroups
      If Not objGroup.DeleteGroup Then
        Set ThisItem = .ListItems.Add(, objGroup.Name, objGroup.Name, "GROUP", "GROUP")
      End If
    Next
      
    ' If no items are selected then try to select the first one.
    If (lvList_SelectedCount = 0) And (.ListItems.Count > 0) Then
      .SelectedItem = .ListItems(1)
      .SelectedItem.EnsureVisible
    End If
    
    .LabelWrap = True
    
  End With

TidyUpAndExit:
  Set ThisItem = Nothing
  Set objGroup = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Private Sub lvList_RefreshGroup()
  ' Populate the listview with the group security categories.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sCurrentGroup As String
  Dim ThisItem As ComctlLib.ListItem

  fOK = True
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  
  ' Update the caption.
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group -  Security Categories"
  
  With lvList
  
    ' JDM - 12/12/01 - Fault 2809 - Turn off label wrap so that icons are arranged properly in large icon mode
    .LabelWrap = False
   
    ' Add the security categories to the listview.
    Set ThisItem = .ListItems.Add(, "USERS", "User Logins", "CLOSEDFOLDER", "CLOSEDFOLDER")
    Set ThisItem = .ListItems.Add(, "TABLESVIEWS", "Data Permissions", "CLOSEDFOLDER", "CLOSEDFOLDER")
    Set ThisItem = .ListItems.Add(, "SYSTEM", "System Permissions", "SYSTEM", "SYSTEM")

    ' If no items are selected then try to select the first one.
    If (lvList_SelectedCount = 0) And (.ListItems.Count > 0) Then
      .SelectedItem = .ListItems(1)
      .SelectedItem.EnsureVisible
    End If
    
    ' Turn label wrapping on
    .LabelWrap = True
    
  End With
  
TidyUpAndExit:
  Set ThisItem = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Private Sub lvList_RefreshUsers()
  ' Populate the listview with the logins for the selected user group.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sCurrentGroup As String
  Dim ThisItem As ComctlLib.ListItem
  Dim objUser As SecurityUser
  Dim sIcon As String

  fOK = True
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  
  ' Update the caption.
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group -  User logins"
        
  ' Fill in the users collection information if is has not yet been read.
  If Not gObjGroups(sCurrentGroup).Users_Initialised Then
    InitialiseUsersCollection gObjGroups(sCurrentGroup)
  End If
        
  With lvList
    .LabelWrap = False
    
    .ColumnHeaders.Add , , "User Login"
    
    ' Add the users to the listview.
    For Each objUser In gObjGroups(sCurrentGroup).Users
      If (Not objUser.DeleteUser) And _
        (objUser.MovedUserTo = vbNullString Or objUser.MovedUserTo = sCurrentGroup) Then
        
        ' JPD20020430 Fault 3813 - Changed the key to make sure it is not a numeric value,
        ' as the 'add' function as causing an error when the key was numeric (eg. "1", "11").
        Select Case objUser.LoginType
          Case iUSERTYPE_TRUSTEDUSER
            sIcon = "USER"
        
          Case iUSERTYPE_TRUSTEDGROUP
            sIcon = "GROUP"
        
          Case iUSERTYPE_SQLLOGIN
            sIcon = "USER_SQL"
          
          ' NPG20090204 Fault 11931
          Case iUSERTYPE_ORPHANUSER
            sIcon = "USER_ORPHAN"
            
          Case iUSERTYPE_ORPHANGROUP
            sIcon = "GROUP_ORPHAN"
            
        End Select
        
        Set ThisItem = .ListItems.Add(, "key" & objUser.UserName, objUser.UserName, sIcon, sIcon)

      End If
    Next
          
    ' If no items are selected then try to select the first one.
    If (lvList_SelectedCount = 0) And (.ListItems.Count > 0) Then
      .SelectedItem = .ListItems(1)
      .SelectedItem.EnsureVisible
    End If
    
    .LabelWrap = True
  
  End With
  
TidyUpAndExit:
  Set ThisItem = Nothing
  Set objUser = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub





Public Function lvList_SelectedCount() As Integer
  ' Return the count of views selected in the listview.
  Dim iLoop As Integer
  
  lvList_SelectedCount = 0
  
  ' Loop through the list view items counting how many
  ' are currently selected.
  For iLoop = 1 To lvList.ListItems.Count
    If lvList.ListItems(iLoop).Selected = True Then
      lvList_SelectedCount = lvList_SelectedCount + 1
    End If
  Next iLoop

End Function

' NPG20090206 Fault 11931
Public Function lvList_OrphanCount() As Integer
  ' Return the count of orphaned windows logins displayed in the listview.
  Dim iLoop As Integer
  
  lvList_OrphanCount = 0
  
  ' Loop through the list view items counting how many
  ' are orphans (Have no Login in SQL).
  For iLoop = 1 To lvList.ListItems.Count
    If Right(Trim(lvList.ListItems(iLoop).Icon), 6) = "ORPHAN" Then
      lvList_OrphanCount = lvList_OrphanCount + 1
    End If
  Next iLoop

End Function

' NPG20090206 Fault 11931
Public Function lvList_SelectAllOrphans() As Integer
  ' Select all orphans in this listview.
  Dim iLoop As Integer
    
  ' Loop through the list view items and select them if they're orphans
  For iLoop = 1 To lvList.ListItems.Count
    If Right(Trim(lvList.ListItems(iLoop).Icon), 6) = "ORPHAN" Then
      lvList.ListItems(iLoop).Selected = True
    Else
      lvList.ListItems(iLoop).Selected = False
    End If
  Next iLoop

End Function

' NPG20090206 Fault 11931
Public Function lvList_SelectedOrphanCount() As Integer
  ' Return the count of orphaned windows logins SELECTED in the listview.
  Dim iLoop As Integer
  
  lvList_SelectedOrphanCount = 0
  
  ' Loop through the list view items counting how many
  ' are orphans (Have no Login in SQL).
  For iLoop = 1 To lvList.ListItems.Count
    If Right(Trim(lvList.ListItems(iLoop).Icon), 6) = "ORPHAN" And lvList.ListItems(iLoop).Selected = True Then
      lvList_SelectedOrphanCount = lvList_SelectedOrphanCount + 1
    End If
  Next iLoop

End Function



Private Sub RefreshStatusBar()
  Dim sMessage As String
  
  sMessage = vbNullString
  
  If Not trvConsole.SelectedItem Is Nothing Then
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        sMessage = " " & Trim(Str(lvList.ListItems.Count)) & " user group" & IIf(lvList.ListItems.Count <> 1, "s", vbNullString) & _
          ", " & Trim(Str(lvList_SelectedCount)) & " selected."
          
      Case "GROUP"
        sMessage = " User group security categories."
        
      Case "USERS"
        sMessage = " " & Trim(Str(lvList.ListItems.Count)) & " user login" & IIf(lvList.ListItems.Count <> 1, "s", vbNullString) & _
          ", " & Trim(Str(lvList_SelectedCount)) & " selected."
          
      Case "TABLESVIEWS"
        sMessage = Trim(Str(lvList.ListItems.Count)) & " table" & IIf(lvList.ListItems.Count <> 1, "s", vbNullString) & _
          " / view" & IIf(lvList.ListItems.Count <> 1, "s", vbNullString) & _
          ", " & Trim(Str(lvList_SelectedCount)) & " selected."
          
      Case "TABLE"
        sMessage = " Table columns."
          
      Case "VIEW"
        sMessage = " View columns."
          
      Case "SYSTEM"
        sMessage = " System permissions."
    End Select
  End If
  
  sbStatus.Panels(1).Text = sMessage
  
End Sub


Public Sub EditMenu(psMenuItem As String)

  Dim sCurrentGroup As String
  Dim objGroup As SecurityGroup
  Dim sNodeKeyBeforeSave As String
  Dim iViewTypeBeforeSave As String
  Dim nodTemp As SSNode
  
  'Get the selected item type
  If Not trvConsole.SelectedItem Is Nothing Then
    ' Process the menu selection depending on what is currently selected.
    Select Case psMenuItem
      
      Case "ID_SecurityNew"
        EditMenu_New

      Case "ID_SecurityCopy"
        EditMenu_CopyGroup

      Case "ID_SecurityAutomaticAdd"
        'NHRD18022003 Fault 3413 If the user is not a System Administrator
        'then there is no point going any further
        If Not gbUserCanManageLogins Then
          MsgBox "New users can only be created by a system administrator.", vbInformation + vbOKOnly, App.Title
        Else
          EditMenu_AutomaticAdd
        End If
        
      Case "ID_SecurityDelete"
        EditMenu_Delete
      
      Case "ID_SecurityMove"
        EditMenu_Move
      
      Case "ID_SecurityResetPassword"
        EditMenu_ResetPassword
      
      Case "ID_FindUser"
        EditMenu_FindUser
      
      Case "ID_SecurityProperties"
        EditMenu_Properties
      
      Case "ID_SecuritySelectAll"
        EditMenu_SelectAll
     
      Case "ID_SecurityPrint"
        EditMenu_Print
        
      Case "ID_SecuritySave"
        If Application.Changed Then
          If MsgBox("Save changes." & vbNewLine & _
                    "Are you sure ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            
            'TM20011015 Fault 2914
            'Code added to remember the view settings and selected node before the
            'save so that they can be re-set after the save.
            sNodeKeyBeforeSave = trvConsole(trvConsole.SelectedItem).Key
            iViewTypeBeforeSave = Me.lvList.View
            
            If ApplyChanges Then
              Application.Changed = False
              'TM20010920 Fault 2530
              'Refresh the currently selected group collection.
              sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
              
              'Collapse all the current Security Group nodes that are not
              'for the current group. This is done so the groups can be
              're-initialised when expanded.
              'TM20011107 Fault 3104 - the current group may no longer exist.
              If sCurrentGroup <> vbNullString Then
                For Each nodTemp In trvConsole.Nodes
                  If UCase$(nodTemp.Text) <> UCase$(sCurrentGroup) Then
                    nodTemp.Expanded = False
                  End If
                Next nodTemp
                Set nodTemp = Nothing
                lvList.View = iViewTypeBeforeSave
                trvConsole(sNodeKeyBeforeSave).Selected = True
                
                ' JDM - 05/04/04 - Fault 8378 - Have to do this twice as code trigger above resets these settings
                lvList.View = iViewTypeBeforeSave
                trvConsole(sNodeKeyBeforeSave).Selected = True
                
                InitialiseGroupCollections sCurrentGroup, False
              End If
              UpdateRightPane
            End If
          End If
        End If
      
      Case "ID_LargeIcons"
        ChangeView lvwIcon
       
      Case "ID_SmallIcons"
        ChangeView lvwSmallIcon
       
      Case "ID_List"
        ChangeView lvwList
       
      Case "ID_Details"
        ChangeView lvwReport
      
      'NHRD02062003 Fault 2173
      Case "ID_CheckAll"
        EditMenu_ToggleSystemPermissions ("CheckAll")
      'NHRD02062003 Fault 2173
      Case "ID_UnCheckAll"
        EditMenu_ToggleSystemPermissions ("UnCheckAll")
    
    End Select
    
    RefreshSecurityMenu
  End If

End Sub

Private Function RefreshDataPermissionsList()
              
  lvList_Refresh
              
End Function



Private Sub ssgrdColumns_Refresh()
  ' Load the column grid with the column info for the currently selected view/table.
  ssgrdColumns.RemoveAll
  
  If Not trvConsole.SelectedItem Is Nothing Then
    Select Case trvConsole.SelectedItem.DataKey
      Case "TABLE"
        ssgrdColumns_RefreshTableColumns
        
      Case "VIEW"
        ssgrdColumns_RefreshViewColumns
    End Select
  End If
  
  RefreshStatusBar

End Sub


Private Sub ssgrdColumns_RefreshViewColumns()
  ' Load the column grid with the column info for the currently selected view.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fAllColumnSelect As Boolean
  Dim fAllColumnUpdate As Boolean
  Dim sViewName As String
  Dim sCurrentGroup As String
  Dim objColumn As SecurityColumn
  
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  sViewName = Mid(trvConsole.SelectedItem.Key, Len(trvConsole.SelectedItem.Parent.Key) + 4)
  
  ' Update the right pane caption.
  lblRightPaneCaption.Caption = " '" & sCurrentGroup & "' group - '" & sViewName & "' view columns"
  
  With ssgrdColumns.Columns("Select")
    .Locked = False
    .ForeColor = vbBlack
  End With

  fAllColumnUpdate = True
  fAllColumnSelect = True
  For Each objColumn In gObjGroups(sCurrentGroup).Views(sViewName).Columns
    ssgrdColumns.AddItem objColumn.Name & _
      vbTab & objColumn.SelectPrivilege & _
      vbTab & objColumn.UpdatePrivilege
      
    ' Check for all columns being updateable and selectable
    If fAllColumnSelect And Not objColumn.SelectPrivilege Then
      fAllColumnSelect = False
    End If
    If fAllColumnUpdate And Not objColumn.UpdatePrivilege Then
      fAllColumnUpdate = False
    End If
  Next
  
  ssgrdColumns.AddItem "(All Columns)" & _
    vbTab & fAllColumnSelect & _
    vbTab & fAllColumnUpdate, 0

TidyUpAndExit:
  Set objColumn = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Private Sub EditMenu_SelectAll()

'NPG20080611 Fault 13196
'  ' Select all items in the list view.
'  lvList_SelectAll
'  lvList.SetFocus
'
'  'TM20020116 Fault 3361
'  RefreshStatusBar

    Select Case trvConsole.SelectedItem.DataKey
      Case "SYSTEM"
        EditMenu_ToggleSystemPermissions ("CheckAll")
      Case Else
        ' Select all items in the list view.
        lvList_SelectAll
        lvList.SetFocus
    End Select
  
  'TM20020116 Fault 3361
  RefreshStatusBar
End Sub

Private Sub EditMenu_New()
  ' Add a new item to the list view depending on the active view,
  ' and type of object selected in the treeview.
  If ActiveView Is trvConsole Then
    
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        Group_New
      Case "GROUP"
        Group_New
      Case "USERS"
        User_New
    End Select
    
  ElseIf ActiveView Is lvList Then
  
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        Group_New
      Case "USERS"
        User_New
    End Select
  
  End If
  
End Sub

Private Sub EditMenu_Delete()
  ' Delete the selected item(s).
  
  If ActiveView Is trvConsole Then
    
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUP"
        Group_Delete
    End Select
  
  ElseIf ActiveView Is lvList Then
    
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        Group_Delete
      Case "USERS"
        User_Delete
    End Select
  
  End If
  
End Sub

Private Sub EditMenu_ResetPassword()

  Dim objForm As New frmPasswordMaintenance
  
  If ActiveView Is lvList Then
    If trvConsole.SelectedItem.DataKey = "USERS" Then
      If (lvList_SelectedCount = 1) Then
  
        objForm.ShowAllUsers = False
        objForm.UserName = lvList.SelectedItem.Text
        objForm.SecurityGroup = WhichGroup(trvConsole.SelectedItem)
        objForm.Initialise
        objForm.Show vbModal

      End If
    End If
  End If

  Unload objForm

End Sub

Private Sub EditMenu_Properties()
  ' Change the properties of the selected item.
  
  If ActiveView Is trvConsole Then
  
    Select Case trvConsole.SelectedItem.DataKey
      Case "TABLE"
        TableView_Properties
      Case "VIEW"
        TableView_Properties
      Case "GROUP"            'MH20000410
        Group_Properties
    End Select
    
  ElseIf ActiveView Is lvList Then
  
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS"
        Group_Properties
      Case "TABLESVIEWS"
        TableView_Properties
      Case "USERS"
    End Select
    
  End If
  
End Sub

Private Sub EditMenu_Print()
  Dim objGroup As SecurityGroup
  Dim objPrintDef As clsPrintDef
  Dim astrGroups() As String
  Dim strGroups As String
  
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim mlngBottom As Long
  Dim fCancelled As Boolean
  
  ReDim gasPrintGroups(0)
  ReDim gasPrintOptions(1)
  Call ResetPrintArray(1, False)
  glngPageNum = 0
  gstrPrintGroupName = vbNullString
  
  ' Different defaults for what to print.
  If ActiveView Is trvConsole Then
    
    Select Case trvConsole.SelectedItem.DataKey
      Case "GROUPS", "GROUP"
        frmSecurityPrintOptions.SetDefaultOptions True, True, False, True
        strGroups = trvConsole.SelectedItem.Text & ";"
      Case "USERS"
        frmSecurityPrintOptions.SetDefaultOptions True, False, False, False
        strGroups = trvConsole.SelectedItem.Parent.Text & ";"
      Case "TABLESVIEWS"
        frmSecurityPrintOptions.SetDefaultOptions False, True, False, False
        strGroups = trvConsole.SelectedItem.Parent.Text & ";"
      Case "SYSTEM"
        frmSecurityPrintOptions.SetDefaultOptions False, False, False, True
        strGroups = trvConsole.SelectedItem.Parent.Text & ";"
    End Select
 
  ElseIf ActiveView Is lvList Then
    
    Select Case lvList.SelectedItem.Text
      Case "User Logins"
        frmSecurityPrintOptions.SetDefaultOptions True, False, False, False
        strGroups = trvConsole.SelectedItem.Text & ";"
        
      Case "Data Permissions"
        frmSecurityPrintOptions.SetDefaultOptions False, True, False, False
        strGroups = trvConsole.SelectedItem.Text & ";"
        
      Case "System Permissions"
        frmSecurityPrintOptions.SetDefaultOptions False, False, False, True
        strGroups = trvConsole.SelectedItem.Text & ";"
        
      Case Else
        For iLoop = 1 To lvList.ListItems.Count
          If lvList.ListItems.Item(iLoop).Selected = True Then
            strGroups = strGroups & lvList.ListItems.Item(iLoop).Text & ";"
          End If
        Next iLoop
        
    End Select
    
    
  End If
 
  ' Pop up dialog to see what they want to print
  frmSecurityPrintOptions.Show vbModal
  fCancelled = frmSecurityPrintOptions.Cancelled
  Set frmSecurityPrintOptions = Nothing
  
  If fCancelled Then
    Exit Sub
  End If
  
  'Start the progress bar here
  With gobjProgress
    '.AviFile = App.Path & "\videos\table.Avi"
    '.AviFile = App.Path & "\videos\DB_Transfer.Avi"
    .AVI = dbTransfer
    .Caption = "Retrieving Print information ..."
    .MainCaption = "Printing"
    .NumberOfBars = 0
    .Time = False
    .Cancel = False
    .OpenProgress
  End With
  
  'Need to initialise everything anyway
  astrGroups = Split(strGroups, ";")
  
  'NHRD28062004 Fault 8520
  If strGroups = "User Groups;" Then
        For Each objGroup In gObjGroups
          InitialiseGroupCollections objGroup.Name, False
          If Not objGroup.Users_Initialised Then
            InitialiseUsersCollection objGroup
          End If
        Next objGroup
  Else      ' JDM - 30/03/2004 - Fault 8224 - Only initialise the groups we actually need
        For iLoop = LBound(astrGroups) To UBound(astrGroups) - 1
          InitialiseGroupCollections astrGroups(iLoop), False
          If Not gObjGroups(astrGroups(iLoop)).Users_Initialised Then
            InitialiseUsersCollection gObjGroups(astrGroups(iLoop))
          End If
        Next iLoop
  End If

  gobjProgress.CloseProgress
  
  Set objPrintDef = New clsPrintDef
  'NHRD04032004 Fault 7880 Added print object stuff.
  If objPrintDef.IsOK Then
      With objPrintDef
          If .PrintStart(False) Then
              If ActiveView Is trvConsole Then
                  Select Case trvConsole.SelectedItem.DataKey
                      Case "GROUPS"
                          'gasPrintOptions(1).PrintLPaneGROUPS = True
                          gObjGroups.PrintSecurity 0
                      Case "GROUP"
                          'gasPrintOptions(1).PrintLPaneGROUP = True
                          gObjGroups(trvConsole.SelectedItem.Text).PrintSecurity 0
                      Case "USERS", "SYSTEM", "TABLESVIEWS"
                          gObjGroups(trvConsole.SelectedItem.Parent.Text).PrintSecurity 0
                  End Select
              ElseIf ActiveView Is lvList Then
                  If trvConsole.SelectedItem.DataKey = "GROUPS" Then
                  
                      If (lvList_SelectedCount > 1) Then
                      
                          ' Read the names of the groups to be printed from the listview into an array.
                          For iLoop = 1 To lvList.ListItems.Count
                              If lvList.ListItems(iLoop).Selected = True Then
                                  iNextIndex = UBound(gasPrintGroups) + 1
                                  ReDim Preserve gasPrintGroups(iNextIndex)
                                  gasPrintGroups(iNextIndex) = lvList.ListItems(iLoop).Text
                              End If
                          Next iLoop
                      Else
                          ' Read the name of the group to be printed from the treeview into an array.
                          ReDim gasPrintGroups(1)
                          gasPrintGroups(1) = lvList.SelectedItem.Text
                      End If
                        
                      'gasPrintOptions(1).PrintRPaneGROUPS = True
                      gObjGroups.PrintSecurity 0
                        
                  Else
                  
                    ' User logins, Data or System Permissions
                    iNextIndex = UBound(gasPrintGroups) + 1
                    ReDim Preserve gasPrintGroups(iNextIndex)
                    gasPrintGroups(iNextIndex) = trvConsole.SelectedItem.Text
                    gObjGroups(trvConsole.SelectedItem.Text).PrintSecurity 0
                  
                  End If
              End If
          End If
      End With
  End If

  'End Document Print
  Printer.EndDoc
  'Printer.KillDoc
  
End Sub
Private Sub EditMenu_ToggleSystemPermissions(strToggleValue As String)
  'JPD 20030911 Fault 6033 & Fault 6034
  'NHRD02062003 Fault 2173
  ' Load the System Permissions treeview with the selected user group's permissions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sCurrentGroup As String
  Dim objOriginalNode As SSActiveTreeView.SSNode
  Dim objNode As SSActiveTreeView.SSNode
  Dim strPicture As String
  Dim fSysMgr As Boolean
  Dim fSecMgr As Boolean
  Dim sMessage As String
  
  fOK = True
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  Set objOriginalNode = sstrvSystemPermissions.SelectedItem
  
  fSysMgr = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed
  fSecMgr = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed

  strPicture = IIf(strToggleValue = "CheckAll", "SYSIMG_TICK", "SYSIMG_NOTICK")
  
  Select Case strToggleValue
    Case "CheckAll"
      If fSysMgr Or fSecMgr Then
        sMessage = "All System Permissions (excluding Module Access) are already granted for this user group as it has access to the " & IIf(fSysMgr, "System Manager", "Security Manager") & " module."
        MsgBox sMessage, vbOKOnly + vbInformation, App.ProductName
        Exit Sub
      Else
        sMessage = "All System Permissions (excluding Module Access) will be granted for this user group."
          
        If MsgBox(sMessage & vbNewLine & vbNewLine & _
            "Are you sure you want to continue ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
          Exit Sub
        End If
      End If
      
    Case "UnCheckAll"
      If fSysMgr Or fSecMgr Then
        sMessage = "No System Permissions (excluding Module Access) can be revoked for this user group as it has access to the " & IIf(fSysMgr, "System Manager", "Security Manager") & " module."
        MsgBox sMessage, vbOKOnly + vbInformation, App.ProductName
        Exit Sub
      Else
        sMessage = "All System Permissions (excluding Module Access) will be revoked for this user group."
        
        If MsgBox(sMessage & vbNewLine & vbNewLine & _
            "Are you sure you want to continue ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
          Exit Sub
        End If
      End If
  End Select

  'Change permissions
  For Each objNode In sstrvSystemPermissions.Nodes
    If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
      If objNode.Parent.Key <> "C_MODULEACCESS" Then    ' The item is not a module access item.
      
        If objNode.Key = "P_ACCORD_SENDRECORD" Then
        
          If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed = True Or _
            gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed = True Then
              objNode.Image = strPicture
              gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = (strToggleValue = "CheckAll")
          End If
        Else
          objNode.Image = strPicture
          gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = (strToggleValue = "CheckAll")
        End If
      
      End If
    End If
  Next objNode
  Set objNode = Nothing

  ' Apply any system permission constraints.
  RelateSystemPermissions

  ' Enable the apply button
  Application.Changed = True

  'MH20011019 Fault 2983
  gObjGroups(sCurrentGroup).RequireLogout = True

  RefreshSecurityMenu

  If Not gobjCurrentNode Is Nothing Then
    Set sstrvSystemPermissions.SelectedItem = gobjCurrentNode
  End If
  
  sstrvSystemPermissions.SetFocus

'  If strToggleValue = "CheckAll" Then
'    MsgBox "All System Permissions have been granted.", vbInformation, App.ProductName
'  Else
'    MsgBox "All System Permissions have been revoked.", vbInformation, App.ProductName
'  End If

TidyUpAndExit:
  Set objNode = Nothing
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub lvList_SelectAll()
  ' Select all items in the list view.
  Dim iLoop As Integer
  
  ' Loop through the list view items marking each one as selected.
  For iLoop = 1 To lvList.ListItems.Count
    lvList.ListItems(iLoop).Selected = True
  Next iLoop

End Sub



Public Sub ChangeView(piViewStyle As Integer)
  ' Set the view style for the list view.
  lvList.View = piViewStyle
  RefreshSecurityMenu
  
  Select Case Left(trvConsole.SelectedItem.Key, 2)
    Case "RT"
      giListView_ViewGroups = piViewStyle
    Case "GP"
      giListView_ViewCategories = piViewStyle
    Case "US"
      giListView_ViewUsers = piViewStyle
    Case "TV"
      giListView_ViewTables = piViewStyle
  End Select
  
End Sub


Private Function TableView_Properties() As Boolean
  ' Display the properties of the selected tables/views.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fSelected As Boolean
  Dim fSomethingChanged As Boolean
  Dim iCount As Integer
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim fNoLookups As Boolean
  Dim fAllLookups As Boolean
  Dim sKey As String
  Dim sName As String
  Dim sCurrentGroupName As String
  Dim sCurrentTableViewName As String
  Dim objTableView As SecurityTable
  Dim objColumn As SecurityColumn
  Dim fSelectAllGranted As Boolean
  Dim fSelectNoneGranted As Boolean
  Dim fInsertAllGranted As Boolean
  Dim fInsertNoneGranted As Boolean
  Dim fUpdateAllGranted As Boolean
  Dim fUpdateNoneGranted As Boolean
  Dim fDeleteAllGranted As Boolean
  Dim fDeleteNoneGranted As Boolean
  Dim fViewMenuAllGranted As Boolean
  Dim fViewMenuNoneGranted As Boolean
  Dim fViewRecordEditAllGranted As Boolean
  Dim fViewRecordEditNoneGranted As Boolean
  Dim iSelectPermission As ColumnPrivilegeStates
  Dim iUpdatePermission As ColumnPrivilegeStates
  Dim iInsertPermission As ColumnPrivilegeStates
  Dim iDeletePermission As ColumnPrivilegeStates
  Dim iViewMenuPermission As ColumnPrivilegeStates
  Dim iViewRecordEditPermission As ColumnPrivilegeStates
  Dim avSelectedTablesViews() As Variant
  Dim frmProps As frmTableViewProperties
  Dim fMultiParentChildren As Boolean
  Dim iMultiParentChildren_All As Integer
  Dim iMultiParentChildren_Any As Integer
  Dim sMessage As String
  Dim sTempMessage As String
  Dim sTableName As String
  Dim iIndex As Integer
  
  ReDim avSelectedTablesViews(0)
  fOK = True
  sCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  fSomethingChanged = False
  fNoLookups = True
  fAllLookups = True
  fMultiParentChildren = False
  iMultiParentChildren_All = 0
  iMultiParentChildren_Any = 0
  
  ' Construct an array of the table/view objects that are selected.
  If ActiveView Is lvList Then
    For iLoop = 1 To lvList.ListItems.Count
      If lvList.ListItems(iLoop).Selected = True Then
        iNextIndex = UBound(avSelectedTablesViews) + 1
        ReDim Preserve avSelectedTablesViews(iNextIndex)
        
        sKey = lvList.ListItems(iLoop).Key
        If (Left(sKey, 3) = "TB_") Then
          sCurrentTableViewName = Mid(sKey, Len("TB_TV_GP_" & sCurrentGroupName) + 1)
          Set objTableView = gObjGroups(sCurrentGroupName).Tables(sCurrentTableViewName)
        Else
          sCurrentTableViewName = Mid(sKey, Len("VW_TV_GP_" & sCurrentGroupName) + 1)
          Set objTableView = gObjGroups(sCurrentGroupName).Views(sCurrentTableViewName)
        End If
        Set avSelectedTablesViews(iNextIndex) = objTableView

        If objTableView.TableType = tabLookup Then
          fNoLookups = False
        Else
          fAllLookups = False
        End If

        Set objTableView = Nothing
      End If
    Next iLoop
  
    If UBound(avSelectedTablesViews) > 1 Then
      sName = Trim(Str(UBound(avSelectedTablesViews))) & " selected tables/views"
    Else
      sName = lvList.SelectedItem.Text
    End If
  Else
    ReDim avSelectedTablesViews(1)
    
    sName = trvConsole.SelectedItem.Text
    sKey = trvConsole.SelectedItem.Key

    If (Left(sKey, 3) = "TB_") Then
      sCurrentTableViewName = Mid(sKey, Len("TB_TV_GP_" & sCurrentGroupName) + 1)
      Set objTableView = gObjGroups(sCurrentGroupName).Tables(sCurrentTableViewName)
    Else
      sCurrentTableViewName = Mid(sKey, Len("VW_TV_GP_" & sCurrentGroupName) + 1)
      Set objTableView = gObjGroups(sCurrentGroupName).Views(sCurrentTableViewName)
    End If
    Set avSelectedTablesViews(1) = objTableView
    
    If objTableView.TableType = tabLookup Then
      fNoLookups = False
    Else
      fAllLookups = False
    End If
    
    Set objTableView = Nothing
  End If
  
  ' Get the aggregate permissions for the selected tables/views.
  fSelectAllGranted = True
  fSelectNoneGranted = True
  fInsertAllGranted = True
  fInsertNoneGranted = True
  fUpdateAllGranted = True
  fUpdateNoneGranted = True
  fDeleteAllGranted = True
  fDeleteNoneGranted = True
  fViewMenuAllGranted = True
  fViewMenuNoneGranted = True
  
  For iLoop = 1 To UBound(avSelectedTablesViews)
    If avSelectedTablesViews(iLoop).SelectPrivilege <> giPRIVILEGES_ALLGRANTED Then
      fSelectAllGranted = False
    End If
    If avSelectedTablesViews(iLoop).SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
      fSelectNoneGranted = False
    End If
    If avSelectedTablesViews(iLoop).UpdatePrivilege <> giPRIVILEGES_ALLGRANTED Then
      fUpdateAllGranted = False
    End If
    If avSelectedTablesViews(iLoop).UpdatePrivilege <> giPRIVILEGES_NONEGRANTED Then
      fUpdateNoneGranted = False
    End If
    If avSelectedTablesViews(iLoop).InsertPrivilege Then
      fInsertNoneGranted = False
    Else
      fInsertAllGranted = False
    End If
    If avSelectedTablesViews(iLoop).DeletePrivilege Then
      fDeleteNoneGranted = False
    Else
      fDeleteAllGranted = False
    End If
    If avSelectedTablesViews(iLoop).HideFromMenu Then
      fViewMenuNoneGranted = False
    Else
      fViewMenuAllGranted = False
    End If
    
    If avSelectedTablesViews(iLoop).ParentCount > 1 Then
      fMultiParentChildren = True
      
      If avSelectedTablesViews(iLoop).ParentJoinType = 1 Then
        iMultiParentChildren_All = iMultiParentChildren_All + 1
      Else
        iMultiParentChildren_Any = iMultiParentChildren_Any + 1
      End If
    End If
  Next iLoop
  
  If fSelectAllGranted Then
    iSelectPermission = giPRIVILEGES_ALLGRANTED
  ElseIf fSelectNoneGranted Then
    iSelectPermission = giPRIVILEGES_NONEGRANTED
  Else
    iSelectPermission = giPRIVILEGES_SOMEGRANTED
  End If
  
  If fUpdateAllGranted Then
    iUpdatePermission = giPRIVILEGES_ALLGRANTED
  ElseIf fUpdateNoneGranted Then
    iUpdatePermission = giPRIVILEGES_NONEGRANTED
  Else
    iUpdatePermission = giPRIVILEGES_SOMEGRANTED
  End If
  
  If fInsertAllGranted Then
    iInsertPermission = giPRIVILEGES_ALLGRANTED
  ElseIf fInsertNoneGranted Then
    iInsertPermission = giPRIVILEGES_NONEGRANTED
  Else
    iInsertPermission = giPRIVILEGES_SOMEGRANTED
  End If
  
  If fDeleteAllGranted Then
    iDeletePermission = giPRIVILEGES_ALLGRANTED
  ElseIf fDeleteNoneGranted Then
    iDeletePermission = giPRIVILEGES_NONEGRANTED
  Else
    iDeletePermission = giPRIVILEGES_SOMEGRANTED
  End If
  
  If fViewMenuAllGranted Then
    iViewMenuPermission = giPRIVILEGES_ALLGRANTED
  ElseIf fViewMenuNoneGranted Then
    iViewMenuPermission = giPRIVILEGES_NONEGRANTED
  Else
    iViewMenuPermission = giPRIVILEGES_SOMEGRANTED
  End If
  
  ' Display the table/view properties form.
  Set frmProps = New frmTableViewProperties
  
  With frmProps
    ' Initialise the properties form with the properties of the current table/view.
    .TableViewName = sName
    .SecurityGroupName = sCurrentGroupName
    .LookupTableStatus = IIf(fNoLookups, LOOKUPTABLES_NONE, IIf(fAllLookups, LOOKUPTABLES_ALL, LOOKUPTABLES_SOME))

    .SelectPermission = iSelectPermission
    .UpdatePermission = iUpdatePermission
    .InsertPermission = iInsertPermission
    .DeletePermission = iDeletePermission
    .HideFromMenu = iViewMenuPermission
    .MultiParentChildren = fMultiParentChildren
    .MultiParentJoinType = IIf(iMultiParentChildren_All > iMultiParentChildren_Any, 1, 0)
    .UserGroup = sCurrentGroupName

    ' Call the table/view properties form.
    .Show vbModal
    
    ' Check whether the user select OK or cancel.
    If .Tag = "OK" Then
    
      '******************************************************
      If .SelectPermission = giPRIVILEGES_NONEGRANTED _
        And .UpdatePermission = giPRIVILEGES_NONEGRANTED _
        And .InsertPermission = giPRIVILEGES_NONEGRANTED _
        And .DeletePermission = giPRIVILEGES_NONEGRANTED Then
        
        sMessage = vbNewLine
        For iLoop = 1 To UBound(avSelectedTablesViews)
          Set objTableView = avSelectedTablesViews(iLoop)
          sTempMessage = CheckTidyPermissionOnChilds(objTableView, sCurrentGroupName, avSelectedTablesViews)
          Set objTableView = Nothing
          
          ' Ensure duplicates are not added to the message.
          Do While Len(sTempMessage) > 0
            iIndex = InStr(sTempMessage, vbNewLine)
            If iIndex = 0 Then
              sTempMessage = vbNullString
            Else
              sTableName = Left(sTempMessage, iIndex - 1)
              sTempMessage = Mid(sTempMessage, iIndex + 2)
              
              If InStr(sMessage, vbNewLine & sTableName & vbNewLine) = 0 Then
                sMessage = sMessage & sTableName & vbNewLine
              End If
            End If
          Loop
        Next iLoop
        
        If sMessage <> vbNewLine Then
          sMessage = "Removing 'read' permission will also remove all permissions to the following child tables:" & vbNewLine & _
            sMessage & vbNewLine & _
            "Do you wish to continue?"
        
          If (MsgBox(sMessage, vbYesNo + vbQuestion, App.Title) <> vbYes) Then
            fOK = False
            GoTo TidyUpAndExit
          End If
        End If
      End If
      '******************************************************
      
      ' Save the new properties to the selected tables/views.
      For iLoop = 1 To UBound(avSelectedTablesViews)
        Set objTableView = avSelectedTablesViews(iLoop)
      
        ' If the table/view is a top-level table/view then
        ' check if the select permission has changed from All/Some to None, or vice versa.
        ' If so then flag all children of this table as changed. This is done as the
        ' permitted child views on the children need to be recalculated.
          ' Do nothing if the given table.view is a child or lookup table.
        If ((objTableView.TableType <> tabChild) And (objTableView.TableType <> tabLookup)) Then
          If ((.SelectPermission <> giPRIVILEGES_NONEGRANTED) And (objTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED)) Or _
            ((.SelectPermission = giPRIVILEGES_NONEGRANTED) And (objTableView.SelectPrivilege <> giPRIVILEGES_NONEGRANTED)) Then
            FlagChildrenChanged objTableView, sCurrentGroupName
          End If
        End If

        If (objTableView.SelectPrivilege <> IIf(objTableView.TableType = tabLookup, giPRIVILEGES_ALLGRANTED, .SelectPermission)) Or _
          (objTableView.UpdatePrivilege <> .UpdatePermission) Or _
          (objTableView.InsertPrivilege <> (.InsertPermission = giPRIVILEGES_ALLGRANTED)) Or _
          (objTableView.DeletePrivilege <> (.DeletePermission = giPRIVILEGES_ALLGRANTED)) Or _
          (objTableView.HideFromMenu <> (.HideFromMenu = giPRIVILEGES_ALLGRANTED)) Then
          objTableView.Changed = True
        End If
        objTableView.SelectPrivilege = IIf(objTableView.TableType = tabLookup, giPRIVILEGES_ALLGRANTED, .SelectPermission)
        objTableView.UpdatePrivilege = .UpdatePermission
        objTableView.InsertPrivilege = (.InsertPermission = giPRIVILEGES_ALLGRANTED)
        objTableView.DeletePrivilege = (.DeletePermission = giPRIVILEGES_ALLGRANTED)
        
        If fAllLookups And (.HideFromMenu <> giPRIVILEGES_SOMEGRANTED) Then
          objTableView.HideFromMenu = .HideFromMenu
        End If
        
        If Not objTableView.Changed Then
          If (objTableView.TableType = tabChild) And _
            (objTableView.ParentJoinType <> .MultiParentJoinType) Then
            objTableView.Changed = True
          End If
        End If
        objTableView.ParentJoinType = .MultiParentJoinType
        
        If objTableView.Changed Then
          fSomethingChanged = True
        End If
        
        ' Propogate the select and update privileges to all the columns for this table/view
        ' if the privilege has changed.
        If .SelectPermission <> giPRIVILEGES_SOMEGRANTED Then
          For Each objColumn In objTableView.Columns
            objColumn.Changed = (objColumn.SelectPrivilege <> IIf(objTableView.TableType = tabLookup, True, (.SelectPermission = giPRIVILEGES_ALLGRANTED)))
            objColumn.SelectPrivilege = IIf(objTableView.TableType = tabLookup, True, (.SelectPermission = giPRIVILEGES_ALLGRANTED))
          
            If objColumn.Changed Then
              fSomethingChanged = True
            End If
          Next objColumn
          Set objColumn = Nothing
        End If

        If .UpdatePermission <> giPRIVILEGES_SOMEGRANTED Then
          For Each objColumn In objTableView.Columns
            objColumn.Changed = (objColumn.UpdatePrivilege <> (.UpdatePermission = giPRIVILEGES_ALLGRANTED))
            objColumn.UpdatePrivilege = (.UpdatePermission = giPRIVILEGES_ALLGRANTED)
          
            If objColumn.Changed Then
              fSomethingChanged = True
            End If
          Next objColumn
          Set objColumn = Nothing
        End If
      Next iLoop
        
      ' Check if any of the selected changes could not be made due to
      ' child tables not having any 'readable'' parents.
      ' Tidy up the permissions first.
      For Each objTableView In gObjGroups(sCurrentGroupName).Tables
        TidyPermissionOnChilds objTableView, sCurrentGroupName
      Next objTableView
      Set objTableView = Nothing
      
      For Each objTableView In gObjGroups(sCurrentGroupName).Views
        TidyPermissionOnChilds objTableView, sCurrentGroupName
      Next objTableView
      Set objTableView = Nothing
      
      If .SelectPermission <> giPRIVILEGES_NONEGRANTED Then
        sMessage = vbNewLine
        iCount = 0
        For iLoop = 1 To UBound(avSelectedTablesViews)
          Set objTableView = avSelectedTablesViews(iLoop)
          
          If objTableView.SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
            ' The 'read' permission must have been force to 'none' as part of the 'tidy up',
            ' so inform the user.
            sMessage = sMessage & objTableView.Name & vbNewLine
            iCount = iCount + 1
          End If
          Set objTableView = Nothing
        Next iLoop
        
        If iCount > 0 Then
          If iCount = 1 Then
            sMessage = "Permissions cannot be granted to the following table as it has no parents that have 'read' permission granted:" & vbNewLine & _
              sMessage
          Else
            sMessage = "Permissions cannot be granted to the following tables as they have no parents that have 'read' permission granted:" & vbNewLine & _
              sMessage
          End If
          
          MsgBox sMessage, vbOKOnly + vbInformation, App.Title
        End If
      End If
      
      ' Enable the apply button
      If fSomethingChanged Then
        gObjGroups(sCurrentGroupName).Changed = True
        gObjGroups(sCurrentGroupName).RequireLogout = True  'MH20010410
        Application.Changed = True
      End If
      
      UpdateRightPane

      If ActiveView Is lvList Then
        ' Ensure the original items are still selected in the listview if it still exists.
        For iLoop = 1 To lvList.ListItems.Count
          fSelected = False
          For iNextIndex = 1 To UBound(avSelectedTablesViews)
            If (lvList.ListItems(iLoop).Key = "TB_TV_GP_" & sCurrentGroupName & Trim(avSelectedTablesViews(iNextIndex).Name)) Or _
              (lvList.ListItems(iLoop).Key = "VW_TV_GP_" & sCurrentGroupName & Trim(avSelectedTablesViews(iNextIndex).Name)) Then
              
              fSelected = True
              Exit For
            End If
          Next iNextIndex
          
          lvList.ListItems(iLoop).Selected = fSelected
        Next iLoop
      End If
    End If
  End With
  
  Unload frmProps
  Set frmProps = Nothing
  
TidyUpAndExit:
  Set objTableView = Nothing
  TableView_Properties = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Function lvList_ValidKey(ByVal pvKey As Variant) As Boolean
  On Error GoTo ErrorTrap
  
  If lvList.ListItems(pvKey).Key = pvKey Then
    lvList_ValidKey = True
  End If
  
  Exit Function
  
ErrorTrap:
  lvList_ValidKey = False
  If Err.Number <> 35601 Then
    MsgBox "Runtime error '" & Trim(Str(Err.Number)) & "'." & _
      vbCr & vbCr & Err.Description, _
      vbExclamation + vbOKOnly, "Microsoft Visual Basic"
  End If
  Err = False
  
End Function


Private Function PrivilegeDescription(piState As ColumnPrivilegeStates) As String
  ' Return a string describing the column privilege states.
  Select Case piState
    Case giPRIVILEGES_ALLGRANTED
      PrivilegeDescription = "Full"
    Case giPRIVILEGES_NONEGRANTED
      PrivilegeDescription = "None"
    Case giPRIVILEGES_SOMEGRANTED
      PrivilegeDescription = "Partial"
    Case Else
      PrivilegeDescription = "<unknown>"
  End Select
  
End Function

Private Sub sstrvSystemPermissions_Initialise()
  ' Populate the System Permissions treeview with the required nodes.
  Dim sSQL As String
  Dim sFileName As String
  Dim sFirstKey As String
  Dim objCategoryNode As SSActiveTreeView.SSNode
  Dim rsCategories As New ADODB.Recordset
  Dim rsItems As New ADODB.Recordset
  Dim objImage As Object
  Dim strSQLWhere As String
  
  sFirstKey = vbNullString
  strSQLWhere = vbNullString
  
  With sstrvSystemPermissions
    ' Clear any existing nodes.
    .Nodes.Clear
    
    ' Clear any existing images.
    For Each objImage In imgSystemPermissions.ListImages
      If objImage.Tag <> "SYSTEM" Then
        objImage.Key = vbNullString
      End If
    Next objImage
    Set objImage = Nothing

    sSQL = "SELECT * FROM ASRSysPermissionCategories"

    ' Limit system permission if CMG module is not enabled
    'If Not frmSystem.IsModuleEnabled(CMG) Then
    If Not IsModuleEnabled(modCMG) Then
      strSQLWhere = strSQLWhere & IIf(InStr(strSQLWhere, "WHERE") > 0, " AND ", " WHERE ") & "ASRSysPermissionCategories.categoryKey <> 'CMG'"
    End If
    
    ' Hide Accord if module not enabled
    If Not IsModuleEnabled(modAccord) Then
      strSQLWhere = strSQLWhere & IIf(InStr(strSQLWhere, "WHERE") > 0, " AND ", " WHERE ") & "ASRSysPermissionCategories.categoryKey <> 'ACCORD'"
    End If
    
    ' Hide Workflow if module not enabled
    If Not IsModuleEnabled(modWorkflow) Then
      strSQLWhere = strSQLWhere & IIf(InStr(strSQLWhere, "WHERE") > 0, " AND ", " WHERE ") & "ASRSysPermissionCategories.categoryKey <> 'WORKFLOW'"
    End If
    
    ' Hide Workflow if module not enabled
    If Not IsModuleEnabled(modVersionOne) Then
      strSQLWhere = strSQLWhere & IIf(InStr(strSQLWhere, "WHERE") > 0, " AND ", " WHERE ") & "ASRSysPermissionCategories.categoryKey <> 'VERSION1'"
    End If
    
    ' Hide Nine Box Grid Reports if module not enabled
    If Not IsModuleEnabled(modNineBoxGrid) Then
      strSQLWhere = strSQLWhere & IIf(InStr(strSQLWhere, "WHERE") > 0, " AND ", " WHERE ") & "ASRSysPermissionCategories.categoryKey <> 'NINEBOXGRID'"
    End If
        
    ' Add order by clause
    sSQL = sSQL & strSQLWhere & " ORDER BY listOrder, description"
    
    rsCategories.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsCategories
      Do While Not .EOF
      
        If Not IsNull(!Picture) Then
          ' Add the image to the image list.
          sFileName = ReadPicture(!Picture)
          imgSystemPermissions.ListImages.Add , "IMG_" & !CategoryKey, LoadPicture(sFileName)
          Kill sFileName
    
          sstrvSystemPermissions.Images.Add !categoryID, !CategoryKey, "IMG_" & !CategoryKey
          Set objCategoryNode = sstrvSystemPermissions.Nodes.Add(, , "C_" & !CategoryKey, !Description, "IMG_" & !CategoryKey)
        Else
          Set objCategoryNode = sstrvSystemPermissions.Nodes.Add(, , "C_" & !CategoryKey, !Description, "SYSIMG_UNKNOWN")
        End If
      
        If sFirstKey = vbNullString Then
          sFirstKey = "C_" & !CategoryKey
        End If
        
        objCategoryNode.Font.Bold = True
          
        ' Add the category items.
        sSQL = "SELECT *" & _
          " FROM ASRSysPermissionItems" & _
          " WHERE categoryID = " & Trim(Str(!categoryID)) & _
          " ORDER BY listOrder, description"
        
        rsItems.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        With rsItems
          Do While Not .EOF
            'sstrvSystemPermissions.Nodes.Add "C_" & rsCategories!CategoryKey, tvwChild, "P_" & rsCategories!CategoryKey & "_" & !ItemKey, !Description, "SYSIMG_NOTICK"
            sstrvSystemPermissions.Nodes.Add "C_" & rsCategories!CategoryKey, tvwChild, "P_" & rsCategories!CategoryKey & "_" & !ItemKey, !Description, IIf(mblnReadOnly, "SYSIMG_GREYNOTICK", "SYSIMG_NOTICK")

            .MoveNext
          Loop

          .Close
        End With
        
        objCategoryNode.Expanded = True
        Set objCategoryNode = Nothing
        
        .MoveNext
      Loop
      
      .Close
    End With
    
    Set rsCategories = Nothing
    Set rsItems = Nothing
    
    If sFirstKey <> vbNullString Then
      .SelectedItem = .Nodes(sFirstKey)
    End If
  End With

End Sub
Private Function ReadPicture(pobjPictureField As ADODB.Field) As String
  Dim iLoop As Integer
  Dim iChunks As Integer
  Dim iTempFile As Integer
  Dim iFragment As Integer
  Dim lngColumnSize As Long
  Dim sTempName As String
  Dim abytChunk() As Byte
  
  Const ChunkSize = 2 ^ 14

  sTempName = GetTemporaryFileName
  iTempFile = 1
  Open sTempName For Binary Access Write As iTempFile
  
  lngColumnSize = pobjPictureField.ActualSize
  iChunks = lngColumnSize \ ChunkSize
  iFragment = lngColumnSize Mod ChunkSize
      
  ReDim abytChunk(iFragment)
  abytChunk() = pobjPictureField.GetChunk(iFragment)
  Put iTempFile, , abytChunk()
      
  For iLoop = 1 To iChunks
    ReDim abytChunk(ChunkSize)
    abytChunk() = pobjPictureField.GetChunk(ChunkSize)
    Put iTempFile, , abytChunk()
  Next iLoop
  Close iTempFile
  
  ReadPicture = sTempName
  
End Function


Private Function GetTemporaryFileName() As String
  ' Return a unique temporary file name.
  Dim sTmpPath As String
  Dim sTmpName As String
  
  sTmpPath = Space(1024)
  sTmpName = Space(1024)

  Call GetTempPath(1024, sTmpPath)
  Call GetTempFileName(sTmpPath, "_T", 0, sTmpName)
  
  sTmpName = Trim(sTmpName)
  If Len(sTmpName) > 0 Then
    sTmpName = Left(sTmpName, Len(sTmpName) - 1)
  Else
    sTmpName = vbNullString
  End If
    
  GetTemporaryFileName = Trim(sTmpName)
  
End Function


Private Sub ToggleSystemPermission(pobjNode As SSActiveTreeView.SSNode)
  ' Toggle the state of the given node in the system permission treeview.
  Dim fApplyPermission As Boolean
  Dim iIndex As Integer
  Dim sCurrentGroup As String
  Dim sMessage As String
  Dim sItemKey As String
  Dim objNode As SSActiveTreeView.SSNode

  If mblnReadOnly Then
    Exit Sub
  End If

  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
   
  ' Do nothing if the node is locked.
  If pobjNode.Image = "SYSIMG_GREYTICK" Or pobjNode.Image = "SYSIMG_GREYNOTICK" Then
    ' Tell the user why they cannot change the item.
    'NHRD05022004 Fault 3621 Changed 'revoke' to 'alter' as the box can be SYSIMG_GREYNOTICK which isn't a revoke.
    sMessage = "Unable to alter this permission."
    
    If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
      If (pobjNode.Key <> "P_MODULEACCESS_SSINTRANET") Then
        sMessage = sMessage & vbNewLine & _
         "It is required if permission is granted to run the Security Manager."
      End If
    ElseIf gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
      If (pobjNode.Key <> "P_MODULEACCESS_SSINTRANET") Then
        sMessage = sMessage & vbNewLine & _
         "It is required if permission is granted to run the Security Manager."
      End If
    ElseIf (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_DIARY_SYSTEMEVENTS").Allowed) And _
      (pobjNode.Key = "P_DIARY_MANUALEVENTS") Then
      sMessage = sMessage & vbNewLine & _
        "It is required if permission is granted to view System Generated Diary Events."
    ElseIf (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EVENTLOG_PURGE").Allowed) And _
      (pobjNode.Key = "P_EVENTLOG_VIEWALL") Then
      sMessage = sMessage & vbNewLine & _
        "It is required if permission is granted to purge event log entries."
    ElseIf (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EVENTLOG_PURGE").Allowed) And _
      (pobjNode.Key = "P_EVENTLOG_DELETE") Then
      sMessage = sMessage & vbNewLine & _
        "It is required if permission is granted to purge event log entries."
    ElseIf ((gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed) Or _
      (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_BLOCK").Allowed)) And _
      (pobjNode.Key = "P_ACCORD_VIEWTRANSFER") Then
      sMessage = sMessage & vbNewLine & _
        "It is required if permission is granted to block or create transfers."
    ElseIf (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_WORKFLOW_ADMINISTER").Allowed) And _
      (pobjNode.Key = "P_WORKFLOW_VIEWLOG") Then
      sMessage = sMessage & vbNewLine & _
        "It is required if permission is granted to administer workflows."
    Else
      iIndex = InStrRev(pobjNode.Key, "_")
      If iIndex > 0 Then
        sItemKey = Mid(pobjNode.Key, iIndex + 1)
              
        Select Case UCase$(sItemKey)
          Case "EDIT"
            sMessage = sMessage & vbNewLine & _
              "It is required if permission is granted to create " & _
              pobjNode.Parent.Text & IIf(UCase$(Right(pobjNode.Parent.Text, 1)) = "S", ".", " definitions.")
          Case "VIEW"
            sMessage = sMessage & vbNewLine & _
              "It is required if permission is granted to create or edit  " & _
              pobjNode.Parent.Text & IIf(UCase$(Right(pobjNode.Parent.Text, 1)) = "S", ".", " definitions.")
        End Select
      End If
    End If

    MsgBox sMessage, vbOKOnly + vbExclamation, App.ProductName
    Set sstrvSystemPermissions.SelectedItem = pobjNode
    sstrvSystemPermissions.SetFocus
    Exit Sub
  End If
  
  ' IF ITS SEC/SYS BEING REVOKED AND ITS THE CURRENT USERS ROLE, STOPEM !
  If (pobjNode.Key = "P_MODULEACCESS_SECURITYMANAGER") And gsUserName <> "sa" Then
    If (gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed) Then
      
      If Not gObjGroups(sCurrentGroup).Users_Initialised Then
        InitialiseUsersCollection gObjGroups(sCurrentGroup)
      End If
      Dim intCount As Integer
      Dim fFound As Boolean
      
      For intCount = 1 To gObjGroups(sCurrentGroup).Users.Count
        If LCase(gObjGroups(sCurrentGroup).Users.Item(intCount).UserName) = LCase(gsUserName) Then
          fFound = True
          Exit For
        End If
      Next intCount
      
      If fFound Then
        MsgBox "You cannot revoke 'Security Manager' permission. You" & vbNewLine & _
               "are logged in as a username contained in this group.", vbExclamation + vbOKOnly, App.Title
        Set sstrvSystemPermissions.SelectedItem = pobjNode
        sstrvSystemPermissions.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  ' IF ITS AN INTRANET ITEM AND INTRANET ACCESS IS REVOKED, STOPEM !
  
  ' RH 31/05/01 - Check the parent isnt nothing, otherwise get a runtime error
  ' when clicking on the node titles (CROSSTABS, SYSTEM MANAGER, CUSTOM REPORTS etc).
  If Not pobjNode.Parent Is Nothing Then
    
    If (pobjNode.Parent.Key = "C_INTRANET") And _
      (pobjNode.Image = "SYSIMG_NOTICK") And _
      (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_INTRANET").Allowed) Then
       
      MsgBox "Unable to grant this permission as access to the Data Manager Intranet (multiple record access) module is denied.", vbOKOnly + vbExclamation, App.ProductName
      Set sstrvSystemPermissions.SelectedItem = pobjNode
      sstrvSystemPermissions.SetFocus
      Exit Sub
    End If
  End If
    
  ' JDM - Fault 9832 - 24/02/2005 - If its an Accord Export - make sure they have system or security access
  If IsModuleEnabled(modAccord) Then

    If (pobjNode.Key = "P_ACCORD_SENDRECORD") And (pobjNode.Image = "SYSIMG_NOTICK") _
      And (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed) Then
        MsgBox "Unable to grant this permission as write access to the System Manager module is denied.", vbOKOnly + vbExclamation, App.ProductName
        Set sstrvSystemPermissions.SelectedItem = pobjNode
        sstrvSystemPermissions.SetFocus
        Exit Sub
    End If
    
    If (pobjNode.Key = "P_MODULEACCESS_SYSTEMMANAGER") And _
      (pobjNode.Image = "SYSIMG_TICK") Then
      sstrvSystemPermissions.Nodes.Item("P_ACCORD_SENDRECORD").Image = "SYSIMG_NOTICK"
      gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed = False
    End If
  
    If (pobjNode.Key = "P_MODULEACCESS_SECURITYMANAGER") And _
      (pobjNode.Image = "SYSIMG_TICK") Then
      sstrvSystemPermissions.Nodes.Item("P_ACCORD_SENDRECORD").Image = "SYSIMG_NOTICK"
      gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed = False
    End If
  
  End If
  
  If Left(pobjNode.Key, 1) = "P" Then
    fApplyPermission = (pobjNode.Image = "SYSIMG_NOTICK")
    
    pobjNode.Image = IIf(fApplyPermission, "SYSIMG_TICK", "SYSIMG_NOTICK")
    gObjGroups(sCurrentGroup).SystemPermissions.Item(pobjNode.Key).Allowed = fApplyPermission
      
    ' Apply specific functionality to specific system permission items.
    Select Case pobjNode.Key
      Case "P_MODULEACCESS_SYSTEMMANAGER"
        If fApplyPermission And (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed) Then
        
          If MsgBox("Permission to run the System Manager requires the" & vbNewLine & _
            "user group to have full access to all tables and views," & vbNewLine & _
            "and all other system permissions." & vbNewLine & _
            "These permissions will be granted automatically." & vbNewLine & vbNewLine & _
            "Are you sure you want to continue ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
            
            ' Reset the permission flag if the user does not want to grant
            ' the permission to run the System Manager.
            pobjNode.Image = "SYSIMG_NOTICK"
            gObjGroups(sCurrentGroup).SystemPermissions.Item(pobjNode.Key).Allowed = False
          End If
        End If
    
      Case "P_MODULEACCESS_SECURITYMANAGER"
        If fApplyPermission And (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed) Then
          If MsgBox("Permission to run the Security Manager requires the" & vbNewLine & _
            "user group to have full access to all tables and views," & vbNewLine & _
            "and all other system permissions." & vbNewLine & _
            "These permissions will be granted automatically." & vbNewLine & vbNewLine & _
            "Are you sure you want to continue ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
            
            ' Reset the permission flag if the user does not want to grant
            ' the permission to run the Security Manager.
            pobjNode.Image = "SYSIMG_NOTICK"
            gObjGroups(sCurrentGroup).SystemPermissions.Item(pobjNode.Key).Allowed = False
          End If
        End If
    End Select
      
    ' Apply any system permission constraints.
    RelateSystemPermissions
    
    ' Enable the apply button
    Application.Changed = True
    
    'MH20011019 Fault 2983
    gObjGroups(sCurrentGroup).RequireLogout = True
    
    RefreshSecurityMenu
    
    Set sstrvSystemPermissions.SelectedItem = pobjNode
    sstrvSystemPermissions.SetFocus
  End If

End Sub

Private Sub RelateSystemPermissions()
  ' Handle any relations between system permissions.
  Dim iIndex As Integer
  Dim sItemKey As String
  Dim sCategoryKey As String
  Dim sCurrentGroup As String
  Dim sAssociatedItemKey As String
  Dim sAssociatedItemKey2 As String
  Dim objNode As SSActiveTreeView.SSNode
  
  sCurrentGroup = WhichGroup(trvConsole.SelectedItem)
  
  '
  ' MODULE ACCESS constraints
  '
  ' If a user group is granted permission to run the System or Security Managers then
  ' they must have full access to all table/views, and be granted all other system permissions.
  If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
    gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
    
    ' Grant all permissions on all columns in all tables and views.
    gObjGroups(sCurrentGroup).Tables.GrantAll (sCurrentGroup)
    gObjGroups(sCurrentGroup).Views.GrantAll (sCurrentGroup)
    gObjGroups(sCurrentGroup).RequireLogout = True    'MH20010410

    ' Grant all other system permissions.
    For Each objNode In sstrvSystemPermissions.Nodes
      If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
        If objNode.Parent.Key <> "C_MODULEACCESS" Then    ' The item is not a module access item.
          objNode.Image = "SYSIMG_GREYTICK"
          gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = True
        Else
          Select Case objNode.Key
          Case "P_MODULEACCESS_SYSTEMMANAGERRO"
            If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
              objNode.Image = "SYSIMG_GREYTICK"
              gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = True
            End If
          
          Case "P_MODULEACCESS_SECURITYMANAGERRO"
            If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
              objNode.Image = "SYSIMG_GREYTICK"
              gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = True
            End If

          End Select
        End If
      End If
    Next objNode
    Set objNode = Nothing

  End If

  ' If a user group is not granted permission to run the System or Security Managers then
  ' do not lock other system permissions.
  If (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed) Or _
    (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed) Then

    ' Unlock all other system permissions.
    For Each objNode In sstrvSystemPermissions.Nodes
      If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
        If gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed Then
          If objNode.Parent.Key <> "C_MODULEACCESS" Then    ' The item is not a module access item.
            If (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed) And _
               (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed) Then
              objNode.Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
            End If
          Else
            Select Case objNode.Key
            Case "P_MODULEACCESS_SYSTEMMANAGERRO"
              If Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
                objNode.Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
              End If
            Case "P_MODULEACCESS_SECURITYMANAGERRO"
              If Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
                objNode.Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
              End If
            End Select
          End If
        End If
      End If
    Next objNode
    Set objNode = Nothing
  End If
  
  Dim fAllowed, fAllowedRO As Boolean
  For Each objNode In sstrvSystemPermissions.Nodes

    Select Case objNode.Key
      Case "P_MODULEACCESS_SYSTEMMANAGER"
        fAllowed = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed
        objNode.Image = IIf(fAllowed, "SYSIMG_TICK", "SYSIMG_NOTICK")
      Case "P_MODULEACCESS_SYSTEMMANAGERRO"
        'If System Manager RW is unticked then the System Manager RO will be a grey ticked
        fAllowed = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed
        fAllowedRO = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGERRO").Allowed
        
        If fAllowed Then
          objNode.Image = "SYSIMG_GREYTICK"
        Else
          If fAllowedRO Then
            objNode.Image = "SYSIMG_TICK"
          Else
            objNode.Image = "SYSIMG_NOTICK"
          End If
        End If
      Case "P_MODULEACCESS_SECURITYMANAGER"
        fAllowed = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
        objNode.Image = IIf(fAllowed, "SYSIMG_TICK", "SYSIMG_NOTICK")
      Case "P_MODULEACCESS_SECURITYMANAGERRO"
        'If Security Manager RW is unticked then the Security Manager RO will be a grey ticked
        fAllowed = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
        fAllowedRO = gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGERRO").Allowed
        
        If fAllowed Then
          objNode.Image = "SYSIMG_GREYTICK"
        Else
          If fAllowedRO Then
            objNode.Image = "SYSIMG_TICK"
          Else
            objNode.Image = "SYSIMG_NOTICK"
          End If
        End If
    End Select
  Next objNode

  ' JPD 9/4/01 - If the user is denied Full access to the intranet module, deny access to all Intranet
  ' items, and lock them.
  If (Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_MODULEACCESS_INTRANET").Allowed) Then
    
    ' Unlock all other system permissions.
    For Each objNode In sstrvSystemPermissions.Nodes
      If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
        If objNode.Parent.Key = "C_INTRANET" Then    ' The item is not a module access item.
          If gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed Then
            objNode.Image = "SYSIMG_NOTICK"
            gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed = False
          End If
        End If
      End If
    Next objNode
    Set objNode = Nothing
  End If
  
  '
  ' DIARY constraints
  '
  ' If a user group is granted permission to view System Generated Diary Events then
  ' they must have permission to view Manually Entered Diary Events.
  If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_DIARY_SYSTEMEVENTS").Allowed Then
    ' Grant permission to View Manual Diary Events.
    sstrvSystemPermissions.Nodes("P_DIARY_MANUALEVENTS").Image = "SYSIMG_GREYTICK"
    gObjGroups(sCurrentGroup).SystemPermissions.Item("P_DIARY_MANUALEVENTS").Allowed = True
  End If
    
  ' If a user group is not granted permission to view System Generated Diary Events then
  ' do not lock their permission to view Manually Entered Diary Events.
  If Not gObjGroups(sCurrentGroup).SystemPermissions.Item("P_DIARY_SYSTEMEVENTS").Allowed Then
    ' Grant permission to View Manual Diary Events.
    If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_DIARY_MANUALEVENTS").Allowed Then
      sstrvSystemPermissions.Nodes("P_DIARY_MANUALEVENTS").Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
    End If
  End If

  '
  ' EVENT LOG constraints
  '
  ' If a user group is granted permission to purge event log entries then
  ' they must have permission to view all, and delete event log entries.
  If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EVENTLOG_PURGE").Allowed Then
    ' Grant permission to View all and delete event log entries.
    sstrvSystemPermissions.Nodes("P_EVENTLOG_DELETE").Image = "SYSIMG_GREYTICK"
    gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EVENTLOG_DELETE").Allowed = True
    sstrvSystemPermissions.Nodes("P_EVENTLOG_VIEWALL").Image = "SYSIMG_GREYTICK"
    gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EVENTLOG_VIEWALL").Allowed = True
  End If

  '
  ' WORKFLOW constraints
  '
  ' If a user group is granted permission to Administer workflows then
  ' they must have permission to view the workflow log.
  If IsModuleEnabled(modWorkflow) Then
    If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_WORKFLOW_ADMINISTER").Allowed Then
      ' Grant permission to View the log.
      sstrvSystemPermissions.Nodes("P_WORKFLOW_VIEWLOG").Image = "SYSIMG_GREYTICK"
      gObjGroups(sCurrentGroup).SystemPermissions.Item("P_WORKFLOW_VIEWLOG").Allowed = True
    End If
  End If

  '
  ' ACCORD constraints
  '
  ' If a user group is granted permission to block or export transfers then
  ' they must have permission to view all transfers.
  'JPD 20041217 Fault 9647
  If IsModuleEnabled(modAccord) Then
    If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_BLOCK").Allowed _
      Or gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_SENDRECORD").Allowed _
      Or gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_VIEWARCHIVE").Allowed Then
      ' Grant permission to View all transfers.
      sstrvSystemPermissions.Nodes("P_ACCORD_VIEWTRANSFER").Image = "SYSIMG_GREYTICK"
      gObjGroups(sCurrentGroup).SystemPermissions.Item("P_ACCORD_VIEWTRANSFER").Allowed = True
    End If
  End If
  
  '
  ' EMAIL QUEUE constraints
  '
  ' If a user group is granted permission to rebuild/purge email queue entries then
  ' they must have permission to view the email queue entries.
  If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EMAIL_REBUILDPURGE").Allowed Then
    ' Grant permission to View email queue entries.
    sstrvSystemPermissions.Nodes("P_EMAIL_VIEW").Image = "SYSIMG_GREYTICK"
    gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EMAIL_VIEW").Allowed = True
  Else
    If gObjGroups(sCurrentGroup).SystemPermissions.Item("P_EMAIL_VIEW").Allowed Then
      sstrvSystemPermissions.Nodes("P_EMAIL_VIEW").Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
    End If
  End If

  '
  ' GENERAL DEFINITION constraints
  '
  ' If a user group is granted permission to create New definitions (eg. Data Transfer definition) then
  ' they must have permission to Edit and View the definitions.
  For Each objNode In sstrvSystemPermissions.Nodes
    If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
      iIndex = InStrRev(objNode.Key, "_")
      If iIndex > 0 Then
        sItemKey = Mid(objNode.Key, iIndex + 1)
        sCategoryKey = Left(objNode.Key, InStrRev(objNode.Key, "_"))
        
        If UCase$(sItemKey) = "NEW" Then
          sAssociatedItemKey = sCategoryKey & "EDIT"
          sAssociatedItemKey2 = sCategoryKey & "VIEW"
          
          If gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed Then
            ' Grant the associated permission items.
            gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey).Allowed = True
            sstrvSystemPermissions.Nodes(sAssociatedItemKey).Image = "SYSIMG_GREYTICK"
          
            gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey2).Allowed = True
            sstrvSystemPermissions.Nodes(sAssociatedItemKey2).Image = "SYSIMG_GREYTICK"
          Else
            ' Unlock the associated permissions.
            If gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey).Allowed Then
              sstrvSystemPermissions.Nodes(sAssociatedItemKey).Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
            End If
          
            If gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey2).Allowed Then
              sstrvSystemPermissions.Nodes(sAssociatedItemKey2).Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
            End If
          End If
        End If
      End If
    End If
  Next objNode
  Set objNode = Nothing

  ' If a user group is granted permission to create Edit definitions (eg. Data Transfer definition) then
  ' they must have permission to View the definitions.
  For Each objNode In sstrvSystemPermissions.Nodes
    If Left(objNode.Key, 1) = "P" Then                  ' The node is an Item node, rather than a Category node.
      iIndex = InStrRev(objNode.Key, "_")
      If iIndex > 0 Then
        sItemKey = Mid(objNode.Key, iIndex + 1)
        sCategoryKey = Left(objNode.Key, InStrRev(objNode.Key, "_"))
        
        If UCase$(sItemKey) = "EDIT" Then
          sAssociatedItemKey = sCategoryKey & "VIEW"
          
          If gObjGroups(sCurrentGroup).SystemPermissions.Item(objNode.Key).Allowed Then
            ' Grant the associated permission items.
            gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey).Allowed = True
            sstrvSystemPermissions.Nodes(sAssociatedItemKey).Image = "SYSIMG_GREYTICK"
          Else
            ' Unlock the associated permissions.
            If gObjGroups(sCurrentGroup).SystemPermissions.Item(sAssociatedItemKey).Allowed Then
              sstrvSystemPermissions.Nodes(sAssociatedItemKey).Image = IIf(mblnReadOnly, "SYSIMG_GREYTICK", "SYSIMG_TICK")
            End If
          End If
        End If
      End If
    End If
  Next objNode
  Set objNode = Nothing

End Sub

Private Sub InitialiseGroupCollections(psGroupName As String, bShowProgress As Boolean)
  ' Initialise the given group's collections.
  Dim fOK As Boolean
  
  fOK = True
  
  ' Do nothing if the group is already initialised.
  If Not gObjGroups(psGroupName).Initialised Then
  
    ' Initialise the group's table, view and system permission collections.
    If InitialiseGroup(gObjGroups(psGroupName), bShowProgress) Then
    
      ' Add the table and view nodes to the console treeview.
      fOK = trvConsole_LoadTablesViews(trvConsole.Nodes("TV_GP_" & psGroupName))
    End If
  End If

End Sub

Private Sub Group_Properties()

  Dim objNewGroup As SecurityGroup
  Dim objNode As SSActiveTreeView.SSNode
  Dim objUser As SecurityUser
  Dim vArray As Variant
  Dim sCurrentGroup As String
  Dim sNewGroup As String
  Dim fAccessSet As Boolean
  
  ' Get group name depending on what object the user has selected
  If ActiveView Is trvConsole Then
    sCurrentGroup = Trim$(WhichGroup(trvConsole.SelectedItem))
  ElseIf ActiveView Is lvList Then
    sCurrentGroup = lvList.SelectedItem.Text
  End If

  Load frmNewGroup
  With frmNewGroup
    .Initialise GROUPACTION_EDIT
    
    Do While .Tag <> "Cancel"
      
      ' Display the new group form.
      .GroupName = sCurrentGroup
      .Show vbModal
      sNewGroup = Trim$(.GroupName)

      ' Check whether the user select OK or cancel.
      If .Tag = "OK" Then
        If sNewGroup <> sCurrentGroup Then

          If Group_Valid(sNewGroup) Then
  
            ' Ensure that the group is initialised
            If Not gObjGroups(sCurrentGroup).Initialised Then
              InitialiseGroup gObjGroups(sCurrentGroup), True
            End If
  
            ' Initialise the user group if necessary.
            If Not gObjGroups(sCurrentGroup).Users_Initialised Then
              InitialiseUsersCollection gObjGroups(sCurrentGroup)
            End If
            
            'Remove the old group from the tree view and add the new group
            trvConsole.Nodes.Remove "GP_" & sCurrentGroup
            AddGroup gObjGroups, sNewGroup
            
            gObjGroups(sNewGroup).Initialised = gObjGroups(sCurrentGroup).Initialised
            gObjGroups(sNewGroup).Users_Initialised = gObjGroups(sCurrentGroup).Users_Initialised
            
            'Copy the existing group properties
            Set gObjGroups(sNewGroup).Tables = gObjGroups(sCurrentGroup).Tables.Clone
            Set gObjGroups(sNewGroup).Views = gObjGroups(sCurrentGroup).Views.Clone
            Set gObjGroups(sNewGroup).SystemPermissions = gObjGroups(sCurrentGroup).SystemPermissions.Clone
            gObjGroups(sNewGroup).OriginalName = gObjGroups(sCurrentGroup).Name
            gObjGroups(sNewGroup).Changed = True

            ' NPG20080611 Fault 13357
            ' Set the copy flag to true so that existing security is retained for renamed security groups
            gObjGroups(sNewGroup).CopyGroup = sCurrentGroup

            ' Copy/set the utility/report access if required.
            fAccessSet = False
            If Len(gObjGroups(sCurrentGroup).AccessCopyGroup) > 0 Then
              fAccessSet = True
              gObjGroups(sNewGroup).AccessCopyGroup = gObjGroups(sCurrentGroup).AccessCopyGroup
            Else
              vArray = gObjGroups(sCurrentGroup).AccessConfiguration
              
              If IsArray(vArray) Then
                If UBound(vArray, 2) > 0 Then
                  fAccessSet = True
                  gObjGroups(sNewGroup).AccessConfiguration = gObjGroups(sCurrentGroup).AccessConfiguration
                End If
              End If
            End If
            
            If Not fAccessSet Then
              gObjGroups(sNewGroup).AccessCopyGroup = sCurrentGroup
            End If

            'Move all of the existing users to the new group
            For Each objUser In gObjGroups(sCurrentGroup).Users
              Call MoveUserToNewGroup(objUser.Login, sNewGroup, sCurrentGroup)
            Next
            Set objUser = Nothing
            
            Call AddGroupToTreeView(sNewGroup)
  
            ' Add the table and view nodes to the console treeview.
            trvConsole_LoadTablesViews trvConsole.Nodes("TV_GP_" & sNewGroup)
            
            'Delete the old group
            Call DeleteGroup(sCurrentGroup)
            
            Application.Changed = True
          End If
        End If
        
        .Tag = "Cancel"
      End If

    Loop
  End With
  
  ' Unload the new group form
  Unload frmNewGroup

End Sub


Private Function Group_Valid(sGroupName As String) As Boolean

  Dim fValuesOK As Boolean
  
  fValuesOK = True
        
  ' Check that the new group name is acceptable
  ' Check if the user login is a reserved login.
  If fValuesOK And (UCase$(sGroupName) = "SA") Then
    MsgBox "'" & sGroupName & "' is a reserved system login." & vbNewLine & _
           "Please enter another name.", vbInformation + vbOKOnly, App.Title
    fValuesOK = False
  End If
    
  ' Check first for keyword
  If fValuesOK And Database.IsKeyword(sGroupName) Then
    MsgBox "'" & sGroupName & "' is a reserved word." & vbNewLine & _
            "Please enter another name.", vbInformation + vbOKOnly, App.Title
    fValuesOK = False
  End If
    
  ' Check for it being a group or user name already in use.
  If fValuesOK And (IsUserNameInUse(sGroupName, gObjGroups)) Then
    MsgBox "'" & sGroupName & "' is already used as the name for a user group or user." & vbNewLine & _
           "Please enter another name.", vbInformation + vbOKOnly, App.Title
    fValuesOK = False
  End If

  Group_Valid = fValuesOK

End Function


Private Sub AddGroupToTreeView(sGroupName As String)
          
  Dim nodGroup As SSNode
  Dim nodUsers As SSNode
  Dim nodSysPriv As SSNode
  Dim nodTableViews As SSNode
          
  ' Add the group to the tree view with underlying folders

  Set nodGroup = trvConsole.Nodes.Add("RT", tvwChild, _
                "GP_" & sGroupName, sGroupName, "GROUP", , "GROUP")
  nodGroup.Sorted = False
      
  ' Add the Users Folder node.
  Set nodUsers = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "US_" & nodGroup.Key, "User Logins", "CLOSEDFOLDER", "OPENFOLDER", "USERS")
  With nodUsers
    .ExpandedImage = "OPENFOLDER"
    .Sorted = True
  End With
          
  ' Add the Tables/Views folder node.
  
  'MH20010208 Fault 1825
  'Set nodTableViews = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "TV_" & nodGroup.Key, "Tables / Views", "CLOSEDFOLDER", , "TABLESVIEWS")
  Set nodTableViews = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "TV_" & nodGroup.Key, "Data Permissions", "CLOSEDFOLDER", "OPENFOLDER", "TABLESVIEWS")
  With nodTableViews
    .ExpandedImage = "OPENFOLDER"
    .Sorted = True
  End With
          
  ' Add the System Permissions node.
  Set nodSysPriv = trvConsole.Nodes.Add(nodGroup.Key, tvwChild, "SY_" & nodGroup.Key, "System Permissions", "SYSTEM", , "SYSTEM")
  nodSysPriv.Sorted = True
          
  ' Select the new group node.
  'Make sure new table is visible and selected in the treeview.
  With nodGroup
    .EnsureVisible
    .Expanded = True
  End With
  trvConsole.SelectedItem = nodGroup
  
  'TM10082004 8430 - Refresh the tree and populate with the tables and columns etc.
  trvConsole_LoadTablesViews trvConsole.Nodes("TV_GP_" & sGroupName)
  
  ' Release the nodes
  Set nodGroup = Nothing
  Set nodUsers = Nothing
  Set nodTableViews = Nothing
  Set nodSysPriv = Nothing

End Sub


Private Sub DeleteGroup(sGroupName As String)

  ' See if the group exists in SQL Server
  If gObjGroups(sGroupName).NewGroup Then
    ' The group has not yet been added to sql server so just remove the item from the tree
    ' and the security groups collection
    gObjGroups.Remove (sGroupName)
  Else
    ' The group does exist in sql server so mark it for deletion and remove it from the tree
    With gObjGroups(sGroupName)
      .RequireLogout = True     'MH20010410
      .DeleteGroup = True
      .Changed = True
    End With
  End If

End Sub

Private Sub EditMenu_AutomaticAdd()

  ' Adds multiple users to the selected security group
  Dim sCurrentGroupName As String
  Dim bExit As Boolean
  Dim bOK As Boolean
  Dim frmNewUsers As frmNewMultipleUser
  Dim frmNewUsersReport As frmNewMultipleUserReport
  Dim avReportList() As Variant            ' 2D - (Username, Password, status)
  Dim alngSourceTables() As Long
  Dim rsRecords As ADODB.Recordset
  Dim iCreateUser As CreateUserStatus
  Dim strUserName As String
  Dim strPassword As String
  Dim strPersonName As String
  Dim lngRetryCount As Long
  Dim iUsersCreated As Integer
  Dim strCurrentLogin As String
  Dim bBypassPolicy As Boolean
  
  Dim iCreateMode As SecurityMgr.CreateUserMode
  Dim bForcePasswordChange As Boolean
  Dim bCheckPolicy As Boolean
  Dim lngFilterExprID As Long
  Dim lngSQLUserNameExprID As Long
  Dim lngWindowsUserNameExprID As Long
  Dim lngPasswordExprID As Long
  
  Dim oUserNames As clsExprExpression
  Dim oPasswords As clsExprExpression
  Dim oFilter As clsExprExpression

  Dim strPersonnelTable As String
  Dim strCheckLoginFieldCode As String
  Dim strPersonNameCode As String
 
  Dim strUserDomainName As String
  Dim strWindowsUserNameCode As String
  Dim strUserNameCode As String
  Dim strPasswordCode As String
  Dim strFilterCode As String
  Dim strSQL As String
  Dim bApplicationChanged As Boolean
  
  Dim objBackupGroup As SecurityGroup
  

  On Error GoTo ErrorTrap

  ReDim mvarUDFsRequired(0)
  
  ' Store change status in case they cancel auto user add
  bApplicationChanged = Application.Changed

  ' JDM - 13/11/01 - Fault 3137 - Need to check if login field has been defiled
  If glngLoginColumnID = 0 Then
    MsgBox "The Personnel login field has not been defined.", vbExclamation, Application.Name
    Exit Sub
  End If

  ' Get the name of the currently selected group.
  sCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  strUserNameCode = vbNullString
  strPasswordCode = vbNullString
  strFilterCode = vbNullString
  bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)   ' Default - Off
  bExit = False

  ' Load automatic add form
  Set frmNewUsers = New frmNewMultipleUser
  frmNewUsers.SecurityGroupName = sCurrentGroupName
  frmNewUsers.Show vbModal
  
  ' If cancelled lets get outta here...
  If frmNewUsers.Cancelled Then
    bExit = True
  Else
    iCreateMode = frmNewUsers.CreateUserMode
    bForcePasswordChange = frmNewUsers.ForceChangePassword
    lngFilterExprID = frmNewUsers.ID_FilterExpr
    lngSQLUserNameExprID = frmNewUsers.ID_SQLUserNameExpr
    lngPasswordExprID = frmNewUsers.ID_PasswordExpr
    lngWindowsUserNameExprID = frmNewUsers.ID_WindowsUserNameExpr
    strUserDomainName = frmNewUsers.DomainName
  End If

  ' Tidy up the form
  Set frmNewUsers = Nothing
  
  If Not bExit Then
    
    With gobjProgress
      '.AviFile = vbNullString
      .AVI = dbAutoAdd
      .NumberOfBars = 1
      .Bar1MaxValue = 5
      .Caption = "Processing..."
      .MainCaption = "Add Users"
      .Cancel = False
      .Time = False
      .HidePercentages = False
      .OpenProgress
    End With
    
    ' Initialise the user group if necessary.
    If Not gObjGroups(sCurrentGroupName).Users_Initialised Then
      InitialiseUsersCollection gObjGroups(sCurrentGroupName)
    End If
    
    ' Generate code for the SQL usernames
    If iCreateMode = iUSERCREATE_SQLLOGIN Then
      ReDim alngSourceTables(2, 0)
      Set oUserNames = New clsExprExpression
      bOK = oUserNames.Initialise(glngPersonnelTableID, lngSQLUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER)
      If bOK Then
        bOK = oUserNames.RuntimeCalculationCode(alngSourceTables, strUserNameCode, True)
     
        ' Load any required UDFs
        If bOK And gbEnableUDFFunctions Then
          bOK = oUserNames.UDFCalculationCode(alngSourceTables, mvarUDFsRequired(), True)
        End If
      
      End If
      Set oUserNames = Nothing
      gobjProgress.UpdateProgress False
    
      ' Generate code for the passwords
      ReDim alngSourceTables(2, 0)
      Set oPasswords = New clsExprExpression
      bOK = oPasswords.Initialise(glngPersonnelTableID, lngPasswordExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER)
      If bOK Then
        bOK = oPasswords.RuntimeCalculationCode(alngSourceTables, strPasswordCode, True)
      
        ' Load any required UDFs
        If bOK And gbEnableUDFFunctions Then
          bOK = oPasswords.UDFCalculationCode(alngSourceTables, mvarUDFsRequired(), True)
        End If
      
      End If
      Set oPasswords = Nothing
      gobjProgress.UpdateProgress False
    
    ' Windows accounts (manual)
    ElseIf iCreateMode = iUSERCREATE_WINDOWSMANUAL Then
    
      ReDim alngSourceTables(2, 0)
      Set oUserNames = New clsExprExpression
      bOK = oUserNames.Initialise(glngPersonnelTableID, lngWindowsUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER)
      If bOK Then
        bOK = oUserNames.RuntimeCalculationCode(alngSourceTables, strUserNameCode, True)
     
        ' Load any required UDFs
        If bOK And gbEnableUDFFunctions Then
          bOK = oUserNames.UDFCalculationCode(alngSourceTables, mvarUDFsRequired(), True)
        End If
      
        ' Bolt on the domain string
        If Len(strUserDomainName) > 0 Then
          strUserNameCode = "'" & strUserDomainName & "\' + " & strUserNameCode
        End If
      
      End If
      Set oUserNames = Nothing
      gobjProgress.UpdateProgress False
    
      ' Integrated security has no need to define passwords
      strPasswordCode = vbNullString
    
    End If
    
    'NHRD08072003 Fault 3826 Added the ability to ignore the fact
    'that a filter has not been selected.
    'Before it was compulsory also see Refeshbuttons in frmNewMultipleUser
    If lngFilterExprID > 0 Then
      ' Generate code for the filter
      ReDim alngSourceTables(2, 0)
      Set oFilter = New clsExprExpression
      bOK = oFilter.Initialise(glngPersonnelTableID, lngFilterExprID, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
      If bOK Then
        bOK = oFilter.RuntimeCalculationCode(alngSourceTables, strFilterCode, True)
        
        ' Load any required UDFs
        If bOK And gbEnableUDFFunctions Then
          bOK = oFilter.UDFCalculationCode(alngSourceTables, mvarUDFsRequired(), True)
        End If
        
        If LenB(strFilterCode) <> 0 Then
          strFilterCode = " WHERE (" & strFilterCode & " = 1)"
        End If
      End If
      Set oFilter = Nothing
    End If
    
    gobjProgress.UpdateProgress False

    ' Generate code to filter out existing users
    strPersonnelTable = modExpression.GetTableName(glngPersonnelTableID)
    strCheckLoginFieldCode = strPersonnelTable & "." & modExpression.GetColumnName(glngLoginColumnID)
    strPersonNameCode = strPersonnelTable & "." & modExpression.GetColumnName(glngForenameColumnID) _
          & " + ' ' + " & strPersonnelTable & "." & modExpression.GetColumnName(glngSurnameColumnID)

    ' Build array of all users to be added and a password
    '    strSQL = "Select ID, " & strPersonNameCode & " As PersonName," _
    '      & strCheckLoginFieldCode & " As CurrentLogin, " _
    '      & strUserNameCode & " As UserName, " _
    '      & strPasswordCode & " As Password" _
    '      & " FROM " & strPersonnelTable _
    '      & strFilterCode
    
    ' Create dynamic User defined functions
    UDFFunctions mvarUDFsRequired, True
    
    ' Build array of all users to be added and a password
    strSQL = "Select ID, " & strPersonNameCode & " As PersonName," _
      & strCheckLoginFieldCode & " As CurrentLogin, " _
      & strUserNameCode & " As UserName, " _
      & IIf(LenB(strPasswordCode) <> 0, strPasswordCode, "''") & " As Password" _
      & " FROM " & strPersonnelTable _
      & strFilterCode _
      & " ORDER BY UserName "
      '& " ORDER by " & strUserNameCode
    
    Set rsRecords = modExpression.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)
    gobjProgress.UpdateProgress False
   
    ReDim avReportList(rsRecords.RecordCount, 4)
    iUsersCreated = 0
    
    ' Open in forward only for rest of code
    rsRecords.Close
    Set rsRecords = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    
    ' Create copy of the current security group
    Set objBackupGroup = gObjGroups(sCurrentGroupName).Clone(True)
    
    ' Loop creating users
    
    With rsRecords
    
      If Not .EOF And Not .BOF Then
        .MoveFirst
        Do While Not .EOF
          
          ' JDM - 06/12/01 - Stop any apostrophees being used in username
          ' JDM - 12/12/01 - Strip username down to 48 because that's the max length in the audit log
          'JPD 20030801 Fault 6345
          If IsNull(.Fields("UserName").Value) Then
            strUserName = vbNullString
          Else
            'MH20051007 Fault 10396
            'strUserName = Left(Replace(Trim(.Fields("UserName")), "'", vbnullstring), 48)
            strUserName = Left(Trim(.Fields("UserName").Value), 48)
          End If
          
          
          If IsNull(.Fields("Password").Value) Then
            strPassword = vbNullString
          Else
            'MH20051007 Fault 10396
            'strPassword = Replace(.Fields("Password"), "'", vbnullstring)
            strPassword = .Fields("Password").Value
          End If
          
          'JDM - 23/04/02 - Fault 3794 - Apostophees in passwords causing errors
          strCurrentLogin = IIf(IsNull(.Fields("CurrentLogin").Value), vbNullString, .Fields("CurrentLogin").Value)
          
          ' Don't add if PR record already has a login name
          If LenB(strCurrentLogin) = 0 And (Not IsAlreadyNewUser(strUserName, gObjGroups)) Then
            lngRetryCount = 0
            iCreateUser = CreateUser(strUserName, strPassword, sCurrentGroupName, .Fields("ID").Value, bForcePasswordChange, iCreateMode, bBypassPolicy)
           
            If iCreateMode = iUSERCREATE_SQLLOGIN Then
              Do While iCreateUser = iFAILED_USERNAMEISUSED
                lngRetryCount = lngRetryCount + 1
                
                'MH20051007 Fault 10396
                ''' JDM - 06/12/01 - Stop any apostrophees being used in username
                '''strUserName = Trim(.Fields("UserName")) & LTrim(Str(lngRetryCount))
                ''strUserName = Left(Replace(Trim(.Fields("UserName")), "'", vbnullstring), 48) & LTrim(Str(lngRetryCount))
                strUserName = Left$(strUserName, 48) & LTrim(Str(lngRetryCount))
                
                iCreateUser = CreateUser(strUserName, strPassword, sCurrentGroupName, .Fields("ID").Value, bForcePasswordChange, iCreateMode, bBypassPolicy)
              Loop
            End If
          Else
            strUserName = IIf(LenB(strCurrentLogin) = 0, strUserName, strCurrentLogin)
            iCreateUser = iSUCCESS_USERALREADYADDED
          End If
        
          ' Add to reporting array
          avReportList(iUsersCreated, 1) = strUserName
          avReportList(iUsersCreated, 2) = strPassword
          avReportList(iUsersCreated, 3) = iCreateUser
          avReportList(iUsersCreated, 4) = .Fields("PersonName")
          iUsersCreated = iUsersCreated + 1
          rsRecords.MoveNext

        Loop
      End If
  
      .Close
    End With
    Set rsRecords = Nothing
    
    gobjProgress.UpdateProgress False
    gobjProgress.CloseProgress

    ' JDM - 23/04/2004 - Fault 8542 - Clear up UDFs
    UDFFunctions mvarUDFsRequired, False

    ' Show report of users created
    If iUsersCreated > 0 Then
      Set frmNewUsersReport = New frmNewMultipleUserReport
      frmNewUsersReport.CreateMode = iCreateMode
      frmNewUsersReport.SecurityGroupName = sCurrentGroupName
      frmNewUsersReport.ReportData = avReportList
      frmNewUsersReport.Show vbModal
      
      If frmNewUsersReport.Cancelled Then
        gObjGroups(sCurrentGroupName).Replace objBackupGroup
        Application.Changed = bApplicationChanged
      End If
      
      Set frmNewUsersReport = Nothing
    Else
      MsgBox "No users found meeting selected criteria.", vbInformation, "Automatic User Add"
    End If
    
    ' Refresh the display
    UpdateRightPane
    
  End If
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  gobjProgress.CloseProgress
  MsgBox "An Error has occured while adding users", vbCritical, Application.Name
  Resume TidyUpAndExit

End Sub

Private Function CreateUser(pstrUsername As String, pstrPassword As String _
    , pstrGroup As String, plngPersonnelRecordID As Long, pbForcePasswordChange As Boolean _
    , piCreateMode As SecurityMgr.CreateUserMode, pbBypassPolicy As Boolean) As CreateUserStatus

  ' Add a new user to the specified group.
  Dim strPassword As String
  Dim strCurrentGroupName As String
  Dim sSQL As String
  Dim strUserLogin As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim bOK As Boolean
  Dim bUserFound As Boolean
  
  ' Get the name of the currently selected group.
  strCurrentGroupName = pstrGroup
  strUserLogin = pstrUsername
  strPassword = pstrPassword
  bOK = True
  
  ' Initialise the user group if necessary.
  If Not gObjGroups(strCurrentGroupName).Users_Initialised Then
    InitialiseUsersCollection gObjGroups(strCurrentGroupName)
  End If
  
  ' Check that a user name has been entered.
  If strUserLogin = vbNullString Then
    CreateUser = iFAILED_USERNAMEISBLANK
    bOK = False
  End If

  ' Check for it being a group or user name already in use.
  If bOK And IsUserNameInUse(strUserLogin, gObjGroups) Then
    CreateUser = iFAILED_USERNAMEISUSED
    bOK = False
  End If
  
  'TM20020122 Fault 3374
  ' Check that the username is less or equal to the length of the Login field.
  If bOK And Not (GetColumnSize(glngLoginColumnID) >= Len(pstrUsername)) Then
    CreateUser = iFAILED_USERNAMEGREATERTHANLOGINSIZE
    bOK = False
  End If

  'TM20020122 Fault 3229
  'Check that the username is less than the maximum username length.
  If bOK And Len(Trim(strUserLogin)) > giMAXIMUMUSERNAMELENGTH Then
    CreateUser = iFAILED_USERNAMEISTOOLONG
    bOK = False
  End If

  ' Checks specific to SQL authentications
  If piCreateMode = iUSERCREATE_SQLLOGIN Then

    ' Check if the user login is a reserved login.
    If bOK And (UCase$(strUserLogin) = "SA") Then
      CreateUser = iFAILED_USERNAMEISRESERVED
      bOK = False
    End If
      
    ' Check if the user login is an SQL keyword.
    If bOK And Database.IsKeyword(strUserLogin) Then
      CreateUser = iFAILED_USERNAMEISKEYWORD
      bOK = False
    End If
    
    ' Check if the username is numeric
    If bOK And IsNumeric(strUserLogin) Then
      CreateUser = iFAILED_USERNAMEISNUMERIC
      bOK = False
    End If
     
    ' Check that login name of same name does not exist.
    If bOK And IsSQLLoginNameInUse(strUserLogin) Then
      CreateUser = iWARNING_LOGINEXISTS
      bOK = True
    End If
     
    'TM20020122 Fault 3350
    'Check that the password is greater than the minimum password length.
    If Not pbBypassPolicy Then
      If bOK And Not IsMinimumPasswordLength(strPassword) Then
        CreateUser = iFAILED_PASSWORDNOTMINIMUM
        bOK = False
      End If
      
      If bOK Then
        If Not CheckPasswordComplexity(strUserLogin, strPassword) Then
          CreateUser = iFAILED_PASSWORDNOTCOMPLEX
          bOK = False
        End If
      End If
    End If
    
  End If

  ' Checks specific to windows authentication
  If bOK = True And piCreateMode = iUSERCREATE_WINDOWSMANUAL Then

    If bOK And Not CheckNTAccountExist(strUserLogin) Then
      CreateUser = iFAILED_NTACCOUNTNOTEXIST
      bOK = False
    End If
  
  End If
  
  If bOK Then
    ' Check if the user already exists in another group but is marked as deleted.
    bUserFound = False
    
    For Each objGroup In gObjGroups
      If Not objGroup.Users_Initialised Then
        InitialiseUsersCollection objGroup
      End If
      
      For Each objUser In objGroup.Users
        If UCase$(strUserLogin) = UCase$(objUser.Login) Then
          ' Mark the user as undeleted and move it to the new group.
          objUser.DeleteUser = False
          MoveUserToNewGroup strUserLogin, strCurrentGroupName, objGroup.Name
          bUserFound = True
          Exit For
        End If
      Next objUser
      Set objUser = Nothing
      
      If bUserFound Then
        Exit For
      End If
    Next objGroup
    Set objGroup = Nothing
    
    ' Add the user to the group.
    If Not bUserFound Then
       ' AE20080425 Fault #12827
'      gObjGroups(strCurrentGroupName).Users.Add strUserLogin, True, False, True, _
'        vbNullString, vbNullString, strUserLogin, strPassword, , plngPersonnelRecordID, pbForcePasswordChange, _
'        IIf(piCreateMode = iUSERCREATE_SQLLOGIN, iUSERTYPE_SQLLOGIN, iUSERTYPE_TRUSTEDUSER), Not pbBypassPolicy

      gObjGroups(strCurrentGroupName).Users.Add strUserLogin, True, False, True, _
        vbNullString, vbNullString, strUserLogin, strPassword, , plngPersonnelRecordID, pbForcePasswordChange, _
        IIf(piCreateMode = iUSERCREATE_SQLLOGIN, iUSERTYPE_SQLLOGIN, iUSERTYPE_TRUSTEDUSER), Not pbBypassPolicy
    End If
    
    ' Enable the apply button
    Application.Changed = True
    CreateUser = IIf(CreateUser = iWARNING_LOGINEXISTS, iWARNING_LOGINEXISTS, iSUCCESS)
    
  End If

End Function

Private Sub EditMenu_CopyGroup()

  ' Copy a group.
  Dim fValuesOK As Boolean
  Dim sGroupName As String
  Dim sCurrentGroupName As String
  Dim fAccessSet As Boolean
  Dim vArray As Variant
  
  ' Get group name depending on what object the user has selected
  If ActiveView Is trvConsole Then
    sCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  ElseIf ActiveView Is lvList Then
    sCurrentGroupName = lvList.SelectedItem.Text
  End If

  ' Initialise the user group if necessary.
  If Not gObjGroups(sCurrentGroupName).Initialised Then
    InitialiseGroupCollections sCurrentGroupName, True
  End If

  Load frmNewGroup
  With frmNewGroup
    .Initialise GROUPACTION_COPY
    
    Do While .Tag <> "Cancel"

      ' Display the new group form.
      .GroupName = "Copy_of_" & gObjGroups(sCurrentGroupName).Name
      .CopyGroup
      .Show vbModal

      ' Check whether the user select OK or cancel.
      If .Tag = "OK" Then

        sGroupName = Trim(.GroupName)
        fValuesOK = Group_Valid(sGroupName)

        If fValuesOK Then

          ' Paste clone of current security group
          gObjGroups.Paste gObjGroups(sCurrentGroupName).Clone(False, sGroupName)

          ' Copy/set the utility/report access if required.
          fAccessSet = False
          If Len(gObjGroups(sCurrentGroupName).AccessCopyGroup) > 0 Then
            fAccessSet = True
            gObjGroups(sGroupName).AccessCopyGroup = gObjGroups(sCurrentGroupName).AccessCopyGroup
          Else
            vArray = gObjGroups(sCurrentGroupName).AccessConfiguration
            
            If IsArray(vArray) Then
              If UBound(vArray, 2) > 0 Then
                fAccessSet = True
                gObjGroups(sGroupName).AccessConfiguration = gObjGroups(sCurrentGroupName).AccessConfiguration
              End If
            End If
          End If
          
          If Not fAccessSet Then
            gObjGroups(sGroupName).AccessCopyGroup = sCurrentGroupName
          End If

          'JPD 20071203 Faults 12580, 12670, 12671
          If Len(gObjGroups(sCurrentGroupName).CopyGroup) > 0 Then
            gObjGroups(sGroupName).CopyGroup = gObjGroups(sCurrentGroupName).CopyGroup
          Else
            If gObjGroups(sCurrentGroupName).Changed _
              And gObjGroups(sCurrentGroupName).NewGroup Then
              
              gObjGroups(sGroupName).CopyGroup = ""
            Else
              gObjGroups(sGroupName).CopyGroup = sCurrentGroupName
            End If
          End If

          Call AddGroupToTreeView(sGroupName)

          ' Enable the apply button
          Application.Changed = True
          .Tag = "Cancel"
          
        End If
      End If
    Loop
  End With

  ' Unload the new group form
  Unload frmNewGroup

End Sub

Private Sub EditMenu_Move()

  ' Move item to the list view depending on the active view,
  ' and type of object selected in the treeview.
  If ActiveView Is lvList Then
  
    Select Case trvConsole.SelectedItem.DataKey
      Case "USERS"
        User_Move
    End Select
  
  End If

End Sub

Private Function User_Move()

  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim asSelectedUsers() As Variant
  Dim strCurrentGroupName As String
  Dim strMoveToGroupName As String
  Dim frmMove As frmMoveUser
  Dim objUser As SecurityUser
  Dim bExit As Boolean
  Dim lngSelectedCount As Long
  
  bExit = False
  lngSelectedCount = 0
  
  ' Get the name of the currently selected group.
  strCurrentGroupName = WhichGroup(trvConsole.SelectedItem)
  
  ' Initialise the user group if necessary.
  If Not gObjGroups(strCurrentGroupName).Users_Initialised Then
    InitialiseUsersCollection gObjGroups(strCurrentGroupName)
  End If
  
  ' If we have more than one selection then question the multi-move.
  ReDim asSelectedUsers(0)
  If (lvList_SelectedCount > 1) Then
    ' Read the names of the users to be moved from the listview into an array.
    For iLoop = 1 To lvList.ListItems.Count
      If lvList.ListItems(iLoop).Selected = True Then
        iNextIndex = UBound(asSelectedUsers) + 1
        ReDim Preserve asSelectedUsers(iNextIndex)
        asSelectedUsers(iNextIndex) = lvList.ListItems(iLoop).Text
      End If
    Next iLoop
  Else
    ' Read the name of the group to be moved from the treeview into an array.
    ReDim asSelectedUsers(1)
    asSelectedUsers(1) = lvList.SelectedItem.Text
  End If
  
  lngSelectedCount = UBound(asSelectedUsers)
  
  ' get which group to move to.
  Set frmMove = New frmMoveUser
  With frmMove
    .MoveToGroupName = strCurrentGroupName
    .Show vbModal
    If Not .Cancelled Then
      strMoveToGroupName = .MoveToGroupName
    Else
      bExit = True
    End If
  End With
  Set frmMove = Nothing

'********************************************************************************
  'TM20020122 Fault 3353 - Check if the user is logged into the system.
  Dim sCurrentUser As String
  Dim iFailedCount As Integer
  Dim sMessage As String
  Dim iAnswer As Integer
  
  If Not bExit And Not strMoveToGroupName = strCurrentGroupName Then
    iFailedCount = 0
    sCurrentUser = vbNullString
    sMessage = vbNullString
  
    For iLoop = 1 To UBound(asSelectedUsers)
      sCurrentUser = CStr(gObjGroups(strCurrentGroupName).Users(asSelectedUsers(iLoop)).UserName)
      'MH20061017 Fault 11376
      'If UserLoggedIn(sCurrentUser) <> vbNullString Then
      If GetCurrentUsersAppName(sCurrentUser) <> vbNullString Then
        asSelectedUsers(iLoop) = vbNullString
        iFailedCount = iFailedCount + 1
        sMessage = sMessage & vbNewLine & sCurrentUser
      End If
    Next iLoop
  
    If iFailedCount > 0 Then
      
      If lngSelectedCount > 1 And (lngSelectedCount > iFailedCount) Then
        sMessage = sMessage & vbNewLine & vbNewLine & _
                  "Do you wish to move the selected users that are not logged in to the system?"
      End If
      
      If iFailedCount = 1 Then
        iAnswer = MsgBox("Cannot move the following user as the user is logged into the system." & _
                          vbNewLine & sMessage _
                          , IIf((lngSelectedCount > iFailedCount), vbExclamation + vbYesNo, vbExclamation + vbOKOnly), App.Title)
      Else
        iAnswer = MsgBox("Cannot move the following users as the users are logged into the system." & _
                          vbNewLine & sMessage _
                          , IIf((lngSelectedCount > iFailedCount), vbExclamation + vbYesNo, vbExclamation + vbOKOnly), App.Title)
      End If
      
      Select Case iAnswer
      Case vbYes
        bExit = False
      Case vbNo:
        bExit = True
      Case vbOK:
        'NHRD22062006 11066 Added the vbOK option to
        'cover all eventualities from the msg above
        bExit = True
      End Select
    End If
  End If
'********************************************************************************

  If Not bExit And Not strMoveToGroupName = strCurrentGroupName Then

    ' Initialise the new user group if necessary.
    If Not gObjGroups(strMoveToGroupName).Initialised Then
      InitialiseGroup gObjGroups(strMoveToGroupName), False
    End If

    'JPD 20030728 Fault 6304
    If Not gObjGroups(strMoveToGroupName).Users_Initialised Then
      InitialiseUsersCollection gObjGroups(strMoveToGroupName)
    End If

    ' Move the selected users
    For iLoop = 1 To UBound(asSelectedUsers)
      If asSelectedUsers(iLoop) <> vbNullString Then
        ' Show this user in new group
        Set objUser = gObjGroups(strCurrentGroupName).Users(asSelectedUsers(iLoop)).Clone
      
        ' Is User already defined in selected group
        If gObjGroups(strMoveToGroupName).Users.Paste(objUser) Then
          gObjGroups(strMoveToGroupName).Users(asSelectedUsers(iLoop)).MovedUserFrom = strCurrentGroupName
        Else
          gObjGroups(strMoveToGroupName).Users(asSelectedUsers(iLoop)).MovedUserFrom = vbNullString
          gObjGroups(strMoveToGroupName).Users(asSelectedUsers(iLoop)).MovedUserTo = vbNullString
        End If
        
        ' Move the current user
        If Not gObjGroups(strCurrentGroupName).Users(asSelectedUsers(iLoop)).NewUser Then
          gObjGroups(strCurrentGroupName).Users(asSelectedUsers(iLoop)).MovedUserTo = strMoveToGroupName
          gObjGroups(strCurrentGroupName).Users(asSelectedUsers(iLoop)).Changed = True
        Else
          gObjGroups(strCurrentGroupName).Users.Remove (asSelectedUsers(iLoop))
          gObjGroups(strMoveToGroupName).Users(asSelectedUsers(iLoop)).Changed = True
        End If
      End If
    Next iLoop
    ' Update display
    Application.Changed = True
    UpdateRightPane
  End If

End Function

Private Sub ResetHelpContextID()
  'Reset the HelpContextID so that the correct help fpage is called up when
  'F1 is pressed
  Select Case trvConsole.SelectedItem.DataKey
     Case "SYSTEM"
        Me.HelpContextID = 8049
      
      Case "GROUPS"
        Me.HelpContextID = 8050
      
      Case "GROUP"
        Me.HelpContextID = 8051
      
      Case "USERS"
        Me.HelpContextID = 8052
      
      Case "TABLESVIEWS"
        Me.HelpContextID = 8053

  End Select

End Sub

' Locates a user and opens up the group for that user
Private Sub EditMenu_FindUser()

  Dim objForm As New frmFindUser
  Dim strFoundGroup As String
  Dim iCount As Integer
  Dim iIndex As Integer
  Dim fOK As Boolean
  
  fOK = objForm.Initialise
  
  If fOK Then
  
    objForm.Show vbModal
  
    If Not objForm.Cancelled Then
      IsUserNameInUse objForm.SelectedUser, gObjGroups, strFoundGroup
  
      ' Find the group in the treeview
      trvConsole.SelectedNodes.Clear
      For iCount = 1 To trvConsole.Nodes.Count
        If LCase$(trvConsole.Nodes(iCount).Text) = LCase$(strFoundGroup) Then
          trvConsole.Nodes(iCount).Selected = True
          trvConsole_NodeClick trvConsole.SelectedItem
          trvConsole.SelectedItem.Expanded = True
          trvConsole.SelectedItem.Child.FirstSibling.Selected = True
          trvConsole_NodeClick trvConsole.SelectedItem
        End If
      Next iCount
  
    End If
  End If

  Unload objForm

End Sub
