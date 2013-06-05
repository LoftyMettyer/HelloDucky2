VERSION 5.00
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "coa_line.ocx"
Begin VB.Form frmAuditSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audit Log"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5082
   Icon            =   "frmAuditSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2790
      TabIndex        =   9
      Top             =   5265
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4095
      TabIndex        =   10
      Top             =   5265
      Width           =   1200
   End
   Begin VB.Frame fraAudit 
      Caption         =   "Audit :"
      Height          =   5100
      Left            =   135
      TabIndex        =   11
      Top             =   45
      Width           =   5160
      Begin VB.ComboBox cboID 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4605
         Width           =   2500
      End
      Begin VB.ComboBox cboDescription 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4200
         Width           =   2500
      End
      Begin VB.ComboBox cboTime 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1340
         Width           =   2500
      End
      Begin VB.ComboBox cboUser 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1755
         Width           =   2500
      End
      Begin VB.ComboBox cboAuditDate 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   915
         Width           =   2500
      End
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   2500
      End
      Begin VB.ComboBox cboColumn 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2565
         Width           =   2500
      End
      Begin VB.ComboBox cboAuditTable 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2500
      End
      Begin VB.ComboBox cboNewValue 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3390
         Width           =   2500
      End
      Begin VB.ComboBox cboModule 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3795
         Width           =   2500
      End
      Begin VB.ComboBox cboOldValue 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2985
         Width           =   2500
      End
      Begin COALine.COA_Line ASRDummyLine1 
         Height          =   30
         Left            =   180
         Top             =   755
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   53
      End
      Begin VB.Label lblID 
         Caption         =   "Audit ID Column :"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   4680
         Width           =   1680
      End
      Begin VB.Label LblDescription 
         Caption         =   "Description Column :"
         Height          =   330
         Left            =   180
         TabIndex        =   24
         Top             =   4275
         Width           =   2085
      End
      Begin VB.Label lblTime 
         Caption         =   "Audit Time Column : "
         Height          =   330
         Left            =   180
         TabIndex        =   21
         Top             =   1395
         Width           =   1815
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1815
         Width           =   2010
      End
      Begin VB.Label lblChangeDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Date Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   975
         Width           =   1920
      End
      Begin VB.Label lblTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Table Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2220
         Width           =   2010
      End
      Begin VB.Label lblColumn 
         BackStyle       =   0  'Transparent
         Caption         =   "Column Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   2625
         Width           =   1755
      End
      Begin VB.Label lblAuditTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblStartSession 
         BackStyle       =   0  'Transparent
         Caption         =   "New Value Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   3450
         Width           =   1755
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   3855
         Width           =   1455
      End
      Begin VB.Label lblReason 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Value Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   3035
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmAuditSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReadOnly As Boolean
Public Changed As Boolean

Private Sub cboAuditTable_Click()

  Dim lngAuditTable As Long
  Dim objctl As Control

  lngAuditTable = GetComboItem(cboAuditTable)

  ' Clear the current contents of the combos.
  For Each objctl In Me
    If TypeOf objctl Is ComboBox And Not objctl.Name = "cboAuditTable" Then
      With objctl
        .Clear
        AddItemToComboBox objctl, "<None>", 0
      End With
    End If
  Next objctl

  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngAuditTable

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> lngAuditTable Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
        
          If !DataType = dtTIMESTAMP Then
            AddItemToComboBox cboAuditDate, !ColumnName, !ColumnID
          End If

          If !DataType = dtVARCHAR Then

            If !Size >= 50 Then
              AddItemToComboBox cboUser, !ColumnName, !ColumnID
              AddItemToComboBox cboModule, !ColumnName, !ColumnID
            End If

            If !Size = 8 Then
              AddItemToComboBox cboTime, !ColumnName, !ColumnID
            End If

            If !Size >= 200 Then
              AddItemToComboBox cboTable, !ColumnName, !ColumnID
              AddItemToComboBox cboColumn, !ColumnName, !ColumnID
            End If
  
            If !Size >= 255 Then
              AddItemToComboBox cboDescription, !ColumnName, !ColumnID
            End If
  
            If !MultiLine = True Then
              AddItemToComboBox cboOldValue, !ColumnName, !ColumnID
              AddItemToComboBox cboNewValue, !ColumnName, !ColumnID
            End If
          
          ElseIf !DataType = dtINTEGER Then
              AddItemToComboBox cboID, !ColumnName, !ColumnID
          
          End If
            
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOk_Click()

  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If

End Sub

Private Function SaveChanges() As Boolean
  
  SaveChanges = False
   
  Screen.MousePointer = vbHourglass
  
  SaveParam gsPARAMETERKEY_AUDITTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboAuditTable)
  SaveParam gsPARAMETERKEY_AUDITDATECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboAuditDate)
  SaveParam gsPARAMETERKEY_AUDITTIMECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTime)
  SaveParam gsPARAMETERKEY_AUDITUSERCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboUser)
  SaveParam gsPARAMETERKEY_AUDITTABLECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTable)
  SaveParam gsPARAMETERKEY_AUDITCOLUMNCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboColumn)
  SaveParam gsPARAMETERKEY_AUDITOLDVALUECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboOldValue)
  SaveParam gsPARAMETERKEY_AUDITNEWVALUECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboNewValue)
  SaveParam gsPARAMETERKEY_AUDITMODULECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboModule)
  SaveParam gsPARAMETERKEY_AUDITDESCRIPTIONCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboDescription)
  SaveParam gsPARAMETERKEY_AUDITIDCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboID)
  
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault
  
End Function


Private Sub SaveParam(strKey As String, strType As String, lngValue As Long)

  With recModuleSetup

    .Index = "idxModuleParameter"
    .Seek "=", gsMODULEKEY_AUDIT, strKey
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_AUDIT
      !parameterkey = strKey
    Else
      .Edit
    End If
    !ParameterType = strType
    !parametervalue = lngValue
    .Update

  End With

End Sub

Private Function ReadParam(strKey As String) As Long

  With recModuleSetup
    .Index = "idxModuleParameter"

    .Seek "=", gsMODULEKEY_AUDIT, strKey
    If .NoMatch Then
      ReadParam = 0
    Else
      ReadParam = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

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

  Me.ReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If Me.ReadOnly Then
    ControlsDisableAll Me
  End If

  PopulateBaseTableCombos
  InitialiseCombos
  
  Me.Changed = False
  Screen.MousePointer = vbDefault

End Sub

Private Sub InitialiseCombos()

  SetComboItem cboAuditTable, ReadParam(gsPARAMETERKEY_AUDITTABLE)
  SetComboItem cboAuditDate, ReadParam(gsPARAMETERKEY_AUDITDATECOLUMN)
  SetComboItem cboTime, ReadParam(gsPARAMETERKEY_AUDITTIMECOLUMN)
  SetComboItem cboUser, ReadParam(gsPARAMETERKEY_AUDITUSERCOLUMN)
  SetComboItem cboTable, ReadParam(gsPARAMETERKEY_AUDITTABLECOLUMN)
  SetComboItem cboColumn, ReadParam(gsPARAMETERKEY_AUDITCOLUMNCOLUMN)
  SetComboItem cboOldValue, ReadParam(gsPARAMETERKEY_AUDITOLDVALUECOLUMN)
  SetComboItem cboNewValue, ReadParam(gsPARAMETERKEY_AUDITNEWVALUECOLUMN)
  SetComboItem cboModule, ReadParam(gsPARAMETERKEY_AUDITMODULECOLUMN)
  SetComboItem cboDescription, ReadParam(gsPARAMETERKEY_AUDITDESCRIPTIONCOLUMN)
  SetComboItem cboID, ReadParam(gsPARAMETERKEY_AUDITIDCOLUMN)
 
End Sub

Private Sub PopulateBaseTableCombos()
  
  Dim lngPostTable As Long
  
  cboAuditTable.Clear
  AddItemToComboBox cboAuditTable, "<None>", 0
  cboAuditTable.ListIndex = 0

  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If Not !Deleted Then
        AddItemToComboBox cboAuditTable, !TableName, !TableID
      End If
      .MoveNext
    Loop
  End With

End Sub


