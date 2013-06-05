VERSION 5.00
Begin VB.Form frmGlobalFunctionsTableValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lookup Table Value"
   ClientHeight    =   3465
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
   Icon            =   "frmGlobalFunctionsTableValue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1455
      TabIndex        =   3
      Top             =   2955
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2715
      TabIndex        =   4
      Top             =   2955
      Width           =   1200
   End
   Begin VB.ListBox lstRecords 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   1815
      Left            =   1095
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1000
      Width           =   2820
   End
   Begin VB.ComboBox cboColumn 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1110
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2805
   End
   Begin VB.ComboBox cboTable 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1110
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   200
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value :"
      Height          =   195
      Index           =   2
      Left            =   195
      TabIndex        =   7
      Top             =   1060
      Width           =   495
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column :"
      Height          =   195
      Index           =   1
      Left            =   200
      TabIndex        =   6
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table :"
      Height          =   195
      Index           =   0
      Left            =   200
      TabIndex        =   5
      Top             =   260
      Width           =   495
   End
End
Attribute VB_Name = "frmGlobalFunctionsTableValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancelled As Boolean
Private mbLoading As Boolean
Private datData As DataMgr.clsDataAccess
Private miColumnDataType As SQLDataType

Public Property Get LookupTableID() As Long
  If cboTable.ListIndex <> -1 Then
    LookupTableID = cboTable.ItemData(cboTable.ListIndex)
  End If
End Property

Public Property Get LookupColumnID() As Long
  If cboColumn.ListIndex <> -1 Then
    LookupColumnID = cboColumn.ItemData(cboColumn.ListIndex)
  End If
End Property


Public Function Initialise(lTableID As Long, lColumnID As Long, iColumnDataType As SQLDataType, Optional strValue As String)

  Dim lngCount As Long
  
  Set datData = New clsDataAccess
  
  miColumnDataType = iColumnDataType
  
  GetTables
  
  'Check for at least one table
  Initialise = (cboTable.ListCount > 0)
  If Initialise Then
    SetComboItem cboTable, lTableID
    If cboTable.ListIndex = -1 Then
      cboTable.ListIndex = 0
    End If


    SetComboItem cboColumn, lColumnID
    If cboColumn.ListIndex = -1 And cboColumn.ListCount > 0 Then
      cboColumn.ListIndex = 0
    End If


    If lstRecords.ListCount > 0 Then
      For lngCount = 0 To lstRecords.ListCount - 1
        If Trim$(lstRecords.List(lngCount)) = Trim$(strValue) Then
          lstRecords.ListIndex = lngCount
          Exit For
        End If
      Next
  
      If lstRecords.ListIndex = -1 Then
        lstRecords.ListIndex = 0
      End If
    End If

  Else
    COAMsgBox "There are no lookup columns which match the selected data type", vbExclamation
  
  End If

End Function

Public Property Get Cancelled() As Boolean

  Cancelled = mbCancelled

End Property

Private Sub cboColumn_Click()

  If Not mbLoading Then
    GetRecords
  End If

End Sub

Private Sub cboTable_Click()

  If Not mbLoading Then
    GetColumns
  End If

End Sub

Private Sub cmdCancel_Click()

  mbCancelled = True
  Me.Hide

End Sub

Private Sub cmdOK_Click()

  mbCancelled = False
  Me.Hide

End Sub

Private Sub Form_Activate()
  If lstRecords.Enabled Then
    lstRecords.SetFocus
  End If
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'  If UnloadMode = vbFormControlMenu Then
'        mbCancelled = True
'        Cancel = True
'        Me.Hide
'  End If
'
'End Sub

Private Sub GetTables()

  Dim rsTables As Recordset
  Dim strSQL As String

  strSQL = "SELECT DISTINCT ASRSysTables.TableName, ASRSysTables.TableID " & _
           "FROM ASRSysTables " & _
           "JOIN ASRSysColumns on ASRSysColumns.TableID = ASRSysTables.TableID " & _
           "WHERE TableType = 3"
           
  If miColumnDataType <> 0 Then
    strSQL = strSQL & " AND ASRSysColumns.DataType = " & CStr(miColumnDataType)
  End If
  'Set rsTables = datGeneral.GetLookupTables
  Set rsTables = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  With cboTable
    
    .Clear
    mbLoading = True
    Do While Not rsTables.EOF
      If gcoTablePrivileges.Item(rsTables!TableName).AllowSelect Then
        .AddItem rsTables!TableName
        .ItemData(.NewIndex) = rsTables!TableID
      End If
      rsTables.MoveNext
    Loop
    
    mbLoading = False
    
    .Enabled = (.ListCount > 1)
    .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
    If .ListCount > 0 Then
      .ListIndex = 0
    End If

  End With
        
End Sub

Private Sub GetColumns()
  
  Dim rsColumns As Recordset
  Dim strSQL As String
    
  Screen.MousePointer = vbHourglass
  'Set rsColumns = datGeneral.GetColumns(cboTable.ItemData(cboTable.ListIndex))

  strSQL = "SELECT ColumnName, ColumnID " & _
           "FROM ASRSysColumns " & _
           "WHERE TableID = " & CStr(cboTable.ItemData(cboTable.ListIndex))
  If miColumnDataType <> 0 Then
    strSQL = strSQL & " AND DataType = " & CStr(miColumnDataType)
  End If
  
  Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  
  With cboColumn
    .Clear
    
    mbLoading = True
        
    Do While Not rsColumns.EOF
      If Mid$(rsColumns!ColumnName, 1, 2) <> "ID" Then
        .AddItem rsColumns!ColumnName
        .ItemData(.NewIndex) = rsColumns!ColumnID
      End If
      
      rsColumns.MoveNext
    Loop
    
    mbLoading = False
    
    .Enabled = (.ListCount > 1)
    .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
    If .ListCount > 0 Then
      .ListIndex = 0
    Else
      'Need to clear the listview
      With lstRecords
        .Clear
        .Enabled = False
        .BackColor = vbButtonFace
      End With
    End If

  End With
    
  Screen.MousePointer = vbDefault
    
End Sub

Private Sub GetRecords()

  Dim rsRecords As Recordset

  On Error GoTo Err_Trap

  Set rsRecords = datGeneral.GetLookupRecords(cboTable.Text, cboColumn.Text)

  With lstRecords
    .Clear
    Do While Not rsRecords.EOF
      If Not IsNull(rsRecords!Record) Then
        .AddItem rsRecords!Record
        .ItemData(.NewIndex) = rsRecords!RecordID
      End If
      rsRecords.MoveNext
    Loop
        
    .Enabled = (.ListCount > 0)
    cmdOK.Enabled = (.ListCount > 0)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
  
  End With
    
Exit Sub
    
Err_Trap:
    COAMsgBox Err.Description
    
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

Private Sub lstRecords_DblClick()
  cmdOK_Click
End Sub

