VERSION 5.00
Begin VB.Form frmWorkflowEmailAddressColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Address Column"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5069
   Icon            =   "frmWorkflowEmailAddressColumn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboColumn 
      Height          =   315
      Left            =   1000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   200
      Width           =   3000
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   1400
      TabIndex        =   2
      Top             =   700
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Label lblColumn 
      Caption         =   "Column :"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   255
      Width           =   810
   End
End
Attribute VB_Name = "frmWorkflowEmailAddressColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mfChanged As Boolean
Private mlngColumnID As Long
Private mlngPersonnelTableID As Long
Private msSelectedColumns As String

Private mblnReadOnly As Boolean

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property


Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property


Public Sub Initialize(plngColumnID As Long, _
  plngPersonnelTableID As Long, _
  psSelectedColumns As String)
  
  mlngColumnID = plngColumnID
  mlngPersonnelTableID = plngPersonnelTableID
  msSelectedColumns = psSelectedColumns

  GetColumns

  mfChanged = False
  
  RefreshControls
  
End Sub



Public Property Get ColumnID() As Long
  If cboColumn.ListCount > 0 Then
    ColumnID = cboColumn.ItemData(cboColumn.ListIndex)
  Else
    ColumnID = 0
  End If
  
End Property

Public Property Get ColumnName() As String
  If cboColumn.ListCount > 0 Then
    ColumnName = cboColumn.List(cboColumn.ListIndex)
  Else
    ColumnName = ""
  End If
  
End Property


Private Sub GetColumns()
  ' Populate the columns combo.
  Dim sSQL As String
  Dim rsTemp As dao.Recordset
  Dim iDefaultItem As Integer
  Dim sTableName As String

  iDefaultItem = 0

  cboColumn.Clear

  If mlngPersonnelTableID > 0 Then
    sSQL = "SELECT tmpColumns.columnID, tmpColumns.columnName" & _
      " FROM tmpColumns" & _
      " WHERE (tmpColumns.deleted = FALSE)" & _
      " AND (tmpColumns.tableID = " & CStr(mlngPersonnelTableID) & ")" & _
      " AND (tmpColumns.dataType = " & CStr(dtVARCHAR) & ")" & _
      " AND ((tmpColumns.columnID NOT IN (" & msSelectedColumns & "))" & _
      "   OR (tmpColumns.columnID = " & CStr(mlngColumnID) & "))" & _
      " ORDER BY tmpColumns.columnName"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    While Not rsTemp.EOF
      cboColumn.AddItem rsTemp!ColumnName
      cboColumn.ItemData(cboColumn.NewIndex) = rsTemp!ColumnID

      If mlngColumnID = rsTemp!ColumnID Then
        iDefaultItem = cboColumn.NewIndex
      End If

      rsTemp.MoveNext
    Wend
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  If cboColumn.ListCount = 0 Then
    sSQL = "SELECT tmpTables.tableName" & _
      " FROM tmpTables" & _
      " WHERE tmpTables.tableID = " & CStr(mlngPersonnelTableID)
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsTemp.EOF And rsTemp.BOF) Then
      sTableName = rsTemp!TableName
    End If

    rsTemp.Close
    Set rsTemp = Nothing

    If Len(sTableName) = 0 Then
      MsgBox "All available columns have already been selected.", vbOKOnly + vbExclamation, Application.Name
    Else
      MsgBox "All '" & sTableName & "' columns have already been selected.", vbOKOnly + vbExclamation, Application.Name
    End If

    cmdCancel_Click
  Else
    cboColumn.ListIndex = iDefaultItem
  End If
  
End Sub




Private Sub RefreshControls()
  ' Disable the OK button as required.
  cmdOk.Enabled = (ColumnID <> mlngColumnID)
    
End Sub



Private Sub cboColumn_Click()
  mfChanged = True
  RefreshControls
  
End Sub


Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
  
End Sub

Private Sub cmdOK_Click()
  Cancelled = False
  Me.Hide
  
End Sub


Private Sub Form_Initialize()
  mblnReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)

End Sub

Private Sub Form_Load()
  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  fraOKCancel.BorderStyle = vbBSNone

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
        Cancel = True   'MH20021105 Fault 4694
    End Select
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub


