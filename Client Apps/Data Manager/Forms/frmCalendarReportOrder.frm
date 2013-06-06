VERSION 5.00
Begin VB.Form frmCalendarReportOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Sort Order"
   ClientHeight    =   2130
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
   HelpContextID   =   1070
   Icon            =   "frmCalendarReportOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   3825
      Begin VB.ComboBox cboColumns 
         Height          =   315
         ItemData        =   "frmCalendarReportOrder.frx":000C
         Left            =   1065
         List            =   "frmCalendarReportOrder.frx":0013
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2565
      End
      Begin VB.OptionButton optAsc 
         Caption         =   "&Ascending"
         Height          =   225
         Left            =   1065
         TabIndex        =   1
         Top             =   780
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Descending"
         Height          =   225
         Left            =   1065
         TabIndex        =   2
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label lblColumns 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   3
      Top             =   1605
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   1365
      TabIndex        =   4
      Top             =   1605
      Width           =   1200
   End
End
Attribute VB_Name = "frmCalendarReportOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmForm As frmCalendarReport
Private mblnNew As Boolean

Private mavSelectedColumns() As String
Private mcolColumnDetails As CColumnPrivileges

Private mblnUserCancelled As Boolean

Public Function Initialise(pblnNew As Boolean, pfrmForm As frmCalendarReport, _
                            pcolColumnsList As CColumnPrivileges, psSelectedColumns As String, _
                            Optional plngColumnID As Long, Optional pstrSortOrder As String) As Boolean

  On Error GoTo Initialise_ERROR
  
  Dim pvarOldPosition As Variant
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim plngExclusion As Long
  Dim aryTemp() As String
  
  Set mfrmForm = pfrmForm
  mblnNew = pblnNew
  
  Set mcolColumnDetails = pcolColumnsList
  
  aryTemp = Split(psSelectedColumns, ",")
  
  mavSelectedColumns = aryTemp
  
  pintLoop = PopulateCombo(plngColumnID)
  
  If pblnNew Then
    If pintLoop = 0 Then
      COAMsgBox "No columns on the base table.", vbExclamation + vbOKOnly, mfrmForm.Caption
      Initialise = False
      Exit Function
    End If
     
    If (mfrmForm.grdOrder.Rows = pintLoop And pintLoop > 0) Or (cboColumns.ListCount = 0) Then
      COAMsgBox "You have selected all base columns in the sort order.", vbExclamation + vbOKOnly, mfrmForm.Caption
      Initialise = False
      Exit Function
    End If
     
    ' Now set the first item in the combobox
    cboColumns.ListIndex = 0
    
  Else
    mfrmForm.grdOrder.MoveFirst
    Do Until mfrmForm.grdOrder.Columns("ColumnID").Value = plngColumnID
      mfrmForm.grdOrder.MoveNext
    Loop
      
  End If

  If UCase(pstrSortOrder) = "DESC" Then
    optDesc.Value = True
  Else
    optAsc.Value = True
  End If
  
  Initialise = True
  Exit Function
  
Initialise_ERROR:
  
  Initialise = False
  COAMsgBox "Error initialising the order form." & vbCrLf & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, mfrmForm.Caption

End Function
Private Function IsAlreadyAnOrderColumn(plngColumnID As Long) As Boolean
  
  Dim i As Integer
  
  For i = 0 To UBound(mavSelectedColumns) Step 1
    If plngColumnID = mavSelectedColumns(i) Then
      IsAlreadyAnOrderColumn = True
      Exit Function
    End If
  Next i

  IsAlreadyAnOrderColumn = False
  
End Function
Private Function PopulateCombo(SelectedID As Long) As Integer

  Dim lngColumnCount As Long
  Dim lngID As Long

  lngColumnCount = 0

  Dim objColumn As CColumnPrivilege

  cboColumns.Clear

  For Each objColumn In mcolColumnDetails

    lngID = objColumn.ColumnID
    
    If objColumn.ColumnType <> ColumnTypes.colSystem _
      And objColumn.ColumnType <> ColumnTypes.colLink _
      And objColumn.DataType <> SQLDataType.sqlOle _
      And objColumn.DataType <> SQLDataType.sqlVarBinary Then
      
      If lngID = SelectedID Then
        cboColumns.AddItem objColumn.ColumnName
        cboColumns.ItemData(cboColumns.NewIndex) = objColumn.ColumnID
        cboColumns.ListIndex = cboColumns.NewIndex
  
      ElseIf Not (IsAlreadyAnOrderColumn(lngID)) Then
        cboColumns.AddItem objColumn.ColumnName
        cboColumns.ItemData(cboColumns.NewIndex) = objColumn.ColumnID
  
      End If
    
    End If
    
    lngColumnCount = lngColumnCount + 1
  Next objColumn

  PopulateCombo = lngColumnCount

End Function
Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Private Sub cmdAction_Click(Index As Integer)

  Dim pvarbookmark As Variant
  Dim plngRow As Long
  
  ' OK pressed
  
  If Index = 1 Then
    
    If cboColumns.Text = "" Then
      COAMsgBox "You must select a column.", vbExclamation + vbOKOnly, mfrmForm.Caption
      Exit Sub
    End If

    If mblnNew = True Then
      mfrmForm.grdOrder.MoveLast
      mfrmForm.grdOrder.AddItem _
        CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
        mfrmForm.cboBaseTable.Text & "." & Me.cboColumns.Text & vbTab & _
        IIf(optAsc.Value, "Asc", "Desc")
        mfrmForm.grdOrder.MoveLast
       mfrmForm.grdOrder.SelBookmarks.Add mfrmForm.grdOrder.Bookmark
    Else
    
      With mfrmForm.grdOrder
    
        plngRow = .AddItemRowIndex(.Bookmark)
        pvarbookmark = .GetBookmark(plngRow)
 
        .RemoveItem plngRow
        mfrmForm.grdOrder.AddItem _
        CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
        mfrmForm.cboBaseTable.Text & "." & Me.cboColumns.Text & vbTab & _
        IIf(optAsc.Value, "Asc", "Desc"), plngRow
        .SelBookmarks.Add .AddItemBookmark(plngRow)
      
        .AddItemRowIndex (.Bookmark)
        .Bookmark = pvarbookmark
      
      End With
      
    End If
  
  Else
    mblnUserCancelled = True
  End If

  'AE20071025 Fault #6797
  'Unload Me
  Me.Hide

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



