VERSION 5.00
Begin VB.Form frmMailMergeOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail Merge Order"
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
   HelpContextID   =   1046
   Icon            =   "frmMailMergeOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   1365
      TabIndex        =   7
      Top             =   1605
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   6
      Top             =   1605
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.OptionButton optDesc 
         Caption         =   "&Descending"
         Height          =   225
         Left            =   975
         TabIndex        =   3
         Top             =   1080
         Width           =   1485
      End
      Begin VB.OptionButton optAsc 
         Caption         =   "&Ascending"
         Height          =   225
         Left            =   975
         TabIndex        =   2
         Top             =   780
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   780
         Width           =   525
      End
      Begin VB.Label lblColumns 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmMailMergeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmForm As Form
Private mblnNew As Boolean
Private mblnUserCancelled As Boolean

Public Function Initialise(pblnNew As Boolean, pfrmForm As Form, Optional plngColExprID As Long, Optional pstrSortOrder As String) As Boolean

  On Error GoTo Initialise_ERROR
  
  Dim pvarOldPosition As Variant
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim plngExclusion As Long
  
  Set mfrmForm = pfrmForm
  mblnNew = pblnNew
  
  pintLoop = PopulateCombo(plngColExprID)
  
  If pblnNew Then
  
    If pintLoop = 0 Then
      COAMsgBox "You must add a column to the definition before you can add to the sort order.", vbExclamation + vbOKOnly, mfrmForm.Caption
      Initialise = False
      Exit Function
    End If
     
    If (mfrmForm.grdReportOrder.Rows = pintLoop And pintLoop > 0) Or (cboColumns.ListCount = 0) Then
      COAMsgBox "You have selected all existing columns in the sort order." & vbCrLf & _
             "To add more sort order columns, you must add more columns to the definition.", vbExclamation + vbOKOnly, mfrmForm.Caption
      Initialise = False
      Exit Function
    End If
     
    ' Now set the first item in the combobox
    cboColumns.ListIndex = 0
    
  Else
  
    mfrmForm.grdReportOrder.MoveFirst
    Do Until mfrmForm.grdReportOrder.Columns("ColExprID").Value = plngColExprID
      mfrmForm.grdReportOrder.MoveNext
    Loop
      
  End If

  If Left(pstrSortOrder, 1) = "D" Then
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


Private Function IsntAlreadyAnOrderColumn(plngColExprID As Long) As Boolean

  Dim pvarOldPosition As Variant
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With mfrmForm.grdReportOrder
    
    ' Store the old position so we can return it after we have looped thru the grid
    pvarOldPosition = .Bookmark

    .Redraw = False
    .MoveFirst
    pintLoop = 0
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
        IsntAlreadyAnOrderColumn = False
        .Redraw = True
        Exit Function
      End If
      pintLoop = pintLoop + 1
    Loop
  
    .Bookmark = pvarOldPosition
    .Redraw = True
    '.SelBookmarks.Add .Bookmark
    '.MoveLast
  
  End With

  IsntAlreadyAnOrderColumn = True

End Function

Private Sub cmdAction_Click(Index As Integer)

  Dim pvarbookmark As Variant
  Dim plngRow As Long
  Dim pstrRow As String
  
  If cboColumns.Text = "" Then
    COAMsgBox "You must select a column.", vbExclamation + vbOKOnly, mfrmForm.Caption
    Exit Sub
  End If
  ' OK pressed
  
  If Index = 1 Then
    
    If mblnNew = True Then
      mfrmForm.grdReportOrder.MoveLast
      mfrmForm.grdReportOrder.AddItem _
        CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
        Me.cboColumns.Text & vbTab & _
        IIf(Me.optAsc.Value, "Ascending", "Descending")
        mfrmForm.grdReportOrder.MoveLast
       mfrmForm.grdReportOrder.SelBookmarks.Add mfrmForm.grdReportOrder.Bookmark
    Else
    
      With mfrmForm.grdReportOrder
        plngRow = .AddItemRowIndex(.Bookmark)
 
        pstrRow = CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
                  Me.cboColumns.Text & vbTab & _
                  IIf(Me.optAsc.Value, "Ascending", "Descending")
                  
        .RemoveItem plngRow
        .AddItem pstrRow, plngRow
        .Bookmark = .AddItemBookmark(plngRow)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
        
      End With
      
    End If
  'AE20071025 Fault #6797
  ElseIf Index = 0 Then
    mblnUserCancelled = True
  End If

  'AE20071025 Fault #6797
  'Unload Me
  Me.Hide
  
End Sub


Private Function PopulateCombo(SelectedID As Long) As Integer

  Dim lngColumnCount As Long
  Dim pintLoop As Integer
  Dim lngID As Long
  
  With mfrmForm.ListView2
    
    pintLoop = 1
    lngColumnCount = 0
    Do While pintLoop <= .ListItems.Count
        
      With .ListItems(pintLoop)
        If Left(.Key, 1) = "C" Or mfrmForm.Name = "frmMatchDef" Then
          lngID = Val(Mid(.Key, 2))
          
          If lngID = SelectedID Then
            cboColumns.AddItem .Text
            cboColumns.ItemData(cboColumns.NewIndex) = lngID
            cboColumns.ListIndex = cboColumns.NewIndex
          ElseIf IsntAlreadyAnOrderColumn(lngID) Then
            cboColumns.AddItem .Text
            cboColumns.ItemData(cboColumns.NewIndex) = lngID
          End If
        
          lngColumnCount = lngColumnCount + 1
        End If
      End With
    
      pintLoop = pintLoop + 1
    Loop
    
  End With
      
  PopulateCombo = lngColumnCount

End Function

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

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property


