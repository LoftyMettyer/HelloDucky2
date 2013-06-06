VERSION 5.00
Begin VB.Form frmExportOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Order"
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
   HelpContextID   =   1037
   Icon            =   "frmExportOrder.frx":0000
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
      TabIndex        =   2
      Top             =   0
      Width           =   3825
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2655
      End
      Begin VB.OptionButton optAsc 
         Caption         =   "&Ascending"
         Height          =   225
         Left            =   975
         TabIndex        =   4
         Top             =   780
         Value           =   -1  'True
         Width           =   1890
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Descending"
         Height          =   225
         Left            =   975
         TabIndex        =   3
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label lblColumns 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   750
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
      TabIndex        =   1
      Top             =   1605
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   1455
      TabIndex        =   0
      Top             =   1605
      Width           =   1200
   End
End
Attribute VB_Name = "frmExportOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmForm As frmExport
Private mblnNew As Boolean

Public Function Initialise(pblnNew As Boolean, pfrmForm As frmExport, Optional plngColExprID As Long, Optional pstrSortOrder As String) As Boolean

  On Error GoTo Initialise_ERROR
  
  Dim pvarOldPosition As Variant
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim plngExclusion As Long
  Dim strAlreadyAdded As String
  Dim lngID As Long
  
  Set mfrmForm = pfrmForm
  mblnNew = pblnNew
  
  If pblnNew Then
    
    With mfrmForm.grdColumns
     
      ' Store the old position so we can return it after we have looped thru the grid
      pvarOldPosition = .Bookmark
      
      'AE20071024 Fault #8474
      Screen.MousePointer = vbHourglass
      .Redraw = False
      
      ' Loop thru the export grid, adding data to the combo if they are columns
      .MoveFirst
      strAlreadyAdded = " "
        Do Until pintLoop = .Rows
          pvarbookmark = .GetBookmark(pintLoop)
          If .Columns("Type").CellText(pvarbookmark) = "C" Then
            lngID = .Columns("ColExprID").CellText(pvarbookmark)
            If IsntAlreadyAnOrderColumn(lngID) = True Then
              If InStr(strAlreadyAdded, " " & CStr(lngID) & " ") = 0 Then
                cboColumns.AddItem .Columns("Data").CellText(pvarbookmark)
                cboColumns.ItemData(Me.cboColumns.NewIndex) = lngID
                strAlreadyAdded = strAlreadyAdded & CStr(lngID) & " "
              End If
            End If
          End If
          pintLoop = pintLoop + 1
        Loop
    
     .Bookmark = pvarOldPosition
     .SelBookmarks.Add .Bookmark
     
     'AE20071024 Fault #8474
     .Redraw = True
      Screen.MousePointer = vbNormal
      
'      If (mfrmForm.grdExportOrder.Rows = pintLoop And pintLoop > 0) Or (cboColumns.ListCount = 0) Then
'        COAMsgBox "You have selected all existing export columns in the sort order." & vbCrLf & "To add more sort order columns, you must add more columns to the export definition.", vbExclamation + vbOKOnly, "Export"
'        Initialise = False
'        Exit Function
'      End If
'
'      If pintLoop = 0 Then
'       COAMsgBox "You can order the export by columns selected in the export." & vbCrLf & "You must select a column for the export before you can sort by it.", vbExclamation + vbOKOnly, "Export"
'       Initialise = False
'       Exit Function
'      End If
     
        If cboColumns.ListCount = 0 Then
          If mfrmForm.grdColumns.Rows = 0 Then
            COAMsgBox "You must add a column to the export before you can add to the sort order.", vbExclamation + vbOKOnly, "Export"
          Else
            COAMsgBox "You must add more columns to the export before you can add to the sort order.", vbExclamation + vbOKOnly, "Export"
          End If
        Initialise = False
        Exit Function
        End If
     
     
    End With
    
    ' Now set the first item in the combobox, disabling combo if only 1 column available
    
    With cboColumns
      If .ListCount = 1 Then
        .ListIndex = 0
        .Enabled = False
        .BackColor = vbButtonFace
      ElseIf .ListCount > 1 Then
        .ListIndex = 0
        .Enabled = True
        .BackColor = vbWindowBackground
      End If
    End With
     
  Else
  
    With mfrmForm.grdColumns
     
       ' Store the old position so we can return it after we have looped thru the grid
       pvarOldPosition = .Bookmark
       
        plngExclusion = mfrmForm.grdExportOrder.Columns("ColExprID").CellValue(mfrmForm.grdExportOrder.Bookmark)
       
       ' Loop thru the export grid, adding data to the combo if they are columns
       .MoveFirst
         Do Until pintLoop = .Rows
           pvarbookmark = .GetBookmark(pintLoop)
           If .Columns("Type").CellText(pvarbookmark) = "C" Then
             If IsntAlreadyAnOrderColumn2(.Columns("ColExprID").CellText(pvarbookmark), plngExclusion) = True Then
               Me.cboColumns.AddItem .Columns("Data").CellText(pvarbookmark)
               Me.cboColumns.ItemData(Me.cboColumns.NewIndex) = .Columns("ColExprID").CellValue(pvarbookmark)
             End If
           End If
           pintLoop = pintLoop + 1
         Loop
      
       .Bookmark = pvarOldPosition
       .SelBookmarks.Add .Bookmark
     
      mfrmForm.grdExportOrder.MoveFirst
      
      Do Until mfrmForm.grdExportOrder.Columns("ColExprID").Value = plngExclusion
        mfrmForm.grdExportOrder.MoveNext
      Loop
      
    End With
  
    SetComboText cboColumns, mfrmForm.grdExportOrder.Columns("Column").CellText(mfrmForm.grdExportOrder.Bookmark)
  
    Select Case mfrmForm.grdExportOrder.Columns("Sort Order").CellText(mfrmForm.grdExportOrder.Bookmark)
      Case "Ascending": optAsc.Value = True
      Case Else: optDesc.Value = True
    End Select
  
  End If
  
  Initialise = True
  Exit Function
  
Initialise_ERROR:
  
  Initialise = False
  COAMsgBox "Error initialising the Export Order form." & vbCrLf & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, "Export"

End Function


Private Function IsntAlreadyAnOrderColumn(plngColExprID As Long) As Boolean

  Dim pvarOldPosition As Variant
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With mfrmForm.grdExportOrder
    
    ' Store the old position so we can return it after we have looped thru the grid
    pvarOldPosition = .Bookmark
    
    ' Loop thru the export grid, adding data to the combo if they are columns
    
    'AE20071024 Fault #8474
    .Redraw = False
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          'AE20071024 Fault #8474
          .Bookmark = pvarOldPosition
          .Redraw = True
          IsntAlreadyAnOrderColumn = False
          Exit Function
        End If
        pintLoop = pintLoop + 1
      Loop
  
    .Bookmark = pvarOldPosition
    '.SelBookmarks.Add .Bookmark
    '.MoveLast
    'AE20071024 Fault #8474
    .Redraw = True
  End With

  IsntAlreadyAnOrderColumn = True

End Function

Private Function IsntAlreadyAnOrderColumn2(plngColExprID As Long, plngExclusionID As Long) As Boolean

  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With mfrmForm.grdExportOrder
    
    ' Loop thru the export grid, adding data to the combo if they are columns
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          If .Columns("ColExprID").CellText(pvarbookmark) <> plngExclusionID Then
            IsntAlreadyAnOrderColumn2 = False
            Exit Function
          End If
        End If
        pintLoop = pintLoop + 1
      Loop
  
    .MoveFirst
    
  End With

  IsntAlreadyAnOrderColumn2 = True

End Function

Private Sub cmdAction_Click(Index As Integer)

  Dim pvarbookmark As Variant
  Dim plngRow As Long
  
  If cboColumns.Text = "" Then
    COAMsgBox "You must select a column.", vbExclamation + vbOKOnly, "Export"
    Exit Sub
  End If
  ' OK pressed
  
  If Index = 1 Then
    
    If mblnNew = True Then
      mfrmForm.grdExportOrder.MoveLast
      mfrmForm.grdExportOrder.AddItem _
        CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
        Me.cboColumns.Text & vbTab & _
        IIf(Me.optAsc.Value, "Ascending", "Descending")
        mfrmForm.grdExportOrder.MoveLast
       mfrmForm.grdExportOrder.SelBookmarks.Add mfrmForm.grdExportOrder.Bookmark
    Else
    
      With mfrmForm.grdExportOrder
    
        plngRow = .AddItemRowIndex(.Bookmark)
        pvarbookmark = .GetBookmark(plngRow)
 
        .RemoveItem plngRow
        mfrmForm.grdExportOrder.AddItem _
        CStr(Me.cboColumns.ItemData(Me.cboColumns.ListIndex)) & vbTab & _
        Me.cboColumns.Text & vbTab & _
        IIf(Me.optAsc.Value, "Ascending", "Descending"), plngRow
        .SelBookmarks.Add .AddItemBookmark(plngRow)
      
        .AddItemRowIndex (.Bookmark)
        .Bookmark = pvarbookmark
      
      End With
      
    End If
    
    ' RH 15/02/01
    mfrmForm.Changed = True
    
  End If

  Unload Me
  
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



