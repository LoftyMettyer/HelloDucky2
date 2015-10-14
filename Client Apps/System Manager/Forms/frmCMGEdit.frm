VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmCMGEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit CMG & Centrefile Layout"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5083
   Icon            =   "frmCMGEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   400
      Index           =   0
      Left            =   3660
      TabIndex        =   7
      Top             =   1875
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2415
      TabIndex        =   6
      Top             =   1875
      Width           =   1200
   End
   Begin VB.Frame frmDetails 
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   4755
      Begin VB.CheckBox chkExportItem 
         Caption         =   "&Export"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtLayoutItem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   3150
      End
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   100
         Text            =   "0"
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "&Field Size :"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   780
         Width           =   765
      End
      Begin VB.Label lblExportItem 
         AutoSize        =   -1  'True
         Caption         =   "E&xport Item :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmCMGEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbLoading As Boolean
Private mfrmForm As frmCMGSetup
Private mblnUserCancelled As Boolean
Private mblnUserChanged As Boolean

Public Function Initialise(sLayoutItem As String, iSize As Integer, fExported As Boolean, Optional blnEditing As Boolean, Optional pfrmForm As Form) As Boolean
  
  mbLoading = True
  mblnUserChanged = False
  mblnUserCancelled = False
  
  Set mfrmForm = pfrmForm
  
  Me.txtLayoutItem.Text = sLayoutItem
  Me.spnMaxSize.value = iSize
  Me.chkExportItem.value = IIf(fExported, 1, 0)
           
  Me.spnMaxSize.Enabled = Not (mfrmForm.chkUseCSV.value = 1)
  Me.spnMaxSize.Enabled = Me.chkExportItem.value
           
  If Me.txtLayoutItem.Text = "Record Identifier" Or Me.txtLayoutItem.Text = "Output Column" Then
    Me.chkExportItem.value = 1
    Me.chkExportItem.Enabled = False
  Else
    Me.chkExportItem.Enabled = True
  End If
           
  Initialise = True
  mbLoading = False
  Me.Show vbModal

End Function

Private Sub chkExportItem_Click()
  If Not mbLoading Then
    Me.spnMaxSize.Enabled = Me.chkExportItem.value
    mblnUserChanged = True
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub spnMaxSize_Change()
  If Not mbLoading Then
    mblnUserChanged = True
  End If
End Sub

Private Sub cmdAction_Click(Index As Integer)
 
  ' OK pressed
  If Index = 1 Then
  
    Dim intSourceRow As Integer
    Dim strSourceRow As String
    Dim intDestinationRow As Integer
    Dim strDestinationRow As String
        
    intDestinationRow = mfrmForm.grdCMGLayout.AddItemRowIndex(mfrmForm.grdCMGLayout.Bookmark)
    mfrmForm.grdCMGLayout.MoveNext
    strDestinationRow = Me.txtLayoutItem.Text & vbTab & CStr(Me.spnMaxSize.value) & vbTab & IIf(chkExportItem.value = 0, "False", "True")
    
    mfrmForm.grdCMGLayout.RemoveItem intDestinationRow
    mfrmForm.grdCMGLayout.AddItem strDestinationRow, intDestinationRow
    
    mfrmForm.grdCMGLayout.SelBookmarks.RemoveAll
    mfrmForm.grdCMGLayout.MoveNext
    mfrmForm.grdCMGLayout.Bookmark = mfrmForm.grdCMGLayout.AddItemBookmark(intDestinationRow)
    ' mfrmForm.grdCMGLayout.SelBookmarks.Add mfrmForm.grdCMGLayout.AddItemBookmark(intDestinationRow)
    
'    UpdateButtonStatus
    
  ElseIf Index = 0 Then
    mblnUserCancelled = True
  End If

  Me.Hide
  
  ' Set objColumn = Nothing
  
End Sub

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Property Get UserChanged() As Boolean
  UserChanged = mblnUserChanged
End Property
