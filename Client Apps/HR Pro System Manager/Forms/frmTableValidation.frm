VERSION 5.00
Begin VB.Form frmTableValidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overlapping Dates"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5088
   Icon            =   "frmTableValidation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOverlap 
      Caption         =   "Overlapping Event : "
      Height          =   2175
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   5910
      Begin VB.ComboBox cboOverlapColumnType 
         Height          =   315
         ItemData        =   "frmTableValidation.frx":000C
         Left            =   1485
         List            =   "frmTableValidation.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1215
         Width           =   2010
      End
      Begin VB.TextBox txtOverlapMessage 
         Height          =   330
         Left            =   1475
         TabIndex        =   5
         Top             =   1620
         Width           =   4155
      End
      Begin VB.ComboBox cboOverlapColumnEndSession 
         Height          =   315
         ItemData        =   "frmTableValidation.frx":0010
         Left            =   3645
         List            =   "frmTableValidation.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   765
         Width           =   2010
      End
      Begin VB.ComboBox cboOverlapColumnEndDate 
         Height          =   315
         ItemData        =   "frmTableValidation.frx":0014
         Left            =   1485
         List            =   "frmTableValidation.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   765
         Width           =   2010
      End
      Begin VB.ComboBox cboOverlapColumnStartSession 
         Height          =   315
         ItemData        =   "frmTableValidation.frx":0018
         Left            =   3645
         List            =   "frmTableValidation.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2010
      End
      Begin VB.ComboBox cboOverlapColumnStartDate 
         Height          =   315
         ItemData        =   "frmTableValidation.frx":001C
         Left            =   1485
         List            =   "frmTableValidation.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label lblOverlapType 
         Caption         =   "Event Type :"
         Height          =   330
         Left            =   225
         TabIndex        =   12
         Top             =   1270
         Width           =   1140
      End
      Begin VB.Label lblOverlapMessage 
         AutoSize        =   -1  'True
         Caption         =   "Message :"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   1710
         Width           =   870
      End
      Begin VB.Label lblOverlapColumnEnd 
         AutoSize        =   -1  'True
         Caption         =   "Event End :"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   825
         Width           =   990
      End
      Begin VB.Label lblOverlapColumnStart 
         AutoSize        =   -1  'True
         Caption         =   "Event Start  :"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   375
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   3420
      TabIndex        =   6
      Top             =   2340
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4725
      TabIndex        =   7
      Top             =   2340
      Width           =   1200
   End
End
Attribute VB_Name = "frmTableValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancelled As Boolean
Private mbReadOnly As Boolean
Private mbLoading As Boolean
Private objValidationObject As clsTableValidation
Private miValidationType As ValidationType

Public Property Get ValidationObject() As clsTableValidation
  Set ValidationObject = objValidationObject
End Property

Public Property Let ValidationObject(ByVal NewValue As clsTableValidation)
  Set objValidationObject = NewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(ByVal NewValue As Boolean)
  If Not mbLoading Then
    cmdOK.Enabled = NewValue And Not mbReadOnly
  End If
End Property

Private Sub cboOverlapColumnEndDate_Click()
  txtOverlapMessage.Text = GenerateMessage
  Me.Changed = True
End Sub

Private Sub cboOverlapColumnEndSession_Click()
  Me.Changed = True
End Sub

Private Sub cboOverlapColumnStartDate_Click()
  txtOverlapMessage.Text = GenerateMessage
  Me.Changed = True
End Sub

Private Sub cboOverlapColumnStartSession_Click()
  Me.Changed = True
End Sub

Private Sub cboOverlapColumnType_Click()
  Me.Changed = True
End Sub

Private Sub Form_Load()

  miValidationType = VALIDATION_OVERLAP
  mbCancelled = True
  mbReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

End Sub


Private Sub cmdCancel_Click()

  If Me.Changed Then
    Select Case MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
    Case vbYes
      cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide

End Sub

Private Sub cmdOK_Click()

  If ValidDefinition = False Then
    Exit Sub
  End If

  SaveDefinition
  mbCancelled = False
  Me.Hide

End Sub

Private Function ValidDefinition() As Boolean
  
  Dim sValidationMessage As String
  
  ValidDefinition = True
  sValidationMessage = vbNullString
  
  Select Case miValidationType
    Case VALIDATION_MANDATORY
    Case VALIDATION_UNIQUE
    Case VALIDATION_DUPLICATE
    Case VALIDATION_OVERLAP
    
      If GetComboItem(cboOverlapColumnStartDate) = 0 Then
        sValidationMessage = sValidationMessage & "Overlap Event Start is not defined" & vbNewLine
      End If
    
      If GetComboItem(cboOverlapColumnEndDate) = 0 Then
        sValidationMessage = sValidationMessage & "Overlap Event End is not defined" & vbNewLine
      End If
    
    Case VALIDATION_CUSTOM
        
  End Select
  
  If Len(sValidationMessage) > 0 Then
    MsgBox sValidationMessage, vbInformation, "Validation"
    ValidDefinition = False
  End If
  
End Function

Private Sub SaveDefinition()

  objValidationObject.ValidationType = miValidationType
  objValidationObject.EventStartdateColumnID = GetComboItem(cboOverlapColumnStartDate)
  objValidationObject.EventStartSessionColumnID = GetComboItem(cboOverlapColumnStartSession)
  objValidationObject.EventEnddateColumnID = GetComboItem(cboOverlapColumnEndDate)
  objValidationObject.EventEndSessionColumnID = GetComboItem(cboOverlapColumnEndSession)
  objValidationObject.EventTypeColumnID = GetComboItem(cboOverlapColumnType)
  objValidationObject.Message = txtOverlapMessage.Text

End Sub

Public Function PopulateControls() As Boolean

  On Error GoTo ErrorTrap

  Dim iCount As Integer
  Dim bOK As Boolean
  Dim lngTableID As Long

  bOK = True
  lngTableID = objValidationObject.TableID
  mbLoading = True

  ' Clear existing definitions
  cboOverlapColumnStartDate.Clear
  cboOverlapColumnStartSession.Clear
  cboOverlapColumnEndDate.Clear
  cboOverlapColumnEndSession.Clear
  cboOverlapColumnType.Clear
 
  PopulateComboWithColumns cboOverlapColumnStartDate, lngTableID, False, dtTIMESTAMP, True, objValidationObject.EventStartdateColumnID
  PopulateComboWithColumns cboOverlapColumnStartSession, lngTableID, True, dtVARCHAR, False, -1
  PopulateComboWithColumns cboOverlapColumnEndDate, lngTableID, False, dtTIMESTAMP, False, -1
  PopulateComboWithColumns cboOverlapColumnEndSession, lngTableID, True, dtVARCHAR, False, -1
  PopulateComboWithColumns cboOverlapColumnType, lngTableID, True, dtVARCHAR, False, -1
  
  SetComboItem cboOverlapColumnStartDate, objValidationObject.EventStartdateColumnID
  SetComboItem cboOverlapColumnStartSession, objValidationObject.EventStartSessionColumnID
  SetComboItem cboOverlapColumnEndDate, objValidationObject.EventEnddateColumnID
  SetComboItem cboOverlapColumnEndSession, objValidationObject.EventEndSessionColumnID
  SetComboItem cboOverlapColumnType, objValidationObject.EventTypeColumnID
  
  txtOverlapMessage.Text = objValidationObject.Message
  
TidyUpAndExit:
  mbLoading = False
  PopulateControls = bOK
  Exit Function
  
ErrorTrap:
  bOK = False

End Function


Private Sub PopulateComboWithColumns(ByRef cboTemp As ComboBox, ByVal plngTableID As Long, ByVal AllowNone As Boolean, ByVal DataType As DataTypes, IsMandatory As Boolean, ByVal plngAddCurrent As Long)

  If AllowNone Then
    With cboTemp
      .Clear
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End With
  End If


  With recColEdit
    .Index = "idxName"
    .Seek ">=", plngTableID

    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> plngTableID Then
          Exit Do
        End If

        If (Not !Deleted) And (!DataType = DataType) And (!Mandatory = IsMandatory Or Not IsMandatory Or plngAddCurrent = !ColumnID) Then

          cboTemp.AddItem (!ColumnName)
          cboTemp.ItemData(cboTemp.NewIndex) = !ColumnID

        End If

        .MoveNext
      Loop
    End If
  End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    mbCancelled = True
  End If

End Sub

Private Sub optType_Click(Index As Integer)
  miValidationType = Index
End Sub

Private Function GenerateMessage() As String
  GenerateMessage = cboOverlapColumnStartDate.Text & " and " & cboOverlapColumnEndDate.Text & " overlaps with another record."
End Function

Private Sub txtOverlapMessage_Click()
  Me.Changed = True
End Sub
