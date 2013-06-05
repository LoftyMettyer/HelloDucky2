VERSION 5.00
Begin VB.Form frmEmailDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Address Definition"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1072
   Icon            =   "frmEmailDef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4780
      Begin VB.TextBox txtEmailAddress 
         Height          =   315
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   4
         Top             =   700
         Width           =   2955
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   3
         Top             =   300
         Width           =   2955
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address :"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   2
         Top             =   765
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3735
      TabIndex        =   6
      Top             =   1455
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2490
      TabIndex        =   5
      Top             =   1455
      Width           =   1200
   End
End
Attribute VB_Name = "frmEmailDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSelectedID As Long
Private mblnCancelled As Boolean

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngSelectedID
End Property

Public Sub Initialise(blnNew As Boolean, blnCopy As Boolean, Optional lngSelectedID As Long)

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  If Not blnNew And lngSelectedID > 0 Then
    mlngSelectedID = lngSelectedID
    strSQL = "SELECT * FROM ASRSysEmailAddress WHERE EmailID = " & CStr(mlngSelectedID)
    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
    If rsTemp.BOF And rsTemp.EOF Then
      COAMsgBox "This Email definition has been deleted by another user.", vbCritical, "Email Definition"
      mlngSelectedID = 0
      Exit Sub
    End If
  
    txtName.Text = IIf(blnCopy, "Copy of ", "") & rsTemp!Name
    txtEmailAddress.Text = rsTemp!Fixed
    txtEmailAddress.Tag = rsTemp!Fixed

    Changed = blnCopy
    If blnCopy Then
      mlngSelectedID = 0
    End If
  Else
    Changed = False
  End If

End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()

  If Trim(txtName.Text) = vbNullString Then
    COAMsgBox "You must give this definition a name", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Sub
  End If

  If Trim(txtEmailAddress.Text) = vbNullString Then
    COAMsgBox "You must enter an email address", vbExclamation, Me.Caption
    txtEmailAddress.SetFocus
    Exit Sub
  End If

'  If InStr(txtEmailAddress.Text, "'") > 0 Then
'    COAMsgBox "The email address cannot contain apostraphes.", vbExclamation, Me.Caption
'    txtEmailAddress.SetFocus
'    Exit Sub
'  End If

  If UniqueName(txtName.Text) = False Then
    COAMsgBox "An email address called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Sub
  End If

  If mlngSelectedID > 0 Then
    'Check if the email address has changed...
    If txtEmailAddress.Text <> txtEmailAddress.Tag Then
      If CheckForUsage = False Then
        Exit Sub
      End If
    End If
  End If

  SaveDefinition
  Me.Hide
  mblnCancelled = False

End Sub

Private Sub Form_Activate()
  mblnCancelled = True
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


Private Sub txtEmailAddress_Change()
  Changed = True
End Sub

Private Sub txtEmailAddress_GotFocus()
  With txtEmailAddress
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtName_Change()
  Changed = True
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub SaveDefinition()
  
  Dim objInsUpd As clsDataInsertUpdate
  Dim blnNew As Boolean

  blnNew = (mlngSelectedID = 0)

  Set objInsUpd = New clsDataInsertUpdate

  objInsUpd.AddColumn "TableID", 0
  objInsUpd.AddColumn "Name", txtName.Text, True
  objInsUpd.AddColumn "Fixed", txtEmailAddress.Text, True
  objInsUpd.AddColumn "Type", 0
  objInsUpd.AddColumn "ColumnID", 0
  objInsUpd.AddColumn "ExprID", 0

  If blnNew Then
'    objInsUpd.AddColumn "EmailID", CStr(UniqueColumnValue("ASRSysEmailAddress", "EmailID"))
    objInsUpd.AddColumn "EmailID", CStr(GetUniqueID("emailaddress", "ASRSysEmailAddress", "EmailID"))
  End If

  mlngSelectedID = objInsUpd.InsertUpdate("ASRSysEmailAddress", "EmailID", mlngSelectedID)

  Set objInsUpd = Nothing

  If blnNew Then
    Call UtilCreated(utlEmailAddress, mlngSelectedID)
  Else
    Call UtilUpdateLastSaved(utlEmailAddress, mlngSelectedID)
  End If

End Sub

Public Sub PrintDef(lngEmailGroupID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim strSQL As String

  strSQL = "SELECT * FROM ASRSysEmailAddress WHERE EmailID = " & CStr(lngEmailGroupID)
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This Email definition has been deleted by another user.", vbCritical, "Email Definition"
    Exit Sub
  End If


  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      
      If .PrintStart(False) Then
      
        .PrintHeader "Email Address : " & rsTemp!Name
        .PrintNormal "Name : " & rsTemp!Name
        .PrintNormal "Email Address : " & rsTemp!Fixed
        .PrintEnd
        .PrintConfirm "Email Address : " & rsTemp!Name, "Email Address Definition"
    
      End If
    End With
  
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Cross Tab Definition Failed", vbCritical, "Email Address Definition"

End Sub


Private Function CheckForUsage() As Boolean

  Dim frmProp As frmDefProp
  Dim strMBText As String
  Dim lngCount As Long

  CheckForUsage = True
  
  Set frmProp = New frmDefProp
  frmProp.CheckForUseage "EMAIL ADDRESS", mlngSelectedID

  With frmProp.List1
    If .List(0) <> "<None>" And _
       .List(0) <> "<Error Checking Usage>" Then
  
      strMBText = "Updating this email address will effect the following definitions:" & vbCrLf & vbCrLf
      For lngCount = 0 To .ListCount
        strMBText = strMBText & .List(lngCount) & vbCrLf
      Next
      strMBText = strMBText & vbCrLf & "Continue anyway?"
      CheckForUsage = (COAMsgBox(strMBText, vbExclamation + vbYesNo, Me.Caption) = vbYes)
  
    End If
  End With

  Set frmProp = Nothing

End Function

Private Function UniqueName(sName As String) As Boolean

  Dim rsName As Recordset
  Dim sSQL As String
    
  sSQL = "SELECT * FROM ASRSysEmailAddress" & _
         " WHERE Name = '" & Replace(sName, "'", "''") & "' AND EmailID <> " & CStr(mlngSelectedID)
    
  Set rsName = datGeneral.GetReadOnlyRecords(sSQL)
  UniqueName = (rsName.BOF And rsName.EOF)
  rsName.Close
    
  Set rsName = Nothing

End Function


