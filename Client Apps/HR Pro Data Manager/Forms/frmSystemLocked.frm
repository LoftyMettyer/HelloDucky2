VERSION 5.00
Begin VB.Form frmSystemLocked 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Locked"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1003
   Icon            =   "frmSystemLocked.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLockMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HR Pro is locked until the external OLE/photo editing application is terminated."
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   3810
   End
End
Attribute VB_Name = "frmSystemLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miLockType As HRProDataMgr.LockType
Private mlngProcessID As Long
Private mbIsFileHandleOK As Boolean

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const STILL_ACTIVE = &H103

Public Property Let LockType(ByVal piNewValue As HRProDataMgr.LockType)
  ' Set the lock type.
  miLockType = piNewValue

  Select Case miLockType
    Case giLOCKTYPE_PHOTO
      lblLockMessage.Caption = "HR Pro is locked until the photo editing application is terminated."
    
    Case giLOCKTYPE_OLE
      lblLockMessage.Caption = "HR Pro is locked until the OLE document application is terminated."
      
    Case giLOCKTYPE_CRYSTAL
      lblLockMessage.Caption = "HR Pro is locked until the Crystal Reports application is terminated."
      
    Case Else
      Unload Me
  End Select
  
End Property

Public Property Let ProcessID(ByVal plngNewValue As Long)
  
  mlngProcessID = plngNewValue

End Property

'Public Property Set OLEControl(ByVal pctlNewValue As VB.Control)
'  LockType = giLOCKTYPE_OLE
'  Set mctlOLEControl = pctlNewValue
'
'End Property


Private Sub Form_Activate()
  
  Dim retVAL As Long
  Dim strError As String
  Dim iAppTrap As Integer
  
  mbIsFileHandleOK = True
  iAppTrap = 0
  
  If miLockType = giLOCKTYPE_OLE Or miLockType = giLOCKTYPE_PHOTO Then
    
    If mlngProcessID = 0 Then
      mbIsFileHandleOK = False
      Unload Me
    Else
     
      'Loop while the process is active
      Do
        iAppTrap = IIf(iAppTrap < 2, iAppTrap + 1, iAppTrap)
  
        'Get the status of the process
        GetExitCodeProcess mlngProcessID, retVAL
        DoEvents: Sleep 100

      Loop While retVAL = STILL_ACTIVE
      
      mbIsFileHandleOK = IIf(iAppTrap > 1, True, False)
      Unload Me
      
    End If
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

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub

' Was the file handle successfully grabbed
Public Property Get IsFileHandleOK() As Boolean
  IsFileHandleOK = mbIsFileHandleOK
End Property



