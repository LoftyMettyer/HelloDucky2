VERSION 5.00
Begin VB.Form frmSystemLocked 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intranet Locked"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLockMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HR Pro is locked until the file editing application is terminated."
      Height          =   495
      Left            =   200
      TabIndex        =   0
      Top             =   160
      Width           =   3495
   End
End
Attribute VB_Name = "frmSystemLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''Const giLOCKTYPE_PHOTO = 1
'''Const giLOCKTYPE_OLE = 2
'''Const giLOCKTYPE_CRYSTAL = 4

'''Private miLockType As Integer
Private mlngProcessID As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const STILL_ACTIVE = &H103

Private Sub Form_Activate()
  Dim retVAL As Long
  
  Do
    'Get the status of the process
    GetExitCodeProcess mlngProcessID, retVAL
  
    DoEvents: Sleep 100

    'Loop while the process is active
  Loop While retVAL = STILL_ACTIVE
  
  Unload Me

End Sub


Public Property Let ProcessID(ByVal plngNewValue As Long)
  
'''  LockType = giLOCKTYPE_OLE
  mlngProcessID = plngNewValue

End Property

