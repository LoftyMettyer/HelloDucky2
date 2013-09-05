VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmEmailSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Email Recipients"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1034
   Icon            =   "frmEmailSel.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4200
      TabIndex        =   2
      Top             =   3665
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4200
      TabIndex        =   1
      Top             =   3165
      Width           =   1200
   End
   Begin VB.CheckBox chkIncRecDesc 
      Caption         =   "Include Record Description in message"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4140
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4815
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmEmailSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function MAPISignon() As Integer
  'Begin a MAPI session
  Screen.MousePointer = 11
  On Error Resume Next
  MAPISignon = True
  If MAPISession1.SessionID = 0 Then
    'No session currently exists
    MAPISession1.DownLoadMail = False
    MAPISession1.SignOn
    If Err > 0 Then
      If Err <> 32001 And Err <> 32003 Then
        'MH20020820 Fault 4317
        'MsgBox Error$, 48, "Mail Error"
        MsgBox "Email not configured correctly." & vbCrLf & _
               IIf(Err.Description <> vbNullString, "(" & Trim(Err.Description) & ")", ""), _
               vbExclamation, "Mail Error"
      End If
      MAPISignon = False
    Else
      MAPIMessages1.SessionID = MAPISession1.SessionID
    End If
  End If
  Screen.MousePointer = 0
End Function

Function MAPIsignoff() As Integer
  'End a MAPI session
  Screen.MousePointer = 11
  On Error Resume Next
  MAPIsignoff = True
  If MAPISession1.SessionID <> 0 Then
    'Session currently exists
    MAPISession1.SignOff
    If Err > 0 Then
      MsgBox Error$, 48, "Mail Error"
      MAPIsignoff = False
    Else
      MAPIMessages1.SessionID = 0
    End If
  End If
  Screen.MousePointer = 0
End Function



