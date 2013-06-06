VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   2205
   ClientTop       =   2655
   ClientWidth     =   6135
   ClipControls    =   0   'False
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
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   873
      Calculator      =   "frmSplash.frx":038A
      Caption         =   "frmSplash.frx":03AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSplash.frx":0410
      Keys            =   "frmSplash.frx":042E
      Spin            =   "frmSplash.frx":0478
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2086338565
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBMask6Ctl.TDBMask TDBMask1 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "frmSplash.frx":04A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmSplash.frx":0506
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "&&&&&&&&&&"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "TDBMask1__"
      Value           =   "TDBMask1"
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   0
      Picture         =   "frmSplash.frx":0548
      Top             =   0
      Width           =   6150
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
 '*********************************************************************
  '  Well you'll never believe this one but we need to load the APEX
  '  controls on the form otherwise the codejock stuff doesnt always
  '  get applied to them in recedit!?!?  ' AE20090728
  '*********************************************************************
  
  
  'Image1.Move 60, 60, Me.ScaleWidth - 120, lblCopyRight.Top - 150
    
'  With lblVersion
'    .Left = 0
'    .Width = Me.ScaleWidth
'    .Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'  End With
'
'  With Line2
'    .X1 = 75
'    .X2 = Me.ScaleWidth - 150
'  End With
'
'  With lblCopyRight
'    .Left = 50
'    .Width = Me.ScaleWidth - 100
'    .Caption = "Copyright © 1997-" & Format(Date, "yyyy") & ", COA Solutions Ltd"
'  End With
'
'  With lblDisclaimer
'    .Left = 100
'    .Width = Me.ScaleWidth - 200
'  End With
  
  UI.frmAtCenter Me
  
End Sub

