VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.0#0"; "COA_Line.ocx"
Begin VB.Form frmWorkflowStepDetails 
   Caption         =   "Workflow Step Details"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1144
   Icon            =   "frmWorkflowStepDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picResizer 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraButtons 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   3840
      TabIndex        =   43
      Top             =   8220
      Width           =   4000
      Begin VB.CommandButton cmdOpenWebForm 
         Caption         =   "O&pen"
         Height          =   400
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdResendEmail 
         Caption         =   "&Resend"
         Height          =   400
         Left            =   1400
         TabIndex        =   45
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   2800
         TabIndex        =   46
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   8020
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7770
      Begin VB.Frame fraSpecificDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   2200
         Index           =   5
         Left            =   100
         TabIndex        =   38
         Top             =   5720
         Width           =   7550
         Begin VB.TextBox txtStoredDataMessage 
            BackColor       =   &H8000000F&
            Height          =   800
            Left            =   1250
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   40
            Top             =   250
            Width           =   6200
         End
         Begin SSDataWidgets_B.SSDBGrid grdStoredDataValues 
            Height          =   1000
            Left            =   1250
            TabIndex        =   42
            Top             =   1130
            Width           =   6200
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   2
            AllowUpdate     =   0   'False
            AllowGroupSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowColumnMoving=   0
            AllowGroupSwapping=   0   'False
            AllowColumnSwapping=   0
            AllowGroupShrinking=   0   'False
            AllowColumnShrinking=   0   'False
            AllowDragDrop   =   0   'False
            SelectTypeCol   =   0
            SelectTypeRow   =   0
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            CellNavigation  =   1
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3200
            Columns(0).Caption=   "Column"
            Columns(0).Name =   "Column"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Value"
            Columns(1).Name =   "Value"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   10936
            _ExtentY        =   1764
            _StockProps     =   79
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin COALine.COA_Line linStoredData 
            Height          =   30
            Left            =   100
            Top             =   100
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   53
         End
         Begin VB.Label lblStoredDataMessage 
            BackStyle       =   0  'Transparent
            Caption         =   "Message :"
            Height          =   195
            Left            =   105
            TabIndex        =   39
            Top             =   315
            Width           =   915
         End
         Begin VB.Label lblStoredDataValues 
            BackStyle       =   0  'Transparent
            Caption         =   "Values :"
            Height          =   195
            Left            =   105
            TabIndex        =   41
            Top             =   1185
            Width           =   795
         End
      End
      Begin VB.Frame fraSpecificDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   1600
         Index           =   2
         Left            =   100
         TabIndex        =   32
         Top             =   4030
         Width           =   7550
         Begin VB.TextBox txtWebFormMessage 
            BackColor       =   &H8000000F&
            Height          =   1000
            Left            =   1350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   37
            Top             =   650
            Visible         =   0   'False
            Width           =   6200
         End
         Begin SSDataWidgets_B.SSDBGrid grdWebFormEnteredValues 
            Height          =   1000
            Left            =   1250
            TabIndex        =   36
            Top             =   550
            Width           =   6200
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            AllowUpdate     =   0   'False
            AllowGroupSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowColumnMoving=   0
            AllowGroupSwapping=   0   'False
            AllowColumnSwapping=   0
            AllowGroupShrinking=   0   'False
            AllowColumnShrinking=   0   'False
            AllowDragDrop   =   0   'False
            SelectTypeCol   =   0
            SelectTypeRow   =   0
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            CellNavigation  =   1
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   185
            Columns.Count   =   3
            Columns(0).Width=   3519
            Columns(0).Caption=   "Identifier"
            Columns(0).Name =   "Identifier"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2646
            Columns(1).Caption=   "Type"
            Columns(1).Name =   "Type"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   4710
            Columns(2).Caption=   "Value"
            Columns(2).Name =   "Value"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   10936
            _ExtentY        =   1764
            _StockProps     =   79
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin COALine.COA_Line linWebForm 
            Height          =   30
            Left            =   100
            Top             =   100
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   53
         End
         Begin VB.Label lblWebFormEnteredValues 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Values :"
            Height          =   195
            Left            =   105
            TabIndex        =   35
            Top             =   610
            Width           =   570
         End
         Begin VB.Label lblWebFormUser 
            BackStyle       =   0  'Transparent
            Caption         =   "User :"
            Height          =   195
            Left            =   105
            TabIndex        =   33
            Top             =   255
            Width           =   615
         End
         Begin VB.Label lblWebFormUserValue 
            BackStyle       =   0  'Transparent
            Caption         =   "WebFormUserValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1245
            TabIndex        =   34
            Top             =   250
            Width           =   6000
         End
      End
      Begin VB.Frame fraSpecificDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   1960
         Index           =   3
         Left            =   100
         TabIndex        =   23
         Top             =   1960
         Width           =   7550
         Begin VB.TextBox txtEmailCCValue 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   610
            Width           =   4575
         End
         Begin VB.TextBox txtEmailAddressValue 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   250
            Width           =   4575
         End
         Begin VB.TextBox txtEmailMessage 
            BackColor       =   &H8000000F&
            Height          =   1000
            Left            =   1335
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   31
            Top             =   1270
            Width           =   6015
         End
         Begin COALine.COA_Line linEmail 
            Height          =   30
            Left            =   100
            Top             =   100
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   53
         End
         Begin VB.Label lblEmailCc 
            BackStyle       =   0  'Transparent
            Caption         =   "Email Copy :"
            Height          =   195
            Left            =   105
            TabIndex        =   26
            Top             =   615
            Width           =   1110
         End
         Begin VB.Label lblEmailSubjectValue 
            BackStyle       =   0  'Transparent
            Caption         =   "EmailSubjectValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1380
            TabIndex        =   29
            Top             =   975
            Width           =   5820
         End
         Begin VB.Label lblEmailSubject 
            BackStyle       =   0  'Transparent
            Caption         =   "Email Subject :"
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   975
            Width           =   1230
         End
         Begin VB.Label lblEmailMessage 
            BackStyle       =   0  'Transparent
            Caption         =   "Message :"
            Height          =   195
            Left            =   100
            TabIndex        =   30
            Top             =   1330
            Width           =   735
         End
         Begin VB.Label lblEmailAddress 
            BackStyle       =   0  'Transparent
            Caption         =   "Email To :"
            Height          =   195
            Left            =   105
            TabIndex        =   24
            Top             =   255
            Width           =   915
         End
      End
      Begin VB.Frame fraSpecificDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   470
         Index           =   4
         Left            =   100
         TabIndex        =   20
         Top             =   1410
         Width           =   7550
         Begin COALine.COA_Line linDecision 
            Height          =   30
            Left            =   100
            Top             =   100
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   53
         End
         Begin VB.Label lblDecisionFlowValue 
            BackStyle       =   0  'Transparent
            Caption         =   "DecisionFlowValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1425
            TabIndex        =   22
            Top             =   255
            Width           =   5775
         End
         Begin VB.Label lblDecisionFlow 
            BackStyle       =   0  'Transparent
            Caption         =   "Decision Flow :"
            Height          =   195
            Left            =   105
            TabIndex        =   21
            Top             =   255
            Width           =   1350
         End
      End
      Begin VB.Frame fraBasicDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   1060
         Left            =   100
         TabIndex        =   1
         Top             =   240
         Width           =   7550
         Begin VB.Label lblCompletionCount 
            BackStyle       =   0  'Transparent
            Caption         =   "Completions :"
            Height          =   195
            Left            =   105
            TabIndex        =   14
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblCompletionCountValue 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1425
            TabIndex        =   15
            Top             =   840
            Width           =   105
         End
         Begin VB.Label lblFailureCount 
            BackStyle       =   0  'Transparent
            Caption         =   "Failures :"
            Height          =   195
            Left            =   2805
            TabIndex        =   16
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lblFailureCountValue 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3885
            TabIndex        =   17
            Top             =   840
            Width           =   90
         End
         Begin VB.Label lblTimeoutCount 
            BackStyle       =   0  'Transparent
            Caption         =   "Timeouts :"
            Height          =   195
            Left            =   5400
            TabIndex        =   18
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblTimeoutCountValue 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6330
            TabIndex        =   19
            Top             =   840
            Width           =   90
         End
         Begin VB.Label lblDurationValue 
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6330
            TabIndex        =   13
            Top             =   480
            Width           =   1000
         End
         Begin VB.Label lblDuration 
            BackStyle       =   0  'Transparent
            Caption         =   "Duration :"
            Height          =   195
            Left            =   5400
            TabIndex        =   12
            Top             =   480
            Width           =   900
         End
         Begin VB.Label lblEndTimeValue 
            BackStyle       =   0  'Transparent
            Caption         =   "99/99/9999  00:00"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3885
            TabIndex        =   11
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblEndTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Completed :"
            Height          =   240
            Left            =   2805
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblStartTimeValue 
            BackStyle       =   0  'Transparent
            Caption         =   "99/99/9999  00:00"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1425
            TabIndex        =   9
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblStartTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Activated :"
            Height          =   195
            Left            =   105
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblStatusValue 
            BackStyle       =   0  'Transparent
            Caption         =   "StatusValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6330
            TabIndex        =   7
            Top             =   120
            Width           =   1000
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Status :"
            Height          =   195
            Left            =   5400
            TabIndex        =   6
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblCaptionValue 
            BackStyle       =   0  'Transparent
            Caption         =   "CaptionValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3885
            TabIndex        =   5
            Top             =   120
            Width           =   945
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Caption :"
            Height          =   195
            Left            =   2805
            TabIndex        =   4
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblElementType 
            BackStyle       =   0  'Transparent
            Caption         =   "Element Type :"
            Height          =   195
            Left            =   105
            TabIndex        =   2
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label lblElementTypeValue 
            BackStyle       =   0  'Transparent
            Caption         =   "ElementTypeValue"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1425
            TabIndex        =   3
            Top             =   120
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmWorkflowStepDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SCROLLBAR_WIDTH = 250
Private Const MINFORMWIDTH = 9750
Private Const GAPOVERBUTTONS = 80
Private Const GAPUNDERBUTTONS = 620

Private Const MINHEIGHT_EMAIL = 4800
Private Const MINHEIGHT_WEBFORM = 3900
Private Const MINHEIGHT_STOREDDATA = 4500

Private Const COLUMN1LABELLEFT = 120
Private Const COLUMN1LABELWIDTH = 1320
Private Const COLUMN2LABELWIDTH = 1100
Private Const COLUMN3LABELWIDTH = 940
Private Const GAPBETWEENCOLUMNS = 320

Private Const BUTTON_GAP = 260
Private Const BUTTON_WIDTH = 1400
Private Const BUTTON_HEIGHT = 400

' Flag to store if we are currently resizing the form
Private mblnSizing As Boolean

Private mlngWorkflowStepID As Long
Private mlngWorkflowInstanceID As Long
Private mlngWorkflowElementID As Long
Private mlngElementType As ElementType
Private miStatus As Integer
Private mfAdministrator As Boolean
Private malngHypertextLinkedSteps() As Long

Private mfrmWorkflowLog As frmWorkflowLogDetails
Private mfSizeable As Boolean
Private msngWebFormValueColumnWidth As Single
Private msngStoredDataValueColumnWidth As Single

Private masEmailAddresses() As String
Private masActionedUserNames() As String
Private masActionedDelegatedEmailAddresses() As String

Private masngFormDimensions() As Single

Public Enum WorkflowWebFormItemTypes
  giWFFORMITEM_UNKNOWN = -1
  giWFFORMITEM_BUTTON = 0
  giWFFORMITEM_DBVALUE = 1
  giWFFORMITEM_LABEL = 2
  giWFFORMITEM_INPUTVALUE_CHAR = 3
  giWFFORMITEM_WFVALUE = 4
  giWFFORMITEM_INPUTVALUE_NUMERIC = 5
  giWFFORMITEM_INPUTVALUE_LOGIC = 6
  giWFFORMITEM_INPUTVALUE_DATE = 7
  giWFFORMITEM_FRAME = 8
  giWFFORMITEM_LINE = 9
  giWFFORMITEM_IMAGE = 10
  giWFFORMITEM_INPUTVALUE_GRID = 11
  giWFFORMITEM_FORMATCODE = 12 ' NB. Only used in emails.
  giWFFORMITEM_INPUTVALUE_DROPDOWN = 13
  giWFFORMITEM_INPUTVALUE_LOOKUP = 14
  giWFFORMITEM_INPUTVALUE_OPTIONGROUP = 15
  giWFFORMITEM_CALC = 16
  giWFFORMITEM_INPUTVALUE_FILEUPLOAD = 17
  giWFFORMITEM_FILEATTACHMENT = 18
  giWFFORMITEM_DBFILE = 19
  giWFFORMITEM_WFFILE = 20
End Enum

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Const EM_GETLINECOUNT  As Long = &HBA

Private Function DoHeaderInfo() As Boolean

  ' Populate all the relevant header fields and position the controls
  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim rstTemp As Recordset
  Dim rsAddresses As ADODB.Recordset
  Dim rsPendingSteps As ADODB.Recordset
  Dim sDateFormat As String
  Dim fraTemp As Frame
  Dim sTemp As String
  Dim sTemp2 As String
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim fFound As Boolean
  Dim fCanOpenWebForm As Boolean
  Dim iIndex As Integer
  Dim fRequired As Boolean
  Dim sngYMove As Single
  Dim sHypertextLinkedSteps As String
  
  sDateFormat = DateFormat
  mfSizeable = True
  
  ReDim masEmailAddresses(0)
  ReDim masActionedUserNames(0)
  ReDim masActionedDelegatedEmailAddresses(0)
  ReDim malngHypertextLinkedSteps(0)
  
  If Not mfAdministrator Then
    ' The current user does NOT have permission to Administer workflows.
    ' Only display Email message of the email was sent the the current user.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('spASRGetWorkflowEmailAddresses')" & _
      "     AND sysstat & 0xf = 4"
    Set rsAddresses = datGeneral.GetRecords(sSQL)

    iCount = rsAddresses!objectCount
    rsAddresses.Close
    Set rsAddresses = Nothing
          
    If iCount > 0 Then
      sSQL = "exec spASRGetWorkflowEmailAddresses"
      Set rsAddresses = datGeneral.GetReadOnlyRecords(sSQL)

      With rsAddresses
        Do While Not .EOF
          If Len(Trim(IIf(IsNull(!address), "", !address))) > 0 Then
            ReDim Preserve masEmailAddresses(UBound(masEmailAddresses) + 1)
            masEmailAddresses(UBound(masEmailAddresses)) = UCase(Trim(!address))
          End If
          
          .MoveNext
        Loop

        .Close
      End With

      Set rsAddresses = Nothing
    End If
  End If
  
  ' Get the delegated email addresses (if any)
  sSQL = "SELECT DISTINCT ASRSysWorkflowStepDelegation.delegateEmail" & _
    " FROM ASRSysWorkflowStepDelegation" & _
    " WHERE ASRSysWorkflowStepDelegation.stepID = " & CStr(mlngWorkflowStepID)
  Set rstTemp = datGeneral.GetReadOnlyRecords(sSQL)
  
  With rstTemp
    Do While Not .EOF
      If Len(Trim(IIf(IsNull(!delegateEmail), "", !delegateEmail))) > 0 Then
        ReDim Preserve masActionedDelegatedEmailAddresses(UBound(masActionedDelegatedEmailAddresses) + 1)
        masActionedDelegatedEmailAddresses(UBound(masActionedDelegatedEmailAddresses)) = Trim(!delegateEmail)
      End If
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rstTemp = Nothing

  sSQL = "SELECT ASRSysWorkflowInstanceSteps.status," & _
    "   ASRSysWorkflowInstanceSteps.instanceID," & _
    "   ASRSysWorkflowInstanceSteps.elementID," & _
    "   ASRSysWorkflowInstanceSteps.activationDateTime," & _
    "   isnull(ASRSysWorkflowInstanceSteps.message, '') AS [message]," & _
    "   ASRSysWorkflowInstanceSteps.completionDateTime," & _
    "   isnull(ASRSysWorkflowInstanceSteps.userEmail, '') AS [userEmail]," & _
    "   isnull(ASRSysWorkflowElements.emailSubject, '') AS [emailSubject]," & _
    "   isnull(ASRSysWorkflowInstanceSteps.userName, '') AS [userName]," & _
    "   ASRSysWorkflowElements.type," & _
    "   CASE" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 0 THEN 'Begin'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 1 THEN 'Terminator'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 2 THEN 'Web Form'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 3 THEN 'Email'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 4 THEN 'Decision'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 5 THEN 'Stored Data'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 6 THEN 'And'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 7 THEN 'Or'" & _
    "   END AS [typeDesc]," & _
    "   CASE WHEN [ASRSysWorkflowInstanceSteps].[activationDateTime] IS null OR [ASRSysWorkflowInstanceSteps].[completionDateTime] IS null THEN -1" & _
    "    ELSE datediff(s, [ASRSysWorkflowInstanceSteps].[activationDateTime], [ASRSysWorkflowInstanceSteps].[completionDateTime])" & _
    "   END AS [Duration],"

  sSQL = sSQL & _
    "   CASE" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_ONHOLD) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_ONHOLD) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGENGINEACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGENGINEACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGUSERACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_COMPLETED) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_COMPLETED) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_FAILED) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILED) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_FAILEDACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILEDACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_INPROGRESS) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_INPROGRESS) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_TIMEOUT) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_TIMEOUT) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGUSERCOMPLETION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERCOMPLETION) & "'" & _
    "   END AS [statusDesc]," & _
    "   isnull(ASRSysWorkflowElements.caption, '') AS [caption]," & _
    "   isnull(ASRSysWorkflowInstanceSteps.decisionFlow, 0) AS [decisionFlow]," & _
    "   ASRSysWorkflowElements.identifier," & _
    "   ASRSysWorkflowElements.dataAction," & _
    "   ASRSysWorkflowElements.dataTableID," & _
    "   ASRSysWorkflowElements.dataRecord," & _
    "   ASRSysWorkflowElements.emailID," & _
    "   ASRSysWorkflowElements.emailRecord," & _
    "   ASRSysWorkflowElements.trueFlowIdentifier," & _
    "   isnull([ASRSysWorkflowInstanceSteps].completionCount,0) AS [completionCount]," & _
    "   isnull([ASRSysWorkflowInstanceSteps].failedCount,0) AS [failedCount]," & _
    "   isnull([ASRSysWorkflowInstanceSteps].timeoutCount,0) AS [timeoutCount],"
    
  sSQL = sSQL & _
    "   ltrim(rtrim(isnull(ASRSysWorkflowInstanceSteps.emailCC, ''))) AS [emailCC]," & _
    "   isnull([ASRSysWorkflowInstanceSteps].hypertextLinkedSteps,'') AS [hypertextLinkedSteps]"

  sSQL = sSQL & _
    " FROM ASRSysWorkflowInstanceSteps" & _
    " INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & _
    " WHERE ASRSysWorkflowInstanceSteps.ID = " & CStr(mlngWorkflowStepID)
  Set rstTemp = datGeneral.GetReadOnlyRecords(sSQL)
  
  With rstTemp
    If (.EOF And .BOF) Then
      MsgBox "This record no longer exists in the workflow log.", vbExclamation + vbOKOnly, "Workflow Log"
      DoHeaderInfo = False
      GoTo TidyUpAndExit
    End If
  
    lblElementTypeValue.Caption = Replace(.Fields("typeDesc"), "&", "&&")
    lblCaptionValue.Caption = Replace(IIf(Len(.Fields("caption")) = 0, "<None>", .Fields("caption")), "&", "&&")
    lblStatusValue.Caption = Replace(.Fields("statusDesc"), "&", "&&")
    lblStartTimeValue.Caption = Replace(IIf(IsNull(.Fields("activationDateTime")), "<Not activated>", _
      Format(.Fields("activationDateTime"), sDateFormat & " hh:nn")), "&", "&&")
    lblEndTimeValue.Caption = Replace(IIf(IsNull(.Fields("completionDateTime")), "<Not completed>", _
      Format(.Fields("completionDateTime"), sDateFormat & " hh:nn")), "&", "&&")
    lblDurationValue.Caption = Replace(IIf(.Fields("Duration").Value = -1, "", FormatEventDuration(.Fields("Duration").Value)), "&", "&&")

    lblCompletionCountValue.Caption = .Fields("completionCount")
    lblFailureCountValue.Caption = .Fields("failedCount")
    lblTimeoutCountValue.Caption = .Fields("timeoutCount")

    mlngElementType = .Fields("type").Value
    miStatus = .Fields("status").Value
    mlngWorkflowInstanceID = .Fields("instanceID").Value
    mlngWorkflowElementID = .Fields("elementID").Value
    
    lblTimeoutCount.Visible = (mlngElementType = elem_WebForm)
    lblTimeoutCountValue.Visible = (mlngElementType = elem_WebForm)
    
    Select Case mlngElementType
      Case elem_WebForm
        sTemp = IIf(IsNull(.Fields("userEmail").Value), "", .Fields("userEmail").Value)
        If Len(Trim(sTemp)) = 0 Then
          sTemp = IIf(IsNull(.Fields("userName").Value), "", .Fields("userName").Value)
        End If
        
        'JPD 20070102 Fault 11750
        sTemp2 = UCase(Trim(sTemp))
        Do While Len(sTemp2) > 0
          iIndex = InStr(sTemp2, ";")
          
          If iIndex > 0 Then
            ReDim Preserve masActionedUserNames(UBound(masActionedUserNames) + 1)
            masActionedUserNames(UBound(masActionedUserNames)) = Trim(Left(sTemp2, iIndex - 1))
            sTemp2 = Trim(Mid(sTemp2, iIndex + 1))
          Else
            ReDim Preserve masActionedUserNames(UBound(masActionedUserNames) + 1)
            masActionedUserNames(UBound(masActionedUserNames)) = sTemp2
            sTemp2 = ""
          End If
        Loop

        If UBound(masActionedDelegatedEmailAddresses) > 0 Then
          sTemp = sTemp & " - delegated to "
          
          For iCount2 = 1 To UBound(masActionedDelegatedEmailAddresses)
            sTemp = sTemp & _
              IIf(iCount2 > 1, "; ", "") & _
              masActionedDelegatedEmailAddresses(iCount2)
          Next iCount2
        End If
        lblWebFormUserValue.Caption = Replace(sTemp, "&", "&&")

        If miStatus = giWFSTEPSTATUS_FAILED Then ' Failed - display the message not the values
          sTemp = .Fields("message").Value
          
          If mfAdministrator Then
            sTemp = Replace(Replace(sTemp, "&", "&&"), vbCrLf, vbTab)
            sTemp = Replace(sTemp, vbCr, vbTab)
            sTemp = Replace(sTemp, vbLf, vbTab)
            txtWebFormMessage.Text = Replace(sTemp, vbTab, vbNewLine)
          Else
            ' The current user does NOT have permission to Administer workflows.
            ' Do NOT display Stored Data message.
            txtWebFormMessage.Text = IIf(Len(Trim(sTemp)) = 0, "", "<You do not have permission to view the message>")
          End If
        
          lblWebFormEnteredValues.Caption = "Message :"
        Else
          lblWebFormEnteredValues.Caption = "Values :"
        End If
        
      Case elem_Email
        sTemp = .Fields("userEmail").Value
        
        'JPD 20070102 Fault 11750
        sTemp2 = UCase(Trim(sTemp))
        Do While Len(sTemp2) > 0
          iIndex = InStr(sTemp2, ";")
          
          If iIndex > 0 Then
            ReDim Preserve masActionedUserNames(UBound(masActionedUserNames) + 1)
            masActionedUserNames(UBound(masActionedUserNames)) = Trim(Left(sTemp2, iIndex - 1))
            sTemp2 = Trim(Mid(sTemp2, iIndex + 1))
          Else
            ReDim Preserve masActionedUserNames(UBound(masActionedUserNames) + 1)
            masActionedUserNames(UBound(masActionedUserNames)) = sTemp2
            sTemp2 = ""
          End If
        Loop
        
        If UBound(masActionedDelegatedEmailAddresses) > 0 Then
          sTemp = sTemp & " - delegated to "
          
          For iCount2 = 1 To UBound(masActionedDelegatedEmailAddresses)
            sTemp = sTemp & _
              IIf(iCount2 > 1, "; ", "") & _
              masActionedDelegatedEmailAddresses(iCount2)
          Next iCount2
        End If
        txtEmailAddressValue.Text = sTemp

        sngYMove = 0
        
        sTemp = .Fields("emailCC").Value
        txtEmailCCValue.Text = sTemp

        fRequired = (Len(sTemp) > 0)
        lblEmailCc.Visible = fRequired
        txtEmailCCValue.Visible = fRequired

        sTemp = .Fields("emailSubject").Value
        lblEmailSubjectValue.Caption = Replace(sTemp, "&", "&&")
        
        If mfAdministrator Then
          sTemp = .Fields("message").Value
          sTemp = Replace(Replace(sTemp, "&", "&&"), vbCrLf, vbTab)
          sTemp = Replace(sTemp, vbCr, vbTab)
          sTemp = Replace(sTemp, vbLf, vbTab)
          txtEmailMessage.Text = Replace(sTemp, vbTab, vbNewLine)
        Else
          fFound = False
          For iCount = 1 To UBound(masEmailAddresses)
            For iCount2 = 1 To UBound(masActionedUserNames)
              If (masActionedUserNames(iCount2) = masEmailAddresses(iCount)) Then
                fFound = True
                Exit For
              End If
            Next iCount2
          
            For iCount2 = 1 To UBound(masActionedDelegatedEmailAddresses)
              If (UCase(masActionedDelegatedEmailAddresses(iCount2)) = masEmailAddresses(iCount)) Then
                fFound = True
                Exit For
              End If
            Next iCount2
            
            If fFound Then
              Exit For
            End If
          Next iCount
          
          If fFound Then
            ' Email was for the current user.
            sTemp = .Fields("message").Value
            sTemp = Replace(Replace(sTemp, "&", "&&"), vbCrLf, vbTab)
            sTemp = Replace(sTemp, vbCr, vbTab)
            sTemp = Replace(sTemp, vbLf, vbTab)
            txtEmailMessage.Text = Replace(sTemp, vbTab, vbNewLine)
          Else
            txtEmailMessage.Text = "<You do not have permission to view the message>"
          End If
        End If
        
        sHypertextLinkedSteps = Trim(.Fields("hypertextLinkedSteps").Value)
        Do While Len(sHypertextLinkedSteps) > 0
          iIndex = InStr(sHypertextLinkedSteps, vbTab)
          
          If iIndex > 0 Then
            sTemp = Trim(Left(sHypertextLinkedSteps, iIndex - 1))
            sHypertextLinkedSteps = Trim(Mid(sHypertextLinkedSteps, iIndex + 1))
          Else
            sTemp = sHypertextLinkedSteps
            sHypertextLinkedSteps = ""
          End If
        
          If Val(sTemp) > 0 Then
            ReDim Preserve malngHypertextLinkedSteps(UBound(malngHypertextLinkedSteps) + 1)
            malngHypertextLinkedSteps(UBound(malngHypertextLinkedSteps)) = Val(sTemp)
          End If
        Loop
        
      Case elem_Decision
        mfSizeable = False
        sTemp = IIf(.Fields("decisionFlow").Value = 0, "False", "True")
        lblDecisionFlowValue.Caption = Replace(sTemp, "&", "&&")
        
      Case elem_StoredData
        sTemp = .Fields("message").Value
        
        If mfAdministrator Then
          sTemp = Replace(Replace(sTemp, "&", "&&"), vbCrLf, vbTab)
          sTemp = Replace(sTemp, vbCr, vbTab)
          sTemp = Replace(sTemp, vbLf, vbTab)
          txtStoredDataMessage.Text = Replace(sTemp, vbTab, vbNewLine)
        Else
          ' The current user does NOT have permission to Administer workflows.
          ' Do NOT display Stored Data message.
          txtStoredDataMessage.Text = IIf(Len(Trim(sTemp)) = 0, "", "<You do not have permission to view the message>")
        End If
  
      Case Else
        ' No specific details to be displayed.
        ' elem_Begin, elem_Terminator, elem_SummingJunction
        ' elem_Or, elem_Connector1, elem_Connector2
        mfSizeable = False
    End Select
  
    .Close
  End With
  Set rstTemp = Nothing
  
  For Each fraTemp In fraSpecificDetails
    fraTemp.Visible = (fraTemp.Index = mlngElementType)
  Next fraTemp
  Set fraTemp = Nothing
  
  cmdResendEmail.Visible = (mlngElementType = elem_Email)
  cmdResendEmail.Enabled = (miStatus = giWFSTEPSTATUS_COMPLETED) _
    Or (miStatus = giWFSTEPSTATUS_PENDINGENGINEACTION)
  
  cmdOpenWebForm.Visible = (mlngElementType = elem_WebForm)
  
  fCanOpenWebForm = (miStatus = giWFSTEPSTATUS_PENDINGUSERACTION) _
    Or (miStatus = giWFSTEPSTATUS_PENDINGUSERCOMPLETION)
    
  If fCanOpenWebForm And (Not mfAdministrator) Then
    ' The current user does NOT have permission to Administer workflows.
    ' Only allow them to open the web form if it was actioned to them.
    fCanOpenWebForm = False

    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('spASRCheckPendingWorkflowSteps')" & _
      "     AND sysstat & 0xf = 4"
    Set rsPendingSteps = datGeneral.GetRecords(sSQL)
  
    iCount = rsPendingSteps!objectCount
    rsPendingSteps.Close
    Set rsPendingSteps = Nothing
          
    If iCount > 0 Then
      sSQL = "exec spASRCheckPendingWorkflowSteps"
      Set rsPendingSteps = datGeneral.GetReadOnlyRecords(sSQL)
  
      With rsPendingSteps
        Do While Not .EOF
          If !ID = mlngWorkflowStepID Then
            fCanOpenWebForm = True
            Exit Do
          End If
          
          .MoveNext
        Loop
  
        .Close
      End With
  
      Set rsPendingSteps = Nothing
    End If
  End If
  
  cmdOpenWebForm.Enabled = fCanOpenWebForm
  
  DoHeaderInfo = True

TidyUpAndExit:
  Set rstTemp = Nothing
  Exit Function

ErrorTrap:

  MsgBox "Error whilst populating workflow step detail." & vbCrLf & "(" & Err.Description & ")"
  DoHeaderInfo = False

End Function


Public Function Initialise(plngInstanceStepID As Long, _
  pfrmWorkflowLog As frmWorkflowLogDetails) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  mlngWorkflowStepID = plngInstanceStepID
  Set mfrmWorkflowLog = pfrmWorkflowLog
  mfAdministrator = datGeneral.SystemPermission("WORKFLOW", "ADMINISTER")
  
  ' Let user know we are doing something, and dont redraw the form until the controls
  ' have all been repositioned
  Screen.MousePointer = vbHourglass
  Me.AutoRedraw = False
  
  ' Display the correct header information on the form
  If fOK Then fOK = DoHeaderInfo
  If fOK Then fOK = PopulateDetailsGrid

  If (mlngElementType >= LBound(masngFormDimensions, 2)) _
    And (mlngElementType <= UBound(masngFormDimensions, 2)) Then
  
    Me.Height = masngFormDimensions(0, mlngElementType)
    Me.Width = masngFormDimensions(1, mlngElementType)
  End If

  If fOK Then Form_Resize

  Dim sngMinFormHeight As Single
  Select Case mlngElementType
    Case elem_WebForm
      sngMinFormHeight = MINHEIGHT_WEBFORM

    Case elem_Email
      sngMinFormHeight = MINHEIGHT_EMAIL

    Case elem_StoredData
      sngMinFormHeight = MINHEIGHT_STOREDDATA

    Case Else
      sngMinFormHeight = fraButtons.Top _
         + fraButtons.Height _
         + GAPUNDERBUTTONS
  End Select
  Hook Me.hWnd, MINFORMWIDTH, CLng(sngMinFormHeight)

  ' Let user know we have finished, and can now redraw the form
  Screen.MousePointer = vbNormal
  Me.AutoRedraw = True

TidyUpAndExit:
  Initialise = fOK
  Exit Function

ErrorTrap:
  Initialise = False
  MsgBox "Error retrieving details for this workflow step." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Workflow Log"
  GoTo TidyUpAndExit
  
End Function

Private Function PopulateDetailsGrid() As Boolean
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rstTemp As Recordset
  Dim fOK As Boolean
  Dim sAddString As String
  Dim fShowDetails As Boolean
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim sTemp As String
  Dim iIndex As Integer
  
  fOK = True
  
  If mlngElementType = elem_WebForm Then
    grdWebFormEnteredValues.RemoveAll
    
    If (miStatus = giWFSTEPSTATUS_COMPLETED) _
      Or (miStatus = giWFSTEPSTATUS_PENDINGUSERCOMPLETION) Then
      fShowDetails = mfAdministrator
      
      If Not fShowDetails Then
        ' Was the form actioned for the current user?
        ' If it was then we can show the details.
        For iCount = 1 To UBound(masActionedUserNames)
          If (UCase(Trim(gsUserName)) = masActionedUserNames(iCount)) Then
            fShowDetails = True
            Exit For
          End If
        Next iCount
        
        If Not fShowDetails Then
          For iCount = 1 To UBound(masEmailAddresses)
            For iCount2 = 1 To UBound(masActionedUserNames)
              If (masEmailAddresses(iCount) = masActionedUserNames(iCount2)) Then
                fShowDetails = True
                Exit For
              End If
            Next iCount2
            
            For iCount2 = 1 To UBound(masActionedDelegatedEmailAddresses)
              If (UCase(masActionedDelegatedEmailAddresses(iCount2)) = masEmailAddresses(iCount)) Then
                fShowDetails = True
                Exit For
              End If
            Next iCount2
          
            If fShowDetails Then
              Exit For
            End If
          Next iCount
        End If
      End If

      If fShowDetails Then
        sSQL = "SELECT ASRSysWorkflowElementItems.inputType," & _
          " ASRSysWorkflowElementItems.itemType," & _
          " ASRSysWorkflowInstanceValues.identifier," & _
          " isnull(ASRSysWorkflowInstanceValues.value, '') AS [value]," & _
          " isnull(ASRSysWorkflowInstanceValues.fileUpload_fileName, '') AS [fileUpload_fileName]," & _
          " isnull(ASRSysWorkflowInstanceValues.valueDescription, '') AS [valueDescription]," & _
          " ASRSysWorkflowElementItems.lookupColumnID" & _
          " FROM ASRSysWorkflowInstanceValues" & _
          " INNER JOIN ASRSysWorkflowInstanceSteps ON ASRSysWorkflowInstanceValues.instanceID = ASRSysWorkflowInstanceSteps.instanceID" & _
          "   AND ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowInstanceSteps.elementID" & _
          " INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowInstanceValues.identifier = ASRSysWorkflowElementItems.identifier" & _
          "   AND ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElementItems.elementID" & _
          " WHERE ASRSysWorkflowInstanceSteps.ID = " & CStr(mlngWorkflowStepID) & _
          " ORDER BY ASRSysWorkflowElementItems.identifier"
      
        Set rstTemp = datGeneral.GetReadOnlyRecords(sSQL)
      
        With grdWebFormEnteredValues
          Do Until rstTemp.EOF
            sAddString = ""
            
            Select Case rstTemp.Fields("itemType")
              Case giWFFORMITEM_BUTTON
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Button" & vbTab
                If Len(rstTemp.Fields("value")) > 0 Then
                  sAddString = sAddString & _
                    IIf(rstTemp.Fields("value") = "1", "Pressed", "Not pressed")
                End If
              
              Case giWFFORMITEM_INPUTVALUE_GRID
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Record Selector" & vbTab & _
                  rstTemp.Fields("valueDescription")
                
              Case giWFFORMITEM_INPUTVALUE_NUMERIC
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Numeric" & vbTab & _
                  datGeneral.ConvertNumberForDisplay(rstTemp.Fields("value"))

              Case giWFFORMITEM_INPUTVALUE_LOGIC
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Logic" & vbTab
                If Len(rstTemp.Fields("value")) > 0 Then
                  sAddString = sAddString & _
                    IIf(rstTemp.Fields("value") = "1", "True", "False")
                End If

              Case giWFFORMITEM_INPUTVALUE_DATE
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Date" & vbTab
                If (Len(rstTemp.Fields("value")) > 0) _
                  And (UCase(Trim((rstTemp.Fields("value")))) <> "NULL") Then
                  ' Format date for the locale
                  sAddString = sAddString & _
                    ConvertSQLDateToLocale(rstTemp.Fields("value"))
                Else
                  sAddString = sAddString & _
                    "<undefined>"
                End If

              Case giWFFORMITEM_INPUTVALUE_CHAR
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Character" & vbTab & _
                  rstTemp.Fields("value")
            
              Case giWFFORMITEM_INPUTVALUE_DROPDOWN
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Dropdown" & vbTab & _
                  rstTemp.Fields("value")
            
              Case giWFFORMITEM_INPUTVALUE_LOOKUP
                    sAddString = _
                      rstTemp.Fields("identifier") & vbTab & _
                      "Lookup" & vbTab
                
                Select Case datGeneral.GetColumnDataType(rstTemp.Fields("lookupColumnID"))
                  Case -7 ' Logic
                    If Len(rstTemp.Fields("value")) > 0 Then
                      sAddString = sAddString & _
                        IIf(rstTemp.Fields("value") = "1", "True", "False")
                    End If
                  
                  Case 2, 4 ' Numeric, Integer
                    sAddString = sAddString & _
                      datGeneral.ConvertNumberForDisplay(rstTemp.Fields("value"))

                  Case 11 ' Date
                    If (Len(rstTemp.Fields("value")) > 0) _
                      And (UCase(Trim((rstTemp.Fields("value")))) <> "NULL") Then
                      ' Format date for the locale
                      sAddString = sAddString & _
                        ConvertSQLDateToLocale(rstTemp.Fields("value"))
                    Else
                      sAddString = sAddString & _
                        "<undefined>"
                    End If
                  
                  Case Else
                    sAddString = sAddString & _
                      rstTemp.Fields("value")
                End Select

              Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "Option Group" & vbTab & _
                  rstTemp.Fields("value")
            
              Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
                sTemp = rstTemp.Fields("fileUpload_fileName")
                iIndex = InStrRev(sTemp, "\")
                If iIndex > 0 Then
                  sTemp = Mid(sTemp, iIndex + 1)
                End If
                If Len(sTemp) = 0 Then
                  sTemp = "<none>"
                End If
                
                sAddString = _
                  rstTemp.Fields("identifier") & vbTab & _
                  "File Upload" & vbTab & _
                  sTemp
            End Select
            
            If Len(sAddString) > 0 Then
              .AddItem sAddString
            End If
    
            rstTemp.MoveNext
          Loop
          Set rstTemp = Nothing
        End With
      End If
    End If
    
    ResizeGridColumns grdWebFormEnteredValues
  
  ElseIf mlngElementType = elem_StoredData Then
    grdStoredDataValues.RemoveAll
      
    If (miStatus = giWFSTEPSTATUS_COMPLETED _
      Or miStatus = giWFSTEPSTATUS_FAILED _
      Or miStatus = giWFSTEPSTATUS_FAILEDACTION) Then
      
      If mfAdministrator Then
        sSQL = "SELECT DISTINCT ASRSysTables.tableName + '.' + ASRSysColumns.columnName AS [tableColumn]," & _
          " CASE" & _
          "   WHEN ASRSysWorkflowElementColumns.valueType = 0 THEN ASRSysWorkflowElementColumns.value" & _
          "   WHEN ASRSysWorkflowElementColumns.valueType = 2 OR ASRSysWorkflowElementColumns.valueType = 3 THEN isnull(DBV.value, '')" & _
          "   WHEN ASRSysWorkflowElementColumns.valueType = 1 AND (ASRSysColumns.dataType = -3 OR ASRSysColumns.dataType = -4) THEN isnull(ASRSysWorkflowInstanceValues.fileUpload_fileName, '')" & _
          "   ELSE isnull(ASRSysWorkflowInstanceValues.value, '')" & _
          " END AS [value]," & _
          " ASRSysWorkflowElementItems.InputType," & _
          " ASRSysColumns.dataType"
        sSQL = sSQL & _
          " FROM ASRSysWorkflowElementColumns" & _
          " INNER JOIN ASRSysWorkflowInstanceSteps ON ASRSysWorkflowElementColumns.elementID = ASRSysWorkflowInstanceSteps.elementID" & _
          " INNER JOIN ASRSysColumns ON ASRSysWorkflowElementColumns.columnID = ASRSysColumns.columnID" & _
          " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
          " INNER JOIN ASRSysWorkflowInstances ON ASRSysWorkflowInstanceSteps.instanceID = ASRSysWorkflowInstances.ID" & _
          " LEFT OUTER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementColumns.WFFormIdentifier = ASRSysWorkflowElements.identifier" & _
          "   AND ASRSysWorkflowInstances.workflowID = ASRSysWorkflowElements.workflowID" & _
          "   AND LEN(ASRSysWorkflowElementColumns.WFFormIdentifier) > 0" & _
          " LEFT OUTER JOIN ASRSysWorkflowInstanceValues ON ASRSysWorkflowInstanceSteps.instanceID = ASRSysWorkflowInstanceValues.instanceID" & _
          "   AND ASRSysWorkflowElementColumns.WFValueIdentifier = ASRSysWorkflowInstanceValues.identifier" & _
          "   AND ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID" & _
          " LEFT OUTER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowInstanceValues.identifier = ASRSysWorkflowElementItems.identifier" & _
          "   AND ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID" & _
          " LEFT OUTER JOIN ASRSysWorkflowInstanceValues DBV ON ASRSysWorkflowInstanceSteps.instanceID = DBV.instanceID" & _
          "   AND ASRSysWorkflowElementColumns.columnID = DBV.columnID" & _
          "   AND DBV.elementID = ASRSysWorkflowElementColumns.elementID" & _
          " WHERE ASRSysWorkflowInstanceSteps.ID = " & CStr(mlngWorkflowStepID) & " ORDER BY [tableColumn]"
        Set rstTemp = datGeneral.GetReadOnlyRecords(sSQL)
      
        With grdStoredDataValues
          Do Until rstTemp.EOF
            sAddString = rstTemp.Fields("tableColumn") & vbTab

            Select Case rstTemp.Fields("dataType")
              Case sqlNumeric, sqlInteger ' Numeric
                If (Len(rstTemp.Fields("value")) > 0) Then
                  sAddString = sAddString & _
                    datGeneral.ConvertNumberForDisplay(rstTemp.Fields("value"))
                End If

              Case sqlBoolean ' Logic
                If Len(rstTemp.Fields("value")) > 0 Then
                  sAddString = sAddString & _
                    IIf((rstTemp.Fields("value") = "1") Or (UCase(Trim(rstTemp.Fields("value"))) = "TRUE"), "True", "False")
                Else
                    sAddString = sAddString & ""
                End If

              Case sqlDate ' Date
                If (Len(rstTemp.Fields("value")) > 0) _
                  And (UCase(Trim((rstTemp.Fields("value")))) <> "NULL") Then
                  ' Format date for the locale
                  sAddString = sAddString & _
                    ConvertSQLDateToLocale(rstTemp.Fields("value"))
                Else
                  sAddString = sAddString & _
                    "<undefined>"
                End If
              
              Case sqlOle         ' OLE columns
                sTemp = rstTemp.Fields("value")
                iIndex = InStrRev(sTemp, "\")
                If iIndex > 0 Then
                  sTemp = Mid(sTemp, iIndex + 1)
                End If
                If Len(sTemp) = 0 Then
                  sTemp = "<none>"
                End If
                  
                sAddString = sAddString & _
                  sTemp
                    
              Case sqlVarBinary   ' Photo columns
                sTemp = rstTemp.Fields("value")
                iIndex = InStrRev(sTemp, "\")
                If iIndex > 0 Then
                  sTemp = Mid(sTemp, iIndex + 1)
                End If
                If Len(sTemp) = 0 Then
                  sTemp = "<none>"
                End If
              
                sAddString = sAddString & _
                  sTemp
                    
              Case Else ' Character
                sAddString = sAddString & _
                  rstTemp.Fields("value")
            End Select
    
            .AddItem sAddString
    
            rstTemp.MoveNext
          Loop
          Set rstTemp = Nothing
        End With
      Else
        ' Not Workflow Administrator
      End If
    End If
    
    ResizeGridColumns grdStoredDataValues
  End If

TidyUpAndExit:
  PopulateDetailsGrid = fOK
  Exit Function
  
ErrorTrap:
  MsgBox "Error retrieving details for this workflow step." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Workflow Log"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Sub ResizeGridColumns(pctlGrid As SSDBGrid)
  ' Size the visible columns in the given grid to fit the text.
  ' If the columns are then not as wide as the grid, stretch out the last visible column.

  Dim iLastVisibleColumn As Integer
  Dim iColumn As Integer
  Dim iRow As Integer
  Dim lngTextWidth As Long
  Dim varBookmark As Variant
  Dim varOriginalPos As Variant
  Dim fVerticalScrollRequired As Boolean
  Dim fHorizontalScrollRequired As Boolean
  
  Const SCROLLWIDTH = 255
  
  iLastVisibleColumn = -1
  lngTextWidth = 0
  
  With pctlGrid
    varOriginalPos = .Bookmark

    .Redraw = False
    .MoveFirst
    
    For iColumn = 0 To .Columns.Count - 1 Step 1
      lngTextWidth = Me.TextWidth(.Columns(iColumn).Caption)

      If .Columns(iColumn).Visible Then
        iLastVisibleColumn = iColumn
        
        For iRow = 0 To .Rows - 1 Step 1
          varBookmark = .AddItemBookmark(iRow)

          If BigTextWidth(.Columns(iColumn).CellText(varBookmark), 0) > lngTextWidth Then
            lngTextWidth = BigTextWidth(.Columns(iColumn).CellText(varBookmark), 0)
          End If
        Next iRow

        .Columns(iColumn).Width = lngTextWidth + 195
      End If
      lngTextWidth = 0
    Next iColumn

    If iLastVisibleColumn >= 0 Then
      ' Stretch out the last column if required
      fVerticalScrollRequired = (.Rows > .VisibleRows)
      
      If .Columns(iLastVisibleColumn).Left + .Columns(iLastVisibleColumn).Width _
        < (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) Then
      
        .Columns(iLastVisibleColumn).Width = _
          (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) - .Columns(iLastVisibleColumn).Left - 25
      End If
    End If
    
    .Bookmark = varOriginalPos
    .Redraw = True
  End With

End Sub



Private Sub cmdOK_Click()
  Unload Me

End Sub


Private Sub cmdOpenWebForm_Click()
  Dim strExePath As String
  Dim fIsDLL As Boolean
  
  strExePath = GetDefaultBrowserApplication(fIsDLL)

  If Len(Trim(strExePath)) > 1 Then
    OpenWebForm mlngWorkflowInstanceID, mlngWorkflowElementID
  Else
    MsgBox "Unable to open selected Workflow form." & vbCrLf & vbCrLf & "Please contact your system administrator.", vbExclamation + vbOKOnly, "Workflow"
  End If

End Sub

Private Sub cmdResendEmail_Click()
  Dim sTemp As String
  Dim sCaption As String
  Dim sSQL As String
  Dim sSQL2 As String
  Dim sEmailTo As String
  Dim sEmailCopyTo As String
  Dim sMessage As String
  Dim sMessageWhole_To As String
  Dim sMessageWhole_CopyTo As String
  Dim sMessageSuffix_Hypertextlink As String
  Dim sMessageSuffix_Resent As String
  Dim sSubject As String
  Dim datData As HRProDataMgr.clsDataAccess
  Dim iLoop As Integer
  Dim lngElementID As Long
  Dim sURL As String
  Dim sUser As String
  Dim sPassword As String
  Dim sEncryptedString As String
  Dim rstCaptions As Recordset
  
  On Error GoTo ErrorTrap

  'MH20061219 Fault 11839
  If GetSystemSetting("email", "method", 1) = 0 Then
    MsgBox "Unable to resend this message as server side emails are currently disabled." & vbCrLf & _
           "Please contact your system administrator.", vbCritical, Me.Caption
    Exit Sub
  End If

  Set datData = New clsDataAccess
  
  sEmailTo = txtEmailAddressValue.Text
  sEmailCopyTo = txtEmailCCValue.Text
  sMessage = Replace(txtEmailMessage.Text, vbNewLine, vbCr)
  sSubject = Replace(lblEmailSubjectValue.Caption, "&&", "&")

  If UBound(malngHypertextLinkedSteps) > 0 Then
    sURL = WorkflowURL
    ReadWebLogon sUser, sPassword
    
    sMessageSuffix_Hypertextlink = _
      "Click on the following link" _
      & IIf(UBound(malngHypertextLinkedSteps) = 1, "", "s") _
      & " to action:" & vbNewLine

    For iLoop = 1 To UBound(malngHypertextLinkedSteps)
      lngElementID = malngHypertextLinkedSteps(iLoop)
  
      sEncryptedString = EncryptQueryString(mlngWorkflowInstanceID, lngElementID, sUser, sPassword)

      sSQL2 = "SELECT isnull(WE.caption, '') AS [caption]" _
        & " FROM ASRSysWorkflowElements WE" _
        & " WHERE WE.ID = " & CStr(lngElementID)
      Set rstCaptions = datData.OpenRecordset(sSQL2, adOpenForwardOnly, adLockReadOnly)

      sCaption = IIf(IsNull(rstCaptions!Caption), "", rstCaptions!Caption)

      Set rstCaptions = Nothing
  
      sMessageSuffix_Hypertextlink = sMessageSuffix_Hypertextlink & vbNewLine _
        & sCaption & " - " & vbNewLine _
        & "<" & sURL & "?" & sEncryptedString & ">"
    Next iLoop
  
    sMessageSuffix_Hypertextlink = sMessageSuffix_Hypertextlink & vbNewLine & vbNewLine _
      & "Please make sure that the link has not been cut off by your display. If it has been cut off you will need to copy and paste it into your browser."
  End If
  
  'JPD 20061130 Fault 10998
  sMessageSuffix_Resent = _
    "Resent on " & Format(Now, DateFormat) & _
    " at " & Format(Now, "hh:nn") & _
    " by " & gsUserName
  
  sMessageWhole_To = _
      sMessage & vbNewLine & vbNewLine _
      & sMessageSuffix_Hypertextlink & vbNewLine & vbNewLine _
      & sMessageSuffix_Resent
  
  sSQL = "INSERT ASRSysEmailQueue(RecordDesc, ColumnValue, DateDue, UserName, [Immediate], RecalculateRecordDesc, RepTo, MsgText, Subject, WorkflowInstanceID)" & _
    " VALUES ('', '', getdate(), 'HR Pro Workflow', 1, 0, '" & Replace(sEmailTo, "'", "''") & "'" & _
    ", '" & Replace(sMessageWhole_To, "'", "''") & "', '" & Replace(sSubject, "'", "''") & "'," & CStr(mlngWorkflowInstanceID) & ")"
  datData.ExecuteSql (sSQL)

  If Len(sEmailCopyTo) > 0 Then
    sMessageWhole_CopyTo = _
      "You have been copied in on the following HR Pro Workflow email with recipients:" & vbNewLine _
      & vbTab & sEmailTo & vbNewLine & vbNewLine _
      & sMessage & vbNewLine & vbNewLine _
      & sMessageSuffix_Resent
  
    sSQL = "INSERT ASRSysEmailQueue(RecordDesc, ColumnValue, DateDue, UserName, [Immediate], RecalculateRecordDesc, RepTo, MsgText, Subject, WorkflowInstanceID)" & _
      " VALUES ('', '', getdate(), 'HR Pro Workflow', 1, 0, '" & Replace(sEmailCopyTo, "'", "''") & "'" & _
      ", '" & Replace(sMessageWhole_CopyTo, "'", "''") & "', '" & Replace(sSubject, "'", "''") & "'," & CStr(mlngWorkflowInstanceID) & ")"
    datData.ExecuteSql (sSQL)
  End If

  sTemp = gsUserName
  gsUserName = "HR Pro Workflow"
  objEmail.SendImmediateEmails
  gsUserName = sTemp

  MsgBox "Email resent.", vbInformation + vbOKOnly, "Workflow Log"

TidyUpAndExit:
  Set datData = Nothing
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If

End Sub


Private Sub Form_Load()
  Dim fraFrame As Frame
  Dim sngWidth As Single
  Dim sngHeight As Single
  Dim iLoop As Integer
  Dim sTemp As String
  
  RemoveIcon Me
  
  msngWebFormValueColumnWidth = grdWebFormEnteredValues.Columns("Value").Width
  msngStoredDataValueColumnWidth = grdStoredDataValues.Columns("Value").Width
  
  grdWebFormEnteredValues.RowHeight = 239
  grdStoredDataValues.RowHeight = 239
  
  ' Retrieve the size of the form when last viewed
  sngHeight = GetPCSetting("WorkflowLogStepDetails", "Height", Me.Height)
  sngWidth = GetPCSetting("WorkflowLogStepDetails", "Width", Me.Width)

'  Me.Height = sngHeight
'  Me.Width = sngWidth

  ReDim masngFormDimensions(1, elem_Connector2)
  For iLoop = elem_Begin To elem_Connector2
    Select Case iLoop
      Case elem_Begin
        sTemp = "Begin"
      Case elem_Terminator
        sTemp = "Terminator"
      Case elem_WebForm
        sTemp = "WebForm"
      Case elem_Email
        sTemp = "Email"
      Case elem_Decision
        sTemp = "Decision"
      Case elem_StoredData
        sTemp = "StoredData"
      Case elem_SummingJunction
        sTemp = "SummingJunction"
      Case elem_Or
        sTemp = "Or"
      Case elem_Connector1
        sTemp = "Connector1"
      Case elem_Connector2
        sTemp = "Connector2"
      Case Else
        sTemp = ""
    End Select
    
    masngFormDimensions(0, iLoop) = GetPCSetting("WorkflowLogStepDetails", "Height_" & sTemp, sngHeight)
    masngFormDimensions(1, iLoop) = GetPCSetting("WorkflowLogStepDetails", "Width_" & sTemp, sngWidth)
  Next iLoop
  
  fraBasicDetails.BackColor = vbButtonFace
  fraButtons.BackColor = vbButtonFace
  For Each fraFrame In fraSpecificDetails
    fraFrame.BackColor = vbButtonFace
    fraFrame.Top = fraBasicDetails.Top + fraBasicDetails.Height
  Next fraFrame
  Set fraFrame = Nothing
  
  txtWebFormMessage.Top = grdWebFormEnteredValues.Top
  
  cmdOpenWebForm.Left = cmdResendEmail.Left
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Store the size of the form for retrieval when next viewed
  Dim sTemp As String
  
  SavePCSetting "WorkflowLogStepDetails", "Height", Me.Height
  SavePCSetting "WorkflowLogStepDetails", "Width", Me.Width

  Select Case mlngElementType
    Case elem_Begin
      sTemp = "Begin"
    Case elem_Terminator
      sTemp = "Terminator"
    Case elem_WebForm
      sTemp = "WebForm"
    Case elem_Email
      sTemp = "Email"
    Case elem_Decision
      sTemp = "Decision"
    Case elem_StoredData
      sTemp = "StoredData"
    Case elem_SummingJunction
      sTemp = "SummingJunction"
    Case elem_Or
      sTemp = "Or"
    Case elem_Connector1
      sTemp = "Connector1"
    Case elem_Connector2
      sTemp = "Connector2"
    Case Else
      sTemp = ""
  End Select
  
  SavePCSetting "WorkflowLogStepDetails", "Height_" & sTemp, Me.Height
  SavePCSetting "WorkflowLogStepDetails", "Width_" & sTemp, Me.Width
  
End Sub


Private Sub Form_Resize()
  Dim lngColumn1Left As Long
  Dim lngColumn2Left As Long
  Dim lngColumn3Left As Long
  Dim lngColumn1DataLeft As Long
  Dim lngColumn2DataLeft As Long
  Dim lngColumn3DataLeft As Long
  'Dim sngMinFormHeight As Single
  Dim fraTemp As Frame
  Dim lngHeight As Long
  Dim lngLines As Long

  'JPD 20030908 Fault 5756
  DisplayApplication

  If mblnSizing Then Exit Sub

  mblnSizing = True

  Select Case mlngElementType
'    Case elem_WebForm
'      sngMinFormHeight = MINHEIGHT_WEBFORM
'
'    Case elem_Email
'      sngMinFormHeight = MINHEIGHT_EMAIL

    Case elem_Decision
      fraDetails.Height = (2 * fraBasicDetails.Top) _
        + fraBasicDetails.Height _
        + fraSpecificDetails(4).Height

'    Case elem_StoredData
'      sngMinFormHeight = MINHEIGHT_STOREDDATA

    Case Else
      ' elem_Begin, elem_Terminator, elem_SummingJunction
      ' elem_Or, elem_Connector1, elem_Connector2
      fraDetails.Height = (2 * fraBasicDetails.Top) _
        + fraBasicDetails.Height
  End Select

'  If Me.Height < sngMinFormHeight Then Me.Height = sngMinFormHeight
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 200)
'  If Me.Width < MINFORMWIDTH Then Me.Width = MINFORMWIDTH
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)

  If mfSizeable Then
    fraDetails.Height = Me.Height _
      - fraDetails.Top _
      - fraButtons.Height _
      - GAPOVERBUTTONS _
      - GAPUNDERBUTTONS

    If Not fraSpecificDetails(mlngElementType) Is Nothing Then
      fraSpecificDetails(mlngElementType).Height = fraDetails.Height _
        - fraSpecificDetails(mlngElementType).Top _
        - 100
    End If

    Select Case mlngElementType
      Case elem_WebForm
'        grdWebFormEnteredValues.Height = fraSpecificDetails(mlngElementType).Height _
'          - grdWebFormEnteredValues.Top _
'          - 100
        grdWebFormEnteredValues.Height = 1500
        txtWebFormMessage.Height = grdWebFormEnteredValues.Height

      Case elem_Email
        If Me.Height < fraDetails.Top _
            + (fraSpecificDetails(mlngElementType).Top + txtEmailMessage.Top + 1000) _
            + fraButtons.Height _
            + GAPOVERBUTTONS _
            + GAPUNDERBUTTONS Then

'          Me.Height = fraDetails.Top _
'            + (fraSpecificDetails(mlngElementType).Top + txtEmailMessage.Top + 1000) _
'            + fraButtons.Height _
'            + GAPOVERBUTTONS _
'            + GAPUNDERBUTTONS

          fraDetails.Height = Me.Height _
            - fraDetails.Top _
            - fraButtons.Height _
            - GAPOVERBUTTONS _
            - GAPUNDERBUTTONS

          If Not fraSpecificDetails(mlngElementType) Is Nothing Then
            fraSpecificDetails(mlngElementType).Height = fraDetails.Height _
              - fraSpecificDetails(mlngElementType).Top _
              - 100
          End If
        Else
          txtEmailMessage.Height = fraSpecificDetails(mlngElementType).Height _
            - txtEmailMessage.Top _
            - 100
        End If

      Case elem_StoredData
        txtStoredDataMessage.Height = (fraSpecificDetails(mlngElementType).Height _
          - txtStoredDataMessage.Top _
          - 240) / 3
        grdStoredDataValues.Top = txtStoredDataMessage.Top _
          + txtStoredDataMessage.Height _
          + 240
        grdStoredDataValues.Height = fraSpecificDetails(mlngElementType).Height _
          - grdStoredDataValues.Top _
          - 100
        lblStoredDataValues.Top = grdStoredDataValues.Top + 60
    End Select
  End If

  'AE20071106 Fault #12195
  'fraDetails.Width = Me.Width - 345
  fraDetails.Width = Me.ScaleWidth - (Me.fraDetails.Left * 2)

  fraBasicDetails.Width = fraDetails.Width _
    - 220
  For Each fraTemp In fraSpecificDetails
    fraTemp.Width = fraBasicDetails.Width
  Next fraTemp
  Set fraTemp = Nothing

  lngColumn1Left = COLUMN1LABELLEFT
  lngColumn1DataLeft = COLUMN1LABELLEFT + COLUMN1LABELWIDTH

  lngColumn2Left = (fraBasicDetails.Width / 2.775)
  lngColumn2DataLeft = lngColumn2Left + COLUMN2LABELWIDTH

  lngColumn3Left = (fraBasicDetails.Width / 1.44)
  lngColumn3DataLeft = lngColumn3Left + COLUMN3LABELWIDTH

  ' FIRST COLUMN

  ' Basic Details
  lblElementType.Left = lngColumn1Left
  lblElementTypeValue.Left = lngColumn1DataLeft
  lblElementTypeValue.Width = lngColumn2Left - lngColumn1DataLeft - GAPBETWEENCOLUMNS

  lblStartTime.Left = lngColumn1Left
  lblStartTimeValue.Left = lngColumn1DataLeft
  lblStartTimeValue.Width = lngColumn2Left - lngColumn1DataLeft - GAPBETWEENCOLUMNS

  lblCompletionCount.Left = lngColumn1Left
  lblCompletionCountValue.Left = lngColumn1DataLeft
  lblCompletionCountValue.Width = lngColumn2Left - lngColumn1DataLeft - GAPBETWEENCOLUMNS

  ' Decision Details
  linDecision.Left = lngColumn1Left
  linDecision.Width = fraBasicDetails.Width - lngColumn1Left - lngColumn1Left

  lblDecisionFlow.Left = lngColumn1Left
  lblDecisionFlowValue.Left = lngColumn1DataLeft
  lblDecisionFlowValue.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  ' Email Details
  linEmail.Left = lngColumn1Left
  linEmail.Width = fraBasicDetails.Width - lngColumn1Left - lngColumn1Left

  lblEmailAddress.Left = lngColumn1Left
  txtEmailAddressValue.Left = lngColumn1DataLeft
  txtEmailAddressValue.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left
  'Resize the textbox as required.
  lngLines = SendMessage(txtEmailAddressValue.hWnd, EM_GETLINECOUNT, 0, ByVal 0)
  picResizer.Font = txtEmailAddressValue.Font
  picResizer.FontSize = txtEmailAddressValue.FontSize
  lngHeight = picResizer.TextHeight("Xy") + 30
  txtEmailAddressValue.Height = lngHeight * lngLines

  lblEmailCc.Left = lngColumn1Left
  lblEmailCc.Top = txtEmailAddressValue.Top + txtEmailAddressValue.Height + 105
  txtEmailCCValue.Left = lngColumn1DataLeft
  txtEmailCCValue.Top = lblEmailCc.Top
  txtEmailCCValue.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left
  'Resize the textbox as required.
  lngLines = SendMessage(txtEmailCCValue.hWnd, EM_GETLINECOUNT, 0, ByVal 0)
  picResizer.Font = txtEmailCCValue.Font
  picResizer.FontSize = txtEmailCCValue.FontSize
  lngHeight = picResizer.TextHeight("Xy") + 30
  txtEmailCCValue.Height = lngHeight * lngLines

  lblEmailSubject.Left = lngColumn1Left
  lblEmailSubject.Top = IIf(lblEmailCc.Visible, txtEmailCCValue.Top + txtEmailCCValue.Height + 105, _
  txtEmailAddressValue.Top + txtEmailAddressValue.Height + 105)

  lblEmailSubjectValue.Top = lblEmailSubject.Top
  lblEmailSubjectValue.Left = lngColumn1DataLeft
  lblEmailSubjectValue.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  lblEmailMessage.Left = lngColumn1Left
  lblEmailMessage.Top = lblEmailSubject.Top + lblEmailSubject.Height + 165
  txtEmailMessage.Left = lngColumn1DataLeft
  txtEmailMessage.Top = lblEmailSubject.Top + lblEmailSubject.Height + 105
  txtEmailMessage.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  ' Web Form Details
  linWebForm.Left = lngColumn1Left
  linWebForm.Width = fraBasicDetails.Width - lngColumn1Left - lngColumn1Left

  lblWebFormUser.Left = lngColumn1Left
  lblWebFormUserValue.Left = lngColumn1DataLeft
  lblWebFormUserValue.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  grdWebFormEnteredValues.Visible = (miStatus <> giWFSTEPSTATUS_FAILED)
  txtWebFormMessage.Visible = (miStatus = giWFSTEPSTATUS_FAILED)

  lblWebFormEnteredValues.Left = lngColumn1Left
  grdWebFormEnteredValues.Left = lngColumn1DataLeft
  grdWebFormEnteredValues.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left
  txtWebFormMessage.Left = grdWebFormEnteredValues.Left
  txtWebFormMessage.Width = grdWebFormEnteredValues.Width

  ' Set the grid height and width
  With grdWebFormEnteredValues
    If .Columns("Value").Width < msngWebFormValueColumnWidth Then
      .Columns("Value").Width = msngWebFormValueColumnWidth
    Else
      .Columns("Value").Width = .Width - .Columns("Value").Left - 25 _
        - IIf(.VisibleRows < .Rows, SCROLLBAR_WIDTH, 0)
    End If
  End With

  ' Stored Data Details
  linStoredData.Left = lngColumn1Left
  linStoredData.Width = fraBasicDetails.Width - lngColumn1Left - lngColumn1Left

  lblStoredDataMessage.Left = lngColumn1Left
  txtStoredDataMessage.Left = lngColumn1DataLeft
  txtStoredDataMessage.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  lblStoredDataValues.Left = lngColumn1Left
  grdStoredDataValues.Left = lngColumn1DataLeft
  grdStoredDataValues.Width = fraBasicDetails.Width - lngColumn1DataLeft - lngColumn1Left

  ' Set the grid height and width
  With grdStoredDataValues
    If .Columns("Value").Width < msngStoredDataValueColumnWidth Then
      .Columns("Value").Width = msngStoredDataValueColumnWidth
    Else
      .Columns("Value").Width = .Width - .Columns("Value").Left - 25
    End If
  End With


  ' SECOND COLUMN

  ' Basic Details
  lblCaption.Left = lngColumn2Left
  lblCaptionValue.Left = lngColumn2DataLeft
  lblCaptionValue.Width = lngColumn3Left - lngColumn2DataLeft - GAPBETWEENCOLUMNS

  lblEndTime.Left = lngColumn2Left
  lblEndTimeValue.Left = lngColumn2DataLeft
  lblEndTimeValue.Width = lngColumn3Left - lngColumn2DataLeft - GAPBETWEENCOLUMNS

  lblFailureCount.Left = lngColumn2Left
  lblFailureCountValue.Left = lngColumn2DataLeft
  lblFailureCountValue.Width = lngColumn3Left - lngColumn2DataLeft - GAPBETWEENCOLUMNS

  ' THIRD COLUMN

  ' Basic Details
  lblStatus.Left = lngColumn3Left
  lblStatusValue.Left = lngColumn3DataLeft
  lblStatusValue.Width = fraBasicDetails.Width - lngColumn3Left - lngColumn1DataLeft

  lblDuration.Left = lngColumn3Left
  lblDurationValue.Left = lngColumn3DataLeft
  lblDurationValue.Width = fraBasicDetails.Width - lngColumn3Left - lngColumn1DataLeft

  lblTimeoutCount.Left = lngColumn3Left
  lblTimeoutCountValue.Left = lngColumn3DataLeft
  lblTimeoutCountValue.Width = fraBasicDetails.Width - lngColumn3Left - lngColumn1DataLeft

  fraButtons.Top = fraDetails.Top _
    + fraDetails.Height _
    + GAPOVERBUTTONS
  fraButtons.Left = Me.ScaleWidth _
    - Me.fraButtons.Width _
    - Me.fraDetails.Left
  'AE20071203 Fault #12195
'  fraButtons.Left = Me.Width _
'    - fraButtons.Width _
'    - 275

  If Not mfSizeable Then
'    Me.Height = fraButtons.Top _
'      + fraButtons.Height _
'      + GAPUNDERBUTTONS
  End If

  mblnSizing = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

