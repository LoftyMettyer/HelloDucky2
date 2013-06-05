VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmEmailLink 
   Caption         =   "Email Link"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   705
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1016
   Icon            =   "frmEmailLink.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDocument 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   1320
      Picture         =   "frmEmailLink.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picDocument 
      Height          =   465
      Index           =   2
      Left            =   2520
      Picture         =   "frmEmailLink.frx":08D6
      ScaleHeight     =   405
      ScaleWidth      =   465
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picDocument 
      Height          =   480
      Index           =   1
      Left            =   1920
      Picture         =   "frmEmailLink.frx":11A0
      ScaleHeight     =   420
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picNoDrop 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      Picture         =   "frmEmailLink.frx":1A6A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   6975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select File Attachment"
      Filter          =   "All Files|*.*"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7950
      TabIndex        =   27
      Top             =   6990
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   6645
      TabIndex        =   26
      Top             =   6990
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   45
      TabIndex        =   28
      Top             =   45
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "De&finition"
      TabPicture(0)   =   "frmEmailLink.frx":1D74
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraLinkTypeDetails(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraLinkTypeDetails(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraLinkTypeDetails(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmDefinition(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmDefinition(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Co&ntent"
      TabPicture(1)   =   "frmEmailLink.frx":1D90
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmContent"
      Tab(1).ControlCount=   1
      Begin VB.Frame frmContent 
         Height          =   6240
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   8800
         Begin VB.TextBox txtAttachment 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1900
            Width           =   6645
         End
         Begin VB.CommandButton cmdAttachmentClear 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   20.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8295
            MaskColor       =   &H000000FF&
            TabIndex        =   44
            Top             =   1900
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtContent 
            Height          =   3720
            Index           =   1
            Left            =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2350
            Width           =   5595
         End
         Begin VB.TextBox txtContent 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   42
            Top             =   1500
            Width           =   6930
         End
         Begin VB.TextBox txtRecipients 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   300
            Width           =   6930
         End
         Begin VB.TextBox txtRecipients 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   700
            Width           =   6930
         End
         Begin VB.TextBox txtRecipients 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   2
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1100
            Width           =   6930
         End
         Begin VB.CommandButton cmdRecipients 
            Caption         =   "To..."
            Height          =   315
            Index           =   0
            Left            =   200
            TabIndex        =   38
            Top             =   300
            Width           =   1305
         End
         Begin VB.CommandButton cmdRecipients 
            Caption         =   "Cc..."
            Height          =   315
            Index           =   1
            Left            =   200
            TabIndex        =   37
            Top             =   700
            Width           =   1305
         End
         Begin VB.CommandButton cmdRecipients 
            Caption         =   "Bcc..."
            Height          =   315
            Index           =   2
            Left            =   200
            TabIndex        =   36
            Top             =   1100
            Width           =   1305
         End
         Begin VB.CommandButton cmdAttachment 
            Caption         =   "Attachment..."
            Height          =   315
            Left            =   200
            TabIndex        =   35
            Top             =   1900
            Width           =   1305
         End
         Begin ComctlLib.TreeView sstrvAvailable 
            Height          =   3720
            Left            =   180
            TabIndex        =   34
            Top             =   2350
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   6562
            _Version        =   327682
            HideSelection   =   0   'False
            Indentation     =   556
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDragMode     =   1
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Subject :"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   46
            Top             =   1560
            Width           =   645
         End
      End
      Begin VB.Frame frmDefinition 
         Caption         =   "Link Type :"
         Height          =   4105
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2495
         Width           =   1480
         Begin VB.OptionButton optLinkType 
            Caption         =   "Co&lumn"
            Height          =   255
            Index           =   0
            Left            =   200
            TabIndex        =   9
            Tag             =   "0"
            Top             =   360
            Width           =   1125
         End
         Begin VB.OptionButton optLinkType 
            Caption         =   "&Record"
            Height          =   255
            Index           =   1
            Left            =   200
            TabIndex        =   10
            Tag             =   "1"
            Top             =   760
            Width           =   1125
         End
         Begin VB.OptionButton optLinkType 
            Caption         =   "D&ate"
            Height          =   255
            Index           =   2
            Left            =   200
            TabIndex        =   11
            Tag             =   "2"
            Top             =   1160
            Width           =   930
         End
      End
      Begin VB.Frame frmDefinition 
         Height          =   2075
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8800
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   700
            Width           =   6540
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8340
            TabIndex        =   5
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtTitle 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   2
            Top             =   300
            Width           =   6855
         End
         Begin GTMaskDate.GTMaskDate cboEffectiveDate 
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   1095
            Width           =   1440
            _Version        =   65537
            _ExtentX        =   2540
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            NullText        =   "__/__/____"
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoSelect      =   -1  'True
            Text            =   "  /  /"
            MaskCentury     =   2
            SpinButtonEnabled=   0   'False
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTips        =   0   'False
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date :"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   6
            Top             =   1155
            Width           =   1410
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Filter :"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   760
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   1
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame fraLinkTypeDetails 
         Caption         =   "Column Related Link :"
         Height          =   4105
         Index           =   0
         Left            =   1720
         TabIndex        =   12
         Top             =   2495
         Width           =   7200
         Begin VB.ListBox lstColumnLinkColumns 
            Height          =   3660
            Left            =   200
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   300
            Width           =   6825
         End
      End
      Begin VB.Frame fraLinkTypeDetails 
         Caption         =   "Date Related Link :"
         Height          =   4105
         Index           =   2
         Left            =   1720
         TabIndex        =   18
         Top             =   2495
         Visible         =   0   'False
         Width           =   7200
         Begin VB.CheckBox chkDateAmendments 
            Caption         =   "E&mail recipients any data changes"
            Height          =   240
            Left            =   200
            TabIndex        =   25
            Top             =   1200
            Value           =   1  'Checked
            Width           =   4500
         End
         Begin VB.ComboBox cboDateLinkDirection 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEmailLink.frx":1DAC
            Left            =   3480
            List            =   "frmEmailLink.frx":1DBC
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   705
            Width           =   1300
         End
         Begin VB.ComboBox cboDateLinkColumn 
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   300
            Width           =   5895
         End
         Begin VB.ComboBox cboDateLinkOffsetPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEmailLink.frx":1DDC
            Left            =   2040
            List            =   "frmEmailLink.frx":1DEC
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   705
            Width           =   1300
         End
         Begin COASpinner.COA_Spinner spnDateLinkOffset 
            Height          =   315
            Left            =   1125
            TabIndex        =   22
            Top             =   705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin VB.Label lblDateLinkColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Left            =   200
            TabIndex        =   19
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lblDateLinkOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset :"
            Height          =   195
            Left            =   200
            TabIndex        =   21
            Top             =   760
            Width           =   570
         End
      End
      Begin VB.Frame fraLinkTypeDetails 
         Caption         =   "Record Related Link :"
         Height          =   4105
         Index           =   1
         Left            =   1720
         TabIndex        =   14
         Top             =   2495
         Visible         =   0   'False
         Width           =   7200
         Begin VB.CheckBox chkRecordLinkRecord 
            Caption         =   "&Update Record"
            Height          =   240
            Index           =   1
            Left            =   200
            TabIndex        =   16
            Top             =   760
            Width           =   1980
         End
         Begin VB.CheckBox chkRecordLinkRecord 
            Caption         =   "&Insert Record"
            Height          =   240
            Index           =   0
            Left            =   200
            TabIndex        =   15
            Top             =   360
            Width           =   1980
         End
         Begin VB.CheckBox chkRecordLinkRecord 
            Caption         =   "D&elete Record"
            Height          =   240
            Index           =   2
            Left            =   200
            TabIndex        =   17
            Top             =   1160
            Width           =   1980
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmailLink.frx":1E0C
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmailLink.frx":235E
            Key             =   "IMG_CALC"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmailLink.frx":28B0
            Key             =   "IMG_CUSTOM"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmailLink.frx":2E02
            Key             =   "IMG_NO"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEmailLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mlngTableID As Long
Private mstrTableName As String
Private mblnCancelled As Boolean
Private mblnLoading As Boolean
Private msDocumentsPath As String

Private mobjTempLink As clsEmailLink
Private mblnReadOnly As Boolean

' Flag to see if any changes have been made by the user
'Private mblnChanged As Boolean

Private mcolAvailableComponents As Collection
Private mcolRecipients() As Collection


Public Property Let Changed(ByVal value As Boolean)
  If Not mblnLoading Then
    cmdOk.Enabled = value
  End If
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property


' Return the character position under the mouse.
Public Function TextBoxCursorPos(ByVal txt As TextBox, ByVal x As Single, ByVal y As Single) As Long
    ' Convert the position to pixels.
    x = x \ Screen.TwipsPerPixelX
    y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    TextBoxCursorPos = SendMessageLong(txt.hWnd, EM_CHARFROMPOS, 0&, CLng(x + y * &H10000)) And &HFFFF&
End Function


Public Property Get Cancelled() As Variant
  Cancelled = mblnCancelled
End Property

'Public Property Let AllowOffset(ByVal blnNewValue As Boolean)
'  optImmediate.Enabled = (blnNewValue And Not mblnReadOnly)
'  optOffset.Enabled = (blnNewValue And Not mblnReadOnly)
'End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
  mstrTableName = GetTableName(mlngTableID)
End Property


Public Property Get EmailLink() As clsEmailLink

  Set EmailLink = mobjTempLink

End Property

Public Property Let EmailLink(ByVal objNewValue As clsEmailLink)

  If mobjTempLink Is Nothing Then
    Set mobjTempLink = New clsEmailLink
  End If
  Set mobjTempLink = objNewValue

End Property


Public Sub PopulateControls()

  mblnCancelled = False
  SSTab1.Tab = 0

  Call PopulateCombos
  Call PopulateColumns(mlngTableID)

  With mobjTempLink

    txtTitle.Text = .Title
    txtFilter.Tag = .FilterID
    txtFilter.Text = GetExpressionName(.FilterID)
    cboEffectiveDate.DateValue = .EffectiveDate
    
    optLinkType(0).value = (.LinkType = 0)
    optLinkType(1).value = (.LinkType = 1)
    optLinkType(2).value = (.LinkType = 2)
    
    chkRecordLinkRecord(0).value = IIf(.RecordInsert, vbChecked, vbUnchecked)
    chkRecordLinkRecord(1).value = IIf(.RecordUpdate, vbChecked, vbUnchecked)
    chkRecordLinkRecord(2).value = IIf(.RecordDelete, vbChecked, vbUnchecked)
    
    SetComboItem cboDateLinkColumn, .DateColumnID
    If .DateColumnID > 0 Then
      cboDateLinkDirection.ListIndex = IIf(.DateOffset < 0, 0, 1)
      spnDateLinkOffset.value = Abs(.DateOffset)
      SetComboItem cboDateLinkOffsetPeriod, .DatePeriod
      chkDateAmendments.value = IIf(.DateAmendment, vbChecked, vbUnchecked)
    End If

    'optOffset.value = Not (.Immediate)
    'Call OffsetEnabled(Not .Immediate)
    'spnOffset.value = .Offset
    'txtSubject = .Subject
    'SetComboItem cboImportance, .Importance
    'SetComboItem cboSensitivity, .Sensitivity
    'chkIncludeRecDesc = IIf(.IncRecordDesc, vbChecked, vbUnchecked)
    'chkIncludeColumn = IIf(.IncColumnDetails, vbChecked, vbUnchecked)
    'chkIncludeUsername = IIf(.IncUsername, vbChecked, vbUnchecked)

    'chkWhenInsert.value = vbUnchecked 'IIf(.EmailInsert = True, vbChecked, vbUnchecked)
    'chkWhenUpdate.value = vbChecked 'IIf(.EmailUpdate = True, vbChecked, vbUnchecked)
    'chkWhenDelete.value = vbUnchecked 'IIf(.EmailDelete = True, vbChecked, vbUnchecked)

    'txtBody.Text = .Text
    
    
    'MH20090520
    .SubjectContent.SetTextboxFromContent txtContent(0)
    .BodyContent.SetTextboxFromContent txtContent(1)
    
    txtAttachment.Text = .Attachment
    cmdAttachmentClear.Enabled = (txtAttachment.Text <> "" And Not mblnReadOnly)
    
    Set mcolRecipients(0) = .RecipientsTo
    Set mcolRecipients(1) = .RecipientsCc
    Set mcolRecipients(2) = .RecipientsBcc
    
    
  End With

  'Call PopulateRecipients(True)
  'Call PopulateAttachments
  'Call UpdateAttachmentButtonStatus

  RefreshRecipients
  PopulateComponents
  
  mblnLoading = False

End Sub


Private Sub cboDateLinkColumn_Click()
  Changed = True
End Sub

Private Sub cboDateLinkDirection_Click()
  Changed = True
End Sub

Private Sub cboDateLinkOffsetPeriod_Click()
  Changed = True
End Sub

Private Sub cboEffectiveDate_Change()
  Changed = True
End Sub

Private Sub cboEffectiveDate_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    cboEffectiveDate.DateValue = Date
  End If
End Sub

Private Sub cboEffectiveDate_LostFocus()

  If IsEmpty(cboEffectiveDate.DateValue) Then
     MsgBox "You must enter a date.", vbOKOnly + vbExclamation, App.Title
  End If
  
'  If IsNull(cboEffectiveDate.DateValue) Then
'     cboEffectiveDate.ForeColor = vbRed
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboEffectiveDate.ForeColor = vbWindowText
'     cboEffectiveDate.DateValue = Null
'     cboEffectiveDate.SetFocus
'     Exit Sub
'  End If

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboEffectiveDate

End Sub


Private Sub PopulateCombos()

  With cboDateLinkOffsetPeriod
    .Clear
    .AddItem "Day(s)": .ItemData(.NewIndex) = iTimePeriodDays
    .AddItem "Week(s)": .ItemData(.NewIndex) = iTimePeriodWeeks
    .AddItem "Month(s)": .ItemData(.NewIndex) = iTimePeriodMonths
    .AddItem "Year(s)": .ItemData(.NewIndex) = iTimePeriodYears
    .ListIndex = -1
  End With

  With cboDateLinkDirection
    .Clear
    .AddItem "Before"
    .AddItem "After"
    .ListIndex = -1
  End With

End Sub

Private Sub chkDateAmendments_Click()
  Changed = True
End Sub

Private Sub chkRecordLinkRecord_Click(index As Integer)
  Changed = True
End Sub

Private Sub cmdRecipients_Click(index As Integer)
  
  Dim objEmail As clsEmailAddr

  Set objEmail = New clsEmailAddr
  
  With objEmail
    .EmailIDs = mcolRecipients(index)
    .TableID = mlngTableID
  
    If .SelectEmail(mblnReadOnly, True) Then
      Set mcolRecipients(index) = .EmailIDs
    End If

  End With

  RefreshRecipients

End Sub

Private Sub RefreshRecipients()

  Dim intModeIndex As Integer
  Dim lngRecipientIndex As Long
  Dim lngRecipientID As Long
  Dim strRecipientName As String
  Dim strOutput As String
  
  
  'Update all address boxes with any amended names and remove any that have been deleted.
  For intModeIndex = 0 To 2
    With txtRecipients(intModeIndex)
      strOutput = vbNullString

      For lngRecipientIndex = mcolRecipients(intModeIndex).Count To 1 Step -1
        lngRecipientID = mcolRecipients(intModeIndex).Item(lngRecipientIndex)
        strRecipientName = GetEmailAddressName(lngRecipientID)

        If strRecipientName = vbNullString Then
          mcolRecipients(intModeIndex).Remove lngRecipientIndex
        Else
          strOutput = strRecipientName & _
            IIf(strOutput <> vbNullString, "; " & strOutput, "")
        End If
      Next

      .Text = strOutput
    End With
  Next

End Sub

Private Sub cmdAttachment_Click()

  Dim frmFileSel As frmEmailLinkAttachmentSel

  Set frmFileSel = New frmEmailLinkAttachmentSel

  If Trim(gstrEmailAttachmentPath) = vbNullString Then
    MsgBox "You will need to set up an email path in configuration prior to adding email attachments", vbExclamation, "Email Link"
    Exit Sub
  End If

  With frmFileSel
    .Show vbModal
    If .Cancelled = False Then
      txtAttachment.Text = .FileName
      cmdAttachmentClear.Enabled = (txtAttachment.Text <> vbNullString And Not mblnReadOnly)
    End If
  End With

  Set frmFileSel = Nothing

End Sub


Private Sub cmdAttachmentClear_Click()

  If MsgBox("Are you sure you wish to remove this attachment ?", vbYesNo + vbQuestion, "Remove Attachment") = vbYes Then
    txtAttachment.Text = ""
  End If
  cmdAttachmentClear.Enabled = (txtAttachment.Text <> "")
  
End Sub


Private Sub cmdCancel_Click()
  mblnCancelled = True
  UnLoad Me

  'Me.Hide
End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Where Clause' expression selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise mlngTableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
    End If
  
  End With


TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdOK_Click()
  
  Dim strErrorText As String
  Dim intIndex As Integer
  
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If IsEmpty(cboEffectiveDate.DateValue) Then
    MsgBox "You must enter a date.", vbOKOnly + vbExclamation, App.Title
    Exit Sub
  End If
  
  If ValidateGTMaskDate(cboEffectiveDate) = False Then
    Exit Sub
  End If
  
  With mobjTempLink

    mblnCancelled = False
    
    
    If Len(Trim(txtTitle.Text)) = 0 Then
      SSTab1.Tab = 0
      MsgBox "Please enter a title for this email link.", vbExclamation
      Exit Sub
    End If
    
    If IsNull(cboEffectiveDate.DateValue) Then
      SSTab1.Tab = 0
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly, App.Title
      Exit Sub
    End If


    If optLinkType(0).value Then
      If lstColumnLinkColumns.SelCount < 1 Then
        SSTab1.Tab = 0
        MsgBox "Please select at least one related column.", vbExclamation
        Exit Sub
      End If
    
    ElseIf optLinkType(1).value Then
      If chkRecordLinkRecord(0).value = vbUnchecked And _
         chkRecordLinkRecord(1).value = vbUnchecked And _
         chkRecordLinkRecord(2).value = vbUnchecked Then
        SSTab1.Tab = 0
        MsgBox "Please select Insert, Update and/or Delete.", vbExclamation
        Exit Sub
      End If
    End If
    
    
    If mcolRecipients(0).Count < 1 Then
      SSTab1.Tab = 1
      MsgBox "Please select at least one recipient to send the email to.", vbExclamation
      Exit Sub
    End If
    
    
    If txtContent(0).Text & txtContent(1).Text & txtAttachment.Text = vbNullString Then
      SSTab1.Tab = 1
      MsgBox "Please include either a subject, content or an attachment to the email.", vbExclamation
      txtContent(0).SetFocus
      Exit Sub
    End If
    
    
    strErrorText = .SubjectContent.SetContentFromTextbox(txtContent(0), mcolAvailableComponents, mlngTableID)
    If strErrorText <> vbNullString Then
      SSTab1.Tab = 1
      MsgBox strErrorText, vbExclamation
      txtContent(0).SetFocus
      Exit Sub
    End If

    strErrorText = .BodyContent.SetContentFromTextbox(txtContent(1), mcolAvailableComponents, mlngTableID)
    If strErrorText <> vbNullString Then
      SSTab1.Tab = 1
      MsgBox strErrorText, vbExclamation
      txtContent(1).SetFocus
      Exit Sub
    End If
    

    
    .Title = txtTitle.Text
    .FilterID = txtFilter.Tag
    .EffectiveDate = cboEffectiveDate.DateValue
    .TableID = mlngTableID
    
    If optLinkType(2).value Then
      .LinkType = 2
    ElseIf optLinkType(1).value Then
      .LinkType = 1
    Else
      .LinkType = 0
    End If

    .RecordInsert = (chkRecordLinkRecord(0).value = vbChecked)
    .RecordUpdate = (chkRecordLinkRecord(1).value = vbChecked)
    .RecordDelete = (chkRecordLinkRecord(2).value = vbChecked)

    If cboDateLinkColumn.ListIndex >= 0 And cboDateLinkOffsetPeriod.ListIndex >= 0 Then
      .DateColumnID = cboDateLinkColumn.ItemData(cboDateLinkColumn.ListIndex)
      .DateOffset = spnDateLinkOffset.value * IIf(cboDateLinkDirection.ListIndex = 1, 1, -1)
      .DatePeriod = cboDateLinkOffsetPeriod.ItemData(cboDateLinkOffsetPeriod.ListIndex)
      .DateAmendment = (chkDateAmendments.value = vbChecked)
    Else
      .DateColumnID = 0
      .DateOffset = 0
      .DatePeriod = 0
      .DateAmendment = True
    End If

    .Attachment = txtAttachment.Text

    .RecipientsTo = mcolRecipients(0)
    .RecipientsCc = mcolRecipients(1)
    .RecipientsBcc = mcolRecipients(2)
    

    .Columns = New Collection
    For intIndex = 0 To lstColumnLinkColumns.ListCount - 1
      If lstColumnLinkColumns.Selected(intIndex) Then
        .Columns.Add lstColumnLinkColumns.ItemData(intIndex), CStr(lstColumnLinkColumns.ItemData(intIndex))
      End If
    Next
  
  End With

  'Prompt to rebuild if this is a date related link
  If optLinkType(2).value = False Then
    Application.ChangedEmailLink = True
  End If
  Me.Hide

End Sub

Private Sub Form_Activate()
  Form_Resize
End Sub

Private Sub Form_Load()

  mblnLoading = True

  ReDim mcolRecipients(2)
  Set mcolRecipients(0) = New Collection
  Set mcolRecipients(1) = New Collection
  Set mcolRecipients(2) = New Collection


  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  'JPD 20041115 Fault 8970
  UI.FormatGTDateControl cboEffectiveDate
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    cmdFilter.Enabled = True
    'cmdAddresses.Enabled = True
  End If
  
  msDocumentsPath = GetPCSetting("DataPaths", "documentspath_" & gsDatabaseName, App.Path)

  RemoveIcon Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer

  If UnloadMode = vbFormControlMenu Then
    mblnCancelled = True
  End If

  If mblnCancelled = True Then
    
    If Changed = True Then  'And cmdOK.Enabled Then
      
      'pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
      pintAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, App.Title)
      
      If pintAnswer = vbYes Then
        cmdOK_Click
        Cancel = True   'MH20021105 Fault 4694
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Cancel = True
        Exit Sub
      End If
    
    End If
  
  End If

End Sub

Private Sub Form_Resize()
  
  Dim GAP As Long
  Dim lngLeft As Long
  Dim lngTop As Long
  Dim lngWidth As Long
  Dim lngHeight As Long
  
  On Local Error Resume Next
  
  
  'JPD 20030908 Fault 5756
  DisplayApplication

  
  GAP = 100


  'BUTTONS & TAB CONTROL
  lngLeft = Me.ScaleWidth - (cmdCancel.Width + GAP)
  lngTop = Me.ScaleHeight - (cmdCancel.Height + GAP)
  cmdCancel.Move lngLeft, lngTop
  
  lngLeft = lngLeft - (cmdOk.Width + GAP)
  cmdOk.Move lngLeft, lngTop
  
  lngWidth = Me.ScaleWidth - (GAP * 2)
  lngHeight = lngTop - (GAP * 2)
  SSTab1.Move GAP, GAP, lngWidth, lngHeight
  
  
  GAP = 200


  'FRAMES
  lngLeft = GAP
  lngTop = 240 + GAP
  lngWidth = lngWidth - (frmContent.Left + GAP)
  lngHeight = lngHeight - (frmContent.Top + GAP)
  
  frmDefinition(0).Move lngLeft, lngTop, lngWidth
  frmContent.Move lngLeft, lngTop, lngWidth, lngHeight

  
  lngLeft = GAP
  lngTop = frmDefinition(0).Top + frmDefinition(0).Height + 100
  lngWidth = frmDefinition(1).Width
  lngHeight = SSTab1.Height - (lngTop + GAP)
  
  frmDefinition(1).Move lngLeft, lngTop, lngWidth, lngHeight
  
  lngLeft = frmDefinition(1).Left + frmDefinition(1).Width + GAP
  lngWidth = SSTab1.Width - (lngLeft + GAP)
  
  fraLinkTypeDetails(0).Move lngLeft, lngTop, lngWidth, lngHeight
  fraLinkTypeDetails(1).Move lngLeft, lngTop, lngWidth, lngHeight
  fraLinkTypeDetails(2).Move lngLeft, lngTop, lngWidth, lngHeight
  
  
  'DEFINITION TAB CONTROLS
  lngWidth = frmDefinition(0).Width - GAP
  
  txtTitle.Width = lngWidth - txtTitle.Left
  cmdFilter.Left = lngWidth - cmdFilter.Width
  txtFilter.Width = cmdFilter.Left - txtFilter.Left
  
  lngWidth = fraLinkTypeDetails(0).Width - GAP
  lngHeight = fraLinkTypeDetails(0).Height - GAP
  
  lstColumnLinkColumns.Width = lngWidth - (lstColumnLinkColumns.Left)
  lstColumnLinkColumns.Height = lngHeight - (lstColumnLinkColumns.Top)
  cboDateLinkColumn.Width = lngWidth - (cboDateLinkColumn.Left)


  'CONTENT TAB CONTROLS
  lngWidth = frmContent.Width - GAP
  lngHeight = frmContent.Height - GAP
  
  txtRecipients(0).Width = lngWidth - txtRecipients(0).Left
  txtRecipients(1).Width = lngWidth - txtRecipients(0).Left
  txtRecipients(2).Width = lngWidth - txtRecipients(0).Left
  txtContent(0).Width = lngWidth - txtContent(0).Left
  cmdAttachmentClear.Left = lngWidth - cmdAttachmentClear.Width
  txtAttachment.Width = cmdAttachmentClear.Left - txtAttachment.Left
  txtContent(1).Height = lngHeight - txtContent(1).Top
  txtContent(1).Width = lngWidth - txtContent(1).Left
  sstrvAvailable.Height = lngHeight - sstrvAvailable.Top

  On Local Error GoTo 0

End Sub


Private Sub Form_Terminate()
  Set mcolAvailableComponents = Nothing
  Set mcolRecipients(0) = Nothing
  Set mcolRecipients(1) = Nothing
  Set mcolRecipients(2) = Nothing
  'Set mcolRecipients = Nothing
End Sub

Private Sub lstColumnLinkColumns_ItemCheck(Item As Integer)
  Changed = True
End Sub

Private Sub optLinkType_Click(index As Integer)
  fraLinkTypeDetails(0).Visible = (index = 0)
  fraLinkTypeDetails(1).Visible = (index = 1)
  fraLinkTypeDetails(2).Visible = (index = 2)
  Changed = True
End Sub

Private Sub spnDateLinkOffset_Change()
  
  Dim blnOffset As Boolean
  
  blnOffset = (spnDateLinkOffset.value > 0)
  
  cboDateLinkOffsetPeriod.Enabled = blnOffset
  cboDateLinkOffsetPeriod.BackColor = IIf(blnOffset, vbWindowBackground, vbButtonFace)
  
  cboDateLinkDirection.Enabled = blnOffset
  cboDateLinkDirection.BackColor = IIf(blnOffset, vbWindowBackground, vbButtonFace)
  
  If blnOffset Then
    If cboDateLinkOffsetPeriod.ListIndex < 0 Then
      cboDateLinkOffsetPeriod.ListIndex = 0
    End If
    If cboDateLinkDirection.ListIndex < 0 Then
      cboDateLinkDirection.ListIndex = 0
    End If
  Else
    cboDateLinkOffsetPeriod.ListIndex = -1
    cboDateLinkDirection.ListIndex = -1
  End If

  Changed = True

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

'  If Not mblnReadOnly Then
'    frmDefinition(0).Enabled = (SSTab1.Tab = 0)
'    frmDefinition(1).Enabled = (SSTab1.Tab = 0)
'    fraLinkTypeDetails(0).Enabled = (SSTab1.Tab = 0)
'    fraLinkTypeDetails(1).Enabled = (SSTab1.Tab = 0)
'    fraLinkTypeDetails(2).Enabled = (SSTab1.Tab = 0)
'    frmContent.Enabled = (SSTab1.Tab = 1)
'  End If

  frmDefinition(0).Visible = (SSTab1.Tab = 0)
  frmDefinition(1).Visible = (SSTab1.Tab = 0)
  fraLinkTypeDetails(0).Visible = (SSTab1.Tab = 0)
  fraLinkTypeDetails(1).Visible = (SSTab1.Tab = 0)
  fraLinkTypeDetails(2).Visible = (SSTab1.Tab = 0)
  frmContent.Visible = (SSTab1.Tab = 1)

End Sub

Private Sub sstrvAvailable_DblClick()
  InsertColumn 1, txtContent(1).SelStart
End Sub

Private Sub txtAttachment_Change()
  Changed = True
End Sub

Private Sub txtContent_Change(index As Integer)
  Changed = True
End Sub

Private Sub txtContent_GotFocus(index As Integer)
  cmdOk.Default = False
End Sub

Private Sub txtContent_LostFocus(index As Integer)
  cmdOk.Default = True
End Sub

Private Sub txtFilter_Change()
  Changed = True
End Sub

Private Sub txtRecipients_Change(index As Integer)
  Changed = True
End Sub

Private Sub txtTitle_Change()
  Changed = True
End Sub

Private Sub txtTitle_GotFocus()
  With txtTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub


'Private Sub AddNode(objParent As SSNode, strCode As String, intID As Integer, strPrefix As String, strText As String, strImage As String)
Private Sub AddNode(objParent As Node, strCode As String, intID As Integer, strPrefix As String, strText As String, strImage As String)

  'Dim objNode As SSNode
  Dim objNode As Node

  Set objNode = sstrvAvailable.Nodes.Add(objParent, tvwChild, strCode & CStr(intID), strText, strImage, strImage)
  objNode.Tag = strPrefix & strText
  
  'MH20090804 Fault HRPRO-193 (also refer clsLinkContent)
  If strCode <> "E" Then
    'These need to be reversed like this so we can look up the key based on the text
    mcolAvailableComponents.Add strCode & CStr(intID), objNode.Tag
  End If

End Sub

Public Sub PopulateComponents()

  'Key Prefixes:
  '
  ' C = Column
  ' E = Expression
  ' X = Special/Custom/Function
  ' Z = Heading (no action)
  '

  Set mcolAvailableComponents = New Collection

  'Dim objParent As SSNode
  Dim objParent As Node

  With sstrvAvailable
    .ImageList = ImageList1
    .Nodes.Clear

    PopulateColumnNodes mlngTableID, mstrTableName

    With recRelEdit
      .index = "idxChildID"
      .MoveFirst
      .Seek "=", mlngTableID
    
      If Not .NoMatch Then
        Do While (Not .EOF)
          If !childID <> mlngTableID Then
            Exit Do
          End If
        
          PopulateColumnNodes !parentID, GetTableName(!parentID)

          .MoveNext
        Loop
      End If
    End With
    
    
    PopulateCalculationNodes
    
    
    ' Set objParent = sstrvAvailable.Nodes.Add(, , "ZFunctions", "Functions", "IMG_CALC", "IMG_CALC")
    Set objParent = sstrvAvailable.Nodes.Add(, , "ZFunctions", "Functions", "IMG_CUSTOM", "IMG_CUSTOM")
    objParent.Expanded = True
    ' AddNode objParent, "X", 0, "Function: ", "Current User", "IMG_CALC"
    AddNode objParent, "X", 0, "Function: ", "Current User", "IMG_CUSTOM"
    'AddNode objParent, "X", 1, "Old Column Value", "IMG_CALC"
  
  End With
  
End Sub


Private Sub PopulateColumnNodes(lngTableID As Long, strTableName As String)

  'Dim objParent As SSNode
  Dim objParent As Node
  
  Set objParent = sstrvAvailable.Nodes.Add(, , "T" & CStr(lngTableID), strTableName, "IMG_TABLE", "IMG_TABLE")
  objParent.Expanded = True
  With recColEdit
    .index = "idxName"
    .Seek ">=", lngTableID

    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> lngTableID Then
          Exit Do
        End If


        If (Not !Deleted) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
          (!ControlType <> giCTRL_OLE) And _
          (!ControlType <> giCTRL_PHOTO) And _
          (!ControlType <> giCTRL_LINK) Then

          AddNode objParent, "C", !ColumnID, strTableName & ".", !ColumnName, "IMG_TABLE"

        End If

        .MoveNext
      Loop
    End If
  End With

  If objParent.Children = 0 Then
    sstrvAvailable.Nodes.Remove objParent.index
  End If

End Sub


Private Sub PopulateColumns(lngTableID As Long)

  lstColumnLinkColumns.Clear
  cboDateLinkColumn.Clear
  
  With recColEdit
    .index = "idxName"
    .Seek ">=", lngTableID

    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> lngTableID Then
          Exit Do
        End If


        If (Not !Deleted) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
          (!ControlType <> giCTRL_OLE) And _
          (!ControlType <> giCTRL_PHOTO) And _
          (!ControlType <> giCTRL_LINK) Then

          lstColumnLinkColumns.AddItem (!ColumnName)
          lstColumnLinkColumns.ItemData(lstColumnLinkColumns.NewIndex) = !ColumnID
          lstColumnLinkColumns.Selected(lstColumnLinkColumns.NewIndex) = mobjTempLink.IsColumnSelected(!ColumnID)

          If !DataType = sqlDate Then
            cboDateLinkColumn.AddItem !ColumnName
            cboDateLinkColumn.ItemData(cboDateLinkColumn.NewIndex) = !ColumnID
          End If

        End If

        .MoveNext
      Loop
    End If
  End With

  With cboDateLinkColumn
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
  End With

End Sub


Private Sub PopulateCalculationNodes()

  'Dim objParent As SSNode
  Dim objParent As Node
  
  Set objParent = sstrvAvailable.Nodes.Add(, , "ZCalculations", "Calculations", "IMG_CALC", "IMG_CALC")
  objParent.Expanded = True
  With recExprEdit
    .index = "idxExprName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF

      If ((!Type = giEXPR_COLUMNCALCULATION) Or _
        (!Type = giEXPR_RECORDDESCRIPTION) Or _
        (!Type = giEXPR_EMAIL)) And _
        (!ParentComponentID = 0) And _
        (Not !Deleted) And _
        ((!UserName = gsUserName) Or (!Access <> ACCESS_HIDDEN)) And _
        (!TableID = mlngTableID) Then

        AddNode objParent, "E", !ExprID, "Calculation: ", !Name, "IMG_CALC"
      End If
      
      .MoveNext
    Loop
  
  End With
  
  If objParent.Children = 0 Then
    sstrvAvailable.Nodes.Remove objParent.index
  End If

End Sub


'MH20090520
'Private Sub sstrvAvailable_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  If Button = vbLeftButton Then
'    If Not (sstrvAvailable.SelectedItem Is Nothing) Then
'      If Not (sstrvAvailable.SelectedItem.Parent Is Nothing) Then
'        sstrvAvailable.Drag vbBeginDrag
'      End If
'    End If
'  End If
'End Sub

Private Sub sstrvAvailable_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  sstrvAvailable.SelectedItem = sstrvAvailable.HitTest(x, y)
  
  If Button = vbLeftButton Then
    If Not (sstrvAvailable.SelectedItem Is Nothing) Then
      If Not (sstrvAvailable.SelectedItem.Parent Is Nothing) Then
        sstrvAvailable.Drag vbBeginDrag
      End If
    End If
  End If

End Sub


Private Sub sstrvAvailable_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  'Key Prefixes:
  '
  ' C = Column
  ' E = Expression
  ' X = Special/Custom/Function
  ' Z = Heading (no action)
  
  If sstrvAvailable.SelectedItem Is Nothing Then
    sstrvAvailable.SelectedItem = sstrvAvailable.HitTest(x, y)
  End If
  
  If Not (sstrvAvailable.SelectedItem Is Nothing) Then
    Select Case Left(sstrvAvailable.SelectedItem.key, 1)
      Case "X"
          Source.DragIcon = picDocument(0).Picture
      Case "C"
          Source.DragIcon = picDocument(1).Picture
      Case "E"
          Source.DragIcon = picDocument(2).Picture
      Case Else
          Source.DragIcon = picDocument(0).Picture
    End Select
  End If
End Sub

Private Sub frmContent_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Source.DragIcon = picNoDrop.Picture
End Sub

Private Sub txtContent_DragOver(index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  If TypeOf Source Is TreeView Then
    
    'sstrvAvailable.SelectedItem = sstrvAvailable.HitTest(x, y)
    ' txtContent.SelectedItem = txtContent.HitTest(x, y)
    
    If Not (sstrvAvailable.SelectedItem Is Nothing) Then
      Select Case Left(sstrvAvailable.SelectedItem.key, 1)
        Case "X"
            Source.DragIcon = picDocument(0).Picture
        Case "C"
            Source.DragIcon = picDocument(1).Picture
        Case "E"
            Source.DragIcon = picDocument(2).Picture
        Case Else
            Source.DragIcon = picDocument(0).Picture
      End Select
    End If
    
    ' Source.DragIcon = picDocument(0).Picture
    txtContent(index).SelStart = TextBoxCursorPos(txtContent(index), x, y)
    txtContent(index).SelLength = 0
  Else
    Source.DragIcon = picNoDrop.Picture
  End If
End Sub

Private Sub txtContent_DragDrop(index As Integer, Source As Control, x As Single, y As Single)

  Dim strFieldText As String
  Dim lngStart As Long

  If TypeOf Source Is TreeView Then
    lngStart = TextBoxCursorPos(txtContent(index), x, y)
    InsertColumn index, lngStart
  End If

End Sub


Private Sub InsertColumn(index As Integer, lngStart As Long)
  Dim lngStartMergePoint As Long
  Dim lngEndMergePoint As Long
  Dim strNodeText As String

  strNodeText = ""
  If Not sstrvAvailable.SelectedItem Is Nothing Then
    strNodeText = sstrvAvailable.SelectedItem.Tag
  End If
  
  If strNodeText <> "" Then
    If lngStart > 0 Then
      lngStartMergePoint = InStr(lngStart, txtContent(index).Text, strDelimStart)
      lngEndMergePoint = InStr(lngStart, txtContent(index).Text, strDelimStop)
      If (lngStartMergePoint = 0 And lngEndMergePoint > 0) Or _
          (lngStartMergePoint > lngEndMergePoint) Then
        lngStart = lngEndMergePoint
      End If
    End If

    txtContent(index).SelStart = lngStart
    txtContent(index).SelLength = 0
    txtContent(index).SelText = strDelimStart & strNodeText & strDelimStop
  End If

End Sub


Private Function GetEmailAddressName(lngRecipientID As Long) As String
  
  On Error GoTo ErrorTrap
  
  GetEmailAddressName = vbNullString
  
  With recEmailAddrEdit
    .index = "idxID"
    .Seek "=", lngRecipientID
    If Not .NoMatch Then
      If Not !Deleted Then
        GetEmailAddressName = !Name
      End If
    End If
  End With
  
  Exit Function
  
ErrorTrap:

End Function


