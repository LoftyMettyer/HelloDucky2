VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmCMGSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CMG & Centrefile Setup"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5078
   Icon            =   "frmCMGSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkReverseOutput 
      Caption         =   "Reverse Output / Last Change Date"
      Height          =   195
      Left            =   315
      TabIndex        =   9
      Top             =   3555
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Frame frmOutput 
      Caption         =   "Output : "
      Height          =   1380
      Left            =   120
      TabIndex        =   28
      Top             =   2535
      Width           =   7215
      Begin VB.CheckBox chkIgnoreEmptyFields 
         Caption         =   "I&gnore Empty Fields"
         Height          =   285
         Left            =   195
         TabIndex        =   8
         Top             =   630
         Width           =   2850
      End
      Begin VB.CheckBox chkUseCSV 
         Caption         =   "Use Comma &Separated Values"
         Height          =   240
         Left            =   195
         TabIndex        =   7
         Top             =   315
         Width           =   3075
      End
   End
   Begin VB.Frame frmCMG 
      Caption         =   "Layout : "
      Height          =   2355
      Left            =   120
      TabIndex        =   18
      Top             =   135
      Width           =   7215
      Begin VB.CommandButton cmdDefault 
         Caption         =   "De&fault Order"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5640
         TabIndex        =   6
         Top             =   1740
         Width           =   1380
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5640
         TabIndex        =   3
         Top             =   300
         Width           =   1380
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move &Down"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5640
         TabIndex        =   5
         Top             =   1260
         Width           =   1380
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move &Up"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5640
         TabIndex        =   4
         Top             =   780
         Width           =   1380
      End
      Begin VB.CheckBox chkExportLastDate 
         Caption         =   "Expo&rt Last Change Date"
         Height          =   300
         Left            =   195
         TabIndex        =   16
         Top             =   4545
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CheckBox chkExportFileCode 
         Caption         =   "&Export File Code"
         Height          =   300
         Left            =   195
         TabIndex        =   10
         Top             =   3015
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chkExportFieldCode 
         Caption         =   "Ex&port Field Code"
         Height          =   300
         Left            =   195
         TabIndex        =   13
         Top             =   3780
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   195
         Picture         =   "frmCMGSetup.frx":000C
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   20
         Top             =   4200
         Width           =   225
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   195
         Picture         =   "frmCMGSetup.frx":0156
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   19
         Top             =   3420
         Width           =   225
      End
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Index           =   4
         Left            =   3960
         TabIndex        =   17
         Top             =   4545
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
         Enabled         =   0   'False
         MaximumValue    =   100
         Text            =   "0"
      End
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
         Enabled         =   0   'False
         MaximumValue    =   100
         Text            =   "0"
      End
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   15
         Top             =   4155
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   3390
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
      Begin COASpinner.COA_Spinner spnMaxSize 
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   14
         Top             =   3765
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
         Enabled         =   0   'False
         MaximumValue    =   100
         Text            =   "0"
      End
      Begin SSDataWidgets_B.SSDBGrid grdCMGLayout 
         Height          =   1840
         Left            =   195
         TabIndex        =   2
         Top             =   300
         Width           =   5295
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         GroupHeadLines  =   0
         Col.Count       =   3
         stylesets.count =   6
         stylesets(0).Name=   "ssetHeaderDisabled"
         stylesets(0).ForeColor=   -2147483631
         stylesets(0).BackColor=   -2147483633
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmCMGSetup.frx":02A0
         stylesets(1).Name=   "ssetSelected"
         stylesets(1).ForeColor=   -2147483634
         stylesets(1).BackColor=   -2147483635
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frmCMGSetup.frx":02BC
         stylesets(2).Name=   "ssetEnabled"
         stylesets(2).ForeColor=   -2147483640
         stylesets(2).BackColor=   -2147483643
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "frmCMGSetup.frx":02D8
         stylesets(3).Name=   "ssetHeaderEnabled"
         stylesets(3).ForeColor=   -2147483630
         stylesets(3).BackColor=   -2147483633
         stylesets(3).HasFont=   -1  'True
         BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(3).Picture=   "frmCMGSetup.frx":02F4
         stylesets(4).Name=   "ssetCheckBoxSelected"
         stylesets(4).ForeColor=   -2147483640
         stylesets(4).BackColor=   -2147483635
         stylesets(4).HasFont=   -1  'True
         BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(4).Picture=   "frmCMGSetup.frx":0310
         stylesets(5).Name=   "ssetDisabled"
         stylesets(5).ForeColor=   -2147483631
         stylesets(5).BackColor=   -2147483633
         stylesets(5).HasFont=   -1  'True
         BeginProperty stylesets(5).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(5).Picture=   "frmCMGSetup.frx":032C
         CheckBox3D      =   0   'False
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   1
         StyleSet        =   "ssetEnabled"
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         ActiveRowStyleSet=   "ssetSelected"
         Columns.Count   =   3
         Columns(0).Width=   5556
         Columns(0).Caption=   "Export Item"
         Columns(0).Name =   "Export Item"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(0).HasBackColor=   -1  'True
         Columns(0).BackColor=   16777215
         Columns(1).Width=   2143
         Columns(1).Caption=   "Field Size"
         Columns(1).Name =   "Field Size"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(1).HasBackColor=   -1  'True
         Columns(1).BackColor=   16777215
         Columns(2).Width=   1640
         Columns(2).Caption=   "Exported"
         Columns(2).Name =   "Exported"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   2
         Columns(2).HasBackColor=   -1  'True
         Columns(2).BackColor=   16777215
         TabNavigation   =   1
         _ExtentX        =   9340
         _ExtentY        =   3246
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSize 
         Caption         =   "Field Size :"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3075
         TabIndex        =   27
         Top             =   4590
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSize 
         Caption         =   "Field Size :"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   3075
         TabIndex        =   26
         Top             =   3045
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSize 
         Caption         =   "Field Size :"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3075
         TabIndex        =   25
         Top             =   3825
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblOutputColumn 
         Caption         =   "Output Column"
         Height          =   300
         Left            =   480
         TabIndex        =   24
         Top             =   4215
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label lblSize 
         Caption         =   "Field Size :"
         Height          =   315
         Index           =   3
         Left            =   3075
         TabIndex        =   23
         Top             =   4215
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSize 
         Caption         =   "Field Size :"
         Height          =   315
         Index           =   1
         Left            =   3075
         TabIndex        =   22
         Top             =   3450
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblRecordID 
         Caption         =   "Record Identifier"
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   3450
         Visible         =   0   'False
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6105
      TabIndex        =   1
      Top             =   4050
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4815
      TabIndex        =   0
      Top             =   4050
      Width           =   1200
   End
End
Attribute VB_Name = "frmCMGSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfChanged As Boolean
Private mbReadOnly As Boolean

' Load the CMG settings
Private Sub LoadSettings()
  Dim iRow As Integer

  chkUseCSV.value = IIf(gbCMGExportUseCSV, 1, 0)
  chkIgnoreEmptyFields.value = IIf(gbCMGIgnoreBlanks, 1, 0)
  chkReverseOutput.value = IIf(gbCMGReverseDateChanged, 1, 0)
  chkExportFileCode.value = IIf(gbCMGExportFileCode, 1, 0)
  chkExportFieldCode.value = IIf(gbCMGExportFieldCode, 1, 0)
  chkExportLastDate.value = IIf(gbCMGExportLastChangeDate, 1, 0)
  spnMaxSize(0).value = giCMGExportFileCodeSize
  spnMaxSize(1).value = giCMGEXportRecordIDSize
  spnMaxSize(2).value = giCMGExportFieldCodeSize
  spnMaxSize(3).value = giCMGExportOutputColumnSize
  spnMaxSize(4).value = giCMGExportLastChangeDateSize
  
  'NPG20090313 Fault 13595
  If giCMGEXportRecordIDOrderID + giCMGExportFileCodeOrderID + giCMGExportFieldCodeOrderID + giCMGExportOutputColumnOrderID + giCMGExportLastChangeDateOrderID = 0 Then
    ' This is the first time CMG setup has been run on this version
    ' so use default order and use reverseoutput option if required
    Me.grdCMGLayout.AddItem "Export File Code" & vbTab & CStr(giCMGExportFileCodeSize) & vbTab & IIf(gbCMGExportFileCode, True, False)
    Me.grdCMGLayout.AddItem "Record Identifier" & vbTab & CStr(giCMGEXportRecordIDSize) & vbTab & True
    Me.grdCMGLayout.AddItem "Export Field Code" & vbTab & CStr(giCMGExportFieldCodeSize) & vbTab & IIf(gbCMGExportFieldCode, True, False)
    If gbCMGReverseDateChanged Then
      Me.grdCMGLayout.AddItem "Export Last Change Date" & vbTab & CStr(giCMGExportLastChangeDateSize) & vbTab & IIf(gbCMGExportLastChangeDate, True, False)
      Me.grdCMGLayout.AddItem "Output Column" & vbTab & CStr(giCMGExportOutputColumnSize) & vbTab & True
    Else
      Me.grdCMGLayout.AddItem "Output Column" & vbTab & CStr(giCMGExportOutputColumnSize) & vbTab & True
      Me.grdCMGLayout.AddItem "Export Last Change Date" & vbTab & CStr(giCMGExportLastChangeDateSize) & vbTab & IIf(gbCMGExportLastChangeDate, True, False)
    End If
  Else
  ' Use saved settings
    For iRow = 0 To 4
      If giCMGEXportRecordIDOrderID = iRow Then Me.grdCMGLayout.AddItem "Record Identifier" & vbTab & CStr(giCMGEXportRecordIDSize) & vbTab & True
      If giCMGExportFileCodeOrderID = iRow Then Me.grdCMGLayout.AddItem "Export File Code" & vbTab & CStr(giCMGExportFileCodeSize) & vbTab & IIf(gbCMGExportFileCode, True, False)
      If giCMGExportFieldCodeOrderID = iRow Then Me.grdCMGLayout.AddItem "Export Field Code" & vbTab & CStr(giCMGExportFieldCodeSize) & vbTab & IIf(gbCMGExportFieldCode, True, False)
      If giCMGExportOutputColumnOrderID = iRow Then Me.grdCMGLayout.AddItem "Output Column" & vbTab & CStr(giCMGExportOutputColumnSize) & vbTab & True
      If giCMGExportLastChangeDateOrderID = iRow Then Me.grdCMGLayout.AddItem "Export Last Change Date" & vbTab & CStr(giCMGExportLastChangeDateSize) & vbTab & IIf(gbCMGExportLastChangeDate, True, False)
    Next iRow
  End If
  
End Sub
Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
End Property
Private Function ValidateSetup() As Boolean
  ValidateSetup = True
End Function

' Save the CMG Export Details
Private Function SaveChanges() As Boolean
  Dim iRow As Integer
  'AE20071119 Fault #12607
  SaveChanges = False
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass

  gbCMGExportUseCSV = chkUseCSV.value
  gbCMGIgnoreBlanks = chkIgnoreEmptyFields.value
  gbCMGReverseDateChanged = chkReverseOutput.value
  
  gbCMGExportFileCode = chkExportFileCode.value
  gbCMGExportFieldCode = chkExportFieldCode.value
  gbCMGExportLastChangeDate = chkExportLastDate.value
  giCMGExportFileCodeSize = spnMaxSize(0).value
  giCMGEXportRecordIDSize = spnMaxSize(1).value
  giCMGExportFieldCodeSize = spnMaxSize(2).value
  giCMGExportOutputColumnSize = spnMaxSize(3).value
  giCMGExportLastChangeDateSize = spnMaxSize(4).value

  'NPG20090313 Fault 13595
  With grdCMGLayout
    .Redraw = False
    If .Rows > 0 Then
      .MoveFirst
    End If
  
    For iRow = 0 To .Rows - 1
        Select Case .Columns(0).Text
        
          Case "Record Identifier"
            giCMGEXportRecordIDOrderID = iRow
            giCMGEXportRecordIDSize = .Columns(1).value
            
          Case "Export File Code"
            giCMGExportFileCodeOrderID = iRow
            gbCMGExportFileCode = .Columns(2).value
            giCMGExportFileCodeSize = .Columns(1).value
            
          Case "Export Field Code"
            giCMGExportFieldCodeOrderID = iRow
            gbCMGExportFieldCode = .Columns(2).value
            giCMGExportFieldCodeSize = .Columns(1).value
            
          Case "Output Column"
            giCMGExportOutputColumnOrderID = iRow
            giCMGExportOutputColumnSize = .Columns(1).value
            
          Case "Export Last Change Date"
            giCMGExportLastChangeDateOrderID = iRow
            gbCMGExportLastChangeDate = .Columns(2).value
            giCMGExportLastChangeDateSize = .Columns(1).value
            
        End Select
      .MoveNext
    Next iRow
    
    If .Rows > 0 Then
      .MoveFirst
    End If
    .Redraw = True
  End With
  
  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbNormal
End Function

Private Sub chkExportFieldCode_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkExportFileCode_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkExportLastDate_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkIgnoreEmptyFields_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkReverseOutput_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkUseCSV_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub cmdCancel_Click()
  'AE20071119 Fault #12607
'  Dim pintAnswer As Integer
'    If Changed = True And cmdOk.Enabled Then
'      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
'      If pintAnswer = vbYes Then
'        'AE20071108 Fault #12551
'        'Using Me.MousePointer = vbNormal forces the form to be reloaded
'        'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
'        'Me.MousePointer = vbHourglass
'        Screen.MousePointer = vbHourglass
'        cmdOK_Click 'This is just like saving
'        'Me.MousePointer = vbNormal
'        Screen.MousePointer = vbNormal
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdDefault_Click()
  Dim pintAnswer As Integer
  pintAnswer = MsgBox("This will reset the export to the default order" & _
            vbCrLf & "Do you want to continue ?", vbQuestion + vbYesNo, App.Title)
  If pintAnswer = vbYes Then
    Me.grdCMGLayout.RemoveAll
    
    Me.grdCMGLayout.AddItem "Export File Code" & vbTab & CStr(giCMGExportFileCodeSize) & vbTab & IIf(gbCMGExportFileCode, True, False)
    Me.grdCMGLayout.AddItem "Record Identifier" & vbTab & CStr(giCMGEXportRecordIDSize) & vbTab & True
    Me.grdCMGLayout.AddItem "Export Field Code" & vbTab & CStr(giCMGExportFieldCodeSize) & vbTab & IIf(gbCMGExportFieldCode, True, False)
    Me.grdCMGLayout.AddItem "Output Column" & vbTab & CStr(giCMGExportOutputColumnSize) & vbTab & True
    Me.grdCMGLayout.AddItem "Export Last Change Date" & vbTab & CStr(giCMGExportLastChangeDateSize) & vbTab & IIf(gbCMGExportLastChangeDate, True, False)
  
    Me.grdCMGLayout.Refresh
    
    Changed = True
    RefreshButtons
    
  End If
End Sub

Private Sub cmdEdit_Click()
  frmCMGEdit.Initialise grdCMGLayout.Columns(0).Text, grdCMGLayout.Columns(1).Text, grdCMGLayout.Columns(2).Text, True, Me
  
  If (frmCMGEdit.UserChanged And Not frmCMGEdit.UserCancelled) Then
    Changed = True
    RefreshButtons
  End If

  UnLoad frmCMGEdit
  
End Sub

Private Sub cmdMoveDown_Click()
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdCMGLayout.AddItemRowIndex(grdCMGLayout.Bookmark)
  strSourceRow = grdCMGLayout.Columns(0).Text & vbTab & grdCMGLayout.Columns(1).Text & vbTab & grdCMGLayout.Columns(2).Text
  
  intDestinationRow = intSourceRow + 1
  grdCMGLayout.MoveNext
  strDestinationRow = grdCMGLayout.Columns(0).Text & vbTab & grdCMGLayout.Columns(1).Text & vbTab & grdCMGLayout.Columns(2).Text
  
  grdCMGLayout.RemoveItem intDestinationRow
  grdCMGLayout.RemoveItem intSourceRow
  
  grdCMGLayout.AddItem strDestinationRow, intSourceRow
  grdCMGLayout.AddItem strSourceRow, intDestinationRow
  
  grdCMGLayout.SelBookmarks.RemoveAll
  grdCMGLayout.MoveNext
  grdCMGLayout.Bookmark = grdCMGLayout.AddItemBookmark(intDestinationRow)
  'grdCMGLayout.SelBookmarks.Add grdCMGLayout.AddItemBookmark(intDestinationRow)
  
  UpdateButtonStatus

  Changed = True
  RefreshButtons
  
End Sub

Private Sub cmdMoveUp_Click()
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdCMGLayout.AddItemRowIndex(grdCMGLayout.Bookmark)
  strSourceRow = grdCMGLayout.Columns(0).Text & vbTab & grdCMGLayout.Columns(1).Text & vbTab & grdCMGLayout.Columns(2).Text
  
  intDestinationRow = intSourceRow - 1
  grdCMGLayout.MovePrevious
  strDestinationRow = grdCMGLayout.Columns(0).Text & vbTab & grdCMGLayout.Columns(1).Text & vbTab & grdCMGLayout.Columns(2).Text
  
  grdCMGLayout.RemoveItem intSourceRow
  grdCMGLayout.RemoveItem intDestinationRow
  
  grdCMGLayout.AddItem strSourceRow, intDestinationRow
  grdCMGLayout.AddItem strDestinationRow, intSourceRow
  
  grdCMGLayout.SelBookmarks.RemoveAll
  grdCMGLayout.MovePrevious
  grdCMGLayout.Bookmark = grdCMGLayout.AddItemBookmark(intDestinationRow)
  'grdCMGLayout.SelBookmarks.Add grdCMGLayout.AddItemBookmark(intDestinationRow)

  UpdateButtonStatus
  
  Changed = True
  RefreshButtons
  
End Sub



Private Sub cmdOK_Click()
  
  'AE20071119 Fault #12607
  'If ValidateSetup Then
    'SaveChanges
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If

End Sub

Private Sub Form_Activate()
  UpdateButtonStatus
End Sub

Private Sub Form_Load()
  
  LoadSettings
  
  ' Disable if read only
  mbReadOnly = (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
  ControlsDisableAll Me, Not mbReadOnly
      
  Changed = False
  RefreshButtons
    
End Sub

Private Sub RefreshButtons()

  Dim bCMGExportUseCSV As Boolean
  
  If Not mbReadOnly Then
    bCMGExportUseCSV = (chkUseCSV.value = vbChecked)
  
    spnMaxSize(0).Enabled = IIf(chkExportFileCode.value = vbChecked, True, False) And Not bCMGExportUseCSV
    spnMaxSize(1).Enabled = Not bCMGExportUseCSV
    spnMaxSize(2).Enabled = IIf(chkExportFieldCode.value = vbChecked, True, False) And Not bCMGExportUseCSV
    spnMaxSize(3).Enabled = Not bCMGExportUseCSV
    spnMaxSize(4).Enabled = IIf(chkExportLastDate.value = vbChecked, True, False) And Not bCMGExportUseCSV
    
    lblSize(0).Enabled = spnMaxSize(0).Enabled
    lblSize(1).Enabled = spnMaxSize(1).Enabled
    lblSize(2).Enabled = spnMaxSize(2).Enabled
    lblSize(3).Enabled = spnMaxSize(3).Enabled
    lblSize(4).Enabled = spnMaxSize(4).Enabled
  
    cmdOk.Enabled = mfChanged
  End If

End Sub

Private Function UpdateButtonStatus()

  On Error Resume Next
  
  Dim tempItem As ListItem, iCount As Integer
  
    ' disable move up button if required
    cmdMoveUp.Enabled = (grdCMGLayout.AddItemRowIndex(grdCMGLayout.Bookmark) > 0)
    cmdMoveDown.Enabled = (grdCMGLayout.AddItemRowIndex(grdCMGLayout.Bookmark) < 4)

  'TM20020508 Fault 3790
  '   Call CheckListViewColWidth(ListView1)
  ' Call CheckListViewColWidth(ListView2)

  grdCMGLayout.Refresh
  
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the user cancels or tries to close the form
  'AE20071119 Fault #12607
  'If UnloadMode <> vbFormCode And cmdOK.Enabled Then
  If Changed Then
    Select Case MsgBox("Apply module changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        'AE20071119 Fault #12607
        'SaveChanges
        Cancel = (Not SaveChanges)
    End Select
  End If
End Sub


Private Sub grdCMGLayout_Click()
  UpdateButtonStatus
End Sub

Private Sub grdCMGLayout_DblClick()
  frmCMGEdit.Initialise grdCMGLayout.Columns(0).Text, grdCMGLayout.Columns(1).Text, grdCMGLayout.Columns(2).Text, True, Me
'  UpdateOrderButtons
'
 'AE20071025 Fault #6797
  If (frmCMGEdit.UserChanged And Not frmCMGEdit.UserCancelled) Then
    Changed = True
    RefreshButtons
  End If

  UnLoad frmCMGEdit
  
End Sub

Private Sub spnMaxSize_Change(Index As Integer)
  Changed = True
  RefreshButtons
End Sub

Private Sub RefreshReportOrderGrid()
'
'  With grdReportOrder
'    .Refresh
'    .Enabled = True
'    .AllowUpdate = (Not mblnReadOnly)
'    .Columns(1).Locked = True
'    .CheckBox3D = False
'
'    If mblnReadOnly Then
'      .HeadStyleSet = "ssetHeaderDisabled"
'      .StyleSet = "ssetDisabled"
'      .ActiveRowStyleSet = "ssetDisabled"
'      .SelectTypeRow = ssSelectionTypeNone
'      .SelectTypeCol = ssSelectionTypeNone
'      .RowNavigation = ssRowNavigationAllLock
'    Else
'      .HeadStyleSet = "ssetHeaderEnabled"
'      .StyleSet = "ssetEnabled"
'      .ActiveRowStyleSet = "ssetSelected"
'      .SelectTypeRow = ssSelectionTypeSingleSelect
'      .SelectByCell = False
'      .SelectTypeCol = ssSelectionTypeNone
'      .RowNavigation = ssRowNavigationLRLock
'
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'    End If
'
'  End With
'
'  grdReportOrder_RowColChange 0, 0
'
'  UpdateOrderButtons
'
End Sub

