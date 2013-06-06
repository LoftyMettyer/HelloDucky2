VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diary Events"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1028
   Icon            =   "frmDiary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid grdPrint 
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   7845
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Alarm"
      stylesets(0).ForeColor=   -2147483643
      stylesets(0).BackColor=   -2147483643
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
      stylesets(0).Picture=   "frmDiary.frx":000C
      stylesets(0).AlignmentPicture=   4
      stylesets(0).PictureMetaWidth=   35
      stylesets(0).PictureMetaHeight=   35
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
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   -2147483640
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "DiaryEventID"
      Columns(0).Name =   "DiaryEventID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   4
      Columns(0).StyleSet=   "Alarm"
      Columns(1).Width=   873
      Columns(1).Caption=   "Alarm"
      Columns(1).Name =   "Alarm"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2540
      Columns(2).Caption=   "Date / Time"
      Columns(2).Name =   "Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   10372
      Columns(3).Caption=   "Title"
      Columns(3).Name =   "Title"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   13838
      _ExtentY        =   5318
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame fraViewType 
      BorderStyle     =   0  'None
      Caption         =   "View By Day"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9030
      Begin SSDataWidgets_B.SSDBGrid grdViewByDay 
         Height          =   4100
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   8820
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":05A6
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   15557
         _ExtentY        =   7232
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label lblSelectedDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblSelectedDay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1830
      End
   End
   Begin VB.Frame fraViewType 
      BorderStyle     =   0  'None
      Caption         =   "View By Week"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9030
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   0
         Left            =   645
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":0995
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   1
         Left            =   645
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":0D84
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   2
         Left            =   645
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":1173
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   3
         Left            =   4080
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":1562
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   4
         Left            =   4080
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":1951
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   5
         Left            =   4080
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":1D40
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBGrid grdViewByWeek 
         Height          =   975
         Index           =   6
         Left            =   4080
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3345
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   4
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).BackColor=   16184819
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
         stylesets(0).Picture=   "frmDiary.frx":212F
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
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
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Alarm"
         Columns(2).Width=   900
         Columns(2).Caption=   "EventTime"
         Columns(2).Name =   "EventTime"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   12356
         Columns(3).Caption=   "EventTitle"
         Columns(3).Name =   "EventTitle"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   5900
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   7530
         TabIndex        =   13
         Top             =   360
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   7530
         TabIndex        =   12
         Top             =   1440
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   7530
         TabIndex        =   11
         Top             =   2520
         Width           =   520
      End
      Begin VB.Label lblDayTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   7530
         TabIndex        =   10
         Top             =   3600
         Width           =   520
      End
      Begin VB.Label lblSelectedWeek 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblSelectedWeek"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2040
      End
      Begin VB.Label lblWeekNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblWeekNo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.Frame fraViewType 
      BorderStyle     =   0  'None
      Caption         =   "View By Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9030
      Begin MSComCtl2.MonthView mvwViewbyMonth 
         Height          =   4515
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   7964
         _Version        =   393216
         ForeColor       =   6697779
         BackColor       =   16184819
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
         MaxSelCount     =   60
         MonthColumns    =   3
         MonthRows       =   2
         MonthBackColor  =   16184819
         ScrollRate      =   1
         StartOfWeek     =   139984897
         TitleBackColor  =   6697779
         TitleForeColor  =   15988214
         TrailingForeColor=   -2147483643
         CurrentDate     =   36526
         MaxDate         =   401768
         MinDate         =   2
      End
   End
   Begin VB.Frame fraViewType 
      BorderStyle     =   0  'None
      Caption         =   "View List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8970
      Begin SSDataWidgets_B.SSDBGrid grdViewByList 
         Height          =   4575
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   8820
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         stylesets.count =   1
         stylesets(0).Name=   "Alarm"
         stylesets(0).ForeColor=   -2147483643
         stylesets(0).BackColor=   14811135
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
         stylesets(0).Picture=   "frmDiary.frx":251E
         stylesets(0).AlignmentPicture=   0
         stylesets(0).PictureMetaWidth=   28
         stylesets(0).PictureMetaHeight=   28
         DividerType     =   0
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         CellNavigation  =   1
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         BackColorEven   =   16184819
         BackColorOdd    =   16184819
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "DiaryEventID"
         Columns(0).Name =   "DiaryEventID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(0).StyleSet=   "Alarm"
         Columns(1).Width=   529
         Columns(1).Caption=   "Alarm"
         Columns(1).Name =   "Alarm"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3360
         Columns(2).Caption=   "Date / Time"
         Columns(2).Name =   "Date"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   10583
         Columns(3).Caption=   "Title"
         Columns(3).Name =   "Title"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   15557
         _ExtentY        =   8070
         _StockProps     =   79
         BackColor       =   16184819
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmDiary.frx":290D
   End
End
Attribute VB_Name = "frmDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mintVIEWBYDAY As Integer = 0
Private Const mintVIEWBYWEEK As Integer = 1
Private Const mintVIEWBYMONTH As Integer = 2
Private Const mintVIEWBYLIST As Integer = 3

Private DiaryPrint As clsPrintGrid
Private mdtLastStartDate As Date
Private mintCurrentView As Integer


Public Sub Initialise()

  Dim lngViewMode As Long
  
  frmMain.tmrDiary.Enabled = False
  gobjDiary.DiaryEventID = 0

  If Not gobjDiary.ViewingAlarms Then
    'gobjDiary.FilterEventType = GetUserSetting("Diary", "FilterEventType", 0)
    'gobjDiary.FilterAlarmStatus = GetUserSetting("Diary", "FilterAlarmStatus", 0)
    'gobjDiary.FilterPastPresent = GetUserSetting("Diary", "FilterPastPresent", 0)
    'gobjDiary.FilterOnlyMine = GetUserSetting("Diary", "OnlyMine ManualEvents", False)
    gobjDiary.FilterEventType = 0
    gobjDiary.FilterAlarmStatus = 0
    gobjDiary.FilterPastPresent = 0
    gobjDiary.FilterOnlyMine = False
    lngViewMode = GetUserSetting("Diary", "ViewMode", 1)
  Else
    lngViewMode = gobjDiary.CurrentView
  End If

  frmDiary.Caption = gobjDiary.FilterText
  ChangeView (lngViewMode)  'This does a refresh too


fraViewType(lngViewMode).ZOrder 0

End Sub


Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  'TM20011106 Fault 3100
  Cancel = True
End Sub

Private Sub Form_GotFocus()

  Select Case gobjDiary.CurrentView
  Case mintVIEWBYDAY
    grdViewByDay.SetFocus
  
  Case mintVIEWBYMONTH
    mvwViewbyMonth.SetFocus
  
  Case mintVIEWBYLIST
    grdViewByList.SetFocus
  
  End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim fManyEvents As Boolean
  
Select Case KeyCode
Case vbKeyF1
  If ShowAirHelp(Me.HelpContextID) Then
    KeyCode = 0
  End If
End Select

  'Holding CTRL
  If Shift And vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyN
      If DiaryToolBar("New").Enabled Then
        gobjDiary.DiaryEventID = 0   'New Record
        frmDiaryDetail.Show vbModal
        gobjDiary.RefreshDiaryData
      End If

    Case vbKeyR
      If DiaryToolBar("Repeat").Enabled Then
        If EventStillExists Then
          frmDiaryDuplicate.Show vbModal
          gobjDiary.RefreshDiaryData
        End If
      End If

    Case vbKeyE
      If DiaryToolBar("Edit").Enabled Then
        Call DiaryEditEvent
        gobjDiary.RefreshDiaryData
      End If
    
    Case vbKeyD
      If DiaryToolBar("Delete").Enabled Then
        gobjDiary.DeleteCurrentEntry
        gobjDiary.RefreshDiaryData
      End If
  
    Case vbKeyX
      If DiaryToolBar("Cut").Enabled Then
        gobjDiary.CutEntries blnDelete:=True
        gobjDiary.RefreshDiaryData
      End If

    Case vbKeyC
      If DiaryToolBar("Copy").Enabled Then
        gobjDiary.CutEntries blnDelete:=False
        gobjDiary.RefreshDiaryData
      End If

    Case vbKeyV
      If DiaryToolBar("Paste").Enabled Then
        gobjDiary.PasteEntries
        gobjDiary.RefreshDiaryData
      End If

    Case vbKeyP
      If DiaryToolBar("Print").Enabled Then
        Call PrintDiaryEvents
      End If

    Case vbKeyG
      If DiaryToolBar("Goto").Enabled Then
        frmDiaryGoTo.Show vbModal
        gobjDiary.RefreshDiaryData
      End If

    Case vbKeyL
      If DiaryToolBar("Alarm").Enabled Then
        frmDiaryAlarmSet.Show vbModal
        gobjDiary.RefreshDiaryData
      End If
      
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    End Select
  End If
  
  'Don't need CTRL
  Select Case KeyCode
  'Case vbKeyF4
  'NHRD28112002 Fault 4648 used vbkeyF3 instead as this
  'causes less confusion with the other F4 functions
  Case vbKeyF3
    If DiaryToolBar("Edit").Enabled Then
      Call DiaryEditEvent
      gobjDiary.RefreshDiaryData
    End If
  
  Case vbKeyF5
    gobjDiary.RefreshDiaryData

  Case vbKeyF7
    If DiaryToolBar("MonthView").Enabled Then
      Call ChangeView(2)
      gobjDiary.DiaryEventID = 0
      mvwViewbyMonth.Value = gobjDiary.DateSelected
      mvwViewbyMonth.SetFocus
    End If

  Case vbKeyF8
    If DiaryToolBar("WeekView").Enabled Then
      Call ChangeView(1)
    End If

  Case vbKeyF9
    If DiaryToolBar("DayView").Enabled Then
      Call ChangeView(0)
      grdViewByDay.SetFocus
      
    End If

  Case vbKeyF10
    'Debug.Print Me.ActiveControl.Name
    If DiaryToolBar("ListView").Enabled Then
      Call ChangeView(3)
      'MH20010830 Fault 2766
      Do While ActiveControl.Name <> grdViewByList.Name
        DoEvents
        grdViewByList.SetFocus
      Loop
    End If
    'Debug.Print Me.ActiveControl.Name
  
  Case vbKeyReturn
    'If SSActiveToolBars1.Tools(11).Enabled Then   'Edit
    If DiaryToolBar("Edit").Enabled Then
      Call DiaryEditEvent
      gobjDiary.RefreshDiaryData
    End If

  Case vbKeyDelete
    If DiaryToolBar("Cut").Enabled Then
      gobjDiary.CutEntries blnDelete:=True
    Else
      gobjDiary.DeleteCurrentEntry
    End If
    gobjDiary.RefreshDiaryData
    
  Case vbKeyEscape
    Unload Me

  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set frmDiary = Nothing

'  'Save the current view... for next time !
'  If Not mblnViewingAlarmedEvents Then
'    SaveUserSetting "Diary", "ViewMode", gobjDiary.CurrentView
'    SaveUserSetting "Diary", "FilterEventType", gobjDiary.FilterEventType
'    SaveUserSetting "Diary", "FilterAlarmStatus", gobjDiary.FilterAlarmStatus
'    SaveUserSetting "Diary", "FilterPastPresent", gobjDiary.FilterPastPresent
'  End If

  With frmMain
    .tmrDiary.Enabled = gblnDiaryConstCheck
    .tmrDiary.Interval = 1   'Force diary check
    .RefreshMainForm Me, True
  End With

End Sub


Private Sub grdViewByDay_Click()
  If Trim$(grdViewByDay.Columns(0).Text) <> vbNullString Then
    gobjDiary.DiaryEventID = CLng(grdViewByDay.Columns(0).Text)
  End If
End Sub

Private Sub grdViewByDay_DblClick()
  Call DiaryEditEvent
End Sub


Private Sub grdViewByDay_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmDiary.grdViewByDay_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    
  If Me.grdViewByDay.SelBookmarks.Count > 1 Then
    'Me.cmdView.Enabled = False
  ElseIf Me.grdViewByDay.SelBookmarks.Count = 1 Then
    Me.grdViewByDay.SelBookmarks.RemoveAll
    Me.grdViewByDay.SelBookmarks.Add Me.grdViewByDay.Bookmark
    'Me.cmdView.Enabled = True
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
End Sub

Private Sub grdViewByDay_RowLoaded(ByVal Bookmark As Variant)
    
  With grdViewByDay.Columns("Alarm")

    'MH20010131 Fault 1752
    'Old way was not working for french regional settings
    'If .Value <> 0 Then
    If CStr(.Value) = "True" Then
      .CellStyleSet "Alarm"
      .Text = Space$(10) & "Alarm"
    Else
      .Text = vbNullString
    End If

  End With

End Sub

Private Sub grdViewByList_Click()

  Dim lngEventID As Long
  
  lngEventID = 0
  If grdViewByList.SelBookmarks.Count = 1 Then
    If Trim$(grdViewByList.Columns(0).Text) <> vbNullString Then
      lngEventID = CLng(grdViewByList.Columns(0).Text)
    End If
  End If

  gobjDiary.DiaryEventID = lngEventID

End Sub

Private Sub grdViewByList_DblClick()
    
  ' RH 09/10/00 - Bug - workaround for blank grid line bug. do nothing if
  '                     user has clicked on a blank line. Pending response
  '                     from sheridan support. also in event log and find
  '                     window
  If Trim$(grdViewByList.Columns(0).Text) <> vbNullString Then
    Call DiaryEditEvent
  End If
  
End Sub


Private Sub grdPrint_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  Call DiaryPrint.PrintInitialise(ssPrintInfo)
  grdPrint.Width = IIf(DiaryPrint.Portrait, 7845, 13000)
  grdPrint.Columns("Title").Width = grdPrint.Width - 1965
  Me.Font.Size = 8.25
End Sub


Private Sub grdViewByList_KeyDown(KeyCode As Integer, Shift As Integer)
'NHRD25092006 Fault 11065 Was replicating the Form KeyDown so commented out
'    Form_KeyDown KeyCode, Shift
End Sub

Private Sub grdViewByList_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmDiary.grdViewByList_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    
  If Me.grdViewByList.SelBookmarks.Count > 1 Then
    'Me.cmdView.Enabled = False
  ElseIf Me.grdViewByList.SelBookmarks.Count = 1 Then
    Me.grdViewByList.SelBookmarks.RemoveAll
    Me.grdViewByList.SelBookmarks.Add Me.grdViewByList.Bookmark
    'Me.cmdView.Enabled = True
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
End Sub

Private Sub grdViewByList_RowLoaded(ByVal Bookmark As Variant)

  If Val(Bookmark) > 0 Then
    grdViewByList.RowBookmark Bookmark
    With grdViewByList.Columns("Alarm")

      If .Value <> 0 Then
        .CellStyleSet "Alarm"
        .Text = Space$(10) & "Alarm"
      Else
        .Text = vbNullString
      End If

    End With
  End If

End Sub


Private Sub grdPrint_RowPrint(ByVal Bookmark As Variant, ByVal PageNumber As Long, Cancel As Integer)

  Dim lngLen As Long

  If Val(Bookmark) > 0 Then
    grdPrint.RowBookmark Bookmark
    With grdPrint.Columns("Alarm")

      If .Text = "0" Then
        .Text = vbNullString
      Else
        .Text = "Alarm"
      End If

    End With
  
    'MH20000307
    With grdPrint.Columns("Title")
      lngLen = Len(.Text)
      Do While Me.TextWidth(.Text & "...") > .Width
        lngLen = lngLen - 1
        .Text = Left(.Text, lngLen) & "..."
      Loop
    End With

  End If

End Sub


Private Sub grdViewByList_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    StartLocation = 0
  End If
  NewLocation = CLng(StartLocation) + NumberOfRowsToMove
End Sub
Private Sub grdViewByList_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Call gobjDiary.ListUnboundReadData(RowBuf, StartLocation, ReadPriorRows)
End Sub

Private Sub grdPrint_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    StartLocation = 0
  End If
  NewLocation = CLng(StartLocation) + NumberOfRowsToMove
End Sub
Private Sub grdPrint_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Call gobjDiary.ListUnboundReadData(RowBuf, StartLocation, ReadPriorRows)
End Sub

Private Sub grdViewByWeek_DblClick(Index As Integer)
  Call DiaryEditEvent
End Sub

Private Sub grdViewByWeek_GotFocus(Index As Integer)
  Dim intCount As Integer
  For intCount = 0 To 6
    grdViewByWeek(intCount).SelBookmarks.RemoveAll
  Next
  grdViewByWeek(Index).SelBookmarks.Add grdViewByWeek(Index).GetBookmark(0)
End Sub


Private Sub grdViewByWeek_RowColChange(Index As Integer, ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmDiary.grdViewByWeek_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    
  If Me.grdViewByWeek.Item(Index).SelBookmarks.Count > 1 Then
    'Me.cmdView.Enabled = False
  ElseIf Me.grdViewByWeek.Item(Index).SelBookmarks.Count = 1 Then
    Me.grdViewByWeek.Item(Index).SelBookmarks.RemoveAll
    Me.grdViewByWeek.Item(Index).SelBookmarks.Add Me.grdViewByWeek.Item(Index).Bookmark
    '
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdViewByWeek_RowLoaded(Index As Integer, ByVal Bookmark As Variant)

  With grdViewByWeek(Index).Columns("Alarm")
    
    'MH20010131 Fault 1752
    'Old way was not working for french regional settings
    'If .Value <> 0 Then
    If CStr(.Value) = "True" Then
      .CellStyleSet "Alarm"
      .Text = Space$(10) & "Alarm"
    Else
      .Text = vbNullString
    End If
  End With

End Sub

Private Sub mvwViewbyMonth_DateClick(ByVal DateClicked As Date)
  gobjDiary.DiaryEventID = 0
  gobjDiary.DateSelected = DateClicked
End Sub

Private Sub mvwViewbyMonth_DateDblClick(ByVal DateDblClicked As Date)
  gobjDiary.DiaryEventID = 0
  Call ChangeView(0)
  gobjDiary.RefreshDiaryData
End Sub

Private Sub mvwViewbyMonth_GotFocus()
  'The DoEvents command is required so that the daybold property of
  'the month view is set when the monthview is the startup view
  If gobjDiary.CurrentView = 2 Then
    DoEvents
    gobjDiary.RefreshDiaryData
  End If
End Sub

'Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  Dim intMovement As Integer

  Select Case Tool.Name
  'Case "Day"
  Case "DayView"
    Call ChangeView(0)
    grdViewByDay.SetFocus
  
  'Case "Week"
  Case "WeekView"
    Call ChangeView(1)
  
  'Case "Month"
  Case "MonthView"
    Call ChangeView(2)
    gobjDiary.DiaryEventID = 0
    mvwViewbyMonth.Value = gobjDiary.DateSelected
    mvwViewbyMonth.SetFocus
  
  'Case "List"
  Case "ListView"
    Call ChangeView(3)
    grdViewByList.SetFocus

  Case "Previous"
    gobjDiary.DiaryEventID = 0
    gobjDiary.MoveDate (-1)
    'gobjDiary.RefreshDiaryData
  
  Case "Next"
    gobjDiary.DiaryEventID = 0
    gobjDiary.MoveDate (1)
    'gobjDiary.RefreshDiaryData

  Case "New"
    gobjDiary.DiaryEventID = 0   'New Record
    frmDiaryDetail.Show vbModal
    gobjDiary.RefreshDiaryData
  
  'Case "Duplicate"
  Case "Repeat"
    If EventStillExists Then
      frmDiaryDuplicate.Show vbModal
      gobjDiary.RefreshDiaryData
    End If

  Case "Edit"
    Call DiaryEditEvent
    'NHRD08102004 Fault 7441 commented this
    ' out as it is giving warning message twice and
    ' gobjDiary.RefreshDiaryData is done earlier anyway.

  Case "Delete"
    gobjDiary.DeleteCurrentEntry
    gobjDiary.RefreshDiaryData

  Case "Alarm"
    frmDiaryAlarmSet.Show vbModal
    gobjDiary.RefreshDiaryData

  'Case "Go To"
  Case "Goto"
    frmDiaryGoTo.Show vbModal
    gobjDiary.RefreshDiaryData
  
  Case "Refresh"
    gobjDiary.DiaryEventID = 0
    gobjDiary.RefreshDiaryData
  
  Case "Filter"
    frmDiaryFilter.Show vbModal

  Case "ClearFilter"
    Call ClearFilter
    gobjDiary.RefreshDiaryData

  Case "Print"
    Call PrintDiaryEvents

  Case "Cut"
    gobjDiary.CutEntries blnDelete:=True
    gobjDiary.RefreshDiaryData

  Case "Copy"
    gobjDiary.CutEntries blnDelete:=False
    gobjDiary.RefreshDiaryData

  Case "Paste"
    gobjDiary.PasteEntries
    gobjDiary.RefreshDiaryData

  End Select
  
  'Call EnableToolBar

End Sub


Public Sub ChangeView(intNewViewType As Integer)

  gobjDiary.ChangeView (intNewViewType)
  gobjDiary.RefreshDiaryData
  fraViewType(intNewViewType).ZOrder 0
  mintCurrentView = intNewViewType
  
End Sub


Private Sub DiaryEditEvent()
  
  Dim rsTables As Recordset
  
  If EventStillExists Then
    frmDiaryDetail.Show vbModal
    'Call EnableToolBar
  End If

End Sub


'Private Sub EnableToolBar()
'
'  'This sub should be called after any
'  'other form has been shown modally
'
'  With SSActiveToolBars1
'    .Redraw = False
'    .Enabled = False
'    .Enabled = True
'    .Redraw = True
'  End With
'
'End Sub


Private Sub PrintDiaryEvents()

  ' RH 19/09/00 - BUG 950 - Grids return 'Cancelled by user' COAMsgBox if a printer
  '                         is not connected to the machine, so use the check in
  '                         clsPrintDef to check for printer existance first
  Dim frmDiaryPrint As HRProDataMgr.frmDiaryPrintOptions
  Dim objPrintDef As clsPrintDef
  Dim bCancelled As Boolean
  Dim strErrorMessage As String
  Dim strDateRange As String
  
  Set frmDiaryPrint = New HRProDataMgr.frmDiaryPrintOptions
  Set objPrintDef = New clsPrintDef
  
  'NHRD16072004 Fault 8738
  If objPrintDef.IsOK = False Then Exit Sub
  
  'frmDiaryPrint.EnableDefault = Not (mintCurrentView = mintVIEWBYMONTH)
  strDateRange = vbNullString
  frmDiaryPrint.Initialise
  frmDiaryPrint.Show vbModal
  
  If objPrintDef.IsOK Then
  
    Set objPrintDef = Nothing
    gobjDiary.Printing = True
  
    With frmDiaryPrint
      bCancelled = .Cancelled
      If Not bCancelled Then
        strDateRange = gobjDiary.GetPrintData(.ReturnDateType, .RangeStartDate, .RangeEndDate, .StartDateExpressionID, .EndDateExpressionID)
        strErrorMessage = .NoRecordsMessage
      End If
    End With
    
    If Not bCancelled Then
      With grdPrint
        If .Rows = 0 Then
          COAMsgBox strErrorMessage, vbInformation, "Diary Print"
          
        Else
          'This bit prints using the PrintGrid object
          Set DiaryPrint = New clsPrintGrid
          DiaryPrint.Heading = Me.Caption & strDateRange
          DiaryPrint.Grid = grdPrint
          DiaryPrint.SuppressPrompt = Not gbPrinterPrompt
          DiaryPrint.PrintGrid
          Set DiaryPrint = Nothing
          
          Dim objDefPrinter As cSetDfltPrinter
          Set objDefPrinter = New cSetDfltPrinter
          Do
            objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
          Loop While Printer.DeviceName <> gstrDefaultPrinterName
          Set objDefPrinter = Nothing
          
        End If
      End With
    
      gobjDiary.Printing = False
    End If
  
    gobjDiary.RefreshDiaryData

  End If

End Sub


Private Sub CheckKeyDown(grdTemp As SSDBGrid, KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
  Case 46   'Delete Key
    'If frmDiary.SSActiveToolBars1.Tools(8).Enabled Then   'Check for delete button
    If DiaryToolBar("Delete").Enabled Then
      gobjDiary.DeleteCurrentEntry
      gobjDiary.RefreshDiaryData
    End If
    
  Case 35   'End key
    If Shift = 2 Then   'Ctrl Key
      grdTemp.MoveLast
      gobjDiary.DiaryEventID = CLng(grdTemp.Columns(0).Text)
      gobjDiary.RefreshDiaryData
    End If

  Case 36   'Home key
    If Shift = 2 Then   'Ctrl Key
      grdTemp.MoveFirst
      gobjDiary.DiaryEventID = CLng(grdTemp.Columns(0).Text)
      gobjDiary.RefreshDiaryData
    End If
  
  End Select
  Screen.MousePointer = vbDefault

End Sub


Private Sub ClearFilter()

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
    
    
  strMBText = "Are you sure you want to clear the current filter?"
  intMBButtons = vbInformation + vbYesNo
  strMBTitle = "Clear Filter"
  intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
  If intMBResponse = vbYes Then
    gobjDiary.FilterEventType = 0
    gobjDiary.FilterAlarmStatus = 0
    gobjDiary.FilterPastPresent = 0
    gobjDiary.FilterOnlyMine = False
    frmDiary.Caption = gobjDiary.FilterText
  End If

End Sub


Private Sub grdViewByWeek_Click(Index As Integer)
  SetDate Index
End Sub

Private Sub lblDayTitle_Click(Index As Integer)
  SetDate Index
End Sub

Private Sub SetDate(intIndex As Integer)
  
  'Setting the date can cause an error if its outside the
  'the monthview mindate and maxdate
  On Local Error GoTo ExitSub
  
  gobjDiary.DateSelected = lblDayTitle(intIndex).Tag
  gobjDiary.DiaryEventID = 0
  
  With grdViewByWeek(intIndex)
    If Trim$(.Columns(0).Text) <> vbNullString Then
      gobjDiary.DiaryEventID = CLng(.Columns(0).Text)
    End If
    '.SetFocus
  End With

ExitSub:

End Sub


Private Sub Form_Resize()

  Dim lngCount As Long
  Dim lngLeft As Long
  Dim lngWidth As Long
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  'The monthview determines the size of frmDiary !!!
  With mvwViewbyMonth

    UI.LockWindow Me.hWnd
    
    .Top = 150
    'List View
'    Me.Width = (Me.Width - Me.ScaleWidth) + .Width + 240
'    'Me.Height = (Me.Height - Me.ScaleHeight) + .Height '+ 360
'    Me.Height = (Me.Height - Me.ScaleHeight) + .Height + _
'                (ActiveBar1.Bands(0).Height * Screen.TwipsPerPixelY)

    For lngCount = 0 To 3
      fraViewType(lngCount).Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    Next
    
    'List View
    grdViewByList.Move .Left, .Top, .Width

    'Day View
    lblSelectedDay.Move .Left, .Top, .Width
    grdViewByDay.Move .Left, .Top + (lblSelectedDay.Top + lblSelectedDay.Height), .Width

    'Week View
    lblWeekNo.Move .Left, .Top, .Width
    lblSelectedWeek.Move .Left, lblWeekNo.Top + lblWeekNo.Height + 120, .Width

    lngWidth = (Me.ScaleWidth - ((lblDayTitle(0).Width * 2) + 600)) / 2
    lngLeft = lblDayTitle(0).Width + 240
    For lngCount = 0 To 2
      lblDayTitle(lngCount).Left = 120
      grdViewByWeek(lngCount).Left = lngLeft
      grdViewByWeek(lngCount).Width = lngWidth
    Next

    lngLeft = lngLeft + lngWidth + 120
    For lngCount = 3 To 6
      lblDayTitle(lngCount).Left = Me.ScaleWidth - (lblDayTitle(0).Width + 120)
      grdViewByWeek(lngCount).Left = lngLeft
      grdViewByWeek(lngCount).Width = lngWidth
    Next

    UI.UnlockWindow

  End With

End Sub


Private Function EventStillExists() As Boolean

  Dim rsTables As Recordset
  
  EventStillExists = True
  
  If gobjDiary.DiaryEventID > 0 Then
    Set rsTables = gobjDiary.GetCurrentRecord
      
    If rsTables.BOF And rsTables.EOF Then
      COAMsgBox "This diary event has been deleted by another user.", vbCritical, "Diary"
      gobjDiary.DiaryEventID = 0
      gobjDiary.RefreshDiaryData
      EventStillExists = False
    End If
  End If

End Function

Public Property Get DiaryToolBar() As ActiveBarLibraryCtl.Tools
  Set DiaryToolBar = ActiveBar1.Bands(0).Tools
End Property


Private Sub grdViewByWeek_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    Me.ActiveBar1.Bands("bndDiary").TrackPopup -1, -1
  End If
End Sub

Private Sub grdViewByDay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    Me.ActiveBar1.Bands("bndDiary").TrackPopup -1, -1
  End If
End Sub

Private Sub mvwViewbyMonth_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    Me.ActiveBar1.Bands("bndDiary").TrackPopup -1, -1
  Else
    gobjDiary.DateSelected = mvwViewbyMonth.Value
    gobjDiary.RefreshDiaryData
  End If
End Sub

Private Sub grdViewByList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    Me.ActiveBar1.Bands("bndDiary").TrackPopup -1, -1
  End If
End Sub


Private Sub Form_Load()
  SetColours
End Sub


Private Sub SetColours()

  Dim intIndex As Integer

  SetGridColours grdViewByList
  SetGridColours grdViewByDay

  For intIndex = 0 To 6
    SetGridColours grdViewByWeek(intIndex)
  Next

  With mvwViewbyMonth
    .BackColor = glngDEFAULTDATABACKCOLOUR
    .ForeColor = glngDEFAULTDATAFORECOLOUR
    .MonthBackColor = glngDEFAULTDATABACKCOLOUR
    .TitleBackColor = glngDEFAULTHEADINGFORECOLOUR
    .TitleForeColor = glngDEFAULTHEADINGBACKCOLOUR
    .TrailingForeColor = glngDEFAULTHEADINGFORECOLOUR
  End With

End Sub


Private Sub SetGridColours(grd As SSDBGrid)

  With grd
    .BackColor = glngDEFAULTDATABACKCOLOUR
    .BackColorEven = glngDEFAULTDATABACKCOLOUR
    .BackColorOdd = glngDEFAULTDATABACKCOLOUR
    .ForeColor = glngDEFAULTDATAFORECOLOUR
    .ForeColorEven = glngDEFAULTDATAFORECOLOUR
    .ForeColorOdd = glngDEFAULTDATAFORECOLOUR
    .StyleSets("Alarm").BackColor = glngDEFAULTDATABACKCOLOUR
  End With

End Sub

