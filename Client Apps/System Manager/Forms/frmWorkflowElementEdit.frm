VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmWorkflowElementEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Element"
   ClientHeight    =   9675
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5055
   Icon            =   "frmWorkflowElementEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraElement 
      Height          =   6800
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   550
      Visible         =   0   'False
      Width           =   6100
      Begin VB.TextBox txtEmailCC 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1780
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2300
         Width           =   2580
      End
      Begin VB.CommandButton cmdEmailCC 
         Caption         =   "..."
         Height          =   315
         Left            =   4360
         TabIndex        =   19
         Top             =   2300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.ComboBox cboEmailTable 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":000C
         Left            =   1780
         List            =   "frmWorkflowElementEdit.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   2895
      End
      Begin VB.TextBox txtEmailSubject 
         Height          =   315
         Left            =   1780
         MaxLength       =   200
         TabIndex        =   21
         Top             =   2700
         Width           =   2895
      End
      Begin VB.ComboBox cboEmailElement 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":0010
         Left            =   1780
         List            =   "frmWorkflowElementEdit.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   700
         Width           =   2895
      End
      Begin VB.ComboBox cboEmailRecordSelector 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":0014
         Left            =   1780
         List            =   "frmWorkflowElementEdit.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1100
         Width           =   2895
      End
      Begin VB.ComboBox cboEmailRecord 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":0018
         Left            =   1780
         List            =   "frmWorkflowElementEdit.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "..."
         Height          =   315
         Left            =   4360
         TabIndex        =   16
         Top             =   1900
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1780
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1900
         Width           =   2580
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "&Add ..."
         Height          =   400
         Left            =   4795
         TabIndex        =   27
         Top             =   3200
         Width           =   1200
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "&Edit ..."
         Height          =   400
         Left            =   4795
         TabIndex        =   28
         Top             =   3700
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "&Remove"
         Height          =   400
         Left            =   4795
         TabIndex        =   30
         Top             =   4700
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemoveAllItems 
         Caption         =   "Remo&ve All"
         Height          =   400
         Left            =   4795
         TabIndex        =   31
         Top             =   5200
         Width           =   1200
      End
      Begin VB.CommandButton cmdCopyItem 
         Caption         =   "Cop&y ..."
         Height          =   400
         Left            =   4795
         TabIndex        =   29
         Top             =   4200
         Width           =   1200
      End
      Begin SSDataWidgets_B.SSDBGrid ssgrdItems 
         Height          =   3405
         Left            =   195
         TabIndex        =   26
         Top             =   3195
         Width           =   4480
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         Col.Count       =   10
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
         RowNavigation   =   1
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   10
         Columns(0).Width=   8811
         Columns(0).Caption=   "Items"
         Columns(0).Name =   "Description"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "ItemType"
         Columns(1).Name =   "ItemType"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Caption"
         Columns(2).Name =   "Caption"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "DBColumnID"
         Columns(3).Name =   "DBColumnID"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "DBRecord"
         Columns(4).Name =   "DBRecord"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "WFFormIdentifier"
         Columns(5).Name =   "WFFormIdentifier"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "WFValueIdentifier"
         Columns(6).Name =   "WFValueIdentifier"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "DBWebForm"
         Columns(7).Name =   "DBWebForm"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "DBRecordSelector"
         Columns(8).Name =   "DBRecordSelector"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "CalculationID"
         Columns(9).Name =   "CalculationID"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   7902
         _ExtentY        =   6006
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdAttachClear 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4360
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2900
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdAttachAttachment 
         Caption         =   "..."
         Height          =   315
         Left            =   4045
         TabIndex        =   24
         Top             =   2900
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtAttachAttachment 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   1780
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2900
         Visible         =   0   'False
         Width           =   2270
      End
      Begin VB.CommandButton cmdMoveItemDown 
         Caption         =   "Do&wn"
         Height          =   405
         Left            =   4795
         TabIndex        =   33
         Top             =   6195
         Width           =   1200
      End
      Begin VB.CommandButton cmdMoveItemUp 
         Caption         =   "&Up"
         Height          =   405
         Left            =   4795
         TabIndex        =   32
         Top             =   5700
         Width           =   1200
      End
      Begin VB.Label lblEmailCC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Copy :"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   2355
         Width           =   885
      End
      Begin VB.Label lblEmailTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   200
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblEmailSubject 
         AutoSize        =   -1  'True
         Caption         =   "Email Subject :"
         Height          =   195
         Left            =   200
         TabIndex        =   20
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label lblEmailRecordSelector 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Selector :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   1155
         Width           =   1500
      End
      Begin VB.Label lblEmailWebForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Element :"
         Height          =   195
         Left            =   200
         TabIndex        =   8
         Top             =   765
         Width           =   675
      End
      Begin VB.Label lblEmailRecord 
         Caption         =   "Email Record :"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email To :"
         Height          =   195
         Left            =   200
         TabIndex        =   14
         Top             =   1965
         Width           =   690
      End
      Begin VB.Label lblAttachAttachment 
         AutoSize        =   -1  'True
         Caption         =   "Attachment :"
         Height          =   195
         Left            =   200
         TabIndex        =   22
         Top             =   2960
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.Frame fraElement 
      Height          =   6400
      Index           =   2
      Left            =   6240
      TabIndex        =   43
      Top             =   550
      Visible         =   0   'False
      Width           =   7800
      Begin VB.Frame fraDataRecord 
         Caption         =   "Secondary Record :"
         Height          =   1900
         Index           =   1
         Left            =   3950
         TabIndex        =   57
         Top             =   750
         Width           =   3650
         Begin VB.ComboBox cboDataRecordTable 
            Height          =   315
            Index           =   1
            ItemData        =   "frmWorkflowElementEdit.frx":001C
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":001E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1450
            Width           =   1815
         End
         Begin VB.ComboBox cboDataRecord 
            Height          =   315
            Index           =   1
            ItemData        =   "frmWorkflowElementEdit.frx":0020
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   250
            Width           =   1815
         End
         Begin VB.ComboBox cboDataElement 
            Height          =   315
            Index           =   1
            ItemData        =   "frmWorkflowElementEdit.frx":0024
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":0026
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cboDataRecordSelector 
            Height          =   315
            Index           =   1
            ItemData        =   "frmWorkflowElementEdit.frx":0028
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":002A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label lblDataRecordTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   64
            Top             =   1510
            Width           =   495
         End
         Begin VB.Label lblDataRecord 
            Caption         =   "Record :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   58
            Top             =   315
            Width           =   795
         End
         Begin VB.Label lblDataRecordSelector 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Record Selector :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   62
            Top             =   1110
            Width           =   1470
         End
         Begin VB.Label lblDataWebForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Element :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   60
            Top             =   705
            Width           =   675
         End
      End
      Begin VB.Frame fraDataRecord 
         Caption         =   "Primary Record :"
         Height          =   1900
         Index           =   0
         Left            =   200
         TabIndex        =   48
         Top             =   750
         Width           =   3650
         Begin VB.ComboBox cboDataRecordTable 
            Height          =   315
            Index           =   0
            ItemData        =   "frmWorkflowElementEdit.frx":002C
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":002E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1450
            Width           =   1815
         End
         Begin VB.ComboBox cboDataRecord 
            Height          =   315
            Index           =   0
            ItemData        =   "frmWorkflowElementEdit.frx":0030
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":0032
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   250
            Width           =   1815
         End
         Begin VB.ComboBox cboDataElement 
            Height          =   315
            Index           =   0
            ItemData        =   "frmWorkflowElementEdit.frx":0034
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":0036
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cboDataRecordSelector 
            Height          =   315
            Index           =   0
            ItemData        =   "frmWorkflowElementEdit.frx":0038
            Left            =   1730
            List            =   "frmWorkflowElementEdit.frx":003A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label lblDataRecordTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   55
            Top             =   1515
            Width           =   495
         End
         Begin VB.Label lblDataRecord 
            Caption         =   "Record :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   49
            Top             =   315
            Width           =   885
         End
         Begin VB.Label lblDataRecordSelector 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Record Selector :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   53
            Top             =   1110
            Width           =   1470
         End
         Begin VB.Label lblDataWebForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Element :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   51
            Top             =   705
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdDataRemoveAll 
         Caption         =   "Remo&ve All"
         Height          =   400
         Left            =   6400
         TabIndex        =   70
         Top             =   4300
         Width           =   1200
      End
      Begin VB.CommandButton cmdDataRemove 
         Caption         =   "&Remove"
         Height          =   400
         Left            =   6400
         TabIndex        =   69
         Top             =   3800
         Width           =   1200
      End
      Begin VB.CommandButton cmdDataEdit 
         Caption         =   "&Edit ..."
         Height          =   400
         Left            =   6400
         TabIndex        =   68
         Top             =   3300
         Width           =   1200
      End
      Begin VB.CommandButton cmdDataAdd 
         Caption         =   "&Add ..."
         Height          =   400
         Left            =   6400
         TabIndex        =   67
         Top             =   2800
         Width           =   1200
      End
      Begin VB.ComboBox cboDataAction 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":003C
         Left            =   1050
         List            =   "frmWorkflowElementEdit.frx":0049
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   300
         Width           =   2750
      End
      Begin VB.ComboBox cboDataTable 
         Height          =   315
         Left            =   4750
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   300
         Width           =   2750
      End
      Begin SSDataWidgets_B.SSDBGrid ssgrdColumns 
         Height          =   3400
         Left            =   200
         TabIndex        =   66
         Top             =   2800
         Width           =   6000
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         Col.Count       =   10
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
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   10
         Columns(0).Width=   4180
         Columns(0).Caption=   "Column"
         Columns(0).Name =   "ColumnDesc"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   6165
         Columns(1).Caption=   "Value"
         Columns(1).Name =   "ValueDesc"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "ColumnID"
         Columns(2).Name =   "ColumnID"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "ValueType"
         Columns(3).Name =   "ValueType"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Value"
         Columns(4).Name =   "Value"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "WFFormIdentifier"
         Columns(5).Name =   "WFFormIdentifier"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "WFValueIdentifier"
         Columns(6).Name =   "WFValueIdentifier"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "DBColumnID"
         Columns(7).Name =   "DBColumnID"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "DBRecord"
         Columns(8).Name =   "DBRecord"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "CalculationID"
         Columns(9).Name =   "CalculationID"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10583
         _ExtentY        =   5997
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDataAction 
         Caption         =   "Action :"
         Height          =   195
         Left            =   195
         TabIndex        =   44
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblDataTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   4050
         TabIndex        =   46
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame fraElement 
      Height          =   2100
      Index           =   1
      Left            =   6360
      TabIndex        =   34
      Top             =   7440
      Visible         =   0   'False
      Width           =   5130
      Begin VB.Frame fraDecisionFlow 
         Caption         =   "'True' flow criteria :"
         Height          =   1095
         Left            =   200
         TabIndex        =   37
         Top             =   750
         Width           =   4740
         Begin VB.TextBox txtDecisionFlowExpression 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   650
            Width           =   2655
         End
         Begin VB.CommandButton cmdDecisionFlowExpression 
            Caption         =   "..."
            Height          =   315
            Left            =   4305
            TabIndex        =   42
            Top             =   650
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.ComboBox cboDecisionFlowButton 
            Height          =   315
            Left            =   1650
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   250
            Width           =   2970
         End
         Begin VB.OptionButton optDecisionFlowType 
            Caption         =   "Ca&lculation"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   39
            Top             =   710
            Width           =   1320
         End
         Begin VB.OptionButton optDecisionFlowType 
            Caption         =   "&Button"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   38
            Top             =   310
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboDecisionCaption 
         Height          =   315
         ItemData        =   "frmWorkflowElementEdit.frx":0065
         Left            =   1860
         List            =   "frmWorkflowElementEdit.frx":0075
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   300
         Width           =   2970
      End
      Begin VB.Label lblDecisionCaption 
         AutoSize        =   -1  'True
         Caption         =   "Decision Caption :"
         Height          =   195
         Left            =   195
         TabIndex        =   35
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame fraButtons 
      Height          =   400
      Left            =   6360
      TabIndex        =   71
      Top             =   6960
      Width           =   2600
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   73
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraCaption 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   10785
      Begin VB.TextBox txtIdentifier 
         Height          =   315
         Left            =   5535
         MaxLength       =   200
         TabIndex        =   4
         Top             =   0
         Width           =   2750
      End
      Begin VB.TextBox txtCaption 
         Height          =   315
         Left            =   1300
         MaxLength       =   200
         TabIndex        =   2
         Top             =   0
         Width           =   2900
      End
      Begin VB.Label lblIdentifier 
         AutoSize        =   -1  'True
         Caption         =   "Identifier :"
         Height          =   195
         Left            =   4545
         TabIndex        =   3
         Top             =   60
         Width           =   990
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Caption :"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   60
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmWorkflowElementEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfLoading As Boolean
Private mfReadOnly As Boolean

Private mwfElement As VB.Control
Private mfCanBeEdited As Boolean

Private msngMinFormWidth As Single
Private msngMinFormHeight As Single

Private Const iXGAP = 200
Private Const iYGAP = 200
Private Const iXFRAMEGAP = 150
Private Const iYFRAMEGAP = 100
Private Const iCOMPONENTFRAMEWIDTH = 2000
Private Const iFRAMEWIDTH = 5400
Private Const iFRAMEHEIGHT = 3900

Private Enum MoveDirection
  MOVEDIRECTION_UP = 0
  MOVEDIRECTION_DOWN = 1
End Enum

Private mfrmCallingForm As Form
Private mlngPersonnelTableID As Long
Private mlngBaseTableID As Long
Private miInitiationType As WorkflowInitiationTypes

Private Enum EmailRecipients
  EMAIL_TO = 0
  EMAIL_CC = 1
End Enum


Private mlngEmailID As Long
Private mlngEmailCCID As Long
Private mlngDecisionFlowExpressionID As Long

Private Enum StoredDataTableIndex
  STOREDDATATABLE_PRIMARY = 0
  STOREDDATATABLE_SECONDARY = 1
End Enum

Private maWFPrecedingElements() As VB.Control
Private maWFAllElements() As VB.Control
Private miDataRecordBeingRefreshed As Integer
Private mavIdentifierLog() As Variant

Private mfInitializing As Boolean
Private msInitializeMessage As String

Private miEmailAttachmentType As WorkflowEmailAttachmentTypes
Private msEmailAttachment_File As String
Private msEmailAttachment_WFElementIdentifier As String
Private msEmailAttachment_WFItemIdentifier As String
Private mlngEmailAttachment_DBColumnID As Long
Private miEmailAttachment_DBRecord As Integer
Private msEmailAttachment_DBElementIdentifier As String
Private msEmailAttachment_DBItemIdentifier As String
Private Function GetElementByIdentifier(psIdentifier As String) As VB.Control
  ' Return the element with the given identifier.
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  
  If Len(Trim(psIdentifier)) = 0 Then
    Exit Function
  End If
  
  For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfTemp = maWFPrecedingElements(lngLoop)

    If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(psIdentifier))) Then
      Set GetElementByIdentifier = wfTemp
      Exit For
    End If
    
    Set wfTemp = Nothing
  Next lngLoop
  
End Function


Public Property Set CallingForm(pfrmForm As Form)
  Set mfrmCallingForm = pfrmForm
  mfReadOnly = pfrmForm.ReadOnly
  mlngBaseTableID = pfrmForm.BaseTable
  miInitiationType = pfrmForm.InitiationType

End Property

Public Property Get CallingForm() As Form
  Set CallingForm = mfrmCallingForm
End Property


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Public Property Get CanBeEdited() As Boolean
  CanBeEdited = mfCanBeEdited
End Property

Private Sub FormatScreen()
  ' Format the screen depending on the type of element being edited.
  Dim fraTemp As Frame
  Dim iFrameIndex As Integer
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sRow As String
  Dim fFound As Boolean
  Dim sDescription As String
  
  iFrameIndex = -1
  mfCanBeEdited = True
  mfLoading = True
  
  mfInitializing = True
  msInitializeMessage = ""
  
  txtCaption.MaxLength = 200
  txtCaption.Text = mwfElement.Caption
  
  ' Log the original identifiers of the controls
  ' Column 1 = the control
  ' Column 2 = original identifier
  ' Column 3 = new identifier (populated in cmdOK_Click)
  ' Column 4 = deleted flag (only used in WebForms)
  ' Column 5 = original StoredData table
  ' Column 6 = current StoredData table (defaulted to original value, updated in cmdOK_Click)
  ' NB. Row 0 is for the form itself.
  ReDim mavIdentifierLog(6, 0)
  mavIdentifierLog(2, 0) = mwfElement.Identifier
  mavIdentifierLog(3, 0) = ""
  mavIdentifierLog(4, 0) = False
  
  If mwfElement.ElementType = elem_StoredData Then
    mavIdentifierLog(5, 0) = mwfElement.DataTableID
    mavIdentifierLog(6, 0) = mwfElement.DataTableID
  End If
  
  txtIdentifier.Text = mwfElement.Identifier
  
  Select Case mwfElement.ElementType
    Case elem_Begin
      
    Case elem_Connector1, elem_Connector2
      txtCaption.MaxLength = 1
      
    Case elem_Decision
      iFrameIndex = 1
            
      For iLoop = 0 To cboDecisionCaption.ListCount - 1
        If cboDecisionCaption.ItemData(iLoop) = mwfElement.DecisionCaptionType Then
          cboDecisionCaption.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      mlngDecisionFlowExpressionID = mwfElement.DecisionFlowExpressionID
      optDecisionFlowType(mwfElement.DecisionFlowType).value = True
      DecisionFlowControls_Refresh

    Case elem_Email
      iFrameIndex = 0
      
      ' Get the Email Address details.
      lblEmail.Visible = True
      txtEmail.Visible = True
      cmdEmail.Visible = True
      
      lblEmailCC.Visible = True
      txtEmailCC.Visible = True
      cmdEmailCC.Visible = True
      
      lblEmailRecord.Visible = True
      cboEmailRecord.Visible = True
      
      mlngEmailID = mwfElement.EmailID
      mlngEmailCCID = mwfElement.EmailCCID
      
      cboEmailRecord_Refresh
      
      txtEmailSubject.Text = mwfElement.EMailSubject
      
      miEmailAttachmentType = mwfElement.Attachment_Type
      msEmailAttachment_File = mwfElement.Attachment_File
      msEmailAttachment_WFElementIdentifier = mwfElement.Attachment_WFElementIdentifier
      msEmailAttachment_WFItemIdentifier = mwfElement.Attachment_WFValueIdentifier
      mlngEmailAttachment_DBColumnID = mwfElement.Attachment_DBColumnID
      miEmailAttachment_DBRecord = mwfElement.Attachment_DBRecord
      msEmailAttachment_DBElementIdentifier = mwfElement.Attachment_DBElement
      msEmailAttachment_DBItemIdentifier = mwfElement.Attachment_DBValue
      
      sDescription = ""
      
      Select Case miEmailAttachmentType
        Case giWFEMAILITEM_DBVALUE
          sDescription = "Database value - " & GetColumnName(mlngEmailAttachment_DBColumnID)
        Case giWFEMAILITEM_WFVALUE
          sDescription = "Workflow value - " & msEmailAttachment_WFElementIdentifier & "." & msEmailAttachment_WFItemIdentifier
        Case giWFEMAILITEM_FILEATTACHMENT
          sDescription = "File - '" & msEmailAttachment_File & "'"
      End Select
      txtAttachAttachment.Text = sDescription

      cmdRemoveAllItems.Top = cmdRemoveAllItems.Top - (cmdAddItem.Top - ssgrdItems.Top)
      cmdRemoveItem.Top = cmdRemoveItem.Top - (cmdAddItem.Top - ssgrdItems.Top)
      cmdCopyItem.Top = cmdCopyItem.Top - (cmdAddItem.Top - ssgrdItems.Top)
      cmdEditItem.Top = cmdEditItem.Top - (cmdAddItem.Top - ssgrdItems.Top)
      cmdAddItem.Top = ssgrdItems.Top
      
      asItems = mwfElement.Items
      
      With ssgrdItems
        .RemoveAll
        
        For iLoop2 = 1 To UBound(asItems, 2)
          sRow = ""
          
          sRow = sRow & asItems(1, iLoop2) & vbTab ' Description
          sRow = sRow & asItems(2, iLoop2) & vbTab ' Item Type
          sRow = sRow & asItems(3, iLoop2) & vbTab ' Caption
          sRow = sRow & asItems(4, iLoop2) & vbTab ' DB Column ID
          sRow = sRow & asItems(5, iLoop2) & vbTab ' DB Record
          sRow = sRow & asItems(11, iLoop2) & vbTab ' WF Form Identifier
          sRow = sRow & asItems(12, iLoop2) & vbTab ' WF Value Identifier
          sRow = sRow & asItems(13, iLoop2) & vbTab ' DB Web Form Identifier
          sRow = sRow & asItems(14, iLoop2) & vbTab ' DB RecSel Identifier
          sRow = sRow & asItems(56, iLoop2) & vbTab ' Calculation ID
          
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Update
  
          .SelBookmarks.RemoveAll
        Next iLoop2
        
        If .Rows > 0 Then
          .Bookmark = .AddItemBookmark(0)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End If
      End With
      RefreshExpressionNames
      ResizeGridColumns ssgrdItems
      
    Case elem_Or
      mfCanBeEdited = False
    
    Case elem_StoredData
      iFrameIndex = 2

      miDataRecordBeingRefreshed = -1

      fFound = False
      For iLoop = 0 To cboDataAction.ListCount - 1
        If cboDataAction.ItemData(iLoop) = mwfElement.DataAction Then
          cboDataAction.ListIndex = iLoop
          fFound = True
          Exit For
        End If
      Next iLoop
      
      If Not fFound Then
        cboDataAction.ListIndex = 0
      End If
      
      cboDataTable_Refresh
      
      cboDataRecord_Refresh STOREDDATATABLE_PRIMARY
      
      avColumns = mwfElement.DataColumns
      
      With ssgrdColumns
        .RemoveAll
        
        For iLoop2 = 1 To UBound(avColumns, 2)
          sRow = ""
          
          For iLoop = 1 To UBound(avColumns, 1)
            sRow = sRow & avColumns(iLoop, iLoop2) & vbTab
          Next iLoop
                    
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Update
  
          .SelBookmarks.RemoveAll
        Next iLoop2
        
        If .Rows > 0 Then
          .Bookmark = .AddItemBookmark(0)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End If
      End With
      ResizeGridColumns ssgrdColumns

    Case elem_SummingJunction
      mfCanBeEdited = False
    
    Case elem_Terminator
    
    Case elem_WebForm
      ' Now uses workflow web form designer.
      
  End Select
  
  If Not mfCanBeEdited Then
    Exit Sub
  End If
  
  For Each fraTemp In fraElement
    If fraTemp.Index = iFrameIndex Then
      With fraTemp
        .Left = iXFRAMEGAP
        .Top = fraCaption.Top + fraCaption.Height + iYFRAMEGAP
        
        fraButtons.Top = .Top + .Height + iYFRAMEGAP
        fraButtons.Left = .Left + .Width - fraButtons.Width
        
        msngMinFormWidth = .Left + _
          .Width + iXFRAMEGAP + _
          (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
        msngMinFormHeight = fraButtons.Top + fraButtons.Height + iXFRAMEGAP + _
          (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
        
        Me.Width = msngMinFormWidth
        Me.Height = msngMinFormHeight
        
        .Visible = True
      End With
    End If
    
    fraTemp.Visible = (fraTemp.Index = iFrameIndex)
  Next fraTemp
  Set fraTemp = Nothing
  
  If iFrameIndex < 0 Then
    With fraCaption
      fraButtons.Top = .Top + .Height + iYFRAMEGAP
      fraButtons.Left = .Left + txtCaption.Left + txtCaption.Width - fraButtons.Width
      
      Me.Width = .Left + _
        .Width + iXFRAMEGAP + _
        (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
      Me.Height = fraButtons.Top + fraButtons.Height + iXFRAMEGAP + _
        (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
    End With
  End If
  
  mfLoading = False
  
  Changed = (Len(msInitializeMessage) > 0)
  
  If Len(msInitializeMessage) > 0 Then
    MsgBox msInitializeMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  mfInitializing = False
  
  RefreshScreen
  
End Sub

Private Sub cboDataTable_Refresh()
  ' Populate the Data Table combo and
  ' select the current table if it is still valid.
  Dim fTableOK As Boolean
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = -1
  
  ' Clear the current contents of the combo.
  cboDataTable.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combo for each table that has not been deleted.
    Do While Not .EOF
      fTableOK = False

      If (Not .Fields("deleted")) Then
      
        cboDataTable.AddItem !TableName
        cboDataTable.ItemData(cboDataTable.NewIndex) = !TableID

        If !TableID = mwfElement.DataTableID Then
          iIndex = cboDataTable.NewIndex
        End If

        If ((miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) And (!TableID = mlngPersonnelTableID)) _
          Or ((miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) And (!TableID = mlngBaseTableID)) Then
          
          iDefaultIndex = cboDataTable.NewIndex
        End If
      End If

      .MoveNext
    Loop
  End With

  ' Enable the combo if there are items.
  With cboDataTable
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        If iDefaultIndex >= 0 Then
          iIndex = iDefaultIndex
        Else
          iIndex = 0
        End If
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no tables>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
    End If
  End With
    
End Sub

Private Sub DecisionFlowControls_Refresh()
  ' Populate the TrueFlow combo and
  ' select the current value if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sMsg As String
  
  iIndex = -1

  ' Clear the current contents of the combo.
  cboDecisionFlowButton.Clear

  If optDecisionFlowType(decisionFlowType_Button).value Then
    mlngDecisionFlowExpressionID = 0
  
    If (UBound(maWFPrecedingElements) > 1) Then
      
      ' Add  an item to the combo for each button in the preceding web form.
      iLoop = 2
      Set wfTemp = maWFPrecedingElements(iLoop)
      
      ' Ignore connectors
      ' and Decision elements (see fault 11334) to determine the preceding WebForm
      Do While (wfTemp.ElementType = elem_Connector2) _
        Or (wfTemp.ElementType = elem_Decision)
        
        iLoop = iLoop + 1
        
        If (wfTemp.ElementType = elem_Connector2) Then
          iLoop = iLoop + 1
        End If
        
        Set wfTemp = maWFPrecedingElements(iLoop)
      Loop
      
      If wfTemp.ElementType = elem_WebForm Then
        
        asItems = wfTemp.Items
    
        For iLoop = 1 To UBound(asItems, 2)
          If asItems(2, iLoop) = giWFFORMITEM_BUTTON Then
            cboDecisionFlowButton.AddItem asItems(9, iLoop)
            cboDecisionFlowButton.ItemData(cboDecisionFlowButton.NewIndex) = 1
          End If
        Next iLoop
    
        For iLoop = 0 To cboDecisionFlowButton.ListCount - 1
          If cboDecisionFlowButton.List(iLoop) = mwfElement.TrueFlowIdentifier Then
            iIndex = iLoop
          End If
        Next iLoop
      End If
    End If
  
    ' Enable the combo if there are items.
    If (iIndex < 0) Then
      If (Len(Trim(mwfElement.TrueFlowIdentifier)) > 0) Then
        sMsg = "The previously selected '" & GetDecisionCaptionDescription(cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex), True) & "' Flow Button is no longer valid." & vbCrLf
      End If
      
      If cboDecisionFlowButton.ListCount > 0 Then
        sMsg = sMsg & "A default '" & GetDecisionCaptionDescription(cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex), True) & "' Flow Button has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
      
      mfChanged = True
    End If
    
    With cboDecisionFlowButton
      If .ListCount > 0 Then
        .Enabled = True
        If iIndex < 0 Then
          iIndex = 0
          mfChanged = True
        End If
        .ListIndex = iIndex
      Else
        .Enabled = False
  
        .AddItem "<no buttons>"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
      End If
    End With
  End If
  
  txtDecisionFlowExpression.Text = GetExpressionName(mlngDecisionFlowExpressionID)
  
End Sub

Public Property Set Element(pwfElement As VB.Control)
  Set mwfElement = pwfElement
  
  ReDim maWFPrecedingElements(1)
  Set maWFPrecedingElements(1) = mwfElement
  mfrmCallingForm.PrecedingElements mwfElement, maWFPrecedingElements

  ReDim maWFAllElements(0)
  mfrmCallingForm.AllElements maWFAllElements

  FormatScreen

End Property

Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fEnabled As Boolean
  Dim fSecondaryRecordRequired As Boolean
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim optTemp As OptionButton
  
  fEnabled = (Not mfReadOnly)
  txtCaption.Enabled = fEnabled
  txtCaption.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
  lblCaption.Enabled = fEnabled
  
  fEnabled = (Not mfReadOnly)
  lblIdentifier.Visible = (mwfElement.ElementType = elem_StoredData)
  txtIdentifier.Visible = (mwfElement.ElementType = elem_StoredData)
  txtIdentifier.Enabled = fEnabled
  txtIdentifier.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
  lblIdentifier.Enabled = fEnabled
  
  Select Case mwfElement.ElementType
    Case elem_Begin
    
    Case elem_Connector1, elem_Connector2
      
    Case elem_Decision
      fEnabled = (Not mfReadOnly)
      cboDecisionCaption.Enabled = fEnabled
      cboDecisionCaption.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDecisionCaption.Enabled = fEnabled
      fraDecisionFlow.ForeColor = IIf(fEnabled, vbWindowText, vbApplicationWorkspace)

      For Each optTemp In optDecisionFlowType
        optTemp.Enabled = fEnabled
      Next optTemp
      Set optTemp = Nothing
      
      fEnabled = (Not mfReadOnly) And (optDecisionFlowType(decisionFlowType_Button).value)
      cboDecisionFlowButton.Enabled = fEnabled
      cboDecisionFlowButton.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      
      fEnabled = (optDecisionFlowType(decisionFlowType_Expression).value)
      cmdDecisionFlowExpression.Enabled = fEnabled

    Case elem_Email
      fEnabled = (Not mfReadOnly)
      cmdEmail.Enabled = True
      lblEmail.Enabled = fEnabled
      
      cmdEmailCC.Enabled = True
      lblEmailCC.Enabled = fEnabled
      
      fEnabled = (Not mfReadOnly)
      cboEmailRecord.Enabled = fEnabled
      cboEmailRecord.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblEmailRecord.Enabled = fEnabled

      fEnabled = (Not mfReadOnly) _
        And (cboEmailRecord.ItemData(cboEmailRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
        And (cboEmailElement.ListCount > 0)
      cboEmailElement.Enabled = fEnabled
      cboEmailElement.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblEmailWebForm.Enabled = fEnabled
      
      fEnabled = (Not mfReadOnly) _
        And (cboEmailRecord.ItemData(cboEmailRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
        And (cboEmailRecordSelector.ListCount > 0)
      cboEmailRecordSelector.Enabled = fEnabled
      cboEmailRecordSelector.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblEmailRecordSelector.Enabled = fEnabled
      
      fEnabled = (Not mfReadOnly) _
        And (cboEmailTable.ListCount > 0)
      cboEmailTable.Enabled = fEnabled
      cboEmailTable.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblEmailTable.Enabled = fEnabled
      
      cmdAttachClear.Enabled = (Not mfReadOnly) _
        And (Len(txtAttachAttachment.Text) > 0)
      
      fEnabled = (Not mfReadOnly)
      txtEmailSubject.Enabled = fEnabled
      txtEmailSubject.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblEmailSubject.Enabled = fEnabled
      
      cmdAddItem.Enabled = (Not mfReadOnly)

      With ssgrdItems
        If .Rows = 0 Then
          cmdEditItem.Enabled = False
          cmdRemoveItem.Enabled = False
          cmdRemoveAllItems.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditItem.Enabled = (Not mfReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveItem.Enabled = Not mfReadOnly
          Else
            cmdEditItem.Enabled = False
            cmdRemoveItem.Enabled = False
          End If

          cmdRemoveAllItems.Enabled = Not mfReadOnly
        End If

        cmdCopyItem.Enabled = cmdEditItem.Enabled

        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveItemUp.Enabled = False
            cmdMoveItemDown.Enabled = (.Rows > 1) And (Not mfReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveItemUp.Enabled = (.Rows > 1) And (Not mfReadOnly)
            cmdMoveItemDown.Enabled = False
          Else
            cmdMoveItemUp.Enabled = (.Rows > 1) And (Not mfReadOnly)
            cmdMoveItemDown.Enabled = (.Rows > 1) And (Not mfReadOnly)
          End If
        Else
          cmdMoveItemUp.Enabled = False
          cmdMoveItemDown.Enabled = False
        End If
      End With

    Case elem_Or
      ' Not required.
      
    Case elem_StoredData
      fSecondaryRecordRequired = SecondaryRecordRequired
      
      fEnabled = (Not mfReadOnly)
      cboDataAction.Enabled = fEnabled
      cboDataAction.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataAction.Enabled = fEnabled

      fEnabled = (Not mfReadOnly)
      cboDataTable.Enabled = fEnabled
      cboDataTable.BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataTable.Enabled = fEnabled

      fEnabled = (Not mfReadOnly) And (cboDataRecord(STOREDDATATABLE_PRIMARY).ListCount > 0)
      If fEnabled And (cboDataRecord(STOREDDATATABLE_PRIMARY).ListCount = 1) Then
        fEnabled = (cboDataRecord(STOREDDATATABLE_PRIMARY).ItemData(cboDataRecord(STOREDDATATABLE_PRIMARY).ListIndex) <> giWFRECSEL_UNKNOWN)
      End If
      
      fraDataRecord(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      
      cboDataRecord(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      cboDataRecord(STOREDDATATABLE_PRIMARY).BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataRecord(STOREDDATATABLE_PRIMARY).Enabled = fEnabled

      fEnabled = (Not mfReadOnly) And (cboDataElement(STOREDDATATABLE_PRIMARY).ListCount > 0)
      cboDataElement(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      cboDataElement(STOREDDATATABLE_PRIMARY).BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataWebForm(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      
      fEnabled = (Not mfReadOnly) And (cboDataRecordSelector(STOREDDATATABLE_PRIMARY).ListCount > 0)
      cboDataRecordSelector(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      cboDataRecordSelector(STOREDDATATABLE_PRIMARY).BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataRecordSelector(STOREDDATATABLE_PRIMARY).Enabled = fEnabled

      fEnabled = (Not mfReadOnly) And (cboDataRecordTable(STOREDDATATABLE_PRIMARY).ListCount > 0)
      cboDataRecordTable(STOREDDATATABLE_PRIMARY).Enabled = fEnabled
      cboDataRecordTable(STOREDDATATABLE_PRIMARY).BackColor = IIf(fEnabled, vbWindowBackground, vbButtonFace)
      lblDataRecordTable(STOREDDATATABLE_PRIMARY).Enabled = fEnabled

      fEnabled = (Not mfReadOnly) _
        And (cboDataRecord(STOREDDATATABLE_SECONDARY).ListCount > 0) _
        And fSecondaryRecordRequired
      If fEnabled And (cboDataRecord(STOREDDATATABLE_SECONDARY).ListCount = 1) Then
        fEnabled = (cboDataRecord(STOREDDATATABLE_SECONDARY).ItemData(cboDataRecord(STOREDDATATABLE_SECONDARY).ListIndex) <> giWFRECSEL_UNKNOWN)
      End If
      
      fraDataRecord(STOREDDATATABLE_SECONDARY).Enabled = fEnabled
      
      cboDataRecord(STOREDDATATABLE_SECONDARY).Enabled = fEnabled
      cboDataRecord(STOREDDATATABLE_SECONDARY).BackColor = IIf((fEnabled), vbWindowBackground, vbButtonFace)
      lblDataRecord(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled)

      cboDataElement(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataElement(STOREDDATATABLE_SECONDARY).ListCount > 0)
      cboDataElement(STOREDDATATABLE_SECONDARY).BackColor = IIf((fEnabled) And (cboDataElement(STOREDDATATABLE_SECONDARY).ListCount > 0), vbWindowBackground, vbButtonFace)
      lblDataWebForm(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataElement(STOREDDATATABLE_SECONDARY).ListCount > 0)
      
      cboDataRecordSelector(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataRecordSelector(STOREDDATATABLE_SECONDARY).ListCount > 0)
      cboDataRecordSelector(STOREDDATATABLE_SECONDARY).BackColor = IIf((fEnabled) And (cboDataRecordSelector(STOREDDATATABLE_SECONDARY).ListCount > 0), vbWindowBackground, vbButtonFace)
      lblDataRecordSelector(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataRecordSelector(STOREDDATATABLE_SECONDARY).ListCount > 0)

      cboDataRecordTable(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataRecordTable(STOREDDATATABLE_SECONDARY).ListCount > 0)
      cboDataRecordTable(STOREDDATATABLE_SECONDARY).BackColor = IIf((fEnabled) And (cboDataRecordTable(STOREDDATATABLE_SECONDARY).ListCount > 0), vbWindowBackground, vbButtonFace)
      lblDataRecordTable(STOREDDATATABLE_SECONDARY).Enabled = (fEnabled) And (cboDataRecordTable(STOREDDATATABLE_SECONDARY).ListCount > 0)

      If cboDataAction.ItemData(cboDataAction.ListIndex) = DATAACTION_DELETE Then
        cmdDataAdd.Enabled = False
        cmdDataEdit.Enabled = False
        cmdDataRemove.Enabled = False
        cmdDataRemoveAll.Enabled = False
        
        With ssgrdColumns
          .RemoveAll
          .Enabled = False
        End With
      
        ResizeGridColumns ssgrdColumns
      Else
        cmdDataAdd.Enabled = (Not mfReadOnly)
  
        With ssgrdColumns
          .Enabled = True
          
          If .Rows = 0 Then
            cmdDataEdit.Enabled = False
            cmdDataRemove.Enabled = False
            cmdDataRemoveAll.Enabled = False
          Else
            If .SelBookmarks.Count > 0 Then
              cmdDataEdit.Enabled = (Not mfReadOnly) And _
                (.SelBookmarks.Count = 1)
              cmdDataRemove.Enabled = Not mfReadOnly
            Else
              cmdDataEdit.Enabled = False
              cmdDataRemove.Enabled = False
            End If
  
            cmdDataRemoveAll.Enabled = Not mfReadOnly
          End If
        End With
      End If
      
    Case elem_SummingJunction
      ' Not required.
      
    Case elem_Terminator
    
    Case elem_WebForm
      ' Now uses workflow web form designer.
  
  End Select
  
  cmdOK.Enabled = mfChanged

End Sub

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
      lngTextWidth = TextWidth(.Columns(iColumn).Caption)

      If .Columns(iColumn).Visible Then
        iLastVisibleColumn = iColumn
        
        For iRow = 0 To .Rows - 1 Step 1
          varBookmark = .AddItemBookmark(iRow)

          If TextWidth(Trim(.Columns(iColumn).CellText(varBookmark))) > lngTextWidth Then
            lngTextWidth = TextWidth(Trim(.Columns(iColumn).CellText(varBookmark)))
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

Private Function SecondaryRecordRequired() As Boolean
  ' Return TRUE if the secondary record parameters are required.
  ' ie. if the dataAction is INSERTing into a shared history table.
  Dim iAction As DataAction
  Dim fRequired As Boolean
  Dim lngTableID As Long
  Dim rsInfo As DAO.Recordset
  Dim sSQL As String
  Dim iParentCount As Integer
  
  fRequired = False
  If cboDataAction.ListIndex >= 0 And cboDataTable.ListIndex >= 0 Then
    iAction = cboDataAction.ItemData(cboDataAction.ListIndex)
    lngTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
  
    If iAction = DATAACTION_INSERT Then
      ' We are INSERTing. Are we INSERTing into a shared history?
      sSQL = "SELECT COUNT(*) AS recCount" & _
        " FROM tmpRelations" & _
        " WHERE childID = " & CStr(lngTableID)
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      iParentCount = rsInfo!reccount
      rsInfo.Close
      Set rsInfo = Nothing
  
      fRequired = (iParentCount = 2)
    End If
  End If
  
  SecondaryRecordRequired = fRequired
  
End Function

Private Function ValidateElement() As Boolean
  ' Return TRUE if the element has a valid definition.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim asMessages() As String
  Dim fContinue As Boolean
  Dim frmUsage As frmUsage
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  
  ReDim asMessages(0)
  
  Select Case mwfElement.ElementType
    Case elem_Begin
      ' No validation required.

    Case elem_Connector1, elem_Connector2
      ' No validation required.

    Case elem_Decision
      ' There must be a button/calc selected (as required).
      If optDecisionFlowType(decisionFlowType_Expression).value Then
        ' Expression
        If mlngDecisionFlowExpressionID <= 0 Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "Invalid '" & GetDecisionCaptionDescription(cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex), True) & "' flow calculation selected"
        End If
      Else
        ' Button
        If cboDecisionFlowButton.ItemData(cboDecisionFlowButton.ListIndex) = 0 Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "Invalid '" & GetDecisionCaptionDescription(cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex), True) & "' flow button selected"
        End If
      End If

    Case elem_Email
      If mlngEmailID <= 0 Then
        ReDim Preserve asMessages(UBound(asMessages) + 1)
        asMessages(UBound(asMessages)) = "No Email To address selected"
      End If

      If cboEmailRecord.ItemData(cboEmailRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD Then
        ' If we're working with an Identified record, then we must have an identified element,
        ' and if that element is a WebForm then we must have an identified RecordSelector.
        If cboEmailElement.ListCount = 0 Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "Invalid email record element"
        Else
          For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
            Set wfTemp = maWFPrecedingElements(lngLoop)
      
            If wfTemp.ControlIndex = cboEmailElement.ItemData(cboEmailElement.ListIndex) Then
              If wfTemp.ElementType = elem_WebForm Then
                If cboEmailRecordSelector.ListCount = 0 Then
                  ReDim Preserve asMessages(UBound(asMessages) + 1)
                  asMessages(UBound(asMessages)) = "Invalid email record selector"
                End If
              End If
      
              Exit For
            End If
      
            Set wfTemp = Nothing
          Next lngLoop
        End If
      End If
    
    Case elem_Or
      ' No validation required.

    Case elem_StoredData
      If (Len(Trim(txtIdentifier.Text)) = 0) Then
        ReDim Preserve asMessages(UBound(asMessages) + 1)
        asMessages(UBound(asMessages)) = "No identifier"
      Else
        If Not mfrmCallingForm.IsUniqueIdentifier(txtIdentifier.Text, mwfElement.ControlIndex) Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "Non-unique identifier"
        End If
      End If
      
      If cboDataRecord(STOREDDATATABLE_PRIMARY).ListCount = 0 Then
        ReDim Preserve asMessages(UBound(asMessages) + 1)
        asMessages(UBound(asMessages)) = "Invalid primary record"
      Else
        If cboDataRecord(STOREDDATATABLE_PRIMARY).ItemData(cboDataRecord(STOREDDATATABLE_PRIMARY).ListIndex) = giWFRECSEL_IDENTIFIEDRECORD Then
          ' If we're working with an Identified record, then we must have an identified element,
          ' and if that element is a WebForm then we must have an identified RecordSelector.
          If (cboDataElement(STOREDDATATABLE_PRIMARY).ListCount = 0) Then
            ReDim Preserve asMessages(UBound(asMessages) + 1)
            asMessages(UBound(asMessages)) = "Invalid primary record element"
          Else
            For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
              Set wfTemp = maWFPrecedingElements(lngLoop)
      
              If wfTemp.ControlIndex = cboDataElement(STOREDDATATABLE_PRIMARY).ItemData(cboDataElement(STOREDDATATABLE_PRIMARY).ListIndex) Then
                If wfTemp.ElementType = elem_WebForm Then
                  If (cboDataRecordSelector(STOREDDATATABLE_PRIMARY).ListCount = 0) Then
                    ReDim Preserve asMessages(UBound(asMessages) + 1)
                    asMessages(UBound(asMessages)) = "Invalid primary record selector"
                  End If
                End If
      
                Exit For
              End If
      
              Set wfTemp = Nothing
            Next lngLoop
          End If
        ElseIf cboDataRecord(STOREDDATATABLE_PRIMARY).ItemData(cboDataRecord(STOREDDATATABLE_PRIMARY).ListIndex) = giWFRECSEL_UNKNOWN Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "Invalid primary record"
        End If
      End If
      
      If cboDataAction.ItemData(cboDataAction.ListIndex) <> DATAACTION_DELETE Then
        If ssgrdColumns.Rows = 0 Then
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "No column values have been defined"
        End If
      End If
      
    Case elem_SummingJunction
      ' No validation required.

    Case elem_Terminator
      ' No validation required.

    Case elem_WebForm
      ' Now uses workflow web form designer.
      
  End Select
    
  ' Display the validity failures to the user.
  fContinue = (UBound(asMessages) = 0)
  
  If Not fContinue Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList

    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList (asMessages(iLoop))
    Next iLoop

    Screen.MousePointer = vbDefault
    frmUsage.ShowMessage "Workflow", "The " & mwfElement.ElementTypeDescription & " definition is invalid for the reasons listed below." & _
      vbCrLf & "Do you wish to continue?", UsageCheckObject.Workflow, _
      USAGEBUTTONS_YES + USAGEBUTTONS_NO + USAGEBUTTONS_PRINT, "validation"

    fContinue = (frmUsage.Choice = vbYes)

    UnLoad frmUsage
    Set frmUsage = Nothing
  End If
    
TidyUpAndExit:
  ValidateElement = fContinue
  Exit Function
  
ErrorTrap:
  fContinue = True
  Resume TidyUpAndExit
  
End Function
Private Function EmailName(piRecipient As EmailRecipients, _
  plngTableID As Long) As String
  
  ' Validate the current email, and return the name.
  Dim strEmailName As String
  Dim fValid As Boolean
  Dim lngTempEmailID As Long
  
  fValid = True
  strEmailName = ""
  
  Select Case piRecipient
    Case EMAIL_CC
      lngTempEmailID = mlngEmailCCID
    Case Else
      lngTempEmailID = mlngEmailID
  End Select
  
  If lngTempEmailID > 0 Then
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", lngTempEmailID
  
      fValid = (Not .NoMatch)
      
      If fValid Then
        fValid = Not !Deleted
      End If

      If fValid Then
        ' Definitely valid if the email address is 'fixed'.
        ' ie. doesn't matter what the base table is.
        fValid = (!Type = 0)
        
        If (Not fValid) And (plngTableID > 0) Then
          ' Tied to a table
          fValid = (!TableID = plngTableID)
        End If
        
        If fValid Then
          strEmailName = !Name
        End If
      End If
    End With
  
    If Not fValid Then
      Select Case piRecipient
        Case EMAIL_CC
          mlngEmailCCID = 0
        Case Else
          mlngEmailID = 0
      End Select
    End If
  End If
  
  EmailName = strEmailName
  
End Function

Private Sub cboDataAction_Click()
  If Not mfLoading Then
    cboDataRecord_Refresh STOREDDATATABLE_PRIMARY
    cboDataElement_Refresh STOREDDATATABLE_PRIMARY
    cboDataRecordSelector_Refresh STOREDDATATABLE_PRIMARY
    
    cboDataRecord_Refresh STOREDDATATABLE_SECONDARY
    cboDataElement_Refresh STOREDDATATABLE_SECONDARY
    cboDataRecordSelector_Refresh STOREDDATATABLE_SECONDARY
  End If
  
  Changed = True
  
End Sub

Private Sub cboDataRecord_Click(Index As Integer)
  cboDataElement_Refresh Index
  
  Changed = True
  
End Sub

Private Sub cboDataRecordSelector_Click(Index As Integer)
  cboDataRecordTable_Refresh Index
  
  Changed = True

End Sub

Private Sub cboDataRecordTable_Click(Index As Integer)

  If miDataRecordBeingRefreshed <> Index Then
    miDataRecordBeingRefreshed = (1 - Index)
    cboDataRecord_Refresh (1 - Index)
    miDataRecordBeingRefreshed = -1
  End If
  
  Changed = True

End Sub


Private Sub cboDataTable_Click()
  Dim lngTableID As Long
  Dim fChangeApproved As Boolean
  Dim iLoop As Integer
  Dim fTableChanged As Boolean
  
  lngTableID = 0
  fChangeApproved = True
  fTableChanged = False
  
  With ssgrdColumns
    If .Rows > 0 Then
      .MoveFirst
      lngTableID = GetTableIDFromColumnID(CLng(.Columns(2).Text))
    End If
  End With
  
  If (lngTableID > 0) _
    And (lngTableID <> cboDataTable.ItemData(cboDataTable.ListIndex)) Then
    
    fTableChanged = True
    fChangeApproved = (MsgBox("Changing the table will remove all selected columns. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton1, App.ProductName) = vbYes)
  End If
  
  If fChangeApproved Then
    If fTableChanged Then
      ssgrdColumns.RemoveAll
    
      ResizeGridColumns ssgrdColumns
    End If
    
    Changed = True
    cboDataRecord_Refresh STOREDDATATABLE_PRIMARY
  Else
    'Reselect the original table.
    For iLoop = 0 To cboDataTable.ListCount - 1
      If cboDataTable.ItemData(iLoop) = lngTableID Then
        cboDataTable.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
End Sub

Private Sub cboDataElement_Click(Index As Integer)
  cboDataRecordSelector_Refresh Index
  
  Changed = True

End Sub

Private Sub cboDecisionCaption_Click()
  Changed = True

  fraDecisionFlow.Caption = "'" & GetDecisionCaptionDescription(cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex), True) & "' flow criteria :"
  
End Sub

Private Sub cboEmailRecord_Click()
  cboEmailElement_Refresh
  Changed = True
  
End Sub

Private Sub cboEmailRecordSelector_Click()
  cboEmailTable_Refresh
  Changed = True

End Sub

Private Sub cboEmailElement_Click()
  cboEmailRecordSelector_Refresh
  Changed = True
  
End Sub

Private Sub cboEmailTable_Click()
  Dim lngTableID As Long
  
  lngTableID = -1
  
  If cboEmailTable.ListIndex >= 0 Then
    lngTableID = cboEmailTable.ItemData(cboEmailTable.ListIndex)
  End If
  
  txtEmail.Text = EmailName(EMAIL_TO, lngTableID)
  txtEmailCC.Text = EmailName(EMAIL_CC, lngTableID)
    
  Changed = True

End Sub


Private Sub cboDecisionFlowButton_Click()
  Changed = True
End Sub

Private Sub cmdAddItem_Click()

  Dim sRow As String
  Dim frmItem As New frmWorkflowElementItem
  Dim sDescription As String
  Dim sTemp As String
  Dim lngCalcID As Long
  
  With frmItem
    .Initialize Me, _
      mwfElement.ElementType, _
      giWFEMAILITEM_LABEL, _
      "", _
      0, _
      0, _
      "", _
      "", _
      True, _
      "", _
      "", _
      0, _
      False, _
      ""

    .Show vbModal

    If Not .Cancelled Then
      sDescription = "<unknown>"
      
      Select Case .ItemType
        Case giWFEMAILITEM_DBVALUE
          sDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        Case giWFEMAILITEM_LABEL
          sDescription = "Label - '" & .ItemCaption & "'"
        Case giWFEMAILITEM_WFVALUE
          sDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
        Case giWFEMAILITEM_FORMATCODE
          sDescription = "Formatting - " & FormatDescription(.ItemCaption)
        Case giWFEMAILITEM_CALCULATION
          lngCalcID = .CalculationID
          sDescription = "Calculation - "

          sTemp = GetExpressionName(lngCalcID)
          If Len(Trim(sTemp)) = 0 Then
            sTemp = "<unknown>"
            lngCalcID = 0
          Else
            sTemp = "<" & sTemp & ">"
          End If
          sDescription = sDescription & sTemp
      
      End Select
      
      sRow = sDescription _
        & vbTab & .ItemType _
        & vbTab & .ItemCaption _
        & vbTab & CStr(.ItemDBColumnID) _
        & vbTab & CStr(.ItemDBRecord) _
        & vbTab & .ItemWFFormIdentifier _
        & vbTab & .ItemWFValueIdentifier _
        & vbTab & .ItemDBWebForm _
        & vbTab & .ItemDBRecordSelector _
        & vbTab & CStr(lngCalcID)

      With ssgrdItems
        .AddItem sRow
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .Update

        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      
      Changed = True
    End If
  
    RefreshExpressionNames
    ResizeGridColumns ssgrdItems
  End With

  UnLoad frmItem
  Set frmItem = Nothing

End Sub
Private Sub RefreshExpressionNames()
  ' Refresh any calculation titles in the grid.
  Dim varBookmark As Variant
  Dim lngCalcID As Long
  Dim sDescription As String
  Dim sTemp As String
  Dim iLoop As Integer
  
  UI.LockWindow ssgrdItems.hWnd

  With ssgrdItems
    varBookmark = .Bookmark
    
    .MoveFirst

    For iLoop = 0 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      
      If val(.Columns("ItemType").Text) = giWFEMAILITEM_CALCULATION Then
        lngCalcID = val(.Columns("CalculationID").Text)
        sDescription = "Calculation - "
        sTemp = GetExpressionName(lngCalcID)
        
        If Len(Trim(sTemp)) = 0 Then
          sTemp = "<unknown>"
          lngCalcID = 0
        Else
          sTemp = "<" & sTemp & ">"
        End If
        sDescription = sDescription & sTemp
  
        .Columns("CalculationID").Text = CStr(lngCalcID)
        .Columns("Description").Text = sDescription
        
        .Update
      End If
      
      .MoveNext
    Next iLoop

    .MoveFirst
    .Bookmark = varBookmark
  End With

  UI.UnlockWindow

End Sub



Private Sub cmdAttachAttachment_Click()
  Dim frmItem As New frmWorkflowElementItem
  Dim sDescription As String

  If (miEmailAttachmentType <> giWFEMAILITEM_DBVALUE) _
    And (miEmailAttachmentType <> giWFEMAILITEM_WFVALUE) Then
    
    miEmailAttachmentType = giWFEMAILITEM_FILEATTACHMENT
  End If
  
  With frmItem
    .Initialize Me, _
      mwfElement.ElementType, _
      miEmailAttachmentType, _
      "", _
      mlngEmailAttachment_DBColumnID, _
      miEmailAttachment_DBRecord, _
      msEmailAttachment_WFElementIdentifier, _
      msEmailAttachment_WFItemIdentifier, _
      False, _
      msEmailAttachment_DBElementIdentifier, _
      msEmailAttachment_DBItemIdentifier, _
      0, _
      True, _
      msEmailAttachment_File

    .Show vbModal

    If Not .Cancelled Then
      sDescription = "<unknown>"
      
      Select Case .ItemType
        Case giWFEMAILITEM_DBVALUE
          sDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        Case giWFEMAILITEM_WFVALUE
          sDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
        Case giWFEMAILITEM_FILEATTACHMENT
          sDescription = "File - '" & .FileAttachment & "'"
      End Select

      miEmailAttachmentType = .ItemType
      msEmailAttachment_File = .FileAttachment
      msEmailAttachment_WFElementIdentifier = .ItemWFFormIdentifier
      msEmailAttachment_WFItemIdentifier = .ItemWFValueIdentifier
      mlngEmailAttachment_DBColumnID = .ItemDBColumnID
      miEmailAttachment_DBRecord = .ItemDBRecord
      msEmailAttachment_DBElementIdentifier = .ItemDBWebForm
      msEmailAttachment_DBItemIdentifier = .ItemDBRecordSelector

      txtAttachAttachment.Text = sDescription
      
      Changed = True
    End If
  End With

  UnLoad frmItem
  Set frmItem = Nothing
  
End Sub

Private Sub cmdAttachClear_Click()
  miEmailAttachmentType = giWFEMAILITEM_UNKNOWN
  msEmailAttachment_File = ""
  msEmailAttachment_WFElementIdentifier = ""
  msEmailAttachment_WFItemIdentifier = ""
  mlngEmailAttachment_DBColumnID = 0
  miEmailAttachment_DBRecord = 0
  msEmailAttachment_DBElementIdentifier = ""
  msEmailAttachment_DBItemIdentifier = ""
   
  txtAttachAttachment.Text = ""
   
  Changed = True
  
End Sub


Private Sub cmdCancel_Click()
  mfCancelled = True
  
  UnLoad Me

End Sub

Private Sub cmdCopyItem_Click()
  Dim sRow As String
  Dim lngRow As Long
  Dim frmItem As New frmWorkflowElementItem
  Dim sDescription As String
  Dim lngCalcID As Long
  Dim sTemp As String
  
  ssgrdItems.Bookmark = ssgrdItems.SelBookmarks(0)
  lngRow = ssgrdItems.AddItemRowIndex(ssgrdItems.Bookmark)

  With frmItem
    .Initialize Me, _
      mwfElement.ElementType, _
      CInt(ssgrdItems.Columns("ItemType").Text), _
      ssgrdItems.Columns("Caption").Text, _
      CLng(ssgrdItems.Columns("DBColumnID").Text), _
      CInt(ssgrdItems.Columns("DBRecord").Text), _
      ssgrdItems.Columns("WFFormIdentifier").Text, _
      ssgrdItems.Columns("WFValueIdentifier").Text, _
      True, _
      ssgrdItems.Columns("DBWebForm").Text, _
      ssgrdItems.Columns("DBRecordSelector").Text, _
      val(ssgrdItems.Columns("CalculationID").Text), _
      False, _
      ""

    .Show vbModal

    If Not .Cancelled Then
      sDescription = "<unknown>"
      
      Select Case .ItemType
        Case giWFEMAILITEM_DBVALUE
          sDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        Case giWFEMAILITEM_LABEL
          sDescription = "Label - '" & .ItemCaption & "'"
        Case giWFEMAILITEM_WFVALUE
          sDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
        Case giWFEMAILITEM_FORMATCODE
          sDescription = "Formatting - " & FormatDescription(.ItemCaption)
        Case giWFEMAILITEM_CALCULATION
          lngCalcID = .CalculationID
          sDescription = "Calculation - "

          sTemp = GetExpressionName(lngCalcID)
          If Len(Trim(sTemp)) = 0 Then
            sTemp = "<unknown>"
            lngCalcID = 0
          Else
            sTemp = "<" & sTemp & ">"
          End If
          sDescription = sDescription & sTemp
      End Select

      sRow = sDescription _
        & vbTab & .ItemType _
        & vbTab & .ItemCaption _
        & vbTab & CStr(.ItemDBColumnID) _
        & vbTab & CStr(.ItemDBRecord) _
        & vbTab & .ItemWFFormIdentifier _
        & vbTab & .ItemWFValueIdentifier _
        & vbTab & .ItemDBWebForm _
        & vbTab & .ItemDBRecordSelector _
        & vbTab & CStr(lngCalcID)

      With ssgrdItems
        .AddItem sRow, lngRow
        .Bookmark = .AddItemBookmark(lngRow)
        .Update

        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .AddItemBookmark(lngRow)
      End With

      Changed = True
    End If
  
    RefreshExpressionNames
    ResizeGridColumns ssgrdItems
  End With

  UnLoad frmItem
  Set frmItem = Nothing

End Sub

Private Sub cmdDataAdd_Click()
  Dim sRow As String
  Dim frmColumn As New frmWorkflowElementColumn
  Dim sColumnDescription As String
  Dim sValueDescription As String
  Dim alngColumnsDone() As Long
  Dim iLoop As Integer
  Dim objMisc As Misc
  Dim sTemp As String
  Dim lngCalcID As Long
  
  With ssgrdColumns
    ReDim alngColumnsDone(.Rows)
    
    .MoveFirst
    For iLoop = 1 To .Rows
      alngColumnsDone(iLoop) = CLng(.Columns("ColumnID").Text)
      
      .MoveNext
    Next iLoop
  End With

  With frmColumn
    .Initialize Me, _
      alngColumnsDone, _
      cboDataTable.ItemData(cboDataTable.ListIndex), _
      0, _
      giWFDATAVALUE_FIXED, _
      "", _
      "", _
      "", _
      0, _
      giWFRECSEL_UNKNOWN, _
      True, _
      0
    
    .Show vbModal

    If Not .Cancelled Then
      sColumnDescription = GetColumnName(.ColumnID, True)

      sValueDescription = "<unknown>"
      Select Case .ValueType
        Case giWFDATAVALUE_FIXED
          sValueDescription = "Fixed value - " & .value
          
          If (GetColumnDataType(.ColumnID) = dtTIMESTAMP) _
            And UCase(.value) <> "NULL" Then
            Set objMisc = New Misc
            sValueDescription = "Fixed value - " & objMisc.ConvertSQLDateToLocale(.value)
            Set objMisc = Nothing
          End If
        
        Case giWFDATAVALUE_WFVALUE
          sValueDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
        
        Case giWFDATAVALUE_DBVALUE
          sValueDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        
        Case giWFDATAVALUE_CALC
          lngCalcID = .CalculationID
          sValueDescription = "Calculation - "

          sTemp = GetExpressionName(lngCalcID)
          If Len(Trim(sTemp)) = 0 Then
            sTemp = "<unknown>"
            lngCalcID = 0
          Else
            sTemp = "<" & sTemp & ">"
          End If
          sValueDescription = sValueDescription & sTemp
      End Select

      sRow = sColumnDescription _
        & vbTab & sValueDescription _
        & vbTab & .ColumnID _
        & vbTab & .ValueType _
        & vbTab & .value _
        & vbTab & .ItemWFFormIdentifier _
        & vbTab & .ItemWFValueIdentifier _
        & vbTab & .ItemDBColumnID _
        & vbTab & .ItemDBRecord _
        & vbTab & CStr(lngCalcID)
      
      With ssgrdColumns
        .AddItem sRow
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .Update

        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      ResizeGridColumns ssgrdColumns

      Changed = True
    End If
  End With

  UnLoad frmColumn
  Set frmColumn = Nothing

End Sub



Public Property Get BaseTable() As Long
  BaseTable = mlngBaseTableID
  
End Property

Public Property Let BaseTable(ByVal plngNewValue As Long)
  mlngBaseTableID = plngNewValue

End Property
Public Property Get InitiationType() As WorkflowInitiationTypes
  InitiationType = miInitiationType
  
End Property

Public Property Let InitiationType(ByVal piNewValue As WorkflowInitiationTypes)
  miInitiationType = piNewValue

End Property


Private Sub cmdDataEdit_Click()
  Dim sRow As String
  Dim lngRow As Long
  Dim frmColumn As New frmWorkflowElementColumn
  Dim sColumnDescription As String
  Dim sValueDescription As String
  Dim alngColumnsDone() As Long
  Dim iLoop As Integer
  Dim lngCurrentColumnID As Long
  Dim objMisc As Misc
  Dim sTemp As String
  Dim lngCalcID As Long
  
  With ssgrdColumns
    .Bookmark = .SelBookmarks(0)
    lngCurrentColumnID = CLng(.Columns("ColumnID").Text)
    
    ReDim alngColumnsDone(.Rows)
    
    .MoveFirst
    For iLoop = 1 To .Rows
      If CLng(.Columns("ColumnID").Text) = lngCurrentColumnID Then
        alngColumnsDone(iLoop) = 0
      Else
        alngColumnsDone(iLoop) = CLng(.Columns("ColumnID").Text)
      End If
       
      .MoveNext
    Next iLoop
  End With
  
  ssgrdColumns.Bookmark = ssgrdColumns.SelBookmarks(0)
  lngRow = ssgrdColumns.AddItemRowIndex(ssgrdColumns.Bookmark)

  With frmColumn
    .Initialize Me, _
      alngColumnsDone, _
      cboDataTable.ItemData(cboDataTable.ListIndex), _
      CLng(ssgrdColumns.Columns("ColumnID").Text), _
      CInt(ssgrdColumns.Columns("ValueType").Text), _
      ssgrdColumns.Columns("Value").Text, _
      ssgrdColumns.Columns("WFFormIdentifier").Text, _
      ssgrdColumns.Columns("WFValueIdentifier").Text, _
      CLng(IIf(Len(ssgrdColumns.Columns("DBColumnID").Text) = 0, "0", ssgrdColumns.Columns("DBColumnID").Text)), _
      CInt(IIf(Len(ssgrdColumns.Columns("DBRecord").Text) = 0, "0", ssgrdColumns.Columns("DBRecord").Text)), _
      False, _
      val(ssgrdColumns.Columns("CalculationID").Text)

    .Show vbModal

    If Not .Cancelled Then
      sColumnDescription = GetColumnName(.ColumnID, True)

      sValueDescription = "<unknown>"
      Select Case .ValueType
        Case giWFDATAVALUE_FIXED
          sValueDescription = "Fixed value - " & .value
          
          If (GetColumnDataType(.ColumnID) = dtTIMESTAMP) _
            And UCase(.value) <> "NULL" Then
            Set objMisc = New Misc
            sValueDescription = "Fixed value - " & objMisc.ConvertSQLDateToLocale(.value)
            Set objMisc = Nothing
          End If
          
        Case giWFDATAVALUE_WFVALUE
          sValueDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
          
        Case giWFDATAVALUE_DBVALUE
          sValueDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        
         Case giWFDATAVALUE_CALC
          lngCalcID = .CalculationID
          sValueDescription = "Calculation - "

          sTemp = GetExpressionName(lngCalcID)
          If Len(Trim(sTemp)) = 0 Then
            sTemp = "<unknown>"
            lngCalcID = 0
          Else
            sTemp = "<" & sTemp & ">"
          End If
          sValueDescription = sValueDescription & sTemp
          
      End Select

      sRow = sColumnDescription _
        & vbTab & sValueDescription _
        & vbTab & .ColumnID _
        & vbTab & .ValueType _
        & vbTab & .value _
        & vbTab & .ItemWFFormIdentifier _
        & vbTab & .ItemWFValueIdentifier _
        & vbTab & .ItemDBColumnID _
        & vbTab & .ItemDBRecord _
        & vbTab & CStr(lngCalcID)

      ssgrdColumns.RemoveItem lngRow

      With ssgrdColumns
        .AddItem sRow, lngRow
        .Bookmark = .AddItemBookmark(lngRow)
        .Update

        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .AddItemBookmark(lngRow)
      End With
      ResizeGridColumns ssgrdColumns

      Changed = True
    End If
  End With

  UnLoad frmColumn
  Set frmColumn = Nothing

End Sub

Private Sub cmdDataRemove_Click()
  RemoveItem ssgrdColumns

End Sub

Private Sub cmdDataRemoveAll_Click()
  ssgrdColumns.RemoveAll
  ResizeGridColumns ssgrdColumns
  Changed = True

End Sub

Private Sub cmdDecisionFlowExpression_Click()
  Dim objExpr As CExpression
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    .Initialise 0, mlngDecisionFlowExpressionID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_LOGIC
    .UtilityID = mfrmCallingForm.WorkflowID
    .UtilityBaseTable = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL, mlngPersonnelTableID, _
      IIf(miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED, mlngBaseTableID, 0))
    .WorkflowInitiationType = miInitiationType
    .PrecedingWorkflowElements = maWFPrecedingElements
    .AllWorkflowElements = maWFAllElements
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(mfReadOnly) Then
      mlngDecisionFlowExpressionID = .ExpressionID
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mlngDecisionFlowExpressionID, False

        If .NoMatch Then
          mlngDecisionFlowExpressionID = 0
        End If
      End With
    End If
  
    ' Read the selected expression info.
    txtDecisionFlowExpression.Text = GetExpressionName(mlngDecisionFlowExpressionID)
  End With
  
  Set objExpr = Nothing
  Changed = True

End Sub

Private Sub cmdEditItem_Click()
  Dim sRow As String
  Dim lngRow As Long
  Dim frmItem As New frmWorkflowElementItem
  Dim sDescription As String
  Dim sTemp As String
  Dim lngCalcID As Long

  ssgrdItems.Bookmark = ssgrdItems.SelBookmarks(0)
  lngRow = ssgrdItems.AddItemRowIndex(ssgrdItems.Bookmark)

  With frmItem
    .Initialize Me, _
      mwfElement.ElementType, _
      CInt(ssgrdItems.Columns("ItemType").Text), _
      ssgrdItems.Columns("Caption").Text, _
      CLng(ssgrdItems.Columns("DBColumnID").Text), _
      CInt(ssgrdItems.Columns("DBRecord").Text), _
      ssgrdItems.Columns("WFFormIdentifier").Text, _
      ssgrdItems.Columns("WFValueIdentifier").Text, _
      False, _
      ssgrdItems.Columns("DBWebForm").Text, _
      ssgrdItems.Columns("DBRecordSelector").Text, _
      val(ssgrdItems.Columns("CalculationID").Text), _
      False, _
      ""

    .Show vbModal

    If Not .Cancelled Then
      sDescription = "<unknown>"
      
      Select Case .ItemType
        Case giWFEMAILITEM_DBVALUE
          sDescription = "Database value - " & GetColumnName(.ItemDBColumnID)
        Case giWFEMAILITEM_LABEL
          sDescription = "Label - '" & .ItemCaption & "'"
        Case giWFEMAILITEM_WFVALUE
          sDescription = "Workflow value - " & .ItemWFFormIdentifier & "." & .ItemWFValueIdentifier
        Case giWFEMAILITEM_FORMATCODE
          sDescription = "Formatting - " & FormatDescription(.ItemCaption)
        Case giWFEMAILITEM_CALCULATION
          lngCalcID = .CalculationID
          sDescription = "Calculation - "

          sTemp = GetExpressionName(lngCalcID)
          If Len(Trim(sTemp)) = 0 Then
            sTemp = "<unknown>"
            lngCalcID = 0
          Else
            sTemp = "<" & sTemp & ">"
          End If
          sDescription = sDescription & sTemp
      End Select
      
      sRow = sDescription _
        & vbTab & .ItemType _
        & vbTab & .ItemCaption _
        & vbTab & CStr(.ItemDBColumnID) _
        & vbTab & CStr(.ItemDBRecord) _
        & vbTab & .ItemWFFormIdentifier _
        & vbTab & .ItemWFValueIdentifier _
        & vbTab & .ItemDBWebForm _
        & vbTab & .ItemDBRecordSelector _
        & vbTab & CStr(lngCalcID)

      ssgrdItems.RemoveItem lngRow
        
      With ssgrdItems
        .AddItem sRow, lngRow
        .Bookmark = .AddItemBookmark(lngRow)
        .Update

        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .AddItemBookmark(lngRow)
      End With

      Changed = True
    End If
  
    RefreshExpressionNames
    ResizeGridColumns ssgrdItems
  End With

  UnLoad frmItem
  Set frmItem = Nothing

End Sub

Private Sub cmdEmail_Click()
  ' Display the Email selection form.
  Dim objEmail As clsEmailAddr
  Dim wfTemp As VB.Control
  
  ' Create a new Email object.
  Set objEmail = New clsEmailAddr

  ' Initialize the Email object.
  With objEmail
    .EmailID = mlngEmailID
    
    If cboEmailTable.ListIndex >= 0 Then
      .TableID = cboEmailTable.ItemData(cboEmailTable.ListIndex)
    Else
      .TableID = 0
    End If
    
    ' Instruct the Email object to handle the selection.
    If .SelectEmail(mfReadOnly) Then
      mlngEmailID = .EmailID
      txtEmail.Text = .EmailName

      Changed = True
    Else
      ' Check in case the original Email has been deleted.
      With recEmailAddrEdit
        .Index = "idxID"
        .Seek "=", mlngEmailID

        If .NoMatch Then
          mlngEmailID = 0
          txtEmail.Text = vbNullString
        Else
          If !Deleted Then
            mlngEmailID = 0
            txtEmail.Text = vbNullString
          End If
        End If
      End With
    End If
  End With

  ' Disassociate object variables.
  Set objEmail = Nothing

End Sub

Private Sub cmdEmailCC_Click()
  ' Display the Email selection form.
  Dim objEmail As clsEmailAddr
  Dim wfTemp As VB.Control

  ' Create a new Email object.
  Set objEmail = New clsEmailAddr

  ' Initialize the Email object.
  With objEmail
    .EmailID = mlngEmailCCID

    If cboEmailTable.ListIndex >= 0 Then
      .TableID = cboEmailTable.ItemData(cboEmailTable.ListIndex)
    Else
      .TableID = 0
    End If

    ' Instruct the Email object to handle the selection.
    If .SelectEmail(mfReadOnly) Then
      mlngEmailCCID = .EmailID
      txtEmailCC.Text = .EmailName

      Changed = True
    Else
      ' Check in case the original Email has been deleted.
      With recEmailAddrEdit
        .Index = "idxID"
        .Seek "=", mlngEmailCCID

        If .NoMatch Then
          mlngEmailCCID = 0
          txtEmailCC.Text = vbNullString
        Else
          If !Deleted Then
            mlngEmailCCID = 0
            txtEmailCC.Text = vbNullString
          End If
        End If
      End With
    End If
  End With

  ' Disassociate object variables.
  Set objEmail = Nothing

End Sub


Private Sub cmdMoveItemDown_Click()
  MoveItem ssgrdItems, MOVEDIRECTION_DOWN

End Sub

Private Sub MoveItem(pctlGrid As SSDBGrid, _
  piDirection As MoveDirection)
  Dim iLoop As Integer
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = pctlGrid.AddItemRowIndex(pctlGrid.Bookmark)
  
  For iLoop = 0 To pctlGrid.Columns.Count - 1
    strSourceRow = strSourceRow & pctlGrid.Columns(iLoop).Text & _
      IIf(iLoop = pctlGrid.Columns.Count - 1, "", vbTab)
  Next iLoop
  
  If piDirection = MOVEDIRECTION_UP Then
    intDestinationRow = intSourceRow - 1
    pctlGrid.MovePrevious
  Else
    intDestinationRow = intSourceRow + 1
    pctlGrid.MoveNext
  End If
  
  For iLoop = 0 To pctlGrid.Columns.Count - 1
    strDestinationRow = strDestinationRow & pctlGrid.Columns(iLoop).Text & _
      IIf(iLoop = pctlGrid.Columns.Count - 1, "", vbTab)
  Next iLoop
  
  If piDirection = MOVEDIRECTION_UP Then
    pctlGrid.AddItem strSourceRow, intDestinationRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
    pctlGrid.Update
    
    pctlGrid.RemoveItem intSourceRow + 1
    
    pctlGrid.SelBookmarks.RemoveAll
    pctlGrid.MovePrevious
  Else
    pctlGrid.RemoveItem intDestinationRow
    pctlGrid.RemoveItem intSourceRow
    
    pctlGrid.AddItem strDestinationRow, intSourceRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intSourceRow)
    pctlGrid.Update
    
    pctlGrid.AddItem strSourceRow, intDestinationRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
    pctlGrid.Update
    
    pctlGrid.SelBookmarks.RemoveAll
    pctlGrid.MoveNext
  End If
  
  pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
  pctlGrid.SelBookmarks.Add pctlGrid.AddItemBookmark(intDestinationRow)
  
  Changed = True

End Sub



Private Sub cmdMoveItemUp_Click()
  MoveItem ssgrdItems, MOVEDIRECTION_UP

End Sub


Private Sub cmdOK_Click()
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim iValue As Integer
  Dim optTemp As OptionButton
  Dim awfElements() As VB.Control
  
  ReDim asItems(0)
  ReDim avColumns(0)
  
  If Not ValidateElement Then
    Exit Sub
  End If
  
  mwfElement.Caption = txtCaption.Text
      
  Select Case mwfElement.ElementType
    Case elem_Begin
    
    Case elem_Connector1, elem_Connector2
      
    Case elem_Decision
      mwfElement.DecisionCaptionType = cboDecisionCaption.ItemData(cboDecisionCaption.ListIndex)
      For Each optTemp In optDecisionFlowType
        If optTemp.value Then
          iValue = optTemp.Index
          Exit For
        End If
      Next optTemp
      Set optTemp = Nothing
      mwfElement.DecisionFlowType = iValue
      mwfElement.TrueFlowIdentifier = IIf(mwfElement.DecisionFlowType = decisionFlowType_Button, cboDecisionFlowButton.List(cboDecisionFlowButton.ListIndex), "")
      mwfElement.DecisionFlowExpressionID = IIf(mwfElement.DecisionFlowType = decisionFlowType_Expression, mlngDecisionFlowExpressionID, 0)
      
    Case elem_Email
      mwfElement.EmailID = mlngEmailID
      mwfElement.EmailCCID = mlngEmailCCID
      mwfElement.EmailRecord = cboEmailRecord.ItemData(cboEmailRecord.ListIndex)

      If mwfElement.EmailRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        mwfElement.RecordSelectorWebFormIdentifier = cboEmailElement.List(cboEmailElement.ListIndex)
        mwfElement.RecordSelectorIdentifier = cboEmailRecordSelector.List(cboEmailRecordSelector.ListIndex)
      Else
        mwfElement.RecordSelectorWebFormIdentifier = ""
        mwfElement.RecordSelectorIdentifier = ""
      End If
      
      mwfElement.EMailSubject = txtEmailSubject.Text
      
      mwfElement.Attachment_Type = miEmailAttachmentType
      mwfElement.Attachment_File = msEmailAttachment_File
      mwfElement.Attachment_WFElementIdentifier = msEmailAttachment_WFElementIdentifier
      mwfElement.Attachment_WFValueIdentifier = msEmailAttachment_WFItemIdentifier
      mwfElement.Attachment_DBColumnID = mlngEmailAttachment_DBColumnID
      mwfElement.Attachment_DBRecord = miEmailAttachment_DBRecord
      mwfElement.Attachment_DBElement = msEmailAttachment_DBElementIdentifier
      mwfElement.Attachment_DBValue = msEmailAttachment_DBItemIdentifier

      With ssgrdItems
        ReDim asItems(WFITEMPROPERTYCOUNT, .Rows)
        
        .MoveFirst
        For iLoop2 = 1 To .Rows
          asItems(1, iLoop2) = ssgrdItems.Columns("Description").Text ' Description
          asItems(2, iLoop2) = ssgrdItems.Columns("ItemType").Text ' Item Type
          asItems(3, iLoop2) = ssgrdItems.Columns("Caption").Text ' Caption
          asItems(4, iLoop2) = ssgrdItems.Columns("DBColumnID").Text ' DB Column ID
          asItems(5, iLoop2) = ssgrdItems.Columns("DBRecord").Text ' DB Record
          asItems(11, iLoop2) = ssgrdItems.Columns("WFFormIdentifier").Text ' WF Form Identifier
          asItems(12, iLoop2) = ssgrdItems.Columns("WFValueIdentifier").Text ' WF Value Identifier
          asItems(13, iLoop2) = ssgrdItems.Columns("DBWebForm").Text ' DB Web Form Identifier
          asItems(14, iLoop2) = ssgrdItems.Columns("DBRecordSelector").Text ' DB RecSel Identifier
          asItems(56, iLoop2) = ssgrdItems.Columns("CalculationID").Text ' Calculation ID
          
          .MoveNext
        Next iLoop2
      End With

      mwfElement.Items = asItems

    Case elem_Or
      ' Not required.
      
    Case elem_StoredData
      mwfElement.Identifier = txtIdentifier.Text
      
      mwfElement.DataAction = cboDataAction.ItemData(cboDataAction.ListIndex)
      If cboDataTable.Enabled Then
        mwfElement.DataTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
      Else
        mwfElement.DataTableID = 0
      End If
      
      mwfElement.DataRecord = cboDataRecord(STOREDDATATABLE_PRIMARY).ItemData(cboDataRecord(STOREDDATATABLE_PRIMARY).ListIndex)
      If mwfElement.DataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        mwfElement.RecordSelectorWebFormIdentifier = cboDataElement(STOREDDATATABLE_PRIMARY).List(cboDataElement(STOREDDATATABLE_PRIMARY).ListIndex)
        mwfElement.RecordSelectorIdentifier = cboDataRecordSelector(STOREDDATATABLE_PRIMARY).List(cboDataRecordSelector(STOREDDATATABLE_PRIMARY).ListIndex)
      Else
        mwfElement.RecordSelectorWebFormIdentifier = ""
        mwfElement.RecordSelectorIdentifier = ""
      End If
      If (cboDataRecordTable(STOREDDATATABLE_PRIMARY).ListIndex >= 0) Then
        mwfElement.DataRecordTableID = cboDataRecordTable(STOREDDATATABLE_PRIMARY).ItemData(cboDataRecordTable(STOREDDATATABLE_PRIMARY).ListIndex)
      Else
        mwfElement.DataRecordTableID = 0
      End If

      mwfElement.SecondaryDataRecord = cboDataRecord(STOREDDATATABLE_SECONDARY).ItemData(cboDataRecord(STOREDDATATABLE_SECONDARY).ListIndex)
      If mwfElement.SecondaryDataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        mwfElement.SecondaryRecordSelectorWebFormIdentifier = cboDataElement(STOREDDATATABLE_SECONDARY).List(cboDataElement(STOREDDATATABLE_SECONDARY).ListIndex)
        mwfElement.SecondaryRecordSelectorIdentifier = cboDataRecordSelector(STOREDDATATABLE_SECONDARY).List(cboDataRecordSelector(STOREDDATATABLE_SECONDARY).ListIndex)
      Else
        mwfElement.SecondaryRecordSelectorWebFormIdentifier = ""
        mwfElement.SecondaryRecordSelectorIdentifier = ""
      End If
      If (cboDataRecordTable(STOREDDATATABLE_SECONDARY).ListIndex >= 0) Then
        mwfElement.SecondaryDataRecordTableID = cboDataRecordTable(STOREDDATATABLE_SECONDARY).ItemData(cboDataRecordTable(STOREDDATATABLE_SECONDARY).ListIndex)
      Else
        mwfElement.SecondaryDataRecordTableID = 0
      End If
            
      With ssgrdColumns
        ReDim avColumns(.Columns.Count + 1, .Rows)
        
        .MoveFirst
        For iLoop2 = 1 To .Rows
          For iLoop = 1 To .Columns.Count
            avColumns(iLoop, iLoop2) = .Columns(iLoop - 1).Text
          Next iLoop
          
          .MoveNext
        Next iLoop2
      End With

      ReDim awfElements(1)
      Set awfElements(1) = mwfElement
      mavIdentifierLog(3, 0) = mwfElement.Identifier
      mavIdentifierLog(6, 0) = mwfElement.DataTableID
      mfrmCallingForm.UpdateIdentifiers mwfElement, awfElements, mavIdentifierLog

      mwfElement.DataColumns = avColumns

    Case elem_SummingJunction
      ' Not required.
      
    Case elem_Terminator
    
    Case elem_WebForm
      ' Now uses workflow web form designer.
  
  End Select
    
  mfCancelled = False
  UnLoad Me

End Sub

Private Sub cmdRemoveAllItems_Click()
  ssgrdItems.RemoveAll
  ResizeGridColumns ssgrdItems
  Changed = True

End Sub

Private Sub cmdRemoveItem_Click()
  RemoveItem ssgrdItems

End Sub

Private Sub RemoveItem(pctlGrid As SSDBGrid)
  Dim sRowsToDelete As String
  Dim iCount As Integer
  
  sRowsToDelete = ","

  With pctlGrid
    If .Rows = 1 Then
      .RemoveAll
    Else
      For iCount = 0 To .SelBookmarks.Count - 1
        sRowsToDelete = sRowsToDelete & CStr(.AddItemRowIndex(.SelBookmarks(iCount))) & ","
      Next iCount
      
      For iCount = (.Rows - 1) To 0 Step -1
        If InStr(sRowsToDelete, "," & CStr(iCount) & ",") > 0 Then
          .RemoveItem iCount
        End If
      Next iCount
    End If
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  ResizeGridColumns pctlGrid
  
  Changed = True

End Sub

Private Sub cboDataRecord_Refresh(piIndex As Integer)
  ' Populate the DataRecord combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim fRecSelOK As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim lngLoop As Long
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim lngTableID As Long
  Dim fTableIsInRecSel As Boolean
  Dim fParentTableIsInRecSel As Boolean
  Dim iTableType As enum_TableTypes
  Dim sTableIDs As String
  Dim iAction As DataAction
  Dim iRecordSelection As WorkflowRecordSelectorTypes
  Dim fSecondaryRecordRequired As Boolean
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngFirstDataRecordTable As Long
  Dim sMsg As String
  
  iIndex = -1

  ' Secondary record info only required if the data action is INSERTing into a shared history table.
  fSecondaryRecordRequired = SecondaryRecordRequired
  
  If (piIndex = STOREDDATATABLE_SECONDARY) Then
    lngFirstDataRecordTable = DataRecordTable(STOREDDATATABLE_PRIMARY)
  End If
  
  With cboDataRecord(piIndex)
    If .ListIndex < 0 Then
      iRecordSelection = IIf(piIndex = STOREDDATATABLE_PRIMARY, mwfElement.DataRecord, mwfElement.SecondaryDataRecord)
    Else
      iRecordSelection = .ItemData(.ListIndex)
    End If
    
    ' Clear the current contents of the combo.
    .Clear

    If (cboDataTable.ListIndex >= 0) _
      And (piIndex = STOREDDATATABLE_PRIMARY Or fSecondaryRecordRequired) Then
      
      lngTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
      iAction = cboDataAction.ItemData(cboDataAction.ListIndex)
      
      ' --------------------------------------------------------------------------------------
      ' Can only be Identified Record if we're doing :
      '      1) INSERT into a history of an Identified Record's table (or ascendant)
      '   or 2) DELETE/UPDATE a Identified Record table (or ascendant)
      ' --------------------------------------------------------------------------------------
      fRecSelOK = False
      fTableIsInRecSel = False
      fParentTableIsInRecSel = False
      
      sTableIDs = "0"
      If UBound(maWFPrecedingElements) > 1 Then
        ' Add  an item to the combo for each preceding web form.
        ' Ignore the first item, as it will be the current web form.
        For iLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
          If maWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
            ' Add  an item to the combo for each input item in the preceding web form.
            Set wfTemp = maWFPrecedingElements(iLoop)
            asItems = wfTemp.Items
            For iLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                ' Get an array of the valid table IDs (base table and it's ascendants)
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                                
                For lngLoop = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop) = lngTableID Then
                    fTableIsInRecSel = True
                  End If
                  sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
                Next lngLoop
              End If
            Next iLoop2
            
            Set wfTemp = Nothing
          ElseIf maWFPrecedingElements(iLoop).ElementType = elem_StoredData Then
            ReDim alngValidTables(0)
            TableAscendants maWFPrecedingElements(iLoop).DataTableID, alngValidTables
                            
            'JPD 20061227
            If maWFPrecedingElements(iLoop).DataAction = DATAACTION_DELETE Then
              ' Cannot do anything with a Deleted record, but can use its ascendants.
              ' Remove the table itself from the array of valid tables.
              alngValidTables(1) = 0
            End If
            
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = lngTableID Then
                fTableIsInRecSel = True
              End If
              sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
            Next lngLoop
          End If
        Next iLoop
      End If
      
      sSQL = "SELECT COUNT(*) AS [result]" & _
        " FROM tmpRelations" & _
        " WHERE tmpRelations.parentID IN(" & sTableIDs & ")" & _
        " AND tmpRelations.childID = " & CStr(lngTableID)
      
      If (piIndex = STOREDDATATABLE_SECONDARY) Then
        sSQL = sSQL & _
          " AND tmpRelations.parentID <> " & CStr(lngFirstDataRecordTable)
      End If
      
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      If rsTables!result > 0 Then
        fParentTableIsInRecSel = True
      End If

      rsTables.Close
      Set rsTables = Nothing
      
      Select Case iAction
        Case DATAACTION_INSERT:
          fRecSelOK = fParentTableIsInRecSel

        Case DATAACTION_UPDATE, DATAACTION_DELETE:
          fRecSelOK = fTableIsInRecSel
      End Select
      
      If fRecSelOK Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
        .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
      
        If iRecordSelection = giWFRECSEL_IDENTIFIEDRECORD Then
          iIndex = .NewIndex
        End If
      End If
      
      ' --------------------------------------------------------------------------------------
      ' Can only be Initiator's record if we're doing :
      '      1) INSERT into a history of the Personnel table (or ascendant)
      '   or 2) DELETE/UPDATE the Personnel table (or ascendant)
      ' NB. Must be MANUAL initiation
      ' --------------------------------------------------------------------------------------
      fRecSelOK = (miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL)
      
      If fRecSelOK Then
        ' Get an array of the valid table IDs (base table and it's ascendants)
        ReDim alngValidTables(0)
        TableAscendants mlngPersonnelTableID, alngValidTables
        
        Select Case iAction
          Case DATAACTION_INSERT:
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                If (piIndex = STOREDDATATABLE_SECONDARY) Then
                  If alngValidTables(lngLoop) <> lngFirstDataRecordTable Then
                    fFound = True
                    Exit For
                  End If
                Else
                  fFound = True
                  Exit For
                End If
              End If
            Next lngLoop
        
            fRecSelOK = fFound
          
          Case DATAACTION_UPDATE, DATAACTION_DELETE:
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop

            fRecSelOK = fFound
        End Select
      End If
      
      If fRecSelOK Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
        .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
      
        If iRecordSelection = giWFRECSEL_INITIATOR Then
          iIndex = .NewIndex
        End If
      End If
    
      ' --------------------------------------------------------------------------------------
      ' Can only be Base table's record if we're doing :
      '      1) INSERT into a history of the Base table (or ascendant)
      '   or 2) DELETE/UPDATE the Base table (or ascendant)
      ' NB. Must be TRIGGERED initiation
      ' --------------------------------------------------------------------------------------
      fRecSelOK = (miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED)

      If fRecSelOK Then
        ' Get an array of the valid table IDs (base table and it's ascendants)
        ReDim alngValidTables(0)
        TableAscendants mlngBaseTableID, alngValidTables
        
        Select Case iAction
          Case DATAACTION_INSERT:
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                If (piIndex = STOREDDATATABLE_SECONDARY) Then
                  If alngValidTables(lngLoop) <> lngFirstDataRecordTable Then
                    fFound = True
                    Exit For
                  End If
                Else
                  fFound = True
                  Exit For
                End If
              End If
            Next lngLoop
            
            fRecSelOK = fFound
  
          Case DATAACTION_UPDATE, DATAACTION_DELETE:
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop
            
            fRecSelOK = fFound
        End Select
      End If

      If fRecSelOK Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
        .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD

        If iRecordSelection = giWFRECSEL_TRIGGEREDRECORD Then
          iIndex = .NewIndex
        End If
      End If
    
      ' --------------------------------------------------------------------------------------
      ' Can only be Unidentified Record if we're doing :
      '      1) INSERT into a top level or Lookup table
      '      2) The secondary record when INSERTing into a shared history table
      ' --------------------------------------------------------------------------------------
      iTableType = 0
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", lngTableID
        
        If Not .NoMatch Then
          iTableType = !TableType
        End If
      End With
    
      If ((piIndex = STOREDDATATABLE_SECONDARY And fSecondaryRecordRequired)) _
        Or (((iTableType = iTabParent) Or (iTableType = iTabLookup)) _
          And (iAction = DATAACTION_INSERT)) Then
        
        .AddItem GetRecordSelectionDescription(giWFRECSEL_UNIDENTIFIED)
        .ItemData(.NewIndex) = giWFRECSEL_UNIDENTIFIED
      
        If iRecordSelection = giWFRECSEL_UNIDENTIFIED Then
          iIndex = .NewIndex
        End If
      End If
    End If
    
    ' Enable the combo if there are items.
    .Enabled = (.ListCount > 0)
    
    If (.ListCount > 0) Then
      If iIndex < 0 Then
        
        sMsg = "The previously selected " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " record is no longer valid." & vbCrLf
        sMsg = sMsg & "A default " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " record has been selected."
        
        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If
        
        iIndex = 0
        mfChanged = True
      End If
      
      .ListIndex = iIndex
    Else
      .AddItem IIf(piIndex = STOREDDATATABLE_SECONDARY And Not fSecondaryRecordRequired, "", "<No identifiable record for the specified action and table>")
      .ItemData(.NewIndex) = giWFRECSEL_UNKNOWN
      .ListIndex = .NewIndex
    End If
  End With
    
End Sub

Public Sub PrecedingElements(paWFPrecedingElements As Variant)

  ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
  Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mwfElement
  
  mfrmCallingForm.PrecedingElements mwfElement, paWFPrecedingElements
End Sub
Private Function DataRecordTable(piIndex As Integer) As Long
  ' Return the ID of the table referred to by the specified primary or secondary data record.
  If cboDataRecordTable(piIndex).ListCount > 0 Then
    DataRecordTable = cboDataRecordTable(piIndex).ItemData(cboDataRecordTable(piIndex).ListIndex)
  Else
    DataRecordTable = 0
  End If
End Function

Private Sub cboEmailElement_Refresh()
  ' Populate the Email WebForm combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim wfTemp As VB.Control
  Dim fElementWithRecord As Boolean
  Dim asItems() As String
  Dim fFound As Boolean
  Dim sMsg As String
  
  With cboEmailElement
    ' Clear the current contents of the combo.
    .Clear

    If cboEmailRecord.ItemData(cboEmailRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD Then
      For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
        fElementWithRecord = False
        Set wfTemp = maWFPrecedingElements(lngLoop)
        
        If wfTemp.ElementType = elem_WebForm Then
          asItems = wfTemp.Items
          
          For lngLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
              fElementWithRecord = True
              Exit For
            End If
          Next lngLoop2
        ElseIf wfTemp.ElementType = elem_StoredData Then
          'JPD 20061227
          'If wfTemp.DataAction <> DATAACTION_DELETE Then
            fElementWithRecord = True
          'Else
          '  ' DELETE storedData element are only allowed if they are for child tables
          '  ' ie. have parent record that may be used (as the records themselves will be deleted)
          '  With recTabEdit
          '    .Index = "idxTableID"
          '    .Seek "=", wfTemp.DataTableID
          '
          '    If Not recTabEdit.NoMatch Then
          '      If recTabEdit!TableType = iTabChild Then
          '        fElementWithRecord = True
          '      End If
          '    End If
          '  End With
          'End If
        End If
        
        If fElementWithRecord Then
          .AddItem wfTemp.Identifier
          .ItemData(.NewIndex) = wfTemp.ControlIndex
        End If
        
        Set wfTemp = Nothing
      Next lngLoop
    End If
    
    iIndex = 0
    fFound = (Len(Trim(mwfElement.RecordSelectorWebFormIdentifier)) = 0)
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = mwfElement.RecordSelectorWebFormIdentifier Then
        fFound = True
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop
    
    If Not fFound Then
      sMsg = "The previously selected Email Element is no longer valid."
      
      If .ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Email Element has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
              
      mfChanged = True
    End If
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboEmailElement_Click
    End If
  End With
    
End Sub

Private Sub cboEmailTable_Refresh()
  ' Populate the Email Table combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim fFound As Boolean
  Dim alngValidTables() As Long
  Dim lngBaseTable As Long
  Dim lngCurrentTable As Long
  Dim fTableOK As Boolean
  Dim lngExcludedTableID As Long
  
  lngCurrentTable = 0
  iIndex = -1
  iDefaultIndex = -1
  lngExcludedTableID = 0
  
  If mlngEmailID > 0 Then
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", mlngEmailID

      ' Read the expression's tableID from the recordset.
      If Not .NoMatch Then
        lngCurrentTable = !TableID
      End If
    End With
  End If
  If (lngCurrentTable = 0) And (mlngEmailCCID > 0) Then
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", mlngEmailCCID

      ' Read the expression's tableID from the recordset.
      If Not .NoMatch Then
        lngCurrentTable = !TableID
      End If
    End With
  End If
  
  Select Case cboEmailRecord.ItemData(cboEmailRecord.ListIndex)
    Case giWFRECSEL_INITIATOR
      lngBaseTable = mlngPersonnelTableID
      
    Case giWFRECSEL_TRIGGEREDRECORD
      lngBaseTable = mlngBaseTableID
      
    Case giWFRECSEL_IDENTIFIEDRECORD
      Set wfTemp = GetElementByIdentifier(cboEmailElement.List(cboEmailElement.ListIndex))
      
      If Not wfTemp Is Nothing Then
        If wfTemp.ElementType = elem_WebForm Then
          lngBaseTable = cboEmailRecordSelector.ItemData(cboEmailRecordSelector.ListIndex)
        ElseIf wfTemp.ElementType = elem_StoredData Then
          lngBaseTable = wfTemp.DataTableID
          
          'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
          'If wfTemp.DataAction = DATAACTION_DELETE Then
          '  ' Exclude deleted records (but include their parent records)
          '  lngExcludedTableID = wfTemp.DataTableID
          'End If
        End If
      End If
      
    Case Else
      lngBaseTable = 0
  End Select
  
  ' Get an array of the valid table IDs (base table and it's ascendants)
  ReDim alngValidTables(0)
  TableAscendants lngBaseTable, alngValidTables
  
  ' Clear the current contents of the combo.
  cboEmailTable.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combo for each table that has not been deleted.
    Do While Not .EOF
      fTableOK = (Not .Fields("deleted"))
      
      If fTableOK Then
        fFound = False
        
        For lngLoop = 1 To UBound(alngValidTables)
          If (alngValidTables(lngLoop) = !TableID) _
            And (lngExcludedTableID <> !TableID) Then
            fFound = True
            Exit For
          End If
        Next lngLoop
        
        fTableOK = fFound
      End If
      
      If fTableOK Then
        cboEmailTable.AddItem !TableName
        cboEmailTable.ItemData(cboEmailTable.NewIndex) = !TableID

        If !TableID = lngCurrentTable Then
          iIndex = cboEmailTable.NewIndex
        End If

        If !TableID = lngBaseTable Then
          iDefaultIndex = cboEmailTable.NewIndex
        End If
      End If

      .MoveNext
    Loop
  End With

  If cboEmailTable.ListCount > 0 Then
    If iIndex < 0 Then
      If iDefaultIndex < 0 Then
        iIndex = 0
      Else
        iIndex = iDefaultIndex
      End If
    End If
    
    cboEmailTable.ListIndex = iIndex
  Else
    cboEmailTable_Click
  End If
    
End Sub


Private Sub cboDataRecordTable_Refresh(piIndex As Integer)
  ' Populate the Data Record Table combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim fFound As Boolean
  Dim alngValidTables() As Long
  Dim lngBaseTable As Long
  Dim lngCurrentTable As Long
  Dim fTableOK As Boolean
  Dim lngFirstDataRecordTable As Long
  Dim lngTableID As Long
  Dim iAction As DataAction
  Dim fSecondaryRecordRequired As Boolean
  Dim lngExcludedTableID As Long
  
  lngCurrentTable = 0
  iIndex = -1
  iDefaultIndex = -1
  fSecondaryRecordRequired = SecondaryRecordRequired
  lngExcludedTableID = 0

  If (piIndex = STOREDDATATABLE_SECONDARY) Then
    lngFirstDataRecordTable = DataRecordTable(STOREDDATATABLE_PRIMARY)
  End If
  
  With cboDataRecordTable(piIndex)
    If .ListIndex < 0 Then
      lngCurrentTable = IIf(piIndex = STOREDDATATABLE_PRIMARY, mwfElement.DataRecordTableID, mwfElement.SecondaryDataRecordTableID)
    Else
      lngCurrentTable = .ItemData(.ListIndex)
    End If
  End With
  
  ' Clear the current contents of the combo.
  cboDataRecordTable(piIndex).Clear
  
  If (cboDataTable.ListIndex >= 0) _
    And (piIndex = STOREDDATATABLE_PRIMARY Or fSecondaryRecordRequired) Then
    
    lngTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
    iAction = cboDataAction.ItemData(cboDataAction.ListIndex)
  
    Select Case cboDataRecord(piIndex).ItemData(cboDataRecord(piIndex).ListIndex)
      Case giWFRECSEL_INITIATOR
        lngBaseTable = mlngPersonnelTableID
  
      Case giWFRECSEL_TRIGGEREDRECORD
        lngBaseTable = mlngBaseTableID
  
      Case giWFRECSEL_IDENTIFIEDRECORD
        Set wfTemp = GetElementByIdentifier(cboDataElement(piIndex).List(cboDataElement(piIndex).ListIndex))
  
        If Not wfTemp Is Nothing Then
          If wfTemp.ElementType = elem_WebForm Then
            lngBaseTable = cboDataRecordSelector(piIndex).ItemData(cboDataRecordSelector(piIndex).ListIndex)
          ElseIf wfTemp.ElementType = elem_StoredData Then
            lngBaseTable = wfTemp.DataTableID
            
            'JPD 20061227
            If wfTemp.DataAction = DATAACTION_DELETE Then
              ' Exclude deleted records (but include their parent records)
              lngExcludedTableID = wfTemp.DataTableID
            End If
          End If
        End If
  
      Case Else
        lngBaseTable = 0
    End Select
  
    ' Get an array of the valid table IDs (base table and it's ascendants)
    ReDim alngValidTables(0)
    TableAscendants lngBaseTable, alngValidTables
    
    With recTabEdit
      .Index = "idxName"
  
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If
  
      ' Add  an item to the combo for each table that has not been deleted.
      Do While Not .EOF
        fTableOK = (Not .Fields("deleted"))
  
        If fTableOK And (piIndex = STOREDDATATABLE_SECONDARY) Then
          fTableOK = (lngFirstDataRecordTable <> !TableID)
        End If
        
        If fTableOK Then
          fFound = False
  
          For lngLoop = 1 To UBound(alngValidTables)
            If (alngValidTables(lngLoop) = !TableID) _
              And (lngExcludedTableID <> !TableID) Then
              
              Select Case iAction
                Case DATAACTION_INSERT:
                  fFound = IsChildOfTable(alngValidTables(lngLoop), lngTableID)
                Case DATAACTION_DELETE, DATAACTION_UPDATE
                  fFound = (alngValidTables(lngLoop) = lngTableID)
              End Select
              Exit For
            End If
          Next lngLoop
  
          fTableOK = fFound
        End If
        
        If fTableOK Then
          cboDataRecordTable(piIndex).AddItem !TableName
          cboDataRecordTable(piIndex).ItemData(cboDataRecordTable(piIndex).NewIndex) = !TableID
  
          If !TableID = lngCurrentTable Then
            iIndex = cboDataRecordTable(piIndex).NewIndex
          End If
  
          If !TableID = lngBaseTable Then
            iDefaultIndex = cboDataRecordTable(piIndex).NewIndex
          End If
        End If
  
        .MoveNext
      Loop
    End With
  End If

  cboDataRecordTable(piIndex).Enabled = (cboDataRecordTable(piIndex).ListCount > 0)

  If cboDataRecordTable(piIndex).ListCount > 0 Then
    If iIndex < 0 Then
      If iDefaultIndex < 0 Then
        iIndex = 0
      Else
        iIndex = iDefaultIndex
      End If
    End If

    cboDataRecordTable(piIndex).ListIndex = iIndex
  Else
    cboDataRecordTable_Click piIndex
  End If
    
End Sub



Private Sub cboDataElement_Refresh(piIndex As Integer)
  ' Populate the DataElement combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim fElementWithRecord As Boolean
  Dim asItems() As String
  Dim lngTableID As Long
  Dim iAction As DataAction
  Dim iRecordSelection As WorkflowRecordSelectorTypes
  Dim sElementIdentifier As String
  Dim lngFirstDataRecordTable As Long
  Dim fFound As Boolean
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fSecondaryRecordRequired As Boolean
  
  ' Secondary record info only required if the data action is INSERTing into a shared history table.
  fSecondaryRecordRequired = SecondaryRecordRequired
  
  If (piIndex = STOREDDATATABLE_SECONDARY) Then
    lngFirstDataRecordTable = DataRecordTable(STOREDDATATABLE_PRIMARY)
  End If
  
  With cboDataElement(piIndex)
    If .ListIndex < 0 Then
      sElementIdentifier = IIf(piIndex = STOREDDATATABLE_PRIMARY, mwfElement.RecordSelectorWebFormIdentifier, mwfElement.SecondaryRecordSelectorWebFormIdentifier)
    Else
      sElementIdentifier = .List(.ListIndex)
    End If
    
    ' Clear the current contents of the combo.
    .Clear

    If cboDataTable.ListIndex >= 0 _
      And cboDataRecord(piIndex).ListIndex >= 0 Then
      
      lngTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
      iAction = cboDataAction.ItemData(cboDataAction.ListIndex)
      iRecordSelection = cboDataRecord(piIndex).ItemData(cboDataRecord(piIndex).ListIndex)
      
      If (iRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
        And (piIndex = STOREDDATATABLE_PRIMARY Or fSecondaryRecordRequired) Then
        
        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
          fElementWithRecord = False
          Set wfTemp = maWFPrecedingElements(lngLoop)

          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items

            For lngLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                ' Get an array of the valid table IDs (base table and it's ascendants)
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
                
                ' Check that the inputgrid is for a valid table for the selected action/table.
                Select Case iAction
                  Case DATAACTION_INSERT
                    ' WebForm RecordSelector (or ascendant) must be for a parent of the DataTable
                    fFound = False
                    For lngLoop3 = 1 To UBound(alngValidTables)
                      If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                        If (piIndex = STOREDDATATABLE_SECONDARY) Then
                          If alngValidTables(lngLoop3) <> lngFirstDataRecordTable Then
                            fFound = True
                            Exit For
                          End If
                        Else
                          fFound = True
                          Exit For
                        End If
                      End If
                    Next lngLoop3
                    
                    fElementWithRecord = fFound

                    If fElementWithRecord Then
                      Exit For
                    End If

                  Case DATAACTION_DELETE, DATAACTION_UPDATE
                    ' WebForm RecordSelector must be for the DataTable
                    fFound = False
                    For lngLoop3 = 1 To UBound(alngValidTables)
                      If alngValidTables(lngLoop3) = lngTableID Then
                        fFound = True
                        Exit For
                      End If
                    Next lngLoop3
                    
                    fElementWithRecord = fFound
                                            
                    If fElementWithRecord Then
                      Exit For
                    End If
                
                End Select
              End If
            Next lngLoop2
            
          ElseIf wfTemp.ElementType = elem_StoredData Then
            ' Get an array of the valid table IDs (base table and it's ascendants)
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables
            
            'JPD 20061227
            If wfTemp.DataAction = DATAACTION_DELETE Then
              alngValidTables(1) = 0
            End If
            
            Select Case iAction
              Case DATAACTION_INSERT
                ' StoredData record must be for a parent of the DataTable
                fFound = False
                For lngLoop3 = 1 To UBound(alngValidTables)
                  If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                    If (piIndex = STOREDDATATABLE_SECONDARY) Then
                      If alngValidTables(lngLoop3) <> lngFirstDataRecordTable Then
                        fFound = True
                        Exit For
                      End If
                    Else
                      fFound = True
                      Exit For
                    End If
                  End If
                Next lngLoop3
                                  
                fElementWithRecord = fFound
                
              Case DATAACTION_DELETE, DATAACTION_UPDATE
                ' StoredData record must be for the DataTable
                fFound = False
                For lngLoop3 = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop3) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop3
                                  
                fElementWithRecord = fFound
            End Select
          End If

          If fElementWithRecord Then
            .AddItem wfTemp.Identifier
            .ItemData(.NewIndex) = wfTemp.ControlIndex
          End If

          Set wfTemp = Nothing
        Next lngLoop
      End If
    End If

    ' Enable the combo if there are items.
    .Enabled = (.ListCount > 0)
       
    iIndex = -1
    fFound = (Len(Trim(sElementIdentifier)) = 0)
    
    If .ListCount > 0 Then
      For lngLoop = 0 To .ListCount - 1
        If UCase(Trim(sElementIdentifier)) = UCase(Trim(.List(lngLoop))) Then
          fFound = True
          iIndex = lngLoop
          Exit For
        End If
      Next lngLoop
      
      If iIndex < 0 Then
        iIndex = 0
      End If
      
      .ListIndex = iIndex
    Else
      cboDataElement_Click piIndex
    End If
  
    If Not fFound Then
      sMsg = "The previously selected " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " Record Element is no longer valid."

      If .ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " Record Element has been selected."
        'JPD 20070723 Fault 12405
        'Else
        '  If (piIndex = STOREDDATATABLE_PRIMARY) Then
        '    mwfElement.RecordSelectorWebFormIdentifier = ""
        '    mwfElement.RecordSelectorIdentifier = ""
        '  Else
        '    mwfElement.SecondaryRecordSelectorWebFormIdentifier = ""
        '    mwfElement.SecondaryRecordSelectorIdentifier = ""
        '  End If
      End If

      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
      
      mfChanged = True
    End If
  End With
    
End Sub


Private Sub cboDataRecordSelector_Refresh(piIndex As Integer)
  ' Populate the Data RecordSelector combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sRecordSelector As String
  Dim lngTableID As Long
  Dim iAction As DataAction
  Dim iRecordSelection As WorkflowRecordSelectorTypes
  Dim sElementIdentifier As String
  Dim fRecSelectorValid As Boolean
  Dim lngFirstDataRecordTable As Long
  Dim sMsg As String
  Dim fFound As Boolean
  Dim alngValidTables() As Long
  Dim fSecondaryRecordRequired As Boolean
  
  ' Secondary record info only required if the data action is INSERTing into a shared history table.
  fSecondaryRecordRequired = SecondaryRecordRequired
  
  If (piIndex = STOREDDATATABLE_SECONDARY) Then
    lngFirstDataRecordTable = DataRecordTable(STOREDDATATABLE_PRIMARY)
  End If
  
  With cboDataRecordSelector(piIndex)
    If .ListIndex < 0 Then
      sRecordSelector = IIf(piIndex = STOREDDATATABLE_PRIMARY, mwfElement.RecordSelectorIdentifier, mwfElement.SecondaryRecordSelectorIdentifier)
    Else
      sRecordSelector = .List(.ListIndex)
    End If

    ' Clear the current contents of the combo.
    .Clear

    If cboDataTable.ListIndex >= 0 _
      And cboDataRecord(piIndex).ListIndex >= 0 _
      And cboDataElement(piIndex).ListIndex >= 0 Then
      
      lngTableID = cboDataTable.ItemData(cboDataTable.ListIndex)
      iAction = cboDataAction.ItemData(cboDataAction.ListIndex)
      iRecordSelection = cboDataRecord(piIndex).ItemData(cboDataRecord(piIndex).ListIndex)
      sElementIdentifier = cboDataElement(piIndex).List(cboDataElement(piIndex).ListIndex)
      
      If (iRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
        And (piIndex = STOREDDATATABLE_PRIMARY Or fSecondaryRecordRequired) Then
        
        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
          Set wfTemp = maWFPrecedingElements(lngLoop)

          If wfTemp.ControlIndex = cboDataElement(piIndex).ItemData(cboDataElement(piIndex).ListIndex) Then
            If wfTemp.ElementType = elem_WebForm Then
              asItems = wfTemp.Items
  
              For lngLoop2 = 1 To UBound(asItems, 2)
                If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                  fRecSelectorValid = False
                  
                  ' Get an array of the valid table IDs (base table and it's ascendants)
                  ReDim alngValidTables(0)
                  TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables
                  
                  ' Check that the inputgrid is for a valid table for the selected action/table.
                  Select Case iAction
                    Case DATAACTION_INSERT
                      ' Record selectors must be for a parent of the DataTable
                      fFound = False
                      For lngLoop3 = 1 To UBound(alngValidTables)
                        If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                          If (piIndex = STOREDDATATABLE_SECONDARY) Then
                            If alngValidTables(lngLoop3) <> lngFirstDataRecordTable Then
                              fFound = True
                              Exit For
                            End If
                          Else
                            fFound = True
                            Exit For
                          End If
                        End If
                      Next lngLoop3
                      
                      fRecSelectorValid = fFound
  
                    Case DATAACTION_DELETE, DATAACTION_UPDATE
                      ' Record selectors must be for the DataTable
                      fFound = False
                      For lngLoop3 = 1 To UBound(alngValidTables)
                        If alngValidTables(lngLoop3) = lngTableID Then
                          fFound = True
                          Exit For
                        End If
                      Next lngLoop3
                                            
                      fRecSelectorValid = fFound
                      
                  End Select
                  
                  If fRecSelectorValid Then
                    'JPD 20061010 Fault 11355
                    '.AddItem asItems(1, lngLoop2)
                    .AddItem asItems(9, lngLoop2)
                    .ItemData(.NewIndex) = asItems(44, lngLoop2)
                  End If
                End If
              Next lngLoop2
            End If
            
            Exit For
          End If

          Set wfTemp = Nothing
        Next lngLoop
      End If
    End If

    ' Enable the combo if there are items.
    .Enabled = (.ListCount > 0)
    
    iIndex = -1
    fFound = (Len(Trim(sRecordSelector)) = 0)
    
    If .ListCount > 0 Then
      For lngLoop = 0 To .ListCount - 1
        If UCase(Trim(sRecordSelector)) = UCase(Trim(.List(lngLoop))) Then
          fFound = True
          iIndex = lngLoop
          Exit For
        End If
      Next lngLoop
      
      If iIndex < 0 Then
        iIndex = 0
      End If
      
      .ListIndex = iIndex
    Else
      cboDataRecordSelector_Click piIndex
    End If
  
    If Not fFound Then
      sMsg = "The previously selected " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " Record Selector is no longer valid."

      If .ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default " & IIf(piIndex = STOREDDATATABLE_PRIMARY, "Primary", "Secondary") & " Record Selector has been selected."
        'JPD 20070723 Fault 12405
        'Else
        '  If (piIndex = STOREDDATATABLE_PRIMARY) Then
        '    mwfElement.RecordSelectorIdentifier = ""
        '  Else
        '    mwfElement.SecondaryRecordSelectorIdentifier = ""
        '  End If
      End If

      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
        
      mfChanged = True
    End If
  End With
    
End Sub



Private Sub cboEmailRecordSelector_Refresh()
  ' Populate the Email RecordSelector combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sMsg As String
  Dim fFound As Boolean

  With cboEmailRecordSelector
    ' Clear the current contents of the combo.
    .Clear

    If cboEmailElement.ListIndex >= 0 Then
      For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
        Set wfTemp = maWFPrecedingElements(lngLoop)
  
        If wfTemp.ControlIndex = cboEmailElement.ItemData(cboEmailElement.ListIndex) Then
          If wfTemp.ElementType = elem_WebForm Then
          
            asItems = wfTemp.Items
    
            For lngLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                'JPD 20061010 Fault 11355
                '.AddItem asItems(1, lngLoop2)
                .AddItem asItems(9, lngLoop2)
                .ItemData(.NewIndex) = asItems(44, lngLoop2)
              End If
            Next lngLoop2
          End If
          
          Exit For
        End If
        
        Set wfTemp = Nothing
      Next lngLoop
    End If
    
    iIndex = 0
    fFound = (Len(Trim(mwfElement.RecordSelectorIdentifier)) = 0)
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = mwfElement.RecordSelectorIdentifier Then
        fFound = True
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If Not fFound Then
      sMsg = "The previously selected Email Record Selector is no longer valid."
      
      If .ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Email Record Selector has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
        
      mfChanged = True
    End If

    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboEmailRecordSelector_Click
    End If
  End With
    
End Sub

Private Sub cboEmailRecord_Refresh()
  ' Populate the EmailRecord combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim fElementWithRecord As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sMsg As String
  
  iIndex = -1

  With cboEmailRecord
    ' Clear the current contents of the combo.
    .Clear

    ' --------------------------------------------------------------------------------------
    ' Identified Record
    ' --------------------------------------------------------------------------------------
    For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
      fElementWithRecord = False
      Set wfTemp = maWFPrecedingElements(lngLoop)
        
      If wfTemp.ElementType = elem_WebForm Then
        asItems = wfTemp.Items
          
        For lngLoop2 = 1 To UBound(asItems, 2)
          If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
            fElementWithRecord = True
            Exit For
          End If
        Next lngLoop2
      ElseIf wfTemp.ElementType = elem_StoredData Then
        'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
        'If wfTemp.DataAction <> DATAACTION_DELETE Then
          fElementWithRecord = True
        'Else
        '  ' DELETE storedData element are only allowed if they are for child tables
        '  ' ie. have parent record that may be used (as the records themselves will be deleted)
        '  With recTabEdit
        '    .Index = "idxTableID"
        '    .Seek "=", wfTemp.DataTableID
        '
        '    If Not recTabEdit.NoMatch Then
        '      If recTabEdit!TableType = iTabChild Then
        '        fElementWithRecord = True
        '      End If
        '    End If
        '  End With
        'End If
      End If
        
      If fElementWithRecord Then
        Exit For
      End If
        
      Set wfTemp = Nothing
    Next lngLoop

    If fElementWithRecord Then
      .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
      .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
      
      If mwfElement.EmailRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        iIndex = .NewIndex
      End If
    End If
    
    ' --------------------------------------------------------------------------------------
    ' Initiator's Record
    ' --------------------------------------------------------------------------------------
    If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
      .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
      .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
      iDefaultIndex = .NewIndex
    
      If mwfElement.EmailRecord = giWFRECSEL_INITIATOR Then
        iIndex = .NewIndex
      End If
    End If
    
    ' --------------------------------------------------------------------------------------
    ' Triggered Record
    ' --------------------------------------------------------------------------------------
    If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
      .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
      .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
      iDefaultIndex = .NewIndex
    
      If mwfElement.EmailRecord = giWFRECSEL_TRIGGEREDRECORD Then
        iIndex = .NewIndex
      End If
    End If
    
    ' --------------------------------------------------------------------------------------
    ' Unidentified Record
    ' --------------------------------------------------------------------------------------
    .AddItem GetRecordSelectionDescription(giWFRECSEL_UNIDENTIFIED)
    .ItemData(.NewIndex) = giWFRECSEL_UNIDENTIFIED
  
    If mwfElement.EmailRecord = giWFRECSEL_UNIDENTIFIED Then
      iIndex = .NewIndex
    End If
    
    ' Enable the combo if there are items.
    .Enabled = True
    
    If iIndex < 0 Then
      sMsg = "The previously selected Email Record is no longer valid."
      
      If .ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Email Record has been selected."
      End If
      
      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If
        
      mfChanged = True
        
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  Dim fraTemp As Frame
  
  Const GRIDROWHEIGHT = 239
  
  ReDim mavIdentifierLog(6, 0)
  
  fraCaption.BorderStyle = vbBSNone
  fraButtons.BorderStyle = vbBSNone
  
  ssgrdItems.RowHeight = GRIDROWHEIGHT
  ssgrdColumns.RowHeight = GRIDROWHEIGHT
  
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged Then
      iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If iAnswer = vbYes Then
        Call cmdOK_Click
        If Me.Cancelled Then Cancel = 1
      ElseIf iAnswer = vbNo Then
        Me.Cancelled = True
      ElseIf iAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      Me.Cancelled = True
    End If
  End If

End Sub

Private Sub Form_Resize()
  
  Const FRAMEGAP = 100
  
  If Me.WindowState = vbNormal Then
    
    Select Case mwfElement.ElementType
      
      Case elem_Email
        If Me.Height < msngMinFormHeight Then
          Me.Height = msngMinFormHeight
        End If
        
        If Me.Width < msngMinFormWidth Then
          Me.Width = msngMinFormWidth
        End If
        
        With fraElement(0)
          .Width = Me.Width - .Left - iXFRAMEGAP - (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
          .Height = Me.Height - .Top - iYFRAMEGAP - fraButtons.Height - iXFRAMEGAP - (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
          
          fraButtons.Top = .Top + .Height + iYFRAMEGAP
          fraButtons.Left = .Left + .Width - fraButtons.Width
          
          cmdAddItem.Left = .Width - cmdAddItem.Width - iXGAP
          cmdEditItem.Left = cmdAddItem.Left
          cmdCopyItem.Left = cmdAddItem.Left
          cmdRemoveItem.Left = cmdAddItem.Left
          cmdRemoveAllItems.Left = cmdAddItem.Left
          cmdMoveItemDown.Left = cmdAddItem.Left
          cmdMoveItemDown.Top = .Height - cmdMoveItemDown.Height - iXGAP
          cmdMoveItemUp.Left = cmdAddItem.Left
          cmdMoveItemUp.Top = cmdMoveItemDown.Top - cmdMoveItemUp.Height - iYFRAMEGAP
          
          ssgrdItems.Height = .Height - ssgrdItems.Top - iXGAP
          ssgrdItems.Width = cmdAddItem.Left - ssgrdItems.Left - iXGAP
          
          cboEmailRecord.Width = (ssgrdItems.Width + ssgrdItems.Left) - cboEmailRecord.Left
          cboEmailElement.Width = cboEmailRecord.Width
          cboEmailRecordSelector.Width = cboEmailRecord.Width
          cboEmailTable.Width = cboEmailRecord.Width
          txtEmailSubject.Width = cboEmailRecord.Width
          txtEmail.Width = cboEmailRecord.Width - cmdEmail.Width
          cmdEmail.Left = txtEmail.Left + txtEmail.Width
          txtEmailCC.Width = cboEmailRecord.Width - cmdEmailCC.Width
          cmdEmailCC.Left = txtEmailCC.Left + txtEmailCC.Width
          
          txtAttachAttachment.Width = cboEmailRecord.Width - cmdAttachAttachment.Width - cmdAttachClear.Width
          cmdAttachAttachment.Left = txtAttachAttachment.Left + txtAttachAttachment.Width
          cmdAttachClear.Left = cmdAttachAttachment.Left + cmdAttachAttachment.Width
          
          txtCaption.Left = cboEmailRecord.Left
          txtCaption.Width = cboEmailRecord.Width
        End With

        ResizeGridColumns ssgrdItems
      
      Case elem_StoredData
        If Me.Height < msngMinFormHeight Then
          Me.Height = msngMinFormHeight
        End If
        
        If Me.Width < msngMinFormWidth Then
          Me.Width = msngMinFormWidth
        End If
        
        With fraElement(2)
          .Width = Me.Width - .Left - iXFRAMEGAP - (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
          .Height = Me.Height - .Top - iYFRAMEGAP - fraButtons.Height - iXFRAMEGAP - (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
          
          fraCaption.Width = .Width
          
          fraButtons.Top = .Top + .Height + iYFRAMEGAP
          fraButtons.Left = .Left + .Width - fraButtons.Width
          
          cmdDataAdd.Left = .Width - cmdDataAdd.Width - iXGAP
          cmdDataEdit.Left = cmdDataAdd.Left
          cmdDataRemove.Left = cmdDataAdd.Left
          cmdDataRemoveAll.Left = cmdDataAdd.Left

          ssgrdColumns.Height = .Height - ssgrdColumns.Top - iXGAP
          ssgrdColumns.Width = cmdDataAdd.Left - ssgrdColumns.Left - iXGAP

          fraDataRecord(STOREDDATATABLE_PRIMARY).Width = (fraElement(2).Width - (2 * fraDataRecord(STOREDDATATABLE_PRIMARY).Left) - FRAMEGAP) / 2
          fraDataRecord(STOREDDATATABLE_SECONDARY).Left = fraDataRecord(STOREDDATATABLE_PRIMARY).Left + fraDataRecord(STOREDDATATABLE_PRIMARY).Width + FRAMEGAP
          fraDataRecord(STOREDDATATABLE_SECONDARY).Width = fraDataRecord(STOREDDATATABLE_PRIMARY).Width
          
          cboDataRecord(STOREDDATATABLE_PRIMARY).Width = fraDataRecord(STOREDDATATABLE_PRIMARY).Width - cboDataRecord(STOREDDATATABLE_PRIMARY).Left - FRAMEGAP
          cboDataElement(STOREDDATATABLE_PRIMARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          cboDataRecordSelector(STOREDDATATABLE_PRIMARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          cboDataRecordTable(STOREDDATATABLE_PRIMARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          
          cboDataRecord(STOREDDATATABLE_SECONDARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          cboDataElement(STOREDDATATABLE_SECONDARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          cboDataRecordSelector(STOREDDATATABLE_SECONDARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
          cboDataRecordTable(STOREDDATATABLE_SECONDARY).Width = cboDataRecord(STOREDDATATABLE_PRIMARY).Width
                 
          txtCaption.Left = cboDataAction.Left
          txtCaption.Width = fraDataRecord(STOREDDATATABLE_PRIMARY).Left + fraDataRecord(STOREDDATATABLE_PRIMARY).Width - txtCaption.Left
          
          lblDataTable.Left = fraDataRecord(STOREDDATATABLE_SECONDARY).Left + lblDataRecord(STOREDDATATABLE_SECONDARY).Left
          cboDataTable.Left = lblDataTable.Left + 1050
          cboDataTable.Width = fraDataRecord(STOREDDATATABLE_SECONDARY).Left + fraDataRecord(STOREDDATATABLE_SECONDARY).Width - cboDataTable.Left

          cboDataAction.Width = txtCaption.Width
          
          lblIdentifier.Left = lblDataTable.Left
          txtIdentifier.Left = cboDataTable.Left
          txtIdentifier.Width = cboDataTable.Width
          
               
        End With
        
        ResizeGridColumns ssgrdColumns
    
      Case elem_Decision
        txtCaption.Left = cboDecisionCaption.Left
        txtCaption.Width = cboDecisionCaption.Width
        
        Me.Height = fraElement(1).Height + fraElement(1).Top + iYFRAMEGAP + fraButtons.Height + iXFRAMEGAP + (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
        Me.Width = fraElement(1).Width + fraElement(1).Left + iXFRAMEGAP + (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
    
      Case Else
        Me.Height = fraButtons.Top + fraButtons.Height + iXFRAMEGAP + (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
        Me.Width = fraButtons.Width + fraButtons.Left + iXFRAMEGAP + (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
    End Select
  End If
  
End Sub

Private Sub optDecisionFlowType_Click(Index As Integer)
  Changed = True
  DecisionFlowControls_Refresh
  
End Sub

Private Sub ssgrdColumns_DblClick()
  If Not mfReadOnly Then
    If ssgrdColumns.Rows > 0 Then
      cmdDataEdit_Click
    Else
      cmdDataAdd_Click
    End If
  End If

End Sub

Private Sub ssgrdColumns_GotFocus()
  'JPD 20070321 Fault 11871
  ssgrdColumns.Refresh

End Sub


Private Sub ssgrdColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshScreen
End Sub

Private Sub ssgrdColumns_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  If Not cmdDataEdit.Enabled Then
    RefreshScreen
  End If
End Sub

Private Sub ssgrdItems_DblClick()
  If Not mfReadOnly Then
    If ssgrdItems.Rows > 0 Then
      cmdEditItem_Click
    Else
      cmdAddItem_Click
    End If
  End If

End Sub

Private Sub ssgrdItems_GotFocus()
  'JPD 20070321 Fault 11871
  ssgrdItems.Refresh

End Sub


Private Sub ssgrdItems_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshScreen

End Sub

Private Sub ssgrdItems_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  If Not cmdEditItem.Enabled Then
    RefreshScreen
  End If

End Sub

Private Sub txtCaption_Change()
  Changed = True
  
End Sub

Private Sub txtCaption_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub

Public Property Let Changed(ByVal pfNewValue As Boolean)
  If Not mfLoading Then
    mfChanged = pfNewValue
    RefreshScreen
  End If
End Property
Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
End Property

Private Sub txtEmailSubject_Change()
  Changed = True

End Sub

Private Sub txtEmailSubject_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub


Private Sub txtIdentifier_Change()
  Changed = True

End Sub


Private Sub txtIdentifier_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub



