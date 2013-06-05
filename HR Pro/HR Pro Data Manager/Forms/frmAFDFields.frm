VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAFDFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AFD Postcode Names & Numbers Software"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1007
   Icon            =   "frmAFDFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGrid 
      Caption         =   "Records Available :"
      Height          =   2375
      Left            =   150
      TabIndex        =   0
      Top             =   3800
      Width           =   5580
      Begin SSDataWidgets_B.SSDBGrid grdRecords 
         Height          =   1905
         Left            =   195
         TabIndex        =   19
         Top             =   300
         Width           =   5220
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   8
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
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   1
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   8
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "recno"
         Columns(0).Name =   "recno"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1773
         Columns(1).Caption=   "Forename"
         Columns(1).Name =   "forename"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1482
         Columns(2).Caption=   "Initial(s)"
         Columns(2).Name =   "initial"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1773
         Columns(3).Caption=   "Surname"
         Columns(3).Name =   "surname"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3704
         Columns(4).Caption=   "Street"
         Columns(4).Name =   "street"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   2566
         Columns(5).Caption=   "Town"
         Columns(5).Name =   "Town"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2461
         Columns(6).Caption=   "County"
         Columns(6).Name =   "County"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "Property"
         Columns(7).Name =   "Property"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   9199
         _ExtentY        =   3360
         _StockProps     =   79
         ForeColor       =   0
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
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3225
      TabIndex        =   20
      Top             =   6360
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   400
      Index           =   1
      Left            =   4530
      TabIndex        =   21
      Top             =   6360
      Width           =   1200
   End
   Begin VB.Frame fraIndividual 
      Caption         =   "Selected Record :"
      Height          =   3525
      Left            =   150
      TabIndex        =   32
      Top             =   105
      Width           =   5580
      Begin VB.TextBox txtCounty 
         Height          =   315
         Left            =   1400
         TabIndex        =   15
         Top             =   2605
         Width           =   2500
      End
      Begin VB.CheckBox chkCounty 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   16
         Top             =   2625
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtTelephone 
         Height          =   315
         Left            =   1400
         TabIndex        =   17
         Top             =   2920
         Width           =   2500
      End
      Begin VB.CheckBox chkTelephone 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   18
         Top             =   2940
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtTown 
         Height          =   315
         Left            =   1400
         TabIndex        =   13
         Top             =   2290
         Width           =   2500
      End
      Begin VB.CheckBox chkTown 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   14
         Top             =   2310
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtStreet 
         Height          =   315
         Left            =   1400
         TabIndex        =   9
         Top             =   1660
         Width           =   2500
      End
      Begin VB.CheckBox chkStreet 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   10
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtLocality 
         Height          =   315
         Left            =   1400
         TabIndex        =   11
         Top             =   1975
         Width           =   2500
      End
      Begin VB.CheckBox chkLocality 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   12
         Top             =   1995
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtSurname 
         Height          =   315
         Left            =   1400
         TabIndex        =   5
         Top             =   1030
         Width           =   2500
      End
      Begin VB.CheckBox chkSurname 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   6
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtProperty 
         Height          =   315
         Left            =   1400
         TabIndex        =   7
         Top             =   1345
         Width           =   2500
      End
      Begin VB.CheckBox chkProperty 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   8
         Top             =   1365
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtForename 
         Height          =   315
         Left            =   1400
         TabIndex        =   1
         Top             =   400
         Width           =   2500
      End
      Begin VB.CheckBox chkForename 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   2
         Top             =   420
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.TextBox txtInitials 
         Height          =   315
         Left            =   1400
         TabIndex        =   3
         Top             =   715
         Width           =   2500
      End
      Begin VB.CheckBox chkInitials 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   4
         Top             =   735
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "County :"
         Height          =   300
         Left            =   200
         TabIndex        =   41
         Top             =   2665
         Width           =   1125
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone :"
         Height          =   300
         Left            =   200
         TabIndex        =   40
         Top             =   2980
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Locality :"
         Height          =   300
         Left            =   200
         TabIndex        =   39
         Top             =   2035
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Street :"
         Height          =   300
         Left            =   200
         TabIndex        =   38
         Top             =   1720
         Width           =   1125
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Town :"
         Height          =   300
         Left            =   200
         TabIndex        =   37
         Top             =   2350
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname :"
         Height          =   300
         Left            =   200
         TabIndex        =   36
         Top             =   1090
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Property :"
         Height          =   300
         Left            =   200
         TabIndex        =   35
         Top             =   1405
         Width           =   1125
      End
      Begin VB.Label lblForename 
         BackStyle       =   0  'Transparent
         Caption         =   "Forename :"
         Height          =   300
         Left            =   200
         TabIndex        =   34
         Top             =   460
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial(s) :"
         Height          =   300
         Left            =   200
         TabIndex        =   33
         Top             =   775
         Width           =   1125
      End
   End
   Begin VB.Frame fraMerged 
      Caption         =   "Selected Record :"
      Height          =   3525
      Left            =   150
      TabIndex        =   42
      Top             =   100
      Visible         =   0   'False
      Width           =   5580
      Begin VB.CheckBox chkMergedInitials 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   25
         Top             =   735
         Value           =   1  'Checked
         Width           =   950
      End
      Begin VB.TextBox txtMergedInitials 
         Height          =   315
         Left            =   1380
         TabIndex        =   24
         Top             =   715
         Width           =   2500
      End
      Begin VB.CheckBox chkMergedForename 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   23
         Top             =   420
         Value           =   1  'Checked
         Width           =   950
      End
      Begin VB.TextBox txtMergedForename 
         Height          =   315
         Left            =   1380
         TabIndex        =   22
         Top             =   400
         Width           =   2500
      End
      Begin VB.CheckBox chkMergedSurname 
         Caption         =   "Include"
         Height          =   285
         Left            =   4300
         TabIndex        =   27
         Top             =   1050
         Value           =   1  'Checked
         Width           =   950
      End
      Begin VB.TextBox txtMergedSurname 
         Height          =   315
         Left            =   1380
         TabIndex        =   26
         Top             =   1030
         Width           =   2500
      End
      Begin VB.CheckBox chkMergedTelephone 
         Caption         =   "Include"
         Height          =   270
         Left            =   4300
         TabIndex        =   31
         Top             =   2940
         Value           =   1  'Checked
         Width           =   950
      End
      Begin VB.CheckBox chkMergedAddress 
         Caption         =   "Include"
         Height          =   345
         Left            =   4300
         TabIndex        =   29
         Top             =   1995
         Value           =   1  'Checked
         Width           =   950
      End
      Begin VB.TextBox txtMergedTelephone 
         Height          =   360
         Left            =   1380
         TabIndex        =   30
         Top             =   2920
         Width           =   2500
      End
      Begin VB.TextBox txtMergedAddress 
         Height          =   1575
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1345
         Width           =   2500
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial(s) :"
         Height          =   300
         Left            =   180
         TabIndex        =   47
         Top             =   775
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Forename :"
         Height          =   300
         Left            =   180
         TabIndex        =   46
         Top             =   460
         Width           =   1125
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname :"
         Height          =   300
         Left            =   180
         TabIndex        =   45
         Top             =   1090
         Width           =   1125
      End
      Begin VB.Label lblMergedTelephone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone :"
         Height          =   300
         Left            =   120
         TabIndex        =   44
         Top             =   2980
         Width           =   1020
      End
      Begin VB.Label lblMergedAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         Height          =   285
         Left            =   180
         TabIndex        =   43
         Top             =   1405
         Width           =   855
      End
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Note : You have amended the original data returned from Afd..."
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   165
      TabIndex        =   48
      Top             =   6345
      Visible         =   0   'False
      Width           =   2940
   End
End
Attribute VB_Name = "frmAFDFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Module level var to hold the postcode we will be searching for
Private strPostCode As String
'Module level var to store whether or not the field mappings are individual or merged
Private fIndividual As Boolean
'Module level var to store which form called the Afd routine
Private frmForm As frmRecEdit4 ' Form

Private mbQuickAddressMode As Boolean
Private mobjQAPostcodes() As HRProDataMgr.PostCode

Public Function InitialiseAFD(PostCode As String, fIndiv As Boolean, frmCallingForm As Form, FieldName As String) As Boolean

  Dim temp As String     ' Copy of the postcode, incase AFD detects an old postcode
  Dim RecNo As Long      ' Current record number from the return set
  Dim Result As Long     ' Status of the search
  Dim nnFlags As Long    ' Afd parameters

  ' Caption
  Select Case gfAFDEnabled
    Case AFD_NamesNumbers
      Me.Caption = "AFD Postcode Names & Numbers Software"
    Case AFD_PostCode
      Me.Caption = "AFD Postcode Software"
    Case AFD_PostCodeplus
      Me.Caption = "AFD Postcode Plus Software"
    Case Else
      Me.Caption = "Unknown Postcode Software"
  End Select

  ' Quick Address or AFD mode
  mbQuickAddressMode = False

  'Store the postcode entered from the control on the users recedit form into a
  'module level variable
  strPostCode = Trim(PostCode)
  
  'Leave if the postcode is blank
  If strPostCode = "" Then
    InitialiseAFD = False
    Exit Function
  End If
  
  'Store if the mapped fields are individual or not in a module level variable
  fIndividual = fIndiv
  
  'Set the calling form to a module level variable
  Set frmForm = frmCallingForm
  
  'Show the right frame depending on value of fIndividual
  If fIndividual Then
    fraIndividual.Visible = True
    fraMerged.Visible = False
  Else
    fraIndividual.Visible = False
    fraMerged.Visible = True
  End If

  'Which columns on the grid do we display?
  grdRecords.Columns("Property").Visible = False
  grdRecords.Columns("Town").Visible = Not (gfAFDEnabled = AFD_NamesNumbers)
  grdRecords.Columns("County").Visible = Not (gfAFDEnabled = AFD_NamesNumbers)
  grdRecords.Columns("forename").Visible = (gfAFDEnabled = AFD_NamesNumbers)
  grdRecords.Columns("surname").Visible = (gfAFDEnabled = AFD_NamesNumbers)
  grdRecords.Columns("initial").Visible = (gfAFDEnabled = AFD_NamesNumbers)
  
  'Clear the Results List
  grdRecords.RemoveAll
  
  'Call AFD GetPostcode routines
  Result = GetAFDPostcode(strPostCode, 1, RETURN_ALL_RECORDS)
  
 'If an error, report and exit
  If Result < 0 Then
    ShowError Result
    InitialiseAFD = False
    Exit Function
  End If

 'If there's no error but no records returned, report and exit
  If Result < 1 Then
    ShowError -6
    InitialiseAFD = False
    Exit Function
  End If

  'Test & report if Postcode Change has been detected
  'Temp$ = Trim$(oPostCode.PostcodeFrom)
  'If Temp$ <> "" Then
  '   MsgBox "Postcode: " + Temp$ + " has changed to " + oPostCode.PostCode, 64, "Postcode Change"
  'End If
  
  'Store the flags used in the grids tag property
  grdRecords.Tag = nnFlags&

  'Read all records for the postcode
  For RecNo& = 1 To Result&
    
    'I'll only re-read the DLL for records after the first one
     'If RecNo& > 1 Then
        Result = GetAFDPostcode(strPostCode, RecNo, RETURN_MATCHING_RECORDS)
     'End If

    'If we find a valid result
     If Result& > 0 Then
        'Place it in the records grid
         AddToGrid RecNo&
     End If
   
   Next RecNo&

    If grdRecords.Rows > 0 Then
     'Select the first result
      grdRecords.MoveFirst
      grdRecords.SelBookmarks.Add grdRecords.Bookmark
      grdRecords_SelChange 2, False, True
      'Change width of the last column depending if scrollbar is required or not
      If grdRecords.Rows > 6 Then grdRecords.Columns(3).Width = 1500 Else grdRecords.Columns(3).Width = 1725
      InitialiseAFD = True
    Else
     'Report no Sucess
      MsgBox "Lookup Unsuccessful", 16
      InitialiseAFD = False
    End If
 
End Function

Private Sub cmdAction_Click(Index As Integer)

  Dim objControl As Control
  
  'Let user know somethings happening
  Screen.MousePointer = vbHourglass
  
  'setup the temp column name variables...these store the field names to put the
  'Afd data in.
  Dim tempforename As String
  Dim tempinitials As String
  Dim tempsurname As String
  Dim tempaddress As String
  Dim tempproperty As String
  Dim tempstreet As String
  Dim templocality As String
  Dim temptown As String
  Dim tempcounty As String
  Dim temptelephone As String
  
  'See which command button was pressed and take appropriate action
  Select Case Index
  
    Case 0    ' ok
    
      If Not fIndividual Then ' merged address fields
      
        'For all fields that are mapped correctly, store the field names
        If txtMergedForename.Tag <> 0 Then tempforename = datGeneral.GetColumnName(txtMergedForename.Tag)
        If txtMergedInitials.Tag <> 0 Then tempinitials = datGeneral.GetColumnName(txtMergedInitials.Tag)
        If txtMergedSurname.Tag <> 0 Then tempsurname = datGeneral.GetColumnName(txtMergedSurname.Tag)
        If txtMergedAddress.Tag <> 0 Then tempaddress = datGeneral.GetColumnName(txtMergedAddress.Tag)
        If txtMergedTelephone.Tag <> 0 Then temptelephone = datGeneral.GetColumnName(txtMergedTelephone.Tag)
      
        For Each objControl In frmForm.Controls
      
          'Loop through controls on the user form and copy the text accross if
          'the checkbox is checked.
          If objControl.Tag > 0 And IsNumeric(objControl.Tag) Then
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempforename And chkMergedForename.Value Then objControl.Text = txtMergedForename.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempinitials And chkMergedInitials.Value Then objControl.Text = txtMergedInitials.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempsurname And chkMergedSurname.Value Then objControl.Text = txtMergedSurname.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempaddress And chkMergedAddress.Value Then objControl.Text = txtMergedAddress.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = temptelephone And chkMergedTelephone.Value Then objControl.Text = txtMergedTelephone.Text
            End If
          
        Next objControl
        
      Else ' individual address fields
        
        'For all fields that are mapped correctly, store the field names
        If txtForename.Tag <> 0 Then tempforename = datGeneral.GetColumnName(txtForename.Tag)
        If txtInitials.Tag <> 0 Then tempinitials = datGeneral.GetColumnName(txtInitials.Tag)
        If txtSurname.Tag <> 0 Then tempsurname = datGeneral.GetColumnName(txtSurname.Tag)
        If txtProperty.Tag <> 0 Then tempproperty = datGeneral.GetColumnName(txtProperty.Tag)
        If txtStreet.Tag <> 0 Then tempstreet = datGeneral.GetColumnName(txtStreet.Tag)
        If txtLocality.Tag <> 0 Then templocality = datGeneral.GetColumnName(txtLocality.Tag)
        If txtTown.Tag <> 0 Then temptown = datGeneral.GetColumnName(txtTown.Tag)
        If txtCounty.Tag <> 0 Then tempcounty = datGeneral.GetColumnName(txtCounty.Tag)
        If txtTelephone.Tag <> 0 Then temptelephone = datGeneral.GetColumnName(txtTelephone.Tag)
        
        For Each objControl In frmForm.Controls
          
          'Loop through controls on the user form and copy the text accross if
          'the checkbox is checked.
          
          If objControl.Tag > 0 And IsNumeric(objControl.Tag) Then
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempforename And chkForename.Value Then objControl.Text = txtForename.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempinitials And chkInitials.Value Then objControl.Text = txtInitials.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempsurname And chkSurname.Value Then objControl.Text = txtSurname.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempproperty And chkProperty.Value Then objControl.Text = txtProperty.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempstreet And chkStreet.Value Then objControl.Text = txtStreet.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = templocality And chkLocality.Value Then objControl.Text = txtLocality.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = temptown And chkTown.Value Then objControl.Text = txtTown.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempcounty And chkCounty.Value Then objControl.Text = txtCounty.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = temptelephone And chkTelephone.Value Then objControl.Text = txtTelephone.Text
          End If
          
        Next objControl
        
      End If
      
  End Select
  
  'Unload the Afd screen
  Unload Me
  
  'Return mousepointer to normal
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub AddToGrid(RecNo As Long)
   
  Dim ShowString As String          ' String to add to the grid
  
  'Load the string with Name information

  'if the forename and surname are blank, then it could be a business, so show the
  'business name in the forename field
  
  If Trim(oPostCode.FirstName) = "" And Trim(oPostCode.Surname) = "" Then
      ShowString = Str(RecNo) & vbTab & Trim(oPostCode.Organisation) & vbTab & vbTab & vbTab
    Else
      ShowString = Str(RecNo) & vbTab & Trim(oPostCode.FirstName) & vbTab & Trim(oPostCode.Initial2) & vbTab & _
      Trim(oPostCode.Surname) & vbTab
  End If
    
  'Concatenate the address string depending on values of the afd returned data
  'for Building, HouseNo and Street
  
  If Trim(oPostCode.HouseNo) = "" Then
    If Trim(oPostCode.Building) = "" Then
      ShowString = ShowString & Trim(oPostCode.Street)
    Else
      ShowString = ShowString & Trim(oPostCode.Building) & " " & Trim(oPostCode.Street)
    End If
  Else
    ShowString = ShowString & Trim(oPostCode.HouseNo) & " " & Trim(oPostCode.Street)
  End If
    
  ' Add the town and county
  ShowString = ShowString & vbTab & Trim(oPostCode.Town) & vbTab & Trim(oPostCode.County)
    
  'Add the string to the grid
  grdRecords.AddItem ShowString
     
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub grdRecords_DblClick()

  'Doubleclicking the grid has the same result as clicking the OK button
  cmdAction_Click 0

End Sub

Private Sub grdRecords_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  Dim lngRecNo As Long     ' Current record number
  Dim lngResult As Long    ' Status of the search

  'Find the Record Number stored with the selected item (hidden column)
  lngRecNo = grdRecords.Columns(0).CellText(grdRecords.Bookmark)
  
  'Now call off the postcode record
  If mbQuickAddressMode Then
    oPostCode = mobjQAPostcodes(lngRecNo)
  Else
    lngResult = GetAFDPostcode(strPostCode, lngRecNo, Val(grdRecords.Tag))
  
    'If there was an error, report and exit
    If lngResult < 0 Then
      ShowError lngResult
      Exit Sub
    End If
  End If

 'Load the text fields with the data
  DisplayAddress

  'Reset the warning label
  lblWarning.Visible = False
  
End Sub

Private Sub DisplayAddress()

On Error GoTo DisplayAddressErr

  If fIndividual Then ' Address fields are individual
    
    'Empty textboxes
    ClearIndividual
    
    'Load the textboxes with the data returned from Afd
        
    If Trim(oPostCode.FirstName) = "" And Trim(oPostCode.Surname) = "" Then
      txtForename = Trim(oPostCode.Organisation)
    Else
      txtForename = Trim(oPostCode.FirstName)
      txtInitials = Trim(oPostCode.Initial2)
      txtSurname = Trim(oPostCode.Surname)
    End If
    
    txtProperty = Trim(oPostCode.Building)
    If Trim(oPostCode.HouseNo) = "" Then
      txtStreet = Trim(oPostCode.Street)
    Else
      txtStreet = Trim(oPostCode.HouseNo) & " " & Trim(oPostCode.Street)
    End If
    txtLocality = Trim(oPostCode.Locality)
    txtTown = Trim(oPostCode.Town)
    txtCounty = Trim(oPostCode.County)
    txtTelephone = MakePhoneNo$(oPostCode.Phone)
  
  Else ' Address fields are merged
     
    'Empty textboxes
    ClearMerged
    
    Dim address As String
    
    If Trim(oPostCode.FirstName) = "" And Trim(oPostCode.Surname) = "" Then
      txtMergedForename = Trim(oPostCode.Organisation)
    Else
      txtMergedForename = Trim(oPostCode.FirstName)
      txtMergedInitials = Trim(oPostCode.Initial2)
      txtMergedSurname = Trim(oPostCode.Surname)
    End If
     
    'Concatenate the address data
    If Trim(oPostCode.Building) <> "" Then address = Trim(oPostCode.Building) & vbCrLf
    If Trim(oPostCode.HouseNo) = "" Then
      address = address & Trim(oPostCode.Street) & vbCrLf
    Else
      address = address & Trim(oPostCode.HouseNo) & " " & Trim(oPostCode.Street) & vbCrLf
    End If
    If Trim(oPostCode.Locality) <> "" Then address = address & Trim(oPostCode.Locality) & vbCrLf
    If Trim(oPostCode.Town) <> "" Then address = address & Trim(oPostCode.Town) & vbCrLf
    If Trim(oPostCode.County) <> "" Then address = address & Trim(oPostCode.County) & vbCrLf
    If Trim(oPostCode.PostCode) <> "" Then address = address & Trim(oPostCode.PostCode) & vbCrLf
    
    txtMergedAddress = address
    txtMergedTelephone = MakePhoneNo$(oPostCode.Phone)
  
  End If

Exit Sub

DisplayAddressErr:

MsgBox "Warning...An Error Has Occurred.  Could not display data from Afd Correctly", vbExclamation + vbOKOnly, "PostCode Software"

End Sub

Private Function MakePhoneNo$(i$)

  'Routine to try and return a telephone number into the most common format
  
  On Error GoTo MakePhoneNoErr
  
    Select Case Left$(i$, 4)
      Case Is > "02"
        '#RH 14/09/99 - Amended for new telephone code compatibility
        'MakePhoneNo$ = Trim$(Left$(i$, 4) + " " + Mid$(i$, 5, 10))
        MakePhoneNo$ = Trim$(Left$(i$, 3) + " " + Mid$(i$, 4, 4)) + " " + Mid$(i$, 8)
      Case "011" To "0119"
        MakePhoneNo$ = Trim$(Left$(i$, 4) + " " + Mid$(i$, 5, 3) + " " + Mid$(i$, 8))
      Case "0121", "0131", "0141", "0151", "0161", "0171", "0181", "0191"
        MakePhoneNo$ = Trim$(Left$(i$, 4) + " " + Mid$(i$, 5, 3) + " " + Mid$(i$, 8))
      Case Else
        MakePhoneNo$ = Trim$(Left$(i$, 5) + " " + Mid$(i$, 6, 10))
    End Select
  
  Exit Function
  
MakePhoneNoErr:
  
  MakePhoneNo$ = Trim$(i$)
  Resume Next

End Function


Private Sub ClearIndividual()
      
  txtForename = ""
  txtInitials = ""
  txtSurname = ""
  txtProperty = ""
  txtStreet = ""
  txtLocality = ""
  txtTown = ""
  txtCounty = ""
  txtTelephone = ""

End Sub

Private Sub ClearMerged()

  txtMergedForename = ""
  txtMergedInitials = ""
  txtMergedSurname = ""
  txtMergedAddress = ""
  txtMergedTelephone = ""

End Sub

' The following subs display the 'Data has changed' note on the form if the user
' amends any of the data returned by Afd before copying across to their database.

Private Sub txtCounty_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtForename_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtInitials_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtLocality_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtMergedAddress_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtMergedForename_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtMergedInitials_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtMergedSurname_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtMergedTelephone_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtProperty_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtStreet_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtSurname_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtTelephone_Change()
  lblWarning.Visible = True
End Sub

Private Sub txtTown_Change()
  lblWarning.Visible = True
End Sub

Public Function InitialiseQA(PostCode As String, fIndiv As Boolean, frmCallingForm As Form, FieldName As String) As Boolean

  Dim Result As Long     ' Status of the search
  Dim iNoItems As Integer
  Dim lngCount As Long
  Dim lListItemReturn As Long
  Dim rsBuffer As String * 200

  ' Quick Address or AFD mode
  mbQuickAddressMode = True

  ' Caption
  Select Case giQAddressEnabled
    Case QADDRESS_RAPID
      Me.Caption = "QAS Quick Address Rapid Software"
    Case QADDRESS_PRO3
      Me.Caption = "QAS Quick Address Pro Software"
    Case QADDRESS_WORLDWIDE
      Me.Caption = "QAS Quick Address World Wide Software"
    Case QADDRESS_PRO4
      Me.Caption = "QAS Quick Address Pro Software"
    Case Else
      Me.Caption = "Unknown Postcode Software"
  End Select

  'Store the postcode entered from the control on the users recedit form into a
  'module level variable
  strPostCode = Trim(PostCode)
  
  'Leave if the postcode is blank
  If strPostCode = "" Then
    InitialiseQA = False
    Exit Function
  End If
  
  'Store if the mapped fields are individual or not in a module level variable
  fIndividual = fIndiv
  
  'Set the calling form to a module level variable
  Set frmForm = frmCallingForm
  
  'Show the right frame depending on value of fIndividual
  If fIndividual Then
    fraIndividual.Visible = True
    fraMerged.Visible = False
  Else
    fraIndividual.Visible = False
    fraMerged.Visible = True
  End If

  'Which columns on the grid do we display?
  grdRecords.Columns("Property").Visible = Not (giQAddressEnabled = QADDRESS_RAPID)
  grdRecords.Columns("Town").Visible = True
  grdRecords.Columns("County").Visible = True
  grdRecords.Columns("forename").Visible = False
  grdRecords.Columns("surname").Visible = False
  grdRecords.Columns("initial").Visible = False
  
  'Clear the Results List
  grdRecords.RemoveAll
  
  Result = QAddressGetPostcodes(strPostCode, mobjQAPostcodes)
  
 'If an error, report and exit
  If Result < 0 Then
    ShowError Result
    InitialiseQA = False
    Exit Function
  End If

  'If we find a valid result
  If Result > 0 Then
    'Place it in the records grid
    For lngCount = 1 To Result
      oPostCode = mobjQAPostcodes(lngCount)
      AddToGrid lngCount
    Next lngCount
  End If
   

    If grdRecords.Rows > 0 Then
     'Select the first result
      grdRecords.MoveFirst
      grdRecords.SelBookmarks.Add grdRecords.Bookmark
      grdRecords_SelChange 2, False, True
      'Change width of the last column depending if scrollbar is required or not
      If grdRecords.Rows > 6 Then grdRecords.Columns(3).Width = 1500 Else grdRecords.Columns(3).Width = 1725
      InitialiseQA = True
    Else
     'Report no Sucess
      MsgBox "Lookup Unsuccessful", 16
      InitialiseQA = False
    End If
 
End Function

