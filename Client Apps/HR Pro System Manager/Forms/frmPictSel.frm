VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPictSel 
   Caption         =   "Open Picture"
   ClientHeight    =   4935
   ClientLeft      =   1815
   ClientTop       =   1395
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1023
   Icon            =   "frmPictSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1755
      TabIndex        =   1
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3015
      TabIndex        =   2
      Top             =   4200
      Width           =   1200
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   7170
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "name"
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "type"
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "height"
         Object.Tag             =   ""
         Text            =   "Height"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "width"
         Object.Tag             =   ""
         Text            =   "Width"
         Object.Width           =   1058
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   4185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmPictSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mblnCancelled As Boolean

Private gfLoading As Boolean
Private glngPictureID As Long
Private giPictureType As PictureTypeConstants
Private msExcludedExtensions As String

Private Const MIN_FORM_HEIGHT = 3000
Private Const MIN_FORM_WIDTH = 3000

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get Loading() As Boolean
  ' Return the 'Loading' flag.
  Loading = gfLoading
End Property

Public Property Let Loading(IsLoading As Boolean)
  ' Set the 'Loading' flag.
  gfLoading = IsLoading
  
End Property

Public Property Get SelectedPicture() As Long
  ' Return the picture ID.
  SelectedPicture = glngPictureID
  
End Property

Public Property Let SelectedPicture(lngPictID As Long)
  ' Set the picture ID.
  glngPictureID = lngPictID
  
End Property

Private Sub cmdCancel_Click()
    
    Cancelled = True
    UnLoad Me

End Sub

Private Sub cmdOK_Click()

  Cancelled = False
  SelectPicture
  
End Sub

Private Sub Form_Activate()
  Dim blnExit As Boolean
  Dim strKey As String
  Dim strFileName As String
  Dim sExtension As String
  Dim iIndex As Integer
  Dim fGoodPicture As Boolean
  
  If Me.Loading Then
    
    With recPictEdit
      .Index = "idxName"
      If Not (.BOF And .EOF) Then
        Screen.MousePointer = vbHourglass
      
        .MoveFirst
      
        'gobjProgress.AviFile = ""
        gobjProgress.AVI = dbPicture
        gobjProgress.MainCaption = "Picture Manager"
        gobjProgress.Caption = "HR Pro - System Manager"
        gobjProgress.NumberOfBars = 1
        gobjProgress.Bar1MaxValue = .RecordCount
        gobjProgress.Bar1Caption = "Loading pictures..."
        gobjProgress.Time = True
        gobjProgress.Cancel = True
        gobjProgress.OpenProgress
        
        Do While Not .EOF And Not gobjProgress.Cancelled
          strKey = "I" & Trim(Str(.Fields("pictureID")))
        
          strFileName = ReadPicture
          
          fGoodPicture = True
          If Len(msExcludedExtensions) > 0 Then
            iIndex = InStrRev(.Fields("name"), ".")
            If iIndex > 0 Then
              sExtension = ";" & UCase(Mid(.Fields("name"), iIndex)) & ";"
              iIndex = InStr(msExcludedExtensions, sExtension)
              fGoodPicture = (iIndex = 0)
            End If
          End If
           
          If fGoodPicture Then
            ImageList2.ListImages.Add , strKey, LoadPicture(strFileName)
          End If
          
          Kill strFileName
        
           gobjProgress.UpdateProgress False
           
          .MoveNext
        Loop
        
        .MoveFirst
      
        blnExit = gobjProgress.Cancelled
      End If
    End With
  
    If blnExit Then
      gobjProgress.CloseProgress
      UnLoad Me
      Screen.MousePointer = vbNormal
    Else
      LoadViews
      Me.Loading = False
      
      If gobjProgress.Visible = True Then gobjProgress.CloseProgress
      Screen.MousePointer = vbNormal
    End If
  End If
  
End Sub

Private Sub Form_Load()
  Me.Loading = True
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
    
  RemoveIcon Me
    
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  If Me.WindowState = vbMinimized Then
    Exit Sub
  End If

  ListView1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 600
  
  cmdCancel.Top = ListView1.Height + 100
  cmdCancel.Left = ListView1.Left + ListView1.Width - cmdCancel.Width - 100
  
  cmdOk.Top = cmdCancel.Top
  cmdOk.Left = cmdCancel.Left - cmdOk.Width - 200
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Unhook Me.hWnd
  
End Sub

Private Sub ListView1_DblClick()
  SelectPicture
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SelectPicture
  End If
End Sub

Private Sub LoadViews()
  Dim sKey As String
  Dim ThisItem As ComctlLib.ListItem
  Dim ThisNode As ComctlLib.Node
  Dim sName As String
  Dim lngHeight As Long
  Dim lngWidth As Long
  Dim sExtension As String
  Dim iIndex As Integer
  Dim fGoodPicture As Boolean
  
  Screen.MousePointer = vbHourglass
      
  UI.LockWindow Me.hWnd
  
  ListView1.ListItems.Clear
  
  With recPictEdit
    
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
      
        ' Only display undeleted pictures.
        If Not .Fields("deleted") Then

          If (giPictureType = vbPicTypeNone) Or _
            (giPictureType = .Fields("PictureType")) Then

            fGoodPicture = True
            If Len(msExcludedExtensions) > 0 Then
              iIndex = InStrRev(.Fields("name"), ".")
              If iIndex > 0 Then
                sExtension = ";" & UCase(Mid(.Fields("name"), iIndex)) & ";"
                iIndex = InStr(msExcludedExtensions, sExtension)
                fGoodPicture = (iIndex = 0)
              End If
            End If
           
            If fGoodPicture Then
              sKey = "I" & Trim(Str(.Fields("pictureID")))
              sName = Trim(.Fields("name"))
                
              lngHeight = Me.ScaleY(ImageList2.ListImages(sKey).Picture.Height, vbHimetric, vbPixels)
              lngWidth = Me.ScaleX(ImageList2.ListImages(sKey).Picture.Width, vbHimetric, vbPixels)
                
              Set ThisItem = ListView1.ListItems.Add(, sKey, _
                sName, ImageList2.ListImages(sKey).Index)
              ThisItem.SubItems(1) = Choose(.Fields("PictureType"), "Bitmap", "Metafile", "Icon")
              ThisItem.SubItems(2) = Trim(Str(lngHeight))
              ThisItem.SubItems(3) = Trim(Str(lngWidth))
      
                
              If .Fields("pictureID") = glngPictureID Then
                ThisItem.Selected = True
              End If
            End If
          End If
        End If
        
        .MoveNext
      
      Loop
      
      .MoveFirst
    
    End If
    
  End With
  
  UI.UnlockWindow
      
  Screen.MousePointer = vbNormal
  
  If Not ListView1.SelectedItem Is Nothing Then
    ListView1.SelectedItem.EnsureVisible
  End If
End Sub

Private Function SelectPicture() As Boolean
  If Not ListView1.SelectedItem Is Nothing Then
    glngPictureID = Val(Mid(ListView1.SelectedItem.key, 2))
    SelectPicture = True
    
    UnLoad Me
  End If
End Function


Public Property Get PictureType() As PictureTypeConstants
  ' Return the current picture type.
  PictureType = giPictureType
  
End Property

Public Property Let PictureType(ByVal piNewValue As PictureTypeConstants)
  ' Set the current picture type.
  giPictureType = piNewValue
  
End Property

Public Property Get ExcludedExtensions() As String
  ExcludedExtensions = msExcludedExtensions
  
End Property

Public Property Let ExcludedExtensions(ByVal psNewValue As String)
  ' Excluded Extensions string is a semi-column delimited string
  ' of file extension NOT to be shown.
  ' eg. ".gif;.bmp"
  msExcludedExtensions = ";" & UCase(Replace(psNewValue, " ", "")) & ";"
  
End Property
