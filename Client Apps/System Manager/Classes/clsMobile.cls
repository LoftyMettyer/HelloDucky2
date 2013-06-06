VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''put in globals?
'Dim gobjMobileDefs As Mobile.collections.Definition


Public MobileID As Long
'Private mobjPictureLibrary As Mobile.Pictures

' Edit a mobile designer page
Public Sub Edit()

'  Dim obj As New MobileDesigner.DesignerForm
'  obj.EditDesign

End Sub

Private Sub EditMobileDesign(ByRef plngID As Long)
'
'  On Error GoTo ErrorTrap
'
''  Dim objDesigner As SystemFramework
'  Dim objDesigner As Mobile.Form1
'  Dim bOK As Boolean
'
'  Set objDesigner = New Mobile.Form1
'  LoadPictures
'  Set objDesigner.PictureLibrary = mobjPictureLibrary
'
'  ' Transfer the pictures
'
'
'
'  ' Transfer the object types
'
'  ' Show the designer
'  objDesigner.ID = plngID
'  objDesigner.Edit
'
'TidyUpAndExit:
'  Set objDesigner = Nothing
'  Exit Sub
'
'ErrorTrap:
'  GoTo TidyUpAndExit

End Sub


Private Sub LoadPictures()
'  Dim blnExit As Boolean
'  Dim strKey As String
'  Dim strFileName As String
'  Dim sExtension As String
'  Dim iIndex As Integer
'  Dim fGoodPicture As Boolean
'  Dim objPicture As Mobile.Picture
'  Dim objImage As Picture
'  Dim msExcludedExtensions As String
'
'  Set mobjPictureLibrary = New Mobile.Pictures
'
'  With recPictEdit
'    .Index = "idxName"
'    If Not (.BOF And .EOF) Then
'      Screen.MousePointer = vbHourglass
'
'      .MoveFirst
'
'      'gobjProgress.AviFile = ""
'      gobjProgress.AVI = dbPicture
'      gobjProgress.MainCaption = "Picture Manager"
'      gobjProgress.Caption = Application.Name
'      gobjProgress.NumberOfBars = 1
'      gobjProgress.Bar1MaxValue = .RecordCount
'      gobjProgress.Bar1Caption = "Loading pictures..."
'      gobjProgress.Time = True
'      gobjProgress.Cancel = True
'      gobjProgress.OpenProgress
'
'      Do While Not .EOF And Not gobjProgress.Cancelled
'        strKey = "I" & Trim(Str(.Fields("pictureID")))
'
'        strFileName = ReadPicture
'
'        fGoodPicture = True
'        If Len(msExcludedExtensions) > 0 Then
'          iIndex = InStrRev(.Fields("name"), ".")
'          If iIndex > 0 Then
'            sExtension = ";" & UCase(Mid(.Fields("name"), iIndex)) & ";"
'            iIndex = InStr(msExcludedExtensions, sExtension)
'            fGoodPicture = (iIndex = 0)
'          End If
'        End If
'
'        If fGoodPicture Then
'          Set objPicture = New Mobile.Picture
'          objPicture.ID = .Fields("pictureID").value
'          objPicture.FileName = strFileName
'
'          mobjPictureLibrary.Add objPicture
'        End If
'
'    '    Kill strFileName
'
'         gobjProgress.UpdateProgress False
'
'        .MoveNext
'      Loop
'
'      .MoveFirst
'
'      blnExit = gobjProgress.Cancelled
'    End If
'  End With
'
'  gobjProgress.CloseProgress
'  Screen.MousePointer = vbDefault

End Sub


Private Sub LoadPanels()

  ' Load panels from tbsys_mobilepanels (same for every group)





  ' for each security user group load individual items

'
'
'gobjMobileDefs.header.BackColor = vbYellow
'
'gobjMobileDefs.Main.BackColor = vbYellow
'
'gobjMobileDefs.footer.BackColor = vbYellow
'




End Sub


