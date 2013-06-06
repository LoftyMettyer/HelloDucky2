Attribute VB_Name = "modPictMgr"
Option Explicit

'Common dialog fileopen constants
Const cdlOFNPathMustExist = 2048
Const cdlOFNFileMustExist = 4096
Const cdlOFNExplorer = 524288

Public Function GetPictureType(ThisPicture As Object) As String
  Dim strType As String

  Select Case ThisPicture.Type
    Case vbPicTypeNone
      strType = "None"
    Case vbPicTypeBitmap
      strType = "Bitmap"
    Case vbPicTypeMetafile
      strType = "Metafile"
    Case vbPicTypeIcon
      strType = "Icon"
  End Select

  GetPictureType = strType
End Function

Public Function GetTmpFName() As String
  Dim strTmpPath As String, strTmpName As String
  
  strTmpPath = Space(1024)
  strTmpName = Space(1024)

  Call GetTempPath(1024, strTmpPath)
  Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)
  
  strTmpName = Trim(strTmpName)
  If Len(strTmpName) > 0 Then
    strTmpName = Left(strTmpName, Len(strTmpName) - 1)
  Else
    strTmpName = vbNullString
  End If
  
  GetTmpFName = Trim(strTmpName)
End Function

Public Function JustFileName(ByVal FilePath As String) As String
  Dim strFileName As String
  Dim i As Integer
    
  strFileName = Trim(FilePath)
  i = InStr(strFileName, "\")
  Do While i > 0
    strFileName = Mid(strFileName, i + 1)
    i = InStr(strFileName, "\")
  Loop
  
  JustFileName = strFileName
End Function


Public Function ReadPicture() As String
  Dim strTempName As String
  Dim Fl As Long
  Dim Chunks As Integer
  Dim Fragment As Integer
  Dim Chunk() As Byte
  Dim i As Integer
  Dim TempFile As Integer

  If Not recPictEdit Is Nothing Then
  
    With recPictEdit
      
      If Not (.BOF And .EOF) Then
        strTempName = GetTmpFName
        TempFile = 1
        Open strTempName For Binary Access Write As TempFile
        
        Fl = .Fields("Picture").FieldSize
        Chunks = Fl / ChunkSize
        Fragment = Fl Mod ChunkSize
        
        ReDim Chunk(Fragment)
        Chunk() = .Fields("Picture").GetChunk(0, Fragment)
        Put TempFile, , Chunk()
        
        For i = 1 To Chunks
          ReDim Chunk(ChunkSize)
          Chunk() = .Fields("Picture").GetChunk((ChunkSize * (i - 1)) + Fragment, ChunkSize)
          Put TempFile, , Chunk()
        Next i
        Close TempFile
        
        ReadPicture = strTempName
      End If
    End With
  End If
  
End Function
Public Sub SizeImage(ThisImage As Image)
  Dim sglXSize As Single, sglYSize As Single
  Dim dblXRatio As Double, dblYRatio As Double
          
  With ThisImage
    sglXSize = ThisImage.Parent.ScaleX(.Picture.Width, vbHimetric, ThisImage.Parent.ScaleMode)
    sglYSize = ThisImage.Parent.ScaleY(.Picture.Height, vbHimetric, ThisImage.Parent.ScaleMode)
    If sglXSize <= .Width And sglYSize <= .Height Then
      .Height = sglYSize
      .Width = sglXSize
    Else
      If .Picture.Height > .Picture.Width Then
        dblYRatio = 1
        dblXRatio = .Picture.Width / .Picture.Height
      Else
        dblYRatio = .Picture.Height / .Picture.Width
        dblXRatio = 1
      End If
        
      .Height = .Height * dblYRatio
      .Width = .Width * dblXRatio
    End If
  End With
End Sub

Public Function WritePicture(FileName As String) As Boolean
  'On Error GoTo ErrorTrap
  
  Dim Fl As Long
  Dim Chunks As Integer
  Dim Fragment As Integer
  Dim Chunk() As Byte
  Dim i As Integer
  Dim TempFile As Integer
  
  If Not recPictEdit Is Nothing Then
    With recPictEdit
      If .EditMode <> dbEditNone Then
        TempFile = 1
        Open FileName For Binary Access Read As TempFile
        Fl = LOF(TempFile)
        If Fl > 0 Then
          Chunks = Fl / ChunkSize

          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get TempFile, , Chunk()
          .Fields("Picture").AppendChunk Chunk()
          
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
            Get TempFile, , Chunk()
            .Fields("Picture").AppendChunk Chunk()
          Next i
          .Update
          
          WritePicture = True
        Else
          .Cancel
          
          WritePicture = False
        End If
        Close TempFile
      End If
    End With
  End If
  
End Function


Public Sub SetDateComboFormat(cboDate As GTMaskDate.GTMaskDate)

  Dim sDateFormat As String
  
  sDateFormat = DateFormat
  
  cboDate.Format = sDateFormat
  cboDate.DisplayFormat = sDateFormat
  
  sDateFormat = Replace(sDateFormat, "d", "_")
  sDateFormat = Replace(sDateFormat, "m", "_")
  sDateFormat = Replace(sDateFormat, "y", "_")
  
  cboDate.Text = sDateFormat

End Sub

Public Function DateFormat() As String
  ' Returns the date format.
  ' NB. Windows allows the user to configure totally stupid
  ' date formats (eg. d/M/yyMydy !). This function does not cater
  ' for such stupidity, and simply takes the first occurence of the
  ' 'd', 'M', 'y' characters.
  Dim sSysFormat As String
  Dim sSysDateSeparator As String
  Dim sDateFormat As String
  Dim iLoop As Integer
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  sDateFormat = ""
    
  sSysFormat = UI.GetSystemDateFormat
  sSysDateSeparator = UI.GetSystemDateSeparator
    
  ' Loop through the string picking out the required characters.
  For iLoop = 1 To Len(sSysFormat)
      
    Select Case Mid(sSysFormat, iLoop, 1)
      Case "d"
        If Not fDaysDone Then
          ' Ensure we have two day characters.
          sDateFormat = sDateFormat & "dd"
          fDaysDone = True
        End If
          
      Case "M"
        If Not fMonthsDone Then
          ' Ensure we have two month characters.
          sDateFormat = sDateFormat & "mm"
          fMonthsDone = True
        End If
          
      Case "y"
        If Not fYearsDone Then
          ' Ensure we have four year characters.
          sDateFormat = sDateFormat & "yyyy"
          fYearsDone = True
        End If
          
      Case Else
        sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
    End Select
      
  Next iLoop
    
  ' Ensure that all day, month and year parts of the date
  ' are present in the format.
  If Not fDaysDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "dd"
  End If
    
  If Not fMonthsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "mm"
  End If
    
  If Not fYearsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "yyyy"
  End If
    
  ' Return the date format.
  DateFormat = sDateFormat
  
End Function

