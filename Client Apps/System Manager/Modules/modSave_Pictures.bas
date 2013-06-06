Attribute VB_Name = "modSave_Pictures"
Option Explicit

Public Function SavePictures() As Boolean
  ' Save the new and modified Picture records.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngPictureCount As Long
  Dim rsPictures As ADODB.Recordset
  
  Set rsPictures = New ADODB.Recordset
  fOK = True
  
  ' Open the Pictures table on the server.
  rsPictures.Open "SELECT PictureID FROM ASRSysPictures", gADOCon, adOpenStatic, adLockOptimistic
  lngPictureCount = rsPictures.RecordCount
  rsPictures.Close
  
  rsPictures.Open "ASRSysPictures", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  With rsPictures
    Do While Not .EOF
      recPictEdit.Index = "idxID"
      recPictEdit.Seek "=", !PictureID
      
      OutputCurrentProcess2 recPictEdit.Fields("Name").Value, lngPictureCount
      gobjProgress.UpdateProgress2
      
      If recPictEdit.NoMatch Then
        .Delete
      ElseIf recPictEdit!Deleted = True Then
        .Delete
      ElseIf recPictEdit!Changed = True Then
        '.Edit
        !Name = recPictEdit!Name
        .Update
      End If
      
      .MoveNext
    Loop
    
    If Not (recPictEdit.BOF And recPictEdit.EOF) Then
      recPictEdit.MoveFirst
    End If
    
    Do While Not recPictEdit.EOF
      If recPictEdit!New = True Then
        .AddNew
        If CopyPictureField(recPictEdit!Picture, !Picture) Then
          !PictureID = recPictEdit!PictureID
          !Name = recPictEdit!Name
          !PictureType = recPictEdit!PictureType
          .Update
        Else
          .CancelUpdate
        End If
      End If
      
      recPictEdit.MoveNext
    Loop
  End With
  
  rsPictures.Close
  
TidyUpAndExit:
  Set rsPictures = Nothing
  SavePictures = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error saving pictures"
  Resume TidyUpAndExit
  
End Function


Private Function CopyPictureField(SourceFld As dao.Field, DestCol As ADODB.Field) As Boolean
  On Error GoTo ErrorTrap
  
  Dim lngChunkSize As Long
  Dim intChunks As Integer
  Dim lngOffSet As Long
  Dim Chunk() As Byte
  Dim iLoop As Integer

  intChunks = SourceFld.FieldSize \ ChunkSize
  lngChunkSize = SourceFld.FieldSize Mod ChunkSize
        
  If lngChunkSize > 0 Then
    ReDim Chunk(lngChunkSize)
    Chunk() = SourceFld.GetChunk(0, lngChunkSize)
    DestCol.AppendChunk Chunk()
    lngOffSet = lngOffSet + lngChunkSize
  End If
        
  lngChunkSize = ChunkSize
  For iLoop = 1 To intChunks
    ReDim Chunk(lngChunkSize)
    Chunk() = SourceFld.GetChunk(lngOffSet, lngChunkSize)
    DestCol.AppendChunk Chunk()
    lngOffSet = lngOffSet + lngChunkSize
  Next iLoop
  
  CopyPictureField = True

  Exit Function

ErrorTrap:
  CopyPictureField = False
  
  MsgBox Err.Description, vbOKOnly + vbExclamation, Application.Name
  
  Err = False
  
End Function

