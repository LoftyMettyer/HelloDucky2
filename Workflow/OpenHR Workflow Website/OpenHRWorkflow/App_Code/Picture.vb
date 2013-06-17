Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Picture

  'Public Shared Function GetFileName(id As Integer, Optional name As String = Nothing) As String

  '  If name Is Nothing Then Return ""

  '  Dim ext As String = Path.GetExtension(name)

  '  Return String.Format("Pic{0}{1}", id, ext)

  'End Function

  'Public Shared Function GetUrl(id As Integer, Optional name As String = Nothing) As String

  '  If id = 0 Then Return ""

  '  Dim fileName As String
  '  fileName = Picture.GetFileName(id, name)
  '  fileName = HttpContext.Current.Server.MapPath(fileName)
  '  Picture.LoadPicture(id, fileName)

  '  Return fileName

  'End Function

  Public Shared Function LoadPicture(id As Integer) As String

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("spASRGetPicture", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@piPictureID", SqlDbType.Int).Direction = ParameterDirection.Input
      cmd.Parameters("@piPictureID").Value = id

      Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)

      Dim fs As FileStream
      Dim bw As BinaryWriter

      Const bufferSize As Integer = 100
      Dim outByte(bufferSize - 1) As Byte
      Dim retVal As Long
      Dim startIndex As Long

      If dr.Read Then

        Dim ext As String = Path.GetExtension(CStr(dr("Name")))
        Dim fileName As String = String.Format("Pic-{0}{1}", Guid.NewGuid, ext)
        Dim filePath = HttpContext.Current.Server.MapPath("~/Pictures")

        If Not Directory.Exists(filePath) Then
          Directory.CreateDirectory(filePath)
        End If

        ' Create a file to hold the output.
        fs = New System.IO.FileStream(filePath & "\" & fileName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        bw = New System.IO.BinaryWriter(fs)

        ' Reset the starting byte for a new BLOB.
        startIndex = 0

        ' Read bytes into outbyte() and retain the number of bytes returned.
        retVal = dr.GetBytes(1, startIndex, outByte, 0, bufferSize)

        ' Continue reading and writing while there are bytes beyond the size of the buffer.
        Do While retVal = bufferSize
          bw.Write(outByte)
          bw.Flush()

          ' Reposition the start index to the end of the last buffer and fill the buffer.
          startIndex += bufferSize
          retVal = dr.GetBytes(1, startIndex, outByte, 0, bufferSize)
        Loop

        ' Write the remaining buffer.
        bw.Write(outByte)
        bw.Flush()

        ' Close the output file.
        bw.Close()
        fs.Close()

        Return "~/Pictures/" & fileName
      Else
        Return ""
      End If

    End Using

  End Function

End Class
