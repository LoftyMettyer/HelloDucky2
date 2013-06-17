Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Picture

  Public Shared Function GetUrl(id As Integer, name As String) As String

    If id = 0 Then Return ""

    Dim extension As String = Path.GetExtension(name)

    Dim directory = (Configuration.Server & Configuration.Database).GetHashCode()

    Dim fileName As String = directory & "/" & id.ToString & extension

    Dim filePath As String = HttpContext.Current.Server.MapPath("~/" & fileName)

    If File.Exists(filePath) Then
      Return "~/" & fileName
    Else
      LoadPicture(id, filePath)
      Return "~/" & fileName
    End If

  End Function

  Private Shared Sub LoadPicture(id As Integer, filePath As String)

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

        Dim fileDirectory = Path.GetDirectoryName(filePath)

        If Not Directory.Exists(fileDirectory) Then
          Directory.CreateDirectory(fileDirectory)
        End If

        ' Create a file to hold the output.
        fs = New System.IO.FileStream(filePath, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
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

      End If

    End Using

  End Sub

End Class
