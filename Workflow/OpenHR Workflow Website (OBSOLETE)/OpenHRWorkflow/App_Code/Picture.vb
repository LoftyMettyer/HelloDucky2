Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class Picture

  Private Shared ReadOnly Map As New Dictionary(Of Integer, String)
  Private Shared ReadOnly Folder As String
  Private Shared ReadOnly RootPath As String
  Private Shared ReadOnly FullPath As String
  Private Shared ReadOnly Lock As New Object

  Shared Sub New()

    Folder = (Configuration.Server & Configuration.Database).GetHashCode().ToString()
    RootPath = "~/Pictures/" & Folder
    FullPath = HttpContext.Current.Server.MapPath(RootPath)

    If Directory.Exists(FullPath) Then

      For Each file In Directory.GetFiles(FullPath)

        Dim name = Path.GetFileName(file)
        Dim id = Path.GetFileNameWithoutExtension(file)

        Map.Add(CInt(id), RootPath & "/" & name)
      Next

    End If

  End Sub

  Public Shared Function GetUrl(id As Integer) As String

    If id = 0 Then Return String.Empty

    Dim url As String = String.Empty

    SyncLock Lock

      If Not Map.TryGetValue(id, url) Then

        If Not Directory.Exists(FullPath) Then
          Directory.CreateDirectory(FullPath)
        End If

        Dim file = LoadPicture(id)
        url = RootPath & "/" & Path.GetFileName(file)
        Map.Add(CInt(id), url)

      End If

    End SyncLock

    Return url

  End Function

  Private Shared Function LoadPicture(id As Integer) As String

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

      Dim filePath As String = String.Empty

      If dr.Read Then

        filePath = FullPath & "\" & id.ToString() & Path.GetExtension(CStr(dr("Name")))

        ' Create a file to hold the output.
        fs = New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write)
        bw = New BinaryWriter(fs)

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

      Return filePath
    End Using

  End Function

End Class
