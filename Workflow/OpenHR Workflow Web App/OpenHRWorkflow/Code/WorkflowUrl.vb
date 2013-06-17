
Imports System.Globalization
Imports System.Threading

Public Class WorkflowUrl
	Public InstanceId As Integer
	Public ElementId As Integer
   Public Server As String
   Public Database As String
   Public User As String
   Public Password As String
   Public UserName As String

   Public Shared Function Decrypt(value As String) As WorkflowUrl

      Dim url As New WorkflowUrl

      Try
         'Try the latest encryption method
         'Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
         Dim currentCulture = Thread.CurrentThread.CurrentCulture

         Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
         Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-GB")

         Dim crypt As New Crypt
         value = crypt.DecompactString(value)
         value = crypt.DecryptString(value, "", True)

         'Reset the culture to be the one used by the client. Fault HRPRO-1404
         Thread.CurrentThread.CurrentCulture = currentCulture
         Thread.CurrentThread.CurrentUICulture = currentCulture

         'Extract the required parameters from the decrypted queryString.
         Dim values = value.Split(vbTab(0))

         url.InstanceID = CInt(values(0))
         url.ElementID = CInt(values(1))
         url.User = values(2)
         url.Password = values(3)
         url.Server = values(4)
         url.Database = values(5)
         If values.Count > 6 Then url.UserName = values(6)

      Catch ex As Exception
         'Try the older encryption method
         Try
            Dim crypt As New Crypt
            value = crypt.ProcessDecryptString(value)
            value = crypt.DecryptString(value, "", False)

            Dim values = value.Split(vbTab(0))

            If url.InstanceID = 0 Then url.InstanceID = CInt(values(0))
            If url.ElementID = 0 Then url.ElementID = CInt(values(1))
            url.User = values(2)
            url.Password = values(3)
            url.Server = values(4)
            url.Database = values(5)
         Catch exx As Exception
            Throw New Exception("Invalid workflow url")
         End Try
      End Try

      Return url

   End Function
End Class
