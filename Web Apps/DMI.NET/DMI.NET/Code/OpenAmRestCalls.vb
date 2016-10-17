﻿Imports System.IO
Imports System.Net
Imports System.Web.Script.Serialization

Namespace Code

   Public Class OpenAmRestCalls
      Public Shared Function LoginAndReturnToken(username As String, password As String) As String
         Dim requestHeaders As New NameValueCollection

         requestHeaders.Add("x-openam-username", username)
         requestHeaders.Add("x-openam-password", password)

         Dim openAmAuthenticateUri = ApplicationSettings.OpenAmAuthenticateUri

         Return CallRestApiAndReturnObject(Of LoginResponse)(openAmAuthenticateUri, "POST", requestHeaders).TokenId
      End Function
      Public Shared Function GetIdFromSession(tokenId As String) As String
         Dim requestHeaders As New NameValueCollection
         requestHeaders.Add("iplanetdirectorypro", tokenId)
         
         Dim openAmGetIdFromSessionUri = ApplicationSettings.OpenAmGetIdFromSessionUri

         Return CallRestApiAndReturnObject(Of GetIdFromSessionResponse)(openAmGetIdFromSessionUri, "POST", requestHeaders).Id
      End Function
      Private Shared Function CallRestApiAndReturnObject(Of T)(requestUri As String, requestMethod As String, requestHeaders As NameValueCollection) As T
         Dim request = WebRequest.Create(requestUri)
         request.Method = requestMethod
         request.ContentType = "application/json"
         request.Headers.Add(requestHeaders)

         Try
            Dim webResponse = request.GetResponse()
            Dim webStream = webResponse.GetResponseStream()
            Dim responseReader = New StreamReader(webStream)
            Dim response = responseReader.ReadToEnd()

            Dim serializer = New JavaScriptSerializer()
            Dim getResponseObject = serializer.Deserialize(Of T)(response)
            responseReader.Close()

            Return getResponseObject
         Catch ex As Exception
            Throw
         End Try
      End Function

      Private Class LoginResponse
         Public Property TokenId As String
         Public Property SuccessUrl As String
      End Class

      Private Class GetIdFromSessionResponse
         Public Property Id As String
         Public Property Realm As String
         Public Property Dn As String
         Public Property SuccessUrl As String
         Public Property FullLoginUrl As String
      End Class
   End Class
End Namespace