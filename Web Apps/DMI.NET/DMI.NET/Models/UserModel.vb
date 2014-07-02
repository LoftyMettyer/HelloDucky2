Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.ObjectModel

Namespace Models
  Public Class UserModel

    Private _userName As String
    Private _userNames As Collection(Of SelectListItem)

    Public Property Username As String
      Get
        Return _userName
      End Get
      Set(value As String)
        _userName = value
      End Set
    End Property

    Public ReadOnly Property UserNames As Collection(Of SelectListItem)
      Get
        Return _userNames
      End Get
    End Property

    Public Sub PopulateFromDB()

      Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)

      Try

        Dim rstLogins = objDataAccess.GetDataTable("spASRIntGetAvailableLoginsFromAssembly", CommandType.StoredProcedure)

        _userNames = New Collection(Of SelectListItem)
        For Each objRow In rstLogins.Rows
          Dim objItem As New SelectListItem()
          objItem.Text = objRow("name").ToString()

          _userNames.Add(objItem)
        Next

      Catch ex As Exception

      End Try

    End Sub

    Public Function CreateLogin() As Boolean

      Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)

      Try
        objDataAccess.ExecuteSP("sp_ASRIntNewUser", _
            New SqlParameter("@psUserName", SqlDbType.VarChar, 128) With {.Value = _userName})

      Catch ex As Exception
        Throw

      End Try

      Return True

    End Function

  End Class
End Namespace