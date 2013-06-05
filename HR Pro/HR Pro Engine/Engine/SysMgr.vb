Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class SysMgr
  Implements iSystemManager

#Region "iSystemManager Interface"

  Private objMetadataDB As New Connectivity.AccessDB
  Private mobjCommitDB As New Connectivity.ADOClassic
  Private mobjScript As New ScriptDB.Script

  Public Property CommitDB As Object Implements Interfaces.iSystemManager.CommitDB
    Get
      Return mobjCommitDB.NativeObject
    End Get
    Set(ByVal value As Object)
      mobjCommitDB.NativeObject = value
    End Set
  End Property

  Public Property MetadataDB As Object Implements Interfaces.iSystemManager.MetadataDB
    Get
      Return objMetadataDB.NativeObject
    End Get
    Set(ByVal value As Object)

      Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + value.Name + ";"

      objMetadataDB.DB = New OleDb.OleDbConnection
      objMetadataDB.DB.ConnectionString = conStr
      objMetadataDB.NativeObject = value

    End Set
  End Property

  Public Function Initialise() As Boolean Implements Interfaces.iSystemManager.Initialise

    Dim bOK As Boolean = True

    Try

      Globals.Initialise()
      Globals.MetadataDB = objMetadataDB
      Globals.CommitDB = mobjCommitDB
      Globals.Options.DevelopmentMode = False

      Things.PopulateSystemThings()
      Things.PopulateThings()
      Things.PopulateModuleSettings()

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function

  Public Function CloseSafely() As Boolean Implements iSystemManager.CloseSafely

    Dim bOK As Boolean = True

    Try
      objMetadataDB.DB.Close()
      objMetadataDB.NativeObject.Close()

      objMetadataDB.DB = Nothing
      objMetadataDB.NativeObject = Nothing

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function

  Public ReadOnly Property ReturnErrorLog As ErrorHandler.Errors Implements Interfaces.iSystemManager.ErrorLog
    Get
      Return Globals.ErrorLog
    End Get
  End Property

  Public ReadOnly Property ReturnThings As Things.Collection Implements Interfaces.iSystemManager.Things
    Get
      Return Globals.Things
    End Get
  End Property

  Public ReadOnly Property Script As ScriptDB.Script Implements Interfaces.iSystemManager.Script
    Get
      Return mobjScript
    End Get
  End Property

  Public ReadOnly Property Options As HCMOptions Implements Interfaces.iSystemManager.Options
    Get
      Return Globals.Options
    End Get
  End Property

#End Region


End Class
