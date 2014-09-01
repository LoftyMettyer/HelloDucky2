Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)>
Public Class SysMgr
  Implements ISystemManager


#Region "iSystemManager Interface"

  Private ReadOnly _objMetadataDb As New Connectivity.AccessDb
  Private ReadOnly _mobjCommitDb As New Connectivity.ADOClassic
  Private ReadOnly _mobjScript As New ScriptDB.Script

  Public Property CommitDB As Object Implements ISystemManager.CommitDB
    Get
      Return _mobjCommitDb.NativeObject
    End Get
    Set(ByVal value As Object)
      _mobjCommitDb.NativeObject = CType(value, ADODB.Connection)
    End Set
  End Property

  Public Property MetadataDB As Object Implements ISystemManager.MetadataDB
    Get
      Return _objMetadataDb.NativeObject
    End Get
    Set(ByVal value As Object)

      Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CType(value, DAO.Database).Name & ";"

      _objMetadataDb.Db = New OleDb.OleDbConnection
      _objMetadataDb.Db.ConnectionString = conStr
      _objMetadataDb.NativeObject = CType(value, DAO.Database)

    End Set
  End Property

  Public Function PopulateObjects() As Boolean Implements ISystemManager.PopulateObjects

    Dim bOk As Boolean = True

    Try

      If Options Is Nothing Then
        Globals.Initialise()
      End If

      ' Clear any existing errors
      ErrorLog.Clear()

      Globals.MetadataDb = _objMetadataDb
      Globals.CommitDb = _mobjCommitDb
      Globals.Options.DevelopmentMode = False

      Dim sw As New Stopwatch
      sw.Start()
      PopulateSystemThings()

      PopulateSystemSettings()
      PopulateThings()
      PopulateModuleSettings()

    Catch ex As Exception
      bOk = False
    End Try

    Return bOk

  End Function

  Public Function Initialise() As Boolean Implements ISystemManager.Initialise

    Dim bOk As Boolean = True

    Try
      Globals.Initialise()
      Windows.Forms.Application.EnableVisualStyles()

    Catch ex As Exception
      bOk = False
    End Try

    Return bOk

  End Function

  Public Function CloseSafely() As Boolean Implements ISystemManager.CloseSafely

    Dim bOk As Boolean = True

    Try
      _objMetadataDb.Db.Close()
      _objMetadataDb.NativeObject.Close()

      _objMetadataDb.Db = Nothing
      _objMetadataDb.NativeObject = Nothing

    Catch ex As Exception
      bOk = False
    End Try

    Return bOk

  End Function

  Public ReadOnly Property ReturnTuningLog As TuningReport Implements ISystemManager.TuningLog
    Get
      Return TuningLog
    End Get
  End Property

  Public ReadOnly Property Version As Version Implements ISystemManager.Version
    Get
      Return Reflection.Assembly.GetExecutingAssembly().GetName().Version
    End Get
  End Property

  Public ReadOnly Property ReturnErrorLog As Collections.Errors Implements ISystemManager.ErrorLog
    Get
      Return ErrorLog
    End Get
  End Property

  Public Function GetTable(ByVal id As Integer) As Table Implements ISystemManager.GetTable
    Return Tables.GetById(id)
  End Function

  Public ReadOnly Property Script As ScriptDB.Script Implements ISystemManager.Script
    Get
      Return _mobjScript
    End Get
  End Property

  Public ReadOnly Property Options As [Option] Implements ISystemManager.Options
    Get
      Return Globals.Options
    End Get
  End Property

  Public ReadOnly Property Modifications As Modifications Implements ISystemManager.Modifications
    Get
      Return Globals.Modifications
    End Get
  End Property

	Public Function UpdateLicence(existingLicence As String) As String Implements ISystemManager.UpdateLicence

		Dim objLicence As New UpdateLicence
		objLicence.SetOldLicenceKey(existingLicence)
		Return objLicence.GenerateNewKey

	End Function


#End Region

End Class
