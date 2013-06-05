Imports System.Runtime.InteropServices
Imports SystemFramework.Things

<ClassInterface(ClassInterfaceType.None)>
Public Class SysMgr
  Implements COMInterfaces.ISystemManager

#Region "iSystemManager Interface"

  Private objMetadataDB As New Connectivity.AccessDB
  Private mobjCommitDB As New Connectivity.ADOClassic
  Private mobjScript As New ScriptDB.Script

  Public Property CommitDB As Object Implements COMInterfaces.ISystemManager.CommitDB
    Get
      Return mobjCommitDB.NativeObject
    End Get
    Set(ByVal value As Object)
      mobjCommitDB.NativeObject = CType(value, ADODB.Connection)
    End Set
  End Property

  Public Property MetadataDB As Object Implements COMInterfaces.ISystemManager.MetadataDB
    Get
      Return objMetadataDB.NativeObject
    End Get
    Set(ByVal value As Object)

      Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CType(value, DAO.Database).Name & ";"

      objMetadataDB.DB = New OleDb.OleDbConnection
      objMetadataDB.DB.ConnectionString = conStr
      objMetadataDB.NativeObject = CType(value, DAO.Database)

    End Set
  End Property

  Public Function PopulateObjects() As Boolean Implements COMInterfaces.ISystemManager.PopulateObjects

    Dim bOK As Boolean = True

    Try

      If Options Is Nothing Then
        Globals.Initialise()
      End If

      Globals.MetadataDB = objMetadataDB
      Globals.CommitDB = mobjCommitDB
      Globals.Options.DevelopmentMode = False

      Things.PopulateSystemThings()
      Things.PopulateSystemSettings()
      Things.PopulateThings()
      Things.PopulateModuleSettings()

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function

  Public Function Initialise() As Boolean Implements COMInterfaces.ISystemManager.Initialise

    Dim bOK As Boolean = True

    Try
      Globals.Initialise()
      System.Windows.Forms.Application.EnableVisualStyles()

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function

  Public Function CloseSafely() As Boolean Implements COMInterfaces.ISystemManager.CloseSafely

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

  Public ReadOnly Property ReturnTuningLog As Tuning.Report Implements COMInterfaces.ISystemManager.TuningLog
    Get
      Return Globals.TuningLog
    End Get
  End Property

  Public ReadOnly Property Version As System.Version Implements ISystemManager.Version
    Get
      Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
    End Get
  End Property

  Public ReadOnly Property ReturnErrorLog As ErrorHandler.Errors Implements COMInterfaces.ISystemManager.ErrorLog
    Get
      Return Globals.ErrorLog
    End Get
  End Property

  Public ReadOnly Property ReturnThings As ICollection(Of Table) Implements COMInterfaces.ISystemManager.Tables
    Get
      Return Globals.Tables
    End Get
  End Property

  Public ReadOnly Property Script As ScriptDB.Script Implements COMInterfaces.ISystemManager.Script
    Get
      Return mobjScript
    End Get
  End Property

  Public ReadOnly Property Options As HCMOptions Implements COMInterfaces.ISystemManager.Options
    Get
      Return Globals.Options
    End Get
  End Property

  Public ReadOnly Property Modifications As Modifications Implements COMInterfaces.ISystemManager.Modifications
    Get
      Return Globals.Modifications
    End Get
  End Property

#End Region

End Class
