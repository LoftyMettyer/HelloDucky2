'Imports System.Runtime.InteropServices

'<ClassInterface(ClassInterfaceType.None)> _
'Public Class HCM
'  Implements iSystemManager

'  Private objDatabase As New Connectivity.SQL


'  '  Implements iComPopulate

'  ''Public Shared Connection As Connectivity.SQL
'  'Public Shared MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs
'  ''    Public User As Connectivity.User
'  'Public Shared Things As Things.Collection
'  'Public Shared Workflows As Things.Collection
'  'Public Shared Operators As Things.Collection
'  'Public Shared Functions As Things.Collection
'  'Public Shared SelectedThings As Things.Collection
'  'Public Shared ErrorLog As Phoenix.ErrorHandler.Errors
'  'Public Shared ModuleSetup As Things.Collection

'  'Public Shared ScriptDB As ScriptDB.Script

'  Public Function Initialise() As Boolean Implements iSystemManager.Initialise

'    Dim bOK As Boolean = True

'    Try

'      Globals.Initialise()

'      Globals.MetadataDB = objDatabase
'      Globals.CommitDB = objDatabase

'      'Things = New Things.Collection
'      'Workflows = New Things.Collection
'      'Operators = New Things.Collection
'      'Functions = New Things.Collection
'      'ErrorLog = New Phoenix.ErrorHandler.Errors
'      'ModuleSetup = New Things.Collection

'      'ScriptDB = New ScriptDB.Script

'      ' Connect the different database types
'      'If MetadataProvider = Connectivity.MetadataProvider.LegacyDAO Then
'      '  MetadataDB = New Connectivity.AccessDB
'      'Else
'    Catch ex As Exception
'      bOK = False
'    End Try

'    Return bOK

'  End Function

'  Public Function CloseSafely() As Boolean Implements iSystemManager.CloseSafely
'    Return True
'  End Function

'#Region "iComPopulate Interface"

'  'Private mobjADODBConnection As ADODB.Connection
'  'Private mobjMetadataDB As DAO.DBEngine
'  Private miMetadataProvider As Connectivity.MetadataProvider

'#End Region

'  'Public Property MetadataDB As Object Implements Interfaces.iPhoenix.MetadataDB
'  '  Get
'  '    Return Globals.MetadataDB
'  '  End Get
'  '  Set(ByVal value As Object)
'  '    Globals.MetadataDB = value
'  '  End Set
'  'End Property

'  'Public Property CommitDB As Interfaces.iConnection Implements Interfaces.iPhoenix.CommitDB
'  '  Get
'  '    Return Globals.CommitDB
'  '  End Get
'  '  Set(ByVal value As Interfaces.iConnection)
'  '    Globals.CommitDB = value
'  '  End Set
'  'End Property

'  Public ReadOnly Property ReturnThings As Things.Collection Implements Interfaces.iSystemManager.Things
'    Get
'      Return Globals.Things
'    End Get
'  End Property

'  Public ReadOnly Property ReturnScript As ScriptDB.Script Implements Interfaces.iSystemManager.Script
'    Get

'    End Get
'  End Property

'  Public Property CommitDB As Object Implements Interfaces.iSystemManager.CommitDB
'    Get
'      Return objDatabase
'    End Get
'    Set(ByVal value As Object)
'      objDatabase = value
'    End Set
'  End Property

'  Public ReadOnly Property ErrorLog As ErrorHandler.Errors Implements Interfaces.iSystemManager.ErrorLog
'    Get
'      Return Globals.ErrorLog
'    End Get
'  End Property

'  Public Property MetadataDB As Object Implements Interfaces.iSystemManager.MetadataDB
'    Get
'      Return objDatabase
'    End Get
'    Set(ByVal value As Object)
'      objDatabase = value
'    End Set
'  End Property

'  Public ReadOnly Property Options As HCMOptions Implements Interfaces.iSystemManager.Options
'    Get
'      Return Globals.Options
'    End Get
'  End Property

'  Public Function Login(ByRef UserName As String, ByVal Password As String, ByVal Database As String, ByVal server As String) As Boolean

'    Dim objLogin As Connectivity.Login
'    Dim bOK As Boolean = True

'    Try

'      With objLogin
'        .UseContext = False
'        .UserName = UserName
'        .Password = Password
'        .Database = Database
'        .Server = server
'      End With

'      objDatabase.Login = objLogin
'      objDatabase.Open()

'    Catch ex As Exception
'      bOK = False

'    End Try

'    Return bOK

'  End Function

'  Public Function Populate(ByVal Type As HRProEngine.Things.Enums.Type) As Boolean

'    Dim bOK As Boolean

'    Try
'      bOK = Things.PopulateUtilities(Type)

'    Catch ex As Exception
'      bOK = False
'    End Try


'    Return bOK
'  End Function

'End Class


