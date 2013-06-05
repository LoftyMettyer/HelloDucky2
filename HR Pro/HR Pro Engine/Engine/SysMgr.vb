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

      '     Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + value.Name + ";"
      'Dim conStr As String = "Driver=SQL Server;Server={harpdev01};UID=sa;PWD=asr;Database=phoenix;"""

      'Debug.Print(value.ConnectionString)

      'mobjCommitDB.DB = New OleDb.OleDbConnection
      'mobjCommitDB.DB.ConnectionString = conStr
      'objMetadataDB.DB.Open()

      ''  value.ConnectionString
      ''      mobjCommitDB.NativeObject = New Connectivity.SQL
      'mobjCommitDB.NativeObject.ConnectionString = value.ConnectionString
      ''  mobjCommitDB.Open()

      ''value.

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
      '    objMetadataDB.DB.Open()

      objMetadataDB.NativeObject = value

    End Set
  End Property

  Public Function Initialise() As Boolean Implements Interfaces.iSystemManager.Initialise

    Dim bOK As Boolean = True

    Try

      Globals.Initialise()
      Globals.MetadataDB = objMetadataDB
      Globals.CommitDB = mobjCommitDB

      'MakeAccessDBReady()
      Things.PopulateSystemThings()
      Things.PopulateThings()

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

      ''   objMetadataDB.Close()
      'objMetadataDB = Nothing

      '  Globals.MetadataDB.Close()


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

  '' Place the stored procedures/queries onto the access db (compatability with rest of engine)
  'Private Sub MakeAccessDBReady()

  '  Dim sSQL As String

  '  ' Clear the existing ones (if any)
  '  Try

  '          objMetadataDB.NativeObject.QueryDefs.

  '    objMetadataDB.NativeObject.QueryDefs.Delete("spadmin_gettables")
  '    objMetadataDB.NativeObject.QueryDefs.Delete("spadmin_getcolumns")

  '  Catch ex As Exception


  '  End Try

  '  ' spadmin_gettables
  '  sSQL = "SELECT tmpTables.tableid AS ID" & _
  '      ", tmpTables.tablename AS name" & _
  '      ", 1 AS type" & _
  '      ", '' AS description,tmpTables.[TableType] AS tabletype" & _
  '      ", 0 AS isremoteview" & _
  '      ", [tabletype]" & _
  '      ", tmpTables.[RecordDescExprID] AS recorddescriptionid" & _
  '      ", tmpTables.[AuditInsert] AS auditinsert" & _
  '      ", tmpTables.[AuditDelete] AS auditdelete" & _
  '      ", tmpTables.[DefaultEmailID] AS defaultemailid" & _
  '      ", tmpTables.[DefaultOrderID] AS defaultorderid" & _
  '      " FROM tmpTables;"
  '  objMetadataDB.NativeObject.CreateQueryDef("spadmin_gettables", sSQL)

  '  ' spadmin_getcolumns
  '  sSQL = "SELECT tmpColumns.columnid as [ID]" & _
  '      ", tmpColumns.columnname as [name]" & _
  '      ", 2 as [type]" & _
  '      ",'' AS [description]" & _
  '      ", tmpColumns.[calcExprID] AS [calcid]" & _
  '      ", tmpColumns.[datatype] as [datatype]" & _
  '      ", tmpColumns.[size], [decimals]" & _
  '      ", tmpColumns.[audit], [mandatory], [multiline]	" & _
  '      ", ISNULL(tmpColumns.[dfltvalueexprid],0) as [defaultcalcid]" & _
  '      ", tmpColumns.[convertcase] as [case]" & _
  '      ", tmpColumns.[readOnly] as [isreadonly]" & _
  '      " FROM tmpColumns WHERE tmpColumns.tableid = @parentid" & _
  '      " AND tmpColumns.[ColumnName] NOT LIKE 'ID%'" & _
  '      " ORDER BY tmpColumns.[columnname];"
  '  objMetadataDB.NativeObject.CreateQueryDef("spadmin_getcolumns", sSQL)

  'End Sub

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
