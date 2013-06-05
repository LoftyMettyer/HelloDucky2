'Imports System.IO
'Imports System.Xml.Serialization

Public Class Form1

  'Private objProgress As Phoenix.HCMProgressBar
  Private Initialised As Boolean = False

  'Private Sub DisplayErrors()

  '  For Each objError As Phoenix.ErrorHandler.Error In mobjPhoenix.ErrorLog
  '    '      Debug.Print(String.Join(vbNewLine, ScriptDB.HCM.ErrorLog.ToArray(GetType(String))))
  '    Debug.Print(String.Format("{0}--{1}--{2}", objError.Section, objError.ObjectName, objError.Message))
  '  Next

  '  Debug.Print(mobjPhoenix.ErrorLog.Count)

  '  '   Phoenix.ErrorLog.Count
  '  '    mobjPhoenix

  'End Sub

  Private Sub InitialiseStuff()

    ' objProgress = New Phoenix.HCMProgressBar
    '  AddHandler objProgress.Update1, AddressOf UpdateProgress1
    '  AddHandler objProgress.Update2, AddressOf UpdateProgress2

    '  If Not Initialised Then
    '    mobjPhoenix.Initialise()


    '    'mobjPhoenix.

    '    '   Label1.Text = mobjPhoenix.CommitDB
    '    '. .Login.Database

    '    '   mobjPhoenix.Connection.Open()

    '    '      CurrentPhase.Text = "Populating Objects..."
    '    '  mobjPhoenix.Things.populatethings()
    '    '  Phoenix.Things.PopulateSystemThings()
    '    ' Phoenix.Things.PopulateThings(objProgress)
    '    ' Phoenix.Things.PopulateModuleSettings(objProgress)

    '    Initialised = True
    '  End If

  End Sub


  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    InitialiseStuff()

    '  objProgress.TotalSteps1 = 6

    ' Label1.Text = Phoenix.CommitDB.Login.Database
    ''    Dim objAssembly As ScriptDB.Connectivity.ConnectionType

    ' Generate all the SQL code for our objects
    CurrentPhase.Text = "Generating expression code..."
    '  Phoenix.HCM.ScriptDB.GenerateSQLCodeForObjects(objProgress)


    ' Commit any UDF calculations
    CurrentPhase.Text = "Scripting UDFs..."
    '  Phoenix.HCM.ScriptDB.ScriptColumnCalculations(objProgress)

    ' Comit record descriptions
    CurrentPhase.Text = "Scripting Record Descriptions..."
    '   Phoenix.HCM.ScriptDB.ScriptRecordDescriptions(objProgress)

    ' Commit any triggers
    CurrentPhase.Text = "Scripting Triggas..."
    '    Phoenix.HCM.ScriptDB.ScriptTriggers()

    ' Commit any views
    'ScriptDB.ScriptDB.ScriptViews(objProgress)

    '   DisplayErrors()

  End Sub

#Region "Progress Bar Handling"

  Private Sub UpdateProgress1(ByVal Value As Long)
    ProgressBar1.Value = Value
  End Sub
  Private Sub UpdateProgress2(ByVal Value As Long)
    ProgressBar2.Value = Value
  End Sub

#End Region

  Private Sub RemoteViewScript(ByVal bTurnOn As Boolean)

    'Dim objScript As New Phoenix.ScriptDB.Script

    '  InitialiseStuff()

    CurrentPhase.Text = "Generating tables..."
    'ScriptDB.ScriptDB.ScriptTables(objProgress)
    '  Phoenix.HCM.ScriptDB.ScriptViews()
    '  DisplayErrors()

  End Sub


  Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    RemoteViewScript(True)
  End Sub

  Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    RemoteViewScript(False)
  End Sub


  ' Suck and spit module
  Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


    ''  Dim frmMapObjects As New MapObjects
    'frmMapObjects.Show()
    '    Me.Dispose(False)




    'Dim objFile As System.IO.File

    'CurrentPhase.Text = "Generating port script..."
    'InitialiseStuff()
    'ScriptDB.StructurePort.CreateStatements(objProgress)



    'objFile.WriteAllLines(txtUpdateScript.Text, ScriptDB.StructurePort.GetStatements)


    ''    StructurePort.Export.ScriptStatements(objProgress)







    ''Dim objStreamWriter As New StreamWriter(txtUpdateScript.Text)
    ''Dim x As New XmlSerializer(ScriptDB.HCM.Things(0).GetType)
    ''x.Serialize(objStreamWriter, ScriptDB.HCM.Things(0))
    ''objStreamWriter.Close()
    ''Dim objColumn As ScriptDB.Things.Column
    ''objColumn = ScriptDB.HCM.Things(0).Objects(ScriptDB.Things.Type.Column).Item(0)

    '' Debug.Print(objColumn.ToXML)

    'DisplayErrors()

  End Sub

  Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

    Dim objPhoenix As New HRProEngine.SysMgr

    Dim objADO As New ADODB.Connection
    Dim objDAOEngine As New DAO.DBEngine
    Dim objDAODB As DAO.Database
    Dim objRecordset As DAO.Recordset


    ' THIS IS SYSTEM MGR RECCREATION
    Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\dev\play\AsrTemp_phoenix.mdb"

    objDAODB = objDAOEngine.OpenDatabase("c:\dev\play\AsrTemp_phoenix.mdb", , False, conStr)

    objRecordset = objDAODB.OpenRecordset("spadmin_gettables")
    objRecordset.MoveFirst()
    Do While Not objRecordset.EOF
      Debug.Print(objRecordset.Fields("name").Value().ToString)
      objRecordset.MoveNext()
    Loop
    ' -----------------------------------------------------------------

    ' our stuff
    '    objDAODB.CreateQueryDef("spadmin_gettables",
    '   mobjPhoenix.Mode = Phoenix.Connectivity.MetadataProvider.LegacyDAO
    '    mobjPhoenix.CommitDB = objDAODB

    objPhoenix.MetadataDB = objDAODB
    objPhoenix.Initialise()

    '  mobjPhoenix.CommitDB = objADO



    ' -----------------------------------------------------------------
    ' THIS IS SYSTEM MGR RECCREATION
    objDAODB.Close()




  End Sub

  Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

    Dim conStr As String = "Driver={Microsoft Access Driver (*.mdb)};DBQ=c:\dev\play\AsrTemp_phoenix.mdb;"
    Dim objDAODB As DAO.Database
    Dim objDAOEngine As New DAO.DBEngine

    '= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\dev\play\AsrTemp_phoenix.mdb"
    objDAODB = objDAOEngine.OpenDatabase("c:\dev\play\AsrTemp_phoenix.mdb", , False, conStr)


    Dim ODBCdb As New Odbc.OdbcConnection
    Dim objCommand As New Odbc.OdbcCommand
    Dim objAdapter As New Odbc.OdbcDataAdapter
    Dim objDataset As New System.Data.DataSet
    Dim objRow As System.Data.DataRow

    ODBCdb.ConnectionString = conStr
    ODBCdb.Open()

    With objCommand
      .CommandType = CommandType.StoredProcedure
      .CommandText = "EXECUTE spadmin_gettables"                    'Specify stored procedure to run
      .Connection = ODBCdb
    End With

    objAdapter.SelectCommand = objCommand
    objAdapter.Fill(objDataset)
    For Each objRow In objDataset.Tables(0).Rows
      Debug.Print(objRow.Item("Name").ToString)
    Next

    ODBCdb.Close()
    objDAODB.Close()


  End Sub

  Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

    Dim objPhoenix As New HRProEngine.SysMgr

    Dim sPath As String = "C:\dev\HR Pro\HR Pro Engine\Test Harness\"

    Dim objADO As New ADODB.Connection
    Dim objDAOEngine As New DAO.DBEngine
    Dim objDAODB As DAO.Database
    '    Dim objRecordset As DAO.Recordset
    Dim sADOConnect As String = "Driver=SQL Server;Server={harpqa02};UID=sa;PWD=asr;Database=" & txtDatabase.Text & ";"""
    '  Dim objADOLogin As Phoenix.Connectivity.Login

    ' THIS IS SYSTEM MGR RECCREATION
    Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & "AsrTemp_" & txtDatabase.Text & ".mdb"

    Dim bOK As Boolean


    objDAODB = objDAOEngine.OpenDatabase(sPath & "asrtemp_" & txtDatabase.Text & ".mdb", , False, conStr)
    '    objADO = ADODB

    With objADO
      .ConnectionString = sADOConnect
      .Provider = "SQLOLEDB"
      .CommandTimeout = 0
      .ConnectionTimeout = 5
      .CursorLocation = ADODB.CursorLocationEnum.adUseServer
      .Mode = ADODB.ConnectModeEnum.adModeReadWrite
      '.Properties("Packet Size") = 32767
      .Open()
    End With

    'With objADOLogin
    '  .Database = "phoenix"
    '  .Server = "harpdev01"
    '  .UserName = "sa"
    '  .Password = "asr"
    'End With

    objPhoenix.MetadataDB = objDAODB
    objPhoenix.CommitDB = objADO

    objPhoenix.Initialise()


    objPhoenix.Options.RefreshObjects = True
    bOK = objPhoenix.Script.CreateObjects()


    bOK = objPhoenix.Script.CreateTriggers()


    'objPhoenix.Script.DropViews()
    'objPhoenix.Script.DropTableViews()
    'objPhoenix.Script.CreateTables()

    'objPhoenix.Script.CreateTableViews()
    'objPhoenix.Script.CreateViews()
    'objPhoenix.Script.ApplySecurity()

    bOK = objPhoenix.Script.CreateFunctions



    objPhoenix.CloseSafely()


    'objRecordset = objDAODB.OpenRecordset("spadmin_gettables")
    'objRecordset.MoveFirst()
    'Do While Not objRecordset.EOF
    '  Debug.Print(objRecordset.Fields("name").Value().ToString)
    '  objRecordset.MoveNext()
    'Loop
    '' -----------------------------------------------------------------

    '' our stuff
    ''    objDAODB.CreateQueryDef("spadmin_gettables",
    ''   mobjPhoenix.Mode = Phoenix.Connectivity.MetadataProvider.LegacyDAO
    ''    mobjPhoenix.CommitDB = objDAODB

    'objPhoenix.MetadataDB = objDAODB
    'objPhoenix.Initialise()

    '  mobjPhoenix.CommitDB = objADO

    '    objPhoenix.

    objPhoenix.ReturnErrorLog.OutputToFile("c:\dev\errors.txt")

    'Dim objError As Phoenix.ErrorHandler.Error
    'Dim strMessage As String = vbNullString
    'For Each objError In objPhoenix.ReturnErrorLog
    '  'strMessage = String.Format("{4}{0}{1}{2}-----------------{2}{3}{2}{2}" _
    '  '    , objError.ObjectName, objError.Message, vbNewLine, objError.Detail, strMessage)

    '  strMessage = String.Format("{3}{0}{2}{0}{1}" _
    '      , objError.Message, vbNewLine, strMessage, objError.Detail)

    'Next
    'txtErrors.Text = strMessage


    ' -----------------------------------------------------------------
    ' THIS IS SYSTEM MGR RECCREATION
    ' objDAODB.Close()
    objDAODB = Nothing

  End Sub

  Private Sub cmdDatasource_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDatasource.Click

    Dim objPhoenix As New HRProEngine.SysMgr

    Dim objADO As New ADODB.Connection
    Dim objDAOEngine As New DAO.DBEngine
    Dim objDAODB As DAO.Database
    '    Dim objRecordset As DAO.Recordset
    Dim sADOConnect As String = "Driver=SQL Server;Server={harpdev01};UID=sa;PWD=asr;Database=phoenix;"""
    'Dim objADOLogin As Phoenix.Connectivity.Login

    ' THIS IS SYSTEM MGR RECCREATION
    Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\dev\play\AsrTemp_phoenix.mdb"

    objDAODB = objDAOEngine.OpenDatabase("c:\dev\play\AsrTemp_phoenix.mdb", , False, conStr)
    '    objADO = ADODB

    With objADO
      .ConnectionString = sADOConnect
      .Provider = "SQLOLEDB"
      .CommandTimeout = 0
      .ConnectionTimeout = 5
      .CursorLocation = ADODB.CursorLocationEnum.adUseServer
      .Mode = ADODB.ConnectModeEnum.adModeReadWrite
      '.Properties("Packet Size") = 32767
      .Open()
    End With

    'With objADOLogin
    '  .Database = "phoenix"
    '  .Server = "harpdev01"
    '  .UserName = "sa"
    '  .Password = "asr"
    'End With

    objPhoenix.MetadataDB = objDAODB
    objPhoenix.CommitDB = objADO

    objPhoenix.Initialise()

    '  objPhoenix.Script.
    '    objPhoenix
    '     objPhoenix.ReturnThings

    ' -----------------------------------------------------------------
    ' THIS IS SYSTEM MGR RECCREATION
    objDAODB.Close()




  End Sub

  Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

    Dim objHCM As New HRProEngine.HCM

    '    objHCM.


    objHCM.Initialise()
    objHCM.Login("sa", "asr", "phoenix", "harpdev01")
    objHCM.Populate(HRProEngine.Things.Type.GlobalModify)




  End Sub

End Class
