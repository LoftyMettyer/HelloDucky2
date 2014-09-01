'Imports System.IO
'Imports System.Xml.Serialization

Public Class MainForm

		'Private objProgress As Phoenix.HCMProgressBar
	Private Initialised As Boolean

#Region "Scripting stuff (some unused)"

		Private Sub butScriptDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butScriptDB.Click

				Dim objPhoenix As New SystemFramework.SysMgr

		Dim sPath As String = "C:\dev\Components\System Framework\Test Harness\"

				Dim objADO As New ADODB.Connection
				Dim objDAOEngine As New DAO.DBEngine
				Dim objDAODB As DAO.Database
				'    Dim objRecordset As DAO.Recordset
				Dim sADOConnect As String = String.Format("Driver=SQL Server;Server={0};UID=sa;PWD={2};Database={1};" _
																	, txtServer.Text, txtDatabase.Text, txtPassword.Text)
				'  Dim objADOLogin As Phoenix.Connectivity.Login

				' THIS IS SYSTEM MGR RECCREATION
				Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & "AsrTemp_" & txtDatabase.Text & ".mdb"

				Dim bOK As Boolean

				objDAODB = objDAOEngine.OpenDatabase(sPath & "asrtemp_" & txtDatabase.Text & ".mdb", , False, conStr)

				With objADO
						.ConnectionString = sADOConnect
						.Provider = "SQLOLEDB"
						.CommandTimeout = 0
						.ConnectionTimeout = 5
						.CursorLocation = ADODB.CursorLocationEnum.adUseServer
						.Mode = ADODB.ConnectModeEnum.adModeReadWrite
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
			Dim sw As New Stopwatch
			sw.Start()
			objPhoenix.PopulateObjects()
		Console.WriteLine(String.Format("Populate: {0} secs", sw.Elapsed.TotalSeconds))
		'approx 60 seconds now 5 seconds

		objPhoenix.Options.DevelopmentMode = chkDebugMode.Checked
		objPhoenix.Options.RefreshObjects = True

		sw.Reset()
		sw.Start()
		bOK = objPhoenix.Script.CreateObjects()
		Console.WriteLine(String.Format("Create Objects: {0} milliseconds", sw.ElapsedMilliseconds))

		sw.Restart()
		bOK = objPhoenix.Script.CreateTriggers()
		Console.WriteLine(String.Format("Create Triggers: {0} milliseconds", sw.ElapsedMilliseconds))

		sw.Restart()
		bOK = objPhoenix.Script.CreateFunctions
		Console.WriteLine(String.Format("Create Functions: {0} milliseconds", sw.ElapsedMilliseconds))

		sw.Restart()
		bOK = objPhoenix.Script.ScriptIndexes
		Console.WriteLine(String.Format("Create Indexes: {0} milliseconds", sw.ElapsedMilliseconds))

		'bOK = objPhoenix.Script.CreateTableViews
		'bOK = objPhoenix.Script.CreateViews
		'objPhoenix.Script.DropViews()
		'objPhoenix.Script.DropTableViews()
		'objPhoenix.Script.CreateTables()
		'objPhoenix.Script.CreateTableViews()
		'objPhoenix.Script.CreateViews()
		'objPhoenix.Script.ApplySecurity()


		bOK = objPhoenix.Script.ScriptOvernightStep2

		objPhoenix.CloseSafely()

				'    objPhoenix.ReturnErrorLog.Add(HRProEngine.ErrorHandler.Section.General, "hello", HRProEngine.ErrorHandler.Severity.Error, _
				'"SQLCode_AddCodeLevel", " -- Missing calculation")

				'    If objPhoenix.ReturnErrorLog.Count > 0 Then

				'      objPhoenix.ReturnErrorLog.Show()

				'    End If

		'    objPhoenix.ReturnErrorLog.OutputToFile("c:\dev\errors.txt")

		'   objPhoenix.ReturnTuningLog.OutputToFile("c:\dev\HR Pro\HR Pro Engine\Tuning.log")


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

#End Region

#Region "Scripted Updates"

		Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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
						.CommandText = "EXECUTE spadmin_gettables"										'Specify stored procedure to run
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

		Private Sub butViewObjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butViewObjects.Click

		'    Dim objHRPro As New SystemFramework.HCM
		'    Dim objViewer As New ViewObjects

		'    '    Dim SQLDB As New HRProEngine.Connectivity.SQL
		'    Dim objLogin As SystemFramework.Connectivity.Login

		'    With objLogin
		'        .UseContext = False
		'        .UserName = txtUser2.Text
		'        .Password = txtPassword2.Text
		'        .Database = txtDatabase2.Text
		'        .Server = txtServer2.Text
		'    End With


		'    objHRPro.Connect(objLogin)

		'    objHRPro.Initialise()
		'    objHRPro.PopulateObjects()

		'objViewer.Things = objHRPro.ReturnThings
		'objViewer.ShowDialog()

		'    objHRPro.Disconnect()



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

		Private Sub butImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butImport.Click

		'Dim objHRPro As New SystemFramework.HCM
		'Dim objImport As New ViewObjects

		''    Dim SQLDB As New HRProEngine.Connectivity.SQL
		'Dim objLogin As SystemFramework.Connectivity.Login

		'With objLogin
		'    .UseContext = False
		'    .UserName = txtUser2.Text
		'    .Password = txtPassword2.Text
		'    .Database = txtDatabase2.Text
		'    .Server = txtServer2.Text
		'End With

		'objHRPro.Connect(objLogin)

		'objHRPro.Initialise()
		'objHRPro.PopulateObjects()

		'objImport.Things = objHRPro.ReturnThings
		'objImport.ShowDialog()

		'objHRPro.Disconnect()

		End Sub

#End Region


	Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

		Dim objSysMgr As New SystemFramework.SysMgr

		txtNewKey.Text = objSysMgr.UpdateLicence(txtOldKey.Text)



	End Sub
End Class
