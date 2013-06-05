'Imports System.IO
'Imports System.Xml.Serialization

Public Class MainForm

    'Private objProgress As Phoenix.HCMProgressBar
    Private Initialised As Boolean = False

#Region "Scripting stuff (some unused)"

    Private Sub butScriptDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butScriptDB.Click

        Dim objPhoenix As New SystemFramework.SysMgr

        Dim sPath As String = "C:\dev\HR Pro\HR Pro Engine\Test Harness\"

        Dim objADO As New ADODB.Connection
        Dim objDAOEngine As New DAO.DBEngine
        Dim objDAODB As DAO.Database
        '    Dim objRecordset As DAO.Recordset
        Dim sADOConnect As String = String.Format("Driver=SQL Server;Server={0};UID=sa;PWD=asr;Database={1};" _
                              , txtServer.Text, txtDatabase.Text)
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
        objPhoenix.PopulateObjects()



        objPhoenix.Options.DevelopmentMode = chkDebugMode.Checked
        objPhoenix.Options.RefreshObjects = True
        bOK = objPhoenix.Script.CreateObjects()


        ' bOK = objPhoenix.Script.CreateTableViews
        '    bOK = objPhoenix.Script.CreateViews

        bOK = objPhoenix.Script.CreateTriggers()


        'objPhoenix.Script.DropViews()
        'objPhoenix.Script.DropTableViews()
        'objPhoenix.Script.CreateTables()

        'objPhoenix.Script.CreateTableViews()
        'objPhoenix.Script.CreateViews()
        'objPhoenix.Script.ApplySecurity()

        bOK = objPhoenix.Script.CreateFunctions

        bOK = objPhoenix.Script.ScriptIndexes

        bOK = objPhoenix.Script.ScriptOvernightStep2

        objPhoenix.CloseSafely()



        '    objPhoenix.ReturnErrorLog.Add(HRProEngine.ErrorHandler.Section.General, "hello", HRProEngine.ErrorHandler.Severity.Error, _
        '"SQLCode_AddCodeLevel", " -- Missing calculation")

        '    If objPhoenix.ReturnErrorLog.Count > 0 Then

        '      objPhoenix.ReturnErrorLog.Show()

        '    End If



        objPhoenix.ReturnErrorLog.OutputToFile("c:\dev\errors.txt")

        objPhoenix.ReturnTuningLog.OutputToFile("c:\dev\HR Pro\HR Pro Engine\Tuning.log")


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

    'Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '  Dim objTable1 As HRProEngine.Things.Table
    '  Dim objTable2 As HRProEngine.Things.Table
    '  Dim objColumn As HRProEngine.Things.Column
    '  Dim objColumn2 As HRProEngine.Things.Column
    '  Dim objExpresssion1 As HRProEngine.Things.Expression
    '  Dim objExpresssion2 As HRProEngine.Things.Expression
    '  Dim objComponent1 As HRProEngine.Things.Component

    '  objComponent1 = New HRProEngine.Things.Component

    '  objExpresssion1 = New HRProEngine.Things.Expression
    '  objExpresssion1.Objects.Add(objComponent1)

    '  objColumn = New HRProEngine.Things.Column
    '  objColumn.Name = "column1"

    '  objColumn2 = New HRProEngine.Things.Column
    '  objColumn2.Name = "column1"

    '  objTable1 = New HRProEngine.Things.Table
    '  objTable1.Name = "myatble1"

    '  objTable1.Objects.Add(objColumn)
    '  objTable1.Objects.Add(objExpresssion1)

    '  objExpresssion2 = objExpresssion1.DeepClone
    '  objExpresssion2.Name = "copy expr"

    '  objExpresssion2.Objects.Clear()
    '  objExpresssion2.Objects.Add(objComponent1)
    '  objExpresssion2.Objects.Add(objComponent1)
    '  objExpresssion2.Objects.Add(objComponent1)

    '  objTable2 = objTable1.DeepClone

    '  'objTable2 = objTable1 '.DeepClone

    '  objTable2.Name = "mycloned"

    '  '    objTable1.Objects.Add(objColumn2)
    '  objTable2.Objects(0).Name = "changed"
    '  objTable2.Objects.Add(objExpresssion2)

    '  'MsgBox(objTable1.Objects(0).Name)
    '  'MsgBox(objTable2.Objects(0).Name)
    '  'MsgBox(objTable1.Objects.Count)
    '  'MsgBox(objTable2.Objects.Count)


    '  MsgBox(objExpresssion1.Objects.Count)
    '  MsgBox(objExpresssion2.Objects.Count)


    'End Sub

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

    Private Sub butViewObjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butViewObjects.Click

        Dim objHRPro As New SystemFramework.HCM
        Dim objViewer As New ViewObjects

        '    Dim SQLDB As New HRProEngine.Connectivity.SQL
        Dim objLogin As SystemFramework.Connectivity.Login

        With objLogin
            .UseContext = False
            .UserName = txtUser2.Text
            .Password = txtPassword2.Text
            .Database = txtDatabase2.Text
            .Server = txtServer2.Text
        End With

        objHRPro.Connect(objLogin)

        objHRPro.Initialise()
        objHRPro.PopulateObjects()

        objViewer.Things = objHRPro.ReturnThings
        objViewer.ShowDialog()

        objHRPro.Disconnect()



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

        Dim objHRPro As New SystemFramework.HCM
        Dim objImport As New ViewObjects

        '    Dim SQLDB As New HRProEngine.Connectivity.SQL
        Dim objLogin As SystemFramework.Connectivity.Login

        With objLogin
            .UseContext = False
            .UserName = txtUser2.Text
            .Password = txtPassword2.Text
            .Database = txtDatabase2.Text
            .Server = txtServer2.Text
        End With

        objHRPro.Connect(objLogin)

        objHRPro.Initialise()
        objHRPro.PopulateObjects()

        objImport.Things = objHRPro.ReturnThings
        objImport.ShowDialog()

        objHRPro.Disconnect()

    End Sub

#End Region



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Dim objHRPro As New SystemFramework.HCM
        '    Dim objAudit As New AuditLogForm

        'Dim objLogin As SystemFramework.Connectivity.Login

        'With objLogin
        '  .UseContext = False
        '  .UserName = txtUser2.Text
        '  .Password = txtPassword2.Text
        '  .Database = txtDatabase2.Text
        '  .Server = txtServer2.Text
        'End With

        'objHRPro.Connect(objLogin)

        '    objAudit.Database = objHRPro
        'objAudit.ShowDialog()

        'objHRPro.Disconnect()

        Dim con = String.Format("Initial Catalog={0}; Server={1};User ID={2}; Password={3}; APP={4};",
                             txtDatabase2.Text, txtServer2.Text, txtUser2.Text, txtPassword2.Text, "ScriptDB")

        Dim f As New AuditLogForm With {.ConString = con}
        f.ShowDialog()

    End Sub

End Class
