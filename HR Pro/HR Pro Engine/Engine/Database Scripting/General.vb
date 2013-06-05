Namespace ScriptDB

  <HideModuleName()> _
  Public Module General


    ' Dummy uDF can be replaced at a later time - this is just to get things running for evaluation purposes.
    Public Function CreateUDF(ByRef [Role] As String, ByRef [ObjectName] As String, ByRef [BodyCode] As String, ByRef [SafeDummyUDF] As String) As Boolean

      Dim sSQL As String = String.Empty

      Try
        ScriptDB.DropUDF([Role], [ObjectName])
        sSQL = String.Format("CREATE FUNCTION [{0}].[{1}] {2}", [Role], [ObjectName], [BodyCode], vbNewLine)

        ' Commit
        CommitDB.ScriptStatement(sSQL)


      Catch ex As Exception
        Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.UDFs, [ObjectName], HRProEngine.ErrorHandler.Severity.Error, ex.Message, sSQL)

        ' This didn't work, so put a note in the error log and create a dummy UDF
        sSQL = String.Format("CREATE FUNCTION [{0}].[{1}] {2}" _
          , [Role], [ObjectName], [SafeDummyUDF], vbNewLine)

        CommitDB.ScriptStatement(sSQL)

        Return False

      End Try

      Return True
    End Function

    Public Function DropUDF(ByRef [Role] As String, ByRef [ObjectName] As String) As Boolean

      Dim sSQL As String = String.Empty

      Try

        sSQL = String.Format("IF EXISTS(SELECT o.[name] FROM sys.sysobjects o " & _
          "INNER JOIN sys.sysusers u ON o.[uid] = u.[uid] " & _
          "WHERE o.[name] = '{1}' AND [type] IN ('FN', 'TF') AND u.[name] = '{0}')" & vbNewLine & _
          " DROP FUNCTION [{0}].[{1}]", [Role], [ObjectName])

        ' Commit
        CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.UDFs, [ObjectName], HRProEngine.ErrorHandler.Severity.Error, ex.Message, sSQL)
        Return False

      End Try

      Return True

    End Function

  End Module
End Namespace
