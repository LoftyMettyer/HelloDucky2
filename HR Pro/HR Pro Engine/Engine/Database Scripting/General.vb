Namespace ScriptDB

  <HideModuleName()>
  Public Module General

    Public Function DropUDF(ByVal [Role] As String, ByVal [ObjectName] As String) As Boolean

      Dim sSQL As String = String.Empty

      Try

        sSQL = String.Format("IF EXISTS(SELECT o.[name] FROM sys.sysobjects o " & _
          "INNER JOIN sys.sysusers u ON o.[uid] = u.[uid] " & _
          "WHERE o.[name] = '{1}' AND [type] IN ('FN', 'TF') AND u.[name] = '{0}')" & vbNewLine & _
          " DROP FUNCTION [{0}].[{1}]", [Role], [ObjectName])

        ' Commit
        CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.UDFs, [ObjectName], SystemFramework.ErrorHandler.Severity.Error, ex.Message, sSQL)
        Return False

      End Try

      Return True

    End Function

    Public Function DropProcedure(ByVal [Role] As String, ByVal [ObjectName] As String) As Boolean

      Dim sSQL As String = String.Empty

      Try

        sSQL = String.Format("IF EXISTS(SELECT o.[name] FROM sys.sysobjects o " & _
          "INNER JOIN sys.sysusers u ON o.[uid] = u.[uid] " & _
          "WHERE o.[name] = '{1}' AND [type] = 'P' AND u.[name] = '{0}')" & vbNewLine & _
          " DROP PROCEDURE [{0}].[{1}]", [Role], [ObjectName])

        ' Commit
        CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.General, [ObjectName], SystemFramework.ErrorHandler.Severity.Error, ex.Message, sSQL)
        Return False

      End Try

      Return True

    End Function


  End Module
End Namespace
