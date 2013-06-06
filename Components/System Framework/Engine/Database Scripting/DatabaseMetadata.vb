Imports System.Data.SqlClient

Public Class DatabaseMetadata

   Public Shared Function GetFunctions() As IList(Of ScriptedMetadata)

      Dim sql = "SELECT o.name, m.definition" & vbNewLine &
         "FROM sys.objects o" & vbNewLine &
         "INNER JOIN sys.schemas s ON s.schema_id = o.schema_id" & vbNewLine &
         "INNER JOIN sys.sql_modules m ON m.object_id = o.object_id" & vbNewLine &
         "WHERE o.type IN ('FN', 'TF') AND s.name = 'dbo'"

      Dim ds As DataSet = CType(CommitDB, Connectivity.ADOClassic).ExecSql(sql)

      Dim items = ds.Tables(0).Rows.Cast(Of DataRow)().Select(Function(r) New ScriptedMetadata With {
                   .Name = CStr(r(0)),
                   .Definition = CStr(r(1))}
                ).ToList()

      Return items
   End Function

   Public Shared Function GetTriggers() As IList(Of ScriptedMetadata)

      Dim sql = "SELECT o.name, m.definition" & vbNewLine &
         "FROM sys.objects o" & vbNewLine &
         "INNER JOIN sys.schemas s ON s.schema_id = o.schema_id" & vbNewLine &
         "INNER JOIN sys.sql_modules m ON m.object_id = o.object_id" & vbNewLine &
         "WHERE o.type = 'TR' AND s.name = 'dbo'"

      Dim ds As DataSet = CType(CommitDB, Connectivity.ADOClassic).ExecSql(sql)

      Dim items = ds.Tables(0).Rows.Cast(Of DataRow)().Select(Function(r) New ScriptedMetadata With {
                   .Name = CStr(r(0)),
                   .Definition = CStr(r(1))}
                ).ToList()

      Return items
   End Function

End Class

Public Class ScriptedMetadata
   Public Name As String
   Public Definition As String
End Class
