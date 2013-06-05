Imports System.Data.SqlClient

Public Class DatabaseMetadata

   Public Shared Function GetFunctions() As IList(Of FunctionMetadata)

      Dim sql = "SELECT o.name, m.definition, m.is_schema_bound" & vbNewLine &
         "FROM sys.sysobjects o" & vbNewLine &
         "INNER JOIN sys.sysusers u ON o.uid = u.uid" & vbNewLine &
         "INNER JOIN sys.sql_modules m ON m.object_id = o.id" & vbNewLine &
         "WHERE o.type IN ('FN', 'TF') AND u.name = 'dbo'"

      Dim ds As DataSet = CType(CommitDB, Connectivity.ADOClassic).ExecSql(sql)

      Dim items = ds.Tables(0).Rows.Cast(Of DataRow)().Select(Function(r) New FunctionMetadata With {
                   .Name = CStr(r(0)),
                   .Definition = CStr(r(1)),
                   .IsSchemaBound = CBool(r(2))}
                ).ToList()

      Return items
   End Function

End Class

Public Class FunctionMetadata
   Public Name As String
   Public Definition As String
   Public IsSchemaBound As Boolean
End Class
