Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports Assembly.Classes
Imports Microsoft.SqlServer.Server
Imports System.Collections.Generic
Imports System.Linq
Imports System.Collections.ObjectModel

Partial Public Class Functions

   Public Shared Numbers As New Collection(Of UniqueCode)

   <SqlProcedure(Name:="spstat_flushuniquecode")>
   Public Shared Sub FlushUniqueCode()

      Dim UniqueCode As UniqueCode
      Dim cmd As SqlCommand
      Dim command As String
      Dim objectName As String

      Using conn As New SqlConnection("context connection=true")
         conn.Open()

         Try

            For Each UniqueCode In Numbers.Where(Function(u) u.IsNew = True)

               objectName = String.Format("sequence_{0}", UniqueCode.Code)
               command = String.Format("IF NOT EXISTS (SELECT * FROM sys.sequences WHERE name = N'{0}') " &
                     "BEGIN " &
                     "CREATE SEQUENCE [{0}] START WITH {1}; " &
                     "GRANT UPDATE ON [{0}] TO ASRSysGroup; " &
                     "SELECT NEXT VALUE FOR {0}; " & 
                     "END", objectName, UniqueCode.Value)

               cmd = New SqlCommand(command, conn)
               cmd.ExecuteNonQuery()

            Next

         Catch ex As Exception

         End Try

      End Using

      Numbers.Clear()

   End Sub

   <SqlFunction(Name:="udfstat_getuniquecode", DataAccess:=DataAccessKind.Read)> _
   Public Shared Function GetUniqueCode(Prefix As String, RootValue As Long, RecordID As Integer) As SqlTypes.SqlString

      Dim returnVal As Long
      Dim UniqueCode As UniqueCode
      Dim bFound As Boolean = False

      Try

         Dim code As UniqueCode = Numbers.FirstOrDefault(Function(n) n.Code = Prefix And n.LastRecordID = RecordID)

         If code Is Nothing
            code = New UniqueCode With {
               .Code = Prefix,
               .LastRecordID = RecordID,
               .Value = GetNextSequence(Prefix, RootValue),
               .IsNew = True}
            Numbers.Add(code)

         End If

         returnVal = code.Value

      Catch ex As Exception
         Return ex.Message
      End Try

      Return returnVal.ToString

   End Function

   Private Shared Function GetNextSequence(Prefix As String, RootValue As Long) As Long

      Dim command As String
      Dim objectName As String
      Dim cmd As SqlCommand
      Dim nextValue As Long = RootValue

      Try

         Using conn As New SqlConnection("context connection=true")
            conn.Open()

            objectName = String.Format("sequence_{0}", Prefix)

            command = String.Format("IF OBJECT_ID('{0}') IS NOT NULL " & _
               "SELECT NEXT VALUE FOR {0} ELSE SELECT {1}", objectName, RootValue)
            cmd = New SqlCommand(command, conn)
            nextValue = CType(cmd.ExecuteScalar(), Long)

         End Using

      Catch ex As Exception
         Return RootValue

      End Try

      Return nextValue

   End Function

End Class

