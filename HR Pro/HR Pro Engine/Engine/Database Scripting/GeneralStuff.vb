Option Explicit On

Namespace ScriptDB

  <HideModuleName()> _
  Public Module General

    Public Function GetSQLColumnDatatype(ByVal Type As ScriptDB.ComponentValueTypes) As String

      Dim sSQLType As String = String.Empty

      Select Case CInt(Type)
        Case ComponentValueTypes.String, ComponentValueTypes.Component_String
          sSQLType = "[varchar](MAX)"

        Case ComponentValueTypes.Numeric, ComponentValueTypes.Component_Numeric
          sSQLType = String.Format("[numeric](38,8)")

        Case ComponentValueTypes.Date, ComponentValueTypes.Component_Date
          sSQLType = "[datetime]"

        Case ComponentValueTypes.Logic, ComponentValueTypes.Component_Logic
          sSQLType = "[bit]"

      End Select

      GetSQLColumnDatatype = sSQLType

    End Function

  End Module
End Namespace
