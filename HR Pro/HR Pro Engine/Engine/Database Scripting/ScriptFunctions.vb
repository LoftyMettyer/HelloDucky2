﻿Option Explicit On

Namespace ScriptDB

  Public Module ScriptFunctions

#Region "Numeric Functions"

#End Region

    Public Function ConvertCurrency() As Boolean

      Dim bOK As Boolean = True
      Dim sSQL As String = vbNullString
      Dim objConversionTable As Things.Table
      Dim objNameColumn As Things.Column
      Dim objValueColumn As Things.Column
      Dim objDecimalsColumn As Things.Column
      Dim iKeyID As HCMGuid
      Dim sObjectName As String

      Try

        objConversionTable = Globals.ModuleSetup.GetSetting("MODULE_CURRENCY", "Param_ConversionTable").Table

        If Not objConversionTable Is Nothing Then
          iKeyID = Globals.ModuleSetup.GetSetting("MODULE_CURRENCY", "Param_CurrencyNameColumn").Value
          objNameColumn = objConversionTable.Column(iKeyID)

          iKeyID = Globals.ModuleSetup.GetSetting("MODULE_CURRENCY", "Param_ConversionValueColumn").Value
          objValueColumn = objConversionTable.Column(iKeyID)

          iKeyID = Globals.ModuleSetup.GetSetting("MODULE_CURRENCY", "Param_DecimalColumn").Value
          objDecimalsColumn = objConversionTable.Column(iKeyID)

          sObjectName = "udfsys_convertcurrency"
          sSQL = String.Format("/* ------------------------------------------------- */" & vbNewLine & _
                  "/* HR Pro Currency module user defined function.                  */" & vbNewLine & _
                  "/* Automatically generated by the Advanced DB Scripting Engine.   */" & vbNewLine & _
                  "/* -------------------------------------------------------------- */" & vbNewLine & _
                  "CREATE FUNCTION dbo.[{0}] (" & vbNewLine & _
                  "   @currency numeric(38,8)," & vbNewLine & _
                  "   @from {5}," & vbNewLine & _
                  "   @to {5})" & vbNewLine & _
                  "RETURNS numeric(38,8)" & vbNewLine & _
                  "AS " & vbNewLine & _
                  "BEGIN" & vbNewLine & vbNewLine & _
                  "    DECLARE @result numeric(38,8);" & vbNewLine & _
                  "    SELECT @result = ROUND(ISNULL((@currency / NULLIF((SELECT [{2}] FROM dbo.[{1}] WHERE [{3}] = @from),0))" & vbNewLine &
                  "                      * " & vbNewLine & _
                  "                     (SELECT [{2}] FROM dbo.[{1}] WHERE [{3}] = @to), 0)," & vbNewLine & _
                  "                         ISNULL((SELECT [{4}] FROM dbo.[{1}] WHERE [{3}] = @to), 0))" & vbNewLine & vbNewLine & _
                  "    RETURN ISNULL(@result,0);" & vbNewLine & _
                  "END", sObjectName, objConversionTable.Name, objValueColumn.Name, objNameColumn.Name, objDecimalsColumn.Name _
                    , objNameColumn.DataTypeSyntax)
          ScriptDB.DropUDF("dbo", sObjectName)
          bOK = CommitDB.ScriptStatement(sSQL)

        End If

      Catch ex As Exception
        bOK = False

      End Try


      Return bOK

    End Function

    Public Function UniqueCodeViews() As Boolean

      Dim bOK As Boolean = True
      Dim sSQL As String = vbNullString

      Try

        'sSQL = "spadmin_generateuniquecodes"
        'bOK = CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        bOK = False
      End Try

      Return bOK

    End Function

    Public Function GetFieldFromDatabases() As Boolean

      Dim bOK As Boolean = True
      Dim sObjectName As String = "udfsys_getfieldfromdatabaserecord"
      Dim sSQL As String = vbNullString
      Dim sStatement As String = vbNullString
      Dim aryStatements As New ArrayList
      Dim objComponent As Things.Component
      Dim objPart1 As Things.Component
      Dim objPart3 As Things.Component
      Dim objTable1 As Things.Table
      Dim objTable2 As Things.Table
      Dim objIndex As Things.Index

      Try

        For Each objComponent In Globals.GetFieldsFromDB

          ' The parameters as down a couple of component levels 
          objPart1 = CType(objComponent.Objects(0).Objects(0), Things.Component)
          objPart3 = CType(objComponent.Objects(2).Objects(0), Things.Component)
          objTable1 = Globals.Things.Table(objPart1.TableID)
          objTable2 = Globals.Things.Table(objPart3.TableID)

          ' Even though the user can select different table for parameters 1 and 3 this
          ' would return garbage data so ignore it!
          If objTable1 Is objTable2 Then
            sStatement = String.Format("    IF @searchcolumnid = '{0}-{1}' AND @returncolumnid = '{2}-{3}'" & vbNewLine & _
              "        BEGIN" & vbNewLine & _
              "            SELECT @result = [{5}] FROM dbo.[{4}] WHERE [{6}] = @searchexpression;" & vbNewLine & _
              "            RETURN @result;" & vbNewLine & _
              "        END" & vbNewLine _
            , objPart1.TableID.PadLeft, objPart1.ColumnID.PadLeft _
            , objPart3.TableID.PadLeft, objPart3.ColumnID.PadLeft _
            , objTable1.PhysicalName, objTable1.Column(objPart3.ColumnID).Name _
            , objTable2.Column(objPart1.ColumnID).Name)

            ' Only add if not already done so
            If Not aryStatements.Contains(sStatement) Then

              aryStatements.Add(sStatement)

              ' Put an index on this column
              objIndex = New Things.Index
              objIndex.Name = String.Format("IDX_getfromdb_{0}_{1}", objTable1.Column(objPart1.ColumnID).Name, objTable2.Column(objPart3.ColumnID).Name)
              objIndex.Columns.Add(objTable1.Column(objPart1.ColumnID))
              objIndex.IncludedColumns.Add(objTable2.Column(objPart3.ColumnID))
              objIndex.IncludePrimaryKey = False
              objIndex.IsTableIndex = True

              objTable1.Objects.Add(objIndex)
            End If
          End If
        Next

        ' Build the stored procedure
        sSQL = String.Format("CREATE FUNCTION [dbo].[{0}](" & vbNewLine &
            "    @searchcolumnid AS varchar(17)," & vbNewLine & _
            "    @searchexpression AS nvarchar(MAX)," & vbNewLine & _
            "    @returncolumnid AS varchar(17))" & vbNewLine & _
            "RETURNS nvarchar(MAX)" & vbNewLine & _
            "--WITH SCHEMABINDING" & vbNewLine & _
            "AS" & vbNewLine & "BEGIN" & vbNewLine & _
            "    DECLARE @result nvarchar(MAX);" & vbNewLine & _
            "    SET @result = '';" & vbNewLine & vbNewLine & _
            "{1}" & vbNewLine & _
            "    RETURN @result;" & vbNewLine & _
            "END" _
            , sObjectName, String.Join(vbNewLine, aryStatements.ToArray()))

        ScriptDB.DropUDF("dbo", sObjectName)
        bOK = CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.Triggers, sObjectName, HRProEngine.ErrorHandler.Severity.Error, ex.Message, sSQL)
        bOK = False
      End Try

      Return bOK

    End Function

    'Public Function GeneratePerformanceIndexes() As Boolean

    '  Dim bOK As Boolean = True
    '  Dim objColumn As Things.Column

    '  Try
    '    For Each objColumn In Globals.PerformanceIndexes

    '      ScriptIndex(objColumn, False, True)
    '    Next

    '  Catch ex As Exception
    '    bOK = False

    '  End Try

    '  Return bOK

    'End Function

    'Public Function ScriptIndex(ByRef Column As Things.Column, ByVal Clustered As Boolean, ByVal IncludeForeignKey As Boolean) As Boolean

    '  Dim bOK As Boolean
    '  Dim sSQL As String = String.Empty
    '  Dim sColumns As String
    '  Dim objRelation As Things.Relation

    '  Try

    '    sColumns = Column.Name

    '    If IncludeForeignKey Then
    '      For Each objRelation In Column.Table.Objects(Things.Type.Relation)
    '        If objRelation.RelationshipType = RelationshipType.Parent Then
    '          sColumns = sColumns & ", ID_" & CInt(objRelation.ParentID)
    '        End If
    '      Next
    '    End If



    '    ' Create the new index
    '    sSQL = String.Format("IF EXISTS(SELECT [id] FROM sysindexes WHERE [name] = 'IDX_{1}_{0}')" & _
    '        " DROP INDEX [IDX_{1}_{0}] ON [dbo].[{3}];" & _
    '        " CREATE NONCLUSTERED INDEX [IDX_{1}_{0}] ON [dbo].[{3}] ({2});" & vbNewLine _
    '        , Column.Name, Column.Table.Name, sColumns, Column.Table.PhysicalName)
    '    Globals.CommitDB.ScriptStatement(sSQL)

    '  Catch ex As Exception
    '    Globals.ErrorLog.Add(HRProEngine.ErrorHandler.Section.Views, Column.Table.Name & "--" & Column.Name, HRProEngine.ErrorHandler.Severity.Error, ex.Message, sSQL)
    '    bOK = False

    '  End Try

    '  Return bOK

    'End Function

  End Module


End Namespace
