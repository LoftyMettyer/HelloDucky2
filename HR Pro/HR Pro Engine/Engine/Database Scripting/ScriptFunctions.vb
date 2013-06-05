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
      Dim iKeyID As Integer
      Dim sObjectName As String

      Try

        objConversionTable = Globals.ModuleSetup.Setting("MODULE_CURRENCY", "Param_ConversionTable").Table

        If Not objConversionTable Is Nothing Then
          iKeyID = CInt(Globals.ModuleSetup.Setting("MODULE_CURRENCY", "Param_CurrencyNameColumn").Value)
          objNameColumn = objConversionTable.Columns.GetById(iKeyID)

          iKeyID = CInt(Globals.ModuleSetup.Setting("MODULE_CURRENCY", "Param_ConversionValueColumn").Value)
          objValueColumn = objConversionTable.Columns.GetById(iKeyID)

          iKeyID = CInt(Globals.ModuleSetup.Setting("MODULE_CURRENCY", "Param_DecimalColumn").Value)
          objDecimalsColumn = objConversionTable.Columns.GetById(iKeyID)

          sObjectName = "udfsys_convertcurrency"
          If Not objValueColumn Is Nothing And Not objDecimalsColumn Is Nothing Then
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
      Dim objIndex As New Things.Index
      Dim sVariableName As String
      Dim sSearchExpression As String
      Dim objColumn As Things.Column
      Dim bFound As Boolean

      Try

        For Each objComponent In Globals.GetFieldsFromDB

          ' The parameters as down a couple of component levels 
          objPart1 = objComponent.Components(0).Components(0)
          objPart3 = objComponent.Components(2).Components(0)
          objTable1 = Globals.Tables.GetById(objPart1.TableID)
          objTable2 = Globals.Tables.GetById(objPart3.TableID)

          objColumn = objTable1.Columns.GetById(objPart3.ColumnID)
          If Not objColumn Is Nothing Then
            Select Case objColumn.DataType
              Case ColumnTypes.Date
                sVariableName = "@result_date"
              Case ColumnTypes.Numeric
                sVariableName = "@result_numeric"
              Case ColumnTypes.Logic
                sVariableName = "@result_boolean"
              Case ColumnTypes.Integer
                sVariableName = "@result_integer"
              Case Else
                sVariableName = "@result_string"
            End Select

            ' Make integers typesafe
            Select Case objTable1.Columns.GetById(objPart1.ColumnID).DataType
              Case ColumnTypes.Integer
                sSearchExpression = "convert(numeric(38,8), @searchexpression)"
              Case Else
                sSearchExpression = "@searchexpression"
            End Select

            ' Even though the user can select different table for parameters 1 and 3 this
            ' would return garbage data so ignore it!
            If objTable1 Is objTable2 Then
              sStatement = String.Format("    IF @searchcolumnid = '{0}-{1}' AND @returncolumnid = '{2}-{3}'" & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            SELECT {7} = [{5}] FROM dbo.[{4}] WHERE [{6}] = {8};" & vbNewLine & _
                "            RETURN {7};" & vbNewLine & _
                "        END" & vbNewLine _
              , objPart1.TableID.ToString.PadLeft(8, "0"c), objPart1.ColumnID.ToString.PadLeft(8, "0"c) _
              , objPart3.TableID.ToString.PadLeft(8, "0"c), objPart3.ColumnID.ToString.PadLeft(8, "0"c) _
              , objTable1.PhysicalName, objTable1.Columns.GetById(objPart3.ColumnID).Name _
              , objTable2.Columns.GetById(objPart1.ColumnID).Name, sVariableName, sSearchExpression)

              ' Only add if not already done so
              If Not aryStatements.Contains(sStatement) Then

                aryStatements.Add(sStatement)

                ' Put an index on this column
                bFound = False
                For Each objIndex In objTable1.Indexes
                  If objIndex.Columns.Count > 0 Then
                    If objIndex.Columns(0) Is objTable1.Columns.GetById(objPart1.ColumnID) Then
                      bFound = True
                      Exit For
                    End If
                  End If
                Next

                If Not bFound Then
                  objIndex = New Things.Index
                  objIndex.Name = String.Format("IDX_getfromdb_{0}", objTable1.Columns.GetById(objPart1.ColumnID).Name)
                  objIndex.Columns.Add(objTable1.Columns.GetById(objPart1.ColumnID))
                  objIndex.IncludePrimaryKey = False
                  objIndex.IsTableIndex = True
                End If

                objIndex.IncludedColumns.Add(objTable2.Columns.GetById(objPart3.ColumnID))

                If Not bFound And Not objIndex.Columns(0).Multiline Then
                  objTable1.Indexes.Add(objIndex)
                End If

              End If
            End If
          End If
        Next

        ' Build the stored procedure
        sSQL = String.Format("CREATE FUNCTION [dbo].[{0}](" & vbNewLine &
            "    @searchcolumnid AS varchar(17)," & vbNewLine & _
            "    @searchexpression AS varchar(255)," & vbNewLine & _
            "    @returncolumnid AS varchar(17))" & vbNewLine & _
            "RETURNS sql_variant" & vbNewLine & _
            "AS" & vbNewLine & "BEGIN" & vbNewLine & _
            "    DECLARE @result_string     varchar(255)," & vbNewLine & _
            "            @result_numeric    numeric(38,8)," & vbNewLine & _
            "            @result_integer    integer," & vbNewLine & _
            "            @result_boolean    bit," & vbNewLine & _
            "            @result_date       datetime;" & vbNewLine & vbNewLine & _
            "    SET @result_string = '';" & vbNewLine & _
            "    SET @result_integer = 0;" & vbNewLine & _
            "    SET @result_numeric = 0;" & vbNewLine & vbNewLine & _
            "{1}" & vbNewLine & _
            "    RETURN NULL;" & vbNewLine & _
            "END" _
            , sObjectName, String.Join(vbNewLine, aryStatements.ToArray()))

        ScriptDB.DropUDF("dbo", sObjectName)
        bOK = CommitDB.ScriptStatement(sSQL)

      Catch ex As Exception
        Globals.ErrorLog.Add(SystemFramework.ErrorHandler.Section.Triggers, sObjectName, SystemFramework.ErrorHandler.Severity.Error, ex.Message, sSQL)
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

    'Public Function BankHolidayUpdate() As Boolean

    '  Dim bOK As Boolean = True
    '  Dim sSQL As String = vbNullString
    '  Dim objBankHolidayTable As Things.Table
    '  Dim objHolidayDateColumn As Things.Column
    '  Dim objUpdateColumn As Things.Column
    '  Dim iKeyID As Integer
    '  Dim sObjectName As String
    '  Dim aryUpdates As New ArrayList

    '  Try

    '    objBankHolidayTable = Globals.ModuleSetup.Setting("MODULE_ABSENCE", "Param_TableBHol").Table
    '    If Not objBankHolidayTable Is Nothing Then
    '      iKeyID = Globals.ModuleSetup.Setting("MODULE_ABSENCE", "Param_FieldBHolDate").Value
    '      objHolidayDateColumn = objBankHolidayTable.Column(iKeyID)
    '    End If

    '    For Each objUpdateColumn In Globals.OnBankHolidayUpdate
    '      aryUpdates.Add(String.Format(" --   UPDATE dbo.[{0}] SET [{1}] = [{1}] WHERE @bankholidaydate BETWEEN [{2}] AND [{3}]" _
    '        , objUpdateColumn.Table.PhysicalName, objUpdateColumn.Name, "start_date", "end_date"))
    '    Next

    '    sObjectName = "spsys_updatebankholiday"
    '    sSQL = String.Format("/* -------------------------------------------------------------- */" & vbNewLine & _
    '            "/* HR Pro Bank Holiday Module.                                    */" & vbNewLine & _
    '            "/* Automatically generated by the Advanced DB Scripting Engine.   */" & vbNewLine & _
    '            "/* -------------------------------------------------------------- */" & vbNewLine & _
    '            "CREATE PROCEDURE dbo.[{0}] (@bankholidaydate datetime)" & vbNewLine & _
    '            "AS " & vbNewLine & _
    '            "BEGIN" & vbNewLine & vbNewLine & _
    '            "    DECLARE @icount integer;" & vbNewLine & vbNewLine & _
    '            "{1}" & vbNewLine & _
    '            "END", sObjectName, String.Join(vbNewLine, aryUpdates.ToArray()))
    '    ScriptDB.DropProcedure("dbo", sObjectName)
    '    bOK = CommitDB.ScriptStatement(sSQL)


    '  Catch ex As Exception
    '    bOK = False

    '  End Try

    '  Return bOK

    'End Function

  End Module


End Namespace
