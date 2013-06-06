Attribute VB_Name = "modDocumentManagement"
Option Explicit
Const SP_DOCUMENTLINKSP = "spASRLinkDocument"

Public Function CreateLinkDocumentSP() As Boolean

  Dim bOK As Boolean
  Dim sSQL As String
  Dim sProcCode As String
  Dim sTypeTableName As String
  Dim sCategoryTableName As String
  Dim sTypeColumnName As String
  Dim sCategoryColumnName As String
  Dim sTypeCategoryColumnName As String
      
  ' Drop existing stuff
  bOK = DropProcedure(SP_DOCUMENTLINKSP)

  ' Module defined columns
  sCategoryTableName = GetTableName(GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYTABLE, 0))
  sCategoryColumnName = GetColumnName(GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN, 0), True)
  sTypeTableName = GetTableName(GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPETABLE, 0))
  sTypeCategoryColumnName = GetColumnName(GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN, 0), True)
  sTypeColumnName = GetColumnName(GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECOLUMN, 0), True)
  
  If Len(sCategoryTableName) > 0 And Len(sCategoryColumnName) > 0 And Len(sTypeTableName) > 0 And Len(sTypeCategoryColumnName) > 0 And Len(sTypeColumnName) > 0 Then
  
    ' Funky new procedure
    sSQL = "/* ---------------------------------------------------- */" & vbNewLine & _
              "/* HR Pro Document Management module stored procedure.          */" & vbNewLine & _
              "/* Automatically generated by the System manager.   */" & vbNewLine & _
              "/* ---------------------------------------------------- */" & vbNewLine & _
              "CREATE PROCEDURE dbo.[spASRLinkDocument](" & vbNewLine & _
              "    @DocumentCategory nvarchar(255)," & vbNewLine & _
              "    @DocumentType   nvarchar(255)," & vbNewLine & _
              "    @Parent1Key     nvarchar(100)," & vbNewLine & _
              "    @Parent2Key     nvarchar(100)," & vbNewLine & _
              "    @Key            nvarchar(100)," & vbNewLine & _
              "    @DocumentGuid   uniqueidentifier," & vbNewLine & _
              "    @ToDelete       bit," & vbNewLine & _
              "    @Link           nvarchar(MAX))" & vbNewLine & _
              "AS" & vbNewLine & _
              "BEGIN" & vbNewLine & _
              "    SET NOCOUNT ON;" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    DECLARE @iCount           integer," & vbNewLine & _
              "        @sExecuteSQL          nvarchar(MAX)," & vbNewLine & _
              "        @sWhereClause         nvarchar(MAX)," & vbNewLine & _
              "        @sInsertColumns       nvarchar(MAX)," & vbNewLine & _
              "        @sInsertValues        nvarchar(MAX)," & vbNewLine & _
              "        @sGetParentID         nvarchar(MAX)," & vbNewLine & _
              "        @sParamDefinition     nvarchar(MAX);" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    DECLARE @sTableName       nvarchar(255)," & vbNewLine & _
              "        @sTargetKeyField      nvarchar(255)," & vbNewLine & _
              "        @sTargetLinkColumn    nvarchar(255)," & vbNewLine & _
              "        @sTargetCategoryColumn  nvarchar(255)," & vbNewLine & _
              "        @sTargetTypeColumn    nvarchar(255)," & vbNewLine & _
              "        @sParent1Keyfield      nvarchar(255)," & vbNewLine & _
              "        @sParent1TableName     nvarchar(255)," & vbNewLine & _
              "        @intParent1TableID     integer," & vbNewLine & _
              "        @intParent1RecordID    integer;" & vbNewLine & vbNewLine
    
    sSQL = sSQL & _
              "    DECLARE @iDocCategoryID integer," & vbNewLine & _
              "        @iDocTypeID   integer;" & vbNewLine & vbNewLine

    sSQL = sSQL & _
              "    SET @sExecuteSQL = '';" & vbNewLine & _
              "    SET @sWhereClause = '';" & vbNewLine & _
              "    SET @sInsertColumns = '';" & vbNewLine & _
              "    SET @sInsertValues = '';" & vbNewLine & _
              "    SET @intParent1RecordID = 0;" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    SELECT @iDocCategoryID = [ID] FROM [" & sCategoryTableName & "] WHERE [" & sCategoryColumnName & "] = @DocumentCategory;" & vbNewLine & _
              "    SELECT @iDocTypeID = [ID] FROM [" & sTypeTableName & "] WHERE [" & sTypeCategoryColumnName & "] = @DocumentCategory AND [" & sTypeColumnName & "] = @DocumentType;" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    SELECT @sTableName = ISNULL(t.[TableName],''), @sTargetLinkColumn = ISNULL(c1.[ColumnName],''), @sTargetKeyField = ISNULL(c2.[ColumnName],'')" & vbNewLine & _
              "       , @intParent1TableID = ISNULL(d.[Parent1TableID],''), @sParent1KeyField = ISNULL(c3.[ColumnName],''), @sParent1TableName = ISNULL(pt.[TableName],'')" & vbNewLine & _
              "       , @sTargetCategoryColumn  = ISNULL(c4.[ColumnName],''), @sTargetTypeColumn  = ISNULL(c5.[ColumnName],'') " & vbNewLine & _
              "    FROM dbo.[ASRSysDocumentManagementTypes] d" & vbNewLine & _
              "        INNER JOIN [ASRSysTables] t ON t.[TableID] = d.[TargetTableID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysTables] pt ON pt.[TableID] = d.[Parent1TableID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysColumns] c1 ON c1.[ColumnID] = d.[TargetColumnID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysColumns] c2 ON c2.[ColumnID] = d.[TargetKeyFieldColumnID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysColumns] c3 ON c3.[ColumnID] = d.[Parent1KeyFieldColumnID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysColumns] c4 ON c4.[ColumnID] = d.[TargetCategoryColumnID]" & vbNewLine & _
              "        LEFT JOIN [ASRSysColumns] c5 ON c5.[ColumnID] = d.[TargetTypeColumnID]" & vbNewLine & _
              "        WHERE d.[CategoryRecordID] = @iDocCategoryID AND (d.[TypeRecordID] = @iDocTypeID OR @iDocTypeID IS NULL);" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    IF LEN(@sTargetKeyField) > 0" & vbNewLine & _
              "    BEGIN" & vbNewLine & _
              "        SET @sWhereClause = '[' + @sTargetKeyField + '] = ''' + @Key + ''''" & vbNewLine & _
              "        IF LEN(@sTargetTypeColumn) > 0 SET @sWhereClause = @sWhereClause + 'AND [' + @sTargetTypeColumn + '] = ''' + @DocumentType + '''';" & vbNewLine & vbNewLine & _
              "        IF LEN(@sParent1Keyfield) > 0" & vbNewLine & _
              "        BEGIN" & vbNewLine & _
              "            SET @sGetParentID = 'SELECT @intParent1RecordID = [id] FROM dbo.[' + @sParent1TableName + '] WHERE [' + @sParent1KeyField + '] = ''' + @Parent1Key + '''';" & vbNewLine & _
              "            SET @sParamDefinition = N'@intParent1RecordID integer output';" & vbNewLine & _
              "            EXECUTE sp_executeSQL @sGetParentID, @sParamDefinition, @intParent1RecordID output;" & vbNewLine & _
              "            SET @sWhereClause = @sWhereClause + ' AND [ID_' + convert(nvarchar(10),@intParent1TableID) + '] = ' + convert(nvarchar(10),@intParent1RecordID);" & vbNewLine & _
              "            SET @sInsertColumns = '[ID_' + convert(nvarchar(10),@intParent1TableID) + '], '" & vbNewLine & _
              "            SET @sInsertValues = convert(nvarchar(10),@intParent1RecordID) + ', ' " & vbNewLine & _
              "        END" & vbNewLine & _
              "    END" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    IF LEN(@sTargetCategoryColumn) > 0" & vbNewLine & _
              "    BEGIN" & vbNewLine & _
              "      SET @sInsertColumns = @sInsertColumns + @sTargetCategoryColumn + ', '" & vbNewLine & _
              "      SET @sInsertValues = @sInsertValues + '''' + @DocumentCategory + ''', '" & vbNewLine & _
              "    END" & vbNewLine & vbNewLine

    sSQL = sSQL & _
              "    IF LEN(@sTargetTypeColumn) > 0" & vbNewLine & _
              "    BEGIN" & vbNewLine & _
              "      SET @sInsertColumns = @sInsertColumns + @sTargetTypeColumn + ', '" & vbNewLine & _
              "      SET @sInsertValues = @sInsertValues + '''' + @DocumentType + ''', '" & vbNewLine & _
              "    END" & vbNewLine & vbNewLine

    sSQL = sSQL & _
              "    SET @sInsertColumns = @sInsertColumns + @sTargetKeyField + ', ' + @sTargetLinkColumn;" & vbNewLine & _
              "    SET @sInsertValues = @sInsertValues + '''' + @Key + ''', ''' + @Link + '''';" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
              "    SET @sExecuteSQL = 'IF EXISTS(SELECT [ID] FROM dbo.[' + @sTableName + '] WHERE ' + @sWhereClause + ') ' + CHAR(13) +" & vbNewLine & _
              "        'UPDATE dbo.[' + @sTableName + '] SET [' + @sTargetLinkColumn + '] = ''' + @Link + ''' WHERE ' + @sWhereClause + CHAR(13) +" & vbNewLine & _
              "        'ELSE ' + CHAR(13) + " & vbNewLine & _
              "        'INSERT dbo.[' + @sTableName + '](' + @sInsertColumns + ') VALUES (' + @sInsertValues + ')'" & vbNewLine
  
  
    sSQL = sSQL & _
              "    PRINT @sExecuteSQL;" & vbNewLine & _
              "    EXECUTE sp_executeSQL @sExecuteSQL;" & vbNewLine
  
    sSQL = sSQL & "END"
    
    gADOCon.Execute sSQL, , adExecuteNoRecords
    
  End If
  
  CreateLinkDocumentSP = bOK

End Function
