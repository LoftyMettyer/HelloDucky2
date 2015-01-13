CREATE PROCEDURE dbo.spASRCopyChildRecords(
	@iParentTableID integer,
	@iNewRecordID integer,
	@iOriginalRecordID integer)
WITH EXECUTE AS OWNER
AS
BEGIN

	DECLARE @sqlCopyData nvarchar(MAX) = '';
	DECLARE @childDataColumns TABLE (TableID integer, TableName nvarchar(255), ColumnNames nvarchar(MAX));

	INSERT @childDataColumns (TableID, TableName, ColumnNames)	
		SELECT DISTINCT r.ParentID, t.tablename, d.StringValues
			FROM ASRSysRelations r
			INNER JOIN ASRSysTables t ON t.tableid = r.ChildID
			INNER JOIN ASRSysColumns c ON c.tableid = t.tableid
			CROSS APPLY ( SELECT ', ' + columnname
							FROM ASRSysColumns v2
							WHERE v2.tableid = c.tableid AND v2.datatype <> 4
								FOR XML PATH('') )  d ( StringValues )
			WHERE r.ParentID = @iParentTableID AND t.CopyWhenParentRecordIsCopied = 1;

	SELECT @sqlCopyData = @sqlCopyData + 'INSERT ' + TableName + '(ID_' + CONVERT(varchar(10), TableID) +  ColumnNames + ') SELECT ' 
		+ CONVERT(varchar(10), @iNewRecordID) + ColumnNames 
		+ ' FROM ' + TableName + ' WHERE ID_' + CONVERT(varchar(10), TableID) + ' = ' + CONVERT(varchar(10), @iOriginalRecordID) + ';' + CHAR(13)
		FROM @childDataColumns;

	EXECUTE sp_executesql @sqlCopyData;

END
