CREATE PROCEDURE dbo.[spASRIntCopyRecordPostSave]  (
	@tableID		integer,
	@FromRecordID	integer,
	@ToRecordID		integer)
WITH EXECUTE AS OWNER
AS
BEGIN
	DECLARE @nvarCommand nvarchar(MAX) = '',
			@tablename		varchar(255) = '',
			@updateFields	varchar(MAX) = '',
			@readFields		varchar(MAX) = ''

	SELECT @updateFields = @updateFields + CASE WHEN LEN(@updateFields) > 0 THEN ', ' ELSE '' END + c.columnname + ' = newData.' + c.columnname,
		   @readFields = @readFields + CASE WHEN LEN(@readFields) > 0 THEN ', ' ELSE '' END + c.columnname 
	FROM ASRSysColumns c
	WHERE c.tableid = @tableID AND c.datatype IN (-3, -4);

	SELECT @tablename = t.tablename
		FROM ASRSysTables t
		WHERE t.tableid = @tableID;

	IF LEN(@updateFields) > 0
	BEGIN
		SET @nvarCommand = 'UPDATE ' + @tablename + ' SET ' + @updateFields 
			+ ' FROM (SELECT ' + @readFields + ' FROM ' + @tablename + ' WHERE ID = ' + convert(varchar(10), @FromRecordID)
			+ ') newdata WHERE ID = ' + convert(varchar(10), @ToRecordID);

		EXECUTE sp_executeSQL @nvarCommand;

	END
END
