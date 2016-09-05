CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
(
   @psTableName varchar(255),
   @psFieldName varchar(255),
   @piBatches int = 1000,
   @psWhereClause varchar(MAX) = ''
) 
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @lowid		integer, 
         @maxid		integer,
         @rowcount	integer,
         @start		datetime;

	DECLARE @sSQL				nvarchar(MAX),
			@sParamDefinition	nvarchar(500),
         @disableIndexSQL nvarchar(MAX) = '';
	
	SET @sSQL = 'SELECT @lowid = ISNULL(MIN(ID),0),  @maxid = ISNULL(MAX(ID),0) FROM ' + @psTableName;
	SET @sParamDefinition = N'@lowid int OUTPUT, @maxid int OUTPUT';
	EXEC sp_executesql @sSQL, @sParamDefinition, @lowid OUTPUT, @maxid OUTPUT;

  	-- Disable table scalar table indexes
	SELECT @disableIndexSQL = @disableIndexSQL + 'ALTER INDEX [' + i.name + '] ON ' + t.name + ' DISABLE;' + CHAR(13)
		FROM sys.indexes i
		INNER JOIN sys.tables t ON i.object_id = T.object_id
		WHERE i.type_desc = 'NONCLUSTERED'
			AND i.name IS NOT NULL AND i.name LIKE 'IDX_udftab%' AND OBJECT_NAME(i.object_id) = @pstableName
	EXECUTE sp_executeSQL @disableIndexSQL;

	WHILE 1=1
	BEGIN
		SET @start = GETDATE();
		
		-- Do the update
		SELECT @sSQL = 'UPDATE ' + @psTableName + ' SET ' + @psFieldName + ' = ' + @psFieldName
					+ ' WHERE ID BETWEEN ' + CONVERT(nvarchar(10), @lowid) + ' AND ' + CONVERT(varchar(10),  @lowid + @piBatches - 1)
               + CASE WHEN LEN(@psWhereClause) > 0 THEN ' AND ' + @psWhereClause ELSE '' END
		EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @piBatches;

		SET @rowcount = @@ROWCOUNT;

		-- insert a record to this progress table to check the progress
		INSERT INTO ASRSysOvernightProgress (TableName, RecCount, IDRange, StartDate, EndDate, DurationSecs)
			SELECT @psTableName
				, @rowcount
				, CAST(@lowid as varchar(255)) + '-' + CAST(@lowid + @piBatches - 1 as varchar(255))
				, @start
				, GETDATE()
				, DATEDIFF(ss, @start, GETDATE());

		SET @lowid = @lowid + @piBatches;

		IF @lowid > @maxid
		BEGIN
			CHECKPOINT;
			BREAK;
		END
		ELSE
			CHECKPOINT;
	END

  	-- Rebuild table scalar table indexes
   SET @disableIndexSQL = '';
	SELECT @disableIndexSQL = @disableIndexSQL + 'ALTER INDEX [' + i.name + '] ON ' + t.name + ' REBUILD;' + CHAR(13)
		FROM sys.indexes i
		INNER JOIN sys.tables t ON i.object_id = T.object_id
		WHERE i.type_desc = 'NONCLUSTERED'
			AND i.name IS NOT NULL AND i.name LIKE 'IDX_udftab%' AND OBJECT_NAME(i.object_id) = @pstableName
	EXECUTE sp_executeSQL @disableIndexSQL;

	SET NOCOUNT OFF;
END
