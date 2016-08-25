CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
(
	@psTableName varchar(255),
	@psFieldName varchar(255),
	@piBatches int
) 
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @lowid		integer, 
			@maxid		integer,
			@rowcount	integer,
			@start		datetime;

	DECLARE @sSQL				nvarchar(MAX),
			@sParamDefinition	nvarchar(500);

	-- Determine the number of ID's we'll update in each batch
	IF ISNULL(@piBatches, 0) = 0
		SET @piBatches = 2000;
	
	SET @sSQL = 'SELECT @lowid = ISNULL(MIN(ID),0),  @maxid = ISNULL(MAX(ID),0) FROM ' + @psTableName;
	SET @sParamDefinition = N'@lowid int OUTPUT, @maxid int OUTPUT';
	EXEC sp_executesql @sSQL, @sParamDefinition, @lowid OUTPUT, @maxid OUTPUT;

	WHILE 1=1
	BEGIN
		SET @start = GETDATE();
		
		-- Do the update
		SELECT @sSQL = 'UPDATE ' + @psTableName + ' SET ' + @psFieldName + ' = ' + @psFieldName
					+ ' WHERE ID BETWEEN ' + CONVERT(nvarchar(10), @lowid) + ' AND ' + CONVERT(varchar(10),  @lowid + @piBatches - 1);
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

	SET NOCOUNT OFF;
END
