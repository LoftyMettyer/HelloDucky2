CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
(
	@piTableName varchar(255),
	@piFieldName varchar(255),
	@piBatches int
) 
AS
BEGIN
	SET NOCOUNT ON;

	-- Create the progress table if it doesn't already exist
	IF OBJECT_ID('ASRSysOvernightProgress', N'U') IS NULL
		CREATE TABLE ASRSysOvernightProgress
			(TableName varchar(255)
			, RecCount int
			, IDRange varchar(255)
			, StartDate datetime
			, EndDate datetime
			, DurationMins int);

	DECLARE @lowid int,@highid int,@maxid int;
	DECLARE @rowcount int, @start datetime;

	DECLARE @sSQL				nvarchar(MAX),
			@sParamDefinition	nvarchar(500);

	-- Determine the number of ID's we'll update in each batch
	IF ISNULL(@piBatches, 0) = 0
		SET @piBatches = 2000;
	SET @lowid = 0 ;
	SET @highid = @lowid + @piBatches;
	
	SET @sSQL = 'SELECT @maxid = ISNULL(MAX(ID),0) FROM ' + @piTableName;
	SET @sParamDefinition = N'@maxid int OUTPUT';
	EXEC sp_executesql @sSQL, @sParamDefinition, @maxid OUTPUT;

	WHILE 1=1
	BEGIN
		SET @start = GETDATE();
		
		-- Do the update
		SELECT @sSQL = 'UPDATE ' + @piTableName + ' SET ' + @piFieldName + ' = ' + @piFieldName
					+ ' WHERE ID BETWEEN @lowid AND @highid';
		SET @sParamDefinition = N'@lowid int, @highid int';
		EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @highid;

		SET @rowcount = @@ROWCOUNT;

		-- insert a record to this progress table to check the progress
		INSERT INTO ASRSysOvernightProgress 
			SELECT @piTableName
				, @rowcount
				, CAST(@lowid as varchar(255)) + '-' + CAST(@highid as varchar(255))
				, @start
				, GETDATE()
				, DATEDIFF(n, @start, GETDATE());

		SET @lowid = @lowid + @piBatches;
		SET @highid = @lowid + @piBatches;

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