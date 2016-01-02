CREATE PROCEDURE [dbo].[spASRIntGetEventLogBatchDetails] (
	@piBatchRunID 	integer,
	@piEventID		integer
)
AS
BEGIN

	SET NOCOUNT ON;
	
	DECLARE @sExecString		nvarchar(MAX),
			@sSelectString 		varchar(MAX),
			@sFromString		varchar(MAX),
			@sWhereString		varchar(MAX),
			@sOrderString 		varchar(MAX);

	SET @sSelectString = '';
	SET @sFromString = '';
	SET @sWhereString = '';
	SET @sOrderString = '';

	/* create SELECT statment string */
	SET @sSelectString = 'SELECT 
		 ID, 
		 DateTime,
		 EndTime,
		 IsNull(Duration,-1) AS Duration,
		 Username,
		 CASE Type 
						WHEN 0 THEN ''Unknown''
						WHEN 1 THEN ''Cross Tab'' 
						WHEN 2 THEN ''Custom Report'' 
						WHEN 3 THEN ''Data Transfer'' 
						WHEN 4 THEN ''Export'' 
						WHEN 5 THEN ''Global Add'' 
						WHEN 6 THEN ''Global Delete'' 
						WHEN 7 THEN ''Global Update'' 
						WHEN 8 THEN ''Import'' 
						WHEN 9 THEN ''Mail Merge'' 
						WHEN 10 THEN ''Diary Delete'' 
						WHEN 11 THEN ''Diary Rebuild''
						WHEN 12 THEN ''Email Rebuild''
						WHEN 13 THEN ''Standard Report''
						WHEN 14 THEN ''Record Editing''
						WHEN 15 THEN ''System Error''
						WHEN 16 THEN ''Match Report''
						WHEN 17 THEN ''Calendar Report''
						WHEN 18 THEN ''Envelopes & Labels''
						WHEN 19 THEN ''Label Definition''
						WHEN 20 THEN ''Record Profile''
						WHEN 21	THEN ''Succession Planning''
						WHEN 22 THEN ''Career Progression''
						WHEN 25 THEN ''Workflow Rebuild''
						WHEN 35 THEN ''9-Box Grid Report''
						WHEN 38 THEN ''Talent Report''
						ELSE ''Unknown''  
		 END AS Type,
		 Name,
		 CASE Mode 
			WHEN 1 THEN ''Batch''
			WHEN 0 THEN ''Manual''
			ELSE ''Unknown''
		 END AS Mode, 
		 CASE Status 
				WHEN 0 THEN ''Pending''
		   	WHEN 1 THEN ''Cancelled'' 
				WHEN 2 THEN ''Failed'' 
				WHEN 3 THEN ''Successful'' 
				WHEN 4 THEN ''Skipped'' 
				WHEN 5 THEN ''Error''
				ELSE ''Unknown'' 
		 END AS Status,
		 IsNull(BatchName,'''') AS BatchName,
		 IsNull(convert(varchar,SuccessCount), ''N/A'') AS SuccessCount,
		 IsNull(convert(varchar,FailCount), ''N/A'') AS FailCount,
		 IsNull(convert(varchar,BatchJobID), ''N/A'') AS BatchJobID,
		 IsNull(convert(varchar,BatchRunID), ''N/A'') AS BatchRunID';

	SET @sFromString = ' FROM ASRSysEventLog ';

	IF @piBatchRunID > 0
		BEGIN
			SET @sWhereString = ' WHERE BatchRunID = ' + convert(varchar, @piBatchRunID);
		END
	ELSE
		BEGIN
			SET @sWhereString = ' WHERE ID = ' + convert(varchar, @piEventID);
		END

	SET @sOrderString = ' ORDER BY DateTime ASC ';

	SET @sExecString = @sSelectString + @sFromString + @sWhereString + @sOrderString;

	-- Run generated statement
	EXEC sp_executeSQL @sExecString;
	
END