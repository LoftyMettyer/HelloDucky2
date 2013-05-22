CREATE PROCEDURE [dbo].[spASRIntGetEventLogDetails] (
	@piBatchRunID	integer,
	@piEventID		integer,
	@piExists		integer		OUTPUT
)
AS
BEGIN
	
	DECLARE @sSelectString			varchar(MAX),
			@sFromString			varchar(255),
			@sWhereString			varchar(MAX),
			@sOrderString 			varchar(MAX),
			@sTempExecString		nvarchar(MAX),
			@sTempParamDefinition	nvarchar(500),
			@iCount					integer;

	/****************************************************************************************************************************************/
	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(*) FROM ASRSysEventLog WHERE ID = ' + convert(varchar,@piEventID);

	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;
	SET @piExists = @iCount;
	/****************************************************************************************************************************************/

	SET @sSelectString = '';
	SET @sFromString = '';
	SET @sWhereString = '';
	SET @sOrderString = '';

	/* create SELECT statment string */
	SET @sSelectString = 'SELECT [ID], [EventLogID], IsNull([Notes],'''') AS ''Notes'' ';

	SET @sFromString = ' FROM AsrSysEventLogDetails ';

	IF @piBatchRunID > 0
		BEGIN
			SET @sWhereString = ' WHERE AsrSysEventLogDetails.EventLogID IN (SELECT ID FROM ASRSysEventLog WHERE BatchRunID = ' + convert(varchar, @piBatchRunID) + ')';
		END
	ELSE
		BEGIN
			SET @sWhereString = ' WHERE AsrSysEventLogDetails.EventLogID = ' + convert(varchar, @piEventID);
		END

	SET @sOrderString = ' ORDER BY AsrSysEventLogDetails.[ID] ';
	
	SET @sTempExecString = @sSelectString + @sFromString + @sWhereString + @sOrderString;
	EXEC sp_executesql @sTempExecString;
	
END