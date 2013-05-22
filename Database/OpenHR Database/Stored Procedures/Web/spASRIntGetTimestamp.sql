CREATE PROCEDURE [dbo].[spASRIntGetTimestamp] (
	@piTimestamp 	int 		OUTPUT, 
	@piRecordID		integer,
	@psRealsource	varchar(255)
)
AS
BEGIN
	DECLARE @sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500);

	/* Clean the input string parameters. */
	IF len(@psRealsource) > 0 SET @psRealsource = replace(@psRealsource, '''', '''''');
	
	SET @sTempExecString = 'SELECT @iTimestamp = convert(integer, timestamp) FROM ' + convert(nvarchar(255), @psRealsource) + ' where ID = ' + convert(nvarchar(100), @piRecordID);
	SET @sTempParamDefinition = N'@iTimestamp integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @piTimestamp OUTPUT;
END