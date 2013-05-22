CREATE PROCEDURE [dbo].[spASRIntDeleteEventLogRecords]
(
		@piDeleteType			integer,
		@psSelectedEventIDs		varchar(MAX),
		@pfCanViewAll			bit
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sSQL	nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psSelectedEventIDs) > 0 SET @psSelectedEventIDs = replace(@psSelectedEventIDs, '''', '''''');

	IF (@piDeleteType = 0) OR (@piDeleteType = 1)
	BEGIN
		/* 0 = Delete all the selected rows */
		/* 1 = Delete all the rows shown */
		SET @sSQL = 'DELETE FROM ASRSysEventLog' +
			' WHERE ID IN (' + @psSelectedEventIDs + ')';
		EXEC sp_executesql @sSQL;
	END
	
	IF @piDeleteType = 2
	BEGIN
		/* Delete all the records the user has permission to see */
		IF @pfCanViewAll = 1
		BEGIN
			DELETE FROM [dbo].[ASRSysEventLog];
		END
		ELSE
		BEGIN
			DELETE FROM [dbo].[ASRSysEventLog] 
			WHERE username = SYSTEM_USER;
		END
	END
END