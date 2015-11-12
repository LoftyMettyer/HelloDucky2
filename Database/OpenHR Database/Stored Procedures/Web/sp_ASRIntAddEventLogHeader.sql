CREATE PROCEDURE [dbo].[sp_ASRIntAddEventLogHeader]
(
    @piNewRecordID	integer OUTPUT,   /* Output variable to hold the new record ID. */
    @piType			integer,
    @psName			varchar(150), 
    @psUserName		varchar(50),
    @psBatchName	varchar(50),
    @piBatchRunID	integer,
    @piBatchJobID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@outputTable table (id int NOT NULL);

	INSERT INTO [dbo].[ASRSysEventLog] (
		[DateTime],	[Type],	[Name], [Status], [Username],
		[Mode], [BatchName], [SuccessCount], [FailCount], [BatchRunID], [BatchJobID])
	OUTPUT inserted.ID INTO @outputTable
	VALUES (GETDATE(), @piType, @psName, 0, @psUserName,
		CASE
			WHEN len(@psBatchName) = 0 THEN 0
			ELSE 1
		END,    
		@psBatchName, NULL,NULL,
		CASE
			WHEN @piBatchRunID > 0 THEN @piBatchRunID
			ELSE null
		END,
		CASE 
			WHEN @piBatchJobID > 0 THEN @piBatchJobID
			ELSE null
		END);
                  
    -- Get the ID of the inserted record.
	SELECT @piNewRecordID = id FROM @outputTable;

END
