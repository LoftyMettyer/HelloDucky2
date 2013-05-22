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
  INSERT INTO [dbo].[ASRSysEventLog] (
		[DateTime],	[Type],	[Name], [Status], [Username],
		[Mode], [BatchName], [SuccessCount], [FailCount], [BatchRunID], [BatchJobID])
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
    SELECT @piNewRecordID = MAX(id) FROM [dbo].[ASRSysEventLog];

END
