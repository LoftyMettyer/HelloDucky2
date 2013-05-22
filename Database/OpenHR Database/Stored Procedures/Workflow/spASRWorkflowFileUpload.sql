CREATE PROCEDURE [dbo].[spASRWorkflowFileUpload]
(
	@piElementItemID	integer,
	@piInstanceID		integer,
	@pimgFile			image,
	@psContentType		varchar(MAX),
	@psFileName			varchar(MAX),
	@pfClear			bit
)
AS
BEGIN
	DECLARE	@iElementID		integer,
			@sIdentifier	varchar(MAX);

	SELECT
		@iElementID = elementID,
		@sIdentifier = identifier
	FROM ASRSysWorkflowElementItems
	WHERE id = @piElementItemID;

	UPDATE ASRSysWorkflowInstanceValues 
	SET [TempFileUpload_File] = 
			CASE 
				WHEN @pfClear = 1 THEN null
				ELSE @pimgFile
			END, 
		[TempFileUpload_ContentType] = 
			CASE 
				WHEN @pfClear = 1 THEN null
				ELSE @psContentType
			END, 
		[TempFileUpload_Filename] = 
			CASE 
				WHEN @pfCLear = 1 THEN null
				ELSE @psFileName
			END
	WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceValues.elementID = @iElementID
		AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;

END