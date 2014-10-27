CREATE PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]
(
	@piElementItemID	integer,
	@piInstanceID		integer,
	@piSize				integer			OUTPUT,
	@psFileName			varchar(MAX)	OUTPUT
)
AS
BEGIN
	DECLARE
		@iElementID		integer,
		@sIdentifier	varchar(MAX) 

	SELECT 			
		@piSize = ISNULL(ASRSysWorkflowElementItems.InputSize, 0),
		@iElementID = elementID,
		@sIdentifier = identifier
	FROM ASRSysWorkflowElementItems
	WHERE ASRSysWorkflowElementItems.ID = @piElementItemID;

	SELECT @psFileName = [TempFileUpload_Filename]
	FROM ASRSysWorkflowInstanceValues
	WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceValues.elementID = @iElementID
		AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;

	SELECT ASRSysWorkflowElementItemValues.value
	FROM ASRSysWorkflowElementItemValues
	WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID;
END