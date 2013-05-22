CREATE PROCEDURE [dbo].[spASRWorkflowStepDescription]
(
	@piInstanceStepID	integer,
	@psDescription		varchar(MAX)	OUTPUT
)
AS
BEGIN
	DECLARE
		@iInstanceID			integer,
		@iExprID				integer,
		@iResultType			integer,
		@sResult				varchar(MAX),
		@fResult				bit,
		@dtResult				datetime,
		@fltResult				float,
		@fDescHasWorkflowName	bit,
		@fDescHasElementCaption	bit,
		@sWorkflowName			varchar(MAX),
		@sElementCaption		varchar(MAX);

	-- Get the InstanceID and associated DescriptionExprID of the given step
	SELECT @iInstanceID = isnull(WIS.instanceID, 0),
		@iExprID = isnull(WEs.descriptionExprID, 0),
		@fDescHasWorkflowName = isnull(WEs.descHasWorkflowName, 0),
		@fDescHasElementCaption = isnull(WEs.descHasElementCaption, 0),
		@sWorkflowName = isnull(Ws.name, ''),
		@sElementCaption = isnull(WEs.caption, '')
	FROM ASRSysWorkflowInstanceSteps WIS
	INNER JOIN ASRSysWorkflowElements WEs ON WIS.elementID = WEs.ID
	INNER JOIN ASRSysWorkflows Ws ON WEs.workflowID = Ws.ID
	WHERE WIS.ID = @piInstanceStepID;

	IF @iExprID > 0
	BEGIN
		EXEC [dbo].[spASRSysWorkflowCalculation]
			@iInstanceID,
			@iExprID,
			@iResultType OUTPUT,
			@sResult OUTPUT,
			@fResult OUTPUT,
			@dtResult OUTPUT,
			@fltResult OUTPUT, 
			0;
	END

	IF @fDescHasWorkflowName = 1
	BEGIN
		SET @sResult = @sWorkflowName 
			+ ' - '
			+ isnull(@sResult, '');
	END

	IF @fDescHasElementCaption = 1
	BEGIN
		SET @sResult = @sElementCaption 
			+ ' - '
			+ isnull(@sResult, '');
	END

	SELECT @psDescription = isnull(@sResult, '');
END
