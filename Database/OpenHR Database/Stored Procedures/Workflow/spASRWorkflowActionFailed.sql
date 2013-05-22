CREATE PROCEDURE [dbo].[spASRWorkflowActionFailed]
(
	@piInstanceID		integer,
	@piElementID		integer,
	@psMessage			varchar(MAX)
)
AS
BEGIN
	DECLARE
		@iFailureFlows	integer,
		@iCount			integer;

	-- Check if the failed element has an outbound flow for failures.
	SELECT @iFailureFlows = COUNT(*)
	FROM ASRSysWorkflowElements Es
	INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
		AND Ls.startOutboundFlowCode = 1
	WHERE Es.ID = @piElementID
		AND Es.type = 5; -- 5 = StoredData

	IF @iFailureFlows = 0
	BEGIN
		UPDATE ASRSysWorkflowInstanceSteps
		SET status = 4,	-- 4 = failed
			message = @psMessage,
			failedCount = isnull(failedCount, 0) + 1
		WHERE instanceID = @piInstanceID
			AND elementID = @piElementID;

		UPDATE ASRSysWorkflowInstances
		SET status = 2	-- 2 = error
		WHERE ID = @piInstanceID;
	END
	ELSE
	BEGIN
		UPDATE ASRSysWorkflowInstanceSteps
		SET status = 8,	-- 8 = failed action
			message = @psMessage,
			failedCount = isnull(failedCount, 0) + 1
		WHERE instanceID = @piInstanceID
			AND elementID = @piElementID;

		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.status = 1,
			ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
			ASRSysWorkflowInstanceSteps.completionDateTime = null
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceSteps.elementID IN 
				(SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 1))
			AND (ASRSysWorkflowInstanceSteps.status = 0
				OR ASRSysWorkflowInstanceSteps.status = 3
				OR ASRSysWorkflowInstanceSteps.status = 4
				OR ASRSysWorkflowInstanceSteps.status = 6
				OR ASRSysWorkflowInstanceSteps.status = 8);
						
		-- Set activated Web Forms to be 'pending' (to be done by the user) 
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.status = 2
		WHERE ASRSysWorkflowInstanceSteps.id IN (
			SELECT ASRSysWorkflowInstanceSteps.ID
			FROM ASRSysWorkflowInstanceSteps
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowInstanceSteps.status = 1
				AND ASRSysWorkflowElements.type = 2);
						
		-- Set activated Terminators to be 'completed'
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.status = 3,
			ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
			ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
		WHERE ASRSysWorkflowInstanceSteps.id IN (
			SELECT ASRSysWorkflowInstanceSteps.ID
			FROM ASRSysWorkflowInstanceSteps
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowInstanceSteps.status = 1
				AND ASRSysWorkflowElements.type = 1);
						
		-- Count how many terminators have completed. ie. if the workflow has completed.
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowInstanceSteps
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceSteps.status = 3
			AND ASRSysWorkflowElements.type = 1;
											
		IF @iCount > 0 
		BEGIN
			UPDATE ASRSysWorkflowInstances
			SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
				ASRSysWorkflowInstances.status = 3
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
			
			/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
			is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
		END
	END
END

