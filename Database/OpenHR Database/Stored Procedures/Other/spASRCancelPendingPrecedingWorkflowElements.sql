CREATE PROCEDURE dbo.spASRCancelPendingPrecedingWorkflowElements
			(
				@piInstanceID			integer,
				@piElementID			integer
			)
			AS
			BEGIN
				/* Cancel (ie. set status to 0 for all workflow pending (ie. status 1 or 2) elements that precede the given element.
				This ignores connection elements.
				NB. This does work for elements with multiple inbound flows. */
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 0
				WHERE instanceID = @piInstanceID
					AND elementID IN (SELECT ID FROM [dbo].[udfASRGetAllPrecedingWorkflowElements](@piElementID))
					AND status IN (1, 2, 7) -- 1 = pending engine action, 2 = pending user action, 7 = pending user completion
			END


