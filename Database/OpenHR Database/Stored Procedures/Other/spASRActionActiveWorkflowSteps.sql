CREATE PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
AS
BEGIN
	-- Return a recordset of the workflow steps that need to be actioned by the Workflow service.
	-- Action any that can be actioned immediately. 
	DECLARE
		@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to '2', 3 = Summing Junction check, 4 = Or check
		@iElementType		integer,
		@iInstanceID		integer,
		@iElementID			integer,
		@iStepID			integer,
		@iCount				integer,
		@sStatus			bit,
		@sMessage			varchar(MAX),
		@iTemp				integer, 
		@iTemp2				integer, 
		@iTemp3				integer,
		@sForms 			varchar(MAX), 
		@iType				integer,
		@iDecisionFlow		integer,
		@fInvalidElements	bit, 
		@fValidElements		bit, 
		@iPrecedingElementID	integer, 
		@iPrecedingElementType	integer, 
		@iPrecedingElementStatus	integer, 
		@iPrecedingElementFlow	integer, 
		@fSaveForLater			bit;

	DECLARE stepsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT TOP 5 E.type,
		S.instanceID,
		E.ID,
		S.ID
	FROM ASRSysWorkflowInstanceSteps S
	INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
	WHERE S.status = 1
		AND E.type <> 5 -- 5 = StoredData elements handled in the service
	ORDER BY s.ActivationDateTime;

	OPEN stepsCursor;
	FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @iAction = 
			CASE
				WHEN @iElementType = 1 THEN 1	-- Terminator
				WHEN @iElementType = 2 THEN 2	-- Web form (action required from user)
				WHEN @iElementType = 3 THEN 1	-- Email
				WHEN @iElementType = 4 THEN 1	-- Decision
				WHEN @iElementType = 6 THEN 3	-- Summing Junction
				WHEN @iElementType = 7 THEN 4	-- Or	
				WHEN @iElementType = 8 THEN 1	-- Connector 1
				WHEN @iElementType = 9 THEN 1	-- Connector 2
				ELSE 0					-- Unknown
			END;
		
		IF @iAction = 3 -- Summing Junction check
		BEGIN
			-- Check if all preceding steps have completed before submitting this step.
			SET @fInvalidElements = 0;	
		
			DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT WE.ID,
				WE.type,
				WIS.status,
				WIS.decisionFlow
			FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
			INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
			INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
				AND WIS.instanceID = @iInstanceID;

			OPEN precedingElementsCursor;			
			FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;

			WHILE (@@fetch_status = 0)
				AND (@fInvalidElements = 0)
			BEGIN
				IF (@iPrecedingElementType = 4) -- Decision
				BEGIN
					IF @iPrecedingElementStatus = 3 -- 3 = completed
					BEGIN
						SELECT @iCount = COUNT(*) 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
						WHERE ID = @iElementID;

						IF @iCount = 0 SET @fInvalidElements = 1;
					END
					ELSE
					BEGIN
						SET @fInvalidElements = 1;
					END
				END
				ELSE
				BEGIN
					IF (@iPrecedingElementType = 2) -- WebForm
					BEGIN
						IF @iPrecedingElementStatus = 3 -- 3 = completed
							OR @iPrecedingElementStatus = 6 -- 6 = timeout
						BEGIN
							SET @iTemp3 = CASE
									WHEN @iPrecedingElementStatus = 3 THEN 0
									ELSE 1
								END;

							SELECT @iCount = COUNT(*)
							FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
							WHERE ID = @iElementID;
						
							IF @iCount = 0 SET @fInvalidElements = 1;
						END
						ELSE
						BEGIN
							SET @fInvalidElements = 1;
						END
					END
					ELSE
					BEGIN
						IF (@iPrecedingElementType = 5) -- StoredData
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
								OR @iPrecedingElementStatus = 8 -- 8 = failed action
							BEGIN
								SET @iTemp3 = CASE
										WHEN @iPrecedingElementStatus = 3 THEN 0
										ELSE 1
									END;

								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
								WHERE ID = @iElementID;
							
								IF @iCount = 0 SET @fInvalidElements = 1;
							END
							ELSE
							BEGIN
								SET @fInvalidElements = 1;
							END
						END
						ELSE
						BEGIN
							-- Preceding element must have status 3 (3 =Completed)
							IF @iPrecedingElementStatus <> 3 SET @fInvalidElements = 1;
						END
					END
				END

				FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
			END
			CLOSE precedingElementsCursor;
			DEALLOCATE precedingElementsCursor;
			
			IF (@fInvalidElements = 0) SET @iAction = 1;
		END

		IF @iAction = 4 -- Or check
		BEGIN
			SET @fValidElements = 0;
			-- Check if any preceding steps have completed before submitting this step. 

			DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT WE.ID,
				WE.type,
				WIS.status,
				WIS.decisionFlow
			FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
			INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
			INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
				AND WIS.instanceID = @iInstanceID;

			OPEN precedingElementsCursor;	

			FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
			WHILE (@@fetch_status = 0)
				AND (@fValidElements = 0)
			BEGIN
				IF (@iPrecedingElementType = 4) -- Decision
				BEGIN
					IF @iPrecedingElementStatus = 3 -- 3 = completed
					BEGIN
						SELECT @iCount = COUNT(*)
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
						WHERE ID = @iElementID;
					
						IF @iCount > 0 SET @fValidElements = 1;
					END
				END
				ELSE
				BEGIN
					IF (@iPrecedingElementType = 2) -- WebForm
					BEGIN
						IF @iPrecedingElementStatus = 3 -- 3 = completed
							OR @iPrecedingElementStatus = 6 -- 6 = timeout
						BEGIN
							SET @iTemp3 = CASE
									WHEN @iPrecedingElementStatus = 3 THEN 0
									ELSE 1
								END;

							SELECT @iCount = COUNT(*)
							FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
							WHERE ID = @iElementID;
					
							IF @iCount > 0 SET @fValidElements = 1;
						END
					END
					ELSE
					BEGIN
						IF (@iPrecedingElementType = 5) -- StoredData
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
								OR @iPrecedingElementStatus = 8 -- 8 = failed action
							BEGIN
								SET @iTemp3 = CASE
										WHEN @iPrecedingElementStatus = 3 THEN 0
										ELSE 1
									END;

								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
								WHERE ID = @iElementID;

								IF @iCount > 0 SET @fValidElements = 1;
							END
						END
						ELSE
						BEGIN
							-- Preceding element must have status 3 (3 =Completed)
							IF @iPrecedingElementStatus = 3 SET @fValidElements = 1;
						END
					END
				END

				FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
			END
			CLOSE precedingElementsCursor;
			DEALLOCATE precedingElementsCursor;

			-- If all preceding steps have been completed submit the Or step.
			IF @fValidElements > 0 
			BEGIN
				-- Cancel any preceding steps that are not completed as they are no longer required.
				EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @iInstanceID, @iElementID;

				SET @iAction = 1;
			END
		END

		IF @iAction = 1
		BEGIN
			EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '', @sForms OUTPUT, @fSaveForLater OUTPUT, 0;
		END

		IF @iAction = 2
		BEGIN
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET status = 2
			WHERE id = @iStepID;
		END

		FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
	END

	CLOSE stepsCursor;
	DEALLOCATE stepsCursor;

	DECLARE timeoutCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
		WIS.instanceID,
		WE.ID,
		WIS.ID
	FROM ASRSysWorkflowInstanceSteps WIS
	INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
		AND WE.type = 2 -- WebForm
	WHERE ((WIS.status = 2) OR (WIS.status = 7)) -- Pending user action/completion
		AND isnull(WE.timeoutFrequency,0) > 0
		AND CASE 
				WHEN WE.timeoutPeriod = 0 THEN 
					dateadd(minute, WE.timeoutFrequency, WIS.activationDateTime)
				WHEN WE.timeoutPeriod = 1 THEN 
					dateadd(hour, WE.timeoutFrequency, WIS.activationDateTime)
				WHEN WE.timeoutPeriod = 2 AND WE.timeoutExcludeWeekend = 1 THEN 
					dbo.udfASRAddWeekdays(WIS.activationDateTime, WE.timeoutFrequency)
				WHEN WE.timeoutPeriod = 2 THEN 
					dateadd(day, WE.timeoutFrequency, WIS.activationDateTime)
				WHEN WE.timeoutPeriod = 3 THEN 
					dateadd(week, WE.timeoutFrequency, WIS.activationDateTime)
				WHEN WE.timeoutPeriod = 4 THEN 
					dateadd(month, WE.timeoutFrequency, WIS.activationDateTime)
				WHEN WE.timeoutPeriod = 5 THEN 
					dateadd(year, WE.timeoutFrequency, WIS.activationDateTime)
				ELSE getDate()
			END <= getDate();	

	OPEN timeoutCursor;
	FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
	WHILE (@@fetch_status = 0)
	BEGIN
		-- Set the step status to be Timeout
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.status = 6, -- Timeout
			ASRSysWorkflowInstanceSteps.timeoutCount = isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0) + 1
		WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID;

		-- Activate the succeeding elements on the Timeout flow
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.status = 1,
			ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
			ASRSysWorkflowInstanceSteps.completionDateTime = null
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
			AND ASRSysWorkflowInstanceSteps.elementID IN 
				(SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1))
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
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
			AND ASRSysWorkflowInstanceSteps.status = 3
			AND ASRSysWorkflowElements.type = 1;
								
		IF @iCount > 0 
		BEGIN
			UPDATE ASRSysWorkflowInstances
			SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
				ASRSysWorkflowInstances.status = 3
			WHERE ASRSysWorkflowInstances.ID = @iInstanceID;
			
			-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
			-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table.
		END

		FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
	END

	CLOSE timeoutCursor;
	DEALLOCATE timeoutCursor;
END