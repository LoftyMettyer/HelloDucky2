CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE
				@iQueueID			integer,
				@iWorkflowID		integer,
				@iRecordID			integer,
				@iInstanceID		integer,
				@iStartElementID	integer,
				@iTemp				integer,
				@iBaseTable		integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer,
				@TargetName varchar(MAX);

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID,
				Q.parent1TableID,
				Q.parent1RecordID,
				Q.parent2TableID,
				Q.parent2RecordID,
				WF.baseTable
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID;

				EXEC [dbo].[sp_ASRIntGetRecordDescription] @iBaseTable, @iRecordID, 0, 0, @TargetName OUTPUT;
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, 
					initiatorID, 
					status, 
					userName, 
					parent1TableID,
					parent1RecordID,
					parent2TableID,
					parent2RecordID,
					pageno,
					TargetName)
				OUTPUT inserted.ID INTO @outputTable
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					'<Triggered>',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					0,
					@TargetName)
								
				SELECT @iInstanceID = id FROM @outputTable;
				
				UPDATE ASRSysWorkflowQueue
				SET instanceID = @iInstanceID
				WHERE queueID = @iQueueID	

				-- Create the Workflow Instance Steps records. 
				-- Set the first steps' status to be 1 (pending Workflow Engine action). 
				-- Set all subsequent steps' status to be 0 (on hold). */
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @iWorkflowID
		
				INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
		
				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
				SELECT 
					@iInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 3
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 1
						ELSE 0
					END,
					0,
					0
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
				
				-- Create the Workflow Instance Value records. 
				INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
				SELECT @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElementItems.identifier
				FROM ASRSysWorkflowElementItems 
				INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 2
					AND (ASRSysWorkflowElementItems.itemType = 3 
						OR ASRSysWorkflowElementItems.itemType = 5
						OR ASRSysWorkflowElementItems.itemType = 6
						OR ASRSysWorkflowElementItems.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 13
						OR ASRSysWorkflowElementItems.itemType = 14
						OR ASRSysWorkflowElementItems.itemType = 15
						OR ASRSysWorkflowElementItems.itemType = 17
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 5						
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END
		