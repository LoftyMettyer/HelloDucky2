CREATE PROCEDURE [dbo].spASRWorkflowAscendantRecordID
		(
			@piBaseTableID	integer,
			@piBaseRecordID	integer,
			@piParent1TableID	integer,
			@piParent1RecordID	integer,
			@piParent2TableID	integer,
			@piParent2RecordID	integer,
			@piRequiredTableID	integer,
			@piRequiredRecordID	integer	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iParentTableID		integer,
				@iParentRecordID	integer
		
			SET @piRequiredRecordID = 0
			SET @piParent1TableID = isnull(@piParent1TableID, 0)
			SET @piParent1RecordID = isnull(@piParent1RecordID, 0)
			SET @piParent2TableID = isnull(@piParent2TableID, 0)
			SET @piParent2RecordID = isnull(@piParent2RecordID, 0)
		
			IF @piBaseTableID = @piRequiredTableID
			BEGIN
				SET @piRequiredRecordID = @piBaseRecordID
				RETURN
			END
		
			-- The base table is not the same as the required table.
			-- Check ascendant tables.
			DECLARE ascendantsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysRelations.parentID
			FROM ASRSysRelations
			WHERE ASRSysRelations.childID = @piBaseTableID
		
			OPEN ascendantsCursor
			FETCH NEXT FROM ascendantsCursor INTO @iParentTableID
			WHILE (@@fetch_status = 0) AND (@piRequiredRecordID = 0)
			BEGIN
				-- Get the related record in the parent table (if one exists)
				IF EXISTS 
					(SELECT * 
					FROM dbo.sysobjects 
					WHERE id = object_id(N'[dbo].[spASRSysWorkflowParentRecord]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
				BEGIN
					EXEC [dbo].[spASRSysWorkflowParentRecord]
						@piBaseTableID,
						@piBaseRecordID,
						@iParentTableID,
						@iParentRecordID OUTPUT
				END
				ELSE
				BEGIN
					SET @iParentRecordID = 0
				END
		
				IF @iParentRecordID > 0 
				BEGIN
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iParentTableID,
						@iParentRecordID,
						0,					
						0,					
						0,					
						0,					
						@piRequiredTableID,
						@piRequiredRecordID OUTPUT
				END
			
				FETCH NEXT FROM ascendantsCursor INTO @iParentTableID
			END
			CLOSE ascendantsCursor
			DEALLOCATE ascendantsCursor
			
			IF (@piRequiredRecordID = 0) 
				AND (@piParent1TableID > 0)
				AND (@piParent1RecordID > 0)
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@piParent1TableID,
					@piParent1RecordID,
					0,					
					0,					
					0,					
					0,					
					@piRequiredTableID,
					@piRequiredRecordID OUTPUT
			END

			IF (@piRequiredRecordID = 0) 
				AND (@piParent2TableID > 0)
				AND (@piParent2RecordID > 0)
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@piParent2TableID,
					@piParent2RecordID,
					0,					
					0,					
					0,					
					0,					
					@piRequiredTableID,
					@piRequiredRecordID OUTPUT
			END
		END

