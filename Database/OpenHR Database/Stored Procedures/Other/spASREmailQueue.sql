CREATE PROCEDURE [dbo].[spASREmailQueue] AS
BEGIN
	DECLARE @sSQL varchar(MAX),
		@iQueueID int,
		@iRecordID int,
		@iRecordDescID int,
		@sRecordDesc varchar(MAX)

	DECLARE emailQueue_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysEmailQueue.queueID, 
		ASRSysEmailQueue.recordID, 
		ASRSysTables.recordDescExprID
	FROM ASRSysEmailQueue
	INNER JOIN ASRSysEmailLinks ON ASRSysEmailQueue.LinkID = ASRSysEmailLinks.LinkID
	INNER JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysEmailLinks.ColumnID
	INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID
	WHERE ASRSysEmailQueue.recalculateRecordDesc = 1;
	
	OPEN emailQueue_cursor;
	FETCH NEXT FROM emailQueue_cursor INTO @iQueueID, @iRecordID, @iRecordDescID;

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sRecordDesc = '';
		
		SELECT @sSQL = 'sp_ASRExpr_' + convert(varchar,@iRecordDescID);
		IF EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @sRecordDesc OUTPUT, @iRecordID;
		END

		UPDATE ASRSysEmailQueue SET RecordDesc = @sRecordDesc WHERE queueid = @iQueueID;
		FETCH NEXT FROM emailQueue_cursor INTO @iQueueID, @iRecordID, @iRecordDescID;
	END
	CLOSE emailQueue_cursor;
	DEALLOCATE emailQueue_cursor;
END