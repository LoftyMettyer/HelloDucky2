CREATE PROCEDURE [dbo].[spASRWorkflowRebuild]
AS
BEGIN	
	-- Refresh all scheduled Workflow items in the queue.
	DECLARE @sTableName 	varchar(255),
		@iTableID			int,
		@sSQL				nvarchar(MAX)
	
	-- Get a cursor of the tables in the database.
	DECLARE curTables CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableName, tableID
		FROM ASRSysTables;
	OPEN curTables;

	DELETE FROM ASRSysWorkflowQueue 
	WHERE dateInitiated IS null 
		AND [Immediate] = 0;

	-- Loop through the tables in the database.
	FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
	WHILE @@fetch_status <> -1
	BEGIN
		/* Get a cursor of the records in the current table.  */
		/* Call the Workflow trigger for that table and record  */
		SET @sSQL = '
			DECLARE @iCurrentID	int,
				@sSQL		nvarchar(MAX);
			
			IF EXISTS (SELECT * FROM sysobjects
			WHERE id = object_id(''spASRWorkflowRebuild_' + LTrim(Str(@iTableID)) + ''') 
				AND sysstat & 0xf = 4)
			BEGIN
				DECLARE curRecords CURSOR FOR
				SELECT id
				FROM ' + @sTableName + ';

				OPEN curRecords;

				FETCH NEXT FROM curRecords INTO @iCurrentID;
				WHILE @@fetch_status <> -1
				BEGIN
					SET @sSQL = ''EXEC spASRWorkflowRebuild_' + LTrim(Str(@iTableID)) 
						+ ' '' + convert(varchar(100), @iCurrentID) + '''';
					EXECUTE sp_executeSQL @sSQL;

					FETCH NEXT FROM curRecords INTO @iCurrentID;
				END
				CLOSE curRecords;
				DEALLOCATE curRecords;
			END';

		 EXECUTE sp_executeSQL @sSQL;

		/* Move onto the next table in the database. */ 
		FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
	END

	CLOSE curTables;
	DEALLOCATE curTables;
END