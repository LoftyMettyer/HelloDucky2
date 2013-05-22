CREATE PROCEDURE [dbo].[spASREmailRebuild]
AS
BEGIN	
	/* Refresh all calculated columns in the database. */
	DECLARE @sTableName 	varchar(255),
		@iTableID			integer,
		@sSQL				nvarchar(MAX),
		@sColumnName		varchar(255);

	
	/* Get a cursor of the tables in the database. */
	DECLARE curTables CURSOR FOR
		SELECT tableName, tableID
		FROM ASRSysTables
	OPEN curTables;

	DELETE FROM AsrSysEmailQueue WHERE DateSent Is Null AND [Immediate] = 0;

	/* Loop through the tables in the database. */
	FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
	WHILE @@fetch_status <> -1
	BEGIN
		/* Get a cursor of the records in the current table.  */
		/* Call the diary trigger for that table and record  */
		SET @sSQL = 'DECLARE @iCurrentID	int,
						@sSQL		nvarchar(MAX);
					
					IF EXISTS (SELECT * FROM sysobjects
					WHERE id = object_id(''spASREmailRebuild_' + LTrim(Str(@iTableID)) + ''') 
						AND sysstat & 0xf = 4)
					BEGIN
						DECLARE curRecords CURSOR FOR
						SELECT id
						FROM ' + @sTableName + ';
		
						OPEN curRecords;
		
						FETCH NEXT FROM curRecords INTO @iCurrentID;
						WHILE @@fetch_status <> -1
						BEGIN
							PRINT ''ID : '' + Str(@iCurrentID);
							SET @sSQL = ''EXEC spASREmailRebuild_' + LTrim(Str(@iTableID)) 
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

	EXEC [dbo].spASREmailImmediate '';

END
GO

