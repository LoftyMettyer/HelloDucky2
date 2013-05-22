CREATE PROCEDURE [dbo].[spASRUpdateStatistics]
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sTableName		nvarchar(255),
			@sSchema		nvarchar(255),
			@sVarCommand	nvarchar(MAX);

	-- Checking fragmentation
	DECLARE tables CURSOR FOR
		SELECT sc.[Name], so.[Name]
		FROM sys.sysobjects so
			INNER JOIN sys.sysindexes si ON so.id = si.id
			INNER JOIN sys.schemas sc ON so.uid  = sc.schema_id
		WHERE so.type ='U' AND si.indid < 2 AND si.rows > 0
		ORDER BY sc.name, so.[Name];

	-- Open the cursor
	OPEN tables;

	-- Loop through all the tables in the database running dbcc showcontig on each one
	FETCH NEXT FROM tables INTO @sSchema, @sTableName;

	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @sVarCommand = 'UPDATE STATISTICS [' + @sSchema + '].[' + @sTableName + '] WITH FULLSCAN';
		EXECUTE sp_executeSQL @sVarCommand;
		FETCH NEXT FROM tables INTO @sSchema, @sTableName;
	END

	-- Close and deallocate the cursor
	CLOSE tables;
	DEALLOCATE tables;

END