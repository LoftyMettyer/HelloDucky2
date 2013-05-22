CREATE PROCEDURE [dbo].[spASRDefragIndexes]
	(@maxfrag DECIMAL)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @tablename		varchar(128),
			@sSQL			nvarchar(MAX),
			@objectid		int,
			@objectowner	varchar(255),
			@indexid		int,
			@frag			decimal,
			@indexname		char(255),
			@dbname			sysname,
			@tableid		int,
			@tableidchar	varchar(255);

	-- Checking fragmentation
	DECLARE tables CURSOR FOR
		SELECT sc.[Name]  + '.' + so.[Name]
		FROM sys.sysobjects so
			INNER JOIN sys.sysindexes si ON so.id = si.id
			INNER JOIN sys.schemas sc ON so.uid  = sc.schema_id
		WHERE so.type ='U' AND si.indid < 2 AND si.rows > 0
		ORDER BY sc.name, so.[Name];

	-- Create the temporary table to hold fragmentation information
	DECLARE @fraglist TABLE (
		ObjectName CHAR (255),
		ObjectId INT,
		IndexName CHAR (255),
		IndexId INT,
		Lvl INT,
		CountPages INT,
		CountRows INT,
		MinRecSize INT,
		MaxRecSize INT,
		AvgRecSize INT,
		ForRecCount INT,
		Extents INT,
		ExtentSwitches INT,
		AvgFreeBytes INT,
		AvgPageDensity INT,
		ScanDensity DECIMAL,
		BestCount INT,
		ActualCount INT,
		LogicalFrag DECIMAL,
		ExtentFrag DECIMAL);

	-- Open the cursor
	OPEN tables;

	-- Loop through all the tables in the database running dbcc showcontig on each one
	FETCH NEXT FROM tables INTO @tableidchar;

	WHILE @@FETCH_STATUS = 0
	BEGIN
	
		-- Do the showcontig of all indexes of the table
		INSERT INTO @fraglist 
			EXEC ('DBCC SHOWCONTIG (''' + @tableidchar + ''') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS');

		FETCH NEXT FROM tables INTO @tableidchar;
	END

	-- Close and deallocate the cursor
	CLOSE tables;
	DEALLOCATE tables;

	-- Begin Stage 2: (defrag) declare cursor for list of indexes to be defragged
	DECLARE indexes CURSOR FOR
	SELECT ObjectName, ObjectOwner = schema_name(so.uid), ObjectId, IndexName, ScanDensity
	FROM @fraglist f
	JOIN sysobjects so ON f.ObjectId=so.id
	WHERE ScanDensity <= @maxfrag
		AND INDEXPROPERTY (ObjectId, IndexName, 'IndexDepth') > 0;

	-- Open the cursor
	OPEN indexes

	-- Loop through the indexes
	FETCH NEXT FROM indexes	INTO @tablename, @objectowner, @objectid, @indexname, @frag;

	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET QUOTED_IDENTIFIER ON;

		SET @sSQL = 'ALTER INDEX [' +  RTRIM(@indexname) + '] ON ' + RTRIM(@objectowner) + '.' + RTRIM(@tablename) + ' REBUILD;';

		EXECUTE sp_executeSQL @sSQL;

		SET QUOTED_IDENTIFIER OFF;

		FETCH NEXT FROM indexes INTO @tablename, @objectowner, @objectid, @indexname, @frag;
	END

	-- Close and deallocate the cursor
	CLOSE indexes;
	DEALLOCATE indexes;

END