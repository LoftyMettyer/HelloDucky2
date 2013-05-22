CREATE PROCEDURE [dbo].[sp_ASRIntGetCrossTabTablesInfo]
AS
BEGIN
	/* Return a set of information for all of the tables in the system. */
	/* The information required is :
		table id
		table name
		table type
		string listing the ids of the table's children.
		string listing the ids of the table's parents.
		
	NB. The tables are return in name order. */
	
	SET NOCOUNT ON;
	
	DECLARE	@iTableID	integer,
			@sTableName	sysname,
			@iTableType	integer,
			@sChildren	varchar(MAX),
			@sChildrenNames varchar(2000),
			@sParents	varchar(2000),
			@iChildID	integer,
			@sChildName	varchar(2000),
			@iParentID	integer;

	DECLARE @tableInfo TABLE(
		tableID			integer,
		tableName		sysname,
		tableType		integer,
		childrenString	varchar(MAX),
		childrenNames	varchar(MAX),
		parentsString	varchar(MAX));

	DECLARE tableCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableID,
		tableName,
		tableType
	FROM [dbo].[ASRSysTables]
	         WHERE (SELECT COUNT(*) FROM ASRSysColumns
	         WHERE ASRSysColumns.TableID = ASRSysTables.TableID
	         AND columnType <> 3
	         AND columnType <> 4
	         AND dataType <> -3
	         AND dataType <> -4) > 1;

	OPEN tableCursor;
	FETCH NEXT FROM tableCursor INTO @iTableID, @sTableName, @iTableType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sChildren = '';
		SET @sParents = '';
		SET @sChildrenNames = '';

		DECLARE childCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysRelations.childID, ASRSysTables.TableName
			FROM [dbo].[ASRSysRelations]
			INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID
			WHERE ASRSysRelations.parentID = @iTableID
			ORDER BY ASRSysTables.tableName;

		OPEN childCursor;
		FETCH NEXT FROM childCursor INTO @iChildID, @sChildName;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sChildren = @sChildren + convert(varchar(MAX), @iChildID) + char(9);

			SET @sChildrenNames = @sChildrenNames +	convert(varchar(MAX), @iChildID) + char(9) + convert(varchar(MAX), @sChildName) + char(9);

			FETCH NEXT FROM childCursor INTO @iChildID, @sChildName;
		END
		CLOSE childCursor;
		DEALLOCATE childCursor;

		DECLARE parentCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysRelations.parentID
			FROM [dbo].[ASRSysRelations]
			INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID
			WHERE ASRSysRelations.childID = @iTableID
			ORDER BY ASRSysTables.tableName;

		OPEN parentCursor;
		FETCH NEXT FROM parentCursor INTO @iParentID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sParents = @sParents + convert(varchar(MAX), @iParentID) + char(9);
			FETCH NEXT FROM parentCursor INTO @iParentID;
		END
		CLOSE parentCursor;
		DEALLOCATE parentCursor;

		INSERT INTO @tableInfo (tableID, tableName, tableType, childrenString, childrenNames, parentsString) 
			VALUES(@iTableID, @sTableName, @iTableType, @sChildren, @sChildrenNames, @sParents);

		FETCH NEXT FROM tableCursor INTO @iTableID, @sTableName, @iTableType;
	END
	CLOSE tableCursor;
	DEALLOCATE tableCursor;

	SELECT *
		FROM @tableInfo 
		ORDER BY tableName;

END