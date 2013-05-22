CREATE PROCEDURE [dbo].[sp_ASRIntGetTablesInfo]
AS
BEGIN
	/* Return a set of information for all of the tables in the system. */
	/* The information required is :
		table id
		table name
		table type
		string listing the ids of the table's children.
		string listing the ids of the table's parents.
		string listing the ids of all the tables's relations.
	
	NB. The tables are return in name order. */
	
	SET NOCOUNT ON;
	
	DECLARE	@iTableID		integer,
			@sTableName		sysname,
			@iTableType		integer,
			@sChildren		varchar(MAX),
			@sChildrenNames varchar(MAX),
			@sParents		varchar(MAX),
			@iChildID		integer,
			@sChildName		varchar(255),
			@iParentID		integer,			
			@sRelations		varchar(MAX),
			@sRelationName	varchar(MAX),
			@iRelationID	integer;

	DECLARE @tableInfo TABLE (
		tableID		integer,
		tableName	sysname,
		tableType	integer,
		childrenString	varchar(MAX),
		childrenNames	varchar(MAX),
		parentsString	varchar(MAX),
		relatedString   varchar(MAX));

	DECLARE tableCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableID,
		tableName,
		tableType
	FROM [dbo].[ASRSysTables];

	OPEN tableCursor;
	FETCH NEXT FROM tableCursor INTO @iTableID, @sTableName, @iTableType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sChildren = '';
		SET @sParents = '';
		SET @sChildrenNames = '';
		SET @sRelations = '' ;
		
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
			SET @sChildren = @sChildren + 
				convert(varchar(MAX), @iChildID) + char(9);

			SET @sChildrenNames = @sChildrenNames +
				convert(varchar(MAX), @iChildID) + char(9) + convert(varchar(2000), @sChildName) + char(9);

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
			SET @sParents = @sParents + 
				convert(varchar(2000), @iParentID) + char(9);
			FETCH NEXT FROM parentCursor INTO @iParentID;
		END
		CLOSE parentCursor;
		DEALLOCATE parentCursor;

		DECLARE relatedCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysRelations.childID, ASRSysTables.TableName
			FROM [dbo].[ASRSysRelations]
			INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID
			WHERE ASRSysRelations.parentID = @iTableID
			UNION
			SELECT ASRSysRelations.parentID, ASRSysTables.TableName 
			FROM [dbo].[ASRSysRelations]
			INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID
			WHERE ASRSysRelations.childID = @iTableID
			UNION
			SELECT ASRSysTables.TableID, ASRSysTables.TableName 
			FROM [dbo].[ASRSysTables]
			WHERE ASRSysTables.TableID = @iTableID
			ORDER BY ASRSysTables.tableName;

		OPEN relatedCursor;
		FETCH NEXT FROM relatedCursor INTO @iRelationID, @sRelationName;
		WHILE (@@fetch_status = 0)
		BEGIN
			
			SET @sRelations = @sRelations +
				convert(varchar(MAX), @iRelationID) + char(9) + convert(varchar(2000), @sRelationName) + char(9);

			FETCH NEXT FROM relatedCursor INTO @iRelationID, @sRelationName;
		END
		CLOSE relatedCursor;
		DEALLOCATE relatedCursor;

		INSERT INTO @tableInfo (tableID, tableName, tableType, childrenString, childrenNames, parentsString, relatedString) 
			VALUES (@iTableID, @sTableName, @iTableType, @sChildren, @sChildrenNames, @sParents, @sRelations);

		FETCH NEXT FROM tableCursor INTO @iTableID, @sTableName, @iTableType;
	END
	CLOSE tableCursor;
	DEALLOCATE tableCursor;

	SELECT *
		FROM @tableInfo 
		ORDER BY tableName;

END