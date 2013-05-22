CREATE PROCEDURE [dbo].[sp_ASRIntGetReportColumns] (
	@piBaseTableID 			integer, 
	@piParentTable1ID 		integer, 
	@piParentTable2ID 		integer, 
	@piChildTableID			varchar(MAX)		--  tab delimited string of all the selected child tables
	)
AS
BEGIN
	
	SET NOCOUNT ON;
	
	/* Return a recordset of the columns for the given table IDs.*/
	DECLARE @sUserName		sysname;
	DECLARE @sTemp			varchar(MAX);
	DECLARE @sParameter		varchar(MAX);

	SELECT @sUserName = SYSTEM_USER;

	/* Clean the input string parameters. */
	IF len(@piChildTableID) > 0 SET @piChildTableID = replace(@piChildTableID, '''', '''''');

	CREATE TABLE #ChildTempTable (ChildTableID INT PRIMARY KEY);

	SET @sTemp = @piChildTableID;
	
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
			BEGIN
				SET @sParameter = convert(integer, LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1));
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp));
			END
		ELSE
			BEGIN
				SET @sParameter = convert(integer,@sTemp);		
				SET @sTemp = '';
			END

		INSERT INTO #ChildTempTable (ChildTableID) VALUES (@sParameter);
	END
	
	SELECT 	
		'C' + char(9) +
		convert(varchar(255), ASRSysColumns.tableID) + char(9) +
		convert(varchar(255), ASRSysColumns.columnID)+ char(9) +
		--ASRSysTables.tableName + '.' + 
		ASRSysColumns.columnName + char(9) +	
		convert(varchar(255), ASRSysColumns.defaultDisplayWidth) + char(9) +
		convert(varchar(255), ASRSysColumns.decimals) + char(9) +
		'N' + char(9) +
		CASE 
			WHEN ASRSysColumns.dataType IN (2,4) THEN '1' 
			ELSE '0' 
		END as [columnDefn],
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName  as [display],	
		0 as [order]
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysTables.tableID = @piBaseTableID
		AND ASRSysColumns.columnType NOT IN (3,4) 
		AND ASRSysColumns.dataType NOT IN (-3, -4)
	UNION
	SELECT 	
		'C' + char(9) +
		convert(varchar(255), ASRSysColumns.tableID)+ char(9) +
		convert(varchar(255), ASRSysColumns.columnID) + char(9) +
		--ASRSysTables.tableName + '.' + 
		ASRSysColumns.columnName + char(9) +
		convert(varchar(255), ASRSysColumns.defaultDisplayWidth) + char(9) +
		convert(varchar(255), ASRSysColumns.decimals) + char(9) +
		'N' + char(9) +
		CASE 
			WHEN ASRSysColumns.dataType IN (2,4) THEN '1' 
			ELSE '0' 
		END as [columnDefn],
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName  as [display],	
		1 as [order]
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysTables.tableID = @piParentTable1ID
		AND ASRSysColumns.columnType NOT IN (3,4) 
		AND ASRSysColumns.dataType NOT IN (-3, -4)
	UNION
	SELECT 	
		'C' + char(9) +
		convert(varchar(255), ASRSysColumns.tableID)+ char(9) +
		convert(varchar(255), ASRSysColumns.columnID) + char(9) +
		--ASRSysTables.tableName + '.' + 
		ASRSysColumns.columnName + char(9) +
		convert(varchar(255), ASRSysColumns.defaultDisplayWidth) + char(9) +
		convert(varchar(255), ASRSysColumns.decimals) + char(9) +
		'N' + char(9) +
		CASE 
			WHEN ASRSysColumns.dataType IN (2,4) THEN '1' 
			ELSE '0' 
		END as [columnDefn],
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName  as [display],	
		2 as [order]
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysTables.tableID = @piParentTable2ID
		AND ASRSysColumns.columnType NOT IN (3,4) 
		AND ASRSysColumns.dataType NOT IN (-3, -4)
	UNION
	SELECT 	
		'C' + char(9) +
		convert(varchar(255), ASRSysColumns.tableID)+ char(9) +
		convert(varchar(255), ASRSysColumns.columnID) + char(9) +
		--ASRSysTables.tableName + '.' + 
		ASRSysColumns.columnName + char(9) +
		convert(varchar(255), ASRSysColumns.defaultDisplayWidth) + char(9) +
		convert(varchar(255), ASRSysColumns.decimals) + char(9) +
		'N' + char(9) +
		CASE 
			WHEN ASRSysColumns.dataType IN (2,4) THEN '1' 
			ELSE '0' 
		END as [columnDefn],
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName  as [display],	
		3 as [order]
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysTables.tableID IN (SELECT ChildTableID FROM #ChildTempTable)
		AND ASRSysColumns.columnType NOT IN (3,4) 
		AND ASRSysColumns.dataType NOT IN (-3, -4)
	UNION
	SELECT 	
		'E' + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysExpressions.exprID) + char(9) +
		--'<' + ASRSysTables.TableName + ' Calc> ' + 
		ASRSysExpressions.name + char(9) +
		'0'+ char(9) +
		'0'+ char(9) +
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y' 
			ELSE 'N' 
		END + char(9) +
		'0' as [columnDefn],
		'<' + ASRSysTables.TableName + ' Calc> ' + ASRSysExpressions.name as [display],	
		4 as [order]
	FROM ASRSysExpressions 
		INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.TableID 
	WHERE ASRSysExpressions.tableID = @piBaseTableID
		AND ASRSysExpressions.type = 10 
		AND ASRSysExpressions.parentComponentID = 0 
		AND ((ASRSysExpressions.access <> 'HD') 
			OR (ASRSysExpressions.access = 'HD' AND ASRSysExpressions.username = @sUserName)) 
	UNION
	SELECT 	
		'E' + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysExpressions.exprID) + char(9) +
		--'<' + ASRSysTables.TableName + ' Calc> ' + 
		ASRSysExpressions.name + char(9) +
		'0'+ char(9) +
		'0'+ char(9) +
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y' 
			ELSE 'N' 
		END + char(9) +
		'0' as [columnDefn],
		'<' + ASRSysTables.TableName + ' Calc> ' + ASRSysExpressions.name as [display],	
		4 as [order]
	FROM ASRSysExpressions 
		INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.TableID 
	WHERE ASRSysExpressions.tableID = @piParentTable1ID
		AND ASRSysExpressions.type = 10 
		AND ASRSysExpressions.parentComponentID = 0 
		AND ((ASRSysExpressions.access <> 'HD') 
			OR (ASRSysExpressions.access = 'HD' AND ASRSysExpressions.username = @sUserName)) 
	UNION
	SELECT 	
		'E' + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysExpressions.exprID) + char(9) +
		--'<' + ASRSysTables.TableName + ' Calc> ' + 
		ASRSysExpressions.name + char(9) +
		'0'+ char(9) +
		'0'+ char(9) +
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y' 
			ELSE 'N' 
		END + char(9) +
		'0' as [columnDefn],
		'<' + ASRSysTables.TableName + ' Calc> ' + ASRSysExpressions.name as [display],	
		4 as [order]
	FROM ASRSysExpressions 
		INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.TableID 
	WHERE ASRSysExpressions.tableID = @piParentTable2ID
		AND ASRSysExpressions.type = 10 
		AND ASRSysExpressions.parentComponentID = 0 
		AND ((ASRSysExpressions.access <> 'HD') 
			OR (ASRSysExpressions.access = 'HD' AND ASRSysExpressions.username = @sUserName)) 
	UNION
	SELECT 	
		'E' + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysExpressions.exprID) + char(9) +
		--'<' + ASRSysTables.TableName + ' Calc> ' + 
		ASRSysExpressions.name + char(9) +
		'0'+ char(9) +
		'0'+ char(9) +
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y' 
			ELSE 'N' 
		END + char(9) +
		'0' as [columnDefn],
		'<' + ASRSysTables.TableName + ' Calc> ' + ASRSysExpressions.name as [display],	
		5 as [order]
	FROM ASRSysExpressions 
		INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.TableID 
	WHERE ASRSysExpressions.tableid IN (SELECT convert(integer, ChildTableID) FROM #ChildTempTable)
		AND ASRSysExpressions.type = 10 
		AND ASRSysExpressions.parentComponentID = 0 
		AND ((ASRSysExpressions.access <> 'HD') 
			OR (ASRSysExpressions.access = 'HD' AND ASRSysExpressions.username = @sUserName)) 
	ORDER BY [order], [display];

END