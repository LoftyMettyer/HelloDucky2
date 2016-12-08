CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReport_Commercial] (	
	 @piReportID				integer
	,@piRootID					integer)		
WITH EXECUTE AS OWNER
AS		
BEGIN		
	SET NOCOUNT ON;

	DECLARE @sSQL			nvarchar(MAX),
		@sColumnList		nvarchar(MAX) = '',
		@sFilterList		nvarchar(MAX) = '',
		@sBaseViewName		varchar(255),
		@sBaseViewTableName	nvarchar(255),
		@singleRecordViewName nvarchar(255),
		@sortColumns		nvarchar(MAX) = '',
		@topLevelRootID		integer,
		@topLevelReports_To	nvarchar(MAX) = '',
		@staffNoDataType	tinyint;

	DECLARE @sPersonnelStaffNumberColumn nvarchar(255),
			@sPersonnelReportToStaffNoColumn nvarchar(255);

	DECLARE @allNodes OrgChartRelation;
	DECLARE @outputNodes OrgChartRelation;


	SELECT @topLevelRootID = dbo.udfASRIntOrgChartGetTopLevelID(@piRootID); 

	-- Get module setup parameters
	SELECT @sPersonnelStaffNumberColumn = c.ColumnName, @staffNoDataType = c.datatype
		FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		AND UPPER(s.ParameterKey) = 'PARAM_FIELDSEMPLOYEENUMBER';

	SELECT @sPersonnelReportToStaffNoColumn = c.ColumnName		
		FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		AND UPPER(s.ParameterKey) = 'PARAM_FIELDSMANAGERSTAFFNO';

	SELECT @singleRecordViewName = v.ViewName
		FROM ASRSysSSIViews sv
		INNER JOIN ASRSysViews v ON v.ViewID = sv.ViewID
	WHERE SingleRecordView = 1

	-- Build the sort columns
	SELECT @sortColumns = @sortColumns + ', [' + c.ColumnName + '**' + convert(varchar(8), oc.ColumnID) + ']'
		FROM ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId
		WHERE oc.OrganisationID = @piReportID AND c.datatype <> -3;

	-- Build the column selection definition
	SELECT @sColumnList = @sColumnList + ', base.[' + c.ColumnName + '] AS [' + c.ColumnName + '**' + convert(varchar(8), oc.ColumnID) + ']' + CHAR(13)
		FROM ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId
		WHERE oc.OrganisationID = @piReportID;

	-- Build the filter definition
	SELECT @sFilterList = @sFilterList + CASE WHEN LEN(@sFilterList) > 0 THEN ' AND ' ELSE ' ' END + 
		CASE WHEN c.datatype = -7 THEN /* Logic column (must be the equals operator).	*/								
            CASE
				WHEN  oc.Operator = 1 THEN  '(base.' + c.ColumnName + ' = ' + CASE WHEN  UPPER(oc.Value) = 'TRUE' THEN '1'ELSE '0' END + ')'
			END	
		WHEN (c.datatype = 2) OR (c.datatype = 4)  THEN /* Numeric/Integer column. */
			CASE
					WHEN oc.Operator = 1 THEN '(base.' + c.ColumnName + ' = '  + oc.Value + ')'
					WHEN oc.Operator = 2 THEN '(base.' + c.ColumnName + ' <> ' + oc.Value	+ ')'
					WHEN oc.Operator = 3 THEN '(base.' + c.ColumnName + ' <= ' + oc.Value + ')'
					WHEN oc.Operator = 4 THEN '(base.' + c.ColumnName + ' >= ' + oc.Value + ')'
					WHEN oc.Operator = 5 THEN '(base.' + c.ColumnName + ' > '  + oc.Value + ')'
					WHEN oc.Operator = 6 THEN '(base.' + c.ColumnName + ' < '  + oc.Value	+ ')'
			END
		WHEN (c.datatype = 11) THEN /* Date column. */
			CASE
					WHEN oc.Operator = 1 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' = '''  + oc.Value + '''' ELSE ' IS NULL' END + ')'	
					WHEN oc.Operator = 2 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <> ''' + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
					WHEN oc.Operator = 3 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <= ''' + oc.Value + ''' OR base.' + c.ColumnName + ' IS NULL' ELSE ' IS NULL' END + ')'	
					WHEN oc.Operator = 4 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' >= ''' + oc.Value + '''' ELSE ' IS NULL OR base.' + c.ColumnName + ' IS NOT NULL' END + ')'	
					WHEN oc.Operator = 5 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' > '''  + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
					WHEN oc.Operator = 6 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' < '''  + oc.Value + ''' OR base.' + c.ColumnName + ' IS NULL' ELSE ' IS NULL AND base.' + c.ColumnName + ' IS NOT NULL' END + ')'		
			END
		WHEN ((c.datatype <> -7) AND (c.datatype <> 2) AND (c.datatype <> 4) AND (c.datatype <> 11)) THEN /* Character/Working Pattern column. */
			CASE
					WHEN oc.Operator = 1 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' = '''' OR base.' + c.ColumnName + ' IS NULL' ELSE ' LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END + 	+ ')'									  
					WHEN oc.Operator = 2 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' <> '''' AND base.' + c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END  	+ ')'
					WHEN oc.Operator = 7 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL OR base.' + c.ColumnName + ' IS NOT NULL' ELSE ' LIKE ''%' + replace(oc.Value, '''', '''''') + '%'''  END 	+ ')'
					WHEN oc.Operator = 8 THEN '(base.' + c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL AND base.' + c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''%' +  replace(oc.Value, '''', '''''') + '%'''  END  + ')'					   					   
			END
									
			END	
	FROM ASRSysOrganisationReportFilters oc 
		INNER JOIN ASRSysColumns c ON oc.FieldID = c.columnId
		WHERE oc.OrganisationID = @piReportID;


	-- Get the organisation report specifics
	SELECT @sBaseViewName = v.ViewName, @sBaseViewTableName = t.TableName		
		FROM ASRSysOrganisationReport AS r
		INNER JOIN ASRSysViews v ON  r.BaseViewID = v.ViewID
		INNER JOIN ASRSysTables t ON v.ViewTableID = t.tableID
		WHERE r.ID = @piReportID;


	-- Generate nodes for all records in the base view
	EXECUTE AS CALLER;
	SET @sSQL = 'SELECT 0, ID, '
		+  @sPersonnelStaffNumberColumn + ', ' + @sPersonnelReportToStaffNoColumn + 
		+ ' FROM ' + @sBaseViewName
	INSERT @allNodes (IsGhostNode, EmployeeID, Staff_Number, Reports_To_Staff_Number)
		EXECUTE sp_executeSQL @sSQL;
	REVERT;


	-- Generate all the missing nodes (manager records that exist but are not contained in the selected base view)
	SET @sSQL = 'SELECT DISTINCT 1, nodes.Reports_To_Staff_Number, pr.' + @sPersonnelReportToStaffNoColumn + ' , ISNULL(pr.ID, 0) 
				 FROM @allNodes nodes
				 LEFT JOIN ' + @sBaseViewTableName + ' pr ON pr.' + @sPersonnelStaffNumberColumn + ' = nodes.Reports_To_Staff_Number
				 WHERE nodes.Reports_To_Staff_Number NOT IN (SELECT Staff_Number FROM @allNodes)
					AND nodes.EmployeeID <> @piRootID';

	WHILE 1=1
	BEGIN

		INSERT @allNodes (IsGhostNode, Staff_Number, Reports_To_Staff_Number, EmployeeID)
			EXECUTE sp_executeSQL @sSQL, N'@allNodes OrgChartRelation READONLY, @piRootID int', @allNodes=@allNodes, @piRootID=@piRootID

		IF @@ROWCOUNT < 2
			BREAK;

	END


	-- Data cleasning (to remove?) (somewhere in the above isnulls are not handled properly)
   -- Processing logic is different for character and integer data types. This could be clearer to sort out in the above data generation, but for some reason it isn't that
   -- simple - fix at your own risk!
	IF @staffNoDataType = 12
	BEGIN
		UPDATE @allNodes SET Staff_Number = '00000000' WHERE Staff_Number = ''
		UPDATE @allNodes SET Reports_To_Staff_Number = '00000000' WHERE ISNULL(Reports_To_Staff_Number,'') = ''
	END
	ELSE
	BEGIN
	  	UPDATE @allNodes SET Staff_Number = '0' WHERE Staff_Number = ''
		UPDATE @allNodes SET Reports_To_Staff_Number = '0' WHERE ISNULL(Reports_To_Staff_Number,'') = ''
	END

   UPDATE @allNodes SET Reports_To_Staff_Number = '' WHERE Reports_To_Staff_Number = Staff_Number;

	-- Calculate the top most node of this dataset
	SELECT @topLevelReports_To = Staff_Number
		FROM @allNodes WHERE ISNULL(Reports_To_Staff_Number, '') NOT IN (SELECT Staff_Number FROM @allNodes);

	-- Calculate the hierarchy levels
	WITH employees AS (	
		SELECT [IsGhostNode], 0 HierarchyLevel
			, [Staff_Number]
			, [EmployeeID]
			, [Reports_To_Staff_Number]	
		FROM @allNodes WHERE Staff_Number = @topLevelReports_To
		UNION ALL
		SELECT nodes.IsGhostNode
			, subs.HierarchyLevel + 1
			, nodes.[Staff_Number]
			, nodes.EmployeeID
			, nodes.[Reports_To_Staff_Number]
		FROM @allNodes nodes
			INNER JOIN employees subs ON subs.[Staff_Number] = nodes.Reports_To_Staff_Number ) -- AND subs.HierarchyLevel > 3)
		
	INSERT @outputNodes (IsGhostNode, ManagerRoot, HierarchyLevel, EmployeeID, Staff_Number, Reports_To_Staff_Number)
		SELECT DISTINCT IsGhostNode
			,@topLevelRootID AS ManagerRoot, HierarchyLevel, EmployeeID, Staff_Number, Reports_To_Staff_Number
		FROM employees p;


	-- Format the hidden nodes string
	IF LEN(@sFilterList) > 0
		SET @sFilterList = 'CASE WHEN ' + REPLACE(@sFilterList, CHAR(39), CHAR(39)) + ' THEN 0 ELSE 1 END';
	ELSE
		SET @sFilterList = '0';


	-- Merge in the selected data columns
	SET @sSQL = 'SELECT 0 AS IsVacantPost, nodes.* ,' 
	+ @sFilterList + ' AS [IsFilteredNode]'
	+ @sColumnList
	+ ' FROM ' +  @sBaseViewName + ' base
			RIGHT JOIN @allNodes nodes ON nodes.EmployeeID = base.id 
			WHERE EmployeeID <> @piRootID
		UNION
		SELECT nodes.* ,' 
		+ @sFilterList + ' AS [IsFilteredNode]'
		+ @sColumnList	
		+ ' FROM ' + @singleRecordViewName + ' base 
			RIGHT JOIN @allNodes nodes ON nodes.EmployeeID = base.id 
			WHERE ID = @piRootID
		ORDER BY HierarchyLevel ASC' + @sortColumns;


	EXECUTE AS CALLER;
		EXEC sp_executesql @sSQL,  N'@allNodes OrgChartRelation READONLY, @piRootID int'
		, @allNodes=@outputNodes, @piRootID = @piRootID;
	REVERT;

	-- Return the top most node
	SELECT TOP 1 
		HierarchyLevel -1 AS HierarchyLevel, 
		node.EmployeeID AS EmployeeID, Staff_Number,
		'' AS Reports_To_Staff_Number
	FROM @outputNodes node
		ORDER BY HierarchyLevel ASC;
	
END
