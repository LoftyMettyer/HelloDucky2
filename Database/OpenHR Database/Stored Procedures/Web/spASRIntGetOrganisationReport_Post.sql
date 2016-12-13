CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReport_Post] (	
	    @piReportID				integer
	   ,@piRootID					integer
	   ,@psOrganisationReportType	varchar(50)
      ,@psPostAllocationViewName	varchar(500))		
WITH EXECUTE AS OWNER
AS		
BEGIN		
	SET NOCOUNT ON;

	/*
	Returns two result Dataset for respected organisation report.
		1) Based on Personnel module setup in system manager, Result dataset for selected filter and selected columns for respected organisation report in hierarchylevel order.
		2) Result Dataset for respected organisationReport column's parameters like prefix,suffix,fontsize etc.
	*/

	DECLARE @topLevelPostID					integer,
			@topLevelEmployeeID				integer,
			@sSQL							nvarchar(MAX) = '',
			@sColumnList					varchar(MAX)  = '',
			@sFilterList					nvarchar(max) = '',
			@sJoinList						nvarchar(MAX) = '',
			@sWhereCondition				varchar(MAX)  = '',
			@sOrderCondition				nvarchar(MAX) = '',
			@sVacantPost					nvarchar(MAX) = '0 AS IsVacantPost',
			@sEmployeeInfo					nvarchar(MAX) = '0 AS EmployeeID',
			@sTodayDate						varchar(50) = CONVERT(varchar(50),DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate()))),
			@sPersonnelTableName			varchar(MAX),
			@iPersonnelTableID				integer,			 
			@sPostTableName					varchar(MAX),
			@iPostTableID					integer,
			@sPostAllocationTableName		varchar(MAX),
            @sPostAllocationViewName		varchar(MAX),
			@sHierarchyIdentifierColumn		varchar(MAX),
			@sHierarchyReportsToColumn		varchar(MAX),
			@sPostAllocationStartDateColumn	varchar(MAX),
			@sPostAllocationEndDateColumn	varchar(MAX),
			@iBaseTableId					integer,
            @sBaseViewName					varchar(MAX),
			@iHierarchyTableID				integer,
			@iPostAllocationTableID			integer,
			@sPersonnelTableViewName		varchar(max),
			@UseAppointment					bit = 0,
			@UsePersonnel					bit = 0;
	DECLARE @allNodes OrgChartRelation,
			@ghostNodes OrgChartRelation;


	--Assigned postallocation view name 
	SET @sPostAllocationViewName = @psPostAllocationViewName;

   -- Get BaseViewName, BaseViewTableName base on organisationID
	SELECT @sBaseViewName =v.ViewName, @iBaseTableId = t.TableID
	   FROM ASRSysOrganisationReport AS r
	   INNER JOIN ASRSysViews v ON  r.BaseViewID = v.ViewID
	   INNER JOIN ASRSysTables t ON v.ViewTableID = t.tableID
	   WHERE r.ID = @piReportID;

	-- Get module setup parameters
	SELECT @iPersonnelTableID=t.tableID, @sPersonnelTableName = t.TableName FROM ASRSysModuleSetup s 	 
	   INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID WHERE s.parameterKey like 'Param_Table%' AND s.moduleKey = 'MODULE_PERSONNEL';
	
	SELECT @iPostTableID=t.tableID, @sPostTableName = t.TableName FROM ASRSysModuleSetup s 	 
	   INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID WHERE s.parameterKey like 'Param_PostTable%' AND s.moduleKey = 'MODULE_POST';
	
	SELECT @sHierarchyIdentifierColumn = c.ColumnName
	   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
	   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
	   AND UPPER(s.ParameterKey) = 'PARAM_FIELDIDENTIFIER'; 

	SELECT @sHierarchyReportsToColumn = c.ColumnName
	   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
	   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
	   AND UPPER(s.ParameterKey) = 'PARAM_FIELDREPORTSTO'; 		

    --Get postallocation table name
	SELECT @iPostAllocationTableID = t.TableID
				,@sPostAllocationTableName = t.TableName 
	   FROM ASRSysModuleSetup s INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID 
	   WHERE s.parameterKey LIKE 'Param_Table%' AND s.moduleKey = 'MODULE_HIERARCHY' 
	   AND UPPER(s.ParameterKey) = 'PARAM_TABLEPOSTALLOCATION';

	SELECT @sPostAllocationStartDateColumn = c.ColumnName
	   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
	   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY'
	   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSTARTDATE'; 

	SELECT @sPostAllocationEndDateColumn = c.ColumnName
	   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
	   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
	   AND UPPER(s.ParameterKey) = 'PARAM_FIELDENDDATE'; 

	-- Calculate the post ID of the root employee
	SELECT @topLevelEmployeeID = dbo.udfASRIntOrgChartGetTopLevelID(@piRootID); 
	SET @sSQL = N'SELECT TOP 1 @topLevelPostID = ID_' + CONVERT(varchar(10),@iPostTableID) + ' FROM ' + @sPostAllocationTableName +
						   ' WHERE ID_'+ CONVERT(varchar(10),@iPersonnelTableID) +'=' + CONVERT(varchar(10), @topLevelEmployeeID);	
	EXECUTE sp_executesql @sSQL, N'@topLevelPostID integer OUTPUT', @topLevelPostID=@topLevelPostID output;

	-- Are appointment columns referenced?
	IF EXISTS (SELECT * FROM ASRSysOrganisationColumns oc
			INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId	
			WHERE c.tableID = @iPostAllocationTableID AND oc.OrganisationID = @piReportID)
		SET @UseAppointment = 1;

	-- Are personnel columns referenced?
	IF EXISTS (SELECT * FROM ASRSysOrganisationColumns oc
			INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId	
			WHERE c.tableID = @iPersonnelTableID AND oc.OrganisationID = @piReportID)
		SET @UsePersonnel = 1;

    -- Fetch personnel table view name to build final  columns selection  and wherecondition string.
	SELECT @sPersonnelTableViewName=v.ViewName  
	    FROM  ASRSysOrganisationColumns oc   
	    INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
	    INNER JOIN ASRSysTables t ON c.tableID = t.tableID	
	    INNER JOIN ASRSysViews v ON oc.ViewID = v.ViewID	   
	    WHERE oc.OrganisationID = @piReportID AND UPPER(t.TableName) = UPPER(@sPersonnelTableName);			
		
    -- Build a filter string based on filters selected on filter tab.
  	SELECT @sFilterList = @sFilterList + CASE WHEN LEN(@sFilterList) > 0 THEN ' AND (' ELSE ' (' END +
		CASE t.TableID
			WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app'
			END + '.[' + c.columnName + ']' +
		CASE WHEN c.datatype = -7 THEN /* Logic column (must be the equals operator).	*/								
			CASE WHEN  oc.Operator = 1 THEN ' = ' + CASE WHEN  UPPER(oc.Value) = 'TRUE' THEN '1'ELSE '0' END
			END	+ ')'
		WHEN (c.datatype = 2) OR (c.datatype = 4)  THEN /* Numeric/Integer column. */
			CASE oc.Operator
				WHEN 1 THEN ' = '  + oc.Value
				WHEN 2 THEN ' <> ' + oc.Value
				WHEN 3 THEN ' <= ' + oc.Value
				WHEN 4 THEN ' >= ' + oc.Value
				WHEN 5 THEN ' > '  + oc.Value
				WHEN 6 THEN ' < '  + oc.Value
			END	+ ')'
		WHEN (c.datatype = 11) THEN /* Date column. */
			CASE
					WHEN oc.Operator = 1 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' = '''  + oc.Value + '''' ELSE ' IS NULL' END
					WHEN oc.Operator = 2 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' <> ''' + oc.Value + '''' ELSE ' IS NOT NULL' END
					WHEN oc.Operator = 3 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' <= ''' + oc.Value + ''' OR ' 
						+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
						+ ' IS NULL' ELSE ' IS NULL)' END
					WHEN oc.Operator = 4 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' >= ''' + oc.Value + '''' ELSE ' IS NULL OR [' 
						+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
						+ ' IS NOT NULL)' END
					WHEN oc.Operator = 5 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' > '''  + oc.Value + '''' ELSE ' IS NOT NULL' END
					WHEN oc.Operator = 6 THEN CASE WHEN  LEN(oc.Value) > 0 THEN ' < '''  + oc.Value + ''' OR ' 
						+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
						+ ' IS NULL' ELSE ' IS NULL AND ' 
						+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
						+ ' IS NOT NULL)' END
			END	+ ')'
		WHEN ((c.datatype <> -7) AND (c.datatype <> 2) AND (c.datatype <> 4) AND (c.datatype <> 11)) THEN /* Character/Working Pattern column. */
			CASE
				WHEN oc.Operator = 1 THEN CASE WHEN  LEN(oc.Value) = 0 THEN ' = '''' OR ' 
					+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
					+ ' IS NULL' ELSE ' LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END + ')'
				WHEN oc.Operator = 2 THEN CASE WHEN  LEN(oc.Value) = 0 THEN ' <> '''' AND ' 
					+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
					+ ' IS NOT NULL' ELSE ' NOT LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END + ')'
				WHEN oc.Operator = 7 THEN CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL OR ' 
					+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
					+ ' IS NOT NULL' ELSE ' LIKE ''%' + replace(oc.Value, '''', '''''') + '%'''  END + ')'
				WHEN oc.Operator = 8 THEN CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL AND ' 
					+ CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END + '.[' + c.columnName + ']'
					+ ' IS NOT NULL' ELSE ' NOT LIKE ''%' +  replace(oc.Value, '''', '''''') + '%'''  END + ')'
			END
			END 
		FROM  ASRSysOrganisationReport orpt 
		INNER JOIN ASRSysOrganisationReportFilters oc ON orpt.ID = oc.OrganisationID
		INNER JOIN ASRSysColumns c ON oc.FieldID = c.columnId
		INNER JOIN ASRSysTables t  ON c.tableID = t.tableID	
		WHERE oc.OrganisationID = @piReportID
		ORDER BY t.TableName;

	-- Join the associated tables
	IF @UseAppointment = 1 OR @UsePersonnel = 1
	BEGIN
		SET @sJoinList = ' LEFT JOIN ' + @sPostAllocationViewName + ' app ON app.ID_' + convert(varchar(10), @iPostTableID) + ' = base.ID' +
			' AND (app.' + @sPostAllocationEndDateColumn + ' IS NULL OR '  +
			'app.' + @sPostAllocationEndDateColumn +'>= ''' + @sTodayDate + ''') AND ' +  
			'ISNULL(app.' + @sPostAllocationStartDateColumn + ', ''' + @sTodayDate + ''') <=' + '''' + @sTodayDate + '''';
		SET @sVacantPost = 'CASE WHEN (app.ID_' + CONVERT(varchar(10), @iPersonnelTableID) + ' = 0 OR app.ID IS NULL) AND nodes.IsGhostNode = 0 THEN 1 ELSE 0 END AS IsVacantPost';
	END

	IF @UsePersonnel = 1
	BEGIN
		SET @sJoinList = @sJoinList + ' LEFT JOIN ' + @sPersonnelTableViewName + ' emp ON emp.ID = app.ID_' + CONVERT(varchar(10), @iPersonnelTableID);
		SET @sEmployeeInfo = 'ISNULL(emp.ID, 0) AS EmployeeID';
	END

	-- Selecting columns
	SELECT @sColumnList = @sColumnList + ', ' + CASE t.TableID WHEN @iBaseTableId THEN 'base' WHEN @iPersonnelTableID THEN 'emp' ELSE 'app' END
		+ '.[' + c.ColumnName + '] AS [' + c.ColumnName + '**' + convert(varchar(8), oc.ColumnID) + ']' + CHAR(13)
		FROM ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId
		INNER JOIN ASRSysTables t ON t.TableID = c.tableID
		WHERE oc.OrganisationID = @piReportID
		ORDER BY oc.ColumnID;

	-- Ordering report
	SELECT @sOrderCondition = @sOrderCondition + ', [' + c.ColumnName + '**' + convert(varchar(8), oc.ColumnID) + ']'
		FROM ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId
		WHERE oc.OrganisationID = @piReportID AND c.datatype <> -3;

	EXECUTE AS CALLER;
	SET @sSQL = 'SELECT 0, ID, '
		+  @sHierarchyIdentifierColumn + ', ' + @sHierarchyReportsToColumn + 
		+ ' FROM ' + @sBaseViewName;
	INSERT @allNodes (IsGhostNode, EmployeeID, Staff_Number, Reports_To_Staff_Number)
		EXECUTE sp_executeSQL @sSQL;
	REVERT;
	
	-- Generate all the missing nodes (manager records that exist but are not contained in the selected base view)
	SET @sSQL = 'SELECT DISTINCT 1, nodes.Reports_To_Staff_Number, pr.' + @sHierarchyReportsToColumn + ' , ISNULL(pr.ID, 0) 
				 FROM @allNodes nodes
				 LEFT JOIN ' + @sPostTableName + ' pr ON pr.' + @sHierarchyIdentifierColumn + ' = nodes.Reports_To_Staff_Number
				 WHERE nodes.Reports_To_Staff_Number NOT IN (SELECT Staff_Number FROM @allNodes)
					AND nodes.EmployeeID <> ' + convert(varchar(10), @topLevelPostID);

	DECLARE @iRowCount int = 0

	WHILE 1=1
	BEGIN

		DELETE FROM @ghostNodes;
		INSERT @ghostNodes (IsGhostNode, Staff_Number, Reports_To_Staff_Number, EmployeeID)
			EXECUTE sp_executeSQL @sSQL, N'@allNodes OrgChartRelation READONLY, @topLevelPostID int', @allNodes=@allNodes, @topLevelPostID=@topLevelPostID;

		SET @iRowCount = @@ROWCOUNT;

		INSERT @allNodes
			SELECT * FROM @ghostNodes

		IF @iRowCount < 2
			BREAK;

	END

	-- Calculate the top node
	SELECT TOP 1 @topLevelPostID = EmployeeID FROM @allNodes WHERE Reports_To_Staff_Number 
		NOT IN (Select Staff_Number FROM @allNodes)


	-- Calculate the hierarchy levels
	DECLARE @outputNodes OrgChartRelation;
	WITH posts AS (	
		SELECT [IsGhostNode], 0 HierarchyLevel
			, [Staff_Number]
			, [EmployeeID]
			, [Reports_To_Staff_Number]	
		FROM @allNodes WHERE EmployeeID = @topLevelPostID
		UNION ALL
		SELECT nodes.IsGhostNode
			, subs.HierarchyLevel + 1
			, nodes.[Staff_Number]
			, nodes.EmployeeID
			, nodes.[Reports_To_Staff_Number]
		FROM @allNodes nodes
			INNER JOIN posts subs ON subs.[Staff_Number] = nodes.Reports_To_Staff_Number )	
	INSERT @outputNodes (IsGhostNode, ManagerRoot, HierarchyLevel, EmployeeID, Staff_Number, Reports_To_Staff_Number)
		SELECT DISTINCT IsGhostNode
			,@topLevelPostID AS ManagerRoot, HierarchyLevel, EmployeeID, Staff_Number, Reports_To_Staff_Number
		FROM posts p;

	IF LEN(@sFilterList) > 0
		SET @sFilterList = 'CASE WHEN ' + REPLACE(@sFilterList, CHAR(39), CHAR(39)) + ' THEN 0 ELSE 1 END';
	ELSE
		SET @sFilterList = '0';


	-- Merge in the selected data columns
	SET @sSQL = 'SELECT ' + @sVacantPost + ', nodes.IsGhostNode, nodes.HierarchyLevel, nodes.EmployeeID AS HierarchyID, nodes.Staff_Number AS Post_ID, nodes.Reports_To_Staff_Number AS Reports_To_Post_ID, '
		+ @sEmployeeInfo
		+ @sColumnList	
		+ ', ' + @sFilterList + ' AS [IsFilteredNode] '
		+ ' FROM ' + @sBaseViewName + ' base '
		+ @sJoinList
		+ ' RIGHT JOIN @outputNodes nodes ON nodes.EmployeeID = base.id '
		+ ' ORDER BY HierarchyLevel ASC ' + @sOrderCondition;

	EXECUTE AS CALLER;
		EXEC sp_executesql @sSQL,  N'@outputNodes OrgChartRelation READONLY, @piRootID int'
		, @outputNodes=@outputNodes, @piRootID = @piRootID;
	REVERT;

	SELECT 'unused table?';
	
END
