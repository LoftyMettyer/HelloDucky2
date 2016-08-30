CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReport] (	
	 @piReportID				integer
	,@piRootID					integer
	,@psOrganisationReportType	varchar(50)
)		
AS		
BEGIN		
	SET NOCOUNT ON;

	/*
	Returns two result Dataset for respected organisation report.
		1) Based on Personnel module setup in system manager, Result dataset for selected filter and selected columns for respected organisation report in hierarchylevel order.
		2) Result Dataset for respected organisationReport column's parameters like prefix,suffix,fontsize etc.
	*/

	DECLARE  @RootID								   integer,
			   @iColumnID							   integer,
			   @iTableID							   integer,
			   @iOrganisationID					   integer,
			   @sColumnName						   varchar(50),
			   @sTableName							   varchar(50),
			   @sSQL								      nvarchar(MAX) = '',
			   @sUnionAllSql						   nvarchar(MAX) = '',
			   @sUnionSql							   nvarchar(MAX) = '',
			   @sFinalOrgReportSql					nvarchar(MAX) = '',
			   @sWhereConditionSql					nvarchar(MAX) = '',
			   @sColumnString						   varchar(MAX)  = '',
			   @sTableString						   varchar(MAX)  = '',
			   @sFilterWhereCondition				varchar(MAX)  = '',
			   @sWhereCondition					   varchar(MAX)  = '',
			   @sPreviousTableName					varchar(50)   = '',
			   @sNextTableName						varchar(50)   = '',
			   @sOrgColumnTableName				   varchar(50)	  = 'ASRSysOrganisationColumns',
			   @dTodayDate							   datetime	  = DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate())),
			   @sHierarchyLevel					   varchar(50)	  = '1 AS HierarchyLevel',
			   @sPersonnelTableName				   varchar(50),
			   @iPersonnelTableID					integer,
			   @sPersonnelStartDateColumn			varchar(50),
			   @sPersonnelLeavingDateColumn		varchar(50),
			   @sPersonnelStaffNumberColumn		varchar(50),
			   @sPersonnelReportToStaffNoColumn	varchar(50),
			   @sPersonnelCTEColumn				   varchar(50),
			   @sHierarchyTableName				   varchar(50),
			   @sPostAllocationTableName			varchar(50),
			   @sHierarchyIdentifierColumn		varchar(50),
			   @sHierarchyReportsToColumn			varchar(50),
			   @sHierarchyCTEColumn				   varchar(50),
			   @sPostAllocationStartDateColumn	varchar(50),
			   @sPostAllocationEndDateColumn		varchar(50),
            @sPersonnelJobTitle              varchar(50),
			   @iHierarchyTableID					integer,
			   @iPostAllocationTableID				integer,
			   @iHierarchyIdentifierColumnID		integer,
			   @iHierarchyReportsToColumnID		integer,
			   @iPostAllocationStartDateColumnID	integer,
			   @iPostAllocationEndDateColumnID		integer,
            @sFromString                        nvarchar(MAX) ='';

	SET @iOrganisationID = @piReportID;

	-- Get RootID of top level based on loggedIn userID from user defined scalar function udfASRIntOrgChartGetTopLevelID.
	SELECT @RootID = dbo.udfASRIntOrgChartGetTopLevelID(@piRootID); 

	SELECT  @iPersonnelTableID=t.tableID, @sPersonnelTableName = t.TableName FROM ASRSysModuleSetup s 	 
	INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID WHERE s.parameterKey like 'Param_Table%' AND s.moduleKey = 'MODULE_PERSONNEL';

   SELECT  @sPersonnelJobTitle = t.TableName + '.' + c.ColumnName FROM ASRSysModuleSetup s 
   INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
	AND UPPER(s.ParameterKey) = 'PARAM_FIELDSJOBTITLE';

	-- If report type is commercial system then get required columns details as per system manager personnel module setup.
	IF @psOrganisationReportType = 'COMMERCIAL' 
	   BEGIN

		   SELECT  @sPersonnelLeavingDateColumn = UPPER(t.TableName + '.' + c.ColumnName)		
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID	WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSLEAVINGDATE'; 

		   SELECT  @sPersonnelStartDateColumn = UPPER(t.TableName + '.' + c.ColumnName)		
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSSTARTDATE';

		   SELECT   @sPersonnelStaffNumberColumn = UPPER(t.TableName + '.' + c.ColumnName)	
				      ,@sPersonnelCTEColumn = UPPER('ecte'+ '.' + c.ColumnName)	
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSEMPLOYEENUMBER';

		   SELECT  @sPersonnelReportToStaffNoColumn = UPPER(t.TableName + '.' + c.ColumnName)		
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSMANAGERSTAFFNO';
		
		   --Fetch staff number of root id
		   DECLARE @columnVal TABLE (columnVal nvarchar(100));
		   DECLARE @staff_numbertemp nvarchar(100) ,@staff_number NVARCHAR(100);
		   SET @staff_numbertemp = N'Select '+ @sPersonnelStaffNumberColumn + ' FROM ' + @sPersonnelTableName + ' WHERE ' + @sPersonnelTableName+'.ID = ' + CONVERT(varchar(20),@RootID);
		   INSERT @columnVal  EXEC sp_executesql @staff_numbertemp;
		   SET @staff_number = (SELECT * FROM @columnVal);	
			
		   SET @sSQL = 'WITH Emp_CTE AS (' + CHAR(13) + ' SELECT '+ @sHierarchyLevel + ',' + @sPersonnelTableName+'.ID' + ','
					   + @sPersonnelStaffNumberColumn   + ',' +  @sPersonnelReportToStaffNoColumn + ' AS ''Reports_To_Staff_Number'' ' + ',' +@sPersonnelJobTitle;
		
	   END
	ELSE
	   BEGIN
		   -- If it is postbased system then get required columns details as per system manager personnel module setup.
		   SELECT	 @iHierarchyTableID = t.tableID
				      ,@sHierarchyTableName = t.TableName 
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID 
		   WHERE s.parameterKey LIKE 'Param_Table%' AND s.moduleKey = 'MODULE_HIERARCHY' AND UPPER(s.ParameterKey) = 'PARAM_TABLEHIERARCHY'; 
		
		   SELECT    @sHierarchyIdentifierColumn = t.TableName + '.' + c.ColumnName
				      ,@sHierarchyCTEColumn = 'ecte'+ '.' + c.ColumnName
				      ,@iHierarchyIdentifierColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDIDENTIFIER'; 

		   SELECT    @sHierarchyReportsToColumn = t.TableName + '.' + c.ColumnName
				      ,@iHierarchyReportsToColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDREPORTSTO'; 		

		   SELECT    @iPostAllocationTableID = t.TableID
				      ,@sPostAllocationTableName = t.TableName 
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID 
		   WHERE s.parameterKey LIKE 'Param_Table%' AND s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_TABLEPOSTALLOCATION';

		   SELECT   @sPostAllocationStartDateColumn = t.TableName + '.' + c.ColumnName
				   ,@iPostAllocationStartDateColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY'
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSTARTDATE'; 

		   SELECT    @sPostAllocationEndDateColumn = t.TableName + '.' + c.ColumnName
				      ,@iPostAllocationEndDateColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDENDDATE'; 

		   --Fetch PostID value of root id 
		   DECLARE @sPost_IDColumnVal TABLE (columnVal nvarchar(100));
		   DECLARE @sPost_IDTemp nvarchar(100),@sPost_ID NVARCHAR(100);
		   SET @sPost_IDTemp = N'Select '+ @sPersonnelTableName + REPLACE(@sHierarchyIdentifierColumn,@sHierarchyTableName,'') + ' FROM ' + 
							   @sPersonnelTableName + ' WHERE ' + @sPersonnelTableName+'.ID = ' + CONVERT(varchar(20),@RootID);		
		   INSERT @sPost_IDColumnVal  EXEC sp_executesql @sPost_IDTemp;
		   SET @sPost_ID = (SELECT * FROM @sPost_IDColumnVal);		
		
		   SET @sSQL = 'WITH Emp_CTE AS (' + CHAR(13) + ' SELECT '+ @sHierarchyLevel + ',' + @sPersonnelTableName+'.ID' + ',' +
						   @sHierarchyIdentifierColumn + ',' +  @sHierarchyReportsToColumn + ',' +@sPersonnelJobTitle;	
	   END

      -- Build a filter string based on filters selected on filter tab.
		DECLARE filtercolumn_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT				
					CASE WHEN c.datatype = -7 THEN /* Logic column (must be the equals operator).	*/								
                        CASE WHEN  oc.Operator = 1 THEN  '(' + t.TableName + '.'+ c.ColumnName + ' = ' + CASE WHEN  UPPER(oc.Value) = 'TRUE' THEN '1'ELSE '0' END + ')'
								END	
						  WHEN (c.datatype = 2) OR (c.datatype = 4)  THEN /* Numeric/Integer column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + t.TableName + '.'+ c.ColumnName + ' = '  + oc.Value + ')'	/* Equals. */
									   WHEN oc.Operator = 2	THEN '(' + t.TableName + '.'+ c.ColumnName + ' <> ' + oc.Value	+ ')'/* Not Equal To. */
									   WHEN oc.Operator = 3	THEN '(' + t.TableName + '.'+ c.ColumnName + ' <= ' + oc.Value + ')'/* Less than or Equal To. */
									   WHEN oc.Operator = 4 THEN '(' + t.TableName + '.'+ c.ColumnName + ' >= ' + oc.Value + ')'/* Greater than or Equal To. */
									   WHEN oc.Operator = 5 THEN '(' + t.TableName + '.'+ c.ColumnName + ' > '  + oc.Value + ')'/* Greater than. */
									   WHEN oc.Operator = 6 THEN '(' + t.TableName + '.'+ c.ColumnName + ' < '  + oc.Value	+ ')'/* Less than.*/
								END
						 WHEN (c.datatype = 11) THEN /* Date column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' = '''  + oc.Value + '''' ELSE ' IS NULL' END + ')'	
									   WHEN oc.Operator = 2	THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <> ''' + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 3	THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <= ''' + oc.Value + ''' OR ' + t.TableName + '.'+ c.ColumnName + ' IS NULL' ELSE ' IS NULL' END + ')'	
									   WHEN oc.Operator = 4 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' >= ''' + oc.Value + '''' ELSE ' IS NULL OR ' + t.TableName + '.'+ c.ColumnName + ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 5 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' > '''  + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 6 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' < '''  + oc.Value + ''' OR ' + t.TableName + '.'+ c.ColumnName + ' IS NULL' ELSE ' IS NULL AND ' + t.TableName + '.'+ c.ColumnName + ' IS NOT NULL' END + ')'		
								END
						 WHEN ((c.datatype <> -7) AND (c.datatype <> 2) AND (c.datatype <> 4) AND (c.datatype <> 11)) THEN /* Character/Working Pattern column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' = '''' OR ' + t.TableName + '.'+ c.ColumnName + ' IS NULL' ELSE ' LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''),'*', '%'),'?', '_' ) END + 	+ ')'
									   WHEN oc.Operator = 2	THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' <> '''' AND ' + t.TableName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''),'*', '%'),'?', '_' ) END + 	+ ')'
									   WHEN oc.Operator = 7	THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL OR ' + t.TableName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' LIKE ''%' + replace(oc.Value, '''', '''''') + '%''' END + 	+ ')'
									   WHEN oc.Operator = 8 THEN '(' + t.TableName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL AND ' + t.TableName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''%' +  replace(oc.Value, '''', '''''') + '%''' END + 	+ ')'					   
								END
									
					  END		AS WhereCondition
			  
			FROM  ASRSysOrganisationReportFilters oc
			INNER JOIN ASRSysColumns c ON oc.FieldID = c.columnId		
			INNER JOIN ASRSysTables t ON c.tableID = t.tableID			
			WHERE oc.OrganisationID =@iOrganisationID
			ORDER BY t.TableName;

			OPEN filtercolumn_cursor;
					FETCH NEXT FROM filtercolumn_cursor INTO @sWhereCondition;
					WHILE (@@fetch_status = 0)
					BEGIN
					IF (@sWhereCondition <> '' AND @sFilterWhereCondition ='')
						SET @sFilterWhereCondition =+ ' AND ' +  CONVERT(varchar(MAX), @sWhereCondition);
					ELSE IF (@sWhereCondition <> '')
						SET @sFilterWhereCondition = @sFilterWhereCondition + ' AND ' + CONVERT(varchar(MAX), @sWhereCondition);

					FETCH NEXT FROM filtercolumn_cursor INTO @sWhereCondition;
					END
			
			CLOSE filtercolumn_cursor;		
			DEALLOCATE filtercolumn_cursor;

         -- Build a Column string based on columns selection on column tab.
	   DECLARE columnnames_cursor CURSOR LOCAL FAST_FORWARD FOR 
		   SELECT	 oc.ColumnID
				   ,c.ColumnName
				   ,t.TableName			
		   FROM  ASRSysOrganisationColumns oc
		   INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
		   LEFT JOIN ASRSysViews v ON oc.ViewID = v.ViewID
		   WHERE oc.OrganisationID =@iOrganisationID
		   ORDER BY t.TableName;

		OPEN columnnames_cursor;
			FETCH NEXT FROM columnnames_cursor INTO @iColumnID,@sColumnName,@sTableName;
			WHILE (@@fetch_status = 0)
			BEGIN
			
				IF @sColumnString ='' 
					SET @sColumnString = CONVERT(varchar(50), @sTableName) + '.' + CONVERT(varchar(50), @sColumnName)+ ' AS ' + '''' + CONVERT(varchar(50), @sColumnName) + '**' + CONVERT(varchar(50), @iColumnID)  +'''';
				ELSE
					SET @sColumnString = @sColumnString + ', ' + CONVERT(varchar(50), @sTableName) + '.' + CONVERT(varchar(50), @sColumnName) + ' AS ' + '''' + CONVERT(varchar(50), @sColumnName) + '**' +  CONVERT(varchar(50), @iColumnID) +'''';
			
				IF @sTableString = ''
					BEGIN
						SET @sTableString =  CONVERT(varchar(50), @sTableName);	
						SET @sPreviousTableName = CONVERT(varchar(50), @sTableName);
					END
				ELSE
					BEGIN				
					IF @sPreviousTableName <> @sNextTableName				
						SET @sTableString =  @sTableString + ', ' + CONVERT(varchar(50), @sTableName);
					END

				SET @sPreviousTableName =  CONVERT(varchar(50), @sTableName);
				FETCH NEXT FROM columnnames_cursor INTO @iColumnID, @sColumnName,@sTableName;
				SET @sNextTableName = CONVERT(varchar(50), @sTableName);
			END	

			IF @sColumnString <> ''
				BEGIN					
					SET @sUnionAllSQL = replace(replace(@sSQL, @sHierarchyLevel ,'ecte.HierarchyLevel + 1 AS HierarchyLevel'),'WITH Emp_CTE AS (','')  + ' ,' + @sColumnString + ' FROM '+ @sTableString;
					SET @sUnionSql= replace(replace(@sSQL, @sHierarchyLevel ,'0 AS HierarchyLevel'),'WITH Emp_CTE AS (','')  + ' ,' + @sColumnString + ' FROM '+ @sTableString ;
					SET @sSQL = @sSQL  + ' ,' + @sColumnString + ' FROM '+ @sTableString ;
				END

		
		IF @psOrganisationReportType <> 'COMMERCIAL'
			BEGIN
				--ADD postallocation table name, if no postallocation table column is in selection column list.
				IF CHARINDEX(UPPER(@sPostAllocationTableName), UPPER(@sTableString)) = 0
				BEGIN				
					SET @sSQL = @sSQL  + ' ,' + @sPostAllocationTableName;
					Set @sUnionAllSQL = @sUnionAllSQL  + ' ,' + @sPostAllocationTableName;
					SET @sUnionSql = @sUnionSql  + ' ,' + @sPostAllocationTableName;
				END
			
				--ADD hierarchy table name, if no hierarchy table column is in selection column list.
				IF CHARINDEX(UPPER(@sHierarchyTableName), UPPER(@sTableString)) = 0
					BEGIN
					SET @sSQL = @sSQL  + ' ,' + @sHierarchyTableName;						
					Set @sUnionAllSQL = @sUnionAllSQL  +  ' ,(' + @sHierarchyTableName 
								+ ' INNER JOIN Emp_CTE ecte ON '+ UPPER(@sHierarchyCTEColumn) + ' = ' + UPPER(@sHierarchyReportsToColumn) + ')' ;
					SET @sUnionSql = @sUnionSql  + ' ,' + @sHierarchyTableName;
					END
				ELSE IF CHARINDEX(UPPER(@sHierarchyTableName), UPPER(@sTableString)) > 0
					BEGIN							
						SET @sFromString = SUBSTRING ( @sUnionAllSQL ,CHARINDEX('FROM', UPPER(@sUnionAllSQL)) , LEN(@sUnionAllSQL) );
						SET @sUnionAllSQL = REPLACE(@sUnionAllSQL, @sFromString,'');
						Set @sFromString = REPLACE(UPPER(@sFromString) , @sHierarchyTableName ,' (' + @sHierarchyTableName + 
						' INNER JOIN Emp_CTE ecte ON '+ UPPER(@sHierarchyCTEColumn) + ' = ' + UPPER(@sHierarchyReportsToColumn) + ')' ) ;							
						SET @sUnionAllSQL = @sUnionAllSQL + @sFromString;
					END						
			END		
				
			--ADD personnel table name, if no  personnel table column is in selection column list.
			IF CHARINDEX(UPPER(@sPersonnelTableName), UPPER(@sTableString)) = 0	
			BEGIN									
				SET @sSQL = @sSQL  +  ' ,' + @sPersonnelTableName;
				IF @psOrganisationReportType = 'COMMERCIAL'	
				BEGIN				
				Set @sUnionAllSQL = @sUnionAllSQL  +  ' ,(' + @sPersonnelTableName 
							+ ' INNER JOIN Emp_CTE ecte ON ' + @sPersonnelCTEColumn + ' = ' + @sPersonnelReportToStaffNoColumn +')' ;
				END
				ELSE
					Set @sUnionAllSQL = @sUnionAllSQL+ ' ,' + @sPersonnelTableName;
				SET @sUnionSql = @sUnionSql  + ' ,' + @sPersonnelTableName;
			END
			ELSE IF CHARINDEX(UPPER(@sPersonnelTableName), UPPER(@sTableString)) > 0
			BEGIN				
				IF @psOrganisationReportType = 'COMMERCIAL'	
				BEGIN
					SET @sFromString = SUBSTRING ( @sUnionAllSQL ,CHARINDEX('FROM', UPPER(@sUnionAllSQL)) , LEN(@sUnionAllSQL) );
					SET @sUnionAllSQL = REPLACE(@sUnionAllSQL, @sFromString,'');
										Set @sFromString = REPLACE(UPPER(@sFromString) , @sPersonnelTableName ,' (' + @sPersonnelTableName + 
						' INNER JOIN Emp_CTE ecte ON  ' + @sPersonnelCTEColumn + ' = ' + @sPersonnelReportToStaffNoColumn +')' );
						SET @sUnionAllSQL = @sUnionAllSQL + @sFromString;
				END				
			END
			
			IF @psOrganisationReportType <> 'COMMERCIAL'
			
				BEGIN
					SET @sWhereConditionSql = '(' + @sPersonnelTableName + '.ID = ' + @sPostAllocationTableName + '.ID_'+CONVERT(varchar(10),@iPersonnelTableID)+') 
						AND (' + @sHierarchyTableName + '.ID = ' + @sPostAllocationTableName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID)+')' 
						+ ' AND (' + @sPostAllocationEndDateColumn + ' IS NULL OR '  +
						@sPostAllocationEndDateColumn +'>= ''' + CONVERT(varchar(50),@dTodayDate) + ''') AND ' +  
						@sPostAllocationStartDateColumn + '<=' + '''' + CONVERT(varchar(50),@dTodayDate) + '''' 
						+ @sFilterWhereCondition;

					SET @sSQL = @sSQL + ' WHERE UPPER(' + @sHierarchyReportsToColumn + ') = ''' + UPPER(CONVERT(varchar(100),@sPost_ID)) + ''' AND ' + @sWhereConditionSql;
					Set @sUnionAllSQL = @sUnionAllSQL +  ' WHERE ' + @sWhereConditionSql;
					
					SET @sUnionSql = @sUnionSql + ' WHERE ' + @sPersonnelTableName + '.ID ='  + CONVERT(varchar(10), @RootID) 
					+ ' AND (' + @sPersonnelTableName + '.ID = ' + @sPostAllocationTableName + '.ID_'+CONVERT(varchar(10),@iPersonnelTableID)+') 
					AND (' + @sHierarchyTableName + '.ID = ' + @sPostAllocationTableName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID)+')'  ;	
					
					SET @sFinalOrgReportSql =   @sSQL + CHAR(13) + ' UNION ALL '	+ CHAR(13) + @sUnionAllSQL	+ ')' 
										+ CHAR(13) +' SELECT p.* FROM Emp_CTE p' + CHAR(13)
										+ 'UNION ' + @sUnionSql
										+ ' ORDER BY  hierarchylevel, ' + REPLACE(@sHierarchyReportsToColumn,@sHierarchyTableName+'.','');	

			
				END
			ELSE
				BEGIN
					SET @sWhereConditionSql = '(' + @sPersonnelLeavingDateColumn + ' IS NULL OR '  +
						@sPersonnelLeavingDateColumn + ' >= ''' + CONVERT(varchar(50),@dTodayDate) + ''') AND ' +  
						@sPersonnelStartDateColumn +' <=' + '''' + CONVERT(varchar(50),@dTodayDate) + '''' 
						+ @sFilterWhereCondition;

					SET @sSQL = @sSQL + ' WHERE UPPER(' + @sPersonnelReportToStaffNoColumn+ ') = '+ CONVERT(varchar(100),@staff_number) + ' AND ' + @sWhereConditionSql;
					Set @sUnionAllSQL = @sUnionAllSQL +  ' WHERE ' + @sWhereConditionSql;
					SET @sUnionSql = @sUnionSql + ' WHERE ' + @sPersonnelTableName + '.ID ='  + CONVERT(varchar(10), @RootID); 
					SET @sFinalOrgReportSql =   @sSQL + CHAR(13) + ' UNION ALL '	+ CHAR(13) + @sUnionAllSQL	+ ')' 
										+ CHAR(13) +' SELECT p.* FROM Emp_CTE p' + CHAR(13)
										+ 'UNION ' + @sUnionSql
										+ ' ORDER BY  hierarchylevel, Reports_To_Staff_Number';
					
				END
		
		EXEC (@sFinalOrgReportSql);
			
		CLOSE columnnames_cursor;		
		DEALLOCATE columnnames_cursor;

		-- Return Result dataset for respected organisationReport column's parameters like prefix,suffix etc.
		SELECT	 oc.ColumnID
				,c.ColumnName
				,oc.Prefix
				,oc.Suffix
				,oc.FontSize
				,oc.Height
				,oc.Decimals
				,oc.ConcatenateWithNext
				,t.TableID
				,t.TableName
				,ISNULL(v.ViewID,0) as ViewID
				,ISNULL(v.ViewName,'') as ViewName
				,c.datatype
		FROM  ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
		LEFT JOIN ASRSysViews v ON oc.ViewID = v.ViewID
		WHERE oc.OrganisationID =@iOrganisationID
		ORDER BY t.TableName;

END