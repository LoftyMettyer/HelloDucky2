CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReport] (	
	 @piReportID				integer
	,@piRootID					integer
	,@psOrganisationReportType	varchar(50)
   ,@psPostAllocationViewName	varchar(500)
)		
AS		
BEGIN		
	SET NOCOUNT ON;

	/*
	Returns two result Dataset for respected organisation report.
		1) Based on Personnel module setup in system manager, Result dataset for selected filter and selected columns for respected organisation report in hierarchylevel order.
		2) Result Dataset for respected organisationReport column's parameters like prefix,suffix,fontsize etc.
	*/

	DECLARE  @RootID								      integer,
			   @iColumnID							      integer,
			   @iTableID							      integer,
			   @iOrganisationID					      integer,
			   @sColumnName						      varchar(MAX),
			   @sTableName							      varchar(MAX),
			   @sSQL								         nvarchar(MAX) = '',
			   @sUnionAllSql						      nvarchar(MAX) = '',
			   @sUnionSql							      nvarchar(MAX) = '',
			   @sFinalOrgReportSql					   nvarchar(MAX) = '',
			   @sWhereConditionSql					   nvarchar(MAX) = '',
			   @sColumnString						      varchar(MAX)  = '',
			   @sTableString						      varchar(MAX)  = '',
			   @sFilterWhereCondition				   varchar(MAX)  = '',
			   @sWhereCondition					      varchar(MAX)  = '',
			   @sPreviousTableName					   varchar(MAX)   = '',
			   @sNextTableName						   varchar(MAX)   = '',
			   @sOrgColumnTableName				      varchar(50)	  = 'ASRSysOrganisationColumns',
			   @dTodayDate							      datetime	  = DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate())),
			   @sHierarchyLevel					      varchar(50)	  = '1 AS HierarchyLevel',
			   @sPersonnelTableName				      varchar(MAX),
			   @iPersonnelTableID					   integer,
			   @sPersonnelStartDateColumn			   varchar(MAX),
			   @sPersonnelLeavingDateColumn		   varchar(MAX),
			   @sPersonnelStaffNumberColumn		   varchar(MAX),
			   @sPersonnelReportToStaffNoColumn	   varchar(MAX),
			   @sPersonnelCTEColumn				      varchar(MAX),
			   @sHierarchyTableName				      varchar(MAX),
			   @sPostAllocationTableName			   varchar(MAX),
            @sPostAllocationViewName			   varchar(MAX),
			   @sHierarchyIdentifierColumn		   varchar(MAX),
			   @sHierarchyReportsToColumn			   varchar(MAX),
			   @sHierarchyCTEColumn				      varchar(MAX),
			   @sPostAllocationStartDateColumn	   varchar(MAX),
			   @sPostAllocationEndDateColumn		   varchar(MAX),
            @sPersonnelJobTitle                 varchar(MAX),
            @sReportsToPostIDColumn				   varchar(MAX),				
            @sBaseViewName					         varchar(MAX),
				@sBaseViewTableName					   varchar(MAX),
			   @iHierarchyTableID					   integer,
			   @iPostAllocationTableID				   integer,
			   @iHierarchyIdentifierColumnID		   integer,
			   @iHierarchyReportsToColumnID		   integer,
			   @iPostAllocationStartDateColumnID	integer,
			   @iPostAllocationEndDateColumnID		integer,
            @sFromString                        nvarchar(MAX) ='',
            @sExcludeSql					         nvarchar(MAX)='',
            @sViewName 					            nvarchar(MAX)='',
				@sPersonnelTableViewName			   varchar(max);

	SET @iOrganisationID = @piReportID;
   --Assigned postallocation view name 
	SET @sPostAllocationViewName = @psPostAllocationViewName;
	-- Get RootID of top level based on loggedIn userID from user defined scalar function udfASRIntOrgChartGetTopLevelID.
	SELECT @RootID = dbo.udfASRIntOrgChartGetTopLevelID(@piRootID); 

   -- Get BaseViewName, BaseViewTableName base on organisationID
	SELECT  @sBaseViewName =v.ViewName			
	FROM ASRSysOrganisationReport AS r
	INNER JOIN ASRSysViews v ON  r.BaseViewID = v.ViewID
	INNER JOIN ASRSysTables t ON v.ViewTableID = t.tableID
	WHERE r.ID = @iOrganisationID;

	SELECT  @iPersonnelTableID=t.tableID, @sPersonnelTableName = t.TableName FROM ASRSysModuleSetup s 	 
	INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID WHERE s.parameterKey like 'Param_Table%' AND s.moduleKey = 'MODULE_PERSONNEL';

	-- If report type is commercial system then get required columns details as per system manager personnel module setup.
	IF @psOrganisationReportType = 'COMMERCIAL' 
	   BEGIN

		   SELECT  @sPersonnelLeavingDateColumn = c.ColumnName		
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID	WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSLEAVINGDATE'; 

		   SELECT  @sPersonnelStartDateColumn = c.ColumnName	
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSSTARTDATE';

		   SELECT   @sPersonnelStaffNumberColumn = c.ColumnName	
				      ,@sPersonnelCTEColumn = UPPER('ecte'+ '.' + c.ColumnName)	
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSEMPLOYEENUMBER';

		   SELECT  @sPersonnelReportToStaffNoColumn = c.ColumnName		
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID 
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSMANAGERSTAFFNO';
		
		   --Fetch staff number of root id
		   DECLARE @columnVal TABLE (columnVal nvarchar(100));
		   DECLARE @staff_numbertemp nvarchar(100) ,@staff_number NVARCHAR(100);
		   SET @staff_numbertemp = N'Select '+ @sPersonnelStaffNumberColumn + ' FROM ' + @sBaseViewName + ' WHERE ' + @sBaseViewName+'.ID = ' + CONVERT(varchar(20),@RootID);
		  
		   INSERT @columnVal  EXEC sp_executesql @staff_numbertemp;
		   SET @staff_number = (SELECT * FROM @columnVal);	
			
         SET @sSQL = 'WITH Emp_CTE AS (' + CHAR(13) + ' SELECT '+ @sHierarchyLevel + ',' + @sBaseViewName+'.ID  AS EmployeeID' + ','
					   + @sBaseViewName +'.' + @sPersonnelStaffNumberColumn   + ',' + @sBaseViewName +'.' +  @sPersonnelReportToStaffNoColumn + ' AS ''Reports_To_Staff_Number'' ';
		
	   END
	ELSE
	   BEGIN
		   -- If it is postbased system then get required columns details as per system manager personnel module setup.
		   SELECT	 @iHierarchyTableID = t.tableID
				      ,@sHierarchyTableName = t.TableName 
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID 
		   WHERE s.parameterKey LIKE 'Param_Table%' AND s.moduleKey = 'MODULE_HIERARCHY' AND UPPER(s.ParameterKey) = 'PARAM_TABLEHIERARCHY'; 
		
		   SELECT    @sHierarchyIdentifierColumn = c.ColumnName
				      ,@sHierarchyCTEColumn = 'ecte'+ '.' + c.ColumnName
				      ,@iHierarchyIdentifierColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDIDENTIFIER'; 

		   SELECT    @sHierarchyReportsToColumn = c.ColumnName
				      ,@iHierarchyReportsToColumnID = c.columnID
                  ,@sReportsToPostIDColumn = c.ColumnName
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDREPORTSTO'; 		

         --Get postallocation table name
		   SELECT    @iPostAllocationTableID = t.TableID
				      ,@sPostAllocationTableName = t.TableName 
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysTables t ON s.ParameterValue = t.tableID 
		   WHERE s.parameterKey LIKE 'Param_Table%' AND s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_TABLEPOSTALLOCATION';

		   SELECT   @sPostAllocationStartDateColumn = c.ColumnName
				   ,@iPostAllocationStartDateColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY'
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDSTARTDATE'; 

		   SELECT    @sPostAllocationEndDateColumn = c.ColumnName
				      ,@iPostAllocationEndDateColumnID = c.columnID
		   FROM ASRSysModuleSetup s INNER JOIN ASRSysColumns c ON s.ParameterValue = c.columnID
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
		   AND UPPER(s.ParameterKey) = 'PARAM_FIELDENDDATE'; 

		   --Fetch PostID value of root id 
		   DECLARE @sPost_IDColumnVal TABLE (columnVal nvarchar(100));
		   DECLARE @sPost_IDTemp nvarchar(500),@sPost_ID NVARCHAR(100);		  
		   SET @sPost_IDTemp = N'Select '+ @sBaseViewName +'.'+@sHierarchyIdentifierColumn + ' FROM ' +     
							   @sBaseViewName +' INNER JOIN ' + @sPostAllocationViewName +	 ' ON ' +  
							   @sBaseViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID) +
							   ' WHERE ' +  @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iPersonnelTableID) +'=' + CONVERT(varchar(20),@RootID);	
		
		   INSERT @sPost_IDColumnVal  EXEC sp_executesql @sPost_IDTemp;
		   SET @sPost_ID = (SELECT * FROM @sPost_IDColumnVal);		
		
		   SET @sSQL = 'WITH Emp_CTE AS (' + CHAR(13) + ' SELECT '+ @sHierarchyLevel + ',' + @sBaseViewName +'.ID AS HierarchyID' + ',' +
						   @sBaseViewName+'.'+@sHierarchyIdentifierColumn + ',' +  @sBaseViewName+'.'+@sHierarchyReportsToColumn;	

         --Fetch personnel table view name to build final  columns selection  and wherecondition string.
			SELECT @sPersonnelTableViewName=v.ViewName  
			FROM  ASRSysOrganisationColumns oc   
			INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
			INNER JOIN ASRSysTables t ON c.tableID = t.tableID	
			INNER JOIN ASRSysViews v ON oc.ViewID = v.ViewID	   
			WHERE oc.OrganisationID = @iOrganisationID AND UPPER(t.TableName) = UPPER(@sPersonnelTableName);			
			
			IF (@sPersonnelTableViewName IS NOT NULL)
			BEGIN
				SET @sSQL = @sSQL  + ',' + @sPersonnelTableViewName+'.ID AS EmployeeID';
	   END
			ELSE
			BEGIN
				SET @sSQL = @sSQL  + ', 0 AS EmployeeID';
			END
	   END

      -- Build a filter string based on filters selected on filter tab.
		DECLARE filtercolumn_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT				
					CASE WHEN c.datatype = -7 THEN /* Logic column (must be the equals operator).	*/								
                        CASE WHEN  oc.Operator = 1 THEN  '(' + v.ViewName + '.'+ c.ColumnName + ' = ' + CASE WHEN  UPPER(oc.Value) = 'TRUE' THEN '1'ELSE '0' END + ')'
								END	
						  WHEN (c.datatype = 2) OR (c.datatype = 4)  THEN /* Numeric/Integer column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + v.ViewName + '.'+ c.ColumnName + ' = '  + oc.Value + ')'	/* Equals. */
									   WHEN oc.Operator = 2	THEN '(' + v.ViewName + '.'+ c.ColumnName + ' <> ' + oc.Value	+ ')'/* Not Equal To. */
									   WHEN oc.Operator = 3	THEN '(' + v.ViewName + '.'+ c.ColumnName + ' <= ' + oc.Value + ')'/* Less than or Equal To. */
									   WHEN oc.Operator = 4 THEN '(' + v.ViewName + '.'+ c.ColumnName + ' >= ' + oc.Value + ')'/* Greater than or Equal To. */
									   WHEN oc.Operator = 5 THEN '(' + v.ViewName + '.'+ c.ColumnName + ' > '  + oc.Value + ')'/* Greater than. */
									   WHEN oc.Operator = 6 THEN '(' + v.ViewName + '.'+ c.ColumnName + ' < '  + oc.Value	+ ')'/* Less than.*/
								END
						 WHEN (c.datatype = 11) THEN /* Date column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' = '''  + oc.Value + '''' ELSE ' IS NULL' END + ')'	
									   WHEN oc.Operator = 2	THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <> ''' + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 3	THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' <= ''' + oc.Value + ''' OR ' + v.ViewName + '.'+ c.ColumnName + ' IS NULL' ELSE ' IS NULL' END + ')'	
									   WHEN oc.Operator = 4 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' >= ''' + oc.Value + '''' ELSE ' IS NULL OR ' + v.ViewName + '.'+ c.ColumnName + ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 5 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' > '''  + oc.Value + '''' ELSE ' IS NOT NULL' END + ')'	
									   WHEN oc.Operator = 6 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) > 0 THEN ' < '''  + oc.Value + ''' OR ' + v.ViewName + '.'+ c.ColumnName + ' IS NULL' ELSE ' IS NULL AND ' + v.ViewName + '.'+ c.ColumnName + ' IS NOT NULL' END + ')'		
								END
						 WHEN ((c.datatype <> -7) AND (c.datatype <> 2) AND (c.datatype <> 4) AND (c.datatype <> 11)) THEN /* Character/Working Pattern column. */
								CASE
									   WHEN oc.Operator = 1 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' = '''' OR ' + v.ViewName + '.'+ c.ColumnName + ' IS NULL' ELSE ' LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END + 	+ ')'									  
									   WHEN oc.Operator = 2	THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' <> '''' AND ' + v.ViewName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''' + replace(replace(replace(oc.Value, '''', ''''''), '*ALL', '%'),'?', '_' ) + '''' END  	+ ')'
									   WHEN oc.Operator = 7	THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL OR ' + v.ViewName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' LIKE ''%' + replace(oc.Value, '''', '''''') + '%'''  END 	+ ')'
									   WHEN oc.Operator = 8 THEN '(' + v.ViewName + '.'+ c.ColumnName + CASE WHEN  LEN(oc.Value) = 0 THEN ' IS NULL AND ' + v.ViewName + '.'+ c.ColumnName + ' IS NOT NULL' ELSE ' NOT LIKE ''%' +  replace(oc.Value, '''', '''''') + '%'''  END  + ')'					   					   
								END
									
					  END		AS WhereCondition
			  
			FROM  ASRSysOrganisationReport orpt 
			INNER JOIN ASRSysOrganisationReportFilters oc ON orpt.ID = oc.OrganisationID
			INNER JOIN ASRSysColumns c ON oc.FieldID = c.columnId
			INNER JOIN ASRSysTables t  ON c.tableID = t.tableID	
			INNER JOIN ASRSysViews v   ON  orpt.BaseViewID = v.ViewID	
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
				   , t.TableName + '#' + CONVERT(varchar(MAX),t.TableID)  As TableName
				   ,v.ViewName				   			
		   FROM  ASRSysOrganisationColumns oc
		   INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
		   INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
		   LEFT JOIN ASRSysViews v ON oc.ViewID = v.ViewID
		   WHERE oc.OrganisationID =@iOrganisationID
		   ORDER BY t.TableName;

		OPEN columnnames_cursor;
			FETCH NEXT FROM columnnames_cursor INTO @iColumnID,@sColumnName,@sTableName,@sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
			
				IF @sColumnString ='' 
					SET @sColumnString = CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END  + '.' + CONVERT(varchar(MAX), @sColumnName)+ ' AS ' + '''' + CONVERT(varchar(MAX), @sColumnName) + '**' + CONVERT(varchar(50), @iColumnID)  +'''';
				ELSE
					SET @sColumnString = @sColumnString + ', ' + CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END + '.' + CONVERT(varchar(MAX), @sColumnName) + ' AS ' + '''' + CONVERT(varchar(MAX), @sColumnName) + '**' +  CONVERT(varchar(50), @iColumnID) +'''';
			
				IF @sTableString = ''
					BEGIN
						SET @sTableString =  CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END;	
						SET @sPreviousTableName = CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END;
					END
				ELSE
					BEGIN				
					IF @sPreviousTableName <> @sNextTableName				
						SET @sTableString =  @sTableString + ', ' + CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END;
					END

				SET @sPreviousTableName =  CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END;
				FETCH NEXT FROM columnnames_cursor INTO @iColumnID, @sColumnName,@sTableName,@sViewName;
				SET @sNextTableName = CASE WHEN  @sViewName IS NULL THEN CONVERT(varchar(MAX), @sTableName) ELSE CONVERT(varchar(MAX), @sViewName) END;
			END	

			/* Whilst saving org def we don't have view id for selected columns which are selected from postallocation table.
            So we are replacing postallocation table with postallocation view. for eg: 'Appointment#3' with  ASRSysCV3#APPOINTMENTS#Administrators*/
         IF @psOrganisationReportType <> 'COMMERCIAL'
			   BEGIN
			      SET @sColumnString =  REPLACE (@sColumnString, (@sPostAllocationTableName + '#' + CONVERT(varchar(MAX),@iPostAllocationTableID)) ,@sPostAllocationViewName);
			      SET @sTableString = REPLACE (@sTableString, (@sPostAllocationTableName + '#' + CONVERT(varchar(MAX),@iPostAllocationTableID)) ,@sPostAllocationViewName);			
			   END

			IF @sColumnString <> ''
				BEGIN					
					SET @sUnionAllSQL = REPLACE(REPLACE(@sSQL, @sHierarchyLevel ,'ecte.HierarchyLevel + 1 AS HierarchyLevel'),'WITH Emp_CTE AS (','')  + ' ,' + @sColumnString + ' FROM '+ @sTableString;
					SET @sUnionSql= REPLACE(REPLACE(@sSQL, @sHierarchyLevel ,'0 AS HierarchyLevel'),'WITH Emp_CTE AS (','')  + ' ,' + @sColumnString + ' FROM '+ @sTableString ;
					SET @sSQL = @sSQL  + ' ,' + @sColumnString + ' FROM '+ @sTableString ;
				END

		
		IF @psOrganisationReportType <> 'COMMERCIAL'
			BEGIN
				--Add postallocation view name, if no postallocation table column is in selection column list.
				IF CHARINDEX(UPPER(@sPostAllocationViewName), UPPER(@sTableString)) = 0
				BEGIN				
					SET @sSQL = @sSQL  + ' ,' + @sPostAllocationViewName;
					SET @sUnionAllSQL = @sUnionAllSQL  + ' ,' + @sPostAllocationViewName;
					SET @sUnionSql = @sUnionSql  + ' ,' + @sPostAllocationViewName;
				END
			
				--Add base view name, if no base view name column is in selection column list.
				IF CHARINDEX(UPPER(@sBaseViewName), UPPER(@sTableString)) = 0
					BEGIN
					SET @sSQL = @sSQL  + ' ,' + @sBaseViewName;						
					SET @sUnionAllSQL = @sUnionAllSQL  +  ' ,(' + @sBaseViewName 
										+ ' INNER JOIN Emp_CTE ecte ON '+ UPPER(@sHierarchyCTEColumn) + ' = ' + @sBaseViewName + '.'+ UPPER(@sHierarchyReportsToColumn) + ')' ;
					SET @sUnionSql = @sUnionSql  + ' ,' + @sBaseViewName;
					END
				ELSE IF CHARINDEX(UPPER(@sBaseViewName), UPPER(@sTableString)) > 0
					BEGIN							
						SET @sFromString = SUBSTRING ( @sUnionAllSQL ,CHARINDEX(' FROM ', UPPER(@sUnionAllSQL)) , LEN(@sUnionAllSQL) );
						SET @sUnionAllSQL = REPLACE(@sUnionAllSQL, @sFromString,'');
						SET @sFromString = REPLACE(UPPER(@sFromString) , @sBaseViewName ,' (' + @sBaseViewName + 
											' INNER JOIN Emp_CTE ecte ON '+ UPPER(@sHierarchyCTEColumn) + ' = ' + @sBaseViewName + '.' + UPPER(@sHierarchyReportsToColumn) + ')' ) ;							
						SET @sUnionAllSQL = @sUnionAllSQL + @sFromString;
					END						
			END		
				
				IF @psOrganisationReportType = 'COMMERCIAL'	
				BEGIN				
			IF CHARINDEX(UPPER(@sBaseViewName), UPPER(@sTableString)) > 0
				BEGIN
					SET @sFromString = SUBSTRING ( @sUnionAllSQL ,CHARINDEX(' FROM ', UPPER(@sUnionAllSQL)) , LEN(@sUnionAllSQL) );
					SET @sUnionAllSQL = REPLACE(@sUnionAllSQL, @sFromString,'');
				SET @sFromString = REPLACE(UPPER(@sFromString) , @sBaseViewName ,' (' + @sBaseViewName + 
								  ' INNER JOIN Emp_CTE ecte ON  ' + @sPersonnelCTEColumn + ' = ' + @sBaseViewName + '.' + @sPersonnelReportToStaffNoColumn +')' );
						SET @sUnionAllSQL = @sUnionAllSQL + @sFromString;
				END				
			END
			
			IF @psOrganisationReportType <> 'COMMERCIAL'
			
				BEGIN
					   IF (@sPersonnelTableViewName IS NOT NULL)
						BEGIN						
							SET @sWhereConditionSql = '(' + @sPersonnelTableViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+CONVERT(varchar(10),@iPersonnelTableID)+') 
							AND (' + @sBaseViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID)+')' 							
						END
					   ELSE
					   BEGIN					 
							SET @sWhereConditionSql = '(' + @sBaseViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID)+')' 
						END
					
						SET @sWhereConditionSql =@sWhereConditionSql +  ' AND (' + @sPostAllocationViewName + '.' + @sPostAllocationEndDateColumn + ' IS NULL OR '  +
						@sPostAllocationViewName + '.' + @sPostAllocationEndDateColumn +'>= ''' + CONVERT(varchar(50),@dTodayDate) + ''') AND ' +  
						@sPostAllocationViewName + '.' + @sPostAllocationStartDateColumn + '<=' + '''' + CONVERT(varchar(50),@dTodayDate) + '''' 
						+ @sFilterWhereCondition;

					SET @sSQL = @sSQL + ' WHERE UPPER(' + @sBaseViewName+ '.' + @sHierarchyReportsToColumn + ') = ''' + UPPER(CONVERT(varchar(100),@sPost_ID)) + ''' AND ' + @sWhereConditionSql 					
					
					SET @sUnionAllSQL = @sUnionAllSQL + ' WHERE ' + @sWhereConditionSql
					
					
					SET @sUnionSql = @sUnionSql + '  WHERE ' +  @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iPersonnelTableID) +'=' + CONVERT(varchar(20),@RootID)
									+ ' AND (' + @sBaseViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+ CONVERT(varchar(10),@iHierarchyTableID)+')' 
					IF (@sPersonnelTableViewName IS NOT NULL) 
						SET @sUnionSql = @sUnionSql + ' AND (' + @sPersonnelTableViewName + '.ID = ' + @sPostAllocationViewName + '.ID_'+CONVERT(varchar(10),@iPersonnelTableID)+')' 

               /* Data return from @sExcludeSql is use only to check whether we have to display hierachy level 0's data or not on org report.
					   If hierachy level 0's data match filter criteria and also there is no subordinate(child) level to it then we show hierachy level 0. */
					SET @sExcludeSql = REPLACE(@sUnionSql,'0 AS HierarchyLevel','-1 AS HierarchyLevel') 
									+ ' AND (' + @sPostAllocationViewName + '.' + + @sPostAllocationEndDateColumn + ' IS NULL OR '  +
										@sPostAllocationViewName + '.' + @sPostAllocationEndDateColumn +'>= ''' + CONVERT(varchar(50),@dTodayDate) + ''') AND ' +  
										@sPostAllocationViewName + '.' + @sPostAllocationStartDateColumn + '<=' + '''' + CONVERT(varchar(50),@dTodayDate) + '''' 
										+ CHAR(13) + @sFilterWhereCondition;
					
					SET @sFinalOrgReportSql =   @sSQL + CHAR(13) + ' UNION ALL '	+ CHAR(13) + @sUnionAllSQL	+ ')' 
										+ CHAR(13) +' SELECT * INTO #OrgReportTemp FROM ( '
										+ CHAR(13) +' SELECT p.* FROM Emp_CTE p' + CHAR(13)
										+ 'UNION ' + @sUnionSql + CHAR(13)
										+ ') orgRpt'  + CHAR(13)
                              /* Query for fetching vacant post */
										+ ' SELECT * INTO #NoAppointmentForPostID  FROM ( '+ CHAR(13) + 'SELECT DISTINCT  ot.HierarchyLevel,ot.'+@sReportsToPostIDColumn+' FROM #OrgReportTemp  ot WHERE ot.'+@sReportsToPostIDColumn+' IN' + CHAR(13)+ '( ' + CHAR(13)
										+ ' SELECT ' + @sHierarchyReportsToColumn + ' FROM ' + @sBaseViewName 
										+ CHAR(13) + ' WHERE '+ @sHierarchyIdentifierColumn + ' NOT IN ( Select ' + @sPostAllocationViewName + '.' + @sHierarchyIdentifierColumn + ' FROM ' + @sPostAllocationViewName +' ))' + CHAR(13)+')a'
                              + CHAR(13) +' INSERT INTO #OrgReportTemp ( '+ REPLACE(@sHierarchyLevel,'1 AS ', '') + ', HierarchyID ,' + @sHierarchyIdentifierColumn + ',' + @sHierarchyReportsToColumn + ', EmployeeID )' + CHAR(13)
										+ CHAR(13) + ' SELECT nafp.' + REPLACE(@sHierarchyLevel,'1 AS ', '') + ', 0 AS HierarchyID,' + @sBaseViewName +'.' + @sHierarchyIdentifierColumn + ',' + @sBaseViewName +'.' + @sHierarchyReportsToColumn + ',0'
										+ CHAR(13) + ' FROM ' + @sBaseViewName  + ' INNER JOIN #NoAppointmentForPostID nafp ON nafp.'+@sReportsToPostIDColumn +'='+ @sBaseViewName +'.'+ @sReportsToPostIDColumn
										+ CHAR(13) + ' WHERE '+ @sHierarchyIdentifierColumn + ' NOT IN ( Select ' + @sPostAllocationViewName + '.' + @sHierarchyIdentifierColumn + ' FROM ' + @sPostAllocationViewName +' )'
                              /* End of fetching vacant post */
										+ CHAR(13) + ' SELECT * FROM #OrgReportTemp  ORDER BY  hierarchylevel, ' + REPLACE(@sHierarchyReportsToColumn,@sHierarchyTableName+'.','');
			
				END
			ELSE
				BEGIN
				
					SET @sWhereConditionSql = '(' + @sBaseViewName + '.' + @sPersonnelLeavingDateColumn + ' IS NULL OR '  +
						@sBaseViewName + '.' + @sPersonnelLeavingDateColumn + ' >= ''' + CONVERT(varchar(50),@dTodayDate) + ''') AND ' +  
						@sBaseViewName + '.' + @sPersonnelStartDateColumn +' <=' + '''' + CONVERT(varchar(50),@dTodayDate) + '''' 
						+ @sFilterWhereCondition;

					SET @sSQL = @sSQL + ' WHERE UPPER(' + + @sBaseViewName + '.' + @sPersonnelReportToStaffNoColumn+ ') = '+ CONVERT(varchar(100),@staff_number) + ' AND ' + @sWhereConditionSql;					

					SET @sUnionAllSQL = @sUnionAllSQL + ' WHERE ' + @sWhereConditionSql;					

					SET @sUnionSql = @sUnionSql + ' WHERE ' + @sBaseViewName + '.ID ='  + CONVERT(varchar(10), @RootID);					 

               SET @sExcludeSql = REPLACE(@sUnionSql,'0 AS HierarchyLevel','-1 AS HierarchyLevel') + ' AND ' + @sWhereConditionSql;

					SET @sFinalOrgReportSql =   @sSQL + CHAR(13) + ' UNION ALL '	+ CHAR(13) + @sUnionAllSQL	+ ')' 
										+ CHAR(13) +' SELECT p.* FROM Emp_CTE p' + CHAR(13)
										+ 'UNION ' + @sUnionSql
										+ ' ORDER BY  hierarchylevel, Reports_To_Staff_Number';
					
				END
		
		EXEC (@sFinalOrgReportSql);
      EXEC (@sExcludeSql);
		IF OBJECT_ID('tempdb..#OrgReportTemp') IS NOT NULL
		BEGIN
			DROP TABLE #NoAppointmentForPostID;
			DROP TABLE #OrgReportTemp;
		END  
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
				,ISNULL(v.ViewID,0) AS ViewID
				,ISNULL(v.ViewName,'') AS ViewName
				,c.datatype
		FROM  ASRSysOrganisationColumns oc
		INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
		LEFT JOIN ASRSysViews v ON oc.ViewID = v.ViewID
		WHERE oc.OrganisationID =@iOrganisationID
		ORDER BY t.TableName;      
    
     UPDATE ASRSysUtilAccessLog SET 
              RunBy = system_user, 
              RunDate = getdate(), 
              RunHost = host_name() 
     WHERE UtilID = @iOrganisationID AND Type = 39;

END