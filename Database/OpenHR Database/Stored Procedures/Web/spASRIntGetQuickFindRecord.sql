CREATE PROCEDURE [dbo].[spASRIntGetQuickFindRecord] (
	@plngTableID		integer,
	@plngViewID			integer,
	@plngColumnID		integer,
	@psValue			varchar(MAX),
	@psFilterDef		varchar(MAX),
	@plngRecordID		integer			OUTPUT,
	@psDecimalSeparator	varchar(100),
	@psLocaleDateFormat	varchar(100)
)
AS
BEGIN
	DECLARE 
		@lngTableID			integer,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sRealSource 		varchar(255),
		@sTableName 		varchar(255),
		@iTableType			integer,
		@iChildViewID		integer,
		@sActualUserName	sysname,
		@sColumnName		sysname,
		@iDataType			integer,
		@iResult			integer,
		@sSQL1				nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@fPermitted			bit,
		@iIndex1			integer,
		@iIndex2			integer,
		@iIndex3			integer,
		@iColumnID			integer,
		@iOperatorID		integer,
		@sValue				varchar(MAX),
		@sFilterSQL			nvarchar(MAX),
		@sSubFilterSQL		nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psValue) > 0 SET @psValue = replace(@psValue, '''', '''''')

	SET @plngRecordID = 0
	SET @sFilterSQL = ''

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Get the table ID from the view ID (if required). */
	IF @plngTableID > 0 
	BEGIN
		SET @lngTableID = @plngTableID
	END
	ELSE
	BEGIN
		SELECT @lngTableID = ASRSysViews.viewTableID
		FROM ASRSysViews
		WHERE ASRSysViews.viewID = @plngViewID
	END

	/* Get the table-type. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @lngTableID

	/* Get the column name. */
	SELECT @sColumnName = columnName,
		@iDataType = dataType
	FROM ASRSysColumns
	WHERE columnID = @plngColumnID
	
	/* Check if the current user is a System or Security manager. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1
	END
	ELSE
	BEGIN	
		SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
		FROM ASRSysGroupPermissions
		INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
		WHERE sysusers.uid = @iUserGroupID
			AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
			OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
			AND ASRSysGroupPermissions.permitted = 1
			AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'
	END

	/* Get the real source of the given screen's table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @plngViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @plngViewID
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName
		END 
	END
	ELSE
	BEGIN
		/* Get appropriate child view if required. */
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @lngTableID
			AND role = @sUserGroupName
			
		IF @iChildViewID IS null SET @iChildViewID = 0
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_')
			SET @sRealSource = left(@sRealSource, 255)
		END
	END

	/* Check the user has permission to read the column. */
	SET @fPermitted = 1
	IF @fSysSecMgr = 0
	BEGIN
		SELECT @iResult = COUNT(*)
		FROM sysprotects
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
		WHERE sysprotects.uid = @iUserGroupID
			AND sysprotects.action = 193 
			AND syscolumns.name = @sColumnName
			AND sysobjects.name = @sRealSource
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		
		IF @iResult = 0 SET @fPermitted = 0
	END
	
	IF @fPermitted = 1
	BEGIN
		/* Construct the SQL query for getting the record. */
		SET @sSQL1 = 'SELECT @result = id' +
			' FROM ' + @sRealSource +
			' WHERE ' + @sColumnName 
			
		IF (@iDataType = 11) AND len(@psValue) = 0
		BEGIN
			/* Date column - handle nulls */
			SET @sSQL1 = @sSQL1 +
				' IS null'
		END
		ELSE
		BEGIN
			SET @sSQL1 = @sSQL1 +	' = '
			IF ((@iDataType = 11) OR (@iDataType = 12)) SET @sSQL1 = @sSQL1 + ''''
			SET @sSQL1 = @sSQL1 + @psValue
			IF ((@iDataType = 11) OR (@iDataType = 12)) SET @sSQL1 = @sSQL1 + ''''
		END
	
		IF len(@psFilterDef)> 0 
		BEGIN
			WHILE charindex('	', @psFilterDef) > 0
			BEGIN
				SET @sSubFilterSQL = ''

				SET @iIndex1 = charindex('	', @psFilterDef)
				SET @iIndex2 = charindex('	', @psFilterDef, @iIndex1+1)
				SET @iIndex3 = charindex('	', @psFilterDef, @iIndex2+1)
				
				SET @iColumnID = convert(integer, LEFT(@psFilterDef, @iIndex1-1))
				SET @iOperatorID = convert(integer, SUBSTRING(@psFilterDef, @iIndex1+1, @iIndex2-@iIndex1-1))
				SET @sValue = SUBSTRING(@psFilterDef, @iIndex2+1, @iIndex3-@iIndex2-1)
				
				SET @psFilterDef = SUBSTRING(@psFilterDef, @iIndex3+1, LEN(@psFilterDef) - @iIndex3)

				SELECT @iDataType = dataType,
					@sColumnName = columnName
				FROM ASRSysColumns
				WHERE columnID = @iColumnID
								
				SET @sColumnName = @sRealSource + '.' + @sColumnName

				IF (@iDataType = -7) 
				BEGIN
					/* Logic column (must be the equals operator).	*/
					SET @sSubFilterSQL = @sColumnName + ' = '
				
					IF UPPER(@sValue) = 'TRUE'
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + '1'
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + '0'
					END
				END

				IF ((@iDataType = 2) OR (@iDataType = 4)) 
				BEGIN
					/* Numeric/Integer column. */
					/* Replace the locale decimal separator with '.' for SQL's benefit. */
					SET @sValue = REPLACE(@sValue, @psDecimalSeparator, '.')

					IF (@iOperatorID = 1) 
					BEGIN
						/* Equals. */
						SET @sSubFilterSQL = @sColumnName + ' = ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
	            SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL'
	          END
	        END

					IF (@iOperatorID = 2)
					BEGIN
						/* Not Equal To. */
						SET @sSubFilterSQL = @sColumnName + ' <> ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sSubFilterSQL + ' AND ' + @sColumnName + ' IS NOT NULL'
						END
					END

					IF (@iOperatorID = 3) 
					BEGIN
						/* Less than or Equal To. */
						SET @sSubFilterSQL = @sColumnName + ' <= ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
	            SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL'
	          END
          END
        
					IF (@iOperatorID = 4) 
					BEGIN
						/* Greater than or Equal To. */
						SET @sSubFilterSQL = @sColumnName + ' >= ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL'
						END
					END

					IF (@iOperatorID = 5) 
					BEGIN
						/* Greater than. */
						SET @sSubFilterSQL = @sColumnName + ' > ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
	            SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL'
						END
					END

					IF (@iOperatorID = 6) 
					BEGIN
						/* Less than.*/
						SET @sSubFilterSQL = @sColumnName + ' < ' + @sValue

						IF convert(float, @sValue) = 0
						BEGIN
	            SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL'
						END
					END
				END

				IF (@iDataType = 11) 
				BEGIN
					/* Date column. */
					IF LEN(@sValue) > 0
					BEGIN
						/* Convert the locale date into the SQL format. */
						/* Note that the locale date has already been validated and formatted to match the locale format. */
						SET @iIndex1 = CHARINDEX('mm', @psLocaleDateFormat)
						SET @iIndex2 = CHARINDEX('dd', @psLocaleDateFormat)
						SET @iIndex3 = CHARINDEX('yyyy', @psLocaleDateFormat)
						
						SET @sValue = SUBSTRING(@sValue, @iIndex1, 2) + '/' 
							+ SUBSTRING(@sValue, @iIndex2, 2) + '/' 
							+ SUBSTRING(@sValue, @iIndex3, 4)
					END

					IF (@iOperatorID = 1) 
					BEGIN
						/* Equal To. */
		        IF LEN(@sValue) > 0
		        BEGIN
							SET @sSubFilterSQL = @sColumnName + ' = ''' + @sValue + ''''
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL'
						END
			    END

					IF (@iOperatorID = 2)
					BEGIN
						/* Not Equal To. */
						IF LEN(@sValue) > 0
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' <> ''' + @sValue + ''''
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL'
						END
	        END

					IF (@iOperatorID = 3) 
					BEGIN
						/* Less than or Equal To. */
			      IF LEN(@sValue) > 0 
			      BEGIN
							SET @sSubFilterSQL = @sColumnName + ' <= ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL'
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL'
					  END
	        END

					IF (@iOperatorID = 4) 
					BEGIN
						/* Greater than or Equal To. */
			      IF LEN(@sValue) > 0
			      BEGIN
							SET @sSubFilterSQL = @sColumnName + ' >= ''' + @sValue + ''''
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL'
						END
					END

					IF (@iOperatorID = 5) 
					BEGIN
						/* Greater than. */
			      IF LEN(@sValue) > 0
			      BEGIN
							SET @sSubFilterSQL = @sColumnName + ' > ''' + @sValue + ''''
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL'
						END
				  END

					IF (@iOperatorID = 6)
					BEGIN
						/* Less than. */
		        IF LEN(@sValue) > 0
		        BEGIN
							SET @sSubFilterSQL = @sColumnName + ' < ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL'
						END
						ELSE
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL'
						END
	        END
				END
			
				IF ((@iDataType <> -7) AND (@iDataType <> 2) AND (@iDataType <> 4) AND (@iDataType <> 11)) 
				BEGIN
						/* Character/Working Pattern column. */
					IF (@iOperatorID = 1) 
					BEGIN
						/* Equal To. */
						IF LEN(@sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' = '''' OR ' + @sColumnName + ' IS NULL'
						END
						ELSE
						BEGIN
							/* Replace the standard * and ? characters with the SQL % and _ characters. */
							SET @sValue = replace(@sValue, '''', '''''')
							SET @sValue = replace(@sValue, '*', '%')
							SET @sValue = replace(@sValue, '?', '_')

							SET @sSubFilterSQL = @sColumnName + ' LIKE ''' + @sValue + ''''
						END
					END

					IF (@iOperatorID = 2) 
					BEGIN
						/* Not Equal To. */
						IF LEN(@sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' <> '''' AND ' + @sColumnName + ' IS NOT NULL'
						END
						ELSE
						BEGIN
							/* Replace the standard * and ? characters with the SQL % and _ characters. */
							SET @sValue = replace(@sValue, '''', '''''')
							SET @sValue = replace(@sValue, '*', '%')
							SET @sValue = replace(@sValue, '?', '_')

							SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''' + @sValue + ''''
						END
					END

					IF (@iOperatorID = 7)
					BEGIN
						/* Contains */
						IF LEN(@sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL'
						END
						ELSE
						BEGIN
							/* Replace the standard * and ? characters with the SQL % and _ characters. */
							SET @sValue = replace(@sValue, '''', '''''')

							SET @sSubFilterSQL = @sColumnName + ' LIKE ''%' + @sValue + '%'''
						END
					END

					IF (@iOperatorID = 8) 
					BEGIN
						/* Does Not Contain. */
						IF LEN(@sValue) = 0
						BEGIN
							SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL'
						END
						ELSE
						BEGIN
							/* Replace the standard * and ? characters with the SQL % and _ characters. */
							SET @sValue = replace(@sValue, '''', '''''')

							SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''%' + @sValue + '%'''
						END
					END
				END
			
				IF LEN(@sSubFilterSQL) > 0
				BEGIN
					/* Add the filter code for this grid record into the complete filter code. */
					IF LEN(@sFilterSQL) > 0
					BEGIN
						SET @sFilterSQL = @sFilterSQL + ' AND ('
					END
					ELSE
					BEGIN
						SET @sFilterSQL = @sFilterSQL + '('
					END

					SET @sFilterSQL = @sFilterSQL + @sSubFilterSQL + ')'
				END
			END

			IF LEN(@sFilterSQL) > 0 
			BEGIN
				SET @sSQL1 = @sSQL1 +
					' AND ' + @sFilterSQL
			END
		END

		SET @sParamDefinition = N'@result integer OUTPUT'
		EXEC sp_executesql @sSQL1, @sParamDefinition, @plngRecordID OUTPUT
		
		IF @plngRecordID IS null SET @plngRecordID = 0
	END	
END