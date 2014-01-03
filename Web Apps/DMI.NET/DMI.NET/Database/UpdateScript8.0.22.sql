
/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetTrainingBookingParameters]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetTrainingBookingParameters]
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetSummaryFields]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetSummaryFields]
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetScreenDefinition]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetScreenDefinition]
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetScreenControlsString2]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetScreenControlsString2]
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetFindRecords3]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetFindRecords3]
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetFilterColumns]    Script Date: 02/01/2014 20:50:36 ******/
DROP PROCEDURE [dbo].[sp_ASRIntGetFilterColumns]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetSummaryValues]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetSummaryValues]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetSelfServiceRecordID]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetSelfServiceRecordID]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetParentValues]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetParentValues]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetNavigationLinks]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetNavigationLinks]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetLinks]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetLinks]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetColumnsFromTablesAndViews]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetColumnsFromTablesAndViews]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetChartData]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGetChartData]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGet1000SeparatorFindColumns]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntGet1000SeparatorFindColumns]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntAllTablePermissions]    Script Date: 02/01/2014 20:52:28 ******/
DROP PROCEDURE [dbo].[spASRIntAllTablePermissions]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetColumnPermissions]') AND xtype = 'P')
	DROP PROCEDURE [dbo].[spASRIntGetColumnPermissions]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntCreateCacheTables]') AND xtype = 'P')
	DROP PROCEDURE [dbo].[spASRIntCreateCacheTables]
GO


IF TYPE_ID(N'DataPermissions') IS NOT NULL
	DROP TYPE [dbo].[DataPermissions]

GO

CREATE TYPE [dbo].[DataPermissions] AS TABLE
(
	name	varchar(255)
)

GO



/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetFilterColumns]    Script Date: 02/01/2014 20:50:36 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetFilterColumns]
(
	@plngTableID 	integer, 
	@plngViewID 	integer,
	@psRealSource	varchar(8000) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the IDs and names of the columns available for the given table/view. */
	DECLARE @lngTableID		integer,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sRealSource 		varchar(255),
		@sTableName 		varchar(255),
		@iTableType			integer,
		@iChildViewID		integer,
		@sActualUserName	sysname;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Get the table ID from the view ID (if required). */
	IF @plngTableID > 0 
	BEGIN
		SET @lngTableID = @plngTableID;
	END
	ELSE
	BEGIN
		SELECT @lngTableID = ASRSysViews.viewTableID
		FROM [dbo].[ASRSysViews]
		WHERE ASRSysViews.viewID = @plngViewID;
	END

	/* Get the table-type. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM [dbo].[ASRSysTables]
	WHERE ASRSysTables.tableID = @lngTableID;

	/* Check if the current user is a System or Security manager. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1;
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
			AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS';
	END

	/* Create a temporary table to hold our resultset. */
	DECLARE @ColumnInfo TABLE(
		columnID	integer,
		columnName	sysname,
		dataType	integer,
		readGranted	bit,
		size		integer,
		decimals	integer);

	/* Get the real source of the given screen's table/view. */
	IF (@fSysSecMgr = 1 )
	BEGIN
		/* Populate the temporary table with information on the order for the given table. */
		IF @plngViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM [dbo].[ASRSysViews]
			WHERE viewID = @plngViewID;

	   		INSERT INTO @ColumnInfo (
				columnID, 
				columnName,
				dataType,
				readGranted,
				size,
				decimals)
			(SELECT ASRSysColumns.columnId, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				1,
				ASRSysColumns.size,
				ASRSysColumns.decimals				
			FROM ASRSysColumns
			INNER JOIN ASRSysViewColumns ON ASRSysColumns.columnId = ASRSysViewColumns.columnID
			WHERE ASRSysColumns.tableID = @lngTableID
				AND ASRSysColumns.columnType <> 4
				AND ASRSysColumns.columnType <> 3
				AND ASRSysColumns.dataType <> -3
				AND ASRSysColumns.dataType <> -4
				AND ASRSysViewColumns.viewID = @plngViewID
				AND ASRSysViewColumns.inView=1);
		END
		ELSE
		BEGIN
			IF @iTableType <> 2 /* ie. top-level or lookup */
			BEGIN
				/* RealSource is the table. */	
				SET @sRealSource = @sTableName;
			END 
			ELSE
			BEGIN
				SELECT @iChildViewID = childViewID
				FROM [dbo].[ASRSysChildViews2]
				WHERE tableID = @lngTableID
					AND role = @sUserGroupName;
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_');
					SET @sRealSource = left(@sRealSource, 255);
				END
			END

	   		INSERT INTO @ColumnInfo (
				columnID, 
				columnName,
				dataType,
				readGranted,
				size,
				decimals)
			(SELECT ASRSysColumns.columnId, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				1,
				ASRSysColumns.size,
				ASRSysColumns.decimals				
			FROM ASRSysColumns
			WHERE ASRSysColumns.tableID = @lngTableID
				AND columnType <> 4
				AND columnType <> 3
				AND dataType <> -3
				AND dataType <> -4);
		END
	END
	ELSE
	BEGIN
		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @plngViewID > 0 
			BEGIN	
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM [dbo].[ASRSysViews]
				WHERE viewID = @plngViewID;
			END
			ELSE
			BEGIN
				SET @sRealSource = @sTableName;
			END 
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM [dbo].[ASRSysChildViews2]
			WHERE tableID = @lngTableID
				AND [role] = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sRealSource = left(@sRealSource, 255);
			END
		END

   		 INSERT INTO @ColumnInfo (
			columnID, 
			columnName,
			dataType,
			readGranted,
			size,
			decimals)
		(SELECT 
			ASRSysColumns.columnId,
			syscolumns.name,
			ASRSysColumns.dataType,
			CASE protectType
				WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END,
			ASRSysColumns.size,
			ASRSysColumns.decimals				
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		INNER JOIN syscolumns ON p.id = syscolumns.id
		INNER JOIN ASRSysColumns ON syscolumns.name = ASRSysColumns.columnName
		WHERE p.action = 193 
			AND p.uid = @iUserGroupID
			AND ASRSysColumns.tableID = @lngTableID
			AND ASRSysColumns.columnType <> 4
			AND ASRSysColumns.columnType <> 3
			AND ASRSysColumns.dataType <> -3
			AND ASRSysColumns.dataType <> -4
			AND syscolumns.name <> 'timestamp'
			AND sysobjects.name = @sRealSource
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)));
	END

	SET @psRealSource = @sRealSource;

	/* Return the resultset. */
	SELECT columnID, columnName, dataType, size, decimals
	FROM @ColumnInfo 
	WHERE readGranted = 1
	ORDER BY columnName;
END
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetFindRecords3]    Script Date: 02/01/2014 20:50:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetFindRecords3] (
	@pfError 				bit 			OUTPUT, 
	@pfSomeSelectable 		bit 			OUTPUT, 
	@pfSomeNotSelectable 	bit 			OUTPUT, 
	@psRealSource			varchar(255)	OUTPUT,
	@pfInsertGranted		bit				OUTPUT,
	@pfDeleteGranted		bit				OUTPUT,
	@piTableID 				integer, 
	@piViewID 				integer, 
	@piOrderID 				integer, 
	@piParentTableID		integer,
	@piParentRecordID		integer,
	@psFilterDef			varchar(MAX),
	@piRecordsRequired		integer,
	@pfFirstPage			bit				OUTPUT,
	@pfLastPage				bit				OUTPUT,
	@psLocateValue			varchar(MAX),
	@piColumnType			integer			OUTPUT,
	@piColumnSize			integer			OUTPUT,
	@piColumnDecimals		integer			OUTPUT,
	@psAction				varchar(255),
	@piTotalRecCount		integer			OUTPUT,
	@piFirstRecPos			integer			OUTPUT,
	@piCurrentRecCount		integer,
	@psDecimalSeparator		varchar(255),
	@psLocaleDateFormat		varchar(255)
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the find records for the current user, given the table/view and order IDs.
		@pfError = 1 if errors occured in getting the find records. Else 0.
		@pfSomeSelectable = 1 if some find columns were selectable. Else 0.
		@pfSomeNotSelectable = 1 if some find columns were NOT selectable. Else 0.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@piParentTableID = the ID of the parent table.
		@piParentRecordID = the ID of the associated record in the parent table.
	*/
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 			varchar(10),
		@fSelectGranted 	bit,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString			varchar(MAX),
		@sViewName 			varchar(255),
		@sExecString		nvarchar(MAX),
		@sTempString		varchar(MAX),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iDataType 			integer,
		@iTempAction		integer,
		@iTemp				integer,
		@sRemainingSQL		varchar(MAX),
		@iLastCharIndex		integer,
		@iCharIndex 		integer,
		@sDESCstring		varchar(5),
		@sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@fFirstColumnAsc	bit,
		@sFirstColCode		varchar(MAX),
		@sLocateCode		varchar(MAX),
		@iCount				integer,
		@iGetCount			integer,
		@iSize				integer,
		@iDecimals			integer,
		@bUse1000Separator	bit,
		@sActualLoginName	varchar(250),
		@iIndex1			integer,
		@iIndex2			integer,
		@iIndex3			integer,
		@iColumnID			integer,
		@iOperatorID		integer,
		@sValue				varchar(MAX),
		@sFilterSQL			nvarchar(MAX),
		@sTempLocateValue	varchar(MAX),
		@sSubFilterSQL		nvarchar(MAX),
		@bBlankIfZero		bit,
		@sTempFilterString	varchar(MAX),
		@sJoinSQL			nvarchar(max),
		@psOriginalAction		varchar(255),
		@sThousandColumns		varchar(255),
		@sBlankIfZeroColumns	varchar(255);
	
	/* Clean the input string parameters. */
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''');
	/* Initialise variables. */
	SET @pfError = 0;
	SET @pfSomeSelectable = 0;
	SET @pfSomeNotSelectable = 0;
	SET @sDESCstring = ' DESC';
	SET @fFirstColumnAsc = 1;
	SET @sFirstColCode = '';
	SET @piColumnSize = 0;
	SET @piColumnDecimals = 0;
	SET @bUse1000Separator = 0;
	SET @bBlankIfZero = 0;
	SET @sThousandColumns = '';
	SET @sBlankIfZeroColumns = '';

	SET @sRealSource = '';
	SET @sSelectSQL = '';
	SET @sOrderSQL = '';
	SET @sReverseOrderSQL = '';
	SET @fSelectDenied = 0;
	SET @sExecString = '';
	SET @sFilterSQL = '';
	SET @sTempLocateValue = '';
	
	SET @psOriginalAction = @psAction;
	
	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000;
	SET @psAction = UPPER(@psAction);
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE') AND 
		(@psAction <> 'LOCATEID')
	BEGIN
		SET @psAction = 'MOVEFIRST';
	END
	
	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	-- Create a temporary view of the sysprotects table
	DECLARE @SysProtects TABLE([ID] int, Columns varbinary(8000)
								, [Action] tinyint
								, ProtectType tinyint);
	INSERT INTO @SysProtects
	SELECT p.[ID], p.[Columns], p.[Action], p.ProtectType FROM ASRSysProtectsCache p WHERE p.UID = @iUserGroupID;
	
	-- Create a temporary table to hold the tables/views that need to be joined.
	DECLARE @JoinParents TABLE(tableViewName sysname, tableID integer);

	-- Create a temporary table of the 'select' column permissions for all tables/views used in the order.
	DECLARE @ColumnPermissions TABLE(tableID integer,
								tableViewName	sysname,
								columnName	sysname,
								selectGranted	bit);
								
	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM [dbo].[ASRSysTables]
		WHERE ASRSysTables.tableID = @piTableID;
	
	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1;
		RETURN;
	END
	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM [dbo].[ASRSysViews]
			WHERE viewID = @piViewID;
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM [dbo].[ASRSysChildViews2]
		WHERE tableID = @piTableID
			AND role = @sUserGroupName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END
	SET @psRealSource = @sRealSource;
	
	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	IF len(@psFilterDef)> 0 
	BEGIN
		WHILE charindex('	', @psFilterDef) > 0
		BEGIN
			SET @sSubFilterSQL = '';
			SET @iIndex1 = charindex('	', @psFilterDef);
			SET @iIndex2 = charindex('	', @psFilterDef, @iIndex1+1);
			SET @iIndex3 = charindex('	', @psFilterDef, @iIndex2+1);
				
			SET @iColumnID = convert(integer, LEFT(@psFilterDef, @iIndex1-1));
			SET @iOperatorID = convert(integer, SUBSTRING(@psFilterDef, @iIndex1+1, @iIndex2-@iIndex1-1));
			SET @sValue = SUBSTRING(@psFilterDef, @iIndex2+1, @iIndex3-@iIndex2-1);
			
			SET @psFilterDef = SUBSTRING(@psFilterDef, @iIndex3+1, LEN(@psFilterDef) - @iIndex3);
			SELECT @iDataType = dataType,
				@sColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @iColumnID;
							
			SET @sColumnName = @sRealSource + '.' + @sColumnName;
			IF (@iDataType = -7) 
			BEGIN
				/* Logic column (must be the equals operator).	*/
				SET @sSubFilterSQL = @sColumnName + ' = ';
			
				IF UPPER(@sValue) = 'TRUE'
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '1';
				END
				ELSE
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '0';
				END
			END
			
			IF ((@iDataType = 2) OR (@iDataType = 4)) 
			BEGIN
				/* Numeric/Integer column. */
				/* Replace the locale decimal separator with '.' for SQL's benefit. */
				SET @sValue = REPLACE(@sValue, @psDecimalSeparator, '.');
				IF (@iOperatorID = 1) 
				BEGIN
					/* Equals. */
					SET @sSubFilterSQL = @sColumnName + ' = ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <> ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' AND ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
        
				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' >= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					SET @sSubFilterSQL = @sColumnName + ' > ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 6) 
				BEGIN
					/* Less than.*/
					SET @sSubFilterSQL = @sColumnName + ' < ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
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
					SET @iIndex1 = CHARINDEX('mm', @psLocaleDateFormat);
					SET @iIndex2 = CHARINDEX('dd', @psLocaleDateFormat);
					SET @iIndex3 = CHARINDEX('yyyy', @psLocaleDateFormat);
						
					SET @sValue = SUBSTRING(@sValue, @iIndex1, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex2, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex3, 4);
				END

				IF (@iOperatorID = 1) 
				BEGIN
					/* Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' = ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END

				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					IF LEN(@sValue) > 0 
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <= ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' >= ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' > ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 6)
				BEGIN
					/* Less than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' < ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
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
						SET @sSubFilterSQL = @sColumnName + ' = '''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''' + @sValue + '''';
					END
				END
				
				IF (@iOperatorID = 2) 
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> '''' AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''' + @sValue + '''';
					END
				END

				IF (@iOperatorID = 7)
				BEGIN
					/* Contains */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''%' + @sValue + '%''';
					END
				END
				
				IF (@iOperatorID = 8) 
				BEGIN
					/* Does Not Contain. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''%' + @sValue + '%''';
					END
				END
			END
			
			IF LEN(@sSubFilterSQL) > 0
			BEGIN
				/* Add the filter code for this grid record into the complete filter code. */
				IF LEN(@sFilterSQL) > 0
				BEGIN
					SET @sFilterSQL = @sFilterSQL + ' AND (';
				END
				ELSE
				BEGIN
					SET @sFilterSQL = @sFilterSQL + '(';
				END
				SET @sFilterSQL = @sFilterSQL + @sSubFilterSQL + ')';
			END
		END
	END
	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT ASRSysColumns.tableID
		FROM [dbo].[ASRSysOrderItems]
		INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
		WHERE ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				CASE protectType
				        	WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;
	
	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.tableID,
		ASRSysOrderItems.columnID, 
		ASRSysColumns.columnName,
	    	ASRSysTables.tableName,
		ASRSysOrderItems.ascending,
		ASRSysOrderItems.type,
		ASRSysColumns.dataType,
		ASRSysColumns.size,
		ASRSysColumns.decimals,
		ASRSysColumns.Use1000Separator,
		ASRSysColumns.BlankIfZero
	FROM [dbo].[ASRSysOrderItems]
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
	ORDER BY ASRSysOrderItems.sequence;
	
	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END
	WHILE (@@fetch_status = 0)
	BEGIN

		SET @fSelectGranted = 0;
		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;
				
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName;

					SET @sSelectSQL = @sSelectSQL + @sTempString;
					SET @sThousandColumns = @sThousandColumns + convert(varchar(1),@bUse1000Separator);
					SET @sBlankIfZeroColumns = @sBlankIfZeroColumns + convert(varchar(1),@bBlankIfZero);
					
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							IF @piColumnType = 11 /* Date column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = isnull(convert(varchar(MAX), ' + @sRealSource + '.' + @sColumnName + ', 101), '''')';
							END
							ELSE
							IF @piColumnType = -7 /* Logic column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = replace(isnull(' + @sRealSource + '.' + @sColumnName + ', ''''), '''''''', '''''''''''')' ;
							END
							
							SET @sTempExecString = @sTempExecString
								+ ' FROM ' + @sRealSource
								+ ' WHERE ' + @sRealSource + '.ID = ' + @psLocateValue;
							SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
							EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
						END
					END
					
					SET @sOrderSQL = @sOrderSQL + CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName +
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END;
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1;
			END	
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName;
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + @sColumnTableName + '.' + @sColumnName;
					SET @sSelectSQL = @sSelectSQL + @sTempString;
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							SET @sTempLocateValue = '';
						END
					END
					SET @sOrderSQL = @sOrderSQL + CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName + 
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END;
				END
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @JoinParents
				WHERE tableViewName = @sColumnTableName;
				
				IF @iTempCount = 0
				BEGIN
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = '';
				
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT tableViewName
					FROM @ColumnPermissions
					WHERE tableID = @iColumnTableId
						AND tableViewName <> @sColumnTableName
						AND columnName = @sColumnName
						AND selectGranted = 1;
						
				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = 'CASE';
					SET @sSubString = @sSubString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @JoinParents
					WHERE tableViewname = @sViewName;
					
					IF @iTempCount = 0
					BEGIN
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId);
					END
					FETCH NEXT FROM viewCursor INTO @sViewName;
				END

				CLOSE viewCursor;
				DEALLOCATE viewCursor;

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +	' ELSE NULL END';
					IF @sType = 'F'
					BEGIN
						/* Find column. */
						SET @sTempString = CASE 
								WHEN (len(@sSelectSQL) > 0) THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END;
						SET @sSelectSQL = @sSelectSQL + @sTempString;
							
						SET @sTempString = ' AS [' + @sColumnName + ']';
						SET @sSelectSQL = @sSelectSQL + @sTempString;
						
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType;
							SET @piColumnSize = @iSize;
							SET @piColumnDecimals = @iDecimals;
							SET @fFirstColumnAsc = @fAscending;
							SET @sFirstColCode = @sSubString;
							IF (@psAction = 'LOCATEID')
							BEGIN
								SET @sTempLocateValue = '';
							END
						END
						SET @sOrderSQL = @sOrderSQL + CASE 
								WHEN len(@sOrderSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END + 
							CASE 
								WHEN @fAscending = 0 THEN ' DESC' 
								ELSE '' 
							END;
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1;
				END	
			END
		END
		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;
	END

	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	IF @psAction = 'LOCATEID'
	BEGIN
		SET @psAction = 'LOCATE';
		SET @psLocateValue = @sTempLocateValue;
	END
	
	/* Set the flags that show if no order columns could be selected, or if only some of them could be selected. */
	SET @pfSomeSelectable = CASE WHEN (len(@sSelectSQL) > 0) THEN 1 ELSE 0 END;
	SET @pfSomeNotSelectable = @fSelectDenied;

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN ',' ELSE '' END + 
		@sRealSource + '.ID';

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sRemainingSQL = @sOrderSQL;
		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(',', @sOrderSQL);
		WHILE @iCharIndex > 0 
		BEGIN
 			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', ';
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', ';
			END
			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(',', @sOrderSQL, @iLastCharIndex + 1);
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex);
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring;
	END
	
	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
	IF @piParentTableID > 0 
	BEGIN
		SET @sTempExecString = @sTempExecString + 
			' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
		IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
	END
	ELSE
	BEGIN
		IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
	END
	
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;

	SET @piTotalRecCount = @iCount;
	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sTempString = ',' + @sRealSource + '.ID';
		SET @sSelectSQL = @sSelectSQL + @sTempString;

		SET @sExecString = 'SELECT ' 
		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
		BEGIN
			SET @sTempString = 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
			SET @sExecString = @sExecString + @sTempString;
		END
	
		SET @sTempString = @sSelectSQL + ' FROM ' + @sRealSource;
		SET @sExecString = @sExecString + @sTempString;

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;
		
		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempString = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
			SET @sExecString = @sExecString + @sTempString;

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		
		CLOSE joinCursor;
		DEALLOCATE joinCursor;
		
		IF (@psAction = 'MOVELAST')
		BEGIN
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'MOVENEXT') 
		BEGIN
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
		END
		IF @psAction = 'MOVEPREVIOUS'
		BEGIN
			IF @piFirstRecPos <= @piRecordsRequired
			BEGIN
				SET @iGetCount = @piFirstRecPos - 1;
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;

		END
		/* Add the filter code. */
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			SET @sExecString = @sExecString + @sTempString;
				
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' AND ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' WHERE ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
		IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sReverseOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'LOCATE')
		BEGIN
			SET @sLocateCode = '';
			IF (@piParentTableID > 0) OR (len(@sFilterSQL) > 0) 
			BEGIN
				SET @sLocateCode = @sLocateCode + ' AND (' + @sFirstColCode;
			END
			ELSE
			BEGIN
				SET @sLocateCode = @sLocateCode + ' WHERE (' + @sFirstColCode;
			END;
			
			SET @sJoinSQL = '';
			DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName, 
				tableID
			FROM @JoinParents;
			OPEN joinCursor;
			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sJoinSQL = @sJoinSQL + 
					' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
				FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			END;
			
			CLOSE joinCursor;
			DEALLOCATE joinCursor;

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = replace(isnull(' + @sFirstColCode + ', ''''), '''''''', '''''''''''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL';
						
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' <= ''' + @sTempLocateValue + '''';
					END;
				END
			END
			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = isnull(convert(varchar(MAX),' + @sFirstColCode + ', 101), '''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sFirstColCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
						END
					END;
				END
				ELSE
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + @sFirstColCode + ' IS NULL';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ''' + @sTempLocateValue + '''';
						END
					END;
				END
			END
			IF @piColumnType = -7 /* Logic column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;
				
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
			END
			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue;
					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + @sTempLocateValue;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL';

					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + @sTempLocateValue ;
					END;
				END
			END
			SET @sLocateCode = @sLocateCode + ')';
			SET @sTempString = @sLocateCode;
			SET @sExecString = @sExecString + @sTempString;
		END

		--/* Add in the string to include max number of records available in block */				
		--/* Fault HRPRO-382 */
		--IF @psAction = 'LOCATE'
		--BEGIN
		--	SET @sTempString = ' OR '+@sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource
		--	SET @sTempFilterString = ''
			
		--	IF @piParentTableID > 0 SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						

		--	IF LEN(@sFilterSQL) > 0
		--	BEGIN
		--		IF LEN(@sTempFilterString) > 0
		--			SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	
		--		ELSE
		--			SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	
		--	END
			
		--	--use for 'go to' only:
		--	IF @psOriginalAction = 'LOCATE'
		--	BEGIN
		--		IF LEN(@sLocateCode) > 0 SET @sTempFilterString = @sTempFilterString + @sLocateCode	
		--	END
			
		--	SET @sTempString = @sTempString + @sTempFilterString
		--	SET @sTempString = @sTempString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
		--	SET @sExecString = @sExecString + @sTempString;		
		--END
		
		/* Add the ORDER BY code to the find record selection string if required. */
		SET @sTempString = ' ORDER BY ' + @sOrderSQL;
		SET @sExecString = @sExecString + @sTempString;
				
	END

	/* Check if the user has insert or delete permission on the table. */
	SET @pfInsertGranted = 0;
	SET @pfDeleteGranted = 0;
	IF LEN(@sRealSource) > 0
	BEGIN
		DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT p.action
		FROM @sysprotects p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		WHERE p.protectType <> 206
			AND ((p.action = 195) OR (p.action = 196))
			AND sysobjects.name = @sRealSource;
			
		OPEN tableInfo_cursor;
		FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iTempAction = 195
			BEGIN
				SET @pfInsertGranted = 1;
			END
			ELSE
			BEGIN
				SET @pfDeleteGranted = 1;
			END
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		END
		
		CLOSE tableInfo_cursor;
		DEALLOCATE tableInfo_cursor;
		
	END
	
	/* Set the IsFirstPage, IsLastPage flags, and the page number. */
	IF @psAction = 'MOVEFIRST'
	BEGIN
		SET @piFirstRecPos = 1;
		SET @pfFirstPage = 1;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVENEXT'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount;
		SET @pfFirstPage = 0;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos - @iGetCount;
		IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVELAST'
	BEGIN
		SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1;
		IF @piFirstRecPos < 1 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 1;
	END
	IF @psAction = 'LOCATE'
	BEGIN
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		
		CLOSE joinCursor;
		DEALLOCATE joinCursor;
		
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
		END
		
		SET @sTempExecString = @sTempExecString + @sLocateCode;
		SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iTemp OUTPUT;
		IF @iTemp <=0 
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount + 1;
		END
		ELSE
		BEGIN
			--IF @psOriginalAction = 'LOCATEID'	
			--BEGIN				
			--	SET @piFirstRecPos = @piTotalRecCount - CASE WHEN @piRecordsRequired > @piTotalRecCount THEN @piTotalRecCount - 1 ELSE @piRecordsRequired END;
			--END
			--ELSE 
			--BEGIN
				SET @piFirstRecPos = @piTotalRecCount - @iTemp + 1;
			--END
		END
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	
	-- Return a recordset of the required columns in the required order from the given table/view.
	IF (@pfSomeSelectable = 1)
	BEGIN

		SELECT @sBlankIfZeroColumns AS BlankIfZeroColumns
			, @sThousandColumns AS ThousandColumns

		EXECUTE sp_executeSQL @sExecString;
	END


END
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetScreenControlsString2]    Script Date: 02/01/2014 20:50:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenControlsString2] (
	@piScreenID 	integer,
	@piViewID 		integer,
	@psSelectSQL	varchar(MAX) OUTPUT,
	@psFromDef		varchar(MAX) OUTPUT,
	@piOrderID		integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@iScreenTableID		integer,
		@iScreenTableType	integer,
		@sScreenTableName	varchar(255),
		@iScreenOrderID 	integer,
		@sRealSource 		varchar(255),
		@iChildViewID 		integer,
		@sJoinCode 			varchar(MAX),
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@sColumnTableName 	varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@fUpdateGranted 	bit,
		@sSelectString 		varchar(MAX),
		@iTempCount 		integer,
		@sViewName 			varchar(255),
		@fAscending 		bit,
		@sTableViewName 	varchar(255),
		@iJoinTableID 		integer,
		@sParentRealSource	varchar(255),
		@iParentChildViewID	integer,
		@iParentTableType	integer,
		@sParentTableName	sysname,
		@iColumnType		integer,
		@iLinkTableID		integer,
		@lngPermissionCount	integer,
		@iLinkChildViewID	integer,
		@sLinkRealSource	varchar(255),
		@sLinkTableName		varchar(255),
		@iLinkTableType		integer,
		@sNewBit			varchar(max),
		@iID				integer,
		@iCount				integer,
		@iUserType			integer,
		@sRoleName			sysname,
		@iEmployeeTableID	integer,
		@sActualUserName	sysname,
		@AppName varchar(50),
		@ItemKey varchar(20);

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;


	DECLARE @SysProtects TABLE([ID] int, Columns varbinary(8000)
								, [Action] tinyint
								, ProtectType tinyint)
	INSERT INTO @SysProtects
	SELECT p.[ID], p.[Columns], p.[Action], p.ProtectType FROM ASRSysProtectsCache p
		INNER JOIN SysColumns c ON (c.id = p.id
			AND c.[Name] = 'timestamp'
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0)))
		WHERE p.UID = @iUserGroupID
			AND [ProtectType] IN (204, 205)
			AND [Action] IN (193, 197);


	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @JoinParents TABLE(tableViewName	sysname,
								tableID		integer);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @ColumnPermissions TABLE(tableID		integer,
										tableViewName	sysname,
										columnName	sysname,
										action		int,		
										granted		bit);


	/* Get the table type and name. */
	SELECT @iScreenTableID = ASRSysScreens.tableID,
		@iScreenTableType = ASRSysTables.tableType,
		@sScreenTableName = ASRSysTables.tableName,
		@iScreenOrderID = 
				CASE 
					WHEN ASRSysScreens.orderID > 0 THEN ASRSysScreens.orderID
					ELSE ASRSysTables.defaultOrderID 
				END
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.ScreenID = @piScreenID;

	IF @iScreenOrderID IS NULL SET @iScreenOrderID = 0;

	IF @piOrderID <= 0 SET @piOrderID = @iScreenOrderID;

	/* Get the real source of the given screen's table/view. */
	IF @iScreenTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @piViewID;
		END
		ELSE
		BEGIN
			/* RealSource is the table. */	
			SET @sRealSource = @sScreenTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @iScreenTableID
			AND role = @sRoleName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sScreenTableName, ' ', '_') +
				'#' + replace(@sRoleName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END

	/* Initialise the select and order parameters. */
	SET @psSelectSQL = '';
	SET @psFromDef = '';
	SET @sJoinCode = '';

	/* Loop through the tables used in the screen, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID
	FROM ASRSysControls
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0
	UNION
	SELECT DISTINCT ASRSysColumns.tableID 
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	WHERE ASRSysOrderItems.type = 'O' 
		AND ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @iScreenTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				p.action,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			SELECT @iParentTableType = tableType,
				@sParentTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @iTempTableID

			IF @iParentTableType <> 2 /* ie. top-level or lookup */
			BEGIN
				INSERT INTO @ColumnPermissions
				SELECT 
					@iTempTableID,
					sysobjects.name,
					syscolumns.name,
					p.[action],
					CASE p.protectType
					        	WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @sysprotects p
				INNER JOIN sysobjects ON p.id = sysobjects.id
				INNER JOIN syscolumns ON p.id = syscolumns.id
				WHERE syscolumns.name <> 'timestamp'
					AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @iTempTableID 
						UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			END
			ELSE
			BEGIN
				/* Get permitted child view on the parent table. */
				SELECT @iParentChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTempTableID
					AND role = @sRoleName
					
				IF @iParentChildViewID IS null SET @iParentChildViewID = 0
					
				IF @iParentChildViewID > 0 
				BEGIN
					SET @sParentRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iParentChildViewID) +
						'#' + replace(@sParentTableName, ' ', '_') +
						'#' + replace(@sRoleName, ' ', '_')
					SET @sParentRealSource = left(@sParentRealSource, 255)

					INSERT INTO @ColumnPermissions
					SELECT 
						@iTempTableID,
						@sParentRealSource,
						syscolumns.name,
						p.[action],
						CASE p.protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @sysprotects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					WHERE syscolumns.name <> 'timestamp'
						AND sysobjects.name = @sParentRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				END
			END
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	SET @iUserType = 1

	/*Ascertain application name in order to select by correct item key  */
	SELECT @AppName = APP_NAME()
	IF @AppName = 'OPENHR SELF-SERVICE INTRANET'
	BEGIN
		SET @ItemKey = 'SSINTRANET'
	END
	ELSE
	BEGIN
		SET @ItemKey = 'INTRANET'
	END

	SELECT @iID = ASRSysPermissionItems.itemID
	FROM ASRSysPermissionItems
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	WHERE ASRSysPermissionItems.itemKey = @ItemKey
		AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

	IF @iID IS NULL SET @iID = 0
	IF @iID > 0
	BEGIN
		/* The permission does exist in the current version so check if the user is granted this permission. */
		SELECT @iCount = count(ASRSysGroupPermissions.itemID)
		FROM ASRSysGroupPermissions 
		WHERE ASRSysGroupPermissions.itemID = @iID
			AND ASRSysGroupPermissions.groupName = @sRoleName
			AND ASRSysGroupPermissions.permitted = 1
			
		IF @iCount > 0
		BEGIN
			SET @iUserType = 0
		END
	END

	/* Get the EMPLOYEE table information. */
	SELECT @iEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel'
	IF @iEmployeeTableID IS NULL SET @iEmployeeTableID = 0

	/* Create a temporary table of the column info for all columns used in the screen controls. */
	DECLARE @columnInfo TABLE
	(
		columnID	integer,
		selectGranted	bit,
		updateGranted	bit
	)

	/* Populate the temporary table with info for all columns used in the screen controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID, 
		ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysTables.tableName,
		ASRSysColumns.dataType,
		ASRSysColumns.columnType,
		ASRSysColumns.linkTableID
	FROM ASRSysControls
	LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
	LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID AND ASRSysColumns.columnId = ASRSysControls.columnID
	WHERE screenID = @piScreenID
	AND ASRSysControls.columnID > 0

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0
		SET @fUpdateGranted = 0

		IF @iColumnTableID = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193

			/* Get the update permission on the column. */
			SELECT @fUpdateGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 197

			/* If the column is a link column, ensure that the link table can be seen. */
			IF (@fUpdateGranted = 1) AND (@iColumnType = 4)
			BEGIN
				SELECT @sLinkTableName = tableName,
					@iLinkTableType = tableType
				FROM ASRSysTables
				WHERE tableID = @iLinkTableID

				IF @iLinkTableType = 1
				BEGIN
					/* Top-level table. */
					SELECT @lngPermissionCount = COUNT(sysprotects.uid)
					FROM sysprotects
					INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
					INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
					WHERE sysprotects.uid = @iUserGroupID
						AND sysprotects.action = 193
						AND sysprotects.protectType <> 206
						AND syscolumns.name <> 'timestamp'
						AND syscolumns.name <> 'ID'
						AND sysobjects.name = @sLinkTableName
						AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					IF @lngPermissionCount = 0 
					BEGIN
						/* No permission on the table itself check the views. */
						SELECT @lngPermissionCount = COUNT(ASRSysViews.viewTableID)
						FROM ASRSysViews
						INNER JOIN sysobjects ON ASRSysViews.viewName = sysobjects.name
						INNER JOIN sysprotects ON sysobjects.id = sysprotects.id  
						WHERE ASRSysViews.viewTableID = @iLinkTableID
							AND sysprotects.uid = @iUserGroupID
							AND sysprotects.action = 193
							AND sysprotects.protecttype <> 206

						IF @lngPermissionCount = 0 SET @fUpdateGranted = 0
					END
				END
				ELSE
				BEGIN
					/* Child/history table. */
					SELECT @iLinkChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @iLinkTableID
						AND role = @sRoleName
						
					IF @iLinkChildViewID IS null SET @iLinkChildViewID = 0
						
					IF @iLinkChildViewID > 0 
					BEGIN
						SET @sLinkRealSource = 'ASRSysCV' + 
							convert(varchar(1000), @iLinkChildViewID) +
							'#' + replace(@sLinkTableName, ' ', '_') +
							'#' + replace(@sRoleName, ' ', '_')
						SET @sLinkRealSource = left(@sLinkRealSource, 255)
					END

					SELECT @lngPermissionCount = COUNT(sysobjects.name)
					FROM sysprotects 
					INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
					WHERE sysprotects.uid = @iUserGroupID
						AND sysprotects.protectType <> 206
						AND sysprotects.action = 193
						AND sysobjects.name = @sLinkRealSource
		
					IF @lngPermissionCount = 0 SET @fUpdateGranted = 0
				END
			END

			IF @fSelectGranted = 1 
			BEGIN
				/* Get the select string for the column. */
				IF len(@psSelectSQL) > 0 
					SET @psSelectSQL = @psSelectSQL + ',';
			
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sNewBit = 'convert(varchar(10), ' + @sRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(255), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sNewBit = @sRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
			END
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sColumnTableName
				AND columnName = @sColumnName
				AND action = 193;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				IF len(@psSelectSQL) > 0 
					SET @psSelectSQL = @psSelectSQL + ',';
	
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sNewBit = 'convert(varchar(10), ' + @sColumnTableName + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sNewBit = @sColumnTableName + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
			
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
					FROM @JoinParents
					WHERE tableViewName = @sColumnTableName;

				IF @iTempCount = 0
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
					
			END
			ELSE	
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */
				SET @sSelectString = '';

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @ColumnPermissions
				WHERE tableID = @iColumnTableID
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND action = 193
					AND granted = 1;

				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)

				BEGIN
					/* Column CAN be read from the view. */
					SET @fSelectGranted = 1;

					IF len(@sSelectString) = 0 SET @sSelectString = 'CASE';
	
					IF @iColumnDataType = 11 /* Date */
					BEGIN
						 /* Date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)';
					END
					ELSE
					BEGIN
						 /* Non-date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
					END

					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
						FROM @JoinParents
						WHERE tableViewName = @sViewName;

					IF @iTempCount = 0
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableID);

					FETCH NEXT FROM viewCursor INTO @sViewName;
				END
				CLOSE viewCursor;
				DEALLOCATE viewCursor;

				IF len(@sSelectString) > 0
				BEGIN
					SET @sSelectString = @sSelectString +
						' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']';

					IF len(@psSelectSQL) > 0 
						SET @psSelectSQL = @psSelectSQL + ',';

					SET @psSelectSQL = @psSelectSQL + @sSelectString;
				END
			END

			/* Reset the update permission on the column. */
			SET @fUpdateGranted = 0
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0
		IF @fUpdateGranted IS NULL SET @fUpdateGranted = 0

		IF (@iUserType = 1) 
			AND (@iScreenTableType = 1)
			AND (@iScreenTableID <> @iEmployeeTableID)
		BEGIN
			SET @fUpdateGranted = 0
		END

		INSERT INTO @columnInfo (columnID, selectGranted, updateGranted)
			VALUES (@iColumnId, @fSelectGranted, @fUpdateGranted)

		FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor

	/* Create the order string. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.tableID,
		ASRSysOrderItems.columnID, 
		ASRSysColumns.columnName,
	    	ASRSysTables.tableName,
		ASRSysOrderItems.ascending
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
		AND ASRSysOrderItems.type = 'O'
	ORDER BY ASRSysOrderItems.sequence

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF @iColumnTableId = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName
				AND action = 193

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @JoinParents
				WHERE tableViewName = @sColumnTableName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID)
				END
			END
			ELSE	
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @ColumnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND action = 193
					AND granted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @JoinParents
					WHERE tableViewname = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Add the id and timestamp columns to the select string. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnId, 
		ASRSysColumns.columnName
	FROM ASRSysColumns
	WHERE tableID = @iScreenTableID
		AND columnType = 3

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName
	WHILE (@@fetch_status = 0)
	BEGIN
		IF len(@psSelectSQL) > 0 
			SET @psSelectSQL = @psSelectSQL + ',';

		SET @sNewBit = @sRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
		SET @psSelectSQL = @psSelectSQL + @sNewBit;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	SET @sNewBit = ', CONVERT(integer, ' + @sRealSource + '.TimeStamp) AS timestamp ';
	SET @psSelectSQL = @psSelectSQL + @sNewBit;

	/* Create the FROM code. */
	SET @psFromDef = @sRealSource + '	'
	DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableViewName, tableID
		FROM @JoinParents;

	OPEN joinCursor;
	FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @psFromDef = @psFromDef + @sTableViewName + '	' + convert(varchar(100), @iJoinTableID) + '	';
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
	END
	CLOSE joinCursor;
	DEALLOCATE joinCursor;

	SELECT
		convert(varchar(MAX), case when ASRSysControls.pageNo IS null then '' else ASRSysControls.pageNo end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.tableID IS null then '' else ASRSysControls.tableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.columnID IS null then '' else ASRSysControls.columnID end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.controlType IS null then '' else ASRSysControls.controlType end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.topCoord IS null then '' else ASRSysControls.topCoord end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.leftCoord IS null then '' else ASRSysControls.leftCoord end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.height IS null then '' else ASRSysControls.height end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.width IS null then '' else ASRSysControls.width end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.caption IS null then '' else ASRSysControls.caption end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.backColor IS null then '' else ASRSysControls.backColor end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.foreColor IS null then '' else ASRSysControls.foreColor end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontName IS null then '' else ASRSysControls.fontName end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontSize IS null then '' else ASRSysControls.fontSize end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontBold IS null then '' else ASRSysControls.fontBold end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontItalic IS null then '' else ASRSysControls.fontItalic end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontStrikethru IS null then '' else ASRSysControls.fontStrikethru end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontUnderline IS null then '' else ASRSysControls.fontUnderline end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.displayType IS null then '' else ASRSysControls.displayType end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.tabIndex IS null then '' else ASRSysControls.tabIndex end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.borderStyle IS null then '' else ASRSysControls.borderStyle end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.alignment IS null then '' else ASRSysControls.alignment end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnName IS null then '' else ASRSysColumns.columnName end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnType IS null then '' else ASRSysColumns.columnType end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.datatype IS null then '' else ASRSysColumns.datatype end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.defaultValue IS null then '' else ASRSysColumns.defaultValue end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.size IS null then '' else convert(nvarchar(max),ASRSysColumns.size) end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.decimals IS null then '' else ASRSysColumns.decimals end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.lookupTableID IS null then '' else ASRSysColumns.lookupTableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.lookupColumnID IS null then '' else ASRSysColumns.lookupColumnID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerMinimum IS null then '' else ASRSysColumns.spinnerMinimum end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerMaximum IS null then '' else ASRSysColumns.spinnerMaximum end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerIncrement IS null then '' else ASRSysColumns.spinnerIncrement end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.mandatory IS null then '' else ASRSysColumns.mandatory end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.uniquechecktype IS null then '' when ASRSysColumns.uniquechecktype <> 0 then 1 else 0 end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.convertcase IS null then '' else ASRSysColumns.convertcase end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.mask IS null then '' else rtrim(ASRSysColumns.mask) end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.blankIfZero IS null then '' else ASRSysColumns.blankIfZero end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.multiline IS null then '' else ASRSysColumns.multiline end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.alignment IS null then '' else ASRSysColumns.alignment end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.dfltValueExprID IS null then '' else ASRSysColumns.dfltValueExprID end) + char(9) +
		convert(varchar(MAX), case when isnull(ASRSysColumns.readOnly,0) = 1 then 1 else 0 end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.statusBarMessage IS null then '' else ASRSysColumns.statusBarMessage end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkTableID IS null then '' else ASRSysColumns.linkTableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkOrderID IS null then '' else ASRSysColumns.linkOrderID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkViewID IS null then '' else ASRSysColumns.linkViewID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.Afdenabled IS null then '' else ASRSysColumns.Afdenabled end) + char(9) +
		convert(varchar(MAX), case when ASRSysTables.TableName IS null then '' else ASRSysTables.TableName end) + char(9) +
		convert(varchar(MAX), case when ci.selectGranted IS null then '' else ci.selectGranted end) + char(9) +
		convert(varchar(MAX), case when ci.updateGranted IS null then '' else ci.updateGranted end) + char(9) +
		'' + char(9) +
		convert(varchar(MAX), case when ASRSysControls.pictureID IS null then '' else ASRSysControls.pictureID end)+ char(9) +
		convert(varchar(MAX), case when ASRSysColumns.trimming IS null then '' else ASRSysColumns.trimming end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.Use1000Separator IS null then '' else ASRSysColumns.Use1000Separator end) + char(9) +	
		convert(varchar(MAX), case when ASRSysColumns.lookupFilterColumnID IS null then '' else ASRSysColumns.lookupFilterColumnID end) + char(9) +	
		convert(varchar(MAX), case when ASRSysColumns.LookupFilterValueID IS null then '' else ASRSysColumns.LookupFilterValueID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.OLEType IS null then '' else ASRSysColumns.OLEType end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.MaxOLESizeEnabled IS null then '' else ASRSysColumns.MaxOLESizeEnabled end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.MaxOLESize IS null then '' else ASRSysColumns.MaxOLESize end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateTo IS null then '' else ASRSysControls.NavigateTo end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateIn IS null then '' else ASRSysControls.NavigateIn end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateOnSave IS null then '' else ASRSysControls.NavigateOnSave end) + char(9) +
		convert(varchar(MAX), case when isnull(ASRSysControls.readOnly,0) = 1 then 1 else 0 end)
		AS [controlDefinition],
		ASRSysControls.pageNo AS [pageNo],
		ASRSysControls.controlLevel AS [controlLevel],
		ASRSysControls.tabIndex AS [tabIndex]
	FROM ASRSysControls
	LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
	LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID AND ASRSysColumns.columnId = ASRSysControls.columnID
	LEFT OUTER JOIN @columnInfo ci ON ASRSysColumns.columnId = ci.columnID
	WHERE screenID = @piScreenID
	UNION
	SELECT 
		convert(varchar(MAX), -1) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnId IS null then '' else ASRSysColumns.columnId end)  + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnName IS null then '' else ASRSysColumns.columnName end) 
		AS [controlDefinition],
		0 AS [pageNo],
		0 AS [controlLevel],
		0 AS [tabIndex]
	FROM ASRSysColumns
	WHERE tableID = @iScreenTableID
		AND columnType = 3
	ORDER BY [pageNo],
		[controlLevel] DESC, 
		[tabIndex];

END
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetScreenDefinition]    Script Date: 02/01/2014 20:50:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenDefinition] (
	@piScreenID 		integer,
	@piViewID			integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @iTabCount 		integer,
		@sTabCaptions		varchar(MAX),
		@sTabCaption		varchar(MAX),
		@fSysSecMgr			bit,
		@fInsertGranted		bit,
		@fDeleteGranted		bit,
		@sRealSource		sysname,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@iTableID			integer,
		@iTableType			integer,
		@sTableName			sysname,
		@iTempAction		integer,
		@iChildViewID 		integer,
		@sActualUserName	varchar(250);

	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT
					
	/* Get the table type and name. */
	SELECT @iTableID = ASRSysScreens.tableID,
		@iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.ScreenID = @piScreenID

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
	OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

	/* Get the real source and insert/delete permissions for the table. */
	IF @fSysSecMgr = 1 
	BEGIN
		/* Permission must be granted for System or Security mangers. */
		SET @fInsertGranted = 1
		SET @fDeleteGranted = 1	

		/* Get the realSource of the table. */
		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @piViewID > 0 
			BEGIN
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM ASRSysViews
				WHERE viewID = @piViewID	
			END
			ELSE
			BEGIN
				/* RealSource is the table. */	
				SET @sRealSource = @sTableName
			END 
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
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
	END
	ELSE
	BEGIN

		/* Permission must be read from the database  for Non-System and Non-Security mangers. */
		SET @fInsertGranted = 0
		SET @fDeleteGranted = 0	

		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @piViewID > 0 
			BEGIN	
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM ASRSysViews
				WHERE viewID = @piViewID
			END
			ELSE
			BEGIN
				SET @sRealSource = @sTableName
			END 

			/* Get the insert/delete permissions for the realSource. */
			DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT p.action
				FROM ASRSysProtectsCache p
				INNER JOIN sysobjects ON p.id = sysobjects.id
				WHERE p.UID = @iUserGroupID AND p.action  IN (195, 196)
					AND sysobjects.name = @sRealSource
					AND ProtectType <> 206

			OPEN tableInfo_cursor
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iTempAction = 195
				BEGIN
					SET @fInsertGranted = 1
				END
				ELSE
				BEGIN
					SET @fDeleteGranted = 1	
				END
				FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			END
			CLOSE tableInfo_cursor
			DEALLOCATE tableInfo_cursor
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sRealSource = left(@sRealSource, 255)

				/* Get appropriate child view if required. */
				DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT p.action
					FROM ASRSysProtectsCache p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					WHERE sysobjects.name = @sRealSource
						AND p.UID = @iUserGroupID
						AND p.Action IN(193, 195, 196)
						AND ProtectType <> 206

				OPEN tableInfo_cursor
				FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iTempAction = 195
					BEGIN
						SET @fInsertGranted = 1
					END
					ELSE
					BEGIN
						IF @iTempAction = 196
						BEGIN
							SET @fDeleteGranted = 1	
						END
					END
					FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
				END
				CLOSE tableInfo_cursor
				DEALLOCATE tableInfo_cursor
			END
		END
	END
	
	/* Get the tab page captions info. */
	SET @iTabCount = 0
	SET @sTabCaptions = ''
	
	DECLARE captions_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT caption 
		FROM ASRSysPageCaptions
		WHERE screenID = @piScreenID
		ORDER BY pageIndexID

	OPEN captions_cursor
	FETCH NEXT FROM captions_cursor INTO @sTabCaption
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTabCount > 0 SET @sTabCaptions = @sTabCaptions + char(9) 

		SET @iTabCount = @iTabCount + 1
		SET @sTabCaptions = @sTabCaptions + @sTabCaption
			
		FETCH NEXT FROM captions_cursor INTO @sTabCaption
	END
	CLOSE captions_cursor
	DEALLOCATE captions_cursor

	SELECT @sTableName AS tableName,
		@sRealSource AS realSource,
		@fInsertGranted AS insertGranted,
		@fDeleteGranted AS deleteGranted,
		height,
		width,
		fontName,
		fontSize,
		fontBold,
		fontItalic,
		fontStrikethru,
		fontUnderline,
		@iTabCount AS tabCount,
		@sTabCaptions AS tabCaptions
	FROM ASRSysScreens
	WHERE screenID = @piScreenID
	
END
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetSummaryFields]    Script Date: 02/01/2014 20:50:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetSummaryFields] (
	@piHistoryTableID	integer,
	@piParentTableID 	integer,
	@piParentRecordID	integer,
	@pfCanSelect		bit OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(255),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(255),
		@sColumnName 		varchar(255),
		@fSelectGranted 	bit,
		@iCount				integer,
		@sActualUserName	sysname;

	SET @pfCanSelect = 0;

	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
		OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID

	/* Create a temporary table of the 'read' column permissions for all tables/views used. */
	DECLARE @ColumnPermissions TABLE(
				tableViewName	sysname,
				columnName	sysname,
				granted		bit);

	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO @ColumnPermissions
		SELECT 
			@sParentTableName,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns 
		WHERE ASRSysColumns.tableID = @piParentTableID
	END
	ELSE
	BEGIN
		IF @iParentTableType <> 2 /* ie. top-level or lookup */
		BEGIN

			-- Get list of views/table columns that are summary fields
			DECLARE @SummaryColumns TABLE ([ID] int, [TableName] sysname, [ColumnName] sysname, [ColID] int)
			INSERT @SummaryColumns
				SELECT sysobjects.id, sysobjects.name,
					syscolumns.name, syscolumns.ColID
				FROM sysobjects
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.name IN (SELECT ASRSysTables.tableName
												FROM ASRSysTables
												WHERE ASRSysTables.tableID = @piParentTableID
											UNION SELECT ASRSysViews.viewName
												FROM ASRSysViews
												WHERE ASRSysViews.viewTableID = @piParentTableID)
					AND syscolumns.name IN (SELECT ac.ColumnName
												FROM ASRSysSummaryFields am
												INNER JOIN ASRSysColumns ac ON am.ParentColumnID = ac.ColumnID
												WHERE HistoryTableID = @piHistoryTableID)

			-- Generate security context on selected columns
			INSERT INTO @ColumnPermissions
				SELECT sm.TableName,
					sm.ColumnName,
					CASE p.protectType
						WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM ASRSysProtectsCache p
				INNER JOIN @SummaryColumns sm ON p.id = sm.id
				WHERE p.UID = @iUserGroupID
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) = 0))
					AND p.Action = 193

		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @piParentTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sParentRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sParentTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sParentRealSource = left(@sParentRealSource, 255)
			END

			INSERT INTO @ColumnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM ASRSysProtectsCache p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.UID = @iUserGroupID
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				AND p.Action = 193
		END
	END

	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnName
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnId
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @sColumnName
	WHILE (@@fetch_status = 0) AND (@pfCanSelect = 0)
	BEGIN
		SET @fSelectGranted = 0

		/* Get the select permission on the column. */

		/* Check if the column is selectable directly from the table. */
		SELECT @iCount = COUNT(*)
		FROM @ColumnPermissions
		WHERE columnName = @sColumnName
			AND granted = 1

		IF @iCount > 0 
		BEGIN
			SET @pfCanSelect = 1
		END

		FETCH NEXT FROM columnsCursor INTO @sColumnName
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor


	SELECT DISTINCT ASRSysSummaryFields.sequence, 
    	ASRSysSummaryFields.startOfGroup, 
		ASRSysColumns.columnName, 
		ASRSysColumns.columnId, 
		ASRSysColumns.dataType, 
		ASRSysColumns.size, 
		ASRSysColumns.decimals, 
		ASRSysColumns.controlType, 
		ASRSysColumns.alignment,
		ASRSysColumns.Use1000Separator
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnId
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence;
	
END
GO

/****** Object:  StoredProcedure [dbo].[sp_ASRIntGetTrainingBookingParameters]    Script Date: 02/01/2014 20:50:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetTrainingBookingParameters] (
	@piEmployeeTableID			integer	OUTPUT,
	@piCourseTableID			integer	OUTPUT,
	@piCourseCancelDateColumnID	integer	OUTPUT,
	@piTBTableID				integer	OUTPUT,
	@pfTBTableSelect			bit		OUTPUT,
	@pfTBTableInsert			bit		OUTPUT,
	@pfTBTableUpdate			bit		OUTPUT,
	@piTBStatusColumnID			integer	OUTPUT,
	@pfTBStatusColumnUpdate		bit		OUTPUT,
	@piTBCancelDateColumnID		integer	OUTPUT,
	@pfTBCancelDateColumnUpdate	bit		OUTPUT,
	@pfTBProvisionalStatusExists	bit	OUTPUT,
	@piWaitListTableID			integer	OUTPUT,
	@pfWaitListTableInsert			bit	OUTPUT,
	@pfWaitListTableDelete			bit	OUTPUT,
	@piWaitListCourseTitleColumnID		integer	OUTPUT,
	@pfWaitListCourseTitleColumnUpdate	bit	OUTPUT,
	@pfWaitListCourseTitleColumnSelect	bit	OUTPUT,
	@piBulkBookingDefaultViewID		integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @fOK			bit,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sTempName			sysname,
		@iTempAction		integer,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@sRealSource		sysname,
		@iStatusCount		integer,
		@iChildViewID		integer,
		@sTBTableName		sysname,
		@sWLTableName		sysname,
		@sActualUserName	sysname;
		
	/* Training Booking information. */
	SET @fOK = 1

	SET @piEmployeeTableID = 0

	SET @piCourseTableID = 0
	SET @piCourseCancelDateColumnID = 0

	SET @piTBTableID = 0
	SET @pfTBTableSelect = 0
	SET @pfTBTableInsert = 0
	SET @pfTBTableUpdate = 0
	SET @piTBStatusColumnID = 0
	SET @pfTBStatusColumnUpdate = 0
	SET @piTBCancelDateColumnID = 0
	SET @pfTBCancelDateColumnUpdate = 0
	SET @pfTBProvisionalStatusExists = 0

	SET @piWaitListTableID = 0
	SET @pfWaitListTableInsert = 0
	SET @pfWaitListTableDelete = 0
	SET @piWaitListCourseTitleColumnID = 0
	SET @pfWaitListCourseTitleColumnUpdate = 0
	SET @pfWaitListCourseTitleColumnSelect = 0

	SET @piBulkBookingDefaultViewID = 0
	
	/* Get the current user's group id. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
	OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'TRAINING', @fOK OUTPUT

	/* Get the required training booking module paramaters. */
	IF @fOK = 1
	BEGIN
		/* Get the EMPLOYEE table information. */
		SELECT @piEmployeeTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_EmployeeTable'
		IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0

		/* Get the COURSE table information. */
		SELECT @piCourseTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseTable'
		IF @piCourseTableID IS NULL SET @piCourseTableID = 0

		IF @piCourseTableID > 0
		BEGIN
			SELECT @piCourseCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_CourseCancelDate'
			IF @piCourseCancelDateColumnID IS NULL SET @piCourseCancelDateColumnID = 0
		END

		/* Get the TRAINING BOOKING table information. */
		SELECT @piTBTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_TrainBookTable'
		IF @piTBTableID IS NULL SET @piTBTableID = 0


		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID]				int,
								   [columns]		varbinary(8000),
								   [Action]			tinyint,
								   [ProtectType]	tinyint)
		INSERT INTO @SysProtects
		SELECT [ID], [Columns], [Action], [ProtectType] FROM ASRSysProtectsCache
			WHERE [UID] = @iUserGroupID AND [Action] IN (193, 195, 196, 197)

		IF @piTBTableID > 0
		BEGIN
			SELECT @sTBTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piTBTableID

			SELECT @piTBStatusColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookStatus'
			IF @piTBStatusColumnID IS NULL SET @piTBStatusColumnID = 0

			SELECT @piTBCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookCancelDate'
			IF @piTBCancelDateColumnID IS NULL SET @piTBCancelDateColumnID = 0

			SET @sCommand = 'SELECT @iStatusCount = COUNT(*)' +
				' FROM ASRSysColumnControlValues' +
				' WHERE columnID = ' + convert(nvarchar(100), @piTBStatusColumnID) +
				' AND value = ''P'''
			SET @sParamDefinition = N'@iStatusCount integer OUTPUT'
			EXEC sp_executesql @sCommand, @sParamDefinition, @iStatusCount OUTPUT
			IF @iStatusCount > 0 SET @pfTBProvisionalStatusExists = 1

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				/* System/Security managers must have all permissions granted. */
				SET @pfTBTableSelect = 1
				SET @pfTBTableInsert = 1
				SET @pfTBTableUpdate = 1
				SET @pfTBStatusColumnUpdate = 1
				SET @pfTBCancelDateColumnUpdate = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piTBTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN

					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTBTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (193, 195, 197)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN

						IF @iTempAction = 193
						BEGIN
							SET @pfTBTableSelect = 1
						END
						IF @iTempAction = 195
						BEGIN
							SET @pfTBTableInsert = 1
						END
						IF @iTempAction = 197
						BEGIN
							SET @pfTBTableUpdate = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @tbColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @tbColumnPermissions
					SELECT 
						ASRSysColumns.columnId,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piTBTableID
						AND (ASRSysColumns.columnId = @piTBStatusColumnID
							OR ASRSysColumns.columnId = @piTBCancelDateColumnID))
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfTBStatusColumnUpdate = granted

					FROM @tbColumnPermissions
					WHERE columnID = @piTBStatusColumnID
						AND action = 197
					IF @pfTBStatusColumnUpdate IS NULL SET @pfTBStatusColumnUpdate = 0

					SELECT @pfTBCancelDateColumnUpdate = granted
					FROM @tbColumnPermissions
					WHERE columnID = @piTBCancelDateColumnID
						AND action = 197
					IF @pfTBCancelDateColumnUpdate IS NULL SET @pfTBCancelDateColumnUpdate = 0

				END
			END
		END

		/* Get the waiting list table information. */
		SELECT @piWaitListTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_WaitListTable'
		IF @piWaitListTableID IS NULL SET @piWaitListTableID = 0

		IF @piWaitListTableID > 0
		BEGIN
			SELECT @sWLTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piWaitListTableID

			SELECT @piWaitListCourseTitleColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_WaitListCourseTitle'
			IF @piWaitListCourseTitleColumnID IS NULL SET @piWaitListCourseTitleColumnID = 0

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				SET @pfWaitListTableInsert = 1
				SET @pfWaitListTableDelete = 1
				SET @pfWaitListCourseTitleColumnUpdate = 1
				SET @pfWaitListCourseTitleColumnSelect = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piWaitListTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sWLTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (195, 196)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iTempAction = 195
						BEGIN
							SET @pfWaitListTableInsert = 1
						END
						IF @iTempAction = 196
						BEGIN
							SET @pfWaitListTableDelete = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction

					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @waitListColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @waitListColumnPermissions
					SELECT 
						ASRSysColumns.columnId,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piWaitListTableID
						AND ASRSysColumns.columnId = @piWaitListCourseTitleColumnID)
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfWaitListCourseTitleColumnUpdate = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 197
					IF @pfWaitListCourseTitleColumnUpdate IS NULL SET @pfWaitListCourseTitleColumnUpdate = 0

					SELECT @pfWaitListCourseTitleColumnSelect = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 193
					IF @pfWaitListCourseTitleColumnSelect IS NULL SET @pfWaitListCourseTitleColumnSelect = 0

				END
			END
		END

		/* Get the Bulk Booking default view. */
		SELECT @piBulkBookingDefaultViewID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_BulkBookingDefaultView'
		IF @piBulkBookingDefaultViewID IS NULL SET @piBulkBookingDefaultViewID = 0
	END
END
GO



/****** Object:  StoredProcedure [dbo].[spASRIntAllTablePermissions]    Script Date: 02/01/2014 20:52:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntAllTablePermissions]
(
	@psSQLLogin 		varchar(255)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID		integer,
		@sUserGroupName				sysname,
		@sActualUserName			sysname;

	-- Cached view of the objects 
	DECLARE @SysObjects TABLE([ID]		integer PRIMARY KEY CLUSTERED,
							  [Name]	sysname);
		
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
							  
	INSERT INTO @SysObjects
		SELECT [ID], [Name] FROM sysobjects
		WHERE [Name] LIKE 'ASRSysCV%' AND [XType] = 'v'
		UNION 
		SELECT OBJECT_ID(tableName), TableName 
		FROM ASRSysTables
		WHERE NOT OBJECT_ID(tableName) IS null
		UNION
		SELECT OBJECT_ID(viewName), ViewName 
		FROM ASRSysViews
		WHERE NOT OBJECT_ID(viewName) IS null;

	-- Cached view of the sysprotects table
	DECLARE @SysProtects TABLE([ID]				integer,
							   [columns]		varbinary(8000),
							   [Action]			tinyint,
							   [ProtectType]	tinyint);
	INSERT INTO @SysProtects
	SELECT p.ID, p.Columns, p.Action, p.ProtectType FROM ASRSysProtectsCache p
		INNER JOIN @SysObjects o ON p.ID = o.ID
		WHERE p.UID = @iUserGroupID AND ((p.ProtectType <> 206 AND p.Action <> 193) OR (p.Action = 193 AND p.ProtectType IN (204,205)));

	SELECT UPPER(o.name) AS [name], p.action, ISNULL(cv.tableID,0) AS [tableid]
		FROM @SysProtects p
		INNER JOIN @SysObjects o ON p.id = o.id
		LEFT JOIN ASRSysChildViews2 cv ON cv.childViewID = CASE SUBSTRING(o.Name, 1, 8) WHEN 'ASRSysCV' THEN SUBSTRING(o.Name, 9, CHARINDEX('#',o.Name, 0) - 9) ELSE 0 END
		WHERE p.protectType <> 206
			AND p.action <> 193
	UNION
	SELECT UPPER(o.name) AS [name], 193, ISNULL(cv.tableID,0) AS [tableid]
		FROM sys.columns c
		INNER JOIN @SysProtects p ON (c.object_id = p.id
			AND p.action = 193 
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,c.column_id/8+1,1))&power(2,c.column_id&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,c.column_id/8+1,1))&power(2,c.column_id&7)) = 0)))
		INNER JOIN @SysObjects o ON p.id = o.id
		LEFT JOIN ASRSysChildViews2 cv ON cv.childViewID = CASE SUBSTRING(o.Name, 1, 8) WHEN 'ASRSysCV' THEN SUBSTRING(o.Name, 9, CHARINDEX('#',o.Name, 0) - 9) ELSE 0 END
		WHERE (c.name <> 'timestamp' AND c.name <> 'ID')
			AND p.protectType IN (204, 205) 
		ORDER BY name;

END

GO

/****** Object:  StoredProcedure [dbo].[spASRIntGet1000SeparatorFindColumns]    Script Date: 02/01/2014 20:52:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGet1000SeparatorFindColumns] (
	@pfError 				bit 			OUTPUT, 
	@piTableID 				integer, 
	@piViewID 				integer, 
	@piOrderID 				integer, 
	@ps1000SeparatorCols	varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@sType	 			varchar(10),
		@fSelectGranted 	bit,
		@iCount				integer,
		@bUse1000Separator	bit,
		@sActualLoginName	varchar(250);

	/* Initialise variables. */
	SET @pfError = 0;
	SET @ps1000SeparatorCols = '';
	SET @sRealSource = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID;

	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1;
		RETURN;
	END

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @piViewID;
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piTableID
			AND [role] = @sUserGroupName;
		
		IF @iChildViewID IS null SET @iChildViewID = 0;
		
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END

	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	-- Cached view of sysprotects
	DECLARE @SysProtects TABLE([ID] int, [ProtectType] tinyint, [Columns] varbinary(8000))
	INSERT INTO @SysProtects
		SELECT ID, ProtectType, [Columns] FROM ASRSysProtectsCache
		WHERE [UID] = @iUserGroupID AND Action = 193;

	/* Create a temporary table of the 'select' column permissions for all tables/views used in the order. */
	DECLARE @ColumnPermissions TABLE(
				tableID			integer,
				tableViewName	sysname,
				columnName		sysname,
				selectGranted	bit);

	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	WHERE ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				CASE p.protectType
				    WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @Sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.tableID,
		ASRSysColumns.columnName,
		ASRSysTables.tableName,
		ASRSysOrderItems.type,
		ASRSysColumns.Use1000Separator
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
		AND ASRSysOrderItems.type = 'F'
	ORDER BY ASRSysOrderItems.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @sColumnName, @sColumnTableName, @sType, @bUse1000Separator;

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
					CASE
						WHEN @bUse1000Separator = 1 THEN '1'
						ELSE '0'
					END;
			END
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
	
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
					CASE
						WHEN @bUse1000Separator = 1 THEN '1'
						ELSE '0'
					END;
			END
			ELSE	
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM @ColumnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1;

				IF @iCount > 0 
				BEGIN
					SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
						CASE
							WHEN @bUse1000Separator = 1 THEN '1'
							ELSE '0'
						END;
				END
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @sColumnName, @sColumnTableName, @sType, @bUse1000Separator;
	END

	CLOSE orderCursor;
	DEALLOCATE orderCursor;

END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetChartData]    Script Date: 02/01/2014 20:52:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetChartData] (
	--@piHistoryTableID	integer,
	@piParentTableID 	integer,
--	@piParentRecordID	integer,
	@piColumnId integer,
	@piAggregateType integer
)
AS
BEGIN
	SET NOCOUNT ON;
	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(255),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(255),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sRootTable 		varchar(255),
		@sSelectString 		varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	varchar(255),
		@sTemp				varchar(MAX),
		@sSelectSQL			nvarchar(MAX),
		@sActualUserName	sysname,
		@strTempSepText		varchar(500);
	SET @sSelectSQL = '';
		
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
		INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
		OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS';
	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID;
	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(tableViewName	sysname);
	/* Create a temporary table of the 'read' column permissions for all tables/views used. */
	DECLARE @columnPermissions TABLE (tableViewName	sysname,
				columnName	sysname,
				granted		bit);
	-- Cached view of SysProtects
	DECLARE @SysProtects TABLE([ID] int, [ProtectType] tinyint, [Columns] varbinary(8000));
	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO @ColumnPermissions
		SELECT 
			@sParentTableName,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns 
		WHERE ASRSysColumns.tableID = @piParentTableID;
	END
	ELSE
	BEGIN
		IF @iParentTableType <> 2 /* ie. top-level or lookup */
		BEGIN
		
			-- Get list of views/table columns that are summary fields
			DECLARE @SummaryColumns TABLE ([ID] int, [TableName] sysname, [ColumnName] sysname, [ColID] int)
			INSERT @SummaryColumns
				SELECT sysobjects.id, sysobjects.name,
					syscolumns.name, syscolumns.ColID
				FROM sysobjects
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.name IN (SELECT ASRSysTables.tableName
												FROM ASRSysTables
												WHERE ASRSysTables.tableID = @piParentTableID
											UNION SELECT ASRSysViews.viewName
												FROM ASRSysViews
												WHERE ASRSysViews.viewTableID = @piParentTableID)
					AND syscolumns.name IN (SELECT ac.ColumnName
												FROM ASRSysColumns ac
												WHERE ac.TableID = @piParentTableID);
			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM SysProtects
				WHERE Action = 193;
			-- Generate security context on selected columns
			INSERT INTO @ColumnPermissions
				SELECT sm.TableName,
					sm.ColumnName,
					CASE p.protectType
						WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @SysProtects p
				INNER JOIN @SummaryColumns sm ON p.id = sm.id
				WHERE (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) = 0));
		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @piParentTableID
				AND role = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sParentRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sParentTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sParentRealSource = left(@sParentRealSource, 255);
			END
			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM ASRSysProtectsCache
				WHERE [UID] = @iUserGroupID AND Action = 193;
			INSERT INTO @ColumnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	END
	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	
	/*populate the temp table with info from ssi - i.e. chart column details*/
	
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	--SELECT ASRSysColumns.columnId, 
	--	ASRSysColumns.columnName, 
	--	ASRSysColumns.dataType
	--FROM ASRSysSummaryFields 
	--INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnId
	--WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
	--	AND ASRSysColumns.tableID = @piParentTableID 
	--ORDER BY ASRSysSummaryFields.sequence;
	SELECT ASRSysColumns.columnId, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysColumns
	WHERE ASRSysColumns.ColumnID = @piColumnID;
		
	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	WHILE (@@fetch_status = 0)
	BEGIN
		
		SET @fSelectGranted = 0;
		/* Get the select permission on the column. */
		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM @ColumnPermissions
		WHERE tableViewName = @sParentTableName
			AND columnName = @sColumnName;
		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			SET @sTemp = ',';
			IF LEN(@sSelectSQL) > 0
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sTemp = 'convert(varchar(10), ' + @sParentTableName + '.' + @sColumnName + ', 101) AS [COLUMN]';
				IF @piAggregateType = 0	/* Aggregate = Count */
				BEGIN
					SET @sTemp = @sTemp + ', count(convert(varchar(10), ' + @sParentTableName + '.' + @sColumnName + ', 101)) as [Aggregate]'
				END
				ELSE	/* Aggregate = Total */
				BEGIN
					SET @sTemp = @sTemp + ', sum(convert(varchar(10), ' + @sParentTableName + '.' + @sColumnName + ', 101)) as [Aggregate]'
				END
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END 
			ELSE
			BEGIN
				 /* Non-date */
				SET @sTemp = @sParentTableName + '.' + @sColumnName + ' AS [COLUMN]';
				IF @piAggregateType = 0	/* Aggregate = Count */
				BEGIN
					SET @sTemp = @sTemp + ', count(' + @sParentTableName + '.' + @sColumnName + ') as [Aggregate]'
				END
				ELSE	/* Aggregate = Total */
				BEGIN
					SET @sTemp = @sTemp + ', sum(' + @sParentTableName + '.' + @sColumnName + ') as [Aggregate]'
				END
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END
			/* Add the table to the array of tables/views to join if it has not already been added. */
			SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
				WHERE tableViewName = @sParentTableName;
			IF @iTempCount = 0
			BEGIN
				INSERT INTO @joinParents (tableViewName) VALUES(@sParentTableName);
			END
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = '';
			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @ColumnPermissions
			WHERE tableViewName <> @sParentTableName
				AND columnName = @sColumnName
				AND granted = 1;
			OPEN viewCursor;
			FETCH NEXT FROM viewCursor INTO @sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Column CAN be read from the view. */
				SET @fSelectGranted = 1;
				IF len(@sSelectString) > 0 SET @sSelectString = @sSelectString + ',';
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					SET @sSelectString = @sSelectString + 'convert(varchar(10),' + @sViewName + '.' + @sColumnName + ',101)';
					IF @piAggregateType = 0	/* Aggregate = Count */
					BEGIN
						SET @sSelectString = @sSelectString + ', count(convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)) as [Aggregate]'
					END
					ELSE	/* Aggregate = Total */
					BEGIN
						SET @sSelectString = @sSelectString + ', sum(convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)) as [Aggregate]'
					END
				END
				ELSE
				BEGIN
					SET @sSelectString = @sSelectString + @sViewName + '.' + @sColumnName;
					IF @piAggregateType = 0	/* Aggregate = Count */
					BEGIN
						SET @sSelectString = @sSelectString + ', count(' + @sViewName + '.' + @sColumnName + ') as [Aggregate]'
					END
					ELSE	/* Aggregate = Total */
					BEGIN
						SET @sSelectString = @sSelectString + ', sum(' + @sViewName + '.' + @sColumnName + ') as [Aggregate]'
					END
				END
				/* Add the view to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewName = @sViewName;
				IF @iTempCount = 0
					INSERT INTO @joinParents (tableViewName) VALUES(@sViewName);
				FETCH NEXT FROM viewCursor INTO @sViewName;
			END
			CLOSE viewCursor;
			DEALLOCATE viewCursor;
			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = 'COALESCE(' + @sSelectString + ', NULL) AS [COLUMN]';
				SET @sTemp = ',';
				IF LEN(@sSelectSQL) > 0
					SET @sSelectSQL = @sSelectSQL + @sTemp;
				SET @sTemp = @sSelectString;
				SET @sSelectSQL = @sSelectSQL + @sTemp;
				
			END
		END
		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;
	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = 'SELECT ' + @sSelectSQL ;
		SELECT @iTempCount = COUNT(tableViewName)
			FROM @joinParents;
		IF @iTempCount = 1 
		BEGIN
			SELECT TOP 1 @sRootTable = tableViewName
			FROM @joinParents;
		END
		ELSE
		BEGIN
			SET @sRootTable = @sParentTableName;
		END
		--SET @sTemp = ', row_number()  over ( order by ' +@sViewName + '.' + @sColumnName + ') as ROW_NUMBER FROM ' + @sRootTable;
		SET @sTemp = ' FROM ' + @sRootTable;
		SET @sSelectSQL = @sSelectSQL + @sTemp;
		/* Add the join code. */
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @joinParents;
		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sTableViewName <> @sRootTable
			BEGIN
				SET @sTemp = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRootTable + '.ID' + ' = ' + @sTableViewName + '.ID';
				SET @sSelectSQL = @sSelectSQL + @sTemp
			END
			FETCH NEXT FROM joinCursor INTO @sTableViewName;
		END
		CLOSE joinCursor;
		DEALLOCATE joinCursor;
--		SET @sTemp = ' WHERE ' + @sRootTable + '.id = ' + convert(varchar(255), @piParentRecordID);
--		SET @sSelectSQL = @sSelectSQL + @sTemp;
		
		SET @sTemp = ' GROUP BY ' + @sRootTable + '.' + @sColumnName;
		SET @sSelectSQL = @sSelectSQL + @sTemp;
		
	END
	-- Run the constructed SQL SELECT string.
	EXEC sp_executeSQL @sSelectSQL;
END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetColumnsFromTablesAndViews]    Script Date: 02/01/2014 20:52:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetColumnsFromTablesAndViews]
AS
BEGIN

	SET NOCOUNT ON;

	SELECT UPPER(c.columnName) AS [ColumnName], c.columnType, c.dataType
		, c.columnID, ISNULL(c.uniqueCheckType,0) AS uniqueCheckType
		, UPPER(t.tableName) AS tableViewName
	FROM dbo.ASRSysColumns c
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID
	UNION 
	SELECT UPPER(c.columnName) AS [ColumnName], c.columnType, c.dataType
		, c.columnID, ISNULL(c.uniqueCheckType,0) AS uniqueCheckType
		, UPPER(v.viewName) AS tableViewName 
	FROM dbo.ASRSysColumns c
	INNER JOIN ASRSysViews v ON c.tableID = v.viewTableID 
	LEFT OUTER JOIN ASRSysViewColumns vc ON (v.viewID = vc.viewID 
			AND c.columnID = vc.columnID)
	WHERE vc.inView = 1 OR c.columnType = 3 
	ORDER BY tableViewName;

END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetLinks]    Script Date: 02/01/2014 20:52:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetLinks] 
(
		@plngTableID	integer,
		@plngViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iCount				integer,
		@iUtilType			integer, 
		@iUtilID			integer,
		@iScreenID			integer,
		@iTableID			integer,
		@sTableName			sysname,
		@iTableType			integer,
		@sRealSource		sysname,
		@iChildViewID		integer,
		@sAccess			varchar(MAX),
		@sGroupName			varchar(255),
		@pfPermitted		bit,
		@sActualUserName	sysname,
		@iActualUserGroupID integer,
		@fBaseTableReadable bit,
		@iBaseTableID		integer,
		@sURL				varchar(MAX), 
		@fUtilOK			bit,
		@fDrillDownHidden bit,
		@iLinkType			integer,		-- 0 = Hypertext, 1 = Button, 2 = Dropdown List
		@iElement_Type		integer;		-- 2 = chart

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
	
	IF @plngViewID < 1 
	BEGIN 
		SET @plngViewID = -1;
	END
	SET @fBaseTableReadable = 1;
	
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA'
	BEGIN
		EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sGroupName OUTPUT,
			@iActualUserGroupID OUTPUT;
		
		DECLARE @Phase1 TABLE([ID] INT);
		INSERT INTO @Phase1
			SELECT Object_ID(ASRSysViews.ViewName) 
			FROM ASRSysViews 
			WHERE NOT Object_ID(ASRSysViews.ViewName) IS null
			UNION
			SELECT Object_ID(ASRSysTables.TableName) 
			FROM ASRSysTables 
			WHERE NOT Object_ID(ASRSysTables.TableName) IS null
			UNION
			SELECT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255))
			FROM ASRSysChildViews2
			INNER JOIN ASRSysTables 
				ON ASRSysChildViews2.tableID = ASRSysTables.tableID
			WHERE NOT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255)) IS null;
		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int PRIMARY KEY CLUSTERED);
		INSERT INTO @SysProtects
			SELECT p.[ID] 
			FROM ASRSysProtectsCache p
						INNER JOIN SysColumns c ON (c.id = p.id
							AND c.[Name] = 'timestamp'
							AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
							OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0)))
			WHERE p.UID = @iUserGroupID
				AND p.[ProtectType] IN (204, 205)
				AND p.[Action] = 193			
				AND p.id IN (SELECT ID FROM @Phase1);
		-- Readable tables
		DECLARE @ReadableTables TABLE([Name] sysname PRIMARY KEY CLUSTERED);
		INSERT INTO @ReadableTables
			SELECT OBJECT_NAME(p.ID)
			FROM @SysProtects p;
		
		SET @sRealSource = '';
		IF @plngViewID > 0
		BEGIN
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @plngViewID;
		END
		ELSE
		BEGIN
			SELECT @sRealSource = tableName
			FROM ASRSysTables
			WHERE tableID = @plngTableID;
		END
		SET @fBaseTableReadable = 0
		IF len(@sRealSource) > 0
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM @ReadableTables
			WHERE name = @sRealSource;
		
			IF @iCount > 0
			BEGIN
				SET @fBaseTableReadable = 1;
			END
		END
	END
	DECLARE @Links TABLE([ID]						integer PRIMARY KEY CLUSTERED,
											 [utilityType]	integer,
											 [utilityID]		integer,
											 [screenID]			integer,
											 [LinkType]			integer,
											 [Element_Type]	integer,
											 [DrillDownHidden]				bit);
	INSERT INTO @Links ([ID],[utilityType],[utilityID],[screenID], [LinkType], [Element_Type], [DrillDownHidden])
	SELECT ASRSysSSIntranetLinks.ID,
					ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType,
					ASRSysSSIntranetLinks.Element_Type,
					0
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
			AND tableid = @plngTableID)
			AND (id NOT IN (SELECT linkid 
								FROM ASRSysSSIHiddenGroups
								WHERE groupName = @sGroupName));
	/* Remove any utility links from the temp table where the utility has been deleted or hidden from the current user.*/
	/* Or if the user does not permission to run them. */	
	DECLARE utilitiesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType,
					ASRSysSSIntranetLinks.Element_Type
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
				AND tableid = @plngTableID)
			AND (utilityID > 0 
				OR screenID > 0);
	OPEN utilitiesCursor;
	FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType, @iElement_Type;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iUtilID > 0
		BEGIN
			SET @fUtilOK = 1	;			
			/* Check if the utility is deleted or hidden from the user. */
			EXECUTE [dbo].[spASRIntCurrentAccessForRole]
								@sGroupName,
								@iUtilType,
								@iUtilID,
								@sAccess	OUTPUT;
			IF @sAccess = 'HD' 
			BEGIN
				/* Report/utility is hidden from the user. */
				--HERE FOR CHARTs **************************************************************************************************************************************
				IF @iElement_Type = 2
				BEGIN
					SET @fUtilOK = 1;				
					SET @fDrillDownHidden = 1;
				END
				ELSE
				BEGIN
					SET @fUtilOK = 0;
				END
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has system permission to run this type of report/utility. */
				IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA'
				BEGIN
					SELECT @pfPermitted = ASRSysGroupPermissions.permitted
					FROM ASRSysPermissionItems
					INNER JOIN ASRSysPermissionCategories 
					ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 					
							CASE 
								WHEN @iUtilType = 17 THEN 'CALENDARREPORTS'
								WHEN @iUtilType = 9 THEN 'MAILMERGE'
								WHEN @iUtilType = 2 THEN 'CUSTOMREPORTS'
								WHEN @iUtilType = 25 THEN 'WORKFLOW'
								ELSE ''
							END
					LEFT OUTER JOIN ASRSysGroupPermissions 
					ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
						AND ASRSysGroupPermissions.groupName = @sGroupName
					WHERE ASRSysPermissionItems.itemKey = 'RUN';
					IF (@pfPermitted IS null) OR (@pfPermitted = 0)
					BEGIN
						/* User does not have system permission to run this type of report/utility. */
						--HERE FOR CHARTS**************************************************************************************************************************************
						IF @iElement_Type = 2
						BEGIN
							SET @fUtilOK = 1;
							SET @fDrillDownHidden = 1;
						END
						ELSE
						BEGIN
							SET @fUtilOK = 0;
						END
					END
				END
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has read permission on the report/utility base table or any views on it. */
				SET @iBaseTableID = 0;
				IF @iUtilType = 17 /* Calendar Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCalendarReports
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 2 /* Custom Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCustomReportsName
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 9 /* Mail Merge */
				BEGIN
					SELECT @iBaseTableID = TableID
					FROM ASRSysMailMergeName
					WHERE MailMergeID = @iUtilID;
				END
				/* Not check required for reports/utilities without a base table.
				OR reports/utilities based on the top-level table if the user has read permission on the current view. */
				IF (@iBaseTableID > 0)
					AND((@fBaseTableReadable = 0)
						OR (@iBaseTableID <> @plngTableID))
				BEGIN
					IF (@iLinkType <> 0) -- Hypertext link
						AND (@fBaseTableReadable = 0)
						AND (@iBaseTableID = @plngTableID)
					BEGIN
						/* The report/utility is based on the top-level table, and the user does NOT have read permission
						on the current view (on which Button & DropdownList links are scoped). */
						SET @fUtilOK = 0;
					END
					ELSE
					BEGIN
						SELECT @iCount = COUNT(p.ID)
						FROM @SysProtects p
						WHERE OBJECT_NAME(p.ID) IN (SELECT ASRSysTables.tableName
							FROM ASRSysTables
							WHERE ASRSysTables.tableID = @iBaseTableID
						UNION 
							SELECT ASRSysViews.viewName
								FROM ASRSysViews
								WHERE ASRSysViews.viewTableID = @iBaseTableID
						UNION
							SELECT
								left('ASRSysCV' 
									+ convert(varchar(1000), ASRSysChildViews2.childViewID) 
									+ '#'
									+ replace(ASRSysTables.tableName, ' ', '_')
									+ '#'
									+ replace(@sGroupName, ' ', '_'), 255)
								FROM ASRSysChildViews2
								INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID
								WHERE ASRSysChildViews2.role = @sGroupName
									AND ASRSysChildViews2.tableID = @iBaseTableID);
						IF @iCount = 0 
						BEGIN
							SET @fUtilOK = 0;
						END
					END
				END
			END
			/* For some reason the user cannot use this report/utility, so remove it from the temp table of links. */
			IF @fUtilOK = 0 
			BEGIN
				DELETE FROM @Links
				WHERE utilityType = @iUtilType
					AND utilityID = @iUtilID;
			END
			IF @fDrillDownHidden = 1
			BEGIN
				UPDATE @Links
				SET DrillDownHidden = 1 
				WHERE utilityType = @iUtilType
					AND utilityID = @iUtilID;
			END
			
		END
		
		IF (@iScreenID > 0) AND (UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA')
		BEGIN
			/* Do not display the link if the user does not have permission to read the defined view/table for the screen. */
			SELECT @iTableID = ASRSysTables.tableID, 
				@sTableName = ASRSysTables.tableName,
				@iTableType = ASRSysTables.tableType
			FROM ASRSysScreens
						INNER JOIN ASRSysTables 
						ON ASRSysScreens.tableID = ASRSysTables.tableID
			WHERE screenID = @iScreenID;
			SET @sRealSource = '';
			IF @iTableType  = 2
			BEGIN
				SET @iChildViewID = 0;
				/* Child table - check child views. */
				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTableID
					AND [role] = @sGroupName;
				
				IF @iChildViewID IS null SET @iChildViewID = 0;
				
				IF (@iChildViewID > 0) AND (@fBaseTableReadable = 1)
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTableName, ' ', '_') +
						'#' + replace(@sGroupName, ' ', '_');
				
					SET @sRealSource = left(@sRealSource, 255);
				END
				ELSE
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
			ELSE
			BEGIN
				/* Not a child table - must be the top-level table. Check if the user has 'read' permission on the defined view. */
				SET @sRealSource = '';
				IF @plngViewID > 0
				BEGIN
					SELECT @sRealSource = viewName
					FROM ASRSysViews
					WHERE viewID = @plngViewID;
				END
				ELSE
				BEGIN
					SELECT @sRealSource = tableName
					FROM ASRSysTables
					WHERE tableID = @plngTableID;
				END
			END
			IF len(@sRealSource) > 0
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM @ReadableTables
				WHERE name = @sRealSource;
			
				IF @iCount = 0
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
		END
		FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType, @iElement_Type;
	END
	CLOSE utilitiesCursor;
	DEALLOCATE utilitiesCursor;
	/* Remove the Workflow links if the URL has not been configured. */
	SELECT @sURL = isnull(settingValue , '')
	FROM ASRSysSystemSettings
	WHERE section = 'MODULE_WORKFLOW'		
		AND settingKey = 'Param_URL';	
	
	
	IF LEN(@sURL) = 0
	BEGIN
		DELETE FROM @Links
		WHERE utilityType = 25;
	END
	SELECT ASRSysSSIntranetLinks.*, 
		CASE 
			WHEN ASRSysSSIntranetLinks.utilityType = 9 THEN ASRSysMailMergeName.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 2 THEN ASRSysCustomReportsName.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 17 THEN ASRSysCalendarReports.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 25 THEN 0
			ELSE null
		END AS [baseTable],
		ASRSysColumns.ColumnName as [Chart_ColumnName],
		tvL.DrillDownHidden as [DrillDownHidden]
	FROM ASRSysSSIntranetLinks
			LEFT OUTER JOIN ASRSysMailMergeName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysMailMergeName.MailMergeID
				AND ASRSysSSIntranetLinks.utilityType = 9
			LEFT OUTER JOIN ASRSysCalendarReports 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCalendarReports.ID
				AND ASRSysSSIntranetLinks.utilityType = 17
			LEFT OUTER JOIN ASRSysCustomReportsName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCustomReportsName.ID
				AND ASRSysSSIntranetLinks.utilityType = 2
			LEFT OUTER JOIN ASRSysColumns
			ON ASRSysSSIntranetLinks.Chart_ColumnID = ASRSysColumns.columnId		
			LEFT OUTER JOIN @Links tvL
			ON ASRSysSSIntranetLinks.ID = tvL.ID
	WHERE ASRSysSSIntranetLinks.ID IN (SELECT ID FROM @Links)
	ORDER BY ASRSysSSIntranetLinks.linkOrder;
	
END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetNavigationLinks]    Script Date: 02/01/2014 20:52:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetNavigationLinks]
(
		@plngTableID	integer,
		@plngViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE
		@iCount					integer,
		@iViewID				integer,
		@iUtilType				integer, 
		@iUtilID				integer, 
		@iScreenID				integer, 
		@sURL					varchar(MAX),
		@iTableID				integer,
		@sTableName				sysname,
		@iTableType				integer,
		@sRealSource			sysname,
		@iChildViewID			integer,
		@sAccess				varchar(MAX),
		@fTableViewOK			bit,
		@pfCustomReportsRun		bit,
		@pfCalendarReportsRun	bit,
		@pfMailMergeRun			bit,
		@pfWorkflowRun			bit,
		@sGroupName				varchar(255),
		@sActualUserName		sysname,
		@iActualUserGroupID 	integer, 
		@sViewName				sysname,
		@iLinkType 				integer,			/* 0 = Hypertext, 1 = Button, 2 = Dropdown List */
		@fFindPage				bit

	/* See if the current user can run the defined Reports/Utilties. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER)))  = 'SA'
	BEGIN
		SET @pfCustomReportsRun = 1
		SET @pfCalendarReportsRun = 1
		SET @pfMailMergeRun = 1
		SET @pfWorkflowRun = 1
	END
	ELSE
	BEGIN

		EXEC dbo.spASRIntGetActualUserDetails
			@sActualUserName OUTPUT,
			@sGroupName OUTPUT,
			@iActualUserGroupID OUTPUT
			
		DECLARE @unionTable TABLE (ID int PRIMARY KEY CLUSTERED)

		INSERT INTO @unionTable 
			SELECT Object_ID(ViewName) 
			FROM ASRSysViews 
			WHERE viewID IN (SELECT viewID FROM ASRSysSSIViews)
				AND NOT Object_ID(ViewName) IS null
			UNION
			SELECT Object_ID(TableName) 
			FROM ASRSysTables 
			WHERE tableID IN (SELECT tableID FROM ASRSysSSIViews)
				AND NOT Object_ID(TableName) IS null
				AND tableID NOT IN (SELECT tableID 
					FROM ASRSysViewMenuPermissions 
					WHERE ASRSysViewMenuPermissions.groupName = @sGroupName
						AND ASRSysViewMenuPermissions.hideFromMenu = 1)
			UNION
			SELECT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255))
			FROM ASRSysChildViews2
			INNER JOIN ASRSysTables 
				ON ASRSysChildViews2.tableID = ASRSysTables.tableID
			WHERE NOT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255)) IS null

		DECLARE @readableTables TABLE (name sysname)	
	
		INSERT INTO @readableTables
			SELECT OBJECT_NAME(p.id)
			FROM syscolumns
			INNER JOIN ASRSysProtectsCache p 
				ON (syscolumns.id = p.id
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
			WHERE p.UID = @iActualUserGroupID
				AND syscolumns.name = 'timestamp'
				AND (p.ID IN (SELECT id FROM @unionTable))
				AND p.Action = 193 AND ProtectType IN (204, 205)
				OPTION (KEEPFIXED PLAN)

		SELECT @pfCustomReportsRun = ASRSysGroupPermissions.permitted
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories 
			ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				AND ASRSysPermissionCategories.categoryKey = 	'CUSTOMREPORTS'
		LEFT OUTER JOIN ASRSysGroupPermissions 
			ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
				AND ASRSysGroupPermissions.groupName = @sGroupName
		WHERE ASRSysPermissionItems.itemKey = 'RUN'
	
		SELECT @pfCalendarReportsRun = ASRSysGroupPermissions.permitted
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories 
			ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				AND ASRSysPermissionCategories.categoryKey = 	'CALENDARREPORTS'
		LEFT OUTER JOIN ASRSysGroupPermissions 
			ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
				AND ASRSysGroupPermissions.groupName = @sGroupName
		WHERE ASRSysPermissionItems.itemKey = 'RUN'

		SELECT @pfMailMergeRun = ASRSysGroupPermissions.permitted
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories 
			ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				AND ASRSysPermissionCategories.categoryKey = 	'MAILMERGE'
		LEFT OUTER JOIN ASRSysGroupPermissions 
			ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
				AND ASRSysGroupPermissions.groupName = @sGroupName
		WHERE ASRSysPermissionItems.itemKey = 'RUN'

		SELECT @pfWorkflowRun = ASRSysGroupPermissions.permitted
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories 
			ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				AND ASRSysPermissionCategories.categoryKey = 	'WORKFLOW'
		LEFT OUTER JOIN ASRSysGroupPermissions 
			ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
				AND ASRSysGroupPermissions.groupName = @sGroupName
		WHERE ASRSysPermissionItems.itemKey = 'RUN'
	END

	DECLARE @links TABLE(
		LinkType			integer,
		Text1	 			varchar(200),
		Text2	 			varchar(200),
		SingleRecord		bit,
		LinkToFind			bit,
		TableID				integer,
		ViewID				integer ,
		PrimarySequence		integer,
		SecondarySequence	integer,
		FindPage			integer)

	/* Hypertext links. */
	/* Single Record View UNION Multiple Record Tables/Views UNION Table/View Hypertext Links link */
	INSERT INTO @links
		SELECT 0, linksLinkText, '', 1, 0, tableID, viewID, 0, sequence, 0
		FROM ASRSysSSIViews
		WHERE singleRecordView = 1
			AND LEN(linksLinkText) > 0
		UNION
		SELECT 0, hypertextLinkText, '', 0, 1, tableID, viewID, 2, sequence, 0
		FROM ASRSysSSIViews
		WHERE singleRecordView = 0
			AND LEN(hypertextLinkText) > 0
		UNION
		SELECT 0, linksLinkText, '', 0, 0, tableID, viewID, 1, sequence, 1
		FROM ASRSysSSIViews
		WHERE singleRecordView = 0
			AND tableid = @plngTableID
			AND viewID = @plngViewID

	/* Button links. */
	INSERT INTO @links
	SELECT 1, buttonLinkPromptText, buttonLinkButtonText, 0, 1, tableID, viewID, 0, sequence, 0
	FROM ASRSysSSIViews
	WHERE buttonLink = 1

	/* DropdownList links. */
	INSERT INTO @links
	SELECT 2, dropdownListLinkText, '', 0, 1, tableID, viewID, 0, sequence, 0
	FROM ASRSysSSIViews
	WHERE dropdownListLink = 1


	/* Remove linkToFind links for links to views that are not readable by the user, or those that have no valid links defined for them. */
	DECLARE viewsCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT ISNULL(l.viewID, -1) 'viewID', ASRSysViews.viewName, l.tableID, ASRSysTables.tableName
		FROM @links	l
		LEFT OUTER JOIN ASRSysViews	
			ON l.viewID = ASRSysViews.viewID
		INNER JOIN ASRSysTables
			ON l.tableID = ASRSysTables.tableID

	OPEN viewsCursor
	FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName, @iTableID, @sTableName
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fTableViewOK = 0
		
		IF @iViewID > 0 
		BEGIN 
			SELECT @iCount = COUNT(*)
			FROM @readableTables
			WHERE name = @sViewName
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM @readableTables
			WHERE name = @sTableName
		END 

		IF @iCount > 0
		BEGIN

			DECLARE linksCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysSSIntranetLinks.utilityType,
							ASRSysSSIntranetLinks.utilityID,
							ASRSysSSIntranetLinks.screenID,
							ASRSysSSIntranetLinks.url
			FROM ASRSysSSIntranetLinks
			WHERE tableID = @iTableID 
				AND	viewID = @iViewID
	
			OPEN linksCursor
			FETCH NEXT FROM linksCursor INTO @iUtilType, @iUtilID, @iScreenID, @sURL
			WHILE (@@fetch_status = 0) AND (@fTableViewOK = 0)
			BEGIN
				IF LEN(@sURL) > 0 OR (UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA')
				BEGIN
					SET @fTableViewOK = 1
				END
				ELSE
				BEGIN
					IF @iUtilID > 0
					BEGIN
						/* Check if the utility is deleted or hidden from the user. */
						EXECUTE dbo.spASRIntCurrentAccessForRole
												@sGroupName,
												@iUtilType,
												@iUtilID,
												@sAccess	OUTPUT
	
						IF @sAccess <> 'HD' 
						BEGIN
							IF @iUtilType = 2 AND @pfCustomReportsRun = 1 SET @fTableViewOK = 1
							IF @iUtilType = 17 AND @pfCalendarReportsRun = 1 SET @fTableViewOK = 1
							IF @iUtilType = 9 AND @pfMailMergeRun = 1 SET @fTableViewOK = 1
							IF @iUtilType = 25 AND @pfWorkflowRun = 1 SET @fTableViewOK = 1
						END
					END
	
					IF (@iScreenID > 0) 
					BEGIN
						/* Do not display the link if the user does not have permission to read the defined view/tbale for the screen. */
						SELECT @iTableID = ASRSysTables.tableID, 
							@sTableName = ASRSysTables.tableName,
							@iTableType = ASRSysTables.tableType
						FROM ASRSysScreens
										INNER JOIN ASRSysTables 
										ON ASRSysScreens.tableID = ASRSysTables.tableID
						WHERE screenID = @iScreenID
	
						SET @sRealSource = ''
						IF @iTableType  = 2
						BEGIN
							SET @iChildViewID = 0
	
							/* Child table - check child views. */
							SELECT @iChildViewID = childViewID
							FROM ASRSysChildViews2
							WHERE tableID = @iTableID
								AND role = @sGroupName
							
							IF @iChildViewID IS null SET @iChildViewID = 0
							
							IF @iChildViewID > 0 
							BEGIN
								SET @sRealSource = 'ASRSysCV' + 
									convert(varchar(1000), @iChildViewID) +
									'#' + replace(@sTableName, ' ', '_') +
									'#' + replace(@sGroupName, ' ', '_')
							
								SET @sRealSource = left(@sRealSource, 255)
							END
						END
						ELSE
						BEGIN
							/* Not a child table - must be the top-level table. Check if the user has 'read' permission on the defined view. */
							IF @iViewID > 0 
							BEGIN 
								SELECT @sRealSource = viewName
								FROM ASRSysViews
								WHERE viewID = @iViewID
							END
							ELSE
							BEGIN
								SELECT @sRealSource = tableName
								FROM ASRSysTables
								WHERE tableID = @iTableID
							END 
	
							IF @sRealSource IS null SET @sRealSource = ''
						END
	
						IF len(@sRealSource) > 0
						BEGIN
							SELECT @iCount = COUNT(*)
							FROM @readableTables
							WHERE name = @sRealSource
						
							IF @iCount = 1 SET @fTableViewOK = 1
						END
					END
				END
								
				FETCH NEXT FROM linksCursor INTO @iUtilType, @iUtilID, @iScreenID, @sURL
			END
			CLOSE linksCursor
			DEALLOCATE linksCursor

		END
		
		IF @fTableViewOK = 0
		BEGIN
			IF @iViewID > 0 
			BEGIN
				DELETE FROM @links
				WHERE viewID = @iViewID
			END
			ELSE
			BEGIN
				DELETE FROM @links
				WHERE tableid = @iTableID AND viewID = @iViewID
			END
		END
	
		FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName, @iTableID, @sTableName
	END
	CLOSE viewsCursor
	DEALLOCATE viewsCursor

	SELECT *
	FROM @links
	ORDER BY [primarySequence], [secondarySequence]

END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetParentValues]    Script Date: 02/01/2014 20:52:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetParentValues] (
	@piScreenID 		integer,
	@piParentTableID 	integer,
	@piParentRecordID 	integer
)
AS
BEGIN
	
	SET NOCOUNT ON;
	
	/* Return a recordset of the parent record values required for controls in the given screen. */
	DECLARE 
		@iUserGroupID		integer,
		@sRoleName			sysname,
		@iTempCount 		integer,
		@iParentTableType	integer,
		@sParentTableName	sysname,
		@iParentChildViewID	integer,
		@sParentRealSource	varchar(255),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@sNewBit			varchar(MAX),
		@sSelectString 		varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	varchar(255),
		@sParentSelectSQL	nvarchar(MAX),
		@sTemp				varchar(MAX),
		@fColumns			bit,
		@sSQL				nvarchar(MAX),
		@sActualUserName	sysname;

	SET @sParentSelectSQL  = 'SELECT ';
	SET @fColumns = 0;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(tableViewName sysname);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @columnPermissions TABLE(tableViewName	sysname,
		columnName	sysname,
		granted		bit);

	SELECT @iTempCount = COUNT(*)
	FROM ASRSysControls
	INNER JOIN ASRSysColumns ON ASRSysControls.columnID = ASRSysColumns.columnId
		AND ASRSysColumns.tableID = @piParentTableID
	WHERE ASRSysControls.screenID = @piScreenID
		AND ASRSysControls.columnID > 0;

	IF @iTempCount = 0 RETURN;

	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables
		WHERE tableID = @piParentTableID;

	IF @iParentTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		INSERT INTO @columnPermissions
		SELECT 
			sysobjects.name,
			syscolumns.name,
			CASE p.protectType
			        	WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END 
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		INNER JOIN syscolumns ON p.id = syscolumns.id
		WHERE p.UID = @iUserGroupID
			AND p.action = 193 
			AND syscolumns.name <> 'timestamp'
			AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @piParentTableID 
			UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @piParentTableID)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		SET @sParentRealSource = @sParentTableName;
	END
	ELSE
	BEGIN
		/* Get permitted child view on the parent table. */
		SELECT @iParentChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piParentTableID
			AND role = @sRoleName;
				
		IF @iParentChildViewID IS null SET @iParentChildViewID = 0;
				
		IF @iParentChildViewID > 0 
		BEGIN
			SET @sParentRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iParentChildViewID) +
				'#' + replace(@sParentTableName, ' ', '_') +
				'#' + replace(@sRoleName, ' ', '_');
			SET @sParentRealSource = left(@sParentRealSource, 255);

			INSERT INTO @columnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	END

	/* Populate the temporary table with info for all columns used in the screen controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysControls
	INNER JOIN ASRSysColumns ON ASRSysColumns.columnId = ASRSysControls.columnID
	WHERE ASRSysControls.screenID = @piScreenID
		AND ASRSysControls.columnID > 0
		AND ASRSysColumns.tableID = @piParentTableID;
	
	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;
	
		/* Get the select permission on the column. */
		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM @columnPermissions
		WHERE tableViewName = @sParentRealSource
			AND columnName = @sColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			IF @fColumns = 1
			BEGIN
				SET @sTemp = ',';
				SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;
			END

			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sNewBit = 'convert(varchar(10), ' + @sParentRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
			END
			ELSE
			BEGIN
				 /* Non-date */
				SET @sNewBit = @sParentRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
			END

			SET @fColumns = 1;
			SET @sParentSelectSQL = @sParentSelectSQL + @sNewBit;
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = '';

			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @columnPermissions
			WHERE tableViewName <> @sParentRealSource
				AND columnName = @sColumnName
				AND granted = 1;

			OPEN viewCursor;
			FETCH NEXT FROM viewCursor INTO @sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Column CAN be read from the view. */
				SET @fSelectGranted = 1;

				IF len(@sSelectString) = 0 SET @sSelectString = 'CASE';

				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sSelectString = @sSelectString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)';
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sSelectString = @sSelectString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
				END

				/* Add the view to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
				WHERE tableViewName = @sViewName;

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @joinParents (tableViewName) VALUES(@sViewName);
				END

				FETCH NEXT FROM viewCursor INTO @sViewName;
			END
			CLOSE viewCursor;
			DEALLOCATE viewCursor;

			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = @sSelectString +
					' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']';

				IF @fColumns = 1
				BEGIN
					SET @sTemp = ',';
					SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;
				END

				SET @fColumns = 1;
				SET @sParentSelectSQL = @sParentSelectSQL + @sSelectString;
			END
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	IF @fColumns = 0 RETURN;

	SET @sTemp = ' FROM ' + @sParentRealSource;
	SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

	DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableViewName
	FROM @joinParents;

	OPEN joinCursor;
	FETCH NEXT FROM joinCursor INTO @sTableViewName;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sTemp = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sParentRealSource + '.ID = ' + @sTableViewName + '.ID';
		SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

		FETCH NEXT FROM joinCursor INTO @sTableViewName;
	END
	CLOSE joinCursor;
	DEALLOCATE joinCursor;

	SET @sTemp = ' WHERE ' + @sParentRealSource + '.ID = ' + convert(varchar(100), @piParentRecordID);
	SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

	EXECUTE sp_executeSQL @sParentSelectSQL;
	
END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetSelfServiceRecordID]    Script Date: 02/01/2014 20:52:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetSelfServiceRecordID] (
	@piRecordID		integer 		OUTPUT,
	@piRecordCount	integer 		OUTPUT,
	@piViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iUserGroupID	integer,
		@sUserGroupName			sysname,
		@sActualUserName		sysname,
		@sViewName		sysname,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@iRecordID			integer,
		@iRecordCount		integer, 
		@fSysSecMgr			bit,
		@fAccessGranted		bit;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	SET @iRecordID = 0;
	SET @iRecordCount = 0;

	SELECT @sViewName = viewName
		FROM ASRSysViews
		WHERE viewID = @piViewID;

	IF len(@sViewName) > 0
	BEGIN
		/* Check if the user has permission to read the Self-service view. */
		exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;

		IF @fSysSecMgr = 1
		BEGIN
			SET @fAccessGranted = 1;
		END
		ELSE
		BEGIN
		
			SELECT @fAccessGranted =
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM ASRSysProtectsCache p
				INNER JOIN sysobjects ON p.id = sysobjects.id
				INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.UID = @iUserGroupID
				AND p.action = 193 
				AND syscolumns.name = 'ID'
				AND sysobjects.name = @sViewName
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	
		IF @fAccessGranted = 1
		BEGIN
			SET @sCommand = 'SELECT @iValue = COUNT(ID)' + 
				' FROM ' + @sViewName;
			SET @sParamDefinition = N'@iValue integer OUTPUT';
			EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT;

			IF @iRecordCount = 1 
			BEGIN
				SET @sCommand = 'SELECT @iValue = ' + @sViewName + '.ID ' + 
					' FROM ' + @sViewName;
				SET @sParamDefinition = N'@iValue integer OUTPUT';
				EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordID OUTPUT;
			END
		END
	END

	SET @piRecordID = @iRecordID;
	SET @piRecordCount = @iRecordCount;
END
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetSummaryValues]    Script Date: 02/01/2014 20:52:30 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spASRIntGetSummaryValues] (
	@piHistoryTableID	integer,
	@piParentTableID 	integer,
	@piParentRecordID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(255),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(255),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sRootTable 		varchar(255),
		@sSelectString 		varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	varchar(255),
		@sTemp				varchar(MAX),
		@sSelectSQL			nvarchar(MAX),
		@sActualUserName	sysname,
		@strTempSepText		varchar(500);

	SET @sSelectSQL = '';
		
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
		INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
		OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS';

	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID;

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(tableViewName	sysname);

	/* Create a temporary table of the 'read' column permissions for all tables/views used. */
	DECLARE @columnPermissions TABLE (tableViewName	sysname,
				columnName	sysname,
				granted		bit);

	-- Cached view of SysProtects
	DECLARE @SysProtects TABLE([ID] int, [ProtectType] tinyint, [Columns] varbinary(8000));

	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO @ColumnPermissions
		SELECT 
			@sParentTableName,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns 
		WHERE ASRSysColumns.tableID = @piParentTableID;
	END
	ELSE
	BEGIN
		IF @iParentTableType <> 2 /* ie. top-level or lookup */
		BEGIN

			-- Get list of views/table columns that are summary fields
			DECLARE @SummaryColumns TABLE ([ID] int, [TableName] sysname, [ColumnName] sysname, [ColID] int)
			INSERT @SummaryColumns
				SELECT sysobjects.id, sysobjects.name,
					syscolumns.name, syscolumns.ColID
				FROM sysobjects
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.name IN (SELECT ASRSysTables.tableName
												FROM ASRSysTables
												WHERE ASRSysTables.tableID = @piParentTableID
											UNION SELECT ASRSysViews.viewName
												FROM ASRSysViews
												WHERE ASRSysViews.viewTableID = @piParentTableID)
					AND syscolumns.name IN (SELECT ac.ColumnName
												FROM ASRSysSummaryFields am
												INNER JOIN ASRSysColumns ac ON am.ParentColumnID = ac.ColumnID
												WHERE HistoryTableID = @piHistoryTableID);

			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM ASRSysProtectsCache
				WHERE [UID] = @iUserGroupID AND Action = 193;

			-- Generate security context on selected columns
			INSERT INTO @ColumnPermissions
				SELECT sm.TableName,
					sm.ColumnName,
					CASE p.protectType
						WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @SysProtects p
				INNER JOIN @SummaryColumns sm ON p.id = sm.id
				WHERE (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) = 0));

		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @piParentTableID
				AND role = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sParentRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sParentTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sParentRealSource = left(@sParentRealSource, 255);
			END

			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM ASRSysProtectsCache
				WHERE [UID] = @iUserGroupID AND Action = 193;

			INSERT INTO @ColumnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	END

	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnId, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnId
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence;

	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		/* Get the select permission on the column. */

		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM @ColumnPermissions
		WHERE tableViewName = @sParentTableName
			AND columnName = @sColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			SET @sTemp = ',';
			IF LEN(@sSelectSQL) > 0
				SET @sSelectSQL = @sSelectSQL + @sTemp;

			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sTemp = 'convert(varchar(10), ' + @sParentTableName + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END 
			ELSE
			BEGIN
				 /* Non-date */
				SET @sTemp = @sParentTableName + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END

			/* Add the table to the array of tables/views to join if it has not already been added. */
			SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
				WHERE tableViewName = @sParentTableName;

			IF @iTempCount = 0
			BEGIN
				INSERT INTO @joinParents (tableViewName) VALUES(@sParentTableName);
			END
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = '';

			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @ColumnPermissions
			WHERE tableViewName <> @sParentTableName
				AND columnName = @sColumnName
				AND granted = 1;

			OPEN viewCursor;
			FETCH NEXT FROM viewCursor INTO @sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Column CAN be read from the view. */
				SET @fSelectGranted = 1;

				IF len(@sSelectString) > 0 SET @sSelectString = @sSelectString + ',';

				IF @iColumnDataType = 11 /* Date */
					SET @sSelectString = @sSelectString + 'convert(varchar(10),' + @sViewName + '.' + @sColumnName + ',101)';
				ELSE
					SET @sSelectString = @sSelectString + @sViewName + '.' + @sColumnName;


				/* Add the view to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewName = @sViewName;

				IF @iTempCount = 0
					INSERT INTO @joinParents (tableViewName) VALUES(@sViewName);

				FETCH NEXT FROM viewCursor INTO @sViewName;
			END
			CLOSE viewCursor;
			DEALLOCATE viewCursor;

			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = 'COALESCE(' + @sSelectString + ', NULL) AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sTemp = ',';

				IF LEN(@sSelectSQL) > 0
					SET @sSelectSQL = @sSelectSQL + @sTemp;

				SET @sTemp = @sSelectString;
				SET @sSelectSQL = @sSelectSQL + @sTemp;
				
			END
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = 'SELECT ' + @sSelectSQL ;

		SELECT @iTempCount = COUNT(tableViewName)
			FROM @joinParents;

		IF @iTempCount = 1 
		BEGIN
			SELECT TOP 1 @sRootTable = tableViewName
			FROM @joinParents;
		END
		ELSE
		BEGIN
			SET @sRootTable = @sParentTableName;
		END

		SET @sTemp = ' FROM ' + @sRootTable;
		SET @sSelectSQL = @sSelectSQL + @sTemp;

		/* Add the join code. */
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @joinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sTableViewName <> @sRootTable
			BEGIN
				SET @sTemp = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRootTable + '.ID' + ' = ' + @sTableViewName + '.ID';
				SET @sSelectSQL = @sSelectSQL + @sTemp
			END

			FETCH NEXT FROM joinCursor INTO @sTableViewName;
		END
		CLOSE joinCursor;
		DEALLOCATE joinCursor;

		SET @sTemp = ' WHERE ' + @sRootTable + '.id = ' + convert(varchar(255), @piParentRecordID);
		SET @sSelectSQL = @sSelectSQL + @sTemp;

	END

	-- Run the constructed SQL SELECT string.
	EXEC sp_executeSQL @sSelectSQL;

END

GO

CREATE PROCEDURE [dbo].[spASRIntGetColumnPermissions]
	(@SourceList AS dbo.dataPermissions READONLY)
AS
BEGIN
	
	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
			@sActualUserName	sysname,
			@sRoleName				sysname;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	SELECT UPPER(so.name) AS tableViewName, UPPER(syscolumns.name) AS columnName, p.action
		, CASE p.protectType WHEN 205 THEN 1 WHEN 204 THEN 1 ELSE 0 END AS permission, sl.*
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects so ON p.id = so.id
		INNER JOIN @SourceList sl ON sl.name = so.name
		INNER JOIN syscolumns ON p.id = syscolumns.id 
		WHERE p.uid = @iUserGroupID
			AND (p.action = 193 OR p.action = 197)
			AND syscolumns.name <> 'timestamp'
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0) OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		ORDER BY tableViewName;

END

GO

DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
		IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
		BEGIN
				SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END
		ELSE
		BEGIN
				SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END

		FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

GO

GRANT EXEC ON TYPE::[dbo].[DataPermissions] TO ASRSysGroup

GO