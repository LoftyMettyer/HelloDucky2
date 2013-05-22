CREATE PROCEDURE [dbo].[spASRIntGetNavigationLinks]
(
		@plngTableID	integer,
		@plngViewID		integer
)
AS
BEGIN
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

	SET NOCOUNT ON;

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
			INNER JOIN #SysProtects p 
				ON (syscolumns.id = p.id
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
			WHERE syscolumns.name = 'timestamp'
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