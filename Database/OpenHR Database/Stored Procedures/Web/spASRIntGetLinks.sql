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
		@iElement_Type		integer,		-- 2 = chart
		@isOrgChartEnabled	integer;

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
								WHEN @iUtilType = 35 THEN 'NINEBOXGRID'
								WHEN @iUtilType = 38 THEN 'TALENTREPORTS'
								WHEN @iUtilType = 39 THEN 'ORGREPORTING'

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
				IF @iUtilType = 35 /* 9-Box Grid Reports */
				BEGIN				
					SELECT @iBaseTableID = TableID
					FROM ASRSysCrossTab
					WHERE CrossTabID = @iUtilID
					AND CrossTabType = 4;
				END
				IF @iUtilType = 38 -- Talent Reports
				BEGIN				
					SELECT @iBaseTableID = MatchTableID
					FROM ASRSysTalentReports WHERE ID = @iUtilID;
				END

			    IF @iUtilType = 39 -- Organisation Reports
				BEGIN
					SELECT @iBaseTableID = v.ViewTableID
					FROM ASRSysOrganisationReport r
						INNER JOIN ASRSysViews v ON v.ViewID = r.BaseViewID
						WHERE r.ID = @iUtilID;
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

	/* Remove the organisation chart links if the user does not have run permission from system manager. */
	SELECT @isOrgChartEnabled = isnull(ParameterValue, -1)
	FROM ASRSysModuleSetup
	WHERE ModuleKey = 'MODULE_HIERARCHY' AND ParameterKey = 'Param_DisableSimpleChart';

	IF @isOrgChartEnabled = -1
	BEGIN
		DELETE FROM @Links
		WHERE Element_Type = 6;
	END

	SELECT ASRSysSSIntranetLinks.*, 
		CASE 
			WHEN ASRSysSSIntranetLinks.utilityType = 9 THEN ASRSysMailMergeName.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 2 THEN ASRSysCustomReportsName.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 17 THEN ASRSysCalendarReports.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 35 THEN ASRSysCrossTab.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 38 THEN ASRSysTalentReports.MatchTableID
			WHEN ASRSysSSIntranetLinks.utilityType = 39 THEN ASRSysOrganisationReport.BaseViewID
			WHEN ASRSysSSIntranetLinks.utilityType = 25 THEN 0
			ELSE null
		END AS [baseTable],
		ASRSysColumns.ColumnName as [Chart_ColumnName],
		tvL.DrillDownHidden as [DrillDownHidden]
	FROM ASRSysSSIntranetLinks
			LEFT OUTER JOIN ASRSysMailMergeName 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysMailMergeName.MailMergeID AND ASRSysSSIntranetLinks.utilityType = 9
			LEFT OUTER JOIN ASRSysCalendarReports 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysCalendarReports.ID	AND ASRSysSSIntranetLinks.utilityType = 17
			LEFT OUTER JOIN ASRSysCrossTab 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysCrossTab.CrossTabID AND ASRSysSSIntranetLinks.utilityType = 35
			LEFT OUTER JOIN ASRSysCustomReportsName 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysCustomReportsName.ID	AND ASRSysSSIntranetLinks.utilityType = 2
			LEFT OUTER JOIN ASRSysTalentReports 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysTalentReports.ID AND ASRSysSSIntranetLinks.utilityType = 38
			LEFT OUTER JOIN ASRSysOrganisationReport 
				ON ASRSysSSIntranetLinks.utilityID = ASRSysOrganisationReport.ID AND ASRSysSSIntranetLinks.utilityType = 39
			LEFT OUTER JOIN ASRSysColumns
				ON ASRSysSSIntranetLinks.Chart_ColumnID = ASRSysColumns.columnId		
			LEFT OUTER JOIN @Links tvL
			ON ASRSysSSIntranetLinks.ID = tvL.ID
	WHERE ASRSysSSIntranetLinks.ID IN (SELECT ID FROM @Links)
	ORDER BY ASRSysSSIntranetLinks.linkOrder;
	
END
