CREATE PROCEDURE [dbo].[sp_ASRIntDefUsage] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sExecSQL		nvarchar(MAX),
		@sJobTypeName		varchar(255),
		@sCurrentUser		sysname,
		@sDescription		varchar(MAX),
		@sName				varchar(255), 
		@sUserName			varchar(255), 
		@sAccess			varchar(MAX),
		@fIsBatch			bit,
		@sUtilType			varchar(255),
		@iCompID			integer,
		@iRootExprID		integer,
		@sRoleName			varchar(255),
		@fSysSecMgr			bit,
		@iCount				integer,
		@sActualUserName	sysname,
		@iUserGroupID		integer;
		
	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
	
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	SET @sExecSQL = '';
	SET @sCurrentUser = SYSTEM_USER;

	DECLARE @results TABLE([description] varchar(MAX));
	DECLARE @rootExprs TABLE(exprID integer);

	IF @intType = 11 OR @intType = 12
	BEGIN
		/* Create a table of IDs of the expressions that use the given filter or calc. */
		DECLARE expr_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT componentID 
			FROM ASRSysExprComponents
			WHERE calculationID = @intID
				OR filterID = @intID
				OR (fieldSelectionFilter = @intID AND type = 1)
		OPEN expr_cursor
		FETCH NEXT FROM expr_cursor INTO @iCompID
		WHILE (@@fetch_status = 0)
		BEGIN
			execute sp_ASRIntGetRootExpressionIDs @iCompID, @iRootExprID OUTPUT
			IF @iRootExprID > 0
			BEGIN
				INSERT INTO @rootExprs (exprID) VALUES (@iRootExprID)
			END
			FETCH NEXT FROM expr_cursor INTO @iCompID
		END
		CLOSE expr_cursor
		DEALLOCATE expr_cursor
	END

	IF @intType IN(1, 2, 9, 17, 35, 38)
	BEGIN
		/* Reports & Utilities
		Check for usage in Batch Jobs */
		IF @intType = 1 SET @sJobTypeName = 'CROSS TAB'
		IF @intType = 2 SET @sJobTypeName = 'CUSTOM REPORT'
		IF @intType = 9 SET @sJobTypeName = 'MAIL MERGE' 
		IF @intType = 17 SET @sJobTypeName = 'CALENDAR REPORT'
		IF @intType = 35 SET @sJobTypeName = '9-BOX GRID REPORT'
		IF @intType = 38 SET @sJobTypeName = 'TALENT REPORT'
		
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT ASRSysBatchJobName.Name, 
				ASRSysBatchJobName.UserName, 
				ASRSysBatchJobAccess.Access,
				AsrSysBatchJobName.IsBatch
			FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID
			INNER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.groupname = @sRoleName
			WHERE ASRSysBatchJobDetails.JobType = @sJobTypeName
				AND ASRSysBatchJobDetails.JobID = @intID
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @fIsBatch = 1 BEGIN
				SET @sDescription = 'Batch Job: '
			END ELSE BEGIN
				SET @sDescription = 'Report Pack: '
			END

			IF (@sUserName <> @sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '&lt;Hidden by ' + @sUserName + '&gt;'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor

		SELECT @iCount = COUNT(*) 
		FROM [ASRSysSSIntranetLinks]
		WHERE [ASRSysSSIntranetLinks].[utilityID] = @intID
			AND [ASRSysSSIntranetLinks].[utilityType] = @intType
		IF @iCount > 0
		BEGIN
		   	INSERT INTO @results (description) VALUES ('Self-service intranet link')
		END
	END

	IF @intType = 10
	BEGIN
		/* Picklists 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports, Calendar Reports and Mail Merges*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Cross Tab', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE PickListID =@intID
					AND ASRSysCrossTab.CrossTabType <> 4
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.pickListID = @intID
			UNION
				SELECT DISTINCT '9-Box Grid Report', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE PickListID =@intID
					AND ASRSysCrossTab.CrossTabType = 4
			UNION
				SELECT DISTINCT 'Talent Report', n.Name, n.UserName, a.Access
				FROM ASRSysTalentReports n
				INNER JOIN ASRSysTalentReportAccess a ON n.ID = a.ID
					AND a.groupname = @sRoleName
				WHERE n.BasePicklistID = @intID OR n.MatchPicklistID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.pickList = @intID OR ASRSysExportName.Parent1Picklist = @intID OR ASRSysExportName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.pickListID = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access 
				FROM ASRSysCustomReportsName
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.PickList = @intID 
					OR ASRSysCustomReportsName.Parent1Picklist = @intID 
					OR ASRSysCustomReportsName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END, 
					ASRSysMailMergeName.Name, 
					ASRSysMailMergeName.UserName, 
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName 
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE PickListID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning'
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report'
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Picklist = @intID
					OR ASRSysMatchReportName.Table2Picklist = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSysCalendarReportAccess.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.PickList = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				WHERE ASRSysRecordProfileName.pickListID = @intID
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '&lt;Hidden by ' + @sUserName + '&gt;'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 11
	BEGIN
		/* Filters 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Cross Tab',
					ASRSysCrossTab.Name,
					ASRSysCrossTab.UserName,
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE ASRSysCrossTab.FilterID = @intID
					AND ASRSysCrossTab.CrossTabType <> 4
			UNION
				SELECT DISTINCT '9-Box Grid Report', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE ASRSysCrossTab.FilterID = @intID
					AND ASRSysCrossTab.CrossTabType = 4
			UNION
				SELECT DISTINCT 'Talent Report', n.Name, n.UserName, a.Access
				FROM ASRSysTalentReports n
				INNER JOIN ASRSysTalentReportAccess a ON n.ID = a.ID
					AND a.groupname = @sRoleName
				WHERE n.BaseFilterID = @intID OR n.MatchFilterID = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsName
				LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.Filter = @intID
					OR ASRSysCustomReportsName.Parent1Filter = @intID
					OR ASRSysCustomReportsName.Parent2Filter = @intID
					OR ASRSYSCustomReportsChildDetails.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.FilterID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.Filter = @intID 
					OR ASRSysExportName.Parent1Filter = @intID
					OR ASRSysExportName.Parent2Filter = @intID
					OR ASRSysExportName.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.FilterID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeName.FilterID = @intID
			UNION
				SELECT DISTINCT
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning' 
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report' 
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Filter = @intID
					OR ASRSysMatchReportName.Table2Filter = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.filter = @intID
					OR ASRSysCalendarReportEvents.filterID = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID
				WHERE ASRSysRecordProfileName.FilterID = @intID
					OR ASRSYSRecordProfileTables.FilterID = @intID		

		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '&lt;Hidden by ' + @sUserName + '&gt;'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 12
	BEGIN
		/* Calculation.
		Check for usage in Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports 
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.DescriptionExpr =@intID 
					OR ASRSysCalendarReports.StartDateExpr = @intID 
					OR ASRSysCalendarReports.EndDateExpr = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name,
					ASRSysCustomReportsName.UserName,
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsDetails
				INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.CustomReportID = ASRSysCustomReportsName.ID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' 
					AND ASRSysCustomReportsDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportDetails
				INNER JOIN ASRSysExportName ON ASRSysExportDetails.ID = ASRSysExportName.ID 
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysExportDetails.type) = 'X' 
					AND ASRSysExportDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalItems
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID 
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalItems.ValueType = 4 
					AND ASRSysGlobalItems.ExprID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeColumns.ColumnID = @intID
					AND upper(ASRSysMailMergeColumns.type) = 'E'
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '&lt;Hidden by ' + @sUserName + '&gt;'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + '''';
			END
    
    		INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess;
		END
		
		CLOSE usage_cursor;
		DEALLOCATE usage_cursor;
	END

	/* Return the usage records. */
	SELECT * FROM @results ORDER BY description;

END