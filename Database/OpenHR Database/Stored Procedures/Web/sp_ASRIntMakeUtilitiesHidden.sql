CREATE PROCEDURE [dbo].[sp_ASRIntMakeUtilitiesHidden] (
	@piUtilityType		integer,
	@piUtilityID		integer
) AS
BEGIN
	/* Hide any utilities the use the given picklist/filter/calculation. */
	DECLARE
		@sCurrentUser		sysname,
		@sUtilName			varchar(255),
		@iUtilID			integer,
		@sUtilOwner			varchar(255),
		@sUtilAccess		varchar(MAX),
		@iCount				integer,
		@sJobName			varchar(255),
		@iNonHiddenCount	integer,
		@iScheduled			integer, 
		@sRoleToPrompt		sysname,
		@sCurrentUserGroup	sysname,
		@superCursor		cursor,
		@iTemp				integer,
		@iUserGroupID		integer,
		@sActualUserName	sysname;

	SET @sCurrentUser = SYSTEM_USER;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sCurrentUserGroup OUTPUT,
		@iUserGroupID OUTPUT;

	DECLARE @batchJobIDs		TABLE(id integer)
	DECLARE @calendarReportsIDs TABLE(id integer)
	DECLARE @careerIDs			TABLE(id integer)
	DECLARE @crossTabIDs		TABLE(id integer)
	DECLARE @customReportsIDs	TABLE(id integer)
	DECLARE @dataTransferIDs	TABLE(id integer)
	DECLARE @exportIDs			TABLE(id integer)
	DECLARE @globalAddIDs		TABLE(id integer)
	DECLARE @globalUpdateIDs	TABLE(id integer)
	DECLARE @globalDeleteIDs	TABLE(id integer)
	DECLARE @labelsIDs			TABLE(id integer)
	DECLARE @mailMergeIDs		TABLE(id integer)
	DECLARE @matchReportIDs		TABLE(id integer)
	DECLARE @recordProfileIDs	TABLE(id integer)
	DECLARE @successionIDs		TABLE(id integer)
	DECLARE @filterIDs			TABLE(id integer)
	DECLARE @calculationIDs		TABLE(id integer)
	DECLARE @expressionIDs		TABLE(id integer)
	DECLARE @superExpressionIDs	TABLE(id integer)

	IF (@piUtilityType = 12) OR (@piUtilityType = 11)
	BEGIN
		/* Calculation/Filter. */

    /*---------------------------------------------------*/
    /* Check Calculations/Filters For This Expression		*/
    /* NB. This check must be made before checking the reports/utilities	*/
    /*---------------------------------------------------*/
		INSERT INTO @expressionIDs (id) VALUES (@piUtilityID)
		
		exec spASRIntGetAllExprRootIDs @piUtilityID, @superCursor output
		
		FETCH NEXT FROM @superCursor INTO @iTemp
		WHILE (@@fetch_status = 0)
		BEGIN
			INSERT INTO @superExpressionIDs (id) VALUES (@iTemp)
			
			FETCH NEXT FROM @superCursor INTO @iTemp 
		END
		CLOSE @superCursor
		DEALLOCATE @superCursor

		INSERT INTO @expressionIDs (id) SELECT id FROM @superExpressionIDs

		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExpressions.Name,
				ASRSysExpressions.exprID AS [ID],
				ASRSysExpressions.Username,
				ASRSysExpressions.Access
			FROM ASRSysExpressions
			WHERE ASRSysExpressions.exprID IN (SELECT id FROM @superExpressionIDs)
				AND ASRSysExpressions.type = 10

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @sUtilAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Calculation whose owner is the same */
				IF @sUtilAccess <> 'HD'
				BEGIN
					INSERT INTO @calculationIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @sUtilAccess
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExpressions.Name,
				ASRSysExpressions.exprID AS [ID],
				ASRSysExpressions.Username,
				ASRSysExpressions.Access
			FROM ASRSysExpressions
			WHERE ASRSysExpressions.exprID IN (SELECT id FROM @superExpressionIDs)
				AND ASRSysExpressions.type = 11

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @sUtilAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Filter whose owner is the same */
				IF @sUtilAccess <> 'HD'
				BEGIN
					INSERT INTO @filterIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @sUtilAccess
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/*---------------------------------------------------*/
		/* Check Calendar Reports for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT AsrSysCalendarReports.Name,
			AsrSysCalendarReports.ID,
			AsrSysCalendarReports.Username,
			COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]
		FROM AsrSysCalendarReports
		LEFT OUTER JOIN ASRSYSCalendarReportEvents ON AsrSysCalendarReports.ID = ASRSYSCalendarReportEvents.calendarReportID
		LEFT OUTER JOIN ASRSYSCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID
			AND ASRSYSCalendarReportAccess.access <> 'HD'
			AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE AsrSysCalendarReports.DescriptionExpr IN (SELECT id FROM @expressionIDs)
			OR AsrSysCalendarReports.StartDateExpr IN (SELECT id FROM @expressionIDs)
      OR AsrSysCalendarReports.EndDateExpr IN (SELECT id FROM @expressionIDs)
      OR ASRSysCalendarReports.Filter IN (SELECT id FROM @expressionIDs)
      OR ASRSYSCalendarReportEvents.FilterID IN (SELECT id FROM @expressionIDs)
		GROUP BY AsrSysCalendarReports.Name,
			AsrSysCalendarReports.ID,
 			AsrSysCalendarReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Calendar Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @calendarReportsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Calendar Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @calendarReportsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCalendarReports.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCalendarReports ON ASRSysCalendarReports.ID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Calendar Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @calendarReportsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCalendarReports.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Calendar Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Career Progression for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 2
			AND (ASRSysMatchReportName.table1Filter IN (SELECT id FROM @expressionIDs)
      OR ASRSysMatchReportName.table2Filter IN (SELECT id FROM @expressionIDs))
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Career Progression whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @careerIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Career Progressions are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @careerIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Career Progression'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @careerIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Career Progression in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Cross Tabs for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT AsrSysCrossTab.Name,
			AsrSysCrossTab.[CrossTabID] AS [ID],
			AsrSysCrossTab.Username,
			COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]
		FROM AsrSysCrossTab
		LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID
			AND ASRSYSCrossTabAccess.access <> 'HD'
			AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
		WHERE AsrSysCrossTab.FilterID IN (SELECT id FROM @expressionIDs)
		GROUP BY AsrSysCrossTab.Name,
			AsrSysCrossTab.crossTabID,
 			AsrSysCrossTab.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Cross Tab whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @crossTabIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Cross Tabs are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @crossTabIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCrossTab.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCrossTab ON ASRSysCrossTab.CrossTabID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Cross Tab'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @crossTabIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCrossTab.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a cross tab in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Custom Reports For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysCustomReportsName.Name,
        ASRSysCustomReportsName.ID,
        ASRSysCustomReportsName.Username,
        COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]
       FROM ASRSysCustomReportsName
       LEFT OUTER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = AsrSysCustomReportsDetails.CustomReportID
       LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID
       LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID
				AND ASRSYSCustomReportAccess.access <> 'HD'
        AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysCustomReportsName.Filter IN (SELECT id FROM @expressionIDs)
				OR ASRSysCustomReportsName.Parent1Filter IN (SELECT id FROM @expressionIDs)
        OR ASRSysCustomReportsName.Parent2Filter IN (SELECT id FROM @expressionIDs)
        OR ASRSYSCustomReportsChildDetails.ChildFilter IN (SELECT id FROM @expressionIDs)
        OR(AsrSysCustomReportsDetails.Type = 'E' 
					AND AsrSysCustomReportsDetails.ColExprID IN (SELECT id FROM @expressionIDs))
      GROUP BY ASRSysCustomReportsName.Name,
				ASRSysCustomReportsName.ID,
        ASRSysCustomReportsName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Custom Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @customReportsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Custom Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @customReportsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCustomReportsName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysCustomReportsName ON AsrSysCustomReportsname.ID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Custom Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @customReportsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCustomReportsName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Custom Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Data Transfer For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysDataTransferName.Name,
        ASRSysDataTransferName.DataTransferID AS [ID],
        ASRSysDataTransferName.Username,
        COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]
       FROM ASRSysDataTransferName
       LEFT OUTER JOIN ASRSYSDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID
				AND ASRSYSDataTransferAccess.access <> 'HD'
        AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysDataTransferName.FilterID IN (SELECT id FROM @expressionIDs)
      GROUP BY ASRSysDataTransferName.Name,
				ASRSysDataTransferName.DataTransferID,
        ASRSysDataTransferName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Data Transfer whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @dataTransferIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Data Transfers are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @dataTransferIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysDataTransferName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysDataTransferName ON ASRSysDataTransferName.DataTransferID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Data Transfer'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @dataTransferIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysDataTransferName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Data Transfer in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Envelopes & Labels For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT AsrSysMailMergeName.Name,
        AsrSysMailMergeName.MailMergeID AS [ID],
        AsrSysMailMergeName.Username,
        COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
       FROM AsrSysMailMergeName
			 LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID
       LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
				AND ASRSYSMailMergeAccess.access <> 'HD'
        AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
      WHERE AsrSysMailMergeName.isLabel = 1
        AND ((AsrSysMailMergeName.FilterID IN (SELECT id FROM @expressionIDs))
				OR (AsrSysMailMergeColumns.Type = 'E' 
					AND AsrSysMailMergeColumns.ColumnID IN (SELECT id FROM @expressionIDs)))
      GROUP BY AsrSysMailMergeName.Name,
				AsrSysMailMergeName.MailMergeID,
        AsrSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Envelopes & Labels whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @labelsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Envelopes & Labels are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @labelsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					AsrSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysMailMergeName ON AsrSysMailMergeName.MailMergeID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Envelopes & Labels'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @labelsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					AsrSysMailMergeName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Envelopes & Labels in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Export For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExportName.Name,
        ASRSysExportName.ID,
        ASRSysExportName.Username,
        COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]
       FROM ASRSysExportName
			 LEFT OUTER JOIN AsrSysExportDetails ON ASRSysExportName.ID = AsrSysExportDetails.exportID
       LEFT OUTER JOIN ASRSYSExportAccess ON ASRSysExportName.ID = ASRSYSExportAccess.ID
				AND ASRSYSExportAccess.access <> 'HD'
        AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
      WHERE AsrSysExportName.Filter IN (SELECT id FROM @expressionIDs)
        OR AsrSysExportName.Parent1Filter IN (SELECT id FROM @expressionIDs)
        OR AsrSysExportName.Parent2Filter IN (SELECT id FROM @expressionIDs)
        OR AsrSysExportName.ChildFilter IN (SELECT id FROM @expressionIDs)
        OR (AsrSysExportDetails.Type = 'X' 
					AND AsrSysExportDetails.ColExprID IN (SELECT id FROM @expressionIDs))
      GROUP BY AsrSysExportName.Name,
				AsrSysExportName.ID,
        AsrSysExportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Export whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @exportIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Export are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @exportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysExportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysExportName ON ASRSysExportName.ID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Export'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @exportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysExportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Export in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Add For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
			 LEFT OUTER JOIN AsrSysGlobalItems ON ASRSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE AsrSysGlobalFunctions.Type = 'A' 
				AND ((AsrSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs))
				OR (AsrSysGlobalItems.ValueType = 4 
					AND AsrSysGlobalItems.ExprID IN (SELECT id FROM @expressionIDs)))
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Add whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalAddIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Adds are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalAddIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Add'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalAddIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Add in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) > 0) 
						OR (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Update For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
			 LEFT OUTER JOIN AsrSysGlobalItems ON ASRSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE AsrSysGlobalFunctions.Type = 'U' 
				AND ((AsrSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs))
				OR (AsrSysGlobalItems.ValueType = 4 
					AND AsrSysGlobalItems.ExprID IN (SELECT id FROM @expressionIDs)))
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Update whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalUpdateIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Updates are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalUpdateIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Update'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalUpdateIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Update in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Delete For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE AsrSysGlobalFunctions.Type = 'D' 
				AND AsrSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs)
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Delete whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalDeleteIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Deletes are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalDeleteIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Delete'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalDeleteIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Delete in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Mail Merge For This Expression. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT AsrSysMailMergeName.Name,
        AsrSysMailMergeName.MailMergeID AS [ID],
        AsrSysMailMergeName.Username,
        COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
       FROM AsrSysMailMergeName
			 LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID
       LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
				AND ASRSYSMailMergeAccess.access <> 'HD'
        AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
      WHERE AsrSysMailMergeName.isLabel = 0
        AND ((AsrSysMailMergeName.FilterID IN (SELECT id FROM @expressionIDs))
				OR (AsrSysMailMergeColumns.Type = 'E' 
					AND AsrSysMailMergeColumns.ColumnID IN (SELECT id FROM @expressionIDs)))
      GROUP BY AsrSysMailMergeName.Name,
				AsrSysMailMergeName.MailMergeID,
        AsrSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Mail Merge whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @mailMergeIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Mail Merges are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @mailMergeIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					AsrSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysMailMergeName ON AsrSysMailMergeName.MailMergeID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Mail Merge'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @mailMergeIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					AsrSysMailMergeName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Mail Merge in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Match Report for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 0
			AND (ASRSysMatchReportName.table1Filter IN (SELECT id FROM @expressionIDs)
      OR ASRSysMatchReportName.table2Filter IN (SELECT id FROM @expressionIDs))
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Match Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @matchReportIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Match Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @matchReportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Match Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @matchReportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Match Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Record Profile for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysRecordProfileName.Name,
			ASRSysRecordProfileName.recordProfileID AS [ID],
			ASRSysRecordProfileName.Username,
			COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]
		FROM ASRSysRecordProfileName
		LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID
		LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID
			AND ASRSYSRecordProfileAccess.access <> 'HD'
			AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
		WHERE ASRSysRecordProfileName.FilterID IN (SELECT id FROM @expressionIDs)
			OR ASRSYSRecordProfileTables.FilterID IN (SELECT id FROM @expressionIDs)
		GROUP BY ASRSysRecordProfileName.Name,
			ASRSysRecordProfileName.recordProfileID,
 			ASRSysRecordProfileName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Record Profile whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @recordProfileIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Record Profiles are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @recordProfileIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysRecordProfileName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysRecordProfileName ON ASRSysRecordProfileName.recordProfileID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Record Profile'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @recordProfileIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysRecordProfileName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Record Profile in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Succession Planning for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 1
			AND (ASRSysMatchReportName.table1Filter IN (SELECT id FROM @expressionIDs)
      OR ASRSysMatchReportName.table2Filter IN (SELECT id FROM @expressionIDs))
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Succession Planning whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @successionIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Succession Plannings are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @successionIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Succession Planning'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @successionIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Succession Planning in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END
	END

	IF @piUtilityType = 10
	BEGIN
    /* Picklist */
    
		/*---------------------------------------------------*/
		/* Check Calendar Reports for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT AsrSysCalendarReports.Name,
			AsrSysCalendarReports.ID,
			AsrSysCalendarReports.Username,
			COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]
		FROM AsrSysCalendarReports
		LEFT OUTER JOIN ASRSYSCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID
			AND ASRSYSCalendarReportAccess.access <> 'HD'
			AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE AsrSysCalendarReports.Picklist = @piUtilityID
		GROUP BY AsrSysCalendarReports.Name,
			AsrSysCalendarReports.ID,
 			AsrSysCalendarReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Calendar Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @calendarReportsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Calendar Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @calendarReportsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCalendarReports.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCalendarReports ON ASRSysCalendarReports.ID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Calendar Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @calendarReportsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCalendarReports.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Calendar Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Career Progression for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 2
			AND (ASRSysMatchReportName.table1Picklist = @piUtilityID
      OR ASRSysMatchReportName.table2Picklist = @piUtilityID)
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Career Progression whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @careerIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Career Progressions are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @careerIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Career Progression'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @careerIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Career Progression in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Cross Tabs for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT AsrSysCrossTab.Name,
			AsrSysCrossTab.[CrossTabID] AS [ID],
			AsrSysCrossTab.Username,
			COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]
		FROM AsrSysCrossTab
		LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID
			AND ASRSYSCrossTabAccess.access <> 'HD'
			AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
		WHERE AsrSysCrossTab.PicklistID = @piUtilityID
		GROUP BY AsrSysCrossTab.Name,
			AsrSysCrossTab.crossTabID,
 			AsrSysCrossTab.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Cross Tab whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @crossTabIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Cross Tabs are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @crossTabIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCrossTab.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCrossTab ON ASRSysCrossTab.CrossTabID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Cross Tab'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @crossTabIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCrossTab.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a cross tab in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Custom Reports For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysCustomReportsName.Name,
        ASRSysCustomReportsName.ID,
        ASRSysCustomReportsName.Username,
        COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]
       FROM ASRSysCustomReportsName
       LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID
				AND ASRSYSCustomReportAccess.access <> 'HD'
        AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysCustomReportsName.Picklist = @piUtilityID
				OR ASRSysCustomReportsName.Parent1Picklist = @piUtilityID
				OR ASRSysCustomReportsName.Parent2Picklist = @piUtilityID
      GROUP BY ASRSysCustomReportsName.Name,
				ASRSysCustomReportsName.ID,
        ASRSysCustomReportsName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Custom Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @customReportsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Custom Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @customReportsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysCustomReportsName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysCustomReportsName ON AsrSysCustomReportsname.ID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Custom Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @customReportsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysCustomReportsName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Custom Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Data Transfer For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysDataTransferName.Name,
        ASRSysDataTransferName.DataTransferID AS [ID],
        ASRSysDataTransferName.Username,
        COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]
       FROM ASRSysDataTransferName
       LEFT OUTER JOIN ASRSYSDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID
				AND ASRSYSDataTransferAccess.access <> 'HD'
        AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysDataTransferName.PicklistID = @piUtilityID	
      GROUP BY ASRSysDataTransferName.Name,
				ASRSysDataTransferName.DataTransferID,
        ASRSysDataTransferName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Data Transfer whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @dataTransferIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Data Transfers are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @dataTransferIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysDataTransferName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysDataTransferName ON ASRSysDataTransferName.DataTransferID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Data Transfer'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @dataTransferIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysDataTransferName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Data Transfer in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Envelopes & Labels For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT AsrSysMailMergeName.Name,
        AsrSysMailMergeName.MailMergeID AS [ID],
        AsrSysMailMergeName.Username,
        COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
       FROM AsrSysMailMergeName
       LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
				AND ASRSYSMailMergeAccess.access <> 'HD'
        AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
      WHERE AsrSysMailMergeName.isLabel = 1
        AND AsrSysMailMergeName.PicklistID = @piUtilityID
      GROUP BY AsrSysMailMergeName.Name,
				AsrSysMailMergeName.MailMergeID,
        AsrSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Envelopes & Labels whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @labelsIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Envelopes & Labels are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @labelsIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					AsrSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysMailMergeName ON AsrSysMailMergeName.MailMergeID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Envelopes & Labels'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @labelsIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					AsrSysMailMergeName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Envelopes & Labels in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Export For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExportName.Name,
        ASRSysExportName.ID,
        ASRSysExportName.Username,
        COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]
       FROM ASRSysExportName
       LEFT OUTER JOIN ASRSYSExportAccess ON ASRSysExportName.ID = ASRSYSExportAccess.ID
				AND ASRSYSExportAccess.access <> 'HD'
        AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysExportName.Picklist = @piUtilityID
				OR ASRSysExportName.Parent1Picklist = @piUtilityID
				OR ASRSysExportName.Parent2Picklist = @piUtilityID
      GROUP BY AsrSysExportName.Name,
				AsrSysExportName.ID,
        AsrSysExportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Export whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @exportIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Export are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @exportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysExportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysExportName ON ASRSysExportName.ID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Export'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @exportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysExportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Export in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Add For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysGlobalFunctions.Type = 'A' 
				AND ASRSysGlobalFunctions.PicklistID = @piUtilityID
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Add whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalAddIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Adds are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalAddIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Add'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalAddIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Add in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Update For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysGlobalFunctions.Type = 'U' 
				AND ASRSysGlobalFunctions.PicklistID = @piUtilityID
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Update whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalUpdateIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Updates are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalUpdateIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Update'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalUpdateIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Update in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Global Delete For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysGlobalFunctions.Name,
        ASRSysGlobalFunctions.functionID AS [ID],
        ASRSysGlobalFunctions.Username,
        COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]
       FROM ASRSysGlobalFunctions
       LEFT OUTER JOIN ASRSYSGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID
				AND ASRSYSGlobalAccess.access <> 'HD'
        AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysGlobalFunctions.Type = 'D' 
				AND ASRSysGlobalFunctions.PicklistID = @piUtilityID
      GROUP BY ASRSysGlobalFunctions.Name,
				ASRSysGlobalFunctions.functionID,
        ASRSysGlobalFunctions.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Global Delete whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @globalDeleteIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Global Deletes are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @globalDeleteIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysGlobalFunctions.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalFunctions.FunctionID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Global Delete'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @globalDeleteIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysGlobalFunctions.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Global Delete in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*-----------------------------------------------------------------------*/
		/* Check Mail Merge For This Picklist. */
		/*-----------------------------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT AsrSysMailMergeName.Name,
        AsrSysMailMergeName.MailMergeID AS [ID],
        AsrSysMailMergeName.Username,
        COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
       FROM AsrSysMailMergeName
			 LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID
       LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
				AND ASRSYSMailMergeAccess.access <> 'HD'
        AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
      WHERE AsrSysMailMergeName.isLabel = 0
				AND ASRSysMailMergeName.PicklistID = @piUtilityID
      GROUP BY AsrSysMailMergeName.Name,
				AsrSysMailMergeName.MailMergeID,
        AsrSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Mail Merge whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @mailMergeIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Mail Merges are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @mailMergeIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					AsrSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN AsrSysMailMergeName ON AsrSysMailMergeName.MailMergeID = AsrSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Mail Merge'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @mailMergeIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					AsrSysMailMergeName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Mail Merge in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Match Report for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 0
			AND (ASRSysMatchReportName.table1Picklist = @piUtilityID
OR ASRSysMatchReportName.table2Picklist = @piUtilityID)
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Match Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @matchReportIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Match Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @matchReportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Match Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @matchReportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Match Report in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Record Profile for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysRecordProfileName.Name,
			ASRSysRecordProfileName.recordProfileID AS [ID],
			ASRSysRecordProfileName.Username,
			COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]
		FROM ASRSysRecordProfileName
		LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID
			AND ASRSYSRecordProfileAccess.access <> 'HD'
			AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
		WHERE ASRSysRecordProfileName.PicklistID = @piUtilityID
		GROUP BY ASRSysRecordProfileName.Name,
			ASRSysRecordProfileName.recordProfileID,
 			ASRSysRecordProfileName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Record Profile whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @recordProfileIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Record Profiles are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @recordProfileIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysRecordProfileName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysRecordProfileName ON ASRSysRecordProfileName.recordProfileID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Record Profile'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @recordProfileIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysRecordProfileName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Record Profile in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Succession Planning for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID AS [ID],
			ASRSysMatchReportName.Username,
			COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysMatchReportName
		LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID
			AND ASRSYSMatchReportAccess.access <> 'HD'
			AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name
				FROM sysusers
				INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
					AND ASRSysGroupPermissions.permitted = 1
				INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
 					AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
					OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE sysusers.uid = sysusers.gid
					AND sysusers.uid <> 0)
    WHERE ASRSysMatchReportName.matchReportType = 1
			AND (ASRSysMatchReportName.table1Picklist = @piUtilityID
      OR ASRSysMatchReportName.table2Picklist = @piUtilityID)
		GROUP BY ASRSysMatchReportName.Name,
			ASRSysMatchReportName.MatchReportID,
 			ASRSysMatchReportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Succession Planning whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					INSERT INTO @successionIDs (id) VALUES (@iUtilID)
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Succession Plannings are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @successionIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysMatchReportName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMatchReportName ON ASRSysMatchReportName.matchReportID = ASRSysBatchJobdetails.JobID
				LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
					AND ASRSysBatchJobAccess.access <> 'HD'
					AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
						FROM sysusers
						INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
							AND ASRSysGroupPermissions.permitted = 1
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
							OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
						WHERE sysusers.uid = sysusers.gid
							AND sysusers.uid <> 0)
				WHERE ASRSysBatchJobDetails.JobType = 'Succession Planning'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @successionIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysMatchReportName.Name

			OPEN check_cursor
			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			WHILE (@@fetch_status = 0)
			BEGIN
				exec spASRIntCurrentUserAccess 
					0,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF @sUtilOwner = @sCurrentUser
				BEGIN
					/* Found a Succession Planning in a batch job whose owner is the same */
					IF (@iScheduled <> 1) 
						OR (Len(@sRoleToPrompt) = 0) 
						OR (@sRoleToPrompt = @sCurrentUserGroup)
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END
	END

	/*---------------------------------------------------------*/
	/* Mark all relevent utilities as hidden. */
	/*---------------------------------------------------------*/
	
	/* Calculations */
	UPDATE ASRSysExpressions
	SET access = 'HD'
	WHERE exprID IN (SELECT id FROM @calculationIDs)

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @calculationIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 12

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (12, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 12
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor

	/* Filters */
	UPDATE ASRSysExpressions
	SET access = 'HD'
	WHERE exprID IN (SELECT id FROM @filterIDs)

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @filterIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 11

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (11, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 11
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor

	/* Batch Jobs */
	DELETE FROM ASRSysBatchJobAccess
	WHERE ID IN (SELECT id FROM @batchJobIDs)

	INSERT INTO ASRSysBatchJobAccess
		(ID, groupName, access)
		(SELECT ASRSysBatchJobName.ID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysBatchJobName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysBatchJobName.ID IN (SELECT id FROM @batchJobIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @batchJobIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 0

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (0, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 0
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Calendar Reports */
	DELETE FROM ASRSysCalendarReportAccess
	WHERE ID IN (SELECT id FROM @calendarReportsIDs)

	INSERT INTO ASRSysCalendarReportAccess
		(ID, groupName, access)
		(SELECT ASRSysCalendarReports.ID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysCalendarReports
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysCalendarReports.ID IN (SELECT id FROM @calendarReportsIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @calendarReportsIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 17

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (17, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 17
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Career Progression */
	DELETE FROM ASRSysMatchReportAccess
	WHERE ID IN (SELECT id FROM @careerIDs)

	INSERT INTO ASRSysMatchReportAccess
		(ID, groupName, access)
		(SELECT ASRSysMatchReportName.matchReportID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysMatchReportName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysMatchReportName.matchReportID IN (SELECT id FROM @careerIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @careerIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 24

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (24, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 24
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Cross Tabs */
	DELETE FROM ASRSysCrossTabAccess
	WHERE ID IN (SELECT id FROM @crossTabIDs)

	INSERT INTO ASRSysCrossTabAccess
		(ID, groupName, access)
		(SELECT ASRSysCrossTab.CrossTabID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysCrossTab
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysCrossTab.CrossTabID IN (SELECT id FROM @crossTabIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @crossTabIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 1

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (1, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 1
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Custom Reports */
	DELETE FROM ASRSysCustomReportAccess
	WHERE ID IN (SELECT id FROM @customReportsIDs)

	INSERT INTO ASRSysCustomReportAccess
		(ID, groupName, access)
		(SELECT ASRSysCustomReportsName.ID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysCustomReportsName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysCustomReportsName.ID IN (SELECT id FROM @customReportsIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @customReportsIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 2

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (2, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 2
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Data Transfers */
	DELETE FROM ASRSysDataTransferAccess
	WHERE ID IN (SELECT id FROM @dataTransferIDs)

	INSERT INTO ASRSysDataTransferAccess
		(ID, groupName, access)
		(SELECT ASRSysDataTransferName.dataTransferID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysDataTransferName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysDataTransferName.dataTransferID IN (SELECT id FROM @dataTransferIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @dataTransferIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 3

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (3, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 3
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Exports */
	DELETE FROM ASRSysExportAccess
	WHERE ID IN (SELECT id FROM @exportIDs)

	INSERT INTO ASRSysExportAccess
		(ID, groupName, access)
		(SELECT ASRSysExportName.ID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysExportName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysExportName.ID IN (SELECT id FROM @exportIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @exportIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 4

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (4, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 4
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Global Adds */
	DELETE FROM ASRSysGlobalAccess
	WHERE ID IN (SELECT id FROM @globalAddIDs)

	INSERT INTO ASRSysGlobalAccess
		(ID, groupName, access)
		(SELECT ASRSysGlobalFunctions.functionID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysGlobalFunctions
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysGlobalFunctions.functionID IN (SELECT id FROM @globalAddIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @globalAddIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 5

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (5, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 5
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Global Updates */
	DELETE FROM ASRSysGlobalAccess
	WHERE ID IN (SELECT id FROM @globalUpdateIDs)

	INSERT INTO ASRSysGlobalAccess
		(ID, groupName, access)
		(SELECT ASRSysGlobalFunctions.functionID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysGlobalFunctions
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysGlobalFunctions.functionID IN (SELECT id FROM @globalUpdateIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @globalUpdateIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 7

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (7, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 7
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Global Deletes */
	DELETE FROM ASRSysGlobalAccess
	WHERE ID IN (SELECT id FROM @globalDeleteIDs)

	INSERT INTO ASRSysGlobalAccess
		(ID, groupName, access)
		(SELECT ASRSysGlobalFunctions.functionID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysGlobalFunctions
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysGlobalFunctions.functionID IN (SELECT id FROM @globalDeleteIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @globalDeleteIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 6

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (6, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 6
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Labels */
	DELETE FROM ASRSysMailMergeAccess
	WHERE ID IN (SELECT id FROM @labelsIDs)

	INSERT INTO ASRSysMailMergeAccess
		(ID, groupName, access)
		(SELECT ASRSysMailMergeName.mailMergeID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysMailMergeName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysMailMergeName.mailMergeID IN (SELECT id FROM @labelsIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @labelsIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 18

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (18, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 18
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Mail Merges */
	DELETE FROM ASRSysMailMergeAccess
	WHERE ID IN (SELECT id FROM @mailMergeIDs)

	INSERT INTO ASRSysMailMergeAccess
		(ID, groupName, access)
		(SELECT ASRSysMailMergeName.mailMergeID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysMailMergeName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysMailMergeName.mailMergeID IN (SELECT id FROM @mailMergeIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @mailMergeIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 9

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (9, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 9
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Match Reports */
	DELETE FROM ASRSysMatchReportAccess
	WHERE ID IN (SELECT id FROM @matchReportIDs)

	INSERT INTO ASRSysMatchReportAccess
		(ID, groupName, access)
		(SELECT ASRSysMatchReportName.matchReportID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysMatchReportName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysMatchReportName.matchReportID IN (SELECT id FROM @matchReportIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @matchReportIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 14

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (14, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 14
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Record Profiles */
	DELETE FROM ASRSysRecordProfileAccess
	WHERE ID IN (SELECT id FROM @recordProfileIDs)

	INSERT INTO ASRSysRecordProfileAccess
		(ID, groupName, access)
		(SELECT ASRSysRecordProfileName.recordProfileID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysRecordProfileName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysRecordProfileName.recordProfileID IN (SELECT id FROM @recordProfileIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @recordProfileIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 20

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (20, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 20
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

	/* Succession Planning */
	DELETE FROM ASRSysMatchReportAccess
	WHERE ID IN (SELECT id FROM @successionIDs)

	INSERT INTO ASRSysMatchReportAccess
		(ID, groupName, access)
		(SELECT ASRSysMatchReportName.matchReportID, 
			sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers,
			ASRSysMatchReportName
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.uid <> 0
			AND ASRSysMatchReportName.matchReportID IN (SELECT id FROM @successionIDs))

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id
		FROM @successionIDs
	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iUtilID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @iUtilID
			AND type = 23

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (23, @iUtilID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @iUtilID
				AND type = 23
		END

		FETCH NEXT FROM check_cursor INTO @iUtilID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor         

END