CREATE PROCEDURE [dbo].[sp_ASRIntCheckCanMakeHidden] (
	@piUtilityType		integer,
	@piUtilityID		integer,
	@piResult			integer			OUTPUT,
	@psMessage			varchar(MAX)	OUTPUT
) AS
BEGIN

	SET NOCOUNT ON;

	/* Check if the given picklist/filter/calculation can be made hidden.
	Return 	0 if there's no problem
		1 if it is used only in utilities owned by the current user - we then need to prompt the user if they want to make these utilities hidden too.
		2 if it is used in utilities which are in batch jobs not owned by the current user - Cannot therefore make the utility hidden. 
		3 if it is used in utilities which are not owned by the current user - Cannot therefore make the utility hidden. */
	DECLARE
		@sCurrentUser				sysname,
		@sUtilName					varchar(255),
		@iUtilID					integer,
		@sUtilOwner					varchar(255),
		@sUtilAccess				varchar(MAX),
		@iCount_Owner				integer,
		@sDetails_Owner				varchar(MAX),
		@iCount_NotOwner			integer,
		@sDetails_NotOwner			varchar(MAX),
		@iCount						integer,
		@sJobName					varchar(MAX),
		@sBatchJobDetails_Owner		varchar(255),
		@fBatchJobsOK				bit,
		@sBatchJobDetails_NotOwner	varchar(MAX),
		@iNonHiddenCount			integer,
		@iScheduled					integer, 
		@sRoleToPrompt				sysname,
		@sCurrentUserGroup			sysname,
		@sScheduledUserGroups		varchar(MAX),
		@sScheduledJobDetails		varchar(MAX),
		@superCursor				cursor,
		@iTemp						integer,
		@fSysSecMgr					bit,
		@sActualUserName			sysname,
		@iUserGroupID				integer;

	SET @sCurrentUser = SYSTEM_USER;
	SET @iCount_Owner = 0;
	SET @sDetails_Owner = '';
	SET @iCount_NotOwner = 0;
	SET @sDetails_NotOwner = '';
	SET @sBatchJobDetails_Owner = '';
	SET @sBatchJobDetails_NotOwner = '';
	SET @fBatchJobsOK = 1;
	SET @psMessage = '';
	SET @piResult = 0;
	SET @sScheduledUserGroups = '';
	SET @sScheduledJobDetails = '';

	EXEC spASRIntSysSecMgr @fSysSecMgr OUTPUT;
	
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sCurrentUserGroup OUTPUT,
		@iUserGroupID OUTPUT;

	DECLARE @batchJobIDs TABLE(id integer)
	DECLARE @calendarReportsIDs TABLE(id integer)
	DECLARE @careerIDs TABLE(id integer)
		DECLARE @crossTabIDs TABLE(id integer)
	DECLARE @customReportsIDs TABLE(id integer)
	DECLARE @dataTransferIDs TABLE(id integer)
	DECLARE @exportIDs TABLE(id integer)
	DECLARE @globalAddIDs TABLE(id integer)
		DECLARE @globalUpdateIDs TABLE(id integer)
		DECLARE @globalDeleteIDs TABLE(id integer)
	DECLARE @labelsIDs TABLE(id integer)
		DECLARE @mailMergeIDs TABLE(id integer)
	DECLARE @matchReportIDs TABLE(id integer)
	DECLARE @recordProfileIDs TABLE(id integer)
	DECLARE @successionIDs TABLE(id integer)
	DECLARE @filterIDs TABLE(id integer)
	DECLARE @calculationIDs TABLE(id integer)
	DECLARE @expressionIDs TABLE(id integer)
	DECLARE @superExpressionIDs TABLE(id integer)
	DECLARE @talentReportIDs TABLE(id integer)

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
				ASRSysExpressions.ExprID AS [ID],
				ASRSysExpressions.Username,
				ASRSysExpressions.Access
			FROM ASRSysExpressions
			WHERE ASRSysExpressions.ExprID IN (SELECT id FROM @superExpressionIDs)
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Calculation : ' + @sUtilName + '<BR>'
					INSERT INTO @calculationIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Calculation whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calculation : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calculation : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Filter : ' + @sUtilName + '<BR>'
					INSERT INTO @filterIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Filter whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Filter : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Filter : ' + @sUtilName + '<BR>'
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
		SELECT ASRSysCalendarReports.Name,
			ASRSysCalendarReports.ID,
			ASRSysCalendarReports.Username,
			COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysCalendarReports
		LEFT OUTER JOIN ASRSYSCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSYSCalendarReportEvents.calendarReportID
		LEFT OUTER JOIN ASRSYSCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID
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
		WHERE ASRSysCalendarReports.DescriptionExpr IN (SELECT id FROM @expressionIDs)
			OR ASRSysCalendarReports.StartDateExpr IN (SELECT id FROM @expressionIDs)
			OR ASRSysCalendarReports.EndDateExpr IN (SELECT id FROM @expressionIDs)
			OR ASRSysCalendarReports.Filter IN (SELECT id FROM @expressionIDs)
			OR ASRSYSCalendarReportEvents.FilterID IN (SELECT id FROM @expressionIDs)
		GROUP BY ASRSysCalendarReports.Name,
			ASRSysCalendarReports.ID,
			ASRSysCalendarReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Calendar Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Calendar Report : ' + @sUtilName + '<BR>'
					INSERT INTO @calendarReportsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Calendar Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					17,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calendar Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calendar Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Calendar Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Calendar Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
            		
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
					END
				END            
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Talent Report for this Expression. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysTalentReports.Name,
			ASRSysTalentReports.ID,
			ASRSysTalentReports.Username,
			COUNT (ASRSysTalentReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysTalentReports
		LEFT OUTER JOIN ASRSysTalentReportAccess ON ASRSysTalentReports.ID = ASRSysTalentReportAccess.ID
			AND ASRSysTalentReportAccess.access <> 'HD'
			AND ASRSysTalentReportAccess.groupName NOT IN (SELECT sysusers.name
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
    WHERE (ASRSysTalentReports.BaseFilterID IN (SELECT id FROM @expressionIDs)
      OR ASRSysTalentReports.MatchFilterID IN (SELECT id FROM @expressionIDs))
		GROUP BY ASRSysTalentReports.Name,
			ASRSysTalentReports.ID,
 			ASRSysTalentReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Talent Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Talent Report : ' + @sUtilName + '<BR>'
					INSERT INTO @talentReportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Talent Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
          
				exec spASRIntCurrentUserAccess 
					38,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Talent Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Talent Report : ' + @sUtilName + '<BR>'
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Talent Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @talentReportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysTalentReports.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysTalentReports ON ASRSysTalentReports.ID = ASRSysBatchJobdetails.JobID
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
				WHERE ASRSysBatchJobDetails.JobType = 'Talent Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @talentReportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysTalentReports.Name

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
					/* Found a Talent Report in a batch job whose owner is the same */
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
      
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Talent Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Talent Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
            		
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Career Progression : ' + @sUtilName + '<BR>'
					INSERT INTO @careerIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Career Progression whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					24,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Career Progression : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Career Progression : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Career Progression ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Career Progression in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
		SELECT ASRSysCrossTab.Name,
			ASRSysCrossTab.[CrossTabID] AS [ID],
			ASRSysCrossTab.Username,
			COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]
		FROM ASRSysCrossTab
		LEFT OUTER JOIN ASRSYSCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID
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
		WHERE ASRSysCrossTab.FilterID IN (SELECT id FROM @expressionIDs)
		GROUP BY ASRSysCrossTab.Name,
			ASRSysCrossTab.crossTabID,
			ASRSysCrossTab.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Cross Tab whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Cross Tab : ' + @sUtilName + '<BR>'
					INSERT INTO @crossTabIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Cross Tab whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					1,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Cross Tab : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Cross Tab : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
					
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Cross Tab ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Cross Tab in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			 LEFT OUTER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = ASRSysCustomReportsDetails.CustomReportID
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
				OR(ASRSysCustomReportsDetails.Type = 'E' 
					AND ASRSysCustomReportsDetails.ColExprID IN (SELECT id FROM @expressionIDs))
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Custom Report : ' + @sUtilName + '<BR>'
					INSERT INTO @customReportsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Custom Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					2,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Custom Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Custom Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsname.ID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Custom Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Custom Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Data Transfer : ' + @sUtilName + '<BR>'
					INSERT INTO @dataTransferIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Data Transfer whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					3,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Data Transfer : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Data Transfer : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysDataTransferName ON ASRSysDataTransferName.DataTransferID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Data Transfer ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Data Transfer in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			SELECT ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID AS [ID],
				ASRSysMailMergeName.Username,
				COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
			 FROM ASRSysMailMergeName
			 LEFT OUTER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
			 LEFT OUTER JOIN ASRSYSMailMergeAccess ON ASRSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
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
			WHERE ASRSysMailMergeName.isLabel = 1
				AND ((ASRSysMailMergeName.FilterID IN (SELECT id FROM @expressionIDs))
				OR (ASRSysMailMergeColumns.Type = 'E' 
					AND ASRSysMailMergeColumns.ColumnID IN (SELECT id FROM @expressionIDs)))
			GROUP BY ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID,
				ASRSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Envelopes & Labels whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Envelopes & Labels : ' + @sUtilName + '<BR>'
					INSERT INTO @labelsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Envelopes & Labels whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					18,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Envelopes & Labels : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Envelopes & Labels : ' + @sUtilName + '<BR>'
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
					ASRSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeName.MailMergeID = ASRSysBatchJobdetails.JobID
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
					ASRSysMailMergeName.Name

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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Envelopes & Labels ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Envelopes & Labels in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			 LEFT OUTER JOIN ASRSysExportDetails ON ASRSysExportName.ID = ASRSysExportDetails.exportID
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
			WHERE ASRSysExportName.Filter IN (SELECT id FROM @expressionIDs)
				OR ASRSysExportName.Parent1Filter IN (SELECT id FROM @expressionIDs)
				OR ASRSysExportName.Parent2Filter IN (SELECT id FROM @expressionIDs)
				OR ASRSysExportName.ChildFilter IN (SELECT id FROM @expressionIDs)
				OR (ASRSysExportDetails.Type = 'X' 
					AND ASRSysExportDetails.ColExprID IN (SELECT id FROM @expressionIDs))
			GROUP BY ASRSysExportName.Name,
				ASRSysExportName.ID,
				ASRSysExportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Export whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Export : ' + @sUtilName + '<BR>'
					INSERT INTO @exportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Export whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					4,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Export : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Export : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysExportName ON ASRSysExportName.ID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Export ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Export in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			 LEFT OUTER JOIN ASRSysGlobalItems ON ASRSysGlobalFunctions.functionID = ASRSysGlobalItems.FunctionID
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
				AND ((ASRSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs))
				OR (ASRSysGlobalItems.ValueType = 4 
					AND ASRSysGlobalItems.ExprID IN (SELECT id FROM @expressionIDs)))
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Add : ' + @sUtilName + '<BR>'
					INSERT INTO @globalAddIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Add whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					5,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Add : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Add : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Add ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Add in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			 LEFT OUTER JOIN ASRSysGlobalItems ON ASRSysGlobalFunctions.functionID = ASRSysGlobalItems.FunctionID
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
				AND ((ASRSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs))
				OR (ASRSysGlobalItems.ValueType = 4 
					AND ASRSysGlobalItems.ExprID IN (SELECT id FROM @expressionIDs)))
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Update : ' + @sUtilName + '<BR>'
					INSERT INTO @globalUpdateIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Update whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					7,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Update : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Update : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Update ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Update in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			WHERE ASRSysGlobalFunctions.Type = 'D' 
				AND ASRSysGlobalFunctions.FilterID IN (SELECT id FROM @expressionIDs)
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Delete : ' + @sUtilName + '<BR>'
					INSERT INTO @globalDeleteIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Delete whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					6,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Delete : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Delete : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Delete ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Delete in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			SELECT ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID AS [ID],
				ASRSysMailMergeName.Username,
				COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
			 FROM ASRSysMailMergeName
			 LEFT OUTER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
			 LEFT OUTER JOIN ASRSYSMailMergeAccess ON ASRSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
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
			WHERE ASRSysMailMergeName.isLabel = 0
				AND ((ASRSysMailMergeName.FilterID IN (SELECT id FROM @expressionIDs))
				OR (ASRSysMailMergeColumns.Type = 'E' 
					AND ASRSysMailMergeColumns.ColumnID IN (SELECT id FROM @expressionIDs)))
			GROUP BY ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID,
				ASRSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Mail Merge whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Mail Merge : ' + @sUtilName + '<BR>'
					INSERT INTO @mailMergeIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Mail Merge whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					9,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Mail Merge : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Mail Merge : ' + @sUtilName + '<BR>'
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
					ASRSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeName.MailMergeID = ASRSysBatchJobdetails.JobID
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
					ASRSysMailMergeName.Name

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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Mail Merge ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Mail Merge in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Match Report : ' + @sUtilName + '<BR>'
					INSERT INTO @matchReportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Match Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					14,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Match Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Match Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Match Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Match Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Record Profile : ' + @sUtilName + '<BR>'
					INSERT INTO @recordProfileIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Record Profile whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					20,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Record Profile : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Record Profile : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Record Profile ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Record Profile in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Succession Planning : ' + @sUtilName + '<BR>'
					INSERT INTO @successionIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Succession Planning whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					23,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Succession Planning : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Succession Planning : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Succession Planning ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Succession Planning in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
					END
				END            
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END
		
		/*--------------------------------------------------------------------------------------------------------------------------------------------------------------*/
		/* Ok, all relevant utility definitions have now been checked, so check the counts and act accordingly */
		/*--------------------------------------------------------------------------------------------------------------------------------------------------------------*/
		IF (@iCount_Owner = 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 1) AND
			(len(@sBatchJobDetails_Owner) = 0)
		BEGIN
			SET @piResult = 0
			RETURN
		END
			
		IF (@iCount_Owner > 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 1)
		BEGIN
			/* Can change utils and no utils are contained within batch jobs that cant be changed. */
			SET @psMessage = @sDetails_Owner + @sBatchJobDetails_Owner
			SET @piResult = 1
			RETURN
		END
				
		IF (@iCount_Owner > 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 0)
		BEGIN
			IF Len(@sScheduledUserGroups) > 0 
			BEGIN
				SET @psMessage = @sScheduledJobDetails
				SET @piResult = 4
			END
			ELSE
			BEGIN
				/* Can change utils but abort cos those utils are in batch jobs which cannot be changed. */
				SET @psMessage = @sBatchJobDetails_NotOwner
				SET @piResult = 2
			END
			
			RETURN
		END

		IF @iCount_NotOwner > 0 
		BEGIN
			/* Cannot change utils */
			SET @psMessage = @sDetails_NotOwner
			SET @piResult = 3
			RETURN
		END
	END

	IF @piUtilityType = 10
	BEGIN
		/* Picklist */
		
		/*---------------------------------------------------*/
		/* Check Calendar Reports for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysCalendarReports.Name,
			ASRSysCalendarReports.ID,
			ASRSysCalendarReports.Username,
			COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysCalendarReports
		LEFT OUTER JOIN ASRSYSCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID
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
		WHERE ASRSysCalendarReports.Picklist = @piUtilityID
		GROUP BY ASRSysCalendarReports.Name,
			ASRSysCalendarReports.ID,
			ASRSysCalendarReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Calendar Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Calendar Report : ' + @sUtilName + '<BR>'
					INSERT INTO @calendarReportsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Calendar Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					17,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calendar Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Calendar Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Calendar Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Calendar Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
            		
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
					END
				END            
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*---------------------------------------------------*/
		/* Check Talent Reports for this Picklist. */
		/*---------------------------------------------------*/
		DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysTalentReports.Name,
			ASRSysTalentReports.ID,
			ASRSysTalentReports.Username,
			COUNT (ASRSysTalentReportAccess.Access) AS [nonHiddenCount]
		FROM ASRSysTalentReports
		LEFT OUTER JOIN ASRSysTalentReportAccess ON ASRSysTalentReports.ID = ASRSysTalentReportAccess.ID
			AND ASRSysTalentReportAccess.access <> 'HD'
			AND ASRSysTalentReportAccess.groupName NOT IN (SELECT sysusers.name
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
    WHERE (ASRSysTalentReports.BasePicklistID = @piUtilityID
      OR ASRSysTalentReports.MatchPicklistID = @piUtilityID)
		GROUP BY ASRSysTalentReports.Name,
			ASRSysTalentReports.ID,
 			ASRSysTalentReports.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Talent Report whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Talent Report : ' + @sUtilName + '<BR>'
					INSERT INTO @talentReportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Talent Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
          
				exec spASRIntCurrentUserAccess 
					38,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Talent Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Talent Report : ' + @sUtilName + '<BR>'
				END
			END

			FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		END
		CLOSE check_cursor
		DEALLOCATE check_cursor

		/* Now check that any of these Talent Reports are contained within a batch job */
		SELECT @iCount = COUNT(*)
		FROM @talentReportIDs

		IF @iCount > 0 
		BEGIN
			DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysBatchJobName.[Name],
					ASRSysBatchJobName.[ID],
					convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled],
					ASRSysBatchJobName.roleToPrompt,
					COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
					ASRSysBatchJobName.[Username],
					ASRSysTalentReports.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON AsrSysBatchJobName.ID = AsrSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysTalentReports ON ASRSysTalentReports.ID = ASRSysBatchJobdetails.JobID
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
				WHERE ASRSysBatchJobDetails.JobType = 'Talent Report'
					AND ASRSysBatchJobDetails.JobID IN (SELECT id FROM @talentReportIDs)
				GROUP BY ASRSysBatchJobName.Name,
					ASRSysBatchJobName.ID,
					convert(integer, ASRSysBatchJobName.scheduled),
					ASRSysBatchJobName.roleToPrompt,
					ASRSysBatchJobName.Username,
					ASRSysTalentReports.Name

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
					/* Found a Talent Report in a batch job whose owner is the same */
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
      
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Talent Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Talent Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
            		
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Career Progression : ' + @sUtilName + '<BR>'
					INSERT INTO @careerIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Career Progression whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					24,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Career Progression : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Career Progression : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Career Progression ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Career Progression in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
		SELECT ASRSysCrossTab.Name,
			ASRSysCrossTab.[CrossTabID] AS [ID],
			ASRSysCrossTab.Username,
			COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]
		FROM ASRSysCrossTab
		LEFT OUTER JOIN ASRSYSCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID
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
		WHERE ASRSysCrossTab.PicklistID = @piUtilityID
		GROUP BY ASRSysCrossTab.Name,
			ASRSysCrossTab.crossTabID,
			ASRSysCrossTab.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Cross Tab whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Cross Tab : ' + @sUtilName + '<BR>'
					INSERT INTO @crossTabIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Cross Tab whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					1,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Cross Tab : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Cross Tab : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
					
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Cross Tab ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Cross Tab in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Custom Report : ' + @sUtilName + '<BR>'
					INSERT INTO @customReportsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Custom Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					2,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Custom Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Custom Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsname.ID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Custom Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Custom Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Data Transfer : ' + @sUtilName + '<BR>'
					INSERT INTO @dataTransferIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Data Transfer whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					3,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Data Transfer : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Data Transfer : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysDataTransferName ON ASRSysDataTransferName.DataTransferID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Data Transfer ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Data Transfer in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			SELECT ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID AS [ID],
				ASRSysMailMergeName.Username,
				COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
			 FROM ASRSysMailMergeName
			 LEFT OUTER JOIN ASRSYSMailMergeAccess ON ASRSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
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
			WHERE ASRSysMailMergeName.isLabel = 1
				AND ASRSysMailMergeName.PicklistID = @piUtilityID
			GROUP BY ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID,
				ASRSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Envelopes & Labels whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Envelopes & Labels : ' + @sUtilName + '<BR>'
					INSERT INTO @labelsIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Envelopes & Labels whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					18,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Envelopes & Labels : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Envelopes & Labels : ' + @sUtilName + '<BR>'
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
					ASRSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeName.MailMergeID = ASRSysBatchJobdetails.JobID
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
					ASRSysMailMergeName.Name

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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Envelopes & Labels ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Envelopes & Labels in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			GROUP BY ASRSysExportName.Name,
				ASRSysExportName.ID,
				ASRSysExportName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Export whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Export : ' + @sUtilName + '<BR>'
					INSERT INTO @exportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Export whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					4,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Export : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Export : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysExportName ON ASRSysExportName.ID = ASRSysBatchJobdetails.JobID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Export ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Export in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Add : ' + @sUtilName + '<BR>'
					INSERT INTO @globalAddIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Add whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					5,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Add : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Add : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Add ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Add in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Update : ' + @sUtilName + '<BR>'
					INSERT INTO @globalUpdateIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Update whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					7,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Update : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Update : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Update ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Update in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Global Delete : ' + @sUtilName + '<BR>'
					INSERT INTO @globalDeleteIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Global Delete whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					6,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Delete : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Global Delete : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Global Delete ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Global Delete in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
			SELECT ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID AS [ID],
				ASRSysMailMergeName.Username,
				COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]
			 FROM ASRSysMailMergeName
			 LEFT OUTER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
			 LEFT OUTER JOIN ASRSYSMailMergeAccess ON ASRSysMailMergeName.MailMergeID = ASRSYSMailMergeAccess.ID
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
			WHERE ASRSysMailMergeName.isLabel = 0
				AND ASRSysMailMergeName.PicklistID = @piUtilityID
			GROUP BY ASRSysMailMergeName.Name,
				ASRSysMailMergeName.MailMergeID,
				ASRSysMailMergeName.Username

		OPEN check_cursor
		FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @sUtilOwner, @iNonHiddenCount
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sUtilOwner = @sCurrentUser
			BEGIN
				/* Found a Mail Merge whose owner is the same */
				IF @iNonHiddenCount > 0
				BEGIN
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Mail Merge : ' + @sUtilName + '<BR>'
					INSERT INTO @mailMergeIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Mail Merge whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					9,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Mail Merge : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Mail Merge : ' + @sUtilName + '<BR>'
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
					ASRSysMailMergeName.[Name] AS 'JobName' 
				FROM ASRSysBatchJobDetails
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
				INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeName.MailMergeID = ASRSysBatchJobdetails.JobID
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
					ASRSysMailMergeName.Name

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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Mail Merge ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Mail Merge in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Match Report : ' + @sUtilName + '<BR>'
					INSERT INTO @matchReportIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Match Report whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					14,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Match Report : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Match Report : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Match Report ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Match Report in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Record Profile : ' + @sUtilName + '<BR>'
					INSERT INTO @recordProfileIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Record Profile whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					20,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Record Profile : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Record Profile : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Record Profile ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Record Profile in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
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
					SET @iCount_Owner = @iCount_Owner + 1
					SET @sDetails_Owner = @sDetails_Owner + 'Succession Planning : ' + @sUtilName + '<BR>'
					INSERT INTO @successionIDs (id) VALUES (@iUtilID)
				END
			END
			ELSE
			BEGIN
				/* Found a Succession Planning whose owner is not the same */        
				SET @iCount_NotOwner = @iCount_NotOwner + 1
					
				exec spASRIntCurrentUserAccess 
					23,
					@iUtilID,
					@sUtilAccess	OUTPUT

				IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Succession Planning : <Hidden> by ' + @sUtilOwner + '<BR>'
				END
				ELSE
				BEGIN
					SET @sDetails_NotOwner = @sDetails_NotOwner + 'Succession Planning : ' + @sUtilName + '<BR>'
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
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID
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
					IF (@iScheduled = 1) 
						AND (Len(@sRoleToPrompt) > 0) 
						AND (@sRoleToPrompt <> @sCurrentUserGroup)
					BEGIN
						/*Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
			
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sRoleToPrompt + '<BR>'
						IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sUtilOwner + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sUtilName+ '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0
						BEGIN
							SET @sBatchJobDetails_Owner = @sBatchJobDetails_Owner + 'Batch Job : ' +  @sUtilName + ' (Contains Succession Planning ''' + @sJobName + ''') ' + '<BR>'
							INSERT INTO @batchJobIDs (id) VALUES(@iUtilID)
						END
					END
				END
				ELSE
				BEGIN
					/* Found a Succession Planning in a batch job whose owner is not the same */
					SET @fBatchJobsOK = 0
								
					IF (@sUtilAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : <Hidden> by ' + @sUtilName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sBatchJobDetails_NotOwner = @sBatchJobDetails_NotOwner + 'Batch Job : ' + @sUtilName + '<BR>'
					END
				END            
	
				FETCH NEXT FROM check_cursor INTO @sUtilName, @iUtilID, @iScheduled, @sRoleToPrompt, @iNonHiddenCount, @sUtilOwner, @sJobName
			END
			CLOSE check_cursor
			DEALLOCATE check_cursor
		END

		/*--------------------------------------------------------------------------------------------------------------------------------------------------------------*/
		/* Ok, all relevant utility definitions have now been checked, so check the counts and act accordingly */
		/*--------------------------------------------------------------------------------------------------------------------------------------------------------------*/
		IF (@iCount_Owner = 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 1) AND
			(len(@sBatchJobDetails_Owner) = 0)
		BEGIN
			SET @piResult = 0
								RETURN
		END
					 
		IF (@iCount_Owner > 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 1)
		BEGIN
			/* Can change utils and no utils are contained within batch jobs that cant be changed. */
			SET @psMessage = @sDetails_Owner + @sBatchJobDetails_Owner
			SET @piResult = 1
			RETURN
		END
				
		IF (@iCount_Owner > 0) AND
			(@iCount_NotOwner = 0) AND
			(@fBatchJobsOK = 0)
		BEGIN
			IF Len(@sScheduledUserGroups) > 0 
			BEGIN
				SET @psMessage = @sScheduledJobDetails
				SET @piResult = 4
			END
			ELSE
			BEGIN
				/* Can change utils but abort cos those utils are in batch jobs which cannot be changed. */
				SET @psMessage = @sBatchJobDetails_NotOwner
				SET @piResult = 2
			END
			
			RETURN
		END

		IF @iCount_NotOwner > 0 
		BEGIN
			/* Cannot change utils */
			SET @psMessage = @sDetails_NotOwner
			SET @piResult = 3
			RETURN
		END
	END
END