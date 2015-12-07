CREATE PROCEDURE [dbo].[spASRIntValidateCrossTab] (
	@psUtilName 		varchar(255), 
	@piUtilID 			integer, 
	@piTimestamp 		integer, 
	@piBasePicklistID	integer, 
	@piBaseFilterID 	integer, 
	@piEmailGroupID 	integer,
	@piCategoryID 		integer, 
	@psHiddenGroups 	varchar(MAX), 
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT, /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								3 = definition changed by someone else, overwrite ? */
	@psDeletedFilters 	varchar(MAX)	OUTPUT,
	@psHiddenFilters 	varchar(MAX)	OUTPUT,
	@psJobIDsToHide		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp				integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,
			@sTemp					varchar(MAX),
			@sCurrentID				varchar(MAX),
			@sParameter				varchar(MAX),
			@sExprName  			varchar(MAX),
			@sBatchJobName			varchar(MAX),
			@iBatchJobID			integer,
			@iBatchJobScheduled		integer,
			@sBatchJobRoleToPrompt	varchar(MAX),
			@iNonHiddenCount		integer,
			@sBatchJobUserName		sysname,
			@sJobName				varchar(MAX),
			@sCurrentUserGroup		sysname,
			@fBatchJobsOK			bit,
			@sScheduledUserGroups	varchar(MAX),
			@sScheduledJobDetails	varchar(MAX),
			@sCurrentUserAccess		varchar(MAX),
			@iOwnedJobCount 		integer,
			@sOwnedJobDetails		varchar(MAX),
			@sOwnedJobIDs			varchar(MAX),
			@sNonOwnedJobDetails	varchar(MAX),
			@sHiddenGroupsList		varchar(MAX),
			@sHiddenGroup			varchar(MAX),
			@fSysSecMgr				bit,
			@sActualUserName		sysname,
			@iUserGroupID			integer;

	SET @fBatchJobsOK = 1;
	SET @sScheduledUserGroups = '';
	SET @sScheduledJobDetails = '';
	SET @iOwnedJobCount = 0;
	SET @sOwnedJobDetails = '';
	SET @sOwnedJobIDs = '';
	SET @sNonOwnedJobDetails = '';

	SELECT @sCurrentUser = SYSTEM_USER;
	SET @psErrorMsg = '';
	SET @piErrorCode = 0;
	SET @psDeletedFilters = '';
	SET @psHiddenFilters = '';

	exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;
	
 	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The cross tab has been deleted by another user. Save as a new definition ?';
			SET @piErrorCode = 2;
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysCrossTab
			WHERE CrossTabID = @piUtilID;

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					1, 
					@piUtilID,
					@sAccess	OUTPUT;
		
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The cross tab has been amended by another user and is now Read Only. Save as a new definition ?';
					SET @piErrorCode = 2;
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The cross tab has been amended by another user. Would you like to overwrite this definition ?';
					SET @piErrorCode = 3;
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the cross tab name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCrossTab
			WHERE name = @psUtilName
				AND CrossTabID <> @piUtilID
				AND CrossTabType = 0;
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCrossTab
			WHERE name = @psUtilName AND CrossTabType = 0;
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'A cross tab called ''' + @psUtilName + ''' already exists.';
			SET @piErrorCode = 1;
		END
	END

	IF (@piErrorCode = 0) AND (@piBasePicklistID > 0)
	BEGIN
		/* Check that the Base table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piBasePicklistID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table picklist has been deleted by another user.';
			SET @piErrorCode = 1;
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piBasePicklistID;

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table picklist has been made hidden by another user.';
				SET @piErrorCode = 1;
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseFilterID > 0)
	BEGIN
		/* Check that the Base table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piBaseFilterID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table filter has been deleted by another user.';
			SET @piErrorCode = 1;
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piBaseFilterID;

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table filter has been made hidden by another user.';
				SET @piErrorCode = 1;
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piEmailGroupID > 0)
	BEGIN
		/* Check that the email group exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysEmailGroupName 
		WHERE emailGroupID = @piEmailGroupID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The email group has been deleted by another user.';
			SET @piErrorCode = 1;
		END
	END

	IF (@piErrorCode = 0) AND (@piCategoryID > 0)
	BEGIN
		/* Check that the category exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCategories
		WHERE id = @piCategoryID And _deleted = 1;

		IF @iCount = 1
		BEGIN
			SET @psErrorMsg = 'The category has been deleted by another user.';
			SET @piErrorCode = 1;
		END
	END

	IF (@piErrorCode = 0) AND (@piUtilID > 0) AND (len(@psHiddenGroups) > 0)
	BEGIN
		SELECT @sOwner = userName
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID;

		IF (@sOwner = @sCurrentUser) 
		BEGIN
			EXEC spASRIntGetActualUserDetails
				@sActualUserName OUTPUT,
				@sCurrentUserGroup OUTPUT,
				@iUserGroupID OUTPUT;

			DECLARE @HiddenGroups TABLE(groupName sysname, groupID integer);
			SET @sHiddenGroupsList = substring(@psHiddenGroups, 2, len(@psHiddenGroups)-2);
			WHILE LEN(@sHiddenGroupsList) > 0
			BEGIN
				IF CHARINDEX(char(9), @sHiddenGroupsList) > 0
				BEGIN
					SET @sHiddenGroup = LEFT(@sHiddenGroupsList, CHARINDEX(char(9), @sHiddenGroupsList) - 1);
					SET @sHiddenGroupsList = RIGHT(@sHiddenGroupsList, LEN(@sHiddenGroupsList) - CHARINDEX(char(9), @sHiddenGroupsList));
				END
				ELSE
				BEGIN
					SET @sHiddenGroup = @sHiddenGroupsList;
					SET @sHiddenGroupsList = '';
				END

				INSERT INTO @HiddenGroups (groupName, groupID) (SELECT @sHiddenGroup, uid FROM sysusers WHERE name = @sHiddenGroup);
			END

			DECLARE batchjob_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
				ASRSysBatchJobName.Username,
				ASRSysCrossTab.Name AS 'JobName'
	 		FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID 
			INNER JOIN ASRSysCrossTab ON ASRSysCrossTab.CrossTabID = ASRSysBatchJobDetails.JobID
			LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.access <> 'HD'
				AND ASRSysBatchJobAccess.groupName IN (SELECT name FROM sysusers WHERE uid IN (SELECT groupID FROM @HiddenGroups))
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
				AND ASRSysBatchJobDetails.JobID IN (@piUtilID)
			GROUP BY ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				ASRSysBatchJobName.Username,
				ASRSysCrossTab.Name;

			OPEN batchjob_cursor;
			FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
				@iBatchJobID,
				@iBatchJobScheduled,
				@sBatchJobRoleToPrompt,
				@iNonHiddenCount,
				@sBatchJobUserName,
				@sJobName;
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @sCurrentUserAccess = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
								OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		 						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
						WHEN ASRSysBatchJobName.userName = system_user THEN 'RW'
						ELSE
							CASE
								WHEN ASRSysBatchJobAccess.access IS null THEN 'HD'
								ELSE ASRSysBatchJobAccess.access
							END
					END 
				FROM sysusers b
				INNER JOIN sysusers a ON b.uid = a.gid
				LEFT OUTER JOIN ASRSysBatchJobAccess ON (b.name = ASRSysBatchJobAccess.groupName
					AND ASRSysBatchJobAccess.id = @iBatchJobID)
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobAccess.ID = ASRSysBatchJobName.ID
				WHERE a.Name = @sActualUserName;

				IF @sBatchJobUserName = @sOwner
				BEGIN
					/* Found a Batch Job whose owner is the same. */
					IF (@iBatchJobScheduled = 1) AND
						(len(@sBatchJobRoleToPrompt) > 0) AND
						(@sBatchJobRoleToPrompt <> @sCurrentUserGroup) AND
						(CHARINDEX(char(9) + @sBatchJobRoleToPrompt + char(9), @psHiddenGroups) > 0)
					BEGIN
						/* Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0;
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sBatchJobRoleToPrompt + '<BR>';

						IF @sCurrentUserAccess = 'HD'
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>';
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>';
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0 
						BEGIN
							SET @iOwnedJobCount = @iOwnedJobCount + 1;
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains Cross Tab ' + @sJobName + ')' + '<BR>';
							SET @sOwnedJobIDs = @sOwnedJobIDs +
								CASE 
									WHEN Len(@sOwnedJobIDs) > 0 THEN ', '
									ELSE ''
								END +  convert(varchar(100), @iBatchJobID);
						END
					END
				END			
				ELSE
				BEGIN
					/* Found a Batch Job whose owner is not the same. */
					SET @fBatchJobsOK = 0;
	    
					IF @sCurrentUserAccess = 'HD'
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>';
					END
					ELSE
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>';
					END
				END

				FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
					@iBatchJobID,
					@iBatchJobScheduled,
					@sBatchJobRoleToPrompt,
					@iNonHiddenCount,
					@sBatchJobUserName,
					@sJobName;
			END

			CLOSE batchjob_cursor;
			DEALLOCATE batchjob_cursor;
		END
	END

	IF @fBatchJobsOK = 0
	BEGIN
		SET @piErrorCode = 1;

		IF Len(@sScheduledJobDetails) > 0 
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden from the following user groups :'  + '<BR><BR>' +
				@sScheduledUserGroups  +
				'<BR>as it is used in the following batch jobs which are scheduled to be run by these user groups :<BR><BR>' +
				@sScheduledJobDetails;
		END
		ELSE
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden as it is used in the following batch jobs of which you are not the owner :<BR><BR>' +
				@sNonOwnedJobDetails;
	      	END
	END
	ELSE
	BEGIN
	    	IF (@iOwnedJobCount > 0) 
		BEGIN
			SET @piErrorCode = 4;
			SET @psErrorMsg = 'Making this definition hidden to user groups will automatically make the following definition(s), of which you are the owner, hidden to the same user groups:<BR><BR>' +
				@sOwnedJobDetails + '<BR><BR>' +
				'Do you wish to continue ?';
		END
	END

	SET @psJobIDsToHide = @sOwnedJobIDs;

END