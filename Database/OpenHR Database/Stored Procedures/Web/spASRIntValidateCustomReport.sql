CREATE PROCEDURE [dbo].[spASRIntValidateCustomReport] (
	@psUtilName 				varchar(255), 
	@piUtilID 					integer, 
	@piTimestamp 				integer, 
	@piBasePicklistID			integer, 
	@piBaseFilterID 			integer, 
	@piEmailGroupID 			integer, 
	@piParent1PicklistID		integer, 
	@piParent1FilterID 			integer, 
	@piParent2PicklistID		integer, 
	@piParent2FilterID 			integer,
	/* Category to check it exists in table or not */
	@piCategoryID 				integer,
	
	@piChildFilterID 			varchar(100),			/* tab delimited string of child filter ids */ 
	@psCalculations 			varchar(MAX), 
	@psHiddenGroups 			varchar(MAX), 
	@psErrorMsg					varchar(MAX)	OUTPUT,
	@piErrorCode				varchar(MAX)	OUTPUT, /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								3 = definition changed by someone else, overwrite ? */
	@psDeletedCalcs 			varchar(MAX)	OUTPUT, 
	@psHiddenCalcs 				varchar(MAX)	OUTPUT,
	@psDeletedFilters 			varchar(MAX)	OUTPUT,
	@psHiddenFilters 			varchar(MAX)	OUTPUT,
	@psDeletedOrders			varchar(MAX)	OUTPUT,
	@psJobIDsToHide				varchar(MAX)	OUTPUT,
	@psDeletedPicklists 		varchar(MAX)	OUTPUT,
	@psHiddenPicklists 			varchar(MAX)	OUTPUT
	
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp	integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,
			@sTemp					varchar(MAX),
			@sCurrentID				varchar(100),
			@sParameter				varchar(MAX),
			@sExprName  			varchar(255),
			@sBatchJobName			varchar(255),
			@iBatchJobID			integer,
			@iBatchJobScheduled		integer,
			@sBatchJobRoleToPrompt	varchar(MAX),
			@iNonHiddenCount		integer,
			@sBatchJobUserName		sysname,
			@sJobName				varchar(255),
			@sCurrentUserGroup		sysname,
			@fBatchJobsOK			bit,
			@sScheduledUserGroups	varchar(MAX),
			@sScheduledJobDetails	varchar(MAX),
			@sCurrentUserAccess		varchar(MAX),
			@iOwnedJobCount			integer,
			@sOwnedJobDetails		varchar(MAX),
			@sOwnedJobIDs			varchar(MAX),
			@sNonOwnedJobDetails	varchar(MAX),
			@sHiddenGroupsList		varchar(MAX),
			@sHiddenGroup			varchar(MAX),
			@fSysSecMgr				bit,
			@sActualUserName		sysname,
			@iUserGroupID			integer;

	SET @fBatchJobsOK = 1
	SET @sScheduledUserGroups = ''
	SET @sScheduledJobDetails = ''
	SET @iOwnedJobCount = 0
	SET @sOwnedJobDetails = ''
	SET @sOwnedJobIDs = ''
	SET @sNonOwnedJobDetails = ''

	SELECT @sCurrentUser = SYSTEM_USER
	SET @psErrorMsg = ''
	SET @piErrorCode = 0
	SET @psDeletedCalcs = ''
	SET @psHiddenCalcs = ''
	SET @psDeletedOrders = ''
	SET @psDeletedFilters = ''
	SET @psHiddenFilters = ''
	SET @psDeletedPicklists = ''
	SET @psHiddenPicklists = ''
	--SET @psDeletedCategory = ''

	EXEC spASRIntSysSecMgr @fSysSecMgr OUTPUT
	
 	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCustomReportsName
		WHERE ID = @piUtilID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The report has been deleted by another user. Save as a new definition ?'
			SET @piErrorCode = 2
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysCustomReportsName
			WHERE ID = @piUtilID

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					2, 
					@piUtilID,
					@sAccess	OUTPUT

				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The report has been amended by another user and is now Read Only. Save as a new definition ?'
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The report has been amended by another user. Would you like to overwrite this definition ?'
					SET @piErrorCode = 3
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the report name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCustomReportsName
			WHERE name = @psUtilName
				AND ID <> @piUtilID
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCustomReportsName
			WHERE name = @psUtilName
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'A report called ''' + @psUtilName + ''' already exists.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piBasePicklistID > 0)
	BEGIN
		/* Check that the Base table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piBasePicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piBasePicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piBasePicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piBasePicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseFilterID > 0)
	BEGIN
		/* Check that the Base table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piBaseFilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piBaseFilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piBaseFilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piBaseFilterID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piEmailGroupID > 0)
	BEGIN
		/* Check that the email group exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysEmailGroupName 
		WHERE emailGroupID = @piEmailGroupID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The email group has been deleted by another user.'
			SET @piErrorCode = 1
		END
	END

	--//------------------------------------------------------------

	IF (@piErrorCode = 0) AND (@piCategoryID > 0)
	BEGIN
		/* Check that the category exists. */
		SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysCategories]
		WHERE ID = @piCategoryID And _deleted = 'True'

		IF @iCount = 1
		BEGIN
			SET @psErrorMsg = 'The category has been deleted by another user.'
			SET @piErrorCode = 1
		END
	END

	--//------------------------------------------------------------

	IF (@piErrorCode = 0) AND (@piParent1PicklistID > 0)
	BEGIN
		/* Check that the Parent1 table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piParent1PicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The first parent table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent1PicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piParent1PicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The first parent table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent1PicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent1FilterID > 0)
	BEGIN
		/* Check that the Parent 1 table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piParent1FilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The parent 1 filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent1FilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piParent1FilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The parent 1 table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent1FilterID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent2PicklistID > 0)
	BEGIN
		/* Check that the Parent1 table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piParent2PicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The second parent table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent2PicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piParent2PicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The second parent table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent2PicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent2FilterID > 0)
	BEGIN
		/* Check that the Parent 2 table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piParent2FilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The parent 2 filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent2FilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piParent2FilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The parent 2 table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent2FilterID)
			END
		END
	END

	/* Check that the selected child filters exist and are not hidden. */
	IF (@piErrorCode = 0) AND (LEN(@piChildFilterID) > 0)
	BEGIN
		SET @sTemp = @piChildFilterID

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(char(9), @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			IF @sCurrentID > 0 
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM ASRSysExpressions
				WHERE exprID = convert(integer, @sCurrentID)

				IF @iCount = 0
				BEGIN
					SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedFilters) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							 'One or more of the child filters have been deleted by another user. They will be automatically removed from the report.'
					END
					SET @psDeletedFilters = @psDeletedFilters +
					CASE
						WHEN LEN(@psDeletedFilters) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
					SET @piErrorCode = 1
			 	END
				ELSE
			  	BEGIN
					SELECT @sOwner = userName,
						@sAccess = access
					FROM ASRSysExpressions
					WHERE exprID = convert(integer, @sCurrentID)

					IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @psErrorMsg = 
							@psErrorMsg + 
							CASE
								WHEN LEN(@psHiddenFilters) > 0 THEN ''
								ELSE 
									CASE 
										WHEN LEN(@psErrorMsg) > 0 THEN char(13)
										ELSE ''
									END +
									'One or more of the child filters have been made hidden by another user. They will be automatically removed from the report.'
							END
						SET @psHiddenFilters = @psHiddenFilters +
						CASE
							WHEN LEN(@psHiddenFilters) > 0 THEN ','
							ELSE ''
						END + @sCurrentID
						
						SET @piErrorCode = 1
					END
			  	END
			END
		END
	END

	/* Check that the selected child filters exist and are not hidden. */
	IF (@piErrorCode = 0) AND (LEN(@psDeletedOrders) > 0)
	BEGIN
		SET @sTemp = @psDeletedOrders

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(char(9), @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			IF @sCurrentID > 0 
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM ASRSysOrders
				WHERE OrderID = convert(integer, @sCurrentID)

				IF @iCount = 0
				BEGIN
					SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedOrders) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							 'One or more of the child orders have been deleted by another user. They will be automatically removed from the report.'
					END
					SET @psDeletedOrders = @psDeletedOrders +
					CASE
						WHEN LEN(@psDeletedOrders) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
					SET @piErrorCode = 1
			 	END
			END
		END
	END
	
	/* Check that the selected runtime calculations exists. */
	IF (@piErrorCode = 0) AND (LEN(@psCalculations) > 0)
	BEGIN
		SET @sTemp = @psCalculations

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(',', @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(',', @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(',', @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			SELECT @iCount = COUNT(*)
			FROM ASRSysExpressions
			 WHERE exprID = convert(integer, @sCurrentID)

			IF @iCount = 0
			BEGIN
				SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							'One or more runtime calculations have been deleted by another user. They will be automatically removed from the report.'
					END
				SET @psDeletedCalcs = @psDeletedCalcs +
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
				SET @piErrorCode = 1
			END
			ELSE
			BEGIN
				SELECT @sOwner = userName,
					@sAccess = access
				FROM ASRSysExpressions
				WHERE exprID = convert(integer, @sCurrentID)

				IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @psErrorMsg = 
						@psErrorMsg + 
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ''
							ELSE 
								CASE 
									WHEN LEN(@psErrorMsg) > 0 THEN char(13)
									ELSE ''
								END +
								'One or more runtime calculations have been made hidden by another user. They will be automatically removed from the report.'
						END
					SET @psHiddenCalcs = @psHiddenCalcs +
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ','
							ELSE ''
						END + @sCurrentID
						
					SET @piErrorCode = 1
				END
			END
		END
	END
	
	IF (@piErrorCode = 0) AND (@piUtilID > 0) AND (len(@psHiddenGroups) > 0)
	BEGIN
		SELECT @sOwner = userName
		FROM ASRSysCustomReportsName
		WHERE ID = @piUtilID

		IF (@sOwner = @sCurrentUser) 
		BEGIN
			EXEC spASRIntGetActualUserDetails
				@sActualUserName OUTPUT,
				@sCurrentUserGroup OUTPUT,
				@iUserGroupID OUTPUT

			DECLARE @HiddenGroups TABLE(groupName sysname, groupID integer)
			SET @sHiddenGroupsList = substring(@psHiddenGroups, 2, len(@psHiddenGroups)-2)
			WHILE LEN(@sHiddenGroupsList) > 0
			BEGIN
				IF CHARINDEX(char(9), @sHiddenGroupsList) > 0
				BEGIN
					SET @sHiddenGroup = LEFT(@sHiddenGroupsList, CHARINDEX(char(9), @sHiddenGroupsList) - 1)
					SET @sHiddenGroupsList = RIGHT(@sHiddenGroupsList, LEN(@sHiddenGroupsList) - CHARINDEX(char(9), @sHiddenGroupsList))
				END
				ELSE
				BEGIN
					SET @sHiddenGroup = @sHiddenGroupsList
					SET @sHiddenGroupsList = ''
				END

				INSERT INTO @HiddenGroups (groupName, groupID) (SELECT @sHiddenGroup, uid FROM sysusers WHERE name = @sHiddenGroup)
			END

			DECLARE batchjob_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
				ASRSysBatchJobName.Username,
				ASRSysCustomReportsName.Name AS 'JobName'
	 		FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID 
			INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsName.ID = ASRSysBatchJobDetails.JobID
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
			WHERE ASRSysBatchJobDetails.JobType = 'Custom Report'
				AND ASRSysBatchJobDetails.JobID IN (@piUtilID)
			GROUP BY ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				ASRSysBatchJobName.Username,
				ASRSysCustomReportsName.Name

			OPEN batchjob_cursor
			FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
				@iBatchJobID,
				@iBatchJobScheduled,
				@sBatchJobRoleToPrompt,
				@iNonHiddenCount,
				@sBatchJobUserName,
				@sJobName	
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
				WHERE a.Name = @sActualUserName

				IF @sBatchJobUserName = @sOwner
				BEGIN
					/* Found a Batch Job whose owner is the same. */
					IF (@iBatchJobScheduled = 1) AND
						(len(@sBatchJobRoleToPrompt) > 0) AND
						(@sBatchJobRoleToPrompt <> @sCurrentUserGroup) AND
						(CHARINDEX(char(9) + @sBatchJobRoleToPrompt + char(9), @psHiddenGroups) > 0)
					BEGIN
						/* Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sBatchJobRoleToPrompt + '<BR>'

						IF @sCurrentUserAccess = 'HD'
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0 
						BEGIN
							SET @iOwnedJobCount = @iOwnedJobCount + 1
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains Custom Report ' + @sJobName + ')' + '<BR>'
							SET @sOwnedJobIDs = @sOwnedJobIDs +
								CASE 
									WHEN Len(@sOwnedJobIDs) > 0 THEN ', '
									ELSE ''
								END +  convert(varchar(100), @iBatchJobID)
						END
					END
				END			
				ELSE
				BEGIN
					/* Found a Batch Job whose owner is not the same. */
					SET @fBatchJobsOK = 0
	    
					IF @sCurrentUserAccess = 'HD'
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
					END
				END

				FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
					@iBatchJobID,
					@iBatchJobScheduled,
					@sBatchJobRoleToPrompt,
					@iNonHiddenCount,
					@sBatchJobUserName,
					@sJobName	
			END
			CLOSE batchjob_cursor
			DEALLOCATE batchjob_cursor	

		END
	END

	IF @fBatchJobsOK = 0
	BEGIN
		SET @piErrorCode = 1

		IF Len(@sScheduledJobDetails) > 0 
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden from the following user groups :'  + '<BR><BR>' +
				@sScheduledUserGroups  +
				'<BR>as it is used in the following batch jobs which are scheduled to be run by these user groups :<BR><BR>' +
				@sScheduledJobDetails
		END
		ELSE
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden as it is used in the following batch jobs of which you are not the owner :<BR><BR>' +
				@sNonOwnedJobDetails
	      	END
	END
	ELSE
	BEGIN
	    	IF (@iOwnedJobCount > 0) 
		BEGIN
			SET @piErrorCode = 4
			SET @psErrorMsg = 'Making this definition hidden to user groups will automatically make the following definition(s), of which you are the owner, hidden to the same user groups:<BR><BR>' +
				@sOwnedJobDetails + '<BR><BR>' +
				'Do you wish to continue ?'
		END
	END

	SET @psJobIDsToHide = @sOwnedJobIDs
	
END

GO

