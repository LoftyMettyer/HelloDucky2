CREATE PROCEDURE [dbo].[spASRIntSaveTalentReport] (
	@psName								varchar(255),
	@psDescription				varchar(MAX),
	@piBaseTableID				integer,
	@piBaseSelection			integer,
	@piBasePicklistID			integer,
	@piBaseFilterID				integer,
	@piBaseChildTableID		integer,
	@piBaseChildColumnID	integer,
	@piBaseMinimumRatingColumnID		integer,
	@piBasePreferredRatingColumnID	integer,
	@piMatchTableID				integer,
	@piMatchSelection			integer,
	@piMatchPicklistID		integer,
	@piMatchFilterID			integer,
	@piMatchChildTableID	integer,
	@piMatchChildColumnID	integer,
	@piMatchChildRatingColumnID		integer,
	@piMatchAgainstType		integer,
	@psUserName						varchar(255),
	@psAccess							varchar(MAX),
	@psJobsToHide					varchar(MAX),
	@psJobsToHideGroups		varchar(MAX),
	@psColumns						varchar(MAX),
	@piID									integer					OUTPUT,
	@piCategoryID					integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sTemp					varchar(MAX),
			@sColumnDefn			varchar(MAX),
			@sColumnParam			varchar(MAX),
			@iSequence				integer,
			@sType					varchar(MAX),
			@iColExprID				integer,
			@sHeading				varchar(MAX),
			@iSize					integer,
			@iDP					integer,
			@fIsNumeric				bit,
			@iSortOrderSequence		integer,
			@sSortOrder				varchar(MAX),
			@iCount					integer,
			@fIsNew					bit,
			@sGroup					varchar(255),
			@sAccess				varchar(MAX),
			@sSQL					nvarchar(MAX);

	DECLARE	@outputTable table (id int NOT NULL);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''');
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''');

	SET @fIsNew = 0;

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysTalentReports(
			Name, 
			[Description], 
			BaseTableID, 
			BaseSelection, 
			BasePicklistID, 
			BaseFilterID, 
			BaseChildTableID,
			BaseChildColumnID,
			BaseMinimumRatingColumnID,
			BasePreferredRatingColumnID,
			MatchTableID,
			MatchSelection,
			MatchPicklistID,
			MatchFilterID,
			MatchChildTableID,
			MatchChildColumnID,
			MatchChildRatingColumnID,
			MatchAgainstType,
 			UserName)
		OUTPUT inserted.ID INTO @outputTable
 		VALUES (
 			@psName,
 			@psDescription,
 			@piBaseTableID,
			@piBaseSelection, 
			@piBasePicklistID, 
			@piBaseFilterID, 
			@piBaseChildTableID,
			@piBaseChildColumnID,
			@piBaseMinimumRatingColumnID,
			@piBasePreferredRatingColumnID,
			@piMatchTableID,
			@piMatchSelection,
			@piMatchPicklistID,
			@piMatchFilterID,
			@piMatchChildTableID,
			@piMatchChildColumnID,
			@piMatchChildRatingColumnID,
			@piMatchAgainstType,
 			@psUserName);

		SET @fIsNew = 1;
		-- Get the ID of the inserted record.
		SELECT @piID = id FROM @outputTable;

		EXEC [dbo].[spsys_saveobjectcategories] 38, @piID, @piCategoryID;

	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysTalentReports SET 
			Name = @psName,
			Description = @psDescription,
			BaseTableID = @piBaseTableID, 
			BaseSelection = @piBaseSelection, 
			BasePicklistID = @piBasePicklistID, 
			BaseFilterID = @piBaseFilterID, 
			BaseChildTableID = @piBaseChildTableID,
			BaseChildColumnID = @piBaseChildColumnID,
			BaseMinimumRatingColumnID = @piBaseMinimumRatingColumnID,
			BasePreferredRatingColumnID = @piBasePreferredRatingColumnID,
			MatchTableID = @piMatchTableID,
			MatchSelection = @piMatchSelection,
			MatchPicklistID = @piMatchPicklistID,
			MatchFilterID = @piMatchFilterID,
			MatchChildTableID = @piMatchChildTableID,
			MatchChildColumnID = @piMatchChildColumnID,
			MatchChildRatingColumnID = @piMatchChildRatingColumnID,
			MatchAgainstType = @piMatchAgainstType
		WHERE ID = @piID;

		DELETE FROM ASRSysTalentReportColumns
			WHERE TalentReportID = @piID;

		EXEC [dbo].[spsys_saveobjectcategories] 38, @piID, @piCategoryID;

	END

	/* Create the details records. */
	SET @sTemp = @psColumns;

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sColumnDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1);
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1);
		END
		ELSE
		BEGIN
			SET @sColumnDefn = @sTemp;
			SET @sTemp = '';
		END

		/* Rip out the column definition parameters. */
		SET @iSequence = 0;
		SET @sType = '';
		SET @iColExprID = 0;
		SET @sHeading = '';
		SET @iSize = 0;
		SET @iDP = 0;
		SET @fIsNumeric = 0;
		SET @iSortOrderSequence = 0;
		SET @sSortOrder = '';
		SET @iCount = 0;
		
		WHILE LEN(@sColumnDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sColumnDefn) > 0
			BEGIN
				SET @sColumnParam = LEFT(@sColumnDefn, CHARINDEX('||', @sColumnDefn) - 1);
				SET @sColumnDefn = RIGHT(@sColumnDefn, LEN(@sColumnDefn) - CHARINDEX('||', @sColumnDefn) - 1);
			END
			ELSE
			BEGIN
				SET @sColumnParam = @sColumnDefn;
				SET @sColumnDefn = '';
			END

			IF @iCount = 0 SET @iSequence = convert(integer, @sColumnParam);
			IF @iCount = 1 SET @sType = @sColumnParam;
			IF @iCount = 2 SET @iColExprID = convert(integer, @sColumnParam);
			IF @iCount = 3 SET @iSize = convert(integer, @sColumnParam);
			IF @iCount = 4 SET @iDP = convert(integer, @sColumnParam);
			IF @iCount = 5 SET @fIsNumeric = convert(bit, @sColumnParam);
			IF @iCount = 6 SET @iSortOrderSequence = convert(integer, @sColumnParam);
			IF @iCount = 7 SET @sSortOrder = @sColumnParam;

			SET @iCount = @iCount + 1;
		END

		INSERT ASRSysTalentReportColumns (TalentReportID, Type, ColumnID, SortOrderSequence, SortOrder, Size, Decimals)
			VALUES (@piID, @sType, @iColExprID, @iSortOrderSequence, @sSortOrder, @iSize, @iDP);

	END

	-- Access permissions
	DELETE FROM ASRSysTalentReportAccess WHERE ID = @piID;
	INSERT INTO ASRSysTalentReportAccess (ID, groupName, access)
		(SELECT @piID, sysusers.name,
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
		FROM sysusers
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.name <> 'ASRSysGroup'
			AND convert(integer, sysusers.uid) <> 0);

	SET @sTemp = @psAccess;
	
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1);
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)));
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1);
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)));
	
			IF EXISTS (SELECT * FROM ASRSysCustomReportAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysTalentReportAccess
					SET access = @sAccess
					WHERE ID = @piID
						AND groupName = @sGroup;
		END
	END

	IF (@fIsNew = 1)
	BEGIN
		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (2, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name());
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 2;

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog	(type, utilID, savedBy, savedDate, savedHost)
			VALUES (38, @piID, system_user, getdate(), host_name());
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID AND type = 38;
		END
	END

	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + '''' ;

		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' +@psJobsToHide + ')
				AND groupName IN (' + @psJobsToHideGroups + ')';
		EXEC sp_executesql @sSQL;

		SET @sSQL = 'INSERT INTO ASRSysBatchJobAccess
			(ID, groupName, access)
			(SELECT ASRSysBatchJobName.ID, 
				sysusers.name,
				CASE
					WHEN (SELECT count(*)
						FROM ASRSysGroupPermissions
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
							OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
						WHERE sysusers.Name = ASRSysGroupPermissions.groupname
							AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
					ELSE ''HD''
				END
			FROM sysusers,
				ASRSysBatchJobName
			WHERE sysusers.uid = sysusers.gid
				AND sysusers.uid <> 0
				AND sysusers.name IN (' + @psJobsToHideGroups + ')
				AND ASRSysBatchJobName.ID IN (' + @psJobsToHide + '))';
		EXEC sp_executesql @sSQL;
	END
END