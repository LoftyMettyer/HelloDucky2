CREATE PROCEDURE [dbo].[spASRIntSaveOrganisationReport] (

	@psName						varchar(255),
	@psDescription				varchar(MAX),
	@piCategoryID				integer,
	@piBaseViewID				integer,
	@psUserName					varchar(255),
	@psAccess					varchar(MAX),
	@psFilterDef				varchar(MAX),
	@psColumns					varchar(MAX),
	@piID						integer					OUTPUT
	
)
AS
BEGIN
	
	SET NOCOUNT ON;

	DECLARE	@sTemp					varchar(MAX),
			@sColumnDefn			varchar(MAX),
			@sColumnParam			varchar(MAX),
			@iColumnID				integer,
			@sPrefix				varchar(50),
			@sSuffix				varchar(50),
			@iFontSize				integer,
			@iHeight				integer,
			@iDP					integer,
			@fConcatenateWithNext	bit,
			@iCount					integer,
			@fIsNew					bit,
			@sGroup					varchar(255),
			@sAccess				varchar(MAX),
			@sTempFilter			varchar(MAX),
			@sFilterParam			varchar(MAX),
			@iFieldID				integer,
			@iOperator				integer,
			@sValue					varchar(Max);

			DECLARE	@outputTable table (id int NOT NULL);

	SET @fIsNew = 0;

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysOrganisationReport(
			[Name]
           ,[Description]
           ,[BaseViewID]
           ,[UserName])
		OUTPUT inserted.ID INTO @outputTable
 		VALUES (
 			@psName,
 			@psDescription,
 			@piBaseViewID,
			@psUserName);

		SET @fIsNew = 1;
		-- Get the ID of the inserted record.
		SELECT @piID = id FROM @outputTable;

		EXEC [dbo].[spsys_saveobjectcategories] 39, @piID, @piCategoryID;

	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysOrganisationReport SET 
			Name = @psName,
			Description = @psDescription,
			BaseViewID = @piBaseViewID
		WHERE ID = @piID;

		DELETE FROM ASRSysOrganisationColumns
			WHERE OrganisationID = @piID;

		DELETE FROM ASRSysOrganisationReportFilters
			WHERE OrganisationID = @piID;

		EXEC [dbo].[spsys_saveobjectcategories] 39, @piID, @piCategoryID;

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
		SET @iColumnID = 0;
		SET @sPrefix = '';
		SET @sSuffix = '';
		SET @iFontSize = 0;
		SET @iDP = 0;
		SET @iHeight = 0;
		Set @fConcatenateWithNext = 0;
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

			IF @iCount = 0 SET @iColumnID = convert(integer, @sColumnParam);
			IF @iCount = 1 SET @sPrefix = @sColumnParam;
			IF @iCount = 2 SET @sSuffix = @sColumnParam;
			IF @iCount = 3 SET @iFontSize = convert(integer, @sColumnParam);
			IF @iCount = 4 SET @iDP = convert(integer, @sColumnParam);
			IF @iCount = 4 SET @iHeight = convert(integer, @sColumnParam);
			IF @iCount = 5 SET @fConcatenateWithNext = convert(bit, @sColumnParam);

			SET @iCount = @iCount + 1;
		END

		INSERT ASRSysOrganisationColumns (OrganisationID, ColumnID, Prefix, Suffix, FontSize, Decimals, Height, ConcatenateWithNext)
			VALUES (@piID, @iColumnID, @sPrefix, @sSuffix, @iFontSize, @iDP, @iHeight, @fConcatenateWithNext);

	END

	/* Create the records for filters */
	SET @sTemp = @psFilterDef;

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sTempFilter = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1);
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1);
		END
		ELSE
		BEGIN
			SET @sTempFilter = @sTemp;
			SET @sTemp = '';
		END

		/* Rip out the filter definition parameters. */
		SET @iFieldID = 0;
		SET @iOperator = 0;
		SET @sValue = '';
		SET @iCount = 0;
		
		WHILE LEN(@sTempFilter) > 0
		BEGIN
			IF CHARINDEX('||', @sTempFilter) > 0
			BEGIN
				SET @sFilterParam = LEFT(@sTempFilter, CHARINDEX('||', @sTempFilter) - 1);
				SET @sTempFilter = RIGHT(@sTempFilter, LEN(@sTempFilter) - CHARINDEX('||', @sTempFilter) - 1);
			END
			ELSE
			BEGIN
				SET @sFilterParam = @sTempFilter;
				SET @sTempFilter = '';
			END

			IF @iCount = 0 SET @iFieldID = convert(integer, @sFilterParam);
			IF @iCount = 1 SET @iOperator = convert(integer, @sFilterParam);
			IF @iCount = 2 SET @sValue = @sFilterParam;

			SET @iCount = @iCount + 1;
		END

		INSERT ASRSysOrganisationReportFilters (OrganisationID, FieldID, Operator, Value)
			VALUES (@piID, @iFieldID, @iOperator, @sValue);

	END

	-- Access permissions
	DELETE FROM ASRSysOrganisationReportAccess WHERE ID = @piID;
	INSERT INTO ASRSysOrganisationReportAccess (ID, groupName, access)
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
	
			IF EXISTS (SELECT * FROM ASRSysOrganisationReportAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysOrganisationReportAccess
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
		VALUES (39, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name());
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 39;

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog	(type, utilID, savedBy, savedDate, savedHost)
			VALUES (39, @piID, system_user, getdate(), host_name());
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID AND type = 39;
		END
	END

END
