CREATE PROCEDURE [dbo].[spASRIntSaveCrossTab] (
	@psName				varchar(255),
	@psDescription		varchar(MAX),
	@piTableID			integer,
	@piSelection		integer,
	@piPicklistID		integer,
	@piFilterID			integer,
	@pfPrintFilter		bit,
	@psUserName			varchar(255),
	@piHColID			integer,
	@psHStart			varchar(100),
	@psHStop			varchar(100),
	@psHStep			varchar(100),
	@piVColID			integer,
	@psVStart			varchar(100),
	@psVStop			varchar(100),
	@psVStep			varchar(100),
	@piPColID			integer,
	@psPStart			varchar(100),
	@psPStop			varchar(100),
	@psPStep			varchar(100),
	@piIType			integer,
	@piIColID			integer,
	@pfPercentage		bit,
	@pfPerPage			bit,
	@pfSuppress			bit,
	@pfUse1000Separator	bit,
	@pfOutputPreview	bit,
	@piOutputFormat		integer,
	@pfOutputScreen		bit,
	@pfOutputPrinter	bit,
	@psOutputPrinterName	varchar(MAX),
	@pfOutputSave		bit,
	@piOutputSaveExisting	integer,
	@pfOutputEmail		bit,
	@piOutputEmailAddr	integer,
	@psOutputEmailSubject	varchar(MAX),
	@psOutputEmailAttachAs	varchar(MAX),
	@psOutputFilename	varchar(MAX),
	@psAccess			varchar(MAX),
	@psJobsToHide		varchar(MAX),
	@psJobsToHideGroups	varchar(MAX),
	@piCategoryID		integer,
	@piID				integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
			@fIsNew		bit,
			@sTemp		varchar(MAX),
			@iCount		integer,
			@sGroup		varchar(MAX),
			@sAccess	varchar(MAX),
			@sSQL		nvarchar(MAX);

	DECLARE	@outputTable table (crossTabId int NOT NULL);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''')
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''')

	SET @fIsNew = 0

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysCrossTab (
			Name, 
			Description, 
			TableID, 
			Selection, 
			PicklistID, 
			FilterID, 
 			PrintFilterHeader, 
 			UserName, 
 			HorizontalColID, 
 			HorizontalStart, 
 			HorizontalStop, 
 			HorizontalStep, 
			VerticalColID, 
			VerticalStart, 
			VerticalStop, 
			VerticalStep, 
			PageBreakColID, 
			PageBreakStart, 
			PageBreakStop, 
			PageBreakStep, 
			IntersectionType, 
			IntersectionColID, 
			Percentage, 
			PercentageofPage, 
			SuppressZeros, 
			ThousandSeparators, 
			OutputPreview, 
			OutputFormat, 
			OutputScreen, 
			OutputPrinter, 
			OutputPrinterName, 
			OutputSave, 
			OutputSaveExisting, 
			OutputEmail, 
			OutputEmailAddr, 
			OutputEmailSubject, 
			OutputEmailAttachAs, 
			OutputFileName,
			CrossTabType)
		OUTPUT inserted.crossTabId INTO @outputTable
		VALUES (
			@psName,
			@psDescription,
			@piTableID,
			@piSelection,
			@piPicklistID,
			@piFilterID,
			@pfPrintFilter,
			@psUserName,
			@piHColID,
			@psHStart,
			@psHStop,
			@psHStep,
			@piVColID,
			@psVStart,
			@psVStop,
			@psVStep,
			@piPColID,
			@psPStart,
			@psPStop,
			@psPStep,
			@piIType,
			@piIColID,
			@pfPercentage,
			@pfPerPage,
			@pfSuppress,
			@pfUse1000Separator,
			@pfOutputPreview,
			@piOutputFormat,
			@pfOutputScreen,
			@pfOutputPrinter,
			@psOutputPrinterName,
			@pfOutputSave,
			@piOutputSaveExisting,
			@pfOutputEmail,
			@piOutputEmailAddr,
			@psOutputEmailSubject,
			@psOutputEmailAttachAs,
			@psOutputFilename,
			0 -- Cross tab
		)

		SET @fIsNew = 1

		-- Get the ID of the inserted record.
		SELECT @piID = crossTabId FROM @outputTable;

		Exec [dbo].[spsys_saveobjectcategories] 1, @piID, @piCategoryID

	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysCrossTab SET 
			Name = @psName,
			Description = @psDescription,
			TableID = @piTableID,
			Selection = @piSelection,
			PicklistID = @piPicklistID,
			FilterID = @piFilterID,
			PrintFilterHeader = @pfPrintFilter,
			HorizontalColID = @piHColID,
			HorizontalStart = @psHStart,
			HorizontalStop = @psHStop,
			HorizontalStep = @psHStep,	
			VerticalColID = @piVColID,
			VerticalStart = @psVStart,
			VerticalStop = @psVStop,
			VerticalStep = @psVStep,	
			PageBreakColID = @piPColID,
			PageBreakStart = @psPStart,
			PageBreakStop = @psPStop,
			PageBreakStep = @psPStep,	
			IntersectionType = @piIType,
			IntersectionColID = @piIColID,
			Percentage = @pfPercentage,
			PercentageofPage = @pfPerPage,
			SuppressZeros = @pfSuppress,
			ThousandSeparators = @pfUse1000Separator,
			OutputPreview = @pfOutputPreview,
			OutputFormat = @piOutputFormat,
			OutputScreen = @pfOutputScreen,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName,
			OutputSave = @pfOutputSave,
			OutputSaveExisting = @piOutputSaveExisting,
			OutputEmail = @pfOutputEmail,
			OutputEmailAddr = @piOutputEmailAddr,
			OutputEmailSubject = @psOutputEmailSubject,
			OutputEmailAttachAs = @psOutputEmailAttachAs,
			OutputFileName = @psOutputFilename
		WHERE CrossTabID = @piID

		Exec [dbo].[spsys_saveobjectcategories] 1, @piID, @piCategoryID

	END

	DELETE FROM ASRSysCrossTabAccess WHERE ID = @piID

	INSERT INTO ASRSysCrossTabAccess (ID, groupName, access)
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
			AND sysusers.uid <> 0)

	SET @sTemp = @psAccess
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			IF EXISTS (SELECT * FROM ASRSysCrossTabAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysCrossTabAccess
					SET access = @sAccess
					WHERE ID = @piID
						AND groupName = @sGroup
		END
	END

	IF (@fIsNew = 1)
	BEGIN
		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (1, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 1

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (1, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 1
		END
	END
	
	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + ''''
		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' + @psJobsToHide + ')
				AND groupName IN (' + @psJobsToHideGroups + ')'
		EXEC sp_executesql @sSQL

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
				AND ASRSysBatchJobName.ID IN (' + @psJobsToHide + '))'
		EXEC sp_executesql @sSQL
	END
	
END

