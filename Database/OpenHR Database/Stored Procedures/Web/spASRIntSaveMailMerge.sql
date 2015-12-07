CREATE PROCEDURE [dbo].[spASRIntSaveMailMerge] (
	@psName				varchar(255),
	@psDescription		varchar(MAX),
	@piTableID			integer,
	@piSelection		integer,
	@piPicklistID		integer,
	@piFilterID			integer,
	@UploadTemplate			image = null,
	@UploadTemplateName		nvarchar(255),
	@piOutputFormat			integer,
	@pfOutputSave			bit,
	@psOutputFilename		varchar(MAX),
	@piEmailAddrID		integer,
	@psEmailSubject		varchar(MAX),
	@pfOutputScreen			bit,
	@psUserName			varchar(255),
	@pfEmailAsAttachment	bit,
	@psEmailAttachmentName	varchar(MAX),
	@pfSuppressBlanks		bit,
	@pfPauseBeforeMerge		bit,
	@pfOutputPrinter			bit,
	@psOutputPrinterName	varchar(255),
	@piDocumentMapID			integer,
	@pfManualDocManHeader		bit,	
	@psAccess			varchar(MAX),
	@psJobsToHide		varchar(MAX),
	@psJobsToHideGroups	varchar(MAX),
	@psColumns			varchar(MAX),
	@psColumns2			varchar(MAX),
	@piCategoryID		integer,
	@piID				integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sTemp			varchar(MAX),
			@sColumnDefn	varchar(MAX),
			@sColumnParam	varchar(MAX),
			@iSequence		integer,
			@sType			varchar(MAX),
			@iColExprID		integer,
			@sHeading		varchar(MAX),
			@iSize			integer,
			@iDP			integer,
			@fIsNumeric		bit,
			@fAvge			bit,
			@fCnt			bit,
			@fTot			bit,
			@iSortOrderSequence	integer,
			@sSortOrder		varchar(MAX),
			@fBOC			bit,
			@fPOC			bit,
			@fVOC			bit,
			@fSRV			bit,
			@iCount			integer,
			@fIsNew			bit,
			@sGroup			varchar(255),
			@sAccess		varchar(MAX),
			@sSQL			nvarchar(MAX);

	DECLARE	@outputTable table (MailMergeId int NOT NULL);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''');
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''');
	SET @fIsNew = 0
	/* Insert/update the report header. */
	IF @piID = 0
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysMailMergeName (
			Name, 
			Description, 
			TableID, 
			Selection, 
			PicklistID, 
			FilterID, 
			OutputFormat, 
			OutputSave, 
			OutputFilename, 
			EmailAddrID, 
			EmailSubject, 
			TemplateFileName, 
			UploadTemplate,
			UploadTemplateName,
			OutputScreen, 
			UserName, 
			EMailAsAttachment,
			EmailAttachmentName, 
			SuppressBlanks, 
			PauseBeforeMerge, 
			OutputPrinter,
			OutputPrinterName,
			DocumentMapID,
			ManualDocManHeader,			
			IsLabel, 
			LabelTypeID, 
			PromptStart)
		OUTPUT inserted.MailMergeID INTO @outputTable
		VALUES (
			@psName,
			@psDescription,
			@piTableID,
			@piSelection,
			@piPicklistID,
			@piFilterID,
			@piOutputFormat,
			@pfOutputSave,
			@psOutputFilename,
			@piEmailAddrID,
			@psEmailSubject,
			'',
			@UploadTemplate,
			@UploadTemplateName,
			@pfOutputScreen,
			@psUserName,
			@pfEmailAsAttachment,
			@psEmailAttachmentName,
			@pfSuppressBlanks,
			@pfPauseBeforeMerge,
			@pfOutputPrinter,
			@psOutputPrinterName,
			@piDocumentMapID,
			@pfManualDocManHeader,
			0, 
			0, 
			0);
		SET @fIsNew = 1
		-- Get the ID of the inserted record.
		SELECT @piID = MailMergeId FROM @outputTable;

		/* Insert the category into the table tbsys_objectcategories */
		Exec [dbo].[spsys_saveobjectcategories] 9, @piID, @piCategoryID;

	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysMailMergeName SET 
			Name = @psName,
			Description = @psDescription,
			TableID = @piTableID,
			Selection = @piSelection,
			PicklistID = @piPicklistID,
			FilterID = @piFilterID,
			OutputFormat = @piOutputFormat,
			OutputSave = @pfOutputSave,
			OutputFilename = @psOutputFilename,
			EmailAddrID = @piEmailAddrID,
			EmailSubject = @psEmailSubject,
			OutputScreen = @pfOutputScreen,
			EMailAsAttachment = @pfEmailAsAttachment,
			EmailAttachmentName = @psEmailAttachmentName,
			SuppressBlanks = @pfSuppressBlanks,
			PauseBeforeMerge = @pfPauseBeforeMerge,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName,
			DocumentMapID = @piDocumentMapID,
			ManualDocManHeader = @pfManualDocManHeader,
			UploadTemplate = @UploadTemplate,
			UploadTemplateName = @UploadTemplateName,
			IsLabel = 0,
			LabelTypeID = 0,
			PromptStart = 0
		WHERE MailMergeID = @piID;
		/* Delete existing report details. */
		DELETE FROM ASRSysMailMergeColumns
		WHERE MailMergeID = @piID;

		/* Update the category into the table tbsys_objectcategories */
		Exec [dbo].[spsys_saveobjectcategories] 9, @piID, @piCategoryID;

	END
	/* Create the details records. */
	SET @sTemp = @psColumns
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sColumnDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1);
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1);
			IF len(@sTemp) <= 7000
			BEGIN
				SET @sTemp = @sTemp + LEFT(@psColumns2, 1000);
				IF len(@psColumns2) > 1000
				BEGIN
					SET @psColumns2 = SUBSTRING(@psColumns2, 1001, len(@psColumns2) - 1000);
				END
				ELSE
				BEGIN
					SET @psColumns2 = '';
				END
			END
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
		SET @iSize = 0;
		SET @iDP = 0;
		SET @fIsNumeric = 0;
		SET @iSortOrderSequence = 0;
		SET @sSortOrder = '';
		SET @fBOC = 0;
		SET @fPOC = 0;
		SET @fVOC = 0;
		SET @fSRV = 0;
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
		INSERT ASRSysMailMergeColumns (MailMergeID,Type, ColumnID, SortOrderSequence, SortOrder, Size, Decimals)
		VALUES (@piID, @sType, @iColExprID, @iSortOrderSequence, @sSortOrder, @iSize, @iDP);
	END
	DELETE FROM ASRSysMailMergeAccess WHERE ID = @piID;
	INSERT INTO ASRSysMailMergeAccess (ID, groupName, access)
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
			AND sysusers.uid <> 0);
	SET @sTemp = @psAccess;
	
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1);
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)));
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1);
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)));
	
			IF EXISTS (SELECT * FROM ASRSysMailMergeAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysMailMergeAccess
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
		VALUES (9, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name());
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 9;
		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (9, @piID, system_user, getdate(), host_name());
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 9;
		END
	END
	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + '''';
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
