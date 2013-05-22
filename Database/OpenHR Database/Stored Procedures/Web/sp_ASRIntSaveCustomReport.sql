CREATE PROCEDURE [dbo].[sp_ASRIntSaveCustomReport] (
	@psName						varchar(255),
	@psDescription				varchar(MAX),
	@piBaseTableID				integer,
	@pfAllRecords				bit,
	@piPicklistID				integer,
	@piFilterID					integer,
	@piParent1TableID			integer,
	@piParent1FilterID			integer,
	@piParent2TableID			integer,
	@piParent2FilterID			integer,
	@pfSummary					bit,
	@pfPrintFilterHeader		bit,
	@psUserName					varchar(255),
	@pfOutputPreview			bit,
	@piOutputFormat				integer,
	@pfOutputScreen				bit,
	@pfOutputPrinter			bit,
	@psOutputPrinterName		varchar(MAX),
	@pfOutputSave				bit,
	@piOutputSaveExisting		integer,
	@pfOutputEmail				bit,
	@piOutputEmailAddr			integer,
	@psOutputEmailSubject		varchar(MAX),
	@psOutputEmailAttachAs		varchar(MAX),
	@psOutputFilename			varchar(MAX),
	@pfParent1AllRecords		bit,
	@piParent1Picklist			integer,
	@pfParent2AllRecords		bit,
	@piParent2Picklist			integer,
	@psAccess					varchar(MAX),
	@psJobsToHide				varchar(MAX),
	@psJobsToHideGroups			varchar(MAX),
	@psColumns					varchar(MAX),
	@psColumns2					varchar(MAX),
	@psChildString				varchar(MAX),
	@piID						integer					OUTPUT,
	@pfIgnoreZeros				bit
)
AS
BEGIN
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
			@fAvge					bit,
			@fCnt					bit,
			@fTot					bit,
			@fHidden				bit,
			@fGroupWithNext			bit,
			@iSortOrderSequence		integer,
			@sSortOrder				varchar(MAX),
			@fBOC					bit,
			@fPOC					bit,
			@fVOC					bit,
			@fSRV					bit,
			@fRepetition 			integer,
			@iCount					integer,
			@fIsNew					bit,
			@iChildTableID			integer,
			@iChildFilterID			integer,
			@iChildOrderID			integer,
			@iChildMaxRecords		integer,
			@sChildDefn				varchar(MAX),
			@sChildParam			varchar(MAX),
			@sGroup					varchar(255),
			@sAccess				varchar(MAX),
			@sSQL					nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''')
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''')

	SET @fIsNew = 0

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysCustomReportsName (
			Name, 
			[Description], 
			BaseTable, 
			AllRecords, 
			Picklist, 
			Filter, 
 			Parent1Table, 
 			Parent1Filter, 
 			Parent2Table, 
 			Parent2Filter, 
 			Summary,
 			IgnoreZeros, 
 			PrintFilterHeader, 
 			UserName, 
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
 			Parent1AllRecords, 
 			Parent1Picklist, 
 			Parent2AllRecords, 
 			Parent2Picklist)
 		VALUES (
 			@psName,
 			@psDescription,
 			@piBaseTableID,
 			@pfAllRecords,
 			@piPicklistID,
 			@piFilterID,
 			@piParent1TableID,
 			@piParent1FilterID,
 			@piParent2TableID,
 			@piParent2FilterID,
 			@pfSummary,
 			@pfIgnoreZeros,
 			@pfPrintFilterHeader,
 			@psUserName,
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
 			@pfParent1AllRecords,
 			@piParent1Picklist,
 			@pfParent2AllRecords,
 			@piParent2Picklist
		)

		SET @fIsNew = 1
		/* Get the ID of the inserted record.*/
		SELECT @piID = MAX(ID) FROM ASRSysCustomReportsName
	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSYSCustomReportsName SET 
			Name = @psName,
			Description = @psDescription,
			BaseTable = @piBaseTableID,
			AllRecords = @pfAllRecords,
			Picklist = @piPicklistID,
			Filter = @piFilterID,
			Parent1Table = @piParent1TableID,
			Parent1Filter = @piParent1FilterID,
			Parent2Table = @piParent2TableID,
			Parent2Filter = @piParent2FilterID,
			Summary = @pfSummary,
			IgnoreZeros = @pfIgnoreZeros,
			PrintFilterHeader = @pfPrintFilterHeader,
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
			OutputFileName = @psOutputFilename,
			Parent1AllRecords = @pfParent1AllRecords,
			Parent1Picklist = @piParent1Picklist,
			Parent2AllRecords = @pfParent2AllRecords,
			Parent2Picklist = @piParent2Picklist
		WHERE ID = @piID

		/* Delete existing report details. */
		DELETE FROM ASRSysCustomReportsDetails 
		WHERE customReportID = @piID
	END

	/* Create the details records. */
	SET @sTemp = @psColumns

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sColumnDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)

			IF len(@sTemp) <= 7000
			BEGIN
				SET @sTemp = @sTemp + LEFT(@psColumns2, 1000)
				IF len(@psColumns2) > 1000
				BEGIN
					SET @psColumns2 = SUBSTRING(@psColumns2, 1001, len(@psColumns2) - 1000)
				END
				ELSE
				BEGIN
					SET @psColumns2 = ''
				END
			END
		END
		ELSE
		BEGIN
			SET @sColumnDefn = @sTemp
			SET @sTemp = ''
		END

		/* Rip out the column definition parameters. */
		SET @iSequence = 0
		SET @sType = ''
		SET @iColExprID = 0
		SET @sHeading = ''
		SET @iSize = 0
		SET @iDP = 0
		SET @fIsNumeric = 0
		SET @fAvge = 0
		SET @fCnt = 0
		SET @fTot = 0
		SET @fHidden = 0
		SET @fGroupWithNext = 0
		SET @iSortOrderSequence = 0
		SET @sSortOrder = ''
		SET @fBOC = 0
		SET @fPOC = 0
		SET @fVOC = 0
		SET @fSRV = 0
		SET @fRepetition = 0
		SET @iCount = 0
		
		WHILE LEN(@sColumnDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sColumnDefn) > 0
			BEGIN
				SET @sColumnParam = LEFT(@sColumnDefn, CHARINDEX('||', @sColumnDefn) - 1)
				SET @sColumnDefn = RIGHT(@sColumnDefn, LEN(@sColumnDefn) - CHARINDEX('||', @sColumnDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sColumnParam = @sColumnDefn
				SET @sColumnDefn = ''
			END

			IF @iCount = 0 SET @iSequence = convert(integer, @sColumnParam)
			IF @iCount = 1 SET @sType = @sColumnParam
			IF @iCount = 2 SET @iColExprID = convert(integer, @sColumnParam)
			IF @iCount = 3 SET @sHeading = @sColumnParam
			IF @iCount = 4 SET @iSize = convert(integer, @sColumnParam)
			IF @iCount = 5 SET @iDP = convert(integer, @sColumnParam)
			IF @iCount = 6 SET @fIsNumeric = convert(bit, @sColumnParam)
			IF @iCount = 7 SET @fAvge = convert(bit, @sColumnParam)
			IF @iCount = 8 SET @fCnt = convert(bit, @sColumnParam)
			IF @iCount = 9 SET @fTot = convert(bit, @sColumnParam)
			IF @iCount = 10 SET @fHidden = convert(bit, @sColumnParam)
			IF @iCount = 11 SET @fGroupWithNext = convert(bit, @sColumnParam)
			IF @iCount = 12 SET @iSortOrderSequence = convert(integer, @sColumnParam)
			IF @iCount = 13 SET @sSortOrder = @sColumnParam
			IF @iCount = 14 SET @fBOC = convert(bit, @sColumnParam)
			IF @iCount = 15 SET @fPOC = convert(bit, @sColumnParam)
			IF @iCount = 16 SET @fVOC = convert(bit, @sColumnParam)
			IF @iCount = 17 SET @fSRV = convert(bit, @sColumnParam)
			IF @iCount = 18 SET @fRepetition = convert(integer, @sColumnParam)

			SET @iCount = @iCount + 1
		END

		INSERT ASRSysCustomReportsDetails 
			(customReportID, sequence, type, colExprID, heading, size, dp, isNumeric, avge, 
			cnt, tot, hidden, GroupWithNextColumn, 	sortOrderSequence, sortOrder, boc, poc, voc, srv, repetition) 
		VALUES (@piID, @iSequence, @sType, @iColExprID, @sHeading, @iSize, @iDP, @fIsNumeric, @fAvge, 
			@fCnt, @fTot, @fHidden, @fGroupWithNext, @iSortOrderSequence, @sSortOrder, @fBOC, @fPOC, @fVOC, @fSRV, @fRepetition)

	END

	/* Create the table records. */

	IF (@fIsNew = 0)
	BEGIN
		/* Delete existing report child tables. */
		DELETE FROM ASRSysCustomReportsChildDetails 
		WHERE customReportID = @piID
	END

	SET @sTemp = @psChildString

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sChildDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)
		END
		ELSE
		BEGIN
			SET @sChildDefn = @sTemp
			SET @sTemp = ''
		END

		/* Rip out the column definition parameters. */
		SET @iChildTableID = 0
		SET @iChildFilterID = 0
		SET @iChildOrderID = 0
		SET @iChildMaxRecords = 0
		SET @iCount = 0
		
		WHILE LEN(@sChildDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sChildDefn) > 0
			BEGIN
				SET @sChildParam = LEFT(@sChildDefn, CHARINDEX('||', @sChildDefn) - 1)
				SET @sChildDefn = RIGHT(@sChildDefn, LEN(@sChildDefn) - CHARINDEX('||', @sChildDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sChildParam = @sChildDefn
				SET @sChildDefn = ''
			END

			IF @iCount = 0 SET @iChildTableID = convert(integer, @sChildParam)
			IF @iCount = 1 SET @iChildFilterID = convert(integer, @sChildParam)
			IF @iCount = 2 SET @iChildOrderID = convert(integer, @sChildParam)
			IF @iCount = 3 SET @iChildMaxRecords = convert(integer, @sChildParam)
	
			SET @iCount = @iCount + 1
		END

		INSERT ASRSysCustomReportsChildDetails 
			(customReportID, childtable, childfilter, childorder, childmaxrecords) 
		VALUES (@piID, @iChildTableID, @iChildFilterID, @iChildOrderID, @iChildMaxRecords)

	END

	DELETE FROM ASRSysCustomReportAccess WHERE ID = @piID
	INSERT INTO ASRSysCustomReportAccess (ID, groupName, access)
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
	
			IF EXISTS (SELECT * FROM ASRSysCustomReportAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysCustomReportAccess
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
		VALUES (2, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 2

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (2, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 2
		END
	END

	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + ''''

		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' +@psJobsToHide + ')
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



GO

