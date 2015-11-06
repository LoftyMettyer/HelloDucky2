CREATE PROCEDURE [dbo].[spASRIntSaveCalendarReport]
	(
	@psName						varchar(255),
	@psDescription				varchar(MAX),
	@piBaseTable				integer,
	@pfAllRecords				bit,
	@piPicklist					integer,
	@piFilter					integer,
	@pfPrintFilterHeader		bit,
	@psUserName					varchar(255),
	@piDescription1				integer,
	@piDescription2				integer,
	@piDescriptionExpr			integer,
	@piRegion					integer,
	@pfGroupByDesc				bit,
	@psDescSeparator			varchar(100),
	@piStartType				integer,
	@psFixedStart				varchar(100),
	@piStartFrequency			integer,
	@piStartPeriod				integer,
	@piStartDateExpr			integer,
	@piEndType					integer,
	@psFixedEnd					varchar(100),
	@piEndFrequency				integer,
	@piEndPeriod				integer,
	@piEndDateExpr				integer,
	@pfShowBankHols				bit,
	@pfShowCaptions				bit,
	@pfShowWeekends				bit,
	@pfStartOnCurrentMonth		bit,
	@pfIncludeWorkdays			bit,
	@pfIncludeBankHols			bit,
	@pfOutputPreview			bit,
	@piOutputFormat				integer,
	@pfOutputScreen				bit,
	@pfOutputPrinter			bit,
	@psOutputPrinterName		varchar(MAX),
	@pfOutputSave				bit,
	@piOutputSaveExisting		integer,
	@pfOutputEmail				bit,
	@pfOutputEmailAddr			integer,
	@psOutputEmailSubject		varchar(MAX),
	@psOutputEmailAttachAs		varchar(MAX),
	@psOutputFilename			varchar(MAX),
	@psAccess					varchar(MAX),
	@psJobsToHide				varchar(MAX),
	@psJobsToHideGroups			varchar(MAX),
	@psEvents					varchar(MAX),
	@psEvents2					varchar(MAX),
	@psOrderString				varchar(MAX),
	@piCategoryID				integer,
	@piID						integer	OUTPUT
	)
AS
BEGIN 

	SET NOCOUNT ON;

	DECLARE	@sTemp					varchar(MAX),
			@iCount					integer,
			@fIsNew					bit,
			@sEventDefn				varchar(MAX),
			@sEventParam			varchar(MAX),
			@sEventKey				varchar(MAX),
			@sEventName				varchar(MAX),
			@iEventTableID			integer,
			@iEventFilterID			integer,
			@iEventStartDateID		integer,
			@iEventStartSessionID	integer,
			@iEventEndDateID		integer,
			@iEventEndSessionID		integer,
			@iEventDurationID		integer,
			@iLegendType			integer,
			@sLegendCharacter		varchar(2),
			@iLegendLookupTableID	integer,
			@iLegendLookupColumnID	integer,
			@iLegendLookupCodeID	integer,
			@iLegendEventColumnID	integer,
			@iEventDesc1ColumnID	integer,
			@iEventDesc2ColumnID	integer,
			@sOrderDefn				varchar(MAX),
			@sOrderParam			varchar(MAX),
			@iOrderTableID			integer,
			@iOrderColumnID			integer,
			@iOrderSequence			integer,
			@sOrderType				varchar(20),
			@sGroup					varchar(255),
			@sAccess				varchar(MAX),
			@sSQL					nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''');
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''');

	SET @fIsNew = 0;

	/* Insert/update the report header. */
	IF @piID = 0
	BEGIN
		/* Creating a new report. */
		INSERT ASRSYSCalendarReports (
			Name, 
			[Description], 
			BaseTable, 
			AllRecords, 
			Picklist, 
			Filter, 
			PrintFilterHeader, 
			UserName, 
			Description1, 
			Description2, 
			DescriptionExpr, 
			Region,
			GroupByDesc,
			DescriptionSeparator, 
			StartType, 
			FixedStart, 
			StartFrequency,
			StartPeriod,
			StartDateExpr,
			EndType,
			FixedEnd,
			EndFrequency,
			EndPeriod,
			EndDateExpr,
			ShowBankHolidays,
			ShowCaptions,
			ShowWeekends,
			StartOnCurrentMonth, 
			IncludeWorkingDaysOnly, 
			IncludeBankHolidays,
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
			OutputFileName)
		VALUES (
			@psName,
			@psDescription,
			@piBaseTable,
			@pfAllRecords,
			@piPicklist,
			@piFilter,
			@pfPrintFilterHeader,
			@psUserName,
			@piDescription1,
			@piDescription2,
			@piDescriptionExpr,
			@piRegion,
			@pfGroupByDesc,
			@psDescSeparator,
			@piStartType,
			@psFixedStart,
			@piStartFrequency,
			@piStartPeriod,
			@piStartDateExpr,
			@piEndType,
			@psFixedEnd,
			@piEndFrequency,
			@piEndPeriod,
			@piEndDateExpr,
			@pfShowBankHols,
			@pfShowCaptions,
			@pfShowWeekends,
			@pfStartOnCurrentMonth,
			@pfIncludeWorkdays,
			@pfIncludeBankHols,
			@pfOutputPreview,
			@piOutputFormat,
			@pfOutputScreen,
			@pfOutputPrinter,
			@psOutputPrinterName,
			@pfOutputSave,
			@piOutputSaveExisting,
			@pfOutputEmail,
			@pfOutputEmailAddr,
			@psOutputEmailSubject,
			@psOutputEmailAttachAs,
			@psOutputFilename
		);
		
		SET @fIsNew = 1;
		/* Get the ID of the inserted record.*/
		SELECT @piID = MAX(ID) FROM ASRSysCalendarReports;

		Exec [dbo].[spsys_saveobjectcategories] 17 , @piID, @piCategoryID

	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysCalendarReports SET 
			Name = @psName,
			[Description] = @psDescription, 
			BaseTable = @piBaseTable, 
			AllRecords = @pfAllRecords, 
			Picklist = @piPicklist, 
			Filter = @piFilter,
			PrintFilterHeader = @pfPrintFilterHeader,
			Description1 = @piDescription1,
			Description2 = @piDescription2,
			DescriptionExpr = @piDescriptionExpr,
			Region = @piRegion,
			GroupByDesc = @pfGroupByDesc,
			DescriptionSeparator = @psDescSeparator,
			StartType = @piStartType,
			FixedStart = @psFixedStart, 
			StartFrequency = @piStartFrequency,
			StartPeriod = @piStartPeriod,
			StartDateExpr = @piStartDateExpr,
			EndType = @piEndType,
			FixedEnd = @psFixedEnd, 
			EndFrequency = @piEndFrequency,
			EndPeriod = @piEndPeriod,
			EndDateExpr = @piEndDateExpr,
			ShowBankHolidays = @pfShowBankHols,
			ShowCaptions = @pfShowCaptions,
			ShowWeekends = @pfShowWeekends,
			StartOnCurrentMonth = @pfStartOnCurrentMonth,
			IncludeWorkingDaysOnly = @pfIncludeWorkdays,
			IncludeBankHolidays = @pfIncludeBankHols,
			OutputPreview = @pfOutputPreview,
			OutputFormat = @piOutputFormat,
			OutputScreen = @pfOutputScreen,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName, 
			OutputSave = @pfOutputSave,
			OutputSaveExisting = @piOutputSaveExisting,
			OutputEmail = @pfOutputEmail,
			OutputEmailAddr = @pfOutputEmailAddr,
			OutputEmailSubject = @psOutputEmailSubject,
			OutputEmailAttachAs = @psOutputEmailAttachAs,
			OutputFileName = @psOutputFilename  
			WHERE ID = @piID;
		
		Exec [dbo].[spsys_saveobjectcategories] 17 , @piID, @piCategoryID

		/* Delete existing report event details. */
		DELETE FROM ASRSysCalendarReportEvents 
		WHERE calendarReportID = @piID;
	END

	/* Create the report's event details records. */
	SET @sTemp = @psEvents;

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sEventDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)

			IF len(@sTemp) <= 7000
			BEGIN
				SET @sTemp = @sTemp + LEFT(@psEvents2, 1000)
				IF len(@psEvents2) > 1000
				BEGIN
					SET @psEvents2 = SUBSTRING(@psEvents2, 1001, len(@psEvents2) - 1000)
				END
				ELSE
				BEGIN
					SET @psEvents2 = ''
				END
			END
		END
		ELSE
		BEGIN
			SET @sEventDefn = @sTemp
			SET @sTemp = ''
		END

		/* Rip out the event definition parameters. */
		SET @sEventKey = '';
		SET @sEventName = '';
		SET @iEventTableID = 0;
		SET @iEventFilterID = 0;
		SET @iEventStartDateID = 0;
		SET @iEventStartSessionID = 0;
		SET @iEventEndDateID = 0;
		SET @iEventEndSessionID = 0;
		SET @iEventDurationID = 0;
		SET @iLegendType = 0;
		SET @sLegendCharacter = '';
		SET @iLegendLookupTableID = 0;
		SET @iLegendLookupColumnID = 0;
		SET @iLegendLookupCodeID = 0;
		SET @iLegendEventColumnID = 0;
		SET @iEventDesc1ColumnID = 0;
		SET @iEventDesc2ColumnID = 0;
		
		SET @iCount = 0;
		
		WHILE LEN(@sEventDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sEventDefn) > 0
			BEGIN
				SET @sEventParam = LEFT(@sEventDefn, CHARINDEX('||', @sEventDefn) - 1)
				SET @sEventDefn = RIGHT(@sEventDefn, LEN(@sEventDefn) - CHARINDEX('||', @sEventDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sEventParam = @sEventDefn
				SET @sEventDefn = ''
			END

			IF @iCount = 0 SET @sEventKey = @sEventParam;
			IF @iCount = 1 SET @sEventName = @sEventParam;
			IF @iCount = 2 SET @iEventTableID = convert(integer, @sEventParam);
			IF @iCount = 3 SET @iEventFilterID = convert(integer, @sEventParam);
			IF @iCount = 4 SET @iEventStartDateID = convert(integer, @sEventParam);
			IF @iCount = 5 SET @iEventStartSessionID = convert(integer, @sEventParam);
			IF @iCount = 6 SET @iEventEndDateID = convert(integer, @sEventParam);
			IF @iCount = 7 SET @iEventEndSessionID = convert(integer, @sEventParam);
			IF @iCount = 8 SET @iEventDurationID = convert(integer, @sEventParam);
			IF @iCount = 9 SET @iLegendType = convert(integer, @sEventParam);
			
			IF (@iCount = 10)
				BEGIN 
					IF @iLegendType = 0
						BEGIN
							SET @sLegendCharacter = LEFT(@sEventParam,2);
						END	
					ELSE
						BEGIN
							SET @sLegendCharacter = '';
						END
				END
			IF @iCount = 11 SET @iLegendLookupTableID = convert(integer, @sEventParam);
			IF @iCount = 12 SET @iLegendLookupColumnID = convert(integer, @sEventParam);
			IF @iCount = 13 SET @iLegendLookupCodeID = convert(integer, @sEventParam);
			IF @iCount = 14 SET @iLegendEventColumnID = convert(integer, @sEventParam);
			IF @iCount = 15 SET @iEventDesc1ColumnID = convert(integer, @sEventParam);
			IF @iCount = 16 SET @iEventDesc2ColumnID = convert(integer, @sEventParam);

			SET @iCount = @iCount + 1;
		END

		INSERT ASRSysCalendarReportEvents (EventKey, CalendarReportID, Name, TableID, FilterID, 
				EventStartDateID, EventStartSessionID, EventEndDateID, EventEndSessionID, 
				EventDurationID, LegendType, LegendCharacter, LegendLookupTableID, LegendLookupColumnID, 
				LegendLookupCodeID, LegendEventColumnID, EventDesc1ColumnID, EventDesc2ColumnID)
		VALUES (@sEventKey, @piID, @sEventName, @iEventTableID, @iEventFilterID, 
				@iEventStartDateID, @iEventStartSessionID, @iEventEndDateID, @iEventEndSessionID, 
				@iEventDurationID, @iLegendType, @sLegendCharacter, @iLegendLookupTableID, @iLegendLookupColumnID, 
				@iLegendLookupCodeID, @iLegendEventColumnID, @iEventDesc1ColumnID, @iEventDesc2ColumnID);

	END


	/* Create the report's sort order details records. */
	IF (@fIsNew = 0)
	BEGIN
		/* Delete existing report sort order details. */
		DELETE FROM ASRSysCalendarReportOrder 
		WHERE calendarReportID = @piID;
	END

	SET @sTemp = @psOrderString;

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sOrderDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1);
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1);
		END
		ELSE
		BEGIN
			SET @sOrderDefn = @sTemp;
			SET @sTemp = '';
		END

		/* Rip out the column definition parameters. */
		SET @iOrderTableID = 0;
		SET @iOrderColumnID = 0;
		SET @iOrderSequence = 0;
		SET @sOrderType = '';
		
		SET @iCount = 0;

		WHILE LEN(@sOrderDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sOrderDefn) > 0
			BEGIN
				SET @sOrderParam = LEFT(@sOrderDefn, CHARINDEX('||', @sOrderDefn) - 1);
				SET @sOrderDefn = RIGHT(@sOrderDefn, LEN(@sOrderDefn) - CHARINDEX('||', @sOrderDefn) - 1);
			END
			ELSE
			BEGIN
				SET @sOrderParam = @sOrderDefn;
				SET @sOrderDefn = '';
			END

			--IF @iCount = 0 SET @iOrderTableID = convert(integer, @sOrderParam)
			IF @iCount = 0 SET @iOrderColumnID = convert(integer, @sOrderParam);
			IF @iCount = 1 SET @iOrderSequence = convert(integer, @sOrderParam);
			IF @iCount = 2 SET @sOrderType = @sOrderParam;
	
			SET @iCount = @iCount + 1;
		END

		SELECT @iOrderTableID = ASRSysColumns.TableID
		FROM ASRSysColumns
		WHERE ASRSysColumns.ColumnID = @iOrderColumnID;
		
		INSERT ASRSysCalendarReportOrder 
			(CalendarReportID, TableID, ColumnID, OrderSequence, OrderType) 
		VALUES (@piID, @iOrderTableID, @iOrderColumnID, @iOrderSequence, @sOrderType);

	END
	
	DELETE FROM ASRSysCalendarReportAccess WHERE ID = @piID
	INSERT INTO ASRSysCalendarReportAccess (ID, groupName, access)
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
	
			IF EXISTS (SELECT * FROM ASRSysCalendarReportAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysCalendarReportAccess
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
		VALUES (17, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name());
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND [type] = 17;

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				([type], utilID, savedBy, savedDate, savedHost)
			VALUES (17, @piID, system_user, getdate(), host_name());
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND [type] = 17;
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

