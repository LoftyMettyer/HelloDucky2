
/* ---------------------------------------------------- */
/* Update the database from version 2.9 to version 2.10  */
/* ----------------------------------------------------	*/

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@sDBVersion varchar(10),
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @NVarCommand nvarchar(4000),
	@sColumnDataType varchar(8000),
	@iDateFormat varchar(255),
	@sSQLVersion nvarchar(20)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 2.9 or 2.10. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.37') and (@sDBVersion <> '2.9') and (@sDBVersion <> '2.10')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* Get the SQL Server version */
SET @sSQLVersion = SUBSTRING(@@VERSION, 30, 8)


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 120 - Amending System Permissions Table'

	EXEC sp_rename 'ASRSysPermissionCategories.categoryID', 'ToBeDeleted', 'COLUMN' 
	
	SELECT @NVarCommand = 'ALTER TABLE ASRSysPermissionCategories ADD [categoryID] [int]'
	EXEC sp_executesql @NVarCommand
	
	SET @NVarCommand = 'UPDATE ASRSysPermissionCategories SET [categoryID] = [ToBeDeleted]'
	EXEC sp_executesql @NVarCommand

	SELECT @iRecCount = count(id) FROM sysobjects
	WHERE name = 'PK_ASRSysPermissionCategories' AND type = 'K'

	if @iRecCount = 1
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysPermissionCategories
					DROP CONSTRAINT PK_ASRSysPermissionCategories'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @NVarCommand = 'ALTER TABLE ASRSysPermissionCategories DROP COLUMN [ToBeDeleted]'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 2 of 120 - Adding new columns to Column Definitions'

		/* Add new digit group columns */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysColumns')
		and name = 'Trimming'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
						[Trimming] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SET @NVarCommand = 'UPDATE ASRSysColumns SET [Trimming] = 0'
			EXEC sp_executesql @NVarCommand

		END


		/* Add new digit group columns */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysColumns')
		and name = 'Use1000Separator'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
						[Use1000Separator] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SET @NVarCommand = 'UPDATE ASRSysColumns SET [Use1000Separator] = 0'
			EXEC sp_executesql @NVarCommand

		END


/* ------------------------------------------------------------- */
PRINT 'Step 3 of 120 - Adding Tables for Match Reports'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysMatchReportName]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysMatchReportName] (
						[MatchReportID] [int] IDENTITY (1, 1) NOT NULL ,
						[Name] [varchar] (50) NOT NULL ,
						[Description] [varchar] (255) NOT NULL ,
						[Table1ID] [int] NOT NULL ,
						[Table1AllRecords] [bit] NOT NULL ,
						[Table1Picklist] [int] NOT NULL ,
						[Table1Filter] [int] NOT NULL ,
						[Table2ID] [int] NOT NULL ,
						[Table2AllRecords] [bit] NOT NULL ,
						[Table2Picklist] [int] NOT NULL ,
						[Table2Filter] [int] NOT NULL ,
						[Access] [varchar] (2) NOT NULL ,
						[UserName] [varchar] (50) NOT NULL ,
						[NumRecords] [int] NOT NULL ,
						[OutputPreview] [bit] NOT NULL ,
						[OutputFormat] [int] NOT NULL ,
						[OutputScreen] [bit] NOT NULL ,
						[OutputPrinter] [bit] NOT NULL ,
						[OutputPrinterName] [varchar] (255) NOT NULL ,
						[OutputSave] [bit] NOT NULL ,
						[OutputSaveExisting] [int] NOT NULL ,
						[OutputEmail] [bit] NOT NULL ,
						[OutputEmailAddr] [int] NOT NULL ,
						[OutputEmailSubject] [varchar] (255) NOT NULL ,
						[OutputFilename] [varchar] (255) NOT NULL ,
						[Timestamp] [timestamp] NOT NULL 
					) ON [PRIMARY] '
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysMatchReportDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysMatchReportDetails]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysMatchReportDetails] (
						[MatchReportID] [int] NOT NULL ,
						[ColType] [char] (1) NOT NULL ,
						[ColExprID] [int] NOT NULL ,
						[ColSize] [int] NOT NULL ,
						[ColDecs] [int] NOT NULL ,
						[ColHeading] [varchar] (255) NOT NULL ,
						[ColSequence] [int] NOT NULL ,
						[SortOrderSeq] [int] NOT NULL ,
						[SortOrderDirection] [varchar] (4) NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysMatchReportBreakdown]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysMatchReportBreakdown]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysMatchReportBreakdown] (
						[MatchReportID] [int] NOT NULL ,
						[MatchRelationID] [int] NOT NULL ,
						[ColType] [varchar] (1) NOT NULL ,
						[ColExprID] [int] NOT NULL ,
						[ColSize] [int] NOT NULL ,
						[ColDecs] [int] NOT NULL ,
						[ColHeading] [varchar] (255) NOT NULL ,
						[ColSequence] [int] NOT NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysMatchReportTables]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysMatchReportTables]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysMatchReportTables] (
						[MatchReportID] [int] NOT NULL ,
						[MatchRelationID] [int] IDENTITY (1, 1) NOT NULL ,
						[Table1ID] [int] NOT NULL ,
						[Table2ID] [int] NOT NULL ,
						[RequiredExprID] [int] NOT NULL ,
						[PreferredExprID] [int] NOT NULL ,
						[MatchScoreExprID] [int] NOT NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

		END



/* ------------------------------------------------------------- */
PRINT 'Step 4 of 120 - Adding new columns to Match Reports'

		/* Adding Calculation Columns to Calendar Reports definition table */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysMatchReportName')
		and name = 'MatchReportType'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysMatchReportName] ADD
						[MatchReportType] [integer] NULL,
						[ScoreMode] [integer] NULL,
						[ScoreCheck] [bit] NULL,
						[ScoreLimit] [integer] NULL,
						[EqualGrade] [bit] NULL,
						[ReportingStructure] [bit] NULL,
						[PrintFilterHeader] [bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysMatchReportName SET
						[MatchReportType] = 0,
						[ScoreMode] = 0,
						[ScoreCheck] = 0,
						[ScoreLimit] = 0,
						[EqualGrade] = 0,
						[ReportingStructure] = 0,
						[PrintFilterHeader] = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 5 of 120 - Adding Delete Trigger for Match Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysMatchReportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysMatchReportName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysMatchReportName ON dbo.ASRSysMatchReportName 
					FOR DELETE AS

					delete from ASRSysExpressions where type in (15, 16, 17) and (
					exprid IN (select RequiredExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)) OR
					exprid IN (select PreferredExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)) OR
					exprid IN (select MatchScoreExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)))

					delete from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)
					delete from ASRSysMatchReportDetails WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)
					delete from ASRSysMatchReportBreakdown WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)'
		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 6 of 120 - Updating System Permissions for Match Reports'

		/* Adding System Permissions for Match Reports */
		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 23

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(23,
					'Match Reports',
					'',
					10,
					'MATCHREPORTS')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 23

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000000000000680300001600000028000000100000002000000001001800000000004003000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000808080808080000000000000000000000000000000000000808080808080000000000000000000000000000000808080000000000000000000000000000000000000000000000000000000000000808080000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000808080FFFFFF808080FFFFFF000000000000000000000000000000000000000000000000000000000000FFFFFF808080FFFFFFFFFFFFFFFFFF808080FFFFFFFFFFFF000000000000000000000000000000000000000000000000808080FFFFFFFFFFFF000000FFFFFFFFFFFF000000000000000000000000000000000000000000000000000000C0C0C0FFFFFFFFFFFF000000FFFFFF000000000000FFFFFFFFFFFFC0C0C0000000000000000000000000FFFF00808000000000C0C0C0FFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFC0C0C000000000FFFF000000000000808000000000FFFF0080800000000000000000000000000000000000000000000000000000FFFF000000800000800000808000808000000000FFFF0080800000000000000000000000000000000000000000FFFF000000800000FF0000800000808000808000808000000000000000000000000000000000000000000000000000000000FF0000FF0000800000FF0000808000808000808000000000000000000000000000000000000000000000000000000000FF0000FF0000FF0000800000808000808000000000000000000000000000000000000000000000000000000000000000FF0000FF0000FF0000FF0000808000000000000000000000000000000000000000000000000000000000000000000000000000FF0000FF0000FF0000808000000000000000000000000000000000000000000000000000000000000000000000000000000000FF0000FF0000FFFF0000E7E70000DFFB0000FC3F0000F00F0000E0070000E007000080010000000000000000000003C0000007C000000FE000001FE000003FF000003FF80000


			DELETE FROM ASRSysPermissionItems WHERE itemid in (101,102,103,104,105)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (101,'New',10,23,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (102,'Edit',20,23,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (103,'View',30,23,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (104,'Delete',40,23,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (105,'Run',50,23,'RUN')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (101,102,103,104,105)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 101, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 102, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 103, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 104, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 105, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END


/* ------------------------------------------------------------- */
PRINT 'Step 7 of 120 - Updating System Permissions for Succession Planning'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 38

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			INSERT INTO ASRSysPermissionCategories
			(categoryID, description, picture, listOrder, categoryKey)
			VALUES(38, 'Succession Planning', '', 10, 'SUCCESSION')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 38

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00080088087700000000887787BB70700000877887BBB7B700088877787BBBB7008777777807BBB700877877778BBBB700877877888660060087778778000FF0000877777800FF000000888880000FFF000000000077FFFF0000000000770FFFF000000000770FF000000000007770FF0000000000077777000000000000000000B23F0000C0170000C0030000800300000083000000030000000100000083000080830000C1010000FE010000FE000000FE010000FE010000FF010000FF83000000

			DELETE FROM ASRSysPermissionItems WHERE itemid in (134,135,136,137,138)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (134,'New',10,38,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (135,'Edit',20,38,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (136,'View',30,38,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (137,'Delete',40,38,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (138,'Run',50,38,'RUN')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (134,135,136,137,138)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 134, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 135, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 136, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 137, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 138, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 8 of 120 - Updating System Permissions for Career Progression'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 39

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			INSERT INTO ASRSysPermissionCategories
			(categoryID, description, picture, listOrder, categoryKey)
			VALUES(39, 'Career Progression', '', 10, 'CAREER')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 39

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000660060770000000000FF07BB707000000FF007BBB7B7000000FFF07BBBB700077FFFF007BBB7000770FFFF0BBBB7000770FF008778808007770FF00887780000777770087788000000000088877780000000087777778000000008778777780000000877877880000000087778778000000000877777800000000008888800803F0000C0170000C0030000800300000083000000030000000500000083000080830000C1010000FE010000FE000000FE010000FE010000FF010000FF83000000

			DELETE FROM ASRSysPermissionItems WHERE itemid in (139,140,141,142,143)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (139,'New',10,39,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (140,'Edit',20,39,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (141,'View',30,39,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (142,'Delete',40,39,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (143,'Run',50,39,'RUN')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (139,140,141,142,143)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 139, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 140, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 141, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 142, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 143, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 9 of 120 - Adding Tables for Calendar Reports'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysCalendarReports]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysCalendarReports] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[Name] [varchar] (50) NULL ,
						[Description] [varchar] (255) NULL ,
						[BaseTable] [int] NULL ,
						[AllRecords] [bit] NOT NULL ,
						[Picklist] [int] NULL ,
						[Filter] [int] NULL ,
						[Access] [varchar] (2) NULL ,
						[UserName] [varchar] (50) NULL ,
						[Description1] [int] NOT NULL ,
						[Description2] [int] NULL ,
						[Region] [int] NULL ,
						[GroupByDesc] [bit] NOT NULL ,
						[StartType] [int] NOT NULL ,
						[FixedStart] [datetime] NULL ,
						[StartFrequency] [int] NULL ,
						[StartPeriod] [int] NULL ,
						[EndType] [int] NOT NULL ,
						[FixedEnd] [datetime] NULL ,
						[EndFrequency] [int] NULL ,
						[EndPeriod] [int] NULL ,
						[ShowBankHolidays] [bit] NOT NULL ,
						[ShowCaptions] [bit] NOT NULL ,
						[ShowWeekends] [bit] NOT NULL ,
						[IncludeWorkingDaysOnly] [bit] NOT NULL ,
						[OutputPreview] [bit] NOT NULL ,
						[OutputFormat] [int] NOT NULL ,
						[OutputScreen] [bit] NOT NULL ,
						[OutputPrinter] [bit] NOT NULL ,
						[OutputPrinterName] [varchar] (255) NOT NULL ,
						[OutputSave] [bit] NOT NULL ,
						[OutputSaveExisting] [int] NOT NULL ,
						[OutputEmail] [bit] NOT NULL ,
						[OutputEmailAddr] [int] NOT NULL ,
						[OutputEmailSubject] [varchar] (255) NOT NULL ,
						[OutputFilename] [varchar] (255) NOT NULL ,
						[Timestamp] [timestamp] NULL 
					) ON [PRIMARY] '
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysCalendarReportEvents]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysCalendarReportEvents]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysCalendarReportEvents] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[EventKey] [varchar] (50) NOT NULL ,
						[CalendarReportID] [int] NOT NULL ,
						[Name] [varchar] (50) NULL ,
						[TableID] [int] NOT NULL ,
						[FilterID] [int] NOT NULL ,
						[EventStartDateID] [int] NOT NULL ,
						[EventStartSessionID] [int] NULL ,
						[EventEndDateID] [int] NOT NULL ,
						[EventEndSessionID] [int] NULL ,
						[EventDurationID] [int] NULL ,
						[LegendType] [int] NULL ,
						[LegendCharacter] [char] (1) NULL ,
						[LegendLookupTableID] [int] NULL ,
						[LegendLookupColumnID] [int] NULL ,
						[LegendLookupCodeID] [int] NULL ,
						[LegendEventColumnID] [int] NULL ,
						[EventDesc1ColumnID] [int] NULL ,
						[EventDesc2ColumnID] [int] NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysCalendarReportOrder]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysCalendarReportOrder]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysCalendarReportOrder] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[CalendarReportID] [int] NOT NULL ,
						[TableID] [int] NOT NULL ,
						[ColumnID] [int] NOT NULL ,
						[OrderSequence] [int] NOT NULL ,
						[OrderType] [varchar] (4) NOT NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 10 of 120 - Adding IncludeBankHolidays Column to Calendar Reports'

		/* Adding Include Bank Holidays Column to Calendar Reports definition table */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCalendarReports')
		and name = 'IncludeBankHolidays'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCalendarReports ADD
						[IncludeBankHolidays] [bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCalendarReports SET
						[IncludeBankHolidays] = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 11 of 120 - Adding PrintFilterHeader Column to Calendar Reports'

		/* Adding PrintFilterHeader Column to Calendar Reports definition table */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCalendarReports')
		and name = 'PrintFilterHeader'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCalendarReports ADD
						[PrintFilterHeader] [bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCalendarReports SET
						[PrintFilterHeader] = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 12 of 120 - Adding Calculation Columns to Calendar Reports'

		/* Adding Calculation Columns to Calendar Reports definition table */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCalendarReports')
		and name = 'StartDateExpr'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCalendarReports ADD
						[StartDateExpr] [integer] NULL,
						[EndDateExpr] [integer] NULL,
						[DescriptionExpr] [integer] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCalendarReports SET
						[StartDateExpr] = 0,
						[EndDateExpr] = 0,
						[DescriptionExpr] = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 13 of 120 - Adding Delete Trigger for Calendar Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysCalendarReports]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysCalendarReports]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCalendarReports ON dbo.ASRSysCalendarReports 
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysCalendarReportEvents WHERE ASRSysCalendarReportEvents.CalendarReportID IN (SELECT ID FROM deleted)
    						DELETE FROM ASRSysCalendarReportOrder WHERE ASRSysCalendarReportOrder.CalendarReportID IN (SELECT ID FROM deleted)
					END'
		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 14 of 120 - Updating System Permissions for Calendar Reports'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 24

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(24,
					'Calendar Reports',
					'',
					10,
					'CALENDARREPORTS')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 24

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000BBB0000BB000000BB0BB000BB000000000BB000BB000000000BB000BB08FF7F0BBB00B0BB08FF7F000BB0BBBB087770BB0BB00BBB08FF7F0BBB0FF0BB08FF7FF0001FF000087777711117700008FF7FF7FF7FF00008FF7FF7FF7FF0000844448888888000084444888888800008888888888880000FC790000F8300000F0100000F810000000000000000000000000000000000000000000000001000000070000000700000007000000070000000700000007000000


			DELETE FROM ASRSysPermissionItems WHERE itemid in (106,107,108,109,110)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (106,'New',10,24,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (107,'Edit',20,24,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (108,'View',30,24,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (109,'Delete',40,24,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (110,'Run',50,24,'RUN')


			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (106,107,108,109,110)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 106, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 107, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 108, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 109, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 110, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
		END


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 120 - Adding Tables for Record Profile'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysRecordProfileName]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysRecordProfileName] (
						[RecordProfileID] [int] IDENTITY (1, 1) NOT NULL ,
						[Name] [varchar] (50) NULL ,
						[Description] [varchar] (255) NULL ,
						[BaseTable] [int] NULL ,
						[AllRecords] [bit] NOT NULL ,
						[PicklistID] [int] NULL ,
						[FilterID] [int] NULL ,
						[Access] [varchar] (2) NULL ,
						[UserName] [varchar] (50) NULL ,
						[DefaultOutput] [int] NULL ,
						[DefaultExportTo] [int] NULL ,
						[DefaultSave] [bit] NOT NULL ,
						[DefaultSaveAs] [varchar] (255) NULL ,
						[DefaultCloseApp] [bit] NOT NULL ,
						[TimeStamp] [timestamp] NULL ,
						[OrderID] [int] NULL ,
						[Orientation] [int] NULL ,
						[PageBreak] [bit] NULL 
					) ON [PRIMARY] '
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysRecordProfileDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysRecordProfileDetails]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysRecordProfileDetails] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[RecordProfileID] [int] NOT NULL ,
						[Sequence] [int] NULL ,
						[Type] [char] (1) NULL ,
						[ColumnID] [int] NULL ,
						[Heading] [varchar] (50) NULL ,
						[Size] [int] NULL ,
						[DP] [int] NULL ,
						[IsNumeric] [bit] NULL ,
						[TableID] [int] NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

			if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysRecordProfileTables]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
			drop table [dbo].[ASRSysRecordProfileTables]

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysRecordProfileTables] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[RecordProfileID] [int] NOT NULL ,
						[TableID] [int] NOT NULL ,
						[FilterID] [int] NULL ,
						[OrderID] [int] NULL ,
						[MaxRecords] [int] NULL ,
						[Orientation] [int] NULL ,
						[PageBreak] [bit] NULL ,
						[Sequence] [int] NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 16 of 120 - Adding new columns for Record Profile'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysRecordProfileName]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN
			SELECT @iRecCount = count(id) FROM syscolumns
			where id = (select id from sysobjects where name = 'ASRSysRecordProfileName')
				and name = 'DefaultOutput'

			if @iRecCount = 1
			BEGIN	
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
							DROP COLUMN [DefaultOutput] '
				EXEC sp_executesql @NVarCommand
				
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
							DROP COLUMN [DefaultExportTo] '
				EXEC sp_executesql @NVarCommand

				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
							DROP COLUMN [DefaultSave] '
				EXEC sp_executesql @NVarCommand					
				
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
							DROP COLUMN [DefaultSaveAs] '
				EXEC sp_executesql @NVarCommand			  
				
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
							DROP COLUMN [DefaultCloseApp] '
				EXEC sp_executesql @NVarCommand
			END

			SELECT @iRecCount = count(id) FROM syscolumns
			where id = (select id from sysobjects where name = 'ASRSysRecordProfileName')
				and name = 'OutputPreview'

			if @iRecCount = 0
			BEGIN
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] 
						ADD  
						[OutputPreview] [bit] NULL , 
						[OutputFormat] [int] NULL ,
						[OutputScreen] [bit] NULL ,
						[OutputPrinter] [bit] NULL ,
						[OutputPrinterName] [varchar] (255) NULL ,
						[OutputSave] [bit] NULL ,
						[OutputSaveExisting] [int] NULL ,
						[OutputEmail] [bit] NULL ,
						[OutputEmailAddr] [int] NULL ,
						[OutputEmailSubject] [varchar] (255) NULL ,
						[OutputFilename] [varchar] (255) NULL ,
						[IndentRelatedTables] [bit] NULL ,
						[SuppressEmptyRelatedTableTitles] [bit] NULL ,
						[SuppressTableRelationshipTitles] [bit] NULL 
						'
				EXEC sp_executesql @NVarCommand
			END


			SELECT @iRecCount = count(id) FROM syscolumns
			where id = (select id from sysobjects where name = 'ASRSysRecordProfileName')
			and name = 'PrintFilterHeader'

			if @iRecCount = 0
			BEGIN
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysRecordProfileName] ADD
							[PrintFilterHeader] [bit] NULL'
				EXEC sp_executesql @NVarCommand
				SELECT @NVarCommand = 'UPDATE ASRSysRecordProfileName SET
							[PrintFilterHeader] = 0'
				EXEC sp_executesql @NVarCommand
			END
		END



/* ------------------------------------------------------------- */
PRINT 'Step 17 of 120 - Adding Delete Trigger for Record Profile'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysRecordProfileName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysRecordProfileName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysRecordProfileName ON dbo.ASRSysRecordProfileName 
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysRecordProfileDetails WHERE ASRSysRecordProfileDetails.RecordProfileID IN (SELECT ID FROM deleted)
    						DELETE FROM ASRSysRecordProfileTables WHERE ASRSysRecordProfileTables.RecordProfileID IN (SELECT ID FROM deleted)
					END'
		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 18 of 120 - Updating System Permissions for Record Profile'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 34

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(34,
					'Record Profile',
					'',
					10,
					'RECORDPROFILE')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 34

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000FFFFFFFFFFF000000000000000F0000FFFFFF00770F0000FFFFFF0F070F0000F0F7770FF00F0000FFFFFF00000F0000F0F777FFFF0F0000FFFFFFFFFF0F0000F0F777777F0F0000FFFFFFFFFF0F0000F0F77FCCCF0F0000FFFFFFC7CF0F0000F0F77FCCCF0F0000FFFFFFFFFF00000000000000000000C0010000C00100008001000080010000800100008001000080010000800100008001000080010000800100008001000080010000800100008001000080070000


			DELETE FROM ASRSysPermissionItems WHERE itemid in (120,121,122,123,124)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (120,'New',10,34,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (121,'Edit',20,34,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (122,'View',30,34,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (123,'Delete',40,34,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (124,'Run',50,34,'RUN')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (120,121,122,123,124)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 120, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 121, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 122, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 123, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 124, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 19 of 120 - Creating Definition Access for Record Profile'

		SELECT @iRecCount = count(sysobjects.id)
		FROM sysobjects 
		WHERE name = 'ASRSysRecordProfileAccess'

		IF @iRecCount = 0 
		BEGIN
			CREATE TABLE [dbo].[ASRSysRecordProfileAccess] (
				[GroupName] [varchar] (256) NOT NULL ,
				[Access] [varchar] (2) NOT NULL ,
				[ID] [int] NOT NULL 
			) ON [PRIMARY]


			SELECT @NVarCommand = 'INSERT INTO ASRSysRecordProfileAccess 
				(groupName, access, id)
				(SELECT sysusers.name,
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
					ELSE
						ASRSysRecordProfileName.access
				END,
				recordProfileID
			FROM ASRSysRecordProfileName,
				sysusers
			WHERE sysusers.uid = sysusers.gid
				and sysusers.uid <> 0)'

			exec sp_sqlexec @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 20 of 120 - Adding Tables for Envelopes & Labels'

		/* Adding Columns to Mail Merge required for Envelopes & Labels */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysMailMergeName')
		and name = 'IsLabel'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysMailMergeName ADD 
						[IsLabel] [bit] NULL ,
						[LabelTypeID] [int] NULL ,
						[PromptStart] [int] NULL '
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysMailMergeName SET
						IsLabel = 0'
			EXEC sp_executesql @NVarCommand
		 
			SELECT @NVarCommand = 'ALTER TABLE ASRSysMailMergeColumns ADD 
						[ColumnOrder] [int] NULL ,
						[StartOnNewLine] [bit] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysMailMergeColumns SET
						StartOnNewLine = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 21 of 120 - Updating System Permissions for Envelopes & Labels'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 29

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(29,
					'Envelopes & Labels',
					'',
					10,
					'LABELS')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 29

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x424DB60300000000000036000000280000001200000010000000010018000000000080030000130B0000130B00000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4900FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4F00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6500FFFFFFFFFFFF4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4DFFFFFF7200FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF6D00FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFA64D4DA64D4DA64D4DA64D4DA64D4DFFFFFFFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF4E00FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF5C00FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFA64D4DA64D4DA64D4DA64D4DA64D4DA64D4DFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF6F00FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF7200FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4DA64D4DA6FFFFFF4D4D4DFFFFFF6D00FFFFFFFFFFFF4D4D4DFFFFFF4D4D4D4D4D4D4D4D4D4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4DA64D4DA6FFFFFF4D4D4DFFFFFF5C00FFFFFFFFFFFF4D4D4DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4D4D4DFFFFFF6500FFFFFFFFFFFF4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4DFFFFFF7300FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6400FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6F00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF5C00

			DELETE FROM ASRSysPermissionItems WHERE itemid in (111,112,113,114,115)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (115,'New',10,29,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (111,'Edit',20,29,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (112,'View',30,29,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (113,'Delete',40,29,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (114,'Run',50,29,'RUN')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (111,112,113,114,115)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 111, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 112, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 113, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 114, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 115, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 22 of 120 - Adding Tables for Label Definition'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysLabelTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysLabelTypes] (
						[LabelTypeID] [int] IDENTITY (1, 1) NOT NULL ,
						[Type] [varchar] (25) NOT NULL ,
						[LabelHeight] [float] NOT NULL ,
						[LabelWidth] [float] NOT NULL ,
						[TopMargin] [float] NOT NULL ,
						[SideMargin] [float] NOT NULL ,
						[VerticalPitch] [float] NOT NULL ,
						[HorizontalPitch] [float] NOT NULL ,
						[NumberAcross] [int] NOT NULL ,
						[NumberDown] [int] NOT NULL ,
						[ASRDefined] [bit] NOT NULL ,
						[LabelSupplier] [varchar] (25) NULL ,
						[Name] [varchar] (50) NOT NULL ,
						[Description] [varchar] (255) NOT NULL ,
						[Access] [varchar] (2) NOT NULL ,
						[Username] [varchar] (50) NOT NULL ,
						[Timestamp] [timestamp] NOT NULL ,
						[PageOrientation] [bit] NOT NULL ,
						[PageWidth] [float] NOT NULL ,
						[PageHeight] [float] NOT NULL ,
						[PageTypeID] [int] NOT NULL 
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

		END

/* ------------------------------------------------------------- */
PRINT 'Step 23 of 120 - Updating System Permissions for Label Definition'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 30

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(30,
					'Label Definition',
					'',
					10,
					'LABELDEFINITION')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 30

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x424D360300000000000036000000280000001000000010000000010018000000000000030000230B0000230B00000000000000000000FFFFFFFFFFFFA5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5A5FFFFFF4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFF0000FF0000FF0000FF00004A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFF0000FF0000FF0000FF00004A4A4A4AFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4AFFFF4A4A4A4AFFFFFFFFFF4A4AA5FFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4AFFFF4AFFFF4A4A4AFFFFFF4A4AA5FFFFFF4A4A4AA5A5A5FFFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4AFFFF4AFFFF4A4A4AFFFFFFFFFFFFFFFFFF4A4A4AA5A5A5FFFFFF4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4AFFFF4AFFFF4A4A4A4A4A4A4A4A4A4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4AFFFF4AFFFF4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4A4A4A4A4A4A4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4A4AA54A4AA54A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4A4A4A4A4A4AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF


			DELETE FROM ASRSysPermissionItems WHERE itemid in (116,117,118,119)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (116,'New',10,30,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (117,'Edit',20,30,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (118,'View',30,30,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (119,'Delete',40,30,'DELETE')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (116,117,118,119)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 116, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 117, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 118, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 119, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 24 of 120 - Adding Tables for Email Group Definition'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysEmailGroupName]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN
			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysEmailGroupName] (
					[EmailGroupID] [int] IDENTITY (1, 1) NOT NULL ,
					[Name] [varchar] (50) NOT NULL ,
					[Description] [varchar] (255) NULL ,
					[UserName] [varchar] (50) NOT NULL ,
					[Access] [varchar] (2) NOT NULL,
					[TimeStamp] [timestamp]
					) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand
		END

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysEmailGroupItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN
			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysEmailGroupItems] (
						[EmailGroupID] [int] NOT NULL ,
						[EmailDefID] [int] NOT NULL 
						) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 25 of 120 - Updating System Permissions for Email Groups'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 35

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			INSERT INTO ASRSysPermissionCategories
			(categoryID, description, picture, listOrder, categoryKey)
			VALUES
			(35, 'Email Groups', '', 10, 'EMAILGROUPS')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 35

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000100000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00000000000000000000600060000000000000FF0000000000000FF000050000000000FFF070000000077FFFF0700000000770FFFF070000000770FF007700000007770FF07770000000777770700FFFB000000008770FBF80000008888808F8B00000000000FF8FF0000000B8FFBFF8B00000008FBFFFBF800000000000000000FFFFFFFF80FFFFFFC1FFFFFFC01FFFFF803FFFFF003FFFFF001FFFFF001FFFFF0000FFFF8000FFFFC000FFFFF000FFFFF800FFFFF800FFFFF800FFFFF800FFFF00


			DELETE FROM ASRSysPermissionItems WHERE itemid in (125,126,127,128)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (125,'New',10,35,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (126,'Edit',20,35,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (127,'View',30,35,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (128,'Delete',40,35,'DELETE')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (125,126,127,128)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 125, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 126, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 127, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 128, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 26 of 120 - Updating Email Address Definitions'

		DECLARE @EmailID int
		DECLARE @Fixed varchar(8000)
		DECLARE @SQL varchar(8000)

		DECLARE HRProCursor CURSOR
		FOR SELECT DISTINCT Fixed FROM ASRSysEmailAddress WHERE Type = 0

		set nocount on

		OPEN HRProCursor
		FETCH NEXT FROM HRProCursor INTO @Fixed
		WHILE @@FETCH_STATUS = 0
		BEGIN

			SELECT TOP 1 @EmailID = EmailID FROM ASRSysEmailAddress WHERE Fixed = @Fixed ORDER BY EmailID

			SELECT @SQL = 'UPDATE ASRSysTables SET DefaultEmailID = ' + CONVERT(varchar,@EmailID) + ' WHERE DefaultEmailID IN
					(SELECT EmailID FROM ASRSysEmailAddress WHERE lower(Fixed) = lower(''' + replace(@Fixed,'''','''''') + '''))'
			EXECUTE sp_sqlexec @SQL

			SELECT @SQL = 'UPDATE ASRSysEmailLinksRecipients SET RecipientID = ' + CONVERT(varchar,@EmailID) + ' WHERE RecipientID IN
					(SELECT EmailID FROM ASRSysEmailAddress WHERE lower(Fixed) = lower(''' + replace(@Fixed,'''','''''') + '''))'
			EXECUTE sp_sqlexec @SQL

			SELECT @SQL = 'UPDATE ASRSysMailMergeName SET EmailAddrID = ' + CONVERT(varchar,@EmailID) + ' WHERE EmailAddrID IN
					(SELECT EmailID FROM ASRSysEmailAddress WHERE lower(Fixed) = lower(''' + replace(@Fixed,'''','''''') + '''))'
			EXECUTE sp_sqlexec @SQL

			SELECT @SQL = 'DELETE FROM ASRSysEmailAddress WHERE EmailID <> ' + CONVERT(varchar,@EmailID) + ' AND EmailID IN
					(SELECT EmailID FROM ASRSysEmailAddress WHERE lower(Fixed) = lower(''' + replace(@Fixed,'''','''''') + '''))'
			EXECUTE sp_sqlexec @SQL

			FETCH NEXT FROM HRProCursor INTO @Fixed
		END

		CLOSE HRProCursor
		DEALLOCATE HRProCursor

		set nocount off

		SELECT @SQL = 'UPDATE ASRSysEmailAddress SET TableID = 0 WHERE Type = 0'
		EXECUTE sp_sqlexec @SQL



/* ------------------------------------------------------------- */
PRINT 'Step 27 of 120 - Updating System Permissions for Email Addresses'

		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 36

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
			(categoryID, description, picture, listOrder, categoryKey)
			VALUES
			(36, 'Email Addresses', '', 10, 'EMAILADDRESSES')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 36

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000100000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00000000000000000000600060000000000000FF0000000000000FF000000000000000FFF000000000077FFFF0000000000770FFFF000000000770FF000000000007770FF00000000000777770FFBFFFB00000000FBFFFBF80000000F8F888F8B0000000FF8FFF8FF0000000B8FFBFF8B00000008FBFFFBF800000000000000000FFFFFFFF80FFFFFFC1FFFFFFC1FFFFFF80FFFFFF00FFFFFF007FFFFF00FFFFFF0000FFFF8000FFFFC000FFFFF800FFFFF800FFFFF800FFFFF800FFFFF800FFFF00


			DELETE FROM ASRSysPermissionItems WHERE itemid in (129,130,131,132)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (129,'New',10,36,'NEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (130,'Edit',20,36,'EDIT')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (131,'View',30,36,'VIEW')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (132,'Delete',40,36,'DELETE')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (129,130,131,132)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 129, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 130, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 131, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 132, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END

		UPDATE ASRSysPermissionCategories SET Description = 'Email Queue' WHERE Description = 'Email Generation'



/* ------------------------------------------------------------- */
PRINT 'Step 28 of 120 - Updating System Permissions for Email Queue'

		/* Updating System Permissions Picture Icon for Email Queue */
		DELETE FROM ASRSysPermissionCategories WHERE categoryID = 18

		--SET IDENTITY_INSERT ASRSysPermissionCategories ON

		/* The record doesn't exist, so create it. */
		INSERT INTO ASRSysPermissionCategories
			(categoryID, 
				description, 
				picture, 
				listOrder, 
				categoryKey)
			VALUES(18,
				'Email Queue',
				'',
				10,
				'EMAIL')

		--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

		SELECT @ptrval = TEXTPTR(picture) 
		FROM ASRSysPermissionCategories
		WHERE categoryID = 18

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000BFFFBFFFB00000008FBFFFBF80000000F8F888F8B0000000FF8FFF8FF0000000B8FFBFF8B0FB00008FBFFFBF80F8000000000000008B0000000FF8FFF8FF0000000B8FFBFF8B0B000008FBFFFBF808000000000000000B0000000FF8FFF8FF0000000B8FFBFF8B00000008FBFFFBF800000000000000000001F0000001F0000001F0000001F000000030000000300000003000000030000E0000000E0000000E0000000E0000000F8000000F8000000F8000000F800000000



/* ------------------------------------------------------------- */
PRINT 'Step 29 of 120 - Adding new columns for Cross Tabs'

		/* Adding Columns to Cross Tabs required for new functionality */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
		and name = 'OutputPreview'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab ADD
						[OutputPreview] [bit] NULL,
						[OutputFormat] [int] NULL,
						[OutputScreen] [bit] NULL,
						[OutputPrinter] [bit] NULL,
						[OutputPrinterName] [varchar] (255) NULL,
						[OutputSave] [bit] NULL,
						[OutputSaveExisting] [int] NULL,
						[OutputEmail] [bit] NULL,
						[OutputEmailAddr] [int] NULL,
						[OutputEmailSubject] [varchar] (255) NULL,
						[OutputFilename] [varchar] (255) NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET
						[OutputPreview] = 1,
						[OutputFormat] = 0,
						[OutputScreen] = CASE WHEN [DefaultCloseApp]=1 THEN 0 ELSE 1 END,
						[OutputPrinter] = 0,
						[OutputPrinterName] = '''',
						[OutputSave] = [DefaultSave],
						[OutputSaveExisting] = 0,
						[OutputEmail] = 0,
						[OutputEmailAddr] = 0,
						[OutputEmailSubject] = '''',
						[OutputFilename] = [DefaultSaveAs]'
			EXEC sp_executesql @NVarCommand
			
			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET [OutputFormat] = 4 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 1'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET [OutputFormat] = 3 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 2'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET [OutputFormat] = 2 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET [OutputPreview] = 0, [OutputPrinter] = 1 WHERE [DefaultOutput] = 0'
			EXEC sp_executesql @NVarCommand

		END



/* ------------------------------------------------------------- */
PRINT 'Step 30 of 120 - Removing old columns for Cross Tabs'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
			and name = 'DefaultOutput'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET [OutputScreen] = 1 WHERE [OutputFormat] = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP COLUMN DefaultOutput'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
			and name = 'DefaultExportTo'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP COLUMN DefaultExportTo'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM sysobjects
		WHERE name = 'DF_ASRSysCrossTab_DefaultSave' AND type = 'D'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP CONSTRAINT DF_ASRSysCrossTab_DefaultSave'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
			and name = 'DefaultSave'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP COLUMN DefaultSave'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
			and name = 'DefaultSaveAs'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP COLUMN DefaultSaveAs'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM sysobjects 
		WHERE name = 'DF_ASRSysCrossTab_DefaultCloseApp' AND type = 'D'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP CONSTRAINT DF_ASRSysCrossTab_DefaultCloseApp'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCrossTab')
			and name = 'DefaultCloseApp'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCrossTab
								DROP COLUMN DefaultCloseApp'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 31 of 120 - Adding new columns for Custom Reports'

		/* Adding Columns to Custom Reports required for new functionality */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
		and name = 'OutputPreview'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName ADD
						[OutputPreview] [bit] NULL,
						[OutputFormat] [int] NULL,
						[OutputScreen] [bit] NULL,
						[OutputPrinter] [bit] NULL,
						[OutputPrinterName] [varchar] (255) NULL,
						[OutputSave] [bit] NULL,
						[OutputSaveExisting] [int] NULL,
						[OutputEmail] [bit] NULL,
						[OutputEmailAddr] [int] NULL,
						[OutputEmailSubject] [varchar] (255) NULL,
						[OutputFilename] [varchar] (255) NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET
						[OutputPreview] = 1,
						[OutputFormat] = 0,
						[OutputScreen] = CASE WHEN [DefaultCloseApp]=1 THEN 0 ELSE 1 END,
						[OutputPrinter] = 0,
						[OutputPrinterName] = '''',
						[OutputSave] = [DefaultSave],
						[OutputSaveExisting] = 0,
						[OutputEmail] = 0,
						[OutputEmailAddr] = 0,
						[OutputEmailSubject] = '''',
						[OutputFilename] = [DefaultSaveAs]'
			EXEC sp_executesql @NVarCommand


			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET [OutputFormat] = 4 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 1'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET [OutputFormat] = 3 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 2'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET [OutputFormat] = 2 WHERE [DefaultOutput] = 1 AND [DefaultExportTo] = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET [OutputPreview] = 0, [OutputPrinter] = 1 WHERE [DefaultOutput] = 0'
			EXEC sp_executesql @NVarCommand


			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsDetails ADD 
						[Hidden] [bit] NULL ,
						[GroupWithNextColumn] [bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsDetails SET
						Hidden = 0,
						GroupWithNextColumn = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsChildDetails ADD 
						[ChildOrder] [int] NULL '
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsChildDetails SET
						ChildOrder = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 32 of 120 - Updating System Permissions for Cross Tabs'

		/* Changing name of "Cross-tabs" Permission Category */
		/*
		SELECT @NVarCommand = 'Update ASRSysPermissionCategories Set Description = ''Cross Tabs'' Where categoryID = 3'
		EXEC sp_executesql @NVarCommand
		*/

		/* Updating System Permissions Picture Icon for Cross Tabs */
		DELETE FROM ASRSysPermissionCategories WHERE categoryID = 3

		--SET IDENTITY_INSERT ASRSysPermissionCategories ON

		INSERT INTO ASRSysPermissionCategories
			(categoryID, 
				description, 
				picture, 
				listOrder, 
				categoryKey)
			VALUES(3,
				'Cross Tabs',
				'',
				10,
				'CROSSTABS')

		--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

		SELECT @ptrval = TEXTPTR(picture) 
		FROM ASRSysPermissionCategories
		WHERE categoryID = 3

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000FFFFFFCFFFFFF000F800FCCCFFFFF000F8B0CCCCCFFFF000F8B0FFCFFFFFF000F8B0FFCFFFCFF000F8B0FFCFFFCCF000F8B0FFFCCCCCC000F888FFFFFFCCF000FFFFFFFFFFCFF000F800F8000000F000F8B0F8BBBBB0F000F888F8888888F000FFFFFFFFFFFFF000000000000000000FFFF000000010000000100000001000000010000000100000001000000010000000100000001000000010000000100000001000000010000000100000001000000



/* ------------------------------------------------------------- */
PRINT 'Step 33 of 120 - Removing old columns for Custom Reports'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
			and name = 'DefaultOutput'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsName SET [OutputScreen] = 1 WHERE [OutputFormat] = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP COLUMN DefaultOutput'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
			and name = 'DefaultExportTo'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP COLUMN DefaultExportTo'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM sysobjects
		WHERE name = 'DF_ASRSysQuickReportsName_DefaultSave' AND type = 'D'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP CONSTRAINT DF_ASRSysQuickReportsName_DefaultSave'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
			and name = 'DefaultSave'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP COLUMN DefaultSave'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
			and name = 'DefaultSaveAs'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP COLUMN DefaultSaveAs'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM sysobjects 
		WHERE name = 'DF_ASRSysQuickReportsName_DefaultCloseApp' AND type = 'D'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP CONSTRAINT DF_ASRSysQuickReportsName_DefaultCloseApp'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
			and name = 'DefaultCloseApp'

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsName
								DROP COLUMN DefaultCloseApp'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 34 of 120 - Adding new columns for Export'

		/* Adding Columns to Export required for new functionality */
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysExportName')
		and name = 'OutputFormat'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysExportName ADD
						[OutputFormat] [int] NULL,
						[OutputSave] [bit] NULL,
						[OutputSaveExisting] [int] NULL,
						[OutputEmail] [bit] NULL,
						[OutputEmailAddr] [int] NULL,
						[OutputEmailSubject] [varchar] (255) NULL,
						[OutputFilename] [varchar] (255) NULL'
			EXEC sp_executesql @NVarCommand
		END


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysExportDetails')
		and name = 'Heading'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysExportDetails ADD
						[Heading] [varchar] (50) NULL'

			EXEC sp_executesql @NVarCommand
		END


		SELECT @NVarCommand = 'UPDATE ASRSysExportName SET
					[OutputSave] = 1,
					[OutputSaveExisting] = case when appendtofile = 1 then 3 else 0 end,
					[OutputEmail] = 0,
					[OutputEmailAddr] = 0,
					[OutputEmailSubject] = '''',
					[OutputFilename] = [OutputName],
					[OutputFormat] = case
					when outputtype = ''D'' then 1
					when outputtype = ''F'' then 7
					when outputtype = ''C'' then 8
					when outputtype = ''S'' then 99
					else 0 end 
				WHERE [OutputFormat] IS NULL'

		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 35 of 120 - Adding new column to Import'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysImportDetails')
		and name = 'LookupEntries'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysImportDetails] ADD [LookupEntries] [bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysImportDetails SET LookupEntries = 0'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 36 of 120 - Adding new columns to Batch Jobs'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysBatchJobName')
		and name = 'LockSpid'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysBatchJobName] ADD
						[LockSpid] [integer] NULL,
						[LockLoginTime] [datetime] NULL'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 37 of 120 - Creating Definition Access for Batch Jobs'

		SELECT @iRecCount = count(sysobjects.id)
		FROM sysobjects 
		WHERE name = 'ASRSysBatchJobAccess'

		IF @iRecCount = 0 
		BEGIN
			CREATE TABLE [dbo].[ASRSysBatchJobAccess] (
				[GroupName] [varchar] (256) NOT NULL ,
				[Access] [varchar] (2) NOT NULL ,
				[ID] [int] NOT NULL 
			) ON [PRIMARY]

			SELECT @NVarCommand = 'INSERT INTO ASRSysBatchJobAccess
				(groupName, access, id)
				(SELECT sysusers.name,
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
					ELSE
						ASRSysBatchJobName.access
				END,
				ID
			FROM ASRSysBatchJobName,
				sysusers
			WHERE sysusers.uid = sysusers.gid
				and sysusers.uid <> 0)'

			exec sp_sqlexec @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 38 of 120 - Adding new procedure for Batch Jobs'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRLockWriteBatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRLockWriteBatch]

		EXEC('CREATE Procedure spASRLockWriteBatch (@BatchJobID int, @LockedByOther int OUTPUT)
		AS
		BEGIN

			DECLARE @OrigTranCount int
			DECLARE @Realspid int

			SET @OrigTranCount = @@trancount
			IF @OrigTranCount = 0 BEGIN TRANSACTION

			SELECT @LockedByOther = COUNT(ID) FROM ASRSysBatchJobName
			JOIN master..sysprocesses syspro ON spid = LockSpid
			WHERE LockLoginTime = syspro.login_time AND LockSpid <> @@spid
			AND ID = @BatchJobID

			IF @LockedByOther = 0
			BEGIN

				--Need to get spid of parent process
				SELECT @Realspid = a.spid
				FROM master..sysprocesses a
				FULL OUTER JOIN master..sysprocesses b
					ON a.hostname = b.hostname
					AND a.hostprocess = b.hostprocess
					AND a.spid <> b.spid
				WHERE b.spid = @@Spid

				--If there is no parent spid then use current spid
				--IF @Realspid is null SET @Realspid = @@spid


				UPDATE ASRSysBatchJobName SET
				LockSpid = @Realspid,
				LockLoginTime = (
					SELECT login_time
					FROM master..sysprocesses
					WHERE spid = @Realspid)
				WHERE ID = @BatchJobID

			END

			IF @OrigTranCount = 0 COMMIT TRANSACTION

		END')		



/* ------------------------------------------------------------- */
PRINT 'Step 39 of 120 - Amending Delete Trigger for Event Log'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeEventLog]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_ASRSysPurgeEventLog]

		EXEC('CREATE TRIGGER INS_ASRSysPurgeEventLog ON ASRSysEventLog 
		FOR INSERT 
		AS 
		DECLARE @intFrequency int, 
			@strPeriod char(2) 

		SELECT @intFrequency = Frequency 
		FROM ASRSysEventLogPurge 

		SELECT @strPeriod = Period 
		FROM ASRSysEventLogPurge 

		IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL) 
		BEGIN 
			IF @strPeriod = ''dd'' 
			BEGIN 
				DELETE FROM ASRSysEventLog 
				WHERE [DateTime] < DATEADD(dd,-@intfrequency,getdate()) 
			END 

			IF @strPeriod = ''wk''
			BEGIN 
				DELETE FROM ASRSysEventLog 
				WHERE [DateTime] < DATEADD(wk,-@intfrequency,getdate()) 
			END 

			IF @strPeriod = ''mm'' 
			BEGIN 
				DELETE FROM ASRSysEventLog 
				WHERE [DateTime] < DATEADD(mm,-@intfrequency,getdate()) 
			END 

			IF @strPeriod = ''yy'' 
			BEGIN 
				DELETE FROM ASRSysEventLog 
				WHERE [DateTime] < DATEADD(yy,-@intfrequency,getdate()) 
			END 

			DELETE FROM ASRSysEventLogDetails 
			WHERE [EventLogID] NOT IN (SELECT ID FROM AsrSysEventLog) 
		END')



/* ------------------------------------------------------------- */
PRINT 'Step 40 of 120 - Adding new columns for Functions'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysFunctions')
		and name = 'UDF'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysFunctions ADD [UDF] [bit]'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysFunctions SET udf = 0'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysFunctions SET udf = 1 WHERE functionid IN (12, 20, 30, 46, 47)'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 41 of 120 - Adding Expression Shortcut Keys functionality'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysOperators')
		and LOWER(name) = 'shortcutkeys'

		if @iRecCount = 0
		BEGIN
			
			SELECT @NVarCommand = 'ALTER TABLE ASRSysOperators ADD
						[ShortcutKeys] [varchar] (20) NULL '
			EXEC sp_executesql @NVarCommand
		END

			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''+''
						WHERE OperatorID = 1'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''-''
						WHERE OperatorID = 2'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''*''
						WHERE OperatorID = 3'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''/''
						WHERE OperatorID = 4'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''A''
						WHERE OperatorID = 5'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''O''
						WHERE OperatorID = 6'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''=''
						WHERE OperatorID = 7'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = NULL
						WHERE OperatorID = 8'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''<''
						WHERE OperatorID = 9'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''>''
						WHERE OperatorID = 10'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = NULL
						WHERE OperatorID = 11'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = NULL
						WHERE OperatorID = 12'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''N''
						WHERE OperatorID = 13'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = NULL
						WHERE OperatorID = 14'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''^''
						WHERE OperatorID = 15'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = ''%''
						WHERE OperatorID = 16'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysOperators SET
							[ShortcutKeys] = NULL
						WHERE OperatorID = 17'
			EXEC sp_executesql @NVarCommand


		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysFunctions')
		and lower(name) = 'shortcutkeys'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysFunctions ADD
						[ShortcutKeys] [varchar] (20) NULL '
			EXEC sp_executesql @NVarCommand
		END
			SELECT @NVarCommand = 'UPDATE ASRSysFunctions SET
							[ShortcutKeys] = ''()''
						WHERE FunctionID = 27'
			EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 42 of 120 - Updating function names'

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 22'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF)
       			VALUES                (22, ''Weekdays from Start and End Dates'', 2, 0, ''Date/Time'', ''sp_ASRFn_WeekdaysFromStartAndEndDates'', 0, 1, NULL, 0)'
		EXEC sp_executesql @NVarCommand


		UPDATE ASRSysFunctions SET FunctionName = 'Convert Numeric to Character' where functionid = 3
		UPDATE ASRSysFunctions SET FunctionName = 'Convert to Proper Case' where functionid = 12
		UPDATE ASRSysFunctions SET FunctionName = 'Weekdays between Two Dates' where functionid = 22
		UPDATE ASRSysFunctions SET FunctionName = 'Whole Months between Two Dates' where functionid = 26
		UPDATE ASRSysFunctions SET FunctionName = 'Days between Two Dates' where functionid = 45
		UPDATE ASRSysFunctions SET FunctionName = 'Working Days between Two Dates' where functionid = 46
		UPDATE ASRSysFunctions SET FunctionName = 'Absence between Two Dates' where functionid = 47
		UPDATE ASRSysFunctions SET FunctionName = 'Field Changed between Two Dates' where functionid = 53
		UPDATE ASRSysFunctions SET FunctionName = 'Whole Years between Two Dates' where functionid = 54


/* ------------------------------------------------------------- */
PRINT 'Step 43 of 120 - Adding new Function First Day of Month'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_FirstDayOfMonth]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_FirstDayOfMonth]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_FirstDayOfMonth
					(
					@pdtResult 	datetime OUTPUT,
					@pdtDate 	datetime
					)
					AS
					BEGIN
						SET @pdtResult = dateadd(dd, 1 - datepart(dd, @pdtDate), @pdtDate)
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 55'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (55, ''First Day of Month'', 4, 0, ''Date/Time'', ''sp_ASRFn_FirstDayOfMonth'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 55'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (55, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 44 of 120 - Adding new Function Last Day of Month'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_LastDayOfMonth]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_LastDayOfMonth]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_LastDayOfMonth
					(
					@pdtResult 	datetime OUTPUT,
					@pdtDate 	datetime
					)
					AS
					BEGIN
						SET @pdtResult = dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, @pdtDate), @pdtDate)))
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 56'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (56, ''Last Day of Month'', 4, 0, ''Date/Time'', ''sp_ASRFn_LastDayOfMonth'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 56'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (56, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 45 of 120 - Adding new Function First Day of Year'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_FirstDayOfYear]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_FirstDayOfYear]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_FirstDayOfYear
					(
					@pdtResult 	datetime OUTPUT,	
					@pdtDate 	datetime
					)
					AS
					BEGIN
						SET @pdtResult = dateadd(dd, 1 - datepart(dy, @pdtDate), @pdtDate)
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 57'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (57, ''First Day of Year'', 4, 0, ''Date/Time'', ''sp_ASRFn_FirstDayOfYear'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 57'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (57, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 46 of 120 - Adding new Function Last Day of Year'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_LastDayOfYear]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_LastDayOfYear]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_LastDayOfYear
					(
					@pdtResult 	datetime OUTPUT,
					@pdtDate 	datetime
					)
					AS
					BEGIN
						SET @pdtResult = dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, @pdtDate), @pdtDate)))
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 58'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (58, ''Last Day of Year'', 4, 0, ''Date/Time'', ''sp_ASRFn_LastDayOFYear'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 58'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (58, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 47 of 120 - Adding new Function Name Of Month'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_NameOfMonth]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_NameOfMonth]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_NameOfMonth
					(
					@psResult	varchar(8000) OUTPUT,
					@pdtDate 	datetime
					)
					AS
					BEGIN
						SET @psResult = datename(month, @pdtDate)
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 59'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (59, ''Name of Month'', 1, 0, ''Date/Time'', ''sp_ASRFn_NameOfMonth'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 59'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (59, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 48 of 120 - Adding new Function Name Of Day'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_NameOfDay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_NameOfDay]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_NameOfDay
					(
					@psResult	varchar(8000) OUTPUT,
					@pdtDate 	datetime		
					)
					AS
					BEGIN
						SET @psResult = datename(weekday, @pdtDate)
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 60'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (60, ''Name of Day'', 1, 0, ''Date/Time'', ''sp_ASRFn_NameOfDay'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 60'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (60, 1, 4, '<Date>')



/* ------------------------------------------------------------- */
PRINT 'Step 49 of 120 - Adding new Function Is Field Populated'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsPopulated_1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_IsPopulated_1]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsPopulated_1
					(
					@pfResult	bit OUTPUT,
					@psString	varchar(8000)
					)
					AS
					BEGIN
						SET @pfResult = 1
					
						IF len(@psString) = 0 
						BEGIN
							SET @pfResult = 0
						END

						IF @psString IS null
						BEGIN
							SET @pfResult = 0
						END
					END'
		EXEC sp_executesql @NVarCommand

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsPopulated_2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_IsPopulated_2]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsPopulated_2
					(
					@pfResult	bit OUTPUT,
					@pdblNumeric	float
					)
					AS
					BEGIN
						SET @pfResult = 1
					
						IF @pdblNumeric = 0 
						BEGIN
							SET @pfResult = 0
						END
					
						IF @pdblNumeric IS null
						BEGIN
							SET @pfResult = 0
						END
					END'
		EXEC sp_executesql @NVarCommand

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsPopulated_3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_IsPopulated_3]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsPopulated_3
					(
					@pfResult	bit OUTPUT,
					@pfLogic	bit
					)
					AS
					BEGIN
						SET @pfResult = 1

						IF @pfLogic = 0 
						BEGIN
							SET @pfResult = 0
						END
			
						IF @pfLogic IS null
						BEGIN
							SET @pfResult = 0
						END
					END'
		EXEC sp_executesql @NVarCommand

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsPopulated_4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRFn_IsPopulated_4]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsPopulated_4
					(
					@pfResult	bit OUTPUT,
					@pdtDate	datetime
					)
					AS
					BEGIN
						SET @pfResult = 1
					
						IF @pdtDate IS null
						BEGIN
							SET @pfResult = 0
						END
					END'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 61'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys)
       			VALUES                (61, ''Is Field Populated'', 3, 0, ''Comparison'', ''sp_ASRFn_IsPopulated'', 0, 1, NULL)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 61'
		EXEC sp_executesql @NVarCommand
		INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
			VALUES                         (61, 1, 0, '<Field>')



/* ------------------------------------------------------------- */
PRINT 'Step 50 of 120 - Updating System Permissions for Lookup Table Menu'

		DECLARE @iCount int
		DECLARE @fPermitted bit

		SELECT @iCount = COUNT(*) 
		FROM ASRSysPermissionCategories
		WHERE categoryID = 37

		IF @iCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
			VALUES (37, 'Menu', '', 10, 'MENU')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 37

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000000000000000000000FFFFF00000000000F000F00000000000FFFFF00000000000F000F00000000000FFFFF00000000000F000F00000000000FFFFF0000000000000000000000FFFF444444FFF000FFFF444444FFF00000000000000000000000000000000000000000000000000FFFF0000FFFF0000F80F0000F80F0000F80F0000F80F0000F80F0000F80F0000F80F0000F80F000000010000000100000001000000010000FFFF0000FFFF0000

			INSERT INTO ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
			VALUES (133, 'View lookup table menu', 10, 37, 'VIEWLOOKUPTABLES')

			DECLARE curGroups CURSOR LOCAL FAST_FORWARD FOR 
			SELECT name
			FROM sysusers
			WHERE uid = gid
				AND uid <> 0 
			OPEN curGroups
			FETCH NEXT FROM curGroups INTO @sGroup
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @fPermitted = permitted
				FROM ASRSysGroupPermissions
				WHERE groupName = @sGroup
					AND itemID = 133
			
				IF @fPermitted IS null
				BEGIN
					INSERT INTO ASRSysGroupPermissions (itemID, groupName, permitted)
					VALUES (133, @sGroup, 1)
				END
			
				FETCH NEXT FROM curGroups INTO @sGroup
			END
			CLOSE curGroups
			DEALLOCATE curGroups
		END

/* ------------------------------------------------------------- */
PRINT 'Step 51 of 120 - Adding ASRSysColours and Populate with Colour Codes'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysColours]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		drop table [dbo].[ASRSysColours]

		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysColours] (
      						[ColOrder] [int] NULL,
      						[ColValue] [int] NULL,
      						[ColDesc] [varchar] (50) NULL,
      						[WordColourIndex] [int] NULL,
						[CalendarLegendColour] [bit] NULL
					) ON [PRIMARY]'

		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(1, 16777215, ''White'', 8, 0)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(2, 16777164, ''Light Turquoise'', 8, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(3, 13434828, ''Light Green'', 8, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(4, 13434879, ''Light Yellow'', 7, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(5, 16764057, ''Pale Blue'', 8, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(6, 16751052, ''Lavender'', 5, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(7, 13408767, ''Rose'', 5, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(8, 10079487, ''Tan'', 7, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(9, 12632256, ''Grey 25%'', 16, 0)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(10, 16776960, ''Turquoise'', 3, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(11, 16711935, ''Pink'', 5, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(12, 65535, ''Yellow'', 7, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(13, 16763904, ''Sky Blue'', 3, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(14, 13421619, ''Aqua'', 3, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(15, 52479, ''Gold'', 7, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(16, 9868950, ''Grey 40%'', 15, 0)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(17, 16737843, ''Light Blue'', 2, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(18, 39423, ''Light Orange'', 6, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(19, 8421504, ''Grey 50%'', 15, 0)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(20, 13395456, ''Blue Grey'', 2, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(21, 52377, ''Lime'', 11, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(22, 26367, ''Orange'', 6, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(23, 6723891, ''Sea Green'', 11, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(24, 6697881, ''Plum'', 12, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(25, 16711680, ''Blue'', 2, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(26, 8421376, ''Teal'', 10, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(27, 8388736, ''Violet'', 12, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(28, 10040115, ''Indigo'', 12, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(29, 32896, ''Dark Yellow'', 14, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(30, 65280, ''Bright Green'', 4, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(31, 255, ''Red'', 6, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(32, 13209, ''Brown'', 1, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(33, 6697728, ''Dark Teal'', 1, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(34, 8388608, ''Dark Blue'', 9, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(35, 32768, ''Green'', 11, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(36, 128, ''Dark Red'', 13, 1)'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'INSERT INTO ASRSysColours(ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
		VALUES(37, 0, ''Black'', 1, 0)'
		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 52 of 120 - Altering SettingValue in ASRUserSettings table'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysUserSettings')
		and name = 'SettingValue'

		if @iRecCount > 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysUserSettings ALTER COLUMN
						[SettingValue] [varchar] (255) NULL '
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 53 of 120 - Adding new column to screen definitions'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysScreens')
		and name = 'SSIntranet'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysScreens]
			                       ADD [SSIntranet] [bit] NULL'
			EXEC sp_executesql @NVarCommand
		END



/* ------------------------------------------------------------- */
PRINT 'Step 54 of 120 - Adding new table for Self Service Intranet'

		if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysSSIntranetLinks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN

			SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysSSIntranetLinks] (
									[linkType] [int] NULL ,
									[linkOrder] [int] NULL ,
									[prompt] [varchar] (100) NULL ,
									[text] [varchar] (100) NULL ,
									[screenID] [int] NULL ,
									[pageTitle] [varchar] (100) NULL ,
									[url] [varchar] (100) NULL ,
									[viewID] [int] NULL 
									) ON [PRIMARY]'
			EXEC sp_executesql @NVarCommand

		END



/* ------------------------------------------------------------- */
PRINT 'Step 55 of 120 - Adding new keywords'

		/* Add new keyword restrictions */

		SELECT @NVarCommand = 'DELETE FROM ASRSysKeywords WHERE LOWER(ASRSysKeywords.Keyword) = ''dbo'''
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysKeywords (Provider, Keyword) VALUES(''Microsoft SQL Server'', ''dbo'')'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysKeywords (Provider, Keyword) VALUES(''Microsoft SQL Server'', ''DBO'')'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 56 of 120 - Amending Permissions Stored Procedures'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRAllTablePermissions]

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissionsForGroup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRAllTablePermissionsForGroup]

		EXEC('CREATE PROCEDURE [dbo].sp_ASRAllTablePermissions 
		AS
		BEGIN
			/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
			DECLARE @iUserGroupID	int

			/* Initialise local variables. */
			SELECT @iUserGroupID = sysusers.gid
			FROM sysusers
			WHERE sysusers.name = CURRENT_USER

			SELECT sysobjects.name, sysprotects.action
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND sysprotects.action <> 193
				AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
			UNION
			SELECT sysobjects.name, 193
			FROM syscolumns
			INNER JOIN sysprotects ON (syscolumns.id = sysprotects.id
				AND sysprotects.action = 193 
				AND sysprotects.uid = @iUserGroupID
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE syscolumns.name = ''timestamp''
				AND ((sysprotects.protectType = 205) 
				OR (sysprotects.protectType = 204))
			ORDER BY sysobjects.name
		END')

		EXEC('CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]
		(
			@psGroupName sysname
		)
		AS
		BEGIN
			/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
			DECLARE @iUserGroupID	int

			/* Initialise local variables. */
			SELECT @iUserGroupID = sysusers.gid
			FROM sysusers
			WHERE sysusers.name = @psGroupName

			SELECT sysobjects.name, sysprotects.action
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
			ORDER BY sysobjects.name
		END')


/* ------------------------------------------------------------- */
PRINT 'Step 57 of 120 - Updating Absence Breakdown Stored Procedures'

-- Removed - is updated later in update script

/* ------------------------------------------------------------- */
PRINT 'Step 58 of 120 - Update Bradford Factor Stored Procedures'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASR_Bradford_MergeAbsences]

		SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASR_Bradford_MergeAbsences
					(
					@pdReportStart	  	datetime,
					@pdReportEnd		datetime,
					@pcReportTableName	char(30)
					)
					AS
					BEGIN
						declare @sSql as char(8000)

						/* Variables to hold current absence record */
						declare @pdStartDate as datetime
						declare @pdEndDate as datetime
						declare @pcStartSession as char(2)
						declare @pfDuration as float
						declare @piID as integer
						declare @piPersonnelID as integer
						declare @pbContinuous as bit

						/* Variables to hold last absence record */
						declare @pdLastStartDate as datetime
						declare @pcLastStartSession as char(2)
						declare @pfLastDuration as float
						declare @piLastID as integer
						declare @piLastPersonnelID as integer

						/* Open the passed in table */
						set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' FOR UPDATE OF Start_Date, Start_Session, Duration,Included_Days''
						execute(@sSQL)
						open BradfordIndexCursor

						/* Loop through the records in the bradford report table */
						Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
						while @@FETCH_STATUS = 0
						begin

							if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
							begin
								Set @pdLastStartDate = @pdStartDate
								Set @pcLastStartSession = @pcStartSession
								Set @pfLastDuration = @pfDuration
								Set @piLastID = @piID
			
							end
							else
							begin

								Set @pfLastDuration = @pfLastDuration + @pfDuration

								/* update start date */
								set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(10), @pfLastDuration) + '', Included_Days = '' + Convert(Char(10), @pfLastDuration) + '' WHERE CURRENT OF BradFordIndexCursor''
								execute(@sSQL)

								/* Delete the previous record from our collection */
								set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = '' + Convert(varchar(10),@piLastId)
								execute(@sSQL)

								Set @piLastID = @piID

							end

							/* Get next absence record */
							Set @piLastPersonnelID = @piPersonnelID
							
							Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
						end

						close BradfordIndexCursor
						deallocate BradfordIndexCursor

					END'
		EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 59 of 120 - Add Drop Temp Objects Stored Procedure'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRDropTempObjects]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRDropTempObjects]

		EXEC('CREATE PROCEDURE spASRDropTempObjects
		AS
		BEGIN

			DECLARE	@sObjectName varchar(2000),
					@sUsername varchar(2000),
					@sXType varchar(50)
						
			DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
			SELECT [dbo].[sysobjects].[name], [dbo].[sysusers].[name], [dbo].[sysobjects].[xtype]
			FROM [dbo].[sysobjects] 
					INNER JOIN [dbo].[sysusers]
					ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
			WHERE LOWER([dbo].[sysusers].[name]) != ''dbo'' 
					AND (OBJECTPROPERTY(id, N''IsUserTable'') = 1
						OR OBJECTPROPERTY(id, N''IsProcedure'') = 1)

			OPEN tempObjects
			FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
			WHILE (@@fetch_status <> -1)
			BEGIN
				
				IF UPPER(@sXType) = ''U''
					-- user table
					BEGIN
						EXEC (''DROP TABLE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''P''
					-- procedure
					BEGIN
						EXEC (''DROP PROCEDURE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END
				
				FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
				
			END
			CLOSE tempObjects
			DEALLOCATE tempObjects
			
			EXEC (''DELETE FROM [dbo].[ASRSysSQLObjects]'')

		END')


/* ------------------------------------------------------------- */
PRINT 'Step 60 of 120 - Dropping Obsolete Objects'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysBatchLock]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		drop table [dbo].[ASRSysBatchLock]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRInsertNewRecord]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRInsertNewRecord]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetDataTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRGetDataTransferDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRGetTransferDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRJoinTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRJoinTransferDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRPrimaryTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRPrimaryTransferDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRSecondaryTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRSecondaryTransferDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRSecondaryJoinDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRSecondaryJoinDetails]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAddToPicklist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRAddToPicklist]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRPicklistWhere]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRPicklistWhere]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRUniqueCheck]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRUniqueCheck]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRUserValidation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRUserValidation]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRRecordAmended]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRRecordAmended]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCopyDataExport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRCopyDataExport]

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetBatchLock]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASRGetBatchLock]

/* ------------------------------------------------------------- */
PRINT 'Step 61 of 120 - Adding HeadingText to Labels & Envelopes'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysMailMergeColumns')
	and name = 'Headingtext'
	
	if @iRecCount = 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysMailMergeColumns] ADD [Headingtext] [varchar] (50)'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 62 of 120 - Adding Thousand Separator to Cross Tabs'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysCrossTab')
	and name = 'ThousandSeparators'
	
	if @iRecCount = 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysCrossTab] ADD [ThousandSeparators] [bit]'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysCrossTab] SET ThousandSeparators = 0'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 63 of 120 - Adding Lookup Filter Column'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysColumns')
and name = 'LookupFilter'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] ADD
				[LookupFilterID] [int] NULL,
				[LookupFilter] [bit] NULL'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 64 of 120 - Removing redundant Self-service Intranet columns'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
	and name = 'viewID'
	
	if @iRecCount > 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysSSIntranetLinks] DROP COLUMN [viewID]'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 65 of 120 - Updating Absence Breakdown Stored Procedures'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_AbsenceBreakdown_Run]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[sp_ASR_AbsenceBreakdown_Run]

		execute('CREATE PROCEDURE sp_ASR_AbsenceBreakdown_Run
		(
		@pdReportStart      datetime,
		@pdReportEnd    datetime,
		@pcReportTableName  char(30)
		) 
		AS 
		begin
		declare @pdStartDate as datetime
		declare @pdEndDate as datetime
		declare @pcStartSession as char(2)
		declare @pcEndSession as char(2)
		declare @pcType as char(50)
		declare @pcRecordDescription as char(100)

		declare @pfDuration as float
		declare @pdblSun as float
		declare @pdblMon as float
		declare @pdblTue as float
		declare @pdblWed as float
		declare @pdblThu as float
		declare @pdblFri as float
		declare @pdblSat as float

		declare @sSQL as char(8000)
		declare @piParentID as integer
		declare @piID as integer
		declare @pbProcessed as bit

		declare @pdTempStartDate as datetime
		declare @pdTempEndDate as datetime
		declare @pcTempStartSession as char(2)
		declare @pcTempEndSession as char(2)

		declare @pfCount as float
		declare @psVer as char(80)

		/* Alter the structure of the temporary table so it can hold the text for the days */
		Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ALTER COLUMN Hor NVARCHAR(10)''
		execute(@sSQL)
		Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ADD Processed BIT''
		execute(@sSQL)
		Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ADD DisplayOrder INT''
		execute(@sSQL)
		Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ALTER COLUMN Value decimal(10,5)''
		execute(@sSQL)

		/* Load the values from the temporary cursor */
		Set @sSQL = ''DECLARE AbsenceBreakdownCursor CURSOR STATIC FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM '' + @pcReportTableName
		execute(@sSQL)
		open AbsenceBreakdownCursor

		/* Loop through the records in the absence breakdown report table */
		Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed
		while @@FETCH_STATUS = 0
			begin

			Set @pdblSun = 0
			Set @pdblMon = 0
			Set @pdblTue = 0
			Set @pdblWed = 0
			Set @pdblThu = 0
			Set @pdblFri = 0
			Set @pdblSat = 0

			/* If blank leaving date set it to todays date */
			if @pdEndDate = Null set @pdEndDate = getdate()

			/* The absence should only calculate for absence within the reporting period */
			set @pdTempStartDate = @pdStartDate
			set @pcTempStartSession = @pcStartSession
			set @pdTempEndDate = @pdEndDate
			set @pcTempEndSession = @pcEndSession

			if @pdStartDate <  @pdReportStart
				begin
				set @pdTempStartDate = @pdReportStart
				set @pcTempStartSession = ''AM''
				end
			if @pdEndDate >  @pdReportEnd
				begin
				set @pdTempEndDate = @pdReportEnd
				set @pcTempEndSession = ''PM''
				end

			/* Calculate the days this absence takes up */
			execute sp_ASR_AbsenceBreakdown_Calculate @pfDuration OUTPUT, @pdblMon OUTPUT, @pdblTue OUTPUT, @pdblWed OUTPUT, @pdblThu OUTPUT, @pdblFri OUTPUT, @pdblSat OUTPUT, @pdblSun OUTPUT, @pdTempStartDate, @pcTempStartSession, @pdTempEndDate, @pcTempEndSession, @piParentID

			/* Strip out dodgy characters */
			set @pcRecordDescription = replace(@pcRecordDescription,'''''''','''')
			set @pcType = replace(@pcType,'''''''','''')

			/* Add Mondays records */
			if @pdblMon > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 0) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblMon) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',1,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',1)''
				execute(@sSQL)
				end

			/* Add Tuesday records */
			if @pdblTue > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 1) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblTue) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',2,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',2)''
				execute(@sSQL)
				end

			/* Add Wednesdays records */
			if @pdblWed > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 2) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblWed) +  '','''''' + convert(varchar(20),@pdStartDate) +  '''''',3,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',3)''
				execute(@sSQL)
				end

			/* Add new records depending on how many Thursdays were found */
			if @pdblThu > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 3) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblThu) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',4,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',4)''
				execute(@sSQL)
				end

			/* Add new records depending on how many Fridays were found */
			if @pdblFri > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 4) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblFri) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',5,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',5)''
				execute(@sSQL)
				end

			/* Add new records depending on how many Saturdays were found */
			if @pdblSat > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 5) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSat) + '',''''''+ convert(varchar(20),@pdStartDate) + '''''',6,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',6)''
				execute(@sSQL)
				end

			/* Add new records depending on how many Sundays were found */
			if @pdblSun > 0
				begin
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 6) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSun) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',7,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',0)''
				execute(@sSQL)
				end

			/* Calculate total duraton of absence */
			set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun

			if @pfDuration > 0
				begin
				/* Write records for average, totals and count */
				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Total'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',8)''
				execute(@sSQL)

				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Count'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),1) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',10,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',10)''
				execute(@sSQL)

				set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Average'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'''''' + convert(varchar(20),@pdEndDate) +'''''',9)''
				execute(@sSQL)
				end

			/* Process next record */
			Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed

			end

		/* Delete this record from our collection as it''s now been processed */
		set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Processed IS NULL''
		execute(@sSQL)

		Set @sSQL = ''DECLARE CalculateAverage CURSOR STATIC FOR SELECT Ver,(SUM(Value) / COUNT(Value)) / COUNT(Value) FROM '' + @pcReportTableName + '' WHERE hor = ''''Average'''' GROUP BY Ver''
		execute(@sSQL)
		open CalculateAverage

		Fetch Next From CalculateAverage Into @psVer, @pfCount
		while @@FETCH_STATUS = 0
			begin
      			Set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Value = '' + Convert(varchar(10),@pfCount) + '' WHERE Ver =  '''''' + @psVer + '''''' AND Hor = ''''Average''''''
			execute(@sSQL)
				Fetch Next From CalculateAverage Into @psVer, @pfCount
			end

		/* Tidy up */
		close AbsenceBreakdownCursor
		close CalculateAverage
		deallocate AbsenceBreakdownCursor
		deallocate CalculateAverage

		END')


/* ------------------------------------------------------------- */
PRINT 'Step 66 of 120 - Adding Maternity and Parential Leave Functions'

		DELETE FROM ASRSysFunctions WHERE functionID IN (62, 63, 64)

		SELECT @NVarCommand = 'INSERT ASRSysFunctions (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ShortcutKeys)
					VALUES (62, ''Parental Leave Entitlement'', 2, 1, ''Absence'', ''spASRSysFnParentalLeaveEntitlement'', 0, 0, 0, null)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT ASRSysFunctions (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ShortcutKeys)
					VALUES (63, ''Parental Leave Taken'', 2, 1, ''Absence'', ''spASRSysFnParentalLeaveTaken'', 0, 0, 0, null)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT ASRSysFunctions (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ShortcutKeys)
					VALUES (64, ''Maternity Expected Return Date'', 4, 0, ''Absence'', ''spASRSysFnMaternityExpectedReturn'', 0, 0, 0, null)'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 67 of 120 - Adding Maternity and Parential Leave Functions'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[spASRMaternityExpectedReturn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRMaternityExpectedReturn]

		execute('CREATE PROCEDURE dbo.spASRMaternityExpectedReturn (
			@pdblResult datetime OUTPUT,
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(8000)
			)
			AS
			BEGIN

				IF LOWER(@Ordinary) = ''ordinary''
					IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,26,@LeaveStart)
					ELSE
						IF DateDiff(d,''04/30/2000'', @EWCDate) >= 0
							SET @pdblResult = Dateadd(ww,18,@LeaveStart)
						ELSE
							SET @pdblResult = Dateadd(ww,14,@LeaveStart)
				ELSE
					IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,52,@LeaveStart)
					ELSE
						--29 weeks from baby birth date (but return on the monday before!)
						SET @pdblResult = DateAdd(d,203 - datepart(dw,DateAdd(d,-2,@BabyBirthDate)),@BabyBirthDate)

			END')


		if exists (select * from sysobjects where id = object_id(N'[dbo].[spASRParentalLeaveEntitlement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRParentalLeaveEntitlement]

		execute('CREATE PROCEDURE dbo.spASRParentalLeaveEntitlement (
			@pdblResult    float OUTPUT,
			@DateOfBirth datetime,
			@AdoptedDate datetime,
			@Disabled bit
			)
			AS
			BEGIN

			DECLARE @Today datetime
			DECLARE @ChildAge int
			DECLARE @Adopted bit
			DECLARE @YearsOfResponsibility int
			DECLARE @StartDate datetime

			--Check if we should used the Date of Birth or the Date of Adoption column...
			SET @Adopted = 0
			SET @StartDate = @DateOfBirth
			IF NOT @AdoptedDate IS NULL
			BEGIN
				SET @Adopted = 1
				SET @StartDate = @AdoptedDate
			END

			--Set variables based on this date...
			--(years of responsibility = years since born or adopted)
			SET @Today = getdate()
			EXEC sp_ASRFn_WholeYearsBetweenTwoDates @ChildAge OUTPUT, @DateOfBirth, @Today
			EXEC sp_ASRFn_WholeYearsBetweenTwoDates @YearsOfResponsibility OUTPUT, @StartDate, @Today


			SELECT @pdblResult = CASE
				WHEN @Disabled = 0 And @Adopted = 0 And @ChildAge < 5
					THEN
					65

				WHEN @Disabled = 0 And @Adopted = 1 And @ChildAge < 18
					And @YearsOfResponsibility < 5 THEN
					65

				WHEN @Disabled = 1 And @Adopted = 0 And @ChildAge < 18 
					And DateDiff(d,''12/15/1994'',@DateOfBirth) >= 0 THEN
					90

				WHEN @Disabled = 1 And @Adopted = 1 And @ChildAge < 18 
				And DateDiff(d,''12/15/1994'',@AdoptedDate) >= 0 THEN
					90

				ELSE
					0
				END

			END')


/* ------------------------------------------------------------- */
PRINT 'Step 68 of 120 - Adding New Function Column'

			/* Adding Columns to Label/Envelope template table */
			SELECT @iRecCount = count(id) FROM syscolumns
			where id = (select id from sysobjects where name = 'ASRSysFunctions')
			and name = 'ExcludeExprTypes'

			if @iRecCount = 0
			BEGIN
				SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysFunctions] ADD
							[ExcludeExprTypes] [varchar](50) NULL'
				EXEC sp_executesql @NVarCommand
				SELECT @NVarCommand = 'UPDATE ASRSysFunctions SET [ExcludeExprTypes] = '''''
				EXEC sp_executesql @NVarCommand
				SELECT @NVarCommand = 'UPDATE ASRSysFunctions SET [ExcludeExprTypes] = ''4 15 16 17'' WHERE functionid IN (30, 47)'
				EXEC sp_executesql @NVarCommand
			END



/* ------------------------------------------------------------- */
PRINT 'Step 69 of 120 - Updating icon and description for Label Definitions'

UPDATE asrsysPermissionCategories SET
	Description = 'Envelope & Label Templates'
	WHERE categoryID = 30

SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 30
WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000400100000000000000000000000100000000000000000000800080008000000080800000008000000080800000008000C0C0C000C0DCC000F0CAA60080808000FF00FF00FF000000FFFF000000FF000000FFFF000000FF00FFFFFF00F0FBFF00A4A0A000D4F0FF00B1E2FF008ED4FF006BC6FF0048B8FF0025AAFF0000AAFF000092DC00007AB90000629600004A730000325000D4E3FF00B1C7FF008EABFF006B8FFF004873FF002557FF000055FF000049DC00003DB900003196000025730000195000D4D4FF00B1B1FF008E8EFF006B6BFF004848FF002525FF000000FF000000DC000000B900000096000000730000005000E3D4FF00C7B1FF00AB8EFF008F6BFF007348FF005725FF005500FF004900DC003D00B900310096002500730019005000F0D4FF00E2B1FF00D48EFF00C66BFF00B848FF00AA25FF00AA00FF009200DC007A00B900620096004A00730032005000FFD4FF00FFB1FF00FF8EFF00FF6BFF00FF48FF00FF25FF00FF00FF00DC00DC00B900B900960096007300730050005000FFD4F000FFB1E200FF8ED400FF6BC600FF48B800FF25AA00FF00AA00DC009200B9007A009600620073004A0050003200FFD4E300FFB1C700FF8EAB00FF6B8F00FF487300FF255700FF005500DC004900B9003D00960031007300250050001900FFD4D400FFB1B100FF8E8E00FF6B6B00FF484800FF252500FF000000DC000000B9000000960000007300000050000000FFE3D400FFC7B100FFAB8E00FF8F6B00FF734800FF572500FF550000DC490000B93D0000963100007325000050190000FFF0D400FFE2B100FFD48E00FFC66B00FFB84800FFAA2500FFAA0000DC920000B97A000096620000734A000050320000FFFFD400FFFFB100FFFF8E00FFFF6B00FFFF4800FFFF2500FFFF0000DCDC0000B9B90000969600007373000050500000F0FFD400E2FFB100D4FF8E00C6FF6B00B8FF4800AAFF2500AAFF000092DC00007AB90000629600004A73000032500000E3FFD400C7FFB100ABFF8E008FFF6B0073FF480057FF250055FF000049DC00003DB90000319600002573000019500000D4FFD400B1FFB1008EFF8E006BFF6B0048FF480025FF250000FF000000DC000000B90000009600000073000000500000D4FFE300B1FFC7008EFFAB006BFF8F0048FF730025FF570000FF550000DC490000B93D00009631000073250000501900D4FFF000B1FFE2008EFFD4006BFFC60048FFB80025FFAA0000FFAA0000DC920000B97A000096620000734A0000503200D4FFFF00B1FFFF008EFFFF006BFFFF0048FFFF0025FFFF0000FFFF0000DCDC0000B9B900009696000073730000505000F2F2F200E6E6E600DADADA00CECECE00C2C2C200B6B6B600AAAAAA009E9E9E0092929200868686007A7A7A006E6E6E0062626200565656004A4A4A003E3E3E0032323200262626001A1A1A000E0E0E0000000000000000000A0000000000000000000000000000000A0A00000000000000000000000000000A0A0A000000000000000000000000000A0F070A0000000000000000000000000A0F070A070000000000000000000000070A0F070A0000000011111111111111000A0F070A070000001111111100000000070A0F070A0000001111111111111111110A0F070A070000111111110000000000070A0F070A0000111111111111111111110A0F0A0A070011111111111111111110070A0A0A0A0011000000111111111110100A0A0A00001111111111111111111111110000000000000000000000000000000000000000000000000000000000000000000000C07FFFFFFF3FFFFFFF1FFFFFFF0FFFFFFF07FFFF0001FFFF0001FFFF0001FFFF0001FFFF0001FFFF0000FFFF0000FFFF0001FFFF0003FFFF0003FFFFFFFFFFFF00


/* ------------------------------------------------------------- */
PRINT 'Step 70 of 120 - Adding new columns to Envelopes & Labels'

/* Adding Columns to Label/Envelope template table */
SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysLabelTypes')
and name = 'IsEnvelope'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysLabelTypes] ADD
				[IsEnvelope] [bit] NULL,
				[MeasurementMethod] [integer] NULL'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysLabelTypes SET
				[IsEnvelope] = 0,
				[MeasurementMethod] = 0'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 71 of 120 - Adding page sizes for envelopes & labels'

if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysPageSizes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ASRSysPageSizes]

SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysPageSizes] (
			[PageSizeID] [int] NOT NULL ,
			[Name] [char] (20) NOT NULL ,
			[Width] [float] NOT NULL ,
			[Height] [float] NOT NULL ,
			[DisplayOrder] [int] NOT NULL ,
			[WordTemplateID] [int] NOT NULL ,
			[IsEnvelope] [bit] NOT NULL
		) ON [PRIMARY]'
EXEC sp_executesql @NVarCommand

DELETE FROM asrSysPageSizes

-- Label Types
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (1,''Custom'',0,0,100,0,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (2,''A4'',21,29.7,1,7,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (3,''A5'',14.8,21,10,9,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (4,''Letter'',21.59,27.94,10,2,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (5,''Mini'',10.48,12.7,10,7,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (6,''Vertical Half Sheet'',10.79,25.4,10,9,0)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (7,''B5'',18.2,25.7,10,2,6)'
EXEC sp_executesql @NVarCommand

-- Envelope Types
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (8,''Custom'',0,0,100,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (9,''B4'',35.3,25,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (10,''B5'',25,17.6,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (11,''B6'',17.6,12.5,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (12,''C3'',45.8,32.4,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (13,''C4'',32.4,22.9,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (14,''C5'',22.9,16.2,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (15,''C6'',16.2,11.4,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (16,''C65'',22.9,11.4,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (17,''E4'',22.0,31.0,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (18,''E5'',15.5,22.0,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (19,''E6'',11.0,15.5,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (20,''E65 / DL'',11.0,22.0,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (21,''M5'',15.5,22.3,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (22,''M65'',11.2,22.3,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (23,''US Legal'',35.56,21.59,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (24,''US Letter'',27.94,21.59,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (24,''US Letter'',27.94,21.59,10,0,1)'
EXEC sp_executesql @NVarCommand
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (25,''A4'',29.7,21,1,0,1)'
EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 72 of 120 - Adding new columns to Envelopes & Labels'

/* Adding Columns to Label/Envelope template table */
SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysLabelTypes')
and name = 'FromTop'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysLabelTypes] ADD
				[FromTop] [float] NULL,
				[FromLeft] [float] NULL,
				[FromTopAuto] [bit] NULL,
				[FromLeftAuto] [bit] NULL'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysLabelTypes SET
				[FromTop] = 0,
				[FromLeft] = 0,
				[FromTopAuto] = 1,
				[FromLeftAuto] = 1'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 73 of 120 - Creating Access Table for Global Functions'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysGlobalAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysGlobalAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysGlobalAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysGlobalFunctions.access
			END,
			functionID
		FROM ASRSysGlobalFunctions,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysGlobalFunctions] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 74 of 120 - Creating Access Table for Export'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysExportAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysExportAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysExportAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysExportName.access
			END,
			ID
		FROM ASRSysExportName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysExportName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 75 of 120 - Creating Access Table for Import'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysImportAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysImportAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysImportAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysImportName.access
			END,
			ID
		FROM ASRSysImportName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysImportName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 76 of 120 - Creating Access Table for Data Transfer'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysDataTransferAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysDataTransferAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysDataTransferAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysDataTransferName.access
			END,
			DataTransferID
		FROM ASRSysDataTransferName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysDataTransferName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 77 of 120 - Creating Access Table for Match Reports, Career Progrression & Succession Planning'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysMatchReportAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysMatchReportAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysMatchReportAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysMatchReportName.access
			END,
			MatchReportID
		FROM ASRSysMatchReportName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysMatchReportName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 78 of 120 - Adding Event Log v2 modifications.'

/* Adding Columns to Label/Envelope template table */
SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysEventLog')
and name = 'EndTime'

if @iRecCount = 0
BEGIN

	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysEventLog] 
				ADD [EndTime] datetime NULL,
				    [Duration] numeric(18,0) NULL,
			   	    [BatchJobID] int NULL'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 79 of 120 - Creating version information table'

if exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysVersionInformation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ASRSysVersionInformation]

/*
SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysVersionInformation] (
	[ChangeID] [int] NOT NULL ,
	[Description] [varchar] (1000) NOT NULL ,
	[Area] [varchar] (100) NOT NULL ,
	[SQL_2000_Only] [bit] NOT NULL ,
	[Version] [varchar] (50) NOT NULL ,
	[HRPro_Module_Code] [int] NOT NULL 
) ON [PRIMARY]'
EXEC sp_executesql @NVarCommand
*/

/* ------------------------------------------------------------- */
PRINT 'Step 80 of 120 - Adding Delete Trigger for Event Log'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysEventLog]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysEventLog]

		SELECT @NVarCommand = 'CREATE TRIGGER [DEL_ASRSysEventLog] ON [dbo].[ASRSysEventLog] 
														FOR DELETE 
														AS
														BEGIN
															DELETE FROM ASRSysEventLogDetails WHERE ASRSysEventLogDetails.EventLogID IN (SELECT ID FROM deleted)
														END'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 81 of 120 - Updating Delete Trigger for Custom Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysCustomReportsName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysCustomReportsName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCustomReportsName ON ASRSysCustomReportsName
														FOR DELETE 
														AS
														BEGIN
															DELETE FROM ASRSysCustomReportsDetails WHERE CustomReportID IN (SELECT ID FROM Deleted)
															DELETE FROM ASRSysCustomReportsChildDetails WHERE CustomReportID IN (SELECT ID FROM Deleted)
														END'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 82 of 120 - Creating Access Table for Cross Tabs'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysCrossTabAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysCrossTabAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysCrossTabAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
			 		ASRSysCrossTab.access
			END,
			crossTabID
	 	FROM ASRSysCrossTab,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysCrossTab] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 83 of 120 - Creating Access Table for Custom Reports'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysCustomReportAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysCustomReportAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysCustomReportAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
			 		ASRSysCustomReportsName.access
			END,
			ID
	 	FROM ASRSysCustomReportsName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysCustomReportsName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 84 of 120 - Creating Access Table for Mail Merge and Envelopes & Labels'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysMailMergeAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysMailMergeAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysMailMergeAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
					ASRSysMailMergeName.access
			END,
			MailMergeID
		FROM ASRSysMailMergeName,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysMailMergeName] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 85 of 120 - Creating Access Table for Calendar Reports'

	SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects 
	WHERE name = 'ASRSysCalendarReportAccess'

	IF @iRecCount = 0 
	BEGIN
		CREATE TABLE [dbo].[ASRSysCalendarReportAccess] (
			[GroupName] [varchar] (256) NOT NULL ,
			[Access] [varchar] (2) NOT NULL ,
			[ID] [int] NOT NULL 
		) ON [PRIMARY]


		SELECT @NVarCommand = 'INSERT INTO ASRSysCalendarReportAccess 
			(groupName, access, id)
			(SELECT sysusers.name,
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
				ELSE
			 		ASRSysCalendarReports.access
			END,
			ID
	 	FROM ASRSysCalendarReports,
			sysusers
		WHERE sysusers.uid = sysusers.gid
			and sysusers.uid <> 0)'

		exec sp_sqlexec @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysCalendarReports] 
					DROP COLUMN [access] '
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 86 of 120 - Add Email Attach As columns where required'


	declare @TableName varchar(255)

	SET @TableName = 'ASRSysCrossTab'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END

	SET @TableName = 'ASRSysCalendarReports'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END

	SET @TableName = 'ASRSysCustomReportsName'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END

	SET @TableName = 'ASRSysExportName'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END

	SET @TableName = 'ASRSysMatchReportName'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END

	SET @TableName = 'ASRSysRecordProfilename'
	if not exists(SELECT id FROM syscolumns where id = object_ID(@TableName) and name = 'OutputEmailAttachAs')
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].['+@TableName+'] ADD [OutputEmailAttachAs] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE [dbo].['+@TableName+'] SET [OutputEmailAttachAs] ='''''
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 87 of 120 - Amending batch job locking stored procedure'

	if exists (select * from sysobjects where id = object_id(N'[dbo].[spASRLockWriteBatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRLockWriteBatch]

	execute('CREATE Procedure spASRLockWriteBatch (@BatchJobID int, @Clearlock bit, @LockedByOther int OUTPUT)
		AS
		BEGIN

			DECLARE @OrigTranCount int
			DECLARE @Realspid int

			SET @OrigTranCount = @@trancount
			IF @OrigTranCount = 0 BEGIN TRANSACTION

			SELECT @LockedByOther = COUNT(ID) FROM ASRSysBatchJobName
			JOIN master..sysprocesses syspro ON spid = LockSpid
			WHERE LockLoginTime = syspro.login_time AND LockSpid <> @@spid
			AND ID = @BatchJobID

			IF @LockedByOther = 0
			BEGIN

				--Need to get spid of parent process
				SELECT @Realspid = a.spid
				FROM master..sysprocesses a
				FULL OUTER JOIN master..sysprocesses b
					ON a.hostname = b.hostname
					AND a.hostprocess = b.hostprocess
					AND a.spid <> b.spid
				WHERE b.spid = @@Spid

				--If there is no parent spid then use current spid
				--IF @Realspid is null SET @Realspid = @@spid

				IF @Clearlock = 0
					UPDATE ASRSysBatchJobName SET
					LockSpid = @Realspid,
					LockLoginTime = (
						SELECT login_time
						FROM master..sysprocesses
						WHERE spid = @Realspid)
					WHERE ID = @BatchJobID
				ELSE
					UPDATE ASRSysBatchJobName SET
					LockSpid = 0,
					LockLoginTime = null
					WHERE ID = @BatchJobID

			END

			IF @OrigTranCount = 0 COMMIT TRANSACTION

		END')


/* ------------------------------------------------------------- */
PRINT 'Step 88 of 120 - Modifying Lookup Filter Column'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysColumns')
and name = 'LookupFilter'

if @iRecCount > 0
BEGIN
	SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [LookupFilter]'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [LookupFilterID]'
	EXEC sp_executesql @NVarCommand
END

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysColumns')
and name = 'LookupFilterColumnID'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] ADD
				[LookupFilterColumnID] [int] NULL,
				[LookupFilterValueID] [int] NULL'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 89 of 120 - Modifying get control details stored procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetControlDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetControlDetails]

execute('CREATE PROCEDURE [dbo].[sp_ASRGetControlDetails] (
					@piScreenID int)
		AS
		BEGIN
					SELECT ASRSysControls.*, 
								ASRSysColumns.columnName, 
								ASRSysColumns.columnType, 
								ASRSysColumns.datatype,
								ASRSysColumns.defaultValue,
								ASRSysColumns.size, 
								ASRSysColumns.decimals, 
								ASRSysColumns.lookupTableID, 
								ASRSysColumns.lookupColumnID, 
								ASRSysColumns.lookupFilterColumnID, 
								ASRSysColumns.lookupFilterValueID, 
								ASRSysColumns.spinnerMinimum, 
								ASRSysColumns.spinnerMaximum, 
								ASRSysColumns.spinnerIncrement, 
								ASRSysColumns.mandatory, 
								ASRSysColumns.uniquecheck,
								ASRSysColumns.convertcase, 
								ASRSysColumns.mask, 
								ASRSysColumns.blankIfZero, 
								ASRSysColumns.multiline, 
								ASRSysColumns.alignment AS colAlignment, 
								ASRSysColumns.calcExprID, 
								ASRSysColumns.gotFocusExprID, 
								ASRSysColumns.lostFocusExprID, 
								ASRSysColumns.dfltValueExprID, 
								ASRSysColumns.calcTrigger, 
								ASRSysColumns.readOnly, 
								ASRSysColumns.statusBarMessage, 
								ASRSysColumns.errorMessage, 
								ASRSysColumns.linkTableID,
								ASRSysColumns.linkViewID,
								ASRSysColumns.linkOrderID,
								ASRSysColumns.Afdenabled,
								ASRSysColumns.OleOnServer,
								ASRSysTables.TableName,
								ASRSysColumns.Trimming,
								ASRSysColumns.Use1000Separator
					FROM ASRSysControls 
					LEFT OUTER JOIN ASRSysTables 
								ON ASRSysControls.tableID = ASRSysTables.tableID 
					LEFT OUTER JOIN ASRSysColumns 
								ON ASRSysColumns.tableID = ASRSysControls.tableID 
								AND ASRSysColumns.columnID = ASRSysControls.columnID
					WHERE ASRSysControls.ScreenID = @piScreenID
					ORDER BY ASRSysControls.PageNo, 
								ASRSysControls.ControlLevel DESC, 
								ASRSysControls.tabIndex
		END')


/* ------------------------------------------------------------- */
PRINT 'Step 90 of 120 - Updating Page Sizes'

SET @NVarCommand = 'UPDATE asrSysPageSizes SET IsEnvelope = 0 WHERE PageSizeID = 7'
EXEC sp_executesql @NVarCommand

DELETE FROM asrSysPageSizes WHERE PageSizeId = 24
SET @NVarCommand = 'INSERT asrSysPageSizes (PageSizeID,Name,Width,Height,DisplayOrder,wordtemplateID,IsEnvelope) VALUES (24,''US Letter'',27.94,21.59,10,0,1)'
EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 91 of 120 - Creating menu view permissions'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE name = 'ASRSysViewMenuPermissions'

IF @iRecCount = 0 
BEGIN
	CREATE TABLE [dbo].[ASRSysViewMenuPermissions] (
		[TableID] [int] NOT NULL ,
		[TableName] [varchar] (128) NOT NULL ,
		[groupName] [varchar] (255) NOT NULL ,
		[HideFromMenu] [bit] NOT NULL
	) ON [PRIMARY]
END


/* ------------------------------------------------------------- */
PRINT 'Step 92 of 120 - Load summary fields'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetSummaryFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetSummaryFields]

execute('CREATE PROCEDURE sp_ASRGetSummaryFields (
	@piHistoryTableID	int,
	@piParentTableID 	int)
AS
BEGIN
	SELECT DISTINCT ASRSysSummaryFields.sequence, 
	    	ASRSysSummaryFields.startOfGroup, 
		ASRSysColumns.columnName, 
		ASRSysColumns.columnID, 
		ASRSysColumns.tableID, 
		ASRSysColumns.dataType, 
		ASRSysColumns.size, 
		ASRSysColumns.decimals, 
		ASRSysColumns.controlType, 
		ASRSysColumns.columnType, 
		ASRSysColumns.multiline,
		ASRSysColumns.alignment,
		ASRSysColumns.BlankIfZero,
		ASRSysColumns.Use1000Separator,		
	    ASRSysSummaryFields.StartOfColumn
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence
END')


/* ------------------------------------------------------------- */
PRINT 'Step 93 of 120 - Remove obsolete column definition information'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'DigitGrouping'

if @iRecCount = 1
BEGIN	
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [DigitGrouping] '
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [DigitSeparator] '
	EXEC sp_executesql @NVarCommand

END



/* ------------------------------------------------------------- */
PRINT 'Step 94 of 120 - Adding additional new columns to Envelopes & Labels'

/* Adding Columns to Label/Envelope template table */
SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysLabelTypes')
and name = 'HeadingFontName'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysLabelTypes] ADD
				[HeadingFontName] [varchar] (255) NULL,
				[HeadingFontSize] [integer] NULL,
				[HeadingFontColour] [float] NULL,
				[HeadingFontBold] [bit] NULL,
				[HeadingFontItalic] [bit] NULL,
				[HeadingFontUnderline] [bit] NULL,
				[StandardFontName] [varchar] (255) NULL,
				[StandardFontSize] [integer] NULL,
				[StandardFontColour] [float] NULL,
				[StandardFontBold] [bit] NULL,
				[StandardFontItalic] [bit] NULL,
				[StandardFontUnderline] [bit] NULL'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysLabelTypes SET
				[HeadingFontName] = ''Tahoma'',
				[HeadingFontSize] = 12,
				[HeadingFontColour] = 0,
				[HeadingFontBold] = 1,
				[HeadingFontItalic] = 0,
				[HeadingFontUnderline] = 0,
				[StandardFontName] = ''Tahoma'',
				[StandardFontSize] = 10,
				[StandardFontColour] = 0,
				[StandardFontBold] = 0,
				[StandardFontItalic] = 0,
				[StandardFontUnderline] = 0'
	EXEC sp_executesql @NVarCommand
END


/* ------------------------------------------------------------- */
PRINT 'Step 95 of 120 - Adding new columns to Calendar Reports'

/* Adding Columns to Calendar reports table */
SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysCalendarReports')
and name = 'DescriptionSeparator'

if @iRecCount = 0
BEGIN
	SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysCalendarReports] ADD
				[DescriptionSeparator] [varchar] (6)'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'UPDATE ASRSysCalendarReports SET [DescriptionSeparator] = '', '''
	EXEC sp_executesql @NVarCommand
END



/* ------------------------------------------------------------- */
Print 'Step 96 of 120 - Adding overnight System Setting'

	-- Make sure that this setting is populated for the first time they use Data Manager
    DELETE from asrsyssystemsettings
    WHERE [Section] = 'overnight' and [SettingKey] = 'last completed'
    INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
    VALUES('overnight', 'last completed', convert(varchar,getdate(),103)+' '+convert(varchar,getdate(),108))


/* ------------------------------------------------------------- */
PRINT 'Step 97 of 120 - Adding new columns to Email Links'

	/* Add new column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysEmailLinks')
	and name = 'EmailInsert'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysEmailLinks ADD 
					[EmailInsert] [bit] NULL,
					[EmailUpdate] [bit] NULL,
					[EmailDelete] [bit] NULL'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysEmailLinks SET [EmailInsert] = 1
					, [EmailUpdate] = 1, [EmailDelete] = 1'
		EXEC sp_executesql @NVarCommand

	END

	/* Add recalculate record description */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysEmailQueue')
	and name = 'RecalculateRecordDesc'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysEmailQueue ADD 
					[RecalculateRecordDesc] [bit] NULL,
					[TableID] int NULL'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysEmailQueue SET [RecalculateRecordDesc] = 1'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 98 of 120 - Adding new Email Links to tables'

	/* Add new column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysTables')
	and name = 'AuditInsert'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysTables ADD 
					[AuditInsert] [bit] NULL,
					[AuditDelete] [bit] NULL,
					[EmailInsert] [int] NULL,
					[EmailDelete] [int] NULL'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysTables SET [EmailInsert] = 0
					, [EmailDelete] = 0, [AuditInsert] = 0, [AuditDelete] = 0'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 99 of 120 - Changing colours in Output options'

SELECT @NVarCommand = 'UPDATE ASRSysColours SET WordColourIndex = 3 WHERE ColOrder = 2'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'UPDATE ASRSysColours SET WordColourIndex = 3 WHERE ColOrder = 5'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'UPDATE ASRSysColours SET WordColourIndex = 4 WHERE ColOrder = 3'
EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 100 of 120 - Adding Delete Trigger for Batch Jobs'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysBatchJobName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysBatchJobName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysBatchJobName ON dbo.ASRSysBatchJobName 
					FOR DELETE AS

					delete from ASRSysBatchJobDetails WHERE BatchJobNameID IN (SELECT ID FROM Deleted)
					delete from ASRSysBatchJobAccess WHERE ID IN (SELECT ID FROM Deleted)'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 101 of 120 - Updating Delete Trigger for Calendar Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysCalendarReports]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysCalendarReports]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCalendarReports ON dbo.ASRSysCalendarReports 
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysCalendarReportEvents WHERE ASRSysCalendarReportEvents.CalendarReportID IN (SELECT ID FROM deleted)
    						DELETE FROM ASRSysCalendarReportOrder WHERE ASRSysCalendarReportOrder.CalendarReportID IN (SELECT ID FROM deleted)
						DELETE FROM ASRSysCalendarReportAccess WHERE ASRSysCalendarReportAccess.ID IN (SELECT ID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 102 of 120 - Adding Delete Trigger for Cross Tabs'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysCrossTab]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysCrossTab]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCrossTab ON dbo.ASRSysCrossTab 
					FOR DELETE AS

					delete from ASRSysCrossTabAccess WHERE ID IN (SELECT CrossTabID FROM Deleted)'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 103 of 120 - Updating Delete Trigger for Custom Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysCustomReportsName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysCustomReportsName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCustomReportsName ON ASRSysCustomReportsName
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysCustomReportsDetails WHERE CustomReportID IN (SELECT ID FROM Deleted)
						DELETE FROM ASRSysCustomReportsChildDetails WHERE CustomReportID IN (SELECT ID FROM Deleted)
						delete from ASRSysCustomReportAccess WHERE ID IN (SELECT ID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 104 of 120 - Updating Delete Trigger for Data Transfer'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysDataTransferName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysDataTransferName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysDataTransferName ON ASRSysDataTransferName
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysDataTransferColumns WHERE DataTransferID IN (SELECT DataTransferID FROM Deleted)
						delete from ASRSysDataTransferAccess WHERE ID IN (SELECT DataTransferID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 105 of 120 - Updating Delete Trigger for Export'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysExportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysExportName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysExportName ON ASRSysExportName
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysExportDetails WHERE ExportID IN (SELECT ID FROM Deleted)
						delete from ASRSysExportAccess WHERE ID IN (SELECT ID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 106 of 120 - Updating Delete Trigger for Global Functions'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysGlobalFunctions]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysGlobalFunctions]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysGlobalFunctions ON ASRSysGlobalFunctions
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysGlobalItems WHERE FunctionID IN (SELECT FunctionID FROM Deleted)
						delete from ASRSysGlobalAccess WHERE ID IN (SELECT FunctionID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 107 of 120 - Updating Delete Trigger for Import'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysImportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysImportName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysImportName ON ASRSysImportName
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysImportDetails WHERE ImportID IN (SELECT ID FROM Deleted)
						delete from ASRSysImportAccess WHERE ID IN (SELECT ID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 108 of 120 - Updating Delete Trigger for Mail Merge'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysMailMergeName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysMailMergeName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysMailMergeName ON ASRSysMailMergeName
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysMailMergeColumns WHERE MailMergeID IN (SELECT MailMergeID FROM Deleted)
						delete from ASRSysMailMergeAccess WHERE ID IN (SELECT MailMergeID FROM Deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 109 of 120 - Updating Delete Trigger for Match Reports'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysMatchReportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysMatchReportName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysMatchReportName ON dbo.ASRSysMatchReportName 
					FOR DELETE AS

					delete from ASRSysExpressions where type in (15, 16, 17) and (
					exprid IN (select RequiredExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)) OR
					exprid IN (select PreferredExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)) OR
					exprid IN (select MatchScoreExprID from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)))

					delete from ASRSysMatchReportTables WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)
					delete from ASRSysMatchReportDetails WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)
					delete from ASRSysMatchReportBreakdown WHERE MatchReportID IN (SELECT MatchReportID FROM Deleted)
					delete from ASRSysMatchReportAccess WHERE ID IN (SELECT MatchReportID FROM Deleted)'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 110 of 120 - Updating Delete Trigger for Record Profile'

		if exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysRecordProfileName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysRecordProfileName]

		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysRecordProfileName ON dbo.ASRSysRecordProfileName 
					FOR DELETE 
					AS
					BEGIN
						DELETE FROM ASRSysRecordProfileDetails WHERE ASRSysRecordProfileDetails.RecordProfileID IN (SELECT recordProfileID FROM deleted)
    						DELETE FROM ASRSysRecordProfileTables WHERE ASRSysRecordProfileTables.RecordProfileID IN (SELECT recordProfileID FROM deleted)
						DELETE FROM ASRSysRecordProfileAccess WHERE ID IN (SELECT recordProfileID FROM deleted)
					END'
		EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 111 of 120 - Update System Permissions to re-arrange layout'

UPDATE ASRSysPermissionItems SET listorder = '10' WHERE itemid = 79
UPDATE ASRSysPermissionItems SET listorder = '20' WHERE itemid = 78
UPDATE ASRSysPermissionItems SET listorder = '30' WHERE itemid = 77
UPDATE ASRSysPermissionItems SET listorder = '40' WHERE itemid = 88
UPDATE ASRSysPermissionItems SET listorder = '10' WHERE itemid = 81
UPDATE ASRSysPermissionItems SET listorder = '20' WHERE itemid = 80


/* ------------------------------------------------------------- */
PRINT 'Step 112 of 120 - Deleting ASRSysSystemSettings for Event Log Email'

DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'development' AND ASRSysSystemSettings.SettingKey = 'eventlog_email_1' 
DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'development' AND ASRSysSystemSettings.SettingKey = 'eventlog_email_2' 
DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'development' AND ASRSysSystemSettings.SettingKey = 'eventlog_email_3' 
DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'development' AND ASRSysSystemSettings.SettingKey = 'eventlog_email_4' 
DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'development' AND ASRSysSystemSettings.SettingKey = 'eventlog_email_enable' 
DELETE FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.Section = 'email' AND ASRSysSystemSettings.SettingKey = 'event log send' 


/* ------------------------------------------------------------- */
PRINT 'Step 113 of 120 - Deleting Obsolete Absence Module Setup Parameters'

DELETE FROM ASRSysModuleSetup
WHERE ModuleKey = 'MODULE_ABSENCE'
AND ParameterKey IN ('Param_FieldTypeInclude', 'Param_FieldTypeBradfordIndex')



/* ------------------------------------------------------------- */
PRINT 'Step 114 of 120 - Adding tableID to Email Queue'

	/* Add new column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysEmailQueue')
	and name = 'TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysEmailQueue ADD 
					[TableID] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 115 of 120 - Updating email batch stored procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASREmailImmediate]

execute('CREATE PROCEDURE spASREmailImmediate(@Username varchar(255))  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue varchar(8000),
		@RecDescID int,
		@RecDesc varchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int,
		@blnEnabled int,
		@RecalculateRecordDesc bit,
		@TableID int,
		@RecipTo varchar(4000),
		@TempText nvarchar(4000)

	/* Loop through all entries which are to be sent */
	DECLARE emailqueue_cursor
	CURSOR LOCAL FAST_FORWARD FOR 
		SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID
		FROM ASRSysEmailQueue
		WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
		And (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
		ORDER BY DateDue

	OPEN emailqueue_cursor
	FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

	WHILE (@@fetch_status = 0)
	BEGIN

		IF @RecalculateRecordDesc = 1
			BEGIN	
				IF @ColumnID > 0
					BEGIN
						SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
							 (SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))
					END
				ELSE IF @TableID > 0
					BEGIN			
						SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = @TableID)
					END
		
				SET @RecDesc = ''''
				SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					EXEC @sSQL @RecDesc OUTPUT, @Recordid
				END
			END

		/* Add table name to record descripion if it is a table entry */
		IF @TableID > 0
			BEGIN
				SELECT @TempText = (SELECT TableName FROM ASRSYSTables WHERE TableID = @TableID)
				SET @RecDesc = @TempText + '' : '' + @RecDesc
			END		
	
		IF @ColumnID > 0
			BEGIN
				SELECT @sSQL = ''spASRSysEmailSend_'' + convert(varchar,@LinkID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SELECT @emailDate = getDate()
						EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''
					END
			END
		ELSE IF @TableID > 0
			BEGIN
				SET @sSQL = ''spASRSysEmailAddr_'' + convert(varchar,@LinkID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SELECT @emailDate = getDate()
						EXEC @hResult = @sSQL @RecipTo OUTPUT,0
						EXEC @hResult = master.dbo.xp_sendmail  @recipients=@RecipTo,  @subject=@columnvalue,  @message=@RecDesc, @no_output=''True''
					END
			END

		IF @hResult = 0
		BEGIN
			UPDATE ASRSysEmailQueue SET DateSent = @emailDate
			WHERE QueueID = @QueueID
		END

		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

	END
	CLOSE emailqueue_cursor
	DEALLOCATE emailqueue_cursor

END')

/* ------------------------------------------------------------- */
PRINT 'Step 116 of 120 - Adding table email link stored procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAuditTable]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAuditTable]

execute('CREATE PROCEDURE sp_ASRAuditTable (
	@piTableID int,
	@piRecordID int,
	@psRecordDesc varchar(255),
	@psValue varchar(255))
AS
BEGIN	

	DECLARE @sTableName varchar(8000)

	/* Get the table name for the given column. */
	SELECT @sTableName = tablename 
	FROM asrsystables
	WHERE asrsystables.tableid = @piTableID

	IF @sTableName IS NULL SELECT @sTableName = ''<Unknown>''

	/* Insert a record into the Audit Trail table. */
	INSERT INTO ASRSysAuditTrail 
		(userName, dateTimeStamp, tablename, recordID, recordDesc, columnname, oldValue, newValue,ColumnID, Deleted)
	VALUES 
		(user, getDate(), @sTableName, @piRecordID, @psRecordDesc, '''', '''', @psValue,0, 0)

END')


/* ------------------------------------------------------------- */
PRINT 'Step 117 of 120 - Changing separator column type to bit'
EXECUTE('ALTER TABLE asrsyscolumns ALTER COLUMN use1000separator bit')


/* ------------------------------------------------------------- */
PRINT 'Step 118 of 120 - Updating Permission Items'

UPDATE ASRSysPermissionItems SET Description = 'Absence Breakdown' WHERE ItemID = 95
UPDATE ASRSysPermissionItems SET Description = 'Absence Calendar' WHERE ItemID = 96
UPDATE ASRSysPermissionItems SET Description = 'Bradford Factor' WHERE ItemID = 97
UPDATE ASRSysPermissionItems SET Description = 'Stability Index' WHERE ItemID = 98
UPDATE ASRSysPermissionItems SET Description = 'Turnover Report' WHERE ItemID = 99


/* ------------------------------------------------------------- */
PRINT 'Step 119 of 120 - Updating case sensitive compare'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRCaseSensitiveCompare]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCaseSensitiveCompare]

execute('CREATE PROCEDURE sp_ASRCaseSensitiveCompare
(
	@pfResult		bit OUTPUT,
	@psStringA 		varchar(8000),
	@psStringB		varchar(8000)
)
AS
BEGIN
	/* Return 1 if the given string are exactly equal. */
	DECLARE @iPosition	integer

	SET @pfResult = 0

	IF (@psStringA IS NULL) AND (@psStringB IS NULL) SET @pfResult = 1

	IF (@pfResult = 0) AND (NOT @psStringA IS NULL) AND (NOT @psStringB IS NULL)
	BEGIN

		/* LEN() does not look at trailing spaces, so force it too by adding some quotations at the end. */
		IF LEN(@psStringA+'''''''') = LEN(@psStringB+'''''''')
		BEGIN
			SET @pfResult = 1

			SET @iPosition = 1
			WHILE @iPosition <= LEN(@psStringA) 
			BEGIN
				IF ASCII(SUBSTRING(@psStringA, @iPosition, 1)) <> ASCII(SUBSTRING(@psStringB, @iPosition, 1))
				BEGIN
					SET @pfResult = 0
					BREAK
				END

				SET @iPosition = @iPosition + 1
			END
		END
	END
END')


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 120 of 120 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.10')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.10')


/* ------------------------------------------- */
/* Grant permission to email stored procedures */
/* ------------------------------------------- */
SELECT @NVarCommand = 'USE master
GRANT ALL ON master..xp_StartMail TO public
GRANT ALL ON master..xp_SendMail TO public'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE '+@DBName
EXEC sp_executesql @NVarCommand

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.10 Of HR Pro'
