
/* -------------------------------------------------- */
/* Update the database from version 26 to version 27. */
/* -------------------------------------------------- */

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
        @NVarCommand nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */
SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

if @sDBVersion = ''
BEGIN
  SELECT @sDBVersion = SystemManagerVersion FROM ASRSysConfig
END


/* Exit if the database is not version 26 or 27. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.27') and (@sDBVersion <> '1.28')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 11 - Create index on system table'


SELECT @iRecCount = count(id) FROM sysindexes WHERE name = 'Section, SettingKey'
if @iRecCount = 0
BEGIN
  CREATE UNIQUE CLUSTERED
    INDEX [Section, SettingKey] ON [dbo].[ASRSysSystemSettings] ([Section], [SettingKey])
END


/* ---------------------------- */

PRINT 'Step 2 of 11 - Amending definition tables'

ALTER TABLE ASRSYSEventlog ALTER COLUMN name varchar(150)


SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
and name = 'Parent1AllRecords'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysCustomReportsName ADD Parent1AllRecords bit null
  ALTER TABLE ASRSysCustomReportsName ADD Parent1Picklist int null
  ALTER TABLE ASRSysCustomReportsName ADD Parent2AllRecords bit null
  ALTER TABLE ASRSysCustomReportsName ADD Parent2Picklist int null

  SELECT @NVarCommand = 'UPDATE asrsyscustomreportsname set Parent1AllRecords = 1, Parent1Picklist = 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsyscustomreportsname set Parent1AllRecords = 0 where Parent1Filter > 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsyscustomreportsname set Parent2AllRecords = 1, Parent2Picklist = 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsyscustomreportsname set Parent2AllRecords = 0 where Parent2Filter > 0'
  EXEC sp_executesql @NVarCommand

END


SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysExportName')
and name = 'Parent1AllRecords'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysExportName ADD Parent1AllRecords bit null
  ALTER TABLE ASRSysExportName ADD Parent1Picklist int null
  ALTER TABLE ASRSysExportName ADD Parent2AllRecords bit null
  ALTER TABLE ASRSysExportName ADD Parent2Picklist int null

  SELECT @NVarCommand = 'UPDATE asrsysexportname set Parent1AllRecords = 1, Parent1Picklist = 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsysexportname set Parent1AllRecords = 0 where Parent1Filter > 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsysexportname set Parent2AllRecords = 1, Parent2Picklist = 0'
  EXEC sp_executesql @NVarCommand
  SELECT @NVarCommand = 'UPDATE asrsysexportname set Parent2AllRecords = 0 where Parent2Filter > 0'
  EXEC sp_executesql @NVarCommand

END


/* ---------------------------- */

PRINT 'Step 3 of 11 - Amending Summary Field Definition'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysSummaryFields')
and name = 'StartOfColumn'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysSummaryFields ADD StartOfColumn bit null
END


SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysTables')
and name = 'ManualSummaryColumnBreaks'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysTables ADD ManualSummaryColumnBreaks bit null
END

/* ---------------------------- */

PRINT 'Step 4 of 11 - Amending Operators Table'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysOperators')
and name = 'CastAsFloat'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysOperators ADD CastAsFloat bit not null default 0
  SELECT @NVarCommand = 'Update ASRSysOperators set CastAsFloat = 1 where OperatorID = 4'
  exec sp_executesql @NVarCommand
END

/* ---------------------------- */

PRINT 'Step 5 of 11 - Updating Intranet permission items'


update ASRSysPermissionItems Set description = 'Intranet (Full access)' where itemid = 4

delete from ASRSysPermissionItems where itemid = 100
insert ASRSysPermissionItems(itemID, description, listOrder, categoryID, itemKey)
values(100, 'Intranet (Self-service access)', 70, 1, 'INTRANET_SELFSERVICE')

/* ---------------------------- */

PRINT 'Step 6 of 11 - Adding new System Permissions information.'

DECLARE @res1 int,
        @res2 int

-- Find out if the old Standard Reports Run setting still exists.
SELECT DISTINCT @res1 = COUNT(*) FROM ASRSysGroupPermissions WHERE ItemID = 51

-- Find out if the new itemIDs exist.
SELECT DISTINCT @res2 = COUNT(*) FROM ASRSysGroupPermissions WHERE ItemID IN (95,96,97,98,99)

IF @res1 > 0 AND @res2 = 0
  BEGIN
    -- Insert the new items into the ASRSysPermissionItems table.
    INSERT INTO ASRSysPermissionItems (itemid, description, listOrder, categoryID, itemKey)
       VALUES                     (95, 'Run Absence Breakdown', 10, 13, 'RUN_AB')
    INSERT INTO ASRSysPermissionItems (itemid, description, listOrder, categoryID, itemKey)
       VALUES                     (96, 'Run Absence Calender', 20, 13, 'RUN_AC')
    INSERT INTO ASRSysPermissionItems (itemid, description, listOrder, categoryID, itemKey)
       VALUES                     (97, 'Run Bradford Factor', 30, 13, 'RUN_BF')
    INSERT INTO ASRSysPermissionItems (itemid, description, listOrder, categoryID, itemKey)
       VALUES                     (98, 'Run Stability Index', 40, 13, 'RUN_SI')
    INSERT INTO ASRSysPermissionItems (itemid, description, listOrder, categoryID, itemKey)
       VALUES                     (99, 'Run Turnover Report', 50, 13, 'RUN_TR')

    -- Update the permissions based upon the old Standard Reports Run setting.
    INSERT INTO ASRSysGroupPermissions 
       SELECT DISTINCT 95,GroupName,Permitted FROM ASRSysGroupPermissions WHERE ItemID = 51
    -- Except for the Absence Calender this should be defaulted to 'Permitted'.
    INSERT INTO ASRSysGroupPermissions 
       SELECT DISTINCT 96,GroupName,1 FROM ASRSysGroupPermissions WHERE ItemID = 51
    INSERT INTO ASRSysGroupPermissions 
       SELECT DISTINCT 97,GroupName,Permitted FROM ASRSysGroupPermissions WHERE ItemID = 51
    INSERT INTO ASRSysGroupPermissions 
       SELECT DISTINCT 98,GroupName,Permitted FROM ASRSysGroupPermissions WHERE ItemID = 51
    INSERT INTO ASRSysGroupPermissions 
       SELECT DISTINCT 99,GroupName,Permitted FROM ASRSysGroupPermissions WHERE ItemID = 51

    -- Remove old systen permission items.
    DELETE FROM ASRSysPermissionItems WHERE ItemID = 51
    DELETE FROM ASRSysGroupPermissions WHERE ItemID = 51
  END 



/* ---------------------------- */

/* Start of Currency Module information. */
PRINT 'Step 7 of 11 - Adding Currency Module information.'

DELETE FROM ASRSysFunctions WHERE functionID = 51

INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
       VALUES                (51, 'Convert Currency', 2, 0, 'General', 'sp_ASRFn_ConvertCurrency', 0, 1)



DELETE FROM ASRSysFunctionParameters WHERE functionID = 51
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (51, 1, 2, '<Numeric Value>')
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (51, 2, 1, '<From Currency>')
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (51, 3, 1, '<To Currency>')


SELECT @iRecCount = count(ModuleKey) FROM ASRSysModuleSetup WHERE ModuleKey = 'MODULE_CURRENCY'

if @iRecCount = 0
BEGIN
  INSERT INTO ASRSysModuleSetup  (ModuleKey, ParameterKey, ParameterValue, ParameterType)
         VALUES                  ('MODULE_CURRENCY', 'Param_ConversionTable', 0, 'PType_TableID')
  INSERT INTO ASRSysModuleSetup  (ModuleKey, ParameterKey, ParameterValue, ParameterType)
         VALUES                  ('MODULE_CURRENCY', 'Param_CurrencyNameColumn', 0, 'PType_ColumnID')
  INSERT INTO ASRSysModuleSetup  (ModuleKey, ParameterKey, ParameterValue, ParameterType)
         VALUES                  ('MODULE_CURRENCY', 'Param_ConversionValueColumn', 0, 'PType_ColumnID')
  INSERT INTO ASRSysModuleSetup  (ModuleKey, ParameterKey, ParameterValue, ParameterType)
         VALUES                  ('MODULE_CURRENCY', 'Param_DecimalColumn', 0, 'PType_ColumnID')
END


DELETE FROM ASRSysFunctions WHERE functionID = 51

INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
       VALUES                (51, 'Convert Currency', 2, 0, 'Numeric', 'sp_ASRFn_ConvertCurrency', 0, 1)


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ConvertCurrency]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ConvertCurrency]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_ConvertCurrency
(
	  @pfResult	Float OUTPUT
	, @pfValue	Float
	, @psFromCurr	VarChar(8000)
	, @psToCurr	VarChar(8000)
)
AS
BEGIN
	DECLARE
			  @sCConvTable 	SysName
			, @sCConvExRateCol	SysName
			, @sCConvCurrDescCol	SysName
			, @sCConvDecCol	SysName
			, @sCommandString	nvarchar(4000)
			, @sParamDefinition	nvarchar(500)
	
	-- Get the name of the Currency Conversion table and Currency Description column.
	SELECT @sCConvCurrDescCol = ASRSysColumns.ColumnName, @sCConvTable = ASRSysTables.TableName 
	FROM ASRSysModuleSetup 
     		INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
               		 INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID 
	WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_CurrencyNameColumn''

	-- Get the name of the Exchange Rate column.
	SELECT @sCConvExRateCol = ASRSysColumns.ColumnName
	FROM ASRSysModuleSetup 
     		INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
        WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_ConversionValueColumn''

	-- Get the name of the Decimals column.
	SELECT @sCConvDecCol = ASRSysColumns.ColumnName
	FROM ASRSysModuleSetup 
     		INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
        WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_DecimalColumn''

	IF (NOT @sCConvTable IS NULL) AND (NOT @sCConvCurrDescCol IS NULL) AND (NOT @sCConvExRateCol IS NULL) AND (NOT @sCConvDecCol IS NULL)
	  -- Create the SQL string that returns the Coverted value.
	  BEGIN
	    SET @sCommandString = ''SELECT @pfResult = (SELECT ROUND((('' + LTRIM(RTRIM(STR(@pfValue,20,6)))
									    + '' / '' 
									    + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol + '' FROM '' + @sCConvTable + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psFromCurr + '''''')) '' 
									    + '' * '' 
									    + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol + '' FROM '' + @sCConvTable + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + '''''')) '' 
									    + '' , '' 
									    + '' (SELECT '' + @sCConvTable + ''.'' + @sCConvDecCol + '' FROM '' + @sCConvTable + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + ''''''))) ''

	    SET @sParamDefinition = N''@pfResult float output''

            EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfResult output
          END
	ELSE
	  SET @pfResult = NULL
END'

exec sp_executesql @NVarCommand

/* ---------------------------- */

PRINT 'Step 8 of 11 - Updating Summary Fields Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetSummaryFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetSummaryFields]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRGetSummaryFields (
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
	    	ASRSysSummaryFields.StartOfColumn
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence
END'

exec sp_executesql @NVarCommand



/* ---------------------------- */

PRINT 'Step 9 of 11 - Updating Permissions Stored Procedure'


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRSystemPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRSystemPermission]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRSystemPermission
(
	@pfPermissionGranted 	bit OUTPUT,
	@psCategoryKey	varchar(50),
	@psPermissionKey	varchar(50)
)
AS
BEGIN
	/* Return 1 if the given permission is granted to the current user, 0 if it is not.	*/
	DECLARE @fGranted bit


	/* MH20010222 - This needs to be System_User and not Current_User ! */
	/* TM20011114 - Need to check ''syslogins.loginname'' not ''syslogins.name'' ! */
	SELECT @fGranted = sysAdmin FROM master..syslogins WHERE loginname = SYSTEM_USER

	IF @fGranted = 0
	BEGIN

		SELECT @fGranted = ASRSysGroupPermissions.permitted
		FROM ASRSysGroupPermissions
			INNER JOIN ASRSysPermissionItems 
				ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
			INNER JOIN ASRSysPermissionCategories
				ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID,
		sysusers a
			INNER JOIN sysusers b 
				ON a.uid = b.gid
		WHERE b.name = CURRENT_USER
			AND ASRSysPermissionItems.itemKey = @psPermissionKey
			AND ASRSysGroupPermissions.groupName = a.name
			AND ASRSysPermissionCategories.categoryKey = @psCategoryKey

	END


	IF @fGranted IS NULL
	BEGIN
		SET @fGranted = 0
	END

	SET @pfPermissionGranted = @fGranted
END'

exec sp_executesql @NVarCommand


/* ---------------------------- */

PRINT 'Step 10 of 11 - Updating Permission Descriptions'

Update ASRSysPermissionItems Set Description = 'Run Absence Calendar' Where itemID = 96
Update ASRSysPermissionItems Set Description = 'Batch Log On' Where ItemID = 94


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 11 of 11 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.28')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.28')

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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.28 Of HR Pro'
