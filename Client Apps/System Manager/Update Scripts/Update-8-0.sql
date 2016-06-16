/* --------------------------------------------------- */
/* Update the database from version 7.0 to version 8.0 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(MAX),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @sSPCode nvarchar(MAX)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON;
SET @DBName = DB_NAME();

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '7.0') and (@sDBVersion <> '8.0')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
IF (@iSQLVersion < 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step - Changes to Shared Table Transfer for PAE Defaults'
/* ------------------------------------------------------------- */
	
	-- Add new mappings for Employee transfer
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 224 AND TransferTypeID = 0
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (224,0,0,''PAE Worker Postponement Applies'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (225,0,0,''PAE Worker Postponement End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (226,0,0,''PAE EJ Postponement Applies'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (227,0,0,''PAE Postponement Notice Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (228,0,0,''PAE Default Pension Scheme'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------- */
PRINT 'Step - Generare sysprotects cache'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysProtectsCache]') AND type = N'U')
		DROP TABLE [dbo].[ASRSysProtectsCache];

	SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysProtectsCache]
	(
		[ID] int NOT NULL,
		[Action] tinyint NOT NULL,
		[Columns] varbinary(8000),
		[ProtectType] int NOT NULL,
		[UID] integer NOT NULL
	)'
	EXEC sp_executesql @NVarCommand

	IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[[ASRSysProtectsCache]]') AND name = N'IDX_ProtectsCache_UID')
		DROP INDEX [IDX_ProtectsCache_UID] ON [dbo].[ASRSysProtectsCache]
	EXEC sp_executesql N'CREATE CLUSTERED INDEX [IDX_ProtectsCache_UID] ON [dbo].[ASRSysProtectsCache] ([UID] ASC)';

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRPostSystemSave]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spASRPostSystemSave];

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRPostSystemSave]
		AS
		BEGIN

			IF OBJECT_ID(''ASRSysProtectsCache'') IS NOT NULL 
				DELETE FROM ASRSysProtectsCache;

			INSERT ASRSysProtectsCache 
			SELECT ID, Action, Columns, ProtectType , uid
			   FROM sysprotects;

		END'
	EXEC sp_executesql @NVarCommand;
	EXEC sp_executesql N'spASRPostSystemSave';


/* ------------------------------------------------------- */
PRINT 'Step - Ensure the required permissions are granted'
/* ------------------------------------------------------- */

	DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
	SELECT sysobjects.name, sysobjects.xtype
	FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
	WHERE (((sysobjects.xtype = 'p' OR sysobjects.xtype = 'pc') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%' OR sysobjects.name LIKE 'spadmin%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')
	--IF (@@ERROR <> 0) goto QuitWithRollback
 
	OPEN curObjects
	FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
	WHILE (@@fetch_status = 0)
	BEGIN
		IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'PC' OR rtrim(@sObjectType) = 'FN'
		BEGIN
			IF @sObject LIKE 'sp_ASRExpr_%' OR @sObject LIKE 'sp_ASRDfltExpr_%' OR @sObject LIKE 'spASREmail_%' OR @sObject LIKE 'spASRUpdateOLEField_%'
				SET @NVarCommand = 'REVOKE EXECUTE ON [' + @sObject + '] TO [ASRSysGroup]'
			ELSE              
				SET @NVarCommand = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
			END
		ELSE
		BEGIN
			SET @NVarCommand = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
		END
 
		EXECUTE sp_executeSQL @NVarCommand
	
		FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
	END
	CLOSE curObjects
	DEALLOCATE curObjects
 

	/* For the reset password functionality */
	GRANT EXEC ON spadmin_commitresetpassword TO [openhr2iis]


/* ------------------------------------------------------- */
PRINT 'Step - Remove hotfix for resetting passwords '
/* ------------------------------------------------------- */

DECLARE @DeleteJobID uniqueidentifier;
SELECT @DeleteJobID = [job_id] FROM msdb.dbo.sysjobs WHERE name = 'OpenHR Password Resets'
IF NOT @DeleteJobID IS NULL
	EXEC msdb.dbo.sp_delete_job @job_id=@DeleteJobID, @delete_unused_schedule=1


/* ------------------------------------------------------- */
PRINT 'Step - New functions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysFunctions', 'U') AND name = 'ExcludeFromSysMgr')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysFunctions ADD ExcludeFromSysMgr bit NULL;';
		EXEC sp_executesql N'UPDATE ASRSysFunctions SET ExcludeFromSysMgr = 0;';

	END

	DELETE FROM ASRSysFunctions WHERE functionID IN (78, 79, 80)
	DELETE FROM ASRSysFunctionParameters WHERE functionID IN (78, 79, 80)

	EXEC sp_executesql N'INSERT ASRSysFunctions (FunctionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ExcludeFromSysMgr)
		VALUES (78, ''Last Run Date'', 4, 0, ''Date/Time'', '''',0, 1, 0, 1);'
	EXEC sp_executesql N'INSERT ASRSysFunctions (FunctionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ExcludeFromSysMgr)
		VALUES (79, ''Base Record ID'', 2, 0, ''General'', '''',0, 1, 0, 0);'
	EXEC sp_executesql N'INSERT ASRSysFunctions (FunctionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, UDF, ExcludeFromSysMgr)
		VALUES (80, ''Create Exact Date'', 4, 0, ''Date/Time'', '''',0, 1, 0, 0);'
	INSERT ASRSysFunctionParameters (functionID, parameterIndex, parameterType, parameterName)
		VALUES (80, 1, 2, '<Day>');
	INSERT ASRSysFunctionParameters (functionID, parameterIndex, parameterType, parameterName)
		VALUES (80, 2, 2, '<Month>');
	INSERT ASRSysFunctionParameters (functionID, parameterIndex, parameterType, parameterName)
		VALUES (80, 3, 2, '<Year>');


	-- Parentheses
	DELETE FROM tbstat_componentcode WHERE id IN (79, 80) AND isoperator = 0

	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount], [maketypesafe])
		VALUES (79, '[ID]', 4, 'Base Record ID', 0, 0, 0, 0);
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount], [maketypesafe])
		VALUES (80, 'dbo.udfASRCreateDate({0}, {1}, {2})', 4, 'Create Exact Date', 0, 0, 0, 0);


	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[udfASRCreateDate]') AND type in (N'FN'))
		DROP FUNCTION [dbo].[udfASRCreateDate];

	SELECT @NVarCommand = 'CREATE FUNCTION dbo.udfASRCreateDate(@day float, @month float, @year float)
		RETURNS datetime
		AS
		BEGIN

			DECLARE @date varchar(20);

			IF @day < 1 OR @month < 1 OR @year < 1 OR @month > 12 OR @day > 31 OR @year > 9999  RETURN NULL;
			SET @date = CONVERT(varchar(2), @month) + ''/'' + CONVERT(varchar(2), @day) + ''/'' + CONVERT(varchar(4), @year);

			IF ISDATE(@date) = 0
				RETURN NULL;

			RETURN @date;

		END';
		EXEC sp_executesql @NVarCommand;


	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[udfsys_FieldChangedSinceLastExport]') AND type in (N'FN'))
		DROP FUNCTION [dbo].[udfsys_FieldChangedSinceLastExport];

	SELECT @NVarCommand = 'CREATE FUNCTION [dbo].[udfsys_FieldChangedSinceLastExport](
		@columnID	integer,
		@FromDate	datetime,
		@recordID	integer
	)
	RETURNS bit
	AS
	BEGIN

		DECLARE @result	bit = 0;
		
		SELECT @result = CASE WHEN
				EXISTS(SELECT [DateTimeStamp] FROM dbo.[ASRSysAuditTrail]
					WHERE [ColumnID] = @columnid
					AND @recordID = [RecordID] 
					AND [DateTimeStamp] >= @FromDate)
				THEN 1 ELSE 0 END;

		RETURN @result;
	END';
	EXEC sp_executesql @NVarCommand;



/* ------------------------------------------------------- */
PRINT 'Step - Export additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'TransformFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD TransformFile nvarchar(MAX) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XMLDataNodeName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XMLDataNodeName nvarchar(50) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LastSuccessfulOutput')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LastSuccessfulOutput datetime NULL;';
		
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'AuditChangesOnly')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD AuditChangesOnly bit NULL;';


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '8.0';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.0.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.0.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.0')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
EXEC dbo.spsys_setsystemsetting 'database', 'refreshstoredprocedures', 1;


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF;

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.0 Of OpenHR'