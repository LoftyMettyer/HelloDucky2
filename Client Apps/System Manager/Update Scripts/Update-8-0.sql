/* --------------------------------------------------- */
/* Update the database from version 7.0 to version 8.0 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion int,
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
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
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