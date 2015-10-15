/* --------------------------------------------------- */
/* Update the database from version 8.1 to version 8.2 */
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
IF (@sDBVersion <> '8.1') and (@sDBVersion <> '8.2')
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


/* ------------------------------------------------------- */
PRINT 'Step - Mail Merge additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'UploadTemplate')
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName ADD UploadTemplate varbinary(MAX) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'UploadTemplateName')
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName ADD UploadTemplateName nvarchar(255) NULL;';


/* ------------------------------------------------------- */
PRINT 'Step - Audit Changes'
/* ------------------------------------------------------- */

ALTER TABLE [ASRSysAuditAccess] ALTER COLUMN [ComputerName] varchar(255);


/* ------------------------------------------------------- */
PRINT 'Step - Calculation Changes'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_weekdaysbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_weekdaysbetweentwodates];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_weekdaysbetweentwodates](
			@datefrom AS datetime,
			@dateto AS datetime)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
	
			DECLARE @result integer;

			SELECT @result = CASE 
				WHEN DATEDIFF (day, @datefrom, @dateto) < 0 THEN 0
				ELSE (DATEDIFF(dd, @datefrom, @dateto) + 1)
					- (DATEDIFF(wk, @datefrom, @dateto) * 2)
					- (CASE WHEN DATEPART(dw, @datefrom) = 1 THEN 1 ELSE 0 END)
					- (CASE WHEN DATEPART(dw, @dateto) = 7 THEN 1 ELSE 0 END)
					END;
				
			RETURN ISNULL(@result,0);
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_fieldlastchangedate]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_fieldlastchangedate];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_fieldlastchangedate](
			@colrefID	varchar(32),
			@recordID	integer
		)
		RETURNS datetime
		WITH SCHEMABINDING
		AS
		BEGIN

			DECLARE @result		datetime,
					@tableid	integer,
					@columnid	integer;
		
			SET @tableid = SUBSTRING(@colrefID, 1, 8);
			SET @columnid = SUBSTRING(@colrefID, 10, 8);

			SELECT TOP 1 @result = DATEADD(dd, 0, DATEDIFF(dd, 0, [DateTimeStamp])) FROM dbo.[ASRSysAuditTrail]
				WHERE [ColumnID] = @columnid AND [TableID] = @tableID
					AND @recordID = [RecordID]
				ORDER BY [DateTimeStamp] DESC ;

			RETURN @result;

		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysStringToTable]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsysStringToTable];

	EXEC sp_executesql N'CREATE FUNCTION dbo.[udfsysStringToTable] (           
		  @String nvarchar(MAX),
		  @delimiter nvarchar(2))
	RETURNS @Table TABLE( Splitcolumn nvarchar(MAX)) 
	BEGIN

		DECLARE @Xml AS XML;
		SET @Xml = cast((''<A>''+replace(@String,@delimiter,''</A><A>'')+''</A>'') AS XML);

		INSERT INTO @Table SELECT LTRIM(RTRIM(A.value(''.'', ''nvarchar(max)''))) AS [Column] FROM @Xml.nodes(''A'') AS FN(A);
		RETURN;

	END'


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

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XSDFileName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XSDFileName nvarchar(255) NULL, PreserveTransformPath bit, PreserveXSDPath bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitXMLNodesFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitXMLNodesFile bit;';

	ALTER TABLE [ASRSysExportName] ALTER COLUMN [HeaderText] varchar(MAX);
	ALTER TABLE [ASRSysExportName] ALTER COLUMN [FooterText] varchar(MAX);

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'StripDelimiterFromData')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD StripDelimiterFromData bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFile bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFileSize')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFileSize int;';


/* ------------------------------------------------------- */
PRINT 'Step - Workflow additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElements', 'U') AND name = 'RequiresAuthentication')
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElements ADD RequiresAuthentication bit NULL;';


/* ------------------------------------------------------- */
PRINT 'Step - Branding'
/* ------------------------------------------------------- */

	EXEC sp_executesql N'UPDATE ASRSysPermissionCategories SET [description] = ''OpenHR Web'' WHERE categoryID = 19';
	EXEC sp_executesql N'UPDATE ASRSysPermissionItems SET [description] = ''OpenHR Web'' WHERE itemID = 4';


/* ------------------------------------------------------- */
PRINT 'Step - Database Hardening'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowValidateService]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowValidateService];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRWorkflowValidateService](@allow bit OUTPUT)
		WITH ENCRYPTION
		AS
		BEGIN

			DECLARE @iUserGroupID		integer,
				@sUserGroupName			sysname,
				@fSysSecMgr				bit,
				@sActualUserName		sysname;

			EXEC [dbo].[spASRIntGetActualUserDetails]
				@sActualUserName OUTPUT,
				@sUserGroupName OUTPUT,
				@iUserGroupID OUTPUT;

			SELECT @allow = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
				FROM ASRSysGroupPermissions gp
					INNER JOIN ASRSysPermissionItems pi ON gp.itemID = pi.itemID
					INNER JOIN ASRSysPermissionCategories pc ON pi.categoryID = pc.categoryID
					INNER JOIN sys.database_principals u ON gp.groupName = u.name
				WHERE u.principal_id = @iUserGroupID
					AND pi.itemKey = ''SYSTEMMANAGER'' AND gp.permitted = 1 AND pc.categorykey = ''MODULEACCESS'';

		END'


	EXECUTE sp_executeSQL N'CREATE PROCEDURE #spASRTempHardenTables(@tablename nvarchar(MAX), @adminrole nvarchar(MAX), @allusersHaveSelect bit)
	AS
	BEGIN

		DECLARE @NVarCommand nvarchar(MAX) = '''';

		SELECT @NVarCommand = @NVarCommand + ''REVOKE SELECT, UPDATE, INSERT, DELETE ON ['' + @tablename + ''] TO ['' + U.name + ''];''
			FROM sys.database_permissions P 
			JOIN sys.tables T ON P.major_id = T.object_id 
			JOIN sysusers U ON U.uid = P.grantee_principal_id
			WHERE t.name = @tablename;			
		EXECUTE sp_executeSQL @NVarCommand;

		SET @NVarCommand = ''GRANT SELECT, INSERT, UPDATE, DELETE ON ['' + @tablename + ''] TO ['' + @adminrole + ''];'';
		EXECUTE sp_executeSQL @NVarCommand;	

		IF @allusersHaveSelect = 1
		BEGIN
			SET @NVarCommand = ''GRANT SELECT ON ['' + @tablename + ''] TO [ASRSysGroup];'';
			EXECUTE sp_executeSQL @NVarCommand;
		END
	END';


	IF EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmins' AND type = 'R')
	BEGIN
		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_droprolemember @rolename = [ASRSysAdmins], @membername = [' + members.[name] + '];'
			FROM sys.database_role_members AS rolemembers
				JOIN sys.database_principals AS roles ON roles.[principal_id] = rolemembers.[role_principal_id]
				JOIN sys.database_principals AS members ON members.[principal_id] = rolemembers.[member_principal_id]
			WHERE roles.[name]='ASRSysAdmins';

		EXEC sp_executeSQL @NVarCommand;
		EXEC sp_executeSQL N'DROP ROLE [ASRSysAdmins];'
	END

	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmin' AND type = 'R')
	BEGIN
		EXECUTE sp_executesql N'CREATE ROLE [ASRSysAdmin] AUTHORIZATION [dbo];';

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_addrolemember @rolename = [ASRSysAdmin], @membername = [' + gp.groupName + '];'
			FROM ASRSysGroupPermissions gp
			INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
			WHERE pi.itemID IN (1) AND gp.permitted = 1;
		
		EXECUTE sp_executesql @NVarCommand;

		EXEC #spASRTempHardenTables 'ASRSysColours', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysColumnControlValues', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysConfig', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysControls', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysDiaryLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinksColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinksRecipients', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysFunctionParameters', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysFunctions', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysGroups', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysHistoryScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysKeywords', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysLinkContent', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysModuleRelatedColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysModuleSetup', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOperatorParameters', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOperators', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookEvents', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookFolders', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinksColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinksDestinations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysPermissionCategories', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysPictures', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysRelations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIHiddenGroups', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIntranetLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIViews', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSummaryFields', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTables', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTableTriggers', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTableValidations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViewColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViewScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViews', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'tbsys_MobileFormElements', 'ASRSysAdmin', 0;
		EXEC #spASRTempHardenTables 'tbsys_MobileFormLayout', 'ASRSysAdmin', 0;
		EXEC #spASRTempHardenTables 'tbsys_MobileGroupWorkflows', 'ASRSysAdmin', 0;

	END
	
	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysWorkflowAdmin' AND type = 'R')
	BEGIN
		EXECUTE sp_executesql N'CREATE ROLE [ASRSysWorkflowAdmin] AUTHORIZATION [dbo];';

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_addrolemember @rolename = [ASRSysWorkflowAdmin], @membername = [' + gp.groupName + '];'
			FROM ASRSysGroupPermissions gp
			INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
			WHERE pi.itemID IN (1, 151, 152) AND gp.permitted = 1;
		EXECUTE sp_executesql @NVarCommand;

		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementItems', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementItemValues', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElements', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementValidations', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstances', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstanceSteps', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstanceValues', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowLinks', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowQueue', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowQueueColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowStepDelegation', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowTriggeredLinkColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowTriggeredLinks', 'ASRSysWorkflowAdmin', 0;

	END



	DROP PROCEDURE #spASRTempHardenTables


PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '8.2';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.2.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.2.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.2')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.2 Of OpenHR'