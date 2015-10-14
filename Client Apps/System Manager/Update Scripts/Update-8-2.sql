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


	IF EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmins' AND type = 'R')
	BEGIN
		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'ALTER ROLE [ASRsysAdmins] DROP MEMBER ['+ members.[name]+ '];'
			FROM sys.database_role_members AS rolemembers
				JOIN sys.database_principals AS roles ON roles.[principal_id] = rolemembers.[role_principal_id]
				JOIN sys.database_principals AS members ON members.[principal_id] = rolemembers.[member_principal_id]
			WHERE roles.[name]='ASRsysAdmins';

		EXEC sp_executeSQL @NVarCommand;
		EXEC sp_executeSQL N'DROP ROLE [ASRSysAdmins];'
	END

	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmin' AND type = 'R')
	BEGIN
		EXECUTE sp_executesql N'CREATE ROLE [ASRSysAdmin] AUTHORIZATION [dbo];';

		SET @NVarCommand = '';
		SELECT DISTINCT @NVarCommand = @NVarCommand + 'ALTER ROLE [ASRSysAdmin] ADD MEMBER ' + gp.groupName + ';'
			FROM ASRSysGroupPermissions gp
			INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
			WHERE pi.itemID IN (1) AND gp.permitted = 1;
		
		EXECUTE sp_executesql @NVarCommand;

	END




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