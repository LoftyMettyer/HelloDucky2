/* --------------------------------------------------- */
/* Update the database from version 8.2 to version 8.3 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(MAX),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit,
	@newDesktopImageID	integer,
	@picname			varchar(255),
	@picturetype		integer,
	@oldDesktopImageID	integer,
	@newMobileHeaderID	integer,
	@newMobileFooterID	integer;
	
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
IF (@sDBVersion <> '8.2') and (@sDBVersion <> '8.3')
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


/* ------------------------------------------------------- */
PRINT 'Step - Workspace Integration'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetWorkflowIDFromName]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetWorkflowIDFromName];
	EXEC sp_executesql N'CREATE PROCEDURE spASRGetWorkflowIDFromName(
		@name varchar(255),
		@id integer OUTPUT)
	AS
	BEGIN

		IF (SELECT COUNT(id) FROM ASRSysWorkflows WHERE Name = @name) = 1
			SELECT @id = id FROM ASRSysWorkflows WHERE Name = @name;
		ELSE
			SET @id = 0;

	END'


/* ------------------------------------------------------- */
PRINT 'Step - Calculation Updates'
/* ------------------------------------------------------- */

	UPDATE tbstat_componentcode SET [precode] = 'POWER(', [aftercode] = ')', [code] = ', ' WHERE ID = 15 AND isoperator = 1;

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_IsValidNINumber]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRFn_IsValidNINumber];

	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[sp_ASRFn_IsValidNINumber]
		(
			@result integer OUTPUT,
			@input varchar(MAX)
		)
		AS
		BEGIN
			DECLARE @ValidPrefixes varchar(MAX);
			DECLARE @ValidSuffixes varchar(MAX);
			DECLARE @Prefix varchar(MAX);
			DECLARE @Suffix varchar(MAX);
			DECLARE @Numerics varchar(MAX);
			SET @result = 1;
			IF ISNULL(@input,'''') = '''' RETURN
			SET @ValidPrefixes = 
				''/AA/AB/AE/AH/AK/AL/AM/AP/AR/AS/AT/AW/AX/AY/AZ'' +
				''/BA/BB/BE/BH/BK/BL/BM/BT'' +
				''/CA/CB/CE/CH/CK/CL/CR'' +
				''/EA/EB/EE/EH/EK/EL/EM/EP/ER/ES/ET/EW/EX/EY/EZ'' +
				''/GY'' +
				''/HA/HB/HE/HH/HK/HL/HM/HP/HR/HS/HT/HW/HX/HY/HZ'' +
				''/JA/JB/JC/JE/JG/JH/JJ/JK/JL/JM/JN/JP/JR/JS/JT/JW/JX/JY/JZ'' +
				''/KA/KB/KC/KE/KH/KK/KL/KM/KP/KR/KS/KT/KW/KX/KY/KZ'' +
				''/LA/LB/LE/LH/LK/LL/LM/LP/LR/LS/LT/LW/LX/LY/LZ'' +
				''/MA/MW/MX'' +
				''/NA/NB/NE/NH/NL/NM/NP/NR/NS/NW/NX/NY/NZ'' +
				''/OA/OB/OE/OH/OK/OL/OM/OP/OR/OS/OX'' +
				''/PA/PB/PC/PE/PG/PH/PJ/PK/PL/PM/PN/PP/PR/PS/PT/PW/PX/PY'' +
				''/RA/RB/RE/RH/RK/RM/RP/RR/RS/RT/RW/RX/RY/RZ'' +
				''/SA/SB/SC/SE/SG/SH/SJ/SK/SL/SM/SN/SP/SR/SS/ST/SW/SX/SY/SZ'' +
				''/TA/TB/TE/TH/TK/TL/TM/TP/TR/TS/TT/TW/TX/TY/TZ'' +
				''/WA/WB/WE/WK/WL/WM/WP'' +
				''/YA/YB/YE/YH/YK/YL/YM/YP/YR/YS/YT/YW/YX/YY/YZ'' +
				''/ZA/ZB/ZE/ZH/ZK/ZL/ZM/ZP/ZR/ZS/ZT/ZW/ZX/ZY/'';
			SET @ValidSuffixes = ''/ /A/B/C/D/'';
			SET @Prefix = ''/''+left(@input+''  '',2)+''/''
			SET @Suffix = ''/''+substring(@input+'' '',9,1)+''/''
			SET @Numerics = SUBSTRING(@input,3,6)
			IF charindex(@Prefix,@ValidPrefixes) = 0 OR charindex(@Suffix,@ValidSuffixes) = 0 OR ISNUMERIC(@Numerics) = 0
				SET @result = 0;
		END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isnivalid]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsys_isnivalid];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_isnivalid](
			@input AS nvarchar(MAX))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit;
			
			DECLARE @ValidPrefixes varchar(MAX);
			DECLARE @ValidSuffixes varchar(MAX);
			DECLARE @Prefix varchar(MAX);
			DECLARE @Suffix varchar(MAX);
			DECLARE @Numerics varchar(MAX);

			SET @result = 1;
			IF ISNULL(@input,'''') = '''' RETURN 1

			SET @ValidPrefixes = 
				''/AA/AB/AE/AH/AK/AL/AM/AP/AR/AS/AT/AW/AX/AY/AZ'' +
				''/BA/BB/BE/BH/BK/BL/BM/BT'' +
				''/CA/CB/CE/CH/CK/CL/CR'' +
				''/EA/EB/EE/EH/EK/EL/EM/EP/ER/ES/ET/EW/EX/EY/EZ'' +
				''/GY'' +
				''/HA/HB/HE/HH/HK/HL/HM/HP/HR/HS/HT/HW/HX/HY/HZ'' +
				''/JA/JB/JC/JE/JG/JH/JJ/JK/JL/JM/JN/JP/JR/JS/JT/JW/JX/JY/JZ'' +
				''/KA/KB/KC/KE/KH/KK/KL/KM/KP/KR/KS/KT/KW/KX/KY/KZ'' +
				''/LA/LB/LE/LH/LK/LL/LM/LP/LR/LS/LT/LW/LX/LY/LZ'' +
				''/MA/MW/MX'' +
				''/NA/NB/NE/NH/NL/NM/NP/NR/NS/NW/NX/NY/NZ'' +
				''/OA/OB/OE/OH/OK/OL/OM/OP/OR/OS/OX'' +
				''/PA/PB/PC/PE/PG/PH/PJ/PK/PL/PM/PN/PP/PR/PS/PT/PW/PX/PY'' +
				''/RA/RB/RE/RH/RK/RM/RP/RR/RS/RT/RW/RX/RY/RZ'' +
				''/SA/SB/SC/SE/SG/SH/SJ/SK/SL/SM/SN/SP/SR/SS/ST/SW/SX/SY/SZ'' +
				''/TA/TB/TE/TH/TK/TL/TM/TP/TR/TS/TT/TW/TX/TY/TZ'' +
				''/WA/WB/WE/WK/WL/WM/WP'' +
				''/YA/YB/YE/YH/YK/YL/YM/YP/YR/YS/YT/YW/YX/YY/YZ'' +
				''/ZA/ZB/ZE/ZH/ZK/ZL/ZM/ZP/ZR/ZS/ZT/ZW/ZX/ZY/'';

			SET @ValidSuffixes = ''/ /A/B/C/D/'';

			SET @Prefix = ''/''+left(@input+''  '',2)+''/''
			SET @Suffix = ''/''+substring(@input+'' '',9,1)+''/''
			SET @Numerics = SUBSTRING(@input,3,6)

			IF charindex(@Prefix,@ValidPrefixes) = 0 OR charindex(@Suffix,@ValidSuffixes) = 0 OR ISNUMERIC(@Numerics) = 0
				SET @result = 0;
				
			RETURN @result;
			
		END';



/* ------------------------------------------------------- */
PRINT 'Step - Organisation Reports'
/* ------------------------------------------------------- */

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysOrganisationReport]') AND xtype in (N'U'))
	BEGIN

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReport](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[Name] [varchar](50) NOT NULL,
			[Description] [varchar](255) NOT NULL,
			[BaseViewID] [int] NOT NULL,
			[UserName] [varchar](50) NOT NULL,
			[Timestamp] [timestamp] NOT NULL)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReportAccess](
			[GroupName] varchar(256) NOT NULL,
			[Access] varchar(2) NOT NULL,
			[ID] int NOT NULL)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationColumns](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[OrganisationID] [int] NOT NULL,
			[ColumnID] [int] NOT NULL,
			[Prefix] [varchar](50) NULL,
			[Suffix] [varchar](50) NULL,
			[FontSize] int,
			[Decimals] int,
			[Height] int,
			[ConcatenateWithNext] bit)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReportFilters](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[OrganisationID] int NOT NULL,
			[FieldID] int NOT NULL,
			[Operator] [int] NOT NULL,
			[Value] nvarchar(MAX) NOT NULL)';

	END


	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysOrganisationColumns', 'U') AND name = 'ViewID')
		EXEC sp_executesql N'ALTER TABLE ASRSysOrganisationColumns ADD ViewID int NULL;';
		
	-- Insert the system permissions for 9-Box Grid Reports and new picture too
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 47)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [picture])
			VALUES (47, 'Organisation Reports', 10, 'ORGREPORTING',0x00000100010010100000010008006805000016000000280000001000000020000000010008000000000000010000000000000000000000010000000100000000000032302E00655832006E63570071665B00756B6000796D60007C7063007A7067007D7267007D7268007F776E0082776B0080776D00857C7100847D7400A2820D00D0A400008B8176008A8279008A847D008F867C00DFC76E008F8982008F8A8400938B8200938E880094908A009891880098948E009C968E0098959000A19B9300A59F9800A8A39D00ABA7A100B0A8A000B8B2AA00B6B5B100BEBBB600C2C0BB00C3C2BD00C4C3BE00C5C4BF00C7C6C100C8C7C000CAC9C400CDCAC700CDCCC600CFCEC900D0CEC900D2D1CC00D4D3CE00D2D4CE00D6D4CE00DED9C800D7D6D100D9D6D000DAD8D200DADAD400DDDAD400DFDDD700DEDCD800E0DED900E3E0DA00E2E2DC00E4E2DC00E6E4DE00EFEBDB00E9E6E000EBE8E300ECEAE400F0EDE800F2F0EA00F4F2EC00F6F4EE00F8F6F00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF0000000000000000000000000000000000002F2323232323232323232323232F000022354141414141414141414135220000213F271F1F1F1F1F1F1F1F273F210000203B1D35393B3B3D3F41451B3E2000001D391A390E123F410C0A46183B1D00001D351732FF0E4139FF0C4512391D0000192F1237111101021011440F3419000012310E41414512014649490E31120000122C0C434525FF0949494C08311200000E2A0A4546161111494C4C052F0E00000C2A05484849494C4C4C4C042C0C0000092A1D03030303030404031D2A09000009262A2A2A2A2A2A282A2A2A2609000024060606060606060606060606240000000000000000000000000000000000FFFF00008001000080010000800100008001000080010000800100008001000080010000800100008001000080010000800100008001000080010000FFFF000000);															   
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (174, 47,'New', 10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (175, 47,'Edit', 20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (176, 47,'View', 30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (177, 47,'Delete', 40, 'DELETE');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (178, 47,'Run', 40, 'RUN');


		-- Clone existing security based on system admin permissions
		DELETE FROM ASRSysGroupPermissions WHERE itemid IN (174, 175, 177, 178)
		INSERT ASRSysGroupPermissions (itemID, groupName, permitted)
			SELECT 174, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 1 AND permitted = 1
			UNION
			SELECT 175, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 1 AND permitted = 1
			UNION
			SELECT 176, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 1 AND permitted = 1
			UNION
			SELECT 177, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 1 AND permitted = 1
			UNION
			SELECT 178, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 1 AND permitted = 1

	END

	IF TYPE_ID(N'OrgChartRelation') IS NULL 
	BEGIN
		CREATE TYPE OrgChartRelation AS TABLE 
		( IsGhostNode bit
			, ManagerRoot int
			, HierarchyLevel int
			, EmployeeID int
			, Staff_Number varchar(255)
			, Reports_To_Staff_Number varchar(255));
		GRANT EXECUTE ON TYPE::OrgChartRelation TO ASRSysGroup;
	END


/* ------------------------------------------------------- */
PRINT 'Step - SQL 2016 Support'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRSQLVersion]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRSQLVersion]

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfASRSQLVersion]()
	RETURNS integer
	AS
	BEGIN
		RETURN convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY(''ProductVersion'')))
	END'



	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetActualUserDetails]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetActualUserDetails];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRGetActualUserDetails]
	(
			@psUserName sysname OUTPUT,
			@psUserGroup sysname OUTPUT,
			@piUserGroupID integer OUTPUT,
			@piModuleKey varchar(20)
	)
	AS
	BEGIN
		DECLARE @iFound		int
		DECLARE @sSQLVersion int

	   SET @sSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY(''ProductVersion'')));

		SELECT @iFound = COUNT(*) 
		FROM sysusers usu 
		LEFT OUTER JOIN	(sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
		LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
		WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
			AND (usg.issqlrole = 1 OR usg.uid IS null)
			AND lo.loginname = system_user
			AND CASE
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
			END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''

		IF (@iFound > 0)
		BEGIN
			SELECT	@psUserName = usu.name,
				@psUserGroup = CASE 
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
				END,
				@piUserGroupID = usg.gid
			FROM sysusers usu 
			LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
			LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
			WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
				AND (usg.issqlrole = 1 OR usg.uid IS null)
				AND lo.loginname = system_user
				AND CASE 
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
					END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
				AND CASE 
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
					END IN (
								SELECT [groupName]
								FROM dbo.[ASRSysGroupPermissions]
								WHERE itemID IN (
																	SELECT [itemID]
																	FROM dbo.[ASRSysPermissionItems]
																	WHERE categoryID = 1
																	AND itemKey LIKE @piModuleKey + ''%''
																)  
								AND [permitted] = 1
		)
		END
		ELSE
		BEGIN
			SELECT @psUserName = usu.name, 
				@psUserGroup = CASE
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
				END,
				@piUserGroupID = usg.gid
			FROM sysusers usu 
			LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
			LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
			WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
				AND (usg.issqlrole = 1 OR usg.uid IS null)
				AND is_member(lo.loginname) = 1
				AND CASE
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
				END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
				AND CASE 
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
					END IN (
								SELECT [groupName]
								FROM dbo.[ASRSysGroupPermissions]
								WHERE itemID IN (
																	SELECT [itemID]
																	FROM dbo.[ASRSysPermissionItems]
																	WHERE categoryID = 1
																	AND itemKey LIKE @piModuleKey + ''%''
																)  
								AND [permitted] = 1
		)
		END

		IF @psUserGroup <> ''''
		BEGIN
			DELETE FROM [ASRSysUserGroups] 
			WHERE [UserName] = SUSER_NAME()

			INSERT INTO [ASRSysUserGroups] 
			VALUES 
			(
				CASE
					WHEN @sSQLVersion <= 8 THEN USER_NAME()
					ELSE SUSER_NAME()
				END,
				@psUserGroup
			)
		END

	END';


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(MAX)	OUTPUT,
			@psMessage		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID			integer,
				@iStepID				integer,
				@iElementID				integer,
				@iRecordID				integer,
				@iRecordCount			integer,
				@sTargetName			nvarchar(MAX) = '''',
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@bUseAsTargetIdentifier bit,
				@iTemp					integer,
				@iStartElementID		integer,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms					varchar(MAX),
				@iCount					integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled				bit,
				@fHasTargetIdentifier bit,
				@iElementType			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(255),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sSPName				varchar(255),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fSaveForLater			bit,
				@fResult	bit;
		
   	   SET @iSQLVersion = dbo.udfASRSQLVersion();

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT @fExternallyInitiated = CASE
					WHEN initiationType = 2 THEN 1
					ELSE 0
				END,
				@fEnabled = [enabled],
				@fHasTargetIdentifier = [HasTargetIdentifier]
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;
		
			IF @fExternallyInitiated = 1
			BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = ''<External>'';
			END
			ELSE
			BEGIN
				SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT,
						@sTargetName OUTPUT;
				END

				IF @fHasTargetIdentifier = 1
					SET @sTargetName = ''<Unidentified>'';
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC [dbo].[spASRWorkflowUsesInitiator]
						@piWorkflowID,
						@fUsesInitiator OUTPUT;
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.'';
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.'';
						END
			
						RETURN
					END	
				END
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel'';
					END
		
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
			END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[TargetName],
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				[pageno])
			OUTPUT inserted.ID INTO @outputTable
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@sTargetName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = id FROM @outputTable;
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
		
			INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0);
		
			INSERT INTO [dbo].[ASRSysWorkflowInstanceSteps] (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 1
					ELSE 0
				END,
				0,
				0
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID;
		
			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5;
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
					
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
				OPEN immediateSubmitCursor;
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''spASRWorkflowInsertNewRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL;
		
								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''spASRWorkflowUpdateRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC [dbo].[spASRRecordDescription]
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''spASRWorkflowDeleteRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							EXEC [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT;
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(MAX), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 4,	-- 4 = failed
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
		
								SET @psMessage = @sStoredDataMsg;
								RETURN;
							END
							ELSE
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 8,	-- 8 = failed action
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE [instanceID] = @piInstanceID
									AND [elementID] = @iElementID;
		
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
									SET ASRSysWorkflowInstanceSteps.status = 1
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT,
							0;
					END
		
					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				END
				CLOSE immediateSubmitCursor;
				DEALLOCATE immediateSubmitCursor;
		
				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM [dbo].[ASRSysWorkflowInstanceSteps]
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;
			END						
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE @succeedingSteps table(stepID int)
			
			INSERT INTO @succeedingSteps 
				(stepID) VALUES (-1)
		
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM [dbo].[ASRSysWorkflowInstanceSteps]
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
			OPEN formsCursor;
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
				INSERT INTO @succeedingSteps 
				(stepID) VALUES (@iStepID)
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			END
		
			CLOSE formsCursor;
			DEALLOCATE formsCursor;
		
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET ASRSysWorkflowInstanceSteps.status = 2, 
				userName = @sActualLoginName
			WHERE ASRSysWorkflowInstanceSteps.ID IN (SELECT stepID FROM @succeedingSteps)
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRMobileInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRMobileInstantiateWorkflow];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRMobileInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@psKeyParameter	varchar(max),			
			@psPWDParameter	varchar(max),			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(MAX)	OUTPUT,
			@psMessage		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID			integer,
				@iStepID				integer,
				@iElementID				integer,
				@iRecordID				integer,
				@iRecordCount			integer,
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@bUseAsTargetIdentifier bit,
				@iTemp					integer,
				@iStartElementID		integer,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms					varchar(MAX),
				@iCount					integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled				bit,
				@iElementType			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(255),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sSPName				varchar(255),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fSaveForLater			bit,
				@fResult	bit;
			
         SELECT @iSQLVersion = dbo.udfASRSQLVersion();

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT
			-- @fExternallyInitiated = CASE
			--		WHEN initiationType = 2 THEN 1
			--		ELSE 0
			--	END,
				@fEnabled = enabled
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;

			--IF @fExternallyInitiated = 1
			--BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = @psKeyParameter;
			--END
			--ELSE
			--BEGIN
				--SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysMobileGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@psKeyParameter,			
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT;
				END
			
			print @iRecordID;
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC [dbo].[spASRWorkflowUsesInitiator]
						@piWorkflowID,
						@fUsesInitiator OUTPUT;
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.'';
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.'';
						END
			
						RETURN
					END	
				END
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel'';
					END
		
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
			--END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				pageno)
			OUTPUT inserted.ID INTO @outputTable
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = id FROM @outputTable;
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
		
			INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0);
		
			INSERT INTO [dbo].[ASRSysWorkflowInstanceSteps] (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 1
					ELSE 0
				END,
				0,
				0
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID;
		
			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5;
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
					
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
				OPEN immediateSubmitCursor;
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult	OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''spASRWorkflowInsertNewRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL;
		
								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''spASRWorkflowUpdateRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC [dbo].[spASRRecordDescription]
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''spASRWorkflowDeleteRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							EXEC [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT;
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(MAX), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 4,	-- 4 = failed
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
		
								SET @psMessage = @sStoredDataMsg;
								RETURN;
							END
							ELSE
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 8,	-- 8 = failed action
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE [instanceID] = @piInstanceID
									AND [elementID] = @iElementID;
		
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
									SET ASRSysWorkflowInstanceSteps.status = 1
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT,
							0;
					END
		
					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				END
				CLOSE immediateSubmitCursor;
				DEALLOCATE immediateSubmitCursor;
		
				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM [dbo].[ASRSysWorkflowInstanceSteps]
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;
			END						
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE @succeedingSteps table(stepID int)
			
			INSERT INTO @succeedingSteps 
				(stepID) VALUES (-1)
		
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM [dbo].[ASRSysWorkflowInstanceSteps]
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
			OPEN formsCursor;
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
				INSERT INTO @succeedingSteps 
				(stepID) VALUES (@iStepID)
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			END
		
			CLOSE formsCursor;
			DEALLOCATE formsCursor;
		
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET ASRSysWorkflowInstanceSteps.status = 2, 
				userName = @sActualLoginName
			WHERE ASRSysWorkflowInstanceSteps.ID IN (SELECT stepID FROM @succeedingSteps)
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
(
	@piInstanceID		integer,
	@piElementID		integer,
	@succeedingElements	cursor varying output,
	@psTo				varchar(MAX)
)
AS
BEGIN
	-- Action any immediate elements (Or, Decision and StoredData elements) and return the IDs of the workflow elements that 
	-- succeed them.
	-- This ignores connection elements.
	DECLARE
		@iTempID				integer,
		@iElementID				integer,
		@iElementType			integer,
		@iFlowCode				integer,
		@bUseAsTargetIdentifier	bit,
		@iTrueFlowType			integer,
		@iExprID				integer,
		@iResultType			integer,
		@sValue					varchar(MAX),
		@sResult				varchar(MAX),
		@fResult				bit,
		@dtResult				datetime,
		@fltResult				float,
		@iValue					integer,
		@iPrecedingElementType	integer, 
		@iPrecedingElementID	integer, 
		@iCount					integer,
		@iStepID				integer,
		@curRecipients			cursor,
		@sEmailAddress			varchar(MAX),
		@fDelegated				bit,
		@sDelegatedTo			varchar(MAX),
		@iSQLVersion			integer,
		@fStoredDataOK			bit, 
		@sStoredDataMsg			varchar(MAX), 
		@sStoredDataSQL			varchar(MAX), 
		@iStoredDataTableID		integer,
		@sStoredDataTableName	varchar(MAX),
		@iStoredDataAction		integer, 
		@iStoredDataRecordID	integer,
		@sStoredDataRecordDesc	varchar(MAX),
		@sStoredDataWebForms	varchar(MAX),
		@sStoredDataSaveForLater bit,
		@sSPName				varchar(MAX),
		@iNewRecordID			integer,
		@sEvalRecDesc			varchar(MAX),
		@iResult				integer,
		@iFailureFlows			integer,
		@fDeadlock				bit,
		@iErrorNumber			integer,
		@iRetryCount			integer,
		@iDEADLOCKERRORNUMBER	integer,
		@iMAXRETRIES			integer,
		@fIsDelegate			bit;

	SET @iDEADLOCKERRORNUMBER = 1205;
	SET @iMAXRETRIES = 5;
					
   SELECT @iSQLVersion = dbo.udfASRSQLVersion();
					
	DECLARE @elements table
	(
		elementID		integer,
		elementType		integer,
		processed		tinyint default 0,
		trueFlowType	integer,
		trueFlowExprID	integer
	);
					
	INSERT INTO @elements 
		(elementID,
		elementType,
		processed,
		trueFlowType,
		trueFlowExprID)
	SELECT SUCC.id,
		E.type,
		0,
		ISNULL(E.trueFlowType, 0),
		ISNULL(E.trueFlowExprID, 0)
	FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 0) SUCC
	INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID;
		
	SELECT @iCount = COUNT(*)
	FROM @elements
	WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
		AND processed = 0;

	WHILE @iCount > 0
	BEGIN
		UPDATE @elements
		SET processed = 1
		WHERE processed = 0;

		-- Action any succeeding immediate elements (Decision, Or and StoredData elements)
		DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT E.elementID,
			E.elementType,
			E.trueFlowType, 
			E.trueFlowExprID
		FROM @elements E
		WHERE (E.elementType = 4 OR (@iSQLVersion >= 9 AND E.elementType = 5) OR E.elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
			AND E.processed = 1;

		OPEN immediateCursor;
		FETCH NEXT FROM immediateCursor INTO 
			@iElementID, 
			@iElementType, 
			@iTrueFlowType, 
			@iExprID;
		WHILE (@@fetch_status = 0)
		BEGIN
			-- Submit the immediate elements, and get their succeeding elements
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 3,
				ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
				ASRSysWorkflowInstanceSteps.message = '''',
				ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

			SET @iFlowCode = 0;

			IF @iElementType = 4 -- Decision
			BEGIN
				IF @iTrueFlowType = 1
				BEGIN
					-- Decision Element flow determined by a calculation
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iExprID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0;

					SET @iValue = convert(integer, @fResult);
				END
				ELSE
				BEGIN
					-- Decision Element flow determined by a button in a preceding web form
					SET @iPrecedingElementType = 4; -- Decision element
					SET @iPrecedingElementID = @iElementID;

					WHILE (@iPrecedingElementType = 4)
					BEGIN
						SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
							@iPrecedingElementType = isnull(WE.type, 0)
						FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPrecedingElementID) PE
						INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
						INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
							AND WIS.instanceID = @piInstanceID;

						SET @iPrecedingElementID = @iTempID;
					END
					
					SELECT @sValue = ISNULL(IV.value, ''0'')
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
					WHERE IV.elementID = @iPrecedingElementID
					AND IV.instanceid = @piInstanceID
						AND E.ID = @iElementID;

					SET @iValue = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END;
				END
				
				IF @iValue IS null SET @iValue = 0;
				SET @iFlowCode = @iValue;

				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
			END
			ELSE IF @iElementType = 7 -- Or
			BEGIN
				EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @piInstanceID, @iElementID;
			END
			ELSE IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
			BEGIN
				SET @fStoredDataOK = 1;
				SET @sStoredDataMsg = '''';
				SET @sStoredDataRecordDesc = '''';

				EXEC [spASRGetStoredDataActionDetails]
					@piInstanceID,
					@iElementID,
					@sStoredDataSQL			OUTPUT, 
					@iStoredDataTableID		OUTPUT,
					@sStoredDataTableName	OUTPUT,
					@iStoredDataAction		OUTPUT, 
					@iStoredDataRecordID	OUTPUT,
					@bUseAsTargetIdentifier OUTPUT,
					@fResult OUTPUT;

				IF @fResult = 1
				BEGIN
					IF @iStoredDataAction = 0 -- Insert
					BEGIN
						SET @sSPName  = ''sp_ASRInsertNewRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@sStoredDataSQL;

								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE IF @iStoredDataAction = 1 -- Update
					BEGIN
						SET @sSPName  = ''sp_ASRUpdateRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@sStoredDataSQL,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID,
									null;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE IF @iStoredDataAction = 2 -- Delete
					BEGIN
						EXEC spASRRecordDescription
							@iStoredDataTableID,
							@iStoredDataRecordID,
							@sStoredDataRecordDesc OUTPUT;

						SET @sSPName  = ''sp_ASRDeleteRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE
					BEGIN
						SET @fStoredDataOK = 0;
						SET @sStoredDataMsg = ''Unrecognised data action.'';
					END

					IF (@fStoredDataOK = 1)
						AND ((@iStoredDataAction = 0)
							OR (@iStoredDataAction = 1))
					BEGIN

						exec [dbo].[spASRStoredDataFileActions]
							@piInstanceID,
							@iElementID,
							@iStoredDataRecordID;
					END

					IF @fStoredDataOK = 1
					BEGIN
						SET @sStoredDataMsg = ''Successfully '' +
							CASE
								WHEN @iStoredDataAction = 0 THEN ''inserted''
								WHEN @iStoredDataAction = 1 THEN ''updated''
								ELSE ''deleted''
							END + '' record'';

						IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
						BEGIN
							IF @iStoredDataRecordID > 0 
							BEGIN	
								EXEC [dbo].[spASRRecordDescription] 
									@iStoredDataTableID,
									@iStoredDataRecordID,
									@sEvalRecDesc OUTPUT
								IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
							END
						END

						IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';

						UPDATE ASRSysWorkflowInstanceValues
						SET ASRSysWorkflowInstanceValues.value = convert(varchar(255), @iStoredDataRecordID), 
							ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
						WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceValues.elementID = @iElementID
							AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
							AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 3,
							ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

						IF @bUseAsTargetIdentifier = 1
						BEGIN
							EXEC [dbo].[spASRRecordDescription] @iStoredDataTableID, @iStoredDataRecordID, @sEvalRecDesc OUTPUT;
							UPDATE ASRSysWorkflowInstances SET TargetName = @sEvalRecDesc WHERE ID = @piInstanceID;
						END

					END
					ELSE
					BEGIN
						-- Check if the failed element has an outbound flow for failures.
						SELECT @iFailureFlows = COUNT(*)
						FROM ASRSysWorkflowElements Es
						INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
							AND Ls.startOutboundFlowCode = 1
						WHERE Es.ID = @iElementID
							AND Es.type = 5; -- 5 = StoredData

						IF @iFailureFlows = 0
						BEGIN
							UPDATE ASRSysWorkflowInstanceSteps
							SET status = 4,	-- 4 = failed
								message = @sStoredDataMsg,
								failedCount = isnull(failedCount, 0) + 1,
								completionCount = isnull(completionCount, 0) - 1
							WHERE instanceID = @piInstanceID
								AND elementID = @iElementID;

							UPDATE ASRSysWorkflowInstances
							SET status = 2	-- 2 = error
							WHERE ID = @piInstanceID;
						END
						ELSE
						BEGIN
							UPDATE ASRSysWorkflowInstanceSteps
							SET status = 8,	-- 8 = failed action
								message = @sStoredDataMsg,
								failedCount = isnull(failedCount, 0) + 1,
								completionCount = isnull(completionCount, 0) - 1
							WHERE instanceID = @piInstanceID
								AND elementID = @iElementID;

							INSERT INTO @elements 
								(elementID,
								elementType,
								processed,
								trueFlowType,
								trueFlowExprID)
							SELECT SUCC.id,
								E.type,
								0,
								isnull(E.trueFlowType, 0),
								isnull(E.trueFlowExprID, 0)
							FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1) SUCC
							INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
							WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
						END
					END
				END
				ELSE				
				BEGIN
					SET @fStoredDataOK = 0;

					-- Check if the failed element has an outbound flow for failures.
					SELECT @iFailureFlows = COUNT(*)
					FROM ASRSysWorkflowElements Es
					INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
						AND Ls.startOutboundFlowCode = 1
					WHERE Es.ID = @iElementID
						AND Es.type = 5; -- 5 = StoredData

					IF @iFailureFlows = 0
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET completionCount = isnull(completionCount, 0) - 1
						WHERE instanceID = @piInstanceID
							AND elementID = @iElementID;
					END
					ELSE
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET completionCount = isnull(completionCount, 0) - 1
						WHERE instanceID = @piInstanceID
							AND elementID = @iElementID;

						INSERT INTO @elements 
							(elementID,
							elementType,
							processed,
							trueFlowType,
							trueFlowExprID)
						SELECT SUCC.id,
							E.type,
							0,
							isnull(E.trueFlowType, 0),
							isnull(E.trueFlowExprID, 0)
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1) SUCC
						INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
						WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
					END
				END;
			END

			IF (@iElementType <> 5) OR (@fStoredDataOK = 1)
			BEGIN
				-- Get this immediate element''s succeeding elements
				INSERT INTO @elements 
					(elementID,
					elementType,
					processed,
					trueFlowType,
					trueFlowExprID)
				SELECT SUCC.id,
					E.type,
					0,
					isnull(E.trueFlowType, 0),
					isnull(E.trueFlowExprID, 0)
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, @iFlowCode) SUCC
				INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
				WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
			END

			FETCH NEXT FROM immediateCursor INTO 
				@iElementID, 
				@iElementType, 
				@iTrueFlowType, 
				@iExprID;
		END
		CLOSE immediateCursor;
		DEALLOCATE immediateCursor;

		UPDATE @elements
		SET processed = 2
		WHERE processed = 1;

		SELECT @iCount = COUNT(*)
		FROM @elements
		WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
			AND processed = 0;
	END

	SELECT @iCount = COUNT(*)
	FROM @elements
	WHERE elementType = 2; -- 2=WebForm

	IF (@iCount > 0) AND len(ltrim(rtrim(@psTo))) > 0 
	BEGIN
		SELECT @iStepID = ASRSysWorkflowInstanceSteps.ID
		FROM ASRSysWorkflowInstanceSteps
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;

		DECLARE @recipients TABLE (
			emailAddress	varchar(MAX),
			delegated		bit,
			delegatedTo		varchar(MAX),
			isDelegate		bit
		);

		exec [dbo].[spASRGetWorkflowDelegates] 
			@psTo, 
			@iStepID, 
			@curRecipients output;
		FETCH NEXT FROM @curRecipients INTO 
				@sEmailAddress,
				@fDelegated,
				@sDelegatedTo,
				@fIsDelegate;
		WHILE (@@fetch_status = 0)
		BEGIN
			INSERT INTO @recipients
				(emailAddress,
				delegated,
				delegatedTo,
				isDelegate)
			VALUES (
				@sEmailAddress,
				@fDelegated,
				@sDelegatedTo,
				@fIsDelegate
			);
			
			FETCH NEXT FROM @curRecipients INTO 
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate;
		END
		CLOSE @curRecipients;
		DEALLOCATE @curRecipients;

		DELETE FROM ASRSysWorkflowStepDelegation
		WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType = 2) -- 2 = WebForm
				AND (ASRSysWorkflowInstanceSteps.status = 0
					OR ASRSysWorkflowInstanceSteps.status = 2
					OR ASRSysWorkflowInstanceSteps.status = 6
					OR ASRSysWorkflowInstanceSteps.status = 8
					OR ASRSysWorkflowInstanceSteps.status = 3));

		INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
		SELECT DISTINCT RECS.emailAddress, WIS.ID
		FROM @recipients RECS, 
			ASRSysWorkflowInstanceSteps WIS
		WHERE RECS.isDelegate = 1
			AND WIS.instanceID = @piInstanceID
				AND WIS.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType = 2) -- 2 = WebForm
				AND (WIS.status = 0
					OR WIS.status = 2
					OR WIS.status = 6
					OR WIS.status = 8
					OR WIS.status = 3);
	END

	UPDATE ASRSysWorkflowInstanceSteps
	SET ASRSysWorkflowInstanceSteps.status = 1,
		ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
		ASRSysWorkflowInstanceSteps.completionDateTime = null,
		ASRSysWorkflowInstanceSteps.userEmail = CASE
			WHEN (SELECT ASRSysWorkflowElements.type 
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @psTo -- 2 = Web Form element
			ELSE ASRSysWorkflowInstanceSteps.userEmail
		END
	WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceSteps.elementID IN 
			(SELECT E.elementID
			FROM @elements E
			WHERE E.elementType <> 7 -- 7 = Or
				AND (E.elementType <> 5 OR @iSQLVersion <= 8) -- 5 = StoredData
				AND E.elementType <> 4) -- 4 = Decision
		AND (ASRSysWorkflowInstanceSteps.status = 0
			OR ASRSysWorkflowInstanceSteps.status = 2
			OR ASRSysWorkflowInstanceSteps.status = 6
			OR ASRSysWorkflowInstanceSteps.status = 8
			OR ASRSysWorkflowInstanceSteps.status = 3);

	UPDATE ASRSysWorkflowInstanceSteps
	SET ASRSysWorkflowInstanceSteps.status = 2
	WHERE ASRSysWorkflowInstanceSteps.id IN (
		SELECT ASRSysWorkflowInstanceSteps.ID
		FROM ASRSysWorkflowInstanceSteps
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowInstanceSteps.status = 1
			AND ASRSysWorkflowElements.type = 2);

	-- Return the cursor of succeeding elements. 
	SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
		SELECT elementID 
		FROM @elements E
		WHERE E.elementType <> 7 -- 7 = Or
			AND E.elementType <> 4 -- 4 = Decision
			AND (E.elementType <> 5 OR @iSQLVersion <= 8); -- 5 = StoredData

	OPEN @succeedingElements;
END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetStoredDataActionDetails];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
	(
		@piInstanceID		integer,
		@piElementID		integer,
		@psSQL				varchar(MAX)	OUTPUT, 
		@piDataTableID		integer			OUTPUT,
		@psTableName		varchar(255)	OUTPUT,
		@piDataAction		integer			OUTPUT, 
		@piRecordID			integer			OUTPUT,
		@bUseAsTargetIdentifier	bit OUTPUT,
		@pfResult	bit OUTPUT
	)
	AS
	BEGIN
		DECLARE 
			@iPersonnelTableID			integer,
			@iInitiatorID				integer,
			@iDataRecord				integer,
			@sIDColumnName				varchar(MAX),
			@iColumnID					integer, 
			@sColumnName				varchar(MAX), 
			@iColumnDataType			integer, 
			@sColumnList				varchar(MAX),
			@sValueList					varchar(MAX),
			@sValue						varchar(MAX),
			@sRecSelWebFormIdentifier	varchar(MAX),
			@sRecSelIdentifier			varchar(MAX),
			@iTempTableID				integer,
			@iSecondaryDataRecord		integer,
			@sSecondaryRecSelWebFormIdentifier	varchar(MAX),
			@sSecondaryRecSelIdentifier	varchar(MAX),
			@sSecondaryIDColumnName		varchar(MAX),
			@iSecondaryRecordID			integer,
			@iElementType				integer,
			@iWorkflowID				integer,
			@iID						integer,
			@sWFFormIdentifier			varchar(MAX),
			@sWFValueIdentifier			varchar(MAX),
			@iDBColumnID				integer,
			@iDBRecord					integer,
			@sSQL						nvarchar(MAX),
			@sParam						nvarchar(MAX),
			@sDBColumnName				nvarchar(MAX),
			@sDBTableName				nvarchar(MAX),
			@iRecordID					integer,
			@sDBValue					varchar(MAX),
			@iDataType					integer, 
			@iValueType					integer, 
			@iSDColumnID				integer,
			@fValidRecordID				bit,
			@iBaseTableID				integer,
			@iBaseRecordID				integer,
			@iRequiredTableID			integer,
			@iRequiredRecordID			integer,
			@iDataRecordTableID			integer,
			@iSecondaryDataRecordTableID	integer,
			@iParent1TableID			integer,
			@iParent1RecordID			integer,
			@iParent2TableID			integer,
			@iParent2RecordID			integer,
			@iInitParent1TableID		integer,
			@iInitParent1RecordID		integer,
			@iInitParent2TableID		integer,
			@iInitParent2RecordID		integer,
			@iEmailID					integer,
			@iType						integer,
			@fDeletedValue				bit,
			@iTempElementID				integer,
			@iCount						integer,
			@iResultType				integer,
			@sResult					varchar(MAX),
			@fResult					bit,
			@dtResult					datetime,
			@fltResult					float,
			@iCalcID					integer,
		  @maxSize                float,
			@iSize						integer,
			@iDecimals					integer,
			@iTriggerTableID			integer;
			
		SET @psSQL = '''';
		SET @pfResult = 1;
		SET @piRecordID = 0;

		SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
		FROM ASRSysModuleSetup
		WHERE moduleKey = ''MODULE_PERSONNEL''
			AND parameterKey = ''Param_TablePersonnel'';

		IF @iPersonnelTableID = 0
		BEGIN
			SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
			AND parameterKey = ''Param_TablePersonnel'';
		END

		SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
			@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
			@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
			@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
			@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
		FROM ASRSysWorkflowInstances
		WHERE ASRSysWorkflowInstances.ID = @piInstanceID;

		SELECT @piDataAction = dataAction,
			@piDataTableID = dataTableID,
			@iDataRecord = dataRecord,
			@sRecSelWebFormIdentifier = recSelWebFormIdentifier,
			@sRecSelIdentifier = recSelIdentifier,
			@iSecondaryDataRecord = secondaryDataRecord,
			@sSecondaryRecSelWebFormIdentifier = secondaryRecSelWebFormIdentifier,
			@sSecondaryRecSelIdentifier = secondaryRecSelIdentifier,
			@iDataRecordTableID = dataRecordTable,
			@iSecondaryDataRecordTableID = secondaryDataRecordTable,
			@iWorkflowID = workflowID,
			@iTriggerTableID = ASRSysWorkflows.baseTable,
			@bUseAsTargetIdentifier = ISNULL(UseAsTargetIdentifier, 0)
		FROM ASRSysWorkflowElements
		INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID
		WHERE ASRSysWorkflowElements.ID = @piElementID;

		SELECT @psTableName = tableName
		FROM ASRSysTables
		WHERE tableID = @piDataTableID;

		IF @iDataRecord = 0 -- 0 = Initiator''s record
		BEGIN
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iPersonnelTableID,
				@iInitiatorID,
				@iInitParent1TableID,
				@iInitParent1RecordID,
				@iInitParent2TableID,
				@iInitParent2RecordID,
				@iDataRecordTableID,
				@piRecordID	OUTPUT;

			IF @piDataTableID = @iDataRecordTableID
			BEGIN
				SET @sIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
			END
		END

		IF @iDataRecord = 4 -- 4 = Triggered record
		BEGIN
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iTriggerTableID,
				@iInitiatorID,
				@iInitParent1TableID,
				@iInitParent1RecordID,
				@iInitParent2TableID,
				@iInitParent2RecordID,
				@iDataRecordTableID,
				@piRecordID	OUTPUT;

			IF @piDataTableID = @iDataRecordTableID
			BEGIN
				SET @sIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
			END
		END

		IF @iDataRecord = 1 -- 1 = Identified record
		BEGIN
			SELECT @iElementType = ASRSysWorkflowElements.type
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
				AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
			IF @iElementType = 2
			BEGIN
				 -- WebForm
				SELECT @sValue = ISNULL(IV.value, ''0''),
					@iTempTableID = EI.tableID,
					@iParent1TableID = IV.parent1TableID,
					@iParent1RecordID = IV.parent1RecordID,
					@iParent2TableID = IV.parent2TableID,
					@iParent2RecordID = IV.parent2RecordID
				FROM ASRSysWorkflowInstanceValues IV
				INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
				INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
				WHERE IV.instanceID = @piInstanceID
					AND IV.identifier = @sRecSelIdentifier
					AND Es.identifier = @sRecSelWebFormIdentifier
					AND Es.workflowID = @iWorkflowID
					AND IV.elementID = Es.ID;
			END
			ELSE
			BEGIN
				-- StoredData
				SELECT @sValue = ISNULL(IV.value, ''0''),
					@iTempTableID = Es.dataTableID,
					@iParent1TableID = IV.parent1TableID,
					@iParent1RecordID = IV.parent1RecordID,
					@iParent2TableID = IV.parent2TableID,
					@iParent2RecordID = IV.parent2RecordID
				FROM ASRSysWorkflowInstanceValues IV
				INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
					AND IV.identifier = Es.identifier
					AND Es.workflowID = @iWorkflowID
					AND Es.identifier = @sRecSelWebFormIdentifier
				WHERE IV.instanceID = @piInstanceID;
			END

			SET @piRecordID = 
				CASE
					WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
					ELSE 0
				END;
	
			SET @iBaseTableID = @iTempTableID;
			SET @iBaseRecordID = @piRecordID;
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iBaseTableID,
				@iBaseRecordID,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				@iDataRecordTableID,
				@piRecordID	OUTPUT;

			IF @piDataTableID = @iDataRecordTableID
			BEGIN
				SET @sIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
			END
		END

		SET @fValidRecordID = 1
		IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
		BEGIN
			EXEC [dbo].[spASRWorkflowValidTableRecord]
				@iDataRecordTableID,
				@piRecordID,
				@fValidRecordID	OUTPUT;

			IF @fValidRecordID = 0
			BEGIN
				-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
				EXEC [dbo].[spASRWorkflowActionFailed]
					@piInstanceID, 
					@piElementID, 
					''Stored Data primary record has been deleted or not selected.'';

				SET @psSQL = '''';
				SET @pfResult = 0;
				RETURN;
			END
		END

		IF @piDataAction = 0 -- Insert
		BEGIN
			IF @iSecondaryDataRecord = 0 -- 0 = Initiator''s record
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iPersonnelTableID,
					@iInitiatorID,
					@iInitParent1TableID,
					@iInitParent1RecordID,
					@iInitParent2TableID,
					@iInitParent2RecordID,
					@iSecondaryDataRecordTableID,
					@iSecondaryRecordID	OUTPUT;

				IF @piDataTableID = @iSecondaryDataRecordTableID
				BEGIN
					SET @sSecondaryIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
				END
			END
		
			IF @iSecondaryDataRecord = 4 -- 4 = Triggered record
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iTriggerTableID,
					@iInitiatorID,
					@iInitParent1TableID,
					@iInitParent1RecordID,
					@iInitParent2TableID,
					@iInitParent2RecordID,
					@iSecondaryDataRecordTableID,
					@iSecondaryRecordID	OUTPUT;
	
				IF @piDataTableID = @iSecondaryDataRecordTableID
				BEGIN
					SET @sSecondaryIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
				END
			END

			IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
			BEGIN
				SELECT @iElementType = ASRSysWorkflowElements.type
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)));
	
				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iTempTableID = EI.tableID,
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
					INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
					WHERE IV.instanceID = @piInstanceID
						AND IV.identifier = @sSecondaryRecSelIdentifier
						AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
						AND Es.workflowID = @iWorkflowID
						AND IV.elementID = Es.ID;
				END
				ELSE
				BEGIN
					-- StoredData
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iTempTableID = Es.dataTableID,
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
						AND IV.identifier = Es.identifier
						AND Es.workflowID = @iWorkflowID
						AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
					WHERE IV.instanceID = @piInstanceID;
				END

				SET @iSecondaryRecordID = 
					CASE
						WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
						ELSE 0
					END;
			
				SET @iBaseTableID = @iTempTableID;
				SET @iBaseRecordID = @iSecondaryRecordID;
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iBaseTableID,
					@iBaseRecordID,
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					@iSecondaryDataRecordTableID,
					@iSecondaryRecordID	OUTPUT;

				IF @piDataTableID = @iSecondaryDataRecordTableID
				BEGIN
					SET @sSecondaryIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
				END
			END

			SET @fValidRecordID = 1;
			IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
			BEGIN
				EXEC [dbo].[spASRWorkflowValidTableRecord]
					@iSecondaryDataRecordTableID,
					@iSecondaryRecordID,
					@fValidRecordID	OUTPUT;

				IF @fValidRecordID = 0
				BEGIN
					-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
					EXEC [dbo].[spASRWorkflowActionFailed] 
						@piInstanceID, 
						@piElementID, 
						''Stored Data secondary record has been deleted or not selected.'';

					SET @psSQL = '''';
					SET @pfResult = 0;
					RETURN;
				END
			END

		END

		IF @piDataAction = 0 OR @piDataAction = 1
		BEGIN
			/* INSERT or UPDATE. */
			SET @sColumnList = '''';
			SET @sValueList = '''';

			DECLARE @dbValues TABLE (
				ID integer, 
				wfFormIdentifier varchar(1000),
				wfValueIdentifier varchar(1000),
				dbColumnID int,
				dbRecord int,
				value varchar(MAX));

			INSERT INTO @dbValues (ID, 
				wfFormIdentifier,
				wfValueIdentifier,
				dbColumnID,
				dbRecord,
				value) 
			SELECT EC.ID,
				EC.wfformidentifier,
				EC.wfvalueidentifier,
				EC.dbcolumnid,
				EC.dbrecord, 
				''''
			FROM ASRSysWorkflowElementColumns EC
			WHERE EC.elementID = @piElementID
				AND EC.valueType = 2;
			
			DECLARE dbValuesCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ID,
				wfFormIdentifier,
				wfValueIdentifier,
				dbColumnID,
				dbRecord
			FROM @dbValues;
			OPEN dbValuesCursor;
			FETCH NEXT FROM dbValuesCursor INTO @iID,
				@sWFFormIdentifier,
				@sWFValueIdentifier,
				@iDBColumnID,
				@iDBRecord;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @fDeletedValue = 0;

				SELECT @sDBTableName = tbl.tableName,
					@iRequiredTableID = tbl.tableID, 
					@sDBColumnName = col.columnName,
					@iDataType = col.dataType
				FROM ASRSysColumns col
				INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
				WHERE col.columnID = @iDBColumnID;

				SET @sSQL = ''SELECT @sDBValue = ''
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
						ELSE ''convert(varchar(MAX),''
					END
					+ @sDBTableName + ''.'' + @sDBColumnName
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN '', 101)''
						ELSE '')''
					END
					+ '' FROM '' + @sDBTableName 
					+ '' WHERE '' + @sDBTableName + ''.ID = '';

				SET @iRecordID = 0;

				IF @iDBRecord = 0
				BEGIN
					-- Initiator''s record
					SET @iRecordID = @iInitiatorID;
					SET @iParent1TableID = @iInitParent1TableID;
					SET @iParent1RecordID = @iInitParent1RecordID;
					SET @iParent2TableID = @iInitParent2TableID;
					SET @iParent2RecordID = @iInitParent2RecordID;
					SET @iBaseTableID = @iPersonnelTableID;
				END			

				IF @iDBRecord = 4
				BEGIN
					-- Trigger record
					SET @iRecordID = @iInitiatorID;
					SET @iParent1TableID = @iInitParent1TableID;
					SET @iParent1RecordID = @iInitParent1RecordID;
					SET @iParent2TableID = @iInitParent2TableID;
					SET @iParent2RecordID = @iInitParent2RecordID;

					SELECT @iBaseTableID = isnull(WF.baseTable, 0)
					FROM ASRSysWorkflows WF
					INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
						AND WFI.ID = @piInstanceID;
				END
			
				IF @iDBRecord = 1
				BEGIN
					-- Identified record
					SELECT @iElementType = ASRSysWorkflowElements.type, 
						@iTempElementID = ASRSysWorkflowElements.ID
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)));

					IF @iElementType = 2
					BEGIN
						 -- WebForm
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iBaseTableID = EI.tableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sWFValueIdentifier
							AND Es.identifier = @sWFFormIdentifier
							AND Es.workflowID = @iWorkflowID
							AND IV.elementID = Es.ID;
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iBaseTableID = isnull(Es.dataTableID, 0),
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND IV.identifier = Es.identifier
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sWFFormIdentifier
						WHERE IV.instanceID = @piInstanceID;
					END

					SET @iRecordID = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END;
				END

				SET @iBaseRecordID = @iRecordID;

				SET @fValidRecordID = 1;
			
				IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
				BEGIN
					SET @fValidRecordID = 0;

					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iBaseTableID,
						@iBaseRecordID,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						@iRequiredTableID,
						@iRequiredRecordID	OUTPUT;

					SET @iRecordID = @iRequiredRecordID;

					IF @iRecordID > 0 
					BEGIN
						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iRequiredTableID,
							@iRecordID,
							@fValidRecordID	OUTPUT;
					END

					IF @fValidRecordID = 0
					BEGIN
						IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
						BEGIN
							SELECT @iCount = COUNT(*)
							FROM ASRSysWorkflowQueueColumns QC
							INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
							WHERE WFQ.instanceID = @piInstanceID
								AND QC.columnID = @iDBColumnID;

							IF @iCount = 1
							BEGIN
								SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
								FROM ASRSysWorkflowQueueColumns QC
								INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
								WHERE WFQ.instanceID = @piInstanceID
									AND QC.columnID = @iDBColumnID;

								SET @fValidRecordID = 1;
								SET @fDeletedValue = 1;
							END
						END
						ELSE
						BEGIN
							IF @iDBRecord = 1
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM ASRSysWorkflowInstanceValues IV
								WHERE IV.instanceID = @piInstanceID
									AND IV.columnID = @iDBColumnID
									AND IV.elementID = @iTempElementID;

								IF @iCount = 1
								BEGIN
									SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '''')))
									FROM ASRSysWorkflowInstanceValues IV
									WHERE IV.instanceID = @piInstanceID
										AND IV.columnID = @iDBColumnID
										AND IV.elementID = @iTempElementID;

									SET @fValidRecordID = 1;
									SET @fDeletedValue = 1;
								END
							END
						END
					END

					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed]
							@piInstanceID, 
							@piElementID, 
							''Stored Data column database value record has been deleted or not selected.'';

						SET @psSQL = '''';
						SET @pfResult = 0;
						RETURN;
					END
				END

				IF (@iDataType <> -3)
					AND (@iDataType <> -4)
				BEGIN
					IF @fDeletedValue = 0
					BEGIN
						SET @sSQL = @sSQL + convert(nvarchar(255), @iRecordID);
						SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
						EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
					END

					UPDATE @dbValues
					SET value = @sDBValue
					WHERE ID = @iID;
				END
			
				FETCH NEXT FROM dbValuesCursor INTO @iID,
					@sWFFormIdentifier,
					@sWFValueIdentifier,
					@iDBColumnID,
					@iDBRecord;
			END
			CLOSE dbValuesCursor;
			DEALLOCATE dbValuesCursor;

			DECLARE columnCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT EC.columnID,
				SC.columnName,
				SC.dataType,
				CASE
					WHEN EC.valueType = 0 THEN  -- Fixed Value
						CASE
							WHEN SC.dataType = -7 THEN
								CASE 
									WHEN UPPER(EC.value) = ''TRUE'' THEN ''1''
									ELSE ''0''
								END
							ELSE EC.value
						END
					WHEN EC.valueType = 1 THEN -- Workflow Value
						(SELECT IV.value
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
						INNER JOIN ASRSysWorkflowElements WE2 ON WE.workflowID = WE2.workflowID
						WHERE WE.identifier = EC.WFFormIdentifier
							AND WE2.id = @piElementID
							AND IV.instanceID = @piInstanceID
							AND IV.identifier = EC.WFValueIdentifier)
					ELSE '''' -- Database Value. Handle below to avoid collation conflict.
					END AS [value], 
					EC.valueType, 
					EC.ID,
					EC.calcID,
					isnull(SC.size, 0),
					isnull(SC.decimals, 0)
			FROM ASRSysWorkflowElementColumns EC
			INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
			WHERE EC.elementID = @piElementID
				AND ((SC.dataType <> -3) AND (SC.dataType <> -4));

			OPEN columnCursor;
			FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
				BEGIN
					SELECT @sValue = dbV.value
					FROM @dbValues dbV
					WHERE dbV.ID = @iSDColumnID;
				END

				IF @iValueType = 3 -- Calculated Value
				BEGIN
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iCalcID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0;

					IF @iColumnDataType = 12 SET @sResult = LEFT(@sResult, @iSize); -- Character
					IF @iColumnDataType = 2 -- Numeric
					BEGIN
						SET @maxSize = convert(float, ''1'' + REPLICATE(''0'', @iSize - @iDecimals))
						IF @fltResult >= @maxSize SET @fltResult = 0;
						IF @fltResult <= (-1 * @maxSize) SET @fltResult = 0;
					END

					SET @sValue = 
						CASE
							WHEN @iResultType = 2 THEN ltrim(rtrim(STR(@fltResult, 8000, @iDecimals)))
							WHEN @iResultType = 3 THEN 
								CASE 
									WHEN @fResult = 1 THEN ''1''
									ELSE ''0''
								END
							WHEN (@iResultType = 4) THEN
								CASE 
									WHEN @dtResult is NULL THEN ''NULL''
									ELSE convert(varchar(100), @dtResult, 101)
								END
							ELSE convert(varchar(MAX), @sResult)
						END;
				END

				IF @piDataAction = 0 
				BEGIN
					/* INSERT. */
					SET @sColumnList = @sColumnList
						+ CASE
							WHEN LEN(@sColumnList) > 0 THEN '',''
							ELSE ''''
						END
						+ @sColumnName;

					SET @sValueList = @sValueList
						+ CASE
							WHEN LEN(@sValueList) > 0 THEN '',''
							ELSE ''''
						END
						+ CASE
							WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
							WHEN @iColumnDataType = 11 THEN
								CASE 
									WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
									ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
								END
							WHEN LEN(@sValue) = 0 THEN ''0''
							ELSE isnull(@sValue, 0) -- integer, logic, numeric
						END;
				END
				ELSE
				BEGIN
					/* UPDATE. */
					SET @sColumnList = @sColumnList
						+ CASE
							WHEN LEN(@sColumnList) > 0 THEN '',''
							ELSE ''''
						END
						+ @sColumnName
						+ '' = ''
						+ CASE
							WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
							WHEN @iColumnDataType = 11 THEN
								CASE 
									WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
									ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
								END
							WHEN LEN(@sValue) = 0 THEN ''0''
							ELSE isnull(@sValue, 0) -- integer, logic, numeric
						END;
				END

				DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
					AND columnID = @iColumnID;

				INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
					(instanceID, elementID, identifier, columnID, value, emailID)
					VALUES (@piInstanceID, @piElementID, '''', @iColumnID, @sValue, 0);

				FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
			END

			CLOSE columnCursor;
			DEALLOCATE columnCursor;

			IF @piDataAction = 0 
			BEGIN
				/* INSERT. */
				IF @iDataRecord <> 3 -- 3 = Unidentified record
				BEGIN
					SET @sColumnList = @sColumnList
						+ CASE
							WHEN LEN(@sColumnList) > 0 THEN '',''
							ELSE ''''
						END
						+ @sIDColumnName;
	
					SET @sValueList = @sValueList
						+ CASE
							WHEN LEN(@sValueList) > 0 THEN '',''
							ELSE ''''
						END
						+ convert(varchar(255), @piRecordID);

					IF @piDataAction = 0 -- Insert
						AND (@iSecondaryDataRecord = 0 -- 0 = Initiator''s record
							OR @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
							OR @iSecondaryDataRecord = 4) -- 4 = Triggered record
					BEGIN
						SET @sColumnList = @sColumnList
							+ CASE
								WHEN LEN(@sColumnList) > 0 THEN '',''
								ELSE ''''
							END
							+ @sSecondaryIDColumnName;
				
						SET @sValueList = @sValueList
							+ CASE
								WHEN LEN(@sValueList) > 0 THEN '',''
								ELSE ''''
							END
							+ convert(varchar(255), @iSecondaryRecordID);
					END
				END
			END

			IF LEN(@sColumnList) > 0
			BEGIN
				IF @piDataAction = 0 
				BEGIN
					/* INSERT. */
					SET @psSQL = ''INSERT INTO '' + @psTableName
						+ '' ('' + @sColumnList + '')''
						+ '' VALUES('' + @sValueList + '')'';
					SET @pfResult = 1;
				END
				ELSE
				BEGIN
					/* UPDATE. */
					SET @psSQL = ''UPDATE '' + @psTableName
						+ '' SET '' + @sColumnList
						+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
					SET @pfResult = 1;
				END
			END
		END

		IF @piDataAction = 2
		BEGIN
			/* DELETE. */
			SET @psSQL = ''DELETE FROM '' + @psTableName
				+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
			SET @pfResult = 1;
		END	

		IF (@piDataAction = 0) -- Insert
		BEGIN
			SET @iParent1TableID = isnull(@iDataRecordTableID, 0);
			SET @iParent1RecordID = isnull(@piRecordID, 0);
			SET @iParent2TableID = isnull(@iSecondaryDataRecordTableID, 0);
			SET @iParent2RecordID = isnull(@iSecondaryRecordID, 0);
		END
		ELSE
		BEGIN	-- Update or Delete
			exec [dbo].[spASRGetParentDetails]
				@piDataTableID,
				@piRecordID,
				@iParent1TableID	OUTPUT,
				@iParent1RecordID	OUTPUT,
				@iParent2TableID	OUTPUT,
				@iParent2RecordID	OUTPUT;
		END

		UPDATE ASRSysWorkflowInstanceValues
		SET ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID, 
			ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
			ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID, 
			ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
		WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceValues.elementID = @piElementID
			AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
			AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;

		IF (@piDataAction = 2) -- Delete
		BEGIN
			DECLARE curColumns CURSOR LOCAL FAST_FORWARD FOR 
			SELECT columnID
			FROM [dbo].[udfASRWorkflowColumnsUsed] (@iWorkflowID, @piElementID, 0);

			OPEN curColumns;

			FETCH NEXT FROM curColumns INTO @iDBColumnID;
			WHILE (@@fetch_status = 0)
			BEGIN
				DELETE FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
					AND columnID = @iDBColumnID;

				SELECT @sDBTableName = tbl.tableName,
					@iRequiredTableID = tbl.tableID, 
					@sDBColumnName = col.columnName,
					@iDataType = col.dataType
				FROM ASRSysColumns col
				INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
				WHERE col.columnID = @iDBColumnID;

				SET @sSQL = ''SELECT @sDBValue = ''
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
						ELSE ''convert(varchar(MAX),''
					END
					+ @sDBTableName + ''.'' + @sDBColumnName
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN '', 101)''
						ELSE '')''
					END
					+ '' FROM '' + @sDBTableName 
					+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);

				SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
				EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;

				INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
					(instanceID, elementID, identifier, columnID, value, emailID)
					VALUES (@piInstanceID, @piElementID, '''', @iDBColumnID, @sDBValue, 0);
					
				FETCH NEXT FROM curColumns INTO @iDBColumnID;
			END
			CLOSE curColumns;
			DEALLOCATE curColumns;

			DECLARE curEmails CURSOR LOCAL FAST_FORWARD FOR 
			SELECT emailID,
				type,
				colExprID
			FROM [dbo].[udfASRWorkflowEmailsUsed] (@iWorkflowID, @piElementID, 0);

			OPEN curEmails;

			FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
			WHILE (@@fetch_status = 0)
			BEGIN
				DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
					AND emailID = @iEmailID;

				IF @iType = 1 -- Column
				BEGIN
					SELECT @sDBTableName = tbl.tableName,
						@iRequiredTableID = tbl.tableID, 
						@sDBColumnName = col.columnName,
						@iDataType = col.dataType
					FROM [dbo].[ASRSysColumns] col
					INNER JOIN [dbo].[ASRSysTables] tbl ON col.tableID = tbl.tableID
					WHERE col.columnID = @iDBColumnID;

					SET @sSQL = ''SELECT @sDBValue = ''
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
							ELSE ''convert(varchar(MAX),''
						END
						+ @sDBTableName + ''.'' + @sDBColumnName
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN '', 101)''
							ELSE '')''
						END
						+ '' FROM '' + @sDBTableName 
						+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);

					SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
					EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
				END
				ELSE
				BEGIN
					EXEC [dbo].[spASRSysEmailAddr]
						@sDBValue OUTPUT,
						@iEmailID,
						@piRecordID;
				END

				INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
					(instanceID, elementID, identifier, columnID, value, emailID)
					VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID);
					
				FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
			END
			CLOSE curEmails;
			DEALLOCATE curEmails;
		END
	END'


/* ------------------------------------------------------- */
PRINT 'Step - SQL Metadata Stored Proc'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetSQLMetadata]') AND xtype = 'P')
		DROP PROCEDURE [dbo].spASRGetSQLMetadata;
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRGetSQLMetadata](
	@sServerName nvarchar(128) OUTPUT,
	@sDBName nvarchar(128) OUTPUT)
	AS
	BEGIN
			SET @sServerName = CONVERT(nvarchar(128), SERVERPROPERTY(''ServerName''));
			SET @sDBName = db_name();
	END'


/* ------------------------------------------------------- */
PRINT 'Step - General Updates'
/* ------------------------------------------------------- */


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRPostSystemSave]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRPostSystemSave];

	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRPostSystemSave]
		AS
		BEGIN

			SET NOCOUNT ON;

			DECLARE @iBlockPatch int = 0;

			IF OBJECT_ID(''ASRSysProtectsCache'') IS NOT NULL 
				DELETE FROM ASRSysProtectsCache;

			INSERT ASRSysProtectsCache ([ID], [Action], [Columns], [ProtectType], [UID])
				SELECT p.ID, Action, Columns, ProtectType , p.uid
					FROM sys.sysprotects p
					INNER JOIN sys.sysobjects o ON o.id = p.id
					WHERE o.xtype = ''V'' AND p.uid < @iBlockPatch
					ORDER BY p.uid, name;

			INSERT ASRSysProtectsCache ([ID], [Action], [Columns], [ProtectType], [UID])
				SELECT p.ID, Action, Columns, ProtectType , p.uid
					FROM sys.sysprotects p
					INNER JOIN sys.sysobjects o ON o.id = p.id
					WHERE o.xtype = ''V'' AND p.uid >= @iBlockPatch
					ORDER BY p.uid, name;

		END';


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRMakeLoginsProcessAdmin]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRMakeLoginsProcessAdmin];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRIsServer64Bit]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfASRIsServer64Bit]

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfASRIsServer64Bit]()
		RETURNS int
		AS
		BEGIN

			DECLARE @bIs64Bit bit
			SELECT @bIs64Bit = CASE PATINDEX (''%X64)%'' , @@version)
					WHEN 0 THEN 0
					ELSE 1
				END
			RETURN @bIs64Bit

		END'


/* ------------------------------------------------------- */
PRINT 'Step - Overnight Metrics'
/* ------------------------------------------------------- */

	-- Create the progress table if it doesn't already exist
	IF OBJECT_ID('ASRSysOvernightProgress', N'U') IS NULL
		EXEC sp_executesql N'CREATE TABLE ASRSysOvernightProgress
			(TableName varchar(255)
			, RecCount int
			, IDRange varchar(255)
			, StartDate datetime
			, EndDate datetime
			, DurationMins int)';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysOvernightProgress', 'U') AND name = 'DurationSecs')
		EXEC sp_executesql N'ALTER TABLE ASRSysOvernightProgress ADD DurationSecs int';


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSysOvernightTableUpdate]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRSysOvernightTableUpdate];

	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
	(
		@psTableName varchar(255),
		@psFieldName varchar(255),
		@piBatches int
	) 
	AS
	BEGIN
		SET NOCOUNT ON;

		DECLARE @lowid		integer, 
				@maxid		integer,
				@rowcount	integer,
				@start		datetime;

		DECLARE @sSQL				nvarchar(MAX),
				@sParamDefinition	nvarchar(500);

		-- Determine the number of ID''s we''ll update in each batch
		IF ISNULL(@piBatches, 0) = 0
			SET @piBatches = 2000;
	
		SET @sSQL = ''SELECT @lowid = ISNULL(MIN(ID),0),  @maxid = ISNULL(MAX(ID),0) FROM '' + @psTableName;
		SET @sParamDefinition = N''@lowid int OUTPUT, @maxid int OUTPUT'';
		EXEC sp_executesql @sSQL, @sParamDefinition, @lowid OUTPUT, @maxid OUTPUT;

		WHILE 1=1
		BEGIN
			SET @start = GETDATE();
		
			-- Do the update
			SELECT @sSQL = ''UPDATE '' + @psTableName + '' SET '' + @psFieldName + '' = '' + @psFieldName
						+ '' WHERE ID BETWEEN '' + CONVERT(nvarchar(10), @lowid) + '' AND '' + CONVERT(varchar(10),  @lowid + @piBatches - 1);
			EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @piBatches;

			SET @rowcount = @@ROWCOUNT;

			-- insert a record to this progress table to check the progress
			INSERT INTO ASRSysOvernightProgress (TableName, RecCount, IDRange, StartDate, EndDate, DurationSecs)
				SELECT @psTableName
					, @rowcount
					, CAST(@lowid as varchar(255)) + ''-'' + CAST(@lowid + @piBatches - 1 as varchar(255))
					, @start
					, GETDATE()
					, DATEDIFF(ss, @start, GETDATE());

			SET @lowid = @lowid + @piBatches;

			IF @lowid > @maxid
			BEGIN
				CHECKPOINT;
				BREAK;
			END
			ELSE
				CHECKPOINT;
		END

		SET NOCOUNT OFF;
	END'


/* ------------------------------------------------------- */
PRINT 'Step - Performance Improvements'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[tbsys_intransactiontrigger]') AND xtype = 'U')
		DROP TABLE [dbo].[tbsys_intransactiontrigger];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[InTriggerContext]') AND xtype = 'V')
		DROP VIEW [dbo].[InTriggerContext];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spsys_TrackTriggerInsert]') AND xtype = 'P')
		DROP PROCEDURE [dbo].spsys_TrackTriggerInsert;

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spsys_TrackTriggerClear]') AND xtype = 'P')
		DROP PROCEDURE [dbo].spsys_TrackTriggerClear;

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysGetContextTable]') AND xtype = 'TF')
		DROP FUNCTION [dbo].udfsysGetContextTable;


	EXEC sp_executesql N'CREATE FUNCTION [dbo].udfsysGetContextTable()
	  RETURNS @Context TABLE([TableFromId] integer, [NestLevel] tinyint, [ActionType] tinyint)
	  WITH SCHEMABINDING
	AS
	BEGIN

	  DECLARE @buffer varchar(128) = rtrim(replace(convert(varchar(128),CONTEXT_INFO()), char(0), char(32)));
	  DECLARE @fPtr1 int = CHARINDEX(CHAR(2),@buffer),
			  @rPtr int = CHARINDEX(CHAR(3),@buffer);
	  DECLARE @fPtr2 int = CHARINDEX(CHAR(2),@buffer, @fPtr1+1);
		  
	  WHILE @rPtr > 0
	  BEGIN

		INSERT INTO @Context
			SELECT convert(integer, SUBSTRING(@buffer,1,abs(@fPtr1-1))),
				convert(tinyint, SUBSTRING(@buffer, @fPtr1+1, @fPtr2-@fPtr1-1)), 
				convert(tinyint, SUBSTRING(@buffer, @fPtr2+1, @rPtr-@fPtr2-1))
			WHERE @rPtr > NULLIF(@fPtr1,0)+1;

		SET @buffer = SUBSTRING(@buffer,@rPtr+1,128);
		SET @fPtr1 = CHARINDEX(CHAR(2),@buffer);
		SET @fPtr2 = CHARINDEX(CHAR(2),@buffer, @fPtr1+1);
		SET @rPtr = CHARINDEX(CHAR(3),@buffer);

	  END

	  RETURN;

	END';


	EXEC sp_executesql N'CREATE VIEW [dbo].InTriggerContext
	  WITH SCHEMABINDING
	AS
	SELECT TOP 16 [TableFromId], [NestLevel], [ActionType]
	   FROM dbo.udfsysGetContextTable()';


	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spsys_TrackTriggerInsert](@TableFromID integer, @actionType tinyint, @NestLevel tinyint)
	AS
	BEGIN

	   BEGIN TRY

		IF ISNULL(len(@TableFromID),0) = 0
		   RAISERROR(''Context Key may not by null or empty.'',11,1);

		DECLARE @buffer varchar(128) = '''';

		SELECT @buffer += convert(varchar(125),[TableFromId]) + CHAR(2) + convert(varchar(3),[NestLevel]) + CHAR(2) + convert(varchar(3),[ActionType]) + CHAR(3)
		  FROM [InTriggerContext]
		  WHERE [TableFromId] != @TableFromID;

		IF LEN(@buffer) + LEN(@TableFromID) + LEN(@NestLevel)  > 126
		   RAISERROR(''Context buffer overflow.'',11,1);

		IF ISNULL(len(@NestLevel),0) > 0
		   SELECT @buffer += convert(varchar(125), @TableFromID) + CHAR(2) + convert(varchar(3),@NestLevel) + CHAR(2) + convert(varchar(3), @actionType) + CHAR(3)

		DECLARE @varbin varbinary(128) = convert(varbinary(128),@buffer);
		SET CONTEXT_INFO @varbin;

	  END TRY
	  BEGIN CATCH
		DECLARE @ErrMsg nvarchar(4000)=isnull(ERROR_MESSAGE(),''Error caught in setContextValue''), @ErrSeverity int=ERROR_SEVERITY();
	  END CATCH

	  FINALLY:

	  if @ErrSeverity > 0  RAISERROR(@ErrMsg, @ErrSeverity, 1);

	  RETURN isnull(len(@buffer),0);

	END';

	EXEC sp_executesql N'CREATE PROCEDURE dbo.spsys_TrackTriggerClear(@TableFromID integer)
	AS
	BEGIN

		DECLARE @buffer varchar(128) = '''',
				  @varBin varbinary(128);

		SELECT @buffer += convert(varchar(125),[TableFromId]) + CHAR(2) + convert(varchar(3),[NestLevel]) + CHAR(2) + convert(varchar(3),[ActionType]) + CHAR(3)
			  FROM [InTriggerContext]
			  WHERE [TableFromId] <> @TableFromID

		SET @varBin = convert(varbinary(128), @buffer);
	   SET CONTEXT_INFO @varBin;

	END';


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSysOvernightTableUpdate]') AND xtype = 'P')
		DROP PROCEDURE [dbo].spASRSysOvernightTableUpdate;

	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
	(
	   @psTableName varchar(255),
	   @psFieldName varchar(255),
	   @piBatches int = 1000,
	   @psWhereClause varchar(MAX) = ''''
	) 
	AS
	BEGIN
		SET NOCOUNT ON;

		DECLARE @lowid		integer, 
			 @maxid		integer,
			 @rowcount	integer,
			 @start		datetime;

		DECLARE @sSQL				nvarchar(MAX),
				@sParamDefinition	nvarchar(500),
			 @disableIndexSQL nvarchar(MAX) = '''';
	
		SET @sSQL = ''SELECT @lowid = ISNULL(MIN(ID),0),  @maxid = ISNULL(MAX(ID),0) FROM '' + @psTableName;
		SET @sParamDefinition = N''@lowid int OUTPUT, @maxid int OUTPUT'';
		EXEC sp_executesql @sSQL, @sParamDefinition, @lowid OUTPUT, @maxid OUTPUT;

  		-- Disable table scalar table indexes
		SELECT @disableIndexSQL = @disableIndexSQL + ''ALTER INDEX ['' + i.name + ''] ON '' + t.name + '' DISABLE;'' + CHAR(13)
			FROM sys.indexes i
			INNER JOIN sys.tables t ON i.object_id = T.object_id
			WHERE i.type_desc = ''NONCLUSTERED''
				AND i.name IS NOT NULL AND i.name LIKE ''IDX_udftab%'' AND OBJECT_NAME(i.object_id) = @pstableName
		EXECUTE sp_executeSQL @disableIndexSQL;

		WHILE 1=1
		BEGIN
			SET @start = GETDATE();
		
			-- Do the update
			SELECT @sSQL = ''UPDATE '' + @psTableName + '' SET '' + @psFieldName + '' = '' + @psFieldName
						+ '' WHERE ID BETWEEN '' + CONVERT(nvarchar(10), @lowid) + '' AND '' + CONVERT(varchar(10),  @lowid + @piBatches - 1)
				   + CASE WHEN LEN(@psWhereClause) > 0 THEN '' AND '' + @psWhereClause ELSE '''' END
			EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @piBatches;

			SET @rowcount = @@ROWCOUNT;

			-- insert a record to this progress table to check the progress
			INSERT INTO ASRSysOvernightProgress (TableName, RecCount, IDRange, StartDate, EndDate, DurationSecs)
				SELECT @psTableName
					, @rowcount
					, CAST(@lowid as varchar(255)) + ''-'' + CAST(@lowid + @piBatches - 1 as varchar(255))
					, @start
					, GETDATE()
					, DATEDIFF(ss, @start, GETDATE());

			SET @lowid = @lowid + @piBatches;

			IF @lowid > @maxid
			BEGIN
				CHECKPOINT;
				BREAK;
			END
			ELSE
				CHECKPOINT;
		END

  		-- Rebuild table scalar table indexes
	   SET @disableIndexSQL = '''';
		SELECT @disableIndexSQL = @disableIndexSQL + ''ALTER INDEX ['' + i.name + ''] ON '' + t.name + '' REBUILD;'' + CHAR(13)
			FROM sys.indexes i
			INNER JOIN sys.tables t ON i.object_id = T.object_id
			WHERE i.type_desc = ''NONCLUSTERED''
				AND i.name IS NOT NULL AND i.name LIKE ''IDX_udftab%'' AND OBJECT_NAME(i.object_id) = @pstableName
		EXECUTE sp_executeSQL @disableIndexSQL;

		SET NOCOUNT OFF;
	END'

	-- Default the overnight stop process to the personnel leaving date
	DECLARE @overnightColumn int,
			@batchsize int,
			@ignoreArchive bit;

	SELECT @overnightColumn = SettingValue
		FROM ASRSysSystemSettings
		WHERE Section = 'overnight' AND SettingKey = 'archivecolumn';

	IF @overnightColumn IS NULL
	BEGIN
		SELECT @overnightColumn = ParameterValue
			FROM ASRSysModuleSetup
			WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey =   'Param_FieldsLeavingDate' AND ParameterType = 'PType_ColumnID';

		EXEC spsys_setsystemsetting 'overnight', 'archivecolumn', @overnightColumn;

	END

	SELECT @batchsize = SettingValue
		FROM ASRSysSystemSettings
		WHERE Section = 'overnight' AND SettingKey = 'batchsize';

	IF @batchsize IS NULL
		EXEC spsys_setsystemsetting 'overnight', 'batchsize', 1000;


	SELECT @ignoreArchive = SettingValue
		FROM ASRSysSystemSettings
		WHERE Section = 'overnight' AND SettingKey = 'ignorearchived';

	IF @ignoreArchive IS NULL
		EXEC spsys_setsystemsetting 'overnight', 'ignorearchived', 0;


/* ------------------------------------------------------- */
PRINT 'Step - Data Protection Enhancements'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_IsPersonnelSubordinateOfUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_ByID_HasPersonnelSubordinateUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_ByID_HasPostSubordinateUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_ByID_HasPostSubordinateUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_ByID_IsPersonnelSubordinateOfUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_ByID_IsPersonnelSubordinateOfUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_ByID_IsPostSubordinateOfUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_ByID_IsPostSubordinateOfUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_HasPersonnelSubordinateUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_HasPersonnelSubordinateUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_HasPostSubordinateUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_HasPostSubordinateUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_IsPersonnelSubordinateOfUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_IsPersonnelSubordinateOfUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRFn_IsPostSubordinateOfUser]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udf_ASRFn_IsPostSubordinateOfUser]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getfieldfromdatabaserecord]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord];

	-- Remove deleted flag direct table access
	IF EXISTS(SELECT * FROM sys.syscolumns c
		INNER JOIN ASRSysTables t ON OBJECT_NAME(c.id) LIKE 'tbuser_' + TableName
		WHERE c.name = '_deleted')
	BEGIN

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand + 'IF EXISTS(SELECT * FROM dbo.sysobjects WHERE name = ''trsys_' + TableName + '_d01'' AND xtype = ''TR'')
				DROP TRIGGER [dbo].[trsys_' + TableName + '_d01];' + CHAR(13)
			FROM ASRSysTables;
		EXECUTE sp_executeSQL @NVarCommand;

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand + 'DELETE FROM dbo.tbuser_' + TableName + ' WHERE _deleted = 1;' + CHAR(13)
			FROM ASRSysTables;
		EXECUTE sp_executeSQL @NVarCommand;
	END


/* ------------------------------------------------------- */
PRINT 'Step - Unique Code Enhancements'
/* ------------------------------------------------------- */

	GRANT CREATE SEQUENCE ON SCHEMA::dbo TO [ASRSysGroup] 

	SET @NVarCommand = '';
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[tbsys_uniquecodes]') AND xtype = 'U')
	BEGIN
		SELECT @NVarCommand = @NVarCommand + 'IF NOT EXISTS (SELECT * FROM sys.sequences WHERE name = N''sequence_' + CodePrefix + ''')
			CREATE SEQUENCE [dbo].[sequence_' + CodePrefix + '] START WITH ' + convert(nvarchar(MAX), MaxCodeSuffix + 1) + ';' + CHAR(13)
			FROM tbsys_uniquecodes
			WHERE ISNULL(CodePrefix, '') <> '';
		EXECUTE sp_executeSQL @NVarCommand;

		EXECUTE sp_executeSQL N'DROP TABLE dbo.tbsys_uniquecodes';
	END

	SET @NVarCommand = '';	
	SELECT @NVarCommand = @NVarCommand + 'GRANT UPDATE ON dbo.[' + name + '] TO ASRSysGroup;' FROM sys.sequences;
	EXECUTE sp_executesql @NVarCommand;

	UPDATE ASRSysFunctions SET spName = 'sp_ASRFn_GetUniqueCode @piInstanceID,' WHERE functionID = 43

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_GetUniqueCode]') AND xtype = 'p')
		DROP PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode]
		(
			@piInstanceID	int,
			@piResult		int OUTPUT,
			@psCodePrefix	varchar(MAX) = '''',
			@piSuffixRoot	int=1
		)
		AS
		BEGIN
			SELECT @piResult = [dbo].[udfstat_getuniquecode] (@psCodePrefix, @piSuffixRoot, @piInstanceID);
		END';


/* ------------------------------------------------------- */
PRINT 'Step - Document Management Enhancements'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysDocumentManagementTypes', 'U') AND name = 'TargetTitleColumnID')
		EXEC sp_executesql N'ALTER TABLE ASRSysDocumentManagementTypes ADD TargetTitleColumnID int NULL;';


/* ------------------------------------------------------------- */
PRINT 'Step - Image rebranding updates'
/* ------------------------------------------------------------- */

	-- Create system tracking column
	IF NOT EXISTS(SELECT ID FROM syscolumns	WHERE ID = (SELECT ID FROM sysobjects where [name] = 'ASRSysPictures') AND [name] = 'GUID')
		EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysPictures] ADD [GUID] [uniqueidentifier] NULL;';

	-- Generic image update routine
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_writepicture]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_writepicture];
	EXECUTE sp_executeSQL  N'CREATE PROCEDURE spadmin_writepicture(@guid uniqueidentifier, @name varchar(255), @picturetype integer, @pictureID integer OUTPUT, @pictureHex varbinary(MAX))
	AS
	BEGIN

		IF NOT EXISTS(SELECT [guid] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid)	
		BEGIN

			SELECT @pictureID = ISNULL(MAX(PictureID), 0) + 1 FROM dbo.[ASRSysPictures];

			INSERT [ASRSysPictures] (PictureID, Name, PictureType, [guid], [Picture]) 
				SELECT @pictureID, @name, @picturetype, @guid, @pictureHex;

		END
		ELSE
		BEGIN
			SELECT @pictureID = [PictureID] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid;
			UPDATE [ASRSysPictures] SET [Name] = @name, Picture = @pictureHex WHERE [guid] = @guid;
		END

	END';

	-- Add/update images
	EXEC dbo.spadmin_writepicture '7410CCC5-01EF-46F0-9D9F-9323A93B4573', 'Advanced Background.jpg', 1, @newDesktopImageID OUTPUT, 0xFFD8FFE000104A46494600010101006000600000FFE100584578696600004D4D002A00000008000401310002000000110000003E511000010000000101000000511100040000000100002E23511200040000000100002E230000000041646F626520496D61676552656164790000FFDB0043000201010201010202020202020202030503030303030604040305070607070706070708090B0908080A0807070A0D0A0A0B0C0C0C0C07090E0F0D0C0E0B0C0C0CFFDB004301020202030303060303060C0807080C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0CFFC000110800F5049D03012200021101031101FFC4001F0000010501010101010100000000000000000102030405060708090A0BFFC400B5100002010303020403050504040000017D01020300041105122131410613516107227114328191A1082342B1C11552D1F02433627282090A161718191A25262728292A3435363738393A434445464748494A535455565758595A636465666768696A737475767778797A838485868788898A92939495969798999AA2A3A4A5A6A7A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6D7D8D9DAE1E2E3E4E5E6E7E8E9EAF1F2F3F4F5F6F7F8F9FAFFC4001F0100030101010101010101010000000000000102030405060708090A0BFFC400B51100020102040403040705040400010277000102031104052131061241510761711322328108144291A1B1C109233352F0156272D10A162434E125F11718191A262728292A35363738393A434445464748494A535455565758595A636465666768696A737475767778797A82838485868788898A92939495969798999AA2A3A4A5A6A7A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6D7D8D9DAE2E3E4E5E6E7E8E9EAF2F3F4F5F6F7F8F9FAFFDA000C03010002110311003F00FDFCA28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800AF03F8C9F1CEE2D7E285A2E9926EB5F0FCA43007E5B893A480FB632BFF007D1EF5E8DF1E7E252FC38F0448F0C9B751BFCC16A075538F99FF00E020FE656BE5369F73649249E493DEBF8D7E941E2B57CBA54386B27AAE15938D5AB28BD63CAD4A9C7D5B4A6FC943A499FA3703E411ACA58DC446F1D6314FADF46FF4FBFB1F6AF8735FB7F146856BA85A36FB7BB8C4887B8CF507DC1C83EE2AED7817ECA3F138596A32786EEA4C43744CB6658FDD93F893FE04391EE0FAD7BED7F43785BC79438BB8768E6F4ECAA7C3522BECD48DB997A3D251FEEC9753E473DCA679763258796DBC5F74F6FF0027E6828A28AFD10F1C28A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0029B34CB6F0B49232A47182CCCC70140EA4D3ABC83F6B3F8A7FF08C786D341B49317BAB2E672A798E0E87FEFB3C7D0357CA71BF1761386724AF9D633E1A71D175949E918AF393B2F2577B267A394E5B531F8B861696F27BF65D5FC91E41F1AFE2637C4AF1CDC5D46CDF61B7FDC5A29ED183F7BEAC727F103B5723E7557F3E8F3EBFC88CFB38C5E7198D6CD31F2E6AB564E527E6FA2EC96C97449247F47E13054F0D4634292B462ACBFAFCCBD63AA4DA65EC3716F23433DBB892375EA8C0E411F435F617C29F8830FC4BF055AEA51ED5988F2EE631FF002CA518DC3E87823D88AF8BFCFAF48FD9A3E2A7FC207E375B3BA936E9BAB9586524FCB149FC0FF99C1F639ED5FB67D1DFC46FF56788560B172B61715684AFB467F627E566F964F6E5777F0A3E5F8D321FAF60BDAD25FBCA7AAF35D57EABCD799F56514515FE9B1F8385145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140147C4DE22B5F0968179A95EC9E5DAD9466590F7C0EC3DC9C003B922BE25F1D78D2EBC7DE2CBED5AF3FD75E49B82E788D470AA3D80007E15EB1FB65FC5B179A8C5E15B397F756A44F7C54FDE931944FF808393EE47715E0FF0068FF006ABFCF3FA4D7888F39CE170F60A57A1856F9ADB4AAECFF00F00578AFEF39F4B1FB8787F907D5B0BF5EAABDFA9B7947A7DFBFA58B5E651E6555FB47FB547DA3FDAAFE5DF66CFD0B90B5E651E6555FB47FB547DA3FDAA3D9B0E43EC0FD99BE2A7FC2C6F022DBDD49BB54D27104F9EB227F049F88183EEA4F715E915F10FC1AF8A127C2DF1ED9EA6199AD49F26EE35FF9690B11BBF1180C3DD457DB3637B0EA5650DC5BC8B341708B246EA72AEA46411EC41AFF004D3E8FBE233E26E1E585C5CAF8AC2DA13BEF28DBDC9F9DD2E593DF9A2DBDD1F80F1B643FD9D8DF694D7EEEA6ABC9F55FAAF27E44B451457EF47C585145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400572FF18BE2541F0A7C0379AB4BB1A651E55AC4C7FD74CDF747D0724FB29AEA2BE38FDACFE317FC2C5F880D61672EED27442D04454E56697FE5A49EE32368F65C8FBD5F9478C9E204784F876A62A93FF68ABEE525FDE6B595BB417BDDAFCA9EE7D4708E42F34C7C69C97EEE3ACBD3B7CDE9E977D0F3AD4F579B59D46E2EEEA469AE2EA46965918F2ECC7249FC6A1F37FCE6AAF9D479D5FE5C54E69C9CE6DB6F56DEEDBEACFE948D351565B16BCDFF0039A3CDFF0039AABE751E7547207216BCDFF39A3CDFF39AABE751E751C81C85AF37FCE6BE9AFD8BBE2EFF006CE8D2F856F24FF49D3D4CD6449FBF0E7E64FAA9391EC7D16BE5BF3AB4BC1DE31BCF03F89EC756B17D9756328913D1BD54FB11907D8D7E85E17F1BD6E13E20A39A42EE9FC3522BED5395B9BE6B4947FBC974B9E1F126470CCF033C33F8B78BED25B7DFB3F267E8551593E04F19D9FC41F08D86B162D9B7BE8848013931B74643EEAC083EE2B5ABFD57C262E8E2A843138792942694A2D6CD3574D7935A9FCC75A94E94DD3A8AD24ECD766B70A28A2BA0CC28A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A2A2BEBE874CB19AE6E24586DEDD1A5964638545519249F40066A652514E5276486936EC8F32FDAC3E327FC2ABF87525BDACBB758D6835BDB6D3F34498F9E5FC01C0FF006981EC6BE29F3ABA6F8F3F16E6F8C5F11EF35466916CD0F9165137FCB38549DBC7AB72C7DDBDAB8DF33DEBFCC7F19B8FA5C57C433AF45DF0F46F0A4BA349EB3F59BD7BF2F2A7B1FD37C1BC3AB2BCBE309AFDE4FDE97AF45F25A7ADD973CEA3CEAA7E67BD1E67BD7E49C87D672173CEA3CEAA7E67BD1E67BD1C81C85CF3A8F3AA9F99EF4799EF472072173CEA3CEAA7E67BD1E67BD1C81C87D09FB117C65FF8477C4F2785EFA6DB67ABB799685BA47718C6DFF81818FAAAFA9AFAC2BF33ED2FA4B1BA8E786578A685C3C6EA48646072083EA0D7DF1FB3E7C5C8BE32FC36B4D4B728D421FF0047BE8C63E59940C9C760DC30FAE3B57F71FD197C41FAD60E5C2D8D97BF453952BF585FDE8FAC1BBAFEEBB2D227E1FE2670EFB1AAB34A2BDD9E92F29747F35BF9AEECEE28A28AFEB23F270A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800AF9DFF6F1F8D7FD81A043E10B097FD33535135F953CC5003F2A7D5D867FDD5F46AF6FF883E39B1F86DE0CD435CD45B6DAE9F1191803F3487A2A0F766200F735F9C9E3CF1D5F7C44F186A1AD6A126FBAD42632BF3C20E8AA3FD9550147B015FCE7F48AE3FF00EC8C9FFB13092B57C5269DB78D2DA4FF00EDFF008579737548FD33C35E1B78DC6FD7EB2FDDD27A79CFA7FE03BBF3B15FCFF6A3CFF6AA3E77BD1E77BD7F017B33FA1B94BDE7FB51E7FB551F3BDE8F3BDE8F661CA5EF3FDA8F3FDAA8F9DEF479DEF47B30E52F79FED479FED547CEF7A3CEF7A3D987297BCFF6A3CFF6AA3E77BD1E77BD1ECC394BDE7FB57A87EC9FF1B3FE1527C4B856EE4F2F45D60ADB5E64FCB173F24BFF00012793FDD66AF21F3BDE8F3BDEBD9E1DCEB1592E654735C0BB54A52525E7DD3F292BA6BAA6CE2CCB2DA58EC2CF095D5E33567FE6BCD3D5799FA940E4515E29FB127C6FFF008599F0E468F7D36FD67C3CAB0B6E3F34F6FD237F7231B4FD013CB57B5D7FAA1C2DC4785CFB2AA19B60DFB95629DBAA7B4A2FCE2EE9F9A3F93336CB2B65F8CA983AFF00141DBD5746BC9AD50514515F4079A145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450079FF8DFF6AEF86BF0D7C5175A2F883C75E18D1F57B2D9E7D9DDDFC714D0EE4575DCA4E46559587B11593FF0DCFF0006FF00E8A6782FFF0006917F8D7E5BFF00C158BFE4FF00FC7DFF0070EFFD36DAD7CEB5F8566FE2C63B078FAD848508354E728A6DBBB519357DFC8FEA6E1EF00F2CCC32AC363EA626A2955A709B4946C9CA2A4D2D36573FA08F86FF0015FC35F187439753F0AEB9A5F8834F82736B25C585C2CD1A4A1558A12BC6E0AEA71E8C2BA0AF8C3FE0863FF2697E22FF00B1BAE7FF0048ECABECFAFD7387B339E6196D1C6D44939ABB4B647F3EF17E4B4F28CE71196519394694B9537BBDB7B681451457B27CD85145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400572FF133E36F847E0C4167278B3C49A3F8763D41996D9AFEE9601395C160BB8F38DC33F515D457E7E7FC17AFFE453F86BFF5F77FFF00A0415F3FC519C54CAF2BAB8FA5152942DA3DB5925D3D4FACE05E1EA59EE79432AAF2718D472BB56BAB4652D2FA743EADFF0086E7F837FF004533C17FF8348BFC68FF0086E7F837FF004533C17FF8348BFC6BF0B68AFC67FE232661FF0040F0FBE5FE67F497FC4B8E53FF0041753EE8FF0091FD14514515FD0C7F2085145140051451400514514005145703FB4A7C6887E067C2BBED5B721D426FF46D3E26E7CC9D81C1C7A28CB1F65C7522BCFCDB34C365B82AB8FC64B969D38B949F9257F9B7B25D5E87560707571788861A82BCA6D24BCDFF005A9F3AFEDFFF001CFF00E124F1545E0FD3E6CD8E8EC25BE287896E48E10FB229FF00BE988EAB5F39F9D55EF7559752BC9AE2E2479AE2E1CC9248EDB9A46272493DC92739A8FED15FE5DF1A71362B88F39AD9BE2B79BF757F2C56918AF45BF7777BB3FAF321C92965781A782A5F656AFBB7BBF9BFB959173CEA3CEAA7F68A3ED15F2FECCF6390B9E751E7553FB451F68A3D9872173CEA3CEAA7F68A3ED147B30E42E79D479D54FED147DA28F661C85CF3A8F3AA9FDA28FB451ECC390B9E751E7553FB451F68A3D98721DC7C11F8B575F06BE24E9FAE5BEF92385FCBBA841C7DA206E1D3EB8E47A3007B57E8F685AE5AF89745B5D42C665B8B3BE89678645E8E8C320FE46BF2A7ED15F5C7FC13BBE3C7DBED6E3C0DA8CDFBDB70D75A5966FBC9D64887D0FCE07A17F4AFE9AFA38F1E3CBB30970EE325FBAAEEF4EFF66A5B6F49A56FF128A5BB3F25F14B867EB38559A505EFD3D25E70EFFF006EBFC1BEC7D51451457F711FCF614514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145007E30FFC158BFE4FFF00C7DFF70EFF00D36DAD7CEB5F457FC158BFE4FF00FC7DFF0070EFFD36DAD7CEB5FC69C51FF239C5FF00D7DA9FFA5B3FD1FE06FF00926F2FFF00AF14BFF4DC4FD59FF8218FFC9A5F88BFEC6EB9FF00D23B2AFB3EBE30FF008218FF00C9A5F88BFEC6EB9FFD23B2AFB3EBFA7F817FE44385FF000FEACFE1EF147FE4ABC77F8DFE4828A28AFAC3E0428A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A002BF3F3FE0BD7FF229FC35FF00AFBBFF00FD020AFD03AFCFCFF82F5FFC8A7F0D7FEBEEFF00FF004082BE27C46FF927713E91FF00D2E27E9BE0E7FC96382F59FF00E9B99F9B5451457F269FDF67F451451457F741FE5B851451400514514005145140013815F9E7FB67FC7BFF0085D1F15A58ECA6F3342D0B75AD960FCB3367F7937FC098000FF7557DEBE96FDBCFE3E7FC2A6F85FF00D8F613797AE78915A08CA9F9ADEDFA4B27B1390A3FDE247DDAF803CDFAD7F21FD23F8EF9E51E17C1CB4569D5B77DE10F97C6FF00EDDECCFDE3C25E17B45E73885BDE30FCA52FFDB57FDBC59F3A8F3AAB79BF5A3CDFAD7F25F21FB872967CEA3CEAADE6FD68F37EB472072967CEA3CEAADE6FD68F37EB472072967CEA3CEAADE6FD68F37EB472072967CEA3CEAADE6FD68F37EB472072967CEA3CEAADE6FD68F37EB472072967CEA3CEAADE6FD68F37EB472072967CEAD2F0978BEF7C0FE27B1D634D9BC8BED36759E17F46539C11DC1E8477048AC4F37EB479BF5AD28CE74AA46AD26E328B4D35A34D6A9AF34C9A946338B84D5D3D1AEE99FA9FF083E27D8FC62F877A6F8834F2A23BE8B3245BB2D6F28E1E33EEAD91EE307A115D357C27FF0004F6F8FDFF000807C406F0AEA3315D27C48E04058FCB05DF453EC1C614FB84F7AFBB2BFD29F0C78DA1C4F9153C749AF6D1F76A2ED34B576ED25692F5B7467F2271970E4B26CCA7865F03F7A0FBC5F4F55B3F4BF50A28A2BF433E5028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A00FC61FF0082B17FC9FF00F8FBFEE1DFFA6DB5AF9D6BE8AFF82B17FC9FFF008FBFEE1DFF00A6DB5AF9D6BF8D38A3FE4738BFFAFB53FF004B67FA3FC0DFF24DE5FF00F5E297FE9B89FAB3FF000431FF00934BF117FD8DD73FFA47655F67D7C61FF0431FF934BF117FD8DD73FF00A47655F67D7F4FF02FFC8870BFE1FD59FC3DE28FFC9578EFF1BFC90514515F587C08514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140057E7E7FC17AFFE453F86BFF5F77FFF00A0415FA075F9F9FF0005EBFF00914FE1AFFD7DDFFF00E81057C4F88DFF0024EE27D23FFA5C4FD37C1CFF0092C705EB3FFD3733F36A8A28AFE4D3FBECFE8A28A28AFEE83FCB70A28A2800A28A2800AA9AF6BB69E18D12F352BF9E3B5B1B185A79E573F2C68A09627E8055BAF917FE0A65FB428D3B4EB7F87FA64FFBEBA0B77AB95FE18FAC50FF00C088DE47A2A76635F27C6DC554387726AD9A56D5C55A2BF9A6F48C7EFD5F68A6FA1EFF000CE435738CC69E0696D27793ED15BBFF002EEDA5D4F9AFF680F8CD79F1CFE2A6A5E20B8DD1C333F9567031CFD9ADD72113EB8E4E382CCC7BD717E73555F3A8F3ABFCD4CC3195F1D8AA98CC54B9AA5493949BEADBBB3FB370981A586A30C3D0568C5249764B42D79CD479CD557CEA3CEAE3F66747B32D79CD479CD557CEA3CEA3D987B32D79CD479CD557CEA3CEA3D987B32D79CD5F407ECE7FB185E7C6AF827E22F11486482F24431E8084ED4B892339919BD55B06307B1DC7B0AF22F81FF0AEFBE37FC4FD2FC3763946BE9733CC1722DA15E6490FD173807A9C0EF5FA9FE14F0BD8F827C3563A3E9B02DB69FA6C0B6F046BFC2AA3033EA7B927924935FBDF823E18D0E20AB5B30CD217C3C13825FCD392B5D7F813BF949C5F467E53E2671954CA29D3C260A56AD26A4FCA29F5FF001356F44FB9F9237293595CC90CC8D14D0B147465C323038208EC4547E7357D21FF000521F801FF00080F8EE2F18E9B06DD27C47215BB0ABF2C1798C93EC24505BFDE57F6AF997CEAFC9F8B385F119066D5B2AC56F4DE8FA4A2F58C97AAB3F27A6E8FBEE1FCDA8E6B80A78FA1B496ABB3EA9FA3FBF7EA5AF39A8F39AAAF9D479D5F39ECCF63D996BCE6A3CE6AABE751E751ECC3D996BCE6A3CE6AABE751E751ECC3D99721BC92DE65923768E48C86565386523A106BF4B7F642F8F49F1F3E115ADE5C48BFDB5A6E2CF5241D4C807CB263D1D70DE99DC3B57E6379D5EADFB1DFED02DF013E305ADDDCCCCBA1EA98B3D4D79DAB193F2CB8F58DBE6F5DBB80EB5FAD7839C6EF86F3C8AAF2B61EBDA153B2FE59FF00DBADEBFDD72EB63E0BC44E15FED7CB1BA4AF5A9DE51EEFBC7E6B6F348FD36A29B14CB3C4B246CAF1B80CACA72181E841A757FA13BEA8FE480A28A2800A28A2800A28A2800A28A2800A28AF1FFDB9FF0068CD6BF657FD9F2FFC61A1E836FAF5D5A5C436F22DC4AC90D9A484A89DC28CBA890C6BB415CF983E618E7971D8CA584C3CF155F48C136EC9BD16FA23BB2BCB6BE618BA781C324E7524A314DA4AEDD96AF43D7A6996DE16924658E38C16666385503A926BC37E2D7FC1493E0CFC1C96482FFC6961A95F4791F65D2036A0F91D54B460C6A7D99857E4DFC7AFDB33E24FED297527FC255E28BEB9B1662574DB76FB3D8C633903C94C2B63B33EE6F735E5F5F88E6FE314DB70CB28AB7F34FF00F914F4FF00C09FA1FD39C3FF00473A4A2AA677896DFF002D3D17FE0724EFF28AF53F4C3C79FF0005E1F0DD848C9E19F00EB5AA2F2165D4AFE3B1C7A1DA8B2E7E991F8579A6B7FF0005DAF1F4EEFF00D9BE0DF085AA91F20B96B8B82A73DCABA678C7A7F4AF86A8AF87C4F88DC4359DFEB1CABB46315FA5FF0013F4FC17837C23868D961149F794A6EFF2E6B7DC91F6A5A7FC1743E2A25C299FC35F0FE48B9DCB1D9DE231FA13727F95757E14FF0082F36B96D228D73E1DE9578BC066B1D4E4B623D480E9267E991F5AF8028AC28F881C414DDE38997CD45FE699D388F097846B47965828AF4728FF00E93247EBD7C1AFF82C5FC23F89F790D9EAD3EA9E0CBD98ED0755841B52DFF5DA32C147BC81057D47A16BD63E28D22DF50D36F2D351B0BB41241736B32CD0CCA7A32BA92187B835FCF157B1FEC95FB6FF008D7F644F14C53E8B7B25F6812C81AFB44B9909B5BA5CFCC5473E5498E9228CE40C861953F75C3FE2ED78D454B3782717F6A2ACD79B5B3F95BD19F96F167D1EF0B2A52AFC3F51C66B5E49BBC5F9296E9F6E6BABEED2D4FDC5A2B8CF801F1E3C3FFB497C2CD37C5BE1BB869B4FD414878A4C09AD255FBF0C8A09DAEA7F02082090413D9D7EEF87C453AF4E35A8B528C95D35B34F667F2BE2B0B5B0D5A587C445C6716D34F469AD1A61451456C738514514005797FED01FB64FC39FD9920C78BBC4969677ECBBE3D3A006E2FA41DBF74996507B33ED5F7AF90FFE0A09FF00056E9B45D4AFBC17F0A2F2359ADD9A0D47C44A036D604868ED73C71D0CBCF7D9D9EBF3AF56D5EEB5ED4E7BDBEBAB8BDBCBA732CD3CF2192599CF25999B2493EA6BF21E2AF14E8E0AA3C2E5715526B4727F0A7D95B597DE97A9FD09C07E04E2732A51C767927469CB5505F1B5DDDEEA09F6B37DD23F46BE257FC17874FB6BA787C1FE01BBBC841F96EB57BE5B727FED8C6AFF00FA33F0F4F2BD5FFE0B93F15EF0BADAE83E03B346C6D3F62BA9245F5E4DC6D39FF77FC6BE31A2BF29C57885C415E5778871F28A4BF257FBD9FBD607C21E12C2C546383527DE4E526FEF76FB9247DA3A3FFC172FE2A5A145BCF0F780EF2351F3116975148DF889CAFF00E3B5EABF0CFF00E0BBFA4DE4F1C3E30F01DFD847C06BAD26F56EB9FF00AE5204C0FF0081935F9B3453C2F887C41425758872F2924FF357FB9863BC20E12C541C6584517DE2E516BEE76FBD347EF27C06FDAC3E1FFED2FA734DE0DF1258EA9344BBA6B36CC37900F5685C07DB9E370054F626BD12BF9E3F0F788B50F096B76DA96957D79A6EA366E24B7BAB599A19A061D195D482A7DC1AFD2CFF0082777FC1573FE166EA561E05F8993C50EBF72EB6FA66B414247A8B9E1629C0C04949C0571857270406C16FD6B84FC50A18FA91C2663154EA3D1497C2DF6D758B7D2EDA7DD688FE7FE3EF0371594D1966193CDD6A31D6516BDF8AEFA6924BAD926BB3576BE59FF82B7581B3FDBDBC6726EDDF6A8B4F940C7DDC58C098FF00C733F8D7CDB5F597FC168349FECDFDB5279B6C63FB4345B3B8CA8C138DF1F3EA7F77F9015F26D7E25C5D4F933BC5AFFA7937F7C9BFD4FE9AF0FAB7B4E19C04BFE9CD35F7452FD0FD59FF008218FF00C9A5F88BFEC6EB9FFD23B2AFB3EBE30FF8218FFC9A5F88BFEC6EB9FF00D23B2AFB3EBFA6B817FE44385FF0FEACFE25F147FE4ABC77F8DFE48F37FDAA3F691B1FD94BE1449E2FD5347D5B58D3ADEEA2B6B84D3C2192DC4990B236F651B776D5EB9CBAD7CCFF00F0FD7F877FF42878D3FEF9B6FF00E3B5F5AFC73F85567F1C7E0F7893C237DB56DF5FB096D03B0CF92E47EEE4FAA38561EEA2BF03FC49E1EBCF09788AFF0049D4216B7BFD32E64B4B989BEF452C6C51D4FB860457C7F88DC4D9CE4D88A53C1C92A534F78A7692DF5F34D5BE67E89E0EF04F0E71260EBC331849D7A525B4DABC24B4D17669A7F23F51FC3BFF0005C3F86FAE7882C6CA6F0DF8B6C22BCB88E07BA9D6DFCBB65660A646C484ED5CE4E06702BED4AFE75EBF6E3FE09D3F1CBFE17EFEC8DE15D5269BCED534B83FB1F51E72DE7DB809B9BFDA78FCB90FFD74A3C39E38C5E6D88AB84CC1A72494A3649689DA4B4F54D7CC3C62F0C72FC83074330CA22D41C9C669B72D5ABC5EBB6D24FE47B8514515FAE1FCFA145145001451450015F21FC5FF00F82C97C3EF845F1435EF0BCBA0F89F549F40BC92C66B9B4583C99258CED70BBA407860CB923A8F4AFA03F69EF8C90FECFDF003C59E3094AF99A2D83C96CADF764B86C24087D9A5641F435F8317B7B36A57B35C5C48D34F70ED249231CB3B139249F524D7E57E23F19E2B27951C3E01A5395E4EE93B4765BF777FB8FDDBC1BF0DF03C450C4633358B74A0D46366D5E5BCB55D95B4FEF1FA87FF000FD7F877FF0042878D3FEF9B6FFE3B5F407EC7DFB61E97FB64F85356D6B45D075CD1F4DD2AE96CBCDD404605CCBB37B2A6C66FBAAC84E71F7C57E1B81B8E0724F415FBA1FB107C081FB397ECC1E14F0CC90F95A9476A2F353C8C31BB9BF79283EBB0B7960FF7635AF3FC3DE2ACEB39C74E38A9AF6508DDDA295DBD22AFF7BF91EC78BDC07C33C3795D396069C957AB2B46F36ED15AC9D9EFD17FDBC7AC514566F8C356BCD03C25AA5F69FA7B6AD7F65672CF6D62B2796D792AA16488360EDDEC02E70719CE0F4AFD8A52518B93E87F39D3839C9423BBD3B7E2F45F334ABCD7E317ED89F0C7E023491F8ABC67A2E9D790FDEB2494DC5E2FD60883483EA5715F957FB46FFC14FBE2B7ED033DCDAFF6C3F84F4290951A6E8CCD6F95F4926FF5AF91C11B829FEE8AF9D9DDA472CC4B331C924F24D7E299D78C34E0DD3CAE8F37F7A7A2F9456B6F569F91FD31C37F475AD522AAE7988E4FEE53B37F39BD13F48C9799FA97F11BFE0B99F0F7407923F0DF867C49E22913EEC9398EC2DE4FA312EFF9A0AF26F117FC178BC5973393A4F807C3D631EEE16EEF66BA6039E32A23E7A738EC78E78F8328AF80C57897C415DE95B91768C62BF169BFC4FD6301E0AF08E1A36786751F79CE4FF04D47F03ED07FF82E7FC592E76F877E1D85CF00D8DE1207FE055741A0FF00C1783C616F719D53C09E1BBC8B3CADADDCD6CD8FAB799FCABE0FA2B8A9F1F71041DD62A5F3B3FCD33D4ABE14709548F2CB030F9392FC534CFD51F859FF0005C4F877E299A287C51E1FF10784E4931BA68F6EA16B17AE5902C9F9466BEB3F85DF18BC2BF1AFC3A356F09EBDA5EBF6190AD259CE243131FE175FBC8DFECB007DABF9FBAEA3E117C68F147C08F195BEBDE13D66F346D4ADC83BE16F9265FEE4887E5910F7560457D7E4BE2E63E8CD47328AA90EAD2E592F3D3DD7E965EA7E7BC4BF47DCA7114DCF25A92A353A464DCA0FCB5BC97ADE5E8CFE8028AF9F7F601FDBBB4CFDB2FC072ADC470E97E32D1635FED5B043FBB901E05C4392498D8F041C946382482ACDF4157EF996E6587C7E1A38BC2CB9A12574FF0047D9AD9AE87F27E7393E332AC64F018E8385483B35F934FAA6B54D68D0514515DC79615F24FC6EFF0082BFF823E05FC58D77C23A87867C55797DA0DC9B59A6B616FE548C00395DD2038E7B815F5B57E1F7FC143FFE4F67E247FD85DFFF00415AFCF7C46E22C6E4F82A55F02D294A767749E966FAFA1FAF783BC1F96F10E675B0B99C5CA31A7CCACDAD79A2BA7933EEDF0F7FC170BE1DEBDAFD8D8FFC22FE30B6FB6DC47079D22DB6C8B7305DC712E7033938AFB52BF9D7AFDE0F0B7ED0BA1E95FB2DE83F11BC49A9DBE9BA4DC68369A9DD5C4AD9DAD2428C500EACE59B6850092DC019AF27C3BE34C5E69F588E6535FBB4A49D94525AF35FD343DEF183C36C0645F53964B4E5FBD728B57726E5EEF2A5D6EEEF447A416DA327803A9AF9DBE3DFFC151FE117C06BA9ACA4D724F136AF012AF65A1A2DD18DBD1A52CB12E0F046F2C3FBB5F9F9FB6C7FC14DFC61FB51EA379A468F35D785FC0F931A69F049B67D413FBD72EBF7B3D7CB1F20C807791B8FCC75E2F11F8B8E337432782697DB975FF0C74FBDFDC7D3706FD1F54E94715C4551C5BD7D9C1ABAFF0014B5D7BA8EDFCC7E8178EFFE0BC7AC5C4ECBE19F87FA6DA44BC2C9AA5FBDC33FB948D63DBF4DC7EBDAB83BDFF82E07C5FBB915A3D2BC076A17AAC7A75C10DF5DD704FE58AF8E68AFCE6BF1F710559734B1325E965F9247EC585F09F84B0F1E5860A2FF00C4E527FF0093367DB5A47FC1753E2643719D43C2BE05B9878F96DE1BA81BDF969DFF00957ACFC33FF82EDF86754B88E1F17782358D194E01B8D32ED2F973EA51C44547D0B1FAD7E66515D185F11B88284AFEDF9976928BFD2FF73472E3FC1DE12C54395E1141F784A516BF1B7DE99FBD5F03BF69FF0001FED1FA5B5D7837C4BA7EB0D1AEE9AD958C7756E3FDB85C0914678C9183D89AEFABF9E7F0AF8B354F037886D756D1B50BCD2F54B1904B6F756B2B45342C3BAB2F22BF52BFE09B3FF0534FF868AB887C0FE387B7B7F1A47113657C804716B4AA32C0A8E1670016C2FCAC01202E307F5CE10F132866756383C74553AAF44D7C327DB5D537D136EFDEF647F3E7885E09E2724A12CC72B9BAD423AC935EFC177D34925D5A49ADED64DAFB3A8A28AFD50FC2028A28A0028A2B17E20FC43D17E15783AFBC41E22D4ADB49D1F4D88CB71733B61500EC075663D02804924000938A9A952308B9CDD92D5B7B24694A94EACD53A69CA4DD924AEDB7B24BAB66D578F7C7BFDBD3E15FECE134D6BE22F145A3EAD09DADA658037978ADFDD644C88CFF00D742A2BF3E3F6CFF00F82B578BBE395F5E687E079AF3C21E11C98FCD89FCBD4B515F5924539894FF00710E70486660703E4077691CB312CCC72493C935F8BF1178B94E94DD0CA20A76FB72BF2FC92B37EADAF468FE92E0EFA3ED6C4538E2B882A3A69EBECE16E6FF00B7A4EE93F249FAA7A1FA45F10FFE0BC7A65B4F2C7E14F87F7D791F223B9D57505B73EC4C51ABFF00E8C15E65AB7FC1747E284F3B7D8BC33E03B584820096DAEA675EBCEE13A8E98FE1EDF857C53457E6D89F113882B4AEF10E3E51515F92BFDECFD9F05E0FF08E1A3CB1C1A93EF294A4DFDF2B7DC91F63D87FC170BE2F5986F3349F015D6EE865D3EE46DFA6DB85FD6BB1F09FFC1793C496B2AFF6E7C3FD0EFE3FE2FB06A12DA1FC37ACBFE7D2BE07A2B1A3C7DC4149DE38A97CECFF0034CE9C47853C255D72CF0305E8E51FFD25A3F58BE187FC16CBE1578BDE387C4163E22F09CEC7E7926B61796ABFF0388990FF00DFB15F4F7C2EF8D9E11F8D9A435F784FC47A3F882DD00321B2B9591E1CF41226772138E8C01AFE7FEBEDCFF821FF00C1793C57F1E75CF1ACCAC2CFC2761F6684E71BAE6E72A3EA044B2E4762E87D2BF43E0FF12336C763E965F88A71A9CEED75EEB4B76F4BA7649BB595FB9F8FF889E0CE4195E535F37C1D59D2F66AFCADA9C5B6D25157B495DB4AFCCEDBD99FA99451457EEC7F2C85145784FEDBFF00B78786FF00634F0947F6855D63C55A9465B4ED223936B30E479D31EA9102319EAC410A3862BC79866187C0E1E58AC5494611DDBFEB56FA25AB3D2CA728C66678B8607014DCEA4DD925F9BE892DDB7A25AB3D8FC5FE34D1FE1FF87EE356D7754D3F47D32D4665BABD9D60863FAB31039F4EF5F277C65FF82D4FC2FF00004D25AF866CF58F1B5DC648F32DD3EC76791FF4D251BCF3DD6320F5CF4CFE6EFED0FF00B51F8DBF6A2F161D5BC5FAC4D79E5B136B651E63B3B153FC31459C2F18058E59B0373135E7B5F83E7BE2F62EA4DD3CAA0A11FE692BC9F9DBE15E8F98FEA8E15FA3DE028D38D5CF6A3A93EB083E582F26FE297AAE5F47B9F74F8AFF00E0BB9E38BCB873A1F827C29A7C24FC8B7D2DC5E328F728D167F215CE47FF0005C1F8BE97AD31D27C06F1B7484E9D73B17E845C6EFCCD7C73457C4D4E3CCFE72E678A97CACBF04923F4EA3E15F09D28F247030B79DDBFBDB6CFBCBC19FF0005E0F1659CCBFF00091780FC3BA947D1BFB3AEE6B26FA8F33CEFCBF5AFA2BE087FC1613E12FC579E1B3D62E2FF00C13A8CC4263558C1B52C7D27425547FB520415F907457A997F89D9F61A4BDA545523DA497E6ACFF13C3CE3C11E15C6C1AA545D1977849FFE932E68FE0BD4FE88B4DD4ADF58B086EACEE21BAB5B841245342E248E553C86561C107D454D5F883FB23FEDD9E38FD913C4119D22F1F52F0DC926EBCD0EEA426D6607EF3275F2A4FF006D7A9037060315FAFDFB357ED2DE19FDAA3E195AF89BC3375BE37C477769211F68D3E6C65A2917B11D8F46182320D7EE1C25C7382CF23ECE2B92AA57706EFA778BEABEE6BB753F983C40F0BF32E1897B69BF6B8793B2A895ACFB496BCAFB6AD3E8EFA2F40A28A2BED8FCC828A28A0028A2BE6CFDBEBFE0A25A2FEC79A28D274F8A0D6BC757F0896D6C1C9F26CD0E409A723076E41C20219B1D547CD5E7E679A61B2FC34B158B972C23D7F44BAB7D123D6C9323C6E6F8C86032FA6E7525B25D3BB6F649756F43DC7E26FC5BF0CFC19F0D49AC78AB5CD3741D363E3CEBB9847BCFF7507576FF00654127D2BE43F8BBFF0005C6F01F85A796DFC1FE1DD6BC5922640B9B871A75ABFA15DC1E43F4645AFCE1F8C9F1C3C55F1FFC6336BDE2ED6AF359D425C8532B623B75CE764683E58D07F75401F8D7275F82E79E2E63EB4DC32C8AA70E8DA5293FBEF15E967EA7F5770BFD1F72AC3535533B9BAD53AC62DC60BCAEAD27EB78FA1F6F6BFF00F05D7F88D7375BB4BF08F826CE1C9C25D25D5CB63B7CCB2C63F4AC8D3BFE0B83F17ACCB79BA4F80EEC363FD6E9F72BB7E9B6E075F7CF4AF8E28AF8D971D67F29733C54BF05F82563F46A7E16F0A423C8B030B79DDBFBDB6FF13F413C09FF0005E4D521B88D3C4FF0FB4FB9858E1E6D2F50781907A88E4570DC76DEBF5ED5F537ECF7FF000530F84FFB44DDC16163AE3E83ADDC61534DD6905ACB237A23E4C4E49E8AAE58FA57E2AD15EE657E296778592F6F25563DA4927F27149DFD6FE87CBE79E0570CE360FEA90961E7D1C64DABF9C64DAB792E5F53FA28AF817FE0BCD61E67C3CF87775BB1E4EA37716DC75DF146739F6D9FAD7857EC2DFF000550F127ECF5A859F877C67717BE25F0392225691BCDBDD217A031313978C71FBB63C01F2EDC61BE83FF0082D1EB3A6FC47FD903C0DE28D16F2DB52D26EB5F85EDAEA1F984B1CB6970C181C703E40083820E011915FA3671C5382CFF00863152C36938C53941EEAD24EFE6B4D1FDF67A1F8DF0EF02E65C29C6F818633DEA7294942A2F86578495BCA5AEB17F26D6A7E61D14515FCE27F651FD14514515FDD07F96E145145001451450072FF19FE2A587C15F869AB78935120C3A7425A38B761AE253C2463DD9881EC327A0AFCA4F1B78DB50F883E2DD475BD52E3CFD43549DAE277EDB98F403B28E8076000AF7CFF82977ED17FF000B13E2447E0ED32E3768FE1790FDA8A9F96E2F7186FC2304A0FF0068C9D78AF98BCCF7AFE1BF1CF8DBFB6737FECDC34AF430CDAD36954DA4FE5F0AF46D68CFEACF0A784FFB372DFAF578FEF6B24FCD43ECAF9FC4FD527B177CD3FDEA3CD3FDEAA5E67BD1E67BD7E1BECCFD53D9977CD3FDEA3CD3FDEAA5E67BD1E67BD1ECC3D9977CD3FDEA3CD3FDEAA5E67BD1E67BD1ECC3D9977CD3FDEA3CD3FDEAA5E67BD7B0FEC4DFB3E37ED09F196DADEEA32DA068DB6F7536C7CAE80FC90FD64618FF007439ED5E9E4B9262734C752CBB08AF52A49457CFABF24B56FA24D9C39A63A865F84A98DC4BB4209B7FE4BCDBD177763EB5FF008272FECF6DF0C7E19378A352876EB5E284578C3AE1ADAD3AA2FD5FEF9F6D9DC57D1D4D8A358635445554501555460003B0A757FA4BC2FC3B86C8F2BA395E13E1A6AD7EADEF293F393BBFC363F88F3DCE2BE6B8FA98FC47C5377B765D12F24AC8E6FE2F7C30D3FE327C39D57C37A92FFA3EA50945900CB5BC8394917DD5803EF8C1E09AFCA3F1EF83752F86FE32D4B41D5A3F2350D2E76B7997B123A303DD48C107B820D7EC057C7FFF000543FD9DFF00B5B45B7F885A5C3FE91A7AADAEAEA8BCC909388E63FEE13B49EB865ECB5F8EF8F7C0BFDA9962CEB0B1FDF61D7BD6DE54F77FF803F797939791FA4F847C53F52C7FF65E21FEEEB3D3CA7D3FF02DBD794F8A3CD3FDEA3CD3FDEAA5E67BD1E67BD7F14FB33FA87D9977CD3FDEA3CD3FDEAA5E67BD1E67BD1ECC3D9977CD3FDEA3CD3FDEAA5E67BD1E67BD1ECC3D9977CD3FDEA3CD3FDEAA5E67BD1E67BD1ECC3D99FA21FF0004DFFDA207C49F86CDE13D4AE37EB5E178C2C059B2D7167D10FD633843E8367A9AFA52BF21FE08FC5ED43E07FC4ED2BC49A7B6E92C251E7439C2DCC2789233FEF2E467B1C1EA2BF597C19E2FB0F1FF0084F4ED6F4B985C69FAA5BA5CDBC98C655864647623A11D8822BFB9FC0EE37FED8C9FFB3B152BD7C3A51D7794368CBCDAF85FA26F567F29F8ADC29FD97997D7682FDD56BBF253FB4BE7F12F5696C69D14515FB79F958514514005145140051451400573DF167E1BD87C61F865AF785B545CD8EBF632D94A76EE31EF520381FDE53861E85457434567569C6A41D39ABA6ACD774F735A15AA51A91AD49DA5169A6BA35AA7F267F3F7ADFC24F10E8BF14350F06AE957D7DE22D36FA5D39ECACE079E59258DCA108AA096C91C6072315F4C7C0CFF0082337C50F89D14377E24934EF02E9F200D8BC3F69BD20F71046703E8EE847A57EA9685F0E7C3FE18F136ABAD69DA2E9765ABEB8EB26A17B0DB225C5E10AAA37B81B9B014704E33CF524D6D57E4795F84380A551CF1D51D457768AF755AFA5DEEDDB7B58FE81CF3E90D9AD7A31A596518D2765CD27EF3E6B6BCABE14AF7B5D4B4EC7C7BF0C3FE08A3F09FC210C6FE20B8F1078BAE47FAC13DD7D8EDDBE890ED71F8C86BD97C39FB047C18F0B2A8B5F86BE12976F4FB658ADE76C73E76ECFE35EBB457E8583E17CA30AAD430D05E7CA9BFBDDDFE27E4798F1D710E3A4E58AC6D497973B4BFF018B515F71E5FABFEC4DF07F5BB7F2E6F863E0545C119B7D16DEDDB9FF6A3553FAF15E15F1E7FE08BDF0D3E2069B71378364BEF04EAFB49895267BBB176FF006E3909700F4CA3803FBA7A57D8B453C770CE538C83A788C3C1F9F2A4FE4D59AF931657C6F9FE5D5555C262EA45AE8E4DC5FAC5DE2FE68FC0EFDA07F67BF147ECCBF11EEBC31E2AB1FB2DF403CC8668C96B7BD88E76CB13E06E4383E841041008207115FB19FF00055EFD9BEDFE39FECB5AA6AF0C0A75EF03C6FABD9CA07CCD028CDCC59FEE98C17C776893DEBF1CEBF9938DB867FB1330787836E9C97345BDEDD9F9A7F859F53FB73C33E365C4D942C5D44A3560F96696D7B5D35E524EFE4EEB5B5CFB17FE08D5FB49CDF0BFF6836F04DEDC37F61F8E14C71A331D905F229689C0EDBD434671D498F3F7457EB1D7F3DFF0EFC6571F0EBE20687E20B4245D687A841A8438FEFC522C83F55AFE82AC2FA3D4EC61B885B743711AC91B7F7958641FC8D7EB7E10E6D3AF97D5C14DDFD934D7A4AFA7C9A6FE67F3F7D21B21A785CDA8667495BDBC5A979CA1657F5719457C89A8A28AFD74FE7B0AF877FE0AFDFB704DF0A7C32BF0CFC2F76D0EBFAFDBF9BABDD42F87B0B36C811291D2497073DC47DBE7047D91F11FC79A7FC2DF006B5E24D564F2B4DD0ACA5BEB861D764685881EAC71803B92057E0AFC5FF8A1AA7C6AF89FAE78AF5890C9A8EBD78F752FCD911827E58D7FD9450AAA3B0502BF2FF14389A797E0560B0EED52B5EED6EA2B7F9BD9795FA9FB9781BC134F37CD259963237A387B349ED29BF8579A8DB99AEFCB7D19CDD3EDADA4BCB88E1863796695822222966763C0000EA49ED4CAFD1EFF00823AFEC3B6C9A443F173C536426B999D97C376D3C60AC28A70D7983FC4581543C6029619DCA47E0BC37C3F5F39C74707434EB27D2315BBFD12EADA3FAAF8CF8B70BC3995CF31C56B6D231D9CA4F64BF16DF449BD76393FD92FFE08B3AA78E34CB5D73E286A175E1DB3B8512C7A2D9EDFB7B29E479CEC0AC39FEE00CD8383B08C57D91E05FF0082717C13F87F691C76BF0F743BE68F04C9A9AB6A0F2118E4F9C58738E8001ED8E2BDBA8AFE9EC9F82727CBA9A852A2A52EB2925293FBF6F45647F10711789BC459C56752B626508F4841B8452ED64EEFD64DB3CCEFFF00630F843A95BF9727C2FF0000AAE7398B41B589BFEFA5407F5AF31F89BFF048DF827F10ED2416BA05E785EF1BA5CE917D2211E9FBB90BC58FA203EF5F4D515EA62B87F2BC44792B61E125FE15F83B5D7C8F0F05C5D9E60E7ED30D8BAB17E5395BE6AF67F347E417ED7FFF0004A1F1AFECD9A65CEBDA1CFF00F0997856D81927B8B787CBBCB15EE6587272A3FBE8480012428AF95036D391C11D0D7F45046E183C83D457E4FF00FC15B7F622B4FD9FFC716BE37F0BD9ADAF857C51398AE2D625DB169B7B82DB500E163914332AF452AE0606D03F10E3EF0EE9E5F45E63965FD9AF8A2F5E5BF54F76BBA776B7BDB6FE9CF09FC60AD9BE2164D9D5BDB3F826925CED6AE324B452B6A9AB27B593B5FE6DF8D9FB40788FF683BBF0FDD789E786F6FF00C3BA3C5A24579B5BCEBA823925911A62490D20F358160064004E5B2C789A28AFC92BE22A569BAB564E527BB7BBE87F416170B470D4950C3C5460B64B44AEEFA2E87EACFF00C10C7FE4D2FC45FF006375CFFE91D957D9F5F187FC10C7FE4D2FC45FF6375CFF00E91D957D9F5FD6DC0BFF00221C2FF87F567F9FBE28FF00C9578EFF001BFC9057E45FFC1637E041F855FB54BF882D61F2F4BF1D5B0D41081851751E23B851EE4F9721F79ABF5D2BE5BFF82BAFC08FF85C3FB255FEA96B0F99AA782251AC4440F98DB8056E173FDD119F30FF00D7115C7E2264DFDA192D451579D3F7E3FF006EEEBE71BFCEC7A5E0FF00127F64712D17376A75BF772FFB7ADCAFE5251D7B5CFC7BAFBC3FE086BF1CBFB03E26F89BE1FDDCC16DFC416C353B0566FF0097983891547ABC4DB8FB41F9FC1F5D97ECF7F172EBE03FC6DF0BF8C2D7CC67D07508EE64443833439DB2C79FF6E32EBFF02AFE6DE17CD9E599A51C6748CB5FF0BD25F837F33FB378E387D677916272DB7BD38BE5FF001C7DE8FF00E4C95FCAE7EFB5155F47D5EDBC41A45ADFD94C971677D0A5C412A1F9658DC06561EC4106AC57F64C649ABA3FCE4945C5F2CB70A28A2992145145007E7FFF00C1747E38FF00667853C27F0EED66DB2EA729D6B5050707C98F747029F5567329FAC2B5F9AF5EC1FB78FC72FF008685FDAAFC5DE20866F3B4D4BA361A6907E536D07EED187B3ED327D6435E3F5FC83C699C7F696715B131778A7CB1FF000C745F7EFF0033FD0EF0D7877FB178770D8392B4DAE79F7E69EAD3F4BA8FC8F7AFF826AFC07FF85FBFB5D786ACAE21F3B49D0DCEB5A88C657CA8086453DB0F2989083D98D7ED757C47FF000445F811FF000877C0FD63C757706DBDF17DD7D9ED1D9791696E4AE41EDBA63267D7CB535F6E57EF5E18E4DF51C9A35A6BDFACF9DFA6D15F76BF33F947C6EE24FED3E249E1E9BBD3C3AF66BFC5BCDFAF37BAFF00C2828A28AFD10FC7CFC57FF82987ECFCDF00FF006B9F105A5ADBF97A4F891C6B5A6845C2EC9D8974503A6D9848A07F742FAD4DF007FE097DF173E3E4505E47A10F0CE8F36185FEB85AD432F5CA45832B6472084DA78F9BBD7EC46B1F0E7C3FE21F1669BAF5FE8BA5DEEB5A3A3C7617B3DB2493D9872A5BCB6232B9DA391CF5F539DAAFC93FE212E0AAE61571588A8FD9CA4DC611D2D7D6CDEBA5EE9249696D4FE81FF88FD9950CA6860709457B6841465526EF76B4BA8AB6AD24DB6DEB7D0F863E13FF00C10C7C17A0C51CDE31F14EB7E22B9032D0D8A2D85B67D0E77BB63D432E7D074AF77F09FF00C136FE07F836058ED7E1DE8B7181CB5FB4B7CCDEFF00BE76FD38F4AF70A2BEDF03C1F92E1236A38687AB5CCFEF95DFE27E639A7889C4B9849CB138DA9E9193847FF01872AFC0F37BAFD8DFE11DE40D1B7C2FF87CAADD4A787AD236FC0AC608FC0D798FC59FF824C7C17F89D652FD97C3F37856F9C1DB75A3DCB45B4F6FDD3EE8B19F4507DC57D2D457662B87F2CC44392B61E125FE15F83B5D7C8F3703C5B9DE0EA2AB86C5D48BF29CBF157B3F468FC52FDB67F601F14FEC69ADC33DD4835CF0A6A1218ECB5882228A1BAF9532F3E5C980481921802413860BE095FD007C69F849A4FC76F85BADF84B5C844BA6EB76CD039C65A16EA92AFF00B48C1587BA8AFC15F1F782EF7E1C78E759F0F6A4A23D4343BE9AC2E54741244E51B1ED9535FCE9E21707D3C97131AB85BFB1A97B27AF2B5BABF55ADD5F5DD74BBFEC5F087C44ADC4982A9431F6FAC51B5DAD14E2F695BA3BAB492D366AD7B2EB3F656F8F97DFB33FC78F0FF8BEC9A431E9F7012F60538FB5DABFCB3467B1CA92467A3053D40AFDDED2356B7D7B4AB5BEB3996E2D2F2249E0957EEC88C032B0F620835FCEFD7ED87FC1337E2049F11BF621F01DD4D2192E34FB47D2E4CF5516D2BC283FEFDA21FC6BEA7C1CCDA6AB57CB64FDD6B9D7934D27F7DD7DC7C37D23321A6F0F86CE60BDE52F6727DD34E51BFA352FBCF78A28A2BF7A3F9482BF0FBFE0A1FF00F27B3F123FEC2EFF00FA0AD7EE0D7E1F7FC143FF00E4F67E247FD85DFF00F415AFC87C63FF0091650FFAF9FF00B6C8FE85FA39FF00C8EF13FF005EBFF6F89E315E95F147F6A5F127C54F82DE07F01DE482DF40F04DBBC7143139C5E4A6472B3483A6523611A8E700391F7C81E6B56B43D12F3C4DADD9E9BA7DBCB797FA84E96D6D044BB9E695D82A228EE4B1007D6BF9FB0F89AF4E32A545B4A6B95A5D55D3B7DE91FD6F8AC1E1AB4A9D7C4453749B945BFB2ECD37F7368D4F869F0BFC41F18FC6369E1FF0BE9379AD6B17CD88ADADD32D8EECC4E02A8EA59885039240AFBEBF67FF00F8219C26CADEFBE25F89A6F3DB0CDA5E8980A9DF0F70EA727B10A83A70C7AD7D41FB0A7EC63A3FEC7DF09E0B3586DEE7C55AA4692EB7A881B9A593AF9487A889338038C9CB1193C7B857F40709F85D84A146388CDA3CF51EBCBF663E4EDF13EF7D3A59EEFF009278FBC72CC313899E0F87E5ECA8C5B5CF6F7E7E6AFF000C7B5973756D5ECBC0FC23FF0004C3F81BE0E81161F01D8DF48BC992FEE67BB673C750EE57B74000F6E4D74AFF00B0AFC1B91194FC33F06618638D32307F3C57ABD15FA4D3C872CA71E5861E9A5E508FF91F8CD6E2CCEEB4B9EAE32AC9F77527FE67CD3F127FE0925F047E20D9C8B6FE1DBAF0D5DB8C0BAD26FA48D97FED9C85E2FF00C733FA57C0FF00B6CFFC131FC59FB2559C9AF58DCFFC251E0BDE11AFE28BCB9EC0B70A278F276A93C0752549C6769201FD8EAADACE8D6BE22D22EB4FBFB786F2C6FA2682E20954347346C0AB2B03C1041208AF9BCFFC3DCA731A2D53A6A954E928AB6BE695935DF4BF668FB3E13F17B88327C445D6AD2AF46FEF426DCB4FEEC9DDC5F6D6DDD33F9E0AB9E1EF105EF8535EB2D534DB99ACF50D3674BAB5B889B6BC12A3064753EA1803F857A97EDD3FB387FC32CFED29AF78620121D21C8BFD25DCE4B5A4B92809EE50868C9EE6327BD790D7F2E6330B5B05899E1EAE938369FAA7D3F467F72E5F8EC3E6582A78BA1EF53AB15257EAA4AFAAFC1AF91FBBBFB20FC7F83F69BFD9E7C37E2F8FCB4BBBEB7F2B50853A417719D92AE3B02C0B283CED653DEBD2EBF3DFFE083FF13E4B8D0BC79E0D9A43E5DACD6FAC5A267FE7A0314C7FF21C1F99FC7F422BFAE78473696659450C5CFE26AD2FF145D9BF9B57F99FE7CF883C3F1C978871597D356846578FF86494A2BE49DBE41451457D21F1A4779790E9D692DC5C4B1C16F02192492460A91A8192C49E0003924D7E38FF00C1483F6EBBCFDAD3E253E99A4DC4D0F80B419CAE9B6E32A2FA519537720EA49C9080FDD43D0166CFD8FF00F0598FDA724F853F03AD7C0FA5DC795AC78E77A5D146F9A1B04C7980F71E6B109E8544A2BF28ABF05F15F8AA6EAFF63619DA2ACEA5BAB7AA8FA2566FBB6BB1FD59E02702538D1FF593191BCA4DAA49F44B494FD5BBC57649F70AE83E18FC2BF117C66F18DAF87FC2DA3DE6B7AC5E7FABB7B64C9007566270A8A3232CC428EE4537E187C37D5FE307C41D23C31A0DB7DAF57D6EE56D6DA3CE1771EACC7B2A8CB31ECA09ED5FB5DFB1DFEC7FE1CFD8FF00E1943A3E9314775AC5D22BEAFAAB2626D426193FF018D724220E00E4E58B31F88E0BE0BAD9ED76DBE4A30F8A5D6FFCB1F3FC12D5F44FF4EF12BC49C370B61A318C554C454F821D12FE6975E54F4496B27A2B59B5F20FC0BFF82164975650DDFC46F16BDAC92005B4DD09159A3E87E6B89148CF5042C647A31AFA13C33FF048DF813E1EB78D66F09DE6AD34783E75EEAF75B988F558E4443F4DB8E2BE96A2BFA132FE05C8B07051861A327DE6B99FFE4D75F7247F22E6DE29714E6151CEA632705DA9BE44BCBDDB37F36DF99F3C6B1FF04A7F80BACA7CDE044B77C6D0F6FAADEC6579CF4136D3F8835E59F107FE0869F0EB5D89DBC3BE25F14787EE187CA27315F5BAFA7C8551FF00F227E55F6D515D18AE0DC8F10B96A6161F28A8BFBE36671E07C46E27C24B9A8E3AAFFDBD2735F74F997E07E49FC64FF82307C56F8790CD75E1F9747F1A59C7CAA59CBF66BC2BEA62970BF82C8C6BEECFF8266FECF773FB3AFEC9FA2D86A9672D8EBDAE48FAC6A70CA852486497012365232ACB12440A9E8C1ABE80A2B8323E03CB329C73C760F9AEE2D59BBA576B55D7A5B56CF5789FC54CEF3FCAE395E63CAD2929394572B9593494ACF96DADF44B54828A28AFB43F363CA7F6CAFDA9F4BFD917E09DF789AF563BAD4A43F65D26C59B06FAE981DAA71C84500B31ECAA40E4A83F897F13FE276B9F193C79A97897C477F36A5AC6AB299A79A43F92A8E8A8A30AAA3800003815EE9FF0548FDA7E4FDA2FF698D42D2CEE3CCF0DF8399F4AD3951898E5756C4F38ED9771804705234AF9B6BF96FC44E2A9E698F787A4FF007349B49746D68E5E7D9797AB3FBA3C1EE03A7916531C6578FF00B4D74A526F78C5EB182EDA59CBFBDA3D9057D8DFB1EFFC120FC53F1DB4AB3F1178D2F26F07786EE944B04022DDA95F46790C11BE58548E433E49EBB0820D74DFF0483FD85ADFE266AA3E2878B2CD6E345D26E0C7A259CA994BDB943F34EC0F548CF0A3BB83FDCC1FD3CAFA4E01F0E696328C732CD137096B186D75FCD27BD9F45A5F7D8F8DF15BC62AF97626593644D2A91D2752C9F2BFE58A7A5D756EF67A257575F3EFC38FF825D7C11F871671A2F836DF5BB85186B9D6277BC793DCA12231FF0001415DC4FF00B1A7C23B8B530B7C2FF87E108032BE1FB556FF00BE8267F5AF4AA2BF6AA390E5B461C94B0F04BCA31FF23F99F15C559CE22A7B5AF8BA929777397F9E9F23E63F8B9FF048EF833F136CA5FB0E8975E12D4181DB75A45CBAA83DB30C85A3C7B2AA93EBD31F9DFF00B67FFC13DBC61FB1CDEC77978D1EBBE14BB97CAB6D62D632AAAE7388E64E4C4E40C8E4A9ECC48207ED6565F8D7C17A57C46F09EA1A1EB9636FA9693AA42D05D5B4EBB92543FD4750472080460815F2BC47E1DE57995193A1054AAF4945595FFBC968D3EF6BF9F43EF3837C60CF326C44562EACABD0BFBD19BE676EF193D535D15F97A5BAAFE7AEBD63F637FDABF59FD90FE3259F88B4F32DCE973E2DF57D383612FED89E47A0917EF2376231F74B02DFDB43F669BAFD947F681D63C29234B369C08BCD2AE5C7373692126324F765C3231E9BA36C718AF29AFE69FF006BCAB1DA5E15694BEE69FE2BF06BC8FED4B65F9EE57AA55285787DF192FC1FE29F668FE853C0FE34D37E23783B4BD7B47BA4BCD2B58B68EEED675E9246EA194E3B1E790790720F35A95F02FF00C10EFF00689935FF0007F883E1A6A171E64BA17FC4DB49563C8B791F6CE83FD959591BEB3B57DF55FD75C379D4336CBA963A1A392D57692D1AFBF6F2B1FE7CF19F0D54C8739AF95CF5507EEBEF17AC5FAD9EBE7741451457B87CB9E4BFB6A7ED51A7FEC8BF03350F12DC2C373AB4C7EC9A45939FF8FBBA6076E40E762005D8E47CAB8072457E23F8E7C71AB7C4BF186A5AFEBB7D36A5ABEAD3B5CDD5CCA7E695DBAFB003A00300000000002BE90FF82B7FED1D27C6DFDA86F343B3B869341F0286D2EDD03E637B9CE6E64C7AEF023FA423D6BE59AFE5BF12389A799E652C3537FBAA4DC52E8E4B494BEFD1792D3767F747835C134F25C9A38CAD1FF68C425293EAA2F58C7CB4D5FF0079D9EC82BE92FD8BBFE09A1E32FDADA28F5A9A4FF845FC1BBF0354B988BC97B838616F1E46FC6082E48407232482B55FFE09B7FB1AFF00C35C7C6A23548E65F07F86C25DEACEBC7DA093FBBB607B19086248E888DC82457ECB693A4DAE83A5DB58D8DB4167656712C30410C6238E18D4615554701400000380057A7E1FF00C3348FF00686617F629DA315A73B5BEBBA8ADB4D5BEAADAF89E2DF8B353229FF64E536FAC357949ABAA69EC927A3935AEBA256D1DF4F9D7E14FFC127BE0AFC32B287ED1E1B93C517D1E0B5DEB372F37987FEB92958B1EDB3F135E951FEC6FF08E2B6108F85FF0F76AAEDC9F0F5A3363FDE31E73EF9CD7A4515FBD61B87F2CC3C3928E1E115FE15F8E977F33F94B19C599DE2EA7B5C4E2EA49F9CE5F82BD97A23C1FC6FF00F04CBF81FE3BB7659FC05A6E9F237DD974D965B26438C6408D829FA1523BE335F27FED2DFF000441BED034CB8D53E186B936B3E482FF00D8DAA944B971E914EA1519BD15D5381F789AFD28A2BCACD782325C7C1C6AD08C5FF34528B5E775BFCD347BD90F89DC4B94D553A18A94E2BECCDB9C5AED6936D7FDBAD3F33F9E4F11787350F086BB75A5EAB6575A76A3632186E2DAE6231CB0B8EAACA7907EB5D4DBFED03E248BE02DD7C3792E22BAF0BCFA9C5AB4314C19A4B19D15D4F94DB80557DF965208C8046D258B7E977FC158FF00621B3F8D7F0BEF3C7DA0D9471F8C7C2F6E67BA3127CDAAD920CBA301F7A48D46E43D48564E72BB7F26EBF9BF89B20C5F0F63A586E67CB24ED25A7345E8D3FC9AF47D8FECBE09E2CC07176570C6F2253A725CD17AF24E3AA69F67BC5F6BA7D50514515F267DF9FD14514515FDD07F96E1451450015E49FB687ED111FECE9F05EF350B7917FB7353CD969687A8958732E3D235F9BD33B477AF59965582267765444059998E0281D4935F955FB6F7ED20DFB457C6ABABAB598B787F47DD65A52F40F183F34DF59186EF5DA101E95F9778B5C6BFEAFE4B2F612B57AD78C3BAFE69FF00DBA9E9FDE713F43F0D784DE779AA5557EE695A53ECFB47FEDE7BF9267944F7925D4EF2492349248C59DD8E5989E4927B934DF37DEAB79A3FBD479A3FBD5FC0DCB7D59FD93ECCB3E6FBD1E6FBD56F347F7A8F347F7A8E40F6659F37DE8F37DEAB79A3FBD479A3FBD47207B32CF9BEF479BEF55BCD1FDEA3CD1FDEA3903D9976DA392F6E6386149269A6608888A599D89C0000E4927B57EAA7EC69FB3D47FB3B7C17B2D3A78D3FB7352C5EEAB20C13E730E23CFF007635C2F5C12188FBD5F237FC1303F671FF00858FF11A4F1AEA90799A3785E402D15C656E2F71953F48C10FFEF14F7AFD11AFEB4F00381FEAF425C458B8FBD3BC69DFA47ED4BFEDE7A2F24FA48FE6CF19F8ABDAD68E478697BB0B4AA79CBA47E4B57E6D7541451457F4B1F830554D7B43B4F13E8979A6DFC11DD58DFC2F6F71138CAC88C0AB29FA826ADD153384671709ABA7A34FA951938B528BB347E46FED33F046F3F679F8C3AA7876E3CC7B58DBCFD3E771FF001F36CE4EC6FA8C156C7F12B5703E6FBD7E957FC145FF0066EFF85D9F075F58D36DFCCF1178555EEA0D8B97BAB7C665878E49C0DEA39395C0FBC6BF333CD1FDEAFF003F3C4FE09970EE773C3D35FB9A9EF537FDD6F58FAC5E9DED66F73FB4BC3EE278E7B94C6B4DFEF61EECD79AEBE925AFADD742CF9BEF479BEF55BCD1FDEA3CD1FDEAFCEB90FB8F6659F37DE8F37DEAB79A3FBD479A3FBD47207B32CF9BEF479BEF55BCD1FDEA3CD1FDEA3903D9967CDF7AFB43FE095FFB487D96F2EBE1CEAB71FBBB82F79A333B7DD7FBD2C03EA332003B893B915F1379A3FBD57BC33E29BCF07788AC756D36E5ED750D3674B9B7994F31C884329FCC57D5F0571357E1ECE29667475517692FE683F897DDAAECD27D0F9DE2AE1BA59D65953015776AF17FCB25B3FD1F74DAEA7ED5D15C3FECE7F1BAC7F684F845A4F89ACF64725D27977902FF00CBADCAE0491FAE01E467AA953DEBB8AFF457038EA38CC3C3178697342694A2D754D5D1FC3B8CC256C2D79E1ABC796706D35D9AD18514515D4738514514005145140051451400514514005145140051451401535ED16DFC49A1DE69D74BE65AEA103DB4CBFDE4752AC3F226BF9E8D4AC5F4BD46E2D642AD25BC8D1315FBA4A920E3F2AFE88ABF9F6F8AB6F1D9FC51F12431A848E2D56E911474502560057E1DE34535CB84A9D6F35FFA49FD41F46DAD2F698FA3D2D4DFFE96BF5302BF7E3F67BBB92FFE01781E799B74B3787EC2476C63731B68C93C57E03D7EFA7ECDFF00F26EFE02FF00B1734FFF00D268EB8BC19FF79C4AFEEC7F367A7F4904BEA5827FDF9FE513B4A28A2BF7E3F930F907FE0B4DF165BC0BFB27C3A0DBCA63BAF186A715A38070C6DE2CCD21FF00BED22523B873F43F92B5F777FC177FC64D7BF18BC0BE1FDDFBBD374697510BCF06E2731E7D3FE5D857C235FCADE2663DE273FAB1E94D462BE4AEFF0016CFEEFF0004F2A8E0F84E84EDEF55729BF9BE55FF0092C62759F027E15DD7C70F8C9E19F08DA3324BE20D461B3691467C98D9879927FC013737D16BF7BBC2FE1AB1F06786B4FD1F4DB75B5D3B4AB68ED2D615FBB1451A85451F450057E51FFC114BE1F278AFF6BC9F579A3DD1F85F45B8BB89C8CED9A4648147B1292CBF91AFD6AAFD33C20CB234B2EA98E6BDEA92B2FF000C7FE0B7F723F14FA436793AF9C51CB22FDCA30E66BFBD37FF00C8A8DBD58514515FAE9FCF61451450015E67FB637C168FF681FD9A3C61E1730F9D777960F2D800B965BA8BF7B0E3BF3222838EA188EF5E994573E2F0B4F1342787AAAF19A69FA35667665F8EAB83C553C5D0769D392927E71775F8A3F9D7A2BBBFDA87C12BF0DFF690F1E6851C7E5DBE97AF5EC102E31FBA133F97C7BA6D35C257F12E2284A8D59519EF16D3F54EC7FA6783C543138786229ED38A92F46AE8FD59FF008218FF00C9A5F88BFEC6EB9FFD23B2AFB3EBE30FF8218FFC9A5F88BFEC6EB9FF00D23B2AFB3EBFADB817FE44385FF0FEACFF003FBC51FF0092AF1DFE37F920AAFABE956FAF6957563790ADC5A5E44F04F137DD911815653EC4122AC515F56D26ACCF838C9C5F34773F037F690F83B73F003E3A78A3C1F75BD8E877EF0C2EC306680FCD0C9FF028D91BF1AE26BF403FE0B9DF01FF00B33C55E17F88D670ED87548CE8BA9328E3CE8C192063FED347E62FD215AFCFFAFE38E2AC9DE579AD6C1FD94EF1FF000BD57E0EDEA8FF0046380F8896799161B316FDE946D2FF001C7497DED5D79347EC47FC1237E39FFC2DFF00D9134CD36E66F3354F054ADA2CC09F98C2A035BB63B28898463D4C4D5F5057E4B7FC1187E39FFC2B6FDA7E6F0BDD4DE5E9BE3AB336C0138517708696127EABE720F5322D7EB4D7F48F87B9C7F686494A5277953F71FF00DBBB7DF1B3F53F8CFC5EE1DFEC8E26AF182B42AFEF23E92BDD7CA4A49795828A28AFB73F310AF17FF82827C73FF867DFD937C59AD433793AA5E5B7F65E9A436D7FB45C7EED597FDA452F27FDB335ED15F99FFF0005CEF8E7FDB3E3DF0BFC3DB49B741A2C0757D4155B833CB9489587F79230CDF49EBE578DB38FECDC9AB6222ED26B963FE2969F82BBF91F79E19F0EFF006D711E1B09257827CF3EDCB0D5A7E52768FCCF81EB53C0FE0EBEF887E34D2741D323F3B52D6AF22B1B54FEF4B2B8451F991CD65D7D8FFF000459F811FF000B17F692BCF175D43BF4EF035A19632C320DE4E1A38863BE104CD9ECCA9EC6BF96321CAE59966147051FB7249F92DE4FE49367F76715E7D0C9B28C466753FE5DC5B5E72DA2BE72697CCFD41F853F0E6C7E10FC34D07C2FA62EDB1D06C62B188EDC1708A14B9FF69882C7DC9AE828A2BFB369D38D382A705649592EC96C7F9B95AB4EB54955AAEF2936DB7D5BD5BF9B0A28A2B4320A28A2800A28A2800A28A2800AFC5DFF0082A9F85E1F0AFEDD9E384B70AB0DF3DADF003B34B6B133E7EAFB8FE35FB455F903FF0005958D53F6DCD44AAAA97D26C8B103EF1D8464FE000FC2BF2BF17A9A964B093DD548FF00E93247EEFF0047AAD28F125482DA5465F84A0CF956BF5BBFE08A57725CFEC69223B6E5B7F105DC718C7DD5D90B63F36279F5AFC91AFD6AFF0082267FC99C5D7FD8C777FF00A2A0AFCDFC267FF0BBFF006E4BF43F65F1F97FC62DFF007121F948FAFA8A28AFE9C3F88C2BF0FBFE0A1FFF0027B3F123FEC2EFFF00A0AD7EE0D7E1F7FC143FFE4F67E247FD85DFFF00415AFC87C63FF91650FF00AF9FFB6C8FE85FA39FFC8EF13FF5EBFF006F89E315F5B7FC11ABE0A47F137F6ACFEDEBC87CCB1F04D935FAE4654DD39F2E107E9991C7BC62BE49AFD3DFF821178356C3E0AF8E3C4181E66A9ADC7A7938E48B781641FADC9FD6BF29F0F72F8E333EA109AD22DC9FFDBA9B5F8D8FDE7C5ECDE797F0A62AA537694D282FFB7DA4FF00F25E63EEBA28A2BFAD0FE010A28A2800A28A2803F39FFE0BCFE088D2F3E1CF89238C09A44BCD36E24C72554C524433EC5A63F8D7E77D7EA17FC176EC124F813E0BBA25BCC875E68947F090F6F2139F7F907EB5F97B5FCABE26D154F886B35F6945FF00E4A97E87F78782589955E10C3297D9738FCB9E4D7E67D89FF0445D7DB4BFDAF351B4DE15353F0E5CC5B49FBCCB35BC831EE0237E04D7EB257E41FF00C11A3FE4F6AC3FEC117BFF00A00AFD7CAFD73C259B964567D2725F827FA9FCF9E3F5351E29BAEB4A0DFDF25F920A28ACEF17F8921F07784F54D5EE3FE3DF4AB496F25F9B6FCB1A173C9E9C0EB5FA64A4A317296C8FC529C2539284756F447E33FF00C14EBE313FC64FDB33C5932CBE658F87E61A1598CE422DBE564C7B198CADC7F7BBF5AF00AB3ACEAF71AFEB1757D74FE65D5ECCF3CCE7F8DDD8B31FC49355ABF89F34C74F1B8CAB8B9EF39397DEEFF81FE9964795D3CB72FA180A5B538463F724AFF3DCFD10FF00821A7ECF915C49E25F8997D06E92DDFF00B134A2CBF7095592E241EF868D011D8C83D6BF462BC4FF00E09CFF000ED7E19FEC57F0FEC7CBD92DF69ABAACA7F899AE98DC0CFB859147B0503B57B657F59705E571CBF26A1412D5C54A5FE296AFEEBDBD11FC0BE2567B3CDB893158993BC549C23FE183E556F5B737AB61451457D41F0A145145001451450015E5BFB6A7C666F803FB2DF8CFC510C9E4DF59D834162C0FCCB733110C2C3BFCAF22B71D94F4EB5EA55F0FFF00C1747C7ADA37C01F0A7876393CB6D775A373201FF2D22B789B23E9BE68CFD5457CFF0015664F019462315176718BB7ABD17E2D1F5DC0793C734E21C26066AF194D732EF18FBD25F38A67E5B96DC727927A9AD5F02F83EF3E21F8DB47D034E4F3350D72F61B0B653FC524AE117F5615955F47FF00C1277C0ABE37FDB8FC26D2A7996FA2A5CEA720DB9E6385C467DB123C673ED5FC9393E07EBB8EA384FE79463F26D26FEE3FD02E21CD165B95E2330FF9F5094BD5C62DA5F37A1FAF3F0A3E1B69DF07BE1A687E16D26311E9FA0D945670FCBB4B84500B9FF698E589EE589AE828A2BFB4A9538D382A705649592EC96C7F9A95AB54AD5255AABBCA4DB6DEEDBD5BF9B0A28A2B4320A28A2803E0FF00F82E9FC238F57F855E12F1B4312FDAB45D41B4BB8651F33413A1752C7D15E2C0F798D7E63D7ECDFF00C158B415D73F60DF1AB6CDD2D8B595D47D3E52B790863CFF00B0CFD2BF192BF98FC58C1C68E79ED22BF8908C9FAEB1FCA28FEDCF00F319E2785FD8CDFF000AA4E0BD1A8CFF003933DC3FE09C1F1464F851FB687816F04BE5DBEA97E347B907EEBA5D7EE46EF60EC8DF541DABF6E2BF9E6F0B6BD27857C4FA6EA90E7CED36EA2BA4C1C1DC8E187EA2BFA178665B885648D9648E4019594E5581E841AFB5F06B192961311857B465192FFB7935FF00B69F9AFD23B2E8D3CC3078E4B59C2517FF006E34D7FE963AB96F8E1F11E2F83FF073C51E299B695F0FE97717CAADFF002D1E38D9913FE04C02FD4D7535F317FC15F7C66DE13FD877C416EADB64D7AF6CF4E5201CE3CE59987E2B0B0E7B135FA96798E783CBABE296F0849AF549DBF13F0BE17CAD6639C617012DAA54845FA392BFE173F1EB52D467D6351B8BBBA964B8BABA91A69A5739691D892CC4FA9249A868AE97E0CF818FC4EF8BDE15F0D8566FEDED5ED74F38EC2599509CF6C06273DABF8C29D39D5A8A11D5C9DBE6CFF492B56A7428CAACF48C536FC9257FC8FD85FF008267FC068FE027EC8BE1BB7921F2F56F114435CD4588C3192750C8A7B8D9108D48F5563DEBDFA996D6D1D9DBC70C28B1C5128444518555030001ED4FAFED6CB7034F0584A784A5F0C2292F92DFE7BB3FCD0CEB34AD9963EB6615FE2A92727F377B7A2D9790514515DC7981451450023A2C88559432B0C10470457E147EDA3F0657E007ED43E32F0BC117936363A834D62B8E16DA6026840F5DA8EAB9F5535FBB15F95FFF0005CAF052E8DFB4978735B8D76AEB7A0AC7271F7A486690139FF71E31F857E55E2E65F1AD9447156F7A9C96BE52D1FE3CBF71FBC7D1F7379E1F8827816FDDAD07A7F7A1EF27F25CDF79F13D14515FCD87F681FD14514515FDD07F96E1451597E38F19E9FF000EBC1FA96BBAB5C2DAE9BA4DBBDD5C487F851464E07727A01DC902B3A95214E0EA5476495DB7B24B76CBA74E7526A9D3576DD925BB6F647CE3FF00053FFDA5BFE154FC2B5F08E9771E5EBDE2D8D92628D86B6B2E9237A8F30E631EA3CCEE2BF36F7D74DF1FFE366A1F1FBE2DEB1E29D4328FA84DFB8877645AC0BF2C718FF7540C9EE727A9AE37CEFF0039AFE05F11B8AA7C459CCF169BF651F769AED15D7D64EF27EB6E88FEE1E01E138E4394C30B25FBC97BD37FDE7D3D22B45E97EA5CDF46FAA7E77F9CD1E77F9CD7C1FB13ED790B9BE8DF54FCEFF39A3CEFF39A3D88721737D1BEA9F9DFE73479DFE7347B10E42E6FAD8F87FE08D4BE26F8DB4BF0FE9107DA352D5EE12DA04EDB98F527B281924F6009ED5CDF9DFE735F7D7FC125FF00668FECCD1AEBE256AF6F8B8BE0F67A2AB8E5210712CE3FDE23603D70AFD9857D7703F08D5E20CDE965F0BA8EF37FCB05BBF57B2F368F95E32E23A591655531F3F8B682EF37B2F4EAFC933EADF823F08F4DF819F0BB48F0BE96375BE97085794AED6B994F324ADEECC49F6E07415D5D1457FA0985C2D2C351861E84796104A292D924AC97C91FC2D89C455C455957ACF9A526DB6F76DEAD8514515D0621451450015F96BFF050CFD9ACFC00F8D335E69F6FE5F86FC4C5EF2C76AE12DE4CFEF60F4F958E40ECAEA3B1AFD4AAF32FDAE3F67BB6FDA53E0A6A5A0308D3548C7DAB4B9DC7FA9B94076E4F656C946F6627A815F9C78A1C18B88B26952A6BF7D4FDEA7EBD63E925A7AD9F43EFBC39E2C79166F1A955FEE6A7BB3F4E92FF00B75EBE975D4FC86DF46FA8F56B0B9D0B54B9B1BC864B6BCB395A09E1906D78A45255948EC41041FA557F3BFCE6BF83A542517CB2DCFEDA8A4D5D6C5CDF46FAA7E77F9CD1E77F9CD4FB11F21737D1BEA9F9DFE73479DFE7347B10E42E6FA37D53F3BFCE68F3BFCE68F621C87D41FF0004D1FDA5CFC1DF8C0BE1CD4EE3CBF0FF008B9D6DC973F2DADDF48A4F60D9D8DFEF29270B5FA655F85A2E369C83FAD7EAE7FC13F3F6991FB467C0FB7FB75C799E25F0EECB1D4C13F34DC7EEA7FF00B68A0E4FF7D1F8C62BFA93C05E306E9CB877152D637952BF6DE51F97C4BFEDEEC8FE6EF1B783F9251CFF000D1D1DA352DDF68CBE7F0BF3E5EE7BB514515FD2C7F3B8514514005145140051457C0FFB797FC15E57C07AADF783FE15BDADE6A96AC60BDF103AACD6F6CE3AADB2F2B2303C17605460E037DE1E2E7BC4182CA30FF59C6CACBA25AB93EC97F497568FA5E16E12CCB88319F53CB61CCF76DE918AEF27D3D356FA267D9DF15FE39783FE06687FDA3E2EF11695A05AB0263FB54E1649F1D4471F2F211E8809AF943E2B7FC1713E1EF85A5960F0A787F5EF164D1FDD9E52BA7DAC9F4660D27E718AFCC5F1A78E759F88FE23B8D635FD52FB58D52EDB74B7579334D23FE2C7A0EC0703B56557E199C78BB9956938E5F054A3D1BF7A5F8FBABD2CFD4FEA3E1EFA3E64D86829E6D5255E7D527C90FC3DE7EBCCAFD91F6A78DFFE0B91F1335A9245D0FC3FE12D0E0627699229AEE75FF8117543FF007C76FC2BCDB5EFF82AFF00C78D75F8F1B0B28F39F2ED74AB38C0EBFC5E516EFD338AF02F0EF85F53F17EA4B67A4E9D7DAA5E3FDD82D2DDE791BE8AA09AF58F09FF00C13C7E3678CC2FD8FE1BF8921DE703EDF12D87A75F3D931D475F7F435F2B1E20E27CC1FEEAAD59FF00839BF2858FBA9708F04E5114ABE1F0F4FF00EBE7237F7D46D8DBEFF82847C6BD4227493E247899564393E55C0888E73C15008FC2B32F3F6DCF8C37D018DFE2778E555BBC7ACCF1B7E6AC0FEB5E97A57FC1207E3B6A3B3CEF0D69B61B8127CFD62D5B67D7CB76EBED9ADAD3FF00E08B3F1A2F43798BE13B4DB8C097542777D3646DFAE2B5594F1655DE15FE7CFF00A984B3EE01A1A2AB855E9ECDFE499E137BFB58FC54D460F2EE3E25FC40B88F39D927886ED973F4325705737325E5C4934D23CB34AC5DDDD8B33B1E4924F524F7AFB26CBFE086DF16EE228DA4D77E1FDBEEFBC8D7F76CC9F95B11FAD7C675E2E7596E6B84E4799C651E6BF2F33BDED6BDB57DD1F4BC379D643987B45924E12E4B737224AD7BDAF64B7B3B7A057EFA7ECDFF00F26EFE02FF00B1734FFF00D268EBF02EBF7D3F66FF00F9377F017FD8B9A7FF00E93475FA6F833FEF589FF0C7F367E27F490FF71C17F8E7F923B4A28A2BF7F3F92CFC85FF0082CCEB1FDA7FB6CDEC3B837F67E8F656F80B8DB956971EFF00EB339F7AF946BE9BFF0082BE7FC9F6F89BFEBD2C3FF4963AF992BF8E78BA4E59DE2DBFF9F93FC1B47FA31E1EC1438630097FCF9A6FEF8A67E87FFC104B4157BDF89DAA347F3469A6DAC4FC746372CE3D7F863F6FAF6FD19AFCFCFF00820A7FC8A7F12BFEBEEC3FF409EBF40EBFA37C378A8F0E61EDFDE7FF0093C8FE37F19AA39F18E32FD3917FE5380514515F727E5E14514500145145007E29FF00C150B485D0FF006F0F8850A8501EE6DAE3E5E9996CE090FE3F3F3EF5E075F457FC158BFE4FFF00C7DFF70EFF00D36DAD7CEB5FC67C4D151CE31715FF003F6A7FE96CFF004878264E5C3B8094B774297FE9B89FAB3FF0431FF934BF117FD8DD73FF00A47655F67D7C61FF000431FF00934BF117FD8DD73FFA47655F67D7F50702FF00C8870BFE1FD59FC39E28FF00C9578EFF001BFC90514515F587C09E4FFB6F7C095FDA33F660F15F8663844DA949686EF4CE32C2EE1FDE4407A6E2BB09F4735F85E46D383C11D457F4515F8A5FF052CF813FF0A0FF006BBF1359DBC3E4E93AEC9FDB7A700BB5447392CEAA318012512A003B28AFC3FC62C9AF0A39A416DEE4BD378FEABE68FE9FFA3AF1272D4C464755EFFBC87AAB466BE6B95DBC9B3C63C0FE31BEF879E34D275ED324F2752D16F22BEB57FEECB138753F981C57EFB7C2EF88363F163E1BE83E26D34FFA0EBF610DFC209C94591036D3FED0CE0FA106BF9F5AFD56FF008225FC71FF0084E3F679D53C1B7536EBEF05DE968109E7EC9705A45C7AE251303E80A0F4AF13C21CE3D8661532F9BD2AABAFF1475FC637FB91F4BF484E1DFACE534B36A6BDEA12B4BFC13B2D7D24A36FF133ED2A28A2BFA2CFE3B21D4B528347D3AE2F2EA54B7B5B58DA69A573858D1412CC4FA0009AFC14FDA4BE304FF1F7E3C78ABC613F99FF0013CD4249A057FBD1403E4850FF00BB12A2FF00C06BF56BFE0AC7F1C7FE14E7EC7FAD5A5BCDE5EA7E3071A1DB807E6F2E404CE7E9E4ABAFD5C57E3757E03E31671CF88A396C1E915CF2F57A2FB95DFF00DBC7F59FD1D7877D9613119D545AD47ECE3FE18EB27E8E565EB10AFD98FF0082567C08FF008521FB2168725CC3E5EABE2D63AEDDE47CCAB2AA8857D78856338E30CCDEF5F95BFB25FC1197F68AFDA2BC2BE115576B6D4EF55AF5973FBBB58F324CD9EC7CB5603DC81DEBF77AD6D63B1B58E186358A18542468830A8A060003D00AAF07B26E7AD5B349AD23EE47D5EB2FB9597CD91F48AE24F6787C3E4749EB37ED27E8AEA2BD1BE67EB1449451457EF87F27851457CD7FB78FFC146F40FD8FEC3FB22C21875FF1C5DC5E643A7EFF00DCD8A9FBB2DC107201EA106198775043579F99E6985CBF0F2C56326A308F5FD12EADF647AD926478ECDF191C0E5F4DCEA4BA2E8BAB6F649756F43E85F1378AB4BF056893EA5AC6A363A4E9B6A374D75793AC10C43D59D8803F135F2F7C61FF0082C87C22F86D34D6BA3CDAB78CAF23257FE25B6FE5DB061EB2CA5411FED20715F987F1E3F698F1B7ED29E243A978C35EBCD51958982DB77976B680F68A25F957D32064F726B83AFC373AF17F15524E195D3508FF0034B593F3B6CBE7CC7F5170CFD1E7034A0AAE7959D49F58C3DD8AF2726B9A5EAB94FBBBC7FF00F05DBF186A5248BE19F04F877488CE42B6A37135F483DFE431007D8823EB5E57E24FF82BCFC76D79DFECFE27D3F4947CFC967A45B1007A03223B0FAE73C75AF9A2DADA4BCB848A18DE59642151114B3313D0003AD7A7F833F625F8B9E3F8E3934BF875E2D9219705269F4F7B689C138C87942A91EE0FAFA1AF8B9715712E612B53AF524FB42EBF08247E951E03E0BCA209D6C35182EF52D2FC6A366D5EFF00C1477E386A13F9927C46D795B18C46228D7F25403F4AC96FDB9BE31B0C7FC2CCF1A73E9AACBFE35DE691FF000493F8F3AA441DFC1B6F66AC011F68D62CC120FB2CA48C770706B7ECFF00E08C1F1AAEAE02496FE18B753FF2D24D50151FF7CA93FA568B2DE2DABAB8577EBCFF00A984B39E00C3E8AA6157A7B27F91E247F6BEF8B4C307E287C4420F51FF000925E7FF001CAE2FC4DE2DD57C6BAAB5F6B3A9EA1AB5F32843717B70F712951D06E724E07A66BEB7B0FF00821EFC5EBC0DE66ADE02B5DBD04BA85C9DDF4DB6EDFAD7CDBFB43FC0BD57F66AF8C3ABF8275BB9D3EF354D1441E7CB62EEF6EC65823986D2EAAC70B20072A3906BCDCDF29CEF0B4155CC6138C1BB2E66ED7B37D5F64CF6F87F88386B1D89950C9AA5395451BB50493E54D26EE92D2ED7DE7155FAD5FF00044CFF009338BAFF00B18EEFFF0045415F92B5FAD5FF00044CFF009338BAFF00B18EEFFF0045415F51E13FFC8F7FEDC97E87C2F8FBFF0024B7FDC487FEDC7D7D451457F4E1FC4615F87DFF00050FFF0093D9F891FF006177FF00D056BF706BF0FBFE0A1FFF0027B3F123FEC2EFFF00A0AD7E43E31FFC8B287FD7CFFDB647F42FD1CFFE47789FFAF5FF00B7C4F18AFD75FF00822ED925AFEC590C8BBB75D6B77923E4F71E5AF1F828AFC8AAFD7BFF0082337FC99358FF00D85EF7FF00425AF85F0917FC2E3FFAF72FCE27EA7F4806D70BAFFAFB0FCA47D5B451457F4C1FC4E14514500145145007C3FF00F05D7FF9377F087FD8C63FF49A6AFCB7AFD48FF82EBFFC9BBF843FEC631FFA4D357E5BD7F2E78A7FF25054FF000C7F23FB9BC09FF924A97F8E7FFA51F567FC11A3FE4F6AC3FEC117BFFA00AFD7CAFC83FF0082347FC9ED587FD822F7FF004015FAF95FA9F847FF002237FF005F25F944FC27E903FF0025447FEBD43F3985797FEDB1AE37877F642F89774B208DFF00E11BBE895CB6DDAD242D18C1F5CB71EF8AF50AF17FF82897FC992FC48FFB0437FE86B5F7B9D4DC32EAF25BA84FFF004967E57C334D54CE3094E5B3AB4D7DF347E1FD14515FC547FA587F421F0F7405F0A780743D2D50C6BA6E9F05A8423054246AB8E3D315B14515FDCD4E0A11505B2D0FF2EEAD49549BA92DDBBFDE1451455198514514005145140057E69FFC179F5C6B8F1F7C3AD37F86D34FBCB903777964894F1DBFD50E7BFE15FA595F96FF00F05D7FF9388F087FD8B83FF4A66AFCF7C50A8E3C3D552EAE0BFF00264FF43F5EF0369A97175093FB31A8D7FE00D7EA7C3F5BDF0EFE28788BE126BEDAAF8635AD4B41D49A2680DCD8CED0CA6362095DCBCE0951C7B0AC1ADCF027C32F127C52D466B3F0CF87B5CF115DDBC7E74B0697612DE4912640DCCB1A9217240C9E32457F2FE1FDAFB45EC2FCDD2D7BDFCADA9FDC58CF61EC64B156E4B6BCD6E5B79DF4B7A9DD7FC373FC64FF00A299E34FFC1A4BFE347FC373FC64FF00A299E34FFC1A4BFE359DFF000C81F16BFE8977C45FFC26EF3FF8DD1FF0C81F16BFE8977C45FF00C26EF3FF008DD7B5ED73CEF57FF273E67D870BFF002E1FEEA668FF00C373FC64FF00A299E34FFC1A4BFE347FC373FC64FF00A299E34FFC1A4BFE359DFF000C81F16BFE8977C45FFC26EF3FF8DD1FF0C81F16BFE8977C45FF00C26EF3FF008DD1ED73CEF57FF270F61C2FFCB87FBA99A3FF000DCFF193FE8A678D3FF0692FF8D1FF000DCFF193FE8A678D3FF0692FF8D677FC3207C5AFFA25DF117FF09BBCFF00E3747FC3207C5AFF00A25DF117FF0009BBCFFE3747B5CF3BD5FF00C9C3D870BFF2E1FEEA637C63FB5AFC4EF885E1ABAD1B5CF1E78A356D26F942DC5A5D6A12490CC030601949C1C100FD4579E57A2FFC3207C5AFFA25DF117FF09BBCFF00E3747FC3207C5AFF00A25DF117FF0009BBCFFE375CB88C2E675E5CD5E1524FBB527F99E86131D92E162E186A94A09BBDA2E095FBD935A9E755FD0BF843FE453D2FFEBD22FF00D0057E16FF00C3207C5AFF00A25DF117FF0009BBCFFE375FBA9E1685EDFC31A6C722B472476B12B2B0C329083208AFD8BC1DC2D7A33C5FB6838DD42D74D7F3F73F9CFE9158EC36269E03EAF5233B3AB7E569DBF87BD997EBE21FF82EBEAFE4FECE9E11B0DCBFE93E245B8DB8E4F976B3AE73EDE6FEA2BEDEAF837FE0BC3FF249BC05FF006179FF00F448AFBFF1024E3C3F896BF957E3248FC9BC25829F176093FE66FEE8C99F9935EE5FF04D4D057C47FB737C3BB764F3047A849758C03830C12CC0F3E9B33F85786D7D15FF00049DFF0093FF00F00FFDC47FF4DB755FCC5C35153CDF0B17B3A94D7FE4C8FEDDE34A8E9F0F63EA477546ABFBA123F67A8A28AFECE3FCDD0A28A2800A28A2800AFCEFFF0082F7690AD69F0BEFD768657D4ADDB8F998116ACBF961BFEFAAFD10AFCFCFF82F5FFC8A7F0D7FEBEEFF00FF004082BE1FC488A7C3989BFF0077FF004B89FA7F835271E31C1DBBCFFF004DCCFCDAA28A2BF93CFEF93FA28A28A2BFBA0FF2DC2BE09FF82BA7ED41E6DD5AFC31D1EE7E58765EEB8C8DD5B8686DCFD062461EF1FA1AFAE7F699F8F1A7FECDDF06758F155FF9723D9C7E5D95BB1C7DAEE5B2228FD705B924745563DABF18FC5BE32D43C73E27D4359D56E9EF352D52E1EEAE667EB248E4B31F6E4F41C0E95F8678D5C5CF0981592E19FEF2B2BCEDD21DBFEDF7A7A277DCFDDBC13E0DFAF635E75898FEEE8BB42FD67DFF00EDC5AFAB5D99179DEF479DEF54FCFA3CFAFE50F627F57FB32E79DEF479DEF54FCFA3CFA3D887B32E79DEF479DEF54FCFA3CFA3D887B32E79DEF479DEF54FCFA3CFA3D887B33D3BF65AF80F7DFB4A7C68D27C316BE6476B33F9FA85C27FCBADAA11E63FA67042AE7AB328AFD90F0EF87ECFC27A058E97A6DBC769A7E9D025B5B4118C2C51A285551EC0002BE73FF8262FECC3FF000A37E0A26BDAA5BF97E24F18225DCC187CD6B6BD618BD89077B0F5600FDDAFA62BFB2BC23E0D592E53F59AF1B56AF693EEA3F663F76AFCDD9EC7F18F8B9C5FFDB19B3C2E1E57A342F15DA52FB52FBD59792BADC28A28AFD60FCA028A28A0028A28A0028A28A00FCEEFF82B57ECC87C21E30B7F88DA3DBEDD375E716FAB2C6BF2C17607C929C741228C138FBE99272E2BE32F3BDEBF6F3E2CFC32D2FE32FC38D63C31AC45E669FAC5B341271968CF55917FDA460AC0FAA8AFC5DF8C3F0CF54F827F13358F0B6B1184BFD1EE0C2EC061665EA922FF00B2EA5587B30AFE44F19782D65F997F6A61A3FBAAEEEEDB2A9BBFFC0BE25E7CDD11FD77E0C7177F69E5DFD978897EF68256EF2A7B27FF006EFC2FCB97B98BE77BD1E77BD53F3E8F3EBF18F627ED5ECCB9E77BD1E77BD53F3E8F3E8F621ECCB9E77BD1E77BD53F3E8F3E8F621ECCB9E77BD7AC7EC61FB494DFB33FC73D375B92491B45BBFF0042D5A15E77DB3919603BB2101C7AED23A135E37E7D1E7D7765B8AAF80C5D3C6E15F2CE9B524FCD7E8F66BAAD0E2CCB2BA18FC2D4C1E255E15138B5E4FF005EA9F47A9FBC9A7EA10EAB610DD5ACD1DC5B5CC6B2C52C6DB96446195607B820839A9ABE38FF008249FED49FF0B03C0171F0F756B80DAB78663F374D2C7E6B8B1240D9EE626207FBAE807DD35F63D7F7B70CE7D4339CB696634369AD57F2C968D3F47F7AB3EA7F0371470FD7C9333AB97623783D1FF345EB192F55F73BAE81451457BC780145145007C61FF0584FDB12EBE0B7C36B5F01F87AE8DBF883C6103BDECF1B624B3B0C946C7A34CDB901ECA92742411F94D5EE5FF0522F89D37C53FDB4BC7774F23343A4EA0DA35BAF68D2D7F7240F62E8EDF5635E1B5FC8FC759E54CCF37AB36FDC83718AE89276BFCDEBF3F23FD04F0B785E8E49C3D429C63FBCA91539BEAE5257B3FF000AB457A5F76CBFE17F0BEA3E36F1159691A4595C6A3A9EA532DBDB5B5BA17927918E02A81DEBF4D3F648FF0082327867C17A5DA6B1F141BFE124D7245121D26194A69F644F3B5D970D330EFC84EA30C304F19FF0437FD9DAD6F4788BE276A16EB2DC5A4C745D20BAFF00A86D8AF7122FB9578D030E80C833C915FA315FA57873C0984A9848E6B98414DCB58C5EA92DAED756FA5F44ADD76FC5FC64F14F1F4B1F3C8B28A8E9C69E939C5DA4E4D5F953DE295ECDAB36EEAF65AE4F837C05A1FC3AD1934EF0FE8FA5E87611FDDB7B0B54B7887FC050015AD4515FB64211845460AC97447F33D4A93A9273A8DB6F76F56FE6145145519857F3AF5FD067C4CF14A781BE1BF8835B924F2A3D1F4DB9BE673FC022899C9EFD36FA57F3E75F8478D151396121D52A8FEFE4FF00267F547D1AE94953CC2A3D9BA4BE6BDA5FF3415FBE9FB37FFC9BBF80BFEC5CD3FF00F49A3AFC0BAFDF4FD9BFFE4DDFC05FF62E69FF00FA4D1D73F833FEF589FF000C7F36767D243FDC705FE39FE48ED28A28AFDFCFE4B3F20FFE0B31A3FF00667EDB57D36DDBFDA1A4595C677677610C59F6FF00578C7B7BD7CA75F767FC177BC1AD63F1A3C0FE20C7C9AA68B2E9E0E3A9B79CB9FD2E47F9C57C275FC85C71877473EC541FF3B7FF00815A5FA9FE86785F8B588E14C0D45D29A8FF00E00DC7F43F453FE0823AD0FF008BA1A7330DDFF12DB98D71C9FF008FA5739FFBE3F3AFD14AFCA1FF008222F8F53C37FB556A9A2CCC153C47A1CD1C43BB4D0BA4A3FF0021896BF57ABF7BF0BF12AAF0FD282FB0E517FF0081397E5247F2878E583950E2EAF51ED523092FFC0147F38B0A28A2BF423F220A28A2800A28A2803F177FE0AB3751DDFEDF7E3F68DB72ABD8213EEBA7DB291F810457CF35EA7FB6F78C17C77FB5E7C46D4924F3636D7AEA089F24EF8E2730A1193D0AA0C7B7A74AF2CAFE2EE20ACAB66989AB1DA5526FEF9367FA51C2587961F23C161E5BC69534FE508A3F567FE0863FF2697E22FF00B1BAE7FF0048ECABECFAF8C3FE0863FF002697E22FFB1BAE7FF48ECABECFAFEA5E05FF00910E17FC3FAB3F857C51FF0092AF1DFE37F920A28A2BEB0F810AF88FFE0B73F023FE133F81BA3F8EACE1DD79E0FBAF22ED80E4DA5C154C9F5DB308F03B798C6BEDCAE7FE2AFC3BB1F8B9F0D75EF0BEA43FD075FB19AC6638C9412215DC3DD49047B815E2F116531CCF2DAD827BCA3A79496B17F7A47D37077104B25CEB0F99C76A725CDE717A497CE2DFCCFE7DEBE8FF00F8255FC72FF8529FB60E831DC4DE5697E2C0742BBC9F9774A4792D8E99132C6327A066F7AF06F1DF832FBE1CF8DB57F0FEA91F93A9689792D8DCA765923728D8F6C8383E959D657B369B7B0DC5BC8D0CF6EEB247229C323039041F50457F21E5B8DAB9763A9E262AD2A724EDE8F54FD7667FA139CE5B4338CAEAE0A6EF0AD06AFBFC4B492F4D1A3FA23A2B86FD9A3E30C3F1FBE027853C6116D0DAE69F1CD3AAFDD8E71F24C83D965575FC2BA3F1E78CACBE1D782358F106A52797A7E876535FDCB7A4712176FC70A6BFB3A962A954A0B1317EE34A49F9357BFDC7F9B95F035E8E2A5829C7F79193835D7993B5BEFD0FCB8FF82D6FC71FF84FBF690B0F08DACDBEC7C13641255046DFB5DC05924E475C46205F62187AD7C695B9F133C7D7DF153E226B9E25D498B5F6BD7D35FCDF36EDAD2396DA3D8670076005635B5B4979711C30C6F2CD2B0444452CCEC780001D493DABF8DF8833496659956C6BFB7276F4DA2BE49247FA35C2791C325C9B0F96C7FE5DC527E72DE4FE726D9FA21FF042EF80FB9BC55F122F21FBB8D0B4D665FF00765B8719FF00B62A08F571EB5FA2B5E73FB24FC118FF00676FD9D3C27E11548D6EB4DB156BE65C1125D49FBC9CE7B8F31980CFF0803B57A357F56708E4FF00D9794D1C235EF5AF2FF13D5FDDB7A23F837C42E23FEDCCFF00118F8BBC1CB961FE08E91FBD2E6F56C28A28AFA43E2CF29FDB4BF695B7FD947F67CD63C58EB14DA8AE2CF4AB790FCB7177267CB07D5540691871958DB1CD7E1F78BBC5DA978F7C51A86B5AC5E4DA86ABAA4EF73757331CBCD231CB31FF0001C0E838AFBAFF00E0BBBF13A6BAF1FF0081FC1B1C8CB6F63A7CBACCC83A48F348618C9F7510C98FF7CFB57C075FCC9E2967953179B3C127FBBA364974726AEDFAEB6F9799FDB7E0570BD1CBF208E6528FEF710DB6FAA826D463E9A3979DF5D9057DCDFB0A7FC1222E3E2F68163E2EF89335F68FA15E0135969107EEAF2FA33CAC92B11FBA8DBB0037B039CA704F8FFF00C1313F676B5FDA2FF6ADD2ACF54805C687E1D85B5ABF8997725C2C4CA238DBB61A578F20F550C3E9FB455E9F86BC1387CC612CCB305CD04ED18F46D6EDF74B64BABBDF45AF89E34F89B8CCA2A4726CA65C95651E69CD6F14F68C7B376BB7BA56B6AEEB8BF84BFB3B781BE04E9EB6FE11F0AE8BA1285DAD2DBDB8FB44A3FDB95B323FD598D7694515FD05470F4A8C153A31518AD92492FB91FC8F8AC557C4D475B1137393DDC9B6DFAB7A8514515B1CE15F8C3FF0562FF93FFF001F7FDC3BFF004DB6B5FB3D5F88FF00F0525F14A78C3F6E3F88D771BF98B0EA2B644E31836F0C76E47E06223F0AFC97C62A896514A1D5D44FEE8CFF00CD1FBFFD1D69C9F10622A2D95192F9B9D3B7E4CF0FAFD6AFF82267FC99C5D7FD8C777FFA2A0AFC95AFD6AFF82267FC99C5D7FD8C777FFA2A0AFCF7C27FF91EFF00DB92FD0FD77C7DFF00925BFEE243FF006E3EBEA28A2BFA70FE230AFC3EFF008287FF00C9ECFC48FF00B0BBFF00E82B5FB835F87DFF00050FFF0093D9F891FF006177FF00D056BF21F18FFE45943FEBE7FEDB23FA17E8E7FF0023BC4FFD7AFF00DBE278C57EBDFF00C119BFE4C9AC7FEC2F7BFF00A12D7E4257EBDFFC119BFE4C9AC7FEC2F7BFFA12D7C2F847FF0023C7FF005EE5F9C4FD4FE903FF0024BC7FEBEC3F299F56D14515FD307F1385145140051451401F0FFF00C175FF00E4DDFC21FF006318FF00D269ABF2DEBF523FE0BAFF00F26EFE10FF00B18C7FE934D5F96F5FCB9E29FF00C94153FC31FC8FEE6F027FE492A5FE39FF00E947D59FF0468FF93DAB0FFB045EFF00E802BF5F2BF20FFE08D1FF0027B561FF00608BDFFD0057EBE57EA7E11FFC88DFFD7C97E513F09FA40FFC9511FF00AF50FCE615E55FB72E8FFDB9FB1CFC4C84A87F2FC397B7182DB7FD544D2E7F0D99C77C57AAD64F8F3C2E9E37F036B5A2C8C163D62C67B2627A01246C87FF0042AFD1730A0EBE16A515F6A325F7A68FC7F29C52C2E3A8E25FD89C65F734FF0043F9EDA2A6BFB19B4BBE9ADAE23686E2DE468A4461F32329C107E845435FC44EEB467FA6E9A6AE8FE84BC07E205F16781B45D5564F31752B182EC3F1F3892356CF1C739EDC56B578AFFC13B3E232FC4EFD8BBE1FDF093CC9ACB4C5D2E6E72CAF6A4DBF3EE4461BDF766BDAABFB6B2DC52C4E1296263B4E3197DE933FCCBCEB032C16615F0735674E728FFE0326BF40A28A2BB8F3028A28A0028A28A002BF32FF00E0BC7A2983E29780352D802DD695736C1F9CB79732B63D38F37F5AFD34AF85FF00E0BB1E037D5FE0A7833C4691EEFEC4D5E5B27207DC4B98B767E9BADD47D48AF87F1230CEB70F6214778F2CBEE92BFE173F50F06B1B1C3717611CB697347FF028492FC6C7E5FD7D95FF00043EF10AE97FB5A6B164EFB5754F0D5C468BC7CD224F6EE3DF8557E95F1AD7B6FF00C139FE2747F09BF6CFF01EA53C823B4BABFF00ECC9CB1C2EDB94680163E8AD22B67FD9AFE6FE13C5AC367386AD2D129C6FE8DD9FE0CFECAE3ECBE58DE1CC6E1A0AF274E4D2EED2BA5F368FDBBA28A2BFB20FF39428A28A0028A28A0028A28A0028A28A002BE1FF00F82EBE91E77ECEDE10BFF2F3F66F118B7DFBBEEF996D33631EFE575ED8F7AFB82BE5FF00F82C1F834F8ABF61ED72E95773E837F67A8803AFFAE101FC96627E80D7CBF1B61DD6C8B1505FC8DFFE03EF7E87DCF8678B587E29C0D497FCFC8C7FF02F77F53F1DEBDE3FE0993ADFF607EDD7F0F27DDB7CCBC9ADB38CFF00ADB59A2C7E3BF1F8D783D75DF00BC74BF0C3E39F83BC44EDB62D0F5AB3BE949E8523991981F62A08FC6BF94727C4AC3E3E8621ED09C65F74933FBDF88B072C66558AC2477A94E71FFC0A2D7EA7EFD51406DC323907A1A2BFB58FF33C28A28A0028A28A002BF3E7FE0BDB751A7873E18C25BF7B25CEA2EABEA156D813F86E1F9D7E8357E63FFC177FC64B7FF17BC0BE1F0DB9B4BD227BF65CFDDFB44DB3FF006DBF97B57C1F8995953E1DAE9FDAE54BFF00038BFC933F56F0530F2ABC618592DA2A6DFA7B392FCDA3E0FA28A2BF94CFEF23FA28A28AF00FF828CFED50BFB31FC04B86B0B858FC51E24DD61A4A83F3C391FBDB803FE99A9183FDF74ED9AFED8CD732A197E0EA63712ED0826DFF0092F36F45E6CFF32B26CA7119A63A965F85579D4692FD5BF24AEDF6499F17FF00C1543F6A9FF85D5F19BFE117D26E3CCF0E7836478328D94BBBCE92C9C750B8F2D7AF472386AF963CE6AAAF746472CCC5998E493D49A4F3FDEBF86F3ECCF119B63EAE6189F8A6EFE8B6497925648FF41387F87F0F9465D4B2EC37C34D5AFDDEEDBF36EEDFA96FCE6A3CE6AA9E7FBD1E7FBD78FEC59ECFB32DF9CD479CD553CFF7A3CFF7A3D8B0F665BF39A8F39AAA79FEF479FEF47B161ECCB7E7357D11FF0004DBFD981BF68FF8EF0DD6A56FE6F85FC2A52FB51DC998EE5F3FB9B73D8EF604B0EE88E38C8AF9D349B1B8D7B55B5B1B3864B9BCBC95208218C6E79646215540EE49207E35FB43FB17FECDB6FF00B2E7C06D2FC3DB636D5E61F6DD5E7439F3AEDC0DC01EEA80045F5080F526BF4AF0C383FF00B5F35556BABD1A3694BB37F663F36AEFC935D4FCABC5BE30FEC3CA1D1A12B57AF78C7BA5F6A5F24ECBCDA7D19EB1451457F611FC481451450014514500145145001451450015F17FFC15E3F65C3E35F01C1F11F47B7DDAA786D041AAA46BCCF664F121EE4C4C79FF0061D89E12BED0AAFAAE976FAE6997165790C773697913413C322EE49518156523B8209047BD783C4D90D0CE72DAB97D7DA4B47FCB25B3F93FBD5D753E8385B886BE4999D2CCA86F07AAFE68BD251F9AFB9D9F43F047CE6A3CE6AF4FFDB5BF670B9FD967E3D6A9E1FDB2368F707EDBA3CEDCF9D6AE4ED04F76420A37A94CF422BC97CFF7AFE1DC7E575F05899E1310AD38369AF35FA767D51FE81E5B8DA18FC2D3C6E165CD0A89493F27FAF75D1E85BF39A8F39AAA79FEF479FEF5C7EC59DBECCB7E7351E73554F3FDE8F3FDE8F62C3D996FCE6A3CE6AA9E7FBD1E7FBD1EC587B33B3F82FF0017F56F819F14745F1568F26DBED1EE04C1376D5B84E8F137FB2E8594FB357ED77C2BF895A5FC61F875A3F89F459BCED375AB65B98492372678646C746560558766522BF07BCFF7AFB87FE08EFF00B562F863C5975F0C758BADB63AE3B5DE8CCE788AE80FDE439F49146E03A6E43DDEBF63F07F8A1E5D8F796621FEEAB3D3B29ECBFF0002F85F9F29F88F8D9C16F31CB7FB5B0D1BD5A0B5B6F2A7BBFF00C07E25E5CC7E90514515FD527F1E851451401F80BFB40BCD27C79F1B35C22C770DAFDF99554E555BED12640FC6B91AF5AFDBBBC0F2FC3BFDB13E2369B24662DDAE5C5EC4A463115C37DA23C7B6C956BC96BF89734A52A58CAB4A7BC6524FD5367FA6991E2215F2DC3D6A7F0CA106BD1C5347EC27FC11D20B787F61CD15A160649751BE69C03D1FCE207FE3A12BEA4AF803FE0867F1E6D6EFC1BE26F873797089A8595D1D674E8DCE1A685D552655FF0071951B1D7F7A4F4071F7FD7F5770362E9E2322C34A97D98A8BF271D1FE57F99FC13E28E5F5B09C538D8565F14DCD79C67EF2B7C9DBD535D028A28AFAC3E0428A2B07E26FC4ED0BE0E781F50F12789351B7D2F47D323324F3CA7F2551D5998F014649240009A8A95214E0EA546925AB6F44977669468D4AD5234A945CA52692495DB6F649756CF9FF00FE0ADBF1D61F841FB22EADA5C770B1EAFE3561A3DAC79F98C470D70D8FEEF940A13D8CAB5F8EB5ECDFB72FED79A87ED89F19E6D72449ACF41D3D4DAE8D6123736D06725D8038F3243F3311FECAE48506BC66BF93B8F388A39C6692AD47F8705CB1F34B77F36DFCAC7F7D7853C1F3E1DC8A387C42FDF547CF3F26D24A37FEEA493E97BD82BF7D3F66FF00F9377F017FD8B9A7FF00E93475F8175FBE9FB37FFC9BBF80BFEC5CD3FF00F49A3AFB5F067FDEB13FE18FE6CFCD7E921FEE382FF1CFF24769451457EFE7F259F1EFFC16B3E14378DBF655B5F11411EEB9F07EAB15C48C0648B79BF72E3FEFE34273E8A7EA3F266BFA0CF899F0FF004FF8ADF0EF5CF0CEA8A5F4FD7AC66B1B8C7DE54910A923D186720F6201AFC15F8B9F0C754F82FF001375CF0AEB11F97A968378F693606164DA7E575FF65970C0F70C0D7F3BF8BD93CA963A9E6315EED45CAFFC51FF0038DADE8CFEC0FA3CF1142BE575B279BF7E8CB9A2BFB92DEDE92BDFFC48D6FD9B3E304BF00BE3CF857C611ABBAE85A8473CE89F7A580FC9320F768D9D7F1AFDE8D1F57B5F106916B7F633C775677D0A5C5BCD19DC9346E032B03DC10411F5AFE77EBF4C3FE08F9FB735B7883C336DF09BC5178B16ABA6A91E1EB899F8BC83926D727F8E3E4A0EE9C0C6C195E13F1253C2626796621DA355A717D39F6B7FDBCAD6F349752BC7CE0DAB8FC153CEF091E69D04D4D2DDD37ADFFEDC776FCA4DEC8FBE28A28AFE893F8F428A28A002B8BFDA2FE2FDAFC04F81DE28F17DD3461743B0927855CF134E46D863FF0081C8517FE055DA57E5CFFC15E3F6E7B2F8C1ACC3F0E3C237CB77E1FD16E3CFD5AF207CC77F74B90B1291C3471E492790CE411F7013F2FC5FC434B27CB678893F7DA6A0BAB93DBE4B77E5F23EE3C3DE11AFC439CD2C2462FD945A9547D1413D75EF2D979BEC99F10DE5E4BA85DCB71348D2CD339924763967627249FA9A8E8A2BF8FF007D59FE882565647EACFF00C10C7FE4D2FC45FF006375CFFE91D957D9F5F187FC10C7FE4D2FC45FF6375CFF00E91D957D9F5FD79C0BFF00221C2FF87F567F9EBE28FF00C9578EFF001BFC90514515F587C0851451401F93FF00F05A6F811FF0AEBF691B3F175AC3B34FF1CDA79B2151855BB802C728FC50C2DEE598FAD7C6F5FB33FF000552F811FF000BBFF641D724B787CDD5BC2646BB67B47CC44408997D4E61690E3BB2AFB57E3357F2BF895937D433A9CE0BDCABEFAF57F17FE4D77E8D1FDDDE0B7127F6AF0D53A551DEA50FDDBF45F03FFC05A5EA99FA73FF000432F8D3FDBBF0B7C51E04BA9B371E1FBC5D4ACD58F2609C6D751ECB2264FBCDF976BFF0599F8DADF0D7F6575F0EDACDE5EA1E38BD5B2201C37D962C4B311F88890FB4A6BE1CFF008256FC613F08BF6CFF000CAC9379361E27DFA15D0FEFF9F8F247FDFF00587F0CD75FFF00059EF8C3FF000B07F6B15F0FC12B359782F4F8ACCA86CA7DA250269587BED6890FBC55F5586E2A71E07953BFEF13F62BD1EBFF00A45D2F43E0F1DC06AA789F0ADCBFBA945621F6BC7DDFC6A72C9F933E45AFA3BFE095DF027FE1777ED7DA149710F9BA5784C1D76F323E526223C95F4E6668CE3BAAB57CE35FAC1FF0458F813FF0AEFF0066FBCF175DC3E5EA1E38BBF3222C30C2D202D1C5F4CB999BB64329F435F13C01937F68E754A125EE43DF97A4765F39597A33F4CF16B893FB1F86ABD583B54A8BD9C7D6774DAF48F335E691F64514515FD687F00851451401F92BFF0005B3FF0093C7B5FF00B172D3FF0046CF5F20D7DC9FF05D7F0449A6FC78F07F887CB2B6FAB686D63BC0E1E4B79DD9BF1DB711FE18AF86EBF90F8EA94A9E7F8A8CBF9AFF0026935F833FD0AF0B6BC2AF0A60650DB912F9C5B4FF0014CFBE7FE083505BB7C47F8852B63ED69A6DA2C63BEC32B97FD4257E97D7E3AFFC1253E3CDA7C13FDAD2CADB54B85B5D2FC5F6ADA2C9239FDDC733B2BC0C7EB2204CF41E69270391FB155FB978538CA757228D287C54E524FE6F997E0FF0F23F977C78CBEB50E299E22A2F76AC20E2FA68945AF935AAF35DC28A28AFD28FC5C28A29B2CAB0C6CEECAA8A0B3331C000773401CEFC61F89FA77C16F85DAF78AF569163B0D06CA4BB93271E6151F2A0FF0069DB6A81DCB0AFC0BF16F89AEFC6BE2AD4F59BF7F32FB56BB96F6E1FFBF248E5D8FE2CC6BEC4FF0082AC7FC1412D7E3DEA4BF0FF00C197866F09E9371E66A17D1B7EEF57B843F2AA7F7A143C83D1DB0C385563F15D7F32F89DC4D4B32C7470B857CD4E95F55B393DEDDD2B249FADB43FB6BC10E09AF92E593C763A3CB5B1167CAF78C15F953ECDB6DB5DAC9EA9A0AFD6AFF82267FC99C5D7FD8C777FFA2A0AFC95AFD6AFF82267FC99C5D7FD8C777FFA2A0A9F09FF00E47BFF006E4BF42BC7DFF925BFEE243FF6E3EBEA28A2BFA70FE230AFC3EFF8287FFC9ECFC48FFB0BBFFE82B5FB835F87DFF050FF00F93D9F891FF6177FFD056BF21F18FF00E45943FEBE7FEDB23FA17E8E7FF23BC4FF00D7AFFDBE278C57EBDFFC119BFE4C9AC7FEC2F7BFFA12D7E4257EBDFF00C119BFE4C9AC7FEC2F7BFF00A12D7C2F847FF23C7FF5EE5F9C4FD4FE903FF24BC7FEBEC3F299F56D14515FD307F1385145140051451401F0FF00FC175FFE4DDFC21FF6318FFD269ABF2DEBF523FE0BAFFF0026EFE10FFB18C7FE934D5F96F5FCB9E29FFC94153FC31FC8FEE6F027FE492A5FE39FFE947D59FF000468FF0093DAB0FF00B045EFFE802BF5F2BF20FF00E08D1FF27B561FF608BDFF00D0057EBE57EA7E11FF00C88DFF00D7C97E513F09FA40FF00C9511FFAF50FCE61451457EA07E1E7E24FFC1483E0EB7C14FDB1BC6360B1F9763AB5D7F6CD91DB8568AE7F7842FB2C8644FF00805786D7EA6FFC169FF66497E237C25D37E20E936C65D4BC1BBA1D4022E5A5B090E771EE7CA939F6596427A57E5957F2371D64B2CB338AB4ADEE49F347D25ADBE4EEBE47FA0DE16F12C33AE1CC3D7BDEA534A9CFBF3452577FE25697CCFD0FFF00821AFED0B15B4DE24F867A85C2ABDC37F6D692AEDF7D82AA5C4633DF6AC4E1476121F5AFD19AFE7CFE1A7C46D5FE11F8FB49F1368374D67AB68B72B736D28E818750C3BAB0CAB0E854907835FB5DFB1C7ED87E1CFDB0BE19C7ABE9522DA6B166AB1EAFA53BE65B0948EA3FBD1B60957EE01070C180FD6FC2BE28A75F08B29AF2B54A77E5BFDA8EF65E71EDDADD99FCFF00E3B70356C2E60F8830B1BD1AB6E7B7D99ED77D94B4D7F9AF7DD5FD7A8A28AFD78FE790A28A2800A28A2800AF27FDB87E0AB7ED01FB2BF8CBC3504266D426B137560ABF79AE602268947FBCC813E8E6BD628AE6C66169E2B0F3C355F8669C5FA356676E5B8FAB81C5D2C6D0F8E9CA325EB169AFC8FE75C8DA707823A8A7417125ACE92C4ED1C91B064753B5948E4107B115F4B7FC1533F65593F671FDA36F350D3ED7CBF0B78C9E4D4B4E28B88E094906783D06D76DC00E024883B1AF99EBF8C334CB6B65F8CA983AFA4A0EDFE4D793566BC99FE926439CE1F37CBA966385778548A6BCBBA7E69DD3F347EE8FEC53FB475B7ED49FB3B683E2849233A9F942CF57897FE585EC6009063B06C89147F7645AF57AFC51FD80FF006D9D43F636F8A4D71324D7FE13D68AC3ACD8C67E72A33B678B271E6264F078652CA7190CBFB21F0CFE27E81F18FC1765E22F0CEA96BAC68FA826F86E206E3DD581C1561D0AB00C0F04035FD3DC0BC5B4B38C1461397EFE0AD25D5DB4E65DD3EBD9E9DAFF00C41E29787F5F87733954A507F55A8DB8496CAFAF23ECE3D2FBC6CFBDB7A8A28AFBA3F2D0A28A2800A28A2800A2B2BC3BE3AD17C5F7BA85B693AB69BA9DC6932882F63B5B9499AD24232124DA4ED6C7383CD6AD4C6719AE68BBA2EA539D3972CD34FB3D37D7F20AE4FE3C7C358FE31FC15F1578564DBFF13FD2AE2CA366E91C8F1B047FF80BED6FA8AEB28A9AD463569CA95457524D3F47A32F0D88A987AD1AF49DA5169A7D9A775F89FCEE5ED94DA6DECD6F711B433DBBB47246C30C8C0E0823D4115157D45FF0568FD9C64F81DFB525F6B16B014D07C745F56B5603E54B8247DA63CFA890EFC76599476AF976BF8B738CB6AE5F8DAB82ABBC1B5EABA3F9AB35EA7FA55C3D9D51CDF2DA39961FE1A9152F47D57AC5DD3F347EDE7FC13CBE3BC7FB417EC9BE15D55A659352D32DC691A90CE596E2DC042CDEEE9B24FF00B695ED95F8E1FF0004CAFDB617F64DF8B52D8EB9349FF085F8A0A43A8100B7D8251C477217A903255C0E4A9CF25403FB13617F06AB630DD5ACD0DCDADCC6B2C3344E1E395186559587041041047041AFEA0E03E24A79B6590BBFDED34A335D6EB452F492D7D6EBA1FC37E2B706D6C833BA9CB1FDC566E74DF4B37771F58B76B76B3EA4D451457DB1F9985145140057E21FFC1447E3447F1D7F6BEF186AF6B2ACDA6D9DC8D2EC595B72B456E3CADCA7FBAECAEE3FDFAFD11FF829F7EDCF65FB36FC2FBBF0BE877CADE3CF11DB9860489FF79A55BB82AD72C47DD6C6446383BBE6E429CFE4157E0BE2E71153AB2865341DF95F34EDD1DAD15EB66DB5E87F577D1F783EB508D5CFF151E5535C94EFD6374E52F46D249F5B4BA0514515F889FD347F445A86A106936135D5D4D1DBDB5B46D2CD2C8C1522451966627800004926BF16FF006E7FDA8E7FDAA3E3EEA5ADC724C341B126C74581B8096C84E1C8ECD2365CE791B82E48515F697FC163BF6B51F0DBE1A43F0DF47B9DBAD78B22F3352646F9ADAC012361F433302BFEE23823E615F979E7FBD7EC1E2F712BAF5A39361DFBB0D67E72E91F92D5F9BEE8FE69F017827D861A5C438A8FBD52F1A77E91FB52FF00B79AB2F24FA48BFE751E7550F3FDE8F3FDEBF11F627F467B32FF009D479D543CFF007A3CFF007A3D887B32FF009D479D543CFF007A3CFF007A3D887B32FF009D479D543CFF007AE9FE0CFC2ED5BE39FC51D13C27A1C7E66A5AE5D2DBC6483B625EAF2363F851033B7B29AD28E0E756A4695357949A492EADE8919622A53A14A55EB3518C536DBD924AEDBF447D8BFF000473FD95FF00E139F1CDCFC4BD62DF7697E1B90DB6928EBF2DC5E15F9A41ED12B0C7FB4E0839435FA5F5CCFC1BF851A4FC0DF85FA2784F448FCBD3743B65B78C91F34ADD5E46C7F13B9676F7635D357F64F0770DD3C932C860E3F1EF37DE4F7F92D97923FCF9E3EE2CA9C439C54C7BBAA7F0C17682DBE6F593F36FA0514515F527C605145140051451400514514005145140051451401F3AFF00C14B3F655FF8696F8073CFA6DBF9BE2AF0AEFBFD336A8DF70B8FDF5BFF00C0D402077744ED9AFC7D32953839FCABFA0CAFC8EFF82AF7ECA7FF000CFDF1D0F88B49B7F2FC2FE35792EE1083E4B4BB0733C3C740490EBD386207DC35F8378C1C26A6A39DE1D6AAD1A9F9465FFB6BFF00B74FE9AF00F8CAD3970E62A5BDE54AFDF7943E7F12FF00B7BBA3E64F3A8F3AA879FEF479FEF5F817B13FA93D997FCEA3CEAA1E7FBD1E7FBD1EC43D997FCEA3CEAA1E7FBD1E7FBD1EC43D997FCEAB5A1788AF3C31ADD9EA5A7DC4B697DA7CC9716F3C670F0C88C19581F5040358DE7FBD1E7FBD38D3717CD1DC52A2A4B964AE99FB91FB1B7ED2367FB537C05D23C4F118E3D482FD9355B743FF001ED768079831D95B21D47F75D7BE6BD4EBF1F7FE0971FB5B7FC3397C7D874BD5AEBC9F0A78C192C6F8C8D88ED27CE20B83D800CC558F002B927EE8AFD82AFEC1E03E26FED9CAE352A3FDEC3DD9FAADA5FF006F2D7D6EBA1FC11E287064B8773A9D1A6BF7353DEA6FC9BD63EB17A77B59BDC28A28AFB43F393F37FF00E0B7FF00B34DC5BEBBA2FC52D36D99ED2E224D27592809F2A45C98266F66526324E002918EAD5F9EF5FD0978F3C0BA4FC4EF06EA5E1FD76CA1D4348D5A06B6BAB7907CB2237EA08E08239040230457E3E7EDC5FF0004E5F14FEC95ADDC6A56515D6BFE059A426DF548E3DCF6609E23B90BF7187037E023646304ED1FCF3E27706D7A58A966F848F3539EB34BECCBABF47BDFA3BDF747F5EF823E2361ABE061C3F8F9A8D6A7A536DE938F48A7FCD1D92EB1B5AF667847803C7FACFC2DF1969FE20F0FEA171A5EB1A5CA26B6B985B0D1B0FD0A919054E43024104122BF49FF00671FF82DBF84FC47A4DAD8FC4AD3AEFC3BAB28092EA361035CD84DEAE51732C79FEE857FAF6AFCC0A2BE0787B8B331C9A6DE0A7EEBDE2D5E2FE5D1F9A699FAC717700E4FC494E31CCA9BE68E919C5DA4976BEA9AF269AEB63F72349FDBF7E0BEB566B3C3F12BC2A88DDA7BC16EFEBF764DADDFD2A0F10FFC1443E09F866DFCCB9F891E1C91719C5A4AD76DFF007CC4AC7F0C57E1F515F78FC64CCB96CA842FFF006F7E57FD4FCB23F473C979EEF1556DDBDCBFDFCBFA1FA99F1A7FE0B7FE03F0ADB4D6FE09D1756F155F6088EE2E57EC3640F63F36656F5C6C5CFA8EDF01FED31FB60F8EBF6B1F104779E2CD53CCB4B662D69A65AA98AC6CF3DD23C9CB60E37B166238CE38AF2FAF5AFD95BF62FF001B7ED71E294B4F0EE9ED06930C816FB59B952B6764BDFE6FE37C748D72C723381961F299971467BC435160DB72527A420AC9FAF576DFDE6D2DF43EF326E06E16E10A32CC14545C56B56A3BC97A37649BDAD149BDB5327F660FD9A3C45FB55FC55B3F0C787E06F9C892FAF5909874EB707E695CFE807566200EB55FF6A0F833FF000CF7F1FF00C55E0D59A6B88742BD30C134A0092585807899B181B8A32938E326BF67BF655FD93BC2BFB23FC3B5D0FC396FE65CDC6D9351D4A65FF48D4A5031B98FF0A8E76A0E1413D4924FE74FFC16CFE1EFFC22FF00B5959EB71AFEE7C4FA2C133B7ACD0B342C3F0448BF3AF7F88B807FB2387E38BABAD7E75CD6D945A6B957CED77DFC8F93E0FF00161F10716CF014172E1BD9CB92EBDE94934DC9F6BC79ACBA2DF56EDF1DD7EFA7ECDFFF0026EFE02FFB1734FF00FD268EBF02EBF7D3F66FFF009377F017FD8B9A7FFE93475EAF833FEF589FF0C7F36783F490FF0071C17F8E7F923B4A28A2BF7F3F92C2BE1BFF0082BEFEC393FC54F0EAFC4CF0AD999B5ED0EDFCBD66DA25F9EFAD10122603F8A4886723A94FF7003F72515E3E7D92E1F36C14F0589DA5B3EA9AD9AF35F8ABAD99F45C2BC4D8BC8332A799E0FE28EEBA4A2F78BF27F83B35AA3F9D7A7DB5CC967711CD0C8F14D130747462AC8C390411D083DEBF48FF00E0A09FF049293C59AA5F78D7E15DAC11DE5C169F50F0F2E23499BAB496A785563D4C47009CED23853F9CBAF6817DE16D66E34ED4ECEEB4FD42CDCC73DB5CC4D14D0B0EAACAC0107D8D7F27F1070D63B26C47B1C5474FB325F0CBCD3EFDD6E8FEF8E11E34CAF88F06B13809AE6B7BD07F145F66BB766B47D0FB73F657FF0082D46BDF0FF4CB5D17E2569B71E2AB1B70234D5AD195752551C7EF158849881FC44A31EAC58926BEBAF05FFC153FE06F8D2D1641E348F4B988CB41A8D94F6EE9F56D850FFC058D7E2ED15F4994F89F9D60A9AA536AAC56DCE9B7F7A69BF9DCF8DE20F03F86B33ACF114E32A127ABF66D28B7FE169A5FF6ED91FB877BFF00050BF827611AB49F123C34C1A3128F2E7321C7B850486FF64F3ED5E75F123FE0B1DF05FC136B27F65EA5AC78AEE972162D3B4E9235DDEEF3F9631EEBBBF1AFC81A2BD4C478C19BCE3CB4A9D38F9DA4DFE32B7DE99E1E0FE8F1C3F4E6A55EB559AED78A4FD6D1BFDCD1F577ED67FF000569F1D7ED11A5DCE87A0C2BE09F0D5C868E68AD27325EDEA1E36C93E170A475540B9C904B0AF94E0824BA9D22891A492460A88A373313C0007726B47C1FE0CD5BE20F88ED747D0F4DBED5F54BE7D905ADA42D34B29F65519E0724F40064D7E9DFFC13BBFE095D0FC0EBEB3F1B7C448ED6FBC5B0E25B0D315966B7D1DBB48CC32B24E3B119543C82CD865F9FCBB2DCEB8B31DCF564E497C537F0C57649595FB456FF007B3EBB38CE786F8072B74E842306F58D38FC737DDB7776EF395ECB457764FE00FDA53F654F10FECB69E118FC49B63BFF001469035492D76157D3D8CAEBE439CF2E1446C7A619CAE0EDDCDE615FA43FF05E8F0879DE1EF873AFAA1FF47B9BDD3E57C1C1F31629101EDC79727E67D38FCDEAF3B8C326A79566D570546FC91E5B5F7778A7F9DCF63C3BE24AD9EE414733C4DBDA4DCF9ADA24D4E4925F24BCFBEA7EACFF00C10C7FE4D2FC45FF006375CFFE91D957D9F5F187FC10C7FE4D2FC45FF6375CFF00E91D957D9F5FD2BC0BFF00221C2FF87F567F15F8A3FF00255E3BFC6FF241451457D61F021451450047756B1DF5AC90CD1ACB0CCA5244719575230411E8457E10FED6BF0464FD9DBF68BF1678459245B5D32F99AC59BAC96B27EF2039EE7CB65071FC408ED5FBC55F9D7FF05D1F8119FF008453E2459C1D33A16A6CA3A7DE96DD8E3FEDB2927FD81E95F97F8AD937D6F2958B82F7A8BBFF00DBAF497E8FD133F71F017893EA19FBCBEA3F73111B7FDBF1BB8FDEB9A3EAD1F9EDA06B973E19D76CB52B393CABCD3E78EE6071FC12230653F8102B4FE29FC43BDF8B5F12B5EF146A3B56FBC41A84DA84CAA4958DA472FB573FC2B9C0F602B068AFE6BF6D3F67ECAFEEDEF6F35757FC59FDA2F0F49D555DC573A4D5FAD9B4DAF4BA46F7C2CF8777FF0016FE24E85E17D31775FEBD7D158C248E11A460BB8FFB2B924FB035FBE9E04F0658FC39F04E91E1FD2E3F274DD12CE2B1B64EEB1C681173EF80327B9AFCC9FF0082237C08FF0084CFE39EB1E3ABB87759F83ED7C8B462383777019323D76C225C8EDE629AFD4CAFE88F08F26F6197CF309AF7AABB2FF0C74FC657FB91FC7FF483E24FAD66F4B28A4FDDA11BCBFC73B3FC23CB6F561451457EB87F3E851451401F35FF00C1547F66B9FF00688FD97AF25D2EDDAE3C41E1194EAF651C6BBA4B8455227840EA4B464B00392D1A0EF5F8D75FD1457E6CFF00C1487FE095FA8D96BDA978FBE18E9CD7DA7DE31B9D4F40B54CCD6AE725E5B741CBA31E4C6BF32927682BC27E2FE2870756C54966D828F3492B4E2B7696D24BAD968FCADD99FD29E06F88B86C0C5E4199CD423277A727A24DEF06FA5DEB16F4BB69BD51F9FA1B69C8E08E86BEF2FD8DFF00E0B2F79F0EFC3D67E1BF89D65A86BD65668B0DB6B56987BD541C013A31025C0FE304360721C9CD7C1F246D0C8C8EACACA70CA46083E869B5F8CE47C418ECA6BFB7C14F95BD1ADD35D9AFE9AE8CFE90E26E13CAF8830BF55CCE9F325AA6B4945F78B5AAF35B3EA99FB73E0FFF0082907C11F1B592CD6DF10B44B4DCBB8C7A86FB174E99044AABC8CF6C83DB239AD3BFFDBDBE0CE9D6E6593E25F84594768AFD656FFBE5727F4AFC33A2BF448F8C79928DA5420DF7F7BF2BFEA7E3B53E8E792B9DE189AAA3DBDC6FEFE55F91FAEDF153FE0B2FF07BC0B6B22E8B71ACF8C2F1410896164D043BBFDA9270981EEAADF435F0CFED6DFF00053BF881FB52DA5CE8EAF1F857C273921B4BB0918BDD27A4F370D20FF640543C654919AF9BEBA4F857F08BC4DF1BBC616FA0F85746BDD6F55B93F2C36E99083A1776385441DD98851DCD7CDE6DC799E670BEABCDCB1969CB4D357BF4EB277ED7B3EC7DA70FF855C2FC3AFEBDC9CD286BCF5649F2DBAECA0ADDED75DCC9F0C786350F1A788ACB49D26CEE350D4F52996DED6DA042F24F231C2AA81DC9AF64FDB3FF00629D4BF637B5F03C7AA5E7DAEFBC4DA5C9737A100F2ADAE924FDE428C3EF04492104F76DC780401FA29FF04FFF00F826D68FFB25D847E20D71AD75BF1F5C4455EE946EB7D2D5861A3833C963C869080482400A09DDC47FC1727E1DFF006FFECE9E1BF11C68CD378775A10B9C7DC86E2360C73FF5D23847E35F413F0E6A60F87ABE3F1ABF7F64D47F9629A6EFDDB57F4F5B9F274FC64A59871761729CB5FF00B2B938CA56F8E4E2D46D7D5454AD6EB2F4B5FF002C2BF5ABFE0899FF00267175FF00631DDFFE8A82BF256BF5ABFE0899FF00267175FF00631DDFFE8A82B87C27FF0091EFFDB92FD0F4BC7DFF00925BFEE243FF006E3EBEA28A2BFA70FE230AFC3EFF008287FF00C9ECFC48FF00B0BBFF00E82B5FB835F87DFF00050FFF0093D9F891FF006177FF00D056BF21F18FFE45943FEBE7FEDB23FA17E8E7FF0023BC4FFD7AFF00DBE278C57EBDFF00C119BFE4C9AC7FEC2F7BFF00A12D7E4257EBDFFC119BFE4C9AC7FEC2F7BFFA12D7C2F847FF0023C7FF005EE5F9C4FD4FE903FF0024BC7FEBEC3F299F56D14515FD307F1385145140051451401F0FFF00C175FF00E4DDFC21FF006318FF00D269ABF2DEBF523FE0BAFF00F26EFE10FF00B18C7FE934D5F96F5FCB9E29FF00C94153FC31FC8FEE6F027FE492A5FE39FF00E947D59FF0468FF93DAB0FFB045EFF00E802BF5F2BF20FFE08D1FF0027B561FF00608BDFFD0057EBE57EA7E11FFC88DFFD7C97E513F09FA40FFC9511FF00AF50FCE61451457EA07E1E57D574BB6D774BB9B1BCB786EACEF22682786550D1CD1B02ACAC0F04104820F506BF1A7FE0A1DFB0CEA3FB21FC4D92E2C209EE7C0DAE4ACFA55E72C2D89C936B2B7691403B49FBEA32390C17F67AB0BE25FC34D0FE30F82350F0E78934DB7D5746D523F2AE2DE51C30EA082395607043020820104115F1FC65C27473DC27B36F96AC7584BB3ECFC9F5EDBF93FD0FC39E3FC470BE61ED9272A33B2A90EE96D25FDE8EB6E8D5D3B5EEBF9F4AE83E187C55F117C18F195AF883C2FAB5E68BABD99FDDDC5BBE091C65187474381956054F706BE9CFDB2FFE092FE30F8157D75ACF8260BDF18F847264DB0C7E66A3A7AFA491A8CC8A3FBE83A025957A9F911D1A372ACA5594E0823906BF97B32CAF1F9462BD9E262E9CE2EE9FA758B5BFAA7F89FDC99367D94F1060BDB60A71AB4E4AD24ECED7DE338BDBD1AD7CD1FA1DF033FE0BA725AD94367F11BC2525CC91801F52D0DD55A4ED936F21033DC91201E8A2BDFF00C3FF00F057BF813ACDB092E3C51A8692C467CBBBD1AE9987B7EEA371FAD7E38D15F5F80F14F3DC34392728D4B7F3C75FBE2E2DFCEECFCF336F02B85B1B51D4A709D16FA539597DD2524BD159791FB2BA9FFC15C3E02D841BE1F195CDF373F243A2DF06FF00C7E251CFD6BCF3C77FF05C9F86DA246CBA0F877C57AF5C283832A4567037A7CC5D9FFF001CEFF857E56D15D188F16B3DA8AD0E4879A8BFFDBA5239707F47FE17A32E6A8EAD4F294D25FF0092462FF13ECDF8BFFF0005B6F897E358A6B7F0B697A1F836DE4CED9957EDF7883D37C8047F8F95F957A3FF00C11A3E2F78A7E317C77F1F6A5E2AF106ADAFDE1D220C497B72D2F963CE3C20270ABECA001E95F9D75F7D7FC107F47BC1F11BC797FF0065B8FB0B69B0402E7CA3E4993CDCECDD8C6EC738CE7153C23C4199E65C4586FAE5694D5DE97D3E197D9564BEE1F885C239264BC1F8D59761E14DF2C55D2BC9FBF1D399DE4FEF3F4BA8A28AFE9C3F890F37FDAABF667D0FF6AFF83D7FE15D697C9693F7F617AA9BA4D3EE541D92AFAF521871B9598646723F143E3BFC08F127ECE3F126FBC2DE29B16B3D46CCEE475CB437711276CD1363E646C707A8208201040FDF7AF2FF00DA9BF647F087ED71E06FEC8F135A15BAB60CDA7EA76E02DDE9EE7A9462395381B90FCAD81D08047E75C77C0B0CEA9FD670D68D78AD3B49767E7D9FC9E9B7EC5E1678A35386AB3C1E32F2C2CDDDA5AB83FE68F74FED47AEEB5D1FE13D7A17ECFBFB52F8E3F660F11B6A3E0ED72E34F13106E6CDFF007B67780769226F95B8C80C30C327045771FB57FF00C13ABE217ECA979717579A7BEBDE17424C7ADE9D1B490A276F3D79684F4CEEF973C066AF05AFE70AD87C7E538AE5A8A54AAC76DD3F54D74F35A33FB2B0B8BCAB3FC0735270AF426B5DA517E4D3D9AECD2699FA4DF083FE0BB3A4DDDBC30F8F3C177D6571C07BBD0E559E273EBE4CACAC83DBCC73FCABDA742FF82BCFC07D5EDD64B8F14DFE96C464C775A35DB30F6FDD46E3DFAE2BF1C28AFB5C1F8AD9ED08F2CDC2A79CA3AFFE4AE27E6999780DC2D8A9B9D28D4A37E909E9FF0093A9D8FD98BDFF0082B57C03B583747E369AE1B38D91E897E1BFF1E840FD6B8DF15FFC16DBE10E871B7F67D8F8C35A93F87C8B08E18CFD4C92291FF7C9EA3F0FC99A2B7ADE2E6793568C69C7D22FF5933930FF0047DE18A6EF39559F939C7FF6D8459F7F7C4BFF0082EFEB97B1C917843C09A6E9C7A25CEAD78F7448F5F2E311807DB791F5E95F317C6AFDBDFE2D7C7C8A6B7D7FC63A92E9B3139B0B0C595A953FC2CB105F3073FF002D0B1F7AF1EAB1A5E9575AE6A10D9D95B5C5E5D5C3048A18233249231ECAA3927D857CAE65C5F9D661EE62311269F45EEA7E568D93F9DCFBDC9BC3DE1BCA3F7983C242325AF34BDE92F3529B6D7C9A3F4BBFE083DFF249BC7BFF0061783FF449AFBCABE3CFF8239FC04F187C0EF845E263E2ED06FBC3F26B9A84573670DE0093BC6B1ED25A3CEF8F9E30E01F6AFB0EBFA5380E8D4A390E1A9D58B8C927A3567F137B3F23F8B7C54C551C4715E32B61E6A71725669A69DA114ECD69A356F50A28A2BEB8FCF8F21FDB6FF654B0FDAEFE065FF872630DBEB16C7ED9A3DE38FF008F5BA5076863D7CB704A375E1B3825457E25F8E7C0FAB7C34F186A5A06BB63369BABE933B5B5D5B4A3E689D7AFB107A82320820824106BFA12AF9C7F6F5FF8279E87FB62E84BA8D9C90689E37D3E2D969A894FDDDDA0E90DC01C95CF461964CF191953F97F885C0CF3687D7B04BF7F1566BF9D76FF0012E9DF67D2DFB97843E284720A8F2CCCDBFAB4DDD3DFD9C9EEEDFCAFAA5B3D575BFE32D7D1BFB1C7FC14B7C6FF00B2541168FB53C4DE0F572DFD93772946B5C9CB7D9E5C131E49C952193249DA0926BC93E37FECFDE2FF00D9D3C612687E30D16EB49BC524C4EEBBA0BB41FC71483E5917DC1E3A1C1C8AE36BF9EF098CC7E538BF6945CA9558E8FA3F469EEBC9AB1FD758FCBB2ACFF01ECB1318D7A13D5754FB38C96CFB34D3F33F60BE17FF00C161FE0BF8FACE3FED4D5354F095E30F9A0D4EC5DD7777C49089171E85B6E7D01E2BD087FC141FE0A9B469BFE164786762B6CC7DA0EFCFFBB8DD8F7C62BF0EE8AFD130FE3066F0872D5A74E4FBD9A7F3B4ADF7247E3F8CFA3C70FD4A9CF42B5582ED78B4BD2F1BFDED9FB3BE34FF0082ADFC0BF06DBC857C60DABDC274834ED3EE2667FA39411FE6E2BE5FFDA23FE0B89A9EBFA75CE9BF0D7C3ADA2098145D5F56292DCA0F5481731AB7BB338FF67BD7C05401B8E0724F415E6E67E296798B83A70946927FC8B5FBDB6D7AAB1ECE49E0670C65F5156AB195792D7F78D38DFF00C31514FD25745FF1478A752F1B7886F356D62FAEF53D4F5090CD737573299269DCF52CC7926BD0B4DFD93BC4977FB2BEA9F16AE17EC7E1FB3D421B0B549233BF50DEC524950F6447D899C10CC587050D7BFF00EC25FF00049BD7BE31EA367E26F88B6979E1EF08A159A2D3E5062BED5C641036FDE8A23DD8E188FBA30438FB83F6FCF863637BFB0378EF41D32C60B3D3F49D196E2DADADE3F2E3B78ED1D2701557A2A88BA74C0F4AE8C8FC3FC562B015F33C7A714A12704EFCD2959B527D6D7F9BDF6DF938A3C5BC0E0735C264994B8CDBA908D492B38C21CC938C7A395B476D22B4F8B6FC4FA28A2BF313F6E3DD3E3C69DABFED07F16F5CF17EB5AC66FB5AB832F962DCB2DB46388E25F9FEEA20551F4CF535C8FF00C287FF00A8B7FE4AFF00F674515F5D5B074AB54955AAAF2936DB6DEADEADEE7E7985CC2BE1E8C70F41A8C229249256492B24B4E883FE143FFD45BFF257FF00B3A3FE143FFD45BFF257FF00B3A28ACFFB370DFCBF8BFF00337FEDAC67F3FE0BFC83FE143FFD45BFF257FF00B3A3FE143FFD45BFF257FF00B3A28A3FB370DFCBF8BFF30FEDAC67F3FE0BFC83FE143FFD45BFF257FF00B3A3FE143FFD45BFF257FF00B3A28A3FB370DFCBF8BFF30FEDAC67F3FE0BFC83FE143FFD45BFF257FF00B3AFBD3FE08CDFB34E95E11B1F13F8E269BFB435933FF63DB3343B05A44112590AF27E672C809E30131FC468A2BED3C3FCAF0BFDB54E7C9AC549ADF476DCFCD7C5ACEF1DFEADD6A6AA3B49C53B24AE9BD568AF67D7BAD363EEEA28A2BFA28FE3D0A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800AF29FDB6BE07E93FB407ECD1E28D0F56FDDFD9ED24D46CEE15433DA5CC28CE9201F9A919195761919CD14571E6587A75F0B528D5578CA2D35DD58F4B26C556C363E8E22849C6719C5A6BA34D1F8E7FF0A1FF00EA2DFF0092BFFD9D1FF0A1FF00EA2DFF0092BFFD9D1457F26FF66E1BF97F17FE67F7D7F6D633F9FF0005FE41FF000A1FFEA2DFF92BFF00D9D1FF000A1FFEA2DFF92BFF00D9D1451FD9B86FE5FC5FF987F6D633F9FF0005FE41FF000A1FFEA2DFF92BFF00D9D1FF000A1FFEA2DFF92BFF00D9D1451FD9B86FE5FC5FF987F6D633F9FF0005FE41FF000A1FFEA2DFF92BFF00D9D1FF000A1FFEA2DFF92BFF00D9D1451FD9B86FE5FC5FF987F6D633F9FF0005FE41FF000A1FFEA2DFF92BFF00D9D7EBB7FC13E3E21EADF117F65BD064D6EE7EDDA96925F4B92EC8C35DAC3808EC39F9B61504E4962A5BBE01457E8DE19518D0CD251A5A2941DD5DEB66ADF71F8EF8D9889E2B2384EBD9B8D4567657574EFAA5D74BAF25D91ED9451457EF07F2A8532E6DA3BDB79219A349A1994A3A3AEE57523041078208ED451405EDAA3E64F8E5FF000491F843F18EEE6BEB3D36F3C1BA94C4B34BA248B140EDEF03068C0F6409F5AF8E7F696FF825345F00C49343E3A935287C979D51F4711B00158E0B09CE7A75C0EBD28A2BF37E2EE0FC9A5859E2D61E2A7DD5E3F845A4FE68FD9BC3DF10F88E38FA58078B94A9BE92B4BEE724DAF933E7FF00F8509FF516FF00C95FFECEBDFBF677FF008252AFC75923693C78DA6C7B164655D17CD620A86201F3C63D3383F4ED4515F98F09F0EE5D8CC72A589A7CD1ED792FC9A3F6EE3EE31CE32ECB2588C156E5977E58BFCE2D1F5A7C1EFF008236FC23F86D730DD6B11EB1E32BC8F0D8D4AE0476A187711441723FD976715F53787BC39A7F84745B7D3749B0B3D2F4EB35D905ADA40B0C302F5C2A280AA3D80A28AFE82CB725C065F1E5C1528C3BD96AFD5EEFE6CFE47CEB89B35CDE6A799E22556DB293D17A2D97C922E57907ED55FB11F82FF6C3FEC36F161D5E393C3FE78B5934FB85858897CBDE1B723647EED48F4E7D4D1457563303431745E1F1305283DD3D53B3BAFC55CE1CB733C565F898E2F0551D3A91BDA49D9ABA69FDE9B47907FC3933E0DFFCFD78D3FF0006517FF19AFAB3C1BE16B5F0378434AD12C8CAD67A3D9C3630195B7398E2408BB8E064E1464E0514571E5B9165F97CA52C1518C1CB7B2B5CF4739E29CDF368C619962255546ED733BD9BDEC69514515EB1E0051451400579EFC73FD95BE1FF00ED216021F18F8674FD5A68D0C715DED315E403D126421C0CF3B73B49EA0D1456188C2D1C45374ABC54A2F74D269FC99D583C762309596230B52509ADA516D35E8D599F24FC58FF0082197849AD6F2FBC33E36D77478E08DE6305FD9C7A80E013B5595A220718E771FAD7C67F12FF0064AFF8577E27934DFF008483ED9B013E67D87CBCE1D97A7987FBB9EBDE8A2BF16E3CE12CA70708D5C3515172DECE56FBAF65F247F4A7857E20710663527431B89738C56978C6FF007F2DDFCDB27F857FB1F7FC2CCF11A69FFF000917D877BC69E67D83CCC6E60BD3CC1D339EB5F677C25FF821AF82ECE0B5BCF1478C35FD78488B3791656F1E9F13679DAD932B118E3E5653EE28A2AF81384B28C5D2956C4D15292DAEE56FBAF67F3467E2A7881C4180AF1C3E0B12E1192D6CA29ECBED72DD7C99F5C7C16FD9BFC0FF00B3C690D67E0DF0D69BA1AC8A1259A24DF7170074F3266CC8FEBF331C576F4515FB3E1F0F4A85354A8C5462B6492497C91FCDB8AC657C55575F133739CB7726DB7EADEACF37FDA7BF659F0BFED6DE04B3F0F78ABFB456CAC6F975085EC6658665915248FEF156F94AC8D918EA07A5784FFC3933E0DFFCFD78D3FF0006517FF19A28AF271DC3795E36AFB7C5D08CE7B5DAD743DFCAB8D33DCB687D57018A9D3A776F962ECAEF73DDFF00660FD973C35FB24F806F3C37E159354934FBED41F5290DFCEB349E6BC71467055546DDB12F18EB9AF47A28AF530B85A586A51A1422A318E892D91E163B1D88C657962B153739C9DDB7BB7DD8514515D0728514514005723F1D7E08E83FB457C2ED4BC23E248A69349D50279860709346C8EAEAC8C41DAC194738E991D0D145655A8D3AD4E54AAABC649A69ECD3D1A66F85C555C3568E22849C67169A6B469A774D79A67CDFFF000E4CF837FF003F5E34FF00C1945FFC668FF87267C1BFF9FAF1A7FE0CA2FF00E334515F39FEA5645FF40B0FB8FB1FF8897C53FF0041F53FF023DF3F669FD993C2FF00B287C3D93C37E158EF3EC33DDBDF4D25DCA259A69582A92CC15470A8A00C7415E854515F4386C352C3D28D0A11518C55925B247C8E371D5F195E58AC54DCE72776DEADBF30A28A2B739428A28A0028A28A00F17FDA17FE09FDF0B7F697B89AF3C41E1D8ED75A9BEF6ABA63FD92F18FAB951B643EF22B57C73FB40FFC117B4BF869A1CDABE91E3ED40D9AB616DAF34A49A4E99E6459507FE394515F19C49C2593E2E8D4C4D6A11E7B377578B6FBBE56AFF3B9FA4706F1FF0010E03134B0786C549536D2E576924BB2524EDF2B1F28EB1FB3A7F656AD756BFDB1E67D9A678B77D931BB692338DFED5E81F027F602FF0085D77F0C3FF0967F6679D2347BBFB2FCEC631CFF00AE5F5A28AFC3B26C8B035F1D1A356178DF6BBEFE4EE7F50711714E6985CAA589A156D34AF7E58BE9D9A68FB07E187FC10DFE1F786AE229BC51E24F1078A1E3C1686154D3EDE5F5C85DF260FB480FBD7D67F0A3E0A784FE06F877FB2BC23E1FD2F40B138DE96908569C8E0348FF007A4619FBCE49F7A28AFE88CAB86F2CCB75C1518C1F7B5DFF00E04EEFF13F8F73EE33CEF39D333C4CAA2FE5BDA3FF0080C6D1FC0EA2B8FF008F1F04343FDA33E15EA9E0FF00112DD1D2756F28CA6DA411CC863952552AC41C1DC83B1C8C8EF4515EBD7A34EB53951AAAF19269A7B34F469FA9E061715570D5A188A12719C1A716B74D3BA6BCD33E71FF0087267C1BFF009FAF1A7FE0CA2FFE335EF9FB34FECD3E1DFD94FE1DC9E18F0C49A949A6C97925F137D32CD2F98EA8A7955518C20E31EB4515E4E5FC3795E06AFB7C2508C256B5D2B3B33DFCDB8CB3CCCE87D5B30C54EA42E9DA4EEAEB667A1514515ED9F3215F32FC5EFF00824EFC2EF8D7F12F58F15EB171E2A5D535CB8373722DAFE348839007CAA62240E3D4D14579F98E5383C7C153C6D3538A7749ABEBDCF5F27CFB31CAAACAB65B5A54A52566E2ECDABDEDF7A39BFF0087267C1BFF009FAF1A7FE0CA2FFE335F417ECE5FB3BE81FB2F7C348FC29E1A7D424D2E1B892E54DECC25977C8416F982A8C71E94515CB97F0E65981ABEDB07423095AD74ACEDD8EECDB8CB3BCD287D5B30C54EA42E9DA4EEAEB67F8B3BCA28A2BDA3E6828A28A0028A28A00F32FDA8BF64EF0BFED75E10D3F44F1549AAC767A6DE7DBA13613AC2E64D8C9C9656C8C39E315E1DFF000E4CF837FF003F5E34FF00C1945FFC668A2BC3C770CE558DAAEBE2A846737D5AD743EA32BE34CF72DC3AC2E03153A74D5DA8C5D95DEE779FB39FFC134FE1DFECBDF12E2F15F86E7F1249AA436F25B28BDBC4962DB20C37CA23539E3D6BE82A28AEECBF2DC2E0697B1C1C14237BD968AFDCF2F36CEB1D9A57FACE6156552764AF277765B2FC58514515DC7961451450015E43F1DBF613F859FB45CD2DD7893C27627549B25B52B226CEF19BD59E3C7987FEBA0614515CB8CC0E1F154FD96260A71ED249AFC4EECBF33C6602B2C460AACA9CD758B717F7AB1F2E7C52FF008217F866DACAEAFBC3BE3CD734D86DD1A5F2750B08AF9881CE0323438FAE0D7C81F137F642FF008573AE7D8FFE121FB67CCE37FD83CBFBADB7A7987AD1457E33C73C239460E0AA6168A8B7D9CADBF6BDBEE47F48F85DE20F10663565471B89738C76BC617DBBF2DDFCD995E18FD9A7FE123D76DECBFB6BC9FB41237FD8F76DC027A6F1E95F50FC10FF00822D5A7C4CD1A3D4AFBE225C436EC14B410688039DC33C399C8FFC768A2BCCE09E15CAF1D59AC552E6B5FED4976ECD1ED7899C779EE57878CB015F91BB7D983EAFBC59F49FC26FF8241FC19F865711DC5E697A9F8BAEA32195B5ABBF3220DFF5CA2088C3D9C37E3D6BE95D03C3DA7F85348834FD2EC6CF4DD3ED576436D6B0AC30C2BE8A8A0003D80A28AFDC72EC9F03808F2E0E9461E8926FD5EEFE67F30671C459A66B3E7CCB113AAD6DCD26D2F45B2F9245CA28A2BD23C50A28A28011D16442ACA195860823822BC1BE367FC1347E0EFC739E6BABEF0AC5A2EA731CB5F68AFF00619093D49451E5313EAC84FBF272515C58ECB70B8DA7ECB174E338F6924FEEBEC7A595E738FCB6AFB7CBEB4A94BBC64D5FD6DBAF27A1F31FC5EFF821E68DE1AD2AE352D17E216A96F6B074B7BDD292E64209C0FDE24918FF00C76BE47F197ECBBFF0896B6D67FDB9F68DA8AFBFEC5B3AFB79868A2BF12E3AE13CA704E32C2D1E5BEFACBF26ECBE47F4D785BC7D9F6671A91C7E21CF976BC617E9D54537F32A685FB387F6D6B36B69FDB3E5FDA2411EFF00B26EDB9EF8DF5F44FC09FF00823F47F18E269A4F884D611C792C8BA1F98CC0301807ED031D7AE0D14579BC1FC319663711C98AA5CCBFC525F9347B5E2271BE779660FDAE06B723D35E583EBE7167D13F0E3FE0893F0A7C2CF1CDAF6A1E26F144ABF7E296E96D2D9FFE03128907FDFCAFA5BE147ECF7E07F81B65E47847C2BA2E821976BCB6B6CAB3CA3FDB94E5DFFE04C68A2BF72CB78772CC03BE0E8462FBA5AFDEEEFF0013F97B39E31CEF365CB98E2A7523FCAE4D47FF0001568FE0763451457B47CD051451400514514018BE3CF875A0FC52F0ECBA4789347D375CD326E5EDAF6DD668F3D880C0E18762391DABE4FF008B3FF044BF863E34B99AE7C37A96BDE0F9A43910C720BDB44FA249FBCFFC898F6A28AF1F34E1FCBB3256C6D18CFCDAD57A3566BEF3E8B23E2DCE32677CB3112A69EE93F75FAC5DE2FE68F8DFF68EFF008277FF00C33F6AF25AFF00C261FDADE5B85DDFD95F67CE59C74F39BFB9FAD79C685FB387F6D6B36B69FDB3E5FDA2411EFF00B26EDB9EF8DF4515FCEF9E64180C3E652C3D1A768DD69793EBE6EE7F6070BF1666B8BC9A18BC455E69B8DEFCB15ADBB2497E07D63F003FE08B1A47C45D06DF58D6BC7DA9359CAC035AD96969049F755B895E4907F163EE76AFAF7E007FC13DFE14FECE13C379A1F86E1BDD620219754D51BED976AC3A32161B233EF1AAD1457EDDC3BC2393E128D3C451C3C79EC9DDDE4EFDD39376F958FE63E2FF0010388B1F88AB84C4E2E4E9A6D72AB4535D9A8A8DFE773DAEB37C65E15B3F1DF84355D0F50576B0D66CE6B1B9543B58C52A14700F63B58D1457D84A2A517196A99F9DD3A928494E0ECD6A9F668F94FF00E1C99F06FF00E7EBC69FF8328BFF008CD1FF000E4CF837FF003F5E34FF00C1945FFC668A2BE67FD4AC8BFE8161F71F6DFF00112F8A7FE83EA7FE047FFFD90000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000

	IF EXISTS(SELECT * FROM ASRSysPictures p
		INNER JOIN ASRSysSystemSettings s ON s.SettingValue = p.PictureID
		WHERE s.Section = 'desktopsetting' AND s.SettingKey = 'bitmapid'
			AND p.[name] IN ('Advanced Business Solutions Wallpaper 1024x768.jpg', 'Advanced Business Solutions Wallpaper 1280x800.jpg',
			'Advanced Business Solutions Wallpaper 1440x900.jpg', 'Advanced Business Solutions Wallpaper 2560x1600.jpg',
			'ASRDesktopImagePersonnelnPost.bmp', 'ASR Splash.jpg', 'ASRDesktopImage2005.jpg', 'ASRDesktopImage2005b.jpg',
			'ASRDesktopImage 1024x768.jpg',
			'ASRDesktopImage 1440x900.jpg',			
			'ASRDesktopImage 1600x1200.jpg',
			'ASRDesktopImage-1024x768.jpg',
			'ASRDesktopImage-1440x900.jpg',			
			'ASRDesktopImage-1600x1200.jpg', 
			'COASolutionsDesktopImage-1024x768.jpg',
			'COASolutionsDesktopImage-1600x1200.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201024x768.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201440x900.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201600x1200.jpg',
			'Default Background.jpg',
			'HRProP.bmp', 'HRProPP.bmp', 'HRProPR.bmp', 'HRProPRP.bmp',			
			'HRProPRT.bmp',	'HRProPRTP.bmp', 'HRProPRTS.bmp', 'HRProPS.bmp', 'HRProPT.bmp', 'HRProPTP.bmp', 'HRProPTS.bmp',	'HRProT.bmp'))
	BEGIN
		-- Set backcolour to white, image to our newly inserted one and tile in the centre
		EXEC spsys_setsystemsetting 'desktopsetting', 'backgroundcolour', '16777215';
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmapid', @newDesktopImageID;	
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmaplocation', 2;	
	END

	-- Mobile images --
	EXEC dbo.spadmin_writepicture '023A47FA-A1E7-415E-8F6F-5B344D913160', 'oneadvhead.png', 5, @newMobileHeaderID OUTPUT, 0x89504E470D0A1A0A0000000D494844520000013F000000410802000000289B34F90000000467414D410000B18F0BFC6105000000097048597300000EC300000EC301C76FA8640000196649444154785EED9D696C1D559680DFAFF93FD2FC1AB55A23C468D4EAEE601346A27F1AD0C89646230D8B5A0D466A31E959401E05849A6E48F8310AEA21413D84A59381A6D9C21208B1B3411267C589134CE2381B31B66327C624A4710809590849A6E6BB759E8FCBA7EA95ABEA553976F0D5A7A7F27DF5CACF7EF5D539E7DE5B76C97BBFC6F1C158D6DE605957E3C3C60DDEFA101B426C0CB029C4E6B16C116638B686F830445B806D21B68FA53DC08E103B437CF4D3513A427C1C62D708BB4374FEE4FB4857DDB9DE973F3F767C9AA229796B7C7B15A331188DA1388D414D360E83711852690C31261B8721B3C66193CDF97DED72B1FBBFBF181A3027D93405E1DBAB64D1B882C9C661C8AC71129373D4188AD3184267FCB5C1E583B3BF1CEC34A7D7348552F256D738820EE7AB314CA4C6509CC61063B27118BE271AEFBBFDD4910DE6C49A660218B157311A438CC9C66148A531C4986C1C86CC1A874D360E438E1AC3F74463BFC49D4E95AF16256F558DC3380CC66188D1188AD3188AD3188CC6509CC670AD684C897B62E813733E4D339194BC95350E71B8088D414D360E83711852690C31261B8721B3C661938DC39059E32967F2FEFB868FB49B33699A8967C45E45350E9B6C1C867135FEB0D1FB662842E32426E7A83114A731188DE15AD5B8ABEE74FF2A730E4D73B528792B6A1CC66188D1188CC610A9F1573D1EEDE0A2518133680CC5690C31261B8721B3C661938DC330B9359E2E71271B23F62AC661C8ACF1C78F397569972E781BEACBD1D89059E32426E7A83114A7314C6E8DBFFB64CE74893B0929792D35658CC650A5C6174E96EDA5F5378F89C961930BD518D464E33018872195C61063B2711852690C31268734CB9F7DB74F97B8939692D75CE3508793680C31268BBADD4BCADE6ADBDA68F36AC1680C31261B8721B3C661938DC390A3C6509CC660ACAB9EAEBA337D6F9BD3659A49C588BD4A2E1AAFAF77D9B269D4C09A5457AF3114A731188DA1388D6192697CBEE7B9E91277F253F296D794311A43668D8FB6968D35ADE3B1D1BCBA388D21C664E33064D6386CB27118326B9CC4E40234BE7C70F674893B5508D8AB188741350E9B6C1C866D4D6557C3EDC257A3F62AAA71129373D4188AD3188CC6509CC650BDC6D30B1EA71A51F62AC66188D118C4DE93FE2C51A5463D1C2C8F83A4D218D464E33018872195C61063B27118326B1C36D9380C99354E6EF2F43D7D539392F75E6D19A36E10E33054D278D7FCB2A5951AF53055B156C8F96A9CC4E41C3586E2348609D178FA9EBEA94BC05EC5A86B88D7F8FC57654B63DAE7EDE5F238E8B030911A43711A438CC9C66148A531C4986C1C860A1A5F3E307D4FDFD4A6E42DAB2D631C06E3ADC1680CBDCD653FC76D6D4DA3435C49348618938DC39059E32426E7A83114A73154D2B8EBF6D387A7173C4E794ADEBBB565AAD4B8B5316296887666A8BC116CD4C63ACA95A3C630911A43711A438CC9C66148A8F1EEBA733DD30B1EAF1102F62AAA712A9387DACB66061B966EA8603515B20AACC4986C1C86541A438CC9C661C8A27105938DC390A3C69058E3EFF6CF39F1D9B5331B74FCD8C0F050FBE9C195678FBEFCEDC093970ECFAEC4B923CFB1CF579F6DF872E89AAA144ADE3B35658CC3905CE36D73CA4E9AB6B5C985E59EA88C9A0A59C7BA8CC310A33114A73114A731188DA1388D4135EEBAEFE4C0B5B0E0F1C4E79FA02BAE7A7D77783D3FC986F88CCCF86F8E3FB528794B47ECAD46E3C85922A2314F61EF9A86E8D12CEA641538178DA1388D21C664E33064D6386CB27118926BDC71CB99BE295FE21233F1AD1A632B71B17FEE99C1A5535463ECADF5A989D0189268BC6751D9C660235B5E5D3F1A9C3BA26692D887BC5AC7AB0BD23889C9396A0CC5690C4663A8ACF1B94353BBC4452AD4BAD2779F51AE08D098686CDEC024A7E4BD5D5B2695C6200EAFAC8F8EAB64CBAAAE50293E139C75CE29AC31188DA1388D414D360E83711852690C31261B8721B3C6ED33AEEC79704A97B864C82E3DEEAD338E154EDF1D64E6E6CD4C5A4ADE5B23F666D3F8D30A35AD4666B5774B851594DBE79447ADABD1386CB27118326B9CC4E41C35866A34EEB8F364FF142E71CBDE1AA9269829E2B06FAF623486189351776D63D940D3DAE68CDAAB2070A571699E528173D718265263284E6388317947DD37DD4BCD079C0B949DC343ED805AE6A91C214F3E7BF4E5AB106F2BD177C724CFA54BDE9B017B15E330446A7CBCAB6C60B00DF7785A1B4350E05551770ED2A89C35441B8D21C664E330A4D218624C360E43668DC3261B87A13A8D2FECCFFF9E3EE20FD5A03DA77BEBE8C464B3732487FB07808D175E7C69DEBC279A57B4B0CDC63FFED33F5F3FE3EF83FC71E1BF7DB3E767F67B4D02FA36D7FFEEBFFEF397B3FEDDBCE152A9243FCED30B1742DBB6ED6C1F3838A1D54AC97BA3D6092C1887C1380CAAF1D60AB34404E4606A6D34AE94690787B832680CC5690CC5690C466348A3F1E55DB3878FE43C8D899CE30EF05E3A3C3B72FA9433588CAD6F68D053FCFEFB1F601B87653B68EFADB7DEB4F1D51A73F049C577FB672C9A57A36F58E0C759BDE67DFD31F547631B99D9D6DF4371F8F62AAA71D864E3309C8E5A44D5DF3A9A571B54E384A35C8A711852690CC5690C31261B87218BC6154C46DDF63B4EF56E349F68F59C195C6A4EDF8AF4D699FA3068ECDD77DFA3A735273467B604289EFDD32BAF6EDBBE03BA3E5A71FEE05DF6B09392631D8D3BDBB7C8DB5EBF71D382A79E92484B2AA13FDACC9933F991DF78F34DB68317AC8228794B02F62A311A83D1182AD5C61F34BA605B4963D0A45A31EA06310E43728D77CEF1F62D1AD53889C9396A0CF96ABCED9673070AB9A72F85BA23BCBF7C21A2068D7DE491DFB0BD7E5D2BC42493983F89AADC045CE9BB2F7EB5D6AEDD9DFC1EE447E6F740D3D403C9C5F01CF1ED558CC39045E380C91FFAD9B5986C1C86CC1A83D118E235EE6BCE6771B571188CC3904A638831D9380C5B6EB8B8E7C98266715DC21C3A652B7171DF8F4F75FC888D5FFFEA3A4E53BCE5089CA0181B3C6625DCEA8BD031A700BD75096B7E1C465D9267368226C75CCED252F25EAF75041D1672D1586A63B137C8C4681C34B977ACBD4A668D93989CA3C6B0F9862B1FCD3E7938E7EB7790E48B99BA565CFFB3DABFF8F5ACBF627BFB5BD73DF8AFB7249456B8FA7342D5916A3E0981C9A5B9BAA9C984E25C1C2E79AFF9F62AC661301A836ABC798EB77789D7BEC06ABCBAD13BD4ECFAB7F8F68AC66B1A47ED6DA977619947D4DDD8E4752E72ACF0BF44E095F56ECE898D6D73BC034BBCCD4D63345E5DEFC6A8A9933BE68F8E75B14D4F6BA3D5786B93F3561EB157658689D418D464E330188721AC71DB9DA7BB8B5DF0E8F2D8D0695A89F52FFD8DC41309BF907C3269AAAB2BA41258A1FEE79746799CCB80966FAF924AE363FE74D1B99121A8558DCEDEE50DDE517F52F73B7F66481EF179B8C78D726952BDDFFF7BB154C5DDFE10F4A50BE599241C4660C9B765544C86B8BEE82A0B8CC9B2B3ACDC426304963DE508F4AF6970EA62ACEC23FDF2926034568CC61063B27118326B1C36D9380CBEC0E7F6BDFCC567C58E5E42C4E4500809B99F7FF8776CBFFAE40F7AD7FDAD3E45266C0E18C9B5A1AE904D60E2B0A4D0984C40AE260897BC576B1D41878571355ED9E860635983D383204C3426E4D2762D7226AF6F2ABB87BDF4D0E891E08C6FF88CC684DF0D4DE5804CE7F1AE517BD941269F649289804C34E680ECB6AADE99CC231CF42F04C45E22B02CE71A6875DB088FB1B2908B47B557310E837118623486E23406DFDEEF3A1E3F3178C87C6605614ECD3014BAA84BE8F8D5CFFFD23C05970ECF36070C7375D53DD53163C73B375CDC3FC3F45743368161D7EE4E3F77298FF06563C45EC5380CAA71D864BCDD34C7F1F590F7658F339947A2B1E4D5082C324B524D1CEE6B2DE7D5347C665BCB6318F483361A8BBDC467D40559D145A76C109F83E5314EA23439B6C0363DD84B4360C9AB4566FA35A986E2348618938DC310A5F1FFB5FDCBC9BE024B5C0379AF392F0D126689BDA82BB1378C39A6214775F1B079714D5DDD4D3FBFFDA6BDAB12D988B7324FCBAB16CDAB3DD6969BC399EF19D66218528D1A2825EFE55A8771188CC310D4F89D06EFCB5E67088D0DCC547B65031078FB02B783D84B468D5AA234FB3737387BF1591A115502350159ECE5516232C897DA8FD2082CE0A4696A6F700E996DE90F0A2CA8C6494CCE516388D478C3ADA73F99E87BFAE2479B9F7DFCAF8912DBDFBACEF41BCC3183A42AAA63E86B9D817BBA6402FEE39733CD3E913CF9DB31AF92176E78358F50DC5B97596081F0CBAF57D2E95494BC3FD53AC4E148938DC380BD3DBE751D8BCB4198788BC36CF0481C167BA1D7DF4D42F13A3FAD6D5FE02CC56434669B260119A460D6916A1E8DBDE4D83433814C82CD01C5640DC840E38C547B5157EC558CC3904A632840E30BBB9F3B5E7C891B26C6DE531D3F92841987CD5306734C8593DBEC990182E7A34D3706F5D32F51DAEC6C2056CB9E44ECD77FEF82B67C29703918F708F15CE9BB2FF31DC2D4BDB2C663DEBC27CC53E33262AF12A331A8C003BE69128737CE75DB62EF11BF7FC5BDE5BC1A93692A33924BCFE639CE5E2D86D95EDE303AC42523D5D8ABA35C3409C594AF44E90F028B31A52ADE39DFD3684C828DBD12CCA5425EE7A7DCC65EC5380C99354E6272058D2F6F7BF0C4D1092A71C370F29933529021652ADEE5CFFE30D81F416F9D39A6E08E5CC55DF5C446222419B2CAC636124ACC947EE26AF02561D85F5EAB91967CDB446339ACBE242DD405E6074F0E35B02CAE942C3AF8543C25EFA55A877118E2357EDF0F83A828364AECC5E4967BCB3D924E8B906A6F973FC2A485F17BBEB13846B24D93CC1993C55E1E252603CDC95CE3ED1EF94B00845C4CEE6E76F34C884AE3E574D2F019873F9A5F3E20D2D264D40A938DBA8689D418C4E18D779DEC9DB812B712E1654FD4B71272B1D73C15E662FF5C734021C95076254C9C24D725020777406C798AE81AEC3788E4C458D32FF573F0D2504D10AEF286241CA60C96452F0919B157310E43258DD73479075A1C1BE63A3E5A3C9A57138DE9C7D50F9ABC832DE5EC1A88C9079BDD2897FABCB2D10D5653097FBCC8CD39B181BDB2C1A3DACB97ABFD2F25AF465AC064823001B9A5DE0D71D183B76CE86D1238DCB9C875F2B8B9C94D2F49522D186F0D46638831D9380C09355E7BEB37FB0AB9A72F03E151252977499B755237863383113F0827B4D92D3988A452112723BD22968ADE613315C2AC1C24C64C1DD332578774F4D665CE9F81BA97DF362D79015CF2FE38D6DE20466388D458D0A45A91BC5A508115155891800C086C508D1519AC96BC5ACB6345F36A45F36A083A9CAFC69058E36F77CEBF2A256E25C2A52F21F7D5277FD0B5E27AD31F41D489EB72E6EA96318F6B26109FD9873D2B8D3F49864C8035FD4124B58E3948422A252009217F4E35F8ECDBAB187B15E330A8C661938DC39059E32426E7A83118758318872195C630A2F1E5AD0F7ED99FF33D7DB970E9F0ECF259B8EFC704DE242157387B34E29689EA57326BBD1A9318EB881479F5A34D3706336DE9940D726CF3C2209552EB0C245C051D43DBB6ED09EF4C2A792FD63A820EE7A531188DA1388D414D360E837118326B0C99355E77E7A943F9DFD397176ED6D78F96CB9FFDA16471C1F3B2129123AEE34E2027814828EEC58F27856783C2C404D524A97572F86D985F452A347F4EB206ABE4BDE0DBAB188DC1082CB42FF686BABCF573A335A6198DFFDCEB2699DE6E18A331E571A4C9A26E6FAB7B5CDBE44AE2A0C6EB9BDC22CD2DFEA835D0D4DE5D8BDCB33CA60AC8396A0C95345E73EBD9CE572655AA1C899BDDE9ADC35ECADD71A7881C15663BAB19AC0A423CC4ABF89C96F0CB6E045E240724142862B5C7BC244892D43A15991760C1E1FE81993367D6373424B99DD0B75749AEF1BE16AFBFDDEBDDE29D1A72B347FB5BDC48F2AE25EE1151693B17BBC1E70F1794ED659BDD3ADF701B6D0BDC6EFD6DEEF1AB41D7F341937BFCBCCB8D54CB8CF1C196F2C6CE856E168A3DD99668BC7789137BD3A3E59521B443CD4EE6AF06DC0876CF1AC7704F99D343DEF627DCE0F3BE25D661301A43111AB7D45C6C9F7F156783D22202933C8F3BD44C9C89BC39217EED472A3431AE6A3CA9321ADEE353EB74F4DD617E21A9D8B53B695555F2FEB7D611743889C6D8DB8D424F3861D0186309C53D9B9CA2F84CFBB4D5F57CFC8AB7BCD14563F4C3643A91F0408BDBD8C0B5B9D5BD1099D9E038588AB14B1B9CC6ECC6F6A6B95ED71BAE9F2F61E3A34EE0958DAEC7EDECAFCA147B416446E3A3ED2E387340EC6DE7E0BECC74B6364D7C797C79D343C33D577F3628158F3CF29B854F3FF579E7E3F6A40CD25B47AD5B698835DFF5CCB22A23C7D81824E19C535AAA09BF40E0E55390BFD111C388BD8AD11822355E37D77BB7D1694CBCC5E42D0BDCE3E6DF391BD9A6BDD7E802F2B685DED6052EAF2614E370EB5CD729334C04E43D4BBC3D6F39BD65C209AB65DA898D8EC58E967BBD0DBF75026372D712D7B3E25E37E174B0D965D4F4C84ACC754DCE5E126C363E5AE80496FB1649A15735BA1E09CEC195D513A0F1AADBBEDE3FF5FE8981DC804AA30023AE9E195CEA86B274E8B8EF0E52624ECD98A9915C2ADE205A97265CD29C0ACDCCF38DED5556BFB27672DC1B184ADEE25A8771188CC3A01A1B9341D2E9658DAE12966D41AB62456B6390A43A8856C50A55310E4B851C591E070996C73AB8A56879ACC4986C1C86341A9FDB35054ADC48B077DEBC27A8BB522DFA0992EF1FCD2024EAB8712EA34A068EA9C3D4B98C392BD50C3E73E9BCFBEE7BC61D792E798B7C7B15E330188721466308DAAB18872183C68AD1188AD318D268FCDDD6C74F1C9932256E118C06EAAA39D656945A41F82E7A8188F92EEC469A9D7C42B89AB59309F1ED0D1263B2711832680C4663284E63284E6308987C65EDACE14F474BDCA39F0DE9F614826CEDE991BF4E9C81548BAB88AB31E134A86E4143560A4EEAB45358603276BD23227EF47B0C15167E2784F0CB6711BFEEAAE4FDA1D6611C86188DA1688DC3261B8721B3C6494C4EA571F36D673ADF09FE5ABB3FEDEDDCB337D8335590A237DBEDA6907CBC0A07444E1E5FFF7D8DB9E1569F85A2D555A406061198F740B00DAE821692BF9F6AC6AE9294BE23F62AC66148A531C4986CEC558CC310A331188DA1388D414D360EC3BBB5E7773C6F4A5C82EECE8F77EFE8D885C3C1FEC98F14BDF7DFFF40F2490B43F2B419078C1572138244364D6513A932F00B6FF835EF6C9777E9A41BC50CB68B43DE99EDDE893F7887FFC1BE2A8A6004D6CB07E030EF24EDE87735C93351970F62BCD8FBFC8DDEF3B50EA33114A731188105E33064D6386CB27118326BEC9B7C69FD439125EE814FBAB177CFDEFD3C62B279F61A2655DA2C7EE283F1846D5577FC39D8A147BDF3A13FCF10D9AE5C701AE3B939C258B876E87717F059D3FBD4ABB2AA4B9EC7A5E43D77A3C3399C5863285A63284E63A852E395779DEC8EAE0C297789BA87BA3F65037B31D9EC3099E14AFFC28B2F65FEF3FFC9479BC36B3034B22968A3FB47808709BD35EDEBF5F1715846B9B9887059094F024B22CD5B35FD95C8FC6737C883F820E26FD91FB157C9AC71D864E330E4A231C4986C1C86541A438CC978BBECB6B33B0BF9270693014E176AAD54B79806B992F8DF64EB8D41A61F5B780A79505752E868C88489A5991BE9F4D107EC3193C155462E2E61B12389BC7F23096DDBB6FB431025D31FA4E43D5B5BA6388DC1680C456B0C796BFCEDD6F9C707A7E42C6E42A4D692BFF3900173E2C620A9B20C0EA586E0597D43FE4C02EB705AC2379FF99EC15DBB3BF92CC0F40709D8AB188D21C664E33064D6386CB2B157310E432A8D21A5C6973E78E8CBBE142950FFC091BDFB0E98CE3054C59D7BF64ED1B92543F2B5CD1ABECC387322A85D2BB7F3E7CF777777B7B5B5F158EE8A69590596C401E2B203256ACD73F7A7BDD456A6330325EF9990BD8A71185205E41C350623B0601C86CC1A5732F9BDBB4E77AD36BFB571F9B8730FD52F0E9B7E031F21BBED3F90C33FC5C805B9B3347E9CB312C9FF7D998CEB921E9BFEF121618E6D0F3FFCB0649BCB962D939E59B3663DF3CC33C3C3C3F2A56D083CDE3856182DDAE3EF5B54CCAA529D9288BF6A4BDD1B3F06E1DB1BC4082C18872195C61063B2711832680C4663A852E3A5B79DDDF1728654992B2B9F0D411587CD534164588BDD92783E3148DD4B33FD4970FFF63E74E286D1533FF53D3DC4C9A85A77707070DEBC79B2BD76ED5A79FF622FD2CA973476937D6CA30636DF2801A4CDFC0809A78ECC92492ED67CEE9C1B7BF6EE0FF61B92D5BD0B6B1DC66130022BC5690C456B1C36D9380CBEC0DF6E5EF0E7FE2C0B1EE5CACA278490681933E54B6ACD9E923C4F92A51D9C3177DF7D4F7CAD5509FDBB1CF110AFC4DE44696790A81166926439C5F1962FC99C8361962F5F79E5959B6FBE59F58E6E279BEDF78A850B900E8F27993A0AAED9183832C85941CED5D37738FEAA2D756F7D4383E90F52F29EF6ED558CC360EC558CC6A0261B87C1380CA9348618938DBD8A7118623406DFDE2B2DB34E1ECA7E4F1F36F291F0C8361F15049F0DC26E923BF13849626F3524B457675CD2D97BECC9B269631B8931EAE267676767B92BAAA1B13C925A470461427AB2B51CB81A9CD60A8F9947121C760E7ED67A0264C6B757C95DE32426E7A831188105E330446AFCC66D5F77A62E71AF31D6AF6B257FA6E832FDE392709595AE859009D5A4035715A6761132AEAC1DDBA42A467891794C1B7ECD7EC7005C6876BC3366C9246F9E24A2D04923F920E217BD95BCFFA975041D16624C36F62AC66148A53114AD3154D0F8EC765BE21236132E750C0F3B714D4DF8DA039F744BA08EA7EF70FFC4C4673F0F1DFDA3A4BCB7840B4ECC295B0913BE009F71232E140FFCA2EC58A0912A474818DB34CD66A3DCA5EDD249FB4D47405199221270789C771BE2E8B6DBCDEF8A8F920FD4741A489B79AB8F3E36C7F48F72ECF8FF03A7C26E60E34A36DF0000000049454E44AE426082
	EXEC dbo.spadmin_writepicture '6F82D2D6-5837-4D70-B27D-955232B3D942', 'arrows.png', 5, @newMobileFooterID OUTPUT, 0x89504E470D0A1A0A0000000D4948445200000140000000330802000000693346260000000467414D410000B18F0BFC6105000000097048597300000EC300000EC301C76FA86400000C0749444154785EED9C8D6E1CB91184F3FE2F76872827C191E1B32FFECB5D2E798D7CBA124AAD26D9D3D2AC602C40E083B0DBD364F5B08B4BEF2CE0BFFDF9DFFF6D369B6BE4FEFEFED906FEE33F7FC6B705FF2675088E5C36ADCFC575494C9129CDB426CC7659DD8BAF73BF42E8675EB67D17BFEB8BEB92982253C6B4BC81EFDEDDFFFAF1538C4CF9EDF317325370841BB8B9BDEBDC06B331670ABE1A147FBEF9A5B328B777EF3ABAACC9BB7FBE4FC1916FDF7F6FEA3641F7FEFD87141CF9FAED3BEB7CA84B026924A7F8199A158AEDAE2967DCF56C03D3DA9FFEFE8FC3E2B84A0E9987C5D15AD20E8BFBFCE56B47B70F0D63C24363513F6987D657C3C83CB4BE74DF7FF898E2AF035D66EBE86214D20EF70609A4919CE2AFA65F216C774D39E9AE671B1831320E8B930FA02E4E0D134571CC60DD8B58BFAFAB8583DAFA320AD4D69751045D49575F811A06BC48972232CAA1AE8D021F3FFD2B5D7D1DCD0AC576D79493EE7ADAC05E38B12A8EC131AD284EC782288A4BBAE7ADEF8641612CDAE93458E9C686C1CAFAD128D031744DD25D1D47D128501C47360A30848129E1A5342B14DB5D29419C77D7E306B60F9CB72A8EB812BC82D3E27C2C78C269718C95AE673B697D37CC134E8D65A3B8BC95F5350F452A7365FD51978FCC94F3222CA717FC9DEADA28D69DEE0D1BC569E78FA354E16A6560BBEBEDDCF5B881FDF14CA79D3D16671FD07EBF1E8B8B0DE36D511C633509B3D98BAFB67ED4E575A16BA37CFBFE7BBCF79466A3B026FE37CC687D1B8569D3BDBF0EB7005FC61A529A8D72A86BA33024DE7B4AEB43619AA4AE506C77C11BB9EB610323A3C15AACA2B8145F1597162B3626A639CE3CBC1D8B7B29A99E95B1523D2BDD31BEB27E340A6F0F0D5D53E8D24BA7C5B8EAF1DE48C751AA276EFB98D6A75F216C774D75C7F8EBDCF5B081C7D59F1637FA635ADCD41FD306FB5820A6C819EB27A388A96E320AF8D628C069C92830BDB56414114F3C079B8CC6F5AD4589641431DE1A2FC61E8DB7F6228A0AE3E2EBEA76D79BBAEBE921D666B3B93AF606DE6CAE98BD81379B2B666FE0CDE68AC91B98EFD9E921D814BE3DA7479D2BE297EF02668B4F02CE131FED143475BF7EFB1E9F40AC60A6A6EEAF1F3FF9214701BA9DE74C17D7EDD3AC506C774D39E3AE671B988C9F6F7E890FC156E8F9DB61719445DA613398A7A9DB0445740F17055DD2E273C515B777EFE273C5157ADE7868683DD2EC98EFE6F60E9ABAF820C5132490D6345F936685B0DDB5E28CBB9E6D60FF3C5017271F1C16A7869176581CF368C2FE677981750F8D25A3406D7D19056AEBCB28D24D97124D5D1905E85CBA149151001FA44B09FDFC00877BA349B342B1DD35E5A4BB9E36B07FE95252511C579D591447414E2B8A73C3E0D0FA1D6C14288C15750BEBBB61A2B0BE8D0285AE1B06C5B1D0D7B551A0388E9ABA7DFA15C276D794FE1AAEDCF5B481A30F525284784C5B154729310D565F7E7C2C8862513A44A388D5A244A3C0CAFAD128B0B27E6C18D095A92EB1A4BB3A8EA25160751C8DBA48A41C20188D02E78FA3668562BB2B2588F3EE7ADCC0FE78E60EBDD66371F601F579ADA7C579128CA217D3E2A2AED67A65FD26A3EED458B178ADF5D4FA360AD37AADA7D677F1BE2306A41C68EADA2814EF7D4217531A7812DFEFF43872F19447322F1898725E44BF42D8EE9A76F922EE7AD8C04C2D19FEF23AF6C6533C663F6FD5AAB864A655718C72C378EB8F1917F75262C378BB3296755917DE5A77B4BE1BC69AA46A23F44969D265DDF436E9D230CD20B7A56A23D66548A11B8DC2DB586D4C8B46E16DA1DBA7592170553EE12FAFB7BB4CEC57AA3652BBEB6103DB0794A831D3E25C8A977E555C5AFA5571E3D2AFACDF0115E9F2B7D61DEF6E6AFD6414985A7F54491BD58C7717B781D3461537281E47D69551206D5433DEDDC9E368ACD0264B15F277BBEB4DDDF5B881212DBD825E744052C1D8754629E8085715F1C20137A3601CAB48D49D8E6DE2B1B13DD39A1599EAC6B1DCBB820C7550913896AE28D8D48D4B3A1D5BE8C6B15ED2640B053D96178A1CEAF6E957E8F87617C4B1C51ABEC85D4F0FB1369BCDD5B137F06673C5EC0DBCD95C317903F36FEEF88FF282F8D5ABE0B269D4F6A374E3F79C821FA51BBF23153475FBF42B04D4C94DC129975DC6FE6AFF28DD57BB2B6FE0BBDE7FBD4D8E1F7E1650D64FB35FFC4698ADA99B1EB14EE13E9BBA37B777F199C18AA62EBBC84F470B4840373EC25971FFFE839FCA16F475497BC5239C82668562BB6BCA19773DDBC0588ACA0E8B930F483B2C8EB2483B2C4EBA8716A42AE91E5A5FBA87C662E19ABAA4C1A1F5F5F30C9E4EF18474E96EAD2BA3C0E1E9CA54A4A59F5B46F4CB0AC929FE6AFA15C276D79493EE7ADAC0C8C8075017E75FBAEAE268AAD2A028CE0D83DAFA54A5B4DAFA360A14C68ABAB5F5D530A8ADDFD4B551801EA7AB91A6AE8C22C67F62191B056ADD3ECD0A61BB6BC549773D6DE0E8839414F187AE288A73C3A0284EC78259E9C686C1CA82D12850E8DA286265FDBEAE1B06742525181B45AC8EA3D830581D477D5D1B0518B2D2EDD3AC506C77A50471DE5D8F1BD81FCF5C53EAAA38FB40073A4C8B73C39C3F2DCEBACC26DD585C4472E458776A41EB3A6D6A2C1BC5E5AD74D530EBF262AA6BA37842FC9D72C00D73DAF438B25190F30B82290D6C144F383D8E46DDFA383AA45F216C77AD74BD742EE0A5EE7ADCC03E16E8B4AB1C8BF3872E39B12B298D1AB45E94C5DBA23817C425EB8ED6A7125D22C7AB335A3F9554182BEA7A7546EBC7A5F01E18AD1F75D12A74BD145CF2DE63E6941697C2BAE371E44B2C05136ACD511F756349B1D729ADCFB4C2D587C27617BC91BB1E36F038785A9C07F357F15571C99AABE218A5B8AC391627784D8438571549F39B64CD956E340A6FE3FC535D190556D68F0DE3ADEF4BF39B6814DE8EF7259251C0F3A7E328E9A6FB32D128BC2539DDD74BF10C8715C2D8855597157717B6BB44EA7272D7C306F6679805A6C5D907FE0C8BC5396DFA19E645890D1E17746AFDB191B56EAC39DDBC906E5C502F7A3C16C6464EAD9F8C22A6BA63CDDED25E5298EA2A129774BA5DDD4AEB7AAD62CD632B5F44BF42FE6E77B9536FE1AEA787589BCDE6EAD81B78B3B962F606DE6CAE98BD81379B2B266F60BEDFFB5B7501DFE9E3B7EA02BE70A7C81466EBE892139F40ACF8E3AF1F4B5270CAFDFB0FF109C48A8BEB92D6D1FDFCE56B7CF2B1E2E2BABF7DFED27CBED5AC506C774D39A3FB6C0323F6F35FBF293BB242CFC1E243BF299445DA6171E892D669861E691E2E8A740F8DA5478B1DDD9BDB3B56E65017A330E1A1F56918698726404DBABC4897127AA4E9E7A82BA41B9FA34E410ED18E6EBF42D8EE5A71C65DCF36309622E3B0383F16AF8BA31ECA22EDB038350C6AEB5BB7B6BE8C22DDDA58D6ADADAF86416D7D19050E756514A88F05EBD6C7918C02F8205D4A90A0CC5A574681C3BDD1AC506C774D39E9AEA70DAC8FE7695284B87D0045713A1644511C3338ADD6D5C289C282360A14C672C3A0B07E5FD70D83C2FAD8D76985AE8D228AE3A8A96BA340B1376C1451EC8D7E85B0DD35A5AFBBEAF2D306D6B1E0E956C5451FC0AAB8D830B12A2E360C5616F4B1A00AFBBA536325A3C04AD74691EECAFA368A17706AFDD430581D47D128B0D2B551342D7F6B5DABAF8EA3681428B67AB342B1DD9572C479773D6E601F0B28B98563710CD325726C9D69719E847B96EAF4CB8F1B46011E325A30EAC621290D6C148AF79094039E84BBF09051D70DC3AC71484A03358C493C64AA6BA3D0120F4136A579121A6CDDF138B251F8EB1E4F8F231BA5D6F524D4E9211493D2A059A1D8EE62C81BB9EB610333B356813C42F17E9464FCA1ABA6AE8A4BABBF2ACEBA5AFD625188E8928456168C46E1AD7593B1D20D462FC63488BA68AD742DA4AEAFACDFD7F5C222045A25D457BA5AD895AE85681F6FBD4BB54A9178835137A541B34248F3A445886C77F17AD4B5D0D45D0F1BD8C7821B3F2D6E2C655A9C17CEED5C15371A6E8C800DE77F771DEA8E91A86BA3F81098EA26A3C0D4FAACBE546414182B11C92831128FA36414182B816414455225A2D07504DC7419055C8923297E58217FB7BBA6BA9772D7C306665EA10C2529128DC56B0563B90E3AE2B1D11CBC56308E55041CE9EB2A028EBC54374A4CC736752DD11CDBD775041489639BBA96E88F7504148963A199A9C9158FC9D3923CA72B8F41473C3656EEDB8963150147FABA8A80232FD58D12D3B14D5D4BACC63E3DC4DA6C3657C7DEC09BCD15B337F06673C5DCDFDFFF1FFF5CD8DF1565DE5B0000000049454E44AE426082
	
	IF EXISTS(SELECT * FROM ASRSysPictures p  
		INNER JOIN tbsys_mobileformlayout m ON m.HeaderLogoID = p.PictureID
		WHERE (p.Name = 'absHeader.png' OR p.name = 'AdvBS_WHITE.png'))
	BEGIN
		UPDATE m
		SET HeaderBackColor = 16777215,
			HeaderPictureLocation = 3,
			HeaderLogoID = @newMobileHeaderID,
			HeaderLogoWidth = 320,
			HeaderLogoHeight = 65,
			HeaderLogoHorizontalOffset = 0,
			HeaderLogoVerticalOffset = 0,
			HeaderLogoHorizontalOffsetBehaviour = 0,
			HeaderLogoVerticalOffsetBehaviour = 0,
			FooterBackColor = 16777215,
			FooterPictureID = @newMobileFooterID,
			FooterPictureLocation = 3
		FROM tbsys_mobileformlayout m
		INNER JOIN ASRSysPictures p ON m.HeaderLogoID = p.PictureID
		WHERE (p.Name = 'absHeader.png' OR p.name = 'AdvBS_WHITE.png')
	END; 

	
/* ------------------------------------------------------- */
PRINT 'Step - Updating Support Information'
/* ------------------------------------------------------- */
   IF (SELECT COUNT(SettingValue) FROM ASRSysSystemSettings WHERE Section = 'support' AND SettingKey = 'email') = 1
			UPDATE ASRSysSystemSettings SET SettingValue = 'ohrsupport@oneadvanced.com' WHERE Section = 'support' AND SettingKey = 'email';
	ELSE
	   	INSERT INTO ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES ('support','email','ohrsupport@oneadvanced.com');

   IF (SELECT COUNT(SettingValue) FROM ASRSysSystemSettings WHERE Section = 'support' AND SettingKey = 'webpage') = 1
			UPDATE ASRSysSystemSettings SET SettingValue = 'https://customers.oneadvanced.com' WHERE Section = 'support' AND SettingKey = 'webpage';
	ELSE
	   	INSERT INTO ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES ('support','webpage','https://customers.oneadvanced.com');

/* ------------------------------------------------------- */
PRINT 'Step - Updating Module Setup'
/* ------------------------------------------------------- */
   IF NOT EXISTS ( SELECT 1 FROM ASRSysModuleSetup WHERE ModuleKey = 'MODULE_HIERARCHY' AND ParameterKey = 'Param_DisableSimpleChart' )
        BEGIN
		   INSERT INTO ASRSysModuleSetup (ModuleKey, ParameterKey, ParameterValue, ParameterType)
		   VALUES ('MODULE_HIERARCHY', 'Param_DisableSimpleChart', 0, 'PType_Other')
	    END

/* ------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'
/* ------------------------------------------------------- */

	EXEC spsys_setsystemsetting 'database', 'version', '8.3';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.3.1';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.3.1';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'outlook service 2', 'minimum version', '1.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.3')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.3 Of OpenHR'