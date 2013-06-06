DROP PROCEDURE [dbo].[spASRIntGetSelfServiceRecordID]
DROP PROCEDURE [dbo].[spASRIntGetActualUserDetails]
DROP PROCEDURE [dbo].[sp_ASRIntGetPersonnelParameters]
DROP PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run]
DROP PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
DROP PROCEDURE [dbo].[sp_ASRIntCheckLogin]
DROP PROCEDURE [dbo].[sp_ASRUniqueObjectName]

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_ASRIntGetUserGroup]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[sp_ASRIntGetUserGroup]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRIntGetUserGroup]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spASRIntGetUserGroup]
GO

CREATE PROCEDURE [dbo].[spASRIntGetSelfServiceRecordID] (
	@piRecordID		integer 		OUTPUT,
	@piRecordCount	integer 		OUTPUT,
	@piViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sViewName		sysname,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@iRecordID			integer,
		@iRecordCount		integer, 
		@fSysSecMgr			bit,
		@fAccessGranted		bit;
		
	SET @iRecordID = 0;
	SET @iRecordCount = 0;

	SELECT @sViewName = viewName
		FROM ASRSysViews
		WHERE viewID = @piViewID;

	IF len(@sViewName) > 0
	BEGIN
		/* Check if the user has permission to read the Self-service view. */
		exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;

		IF @fSysSecMgr = 1
		BEGIN
			SET @fAccessGranted = 1;
		END
		ELSE
		BEGIN
		
			SELECT @fAccessGranted =
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM #sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action = 193 
				AND syscolumns.name = 'ID'
				AND sysobjects.name = @sViewName
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	
		IF @fAccessGranted = 1
		BEGIN
			SET @sCommand = 'SELECT @iValue = COUNT(ID)' + 
				' FROM ' + @sViewName;
			SET @sParamDefinition = N'@iValue integer OUTPUT';
			EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT;

			IF @iRecordCount = 1 
			BEGIN
				SET @sCommand = 'SELECT @iValue = ' + @sViewName + '.ID ' + 
					' FROM ' + @sViewName;
				SET @sParamDefinition = N'@iValue integer OUTPUT';
				EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordID OUTPUT;
			END
		END
	END

	SET @piRecordID = @iRecordID;
	SET @piRecordCount = @iRecordCount;
END
GO

CREATE PROCEDURE [dbo].[spASRIntGetUserGroup]
	( 
	@psItemKey				varchar(50),
	@psUserGroup			varchar(250)	OUTPUT,
	@iSelfServiceUserType	integer			OUTPUT,
	@fSelfService			bit				OUTPUT
	)
AS
BEGIN

	DECLARE @sPermissionItemKey varchar(500),
		@iSSIntranetCount AS integer,
		@sIntranet_SelfService AS varchar(255),
		@sIntranet AS varchar(255);
	
	set @psUserGroup = '';
	/* SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements. */
	SET NOCOUNT ON;
	SET @psUserGroup = (SELECT CASE 
		WHEN (usg.uid IS null) THEN null
		ELSE usg.name
	END as groupname
	FROM sysusers usu 
	LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
	LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
	WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
		AND (usg.issqlrole = 1 OR usg.uid IS null)
		AND lo.loginname = SYSTEM_USER
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END IN (
				SELECT [groupName]
				FROM [dbo].[ASRSysGroupPermissions]
				WHERE itemID IN (
					SELECT [itemID]
					FROM [dbo].[ASRSysPermissionItems]
					WHERE categoryID = 1
					AND itemKey LIKE '%' + @psItemKey + '%'
				)  
				AND [permitted] = 1))
END

	SET @sIntranet = (SELECT itemKey FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'INTRANET');
	
	SET @sIntranet_SelfService = (SELECT itemKey FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'INTRANET_SELFSERVICE');
	
	SET @iSSIntranetCount = (SELECT count(*) FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'SSINTRANET');
		
	If (@sIntranet is null) and (@sIntranet_SelfService is null) and (@iSSINTRANETcount = 0)
	/* No permissions at all  */
	BEGIN
		SET @sPermissionItemKey = 'NO PERMS'
		SET @iSelfServiceUserType = 0
		SET @fSelfService = 0
	END
	
	IF @sIntranet = 'INTRANET'
	/* IF DMI Multi automatically*/ 
	BEGIN
		SET @sPermissionItemKey = 'INTRANET'
		SET @iSelfServiceUserType = 1
		SET @fSelfService = 0
	END
	
	IF (@sIntranet_SelfService = 'INTRANET_SELFSERVICE') and (@iSSIntranetCount = 0)
	/* IF DMI Single Only*/ 
	BEGIN
		SET @sPermissionItemKey = 'INTRANET'
		SET @iSelfServiceUserType = 2
		SET @fSelfService = 0
	END	
	
	IF (@sIntranet_SelfService = 'INTRANET_SELFSERVICE') and (@iSSIntranetCount = 1)
	/* IF DMI Single And SSI */ 
	BEGIN
		SET @sPermissionItemKey = 'SSINTRANET'
		SET @iSelfServiceUserType = 3
		SET @fSelfService = 1
	END	
	
	IF  @iSSIntranetCount = 1 and (@sIntranet is null and  @sIntranet_SelfService is null)
	/* IF SSI Only */ 
	BEGIN
		SET @sPermissionItemKey = 'SSINTRANET'
		SET @iSelfServiceUserType = 4
		SET @fSelfService = 1
	END
GO

CREATE PROCEDURE [dbo].[spASRIntGetActualUserDetails]
(
		@psUserName sysname OUTPUT,
		@psUserGroup sysname OUTPUT,
		@piUserGroupID integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iFound				integer,
		@sSQLVersion			integer,
		@sProgramName			varchar(500),
		@sPermissionItemKey		varchar(500),
		@psItemKey				varchar(50),
		@iSelfServiceUserType	integer,
		@fSelfService			bit;

	SET @psUserGroup = '';
	SET @sPermissionItemKey = '';
	SET @sProgramName = '';
	SET @iSelfServiceUserType = 0;
	SET @psItemKey = 'INTRANET';
	
	EXEC	[dbo].[spASRIntGetUserGroup]
			@psItemKey = 'INTRANET',
			@psUserGroup = @psUserGroup OUTPUT,
			@iSelfServiceUserType = @iSelfServiceUserType OUTPUT,
			@fSelfService = @fSelfService OUTPUT

	IF @psUserGroup IS NULL
	BEGIN
		SET @sPermissionItemKey = 'NO PERMS'
	END

	SET @sSQLVersion = convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))

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
		END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'

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
				END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'
			AND CASE 
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
				END IN (
					SELECT [groupName]
					FROM [dbo].[ASRSysGroupPermissions]
					WHERE itemID IN (
									SELECT [itemID]
									FROM [dbo].[ASRSysPermissionItems]
									WHERE categoryID = 1
									AND itemKey LIKE @sPermissionItemKey + '%'
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
			END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'
			AND CASE 
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
				END IN (
					SELECT [groupName]
					FROM [dbo].[ASRSysGroupPermissions]
					WHERE itemID IN (
														SELECT [itemID]
														FROM [dbo].[ASRSysPermissionItems]
														WHERE categoryID = 1
														AND itemKey LIKE @sPermissionItemKey + '%'
													)  
					AND [permitted] = 1
	)
	END

	IF @psUserGroup <> ''
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
END


GO

CREATE PROCEDURE [dbo].[sp_ASRIntCheckLogin] (
	@piSuccessFlag			integer			OUTPUT,
	@psErrorMessage			varchar(MAX)	OUTPUT,
	@piMinPassordLength		integer			OUTPUT,
	@psIntranetAppVersion	varchar(50),
	@piPasswordLength		integer,
	@piUserType				integer			OUTPUT,
	@psUserGroup			varchar(250)	OUTPUT,
	@iSelfServiceUserType   integer			OUTPUT
)
AS
BEGIN
	/* Check that the current user is okay to login. */
	/* 	@pfLoginOK 	= 0 if the login was NOT okay
				= 1 if the login was okay (without warnings)
				= 2 if the login was okay (but the user's password has expired)
		@psErrorMessage is the description of the login failure if @pfLoginOK = 0
		@piMinPassordLength is the configured minimum password length
		@psIntranetAppVersion is the intranet application version passed into the stored procedure (set as a session variable in the global.asa file. 
		@piPasswordLength is the length of the user's current password. 
		@fDmiNetUserType is te DMI or SSI path 1 for SSi 0 for DMI

	*/
	SET NOCOUNT ON;
	
	DECLARE @iSysAdminRoles				integer,
		@sLockUser						sysname,
		@sHostName						varchar(MAX),
		@sLoginName						varchar(255),
		@sProgramName					varchar(255),
		@sRoleName						sysname,
		@source 						varchar(30),
		@desc 							varchar(200),
		@fIntranetEnabled				bit,
		@sIntranetDBVersion				varchar(50),
		@sIntranetDBMajor				varchar(50),
		@sIntranetDBMinor				varchar(50),
		@sIntranetDBRevision			varchar(50),
		@sIntranetAppMajor				varchar(50),
		@sIntranetAppMinor				varchar(50),
		@sIntranetAppRevision			varchar(50),
		@sMinIntranetVersion			varchar(50),
		@sMinIntranetMajor				varchar(50),
		@sMinIntranetMinor				varchar(50),
		@sMinIntranetRevision			varchar(50),
		@iPosition1 					integer,
		@iPosition2 					integer,
		@fValidIntranetAppVersion		bit,
		@fValidIntranetDBVersion		bit,
		@fValidMinIntranetVersion		bit,
		@iMinPasswordLength				integer,
		@iChangePasswordFrequency		integer,
		@sChangePasswordPeriod			varchar(1),
		@dtPasswordLastChanged			datetime,
		@fPasswordForceChange			bit,
		@sDomain						varchar(MAX),
		@iCount							integer,
		@iPriority						integer,
		@sDescription					varchar(MAX),
		@sLockTime						varchar(255),
		@sValue							varchar(MAX),
		@iValue							integer,
		@iFullUsers						integer,
		@iSSUsers						integer,
		@iSSIUsers						integer,
		@iTemp							integer,
		@fSelfService					bit, 
		@fValidSYSManagerVersion		bit,
		@sSYSManagerMajor				varchar(50),
		@sSYSManagerMinor				varchar(50),
		@sSYSManagerVersion				varchar(50),
		@sActualUserName				sysname, 
		@iActualUserGroupID				integer,
		@iFullIntItemID					integer,
		@iSSIntItemID					integer,
		@iSSIIntItemID					integer,
		@iSID							binary(85),
		@sSQLVersion					int,
		@sLockMessage					varchar(200),
		@fNewSettingFound				bit,
		@iCurrentItemKey				integer,
		@sCurrentItemKey				varchar(50),
		@psItemKey						varchar(50),
		@fOldSettingFound				bit;
		
	SET @piSuccessFlag = 1;
	SET @psErrorMessage = '';
	SET @piMinPassordLength = 0;
	SET @piUserType = 0;
	SET @iFullUsers = 0;
	SET @iSSUsers = 0;
	SET @iSSIUsers = 0;
	SET @fSelfService = 0;
	SET @iSelfServiceUserType = 0;
	SET @psUserGroup = '';
	
	SET @psItemKey = 'INTRANET';
	EXEC	[dbo].[spASRIntGetUserGroup]
			@psItemKey = 'INTRANET',
			@psUserGroup = @psUserGroup OUTPUT,
			@iSelfServiceUserType = @iSelfServiceUserType OUTPUT,
			@fSelfService = @fSelfService OUTPUT				        

	IF @psUserGroup IS NULL
	BEGIN
		SET @piSuccessFlag = 0
		SET @psErrorMessage = 'The user is not a member of any OpenHR user group or has no permissions to use this system.'
	END
		
	IF current_user = 'dbo'
	BEGIN
		SET @piSuccessFlag = 0;
		SET @psErrorMessage = 'SQL Server system administrators cannot use the intranet module.';
	END
	ELSE
	BEGIN
		/* Fault 3901 */
		SELECT @iSysAdminRoles = sysAdmin + securityAdmin + serverAdmin + setupAdmin + processAdmin + diskAdmin + dbCreator
		FROM master..syslogins
		WHERE name = system_user
		IF @iSysAdminRoles > 0 
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'Users assigned to fixed SQL Server roles cannot use the intranet module.'
		END
	END
	/* Check if anyone has locked the system. */
	IF @piSuccessFlag = 1
	BEGIN
		DECLARE @tmpSysProcess1 TABLE(
			hostname nvarchar(50), 
			loginname nvarchar(50),
			program_name nvarchar(50),
			hostprocess int, sid binary(86),
			login_time datetime,
			spid smallint,
			uid smallint);
			
		INSERT @tmpSysProcess1 EXEC dbo.spASRGetCurrentUsers;
		
		
		SELECT TOP 1 @iPriority = ASRSysLock.priority,
			@sLockUser = ASRSysLock.username,
			@sLockTime = convert(varchar(255), ASRSysLock.lock_time, 100),
			@sHostName = ASRSysLock.hostname,
			@sDescription = ASRSysLock.description
		FROM ASRSysLock
		LEFT OUTER JOIN @tmpSysProcess1 syspro 
			ON ASRSysLock.spid = syspro.spid AND ASRSysLock.login_time = syspro.login_time
		WHERE priority = 2 
			OR syspro.spid IS not null
		ORDER BY priority
		IF (NOT @iPriority IS NULL) AND (@iPriority <> 3)
		BEGIN
			/* Get the lock message set in the System Manager */
			SET @sLockMessage = ''
			EXEC sp_ASRIntGetSystemSetting 'messaging', 'lockmessage', 'lockmessage', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
			
			IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND LTRIM(RTRIM(@sValue)) <> ''
			BEGIN
				SET @sLockMessage = @sValue + '<BR><BR>'
			END
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The database has been locked.<P>' + 
				CASE @iPriority
				WHEN 2 THEN
					@sLockMessage 
				ELSE ''
				END
			    + 'User :  ' + @sLockUser + '<BR>' +
				  'Date/Time :  ' + @sLockTime +  '<BR>' +
				  'Machine :  ' + @sHostName +  '<BR>' +
				  'Type :  ' + @sDescription
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
	
		/* Get the current System Manager version */
		SET @sSYSManagerVersion = ''
		exec sp_ASRIntGetSystemSetting 'database', 'version', 'version', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @sSYSManagerVersion = @sValue
		END
		/* Get the intranet version. */
		SET @sIntranetDBVersion = ''
		IF @fSelfService = 0
		BEGIN
			exec dbo.sp_ASRIntGetSystemSetting 'intranet', 'version', 'intranetVersion', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		END
		ELSE
		BEGIN
			exec dbo.sp_ASRIntGetSystemSetting 'ssintranet', 'version', '', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		END
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @sIntranetDBVersion = @sValue
		END
		/* Get the minimum intranet version. */
		SET @sMinIntranetVersion = ''
		IF @fSelfService = 0
		BEGIN
			exec dbo.sp_ASRIntGetSystemSetting 'intranet', 'minimum version', 'minIntranetVersion', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		END
		ELSE
		BEGIN
			exec dbo.sp_ASRIntGetSystemSetting 'ssintranet', 'minimum version', '', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		END
		
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @sMinIntranetVersion = @sValue
		END
		/* Get the minimum password length. */
		SET @iMinPasswordLength = 0
		exec dbo.sp_ASRIntGetSystemSetting 'password', 'minimum length', 'minimumPasswordLength', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @iMinPasswordLength = convert(integer, @sValue)
		END
		SET @piMinPassordLength = @iMinPasswordLength
		/* Get the password change frequency. */
		SET @iChangePasswordFrequency = 0
		exec dbo.sp_ASRIntGetSystemSetting 'password', 'change frequency', 'changePasswordFrequency', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @iChangePasswordFrequency = convert(integer, @sValue)
		END
		/* Get the password change period. */
		SET @sChangePasswordPeriod = ''
		exec dbo.sp_ASRIntGetSystemSetting 'password', 'change period', 'changePasswordFrequency', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
		BEGIN
			SET @sChangePasswordPeriod = UPPER(@sValue)
		END
	END
	/* Check the database version is the right one for the application version. */
	IF @piSuccessFlag = 1
	BEGIN
		/* Extract the Intranet application version parts from the given version string. */	
		SET @fValidIntranetAppVersion = 1
		SET @iPosition1 = charindex('.', @psIntranetAppVersion)
		IF @iPosition1 = 0 SET @fValidIntranetAppVersion = 0
		IF @fValidIntranetAppVersion = 1
		BEGIN
			SET @iPosition2 = charindex('.', @psIntranetAppVersion, @iPosition1 + 1)
			IF @iPosition2 = 0 SET @fValidIntranetAppVersion = 0
		END
		IF @fValidIntranetAppVersion = 1
		BEGIN
			SET @sIntranetAppMajor = left(@psIntranetAppVersion, @iPosition1 - 1)
			SET @sIntranetAppMinor = substring(@psIntranetAppVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
			SET @sIntranetAppRevision = substring(@psIntranetAppVersion, @iPosition2 + 1, len(@psIntranetAppVersion) - @iPosition2)
		END
		ELSE
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'Invalid intranet application version.'
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		/* Extract the Intranet database version parts from the version string. */	
		SET @fValidIntranetDBVersion = 1
		SET @iPosition1 = charindex('.', @sIntranetDBVersion)
		IF @iPosition1 = 0 SET @fValidIntranetDBVersion = 0
		IF @fValidIntranetDBVersion = 1
		BEGIN
			SET @iPosition2 = charindex('.', @sIntranetDBVersion, @iPosition1 + 1)
			IF @iPosition2 = 0 SET @fValidIntranetDBVersion = 0
		END
		IF @fValidIntranetDBVersion = 1
		BEGIN
			SET @sIntranetDBMajor = left(@sIntranetDBVersion, @iPosition1 - 1)
			SET @sIntranetDBMinor = substring(@sIntranetDBVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
			SET @sIntranetDBRevision = substring(@sIntranetDBVersion, @iPosition2 + 1, len(@sIntranetDBVersion) - @iPosition2)
		END
		ELSE
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'Invalid intranet database version.'
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		/* Extract the Minimum Intranet version parts from the version string. */	
		SET @fValidMinIntranetVersion = 1
		SET @iPosition1 = charindex('.', @sMinIntranetVersion)
		IF @iPosition1 = 0 SET @fValidMinIntranetVersion = 0
		IF @fValidMinIntranetVersion = 1
		BEGIN
			SET @iPosition2 = charindex('.', @sMinIntranetVersion, @iPosition1 + 1)
			IF @iPosition2 = 0 SET @fValidMinIntranetVersion = 0
		END
		IF @fValidMinIntranetVersion = 1
		BEGIN
			SET @sMinIntranetMajor = left(@sMinIntranetVersion, @iPosition1 - 1)
			SET @sMinIntranetMinor = substring(@sMinIntranetVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
			SET @sMinIntranetRevision = substring(@sMinIntranetVersion, @iPosition2 + 1, len(@sMinIntranetVersion) - @iPosition2)
		END
	END
	
	/* Check the System Manager database version is the right one for the intranet version. */
	IF @piSuccessFlag = 1
	BEGIN
		/* Extract the System Manager database version parts from the given version string. */	
		SET @fValidSYSManagerVersion = 1
		SET @iPosition1 = charindex('.', @sSYSManagerVersion)
		IF @iPosition1 = 0 SET @fValidSYSManagerVersion = 0
		IF @fValidSYSManagerVersion = 1
		BEGIN
			SET @sSYSManagerMajor = left(@sSYSManagerVersion, @iPosition1 - 1)
			SET @sSYSManagerMinor = substring(@sSYSManagerVersion, @iPosition1 + 1, len(@sSYSManagerVersion) - @iPosition1)
		END
		ELSE
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'Invalid System Manager database version.'
		END
	END
	
	IF @piSuccessFlag = 1
	BEGIN
		/* Check the application version against the one for the current database. */
		IF (convert(integer, @sIntranetAppMajor) < convert(integer, @sIntranetDBMajor)) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) < convert(integer, @sIntranetDBMinor))) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sIntranetDBMinor)) AND (convert(integer, @sIntranetAppRevision) < convert(integer, @sIntranetDBRevision))) 
		BEGIN
			/* Application is too old for the database. */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The intranet application is out of date.' 
												+ '<BR>Please ask the System Administrator to update the intranet application.'
												+ '<BR><BR>'
												+ 'Database Name : ' + db_name()
												+ '<BR><BR>'
												+ 'OpenHR System Manager Version : ' + @sSYSManagerVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Database Version : ' + @sIntranetDBVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Application Version : ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision				
												
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		/* Check the application version against the one for the current database. */
		IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sIntranetDBMajor)) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sIntranetDBMinor))) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sIntranetDBMinor)) AND (convert(integer, @sIntranetAppRevision) > convert(integer, @sIntranetDBRevision))) 
		BEGIN
			/* Database is too old for the appplication. */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The database is out of date.' 
												+ '<BR>Please ask the System Administrator to update the database for use with version ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision + ' of the intranet.'
												+ '<BR><BR>'
												+ 'Database Name : ' +  db_name()
												+ '<BR><BR>'
												+ 'OpenHR System Manager Version : ' + @sSYSManagerVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Database Version : ' + @sIntranetDBVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Application Version : ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision				
												
			IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sSYSManagerMajor)) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sSYSManagerMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sSYSManagerMinor))) 
			BEGIN
				SET @psErrorMessage = @psErrorMessage + '<BR><BR>'
																							+ '<FONT COLOR="Red"><B>Please note that the System Manager version also requires updating to version ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.</B></FONT>' 
			END					
												
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		/* Check the application version against the one for the current database. */
		IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sSYSManagerMajor)) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sSYSManagerMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sSYSManagerMinor))) 
		BEGIN
			/* Database is too old for the appplication. */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The database is out of date.' 
				+ '<BR>Please ask the System Administrator to update the System Manager version to ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.'
				+ '<BR><BR>'
				+ 'Database Name : ' + db_name()
				+ '<BR><BR>'
				+ 'OpenHR System Manager Version : ' + @sSYSManagerVersion
				+ '<BR><BR>'
				+ 'OpenHR Intranet Database Version : ' + @sIntranetDBVersion
				+ '<BR><BR>'
				+ 'OpenHR Intranet Application Version : ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision				
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		IF (CONVERT(varchar, @@SERVERNAME) <> CONVERT(varchar, SERVERPROPERTY('servername')))
		BEGIN
			/* Microsoft SQL Server has been renamed */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The Microsoft SQL Server has been renamed but the operation is incomplete.' 
				+ '<BR>Please contact your System Administrator.'
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		EXEC sp_ASRIntGetSystemSetting 'platform', 'SQLServerVersion', 'SQLServerVersion', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		
		IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND SUBSTRING(LTRIM(RTRIM(@sValue)),1,1) <> @sSQLVersion
		BEGIN
			/* Microsoft SQL Version has been upgraded */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The Microsoft SQL Version has been upgraded.' 
				+ '<BR>Please ask the System Administrator to save the update in the System Manager.'
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		EXEC sp_ASRIntGetSystemSetting 'platform', 'DatabaseName', 'DatabaseName', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND UPPER(LTRIM(RTRIM(@sValue))) <> UPPER(DB_NAME())
		BEGIN
			/* The database name changed */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The database name has changed.' 
				+ '<BR>Please ask the System Administrator to save the update in the System Manager.'
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		EXEC sp_ASRIntGetSystemSetting 'platform', 'ServerName', 'ServerName', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
		
		IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1))
		BEGIN
			IF LTRIM(RTRIM(@sValue)) = '.' SELECT @sValue = @@SERVERNAME
			IF UPPER(@sValue) <> UPPER(@@SERVERNAME)
			BEGIN
				/* The database has moved to a different Microsoft SQL Server */
				SET @piSuccessFlag = 0
				SET @psErrorMessage = 'The database has moved to a different Microsoft SQL Server.' 
					+ '<BR>Please ask the System Administrator to save the update in the System Manager.'
			END
		END
	END
	IF @piSuccessFlag = 1
	BEGIN
		EXEC sp_ASRIntGetSystemSetting 'database', 'refreshstoredprocedures', 'refreshstoredprocedures', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
			
		IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND LTRIM(RTRIM(@sValue)) = 1
		BEGIN
			/* Database is too old for the appplication. */
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The database is out of date.' 
				+ '<BR>Please ask the System Administrator to save the update in the System Manager.'
				+ '<BR><BR>'
				+ 'Database Name : ' + db_name()
				+ '<BR><BR>'
				+ 'OpenHR System Manager Version : ' + @sSYSManagerVersion
				+ '<BR><BR>'
				+ 'OpenHR Intranet Database Version : ' + @sIntranetDBVersion
				+ '<BR><BR>'
				+ 'OpenHR Intranet Application Version : ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision	
		END
	END
	IF (@piSuccessFlag = 1) AND (@fValidMinIntranetVersion = 1)
	BEGIN
		/* Check the application version against the minimum one for the current database. */
		IF (convert(integer, @sIntranetAppMajor) < convert(integer, @sMinIntranetMajor)) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sMinIntranetMajor)) AND (convert(integer, @sIntranetAppMinor) < convert(integer, @sMinIntranetMinor))) 
			OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sMinIntranetMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sMinIntranetMinor)) AND (convert(integer, @sIntranetAppRevision) < convert(integer, @sMinIntranetRevision))) 
		BEGIN
			/* Application is older than the minimum required */
			SET @piSuccessFlag = 0
			--SET @psErrorMessage = 'The intranet application is out of date. You require version ' + @sMinIntranetVersion + ' or later. Contact your administrator to update it.'
			SET @psErrorMessage = 'The intranet application is out of date.' 
												+ '<BR>Please ask the System Administrator to update the intranet application.'
												+ '<BR><BR>'
												+ 'Database Name : ' + db_name()
												+ '<BR><BR>'
												+ 'OpenHR System Manager Version : ' + @sSYSManagerVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Database Version : ' + @sIntranetDBVersion
												+ '<BR><BR>'
												+ 'OpenHR Intranet Application Version : ' + @sIntranetAppMajor + '.' + @sIntranetAppMinor + '.' + @sIntranetAppRevision				
		END
	END
	-- Get licence details
	IF @piSuccessFlag = 1
	BEGIN
		EXEC dbo.spASRIntGetLicenceInfo @fSelfService, @piSuccessFlag OUTPUT,
					@fIntranetEnabled OUTPUT, @iSSUsers OUTPUT,
					@iFullUsers OUTPUT,	@iSSIUsers OUTPUT,
					@psErrorMessage OUTPUT
	END
	-- Check that the user belongs to a valid role in the selected database.
	IF @piSuccessFlag = 1
	BEGIN
		EXEC dbo.spASRIntGetActualUserDetails
			@sActualUserName OUTPUT,
			@sRoleName OUTPUT,
			@iActualUserGroupID OUTPUT					        

		IF @sRoleName IS NULL
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'The  user is not a member of any OpenHR user group.'
		END
	--END
	--IF @piSuccessFlag = 1
	--BEGIN
		/* Check that the user is permitted to use the Intranet module. */
		/* First check that this permission exists in the current version. */
		SELECT @iSSIIntItemID = ASRSysPermissionItems.itemID
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		WHERE ASRSysPermissionItems.itemKey = 'SSINTRANET'
			AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'
		IF @iSSIIntItemID IS NULL SET @iSSIIntItemID = 0
		IF @fSelfService = 1
		BEGIN
			/* The permission does exist in the current version so check if the user is granted this permission. */
			SELECT @iCount = count(ItemID)
			FROM ASRSysGroupPermissions 
			WHERE ASRSysGroupPermissions.itemID = @iSSIIntItemID
				AND ASRSysGroupPermissions.groupName = @sRoleName
				AND ASRSysGroupPermissions.permitted = 1
			IF @iCount = 0
			BEGIN				
				SET @piSuccessFlag = 0
				SET @psErrorMessage = 'You are not permitted to use the Self-service Intranet module with this user name.'
			END
		END
		ELSE
		BEGIN
			SELECT @iFullIntItemID = ASRSysPermissionItems.itemID
			FROM ASRSysPermissionItems
			INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
			WHERE ASRSysPermissionItems.itemKey = 'INTRANET'
				AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'
			IF @iFullIntItemID IS NULL SET @iFullIntItemID = 0
		
			SELECT @iSSIntItemID = ASRSysPermissionItems.itemID
			FROM ASRSysPermissionItems
			INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
			WHERE ASRSysPermissionItems.itemKey = 'INTRANET_SELFSERVICE'
				AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'
			IF @iSSIntItemID IS NULL SET @iSSIntItemID = 0
		
			IF @iFullIntItemID > 0
			BEGIN
				/* The permission does exist in the current version so check if the user is granted this permission. */
				SELECT @iCount = count(ItemID)
				FROM ASRSysGroupPermissions 
				WHERE ASRSysGroupPermissions.itemID = @iFullIntItemID
					AND ASRSysGroupPermissions.groupName = @sRoleName
					AND ASRSysGroupPermissions.permitted = 1
				
				IF @iCount = 0
				BEGIN
					IF @iSSIntItemID > 0
					BEGIN
						/* The permission does exist in the current version so check if the user is granted this permission. */
						SELECT @iCount = count(ItemID)
						FROM ASRSysGroupPermissions 
						WHERE ASRSysGroupPermissions.itemID = @iSSIntItemID
							AND ASRSysGroupPermissions.groupName = @sRoleName
							AND ASRSysGroupPermissions.permitted = 1
						IF @iCount = 0
						BEGIN				
							SET @piSuccessFlag = 0
							SET @psErrorMessage = 'You are not permitted to use the Data Manager Intranet module with this user name.'
						END
						ELSE
						BEGIN
							SET @piUserType = 1
						END
					END
				END
			END
		END
	END	
	IF @piSuccessFlag = 1
	BEGIN
		IF @fSelfService = 1
		BEGIN
			SET @iTemp = @iSSIUsers
		END
		ELSE
		BEGIN
			IF @piUserType = 1
			BEGIN
				SET @iTemp = @iSSUsers
			END
			ELSE
			BEGIN
				SET @iTemp = @iFullUsers
			END
		END
		SET @iValue = 0
		/* Don't use uid as it sometimes is 0 when youdon't expect it to be. */
		DECLARE @tmpSysProcess2 TABLE (
			hostname		nvarchar(50),
			loginname		nvarchar(50),
			program_name	nvarchar(50),
			hostprocess		int,
			[sid]			binary(86), 
			ogin_time		datetime,
			spid			smallint,
			[uid]			smallint);
			
		INSERT @tmpSysProcess2 EXEC [dbo].[spASRGetCurrentUsers]
		
		DECLARE users_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT sid
			FROM @tmpSysProcess2
			WHERE program_name = APP_NAME();
			
		OPEN users_cursor
		FETCH NEXT FROM users_cursor INTO @iSID
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @fSelfService = 1
			BEGIN
				SET @iValue = @iValue + 1
			END
			ELSE
			BEGIN
				/* Check if the process is run by the same type of user as the current user. */
				/* Get the user's group name. */
				SELECT @sRoleName = usg.name
				FROM sysusers usu
				left outer join
					(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
				WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) 
					AND (usg.issqlrole = 1 OR usg.uid is null) 
					AND usu.sid = @iSID 
					AND not (usg.name like 'ASRSys%') 
					AND not (usg.name like 'db_owner')
				IF @piUserType = 1
				BEGIN
					/* Self-service users. */
					IF @iSSIntItemID > 0
					BEGIN
						/* The permission does exist in the current version so check if the user is granted this permission. */
						SELECT @iCount = count(ItemID)
						FROM ASRSysGroupPermissions 
						WHERE ASRSysGroupPermissions.itemID = @iSSIntItemID
							AND ASRSysGroupPermissions.groupName = @sRoleName
							AND ASRSysGroupPermissions.permitted = 1
						IF @iCount > 0 SET @iValue = @iValue + 1
					END
				END
				ELSE
				BEGIN
					/* Full access users. */
					IF @iFullIntItemID > 0
					BEGIN
						/* The permission does exist in the current version so check if the user is granted this permission. */
						SELECT @iCount = count(*)
						FROM ASRSysGroupPermissions 
						WHERE ASRSysGroupPermissions.itemID = @iFullIntItemID
						AND ASRSysGroupPermissions.groupName = @sRoleName
						AND ASRSysGroupPermissions.permitted = 1
						IF @iCount > 0 SET @iValue = @iValue + 1
					END
				END
			END
			FETCH NEXT FROM users_cursor INTO @iSID
		END
		
		CLOSE users_cursor
		DEALLOCATE users_cursor
		IF @iValue > @iTemp
		BEGIN
			SET @piSuccessFlag = 0
			SET @psErrorMessage = 'Unable to logon. You have reached the maximum number of licensed ' + 
				CASE
					WHEN @fSelfService = 1 THEN 'Self-service Intranet'
					WHEN @piUserType = 1 THEN 'Data Manager Intranet (single record)'
					ELSE 'Data Manager Intranet (multiple record)'
				END +
				' users.'
		END
	END
	/* Check if the password has expired */
	SELECT @sSQLVersion = dbo.udfASRSQLVersion()
	IF @piSuccessFlag = 1 AND @sSQLVersion < 9
	BEGIN
		SELECT @dtPasswordLastChanged = lastChanged, 
			@fPasswordForceChange = forceChange
		FROM ASRSysPasswords
		WHERE userName = system_user
		IF @dtPasswordLastChanged IS NULL
		BEGIN
			/* User not in the password table. So add them. */
			SET @dtPasswordLastChanged = GETDATE()
			SET @fPasswordForceChange = 0
			INSERT INTO ASRSysPasswords (username, lastChanged, forceChange)
			VALUES (LOWER(system_user), @dtPasswordLastChanged, @fPasswordForceChange)
		END
		ELSE
		BEGIN
			IF (@iMinPasswordLength <> 0) OR (@iChangePasswordFrequency <> 0) 
			BEGIN
				/* Check for minimum length. */
				IF (@iMinPasswordLength > @piPasswordLength) SET @fPasswordForceChange = 1
    
				/* Check for Date last changed. */
				IF (@iChangePasswordFrequency > 0) AND (@fPasswordForceChange = 0)
				BEGIN
					IF @sChangePasswordPeriod = 'D' 
					BEGIN
						IF DATEADD(day, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
					END
					IF @sChangePasswordPeriod = 'W' 
					BEGIN
						IF DATEADD(week, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
					END
					IF @sChangePasswordPeriod = 'M' 
					BEGIN
						IF DATEADD(month, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
					END
					IF @sChangePasswordPeriod = 'Y' 
					BEGIN
						IF DATEADD(year, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
					END
				END
			END
		END
		IF @fPasswordForceChange = 1 SET @piSuccessFlag = 2
	END
	
END
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetPersonnelParameters] (
	@piEmployeeTableID	integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @fOK			bit,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sActualUserName	sysname;

	/* Personnel information. */
	SET @fOK = 1;
	SET @piEmployeeTableID = 0;

	/* Get the current user's group id. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
	OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'PERSONNEL', @fOK OUTPUT;

	/* Get the required training booking module paramaters. */
	IF @fOK = 1
	BEGIN
		/* Get the EMPLOYEE table information. */
		SELECT @piEmployeeTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_PERSONNEL'
			AND parameterKey = 'Param_TablePersonnel';
		IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0;
	END
END
GO

CREATE PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run]
(
	@pdReportStart      datetime,
	@pdReportEnd		datetime,
	@pcReportTableName  char(30)
) 
AS 
BEGIN

	SET NOCOUNT ON;

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

	declare @sSQL as varchar(MAX)
	declare @piParentID as integer
	declare @piID as integer
	declare @pbProcessed as bit

	declare @pdTempStartDate as datetime
	declare @pdTempEndDate as datetime
	declare @pcTempStartSession as char(2)
	declare @pcTempEndSession as char(2)
	declare @sTempEndDate as varchar(50)

	declare @pfCount as float
	declare @psVer as char(80)

	/* Alter the structure of the temporary table so it can hold the text for the days */
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ALTER COLUMN Hor NVARCHAR(10)'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ADD Processed BIT'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ADD DisplayOrder INT'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ALTER COLUMN Value decimal(10,5)'
	execute(@sSQL)

	/* Load the values from the temporary cursor */
	Set @sSQL = 'DECLARE AbsenceBreakdownCursor CURSOR STATIC FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM ' + @pcReportTableName
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

		/* The absence should only calculate for absence within the reporting period */
		set @pdTempStartDate = @pdStartDate
		set @pcTempStartSession = @pcStartSession
		set @pdTempEndDate = @pdEndDate
		set @pcTempEndSession = @pcEndSession

		--/* If blank leaving date set it to todays date */
		if @pdTempEndDate is Null set @pdTempEndDate = getdate()

		if @pdStartDate <  @pdReportStart
			begin
			set @pdTempStartDate = @pdReportStart
			set @pcTempStartSession = 'AM'
			end
		if @pdTempEndDate >  @pdReportEnd
			begin
			set @pdTempEndDate = @pdReportEnd
			set @pcTempEndSession = 'PM'
			end

		set @sTempEndDate = case when @pdEndDate is null then 'null' else '''' + convert(varchar(40),@pdEndDate) + '''' end

		/* Calculate the days this absence takes up */
		execute sp_ASR_AbsenceBreakdown_Calculate @pfDuration OUTPUT, @pdblMon OUTPUT, @pdblTue OUTPUT, @pdblWed OUTPUT, @pdblThu OUTPUT, @pdblFri OUTPUT, @pdblSat OUTPUT, @pdblSun OUTPUT, @pdTempStartDate, @pcTempStartSession, @pdTempEndDate, @pcTempEndSession, @piParentID

		/* Strip out dodgy characters */
		set @pcRecordDescription = replace(@pcRecordDescription,'''','')
		set @pcType = replace(@pcType,'''','')

		/* Add Mondays records */
		if @pdblMon > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 0) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblMon) + ',''' + convert(varchar(20),@pdStartDate) + ''',1,1,' + @sTempEndDate + ',1)'
			execute(@sSQL)
			end

		/* Add Tuesday records */
		if @pdblTue > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 1) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblTue) +  ',''' + convert(varchar(20),@pdStartDate) + ''',2,1,' + @sTempEndDate +',2)'
			execute(@sSQL)
			end

		/* Add Wednesdays records */
		if @pdblWed > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 2) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblWed) +  ',''' + convert(varchar(20),@pdStartDate) +  ''',3,1,' + @sTempEndDate +',3)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Thursdays were found */
		if @pdblThu > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 3) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblThu) +  ',''' + convert(varchar(20),@pdStartDate) + ''',4,1,' + @sTempEndDate +',4)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Fridays were found */
		if @pdblFri > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 4) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblFri) + ',''' + convert(varchar(20),@pdStartDate) + ''',5,1,' + @sTempEndDate +',5)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Saturdays were found */
		if @pdblSat > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 5) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblSat) + ','''+ convert(varchar(20),@pdStartDate) + ''',6,1,' + @sTempEndDate +',6)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Sundays were found */
		if @pdblSun > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 5) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblSun) + ',''' + convert(varchar(20),@pdStartDate) + ''',7,1,' + @sTempEndDate +',0)'
			execute(@sSQL)
			end

		/* Calculate total duraton of absence */
		set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun

		if @pfDuration > 0
			begin
			/* Write records for average, totals and count */
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Total'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pfDuration) + ',''' + convert(varchar(20),@pdStartDate) + ''',9,1,' + @sTempEndDate +',8)'
			execute(@sSQL)

			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Count'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),1) + ',''' + convert(varchar(20),@pdStartDate) + ''',10,1,' + @sTempEndDate +',10)'
			execute(@sSQL)

			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Average'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pfDuration) + ',''' + convert(varchar(20),@pdStartDate) + ''',9,1,' + @sTempEndDate +',9)'
			execute(@sSQL)
			end

		/* Process next record */
		Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed

		end

	/* Delete this record from our collection as it's now been processed */
	set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Processed IS NULL'
	execute(@sSQL)

	Set @sSQL = 'DECLARE CalculateAverage CURSOR STATIC FOR SELECT Ver,(SUM(Value) / COUNT(Value)) / COUNT(Value) FROM ' + @pcReportTableName + ' WHERE hor = ''Average'' GROUP BY Ver'
	execute(@sSQL)
	open CalculateAverage

	Fetch Next From CalculateAverage Into @psVer, @pfCount
	while @@FETCH_STATUS = 0
		begin
  			Set @sSQL = 'UPDATE ' + @pcReportTableName + ' SET Value = ' + Convert(varchar(10),@pfCount) + ' WHERE Ver =  ''' + @psVer + ''' AND Hor = ''Average'''
		execute(@sSQL)
			Fetch Next From CalculateAverage Into @psVer, @pfCount
		end

	/* Tidy up */
	close AbsenceBreakdownCursor
	close CalculateAverage
	deallocate AbsenceBreakdownCursor
	deallocate CalculateAverage

END
GO

CREATE PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pbOmitBeforeStart	bit,
	@pbOmitAfterEnd	bit,
	@pcReportTableName	char(30)
)
AS
BEGIN

	SET NOCOUNT ON;

	declare @piID as integer;
	declare @pdStartDate as datetime;
	declare @pdEndDate as datetime;
	declare @iDuration as float;
	declare @pbDeleteThisAbsence as bit;
	declare @sSQL as varchar(MAX);

	set @sSQL = 'DECLARE BradfordIndexCursor CURSOR FOR SELECT Absence_ID, Start_Date, End_Date, Duration FROM ' + @pcReportTableName;
	execute(@sSQL);
	open BradfordIndexCursor;

	Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
	while @@FETCH_STATUS = 0
		begin
			set @pbDeleteThisAbsence = 0;
			if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1;
			if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1;
			if @iDuration = 0 set @pbDeleteThisAbsence = 1;

			if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1;
			if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1;

			if @pbDeleteThisAbsence = 1
				begin
					set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Absence_ID = Convert(Int,' + Convert(char(10),@piId) + ')';
					execute(@sSQL);
				end

			Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
		end

	close BradfordIndexCursor;
	deallocate BradfordIndexCursor;

END
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetUserGroup]
	( 
	@psItemKey				varchar(50),
	@psUserGroup			varchar(250)	OUTPUT
	)
AS
BEGIN
	set @psUserGroup = '';
	/* SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements. */
	SET NOCOUNT ON;
	SET @psUserGroup = (SELECT CASE 
		WHEN (usg.uid IS null) THEN null
		ELSE usg.name
	END as groupname
	FROM sysusers usu 
	LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
	LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
	WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
		AND (usg.issqlrole = 1 OR usg.uid IS null)
		AND lo.loginname = SYSTEM_USER
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END IN (
				SELECT [groupName]
				FROM [dbo].[ASRSysGroupPermissions]
				WHERE itemID IN (
					SELECT [itemID]
					FROM [dbo].[ASRSysPermissionItems]
					WHERE categoryID = 1
					AND itemKey LIKE '%' + @psItemKey + '%'
				)  
				AND [permitted] = 1))
END
GO


CREATE PROCEDURE [dbo].[sp_ASRUniqueObjectName](
		  @psUniqueObjectName sysname OUTPUT
		, @Prefix sysname
		, @Type int)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @NewObj 		as sysname
		, @Count 			as integer
		, @sUserName		as sysname
		, @sCommandString	nvarchar(MAX)	
 		, @sParamDefinition	nvarchar(500);

	SET @sUserName = SYSTEM_USER;
	SET @Count = 1;
	SET @NewObj = @Prefix + CONVERT(varchar(100),@Count);

	WHILE (EXISTS (SELECT * FROM sysobjects WHERE id = object_id(@NewObj) AND sysstat & 0xf = @Type))
		OR (EXISTS (SELECT * FROM ASRSysSQLObjects WHERE Name = @NewObj AND Type = @Type))
		BEGIN
			SET @Count = @Count + 1;
			SET @NewObj = @Prefix + CONVERT(varchar(10),@Count);
		END

	INSERT INTO [dbo].[ASRSysSQLObjects] ([Name], [Type], [DateCreated], [Owner])
		VALUES (@NewObj, @Type, GETDATE(), @sUserName);

	SET @sCommandString = 'SELECT @psUniqueObjectName = ''' + @NewObj + '''';
	SET @sParamDefinition = N'@psUniqueObjectName sysname output';
	EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psUniqueObjectName output;

END
GO


DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
     INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
    OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
    AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
    IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
    BEGIN
        SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
    END
    ELSE
    BEGIN
        SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
    END

    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

