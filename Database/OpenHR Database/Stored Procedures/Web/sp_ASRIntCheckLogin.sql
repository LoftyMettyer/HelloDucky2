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
		@sPermissionItemKey varchar(500),
		@iSSIntranetCount AS integer,
		@sIntranet_SelfService AS varchar(255),
		@sIntranet AS varchar(255),
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
	SET @sPermissionItemKey = '';

	/* Deriving the User-group at the correct time especially after new users created was crucial so used this bit of code from later to do it */
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
									AND itemKey LIKE '%INTRANET%'
								)  
					AND [permitted] = 1))
	
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

	/*' Check if the current user is a SQL Server System Administrator.	We do not allow these users to login to the intranet module. */	
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