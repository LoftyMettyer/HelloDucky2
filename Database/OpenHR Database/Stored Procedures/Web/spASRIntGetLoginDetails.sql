CREATE PROCEDURE spASRIntGetLoginDetails
AS
BEGIN

	DECLARE @psUserName			nvarchar(MAX), 
			@psUserGroup		nvarchar(MAX),
			@piUserGroupID		integer,
			@licenseKey			varchar(MAX),
			@sSysManagerVersion	varchar(200),
			@sIntranetDBVersion	varchar(200),
			@bIsLocked			bit = 0,
			@sLockMessage		varchar(MAX) = '',
			@bUpdateInProgress	bit = 0;

	-- Get security info for this user
	EXEC [dbo].[spASRIntGetActualUserDetails] @psUserName OUTPUT, @psUserGroup OUTPUT, @piUserGroupID OUTPUT
	SELECT @psUserName, @psUserGroup, @piUserGroupID

	-- DB information
	SELECT @licenseKey = SettingValue FROM ASRSysSystemSettings WHERE section = 'Licence' AND SettingKey = 'Key';		
	SELECT @sSysManagerVersion = SettingValue FROM ASRSysSystemSettings WHERE section = 'database' AND SettingKey = 'version';		
	SELECT @sIntranetDBVersion = SettingValue FROM ASRSysSystemSettings WHERE section = 'intranet' AND SettingKey = 'version';		

	-- Lock information
	IF EXISTS(SELECT * FROM ASRSysLock WHERE Priority = 1)
	BEGIN
		SELECT @sLockMessage = SettingValue FROM ASRSysSystemSettings WHERE section = 'messaging' AND SettingKey = 'lockmessage';
		SET @bUpdateInProgress = 1;
	END

	IF EXISTS(SELECT * FROM ASRSysLock WHERE Priority = 2)
	BEGIN
		SET @bIsLocked = 1;
		SET @bUpdateInProgress = 1;
	END

	SELECT @licenseKey AS LicenseKey, @sSysManagerVersion AS SysMgrDBVersion, @sIntranetDBVersion AS IntDBVersion,
			@bUpdateInProgress AS UpdateInProgress, @bIsLocked AS IsLocked, @sLockMessage AS LockMessage

	-- Permissions for this user
	SELECT c.categoryKey, i.itemKey, i.categoryID, g.itemID, g.permitted AS permitted
		FROM [ASRSysGroupPermissions] g
			INNER JOIN ASRSysPermissionItems i ON i.itemID = g.itemID
			INNER JOIN ASRSysPermissionCategories c ON c.categoryID = i.categoryID
		WHERE g.groupName = @psUserGroup
		ORDER BY i.categoryID;

	-- Server roles
	SELECT IS_SRVROLEMEMBER('serveradmin') AS IsServerAdmin
		, IS_SRVROLEMEMBER('securityadmin') AS IsSecurityAdmin
		, IS_SRVROLEMEMBER('sysadmin') AS IsSysAdmin;

END
