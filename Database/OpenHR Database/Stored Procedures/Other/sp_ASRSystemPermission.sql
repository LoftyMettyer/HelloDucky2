CREATE PROCEDURE [dbo].[sp_ASRSystemPermission]
(
	@pfPermissionGranted 	bit OUTPUT,
	@psCategoryKey			varchar(50),
	@psPermissionKey		varchar(50),
	@psSQLLogin 			varchar(200)
)
AS
BEGIN
	
	-- Return 1 if the given permission is granted to the current user, 0 if it is not.
	DECLARE @fGranted bit,
			@sGroupName varchar(255);

	-- Is logged in user a system administrator
	SELECT @fGranted = sysAdmin FROM master..syslogins WHERE loginname = @psSQLLogin;

	IF @fGranted = 0
	BEGIN
		SELECT @sGroupName = usg.name
		FROM sysusers usu
		left outer join
		(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
		WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
			(usg.issqlrole = 1 or usg.uid is null) and
			usu.name = @psSQLLogin AND not (usg.name like 'ASRSys%')
			AND not (usg.name = 'db_owner');

		SELECT @fGranted = ASRSysGroupPermissions.permitted
		FROM ASRSysGroupPermissions
			INNER JOIN ASRSysPermissionItems 
				ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
			INNER JOIN ASRSysPermissionCategories
				ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID
		WHERE ASRSysPermissionItems.itemKey = @psPermissionKey
			AND ASRSysGroupPermissions.groupName = @sGroupName
			AND ASRSysPermissionCategories.categoryKey = @psCategoryKey;
	END


	IF @fGranted IS NULL
	BEGIN
		SET @fGranted = 0;
	END

	SET @pfPermissionGranted = @fGranted;

END