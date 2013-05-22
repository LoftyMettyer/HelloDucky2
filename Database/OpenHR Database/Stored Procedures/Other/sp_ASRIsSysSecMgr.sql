CREATE PROCEDURE [dbo].[sp_ASRIsSysSecMgr] (
	@psGroupName		sysname,
	@pfSysSecMgr		bit	OUTPUT
)
AS
BEGIN
	DECLARE @iUserGroupID integer;

	/* Get the current user's group ID. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = @psGroupName;

	SELECT @pfSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER' OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';
END