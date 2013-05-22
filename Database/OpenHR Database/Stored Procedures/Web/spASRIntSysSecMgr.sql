CREATE PROCEDURE [dbo].[spASRIntSysSecMgr]
(
		@pfSysSecMgr bit OUTPUT
)
AS
BEGIN
	DECLARE
		@sRoleName			sysname,
		@sActualUserName	sysname,
		@iActualUserGroupID	integer;
		
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iActualUserGroupID OUTPUT;
					
	SELECT @pfSysSecMgr = 
		CASE
			WHEN (SELECT count(*)
				FROM ASRSysGroupPermissions
				INNER JOIN ASRSysPermissionItems 
					ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
				INNER JOIN ASRSysPermissionCategories 
					ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
				WHERE ASRSysGroupPermissions.groupname = @sRoleName
					AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 1
			ELSE 0
		END;
END