﻿CREATE PROCEDURE dbo.spASRIntSetupTablesCollection
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID		integer,
			@sActualUserName	sysname,
			@sRoleName			sysname,
			@SysSecPerms		integer = 0;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;
	
	SELECT @SysSecPerms = COUNT(*) 
		FROM ASRSysGroupPermissions
			INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
			INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
			INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name AND a.name = @sRoleName
		WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER' OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
			AND ASRSysGroupPermissions.permitted = 1 AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';

	-- Security Info
	SELECT @sActualUserName AS [ActualLogin], @sRoleName AS [UserGroup], SYSTEM_USER AS [UserName]
		, CASE WHEN @SysSecPerms > 0 THEN 1 ELSE 0 END AS [IsSysSecMgr];

	-- Views
	SELECT v.viewID, UPPER(v.viewName) AS [viewname], t.tableID
		, UPPER(t.tableName) AS [tablename], t.tableType, t.defaultOrderID, t.recordDescExprID
		FROM ASRSysViews v
			INNER JOIN ASRSysTables t ON v.viewTableID = t.tableID;

END
