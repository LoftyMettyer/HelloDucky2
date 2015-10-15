CREATE PROCEDURE [dbo].[spASRWorkflowValidateService](@allow bit OUTPUT)
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
			AND pi.itemKey = 'SYSTEMMANAGER' AND gp.permitted = 1 AND pc.categorykey = 'MODULEACCESS';

END