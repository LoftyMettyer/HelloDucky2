CREATE PROCEDURE [dbo].[sp_ASRIntGetPersonnelParameters] (
	@piEmployeeTableID	integer	OUTPUT
)
AS
BEGIN
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