CREATE PROCEDURE [dbo].[sp_ASRIntGetExprCalcs] (
	@piCurrentExprID	integer,
	@piBaseTableID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the calc definitions. */
	DECLARE	 @sUserName SYSNAME,
			 @fSysSecMgr BIT,
			 @sRoleName VARCHAR(255),
			 @sActualUserName	VARCHAR(250),
			 @iActualUserGroupID INTEGER

	SET @sUserName = SYSTEM_USER;
	
	--Determine if user is an admin
	EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sRoleName OUTPUT,
			@iActualUserGroupID OUTPUT;

	SELECT @fSysSecMgr = 
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE ASRSysGroupPermissions.groupname = @sRoleName
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 1
				ELSE 0
			END;


	SELECT Name + char(9) +
		convert(varchar(255), exprID) + char(9) +
		userName AS [definitionString],
		[Description]
	FROM [dbo].[ASRSysExpressions]
	WHERE ExprID <> @piCurrentExprID
		AND Type = 10
		AND TableID = @piBaseTableID
		AND parentComponentID = 0
		AND (Username = @sUserName OR access <> 'HD' OR @fSysSecMgr = 1)
	ORDER BY name;
END