CREATE PROCEDURE [dbo].[sp_ASRIntGetSystemPermissions]
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the IDs and names of the views of the given table for use in the link find page. */
	DECLARE @sGroupName			varchar(255)
	DECLARE @sActualUserName	varchar(250),
			@iActualUserGroupID	integer;

	/* Check if the current user is a System or Security manager. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SELECT ASRSysPermissionCategories.categoryKey + '_' + 	ASRSysPermissionItems.itemkey AS [key],
			1 AS [permitted]
		FROM ASRSysPermissionCategories
		INNER JOIN ASRSysPermissionItems ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID;
	END
	ELSE
	BEGIN	
		EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sGroupName OUTPUT,
			@iActualUserGroupID OUTPUT;
					
		SELECT ASRSysPermissionCategories.categoryKey + '_' + ASRSysPermissionItems.itemKey AS [key],
			CASE
				WHEN NOT ASRSysGroupPermissions.permitted IS NULL THEN ASRSysGroupPermissions.permitted
				ELSE 0
			END AS [permitted]
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		LEFT OUTER JOIN ASRSysGroupPermissions ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
			AND ASRSysGroupPermissions.groupName = @sGroupName;
	END
END