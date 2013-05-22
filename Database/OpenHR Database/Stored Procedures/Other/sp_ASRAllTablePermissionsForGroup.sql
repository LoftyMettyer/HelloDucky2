CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]
(
	@psGroupName sysname
)
AS
BEGIN
	-- Return parameters showing what permissions the current user has on all of the tables.
	DECLARE @iUserGroupID	integer;

	-- Initialise local variables.
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = @psGroupName;

	SELECT sysobjects.name, sysprotects.action
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
		AND sysprotects.protectType <> 206
		AND (sysobjects.xtype = 'u' or sysobjects.xtype = 'v')
		AND (sysobjects.Name NOT LIKE 'ASRSYS%' OR sysobjects.Name LIKE 'ASRSYSCV%')
	ORDER BY sysobjects.name;
	
END