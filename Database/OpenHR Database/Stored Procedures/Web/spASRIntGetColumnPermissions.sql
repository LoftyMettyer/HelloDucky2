CREATE PROCEDURE [dbo].[spASRIntGetColumnPermissions]
	(@psSourceArray varchar(MAX))
AS
BEGIN
	
	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
			@sActualUserName	sysname,
			@sRoleName				sysname;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	SELECT sysobjects.name AS tableViewName, syscolumns.name AS columnName, p.action
		, CASE p.protectType WHEN 205 THEN 1 WHEN 204 THEN 1 ELSE 0 END AS permission
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects ON p.id = sysobjects.id INNER JOIN syscolumns ON p.id = syscolumns.id 
		WHERE p.uid = @iUserGroupID
			AND p.action = 193 or p.action = 197
			AND syscolumns.name <> 'timestamp' AND sysobjects.name IN (@psSourceArray)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0) OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		ORDER BY tableViewName;

END

