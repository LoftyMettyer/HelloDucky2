﻿CREATE PROCEDURE [dbo].[spASRIntGetColumnPermissions]
	(@SourceList AS dbo.dataPermissions READONLY)
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

	SELECT UPPER(so.name) AS tableViewName, UPPER(syscolumns.name) AS columnName, p.action
		, CASE p.protectType WHEN 205 THEN 1 WHEN 204 THEN 1 ELSE 0 END AS permission, sl.*
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects so ON p.id = so.id
		INNER JOIN @SourceList sl ON sl.name = so.name
		INNER JOIN syscolumns ON p.id = syscolumns.id 
		WHERE p.uid = @iUserGroupID
			AND (p.action = 193 OR p.action = 197)
			AND syscolumns.name <> 'timestamp'
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0) OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		ORDER BY tableViewName;

END
