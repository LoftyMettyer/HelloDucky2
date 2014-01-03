CREATE PROCEDURE [dbo].[spASRIntAllTablePermissions]
(
	@psSQLLogin 		varchar(255)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID		integer,
		@sUserGroupName				sysname,
		@sActualUserName			sysname;

	-- Cached view of the objects 
	DECLARE @SysObjects TABLE([ID]		integer PRIMARY KEY CLUSTERED,
							  [Name]	sysname);
		
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
							  
	INSERT INTO @SysObjects
		SELECT [ID], [Name] FROM sysobjects
		WHERE [Name] LIKE 'ASRSysCV%' AND [XType] = 'v'
		UNION 
		SELECT OBJECT_ID(tableName), TableName 
		FROM ASRSysTables
		WHERE NOT OBJECT_ID(tableName) IS null
		UNION
		SELECT OBJECT_ID(viewName), ViewName 
		FROM ASRSysViews
		WHERE NOT OBJECT_ID(viewName) IS null;

	-- Cached view of the sysprotects table
	DECLARE @SysProtects TABLE([ID]				integer,
							   [columns]		varbinary(8000),
							   [Action]			tinyint,
							   [ProtectType]	tinyint);
	INSERT INTO @SysProtects
	SELECT p.ID, p.Columns, p.Action, p.ProtectType FROM ASRSysProtectsCache p
		INNER JOIN @SysObjects o ON p.ID = o.ID
		WHERE p.UID = @iUserGroupID AND ((p.ProtectType <> 206 AND p.Action <> 193) OR (p.Action = 193 AND p.ProtectType IN (204,205)));

	SELECT UPPER(o.name) AS [name], p.action, ISNULL(cv.tableID,0) AS [tableid]
		FROM @SysProtects p
		INNER JOIN @SysObjects o ON p.id = o.id
		LEFT JOIN ASRSysChildViews2 cv ON cv.childViewID = CASE SUBSTRING(o.Name, 1, 8) WHEN 'ASRSysCV' THEN SUBSTRING(o.Name, 9, CHARINDEX('#',o.Name, 0) - 9) ELSE 0 END
		WHERE p.protectType <> 206
			AND p.action <> 193
	UNION
	SELECT UPPER(o.name) AS [name], 193, ISNULL(cv.tableID,0) AS [tableid]
		FROM sys.columns c
		INNER JOIN @SysProtects p ON (c.object_id = p.id
			AND p.action = 193 
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,c.column_id/8+1,1))&power(2,c.column_id&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,c.column_id/8+1,1))&power(2,c.column_id&7)) = 0)))
		INNER JOIN @SysObjects o ON p.id = o.id
		LEFT JOIN ASRSysChildViews2 cv ON cv.childViewID = CASE SUBSTRING(o.Name, 1, 8) WHEN 'ASRSysCV' THEN SUBSTRING(o.Name, 9, CHARINDEX('#',o.Name, 0) - 9) ELSE 0 END
		WHERE (c.name <> 'timestamp' AND c.name <> 'ID')
			AND p.protectType IN (204, 205) 
		ORDER BY name;

END