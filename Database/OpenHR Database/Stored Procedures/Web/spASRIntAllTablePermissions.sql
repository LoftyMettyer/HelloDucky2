CREATE PROCEDURE [dbo].[spASRIntAllTablePermissions]
(
	@psSQLLogin 		varchar(255)
)
AS
BEGIN
;
	SET NOCOUNT ON

	-- Cached view of the objects 
	DECLARE @SysObjects TABLE([ID]		integer PRIMARY KEY CLUSTERED,
							  [Name]	sysname);
							  
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
	SELECT p.ID, p.Columns, p.Action, p.ProtectType FROM #SysProtects p
		INNER JOIN @SysObjects o ON p.ID = o.ID
		WHERE ((p.ProtectType <> 206 AND p.Action <> 193) OR (p.Action = 193 AND p.ProtectType IN (204,205)));


	SELECT o.name, p.action
	FROM @SysProtects p
	INNER JOIN @SysObjects o ON p.id = o.id
	WHERE p.protectType <> 206
		AND p.action <> 193
	UNION
	SELECT o.name, 193
	FROM syscolumns
	INNER JOIN @SysProtects p ON (syscolumns.id = p.id
		AND p.action = 193 
		AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
		AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
		OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
		AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
	INNER JOIN @SysObjects o ON p.id = o.id
	WHERE (syscolumns.name <> 'timestamp' AND syscolumns.name <> 'ID')
		AND p.protectType IN (204, 205) 
	ORDER BY o.name;

END