CREATE PROCEDURE [spASRGetAllTableAndViewColumnPermissionsForGroup](
	@piUID int)
AS
BEGIN

	SET NOCOUNT ON

	-- Cached cut down view of the sysprotects table
	DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
	INSERT @SysProtects
		SELECT [ID], [Action], [ProtectType], [Columns] FROM sysprotects
		WHERE [UID] = @piUID	AND [Action] IN (193, 197)
			AND [ProtectType] = 205

	DECLARE @Phase1 TABLE([TableViewName] sysname, [Name] sysname, [Select] smallint, [Edit] smallint)
		INSERT @Phase1
		SELECT o.name, c.name
			,CASE [Action] WHEN 193 THEN 1 ELSE 0 END
			,CASE [Action] WHEN 197 THEN 1 ELSE 0 END
		FROM @sysprotects p
		INNER JOIN sysobjects o ON p.id = o.id 
		INNER JOIN syscolumns c ON p.id = c.id 
		WHERE c.name <> 'timestamp' 
			AND (o.Name NOT LIKE 'ASRSYS%' OR o.Name LIKE 'ASRSYSCV%')
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0 
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
			 OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0 
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0))

	SELECT [TableViewName], [Name]
		, SUM([Select]) AS [Select]
		, SUM([Edit]) AS [Edit]
	FROM @Phase1		
	GROUP BY [TableViewName], [Name]
	ORDER BY [TableViewName], [Name]

end
