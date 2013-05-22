CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissions] 
	(
	@psSQLLogin 		varchar(200)
	)
	AS
	BEGIN

		SET NOCOUNT ON

		/* Return parameters showing what permissions the current user has on all of the tables. */
		DECLARE @iUserGroupID	int

		/* Initialise local variables. */
		SELECT @iUserGroupID = usg.gid
		FROM sysusers usu
		left outer join
		(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
		WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
			(usg.issqlrole = 1 or usg.uid is null) and
			usu.name = @psSQLLogin AND not (usg.name like 'ASRSys%');

		-- Cached cut down view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
		INSERT @SysProtects
			SELECT [ID],[Action],[ProtectType], [Columns] FROM sysprotects
			WHERE [UID] = @iUserGroupID;

		-- Cached version of the Base table IDs
		DECLARE @BaseTableIDs TABLE([ID] int PRIMARY KEY CLUSTERED, [BaseTableID] int)
		INSERT @BaseTableIDs
			SELECT DISTINCT o.ID, v.TableID
			FROM sysobjects o
			INNER JOIN dbo.ASRSysChildViews2 v ON v.ChildViewID = CONVERT(integer,SUBSTRING(o.Name,9,PATINDEX ( '%#%' , o.Name) - 9))
			WHERE Name LIKE 'ASRSYSCV%';

		SELECT o.name, p.action, bt.BaseTableID
		FROM @SysProtects p
		INNER JOIN sysobjects o ON p.id = o.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE p.protectType <> 206
			AND p.action <> 193
			AND o.xtype = 'v'
			AND (o.Name NOT LIKE 'ASRSYS%' OR o.Name LIKE 'ASRSYSCV%')
		UNION
		SELECT o.name, 193, bt.BaseTableID
		FROM syscolumns
		INNER JOIN @SysProtects p ON (syscolumns.id = p.id
			AND p.action = 193 
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
		INNER JOIN sysobjects o ON o.id = p.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE (syscolumns.name <> 'timestamp' AND syscolumns.name <> 'ID')
			AND p.protectType IN (204, 205) 
			AND o.[xtype] = 'V'
		ORDER BY o.name;

	END
GO

