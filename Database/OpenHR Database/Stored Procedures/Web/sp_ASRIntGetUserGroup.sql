CREATE PROCEDURE [dbo].[sp_ASRIntGetUserGroup]
	( 
	@psItemKey				varchar(50),
	@psUserGroup			varchar(250)	OUTPUT
	)
AS
BEGIN
	set @psUserGroup = '';
	/* SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements. */
	SET NOCOUNT ON;
	SET @psUserGroup = (SELECT CASE 
		WHEN (usg.uid IS null) THEN null
		ELSE usg.name
	END as groupname
	FROM sysusers usu 
	LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
	LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
	WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
		AND (usg.issqlrole = 1 OR usg.uid IS null)
		AND lo.loginname = SYSTEM_USER
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'
		AND CASE 
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
			END IN (
				SELECT [groupName]
				FROM [dbo].[ASRSysGroupPermissions]
				WHERE itemID IN (
					SELECT [itemID]
					FROM [dbo].[ASRSysPermissionItems]
					WHERE categoryID = 1
					AND itemKey LIKE '%' + @psItemKey + '%'
				)  
				AND [permitted] = 1))
END
