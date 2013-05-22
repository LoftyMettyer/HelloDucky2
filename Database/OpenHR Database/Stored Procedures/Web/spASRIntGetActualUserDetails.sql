CREATE PROCEDURE [dbo].[spASRIntGetActualUserDetails]
(
		@psUserName sysname OUTPUT,
		@psUserGroup sysname OUTPUT,
		@piUserGroupID integer OUTPUT
)
AS
BEGIN
	DECLARE @iFound		int
	DECLARE @sSQLVersion int
  DECLARE @sProgramName varchar(500)
  DECLARE @sPermissionItemKey varchar(500)
  
  SET @sPermissionItemKey = ''
  SET @sProgramName = ''
  
  SELECT @sProgramName = program_name FROM sys.sysprocesses WHERE spid = @@SPID
  IF CHARINDEX('Self-service', @sProgramName) > 0 
    BEGIN
      SET @sPermissionItemKey = 'SSINTRANET'
    END
  ELSE
    BEGIN
      SET @sPermissionItemKey = 'INTRANET'
    END
  
	SET @sSQLVersion = convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))

	SELECT @iFound = COUNT(*) 
	FROM sysusers usu 
	LEFT OUTER JOIN	(sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
	LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
	WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
		AND (usg.issqlrole = 1 OR usg.uid IS null)
		AND lo.loginname = system_user
		AND CASE
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
		END NOT LIKE 'ASRSys%' AND usg.name NOT LIKE 'db_owner'

	IF (@iFound > 0)
	BEGIN
		SELECT	@psUserName = usu.name,
			@psUserGroup = CASE 
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
			END,
			@piUserGroupID = usg.gid
		FROM sysusers usu 
		LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
		LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
		WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
			AND (usg.issqlrole = 1 OR usg.uid IS null)
			AND lo.loginname = system_user
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
														AND itemKey LIKE @sPermissionItemKey + '%'
													)  
					AND [permitted] = 1
	)			
	END
	ELSE
	BEGIN
		SELECT @psUserName = usu.name, 
			@psUserGroup = CASE
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
			END,
			@piUserGroupID = usg.gid
		FROM sysusers usu 
		LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
		LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
		WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
			AND (usg.issqlrole = 1 OR usg.uid IS null)
			AND is_member(lo.loginname) = 1
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
														AND itemKey LIKE @sPermissionItemKey + '%'
													)  
					AND [permitted] = 1
	)
	END

	IF @psUserGroup <> ''
	BEGIN
		DELETE FROM [ASRSysUserGroups] 
		WHERE [UserName] = SUSER_NAME()

		INSERT INTO [ASRSysUserGroups] 
		VALUES 
		(
			CASE
				WHEN @sSQLVersion <= 8 THEN USER_NAME()
				ELSE SUSER_NAME()
			END,
			@psUserGroup
		)
	END
END
