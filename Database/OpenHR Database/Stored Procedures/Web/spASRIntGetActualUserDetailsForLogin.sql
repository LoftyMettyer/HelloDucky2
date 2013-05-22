CREATE Procedure spASRIntGetActualUserDetailsForLogin
(
		@psLogin sysname,
		@psUserName sysname OUTPUT,
		@psUserGroup sysname OUTPUT,
		@piUserGroupID integer OUTPUT
)
AS
BEGIN
	DECLARE @iFound		int

	SELECT @iFound = COUNT(*) 
	FROM sysusers usu 
	LEFT OUTER JOIN	(sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
	LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
	WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
		AND (usg.issqlrole = 1 OR usg.uid IS null)
		AND lo.loginname = @psLogin
		AND CASE
			WHEN (usg.uid IS null) THEN null
			ELSE usg.name
		END NOT LIKE 'ASRSys%'

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
			AND lo.loginname = @psLogin
			AND CASE 
				WHEN (usg.uid IS null) THEN null
				ELSE usg.name
				END NOT LIKE 'ASRSys%'
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
			END NOT LIKE 'ASRSys%'
	END
END

GO

