/****** Object:  StoredProcedure [dbo].[spASRIntGetUserGroup]    Script Date: 13/09/2013 08:57:58 ******/
DROP PROCEDURE [dbo].[spASRIntGetUserGroup]
GO

/****** Object:  StoredProcedure [dbo].[spASRIntGetUserGroup]    Script Date: 13/09/2013 08:58:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[spASRIntGetUserGroup]
	( 
	@psItemKey				varchar(50),
	@psUserGroup			varchar(250)	OUTPUT,
	@iSelfServiceUserType	integer			OUTPUT,
	@fSelfService			bit				OUTPUT
	)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sPermissionItemKey varchar(500),
		@iSSIntranetCount AS integer,
		@sIntranet_SelfService AS varchar(255),
		@sIntranet AS varchar(255);
	
	SET @psUserGroup = '';

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

	SET @sIntranet = (SELECT itemKey FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'INTRANET');
	
	SET @sIntranet_SelfService = (SELECT itemKey FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'INTRANET_SELFSERVICE');
	
	SET @iSSIntranetCount = (SELECT count(*) FROM ASRSysPermissionItems inner join ASRSysGroupPermissions ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	WHERE ASRSysGroupPermissions.groupName = @psUserGroup and permitted = 1 and categoryID = 1
	and ASRSysPermissionItems.itemKey = 'SSINTRANET');
		
	IF (@sIntranet is null) and (@sIntranet_SelfService is null) and (@iSSINTRANETcount = 0)
	/* No permissions at all  */
	BEGIN
		SET @sPermissionItemKey = 'NO PERMS'
		SET @iSelfServiceUserType = 0
		SET @fSelfService = 0
	END
	
	IF @sIntranet = 'INTRANET'
	/* IF DMI Multi automatically*/ 
	BEGIN
		SET @sPermissionItemKey = 'INTRANET'
		SET @iSelfServiceUserType = 1
		SET @fSelfService = 0
	END
	
	IF (@sIntranet_SelfService = 'INTRANET_SELFSERVICE') and (@iSSIntranetCount = 0)
	/* IF DMI Single Only*/ 
	BEGIN
		SET @sPermissionItemKey = 'INTRANET'
		SET @iSelfServiceUserType = 2
		SET @fSelfService = 0
	END	
	
	IF (@sIntranet_SelfService = 'INTRANET_SELFSERVICE') and (@iSSIntranetCount = 1)
	/* IF DMI Single And SSI */ 
	BEGIN
		SET @sPermissionItemKey = 'SSINTRANET'
		SET @iSelfServiceUserType = 3
		SET @fSelfService = 1
	END	
	
	IF  @iSSIntranetCount = 1 and (@sIntranet is null and  @sIntranet_SelfService is null)
	/* IF SSI Only */ 
	BEGIN
		SET @sPermissionItemKey = 'SSINTRANET'
		SET @iSelfServiceUserType = 4
		SET @fSelfService = 1
	END

	IF (@sIntranet = 'INTRANET') and (@iSSIntranetCount = 1)
	/* IF DMI Multi and SSI */
	BEGIN
		SET @sPermissionItemKey = 'INTRANET'
		SET @iSelfServiceUserType = 5
		SET @fSelfService = 0
	END

GO

DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
		IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
		BEGIN
				SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END
		ELSE
		BEGIN
				SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END

		FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

GO

GRANT EXEC ON TYPE::[dbo].[DataPermissions] TO ASRSysGroup

GO

EXEC spsys_setsystemsetting 'database', 'version', '8.0';
EXEC spsys_setsystemsetting 'intranet', 'version', '8.0.22';
EXEC spsys_setsystemsetting 'ssintranet', 'version', '8.0.22';

GO