
Declare @iRecCount int

---------------------------------------------------
PRINT 'Table ASRSysLock'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects
WHERE name = 'ASRSysLock'

if @iRecCount = 0
BEGIN
	CREATE TABLE [dbo].[ASRSysLock] (
		[Priority] [int] NULL ,
		[Description] [varchar] (50) NULL ,
		[Username] [varchar] (50) NULL ,
		[Hostname] [varchar] (50) NULL ,
		[Lock_Time] [datetime] NULL ,
		[Login_Time] [datetime] NULL ,
		[SPID] [int] NULL 
	) ON [PRIMARY]
END

SELECT @iRecCount = count(sysobjects.id)
	FROM sysobjects
	WHERE name = 'ASRSysCurrentSessions'


if @iRecCount = 0
BEGIN
	CREATE TABLE ASRSysCurrentSessions(
		[IISServer]		nvarchar(255),
		[Username]		nvarchar(128),
		[Hostname]		nvarchar(255),
		[SessionID]		nvarchar(255),
		[loginTime]		datetime,
		[WebArea]	varchar(255));
END

---------------------------------------------------
PRINT 'Procedure spASRGetCurrentUsers'

if exists (select * from sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASRGetCurrentUsers]

EXEC('CREATE PROCEDURE [dbo].[spASRGetCurrentUsers]
	AS
	BEGIN
		SET NOCOUNT ON

		IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
		   AND APP_NAME() NOT LIKE ''HR Pro Workflow%''
		   AND APP_NAME() NOT LIKE ''HR Pro Outlook%''
		   AND APP_NAME() NOT LIKE ''HR Pro Server.Net%''
		BEGIN
			EXEC sp_ASRIntCheckPolls
		END

		SELECT DISTINCT hostname, loginame, program_name, hostprocess, sid, login_time, spid
	    FROM master..sysprocesses
	    WHERE program_name like ''HR Pro%''
				  AND    program_name NOT LIKE ''HR Pro Workflow%''
				  AND    program_name NOT LIKE ''HR Pro Outlook%''
				  AND    program_name NOT LIKE ''HR Pro Server.Net%''
	    AND dbid in ( 
	                   SELECT dbid FROM master..sysdatabases
	                   WHERE name = DB_NAME())
	     ORDER BY loginame

	END')


---------------------------------------------------
PRINT 'Procedure spASRGetCurrentUsersFromMaster'

if exists (select * from sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsersFromMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASRGetCurrentUsersFromMaster]

EXEC('CREATE PROCEDURE [dbo].spASRGetCurrentUsersFromMaster
		    AS
		    BEGIN
		
				SET NOCOUNT ON
		
				SELECT p.hostname, p.loginame, p.program_name, p.hostprocess
					   , p.sid, p.login_time, p.spid
				FROM     master..sysprocesses p
				JOIN     master..sysdatabases d ON d.dbid = p.dbid
				WHERE    p.program_name LIKE ''HR Pro%''
				  AND    p.program_name NOT LIKE ''HR Pro Workflow%''
				  AND    p.program_name NOT LIKE ''HR Pro Outlook%''
				  AND    p.program_name NOT LIKE ''HR Pro Server.Net%''
				  AND    d.name = db_name()
				ORDER BY loginame
		
				SET NOCOUNT OFF
		
		    END')


---------------------------------------------------
PRINT 'Procedure sp_ASRLockCheck'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockCheck]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockCheck]

exec('CREATE PROCEDURE sp_ASRLockCheck AS
  BEGIN

    SET NOCOUNT ON

    SELECT ASRSysLock.* FROM ASRSysLock
    LEFT OUTER JOIN master..sysprocesses syspro 
      ON asrsyslock.spid = syspro.spid AND asrsyslock.login_time = syspro.login_time
    WHERE Priority = 2 OR syspro.spid IS NOT NULL
    ORDER BY Priority

  END')

---------------------------------------------------
PRINT 'Procedure sp_ASRLockDelete'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockDelete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockDelete]

exec('CREATE Procedure sp_ASRLockDelete (@LockType int)
AS
BEGIN
	DELETE FROM ASRSysLock WHERE Priority = @LockType
END')


---------------------------------------------------
PRINT 'Procedure sp_ASRLockWrite'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockWrite]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockWrite]

exec('CREATE Procedure sp_ASRLockWrite (@LockType int)
AS
BEGIN

	DECLARE @LockDesc varchar(50)
	DECLARE @OrigTranCount int

	SELECT @LockDesc = case @LockType
	WHEN 1 THEN ''Saving''
	WHEN 2 THEN ''Manual''
	WHEN 3 THEN ''Read Write''
	ELSE ''''
	END

	IF @LockDesc <> ''''
	BEGIN

		SET @OrigTranCount = @@trancount
		IF @OrigTranCount = 0 BEGIN TRANSACTION

		DELETE FROM ASRSysLock WHERE Priority = @LockType

		INSERT ASRSysLock (Priority, Description, Username, Hostname, Lock_Time, Login_Time, SPID)
		SELECT @LockType, @LockDesc, system_user, host_name(), getdate(), Login_Time, @@spid FROM master..sysprocesses WHERE spid = @@spid

		IF @OrigTranCount = 0 COMMIT TRANSACTION

	END

END')



---------------------------------------------------
PRINT 'Grant permission to database objects'

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)

DECLARE curNonDBOGroups CURSOR LOCAL FAST_FORWARD FOR 
SELECT name 
FROM sysusers
INNER JOIN ASRSysGroupPermissions nonSysMgrs ON (sysusers.name = nonSysMgrs.groupName)
INNER JOIN ASRSysPermissionItems nonSysMgrPerms ON nonSysMgrs.itemID = nonSysMgrPerms.itemID
            AND nonSysMgrPerms.categoryID = 1
            AND nonSysMgrPerms.itemKey = 'SYSTEMMANAGER'
            AND nonSysMgrs.permitted = 0
INNER JOIN ASRSysGroupPermissions nonSecMgrs ON (sysusers.name = nonSecMgrs.groupName)
INNER JOIN ASRSysPermissionItems nonSecMgrPerms ON nonSecMgrs.itemID = nonSecMgrPerms.itemID
            AND nonSecMgrPerms.categoryID = 1
            AND nonSecMgrPerms.itemKey = 'SECURITYMANAGER'
            AND nonSecMgrs.permitted = 0
WHERE sysusers.gid = sysusers.uid
            AND sysusers.uid > 0

OPEN curNonDBOGroups
FETCH NEXT FROM curNonDBOGroups INTO @sGroup
WHILE (@@fetch_status = 0)
BEGIN

		SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [ASRSysLock] TO [' + @sGroup + ']'
                EXEC(@sSQL)

                SET @sSQL = 'GRANT EXEC ON [sp_ASRLockCheck] TO [' + @sGroup + ']'
                EXEC(@sSQL)

                SET @sSQL = 'GRANT EXEC ON [sp_ASRLockDelete] TO [' + @sGroup + ']'
                EXEC(@sSQL)

                SET @sSQL = 'GRANT EXEC ON [sp_ASRLockWrite] TO [' + @sGroup + ']'
                EXEC(@sSQL)

            FETCH NEXT FROM curNonDBOGroups INTO @sGroup
END

CLOSE curNonDBOGroups
DEALLOCATE curNonDBOGroups


---------------------------------------------------

---Just in case we have moved SQL versions...
---(Ref 11375-11379 inclusive)
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[ASRTempSysProcesses]') and OBJECTPROPERTY(id, N'IsTable') = 1)
DROP TABLE [dbo].[ASRTempSysProcesses]


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
--delete from asrsyssystemsettings
--where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
--insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
--values('database', 'refreshstoredprocedures', 1)

