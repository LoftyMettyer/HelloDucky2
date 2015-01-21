/* ----------------------------------------------------------------------- */
/* Variable declarations                                                   */
/*                                                                         */
/* NB. Variable with the naming convention @sSPCode_<n> are declared below */
/* ----------------------------------------------------------------------- */
DECLARE @iCount integer,
    @sDBVersion varchar(10),
    @sTemp varchar(10),
    @iTemp integer,
    @iTemp2 integer,
    @iMajor integer,
    @iMinor integer,
    @iRevision integer,
    @NVarCommand nvarchar(4000),
    @JobID BINARY(16),
    @ReturnCode integer,
    @sJobName nvarchar(4000),
    @sDBName nvarchar(4000),
    @sErrMsg nvarchar(4000),
    @sVersion varchar(100),
    @sGroup sysname,
    @sObject sysname,
    @sObjectType char(2),
    @sSQL varchar(8000),
    @iCurrentUserCount integer,
    @iHRProLockCount integer,
    @iSQLVersion integer;

/* -------------------------- */
/* Set the new version number */
/* -------------------------- */
SET @sVersion = '8.0.28'
SET NOCOUNT ON;

/* ------------------------------------ */
/* Get the name of the current database */
/* ------------------------------------ */
SELECT @sDBName = master..sysdatabases.name
FROM master..sysdatabases
INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid
WHERE master..sysprocesses.spid = @@spid

/* -------------------------------------- */
/* Check SQL Server version compatibility */
/* -------------------------------------- */
SELECT @iSQLVersion = convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))
IF (@iSQLVersion < 10)
BEGIN
    Print '+--------------------------------------------------------------------------+'
    Print '|                                                                          |'
    Print '|                            SCRIPT FAILURE                                |'
    Print '|                                                                          |'
    Print '| This version of OpenHR is only compatible with SQL Server 2008 or later. |'
    Print '| Please upgrade SQL Server before upgrading to this version of OpenHR.    |'
    Print '|                                                                          |'
    Print '+--------------------------------------------------------------------------+'
		RETURN
END

IF @sDBName = 'master'
BEGIN
    Print '+-----------------------------------------------------------------------+'
    Print '|                                                                       |'
    Print '|                            SCRIPT FAILURE                             |'
    Print '|                                                                       |'
    Print '|        This script should not be run on the ''master'' database.      |'
    Print '|                                                                       |'
    Print '+-----------------------------------------------------------------------+'
    RETURN
END

IF IS_SRVROLEMEMBER('systemadmin') = 0
BEGIN
    Print '+-----------------------------------------------------------------------+'
    Print '|                                                                       |'
    Print '|                            SCRIPT FAILURE                             |'
    Print '|                                                                       |'
    Print '| This script can only be run by a member of the ''systemadmin'' role.  |'
    Print '|                                                                       |'
    Print '+-----------------------------------------------------------------------+'
    RETURN
END


GO


DECLARE @sVersion varchar(10) = '8.0.42'

EXEC spsys_setsystemsetting 'database', 'version', '8.0';
EXEC spsys_setsystemsetting 'intranet', 'version', @sVersion;
EXEC spsys_setsystemsetting 'ssintranet', 'version', @sVersion;


/*---------------------------------------------*/
/* Insert a record into the Audit Access table */
/*---------------------------------------------*/
INSERT INTO ASRSysAuditAccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
VALUES (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'Intranet','v8.0.28' + @sVersion)

Print ''
Print '+-----------------------------------------------------------------------+'
Print '|                                                                       |'
Print '|                            SCRIPT SUCCESS                             |'
Print '|                                                                       |'
Print '+-----------------------------------------------------------------------+'

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

GoTo EndSave

QuitWithRollback:
    Print ''
    Print '+-----------------------------------------------------------------------+'
    Print '|                                                                       |'
    Print '|                            SCRIPT FAILURE                             |'
    Print '|                                                                       |'
    Print '|                    See error messages listed above                    |'
    Print '|                                                                       |'
    Print '+-----------------------------------------------------------------------+'
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION

EndSave:
