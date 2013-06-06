/* -------------------------------------------------- */
/* Update the database from version 24 to version 25. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @AuditCommand nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */

/* Check if the database version column exists. */

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'databaseVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0
BEGIN
	/* The database version column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [databaseVersion] [int] NULL 
END

/* Check if the refreshStoredProcedures column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'refreshStoredProcedures'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The refreshStoredProcedures column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [refreshStoredProcedures] [bit] NULL 
END

/* Check if the systemManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'systemManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The systemManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SystemManagerVersion] [varchar] (50)NULL 
END

/* Check if the securityManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'securityManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The securityManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SecurityManagerVersion] [varchar] (50)NULL 
END

/* Check if the DataManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DataManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The DataManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [DataManagerVersion] [varchar] (50)NULL 
END

/* Check if the IntranetVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'IntranetVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The IntranetVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [IntranetVersion] [varchar] (50)NULL 
END


SET @sCommand = N'SELECT @iDBVersion = databaseVersion
	FROM ASRSysConfig'
SET @sParam = N'@iDBVersion integer OUTPUT'
execute sp_executesql @sCommand, @sParam, @iDBVersion OUTPUT

IF @iDBVersion IS null SET @iDBVersion = 0

/* Exit if the database is not version 24 or 25. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 24) or (@iDBVersion > 25)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ---------------------------- */

PRINT 'Step 1 of 7 - Adding Function Round To Nearest Number'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_RoundToNearestNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_RoundToNearestNumber]

EXEC('CREATE PROCEDURE sp_ASRFn_RoundToNearestNumber
(
	@pfReturn 		float OUTPUT,
	@pfNumberToRound 	float,
	@pfNearestNumber	float
)
AS
BEGIN

	declare @pfRemainder as float

	/* Calculate the remainder. Cannot use the % because it only works on integers and not floats. */
	set @pfReturn = 0
	set @pfRemainder = @pfNumberToRound - (floor(@pfNumberToRound / @pfNearestNumber) * @pfNearestNumber)

	/* Formula for rounding to the nearest specified number */
	if @pfRemainder < (@pfNearestNumber / 2) set @pfReturn = @pfNumberToRound - @pfRemainder
		else set @pfReturn = @pfNumberToRound + @pfNearestNumber - @pfRemainder

END')

/* ---------------------------- */

PRINT 'Step 2 of 7 - Adding Function Round Up To Nearest Whole Number'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_RoundUpToNearestWholeNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_RoundUpToNearestWholeNumber]

EXEC('CREATE PROCEDURE sp_ASRFn_RoundUpToNearestWholeNumber 
(
	@piResult 	integer OUTPUT,	
	@pdblNumber 	float
)
AS
BEGIN
	SET @piResult = ceiling(@pdblNumber)
END')


/* ---------------------------- */

PRINT 'Step 3 of 7 - Adding Is Overnight Process Function'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsOvernightProcess]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_IsOvernightProcess]

EXEC('CREATE PROCEDURE sp_ASRFn_IsOvernightProcess
(
    @result integer OUTPUT
)

AS

SELECT @result = UpdatingDateDependentColumns FROM ASRSysConfig')

/* ---------------------------- */

PRINT 'Step 4 of 7 - Adding function parameters'

delete from asrsysfunctions where functionid in (48,49,50)

insert into asrsysfunctions
(functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
values(48, 'Round Up to Nearest Whole Number', 2, 0, 'Numeric', 'sp_ASRFn_RoundUpToNearestWholeNumber', 0, 1)

insert into asrsysfunctions
(functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
values(49, 'Round to Nearest Number', 2, 0, 'Numeric', 'sp_ASRFn_RoundToNearestNumber', 0, 1)

delete from asrsysfunctionparameters where functionid in (48,49)

insert into asrsysfunctionparameters
(functionID, parameterIndex, parameterType, parameterName)
values(48, 1, 2, '<Numeric>')

insert into asrsysfunctionparameters
(functionID, parameterIndex, parameterType, parameterName)
values(49, 2, 2, '<Nearest Number>')

insert into asrsysfunctionparameters
(functionID, parameterIndex, parameterType, parameterName)
values(49, 1, 2, '<Number To Round>')

INSERT INTO ASRSysFunctions
(functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
VALUES
(50,'Is Overnight Process', 3, 0, 'General', 'sp_ASRFn_IsOvernightProcess', 0, 0)

/* ---------------------------- */

PRINT 'Step 5 of 7 - Amending audit table'

alter table asrsysauditpermissions alter column viewtablename varchar(128)
alter table asrsysauditpermissions alter column columnname varchar(128)

/* ---------------------------- */

PRINT 'Step 6 of 7 - Amending export definition table'

alter table ASRSysExportName alter column Header int
alter table ASRSysExportName add HeaderText varchar(255)


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 7 of 7 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 25,
	systemManagerVersion = '1.1.23',
	securityManagerVersion = '1.1.23',
	dataManagerVersion = '1.1.23'

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
/* UPDATE ASRSysConfig SET refreshstoredprocedures = 1 */

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 25 Has Converted Your HR Pro Database To Use V1.1.23 Of HR Pro'
