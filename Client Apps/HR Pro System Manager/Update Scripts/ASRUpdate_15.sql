/* -------------------------------------------------- */
/* Update the database from version 14 to version 15. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16)


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

/* Exit if the database is not version 14 or 15. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 14) or (@iDBVersion > 15)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------- */
/* Drop and recreate sp_ASRGetAuditTrail */
/* ------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetAuditTrail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetAuditTrail]

EXEC('CREATE PROCEDURE sp_ASRGetAuditTrail (
	@piAuditType	int,
	@psOrder 	varchar(200))
AS
BEGIN
	DECLARE @sSQL varchar(2000)

	IF @piAuditType = 1
	BEGIN
		SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
			ASRSysAuditTrail.dateTimeStamp AS [Date / Time], 
			ASRSysTables.tableName AS [Table], 
			ASRSysColumns.columnName AS [Column], 
			ASRSysAuditTrail.oldValue AS [Old Value], 
			ASRSysAuditTrail.newValue AS [New Value], 
			ASRSysAuditTrail.recordDesc AS [Record Description],
			ASRSysColumns.columnID AS [ColumnID],
			ASRSysAuditTrail.id
			FROM ASRSysAuditTrail 
			INNER JOIN ASRSysTables ON ASRSysAuditTrail.TableID = ASRSysTables.TableID 
			INNER JOIN ASRSysColumns ON ASRSysAuditTrail.ColumnID = ASRSysColumns.columnID ''

		IF LEN(@psOrder) >0
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END
	ELSE IF @piAuditType = 2
	BEGIN
		SET @sSQL =  ''SELECT userName AS [User], 
			dateTimeStamp AS [Date / Time],
			groupName AS [User Group],
			viewTableName AS [View / Table],
			columnName AS [Column], 
			action AS [Action],
			permission AS [Permission], 
			id
			FROM ASRSysAuditPermissions ''

		IF LEN(@psOrder) > 0
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END	
	END
	ELSE IF @piAuditType = 3
	BEGIN
		SET @sSQL = ''SELECT userName AS [User],
    			dateTimeStamp AS [Date / Time],
			groupName AS [User Group], 
			userLogin AS [User Login],
			[Action], 
			id
			FROM ASRSysAuditGroup ''

		IF LEN(@psOrder) > 0 
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END
	ELSE IF @piAuditType = 4
	BEGIN
		SET @sSQL = ''SELECT DateTimeStamp AS [Date / Time],
    			UserGroup AS [User Group],
			UserName AS [User], 
			ComputerName AS [Computer Name],
			HRProModule AS [HR Pro Module],
			Action AS [Action], 
			id
			FROM ASRSysAuditAccess ''

		IF LEN(@psOrder) > 0 
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END

END')


/* ------------------------------------------ */
/* Drop and recreate sp_ASRFn_AbsenceDuration */
/* ------------------------------------------ */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AbsenceDuration]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_AbsenceDuration]

EXEC('CREATE PROCEDURE sp_ASRFn_AbsenceDuration (
	@pdblResult		float OUTPUT,						
	@pdtStartDate		datetime,						
	@psStartSession	varchar(255),						
	@pdtEndDate		datetime,					
	@psEndSession		varchar(255),					
	@iPersonnelID           	int)                   			

AS
BEGIN

/* Used to work out if we can hit child tables directly, or via childviews */
DECLARE @iUserGroupID			int
DECLARE @fSysSecMgr				bit

/* Personnel Table ID and name...used for static region/wp and for ID_xx purposes */
DECLARE @sPersonnelTable 			varchar(255)
DECLARE @iPersonnelTableID        		int

/* The Bank Holiday Region (Primary) Table which contains England, Scotland, Wales etc. */
DECLARE @iBHolRegionTableID			int		
DECLARE @sBHolRegionTableName		sysname	
DECLARE @sBHolRegionColumnName		sysname	

/* The Bank Holiday Instance (Child) Table which contains 25/12/00, 26/12/00 etc. */
DECLARE @iBHolTableID				int		
DECLARE @sBHolTableName			sysname	
DECLARE @sBHolDateColumnName		sysname	

/* Flag storing if the Bank Hols are setup ok and therefore if we should use them or not */
DECLARE @fBHolSetupOK           			bit

/* ID of the persons region...used to work out which dates from the BHol Instance table apply to the employee */
DECLARE @iBHolRegionID			int

/* Date variables used when working out the next change date for historic WP/Regions - If applicable */
DECLARE @dTempDate				datetime
DECLARE @dNextChange_Region		datetime
DECLARE @dNextChange_WP			datetime

/* Date variable used to cycle through dates between start date and end date */
DECLARE @dtCurrentDate			datetime

/* Flag stating if we are using historic region setup (True) or static (False) */
DECLARE @fHistoricRegion			bit

/* Variables to hold the relevant region table/column names */
DECLARE @sStaticRegionColumnName 		varchar(255)
DECLARE @sHistoricRegionTableName 		varchar(255)
DECLARE @sHistoricRegionColumnName 		varchar(255)
DECLARE @sHistoricRegionDateColumnName 	varchar(255)

/* Flag stating if we are using historic wp setup (True) or static (False) */
DECLARE @fHistoricWP				bit

/* Variables to hold the relevant wp table/column names */
DECLARE @sStaticWPColumnName 		varchar(255)
DECLARE @sHistoricWPTableName 		varchar(255)
DECLARE @sHistoricWPColumnName 		varchar(255)
DECLARE @sHistoricWPDateColumnName 	varchar(255)

/* The current wp/region being used in the calculation */
DECLARE @psWorkPattern	        		varchar(255)   
DECLARE @psPersonnelRegion			varchar(255)   

/* Flags derived from @psWorkPattern */
DECLARE @fWorkAM				bit
DECLARE @fWorkPM				bit
DECLARE @fWorkOnSundayAM			bit
DECLARE @fWorkOnSundayPM			bit
DECLARE @fWorkOnMondayAM			bit
DECLARE @fWorkOnMondayPM			bit
DECLARE @fWorkOnTuesdayAM			bit
DECLARE @fWorkOnTuesdayPM			bit
DECLARE @fWorkOnWednesdayAM		bit
DECLARE @fWorkOnWednesdayPM		bit
DECLARE @fWorkOnThursdayAM		bit
DECLARE @fWorkOnThursdayPM		bit
DECLARE @fWorkOnFridayAM			bit
DECLARE @fWorkOnFridayPM			bit
DECLARE @fWorkOnSaturdayAM		bit
DECLARE @fWorkOnSaturdayPM			bit
DECLARE @iDayOfWeek				int
DECLARE @sCommandString			nvarchar(4000)
DECLARE @iCount				int
DECLARE @sParamDefinition 			nvarchar(500)

/* Initialise the result to be 0 */
SET @pdblResult = 0

/* Get the current users group ID */
SELECT @iUserGroupID = sysusers.gid
FROM sysusers
WHERE sysusers.name = CURRENT_USER

/* Check if the current user is a System or Security manager. */
SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
FROM ASRSysGroupPermissions
INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
	OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

/* Get the ID of the BHol Region Table (which contains England, Scotland etc */
SELECT @iBHolRegionTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
AND ParameterKey = ''Param_TableBHolRegion''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Region Table (which contains England, Scotland etc */
SELECT @sBHolRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolRegionTableID

/* Get the name of the BHol Region column in the BHol Region Table */
SELECT @sBHolRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
	AND ParameterKey = ''Param_FieldBHolRegion''
	AND ParameterType = ''PType_ColumnID''

/* Get the ID of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @iBHolTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
AND ParameterKey = ''Param_TableBHol''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @sBHolTableName = AsrSysTables.TableName 
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolTableID

/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sBHolTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' AND ParameterKey = ''Param_TableBHol'' AND ParameterType = ''PType_TableID''))
END

/* Get the name of the BHol Date column. */
SELECT @sBHolDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
	AND ParameterKey = ''Param_FieldBHolDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to state whether BHols have been setup correctly or Not */
IF ((NOT @iBHolRegionTableID IS NULL)
AND (NOT @sBHolRegionTableName IS NULL) 
AND (NOT @sBHolRegionColumnName IS NULL) 
AND (NOT @iBHolTableID IS NULL) 
AND (NOT @sBHolTableName IS NULL) 
AND (NOT @sBHolDateColumnName IS NULL))
BEGIN
	SET @fBHolSetupOK = 1
END
ELSE
BEGIN
	SET @fBHolSetupOK = 0
END

/* Get the ID of the Personnel Table */
SELECT @iPersonnelTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
AND ParameterKey = ''Param_TablePersonnel''
AND ParameterType = ''PType_TableID''

/* Get the name of the Personnel Table */
SELECT @sPersonnelTable = AsrSysTables.TableName 
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_TablePersonnel''
	AND ParameterType = ''PType_TableID''

/* Get the Region Setup - Static Region*/
SELECT @sStaticRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsRegion''
	AND ParameterType = ''PType_ColumnID''
	
/* Get the Region Setup - Historic Region*/
SELECT @sHistoricRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegionTable''
	AND ParameterType = ''PType_TableID''


/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sHistoricRegionTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHRegionTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegion''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricRegionDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL''
	AND ParameterKey = ''Param_FieldsHRegionDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of regions we are to use */
IF @sStaticRegionColumnName is null
BEGIN
	IF (@sHistoricRegionTableName is null) OR (@sHistoricRegionColumnName is null) OR (@sHistoricRegionDateColumnName is null)
	BEGIN
		SET @pdblResult = 0
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricRegion = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricRegion = 0
END

/* Get the WP Setup - Static WP*/
SELECT @sStaticWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

/* Get the Region Setup - Historic WP */
SELECT @sHistoricWPTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL''
	AND ParameterKey = ''Param_FieldsHWorkingPatternTable''
	AND ParameterType = ''PType_TableID''

/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sHistoricWPTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHWorkingPatternTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL''
	AND ParameterKey = ''Param_FieldsHWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricWPDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPatternDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of wp we are to use */
IF @sStaticWPColumnName is null
BEGIN
	IF (@sHistoricWPTableName is null) OR (@sHistoricWPColumnName is null) OR (@sHistoricWPDateColumnName is null)
	BEGIN
		SET @pdblResult = 0
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricWP = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricWP = 0
END

/* Calculate the Absence Duration if all parameters have been provided. */
IF (NOT @pdtStartDate IS NULL) AND (NOT @psStartSession IS NULL) AND (NOT @pdtEndDate IS NULL) AND (NOT @psEndSession IS NULL)
BEGIN

	SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
	SET @pdtEndDate = convert(datetime, convert(varchar(20), @pdtEndDate, 101))
	SET @dtCurrentDate  = @pdtStartDate

	/* If we are using static wp and static region, do it the simple way */
	IF (@fHistoricRegion = 0) AND (@fHistoricWP = 0)
	BEGIN

		/* Get The Employees Working Pattern */
		SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
			
		/* Get The Employees Region */
		SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT

		/* Get the Region ID for the persons Region */
		SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +  '' FROM '' + @sBHolRegionTableName + '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
		SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT

		/* Determine which days are work days from the given work pattern. */
		SET @fWorkOnSundayAM = 0
		SET @fWorkOnSundayPM = 0
		SET @fWorkOnMondayAM = 0
		SET @fWorkOnMondayPM = 0
		SET @fWorkOnTuesdayAM = 0
		SET @fWorkOnTuesdayPM = 0
		SET @fWorkOnWednesdayAM = 0
		SET @fWorkOnWednesdayPM = 0
		SET @fWorkOnThursdayAM = 0
		SET @fWorkOnThursdayPM = 0
		SET @fWorkOnFridayAM = 0
		SET @fWorkOnFridayPM = 0
		SET @fWorkOnSaturdayAM = 0
		SET @fWorkOnSaturdayPM = 0

		IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
		IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
		IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
		IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
		IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
		IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
		IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
		IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
		IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
		IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
		IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
		IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
		IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
		IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

		WHILE @dtCurrentDate <= @pdtEndDate
		BEGIN

			/* Check if the current date is a work day. */
			SET @fWorkAM = 0
			SET @fWorkPM = 0
			SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)

			IF @iDayOfWeek = 1 
			BEGIN
				SET @fWorkAM = @fWorkOnSundayAM
				SET @fWorkPM = @fWorkOnSundayPM
			END
			IF @iDayOfWeek = 2
			BEGIN
				SET @fWorkAM = @fWorkOnMondayAM
				SET @fWorkPM = @fWorkOnMondayPM
			END
			IF @iDayOfWeek = 3
			BEGIN
				SET @fWorkAM = @fWorkOnTuesdayAM
				SET @fWorkPM = @fWorkOnTuesdayPM
			END
			IF @iDayOfWeek = 4
			BEGIN
				SET @fWorkAM = @fWorkOnWednesdayAM
				SET @fWorkPM = @fWorkOnWednesdayPM
			END
			IF @iDayOfWeek = 5
			BEGIN
				SET @fWorkAM = @fWorkOnThursdayAM
				SET @fWorkPM = @fWorkOnThursdayPM
			END
			IF @iDayOfWeek = 6
			BEGIN
				SET @fWorkAM = @fWorkOnFridayAM
				SET @fWorkPM = @fWorkOnFridayPM
			END
			IF @iDayOfWeek = 7
			BEGIN
				SET @fWorkAM = @fWorkOnSaturdayAM
				SET @fWorkPM = @fWorkOnSaturdayPM
			END

			IF (@fWorkAM = 1) OR (@fWorkPM = 1)
			BEGIN
				IF @fBHolSetupOK = 1
				BEGIN

					/* Check that the current date is not a company holiday. */
					SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' + '' FROM '' + @sBHolTableName + 
 			 				               '' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
   								  '''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
					SET @sParamDefinition = N''@count int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT
	
					IF @iCount = 0
					BEGIN
						IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM'')) SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @dtCurrentDate = @pdtStartDate
							BEGIN	
								IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @dtCurrentDate = @pdtEndDate
								BEGIN
									IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
									IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
								END
								ELSE
								BEGIN
									IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
									IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
								END
							END
						END
					END
				END
				ELSE
				BEGIN

					IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)
					BEGIN	
						IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
						IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM'')) SET @pdblResult = @pdblResult + 0.5
					END
					ELSE
					BEGIN
						IF @dtCurrentDate = @pdtStartDate
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @dtCurrentDate = @pdtEndDate
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
						END
					END
				END
			END
		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1

		END

	END
	ELSE  /* else for if we are using all static or not */
	BEGIN 

		WHILE @dtCurrentDate <= @pdtEndDate
		BEGIN

			/* We are using a historic region, so ensure we have the right region for the @dCurrentDate */
			IF @fHistoricRegion = 1
			BEGIN
				/* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */
				IF (@dnextchange_region IS NULL) OR ((@dtCurrentDate >= @dNextChange_Region) And (@dtCurrentDate <> ''12/31/9999''))
				BEGIN

					/* Get The Employees Region For @dCurrentDate */
					SET @sCommandString = ''SELECT TOP 1 @psPersonnelRegion = '' + @sHistoricRegionColumnName +
								  '' FROM '' + @sHistoricRegionTableName +
								  '' WHERE '' + @sHistoricRegionDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' DESC'' 
					SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT


					/* Get the Region ID for the persons Region */
					SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
	               						 '' FROM '' + @sBHolRegionTableName + 
							              '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
					SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
						
					/* Get the date of next change for the Region */
					SET @sCommandString = ''SELECT TOP 1 @dTempDate = '' + @sHistoricRegionDateColumnName +
								  '' FROM '' + @sHistoricRegionTableName +
 								  '' WHERE '' + @sHistoricRegionDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' ASC'' 
					SET @sParamDefinition = N''@dTempDate datetime OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dTempDate OUTPUT
						
					IF @dTempDate IS NULL
					BEGIN
						SET @dNextChange_Region = ''12/31/9999''
					END
					ELSE
					BEGIN
						SET @dNextChange_Region = @dTempDate
					END
				END

			END
			ELSE
			BEGIN
				/* We are using a static region, so get it */
				SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
				SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT

				/* Get the Region ID for the persons Region */
				SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
                  					               '' FROM '' + @sBHolRegionTableName + 
						               '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
				SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
			END

			/* We are using a historic wp so ensure we are getting the right wp for @dCurrentDate */
			IF @fHistoricWP = 1  
			BEGIN
				IF (@dnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dNextChange_WP) And (@dtCurrentDate <> ''12/31/9999''))
				BEGIN
					/* Get The Employees WP For @dCurrentDate */
					SET @sCommandString = ''SELECT TOP 1 @psWorkPattern = '' + @sHistoricWPColumnName +
								  '' FROM '' + @sHistoricWPTableName +
								  '' WHERE '' + @sHistoricWPDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricWPDateColumnName + '' DESC'' 
					SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
					/* Get The next change date for WP */
					SET @sCommandString = ''SELECT TOP 1 @dTempDate = '' + @sHistoricWPDateColumnName +
								  '' FROM '' + @sHistoricWPTableName +
								  '' WHERE '' + @sHistoricWPDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricWPDateColumnName + '' ASC'' 
					SET @sParamDefinition = N''@dTempDate datetime OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dTempDate OUTPUT
					IF @dTempDate IS NULL
					BEGIN
						SET @dNextChange_WP = ''12/31/9999''
					END
					ELSE
					BEGIN
						SET @dNextChange_WP = @dTempDate
					END
				END
			END
			ELSE
			BEGIN
				/* We are using a static wp, so get it */
				SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
				SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
			END

			/* Determine which days are work days from the given work pattern. */
			SET @fWorkOnSundayAM = 0
			SET @fWorkOnSundayPM = 0
			SET @fWorkOnMondayAM = 0
			SET @fWorkOnMondayPM = 0
			SET @fWorkOnTuesdayAM = 0
			SET @fWorkOnTuesdayPM = 0
			SET @fWorkOnWednesdayAM = 0
			SET @fWorkOnWednesdayPM = 0
			SET @fWorkOnThursdayAM = 0
			SET @fWorkOnThursdayPM = 0
			SET @fWorkOnFridayAM = 0
			SET @fWorkOnFridayPM = 0
			SET @fWorkOnSaturdayAM = 0
			SET @fWorkOnSaturdayPM = 0
		
			IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
			IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
			IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
			IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
			IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
			IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
			IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
			IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
			IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
			IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
			IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
			IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
			IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
			IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

			/* Check if the current date is a work day. */
			SET @fWorkAM = 0
			SET @fWorkPM = 0
			SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)
			IF @iDayOfWeek = 1 
			BEGIN
				SET @fWorkAM = @fWorkOnSundayAM
				SET @fWorkPM = @fWorkOnSundayPM
			END
			IF @iDayOfWeek = 2
			BEGIN
				SET @fWorkAM = @fWorkOnMondayAM
				SET @fWorkPM = @fWorkOnMondayPM
			END
			IF @iDayOfWeek = 3
			BEGIN
				SET @fWorkAM = @fWorkOnTuesdayAM
				SET @fWorkPM = @fWorkOnTuesdayPM
			END
			IF @iDayOfWeek = 4
			BEGIN
				SET @fWorkAM = @fWorkOnWednesdayAM
				SET @fWorkPM = @fWorkOnWednesdayPM
			END
			IF @iDayOfWeek = 5
			BEGIN
				SET @fWorkAM = @fWorkOnThursdayAM
				SET @fWorkPM = @fWorkOnThursdayPM
			END
			IF @iDayOfWeek = 6
			BEGIN
				SET @fWorkAM = @fWorkOnFridayAM
				SET @fWorkPM = @fWorkOnFridayPM
			END
			IF @iDayOfWeek = 7
			BEGIN
				SET @fWorkAM = @fWorkOnSaturdayAM
				SET @fWorkPM = @fWorkOnSaturdayPM
			END

			IF (@fWorkAM = 1) OR (@fWorkPM = 1)
			BEGIN
				IF @fBHolSetupOK = 1
				BEGIN

					/* Check that the current date is not a company holiday. */
					SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' +
              							  '' FROM '' + @sBHolTableName + 
							               '' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
  								  '''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
					SET @sParamDefinition = N''@count int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT
					IF @iCount = 0
					BEGIN
						IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM'')) SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN

							IF @dtCurrentDate = @pdtStartDate
							BEGIN	
								IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @dtCurrentDate = @pdtEndDate
								BEGIN
									IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
									IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
								END
								ELSE
								BEGIN
									IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
									IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
								END
							END
						END
					END
				END
				ELSE
				BEGIN
					/* We arent using Bholidays, so just add to the result */
					IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)
					BEGIN	
						IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
						IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM'')) SET @pdblResult = @pdblResult + 0.5
					END
					ELSE
					BEGIN

						IF @dtCurrentDate = @pdtStartDate
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @dtCurrentDate = @pdtEndDate
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
						END
					END
				END
			END
		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1

		END

	END /* end for if we are using all static or not */

END /* end for if all the parameters have been provided */

END')



/* ------------------------ */
/* Create RecordDesc column */
/* ------------------------ */

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'RecordDesc'
	AND sysobjects.name = 'ASRSysEmailQueue'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysEmailQueue]
		ADD [RecordDesc] [varchar] (255) NULL 
END


/* ---------------------------------- */
/* Drop and recreate sp_ASREmailQueue */
/* ---------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASREmailQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASREmailQueue]

EXEC('CREATE PROCEDURE sp_ASREmailQueue AS
BEGIN

Declare @sSQL varchar(8000)
declare @queueid int
declare @recordid int
declare @recorddescid int
declare @recorddesc varchar(8000)


DECLARE emailqueue_cursor CURSOR LOCAL FAST_FORWARD FOR 
SELECT QueueID, RecordID, ASRSysTables.RecordDescExprID as RecDescID 
FROM ASRSysEmailQueue
JOIN ASRSysEmailLinks ON ASRSysEmailQueue.LinkID = ASRSysEmailLinks.LinkID
JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysEmailLinks.ColumnID
JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID

OPEN emailqueue_cursor
FETCH NEXT FROM emailqueue_cursor INTO @queueid, @recordid, @recorddescid

WHILE (@@fetch_status = 0)
BEGIN

	SET @RecordDesc = ''''
	SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecordDescID)
	IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
	BEGIN
		EXEC @sSQL @RecordDesc OUTPUT, @Recordid
	END

	UPDATE ASRSysEmailQueue SET RecordDesc = @recordDesc WHERE queueid = @queueid
	FETCH NEXT FROM emailqueue_cursor INTO @queueid, @recordid, @recorddescid
END
CLOSE emailqueue_cursor
DEALLOCATE emailqueue_cursor

END')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Set the flag to refresh the stored procedures               */
/* ----------------------------------------------------------- */

UPDATE ASRSysConfig
SET databaseVersion = 15,
	systemManagerVersion = '1.1.13',
	securityManagerVersion = '1.1.13',
	dataManagerVersion = '1.1.13'

/* RH Note : intranet version is prob already 0.0.3 as JPD releases fixed sp's to QA as an when done */
