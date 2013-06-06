
/* -------------------------------------------------- */
/* Update the database from version 28 to version 29. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@sDBVersion varchar(10),
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @NVarCommand nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */
SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

if @sDBVersion = ''
BEGIN
  SELECT @sDBVersion = SystemManagerVersion FROM ASRSysConfig
END


/* Exit if the database is not version 28 or 29. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.28') and (@sDBVersion <> '1.29')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END



/* ---------------------------- */

PRINT 'Step 1 of 9 - Updating permission icons.'

SELECT @ptrval = TEXTPTR(picture) 
FROM ASRSysPermissionCategories
WHERE categoryID = 6

WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000100000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000000000000000000087777778000000008777997880000090877777788000099000077778800099999907777880099999990FFFF880009999990888808000099000077778000000907FFFFFF880000000888888808000000087777778000000008FFFFFF88000000087777777800000000888888880FFFFFFFFFE03FFFFFC01FFFFF400FFFFE400FFFFC000FFFF8000FFFF0000FFFF8000FFFFC000FFFFE400FFFFF400FFFFFC00FFFFFC00FFFFFC00FFFFFE01FFFF

SELECT @ptrval = TEXTPTR(picture) 
FROM ASRSysPermissionCategories
WHERE categoryID = 10

WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000100000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000000000000000000087777778000000008777997880000000877777788000009907777778800999999077777880099999990FFFF880099999908888808000009907777778000000008FFFFFF880000000888888808000000087777778000000008FFFFFF88000000087777777800000000888888880FFFFFFFFFE03FFFFFC01FFFFF400FFFFF000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFF000FFFFF400FFFFFC00FFFFFC00FFFFFC00FFFFFE01FFFF


SELECT @ptrval = TEXTPTR(picture) 
FROM ASRSysPermissionCategories
WHERE categoryID = 20

WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000100000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000000000000000000087777778000000408777997880000440877777788000444444477778800444444447777880444444444FFFF880044444444888808000444444477778000004407FFFFFF880000040888888808000000087777778000000008FFFFFF88000000087777777800000000888888880FFFFFFFFFE03FFFFFC01FFFFF400FFFFE400FFFFC000FFFF8000FFFF0000FFFF8000FFFFC000FFFFE400FFFFF400FFFFFC00FFFFFC00FFFFFC00FFFFFE01FFFF


/* ---------------------------- */

PRINT 'Step 2 of 9 - Amending Audit Table Definition.'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysAuditTrail')
and name = 'ColumnID'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysAuditTrail ADD ColumnID int null

  SELECT @NVarCommand = 'UPDATE ASRSysAuditTrail SET ColumnID = 0'
  EXEC sp_executesql @NVarCommand
END


SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysAuditTrail')
and name = 'Deleted'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysAuditTrail ADD Deleted bit

  SELECT @NVarCommand = 'UPDATE ASRSysAuditTrail SET Deleted = 0'
  EXEC sp_executesql @NVarCommand
END


/* ---------------------------- */

PRINT 'Step 3 of 9 - Amending Export Definitions.'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysExportName')
and name = 'Footer'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysExportName ADD Footer int null
  ALTER TABLE ASRSysExportName ADD FooterText varchar(255) null

  SELECT @NVarCommand = 'UPDATE ASRSysExportName SET Footer = 0, FooterText = '''''
  EXEC sp_executesql @NVarCommand
END


/* ---------------------------- */

PRINT 'Step 4 of 9 - Updating functions.'

DELETE FROM ASRSysFunctions WHERE functionID = 52
DELETE FROM ASRSysFunctions WHERE functionID = 53

INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
       VALUES                (52,'Field Last Change Date',4,0,'Audit','sp_ASRFn_AuditFieldLastChangeDate',0,1)
INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
       VALUES                (53,'Field Changed Between Two Dates',3,0,'Audit','sp_ASRFn_AuditFieldChangedBetweenDates',0, 1)


DELETE FROM ASRSysFunctionParameters WHERE functionID = 52
DELETE FROM ASRSysFunctionParameters WHERE functionID = 53

INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (52, 1, 100, '<Audit Field>')
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (53, 1, 100, '<Audit Field>')
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (53, 3, 4, '<To Date>')
INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       VALUES                         (53, 2, 4, '<From Date>')

/* ---------------------------- */

PRINT 'Step 5 of 9 - Updating permission item information for the Diary.'

DELETE FROM ASRSysPermissionItems WHERE ItemID IN (18, 19)

INSERT  INTO ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey) 
         VALUES (18, 'View and edit manually entered events', 20, 5, 'MANUALEVENTS')

INSERT  INTO ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey) 
         VALUES (19, 'View and edit system generated events', 10, 5, 'SYSTEMEVENTS')

/* ---------------------------- */

PRINT 'Step 6 of 9 - Updating Permissions Stored Procedure'


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsOvernightProcess]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_IsOvernightProcess]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsOvernightProcess
(
    @result integer OUTPUT
)
AS
BEGIN
	DECLARE
		@iCount		integer,
		@sTempExecString	nvarchar(4000),
		@sTempParamDefinition	nvarchar(500),
		@sValue		varchar(8000)

	/* Check if the ''ASRSysSystemSettings'' table exists. */
	SELECT @iCount = count(*)
	FROM sysobjects 
	WHERE name = ''ASRSysSystemSettings''
		
	IF @iCount = 1
	BEGIN
		/* The ASRSysSystemSettings table exists. See if the required records exists in it. */
		SET @sTempExecString = ''SELECT @sValue = settingValue'' +
			'' FROM ASRSysSystemSettings'' +
			'' WHERE section = ''''database'''''' +
			'' AND settingKey = ''''updatingdatedependantcolumns''''''
		SET @sTempParamDefinition = N''@sValue varchar(8000) OUTPUT''
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT
	
		IF NOT @sValue IS NULL
		BEGIN
			SET @result = convert(bit, @sValue)
		END
	END
	ELSE
	BEGIN
		SELECT @result = UpdatingDateDependentColumns FROM ASRSysConfig
	END
END'

exec sp_executesql @NVarCommand


/* ---------------------------- */

PRINT 'Step 7 of 9 - Updating Overnight Email Processing'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASREmailBatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASREmailBatch]

SELECT @NVarCommand = '
CREATE PROCEDURE sp_ASREmailBatch  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue datetime,
		@RecDescID int,
		@RecDesc nvarchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int,
		@blnEnabled int

	/* Clear Servers Inbox */
	/* Doing this just before sending messages means that any failure return messages will */
	/* stay in the servers inbox until this sp is run again - could be useful for support ? */

	SELECT @blnEnabled = SettingValue FROM ASRSysSystemSettings
	WHERE [Section] = ''email'' and [SettingKey] = ''overnight enabled''

	IF @blnEnabled = 0
	BEGIN
		RETURN
	END


	DECLARE @message_id varchar(255)

	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	WHILE not @message_ID is null
	BEGIN
		EXEC master.dbo.xp_deletemail @message_id
		SET @message_id = null
		EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	END


	/* Purge email queue */
	EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue''

	/* Loop through all entries which are to be sent */
	DECLARE emailqueue_cursor
	CURSOR LOCAL FAST_FORWARD FOR 
		SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue
		FROM ASRSysEmailQueue
		WHERE DateSent IS Null And DateDue <= GetDate()
		ORDER BY DateDue
	OPEN emailqueue_cursor
	FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue

	WHILE (@@fetch_status = 0)
	BEGIN

		SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
					 (SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))

		SET @RecDesc = ''''
		SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @RecDesc OUTPUT, @Recordid
		END


		SELECT @sSQL = ''sp_ASREmailSend_'' + convert(varchar,@LinkID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			SELECT @emailDate = getDate()
		             EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''

			IF @hResult = 0
			BEGIN
				UPDATE ASRSysEmailQueue SET DateSent = @emailDate
				WHERE QueueID = @QueueID
			END
		END

		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue
	END
	CLOSE emailqueue_cursor
	DEALLOCATE emailqueue_cursor

END'

exec sp_executesql @NVarCommand



/* ---------------------------- */

PRINT 'Step 8 of 9 - Updating Audit Stored Procedures'


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAudit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAudit]


SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRAudit (
	@piColumnID int,
	@piRecordID int,
	@psRecordDesc varchar(255),
	@psOldValue varchar(255),
	@psNewValue varchar(255))
AS
BEGIN	

	DECLARE @sTableName varchar(8000)
	DECLARE @sColumnName varchar(8000)

	/* Get the table name for the given column. */
	SELECT @sTableName = tablename 
	FROM asrsystables, asrsyscolumns
	WHERE asrsystables.tableid = asrsyscolumns.tableid
	AND asrsyscolumns.columnid = @piColumnID

	/* Get the column name for the given column. */
	SELECT @sColumnName = columnname
	FROM asrsyscolumns
	WHERE asrsyscolumns.columnid = @piColumnID

	IF @sTableName IS NULL SELECT @sTableName = ''<Unknown>''

	/* Insert a record into the Audit Trail table. */
	INSERT INTO ASRSysAuditTrail 
		(userName, dateTimeStamp, tablename, recordID, recordDesc, columnname, oldValue, newValue,ColumnID, Deleted)
	VALUES 
		(user, getDate(), @sTableName, @piRecordID, @psRecordDesc, @sColumnName, @psOldValue, @psNewValue,@piColumnID, 0)


/*	DECLARE @iTableID int

	Get the table ID for the given column. 
	SELECT @iTableID = tableID 
	FROM ASRSysColumns
	WHERE columnID = @piColumnID

	IF @iTableID IS NULL SELECT @iTableID = 0

	 Insert a record into the Audit Trail table. 
	INSERT INTO ASRSysAuditTrail 
		(userName, dateTimeStamp, tableID, recordID, recordDesc, columnID, oldValue, newValue)
	VALUES 
		(user, getDate(), @iTableID, @piRecordID, @psRecordDesc, @piColumnID, @psOldValue, @psNewValue)
*/

END'

exec sp_executesql @NVarCommand



if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAuditLogPurge]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAuditLogPurge]


SELECT @NVarCommand = 'CREATE PROCEDURE [sp_ASRAuditLogPurge] AS

DECLARE @intFrequency int,
                  @strPeriod char(2)

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Users''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Users''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)

BEGIN

  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Permissions''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Permissions''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)

BEGIN
  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Data''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Data''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
BEGIN

  IF @strPeriod = ''dd''
  BEGIN
      UPDATE AsrSysAuditTrail SET Deleted = 1 WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    UPDATE AsrSysAuditTrail SET Deleted = 1 WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    UPDATE AsrSysAuditTrail SET Deleted = 1 WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    UPDATE AsrSysAuditTrail SET Deleted = 1  WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END

END

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Access''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Access''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
BEGIN

  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END'

exec sp_executesql @NVarCommand



if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]


SELECT @NVarCommand = 'CREATE Procedure sp_ASRFn_AuditFieldChangedBetweenDates
(
	@Result bit OUTPUT,
	@ColumnID int,
	@FromDate datetime,
	@ToDate datetime,
	@RecordID int
)

As

declare @Found as int

Begin

	set @Result = 0
		
	set @Found = (Select Count(DateTimeStamp) From ASRSysAuditTrail Where ColumnID = @ColumnID
           		And RecordID = @RecordID
		And DateTimeStamp >= @FromDate And DateTimeStamp <= @ToDate)

	if @found > 0 set @Result = 1

End'

exec sp_executesql @NVarCommand



if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AuditFieldLastChangeDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_AuditFieldLastChangeDate]


SELECT @NVarCommand = 'CREATE Procedure sp_ASRFn_AuditFieldLastChangeDate
(
	@Result datetime OUTPUT,
	@ColumnID int,
	@RecordID int
)

As

Begin

        set @Result = (Select Top 1 DateTimeStamp From ASRSysAuditTrail Where ColumnID = @ColumnID And @RecordID = RecordID)

End'

exec sp_executesql @NVarCommand



if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetAuditTrail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetAuditTrail]


SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRGetAuditTrail (
	@piAuditType	int,
	@psOrder 	varchar(200))
AS
BEGIN
	DECLARE @sSQL varchar(2000)

	IF @piAuditType = 1
	BEGIN
/*		SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
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
*/
		SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
			ASRSysAuditTrail.dateTimeStamp AS [Date / Time], 
			ASRSysAuditTrail.tableName AS [Table], 
			ASRSysAuditTrail.columnName AS [Column], 
			ASRSysAuditTrail.oldValue AS [Old Value], 
			ASRSysAuditTrail.newValue AS [New Value], 
			ASRSysAuditTrail.recordDesc AS [Record Description],
			ASRSysAuditTrail.id
			FROM ASRSysAuditTrail
			WHERE Deleted = 0 ''

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

END'

exec sp_executesql @NVarCommand



/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 9 of 9 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.29')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.29')

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.29 Of HR Pro'
