
/* ----------------------------------------------------- */
/* Update the database from version 2.15 to version 2.16 */
/* ----------------------------------------------------- */

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
        @NVarCommand nvarchar(4000),
	@sColumnDataType varchar(8000),
	@iDateFormat varchar(255),
	@sSQLVersion nvarchar(20)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 2.15 or 2.16. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '2.15') and (@sDBVersion <> '2.16')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 29 - Removing obsolete email stored procedures'

DECLARE @SPName varchar(8000)
DECLARE @SQL varchar(8000)

DECLARE HRProCursor CURSOR
FOR select name from sysobjects where name like 'spASRSysEmailAddr_%' order by name

set nocount on

OPEN HRProCursor
FETCH NEXT FROM HRProCursor INTO @SPName
WHILE @@FETCH_STATUS = 0
BEGIN
	SELECT @SQL = 'DROP PROCEDURE ' + @SPName
	--PRINT @SQL
	EXECUTE sp_sqlexec @SQL
	FETCH NEXT FROM HRProCursor INTO @SPName
END

CLOSE HRProCursor
DEALLOCATE HRProCursor

set nocount off

/* ------------------------------------------------------------- */

PRINT 'Step 2 of 29 - Amending overnight email stored procedure'

if exists (select * from dbo.sysobjects 
where id = object_id(N'[dbo].[spASREmailBatch]') 
and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASREmailBatch]


SELECT @NVarCommand = 'CREATE PROCEDURE spASREmailBatch AS
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

	SELECT @blnEnabled = SettingValue FROM ASRSysSystemSettings
	WHERE [Section] = ''email'' and [SettingKey] = ''overnight enabled''

	IF @blnEnabled = 0
	BEGIN
		RETURN
	END

	-- Clear Servers Inbox
	-- Doing this just before sending messages means that any failure return messages will
	-- stay in the servers inbox until this sp is run again - could be useful for support ?

	-- DECLARE @message_id varchar(255)
	-- EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	-- WHILE not @message_ID is null
	-- BEGIN
	--	EXEC master.dbo.xp_deletemail @message_id
	--	SET @message_id = null
	--	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	-- END


	/* Purge email queue */
	EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue''


	/* Send all emails waiting to be sent regardless of username */
	EXEC spASREmailImmediate ''''

END'

EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 3 of 29 - Adding new system group'


DECLARE @SQLVer varchar(8000)
DECLARE @ASRSysGroup varchar(8000)
--DECLARE @SQL varchar(8000)
DECLARE @UserName varchar(8000)

SET @ASRSysGroup = 'ASRSysGroup'

IF not exists(SELECT Name FROM sysusers WHERE Name = @ASRSysGroup)
BEGIN

  SELECT @SQLVer = substring(@@version,charindex('-',@@version)+2,1)

  IF @SQLVer = 7
  BEGIN
    SET @SQL = 'sp_addrole '''+@ASRSysGroup+''''
    EXECUTE sp_sqlexec @SQL
  END
  ELSE
  BEGIN
    SET @SQL = 'sp_addgroup '''+@ASRSysGroup+''''
    EXECUTE sp_sqlexec @SQL

    SET @SQL = 'GRANT CREATE FUNCTION TO [ASRSysGroup]'
    EXECUTE sp_sqlexec @SQL

  END

  SET @SQL = 'GRANT CREATE PROCEDURE TO [ASRSysGroup]'
  EXECUTE sp_sqlexec @SQL

  SET @SQL = 'GRANT CREATE TABLE TO [ASRSysGroup]'
  EXECUTE sp_sqlexec @SQL

  DECLARE HRProCursor CURSOR
  FOR SELECT sysusers.Name FROM sysusers
      INNER JOIN master..syslogins
      ON master..syslogins.sid = sysusers.sid
      WHERE sysusers.name <> 'dbo'

  OPEN HRProCursor
  FETCH NEXT FROM HRProCursor INTO @UserName
  WHILE @@FETCH_STATUS = 0
  BEGIN
    SELECT @SQL = 'sp_addrolemember '''+@ASRSysGroup+''', '''+@UserName+''''
    EXECUTE sp_sqlexec @SQL
    FETCH NEXT FROM HRProCursor INTO @UserName
  END

  CLOSE HRProCursor
  DEALLOCATE HRProCursor

END


/* ------------------------------------------------------------- */
PRINT 'Step 4 of 29 - Update last audit change date'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AuditFieldLastChangeDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRFn_AuditFieldLastChangeDate]


	SELECT @NVarCommand = 'CREATE Procedure sp_ASRFn_AuditFieldLastChangeDate
		(
			@Result datetime OUTPUT,
			@ColumnID int,
			@RecordID int
		)
		
		As
		
		Begin
		
		        set @Result = (Select Top 1 DateTimeStamp From ASRSysAuditTrail Where ColumnID = @ColumnID And @RecordID = RecordID order by DateTimeStamp desc)
		
		End'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 5 of 29 - Update last audit field changed between two dates'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
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
				And DateTimeStamp >= @FromDate  And DateTimeStamp <= @ToDate+1)
		
			if @found > 0 set @Result = 1
		
		End'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 6 of 29 - Adding new columns to Column Definition (Quick Address)'

	/* Add newQuick Address columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'QAddressEnabled'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[QAddressEnabled] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [QAddressEnabled] = 0'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 7 of 29 - Clear any null values in the database'
SET @NVarCommand = 'UPDATE ASRSysColumns SET [QAddressEnabled] = 0 WHERE [QAddressEnabled] is null'
EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 8 of 29 - Adding new columns to Column Definition (Quick Address) - definition'

	/* Add newQuick Address columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'QAAddress'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[QAIndividual] [bit] NULL,
					[QAAddress] [int] NULL,
					[QAProperty] [int] NULL,
					[QAStreet] [int] NULL,
					[QALocality] [int] NULL,
					[QATown] [int] NULL,
					[QACounty] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [QAIndividual] = 0, 
					[QAAddress] = 0,
					[QAProperty] = 0,
					[QAStreet] = 0,
					[QALocality] = 0,
					[QATown] = 0,
					[QACounty] = 0'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 9 of 29 - Adding new columns to Column Definition (Lookup Filter Operator)'

	/* Add new LookupFilterOperator columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'LookupFilterOperator'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[LookupFilterOperator] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns
					SET lookupFilterOperator =
						CASE
							WHEN lookupFilterColumnID > 0 THEN
								CASE
									WHEN (dataType = 11) THEN 7
									WHEN (dataType = -7) OR (dataType = 2) OR (dataType = 4) THEN 1
									WHEN (dataType = 12) OR (dataType = -3) OR (dataType = -1) THEN 14
									ELSE 0
								END
							ELSE 0
						END
					WHERE lookupFilterOperator IS null'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 10 of 29 - Amending Table Permissions Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRAllTablePermissions]


	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].sp_ASRAllTablePermissions 
			AS
			BEGIN
				/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
				DECLARE @iUserGroupID	int

				/* Initialise local variables. */
			--{
			-- MH20040106 Fault 5627 - Ignore ASRSysGroup
			--	SELECT @iUserGroupID = sysusers.gid
			--	FROM sysusers
			--	WHERE sysusers.name = CURRENT_USER
				SELECT @iUserGroupID = usg.gid
				FROM sysusers usu
				left outer join
				(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
				WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
					(usg.issqlrole = 1 or usg.uid is null) and
					usu.name = CURRENT_USER AND not (usg.name like ''ASRSys%'')
			--}


				SELECT sysobjects.name, sysprotects.action
				FROM sysprotects 
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				WHERE sysprotects.uid = @iUserGroupID
					AND sysprotects.protectType <> 206
					AND sysprotects.action <> 193
					AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
				UNION
				SELECT sysobjects.name, 193
				FROM syscolumns
				INNER JOIN sysprotects ON (syscolumns.id = sysprotects.id
					AND sysprotects.action = 193 
					AND sysprotects.uid = @iUserGroupID
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				WHERE syscolumns.name = ''timestamp''
					AND ((sysprotects.protectType = 205) 
					OR (sysprotects.protectType = 204))
				ORDER BY sysobjects.name
			END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 11 of 29 - Amending Module Permissions Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRIsSysSecMgr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRIsSysSecMgr]


	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRIsSysSecMgr (
				@psGroupName		sysname,
				@pfSysSecMgr		bit	OUTPUT
			)
			AS
			BEGIN
				DECLARE @iUserGroupID integer

				/* Get the current user''s group ID. */
			--{
			-- MH20040106 Fault 5627 - Ignore ASRSysGroup
			--	SELECT @iUserGroupID = sysusers.gid
			--	FROM sysusers
			--	WHERE sysusers.name = @psGroupName

				SELECT @iUserGroupID = usg.gid
				FROM sysusers usu
				left outer join
				(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
				WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
					(usg.issqlrole = 1 or usg.uid is null) and
					usu.name = @psGroupName AND not (usg.name like ''ASRSys%'')
			--}


				SELECT @pfSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
				FROM ASRSysGroupPermissions
				INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
				INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
				WHERE sysusers.uid = @iUserGroupID
					AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER'' OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
					AND ASRSysGroupPermissions.permitted = 1
					AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''
			END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 12 of 29 - Amending Screen Controls Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetControlDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRGetControlDetails]


	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].sp_ASRGetControlDetails 
		(
			@piScreenID int
		)
		AS
		BEGIN
			SELECT ASRSysControls.*, 
				ASRSysColumns.columnName, 
				ASRSysColumns.columnType, 
				ASRSysColumns.datatype,
				ASRSysColumns.defaultValue,
				ASRSysColumns.size, 
				ASRSysColumns.decimals, 
				ASRSysColumns.lookupTableID, 
				ASRSysColumns.lookupColumnID, 
				ASRSysColumns.lookupFilterColumnID, 
				ASRSysColumns.lookupFilterOperator, 
				ASRSysColumns.lookupFilterValueID, 
				ASRSysColumns.spinnerMinimum, 
				ASRSysColumns.spinnerMaximum, 
				ASRSysColumns.spinnerIncrement, 
				ASRSysColumns.mandatory, 
				ASRSysColumns.uniquecheck,
				ASRSysColumns.convertcase, 
				ASRSysColumns.mask, 
				ASRSysColumns.blankIfZero, 
				ASRSysColumns.multiline, 
				ASRSysColumns.alignment AS colAlignment, 
				ASRSysColumns.calcExprID, 
				ASRSysColumns.gotFocusExprID, 
				ASRSysColumns.lostFocusExprID, 
				ASRSysColumns.dfltValueExprID, 
				ASRSysColumns.calcTrigger, 
				ASRSysColumns.readOnly, 
				ASRSysColumns.statusBarMessage, 
				ASRSysColumns.errorMessage, 
				ASRSysColumns.linkTableID,
				ASRSysColumns.linkViewID,
				ASRSysColumns.linkOrderID,
				ASRSysColumns.Afdenabled,
				ASRSysColumns.OleOnServer,
				ASRSysTables.TableName,
				ASRSysColumns.Trimming,
				ASRSysColumns.Use1000Separator,
				ASRSysColumns.QAddressEnabled
			FROM ASRSysControls 
			LEFT OUTER JOIN ASRSysTables 
				ON ASRSysControls.tableID = ASRSysTables.tableID 
			LEFT OUTER JOIN ASRSysColumns 
				ON ASRSysColumns.tableID = ASRSysControls.tableID 
					AND ASRSysColumns.columnID = ASRSysControls.columnID
			WHERE ASRSysControls.ScreenID = @piScreenID
			ORDER BY ASRSysControls.PageNo, 
				ASRSysControls.ControlLevel DESC, 
				ASRSysControls.tabIndex
		END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 29 - Add new  Hierarchy functions'

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 65'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (65, ''Is Post Subordinate Of'', 3, 0, ''Personnel'', ''spASRSysFnIsPostSubordinateOf'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 65'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (65, 1, 0, ''<Identifier>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 66'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (66, ''Is Post Subordinate Of User'', 3, 0, ''Personnel'', ''spASRSysFnIsPostSubordinateOfUser'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 66'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (66, 1, 1, ''<Login>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (66, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 67'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (67, ''Is Personnel Subordinate Of'', 3, 0, ''Personnel'', ''spASRSysFnIsPersonnelSubordinateOf'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 67'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (67, 1, 0, ''<Identifier>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (67, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 68'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (68, ''Is Personnel Subordinate Of User'', 3, 0, ''Personnel'', ''spASRSysFnIsPersonnelSubordinateOfUser'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 68'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (68, 1, 1, ''<Login>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (68, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 69'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (69, ''Has Post Subordinate'', 3, 0, ''Personnel'', ''spASRSysFnHasPostSubordinate'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 69'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (69, 1, 0, ''<Identifier>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 70'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (70, ''Has Post Subordinate User'', 3, 0, ''Personnel'', ''spASRSysFnHasPostSubordinateUser'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 70'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (70, 1, 1, ''<Login>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (70, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 71'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (71, ''Has Personnel Subordinate'', 3, 0, ''Personnel'', ''spASRSysFnHasPersonnelSubordinate'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 71'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (71, 1, 0, ''<Identifier>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (71, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 72'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (72, ''Has Personnel Subordinate User'', 3, 0, ''Personnel'', ''spASRSysFnHasPersonnelSubordinateUser'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 72'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (72, 1, 1, ''<Login>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (72, 2, 4, ''<Date>'')'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 14 of 29 - Security Settings'

if not exists (select * from asrsyssystemsettings where Section = 'Misc' and SettingKey = 'cfg_pcl')
begin
	insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	values('misc', 'cfg_pcl', '1')

	insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	values('misc', 'cfg_ba', '3')

	insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	values('misc', 'cfg_ld', '300')
end

if not exists (select * from asrsyssystemsettings where Section = 'Misc' and SettingKey = 'cfg_rt')
begin
	insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	values('misc', 'cfg_rt', '3600')
end


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 29 - Update output options'

--Make sure that "Preview on screen" is unchecked for "Data Only"
UPDATE ASRSysCustomReportsName SET OutputPreview = 0 WHERE OutputFormat = 0
UPDATE ASRSysCrossTab SET OutputPreview = 0 WHERE OutputFormat = 0

--No longer allow Pivot table output for Bradford Factor'
UPDATE ASRSysSystemSettings SET SettingValue = 4
WHERE [Section] = 'bradfordfactor' AND SettingKey = 'format' AND SettingValue = 6


/* ------------------------------------------------------------- */
PRINT 'Step 16 of 29 - Drop Obsolete Table'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysVersionInformation]') and OBJECTPROPERTY(id, N'IsTable') = 1)
	drop table [dbo].[ASRSysVersionInformation]


/* ------------------------------------------------------------- */
PRINT 'Step 17 of 29 - Drop Obsolete Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetUserGroups]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRGetUserGroups]


/* ------------------------------------------------------------- */
PRINT 'Step 18 of 29 - Drop Obsolete Rule'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rulAccess]'))
	drop rule dbo.rulAccess

/* ------------------------------------------------------------- */
PRINT 'Step 19 of 29 - Replace immediate email procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASREmailImmediate]


	SELECT @NVarCommand = 'CREATE PROCEDURE spASREmailImmediate(@Username varchar(255))  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue varchar(8000),
		@RecDescID int,
		@RecDesc varchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int,
		@blnEnabled int,
		@RecalculateRecordDesc bit,
		@TableID int,
		@RecipTo varchar(4000),
		@TempText nvarchar(4000)

	/* Loop through all entries which are to be sent */
	DECLARE emailqueue_cursor
	CURSOR LOCAL FAST_FORWARD FOR 
		SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID
		FROM ASRSysEmailQueue
		WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
		And (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
		ORDER BY DateDue

	OPEN emailqueue_cursor
	FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

	WHILE (@@fetch_status = 0)
	BEGIN

		IF @RecalculateRecordDesc = 1
			BEGIN	
				IF @ColumnID > 0
					BEGIN
						SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
							 (SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))
					END
				ELSE IF @TableID > 0
					BEGIN			
						SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = @TableID)
					END
		
				SET @RecDesc = ''''
				SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					EXEC @sSQL @RecDesc OUTPUT, @Recordid
				END
			END

		/* Add table name to record descripion if it is a table entry */
		IF @TableID > 0
			BEGIN
				SELECT @TempText = (SELECT TableName FROM ASRSYSTables WHERE TableID = @TableID)
				SET @RecDesc = @TempText + '' : '' + @RecDesc
			END		
	
		IF @ColumnID > 0
			BEGIN
				SELECT @sSQL = ''spASRSysEmailSend_'' + convert(varchar,@LinkID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SELECT @emailDate = getDate()
						EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''
					END
			END
		ELSE IF @TableID > 0
			BEGIN
				SET @sSQL = ''spASRSysEmailAddr''
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SELECT @emailDate = getDate()
						EXEC @hResult = @sSQL @RecipTo OUTPUT, @LinkID, 0
						EXEC @hResult = master.dbo.xp_sendmail  @recipients=@RecipTo,  @subject=@columnvalue,  @message=@RecDesc, @no_output=''True''
					END
			END

		IF @hResult = 0
		BEGIN
			UPDATE ASRSysEmailQueue SET DateSent = @emailDate
			WHERE QueueID = @QueueID
		END

		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

	END
	CLOSE emailqueue_cursor
	DEALLOCATE emailqueue_cursor

END'

	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */

PRINT 'Step 20 of 29 - Updating Get Messages Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetMessages]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRGetMessages]


	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRGetMessages AS
	BEGIN
	DECLARE @iDBID	integer,
		@iID		integer,
		@dtLoginTime	datetime,
		@sLoginName	varchar(256),
		@iCount	integer,
		@Realspid integer

	--MH20040211 Fault 8062
	--{
	--Need to get spid of parent process
	SELECT @Realspid = a.spid
	FROM master..sysprocesses a
	FULL OUTER JOIN master..sysprocesses b
		ON a.hostname = b.hostname
		AND a.hostprocess = b.hostprocess
		AND a.spid <> b.spid
	WHERE b.spid = @@Spid

	--If there is no parent spid then use current spid
	IF @Realspid is null SET @Realspid = @@spid
	--}


	/* Get the current user''s process information. */
	SELECT @iDBID = dbID,
		@dtLoginTime = login_time,
		@sLoginName = loginame
	FROM master..sysprocesses
	WHERE spid = @Realspid


	/* Return the recordset of messages. */
	SELECT ''Message from user '''''' + ltrim(rtrim(messageFrom)) + 
		'''''' using '' + ltrim(rtrim(messageSource)) + 
		'' ('' + convert(varchar(100), messageTime, 100) +'')'' + 
		char(10) + message
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime
		AND spid = @Realspid

	/* Remove any messages that have just been picked up. */
	DELETE
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime
		AND spid = @Realspid

	/* Remove any orphaned messages. */
	/* NB. This is done via a cursor to avoid any possible collation conflict between ASRSysMessages.loginName and sysprocesses.loginame. */
	DECLARE messages_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT id,
		loginName, 
		dbID, 
		loginTime 
	FROM ASRSysMessages
	OPEN messages_cursor
	FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM master..sysprocesses
		WHERE loginame =  @sLoginName
			AND dbID = @iDBID
			AND login_time = @dtLoginTime

		IF @iCount = 0
		BEGIN
			DELETE FROM ASRSysMessages 
			WHERE id = @iID
		END
			
		FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	END
	CLOSE messages_cursor 
	DEALLOCATE messages_cursor 
	END'

	EXEC sp_executesql @NVarCommand
	
/* ------------------------------------------------------------- */

PRINT 'Step 21 of 29 - Updating Send Messages Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRSendMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRSendMessage]


	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRSendMessage 
		(
			@psMessage	varchar(8000)
		)
		AS
		BEGIN
			DECLARE @iDBid	integer,
				@iSPid		integer,
				@iUid		integer,
				@sLoginName	varchar(256),
				@dtLoginTime	datetime, 
				@sCurrentUser	varchar(256),
				@sCurrentApp	varchar(256),
				@Realspid integer

			--MH20040224 Fault 8062
			--{
			--Need to get spid of parent process
			SELECT @Realspid = a.spid
			FROM master..sysprocesses a
			FULL OUTER JOIN master..sysprocesses b
				ON a.hostname = b.hostname
				AND a.hostprocess = b.hostprocess
				AND a.spid <> b.spid
			WHERE b.spid = @@Spid

			--If there is no parent spid then use current spid
			IF @Realspid is null SET @Realspid = @@spid
			--}


			/* Get the process information for the current user. */
			SELECT @iDBid = dbid, 
				@sCurrentUser = loginame,
				@sCurrentApp = program_name
			FROM master..sysprocesses
			WHERE spid = @@spid

			/* Get a cursor of the other logged in HR Pro users. */
			DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT spid, loginame, uid, login_time
				FROM master..sysprocesses
				WHERE program_name LIKE ''HR Pro%''
				AND dbid = @iDBid
				AND (spid <> @@spid and spid <> @Realspid)

			OPEN logins_cursor
			FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Create a message record for each HR Pro user. */
				INSERT INTO ASRSysMessages 
					(loginname, message, loginTime, dbid, uid, spid, messageTime, messageFrom, messageSource) 
					VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, getdate(), @sCurrentUser, @sCurrentApp)

				FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
			END
			CLOSE logins_cursor
			DEALLOCATE logins_cursor
		END'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */

PRINT 'Step 22 of 29 - Update Email Address Sequential Numbering'

DECLARE @MaxEmailID integer
if not exists (select * from asrsyssystemsettings where Section = 'autoid' and SettingKey = 'emailaddress')
begin
  SELECT @MaxEmailID = MAX(EmailID) FROM ASRSysEmailAddress 
  INSERT INTO ASRSysSystemSettings (Section, SettingKey, SettingValue)
  VALUES ('autoid', 'emailaddress', @MaxEmailID)
end

/* ------------------------------------------------------------- */

PRINT 'Step 23 of 29 - Removing obsolete Hierarchy functions'

DELETE FROM ASRSysFunctions WHERE functionID IN (65,67,69,71)
DELETE FROM ASRSysFunctionParameters WHERE functionID in (65,66,67,68,69,70,71,72)


/* ------------------------------------------------------------- */

PRINT 'Step 24 of 29 - Updating Hierarchy functions'

UPDATE ASRSysFunctions SET functionName = 'Is Post That Reports To Current User' WHERE functionID = 66
UPDATE ASRSysFunctions SET functionName = 'Is Personnel That Reports To Current User' WHERE functionID = 68
UPDATE ASRSysFunctions SET functionName = 'Is Post That Current User Reports To' WHERE functionID = 70
UPDATE ASRSysFunctions SET functionName = 'Is Personnel That Current User Reports To' WHERE functionID = 72


/* ------------------------------------------------------------- */
PRINT 'Step 25 of 29 - New system permissions stored procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRSystemPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRSystemPermission]


	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRSystemPermission
(
	@pfPermissionGranted 	bit OUTPUT,
	@psCategoryKey	varchar(50),
	@psPermissionKey	varchar(50),
	@psSQLLogin 		varchar(200)
)
AS
BEGIN
	
	-- Return 1 if the given permission is granted to the current user, 0 if it is not.

	DECLARE @fGranted bit
	DECLARE @sGroupName varchar(8000)

	-- Is logged in user a system administrator
	SELECT @fGranted = sysAdmin FROM master..syslogins WHERE loginname = @psSQLLogin

	IF @fGranted = 0
	BEGIN
		SELECT @sGroupName = usg.name
		FROM sysusers usu
		left outer join
		(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
		WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
			(usg.issqlrole = 1 or usg.uid is null) and
			usu.name = @psSQLLogin AND not (usg.name like ''ASRSys%'')

		SELECT @fGranted = ASRSysGroupPermissions.permitted
		FROM ASRSysGroupPermissions
			INNER JOIN ASRSysPermissionItems 
				ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
			INNER JOIN ASRSysPermissionCategories
				ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID
		WHERE ASRSysPermissionItems.itemKey = @psPermissionKey
			AND ASRSysGroupPermissions.groupName = @sGroupName
			AND ASRSysPermissionCategories.categoryKey = @psCategoryKey
	END


	IF @fGranted IS NULL
	BEGIN
		SET @fGranted = 0
	END

	SET @pfPermissionGranted = @fGranted

	END'

	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'GRANT EXEC ON [sp_ASRSystemPermission] TO [ASRSysGroup]'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 26 of 29 - New read all table permissions procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAllTablePermissions]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].sp_ASRAllTablePermissions 
	(
	@psSQLLogin 		varchar(200)
	)
	AS
	BEGIN
		/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
		DECLARE @iUserGroupID	int

		/* Initialise local variables. */
		SELECT @iUserGroupID = usg.gid
		FROM sysusers usu
		left outer join
		(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
		WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
			(usg.issqlrole = 1 or usg.uid is null) and
			usu.name = @psSQLLogin AND not (usg.name like ''ASRSys%'')

		SELECT sysobjects.name, sysprotects.action
		FROM sysprotects 
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		WHERE sysprotects.uid = @iUserGroupID
			AND sysprotects.protectType <> 206
			AND sysprotects.action <> 193
			AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
		UNION
		SELECT sysobjects.name, 193
		FROM syscolumns
		INNER JOIN sysprotects ON (syscolumns.id = sysprotects.id
			AND sysprotects.action = 193 
			AND sysprotects.uid = @iUserGroupID
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		WHERE syscolumns.name = ''timestamp''
			AND ((sysprotects.protectType = 205) 
			OR (sysprotects.protectType = 204))
		ORDER BY sysobjects.name
	END'

	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 27 of 29 - Cleanup to drop temporary UDFs'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRDropTempObjects]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASRDropTempObjects]


	SELECT @NVarCommand = 'CREATE PROCEDURE spASRDropTempObjects
		AS
		BEGIN

			DECLARE	@sObjectName varchar(2000),
					@sUsername varchar(2000),
					@sXType varchar(50)
						
			DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
			SELECT [dbo].[sysobjects].[name], [dbo].[sysusers].[name], [dbo].[sysobjects].[xtype]
			FROM [dbo].[sysobjects] 
					INNER JOIN [dbo].[sysusers]
					ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
			WHERE LOWER([dbo].[sysusers].[name]) != ''dbo''
					AND (OBJECTPROPERTY(id, N''IsUserTable'') = 1
						OR OBJECTPROPERTY(id, N''IsProcedure'') = 1
						OR OBJECTPROPERTY(id, N''IsTableFunction'') = 1)

			OPEN tempObjects
			FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
			WHILE (@@fetch_status <> -1)
			BEGIN		
				IF UPPER(@sXType) = ''U''
					-- user table
					BEGIN
						EXEC (''DROP TABLE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''P''
					-- procedure
					BEGIN
						EXEC (''DROP PROCEDURE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''TF''
					-- UDF
					BEGIN
						EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				
				FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
				
			END
			CLOSE tempObjects
			DEALLOCATE tempObjects
			
			EXEC (''DELETE FROM [dbo].[ASRSysSQLObjects]'')

		END'

	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 28 of 29 - New Case Sensitive Comparison stored procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRCaseSensitiveCompare]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCaseSensitiveCompare]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRCaseSensitiveCompare
(
	@pfResult		bit OUTPUT,
	@psStringA 		varchar(8000),
	@psStringB		varchar(8000)
)
AS
BEGIN
	/* Return 1 if the given string are exactly equal. */
	DECLARE @iPosition	integer

	SET @pfResult = 0

	IF (@psStringA IS NULL) AND (@psStringB IS NULL) SET @pfResult = 1

	IF (@pfResult = 0) AND (NOT @psStringA IS NULL) AND (NOT @psStringB IS NULL)
	BEGIN

		/* LEN() does not look at trailing spaces, so force it too by adding some quotations at the end. */
		SET @psStringA = @psStringA + ''''''''
		SET @psStringB = @psStringB + ''''''''

		IF LEN(@psStringA) = LEN(@psStringB)
		BEGIN
			SET @pfResult = 1

			SET @iPosition = 1
			WHILE @iPosition <= LEN(@psStringA) 
			BEGIN
				IF ASCII(SUBSTRING(@psStringA, @iPosition, 1)) <> ASCII(SUBSTRING(@psStringB, @iPosition, 1))
				BEGIN
					SET @pfResult = 0
					BREAK
				END

				SET @iPosition = @iPosition + 1
			END
		END
	END
END'

	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 29 of 29 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.16')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.16')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.16')

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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.16 Of HR Pro'
