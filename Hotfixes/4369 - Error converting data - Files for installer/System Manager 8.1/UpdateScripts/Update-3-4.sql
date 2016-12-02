
/* --------------------------------------------------- */
/* Update the database from version 3.3 to version 3.4 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@iSQLVersion numeric(3,1),
	@Command varchar(8000),
	@NVarCommand nvarchar(4000),
    @sObject sysname,
    @sObjectType char(2)

DECLARE @sSQL varchar(8000)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
DECLARE @sSPCode_5 nvarchar(4000)
DECLARE @sSPCode_6 nvarchar(4000)
DECLARE @sSPCode_7 nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON
SET @DBName = DB_NAME()
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 3.3 or 3.4. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '3.3') and (@sDBVersion <> '3.4')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
PRINT 'Step 1 of X - Populating Email configuration'

  DECLARE @SPText varchar(8000)
  DECLARE @Temp varchar(8000)
  DECLARE @method varchar(8000)
  DECLARE @server varchar(8000)
  DECLARE @account varchar(8000)
  DECLARE @qainfo varchar(8000)

  DECLARE @start int
  DECLARE @finish int

  DECLARE    HRProCursor
  CURSOR FOR
  SELECT     syscomments.text
  FROM       syscomments
  INNER JOIN sysobjects
          ON syscomments.id = sysobjects.id
  WHERE      sysobjects.name = 'spASRSendMail'
  ORDER BY   syscomments.colid

  OPEN  HRProCursor
  FETCH NEXT
  FROM  HRProCursor
  INTO  @Temp

  SET @SPText = ''
  WHILE @@FETCH_STATUS = 0
  BEGIN
    SET @SPText = @SPText + @Temp

    FETCH NEXT
    FROM  HRProCursor
    INTO  @Temp
  END

  CLOSE HRProCursor
  DEALLOCATE HRProCursor

  if charindex('xp_SMTPsendmail80',@sptext) > 0
    set @method = '3'

    --SMTP Server
    SET @start = charindex('@address',@sptext)
    if @start > 0
    BEGIN
      SET @start = charindex('''',@sptext,@start)+1
      SET @finish = charindex('''',@sptext,@start)
      if @start > 0 and @finish > 0
        SET @server = substring(@sptext,@start,@finish-@start)
    END

  --SMTP Mailbox
  SET @start = charindex('@from',@sptext)
  if @start > 0
  BEGIN
    SET @start = charindex('''',@sptext,@start)+1
    SET @finish = charindex('''',@sptext,@start)
    if @start > 0 and @finish > 0
      SET @account = substring(@sptext,@start,@finish-@start)
  END

  else if charindex('sp_send_dbmail',@sptext) > 0
    set @method = '2'
  else if charindex('xp_sendmail',@sptext) > 0
    set @method = '1'
  else
    set @method = '0'

  if charindex('QA Info: ',@sptext) > 0
    set @qainfo = '1'
  else
    set @qainfo = '0'


  delete from asrsyssystemsettings
  where [Section] = 'email' and [SettingKey] = 'method'
  insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
  values('email', 'method', @method)

  delete from asrsyssystemsettings
  where [Section] = 'email' and [SettingKey] = 'server'
  insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
  values('email', 'server', @server)

  delete from asrsyssystemsettings
  where [Section] = 'email' and [SettingKey] = 'account'
  insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
  values('email', 'account', @account)

  delete from asrsyssystemsettings
  where [Section] = 'email' and [SettingKey] = 'qa info'
  insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
  values('email', 'qa info', @qainfo)


/* ------------------------------------------------------------- */
PRINT 'Step 2 of X - Creating/modifying Workflow tables'

	/* ASRSysWorkflowElements - Add new DataRecordTable column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'DataRecordTable'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						DataRecordTable [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.dataRecordTable =
			isnull(CASE
				WHEN ASRSysWorkflowElements.dataRecord = 0 THEN -- Initiator
					(SELECT TOP 1 convert(integer, ISNULL(MS1.parameterValue, ''0''))
					FROM ASRSysModuleSetup MS1
					WHERE MS1.moduleKey = ''MODULE_PERSONNEL''
						AND MS1.parameterKey = ''Param_TablePersonnel'')

 				WHEN ASRSysWorkflowElements.dataRecord = 1 THEN -- Identified
					(SELECT TOP 1
						CASE 
							WHEN WE1.type = 2 THEN -- WebForm
								(SELECT TOP 1 WEI1.tableID
								FROM ASRSysWorkflowElementItems WEI1
								WHERE WEI1.elementID = WE1.ID
									AND WEI1.identifier = ASRSysWorkflowElements.recSelIdentifier)
							WHEN WE1.type = 5 THEN -- StoredData
								WE1.dataTableID 
							ELSE 0
						END
					FROM ASRSysWorkflowElements WE1
					WHERE WE1.identifier = ASRSysWorkflowElements.recSelWebFormIdentifier
					AND WE1.workflowID = ASRSysWorkflowElements.workflowID)

				WHEN ASRSysWorkflowElements.dataRecord = 4 THEN -- Triggered
					(SELECT TOP 1 WF1.baseTable
					FROM ASRSysWorkflows WF1
					WHERE WF1.ID = ASRSysWorkflowElements.workflowID)

				ELSE 0
			END, 0)
		WHERE ASRSysWorkflowElements.type = 5 -- StoredData elements'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new SecondaryDataRecordTable column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'SecondaryDataRecordTable'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						SecondaryDataRecordTable [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.secondaryDataRecordTable =
			isnull(CASE
				WHEN ASRSysWorkflowElements.secondaryDataRecord = 0 THEN -- Initiator
					(SELECT TOP 1 convert(integer, ISNULL(MS2.parameterValue, ''0''))
					FROM ASRSysModuleSetup MS2
					WHERE MS2.moduleKey = ''MODULE_PERSONNEL''
						AND MS2.parameterKey = ''Param_TablePersonnel'')

 				WHEN ASRSysWorkflowElements.secondaryDataRecord = 1 THEN -- Identified
					(SELECT TOP 1
						CASE 
							WHEN WE2.type = 2 THEN -- WebForm
								(SELECT TOP 1 WEI2.tableID
								FROM ASRSysWorkflowElementItems WEI2
								WHERE WEI2.elementID = WE2.ID
									AND WEI2.identifier = ASRSysWorkflowElements.secondaryRecSelIdentifier)
							WHEN WE2.type = 5 THEN -- StoredData
								WE2.dataTableID 
							ELSE 0
						END
					FROM ASRSysWorkflowElements WE2
					WHERE WE2.identifier = ASRSysWorkflowElements.secondaryRecSelWebFormIdentifier
					AND WE2.workflowID = ASRSysWorkflowElements.workflowID)

				WHEN ASRSysWorkflowElements.secondaryDataRecord = 4 THEN -- Triggered
					(SELECT TOP 1 WF2.baseTable
					FROM ASRSysWorkflows WF2
					WHERE WF2.ID = ASRSysWorkflowElements.workflowID)

				ELSE 0
			END, 0)
		WHERE ASRSysWorkflowElements.type = 5 -- StoredData elements'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new RecordTableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'RecordTableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						RecordTableID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.RecordTableID =
			isnull(CASE
				WHEN ASRSysWorkflowElementItems.DBRecord = 0 THEN -- Initiator
					(SELECT TOP 1 convert(integer, ISNULL(MS1.parameterValue, ''0''))
					FROM ASRSysModuleSetup MS1
					WHERE MS1.moduleKey = ''MODULE_PERSONNEL''
						AND MS1.parameterKey = ''Param_TablePersonnel'')

 				WHEN ASRSysWorkflowElementItems.DBRecord = 1 THEN -- Identified
					(SELECT TOP 1
						CASE 
							WHEN WE1.type = 2 THEN -- WebForm
								(SELECT TOP 1 WEI1.tableID
								FROM ASRSysWorkflowElementItems WEI1
								WHERE WEI1.elementID = WE1.ID
									AND WEI1.identifier = ASRSysWorkflowElementItems.WFValueIdentifier)
							WHEN WE1.type = 5 THEN -- StoredData
								WE1.dataTableID 
							ELSE 0
						END
					FROM ASRSysWorkflowElements WE1
					WHERE WE1.identifier = ASRSysWorkflowElementItems.WFFormIdentifier
					AND WE1.workflowID = (SELECT ASRSysWorkflowElements.workflowID
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID))

				WHEN ASRSysWorkflowElementItems.DBRecord = 4 THEN -- Triggered
					(SELECT TOP 1 WF1.baseTable
					FROM ASRSysWorkflows WF1
					WHERE WF1.ID = (SELECT ASRSysWorkflowElements.workflowID
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID))

				ELSE 0
			END, 0)
		WHERE ASRSysWorkflowElementItems.itemType = 11 -- RecSel items'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new parent1TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'parent1TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						parent1TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new parent1RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'parent1RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						parent1RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new parent2TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'parent2TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						parent2TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new parent2RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'parent2RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						parent2RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new emailID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'EmailID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						EmailID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceValues
		SET ASRSysWorkflowInstanceValues.EmailID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstances - Add new parent1TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
	and name = 'parent1TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD 
						parent1TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstances - Add new parent1RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
	and name = 'parent1RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD 
						parent1RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstances - Add new parent2TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
	and name = 'parent2TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD 
						parent2TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstances - Add new parent2RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
	and name = 'parent2RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD 
						parent2RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueue - Add new parent1TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueue')
	and name = 'parent1TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueue ADD 
						parent1TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueue - Add new parent1RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueue')
	and name = 'parent1RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueue ADD 
						parent1RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueue - Add new parent2TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueue')
	and name = 'parent2TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueue ADD 
						parent2TableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueue - Add new parent2RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueue')
	and name = 'parent2RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueue ADD 
						parent2RecordID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueue - Add new instanceID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueue')
	and name = 'InstanceID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueue ADD 
						InstanceID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowQueue
		SET ASRSysWorkflowQueue.InstanceID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowQueueColumns - Add new emailID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowQueueColumns')
	and name = 'EmailID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowQueueColumns ADD 
						EmailID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowQueueColumns
		SET ASRSysWorkflowQueueColumns.EmailID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new TrueFlowType column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'TrueFlowType'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						TrueFlowType [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.TrueFlowType = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new TrueFlowExprID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'TrueFlowExprID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						TrueFlowExprID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.TrueFlowExprID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new DescriptionExprID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'DescriptionExprID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						DescriptionExprID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.DescriptionExprID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new WebFormFGColor column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'WebFormFGColor'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						WebFormFGColor [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.WebFormFGColor = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new Orientation column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'Orientation'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						Orientation [smallint] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.Orientation = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new RecordOrderID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'RecordOrderID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						RecordOrderID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.RecordOrderID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new RecordFilterID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'RecordFilterID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						RecordFilterID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.RecordFilterID = 0'

		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Modifying Expression tables'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExpressions')
	and name = 'UtilityID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExpressions ADD 
						UtilityID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysExpressions
		SET ASRSysExpressions.UtilityID = 0'

		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExprComponents')
	and name = 'WorkflowElement'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExprComponents ADD 
						WorkflowElement [varchar](200) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExprComponents')
	and name = 'WorkflowItem'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExprComponents ADD 
						WorkflowItem [varchar](200) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExprComponents')
	and name = 'WorkflowRecord'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExprComponents ADD 
						WorkflowRecord [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExprComponents')
	and name = 'WorkflowRecordTableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExprComponents ADD 
						WorkflowRecordTableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 4 of X - Creating/modifying Workflow stored procedures'

	----------------------------------------------------------------------
	-- spASRColumnsUsedInExpression
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRColumnsUsedInExpression]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRColumnsUsedInExpression]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRColumnsUsedInExpression]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRColumnsUsedInExpression
		(
			@piExpressionID		integer,
			@pcurColumns		cursor varying output
		)
		AS
		BEGIN
			-- Return the IDs of the columns used in the given expression.
			DECLARE
				@iType			integer,
				@iID			integer,
				@iExprID		integer,
				@iColumnID		integer,
				@curSubColumns	cursor
		
			CREATE TABLE #curColumns (columnID integer)
		
			-- Record the columns used by field components.
			INSERT INTO #curColumns
			SELECT DISTINCT EC.fieldColumnID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID = @piExpressionID
				AND EC.type = 1 -- Field component
		
			-- Check sub-expressions.
			DECLARE curSubExpressions CURSOR LOCAL FAST_FORWARD FOR 
			SELECT
				EC.type, 
				CASE 
					WHEN EC.type = 1 THEN EC.fieldSelectionFilter -- Field filter
					WHEN EC.type = 2 THEN EC.componentID -- Function
					WHEN EC.type = 3 THEN EC.calculationID -- Calculation
					WHEN EC.type = 10 THEN EC.filterID -- Filter
				END
			FROM ASRSysExprComponents EC
			WHERE EC.exprID = @piExpressionID
				AND ((EC.type = 1 AND EC.fieldSelectionFilter > 0)
					OR (EC.type = 2)
					OR (EC.type = 3)
					OR (EC.type = 10))
		
			OPEN curSubExpressions
			FETCH NEXT FROM curSubExpressions INTO @iType, @iID
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iType = 2
				BEGIN
					-- Get the columns used in as follows:
					-- 1) Function component sub-expressions
					DECLARE curFunctionSubExpressions CURSOR LOCAL FAST_FORWARD FOR 
					SELECT
						E.exprID
					FROM ASRSysExpressions E
					WHERE E.parentComponentID = @iID
		
					OPEN curFunctionSubExpressions
					FETCH NEXT FROM curFunctionSubExpressions INTO @iExprID
					WHILE (@@fetch_status = 0)
					BEGIN
						EXEC spASRColumnsUsedInExpression @iExprID, @curSubColumns OUTPUT
		
						FETCH NEXT FROM @curSubColumns INTO @iColumnID
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #curColumns (columnID) VALUES (@iColumnID)
									
							FETCH NEXT FROM @curSubColumns INTO @iColumnID
						END
						CLOSE @curSubColumns
						DEALLOCATE @curSubColumns
		
						FETCH NEXT FROM curFunctionSubExpressions INTO @iExprID
					END
					CLOSE curFunctionSubExpressions
					DEALLOCATE curFunctionSubExpressions
				END
				ELSE
				BEGIN
					-- Get the columns used in as follows:
					-- 1) Field component filters
					-- 2) Calculation components
					-- 3) Filter components
					EXEC spASRColumnsUsedInExpression @iID, @curSubColumns OUTPUT
		
					FETCH NEXT FROM @curSubColumns INTO @iColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #curColumns (columnID) VALUES (@iColumnID)
								
						FETCH NEXT FROM @curSubColumns INTO @iColumnID
					END
					CLOSE @curSubColumns
					DEALLOCATE @curSubColumns
				END
		
				FETCH NEXT FROM curSubExpressions INTO @iType, @iID
			END
			CLOSE curSubExpressions
			DEALLOCATE curSubExpressions
		
			/* Return the cursor of columns. */
			SET @pcurColumns = CURSOR FORWARD_ONLY STATIC FOR
				SELECT columnID 
				FROM #curColumns
			OPEN @pcurColumns
		
			DROP TABLE #curColumns
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowColumnsUsed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowColumnsUsed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowColumnsUsed]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowColumnsUsed]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRWorkflowColumnsUsed
		(
			@piWorkflowID		integer,
			@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
			@pfDeleteTrigger	bit,		-- 1 when the deleted record is the trigger record
			@curColumnsUsed		cursor varying output		
		)
		AS
		BEGIN
			DECLARE
				@iBaseTableID		integer,
				@sIdentifier		varchar(8000),
				@iElementType		integer,
				@iEmailType			integer,
				@iEmailColumnID		integer,
				@iEmailExprID		integer,
				@iExprColumnID		integer,
				@curExprColumns		cursor
		
			CREATE TABLE #columnsUsed (columnID integer)
		
			-- Get the basic info of the given Workflow/element
			IF @pfDeleteTrigger = 1
			BEGIN
				-- Get the table of the deleted record.
				SELECT @sIdentifier = '''',
					@iBaseTableID = isnull(WF.baseTable, 0)
				FROM ASRSysWorkflows WF 
				WHERE WF.ID = @piWorkflowID
			END
			ELSE
			BEGIN
				-- Get the table of the deleted record
				-- and the identifier of the StoredData element.
				SELECT @sIdentifier = isnull(WFE.identifier, ''''),
					@iBaseTableID = isnull(WFE.dataTableID, 0)
				FROM ASRSysWorkflowElements WFE
				WHERE WFE.ID = @piElementID
					AND WFE.type = 5 -- StoredData
					AND WFE.dataAction = 2 -- Delete
			END
		
			----------------------------------------------------------------------------
			-- Determine which fields from the Deleted record are used in Email elements
			-- 1) Email items
			----------------------------------------------------------------------------
			INSERT INTO #columnsUsed
			SELECT WEI.dbColumnID
			FROM ASRSysWorkflowElementItems WEI
			INNER JOIN ASRSysWorkflowElements WE ON WEI.elementID = WE.ID
			INNER JOIN ASRSysColumns Cols ON WEI.dbColumnID = Cols.columnID
			WHERE WE.workflowID = @piWorkflowID
				AND WE.type = 3 -- email
				AND WEI.itemType = 1 -- DBValue	
				AND Cols.tableID = @iBaseTableID
				AND (((@pfDeleteTrigger = 1) AND (WEI.dbRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (WEI.dbRecord = 1) -- Identified
						AND (WEI.recSelWebFormIdentifier = @sIdentifier)))
		
			----------------------------------------------------------------------------
			-- Determine which fields from the Deleted record are used in WebForm elements
			-- 1) WebForm DBValues
			----------------------------------------------------------------------------
			INSERT INTO #columnsUsed
			SELECT WEI.dbColumnID
			FROM ASRSysWorkflowElementItems WEI
			INNER JOIN ASRSysWorkflowElements WE ON WEI.elementID = WE.ID
			INNER JOIN ASRSysColumns Cols ON WEI.dbColumnID = Cols.columnID
			WHERE WE.workflowID = @piWorkflowID
				AND WE.type = 2 -- WebForm
				AND WEI.itemType = 1 -- DBValue	
				AND Cols.tableID = @iBaseTableID
				AND (((@pfDeleteTrigger = 1) AND (WEI.dbRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (WEI.dbRecord = 1) -- Identified
						AND (WEI.WFFormIdentifier = @sIdentifier)))
		
			----------------------------------------------------------------------------
			-- Determine which fields from the Deleted record are used in StoredData elements
			-- 1) StoredData DBValues
			----------------------------------------------------------------------------
			INSERT INTO #columnsUsed
			SELECT WEC.dbColumnID
			FROM ASRSysWorkflowElementColumns WEC
			INNER JOIN ASRSysWorkflowElements WE ON WEC.elementID = WE.ID
			INNER JOIN ASRSysColumns Cols ON WEC.dbColumnID = Cols.columnID
			WHERE WE.workflowID = @piWorkflowID
				AND WE.type = 5 -- StoredData
				AND WEC.valueType = 2 -- DBValue	
				AND Cols.tableID = @iBaseTableID
				AND (((@pfDeleteTrigger = 1) AND (WEC.dbRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (WEC.dbRecord = 1) -- Identified
						AND (WEC.WFFormIdentifier = @sIdentifier)))
		
			----------------------------------------------------------------------------
			-- Determine which fields from the'


	SET @sSPCode_1 = ' Deleted record are used in Expressions
			----------------------------------------------------------------------------
			INSERT INTO #columnsUsed
			SELECT EC.fieldColumnID
			FROM ASRSysExprComponents EC
			INNER JOIN ASRSysExpressions EXPRS ON EC.exprID = EXPRS.exprID
			INNER JOIN ASRSysColumns Cols ON EC.fieldColumnID = Cols.columnID
			WHERE EXPRS.utilityID = @piWorkflowID
				AND EC.type = 12 -- WFField	
				AND Cols.tableID = @iBaseTableID
				AND (((@pfDeleteTrigger = 1) AND (EC.workflowRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (EC.workflowRecord = 1) -- Identified
						AND (EC.workflowElement = @sIdentifier)))

			----------------------------------------------------------------------------
			-- Return a recordset of the columns in the deleted record''s table that are used
			-- elsewhere in the Workflow
			----------------------------------------------------------------------------
			SET @curColumnsUsed = CURSOR FORWARD_ONLY STATIC FOR
			SELECT DISTINCT columnID
			FROM #columnsUsed 
		
			OPEN @curColumnsUsed
		
			DROP TABLE #columnsUsed 
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRWorkflowEmailsUsed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowEmailsUsed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowEmailsUsed]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowEmailsUsed]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRWorkflowEmailsUsed
		(
			@piWorkflowID		integer,
			@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
			@pfDeleteTrigger	bit,		-- 1 when the deleted record is the trigger record
			@curEmailsUsed		cursor varying output		
		)
		AS
		BEGIN
			DECLARE
				@iBaseTableID		integer,
				@sIdentifier		varchar(8000)

			CREATE TABLE #emailsUsed 
				(emailID integer,
				type integer,
				colExprID integer)
		
			-- Get the basic info of the given Workflow/element
			IF @pfDeleteTrigger = 1
			BEGIN
				-- Get the table of the deleted record.
				SELECT @sIdentifier = '''',
					@iBaseTableID = isnull(WF.baseTable, 0)
				FROM ASRSysWorkflows WF 
				WHERE WF.ID = @piWorkflowID
			END
			ELSE
			BEGIN
				-- Get the table of the deleted record
				-- and the identifier of the StoredData element.
				SELECT @sIdentifier = isnull(WFE.identifier, ''''),
					@iBaseTableID = isnull(WFE.dataTableID, 0)
				FROM ASRSysWorkflowElements WFE
				WHERE WFE.ID = @piElementID
					AND WFE.type = 5 -- StoredData
					AND WFE.dataAction = 2 -- Delete
			END
		
			----------------------------------------------------------------------------
			-- Determine which fields from the Deleted record are used in Email elements
			-- 1) Email address
			----------------------------------------------------------------------------
			INSERT INTO #emailsUsed
			SELECT WFE.emailID,
				EA.type,
				CASE
					WHEN EA.type = 1 THEN EA.columnID -- Column
					ELSE EA.exprID -- Calculated
				END
			FROM ASRSysWorkflowElements WFE
			INNER JOIN ASRSysEmailAddress EA ON WFE.emailID = EA.emailID
			WHERE WFE.workflowID = @piWorkflowID
				AND WFE.type = 3 -- Email
				AND EA.tableID = @iBaseTableID
				AND ((EA.type = 1) OR (EA.type = 2))
				AND (((@pfDeleteTrigger = 1) AND (WFE.emailRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (WFE.emailRecord = 1) -- Identified
						AND (WFE.recSelWebFormIdentifier = @sIdentifier)))
		
			----------------------------------------------------------------------------
			-- Return a cursor of the columns in the deleted record''s table that are used
			-- elsewhere in the Workflow
			----------------------------------------------------------------------------
			SET @curEmailsUsed = CURSOR FORWARD_ONLY STATIC FOR
			SELECT DISTINCT emailID,
				type,
				colExprID
			FROM #emailsUsed 

			OPEN @curEmailsUsed
		
			DROP TABLE #emailsUsed 
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetParentDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetParentDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetParentDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetParentDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetParentDetails
		(
			@piBaseTableID		integer,
			@piBaseRecordID		integer,
			@piParent1TableID	integer	OUTPUT,
			@piParent1RecordID	integer	OUTPUT,
			@piParent2TableID	integer	OUTPUT,
			@piParent2RecordID	integer	OUTPUT
		)
		AS
		BEGIN
			-- Return the parent table IDs and related record IDs for the given table and record.
			-- Return 0 if no table/record exists.
			DECLARE
				@sSQL		nvarchar(4000),
				@sParam		nvarchar(4000),
				@sTableName	nvarchar(4000)
		
			SET @piParent1TableID = 0
			SET @piParent1RecordID = 0
			SET @piParent2TableID = 0
			SET @piParent2RecordID = 0
		
			SELECT @sTableName = tableName
			FROM ASRSysTables 
			WHERE tableID = @piBaseTableID
		
			SELECT TOP 1 @piParent1TableID = isnull(parentID, 0)
			FROM ASRSysRelations 
			WHERE childID = @piBaseTableID
			ORDER BY parentID ASC
		
			SELECT TOP 1 @piParent2TableID = isnull(parentID, 0)
			FROM ASRSysRelations 
			WHERE childID = @piBaseTableID
				AND parentID <> @piParent1TableID
			ORDER BY parentID ASC
		
			IF (LEN(@sTableName) > 0) AND (@piBaseRecordID > 0)
			BEGIN
				IF (@piParent1TableID > 0)
				BEGIN
					SET @sSQL = ''SELECT @piParent1RecordID = isnull(ID_'' + convert(nvarchar(4000), @piParent1TableID) + '',0)''
						+ '' FROM '' + @sTableName
						+ '' WHERE ID = '' + convert(varchar(8000), @piBaseRecordID)
					SET @sParam = N''@piParent1RecordID integer OUTPUT''
					EXEC sp_executesql @sSQL, @sParam, @piParent1RecordID OUTPUT
				END
		
				IF @piParent2TableID > 0 
				BEGIN
					SET @sSQL = ''SELECT @piParent2RecordID = isnull(ID_'' + convert(nvarchar(4000), @piParent2TableID) + '',0)''
						+ '' FROM '' + @sTableName
						+ '' WHERE ID = '' + convert(varchar(8000), @piBaseRecordID)
					SET @sParam = N''@piParent2RecordID integer OUTPUT''
					EXEC sp_executesql @sSQL, @sParam, @piParent2RecordID OUTPUT
				END		
			END
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowAscendantRecordID
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowAscendantRecordID]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowAscendantRecordID]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowAscendantRecordID]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRWorkflowAscendantRecordID
		(
			@piBaseTableID	integer,
			@piBaseRecordID	integer,
			@piParent1TableID	integer,
			@piParent1RecordID	integer,
			@piParent2TableID	integer,
			@piParent2RecordID	integer,
			@piRequiredTableID	integer,
			@piRequiredRecordID	integer	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iParentTableID		integer,
				@iParentRecordID	integer
		
			SET @piRequiredRecordID = 0
			SET @piParent1TableID = isnull(@piParent1TableID, 0)
			SET @piParent1RecordID = isnull(@piParent1RecordID, 0)
			SET @piParent2TableID = isnull(@piParent2TableID, 0)
			SET @piParent2RecordID = isnull(@piParent2RecordID, 0)
		
			IF @piBaseTableID = @piRequiredTableID
			BEGIN
				SET @piRequiredRecordID = @piBaseRecordID
				RETURN
			END
		
			-- The base table is not the same as the required table.
			-- Check ascendant tables.
			DECLARE ascendantsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysRelations.parentID
			FROM ASRSysRelations
			WHERE ASRSysRelations.childID = @piBaseTableID
		
			OPEN ascendantsCursor
			FETCH NEXT FROM ascendantsCursor INTO @iParentTableID
			WHILE (@@fetch_status = 0) AND (@piRequiredRecordID = 0)
			BEGIN
				-- Get the related record in the parent table (if one exists)
				IF EXISTS 
					(SELECT * 
					FROM dbo.sysobjects 
					WHERE id = object_id(N''[dbo].[spASRSysWorkflowParentRecord]'') AND OBJECTPROPERTY(id, N''IsProcedure'') = 1)
				BEGIN
					EXEC spASRSysWorkflowParentRecord
						@piBaseTableID,
						@piBaseRecordID,
						@iParentTableID,
						@iParentRecordID OUTPUT
				END
				ELSE
				BEGIN
					SET @iParentRecordID = 0
				END
		
				IF @iParentRecordID > 0 
				BEGIN
					EXEC spASRWorkflowAscendantRecordID
						@iParentTableID,
						@iParentRecordID,
						0,					
						0,					
						0,					
						0,					
						@piRequiredTableID,
						@piRequiredRecordID OUTPUT
				END
			
				FETCH NEXT FROM ascendantsCursor INTO @iParentTableID
			END
			CLOSE ascendantsCursor
			DEALLOCATE ascendantsCursor
			
			IF (@piRequiredRecordID = 0) 
				AND (@piParent1TableID > 0)
				AND (@piParent1RecordID > 0)
			BEGIN
				EXEC spASRWorkflowAscendantRecordID
					@piParent1TableID,
					@piParent1RecordID,
					0,					
					0,					
					0,					
					0,					
					@piRequiredTableID,
					@piRequiredRecordID OUTPUT
			END

			IF (@piRequiredRecordID = 0) 
				AND (@piParent2TableID > 0)
				AND (@piParent2RecordID > 0)
			BEGIN
				EXEC spASRWorkflowAscendantRecordID
					@piParent2TableID,
					@piParent2RecordID,
					0,					
					0,					
					0,					
					0,					
					@piRequiredTableID,
					@piRequiredRecordID OUTPUT
			END
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowValidTableRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowValidTableRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowValidTableRecord]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowValidTableRecord]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRWorkflowValidTableRecord
			@piTableID	integer,
			@piRecordID	integer,
			@pfValid	bit			OUTPUT
		AS
		BEGIN
			DECLARE
				@sSQL	nvarchar(4000),
				@sParam	nvarchar(500)
				
				SET @pfValid = 0
		
				IF EXISTS (SELECT *
					FROM dbo.sysobjects
					WHERE id = object_id(N''[dbo].[udf_ASRWorkflowValidTableRecord]'')
						AND OBJECTPROPERTY(id, N''IsScalarFunction'') = 1)
				BEGIN
					SET @sSQL = ''SET @pfValid = [dbo].[udf_ASRWorkflowValidTableRecord]('' 
						+ convert(nvarchar(4000), @piTableID) 
						+ '', '' 
						+ convert(nvarchar(4000), @piRecordID)
						+ '')''
					SET @sParam = N''@pfValid bit OUTPUT''
					EXEC sp_executesql @sSQL, @sParam, @pfValid OUTPUT
				END	
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetStoredDataActionDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetStoredDataActionDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetStoredDataActionDetails
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psSQL				varchar(8000)	OUTPUT, 
				@piDataTableID		integer			OUTPUT,
				@psTableName		varchar(8000)	OUTPUT,
				@piDataAction		integer			OUTPUT, 
				@piRecordID			integer			OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iPersonnelTableID			integer,
					@iInitiatorID				integer,
					@iDataRecord				integer,
					@sIDColumnName				varchar(8000),
					@iColumnID					integer, 
					@sColumnName				varchar(8000), 
					@iColumnDataType			integer, 
					@sColumnList				varchar(8000),
					@sValueList					varchar(8000),
					@sValue						varchar(8000),
					@sRecSelWebFormIdentifier	varchar(8000),
					@sRecSelIdentifier			varchar(8000),
					@iTempTableID				integer,
					@iSecondaryDataRecord		integer,
					@sSecondaryRecSelWebFormIdentifier	varchar(8000),
					@sSecondaryRecSelIdentifier	varchar(8000),
					@sSecondaryIDColumnName		varchar(8000),
					@iSecondaryRecordID			integer,
					@iElementType				integer,
					@iWorkflowID				integer,
					@iID int,
					@sWFFormIdentifier varchar(8000),
					@sWFValueIdentifier varchar(8000),
					@iDBColumnID int,
					@iDBRecord int,
					@sSQL nvarchar(4000),
					@sParam nvarchar(4000),
					@sDBColumnName nvarchar(4000),
					@sDBTableName nvarchar(4000),
					@iRecordID int,
					@sDBValue varchar(8000),
					@iDataType int, 
					@iValueType int, 
					@iSDColumnID int,
					@fValidRecordID	bit,
					@iBaseTableID	integer,
					@iBaseRecordID	integer,
					@iRequiredTableID	integer,
					@iRequiredRecordID	integer,
					@iDataRecordTableID	integer,
					@iSecondaryDataRecordTableID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iInitParent1TableID	int,
					@iInitParent1RecordID	int,
					@iInitParent2TableID	int,
					@iInitParent2RecordID	int,
					@curColumns		cursor,
					@curEmails		cursor,
					@iEmailID		int,
					@iType			int,
					@fDeletedValue	bit,
					@iTempElementID	integer,
					@iCount			integer,
					@iTriggerTableID	int
						
				SET @psSQL = ''''
				SET @piRecordID = 0
			
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SELECT @piDataAction = dataAction,
					@piDataTableID = dataTableID,
					@iDataRecord = dataRecord,
					@sRecSelWebFormIdentifier = recSelWebFormIdentifier,
					@sRecSelIdentifier = recSelIdentifier,
					@iSecondaryDataRecord = secondaryDataRecord,
					@sSecondaryRecSelWebFormIdentifier = secondaryRecSelWebFormIdentifier,
					@sSecondaryRecSelIdentifier = secondaryRecSelIdentifier,
					@iDataRecordTableID = dataRecordTable,
					@iSecondaryDataRecordTableID = secondaryDataRecordTable,
					@iWorkflowID = workflowID,
					@iTriggerTableID = ASRSysWorkflows.baseTable
				FROM ASRSysWorkflowElements
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID
				WHERE ASRSysWorkflowElements.ID = @piElementID
			
				SELECT @psTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @piDataTableID
			
				IF @iDataRecord = 0 -- 0 = Initiator''s record
				BEGIN
					EXEC spASRWorkflowAscendantRecordID
						@iPersonnelTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2T'


	SET @sSPCode_1 = 'ableID,
						@iInitParent2RecordID,
						@iDataRecordTableID,
						@piRecordID	OUTPUT
			
					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
					END
				END

				IF @iDataRecord = 4 -- 4 = Triggered record
				BEGIN
					EXEC spASRWorkflowAscendantRecordID
						@iTriggerTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2TableID,
						@iInitParent2RecordID,
						@iDataRecordTableID,
						@piRecordID	OUTPUT
			
					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
					END
				END

				IF @iDataRecord = 1 -- 1 = Identified record
				BEGIN
					SELECT @iElementType = ASRSysWorkflowElements.type
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
					
					IF @iElementType = 2
					BEGIN
						 -- WebForm
						SELECT @piRecordID = 
							CASE
								WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
								ELSE 0
							END,
							@iTempTableID = EI.tableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sRecSelIdentifier
							AND Es.identifier = @sRecSelWebFormIdentifier
							AND Es.workflowID = @iWorkflowID
							AND IV.elementID = Es.ID
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @piRecordID = 
							CASE
								WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
								ELSE 0
							END,
							@iTempTableID = Es.dataTableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND IV.identifier = Es.identifier
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sRecSelWebFormIdentifier
						WHERE IV.instanceID = @piInstanceID
					END
				
					SET @iBaseTableID = @iTempTableID
					SET @iBaseRecordID = @piRecordID
					EXEC spASRWorkflowAscendantRecordID
						@iBaseTableID,
						@iBaseRecordID,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						@iDataRecordTableID,
						@piRecordID	OUTPUT

					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
					END
				END
			
				SET @fValidRecordID = 1
				IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
				BEGIN
					EXEC spASRWorkflowValidTableRecord
						@iDataRecordTableID,
						@piRecordID,
						@fValidRecordID	OUTPUT

					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Stored Data primary record has been deleted or not selected.''
						SET @psSQL = ''''
						RETURN
					END
				END

				IF @piDataAction = 0 -- Insert
				BEGIN
					IF @iSecondaryDataRecord ='


	SET @sSPCode_2 = ' 0 -- 0 = Initiator''s record
					BEGIN
						EXEC spASRWorkflowAscendantRecordID
							@iPersonnelTableID,
							@iInitiatorID,
							@iInitParent1TableID,
							@iInitParent1RecordID,
							@iInitParent2TableID,
							@iInitParent2RecordID,
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID	OUTPUT
				
						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END
					END
					
					IF @iSecondaryDataRecord = 4 -- 4 = Triggered record
					BEGIN
						EXEC spASRWorkflowAscendantRecordID
							@iTriggerTableID,
							@iInitiatorID,
							@iInitParent1TableID,
							@iInitParent1RecordID,
							@iInitParent2TableID,
							@iInitParent2RecordID,
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID	OUTPUT
				
						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END
					END

					IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
					BEGIN
						SELECT @iElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)))
				
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @iSecondaryRecordID = 
								CASE
									WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
									ELSE 0
								END,
								@iTempTableID = EI.tableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sSecondaryRecSelIdentifier
								AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @iSecondaryRecordID = 
								CASE
									WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
									ELSE 0
								END,
								@iTempTableID = Es.dataTableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
						
						SET @iBaseTableID = @iTempTableID
						SET @iBaseRecordID = @iSecondaryRecordID
						EXEC spASRWorkflowAscendantRecordID
							@iBaseTableID,
							@iBaseRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID	OUTPUT

						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END
					END

					SET @fValidRecordID = 1
					IF (@iSecondaryDataRe'


	SET @sSPCode_3 = 'cord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
					BEGIN
						EXEC spASRWorkflowValidTableRecord
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Stored Data secondary record has been deleted or not selected.''

							SET @psSQL = ''''
							RETURN
						END
					END
				END

				IF @piDataAction = 0 OR @piDataAction = 1
				BEGIN
					/* INSERT or UPDATE. */
					SET @sColumnList = ''''
					SET @sValueList = ''''

					CREATE TABLE #dbValues (ID integer, 
						wfFormIdentifier varchar(1000),
						wfValueIdentifier varchar(1000),
						dbColumnID int,
						dbRecord int,
						value varchar(8000))

					INSERT INTO #dbValues (ID, 
						wfFormIdentifier,
						wfValueIdentifier,
						dbColumnID,
						dbRecord,
						value) 
					SELECT EC.ID,
						EC.wfformidentifier,
						EC.wfvalueidentifier,
						EC.dbcolumnid,
						EC.dbrecord, 
						''''
					FROM ASRSysWorkflowElementColumns EC
					WHERE EC.elementID = @piElementID
						AND EC.valueType = 2
						
					DECLARE dbValuesCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ID,
						wfFormIdentifier,
						wfValueIdentifier,
						dbColumnID,
						dbRecord
					FROM #dbValues
					OPEN dbValuesCursor
					FETCH NEXT FROM dbValuesCursor INTO @iID,
						@sWFFormIdentifier,
						@sWFValueIdentifier,
						@iDBColumnID,
						@iDBRecord
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @fDeletedValue = 0

						SELECT @sDBTableName = tbl.tableName,
							@iRequiredTableID = tbl.tableID, 
							@sDBColumnName = col.columnName,
							@iDataType = col.dataType
						FROM ASRSysColumns col
						INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
						WHERE col.columnID = @iDBColumnID

						SET @sSQL = ''SELECT @sDBValue = ''
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
								ELSE ''convert(varchar(8000),''
							END
							+ @sDBTableName + ''.'' + @sDBColumnName
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN '', 101)''
								ELSE '')''
							END
							+ '' FROM '' + @sDBTableName 
							+ '' WHERE '' + @sDBTableName + ''.ID = ''

						SET @iRecordID = 0

						IF @iDBRecord = 0
						BEGIN
							-- Initiator''s record
							SET @iRecordID = @iInitiatorID
							SET @iParent1TableID = @iInitParent1TableID
							SET @iParent1RecordID = @iInitParent1RecordID
							SET @iParent2TableID = @iInitParent2TableID
							SET @iParent2RecordID = @iInitParent2RecordID

							SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
							FROM ASRSysModuleSetup
							WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_TablePersonnel''
						END			

						IF @iDBRecord = 4
						BEGIN
							-- Trigger record
							SET @iRecordID = @iInitiatorID
							SET @iParent1TableID = @iInitParent1TableID
							SET @iParent1RecordID = @iInitParent1RecordID
							SET @iParent2TableID = @iInitParent2TableID
							SET @iParent2RecordID = @iInitParent2RecordID

							SELECT @iBaseTableID = isnull(WF.baseTable, 0)
							FROM ASRSysWorkflows WF
							INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
								AND WFI.ID = @piInstanceID
						END
						
						IF @iDBRecord = 1
						BEGIN
							-- Identified record
							SELECT @iElementType = ASRSysWorkflowElements.type, 
								@iTempElementID = ASRSysWorkflowElements.ID
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
								AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdent'


	SET @sSPCode_4 = 'ifier)))

							IF @iElementType = 2
							BEGIN
								 -- WebForm
								SELECT @iRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@iBaseTableID = EI.tableID,
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
								INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
								WHERE IV.instanceID = @piInstanceID
									AND IV.identifier = @sWFValueIdentifier
									AND Es.identifier = @sWFFormIdentifier
									AND Es.workflowID = @iWorkflowID
									AND IV.elementID = Es.ID
							END
							ELSE
							BEGIN
								-- StoredData
								SELECT @iRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@iBaseTableID = isnull(Es.dataTableID, 0),
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
									AND IV.identifier = Es.identifier
									AND Es.workflowID = @iWorkflowID
									AND Es.identifier = @sWFFormIdentifier
								WHERE IV.instanceID = @piInstanceID
							END
						END

						SET @iBaseRecordID = @iRecordID

						SET @fValidRecordID = 1
						
						IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
						BEGIN
							SET @fValidRecordID = 0

							EXEC spASRWorkflowAscendantRecordID
								@iBaseTableID,
								@iBaseRecordID,
								@iParent1TableID,
								@iParent1RecordID,
								@iParent2TableID,
								@iParent2RecordID,
								@iRequiredTableID,
								@iRequiredRecordID	OUTPUT

							SET @iRecordID = @iRequiredRecordID

							IF @iRecordID > 0 
							BEGIN
								EXEC spASRWorkflowValidTableRecord
									@iRequiredTableID,
									@iRecordID,
									@fValidRecordID	OUTPUT
							END

							IF @fValidRecordID = 0
							BEGIN
								IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
								BEGIN
									SELECT @iCount = COUNT(*)
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.columnID = @iDBColumnID

									IF @iCount = 1
									BEGIN
										SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
										FROM ASRSysWorkflowQueueColumns QC
										INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
										WHERE WFQ.instanceID = @piInstanceID
											AND QC.columnID = @iDBColumnID

										SET @fValidRecordID = 1
										SET @fDeletedValue = 1
									END
								END
								ELSE
								BEGIN
									IF @iDBRecord = 1
									BEGIN
										SELECT @iCount = COUNT(*)
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.columnID = @iDBColumnID
											AND IV.elementID = @iTempElementID

										IF @iCount = 1
										BEGIN
											SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '''')))
											FROM ASRSysWorkflowInstanceValues IV
											WHERE IV.instanceID = @piInstanceID
												AND IV.columnID = @iDBColumnID
												AND IV.elementID = @iTempElementID

											SET @fValidRecordID = 1
											SET @fDeletedValue = 1
										END
									END
								END
							END

							IF @fValidRecordID = 0
							BEGIN
'


	SET @sSPCode_5 = '								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Stored Data column database value record has been deleted or not selected.''

								SET @psSQL = ''''
								RETURN
							END
						END

						IF @fDeletedValue = 0
						BEGIN
							SET @sSQL = @sSQL + convert(nvarchar(4000), @iRecordID)
							SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT
						END

						UPDATE #dbValues
						SET value = @sDBValue
						WHERE ID = @iID
						
						FETCH NEXT FROM dbValuesCursor INTO @iID,
							@sWFFormIdentifier,
							@sWFValueIdentifier,
							@iDBColumnID,
							@iDBRecord
					END
					CLOSE dbValuesCursor
					DEALLOCATE dbValuesCursor
			
					DECLARE columnCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT EC.columnID,
						SC.columnName,
						SC.dataType,
						CASE
							WHEN EC.valueType = 0 THEN  -- Fixed Value
								CASE
									WHEN SC.dataType = -7 THEN
										CASE 
											WHEN UPPER(EC.value) = ''TRUE'' THEN ''1''
											ELSE ''0''
										END
									ELSE EC.value
								END
							WHEN EC.valueType = 1 THEN -- Workflow Value
								(SELECT IV.value
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
								INNER JOIN ASRSysWorkflowElements WE2 ON WE.workflowID = WE2.workflowID
								WHERE WE.identifier = EC.WFFormIdentifier
									AND WE2.id = @piElementID
									AND IV.instanceID = @piInstanceID
									AND IV.identifier = EC.WFValueIdentifier)
							ELSE '''' -- Database Value. Handle below to avoid collation conflict.
							END AS [value], 
							EC.valueType, 
							EC.ID
					FROM ASRSysWorkflowElementColumns EC
					INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
					WHERE EC.elementID = @piElementID
			
					OPEN columnCursor
					FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
						BEGIN
							SELECT @sValue = dbV.value
							FROM #dbValues dbV
							WHERE dbV.ID = @iSDColumnID
						END

						IF @piDataAction = 0 
						BEGIN
							/* INSERT. */
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sColumnName
			
							SET @sValueList = @sValueList
								+ CASE
									WHEN LEN(@sValueList) > 0 THEN '',''
									ELSE ''''
								END
								+ CASE
									WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
									WHEN @iColumnDataType = 11 THEN
										CASE 
											WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
											ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
										END
									ELSE isnull(@sValue, 0) -- integer, logic, numeric
								END
						END
						ELSE
						BEGIN
							/* UPDATE. */
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sColumnName
								+ '' = ''
								+ CASE
									WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
									WHEN @iColumnDataType = 11 THEN
										CASE 
											WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
											ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
										END
									ELSE isnull(@sValu'


	SET @sSPCode_6 = 'e, 0) -- integer, logic, numeric
								END
						END

						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
							AND elementID = @piElementID
							AND columnID = @iDBColumnID

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', @iColumnID, @sValue, 0)
			
						FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID
					END
			
					CLOSE columnCursor
					DEALLOCATE columnCursor
			
					DROP TABLE #dbValues
			
					IF @piDataAction = 0 

					BEGIN
						/* INSERT. */
						IF @iDataRecord <> 3 -- 3 = Unidentified record
						BEGIN
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sIDColumnName
				
							SET @sValueList = @sValueList
								+ CASE
									WHEN LEN(@sValueList) > 0 THEN '',''
									ELSE ''''
								END
								+ convert(varchar(8000), @piRecordID)

							IF @piDataAction = 0 -- Insert
								AND (@iSecondaryDataRecord = 0 -- 0 = Initiator''s record
									OR @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
									OR @iSecondaryDataRecord = 4) -- 4 = Triggered record
							BEGIN
								SET @sColumnList = @sColumnList
									+ CASE
										WHEN LEN(@sColumnList) > 0 THEN '',''
										ELSE ''''
									END
									+ @sSecondaryIDColumnName
							
								SET @sValueList = @sValueList
									+ CASE
										WHEN LEN(@sValueList) > 0 THEN '',''
										ELSE ''''
									END
									+ convert(varchar(8000), @iSecondaryRecordID)
							END
						END
					END
			
					IF LEN(@sColumnList) > 0
					BEGIN
						IF @piDataAction = 0 
						BEGIN
							/* INSERT. */
							SET @psSQL = ''INSERT INTO '' + @psTableName
								+ '' ('' + @sColumnList + '')''
								+ '' VALUES('' + @sValueList + '')''
						END
						ELSE
						BEGIN
							/* UPDATE. */
							SET @psSQL = ''UPDATE '' + @psTableName
								+ '' SET '' + @sColumnList
								+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @piRecordID)
						END
					END
				END
			
				IF @piDataAction = 2
				BEGIN
					/* DELETE. */
					SET @psSQL = ''DELETE FROM '' + @psTableName
						+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @piRecordID)
				END	

				IF (@piDataAction = 0) -- Insert
				BEGIN
					SET @iParent1TableID = isnull(@iDataRecordTableID, 0)
					SET @iParent1RecordID = isnull(@piRecordID, 0)
					SET @iParent2TableID = isnull(@iSecondaryDataRecordTableID, 0)
					SET @iParent2RecordID = isnull(@iSecondaryRecordID, 0)
				END
				ELSE
				BEGIN	-- Update or Delete
					exec spASRGetParentDetails
						@piDataTableID,
						@piRecordID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT
				END

				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID, 
					ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
					ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID, 
					ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
					AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0

				IF (@piDataAction = 2) -- Delete
				BEGIN
					EXEC spASRWorkflowColumnsUsed @iWorkflowID, 
						@piElementID,
						0, 
						@curColumns OUTPUT

					FETCH NEXT FROM @curColumns INTO @iDBColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE inst'


	SET @sSPCode_7 = 'anceID = @piInstanceID
							AND elementID = @piElementID
							AND columnID = @iDBColumnID

						SELECT @sDBTableName = tbl.tableName,
							@iRequiredTableID = tbl.tableID, 
							@sDBColumnName = col.columnName,
							@iDataType = col.dataType
						FROM ASRSysColumns col
						INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
						WHERE col.columnID = @iDBColumnID

						SET @sSQL = ''SELECT @sDBValue = ''
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
								ELSE ''convert(varchar(8000),''
							END
							+ @sDBTableName + ''.'' + @sDBColumnName
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN '', 101)''
								ELSE '')''
							END
							+ '' FROM '' + @sDBTableName 
							+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(8000), @piRecordID)

						SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', @iDBColumnID, @sDBValue, 0)
								
						FETCH NEXT FROM @curColumns INTO @iDBColumnID
					END
					CLOSE @curColumns
					DEALLOCATE @curColumns

					EXEC spASRWorkflowEmailsUsed @iWorkflowID, 
						@piElementID,
						0, 
						@curEmails OUTPUT

					FETCH NEXT FROM @curEmails INTO @iEmailID, @iType, @iDBColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
							AND elementID = @piElementID
							AND emailID = @iEmailID

						IF @iType = 1 -- Column
						BEGIN
							SELECT @sDBTableName = tbl.tableName,
								@iRequiredTableID = tbl.tableID, 
								@sDBColumnName = col.columnName,
								@iDataType = col.dataType
							FROM ASRSysColumns col
							INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
							WHERE col.columnID = @iDBColumnID

							SET @sSQL = ''SELECT @sDBValue = ''
								+ CASE
									WHEN @iDataType = 12 THEN ''''
									WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
									ELSE ''convert(varchar(8000),''
								END
								+ @sDBTableName + ''.'' + @sDBColumnName
								+ CASE
									WHEN @iDataType = 12 THEN ''''
									WHEN @iDataType = 11 THEN '', 101)''
									ELSE '')''
								END
								+ '' FROM '' + @sDBTableName 
								+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(8000), @piRecordID)

							SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT
						END
						ELSE
						BEGIN
							EXEC spASRSysEmailAddr
								@sDBValue OUTPUT,
								@iEmailID,
								@piRecordID
						END

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID)
								
						FETCH NEXT FROM @curEmails INTO @iEmailID, @iType, @iDBColumnID
					END
					CLOSE @curEmails
					DEALLOCATE @curEmails
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5
		+ @sSPCode_6
		+ @sSPCode_7)

	----------------------------------------------------------------------
	-- spASRActionOrsAndGetSucceedingWorkflowElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionOrsAndGetSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionOrsAndGetSucceedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRActionOrsAndGetSucceedingWorkflowElements]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRActionOrsAndGetSucceedingWorkflowElements
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@piValue		integer,
			@succeedingElements	cursor varying output
		)
		AS
		BEGIN
			-- Action any Or elements and return the IDs of the workflow elements that 
			-- succeed the given element.
			-- This ignores connection elements.
			-- NB. This does work for elements with multiple outbound flows. 
			DECLARE
				@iConnectorPairID	integer,
				@iElementID	integer,
				@superCursor		cursor,
				@iTemp		integer,
				@sForms		varchar(8000)
					
			CREATE TABLE #succeedingElements (elementID integer)
		
			/* Get the non-connector and non-or elements. */
			INSERT INTO #succeedingElements
			SELECT L.endElementID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND ((L.startOutboundFlowCode = @piValue) OR 
					(@piValue = 0 and L.startOutboundFlowCode = -1))
				AND E.type <> 8 -- 8 = Connector 1
				AND E.type <> 7 -- 7 = Or
		
			-- Action any succeeding Or elements
			DECLARE orCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT L.endElementID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND ((L.startOutboundFlowCode = @piValue) OR 
					(@piValue = 0 and L.startOutboundFlowCode = -1))
				AND E.type = 7 -- 7 = Or
			OPEN orCursor
			FETCH NEXT FROM orCursor INTO @iElementID
			WHILE (@@fetch_status = 0)
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
		
				EXEC spASRCancelPendingPrecedingWorkflowElements @piInstanceID, @iElementID, @iElementID
				EXEC spASRSubmitWorkflowStep @piInstanceID, @iElementID, '''', '''', @sForms OUTPUT
		
				EXEC spASRActionOrsAndGetSucceedingWorkflowElements 
					@piInstanceID, 
					@iElementID, 
					0, 
					@superCursor OUTPUT	
		
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
					
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
		
				FETCH NEXT FROM orCursor INTO @iElementID
			END
			CLOSE orCursor
			DEALLOCATE orCursor
				
			DECLARE succeedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.connectionPairID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND ((L.startOutboundFlowCode = @piValue) OR 
					(@piValue = 0 and L.startOutboundFlowCode = -1))
				AND E.type = 8 -- 8 = Connector 1
				
			OPEN succeedingConnectorsCursor
			FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC spASRActionOrsAndGetSucceedingWorkflowElements @piInstanceID, @iConnectorPairID, 0, @superCursor OUTPUT	
		
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
							
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
				
				FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
			END
			CLOSE succeedingConnectorsCursor
			DEALLOCATE succeedingConnectorsCursor
				
			-- Return the cursor of succeeding elements. 
			SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
				SELECT elementID 
				FROM #succeedingElements
			OPEN @succeedingElements
				
			DROP TABLE #succeedingElements
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetWorkflowEmailMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowEmailMessage
					(
						@piInstanceID		integer,
						@piElementID		integer,
						@psMessage		varchar(8000)	OUTPUT, 
						@pfOK	bit	OUTPUT
					)
					AS
					BEGIN
						DECLARE 
							@iInitiatorID		integer,
							@sCaption		varchar(8000),
							@iItemType		integer,
							@iDBColumnID		integer,
							@iDBRecord		integer,
							@sWFFormIdentifier	varchar(8000),
							@sWFValueIdentifier	varchar(8000),
							@sValue		varchar(8000),
							@sTableName		sysname,
							@sColumnName		sysname,
							@iRecordID		integer,
							@sSQL			nvarchar(4000),
							@sSQLParam		nvarchar(4000),
							@iCount		integer,
							@iElementID		integer,
							@superCursor		cursor,
							@iTemp		integer,
							@hResult 		integer,
							@objectToken 		integer,
							@sQueryString		varchar(8000),
							@sURL			varchar(8000), 
							@sEmailFormat		varchar(8000),
							@iEmailFormat		integer,
							@iSourceItemType	integer,
							@dtTempDate		datetime, 
							@sParam1	varchar(8000),
							@sDBName	sysname,
							@sRecSelWebFormIdentifier	varchar(8000),
							@sRecSelIdentifier	varchar(8000),
							@iElementType		integer,
							@iWorkflowID		integer, 
							@fValidRecordID	bit,
							@iBaseTableID	integer,
							@iBaseRecordID	integer,
							@iRequiredTableID	integer,
							@iRequiredRecordID	integer,
							@iParent1TableID	int,
							@iParent1RecordID	int,
							@iParent2TableID	int,
							@iParent2RecordID	int,
							@iInitParent1TableID	int,
							@iInitParent1RecordID	int,
							@iInitParent2TableID	int,
							@iInitParent2RecordID	int,
							@fDeletedValue		bit,
							@iTempElementID		integer,
							@iColumnID			integer
									
						SET @pfOK = 1
						SET @psMessage = ''''
					
						exec spASRGetSetting 
							''email'',
							''date format'',
							''103'',
							0,
							@sEmailFormat		OUTPUT
			
						SET @iEmailFormat = convert(integer, @sEmailFormat)
						
						SELECT @sURL = parameterValue
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_URL''
		
						IF upper(right(@sURL, 5)) <> ''.ASPX''
							AND right(@sURL, 1) <> ''/''
							AND len(@sURL) > 0
						BEGIN
							SET @sURL = @sURL + ''/''
						END
			
						SELECT @sParam1 = parameterValue
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''		
							AND parameterKey = ''Param_Web1''
						
						SET @sDBName = db_name()
			
						SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
							@iWorkflowID = ASRSysWorkflowInstances.workflowID,
							@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
							@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
							@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
							@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
						FROM ASRSysWorkflowInstances
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					
						DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT EI.caption,
							EI.itemType,
							EI.dbColumnID,
							EI.dbRecord,
							EI.wfFormIdentifier,
							EI.wfValueIdentifier, 
							EI.recSelWebFormIdentifier,
							EI.recSelIdentifier
						FROM ASRSysWorkflowElementItems EI
						WHERE EI.elementID = @piElementID
						ORDER BY EI.ID
					
						OPEN itemCursor
						FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier
						WHILE (@@fetch_status = 0)
						BEGIN
							SET @sValue = ''''

							IF @iItemType = 1
							BEGIN
								SET @fDeletedValue = 0

								/* Database value. */
								SELECT @sTableName = ASRSysTables.tableName, 
									@iRequiredTableID = ASRSysTables.tableID, 
									@sColumnName = ASRSysColumns.columnNam'


	SET @sSPCode_1 = 'e, 
									@iSourceItemType = ASRSysColumns.dataType
								FROM ASRSysColumns
								INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
								WHERE ASRSysColumns.columnID = @iDBColumnID
					
								IF @iDBRecord = 0
								BEGIN
									-- Initiator''s record
									SET @iRecordID = @iInitiatorID
									SET @iParent1TableID = @iInitParent1TableID
									SET @iParent1RecordID = @iInitParent1RecordID
									SET @iParent2TableID = @iInitParent2TableID
									SET @iParent2RecordID = @iInitParent2RecordID

									SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
									FROM ASRSysModuleSetup
									WHERE moduleKey = ''MODULE_WORKFLOW''
									AND parameterKey = ''Param_TablePersonnel''
								END			

								IF @iDBRecord = 4
								BEGIN
									-- Trigger record
									SET @iRecordID = @iInitiatorID
									SET @iParent1TableID = @iInitParent1TableID
									SET @iParent1RecordID = @iInitParent1RecordID
									SET @iParent2TableID = @iInitParent2TableID
									SET @iParent2RecordID = @iInitParent2RecordID

									SELECT @iBaseTableID = isnull(WF.baseTable, 0)
									FROM ASRSysWorkflows WF
									INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
										AND WFI.ID = @piInstanceID
								END
			
								IF @iDBRecord = 1
								BEGIN
									-- Previously identified record.
									SELECT @iElementType = ASRSysWorkflowElements.type, 
										@iTempElementID = ASRSysWorkflowElements.ID
									FROM ASRSysWorkflowElements
									WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
										AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
			
									IF @iElementType = 2
									BEGIN
										 -- WebForm
										SELECT @iRecordID = 
											CASE
												WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
												ELSE 0
											END,
											@iBaseTableID = EI.tableID,
											@iParent1TableID = IV.parent1TableID,
											@iParent1RecordID = IV.parent1RecordID,
											@iParent2TableID = IV.parent2TableID,
											@iParent2RecordID = IV.parent2RecordID
										FROM ASRSysWorkflowInstanceValues IV
										INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
										INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
										WHERE IV.instanceID = @piInstanceID
											AND IV.identifier = @sRecSelIdentifier
											AND Es.identifier = @sRecSelWebFormIdentifier
											AND Es.workflowID = @iWorkflowID
											AND IV.elementID = Es.ID
									END
									ELSE
									BEGIN
										-- StoredData
										SELECT @iRecordID = 
											CASE
												WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
												ELSE 0
											END,
											@iBaseTableID = isnull(Es.dataTableID, 0),
											@iParent1TableID = IV.parent1TableID,
											@iParent1RecordID = IV.parent1RecordID,
											@iParent2TableID = IV.parent2TableID,
											@iParent2RecordID = IV.parent2RecordID
										FROM ASRSysWorkflowInstanceValues IV
										INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
											AND IV.identifier = Es.identifier
											AND Es.workflowID = @iWorkflowID
											AND Es.identifier = @sRecSelWebFormIdentifier
										WHERE IV.instanceID = @piInstanceID
									END
								END		
			
								SET @iBaseRecordID = @iRecordID

								IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
								BEGIN
									SET @fValidRecordID = 0

									EXEC spASRWorkflowAscendantRecordID
										@iBaseTableID,
										@iBaseRecordID,
										@iParent1TableID,
										@iParent1RecordID,
										@iParent2TableID,
										@iParent2RecordID,
										@iRequiredTableID,
										@iRequiredRecordI'


	SET @sSPCode_2 = 'D	OUTPUT

									SET @iRecordID = @iRequiredRecordID

									IF @iRecordID > 0 
									BEGIN
										EXEC spASRWorkflowValidTableRecord
											@iRequiredTableID,
											@iRecordID,
											@fValidRecordID	OUTPUT
									END

									IF @fValidRecordID = 0
									BEGIN
										IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
										BEGIN
											SELECT @iCount = COUNT(*)
											FROM ASRSysWorkflowQueueColumns QC
											INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
											WHERE WFQ.instanceID = @piInstanceID
												AND QC.columnID = @iDBColumnID

											IF @iCount = 1
											BEGIN
												SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
												FROM ASRSysWorkflowQueueColumns QC
												INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
												WHERE WFQ.instanceID = @piInstanceID
													AND QC.columnID = @iDBColumnID

												SET @fValidRecordID = 1
												SET @fDeletedValue = 1
											END
										END
										ELSE
										BEGIN
											IF @iDBRecord = 1
											BEGIN
												SELECT @iCount = COUNT(*)
												FROM ASRSysWorkflowInstanceValues IV
												WHERE IV.instanceID = @piInstanceID
													AND IV.columnID = @iDBColumnID
													AND IV.elementID = @iTempElementID

												IF @iCount = 1
												BEGIN
													SELECT @sValue = rtrim(ltrim(isnull(IV.value , '''')))
													FROM ASRSysWorkflowInstanceValues IV
													WHERE IV.instanceID = @piInstanceID
														AND IV.columnID = @iDBColumnID
														AND IV.elementID = @iTempElementID

													SET @fValidRecordID = 1
													SET @fDeletedValue = 1
												END
											END
										END
									END

									IF @fValidRecordID  = 0
									BEGIN
										SET @psMessage = ''''
										SET @pfOK = 0

										RETURN
									END
								END

								IF @fDeletedValue = 0
								BEGIN
									SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
										'' FROM '' + @sTableName +
										'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
									SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
									EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
								END					
								IF @sValue IS null SET @sValue = ''''
					
								/* Format dates */
								IF @iSourceItemType = 11
								BEGIN
									IF len(@sValue) = 0
									BEGIN
										SET @sValue = ''<undefined>''
									END
									ELSE
									BEGIN
										SET @dtTempDate = convert(datetime, @sValue)
										SET @sValue = convert(varchar(8000), @dtTempDate, @iEmailFormat)
									END
								END
			
								/* Format logics */
								IF @iSourceItemType = -7
								BEGIN
									IF @sValue = 0 
									BEGIN
										SET @sValue = ''False''
									END
									ELSE
									BEGIN
										SET @sValue = ''True''
									END
								END	
			
								SET @psMessage = @psMessage
									+ @sValue
							END
							
							IF @iItemType = 2
							BEGIN
								/* Label value. */
								SET @psMessage = @psMessage
									+ @sCaption
							END
					
							IF @iItemType = 4
							BEGIN
								/* Workflow value. */
								SELECT @sValue = ASRSysWorkflowInstanceValues.value, 
									@iSourceItemType = ASRSysWorkflowElementItems.itemType,
									@iColumnID = ASRSysWorkflowElementItems.lookupColumnID
								FROM ASRSysWorkflowInstanceValues
								INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
								INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID
								WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
									AND '


	SET @sSPCode_3 = 'ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
									AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
									AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
					
								IF @sValue IS null SET @sValue = ''''
			
								IF @iSourceItemType = 14 -- Lookup, need to get the column data type
								BEGIN
									SELECT @iSourceItemType = 
										CASE
											WHEN ASRSysColumns.dataType = -7 THEN 6 -- Logic
											WHEN ASRSysColumns.dataType = 2 THEN 5 -- Numeric
											WHEN ASRSysColumns.dataType = 4 THEN 5 -- Integer
											WHEN ASRSysColumns.dataType = 11 THEN 7 -- Date
											ELSE 3
										END
									FROM ASRSysColumns
									WHERE ASRSysColumns.columnID = @iColumnID
								END
										
								/* Format dates */
								IF @iSourceItemType = 7
								BEGIN
									IF len(@sValue) = 0 OR @sValue = ''null''
									BEGIN
										SET @sValue = ''<undefined>''
									END
									ELSE
									BEGIN
										SET @dtTempDate = convert(datetime, @sValue)
										SET @sValue = convert(varchar(8000), @dtTempDate, @iEmailFormat)
									END
								END
					
								/* Format logics */
								IF @iSourceItemType = 6
								BEGIN
									IF @sValue = 0 
									BEGIN
										SET @sValue = ''False''
									END
									ELSE
									BEGIN
										SET @sValue = ''True''
									END
								END			
			
								SET @psMessage = @psMessage
									+ @sValue
							END
			
							IF @iItemType = 12
							BEGIN
								/* Formatting option. */
								/* NB. The empty string that precede the char codes ARE required. */
								SET @psMessage = @psMessage +
									CASE
										WHEN @sCaption = ''L'' THEN '''' + char(13) + ''--------------------------------------------------'' + char(13)
										WHEN @sCaption = ''N'' THEN '''' + char(13)
										WHEN @sCaption = ''T'' THEN '''' + char(9)
										ELSE ''''
									END
							END
			
					
							FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier
						END
						CLOSE itemCursor
						DEALLOCATE itemCursor
					
						/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
						CREATE TABLE #succeedingElements (elementID integer)
					
						EXEC spASRActionOrsAndGetSucceedingWorkflowElements @piInstanceID, @piElementID, 0, @superCursor OUTPUT
					
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
					
						SELECT @iCount = COUNT(*)
						FROM #succeedingElements SE
						INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.id
						WHERE WE.type = 2 -- 2 = Web Form element
					
						IF @iCount > 0 
						BEGIN
							SET @psMessage = @psMessage + CHAR(13) + CHAR(13)
								+ ''Click on the following link''
								+ CASE
									WHEN @iCount = 1 THEN ''''
									ELSE ''s''
								END
								+ '' to action:''
								+ CHAR(13)
					
							DECLARE elementCursor CURSOR LOCAL FAST_FORWARD FOR 
							SELECT SE.elementID, ISNULL(WE.caption, '''')
							FROM #succeedingElements SE
							INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.ID
						
							OPEN elementCursor
							FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
							WHILE (@@fetch_status = 0)
							BEGIN
								EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
								IF @hResult <> 0
								BEGIN
									SET @sQueryString = ''''
								END
								ELSE
								BEGIN
									EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @sQueryStrin'


	SET @sSPCode_4 = 'g OUTPUT, @piInstanceID, @iElementID, @sParam1, @@servername, @sDBName
									IF @hResult <> 0 
									BEGIN
										SET @sQueryString = ''''
									END

									EXEC @hResult = sp_OADestroy @objectToken 
								END
											
								IF LEN(@sQueryString) = 0 
								BEGIN
									SET @psMessage = @psMessage + CHAR(13) +
										@sCaption + '' - Error constructing the query string. Please contact your system administrator.''
								END
								ELSE
								BEGIN
									SET @psMessage = @psMessage + CHAR(13) +
										@sCaption + '' - '' + CHAR(13) + 
										''<'' + @sURL + ''?'' + @sQueryString + ''>''
								END
								
								FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
							END
							CLOSE elementCursor
					
							DEALLOCATE elementCursor

							SET @psMessage = @psMessage + CHAR(13) + CHAR(13)
								+ ''Please make sure that the link''
								+ CASE
									WHEN @iCount = 1 THEN '' has''
									ELSE ''s have''
								END
								+ '' not been cut off by your display.'' + CHAR(13)
								+ ''If ''
								+ CASE
									WHEN @iCount = 1 THEN ''it has''
									ELSE ''they have''
								END
								+ '' been cut off you will need to copy and paste ''
								+ CASE
									WHEN @iCount = 1 THEN ''it''
									ELSE ''them''
								END
								+ '' into your browser.''
						END
					
						DROP TABLE #succeedingElements
					END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

	----------------------------------------------------------------------
	-- spASRGetWorkflowFormItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFormItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFormItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowFormItems
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psErrorMessage	varchar(8000)	OUTPUT,
				@piBackColour	integer	OUTPUT,
				@piBackImage	integer	OUTPUT,
				@piBackImageLocation	integer	OUTPUT,
				@piWidth	integer	OUTPUT,
				@piHeight	integer	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iID			integer,
					@iItemType		integer,
					@iDBColumnID		integer,
					@iDBColumnDataType	integer,
					@iDBRecord		integer,
					@sWFFormIdentifier	varchar(8000),
					@sWFValueIdentifier	varchar(8000),
					@sValue		varchar(8000),
					@sSQL			nvarchar(4000),
					@sSQLParam		nvarchar(4000),
					@sTableName		sysname,
					@sColumnName		sysname,
					@iInitiatorID		integer,
					@iRecordID		integer,
					@iStatus		integer,
					@iCount		integer,
					@iWorkflowID		integer,
					@iElementType		integer, 
					@iType integer,
					@fValidRecordID	bit,
					@iBaseTableID	integer,
					@iBaseRecordID	integer,
					@iRequiredTableID	integer,
					@iRequiredRecordID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iInitParent1TableID	int,
					@iInitParent1RecordID	int,
					@iInitParent2TableID	int,
					@iInitParent2RecordID	int,
					@fDeletedValue		bit,
					@iTempElementID		integer,
					@iColumnID	integer
						
				/* Check the given instance still exists. */
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				IF @iCount = 0
				BEGIN
					SET @psErrorMessage = ''This workflow step is invalid. The workflow process may have been completed.''
					RETURN
				END
			
				/* Check if the step has already been completed! */
				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iStatus = 3
				BEGIN
					SET @psErrorMessage = ''This workflow step has already been completed.''
					RETURN
				END

				IF @iStatus = 6
				BEGIN
					SET @psErrorMessage = ''This workflow step has timed out.''
					RETURN
				END
			
				IF @iStatus = 0
				BEGIN
					SET @psErrorMessage = ''This workflow step is invalid. It may no longer be required due to the results of other workflow steps.''
					RETURN
				END
			
				SET @psErrorMessage = ''''
			
				SELECT 			
					@piBackColour = isnull(webFormBGColor, 16777166),
					@piBackImage = isnull(webFormBGImageID, 0),
					@piBackImageLocation = isnull(webFormBGImageLocation, 0),
					@piWidth = isnull(webFormWidth, -1),
					@piHeight = isnull(webFormHeight, -1),
					@iWorkflowID = workflowID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.ID = @piElementID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				CREATE TABLE #itemValues (ID integer, value varchar(8000), type integer)	
			
				DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowElementItems.ID,
					ASRSysWorkflowElementItems.itemType,
					ASRSysWorkflowElementItems.dbColumnID,
					ASRSysWorkflowElementItems.dbRecord,
					ASRSysWorkflowElementItems.wfFormIdentifier,
					ASRSysWorkflowElementItems.wfValueIdentifier
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.elementID = @piElementID
					AND (ASRSysWorkflowElementItems.itemType = 1 OR ASRSysWorkflowElementItems.itemType = 4)
			
				OPEN itemCursor
		'


	SET @sSPCode_1 = '		FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iItemType = 1
					BEGIN
						SET @fDeletedValue = 0

							/* Database value. */
						SELECT @sTableName = ASRSysTables.tableName, 
							@iRequiredTableID = ASRSysTables.tableID, 
							@sColumnName = ASRSysColumns.columnName,
							@iDBColumnDataType = ASRSysColumns.dataType
						FROM ASRSysColumns
						INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
						WHERE ASRSysColumns.columnID = @iDBColumnID
			
						SET @iType = @iDBColumnDataType
		
						IF @iDBRecord = 0
						BEGIN
							-- Initiator''s record
							SET @iRecordID = @iInitiatorID
							SET @iParent1TableID = @iInitParent1TableID
							SET @iParent1RecordID = @iInitParent1RecordID
							SET @iParent2TableID = @iInitParent2TableID
							SET @iParent2RecordID = @iInitParent2RecordID

							SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
							FROM ASRSysModuleSetup
							WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_TablePersonnel''
						END			

						IF @iDBRecord = 4
						BEGIN
							-- Trigger record
							SET @iRecordID = @iInitiatorID
							SET @iParent1TableID = @iInitParent1TableID
							SET @iParent1RecordID = @iInitParent1RecordID
							SET @iParent2TableID = @iInitParent2TableID
							SET @iParent2RecordID = @iInitParent2RecordID

							SELECT @iBaseTableID = isnull(WF.baseTable, 0)
							FROM ASRSysWorkflows WF
							INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
								AND WFI.ID = @piInstanceID
						END

						IF @iDBRecord = 1
						BEGIN
							-- Identified record.
							SELECT @iElementType = ASRSysWorkflowElements.type, 
								@iTempElementID = ASRSysWorkflowElements.ID
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
								AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)))
								
							IF @iElementType = 2
							BEGIN
								 -- WebForm
								SELECT @iRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@iBaseTableID = EI.tableID,
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
								INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
								WHERE IV.instanceID = @piInstanceID
									AND IV.identifier = @sWFValueIdentifier
									AND Es.identifier = @sWFFormIdentifier
									AND Es.workflowID = @iWorkflowID
									AND IV.elementID = Es.ID
							END
							ELSE
							BEGIN
								-- StoredData
								SELECT @iRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@iBaseTableID = isnull(Es.dataTableID, 0),
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
									AND IV.identifier = Es.identifier
									AND Es.workflowID = @iWorkflowID
									AND Es.identifier = @sWFFormIdentifier
								WHERE IV.instanceID = @piInstanceID
							END
						END	
						
						SET @iBaseRecordID = @iRecordID

						IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
						BEGIN
							SET @fValidRecordID = 0

							EXEC spASRWo'


	SET @sSPCode_2 = 'rkflowAscendantRecordID
								@iBaseTableID,
								@iBaseRecordID,
								@iParent1TableID,
								@iParent1RecordID,
								@iParent2TableID,
								@iParent2RecordID,
								@iRequiredTableID,
								@iRequiredRecordID	OUTPUT

							SET @iRecordID = @iRequiredRecordID

							IF @iRecordID > 0 
							BEGIN
								EXEC spASRWorkflowValidTableRecord
									@iRequiredTableID,
									@iRecordID,
									@fValidRecordID	OUTPUT
							END

							IF @fValidRecordID = 0
							BEGIN
								IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
								BEGIN
									SELECT @iCount = COUNT(*)
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.columnID = @iDBColumnID

									IF @iCount = 1
									BEGIN
										SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
										FROM ASRSysWorkflowQueueColumns QC
										INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
										WHERE WFQ.instanceID = @piInstanceID
											AND QC.columnID = @iDBColumnID

										SET @fValidRecordID = 1
										SET @fDeletedValue = 1
									END
								END
								ELSE
								BEGIN
									IF @iDBRecord = 1
									BEGIN
										SELECT @iCount = COUNT(*)
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.columnID = @iDBColumnID
											AND IV.elementID = @iTempElementID

										IF @iCount = 1
										BEGIN
											SELECT @sValue = rtrim(ltrim(isnull(IV.value , '''')))
											FROM ASRSysWorkflowInstanceValues IV
											WHERE IV.instanceID = @piInstanceID
												AND IV.columnID = @iDBColumnID
												AND IV.elementID = @iTempElementID

											SET @fValidRecordID = 1
											SET @fDeletedValue = 1
										END
									END
								END
							END

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Web Form item record has been deleted or not selected.''
											
								SET @psErrorMessage = ''Error loading web form. Web Form item record has been deleted or not selected.''
								RETURN
							END
						END
							
						IF @fDeletedValue = 0
						BEGIN
							IF @iDBColumnDataType = 11 -- Date column, need to format into MM\DD\YYYY
							BEGIN
								SET @sSQL = ''SELECT @sValue = convert(varchar(100), '' + @sTableName + ''.'' + @sColumnName + '', 101)''
							END
							ELSE
							BEGIN
								SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName
							END
							
							SET @sSQL = @sSQL +
									'' FROM '' + @sTableName +
									'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
							SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
						END
					END
					ELSE
					BEGIN
						/* Workflow value. */
						SELECT @sValue = ASRSysWorkflowInstanceValues.value, 
							@iType = ASRSysWorkflowElementItems.itemType,
							@iColumnID = ASRSysWorkflowElementItems.lookupColumnID
						FROM ASRSysWorkflowInstanceValues
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
						INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.ElementID
						WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
							AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
							AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
		
						IF @iType = 14 -- Lookup, need to get th'


	SET @sSPCode_3 = 'e column data type
						BEGIN
							SELECT @iType = 
								CASE
									WHEN ASRSysColumns.dataType = -7 THEN 6 -- Logic
									WHEN ASRSysColumns.dataType = 2 THEN 5 -- Numeric
									WHEN ASRSysColumns.dataType = 4 THEN 5 -- Integer
									WHEN ASRSysColumns.dataType = 11 THEN 7 -- Date
									ELSE 3
								END
							FROM ASRSysColumns
							WHERE ASRSysColumns.columnID = @iColumnID
						END
					END
			
					INSERT INTO #itemValues (ID, value, type)
					VALUES (@iID, @sValue, @iType)
			
					FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				END
				CLOSE itemCursor
				DEALLOCATE itemCursor
			
				SELECT thisFormItems.*, 
					#itemValues.value, 
					#itemValues.type AS [sourceItemType]
				FROM ASRSysWorkflowElementItems thisFormItems
				LEFT OUTER JOIN #itemValues ON thisFormItems.ID = #itemValues.ID
				WHERE thisFormItems.elementID = @piElementID
				ORDER BY thisFormItems.ZOrder DESC
				DROP TABLE #itemValues
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3)

	----------------------------------------------------------------------
	-- spASRSubmitWorkflowStep
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRSubmitWorkflowStep
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psFormInput1		varchar(8000),
			@psFormInput2		varchar(8000),
			@psFormElements		varchar(8000)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iIndex1		integer,
				@iIndex2		integer,
				@iID			integer,
				@sID			varchar(8000),
				@sValue		varchar(8000),
				@iElementType		integer,
				@iPreviousElementID	integer,
				@iValue		integer,
				@hResult		integer,
				@sTo			varchar(8000),
				@sMessage		varchar(8000),
				@iEmailID		integer,
				@iEmailRecord		integer,
				@iEmailRecordID	integer,
				@sSQL			nvarchar(4000),
				@iCount		integer,
				@superCursor		cursor,
				@curDelegatedRecords	cursor,
				@fDelegate		bit,
				@fDelegationValid	bit,
				@fCopyDelegateEmail	bit,
				@iDelegateEmailID	integer,
				@iDelegateRecordID	integer,
				@sTemp		varchar(8000),
				@sDelegateTo		varchar(8000),
				@sAllDelegateTo	varchar(8000),
				@iCurrentStepID	int,
				@sDelegatedMessage	varchar(8000),
				@iTemp		integer, 
				@iPrevElementType	integer,
				@iWorkflowID		integer,
				@sRecSelIdentifier	varchar(8000),
				@sRecSelWebFormIdentifier	varchar(8000), 
				@iStepID int,
				@iElementID int,
				@sUserName varchar(8000),
				@sUserEmail varchar(8000), 
				@sValueDescription	varchar(8000),
				@iTableID		integer,
				@iRecDescID		integer,
				@sEvalRecDesc	varchar(8000),
				@sExecString		nvarchar(4000),
				@sParamDefinition	nvarchar(500),
				@sIdentifier		varchar(8000),
				@iItemType		integer,
				@iDataAction		integer, 
				@fValidRecordID	bit,
				@iEmailTableID integer,
				@iEmailType integer,
				@iBaseTableID	integer,
				@iBaseRecordID	integer,
				@iRequiredRecordID	integer,
				@iParent1TableID	int,
				@iParent1RecordID	int,
				@iParent2TableID	int,
				@iParent2RecordID	int,
				@iTempElementID		integer,
				@iTrueFlowType	integer,
				@iExprID		integer,
				@iResultType	integer,
				@sResult		varchar(8000),
				@fResult		bit,
				@dtResult		datetime,
				@fltResult		float,
				@sEmailSubject	varchar(200)

			SELECT @iCurrentStepID = ID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

			SET @fCopyDelegateEmail = 1
			SELECT @sTemp = LTRIM(RTRIM(UPPER(ISNULL(parameterValue, ''''))))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_CopyDelegateEmail''
			IF @sTemp = ''FALSE''
			BEGIN
				SET @fCopyDelegateEmail = 0
			END

			SET @iDelegateEmailID = 0
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail''
			SET @iDelegateEmailID = convert(integer, @sTemp)

			SET @psFormElements = ''''
						
			-- Get the type of the given element 
			SELECT @iElementType = E.type,
				@iEmailID = E.emailID,
				@iEmailRecord = E.emailRecord, 
				@iWorkflowID = E.workflowID,
				@sRecSelIdentifier = E.RecSelIdentifier, 
				@sRecSelWebFormIdentifier = E.RecSelWebFormIdentifier, 
				@iTableID = E.dataTableID,
				@iDataAction = E.dataAction, 
				@iTrueFlowType = isnull(E.trueFlowType, 0), 
				@iExprID = isnull(E.trueFlowExprID, 0), 
				@sEmailSubject = ISNULL(E.emailSubject, '''')
			FROM ASRSysWorkflowElements E
			WHERE E.ID = @piElementID

			IF @iElementType = 5 -- Stored Data element
			BEGIN
				SET @sValue = @psFormInput1
				SET @sValueDescription = ''''
				SET @sMessage = ''Successfully '' +
					CASE
						WHEN @iDataAction = 0 THEN ''inserted''
						WHEN @iDataAction = 1 THEN ''updated''
						ELSE ''deleted''
					END + '' record''

				IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
				BEGIN
					SET @sValueDescription = @psFormInput2
				END
				ELSE
				BEGIN
					SE'


	SET @sSPCode_1 = 'T @iTemp = convert(integer, @sValue)
					IF @iTemp > 0 
					BEGIN	
						EXEC spASRRecordDescription 
							@iTableID,
							@iTemp,
							@sEvalRecDesc OUTPUT
						IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc
					END
				END

				IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')''

				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.value = @sValue, 
					ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
					AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0
			END
			ELSE
			BEGIN
				-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
				WHILE (charindex(CHAR(9), @psFormInput1) > 0) OR (charindex(CHAR(9), @psFormInput2) > 0)
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1)
					IF @iIndex1 > 0
					BEGIN
						SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''')

						SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1)
						IF @iIndex2 > 0	
						BEGIN
							SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1)

							SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2)
						END
						ELSE
						BEGIN
							SET @iIndex2 = charindex(CHAR(9), @psFormInput2)
							SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, len(@psFormInput1)-@iIndex1) +
								LEFT(@psFormInput2, @iIndex2-1)

							SET @psFormInput1 = ''''
							SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
						END
					END
					ELSE
					BEGIN
						SET @iIndex1 = charindex(CHAR(9), @psFormInput2)
						SET @iIndex2 = charindex(CHAR(9), @psFormInput2, @iIndex1+1)

						SET @sID = replace(@psFormInput1, '''''''', '''''''''''') +
							replace(LEFT(@psFormInput2, @iIndex1-1), '''''''', '''''''''''')
						SET @sValue = SUBSTRING(@psFormInput2, @iIndex1+1, @iIndex2-@iIndex1-1)

						SET @psFormInput1 = ''''
						SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
					END
					SET @sValue = left(@sValue, 1000)

					--Get the record description (for RecordSelectors only)
					SET @sValueDescription = ''''

					-- Get the WebForm item type, etc.
					SELECT @sIdentifier = EI.identifier,
						@iItemType = EI.itemType,
						@iTableID = EI.tableID
					FROM ASRSysWorkflowElementItems EI
					WHERE EI.ID = convert(integer, @sID)

					SET @iParent1TableID = 0
					SET @iParent1RecordID = 0
					SET @iParent2TableID = 0
					SET @iParent2RecordID = 0

					IF @iItemType = 11 -- Record Selector
					BEGIN
						-- Get the table record description ID. 
						SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
						FROM ASRSysTables 
						WHERE ASRSysTables.tableID = @iTableID

						SET @iTemp = convert(integer, isnull(@sValue, ''0''))

						-- Get the record description. 
						IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
						BEGIN
							SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(4000), @iRecDescID) + '' @recDesc OUTPUT, @recID''
							SET @sParamDefinition = N''@recDesc varchar(8000) OUTPUT, @recID integer''
							EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc
						END

						-- Record the selected record''s parent details.
						exec spASRGetParentDetails
							@iTableID,
							@iTemp,
							@iParent1TableID	OUTPUT,
							@iParent1RecordID	OUTPUT,
							@iParent2TableID	OUTPUT,
							@iParent2RecordID	OUTPUT
					END

					UP'


	SET @sSPCode_2 = 'DATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
						ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
						ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
						ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
						ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier
				END
			END
					
			SET @hResult = 0
			SET @sTo = ''''
		
			IF @iElementType = 3 -- Email element
			BEGIN
				-- Get the email recipient. 
				SET @sTo = ''''
				SET @iEmailRecordID = 0
				SET @sSQL = ''spASRSysEmailAddr''

				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @fValidRecordID = 1

					SELECT @iEmailTableID = isnull(tableID, 0),
						@iEmailType = isnull(type, 0)
					FROM ASRSysEmailAddress
					WHERE emailID = @iEmailID

					IF @iEmailType = 0 
					BEGIN
						SET @iEmailRecordID = 0
					END
					ELSE
					BEGIN
						SET @iTempElementID = 0

						-- Get the record ID required. 
						IF (@iEmailRecord = 0) OR (@iEmailRecord = 4)
						BEGIN
							/* Initiator record. */
							SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID,
								@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
								@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
								@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
								@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
							FROM ASRSysWorkflowInstances
							WHERE ASRSysWorkflowInstances.ID = @piInstanceID

							SET @iBaseRecordID = @iEmailRecordID

							IF @iEmailRecord = 4
							BEGIN
								-- Trigger record
								SELECT @iBaseTableID = isnull(WF.baseTable, 0)
								FROM ASRSysWorkflows WF
								INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
									AND WFI.ID = @piInstanceID
							END
							ELSE
							BEGIN
								-- Initiator''s record
								SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
								FROM ASRSysModuleSetup
								WHERE moduleKey = ''MODULE_WORKFLOW''
								AND parameterKey = ''Param_TablePersonnel''
							END
						END
		
						IF @iEmailRecord = 1
						BEGIN
							SELECT @iPrevElementType = ASRSysWorkflowElements.type,
								@iTempElementID = ID
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
								AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))

							IF @iPrevElementType = 2
							BEGIN
								 -- WebForm
								SELECT @iEmailRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@iBaseTableID = EI.tableID,
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
								INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
								WHERE IV.instanceID = @piInstanceID
									AND IV.identifier = @sRecSelIdentifier
									AND Es.identifier = @sRecSelWebFormIdentifier
									AND Es.workflowID = @iWorkflowID
									AND IV.elementID = Es.ID
							END
							ELSE
							BEGIN
								-- StoredData
								SELECT @iEmailRecordID = 
									CASE
										WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
										ELSE 0
									END,
									@'


	SET @sSPCode_3 = 'iBaseTableID = isnull(Es.dataTableID, 0),
									@iParent1TableID = IV.parent1TableID,
									@iParent1RecordID = IV.parent1RecordID,
									@iParent2TableID = IV.parent2TableID,
									@iParent2RecordID = IV.parent2RecordID
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
									AND IV.identifier = Es.identifier
									AND Es.workflowID = @iWorkflowID
									AND Es.identifier = @sRecSelWebFormIdentifier
								WHERE IV.instanceID = @piInstanceID
							END

							SET @iBaseRecordID = @iEmailRecordID
						END

						SET @fValidRecordID = 1
						IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
						BEGIN
							SET @fValidRecordID = 0

							EXEC spASRWorkflowAscendantRecordID
								@iBaseTableID,
								@iBaseRecordID,
								@iParent1TableID,
								@iParent1RecordID,
								@iParent2TableID,
								@iParent2RecordID,
								@iEmailTableID,
								@iRequiredRecordID	OUTPUT

							SET @iEmailRecordID = @iRequiredRecordID

							IF @iRequiredRecordID > 0 
							BEGIN
								EXEC spASRWorkflowValidTableRecord
									@iEmailTableID,
									@iEmailRecordID,
									@fValidRecordID	OUTPUT
							END

							IF @fValidRecordID = 0
							BEGIN
								IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
								BEGIN
									SELECT @sTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.emailID = @iEmailID

									IF len(@sTo) > 0 SET @fValidRecordID = 1
								END
								ELSE
								BEGIN
									IF @iEmailRecord = 1
									BEGIN
										SELECT @sTo = rtrim(ltrim(isnull(IV.value , '''')))
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.emailID = @iEmailID
											AND IV.elementID = @iTempElementID
										IF len(@sTo) > 0 SET @fValidRecordID = 1

									END
								END
							END

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Email record has been deleted or not selected.''
											
								SET @hResult = -1
							END
						END
					END

					IF @fValidRecordID = 1
					BEGIN
						/* Get the recipient address. */
						IF len(@sTo) = 0
						BEGIN
							EXEC @hResult = @sSQL @sTo OUTPUT, @iEmailID, @iEmailRecordID
							IF @sTo IS null SET @sTo = ''''
						END

						IF LEN(rtrim(ltrim(@sTo))) = 0
						BEGIN
							-- Email step failure if no known recipient.
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''No email recipient.''
										
							SET @hResult = -1
						END
					END
				END
		
				IF LEN(rtrim(ltrim(@sTo))) > 0
				BEGIN
					IF (rtrim(ltrim(@sTo)) = ''@'')
						OR (charindex('' @ '', @sTo) > 0)
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.userEmail = @sTo
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

						EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Invalid email recipient.''
						
						SET @hResult = -1
					END
					ELSE
					BEGIN
						/* Build the email message. */
						EXEC spASRGetWorkflowEmailMessage @piInstanceID, @piElementID, @sMessage OUTPUT, @fValidRecordID OUTPUT
		
						IF @fValidRecordID = 1
						BEGIN
							exec spASRDelegateWorkflowEmail 
								@sTo,
								@sMessage,
								@iCurrentStepID,
								@sEmailSubject,
		'


	SET @sSPCode_4 = '						0
						END
						ELSE
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Email item database value record has been deleted or not selected.''
										
							SET @hResult = -1
						END
					END
				END
			END
		
			IF @hResult = 0
			BEGIN
				/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.userEmail = CASE
						WHEN @iElementType = 3 THEN @sTo
						ELSE ASRSysWorkflowInstanceSteps.userEmail
					END,
					ASRSysWorkflowInstanceSteps.message = CASE
						WHEN @iElementType = 3 THEN LEFT(@sMessage, 8000)
						WHEN @iElementType = 5 THEN LEFT(@sMessage, 8000)
						ELSE ''''
					END
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iElementType = 4 -- Decision element
				BEGIN
					IF @iTrueFlowType = 1
					BEGIN
						-- Decision Element flow determined by a calculation
						EXEC [spASRSysWorkflowCalculation]
							@piInstanceID,
							@iExprID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT

						SET @iValue = convert(integer, @fResult)
					END
					ELSE
					BEGIN
						-- Decision Element flow determined by a button in a preceding web form
						SET @iPrevElementType = 4 -- Decision element
						SET @iPreviousElementID = @piElementID

						WHILE (@iPrevElementType = 4)
						BEGIN
							/* Get the ID of the elements that precede the Decision element. */
							CREATE TABLE #precedingElements (elementID integer)
		
							EXEC spASRGetPrecedingWorkflowElements @iPreviousElementID, @superCursor OUTPUT
			
							FETCH NEXT FROM @superCursor INTO @iTemp
							WHILE (@@fetch_status = 0)
							BEGIN
								INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
								
								FETCH NEXT FROM @superCursor INTO @iTemp 
							END
							CLOSE @superCursor
							DEALLOCATE @superCursor
		
							SELECT TOP 1 @iPreviousElementID = elementID
							FROM #precedingElements
		
							DROP TABLE #precedingElements
					
							SELECT @iPrevElementType = ASRSysWorkflowElements.type
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.ID = @iPreviousElementID
						END
					
						SELECT @iValue = 
							CASE
								WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
								ELSE 0
							END
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
						WHERE IV.elementID = @iPreviousElementID
							AND IV.instanceid = @piInstanceID
							AND E.ID = @piElementID
					END
				
					IF @iValue IS null SET @iValue = 0
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
					CREATE TABLE #succeedingElements2 (elementID integer)
		
					EXEC spASRGetDecisionSucceedingWorkflowElements @piElementID, @iValue, @superCursor OUTPUT
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #succeedingElements2 (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 1,
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
						ASR'


	SET @sSPCode_5 = 'SysWorkflowInstanceSteps.completionDateTime = null
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT #succeedingElements2.elementID 
							FROM #succeedingElements2)
						AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 3)
		
					DROP TABLE #succeedingElements2
				END
				ELSE
				BEGIN
					CREATE TABLE #succeedingElements (elementID integer)
		
					EXEC spASRActionOrsAndGetSucceedingWorkflowElements @piInstanceID, @piElementID, 0, @superCursor OUTPUT
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor

					-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
					IF @iElementType = 2 -- WebForm
					BEGIN
						SELECT @sUserName = isnull(WIS.userName, ''''),
							@sUserEmail = isnull(WIS.userEmail, '''')
						FROM ASRSysWorkflowInstanceSteps WIS
						WHERE WIS.instanceID = @piInstanceID
							AND WIS.elementID = @piElementID

						-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
						DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT ASRSysWorkflowInstanceSteps.ID,
							ASRSysWorkflowInstanceSteps.elementID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowElements.type = 2
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 3)
						OPEN formsCursor
						FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
						WHILE (@@fetch_status = 0) 
						BEGIN
							SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)

							DELETE FROM ASRSysWorkflowStepDelegation
							WHERE stepID = @iStepID

							INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
								(SELECT WSD.delegateEmail, @iStepID
								FROM ASRSysWorkflowStepDelegation WSD
								WHERE WSD.stepID = @iCurrentStepID)
							
							-- Change the step status to be 2 (pending user input). 
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 2, 
								ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.completionDateTime = null,
								ASRSysWorkflowInstanceSteps.userName = @sUserName,
								ASRSysWorkflowInstanceSteps.userEmail = @sUserEmail 
							WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 3)
							
							FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
						END
						CLOSE formsCursor
						DEALLOCATE formsCursor

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.completionDateTime = null
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
								(SELECT ASRSysWorkflowElements.ID
								FROM A'


	SET @sSPCode_6 = 'SRSysWorkflowElements
								WHERE ASRSysWorkflowElements.type = 2)
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 3)
					END
					ELSE
					BEGIN
						DELETE FROM ASRSysWorkflowStepDelegation
						WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
							FROM ASRSysWorkflowInstanceSteps
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT #succeedingElements.elementID
									FROM #succeedingElements)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 3))
						
						INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
						(SELECT WSD.delegateEmail,
							SuccWIS.ID
						FROM ASRSysWorkflowStepDelegation WSD
						INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
						INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
							AND SuccWIS.elementID IN (SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND (SuccWIS.status = 0
								OR SuccWIS.status = 2
								OR SuccWIS.status = 3)
						INNER JOIN ASRSysWorkflowElements SuccWE ON SuccWIS.elementID = SuccWE.ID
							AND SuccWE.type = 2
						WHERE WSD.stepID = @iCurrentStepID)

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.completionDateTime = null,
							ASRSysWorkflowInstanceSteps.userEmail = CASE
								WHEN (SELECT ASRSysWorkflowElements.type 
									FROM ASRSysWorkflowElements 
									WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
								ELSE ASRSysWorkflowInstanceSteps.userEmail
							END
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 3)
					END
					
					DROP TABLE #succeedingElements
				END
			
				-- Set activated Web Forms to be ''pending'' (to be done by the user) 
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2)
		
				-- Set activated Terminators to be ''completed'' 
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate()
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowElements.type = 1)
		
				-- Count how many terminators have completed. ie. if the workflow has completed. 
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type ='


	SET @sSPCode_7 = ' 1
							
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					
					-- Steps pending action are no longer required.
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
							OR ASRSysWorkflowInstanceSteps.status = 2) -- 2 = Pending User Action
				END

				IF @iElementType = 3 -- Email element
					OR @iElementType = 5 -- Stored Data element
				BEGIN
					exec spASREmailImmediate ''HR Pro Workflow''
				END
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5
		+ @sSPCode_6
		+ @sSPCode_7)

	----------------------------------------------------------------------
	-- spASRWorkflowValidRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowValidRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowValidRecord]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowValidRecord]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRWorkflowValidRecord
			@piInstanceID				integer,
			@piRecordType				integer,
			@piRecordID					integer,
			@sElementIdentifier			varchar(8000),
			@sElementItemIdentifier		varchar(8000),
			@pfValid					bit				OUTPUT
		AS
		BEGIN
			DECLARE
				@sSQL					nvarchar(4000),
				@iTableID				integer,
				@sTableName				nvarchar(4000),
				@iWorkflowID			integer,
				@sParam					nvarchar(500),
				@iRecCount				integer,
				@iElementType			integer
		
			SET @pfValid = 0
		
			SELECT @iWorkflowID = WF.ID,
				@iTableID = 
					CASE
						WHEN @piRecordType = 4 THEN isnull(WF.baseTable, 0)
						ELSE 0
					END
			FROM ASRSysWorkflows WF
			INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
				AND WFI.ID = @piInstanceID
		
			IF @piRecordType = 0
			BEGIN
				-- Initiator''s record
				SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel''
			END
		
			IF @piRecordType = 1
			BEGIN
				-- Identified record
				SELECT @iElementType = ASRSysWorkflowElements.type,
					@iTableID = 
						CASE
							WHEN ASRSysWorkflowElements.type = 5 THEN isnull(ASRSysWorkflowElements.dataTableID, 0)
							ELSE 0
						END
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)))
		
				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @iTableID = WFEI.tableID
					FROM ASRSysWorkflowElementItems WFEI
					INNER JOIN ASRSysWorkflowElements WFE ON WFEI.elementID = WFE.ID
						AND WFE.identifier = @sElementIdentifier
						AND WFE.workflowID = @iWorkflowID
					WHERE WFEI.identifier = @sElementItemIdentifier
				END
			END
		
			exec spASRWorkflowValidTableRecord
				@iTableID,
				@piRecordID,
				@pfValid	OUTPUT
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetActualUserDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetActualUserDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetActualUserDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetActualUserDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRGetActualUserDetails
		(
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
				AND lo.loginname = system_user
				AND CASE
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
				END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
		
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
						END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
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
					END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
			END
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetWorkflowGridItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowGridItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowGridItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowGridItems]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowGridItems
			(
				@piInstanceID		integer,
				@piElementItemID	integer, 
				@pfOK				bit	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iTableID 		integer,
					@iOrderID			integer,
					@iFilterID			integer,
					@sFilterSQL			varchar(8000),
					@sFilterUDF			varchar(8000),
					@sRecSelWebFormIdentifier	varchar(200),
					@sRecSelIdentifier	varchar(200),
					@iDBRecord		integer,
					@iInitiatorID		integer,
					@sSQL			varchar(8000),
					@sOrderItemType	varchar(8000),
					@sSelectSQL		varchar(8000),
					@sOrderSQL		varchar(8000),
					@sBaseTableName	sysname,
					@fAscending		bit,
					@sColumnName		sysname,
					@sTempTableName	sysname,
					@iTempTableID		integer,
					@iTempTableType	integer,
					@iTempCount		integer,
					@iDataType		integer,
					@iRecordID		integer,
					@iPersonnelTableID	integer,
					@iWorkflowID		integer,
					@iElementType		integer, 
					@fValidRecordID	bit,
					@iElementID	integer,
					@iBaseTableID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iRecordTableID		int,
					@iTriggerTableID	integer
			
				SET @pfOK = 1

				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT 			
					@iTableID = ASRSysWorkflowElementItems.tableID,
					@iElementID = ASRSysWorkflowElementItems.elementiD,
					@sRecSelWebFormIdentifier = isnull(ASRSysWorkflowElementItems.wfFormIdentifier, ''''),
					@sRecSelIdentifier = isnull(ASRSysWorkflowElementItems.wfValueIdentifier, 0),
					@iDBRecord = ASRSysWorkflowElementItems.dbRecord,
					@iOrderID = 
						CASE
							WHEN isnull(ASRSysWorkflowElementItems.recordOrderID, 0) > 0 THEN ASRSysWorkflowElementItems.recordOrderID
							ELSE ASRSysTables.defaultOrderID
						END,
					@iFilterID = isnull(ASRSysWorkflowElementItems.recordFilterID, 0),
					@iRecordTableID = ASRSysWorkflowElementItems.recordTableID,
					@sBaseTableName = ASRSysTables.tableName
				FROM ASRSysWorkflowElementItems
				INNER JOIN ASRSysTables ON ASRSysWorkflowElementItems.tableID = ASRSysTables.tableID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID, 
					@iTriggerTableID = ASRSysWorkflows.baseTable,
					@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SET @sSelectSQL = ''''
				SET @sOrderSQL = ''''
			
				CREATE TABLE #joinParents
				(
					tableID		integer
				)	
			
				DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT 
					ASRSysColumns.columnName,
					ASRSysColumns.dataType,
					ASRSysColumns.tableID,
					ASRSysTables.tableType,
					ASRSysTables.tableName,
					upper(isnull(ASRSysOrderItems.type, '''')),
					ASRSysOrderItems.ascending
				FROM ASRSysOrderItems
				INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
				INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
				WHERE ASRSysOrderItems.orderID = @iOrderID
				ORDER BY ASRSysOrderItems.type,
					ASRSysOrderItems.sequence
			
				OPEN orderCursor
				FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @sOrderItemType = ''F''
					BEG'


	SET @sSPCode_1 = 'IN
						SET @sSelectSQL = @sSelectSQL +
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN '',''
								ELSE ''''
							END +
							@sTempTableName + ''.'' + @sColumnName
					END
			
					IF @sOrderItemType = ''O''
					BEGIN
						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN '',''
								ELSE '' ''
							END + 
							@sTempTableName + ''.'' + @sColumnName +
							CASE 
								WHEN @fAscending = 0 THEN '' DESC'' 
								ELSE '''' 
							END				
					END
			
					IF @iTableID <> @iTempTableID
					BEGIN
						SELECT @iTempCount = COUNT(tableID)
						FROM #joinParents
						WHERE tableID = @iTempTableID
			
						IF @iTempCount = 0
						BEGIN
							INSERT INTO #joinParents (tableID) VALUES(@iTempTableID)
						END
					END
			
					FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				END
				CLOSE orderCursor
				DEALLOCATE orderCursor
			
				IF len(@sSelectSQL) > 0 
				BEGIN
					SET @sSelectSQL = ''SELECT '' + @sSelectSQL + '','' +
						@sBaseTableName + ''.id'' +
					'' FROM '' + @sBaseTableName
			
					DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ASRSysTables.tableName, 
						#joinParents.tableID
					FROM #joinParents
					INNER JOIN ASRSysTables ON #joinParents.tableID = ASRSysTables.tableID
			
					OPEN joinCursor
					FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @sSelectSQL = @sSelectSQL + 
							'' LEFT OUTER JOIN '' + @sTempTableName + '' ON '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iTempTableID) + '' = '' + @sTempTableName + ''.ID''
			
						FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					END
					CLOSE joinCursor
					DEALLOCATE joinCursor
			
					IF @iDBRecord = 0 -- ie. based on the initiator''s record
					BEGIN
						SET @iBaseTableID = @iPersonnelTableID
						SET @iRecordID = @iInitiatorID
					END

					IF @iDBRecord = 4 -- ie. based on the triggered record
					BEGIN
						SET @iBaseTableID = @iTriggerTableID
						SET @iRecordID = @iInitiatorID
					END
			
					IF @iDBRecord = 1 -- ie. based on a previously identified record
					BEGIN
						SELECT @iElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
		
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @iRecordID = 
								CASE
									WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
									ELSE 0
								END,
								@iTempTableID = EI.tableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sRecSelIdentifier
								AND Es.identifier = @sRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @iRecordID = 
								CASE
									WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
									ELSE 0
								END,
								@iTempTableID = Es.dataTableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN AS'


	SET @sSPCode_2 = 'RSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
			
						SET @iBaseTableID = @iTempTableID
					END
			
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						EXEC spASRWorkflowAscendantRecordID
							@iBaseTableID,
							@iRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRecordTableID,
							@iRecordID	OUTPUT

						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iRecordTableID) + '' = '' + convert(varchar(100), @iRecordID)

						SET @fValidRecordID = 1

						EXEC spASRWorkflowValidTableRecord
							@iRecordTableID,
							@iRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID  = 0
						BEGIN
							SET @pfOK = 0

							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @iElementID, ''Web Form record selector item record has been deleted or not selected.''
							
							-- Need to return a recordset of some kind.
							SELECT '''' AS ''Error''

							RETURN
						END
					END

					IF @iFilterID > 0 
					BEGIN
						SET @sFilterUDF = ''dbo.udf_ASRWFExpr_'' + convert(varchar(8000), @iFilterID)

						IF EXISTS(
							SELECT Name
							FROM sysobjects
							WHERE id = object_id(@sFilterUDF)
							AND sysstat & 0xf = 0)
						BEGIN
							SET @sFilterSQL = 
								CASE
									WHEN (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4) THEN '' AND ''
									ELSE '' WHERE ''
								END 
								+ @sBaseTableName + ''.ID  IN (SELECT id FROM '' + @sFilterUDF + ''('' + convert(varchar(8000), @piInstanceID) + ''))''
						END
					END

					SET @sOrderSQL = '' ORDER BY '' + @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						@sBaseTableName + ''.ID''

					EXEC (@sSelectSQL 
						+ @sFilterSQL
						+ @sOrderSQL)
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRInstantiateWorkflow
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRInstantiateWorkflow
		(
			@piWorkflowID	integer,			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(8000)	OUTPUT,
			@psMessage		varchar(8000)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID		integer,
				@iStepID		integer,
				@iElementID		integer,
				@iRecordID		integer,
				@iRecordCount		integer,
				@sSQL			nvarchar(4000),
				@hResult		integer,
				@sActualLoginName sysname,
				@fUsesInitiator	bit, 
				@iTemp int,
				@iStartElementID int,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms	varchar(8000),
				@iCount	integer,
				@superCursor	cursor	
		
			SET @iInitiatorID = 0
			SET @psFormElements = ''''
			SET @psMessage = ''''
			SET @iParent1TableID = 0
			SET @iParent1RecordID = 0
			SET @iParent2TableID = 0
			SET @iParent2RecordID = 0
		
			SET @sActualLoginName = SUSER_SNAME()
			
			SET @sSQL = ''spASRSysGetCurrentUserRecordID''
			IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
			BEGIN
				SET @hResult = 0
		
				EXEC @hResult = @sSQL 
					@iRecordID OUTPUT,
					@iRecordCount OUTPUT
			END
		
			IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
			IF @iInitiatorID = 0 
			BEGIN
				/* Unable to determine the initiator''s record ID. Is it needed anyway? */
				EXEC spASRWorkflowUsesInitiator
					@piWorkflowID,
					@fUsesInitiator OUTPUT	
			
				IF @fUsesInitiator = 1
				BEGIN
					IF @iRecordCount = 0
					BEGIN
						/* No records for the initiator. */
						SET @psMessage = ''Unable to locate your personnel record.''
					END
					IF @iRecordCount > 1
					BEGIN
						/* More than one record for the initiator. */
						SET @psMessage = ''You have more than one personnel record.''
					END
			
					RETURN
				END	
			END
			ELSE
			BEGIN
				SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel''

				exec spASRGetParentDetails
					@iTableID,
					@iInitiatorID,
					@iParent1TableID	OUTPUT,
					@iParent1RecordID	OUTPUT,
					@iParent2TableID	OUTPUT,
					@iParent2RecordID	OUTPUT
			END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO ASRSysWorkflowInstances (workflowID, 
				initiatorID, 
				status, 
				userName, 
				parent1TableID,
				parent1RecordID,
				parent2TableID,
				parent2RecordID)
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID)
						
			SELECT @piInstanceID = MAX(id)
			FROM ASRSysWorkflowInstances
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */

			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID

			CREATE TABLE #succeedingElements (elementID integer)

			EXEC spASRGetSucceedingWorkflowElements @iStartElementID, @superCursor OUTPUT

			FETCH NEXT FROM @superCursor INTO @iTemp
			WHILE (@@fetch_status = 0)
			BEGIN
				INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
				
				FETCH NEXT FROM @superCursor INTO @iTemp 
			END
			CLOSE @superCursor
			DEALLOCATE @superCursor
		
			INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
			'


	SET @sSPCode_1 = '			FROM #succeedingElements) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
						FROM #succeedingElements) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID
		
			DROP TABLE #succeedingElements

			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 7 -- Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
			WHILE @iCount > 0 
			BEGIN
				DECLARE orCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 7 -- Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		

				OPEN orCursor
				FETCH NEXT FROM orCursor INTO @iElementID
				WHILE (@@fetch_status = 0) 
				BEGIN
					EXEC spASRSubmitWorkflowStep @piInstanceID, @iElementID, '''', '''', @sForms OUTPUT

					FETCH NEXT FROM orCursor INTO @iElementID
				END
				CLOSE orCursor
				DEALLOCATE orCursor

				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 7 -- Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
			END						

			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM ASRSysWorkflowInstanceSteps
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWor'


	SET @sSPCode_2 = 'kflowInstanceSteps.instanceID = @piInstanceID		
		
			OPEN formsCursor
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)
		
				/* Change the step''s status to be 2 (pending user input). */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2, 
					userName = @sActualLoginName
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
			END
			CLOSE formsCursor
			DEALLOCATE formsCursor
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRInstantiateTriggeredWorkflows
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateTriggeredWorkflows]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRInstantiateTriggeredWorkflows
		AS
		BEGIN
			DECLARE
				@iQueueID			integer,
				@iWorkflowID		integer,
				@iRecordID			integer,
				@iInstanceID		integer,
				@iStartElementID	integer,
				@iTemp				integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer,
				@superCursor		cursor	
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID,
				Q.parent1TableID,
				Q.parent1RecordID,
				Q.parent2TableID,
				Q.parent2RecordID
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, 
					initiatorID, 
					status, 
					userName, 
					parent1TableID,
					parent1RecordID,
					parent2TableID,
					parent2RecordID)
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					''<Triggered>'',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID)
								
				SELECT @iInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
				
				UPDATE ASRSysWorkflowQueue
				SET instanceID = @iInstanceID
				WHERE queueID = @iQueueID	

				-- Create the Workflow Instance Steps records. 
				-- Set the first steps'' status to be 1 (pending Workflow Engine action). 
				-- Set all subsequent steps'' status to be 0 (on hold). */
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @iWorkflowID
		
				CREATE TABLE #succeedingElements (elementID integer)
		
				EXEC spASRGetSucceedingWorkflowElements @iStartElementID, @superCursor OUTPUT
		
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
		
				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime)
				SELECT 
					@iInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 3
						WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
							FROM #succeedingElements) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
							FROM #succeedingElements) THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						ELSE null
					END
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
				
				DROP TABLE #succeedingElements
		
				-- Create the Workflow Instance Value records. 
				INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
				SELECT @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElementItems.identifier
				FROM ASRSysWorkflowElementItems 
				INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowEle'


	SET @sSPCode_1 = 'ments.type = 2
					AND (ASRSysWorkflowElementItems.itemType = 3 
						OR ASRSysWorkflowElementItems.itemType = 5
						OR ASRSysWorkflowElementItems.itemType = 6
						OR ASRSysWorkflowElementItems.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 13
						OR ASRSysWorkflowElementItems.itemType = 14
						OR ASRSysWorkflowElementItems.itemType = 15
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 5						
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END
		'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRWorkflowLogPurge
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowLogPurge]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowLogPurge]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowLogPurge]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE spASRWorkflowLogPurge 
		AS
		BEGIN
			DECLARE 
				@sUnit			char(1),
				@iPeriod		int,
				@dtPurgeDate	datetime,
				@dtToday		datetime
			
			-- Get purge period details
			SELECT @sUnit = unit, 
				@iPeriod = (period * -1)
			FROM ASRSysPurgePeriods 
			WHERE purgeKey =  ''WORKFLOW''
		
			IF (@sUnit IS NOT NULL) AND (@iPeriod IS NOT NULL)
			BEGIN
				SET @dtToday = convert(datetime,convert(varchar,getdate(),101))
			
				-- Calculate the purge date 
				SET @dtPurgeDate = 
					CASE 
						WHEN @sUnit = ''D'' THEN dateadd(dd, @iPeriod, @dtToday)
						WHEN @sUnit = ''W'' THEN dateadd(ww, @iPeriod, @dtToday)
						WHEN @sUnit = ''M'' THEN dateadd(mm, @iPeriod, @dtToday)
						ELSE dateadd(yy, @iPeriod, @dtToday)
					END
		
				DELETE FROM ASRSysWorkflowInstances 
				WHERE NOT completionDateTime IS null
					AND convert(datetime, convert(varchar(20), completionDateTime, 101)) <= convert(datetime, convert(varchar(20), @dtPurgeDate, 101))
			END
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowStepDescription
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowStepDescription]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowStepDescription]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowStepDescription]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRWorkflowStepDescription
		(
			@piInstanceStepID	integer,
			@psDescription		varchar(8000)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInstanceID	integer,
				@iExprID		integer,
				@iResultType	integer,
				@sResult		varchar(8000),
				@fResult		bit,
				@dtResult		datetime,
				@fltResult		float
		
			-- Get the InstanceID and associated DescriptionExprID of the given step
			SELECT @iInstanceID = isnull(WIS.instanceID, 0),
				@iExprID = isnull(WEs.descriptionExprID, 0)
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WEs ON WIS.elementID = WEs.ID
			WHERE WIS.ID = @piInstanceStepID
		
			IF @iExprID > 0
			BEGIN
				EXEC [spASRSysWorkflowCalculation]
					@iInstanceID,
					@iExprID,
					@iResultType OUTPUT,
					@sResult OUTPUT,
					@fResult OUTPUT,
					@dtResult OUTPUT,
					@fltResult OUTPUT
			END
		
			SELECT @psDescription = isnull(@sResult, '''')
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRCancelPendingPrecedingWorkflowElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRCancelPendingPrecedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRCancelPendingPrecedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRCancelPendingPrecedingWorkflowElements]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRCancelPendingPrecedingWorkflowElements
			(
				@piInstanceID			integer,
				@piElementID			integer,
				@piSourceElementID		integer
			)
			AS
			BEGIN
				/* Cancel (ie. set status to 0 for all workflow pending (ie. status 1 or 2) elements that precede the given element.
				This ignores connection elements.
				NB. This does work for elements with multiple inbound flows. */
				DECLARE
					@iConnectorPairID	integer,
					@iElementID		integer,
					@iStepID		integer,
					@superCursor		cursor,
					@iTemp		integer
				
				CREATE TABLE #precedingElements (elementID integer)
			
				EXEC spASRGetPrecedingWorkflowElements @piElementID, @superCursor output
				
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
					
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
			
				/* Return the recordset of preceding elements. */
				DECLARE elementsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					S.ID
				FROM #precedingElements E
				INNER JOIN ASRSysWorkflowInstanceSteps S ON E.elementID = S.elementID
				WHERE S.instanceID = @piInstanceID
					AND (S.status = 0 OR S.status = 1 OR S.status = 2)
					AND (E.elementID <> @piSourceElementID)

				OPEN elementsCursor
				FETCH NEXT FROM elementsCursor INTO @iElementID, @iStepID
				WHILE (@@fetch_status = 0)
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 0
					WHERE ID = @iStepID
			
					EXEC spASRCancelPendingPrecedingWorkflowElements @piInstanceID, @iElementID, @piSourceElementID			
			
					FETCH NEXT FROM elementsCursor INTO @iElementID, @iStepID
				END
				CLOSE elementsCursor
				DEALLOCATE elementsCursor
			
				DROP TABLE #precedingElements
			END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRActionActiveWorkflowSteps
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRActionActiveWorkflowSteps
		AS
		BEGIN
			/* Return a recordset of the workflow steps that need to be actioned by the Workflow service.
			Action any that can be actioned immediately. */
			DECLARE
				@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
				@iElementType		integer,
				@iInstanceID		integer,
				@iElementID			integer,
				@iStepID			integer,
				@iCount				integer,
				@sStatus			bit,
				@sMessage			varchar(8000),
				@superCursor		cursor,
				@superCursor2		cursor,
				@iTemp				integer, 
				@iTemp2				integer, 
				@sForms 			varchar(8000), 
				@iType				integer,
				@iDecisionFlow		integer,
				@iInvalidDecisionCount	integer
		
			DECLARE stepsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.type,
				S.instanceID,
				E.ID,
				S.ID
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE S.status = 1
				AND E.type <> 5 -- 5 = StoredData elements handled in the service
		
			OPEN stepsCursor
			FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @iAction = 
					CASE
						WHEN @iElementType = 1 THEN 1	-- Terminator
						WHEN @iElementType = 2 THEN 2	-- Web form (action required from user)
						WHEN @iElementType = 3 THEN 1	-- Email
						WHEN @iElementType = 4 THEN 1	-- Decision
						WHEN @iElementType = 6 THEN 3	-- Summing Junction
						WHEN @iElementType = 7 THEN 4	-- Or	
						WHEN @iElementType = 8 THEN 1	-- Connector 1
						WHEN @iElementType = 9 THEN 1	-- Connector 2
						ELSE 0					-- Unknown
					END
				
				IF @iAction = 3 -- Summing Junction check
				BEGIN
					/* Check if all preceding steps have completed before submitting this step. */
					SET @iInvalidDecisionCount = 0
		
					CREATE TABLE #precedingElements (elementID integer)
				
					EXEC spASRGetPrecedingWorkflowElements @iElementID, @superCursor OUTPUT
			
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
		
						/* Check that the preceding element, if it was a Decision element,
						not only completed, but also followed onto the path to the current element. */
						IF (@iInvalidDecisionCount = 0) 
						BEGIN
							SELECT @iType = WE.type,
								@iDecisionFlow = WIS.decisionFlow
							FROM ASRSysWorkflowInstanceSteps WIS
							INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
							WHERE WIS.instanceID = @iInstanceID
								AND WE.ID = @iTemp
		
							IF (@iType = 4) -- Decision
							BEGIN
								CREATE TABLE #succeedingElements2 (elementID integer)
								
								EXEC spASRGetDecisionSucceedingWorkflowElements @iTemp, @iDecisionFlow, @superCursor2 OUTPUT
								
								FETCH NEXT FROM @superCursor2 INTO @iTemp2
								WHILE (@@fetch_status = 0)
								BEGIN
									INSERT INTO #succeedingElements2 (elementID) VALUES (@iTemp2)
									
									FETCH NEXT FROM @superCursor2 INTO @iTemp2
								END
								CLOSE @superCursor2
								DEALLOCATE @superCursor2
								
								SELECT @iCount = COUNT(*)
								FROM #succeedingElements2
								WHERE elementID = @iElementID
		
								IF @iCount = 0 SET @iInvalidDecisionCount = @iInvalidDecisionCount + 1
								
								DROP TABLE #succeedingElements2
							END
						END
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
					
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowInstanceSteps WIS
					INNER JOIN #precedingElements PE ON WIS.elementID = PE.elementID
					WHERE WIS.instanceID = @iInstanceID
						AND WIS.status <> 3 -- 3 = completed
		
					/* If all preceding steps have been completed submit the Summing Junction step. *'


	SET @sSPCode_1 = '/
					IF (@iCount = 0) AND (@iInvalidDecisionCount = 0) SET @iAction = 1
		
					DROP TABLE #precedingElements
				END
		
				IF @iAction = 4 -- Or check
				BEGIN
					/* Check if any preceding steps have completed before submitting this step. */
					CREATE TABLE #precedingElements2 (elementID integer)
		
					EXEC spASRGetPrecedingWorkflowElements @iElementID, @superCursor output
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #precedingElements2 (elementID) VALUES (@iTemp)
					
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
		
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowInstanceSteps WIS
					INNER JOIN #precedingElements2 PE ON WIS.elementID = PE.elementID
					WHERE WIS.instanceID = @iInstanceID
						AND WIS.status = 3 -- 3 = completed
		
					/* If all preceding steps have been completed submit the Or step. */
					IF @iCount > 0 
					BEGIN
						/* Cancel any preceding steps that are not completed as they are no longer required. */
						EXEC spASRCancelPendingPrecedingWorkflowElements @iInstanceID, @iElementID, @iElementID
		
						SET @iAction = 1
					END
		
					DROP TABLE #precedingElements2
				END
		
				IF @iAction = 1
				BEGIN
					EXEC spASRSubmitWorkflowStep @iInstanceID, @iElementID, '''', '''', @sForms OUTPUT
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 2
					WHERE id = @iStepID
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
			END

			DECLARE timeoutCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT 
				WIS.instanceID,
				WE.ID,
				WIS.ID
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
				AND WE.type = 2 -- WebForm
			WHERE WIS.status = 2 -- Pending user action
				AND isnull(WE.timeoutFrequency,0) > 0
				AND CASE 
					WHEN WE.timeoutPeriod = 0 THEN datediff(minute, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 1 THEN datediff(Hour, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 2 THEN datediff(day, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 3 THEN datediff(week, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 4 THEN datediff(month, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 5 THEN datediff(year, WIS.activationDateTime, getDate())
					ELSE 0
				END >= WE.timeoutFrequency

			OPEN timeoutCursor
			FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			WHILE (@@fetch_status = 0)
			BEGIN
				-- Set the step status to be Timeout
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 6 -- Timeout
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID

				-- Activate the succeeding elements on the Timeout flow
				CREATE TABLE #succeedingElements3 (elementID integer)
					
				EXEC spASRGetDecisionSucceedingWorkflowElements @iElementID, 1, @superCursor OUTPUT
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements3 (elementID) VALUES (@iTemp)
									
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
					
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT #succeedingElements3.elementID 
						FROM #succeedingElements3)
					AND ASRSysWorkflowInstanceSteps.status = 0
					
				DROP TABLE #succeedingElements3

				/* Set activated Web Forms to be ''pending'


	SET @sSPCode_2 = ''' (to be done by the user) */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2)
					
				/* Set activated Terminators to be ''completed'' */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate()
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1)
					
				/* Count how many terminators have completed. ie. if the workflow has completed. */
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1
										
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @iInstanceID
					
					/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
				END

				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRWorkflowUsesInitiator
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowUsesInitiator]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowUsesInitiator]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowUsesInitiator]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRWorkflowUsesInitiator
			(
				@piWorkflowID		integer,			
				@pfUsesInitiator		bit	OUTPUT
			)
			AS
			BEGIN
				/* Return 1 if the given workflow uses the initiator''s personnel record; else return 0 */
				DECLARE
					@iCount	integer
			
				SET @pfUsesInitiator = 0
			
				/* Initiator''s record used by a Stored Data element action? */
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
					AND (ASRSysWorkflowElements.dataRecord = 0 OR ASRSysWorkflowElements.secondaryDataRecord = 0) -- 0 = Initiator''s record
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
				IF @iCount > 0 SET @pfUsesInitiator = 1
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Stored Data element Database Value item? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementColumns
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementColumns.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
						AND ASRSysWorkflowElementColumns.valueType = 2 -- 2 = Database value
						AND ASRSysWorkflowElementColumns.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END

				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by an Email element address? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
						AND ASRSysWorkflowElements.emailRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by an Email element Database Value item? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
						AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Web Form element Database Value? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
						AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Web Form element Record Selector? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
						AND ASRSysWorkflowElementItems.itemType = 11 -- 11 = Record Selector
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END

				IF @pfUsesInitiator = 0
				BEGIN
					-- Expressions
					SELECT @iCount = COUNT(*)
					FROM ASRSysExprComponents EC
					WHERE EC.exprID in (SELECT E.exprID
							FROM ASRSysExpressions E
							WHERE E.utilityid = '


	SET @sSPCode_1 = '@piWorkflowID)
						AND EC.workflowRecord = 0 -- 0 = Initiator''s record

							
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRWorkflowTriggering
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowTriggering]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowTriggering]

/* ------------------------------------------------------------- */
PRINT 'Step 5 of X - Run time Bradford Factor'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysFunctions')
	and name = 'UDFName'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysFunctions ADD 
						UDFName [nvarchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 73'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, UDFName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes)
			VALUES (73, ''Bradford Factor'', 2, 0, ''Absence'', '''',''udf_ASRFn_BradfordFactor'', 0, 1, NULL, 1, NULL)'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 73'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (73, 1, 4, ''<Start Date>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (73, 2, 4, ''<End Date>'')'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
		VALUES (73, 3, 1, ''<Absence Type(s)>'')'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 6 of X - Performance Indexing'

	-- ASRSysAccordTransferFieldDefinitions
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransferFieldDefinitions') AND name = N'IDX_TransferTypeID')
		DROP INDEX ASRSysAccordTransferFieldDefinitions.[IDX_TransferTypeID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_TransferTypeID] ON ASRSysAccordTransferFieldDefinitions ([TransferTypeID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysAccordTransferTypes
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransferTypes') AND name = N'PK_TransferTypeID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes ALTER COLUMN TransferTypeID INT NOT NULL'
		EXEC sp_executesql @NVarCommand
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysAccordTransferTypes ADD CONSTRAINT
					PK_TransferTypeID PRIMARY KEY NONCLUSTERED 
					(TransferTypeID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysBatchJobName
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysBatchJobName') AND name = N'PK_ASRSysBatchJobName_ID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysBatchJobName ADD CONSTRAINT
					PK_ASRSysBatchJobName_ID PRIMARY KEY NONCLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysBatchJobName') AND name = N'IDX_Name')
		DROP INDEX ASRSysBatchJobName.[IDX_Name]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_Name] ON ASRSysBatchJobName ([Name])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCalendarReportAccess
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCalendarReportAccess') AND name = N'IDX_ID')
		DROP INDEX ASRSysCalendarReportAccess.[IDX_ID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ID] ON ASRSysCalendarReportAccess ([ID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCalendarReportEvents
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCalendarReportEvents') AND name = N'PK_ASRSysCalendarReportEvents')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysCalendarReportEvents ADD CONSTRAINT
					PK_ASRSysCalendarReportEvents PRIMARY KEY NONCLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCalendarReportEvents') AND name = N'IDX_CalendarReportID')
		DROP INDEX ASRSysCalendarReportEvents.[IDX_CalendarReportID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_CalendarReportID] ON ASRSysCalendarReportEvents ([CalendarReportID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCalendarReports
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCalendarReports') AND name = N'PK_ASRSysCalendarReports')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysCalendarReports ADD CONSTRAINT
					PK_ASRSysCalendarReports PRIMARY KEY CLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCalendarReports') AND name = N'IDX_BaseTableID')
		DROP INDEX ASRSysCalendarReports.[IDX_BaseTableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_BaseTableID] ON ASRSysCalendarReports ([BaseTable])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysChildViewParents2
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysChildViewParents2') AND name = N'IDX_ChildViewID')
		DROP INDEX ASRSysChildViewParents2.[IDX_ChildViewID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ChildViewID] ON ASRSysChildViewParents2 ([ChildViewID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysChildViews2
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysChildViews2') AND name = N'PK_ASRSysChildViews2')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysChildViews2 ADD CONSTRAINT
					PK_ASRSysChildViews2 PRIMARY KEY CLUSTERED 
					(ChildViewID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysChildViews2') AND name = N'IDX_TableID')
		DROP INDEX ASRSysChildViews2.[IDX_TableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TableID] ON ASRSysChildViews2 ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysColours
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysColours') AND name = N'IDX_ColOrder')
		DROP INDEX ASRSysColours.[IDX_ColOrder]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ColOrder] ON ASRSysColours ([ColOrder])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysColumnControlValues
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysColumnControlValues') AND name = N'IDX_ColumnSequence')
		DROP INDEX ASRSysColumnControlValues.[IDX_ColumnSequence]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ColumnSequence] ON ASRSysColumnControlValues ([ColumnID], [Sequence])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysColumns
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysColumns') AND name = N'FK_TableID')
		DROP INDEX ASRSysColumns.[FK_TableID]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysColumns') AND name = N'IDX_TableID')
		DROP INDEX ASRSysColumns.[IDX_TableID]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysColumns') AND name = N'PK_ASRSysColumns')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysColumns DROP CONSTRAINT PK_ASRSysColumns'
		EXEC sp_executesql @NVarCommand
	END

	SET @NVarCommand = 'ALTER TABLE dbo.ASRSysColumns ADD CONSTRAINT
				PK_ASRSysColumns PRIMARY KEY CLUSTERED
				(ColumnID) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand

	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TableID] ON ASRSysColumns ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysControls
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysControls') AND name = N'FK_ScreenID')
		DROP INDEX ASRSysControls.[FK_ScreenID]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysControls') AND name = N'IDX_ScreenID')
		DROP INDEX ASRSysControls.[IDX_ScreenID]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysControls') AND name = N'IDX_TableID')
		DROP INDEX ASRSysControls.[IDX_TableID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ScreenID] ON ASRSysControls ([ScreenID])'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TableID] ON ASRSysControls ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCustomReportAccess
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCustomReportAccess') AND name = N'IDX_ID')
		DROP INDEX ASRSysCustomReportAccess.[IDX_ID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ID] ON ASRSysCustomReportAccess ([ID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCustomReportsChildDetails
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCustomReportsChildDetails') AND name = N'IDX_CustomReportID')
		DROP INDEX ASRSysCustomReportsChildDetails.[IDX_CustomReportID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_CustomReportID] ON ASRSysCustomReportsChildDetails ([CustomReportID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCustomReportsDetails
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCustomReportsDetails') AND name = N'IDX_CustomReportID')
		DROP INDEX ASRSysCustomReportsDetails.[IDX_CustomReportID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_CustomReportID] ON ASRSysCustomReportsDetails ([CustomReportID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysCustomReportsName
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysCustomReportsName') AND name = N'PK_ASRSysCustomReportsName_ID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysCustomReportsName ADD CONSTRAINT
					PK_ASRSysCustomReportsName_ID PRIMARY KEY CLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysDataTransferName
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysDataTransferName') AND name = N'IDX_FromTableID')
		DROP INDEX ASRSysDataTransferName.[IDX_FromTableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_FromTableID] ON ASRSysDataTransferName ([FromTableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysDiaryLinks
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysDiaryLinks') AND name = N'IDX_ColumnID')
		DROP INDEX ASRSysDiaryLinks.[IDX_ColumnID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ColumnID] ON ASRSysDiaryLinks ([ColumnID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysDiaryEvents
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysDiaryEvents') AND name = N'IDX_DateEventID')
		DROP INDEX ASRSysDiaryEvents.[IDX_DateEventID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_DateEventID] ON ASRSysDiaryEvents ([EventDate],[DiaryEventsID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysEmailAddresses
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysEmailAddress') AND name = N'IDX_TableID')
		DROP INDEX ASRSysEmailAddress.[IDX_TableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TableID] ON ASRSysEmailAddress ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysEmailLinks
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysEmailLinks') AND name = N'PK_ASREmailLinks')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysEmailLinks ADD CONSTRAINT
					PK_ASREmailLinks PRIMARY KEY CLUSTERED
					(LinkID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysEmailQueue
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysEmailQueue') AND name = N'PK_ASRSysEmailQueue')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysEmailQueue ADD CONSTRAINT
					PK_ASRSysEmailQueue PRIMARY KEY CLUSTERED
					(QueueID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysEventLog
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysEventLog') AND name = N'PK_ASRSysEventLog_ID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysEventLog ADD CONSTRAINT
					PK_ASRSysEventLog_ID PRIMARY KEY CLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysEventLogDetails
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysEventLogDetails') AND name = N'IDX_EventLogID')
		DROP INDEX ASRSysEventLogDetails.[IDX_EventLogID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_EventLogID] ON ASRSysEventLogDetails ([EventLogID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysExprComponents
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysExprComponents') AND name = N'Component_ID')
		DROP INDEX ASRSysExprComponents.[Component_ID]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysExprComponents') AND name = N'IDX_ExprID')
		DROP INDEX ASRSysExprComponents.[IDX_ExprID]
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysExprComponents') AND name = N'PK_ComponentID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysExprComponents ADD CONSTRAINT
						PK_ComponentID PRIMARY KEY NONCLUSTERED 
						(ComponentID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ExprID] ON ASRSysExprComponents ([ExprID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysExpressions
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysExpressions') AND name = N'IDX_ParentComponentID')
		DROP INDEX ASRSysExpressions.[IDX_ParentComponentID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ParentComponentID] ON ASRSysExpressions ([ParentComponentID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysFunctionParameters
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysFunctionParameters') AND name = N'IDX_FunctionID')
		DROP INDEX ASRSysFunctionParameters.[IDX_FunctionID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_FunctionID] ON ASRSysFunctionParameters ([FunctionID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysGroupPremissions
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysGroupPermissions') AND name = N'IDX_GroupName')
		DROP INDEX ASRSysGroupPermissions.[IDX_GroupName]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysGroupPermissions') AND name = N'IDX_ItemID')
		DROP INDEX ASRSysGroupPermissions.[IDX_ItemID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_GroupName] ON ASRSysGroupPermissions ([GroupName])'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ItemID] ON ASRSysGroupPermissions ([ItemID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysImportName
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysImportName') AND name = N'PK_ASRSysImportName')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysImportName ADD CONSTRAINT
					PK_ASRSysImportName PRIMARY KEY CLUSTERED
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysLock
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysLock') AND name = N'IDX_Priority')
		DROP INDEX ASRSysLock.[IDX_Priority]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_Priority] ON ASRSysLock ([Priority])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysMailMergeColumns
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysMailMergeColumns') AND name = N'PK_ASRSysMailMergeColumns')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysMailMergeColumns ADD CONSTRAINT
					PK_ASRSysMailMergeColumns PRIMARY KEY NONCLUSTERED
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysMailMergeColumns') AND name = N'IDX_MailMergeID')
		DROP INDEX ASRSysMailMergeColumns.[IDX_MailMergeID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_MailMergeID] ON ASRSysMailMergeColumns ([MailMergeID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysMailMergeName
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysMailMergeName') AND name = N'IDX_TableID')
		DROP INDEX ASRSysMailMergeName.[IDX_TableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TableID] ON ASRSysMailMergeName ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysMatchReportName
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysMatchReportName') AND name = N'IDX_TypeTableID')
		DROP INDEX ASRSysMatchReportName.[IDX_TypeTableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TypeTableID] ON ASRSysMatchReportName ([MatchReportType], [Table1ID], [Table2ID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysModuleSetup
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysModuleSetup') AND name = N'IDX_ModuleParameterKey')
		DROP INDEX ASRSysModuleSetup.[IDX_ModuleParameterKey]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ModuleParameterKey] ON ASRSysModuleSetup ([ModuleKey],[ParameterKey])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysOperatorParameters
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysOperatorParameters') AND name = N'IDX_OperatorID')
		DROP INDEX ASRSysOperatorParameters.[IDX_OperatorID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_OperatorID] ON ASRSysOperatorParameters ([OperatorID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysOperators
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysOperators') AND name = N'PK_OperatorID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysOperators ADD CONSTRAINT
					PK_OperatorID PRIMARY KEY CLUSTERED 
					(OperatorID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	-- ASRSysOrderItems
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysOrderItems') AND name = N'IDX_OrderSequence')
		DROP INDEX ASRSysOrderItems.[IDX_OrderSequence]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysOrderItems') AND name = N'IDX_ColumnID')
		DROP INDEX ASRSysOrderItems.[IDX_ColumnID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_OrderSequence] ON ASRSysOrderItems ([OrderID],[Sequence])'
	EXEC sp_executesql @NVarCommand
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ColumnID] ON ASRSysOrderItems ([ColumnID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysPageCaptions
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPageCaptions') AND name = N'IDX_ScreenID')
		DROP INDEX ASRSysPageCaptions.[IDX_ScreenID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ScreenID] ON ASRSysPageCaptions ([ScreenID],[PageIndexID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysPageSizes
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPageSizes') AND name = N'PK_PageSizeID')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysPageSizes ADD CONSTRAINT
					PK_PageSizeID PRIMARY KEY CLUSTERED 
					(PageSizeID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysPermissionCategories
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPermissionCategories') AND name = N'IDX_CategoryKeyID')
		DROP INDEX ASRSysPermissionCategories.[IDX_CategoryKeyID]
	SET @NVarCommand = 'CREATE UNIQUE NONCLUSTERED INDEX [IDX_CategoryKeyID] ON ASRSysPermissionCategories ([CategoryID], [CategoryKey])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysPermissionItems
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPermissionItems') AND name = N'IDX_ItemKey')
		DROP INDEX ASRSysPermissionItems.[IDX_ItemKey]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ItemKey] ON ASRSysPermissionItems ([ItemKey])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysPictures
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPictures') AND name = N'PK_ASRSysPictures')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysPictures DROP CONSTRAINT PK_ASRSysPictures'
		EXEC sp_executesql @NVarCommand
	END

	SET @NVarCommand = 'ALTER TABLE dbo.ASRSysPictures ADD CONSTRAINT
				PK_ASRSysPictures PRIMARY KEY CLUSTERED
				(PictureID) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand


	-- ASRSysRecordProfileName
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysRecordProfileName') AND name = N'IDX_BaseTableID')
		DROP INDEX ASRSysRecordProfileName.[IDX_BaseTableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_BaseTableID] ON ASRSysRecordProfileName ([BaseTable])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysRelations
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysRelations') AND name = N'IDX_ParentChildID')
		DROP INDEX ASRSysRelations.[IDX_ParentChildID]
	SET @NVarCommand = 'CREATE UNIQUE CLUSTERED INDEX [IDX_ParentChildID] ON ASRSysRelations ([ParentID],[ChildID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysScreens
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysScreens') AND name = N'PK_ASRSysScreens')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysScreens DROP CONSTRAINT PK_ASRSysScreens'
		EXEC sp_executesql @NVarCommand
	END
	SET @NVarCommand = 'ALTER TABLE dbo.ASRSysScreens ADD CONSTRAINT
				PK_ASRSysScreens PRIMARY KEY NONCLUSTERED (ScreenID) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysScreens') AND name = N'IDX_TableID')
		DROP INDEX ASRSysScreens.[IDX_TableID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_TableID] ON ASRSysScreens ([TableID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysSSIntranetLinks
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSSIntranetLinks') AND name = N'PK_ASRSysSSIntranetLinks')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysSSIntranetLinks ADD CONSTRAINT
					PK_ASRSysSSIntranetLinks PRIMARY KEY CLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSSIntranetLinks') AND name = N'IDX_ViewID')
		DROP INDEX ASRSysSSIntranetLinks.[IDX_ViewID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ViewID] ON ASRSysSSIntranetLinks ([ViewID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysSummaryFields
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSummaryFields') AND name = N'IDX_HistoryTableSequenceID')
		DROP INDEX ASRSysSummaryFields.[IDX_HistoryTableSequenceID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_HistoryTableSequenceID] ON ASRSysSummaryFields ([HistoryTableID],[Sequence])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysTables
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysTables') AND name = N'PK_ASRSysTables')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysTables DROP CONSTRAINT PK_ASRSysTables'
		EXEC sp_executesql @NVarCommand
	END
	SET @NVarCommand = 'ALTER TABLE dbo.ASRSysTables ADD CONSTRAINT
				PK_ASRSysTables PRIMARY KEY CLUSTERED (TableID) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand


	-- ASRSysUniqueCodes
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysUniqueCodes') AND name = N'PK_ASRSysUniqueCodes')
	BEGIN
		SET @NVarCommand = 'DELETE FROM ASRSysUniqueCodes WHERE CodePrefix IS NULL'
		EXEC sp_executesql @NVarCommand
		SET @NVarCommand = 'ALTER TABLE ASRSysUniqueCodes ALTER COLUMN CodePrefix varchar(20) NOT NULL'
		EXEC sp_executesql @NVarCommand
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysUniqueCodes ADD CONSTRAINT
					PK_ASRSysUniqueCodes PRIMARY KEY CLUSTERED
					(CodePrefix) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysUserSettings
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysUserSettings') AND name = N'IDX_SettingID')
		DROP INDEX ASRSysUserSettings.[IDX_SettingID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_SettingID] ON ASRSysUserSettings ([Section],[SettingKey],[UserName])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysUtilAccessLog
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysUtilAccessLog') AND name = N'IDX_UtilIDType')
		DROP INDEX ASRSysUtilAccessLog.[IDX_UtilIDType]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_UtilIDType] ON ASRSysUtilAccessLog ([UtilID],[Type])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysViewColumns
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysViewColumns') AND name = N'IDX_ViewID')
		DROP INDEX ASRSysViewColumns.[IDX_ViewID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_ViewID] ON ASRSysViewColumns ([ViewID],[ColumnID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysViewMenuPermissions
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysViewMenuPermissions') AND name = N'IDX_GroupName')
		DROP INDEX ASRSysViewMenuPermissions.[IDX_GroupName]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_GroupName] ON ASRSysViewMenuPermissions ([GroupName])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysViews
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysViews') AND name = N'PK_ASRSysViews')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysViews ADD CONSTRAINT
					PK_ASRSysViews PRIMARY KEY CLUSTERED 
					(ViewID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysViewScreens
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysViewScreens') AND name = N'IDX_ScreenViewID')
		DROP INDEX ASRSysViewScreens.[IDX_ScreenViewID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_ScreenViewID] ON ASRSysViewScreens ([ScreenID],[ViewID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysWorkflowElements
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysWorkflowElements') AND name = N'PK_ASRSysWorkflowElements')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysWorkflowElements ADD CONSTRAINT
					PK_ASRSysWorkflowElements PRIMARY KEY NONCLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysWorkflowInstanceSteps
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysWorkflowInstanceSteps') AND name = N'IDX_InstanceID')
		DROP INDEX ASRSysWorkflowInstanceSteps.[IDX_InstanceID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_InstanceID] ON ASRSysWorkflowInstanceSteps ([InstanceID])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysWorkflowQueue
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysWorkflowQueue') AND name = N'PK_ASRSysWorkflowQueue')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysWorkflowQueue ADD CONSTRAINT
					PK_ASRSysWorkflowQueue PRIMARY KEY CLUSTERED 
					(QueueID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	-- ASRSysWorkflows
	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysWorkflows') AND name = N'PK_ASRSysWorkflows')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysWorkflows ADD CONSTRAINT
					PK_ASRSysWorkflows PRIMARY KEY CLUSTERED 
					(ID) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END



/* ------------------------------------------------------------- */
PRINT 'Step 7 of X - Update Email Procedure'

	----------------------------------------------------------------------
	-- spASREmailImmediate
	----------------------------------------------------------------------

    --Delete any entries which future due date which have already been sent
	--(These will be recreate when the queue is rebuild overnight)
	SET @sSPCode_0 = 
		'DELETE
         FROM   ASRSysEmailQueue
		 WHERE  Not DateSent is Null
		   AND  datediff(dd,DateDue,getdate()) < 0'
	EXECUTE (@sSPCode_0)


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASREmailImmediate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASREmailImmediate]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASREmailImmediate]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASREmailImmediate(@Username varchar(255)) AS
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
						@DateDue datetime,
						@hResult int,
						@blnEnabled int,
						@RecalculateRecordDesc bit,
						@TableID int
					DECLARE @RecipTo varchar(4000),
						@TempText nvarchar(4000),
						@CC varchar(4000),
						@BCC varchar(4000),
						@Subject varchar(4000),
						@MsgText varchar(8000),
						@Attachment varchar(4000)
		
					DECLARE emailqueue_cursor
					CURSOR LOCAL FAST_FORWARD FOR 
					  SELECT QueueID
					       , ASRSysEmailQueue.LinkID
					       , RecordID
					       , ASRSysEmailQueue.ColumnID
					       , ColumnValue
					       , RecordDesc
					       , RecalculateRecordDesc
					       , TableID
					       , DateDue
					  FROM   ASRSysEmailQueue
					  LEFT OUTER JOIN
					         ASRSysEmailLinks
					      ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
					  WHERE  DateSent IS Null
					    AND  datediff(dd,DateDue,getdate()) >= 0
					    AND  (LOWER(substring(@Username,charindex(''\'',@Username)+1,999)) = LOWER(substring([Username],charindex(''\'',[Username])+1,999))
							  OR @Username = ''''
							  )
					  ORDER BY dateDue

					OPEN emailqueue_cursor
					FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue
		
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
										EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, @Username, @RecipTo OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
									END
							END
						ELSE IF @TableID > 0
							BEGIN
								SET @sSQL = ''spASRSysEmailAddr''
								IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
									BEGIN
										SELECT @emailDate = getDate()
										EXEC @hResult = @sSQL @RecipTo OUTPUT, @LinkID, 0
										SET @Subject = @columnvalue
										SET @MsgText = @RecDesc
										EXEC spASRSendMail @hResult OUTPUT, @RecipTo, '''', '''', @Subject,  @MsgText, ''''
									END
							EN'


		SET @sSPCode_1 = 'D
		
						IF @ColumnID IS null AND @TableID IS null
						BEGIN
							SELECT @emailDate = getDate()
		
							SELECT @RecipTo = RepTo,
								@CC = RepCC,
								@BCC = RepBCC,
								@Subject = Subject,
								@Attachment = Attachment,
								@MsgText = MsgText
							FROM ASRSysEmailQueue 
							WHERE QueueID = @QueueID
		
							IF RTrim(@RecipTo) = ''''
								SET @hResult = 1
							ELSE
								EXEC spASRSendMail @hResult OUTPUT, @RecipTo, '''', '''', @Subject,  @MsgText, ''''
						END
		
						IF @hResult = 0
						BEGIN
							UPDATE ASRSysEmailQueue SET DateSent = @emailDate, RepTo = @RecipTo, RepCC = @CC, RepBCC = @BCC, Subject = @Subject, Attachment = @Attachment
							WHERE QueueID = @QueueID
							
							UPDATE ASRSysEmailQueue SET MsgText = @MsgText
							WHERE QueueID = @QueueID
						END
						FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue
					END
					CLOSE emailqueue_cursor
					DEALLOCATE emailqueue_cursor
				END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

/* ------------------------------------------------------------- */
PRINT 'Step 8 of X - Accord changes'

	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsCompanyCode = 0 WHERE TransferFieldID = 2 AND TransferTypeID = 71'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 9 of X - Sysprocess amendments'

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersFromMaster
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersFromMaster]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersFromMaster]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersFromMaster]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetCurrentUsersFromMaster
		    AS
		    BEGIN
		
				SET NOCOUNT ON
		
				IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				BEGIN
					EXEC sp_ASRIntCheckPolls
				END
		
				SELECT p.hostname, p.loginame, p.program_name, p.hostprocess
					   , p.sid, p.login_time, p.spid
				FROM     master..sysprocesses p
				JOIN     master..sysdatabases d ON     d.dbid = p.dbid
				WHERE    p.program_name LIKE ''HR Pro%''
				  AND    p.program_name NOT LIKE ''HR Pro Workflow%''
				  AND    d.name = db_name()
				ORDER BY loginame
		
				SET NOCOUNT OFF
		
		    END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRInsertToTableFromText
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInsertToTableFromText]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInsertToTableFromText]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRInsertToTableFromText]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRInsertToTableFromText
			@sInputValues VARCHAR(8000),  
		    @sInsertTable SYSNAME AS  
		BEGIN  
			DECLARE @sRecord varchar(8000)
			DECLARE @sValues varchar(8000)
			DECLARE @sValue varchar(8000)
		    DECLARE @iRecordStart smallint
			DECLARE @iValueStart smallint
			DECLARE @sInsertSQL varchar(8000)  
		    
			SET @sRecord = ''''
		
		    WHILE @sInputValues <> ''''
		    BEGIN  
		
		        SET @iRecordStart = CHARINDEX(''##'', @sInputValues)  
		        IF @iRecordStart > 0  
		            BEGIN  
		                SET @sRecord = LEFT(@sInputValues, @iRecordStart-1) + '',''
		                SET @sInputValues = RIGHT(@sInputValues, LEN(@sInputValues) - (@iRecordStart + 1))  
						SET @sValues = ''''
		
						WHILE @sRecord <> ''''
						BEGIN
							SET @iValueStart = CHARINDEX('','', @sRecord)  
							IF @iValueStart > 0
								BEGIN
									SET @sValue = LEFT(@sRecord, @iValueStart-1)
									SET @sRecord = RIGHT(@sRecord, LEN(@sRecord) - @iValueStart)  
								END
							ELSE
								BEGIN
									SET @sRecord = ''''  
								END
		
							IF LEFT(@sValue,2) = ''0x''
							BEGIN
								IF LEN(@sValues) = 0 SET @sValues = '''' + @sValue + ''''
								ELSE SET @sValues = @sValues + '','' + @sValue + ''''
							END
							ELSE
							BEGIN
								IF LEN(@sValues) = 0 SET @sValues = '''''''' + @sValue + ''''''''
								ELSE SET @sValues = @sValues + '','''''' + @sValue + ''''''''
							END
						END
		
					    SET @sInsertSQL = ''INSERT INTO '' + @sInsertTable + '' VALUES('' + @sValues + '')''  
						EXECUTE (@sInsertSQL)
		
		            END  
		        ELSE
					BEGIN  
						SET @sRecord = @sInputValues 
						SET @sInputValues = ''''  
					END
		
		    END  
		
		END'

	EXECUTE (@sSPCode_0)


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsers]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsers]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsers]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRGetCurrentUsers
		    AS
		    BEGIN
		
				SET NOCOUNT ON
		
				DECLARE @Login nvarchar(200)
				DECLARE @hResult int
				DECLARE @objectToken int
				DECLARE @sTableName nvarchar(50)
				DECLARE @sSQL nvarchar(500)
				DECLARE @sCurrentUsers nvarchar(4000)
				DECLARE @sSQLVersion char(2)
		
				SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
				IF @sSQLVersion = ''9''
				BEGIN
		
					SELECT @Login = [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = ''MODULE_SQL''
						AND [ParameterKey] = ''Param_FieldsLoginDetails''
		
					EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
					IF @hResult = 0
					BEGIN
		
						CREATE TABLE #tmpProcesses(HostName varchar(100), Loginame varchar(100), Program_Name varchar(100), HostProcess int, Sid binary(86), Login_Time datetime, spid int)
						EXEC @hResult = sp_OAMethod @objectToken, ''GetCurrentUsersArray'', @sCurrentUsers OUTPUT, @Login
						EXEC dbo.[spASRInsertToTableFromText] @sCurrentUsers, ''#tmpProcesses''
						SELECT * FROM #tmpProcesses
		
					END
		
					EXEC @hResult = sp_OADestroy @objectToken
					DROP TABLE #tmpProcesses
		
				END
				ELSE
				BEGIN
					EXECUTE dbo.[spASRGetCurrentUsersFromMaster]
				END
		
				SET NOCOUNT OFF
		
		    END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- sp_ASRLockCheck
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRLockCheck]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRLockCheck]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRLockCheck]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE sp_ASRLockCheck AS
	BEGIN

		SET NOCOUNT ON

		DECLARE @sSQLVersion char(2)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9'' AND APP_NAME() <> ''HR Pro Workflow Service''
		BEGIN

			CREATE TABLE #tmpProcesses(HostName varchar(100), LoginName varchar(100), Program_Name varchar(100), HostProcess int, Sid binary(86), Login_Time datetime, spid int)
			INSERT #tmpProcesses EXEC dbo.[spASRGetCurrentUsers]

			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN #tmpProcesses syspro 
				ON ASRSysLock.spid = syspro.spid AND ASRSysLock.login_time = syspro.login_time
			WHERE priority = 2 OR syspro.spid IS NOT NULL
			ORDER BY priority

			DROP TABLE #tmpProcesses

		END
		ELSE
		BEGIN

			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN master..sysprocesses syspro 
				ON asrsyslock.spid = syspro.spid AND asrsyslock.login_time = syspro.login_time
			WHERE Priority = 2 OR syspro.spid IS NOT NULL
			ORDER BY Priority

		END

		SET NOCOUNT OFF

	END'

	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRGetCurrentUsersCountOnServer
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountOnServer]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRGetCurrentUsersCountOnServer
					(
						@iLoginCount	integer OUTPUT,
						@psLoginName	varchar(8000)
					)
		    AS
		    BEGIN
		
				DECLARE @Login nvarchar(200)
				DECLARE @sSQLVersion char(2)
				DECLARE @hResult int
				DECLARE @objectToken int
		
				IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				BEGIN
					EXEC sp_ASRIntCheckPolls
				END
		
				SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
				IF @sSQLVersion = ''9''
				BEGIN
		
					SELECT @Login = [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = ''MODULE_SQL''
						AND [ParameterKey] = ''Param_FieldsLoginDetails''
		
					EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
					IF @hResult = 0
					BEGIN
						EXEC @hResult = sp_OAMethod @objectToken, ''CountCurrentLogins'', @iLoginCount OUTPUT, @Login, @psLoginName
					END
		
					EXEC @hResult = sp_OADestroy @objectToken
		
				END
				ELSE
				BEGIN
		
					SELECT   @iLoginCount = COUNT(*)
					FROM     master..sysprocesses p
					WHERE    p.program_name LIKE ''HR Pro%''
					  AND    p.program_name NOT LIKE ''HR Pro Workflow%''
					  AND    p.loginame = @psLoginName
				END
		
		    END'

	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRDropTempObjects
	----------------------------------------------------------------------
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDropTempObjects]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDropTempObjects]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRDropTempObjects]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE spASRDropTempObjects
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
						OR OBJECTPROPERTY(id, N''IsInlineFunction'') = 1
						OR OBJECTPROPERTY(id, N''IsScalarFunction'') = 1
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

				IF UPPER(@sXType) = ''FN''
					-- UDF
					BEGIN
						EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END
				
				FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
				
			END
			CLOSE tempObjects
			DEALLOCATE tempObjects
			
			EXEC (''DELETE FROM [dbo].[ASRSysSQLObjects]'')


			-- Clear out any temporary tables that may have got left behind from the createunique function
			DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
			SELECT [dbo].[sysobjects].[name]
			FROM [dbo].[sysobjects] 
			INNER JOIN [dbo].[sysusers]	ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
			LEFT JOIN ASRSysTables ON sysobjects.[name] = ASRSysTables.TableName
			WHERE LOWER([dbo].[sysusers].[name]) = ''dbo''
				AND OBJECTPROPERTY(sysobjects.id, N''IsUserTable'') = 1
				AND ASRSysTables.TableName IS NULL
				AND [dbo].[sysobjects].[name] LIKE ''tmp%''

			OPEN tempObjects
			FETCH NEXT FROM tempObjects INTO @sObjectName
			WHILE (@@fetch_status <> -1)
			BEGIN		
				EXEC (''DROP TABLE [dbo].['' + @sObjectName + '']'')
				FETCH NEXT FROM tempObjects INTO @sObjectName
			END

			CLOSE tempObjects
			DEALLOCATE tempObjects

		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 10 of X - Modifying Expression Function stored procedures'

	----------------------------------------------------------------------
	-- sp_ASRFn_GetCurrentUser
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_GetCurrentUser]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE sp_ASRFn_GetCurrentUser 
		(
			@psResult	varchar(255) OUTPUT
		)
		AS
		BEGIN
			SET @psResult = 
				CASE 
					WHEN UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW'' THEN ''HR Pro Workflow'' 
					ELSE SUSER_SNAME()
				END
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 11 of X - Modifying ASRSysSSIntranetLinks Table'


	/* ASRSysSSIntranetLinks - Add newWindow column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
	and name = 'newWindow'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
						newWindow[bit] NULL'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 12 of X - Modifying Settings stored procedures'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSaveUserSetting]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSaveUserSetting]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSaveUserSetting]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRSaveUserSetting]
	(
		@sSection		varchar(50),
		@sSettingKey	varchar(50),
		@sUsername		varchar(50),
		@sSettingValue	varchar(255)
	)
	AS
	BEGIN
		SET NOCOUNT ON

		IF EXISTS(SELECT [SettingValue] FROM ASRSysUserSettings WHERE Section = @sSection	 AND SettingKey = @sSettingKey AND UserName = @sUsername)
			UPDATE ASRSysUserSettings SET [SettingValue] = @sSettingValue WHERE Section = @sSection AND SettingKey = @sSettingKey AND UserName = @sUsername
		ELSE
			INSERT ASRSysUserSettings ([Section], [SettingKey], [UserName], [SettingValue]) VALUES (@sSection, @sSettingKey, @sUsername, @sSettingValue)

	END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 13 of X - Modifying Calendar Reports table'

	/* ASRSysCalendarReports - Add StartOnCurrentMonth column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysCalendarReports')
	and name = 'StartOnCurrentMonth'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysCalendarReports ADD 
						StartOnCurrentMonth[bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysCalendarReports SET 
									StartOnCurrentMonth = 0
								WHERE StartOnCurrentMonth IS NULL'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 14 of x - Updating Maternity Entitlement'

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRMaternityExpectedReturn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRMaternityExpectedReturn]

  EXEC('CREATE PROCEDURE [dbo].[spASRMaternityExpectedReturn] (
			@pdblResult datetime OUTPUT,
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(8000)
			)
			AS
			BEGIN

				IF LOWER(@Ordinary) = ''ordinary''
					IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,26,@LeaveStart)
					ELSE
						IF DateDiff(d,''04/30/2000'', @EWCDate) >= 0
							SET @pdblResult = Dateadd(ww,18,@LeaveStart)
						ELSE
							SET @pdblResult = Dateadd(ww,14,@LeaveStart)
				ELSE
					IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,52,@LeaveStart)
					ELSE
						--29 weeks from baby birth date (but return on the monday before!)
						SET @pdblResult = DateAdd(d,203 - datepart(dw,DateAdd(d,-2,@BabyBirthDate)),@BabyBirthDate)

			END')

  GRANT EXEC ON [spASRMaternityExpectedReturn] TO [ASRSysGroup]

/* ------------------------------------------------------------- */
PRINT 'Step 15 of x - Modifying Workflow triggers'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowTriggeredLinks]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	drop trigger [dbo].[DEL_ASRSysWorkflowTriggeredLinks]

	SELECT @NVarCommand = '	CREATE TRIGGER [dbo].[DEL_ASRSysWorkflowTriggeredLinks] 
				   ON  [dbo].[ASRSysWorkflowTriggeredLinks] 
				   FOR DELETE
				AS 
				BEGIN
					DELETE FROM ASRSysWorkflowTriggeredLinkColumns WHERE LinkID IN (SELECT LinkID FROM deleted)
					DELETE FROM ASRSysWorkflowQueue WHERE LinkID IN (SELECT LinkID FROM deleted)
				END'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 16 of x - Modifying Send Messages'

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetMessages]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[sp_ASRGetMessages]

  EXEC('CREATE PROCEDURE sp_ASRGetMessages AS
	BEGIN
		DECLARE @iDBID	integer,
			@iID		integer,
			@dtLoginTime	datetime,
			@sLoginName	varchar(256),
			@iCount	integer,
			@Realspid integer

		-- Need to get spid of parent process
		SELECT @Realspid = a.spid
		FROM master..sysprocesses a
		FULL OUTER JOIN master..sysprocesses b
			ON a.hostname = b.hostname
			AND a.hostprocess = b.hostprocess
			AND a.spid <> b.spid
		WHERE b.spid = @@Spid

		-- If there is no parent spid then use current spid
		IF @Realspid is null SET @Realspid = @@spid

		-- Get the current user process information.
		SELECT @iDBID = dbID,
			@dtLoginTime = login_time,
			@sLoginName = loginame
		FROM master..sysprocesses
		WHERE spid = @Realspid

		-- Return the recordset of messages.
		SELECT ''Message from user '''''' + ltrim(rtrim(messageFrom)) + 
			'''''' using '' + ltrim(rtrim(messageSource)) + 
			'' ('' + convert(varchar(100), messageTime, 100) +'')'' + 
			char(10) + message
		FROM ASRSysMessages
		WHERE loginName = @sLoginName
			AND dbID = @iDBID
			AND loginTime = @dtLoginTime
			AND spid = @Realspid

		-- Remove any messages that have just been picked up.
		DELETE
		FROM ASRSysMessages
		WHERE loginName = @sLoginName
			AND dbID = @iDBID
			AND loginTime = @dtLoginTime
			AND spid = @Realspid

	END')

  GRANT EXEC ON [sp_ASRGetMessages] TO [ASRSysGroup]


/* ------------------------------------------------------------- */
PRINT 'Step 17 of x - Modifying Statutory Redundancy Pay Function'

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRFn_StatutoryRedundancyPay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[sp_ASRFn_StatutoryRedundancyPay]

  EXEC('CREATE PROCEDURE sp_ASRFn_StatutoryRedundancyPay 
	(
		@pdblRedundancyPay	float OUTPUT,
		@pdtStartDate 		datetime,
		@pdtLeaveDate 	datetime,
		@pdtDOB		datetime,
		@pdblWeeklyRate 	float,
		@pdblStatLimit 		float
	)
	AS
	BEGIN
		DECLARE @dtMinAgeBirthday	datetime,
			@dtServiceFrom		datetime,
			@iServiceYears 		integer,
			@iAgeY			integer,
			@iAgeM 		integer,
			@dblRate1 		float,
			@dblRate2 		float,
			@dblRate3 		float,
			@dtTempDate 		datetime,
			@iTempAgeY		integer,
			@iTemp		integer,
			@dblTemp2 		float,
			@iAfterOct2006	bit,
			@iMinAge	integer
	
		SET @pdblRedundancyPay = 0
		SET @iAfterOct2006 = case when datediff(dd,@pdtLeaveDate,''10/01/2006'') <= 0 then 1 else 0 end
	
		if @iAfterOct2006 = 1
			SET @iMinAge = 15
		else
			SET @iMinAge = 18
	
		/* First three parameters are compulsory, so return 0 and exit if they are not set */
		IF (@pdtStartDate IS null) OR (@pdtLeaveDate IS null) OR (@pdtDOB IS null)
		BEGIN
			RETURN
		END
	
		SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
		SET @pdtLeaveDate = convert(datetime, convert(varchar(20), @pdtLeaveDate, 101))
		SET @pdtDOB = convert(datetime, convert(varchar(20), @pdtDOB, 101))


		/* Calc start date */
	   	SET @dtServiceFrom = @pdtStartDate
		if @iAfterOct2006 = 0
		BEGIN
			SET @dtMinAgeBirthday = dateadd(yy, @iMinAge, @pdtDOB)
			IF @dtMinAgeBirthday >= @pdtStartDate
				SET @dtServiceFrom = @dtMinAgeBirthday
		END

	
		/* Calc number of applicable complete yrs the employee has been employed */
		exec sp_ASRFn_WholeYearsBetweenTwoDates @iServiceYears OUTPUT, @dtServiceFrom, @pdtLeaveDate
	
		/* exit if its less than 2 years */
		IF @iServiceYears < 2 
		BEGIN
			RETURN
		END
	
		/* calculate the employees years and months to the leave date */
		exec sp_ASRFn_WholeYearsBetweenTwoDates @iAgeY OUTPUT, @pdtDOB, @pdtLeaveDate
	
		SET @dtTempDate = dateadd(yy, @iAgeY, @pdtDOB)
		exec sp_ASRFn_WholeMonthsBetweenTwoDates @iAgeM OUTPUT, @dtTempDate, @pdtLeaveDate
	
		/* only count up to 20 years for redundancy */
		exec sp_ASRFn_Minimum @iServiceYears OUTPUT, 20, @iServiceYears
	
		/* fill in the rates depending on service and age */
		SET @iTempAgeY = @iAgeY
		SET @dblRate1 = 0
		SET @dblRate2 = 0
		SET @dblRate3 = 0
	
		IF @iTempAgeY >= 41
		BEGIN
			SET @iTemp = @iTempAgeY - 41
			exec sp_ASRFn_Minimum @dblRate1 OUTPUT, @iTemp, @iServiceYears
			SET @iTempAgeY = 41
			SET @iServiceYears = @iServiceYears - @dblRate1
		END
	
		IF @iTempAgeY >= 22
		BEGIN
			SET @iTemp = @iTempAgeY - 22
			exec sp_ASRFn_Minimum @dblRate2 OUTPUT, @iTemp, @iServiceYears
			SET @iTempAgeY = 22
			SET @iServiceYears = @iServiceYears - @dblRate2
		END
	
		IF @iTempAgeY >= @iMinAge
		BEGIN
			SET @iTemp = @iTempAgeY - @iMinAge
			exec sp_ASRFn_Minimum @dblRate3 OUTPUT, @iTemp, @iServiceYears
		END
	
		/* calc the redundancy pay */
		exec sp_ASRFn_Minimum @dblTemp2 OUTPUT, @pdblWeeklyRate, @pdblStatLimit
	
		SET @pdblRedundancyPay = ((@dblRate1 * 1.5) + (@dblRate2) + (@dblRate3 * 0.5)) * @dblTemp2
	
		if @iAfterOct2006 = 0
		begin
			IF @iAgeY = 64 
			BEGIN
				SET @pdblRedundancyPay = @pdblRedundancyPay * (12 - @iAgeM) / 12
			END
		end
	END')

  GRANT EXEC ON [sp_ASRFn_StatutoryRedundancyPay] TO [ASRSysGroup]



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step X of X - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.4')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.4')

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
     INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
    OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASR%')))
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

SELECT @NVarCommand = 'USE master
	GRANT EXECUTE ON sp_OACreate TO public
	GRANT EXECUTE ON sp_OADestroy TO public
	GRANT EXECUTE ON sp_OAGetErrorInfo TO public
	GRANT EXECUTE ON sp_OAGetProperty TO public
	GRANT EXECUTE ON sp_OAMethod TO public
	GRANT EXECUTE ON sp_OASetProperty TO public
	GRANT EXECUTE ON sp_OAStop TO public
	GRANT EXECUTE ON xp_LoginConfig TO public
	GRANT EXECUTE ON xp_EnumGroups TO public'
EXEC sp_executesql @NVarCommand

-- Version specific functions
IF (@iSQLVersion < 11)
BEGIN
	SELECT @NVarCommand = 'USE master
		GRANT EXECUTE ON xp_StartMail TO public
		GRANT EXECUTE ON xp_SendMail TO public';
	EXEC sp_executesql @NVarCommand;
END

SELECT @NVarCommand = 'USE ['+@DBName + ']'
EXEC sp_executesql @NVarCommand


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.4 Of HR Pro'
