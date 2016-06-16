
/* --------------------------------------------------- */
/* Update the database from version 3.4 to version 3.5 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(8000),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(4000),
	@iTableID	integer

DECLARE @sSQL varchar(8000)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
DECLARE @sSPCode_5 nvarchar(4000)
DECLARE @sSPCode_6 nvarchar(4000)
DECLARE @sSPCode_7 nvarchar(4000)
DECLARE @sSPCode_8 nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON
SET @DBName = DB_NAME()

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 3.3 or 3.4. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '3.4') and (@sDBVersion <> '3.5')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2000 or above
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
IF (@iSQLVersion < 8)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of HR Pro', 16, 1)
	RETURN
END

/* ---------------------------------------------------------------------------------- */
PRINT 'Step 1 of 38 - SQL Version reader'

	-- [udfASRSQLVersionNumber]
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRSQLVersion]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRSQLVersion]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRSQLVersion]
	(
	)
	RETURNS integer
	AS
	BEGIN
		RETURN convert(int,convert(float,substring(@@version,charindex(''-'',@@version)+2,2)))
	END'
	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 2 of 38 - Indexing'

	-- ASRSysAccordTransferTypes
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransferTypes') AND name = N'IDX_BaseTableID')
		DROP INDEX ASRSysAccordTransferTypes.[IDX_BaseTableID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_BaseTableID] ON ASRSysAccordTransferTypes ([ASRBaseTableID],[IsVisible])'
	EXEC sp_executesql @NVarCommand


	-- ASRSysSSIHiddenGroups
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSSIHiddenGroups') AND name = N'IDX_LinkID')
		DROP INDEX ASRSysSSIHiddenGroups.[IDX_LinkID]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_LinkID] ON ASRSysSSIHiddenGroups ([LinkID] ASC)'
	EXEC sp_executesql @NVarCommand


	-- ASRSysPurgePeriods
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysPurgePeriods') AND name = N'IDX_PurgeKey')
		DROP INDEX ASRSysPurgePeriods.[IDX_PurgeKey]
	SET @NVarCommand = 'CREATE CLUSTERED INDEX [IDX_PurgeKey] ON ASRSysPurgePeriods ([PurgeKey] ASC)'
	EXEC sp_executesql @NVarCommand


	-- ASRSysSystemSettings
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSystemSettings') AND name = N'Section, SettingKey')
		DROP INDEX ASRSysSystemSettings.[Section, SettingKey]
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysSystemSettings') AND name = N'IDX_SectionSettingKey')
		DROP INDEX ASRSysSystemSettings.[IDX_SectionSettingKey]
	SET @NVarCommand = 'CREATE UNIQUE CLUSTERED INDEX [IDX_SectionSettingKey] ON ASRSysSystemSettings ([Section] ASC, [SettingKey] ASC)'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 3 of 38 - User Setting Stored Procedures'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetUserSetting]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetUserSetting]

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSaveUserSetting]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSaveUserSetting]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetUserSetting]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSaveUserSetting]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetUserSetting]
	(
		@sSection		varchar(50),
		@sSettingKey	varchar(50),
		@sSettingValue	varchar(255) OUTPUT
	)
	AS
	BEGIN
		SET NOCOUNT ON

		SELECT @sSettingValue = [SettingValue]
		FROM ASRSysUserSettings
		WHERE Section = @sSection
			AND SettingKey = @sSettingKey
			AND UserName = System_User
	END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRSaveUserSetting]
	(
		@sSection		varchar(50),
		@sSettingKey	varchar(50),
		@sSettingValue	varchar(255)
	)
	AS
	BEGIN
		SET NOCOUNT ON

		IF EXISTS(SELECT [SettingValue] FROM ASRSysUserSettings WHERE Section = @sSection	 AND SettingKey = @sSettingKey AND UserName = System_User)
			UPDATE ASRSysUserSettings SET [SettingValue] = @sSettingValue WHERE Section = @sSection AND SettingKey = @sSettingKey AND UserName = System_User
		ELSE
			INSERT ASRSysUserSettings ([Section], [SettingKey], [UserName], [SettingValue]) VALUES (@sSection, @sSettingKey, System_User, @sSettingValue)

	END'
	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 4 of 38 - Misc Stored Procedures'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetOrderDefinition]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetOrderDefinition]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRGetOrderDefinition]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRGetOrderDefinition] (
			@piOrderID int) 
		AS
		BEGIN
			/* Return the recordset of order items for the given order. */
			SELECT ASRSysOrderItems.*,
				ASRSysColumns.columnName,
				ASRSysColumns.tableID,
				ASRSysColumns.dataType,
    			ASRSysTables.tableName,
				ASRSysColumns.Size,
				ASRSysColumns.Decimals,
				ASRSysColumns.Use1000Separator
			FROM ASRSysOrderItems
			INNER JOIN ASRSysColumns 
				ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
			INNER JOIN ASRSysTables 
				ON ASRSysTables.tableID = ASRSysColumns.tableID
			WHERE ASRSysOrderItems.orderID = @piOrderID
			ORDER BY ASRSysOrderItems.type, 
				ASRSysOrderItems.sequence
		END'
	EXECUTE (@sSPCode_0)


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAllTablePermissions]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAllTablePermissions]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissions]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = '
	ALTER PROCEDURE [dbo].[sp_ASRAllTablePermissions] 
		(
		@psSQLLogin 		varchar(200)
		)
		AS
		BEGIN

			SET NOCOUNT ON

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

			-- Cached cut down view of the sysprotects table
			DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
			INSERT @SysProtects
				SELECT [ID],[Action],[ProtectType], [Columns] FROM sysprotects
				WHERE [UID] = @iUserGroupID

			SELECT sysobjects.name, p.action
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			WHERE p.protectType <> 206
				AND p.action <> 193
				AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
			UNION
			SELECT sysobjects.name, 193
			FROM syscolumns
			INNER JOIN @SysProtects p ON (syscolumns.id = p.id
				AND p.action = 193 
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
			INNER JOIN sysobjects ON sysobjects.id = p.id
			WHERE syscolumns.name = ''timestamp''
				AND p.protectType IN (204, 205) 
			ORDER BY sysobjects.name

	END'
	EXECUTE (@sSPCode_0)



/* ------------------------------------------------------------- */
PRINT 'Step 5 of 38 - Internal UDF Functions'

	-- [udf_ASRHasFunctionComponent]
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRHasFunctionComponent]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udf_ASRHasFunctionComponent]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udf_ASRHasFunctionComponent]
	(
		@piExpressionID integer,
		@piFunctionID integer
	)
	RETURNS integer
	AS
	BEGIN

		DECLARE @iCount integer
		DECLARE @bFound bit

		SELECT @iCount = COUNT(*)
			FROM ASRSysExprComponents c 
			LEFT OUTER JOIN ASRSysExpressions e ON c.ComponentID = e.parentComponentID
			WHERE c.ExprID = @piExpressionID AND
				((c.Type = 1 AND dbo.udf_ASRHasFunctionComponent(c.FieldSelectionFilter,@piFunctionID) > 0) OR
				(c.Type = 2 AND c.FunctionID = @piFunctionID) OR
				(c.Type = 2 AND dbo.udf_ASRHasFunctionComponent(e.exprID,@piFunctionID) > 0) OR
				(c.Type = 3 AND dbo.udf_ASRHasFunctionComponent(c.CalculationID,@piFunctionID) > 0) OR
				(c.Type = 10 AND dbo.udf_ASRHasFunctionComponent(c.FilterID,@piFunctionID) > 0))

		IF @iCount > 0 SET @bFound = 1
		ELSE SET @bFound = 0
		
		RETURN @bFound

	END'
	EXECUTE (@sSPCode_0)

	-- [udf_ASRHasFunctionComponent]
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udf_ASRRootExpressionID]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udf_ASRRootExpressionID]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udf_ASRRootExpressionID]
	(
		@piComponentID int
	)
	RETURNS int
	AS
	BEGIN

		DECLARE @iRootExpressionID int
		DECLARE @iParentComponentID int

		SELECT @iRootExpressionID = ASRSysExpressions.exprID,
			@iParentComponentID = ASRSysExpressions.parentComponentID
			FROM ASRSysExpressions
			JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID
			WHERE ASRSysExprComponents.componentID = @piComponentID

		IF @iParentComponentID <> 0
		BEGIN
			SELECT @iRootExpressionID = dbo.udf_ASRRootExpressionID(@iParentComponentID)
		END
	
		RETURN @iRootExpressionID

	END'
	EXECUTE (@sSPCode_0)


	-- [udfASRIsServer64Bit]
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRIsServer64Bit]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRIsServer64Bit]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRIsServer64Bit]
		()
		RETURNS int
		AS
		BEGIN

			DECLARE @bIs64Bit bit
			SELECT @bIs64Bit = CASE PATINDEX (''%X64)%'' , @@version)
					WHEN 0 THEN 0
					ELSE 1
				END
			RETURN @bIs64Bit

		END'
	EXECUTE (@sSPCode_0)




/* ------------------------------------------------------------- */
PRINT 'Step 6 of 38 - Modifying column for Data Transfer'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysDataTransferColumns')
	and name = 'FromText'
	and length < 255

	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysDataTransferColumns
                               ALTER COLUMN [FromText] [varchar] (255) NULL '
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 7 of 38 - Creating/modifying Workflow tables'

	-- Create the ASRSysWorkflowElementValidations table
	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowElementValidations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowElementValidations](
				[ID] [int] NOT NULL,
				[ElementID] [int] NOT NULL,
				[ExprID] [int] NOT NULL,
				[Type] [smallint] NOT NULL,
				[Message] [varchar](1000) 
			) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new DescHasWorkflowName column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'DescHasWorkflowName'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						DescHasWorkflowName [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.DescHasWorkflowName = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new DescHasElementCaption column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'DescHasElementCaption'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						DescHasElementCaption [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.DescHasElementCaption = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new Behaviour column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'Behaviour'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						Behaviour [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.Behaviour = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new Mandatory column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'Mandatory'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						Mandatory [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.Mandatory = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new CalcID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'CalcID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						CalcID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.CalcID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new CaptionType column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'CaptionType'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						CaptionType [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.CaptionType = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElementItems - Add new DefaultValueType column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'DefaultValueType'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						DefaultValueType [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElementItems
		SET ASRSysWorkflowElementItems.DefaultValueType = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Add new CompletionCount column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'CompletionCount'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD 
						CompletionCount [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.CompletionCount = 
			CASE 
				WHEN ASRSysWorkflowInstanceSteps.status = 3 THEN 1
				ELSE 0
			END'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Add new FailedCount column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'FailedCount'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD 
						FailedCount [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.FailedCount = 
			CASE 
				WHEN ASRSysWorkflowInstanceSteps.status = 4 THEN 1
				ELSE 0
			END'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Add new TimeoutCount column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'TimeoutCount'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD 
						TimeoutCount [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.TimeoutCount = 
			CASE 
				WHEN ASRSysWorkflowInstanceSteps.status = 6 THEN 1
				ELSE 0
			END'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Extended Value column to varchar(8000)*/
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'Value'
	and length < 8000

	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues
							   ALTER COLUMN Value [varchar] (8000) NULL '
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new EmailCCID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'EmailCCID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						EmailCCID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowElements
		SET ASRSysWorkflowElements.EmailCCID = 0'

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Drop EmailBCCID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'EmailBCCID'

	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements DROP COLUMN
						EmailBCCID'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Add new EmailCC column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'EmailCC'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD 
						EmailCC [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.EmailCC = '''''

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Drop EmailBCC column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'EmailBCC'

	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps DROP COLUMN 
						EmailBCC'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceSteps - Add new HypertextLinkedSteps column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
	and name = 'HypertextLinkedSteps'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD 
						HypertextLinkedSteps [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = '
		UPDATE ASRSysWorkflowInstanceSteps
		SET ASRSysWorkflowInstanceSteps.HypertextLinkedSteps = '''''

		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempValue column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'TempValue'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						TempValue [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempParent1TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'TempParent1TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						TempParent1TableID [integer] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempParent1RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'TempParent1RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						TempParent1RecordID [integer] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempParent2TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'TempParent2TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						TempParent2TableID [integer] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempParent2RecordID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'TempParent2RecordID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						TempParent2RecordID [integer] NULL'
		EXEC sp_executesql @NVarCommand
	END

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowTriggeredLinks]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	drop trigger [dbo].[DEL_ASRSysWorkflowTriggeredLinks]

	SELECT @NVarCommand = '	CREATE TRIGGER [dbo].[DEL_ASRSysWorkflowTriggeredLinks] 
				   ON  [dbo].[ASRSysWorkflowTriggeredLinks] 
				   FOR DELETE
				AS 
				BEGIN
					DELETE FROM ASRSysWorkflowTriggeredLinkColumns WHERE LinkID IN (SELECT LinkID FROM deleted)
				END'
	EXEC sp_executesql @NVarCommand

	/* ASRSysWorkflows - Default the baseTable column for old definitions */
	SELECT TOP 1 @iTableID = convert(int, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_TablePersonnel'

	IF (@iTableID IS null) 
		OR (@iTableID = 0) 
	BEGIN
		SELECT TOP 1 @iTableID = convert(int, isnull(parameterValue, '0'))
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_PERSONNEL'
			AND parameterKey = 'Param_TablePersonnel'
	END

	IF (@iTableID > 0) 
	BEGIN
		UPDATE ASRSysWorkflows
		SET initiationType = 0,
			baseTable = @iTableID
		WHERE isnull(initiationType, 0) = 0
			AND isnull(baseTable, 0) = 0
	END


/* ------------------------------------------------------------- */

PRINT 'Step 8 of 38 - Accord Integration modifications'

	-- Remove obsolete procedures
	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRAccordCanDeleteRecord]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordCanDeleteRecord]

	-- Add force as update column
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferTypes')
	and name = 'ForceAsUpdate'

	IF @iRecCount = 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes ADD
		                       [ForceAsUpdate] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysAccordTransferTypes SET [ForceAsUpdate] = 0'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysAccordTransferTypes SET [ForceAsUpdate] = 1 WHERE TransferTypeID = 71'
		EXEC sp_executesql @NVarCommand
	END

	-- Drop the unused status column
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferTypes')
	and name = 'StatusColumnID'

	IF @iRecCount = 1
	BEGIN
		SET @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes DROP COLUMN [StatusColumnID]'
		EXEC sp_executesql @NVarCommand
	END


	-- Set pension fields to v1.10 spec
	SET @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET [Description] = ''Unused'' WHERE TransferTypeID = 0 AND TransferFieldID IN (57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77)'
	EXEC sp_executesql @NVarCommand


	-- Handle multiple transfer types in payroll check
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordIsRecordInPayroll]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordIsRecordInPayroll]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRAccordIsRecordInPayroll]
		(@iRecordID int,
		@iTransferType int,
		@ProhibitDelete int OUTPUT)
	AS
		BEGIN
		SET NOCOUNT ON
		
		SET @ProhibitDelete = 0

		IF EXISTS(SELECT Status FROM ASRSysAccordTransactions
					WHERE HRProRecordID = @iRecordID
						AND Status IN (10,11)
						AND TransferType = @iTransferType)
			SET @ProhibitDelete = 1

	END'
	EXEC sp_executesql @sSPCode_0


	----------------------------------------------------------------------
	-- spASRAccordPopulateTransaction
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordPopulateTransaction]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordPopulateTransaction]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRAccordPopulateTransaction]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRAccordPopulateTransaction] (
			@piTransactionID int OUTPUT,
			@piTransferType int,
			@piTransactionType int ,
			@piDefaultStatus int,
			@piHRProRecordID int,
			@iTriggerLevel int,
			@pbSendAllFields bit OUTPUT)
			AS
			BEGIN	
		
			-- Return the required user or system setting.
			DECLARE @iCount	integer
			DECLARE @bNewTransaction bit
			DECLARE @iStatus integer
			DECLARE @bCreate bit
			DECLARE @bForceAsUpdate bit
		
			SET @piTransactionID = null
			SET @bCreate = 1
			SET @bForceAsUpdate = 0
		
			SELECT @piTransactionID = TransactionID
				FROM ASRSysAccordTransactionProcessInfo
				WHERE spid = @@SPID AND TransferType = @piTransferType AND RecordID = @piHRProRecordID
		
			-- Could be a null if the trigger was fired from a non Accord module enabled table, e.g. a child updating a parent field
			IF @piTransactionID IS null SET @bNewTransaction = 1
			ELSE SET @bNewTransaction = 0
		
			-- Get a transaction ID for this process and update the temporary Accord table
			IF @bNewTransaction = 1
			BEGIN
				SELECT @iCount = COUNT(*)
					FROM ASRSysSystemSettings
					WHERE section = ''AccordTransfer'' AND settingKey = ''NextTransactionID''
				
				IF @iCount = 0
					INSERT ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES (''AccordTransfer'',''NextTransactionID'',1)
				ELSE
					UPDATE ASRSysSystemSettings SET SettingValue = SettingValue + 1 WHERE section = ''AccordTransfer'' AND settingKey =  ''NextTransactionID''
		
				SELECT @piTransactionID = settingValue 
				FROM ASRSysSystemSettings
				WHERE section = ''AccordTransfer'' AND settingKey =  ''NextTransactionID''
		
				-- If update, has it already been sent?
				IF @piTransactionType = 1
				BEGIN
		
					SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
					WHERE HRProRecordID = @piHRProRecordID
						AND TransferType = @piTransferType
					ORDER BY CreatedDateTime DESC
				
					IF @iStatus IS NULL OR @iStatus = 20
					BEGIN
						SET @piTransactionType = 0
						SET @pbSendAllFields = 1
					END
				END
		
				SELECT @bForceAsUpdate = ForceAsUpdate FROM ASRSysAccordTransferTypes
				WHERE TransferTypeID = @piTransferType
		
				IF @bForceAsUpdate = 1 AND @piTransactionType = 0 SET @piTransactionType = 1
		
				-- Are we trying to delete something thats never been sent?
				IF @piTransactionType = 2
				BEGIN
					SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
					WHERE HRProRecordID = @piHRProRecordID
					ORDER BY CreatedDateTime DESC
				
					IF @iStatus IS NULL	SET @bCreate = 0
					ELSE SET @pbSendAllFields = 1
				END
		
				-- Insert a record into the Accord Transfer table.
				IF @bCreate = 1
				BEGIN
					INSERT INTO ASRSysAccordTransactions
						([TransactionID],[TransferType], [TransactionType], [CreatedUser], [CreatedDateTime], [Status], [HRProRecordID], [Archived])
					VALUES 
						(@piTransactionID, @piTransferType, @piTransactionType, SYSTEM_USER, GETDATE(), @piDefaultStatus, @piHRProRecordID, 0)
		
					INSERT ASRSysAccordTransactionProcessInfo (SPID, TransactionID,TransferType,RecordID) VALUES (@@SPID, @piTransactionID, @piTransferType, @piHRProRecordID)
				END
		
			END
		END'

	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRAccordDeleteTransactionsForRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordDeleteTransactionsForRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordDeleteTransactionsForRecord]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRAccordDeleteTransactionsForRecord]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRAccordDeleteTransactionsForRecord]
	(@iRecordID int
	, @iTransferType int)
	AS
	BEGIN
		SET NOCOUNT ON
		DELETE FROM ASRSysAccordTransactions WHERE HrProRecordID = @iRecordID
			AND TransferType = @iTransferType
	END'

	EXECUTE (@sSPCode_0)









/* ------------------------------------------------------------- */
PRINT 'Step 9 of 38 - Creating/modifying Workflow stored procedures and functions'

	----------------------------------------------------------------------
	-- udfASRWorkflowColumnsUsed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRWorkflowColumnsUsed]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRWorkflowColumnsUsed]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRWorkflowColumnsUsed] ()
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			INSERT @results (id) VALUES (1)
			RETURN
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRWorkflowColumnsUsed] (
			@piWorkflowID		integer,
			@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
			@pfDeleteTrigger	bit			-- 1 when the deleted record is the trigger record
		)
		RETURNS @results TABLE (
			columnID	integer
		)
		AS
		BEGIN
			-- Return a table containing the info of any elements that refer to columns from the 
			-- DeleteTrigger record or the record identified using the given StoredData element.
			DECLARE
				@iBaseTableID		integer,
				@sIdentifier		varchar(8000),
				@iElementType		integer,
				@iEmailType			integer,
				@iEmailColumnID		integer,
				@iEmailExprID		integer,
				@iExprColumnID		integer
		
			-- Create a local table variable to hold the results.
			DECLARE @columnsUsed TABLE (
				columnID	integer
			)
		
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
			INSERT INTO @columnsUsed
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
			INSERT INTO @columnsUsed
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
			INSERT INTO @columnsUsed
			SELECT WEC.dbColumnID
			FROM ASRSysWorkflowElementColumns WEC
			INNER JOIN ASRSysWorkflowElements WE ON WEC.elementID = WE.ID
			INNER JOIN ASRSysColumns Cols ON WEC.dbColumnID = Cols.columnID
			WHERE WE.workflowID = @piWorkflowID
				AND WE.type = 5 -- StoredData
				AND WEC.valueType = 2 -- DBValue	
				AND Cols.tableID = @iBaseTableID
				AND (((@pfDeleteTrigger = 1) AND (WEC.dbRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigg'


	SET @sSPCode_1 = 'er = 0) 
						AND (WEC.dbRecord = 1) -- Identified
						AND (WEC.WFFormIdentifier = @sIdentifier)))
		
			----------------------------------------------------------------------------
			-- Determine which fields from the Deleted record are used in Expressions
			----------------------------------------------------------------------------
			INSERT INTO @columnsUsed
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
		
			-- Read and return the results from the local table variable.
			INSERT @results
				SELECT DISTINCT columnID
				FROM @columnsUsed
		
			RETURN
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- udfASRWorkflowEmailsUsed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRWorkflowEmailsUsed]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRWorkflowEmailsUsed]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRWorkflowEmailsUsed] ()
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			INSERT @results (id) VALUES (1)
			RETURN
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRWorkflowEmailsUsed] (
			@piWorkflowID		integer,
			@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
			@pfDeleteTrigger	bit			-- 1 when the deleted record is the trigger record
		)
		RETURNS @results TABLE (
				emailID		integer,
				type		integer,
				colExprID	integer
		)
		AS
		BEGIN
			-- Return a table containing the info of any email elements that refer to the 
			-- DeleteTrigger record or the record identified using the given StoredData element.
			DECLARE
				@iBaseTableID		integer,
				@sIdentifier		varchar(8000)
		
			-- Create a local table variable to hold the results.
			DECLARE @emailsUsed TABLE (
				emailID		integer,
				type		integer,
				colExprID	integer
			)
		
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
			-- 1) Email To address
			----------------------------------------------------------------------------
			INSERT INTO @emailsUsed
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
			-- 2) Email Copy address
			----------------------------------------------------------------------------
			INSERT INTO @emailsUsed
			SELECT WFE.emailCCID,
				EA.type,
				CASE
					WHEN EA.type = 1 THEN EA.columnID -- Column
					ELSE EA.exprID -- Calculated
				END
			FROM ASRSysWorkflowElements WFE
			INNER JOIN ASRSysEmailAddress EA ON WFE.emailCCID = EA.emailID
			WHERE WFE.workflowID = @piWorkflowID
				AND WFE.type = 3 -- Email
				AND EA.tableID = @iBaseTableID
				AND ((EA.type = 1) OR (EA.type = 2))
				AND (((@pfDeleteTrigger = 1) AND (WFE.emailRecord = 4)) -- Triggered
					OR ((@pfDeleteTrigger = 0) 
						AND (WFE.emailRecord = 1) -- Identified
						AND (WFE.recSelWebFormIdentifier = @sIdentifier)))
		
			-- Read and return the results from the local table variable.
			INSERT @results
				SELECT DISTINCT emailID,
					type,
					colExprID
				FROM @emailsUsed
		
			RETURN
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- udfASRGetAllPrecedingWorkflowElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRGetAllPrecedingWorkflowElements]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRGetAllPrecedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRGetAllPrecedingWorkflowElements] ()
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			INSERT @results (id) VALUES (1)
			RETURN
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRGetAllPrecedingWorkflowElements] (
			@piElementID	integer
		)
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			-- Return a table containing the IDs of ALL elements that precede the given element.
			-- Connectors are ignored.
		
			DECLARE @iRowsAdded integer
		
			-- Create a local table variable to hold the results.
			DECLARE @precedingElements TABLE (
				elementID	integer PRIMARY KEY CLUSTERED,
				type		integer,
				processed	tinyint default 0
			)
		
			-- Add details of preceding elements into the results table.
			-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
			INSERT INTO @precedingElements
			SELECT DISTINCT
				CASE 
					WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
					ELSE E.ID
				END, 
				CASE 
					WHEN E.type = 9 THEN 8 -- 8 = Connector 1
					ELSE E.type
				END, 
				0
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
			WHERE L.endElementID = @piElementID
		
			SET @iRowsAdded = @@rowcount
		
			WHILE @iRowsAdded > 0
			BEGIN
				-- If we''ve just added rows to the results table, process the new rows.
				-- Mark the new rows as ''being processed''.
				UPDATE @precedingElements
				SET processed = 1
				WHERE processed = 0
		
				-- Add details of elements that precede those being processed into the results table.
				-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
				INSERT INTO @precedingElements
				SELECT DISTINCT
					CASE 
						WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
						ELSE E.ID
					END, 
					CASE 
						WHEN E.type = 9 THEN 8 -- 8 = Connector 1
						ELSE E.type
					END, 
					0
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
				INNER JOIN @precedingElements precEl ON L.endElementID = precEl.elementID
				WHERE precEl.processed = 1
					AND CASE 
						WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
						ELSE E.ID
					END NOT IN (SELECT elementID FROM @precedingElements)
					AND CASE 
						WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
						ELSE E.ID
					END <> @piElementID
		
				SET @iRowsAdded = @@rowcount
		
				-- Mark the processed rows as ''been processed''.
				UPDATE @precedingElements
				SET processed = 2
				WHERE processed = 1
			END
		
			-- Read and return the results from the local table variable.
			INSERT @results
				SELECT elementID
				FROM @precedingElements
				WHERE type <> 8
					AND type <> 9
			RETURN
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- udfASRGetPrecedingWorkflowElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRGetPrecedingWorkflowElements]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRGetPrecedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRGetPrecedingWorkflowElements] ()
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			INSERT @results (id) VALUES (1)
			RETURN
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRGetPrecedingWorkflowElements] (
			@piElementID	integer
		)
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			-- Return a table containing the IDs of the elements that precede the given element.
			-- Connectors are ignored.
		
			DECLARE @iRowsAdded integer
		
			-- Create a local table variable to hold the results.
			DECLARE @precedingElements TABLE (
				elementID	integer PRIMARY KEY CLUSTERED,
				type		integer,
				processed	tinyint default 0
			)
		
			-- Add details of preceding elements into the results table.
			-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
			INSERT INTO @precedingElements
			SELECT DISTINCT
				CASE 
					WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
					ELSE E.ID
				END, 
				CASE 
					WHEN E.type = 9 THEN 8 -- 8 = Connector 1
					ELSE E.type
				END, 
				0
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
			WHERE L.endElementID = @piElementID
		
			SET @iRowsAdded = @@rowcount
		
			WHILE @iRowsAdded > 0
			BEGIN
				-- If we''ve just added rows to the results table, process the new rows.
				-- Mark the new rows as ''being processed''.
				UPDATE @precedingElements
				SET processed = 1
				WHERE processed = 0
		
				-- Add details of elements that precede those being processed into the results table.
				-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
				INSERT INTO @precedingElements
				SELECT DISTINCT
					CASE 
						WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
						ELSE E.ID
					END, 
					CASE 
						WHEN E.type = 9 THEN 8 -- 8 = Connector 1
						ELSE E.type
					END, 
					0
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
				INNER JOIN @precedingElements precEl ON L.endElementID = precEl.elementID
				WHERE precEl.processed = 1
					AND precEl.type = 8 -- 8 = Connector 1
					AND CASE 
						WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
						ELSE E.ID
					END NOT IN (SELECT elementID FROM @precedingElements)
		
				SET @iRowsAdded = @@rowcount
		
				-- Mark the processed rows as ''been processed''.
				UPDATE @precedingElements
				SET processed = 2
				WHERE processed = 1
			END
		
			-- Read and return the results from the local table variable.
			INSERT @results
				SELECT elementID
				FROM @precedingElements
				WHERE type <> 8
					AND type <> 9
			RETURN
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- udfASRGetSucceedingWorkflowElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRGetSucceedingWorkflowElements]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRGetSucceedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRGetSucceedingWorkflowElements] ()
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			INSERT @results (id) VALUES (1)
			RETURN
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRGetSucceedingWorkflowElements] (
			@piElementID	integer,
			@piFlowCode		integer
		)
		RETURNS @results TABLE (id integer)
		AS
		BEGIN
			-- Return a table containing the IDs of the elements that succeed the given element
			-- from the given flow.
			-- Connectors are ignored.
		
			DECLARE @iRowsAdded integer
		
			-- Create a local table variable to hold the results.
			DECLARE @succeedingElements TABLE (
				elementID	integer PRIMARY KEY CLUSTERED,
				type		integer,
				processed	tinyint default 0
			)
		
			-- Add details of succeeding elements into the results table.
			-- NB. We skip over the Connector1 elements straight onto the associated Connector 2 elements.
			INSERT INTO @succeedingElements
			SELECT DISTINCT
				CASE 
					WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
					ELSE E.ID
				END, 
				CASE 
					WHEN E.type = 8 THEN 9 -- 9 = Connector 2
					ELSE E.type
				END, 
				0
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND ((L.startOutboundFlowCode = @piFlowCode) OR 
					(@piFlowCode = 0 and L.startOutboundFlowCode = -1))
		
			SET @iRowsAdded = @@rowcount
		
			WHILE @iRowsAdded > 0
			BEGIN
				-- If we''ve just added rows to the results table, process the new rows.
				-- Mark the new rows as ''being processed''.
				UPDATE @succeedingElements
				SET processed = 1
				WHERE processed = 0
		
				-- Add details of elements that succeed those being processed into the results table.
				-- NB. We skip over the Connector1 elements straight onto the associated Connector 2 elements.
				INSERT INTO @succeedingElements
				SELECT DISTINCT
					CASE 
						WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
						ELSE E.ID
					END, 
					CASE 
						WHEN E.type = 8 THEN 9 -- 9 = Connector 2
						ELSE E.type
					END, 
					0
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
				INNER JOIN @succeedingElements succEl ON L.startElementID = succEl.elementID
				WHERE succEl.processed = 1
					AND succEl.type = 9 -- 9 = Connector 2
					AND CASE 
						WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
						ELSE E.ID
					END NOT IN (SELECT elementID FROM @succeedingElements)
		
				SET @iRowsAdded = @@rowcount
		
				-- Mark the processed rows as ''been processed''.
				UPDATE @succeedingElements
				SET processed = 2
				WHERE processed = 1
			END
		
			-- Read and return the results from the local table variable.
			INSERT @results
				SELECT elementID
				FROM @succeedingElements
				WHERE type <> 8
					AND type <> 9
			RETURN
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowActionFailed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowActionFailed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowActionFailed]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowActionFailed]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRWorkflowActionFailed
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psMessage			varchar(8000)
		)
		AS
		BEGIN
			DECLARE
				@iFailureFlows	integer,
				@iCount			integer
		
			-- Check if the failed element has an outbound flow for failures.
			SELECT @iFailureFlows = COUNT(*)
			FROM ASRSysWorkflowElements Es
			INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
				AND Ls.startOutboundFlowCode = 1
			WHERE Es.ID = @piElementID
				AND Es.type = 5 -- 5 = StoredData
		
			IF @iFailureFlows = 0
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 4,	-- 4 = failed
					message = @psMessage,
					failedCount = isnull(failedCount, 0) + 1
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
		
				UPDATE ASRSysWorkflowInstances
				SET status = 2	-- 2 = error
				WHERE ID = @piInstanceID
			END
			ELSE
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 8,	-- 8 = failed action
					message = @psMessage,
					failedCount = isnull(failedCount, 0) + 1
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
		
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 1))
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 3
						OR ASRSysWorkflowInstanceSteps.status = 4
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8)
								
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
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1)
								
				-- Count how many terminators have completed. ie. if the workflow has completed.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1
													
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					
					/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
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
							EXEC [dbo].[spASRSysWorkflowParentRecord]
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
							EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@piParent2TableID,
							@piParent2RecordID,
							0,					
							0,					
							0,					
							0,					
							@piRequiredTableID,
							@piRequiredRecordID OUTPUT
					END
				END
		'

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
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iPersonnelTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2TableID,
						@iInitParent2RecordID,
						'


	SET @sSPCode_1 = '@iDataRecordTableID,
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
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
					EXEC [dbo].[spASRWorkflowValidTableRecord]
						@iDataRecordTableID,
						@piRecordID,
						@fValidRecordID	OUTPUT

					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed]
							@piInstanceID, 
							@piElementID, 
							''Stored Data primary record has been deleted or not selected.''

						SET @psSQL = ''''
						RETURN
					END
				END

				IF @piDataAction = 0 -- Insert
				BEGIN
					IF @iSecond'


	SET @sSPCode_2 = 'aryDataRecord = 0 -- 0 = Initiator''s record
					BEGIN
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
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

					SET @fValidR'


	SET @sSPCode_3 = 'ecordID = 1
					IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
					BEGIN
						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Stored Data secondary record has been deleted or not selected.''

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

					DECLARE @dbValues TABLE (
						ID integer, 
						wfFormIdentifier varchar(1000),
						wfValueIdentifier varchar(1000),
						dbColumnID int,
						dbRecord int,
						value varchar(8000)
					)

					INSERT INTO @dbValues (ID, 
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
					FROM @dbValues
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
'


	SET @sSPCode_4 = '
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

						SET @fValidRecordID = 1
						
						IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
						BEGIN
							SET @fValidRecordID = 0

							EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
								EXEC [dbo].[spASRWorkflowValidTableRecord]
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
											SET @fDe'


	SET @sSPCode_5 = 'letedValue = 1
										END
									END
								END
							END

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC [dbo].[spASRWorkflowActionFailed]
									@piInstanceID, 
									@piElementID, 
									''Stored Data column database value record has been deleted or not selected.''

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

						UPDATE @dbValues
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
							FROM @dbValues dbV
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
											WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS'


	SET @sSPCode_6 = ' null) THEN ''null''
											ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
										END
									ELSE isnull(@sValue, 0) -- integer, logic, numeric
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
					exec [dbo].[spASRGetParentDetails]
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
					DECLARE curColumns CURSOR LOCAL FAST_FORWARD FOR 
					SELECT columnID
					FROM [dbo].[udfASRWorkflowColumnsUsed] (@iWorkflowID, @piE'


	SET @sSPCode_7 = 'lementID, 0)

					OPEN curColumns

					FETCH NEXT FROM curColumns INTO @iDBColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
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
								
						FETCH NEXT FROM curColumns INTO @iDBColumnID
					END
					CLOSE curColumns
					DEALLOCATE curColumns

					DECLARE curEmails CURSOR LOCAL FAST_FORWARD FOR 
					SELECT emailID,
						type,
						colExprID
					FROM [dbo].[udfASRWorkflowEmailsUsed] (@iWorkflowID, @piElementID, 0)

					OPEN curEmails

					FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID
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
							EXEC [dbo].[spASRSysEmailAddr]
								@sDBValue OUTPUT,
								@iEmailID,
								@piRecordID
						END

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID)
								
						FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID
					END
					CLOSE curEmails
					DEALLOCATE curEmails
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
	-- spASRGetWorkflowDelegates
	----------------------------------------------------------------------

	-- Create dummy UDF so collation sequence doesn't cause a problem
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRGetWorkflowDelegatedRecords]')
			AND OBJECTPROPERTY(id, N'IsTableFunction') = 1)
		DROP FUNCTION [dbo].[udfASRGetWorkflowDelegatedRecords]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRGetWorkflowDelegatedRecords]
		(
			@psOriginalRecipient varchar(8000)
		)
		RETURNS @results TABLE (
			id integer,
			emailAddress varchar(8000),
			delegated bit,
			delegatedTo varchar(8000)
		)
		AS
		BEGIN
			-- Will get regenerated in System manager save
			RETURN
		END'
	EXECUTE (@sSPCode_0)


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowDelegates]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowDelegates]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowDelegates]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowDelegates] 
		(
			@psTo			varchar(8000),
			@piStepID		integer,
			@results		cursor varying output
		)
		AS
		BEGIN
			DECLARE
				@iDelegateEmailID	integer,
				@sTemp				varchar(8000),
				@iDelegateRecordID	integer,
				@sDelegateTo		varchar(8000),
				@iCount				integer,
				@sSQL				nvarchar(4000),
				@iInstanceID		integer
		
			IF len(ltrim(rtrim(@psTo))) = 0 RETURN
		
		    DECLARE @recipients TABLE (
		        recordID		integer,
				emailAddress	varchar(8000),
				delegated		bit,
				delegatedTo		varchar(8000),
				processed		tinyint default 0,
				isDelegate		bit
		    )
				
			-- Get the delegate email address definition. 
			SET @iDelegateEmailID = 0
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail''
			SET @iDelegateEmailID = convert(integer, @sTemp)
				
			IF @iDelegateEmailID = 0
			BEGIN
				INSERT INTO @recipients (
					recordID,
					emailAddress,
					delegated,
					delegatedTo,
					processed,
					isDelegate)
				VALUES (
					0, -- Personnel Record ID
					@psTo, -- Email Address(es)
					0, -- Delegated
					'''', -- Delegate Email Address
					2, -- Processed
					0) -- Is Delegate
			END
			ELSE
			BEGIN
				INSERT INTO @recipients 
				SELECT         
					RECnew.ID,
					RECnew.emailAddress,
					RECnew.delegated,
					RECnew.delegatedTo,
					0,
					0
				FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@psTo) RECnew
				WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
					AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
						FROM @recipients RECold
						WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID)
		
				SELECT @iCount = COUNT(*)
				FROM @recipients
				WHERE processed = 0
		
				WHILE @iCount > 0
				BEGIN
					-- Mark the new rows as ''being processed''.
					UPDATE @recipients
					SET processed = 1
					WHERE processed = 0
		
					DECLARE delegatesCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT recordID
					FROM @recipients
					WHERE recordID > 0
						AND processed = 1
						AND delegated = 1
		
					OPEN delegatesCursor
					FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @sDelegateTo = ''''
						SET @sSQL = ''spASRSysEmailAddr''
			
						IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
						BEGIN
							-- Get the delegate''s email address
							EXEC @sSQL @sDelegateTo OUTPUT, @iDelegateEmailID, @iDelegateRecordID
							IF @sDelegateTo IS null SET @sDelegateTo = ''''
						END
		
						IF len(@sDelegateTo) > 0 
						BEGIN
							UPDATE @recipients 
							SET delegatedTo = @sDelegateTo
							WHERE recordID = @iDelegateRecordID
		
							INSERT INTO @recipients 
							SELECT         
								RECnew.ID,
								RECnew.emailAddress,
								RECnew.delegated,
								RECnew.delegatedTo,
								0,
								1
							FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@sDelegateTo) RECnew
							WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
								AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
									FROM @recipients RECold
									WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID)
						END
						ELSE
						BEGIN
							UPDATE @recipients 
							SET delegated = 0
							WHERE recordID = @iDelegateRecordID
						END
		
						FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID
					END
					CLOSE delegatesCursor
					DEALLOCATE delegatesCursor
		
					-- Mark the processed rows as ''been processed''.
					UPDATE @recipients
					SET processed = 2
					WHERE processed = 1
		
					SELECT @iCount = COUNT(*)
					FROM @recipients
					WHERE processed = 0
				END
			END
		
			-- Return the cursor of succeeding elements. 
			SET @results = CURSOR FORWARD_ONLY STA'


	SET @sSPCode_1 = 'TIC FOR
		        SELECT DISTINCT 
					emailAddress,
					delegated,
					delegatedTo,
					isDelegate
		        FROM @recipients
		
			OPEN @results
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRDelegateWorkflowEmail
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDelegateWorkflowEmail]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDelegateWorkflowEmail]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRDelegateWorkflowEmail]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRDelegateWorkflowEmail] 
		(
			@psTo			varchar(8000),
			@psCopyTo		varchar(8000),
			@psMessage		varchar(8000),
			@psMessage_HypertextLinks		varchar(8000),
			@piStepID		integer,
			@psEmailSubject	varchar(8000)
		)
		AS
		BEGIN
			DECLARE
				@sTo				varchar(8000),
				@sAddress			varchar(8000),
				@iInstanceID		integer,
				@curRecipients		cursor,
				@sEmailAddress		varchar(8000),
				@fDelegated			bit,
				@sDelegatedTo		varchar(8000),
				@fIsDelegate		bit
		
			SET @psMessage = isnull(@psMessage, '''')
			SET @psMessage_HypertextLinks = isnull(@psMessage_HypertextLinks, '''')
			IF (len(ltrim(rtrim(@psTo))) = 0) RETURN
		
			-- Get the instanceID of the given step
			SELECT @iInstanceID = instanceID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ID = @piStepID
				
		    DECLARE @recipients TABLE (
				emailAddress	varchar(8000),
				delegated		bit,
				delegatedTo		varchar(8000),
				isDelegate		bit
		    )
		
			exec [dbo].[spASRGetWorkflowDelegates] 
				@psTo, 
				@piStepID, 
				@curRecipients output
			FETCH NEXT FROM @curRecipients INTO 
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate
			WHILE (@@fetch_status = 0)
			BEGIN
				INSERT INTO @recipients
					(emailAddress,
					delegated,
					delegatedTo,
					isDelegate)
				VALUES (
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate
				)
				
				FETCH NEXT FROM @curRecipients INTO 
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate
			END
			CLOSE @curRecipients
			DEALLOCATE @curRecipients
		
			-- Clear out the delegation record for the current step
			DELETE FROM ASRSysWorkflowStepDelegation
			WHERE stepID = @piStepID
		
			INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
			SELECT DISTINCT emailAddress, @piStepID
			FROM @recipients
			WHERE isDelegate = 1
		
			SET @sTo = ''''
			
			DECLARE toCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT ltrim(rtrim(emailAddress))
			FROM @recipients
			WHERE len(ltrim(rtrim(emailAddress))) > 0
		
			OPEN toCursor
			FETCH NEXT FROM toCursor INTO @sAddress
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sTo = @sTo
					+ CASE 
						WHEN len(ltrim(rtrim(@sTo))) > 0 THEN '';''
						ELSE ''''
					END 
					+ @sAddress
		
				FETCH NEXT FROM toCursor INTO @sAddress
			END
			CLOSE toCursor
			DEALLOCATE toCursor
		
			IF len(@sTo) > 0
			BEGIN
				INSERT ASRSysEmailQueue(
					RecordDesc,
					ColumnValue, 
					DateDue, 
					UserName, 
					[Immediate],
					RecalculateRecordDesc, 
					RepTo,
					MsgText,
					WorkflowInstanceID, 
					Subject)
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@sTo,
					@psMessage + @psMessage_HypertextLinks,
					@iInstanceID,
					@psEmailSubject)
			END
		
			IF (len(@psCopyTo) > 0) AND (len(@psMessage) > 0)
			BEGIN
				INSERT ASRSysEmailQueue(
					RecordDesc,
					ColumnValue, 
					DateDue, 
					UserName, 
					[Immediate],
					RecalculateRecordDesc, 
					RepTo,
					MsgText,
					WorkflowInstanceID, 
					Subject)
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@psCopyTo,
					''You have been copied in on the following HR Pro Workflow email with recipients:'' + CHAR(13)
						+ CHAR(9) + @sTo + CHAR(13)	+ CHAR(13)
						+ @psMessage,
					@iInstanceID,
					@psEmailSubject)
			END
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
						@piElementID			integer
					)
					AS
					BEGIN
						/* Cancel (ie. set status to 0 for all workflow pending (ie. status 1 or 2) elements that precede the given element.
						This ignores connection elements.
						NB. This does work for elements with multiple inbound flows. */
						UPDATE ASRSysWorkflowInstanceSteps
						SET status = 0
						WHERE instanceID = @piInstanceID
							AND elementID IN (SELECT ID FROM [dbo].[udfASRGetAllPrecedingWorkflowElements](@piElementID))
							AND status IN (1, 2, 7) -- 1 = pending engine action, 2 = pending user action, 7 = pending user completion
					END
		'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowSubmitImmediatesAndGetSucceedingElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@succeedingElements	cursor varying output,
			@psTo				varchar(8000)
		)
		AS
		BEGIN
			-- Action any immediate elements (Ors & Decisions) and return the IDs of the workflow elements that 
			-- succeed them.
			-- This ignores connection elements.
			DECLARE
				@iTempID		integer,
				@iElementID		integer,
				@iElementType	integer,
				@iFlowCode		integer,
				@iTrueFlowType	integer,
				@iExprID		integer,
				@iResultType	integer,
				@sResult		varchar(8000),
				@fResult		bit,
				@dtResult		datetime,
				@fltResult		float,
				@iValue			integer,
				@iPrecedingElementType	integer, 
				@iPrecedingElementID	integer, 
				@iCount			integer,
				@iStepID		integer,
				@curRecipients		cursor,
				@sEmailAddress		varchar(8000),
				@fDelegated			bit,
				@sDelegatedTo		varchar(8000),
				@fIsDelegate		bit
		
			DECLARE @elements table
			(
				elementID		integer,
				elementType		integer,
				processed		tinyint default 0,
				trueFlowType	integer,
				trueFlowExprID	integer
			)
							
			INSERT INTO @elements 
				(elementID,
				elementType,
				processed,
				trueFlowType,
				trueFlowExprID)
			SELECT SUCC.id,
				E.type,
				0,
				isnull(E.trueFlowType, 0),
				isnull(E.trueFlowExprID, 0)
			FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 0) SUCC
			INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
				
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE (elementType = 4 OR elementType = 7) -- 4=Decision, 7=Or
				AND processed = 0
		
			WHILE @iCount > 0
			BEGIN
				UPDATE @elements
				SET processed = 1
				WHERE processed = 0
		
				-- Action any succeeding immediate elements (Decisions & Ors)
				DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					E.elementType,
					E.trueFlowType, 
					E.trueFlowExprID
				FROM @elements E
				WHERE (E.elementType = 4 OR E.elementType = 7) -- 4=Decision, 7=Or
					AND E.processed = 1
		
				OPEN immediateCursor
				FETCH NEXT FROM immediateCursor INTO 
					@iElementID, 
					@iElementType, 
					@iTrueFlowType, 
					@iExprID
				WHILE (@@fetch_status = 0)
				BEGIN
					-- Submit the immediate elements, and get their succeeding elements
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.userEmail = ASRSysWorkflowInstanceSteps.userEmail,
						ASRSysWorkflowInstanceSteps.message = '''',
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
		
					SET @iFlowCode = 0
		
					IF @iElementType = 4 -- Decision
					BEGIN
						IF @iTrueFlowType = 1
						BEGIN
							-- Decision Element flow determined by a calculation
							EXEC [dbo].[spASRSysWorkflowCalculation]
								@piInstanceID,
								@iExprID,
								@iResultType OUTPUT,
								@sResult OUTPUT,
								@fResult OUTPUT,
								@dtResult OUTPUT,
								@fltResult OUTPUT, 
								0
		
							SET @iValue = convert(integer, @fResult)
						END
						ELSE
						BEGIN
							-- Decision Element flow determined by a button in a preceding web form
							SET @iPrecedingElementType = 4 -- Decision element
							SET @iPrecedingElementID = @iElementID
		
							WHILE (@iPrecedingElementType = 4)
							BEGIN
								SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
									@iPrecedingElementType = isnull(WE.type, 0)
								FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPrecedingElementID) PE
								INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
								INNER JOIN ASRSysWorkflowInstanceSt'


	SET @sSPCode_1 = 'eps WIS ON PE.ID = WIS.elementID
									AND WIS.instanceID = @piInstanceID
		
								SET @iPrecedingElementID = @iTempID
							END
							
							SELECT @iValue = 
								CASE
									WHEN isnumeric(IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
									ELSE 0
								END
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
							WHERE IV.elementID = @iPrecedingElementID
							AND IV.instanceid = @piInstanceID
								AND E.ID = @iElementID
						END
						
						IF @iValue IS null SET @iValue = 0
						SET @iFlowCode = @iValue
		
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
					END
					ELSE IF @iElementType = 7 -- Or
					BEGIN
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @piInstanceID, @iElementID
					END
		
					-- Get this immediate element''s succeeding elements
					INSERT INTO @elements 
						(elementID,
						elementType,
						processed,
						trueFlowType,
						trueFlowExprID)
					SELECT SUCC.id,
						E.type,
						0,
						isnull(E.trueFlowType, 0),
						isnull(E.trueFlowExprID, 0)
					FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, @iFlowCode) SUCC
					INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
					WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements)
		
					FETCH NEXT FROM immediateCursor INTO 
						@iElementID, 
						@iElementType, 
						@iTrueFlowType, 
						@iExprID
				END
				CLOSE immediateCursor
				DEALLOCATE immediateCursor
		
				UPDATE @elements
				SET processed = 2
				WHERE processed = 1
		
				SELECT @iCount = COUNT(*)
				FROM @elements
				WHERE (elementType = 4 OR elementType = 7) -- 4=Decision, 7=Or
					AND processed = 0
			END
		
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE elementType = 2 -- 2=WebForm
		
			IF (@iCount > 0) AND len(ltrim(rtrim(@psTo))) > 0 
			BEGIN
				SELECT @iStepID = ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
		
				DECLARE @recipients TABLE (
					emailAddress	varchar(8000),
					delegated		bit,
					delegatedTo		varchar(8000),
					isDelegate		bit
				)
		
				exec [dbo].[spASRGetWorkflowDelegates] 
					@psTo, 
					@iStepID, 
					@curRecipients output
				FETCH NEXT FROM @curRecipients INTO 
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO @recipients
						(emailAddress,
						delegated,
						delegatedTo,
						isDelegate)
					VALUES (
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate
					)
					
					FETCH NEXT FROM @curRecipients INTO 
							@sEmailAddress,
							@fDelegated,
							@sDelegatedTo,
							@fIsDelegate
				END
				CLOSE @curRecipients
				DEALLOCATE @curRecipients
		
				DELETE FROM ASRSysWorkflowStepDelegation
				WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
					FROM ASRSysWorkflowInstanceSteps
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT E.elementID
							FROM @elements E
							WHERE E.elementType = 2) -- 2 = WebForm
						AND (ASRSysWorkflowInstanceSteps.status = 0
							OR ASRSysWorkflowInstanceSteps.status = 2
							OR ASRSysWorkflowInstanceSteps.status = 6
							OR ASRSysWorkflowInstanceSteps.status = 8
							OR ASRSysWorkflowInstanceSteps.status = 3))
		
				INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
				SELECT DISTINCT RECS.emailAddress, WIS.ID
				FROM @recipients '


	SET @sSPCode_2 = 'RECS, 
					ASRSysWorkflowInstanceSteps WIS
				WHERE RECS.isDelegate = 1
					AND WIS.instanceID = @piInstanceID
						AND WIS.elementID IN 
							(SELECT E.elementID
							FROM @elements E
							WHERE E.elementType = 2) -- 2 = WebForm
						AND (WIS.status = 0
							OR WIS.status = 2
							OR WIS.status = 6
							OR WIS.status = 8
							OR WIS.status = 3)
			END
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 1,
				ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.completionDateTime = null,
				ASRSysWorkflowInstanceSteps.userEmail = CASE
					WHEN (SELECT ASRSysWorkflowElements.type 
						FROM ASRSysWorkflowElements 
						WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @psTo -- 2 = Web Form element
					ELSE ASRSysWorkflowInstanceSteps.userEmail
				END
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType <> 7 -- 7 = Or
						AND E.elementType <> 4) -- 4 = Decision
				AND (ASRSysWorkflowInstanceSteps.status = 0
					OR ASRSysWorkflowInstanceSteps.status = 2
					OR ASRSysWorkflowInstanceSteps.status = 6
					OR ASRSysWorkflowInstanceSteps.status = 8
					OR ASRSysWorkflowInstanceSteps.status = 3)
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 2
			WHERE ASRSysWorkflowInstanceSteps.id IN (
				SELECT ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2)
		
			-- Return the cursor of succeeding elements. 
			SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
				SELECT elementID 
				FROM @elements E
				WHERE E.elementType <> 7 -- 7 = Or
					AND E.elementType <> 4 -- 4 = Decision
		
			OPEN @succeedingElements
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
						@iParent2RecordID	integer
		
					DECLARE @succeedingElements table(elementID int)
				
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
				
						INSERT INTO @succeedingElements 
						SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
				
						INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
						SELECT 
							@iInstanceID, 
							ASRSysWorkflowElements.ID, 
							CASE
								WHEN ASRSysWorkflowElements.type = 0 THEN 3
								WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
								FROM @succeedingElements) THEN 1
								ELSE 0
							END, 
							CASE
								WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
								WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
								FROM @succeedingElements) THEN getdate()
								ELSE null
							END, 
							CASE
								WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
								ELSE null
							END, 
							CASE
								WHEN ASRSysWorkflowElements.type = 0 THEN 1
								ELSE 0
							END,
							0,
							0
						FROM ASRSysWorkflowElements 
						WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
						
						-- Create the Workflow Instance Value records. 
						INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
						SELECT @iInstanceID, ASRSysWorkflowElements.ID, 
							ASRSysWorkflowElementItems.identifier
						FROM ASRSysWorkflowElementItems 
						INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
'


	SET @sSPCode_1 = '							AND ASRSysWorkflowElements.type = 2
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
	-- spASRGetActiveWorkflowStoredDataSteps
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetActiveWorkflowStoredDataSteps]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetActiveWorkflowStoredDataSteps]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetActiveWorkflowStoredDataSteps]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetActiveWorkflowStoredDataSteps
		AS
		BEGIN
			/* Return a recordset of the workflow StoredData steps that need to be actioned by the Workflow service. */
			DECLARE @steps table(ID integer)
		
			INSERT INTO @steps
			SELECT S.ID
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE S.status = 1
				AND E.type = 5 -- 5 = Stored Data
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET status = 5 -- In progress
			WHERE ID IN (SELECT ID FROM @steps)
		
			SELECT S.instanceID AS [instanceID],
				E.ID AS [elementID],
				S.ID AS [stepID]
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE s.ID IN (SELECT ID FROM @steps)
		END
		'

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

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
							(
								@piInstanceID	integer,
								@piElementID	integer,
								@psMessage		varchar(8000)	OUTPUT, 
								@psMessage_HypertextLinks		varchar(8000)	OUTPUT, 
								@psHypertextLinkedSteps			varchar(8000)	OUTPUT, 
								@pfOK			bit	OUTPUT,
								@psTo			varchar(8000)
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
									@iColumnID			integer,
									@iResultType	integer,
									@sResult		varchar(8000),
									@fResult		bit,
									@dtResult		datetime,
									@fltResult		float,
									@iCalcID		integer,
									@iSQLVersion	integer
											
								SET @pfOK = 1
								SET @psMessage = ''''
								SET @psMessage_HypertextLinks = ''''
								SET @psHypertextLinkedSteps = ''''
								SELECT @iSQLVersion = dbo.udfASRSQLVersion()
							
								exec [dbo].[spASRGetSetting]
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
									EI.wfValu'


	SET @sSPCode_1 = 'eIdentifier, 
									EI.recSelWebFormIdentifier,
									EI.recSelIdentifier, 
									EI.calcID
								FROM ASRSysWorkflowElementItems EI
								WHERE EI.elementID = @piElementID
								ORDER BY EI.ID
							
								OPEN itemCursor
								FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID
								WHILE (@@fetch_status = 0)
								BEGIN
									SET @sValue = ''''
		
									IF @iItemType = 1
									BEGIN
										SET @fDeletedValue = 0
		
										/* Database value. */
										SELECT @sTableName = ASRSysTables.tableName, 
											@iRequiredTableID = ASRSysTables.tableID, 
											@sColumnName = ASRSysColumns.columnName, 
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
													@iParent1TableID = IV.parent1T'


	SET @sSPCode_2 = 'ableID,
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
		
											EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
												EXEC [dbo].[spASRWorkflowValidTableRecord]
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
												SET @sValue ='


	SET @sSPCode_3 = ' convert(varchar(8000), @dtTempDate, @iEmailFormat)
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
											AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
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
					
									IF @iItemType = 16
									BEGIN
										/* Calculation. */
										EXEC [dbo].[spASRSysWorkflowCalculation]
											@piInstanceID,
											@iCalcID,
											@iResultType OUTPUT,
											@sResult OUTPUT,
											@fResult OUTPUT,
											@dtResult OUTPUT,
											@fltResult OUTPUT, 
											0
		
										SET @psMessage = @psMessage +
											@sResult
									END
							
									FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFor'


	SET @sSPCode_4 = 'mIdentifier, @sRecSelIdentifier, @iCalcID
								END
								CLOSE itemCursor
								DEALLOCATE itemCursor
							
								/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
								CREATE TABLE #succeedingElements (elementID integer)
							
								EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
									@piInstanceID, 
									@piElementID, 
									@superCursor OUTPUT,
									@psTo
							
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
									SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) + CHAR(13)
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
									WHERE WE.type = 2 -- 2 = Web Form element
								
									OPEN elementCursor
									FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
									WHILE (@@fetch_status = 0)
									BEGIN
			
										IF @iSQLVersion = 8
										BEGIN
											EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
											IF @hResult <> 0
											BEGIN
												SET @sQueryString = ''''
											END
											ELSE
											BEGIN
												EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @sQueryString OUTPUT, @piInstanceID, @iElementID, @sParam1, @@servername, @sDBName
												IF @hResult <> 0 
												BEGIN
													SET @sQueryString = ''''
												END
		
												EXEC @hResult = sp_OADestroy @objectToken 
											END
										END
										ELSE
										BEGIN
											SELECT @sQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @iElementID, @sParam1, @@servername, @sDBName)
										END
													
										IF LEN(@sQueryString) = 0 
										BEGIN
											SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) +
												@sCaption + '' - Error constructing the query string. Please contact your system administrator.''
										END
										ELSE
										BEGIN
											SET @psHypertextLinkedSteps = @psHypertextLinkedSteps
												+ CASE
													WHEN len(@psHypertextLinkedSteps) = 0 THEN char(9)
													ELSE ''''
												END 
												+ convert(varchar(8000), @iElementID)
												+ char(9)
		
											SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) +
												@sCaption + '' - '' + CHAR(13) + 
												''<'' + @sURL + ''?'' + @sQueryString + ''>''
										END
										
										FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
									END
									CLOSE elementCursor
							
									DEALLOCATE elementCursor
		
									SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) + CHAR(13)
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
								'


	SET @sSPCode_5 = '		+ '' been cut off you will need to copy and paste ''
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
		+ @sSPCode_4
		+ @sSPCode_5)


	----------------------------------------------------------------------
	-- spASRGetWorkflowQueryString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowQueryString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowQueryString]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowQueryString
		(
			@piInstanceID	integer,
			@piElementID	integer,
			@psQueryString	varchar(8000)	output
		)
		AS
		BEGIN
			DECLARE
				@hResult		integer,
				@objectToken	integer,
				@sURL			varchar(8000),
				@sParam1		varchar(8000),
				@sDBName		sysname,
				@sSQLVersion	varchar(2)
		
			SET @psQueryString = ''''
			SET @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
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
		
			IF (len(@sURL) > 0)
			BEGIN
				SET @sDBName = db_name()
		
				IF @sSQLVersion = ''8''
				BEGIN			
					EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
			
					IF (@hResult = 0) 
					BEGIN
						EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @psQueryString OUTPUT, @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName
						IF @hResult <> 0
						BEGIN
							SET @psQueryString = ''''
						END
			
						EXEC sp_OADestroy @objectToken
					END
				END
				ELSE
				BEGIN
					SELECT @psQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName)
				END
			
				IF len(@psQueryString) > 0
				BEGIN
					SET @psQueryString = @sURL + ''?'' + @psQueryString
				END
			END
		END'

	EXECUTE (@sSPCode_0)

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
			@psFormElements		varchar(8000)	OUTPUT,
			@pfSavedForLater	bit				OUTPUT
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
				@hTmpResult		integer,
				@sTo			varchar(8000),
				@sCopyTo		varchar(8000),
				@sTempTo		varchar(8000),
				@sMessage		varchar(8000),
				@sMessage_HypertextLinks	varchar(8000),
				@sHypertextLinkedSteps		varchar(8000),
				@iEmailID		integer,
				@iEmailCopyID		integer,
				@iTempEmailID		integer,
				@iEmailLoop		integer,
				@iEmailRecord		integer,
				@iEmailRecordID	integer,
				@sSQL			nvarchar(4000),
				@iCount		integer,
				@superCursor		cursor,
				@curDelegatedRecords	cursor,
				@fDelegate		bit,
				@fDelegationValid	bit,
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
				@sEmailSubject	varchar(200),
				@iTempID	integer,
				@iBehaviour		integer

			SET @pfSavedForLater = 0

			SELECT @iCurrentStepID = ID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

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
				@iEmailCopyID = isnull(E.emailCCID, 0),
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

			--------------------------------------------------
			-- Read the submitted webForm/storedData values
			--------------------------------------------------
			IF @iElementType = 5 -- Stored Data element
			BEGIN
				SET @sValue = @psFormInput1
				SET @sValueDescription = ''''
				SET @sMessage = ''Successfully '' +
					CASE
						WHEN @iDataAction = 0 THEN ''inserted''
						WHEN @iDataAction = 1 THEN '


	SET @sSPCode_1 = '''updated''
						ELSE ''deleted''
					END + '' record''

				IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
				BEGIN
					SET @sValueDescription = @psFormInput2
				END
				ELSE
				BEGIN
					SET @iTemp = convert(integer, @sValue)
					IF @iTemp > 0 
					BEGIN	
						EXEC [dbo].[spASRRecordDescription] 
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
					SET @sValue = left(@sValue, 8000)

					--Get the record description (for RecordSelectors only)
					SET @sValueDescription = ''''

					-- Get the WebForm item type, etc.
					SELECT @sIdentifier = EI.identifier,
						@iItemType = EI.itemType,
						@iTableID = EI.tableID,
						@iBehaviour = EI.behaviour
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
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalR'


	SET @sSPCode_2 = 'ecDesc
						END

						-- Record the selected record''s parent details.
						exec [dbo].[spASRGetParentDetails]
							@iTableID,
							@iTemp,
							@iParent1TableID	OUTPUT,
							@iParent1RecordID	OUTPUT,
							@iParent2TableID	OUTPUT,
							@iParent2RecordID	OUTPUT
					END
					ELSE
					IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
					BEGIN
						SET @pfSavedForLater = 1
					END

					UPDATE ASRSysWorkflowInstanceValues
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

				IF @pfSavedForLater = 1
				BEGIN
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 7
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

					RETURN
				END
			END
					
			SET @hResult = 0
			SET @sTo = ''''
			SET @sCopyTo = ''''
		
			--------------------------------------------------
			-- Process email element
			--------------------------------------------------
			IF @iElementType = 3 -- Email element
			BEGIN
				-- Get the email recipient. 
				SET @iEmailRecordID = 0
				SET @sSQL = ''spASRSysEmailAddr''

				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @iEmailLoop = 0
					WHILE @iEmailLoop < 2
					BEGIN
						SET @hTmpResult = 0
						SET @sTempTo = ''''
						SET @iTempEmailID = 
							CASE 
								WHEN @iEmailLoop = 1 THEN @iEmailCopyID
								ELSE isnull(@iEmailID, 0)
							END

						IF @iTempEmailID > 0 
						BEGIN
							SET @fValidRecordID = 1

							SELECT @iEmailTableID = isnull(tableID, 0),
								@iEmailType = isnull(type, 0)
							FROM ASRSysEmailAddress
							WHERE emailID = @iTempEmailID

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
									SELECT @iPrevElementType'


	SET @sSPCode_3 = ' = ASRSysWorkflowElements.type,
										@iTempElementID = ASRSysWorkflowElements.ID
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

									SET @iBaseRecordID = @iEmailRecordID
								END

								SET @fValidRecordID = 1
								IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
								BEGIN
									SET @fValidRecordID = 0

									EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
										EXEC [dbo].[spASRWorkflowValidTableRecord]
											@iEmailTableID,
											@iEmailRecordID,
											@fValidRecordID	OUTPUT
									END

									IF @fValidRecordID = 0
									BEGIN
										IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
										BEGIN
											SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
											FROM ASRSysWorkflowQueueColumns QC
											INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
											WHERE WFQ.instanceID = @piInstanceID
												AND QC.emailID = @iTempEmailID

											IF len(@sTempTo) > 0 SET @fValidRecordID = 1
										END
										ELSE
										BEGIN
											IF @iEmailRecord = 1
											BEGIN
												SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
												FROM ASRSysWorkflowInstanceValues IV
												WHERE IV.instanceID = @piInstanceID
													AND IV.emailID = @iTempEmailID
													AND IV.elementID = @iTempElementID
												IF len(@sTempTo) > 0 SET @fValidRecordID = 1
											END
										END
									END

									IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Update the ASRSysWorkflowInstanceSteps'


	SET @sSPCode_4 = ' table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
											@piInstanceID, 
											@piElementID, 
											''Email record has been deleted or not selected.''
													
										SET @hTmpResult = -1
									END
								END
							END

							IF @fValidRecordID = 1
							BEGIN
								/* Get the recipient address. */
								IF len(@sTempTo) = 0
								BEGIN
									EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID
									IF @sTempTo IS null SET @sTempTo = ''''
								END

								IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
								BEGIN
									-- Email step failure if no known recipient.
									-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
									EXEC [dbo].[spASRWorkflowActionFailed] 
										@piInstanceID, 
										@piElementID, 
										''No email recipient.''
												
									SET @hTmpResult = -1
								END
							END

							IF @iEmailLoop = 1 
							BEGIN
								SET @sCopyTo = @sTempTo

								IF (rtrim(ltrim(@sCopyTo)) = ''@'')
									OR (charindex('' @ '', @sCopyTo) > 0)
								BEGIN
									SET @sCopyTo = ''''
								END
							END
							ELSE
							BEGIN
								SET @sTo = @sTempTo
							END
						END
						
						SET @iEmailLoop = @iEmailLoop + 1

						IF @hTmpResult <> 0 SET @hResult = @hTmpResult
					END
				END
		
				IF LEN(rtrim(ltrim(@sTo))) > 0
				BEGIN
					IF (rtrim(ltrim(@sTo)) = ''@'')
						OR (charindex('' @ '', @sTo) > 0)
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.userEmail = @sTo,
							ASRSysWorkflowInstanceSteps.emailCC = @sCopyTo
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

						EXEC [dbo].[spASRWorkflowActionFailed] 
							@piInstanceID, 
							@piElementID, 
							''Invalid email recipient.''
						
						SET @hResult = -1
					END
					ELSE
					BEGIN
						/* Build the email message. */
						EXEC [dbo].[spASRGetWorkflowEmailMessage] 
							@piInstanceID, 
							@piElementID, 
							@sMessage OUTPUT, 
							@sMessage_HypertextLinks OUTPUT, 
							@sHypertextLinkedSteps OUTPUT, 
							@fValidRecordID OUTPUT, 
							@sTo
		
						IF @fValidRecordID = 1
						BEGIN
							exec [dbo].[spASRDelegateWorkflowEmail] 
								@sTo,
								@sCopyTo,
								@sMessage,
								@sMessage_HypertextLinks,
								@iCurrentStepID,
								@sEmailSubject
						END
						ELSE
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Email item database value record has been deleted or not selected.''
										
							SET @hResult = -1
						END
					END
				END
			END
		
			--------------------------------------------------
			-- Mark the step as complete
			--------------------------------------------------
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
					ASRSysWorkflowInstanceSteps.emailCC = CASE
						WHEN @iElementType = 3 THEN @sCopyTo
						ELSE ASRSysWorkflowInstanceSteps.emailCC
					END,
					ASRSysWorkflowInstanceSteps.hypertextLinkedSteps = CASE
						WHEN @iElementType = 3 THEN @sHypertextLinkedSteps
						ELSE ASRSysWorkflowInstanceSteps.hypertextLinkedSteps
					END,
					ASRSysWorkflowInstanceSteps.messa'


	SET @sSPCode_5 = 'ge = CASE
						WHEN @iElementType = 3 THEN LEFT(@sMessage, 8000)
						WHEN @iElementType = 5 THEN LEFT(@sMessage, 8000)
						ELSE ''''
					END,
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iElementType = 4 -- Decision element
				BEGIN
					IF @iTrueFlowType = 1
					BEGIN
						-- Decision Element flow determined by a calculation
						EXEC [dbo].[spASRSysWorkflowCalculation]
							@piInstanceID,
							@iExprID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT, 
							0

						SET @iValue = convert(integer, @fResult)
					END
					ELSE
					BEGIN
						-- Decision Element flow determined by a button in a preceding web form
						SET @iPrevElementType = 4 -- Decision element
						SET @iPreviousElementID = @piElementID

						WHILE (@iPrevElementType = 4)
						BEGIN
							SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
								@iPrevElementType = isnull(WE.type, 0)
							FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
							INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
							INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
								AND WIS.instanceID = @piInstanceID

							SET @iPreviousElementID = @iTempID
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
			
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 1,
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.completionDateTime = null
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT SUCC.id FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, @iValue) SUCC)
						AND (ASRSysWorkflowInstanceSteps.status = 0
							OR ASRSysWorkflowInstanceSteps.status = 2
							OR ASRSysWorkflowInstanceSteps.status = 6
							OR ASRSysWorkflowInstanceSteps.status = 8
							OR ASRSysWorkflowInstanceSteps.status = 3)
				END
				ELSE
				BEGIN
					IF @iElementType <> 3 -- 3=Email element
					BEGIN
						-- Do not the following bit when the submitted element is an Email element as 
						-- the succeeding elements will already have been actioned.
						DECLARE @succeedingElements TABLE(elementID integer)
		
						EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
							@piInstanceID, 
							@piElementID, 
							@superCursor OUTPUT,
							''''
		
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO @succeedingElements (elementID) VALUES (@iTemp)
							
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
			'


	SET @sSPCode_6 = '				WHERE WIS.instanceID = @piInstanceID
								AND WIS.elementID = @piElementID

							-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
							DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
							SELECT ASRSysWorkflowInstanceSteps.ID,
								ASRSysWorkflowInstanceSteps.elementID
							FROM ASRSysWorkflowInstanceSteps
							INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND ASRSysWorkflowElements.type = 2
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
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
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
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
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
									(SELECT ASRSysWorkflowElements.ID
									FROM ASRSysWorkflowElements
									WHERE ASRSysWorkflowElements.type = 2)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3)
						END
						ELSE
						BEGIN
							DELETE FROM ASRSysWorkflowStepDelegation
							WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
								FROM ASRSysWorkflowInstanceSteps
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN 
										(SELECT suc.elementID
										FROM @succeedingElements suc)
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3))
							
							I'


	SET @sSPCode_7 = 'NSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
							(SELECT WSD.delegateEmail,
								SuccWIS.ID
							FROM ASRSysWorkflowStepDelegation WSD
							INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
							INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
								AND SuccWIS.elementID IN (SELECT suc.elementID
									FROM @succeedingElements suc)
								AND (SuccWIS.status = 0
									OR SuccWIS.status = 2
									OR SuccWIS.status = 6
									OR SuccWIS.status = 8
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
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3)
						END
					END
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
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
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
					AND ASRSysWorkflowElements.type = 1
							
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
						AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pend'


	SET @sSPCode_8 = 'ing Engine Action
							OR ASRSysWorkflowInstanceSteps.status = 2) -- 2 = Pending User Action
				END

				IF @iElementType = 3 -- Email element
					OR @iElementType = 5 -- Stored Data element
				BEGIN
					exec [dbo].[spASREmailImmediate] ''HR Pro Workflow''
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
		+ @sSPCode_7
		+ @sSPCode_8)

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

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			-- Return a recordset of the workflow steps that need to be actioned by the Workflow service.
			-- Action any that can be actioned immediately. 
			DECLARE
				@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
				@iElementType		integer,
				@iInstanceID		integer,
				@iElementID			integer,
				@iStepID			integer,
				@iCount				integer,
				@sStatus			bit,
				@sMessage			varchar(8000),
				@iTemp				integer, 
				@iTemp2				integer, 
				@iTemp3				integer,
				@sForms 			varchar(8000), 
				@iType				integer,
				@iDecisionFlow		integer,
				@fInvalidElements	bit, 
				@fValidElements	bit, 
				@iPrecedingElementID	integer, 
				@iPrecedingElementType	integer, 
				@iPrecedingElementStatus	integer, 
				@iPrecedingElementFlow	integer, 
				@fSaveForLater		bit
		
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
					-- Check if all preceding steps have completed before submitting this step.
					SET @fInvalidElements = 0		
				
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID

					OPEN precedingElementsCursor
			
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow

					WHILE (@@fetch_status = 0)
						AND (@fInvalidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*) 
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID

								IF @iCount = 0 SET @fInvalidElements = 1
							END
							ELSE
							BEGIN
								SET @fInvalidElements = 1
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END

									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID
								
									IF @iCount = 0 SET @fInvalidElements = 1
								END
								ELSE
								BEGIN
									SET @fInvalidElements = 1
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingE'


	SET @sSPCode_1 = 'lementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END

										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID
									
										IF @iCount = 0 SET @fInvalidElements = 1
									END
									ELSE
									BEGIN
										SET @fInvalidElements = 1
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus <> 3 SET @fInvalidElements = 1
								END
							END
						END

						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					END
					CLOSE precedingElementsCursor
					DEALLOCATE precedingElementsCursor
					
					IF (@fInvalidElements = 0) SET @iAction = 1
				END
		
				IF @iAction = 4 -- Or check
				BEGIN
					SET @fValidElements = 0		
					-- Check if any preceding steps have completed before submitting this step. 
		
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID

					OPEN precedingElementsCursor		
		
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					WHILE (@@fetch_status = 0)
						AND (@fValidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID
							
								IF @iCount > 0 SET @fValidElements = 1
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END

									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID
							
									IF @iCount > 0 SET @fValidElements = 1
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingElementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END

										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID

										IF @iCount > 0 SET @fValidElements = 1
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus = 3 SET @fValidElements = 1
								END
							END
						END

						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					END
					CLOSE precedingElementsCursor
					DEALLOCATE precedingElementsCursor
		
					-- If all preceding steps have been completed submit the Or step.
					IF @fValidElements > '


	SET @sSPCode_2 = '0 
					BEGIN
						-- Cancel any preceding steps that are not completed as they are no longer required.
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @iInstanceID, @iElementID
		
						SET @iAction = 1
					END
				END
		
				IF @iAction = 1
				BEGIN
					EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '''', '''', @sForms OUTPUT, @fSaveForLater OUTPUT
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 2
					WHERE id = @iStepID
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
			END

			CLOSE stepsCursor
			DEALLOCATE stepsCursor

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
				SET ASRSysWorkflowInstanceSteps.status = 6, -- Timeout
					ASRSysWorkflowInstanceSteps.timeoutCount = isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID

				-- Activate the succeeding elements on the Timeout flow
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1))
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 3
						OR ASRSysWorkflowInstanceSteps.status = 4
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8)
					
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
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1)
					
				-- Count how many term'


	SET @sSPCode_3 = 'inators have completed. ie. if the workflow has completed.
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
					
					-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table.
				END

				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			END

			CLOSE timeoutCursor
			DEALLOCATE timeoutCursor
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3)

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

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRInstantiateWorkflow]
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
						@fSaveForLater bit	
					DECLARE @succeedingElements table(elementID int)
				
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
						EXEC [dbo].[spASRWorkflowUsesInitiator]
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
		
						exec [dbo].[spASRGetParentDetails]
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
		
					INSERT INTO @succeedingElements 
					SELECT id 
					FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
				
					INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
					SELECT 
						@piInstanceID, 
						ASRSysWorkflowElements.ID, 
						CASE
							WHEN ASRSysWorkflowElements.ty'


	SET @sSPCode_1 = 'pe = 0 THEN 3
							WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
								FROM @succeedingElements suc) THEN 1
							ELSE 0
						END, 
						CASE
							WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
							WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
								FROM @succeedingElements suc) THEN getdate()
							ELSE null
						END, 
						CASE
							WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
							ELSE null
						END, 
						CASE
							WHEN ASRSysWorkflowElements.type = 0 THEN 1
							ELSE 0
						END,
						0,
						0
					FROM ASRSysWorkflowElements 
					WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID
				
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
							AND (ASRSysWorkflowElements.type = 7 -- Or
								OR ASRSysWorkflowElements.type = 4) -- Decision
							AND ASRSysWorkflowElements.workflowID = @piWorkflowID
							AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
					WHILE @iCount > 0 
					BEGIN
						DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT ASRSysWorkflowInstanceSteps.elementID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.status = 1
							AND (ASRSysWorkflowElements.type = 7 -- Or
								OR ASRSysWorkflowElements.type = 4) -- Decision
							AND ASRSysWorkflowElements.workflowID = @piWorkflowID
							AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
		
						OPEN immediateSubmitCursor
						FETCH NEXT FROM immediateSubmitCursor INTO @iElementID
						WHILE (@@fetch_status = 0) 
						BEGIN
							EXEC [dbo].[spASRSubmitWorkflowStep] 
								@piInstanceID, 
								@iElementID, 
								'''', 
								'''', 
								@sForms OUTPUT, 
								@fSaveForLater OUTPUT
		
							FETCH NEXT FROM immediateSubmitCursor INTO @iElementID
						END
						CLOSE immediateSubmitCursor
						DEALLOCATE immediateSubmitCursor
		
						SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
							FROM ASRSysWorkflowInstanceSteps
							INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
							WHERE ASRSysWorkflowInstanceSteps.status = 1
								AND (ASRSysWorkflowElements.type = 7 -- Or
									OR ASRSysWorkflowElements.type = 4) -- Decision
								AND ASRSysWorkflowElements.workflowID = @piWorkflowID
								AND ASRSysWorkflowInstanceSteps.instanceID '


	SET @sSPCode_2 = '= @piInstanceID		
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
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
				
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
				END
		'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

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

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowFormItems]
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
					@iDefaultValueType		integer,
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
					@iColumnID	integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iCalcID	integer,
					@iSize		integer,
					@iDecimals	integer,
					@sIdentifier	varchar(8000)

				DECLARE @itemValues table(ID integer, value varchar(8000), type integer)	
						
				-- Check the given instance still exists.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				IF @iCount = 0
				BEGIN
					SET @psErrorMessage = ''This workflow step is invalid. The workflow process may have been completed.''
					RETURN
				END
			
				-- Check if the step has already been completed!
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
			
				DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowElementItems.ID,
					ASRSysWorkflowElementItems.itemType,
					ASRSysWorkflowElementItems.dbColumnID,
					ASRSysWorkflowElementItems.dbRecord,
					ASRSysWorkflowElementItems.wfFormIdentifier,
					'


	SET @sSPCode_1 = 'ASRSysWorkflowElementItems.wfValueIdentifier,
					ASRSysWorkflowElementItems.calcID,
					ASRSysWorkflowElementItems.identifier,
					isnull(ASRSysWorkflowElementItems.defaultValueType, 0) AS [defaultValueType],
					isnull(ASRSysWorkflowElementItems.inputSize, 0),
					isnull(ASRSysWorkflowElementItems.inputDecimals, 0)
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.elementID = @piElementID
					AND (ASRSysWorkflowElementItems.itemType = 1 
						OR (ASRSysWorkflowElementItems.itemType = 2 AND ASRSysWorkflowElementItems.captionType = 3)
						OR ASRSysWorkflowElementItems.itemType = 3
						OR ASRSysWorkflowElementItems.itemType = 5
						OR ASRSysWorkflowElementItems.itemType = 6
						OR ASRSysWorkflowElementItems.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 4)
			
				OPEN itemCursor
				FETCH NEXT FROM itemCursor INTO 
					@iID, 
					@iItemType, 
					@iDBColumnID, 
					@iDBRecord, 
					@sWFFormIdentifier, 
					@sWFValueIdentifier, 
					@iCalcID, 
					@sIdentifier, 
					@iDefaultValueType,
					@iSize,
					@iDecimals
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iItemType = 1
					BEGIN
						SET @fDeletedValue = 0

						-- Database value. 
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
									A'


	SET @sSPCode_2 = 'ND Es.workflowID = @iWorkflowID
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

							EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
								EXEC [dbo].[spASRWorkflowValidTableRecord]
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
								EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @piElementID, ''Web Form item record has been deleted or not selected.''
											
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
									'' FROM '' + '


	SET @sSPCode_3 = '@sTableName +
									'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
							SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
						END
					END

					IF @iItemType = 4
					BEGIN
						-- Workflow value.
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
		
						IF @iType = 14 -- Lookup, need to get the column data type
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

					IF @iItemType = 2 
					BEGIN
						-- Label with calculated caption
						EXEC [dbo].[spASRSysWorkflowCalculation]
							@piInstanceID,
							@iCalcID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT, 
							0

						SET @sValue = @sResult
						SET @iType = 3 -- Character
					END

					IF (@iItemType = 3)
						OR (@iItemType = 5)
						OR (@iItemType = 6)
						OR (@iItemType = 7)
						OR (@iItemType = 11)
					BEGIN
						IF @iStatus = 7 -- Previously SavedForLater
						BEGIN
							SELECT @sValue = 
								CASE
									WHEN (@iItemType = 6 AND IVs.value = ''1'') THEN ''TRUE'' 
									WHEN (@iItemType = 6 AND IVs.value <> ''1'') THEN ''FALSE'' 
									WHEN (@iItemType = 7 AND (upper(ltrim(rtrim(IVs.value))) = ''NULL'')) THEN '''' 
									ELSE isnull(IVs.value, '''')
								END
							FROM ASRSysWorkflowInstanceValues IVs
							WHERE IVs.instanceID = @piInstanceID
								AND IVs.elementID = @piElementID
								AND IVs.identifier = @sIdentifier
						END
						ELSE	
						BEGIN
							IF @iDefaultValueType = 3 -- Calculated
							BEGIN
								EXEC [dbo].[spASRSysWorkflowCalculation]
									@piInstanceID,
									@iCalcID,
									@iResultType OUTPUT,
									@sResult OUTPUT,
									@fResult OUTPUT,
									@dtResult OUTPUT,
									@fltResult OUTPUT, 
									0

								IF @iItemType = 3 SET @sResult = LEFT(@sResult, @iSize)
								IF @iItemType = 5
								BEGIN
									IF @fltResult >= power(10, @iSize - @iDecimals) SET @fltResult = 0
									IF @fltResult <= (-1 * power(10, @iSize - @iDecimals)) SET @fltResult = 0
								END

								SET @sValue = 
									CASE
										WHEN @iResultType = 2 THEN STR(@fltResult, 8000, @iDecimals)
										WHEN @iResultType = 3 THEN 
											CASE 
												WHEN @fResult = 1 THEN ''TRUE''
												ELSE ''FALSE''
											END
										WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
										ELSE convert(varchar(8000), @sResult)
									END

								SET @iType = @iResultType
							END
							ELSE
							BEGIN
								SELECT @sValue = isnull(EIs.inputDefault, '''')
								FROM ASRSysWorkflowElementItems EIs
								WHERE EIs.elementID = @piElementID
									AND EIs.ID = @iID
							END
						END
					END		
			
					INSERT INTO @itemValues (ID, value, type)
					VALUES ('


	SET @sSPCode_4 = '@iID, @sValue, @iType)
			
					FETCH NEXT FROM itemCursor INTO 
						@iID, 
						@iItemType, 
						@iDBColumnID, 
						@iDBRecord, 
						@sWFFormIdentifier, 
						@sWFValueIdentifier, 
						@iCalcID, 
						@sIdentifier, 
						@iDefaultValueType,
						@iSize,
						@iDecimals
				END
				CLOSE itemCursor
				DEALLOCATE itemCursor
			
				SELECT thisFormItems.*, 
					IV.value, 
					IV.type AS [sourceItemType]
				FROM ASRSysWorkflowElementItems thisFormItems
				LEFT OUTER JOIN @itemValues IV ON thisFormItems.ID = IV.ID
				WHERE thisFormItems.elementID = @piElementID
				ORDER BY thisFormItems.ZOrder DESC
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

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
						DECLARE @joinParents table(tableID		integer)	
					
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
						FETCH NEXT FROM orderCursor INTO @sColumnNa'


	SET @sSPCode_1 = 'me, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
						WHILE (@@fetch_status = 0)
						BEGIN
							IF @sOrderItemType = ''F''
							BEGIN
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
								FROM @joinParents
								WHERE tableID = @iTempTableID
					
								IF @iTempCount = 0
								BEGIN
									INSERT INTO @joinParents (tableID) VALUES(@iTempTableID)
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
								JP.tableID
							FROM @joinParents JP
							INNER JOIN ASRSysTables ON JP.tableID = ASRSysTables.tableID
					
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
											WHEN isnumeric('


	SET @sSPCode_2 = 'IV.value) = 1 THEN convert(integer, ISNULL(IV.value, ''0''))
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
							END
					
							IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
							BEGIN
								EXEC [dbo].[spASRWorkflowAscendantRecordID]
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
		
								EXEC [dbo].[spASRWorkflowValidTableRecord]
									@iRecordTableID,
									@iRecordID,
									@fValidRecordID	OUTPUT
		
								IF @fValidRecordID  = 0
								BEGIN
									SET @pfOK = 0
		
									-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
									EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @iElementID, ''Web Form record selector item record has been deleted or not selected.''
									
									-- Need to return a recordset of some kind.
									SELECT '''' AS ''Error''
		
									RETURN
								END
							END
		
							IF @iFilterID > 0 
							BEGIN
								SET @sFilterUDF = ''[dbo].udf_ASRWFExpr_'' + convert(varchar(8000), @iFilterID)
		
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
					END
		'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRGetWorkflowItemValues
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowItemValues]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowItemValues]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowItemValues]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowItemValues]
			(
				@piElementItemID	integer,
				@piInstanceID	integer
			)
			AS
			BEGIN
				DECLARE 
					@iItemType			integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iDefaultValueType		integer,
					@iCalcID				integer,
					@iLookupColumnID	integer,
					@sDefaultValue		varchar(8000),
					@sTableName			sysname,
					@sColumnName		sysname,
					@iDataType			integer,
					@sSelectSQL			varchar(8000),
					@iStatus			integer,
					@iElementID			integer,
					@sValue				varchar(8000),
					@sIdentifier		varchar(8000)
				DECLARE @dropdownValues TABLE([value] varchar(255))

				SELECT 			
					@iItemType = ASRSysWorkflowElementItems.itemType,
					@sDefaultValue = ASRSysWorkflowElementItems.inputDefault,
					@iLookupColumnID = ASRSysWorkflowElementItems.lookupColumnID,
					@iElementID = ASRSysWorkflowElementItems.elementID,
					@sIdentifier = ASRSysWorkflowElementItems.identifier,
					@iCalcID = isnull(ASRSysWorkflowElementItems.calcID, 0),
					@iDefaultValueType = isnull(ASRSysWorkflowElementItems.defaultValueType, 0)
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID

				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID

				IF @iStatus = 7 -- Previously SavedForLater
				BEGIN
					SELECT @sValue = isnull(IVs.value, '''')
					FROM ASRSysWorkflowInstanceValues IVs
					WHERE IVs.instanceID = @piInstanceID
						AND IVs.elementID = @iElementID
						AND IVs.identifier = @sIdentifier

					SET @sDefaultValue = @sValue
				END
				ELSE
				BEGIN
					IF @iDefaultValueType = 3 -- Calculated
					BEGIN
						EXEC [dbo].[spASRSysWorkflowCalculation]
							@piInstanceID,
							@iCalcID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT, 
							0

						SET @sDefaultValue = 
							CASE
								WHEN @iResultType = 2 THEN convert(varchar(8000), @fltResult)
								WHEN @iResultType = 3 THEN 
									CASE 
										WHEN @fResult = 1 THEN ''TRUE''
										ELSE ''FALSE''
									END
								WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
								ELSE convert(varchar(8000), @sResult)
							END
					END
				END

				IF @iItemType = 15 -- OptionGroup
				BEGIN
					SELECT ASRSysWorkflowElementItemValues.value,
						CASE
							WHEN ASRSysWorkflowElementItemValues.value = @sDefaultValue THEN 1
							ELSE 0
						END
					FROM ASRSysWorkflowElementItemValues
					WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
					ORDER BY ASRSysWorkflowElementItemValues.sequence
				END

				IF @iItemType = 13 -- Dropdown
				BEGIN
					INSERT INTO @dropdownValues ([value])
						VALUES (null)

					INSERT INTO @dropdownValues ([value])
						SELECT ASRSysWorkflowElementItemValues.value
						FROM ASRSysWorkflowElementItemValues
						WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
						ORDER BY [sequence]

					SELECT [value],
						CASE
							WHEN [value] = @sDefaultValue THEN 1
							ELSE 0
						END
					FROM @dropdownValues 
				END
				
				IF (@iItemType = 14) -- Lookup
					AND (@iLookupColumnID > 0)
				BEGIN
					SELECT @sTableName = ASRSysTables.tableName,
						@sColumnName = ASRSysColumns.columnName,
						@iDataType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iLookupColumnID

					IF @iDataType = 11 -- Date 
						AND UPPER(LTRIM(RTRIM(@sDefaultValue))) = ''NULL''
					BEGIN
						SET @sDefaultValue = ''''
					END

'


	SET @sSPCode_1 = '					SET @sSelectSQL = ''SELECT null AS [value], '' 
						+ CASE 
							WHEN @sDefaultValue = '''' THEN ''1''
							ELSE ''0''
						END
						+ '' UNION SELECT DISTINCT '' + @sTableName + ''.'' + @sColumnName + '' AS [value],''

					IF len(ltrim(rtrim(@sDefaultValue))) = 0 
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ '' 0''
					END
					ELSE
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ '' CASE''
							+ ''   WHEN '' + @sTableName + ''.'' + @sColumnName + '' = ''
							+ CASE
								WHEN (@iDataType = 12) -- Character
									OR (@iDataType = -1) -- WorkingPattern 
									OR (@iDataType = 11) -- Date 
									THEN '''''''' + @sDefaultValue + ''''''''
								ELSE @sDefaultValue 
							END
							+ ''   THEN 1''
							+ ''   ELSE 0''
							+ '' END''
					END
					SET @sSelectSQL = @sSelectSQL
						+ '' FROM '' + @sTableName
						+ '' ORDER BY [value]''

					EXEC (@sSelectSQL)
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

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

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowStepDescription]
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
				@fltResult		float,
				@fDescHasWorkflowName	bit,
				@fDescHasElementCaption	bit,
				@sWorkflowName			varchar(8000),
				@sElementCaption		varchar(8000)
		
			-- Get the InstanceID and associated DescriptionExprID of the given step
			SELECT @iInstanceID = isnull(WIS.instanceID, 0),
				@iExprID = isnull(WEs.descriptionExprID, 0),
				@fDescHasWorkflowName = isnull(WEs.descHasWorkflowName, 0),
				@fDescHasElementCaption = isnull(WEs.descHasElementCaption, 0),
				@sWorkflowName = isnull(Ws.name, ''''),
				@sElementCaption = isnull(WEs.caption, '''')
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WEs ON WIS.elementID = WEs.ID
			INNER JOIN ASRSysWorkflows Ws ON WEs.workflowID = Ws.ID
			WHERE WIS.ID = @piInstanceStepID
		
			IF @iExprID > 0
			BEGIN
				EXEC [dbo].[spASRSysWorkflowCalculation]
					@iInstanceID,
					@iExprID,
					@iResultType OUTPUT,
					@sResult OUTPUT,
					@fResult OUTPUT,
					@dtResult OUTPUT,
					@fltResult OUTPUT, 
					0
			END
		
			IF @fDescHasWorkflowName = 1
			BEGIN
				SET @sResult = @sWorkflowName 
					+ '' - ''
					+ isnull(@sResult, '''')
			END
		
			IF @fDescHasElementCaption = 1
			BEGIN
				SET @sResult = @sElementCaption 
					+ '' - ''
					+ isnull(@sResult, '''')
			END
		
			SELECT @psDescription = isnull(@sResult, '''')
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRActionOrsAndGetSucceedingWorkflowElements - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionOrsAndGetSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionOrsAndGetSucceedingWorkflowElements]

	----------------------------------------------------------------------
	-- spASRGetDecisionSucceedingWorkflowElements - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDecisionSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDecisionSucceedingWorkflowElements]

	----------------------------------------------------------------------
	-- spASRGetPrecedingWorkflowElements - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetPrecedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetPrecedingWorkflowElements]

	----------------------------------------------------------------------
	-- spASRGetSucceedingWorkflowElements - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetSucceedingWorkflowElements]

	----------------------------------------------------------------------
	-- spASRIgnoreOrsAndGetSucceedingWorkflowElements - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRIgnoreOrsAndGetSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIgnoreOrsAndGetSucceedingWorkflowElements]

	----------------------------------------------------------------------
	-- spASRWorkflowColumnsUsed - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowColumnsUsed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowColumnsUsed]

	----------------------------------------------------------------------
	-- spASRWorkflowEmailsUsed - NO LONGER REQUIRED
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowEmailsUsed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowEmailsUsed]

/* ------------------------------------------------------------- */
PRINT 'Step 10 of 38 - Updating Server DLL version check'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetServerDLLVersion]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetServerDLLVersion]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetServerDLLVersion]
		(
			@strVersion varchar(255) OUTPUT
		)
		AS
		BEGIN

			SET NOCOUNT ON

			DECLARE @objectToken int
			DECLARE @hResult int

			IF EXISTS(SELECT SettingValue
					  FROM   ASRSysSystemSettings
					  WHERE  [Section] = ''server dll''
						AND  [SettingKey] = ''disable check''
						AND  [SettingValue] = ''1'')
			BEGIN

				SELECT @strVersion = [SettingValue]
				FROM   ASRSysSystemSettings
				WHERE  [Section] = ''server dll''
				  AND  [SettingKey] = ''minimum version''

			END
			ELSE
			BEGIN
				-- Create Server DLL object
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsGeneral'', @objectToken OUTPUT
				IF @hResult = 0
					EXEC @hResult = sp_OAMethod @objectToken, ''GetVersion'', @strVersion OUTPUT
				ELSE
					SET @strVersion = ''0.0.0''
						
				EXEC sp_OADestroy @objectToken
					
			END

		END'
	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 11 of 38 - Updating Current User Check'

	----------------------------------------------------------------------
	-- spASRGetCurrentUsers
	----------------------------------------------------------------------

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

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetCurrentUsers
		AS
		BEGIN
		
			SET NOCOUNT ON
		
			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				AND APP_NAME() NOT LIKE ''HR Pro Server.Net%''
				AND APP_NAME() NOT LIKE ''HR Pro Workflow%''
				AND APP_NAME() NOT LIKE ''HR Pro Outlook%''
			BEGIN
				EXEC sp_ASRIntCheckPolls
			END
		
			DECLARE @sCurrentUsers nvarchar(4000)
			DECLARE @sSQLVersion char(2)
			DECLARE @Mode smallint
		
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
			IF @@ROWCOUNT = 0 SET @Mode = 0
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			IF (@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8
			BEGIN
				EXECUTE dbo.[spASRGetCurrentUsersFromAssembly]
			END
			ELSE
			BEGIN
				EXECUTE dbo.[spASRGetCurrentUsersFromMaster]
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersAppName
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersAppName]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE spASRGetCurrentUsersAppName
					(
						@psAppName		varchar(8000) OUTPUT,
						@psUserName		varchar(8000)
					)
		    AS
		    BEGIN
		
		        IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
		        BEGIN
		            EXEC sp_ASRIntCheckPolls
		        END
		
		
		        SELECT TOP 1 @psAppName = rtrim(p.program_name)
		        FROM master..sysprocesses p
		        WHERE p.program_name LIKE ''HR Pro%''
					AND p.program_name NOT LIKE ''HR Pro Workflow%''
		            AND p.program_name NOT LIKE ''HR Pro Outlook%''
		            AND p.program_name NOT LIKE ''HR Pro Server.Net%''
					AND p.loginame = @psUsername
		        GROUP BY p.hostname
		               , p.loginame
		               , p.program_name
		               , p.hostprocess
		        ORDER BY p.loginame
		
		    END'

	EXECUTE (@sSPCode_0)


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
		
				--IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				--BEGIN
				--	EXEC sp_ASRIntCheckPolls
				--END
		
				SELECT p.hostname, p.loginame, p.program_name, p.hostprocess
					   , p.sid, p.login_time, p.spid
				FROM master..sysprocesses p
				JOIN master..sysdatabases d ON     d.dbid = p.dbid
				WHERE p.program_name LIKE ''HR Pro%''
					AND p.program_name NOT LIKE ''HR Pro Workflow%''
		            AND p.program_name NOT LIKE ''HR Pro Outlook%''
					AND p.program_name NOT LIKE ''HR Pro Server.Net%''
					AND d.name = db_name()
				ORDER BY loginame
		
		    END'

	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRGetCurrentUsersCountInApp
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountInApp]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
					(
						@piCount		integer		OUTPUT
					)
		    AS
		    BEGIN
		
				SET NOCOUNT ON
		
				DECLARE @sSQLVersion char(2)
				DECLARE @Mode smallint
		
				IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				BEGIN
					EXEC sp_ASRIntCheckPolls
				END
		
				SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
				SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
				IF @@ROWCOUNT = 0 SET @Mode = 0
				
				IF (@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8
				BEGIN
					SELECT @piCount = dbo.[udfASRNetCountCurrentUsersInApp](APP_NAME())
				END
				ELSE
				BEGIN
		
					SELECT @piCount = COUNT(p.Program_Name)
					FROM     master..sysprocesses p
					JOIN     master..sysdatabases d
					  ON     d.dbid = p.dbid
					WHERE    p.program_name = APP_NAME()
					  AND    d.name = db_name()
					GROUP BY p.program_name
				END
		
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

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetCurrentUsersCountOnServer
					(
						@iLoginCount	integer OUTPUT,
						@psLoginName	varchar(8000)
					)
		    AS
		    BEGIN
		
				DECLARE @sSQLVersion char(2)
				DECLARE @Mode smallint
		
				IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
				BEGIN
					EXEC sp_ASRIntCheckPolls
				END
		
				SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
				SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
				IF @@ROWCOUNT = 0 SET @Mode = 0
				
				IF (@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8
				BEGIN
					SELECT @iLoginCount = dbo.[udfASRNetCountCurrentLogins](@psLoginName)
				END
				ELSE
				BEGIN
		
					SELECT @iLoginCount = COUNT(*)
					FROM master..sysprocesses p
					WHERE p.program_name LIKE ''HR Pro%''
						AND p.program_name NOT LIKE ''HR Pro Workflow%''
		                AND p.program_name NOT LIKE ''HR Pro Outlook%''
		                AND p.program_name NOT LIKE ''HR Pro Server.Net%''
					    AND p.loginame = @psLoginName
				END
		
		    END'

	EXECUTE (@sSPCode_0)



/* ------------------------------------------------------------- */
PRINT 'Step 12 of 38 - Audit Log Changes'

	SELECT @NVarCommand = 'ALTER TABLE ASRSysAuditAccess ALTER COLUMN Action VARCHAR(20) NOT NULL'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 38 - Index Defragmentation'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDefragIndexes]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDefragIndexes]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRDefragIndexes]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRDefragIndexes
			(@maxfrag DECIMAL)
		AS
			BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @tablename VARCHAR (128)
			DECLARE @execstr VARCHAR (255)
			DECLARE @objectid INT
			DECLARE @objectowner VARCHAR(255)
			DECLARE @indexid INT
			DECLARE @frag DECIMAL
			DECLARE @indexname CHAR(255)
			DECLARE @dbname sysname
			DECLARE @tableid INT
			DECLARE @tableidchar VARCHAR(255)
			DECLARE @sSQLVersion nvarchar(20)
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			-- Checking fragmentation
			DECLARE tables CURSOR FOR
				SELECT convert(varchar,so.id)
				FROM sysobjects so
				JOIN sysindexes si ON so.id = si.id
				WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
		
			-- Create the temporary table to hold fragmentation information
			CREATE TABLE #fraglist (
				ObjectName CHAR (255),
				ObjectId INT,
				IndexName CHAR (255),
				IndexId INT,
				Lvl INT,
				CountPages INT,
				CountRows INT,
				MinRecSize INT,
				MaxRecSize INT,
				AvgRecSize INT,
				ForRecCount INT,
				Extents INT,
				ExtentSwitches INT,
				AvgFreeBytes INT,
				AvgPageDensity INT,
				ScanDensity DECIMAL,
				BestCount INT,
				ActualCount INT,
				LogicalFrag DECIMAL,
				ExtentFrag DECIMAL)
		
			-- Open the cursor
			OPEN tables
		
			-- Loop through all the tables in the database running dbcc showcontig on each one
			FETCH NEXT FROM tables INTO @tableidchar
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				-- Do the showcontig of all indexes of the table
				INSERT INTO #fraglist 
				EXEC (''DBCC SHOWCONTIG ('' + @tableidchar + '') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS'')
				FETCH NEXT FROM tables INTO @tableidchar
			END
		
			-- Close and deallocate the cursor
			CLOSE tables
			DEALLOCATE tables
		
			-- Begin Stage 2: (defrag) declare cursor for list of indexes to be defragged
			DECLARE indexes CURSOR FOR
			SELECT ObjectName, ObjectOwner = user_name(so.uid), ObjectId, IndexName, ScanDensity
			FROM #fraglist f
			JOIN sysobjects so ON f.ObjectId=so.id
			WHERE ScanDensity <= @maxfrag
				AND INDEXPROPERTY (ObjectId, IndexName, ''IndexDepth'') > 0
		
			-- Open the cursor
			OPEN indexes
		
			-- Loop through the indexes
			FETCH NEXT
			FROM indexes
			INTO @tablename, @objectowner, @objectid, @indexname, @frag
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET QUOTED_IDENTIFIER ON
		
				IF (@sSQLVersion = ''8'')
					SET @execstr = ''DBCC DBREINDEX ('''''' + RTRIM(@objectowner) + ''.'' + RTRIM(@tablename)  + '''''' , '' + RTRIM(@indexname) +'')''
				ELSE
					SET @execstr = ''ALTER INDEX '' +  RTRIM(@indexname) + '' ON '' + RTRIM(@objectowner) + ''.'' + RTRIM(@tablename) + '' REBUILD''
		
				EXEC (@execstr)
		
				SET QUOTED_IDENTIFIER OFF
		
				FETCH NEXT FROM indexes INTO @tablename, @objectowner, @objectid, @indexname, @frag
			END
		
			-- Close and deallocate the cursor
			CLOSE indexes
			DEALLOCATE indexes
		
		
			DROP TABLE #FragList
		END'

	EXECUTE (@sSPCode_0)


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRUpdateStatistics]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRUpdateStatistics]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRUpdateStatistics]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)


	SET @sSPCode_0 = 'ALTER PROCEDURE spASRUpdateStatistics
		AS
			BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @sTableName nvarchar(4000)
			DECLARE @sVarCommand nvarchar(4000)
			DECLARE @sSQLVersion nvarchar(20)
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			-- Checking fragmentation
			DECLARE tables CURSOR FOR
				SELECT so.[Name]
				FROM sysobjects so
				JOIN sysindexes si ON so.id = si.id
				WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
				ORDER BY so.[Name]
		
			-- Open the cursor
			OPEN tables
		
			-- Loop through all the tables in the database running dbcc showcontig on each one
			FETCH NEXT FROM tables INTO @sTableName
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET @sVarCommand = ''UPDATE STATISTICS '' + @sTableName + '' WITH FULLSCAN''
				EXEC (@sVarCommand)
				FETCH NEXT FROM tables INTO @sTableName
			END
		
			-- Close and deallocate the cursor
			CLOSE tables
			DEALLOCATE tables
		
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 14 of 38 - Modifying Functions table'

	/* ASRSysFunctions - Add new [IncludeExprTypes] column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysFunctions')
	and name = 'IncludeExprTypes'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysFunctions ADD 
						IncludeExprTypes [varchar](50) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 74'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime, ShortcutKeys, UDF, excludeExprTypes, includeExprTypes)
			VALUES (74, ''Does Record Exist'', 3, 0, ''Comparison'', '''', 0, 0, NULL, 0, NULL, ''21 22 23'')'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 38 - Modifying Expression tables'

	/* ASRSysExprComponents - Add new WorkflowElementProperty column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExprComponents')
	and name = 'WorkflowElementProperty'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExprComponents ADD 
						WorkflowElementProperty [int] NULL'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 16 of 38 - Adding Trigger Table'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysTrigger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysTrigger](
			[TableID] [int] NULL,
			[RecordID] [int] NULL,
			[SPID] [int] NULL,
			[TimeStamp] [int] NULL
			) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 17 of 38 - Amending Audit Procedure'


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAudit]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAudit]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRAudit] 
		(
			@piColumnID int,
			@piRecordID int,
			@psRecordDesc varchar(255),
			@psOldValue varchar(255),
			@psNewValue varchar(255)
		)
		AS
		BEGIN	
			DECLARE @sTableName varchar(8000),
				    @sColumnName varchar(8000),
				    @sUserName varchar(8000)
		
			/* Get the table name for the given column. */
			/* Get the column name for the given column. */
			SELECT @sTableName = ASRSysTables.tableName,
				@sColumnName = ASRSysColumns.columnName
			FROM ASRSysColumns
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
			WHERE ASRSysColumns.columnID = @piColumnID
		
			IF @sTableName IS NULL SET @sTableName = ''<Unknown>''

			SET @sUsername = user
			IF UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW''
				SET @sUsername = ''HR Pro Workflow''
			ELSE
			BEGIN
				IF EXISTS(SELECT * FROM ASRSysSystemSettings
                          WHERE [Section] = ''database''
                            AND [SettingKey] = ''updatingdatedependantcolumns''
                            AND [SettingValue] = 1)
				BEGIN
				  IF user = ''dbo''
					SET @sUsername = ''HR Pro Overnight Process''
				END
			END			


			/* Insert a record into the Audit Trail table. */
			INSERT INTO ASRSysAuditTrail 
				(userName, 
				dateTimeStamp, 
				tablename, 
				recordID, 
				recordDesc, 
				columnname, 
				oldValue, 
				newValue,
				ColumnID, 
				deleted)
			VALUES 
				(@sUsername, 
				getDate(), 
				@sTableName, 
				@piRecordID, 
				@psRecordDesc, 
				@sColumnName, 
				@psOldValue, 
				@psNewValue,
				@piColumnID,
				0)
		END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 18 of 38 - Modifying column for Calendar Report Events'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysCalendarReportEvents')
		and name = 'LegendCharacter'
		and length < 2

	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [ASRSysCalendarReportEvents]
                               ALTER COLUMN [LegendCharacter] [varchar](2) NULL '
		EXEC sp_executesql @NVarCommand
	END


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 19 of 38 - Modifying ASR System tables for Multiple SSI table functionality'

	/* ASRSysSSIViews - Add new TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIViews')
	and name = 'TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIViews ADD 
						TableID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE [ASRSysSSIViews]
								SET [ASRSysSSIViews].[TableID] = (SELECT [ASRSysViews].[ViewTableID] 
																  FROM [ASRSysViews] 
																  WHERE [ASRSysViews].[ViewID] = [ASRSysSSIViews].[ViewID])'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysSSIntranetLinks - Add new TableID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
	and name = 'TableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
						TableID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE [ASRSysSSIntranetLinks]
								SET [ASRSysSSIntranetLinks].[TableID] = (SELECT [ASRSysViews].[ViewTableID] 
																		  FROM [ASRSysViews] 
																		  WHERE [ASRSysViews].[ViewID] = [ASRSysSSIntranetLinks].[ViewID])'
		EXEC sp_executesql @NVarCommand
	END
		

/* ---------------------------------------------------------------------------------- */
PRINT 'Step 20 of 38 - Updating Export Definitions'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExportName')
	and name = 'Delimiter'
	if @iRecCount = 1
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExportName ALTER COLUMN Delimiter [varchar](7) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExportName')
	and name = 'OtherDelimiter'
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExportName ADD OtherDelimiter [varchar](1) NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysExportName
								SET OtherDelimiter = Delimiter, Delimiter = ''<Other>''
								WHERE Delimiter <> ''<tab>'' AND Delimiter <> '','' '
		EXEC sp_executesql @NVarCommand

	END


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 21 of 38 - Alter Currency Conversion function'

IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertCurrency]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALter PROCEDURE sp_ASRFn_ConvertCurrency
		(
			  @pfResult	Float OUTPUT
			, @pfValue	Float
			, @psFromCurr	VarChar(8000)
			, @psToCurr	VarChar(8000)
		)
		AS
		BEGIN
			DECLARE
					  @sCConvTable 	SysName
					, @sCConvExRateCol	SysName
					, @sCConvCurrDescCol	SysName
					, @sCConvDecCol	SysName
					, @sCommandString	nvarchar(4000)
					, @sParamDefinition	nvarchar(500)
			
			-- Get the name of the Currency Conversion table and Currency Description column.
			SELECT @sCConvCurrDescCol = ASRSysColumns.ColumnName, @sCConvTable = ASRSysTables.TableName 
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
		             			 INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID 
			WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_CurrencyNameColumn''
		
			-- Get the name of the Exchange Rate column.
			SELECT @sCConvExRateCol = ASRSysColumns.ColumnName
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
				WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_ConversionValueColumn''
		
			-- Get the name of the Decimals column.
			SELECT @sCConvDecCol = ASRSysColumns.ColumnName
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
				WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_DecimalColumn''
		
			IF (NOT @sCConvTable IS NULL) AND (NOT @sCConvCurrDescCol IS NULL) AND (NOT @sCConvExRateCol IS NULL) AND (NOT @sCConvDecCol IS NULL) 
		  -- Create the SQL string that returns the Coverted value.
		  BEGIN
				SET @sCommandString = ''SELECT @pfResult = ROUND(ISNULL(('' + LTRIM(RTRIM(STR(@pfValue,20,6))) 
									+ '' / NULLIF((SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol
									 		  + '' FROM '' + @sCConvTable
											  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psFromCurr + ''''''), 0))''
											  + '' * ''
											  + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol
											  + '' FROM '' + @sCConvTable
											  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + ''''''), 0)''
											  + '' , ''
											  + '' ISNULL(''
											  + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvDecCol
											  + '' FROM '' + @sCConvTable
											  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + ''''''), 0))''
		
				SET @sParamDefinition = N''@pfResult float output''
		
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfResult output
			END
			ELSE
				SET @pfResult = NULL
		
		END'

	EXECUTE (@sSPCode_0)



/* ---------------------------------------------------------------------------------- */
PRINT 'Step 22 of 38 - Update Import Definitions'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysImportName')
	and name = 'HeaderLines'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysImportName ADD
							   [HeaderLines] [Int] NULL,
							   [FooterLines] [Int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysImportName SET
							   [HeaderLines] = convert(int,[IgnoreFirstLine]),
							   [FooterLines] = convert(int,[IgnoreLastLine])'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE ASRSysImportName DROP COLUMN
							   [IgnoreFirstLine],
							   [IgnoreLastLine]'
		EXEC sp_executesql @NVarCommand
	END


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 23 of 38 - Read permissions boost'

	----------------------------------------------------------------------
	-- sp_ASRAllTablePermissions
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAllTablePermissions]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAllTablePermissions]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissions]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRAllTablePermissions] 
	(
	@psSQLLogin 		varchar(200)
	)
	AS
	BEGIN

		SET NOCOUNT ON

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

		-- Cached cut down view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
		INSERT @SysProtects
			SELECT [ID],[Action],[ProtectType], [Columns] FROM sysprotects
			WHERE [UID] = @iUserGroupID

		-- Cached version of the Base table IDs
		DECLARE @BaseTableIDs TABLE([ID] int PRIMARY KEY CLUSTERED, [BaseTableID] int)
		INSERT @BaseTableIDs
			SELECT DISTINCT o.ID, v.TableID
			FROM sysobjects o
			INNER JOIN ASRSysChildViews2 v ON v.ChildViewID = CONVERT(integer,SUBSTRING(o.Name,9,PATINDEX ( ''%#%'' , o.Name) - 9))
			WHERE Name LIKE ''ASRSYSCV%''


		SELECT o.name, p.action, bt.BaseTableID
		FROM @SysProtects p
		INNER JOIN sysobjects o ON p.id = o.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE p.protectType <> 206
			AND p.action <> 193
			AND (o.xtype = ''u'' or o.xtype = ''v'')
			AND (o.Name NOT LIKE ''ASRSYS%'' OR o.Name LIKE ''ASRSYSCV%'')
		UNION
		SELECT o.name, 193, bt.BaseTableID
		FROM syscolumns
		INNER JOIN @SysProtects p ON (syscolumns.id = p.id
			AND p.action = 193 
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
		INNER JOIN sysobjects o ON o.id = p.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE syscolumns.name = ''timestamp''
			AND p.protectType IN (204, 205) 
			AND (o.Name NOT LIKE ''ASRSYS%'' OR o.Name LIKE ''ASRSYSCV%'')
		ORDER BY o.name

	END'

	EXECUTE (@sSPCode_0)



	----------------------------------------------------------------------
	-- sp_ASRAllTablePermissionsForGroup
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAllTablePermissionsForGroup]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].sp_ASRAllTablePermissionsForGroup
		(
			@psGroupName sysname
		)
		AS
		BEGIN
			/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
			DECLARE @iUserGroupID	int

			/* Initialise local variables. */
			SELECT @iUserGroupID = sysusers.gid
			FROM sysusers
			WHERE sysusers.name = @psGroupName

			SELECT sysobjects.name, sysprotects.action
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
				AND (sysobjects.Name NOT LIKE ''ASRSYS%'' OR sysobjects.Name LIKE ''ASRSYSCV%'')
			ORDER BY sysobjects.name
		END'
	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRGetAllTableAndViewColumns
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetAllTableAndViewColumns]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetAllTableAndViewColumns]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetAllTableAndViewColumns]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetAllTableAndViewColumns] 
		AS
		BEGIN
		
			SELECT ASRSysColumns.columnName, ASRSysColumns.columnType, ASRSysColumns.dataType
			, ASRSysColumns.columnID, ASRSysColumns.uniqueCheckType, ASRSysColumns.DefaultDisplayWidth
			, ASRSysColumns.Size, ASRSysColumns.Decimals, ASRSysColumns.Use1000Separator
			, ASRSysColumns.OLEType, ASRSysTables.tableName AS tableViewName 
			FROM ASRSysColumns 
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID 
			UNION SELECT ASRSysColumns.columnName, ASRSysColumns.columnType, ASRSysColumns.dataType
			, ASRSysColumns.columnID, ASRSysColumns.uniqueCheckType, ASRSysColumns.DefaultDisplayWidth
			, ASRSysColumns.Size, ASRSysColumns.Decimals, ASRSysColumns.Use1000Separator
			, ASRSysColumns.OLEType, ASRSysViews.viewName AS tableViewName 
			FROM ASRSysColumns 
			INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysViews.viewTableID 
			LEFT OUTER JOIN ASRSysViewColumns ON (ASRSysViews.viewID = ASRSysViewColumns.viewID 
				AND ASRSysColumns.columnID = ASRSysViewColumns.columnID) 
			WHERE ASRSysViewColumns.inView = 1 OR ASRSysColumns.columnType = 3
		
		END'
	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRGetAllTableAndViewColumnPermissionsForGroup
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetAllTableAndViewColumnPermissionsForGroup]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetAllTableAndViewColumnPermissionsForGroup]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetAllTableAndViewColumnPermissionsForGroup]
			AS
			BEGIN
				DECLARE @iDummy Int
			END'
		EXECUTE (@sSPCode_0)

		SET @sSPCode_0 = 'ALTER PROCEDURE [spASRGetAllTableAndViewColumnPermissionsForGroup](
		@piUID int)
	AS
	BEGIN

		SET NOCOUNT ON

		-- Cached cut down view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
		INSERT @SysProtects
			SELECT [ID], [Action], [ProtectType], [Columns] FROM sysprotects
			WHERE [UID] = @piUID	AND [Action] IN (193, 197)
				AND [ProtectType] = 205

		DECLARE @Phase1 TABLE([TableViewName] sysname, [Name] sysname, [Select] smallint, [Edit] smallint)
			INSERT @Phase1
			SELECT o.name, c.name
				,CASE [Action] WHEN 193 THEN 1 ELSE 0 END
				,CASE [Action] WHEN 197 THEN 1 ELSE 0 END
			FROM @sysprotects p
			INNER JOIN sysobjects o ON p.id = o.id 
			INNER JOIN syscolumns c ON p.id = c.id 
			WHERE c.name <> ''timestamp''
				AND (o.Name NOT LIKE ''ASRSYS%'' OR o.Name LIKE ''ASRSYSCV%'')
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0 
				AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
				 OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0 
				AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0))

		SELECT [TableViewName], [Name]
			, SUM([Select]) AS [Select]
			, SUM([Edit]) AS [Edit]
		FROM @Phase1		
		GROUP BY [TableViewName], [Name]
		ORDER BY [TableViewName], [Name]

	END'
	EXECUTE (@sSPCode_0)


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 24 of 38 - 64 bit DLL update'

	----------------------------------------------------------------------
	-- spASRMakeLoginsProcessAdmin
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRMakeLoginsProcessAdmin]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRMakeLoginsProcessAdmin]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRMakeLoginsProcessAdmin]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRMakeLoginsProcessAdmin
		AS
		BEGIN
		
			SET NOCOUNT OFF
		
			DECLARE @cursLogins cursor
			DECLARE @Mode smallint
			DECLARE @sName nvarchar(200)
			DECLARE @tmp_role_member_ids TABLE(id int not null, role_id int null, sub_role_id int null, generation int null)
			DECLARE @generation int
		
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
			IF @@ROWCOUNT = 0 SET @Mode = 0
		
			SET @generation = 0
		
			INSERT INTO @tmp_role_member_ids (id)
				SELECT CAST(rl.uid AS int) AS [ID]
				FROM dbo.sysusers AS rl
				WHERE (rl.issqlrole = 1)and(rl.name=N''ASRSysGroup'')
		
			UPDATE @tmp_role_member_ids SET role_id = id, sub_role_id = id, generation=@generation
			WHILE ( 1=1 )
			BEGIN
				INSERT INTO @tmp_role_member_ids (id, role_id, sub_role_id, generation)
					SELECT a.memberuid, b.role_id, a.groupuid, @generation + 1
						FROM sysmembers AS a INNER JOIN @tmp_role_member_ids AS b
						ON a.groupuid = b.id
						WHERE b.generation = @generation
				IF @@ROWCOUNT <= 0
					BREAK
				SET @generation = @generation + 1
			END
		
			DELETE @tmp_role_member_ids
			WHERE id in ( SELECT CAST(rl.uid AS int) AS [ID]
				FROM dbo.sysusers AS rl
				WHERE (rl.issqlrole = 1)and(rl.name=N''ASRSysGroup''))
		
			UPDATE @tmp_role_member_ids SET generation = 0;
		
			INSERT INTO @tmp_role_member_ids (id, role_id, generation) 
				SELECT distinct id, role_id, 1 FROM @tmp_role_member_ids
		
			DELETE @tmp_role_member_ids WHERE generation = 0
		
			SET @cursLogins = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
				SELECT u.name
					FROM dbo.sysusers AS rl
					INNER JOIN @tmp_role_member_ids AS m ON m.role_id=CAST(rl.uid AS int)
					INNER JOIN dbo.sysusers AS u ON u.uid = m.id
					WHERE (rl.issqlrole = 1)and(rl.name=N''ASRSysGroup'')
		    OPEN @cursLogins
		    FETCH NEXT FROM @cursLogins INTO @sName
		
		    WHILE (@@fetch_status = 0)
		    BEGIN
				FETCH NEXT FROM @cursLogins INTO @sName
		
				PRINT ''--'' + @sName
		
				IF (@Mode = 3)
				BEGIN
					EXEC master..sp_addsrvrolemember @loginame = @sName, @rolename = N''processadmin''
				END
				ELSE
				BEGIN
					EXEC master..sp_dropsrvrolemember @loginame = @sName, @rolename = N''processadmin''
				END
		
			END
			
		END'
	EXECUTE (@sSPCode_0)


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 25 of 38 - Updating Screen Control Properties'


	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysControls')
	and name = 'ReadOnly'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysControls ADD [ReadOnly] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysControls SET [ReadOnly] = 0'
		EXEC sp_executesql @NVarCommand
	END


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 26 of 38 - Updating Screen Control Procedure'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetControlDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetControlDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRGetControlDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].sp_ASRGetControlDetails 
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
						ASRSysColumns.uniquechecktype,
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
						case when isnull(ASRSysColumns.readOnly,0) = 1 then 1 else case when isnull(ASRSysControls.readOnly,0) = 1 then 1 else 0 end end as ''readOnly'', 
						ASRSysColumns.statusBarMessage, 
						ASRSysColumns.errorMessage, 
						ASRSysColumns.linkTableID,
						ASRSysColumns.linkViewID,
						ASRSysColumns.linkOrderID,
						ASRSysColumns.Afdenabled,
						ASRSysTables.TableName,
						ASRSysColumns.Trimming,
						ASRSysColumns.Use1000Separator,
						ASRSysColumns.QAddressEnabled,
						ASRSysColumns.OLEType,
						ASRSysColumns.MaxOLESizeEnabled,
						ASRSysColumns.MaxOLESize,
						ASRSysColumns.AutoUpdateLookupValues			
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
				END
		'

	EXECUTE (@sSPCode_0)

/* ---------------------------------------------------------------------------------- */
PRINT 'Step 27 of 38 - Updating Audit permissions'

	DECLARE @sGroup sysname
	DECLARE @secRW bit
	DECLARE @secRO bit
	DECLARE @cmg bit

	DECLARE curGroups CURSOR LOCAL FAST_FORWARD FOR 
	SELECT name, secRW.permitted, secRO.permitted, cmg.permitted
	FROM sysusers u
	INNER JOIN ASRSysGroupPermissions secRW
            ON u.name = secRW.groupName AND secRW.itemid = 3	-- 'SECURITYMANAGER'
	INNER JOIN ASRSysGroupPermissions secRO
            ON u.name = secRO.groupName AND secRO.itemid = 84	-- 'SECURITYMANAGERRO'
	INNER JOIN ASRSysGroupPermissions cmg
            ON u.name = cmg.groupName AND cmg.itemid = 86		-- 'CMGRUN'

	OPEN curGroups
	FETCH NEXT FROM curGroups INTO @sGroup, @secRW, @secRO, @cmg
	WHILE (@@fetch_status = 0)
	BEGIN

		IF (@secRW = 1)
		BEGIN
			EXEC('GRANT DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditAccess TO [' + @sGroup + ']')
			EXEC('GRANT DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditCleardown TO [' + @sGroup + ']')
			EXEC('GRANT DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditGroup TO [' + @sGroup + ']')
			EXEC('GRANT DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditPermissions TO [' + @sGroup + ']')
			EXEC('GRANT DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditTrail TO [' + @sGroup + ']')
		END
		ELSE
		BEGIN
			EXEC('REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditAccess TO [' + @sGroup + ']')
			EXEC('REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditCleardown TO [' + @sGroup + ']')
			EXEC('REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditGroup TO [' + @sGroup + ']')
			EXEC('REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditPermissions TO [' + @sGroup + ']')
			EXEC('REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditTrail TO [' + @sGroup + ']')
			IF (@secRO = 1)
			BEGIN
				EXEC('GRANT SELECT ON ASRSysAuditAccess TO [' + @sGroup + ']')
				EXEC('GRANT SELECT ON ASRSysAuditCleardown TO [' + @sGroup + ']')
				EXEC('GRANT SELECT ON ASRSysAuditGroup TO [' + @sGroup + ']')
				EXEC('GRANT SELECT ON ASRSysAuditPermissions TO [' + @sGroup + ']')
				EXEC('GRANT SELECT ON ASRSysAuditTrail TO [' + @sGroup + ']')
				EXEC('GRANT UPDATE ON ASRSysAuditTrail(CMGCommitDate, CMGExportDate) TO [' + @sGroup + ']')
			END
			ELSE IF (@cmg = 1)
			BEGIN
				EXEC('GRANT SELECT ON ASRSysAuditTrail TO [' + @sGroup + ']')
				EXEC('GRANT UPDATE ON ASRSysAuditTrail(CMGCommitDate, CMGExportDate) TO [' + @sGroup + ']')
			END
		END

		FETCH NEXT FROM curGroups INTO @sGroup, @secRW, @secRO, @cmg
	END

	CLOSE curGroups
	DEALLOCATE curGroups

	REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditTrail TO ASRSysGroup
	REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditCleardown TO ASRSysGroup
	REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditAccess TO ASRSysGroup
	REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditGroup TO ASRSysGroup
	REVOKE DELETE, INSERT, SELECT, UPDATE ON ASRSysAuditPermissions TO ASRSysGroup

	GRANT SELECT(DateTimeStamp, RecordID, ColumnID) ON ASRSysAuditTrail TO ASRSysGroup
	GRANT INSERT ON ASRSysAuditAccess TO ASRSysGroup


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 28 of 38 - Remove obsolete trigger'

	-- This trigger may exist if hotfix Q000431 has been run. However, it is no longer required in v3.5
	IF EXISTS (SELECT [Name] FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[DEL_ASRSysLock]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP TRIGGER [dbo].[DEL_ASRSysLock]


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 29 of 38 - Remove obsolete functions'

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRLogins]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRLogins]

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRServerDir]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRServerDir]

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRServerFileExists]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRServerFileExists]

	IF EXISTS (SELECT [Name] FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRGenerateSysProcesses]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGenerateSysProcesses]

	IF EXISTS (SELECT [Name] FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRInsertToTableFromText]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInsertToTableFromText]

	IF EXISTS (SELECT [Name] FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRIntModuleEnabled]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIntModuleEnabled]

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASROutlookBatch]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASROutlookBatch]

	IF EXISTS (SELECT [Name] FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRTestProcessAccount]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRTestProcessAccount]


/* ---------------------------------------------------------------------------------- */
PRINT 'Step 30 of 38 - Updating Outlook Processing'

	----------------------------------------------------------------------
	-- spASRNetOutlookBatch
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRNetOutlookBatch]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRNetOutlookBatch]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRNetOutlookBatch]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRNetOutlookBatch]
		(
			@Content varchar(8000) out,
			@AllDayEvent bit out,
			@StartDate datetime out,
			@EndDate datetime out,
			@StartTime varchar(8000) out,
			@EndTime varchar(8000) out,
			@Subject varchar(8000) out,
			@Folder varchar(8000) out,
			@LinkID int,
			@RecordID int,
			@FolderID int,
			@StartDateColumnID int,
			@EndDateColumnID int,
			@FixedStartTime varchar(8000),
			@FixedEndTime varchar(8000),
			@StartTimeColumnID int,
			@EndTimeColumnID int,
			@TimeRange int,
			@Title varchar(8000),
			@SubjectExprID int,
			@RecordDescExprID int,
			@DateFormat varchar(8000),
			@FolderPath varchar(8000),
			@FolderType int,
			@FolderExprID int
		)
			AS
		
				BEGIN
		
				DECLARE @sSQL nvarchar(4000)
				DECLARE @sParamDefinition nvarchar(4000)
		
				DECLARE @CharValue varchar(8000)
		
				DECLARE @Heading varchar(8000)
				DECLARE @TableName varchar(8000)
				DECLARE @ColumnName varchar(8000)
				DECLARE @DataType int
					
				SELECT @sSQL = ''SELECT @StartDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
				FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
				WHERE ColumnID = @StartDateColumnID
				SET @sParamDefinition = N''@StartDate datetime OUTPUT''
				EXEC sp_executesql @sSQL,  @sParamDefinition, @StartDate OUTPUT
		
				SET @EndDate = Null
				IF @EndDateColumnID > 0
				BEGIN
					SELECT @sSQL = ''SELECT @EndDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
					WHERE ColumnID = @EndDateColumnID
					SET @sParamDefinition = N''@EndDate datetime OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @EndDate OUTPUT
					IF rtrim(@EndDate) = '''' SET @EndDate = null
				END
		
				IF @TimeRange = 0
				BEGIN
					SET @AllDayEvent = 1
					SET @StartTime = ''''
					SET @EndTime = ''''
				END
				IF @TimeRange = 1
				BEGIN
					SET @AllDayEvent = 0
					SET @StartTime = @FixedStartTime
					SET @EndTime = @FixedEndTime
				END
				IF @TimeRange = 2
				BEGIN
					SET @AllDayEvent = 0
		
					SELECT @sSQL = ''SELECT @StartTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
					WHERE ColumnID = @StartTimeColumnID
					SET @sParamDefinition = N''@StartTime varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @StartTime OUTPUT
		
					SELECT @sSQL = ''SELECT @EndTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
					WHERE ColumnID = @EndTimeColumnID
					SET @sParamDefinition = N''@EndTime varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @EndTime OUTPUT
		
					IF UPPER(@StartTime) = ''AM''
						SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
						WHERE [Section] = ''outlook'' and [Settingkey] = ''amstarttime''
					IF UPPER(@StartTime) = ''PM''
						SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
						WHERE [Section] = ''outlook'' and [Settingkey] = ''pmstarttime''
					IF UPPER(@EndTime) = ''AM''
						SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
						WHERE [Section] = ''outlook'' and [Settingkey] = ''amendtime''
					IF UPPER(@EndTime) = ''PM''
						SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
						WHERE [Section] = ''outlook'' and [Settingkey] = ''pmendtime''
				END
		
		
				SET @Subject = ''''
				IF @SubjectExprID > 0
				BEGIN
					SET @sSQL = ''DECLARE @hResult int
						IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),'


	SET @sSPCode_1 = '@SubjectExprID)+'''''')
					             BEGIN
					                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@SubjectExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
					                IF @hResult <> 0 SET @Subject = ''''''''
					                SET @Subject = CONVERT(varchar(255), @Subject)
						     END
						     ELSE SET @Subject = ''''''''''
					SET @sParamDefinition = N''@Subject varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
				END
				ELSE
				BEGIN
					IF @RecordDescExprID > 0
					BEGIN
						SET @sSQL = ''DECLARE @hResult int
							IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),@RecordDescExprID)+'''''')
						             BEGIN
						                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@RecordDescExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
						                IF @hResult <> 0 SET @Subject = ''''''''
						                SET @Subject = CONVERT(varchar(255), @Subject)
							     END
							     ELSE SET @Subject = ''''''''''
						SET @sParamDefinition = N''@Subject varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
						IF @Subject <> ''''
							SET @Subject = '': ''+@Subject
					END
					SET @Subject = @Title+@Subject
				END
		
		
				SET @Folder = @FolderPath
				IF @FolderType > 0
				BEGIN
					SET @sSQL = ''DECLARE @hResult int
						IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),@FolderExprID)+'''''')
					             BEGIN
					                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@FolderExprID)+'' @Folder OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
					                IF @hResult <> 0 SET @Folder = ''''''''
						     END
						     ELSE SET @Folder = ''''''''''
					SET @sParamDefinition = N''@Folder varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @Folder OUTPUT
				END
		
		
				DECLARE CursorColumns CURSOR FOR 
				SELECT isnull(ASRSysOutlookLinksColumns.Heading,''''),
					ASRSysTables.TableName,
					ASRSysColumns.ColumnName,
					ASRSysColumns.DataType
				FROM ASRSysOutlookLinksColumns
				JOIN ASRSysColumns
					ON ASRSysColumns.ColumnID = ASRSysOutlookLinksColumns.ColumnID
				JOIN ASRSysTables
					ON ASRSysColumns.TableID = ASRSysTables.TableID
				WHERE LinkID = @LinkID
				ORDER BY [Sequence] DESC
		
				SET @Content = char(13) + @Content
		
				OPEN CursorColumns
				FETCH NEXT FROM CursorColumns
				INTO	@Heading, @TableName, @ColumnName, @DataType
		
				WHILE @@FETCH_STATUS = 0
				BEGIN
		
					IF @Heading <> '''' SET @Heading = @Heading+'': ''
		
					IF @DataType = 12
						SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+isnull([''+@ColumnName+''],'''''''') FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					IF @DataType = 11
						SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] is null then ''''<Empty>'''' else convert(varchar(8000),[''+@ColumnName+''],''+@DateFormat+'') end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					IF @DataType = -7
						SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] = 1 then ''''Y'''' else ''''N'''' end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					IF @DataType <> 11 AND @DataType <> 12 AND @DataType <> -7
						SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+convert(varchar(8000),isnull([''+@ColumnName+''],'''''''')) FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
		
					SET @sParamDefinition = N''@CharValue varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @CharValue OUTPUT
		
					IF @CharValue IS Null SET @Cha'


	SET @sSPCode_2 = 'rValue = ''''
					SET @Content = @CharValue + char(13) + @Content
		
					FETCH NEXT FROM CursorColumns
					INTO	@Heading, @TableName, @ColumnName, @DataType
				END
		
				CLOSE CursorColumns
				DEALLOCATE CursorColumns
		
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)


	----------------------------------------------------------------------
	-- sp_ASRSendMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRSendMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRSendMessage]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRSendMessage]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE sp_ASRSendMessage 
		(
			@psMessage	varchar(8000),
			@psSPIDS	varchar(8000)
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
					AND program_name NOT LIKE ''HR Pro Workflow%''
					AND program_name NOT LIKE ''HR Pro Outlook%''
					AND program_name NOT LIKE ''HR Pro Server.Net%''
				AND dbid = @iDBid
				AND (spid <> @@spid and spid <> @Realspid)
				AND (@psSPIDS = '''' OR charindex('' ''+convert(varchar,spid)+'' '', @psSPIDS)>0)

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

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 31 of 38 - Updating domain integration'

	----------------------------------------------------------------------
	-- spASRGetDomainPolicy
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDomainPolicy]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDomainPolicy]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetDomainPolicy]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRGetDomainPolicy
			(@LockoutDuration int OUTPUT,
			 @lockoutThreshold int OUTPUT,
			 @lockoutObservationWindow int OUTPUT,
			 @maxPwdAge int OUTPUT, 
			 @minPwdAge int OUTPUT,
			 @minPwdLength int OUTPUT, 
			 @pwdHistoryLength int OUTPUT, 
			 @pwdProperties int OUTPUT)
		AS
		BEGIN
		
			SET NOCOUNT ON
			
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 int
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion char(2)
		
			-- Initialise the variables
			SET @LockoutDuration = 0
			SET @lockoutThreshold  = 0
			SET @lockoutObservationWindow  = 0
			SET @maxPwdAge  = 0
			SET @minPwdAge  = 0
			SET @minPwdLength  = 0
			SET @pwdHistoryLength  = 0 
			SET @pwdProperties  = 0
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			-- SQL2000 uses Server DLL, SQL2005 uses server assembly
			IF @sSQLVersion = 8
			BEGIN
		
				/* Create Server DLL object */
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll not found''
				  RAISERROR (@pserrormessage,1,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 1
				END
		
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''getDomainPolicySettings'',@hResult2 OUTPUT, @LockoutDuration OUTPUT
						, @lockoutThreshold OUTPUT, @lockoutObservationWindow OUTPUT, @maxPwdAge OUTPUT
						, @minPwdAge OUTPUT, @minPwdLength OUTPUT, @pwdHistoryLength OUTPUT
						, @pwdProperties OUTPUT
		
				IF @hResult <> 0 
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
				  RAISERROR (@pserrormessage,2,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 2
				END
		
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
		
				EXEC sp_executesql N''EXEC spASRGetDomainPolicyFromAssembly
						@lockoutDuration OUTPUT, @lockoutThreshold OUTPUT,
						@lockoutObservationWindow OUTPUT, @maxPwdAge OUTPUT,
						@minPwdAge OUTPUT, @minPwdLength OUTPUT,
						@pwdHistoryLength OUTPUT, @pwdProperties OUTPUT''
					, N''@lockoutDuration int OUT, @lockoutThreshold int OUT,
						@lockoutObservationWindow int OUT, @maxPwdAge int OUT,
						@minPwdAge int OUT,	@minPwdLength int OUT,
						@pwdHistoryLength int OUT, @pwdProperties int OUT''
					, @LockoutDuration OUT, @lockoutThreshold OUT
					, @lockoutObservationWindow OUT, @maxPwdAge OUT
					, @minPwdAge OUT, @minPwdLength OUT
					, @pwdHistoryLength OUT, @pwdProperties OUT
		
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetDomains
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDomains]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDomains]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetDomains]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetDomains
				(@DomainString varchar(8000) OUTPUT)
		AS
		BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 varchar(255)
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion char(2)
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			-- SQL2000 uses Server DLL, SQL2005 uses server assembly
			IF @sSQLVersion = 8
			BEGIN
		
				-- Create Server DLL object
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll not found''
				  RAISERROR (@pserrormessage,1,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 1
				END
			
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''getDomains'', @hResult2 OUTPUT, @DomainString OUTPUT
			
				IF @hResult <> 0 
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
				  RAISERROR (@pserrormessage,2,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 2
				END
			
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
				SELECT @DomainString = dbo.udfASRNetGetDomains()
			END
		
		END'

	EXECUTE (@sSPCode_0)


	----------------------------------------------------------------------
	-- spASRGetWindowsUsers
	----------------------------------------------------------------------
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWindowsUsers]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWindowsUsers]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWindowsUsers]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRGetWindowsUsers
				(@DomainName varchar(200),
				@UserString varchar(8000) OUTPUT)
		AS
		BEGIN
		
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 varchar(255)
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion char(2)
		
			SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		
			-- SQL2000 uses Server DLL, SQL2005 uses server assembly
			IF @sSQLVersion = 8
			BEGIN
			
				-- Create Server DLL object
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
					EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
					SET @pserrormessage = ''HR Pro Server.dll not found''
					RAISERROR (@pserrormessage,1,1)
					EXEC sp_OADestroy @objectToken
					RETURN 1
				END
		
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''GetUsers'', @UserString OUTPUT, @DomainName
				IF @hResult <> 0 
				BEGIN
					EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
					SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
					RAISERROR (@pserrormessage,2,1)
					EXEC sp_OADestroy @objectToken
				RETURN 2
				END
		
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
				SELECT @UserString = dbo.udfASRNetGetUsers(@DomainName)
			END
		
			
		END'

	EXECUTE (@sSPCode_0)


	-- Reset the process acount to be trusted account
	IF EXISTS(SELECT [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'ProcessAccount'	AND [SettingKey] = 'Mode')
	BEGIN
		UPDATE ASRSysSystemSettings SET [SettingValue] = 1
			WHERE [Section] = 'ProcessAccount'	AND [SettingKey] = 'Mode'
	END


/* ------------------------------------------------------------- */
PRINT 'Step 32 of 38 - Bradford Factor '

	----------------------------------------------------------------------
	-- sp_ASR_Bradford_DeleteAbsences
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASR_Bradford_DeleteAbsences]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
		(
			@pdReportStart	  	datetime,
			@pdReportEnd		datetime,
			@pbOmitBeforeStart	bit,
			@pbOmitAfterEnd	bit,
			@pcReportTableName	char(30)
		)
		AS
		BEGIN
		
			declare @piID as integer
			declare @pdStartDate as datetime
			declare @pdEndDate as datetime
			declare @iDuration as float
			declare @pbDeleteThisAbsence as bit
			declare @sSQL as char(8000)
		
			set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Absence_ID, Start_Date, End_Date, Duration FROM '' + @pcReportTableName
			execute(@sSQL)
			open BradfordIndexCursor
		
			Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration
			while @@FETCH_STATUS = 0
				begin
					set @pbDeleteThisAbsence = 0
					if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1
					if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1
					if @iDuration = 0 set @pbDeleteThisAbsence = 1
		
					if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1
					if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1
		
					if @pbDeleteThisAbsence = 1
						begin
							set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = Convert(Int,'' + Convert(char(10),@piId) + '')''
							execute(@sSQL)
						end
		
					Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration
				end
		
			close BradfordIndexCursor
			deallocate BradfordIndexCursor
		
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 33 of 38 - Locking Update '

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

	SET @sSPCode_0 = 'ALTER PROCEDURE sp_ASRLockCheck AS
	BEGIN

		SET NOCOUNT ON

		DECLARE @sSQLVersion char(2)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9'' AND APP_NAME() <> ''HR Pro Workflow Service'' AND APP_NAME() <> ''HR Pro Outlook Calendar Service''
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

	END'
	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 34 of 38 - Reserved word update'

	SELECT @iRecCount = COUNT(Keyword) FROM ASRSysKeywords
		WHERE provider = 'Microsoft SQL Server' and Keyword = 'sys'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT ASRSysKeywords (Provider, Keyword) 
                               VALUES (''Microsoft SQL Server'',''sys'')'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 35 of 38 - Build Domain List setting'

	IF NOT EXISTS(SELECT SettingValue
				  FROM   ASRSysSystemSettings
				  WHERE  [Section] = 'misc'
				  AND  [SettingKey] = 'autobuilddomainlist')
		INSERT ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES ('misc','autobuilddomainlist',0)

/* ------------------------------------------------------------- */
PRINT 'Step 36 of 38 - Transfer Ownership update'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetOwnersForAllUtilities]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetOwnersForAllUtilities]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetOwnersForAllUtilities]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRGetOwnersForAllUtilities
		AS
		BEGIN

			SET NOCOUNT ON

			DECLARE @UserNames TABLE (UserName varchar(50))

			INSERT @UserNames
				SELECT Username FROM ASRSysBatchJobName
				UNION
				SELECT Username FROM ASRSysCalendarReports
				UNION
				SELECT Username FROM ASRSysCrossTab
				UNION
				SELECT Username FROM ASRSysCustomReportsName
				UNION
				SELECT Username FROM ASRSysDataTransferName
				UNION
				SELECT Username FROM ASRSysExportName
				UNION
				SELECT Username FROM ASRSysExpressions
				UNION
				SELECT Username FROM ASRSysGlobalFunctions
				UNION
				SELECT Username FROM ASRSysImportName
				UNION
				SELECT Username FROM ASRSysLabelTypes
				UNION
				SELECT Username FROM ASRSysMailMergeName
				UNION
				SELECT Username FROM ASRSysMatchReportName
				UNION
				SELECT Username FROM ASRSysPickListName
				UNION
				SELECT Username FROM ASRSysRecordProfileName

			SELECT DISTINCT UserName FROM @UserNames
			WHERE UserName <> '''' AND UserName IS NOT NULL AND UserName <> ''sa''
			ORDER BY UserName

		END'
	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 37 of 38 - spASREmailRebuild update'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASREmailRebuild]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASREmailRebuild]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASREmailRebuild]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASREmailRebuild
		AS
		BEGIN	
			/* Refresh all calculated columns in the database. */
			DECLARE @sTableName 	varchar(255),
				@iTableID		int,
				@sSQL			varchar(8000),
				@sColumnName		varchar(255)
		
			
			/* Get a cursor of the tables in the database. */
			DECLARE curTables CURSOR FOR
				SELECT tableName, tableID
				FROM ASRSysTables
			OPEN curTables
		
			DELETE FROM AsrSysEmailQueue WHERE DateSent Is Null AND [Immediate] = 0
		
			/* Loop through the tables in the database. */
			FETCH NEXT FROM curTables INTO @sTableName, @iTableID
			WHILE @@fetch_status <> -1
			BEGIN
				/* Get a cursor of the records in the current table.  */
				/* Call the diary trigger for that table and record  */
				SET @sSQL = ''DECLARE @iCurrentID	int,
								@sSQL		varchar(8000)
							
							IF EXISTS (SELECT * FROM sysobjects
							WHERE id = object_id(''''spASREmailRebuild_'' + LTrim(Str(@iTableID)) + '''''') 
								AND sysstat & 0xf = 4)
							BEGIN
								DECLARE curRecords CURSOR FOR
								SELECT id
								FROM '' + @sTableName + ''
				
								OPEN curRecords
				
								FETCH NEXT FROM curRecords INTO @iCurrentID
								WHILE @@fetch_status <> -1
								BEGIN
									PRINT ''''ID : '''' + Str(@iCurrentID)
									SET @sSQL = ''''EXEC spASREmailRebuild_'' + LTrim(Str(@iTableID)) 
										+ '' '''' + convert(varchar(100), @iCurrentID) + ''''''''
									EXEC (@sSQL)
				
									FETCH NEXT FROM curRecords INTO @iCurrentID
								END
								CLOSE curRecords
								DEALLOCATE curRecords
							END''
				 EXEC (@sSQL) 
		
				/* Move onto the next table in the database. */ 
				FETCH NEXT FROM curTables INTO @sTableName, @iTableID
			END
		
			CLOSE curTables
			DEALLOCATE curTables
		
			EXEC spASREmailImmediate ''''
		
		END'

	EXECUTE (@sSPCode_0)



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 38 of 38 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.5')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.5.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '3.5.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.5')


SELECT @NVarCommand = 
	'IF EXISTS (SELECT * FROM dbo.sysobjects
			WHERE id = object_id(N''[dbo].[sp_ASRLockCheck]'')
			AND OBJECTPROPERTY(id, N''IsProcedure'') = 1)
		GRANT EXECUTE ON sp_ASRLockCheck TO public'
EXEC sp_executesql @NVarCommand


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.5 Of HR Pro'
