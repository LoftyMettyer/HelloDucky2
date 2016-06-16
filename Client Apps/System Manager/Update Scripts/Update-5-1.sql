/* --------------------------------------------------- */
/* Update the database from version 5.0 to version 5.1 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @ownerGUID uniqueidentifier
DECLARE @newDesktopImageID	integer,
		@picname			varchar(255),
		@oldDesktopImageID	integer;

DECLARE @categorymatch TABLE(tableid integer, categoryid integer);
DECLARE @bCategoriesProcessed bit,
		@configID integer,
		@recruitmentID integer,
		@salaryID integer,
		@generalID integer,
		@disciplineID integer,
		@healthID integer,
		@trainingID integer,
		@skillID integer,
		@benefitID integer,
		@statabsenceID integer,
		@absenceID integer;
		
DECLARE @tableid	integer,
		@categoryid	integer,
		@nextid		integer;

DECLARE @sSPCode nvarchar(max)

DECLARE @admingroups TABLE(groupname nvarchar(255))


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON;
SET @DBName = DB_NAME();

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '5.0') and (@sDBVersion <> '5.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
IF (@iSQLVersion < 9)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
PRINT 'Step - Deletion changes'
/* ------------------------------------------------------------- */

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID(''tbuser_' + tablename + ''', ''U'') AND name = ''_deleted'')
		ALTER TABLE [tbuser_' + TableName + '] ADD [_deleted] bit, [_deleteddate] datetime;' FROM ASRSysTables
			ORDER BY tablename;	
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step - Fusion Message Bus Integration'
/* ------------------------------------------------------------- */

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spSendFusionMessage]') AND xtype = 'P')
		EXECUTE sp_executesql N'CREATE PROCEDURE fusion.spSendFusionMessage(@TableID integer, @RecordID integer)
		AS
		BEGIN
			SET NOCOUNT ON;
		END'



/* ------------------------------------------------------------- */
PRINT 'Step - Calculation framework'
/* ------------------------------------------------------------- */

	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbstat_componentcode', 'U') AND name = 'objectid')
		EXEC sp_executesql N'ALTER TABLE tbstat_componentcode DROP COLUMN [objectid];';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_convertcharactertonumeric]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsys_convertcharactertonumeric];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_divide]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsys_divide];

	-- Clear old component expressions
	DELETE FROM tbstat_componentcode WHERE id IN (16, 61, 25, 4, 27, 42) AND isoperator = 0
	DELETE FROM tbstat_componentcode WHERE id IN (4) AND isoperator = 1
	DELETE FROM tbstat_componentdependancy WHERE id IN (16, 61, 25, 42)

	-- Is Field Empty
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount])
		VALUES (16, 'dbo.udfstat_isfieldempty({0},{1})', 3, 'Is Field Empty', 0, 0, 1);		
	INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (16, 4, '', '', '');

	-- Is Field Populated
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount])
		VALUES (61, 'dbo.udfstat_isfieldpopulated({0},{1})', 3, 'Is Field Populated', 0, 0, 1);		
	INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (61, 4, '', '', '');

	-- Convert character to numeric
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount])
		VALUES (25, 'dbo.udfstat_convertcharactertonumeric({0})', 2, 'Convert Character to Numeric', 0, 0, 0);		

	-- If... Then... Else
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount])
		VALUES (4, 'CASE WHEN {0} THEN {1} ELSE {2} END', 0, 'If... Then... Else...', 0, 0, 1);		

	-- Parentheses
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount], [maketypesafe])
		VALUES (27, '({0})', 0, 'Parentheses', 0, 0, 0, 1);

	-- Divided By
	INSERT [dbo].[tbstat_componentcode] ([id], [precode], [code], [aftercode], [name], [isoperator], [operatortype], [casecount])
		VALUES (4, 'dbo.udfstat_divideby(', ',', ')', 'Divided by', 1, 0, 0);		

	-- Get field from database value
	INSERT [dbo].[tbstat_componentcode] ([id], [code], [datatype], [name], [isoperator], [operatortype], [casecount], [isgetfieldfromdb])
		VALUES (42, '[dbo].[udfsys_getfieldfromdatabaserecord_{3}] ({0}, {1}, {2})', 0, 'Get field from database record', 0, 0, 0, 1);		
	INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (42, 5, '', '', '');


/* ------------------------------------------------------------- */
PRINT 'Step - Object triggers'
/* ------------------------------------------------------------- */

	IF  EXISTS (SELECT * FROM sys.triggers WHERE object_id = OBJECT_ID(N'[dbo].[DEL_ASRSysTables]'))
		DROP TRIGGER [dbo].[DEL_ASRSysTables]

	IF  EXISTS (SELECT * FROM sys.triggers WHERE object_id = OBJECT_ID(N'[dbo].[DEL_ASRSysColumns]'))
		DROP TRIGGER [dbo].[DEL_ASRSysColumns]

	IF  EXISTS (SELECT * FROM sys.triggers WHERE object_id = OBJECT_ID(N'[dbo].[DEL_ASRSysWorkflows]'))
		DROP TRIGGER [dbo].[DEL_ASRSysWorkflows]

	EXECUTE sp_executeSQL N'CREATE TRIGGER [dbo].[DEL_ASRSysColumns] ON [dbo].[ASRSysColumns]
		INSTEAD OF DELETE
		AS
		BEGIN
			SET NOCOUNT ON;

			DELETE FROM [tbsys_columns] WHERE columnid IN (SELECT columnid FROM deleted);
			DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT columnid FROM deleted) AND objecttype = 2;
			
		END'

	EXECUTE sp_executeSQL N'CREATE TRIGGER [dbo].[DEL_ASRSysTables] ON [dbo].[ASRSysTables]
		INSTEAD OF DELETE
		AS
		BEGIN
			SET NOCOUNT ON;

			DELETE FROM [tbsys_tables] WHERE tableid IN (SELECT tableid FROM deleted);
			DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT tableid FROM deleted) AND objecttype = 1;

		END'

	EXECUTE sp_executeSQL N'CREATE TRIGGER [dbo].[DEL_ASRSysWorkflows] ON [dbo].[ASRSysWorkflows]
		INSTEAD OF DELETE
		AS
		BEGIN
			SET NOCOUNT ON;

			DELETE FROM [tbsys_workflows] WHERE id IN (SELECT id FROM deleted);
			DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT id FROM deleted) AND objecttype = 10;

		END'


/* ------------------------------------------------------------- */
PRINT  'Step - Data Cleansing'
/* ------------------------------------------------------------- */

	EXECUTE sp_executeSQL N'UPDATE ASRSysColumns SET lostFocusExprID = 0 WHERE (lostFocusExprID = - 1);';	
	EXECUTE sp_executeSQL N'UPDATE ASRSysColumns SET dfltValueExprID = 0 WHERE (dfltValueExprID = - 1);';
	DELETE FROM [tbsys_scriptedobjects] WHERE objecttype = 1 AND targetid NOT IN (SELECT tableid FROM [tbsys_tables])
	DELETE FROM [tbsys_scriptedobjects] WHERE objecttype = 2 AND targetid NOT IN (SELECT columnid FROM [tbsys_columns])
	DELETE FROM [tbsys_scriptedobjects] WHERE objecttype = 10 AND targetid NOT IN (SELECT id FROM [tbsys_workflows])

	SELECT @nextid = MAX([columnid]) + 1 FROM dbo.[ASRSysColumns];
	UPDATE tbsys_systemobjects SET [nextid] = @nextid + 1 WHERE [viewname] = 'ASRSysColumns'



/* ------------------------------------------------------------- */
PRINT 'Step - Structure changes'
/* ------------------------------------------------------------- */
	
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElementItems', 'U') AND name = 'LookupOrderID')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysWorkflowElementItems ADD LookupOrderID int NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElementitems', 'U') AND name = 'HotSpotIdentifier')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysWorkflowElementItems ADD HotSpotIdentifier VARCHAR(200) NULL;';
		EXEC sp_executesql N'UPDATE ASRSysWorkflowElementItems SET [HotSpotIdentifier] = '''';';
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysScreens', 'U') AND name = 'category')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysScreens ADD
				[category] nvarchar(255),
				[groupscreens] bit;';
		EXECUTE sp_executeSQL N'UPDATE ASRSysScreens SET groupscreens = 0 WHERE groupscreens IS NULL';				
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysScreens', 'U') AND name = 'description')
		EXEC sp_executesql N'ALTER TABLE ASRSysScreens ADD [description] nvarchar(MAX);';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysHistoryScreens', 'U') AND name = 'order')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysHistoryScreens ADD [order] smallint;';
		EXEC sp_executesql N'UPDATE ASRSysHistoryScreens SET [order] = 0;';
	END



/* ------------------------------------------------------------- */
PRINT 'Step - Basic object scripting engine'
/* ------------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_scriptnewcolumn]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_scriptnewcolumn];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.spstat_scriptnewcolumn (@columnid integer OUTPUT, @tableid integer, @columnname varchar(255)
		, @datatype integer, @description varchar(255), @size integer, @decimals integer, @islocked bit, @uniquekey varchar(37), @IsIDColumn bit)
	AS
	BEGIN

		DECLARE @ssql nvarchar(MAX),
				@tablename varchar(255),
				@datasyntax	varchar(255);

		DECLARE @spinnerMinimum integer,
			@spinnerMaximum integer,
			@spinnerIncrement integer,
			@audit bit,
			@duplicate bit,
			@defaultvalue varchar(max),
			@columntype integer,
			@mandatory bit,
			@uniquecheck bit,
			@convertcase smallint,
			@mask varchar(MAX),
			@lookupTableID integer,
			@lookupColumnID integer,
			@controltype integer,
			@alphaonly bit,
			@blankIfZero bit,
			@multiline bit,
			@alignment smallint,
			@calcExprID integer,
			@gotFocusExprID integer,
			@lostFocusExprID integer,
			@calcTrigger smallint,
			@readOnly bit,
			@statusBarMessage varchar(255),
			@errorMessage varchar(255),
			@linkTableID integer, 
			@Afdenabled bit, 
			@Afdindividual integer,
			@Afdforename integer, 
			@Afdsurname integer,
			@Afdinitial integer, 
			@Afdtelephone integer, 
			@Afdaddress integer,
			@Afdproperty integer, 
			@Afdstreet integer, 
			@Afdlocality integer, 
			@Afdtown integer, 
			@Afdcounty integer,
			@dfltValueExprID integer, 
			@linkOrderID integer, 
			@OleOnServer bit, 
			@childUniqueCheck bit,
			@LinkViewID integer, 
			@DefaultDisplayWidth integer, 
			@UniqueCheckType integer,
			@Trimming integer, 
			@Use1000Separator bit,
			@LookupFilterColumnID integer, 
			@LookupFilterValueID integer, 
			@QAddressEnabled integer, 
			@QAIndividual integer, 
			@QAAddress integer, 
			@QAProperty integer, 
			@QAStreet integer,
			@QALocality integer, 
			@QATown integer, 
			@QACounty integer, 
			@LookupFilterOperator integer, 
			@Embedded bit, 
			@OLEType integer, 
			@MaxOLESizeEnabled bit, 
			@MaxOLESize integer,
			@AutoUpdateLookupValues bit, 
			@CalculateIfEmpty bit;

		-- Can we safely create this column?
		IF EXISTS(SELECT [columnid] FROM dbo.[ASRSysColumns] WHERE tableid = @tableid AND columnname = @columnname)
			RETURN;

		SELECT @tablename = [tablename] FROM dbo.[ASRSysTables] WHERE tableid = @tableid;
		EXEC dbo.spASRGetNextObjectIdentitySeed ''ASRSysColumns'', @columnid OUTPUT;
		
		SET @defaultvalue = '''';		
		SET @spinnerMinimum = 0;
		SET @spinnerMaximum = 0;
		SET @spinnerIncrement = 0;
		SET @audit = 0;
		SET @duplicate = 0;
		SET @columntype = 0;
		SET @mandatory = 0;
		SET @uniquecheck = 0;
		SET @convertcase = 0;
		SET @mask = '''';
		SET @lookupTableID = 0;
		SET	@lookupColumnID = 0;
		SET	@controltype = 0;	
		SET @alphaonly = 0;
		SET @blankIfZero = 0;
		SET @multiline = 0;
		SET @alignment = 0;
		SET @calcExprID = 0;
		SET @gotFocusExprID = 0;
		SET @lostFocusExprID = 0;
		SET @calcTrigger = 0;
		SET @readOnly = 0;
		SET @statusBarMessage = '''';
		SET @errorMessage = '''';
		SET @linkTableID = 0; 
		SET @Afdenabled = 0; 
		SET @Afdindividual = 0;
		SET @Afdforename = 0; 
		SET @Afdsurname = 0;
		SET @Afdinitial = 0; 
		SET @Afdtelephone = 0; 
		SET @Afdaddress = 0;
		SET @Afdproperty = 0; 
		SET @Afdstreet = 0; 
		SET @Afdlocality = 0; 
		SET @Afdtown = 0; 
		SET @Afdcounty = 0;
		SET @dfltValueExprID = 0; 
		SET @linkOrderID = 0; 
		SET @OleOnServer = 0; 
		SET @childUniqueCheck = 0;
		SET @LinkViewID = 0; 
		SET @DefaultDisplayWidth = convert(varchar(10),@size);
		SET @UniqueCheckType = 0;
		SET @Trimming = 0;
		SET @Use1000Separator = 0;
		SET @LookupFilterColumnID = 0; 
		SET @LookupFilterValueID = 0; 
		SET @QAddressEnabled = 0; 
		SET @QAIndividual = 0; 
		SET @QAAddress = 0; 
		SET @QAProperty = 0; 
		SET @QAStreet = 0;
		SET @QALocality = 0; 
		SET @QATown = 0; 
		SET @QACounty = 0; 
		SET @LookupFilterOperator = 0; 
		SET @Embedded = 0; 
		SET @OLEType = 0; 
		SET @MaxOLESizeEnabled = 0; 
		SET @MaxOLESize = 0;
		SET @AutoUpdateLookupValues = 0; 
		SET @CalculateIfEmpty = 0;

		-- Is ID column?
		IF @IsIDColumn = 1 SET @columntype = 3;	

		-- Logic
		IF @datatype = -7
		BEGIN
			SET @datasyntax = ''bit'';
			SET @defaultvalue = ''FALSE'';
			SET @controltype = 1;
		END

		-- OLE
		IF @datatype = -4
			SET @controltype = 1;

		-- Photo
		IF @datatype = -3
			SET @controltype = 1024;

		-- Link
		IF @datatype = -2
		BEGIN
			SET @datasyntax = ''varchar(255)'';
			SET @controltype = 2048;
		END

		-- Working Pattern
		IF @datatype = -1
		BEGIN
			SET @datasyntax = ''varchar(14)'';
			SET @controltype = 4096;
		END
		
		-- Numeric
		IF @datatype = 2
		BEGIN
			SET @datasyntax = ''numeric('' + convert(varchar(5),@size) + '','' + @decimals + '')'';
			SET @defaultvalue = 0;	
			SET @controltype = 64;
			SET @DefaultDisplayWidth = convert(varchar(10),@size + @decimals);
		END

		-- Integers
		IF @datatype = 4
		BEGIN
			SET @datasyntax = ''integer'';
			SET @controltype = 64;
		END
		
		-- Date
		IF @datatype = 11
		BEGIN
			SET @datasyntax = ''datetime'';
			SET @controltype = 64;
		END

		-- Character
		IF @datatype = 12
		BEGIN
			SET @datasyntax = ''varchar('' + convert(varchar(5),@size) + '')'';
			SET @controltype = 64;
		END

		-- System objects update
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT @uniquekey, 2, @columnid, ''AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'', ''01/01/1900'',1,@islocked, GETDATE()

		-- Update base table								
		INSERT dbo.[tbsys_columns] ([columnID], [tableID], [columnType], [datatype], [defaultValue], [size], [decimals]
				, [lookupTableID], [lookupColumnID], [controltype], [spinnerMinimum], [spinnerMaximum], [spinnerIncrement], [audit]
				, [duplicate], [mandatory], [uniquecheck], [convertcase], [mask], [alphaonly], [blankIfZero], [multiline], [alignment]
				, [calcExprID], [gotFocusExprID], [lostFocusExprID], [calcTrigger], [readOnly], [statusBarMessage], [errorMessage]
				, [linkTableID], [Afdenabled], [Afdindividual], [Afdforename], [Afdsurname], [Afdinitial], [Afdtelephone], [Afdaddress]
				, [Afdproperty], [Afdstreet], [Afdlocality], [Afdtown], [Afdcounty], [dfltValueExprID], [linkOrderID], [OleOnServer]
				, [childUniqueCheck], [LinkViewID], [DefaultDisplayWidth], [ColumnName], [UniqueCheckType], [Trimming], [Use1000Separator]
				, [LookupFilterColumnID], [LookupFilterValueID], [QAddressEnabled], [QAIndividual], [QAAddress], [QAProperty], [QAStreet]
				, [QALocality], [QATown], [QACounty], [LookupFilterOperator], [Embedded], [OLEType], [MaxOLESizeEnabled], [MaxOLESize]
				, [AutoUpdateLookupValues], [CalculateIfEmpty]) 
			VALUES (@columnid, @tableid, @columntype, @datatype, @defaultvalue, @size, @decimals
				, @lookupTableID, @lookupColumnID, @controltype, @spinnerMinimum, @spinnerMaximum, @spinnerIncrement, @audit
				, @duplicate, @mandatory, @uniquecheck, @convertcase, @mask, @alphaonly, @blankIfZero, @multiline, @alignment
				, @calcExprID, @gotFocusExprID, @lostFocusExprID, @calcTrigger, @readOnly, @statusBarMessage, @errorMessage
				, @linkTableID, @Afdenabled, @Afdindividual, @Afdforename, @Afdsurname, @Afdinitial, @Afdtelephone, @Afdaddress
				, @Afdproperty, @Afdstreet, @Afdlocality, @Afdtown, @Afdcounty, @dfltValueExprID, @linkOrderID, @OleOnServer
				, @childUniqueCheck, @LinkViewID, @DefaultDisplayWidth, @ColumnName, @UniqueCheckType, @Trimming, @Use1000Separator
				, @LookupFilterColumnID, @LookupFilterValueID, @QAddressEnabled, @QAIndividual, @QAAddress, @QAProperty, @QAStreet
				, @QALocality, @QATown, @QACounty, @LookupFilterOperator, @Embedded, @OLEType, @MaxOLESizeEnabled, @MaxOLESize
				, @AutoUpdateLookupValues, @CalculateIfEmpty);

			-- Physically create this column (is regenerated by the System Manager save)
			IF @IsIDColumn = 0
			BEGIN 	
				SET @ssql = N''ALTER TABLE dbo.tbuser_'' + @tablename + '' ADD '' + @columnname + '' '' + @datasyntax;
				EXECUTE sp_executesql @ssql;
			END

		RETURN;

	END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_scriptnewtable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_scriptnewtable];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.[spstat_scriptnewtable](@tableID integer OUTPUT, @tablename varchar(255), @tabletype tinyint, @uniquekey varchar(37))
		AS
		BEGIN

			SET NOCOUNT ON;
			
			DECLARE @ssql nvarchar(MAX),
					@newtableID integer;		

			-- Can we safely create this table?
			SELECT @newtableID = ISNULL(tableID,0) FROM dbo.[tbsys_tables] WHERE [TableName] = @tablename;

			IF @newtableID > 0 RETURN;

			EXEC dbo.spASRGetNextObjectIdentitySeed ''ASRSysTables'', @newtableID OUTPUT;

			-- System objects update
			INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
				SELECT @uniquekey, 1, @newtableID, ''AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'', ''01/01/1900'',1, 1, GETDATE()

			-- System metadata
			INSERT dbo.[tbsys_tables] (TableID, TableType, TableName, DefaultOrderID, RecordDescExprID, DefaultEmailID
					, ManualSummaryColumnBreaks, AuditDelete, AuditInsert, isremoteview)
				VALUES (@newtableID, @tabletype, @tablename, 0, 0, 0, 0, 0, 0, 0)

			-- Physically create this table (is regenerated by the System Manager save)	
			SET @ssql = N''CREATE TABLE dbo.tbuser_'' + @tablename + '' ([ID] integer IDENTITY(1,1) PRIMARY KEY CLUSTERED
								, [updflag] int NULL, [_description] nvarchar(MAX) NULL, [_deleted] bit, [_deleteddate] datetime, [TimeStamp] timestamp NOT NULL);'';
			EXECUTE sp_executesql @ssql;

			-- Create a veiw on this table (is replaced by System Manager save, so no need to be precise)
			SET @ssql = N''CREATE VIEW dbo.['' + @tablename + ''] AS SELECT * FROM dbo.[tbuser_'' + @tablename + ''];'';
			EXECUTE sp_executesql @ssql;

			SET @tableID = @newtableID;

		END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_scriptnewprimaryorder]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_scriptnewprimaryorder];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.[spstat_scriptnewprimaryorder](@tableID integer, @ordername varchar(255))
		AS
		BEGIN

			SET NOCOUNT ON;
			
			DECLARE @ssql			nvarchar(MAX),
					@orderID		integer,
					@columncount	smallint,
					@newtableID		integer;		

			-- Can we safely create this order?
			IF NOT EXISTS(SELECT [tableID] FROM dbo.[ASRSysOrders] WHERE [TableID] = @tableID)
			BEGIN

				SELECT @orderID = MAX(OrderID) + 1 FROM dbo.[ASRSysOrders];
			
				-- Default order
				INSERT dbo.[ASRSysOrders] (OrderID, Name, TableID, [Type])
					VALUES (@orderID, @ordername, @tableID, 1);

				-- Order columns
				INSERT dbo.[ASRSysOrderItems] (OrderID, ColumnID, [Type], Sequence, Ascending)
					SELECT TOP 1 @orderID, [columnID], ''O'', 1,1
					FROM dbo.tbsys_columns WHERE [tableID] = @tableID AND columnname NOT LIKE ''ID%'' ORDER BY [columnID];

				-- Find window items
				INSERT dbo.[ASRSysOrderItems] (OrderID, ColumnID, [Type], Sequence, Ascending)
					SELECT TOP 1 @orderID, [columnID], ''F'', 1,1
					FROM dbo.tbsys_columns WHERE [tableID] = @tableID AND columnname NOT LIKE ''ID%'' ORDER BY [columnID];

				-- Set this as the primary order
				UPDATE dbo.tbsys_tables SET [DefaultOrderID] = @orderID WHERE [TableID] = @tableID;

			END

		END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_scriptnewprimaryscreen]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_scriptnewprimaryscreen];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.[spstat_scriptnewprimaryscreen](@tableID integer, @screenname varchar(255))
	AS
	BEGIN
	
		SET NOCOUNT ON;

		DECLARE @screenID integer;

		-- Can we safely create this order?
		IF NOT EXISTS(SELECT [tableID] FROM dbo.[ASRSysScreens] WHERE [TableID] = @tableID)
		BEGIN

			SELECT @screenID = ISNULL(MAX(ScreenID),0) + 1 FROM dbo.[ASRSysScreens];

			-- Insert base screen
			INSERT INTO [dbo].[ASRSysScreens]
				   ([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID],
					[FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline],
					[GridX], [GridY], [AlignToGrid],
					[DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic],
					[QuickEntry], [SSIntranet], [category], [groupscreens])
			 VALUES
				   (@screenID, @screenname, @tableID, 0, 1600, 7100, 0,
					''Verdana'', 8, 0, 0, 0, 0,
					40, 40, 1,
					0, ''Verdana'', 8, 0, 0,
					0, 0, 0, 0);

			-- Insert top most column (label)
			INSERT INTO [dbo].[ASRSysControls](
					[ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex],
					[TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor],
					[FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline],
					[TabIndex], [ReadOnly],	[NavigateOnSave])
				SELECT TOP 1 @screenID, 0, 2, NULL, 0, 256, 0,
						440, 240, 195, 1485, REPLACE([columnname],''_'','' '') + '':'' , -2147483633, 0, 
						''Verdana'', 8, 0, 0, 0, 0,
						22, 0, 0
					FROM dbo.tbsys_columns WHERE [tableID] = @tableID AND columnname NOT LIKE ''ID%'' ORDER BY [columnID];

			-- Insert top most column (control)
			INSERT INTO [dbo].[ASRSysControls](
					[ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex],
					[TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor],
					[FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline],
					[TabIndex], [ReadOnly],	[NavigateOnSave])
				SELECT TOP 1 @screenID, 0, 1, [tableid], [columnid], [controltype], 0,
						360, 1920, 315, 4950, [columnname], 16777215, 0, 
						''Verdana'', 8, 0, 0, 0, 0,
						21, 0, 0
					FROM dbo.tbsys_columns WHERE [tableID] = @tableID AND columnname NOT LIKE ''ID%'' ORDER BY [columnID];

		END
	END'





/* ------------------------------------------------------------- */
PRINT 'Step - Report Packs'
/* ------------------------------------------------------------- */

	SELECT @iRecCount = count(id) FROM syscolumns WHERE id = (select id from sysobjects where name = 'ASRSysBatchJobName') and name = 'OutputFormat'
	IF @iRecCount = 0
	  BEGIN
	    SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysBatchJobName] ADD
			[IsBatch] [bit] NULL,
			[OutputPreview] [bit] NULL,
			[OutputFormat] [int] NULL,
			[OutputScreen] [bit] NULL,
			[OutputPrinter] [bit] NULL,
			[OutputPrinterName] [varchar](255) NULL,
			[OutputSave] [bit] NULL,
			[OutputSaveExisting] [int] NULL,
			[OutputEmail] [bit] NULL,
			[OutputEmailAddr] [int] NULL,
			[OutputEmailSubject] [varchar](255) NULL,
			[OutputFilename] [varchar](255) NULL,
			[OutputEmailAttachAs] [varchar](255) NULL,
			[OutputTitlePage] [varchar](255) NULL,
			[OutputReportPackTitle] [varchar](255) NULL,
			[OutputOverrideFilter] [varchar](255) NULL,
			[OutputTOC] [bit] NULL,
			[OutputCoverSheet] [bit] NULL';
		EXEC sp_executesql @NVarCommand;
	  END
	  
	SELECT @iRecCount = count(id) FROM syscolumns WHERE id = (select id from sysobjects where name = 'ASRSysBatchJobName') and name = 'OverrideFilterID'
	IF @iRecCount = 0
	  BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysBatchJobName] ADD
			[OverrideFilterID] [int] NULL,
			[OutputRetainPivotOrChart] [bit] NULL';
		EXEC sp_executesql @NVarCommand;

		EXECUTE sp_executeSQL N'UPDATE AsrSysBatchJobName SET IsBatch = 1;';	
	  END 
	  	  
	-- Insert the system permissions for Report Packs and new picture too
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 44)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [picture])
			VALUES (44, 'Report Packs', 10, 'REPORTPACKS',0x0000010001001010000001000800680500001600000028000000100000002000000001000800000000000001000000000000000000000001000000010000000000006F685D00736A5E00746B5F00726C60007E7365007B7467007F746600867E6F008B7F68004FA31A0052A21A0057A01A00EA840000EA880A008D806C009B8B76009C8C760093887800A6967D00B09D7E00E6A24900EBB56C00ECBB78004B32BF003B29D1003C29D100533BC1005442CE009D918100A4988400AFA08800ADA18D00B5A28200B6A48200B4A38400B8A48200BFA88300BFA98600AFACA700B7B4B000B8B6B100C1AC8400C5AF8700C8B28900CAB48B00CDB68D00E6BF8800C5CEA800C8D1A900D7CBB800E5D6BC00EDDABD00A398CB00BAB0C600D8D1C300DBD5C900EADCC200E4DECF00E7E1D700F1ECDF00F6F1E800FCF7ED00FDF8EE00FFFAF10000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00000000000000000000000000000000000000201E1D12080604010000000000000000203A332F0A313A0100000000000000002133150D0C0C3002040100000000000024390E0D191835063A010000000000002539160D191B360930020401000000002A3C34171C32371035063A01000000002A4029292927381336093002000000002D40294040283B1437103506000000002D40292929293D213813360F000000002D40404040403E243B143710000000002E2E2D2D2D2A2A213D2138130000000000002E40404040403E243B140000000000002E2E2D2D2D2A2A213D2100000000000000002E40404040403E2400000000000000002E2E2D2D2D2A2A210000FFFF0000C03F0000C03F0000C00F0000C00F0000C0030000C0030000C0030000C0030000C0030000C0030000C0030000F0030000F0030000FC030000FC03000000);
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (158,44,'New', 10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (159,44,'Edit', 20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (160,44,'View', 30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (161,44,'Delete', 40, 'DELETE');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (162,44,'Run', 40, 'RUN');		
	END
	UPDATE dbo.[ASRSysPermissionCategories] SET picture = 0x4749463839611000100077000021F90401000033002C0000000010001000875D686F5E6A73606C7267747B6F7E86687F8B1AA34F1AA0570084EA0A88EA6C808D768B9B7888937D96A67E9DB049A2E66CB5EB78BBECBF324BD1293BC13B53CE425481919D8498A48DA1AD82A2B582A4B883A8BFA7ACAFB0B4B7B1B6B884ACC18BB4CA8DB6CD88BFE6A8CEC5A9D1C8B8CBD7BCD6E5BDDAEDCB98A3C6B0BAC3D1D8C9D5DBC2DCEACFDEE4D7E1E7DFECF1E8F1F6EDF7FCF1FAFFFFFFF000000704F2789D78CBD6922E9D76F504F2789D79339D78F4935EE90000077687CD7689600820FB00000404F51804F504768FFA0000047687CD7689600820FB00000404F53804F524768FFA00000404F5380520B700000076900700000404F5389D78F49DBD11407F2C407F2C000000010FF40000071D033E00000000000000000100B0299D8F0F000000010FF40000009E921C010FF41D033E0000010000010000070000000E0FF004F4A40000005BEC5E00000000B029600B505C1A200000000D1B805BEAF200000100B0290000000000070E0FF0935E3D04F3389DEE31600B5000000000B02900000000000000000104F3809DEE929D76F5A4426000000000B029D7FF3400000000B02900000000000100000004F35000000004F418767026767DFD0120FB85000F0000000120FB00000077565077563B04F3AC0001000CA05842FC8804F3AC0003006F685D736A5E726C607B7467717171867E6F8B7F684FA31A57A01AEA8400EA880A8D806C9B8B76938878A6967DB09D7EE6A249EBB56CECBB784B32BF3B29D1533BC15442CE9D9181A49884ADA18DB5A282B8A482BFA883AFACA7B7B4B0B8B6B1C1AC84CAB48BCDB68DE6BF88C5CEA8C8D1A9D7CBB8E5D6BCEDDABDA398CBBAB0C6D8D1C3DBD5C9EADCC2E4DECFE7E1D7F1ECDFF6F1E8FCF7EDFFFAF107000078F204CB789D2E92D6F5769D78F20433799DF4789DE95E930700002C7F40A4F4044DD59D240000010000000000000000700000FFFFFFFFFFFF7A789D33799D01000000000077495077585C0820FB00000000010004F52000000277565077495077585C0820FB00000000010004F54000000277565077563B04F53C089B0067081C18220488831F3E6418C8704608191021C6D0D0D0A141840A612C6CF830A28C892E1C302C781044C20C1A57346008C2E3470D21552C6029C3834D0F291BA45040D303C40E3117A018C0F0434D9B1C54A628302240D117272254282173680B000C37B08080600285A54D05601DA8814582AE12AC8E6598C1C40304070E845D3B10430B13220C90B85A5120860B1618101820B6AFE1BE0101003B00 WHERE [categoryID] = 44
	
	-- Adding Report Pack field to Event Log for Report Pack
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysEventLog')
	and name = 'ReportPack'

	IF @iRecCount = 0
	BEGIN
	 SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysEventLog] 
		ADD [ReportPack] bit NULL'
	 EXEC sp_executesql @NVarCommand
	END

	-- Extra file formats
	IF NOT EXISTS(SELECT ID FROM syscolumns	WHERE ID = (SELECT ID FROM sysobjects where [name] = 'ASRSysFileFormats') AND [name] = 'Direction')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysFileFormats] ADD [Direction] tinyint NULL;';
		EXEC sp_executesql N'DELETE FROM ASRSysFileFormats WHERE ID > 922;';
		EXEC sp_executesql N'UPDATE ASRSysFileFormats SET [direction] = 2;';
	END

	IF NOT EXISTS(SELECT * FROM ASRSysFileFormats where ID = 923)
	BEGIN
		EXEC sp_executesql N'INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default], [Direction])
			VALUES (923, ''Word'', ''PDF (*.pdf)'', ''pdf'', NULL, 17, 0, 1);'
		EXEC sp_executesql N'INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default], [Direction])
			VALUES (924, ''Word'', ''Rich Text Format (*.rtf)'', ''rtf'', NULL, 6, 0, 1);'
		EXEC sp_executesql N'INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default], [Direction])
			VALUES (925, ''Word'', ''Plain Text (*.txt)'', ''txt'', NULL, 2, 0, 1);'
		EXEC sp_executesql N'INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default], [Direction])
			VALUES (926, ''Word'', ''Web Page (*.html)'', ''html'', NULL, 8, 0, 1);		'
		EXEC sp_executesql N'INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default], [Direction])
			VALUES (927, ''Excel'', ''Web Page (*.html)'', ''html'', NULL, 44, 0, 1);'
	END


/* ------------------------------------------------------------- */
PRINT 'Step - Menu & Category enhancements'
/* ------------------------------------------------------------- */

	-- Categories
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbsys_objectcategories') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [tbsys_objectcategories](
			[objecttype]	smallint, 
			[objectid]		integer,
			[categoryid]	integer)';
		GRANT INSERT, UPDATE, SELECT, DELETE ON dbo.[tbsys_objectcategories] TO [ASRSysGroup];
	END

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spsys_getobjectcategories]') AND xtype = 'P')
	BEGIN
		EXECUTE sp_executesql N'CREATE PROCEDURE dbo.[spsys_getobjectcategories](@utilityType as integer, @UtilityID as integer, @TableID as integer)
		AS
		BEGIN

			DECLARE @iCount integer;
			-- Code generated by System Manager

		END'
		GRANT EXECUTE ON dbo.[spsys_getobjectcategories] TO [ASRSysGroup];
	END

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getcategory]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsys_getcategory];
	EXECUTE sp_executesql N'CREATE FUNCTION udfsys_getcategory(@categoryID as integer)
		RETURNS nvarchar(MAX)
		AS
		BEGIN

			-- Code generated by System Manager
			RETURN ''''
			
		END'
	GRANT EXECUTE ON dbo.[udfsys_getcategory] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spsys_saveobjectcategories]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_saveobjectcategories];
	EXECUTE sp_executesql N'CREATE PROCEDURE dbo.[spsys_saveobjectcategories](@utilityType as integer, @UtilityID as integer, @CategoryID integer)
		AS
		BEGIN

			SET NOCOUNT ON;

			DELETE FROM dbo.tbsys_objectcategories
				WHERE [objecttype] = @utilityType AND [objectid] = @UtilityID;
			
			IF @CategoryID > 0
			BEGIN
				INSERT dbo.tbsys_objectcategories([objecttype], [objectid], [categoryid])
					VALUES (@utilityType, @UtilityID, @CategoryID);
			END

		END'
	GRANT EXECUTE ON dbo.[spsys_saveobjectcategories] TO [ASRSysGroup];


	-- Generate, configure and populate the categories table
	DECLARE @categorytableid integer,
			@namecolumnid integer

	EXECUTE dbo.spstat_scriptnewtable @categorytableid OUTPUT, 'Object_Categories_Table', 3, 'D749F8E9-9625-47D3-889F-6A673B6C5F4A';
	IF @categorytableid > 0 
	BEGIN
		EXECUTE dbo.spstat_scriptnewcolumn @namecolumnid OUTPUT, @categorytableid, 'ID', 4, 'ID', 0, 0, 1, 'DB0F209B-A0C0-4B8B-9044-8519EB94C718', 1;
		EXECUTE dbo.spstat_scriptnewcolumn @namecolumnid OUTPUT, @categorytableid, 'Category_Name', 12, 'Category name', 50, 0, 1, 'CF24AEC1-28C3-4A44-A56C-B4363C275785', 0;

		-- Column specific details
		UPDATE dbo.tbsys_columns SET uniquecheck = 1, mandatory = 1, uniquechecktype = -1 WHERE columnID = @namecolumnid AND tableID = @categorytableid

		EXECUTE dbo.spstat_scriptnewprimaryorder @categorytableid, 'Category_Name';
		EXECUTE spstat_scriptnewprimaryscreen @categorytableid, 'Category';
		
		-- Configure the module
		EXECUTE dbo.spstat_setdefaultmodulesetting 'MODULE_CATEGORY', 'Param_CategoryTable', @categorytableid, 'PType_TableID';
		EXECUTE dbo.spstat_setdefaultmodulesetting 'MODULE_CATEGORY', 'Param_CatageoryNameColumn', @namecolumnid, 'PType_ColumnID';

		-- Populate with some basic categories
		IF @categorytableid > 0 AND @namecolumnid > 0
		BEGIN
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Personnel')		
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Absence')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Applicant')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Salary')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Documents')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Discipline')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Grievances')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Health & Safety')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Learning & Development')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Pension')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Configuration')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Statutory Leave')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('General')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Recruitment')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Post')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Skills')
			INSERT dbo.tbuser_Object_Categories_Table ([Category_Name]) VALUES ('Benefits')			
		END

		IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spsys_getobjectcategories]') AND xtype = 'P')
			DROP PROCEDURE [dbo].[spsys_getobjectcategories];
		EXECUTE sp_executesql N'CREATE PROCEDURE dbo.[spsys_getobjectcategories](@utilityType as integer, @UtilityID as integer, @TableID as integer)
			AS
			BEGIN

				-- To be regenerated by the system manager save process
				SELECT c.ID, c.[Category_Name] AS [category_name]
					, CASE ISNULL(s.categoryid,0) WHEN 0 THEN 0 ELSE 1 END   AS [selected]
					FROM dbo.[tbuser_Object_Categories_Table] c
						LEFT JOIN tbsys_objectcategories s ON s.CategoryID = c.ID AND s.objecttype = @utilityType AND s.objectid = @UtilityID
					ORDER BY c.[Category_Name];

			END'
		GRANT EXECUTE ON dbo.[spsys_getobjectcategories] TO [ASRSysGroup];
		
	END


	-- Populate catageories from existing definitions
	IF NOT EXISTS(SELECT [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'upgrade' and [SettingKey] = 'categoriesprocessed' AND [SettingValue] = 1)
	BEGIN
		
		-- Generic categories
		SELECT @configID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Configuration';
		SELECT @recruitmentID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Recruitment';
		SELECT @absenceID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Absence';
		SELECT @salaryID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Salary';
		SELECT @generalID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'General';
		SELECT @disciplineID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Discipline';
		SELECT @healthID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Health & Safety';	
		SELECT @trainingID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Learning & Development';	
		SELECT @skillID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Skills';	
		SELECT @benefitID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Benefits';	
		SELECT @statabsenceID = ID FROM tbuser_Object_Categories_Table WHERE [Category_Name] = 'Statutory Leave';	

		-- Build category match lookup
		INSERT @categorymatch
			SELECT ParameterValue AS TableID, ID AS categoryid FROM tbuser_Object_Categories_Table
				INNER JOIN ASRSysModuleSetup ON ModuleKey = 'MODULE_PERSONNEL' AND [ParameterType] = 'PType_TableID'
				WHERE [Category_Name] = 'Personnel'
			UNION
			SELECT ParameterValue AS TableID, ID AS categoryid FROM tbuser_Object_Categories_Table
				INNER JOIN ASRSysModuleSetup ON ModuleKey IN ('MODULE_ABSENCE') AND [ParameterType] = 'PType_TableID'
				WHERE [Category_Name] = 'Absence'
			UNION
			SELECT ParameterValue AS TableID, ID AS categoryid FROM tbuser_Object_Categories_Table
				INNER JOIN ASRSysModuleSetup ON ModuleKey = 'MODULE_TRAININGBOOKING' AND [ParameterType] = 'PType_TableID' AND ParameterKey <> 'Param_EmployeeTable'
				WHERE [Category_Name] = 'Learning & Development'
			UNION
			SELECT ParameterValue AS TableID, ID AS categoryid FROM tbuser_Object_Categories_Table
				INNER JOIN ASRSysModuleSetup ON ModuleKey = 'MODULE_POST' AND [ParameterType] = 'PType_TableID'
				WHERE [Category_Name] = 'Post'
			UNION
				SELECT TableID, @configID AS categoryid FROM tbsys_tables t WHERE t.TableType = 3
			UNION	
				SELECT TableID, @recruitmentID AS categoryid FROM ASRSysTables WHERE (tablename LIKE '%vacanc%' OR tablename LIKE '%applicant%') AND tabletype <> 3
			UNION	
				SELECT TableID, @absenceID AS categoryid FROM ASRSysTables WHERE (tablename LIKE '%absence%' OR tablename LIKE '%holiday%') AND tabletype <> 3
			UNION	
				SELECT TableID, @salaryID AS categoryid FROM ASRSysTables WHERE (tablename LIKE '%salary%' OR tablename LIKE '%deduct%' OR tablename LIKE '%allowance%' )   AND tabletype <> 3
			UNION
				SELECT TableID, @recruitmentID FROM ASRSysTables WHERE tablename LIKE '%applica%' AND tabletype <> 3
			UNION
				SELECT TableID, @disciplineID FROM ASRSysTables WHERE tablename LIKE '%disciplin%' AND tabletype <> 3
			UNION
				SELECT TableID, @trainingID FROM ASRSysTables WHERE tablename LIKE '%training%' AND tabletype <> 3				
			UNION	
				SELECT TableID, @healthID AS categoryid FROM ASRSysTables WHERE (tablename LIKE '%incidents%' OR tablename LIKE '%health%' OR tablename LIKE '%safety%') AND tabletype <> 3
			UNION	
				SELECT TableID, @skillID AS categoryid FROM ASRSysTables 
					WHERE (tablename LIKE '%qualification%' OR tablename LIKE '%competenc%' 
					OR tablename LIKE '%nvq%' OR tablename LIKE '%language%') AND tabletype <> 3
			UNION
				SELECT TableID, @benefitID AS categoryid FROM ASRSysTables
					WHERE (tablename LIKE '%benefit%' OR tablename LIKE '%eye_test%' OR tablename LIKE '%loans%') 
						AND tabletype <> 3			
			UNION
				SELECT TableID, @statabsenceID AS categoryid FROM ASRSysTables
					WHERE (tablename LIKE '%maternity%' OR tablename LIKE '%paternity%' OR tablename LIKE '%adoption%') 
						AND tabletype <> 3			

		DELETE FROM @categorymatch WHERE categoryID IS NULL

		--Make sure a table on maps to one category
		DELETE t FROM (SELECT ROW_NUMBER() OVER (PARTITION BY tableid ORDER BY categoryid) cnt FROM @categorymatch) t WHERE t.cnt > 1
		
		-- Globals Deletes/Updates
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT CASE [type] 
					WHEN 'D' THEN 6
					WHEN 'U' THEN 7
				END	AS [objectType], [FunctionID] AS ID, cat.categoryid
				FROM ASRSysGlobalFunctions g
					INNER JOIN @categorymatch cat ON cat.tableid = g.TableID
				WHERE [type] IN ('D', 'U')

		-- Global Adds
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 5 AS [objectType], [FunctionID] AS ID, cat.categoryid
				FROM ASRSysGlobalFunctions g
					INNER JOIN @categorymatch cat ON cat.tableid = g.ChildTableID
				WHERE [type] = 'A'

		-- Data Transfers
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 3 AS [objectType], [DataTransferID] AS ID, cat.categoryid
				FROM ASRSysDataTransferName d
					INNER JOIN @categorymatch cat ON cat.tableid = d.FromTableID

		-- Record Profile
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 20 AS [objectType], r.[RecordProfileID] AS ID, cat.categoryid
				FROM ASRSysRecordProfileName r
					INNER JOIN @categorymatch cat ON cat.tableid = r.BaseTable

		-- Match Reports / Succession / Career
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT CASE [MatchReportType] 
					WHEN 0 THEN 14 
					WHEN 1 THEN 23
					WHEN 2 THEN 24 
				END	AS [objectType], r.[MatchReportID] AS ID, cat.categoryid
				FROM ASRSysMatchReportName r
					INNER JOIN @categorymatch cat ON cat.tableid = r.Table1ID

		-- Cross Tabs
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 1 AS [objectType], c.CrossTabID AS ID, cat.categoryid
				FROM ASRSysCrossTab c
					INNER JOIN @categorymatch cat ON cat.tableid = c.TableID

		-- Imports
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 8 AS [objectType], i.ID AS ID, cat.categoryid
				FROM ASRSysImportName i
					INNER JOIN @categorymatch cat ON cat.tableid = i.BaseTable
		
		-- Mail Merge / Envelopes
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT CASE [IsLabel] 
					WHEN 0 THEN 9
					WHEN 1 THEN 18
				END	AS [objectType],  MailMergeID AS ID, cat.categoryid
				FROM ASRSysMailMergeName m
					INNER JOIN @categorymatch cat ON cat.tableid = m.TableID

		-- Custom Reports (recognised child tables)
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT DISTINCT 2 AS [objectType], r.ID, cat.categoryid
				FROM ASRSysCustomReportsName r
					INNER JOIN ASRSysCustomReportsChildDetails c ON c.CustomReportID = r.ID
					INNER JOIN @categorymatch cat ON cat.tableid = c.ChildTable
				WHERE r.id IN (SELECT c2.CustomReportID
									FROM ASRSysCustomReportsChildDetails c2			
									GROUP BY c2.CustomReportID
									HAVING COUNT(c2.CustomReportID) < 2)

		-- Custom Reports (unrecognised child tables - revert to base table)
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 2 AS [objectType], r.id AS ID, cat.categoryid
				FROM ASRSysCustomReportsName r
					INNER JOIN @categorymatch cat ON cat.tableid = r.BaseTable
				WHERE r.ID NOT IN( SELECT objectid FROM tbsys_objectcategories WHERE [objectType] = 2)

		-- Calendar reports
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 17 AS [objectType], c.ID AS ID, cat.categoryid
				FROM ASRSysCalendarReports c
					INNER JOIN @categorymatch cat ON cat.tableid = c.BaseTable

		-- Exports 
		INSERT tbsys_objectcategories ([objecttype], [objectid], [categoryid])
			SELECT 4 AS [objectType], e.id AS ID, cat.categoryid
				FROM ASRSysExportName e
					INNER JOIN @categorymatch cat ON cat.tableid = e.BaseTable
				WHERE e.ID NOT IN( SELECT objectid FROM tbsys_objectcategories WHERE [objectType] = 4)

		-- Flag so we don't process again
		EXEC dbo.spsys_setsystemsetting 'upgrade', 'categoriesprocessed', 1;

	END


	-- Menus
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbsys_userusage') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [tbsys_userusage](
			[objecttype]	smallint, 
			[objectid]	integer,
			[username]	varchar(255),
			[lastrun]	datetime,
			[runcount]	integer)';
		GRANT INSERT, UPDATE, SELECT, DELETE ON dbo.[tbsys_userusage] TO [ASRSysGroup];
	END

	IF NOT EXISTS(SELECT ID FROM syscolumns	WHERE ID = (SELECT ID FROM sysobjects where [name] = 'tbsys_userusage') AND [name] = 'lastaction')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE dbo.[tbsys_userusage] ADD [lastaction] integer NULL;';
		EXEC sp_executesql N'UPDATE dbo.[tbsys_userusage] SET [lastaction] = 16384;';
	END

	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbsys_userfavourites') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [tbsys_userfavourites](
			[username]		varchar(255),
			[objecttype]	smallint, 
			[objectid]		integer,
			[dateset]		datetime)';
		GRANT INSERT, UPDATE, SELECT, DELETE ON dbo.[tbsys_userfavourites] TO [ASRSysGroup];
	END

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_updateobjectusage]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_updateobjectusage];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spstat_updateobjectusage](@objecttype integer, @objectid integer, @lastaction integer)
	AS	
	BEGIN	
		SET NOCOUNT ON;

		DECLARE @sUsername varchar(255)

		SET @sUsername = SYSTEM_USER;

		IF NOT EXISTS(SELECT [objectid] FROM dbo.[tbsys_userusage] WHERE [objecttype] = @objecttype AND [objectid] = @objectID AND [username] = @sUsername)
		BEGIN
			INSERT tbsys_userusage (objecttype, objectid, username, lastrun, runcount, [lastaction])
				VALUES (@objecttype, @objectID, @sUsername , GETDATE(), 1, @lastaction)
		END
		ELSE
		BEGIN
			UPDATE dbo.[tbsys_userusage] SET [lastrun] = GETDATE(), [runcount] = [runcount] + 1, [lastaction] = @lastaction
				WHERE [objecttype] = @objecttype AND [objectid] = @objectID AND [username] = @sUsername
		END

	END';
	GRANT EXECUTE ON dbo.[spstat_updateobjectusage] TO [ASRSysGroup];


	IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[dbo].[ASRSysAllObjectNames]'))
		DROP VIEW [dbo].[ASRSysAllobjectNames]
	EXEC sp_executesql N'CREATE VIEW dbo.[ASRSysAllObjectNames]
	AS
		SELECT 25 AS [objectType], [ID], [Name], '''' AS Username, description FROM ASRSysWorkflows
		UNION
		SELECT CASE [IsBatch] 
				WHEN 0 THEN 29
				WHEN 1 THEN 0
			END	, ID,  Name, Username, description FROM ASRSysBatchJobName
		UNION
		SELECT CASE [IsLabel] 
				WHEN 0 THEN 9
				WHEN 1 THEN 18
			END	AS [objectType],  MailMergeID AS ID, Name, Username, description FROM ASRSysMailMergeName
		UNION
		SELECT 2 AS [objectType], ID, Name, Username, description FROM ASRSysCustomReportsName
		UNION
		SELECT 1 AS [objectType], CrossTabID AS ID, Name, Username, description FROM ASRSysCrossTab
		UNION		
		SELECT CASE [MatchReportType] 
				WHEN 0 THEN 14 
				WHEN 1 THEN 23
				WHEN 2 THEN 24 
			END	AS [objectType], MatchReportID AS ID, Name, Username, description FROM ASRSysMatchReportName			
		UNION
		SELECT 4 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysExportName
		UNION		
		SELECT 8 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysImportName
		UNION
		SELECT 3 AS [objectType], DataTransferID AS ID, Name, Username, description FROM ASRSysDataTransferName
		UNION
		SELECT CASE [type] 
				WHEN ''A'' THEN 5
				WHEN ''D'' THEN 6
				WHEN ''U'' THEN 7
			END	AS [objectType], [FunctionID] AS ID, Name, Username, description FROM ASRSysGlobalFunctions
		UNION		
		SELECT 15 AS [objectType], 0 AS ID, ''Absence Breakdown'', '''' AS Username, '''' AS Description
		UNION
		SELECT 16 AS [objectType], 0 AS ID, ''Bradford Factor'', '''' AS Username, '''' AS Description
		UNION
		SELECT 17 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysCalendarReports
		UNION		
		SELECT 20 AS [objectType], RecordProfileID AS ID, Name, Username, description FROM ASRSysRecordProfileName
		UNION
		SELECT 30 AS [objectType], 0 AS ID, ''Turnover'', '''' AS Username, '''' AS Description
		UNION
		SELECT 31 AS [objectType], 0 AS ID, ''Stability Index'', '''' AS Username, '''' AS Description;'
	GRANT SELECT ON dbo.[ASRSysAllobjectNames] TO [ASRSysGroup];


	IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[dbo].[ASRSysAllObjectAccess]'))
		DROP VIEW [dbo].[ASRSysAllObjectAccess]
	EXEC sp_executesql N'CREATE VIEW dbo.[ASRSysAllObjectAccess]
	AS
		SELECT CASE b.[IsBatch] 
					WHEN 0 THEN 29
					WHEN 1 THEN 0
				END	AS [objectType], a.* FROM ASRSysBatchJobAccess a
			INNER JOIN ASRSysBatchJobName b ON a.ID = b.ID
		UNION
		SELECT 1 AS [objectType], * FROM [ASRSysCrossTabAccess]
		UNION
		SELECT 2 AS [objectType], * FROM [ASRSysCustomReportAccess]
		UNION
		SELECT 3 AS [objectType], * FROM [ASRSysDataTransferAccess]
		UNION
		SELECT 4 AS [objectType], * FROM [ASRSysExportAccess]
		UNION		
		SELECT CASE [type] 
					WHEN ''A'' THEN 5
					WHEN ''D'' THEN 6
					WHEN ''U'' THEN 7
				END	AS [objectType], a.* FROM [ASRSysGlobalAccess] a
			INNER JOIN ASRSysGlobalFunctions g ON a.ID = g.functionID	
		UNION
		SELECT 8 AS [objectType], * FROM [ASRSysImportAccess]
		UNION
		SELECT CASE m.[IsLabel] 
					WHEN 0 THEN 9
					WHEN 1 THEN 18
				END	AS [objectType], a.* FROM ASRSysMailMergeAccess a
			INNER JOIN ASRSysMailMergeName m ON a.ID = m.MailMergeID
		UNION
		SELECT CASE [MatchReportType] 
				WHEN 0 THEN 14 
				WHEN 1 THEN 23
				WHEN 2 THEN 24 
			END	AS [objectType], a.* FROM ASRSysMatchReportAccess a
			INNER JOIN ASRSysMatchReportName m ON a.ID = m.MatchReportID			
		UNION
		SELECT 17 AS [objectType], * FROM ASRSysCalendarReportAccess
		UNION
		SELECT 20 AS [objectType], * FROM [ASRSysRecordProfileAccess]';
	GRANT SELECT ON dbo.[ASRSysAllObjectAccess] TO [ASRSysGroup];


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_recentlyrunobjects]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_recentlyrunobjects];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_recentlyrunobjects]
	AS
	BEGIN
		SET NOCOUNT ON;

		SELECT TOP 10 ROW_NUMBER() OVER (ORDER BY [lastrun] DESC) AS ID, u.[objectid], o.[Name], o.[objectType]
			FROM tbsys_userusage u
			INNER JOIN ASRSysAllobjectNames o ON o.[objectType] = u.objecttype AND o.[ID] = u.objectid
			WHERE u.[username] = SYSTEM_USER
			ORDER BY u.[lastrun] DESC

	END';
	GRANT EXECUTE ON dbo.[spstat_recentlyrunobjects] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_clearrecentusage]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_clearrecentusage];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_clearrecentusage]
		AS
		BEGIN
			SET NOCOUNT ON;
			
			DELETE FROM dbo.tbsys_userusage WHERE [username] = SYSTEM_USER;
		END';
	GRANT EXECUTE ON dbo.[spstat_clearrecentusage] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getfavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getfavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_getfavourites]
	AS
	BEGIN
		SET NOCOUNT ON;

		SELECT o.[objectType], f.[objectid], o.[Name]
			FROM tbsys_userfavourites f
			INNER JOIN ASRSysAllobjectNames o ON o.[objectType] = f.[objecttype] AND o.[ID] = f.objectid
			WHERE f.[username] = SYSTEM_USER
			ORDER BY o.[Name]

	END';
	GRANT EXECUTE ON dbo.[spstat_getfavourites] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_addtofavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_addtofavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_addtofavourites](@objecttype integer, @objectid integer)
		AS
		BEGIN
			SET NOCOUNT ON;

			DECLARE @now datetime;
			SET @now = GETDATE();
			
			IF NOT EXISTS(SELECT [username] FROM dbo.tbsys_userfavourites 
								WHERE [username] = SYSTEM_USER AND @objectid = [objectid] AND @objecttype = [objecttype])
			BEGIN
				INSERT dbo.tbsys_userfavourites (username, objecttype, objectid, dateset)
					VALUES (SYSTEM_USER, @objecttype, @objectid, @now);
			END

		END';
	GRANT EXECUTE ON dbo.[spstat_addtofavourites] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_removefromfavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_removefromfavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_removefromfavourites](@objecttype integer, @objectid integer)
		AS
		BEGIN	
			SET NOCOUNT ON;
		
			DELETE FROM dbo.tbsys_userfavourites
				WHERE [username] = SYSTEM_USER AND @objectid = [objectid] AND @objecttype = [objecttype]
		END';
	GRANT EXECUTE ON dbo.[spstat_removefromfavourites] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_clearfavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_clearfavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_clearfavourites]
		AS
		BEGIN
			SET NOCOUNT ON;

			DELETE FROM dbo.tbsys_userfavourites WHERE [username] = SYSTEM_USER;
		END';
	GRANT EXECUTE ON dbo.[spstat_clearfavourites] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRGetHistoryScreens]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRGetHistoryScreens];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[sp_ASRGetHistoryScreens]
			(@piParentScreenID	integer)
		AS
		BEGIN
			-- Return a recordset of the history screens that hang off the given parent screen.
			SELECT CASE WHEN ISNULL(cs.category,0) = 0 OR ISNULL(parentScreen.groupscreens,0) = 0
				THEN t.tableName
				ELSE ISNULL(dbo.udfsys_getcategory(cs.category),t.tableName)
				END AS tableName,
				t.tableID,
				cs.screenID,
				cs.name,
				cs.pictureID,
				t.tableName AS [realsource]
			FROM ASRSysScreens parentScreen
			INNER JOIN ASRSysHistoryScreens h ON parentScreen.screenID = h.parentScreenID
			INNER JOIN ASRSysScreens cs ON h.historyScreenID = cs.screenID
			INNER JOIN ASRSysTables t ON cs.tableID = t.tableID
			WHERE parentScreen.screenID = @piParentScreenID	AND cs.quickEntry = 0
			ORDER BY tableName, h.[order], cs.Name;
		END'
	GRANT EXECUTE ON dbo.[sp_ASRGetHistoryScreens] TO [ASRSysGroup];

	-- Prepopulate the objects with some categories
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbuser_Object_Categories_Table') AND type in (N'U'))
	BEGIN	

		SELECT @tableid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup]
			WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_TablePersonnel' AND [ParameterType] = 'PType_TableID';
		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Personnel'
		SET @NVarCommand = 'UPDATE ASRSysScreens SET category = ' + convert(varchar(10),@categoryid) + '
			WHERE TableID = ' + convert(varchar(10),@tableid) + ' AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Absence'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + 'FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE (Name LIKE ''%Absence%'' OR Name LIKE ''%Maternity%'' OR Name LIKE ''%Paternity%'' 
				OR Name LIKE ''%Adoption%'' OR Name LIKE ''%ASPP%'')
				AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Salary'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE (Name LIKE ''%Allowances%'' OR Name LIKE ''%Benefits%'' OR Name LIKE ''%Bonuses%'' 
				OR Name LIKE ''%Deductions%'' OR Name LIKE ''%Loans%'' OR Name LIKE ''%Pensions%'' OR Name LIKE ''%Salary%'' OR Name LIKE ''%Timesheets%'')
				AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Learning & Development'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE (Name LIKE ''%Competencies%'' OR Name LIKE ''%Course Bookings%''
				OR Name LIKE ''%CPD Summary%'' OR Name LIKE ''%Languages%'' OR Name LIKE ''%Qualifications%'' OR Name LIKE ''%Subscriptions%'' OR Name LIKE ''%Training%'')
				AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;			

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Health & Safety'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE (Name LIKE ''%Eye Tests%'' OR Name LIKE ''%Medicals%'' OR Name LIKE ''%Risk Assessments%'' OR Name LIKE ''%H&S%'')
				AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Applicant'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			WHERE Name LIKE ''%Applicant%'' AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Recruitment'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE Name LIKE ''%Vacancy%'' AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'Training'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 2
			WHERE (Name LIKE ''%Training%'' OR Name LIKE ''%Course%'' OR Name LIKE ''%Delegate%'')
				AND category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		SELECT @categoryid = [ID] FROM dbo.tbuser_Object_Categories_Table WHERE Category_name = 'General'
		SET @NVarCommand = 'UPDATE s SET s.category = ' + convert(varchar(10),@categoryid) + '
			FROM ASRSysScreens s
			INNER JOIN ASRSysTables t ON t.tableid = s.tableid AND t.tabletype = 3
			WHERE category IS NULL';
		EXECUTE sp_executeSQL @NVarCommand;

		EXECUTE sp_executeSQL N'UPDATE ASRSysScreens SET category = 0 WHERE category IS NULL';

	END



/* ------------------------------------------------------------- */
/* Step - Updating workflow stored procedures */
/* ------------------------------------------------------------- */
	----------------------------------------------------------------------
	-- spASRSubmitWorkflowStep
	----------------------------------------------------------------------*/

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSubmitWorkflowStep]
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psFormInput1		varchar(MAX),
				@psFormElements		varchar(MAX)	OUTPUT,
				@pfSavedForLater	bit				OUTPUT,
				@piPageNo	integer
			)
			AS
			BEGIN
				DECLARE
					@iIndex1			integer,
					@iIndex2			integer,
					@iID				integer,
					@sID				varchar(MAX),
					@sValue				varchar(MAX),
					@iElementType		integer,
					@iPreviousElementID	integer,
					@iValue				integer,
					@hResult			integer,
					@hTmpResult			integer,
					@sTo				varchar(MAX),
					@sCopyTo			varchar(MAX),
					@sTempTo			varchar(MAX),
					@sMessage			varchar(MAX),
					@sMessage_HypertextLinks	varchar(MAX),
					@sHypertextLinkedSteps		varchar(MAX),
					@iEmailID			integer,
					@iEmailCopyID		integer,
					@iTempEmailID		integer,
					@iEmailLoop			integer,
					@iEmailRecord		integer,
					@iEmailRecordID		integer,
					@sSQL				nvarchar(MAX),
					@iCount				integer,
					@superCursor		cursor,
					@curDelegatedRecords	cursor,
					@fDelegate			bit,
					@fDelegationValid	bit,
					@iDelegateEmailID	integer,
					@iDelegateRecordID	integer,
					@sTemp				varchar(MAX),
					@sDelegateTo		varchar(MAX),
					@sAllDelegateTo		varchar(MAX),
					@iCurrentStepID		int,
					@sDelegatedMessage	varchar(MAX),
					@iTemp				integer, 
					@iPrevElementType	integer,
					@iWorkflowID		integer,
					@sRecSelIdentifier	varchar(MAX),
					@sRecSelWebFormIdentifier	varchar(MAX), 
					@iStepID			int,
					@iElementID			int,
					@sUserName			varchar(MAX),
					@sUserEmail			varchar(MAX), 
					@sValueDescription	varchar(MAX),
					@iTableID			integer,
					@iRecDescID			integer,
					@sEvalRecDesc		varchar(MAX),
					@sExecString		nvarchar(MAX),
					@sParamDefinition	nvarchar(500),
					@sIdentifier		varchar(MAX),
					@iItemType			integer,
					@iDataAction		integer, 
					@fValidRecordID		bit,
					@iEmailTableID		integer,
					@iEmailType			integer,
					@iBaseTableID		integer,
					@iBaseRecordID		integer,
					@iRequiredRecordID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iTempElementID		integer,
					@iTrueFlowType		integer,
					@iExprID			integer,
					@iResultType		integer,
					@sResult			varchar(MAX),
					@fResult			bit,
					@dtResult			datetime,
					@fltResult			float,
					@sEmailSubject		varchar(200),
					@iTempID			integer,
					@iBehaviour			integer;
		
				SET @pfSavedForLater = 0;
		
				SELECT @iCurrentStepID = ID
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
				SET @iDelegateEmailID = 0;
				SELECT @sTemp = ISNULL(parameterValue, '''')
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_DelegateEmail'';
				SET @iDelegateEmailID = convert(integer, @sTemp);
		
				SET @psFormElements = '''';
						
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
				WHERE E.ID = @piElementID;
		
				--------------------------------------------------
				-- Read the submitted webForm/storedData values
				--------------------------------------------------
				IF @iElementType = 5 -- Stored Data element
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
					SET @sValue = LEFT(@psFormInput1, @iIndex1-1);
					SET @sTemp = SUBSTRING(@psFormInput1, @iIndex1+1, LEN(@psFormInput1) - @iIndex1);
		
					SET @sValueDescription = '''';
					SET @sMessage = ''Successfully '' +
						CASE
							WHEN @iDataAction = 0 THEN ''inserted''
							WHEN @iDataAction = 1 THEN ''updated''
							ELSE ''deleted''
						END + '' record'';
		
					IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
					BEGIN
						SET @sValueDescription = @sTemp;
					END
					ELSE
					BEGIN
						SET @iTemp = convert(integer, @sValue);
						IF @iTemp > 0 
						BEGIN	
							EXEC [dbo].[spASRRecordDescription] 
								@iTableID,
								@iTemp,
								@sEvalRecDesc OUTPUT
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
						END
					END
		
					IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')'';
		
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
						AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
				END
				ELSE
				BEGIN
					-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
					WHILE (charindex(CHAR(9), @psFormInput1) > 0)
					BEGIN
		
						SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
						SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1);
						SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''');
						SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1);
						SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2);
		
						--Get the record description (for RecordSelectors only)
						SET @sValueDescription = '''';
		
						-- Get the WebForm item type, etc.
						SELECT @sIdentifier = EI.identifier,
							@iItemType = EI.itemType,
							@iTableID = EI.tableID,
							@iBehaviour = EI.behaviour
						FROM ASRSysWorkflowElementItems EI
						WHERE EI.ID = convert(integer, @sID);
		
						SET @iParent1TableID = 0;
						SET @iParent1RecordID = 0;
						SET @iParent2TableID = 0;
						SET @iParent2RecordID = 0;
		
						IF @iItemType = 11 -- Record Selector
						BEGIN
							-- Get the table record description ID. 
							SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
							FROM ASRSysTables 
							WHERE ASRSysTables.tableID = @iTableID;
		
							SET @iTemp = convert(integer, isnull(@sValue, ''0''));
		
							-- Get the record description. 
							IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
							BEGIN
								SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(MAX), @iRecDescID) + '' @recDesc OUTPUT, @recID'';
								SET @sParamDefinition = N''@recDesc varchar(MAX) OUTPUT, @recID integer'';
								EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp;
								IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
							END
		
							-- Record the selected record''s parent details.
							exec [dbo].[spASRGetParentDetails]
								@iTableID,
								@iTemp,
								@iParent1TableID	OUTPUT,
								@iParent1RecordID	OUTPUT,
								@iParent2TableID	OUTPUT,
								@iParent2RecordID	OUTPUT;
						END
						ELSE
						IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
						BEGIN
							SET @pfSavedForLater = 1;
						END
		
						IF (@iItemType = 17) -- FileUpload Control
						BEGIN
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.fileUpload_File = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_File
									ELSE null
								END,
								ASRSysWorkflowInstanceValues.fileUpload_ContentType = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_ContentType
									ELSE null
								END,
								ASRSysWorkflowInstanceValues.fileUpload_FileName = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_FileName
									ELSE null
								END
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @piElementID
								AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
						END
						ELSE
						BEGIN
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = @sValue, 
								ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
								ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
								ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
								ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
								ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @piElementID
								AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
						END
					END
		
					IF @pfSavedForLater = 1
					BEGIN
						/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 7
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
						
						/* Remember the page number too  */
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.pageno = @piPageNo
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
						RETURN;
					END
				END
					
				SET @hResult = 0;
				SET @sTo = '''';
				SET @sCopyTo = '''';
		
				--------------------------------------------------
				-- Process email element
				--------------------------------------------------
				IF @iElementType = 3 -- Email element
				BEGIN
					-- Get the email recipient. 
					SET @iEmailRecordID = 0;
					SET @sSQL = ''spASRSysEmailAddr'';
		
					IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SET @iEmailLoop = 0
						WHILE @iEmailLoop < 2
						BEGIN
							SET @hTmpResult = 0;
							SET @sTempTo = '''';
							SET @iTempEmailID = 
								CASE 
									WHEN @iEmailLoop = 1 THEN @iEmailCopyID
									ELSE isnull(@iEmailID, 0)
								END;
		
							IF @iTempEmailID > 0 
							BEGIN
								SET @fValidRecordID = 1;
		
								SELECT @iEmailTableID = isnull(tableID, 0),
									@iEmailType = isnull(type, 0)
								FROM ASRSysEmailAddress
								WHERE emailID = @iTempEmailID;
		
								IF @iEmailType = 0 
								BEGIN
									SET @iEmailRecordID = 0;
								END
								ELSE
								BEGIN
									SET @iTempElementID = 0;
		
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
										WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
										SET @iBaseRecordID = @iEmailRecordID;
		
										IF @iEmailRecord = 4
										BEGIN
											-- Trigger record
											SELECT @iBaseTableID = isnull(WF.baseTable, 0)
											FROM ASRSysWorkflows WF
											INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
												AND WFI.ID = @piInstanceID;
										END
										ELSE
										BEGIN
											-- Initiator''s record
											SELECT @iBaseTableID = convert(integer, ISNULL(parameterValue, ''0''))
											FROM ASRSysModuleSetup
											WHERE moduleKey = ''MODULE_PERSONNEL''
												AND parameterKey = ''Param_TablePersonnel'';
		
											IF @iBaseTableID = 0
											BEGIN
												SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
												FROM ASRSysModuleSetup
												WHERE moduleKey = ''MODULE_WORKFLOW''
												AND parameterKey = ''Param_TablePersonnel'';
											END
										END
									END
		
									IF @iEmailRecord = 1
									BEGIN
										SELECT @iPrevElementType = ASRSysWorkflowElements.type,
											@iTempElementID = ASRSysWorkflowElements.ID
										FROM ASRSysWorkflowElements
										WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
											AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
										IF @iPrevElementType = 2
										BEGIN
											 -- WebForm
											SELECT @sValue = ISNULL(IV.value, ''0''),
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
												AND IV.elementID = Es.ID;
										END
										ELSE
										BEGIN
											-- StoredData
											SELECT @sValue = ISNULL(IV.value, ''0''),
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
											WHERE IV.instanceID = @piInstanceID;
										END
		
										SET @iEmailRecordID = 
											CASE
												WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
												ELSE 0
											END;
		
										SET @iBaseRecordID = @iEmailRecordID;
									END
		
									SET @fValidRecordID = 1;
									IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
									BEGIN
										SET @fValidRecordID = 0;
		
										EXEC [dbo].[spASRWorkflowAscendantRecordID]
											@iBaseTableID,
											@iBaseRecordID,
											@iParent1TableID,
											@iParent1RecordID,
											@iParent2TableID,
											@iParent2RecordID,
											@iEmailTableID,
											@iRequiredRecordID	OUTPUT;
		
										SET @iEmailRecordID = @iRequiredRecordID;
		
										IF @iRequiredRecordID > 0 
										BEGIN
											EXEC [dbo].[spASRWorkflowValidTableRecord]
												@iEmailTableID,
												@iEmailRecordID,
												@fValidRecordID	OUTPUT;
										END
		
										IF @fValidRecordID = 0
										BEGIN
											IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
											BEGIN
												SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
												FROM ASRSysWorkflowQueueColumns QC
												INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
												WHERE WFQ.instanceID = @piInstanceID
													AND QC.emailID = @iTempEmailID;
		
												IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
											END
											ELSE
											BEGIN
												IF @iEmailRecord = 1
												BEGIN
													SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
													FROM ASRSysWorkflowInstanceValues IV
													WHERE IV.instanceID = @piInstanceID
														AND IV.emailID = @iTempEmailID
														AND IV.elementID = @iTempElementID;
		
													IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
												END
											END
										END
		
										IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
										BEGIN
											-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
											EXEC [dbo].[spASRWorkflowActionFailed] 
												@piInstanceID, 
												@piElementID, 
												''Email record has been deleted or not selected.'';
													
											SET @hTmpResult = -1;
										END
									END
								END
		
								IF @fValidRecordID = 1
								BEGIN
									/* Get the recipient address. */
									IF len(@sTempTo) = 0
									BEGIN
										EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID;
										IF @sTempTo IS null SET @sTempTo = '''';
									END
		
									IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Email step failure if no known recipient.
										-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
											@piInstanceID, 
											@piElementID, 
											''No email recipient.'';
												
										SET @hTmpResult = -1;
									END
								END
		
								IF @iEmailLoop = 1 
								BEGIN
									SET @sCopyTo = @sTempTo;
		
									IF (rtrim(ltrim(@sCopyTo)) = ''@'')
										OR (charindex('' @ '', @sCopyTo) > 0)
									BEGIN
										SET @sCopyTo = '''';
									END
								END
								ELSE
								BEGIN
									SET @sTo = @sTempTo;
								END
							END
						
							SET @iEmailLoop = @iEmailLoop + 1;
		
							IF @hTmpResult <> 0 SET @hResult = @hTmpResult;
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
								AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Invalid email recipient.'';
						
							SET @hResult = -1;
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
								@sTo;
		
							IF @fValidRecordID = 1
							BEGIN
								exec [dbo].[spASRDelegateWorkflowEmail] 
									@sTo,
									@sCopyTo,
									@sMessage,
									@sMessage_HypertextLinks,
									@iCurrentStepID,
									@sEmailSubject;
							END
							ELSE
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC [dbo].[spASRWorkflowActionFailed] 
									@piInstanceID, 
									@piElementID, 
									''Email item database value record has been deleted or not selected.'';
										
								SET @hResult = -1;
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
						ASRSysWorkflowInstanceSteps.message = CASE
							WHEN @iElementType = 3 THEN @sMessage
							WHEN @iElementType = 5 THEN @sMessage
							ELSE ''''
						END,
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
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
								0;
		
							SET @iValue = convert(integer, @fResult);
						END
						ELSE
						BEGIN
							-- Decision Element flow determined by a button in a preceding web form
							SET @iPrevElementType = 4; -- Decision element
							SET @iPreviousElementID = @piElementID;
		
							WHILE (@iPrevElementType = 4)
							BEGIN
								SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
									@iPrevElementType = isnull(WE.type, 0)
								FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
								INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
								INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
									AND WIS.instanceID = @piInstanceID;
		
								SET @iPreviousElementID = @iTempID;
							END
					
							SELECT @sValue = ISNULL(IV.value, ''0'')
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
							WHERE IV.elementID = @iPreviousElementID
								AND IV.instanceid = @piInstanceID
								AND E.ID = @piElementID;
		
							SET @iValue = 
								CASE
									WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
									ELSE 0
								END;
						END
				
						IF @iValue IS null SET @iValue = 0;
		
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
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
								OR ASRSysWorkflowInstanceSteps.status = 3);
					END
					ELSE
					BEGIN
						IF @iElementType <> 3 -- 3=Email element
						BEGIN
							IF @iElementType = 2 -- WebForm
							BEGIN
								SELECT @sUserName = isnull(WIS.userName, ''''),
									@sUserEmail = isnull(WIS.userEmail, '''')
								FROM ASRSysWorkflowInstanceSteps WIS
								WHERE WIS.instanceID = @piInstanceID
									AND WIS.elementID = @piElementID;
							END;
									
							-- Do not the following bit when the submitted element is an Email element as 
							-- the succeeding elements will already have been actioned.
							DECLARE @succeedingElements TABLE(elementID integer);
		
							EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
								@piInstanceID, 
								@piElementID, 
								@superCursor OUTPUT,
								'''';
		
							FETCH NEXT FROM @superCursor INTO @iTemp;
							WHILE (@@fetch_status = 0)
							BEGIN
								INSERT INTO @succeedingElements (elementID) VALUES (@iTemp);
							
								FETCH NEXT FROM @superCursor INTO @iTemp;
							END
							CLOSE @superCursor;
							DEALLOCATE @superCursor;
		
							-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
							IF @iElementType = 2 -- WebForm
							BEGIN
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
										OR ASRSysWorkflowInstanceSteps.status = 3);
		
								OPEN formsCursor;
								FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
								WHILE (@@fetch_status = 0) 
								BEGIN
									SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
									DELETE FROM ASRSysWorkflowStepDelegation
									WHERE stepID = @iStepID;
		
									INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
										(SELECT WSD.delegateEmail, @iStepID
										FROM ASRSysWorkflowStepDelegation WSD
										WHERE WSD.stepID = @iCurrentStepID);
								
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
											OR ASRSysWorkflowInstanceSteps.status = 3);
								
									FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
								END
								CLOSE formsCursor;
								DEALLOCATE formsCursor;
		
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
										OR ASRSysWorkflowInstanceSteps.status = 3);
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
											OR ASRSysWorkflowInstanceSteps.status = 3));
							
								INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
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
								WHERE WSD.stepID = @iCurrentStepID);
		
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
										OR ASRSysWorkflowInstanceSteps.status = 3);
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
							AND ASRSysWorkflowElements.type = 2);
		
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
							AND ASRSysWorkflowElements.type = 1);
		
					-- Count how many terminators have completed. ie. if the workflow has completed. 
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.status = 3
						AND ASRSysWorkflowElements.type = 1;
							
					IF @iCount > 0 
					BEGIN
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
							ASRSysWorkflowInstances.status = 3,
							ASRSysWorkflowInstances.pageno = @piPageNo
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
					
						-- Steps pending action are no longer required.
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
								OR ASRSysWorkflowInstanceSteps.status = 2); -- 2 = Pending User Action
					END
		
					IF @iElementType = 3 -- Email element
						OR @iElementType = 5 -- Stored Data element
					BEGIN
						exec [dbo].[spASREmailImmediate] ''OpenHR Workflow'';
					END
				END
			END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowItemValues
	----------------------------------------------------------------------

	IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRGetWorkflowItemValues]')	AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowItemValues];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowItemValues]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowItemValues]
			(
				@piElementItemID	integer,
				@piInstanceID	integer, 
				@piLookupColumnIndex	integer OUTPUT, 
				@piItemType	integer OUTPUT, 
				@psDefaultValue	varchar(8000) OUTPUT
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
					@iOrderID			integer,
					@iTableID			integer,
					@sSelectSQL			varchar(max),
					@sColumnList		varchar(max),
					@sOrderSQL			varchar(max),
					@sJoinSQL			varchar(max),
					@sJoinedTables		varchar(max),
					@fLookupColumnDoneF	bit,
					@sOrderType	char(1),
					@fOrderAsc	bit,
					@sOrderTableName	sysname,
					@sOrderColumnName	sysname,
					@iOrderColumnID	integer,
					@iOrderTableID	integer,
					@sTemp	varchar(max),
					@iCount	integer,
					@iStatus			integer,
					@iElementID			integer,
					@sValue				varchar(8000),
					@sIdentifier		varchar(8000),
					@sLookupFilterColumnName	varchar(8000),
					@iLookupFilterColumnType	int,
					@iLookupOrderID		int;

				SET @piLookupColumnIndex = 0;
								
				DECLARE @dropdownValues TABLE([value] varchar(255));

				SELECT 			
					@iItemType = ASRSysWorkflowElementItems.itemType,
					@sDefaultValue = ASRSysWorkflowElementItems.inputDefault,
					@iLookupColumnID = ASRSysWorkflowElementItems.lookupColumnID,
					@iElementID = ASRSysWorkflowElementItems.elementID,
					@sIdentifier = ASRSysWorkflowElementItems.identifier,
					@iCalcID = isnull(ASRSysWorkflowElementItems.calcID, 0),
					@iDefaultValueType = isnull(ASRSysWorkflowElementItems.defaultValueType, 0),
					@sLookupFilterColumnName = isnull(COLS.columnName, ''''),
					@iLookupFilterColumnType = isnull(COLS.dataType, 0),
					@iLookupOrderID = ASRSysWorkflowElementItems.LookupOrderID
				FROM ASRSysWorkflowElementItems
				LEFT OUTER JOIN ASRSysColumns COLS ON ASRSysWorkflowElementItems.LookupFilterColumnID = COLS.columnID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID;

				SET @piItemType = @iItemType;

				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

				IF @iStatus = 7 -- Previously SavedForLater
				BEGIN
					SELECT @sValue = isnull(IVs.value, '''')
					FROM ASRSysWorkflowInstanceValues IVs
					WHERE IVs.instanceID = @piInstanceID
						AND IVs.elementID = @iElementID
						AND IVs.identifier = @sIdentifier;

					SET @sDefaultValue = @sValue;
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
							0;

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
							END;
					END
				END

				SET @psDefaultValue = @sDefaultValue;

				IF @iItemType = 15 -- OptionGroup
				BEGIN
					SELECT ASRSysWorkflowElementItemValues.value,
						CASE
							WHEN ASRSysWorkflowElementItemValues.value = @sDefaultValue THEN 1
							ELSE 0
						END AS [ASRSysDefaultValueFlag]
					FROM ASRSysWorkflowElementItemValues
					WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
					ORDER BY ASRSysWorkflowElementItemValues.sequence;
				END

				IF @iItemType = 13 -- Dropdown
				BEGIN
					INSERT INTO @dropdownValues ([value])
						SELECT ASRSysWorkflowElementItemValues.value
						FROM ASRSysWorkflowElementItemValues
						WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
						ORDER BY [sequence];

					SELECT [value],
						'''' AS [ASRSysLookupFilterValue]				
					FROM @dropdownValues;
				END
				
				IF (@iItemType = 14) AND (@iLookupColumnID > 0) -- Lookup
				BEGIN
					SELECT @sTableName = ASRSysTables.tableName,
						@sColumnName = ASRSysColumns.columnName,
						@iOrderID = COALESCE(NULLIF(@iLookupOrderID, 0), ASRSysTables.defaultOrderID),
						@iTableID = ASRSysTables.tableID,
						@iDataType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iLookupColumnID;		

					IF @iDataType = 11 -- Date 
						AND UPPER(LTRIM(RTRIM(@sDefaultValue))) = ''NULL''
					BEGIN
						SET @sDefaultValue = '''';
					END

					SET @sColumnList = '''';
					SET @sJoinSQL ='''';
					SET @sOrderSQL = '''';
					SET @fLookupColumnDoneF = 0;
					SET @sJoinedTables = '','';
					SET @iCount = 0;
				
					DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ASRSysOrderItems.type,
						ASRSysTables.tableName,
						ASRSysColumns.columnName,
						ASRSysColumns.columnID,
						ASRSysColumns.tableID,
						ASRSysOrderItems.ascending
					FROM ASRSysOrderItems
					INNER JOIN ASRSysColumns 
						ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
					INNER JOIN ASRSysTables 
						ON ASRSysTables.tableID = ASRSysColumns.tableID
					WHERE ASRSysOrderItems.orderID = @iOrderID
					ORDER BY ASRSysOrderItems.type, 
						ASRSysOrderItems.sequence;

					OPEN orderCursor;
					FETCH NEXT FROM orderCursor INTO 
						@sOrderType, 
						@sOrderTableName,
						@sOrderColumnName,
						@iOrderColumnID,
						@iOrderTableID,
						@fOrderAsc;
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @sOrderType = ''F''
						BEGIN
							IF @iLookupColumnID = @iOrderColumnID
							BEGIN
								SET @fLookupColumnDoneF = 1;
								SET @piLookupColumnIndex = @iCount;
							END;
		
							SET @sColumnList = @sColumnList 
								+ CASE
										WHEN LEN(@sColumnList) > 0 THEN '',''
										ELSE ''''
									END
								+ @sOrderTableName + ''.'' + @sOrderColumnName;

							SET @iCount = @iCount + 1;
						END
						ELSE
						BEGIN
							SET @sOrderSQL = @sOrderSQL 
								+ CASE
										WHEN LEN(@sOrderSQL) > 0 THEN '',''
										ELSE ''''
									END
								+ @sOrderTableName + ''.'' + @sOrderColumnName	
								+CASE
										WHEN @fOrderAsc = 0 THEN '' DESC''
										ELSE ''''
									END;
						END;

						IF @iTableID <> @iOrderTableID
						BEGIN
							SET @sTemp = '','' + CONVERT(varchar(max), @iOrderTableID) + '',''
							IF CHARINDEX(@sTemp, @sJoinedTables) = 0
							BEGIN
								SET @sJoinedTables = @sJoinedTables + CONVERT(varchar(max), @iOrderTableID) + '','';
								
								SET @sJoinSQL = @sJoinSQL 
									+ '' LEFT OUTER JOIN '' + @sOrderTableName
									+ '' ON '' + @sTableName + ''.ID_'' + CONVERT(varchar(max), @iOrderTableID)
									+ ''='' + @sOrderTableName + ''.ID''
							END
						END;

						FETCH NEXT FROM orderCursor INTO 
							@sOrderType, 
							@sOrderTableName,
							@sOrderColumnName,
							@iOrderColumnID,
							@iOrderTableID,
							@fOrderAsc;
					END
					CLOSE orderCursor;
					DEALLOCATE orderCursor;
				
					IF @fLookupColumnDoneF = 0
					BEGIN
						SET @piLookupColumnIndex = @iCount;

						SET @sColumnList = @sColumnList 
							+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
							+ @sTableName + ''.'' + @sColumnName;
					END;

					SET @sSelectSQL = ''SELECT '' + @sColumnList + '','';

					IF len(ltrim(rtrim(@sLookupFilterColumnName))) = 0 
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ ''null AS [ASRSysLookupFilterValue]'';
					END
					ELSE
					BEGIN
						SET @sSelectSQL = @sSelectSQL +
							CASE
								WHEN (@iLookupFilterColumnType = 12) -- Character
									OR (@iLookupFilterColumnType = -1) -- WorkingPattern 
									OR (@iLookupFilterColumnType = -3) THEN -- Photo
									''UPPER(LTRIM(RTRIM('' + @sLookupFilterColumnName + '')))''
								WHEN (@iLookupFilterColumnType = 11) THEN-- Date
									''CASE WHEN '' + @sLookupFilterColumnName + '' IS NULL THEN '''''''' ELSE CONVERT(varchar(100), '' + @sLookupFilterColumnName + '', 112) END''
								ELSE
									@sLookupFilterColumnName
							END 
							+ '' AS [ASRSysLookupFilterValue]'';
					END;

					SET @psDefaultValue = @sDefaultValue;

					SET @sSelectSQL = @sSelectSQL
						+ '' FROM '' + @sTableName 
						+ @sJoinSQL
						+ CASE	
							WHEN len(@sOrderSQL) > 0 THEN '' ORDER BY '' + @sOrderSQL
							ELSE ''''
						END;

					EXEC (@sSelectSQL);
				END;
			END;';

	EXECUTE sp_executeSQL @sSPCode;

/* ------------------------------------------------------------- */
PRINT 'Step - Audit Log Updates'
/* ------------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getaudittrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getaudittrail];
		
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_getaudittrail] (
		@piAuditType	int,
		@psOrder 		varchar(MAX),
		@psFilter		varchar(MAX),
		@piTop			int)
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @sSQL nvarchar(MAX);

		IF @piAuditType = 1
		BEGIN
			SET @sSQL = ''SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time], 
				a.tableName AS [Table], 
				a.columnName AS [Column], 
				a.oldValue AS [Old Value], 
				a.newValue AS [New Value], 
				a.recordDesc AS [Record Description],
				a.id,
				CASE WHEN c.DataType = 2 OR c.DataType = 4 THEN 1 ELSE 0 END AS IsNumeric
				FROM dbo.ASRSysAuditTrail a
				LEFT JOIN dbo.tbsys_columns c ON c.ColumnID = a.ColumnID'';
		END
		ELSE IF @piAuditType = 2
			SET @sSQL =  ''SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group],
				a.viewTableName AS [View / Table],
				a.columnName AS [Column], 
				a.action AS [Action],
				a.permission AS [Permission], 
				a.id
				FROM dbo.ASRSysAuditPermissions a'';
		ELSE IF @piAuditType = 3
			SET @sSQL = ''SELECT {TOP} 
				a.userName AS [User],
    			a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group], 
				a.userLogin AS [User Login],
				a.[Action], 
				a.id
				FROM dbo.ASRSysAuditGroup a'';
		ELSE IF @piAuditType = 4
			SET @sSQL = ''SELECT {TOP} 
				a.DateTimeStamp AS [Date / Time],
				a.UserGroup AS [User Group],
				a.UserName AS [User], 
				a.ComputerName AS [Computer Name],
				a.HRProModule AS [Module],
				a.Action AS [Action], 
				a.id
				FROM dbo.ASRSysAuditAccess a'';
				
		IF LEN(@psFilter) > 0
			SET @sSQL = @sSQL + CHAR(10) + ''WHERE '' + @psFilter;

		IF LEN(@psOrder) > 0
			SET @sSQL = @sSQL + CHAR(10) + @psOrder;
				
		-- Retreive selected data
		IF LEN(@sSQL) > 0 
		BEGIN
			IF ISNULL(@piTop, 0) > 0
				SET @sSQL = REPLACE(@sSQL, ''{TOP}'', ''TOP '' + convert(varchar, @piTop));
				
			EXECUTE sp_executeSQL @sSQL;
		END

	END';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditTrail]') AND name = N'IDX_DateTimeStamp')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX [IDX_DateTimeStamp] ON [dbo].[ASRSysAuditTrail] ([DateTimeStamp] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]';

/* ------------------------------------------------------------- */
/* Step - Image updates */
/* ------------------------------------------------------------- */

	-- Create system tracking column
	IF NOT EXISTS(SELECT ID FROM syscolumns	WHERE ID = (SELECT ID FROM sysobjects where [name] = 'ASRSysPictures') AND [name] = 'GUID')
		EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysPictures] ADD [GUID] [uniqueidentifier] NULL;';

	-- Generic image update routine
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_writepicture]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_writepicture];
	EXECUTE sp_executeSQL  N'CREATE PROCEDURE spadmin_writepicture(@guid uniqueidentifier, @name varchar(255), @pictureID integer OUTPUT, @pictureHex varbinary(MAX))
	AS
	BEGIN

		IF NOT EXISTS(SELECT [guid] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid)	
		BEGIN

			SELECT @pictureID = ISNULL(MAX(PictureID), 0) + 1 FROM dbo.[ASRSysPictures];

			INSERT [ASRSysPictures] (PictureID, Name, PictureType, [guid], [Picture]) 
				SELECT @pictureID, @name, 1, @guid, @pictureHex;

		END
		ELSE
		BEGIN
			SELECT @pictureID = [PictureID] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid;
			UPDATE [ASRSysPictures] SET [Name] = @name, Picture = @pictureHex WHERE [guid] = @guid;
		END

	END';

	-- Add/update images
	EXEC dbo.spadmin_writepicture '7410CCC5-01EF-46F0-9D9F-9323A93B4573', 'Default Background.jpg', @newDesktopImageID OUTPUT, 0xffd8ffe000104a46494600010200006400640000ffec00114475636b7900010004000000640000ffee002641646f62650064c0000000010300150403060a0d000040a00000b5fa0001070900016b04ffdb008400010101010101010101010101010101010101010101010101010101010101010101010101010101010101010202020202020202020202030303030303030303030101010101010102010102020201020203030303030303030303030303030303030303030303030303030303030303030303030303030303030303030303030303ffc2001108017201fe03011100021101031101ffc4014e0001000104030101000000000000000000000704050608020309010a010100010501010100000000000000000000050203040607010809100001040201020403070304030101000003010204050006111207304013141015082050213122351633341732232536607037242711000103020205050908081107080b01000201030411050012213113140641516122321071819142526223153040a1b17282923320c1d1a25373240750f0e1b2c2d24393d33474b4d435b5167683b394a425753663445464e46586966070f1e2a3c38495c526c61712000102030306080a0706040603010100010203001121311204415161712213108191a1b1d132053040f0c1e142722333142050526292a2b282d273243415c243839370f1e253a37463e344600613010001030302050403010101000000000111002131415161718110f091a1b13040c1d120e1f1506070ffda000c03010002110311000001f7f00000000000000000000000000000000000000000000000000000000000000000000001f0b2d54d17aa82f9455d80000000000000e34fbf7df3efa000000000000000000000000000c3b2edc51276296aa6ff006aab9d1ed155e63b769e4f645c0bd2245e458b5c95b0eb32b6b86cdbccec7671d1b58e55f8052d1ed059aa8ecd74b8f729f1eef4e3dce18f716ae73b75547aacc9b57292c6ba4c60f3b9e00000000000000000000000350b718e84a73169a9f60a9b8f8a25e3ed77add5515c9b179b3b40c96454fb78d3a6b73792ecf45079d16719def0ae79b248bd7349997ea3e436bb5552daab951577f9ef67957ca3debb55d3e3dc5baf95bab9d15f2b7573b7573b75d4574e53b6c3d565da0000000000000000000001a13d061350b72c384e670aa6e476deeab973941e6df2c5cb7dca61d98c3d4cdab0639c9ccd8080929a612e7a59ca67ed9897605f8ffba661bf6b536fd13cbabbdf79fa0001c7c5ab1ae63d1d916ec3bdcadd5ce8af95bab9dbaae5996b2ddaa240000000000000000000035c76788f19fb56bd82e7e24b11f637ab4bcffa0007c354f66c2d55d8fd92e333fd30e6129bd7cb36d84be7ee9b3577de697194c40000001f0c563323158acaf98f7795aaea3d5c322de79b6448000000000000000000007849f4073fd72d9313a7da76e75ab92cea1294d3383c2fdb1c0f9ef971d625b14dae2b46b6fb1f3df33a8ecef7abe3feb19c4c61dc24f100000a3a56cb5edbacd7458f72971aef0c7b9864664d362ddedb177b71ee5eae79226d7157cd863b955e0000000000000000000892760fc88ec30daf7b46a7d56e4f6c3e79e8bba9f9c1f4de19b7c06acfea8fc87f2ba79f9eec57c97da658f8f3b7ebafd0bccfcfcfd14f9be92ee36c3eb733e94f3492db4d236c0074bcb2a8e8f3deba3deef2aadf552f7b3d3c75d3ee0b1d7b0789c9a3c2c8edb17331c8f2f352f72b8b91ec51a00000000000000000035e769d7f5536f8cf35ba542d5d146f17cafd6f707e43ed18bebf21a8df6e7cff00b3bf0afd0917fd45c9b3fc2ca9738d6f70d6d70de5e7e9b7cab4372e6f7e8d9db65a6ccedd691b2f47b4da6bb5f3cf7b3caaed4d7d8f400001836331eb155a706fdba3f2738bee76abb85ea736d9a300000000000000000028f0723059b8a887638ef127b8e9f5fe59dd7e17bfefb7c87dae1dd823638c8b1287cbfd5ec17ec42df5e71adb1f9d7a8c11f43734f3c3edbe0b64cdb7ea272bda36e74d94cce372aa28aeef6ef80000001c48aac5ba7b755579556635fe8c7b9cedd79065dbcb65b1400000000000000001668291e8a51d6f501af9b260792dd834d9bf9b6c9eb6fc41df6f1371f8e5aaf5b7a5e9f34fcbfd8618e8da97ce63b3e6da06cf7edb22f52becce1da4df5b71af47398edfb5ba8cbec16b99a0000000018c2d47f4dba9a6a92eabd8b635764c6afb68aa4793b35773c00000000000000002c1ad4b473aaca5f3b1f3dc173f23cebe9badeac6d909b39f2ef55df6f8e3b964bb743461deb99ce9ccb71b26bd2962d725209ea7a566bf2d75da28cc889fe98e5de7bfe90fcc9b45abcb6eef3ed9f7274a930000000051fb4c4d770bed3549f6f2ae8ae2ac7a29edd59c5dab26bfe000000000000000018ee56242ff003375dcf37bd7a4eea9a7629976ff003bbf476a17cbdaa5c6264b3585cdcbe8b5b3ff003df4ddaae1fd26ed351f6c8bcb8876189d0ffb2f81e2bedfc4e471f18d922ac94cefba7c0f72d83d77300000000e0f227cb8ca2254c590ba797317f6d47f4dacd29bf9a5758000000000000000a0b96a3f928abe63e4ebc7c23f495eb628d9b3e81e6993edb09e3177ad5b5137ed1f34c3ab2cc6f730c6a24dd325e44d76fdcf2f07af23dc3efe5c29bdc1e21915e2590c5737173d81dd7f41df3ced35747a0000003e11767c459ebb52d47cc5c29b96faa8896fe0667672b37a3200000000000000038bcc164632d77ac48d172f5545cc4344d823ce5db77764d1789f8ec6fb069de05fd47a2ca1635aa9a2bccb1a8c92dd9c8a8c5b9d16a8fdbd8fd7958bddb98864d7495d30d674f7b79c2f68dbfd3a400000007c791749c25aebb72e45ced5535d37b4445210f9962e7e696330000000000000002cf7b1e369587ce3024b29c3ce00000690ef519e32769d7b6061e37b6ac5be538f7ab74d5f94d37b55a2baf1db97ebec656bbcd677a3bcd673d50e513800000038bc8ae5e0f8f9e4a91539d9e554f551114ac0e7383299663668000000000000038bc8fe4a26c77f1a52899caca2e00000074fb4e8eef905e6ff00458f95a26acaf1fcacaec76556f8795d2517ecf557064e53e806852be8773999e40000007455444b35059162e5e7f1b2a28abb715cc41489172f7db19400000000000000b6ddb314cbc0e53899d9f474b7d0000002c191898148c5491173362bf569ced98109c959a5b7ee41895da3228a4c9a65ec0bdb7fabe6ca515900000002d776c45b310921464b6498b982cd7f1b00928a92a2a66baddd00000000000000623971f1b4a434a9133b91e3660000007c791f48c4db2ed89422a73b7ca8018ce4d18a66517cc7ab35c2b800000000c4f33030ece8f942266ab6ddd186e7475baed9902325b9bd0000000000000383c8d24e1718cac29861f61ba5abe0000014d55113cb40df6c654891d2e000000000000074fb4c712b0f5f6eee7b1b29f5ed355460325159262e66478b9800000000000000a5aa888e56029eba26288d86b28b80000016cb966229581ce30e4738c29300000000000002c1938983c846c83192b76b390318cbc2b35fc6ce63e53be9a80000000000f093e85f9a721c6cad92d576e8f64e2b03918d9d75dd976c747de70a908d956366a05d8f59dabd3f76b5deb35345566bf63388f91e3eb09cf8eaeb7729d4c1d33af6c36bdb46718923ac3b669be7674de51b15ac6d9d755398e0c86d769dbbdcad5dd2cdef9eeb76d3a8fa93c8fb448d172de7f748e5d16cbc2edfe93be6299985a9fb968fe9b728ec55f6eec77271122464b6711f25ad1b56a1e69f54e4160c9c5f4339a755d80d6f6890236570ece8fb45eb122464af93fd8f887a6dca3b16498b9767bf63b3cf7cbeeb3c6bd3ae4fd8f3c8f93000f1c7b7f019ff005cda2468b96d72d9f53d5adb74cb858bd9ce048ec9eb3b6d4535c6b290faf1b36ab5d6ee66385212cc34e60d211b63c8c69db5ed978d154093dad47d330fef6fcedf4d81e29778f9df29c3cef4cb95760f347aa720d7ad9b55de5e7dd235ab6ad4253889a90e3256119fd7310cdc1d9cd4f71caf0f36279984d62db34d9d35fd928ebb798e0c8faa3c87b5757b4f8a7de3e78f61389f7ac87172f4a77ce770ecdc044f3309144c41ed06a7b9e5d859da85bae899561e6ed0ea5b9ebf6c9ab53d54e698321e9572bebd2146ca800783bf43fccfb43a8ee78567c75c6c5fd5adb74beaf7c9961a7b62f59daf35c091d44dd344c373e3e4e8999aaa2bdadd3b77f3bfa6f29cf63a4e7ad73679ef4bdd3caaed7c2fd54e4dd9be9b61a76f03cb7eb7c5e319687f5678ff006df1e3b6f0398e127e00d8f57a5ae8cd7024764b56dba0fd835c8de5222498a979621e7306908c84e7b5df6378877df01fe8df987f44bf32fd5bad7b4ea1b3fa9ee5a59bdf3dd61db34db65eb331c1cfd3d54c4f33071c4a44fa21ccbab407b1eafae5b3ea77fc7caf4cb94f60f3c7a672ae8aa9b35fc7dfde71d433f8e949f35cd9c0117cb42d3574add718dfc1b4e4634a38d9f2a454cc7329139fc6ca637958978b37ed57ac7753554535e3795899bc7496bac8c2e7b152be36fd0df387b91c17e8cefa6acbf0b3c59ef63d1dcb75d6eee159f1dd155378b19160c9c5aeb772b2ddccbf0b3e3c938aa7aa8b6ddb39163658eca7df1a3b9fcffeeefcf3f4b6239b8397e16788d6521eeb6afe6b812313cc41c851b2b1fc945e498b974f55166bf62ae8af2fc2cfc7b2717b69f686e5bba5abd5d6ee6438d960014af229b987415dbc9fcbb2459cc000000186d9f30ff2df1ae8982e5fec7a00000000029aaa38fbe55d1729eaa3b7cab9f9e800000000000000000011a5366c555ae355328f99579f2b000000f8461197696ba38dca25accf7900000000003a7da696aa3b1ed5d15fd7a000000000000000000058edf91ddaa785747dabc986e5ff00a000000535baa3a83cbebbb475ddb79a675bc8b22900000757b4f0f7cefa6be8f69b45cb3dbe7b57e575f45c00000000000000000000011fe0d764a7ce35d3d95d32ce5540000003853ee03aa4af45da3aaf5be37ade459d66ff00956eb6e79f40282bb546f28bc52fb4f77aae7b7cb77b93d000000000000000000000038f88ca172baaba785ca38574c979d457d600000016c8bcac675d92a7b94755fb7d57ed75deb7f2f5be8c9b1477ec54d7e5dfdaeeb52f35557455f40000000000000000000000001d74fb1ceb723c2e5be1768e1729b8e45bcf73ed76d40000001c68f68b0afd35baac772cdaefd8aabb4df2aaaf8afbaa00000000000000000000000000000308d565686dd5c2f5bebbb470b94775fb791e65abce4515973cfa003e146f2d8a2d35514ded1df52fd45dbcd377e800000eab35d04665dc2570fb2e5200000000000000000000000a3c3bd8aeab2d4cf3aef51d77adf5dda385da385cb6bb4f75da39574f0ae9e9aa9f95d171b9e5e2af6f7ed55ea80000e34fb658092b6446759a0a4719d466b08e75b3e01cc36ce8c1b932f7de71b0df547230000000000000000000000070b755b2272e8312f53f9ef5554f0b947cbb473bd4775ea2a6ed159769b857e5454000003c7463ddeab55f4d8b95f238bcee5200c534c9ed5ff8fbb852475fdd6fd05f9a800000000000000000000000000000000000001c68f7955e3d0000035a7e4eed396ee9afcd3df79b0000000000000000000000000000000000000028a3f22d70d9d90ed31200005875b958eb976e32f76fe79db7a80000000000063d9389cfcf6fd8f9400000000000000000006b57c85d9e29e25bdecc7d83c5a53ed1a33d010d7cff00d17583e35edfe827ea57c960003a31ee62fa7cee57ba40765ca4000000000003cfae97ca2ed66fef7f3aea00000000000000000002970afe85fe657d51c68aa45eb1a7eddfddbf3f08cb8f6eb14710df2d90f9b1971dddb79bf48fe6085fe7be9134fd0dcdae52f85027cc5d5a5aee7a0692f74e27adbbc68d2b444e7a5fc9fb2c612f09ac1b769500ec9ab4b10d3be93f2aec56dbb67cdcea9c7a2b9883dede77d3f6075adafccceb1c6702928bf4639775d9cf5ed9b55371d1b457a2731c064a2f7939ef4bdefe75d40000000000000000000439c0ba26ae7c61dc7653ebce31ad5f20f68df0fd31f962d90d9da41f9cbf4d4d3f42f36a5c3bf12f0cdff7fbf50be4fd0ffccffa9a79fa6b95e4bb7436a97c4bddf7cbf4cbe588677bd2b3e8d94f343abf1b9a60762cf63657ce8ea3c8bd81e21f4078f7dc3e7ef5cb8a77bf3d7a6f279e35cda36eb49df61e9cd7f453a2731df0e75d426b80d8fc3ffa07e6bf6c7817d1fe227d01f367b71c03e92f3e3a672897e127f7bf9d7500000000000000000070a3db542e7ea8fc43de319d3a6e61ef7cf35f7e5beb5b31f60f15b8cb626ab7c51dd7d01fd46f93626e1dbeeb07c6bdbfd04fd4af92f5abe42ecf1e72adbb26dc2179dca76afedfe09e2977cf9c7be9ab12cdc0da6d3f789b75fd935f766d4fd46e43db7c9fecbc2b72346e85e45f6be09fa08f9a7eadecf3d1f9e4fa6fe4cd93d577118866c7ef473ce9b086c1ad7a7dc93b579f5d2f945dacdfdefe75d400000000000000000e9b17313d2e7ebe4313463f367e9f953b5e8b7fd9a2a3be53b7f3bb4ec9fd73c6b4dfe02fa2b66fec4e2987e87b0c3bc0ba27a09fa95f25e21a1ec1a43f9cff004cd464dadbdfbbfe7decedfcdbc5eef3f3afbcbf39fd43e6e754e3d4372dcdbafec9afbb36a7ea3721edbe51f64e17b89a3f41f393a8f21f47b96f5fd93d537189a6a07c8ded7c17d4ae3fdb760f5adb21e9cd7f1bcac3f33bac719f6db80fd25e57f61e1f30c1ec1bdfceba80000000000000000e16ea8af8e6f79c743d6318d2e6a15f9e3a4ed17d9dc3eba471a3ae4fb8427f3bf4ad9dfb1b89c07f31f55c139a6d39cf48d5ecdafc96d57dafc27ed4d68f903b45046656d27d9fc3b957e7915daf81c373baf4c707b0ed7699bcc8f1731154c41eeb685d1f46ba1733d81d6b6bc8f172fc98ecfc22c9918dbc5cfba5ece6a5ba793dd978558b27166681d8bd5de37dd3c9becdc275fb66d526781d8b7879f74bd9dd4b7400000000000000011e730dbe2ee3fbd6c8fd59c5f9d7e00001c68f47dabcfbe80000000d44ddb419e35cda3a2aa31acbc3fbe2fd8f95d35534b5512744cd6af6dda4cbd09b051dcb5996048ebb6d1a94bf073fdde55f0c93173330c1900000000000000001af7f3075cc0b9ced53c7d1fcb248eaba67df41e2823b2b8d3577e4d9efc8b7ceba4000000003cdaea9c7722c5cca3b96ef36323979ed3d5459b231f9f9eef0f3ee97e7d74be512bc34ef1794176d596fe3e7d1d29475dbc6b2f0f60b59dae7bd7369000000000000000187e87b0ebcfcb5d76971f2326dc61335e83ad649b5c35d66702fdb145d748e28000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000007fffda0008010100010502ff00c7c9655e2cf9d55632d6b498d735e9e4b94fb9edafaae91a4d92e27aa246260a6420e7ce511096114b8e6d7f2dbbbb81955b4d55a95ce6b1096b5c2c26cb56ccfe5d5d91f63a890a8a8e4f8b8c2663a7474c5b14c59e55c59675cf58ab9d4ab9ce738847a6364953192dab88a8e4fb8368ee3f12a29851c9376e830593fbb8112c8eeb6cc5cff20ef4ec0f74f6b0ac1eef3d72bf78adb16ca9312c075fb97cb4a3a4a990336af0de9634b26063859add9122497da331d6121d8bee8b89164ae7b3939eca4e7b4938b1ce98bca7c39f8739cfc18470d44442b7cf77377e229db3e1d644b3dce4c9ca3d2afb695aaedcea75a91415b053de61d62cb6d9e8ba85a25ef6d6d6a32bb6cb082b12fa3d88343de7f88cc2d8858e2ace9832c656bb5fac43c96c48cdc46b5bf6d5117090c2fc2c6287ecf3919fd25f3bdcddcff87d036c3a52c271663f47d7016323dd719eef3dde7bbcf779eef3de66fd411ce38920f0cdf306919d95dd7e63164190011432ca2c58c3881f0e6476b73d61a67b816248162158b82fea79dee86c3236dda546f03583cd48ed1d331d2088e92e6afbbcf779ef73ddbb1124f4da496ad6b87ce318e2b75e9f63aedc865b6fa2c78c38ccfb6f38598eb10263ac9f8b3643b14a47673876fa65e739ce721abbd61cbc45454f35bcd93eaf55f93a0d9315492d133b6554f9e6783a72ce00e60164bb9f5f3dc66a752c2c520737ca871a995339f4c916b7dc03b4131eb43f61e460d1d621e7d6b02e7b2945c6d6093120c64c48b1d33d00e7a01c58c05cb4822ea78083ce739c88ce81f3910bfabcd77006e960b4a8f420f57acf44ced0cc8fd040b9725343183aed625ec804788166cfaf45743d16c413681e2e737490281ae2a63d9ca6970be634ba4457575ee3c83127bdf533d39a5c6c08e8ad6b5a9e0593b9361e3a39228bd57f39ce46feb7992151993e2b244dde21c51eb23fc188d55cd1bb71744570151368d7645f44d1e3a06a881c9654045d228a4c3419b373d3a7ecb92e2488521533b331405d5c71238ac7a671b070a3b17c35fc30c5f54bce086e33cd1db149ce739019cbfcc499482c0a35adda6c842ab75a2585603fa7a3548adf6311f3f4932c0b1a046ad8d6d36e081cd9ab6d2682aa78258d85e3066ceead60958b9a24f6d776ff004d9e93cde2d948f4c78363caf8b19b19962bff00e8e71bcb9402408fcbcf989159ee7856409b3d77d88d06ab517e83b0783dac8ed82745b0a4608f965ff3fb2384d54207081cd8e3fcba4906988f7b737e4697592af431f61f2daded1aba443f10e6647194ae31043799f162b2333f2c317d52f3902374279797651e2e489c4987a1f4dcfcbb8496153626975565324325498b224465059d9e06cacf34123dc411f3f4930a24cdae3a928493a7b50d656592274e3306d6125cdbb34b97db2a87d46a7e1b9c8d49b316510012487c58a38accb293e98f20c2e7cb9e486334b61327be1d38838457c63c3b0245346bd82746c98eecef86adf2bbdace24281cab80c0e565a4daf41ecf72989b55e63f6ebf4c9fb45dcb8e6c3e1dcd6367f545076e35a36cbb3802c8e1f0ecac3d758b18b29f1e30e33324c8646110ae2be057f96554449972d6647ac9131c20880ccb2a485678ba815b8dd5a4a60f5a54cdab4a8d7faeda56cba5b187ff003c18c547a05d832e0dd9eae3cb857e19d9208d6b7d21c48c47c9b299da2d21357a4f0ecacfd6c8304931c108c0cc2958164a96f965815dd1e5a54e04469244db5242ab0c5f07bb3db46ecf17899572e158556ceb2a3d85498444c6973d6c717084c47165c89a0aed6b2c6c67dd4cecf76b55ebe12aa352c6d14f95f56e918d6b58dc2118164a986b0340ad6c6f2aaa88936e78c895c79ae08051d9e0bc8c1377aedfd56e4697aac8ad9b53696f5806d56956189a2d8133fc7fb967f02bb663b5fd420e4bb99a91968d5e6d03b36288f44444f048460993ac4935d02a119f139c518652cab63c280286df29265062b25cf9135d02a11be1ccb00c446b265b962c40c465d5053ec316d7b4f6b056543975391e56b327122d4ab4e4d723a4750592d6f6cb66b45d774da0d61be14a96188c2165da9a0d70a227c264d0c360c32edca08e28ccf293ad471b11244e3c1ae1c44f0a75c71906b1f2958c68dbf625d2d3d82ff07d2b22eb7aec1f1675a0e3647852ac8808e28ccf84eb568722d590ce4446a7935546a4fb757e458859648b105107e094a30b2759925ad7557947bd836cbb52c8742a746e7e5f02946161664ab27c2ad0c4f2a630c0c9d60596b0a09263c201c71f832650a2b254a34c7d6d5a0bca4cb1044441ceb77c584088df84bb3147c1d7c99af18d826f8fdd989b8e95b0ed5dbeeeb6a143a0ef5f21ecff00658d32f2eb6cda0ddc4ee2b8252770eced46d652ecbabdeca93b06b0fd77b3159695b1ae2fe935f052edbacec6b63b46b34f22c778d3ea475f655f6d16c6d2b6a2357ef1a75a0c9b86a418ccdc75994336f1aa48256577a7f0ee8ea9b1ed355d8e26c7b16e5f50a3b8810fb39577d5103b45596b0aeaeb72d575d343da35ab0288322d7bbbdf506d5ac6c66db759a61c6951a6836519aefb97dff89b2d34d1d90bfc4fda711b5fd6b53dd46def44adc7508265d9f5b4ad66f547732e253b5b9f967743b860edf51c4d63bb1dcb8d721eed76e2b3b0c52cdd12d37ad3a964c19f06ce30371d4254976ffa4327ceb2aeac8bd9fdcdaddea4ee1a94394cb6ab258dcecdaf6bada6d9283611f7f22ecf456dab1086d63ecfd4d7efb4bbaf7be55cf7bf58adbbd37b0da3d2da6ad4ba3504cee88b5f83ac778ae8365dd5ee6dd7d396beb5b4fda17d676cbe9c660abe8e8e7ea9bf6c5b9c3edad28fbaa283b4683a2f6375ed8751ecbfbed5bba5ed6c3be5dcad97e9cea5f0a776753fc73d8e8f52487daceddd4f70363f897ffe67df8ef7cd91b8770ae34fa7bad5be9bc6d0ecd04f4517b9bdb7edee9a2d9e9b5dafd4fea07ea83283e9f692555ebeeb6ecdf74768d62b755ef6fd4bff00d52fb58add8fb23d86a587b176f35bedf50db775fbe3a053c582dafa0ff027607b7517da63ded1b7bf9292c7b810e9e34669c00922ef56d45d1f51d2be9fe8a6d0578acfb27dd0dd7468d57dd3defb0fabd269fda7af17717b4fdabd0a8f69dd3ea075b81afec9ae693fe27a2edbf6d9fdd37771bb7e7ed24aefadc8f62d0f52ff00aa7d9efaee153b56c3b8fd40d14cd724f718b27b2dd8edfe9a9353a5de6a61772ecb75a697dcfb92d876c7b837bf51fad06abf9f5e43edbf602eea10baec5d67b7bb15c6cbda524fef85a53556addafde294fdb8d577da3afeee7b9b2ec6772765fa8ca71c2d6676c7694fdc20cfd1b70ecbeaff00c6745dbfba60d4b6ff0087d4aebbeac1ec1d5cbd9775d9f62afd4e8fb17b5d5d36e40ee8e85bd5be9cd8317bc923b83425ef3fd40ee153b1d85077fea2056ebff36ef2774f7dde296cbbb3df3ee0d0ed14ff00e4ed6ffc2df4f1b7d556c60ecf5dad77892aaa362d6bd0d994f535b168ea75eeea3363dd92135ceef876ea5ee153a6fd404cd7e06d5debda37926e9daed9a676c74afa80a285afd792cfbdbdd0dab7da3b0eee774378a5076e3e9bb67ad657f6ef78a5d1379fa85d9eb6ef61249abeebe81db7ee43fb58eee2f700fddb97de7b1a7ada9d477fab3f6c7b75bd8bb8349f60ba469720bfc0b46c9dace9e2821d6b5c8aefe17a82e44ede6942112be01a140d0349ab97635559711e0ea3aa564ab6a2a6be0d36a9ad6bb963abeb3712226bbafd7c56e89a3b56ce0d6d84385a8eab572b8cdf2eaf76ad935bb22dc504cd7e86c667c2757c0b48d594d514a29d020d9c58da669f0a41f48d36445a7d675ed7c3fc0b46c91a669f30b67aa6b172f8b122c103f46d24af99a968eb8ba76bf2a2d4e9dab519a4e99a7cd91b1d93b58d6755bdbc77753236bf430a7fc2e34ed57607d3eb741afb32d344d36ea5428106b233b44d1dcb2f5dd7ec22d76afacd3c87e8da495f2f4ed427c8815d5f551ae759d7b626d36b941af0e76a3aa59ca8dafd0c3815b51554c0fb262b403399f2099571fd43785666ce338c6a74b7c6735eec18042f87a88b89cf9c9d25641338c802f4a3784572909c67e4a8bca78cf23069eabdf9e973e7679bd20f1f06b7a9df97846fe9719c67190242399e039ed6e7eb77c1e518b3dfb1ea8d985c1c710d7cf4e7fa87e3e03fc09e1393a91cc56af19c67e28a1b054c64903fecba40d15d21b8b610c59f30926cf42c4d83ae8cc544444f3ebf863bf15e338ce32317d50f8478fea6386ad5e338ce338c45737093d419f309c5c404f3e32b64aa32a598385187f723bfd3c6719c67191a42c77b1ed237c25445c58c15c5881c22c01e72a6c6d414b81aa811f1a3637ee930fd32719c6719c60ca40a8ec531b200ff00b6e900663aca2b71f6f887b393895662a8abe28bc273d8ccf7b0b96b9ae4fb80c242a3c6e62f19c6719c6719c622aa67ae74cf752316549c539d710320b8cac90ec655093071802f0155131f611598f9d3df846ec86c353df9b24ebd683c2015aac71e3ba8b6891ebfdc0ad4723a2317161931621b3da1f3da1f3d91f3d819712b5d895a3c6c18c98d1099e175e2a19d9ed42b88d46a7d8b3a98f6222c751bc719c437de9c7853d1092f5eace4df7a1e4478cdf9b552f816123d0042ab594f631a36f90b2b6aaa60575a565b83cdedd7129961f37b6cd4ec9f3ebbec6f1fb487fabf69ef41b5233a591ad6b13c8fd48ffd1fe9dbfe83e68c66470ca90f972491ca21ea33bda5afc2e3678754e5de2cfaae7644b8af0ff56ce63e042fe7995b3996508e66470ff3cca6b27db43dff00bde1d436077d4aecdd541f52708f221cc8b61177cdb3f84eb545f505af4eaab1fa97b6748d2fea1615bcfc97322d7c5d8fea46b221e2fd4bdeb49a0774e8fb80fb1b281510b61fa9308a444fa96be69b46ee0d0efb0737feedebda22bfea5765527713bbe0ee06a9f4edff0041f35b8d87a15c31b8a4da2a983a363dc37c094d9d0eea7fcb6b5ce73dd4ba8fbe8fb1eb71aa2287faa4634ac9f11d06668d33a85b8ccf6d54213ce5871990e2cced9e99677ff20d66187be7ac68c1a7fa73b334bd33bf1ffcd7b35db4acdee444d0f4a851bbe5a5d5ea5b1f6e2d0d75a2fd466df21f3bb4fda4a3aaa39fa8ead680d43b5953a46cbf507b5cbb1d83b7bdb0a2d2eaef35da4d921d1fb9ed8f77b6dbe1eb1adf6df5191dd0dc20e9faad6c6fa80d3f58a8d73e9dbfe83e61551a9ea38ceda65a49b5ab962813a4ee91e547cd2277a91b775ff89c8e88d06f1fb487fab9bbc1e83ebd37d8db6e533dc5a69b07dcd939cd1b772ef1edbb75cc4ec07702eb379ecac9d1f5cfa68fd8fbf1ff00cd7e997fb1cfa9d6b7abb3bffcd7bc6aa5ee935ad637e1dc76ca93dd5ff1df7f73fc77dfdc5eca774a55b77f084676e3e998624abcfa91ff00a3fd3b7fd07cbbded1b54c49649b219575ee72bdd5ba9cab187fc165e5cd31e98baf4ef616db54474ba6ca7dba0244d9f6187691c3fd5cbd83f31abc295e626ab07d954edcc904d4fb093aae0f7033bffb76befd6fe9a3f63efc7ff35fa65fec73ea773b3bff00cd7ea235e340db7b5dbc40dcf5a55444abde35cbbbfefd6bf2a9379ede6ff57bd52e5c6f9abd259775694b7fa07d3f6df0682ff3ea1b6fd7ac29be9dbfe83e59ce46a4a9eb2495e1e806c95d3ed22ff0ab7c8c06458f9b25392de27f0bb8c86d3a44b8d34dea8f56bc7b89a49db087a65bb5ff000b1d3a7967034ab2f59111133b85d82b45b18ddb8ef6cb477d38d8075cece68777a256f73b59b1dbf4fecd76fef7428d9de7edddfefd9dbfa19bace9fb3eb151b7545bf60b79a39a7ed9f7ceef3b4fda697a049da356a7dc2a2e7e9fb74a698fd0bbfb31ba3f602eab6e73b89d829c7b08ddb0ef2ca64bfa70b40ebfda5d42d749d5bcb6c133d1134b888889f6f944f1a76fdb30c957b35659a4adc35a850ddbad6ccc2f71f480cfb9dd358d7e4da6f1a9d32d86fda7d5e365c7244a8dbf7ad82b2df77d528661b686836525f5488513bb54d6b476bbbeb142b69bb6a94ccfe6daa7c9692fe9f6389e5b6a7b9962c3e545d80c2fb0f931d988757e7491711113c693414e9b056c4da35395f2c904a2871771b44b9d7ad8bdb9b9a89d59b7d8561a919702d861beb9f663a0b2aa892d927e65ae5940d4ef636564cb7b62c90dcff0a47cfd4f66d3f57b6a89f2e95f1a276f4bb19aafcb5c5406dc07d7eea2b990ad32336f0382fe4afc644ba7e36a9ab828c00ffeaeffda0008010200010502ff00c7f85ce97674af94e53ee710486cf6cc6e7467a4ab9e867a39d0b9edc6ec24628d1cf6b10b73560c2edb4e3c5de2a9322edb45255151c9f1ea44c52b714f8a77e29498af7e2f39c6719d4f6e24a3371935ab8d735c9f7045adfd2e6aaa3633dea3a633f1b45133e4f598b470970948f6e3e21478d6398b69acc1b7c76af59c9f4f8044b4d7e5d6e3e3e69b61223ceeacea5cfd4b9d0b9e9bb3d2767a4fc51bb15b9c6719c6719c6718c7bc4a22a15be7aaeb93a546f238705a985951e2616c25931dea3f3d3c6a39b839b2c581b011b090c6ec74770dd3e07bc6718608ce2346e87eab5cae9dc2782a262e3c4adce338ce338ce323bba0be76aa0fbd90ace71834664c3ab1bd19d19d19d19d19d19d1908ce6e3dad7a7a7c2de41f4dd28e9180182598687105081e1958898af6629479eb0b3ac6b8c4fd7e76b632428899ce494e4b26ceaa1bc6e1999d39d39d19d38eb8a661469faf17f230c720569164926458a3883fb7ca675a6291715efc5ea5ce31ece155b8adc56e0ba9af1cc54c45454f355a1434df595eefc933b817a7aa83240e7bb59d9276b53dbd2f6f19c66ff7727d62c7cedbec72a05b637f32bbd325e313d7fb7c674a674a674b73a5b9d2dcf4d9850b71c154c56e2b718ce138c88f54779aac5e8c8c5ea37c3b9104a452c7cf979651b75d98faad6ade59c83695bb4bf7bbbc67c5d81e245cd4eac93365f85893a0b39feac4cfcf3a73f0ce7c27e718f1f388ce738ce3049feef99299a3c8cf7a46ae23bdd7c361d9207a660fe343680a299dc512c9b6346cab1c92596fbb6c19124523354da6052e47901941cbb7b92475b943fa133a97c55ce338c73385e338c037f57989531059100e76451aa9503d055cdaecdd554a1939cb088caf34b35b7caa353163e6a12a86a666cfaf58534a149e303273b696ee23b278d493e6b3d36f8fc7c1ff009f19c635bd29e5ec672446558bdecaf5063caf7f54b7c7e5a8bca77524fa352095819395a6f96d13c48fc2830a0cd6dbf368278aa8a8f7897b672d7f93e28bac971fa5fe2f1f61738c6b7cc304e7e584f4933b5c2a39d807fa651f4947d3d29261c398df9353a67caaad337e20e2a064e7e8221419acb7d1d856a6a9cbf24a55c8f575915e89883444b437ad2fc446f18bf87d8e338f2ed6abb106c1a3ceab857122488166486789b157496b65457e50cd6c88dcfe0b8bf0b3a2aab87269daeb71355a16e7f1ba5c06bb4f18f89898dfcac64a468ae5572f86c6718e5e9c55e7e089ce718be5d80e71c56b31555df0b6d76bedf17463b159a74c4c16a8a99423f9238446986e5e855c5ce3ed2673cafe0d4b89dee8fe18c5c63dc8dc55e7e089ce2378473bcb318e7e2358247995fe0d459fb677e9235ed247c63c654ce3ec7e0c46b9f2318c68db736bc7883174e10bd39f9fc111555ac4623dfcf966071e56b315cae5f0511572058161224c47b0be9117dd4b1e7ccc699f35859f34067bd94fc6a33a924fe16174ae4f09115546246610df16b55ca8d685af22bfcab58afc60dacc21bc360d5f9fa028e7abf0320d1dc3b3013113d5c7326b33ae4635263f1465663e7c50e4899224f86c62bd51181690aaff008b06afc573028e72b97ca304aecfd2c47915fe1b038f2a373f3fb2c31879efa6e3a4c87f8a30abb1cf6891ce572fc061e71e644f2c31718e7a311ce572f8288ab8c1a37085f2888ab8c0a37087f8a22b9518c163caaff2a88ab8c62371ef4622aab97c16b55d8d6a37084e7ca304e7e7230a3dee7fc582576291ac45555f214cf853e343b2a79b22c6bfdc5ddea323821c44acacea46d6084aaa689323b591a5249bd288ae0c73c971e1ca8b838928cd1c09a652088170c453389026895214c573a1ca1e257cd4c7bf9f854cc8d10d7e91a341d694247de16398974513c0083324a3e24a1a2ab434bafba24b8a90e51d5cd731d155a0aad71f14e3512fce2e15b2654c82bf226419a44f692bd55af90063cff0a9ad759487cba7aa70169ad0bb12232c055f38ed20c8273a0cd6b7e5b3fd3188857ddc1ff8f6c298f6284a8204593270f164465d71f124065a224bfb3aa7f6e78340c0504b28276c53ce29722c640ea16512651465153d4836793ea9ee90b6bb3b1492240e6574682fb43ad3a92258d86c12634dbdf4e5d4758f5faa8bb3991edbde2cf62799096f666ad8bf14ff0095d7a818d855a19a604bda1798a46c87d55a594e588792499adea5923643b4b27d1bca9892cb328755fef23cb2c5bfd88ef8d652aca4069f5fb13b89ea48fe45b1d9bfaf38e735d67a55af3b9d8d739ab451127cd9fb248649228afaa605838b515fb0cb3cdb822d65cdbd89e241d6e5124c5953fe7122d2d3e5195964db966bc058d6333fbcfb3afc23438d075b90c94dac46deec15e734c915c72548abceca7020ed6b63eaf294df2e03ed36301b892e956518116e5074023965db403b6d25d71c94bd22d82ae2eb0657cd645148aa70aca05ecbf75610aa1d3217c35593c13632b22c08918930fb04329e0baa6c6bc337d4751b6b642516b708d147235c310b27d1a3a8ae8071536bf5b2221be532be7bb2c2315cb1092a93d634695d517a4c57c8349a858b03af282cd908d3b5b649244a1895e906da2b2da7eb721f24882a1a9895c71d2d4c03bad3688855259c03d857eb510a08c8d2d3d8da55fcdf2b2b5b4cca2118a6995a56dad9d7ad69fec24f9cd4f98d8624c99d6e992df9f309d8b653f10844792c6795833142e24d98568641e3b8f32549c1cb9416be4c923fe613f065205c5b09a76f39571a3c2892848090c93204cf80c84138a731d464209ce9d35ed6cf9cd79a54992bf31b0c6ce9ac414c960473de472584f4464c9f9efa4b1e69d2e423674d63630bddca991e3fca31d2643c7f004e99190d2a4495c1584e035e4215df309f8c93246f24b9466a584f4464e9a36908433812a4c6c3ca91254736609ae9321e4298c77780abe1b9739ce7c87e18ae55f871e7f9f13ab95e739f21f9e719cf9d72f19cf8aeff004a3b39ce718ee7c5e9cfd299caf9f72fe3ce739cf86a9ca72ad547675623b1a5c47357ecf19f867e7f0fc33a97ee2e739ce739c45e53c2383d4c5ea62a3b11d88ece711d88af5cfcb3ad73f15fb997f2e739ce739c6bf8ce79f0d511d8b142b9ed079ed133da8d33d31373845ce13ee92a7a6fe739ce739c472a62173a9abe0f19f8675672be13c831e7cc6bfa9af6bd3ee03090ad72386a8eceace739ce739ce73a973a973a9739f11551109670d8afb1b27e15bb51f24516ca7c97ab5c090b15cd51ac98afd6b7197ee3ee0735ae474262e7b42a67b7367a26cf44b9e9133d27e7a4b9e9e742671e17a898a877e7b302ab5ad6a7d8b7a58b6a23c470de286f2c8fbd3845f0ac9ad2ced62a3a8ff7a4997161b12f2917c0b395eda3d7d43a63c6c6099e404131dc50942ef37bd5fcc1da7cf6ef347b6259d4fd8ee3fec80febfda211a36b623e695ad6b1be4755fdc367fdcbcd4838e2c79b2893a59629822d12cbd8ddfc2ff72814af5ee45c756c1b725fd503faf713c9575dfe4dca8b21dbd74938e2c7ff0026e6bf6e4bc815fac7cc829aa44e24eaa46b5ec78df5d0fdfca91ad491985aa07a67eb2f00f18c791f17562bdafd523f1654f22b704221891b55556bf548fc5856c8ae2656d2c9b1c4d522f15948ead99b3fee5e6b7eb548d52113ce5dce9182d6c64784959359650361b3f94543dee23b5dd17e65136cd4a1d0c207f5ca361876705f5b61db7b0ea06fd61ed29001248340863af84cb69e28dee65bd75e9760a7da04d64fd77f75bdb52d735f633c8ed7a7966c5b40a02c35784d41dcdd483481cd96274db835845d661b051acede44f2c79278af91d36d490a3acb956b35b5104936615dad4d96695b3fee5e615c8d4f55d21dba4f499774b3c359652fb8512645cedcd8fab0fb8caa947915ad646ee3fec80febe771ab7d393aad8fcb2f37fb0f7575a056fbbb7fcf20d1c28407ec95a0cafbe6d84adaff00b8d77f75db3fa99a965e7eeb47f8547e7f1abe86d3fccf5bcf99eb79f3fa86075c445b4db157d5cd57f70d9ff72f2e42304c5392696c650e9aa9ce73dd51a3cdb681fe359b9b06bf235f3eab63f2cbcdd60ba76bf941bd56a41dc76aafba8a0febe6c95bf35a6c318920da556fcbe8e12b5266c63292b735b8525256d7fdc6bbfbaed9fd4cd4b2f3f75d624a1215bd7920cac357ca8f1b5c92d915f675a6af3e02ba59c54e748d65b2c22498d9acc29233ecff00b9796739ac6ccb25965ac07a71f6eaab3b985fe3cbec891870e2e6dd405be85fe3cbec80c92d837da048f5c5a5ec84797b752195c3edf5eb09f0b5d06ccd631bb796fee1111a9959b20bd275a50333f948d655e58c7b12d4cb1429d7b651ec5d943671ab72ca40e5ce892cd08c1d92be40db6daf032e6e5964d8930f04c0d9601d8963adb32c3640141959b20da375bd13159b48564dd4d0cf99e5b659de805a6c6a2227db5544cfc7c51d6c4542c428b1906511fec0acc4abb070c1025c9686be61f075b38b9d0e479a157c6286be6486244ea8a91ccaafa538640604b91828130ebec267ae78c78aff002db995c2b51c9ca3d8239c3f61f2e30f12438988d22e2351be33649bdb15f12633d56a487be08b01242968030cb0865490a158cf42a096488cf62b7d294224c8eec2b02146b81eff0081cc893a584c361d1cfb348ad2f96bfa205e4691ac6c309e38173911bb08701fcb1f83857e4c6533570512303ff577ffda0008010300010502ff00c7f94ce53394fbed5c8dceb5ce73ab3af3ab39ceb54c47a2e222bb19026131b4b39d9f209985a79e2c54545fb1d2b9d19d099d2dce13ec70d5c51317140b8a8a9f7090df0eac7cc1b316c1d9efcf8960fc6cd1bb11e8b9f9e45b2911312de660eee43561d9065623f2f230df1fe3ce739d499d499d499cfdae71511d8f6f42f9e31f97728984374e7fba6c6804dc4e94ceac5545c7082ecf4de2c19fab3ab9c19fd15ea4c190837b0bd4db992891fc1472e23917ed11396f9d967f458de18df57ab1a9d4bd59d59d59d59d59d59d584fc7108ec47a3db08df8847ea91f218061ccf393c36aae70b9d2ece87674bb17f2f3a85f72633fac8bf8e260a1cb3b5cd7317ec2409cac5f823fd229ba9ad83201ed8c6719de0f19c267e19ce22e739ce738ee151c1cfcbcdd891c386c13238622b9e3cd76b873640dc8896d560b48ebca2fc35baf0fa6d266cd5613c4c3b14a0864f711aabfdb6fd9eacfc7ed72b9cae23d711f9ce738ab9ce15394f35353ac93f9489d3d099abc86371a4c590c10e8aa596f2920c51b2fa803e86be5696b1afe32e25b415789f9d7b7a306de89d8aa899d7ce7ea5ce94f111d9ce739ce3d7f4f9960d5d856a2c998c470dcbf8e56d5c8ea6132c6192c41ac39050da4c96e1b22eb94e7189c3cb8a8913f083785f90d9fadcdff007bfdd7620da9e411739ce71ebf87980815f8722371cf6ab5cffd2efc169e224c9ef167e2dc74a18070bde9a7b099781b29a0a9b38f384e663c79b54346a601c9e98ff15f15cbf679ce717f1f31163a9dd38bed808221708d460ceeff0069cbd59a80fae63c58f164b67bbb26bf8c6131a4cb65f65218545ce11d9b5093e529f8ac7275247fc7c555e339c4fc71138f8739cf99799acc8b1d471ed59c2639396b9b815e30673057df4ecf7b333586b8caf167e2dc6972e17d4ab49b3133dfcec24b9654293d11047e9880de91f88e7f52a7e3889c7c1cb9ce279773dac45290aa38e8dc67498726234e3355ca12a84c99246ac7cd1a8c9f87c62d94d808fd8ae71760b75c6dd5a63edac0c3f871ca8d9eea509bd6ff001085eac622b95ad46a62af4a73ce35be5c921131a17915ad46a643b2930b136062e2de0171f728b934ab371cdcf4fdbaaa71f05c56e74623713e2ad73d46368da01f4b7c331bab04357e2223531551a8e7f52b19e59e46b315c43a8c2d678270f562a22e28fa3146ef8f19c7c1ac55c41aae318d62003cf8863756081d5889c7c1551a8f238ae18ba7cb12460c4e2635a8d4f0555111fd0677a7c2fa78a063b3d922e7b07e7b1c48e36e7a79d1828fe1aaa35085526081c7c5ce462395e770c483f2af7b588f2b88a28fc78642b479c1242b06d62398d7a285c99f9672c5cfd39fa33945c413dd8c1b59e1bded622abcee18507f1211a3446bcead6a313ca14e8ccfd65708283f0cb2304157e22227d956a2e7a63ce96a78a432331a37995ad4627c08746e300ae5f2a591ce306e22b18d1a782aa8d42995f8107945544c79d5f838fc7c5551a8a47994616b3cab9c8d4215498312915ad46a782f7a311ef711421e9f2842b599c10eac1b59f1219adc413c8a88889e3de327d7ca9b577702356597b6a1d79492644e98eb5b6e955b53191123cb8525ef950d626be130987951e2b63ce872f0b32181c5b1800411447614c203056300c8b3e0b5893e2111d6305ca2171f0b9852e6875c5972a7ed48718e8032403a20987224584288a39910aa8d71af7656cd872dd3a1c7463d846cb4748b7da19323910a9f23a44745890ac13f90be7c11bbde44f492ca3c87b23a27c2e6d5b571990eeeddb212f2a43acaa92b4d65023bc651998d9f05eff9a572148510594361ff0024e9f046f43854b226458a91e5c5949b3b26463c255743fb3b87f731ec3637c8d8e18a457eb35d1cd0a3d6c625ca451c3bf9286bbb891a9c5f48144a1a7d49e828d1890ace55832a23a5da0e65656eb716541d7bd48573d05d92de5ea6051bf5fe6a75a601474b540b497f15ff0087d9b6323a7da1e002442d493a660dd19971515501264788283b46e591b568ef0c5f5e82e6643142d8b6ff00eca4c314bd73598ec975512ae31eef63ac8ec1fa51bf8ceb152ce8c5544cd9dfeada0e3b198e6b5e9b0cd5ae835dabc6245121b5db9b0ad606eacb59871e052092d68e96b23cdb0da620a2cb895df238d5350b7596d56ea37ecb21255641fecbececb3c13654fda631223ed95faf6b7671e3c28d64015c1ac63beea42969ed24edb1103f3392ca8d5e403988c875528f2e8d49b19a3861535947754c2b38e2bdeb2eb76f2f6c020e13e59a35b34b5961af43f675b3ee9b067fc36f8bc8f580be6584c962831f5b9a18f60db9acb23c041b2fdd6919760da278259636d001862faf7f7567651cd79b25a46980f9c44fe3daacf0898930512fbd004a89d333a801646045ba49763e9a66c752f9e0afda0914737619964b614d31f4d5db4461c5129b62b99b671cb7b73651db53a94c120aaac63d6d96d53052252bc37757516eb4b96d68ebd7ec05008106d02ea7a9b24b48ff00616babdcbf2cadc74283e9b6143667cb6bf1b575c88a213862abae0bca109da383082f3c68f25b1e1448b85870cee6448a267cb6bb0a211982ae831df96b264ce9910aa78c48b18a4f8104333020047420c66632be00dceae80e6021c58adf9656e3abe03d4d0a1c85631836ad6d72a920d767b08af6020428ce7d7c023a597d9c385264fceb191630c9f03c0852941122c54c356c090f18c616fcb6bb1f1229582870c0e5adae557c082470c420b24438b2b011234542418467b22c618c400c76fda5fc7e0c4fc7c3e338f20bcae23513e1cf9d5f8b7f2f0b8ce338f20aa899caae71f75a7e7c6719c785ce7e3f0e5133af3f5ae23513eeefcf38ce338ce9fb5d499f8e72d4cea5ce1eb9d0dfb8b8ce338f1464e9cfc1738ce338ce338c72b5b9cb973a39ce338ce13ee5e338ce338ce3c4e5533d57e7acec59099ebbd73fdc7623513eea6fea6f19c6719c6719d3e0f52675672f5ce855c46353c246b9d9eda4f0a8ad5fb818fe954e1c9c6719c6719c6719c671f678ce9ce13c0fcf1b10eec6c58adc62d30f0763523c0dbc17e3488a8e688c969481f4fee04554c43ae7acccf5199d6ccea6675373a933ab39f13a7391a67ae44c5555fb30a79a1bd864735e66b07f7f44e591ede6f03fbd061299dec27781103ea964ce4035ce57bbc818e08ed11c321be6e82b80e89f2f8197d0db1667d8d63fbf27fa3ed35aaf553363b1555cbe4770fdb752fdafcd0c6e290026802c2b1efd8237af03e15d4922737f8c42e2ba996ba593fd10a3b65c9fe2b93233a1c910dc527f15cb186d8122cf694ad92bb94de62ee23738646159693be5b0e36d710802ee47eaaedb0720b842304c97b8058e66e52516aaee35aa94c200e5ee2d47337295d5596b16d0796b7d12b3177199cdb5fb6d60ea5fb5f9ad7627ab31ef68d94739ceb47351ed94058d22b62fbc9a888d4b2d83da9a9ee4d60727fa1ae731d14ed951f678fc3f5d8feb4e23da364833a41c94d5c695ed618dbb342ac6c7d44ce7d76cff00b3ebd4c1b37b2aeb86cd9ab83025d41dd22b36e9ee52d150c60462c08466c0a4056ccdae710d2aa696357064c48d3071bae96fe7ca4850e9e03eeec075f042cdaabe10226a5fb5f984455ce841a51c7f4604e8ef95143ad9805cd9a3749b59feff000aaaa5d63fbf27fa335995d42b78deea06bb1fd283b14af461aaa2258ec33e7c866ab6b232cf5c7d6c4d37fb6d9ff67d33fa79ba66bffb3ec3faaeff002f8dba3df79f29da73e53b4e7f1cbb79f69554a8d3113d1cdc3f6dd4bf6bf2ed6abd7d36c764613a74b4446a4dbe04391fca0195d623b11dbc6f750288e80b1cb1a095ee292a2441293fd19572bd9cdc631a365ecaf733e7a3960eac508ad736a9f1161e9bfdb6cffb3e99fd3cdd335ffd9f6d88e14fa4b2158c3c0d944932b688af8d67536a0b38f922d214635e4759555aa4f1c5959b64f8858fa97ed7e59115ca08be8b2593a8b4d2e2c13ff24afc315c72e5358b6bcffc92bf242896457ec43e875e55b519b30d64bb63af56fc226c515918bb243f4d7f1f85b6ac6f55951b23f3f8815b135eab93581ba8659f5faf5549ab666c5532ad32aa292157cd8409e03ead67188ea6d9a4e515112a9f36147b0048d52c63916af692256eab203232db562b8aca5d85e8fd40cd8b4300f5b0bcb54c7f508acc5fc57edf1e312d262286604d8fb08631fcc424c5b8ad696458c28af359418f85b4af0e23daac8f3ece506459418a474de998b28088cbe8e78c7b285170d65063a7cca0fb78d2a3cb6796a26a3a1b859615851bfec3405762891b9cb5315557c67c58fee8239b05fe8b9630d9606c9110eb5120040cf285d19243658d44a648a6031e8ff5a2185064b302439d5ed91f2ee4b06657c2380af8eac1d4acc707cb575812bca2b6ad3a3a4c1c32d61309f256e3a45637167ae3cc527feaefffda0008010202063f02ff00839b029cd1b5b479a294fa15ac7d9317ad46711796404c6dbe8e2dafd338a1715a93d728f8788fc28fdf8bbbdddafef8bbcfd9e78bc9aa4fd2a08a4a2d8b4fd0a1316ce3de0944d367d421fc5d106c4e53d422e81246611240998f7842472c6d952b8e25779cf5c6cde1c71ee973d71271328d98dea532c48c8297b57572449495cf5c7b95b8857111e6e98bce0bccfda1671e6e01ddaa24e15c0643eca80bd4d7233f0167809a2262df1ff9cc409fd84f9ce8cd9e2b5518dab62e204d7e56c50dd1a236c93c7c1b2488ed4c69afa62e3e2e9e689a7837ad8fe6c7e71fbdd3c0a65d136d42460a331838d3d8687e6553a27cde0f47d2d07c7b6fe022aaeae3e0d31711da3e0372bb32448f07ce363615dad79f8fa75c29e39073e48ddb7da3e5330186ac1ce73f84bc22d116c5b142206bf1e4b67e29aab59ea89f0eeb1789c3b4ee65b8849e426038ca829b394198fa3b85e2f0c1ff00b3bd44f927380467e09c2985f6542507047d43e462e22dca73f82a7057e94d2624ef2c4c59e3684abb00ccea158d262e8e0461b026ee35f9ed654a05b2d2674392b965056b9959cb0970151c0295ef11908cf2fb4321e2b270169aa4f0ff0065c228a1a09f792b55784c2754adcf3ad9c0df72e29455ddefd113ff002d792ee85592ce41cf3e02984e247f989aeb4d3abc2591645914fa5bb361f1b71dcb765ca7d103870f8c03dd00a41d06d1cb5e4e04e1d84de796a000ce4c31dd9ddf23de8b686d7d84a766f4b2927b392849cc77d89c43ca7b395abae1beebef459730ee9ba952aaa4a8f66b9524d2b66a8714e0f74f04a9278824f21079b83081b1f0de4b87406cdef34b8fe8694afa470d7c34c5bf4078d4bd683f6c9848397ac70b9ddc86c62722a7d91c952468948e59c50523e6dcc3a5e5e433914e7bb689f14f24c4358e15c3bd854141d15a73cf8f818c33332eade401aca8439dc430e971c60fc52aecaa427740cd62a66448ecd0706e9dc3f6ced3a93b5c8720cc08e3309c461d416cac4c11c124f952169568e98cfe297bc67768f89d11bd5da6ceb8b873436b191c1c0e3ed993cb9213ad597884c8e0d3030f874de7546908ee4ef1714fe3d926e96c7c2fbb3368ce2db28990e038aef1de0c74886d776f36dce97a53bc4f159cb1f36e3831185c412b4be9b1c9d4cf32b2915d04c48dbc0f7742ccd377789d15015cb3079781672487440f15978c5d4fc6559d7125f606d1893702f65e00630cdfdac4cf910afdee15f79fff00a9f56edb3993eb11cfc6044c70e27fff002b8aab6e20b8c4fd47535a6bb4fed7da31a636ac86db1eb34e0fcb3f37029cce6128d1e293f18d10e383b1390d43ca70e8f5a43cfe8e04af4c058ca2250138b69b7522cbe90a972c530b86ff006d1d514c331feda7aa30785640435ef0c8090f57d3c1a7830cb1f6c8e5491e789ab0d8727f868ea8ae130bfed23aa37d86c3b0dba32a5b4a4f281c2a22cfaaa917976c493410bc2bd4710a9180fb568e719a3de2b76e663d7672ca26871047b423773991f450bef16b78a6e72da5a653b7b2a19a2987ff00c8efefc518fcee7ef44b73f9d7fbd09c432cc9e49983797e757d090852b298bc6d3e1266df1d9ae2ea22678378f028c4fdb4dbc790f1f2c7bac48234a25fe23157d1c863de3fc89ffaa24dad4a6c9acfcd0169b0c6d763c15c4db963445c4fc34f8499b634c578243829e2d4b2266d890ecf81dcbdf04f34674c4c6db1ce226d19fd2bcb324c499a35f6b3ea890b23e5b0e76b29f2cbe126ab6249b786422b121678b4d71216c4cf82908b8f558e71d71bc6ccdbcf17c6cbb9c456e3a9e431ef5a753aabd51fe64fd9f4c7bb6dd5714a361096c69a98bef92e2f4d9c913f547246e70a69f6babafc1c844cdb12470c844cdb1a3c5691a6248f07a234c562fb2a9189621250ace9b3f0f547f2cea17c774f22a5156d7c93e8895d5cf518d96dce431fcc290deb227c8271eeef3abfc29ebe88f7876330a0e4f074e0fbbc34b2243b5133e2b3364688d1e0e6b8ba9b62bf47dda943518f8cefe2575c6db8b3ac9f0b33d98ba9b6267866aecc5d6fc5a6a8ac57c14844cdb124f8a4845e5c491c32117976c68f15908d31a6267c1538242cf14d1121da8af0ccd13175b899f10a32805b903309ad2d84e19a624b54ed4225413ce73406194499d8bd74500e2b213ddf85c39baa29da032e44d9b4a3ae70313b82ee38e496d5788912cb0b65ec2ba9c41bc6f96e82667da358b8d8257a22f38d2c0d5094b8d39bcb6edd3394f36686f74c2d8a1ed22e4ece58bac214b3a047f30da923488becb6e2919c249e8821b6964a6da59a35c5c75252bcc6917194952f4560071a7013650c14069cbe9b45d3313b27481bd6d699e74911465dfc2aea890b3809c53616854864a69ac490ca66e52f0029979eb0bc338d059ed4c806423fb76099fe612bb529b764cc52bff286d2de1dc6540daa45c9d39e2f30da949d5179c6d61339764db070cee15d18903b65ba76a7dab6ca46e54d2778d4a648159deea852986d6a44cd80c5d58215a61c69fc2ba5d295497bba09a686f1cd6c1c238d24ba99aa64039617bb6d4b6d18824a5299ec85e6f21084e19871a377b251749a9b00b61b0db3fcd49139276b4e49c5e432e94e849ea8dc6ed7befb374cf9237b886d61033a4ca248b382e1a329ed1f36b31f2ccb779d16c803caa300a5091894d65200f51105b40011745045f69a59467945c7414af31a45f532e8467ba65d11bddcb973518ddb492a5e6026619f9667decc4eea6bd9cb211bc434e1467ba65d11be2956e6729ca93d71fcba14bd4225884291ac47caada4975159900ce661d02cde2ba7e93bed8e885a99c428ba1264262a654f5612ca2575d20186d84ddb8829707b427cd08c6a2eef9413aab0b7dd95f20d9ae3e7889e216073d8354aa63f9a4a54c1b642bc55e9e586f1583046c046d0fbda0c61db4daa98e7108c2f743578e5553aed3cd070ddecc82c285b4e700c3ddd225f2e89a867add97318531864a0348565169cb0de3a527364fe2b4425684838a725cb6d74089635214cfdd151cf28562191fcab972f5eb649a643a619bb56e731ed4372ba716ae4a76b9fe84ad7929fcc8eb1d30ef78b9967c89eb338f9d4cb7f326ba673e9864fdef3437fd98806433599744e76f1c0c263582876636b21238a5c861589765bc50c9a1728c47ec7f8e2e609284e193652d838d099629b079aa46a22c8796eca61b5a6998221cfe1f9c4381b97bcc4149d45c865f6fb69479cc231c8bbbe504eaac7c9bd2dd9bc539e64deeb8dd802fe797a929f2ca9ae15ddcdcb77217b3ce77bab8290e293f16f9fd224382f24c9517f13b4db6275acf371416b08121a419572cbcd0a7ca40c5373e515e450f2a42b198900dc0aa7b30965f09dd2d52a648189c3017d4dcf44ccd27a21b79abb7d76cf4a614c392b8dc80e39c37dd8ca42309bc12cf2131aac9d21380c0252084f27a72930bc0e39292abb3d7d4446230e7d4a73c3bfc4574fd225d97bc9284b34a12ac6042b0d59899cc659b2c0f94012c341248d7386df4cae2ca5b1ed19f3423028bbbd484eaa41c09bbbe20eaa983dd8e282716cd3f0d01d23244b12a486341a9d54e98185eed5150455454464359484378f6e5719b78d49942317dd0ec959533f2a8e7853dde1892cb691a0fa21cef25aafb44144cf68d532a6a1174ca6fb9b3c6655846093777c8099e6a425085018b6e5cb65742a278d504b23ec9af448416f064a99194e5e40290d6f6ae32be74d9cd0a97c346c8e2b79e1cc605dd089d259803c2e610e51787150f9b92118166815fa53e994270cd4b78acfa04e12a44a4d54f26480ef763ab52b2a7cf29c8c1fee41231247e69ecf1e79698feddb3f3123abb77ba21c75c95d7252fd92a063798152158651a56c83830a9e25c07f3509d40590e61552de38954bf6934871e76edd914d33857a22feccb7bbdb7d5de74c2716996e92997398698664577536c5eff31b5f3831fde3ff0087f2f6b9614f2fb6a338463d4b9a577692fb4279e249a082cbe64c399731f4c7cc609612155964e2220e2bbc1495c865ec8e5b7474428a5296b06b17448016584ebaeae282ee10a4b4b33ae49f9a15872a0714e4f94d39123cab0bc12aeef96152cd58ba25361cdae232a4271b4dcdd09d339a8c32db176f091afb30a7972baec88e784ef6456d904cb31f44271f80524929e5f4e422178ec7a9215765abac9877bd172dd3f3967ed18f9670805e5cc65a2946030557a689e6ca479be8dd4bce848fbeaeb8f8ef7e3575c15875cbe6d378cfa605f75c32331351a1cfae3e33bf8d5d71f1defc6aeb8de851de679d796376e3ab28d7e538beca9495e7065d11bb75d7148390a89117985a90ad0651efdc528693171971c4a330511d101c71c5a9c4d849331aa3e3bdf8d5d717da514af38a45c79d5a919a7c0a4a1e4ef5c139d294a5279216d24de095113cf05a6d6b4b66d009039386fb4a295e7143179f5a96ad249e98de344a56328a18b8b79d28392f2bae2f875cbc3ef18bcfad4a23398f8ef7e3575c5d43ce84fb4aeb8932e2d20e6262fac92ad3120f3b2f6d5d71b0f3a07b6aeb8de6f9d53b294ef2accd38baf38b523312651710f3a10325e575c25a7552beaaabcf1f24d3c90942746d4ab9f29e00cad6b2d0b01265c9c3261c5a53ae9c9137d6a5eb3c171a756119a717dd254bce6b1f1defc6aeb82e36e2d2e2ad20999d7171e71c52331513d3120f3b2f6d5d7171b79d0819028f5c5f75454bce4ccc7f2eb5235189e216a5eb31bb69d7128190288101e5ad65d1619998e38bcf294b56933e9fff00a7ac53eab3abe84bc2d7eaf9448dbf42bf5e5e4f6e24aa1fa548a98d98afd6f2559164658cb1526283eab97d1a457c5e6e100698b9bf66ff00b69eb8bc820a7ea1fbd1757e2f33645d49bebcc817ba291fca6117fb66ef37a6364b4d0d12ff00aa3de6227a378ae894a0ad4d158fba42b9ade68baa12508de6156b6dcce92474423bbfbdcdf42cc92e650725ece34da32cf27d41255446c19452516459167886cd4c5a103454f3d39a26e6d9fbd5e6b39a249121f44de00626545e5e3ce3c8416d636d264610d37f114b006b27eb5af8279c4f64b8ae98fee2e8f768ece939f8ba757d6bbcc5b8db4d9329ad412279a672c486330b3fe2a3f7bc0109f8caa0eb8bcba61c5ba74080db624803c46eb2952d5a04fa22e3c95217988974f8e27038079c692ca36ae28a66a556b294e425ca63facc57fbabfde82de216578b65722499920d5249e51fb3f45aff00da4fe87211ed0e9fa778c6f9ff0087e5411711448f125ff04fea4c7fa63cfe36bc4bb469b4951d404e1cc63bf11c5951e330dbee0934f24949cf25149e7101851f7388173f6ad4f3ecfed709c3206fb1c2d483209f6955ae8009cf28d9670d7352e7cb7fcd08c1b8ceef129782a60cd2404ac6b076acaeb847b43a61cc7b4d6f8b626533bb4ca6725596d91fd17fe6ff00ea86bbc1b174382cb64419113d70bc4bbf09b4151d404e3fa2ff00cdff00d51f3eb677292a2122f5e98197b29cb31c509c63aedd6d53a4a76196719a3e239cdd517b0ae5e398d39e0b6e092c40c35ebb39d653b071421ac3ab797a7332ba13ce63df3aa2bd000e99c17708abe06436f167e00db60959c917b16bb8730af3d9d31b0eac1d201ea80b70a54d13298ea80d340a9c3922f62dc92b327afd11eedd5856991ea8baf5506c50b0fa783783630ff68f98658ab8e4f8a0bc1779a2d9198da0f9a3fd31e7f1bf926cfbcc42a5fb29a9f30e384b0d55c5a801acd0431b81fd15d1fb264957e6ba4c25d6cc9c49983a44358e6ec7100ea3947119887b1a9f8c0493ed2a8392de282e2c92b2664e7309c7f78b8a6d95d5294caf119c9330279286958463708e38a0a74224a9654a8ce600fb39a11ed0e98532e89b6a0411a0db0ee05ced36b2358c878c48c3fdd8ab527789d468ae432e58f9649f7b885ddfd91b4af30e38461da13756a091acd0435826bb0da027d3c66b09c2b2bb8d26765b533b6d8edb855acc1c3e2778a62edaa9d0eb309707ae8aeb14e8942352ba2128600de2e753925a22fa9e72f7b44744283e66ea0dba3243ad27b37fa6b0ac72fb64c93ab2f2d9c50ac3e194518749953d6e3cd179b7160eb309631006f12a9decf4cd071ee76d566848b794f4419288c3644f5e731bc61452a8de91b45bbda949f4d354230e3d657365e6809c3817fb2919b4f174c5f71d5956b30a61e714a68364d6b95396dcb1fe98f3f8ccd5645c476216d367dcb0376358ed7e6a714378f7dbdea5b33bb395654c86c3587308ee115bb7105276c6512fb3c0ef762fb4d2af27d955bc8afd50dffed27f439c0da13d9084f4435ffb49fd0e423da1d3c0d77aa06cb82e2bda4f67953fa61878fc252ae2b52e9cc64ae28f954fc3c3a25fb4ada57f847141c62c7bac3267fb6aa27fc47884485b1f318f929d94cdeeca7cc78f8a2e61d0a2340091e5c51f2c96ca684ce79b8a1af60f4c2352ba219d4af3703e3d8ff143dac74086ae5b257ea313e16cac6c6eebcf38f828ff006847c147fb42372d4d2891a04c856133c895744323d5ba7cdc0bfe09fd498ff4c79fc62fae898b88ecf954c3d8cffb68275aac4f299415aeab26708c78750da1739020ce40ca71fd4b5f8550865e505a5c4cc11aea38a9cb0cbea326546e2bd955398c95c50eeeeae34439f86dfca4f03786ef52a6f10da426f5d2a4aa561d9999e7a7a1180c085909742ef9a0a2542405beb659423da1d3c0f615226f5dbc8f69351cbd9e3e053ef1bcead44939c9842943df3fef0f1f67f2c8f19868afb1bc4cf9441dd645027557cf23c1f36a490c5d3539679a1af60f4c2352ba219d4af370623f63fc70f6b1d020e1bd76d5cc6bd338553dc28cd27cdac700c53c9ba851909dbc91f2e7b6dd388d47571414a87b82764e8ebe053e94c98424aa66828274cf0d38aecce5cb484bed09a9b2790f54b8158b71252c96e42796a0f2523fd31e7f172a55122367e08b3ae0387b6baf1648460bbbee045fbcbbc656582c39e7c422dc3fe23fbb0de11af86da02471097021186ba316dae62f5923da1d078a2dc3fe23fbb0d378fba7141b0172a826c9f1c1c4f725d2cabfcb26453ec934235912d3173e5ee8ce54897eae881b9710bef42e09cc9080892a605264ceed4cbad2a2589023d63fbbc2f3d802cfca2d65499a882275976721a0d108f9a533f2d785e928ceee596cdb28ba9b0700671f30b1eb5b3d79671bcf7657a115e88ec918391d2a39b5421787bd24a72c2710f4f7601b3543670f7b641b74cb81df98bdb77652d17bae1cc4353b8a3975407d83258e7d062e63125272d2f27cb8a2fb1702feeb723fa44065a410da55399b793d301f60c95cc7418bb8c4941cb4bc9ebe68be9dd5ed0d19fe985e1f0c851be92999a5a25657cdc018c7ce63d7b7972c6f6682e68457a236d24612eeb51331c949c6fd89dcb80578fc5d3854f69ca9d43acf4704859e2cc25c7561e7c4c49131532fb50b36b6db9709d3e420368412b28bdfb272c383109505a1bbd933e5ad9aa37a1a3bbbb3e2b62fb282513971e613b60eedb3b26472573572e88371b3b2aba72573572c6ecf6e7283857dd706205a6e6c4f96f7346f194128f2b33f146f44f7e5fddcb8ba670b0055bed68acba610db932d2aed44a73227294f2412c2094032e3cdae141a6c929323a0e9e48386dd9df013968cf9a51bb7d375529f178ba2f760b225f895c09c3629413884894cd8ae3cfd3f4769699c7ba4288ce76475f346d9e4ebff009453c361c61714d3494b7b40ab2cce491c9188630cb6db05e4a85ed90401232e3aca376cb987537f2cda0df324aa433d24469943e860a45ec2ca415317ef54249b7fe718570ac6e92ca0133a0d8a8e58650d9c3875a9cc3b4cb3bc9af2e585a714ac2b8c1789335141d2a4683c705952985e003ca95f514b804ed19e6214107dc5f323a276f244b1d89c3bfdde0685288d0257e7d1186521e435b9b428c88da9de19e7a32c1741a7cf5f965bb9e518b737cd2b7a36403335583c5aad8c363f7adeea48044f684932331c50d34875b6d6d2d73bc65da330a19fa63125a576de44b48008279611716c11f2884a83868744f2114848c2cbb1b401bc90aad124e4f17dda8dcc4a3b0bcd9c1ce0c4b725d4676f6b9bb5cd15c2e27fdb5f545d61bc5a07b2b03aa3d7034dc1d31eff001012345bcc074c4f14ebae9d269e7e98f74803cb3ffc2effda0008010302063f02ff008395f01a6249a98d96d5d1d315091c7d538ed37ca7f762772f0fbb5e6b6246df0b51c1b315fa86ea2d8d3c19cc6c8e0a88ad229c1741f7662d4f246da5247244934733707cd0f8a99718b3c424625e3fba6f8cf9b86b444678a70562c89b66912cbc1257c2e8f4700711450806371eb2cf30f4fd5f751f155675f1705ef57245e5780bc2d89fadd31a23e59768eceaf440445e55905c5dbf57ab13ea7653a865e330181aceae1bccb4e2d3a124f445d58215f47781974a33dd54ba3847fdb5f4c6f9bf888a8f38e3103149f5c7908bcaf12ac6c78e2ee7c456c8d6aa404fa884f441c42fb6eaa7c59380bb881361bc99cf544859052401880365598f567891b787e7de13593b3a25975cf938158f684b128a9fbc32cf48b67c0a40edda358843b948aeb87309ff006d74f65551e7f15bd97c6d96f25ebdf847a61405a69cb0102c4897038c7af43c05c70c9091330e62f13fd205599c9acb8b2c5c6db4046a10ac660d375d409902c232d33c2427b48241e59f4703c55eb20a46b5538270e35902e7cb13c8b6fa0fa782b1b358afd57a201c812600fbdc29c51516b367e0dca1d284e6c875e585e1cfc543ca07cbcace071c73b01067c909ef02e14a5c1d895a34cf9a5cbc17d0ef66c49b397ae705a74496380ab3810850c93e88aec88d3e272f19bc7b31705917931583096d5f0c6d1e2f4f0975d324083de184406f0cbb42bd7d32161f2acf837385ba70feb267252b44ec946e509dd3cd50b66d4fa38518c16cee9e91e7e009cbf574cfc31134f68d044dcb22498262f64221d5666ba543ab85383ff25a17d5a4e41e5a625c2d77cb34525575cfbc839f57566e152b3293d32e02ac9389fd59a61293da9575c20e4af04a2e98561d5db6cf31b226cad48273123a23e33bf895d71f15cfc461f79c24af62a78fd1f41d1a3a0831475dfc47ae3e33bf8d5d7171d75c52331513057eb58359808cb2f0da229e393545d6ec89aaa603a8eca8416d7646c8bc9d1d5154ab938138d47ab450ce9f444c764f0a86117742ada24f4831f1bf2a3f763e2fe547eec7c5fca9ea82d3ae4db36d13d5c32101cfff003b56695678978590b22914fa133e2f245b179c890e0ba8ab598f95236da3cbe88f86a8d96f9fd1179400545d3170ff004e6cfbbe8f05b96edf58e6d10108b044cdbe12ea7b31f762438267826af16adb121644cd55e06f26d891b2256a3a22967d2d1124d06789262faacf09753d98bcaecc4859c13364485913576bc5a4dc5e559124f8299b2363b51236f056364c5b15319cf0de5f8399b2243b31797c333121647def159989648bcbf07f7a2bd98a44951b26634c6d0222d11922d11b33315a08a5be0e6783ef70d6d899ecc493e2b24f6a33aa3ef7839371797644859f46b1608a01e1643b5179564493c37515545f77c5aea2c890890f053364487662f2fc5266c8badc4d76f0ccd9175becc7def15998fbb1a2243c14cc562f2adf14d3133d98a70c8555179d890b3c42afaca5c25424a5504ecb6158a771136d32b1c5cea40cc33c1c43eb9bfb776f1a922c15a98577962f120a9215b24d6549aadd948d528385f980d601396f6cd32da0199b2b6436fb18c655861745c0ed4c932ec8a562fba4040cf645c6dd44f588529b79a0dd537af895e967cfcf0eefb10de22a3b2b2b95bc917f10b4a13a4c7f2ce216741af245c7dd6d0bcc54074c02ebcd80ab2a2a33ead31bc654148ce2b1bc79494a33932825b79b2122b514d301c53cd06d561bc99195b2ae483b971b5cb3281e5944cbed7e34f5c5e55bc00611c2dad333944e965226b79726ea52493316731942314d3a5091b37412264d72688fee98e7ff009653762d469b4246a65939e1d539896df49140970ae55e68bb8875095669d62eb6eb6a5113a28599e06299c5b2ac293f0c3b33d8976452dac6fd2f2f74f4e4012257427ae128c43a84ac816a845f4105073434f61f16c8652a44d1bda992aa2e8a126c818c6de525a54937412325b08de3a96dd730c0052952da28cfcf9e1c562b10dba9bfda0e5e0282d26c870bafff002735ca6bd8d12acb545c71e692ad2b4f5c7cc6f5bdc7dabc25cb1b9c338852ce6509c4d753c17c55f576479ce811f34fbb759365e2403eca4410a5a8e15749cc91d69301c7095397d55358ddbcf3695e69f94a378d282919c56376879a2e1c97d33e98dcef9bde7b42378f29296f393210f7cd3deea46ede5d3b59266376b79a0e0c85499f4c6e0293be027767596a89e25c4a35989e19c4af51818b6de506974ba0912908689b7769e8fa4cfb07a61087f0c80c958bc64682753dbcd0a7dc9df681239ad87310bbd7dc0b68fb265cf0e60177b7092a956b486f0ed4ee050b7547f6f0abb876c9e6b4eb9d047f26b5a7102c24d0eba538b921cc1e34826f9705c3f7401688c4baaeca647902a178cefa7ae8c89af25940394c0c5773be462126cdae62474e4863be4cfe657241cd4bf3e7108c46294b2f2d390d8327343bddf39b7b43f09a185a16a29c1b73e49ca9a559e2780594bdf78d0e7b04c79521386c419e2dabf76e9a4d55acc6810fdf9876523ecf94e1c9de183472d4ecf37d09d8c2d5f957d47a21aeec6bd590e35750973c7c82e630f2029f7652e887c7ddf3c387bf12a226acf6ce873ca567142b1981c405b3222e6500e99cf944270accf7693974b73f3c61ff00d4ff00045fc7a96ac52aa6b6796981802abd84748fcd4074106873c30db33baa71b557395c37fc5ff0986cbb3f75860b12ce1b87f0eecf76a725cc21ceef5dedc24ab5d23e7582a2e8ba159a405d1d0237849dde69faf39724ebaa13de8e4f7b78ddcd295debe09986d0af837073a8ccf9668cf28baa00a603785d871c32a5243291a612f634acbcb13a194a7e784e1c28ab06ecb90d27ad279b5c270385245f29a9c97b9216fe1caf7ada675cb9e0e171655bb4b9774c85d50e987587af5c6e644b42a508c4353de3b327589590e77abea2bc66e8cf34c90759339579a17de3de2b51055cbd40580084778777ad4137a5a8f9d27cad8c2e247af5e6867f849fd23e924333f773499e79c293812e271549190ce2794e49c1f9c2a53ef1524190c92b6c871855ede2029c3ec8973c398f5dedca8ab5d6138f17b7008d741281deada4ab06f57f1d48d0728ff9c4f0a952b1192f0901aebd1cb0ac5f7aa420b9b29090728a4e67c8439ddee4efbf671255385e0fbe99bc99d152f2a1e684b1ddb840fbaa3a53e9e686bbad09b8ea485c87644c2a723acc5e17a5876f6b8933a42f1cabdb95954b3ed590b5ad25583727c939d34a7375c4b0092a7cfda14e99980ee34252fab209d064b49ac3bb9a36fa0f22ade784cfe239b478ece69437812d95172ed67f694539b470b78d4faa6e9e3a8f3f2c39de0fd4a05bf795e89c2b14f4f769959a4cbcf0a4ae737a413cb960b3deaca129c8ab78a6002207f6b2a38507f2cb6b8b34f4658fee7b5f2d31afb177a61a69a9de6e739fde0822375de097138a4091a5b031c537708d11c89a81ac9b61bc5a6f6e9a52279f615330db0cdebf30bae629f4c5cdb9eeb7367afbbd7642b06a9ef94abda2c10f621fbc1179765b5e48bbfe5badf1c9423fb27ff3fe6ecf2425847610903921ceef4372522f567f64cac944d753097b0e27886f2671e8c9c71f2d8f42945149fadc60c0c2776a14899c876cf276467e994242d4a771cd9bc664aadb52279a9af8e12ce342c3c812a09ce5e784e2424a706d4b90565ad479b54231c9bdb9414cf3ecdb178de9621bd9e34ceb0ac0577e5457a2524887ddc45ebaabc29ed42596e77da983cd642b733087010279c7a617dddde2850015c9d60da0884777f77a14537a75ca7cc91a7cd0d774b73dee1e53cd5488f9a6c12184488b2a94883884a6e49776d9e407cff46f29864a8fdc4f547f4ec7e04f5406cb2d6ec582ea65d10776d369bc246490263359647c067f027aa2ac333f613d51ba294eea564a9c91bc6d96c2f574668b8fa52b466227d31bc699692e0ca12907a22ee21095a7489c7f2eda10740af2c5f7da6d6bce520f4c16db6db4b6ab404800eb1963fa767f027aa2e3c90a46622717d969095e7970254e30add36652ad6b5acb2c21e526ea9490659a03aeb6853a2c252091c7c3bb79214de622622eb084a13f7401d11bb75214d9c844c405a18642c65084f546ecb2d5c3f7472c5dc3b69403985bae3fa763f027aa2f2d864ab4a13d5137da6d4467483171b012819044cb0ccfd84f546db2c93ec27aa375b9692cce72b89b73ca56c5f65a6d2e67091382b5b0c959ca509ea853ad227713448e88f9f7585152d7a766f5276641c05e6db4078daa091333b6bc33c434852b3cabcb6c4b0eda11a8706f1e65b52f3cba73c6eda484a330a47f4ecfe04f5406dc6db5369b0148206a1922fb0d3685e70903a22658667ec27aa0adc65a52ce528493d11bb65294b7980908fe65b42f58896190940d02378eb2d29c394a524f441650da032ab5212247588b8c21284664803a3c0cfc7a9c14faf6b148afd7758d9ac5691a7eae9f82d315a70522bf5948d914fa558d911b47eb8a7064e0d911b4629f554fc4abe0f644e27bb5cb51891b7ea1d1131e2f48991753a691ef9e1fb35f2e48a85afcb8a365b97ec88ba17775d3d1131645d752149d227071382d950a94e4e2fa86915f14ad22c9c491b23452266dfa34ab5944050b0c296aec807ebf424da1223e551da36eaf4fd6b7594a94ad027d11f05dfc0aeaf0133d81124fc58bcaed1f11bf885a5089daa2074c5f61695a33a483d1e387118842545669300d06bd338f82d7e04f545e6c00cac52564c50f5f1fd157f08fea4c1d5f4e422e37da899b7c491fc71fa5707f8aae81e369691db51972c25947652250a424ed20d792705c1db6ebc597af8b877aad8633e7d422ab76f717eec1792bbcd144b4da3aa0ea84e1d4ab97b2ca7e711f1ff27fd50ac32ab7612d23b4a32e58f8ff0093fea8f970bbe655a4bce615836dabce26559c8540398e78a34dcb8fae2ee31aba9ce933e6f2d501c6ccdb2286158bbb7e445272b4cad9185bd894eeca489006f156aa08f70ca023ef124f34a0338d406c93da1671e6d7c05c7080d8b498b9836ef8fb4683885bd11ef19414e82475c14361497922641eb82f3c425b169317704d4d39d5d5e98f78cb653a263ae2fb1458b526d1e8d3c1bb3b789fb23ce7274c51a6eef1f5c25853771e0e055b31292879e0ff00155d03c6f7eaecb639cd9e782b576409c2f79fe7cf96d1cd31050aec910bc3aad4aa10c1ec65d422ea7b220e1f0c90a5a6d26c9e682c3c9489227313ce069cf075405a7b40c23109b149ff009c378a19764f48f3f246f4f65b13e3b075f141717d8489c29f55aa3385631f6efbca95b6504acb3247c36827d91031385dda7117aa132a8d43367e58534af51ca6a201e99c39ad3fa842dcc413ba6e541967a6376961abbec83d309561c5d69c4ce5988b65cd0cbabed147453cd09eee41d802f2b49c838ade384e271490bc4ac4eb62676533e78b8eb4d94fb221588c313ba522574e4a8b0e6d7cb09eee6fb0894c6751b39074c09a42b152da569d19874c6ef1280a4f9599a3740ec872e9d2957a083ae1cc51f513cf939e14ac493706d2ce7d1c7d11bb6da6c27d9109c4b0da50f1740a5297546cb3241fe2aba078cc844d56c05abb6e6d757374c2b0e855d2acb09792f0bc920f67371f023142c5091d63d1d10afe11fd49e05136de30afe11fd4983ab8178455a9a8d46de7e985a076c0bc358f2946f4f69c33e2141e78dc0edb87985be6899b23e5bbb6f25a9c85ded2b4e71c5c71bcc4ad295692547ab9e3e696e8554095dcfa67e687bdb1d10e6b4fea10feb4f9f830c7dbff000431a8fea30edfb269e4ba9890e17420fbc2ed39a51f1d7fef2bae3e3aff00de575c07de214e4c54ae6690a965527a61f3eb5e4f41e047f1c7e95c1fe2aba078c5d4db178db0867ed2b9b2f345d1641c394294531f097ca214b4029293642d03b63686b1e5284deecaf6796ce797029dc200a6d46729c88e58388c44a651297183e683ab810e9ec4e4751f29f006d1440128501d86f6479f9e1e08edee972d774c0df4a650427daa79a6383e492b0713785064967cd0f7b63a21cd69fd421fd69f3f061bfd4ff0431a8fea30315fe5ba9e74d3a2509aff003091250cbaf51e05611855f7129999599a53cf1f329ec3a011ac50f5f1c05248f9803693941eacdc08c3ad60e216b0909153532ae6e3879a4f6eecc7ecd7cd0bc3bc64874097b43ae7d1c09c1b4b0a7c3978cb25142dcf5b20ff00155d03c5e42d8af6cc5d1d94c29fc45e9dd9090e58b1ce4f4c29e5769449e5e0529d9ee549c9cd1fe6727a614ac3cf757a901bc7cc2c7ad6cf5c4f793d40f541be923097699544d2de7823de599bd3c284622fef8264699b8e15ba0bdeca94cbcb1336f017fbb645b57a964b564946efde0469729faa0eda558d24684819749871189bb352a7485619896f0916e830e8c4dddb225239a7c0cfcb5dd8bd399cf7746886f0cf4b789193598387c409a0f28d222fe054163248dd57573c6ef117ca3ef3a08fd4614fbcb0a75499485832dbe88387c409a39c1ce22fe094163256eabab9e376bdf5dd2e897eb86f158a7122e2c2a42b619db4f3f01c4776cae9f52cfc3925a2928dcc961bd2b12fd502e2d2ac65eae4484c8f299ca3e5f112bf7c9a7178b978d89e9e0adbe2d88534d3658c3a8824ae46827f661b163ae377c272ca9d705d5ac0405947ed0b469e28695865214daddb959e69d296eb946e54ea7797aee5b672ae6ac6edf580b94f29a6732b38e06f5c1b4998cb4cf49d34c2778e0da4de169d939699237a2a894f8a06330ecb470cab05fdb9725de29c6e9f7025ce8d79b8e37265f2e30fbc9f1cb9250851551cece9a4fa21c71a290f22f50de948194c9bb96dcf0138870256533ca699e9934984979c480b134e91a256db0317bc4ee09903a734ad9e88de61d57913971f8baa56ef3cc380bac89b679bd1f468931b6a1d31b2397c3e24e3308ebca53bb24272486598cb185c46290e3a430b42aeed904aa699f15271bc7dac4a1df9a7169b826b44ce515983a271875e212a377173994dd55cbb20a581657cd18c69283be5e21640954fbc983c90fb8e8c5165ebb22d56724caea848f164842f089c637880c002490b06d925792638a03e84e21bef22ca41dda02db5102c36ca469d1092e247cc6ec4c0fb52b39627ddd85c4e1fbc89c934a01d267725c558c521c65c777f5494a660ecddbaafb3239f24065436bfb76ee792f669c609adc3c8dc9db2a4c851b29a67d76462bbbb72eefaf3841bbb2415cc48f1d90ebae32e3adbcdb72ba9bdd9122939b3d691840ea7b0c393fba4a92409eaa42f788c403f38b5a54d89a86996506b92147173edec95242545349150197c5ef0ab46d1e5963b77559954f473c7c66bf127ae26e2993c698f578af47bb68931ee508446d93ff0bbffda0008010101063f02ff00d1fa393628af36ddbcdf473571fc798fa5fa98eace8bde57c057c44a8b8cc24849ce2a8a9e34f79ebfd07159f2501d77ea22368af4c92ba688c466eae9d55295a654e55c2eecd35658cbd9724a0cbb890f22ec057758f5e625354c669afcbb89f3cb92e2b68be8b0d2b6c88f4531ea62c56fe430d8af8d06b5c65ead175a5129e2c51d623b9f8c65b3fd70ae33b08ec177f0b05f76317d102d97dee2adbcc5e584fdc256589372f33729b4d8385f2c531ba219c2b8a254add38761257d266aaadca0d1adb52d1cd8a9120a7392a2278d71d796cfcc5daff9a43c68279cf90c97ff0033263ea277ef6c7f49c20ef3b125e49024d27ef9a5a4fa5845154545d28a9a51539d17ec3ace027ce4af8b5e342917c91fdb65c755b5f0953e245c68404f02afdbc76fc4829f6b1a5c3fa4b8d2aabddd0649e15c6bcddf4ffd98eb253a534a7ddc5516a9d1fa02fd8f85159913182566e57b7136b6eb59f94c474ecdc2e209e4a2ecdb5ed574a23924dd725cf7f4c9b84b3dacb7cba5c2fab6f984682898537a48220eb553441af366e52e84aae1460c67652f21916eed78c85c757e88e1763b9461e4cac9ba7e1275d215fa29813198ee5a56836d8ca068ba96bbaa978971eb5c89228ba51e8d917bdea0d8a2e106e104dbe7723b9b51efeccf66429f38b156650f4e95a8fcb0244703c294c20c841751150da704a8e3469a45c61e05436cd1795170cc5e2c372e167cc8d31c4235dfade8ba01bbd0362bbcc5ff00ac0a6745ed6bc36f32aaf32e80b8d3ad48da36e3669980db30a89810ae854d785d93af345c99b2b83f46805f0e2ae221b4ab41783b35e62e502ee6e8e992c5744d447496c9c1453a87321a22d539f1eadb22e92541f8b363ab943e48d57efb3634a3e7e0354fb98faa2f0e8f8f1f55f7e1fb6c7653e90fddc7d5fdf07edb1f547e04afc55c759153be94fb3a8afdc5efe2a9af9539bdfef709582513063d4e20bab05438c049a6d505c45eacd713eb4d3ea87476ab950432468ac0e841d089e2d2664bdf525c1371109b6b5257b45d2e517ef7033253ab0a016a97290889c1ff00aa46a8a98f4f503a7024f453babe9add9ee29375e5a456b671f2fca42efe112140850d0752458ac314fde8071af1965478f2475657d96de1a779c124c166b5b509d5ad1eb6aee4435e5469b4dd57e736b859b649277265ba96cc076570693a00155247cca17a38c8f55c145a10168af3e8f20fa5319db2cc2a995d69cd6354d20e0ae85454f02e02d53de55e149aed1a535aff0077a5bc7db0555aa5a5f32eb8ea6897327948b9004dd3e4c89a16a954a2f2a2a735706d6ec80d382a259f5d17e5fdcc10af92aa9e25a60df711764c0aa7355c34a53c015ae34321e14cdfaeae3aa889de4a7d9d151153a74e342645f4757d1d58af683ce4e4efa727d90f317557c3abe1f7f1146344bc5cd4e25ad168bb15ca9bc4e215aa28c3024a6bab843c95c2f5d488889c71c325271d70d7338eb86b5233325aaaae102abb30d43c9f297d25f83073e702390e19a08b25d9912741e573ce69a154554f2aa9c95c5134226a4c6bc6bc6bc6bc6bc6bc1df213680f82a6fe03a05e025414934fc3092f5bce4d3c9a50c0b2a2e854e424f34ba3e2c50b4a125084ba75a2a60f8567bb9e65b59da5a9e35a9bf6c0a0ac4225d24e40554cbced2fa183757c94eaa7397929e3c6cc12aa5a489750a72917460596f50eb5e5325ed12f4afba2ba141f382b4f0a63b63e3ae3b5f017dcc76be05fb98ed278f0dfcb1f8d3dfd2e4b06bece82ab6fb72269128f1cc915f4e4fca5e52739f2922726149c5a976407d25fb98e9e55e9c3603ad1f7f69f2d4b97e653156d979c4e7068c93c628b8ca55154d6849454f02e35e35e35d71ae98cfbbc851f3b62e65f1d298b80b9d958724493bec9a53bfdcc88b471bfbf05d5e14c5bef314bd740900f20ea47053438c9aa69d9bed2a817a24b881321e9873623131855e4090da1d5ca568615caa9c8a98ca1ad7b45ca4bf73dc3ace027456abe24d38eaa19f8289f0e9c751b14f94aa5f165c76e9f25113f571a4cd7be4bdc214d5da1ef2fdcfb0154554cbd6f16af871473e927db4c5534a7beeeefb25964bd1f728aa9db47e72a46430f4d91715cf9b8d29a0074ad391134ae0ea9446b453988f4afd11a27727baf75a045d8aecd75392dcae54e9016c3ac9cbd5c5112889a913453bd82024447851762ef9425c88abaf2172e15355342d71a4ab8d78f6a3e28e13864315092a202d12813bcd9d5c454e8a77254a8dd4763e476420a7d7c6124da22fa4df6abcc94ee039c88b94fe41685f12e9c36ee5d249d6f949a0be14c4ab3ba4aa5689a5b0af930a76690d0f3f5646d7ec6a66209e92a27c78cad238f973362bf6f1ead846539cd7adf7d4fd6e3d749f00d553c5d41c758dc2ef507ed2e3b15ef917dd44c7d5078abf1e3ea5afdec7ee63ea5afa03f731f543e0aa7c54c36a0aa3512e94d0bd3a79715ed2738fdb4eee65d67a7e6f277367c85aba17f57df7698095caecc7651a73ee8c651af4669387cd469a32d7a097adf795c38f7e1dd71def6d0d4913c09dcbc5b557f29526263635fac691099754535faa2cb5f958e6c3b21e2406990275d70b5080266225ef22624ce7abb934f1751168aeba7eb3659934a080aa66e5d298d90458e014a64465b445eff0057ad87ee36f6863bf1815e7996faad3cc8255d546d3aa0e00e9d1afbf88ed0922b90cde8ef0f28aaba6e82d35d09b7134f3d71a31753755136f15d84d0aeb376582b2883d2224a5de1ee2a73a2a623bd96aaad3444be9a8ecdcffe2b2b879ba51b9b6e34efbb19e6cc17c0d997733386209ce4a89f1e3f2661d7fd3a6c9afdf0e9f163d63c0c0f9ac2662f0b85a97bd8a9213c5e73c4a6be2ece2822829cc28889f07b888f9a1f0aaafdaee29876b5aa79dfab8d3d81d25fb5f0f75beffdaf7d5359737ddc44392b945b8cfe95e42371bfd8a62ecf30e2668d0674925a764188129c52f9aa898044f347e2c69c44bf3b39cb1a251e882d379e7badaf9460e51a658785751a1e71d6345c222ae6544d25444aaf3e8d095c6ead5c5c821da3685a136a41269047d6a0ee445e445a574d169897089106442b9ca6648f2e74d9a6654d688b969f37b928ddfab08ef29d756546caa9d35c3576f6838c0ca1eb416db150799132c9b7272baf58e54aa22ebd2bdc4911aeaa8b1c5777b6bed88c5cca9d65179ba10b874d6627df44c3d125b26c4860d5b75a34a101a7c0a8bc8ba953b8d9be68882f4e65346a509ef9e5afc977101e8cbb456f7b4244e638eb44f09263ae61103cd6fd6bdde5717a83e0c6651575cfc23ebb53fbed09e0f74aaea4c1b9e7168ef6a1f83b8801ad7e04e555e84c2b61ab4177d4bb4bf4bbaae7202513be5fa9ef9c83a5cfd6feae3787c90535a29ad3e72d79f032995afe52219f4a68eb8ad2b45d6a98e2184e51ddeb876eec0365a515c763100e8ef160179c07e2c426240a1c66334d900a95471b8d4516c9175838f28897a2bdd765cb70598ec8e6370bc48889ac889742226955c3f7eb6c30b75b67086d9a9ce2d6e023aa48b2d0aab4e90e945ecebd2555ee2350d1928d54290c6d15b91272ae6d9a1a8ecc43473ebc2b4d34511f8882d3d05cd04c654ca883aaada5288bdd817b6c505c53dc252a7ee88a06ec635e9046cd2bd2898d09e1c58c740392675e642aa68531f68486c54b5720a53a3132bd6d88b345e6da6dbe340f76d88af5ddd7d01cbf4b57710012a4bfa6abd18a6b35ed973f427a298ef00a7c6bf6fb882895555a226043c24bce5cbef8a27d69f67d14e535c66a6d0b36a5aae65e9e55c23935c26dbe46fcaa7a2df65bf0e9c4bddc551184db6b555f54a320c97e630b88f98aa2e2ab248ba951e150445e8cca989313fe892a4474e916dd246d7bc4dd17139ef36dfb3afe324b05aba765dcd78856777ad6db535ed29edf932241688cc9a6a5411345e9122c7575737761df994a2b4e8479f97f768aef56a7cea14a277d39b154f0531cc9cf8915d2adc988e2574e9daecf4736873065e68aafc18b25ab35160db5adaa73487fd63fe15734e27cc5ad0dec83d28d8820afd2334f07ba9387a93527292f20a74ae09c3d64be24e444e844c236da5497c49d2bcc898e735ed9fda4e61ee38e79c5a3bda87e0ee6d8fb449d44f345797be5f17be146bb477f061c8be996a1f8f06e9e8ae8144d4229a9130fd5115c1105055d689d6cf4e6e4ee4e88a39f6b1ce83e7512b93fca2265f0e25c2355da4294e3487ab36ccfd5ba9d063424efe1a9e1a12e2c81389e6cb6472383e114f0e5c2ac77de6149284acba6d2aa732a82a5531fd633ff00d2e47f098feb19dfe9723f84c5d24bc66eba7b986d5c35371688fd6a45525d083e2fb0b926a44644ff007b79b73f638a24e99a357e54f68ef267c7f584eff4b7ff006f8569e9929d6d6956dd90f380b45aa5448d4742e23b6e53648692245756c585cea85d0e1a2261e7f4fac73a89cb913aada77f2a262dedba947dd6d1c753d32553753a6920ce9d1ee8a44b4114aaaaea444c68d0d06801fd92f4ae11b6d2abf00a73aaf226283a497b67ca4bf693b9b115ebb9afa1bffdeee23cf274802f2fa45ef7cef1a0f327945f2475ae3630c08039553b74e733d4da610dfa3ce6ba7ee62bde5edaf7f0ec77741b46405e05d7de5e4c0bcdae94d0a8ba8857589742e1339eee7ca2e767c0e27569dfa63aafb25f25d05f89709798c35877144cc43451434aa0e94f368a1d0221cf85844790dc5da4535d43207527cffd3af0a86391d6c941d05d60e0eb4ef7374774d223fb2475454fd5b475cb5a7d63674a66c699bfeaf17f80c756753ffa789fc063fac3fd5617f46c3b1a44dda30f0e4703768639879b3047134f02f7170a44b4414aae28e7565cd44326f9588de4017a47f771023885588ef36fbe54a8a645ce19939513229af3882e1a8ed251b65b06813d101414f0d13dd362d2fa915d2bf8424e5f909c98cadea4ed9af6413a7a7a3191b4f944bda25e75ee2b85de11e522e44c1386b5225afea2742611e7d3a41b5fd71fdcf7b5556889a555752615b8bd72d5b55ec27c94f2d7e0c6de6998a2f9df5a69de5fab1fd34c64681007a397a5575aaf733ba24dbe89447daa09aa7221d5145c44e9d38f573c553933b0a3f0a3a58fe36cfd03c75e6d7a058a7c2aeafc58976852275ea13b109d504caf65d202b95320bd9535f64d10b930fc19226d3f19d54d22a05d425442a6b12451d29e492531b78b912f715ba498b541f69c70d4eb5a93780fd3a294d154545512124ca6069da0315d2262bad3b9a31a57ec5489511134aaaea44c05e6e63959d76a807d57673c9a46498eb06035a553a79aaa4b5764c9728823ce5a04053c9014f126066ca0ff0068dc411c5aa509b64f2977d15da268e4044e552f7456185f55a8cd3f74e84f43e3c57b0ca2f59cfd887397c5846da1ca29e355e725e555ee138e16511d6bf6939d5719bc9d4d87327db25c23d213afac5b5f27a4bd2f8bdedeb0aa6bd96c7b6bfb54e9c6cc11727e0c7b029e7385cb8433a3af79cbd905f413edebf713bc5a594f6c4715275a04eb4b014edb629a4e4208d147cb14d1d644454545762cb8e75131a8989268a8af2a2e0464becd8b896880330932daaeea9a0025a7fcddf5d55f157402247bc4472dee12d1a70fad0e4f2d62cb4f52e55392a869ca9f6490e0b0f4f985aa2c40575cf94e53a8c369ca46a889847efee3372bd250e370f45711c8b08e95172e8fa544dc0f37c48a9d645932c95d79c5cadb608b91a1af55961bd39406ba3979f4e1ae26bf31ea7b7022b89f5eb5d0e10aeb8dfe717d0ae6f725555a2269555d08889cab856185a33e516a577ee07c78475fa833ad13513bf703a708008822294414d089dc271c2ca03a5557f4e955c2000964ad1a6935afa45e97c58475da13fe316fe4f3974fbd5555511134aaae8444e95c2b7134aea579757f934e5efae36cf110b44b5570b4b8efc9afc7846da0411f85579c975aafb929b848029ad4b4261d99686c615ee8a5531a44b9aa275b6d9514a2beb4d0e6a2f2d3951cb75ca2390a7b5ae3481ca443a68e325d890d1534102a8ae3707859bbda48721daeead24b8bb3f35acf5369139113a9e8e333477ae0f92bac18ff006b59f32eb5461d072422740236898ff65f16f085c87c9498fc8b3ca34e98ff00976525f15571a07870c3ca7c3880762daf221d6167f162b70e21e0cb6072ff00b4df9b213e4b08cc4435e6442c56e37cbcf133a3ff0034b5c6f63dbcd7941d74977a26fa41d5c15bf87e045e19b71769bb68d263dc95913e82f19d3ca4ca5d2b80691b37e5493a34c3604f4990e1aea06c7338e112e1abc714301b54a391acfd5341f34ae269515fc48e8f3d7586289a11342226a44f72571c24001d6abfa75e364d210b35a08276dd5e452a7c5847a5a54b58b3ad07a5cf397a357755c74b28a78c979853957080d8f507b21e4369e7b8bcff00a531a3aceaa75dc5d7de1f347deb99d2f9209db3ef26326916d57aac8556bcd9bcf2c23b2d2a5ac59d629f8cf397a357b9d17aeef2349afbe4be4a62a4b95a15d7a764df400f9478cad0e9f28d7b67df5e6e8c6e97880c4d6534b7b44517983fc24790da83f1dce90245c1bdc397166e51eb54b5df2ad4911f323dd6382a385cdb56f56b2c52fdc397bb565ae792308ae96e4a734fb66f20ba39d131eaaf56baaf92eca6a39d79b248268ebe0c29a5c6daa09acd26c6514a6ba96d6898abb7ab48fa293a319f2fee6db847c9cd8cb62b55e6fcaba8edd6b93baa736d264a18f1803a73604ee8e42e1988bad86082ed7624e5057072dbe35539515d545c2adb21d6598e57ee52cb79b8c8e7cf24d3a825ca0da00747b9e77574f9209db35e84c6514d09d96d3eadb4f38979fa71997d63fcae79bd009c89f0f76a7d635ec369da2e9f447a71b57572329cbe48a79ad0f94bd3846da1ca9cabe512f392f2afbd55b668e3fa97cc6fe57397463ca75d3d6bcc9d3c80098ccbd77d5349f9bd01cc9f0fb9ab5116aba89ee4ff0027cfdfc6de4e64697ada7b6f74d75a0af3f2e100050447420a6844fb1acfb4db272aeb597022c955d7caf347e72f8f1ff0870bff00f60b4ff44c26e560b2c3cbab75b5c18f4d35d1b26069a57dd55b6a8e3fcde4b7f2d79fa31b77c891b5d6e16b24f35a1e6f8131b36872a72f392f392f2af75598deb5fecd75802f37a67d18de27aaaa975b64abd65fc62f929e8e110511113422268444e8f7a2912a08a69555d0889ceab856a2aa886a2775117c8e511f87191b4d1e5b8bd904e9e95e6c646d34f966bda35e9e8e8f72571d24104e55f8939d571b36eadb1e6f947f2e9f1605f949d20cafc6e7dcf7a29b848029ac896898d84243442d19d1176ae7c94d609f0e11d97432d68ceb14f96be5afc1dd571d34014e55f8939d70ac42126d9d46e2e85a7a649d81e84d2b8cdf58f72b8a9aba013c9f8fdeaae3a59453e15e644e55c654a8329a839fa4f9d7e04c513aad8f6dce6e84e72c236d0e514f1aaf9c4bcaabee59dc5d3e402768d7a3153d5e4369d91ef73aaf3e05f909eb75836ba9be655f4fe2f7a51573bbc8d0ae9f9ebe426339aecd845d0ba9b1f903adc2e9f87146c7acbda70b499787913a3bbb36fd7bfa9007522fa4a9cbd09a708fdc0d447c961342d39a9a9b4fbec203628029a9053f4d57de0b9f8c2ecf337c76e37388cc4b9dd180871ce69ece328ac8cbeac4d13aba34627f115cb8df6f0adfbaed9a83c49c46e4a2dee64782dec81f8519a5a3b25156a69d5af7b12af777bc3722ee8edffd92178b8e79571991d03778cc2497b78942dbae0e610aaa0ae2e3c75c4bc7119e91023dd36f669d3c164b501198cb26eeeb2e4a69ab5da62aba22848d6c929e4e8c49e1e5e378bc37c1b095d6199e97269ab5be1159f5b276892e3449f227485546949cca8deae5cd65bbda7f399c333ac1152d3192c1178bcdc9f3d60db1b88f2059e2ede1194879b53a67efe9c1ccb8ca622456b5b921e0623302abe538e9080d79d75e0a341e22b3ba6d0b8eba093a3a3bb165333ae36d99893800295524a88a695c4d910389f8799824322d6d5dbdb704603373762a936cb9706df26c2522388796bb4cba698bfa5cf8d2cbc64aebf6f569cb3f11c8e2118082dcace0f9c800dd55fcc8a289dacbd184937bbac0b5b24aa8d94e94d47da926b1644c90de34e61455c18d8ef96cb9b8da667198b29b390d8572e738d547c02bcaa34c6e976e22b15ae52b62ea46b8dda04290ad1a920b9b19321b73664a2b45a534622bb70e25b2c76e734122192dc23b9bd4773eae4b08d19ab918f91c4ea74e026dae744b8c372b924c290d4960953b428eb24619879535a60a65d67c3b6c40d0526749662b28bc83b478c054969a135ae243b0389ac92021b4e4894a9718c1bb476abb490f8bae01351c29a4d7ab4e5c459af71470eb50e72be30a5b97bb684696b1484252457ca4a3521631b8887955722aa570f95b2fd65ba9c76d5d782df75833360d27eed2563bee6c194a768a98571de2ce1b225ffbf2d9444e614deb42261243e9eb174b60be47a4be9fc5dc881c377b7acd2adee4a9668c48971ce7fe4ca8d44438ae37a49c4d19b425702ecae2bba137636d67c8b54eb85c6425c23380f4272807209a4dd643ed2ae645ed62d5c43038925da5964bd97ecd872e6c6767ca9067236bf933cdb79588ec9554bbdcb84e3fe2fe2f24b15cacefb0dc4bddd64eca193f758610e73efdc246e4d6f091a8df94a8f2272d31c40fcee3db0f173122386ef12d1c5322feec1acb23475e61e6c4230102e5aa72e8c047bddfed96e90e2210c69124378c8ba9c260733a0daf212a20ae1b62071059263ef46598d311ae909e7ce20d7349165b7d5dd80655a952894c31c416bfce470bcce1f79f6559e1987c6072273c8cd8421bc0d59a3a3909c5598d13d443ecf5b5e1bbbb1c53736e0712bb2ce15be15c6e318600db235ad9744801f167d7392732654c408d79e22b3dba63b16296c27dca2c790a8e341474c1e744c40bcf2a274e1b930e4312e33a999a9119d6df61c1f39b75a22034ef2e3876e964fce570b47b5c39d618f3387938c4999d35f8776539911ab4c44763c9912c091a16c8914cbaab88dc5107896e10edd70722da9bb5c29d3e2ec9f6a2bef39215197818a39b2e44ae2d093b8860596e17be04850e15daf5761b722dde770d8ec5f29ef1edb6c2f16d0887339a14b4ae2f2ff1071bd8788d96ee6af9de21f1295de05bd8dd2306c255c26a3231090fad975509179717f76e1c5a03c32b22ff00ba3b32fc89625053fc937737a56e0a2a9f5797c1858d378ab86e1c81168c98957cb6477841f6824306ad3b284d05e61d1315a7584915342e3db3fde0b27b236bb0f6a25d20adbf6fab629311fddd5eaf939b361bb5d86f966992e46746598974812263fb36cdd7362c34f91d01a02255a55052ba298dacb5dab9af26b045f4975b8bf06289a139b01281a6e55e2e06e47b4c375551a271b1457a549c8a26b162218e64454522241aa56a9fde0bb5fca05b25d4e27b4e74a871e43685a16db67b7b2e0a33544a1a836249a73161f6e45da7bfc3d7369c81bd33365cc82d1be0597606fa352ad72fca130c884a94a9694c313a6bcf4c98573b936b2653ce487f20380821b4788c90531b95d3896cf0e5a1201c676733b76897fe90d811147f9f970132db362dc223bf572a1486a547729af23cc91b654efe1b851b8ab86e44c79d461a88c5f2d8ec975f52ca8cb6c37289d3754b46544ad70b6c2e2bb08cd42c8ad2dca3511dcd9362af67d823f9f464cd9abc9829d729f0adf083221cc9d298891455c2416f3487cdb6936844889a74ae38abfbc1c56836d7189a36df6c5f5120b8e95e19d88c1df656c08c98ece4d393568c39065f1470ec59ad39b2761c9bddb5894dbba3d5391dc922e839a752a5707676ee309cbab51d65bd6d6e53273998c26d36afbd144d5e69acef8a5491116b813be5e6db6bda22ab43325b4cbaf226b565922db3d4f451704ed8ef16eba8b688aea4296cbee339bb3b7680b6ac2afa48986b88e1f135c62dbaf0e330a3db214fb845dd9c8705adabaa2d3e0c51e3155d09cb8e1c75d3375d76c36871c71c25371c70edf1c8ccccaa44644b5555d7f65c33fee997fcf131698b76e04b644b5c9b9c062e52c2df70038b6f7a534dcc900677e7804998e4448aa24894d4b89b7598b252570dc59736de8cba20d2bd2163367bc0ab66ae0656d29451c5f6f324a6a4bbd47be7084c469f01652d3319b79ba4c82b24a12eba8d5553a317ce1096b705b45bdfbcb51b6521b098bb89e58fb47d58205f4ba9a7167b3db55f588c4cb7982c9705d76afc1478ea620da2a663d1a30bc1ad4c38765b3bd25b7947ac2c85bd325c661359905d96e4af52d57b354f4b1ffeb372b940bdc71cd1a4dc2523d1243a3a51256ef141d8f997cb693a9e6162f9c33c572587cc2f13b8a23396394f6cf331658f1a383ae4d80c1fd6473cc281aa9d6c71e4f7ebb183ecf98f535eca344b9bce53a72862edc4df9d4e263821b410b759d9defaccaa998c607588afeef6e840a8288395c70d54956b5cd1f887f363c5cfc5bddbe4b0436e13b912b80a480afc29326209b6f36a553071c5036ebde2e12fce79ed86fb7576058668018a40a4466fdbd9351f22981adc212d3aeb40d18b6de788675d9fb95e203321938d29b69bb7c4d9a04065a17197f6a4dc611ae7a8a7650529a788b831261c88023768aea69069e7ed5247759db1aaa03cac212778e9a698ba459770910f866c7bd6c458545d8c08f2522b2919b3ceca4fb9b9d727150a895d68223843e0eb94ab7dd4111bcb76944ec096d389b2908ebb1a214a8ce13245d91302ece5445aa5b2c57a922f71270fb7c46fd9fd9324ca03b22eb202500c8dea1b0fb8daeead8ea0caaabaf1c56924dd667a3091e6e67441b4b53cdb9eb04082a0eb0eb4e672aaa2228e8c5f76ded063862d6064cab4f363308e54821b630e3e6c10296ecd19b8a809a4792bf619d7d45a6e771ceba32b3ecbe24aa1ad13f70b74e757f78c70e7015b8d577328b1dda2551bb8de8d971d75c41ed3712dc8d1aaf928a7ab0bc1f2524336658b6f868311c069f062d8f457e288384d3809428615eae94c717b435cad4169b1aebca170314af4d1317e3fcecc399259395721909f962a333ce4b6b164ba11cdb94f5b86221236815eaa82a228a626f13f08f1945bb5b0e34f8e3626a3fe596e8d711d9803f25e9e53111ad42aec6152efe9c5bf87ed6b20a0c096cec165382ebff9570b24d733b80db425eba4953aa9a31c0fff00897ffe7f012f8c275de7710dc5b49339d665a3230df7873ab41b465e37de654a846e29212a76513113840ee0e4de1ae217a1208bbd50366e8e1438770d955423cc8935b56dc21d0e360be8e5e10856c5924cccbf70bdd5dde9d174f7a9bc455790141b6911af56944c587fc429fd9b3b16376e0b250b87f80a1dfa0eeee8b68b3a0f0966651fccdb9b462bac532aaf3e38bac93d5e48772bc94790b1cd1b7b66b6f805eacc81c412a8f32e2f5c1728e7a59e03f7a6d826a4363332c02a319de5608097ceea2570e7155a1e98f5c1972d516e8db8e83b146dd16047b3c771bcad010bc051d9cfa4917312e8d5859cebf27731ccf3314648abc5c48b2561ecbeaba9177cf5aa14ceac26bf2b1078fe7acb0b80ccb80d99a07006294258afdadf79f6d5a570c9c75e7907ac94c88bdc53324014d644b444f0ae2c3164384ddac6d36e0132f56dece4dd668cc922abab405157fe4f0d21223aac8036d0d28cb2db6282db6d35d9106c52889c982624b2d4860e99d97db075a2ca484399b34215ca428a9d3866170fa05b26df65390d97610046289185ada4e911f6483b392598010d348e7cc8a8a8988773e2995737ee97588dcc56224808ed5bc65023cd0d55974e44c1034ce44b93368cab4ccb02c61707e670cf101c242175728bd0a7be509b9321915d88cfb5ca15eb8a229b69e4a1e5481c2760952e28dd1fb5084c90ea3afc676ecee590f213431ea2d21aaa0a51792b8b95dacd2ee8171b2c139a6731f69f667371d10e40bad0b0d6c9c26d17228511175a2e2470ef113f35d8506f8e5b5a7197d0650c68416dba45685e75b7bd5b2e4ac888a8b46d105342263882d37329c916c812644358cfb6d3aae45bab319adb1130e218eccb4d1134e2db79b79c9df7884a7dc276d9c071b09114e10b5bb02343b31ebea552c5fbf38977b94bbaf139f0fcbdf23b843b96f53a444942cb8e382b3244929ecb626ea9a56a5d4d58b971b71b5cee121b913dc6196987441f9af34204e99bc60eeef023a1a34d34da0f6689944510ad3c67c15739ccc6dfd22b8d48750de8920c0de6db571b46b7bb6cc6d930701c4e4a2a966d1c017c6836637475662b55cdb171eb7013acd79762ed47c18e18ff000f597fb3637d9436ad832d0ac2170b54ede99069165353884958caebb9daab6ba5698b8c6e137af703881cdd37096fc0882d3592745395994e4491ebc21707b0ba57c387878a25ccb8de389debddb204918f1f25613b10f2c9d96ee2d8083ba15049717bb3c86e694eb4b779e2a7f66c82b0b6d8ccdbdb346dc57848a4d7c9a222f3e2f1c5f2d262dbae0fdd9d6d1b61b595f971e66b333b6101e9eb68c43e2f6c66a5a597a019a1301bde58f0c63b94651e5055ce9a3afab10ff387020bf3b85f8a5a4b809106cb6acdee384a9b6f90e221b71ae0dbebb76d2b42ca9ad33d0dce1e8370997975bfc9d8b94618f0a2baa9db9a6d4a53785bf31a5ebeace3af12f893f3856e62da77927adb6a83688129a7c1a9b0a40c57e7b570b8bc61bc1b64a94a656f2ad17368bef054f194b2b8c01588cad36251d18896abab9376eeab824daab15cb412aae2edc31f9d3e196e6324ea1dbaf5babafa0b6398064362d9a1bf6e98dd093266368d32a8d73658169e05fcdd44e2fb84d7f6489b19d6d656a9d40605f6d653ae57492936000295af370f7e6fe2db52d5728526dfc40ec08439ed311a9316f6dcb6634a23475e3f684b25d2da664eb68d58de4066ecf822c4db779ab0084456fb694a7f724db7af4d9b2b96b92ab8baf173e33bd95749777dd84186ca5a7b4dd4ddf6ad2be203afadd65a62e7325db9f99c317d2968c9b0283b7b7bf252535baba79595b8da8d500db551cc8be4a189601be0fb7cab8dddf26d1bf6ac436a0c7a9a66171a62504a94f92754441446ab5cda32ac6b8f14c78912ed3333e50e1b321918ac17f1765d0932653892326924aa65ae5a55171c4436d5d843e2ab6c8a51172ac3bab8253d91a5104db96c1a279a0a9cf8b6ed8324fbd7fb6a6d53ac9be006e6d2f949b2822dd5390d4b163e1372ccf4d72f4d5b1d19a1341808fed2b9c9b6a2130b1dc57364b1b37692b5a776c5c50c875a1bce59e712255761273498245ccdb2fb6ea7ca793179e34bb194a76d8c996f0e769cbb5df68d21a722a35041e4a7939c7137882e8920a0c0ddb6e915b175ff00caa5b109bc8d9b8d097ae9235eb268c5c63cb194a7c4ee47b75b762d018a487a7a986f2aae86c9ba1eb4cd8916bfce270b5aadf1da07d9877a75c7e43bb469dcbbb14c8d163cc840e02a922a1e5cc9a7930c7ffe7efcd7b87634a71d578b6d52b235084eee8f6d11b238bb5ce2ced7492ecfcaa6078f8467fb0d1f8ae6558edefd959e1d6ad45ea36fb3aef21e7f67160b6dbc65a48b184e726ef0c836da8dea2d92643d890bae29aa34dae7aa25179f0ddbf8cedd7981c436c64234d06a203a931e640476a28e3d1cd87df44cc4268808aba09531138b0edee42e1ce1e7a11891f5c1a8f6b78e641b793b446df99366b8a6e08f600cb992b62e248e337d9fc3d3ac6dcf47180190a564bd392266c1b47885c450ec5486abcd8b1daada33d250bd02fcab263b6d35b8dc2d4e38c26617dc5dbe5943986944e7c6e992e5b5feee7f70ff008b3597db7fdd4d9e6aef3fc473797daf47172e187c6595ce74f7ae91f66d014658ed428ad109baaea10baa6dae8cb8e28bcddf7c6622dc6facb830da47e422c8797286ccdd8e8a9a34e9c1b28d12c1e22b3533bc23bca46b9c4cc04a95516dd00751746a24c0fe6a55694e302fc9faf97daca8369db57fe882d0e7ad29957362df6989d4896c851e1339a89eae3b42da19aeace796a4bcaab8bd70546b31467eceedddb5b8bd391c65e4b4cf4824691823018ed95732267d1cf8472492c934d59f4341f2194eaf8eab88775b233b7bd58d1d14881443b85b9e5127186aaa88b22318e76c7ca4534d2aa386ec7c5f68993dcb68ee8dce8e68d5c90184d98c7b84597911d90d532ab99c0a89d6152a92b1c39c0b6db8da4a5beda23d0e4b8b7d94a268a20dbd136436d8f5d2eaa11684d26819916dcd4ab8dcb8938b6cd2debccbde66caba3ef372d946e65b2dc724dd78922b4d36a023f5a6d9512a6898876de2a8d7262e96988dc357a247090d5c02286c9a3a2bad1c79640088624993369cc95a0c0be376f7a1f0cf0f1c252275108598502414c6e33ee8a6c4ee174946bd41aa836bca8de65b57173033bd956b9768de44d86c65afb31d5de364d23e405ababd64ae379319bb3e37b138dd9a8c029095c2da3298df536dea1366f266a67a2e2e1c20a927dacf5c67df84b643ba6e290ecd0a8af6d33edf6cd2f572d29cb8e2db8dec6693133da509adc5807dcdb7b5c1feb0b8f3088191a5e5c40b54349292b86fda30ae0af3420d2bd20a1b81bb92386ae0656d6b541c5c92d6b219877c893a1c639ad0b6e352a33a40cbaeb40e3c991b98c22e85aa8e2e5c13c6d6cb8476e3cf75f65d65a171e84f38208e81b244def1024ecd1c69d6d4bb554cc248a369e0be0ab64e7a32cf194e3b21b46dd972000d869d50027373b6c36df3371c71796aa83974f0dfe6de1aca3b8f060c46e7be6c83715fde2cf15e479834754889d59198914528ab87388a0469721ae12b48409b1df408a6fccb3d9e0bd2018342909b13479329aa78312af2cdb9cb58c6ba3d6c561c923294d59890a56db682cb0888493694a793f62e3eff000870bbcfbce1baf3ced82d4e3aebae129b8e38e1c453370cd6aaaba5571ff05f09ff00e5db3ff43c330dee17e1c7624771e761c03b2db4a332f3f976eeb31d632b4c93b9133922229513124a270f58a1ef91dd8929225a2df185f86fd36d11e466386d22bb953336b512a69c51384f86eabff715b3fa2e136bc21c2ee385a4b3d82d2797d11ac45d58f66bd061bb6ed9047dc1c8cc9c2d836882db1ba902b1b20114441a5129849f0385ac91a609676df180c29b069a8e3671248c49cede5c6e976b740ba454705d48d71871e6c747410905cd8c96dc6f68284b45a574e1a9d6de18e1eb7cd633ec6641b2db624a6b68d934e6ca43119b75bda34e28ad1748aaa6063dead56fbab20aa4db73e2b3251a25a549ada812b44b4d6345c12d8ec76cb619a6537a24469b9063a3a87272edcc346a52a637bbb70ed8ae9291b16924dc6d3026c846814945bdb498ee39b31525a2569a712e040b159e1419e2613a1c4b6428d166838dab2612e3b2c835244da25154345a8ad30843c1bc2a242a84243c3d6845154d28a8a912a8a8b87635da1449f097ace469b1da94c153555a784c14b4e8c14cb5f0f5a603f99481e661b5b666ba28d3a484e348a9c82a89dc8326670a5c16dd6378a36e4d333486e0cb7394e416f5b9d5adf58011d0259695d38b3dd5f87ece7a7dbe34a720f5bf24275a1258fd7068bd56ad229862e370b25a27dc22a3431a74cb6c395323230e93eca3129f64de651978d4c72aa6525aa69eeb90ae70a25c61bb915d893a3332e338ad98b8ded18900e34791c1424aa6854c1b167b55b6d2c3ae6d5d66d90634169c772a06d0db8ad3406e64144aae9a261d83728512e109fc9b6873a3b32e2bbb3705d6f6b1df071a7366eb684954d0488b8665c3e14e1a892a3b82ec7931ac56b62430e82d41c65e6a28b8db82ba9516a987613bc2bc3fbabceef0e34dda20b28521532ac8ab2cb643232e8ce8b9a9cb8763d96cd6eb6b4fa5242458ad3652134a524399768fa222f96ab8ff0082f84fff002ed9ff00a1e0a44be14e1a94f9a0093d22c56b79d2469b169b4271c8a46a8db40829cc294c34e5db87acb7271941169c9b6d8721c6c03b2d0b8e34468ca799d9e8c37161468f0e2b2995a8d1596e3b0d0f9adb2d08b609de4c1baef07f0b38e38446e387c3f69333335cc46645114888896aaab86ca5f0a70cbc62d351d9472c16b7ddd8b008d30c342b108f64cb6282289d5144a63d9e3c2bc3706d7bc6f7baad8ed64852767b1deb7548db06a46c7ab9fb797456984956ae1fb34198224093635b21479594b4188bccb0d980926b44a22e1e973385386a5ca90e13b224c9b15adf90fba6b5371e79d8a4e38e12eb555aae2e171b6db465ada20a143b5c715640c59c8d37199161a7366001a90434226038c2e5c273a43f77ba2b4d366ccc659b57b4dd6a00c8daac32dba40b738409540af6aa9dc7eeb0ec968897394af149b8c6b6c3627c8592e6d642bf31a64643aafbbd63cc4b98b4af776b7ae1fb55c5fa226f3221b2b2a894a0ef4223232e8d59a9820b1d9adb6a471111d58511961c7a9ab6ce8023af53d255ee6fb74e19b34c98468e3925c82ca3ef1272c874044e47cfcd86e1dba1c58111ad0dc5871da8d1dbf90cb220d8f8b0a45c1bc2a444aa4445c3d6855255d2aaaab12aaaab889027d8acf360c0100830e5db21498b081b6d1900891de64da8c20d0a0a2022505298deed3c3b62b5ca56c9a5936eb4c085215a351526f6d1a3b6e6cc9452a95a68c1baef07f0b38e38446e387c3f69333335cc46645114888896aaab87664ee15e1b992e41677e54bb1db244878fcf75e7a29b8e168d6ab80856b830edb0db5326e2408ccc38c0ae1299a8311c1b6854cd6aba34ae006f965b6dd7668a2d1cc88cbcf3225a55197c876cca2af9a49826ac767b75a84e9b4dca2b2c1bd4d5b774476afaa7a4ab87675cb86387ae135fc9b6993acb6d9729dd9b62d37b590fc671d7366d3682955d028898916a8764b445b64c570a5dba35b61b10252bad834eac886d3231df575a6c44b30ad445130516cf6cb7daa31baaf9c7b6c28d05837c801b2789a8adb4d93a4db428a54ad0539bec95c3d49e355e444e95c2b87e04e414e444ee2ba5d96b57cb5d5f47dcc584f947dff253ba229a851113c094f77a67d98f38e935f092651f12e15407acbda35a9385f29c2a9af7281d7ef7653be5ab478f1a69de4e4f0f2fbf289f541a07a79cfc3f1775b4e53f585df2d5f7b4f73335f28957c1c89e04eea2a6a5d3eef533104e72544f8f971ea9b5a7e11dab63e0154da17891179f1eb4b69e8f65bfa155afce52f7ee54ed39d54f93e52f7447ce544f1ae29ee4efe2cff5abf608c92f5c3b3e90fdd1f71d2ba57526b25ef0a695c6ac89d3453f125453e1ee55c300f944898a4669d92bce0395bf0b87444c7ac70230f98ca6d1cf0b869953c098cd97339f847155c73e915553c1eff54e46fa89fb2f87badaf318af8953dcd479d153c78515d68b4eed5342a6a54d78a3c99bd31d7e14d4b8eab83de5eaaf8969f62a23570d358b7d6a7ca2d00dfce54c7ac7da653cd02471cf095289e04f0e3a99cd575aa22d57e51b8a84b8fc9a2d7d22a9278699453c78f5d21191f35bed7ded3e35c6624578fce75737def671444a227226afd00af36155795557c7f6005cb4a17ca4d7e3f73cc3a0fe3fd5c50929f63a0887bcaa98ebca21f47684a5e01d2b8f508ea27e11f356c7c01a4cbe0c7ad7a5bd5f2015c6d8f088af5be72ae11282d8a6a422d5e01cd8f58e92f40265f8573634342abce7d7fd7553f410be4afc5f63ce05da4fd9274a63302d517f4f8fdcf4a22f7f1d9a779571552214f9434f8531a64292f3050fe144cb8fc9233c7e9b84281e1a0d3efb1f94c9211fc1b0b97efe89f6f156e3066f3ceae1f8cd571d5011ef0a27c5fa1249c9ac7bcbf63502a73f32f7d31eb429d21a53c4ba531a1c1ef2ae55f11517ecfacf369d1992be24d38d0a47f247f6d971ead9f0997ec513ede3d58a8a2f288200fd373eee2b2642f79155c5f19684c686f3af9ce75d7c5d9f83dcaa66209e9120fc78cbbdc5cde6eddaaf8b3571512424e715454f1a7e80f32a6a5fd3c98a1253e25ef7d9f55553bcb4c7d739f4cbeee3eb8fc78fae3f1e3eb9dfa65f771d874fa56b4f1ae8c75b237df5aafded531eb0ccfbdd44fb6b8ea3408bcf4cc5f48aabee1555a273ae2886af1798c093abe31eafc38fc9ad8ef4148216bc60b4f8f1d528d19398727c6a8f16173cdcfe8ef6f227d1ca8385328eaea72ab462e97d145da2f8b0a242a2a9a151528a8bcca8bab19e3baeb07e734e136be3154c350ae4bb517491b6a55111c0325a08bd4a2182af95ad397f40684954e9c755547e14fbb8d0a2be34fb58eca7d24c763ef83f6d8ec7df07edb1d9fbe1fbb8f213c2bf69171a5c14ef22afdcc759c35ef220fedb1d8cdf2897ed5131d56c07bc29f1fb975514fbdabe92d071da1693d14ce5f48baa9e2c55c457979de253fbd5ea278b1414414e644a27c1f62b985024227ab7d13ac8bc827e7060c0928404a249cc42b454f1e1a6c3b66e800535e622444f87f47e518ea27dd54e94ceba7c38df9c4ea355462be539a94fbc09f0f7bf4550e4becc7052c886fba0d0a92a2aa0a11a8a66a0ae28973b7d7f9647fe13dc1507eb5c4510e8e73f9bf1e331d45815eb17297a03fa74604011040510445352227bc464ddee76fb5463751809172991e0b06f901b82c8bb25c6809d20689506b5a0af3637bb4dc605d22a384d2c9b74b8f35847410549bdb4671c6f68284954ad74fbf0224394fc748cd26d77779c6b33af50e85b351cd91bcb4e6aae3facee1fe9b27f84c283ee13b222baa066e129b860e55c68c88b4af28fcdfb18ff00ef06bf9b4bc35f8c0fd727d9e65f173e36af680f8d3cd1e64c208a208a68444f795abfc5707fb22fb87bfc4171fe6d6ff7dbafb8b46d96cdd35f45b1522f8130fc973b6fba6e97467255a2740eac30e98d1b92046c979c20e1b45e230c0324b46e68ec179b69da657bf9932fceeeab022b2a5a769a02ca0d7e35da150ba11157bd8d11a0a0d75284852a7cade112be0c3518e32b0fb7281e5213cec908b4f82d2b4302ab89a34f7f0d7e303f5c987e636c6f3b044326b69b2ea5510cb36cdcec0e9d5ab1fd55febdff0063c313407223c2aaa19b3643125030cd41ae521e64c3afb9a1b65b374fe4b62a4bf0263faabfd7bfec78df0a36ea24e18363b5db6710a229d764d53af54f0627f0eb5c3ce5ca5dbf75da487ae03122fe570a34e1c8011a4bae6509288b5c9a71eaf87ec42de8a0914f324d1a7ae925b4d7e8e1b63892c276f64ca8b70b6c95980d557b4ec275a6ded98794a0665cc3c988f3a0bed4a872da07e3486490da7997133018126b4545c4ce21dc3da7babb11add37adcf3ef525b8f5de377959726d2bd85ae2eb73be405b2140763330edf1e7a5da7dd49f0788f776b73b7a342c6c93311aecfada491688a5ec9e1ab7311515507da3264ca9069e4996ecb0db69553c9ebd39f11ed7c4f6d6acee4b741862e711e23b7a3ae2e50196d3feb6236a5a3699cc52bd6ca955ee3f36748662448ad93d224c8705a6596812a46e386a82229838dc336572ea004a3ed19ef1418ee517b51e20b4e4971b24d4ae2b25e8e2b3786ad3219aa7522ca9911ca69aa6d5ddf86abf231222c18d3a05ca1c7495261ca1030d8ed05a571894c9283a22e3889d646cb4eac48b95ce5b3060c46d5d912641a036d8ead6bac8896822952225a269c38c70c58525b204a2370bb3c6c8bd4d199b80c223a2d96b452750a9ac5302b3b86ed0fc7af5c223f322bd4e5caebc730117e660e4da8cd9951b2a4eb64ac8332229764e80442f46717b2e0e85e5a2e8ee6e462776be2809a5aa2b82dec04d109b39f2545c1888e02d44729b8a945cb45ae155be1eb18b39baa0673dc7107995d190d09174e44ef621da1eb3bb6cb9c5be45b891372065417a3b502e518f291032fb2eed258d0329a52bd6c3dfe20b8ff0036b7fbed2282f5e61e55fc4b543729df2ca9de5c36d3695374c5b04e72354114f0aae226c93faaf6615a69564d05a717be4e6525c0380b94db213024d6242b5154ef2a6234b0d4fb4274f34a9d70f98754c49943f5823919fc738b9017a502b9bc18233552225522225a9112ad5555574aaaae0264e79c65a793332d3283b526d7538466842085c8945d186e5477df343902c2b6f6ccbb4d3ae66ce02dfe0b9b970d7e303f5c9836dc4cc0e01018aea2134ca48bdf45c49887ad87482abe50eb6cfe782a2f8712e012e96c924b49e89d1b77c02483f4b1b015a3935c16ba7641eb1d5ef68415f9586d96d3338e983603ce6648229e155c478adf658681baf3e54eb177c8b4e27712dd6ceddd2e93ca32b8b3dc71f88da4586c426c5b839922a8ab5186b9c4d7378b083ec5b1458e3d541f675bd9646be4a26c4412b866fbc3e76483786e6b4c48876a7e180cf8af2189b8b0231e547a3ba82bb4114aa2ae6ae8a4e80f399c6d57a7db8c95d2d46971e3cad9d39077a37493e52e2f1fcaed1fda71b1729b7992fa5becee4403811bd59cd7250be428e4aae66590d869414cc55d04384891f85387d1841caa2edaa1c93711528bb67a4b4ebcfaaa72992ae20bd656121dbaf50dc91b987d4c69919dd9c918e8bf56c1838d920f924ab4d1444e18b8c9357243b6a61a7dd2ed3af43cd09c74f9cdd38f997a5710783223a4111861ab9dd90169bc49788b728eef2a8466836b4d4a4e0aeb14c5bef97fb747ba5f2e71d99c2dcf645f8f6b62406d23b0d467909adeb6448ae190e613ea8d295528d3f87acd2592151a1dba2e60af2b4e8b62eb25d20a8a98b85eac725f18171b61425b649557d623bbdc790851e512ed0e3e56a995ccc68be52f244e0b826e6eb6e08cf4c8ed2afe57769c22e466dc04edeed15c0c9e93a5d188bb485166f101b4d9dc6e8fb40fb83215331b1049c42dda332a5953265572952d3a9d837bb6c5b830e8137ebda1575aaf971dfa6da3ba2ba881515170d5b81f328f16f6dda2412d1379b35d0d946cde14eaa924590dbd4d5b404c5e6fce2217b360baf3405d9765151a86c9694eabd29c015efe2648be4a7dd88ca95dafd2b351f96e3ef7ab88069f5252dcae94ecb605968b4c0c485c3b6662388e5c896e8a4a7a29575c71b275f354d646aa4b8b7dead564816cb93fc411a0bcf40652203b1deb75d243827198c9154c9e8a0b9f267d1af4ae1eff00105c7f9b5bfdf355d098ca1a07f4ebc3ad82d5b869bb0f498e9797bfb5554f9b86263acabe2c2a9a36848153caa80b5512ec12d7c187e3396d7323ed1b45f940ea3151afd4eb4af7244025eb473db35f8a77b689d00e257e7e18e9b8348bfe8f297e34ee3222941169b444e64404444c47ff007835fcda5e1afc607eb93b91ee009a1f1d83cbff002ada55b55e936f47ccc4475568db87bbbdcdb37ba955e802a17831bb8af521348df46d5ca38e2f89453e6e1651255b8219ff00cbbb50693c03997be982332100015233254111114a91112e81114c158f824e6c1b73b20a2401b581a5e6eda545242be23bd46171133083591447b6abc8332fb75b7c275c4a98cf9f2ae7701554d4e2b0dbd1d57bcfae1cbfc8e20627ab7261c6dcd9b738d2294955122de0e59686e9a3d5f5ba31c4dfef589fcd17178fe5768fed38d8e2efe5767ff3370ee7049d1332a7118a972a88ad89453bc8a6be3c70aff2497fda7371c4692955b0de6d4d12d1532474b4db444d116badaeb74d6b810144111441114d482894444e844eedf9a8af2b731de2061888fed559565ffc99a8c7b74a2b3b22cbd6f2698ff8aaebff009dae5fd231ff00155d7ff3b5cbfa4619bb5c9d8d3a66f315e7a74bbe2cb987bb6c85b537dfcef38ad34d208d57420a262e021d97ae16a6ddd7f569301e4ad3fe55a1c715ba94db9cfb636e6acdb26a3ca266be5533bc74e4ee5abfc5707fb22fb87bfc4171fe6d6ff7c299ad113f4f8f0823ab913993ce2c4993f8068892be5babd56d17e538a8985225a912a912aeb555d2aabdf5c35312434c0bd9b20189a9651250cda3ce51c7f1e8ff00bdb9869b75c1751e6d4c1c04541a896520eb728e8f1e22baab469c2dddee6d9bdd5aaf401d0bc1891912a71d4250a74355473c4c91771962e046c3cc00b4ae6ccdc6dd104a09faa13313544d3a35e1a89105d246e40beaf9a6415cadbade5105ebafd6eb5a6ac35f8c0fd7277254744ab9936ac7e39aeb8a27cba65f0f71c79d25371d323325d6444b525f0ae19224a3b2ff002a3e7ca689b14fde9117beb8e276e2577b7387af411a9af782b6c91669a174ed1530cfb48d96ca5db26c2b6baf108085c9e38cada09968471f8c0eb43e729d35af70f85d8b8b32af6e5c223ce438a48fac46e311139be380aadc6734e8055cebcd4d38e26ff7ac4fe68b8bc7f2bb47f69c6c7177f2bb3ff99b8773823ff12fff0080c70aff002497fda73711f8805b5dcefd099127913425c2dcd8c571a254d4ab0c5951aebd34ece2012496fdb56f8acc5bc435244900fb002d2cb46ab9962cca67124aa22ae5d68b8aae844d2aaba913133876d1386e336df08a6cc7a2657603282fb51f77ded0b2bb2333c9a1bcc234545542d185beb426112fcd479d19f1ae56ee109b6634b650b91d15681eff002be28f21990d05e1865b0bbdb54846447922288e3cdb55cc709f3d2d9a68a685eb22a262dd65937269ebc5d2e506d71ed90c824cc07e7ca6a236e4b6c4d121b0d93b525714572a2e5425d18e24811c15c9230c67c701aa99b96c7dab86cc0511731bcdc75044e552c5cacd7392dc58bc40cc6dd9f7cd0180b8c127b62d119aa035bd3324d2bca6223ca9dc81c3302e2ccebac6bdb17194dc4247d98acb106e5149b7e436aad0c957658fabaa92222e6a68abdfe20b8ff36b7fbdd489682295555d4889ad71a34343d84fd92f4afc18435edbbd6f9be4a78b4e1a8b0d5941daed1f575c50aa027ab14a0155148abe0c76a1fefc7fc0e188cdf618681a1ef00a0d7beb4ee342c6cd2430ee70571550540d32ba35442a5742f831da87fbf1ff000386025e45902c883f95730198a6522aaa2573ebc1bf6ac86d9aa96e846804daf33466a8041df54a74e32ee5939cdc798414e9d0e2aaf811703b279b76793c0ae66226e3b6c64733037d452335351d2a89f74094a1d048557d71f22d7f03dd92ec358c91dd749c6d0dc2121da759472a36544125544e8c35bc145d86d036d91d353d9664cf953649d6cbab088894444a2226a444e4ee49baf048b12a24b749f3b238f3511f86eb84a4e0c275f56a2390f376448c0813aa99b090542f91e2aa236492f895a6e1837cc4d25ccd48139840bbd87946e5126714bcf46d937b57635a214742aca1db6c0e44c7c93512800a79bcb8bcc4bd9412767ce62431b8bee3e3b36d8d99675718632966efe27d8ad4b1866c97e038dacb7099632c698cbee66306dd245c81a346bc5f9abe14022b93f01c8fb8c871f4cb19b942e6d368c31956af2535f7386fd8676f1f657b637adfa438c577ef65ec765b38efe6fe2655d54d18b258ae2ac2cdb7b0fb6fac632718cce4c92f8eccc81b224c8ea72269c3f65bd31b68af6530305c9222be1f55262bb45d93cdd7bca8aa8a8a2aa98597c293d9b936046b11d8f3bd8f766c56ba0f6c6c4712cab4a83fa7993090ef0e5d9e8a5afdafc5f1a7451f94c8dda69af8017132eb72bac7993a741dc4a24265cdd6386d9990a6929ed9baf9e66a9f54094c3d66bd47db467151c69c0540931240a2a372a23aa85b27dbccbc8a8a8aa8a8a2aa98593c2d708f7468094a2b812bd91756f990b6a6119091346617f4f32637492e711b91faa24dcae338ce464451a2750af8626282be4a2d3168bf5fef106315aae506e8dc0800e4d75e720c96a4832fc973766984236b4a823ba3b926f1c13bb38ccb327deb0baeb710987cd731fb39f75422ac732fdcdc26f67c8aa9a046d8b16ed1e0f655b95c411c2ded87674b297134214e60025e8c35ba5ce14ce26727b05211c71d8b698b6e48d2b6ecb07b075f9520a52b4b9c85b4ca9d94e572cf7828852caeb2e622c378de6b64f331403ae6d32b9eacaf27bddb8c2bd67d548ff1614d1f38be2ee2226a44444ef7bdb8c245bf876cd26d3c1d324c69b2267103b6f94f8c486d4e709a6bd96fb02a4d3b44a9f6b1656f39469f7cb1b3c4116daf01ef0305d6e39b8ae1206c909839222495ad70f5c255d98662b174916433217f32dda23a6c4880cb28d2bf21f6dc6cb43625aabab1607ac93adb3225dafcb6579c90b358751d08af487188ad6eaa5bf0644eabdb31cbcb83b63bc4509b9cd4f72d8f32432536139a92510d890eec3631d7780544532112a55169a71b9ddeeecc59491f7b3611b9325c662e6cbbcc918acbfbab0a5a89cca8b88a370bdc56ca6c31b8461651e9aae403ec4ea426a428c373c971680545a2e8c44df6fb1437e80d5d236cc244acf6d9088ac4e248acbdb18af22f54cf28ae06734e0bb14e3a4b6de6fac2e4726f6c0e37e709b6b54c47e25b3f0c582458a61994580e7103cd5fdc880f1b2ae152dee5a9a7fa8a5b3576a9d9d78f67ddaf5161ccd9b6f38c923ceac765d316db766130d3a109a3224a13aa09cbab0edb1c58c9676783cf8a1cb80e770f2b73d63928ab644d9c6dd533f54555716a78e600b77ca7b28b23cbbe6684e5c0722236aa3586c91f5a9ab9f1799d01d871ae76c1ba3a112e0b73dc961c09fba353a44c6ad82adb729a517366224e067cba68b88ccdeaf1161ca910c676c046448318cb44594611d975c62267d08e38822b4efe20bb72bdc38eddd22efb6d34da3e33e35584da4358edbbbc66de415106a4a2b54d08b863887db70d6d12646e91e526d495e979887746e30b6b2ce55417d5e4cf44ad29839b659cdcf8cd487223a602eb6ad49685b371875a781a75b7441d15a2a26824f7bb55ecac40cbfbebd9bc35ee04792e236f8220891ad01d14d03d65d08e539f5fd8f59e0ef22e65f10d571ea99354f38fd50fdf7597c58eb9d3d16f47df2e9f8b1a3f57c2bad7ddb8e1de24fcdc711f123f70bf3cfda26c0b5ba6d1c228515b006e7ef71518ac912eb27675d71c0779bfdaaf97e762f0adeac972f63b057bb8429326e2cceb6312d1b3aba810805857aaa1b40aaad34e37ebad878ded93d78ff8aaff006e7f87a23526f763f684a7dd61d910bf290970e4b2f281ecc5d45a68d0b5c706cabcc3b8c9283f9c69129a992ad2c5b2e25c3cdd9de8f16e57a830fd5c378a4290759017264aa63f391019b4ca3b95cf8bf88a7428c1189654d61ee276a4c592c820e7740e30e602f371c5f32e11f8f1eb6f120db5d852782d86673721b8f6d180fdaeeb1d63487e39b262bb322546540d6aa38b548e1d83f9cbb25e59e12816f84ec1b4c1e218721964e4bd0ed37f6762f47666c170b2912ec8100bb4b862ed1a171ad9f8da470b5aa3cb5e1bb142bd70b5ca6c78a46dc198cb8dc96209449464d97580043b2ab882e4d88d7b65bb3c5397023136cb1ed31840afc460d33b4d35bd2280ae9144c1bbc19c03c75c1fc70fc864997e3b532d3618d245e6d5e7664a4985c3cfdbf266d02de6753465c7e70634de16bcf117f7bf6322d936d56d59d0e6a1599bb6adaee6607fecc6a24805a2bab97646ab5e762d8ec733783f332bc3452ab587ed927d4771ded6adaa857e8257563f35f6f5e12e27b72f0c4a646f732e56d5890587a2f0dcfb62a4774dc572530ebeb51784765451d399698e3de0bfeeddfd6e673389e6c292ddb9c72d7708d70bd94e89b9ce15c8fbef3527ead1332655ad31c45719bc337bbec3e25b370e3701cb35b7da64cbd6ab7bb0a559e7b61d686dbce9ed10dcf55d65a968a27e6ec2e36f515b4f08f12b52cc536ec5b665c2e76d931a12c9eb08be1188dbd0ba7215155317529969e3160d3f399c457cb5dc78522248b8c0dbb008cdc77271b7b7cb7ce475c05cad9e2e0e7118c9522bb48f65cab95b63da2f13ad62c4608f26ed6f8b46d89351504aa09a8025453dee8045b27daaab0f225722ad33090e8cc074d38a6eab20791c8ca8ea2fccd0ea78471fd5d3ff00d0e47f078a32c5cda4f37652003c4a2838d4ea7e33771fd7f5b1eba70b69e8e92fbd004f871f944991217d23541f1548be1c7ab6807a6952fa4b52ff00d577ffda0008010103013f21ff00cf280aa004ab600caba05210665c2ce85a76aff7f48225b856d2cb9a00af862e888fd9c191ea508e11e8ff00c6b36e9bec019309260a953ac36a46681ba452465ed401be66253bde85081d6486a8172cd0e086249758438a160fb09a98d03536a9a31ea1dd531b060430de0659dbd5b4c8415e6516496a4a5141be4abaa054fca8991d5b844d7a8b837bdcf4a4587266d75d515840d17a2d747714388e2e217012227f0c926cfb034f7553f456b5725ecf915f28a3d587b56c53ca24a9fd43fa0c57ca0568a14f6882fcd6548d89f725ef5607f308c3dea3893553ff00054748241daef4ba218d524cff003d633a00b43808b55957800e559ba869453fc925ec910f3353e0dd5b52f1278e0285b860128996a11b425446741cc88dd827dbad338841bae317868a20c5884127295fc8a77a615c0106a417a55bfa82f1f49301a0541075c4b0cc15914206d5c91c0fd652a866f859b06830dece8b5c3538aa1821daf190594340f4447a48fb52f0a4c1ee192b460e991e2225603b63e4287f27f3f0a9167d2fec4ab27dc6b07857645e8c54f807c0526adc1b9ec0d6b0c85b7df91fbf01a09a5a3b887a8de28e147ec9f4932be5355acf22168389b03b18dd6af7f17c24db56c64b7849157e54a7b9853c0f37512ea021008459d9f0ad885bb00896a22d67a54e9c923cf06afb64a0f8c6b0ca938d1e4bc92cb12f62751a7f358a27cfa409371a37bc5a3549e31c618d1c2824588116ec488098a8d2e0248a331b9912637a4bc55e2e4b28c6d50c506ecb894862643126f5e665eba9103ec43d83f983c2401e8daa516dff009961d22a71916efe57b8a9a9a9a9a28a4d5837fd1f7c3168420604ae410028a85fafd5e7654620ab5176499182e8ea58d9deae732523906a27b2266042040001600b001600fe555599f7d2f92ca05212c20ccb2ba698170cb25f55a3aaa4207f500461ab53989008502af500009a89c95935edddf6d0094b7097da09eed8a03f5a3f926fa16d3ea0e98af1046a7a44dcafd6cfc26bfced457ef5b17ed9f461a5305e61fbe905910b0d8ba24ea3c6a7db43f5d77804f4a88bdd5d35597bd4588c6e428dd9f6ebb7b1cb362a4636255db30254bc1709383f35268e8bb4b820e6f21be9ea3e17c19b02e325b6b3441b59a68f710f44b9423fda6a4a595c36c1036e929526033652c93f70a8a88d302e75cec3434fa197c367ddbd95eb200f599eca57b887c15b43b7bb43eeaf74d3ecbe03b41c7c07749454a9582326e1fe86f292c126c5cf61ed40904b8971fbb69cebf0b5fb3df3c55d21989b9276280e450da82eef818f3767b4d95bac86244402040000c00b01482828e0628d7358635c85060ad38288c246e35764f87cd5b8f2f5a091c501a4328f18d733592d44d16c45e148d0d65d4c9e1d9f61f5111ed765059b628c2a5e4b5416982be4b5c0f10834fe3c93015e9225aec4417ac18e834e5374d0eca8765b9f36dc56593a2f497deb5ff3f8e54c077bf63c51a6cb790eba01e6089c427f66b1c8f9a3dc3dca2a55e0b1f23f6ccd14424723c525ec0fbb9a7030728146e637874ab2012d0c001bc141332183d9d0870ede1a0889d13baa4fb7a94b436d69857e20fd06c140ec134c4ae1d5c88501952d6ec8446dc97cb2b342da9d1c4b21e9b42222868edf49b91d68f608061a013752c4b100dcd78911d3c39bff00604524da39351eb27bd36134916543a1f7f0e32fb5e0912d696d3f02ca38a39ead6c3b41ea537f72d1d6c5e95fe0940807d18db4e7aa47b2a68818321c6e834f9d24e30f23a7e4e28ecf02523fc897dbeeace8f61cd0be4243461df0d64a905981b3652980250bd0fad67d06c53df72029c322901b32a101044110b008af05aa336b203609014e86482266a41c6148490ef2a834a6ec98a17430da0b46b434b7902aa267386d9695cb52165f0c2146e2286ba03dc91c81705408a33e14af0fb53e45d24e8f14058d151c9d822deb566db8af856e4b5baea6a3bcc9fa0fa880a402ab802eaf429b4e4781b1d89e02ba573a6b2d02b24525c20b9d43d0f014478c83b56e93f5fb9b96217d42eaf3d0a5785e62f7956745591717144481a9a28d3214896e8616f4045a99ea5a75265584262f354a8c09b588d284b49ad31a2c09074a7259405e8dab215b67ba8bb0a82272da873187204708082c19d280e0111c848d18048b85a9d4358ef5194ae5c0e8c4a309495629afc8b14d22d67601dcbbc634a48577677609fac4f0831ccff29d93534c35b007ba601ab50145ec397a86850b1df3aadfed42816430355b1404ba75065d3438fb818e46c5d1b038d377a353213c2443b386773176b29f5681b034ed2399579c19c9b67703a5ea18e1fc88d322acab3ddc4cd8966ec0b868a0895bc8844dd0fb3158af58688b585f36bd2cc4e0366804001000003001802b25ab25aa784f4234c012e822cece8048224aba46e2264692c722981c95061e4ce6bf6d68d1fd1243bb463d9bed1c14d82b79a9713d9c07ba43d43ea94ab3c87d41e85e9f69991a60ec957b7269abf7551a20f833f19ef4a05580155c0196b6b11e8d8ec4a28a27ac720cf8bd3abee30a39bc742e7bbc299226bd62e632efd56a600080bd29ba9d1e16840eab78f94502139c00f0767c0151f02d2413c54fd49a1afc0d26457077f074b6a614f2856c43c0089710d2b0dea4b49a04f9698c5ea772a30e6b1630400401001100a88c50d90fa6489169b85acd2642d862492f88568bc31af30822ff00694c7da62e321c82d6b1fa81f187a00955d82b334855a773e3ec54a4dbab62674c5611625c0f8743c0f10f1ce267dbd26869f0fb0dd87b6c6bf6f152df590d02efb1ad2276c987de7ec3c4b524f569db5ce4b38a2813d245d806e775a8d12af4db31fd0dea6e970dbb88aee38a0e4edc3f55a51f5d2cec92599be62f5b06429a9be9048e719a497643c11ceb2b5456053c547da0dde277902c454452b6bb77f33a0ed2eaff002aa7da777445c95320d3cd2b5255889d34375706ad014b2e50190da5c9e78148f915be83a8c1d13295a9e3c927375c8757e9a80ab01756c01956a4f453d67b8eeced50ba2244f5032b4c9f568bf33ba4fc183c2e0aeaaeb0fe762aef223a1b6c058a4d3ec75f807dbd5b7db223014000baab600a8626704be37837f755af0d8e801636edd217a346ba0badd4f2957c122420c2103e0920584a69dbc7a2708ac1ac5ea0a62d627ec493502b0d7743b6519c4385081752093930d3db73747758c88c8a76017049e24423ef66fa1b331b89621d6d25c6b0deb197b58d99ac7359af53cd65a252ead00d55c5124eacd0c110c99a223c9a0e62c4b2c2262082f017a3165a8d7222165d5e4e01f4d405500255b00655daa63e1704b3d8f94672c5ca33b83fab56c8bcfd9599bef00ec294d760cab054de2063de64b5b273faab7b74326d3c3b4f963ed66500984eca9eeadd6a01d319d26597a7775c15646ef1779ea6ee88fa274166fbe4182320944b2a03d7561e328d924752ac8ad8c61b430d339ca86c1e8d7807aa4e4201a0047ad73d4a335cf52eb5860d4b2625bc1b06bad5ce9d61c3c9b1118d5ea233bcb9ad399eb574dd295297683d8821d0e2fa5a247d34c31d40012a3600a99ac776164bc3c39d5b5426b1bb1691ab772637a8f4b1401a01e0758bf000651602eb5bc39ba36b1651d8f75224a24dce9c06fe8ddfb404cca001754b014167b8148eafe070e6b3ad0cdc67939db69a224b31dd5bbddfa4456e5507ed7432d6b380ec9098f056d44a9b144de20483348c986a2e6116088b53e521baaad2645416ac9da0d8e2eb9c0a460c84a00c6a2b165fd348e5c60e69e33d8b5e5510851b37c56f15ac358257f8e74e34610d0fda25752a07d565d207d951bd30a5c1b0ca96eb23808540800000008002c01f48f584aadc06aa70176ad58214c9824965c192776a226daa354643d1359d3c0436019d2663f9b52911a6464b1a559af6156088e15db9eee9aebf6b6d459c84db6f75b140c004be33a04cdc46c542e6d1d8aae7a06b3a62c7d26c9292dd58ba3df742b5a70880e462e67742b2c931d4f6360b152ec8b2008b20b5916c5011ead8695223626190de9a4fee71a60ecaee34b350701007f5a1c4dc5d52f75b45260590d2a541367e1089c657d544005902d92b8ef1b4d47891b51644e7e989ab9c08bd73a217049bfe9e42536c834da355b1445ab4a130f456655dc1b50a890b85a5977ad97b78ca08576751f23dda96d7c1149977777d4e0a2a63c9a8cf4f6d3ed6d194ea272316cc6bb558d964c16aad33c050682f19691737d4cbdbe9280aa004ab600caba05143ae648dc2d973b6d39a28cac04debdcf7268de82d74280e03f8b6db1101ba8dbd66fe071669655269cef4cdb57eaebe10865b83276df78a6095b2b072ad0c4fad8a2a31cb3de56f15a023204ddb0d4d161ced53bfde80c1836778c51c470200c00b01f687c455c02ea2002b27c6e787872b2e35b781840ccd56ba593d25a097487ba8b4d82c7d2cae5751d033a00bd24b8203db5e87436eb5aead8cc6c6ebb7af6fb43737200777576d6a4dc1f4a1773a36ae7e112f36adc07aa8000000401600c01a07865416baec65340bb5b90414b59421b9ed50668bddb770e72df4fb50e3ae73a06bbec52b91adb78b673d0fd9593b89893e5e69a6b4533b8d56442cfd2ced24c03a0d0355b1529ec627c2d80675593d2d5715dc1b2f5f35f1f668acbb168b370ea5f61a392924114dc9fc822ae241103aa8765078be94ae4d7609d96f74566c51b04a6c4a75cad61ad1ef6875754d56efd82843f446f5998804b2bb5c8a7585fd6112c298436e9123e20b34d66e848c679300124922a4989737bd2c117c0a5e28bbbc1d3e188cac28b46cc7e02e988284bb514d19692e23c3dd08a2507a7cf84f61041c280a98e8559ae408b5a20bf29444788ca68bc5323495a5ecf6d08f46880989e47543ae8a86d570eae819097ac9748f64a3b1b824c492590a40b12028b025a85d8291ed4dec1841a30dd65131cc5fe5c47685152af230a449b8a9302d5682863162d1e55ea0db482c5c0bab4fcb1414ec77c89e1726b4a8072e07023870691109425f06122ec8304b63428a6369de3ab26901997453d51d211940bb23451ee792bd99c6d97059a05249fc6c96e1f4a434f07d0db1c712327132b44c7941522d20bcdda1afb5700daf6605a68451845726ea855be1c1ee686e49f0a5e342e4a742864ade825caf97cb3b5de100a0d726baa15a33a5e93444835c358bd6c42554d1f4af15b8a129034b9445228374a022372d14d65a7100939ec955054b80673aa5d5b27fd5000000801001800b01408b4450f2b10dca5566349edb16544b28e471691f8f54984b789441064768454b81a96fea852004fc9054bb2998845c6198b1ac279ae72f8790ec8a84b1edcc0c65db7b2ea446c4029a6e51bc0291165ad20c27b19a9500dfcd2f623c8dadd8a5f7682249194d062688daab5690608bdb50298441acba48fa27b71c218a046664de98382d80312882a65fe651b431aa260783924ac83f195a0dbd9103ad0b3ba18b8f319cd74820eb6e46f660d9159b66983cde5820ab54db58a5b696b0ce04d6f1967730365f1be6853cb1c40eb92f221719485c6958759c82ea2ac01bcc568ccce5de5adadded39b364c7c04aa0bec2d6a6a96d24b26f48bf952ce671d8a009d06ea1566bbe10585940d17b5116e355e8f8f2b5d64f3c41815398ea79733998b0f0a52d056ef0df8b63b512410d76c50d586cdd03207f02352da44415c20913c5c55c802abec49783d40b03d058ec20a5fb85a64ca7aecdc8801c8ac4821a689331766c54c851f4145344630950a9d598cf1981608225f3e04017b1f3d1cb9c5db38fcc6f4ba1cd92001a19b8984501499214ddf0013157bb48ae45da050266e1555910bb5314909eeefc347414732d849390df00128d0fc40d70c40dc242d8a59d4de255cb007243e010ae40eba0293c00ac03db78452016b5d1a070932f62308160ae5c3115b1b40b01d2918d74556fa0349e85a9ed0d3ee732189d21449d2cdbcf01fd0381342731c4b9b197210c8cd483588f91b669b6819202603b3d810b21400986e5b26bb92126eb5154a629768446b2ac560f8eb4683cb58792a08c961c0507b18925402774b042ee1cb101eb84e567bc07bebfcf0246dec0223ad093c569badccaf68b4db7180eba8f02920644c666ac546d7c188646469a8b8f7257094340faa8510fee6ee83084ecabd920e5d9bc2d8dc01480d872a097b6746c98913ca98c08d593b810a1886130322eec858cd3482a3ec0eed81d4a0a148348e34b69271f4c9a956cf06a2db31266e3433831297c625fb605b348776cb016a1d0da4d4404806c1eb01cac09633f02735768e0d1955610f90611d35a0a6ed97016c92ca8ce8f075926d19460bacf7ea3c8a3025f3d9a43c113e37bd62ca708d61ec35ad3c9832c6d940706d5755999e3b02e104dbe2a724688acd20a410e943467ca3a8c6e1e041552949b614fe35a8126cb55301849d66f734d77e6d4893f3890cb261cd550d5ad6e1bb4bd61851374479551d0717c45a9b7c51ac7d47a854893e7286a3ecb04dd17ef79f5bf74b652c9114c120604846b4059fb6f245ac1223671527ce1e0a100d4700bda68cdc2a21d17ca064b040aa7a0f49089285ce89ad9d71d344eaa4b3aa9a0f4922bb0f5eaa900eaaba048ba20b0946b5738c42e730028823fd2df1e8b0466ae139463a07494ddc6bc38b4df5c4b24a73214afd0c3c11336b2385d401ddb2c05e87537914e726252f840bf6e0be6a04e3cb22428e80aeca8d4f10ca666490998b548fc6ad16fec881d2b21159ffe08ad068d46698b937bc964905006b423d50e8286606d794c3e33f137070962511025276321e42666da4ca7d8b60462f9194dedfc1cf293fc3155650ab3e09391c88daf5a1b12dab1418c85e8d24520c50a114001556c0174ad07474e224b45035c65a380041e3c611858016a22537fd4ec1b0228a0d89247143ac8286f5c089103624bb24b29572b1c1c5c9e4510298e57f9b9271664af134404c4f238a5d645417a3a1b7ead2ae668566295f5416c1e84912e340c3e573398eadd8916d4ddcdb54cb2f23857825031211f244bc130dd149853fbb5e45391b60a3ff73ea5ceba25c097c59ada3384a09a484972b62f021f58049804c15c989123624bb04b834304d804285723261a95a5c32894a6098aca849e2e20049c408631e090e896e79ed66ada8b054098feb40180329852e5e0e08b755b8029c8fd3baa53882ab2d28178c2708a0cc3080a9f92b0cf8489a28f01416481d1c07093a8a5826c02342b95932d4935a30d6e8d1e200455b911ff00f32b914dc8520c4830c9261dce68f1f6fd871be64661bf8ebfb093ab00c3616a2d841f38f82de1d0620069ac8a443096a206d0a473731ac5d4afaa0b6af524ab75a3a1b7ead2ae6601622982989e4754bae8892d4e47e9dd529c41559694c35c7c056382e9b521188374d4eb81456f53c63880d5cc9bd684a1876ca31e032fae444881b125d82580a5d3d3633ccb4c612014da7ff229d119292607f259206c6abbaa1c4bb6d37e88f76a28814c2ccb3ef604f0c7d3454d88e6597a17ee5349a81aa077a1f691443d8faf3638082e19b39364a401d77d50f7ac784881730bd233e7a82785675c9c031c2577cdba7de215ba8e8f09f0a4782c6417a0be883b7d25815c176b325d225f80a681031847d1a065200771247d3ebe84c60abb12256d9a337fec362a80af55b0f6e85a00000001005803006df7932b13c59216bb31dfc095087e9c3f3400058000d82c7d248c6443ad0d5a34811044f4701c168dbbfd18108921983988c7c0d0fbce4ba27055e8af7e77f4828a82f59e82cb531886c27cca55d90e821da7bead432b59a3be0e207dfedb4750baeb34ede08a8bb87ba2be98ad867a087e6888852edf8a6ad4848a251013512e3418d05ac1edfb1150374eaf66e9edfc560520627b4caa7ea35ef712964576acac728a8f1627ccd0c84d1ee61ef6867ab2640d49d7ab5d38dd5f42ee34085600003602c7fc0405602bd0bd26fe58ea95a68f801333171984fddd1fa65881be0ecec37a9b17c9f0ea5356ad5a7928cd92b29d165d8957b15aa2fce8509e94c46e5056d60f40ed489c2020716cde810ee201d2e93b150698dc53be0760a00b160c07fc33246513baa68d1a34bd263079e3ca3410b8c9f09906df4c58236027bd7e249ed2947104ca80ead9a9f3926f4961ea9458853f1903e6ab8b6e99b257da91c86e18b96649e20af2757b0ff92f039752f1e98a68d1a3490edacf58b34d811ffb0dee35ae7a00671fc542ed8e6b3c1cbf50d3df1ff98eb39d1f9349fe24c848fa4aa699af61673b2351688fed2445e82822c582c069f44db08b24419650413514df11e7b2a04ef862e888ff00c187becbe1e54e18bc8ac253568d1a3469735f9be094617ab1e8c2a1c7725eede919ecc7e22b2db5019e802ae649d27e5b7ad106becc7cca8f54d89e5f92aec430fc81dfe8028832801d56c5389aeadf653dc5582f3437a8f8eb2089cb8368d4f0d19648721d600157a50f99ca2b6fea2503570190051517a316d794a683848248ca905933292e7fc08a0da13fe357de03f747babf3893f0f7ad1274fca94f819e4ba4ebeb59519cf5aa49f38bbb5f302b929954b74fb24f4ad32352be827e9270b71f727b935e8b1ed289df53005d03dbda85450cd17a003f896d88c002e0bbe11c690d4d1f3350ed8517ace82c6b3917fd543927871e98fa47d13d98487b37a4d2f619c4d3ae439ffd50b5320e9a6a644cc0d0c104800eab60026abf425ae12c948e2836a69611594f7776982b8c01007d888e5f30a7d524ac8854bd604583825d6404e7ef15258d5be3500b8b2e4f032d76dfe4865ee52cc7d58c40f3b232b40a7eccb60b40db6dd5fcd017182803fe76a51c53ac988dda3004e4c8a46b5a61c1493f48ce96d676875abc8bcec03df389e38cfac24e242f10cf30b8aba4751bc023d73d346dcb958d4da35981e1e0824e311918c75b4f61f08cc68cd62b2a81b85e0a43a4e6d11f3e041decc4aef18f0d4b5d1b9059e459596f01333a79d4082b6651641be68f0d4d4c44162147265610714ee8e4f29de59bbfef47022d9a9e13644dab70208085083025183717024172271293a88b7efabcbab2023514d2d06180755aed58f0da0ac86b4bacc3b574a44384619b4fb4a192ccd11950c14275411e826e1c2021a811507b65e1714ab44b92a1f153be5a33312ab73ffb83088332c4034d338a654f5e1244500aae6b4b87ae1340511ab8787ee9bc835780fdda81b2fdc6191daa24774a8ba6cfe34155f2db65c378bc1670b4fdeca80fcd4aa86a1f8bcb0eb1af25347a29a4909461121a94403f218a5e4d2ae5a8b941461348def4864c549af790b546b308ef5af4806ee7905e094549d47c2136c3633c532b8732980d9f75552b145062c0bbae99461e1b66a91636bbba22f54b969a0e3a85b2e49e49106818c64c52c3a17508c1dcd409f95879a759385fce28b5ad35763c1058f7a1758c7fa6a89572134904873511968d2cd66e1d9658280c2a9e44d2ae367205d52a062b14b0c45ce6fa1282897072b709bb3ac1d3117418422167a697348eb82964699f5011b146fa9eced7a0ae79c49049bf351ba63a300cf0a58095062d985137dc1a993bbd38f4ec544c56ebf65020412432153d94cb2936085a81902a13f506c7abd9062d33604cd27b51c85b8fb950e1c0cad648d478dff12a162ad365b0f076a8abee9768fe2201b64564809f8783a91cf8627812e56053a939a300287242c9e207c07e03d81a3803f80c416fd89375bf49e8a6ee00458f56eac3cc896618dcccf34c2da970c916d0e11a7be4cce3684655b0537f64f202e5392bb9269683a5342e098346fb72291d70fbcfe2ce41400028e5710b9541d5e186fa0c1c1a7784410a02f400ce181431601078ae8643c07208a0113231e2ad586332ae77152e308b054af597b320011af92798ac51c071420866594c6bf77a941f1cf5f60355405c4f8cb306bfe55aabeeba62dc70692736295d435529f2d34d4b9882238beb5e6dfcd4afbc9e10ae8a4f06b727962cfbfdd4020b02564272976f023ed13ceda0516390c300770f71b9c64b908e867c50580fb2cf45a66e8f0dce247a66f295751786c1256f04468b58f71f9b315db259a7589b719527764466406a7b30c01da0817ea7f9c9c8284e17b40b312a9ba14c6c50a4a9db26ca112496c6a4080142017556c01536f5184d93f102032b132dcd1b4a56d9c63094cd1b7765f45299ad78184dcdbd02439e1612c084c2959c6336b008610822d65cae75da664f8100bab5b626e8a407dc6501f187a00947402b32a535dd3a7fa2d515ec0dd477e93ee531e78e2114d82c4268949326ecb06f1956c013c49c5dd3cf83af2d5a54b5218b6f84e97f46621041a12b6b15732414fbb385981a6d41ed15b03ee47751466b4024364bb310582f4b14054f02a3a0a8dfc004d5fcc395ce1b0ab4d5f218437952c8bd02618080200680781bfb5f496c92a1c4026995bca1dac176768913991b0084c05c924c55b8a30b5cd891602ca3b6606494b0f966915d7594a89459aebf87f5c99ee1cbea26cde9c4ac94f1273291cac0e46a8ea40c2b312c7bd8cc1b6ac97c8da80e636117de76ca8ecfb284c3cb68575e6d6a510c3a540264e604533599bc2512cdaefc5313ed50e81fc366b469470d0b6d6e243689164a2dd0d4ea485d4d02f94597ecff008b87d3162621505cfc7752bdfd482e358230127b396a96845fedec3100371a0b87f795cb5a46c188107b7d0721bb8357a05da95d20e73e9f5529dd2935095676ca034cf17a04fa87444a677ff4661ad9911851350be6cad9b7929495ba0777b0072362760451a58612751864b2cc366afb6c33744ea3af0128c24eed20b79465986cd3377d63b48c0b7069c4238b6b7498dcaa1bbf91201a251237654bce7a314dad4a8b0a5aa09e5bb05b10d3b61a0a9801182a7581859049591920b336854115e482b922454e5ba539258250265d14340ad313d84d20eabe756daef48c3f6fbffde9681d0bd12b928a46d42324d608484ae4cc19b971c3fc30a5d8f9674aec66ecb84b8ba54c13810e8cefb530a8ddcaea4a75fac03c8ab96e70ef12e0a19dee64f0090a62362024c21f1ec92ea22756e90481797db2a880a4d97bb3c01c2526d92d5304c7bd0164a1e4cca76552a0991640b81a512799912096bf9bda67340255fa06262db854ac60eadfaa4f2a88c0db1589b3501192032b558974904b8d01d98c1298fb21f33e06c28dae7e902995f1a2156ed92a43c9c16d090299541a34a574e23712c02f2c5454fd9cc5cf5cd80568327427ed40be6f6d27ed9ca237f24562c244413175d0162e67127b1942c975a5019821dd292c821eed7bdab3479d4d8eb310ebe11f554b0532798d3d7ff977ffda0008010203013f21ff00cfcfa35c0d70bf6906528471ff00196cc32e075714679e1b7ab2fb54cc41c14e62d43bd08c4d186039bd3933b85cf4cfa54ca7743beddeb9b84a07ab5b75dff828afaef5ad0c3441df3461f31dea7d4048971371fe0e5250777c0340a2ff00c5273eb54995a68d0c6dddad28393fc6acc8b72e7efe6a450ffc14f31718b9dff774de8500f1603ceed482d376d945df83de819ae6cf63f35112baa19ed7f69ab911d87e4fea9d90bdbd6a6167c3d4a588bf34073df59776442892c22c4f6a63d63eca151884ac1ed86575b3a2d7153f7752e252d80005961ce5da52b8a4c935c35c15c34ec7c50755219a68d1a3468d66347c3d6b17b51b7dfc0416a6bc5b347563323952bd5f433a0cfcbfd6aca5b7fb32fa453bfd5d3f3e04e58e18a7acf6bf37ed43c75eebf5dfd6b4c2a2097d1a9cb4efc1f868e8b3a44b5a2489070fe4c8e8deb59a63d18a166f41dc181d1cb69ddf4441b38ac6d9abb668d1a3468d183a0fe3dfefb00f70dbcd81a958b60340d028aa60deb6c7eda9ff003fdc927dae3a6de0827cc54646581b7e2d7464a2c1be87afb50c29495701af92f82f407ef3aa6572fb1060fa9860e2839a07fd1fd56f7b1fd51f98aa48767cfdf1c767b463b20eb2eb5148534545d8a95a1fa81b45132c24e8923fc346af29a2433da7f1574d54350e856a779c6cf5184e9475c1dad97f07b66a18dfd45fad8d3d5fa0865a4f12d68c52f5a590d35b07f01491045e9a8e4d9f93f5429e57ddff00b7b25d623bd333a9f3481600a9a60b1b0d124b460692640428084a9955cab95e68f9b09e65c2300ae57d5706600011351b8f8f2e1b2d424852f162262dd9771d207e49ad95bd87bcda262a9a51762b6c47db4a71b6f953dbf922af4a72f85c75c05703d2b89e94a6452d30abde4f10ba6ad35e8a2ebf767b603ef5237b5fd3fba6de120c60dc49ba2ceaa2da847026ca203d6a30c011208bab0c819170828be13994478d1d080c1508ed7a4d0177de0642204a79dc9422277b36424e4a364c5262866688d7851cd06be03502f11e9fed10ba3f443e43c01c2accc2a4e2f4afa573e031424714d5afaefdd5917dbfba9f30ff0035046c83d684dea4315785b221b86e2d56822b0a98258dab4936a86d95928d5b1201664de032d73bb57274d6beb446c11e600c6237d3351138cc15431784d70baa2689476a418e8369f716280ec5c323e6e371b309e108af67d47eab1047e349c72bd0a5ed838faa25f049628a3314d1a7443ee6f64f9bd78efd6d3ade1f97e3d69e52498935e835954ee12b6c836479815cb4511ae4303e7600bab60bb41cec1f73adc0db6821231369583be141321c8e640920a086b892ce4b9966e0333710cb053929210f660cf42e16ad46f59cc4fd94bc99f68fad151e008a334697b0668a17dc1f83da1bbf1bbd1a736505778fda93c4d218a5a4e86f5b7e6a42335c9a0d5f0fd099f2bd03ad725688bf585f281c2688cd35c15c352f2667e812700841c85c2217d2a354214b317d05df73f345daf3f669442727cbc9f54258ab2937f112cd348aecfdc5d31be90090fd18f5f72d4ddd0f4997bf834522871313fd50b6453a299032dc0b0f4ad22aba634f1339e0bc6020cadb6b3ad7250d270d266dee2e889a1baa4f5a990a98d97506b243016e59bd223150ca0c26c1f8f68fa99b540be6a025a59f089a7ee05e0d4f99f2f5afeec6a7216ed39e8e47519a71cd61c26570fb30e4a31c9e1ec70ee714274317e1ab51c3867afefaced4e071463c0d40400c621bacc333c78d660b392b69e6cd4287a20f728f150506cb6a0f3edd5a53949fa90a8009691cbe08a2ada5a1f6d9b14fa036a3e127dbfba7a597c138408211ec4c44e12160571ce4dee3d48ec5278b5dde5e3f2bf8500488d8b6f003b92eb89696392ab565d7668c50a68d478051468d6fa9b714b3db12ae1def7febe678fa9e4f8a36a8e72f045a94315036c3ed94f72bd6c3f55a51f44a6d72cfe5c7c746c3a44ab12ec7c7b9e79a84c3a9a9d4cfe29f04505052c59bcf9bf6ab72ebacf43cfe2b06c7996803c45a71e56eb8fa4136335e758a0d57c69552e7c351152bcb57edb9cd8cd7eb3f740ff9aa68cbf490c8d0d4f336d9c69a6d45ccfa3e1d9e1acbdd2dfefe79ac3ea1f3e3d668ce0e01f5a741b32a2992e91eb2fc560ef77e969ea57c55274c3f15331015d6c3ab82993b7797dddb7fa7910d03fe7a9f6b7fd78ea66a48796d4c6db7ed5a8a833eaa936b7fa7c36fa58a7eed3534998f6dea364ea503dd1fb97e5dab3383f2475ad32f087aca9265e65b5670768f548ab98fe724559b927e0ce96d8c62d3e5ab2f3f4e0959439df56ac78d9e2dfb9591fe59ad5cfdae90d67662ad98fa7fd57ee89ff152aa72fe3641f90f87c327bfdaf97eae82fcabd383f75266fe29a7a03a7bfea955973f680ac19a1dded534caa7df488ea56ba57f7ffafb4432341d99f6a4748dfc630cb4665db56ac6dfb55752afb9add14d707d242285c1f6ab622dbe9c91eea5271dbf878439befe7348c92fd858893952cb0d79ac055115c2984c6ccd4d343a8acac22c185d68ef202c59b032604c7a956a9959d12b1102341339e345bb0db58040c62833d68157b17a32c8c17c4b838e99a4ab10eea6e664e11389b52d044a065c2347b4d4731cc8c75dbbd4729b0a47ae3deaf62a2153b4849a79d1176e354e38669a9be82bd1a9dcda057a14f946354b8082ef19a5ff008a0e305c4e939d29605a82f0ec484bc144c43f19b1684b82cef61d36a6d32443990349b0c36a4a570815a2e2dd4c561d5a8a61001944dfab6abbe5011b027753317d4b3a4e278a6e1d2281c262278a705898a2629502e30cdb14d1b092efa349b597a3abaf38b38b1ec53e33c8213a8d32fb6e4ae1021a860b945011264200b9369a936eb82923688b031a29c498dcbc92b84f14a5d88f568eecf7ab8736eb2cc3734223b24549a1e741dec98e6906762112c12a6f6eb48e9781c22690cdf03b9d016f860a75681e697e258c414b0049908de2774331c560e4f00c6c55a7b60a1e8ebda695966415e8de919892a306eb647342c3c9d2de331cc454985d458cd8bda91c0342894e46733ad1e4f240e8dc48450c294134ae60c16ce295486cc88757077a9219c480f470f6a403d5458905c9b51260400c04ff00990e13bc5a918e583275ac88a92f04b6bdb3cd6564b1792337f47bd5ef92c8f2413e97adabc682d1e6a03da1cefb41d45f8a36ce9103aa0f4cb650941685ab5817448677b569695d51a06f05ed5bc22735e4160c43e5671f1807442463a964c1d7579fc59a410219166bd84896717d66f562b60dc065da63d2b290c9d5929663b05bb4ad4e5cbe06c903213ba2666d09c74a20b42492064eb36c545b344de096ce1120eb7ac10b991c1704ee02faff000fee8ff695479d6399e03985cdaaff00d8649245926ed691fca9a810d8f746c2646f35b99a151566eab60cdacb1577119841856be81e04fd4701940d6c903a0637681631cdcdcea891a2f59b47b508d06adfc06b41a764ea22e43503f7e4dc9d7a0ba52c8f2da7f3428492222550df17874334f1d2e0d667a8c6cd2ae8b7f8d000662c12b6a9e0251734c0a7adb9e439a6d8b4b2eeae55e6a46835187d69f4c4a3226033d4c6b1185a4575909713a91b22f179d03e8f86832073b4d1deea4ebd8081058bce77a77bc2119e219649ccfad1e69189a89025dbbadeaf7842052e2c5cd6805c3c5e2e4defed5047261a0900b09a21b5e84120552c1c5ad2f2273aadad272216185866522275022e8a4a2edc2d7b97af37dffcad9f531376ab17bd02e96533d26c39a7e975ce8c26671bd3a37e8de48cdbd5ed569c432bcb78fc5736a9dc178fc53fc0e70b82654bb69e941849bb3075593be1b38ab7fab11210678b779dab3c2cc9bc82045ef9bd3f08b18becce1ad7805c713281e657bc303685574de372766094e05ad6d8ace2a6ee3616bbcd6f602386f0c7a5ab33289d092867586fde129a8b99b8bbb035567a6a58bf720cb5481b0df337a41b4bd0bfba3ab3b55cf9fdb1de5da28d2a4e49d7648998c78cb3b1753dc2340ad3f06903efdc86b2f222505c57be8356a4acee8715fda94a8444021351418c4c53840e45b3c113843a1c2a646c19725944fa335749860cb9a58d5b53c9b69420e96190c097dca80cefc458e960b54398ba0762dad9262dcda998d565610e85aead4b9d6f6de8c734c646396f2a62317deaf5b325058de1a8791b3c7ca2de950af57d8eeb75b55e016baacff9422c0a304dd6cc62a610fc9eb56d52e8f0bc60ba59c4d24b892bdf79c41cc426c85a881cd6077894fa3a2628e7de212530e0b992e092c5135560a1cce8cec8bc5a354b83e1a8c90cef1cbb30adec0470da58f5b563153773b0bdce288b633fa02310ef595734a2d66cef5a0caade085ed6cf3597e50910b836caa3048045b0317bc1c88c6897c4948320ca6928003a1336821c833640436d223345665230009c5ed73de8fdb0986a8197f88e86200000c01a3c001cac82e60c0b731a6d5855f2183025b0d12e78786b682843d67315739754f3565197277ebfb4d5ac544a51b4a18a47934990c922c66fd6a4586669758cf7ab196d113d31ed57b1530a9de104d2a88bca066e592f7b6b4b90c1402936a2bd4ae740143d4c3de8a64ec4b75eb125749a82ee1c4e75cf5a9a99ca892190c32599d2de22cdf0884d9848714141081643694da81e0d221243097c5ba53e1242984d918250a8fa064b6d771c628ce1f48c743076f0011f0b408bddb1a9bd20c9c0475b39e734c14f2a55eab7a04782c50bb248ac012dd5dcb7756890f7112724a4de315c96d2f42d1e1240080d808051423b85f3754b7eed01cc0aca6163185b75dbc22c6f2a304109820b16b1e3c02a17e8f6a9a38c488743076f0b677c141d0d3b45316cca2bd5bd064105055116948cdc32def7d6af62a6153bc249a04782c500244b061d02052d77d558e59695596cc80f530f7a9258c48c743076a07834192cb00c66fd6a2c2f2b44c90992154873420e4896519894daedb9fe65aa6ad47d3858a287d804692d6471ede1cad56d3ef2687c0fd36e28a1f6006c2a065e9566cf9fbd81452687e9b86e5e0152a0259fa20b8ab1cf802e2a666153809696b69f7f3c76a2850bbe9c82c2528e23e01e0346b847f88f2d8e6adc5dabe5506b527053d0ff0082daf57b3450a14bc7d3db7f2ab26ff0102a9a56b2a1377b17a3411eed6b6541ff00173f4a2a54a149bc500498fa67403caae16745ad87d67ea8d27d947e86fd51a89e69c8c50183fe4aae9d3a5152a54abb2a89c68c47e92ccd40c52df480ea703e690938173d2ea9ba5a889ea7fc1856061a8f90f9c78050a952a50651bcd1bd5c953a25a07e94ab01ab4a93bc17afc8d1312344bdc8ae595e5947e8d02d8f00f629af10287a19d19a9108908ec950aab52faa2ace59101b611815b5b25d097fe0a15145e13de7efde9cc33cf151e8f528ff0069fba3fd47ee8d73dca38e8d64a0eab40d00c7d270178b1eadbdebd925ee58f555ab79cfbd6f6a40b6c841fc448dd9163a1f05c6c6a05880d9187dea556c26c03dff00eabbc7d2cc931c92bf7cd26a209f45e832755aff00f54163908885028582c6605d294105f434a3b2f06fd9a731575eabb5f22ee9d68424701a07d8a8392619c626036b97e6ac6626528de00c7de077a69f709741ccc382fe1f9e349f4ed85dd1967f915f99ecfe6cfd3deb077e47b06efe68d002b07d9a1c5e7bfdddc791bca5ec57a06b93307060e0a4f7b033d209da1d6ac3d2db6a7ac21e24efbc05c10d5de012d98678561d9279f452f27f908b1312045881e15e67b2a1905b8a1b4b76c1bf82485556774590980de092f04d2bf1d444bd8f04e8d9f1d6763490e57a41ab019aedcc05e5ad0e0cdc44f4a2918e9ddd0293d40e69d03e11c895a3ed720713b37a8e2025b5465bb3382f6c50fb291f23d6dd296b9652c4dc597103b4b6f0740f00255e0a1854d399d6407a439a5f6643d0fca8980137226158b6d273498c60197ceae02ed1e91694c75769e90e5a5b38083d0fca8f11e4430351f72fe05589304cef8acd6e1ccdab24bb67a43f346fefc902f38485793a562f3dfeee433abb51f566dc543e423dd40eeb57a86e62e9abab25d56983164c8922744a84eb7783cd98a784fa9e97925335114949c86546555cab75a522b8791c160c85e448c5281893fe05a8dd9b5fccf650f07db08407a8c54935daf53bf870d493b70e8f681f57582c0e594e965e2b0b233ba81ead6b14cdd0bf548b96992d9220a453c8b84ab9cdf3bf343921205969b49365da22f261c22f2465ec76af3edf50dd7aac18d4df5b1b35d9a281d040762a08d66e5132ddb24ea44def45ec1486c583b4c508528e2191d5e85bd301149c28b2a2f760184bb9b79b7f5c987bd62b8e5a122f84df241c143612a27828e833c0a82ccb13046f0c99bcc60a1e38cd9b3d4c270d4b0270eaa63b1ea568bd0f192ec16866970805f943d513ad794c958183a16a148907621176c5698ac5e7bfdcdbb40a70af967f55354ea8bcfac9e9348e16004a61cb92d190ae373726973264e7c2737ea4ce025e6a426117d7f90f01060e0d8001e0afccf6786bf49ee27990e959171c6f58f1508a79371b09e8af35bfa7bc01d8e8b402b22a3424f82a3430b2974559f7fd6651a9c18c817d1bbadb9f0a3cfb7d79b6fe14c69783a8eb923bfedb52aa597c5cbc2d11325d66b37b6b5e61fc57987f152a50435226c5aeb2d31e483ad9f0b526e0a75597c1e287179eff007140f0505b1a0d8f27c156c1224ea587aa3a53b4a4972add5ead5f6f224860dd18e21d7c185b792d0921735b8f02bb341faa78a88c0820e247a5bda3c169bc820c58020b184dc618a31370335c45d37488444332799ecf0b561ec20ed7a1f836f51b2895f5ab7f372c123e883468c5e3b6bded51514326dbbb123689f03a56058961c8e71cf851e7dbebcdb7fe1ad4f86f2dcf07e4edbd132a6e821bddd84ef87c2c0af2c7659d85b587622f56aeaa8dd94e9768668a365d04ecd4ef847c17fa262190b531689e60bd3c10e7e859bd167b538a9e066dcbcc87657c1f181ac53383362be369bc62f3dfede818095740ad205b7f2e5f6c6f3a59bf43d2fdf8a57af22cc36a26152c435125e2afed38e14bab12f3e087ccd490a2406276b6ebe0def5404802524275b6d62ae9516231d146d46613950f5be1872c25ec5e28e500de3a25350584017a4d565d366b1bf813404660a02011b981569dcefb85a86a45cbeb47400a00d03c030af11100c4091b904736ab71c8bacbafe6940c2748c08e1303895e74abc620c02eb3a2d5a482c25ba0d4a1fa056219422176f0b21b0477a653655b099306c0be76af63b01a86a3fd972a68960c8f497d7d7561b90fa9f92afd8402564c2433baa224e4d0d4351fec86a430e171e9ec7b9a5631bc1e5e68628a4601240939d7c111a3800c8e265c833a8398877849f7fcd291c794405845b80d016faec2b620864793bfdbcff867e0bc991e183c32dd3e8653352b83ea812a801486583936c53714a0711588d6f24a0a80d38246cc0226528b5ce42444015eec99d28cba816d765c132db627bd6a7c8901e409704d06da5da21d6166b91579d9248b39bc48d42698460483a3310f7a3ab764221318a349eac55eff00a0d894ba4925c4a9c71431ccf5e86697af41a1bf705a8848b6ecc15ea9254189a16ef958f5107812d69e49a66b49232ceb6cd6fd62cb6e33dc98d2882940c32a611244b3f6ff00eb633de7d9281d6833a0240162ec6f1e44ca022498fe16fe3065f425af3184d6eed4bb99c5beebfa5314fdf772fd68a60014dc5ef30b6b52ec77ce2f2c65ad0d89ac496290e10c80925caf6a641c122706ee110da6f417dcb003136d09bd42fa2858b1a40674191ad1c48e00b022451b85d2629f3712d57208d82595731490e53377496ab93401d091608c160c72c653598ea66679b1d508e91ec3ed9ea7f76a3f9def69772210653aad04d6246e08bc720e58bda6948cf37e0a3a82d068b1597c3d0c5c6d63dc9bd1a5a05236f1259aba559230939295c9d5a92b0fdb8bbf0c1326205ac89244411b42d992d980fb077340affa191ec88510847993551def77b153b9901e8954ccdef12faa5f7ffe5dffda0008010303013f21ff00cff25735727da4350ffc6cf296c58ae57a8188f076d434b51ca29b8c6ca7a0782f5f200fd15ef0268b74a3f13b8f23a540881fc217c0373472d1b741e851153535a02b428a25d4d23021ff008283df3a1fb6862edf7524cb5621a3ce3fbad40b79e3c1d6cbd43ca9294ca8a666e131e77f5ac820a76e4e57aca7b55cd432f3db73ca1e1bf307926a241dd1487693c2d5350a8f881b94526a6a686a68a1f4ac8c7dfce363b5c39df6a0b1a56afb1ad45d7e6f7f8acd12e6b470e9e1d08d6881e2dfd539db653ec6dfd50456e2df5f976a742ae0264ae6107d68196c17a93f8477fa42e6bae54d4d4d4d4d4dee5fef9c3271c37e81ef14a168657dd6943d26fcf4ad64a87f20428145161869b3f753e2eca85a5895bedebf0e9483e16fd35ad3bb06fc1595a7b1a1e7afd455a85d1a36eb86a0d1a77cedf7dcf57ab5fd53d22b10c8976d9dff0055738d3a5615c9352bd434fc8f2242767c66af5c49d37aaca9166b0c99a39d963a7f7f38a60cbf5a3f2054cd9013786ffdab439a1b7d18a0eb452434a3c0952a55c50346df2da9150e7eeddad2ee91f499ed5809ec067f3453c23a0b03cf1e16264bb0f81dc04a6b61b49438c090062b97a12d0177e077c835256830d5b5a92ad8180b70287b44385b3e1123bd4b7ca6065025ed15a29bdd8f5c77a937f10b3ee549ac2bc6cdd987f150cd434bd4aabf8cb52d72545ad13340d152b2b450e061f77cb8fc0f7a6951f721f69a86043d052d00cdc8f261f4b7ad3502c511d02eb53202d837d726008c2f703325bffe08bbdafde9a618b71f46002f68986d2cd2ad6cb6b87aa3bced49634d4e25b83ec15ed534a06cab6f4ce98fc574b8ef31ecbc0195056c1f6151991d28eb3f4a6a69cb38f18a5cfdd5c6a32fee0a126c7d06a771bd5dab59978179dec0f332648a880596b784509d9d10ef1ac305464023cdafed149aa1887e5b0a6b1e50040ac4f76609077278690ab2cb3d606aee3a829283e11d1f3eb4514bfeb51655f97fa54641c177d71484b7dcdfeacd4d4d58a3eec4f33dabfa072fe2b322f1452de450eea80799794d1d180f0f868fc51003caf9d7635a61c726b61a4190b8e337926911768cc82cc110e125e3a535488b64041189b6060e428dbf823bc2eaecaf6813d0a9681a014c9b6fad05aa68e28a1f114a7ee3ae039e3f744c37a0f3b1f8a4a74d9fd695a2429616427d2810701f5ae9a9e21dab115d1337a8f46897bc5b178a2e0f56839762809191a416a94b2ef275a309bd447944e855a7ea804b4b596a4a0a81cff027ee2d79db46ec20f52efa63b52ef0f55a3f3e135c2aeb8314a3f32de5e7729b8a21bc75914b6688a79ff7ab1e12497ad6fa7a3c3bfc54552be22f94da8b8105634eb12b4b417a2a54299f7ec5ba12d3b785dd75f7a8035fa8a04b8ac4526e28679f082daf806bf6f3022a3027bbd74afe8968dff2277fd61a082ed751dca71d6f2fcbd2691833cafd54ed12684b94738f2f674aba1e4a4771f1b792b94898d265c7837c9d62d968e1e860452750b4780b54a3c8cbc5f3d9a76dda50d1410418fa933fdf50ec2a01e0391abd2d37db1317715aa3bb4a98487bff00550660f01253ad1db5ecb6e343593c17e45633d42bf29fe3fb531071044fbcb5028d3962976e74eaf379a670d35762a5cf822a114502b0669967f65b3975f4de81a8ab7a9f53e56dff00aa4e71bab4c3e1a60a59f4d2a3dedbed8ff6abd3b34ef5e703a7d183fe9a3139550a3d1d7a2b3aef3d2a292a34782ef8dd4075cd5db6a81d057b45f4d409714be69fea9b4bf2a0183c18c62bb20147aef87daa812e2a3f9bf54bdc37ead43883e94898a15133dafebafad4ae6548560f7a711ae04510c7585a1953062f351fc3fbfa738628f7a6dfced5ca76fdf8cf3b57f8b1cb455afbbed654d459da79bd06e76fdfd3121beca256cf61512cf7a881257a3c7edfba67e3927a9461511ba8a33ea1460fa0568cf59fd561fd4e7e9fb29a8e18db42af59ddfaf194762b07fcb151cadf6ba85ec283529414e77febe9c1bdbfea95bdf3681c1fc5820d7f86518c3b7d5d7bf854e6eafd5470b78ff9035ad0dbf7faa0020c7da2812e29343bab48bab5ae1dfe94e38a68f7cfad7e1cfdfda032451abf1eed16af67efc678c5229c6eab8e777ebed64dc14d46367eeac4b6b6b47ff4c84c30d0a1f67dbfbfb4b3b7db4638f6f6de8b8cf7f1f45da9936dbce3e6a34c7d82cf6501731d0e2d59cb90db06243235c5476cb04f842561490c1539ea0f08c9104a09f429573e08850bca47590d1cb7703831ca26722734d35a880776d481302b9a0ca1a9ce0a40c40b1ab04b190c68bc51620c2e379485da6f1534731113d273daa60119021d728e62adbb265f1bc218a118037526044cfa1cd0857c283d4a8a37a21ead2ee0365032a5b1ce2a6131316e0b1b898c6b4d0a196d46f23072d5d0b43dacf07367614d29235de89bc379d4bcb1732deb2e5c0a64ae0b07353573820240ac59c35b778ab94ea46f227450573d5ea46639c50d845097a0331cd1610421624899d8e58be68004987d0eb17656ab247a23732cb8e5b500c24508f44b513be2c9ec120160b76cd28eb087225b18bc54376a177b3bccdc93aa8c8e9313e087ca39a30cca7b83d176ab38d180308259d4226e2349eb08d4f6ba27899a1b8f425412c03a02f4bdabd2334feeb140807a3624cea6098ca849324dc5ec10e1b6e601ccb438e98819da6350244eedca5a8024544eed5d88e464ea7e514515f0a07b96a18988022760253c5328d346a6d3313c4cd472ad15cc5d42f40657aaac631a6311a5307e421076465349de98023612e41733bd069162017a197b1488919894ea64ef4983c12e0cb662f487ab7572b1bbfce4310e84ea0c8c49c3d2b06165b4b070be38a552c812206c5bd5ed56fcec070cb1eb6abfca494b79ec54a50878db6aba2b734ee6112744bab0dd462c389162147519236bd6ac85d10d2fa92de8be42d15a1255cc8ce7f36933026a2465d1bbc12ce958227bf8b50729a4011412466219b6816abad85365c0e8c49deb12628d0c10c4d754f780a8caf6cc4da4c844e04719496890f3290b05c1a45f3458b506098df19124f4b5591d3109b02c6c56da73fc3ac0f1fa5ea96ea388d57895c537b432104a0c3b34a20f012ac21cdb2292215eb345045266d1868b758b4866ac1af12972e58355a7861425ca424e848ca6aca5d0a4b9ca75158e832181d22ff924a757a16b78737a6ca51a79b32716eb592224598d3cd3d51d21386588eb6a70a0028c2b16ce5ce56d472f464d262da0cefd6ad6621cea92912cac2e607c00920a16b038c8fd1280137083603006c54a50d124f46a2e130e913131813a4ce428409e00229345772da6d1aaa7433d5e28c6f0cec2144a7b8254aedb46d458599084d4410c623d1a6c8484742251b0c3a58a6a2048032026ce94a44c9ec98385b3cd487d80ed42b810cb7a3085401859996f7560c5a009baa8c2cb04849100447a2b802302b365bcecdbf9e4b8d4c5459a6ec96e29862ba88bdd4458a08df651d0b0becd11f62c1101619f47bd5da4a4070da7f359534d8d4ad3bf342f8f8c825b705eb419a27542c9074dd1b65b3357a500e41645e251e9075b617d8816844d9b5b16680c5c4e6db9174af6955926623f18c176924ef200bced325df1c1191bda5f3c5446832d3799dec5a62826c2c193227d6f5847146ae096349b27803451fc4583d42468041ace8df96118d8222cbb7b4c44946ff2076f503d08deacbc7eaced8b79a0900b1c681298bb3e30f2e9d3f62277516a738d6527b76a4ad425a12d83049a8d6af8f211abcd63d687bf115548e8052f6626a6a0ecdf19ce6329352a8fcbc60c4e131eac55a0bb8a0c42ecd8bd16b123291add11728db66a6a721d12ef0a6418178921e0130c9809bdb1728fcd2402726adec93deb49e1db3fdb9e28b4991060066736daad3266137365a0bc1d73fd024e9ad4c7d1f73b2fd2f562521d045faeb5384a6731cbc97ce68464eb63d2ae467d5e439571ad998a121481b596889298991dc5955a595d3b0c3b86d912175bd98104492d84032182414304f281141a91b86d379d057219e8f04e36863712a09b2b064c09f4bd446932d3699dacde269a5b4dd57399311de8d35131166fd536ad5faf2d2c8caf8e2ae28f0c2259416c09e94821d44256e64b595c33a8dae1e3221c0304c0555728653ed20e214a92faccb6a633eb224bb0ded7b3ed4be4341d095307f8fe2fd212aa2ae55c97c034f1d56cc5ca1613aa669572cdc632a2e5a8d9f0017c3a49865a98b0c1a20da2ae1ee46367c988abeec98467780934731906248610248a74a814980a1d271daaeee3280bae5ef56dd912f8da42c511f3822521808b2d7d2836496b2267c887a35c63c093a393b52561e9c96c9ec220618cdea7f967318c18e950649b2032422486e461bf8b646c8922e48c966f4b559941577405e9e6760921924646107ad256490047712434ba58940efd017e734d332a03b997bf8069be71743620bba0b1c508c46580da531c6285e08000740b5227db2accb4b163004a301006c0b060a62ddb408d010b4e62b84203ea0292b24a82bbab25a59bba42620040c5b629c6108101217b29d2f139b78495360132640965bb2ddf1e7930791d698087310bd5cbdfc2ec065465eafca6868be000762d4bb2cb588f9c11210484596b6956dd912f8da40c5227db2accb47ef52a03cac9a34058104e6c414184062053a393b3537b3310bd5cbde9e6722c4104aa5800e9525ec804c43010c803260a68accc066c4c00980bf07f25825a9a96a2a5934fa69e07ec02c4c1ef57b17df5f5f08f2a275fbcbaa2a2847d41fb10095056936ddb7f757e73f1e9fb9fbe8a8a8fa770eb4d1a3448fa28335770783908a8b60aa8ce0719f5a5a73b9bbf7e16a8a8a8fa630cd40218a6ad5a3b2a13f8b3417e1e60ef52b28719f3e66b4353e35ac41415dbbcff00c33e047d4f21c510cb929ab56ad1a64015c5b76dfdd4b29e3079eb51082c573ff8b19a68d1a347ea044a8680d66b7280cc5697ae00d8ae65ff0029c1aa9a3468d19d2b4a87f9c94853b4ad056b35a05fe93b08b826bde5daf8a8e29cbfe0a545a469a3468d1a347c08a8a86a2a750a031f40151734640dc67cafed572e9c3eea650399fc46bde75ef599a7802e08f5fda8d28ad4ae23300f7a31905d402eba84db0e91affc04654345b11a9322572d727cd7353bd4d28dbc33f4c7960e7f59a3567cd8f42fee54490e0fcb3ef4acabcbf89b9aaf0f5367cb4c9c909d1c564269e817ff00aa298fa4bfa38aae86bf1d07963aff00d58b564c328de036b97e693258fe8088cbcbcf1dfe2b32160db97cde923ca4afd88a4680166ec4a09816386af5d62033b4a49fbcd2479658b4189b9bc1e060f1d68002c0076fe419f77f8fa19bca7f2f34f1e5ff00ce6da490908eaa0ac2b31d889ef9a977103b281ecd431bb1f67b6ef10e41f0897d05b950da68b3e4263e9fb54c7ae29015791215edd2bddfe2a1b6706a22c46a36ce5f057764b3112248c5e24776ae3051d541e0b6efcab0ade23621ef4dbefcf68a08da1a5225d4a6fac3e28bb19da3ab06378576a07e290b88ea5688535307409daa2574bae9c2dc45d6d7c9696f9c4180e97eb510c1099b88371d49bc17f008eb2880375a6ea2d59bb0a3acb8a92ed443d5fc6b477a422624166e99878a27673601e7432b62f48e24e749e86f1d43b85103ac44deac7da95a3e5c32b41de1b7823131ccd11386f5da10ea88bd5dcfb15beb0f8a7e96724866c8c8b43d7ef34ac1e77b5ec9750a41e1c9e025a955693b57d081da8a19509b8d9acbc23a9a3dc87bd1e6267d63eb8ead1d00080300602a60aa24c3500858d592f6a49325ab10217b93a57bbfc53b3064764b8d63213d1d1d993b546bb1ee1f20d1656ffb9fd828af4213c04b59d5fd1383b16a0017e9a880c710c8dead917c27c520640284daef28b0deeb4321239c201ee77f0d93631a25cb3a0234bbb95dec991d50af75a96936ca238065468cc5a0196928bbb75ef7547089ca66e8f50d8a2bb5899012056b2e9246c445fa9caf61893a947fdaa929371b312795316c48b961de42395525d4284bb89e0c5a3534a19c8b973ab2b911a89647e9289ff0455c865437c077414a15cd12da5a4bd031a51019f22a4af2cb4ce60a60a4488c83313cfdd694723e0758a434ba38fb1eea0d45061369b91264b746a03765daa7c31eff005a3d6ca119f4f09bc9a47d7c0cfbbfc784d2bbed9e8ba903a7dd0fb93dd56eaffc83977a8b5c2f7bdd0e8b4812052365d89efb5321083274bcb9fbcc0d28da9108f5347ab8f1936797ede104ea785a541b21795b5e801a078c510029886cbb48b5f4f1d1a19e09330b22ecac001d286f01bd2ef90a197327a107cbf76db49d09547aeafc15b758e997a24d19182407054270a51225263b4faf8611e30aa16e48db46fe950da47dd7dc9eea97d00af26c1e0a44eb042ca5c09b5f19e52c282596f2171d04e6bddfe3c24747c4cfa7c3c0f10206c162a45cfac19fb8e8159bf97a07bd4900ace89877406f31af8284266e4329965c37e23c64d9e5fb7f08e9219617e291f0ef7da94c31f762d67707178c958bb8abf1dbb8dce45da486a8daaf8862d8c0eb62a23807ac10dee5db2278401051048234132b0b605b50c323ba444e511de86acc0b05d83620baa0f0358b0b804445ae1bb78b4fdbe93072d828f5999fc1c7cd4d6d1df5fd50d2b12396eba4588eed3a7274d721acee9f09384272646569c9deb8fd1ac631a67641643b62b4f41121c8171e8339b5677b696fc3d52a0eb2b02460dc01180bcae86d24a31f193cc4804701986485e6b5756f1165a764e692b22f0598a4a912733616c144c5eaeb31ec00e9a380695440825e590a7301c6b4893309ad822f215776e9282cade1d0da9a26a65384a6437f0bba75d1b510b74d5e9986525dacc1a3b57b9d9061344fe9914a8df0dd0fd653b8bf62af037fa97e3ab60655c86485c6ca9cb1713a04d13dcb32295147db885eafa8bf62a482daaf93da9e1254ed480a818bc782772b2a483c9c9b9d853039ede23da5ec353d44dcb7402cba8216d375325dee4841a86df6f1974ceafd1f3e1a51c8fd0138ab1f55ebb0b685c1230daf9a3278b64dd253a5a037a64651be64122508e0e36a6f0446164481d1b411ad3428b46836144588ba0e96acc9a08438061cc0a7c9db14cf4dbb9a2b3a5594e38380b4c2d160a1eb2c49ac264ed4aa5aeb1862732da7b19ab33c14bb06c3010bbc285efc3bdc6d59c962683c719d5b6db56f1530461a2b12c04401131a5216a87d400a4b60a996b6bb8af03a18ba5f035d49617d8868384eb114acc4d24902144411b9935fb7d3e54fb3e031574a195adbd918d7f8e873a41ead6063b177b5bde9d71cfe98f9ac87d691db2f23188b4cafa66a0ef61860a1988bc24bb17acfc1c522470830c385af42a5a80c1760d5c5e18b8a4630758a01bea1daa7b00284499011a9bc74a7159e78487087b2d90e5ac00270a848262ae0193517e6c20b321c06069498b4844a096f39e2584560a5e0dfc049b253d65951dc7d2b7c5e9bae3b22dc9321700b22f2c564c2193b2a55305c8bc541e40ed152f42e5a8cba560c7964e5844bb18b538e5188844e3762caae9f9d81c2c632346024fb7507d2f3184741a7a54ae3b448eff85158a25f8fbc8fccd489b9ea7c5abd6458f75f8a2d84e0bfe0f6acea3dbd0b7ff2efffda000c030100021103110000100000000000000000000000000000000000000000000000000000000000000000000000075700000000000001880000000000000000000000000001067d4b42d18027af72222380000000000000000000000050bb50b3c7d8593796b3c4a3f0000000000000000000000107a98196c0023400001a5f5d80000000000000000000000f2a00004e67bf800000013596a0000000000000000000001d37492aba6a04000011c99812b700000000000000000000c2d2117d04d7a001b860a9ce739200000000000000000000465c69a54c89c02a280000032b3f40000000000000000006f55ee0be746c7380000000275ab74000000000000000002f4f82a5899f9f00000000001e2a0a4000000000000000008f3cddbdc2ed34400000000104235400000000000000000068804928ceacea4000000007cf16e000000000000000022f34459ae2ad12ee800000004478900000000000000000e8906b716b81737a0300000000b4072400000000000000028400000093af3413e80000000d222e8000000000000000cfa00000010d5f644df000000020018a00000000000000002c00000014ff00f886a2000000075002e80000000000000020800000012f800208000000001f207780000000000000192a00000020800000000000000ceb19d0000000000000005d40000001450000000000000010c08c0000000000029ec49bda186027b8af188b22d2f923e7f38d1dca706400019380be7de416cf80170917bdaed5be0f852f70450b0400026d9e5fe32e0160981f4e5f5a5f0b31b77d288df2f5080001c8836fe25d8474001c8b64a90cadf580180c4ac113e0c000052800000018ba000000000022080000000000000000000002e0000000981800000000000d08000000000000000000018b800000008d640000014cb74000000000000000000000036d800000009cc1240050a93400000000000000000000001fa9b80000000f944d95072000000000000000000000000002322300000001946edc000000000000000000000000000003e314ca000414b3a000000114000000000000000000000002f6ae1c05252e800002763bc9800000000000000000000000264264051000000247a000898000000000000000000000000000000000000003400000058000000000000000000000000000000000000005a00000488000000000006400000000000000000000272007000007740000000000003200000000000000000002eb010f096e8fd3338095d1801c82000000000000000000049c747d08413038ab0811fe2e6fe2000000000000000001476f79b7006c0d262239500d27fc3200000000000000000c67edc19401d2b80622196791c09c5200000000000000002f27806567cc4ec2c1d4203e88af1a20000000000000000005400000de0000000008c325468b9cd000000000000000002b380b8b00000000003d00f24535c7c400000000000000002239b50000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000ffda0008010103013f10ff00cf22672c009420012ae2936433b789d259364db3e061e4f98225848818098be2aed2b25f7ec49fb359016621267685cd68c99818f47fe35f932a0155f09c34974117788e125ac06e246e50c32debc81124c1337a2ce0696708b5c8b96ccd4d6b2a4680a41901734a52a642ce0008c829d177a71e4401f3203446f04d316a2489256a4264ea734c4254b52d82f289d2b58856d56a265b13637abe92d0f42a31144444cdb34a218310d491022a03706135900c01636c26160a2cd85e284932030a30a4b7837c4b4055a8ba28af488a27f09c0b658b3378a3a520880c5996d94877abc36d847acfa2a683f46959956da4474d8f7ac8ae20f4195f6a7113d31b65715d8a821ee267ad40d8b43b7c004a86827fa85a2e39c28f59801d0a08c7c20e8c61366e7fc123bb8662717d5c5ae2a9f81142261020e1f383497c932e9c850103611755c122085819e6e9a009323db48492029af8cca1889956552b53c01d4c00def25c131ea1536ca1461123632dc1590b65f3624bc9050156d8a56fe22aa40ca474d158ab0a07eca7b10cc845f7299496b50f46ad554a3fb2f4a0e3aa5223533346494aa60a98c6c6f7a2d78c9d224c04819a488c66f91d6a719b436e2990608308acb1a01d4a081c2aa740ac090da494350388a5d430616305235b0546409176362f6f430c89dc5f440ee505f1ff009e934ac1ba662c95daa525c3d213a85ebfc548a2a8daf9e946ead968f001e9405cad3ca1c981ee691466e414cb1c84c10ae317ea3f7f2f29ab64815422956b068b4042a1334ea5953952d35bc260497302337a1527958e7436ff004c893ea97a31980216a8b6e6e97a2e90645236453bbbb53ff7fd54d3f66164ead24422ef4a4697312cd1e92142c028a7ad3a5899edd805c49da310dcc900488805d350d26d9a219d8e09485242e937ad29754dfa63178ec45b3a7c1eb58522350c1a2ec94ab1e059808834881ea48812c16cb52f2bc997858121305d550aa522166461944f3560141638c104823f9bf48cbcf51550e82b8d13d2e20da7e68c02f13b1bb4f5fb941de3cfa50fad0778a37d7276a9cac09964a1df587b29afdf265f29aa535b61dd68cc7c88920c518d769810f38092fa158959e0ba91ace30033e6cc251631988243e0600200b054bfd7f552ff005fd54bfd7f552ff5fd54bfd7f548b80ebfd280874b724812a122b824140004262616be4852a70a65a8ab854428971a24e826c36e0185bc1c927030e319662631270539d54064ca0d9ba2e848a151eb05802861cc8312800b0fa6822208908dc47226d43662b8490ac90b9822e4616c96c5be4b4874ac7f605ad41445d83ef322a187ae0435a6cb4411d8a0c8c9109f7d75515944bc794b93694234a41320f54970758abb398e7192fbabda9f48c42790319d2c2b2b85b1e9160290de9747752411f095c5f4fee945d21ba1fba44c71163bca2b525da1ea62823409de50d8994979d69468a85026c0288c08c5e288108236d1b436ae55a28551b90f142fd4def1040b5825d9a67eed0b867083a03966891e138046e937286c075557f9ca99d201ddad2c33a2cf67cd6440f03d49aae2f1cd38894d81d1d56a79a1bbc090276a03670d9d47a958d86342762680870573fe4ae7f3f335cfe7ad2af28304c091dc138a323a4108e60b06f07868438cb80dc4b7dde9e3ea71fdaa081b1094cad5c7969bac8b52acbb149452cb7c619e6b0dbfa285b085b2338c87d9125d42a7e265c42c420020289af788a38b35730610811979e539539011adeae0b1b2c7600a00800d8b141711b1819a0fa4264b80b8dda67ad4679106b860b97a41c0196ded426d33260b76448c5b3498b8319db122c2c86e294a1b2d1238031c908b3f8e44ac470253b044ad1d5d10426464dd689dae07e565c32ac19bab4b6f647c91818da151490d136d0ceabb0acb38d4f7ca8cea5050bc6376e5bfdf86a5c4c96b13d90a54b81923a44b2a335333fa877e0a45d5198652c886ea62ba906f5cd5cde7a55a0a066c02c7666e84dbc166e48cb9222709770ddfbb3e21f9017104048c2205099d3b8544de243b53332aad048c05a1600162a08b79ddad7d7603f057612473d5f42fd861f89ed48813120ee213bb8a7d15157648956019280042cec242082296509555a3c048ef6e412a76410b296d9cb33609c1242688338bd5c34628afd179c64022ad04fee8a3d17cd26975166ae308a2306bb40ae9a9a0739b628d43ba20b35b359e01094226517b04b572fbdc8cc3946509262a52ef5db789116cc512e7f2d9bc965d5d0e27c147421e9f44d164df85b0e97de8f3fea893942c0975dd6deebd1254c9442692774740758a0202c000c00687051baae28cd41d45e807eeafd2c20b62e106d3a197dea0285f08486c4962160695ec3a08d31a4c025c5f1511d8426083037611509326c4bc460ef34d3cae3f0970c2c1828a6c83c0b13a4a004c04504871e73b4a508a0d414abed33c689e90b5841d28873264e8d03091a9003b02d864c45e806594663338959f8caa480c87767da958b5b6e4ece0905cd2e6b0e6ab021440b3192ce445364e85b24152a596c9b0281f4801929194b11cd17f25d84b82042c2266f524c21030a86115cd30000202c060363e98fc65d03a86804b53bd8556f0db3a9dd4a3779ed5947e29898616f0eae09504b582ac03584c160310058a156c3daafa8302c93738a5d1f7329720c81a41b885f432da0614e201cb0585b16ee7528a416e2cc0941258919c549ded465f0909c14221a1b4d0175497d7142abc770b81ea660e2f40180002080580c0056893f4691fedd6cc0c76806d415251bb184ec7e68580e92dd48a8322b5ceee606452405656ea0566dc9a461e84f3a2410b831deb41b75feeadf42f8fa5253d2440ea6c983760dc225a3244238b5614ac4b98428cc117aebd59c2c5a39fac48026e406098c481b8a9cab4513980c8596e960a965076a44b4e6c272cadd80c1898690c13930f6ad86b3de331f01362575b14a1028136a37ef81c03ee10dec092e0568562db740a676b8e8784e86322589166a6a808e94b0848a48b2ee6a7cb648146ccc5961883104ca6cae53cb621108695a92c852da9c80ea6eb480010d800d811551508cb3c8e35bf7ad36cd71ef4bc4be730d9d6ceb16a212602840b02c00582a3fd39e946e8b7d2a72def8861036e83364c104c9144980111c25152526a2c74febbd03478344b5a9c0d2c4e0aacb05e756f16b0072d5fc9551902444c37bade69da92d1480bf22a5dbeac45a00a12615c83b0ab0d0b559098d6465000cb04b2cb48c1163018b4a0194f42542ac0dd2c165e0654e0cabba0196599002a2e0029ee8cfb3246c9bc8cd03ad460ec2ea09ae26daa5ee83edce68d0248588459c89c9a021122494b62824b659698a87ad4f8f02c82b092cf85ad6f45b716e481b3a99c13b21610df3b280a78b15143106a83752b521d523fa4f320906f48856b9df79abafd3b3c144ca26eb733af80016118cdada1727f945860cc5b2e7a357113980284040c95c1a648a341cce2800fc4bb0407401828b0a89bc3ce95157c5043316229012e52071041aa200a77991a4d03a2512d64bc0c8179541796838d15cc6218811f5039e684c5d60135286d5e80db8c56bda332b0508aca2082491d5500542809cb0b138c04b3d61b196595a206977ae9576284f0e15b2d09ec170b3defc9b40fdb48c5c9b72670924a172505eac1f5010523a6a4267571489c6031cc9142744d850698038c9996e0c24824519a78f964d8e100a00c970042146a240870cb2ead9ed97ac0abd6898b15e28b9bce04ec7874ad088df24408c8a5c046a92214ae110a53a66445c0c19a67453e68e78feffba59e6244d35ac4219bcda8c4811c8f409341258f597d617696a2530c316f1fb941bc0cc9e080741668de99f3ef41dc7e1a13ca06421632b005d00bd3cb2032e004640297450640888694b748ac44b0e4be37bd838c65a6c4c8b9fa60984a0009420002ed4be90d069412581235dda0359a024d416e3132ec0084a588c61c210b18080c195a54c0b1004dcde2d2f4578a698c6106097870340a9119042bb25a002169999865f6a023e2a6af03255602b3f5d75c16342e28c523415bf2d246465a2c19060452f0f59db94973308b4c78086040e1230004a086a6295089d5ca86804b429b0f9124026ea120c0a75a6800e5885c0a143c1bdf0af1160bc4d49898d5946d4ff922cb7e996265a96805a0d902546184394d6e08953b4094101de843622ea792d51b7445f423293f9b5654b569b1b6cb1c5a88c40016076b1409c14392c4faff00540e7bf9bf3442da1de0ac002ace20610af736820269d508e412b0401b71a8169fc0bc6532a10423cfa7d44ce400128400095695c0ca292172118efa0000a19d17b11b0b7c974505c6af00a008c20aeb808002b5e87c136374005d6ae65ab6382024e1712b016010a11c24b82f48b94c9fb6af6dabab0dc6152f0d8c4ad515ab62bca09168485b58a2081173c883d9b4db26ea7d11fcb06122191a98c9aac33849035ac04cb848d113822209b4b8141a6c8fac3302e25c2e21490a996420c1953233a8d46111e168d271bb528ccfbfe6a33d5427ae606676e359ce805226381689cc252fa06540a26e7b5a1502b4d58ac1bf1139f3b29001387d3a568fe21481caab014c1258e5f905c6ea08301a3dd5e9a1627e6c56f89b3863670405eeeab76fe0e822b6341a520028009a8fc9ae240d6be56f418a07643110611680583118b67ed43c433c15241895500ad3c03d3863a3342e011477290282e149022e9d11421c5ec331019084ba2c58800fa514a3374012dd56094b02d44b8b99709971199055366a81305754587a542f4d2de9624c8b537d512b59a48de3ca2c3248da282e671916602f60c0625c811801a64dd03368d8dc9601bd84b15a94dd84c90b10d4fdbef41e8824a2e49146d82900c2e5711be83115309c3014f9eca855d9a9cecc364109d4bd8235003fe0221e07000407d2b3e20e12c009550052c0b41155c1608e842d882d8470097115a01134bf70b531e16c685854cf90360b04aa051b220449143c5184ba8613c9202925d494341960a600fb45e71d949dd0c1d14d76e4aa683514df28a20807032b83ce891c88eaf1e44a28000000000800b001600fa42de3d7084c525324c9642524b2111bb37847aa76b6402042842753081d06912aa827a6587c3579d0548b53c048c2073bf799d8b4e4f2a95036c29d808558d2805a886b248b358f114e490a08b29037a1c40994e038e8086cca0834b235196fd1a7b92d4bcae7d624ac8058b102c7a30522046268470844fd35418436049424325b4e4b2832fbaa80b3e0a461c8c2115b384ed7189a6a0e2aca0c3c215d66555c32c29708b4058a9422e7c25172219412f2246d81488b8843b302c0003ed330a04c7d99a4bb4c4b3224f096488357896400ba00941b7de252d07e730b01cc167d174ce7c08a8200255c569e004d1807b21a6411e71076734aa8dd5a71988c468326800956572acb7fe31d91b370ed343771b89aab3e72cdcb8722d928e57ea88e31945494deea1b9648946ad4e1800ce80172243484d7385c854724d82c010788753f330b43914049905170ddd0a64ba690a22142d61a835696440b04c011f6872a6abaa88e25560a9955001b86590d586942684a5192a0d896831329a009929811190e9eda1f2abf44b22cbae983978802a8a4bcc45d0508a12846aa04100d0b1c1b9ebe0e039c9c7d9ae43181d09094b0255605a6652bca016582b96e53253e1b1209382a42e599e52d8c338f02000000802c1e0129b093740663ec9d02a58ca55930818a98037452a705b082442298ecc9919830fb4b27a6e2e9bda26c0b02e051925ec6e53e91c24dd0965389259e3618ec5f8900ba0892ce72eae04410c95c04001f45bcdb21489d34a2d85dd04bb46ce91363b484f6220205230c2442b4c25c31d3fb404c69382b2814e4b858c8261cec0be508304914e48b02b9af46114104c1036651657c22a44daf2157c42366e1a0235bcaae50200083904d15be810d8062c0dd29755fb0414bd2cc7f0dc5e01057a2118e96c05a4746ce80529a953b1ab0434d4c293f0583da1356de32136d31e4e94560598d0a815d2a6998453410c8efe50e8098800f529f089c408a1177329e8f1820ed0e323de565b60ae188123235542021bf3b8950169a39cea0072d2d392855075c2f9a720b454493e044d0dd2b1062d6923a43af2021166e92d41e824d3ecc9006211b5209211f9831e0f9782942257b3421563453128f178665cd03206cc4624c655265c9632b2aa953c5907ae1f215c696f7f8513fe5584a080150b34caec6eae248ec4c1c892a43f688634d091049c71a52d0a6ca866ecf1f2d0853b1f8b06954712c49c1e4900544a248852932732005800250e0b08e40f459bc09c747628a08d0739294907146146b2096c5a20d37eb5081c25e066a0c7b66e8b1b02e14afdcb86fc1c8a0094c5483a187bc8ec9bf0c342c7532400833006a42922ab66e46db451ad4be0cf2c70f444d78e12454d2e2c56a76a03232967468e80b4137370ce181e42598948861a3f22347a0200160b14455c7c1d004196554895efd94b64819588aa0c50b0228ecaaef2edba7464ecc540421a6fb8b99f0b2a55750bd0d89292b8bcb659b0c35b6cdddbec887c12b51c423d5146524817b0228ff0064775bc698334a82b33e40231ccd2c8654808ecca199f0b186a28ce7745aa187c39dd02654b60357f14b10de8299347246b90776a7ee4321b63023acaa5a53247b5fff0073148aaff332663b5e08b31aac13511dd9145ba2e1e154c04e03c496c774027354c85593f1656eb1d62b52d1f8444890281a11549a3d365b11ae72067e345902004691a00b266685531f9fdd7145dd912abde000063d572139a23c198a84de476495c8cb1a98c1b235d92b3f55ce90e9b6a0962b2d4b9f7d2c58e43e565b12e9112d55e6d8a92a6182e50c002fe4a01eac89a93f950ba572a80587a71af4702ec823784ab2dc1a02b0e66a34177fbea067881027f01b104d8cb4716598c01218cc016140b299d40aac5d148ef4084843085081a7d07467286400ba1490d76054cbd1d0d0516214a5a23a4a9cae5d3397cd8e305a2690aac59f82e306fa213aa26ccafd89cf4cd4dc3ca969d7324105f335a9a1af8d4a6b9266031529387fa6da9801136a9c51f34767410160a5fbc1739f599d6dbe0b561deaa7df2d100a8b17d01ac2ed7f8a346a1b305c4c120156032b628de49ccce6cfa12bad44e795f8135dd84a0a0109dd8860efcd2e2908431fd1372e7a20d4062d335bde4d9d3683287b246e1ade54812c39c74313d2ed206814309ab5512e2cfd518c1dee0a445b5d24d620d5d9ab8bc2328822d4c732a589cfd40ae0119ed663704fa7b50f0382bf28aeb00815664a90f64ef95d54584a1e41dc11f7cd9b5a569fe7810c2527ace3819c8b66dd426dd851fcc809a52d66e866a31f2d20b0a041b4581a3a0dcccd1fd6fe151db5b474342f0e253c4c144b12aad526ee18716b0896d68829431a82481885c414a07a480a8a148aad025ac18c57b1f22890f83c0eec0948a4e510578040b80bb48f26b42a8b8c901a70d54c092a07b4222c726aa0546b1750c06aa6bedbc9019d59d85aad2e4cc8090b78a8b8dd20aa3a73613aa83aa0a2b099cb450a65ac042bbfcba18dac5f109535f785408fc264c0e5f0693f7c0bf093f6c390125ac5e54a8b36fdb16cf7465e6cf182d11481228b305024173500e2685af7ac8c8eb37c57e3b4b1c55af1b41ca859a3cecd0b505b6068a20a87e59c66700c48184ad6097a66f920b004fa91504c204ca48eb386f1ca375c6ab405ea57ea0cad7ad82024669e40efefcee6da42369082d4b29f162b1585157be9eec2cbd4866d984ba2ecef848ac2e739488cf650947a05f39893dbc4132eaaeb51c7434db8e573cd29d4f7e8e3e3dd32d0f3d2c42e84f2214d25e7186b4ecc98cc89aa1b0f48e6d81106a55c0f80c4254e64acf2e0bb8d58ff00b84424aabe35b4aa7c2c7a5106b1750c06aaebeb2a07b4222c7268804128d84aaa124ec0f50b043444d048625c4ab15142e84516e8b07805324ae9dcb740e72409390675e19e4d2a6558f106ad01a21a727892d1d33bd039ae8523c4503ad8b310c99dbb0c2492c77eb92004c1284fe2690bfa4de1ccea097c36810411d13dba9c0181a49eb7aeb1478c516d294c21f9800d220356b30fe70200330b54e20070a640acb5bc6c0001421dc12ceab019418510d71d30bee9c82c053dee9d2379816d0e1fd252d90c9374054b53e9caddae1182071848a8aa0eb85f34e4960a9037f70c7deb41c9640e746d51928800119a3def1bd077286a682d42cdbd2c076c6b96016a5144b19b55e9095798d51cdea6d313c31b5d8a8b100e8aa08a10649e0f5fb5db157f151f0222b5b4c39f4e864001db68d3379016d061c38afadd5e1a81069442758d638eb9ddbc44ba98064342141c2053c36bd3e1fd08da72028600e1dd0c2b02c202808d66235dc8766f85629b497e1178200722ab405ebeb1be38a040c0e5431f085fbbef0c541ac8853c2c3038001a78715f5babc750a8d4dac5a7c5840c5ce2cbc3c9b8c07b52311eb870110c33120c09ce68ad106eab5d4400a3e2e9a15450e2c6b9085146b04fa1081624822cf8413b229619d811d0425a9e24e15942c36592aed239d1b5464a20a156680dfdc31f7ad072082aa0eb85f34e4968886d25f845e0801c8aad09b4aa6bf1514e026d4f84c1c4f28719d296afb397d42cb2905c652a0575f4a09684198cd76ba748de605b418de1d37fd8400508c4f4ac0e7600595d04fe38e7a297686e41d895b0d5eacc455a61b1125d89456ed21cd4207924979dec9dbe5fd38999c3751bccecb8c9a56e1dca20d2cc2d7405fab47480f040108180fae372cd986c1620b750a2e9a22acaefc493414681e0a81b60ce304229946333a11490589df26e1bc0a05979fba58bb60baba531383cb0ec830accf10c2b28a68286c185194184ee5fe90184028e00257b054a62072a88de8107056c515cd6dc4c0c61898da8a00758181c2bebde7a4141c3ec2094e0ac82e205e3228d85cb1244192f4f44c1ef284230c40492281338000401000200fbc5907b14182ed20a791998a5697a1daf49289120964982194a19a0cd80007007d240213360e3d9ae0fd570f72bbb590a2d88cad7374ca224c43e8256b04b252f72ca442eda95b18a25d014071742ea6810055002a25065848f0055bacc8a40b5a967405a464adc7622827590d6aa2240632d0aebd71426108ddaa1d5c1e81f7ee12bb057e063d1029a3bcfdfee96a00d6b92f82c7d3ea13c5ddae9457401be524939591d4ae0f4fd5707a67cb401e060da4500384bd08f00041d1559ca5061689ee35b51283a8e684411110446446e22647f8723453d7fbd0772a32c1653d0eb0b68f37ab0dea5e25654361512c5a941798415dd09d52a35d7520896340b4b16dbcb0a6cf64560092d3435a119c09560600d83fe070985605650c1573257996af56834b56e14ecf3f9ab0d014cbc004a0f07d30ec043818817b0d96744780b36626f13839091a7dbcf5ae2f5ae2f5f3ad795aba3139b8dfde512eb20a088d9f525a60a80dc363b165dea6d3283321b4dba8d91433457521562049316bb78a63523bac5b83274cd65093b44126800000000005800b01ff0d0b4806e802fbad766b87bd26941a91ed4c90926b2301b44b69836744113d7786d2e894ba04d7e9b76bcfa54129459b728bd987b144d864c4dd73d45444a989626ce1497c480191c014c4c4452d30255b9438c3711177a68032b7d223caba5c69504ae6172ecb755dff936ce6ed594c2e1d6e46bb35c3dcaeed6c54d04c68d349d6db39344687c356f2ee01745a349eb025e935d8a100a11b88889c259fe202805d5001bab62a56299f2bae953e3c4907b9b2a1c8a521fe934517ab2c209aeb43340b379d5666520620d585291090c0109b8a26800000000010005803e8ec5404ab69c89749a626ad245361bf35f8c5eaef6b24e066c09d1ff008364d298a64e4ccc37db25073130c4811f8031adeb62b87b95ddaf2b5c31e7d29369a4ee54cad366659321e9581a040421c20636a05107a9f41a00438def5a65dea7edd249272d26693045c49eadd8caab9dd502bd023ce13bd4d441bbac9b8978e4aa1e0a5888d90f1a7d0094f925f7501d68ef34f6209536fc06875865075924ba59e288e93a5cd2c20c4945a9b3a85263b24d10dadc0d058bda8ea4c02252220116788da95554775191422208d1a4c8885ba27e7131490dfd890e0b16494ffc15b366804ee4dc372f4d2bdd47601077540985bfea554027b089e91adcc52f7505f15464cf62e63d9fcc5728f9dfdea4ac9cf7a531c73378b74b7a834763d4ddee5502bf84d66a85bcacfd150256032b8a9ae11828b8b311a82e156259d1c3c444d2defad49aa480706bff002ea2d8e46d3b10c6dfc555d27cb4505f26956a2c687f2a609657945daac90553434e7feaaf8a5cb839dced41162c16034fa2ad97da81360fd8d7356d3c0a846b150443308aff00aa3db161d072894951030d634fdc261040175fa139a06b8c6adf25d51a4c0300d02cc554498e340ef051fd51606659602eb2add57ec5b7da7c0548b617491dc3372a7d8724b0027eec1d94869d79a020c02caaec113dccc388492e16ff0056c8f611604d27b3ddd5c05ea0a40294f5c5e0eeacc4a50cc0c50da01bb75caddbff00cee89ad324d94d0683c195b53cab3125cb0b0440020008a6633cd2236493760150a427244ac524c7baf10adc99c25062847b4917b93419db2a957880376b0f03743d972702210c2a1f3617a4735433458b7447872b970405a90552915c456f7f1746ee9306af874817a5a40b1dc867a66f14d49d6f5cb77bb562055281e2264c4de5006044296bcba667e9e0e2ab839d08f67c0f1308c88221e4522cadf25e7081ff000618b3048dac4ab05201680d8230a03a089cd42bb8b4c0e3c374c01aa662453704b012a1525ac1f61cd3e4597486834d1938b00818658d2c429340447fd404886a543b923db7757dd027e69e22749665c1725a2385959a4414e075ca520aa6e5886336f32f0dbc4d9117b8eb0b1ad5767210b35603367642da176bcae9736307eec49b176c1968af83c0629196cc054f8739e1df90bbd40d163909e032322eb719064121634303b952e31fc0e7aacde4376ed3c319132f5243540b3294ec4fa42789341515a78a1a25891a901dcc820bc0887707d7ca9250f0078f5a7a425b570e1a921d818298116e058b5421ae22e943c06df56300c1df52200d699281d90038164277ab9c44284c0cb1d776924c84416bd13e44e15f693dc0445c04276a5bfbc3be6edb215ab4d7c9f2c132ce520c82f682af056579c07e0bb6a3784f829c3f5bab42d4a8e8780d04ff00a5dec223843316d8b0245929781a58b048608d4010da00a39aa0ae412a51010499181184e604526adc83595036be7c55245cdcf4456ff3a1285514e9d03c313a13eb42984a9465221d27d02937063361b6bd32944503e52da0000ef0249280cdc1210520036c6e26f378631ac98b3d5ccd20e8921683866d097dd53059195401f2ae8175a5830dde4cf31818b2dfe1db1dc957e54cca666e29c8e2e620a414126f2269e9eee52b68d6322011127c2e1776f2d3a460c2c5dd0bc0a0a0650d47500f09442dd70115860e9fc2c8f8c950d024272a96c5dda27d6549aa179a73c2a7663b2489e680fa3157e21324818c189b328eb4f594f3e0082f8255029e487cca5613cd5acb161604a8269117c390252574ba7f3ff090b7b2ff00c3d85e26c0680c69b1660771f0202dfda03cb00a927004a3643e23a4c2a0580f17dbe4c17d9e1c1623e3366984bba77750ca22130f898bcd19a06624c02d705523da66718d926ab3127f75d130082955d5e59b602eb4a51db2e94b9481bb70985dbb1a7400229b6810cd84d34de488ee5d42baad13385354783d4eb1c07814639d1326ee5ce24a1375212a41901558015aab37d541119415819b3592809f655290056e9029e9c6e12a8c2840b83c078cddc1ab2a2c016b309adc1444612e2591308d372c7a5fcd8936080d2d51ec9aa840e9a36113132b891125d720b0579d1c53ac6969f22fc3c3865a658e1b835039318a63f97b0bc7085a43ae5adf45946914085806871341e42e6417d4022af03a55b05004608eaf07ccce068db086814c01e911262897514d19cdfa7996f0e085e6e9e02a97687a47225f0a2a7b8492141892634ddbe8c8782a3659b6fa61411d0fdc29839e684e656095a7125a6c0e43b0b1712179503d96059e0ae112dd8b3055c2499777f3cb6c399b241400188dc4161a6b464859c32b62eb19552aab3e12673b1a8932ece5364c95d01f538a1420e226177054016c172258d28f490b83e5848da110ce913706a5715425406509390d52ebd9c2cd1097055d94f183b9af822c218f0f84b8cccb2c098fcb6b40a0712299285e8da63c203500c00583c016521fda20762021b52aacfc865827486cbd4e4422a5345b4f2aa82e88467209604dd2628bfc7f014e230140283228dce0f1dcc4100124826bc84f48efebd0b946fe9dd5010069441084aeb36a77529327a69ee6c427e339036ac9628cc7a229268317cbc6130327b30da8a848a9149c2ae9e384b2d3c2a876be50ed0f8592bac62c6157b6d235a2e74d46f4d25d841aaadf626d3c00c34bb358ef15da6a091431a9ca2ba4dbb81482548217d76821785280b7f2f90b17d9085287db65a6185a6d083de6e22cd628f7d4495c90c1c296481f4230237062cdf47c0d187cacdb41c4f2c9b504195e5cbe907d41a48a8b71fb23826491aeea0c279e7aa53141b1738a87bf54fa44a0cb4d2765469ac6b0f791c1c9a3c5841ba5c8237139dd806a60401c322b51388067a2abdb4723805cc4209ad1cd7a6fa2e1a32c48cd4f85cc0b747e044b15a2f6c226c4d98bb8a108ea262a92c11eb825966d11847248b9f3aa19f45564956180050115a57597c2dc2415226598554173cb5aa7dfc00b3dbb073632c6118218311960d9105fb694c9054964c20c839681b11efa378b5bcd8500a592a540804002088892225913c5405580bab80dda550ae7f34d69aab3c68402e463a915780f54f2653ea5d1c829c929dd3c825fac471b9a70ba3e239c42bac31b485fef126253df5d9645957f86a4cdde16c20bc924e54b04c7aed36d7c22f0a9a63a159019f4c44d4d61da020bb80ec24103466446a462f9f52b3eae8451d834219cc56d320170d898065a98039d354e29e9608d0b0ce8f26c166d5008011a8b8f76748e1188d6e55d1f4b4ba86b37850d2f645b7a2b3fb24573ed928b8c4d25140d13f11386233b57b801bbc3258f6057cbe9fb7b9e0388e2d971041128a33df46ada024486c01e188322b816d9b6a2b46c799ba2ce88ad10852e5859991bd0be962c6f325e5132378b42edac43a4619630860881050b0ca0167b29dc7ff2e7ffda0008010203013f10ff00cf9891746bfcaa473e93488c243f66a200f52acc87a3ff001960b90fa8606f132981a2f959327abd400d228e3c4f7595f5acd5eab572feeac9c74528641b127e13ef5621cf3beac3abb54684ff0089d4f805ec2d2418c897540f7a9e98ca45446484cf1133689b524011c84cb7bd6938984b4c8435e51360e75c6b61789b55c5e817bad765d661df02877689081202444b88a263f862b3d6b06f40fdc50191eafe01a6e01d15f9fc558d47407e26b364eaf898ad4ceaafcf83aa66afdd0cbf359bd889f73e4a60ff0062030ed412b7a8c9fef19ff830a8a8363a2e53452396242b69a4c0ba6aef22b2cd1c44d03ddd8e5b5023fe87ab70ea2e9440ead80ec013d55240bf2ce71103d46a693e2cbbf1f37a40e906d9e97a5ec3e6b95c395d048f66935e2423086a2591d929e63d3bd0124036748028377721818446423646e54d0918589bc4cfed8d8de54e5e791938602cd06423b19068c5a365229d6da126d868464acf03b56b6fac52fae97c7b8fdd2da3d4fdd7fabfb56593a13f1348c047922b70f0777f829298b534361afc9a2563116c92f691d1fc8fdf90f2dc16677cd93225694564a65cbf803b01b1510e6a6c707ef3c14bc4198c87bb0f1b1205e9a527441f99e50e294573e6135d15cba365ea251c83f01eeb3b1a256012fba49f58a0fa8c269fb3abd1ab4b06426f1844ff004a6122580881848b81bc179952062c4de9db34219089246390002e152ddd2f199a51cc4d479e276a21c846f0122a06850063f98308f2bd1923a58f4c7a450f6729a7534f8e694a1fe0a868dd8bafe10ed26bf7d694804c65692e2c45b43444261e00000002c02c01602a7404f2fe0e3e687ac3b8cdf24d92c3a4297869465cd4f6a9ed53daa7b54f6a9ed456625edd45fbc76362c908424d387729b7d9a44f64a0b6028b3ca03401d806591518b3a6d4b77113189381a42fbe32b8d027aa4050540d0e508f54a69a0163ea16aee483b9f92be0123f15fa3a515d07c9b2acf23640fa31448ba638e8fbe8290b78023f4dc4032a71a834ad02b0816ba47ee682c7484b9248572225ae62a5e9e3377783914a8edef5d3ef5cb151609aea9504d81959b5d676a9de19511912444dc8ed14a59a094261edfd51437e9953866d0f9428a9b839bc5423166cca7254533028750ec2f7206ea9fcf0a5631e4e6b44bacbfaac644e00fefdeb38faaf842a0b64efe0f057056c76e50b624c378cd18e024272167ac3a34580522323f76424ddf72159c187f0a782f6e4eaac1eb0537d22177757bb7a415a2a04c1075cf64a14d24b8974aa50942ca955bb538b44150b9b88e244ad632a6760200ea22270d703c33685c8920005e4880814cf49f3b1c6912531b742b1500d08173a79f702faafec7d68390016f02f76facbaff105c528cd3b3c10de5adf97af857f82fd52b9f45faaf8316f88a706a4ea7bdfde8241ca799fc570f8326cfd5a7efc366cb33c012f641ee1bbf769a47bb843b2bed26b46f364f7053e14a10e42955a530280d90ac5a8bf152472a3a507a87385e541c6691821b3e80d05ea04a65db2e50919b66c0420000026b71009372c8736b960d4258b11284411c8d832ee1698726385e378e1a232286b5aa6c72bd493d8fa5645305f0e0ee8f07a0abc52fd2197d0ad015cfea91b071f48026c53b29c1c6b347fbf9a1be5b9af43c0610cfeacfb7dd63673b39fc0cbc17a372c9a700cb6967f14b8b07e53477269156c2ba85311c5b6ab8944a6cb0cc149ce492e6136255502d2ddd69ec59113105902105026441fa64a0050204321a08aa2820612aec919dd0cdd232d9014813510e3c57c2284d86b6a2282da9ffe2324b98c92c0675c43b444e393088c80049a006894c076853eb472ee8ec4fa0d381ddbf065ef46c136441edf9fab20ef4eca221ad323bafeb9f7f17bb7860eaff53ebf7336052fa81d5e5a682eb4148b53489567dc732fc20c98b98cdec9c626af4298d8afc14950ef4fd8c3519e20baa1a5f5248b2ad04577ddf29ca96022b025005a725e920ee7cc49083541095075549fae5b010128918a82fd0c2cdd99c8063a17543468d89a41c44aee1d78462c08c97a017553150b320ccc7759e668c09072e7fb1fad24e29ab2a336cd5d783c19804aa059b5e5d7ee1da46c1be925c60bdc02929089942205665325764d2a37c489d3bb97a16d9a9f82ae8c89ea69c420423b4334ece5cea849d99a9bd98d0437ed4260166538205229344500862c22f2bb22ca9a82e94c1b5972cbbf2ef596680cd08d9304dfc128e320772d1cba844450c226446c9a6b4464439fdd3f0b648d4218603698c828406ed2204c91d16f656a501ee64f4177faad12ac418a0cf8a717801e99c7dc454399f835f8e6b28755aab86912452615d713923cecb8638f0631b2bed369ed9ed5a4553652fd8c8f2509c918e8ff0074c0d6830422ec5940a59629f63c7f157d830d0280d00ca18ac205d2ba9c7a2ac68152df45180e916b69e6d0a78c437a6c565112caa52ab2ab75cd7ba237cd4fdb0121002c440b85191ad4cc27aebda836c3fda7ca7b6083d991bcfd40542eb5cf99ad80a772e3c043c00ccbf6f37cbaec7574a842e4c4f06575f42a465dcd5faedeb4388589176013a318c20a235770a25690445826cc8040342a8d82dfac256c75185eb9c8c47aa944cad44010b41146fa351d5403397ed5120c69d3c59c109b82c4066b8511712cd9908e4f9a6571eafcd5bd8a231c7d93c49c54b6438a3589416c65695e072feaa384d0c2cd98e5906c0e0695c962eab2fd3cd6e47dbfbdeb523a1ab5edc36f08029020c500fdb00a8256a0be64f5dbe7a548706d87ede56a76397f1b76f02b26442040203012121569432fb1209c22a86812d0585caf08c70293d24eb4a8b0660ea619a4f161bc421d380d9201260b0a4961264d308d992e4e752191d4475acd12db7da3c3e6e34b94426744dc751ab95a8e2b618a9504784669540bb48ef061c1f93bfa6f08dee4aec6bd7e5a76c2566c893b926faac817d30560cd200f81e59f8eb50e6eb07e5e3e69eb4af36f00e12ab5d1d5a61efbf5c7db428c6a583f6f0549eea65e069dbbb53526d6af57f05bafd1663688d92e1386b23ad66e349c907d47ad494bb29bee3b4dbd421680a0d8f6f758968d15cd3fc146a6bab074355b015a7dc0e380d4b483bfaa3703b0a556ebaa6af920ad859463b130bdb76f3e9d1012b05420e749a7ecfc7bd4dc1aae9fb3c69aed4e1d5e57c0c04ba4cc5bfd071f34ad94f73d78e3d7ed41402aa5631795df83d74a89e248031d51f05f78a990f83a1a7d2060ae8531eb640c8e97303ace75625608b8b20ec64b5011abea4645bfbb1696aa6c8516c0f5c338277caebcd975621fe43df8adcd8272e987aa62acff00e9c649cad1bb536b9b808bc7769c0af0bec4a838676052b2972809b900152482d7c9dc0c26d93292d55565bafd2346ae854dc18ddd06b1f97e0a59b46bd4fe0739da3c4494af4395d0a6d74cbaaec36ff5ab9ec36fc9ddf8d3ed60b5b5743abe5a93e185d69d363cad2cd2356a7a6dd73b4567e9372722fc6efb6ed4699463e4743db626a6bd8c060f3be6a78b598847633c642816b5c68eedb1ab9dc02d5aa16225eb69d975a9fb66bd71008e668094b0323daea5a04ddf481ab7a0629d3d31d5e20778a2d81e05377ca174609b52a4c56dc18194d17e9a8756d5d0ebfacb536eecac9b07e0efbd211e0ddcbbf4c1efe305b06560fdbc545a90c6abbad0d8f435a9973ec380d3ed6366f75e9b1cfa53a3fbefcabeb4f46ce0df97f583dfe902b066a680b69e4f6f5daada136e3f09e306bb53d65795fe23c15ecfb1b1e8783c9b007377d2f2696fab1737bdd1c73e9350f04e8d395bfbbc66afadec380d3c47bd9b653f073e9bd41c0044307e5d5f7cd20455abf680015602a10474e83aeef18eb530b81abfd6ed4abb686879dfe91919747d8ddd0e9fbf8ac8fc3f87ede9bfda0315742a61a17870ebbbedd6a465f54f4dbae7a52ab2e7c0231743cda81a9a397a1abcb63de9ff004ed7abafc7da93195e6f41cae65db83ccb528b960fcbb14f1e57b7071f4b45daba1e76ad46eab97fae2966fcddf838f9e9f688c8dc58edbfc6ed43094eeeae871ecd4acb70307f7cb7f11bbeba9c1f96dd69b09d6eff00ebd8d24a72a5d5fb01a29514cdc2770ae4cdeb0e3629770c9408770304a4a6899148e79360446b5380eba6089201292661ac0ccaa8c72280e8953b246931a08091a31624e3182d58b86a1f805740a204646e8844112d8509b04d2e049c42001495319939a0b72272604b611457437a782a9112dd040e504da68253a04362608ad8d091a951f69585a16c2049264937a723d38755080357db597b08cce18061d18874a218311c6f02c1ab835a1f4329b5cd059b1595a268bbe3620150c122c44059435331dbfe94e868295e54d1e4dc9df8e9f3e0508426e583015f24228a040a954a43631232c8d1e47b3426b44608d5d1a52662615b7f696d816480d264be2e4492a1b8420bd2b7985d7990e10d42a6a50a2d96aa21b32c91296482f50c3c49ead0d5c1d03834726e694e05762493a458a39e666b2c505391e29f28c28ed8023c2542f3df83a60e14497a5a8d6e871254eb18da98e77c29d085f3ac0241a899d66992f2118099445a9fa8cc69430dc1502a900dcb188144a650a1128856cb3d5f9d6446b50987601c8149500985012a5170e9cebdb63dfa52ab2dda143638809248841051101028c532d9baf3861321d32c0404ee1264c769218661a342813921aa58025a9cf7220ae603a66ac005ccf5107a540d973892a4005e4c45e681aac904d64e0c77d25e62824be87482c0d5015b5815a1e455bf24db684656eabd7203b1b04563718a0cc3483092160c8940669ea0a1303884f50507613849998b3ba6ac3803c1856d217620b50e05c0002005802c0583f97986ea59ff0064103064140446c3344c39c84434c02952a7450f377ab040a030c805d94035e8ab7314a772359a0109cc90458156026f9a2d2542cab8e1380ae882b6dc31428972a135b66ad489c9d21b85c249492842c616618857ba54a305c244172b452051842a160f584412381820b9e566993a35a37414b080cdb374c5a8b351c8028825858994980d5b6a32970051316252c129592b62660c2159564a6464aa4b800c24205b00b682132f888645c42bc04a8dad1c0957c241b261522b4ba71054414528415565753f806361b7d076acfa3346188ca61104384811719d80ab092a2b9211b21a18cc546fce231285a49bcbcc1ca284151b002a61096f62498f040960040334df268b2519576f2f1e00dd22e13e0502248c4376b93b555b9058c4932e7856b028cf80d5864b99641d8f0ac827374adecf02b21b34353047600582297dca7593154c4c414e2e8e6854f01248e2163d824194d1c4da10803aa614a0f54d225a59122d5c023717013c0b92b80bb4479510ba55833990dddf64f69649942720aaeb95cad0a9190a09118089228dee294eb284a172748855242841526ea358a48819108c482d111d65128adaf41728b95222801ba61c51500a92598828e0488a96b488840aa5c0d9605c454ac58091125aa5552ad597440830c2d2b05455349044e594ad642967c0337b100e00e20c4116eff4971001b2a8ca0a35a38b0e04149849b7bcb4370cada40834b10e7f9c36914e220297c162c49cd5d8182481006b36b6c39c2a0e093c126f52dc2034a3a198926084c2190a9ba9b80e8290884abc593c53cc8811deb346e1beea9e8120dc6aa00cf0295856dc41be334202d86536e48900ec571031432331484973a2ea0ce180b6e106f472d4642a8560934a4a90254e0fe07657b402f002d44e89061102108103945d134f12dc5c00f46931612de9a0b84972a57a84f1515cb0399b809212310864f5ae64086308a0e0541119a114485f8e6dc176181140012ef5c48e1a6e8bab87548f6e4d66f132c53276d272bcdb9ea128359113798f1c0566e2d09ba48d934398334ac4752563dcde94939f329808c2b37839a72c48a0a41b14a4607baaefd771841dd421994920946c0006091cc0b06c112d4a593a919a275c27a58bd34ca61a481582c58b25d8c53b693a262c91e78282cb34cf2ea16443494b197b60b43f73d1525871bb4283134b4af628852be0665b58af9109a08dbaa641218468920c04b0674a833aa3397982ced07729502c24e2937451292ac5e2924264e9984d9dc88ccc2782309efb75038260680147617364045a1b24e598d297ccd98cbae5e907152180a5312085827741a02a637f06cb7a61950bc3810dc1486248590d308bb9193f220590720800b1c83522f2358a0980955a73125062dd0b115b5802d009608682e10dc294e851cd1c4bf1708bd5a44dc436a4127925b3ec5a5ee9b46b4f98bfb376d49591682dad328efa102d140a4403d5494406600502e597205d4a4fff004071500baa0c202bca02d0c291422906df451d415220b6ce3000448a406d34d6d52e1fa51c148318a4819dd016232a5f33ac45affc034000010000000010163c3f947f08cac5c09305493112d1475d92bbf6e1b291849f07163c1b0d2fb5f109cb09bcaaca5565a7aff0510ed14871154cd2570e85b804249860daa0b5d628247500123004b853c0442801824004b691460ac9005dc9c5be40c5aa7ed2b8b40dc4012c4b06d46184aba015540050804bd2e71908b88e44d1589318cee0830ea4c3ad58f28362b1022f2179a06a10310140061cb0a4c31158cab8588628dd65d5a507093b26298480c8148078c9d26c60646129186e28d9abc804302c1024ac0c4abad4a6dae6048c129509291b2d27992b759408322235706ac7139cd19de632bc4d3dc09603eb68e83c3f857d2874a500048ab5d55bb4cbe9534b9401d8eaa636f2a36e853cad041f000004001000b0160a38b06a44cb0090a371565a80236849fb068b789bd356046f718548a68b29a507990b758000300014d6209321920525ba84acacd0815827916b68b0c3111a14c51b1802ec4b058818083c6e24263a4eb259f375165059326647d01e103ea0415cca76d56742e67aa2bbb438c24006018034538612aea24555294aa56f51f68585a42c2058624977a083e000008002001602c15068613fb081c00519480510100ba00012d820a9ea910c06116c690683e80932cccf6405496bae5040e252a0252b75a137207c4060e00212975a24f527e540400a01814c4aff003b10515c065fa760f7fb24bcbd463d0bbea7234503b70161d02c7a784826d739f4cf7c734f279f3fdfde676c56cd1ba94bf4940971529bdf151924c7d77a1578ad3d3b5cfae0f55e29387ab3ead3b052ab2e7ef2c3abe01f52c5d913d9a93c7a715b8e4febe8e11635d0eae0a8d697b7edf6ebe0ac05a2e01e6efa15aab79b1e85fd5a364c6d2c7a1f7f3a1b59fbf7f03628a13bfd32ce81ea455a0d4274f079a904461ab10cf267d2b03cfa3eff00c5658ddb3d0cbd86a70b99b1e85deef529732db68740b141e546b479a5d861c52ab2e7fe0202b0157cb2b3fc142219fcfd3cb4135c1303b268f6744b5b3a27c384e492b92b97c19e859277ad6137963d5b56a9752eee0f7a0e7e4cfa8d8ec1d6a1a455bb2d2197fe285132af8f0f69addf052edde4fcd14795f4e5a56804f469b837c0f6943b144f67e4df4fc339fe951ec9d868cc1b8e7d8feaad83a2f1e935818e81ff00250f69cf531e98791f1f3f02d81f156681dcc79f5ac43f1f359c7f2928a03ad06f2a041ab34c1c7d2b2ea59160896504124ed24d43cf774f3d95f18be282e293974447b3ff0006787d15c3c3aed9395a8d1d937584e98c30dbf937c358c53a51f1eb56ed436aa35d696216b536a08fa2d84d2a001cad8a307cccdd846d66fd1844d2d16eadf0462fa489345a4823109eb24ea14b2ce396eb0002a1a058940a00eaeabb6e6d11658e94af60980b288446c8dcda8ba624f55d3b36a5175fd3a47568b464e7fe0846b3a24f736792f4948b67f327b94d7ae91f461ef5b93a7e64a1e69819aad2a79345ad727aafe06b4fba0bfaa3e77483f75a12f56b040fa2a04b8a73d1e9ea63ae13c5599534b17484f4f475a863fd4c3a27a050f26c083b007f1b5c1d000bb216c3221557290d339ab13d02506214c4a219c884ff00aac4886cdcf4c5001063e88222806a45d8f76b568f818ca279ca018bcc85ff00aa0bd2d1a20e8e2c84212236a001556c004956c067e85a4a09c948e18b76fd0300c765a85faef285d5022e856e01007ed956eabf62d9eacf4a24801051020994a91c82e7d2170088310c3b7de3c4c92c09570365c24132a54bcb755ebd52ee48056f51f53f98c2898195a07ef017694332582481b6c1d499894a47e50080383dd72b76ff00f387bee2c6a8e21ab241ab6296293299148d05943800000229b4823109b6464dd8855c6b27600ee9372af118dc43444834808201165688628462848852bc4216ba2e522e9dac05fa83212c2a311542569842ca549836c8bd466040129b83e1dbb19538ca7b423cc906adbc208e4aa102d38b3c86ca0bd48943262482584841e09c40d1c089220e7d314081a41b869283425a2497114c4f842111c235dbfd2d873a6cc6266f8a0a70e5460bd29a0172840a0b9404c45d4309a64cc26069b53a2095b504b7a884c2a0421548600bab47c7262cad1806a16d2b0b8bb79369931b996e6a965eb237208290b75ab850615c4594fc012a401401692d8050a9e8990c224b057a838cb32ce41bd8d4bd85909192506c328c8a0aa540d12251800a4884851241a0c919ba0877130712eb44bb08a294837c41260b653ee9f2812e2a03c0718644ea8d9c2e1ba113f726874477a5c21c0b8201785742dac4b340b414f6881392a1b0e8322425bdc16eddddab5990a490186129104a25292d882525c54142aaaacd02c78062b0259a30905343f6f2c60ecdeb9270f00c6538703f8a52e1a915080845b2b5ac1166c5672d55d4866d036eb5a9d8058301816ea3f5869bb1939efa00563575084a18fca05004b02447428611162245bd28bbe166b9992f5a64c276a04812c60c84514292221916b8cde09b037f08efb8c160516c43e2ab1796062f4e44259c016e000d2e49b7ab87a2ab2052922b015620938047051d751a621d725d90030e5ee5d4bb960b0433204141e5acc913d443b811d4a6260918c8c2b272ba96b52c0b03c5c286902f80c2c81134d316034402569ac26467bae8834ceb08a264a8d8019780836e177ca939a5bd264642cb901c940a9c8d37983841656424345e8acca41798004d001a054bfbaa1544a001b82d623ee5f10434abe73b1974a482865e37fc356bc2283926c95660972667e20bd325c22622018937d13489d7ca440077024402224f85d1c46b75a053c2458bba2928306a583b923a87803821b03c70001fc3f988121a422d3bd42f602d69d41e062c9c3a96e9b358715f3080e621a2158665ae10588861066568bd1f0500012ab6002eab828eb4d118652d9124142c0de612f0283a3d6c559f29b801088bcaff00c66b1da39cb158692def783d0f0064e045092edd148f9088714e555955cab97c450476b046982e5aa1cf8e7ce425d9a37011811017556eb5f3d543fd07c4d3a7b53220a1a4c4b17c4e9f763df1ab1dd7d80d57005d68ea64e10ca1ac65d272376acac8b9c45a1a322b4c53e901a5750d5455dda6b8eddc1d5a03475d20af3a7e693174dea8b7c5656ddcd6ec08c49356dcc04b6536127895949179c1d9e4928bab000db5fc1befca0cd559806f39962578e0c16320b39731210b3180c20a32592962597299fca96d6305aaefdd316112b782d021a265603e6a588b6dddad1e177aaa41217556081a805288015568085be6524622eb7f3d63b400633a5095d5007499858cae261f3562b765c0ba8850090340ac175a4942cf49bd898502d2451506aa5fca406e3bb4e65553018ae55700ecf79212c50238dd56005b40102939801e455c12c2d005b3624dad500533ab28402c8534c50f81383612696085c41282fb80f9c29d48012aba004b4f14b82d974fc260d9795001d925c8afd06ef306614d4e60b84d7b23804cb0f219a4b341ab1946d429c5b2850083904aa95559f09e90e096d6e40194a2cb8a4a2522887e244422442ee094b2ad6fd85d03281608d8e81181de6240c6d0d7114d6d54140b9c82f4c3aacd1bd89a0546e82d57653c6958a371e1122f891e39430b02ae0370130bda826122417a250c0200400160000343c1e3264a6422c0c085953629940984878420f22e6b178d01eeb276643986c9a1a4257db1412c8e946ab840963421949be288a67159464794e34f0ba6fe3293d2899be22e48089442bc04175ab6abf600892dd40e14888200858ac5cf6e41c9bc58b5d5e93413136e88fd4d42efd946d584c4b2aedac41dc25d71702c8d44410043c5804ea8418399318b13594bc9204ee14549b89cf34ff990551642582bb2210d47cd080762fa12cebb9068d33388a67328f229d66a357798182c5409990aac502306c6991b0110219dedf6ed6248b71c85c2179c308689d6a0be290c4002388fa106305832bd02ef61a304e473e9a7767728208cfd44cd9b84448c24637a9861a99306059050888cb34156583108e43c2622830b1504ab05cac51294c302c9410d356d916d6628c2410b2a60c456b532434cca2492e52c66a230cce6e913212124eb3c422b65d5862d25ca76d66c08c9b0963b51ed50800c537090c1c2969d51b1105840c8312226d47ea02204c1100924e406a0de9c2a684e164e0ef2bb38bd25bf4b1cc3b0186310429772e74db903195d7092d093182f776070da61580092204680cf50b43260b8be75252c5132f85040528212292264fb7284893d2d07a14bf8002141b8e28d0b2400034942aa00482891309b9e2a04b8a4a1761f3e695d1e8120e1198752ac6dda874647d42fc81042e5caea4a72abf59a75ee0163a9425a222a2d406c5264258448e2042584b21ccac6e604b085035864f6aa59b81282c59e19f84b533085235437a71f6d32e12e12024c031050285f91694802b5925100b219c0e70b26c1a40bf0165b0b6b0059592cd1e2296e1346765296349a252b175c152f22820088a594ca8622f118bdb2d5499248f214a058ba045c22820f2125af9c51425744189a0d852c54de10c928a9430ba7b3561642140582310b69db76202ce2eac14970339f1b1048e11325e3ed90c346dda1812290315112205d5e24e29c71cd9844049562c16f1ef4b3a250d7786c75b5e765f46094bb228bc911bf1aa009a48aa375d0c18f482082f9157bff00f2e7ffda0008010303013f10ff00cfa191eb5c2a1b03d68471f6635c1a43227fc62b12e0caf4335819bcddf4c1ef516ed5cbf8c51881e9528d2290e45e94cf373863f35a81cd67d4b7ad47af2acf6dfb54b6dd02bd096a3a00e223eacc738d7144ca5b33e03fce923489e31ece8f78a4bcb95b6c21f6aa072911111322371351fe0620d09b1e78a93f456a95e7a513f65fdd62bd3291c0143a2ae61ea15874b87f727b54f006cd9f5c3ed528db2ffc1430a2cb7e03fc0d7834b2b9575fd742d43c90524c0343cc7ac5601873669227a17e68460d3afe20f9a7c9392c7ebdea31d39a362f3ce1a04e8cde774f64762c24214459827cd0ededf812fcf8a12324c121aa8939864dc140eb419723036b5c829706048128a00da8d4682d68d668def6697a281c3454a9beb668f0917276753a35968b0ee7ef73efda20c4376a6fd5a2d9c09101a312e07065fa3cdda9f7269b9d2cf746ca0879b527a16f59a7ecb83e0a8ef42c10e41f9a359debf11f229b927757eb5ea434a0ace56bd5f8cf589ab2239280130a771cf539dd7e113125e2acd8a9c8ebb8e1308a366b4a6d9b4063dea59c413814dea06e8d9507d01464b359761cfeeac98d9454a8e8a9507703b67dbef9fc141f025f913aca1113572d8172172872add69c42249c1b2b87bb6dea475daeafe8db1dabcd351f2d47cb51f2d47cb51f2d259f9a5a240de3e7ad6b39cbb0cf61eb92f5a1e247d11a66464cc930b75543758c0527c61cb6375d898dd835a72f070195160dd8e8176016911bb01a161e0f7655d7ea3931766bf5d52b1ee3f7473bb9fba6ed8ba4fc5463633f8fbe2be23689407477202859b63b20dba94472e85a57220800c03076a754853327719b15b5bd25318599b2023d4a96a7c32f62ac2c66d46f150b9aed50104491a241e0b25608db343d382641aced4141b536646b1321b9b5234288b00027371bace7153c60d8f038e5d5cbd003f997c50ddab50685a4f5ad203b56c53a73fc2b4661b5cf31432d1b98ece9dfd4a7068723f76fa40c19888733fe6c51258516392b9415bb2d21e121cbdd1c8bc86c2bad11c97d6152ea45d9331a196400002c00400180b14b284c004144bab73243102084ab03911847646d5764be0301fa51280324902489a951244b421a414bb080203717374a66a71eea190185d1ba7b6ea8e50705acaf4e510d93ad4b3b61b285d4ed469fc6e0829c335c51844141ada2d52d72573beb506af5a31956884ab7ac3cf8d5926c78119b41e47f4fddb2337eea8cbb5de1d28c8b10f400f5d0c7837685eab2bd6a4b18a84e154e1170e77f4d3e258292224500d43c0353417982a84c2a0a14a12a2dcc44720d85cd564aaaaab42030180aa0b293104144097a4a6e94633104e100c814a4b14a864340d5e02370510e94e4650fbd4f9f2784a41c1ef5a600425f2447b4f4f0862e46283ebefc8fea8c60ed73eafe282cb3b8b34004041f4405a8d94c7f8a8896a3c0032edf75a94ddbf4fde2a34a187945beb8fcd369208e9f931419655f34961c6b4c12a2fced52cb70673744b42004b89778313e957a8d20505c4580de0b909a0be728af6910c20c8ef2d4608669deb0b22605f325a3598bcd4d778b824e4184c3360a8c2a17a380923c9b282914701a27f313a10f64722481114469430d06c477080fcd44ab05bd0ee8a2f76a7b956761a8a4efb2f7b1d83eac5668a1576ead8a3755a8d5f63fbfb9b2285b6ef8e377b1aa0cd82c801c3f3e9bd10e618eb986830c4413a29668846467bb45f0804921a23551b728342b535243e01a6c0654c002a4015a0d9e8651660405c95618d10896a3f6416e42f0c4c990a50a8449720911885940887a4830b9592d439a543720e433246680284202e5468234bd7034b104e1de7eb43865f8a2a958a0839a628a10274a591f70a4d16edda17dd686ca50b4806c0a370340522d30d681e7b3576c3a9bee34b422f6eeebed5227fb53ec358c94f65fd191e9563a5ba9c4988c33ab48e6a29b54b36a75485d8100c20427228c34d9602dc76ac17b56286a0622d31a80845099164ed8520848644709c3bd4ead69c511f926e9b36b3b9a49869cbd41eb44d891c985f7f6a4af89f3f9faacf1523ccd6a0345067c005a28c8e3ee27a7c6d3aba7cf140d2056c81cdd17614470492c611da90f00a44cd00526e5c98f72925ba09cabb197b05b5160023a0b82420de1b4d290ba97e569b1e5743da62156025c0342b3317a4be9a0425bd58b1eae07b98fce28c9c800100c00400c1406c2830640ab28c8ab210492cdcbd6bb88359621ab60e0de91cb2b37dfb64c7152d65e7e67ea224805e964b0e3f7473e7e3ad478ea7c0cc72f8a375344bede0c0686af432f99ae90cb46ee03a7496a0a2d8d1db5ef6e290c2055ec261e5486888dca5ed901858e43dc51b2d3201e20b394ddd9e6b8d3c37de8932d42224ee3e6237a4eca835f7ee999d2e5aac81102c0247ae89a2256b496ef5253c0b910e550340337982ae24454b203e5d2b9c86beb4cb38d211d446b0ed5ad06096d67631fd812d7643305083065f3efd0a23c423d3e9a812e2af3daddddfadb7cd69a99743ce857a8a757c3d399bb4b46bb503d13f7f6c80a00cb52116f30e86bd71d6a748ef976d1b7b117a1861dbf2e5797c1e3e329a2add6223aca4dd1478d6299d4355a4b1bb40373e7d3f142200ef27b0d2fd012f20c8241494103499022e445c44c98447d13b53b48773394b5ff0042c977623911c23a8e8d0bf142319575bacf143485cd0d2a262ab014d1c0c05c7ca6b8a0ee351156501abbabaab75d5ab316fd0fefe3bfd35025c54a52ccfe0f29e99d0c1bfe06efc6bb278a0fbf2eef820783e7d6a7c2301b7f6d39038ece5e78d3ae3ed646a56065fd1cbda6a10a6d61cad7bf62a163ddd0e8fcb7da3e8daab726ce39f9eb906c08470d2dc74e59b32db873a9ad389e2db275cbbdce6a258b342e7c00314114140e458edbf6a492290991a8743ebbba51adf3baeebab4a80e275e5e3e7a67e9204004aba54e3c6b75fd3865d76a839344c3fa1ce5d37a0e00160c1e0f5005df3aec6b46cdbac65e5e7d83ba9241ecf46eeefa6efda2040195c50cae3abc766bd5b70e6a57a1955fa13f2db69a2e41f5795d5fa4f88195a22b44fab2175b2cec697cc3a3f26e724d489896b4bb7c7e4bd5fba253e226af5bb63f7f8acf31c20fc357e1c913f36f4ad0105404d3c05d7a05ea799f267e5c1df53e9a420656a37a4ada96931edf268e003a749f93c6379d3c12c23eabb06af9c50c1418341b9bf3d82b5782ff0081b1f3afdae7cb40cbd0fce0abfa136d6da777b6c54608e6693f278c6f3a62c7d2d381619eaec7bec54c681cfb01abefba1506eecacbd5fc16a823f70e44b8f469592ee1d897ecef58fcc7e41f15ae3e507d186a6248ba95e7d9d2ac1d558f560ac976e47af87bc505323cde5d57e0838fa73b6ee065e9f970543761818375f95e86d454ef97d9c6ceb97dbc6404bc0cbfa39f9a9ff0b9d0363abbbeae0a1f0356ebbaeaf92df6b2d179c9dde0ef18ab03e00f803b1b54930c5f671b0e72fb7d25025b050c9a7569fdbae369cd4b10d9bfab39077cba6f404402c18fe3ee7c0fc95e6afc57b6203e0fab3b07b7d5cf1eb153443cacbc0dbd8e71566cf72eebabe2e30609c8bf2f05b7daa40aa66597f038778c50210180c7da227024ab602a651d5c57a6c73978d620599301fbd8d7a4b509aecacaf3f8307d27e20d7ce5e29365dad7c9a16eb40405f2bf3fa7aedf68e44195a8a815a4f6363dfa5ea123ee43ab5e98eb4004160f07460d5f37782f4c11acad6e5d0e095f6a2fd734e8d3e79d3ed50887cc06ad463238ddcee78c1ef511b0ec706ef1eb448e0faaeeeebf4a56df4355f3ae95d838c1d39ddd7a5aa106733e47ca3ae3ecc5936c67bba1efb0d49189b3a3a3573ea950b2fcacbfa382de28f050c0f2efc17e9431869367a70f76b0d06206879f7fb044882eda406162040016ac97da25b07210c9b14960625f5b931829702a693528a3e4048108041092425c0b0c0f189b83808ba12a331c2210e4c5467332b36aca1b82bca03aade8d968010650440ba042eb1504b838e15500485b030a4423820251cc4c41a8b8a0aa4852232142f80b178a544395013132065850746a5499b95285d4220c428ed4f73ef680a49a0ab3d5f1299254935264c252eb24641da5097432e85296e3cc504118bb085e62922095450050a826668414a8a5f5ff00369614adb7d9e03451c0afc1d3979db6eb8a4b034c721ba22d80b34fd6a1624416e240e110949a62f606811a50e86a50f1256a6dd7145f2006550f2eaea8d87012c4596d8a22cf20ac708c85a205a34bbe0d59990d86198410ca434ddab284cc48b48f2712297f262c432071881acdda415d633200917072a13792236e851c8d1f70f3476761051b5054ed0284981c76099cd1c49d328655b6ec0424529bbc9943d6ab9c8c40665a3d850d885cd6a9c14783a82353408394400911a90a1e75d3e7d30b968a4b52936e49ac206240a80d393cb7fd8eaf6eb40082c14f9fca64031910b02103228ba61c221b4189c4330866651817e1627b544a22128680020908255036a8bc30646e257a88aee22b1287546b0d1edcb60b03449b0059b45383cb2501289eb76d65a2695c069269048c49025ba814fe21d60887715385eca2d5c769a54105c5926f4a17ac2260cb308402889a041990686676df25090e1b9dc44ddd80d5f2b766200df16c4cb7bd3c678150ea32aab2add6eff34c72a15862531605012545a9f411fc27805220403d547630ca2f85693c350dd4d5a8303944d63d06748a6e5c2849365032b16c529840370987945f06c19017d8a3c84662ce474f7e8d5fa5942703e5846184e425470fcfc14350905337340856804888caaecf90a90b2232531ac012987d150119fb22b45e0ac9bd305578060258257820a49034b12e0ec9452976064b03458b21639e493690004b1e401562d882287764d0508585f324c20e6441330a03846401640c8ba2406635bbff00189a82ac59386e080fe0649e2518edb7c88ed714bd194024590323556c70baa1060c2154805c9d449699252e31398184fa50fa28e2d0463c42772050700563b01a39094225515639cdc34400712c820be7c14cbec193cb239f0a0162b27e2c0b042fc8b2c0b901ada35070b41482c410a6ed7b7d0f908b80c933e4449d94f7052b253dc806da8d102268a2e26b5d99e28ce178116010a1707440a7b384982dfba142634456639490125f240e00660f82310e558291e89170112e3026302d6b83900088323000101802c14d633499211241184124b20d0b615822c032498865288a87105c241293116660c2899300281b05699b424c090446f09091a2405410a1132d3559a4d99013015b01b21825ae0c0026252230488001899c512e52495e025daac042d0208210ba4593571dcf160c858e6665c944262a4850a52c4021211d42ed4024944215fda00689307de98b65d6f4b58fe6f00b31b2140db9042cb8ab472ff824142c05f74c641aee882e62c001088eb4a593b26190565840268a0a9142ae4ce30e6e8e68302886d420b6d96d65476fb52c52d22c0a82e0016bb8d14d522729f18bb09a10282553ca0046c0b1a1c75690dfe414cec91b319a5ad443a09024438445668d75c9593380b6a107295e826c484a8b11274a61362d035e79995baadd4b05b356eb78ae20867051a4d4badce623a621bc39017133c0a4b929204fc8c8dd8aa71ccc2b608b030906115e7b7a2610271001c0cba20fba6711398824b0f7d0a67801880b4924b2462d13e39e880cde4761b3b2535043a9643460281c3b511ced2195803996412df1472b8ea41c970b09292e2ad8817904a62018401851a43728f320449b944a9370d1084446d1de98c75b36a34cc40a13b89717843627346bd0020e1030dc40b744a0fa906e136c05cb45f90a02d4e04984cc96ac8e628180e1a1991adb4417ba2ff1c43513b9b695fd85465418124a0911ad4f612889582013794e8e2838d809911d40061400bda6841621eb89c5fda19c44a302ae6011c1601e5095aaad2ae5c30c90005f031c2eb42d838c07436f5979a892001898eca8926d084a1523924805a8d4295b8a907b2e1c419086d442428242bf3d36aa2208bbd010120b83210a38010c423430211448c92bdede566e976eb782e608e3213ac526b4f3309745ba164be68e081585a88debed08876569e880aaf2190af2de2d4190b7613cc0910a54e8a29655044083076328d1a06424850056c0425240ed5655000216242be4c80d08b229372389551110a92c4d07ef4e86324cc229399b2a0c3da25b6518c488dd37b7f063d55b04a0d42aaaaacb7f0ea548926879b911220c09982ad618750edcb02ea0246294b13d787b8ccac7056c8e9318c1286200580002d408d605dde90cb99aa18d9b1d085900a0c482ef527bebea1718021252365a1af45918c4b02a60960d3e88a070da0996c289bc54293362a54b285509815de9a72131022c14a0a291b31438d2c8808984714acb5f80c614524d18934a4d3b2972b3388f0c38a1a9c614c06532e090242eacbea02f05bd365c742865f80a98db37661222157c639bb74c103480125904b9565bdc8008018012930060a85d75b3081c0048c012e14209e21cc8300dc4447156be2c16317388da290b4c51aab8280da0d5624c0f87576000b1000a1402e580160a09280043010f73a28a75402db08070053b1520a195566ab756eb4b4f25c2b0cb6100b001054a1ebf59d09d92ec98b50c3212d1b90781d42075a504f50e655886eaaab9a3845a68888620b016020229137e03e1736e61ad514839a5fbcd0c9035b326465f1b6b30968182010b62cd28d32c28731116793e137c6e073601e8634af6c8499183b14b9d6caa8ab95735a29c80c409b000020002c454a93772a45ba84149850da9d8a9050caab355bab75ae5eb250944e556993aaa9895060aaac17556f50b5688ddc8e5fd60a791730571112f296a0f7d6d42c3004ac002c14b5f9318082bc80480d829eb05a4b11a04051281301fc8f48a54a5145016f979fc7d344e29f001047d7b33a3bfa9b1e8bb252edc72aebaa95f5f0918b9c63bb8ec4bc56a4278d3f7d6dd3ef1cf8298d3ba84cd5bfd362456e53f60852e4628ccfa81d8897a03bd0ee6e38fd9dd40020c7de2296a6ac6afb7d330b71f3fc157bb4fa3172bb8357a05ded43a7d6bbe983dfa518a364072d2fe8e107758a3b41ddea58ec540a729ee38ed07dfadd4efa76d37fd3621919a09ba89fe11b17a3987f10dd96974757e414e19b17772907405da84c8aae5caf55bbde9ff004bfaf9af82d9f6fdb4e41cd77b6280082c7fc00960a604525345eb490fd3d4cbf5fe9b9e9c98d387f900b968db2fa17f6ad62742ec5d7b537dc0be12ef73d2820c0d0b15172a03427fe24633893c0d9f109ce691187e9cc06e0c51d081c87ce6902443a3fbad757897f314d9ab9663e63de9cfc631effed5cc1cdbbeb4ab97fe4881a21ea67f7dff008a816540a2593f9a794adc1a5feea7633823dda913f97e6ae04b737febdbe9629ad915f16073a5091a7b9a336e6dd6901de4089d987fe0ccb2f27e4e4a9d13ee3aecf93e82a808a634eca685b5251ab7ac009fa07095c012bd02863969df447a552f6333e8523f2ab896d639ecbb091de900053a91a5d27eaeb5a29886bb807714795a44111c22591dca62abfa40a52999335199591a5bd220a03fe0a79b828284ef87f27b1402e5d04f9fc54da8ecfea97c7b7f4a5bfabfaad0cbb3faad1cfa7f749c2a5380a53f4b362a275017f44fb039abc9cbfdf25e62a510b405dcdfdd54dc2d52beadff008a0dda2b21c8b8f2244e2454d89d6e02bb88d464b2de2432e207fea89a3e7d7e92e36b47443276c76a2d3ca8e0baf2a14db720ff00d56cecb2f28290020a2241329498d04aa001abf410d6c25c42e778bf571561e7e01f1ec65712d3cd6119572f9b183ec40ac18160140800ca2440d3002769205b2011499049cfde3107c49973320d60212a8114e0892d4508700b340136ff00279e5bbbf994d4cba06ef9bd1a2186765cee2d0da340176cc956eaff00ced2862ec5f727b892e95b345d12149cc259955568326ac9d8b277dc4c8d5f10f117d07419d0f8a05665a1d96794ad786ca0691c30396f24c71eb285c8924f5986ec2eca0325af2ddd53078b54996d04174167c0e393e18997eb014b0c92c4d0e50ff81ee37f0355ba2ca8cdb9f25891ad5b371bb50c00052de29a4b30d385cc900cb3fb66a030c5864e6016a8e9918140fa481a4432257baad6460e489cc45b34e9241d65b6111490c25914770a58911a2ca66b87128946b19fa5205ea60b34c924523a5901aea2002991c44b81a8a0d165ad6fc38c1dae04e21c7666d180a18d925d00082f48814990fc2075d55802500281070e0158da47059150895254dc9e0df0a1ddd407808e3b86052f30188008a9d2420a102900864142428a3948960436402f31e94c1362114852de98c0b7c0fdda143704c96859b4b41d445710632727605ab2a5956c097442720d91e0e360083c228d4daa0d7419e9b8f149169ec5ad27391245300e944a500001002c0000160b504dac684c8d2ad14040c4d16a7700dd99e27e8af6f2ddd4c3886cac072209501b210bca4754d7955974afa5f57903a0a87997a92637adf9940c5b4d393b0359ff00b1b158e221c053fc84a26e48314336e8418a04363043ce8110e9461979c74534044824448aa43e524dec4b4265bab6f0d00925b682cbb39d05eb4729c1c108b088673a595899683dc20d2a0d1030410a06e8e5665a39493aab4ba870392f4036614a4b022d50e2c25ba1260e4d173a11124e805d944d1a71974cb8aaad832c2de060727421b4003529e59210be3e18912a0c96114025621866ac9935c3320870d4869c1b4b28965223042c503c484c25de2ceb86903f32a2db46735c08b8a50a0f440ab689615aaa6ab44e96791492aeb5c7533f728448ab80a67115839d8fcb53fa2b570c24cc44766874ec4518921825df25632ca89400ec621dc51f0d599275dbcb42f8d3aa0015270c127653bf82acac1d55abddf079e5bbbc2ea11677a03622ea75d5b016e44e4cab3f1a49ade07b83b86ada53fbd87ef277a40984aab0017556c0175714902e64bd02240064a9704d23c5b287d7735b30aa3712b2a9164da16a1e5bbff868fe7204c2baa130ed2c757c3622d664c360020cfa50cc85e8e801001802c074f166d3bfcf51bcf60c0931e2a144859bbb6133a901b002c15ccf53813fb56ef1504da013233d62639b6635fbbd2853b2ac7e5d83569a2cc6fa97e0ff005c5ae18108d3b8670722504a08180200e0082876313f0314dd401da464f0406319d80126cc035975ddd94b90398eaaacdba6018b2ebc64ce9e03abae84a36ab32780926915dd25490ed8b449dca443e5bbbc2e2c70b2ef133a8d66ce283eb8f0200e8156f05c52af51725e7a22586d866411cecd1a42658404805b1748288940d306deaf8b70b801a733a95e5bbff968ff009b163c1b1b04626a952bd82ce25097c8032394b08490a6c8528148145493831968b13150c04068fc820e04b740f658c2009c6440004acb2919219411494c0a0ba4c3a4090da93e22089961a1322245d810c2408be6149ea9068bba548a5255a8fca8850828051f7010a56080baab0072b47ac11c5f81beabec0ee2f30dce5eb67072d18670625995960044c98a2e21530a5d6aea8510704c1b001e050c43005b150c4e62cdd23c3d43a0f0bf801619c11c03565440588d1d8b1295cc315c9045994d8903a8376920770808944c82f08d8148a1f130ca21498c78376104994228d817273469a62bdf5f9d906c6d36697a28556eab755dd7c16a658f8a132b28360c49a2c6425d724abc8b8ab648810645f320118926c520785ab7294367869fb098d1c9821722e96a3cc58301c0f18899bf87f7a252c77d116ccd90a6ba92832236322f3562e8446342072af088a0a1a308c64da14c5b7490b50c64e427a94f4752900a2a9c93998610bde6d29336004243374984441422791e8298246206ab9038b3621045ac9203163188d28a851771030b70db86592a60d6f97bb289973444015c146a22c5c82690e8c1152576c054be6cdc4c103768b7de931fae9997246d7fb709980d98597b6f72e94bb4959ea0bd56feff0042e62c6ba1d5c1de93232f18f5fd7ad3f5072c80f690758a5b543863d6041208ac083382292ac809a25d1422409144d4e509415ccb8c28642dd02cc7ce5d0ef3481882a6ea873f1b21c413d6c4c430d9a319e22cba102f5844c98634d719240b5bf72cbc366869857a121dc1b8de697a51722b5178a4999b40722546509b01214a33317a49080648548a60232270a5a89b219eac36489d942330515c1bce8f954b3438a9a10568442c1c04c02437918c5506198cb729424880a04e5116c99d22277908c1b09a06764200120110180e1fb7fdd09cce916391a634a4fce2d5ccd92594aa50b00054461b278e6c66ae4ddcaf603debb935edb83bd6c4cb727d040e8d55be430603a058ec7d61359660134794355da54882c394269012acf4108959b09c7a0ee32c129513e5e697b01ab28b016ed965041ca12510a6c92d43fc00f04f8d8c81355cca21360dcb3a242251134836a548a989815b209a98011d0e97eb9b5c10dc29384e6518c71f081b88c5355f6d11998d980b93866d45a60194cace54acce36450d80d06cc49366e816c8e7aa20351d772c322164937f72a79d76a053d060ebc4a8d730030a16434dc9c147b83ad8e6dccd809685d6e3c2e6090322b86cbf6c44ca0cd99581b965088a8bc860a44aab49defb9ee55ca8c92df376b45e9a0c4ec07a178f50f5a61f483dff00cf3533ba228ee1ea552675a4c7b0f67ff2efffd9

	IF EXISTS(SELECT * FROM ASRSysPictures p
		INNER JOIN ASRSysSystemSettings s ON s.SettingValue = p.PictureID
		WHERE s.Section = 'desktopsetting' AND s.SettingKey = 'bitmapid'
			AND p.[name] IN ('Advanced Business Solutions Wallpaper 1024x768.jpg', 'Advanced Business Solutions Wallpaper 1280x800.jpg',
			'Advanced Business Solutions Wallpaper 1440x900.jpg', 'Advanced Business Solutions Wallpaper 2560x1600.jpg',
			'ASRDesktopImagePersonnelnPost.bmp', 'ASR Splash.jpg', 'ASRDesktopImage2005.jpg', 'ASRDesktopImage2005b.jpg',
			'ASRDesktopImage 1024x768.jpg',
			'ASRDesktopImage 1440x900.jpg',			
			'ASRDesktopImage 1600x1200.jpg',
			'ASRDesktopImage-1024x768.jpg',
			'ASRDesktopImage-1440x900.jpg',			
			'ASRDesktopImage-1600x1200.jpg', 
			'COASolutionsDesktopImage-1024x768.jpg',
			'COASolutionsDesktopImage-1600x1200.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201024x768.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201440x900.jpg',
			'Advanced%20Business%20Solutions%20Wallpaper%201600x1200.jpg',
			'HRProP.bmp', 'HRProPP.bmp', 'HRProPR.bmp', 'HRProPRP.bmp',			
			'HRProPRT.bmp',	'HRProPRTP.bmp', 'HRProPRTS.bmp', 'HRProPS.bmp', 'HRProPT.bmp', 'HRProPTP.bmp', 'HRProPTS.bmp',	'HRProT.bmp'))
	BEGIN
		-- Set backcolour to white, image to our newly inserted one and tile in the centre
		EXEC spsys_setsystemsetting 'desktopsetting', 'backgroundcolour', '16777215';
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmapid', @newDesktopImageID;	
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmaplocation', 2;	
	END
 
/* ------------------------------------------------------------- */
PRINT 'Step - Indexing Update'
/* ------------------------------------------------------------- */

	DECLARE @sql nvarchar(max)

	--(1) too many udf_tab indexes, delete all then let SM add correct ones
	DECLARE @tablename nvarchar(500), @indexname nvarchar(500)
	DECLARE c CURSOR FOR 
	SELECT o.name AS tablename, i.name AS indexname FROM sys.indexes i
	INNER JOIN sys.objects o ON o.object_id = i.object_id
	WHERE i.name LIKE 'IDX_udf%'

	OPEN c
	WHILE (1=1)
	BEGIN
		  FETCH NEXT FROM c INTO @tablename, @indexname
		  IF @@FETCH_STATUS < 0 BREAK
	      
		  SET @sql = N'DROP INDEX [' + @indexname + '] ON [dbo].[' + @tablename + '];'
		EXEC sp_executesql @sql;
	END
	CLOSE c
	DEALLOCATE c

	--(2) udates to defraging index sp
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRDefragIndexes]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRDefragIndexes];
		
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRDefragIndexes]
	AS
	BEGIN
		  SET NOCOUNT ON;
		  DECLARE @objectid int;
		  DECLARE @indexid int;
		  DECLARE @partitioncount bigint;
		  DECLARE @schemaname nvarchar(130); 
		  DECLARE @objectname nvarchar(130); 
		  DECLARE @indexname nvarchar(130); 
		  DECLARE @partitionnum bigint;
		  DECLARE @frag float;
		  DECLARE @command nvarchar(4000); 

		  SELECT object_id AS objectid, index_id AS indexid, partition_number AS partitionnum, avg_fragmentation_in_percent AS frag
		  INTO #work_to_do
		  FROM sys.dm_db_index_physical_stats(DB_ID(),NULL, NULL , NULL, N''LIMITED'')
		  WHERE avg_fragmentation_in_percent > 10.0 -- Allow limited fragmentation
		  AND index_id > 0 -- Ignore heaps

		  DECLARE partitions CURSOR FOR SELECT * FROM #work_to_do;
		  OPEN partitions;
		  WHILE (1=1)
		  BEGIN
				FETCH NEXT FROM partitions INTO @objectid, @indexid, @partitionnum, @frag;
				IF @@FETCH_STATUS < 0 BREAK;
	          
				SELECT @objectname = QUOTENAME(o.name), @schemaname = QUOTENAME(s.name)
				FROM sys.objects AS o
				JOIN sys.schemas as s ON s.schema_id = o.schema_id
				WHERE o.object_id = @objectid;
	          
				SELECT @indexname = QUOTENAME(name)
				FROM sys.indexes
				WHERE  object_id = @objectid AND index_id = @indexid;
	          
				SELECT @partitioncount = count(*)
				FROM sys.partitions
				WHERE object_id = @objectid AND index_id = @indexid;

				IF @frag < 30.0
					  SET @command = N''ALTER INDEX '' + @indexname + N'' ON '' + @schemaname + N''.'' + @objectname + N'' REORGANIZE'';
				IF @frag >= 30.0
					  SET @command = N''ALTER INDEX '' + @indexname + N'' ON '' + @schemaname + N''.'' + @objectname + N'' REBUILD'';
				IF @partitioncount > 1
					  SET @command = @command + N'' PARTITION='' + CAST(@partitionnum AS nvarchar(10));
				EXECUTE sp_executeSQL @command;
		  END
		  CLOSE partitions;
		  DEALLOCATE partitions;

		  DROP TABLE #work_to_do;
	END';
	
	--(3) & (4) add primary key and indexes on accord table
	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAccordTransactions]') AND name = N'PK_ASRSysAccordTransactions')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysAccordTransactions ADD CONSTRAINT PK_ASRSysAccordTransactions PRIMARY KEY CLUSTERED (TransactionID) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAccordTransactions]') AND name = N'IDX_CreatedDateTime_Archived')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX [IDX_CreatedDateTime_Archived] ON [dbo].[ASRSysAccordTransactions] ([CreatedDateTime] ASC, [Archived] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]';

	--(5) delete primary key indexes from views
	DECLARE @table nvarchar(256)
	DECLARE c CURSOR FOR SELECT TableName FROM dbo.tbsys_tables
	OPEN c
	FETCH NEXT FROM c INTO @table
	WHILE @@FETCH_STATUS = 0
	BEGIN
		  IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[' + @table + ']') AND name = N'IDX_ID')
		  BEGIN
				SET @sql = N'DROP INDEX [IDX_ID] ON [dbo].[' + @table + '] WITH ( ONLINE = OFF );'
				EXEC sp_executesql @sql;
		  END
		  FETCH NEXT FROM c INTO @table 
	END
	CLOSE c
	DEALLOCATE c

	--(6) add audit primary keys and indexes
	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditAccess]') AND name = N'PK_ASRSysAuditAccess')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysAuditAccess ADD CONSTRAINT PK_ASRSysAuditAccess PRIMARY KEY CLUSTERED (ID) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditAccess]') AND name = N'IDX_DateTimeStamp')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX IDX_DateTimeStamp ON dbo.ASRSysAuditAccess (DateTimeStamp) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditGroup]') AND name = N'PK_ASRSysAuditGroup')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysAuditGroup ADD CONSTRAINT PK_ASRSysAuditGroup PRIMARY KEY CLUSTERED (ID) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditGroup]') AND name = N'IDX_DateTimeStamp')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX IDX_DateTimeStamp ON dbo.ASRSysAuditGroup (DateTimeStamp) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditPermissions]') AND name = N'PK_ASRSysAuditPermissions')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysAuditPermissions ADD CONSTRAINT PK_ASRSysAuditPermissions PRIMARY KEY CLUSTERED (ID) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

	IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysAuditPermissions]') AND name = N'IDX_DateTimeStamp')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX IDX_DateTimeStamp ON dbo.ASRSysAuditPermissions (DateTimeStamp) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';

/* ------------------------------------------------------------- */
PRINT 'Step - Changes to Shared Table Transfer for RTI/PAE'
/* ------------------------------------------------------------- */
	
	-- Update existing columns for Employee transfer
	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 205) = 'Stay in UK for 6 months or more'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Stay in UK for 183 days or more''  WHERE TransferTypeID = 0 AND TransferFieldID = 205'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 206) = 'Stay in UK less than 6 Months'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Stay in UK less than 183 days''  WHERE TransferTypeID = 0 AND TransferFieldID = 206'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 198) = 'EEA/Commonwealth Citizen'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''EEA Citizen''  WHERE TransferTypeID = 0 AND TransferFieldID = 198'
		EXEC sp_executesql @NVarCommand
	END

	-- Add new mappings for Employee transfer
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 210 AND TransferTypeID = 0
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (210,0,0,''Expat Indicator'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (211,0,0,''Occupational Pension Indicator'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (212,0,0,''SCON'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (213,0,0,''PAE Predicted Pre-Status'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (214,0,0,''PAE Notice Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (215,0,0,''PAE Date Pension Notice Received'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (216,0,0,''PAE Status When Notice Received'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (217,0,0,''PAE Date Opt Out Received'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (218,0,0,''PAE Valid Opt Out Notice'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (219,0,0,''PAE Date Membership Terminated'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
/* Step - Parallel code stream readiness */
/* ------------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_generatesysprotects]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_generatesysprotects];
		
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spadmin_generatesysprotects]
	AS
	BEGIN
		DECLARE @icount integer
	END'
	

/* ------------------------------------------------------------- */
/* Step - Reset Password Parameters */
/* ------------------------------------------------------------- */

	--Get the existing Mobile Setup parameter
	DECLARE @parametervalue NVARCHAR(255);
	SELECT @parametervalue = [dbo].[udfsys_getmodulesetting]('MODULE_MOBILE', 'Param_UniqueEmailColumn')

	--Update Personnel Module Setup
	IF ISNULL(@parametervalue, '0') > 0 
	BEGIN
	EXEC spstat_setmodulesetting
				'MODULE_PERSONNEL',
				'Param_FieldsWorkEmail',
				@parametervalue,
				'PType_ColumnID';
	END

	-- Create the physical reset stuff
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_createsystemlogin]') AND xtype = 'PC')
	BEGIN
		EXECUTE sp_executeSQL N'spadmin_createsystemlogin';
		GRANT EXECUTE ON dbo.spadmin_commitresetpassword TO [OpenHR2IIS];
	END



/* ------------------------------------------------------------- */
/* Step - System farmework/fusion configuration dependencies */
/* ------------------------------------------------------------- */

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spsys_getmodulesetting]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_getmodulesetting];

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spsys_getaccordmodulesetting]')	AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_getaccordmodulesetting];

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spsys_getaccordmodulesetting](
			@moduleKey AS varchar(50),
			@parameterKey AS varchar(50),
			@paramterType AS varchar(50),			
			@parameterValue AS nvarchar(MAX) OUTPUT)
		AS
		BEGIN
			SELECT @parameterValue = [parameterValue] FROM [asrsysModuleSetup] WHERE [ModuleKey] = @moduleKey 
				AND [ParameterKey] = @parameterKey AND [ParameterType] = @paramterType
		END';

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spsys_getmodulesetting](
			@moduleKey AS varchar(50),
			@parameterKey AS varchar(50),
			@paramterType AS varchar(50),			
			@parameterValue AS nvarchar(MAX) OUTPUT)
		AS
		BEGIN
			SELECT @parameterValue = [parameterValue] FROM [asrsysModuleSetup] WHERE [ModuleKey] = @moduleKey 
				AND [ParameterKey] = @parameterKey AND [ParameterType] = @paramterType
		END';


/* ------------------------------------------------------------- */
/* Step - Change 'cancel' icon for Mobile Designer               */
/* ------------------------------------------------------------- */

	UPDATE dbo.ASRSysPictures 
	SET [Picture] = 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D00000025744558746372656174652D6461746500323030392D31312D32335431313A35383A31362D30353A30300FCF81FB00000025744558746D6F646966792D6461746500323030392D30382D32305431333A30353A30362D30353A3030EACD77810000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000A27494441545847C5577750D4D71AFDF9E64DC67999BC6A9279CF90A6269A44A5833431028A41DC200252A52ECB5277A9A24804155450494C2C10A4091215569AD291A6B0F40551C10682C8A2027625E7DD7B75091B277F6767BEF9D5FD9DF39DFBDDEF9ECB01E0FECC60E052A9946B6C6CE4EAEAEAB8CACA4AAEACAC8C2B2929E10A0A0AB8BCBC3CEED4A9535C48B04835302040181414982A16054983C562B9581C341518E02FF7F5F5910A04DEA9DEDE7CA140C0574D4B4BE3323333B9ACAC2C2E3B3B9B3B79F224279148B8C2C242F6ED9A9A9AE9A4D9494343C31FAA101616BA92802647476F1B4F4B4B45556505BABB65181C1CC4E8E828060606D0D9D989F2F232241F3D8288F0D0710F0FB7142F2F77933F52B6B4B45499407575F51B042222C23E080D15EDDABD3B7EA2A8A890000EE0F1E3C72C1E3D7A84C9C9494C4C4CB078F8F021BBA671FDFA75E4E59D4654D4D6097737975D9E9EAE1FFC9E085542718F9DCC64949494C48585056B888202F332323270EDDA350678EFDE189A1A1A20C9CC446AC25E448788102014205A2C62D7928C74349CAFC1DDBB238C546F6F2F8E1C39844D2E8E79EE6ECE5AF4BB0AD0D3A74F2B13282A2A9ABE2116076A848488AB0ACE4858B64F9E3C414D4505F644462291EF855CBE0095427F340585A0852F448DF5469C59FB2D7EB65887681E0F3B4581A82829C683070F70FFFE7DE4E464C3D3C3ADC6D5C5515341E0C48913CA04F2F3F3D98DA0A000153F5F1FC919024E81EF0C0F23EBD04FD8CBE7E35C7028BABF8B456778243A82C3D11E188C366120DA04FEE81004A09344A3A92552B50CB0651D0F2989094CBDB1B131646767C1C9D14EE2E26CFF21C5397EFCB832015AA5F4815028884F4F3FC6C67468E8368EEEDD83543F3F74ED3B809E989DAFC0C5A168F30B422BC9BEC58D0FA9B3075AEC5DD1BAD115ED8E9EE874E2A364F132EC5C618AA46D51E8BB7A15232323F8F160129C9DEC76539CF4F474650239393904DCDB2476FB779357C91F28818C9F0E222D28083D89FBD12E0A41BD950DEA88D4B5ABD7A2C55300A98B079AED37416AE38496F58EB860B81A4D86E668365C83768B8DA85EA48B1D465FE350DC2EDCB973076DAD2D080F174F3A3BDA991E3B764C99009DAB7C2F8F9F0B0ACE30F0EAD252ECF7F783ECFB83682719976BEBE3E24627DC3A9E83D2C51A285BAC89663B67345B3B406A698BDA45DAA8FD4415D763F6A08BE784C6799AE8586985B3F3D410B17A0DCEFC928BDBB76FE37856061C1C6C53535252940908BC3D552337874DDCBC791372B91CF16121A8DE1E8B9ED85DA8B5E0E182BD3351EED5EF3929AE8A25DAA85CA88626732BD47DAE85FAF91A787CA57FFA9DAB9E22B47CAE874E034B1C232422DD5C71E5F265B4B7B7C3DFDF67C2D1D1565D691A7A7BB9FBA6241F66D9D7575723C9CF17BD7B1221DBB61DB56B2C71B7AA06BFFEFA2BA6A6A6D8F1997C0CF57A26A85DA08E263563063EF3F9E34B57D1B6400FBDE644A1FF2E4598F14AE4E79E405F5F1F12F6C6C3D1DEC65F89009FEF9E5A567A8E3592E47D892808DF8C4B3BE220DBB28D8DB92C721BA65EBCC00B12CF9E3DC3CB972FF1ECEE28BA9DBCF1B05DC6AE9F3F7FCE8E532F5E622024161D1FE9E0F24A5BC854CDF0C3222D248487B1DE907D3C13F6B6EBD39408087DF8D216A994CDDD2D022F34C6EC404F740C649BA3D0E4E4864A755DB4BAF2F1E2E9D3E92EF8949C33228410254E9B157D3E18188DB6BF7C8C8ED90BD0F59FA5907DA287BCF7BE40D0860DE8E9E9C1B97367616763D5A244C0CFD75BDEDB7B89CD59AF8D36E88ADB4308C4A29B64DEE62544B5BA1E8AFFF61E2BB8A70FC6D97B746AD1B5801E69DD3C9B7C8841E11674721F4F47D7EBF3FAD9F3E168610E994C860B171A61B3619D5C99805030459B06FDA0839D357A63E2D013198DC6D5EB7076CE8728E2DE4109F70F9471FFC28D98048C8F8FE3CA952B6C4CE99176BCF1FC52C8B84FD0CD7D8A1E12338F5DB33E8595A5393A3A3AD0D6D606EBF56BA79408F8083CE5974995D24C36AD2755EF2140D5979A289CF50E8AB9BFE31CF74F9473FF4687A9352687EFB0D5AFA9A909CDCDCD2C6866E3237731EC168E5E6E3E2ECF087ADD347B211CCCCDD0DADA8AF3E7CFC3DA6AADB2025E9E6ED2A6A68B64C1B907B18B23D23E5B8C4296F52BF00A02DEAAB30A1323A3209E01C42BA0A5A5850DC5AB713D07E227F08014E6886B04FAB9CF5FC767E8E33E43C91C55F859AE636425927CACB75A2B5552C0CDC521B590F47F5A84DFC77C879D0BBF9ACE9C82F76C70C718C99C02E7E6E6A2AAAA0A13B70631E41A86876DDD4C0DB2C281980D8C0E0D63D4251CD7B985D37174BE0E627D84B878F1228E9215D28A67A13C0B1CEC37F81E3890C896D1F2E222F81918A078D6AB31AFE6E6602CBF84C99E9C9C0CE292F0687804FD0BCD70855B801B2ACBF1A8B71FB5B5B5204E8814D9053C69BF849BDC17B845E2E6AC2F10AAB31C69870F3302515B2308816F94FB80ED069E9A0FDF7DA2BFBF8FF5ED10172724A9CC27D2FF07E7B977D1BDDC8AA9438BE8F1D87DF47D64C4C0A9BCD748A603EF1BE049FF2DB6FAD161BCE7118541EE4BDCE6BEC2E9B95A0822F2131B86B367CF6293B3FDC4B7BC351A4A434057274B8B55A95959E9ACC20B4EFE02A1A121F2DE9E8B3AEE7D5CE4FE87CB7A3CC80F65E1D2EC45A4C87E03BFC92D6299DE7E4B15E311FB215FE186A1BF2EC510B7181D6FAB2358D71847F7EF63359290B01B16DF981CA35E5189005D9D2C2C56990ABC3D26BBBA3A59B6FBA3B622525D07156F7D80666E2EDAB90FD934A3557D9510B8460AEDC66B7045B64324630A3C4CA2FF2D35C42F3144ACB78081E7E7E7C1D5C57ED2628D89E91B8684BA58CA6895E98AF8F8B81D6C5ED3757C87BF3FA2D47450FAF63CD65CE8FCA6D9F713E96991D1711E782D3505A7C077B825E82299EF26E05B1D9D514A5A3C31BDD81C1102335323E607DEB064BF1130565961AC2F397CE820EB09B43B2646442040CF08192AAAE899356F3A7B85F434FB69F0594B2099AB8B10227B2CB16EB4EDD2EC77EED88E95C67A12331323E688A8457FA3061437562C5FA66966B2BCE6C71F0E105734C4D6F1DCF43488ACAD11AC638823644A95BEAB86CED98B49857F899ED94B513D470B69F3F511A1630C91256F7ACC69E6B1B1D1303131AAF9DA584F5B8131D38332267413A17848DDAB81BEA6A6BEAE46FEB6AD9B515F5F871B376E90BD40374E1247BC2F348C48EB02818D35ECAC79E05BAFC7567B0724048A90997C94553BCD3AEFF42984880360A0A7213134D4D69EE98A49E352566026010591653A6A2ABADA6A71F6B6569349A447341022B4EBD1B6ABE8E96447C59A109DDF149406DD4350C9C98233A9BF4C3DCE505F5345F14DC5B1A2A24299C04C97FAFB9775B5979AE86A2D49B15AB76622224C4C3AD94FA0969DF6740A4C3E865F727370607F024481BEE0599A4FE8E9AAA5E8EBAA99FEFE5B8A6BF2DF374D29254195209B11E65A1541AFE97D3286EA640EFBD958F3D236396D94DAD97C2B27BD63CACCC450BEC2689994649BB64C47D54F4F474D9DEE25E99E52B11F2C2E2EE6A8ECE5E5E56C5F38732BF8A7EE8CA9227F3A81FF03E0A598E0D6C16B7C0000000049454E44AE426082
	WHERE [Name] = 'absCancel.png';


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '5.1';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v5.1')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
EXEC dbo.spsys_setsystemsetting 'database', 'refreshstoredprocedures', 1;


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF;

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v5.1 Of OpenHR'
