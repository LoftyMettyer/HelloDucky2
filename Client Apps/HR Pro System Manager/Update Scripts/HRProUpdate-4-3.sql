
/* --------------------------------------------------- */
/* Update the database from version 4.2 to version 4.3 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion int,
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16)

DECLARE @sSQL varchar(max)
DECLARE @sSPCode nvarchar(max)

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

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '4.2') and (@sDBVersion <> '4.3')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2005 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion <> 9 AND @iSQLVersion <> 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of HR Pro', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
PRINT 'Step 1 - Create table rename function'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRTableToView]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRTableToView];

	SET @NVarCommand = 'CREATE PROCEDURE dbo.spASRTableToView(@oldname nvarchar(255), @newname nvarchar(255))
		AS
		BEGIN

			DECLARE @sqlCommand nvarchar(MAX);

			IF EXISTS(SELECT name FROM sys.sysobjects WHERE name = @oldname AND xtype = ''U'')
			BEGIN
				EXECUTE sp_rename @oldname, @newname;

				SET @sqlCommand = ''CREATE VIEW dbo.['' + @oldname + ''] AS SELECT * FROM dbo.['' + @newname + ''];'';
				EXECUTE sp_executesql @sqlCommand;
			END

		END'
	EXECUTE (@NVarCommand);


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRUpdateTableStructures]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRUpdateTableStructures];

	SET @NVarCommand = 'CREATE PROCEDURE dbo.spASRUpdateTableStructures(@name nvarchar(255))
		AS
		BEGIN

			DECLARE @sqlCommand nvarchar(MAX);

			IF NOT EXISTS(SELECT o.name FROM sys.sysobjects o INNER JOIN sys.syscolumns c ON c.id = o.id WHERE o.name = ''tbuser_'' + @name AND o.xtype = ''U'' AND c.name = ''guid'')
			BEGIN

				-- Add additional columns
				SET @sqlCommand = ''ALTER TABLE dbo.[tbuser_'' + @name + ''] ADD [guid] uniqueidentifier, updflag integer, deleteddate datetime, recorddescription nvarchar(255), lastsavedby varbinary(85), lastsavedatetime datetime''
				EXECUTE sp_executesql @sqlCommand;
				
				-- Add constraints
				SET @sqlCommand = ''ALTER TABLE dbo.[tbuser_'' + @name + ''] ADD  CONSTRAINT [DF_'' + @name + ''_guid] DEFAULT (newsequentialid()) FOR [guid]''
				EXECUTE sp_executesql @sqlCommand;
			END

		END'
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 2 - Drop existing system triggers'

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'DROP TRIGGER ' +  o.name + ';' + CHAR(13)
		FROM sys.sysobjects o
		INNER JOIN tbsys_tables t ON t.TableName = OBJECT_NAME(o.parent_obj)
		WHERE o.xtype = 'TR' AND (name = 'INS_' + t.TableName OR name = 'UPD_' + t.TableName OR name = 'DEL_' + t.TableName)
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 3 - Rename base user tables'

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'EXECUTE dbo.spASRTableToView ''' + TableName + ''', ''tbuser_' + LOWER(TableName) + ''';'
		FROM ASRSysTables;
	EXECUTE sp_executesql @NVarCommand;

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'EXECUTE dbo.spASRUpdateTableStructures ''' + TableName + ''';'
		FROM ASRSysTables;
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 4 - Rename base system tables'

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysTables'', ''tbsys_tables'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysColumns'', ''tbsys_columns'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysViews'', ''tbsys_views'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysScreens'', ''tbsys_screens'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowElementColumns'', ''tbsys_workflowelementcolumns'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowElementItems'', ''tbsys_workflowelementitems'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowElementItemValues'', ''tbsys_workflowelementitemvalues'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowElements'', ''tbsys_workflowelements'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowElementValidations'', ''tbsys_workflowelementvalidations'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowLinks'', ''tbsys_workflowlinks'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflows'', ''tbsys_workflows'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowTriggeredLinkColumns'', ''tbsys_workflowtriggeredlinkcolumns'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysWorkflowTriggeredLinks'', ''tbsys_workflowtriggeredlinks'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysOrders'', ''tbsys_orders'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysOrderItems'', ''tbsys_orderitems'''
	EXECUTE (@NVarCommand);


/* ------------------------------------------------------------- */
PRINT 'Step 4 - Data Sources structures'




/* ------------------------------------------------------------- */
PRINT 'Step 5 - Drop all HR Pro defined object (schema binding)'

	-- Views
	--SELECT @NVarCommand = @NVarCommand + 'DROP VIEW dbo.[' + o.name + '];'
	--	FROM dbo.sysobjects o
	--	INNER JOIN tbsys_tables t ON t.tablename = o.name
	--	WHERE o.xtype= 'V'
	--EXECUTE sp_executesql @NVarCommand;

	-- Calculations
	SELECT @NVarCommand = @NVarCommand + 'DROP FUNCTION dbo.[' + name + '];'
		FROM dbo.sysobjects
		WHERE name LIKE 'udfcalc_%'
			AND xtype in (N'FN', N'IF', N'TF')
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 6 - Add new calculation procedures'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_absencebetweentwodates]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_absencebetweentwodates]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_firstnamefromforenames]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_firstnamefromforenames]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_getfieldfromdatabaserecord]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_getfunctionparametertype]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getfunctionparametertype]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_initialsfromforenames]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_initialsfromforenames]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_isfieldpopulated]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isfieldpopulated]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_isovernightprocess]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isovernightprocess]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_maternityexpectedreturndate]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_maternityexpectedreturndate]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_nicedate]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_nicedate]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_nicetime]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_nicetime]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_parentalleaveentitlement]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_parentalleaveentitlement]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_parentalleavetaken]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_parentalleavetaken]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_propercase]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_propercase]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_servicelength]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_servicelength]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_statutoryredundancypay]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_statutoryredundancypay]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_uniquecode]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_uniquecode]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_username]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_username]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_wholemonthsbetweentwodates]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_wholemonthsbetweentwodates]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_wholeyearsbetweentwodates]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_wholeyearsbetweentwodates]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_workingdaysbetweentwodates]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_workingdaysbetweentwodates]
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfsys_justdate]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_justdate]

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_wholemonthsbetweentwodates] 
		(
			@date1 	datetime,
			@date2 	datetime
		)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer;
		
			-- Clean dates (trim time part)
			SET @date1 = DATEADD(D, 0, DATEDIFF(D, 0, @date1));
			SET @date2 = DATEADD(D, 0, DATEDIFF(D, 0, @date2));
		
			IF @date1 < @date2
			BEGIN
		
				-- Get the total number of months
				SET @result = DATEDIFF(mm, @date1, @date2);
		      
				-- See if the day field of pvParam2 < pvParam1 day field and if so - 1
				IF DAY(@date2) < DAY(@date1)
				BEGIN
					SET @result = @result -1;
				END
			END
			
			RETURN @result
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_wholeyearsbetweentwodates] (
		     @date1  datetime,
		     @date2  datetime )
		RETURNS integer 
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer = 0;
			
		    -- Get the number of whole years
		    SET @result = YEAR(@date2) - YEAR(@date1);
		
		    -- See if the date passed in months are greater than todays month
		    IF MONTH(@date1) > MONTH(@date2)
		    BEGIN
				SET @result = @result - 1;
		    END
		    
		    -- See if the months are equal and if they are test the day value
		    IF MONTH(@date1) = MONTH(@date2)
		    BEGIN
		        IF DAY(@date1) > DAY(@date2)
		            BEGIN
						SET @result = @result - 1;
		            END
		        END
		        
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_absencebetweentwodates] (
		     @date1		datetime,
		     @date2		datetime,
		     @type		nvarchar(MAX) )
		RETURNS integer 
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer = 0;
			
		    -- Get the number of whole years
		    SET @result = YEAR(@date2) - YEAR(@date1);
		
		    -- See if the date passed in months are greater than todays month
		    IF MONTH(@date1) > MONTH(@date2)
		    BEGIN
				SET @result = @result - 1;
		    END
		    
		    -- See if the months are equal and if they are test the day value
		    IF MONTH(@date1) = MONTH(@date2)
		    BEGIN
		        IF DAY(@date1) > DAY(@date2)
		            BEGIN
						SET @result = @result - 1;
		            END
		        END
		        
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_firstnamefromforenames] 
		(
			@forenames nvarchar(max)
		)
		RETURNS nvarchar(max)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result nvarchar(max);
		
			IF (LEN(@forenames) = 0 ) OR (@forenames IS null)
			BEGIN
				SET @result = '''';
			END
			ELSE
			BEGIN
				IF CHARINDEX('' '', @forenames) > 0
					SET @result = LEFT(@forenames, CHARINDEX('' '', @forenames));
				ELSE
					SET @result = @forenames;
			END
			
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord](
			@searchcolumn AS nvarchar(255),
			@searchexpression AS nvarchar(MAX),
			@returnfield AS nvarchar(255))
		RETURNS nvarchar(MAX)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result nvarchar(MAX);
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getfunctionparametertype]
			(@functionid integer, @parameterindex integer)
		RETURNS integer
		AS
		BEGIN
		
			DECLARE @result integer;
		
			SELECT @result = [parametertype] FROM ASRSysFunctionParameters
				WHERE @functionid = [functionID] AND @parameterindex = [parameterIndex];
		
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_initialsfromforenames] 
		(
			@forenames	varchar(8000),
			@padwithspace bit
		)
		RETURNS nvarchar(10)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result nvarchar(10) = '''';
			DECLARE @icounter integer = 1;
		
			IF LEN(@forenames) > 0 
			BEGIN
				SET @result = UPPER(left(@forenames,1));
		
				WHILE @icounter < LEN(@forenames)
				BEGIN
					IF SUBSTRING(@forenames, @icounter, 1) = '' ''
					BEGIN
						IF @padwithspace = 1
							SET @result = @result + '' '' + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
						ELSE
							SET @result = @result + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
					END
			
					SET @icounter = @icounter +1;
				END
		
				SET @result = @result + '' ''
			
			END
		
			RETURN @result
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isfieldpopulated](
			@inputcolumn as nvarchar(MAX))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit = 0;
			SELECT @result = (
				CASE 
					WHEN @inputcolumn IS NULL THEN 0 
					ELSE
						CASE
		--					WHEN LEN(convert(nvarchar(1),@inputcolumn)) = 0 THEN 0
							WHEN DATALENGTH(@inputcolumn) = 0 THEN 0
							ELSE 1
						END
					END);
		
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isovernightprocess] ()
		RETURNS bit 
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit = 0;
			
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_maternityexpectedreturndate] (
		     @id		integer)
		RETURNS datetime
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result datetime;
			
			SET @result = GETDATE()
		        
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_nicedate](
			@inputdate as datetime)
		RETURNS nvarchar(max)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result varchar(MAX) = '''';
			SELECT @result = CONVERT(nvarchar(2),DATEPART(day, @inputdate))
				+ '' '' + DATENAME(month, @inputdate) 
				+ '' '' + CONVERT(nvarchar(4),DATEPART(YYYY, @inputdate));
		
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_nicetime](
			@inputdate as datetime)
		RETURNS nvarchar(255)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result varchar(255) = '''';
		
			SELECT @result =convert(char(8), @inputdate, 108)
		
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_parentalleaveentitlement] (
		     @id		integer)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer;
			
			SET @result = 10;
		        
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_parentalleavetaken] (
		     @id		integer)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer;
			
			SET @result = 0;
		        
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_propercase](
			@text as nvarchar(max))
		RETURNS nvarchar(max)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @reset	bit = 1;
			DECLARE @result varchar(8000) = '''';
			DECLARE @i		int = 1;
			DECLARE @c		char(1);
		      
			WHILE (@i <= len(@text))
				SELECT @c= substring(@text,@i,1)
					, @result = @result + CASE WHEN @reset=1 THEN UPPER(@c) 
											   ELSE LOWER(@c) END
					, @reset = CASE WHEN @c LIKE ''[a-zA-Z]'' THEN 0
									ELSE 1
									END
					, @i = @i + 1;
		
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_servicelength] (
		     @startdate  datetime,
		     @leavingdate  datetime,
		     @period nvarchar(2))
		RETURNS integer 
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer;
			DECLARE @amount integer;
		
			-- If start date is in the future ignore
			IF @startdate > GETDATE()
				RETURN 0;
			
			-- Trim the leaving date
			IF @leavingdate IS NULL OR @leavingdate > GETDATE()
				SET @leavingdate = GETDATE();
		
		
			SET @amount = [dbo].[udfsys_wholeyearsbetweentwodates]
				(@startdate, @leavingdate);
		
			-- Years	
			IF @period = ''Y'' SET @result = @amount
			
			--Months
			ELSE IF @period = ''M''
				SET @result = [dbo].[udfsys_wholemonthsbetweentwodates]
					(@startdate, @leavingdate) - (@amount * 12);
			
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_statutoryredundancypay](
			@startdate AS datetime,
			@leavingdate AS datetime,
			@dateofbirth AS datetime,
			@weeklyrate AS numeric(10,2),
			@limit as numeric(10,2))
		RETURNS numeric(10,2)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result numeric(10,2);
			DECLARE @service_years integer;
			
			--/* First three parameters are compulsory, so return 0 and exit if they are not set */
			IF (@startdate IS null) OR (@leavingdate IS null) OR (@weeklyrate IS null)
			BEGIN
				SET @result = 0;
				RETURN @result;
			END
		
			-- Calculate service years
			SET @service_years = [dbo].[udfsys_wholeyearsbetweentwodates](@startdate, @leavingdate);
		
			SET @result = @service_years * @weeklyrate;
		
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_uniquecode](
			@prefix AS nvarchar(max),
			@coderoot AS numeric(10,2))
		RETURNS numeric(10,0)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result numeric(10,0);
		
			SET @result = 0;
		
			--SELECT @result = [maxcodesuffix] 
			--	FROM [dbo].[tb_uniquecodes]
			--	WHERE [codeprefix] = @prefix;
		
			-- Update existing value 
			/*
			You can''t run an execute or an update in a UDF, so will have to create an extended
			stored procedure which should be able to do it. Otherwise tack something into the end
			of the update trigger on the base table that calls this function. (code stub already in
			the admin module.
			*/
		
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_username]
			(@userid as integer)
		RETURNS varchar(255)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			RETURN SYSTEM_USER;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_workingdaysbetweentwodates]
		(
			@date1 	datetime,
			@date2 	datetime
		)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
			
			RETURN 0
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_justdate]
		(
			@date 	datetime
		)
		RETURNS datetime
		WITH SCHEMABINDING
		AS
		BEGIN
			RETURN DATEADD(D, 0, DATEDIFF(D, 0, @date));
		END';
	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 7 - Populate code generation tables'

	EXEC sp_executesql N'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[tbstat_componentcode]'') AND type in (N''U''))
		DROP TABLE [dbo].[tbstat_componentcode]'

	EXEC sp_executesql N'CREATE TABLE [dbo].[tbstat_componentcode](
			[objectid] [uniqueidentifier] NULL,
			[code] [nvarchar](max) NULL,
			[datatype] [int] NULL,
			[appendwildcard] [bit] NULL,
			[splitintocase] [bit] NULL,
			[name] [nvarchar](50) NULL,
			[aftercode] [nvarchar](50) NULL,
			[isoperator] [bit] NULL,
			[operatortype] [tinyint] NULL,
			[id] [int] NULL,
			[bypassvalidation] [bit] NULL
		) ON [PRIMARY]'

	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''4ec6c760-2157-492d-9161-24aa7c8a7b35'', N''AND'', NULL, 0, 0, N''And'', NULL, 1, 178, 5, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''ed72c364-c583-4356-b543-d99c188788be'', N''LIKE'', NULL, 1, 0, N''Begins with'', N''+ ''''%'''''', 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''8ef94e11-6693-422d-8099-bedee430083a'', N''+'', NULL, 0, 0, N''Concatenated with'', NULL, 1, 0, 17, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''57bc755b-61b5-41a0-92a0-321aab134b9c'', N''LIKE ''''%'''' +'', NULL, 1, 0, N''Is Contained Within'', NULL, 1, 177, 14, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a34f7387-91a1-40d6-b42f-f8032609cfd6'', N''/ NULLIF('', NULL, 0, 0, N''Divided by'', N'', 0)'', 1, 0, 4, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a9a1e24c-075c-475f-8157-2495a937ae95'', N''NOT LIKE'', NULL, 1, 0, N''Does not begin with'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''2dc26511-020f-4a3b-9839-d1616d358e52'', N''NOT LIKE ''''%'''' +'', NULL, 1, 0, N''Does not contain'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''644bbb16-5dad-4bff-bc8a-667630fd67b1'', N''NOT LIKE ''''%'''' +'', NULL, 0, 0, N''Does not end with'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''6e7c217a-99a6-4467-af8c-931596ce3491'', N''LIKE ''''%'''' + '', NULL, 0, 0, N''Ends with'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''d64f8151-2740-43bd-be5c-b7daaf29d5dc'', N''>'', NULL, 0, 0, N''Is after'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''bdc4b4bd-2557-4198-981b-ab26136ca0be'', N''>='', NULL, 0, 0, N''Is after OR equal to'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''dbeb4e24-6b36-4936-83d5-0d1016974d1f'', N''<'', NULL, 0, 0, N''Is before'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''baad571d-c3c0-42b0-96ea-a5f72cc881d7'', N''<='', NULL, 0, 0, N''Is before OR equal to'', NULL, 1, 177, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''d4521a9e-2974-49ef-849c-0d132aca93a0'', N''='', NULL, 0, 0, N''Is equal to'', NULL, 1, 177, 7, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''14b67bc6-ab84-4bf5-b20c-40c16e94a193'', N''>'', NULL, 0, 0, N''Is greater than'', NULL, 1, 177, 10, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''f54c9ae5-4790-403b-bb66-a026d67df26e'', N''>='', NULL, 0, 0, N''Is greater than OR equal to'', NULL, 1, 177, 12, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''11f18863-ecbd-4930-ab00-544a4fba5162'', N''<'', NULL, 0, 0, N''Is less than'', NULL, 1, 177, 9, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''14dd3e78-331a-47f6-81d2-5ce5df8c6935'', N''<='', NULL, 0, 0, N''Is less than OR equal to'', NULL, 1, 177, 11, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''3543f5d9-eef6-48c7-8aa4-934fe4202700'', N''<>'', NULL, 0, 0, N''Is NOT equal to'', NULL, 1, 177, 8, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''435776a4-6803-4a08-972b-c40480313ce8'', N''-'', NULL, 0, 0, N''Minus'', NULL, 1, 177, 2, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''b35d6bba-d45e-4ec4-bcd0-1d3d3e2d78fc'', N''%'', NULL, 0, 0, N''Modulus'', NULL, 1, 0, 16, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''68a326f9-ca7f-496f-b6e1-0d0f488ac7f6'', N''OR'', NULL, 0, 0, N''Or'', NULL, 1, 178, 6, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''6e51716a-4ac3-49dc-97a5-2bc417e38c2f'', N''+'', NULL, 0, 0, N''Plus'', NULL, 1, 0, 1, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''1acde45c-39a1-4a50-8526-aed3b8e6392b'', N''*'', NULL, 0, 0, N''Times by'', NULL, 1, 0, 3, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a0aefbd0-b295-4598-9432-d4f653eca1ac'', N''*'', NULL, 0, 0, N''To the power of'', NULL, 1, 0, 15, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''e6bd0161-786d-42a8-bdff-8400963e3e89'', N''[dbo].[udfsys_absencebetweentwodates] ({0}, {1}, {2})'', 2, 0, 0, N''Absence between Two Dates'', NULL, 0, 0, 47, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a4776d94-8917-4f5b-ad36-f4104b04e3e0'', N''[dbo].[udf_ASRFn_AbsenceDuration]({0}, {1}, {2}, {3}, @prm_ID)'', 2, 0, 0, N''Absence Duration'', NULL, 0, 0, 30, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''edfa5940-f5ba-47b5-bd93-8f19c35490b3'', N''DATEADD(DD, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, 0, 0, N''Add Days to Date'', NULL, 0, 0, 44, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''51a4dc3e-41a4-4b1a-8df8-8d9a1baed196'', N''DATEADD(MM, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, 0, 0, N''Add Months to Date'', NULL, 0, 0, 23, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''bc6a9215-696d-492c-8acb-95c99f440530'', N''DATEADD(YY, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, 0, 0, N''Add Years to Date'', NULL, 0, 0, 24, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''078108bf-77b2-42a3-b426-42126337f397'', N'''', 2, 0, 0, N''Bradford Factor'', NULL, 0, 0, 73, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''eb449e75-e061-4502-973b-5e3a3e39c2d2'', N'''', 2, 0, 0, N''Convert Character to Numeric'', NULL, 0, 0, 25, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''56b64c0d-84d9-4b15-9c9e-b1fdb42ea4d1'', N'''', 2, 0, 0, N''Convert Currency'', NULL, 0, 0, 51, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''88430aa0-f580-4157-8b2f-c73841cea211'', N''CONVERT(nvarchar(MAX), LEFT({0},{1}))'', 1, 0, 0, N''Convert Numeric to Character'', NULL, 0, 0, 3, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''98e87fe4-bb86-4382-bf53-40fa1275d677'', N''LOWER({0})'', 1, 0, 0, N''Convert to Lowercase'', NULL, 0, 0, 8, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''9e055f03-efe9-4c47-a528-85cd3c57c12a'', N''[dbo].[udfsys_propercase]({0})'', 1, 0, 0, N''Convert to Proper Case'', NULL, 0, 0, 12, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''59a5f6dd-8284-45a2-a68e-01e9f6d2e13e'', N''UPPER({0})'', 1, 0, 0, N''Convert to Uppercase'', NULL, 0, 0, 2, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''302dbbe5-d900-4547-8090-5de3dd3a4970'', N''SYSTEM_USER'', 1, 0, 0, N''Current User'', NULL, 0, 0, 17, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''8a4abce8-984e-4d4f-b1ca-aaef09e1c08d'', N''DATEPART(day, {0})'', 2, 0, 0, N''Day of Date'', NULL, 0, 0, 34, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''b41669c9-59d7-449f-be4f-6d4c6b809db9'', N''DATEPART(week, {0})+1'', 2, 0, 0, N''Day of the Week'', NULL, 0, 0, 28, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''24884a1c-fc85-4bba-8752-cb594c4607f2'', N''DATEDIFF(dd,{0}, {1})+1'', 2, 0, 0, N''Days between Two Dates'', NULL, 0, 0, 45, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''25033092-aa37-406d-ba0e-7b59b81c9b69'', N'''', 3, 0, 0, N''Does Record Exist'', NULL, 0, 0, 74, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a774b4f7-5792-41c5-99fb-301af38f0e68'', N''LEFT({0}, {1})'', 1, 0, 0, N''Extract Characters from the Left'', NULL, 0, 0, 6, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''0d948a6a-e6db-440f-b5fc-25ac323425ae'', N''RIGHT({0}, {1})'', 1, 0, 0, N''Extract Characters from the Right'', NULL, 0, 0, 13, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5c4e830d-6b52-481d-b94e-e6d65912cde2'', N''SUBSTRING({0}, {1}, {2})'', 1, 0, 0, N''Extract Part of a Character String'', NULL, 0, 0, 14, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''f61ea313-4866-4a29-a19f-e2d4fe3db23d'', N'''', 3, 0, 0, N''Field Changed between Two Dates'', NULL, 0, 0, 53, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''532861e4-23ac-474b-ae04-1a85724e7988'', N'''', 4, 0, 0, N''Field Last Change Date'', NULL, 0, 0, 52, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''4be2a715-f36b-4507-8090-9b1159de3aab'', N''DATEADD(dd, 1 - DATEPART(dd,{0}), {0})'', 2, 0, 0, N''First Day of Month'', NULL, 0, 0, 55, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''e3f98ac8-bfbf-4a98-8dd3-89f2830c1c95'', N''DATEADD(dd, 1 - DATEPART(dy, {0}), {0})'', 2, 0, 0, N''First Day of Year'', NULL, 0, 0, 57, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''263f4cc8-7c8d-4c5d-bdea-9e4ced21f078'', N''[dbo].[udfsys_firstnamefromforenames]({0})'', 1, 0, 0, N''First Name from Forenames'', NULL, 0, 0, 21, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''f14ebf8d-98e6-4e36-a1e1-35efd0023c55'', N''CASE WHEN {0} THEN {1} ELSE {2} END'', 0, 0, 0, N''If... Then... Else...'', NULL, 0, 0, 4, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5feedcc3-e731-46b0-b7fe-2027e1e9ded4'', N''[dbo].[udfsys_initialsfromforenames]({0},0)'', 1, 0, 0, N''Initials from Forenames'', NULL, 0, 0, 20, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''7d539e37-6d9f-44b3-a694-7db9638a2502'', N''CASE WHEN {0} < {1} THEN 0 ELSE CASE WHEN {0} > {2} THEN 0 ELSE 1 END END'', 0, 0, 0, N''Is Between'', NULL, 0, 0, 38, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''9171430a-6a23-48cb-bd11-4b025eaaceaf'', N''CASE WHEN {0} < {1} THEN 0 ELSE CASE WHEN {0} > {2} THEN 0 ELSE 1 END END'', 3, 0, 0, N''Is Date In Range'', NULL, 0, 0, NULL, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a9997816-add0-467f-999d-79ef30c2b713'', N''CASE [dbo].[udfsys_isfieldpopulated]({0}) WHEN 0 THEN 1 ELSE 0 END'', 3, 0, 0, N''Is Field Empty'', NULL, 0, 0, 16, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''8caf9f74-dee4-4618-8d59-e292847f202a'', N''CASE [dbo].[udfsys_isfieldpopulated]({0}) WHEN 1 THEN 1 ELSE 0 END'', 3, 0, 0, N''Is Field Populated'', NULL, 0, 0, 61, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''63d90dd1-1fb0-42a7-8135-83cb25293d7b'', N''dbo.[udfsys_isovernightprocess]() '', 3, 0, 0, N''Is Overnight Process'', NULL, 0, 0, 50, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''94127e4f-8046-4516-83a0-2062dd0ea2e6'', N'''', 3, 0, 0, N''Is Personnel That Current User Reports To'', NULL, 0, 0, 72, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''17d67659-4e60-40ee-bb72-763f4f85a645'', N'''', 3, 0, 0, N''Is Personnel That Reports To Current User'', NULL, 0, 0, 68, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''0a0d63a7-d926-4b8c-9f4e-2c3ae3d650ab'', N'''', 3, 0, 0, N''Is Post That Current User Reports To'', NULL, 0, 0, 70, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''cb6680b4-1940-435d-8144-bae2af8f37a1'', N'''', 3, 0, 0, N''Is Post That Reports To Current User'', NULL, 0, 0, 66, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''3e9cae1a-0948-481d-8d0b-9c13ca5d9373'', N''DATEADD(dd, -1, DATEADD(mm, 1, DATEADD(dd, 1 - DATEPART(dd, {0}), {0})))'', 4, 0, 0, N''Last Day of Month'', NULL, 0, 0, 56, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''bc2ce0fc-6e2c-43a2-86f9-0ed45cba129a'', N''DATEADD(dd, -1, DATEADD(yy, 1, DATEADD(dd, 1 - DATEPART(dy, {0}), {0})))'', 4, 0, 0, N''Last Day of Year'', NULL, 0, 0, 58, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''18babc7b-b84e-4ca9-9e10-c630bb004891'', N''LEN({0})'', 2, 0, 0, N''Length of Character Field'', NULL, 0, 0, 7, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''bba7fff7-bd75-4953-abd1-2f70418bbb80'', N''[dbo].[udfsys_maternityexpectedreturndate](@pid)'', 4, 0, 0, N''Maternity Expected Return Date'', NULL, 0, 0, 64, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5acc9ebe-af46-438e-9ebd-2741b42e26e0'', N''CASE WHEN {0} > {1} THEN {0} ELSE {1} END'', 1, 0, 0, N''Maximum'', NULL, 0, 0, 9, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''e03d8884-8835-425c-b268-1eec196917eb'', N''CASE WHEN {0} < {1} THEN {0} ELSE {1} END'', 1, 0, 0, N''Minimum'', NULL, 0, 0, 10, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''7d1376aa-5bdc-4844-9e5d-b3499b807639'', N''DATEPART(MM, {0})'', 2, 0, 0, N''Month of Date'', NULL, 0, 0, 33, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''fbff11aa-2aa9-43c9-b75f-5f2333ff880e'', N''DATENAME(weekday, {0})'', 1, 0, 0, N''Name of Day'', NULL, 0, 0, 60, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''d97bb954-d303-4eeb-90ce-1466287de905'', N''DATENAME(month, {0})'', 1, 0, 0, N''Name of Month'', NULL, 0, 0, 59, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''ccbcd03a-0c7e-47c7-b4bf-d6e8bd7963e8'', N''[dbo].[udfsys_nicedate]({0})'', 1, 0, 0, N''Nice Date'', NULL, 0, 0, 35, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''42a88b07-200f-4785-9c1f-e4b5a97a9001'', N''[dbo].[udfsys_nicetime]({0})'', 1, 0, 0, N''Nice Time'', NULL, 0, 0, 36, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''e59b8c9c-31d1-494f-b9c3-ca0a6a6aef1e'', N'''', 2, 0, 0, N''Number of Working Days per Week'', NULL, 0, 0, 29, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''84044568-fea7-48d5-ae8a-f8178b7ed927'', N''[dbo].[udfsys_parentalleaveentitlement](@pid)'', 2, 0, 0, N''Parental Leave Entitlement'', NULL, 0, 0, 62, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5278a126-c44e-41c5-9e7a-1c890c297d3f'', N''[dbo].[udfsys_parentalleavetaken](@pid)'', 2, 0, 0, N''Parental Leave Taken'', NULL, 0, 0, 63, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''ed3be9d9-28f1-4345-a8c8-ca9f0c18a3a2'', N''({0})'', 0, 0, 0, N''Parentheses'', NULL, 0, 0, 27, 1)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''06e67c1b-c376-4fc9-a260-e9a12022791f'', N'''', 2, 0, 0, N''Remaining Months since Whole Years'', NULL, 0, 0, 19, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''b86c77e6-e393-499e-9114-95a201a316d4'', N''LTRIM(RTRIM({0}))'', 1, 0, 0, N''Remove Leading and Trailing Spaces'', NULL, 0, 0, 5, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5c2244b5-ee8b-4f80-bc9e-defb9ba10b36'', N'''', 4, 0, 0, N''Round Date to Start of Nearest Month'', NULL, 0, 0, 37, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''07e1acb6-6943-4a92-956f-5df24aa2f3d2'', N''FLOOR({0})'', 2, 0, 0, N''Round Down to Nearest Whole Number'', NULL, 0, 0, 31, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''6c8c5ca0-ae52-46fc-9289-06e989c32d6d'', N''ROUND({0} / {1}, 0) * {1}'', 2, 0, 0, N''Round to Nearest Number'', NULL, 0, 0, 49, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''49161588-d050-4f0b-a0cd-3d9d6393f5f3'', N''CEILING({0})'', 2, 0, 0, N''Round Up to Nearest Whole Number'', NULL, 0, 0, 48, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''022a1f4c-b15b-411a-a49f-08ec4c3497e4'', N''CHARINDEX ({1}, {0}, 0) '', 1, 0, 0, N''Search for Character String'', NULL, 0, 0, 11, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''cfa37a8b-4d7b-4abc-ae80-7866219e4469'', N''[dbo].[udfsys_servicelength]({0}, {1}, ''''M'''')'', 2, 0, 0, N''Service Months'', NULL, 0, 0, 40, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''81847039-a90d-476c-88a5-c5e447d77701'', N''[dbo].[udfsys_servicelength]({0}, {1}, ''''Y'''')'', 2, 0, 0, N''Service Years'', NULL, 0, 0, 39, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''2bf404b7-970e-4fdb-9977-00d516a6cc84'', N''[dbo].[udfsys_statutoryredundancypay]({0}, {1}, {2}, {3}, {4})'', 2, 0, 0, N''Statutory Redundancy Pay'', NULL, 0, 0, 41, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''1cce61bf-ee36-4779-83b9-233885440437'', N''(DATEADD(D, 0, DATEDIFF(D, 0, GETDATE())))'', 4, 0, 0, N''System Date'', NULL, 0, 0, 1, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''1b77e32f-756b-4e97-94d2-f0b053b0baca'', N''SYSDATETIME()'', 1, 0, 0, N''System Time'', NULL, 0, 0, 15, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''a8974869-0964-40e9-bbbf-4ac6157bf07f'', N''[dbo].[udfsys_uniquecode] ({0}, {1})'', 0, 0, 0, N''Unique Code'', NULL, 0, 0, 43, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''09e7dfb0-3bc2-4db5-a596-9639eb3e77b5'', N''DATEDIFF (WW, {0}, {1}) - 1'', 2, 0, 0, N''Weekdays between Two Dates'', NULL, 0, 0, 22, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''fbccef52-27be-4ee4-8afa-d8228da2e952'', N''[dbo].[udfsys_wholemonthsbetweentwodates] ({0}, {1})'', 2, 0, 0, N''Whole Months between Two Dates'', NULL, 0, 0, 26, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''1b5082ad-36bb-4bf8-b859-22a1de8f8d2e'', N''[dbo].[udfsys_wholeyearsbetweentwodates] ({0}, {1})'', 2, 0, 0, N''Whole Years between Two Dates'', NULL, 0, 0, 54, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''97880cd2-c73d-4c7e-a4c4-971824b850e6'', N''[dbo].[udfsys_wholeyearsbetweentwodates] ({0}, GETDATE())'', 2, 0, 0, N''Whole Years until Current Date'', NULL, 0, 0, 18, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''f11ffb85-31bc-4b12-9b3d-e4464c868ca4'', N''[dbo].[udfsys_workingdaysbetweentwodates] ({0}, {1})'', 2, 0, 0, N''Working Days between Two Dates'', NULL, 0, 0, 46, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5c9a6256-ac11-456d-92fe-a5e2f5ba4c11'', N''DATEPART(YYYY, {0})'', 2, 0, 0, N''Year of Date'', NULL, 0, 0, 32, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5b636d9f-7589-46d4-bd6a-0e23aef81a51'', N''NOT'', 0, 0, 0, N''Not'', NULL, 1, 180, 13, 0)'
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [appendwildcard], [splitintocase], [name], [aftercode], [isoperator], [operatortype], [id], [bypassvalidation]) VALUES (N''5da8bb7e-f632-4ed0-b236-e042b88f3a1b'', N''[dbo].[udfsys_getfieldfromdatabaserecord] ({0}, {1}, {2})'', 0, 0, 0, N''Get field from database record'', NULL, 0, 0, 42, 0)'


/* ------------------------------------------------------------- */
PRINT 'Step 8 - Administration module stored procedures'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spadmin_getcomponentcode]')
			AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_getcomponentcode]


	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spadmin_getcomponentcode]
	AS
	BEGIN
		SELECT [id], [code], [name], ISNULL([datatype],0) AS [returntype]
			, [appendwildcard], [splitintocase]
			, [aftercode], [isoperator], [operatortype], [aftercode]
			, [bypassvalidation]
		FROM tbstat_componentcode WHERE [id] IS NOT NULL;
	END'
	EXECUTE sp_executeSQL @sSPCode;

	
/* ------------------------------------------------------------- */
/* ------------------------------------------------------------- */

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
--IF (@@ERROR <> 0) goto QuitWithRollback

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
    IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
    BEGIN
        SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END
    ELSE
    BEGIN
        SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END

    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '4.3')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '4.3.0')

delete from asrsyssystemsettings
where [Section] = 'ssintranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('ssintranet', 'minimum version', '4.3.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '4.2.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '4.2.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '4.2.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.3')


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
GRANT EXECUTE ON xp_StartMail TO public
GRANT EXECUTE ON xp_SendMail TO public
GRANT EXECUTE ON xp_LoginConfig TO public
GRANT EXECUTE ON xp_EnumGroups TO public'
--EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE ['+@DBName + ']
GRANT VIEW DEFINITION TO public'
EXEC sp_executesql @NVarCommand


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 0)

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.3 Of HR Pro'
