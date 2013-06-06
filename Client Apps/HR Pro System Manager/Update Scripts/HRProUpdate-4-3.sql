
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

/* ------------------------------------------------------------- */
PRINT 'Step 2 - Rename base user tables'

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'EXECUTE dbo.spASRTableToView ''' + TableName + ''', ''tbuser_' + LOWER(TableName) + ''';'
		FROM ASRSysTables;
	EXECUTE sp_executesql @NVarCommand;

/* ------------------------------------------------------------- */
PRINT 'Step X - Rename base system tables'

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysTables'', ''tbsys_tables'''
	EXECUTE (@NVarCommand);

	SET @NVarCommand = 'EXECUTE spASRTableToView ''ASRSysColumns'', ''tbsys_columns'''
	EXECUTE (@NVarCommand);


/* ------------------------------------------------------------- */
PRINT 'Step 4 - Add new calculation procedures'

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
values('database', 'refreshstoredprocedures', 1)

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.3 Of HR Pro'
