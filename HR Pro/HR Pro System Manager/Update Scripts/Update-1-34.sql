
/* -------------------------------------------------- */
/* Update the database from version 33 to version 34. */
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


/* Exit if the database is not version 33 or 34. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.33') and (@sDBVersion <> '1.34')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 4 - Adding ASR Unique SQL Server Object Checking Functionality.'


SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE lower(name) = 'asrsystempobjects'

IF @iRecCount = 0 
BEGIN
SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysTempObjects]
		(
		  [ID] [int] IDENTITY(1, 1) NOT NULL
		, [Name] [varchar](255) NOT NULL
		, [Type] [varchar](16) NOT NULL
		, [DateCreated] [datetime] NOT NULL
		, [Owner] [varchar](255) NOT NULL 
		) ON [PRIMARY]'

EXEC sp_executesql @NVarCommand
END

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRUniqueObjectName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRUniqueObjectName]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRUniqueObjectName(
	  				  @psUniqueObjectName sysname OUTPUT
					, @Prefix sysname
					, @Type int)
			AS
			BEGIN
				DECLARE   @NewObj 		as sysname
		 			, @Count 		as int
					, @sUserName		as sysname
					, @sCommandString	nvarchar(4000)	
			 		, @sParamDefinition	nvarchar(500)

				SELECT @sUserName = SYSTEM_USER
				SELECT @Count = 1
				SELECT @NewObj = @Prefix + CONVERT(varchar(100),@Count)

				WHILE (EXISTS (SELECT * FROM sysobjects WHERE id = object_id(@NewObj) AND sysstat & 0xf = @Type))
					OR (EXISTS (SELECT * FROM ASRSysTempObjects WHERE Name = @NewObj AND Type = @Type))
					BEGIN
						SELECT @Count = @Count + 1
	    					SELECT @NewObj = @Prefix + CONVERT(varchar(10),@Count)
	  				END

				INSERT INTO ASRSysTempObjects (Name, Type, DateCreated, Owner) VALUES (@NewObj, @Type, GETDATE(), @sUserName)

				SET @sCommandString = ''SELECT @psUniqueObjectName = '''''' + @NewObj + ''''''''

				SET @sParamDefinition = N''@psUniqueObjectName sysname output''

				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psUniqueObjectName output

			END'

EXEC sp_executesql @NVarCommand

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRDropUniqueObject]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRDropUniqueObject]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRDropUniqueObject(
	  			@psUniqueObjectName sysname,
	  			@Type integer)
			AS
			BEGIN
	
				DECLARE 	@sCommandString 	nvarchar(4000),
						@sCurrentUser		varchar(4000),
						@sDBName		varchar(4000)
						
				SET @sCurrentUser = SUSER_SNAME()
				
				SELECT @sDBName = master..sysdatabases.name 
  				FROM master..sysdatabases
	  				INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid
				WHERE master..sysprocesses.spid = @@spid

				IF (EXISTS (SELECT * FROM sysobjects WHERE name = @psUniqueObjectName))
					BEGIN
						IF @Type = 3 
						BEGIN
							SET @sCommandString = ''DROP TABLE ['' + @sCurrentUser + ''].['' + @psUniqueObjectName + '']''
						END

						IF @Type = 4
						BEGIN
							SET @sCommandString = ''DROP PROCEDURE ['' + @sCurrentUser + ''].['' + @psUniqueObjectName + '']''
						END 

						EXECUTE sp_executesql @sCommandString
	  				END
	
				IF (EXISTS (SELECT * FROM ASRSysTempObjects WHERE Name = @psUniqueObjectName AND Type = @Type AND Owner = @sCurrentUser))
					BEGIN
						SET @sCommandString = ''DELETE FROM ASRSysTempObjects WHERE Name = '''''' + @psUniqueObjectName 
												                        + '''''' AND Type = '' + convert(varchar(8000), @Type)
													           + '' AND Owner = '''''' + @sCurrentUser + ''''''''

						EXECUTE sp_executesql @sCommandString
	  				END
			END'

EXEC sp_executesql @NVarCommand

/* ---------------------------- */

PRINT 'Step 2 of 4 - Updating NiceTime function.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_NiceTime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_NiceTime]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_NiceTime 
(
	@psResult 	varchar(8000) OUTPUT,
	@psTimeString	varchar(8000) /* in the format hh:mm:ss (24 hour clock) */
)
AS
BEGIN
	/* Return the given time in the format hh:mm am/pm (12 hour clock) */
	select @psResult = 
	case 
		when len(ltrim(rtrim(@psTimeString))) = 0 then ''''
		else 
			case 
				when isdate(@psTimeString) = 0 then ''***''
				else (convert(varchar(2),((datepart(hour,convert(datetime, @psTimeString)) + 11) % 12) + 1)
					+ '':'' + right(''00'' + datename(minute, convert(datetime, @psTimeString)),2)
					+ case 
						when datepart(hour, convert(datetime, @psTimeString)) > 11 then '' pm''
						else '' am'' 
					end) 
			end 
	end
END'

EXEC sp_executesql @NVarCommand

/* Start of "Whole Years Between Two Dates" expression function information. */

PRINT 'Step 3 of 4 - Adding "Whole Years Between Two Dates" expression function information.'

SELECT @NVarCommand = 'DELETE FROM ASRSysFunctions WHERE functionID = 54'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'INSERT INTO ASRSysFunctions  (functionID, functionName, returnType, timeDependent, category, spName, nonStandard, runtime)
       			VALUES (54, ''Whole Years Between Two Dates'', 2, 0, ''Date/Time'', ''sp_ASRFn_WholeYearsBetweenTwoDates'', 0, 1)'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'DELETE FROM ASRSysFunctionParameters WHERE functionID = 54'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       			VALUES (54, 1, 4, ''<Start Date>'')'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'INSERT INTO ASRSysFunctionParameters  (functionID, parameterIndex, parameterType, parameterName)
       			VALUES (54, 2, 4, ''<End Date>'')'
EXEC sp_executesql @NVarCommand

/* End of "Whole Years Between Two Dates" expression function information. */

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 4 of 4 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.34')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '1.7.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.34')

update asrsyssystemsettings
set [SettingKey] = 'hrpro@hrpro.co.uk'
where [SettingKey] = 'hrpro@hrpro.com'


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
--Not required for v34
--delete from asrsyssystemsettings
--where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
--insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
--values('database', 'refreshstoredprocedures', 1)


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.34 Of HR Pro'
