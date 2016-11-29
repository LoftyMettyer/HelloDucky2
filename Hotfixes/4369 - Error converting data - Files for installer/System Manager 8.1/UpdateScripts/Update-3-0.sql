
/* ----------------------------------------------------- */
/* Update the database from version 2.20 to version 3.0 */
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
	@sSQLVersion nvarchar(20),
	@iSQLVersion numeric(3,1),
	@iTemp integer,
	@sTemp varchar(8000),
	@sTemp2 varchar(8000)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)


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

/* Exit if the database is not version 2.20 or 2.21 or 3.0. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '2.20') and (@sDBVersion <> '2.21') and (@sDBVersion <> '3.0')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 29 - Drop CMG/Centrefile Export'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_GetCMGFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASR_GetCMGFields]

/* ------------------------------------------------------------- */
PRINT 'Step 2 of 29 - Domain Security Reader'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDomainPolicy]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetDomainPolicy]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].spASRGetDomainPolicy
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

  DECLARE @objectToken int
  DECLARE @hResult int
  DECLARE @hResult2 int
  DECLARE @pserrormessage varchar(255)

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

  -- Initialise the variables
  SET @LockoutDuration = 0
  SET @lockoutThreshold  = 0
  SET @lockoutObservationWindow  = 0
  SET @maxPwdAge  = 0
  SET @minPwdAge  = 0
  SET @minPwdLength  = 0
  SET @pwdHistoryLength  = 0 
  SET @pwdProperties  = 0

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
--    EXEC sp_OADestroy @objectToken
    RETURN 2
  END

  --EXEC sp_OADestroy @objectToken

END'

	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 3 of 29 - Adding new columns to Workflow Element Items'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'LeftCoord'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [LeftCoord] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'TopCoord'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [TopCoord] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'Width'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [Width] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'Height'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [Height] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'BackColor'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [BackColor] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'ForeColor'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [ForeColor]  [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'PictureID'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [PictureID] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'PictureBorder'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [PictureBorder] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontName'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontName] [Varchar] (50) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontSize'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontSize] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontBold'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontBold] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontItalic'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontItalic] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontStrikeThru'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontStrikeThru] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		and name = 'FontUnderline'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [FontUnderline] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'Alignment'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [Alignment] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
		
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
			                       	[Alignment] = 0
														 WHERE [Alignment] IS NULL'
			EXEC sp_executesql @NVarCommand	
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'ZOrder'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       [ZOrder] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
		
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
			                       	[ZOrder] = 0
														 WHERE [ZOrder] IS NULL'
			EXEC sp_executesql @NVarCommand	
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
                                    and name = 'TabIndex'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
			                       			[TabIndex] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
			
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
			                                   [TabIndex] = 0
			                       WHERE [TabIndex] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
		                                    and name = 'BackStyle'
		 
		if @iRecCount = 0
		BEGIN
		            SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
		                                   			[BackStyle] [SmallInt] NULL'
		            EXEC sp_executesql @NVarCommand
		 
		            SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
		                                               [BackStyle] = 0
		                                   WHERE [BackStyle] IS NULL'
		            EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'BackColorEven'	 
		if @iRecCount = 0
		BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
	                         				[BackColorEven] [int] NULL'
			EXEC sp_executesql @NVarCommand
		
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [BackColorEven] = 16777215
														 WHERE [BackColorEven] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'BackColorOdd' 
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[BackColorOdd] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [BackColorOdd] = 16777215
														 WHERE [BackColorOdd] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'ColumnHeaders'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[ColumnHeaders] [bit] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [ColumnHeaders] = 1
														 WHERE [ColumnHeaders] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'ForeColorEven'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[ForeColorEven] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [ForeColorEven] = 0
														 WHERE [ForeColorEven] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'ForeColorOdd'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[ForeColorOdd] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [ForeColorOdd] = 0
														 WHERE [ForeColorOdd] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeaderBackColor'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[HeaderBackColor] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeaderBackColor] = -2147483633
														 WHERE [HeaderBackColor] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontName'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontName] [Varchar] (50) NULL'
			EXEC sp_executesql @NVarCommand
 			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeadFontName] = ''Tahoma''
														 WHERE [HeadFontName] IS NULL'
			EXEC sp_executesql @NVarCommand     
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontSize'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontSize] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeadFontSize] = 8
														 WHERE [HeadFontSize] IS NULL'
			EXEC sp_executesql @NVarCommand     
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontBold'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontBold] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeadFontBold] = 0
														 WHERE [HeadFontBold] IS NULL'
			EXEC sp_executesql @NVarCommand     
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontItalic'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontItalic] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
												 [HeadFontItalic] = 0
									 WHERE [HeadFontItalic] IS NULL'
			EXEC sp_executesql @NVarCommand     
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontStrikeThru'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontStrikeThru] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeadFontStrikeThru] = 0
														 WHERE [HeadFontStrikeThru] IS NULL'
			EXEC sp_executesql @NVarCommand     
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'HeadFontUnderline'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
														 [HeadFontUnderline] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [HeadFontUnderline] = 0
														 WHERE [HeadFontUnderline] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'Headlines'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[Headlines] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [Headlines] = 1
														 WHERE [Headlines] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
			and name = 'TableID'
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD
                           				[TableID] [int] NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElementItems SET
																				 [TableID] = 0
														 WHERE [TableID] IS NULL'
			EXEC sp_executesql @NVarCommand      
		END


/* ------------------------------------------------------------- */
PRINT 'Step 4 of 29 - Local View of SysProcesses'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGenerateSysProcesses]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGenerateSysProcesses]

	SELECT @NVarCommand = 'CREATE Procedure spASRGenerateSysProcesses
	AS
	BEGIN
		DECLARE @strDBName nvarchar(100)
		DECLARE @objectToken int
		DECLARE @hResult int
		DECLARE @hMessage varchar(255)
	
		SET @strDBName = DB_Name()
	
		  -- Create Server DLL object
		EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
		IF @hResult <> 0
		BEGIN
			IF EXISTS (SELECT Name FROM dbo.sysobjects WHERE id = object_id(N''[dbo].[ASRTempSysProcesses]'') and OBJECTPROPERTY(id, N''IsTable'') = 1) DROP TABLE [dbo].[ASRTempSysProcesses]
			SELECT * INTO dbo.ASRTempSysProcesses FROM master..sysprocesses
		END	
		ELSE
		BEGIN
			EXEC @hResult = sp_OAMethod @objectToken, ''GenerateProcesses'', @hMessage OUTPUT, @@SERVERNAME, @strDBName
		END
	END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 5 of 29 - Login Check'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRLockCheck]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRLockCheck]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRLockCheck AS
	BEGIN
	
		DECLARE @strDBName varchar(255)
		DECLARE @objectToken int
		DECLARE @hResult int
		DECLARE @hMessage varchar(255)
	
		  -- Create Server DLL object
		EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
		IF @hResult <> 0
		BEGIN
			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN master..sysprocesses syspro 
			ON asrsyslock.spid = syspro.spid and asrsyslock.login_time = syspro.login_time
			WHERE Priority = 2 or syspro.spid IS not null
			ORDER BY Priority
		END	
		ELSE
		BEGIN
			SET @strDBName = DB_Name()
			EXEC @hResult = sp_OAMethod @objectToken, ''GenerateProcesses'', @hMessage OUTPUT, @@SERVERNAME,  @strDBName

			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN dbo.ASRTempSysProcesses syspro 
			ON asrsyslock.spid = syspro.spid and asrsyslock.login_time = syspro.login_time
			WHERE Priority = 2 or syspro.spid IS not null
			ORDER BY Priority

		END

	END'

	EXEC sp_executesql @NVarCommand
	GRANT EXECUTE ON sp_ASRLockCheck TO asrSysGroup	

/* ------------------------------------------------------------- */
PRINT 'Step 6 of 29 - Domain List Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDomains]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetDomains]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRGetDomains]
		(@DomainString varchar(8000) OUTPUT)
	AS
	BEGIN
	
	  DECLARE @objectToken int
	  DECLARE @hResult int
	  DECLARE @hResult2 varchar(255)
	  DECLARE @pserrormessage varchar(255)
	
	  -- Create Server DLL object
	  EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
	  IF @hResult <> 0
	  BEGIN
	    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
	    SET @pserrormessage = ''HR Pro Server.dll not found''
	    RAISERROR (@pserrormessage,1,1)
	    --EXEC sp_OADestroy @objectToken
	    RETURN 1
	  END
	
	  -- Initialise the variables
	
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
	
	  --EXEC sp_OADestroy @objectToken
	
	END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 7 of 29 - Adding new columns to Workflow Instance Steps'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'UserEmail'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [UserEmail] [varchar] (8000) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'UserName'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [UserName] [varchar] (256) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'DecisionFlow'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [DecisionFlow] [smallint] NULL'
			EXEC sp_executesql @NVarCommand
		END

/* ------------------------------------------------------------- */
PRINT 'Step 8 of 29 - Modifying stored procedures for Workflow'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASREmailImmediate]

	SET @sTemp = 'CREATE PROCEDURE [dbo].spASREmailImmediate(@Username varchar(255)) AS
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
			SELECT QueueID, ASRSysEmailQueue.LinkID, RecordID, ASRSysEmailQueue.ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID, DateDue
				FROM ASRSysEmailQueue
				LEFT OUTER JOIN ASRSysEmailLinks ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
				WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
				AND (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
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
								EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, '''', @RecipTo OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
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
					END

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

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDecisionSucceedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetDecisionSucceedingWorkflowElements]

	SELECT @NVarCommand = 'CREATE PROCEDURE dbo.spASRGetDecisionSucceedingWorkflowElements
		(
			@piElementID		integer,
			@piValue		integer,
			@succeedingElements	cursor varying output
		)
		AS
		BEGIN
			/* Return the IDs of the workflow elements that succeed the given element.
			This ignores connection elements.
			NB. This does work for elements with multiple outbound flows. */
			DECLARE
				@iConnectorPairID	integer,
				@superCursor		cursor,
				@iTemp		integer
			
			CREATE TABLE #succeedingElements (elementID integer)
		
			/* Get the non-connector elements. */
			INSERT INTO #succeedingElements
			SELECT L.endElementID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND L.startOutboundFlowCode = @piValue
				AND E.type <> 8 -- 8 = Connector 1
		
			DECLARE succeedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.connectionPairID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND L.startOutboundFlowCode = @piValue
				AND E.type = 8 -- 8 = Connector 1
		
			OPEN succeedingConnectorsCursor
			FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC spASRGetSucceedingWorkflowElements @iConnectorPairID, @superCursor OUTPUT	
				
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
		
			/* Return the cursor of succeeding elements. */
			SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
				SELECT elementID 
				FROM #succeedingElements
			OPEN @succeedingElements
		
			DROP TABLE #succeedingElements
		END'

	EXEC sp_executesql @NVarCommand


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetWorkflowEmailMessage]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetWorkflowEmailMessage
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psMessage		varchar(8000)	OUTPUT
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
				@dtTempDate		datetime
		
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
			
			SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
			FROM ASRSysWorkflowInstances
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID
		
			DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT EI.caption,
				EI.itemType,
				EI.dbColumnID,
				EI.dbRecord,
				EI.wfFormIdentifier,
				EI.wfValueIdentifier
			FROM ASRSysWorkflowElementItems EI
			WHERE EI.elementID = @piElementID
			ORDER BY EI.ID
		
			OPEN itemCursor
			FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iItemType = 1
				BEGIN
					/* Database value. */
					SELECT @sTableName = ASRSysTables.tableName, 
						@sColumnName = ASRSysColumns.columnName, 
						@iSourceItemType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iDBColumnID
		
					IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID
		
					SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
						'' FROM '' + @sTableName +
						'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
					SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
					EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
		
					IF @sValue IS null SET @sValue = ''''
		

					/* Format dates */
					IF @iSourceItemType = 11
					BEGIN
						IF len(@sValue) = 0
						BEGIN
							SET @sValue = ''<null>''
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
						@iSourceItemType = ASRSysWorkflowElementItems.itemType
					FROM ASRSysWorkflowInstanceValues
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
					INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID
					WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
						AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
						AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
		
					IF @sValue IS null SET @sValue = ''''

					/* Format dates */
					IF @iSourceItemType = 7
					BEGIN
						IF len(@sValue) = 0 OR @sValue = ''null''
						BEGIN
							SET @sValue = ''<null>''
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

		
				FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
			END
			CLOSE itemCursor
			DEALLOCATE itemCursor
		
			/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
			CREATE TABLE #succeedingElements (elementID integer)
		
			EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
		
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
						EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @sQueryString OUTPUT, @piInstanceID, @iElementID
						IF @hResult <> 0 
						BEGIN
							SET @sQueryString = ''''
						END
					END
								
					IF LEN(@sQueryString) = 0 
					BEGIN
						SET @psMessage = @psMessage + CHAR(13) +
							@sCaption + '' - Error constructing the query string. Please contact your system administrator.''
					END
					ELSE
					BEGIN
						SET @psMessage = @psMessage + CHAR(13) +
							@sCaption + '' - '' + @sURL + ''/?'' + @sQueryString
					END
					
					FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
				END
				CLOSE elementCursor
		
				DEALLOCATE elementCursor
			END
		
			DROP TABLE #succeedingElements
		END'

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRInstantiateWorkflow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRInstantiateWorkflow]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRInstantiateWorkflow
		(
			@piWorkflowID		integer,			
			@piInstanceID		integer		OUTPUT,
			@psFormElements	varchar(8000)	OUTPUT
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
				@sUserGroupName sysname,
				@iUserGroupID integer
		
			SET @iInitiatorID = 0
			SET @psFormElements = ''''
		
			EXEC spASRIntGetActualUserDetails
				@sActualLoginName OUTPUT,
				@sUserGroupName OUTPUT,
				@iUserGroupID OUTPUT	
			
			SET @sSQL = ''spASRSysGetCurrentUserRecordID''
			IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
			BEGIN
				SET @hResult = 0
		
				EXEC @hResult = @sSQL 
					@iRecordID OUTPUT,
					@iRecordCount OUTPUT
			END
		
			IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO ASRSysWorkflowInstances (workflowID, initiatorID, status, userName)
			VALUES (@piWorkflowID, @iInitiatorID, 0, @sActualLoginName)
						
			SELECT @piInstanceID = MAX(id)
			FROM ASRSysWorkflowInstances
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT ASRSysWorkflowLinks.endElementID
						FROM ASRSysWorkflowLinks
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowLinks.startElementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
							AND ASRSysWorkflowElements.type = 0) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT ASRSysWorkflowLinks.endElementID
						FROM ASRSysWorkflowLinks
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowLinks.startElementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
							AND ASRSysWorkflowElements.type = 0) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END
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
					OR ASRSysWorkflowElementItems.itemType = 0)
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM ASRSysWorkflowInstanceSteps
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowInstanceSteps.status = 1
				AND ASRSysWorkflowElements.type = 2
		
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

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRSubmitWorkflowStep]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRSubmitWorkflowStep
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psFormInput1		varchar(8000),
			@psFormInput2		varchar(8000)
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
				@iTemp		integer
		
			/* Put the submitted form values into the ASRSysWorkflowInstanceValues table. */
			WHILE charindex(CHAR(9), @psFormInput1 + @psFormInput2) > 0
			BEGIN
				SET @iIndex1 = charindex(CHAR(9), @psFormInput1 + @psFormInput2)
				SET @iIndex2 = charindex(CHAR(9), @psFormInput1 + @psFormInput2, @iIndex1+1)
		
				SET @sID = replace(LEFT(@psFormInput1 + @psFormInput2, @iIndex1-1), '''''''', '''''''''''')
				SET @sValue = replace(SUBSTRING(@psFormInput1 + @psFormInput2, @iIndex1+1, @iIndex2-@iIndex1-1), '''''''', '''''''''''')
		
				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.value = @sValue
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND ASRSysWorkflowInstanceValues.identifier = (SELECT ASRSysWorkflowElementItems.identifier
						FROM ASRSysWorkflowElementItems
						WHERE ASRSysWorkflowElementItems.ID = convert(integer, @sID))
		
				IF @iIndex2 > len(@psFormInput1)
				BEGIN
					SET @iIndex2 = @iIndex2 - len(@psFormInput1)
					SET @psFormInput1 = ''''
					SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
				END
				ELSE
				BEGIN
					SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2)
				END
			END
		
			/* Get the type of the given element */
			SELECT @iElementType = E.type,
				@iEmailID = E.emailID,
				@iEmailRecord = E.emailRecord
			FROM ASRSysWorkflowElements E
			WHERE E.ID = @piElementID
		
			SET @hResult = 0
			SET @sTo = ''''			
		
			IF @iElementType = 3 -- Email element
			BEGIN
				/* Get the email recipient. */
				SET @sTo = ''''
				SET @iEmailRecordID = 0
		
				SET @sSQL = ''spASRSysEmailAddr''
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					/* Get the record ID required. */
					IF @iEmailRecord = 0
					BEGIN
						/* Initiator''s record. */
						SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID
						FROM ASRSysWorkflowInstances
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					END
		
					/* Get the recipient''s address. */
					EXEC @hResult = @sSQL @sTo OUTPUT, @iEmailID, @iEmailRecordID
				END
		
				/* Build the email message. */
				EXEC spASRGetWorkflowEmailMessage @piInstanceID, @piElementID, @sMessage OUTPUT
		
				/* Send the email. */
				INSERT ASRSysEmailQueue(
					RecordDesc,
					ColumnValue, 
					DateDue, 
					UserName, 
					[Immediate],
					RecalculateRecordDesc, 
					RepTo,
					MsgText,
					Subject)
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@sTo,
					@sMessage,
					''HR Pro Workflow'')
			END'	

SET @sTemp2 = '
		
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
						ELSE ''''
					END
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iElementType = 4 -- Decision element
				BEGIN
					/* Get the ID of the elements that precede the Decision element. */
					CREATE TABLE #precedingElements (elementID integer)
		
					EXEC spASRGetPrecedingWorkflowElements @piElementID, @superCursor OUTPUT
			
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
				
					SELECT @iValue = convert(integer, IV.value)
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
					WHERE IV.elementID = @iPreviousElementID
						AND IV.instanceid = @piInstanceID
						AND E.ID = @piElementID
				
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
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate()
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT #succeedingElements2.elementID 
							FROM #succeedingElements2)
		
					DROP TABLE #succeedingElements2
				END
				ELSE
				BEGIN
					CREATE TABLE #succeedingElements (elementID integer)
		
					EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 1,
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
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
		
					DROP TABLE #succeedingElements
				END
			
				/* Set activated Web Forms to be ''pending'' (to be done by the user) */
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

				IF @iElementType = 3 -- Email element
				BEGIN
					exec spASREmailImmediate ''HR Pro Workflow''
				END
			END
		END'

	EXEC (@sTemp + @sTemp2)



/* ------------------------------------------------------------- */
PRINT 'Step 9 of 29 - Adding new columns to Accord Transaction Table'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransactions')
	and name = 'EmployeeName'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransactions ADD
		                       [EmployeeName] [varchar] (255) NULL,
		                       [DepartmentCode] [varchar] (255) NULL,
		                       [DepartmentName] [varchar] (255) NULL,
		                       [PayrollCode] [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferFieldDefinitions')
	and name = 'IsEmployeeName'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferFieldDefinitions ADD
		                       [IsEmployeeName] [bit] NULL,
		                       [IsDepartmentCode] [bit] NULL,
		                       [IsDepartmentName] [bit] NULL,
		                       [IsPayrollCode] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		-- Employee (0.74)
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (78,0,0,''12 Months Rolling Sick Days'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (79,0,0,''Current Period Sick Days'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (80,0,0,''Car Engine Size'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (81,0,0,''Employment Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (82,0,0,''Car User Category'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (83,0,0,''OSP Contract Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (84,0,0,''Student Loan'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (85,0,0,''Salary Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (86,0,0,''Use Spinal Points Flag'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, IsPayrollCode) VALUES (87,0,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (88,0,0,''Starter Form Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (89,0,0,''Starter Form Status'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand


		-- Employee (0.76)
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsDepartmentCode = 1, Mandatory = 1, AlwaysTransfer = 1 WHERE TransferFieldID = 36 AND TransferTypeID = 0'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (180,0,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (181,0,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- Salary
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (19,1,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (20,1,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (21,1,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (22,1,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- Allowances
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,2,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,2,1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,2,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,2,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- Loans
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (19,3,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (20,3,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (21,3,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (22,3,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		--Deductions
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,4,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,4,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,4,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,4,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- SSP
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (5, ''SSP'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,5,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,5,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,5,1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,5,1,''Start Session'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,5,1,''End Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,5,1,''End Session'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,5,0,''Reference'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,5,0,''Nominal Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,5,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,5,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,5,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,5,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,5,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,5,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,5,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,5,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,5,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,5,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,5,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,5,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,5,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- SMP
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (6, ''SMP'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,6,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,6,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,6,1,''Expected Week of Confinement'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,6,1,''Start Date'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,6,1,''End Date'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,6,0,''Reference'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,6,0,''Nominal Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,6,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,6,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,6,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,6,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,6,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,6,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,6,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,6,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,6,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,6,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,6,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,6,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,6,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- SPP
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (7, ''SPP'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,7,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,7,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,7,1,''Expected Week of Confinement'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,7,1,''Baby Date of Birth'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,7,1,''SPP Start Date'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,7,1,''SPP Weeks Number'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,7,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,7,0,''Reference'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,7,0,''Nominal Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,7,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,7,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,7,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,7,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,7,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,7,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,7,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,7,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,7,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (18,7,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (19,7,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (20,7,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (21,7,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- SAP
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (8, ''SAP'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,8,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,8,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,8,1,''SAP Start Date'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,8,0,''SAP End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,8,0,''Reference'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,8,0,''Nominal Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,8,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,8,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,8,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,8,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,8,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,8,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,8,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,8,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,8,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (15,8,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (16,8,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (17,8,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (18,8,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- Notes
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (9, ''Notes'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,9,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,9,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,9,0,''Action Flag'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,9,0,''Due Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,9,0,''Note Category'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,9,0,''Note Subject'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,9,0,''Note Text'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (7,9,1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (8,9,1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (9,9,1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (10,9,1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand

		-- Custom
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID) VALUES (10, ''Custom'' ,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,10,0,''Field 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,10,0,''Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,10,0,''Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,10,0,''Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,10,0,''Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,10,0,''Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,10,0,''Field 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,10,0,''Field 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,10,0,''Field 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,10,0,''Field 10'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,10,0,''Field 11'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,10,0,''Field 12'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,10,0,''Field 13'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,10,0,''Field 14'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,10,0,''Field 15'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,10,0,''Field 16'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,10,0,''Field 17'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,10,0,''Field 18'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (18,10,0,''Field 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (19,10,0,''Field 20'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (20,10,0,''Field 21'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (21,10,0,''Field 22'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (22,10,0,''Field 23'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (23,10,0,''Field 24'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (24,10,0,''Field 25'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (25,10,0,''Field 26'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (26,10,0,''Field 27'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (27,10,0,''Field 28'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (28,10,0,''Field 29'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (29,10,0,''Field 30'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Visibility of definition types
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferTypes')
	and name = 'IsVisible'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes ADD
		                       [IsVisible] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferTypes SET IsVisible=1 WHERE TransferTypeID IN(0,1,2,3,4)'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferTypes SET IsVisible=0 WHERE TransferTypeID IN(5,6,7,8,9,10)'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferFieldDefinitions')
	and name = 'GroupBy'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferFieldDefinitions ADD
		                       [GroupBy] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET GroupBy = 0'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET GroupBy = 1 WHERE TransferTypeID = 0 AND TransferFieldID IN (9,10,11,12,13,14)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET GroupBy = 2 WHERE TransferTypeID = 0 AND TransferFieldID IN (25,26,27,28,29)'
		EXEC sp_executesql @NVarCommand
	END

	-- Employee (1.00)
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions
	WHERE TransferFieldID = 182 AND TransferTypeID = 0

	IF @iRecCount = 0
	BEGIN	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (182,0,0,''Tax Code Effective From Date'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (183,0,0,''NI Letter Effective From Date'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (184,0,0,''Employee Category Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (185,0,0,''Nominal Costs Account Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (186,0,0,''Costs Code 1 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (187,0,0,''Costs Code 2 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (188,0,0,''Costs Code 3 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (189,0,0,''Costs Code 4 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (190,0,0,''Costs Code 5 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (191,0,0,''Costs Code 6 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (192,0,0,''Costs Code 7 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (193,0,0,''Costs Code 8 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName, GroupBy) VALUES (194,0,0,''Costs Code 9 Name'',0,0,2,0,0,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET GroupBy = 0 WHERE GroupBy IS NULL'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 10 of 29 - Updating Tidy Up Windows Orphans stored procedure.'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRDeleteInvalidLogins]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRDeleteInvalidLogins]

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRDeleteInvalidLogins
	(
	    @pstrDomainName varchar(100)
	)
	AS
	
		DECLARE @cursLogins cursor
		DECLARE @loginName nvarchar(200)

		-- Are we privileged enough to run this script
		IF IS_SRVROLEMEMBER(''securityadmin'') = 0 and IS_SRVROLEMEMBER(''sysadmin'') = 0
		BEGIN
			RETURN 0
		END

		-- Lets get the invalid accounts into a swish little cursor
		SET @cursLogins = CURSOR LOCAL FAST_FORWARD FOR
		SELECT loginName from master.dbo.syslogins
			WHERE isntname = 1
			AND loginname like @pstrDomainName + ''\%''
			AND ((sid <> SUSER_SID(loginname) and SUSER_SID(loginname) is not null)	OR SUSER_SID(loginname) is null)

		-- Now lets get rid of the invalid accounts
		OPEN @cursLogins
		FETCH NEXT FROM @cursLogins INTO @LoginName
		WHILE (@@fetch_status = 0)
		BEGIN
			EXEC sp_revokelogin @LoginName
			FETCH NEXT FROM @cursLogins INTO @LoginName
		END
	
		-- Tidy Up
		CLOSE @cursLogins 	
		DEALLOCATE @cursLogins

		RETURN 0'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 11 of 29 - Adding new columns to Workflow Element object'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormBGColor'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormBGColor] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormBGImageID'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormBGImageID] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormBGImageLocation'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormBGImageLocation] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END
		
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontName'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontName] [Varchar] (50) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontSize'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontSize] [SmallInt] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontBold'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontBold] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontItalic'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontItalic] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontStrikeThru'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontStrikeThru] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormDefaultFontUnderline'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormDefaultFontUnderline] [Bit] NULL'
			EXEC sp_executesql @NVarCommand
		END
		
		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormHeight'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormHeight] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
		and name = 'WebFormWidth'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD
			                       [WebFormWidth] [Int] NULL'
			EXEC sp_executesql @NVarCommand
		END


/* ------------------------------------------------------------- */
PRINT 'Step 12 of 29 - Altering Self-service Intranet URL column'

		SELECT @NVarCommand = 'ALTER TABLE [ASRSysSSIntranetLinks] 
															ALTER COLUMN [url] VARCHAR(500) NULL'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 29 - Updating Email Queue'

	select @iRecCount = length from syscolumns
	where id = object_id('ASRSysEmailQueue')
	and name like 'ColumnValue'

	IF @iRecCount < 8000
	BEGIN

		--Need to clear the stats before the column size can be altered
		DECLARE @Name varchar(8000)
	
		DECLARE Stats_Cursor
		CURSOR LOCAL FAST_FORWARD FOR 
		select [Name] from sysindexes
		where status > 0 and id in (select id from sysobjects where name like 'ASRSysEmailQueue')
		ORDER BY name
	
		OPEN Stats_Cursor
		FETCH NEXT FROM Stats_Cursor INTO @Name
	
		WHILE (@@fetch_status = 0)
		BEGIN
			SELECT @NVarCommand = 'DROP STATISTICS ASRSysEmailQueue.['+@Name+']'
			EXEC sp_executesql @NVarCommand
			FETCH NEXT FROM Stats_Cursor INTO @Name
		END
		CLOSE Stats_Cursor
		DEALLOCATE Stats_Cursor

		PRINT 'PLEASE NOTE: WARNINGS REGARDING MAXIMUM ROW SIZE CAN BE IGNORED'
		SELECT @NVarCommand = 
		'ALTER TABLE  ASRSysEmailQueue
		 ALTER COLUMN [ColumnValue] [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 14 of 29 - Domain User Reader'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWindowsUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetWindowsUsers]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRGetWindowsUsers]
		(@DomainName varchar(200),
		@UserString varchar(8000) OUTPUT)
	AS
	BEGIN
	
	  DECLARE @objectToken int
	  DECLARE @hResult int
	  DECLARE @hResult2 varchar(255)
	  DECLARE @pserrormessage varchar(255)
	
	  -- Create Server DLL object
	  EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
	  IF @hResult <> 0
	  BEGIN
	    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
	    SET @pserrormessage = ''HR Pro Server.dll not found''
	    RAISERROR (@pserrormessage,1,1)
	    --EXEC sp_OADestroy @objectToken
	    RETURN 1
	  END
	
	  -- Initialise the variables
	
	  -- Populate the variables
	  EXEC @hResult = sp_OAMethod @objectToken, ''GetUsers'', @UserString OUTPUT, @DomainName

	  IF @hResult <> 0 
	  BEGIN
	    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
	    SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
	    RAISERROR (@pserrormessage,2,1)
	    --EXEC sp_OADestroy @objectToken
	    RETURN 2
	  END
	
	  --EXEC sp_OADestroy @objectToken
	
	END'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 15 of 29 - System permissions stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRSystemPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRSystemPermission]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[sp_ASRSystemPermission]
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
			AND not (usg.name = ''db_owner'')

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

	SELECT @NVarCommand = '	GRANT EXECUTE ON sp_ASRSystemPermission TO asrSysGroup'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 16 of 29 - Server DLL Version stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetServerDLLVersion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetServerDLLVersion]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRGetServerDLLVersion]
	(
	@strVersion varchar(255) OUTPUT
	)
	AS
	BEGIN
	DECLARE @objectToken int
	DECLARE @hResult int

	  -- Create Server DLL object
	EXEC @hResult = sp_OACreate ''vbpHRProServer.clsGeneral'', @objectToken OUTPUT
	IF @hResult = 0
		EXEC @hResult = sp_OAMethod @objectToken, ''GetVersion'', @strVersion OUTPUT
	ELSE
		SET @strVersion = ''0.0.0''
		
	--EXEC sp_OADestroy @objectToken
	
	END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 17 of 29 - System Permission Printouts - Amend Accord Transfers'

	SET @NVarCommand = 'UPDATE ASRSysPermissionItems
	SET itemKey = ''VIEWARCHIVE'' WHERE  (itemID = 147)'
	EXEC sp_executesql @NVarCommand

	SET @NVarCommand = 'UPDATE ASRSysPermissionItems
	SET Description = ''View Transfers''
	WHERE  (itemID = 145)'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 18 of 29 - Database Compatability level'

	SELECT @sSQLVersion = substring(@@version,charindex('-',@@version)+2,1)

	IF @sSQLVersion = '8' EXEC sp_dbcmptlevel @DBName, 80
	ELSE
		IF @sSQLVersion = '9'
		BEGIN
			EXEC sp_dbcmptlevel @DBName, 90
			SELECT @NVarCommand = 'GRANT VIEW DEFINITION TO [ASRSysGroup]'
			EXEC sp_executesql @NVarCommand			
		END


/* ------------------------------------------------------------- */
PRINT 'Step 19 of 29 - Tidy up orphaned users stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRTidyUpNonASRUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRTidyUpNonASRUsers]

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRTidyUpNonASRUsers
	AS
		SET NOCOUNT ON
		DECLARE @iGroupID int
	
		SELECT @iGroupID = uid FROM sysusers WHERE isSQLRole = 1 AND name = ''ASRSysGroup''
	
		SELECT * FROM sysusers WHERE uid NOT IN (SELECT uid FROM sysmembers
		INNER JOIN sysusers ON sysusers.uid = sysmembers.memberuid
		WHERE groupuid = @iGroupID)
		AND IsSQLRole = 0 AND NOT (name = ''dbo'' OR name= ''guest'' OR name=''sys'' OR name=''INFORMATION_SCHEMA'')

	RETURN '
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 20 of 29 - Get current users stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetCurrentUsers]

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRGetCurrentUsers
	AS
	BEGIN
		SET NOCOUNT ON
	
		SELECT DISTINCT hostname, loginame, program_name, hostprocess
	    FROM ASRTempSysProcesses
	    WHERE program_name like ''HR Pro%'' 
	    AND dbid in ( 
	                   SELECT dbid FROM master..sysdatabases
	                   WHERE name = DB_NAME())
	     ORDER BY loginame

	END'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 21 of 29 - Adding new column to Diary Events'

	/* Add new digit group columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysDiaryEvents')
	and name = 'LinkID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysDiaryEvents ADD 
						[LinkID] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysDiaryEvents Set LinkID = 0
								WHERE ColumnID = 0'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 22 of 29 - Creating Index for Diary Events'

	SELECT @iRecCount = count(id) FROM sysindexes
	where name = 'ASRSysDiaryEventsIndex1'

	if @iRecCount > 0
		DROP INDEX [ASRSysDiaryEvents].[ASRSysDiaryEventsIndex1]

	SELECT @NVarCommand = 'CREATE INDEX [ASRSysDiaryEventsIndex1]
				ON [dbo].[ASRSysDiaryEvents] ([LinkID], [RowID])
				WITH FILLFACTOR = 80'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 23 of 29 - Removing obsolete Diary procedures'

	DECLARE	@sObjectName varchar(2000)
						
	DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
	SELECT name FROM sysobjects
	WHERE name like 'sp_ASRDiary[_]%' or name like 'sp_ASRDiaryRebuild%'

	OPEN tempObjects
	FETCH NEXT FROM tempObjects INTO @sObjectName
	WHILE (@@fetch_status <> -1)
	BEGIN
		EXEC ('DROP PROCEDURE dbo.[' + @sObjectName + ']')
		FETCH NEXT FROM tempObjects INTO @sObjectName
	END
	CLOSE tempObjects
	DEALLOCATE tempObjects


/* ------------------------------------------------------------- */
PRINT 'Step 24 of 29 - Refresh Security Icons'

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 2
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(2,'Batch Jobs','',10,'BATCHJOBS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 2
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x00000100010010100000010008006805000016000000280000001000000020000000010008000000000000010000000000000000000000010000000100000000000018181800212121006B2929006B313100733131007B4A4200844A4A00845252008C5252008C5A5A00945A5A009C636300946B63009C6B63009C6B6B009C736B009C737300A56B6B00A5737300A57B7300A57B7B00AD7B7B00AD7B840094949400A5848400AD848400A58C8400AD8C8400B5848400B58C8C00AD949400B5949400BD949400B59C9400BD9C9C00BDA59C00A5A5A500BDADA500C6949C00C69C9C00C6A5A500CEA5A500C6ADA500C6ADAD00CEADAD00D6ADAD00CEB5B500D6B5B500DEB5B500D6BDB500D6BDBD00C6C6C600CECECE00DEC6C600DECECE00D6D6D600DED6D600DEDEDE00E7CEC600E7CED600F7D6D600F7DEDE00F7DEE700EFE7E700EFEFEF00F7E7E700FFE7E700FFEFEF00FFF7F700FFF7FF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF000000000000200F070D1C260000000000000000230C13131D16130A0B000000000000210F1E333F3C3D3127120F000000002D0F2139FF45003E431F210F140000001315420025FF4545440029270B0000320E2CFFFFFFFFFF4545434331131B002D094041FFFFFFFFFF46453E3D1A0D002806FF0038FFFF1834FF4500371D0800300740FFFFFFFF350134FF463F161100360C2BFF38FFFFFF350134FF33132200001A0F400041FF3AFF3502401E0F0000003B0F1940FFFF0038FF40210F240000000F040F0F2C40FF402F150F05040000000F2A04170C0707090F13042E0400000000100F003630042D33001A1A00000000000000000007040300000000000000F81F0000E00F0000C00700008003000080030000000100000001000000010000000100000001000080030000800300008003000080030000C8270000FC7F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 21
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(21,'Calculations','',10,'CALCULATIONS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 21
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AB5661466666666666666666666666666B5FFF0FFF0FFF0FFF0FFBCBC07079266B5FFF00EFF0EFF0EFF0EF4F3F00EEF66B5FFF1FFF1FFF1FFF1FFF1FFF2FFEF66B5FFF20EFF0EFF0EFF0EFF0E070EEF66EFFFF2F2F2F1F1F1F0F0BCBC0707EF66EFFF6D6DECECEDF7EFEFF0BCBC070766EFFF0AFFFFFFF4F4F4EDF1F0BC070766EFFF0A0A0A0A0A0A0A0EF2F1F0BC0766EFFFFFFFFFFFFFFFFFFFFFFFFFFFFF66BBEFEFEFEFEFEFEFB5B5B5B5B5B5B5B50A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFFFF0000FFFF0000FFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFF0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 7
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(7,'Global Add','',10,'GLOBALADD')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 7
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AF2BCEFEFEFBCF20A0A0A0A0A0A0A0A07B4ADAD8B8B6D92F00A0A0A0A0A0ABBAD868BA7ADADAD66ECF00A0A0A0A09A6A6A66B65A6A686AD6692F20A0AF3ADA6C7AC48278AC7A6868B6DBC0A0A09A6ACACAC480206ACC76B6C8BEF0A0AB4ACACAC814E06CD8A4827498BEF0A0AB3ACAC4E70B2D3D3B20202486CEF0A799A9959999979DAD4D30648486CBC0A1A58E579A0589ABBDBD381020290F20AA0A0C3A0C3A099C2BAB3B26B48070A0A5279A0F61A7958C15671D3ACBB0A0A0AA0A0C31AC3A099975671D4090A0A0A0A9A58A079A0589ABA9DBBF30A0A0A0A0A799AA052A09A990A0A0A0A0A0A0A0A0AFFFF0000F80F0000F0070000E0030000C00100008001000080010000800100008001000000010000000100000003000000070000000F0000001F000001FF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 8
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(8,'Global Update','',10,'GLOBALUPDATE')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 8
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A9292F70A0A0A0A0A0A0A0A0A0A0A0AF2F7FF7444BCF20A0A0A0A0A0A0A0A07B4F79979744492F00A0A0A0A0A0ABBAD868B1C9A797444ECF00A0A0A0A09A6A6A66B651C9A79744492F20A0AF3ADA6C7AC48278A1C9A797444BC0A0A09A6ACACAC480206AC1C9A79EC440A0AB4ACACAC814E06CD8A481C07746F740AB3ACAC4E70B2D3D3B202026F746F930ABAAC71345596DAD4D306484893930A0A19B255569797BBDBD381020290F20A0A0A9D3497C1C2C2BAB3B26B48070A0A0A0A0A9797C1C2C15671D3ACBB0A0A0A0A0A0A0ABBBABB975671D4090A0A0A0A0A0A0A0A0A19BBBA9DBBF30A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFC7F0000F80F0000F0070000E0030000C0010000800100008001000080000000800000008001000080010000C0030000E0070000F00F0000F81F0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 9
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(9,'Global Delete','',10,'GLOBALDELETE')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 9
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A93250A0A0A0A0A750A0A0A0A0AF2BCEF26264CF20A0A46200A0A0A0A07B4ADAD47262692F053254D0A0A0ABBAD868BA7AD4D47757A26170A0A0A09A6A6A66B65A6A64D4D474DF20A0AF3ADA6C7AC48278AC7A64D4D6DBC0A0A09A6ACACAC48020693E3E3E3E3990A0AB4ACACAC814E069AE3E3484975E31A0AB3ACAC4E70B294E3E30202486C9A750ABAAC713455969475D30648486CBC0A0A19B255569797BBDBD381020290F20A0A0A9D3497C1C2C2BAB3B26B48070A0A0A0A0A9797C1C2C15671D3ACBB0A0A0A0A0A0A0ABBBABB975671D4090A0A0A0A0A0A0A0A0A19BBBA9DBBF30A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFF3E0000F80C0000F0000000E0010000C0010000800100008001000080000000800000008001000080010000C0030000E0070000F00F0000F81F0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 14
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(14,'Filters','',10,'FILTERS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 14
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000400100000000000000000000000000000000000000000000EF966300EF926300EF8E5A00E78A5200DE794A00D6713900CE653100C6552100BD511800BD4D1800FFD7C600FFDBC600FFD3BD00FFCFB500F7C3AD00F7BA9C00F7B29400F7AE8C00F7A27B00B54D1800FFDBCE00F7CBB500F7BEA500F7AA8C00F7A67B00EF9A7300EF8A5200EF865200F7BE9C00F7A68400F79E7300EF966B00EF8A5A00EF824A00F7B69400EF9A6B00F7C7AD00E77D4A00D6693100C6592100B5511800FFFFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000002728290A0000000000000000000000002225221400000000000000000000000022251C140000000000000000000000002625030A0000000000000000000000002625020A0000000000000000000000002617200A000000000000000000000022251320030A000000000000000000220D23132402030A000000000000001C0B1D121E1F200221140000000000031516171118191A01021B14000000020B0C0B0B0D0E0F1011121302140000010101020303040506070809090A000000000000000000000000000000000000000000000000000000000000000000FFFF0000FFFF0000FC3F0000FC3F0000FC3F0000FC3F0000FC3F0000FC3F0000F81F0000F00F0000E0070000C00300008001000080010000FFFF0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 3
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(3,'Cross Tabs','',10,'CROSSTABS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 3
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AF7666666666666666666666666660A0AF7FF19F0BC090909BBB5B5B5B5660A0AF7FFB58B8BF3F0AE19DDBC09B5660A0AF7FFB5FF8BF4AEAE8B8B09BCB5660A0AF7FFB5FFAEFFF1AE19B5AEF0B5660A0AF7FFEFFFAEFFFFF4F3F0AEF0BB660A0AB5FFEFFFB4FFFFFFF4AEAEAE09660A0AB5FFEFEFEFFFFFFFFF19AEBCBC660A0AB5FFFFFFFFFFFFFFFFFFF4F3F0660A0AEFFFB5AE8BFFEFCFAEAE8B8B19660A0ABBFFEFFFAEFFEFFFFFFFFF8BF3660A0ABBFFEFEFEFFFEFEFEFB5B5B5F3660A0ABBFFFFFFFFFFFFFFFFFFFFFFFF660A0ABBBBBBBBEFEFEFEFB5B5B5B5B5F70A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFFFF00008001000080010000800100008001000080010000800100008001000080010000800100008001000080010000800100008001000080010000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 12
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(12,'Custom Reports','',10,'CUSTOMREPORTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 12
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000AF76666666666666666666C0A0A0A0A0AB5FFF3F2F1F1F0BC09B56C0A0A0A0A0AEFFFBC0707F4BBBBB5B56C0A0A0A0A0ABBFFFFFFFFFFF4F419BB6C0A0A0A0A511AFFBCBCEDAEAEAEAE6C666666666C9A2A1AFFFFEDFFF4F4F3BCEDBC09B56C7A162ABCBCB5FFBC0707F191BBB5B56C7A7A162A07B5FFFFFFFFF4F7F419BB6C0A7A7A164B07B4B4ADADADAD07BBBB6C4B4A4B7A164A07D6D6D5D5ADFFF4076CE5E59A9A751644B4B4B4ADAD0707076C59C37A165107BCFFFFFFFFFFFFFFF46C0AA07A75164BEFD6D6D5B4B4ADADADAD0A0AC37A75164AEF09DCDCD6D6D5D5AD0A0A0AE57A751612EF0909090909D6D60A0A0A0A79797474990A0A0A0A0A0A0A800F0000800F0000800F0000800F0000000000000000000000000000000000008000000000000000000000000000000080000000C0000000E0000000F07F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 30
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(30,'Envelope & Label Templates','',10,'LABELDEFINITION')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 30
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A9292F70A0A0A0A0A0A0A0A0A0A0A0A0AF7FF74440A0A0A0A0A0A0A0A0A0A0A0AF7997974440A0A0A0A0A0A0A0A0A0A0A0A1C9A7974440A0A0A0A0A0A0A0A0A0A0A0A1C9A7974440A0A0A0AB5666666666666661C9A797444660A0AB5FF1919F1F0BC09091C9A79EC440A0AEFFFFFFFFFFFF3F319191C07746F740AEFFFEFB5B5F7F7F3F3F3196F746F930AEFFFFFFFFFFFFFF4F3F3F31993930A0AEFFFEFEFB5B5F7F7F3F3F31909660A0AEFFFFFFFFFFFFFFFFFFF464509660A0AEFFFFFFFFFFFFFFFFFFF4D46BC660A0AEFFFFFFFFFFFFFFFFFFFFFFFF3660A0AEF19191919F1F0BC090909BBB5F70A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFC7F0000FC3F0000FC1F0000FE0F0000FF07000080010000800100008000000080000000800100008001000080010000800100008001000080010000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 29
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(29,'Envelopes & Labels','',10,'LABELS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 29
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AB5666666666666666666666666660A0AB5FF1919F1F0BC090909BBB5B5660A0AEFFFFFFFFFFFF3F31919DDF0B5660A0AEFFFEFB5B5F7F7F3F3F319DDBB660A0AEFFFFFFFFFFFFFF4F3F3F31909660A0AEFFFEFEFB5B5F7F7F3F3F31909660A0AEFFFFFFFFFFFFFFFFFFF464509660A0AEFFFFFFFFFFFFFFFFFFF4D46BC660A0AEFFFFFFFFFFFFFFFFFFFFFFFF3660A0AEF19191919F1F0BC090909BBB5F70A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFFFF0000FFFF0000FFFF000080010000800100008001000080010000800100008001000080010000800100008001000080010000FFFF0000FFFF0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 11
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(11,'Mail Merge','',10,'MAILMERGE')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 11
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A040A0AED6D6C6C6C6C6C6C660A04040404040AEFF3BC070707EFF76C0A0A0A0A040A0AEFFFFFFFF4F1F0EF6C0A0A0A0A0A0A0AEFFFFFFFF4F3F1EF6CB5666666666666BBFFFFFFFFF4F3EF6CF7FFF1F1F0BCBC07FFFFFFFFFFF3076CB5FFFFFFFFF4F407FFFFFFFFFFF4076CEFFFBC07FFBBBB09FFFFFFFFFFBCED6CBBFFFFFFFFFFFFBCFFFFFFFFF3916CAEBBFFBCBCFF070709FFFFFFFFF3ED910ABBFFFFFFFFFFFF09BCBCBCBCBCB50A0ABBFFBCBCFFBCBCFF0707F36C0A0A0A0ABBFFFFFFFFFFFFFFFFFFFF6C0A0A0A0ADCD6D6D6D5B4B4ADADADADAD0A0A0A0ADC090909DCDCD6D6D5D5D5AD0A0A0A0ADCD6D6D6D6D5B4B4B4B4ADAD0A0A0A0AF600000082000000F6000000FE00000000000000000000000000000000000000000000000001000000030000000F0000000F0000000F0000000F0000000F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 4
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(4,'Data Transfer','',10,'DATATRANSFER')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 4
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEC6CEA6CAE6CEA6CECFFFFFFFFFFFFFF91EEB5BB09BBB5EE91FFFFFFFFFFFFFFECFF1919191919FFECFFFFFFFFFFFFFF91FFF3191919F3FF91FFFFFFFFFFFFFFEDFFF4F3F3F3F4FFEDFFFFFFFFFFFFFFEDFFF4F3F4F3F4FFED1EFFFFFFFFFFFFF7FFFFFFFFFFFFFF1E461EFFFFFFFFFFF7F007070707071E46466F45FFFFFFFFF7F2F31919191E46176F1694E9FFFFFFB5FFFFFFFF1E46176F1694BDBD46FFFFB507EFF7EB1E01456F169446464646FFEFF5F0BCBCBCF01E1694BD46FFFFFFFFEFFFF5F4F4F4F50194BDBD46FFFFFFFFEFFFFFFFFFFFFF4594BD9446FFFFFFFF0707EFEFEFEFEF6F6F6F6F16FFFFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 10
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(10,'Import','',10,'IMPORT')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 10
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A1C0A0A0A0A0AEC6CEA6CAE6CEA6CEC1E1E0A0A0A0A0A91BCB5BB09BBB5BC1E46010A0A0A0A0AECFF19191919191E4617451E01456F0A91FFF31919191E46176F6F1694946F0AEDFFF4F3F31E46466F161694BDBD6F0AEDFFF4F3F4F31E6F169494BDBD946F0AF7FFFFFFFFFFFF4594BD464646466F0AF7F00707070707F0E9BD460A0A0A0A0AF7F2F3191919F3F2F746460A0A0A0A0AB5FFFFFFFFFFFFFFB50A460A0A0A0A0AB507EFF7EFF7EF07B50A0A0A0A0A0A0AEFFFF0BCBCBCF0FFEF0A0A0A0A0A0A0AEFFFFFF4F4F4FFFFEF0A0A0A0A0A0A0AEFFFFFFFFFFFFFFFEF0A0A0A0A0A0A0A0707EFEFEFEFEF07070A0A0A0A0A0AFFEF0000800F0000800F00008000000080000000800000008000000080000000800F0000800F0000802F0000803F0000803F0000803F0000803F0000803F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 6
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(6,'Export','',10,'EXPORT')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 6
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A1C0A0A0A0A0A0AEC6CEA6CAE6CEA6CEC1E1E0A0A0A0A0A91BCB5BB09BBB5BC9101461E0A0A0A0AECFF1919196F45011E4517461E0A0A0A91FFF319196F9494166F6F17461E0A0AEDFFF4F3F36FBDBD9416166F46461E0AEDFFF4F3F46F94BDBD9494166F1E0A0AF7FFFFFFFF1646464646BD94450A0A0AF7F00707070707F0F746BDE90A0A0A0AF7F2F3191919F3F2F746460A0A0A0A0AB5FFFFFFFFFFFFFFB5460A0A0A0A0A0AB507EFF7EFF7EF07B50A0A0A0A0A0A0AEFFFF0BCBCBCF0FFEF0A0A0A0A0A0A0AEFFFFFF4F4F4FFFFEF0A0A0A0A0A0A0AEFFFFFFFFFFFFFFFEF0A0A0A0A0A0A0A0707EFEFEFEFEF07070A0A0A0A0A0AFFDF0000800F000080070000800300008001000080000000800100008003000080070000800F0000801F0000803F0000803F0000803F0000803F0000803F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 34
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(34,'Record Profile','',10,'RECORDPROFILE')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 34
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000604830006D553D007F654D00DD1B0100836B530084695100897159008C735B008D755D0091786000977F6700AB8C7300AF907700BA9D8A00B0A09000C0A89000C0A8A000C0B0A000CFB7AF00D0B0A000D0B8A000D9B9A700DABAA800DCBBA800D0B8B000D0C0B000E0C0B000E0C8C000E0D0C000E0D0D000E0D8D000F0D8D000F0E0D000F0E0E000F0E8E000FDF5EE00F0F0F000F7F5F400F9F3F100FFF0F000FFF7F000FBF7F600FFF8F000FFF8FF00FEFEFE000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF000000000507070B07070A07070A010000000000022424292B292B2B2924010000000F0101010101010101010129010000000FFF16181818181818180129010000000FFFFFFF2523211E1C180129010000000FFF0DFF1B1B1B201D180129010000000FFFFFFFFF2923221F1801290100000010FF0DFF1B1B1B23201601290100000011FFFFFFFFFF2923232301290100000012FF0DFF1B1B1B25232601290100000014FFFFFFFFFFFF29252601290100000015FF0DFF1B290404042601290100000019FFFFFFFF2904FF04260129070000001AFF0DFF1BFF04040426030A0A0000001BFFFFFFFFFFFFFF26260600000000001B1B1B1B1B1A13130E0D0D00000000E0030000E0030000800300008003000080030000800300008003000080030000800300008003000080030000800300008003000080030000800F0000800F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 13
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(13,'Standard Reports','',10,'STANDARDREPORTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 13
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000AF76666666666666666666C0A0A0A0A0AB5FFF3F2F1F1F0BC09B56C0A0A0A0A0AEFFFBC0707F4BBBBB5B56C0A0A0A0A0ABBFFFFFFFFFFF4F419BB6C0A0A0A0A511AFFBCBCEDAEAEAEAE6C666666666C9A2A1AFFFFEDFFF4F4F3BCEDBC09B56C7A162ABCBCB5FFBC0707F191BBB5B56C7A7A162A07B5FFFFFFFFF4F7F419BB6CFF7A7A164B07B4B4ADADADAD07BBBB6C4B4A4B7A164A07D6D6D5D5ADFFF4076CE5E59A9A751644B4B4B4ADAD0707076C59C37A165107BCFFFFFFFFFFFFFFF46C0AA07A75164BEFD6D6D5B4B4ADADADAD0A0AC37A75164AEF09DCDCD6D6D5D5AD0A0A0AE57A751612EF0909090909D6D60A0A0A0A79797474990A0A0A0A0A0A0A800F0000800F0000800F0000800F0000000000000000000000000000000000000000000000000000000000000000000080000000C0000000E0000000F07F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 15
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(15,'Picklists','',10,'PICKLISTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 15
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AF76D6C12666666666666666666666666F71919F0090909090909090909090966B5FF072B2B244BF71919191919190966B5F34C7552242B24F409B5B5B5B50966EFF44C744C4C2B4BF419191919190966EFFFF24B92ED4BF1FFDDB5B5B5B50766BBFF6D11EFEC6DFFFFFFF4F4F419076607FF1212F2BCECF2FFDCDCDCD6D5076607FFEBEBF7EDECF2FFFFFFFFFFFFBC1207FFF2EBECEBBCFFFFFFFFFFFFFFBC6C07FFFFFFFFFFFFFFFFFFFFFFFFFFFFAE070707070707BBEFEFB5B5B5F7F7F7F70A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AFFFF0000FFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFF0000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 5
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(5,'Diary','',10,'DIARY')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 5
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0AF4BCBCF30A0A0A0A0A0A0A0A0A0A0A0ABCED91EFBCF40A0A0A0A0A0A0A0A0AF407F3B591EDEFF0F40A0A0A0A0A0A0ABCBCF1BCBCB5AE92EFF0F40A0A0A0AF30709B509F209BBB5AEF707F00A0A0ABCBCF109BBBBBBF3BBBBB5AEF7070AF3BCF0B509F3BBB5EFEFBC09B591AE0A0919F0BC09BBBBFF166F93EFBBB5AE1909DD09DDF3BB0907941A746EBBAE0709090909090909F3079A759393B5910AF0BC09090909090909F20793BBAEBC0A0A0ABC0709D6D5D5D60909BCF7EC0A0A0A0A0A0A07B5B5B4B4B4B4DDAEF00A0A0A0A0A0A0AF3F091B4B4B4CFB50A0A0A0A0A0A0A0A0A0AF1BCB4B4ADF30A0A0A0A0A0A0A0A0A0A0A0AF0EF090A0A0A0AF0FF0000F03F0000E00F0000E0030000C0010000C00000008000000080000000000000000001000000010000C0030000F0030000F8070000FE070000FF8F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 24
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(24,'Calendar Reports','',10,'CALENDARREPORTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 24
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00F7666666126666666666666666666666B5BBB5B409B5B409B5B4BBB5B4B5B566B5FFFFB5FFFFB5FFFFB5F4F4B5190766B5FFFFB5FFFFB5FFFFB5F4F4B5190766EFBBB5B4B5B5B4B5B525252525B5B566EFFFFFB5FFFFB5FFFF4DFFFF25F40766EFFFFFB5FFFFB5FFFF75FFFF25F4BC66BBBBB5B4B5B5B4B5B594754D25B5B566BBFFFFB5FFFFB5FFFFB5FFFFB5FFF112BBFFFFB5FFFFB5FFFFB5FFFFB5FFFF66D6D6D6D6D6D5B4B4B4B4B4B4ADADADADD6F4F4F4F4F4F4DCD6D6D6D6D6D6D5ADD6D6D6B5B5B4B4B4B4B4B4CFADADADADF7FFEDFFECFFECFFECFFEBFF6DFFEAF00A11F111F111F2110711F011F411F2F70A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000080000000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 18
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(18,'Email Queue','',10,'EMAIL')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 18
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000314A5A003152630039526300315A6B00395A6B00315A730031637B0039637B00425A6B00426373004A6373004A6B7B00296B8C00316B8C00318CBD004A8CAD00739CAD0073A5AD007BA5B5007BADBD0039A5D600429CC60052B5DE005AB5DE0063B5D60073B5D60063BDE7005AC6EF0063C6E7006BC6EF0063CEEF0063CEF7006BCEF70073CEF70073D6F70073DEF7007BDEF70073D6FF007BE7F7007BE7FF0084ADBD0084B5C6008CBDCE0094BDCE0094C6D60084CEE70084DEF70084E7F7008CE7FF0084EFFF0094E7F7008CF7FF008CFFFF0094F7FF009CF7FF0094FFFF009CFFFF00A5EFFF00B5EFF700A5F7FF00ADF7FF00A5FFFF00B5F7FF00C6F7FF00C6FFFF00CEFFFF00D6FFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF000000000000000000000000000000000000000000000000000000000000000000130D0E0707060804050203010000000012433C31252322221E201B0300000000141A3E36383435343832100303010000143B193524150F2332161B051B0300002A433D1F171D1D161C272609100303012A4333182F39372F1628210A1B051B032C3F1D303E39393931161E0B260910032B2E404342434243413A160C210A1B052D2D2D2B2B2A2A29291313111E0B260900002B2E404342434243413A160C210A00002D2D2D2B2B2A2A29291313111E0B000000002B2E404342434243413A160C000000002D2D2D2B2B2A2A292913131100000000000000000000000000000000FFFF0000FFFF0000000F0000000F000000030000000300000000000000000000000000000000000000000000C0000000C0000000F0000000F0000000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 40
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(40,'Outlook Calendar Queue','',10,'OUTLOOKQUEUE')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 40
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0AF4BCBCF30A0A0A0A0A0A0A0A0A0A0A0ABCED91EFBCF40A0A0A0A0A0A0A0A0AF407F3B591EDEFF0F40A0A0A0A0A0A0ABCBCF1BCBCB5AE92EFF0F40A0A0A0AF30709B509F209BBB5AEF707F00A0A0ABCBCF109BBBBBBF3BBBBB5AEF7070AF3BCF0B509F3BBB5EFEFBC09B591AE0A0919F0BC09BBBBFF166F93EFBBB5AE1909DD09DDF3BB0907941A746EBBAE0709090909090909F3079A759393B5910AF0BC09090909090909F20793BBAEBC0A0A0ABC0709D6D5D5D60909BCF7EC0A0A0A0A0A0A07B5B5B4B4B4B4DDAEF00A0A0A0A0A0A0AF3F091B4B4B4CFB50A0A0A0A0A0A0A0A0A0AF1BCB4B4ADF30A0A0A0A0A0A0A0A0A0A0A0AF0EF090A0A0A0AF0FF0000F03F0000E00F0000E0030000C0010000C00000008000000080000000000000000001000000010000C0030000F0030000F8070000FE070000FF8F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 19
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(19,'Data Manager Intranet','',10,'INTRANET')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 19
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0AF1BCF7FFF7F7F7F7F7F7F7F70A0A0ABCEDAEF7FFFFFFF4F319F0BCF70A0A07ADADADF7FFFFFFFFF4F319F0F70ABBADA7868BF7FFFFBBAECFB5F1F1F7F1ADA6A6A665B5FF09B44F6CADED19F7BAA6A6C7AC8AB5FFD56C71B36C71F3F7ADA6ACACAC6BEFFF912E91B47171F3F7ACACACACAC4EBBFF987898B49098F4F7ACACAC067070BBFFFFC20897BBFFF7F7ACAC702E4F9007FFFFFFFFFFFFF76666B3AC4F55565607FFFFFFFFFFFFB5BC66D5065556979709FFFFFFFFFFFFB566ED19555678BB08BC09BBBAB4B491B407F00A0856789EC2C2785690D3AC06BBF20A0A0A089D98BB98563571D3B3BB0A0A0A0A0A0AF3BBBB9D9D7796DBF10A0A0A0AF0000000E0000000C00000008000000000000000000000000000000000000000000000000000000000000000000000000000000080010000C0070000E00F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 20
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(20,'CMG & Centrefile','',10,'CMG')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 20
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A1C0A0A0A0A0A0AEC6CEA6CAE6CEA6CEC1E1E0A0A0A0A0A91BCB5BB09BBB5BC9101461E0A0A0A0AECFF1919196F45011E4517461E0A0A0A91FFF319196F9494166F6F17461E0A0AEDFFF4F3F36FBDBD9416166F46461E0AEDFFF4F3F46F94BDBD9494166F1E0A0AF7FFFFFFFF1646464646BD94450A0A0AF7F00707070707F0F746BDE90A0A0A0AF7F2F3191919F3F2F746460A0A0A0A0AB5FFFFFFFFFFFFFFB5460A0A0A0A0A0AB507EFF7EFF7EF07B50A0A0A0A0A0A0AEFFFF0BCBCBCF0FFEF0A0A0A0A0A0A0AEFFFFFF4F4F4FFFFEF0A0A0A0A0A0A0AEFFFFFFFFFFFFFFFEF0A0A0A0A0A0A0A0707EFEFEFEFEF07070A0A0A0A0A0AFFDF0000800F000080070000800300008001000080000000800100008003000080070000800F0000801F0000803F0000803F0000803F0000803F0000803F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 23
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(23,'Match Reports','',10,'MATCHREPORTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 23
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x00000100010010100000010008006805000016000000280000001000000020000000010008000000000000010000000000000000000000010000000100000000000084000000FF00000084820000FFFF000000FFFF008181810084828400878787008E8E8E009292920094949400989898009D9D9D00A2A2A200ABABAB00ADADAD00BDBDBD00C2C2C200C6C3C600CFCFCF00D6D6D600D9D9D900DCDCDC00E2E2E200E4E4E400EAEAEA00EEEEEE00F2F2F200F5F5F500F8F8F800FEFEFE0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00FFFFFFFFFFFFFF1EFF1EFFFFFFFFFFFFFFFFFFFFFF1E1A181516181A1EFFFFFFFFFFFF1E1A18140F0E0E0F14181AFFFFFFFFFF1A160F00000000090C0F161AFFFFFF1E18000008FF08FF00000A0F181E1E1A1800FF08FFFFFF08FFFF000A12151A160F0008FFFF00FFFF00000009090C18000012FFFF00FF0000FFFF120000080004030012FFFF00FFFFFF120005000003000403000000000000000005000101030300040300151A1E1E0005000102010303030000161AFFFF1E00000202010203030300161AFFFFFFFF1E0002020201030300141AFFFFFFFFFFFF000202020203001118FF1EFFFFFFFFFF1E0002020203001A1E1EFFFFFFFFFFFFFF1E00020200000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 35
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(35,'Email Addresses','',10,'EMAILADDRESSES')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 35
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000090909001401010001011C00101010001717170018181800050D3400231E3200202020002424240020282000282828002038200030283000303030003B3B3B003C3A3A0030403000305030003068300001074300121A4C001D305C004C3A5B00304F7700404040004545450040484000505050005757570050585000605040006058500070585000406040004078400040785000706050006058600060606000656565006068600068686800707070007777770070787000797979007F7F7F00D66C3000D86C3000DC6E3400DC723600DC743A00DC793F00E1672B00ED6A2B00F867210080704000806050008068500090706000907860008078700090787000E97063004088300000FF00004080400056845F00409050005090500040A050005A9C66007098700050A8700050B0700070A0700070B07000B09850009080600090807000A0807000A0887000B0907000F6874B00FF814E00FF845200FF895200FF8E5A00FF8D6500FC896E00FF986A00FF967B00D6A07B00FFA57600FFA279007038A000205581003D5F86003F5F8A00575E8000635F82004E6B9000436D9F004878A8005A7DA70090309000FF00FF0070A08000638FBE007284B9005796CB006594C5006994C8006898CE00739AC9006DA0D00068A2D60073A8D60076B4D6007DB7D70000FFFF0061FFFF0080808000808880008D8D8D0095959500909890009D9D9D00A0888000A0908000B0908000A5959400A0989000B098900080B0800080B89000A0A09000B4B296008B8EB500A0A0A000A0A8A000A8A8A800B0B0B000B6B6B600B0B8B000B9B9B900C0988000D0A08000C0A89000D0A09000D0A89000D0B09000E7BA8D00FFB18100FFB58500E0A89000E0B09000FFB89600C0A8A000C0B0A000C1B2A700C0B8A000C5B8AB00D0B0A000DBB7A100D0B8A000DDBBA600D6B3B300D0B8B000E7B7A100E0B8A000E5BFA80080E0A00080E8A00090E8B00090F0B000FFC59E00C9C4B300C9C8B300D0C0B000E0C0A000E6C0A800FFD1A900E0C0B000F1CFB600F6D2B300FFDEB800FFE3BE0087B1D70082BCD700C7BACD00C0C0C000C7C7C700CCCCCC00C0D8C000D2C0D100D0D8D000E0D0C000EEDACD00F0D0C000F0D6C000F6DAC400F9DCC400E0D0D000E0D8D000F0D8D000C0E8D000FFE5C500FFE9C700FBE0CA00FCE1CB00FFEFCF00F0E0D000FFE0D000FDE7D500FDE8D800D8CBE000D8E9EC00E4E4E400F0E8E000FFE8E000FFEFEB00FFF0E000F0F0F000FFF0F000FFF8F000FFFFF000FFF8FF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00B4474414130D0B8B00000000000000004AAEAFAE464C451200000000000000006DB0AFAE464C4912000000000000000088D0B188C4464B230000000000000000B3894EE1E54D24923C3B3B3D3D3E5152DB8A1D1D282590823C2283A3B5CEA552A4008D8E7D1F3F873E8453503D3D4F260904C6C18D7D2AA05296B5AA9853CF3A110F92C6C18D2EA55299C7DDE2847C210E0E1C271AC17CA153C9ACD7E3844348A10E2C807C1A86A754ACE1E3E2944342C8853F0E2E40B9D694C9B9E3E2953D2000A6CDC7E5E3CFAC94E1E5E3E0986C6B00A8C798E5C9C7D698C7E3E2DE976C6100AB9D9EE5E2C9B9C9B7D6DECF98873D00D9D79EB6B7CBD4D4CCBAACAC98000000FF000000FF000000FF000000FF0000000000000000000000000000000000000000000000000000000000000000000080000000800000008000000080030000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 36
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(36,'Email Groups','',10,'EMAILGROUPS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 36
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000090909001401010001011C00101010001717170018181800050D3400231E3200202020002424240020282000282828002038200030283000303030003B3B3B003C3A3A0030403000305030003068300001074300121A4C001D305C004C3A5B00304F7700404040004545450040484000505050005757570050585000605040006058500070585000406040004078400040785000706050006058600060606000656565006068600068686800707070007777770070787000797979007F7F7F00D66C3000D86C3000DC6E3400DC723600DC743A00DC793F00E1672B00ED6A2B00F867210080704000806050008068500090706000907860008078700090787000E97063004088300000FF00004080400056845F00409050005090500040A050005A9C66007098700050A8700050B0700070A0700070B07000B09850009080600090807000A0807000A0887000B0907000F6874B00FF814E00FF845200FF895200FF8E5A00FF8D6500FC896E00FF986A00FF967B00D6A07B00FFA57600FFA279007038A000205581003D5F86003F5F8A00575E8000635F82004E6B9000436D9F004878A8005A7DA70090309000FF00FF0070A08000638FBE007284B9005796CB006594C5006994C8006898CE00739AC9006DA0D00068A2D60073A8D60076B4D6007DB7D70000FFFF0061FFFF0080808000808880008D8D8D0095959500909890009D9D9D00A0888000A0908000B0908000A5959400A0989000B098900080B0800080B89000A0A09000B4B296008B8EB500A0A0A000A0A8A000A8A8A800B0B0B000B6B6B600B0B8B000B9B9B900C0988000D0A08000C0A89000D0A09000D0A89000D0B09000E7BA8D00FFB18100FFB58500E0A89000E0B09000FFB89600C0A8A000C0B0A000C1B2A700C0B8A000C5B8AB00D0B0A000DBB7A100D0B8A000DDBBA600D6B3B300D0B8B000E7B7A100E0B8A000E5BFA80080E0A00080E8A00090E8B00090F0B000FFC59E00C9C4B300C9C8B300D0C0B000E0C0A000E6C0A800FFD1A900E0C0B000F1CFB600F6D2B300FFDEB800FFE3BE0087B1D70082BCD700C7BACD00C0C0C000C7C7C700CCCCCC00C0D8C000D2C0D100D0D8D000E0D0C000EEDACD00F0D0C000F0D6C000F6DAC400F9DCC400E0D0D000E0D8D000F0D8D000C0E8D000FFE5C500FFE9C700FBE0CA00FCE1CB00FFEFCF00F0E0D000FFE0D000FDE7D500FDE8D800D8CBE000D8E9EC00E4E4E400F0E8E000FFE8E000FFEFEB00FFF0E000F0F0F000FFF0F000FFF8F000FFFFF000FFF8FF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00B4474414130D0B8B00000000000000004AAEAFAE464C451200000000000000006DB0AFAE464C4912000000000000000088D0B188C4464B230000000000000000B3894EE1E54D24923C3B3B3D3D3E5152DB8A1D1D282590823C2283A3B5CEA552A4008D8E7D1F3F873E8453503D3D4F260904C6C18D7D2AA05296B5AA9853CF3A110F92C6C18D2EA55299C7DDE2847C210E0E1C271AC17CA153C9ACD7E3844348A10E2C807C1A86A754ACE1E3E2944342C8853F0E2E40B9D694C9B9E3E2953D2000A6CDC7E5E3CFAC94E1E5E3E0986C6B00A8C798E5C9C7D698C7E3E2DE976C6100AB9D9EE5E2C9B9C9B7D6DECF98873D00D9D79EB6B7CBD4D4CCBAACAC98000000FF000000FF000000FF000000FF0000000000000000000000000000000000000000000000000000000000000000000080000000800000008000000080030000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 39
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(39,'Career Progression','',10,'CAREER')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 39
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000090909001401010001011C001717170018181800050D3400231E320024242400282828003B3B3B0001074300121A4C001D305C004C3A5B00304F770045454500575757005C5C5C006161610065656500686868007272720077777700797979007F7F7F00D66C3000D86C3000DC6C3000DC6E3400DC723600DC743A00DC793F00E1672B00ED6A2B00F8672100E9706300F6874B00FF814E00FF845200FF895200FF8B5800FF8E5A00FF8E5E00FF915A00FF876100FF8F6300FF8D6500FC896E00FF986A00FF967B00D6A07B00FFA57600FFA27900205581003D5F86003F5F8A00575E8000635F82004E6B9000436D9F004878A8005A7DA700638FBE007284B9005796CB006594C5006994C8006898CE00739AC9006DA0D00068A2D60073A8D60076B4D6007DB7D70000FFFF0061FFFF008080800085858500898989008D8D8D009191910095959500989898009D9D9D008B8EB500A1A1A100A8A8A800AFAFAF00B1B1B100B6B6B600B9B9B900BDBDBD00E7BA8D00FFB18100FFB58500FFB89600D6B3B300FFC29800FFC59E00FFD1A900F6D2B300FFDEB800FFE3BE0087B1D70085B4D70082BCD700C7BACD00C0C0C000C7C7C700CACACA00CCCCCC00D2C0D100D2D2D200D6D6D600FFE5C500FFE9C700FFEBCB00FFEDCB00FFEFCF00D8CBE000D8E9EC00E4E4E400E8E8E800FFEFEB00F0F0F000F6F6F6000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00201F1E1D1D211B1A1D1A144B4D004D0067635F342A2F2A2A2A2A4C4B4B4D4B4D7464635E357C2F282A2A4D4B4B4B4B4D77666460FFFF2F26272A144D4B4B4B4D73777761553A312A2722504B4B4B4B4D255F3070460F2432235D5450504D505000330239443F363B655AFFFF50194D5000000C3D686847457A5A541452504D14000040416A4A4943506D57111854135A0000063842483C3E0015575213156F00000000030D370B0E094D5A5B57570000000007000000007852545B5B5B5400000000000000006B000415545A174D0000000000000000000000010A1405100000000000000000000008000000000000000000000000000000000000000000000000050000000000000000000000000000000000000000000080000000C0000000C0000000C0010000C0030000C0030000E1030000FF030000FF070000FF8F0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 38
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(38,'Succession Planning','',10,'SUCCESSION')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 38
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000090909001401010001011C001717170018181800050D3400231E320024242400282828003B3B3B0001074300121A4C001D305C004C3A5B00304F7700575757005C5C5C006161610065656500696969006E6E6E007272720077777700797979007F7F7F00ED6A2B00F8672100CE7F4600EC754100E6794100E9794000E9706300E1814700E9834D00F6874B00FF814E00E98B5500FF845200FF895200FF8B5800FF8F5B00FF915A00FF876100FF8F6300FC896E00E9936000FF986A00FF967B00D6A07B00DAA57900E9A16F00FFA27900205581003D5F86003F5F8A00575E8000635F82004E6B9000436D9F004878A8005A7DA7000018FE00638FBE007284B9005796CB006594C5006994C8006898CE00739AC9006DA0D00068A2D60073A8D60076B4D6007DB7D70000FFFF0061FFFF008080800084858400898989008C8C8C009191910095959500999999009D9D9D008B8EB500A1A1A100A7A7A700A8A8A800AEAEAE00B1B1B100B6B6B600B9B9B900BDBDBD00EBB38600E7BA8D00FFB18100FFB89600D6B3B300FFC29800FFD1A900F6D2B300FFDEB80087B1D70085B4D70082BCD700B6A9CD00C7BACD00BAC1C200C0C0C000C5C6C500CACACA00CCCCCC00D2C0D100C8D2D400D2D2D200D6D6D600FFE9C700FFEDCB00D8CBE000D8E9EC00E2E2E200E4E4E400EDEDED00FFEFEB00F0F0F000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF0014141414131313131313134B4D004D006D6D5A564F4F4F4F4F4F4C4B4B4D4B4D6D6F6D58567D4F4F4F4F4D4B4B4B4B4D56746F5AFFFF4F194D4F4D4D4B4B4B4D567A7A5A5414524F4D144D4B4B4B4B4D5258526D581019541358341A1A1A1A1A00540014565410166F61FFFF2B2426290000094D6D5A5856766255392F2A261A000052546D6D5A562D71460F20301B5F000004135458194D0238443F353A6500000000010A1304000C3C67674745000000000800000000004041694A494300000000000000000000063742483B3D0000000000000000000000030D360B0E0000000000000000000007000000007700000000000000000000000000006B00000000050000000000000000000000000000000000000000000080000000C0000000C0000000C0010000C1030000C1030000E3030000FF030000FF030000FF870000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 37
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(37,'Menu','',10,'MENU')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 37
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000001000800680500001600000028000000100000002000000001000800000000000001000000000000000000000001000000010000000000008452180084522100845A2100BD9C7B00BD9C8400BDA58400C6A58400CEA59400F7C6B500F7CEB500FFCEB500C6C6C6000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF0000000000000000000000000000000000000604060506040605060406050604060005FFFFFFFFFFFFFF0A0B0A080A0B060006FFFFFFFFFFFFFF0A0A0800080A060006FF00000000FFFF0A0B000B000B050006FFFFFFFFFFFFFF0B090A0A0B09060005FF0000000000FF050706060507060006FFFFFFFFFFFFFF0A0A0B090A0A060006FF000000FFFFFF0A0B000B000B050006FFFFFFFFFFFFFF0B0908000809060005FFFFFFFFFFFFFF0A0B0A080A0B060402010201020102010201020102010207FFFFFFFFFFFFFFFFFFFF01030C020105FFFFFFFFFFFFFFFFFFFF020C0C0C020605070606050706060507010201030100000000000000000000000000000000FFFF00008000000080000000800000008000000080000000800000008000000080000000800000008000000000000000000000000000000000000000FFFF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 16
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(16,'Orders','',10,'ORDERS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 16
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000000080000080000000808000800000008000800080800000C0C0C000C0DCC000F0CAA60004040400080808000C0C0C0011111100161616001C1C1C002222220029292900555555004D4D4D004242420039393900807CFF005050FF009300D600FFECCC00C6D6EF00D6E7E70090A9AD000000330000006600000099000000CC00003300000033330000336600003399000033CC000033FF00006600000066330000666600006699000066CC000066FF00009900000099330000996600009999000099CC000099FF0000CC000000CC330000CC660000CC990000CCCC0000CCFF0000FF660000FF990000FFCC00330000003300330033006600330099003300CC003300FF00333300003333330033336600333399003333CC003333FF00336600003366330033666600336699003366CC003366FF00339900003399330033996600339999003399CC003399FF0033CC000033CC330033CC660033CC990033CCCC0033CCFF0033FF330033FF660033FF990033FFCC0033FFFF00660000006600330066006600660099006600CC006600FF00663300006633330066336600663399006633CC006633FF00666600006666330066666600666699006666CC00669900006699330066996600669999006699CC006699FF0066CC000066CC330066CC990066CCCC0066CCFF0066FF000066FF330066FF990066FFCC00CC00FF00FF00CC009999000099339900990099009900CC009900000099333300990066009933CC009900FF00996600009966330099336600996699009966CC009933FF009999330099996600999999009999CC009999FF0099CC000099CC330066CC660099CC990099CCCC0099CCFF0099FF000099FF330099CC660099FF990099FFCC0099FFFF00CC00000099003300CC006600CC009900CC00CC0099330000CC333300CC336600CC339900CC33CC00CC33FF00CC660000CC66330099666600CC669900CC66CC009966FF00CC990000CC993300CC996600CC999900CC99CC00CC99FF00CCCC0000CCCC3300CCCC6600CCCC9900CCCCCC00CCCCFF00CCFF0000CCFF330099FF6600CCFF9900CCFFCC00CCFFFF00CC003300FF006600FF009900CC330000FF333300FF336600FF339900FF33CC00FF33FF00FF660000FF663300CC666600FF669900FF66CC00CC66FF00FF990000FF993300FF996600FF999900FF99CC00FF99FF00FFCC0000FFCC3300FFCC6600FFCC9900FFCCCC00FFCCFF00FFFF3300CCFF6600FFFF9900FFFFCC006666FF0066FF660066FFFF00FF666600FF66FF00FFFF66002100A5005F5F5F00777777008686860096969600CBCBCB00B2B2B200D7D7D700DDDDDD00E3E3E300EAEAEA00F1F1F100F8F8F800F0FBFF00A4A0A000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AF045446E6E456E0A0A0A0A0A6D0A0A0A0A9345930A0A6E0A0A0A0AEF0AED0A0A0A0A8D450A0ABC0A0A0A0A140A430A0A0A0ABC686E0A0A0A0A0AEF0A0A0AED0A0A0A0A9345EF0A0A0A0AEBED0AED120A0A8D0A0A6E450A0A0A0A0A0A0A0A0A0A0A928DF78D6E930A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0AB48B070A8B8BAE0A0A0A0A0A0A0A0A0A0AADF10A078BBC0A0A0A0A0A0A0A0A0A0AB4CFB58BAE0A0A0A0A0A0A0A0A0A0A0ADDADF28B070A0A0A0A0A0A0A0A0A0A0A0ACFCF8B0A0A0A0A0A0A0A0A0A0A0A0A0AB5ADB50A0A0A0A0A0A0A0A0A0A0A0A0A0ACF0A0A0A0A0A0A0A0A0A0A0AFFFF000080FB0000C6F10000E6F10000E3E00000F1E00000D9FB0000C0FB0000FFFB000088FB0000C8FB0000C1FB0000C1FB0000E3FB0000E3FB0000F7FB0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 17
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(17,'Event Log','',10,'EVENTLOG')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 17
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x00000100010010100000010008006805000016000000280000001000000020000000010008000000000000010000000000000000000000010000000100000000000000630000006B0000634A31006B4A310073523900947B6300947B6B00AD7B63009C846B009C847300B5846B00A58C7300A58C7B00C6947B000031FF000039FF00188CBD0000A5EF0000ADEF0000ADF70073E7FF0073EFFF00AD948400B5948400B59C8400B5948C00B59C8C00BD9C8C00BD9C9400BDA58C00BDA59400C69C8400D69C8C00C6A59400CEA59400C6AD9C00CEAD9C00D6A59400D6AD9C00D6B59C00CEADA500DEB5A500D6B5AD00DEB5AD00D6BDAD00D6BDB500DEBDB500E7BDAD00EFBDAD00D6C6AD00D6C6B500DEC6B500DEC6BD00DECEBD00EFC6B500F7CEB500DEC6C600DECEC600D6CECE00D6D6D600E7D6C600E7D6CE00EFD6CE00EFDECE00FFD6C600F7DECE00FFDECE00EFDED600F7DED600F7E7D600FFE7D600F7E7DE00FFEFDE00D8E9EC00EFEFEF00F7E7E700FFEFE700F7EFEF00FFF7EF00F7F7F700FFFFF7000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00000000040504050405040504050405000000061F1E221E1F1E1F191B1C2204000011094D4D47474343372A27311F050000161307494947474102012A311F040000000C4F4F352D252802020B381F050000110C514F4C39353F422F2C301F04000015140AFF0F101D250202082B170500000018FF510F0F4444460102290D030000111EFFFF2E241D23200E02020D04000016120DFF100F3B01483A02011B0400000024FFFF0F101A23020202451E0500001124FFFF0F0F3C4B514F4F491F040000151417FF0F101D2526210E4F1F0500000033FFFF100F50FFFF51514F22040000003DFFFFFFFFFFFFFFFFFF511F050000003E4036363435343432342D360000E0010000C00100008001000080010000C00100008001000080010000C00100008001000080010000C00100008001000080010000C0010000C0010000C0030000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 13
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(13,'Standard Reports','',10,'STANDARDREPORTS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 13
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000001000800680500001600000028000000100000002000000001000800000000000001000000000000000000000001000000010000000000006D353600703C3C00604830007050400070584000D0683000D16D360080454500884745009552530092555400985755009D565600915957009F5C5B00B1565700B95F56008060500080685000BB615E00806860009863600096756C009C716F00AC6E6E00B1626600B96F7000B3777300BD727300D0704000D1744600E0704000E0744600E0784000C4636300C06E6A00C26C6D00CD6E6E00DA706F00C0727200C0757700CE7E7A00D67D7700A0807000A9827E00E0885000D0827E00E0906000E0987000F0987000F0A07000B4888400A0908000B0908000B0A09000CE818200D3878600DC878600DD898600D78B8900DD908E00DE988B00DD9C9B00DB9F9E00E08F8D00F0A08000FFA88000F0A89000F0B09000FFB09000FFB89000C0A8A000D0A8A000C0B0A000D0B0A000D0B8A000D0B8B000E2A7A300E5A4A600E4A6A900EDA9A800F0A8A800F5B7B600D0C0B000E0C8B000D0C8C000E0C8C000E0D0C000E0D8D000E0E0D000F0E0D000E0E0E000E0E8E000F0E8E000FFE8E000F0F0E000FFF0E000F0F0F000FFF8F0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00003703030303030303030304000000000048FF5D5C5A59585755480400000000004AFF575454614B4B49480400000000004CFFFFFF6363615F5B4B0400000000004CFF57573515131212050303030304004CFFFFFF356362605E563557554804004CFF575736FF575454592C4B494804004C0E0BFF340BFFFF6361375F5B4B040045091A020B022E222220064D4B4B040C082414112723160133321E63614D043A293A192D1C2310102E2222545454040039410CFF172610FFFFFFFF63635F043D4F52390F1824100132302E222220063E3C53513A2B24291D4747434332321F000050402F2A1B4444423231302E22200000404F002A1D000000000000000000800F0000800F0000800F0000800F000080000000800000008000000080000000800000000000000000000000800000000000000000000000C0000000C9FF0000

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 42
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(42,'Workflow','',10,'WORKFLOW')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 42
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000100000000000000000000000100000001000000000000215963002159730021617300218294003192B50031A2C60042B2C60052B2D60052C3D60052D3E70073DBE70094BECE00A5C3C600ADCBD600BDCFD600BDD3DE0084E3F70084EBF70094EBF700B5F3FF00C6C3C600C6D3D600CED7DE00CEDBDE00CEDFDE00CEDFE700DEE3E700C6F3F700C6FBFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF00000000000000000000000000000000000000000D03170000000000151515151500000F0407030F000000000000000000001905110B0A0116000000FFFFFFFF001B061D1413110902000000FFFFFFFF00001A081D1412060C000000FFFFFFFF00000019081C060E0000000000000000000000001006170000000000000000000000000000000000000000000000000000000015150015151500000000000000000000000000000000150000000000000000FFFFFFFFFFFFFF001500000000000000FFFFFFFFFFFFFF001500000000000000FFFFFFFFFFFFFF00150000000000000000000000000000000000000000000000000000000000000000000000000000FFFF0000E3E00000C1C0000080C000000000000080C00000C1C00000E3FF0000F7FF0000C0FF0000807F0000003F0000003F0000003F000080FF0000FFFF0000


/* ------------------------------------------------------------- */
PRINT 'Step 25 of 29 - Modifying stored procedures for Workflow'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetPicture]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetPicture]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetPicture
	(
		@piPictureID		integer
	)
	AS
	BEGIN
		SELECT TOP 1 name, picture
		FROM ASRSysPictures
		WHERE pictureID = @piPictureID
	END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowFormItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetWorkflowFormItems]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetWorkflowFormItems
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
			@iCount		integer
	
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
	
		SET @psErrorMessage = ''''
	
		SELECT 			
			@piBackColour = isnull(webFormBGColor, 16777166),
			@piBackImage = isnull(webFormBGImageID, 0),
			@piBackImageLocation = isnull(webFormBGImageLocation, 0),
			@piWidth = isnull(webFormWidth, -1),
			@piHeight = isnull(webFormHeight, -1)
		FROM ASRSysWorkflowElements
		WHERE ASRSysWorkflowElements.ID = @piElementID
	
		SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
		FROM ASRSysWorkflowInstances
		WHERE ASRSysWorkflowInstances.ID = @piInstanceID
	
		CREATE TABLE #itemValues (ID integer, value varchar(8000))	
	
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
		FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iItemType = 1
			BEGIN
				/* Database value. */
				SELECT @sTableName = ASRSysTables.tableName, 
					@sColumnName = ASRSysColumns.columnName,
					@iDBColumnDataType = ASRSysColumns.dataType
				FROM ASRSysColumns
				INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
				WHERE ASRSysColumns.columnID = @iDBColumnID
	
				IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID

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
			ELSE
			BEGIN
				/* Workflow value. */
				SELECT @sValue = ASRSysWorkflowInstanceValues.value
				FROM ASRSysWorkflowInstanceValues
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
					AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
					AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
			END
	
			INSERT INTO #itemValues (ID, value)
			VALUES (@iID, @sValue)
	
			FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
		END
		CLOSE itemCursor
		DEALLOCATE itemCursor
	
		SELECT thisFormItems.*, 
			#itemValues.value, 
			CASE
				WHEN thisFormItems.itemType = 4 THEN sourceItems.itemType 
				WHEN thisFormItems.itemType = 1 THEN sourceColumns.dataType 
				ELSE null
			END AS [sourceItemType]
		FROM ASRSysWorkflowElementItems thisFormItems
		LEFT OUTER JOIN #itemValues ON thisFormItems.ID = #itemValues.ID
		LEFT OUTER JOIN ASRSysWorkflowElements sourceElements ON thisFormItems.WFFormIdentifier = sourceElements.identifier
			AND len(isnull(thisFormItems.WFFormIdentifier, '''')) > 0 
		LEFT OUTER JOIN ASRSysWorkflowElementItems sourceItems ON sourceElements.id = sourceItems.elementID
			AND thisFormItems.WFValueIdentifier = sourceItems.identifier
		LEFT OUTER JOIN ASRSysColumns sourceColumns ON thisFormItems.DBColumnID = sourceColumns.columnID
			AND thisFormItems.DBColumnID > 0 
		WHERE thisFormItems.elementID = @piElementID
		ORDER BY thisFormItems.ZOrder DESC

		DROP TABLE #itemValues
	END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetSetting]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetSetting]
	
	SET @sTemp = 'CREATE PROCEDURE spASRGetSetting (
		@psSection		varchar(8000),
		@psKey		varchar(8000),
		@psDefault		varchar(8000),
		@pfUserSetting		bit,
		@psResult		varchar(8000)	OUTPUT
	)
	AS
	BEGIN
		/* Return the required user or system setting. */
		DECLARE
			@iCount	integer
	
		IF @pfUserSetting = 1
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM ASRSysUserSettings
			WHERE userName = SYSTEM_USER
				AND section = @psSection		
				AND settingKey = @psKey
	
			SELECT @psResult = settingValue 
			FROM ASRSysUserSettings
			WHERE userName = SYSTEM_USER
				AND section = @psSection		
				AND settingKey = @psKey
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM ASRSysSystemSettings
			WHERE section = @psSection		
				AND settingKey = @psKey
	
			SELECT @psResult = settingValue 
			FROM ASRSysSystemSettings
			WHERE section = @psSection		
				AND settingKey = @psKey
		END
	
		IF @iCount = 0
		BEGIN
			SET @psResult = @psDefault 		
		END
	END'
	
	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRWorkflowLogPurge]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRWorkflowLogPurge]
	
	SET @sTemp = 'CREATE PROCEDURE spASRWorkflowLogPurge 
	AS
	BEGIN
		DECLARE @sUnit char(1),
	        		@iPeriod int,
			@dtPurgeDate datetime,
			@dtToday datetime
	
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
				AND completionDateTime < @dtPurgeDate
		END
	END'
	
	EXEC (@sTemp)


/* ------------------------------------------------------------- */
PRINT 'Step 26 of 29 - Modifying columns for Module Setup'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysModuleSetup')
	and name = 'ParameterValue'
	and length < 1000


	if @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysModuleSetup ALTER COLUMN
					[ParameterValue] [varchar] (1000) NULL '
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 27 of 29 - Modifying columns for Workflow'
	SET @sName = ''

	SELECT @sName = constraintobjs.name
	FROM sysobjects constraintobjs
	INNER JOIN syscolumns ON constraintobjs.id = syscolumns.cdefault
	INNER JOIN sysobjects ON syscolumns.id = sysobjects.id
	WHERE sysobjects.name = 'ASRSysWorkflowInstances'
		and syscolumns.name = 'InitiationDateTime'

	IF @sName IS null SET @sName = ''

	IF LEN(@sName) > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysWorkflowInstances] DROP
			CONSTRAINT [' + @sName + ']'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysWorkflowInstances] ADD 
		CONSTRAINT [DF_ASRSysWorkflowInstances_InitiationDateTime] DEFAULT (getdate()) FOR [InitiationDateTime]'
	EXEC sp_executesql @NVarCommand

	/* Add new record selector identifier columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'RecSelWebFormIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[RecSelWebFormIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'RecSelIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[RecSelIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 28 of 29 - SQL 2005 User lockout configuration'

	IF @sSQLVersion = '9'
	BEGIN
		SELECT @NVarCommand = 'DELETE FROM ASRSysSystemSettings  WHERE Section = ''misc'' AND SettingKey = ''cfg_pcl'''
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES (''misc'', ''cfg_pcl'', ''0'')'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 29 of 29 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.0')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.0.0')

delete from asrsyssystemsettings
where [Section] = 'Server DLL' and [SettingKey] = 'Minimum Version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('Server DLL', 'Minimum Version', '3.0.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.0')


SELECT @NVarCommand = 'USE master
	GRANT EXECUTE ON master..xp_LoginConfig TO public
	GRANT EXECUTE ON master..xp_EnumGroups TO public'
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.0 Of HR Pro'
