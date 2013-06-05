
/* --------------------------------------------------- */
/* Update the database from version 4.1 to version 4.2 */
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

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '4.1') and (@sDBVersion <> '4.2')
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
PRINT 'Step 1 of X - Create New IsValidNINumber function'

	DELETE FROM [ASRSysFunctions] WHERE FunctionID = 75
	INSERT [ASRSysFunctions]
	([functionID],[functionName],[returnType],[timeDependent],[category],[spName],[nonStandard],[runtime],[UDF])
	VALUES
	(75,'Is Valid NI Number',3,0,'Comparison','sp_ASRFn_IsValidNINumber',0,0,0)

	DELETE FROM [ASRSysFunctionParameters] WHERE FunctionID = 75
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(75,1,1,'<National Insurance Number>')

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsValidNINumber]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsValidNINumber];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsValidNINumber]
		(
			@result integer OUTPUT,
			@input varchar(MAX)
		)
		AS
		BEGIN

			DECLARE @ValidPrefixes varchar(MAX);
			DECLARE @ValidSuffixes varchar(MAX);
			DECLARE @Prefix varchar(MAX);
			DECLARE @Suffix varchar(MAX);
			DECLARE @Numerics varchar(MAX);

			SET @result = 1;
			IF ISNULL(@input,'''') = '''' RETURN

			SET @ValidPrefixes = 
				''/AA/AB/AE/AH/AK/AL/AM/AP/AR/AS/AT/AW/AX/AY/AZ'' +
				''/BA/BB/BE/BH/BK/BL/BM/BT'' +
				''/CA/CB/CE/CH/CK/CL/CR'' +
				''/EA/EB/EE/EH/EK/EL/EM/EP/ER/ES/ET/EW/EX/EY/EZ'' +
				''/GY'' +
				''/HA/HB/HE/HH/HK/HL/HM/HP/HR/HS/HT/HW/HX/HY/HZ'' +
				''/JA/JB/JC/JE/JG/JH/JJ/JK/JL/JM/JN/JP/JR/JS/JT/JW/JX/JY/JZ'' +
				''/KA/KB/KE/KH/KK/KL/KM/KP/KR/KS/KT/KW/KX/KY/KZ'' +
				''/LA/LB/LE/LH/LK/LL/LM/LP/LR/LS/LT/LW/LX/LY/LZ'' +
				''/MA/MW/MX'' +
				''/NA/NB/NE/NH/NL/NM/NP/NR/NS/NW/NX/NY/NZ'' +
				''/OA/OB/OE/OH/OK/OL/OM/OP/OR/OS/OX'' +
				''/PA/PB/PC/PE/PG/PH/PJ/PK/PL/PM/PN/PP/PR/PS/PT/PW/PX/PY'' +
				''/RA/RB/RE/RH/RK/RM/RP/RR/RS/RT/RW/RX/RY/RZ'' +
				''/SA/SB/SC/SE/SG/SH/SJ/SK/SL/SM/SN/SP/SR/SS/ST/SW/SX/SY/SZ'' +
				''/TA/TB/TE/TH/TK/TL/TM/TP/TR/TS/TT/TW/TX/TY/TZ'' +
				''/WA/WB/WE/WK/WL/WM/WP'' +
				''/YA/YB/YE/YH/YK/YL/YM/YP/YR/YS/YT/YW/YX/YY/YZ'' +
				''/ZA/ZB/ZE/ZH/ZK/ZL/ZM/ZP/ZR/ZS/ZT/ZW/ZX/ZY/'';

			SET @ValidSuffixes = ''/ /A/B/C/D/'';

			SET @Prefix = ''/''+left(@input+''  '',2)+''/''
			SET @Suffix = ''/''+substring(@input+'' '',9,1)+''/''
			SET @Numerics = SUBSTRING(@input,3,6)

			IF charindex(@Prefix,@ValidPrefixes) = 0 OR charindex(@Suffix,@ValidSuffixes) = 0 OR ISNUMERIC(@Numerics) = 0
				SET @result = 0;

		END';
		
	EXECUTE sp_executeSQL @sSPCode;

/* ------------------------------------------------------------- */
PRINT 'Step 2 of X - Create New IsValidPayrollCharacterSet function'

	DELETE FROM [ASRSysFunctions] WHERE FunctionID = 76
	INSERT [ASRSysFunctions]
	([functionID],[functionName],[returnType],[timeDependent],[category],[spName],[nonStandard],[runtime],[UDF])
	VALUES
	(76,'Is Valid Payroll Character Set',3,0,'Comparison','sp_ASRFn_IsValidForPayrollCharset',0,0,0)

	DELETE FROM [ASRSysFunctionParameters] WHERE FunctionID = 76
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(76,1,1,'<String>')
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(76,2,1,'<Payroll Character Set>')

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsValidForPayrollCharset]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsValidForPayrollCharset];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsValidForPayrollCharset]
		(
			@result integer OUTPUT,
			@input varchar(MAX),
			@Charset varchar(1)
		)
		AS
		BEGIN

			--Charset A - typically Address
			--Charset C - typically Forename
			--Charset D - typically Surname

			DECLARE @ValidCharacters varchar(MAX);
			DECLARE @Index int;


			IF      @Charset = ''A'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). =!"%&*;<>+:?''
			ELSE IF @Charset = ''B'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ''
			ELSE IF @Charset = ''C'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''''
			ELSE IF @Charset = ''D'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). ''
			ELSE IF @Charset = ''G'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). =!"%&*;<>+:?''
			ELSE IF @Charset = ''H'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''. ''
			
			SET @result = 1;
			SET @Index = 1;
			WHILE (@Index <= len(@input) AND @result = 1)
			BEGIN
				IF charindex(substring(@input,@Index,1),@ValidCharacters) = 0
					SET @result = 0;
				SET @Index = @Index + 1;
			END	

		END';
	
	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Create New Replace Characters within a String function'

	DELETE FROM [ASRSysFunctions] WHERE FunctionID = 77
	INSERT [ASRSysFunctions]
	([functionID],[functionName],[returnType],[timeDependent],[category],[spName],[nonStandard],[runtime],[UDF])
	VALUES
	(77,'Replace Characters in a String',1,0,'Character','sp_ASRFn_ReplaceCharsInString',0,1,0)

	DELETE FROM [ASRSysFunctionParameters] WHERE FunctionID = 77
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(77,1,1,'<String>')
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(77,2,1,'<Search For>')
	INSERT [ASRSysFunctionParameters]
	([functionID],[parameterIndex],[parameterType],[parameterName])
	VALUES
	(77,3,1,'<Replace With>')

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ReplaceCharsInString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ReplaceCharsInString];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ReplaceCharsInString]
		(
			@psResult		varchar(MAX) OUTPUT,
			@input varchar(MAX),
			@searchstring varchar(MAX),
			@replacestring varchar(MAX)
		)
		AS
		BEGIN

			IF ISNULL(@input, '''') = '''' RETURN;			
			
			SET @psResult = REPLACE(@input, @searchstring, @replacestring);

		END';
	
	EXECUTE sp_executeSQL @sSPCode;





/* ------------------------------------------------------------- */
PRINT 'Step X of X - '

	
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
values('database', 'version', '4.2')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '4.2.0')

delete from asrsyssystemsettings
where [Section] = 'ssintranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('ssintranet', 'minimum version', '4.2.0')

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
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.2')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.2 Of HR Pro'
