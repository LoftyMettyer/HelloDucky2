
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
PRINT 'Step 1 - Create New IsValidNINumber function'

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
PRINT 'Step 2 - Create New IsValidPayrollCharacterSet function'

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
			WHILE (@Index <= datalength(@input) AND @result = 1)
			BEGIN
				IF charindex(substring(@input,@Index,1),@ValidCharacters) = 0
					SET @result = 0;
				SET @Index = @Index + 1;
			END	

		END';
	
	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 3 - Create New Replace Characters within a String function'

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
PRINT 'Step 4 - Add new formatting columns to ASRSysSSIntranetLinks'

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'UseFormatting')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD UseFormatting bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET UseFormatting = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Formatting_DecimalPlaces')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Formatting_DecimalPlaces int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Formatting_DecimalPlaces = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Formatting_Use1000Separator')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Formatting_Use1000Separator bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Formatting_Use1000Separator = 0'
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Formatting_Prefix')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Formatting_Prefix varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Formatting_Prefix = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Formatting_Suffix')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Formatting_Suffix varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Formatting_Suffix = '''''
	END	
	
--------------------------------------------------------------------------------------------
-- Conditional Formatting Columns
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'UseConditionalFormatting')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD UseConditionalFormatting bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET UseConditionalFormatting = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Operator_1')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Operator_1 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Operator_1 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Value_1')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Value_1 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Value_1 = '''''
	END	

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Style_1')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Style_1 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Style_1 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Colour_1')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Colour_1 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Colour_1 = '''''
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Operator_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Operator_2 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Operator_2 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Value_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Value_2 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Value_2 = '''''
	END	

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Style_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Style_2 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Style_2 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Colour_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Colour_2 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Colour_2 = '''''
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Operator_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Operator_3 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Operator_3 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Value_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Value_3 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Value_3 = '''''
	END	

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Style_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Style_3 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Style_3 = '''''
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'ConditionalFormatting_Colour_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD ConditionalFormatting_Colour_3 varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET ConditionalFormatting_Colour_3 = '''''
	END

--------------------------------------------------------------------------------------------
-- Separator Border Colour Column
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'SeparatorColour')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD SeparatorColour varchar(MAX) NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET SeparatorColour = '''''
	END


--------------------------------------------------------------------------------------------
-- Initial Display Mode for Charts
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'InitialDisplayMode')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD InitialDisplayMode int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET InitialDisplayMode = 0'
	END

--------------------------------------------------------------------------------------------
-- Additional Data columns for Multi-Axis Charts
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_TableID_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_TableID_2 int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_TableID_2 = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ColumnID_2')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ColumnID_2 int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ColumnID_2 = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_TableID_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_TableID_3 int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_TableID_3 = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ColumnID_3')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ColumnID_3 int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ColumnID_3 = 0'
	END		

--------------------------------------------------------------------------------------------
-- Sort Order columns for Charts
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_SortOrderID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_SortOrderID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_SortOrderID = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_SortDirection')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_SortDirection int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_SortDirection = 0'
	END
	

--------------------------------------------------------------------------------------------
-- Colour Code ID column for Charts
--------------------------------------------------------------------------------------------

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ColourID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ColourID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ColourID = 0'
	END
	
	
		
/* ------------------------------------------------------------- */
PRINT 'Step 5 - Modifying Workflow Data Structures'

	/* ASRSysWorkflowElementItems - Add new LookupFilterColumnID column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'LookupFilterColumnID';

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							LookupFilterColumnID [int] NULL';
		EXEC sp_executesql @NVarCommand;
	END

	/* ASRSysWorkflowElementItems - Add new LookupFilterOperator column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'LookupFilterOperator';

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							LookupFilterOperator [int] NULL';
		EXEC sp_executesql @NVarCommand;
	END

	/* ASRSysWorkflowElementItems - Add new LookupFilterValue column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'LookupFilterValue';

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							LookupFilterValue [varchar] (200) NULL';
		EXEC sp_executesql @NVarCommand;
	END

/* ------------------------------------------------------------- */
PRINT 'Step 6 - Modifying Workflow Stored Procedures'


	----------------------------------------------------------------------
	-- spASRGetWorkflowItemValues
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowItemValues]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
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
				@piItemType	integer OUTPUT
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
					@iLookupFilterColumnType	int;

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
					@iLookupFilterColumnType = isnull(COLS.dataType, 0)
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
						'''' AS [ASRSysLookupFilterValue],
						CASE
							WHEN [value] = @sDefaultValue THEN 1
							ELSE 0
						END AS [ASRSysDefaultValueFlag]						
					FROM @dropdownValues;
				END
				
				IF (@iItemType = 14) -- Lookup
					AND (@iLookupColumnID > 0)
				BEGIN
					SELECT @sTableName = ASRSysTables.tableName,
						@sColumnName = ASRSysColumns.columnName,
						@iOrderID = ASRSysTables.defaultOrderID,
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

					SET @sSelectSQL = @sSelectSQL + '','';
					
					IF len(ltrim(rtrim(@sDefaultValue))) = 0 
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ '' 0 AS [ASRSysDefaultValueFlag]'';
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
									THEN '''''''' + REPLACE(@sDefaultValue, '''''''', '''''''''''') + ''''''''
								ELSE @sDefaultValue 
							END
							+ ''   THEN 1''
							+ ''   ELSE 0''
							+ '' END AS [ASRSysDefaultValueFlag]'';
					END;

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

	----------------------------------------------------------------------
	-- spASRGetWorkflowFormItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFormItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFormItems];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		(
			@piInstanceID				integer,
			@piElementID				integer,
			@psErrorMessage				varchar(MAX)	OUTPUT,
			@piBackColour				integer			OUTPUT,
			@piBackImage				integer			OUTPUT,
			@piBackImageLocation		integer			OUTPUT,
			@piWidth					integer			OUTPUT,
			@piHeight					integer			OUTPUT,
			@piCompletionMessageType	integer			OUTPUT,
			@psCompletionMessage		varchar(200)	OUTPUT,
			@piSavedForLaterMessageType	integer			OUTPUT,
			@psSavedForLaterMessage		varchar(200)	OUTPUT,
			@piFollowOnFormsMessageType	integer			OUTPUT,
			@psFollowOnFormsMessage		varchar(200)	OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iID				integer,
				@iItemType			integer,
				@iDefaultValueType	integer,
				@iDBColumnID		integer,
				@iDBColumnDataType	integer,
				@iDBRecord			integer,
				@sWFFormIdentifier	varchar(MAX),
				@sWFValueIdentifier	varchar(MAX),
				@sValue				varchar(MAX),
				@sSQL				nvarchar(MAX),
				@sSQLParam			nvarchar(500),
				@sTableName			sysname,
				@sColumnName		sysname,
				@iInitiatorID		integer,
				@iRecordID			integer,
				@iStatus			integer,
				@iCount				integer,
				@iWorkflowID		integer,
				@iElementType		integer, 
				@iType				integer,
				@fValidRecordID		bit,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iRequiredTableID	integer,
				@iRequiredRecordID	integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@iInitParent1TableID	integer,
				@iInitParent1RecordID	integer,
				@iInitParent2TableID	integer,
				@iInitParent2RecordID	integer,
				@fDeletedValue			bit,
				@iTempElementID			integer,
				@iColumnID				integer,
				@iResultType			integer,
				@sResult				varchar(MAX),
				@fResult				bit,
				@dtResult				datetime,
				@fltResult				float,
				@iCalcID				integer,
				@iSize					integer,
				@iDecimals				integer,
				@iPersonnelTableID		integer,
				@sIdentifier			varchar(MAX);
		
			DECLARE @itemValues table(ID integer, value varchar(MAX), type integer)	
					
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
		
			SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_PERSONNEL''
				AND parameterKey = ''Param_TablePersonnel''
		
			IF @iPersonnelTableID = 0
			BEGIN
				SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel''
			END
						
			SELECT 			
				@piBackColour = isnull(webFormBGColor, 16777166),
				@piBackImage = isnull(webFormBGImageID, 0),
				@piBackImageLocation = isnull(webFormBGImageLocation, 0),
				@piWidth = isnull(webFormWidth, -1),
				@piHeight = isnull(webFormHeight, -1),
				@iWorkflowID = workflowID,
				@piCompletionMessageType = CompletionMessageType,
				@psCompletionMessage = CompletionMessage,
				@piSavedForLaterMessageType = SavedForLaterMessageType,
				@psSavedForLaterMessage = SavedForLaterMessage,
				@piFollowOnFormsMessageType = FollowOnFormsMessageType,
				@psFollowOnFormsMessage = FollowOnFormsMessage
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
				ASRSysWorkflowElementItems.wfValueIdentifier,
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
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 19
					OR ASRSysWorkflowElementItems.itemType = 20
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
				SET @sValue = ''''
		
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
						SET @iBaseTableID = @iPersonnelTableID
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
								AND IV.identifier = @sWFValueIdentifier
								AND Es.identifier = @sWFFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID
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
								AND Es.identifier = @sWFFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
		
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
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
								'' FROM '' + @sTableName +
								'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(100), @iRecordID)
						SET @sSQLParam = N''@sValue varchar(MAX) OUTPUT''
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
					OR (@iItemType = 17)
				BEGIN
					IF @iStatus = 7 -- Previously SavedForLater
					BEGIN
						SELECT @sValue = 
							CASE
								WHEN (@iItemType = 6 AND IVs.value = ''1'') THEN ''TRUE'' 
								WHEN (@iItemType = 6 AND IVs.value <> ''1'') THEN ''FALSE'' 
								WHEN (@iItemType = 7 AND (upper(ltrim(rtrim(IVs.value))) = ''NULL'')) THEN '''' 
								WHEN (@iItemType = 17 AND IVs.fileUpload_File IS null) THEN ''0''
								WHEN (@iItemType = 17 AND NOT IVs.fileUpload_File IS null) THEN ''1''
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
									WHEN @iResultType = 2 THEN STR(@fltResult, 100, @iDecimals)
									WHEN @iResultType = 3 THEN 
										CASE 
											WHEN @fResult = 1 THEN ''TRUE''
											ELSE ''FALSE''
										END
									WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
									ELSE convert(varchar(MAX), @sResult)
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
				VALUES (@iID, @sValue, @iType)
		
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
				IV.type AS [sourceItemType],
				LUFC.ColumnName AS [lookupFilterColumnName],
				LUFC.datatype AS [lookupFilterColumnDataType],
				LUI.ID AS [lookupFilterValueID],
				LUI.ItemType AS [lookupFilterValueType]
			FROM ASRSysWorkflowElementItems thisFormItems
			LEFT OUTER JOIN @itemValues IV ON thisFormItems.ID = IV.ID
			LEFT OUTER JOIN ASRSysColumns LUFC ON thisFormItems.lookupFilterColumnID = LUFC.ColumnID
			LEFT OUTER JOIN ASRSysWorkflowElementItems LUI ON thisFormItems.lookupFilterValue = LUI.Identifier
				AND LUI.elementID = @piElementID
				AND LEN(LUI.Identifier) > 0
			WHERE thisFormItems.elementID = @piElementID
			ORDER BY thisFormItems.ZOrder DESC
		END';

	EXECUTE sp_executeSQL @sSPCode;



/* ------------------------------------------------------------- */
PRINT 'Step 7 - Updating Details for ACS Rebranding'
/* ------------------------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'support' and [SettingKey] = 'email'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('support', 'email', 'service.delivery@advancedcomputersoftware.com')

delete from asrsyssystemsettings
where [Section] = 'support' and [SettingKey] = 'webpage'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('support', 'webpage', 'http://webfirst.advancedcomputersoftware.com')
/* ------------------------------------------------------------- */
-- Update the system permission image for Module Access
	IF EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 1)
	BEGIN
		SELECT @ptrval = TEXTPTR([picture]) 
		FROM dbo.[ASRSysPermissionCategories]
		WHERE categoryID = 1;

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000001000800680500001600000028000000100000002000000001000800000000000000000000000000000000000000000000000000D63C0400A73B1F00ED3E0400AF460D00FE420900E445150097453400FF4B1100A5552200EA501E00FF5318009D504500C7562F00FF5B2000AB623A00FF632800DE673600A6635900FF6B3000A2704C00AB6B6200E9744000BE7D4400FF7C4100B57B7200D7825700FF844400FF844C00FF8C4B00B8867D00DF8B6300FF8C5500FF945300E8936B00BC908900FF9A5500FF9C5B00D19F6100C3999200FFA26000DF9F8400FCA8660075A1B000FFA47900C5A19B00D3AA9100FBB16C00C9A7A200FFAC8500E3AE9300D2ADA80092AFB800AAB0B00086B6C70053B8DD00D9BAB60073BFD800FFC5AA00D9C3BF0031C4FC00D6C9C700FFCCB400FFD0A8004FCAFF00E2CDCA00EBD6B900EAD6CD00EBDAD7008ADAFF00D1DFE3009CE0FF00F0E3E100B0E6FF00F5EBE700F9F2EF00E7F8FF00FDFBF90000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000004D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4C474343494C4D4D4D4D4D4D4C4C4C2F1D2828282222474D4D4D4D464845110E1F2023201F1906404D4D4D444C220C171C272E29231A15064A4D4D4C40060F17202E2E2E241A120C224C4D4D3A101D22263435352533190D184C4D4C4917163B363338352C3B2A0A374C4D4C4A150736361A2A3633191D09414C4D4D42100219281C1C2C21120405284B4C4D4706001B302B2B2B2B2B0503264B3F4D4C1D01193E3D39393D30080B3C46484D4D4C2214192831311E1318424C4C4D4D4D4D4C42372D2626323C4A4D4D4D4D4D4D4D4D4D4C4C4C4C4C4C4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4D4DFFFF0000F81F0000800F0000000700000003000000010000800100000001000000010000800000008000000080000000C0010000E00F0000F81F0000FFFF0000	END
/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 8 - Adding new control type'
/* ------------------------------------------------------------- */

    SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] ALTER COLUMN [ControlType] integer;'
    EXEC sp_executesql @NVarCommand;

    SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysControls] ALTER COLUMN [ControlType] integer;'
    EXEC sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */



/* ------------------------------------------------------------- */
PRINT 'Step X - '

	
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
