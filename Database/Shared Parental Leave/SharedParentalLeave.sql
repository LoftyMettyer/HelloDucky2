/* --------------------------------- */
/* Add Shared Parental Leave Changes */
/* For Version 4.3 & and above only  */
/* Mark Edwynn - February 2015       */
/* --------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion int,
	@NVarCommand nvarchar(MAX),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @sSPCode nvarchar(MAX);

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON;
SET @DBName = DB_NAME();

/* -------------------------------------- */
/* Check SQL Server version compatibility */
/* -------------------------------------- */
SELECT @iSQLVersion = convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))
IF (@iSQLVersion < 10)
BEGIN
    Print '+--------------------------------------------------------------------------+'
    Print '|                                                                          |'
    Print '|                            SCRIPT FAILURE                                |'
    Print '|                                                                          |'
    Print '| This version of OpenHR is only compatible with SQL Server 2008 or later. |'
    Print '| Please upgrade SQL Server before upgrading to this version of OpenHR.    |'
    Print '|                                                                          |'
    Print '+--------------------------------------------------------------------------+'
	RETURN
END

IF @DBName = 'master'
BEGIN
    Print '+-----------------------------------------------------------------------+'
    Print '|                                                                       |'
    Print '|                            SCRIPT FAILURE                             |'
    Print '|                                                                       |'
    Print '|        This script should not be run on the ''master'' database.      |'
    Print '|                                                                       |'
    Print '+-----------------------------------------------------------------------+'
    RETURN
END

IF IS_SRVROLEMEMBER('systemadmin') = 0
BEGIN
    Print '+-----------------------------------------------------------------------+'
    Print '|                                                                       |'
    Print '|                            SCRIPT FAILURE                             |'
    Print '|                                                                       |'
    Print '| This script can only be run by a member of the ''systemadmin'' role.  |'
    Print '|                                                                       |'
    Print '+-----------------------------------------------------------------------+'
    RETURN
END

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */
SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
WHERE [Section] = 'database' AND [SettingKey] = 'version';

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF @sDBVersion NOT IN ('4.3', '5.0', '5.1', '5.2', '8.0', '8.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ------------------------------------ */
/* Check if Script has already been run */
/* ------------------------------------ */
IF EXISTS(SELECT [SettingValue] FROM [dbo].[ASRSysSystemSettings] WHERE [Section] = 'statutory' AND [SettingKey] = 'sharedparentalleave')
BEGIN
	RAISERROR('The database has already been configured for Shared Parental Leave', 16, 1)
	RETURN
END
ELSE 
BEGIN
	/* --------------------------------------- */
	/* Revised Unpaid Parental Leave Procedure */
	/* --------------------------------------- */
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfstat_ParentalLeaveEntitlement]') AND xtype = 'FN')
		DROP FUNCTION [dbo].udfstat_ParentalLeaveEntitlement;
	EXECUTE sp_executeSQL N'CREATE FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement] (
		@DateOfBirth	datetime,
		@AdoptedDate	datetime,
		@Disabled		bit,
		@Region			varchar(MAX))
	RETURNS float
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @pdblResult			float,
			@Today					datetime,
			@ChildAge				integer,
			@ChildAgeLimit			integer,
			@DisabledAgeLimit		integer,
			@Adopted				bit,
			@YearsOfResponsibility	integer,
			@StartDate				datetime,
			@Standard				integer,
			@Extended				integer,
			@Ireland				bit;

		SET @Today = GETDATE();
		SET @Standard = 65;
		SET @Extended = 90;
		SET @ChildAgeLimit = 5;
		SET @DisabledAgeLimit = 18;

		-- All entitlement in UK to age 18 from 5th April 2015
		IF DATEDIFF(d, ''04/05/2015'', @Today) >= 0
			SET @ChildAgeLimit = 18;
				
		IF @Region IN (''Rep of Ireland'', ''Ireland'', ''Eire'')
		BEGIN
			SET @Ireland = 1
			SET @Standard = 70;
			SET @Extended = 70;
			SET @ChildAgeLimit = 8;
			SET @DisabledAgeLimit = 16;
		END;

		IF DATEDIFF(d, ''03-08-2013'', @Today) >= 0
		BEGIN
			   SET @Standard = 90;
			   SET @Extended = 90;
		END;

		-- Check if we should used the Date of Birth or the Date of Adoption column...
		SET @Adopted = 0;
		SET @StartDate = @DateOfBirth;
		IF NOT @AdoptedDate IS NULL
		BEGIN
			SET @Adopted = 1;
			SET @StartDate = @AdoptedDate;
		END;

		-- Set variables based on this date...
		--( years of responsibility = years since born or adopted)
		SELECT @ChildAge = [dbo].[udfsys_wholeyearsbetweentwodates](@DateOfBirth, @Today);
		SELECT @YearsOfResponsibility = [dbo].[udfsys_wholeyearsbetweentwodates](@StartDate, @Today);

		SELECT @pdblResult = CASE
			WHEN @Disabled = 0 AND @Adopted = 0 AND @ChildAge < @ChildAgeLimit
				THEN @Standard
			WHEN @Disabled = 0 AND @Adopted = 1 AND @Ireland = 0 AND @ChildAge < 18 AND @YearsOfResponsibility < 5
				THEN @Standard
			WHEN @Disabled = 0 AND @Adopted = 1 AND @Ireland = 1 AND @ChildAge < @ChildAgeLimit
				THEN @Standard
			WHEN @Disabled = 1 AND @Adopted = 0 AND @ChildAge < @DisabledAgeLimit AND DATEDIFF(d, ''12/15/1994'', @DateOfBirth) >= 0
				THEN @Extended
			WHEN @Disabled = 1 AND @Adopted = 1 AND @ChildAge < @DisabledAgeLimit AND DATEDIFF(d, ''12/15/1994'', @AdoptedDate) >= 0
				THEN @Extended
			ELSE 
				0
			END;

		RETURN ISNULL(@pdblResult, 0);

	END';

	/* ---------------------- */
	/* Create Table Procedure */
	/* ---------------------- */
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spshpl_scriptnewtable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spshpl_scriptnewtable];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spshpl_scriptnewtable](@tableID integer OUTPUT, @tablename varchar(255), @tabletype tinyint, @islocked bit, @uniquekey varchar(37))
		AS
		BEGIN

			SET NOCOUNT ON;
			
			DECLARE @ssql nvarchar(MAX),
					@newtableID integer,
					@ownerID varchar(37);		

			-- Can we safely create this table?
			SELECT @newtableID = ISNULL(tableID,0) FROM dbo.[tbsys_tables] WHERE [TableName] = @tablename;

			IF @newtableID > 0 RETURN;

			EXEC dbo.spASRGetNextObjectIdentitySeed ''ASRSysTables'', @newtableID OUTPUT;

			SELECT @ownerID = [SettingValue] FROM dbo.[ASRSysSystemSettings] WHERE [Section] = ''database'' AND [SettingKey] = ''ownerid''

			-- System objects update
			INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
				SELECT @uniquekey, 1, @newtableID, @ownerID, ''01/01/1900'',1, @islocked, GETDATE()

			-- System metadata
			INSERT dbo.[tbsys_tables] (TableID, TableType, TableName, DefaultOrderID, RecordDescExprID, DefaultEmailID
					, ManualSummaryColumnBreaks, AuditDelete, AuditInsert, isremoteview)
				VALUES (@newtableID, @tabletype, @tablename, 0, 0, 0, 0, 0, 0, 0)

			-- Physically create this table (is regenerated by the System Manager save)	
			SET @ssql = N''CREATE TABLE dbo.tbuser_'' + @tablename + '' ([ID] integer IDENTITY(1,1) PRIMARY KEY CLUSTERED
								, [updflag] int NULL, [_description] nvarchar(MAX) NULL, [_deleted] bit, [_deleteddate] datetime, [TimeStamp] timestamp NOT NULL);'';
			EXECUTE sp_executesql @ssql;

			-- Create a view on this table (is replaced by System Manager save, so no need to be precise)
			SET @ssql = N''CREATE VIEW dbo.['' + @tablename + ''] AS SELECT * FROM dbo.[tbuser_'' + @tablename + ''];'';
			EXECUTE sp_executesql @ssql;

			SET @tableID = @newtableID;

		END';

	/* ----------------------- */
	/* Create Column Procedure */
	/* ----------------------- */
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spshpl_scriptnewcolumn]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spshpl_scriptnewcolumn];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spshpl_scriptnewcolumn] (@columnid integer OUTPUT, @tableid integer, @columnname varchar(255)
		, @datatype integer, @statusBarMessage varchar(255), @size integer, @decimals integer, @Use1000Separator bit, @defaultvalue varchar(max)
		, @islocked bit, @uniquekey varchar(37), @IsIDColumn bit, @IsFKColumn bit, @mandatory bit)
	AS
	BEGIN

		DECLARE @ssql nvarchar(MAX),
				@tablename varchar(255),
				@datasyntax	varchar(255),
				@ownerID varchar(37);

		DECLARE @spinnerMinimum integer,
			@spinnerMaximum integer,
			@spinnerIncrement integer,
			@audit bit,
			@duplicate bit,
			@columntype integer,
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
		
		SET @spinnerMinimum = 0;
		SET @spinnerMaximum = 0;
		SET @spinnerIncrement = 0;
		SET @audit = 0;
		SET @duplicate = 0;
		SET @columntype = 0;
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
		SET @UniqueCheckType = 0;
		SET @Trimming = 1;
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

		-- Logic
		IF @datatype = -7
		BEGIN
			SET @datasyntax = ''bit'';
			SET @defaultvalue = ''FALSE'';
			SET @controltype = 1;
			SET @DefaultDisplayWidth = 1;
		END

		-- OLE
		IF @datatype = -4
		BEGIN
			SET @datasyntax = ''varbinary(max)'';
			SET @controltype = 8;
			SET @DefaultDisplayWidth = 255;
			SET @OLEType = 2; 
			SET @MaxOLESizeEnabled = 1; 
			SET @MaxOLESize = 8000;
		END

		-- Photo
		IF @datatype = -3
		BEGIN
			SET @datasyntax = ''varbinary(max)'';
			SET @controltype = 1024;
			SET @DefaultDisplayWidth = 255;
			SET @OLEType = 2; 
			SET @MaxOLESizeEnabled = 1; 
			SET @MaxOLESize = 8000;
		END

		-- Link
		IF @datatype = -2
		BEGIN
			SET @datasyntax = ''varchar(255)'';
			SET @controltype = 2048;
			SET @columntype = 4;
			SET @DefaultDisplayWidth = 1;
		END

		-- Working Pattern
		IF @datatype = -1
		BEGIN
			SET @datasyntax = ''varchar(14)'';
			SET @controltype = 4096;
			SET @DefaultDisplayWidth = 14;
		END
		
		-- Numeric
		IF @datatype = 2
		BEGIN
			SET @datasyntax = ''numeric('' + convert(varchar(5), @size) + '','' + convert(varchar(5), @decimals) + '')'';
			SET @defaultvalue = 0;	
			SET @controltype = 64;
			SET @DefaultDisplayWidth = convert(varchar(5), @size);
			SET @alignment = 1;
		END

		-- Integers
		IF @datatype = 4
		BEGIN
			SET @datasyntax = ''integer'';
			SET @controltype = 64;
			SET @spinnerMaximum = 10;
			SET @spinnerIncrement = 1;
			SET @alignment = 1;
			-- Is ID column?
			IF @IsIDColumn = 1 OR @IsFKColumn = 1
			BEGIN
				SET @columntype = 3;	
				SET @DefaultDisplayWidth = 1;
			END
			ELSE
				SET @DefaultDisplayWidth = @size;
				SET @size = 10;
			END
		END
		
		-- Date
		IF @datatype = 11
		BEGIN
			SET @datasyntax = ''datetime'';
			SET @controltype = 64;
			SET @DefaultDisplayWidth = 10;
		END

		-- Character
		IF @datatype = 12
		BEGIN
			SET @controltype = 64;
			SET @DefaultDisplayWidth = @size;
			IF @size = 2147483646
				SET @datasyntax = ''nvarchar(max)'';
			ELSE 
				SET @datasyntax = ''varchar('' + convert(varchar(5), @size) + '')'';
		END

		-- System objects update
		SELECT @ownerID = [SettingValue] FROM dbo.[ASRSysSystemSettings] WHERE [Section] = ''database'' AND [SettingKey] = ''ownerid''

		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT @uniquekey, 2, @columnid, @ownerID, ''01/01/1900'',1, @islocked, GETDATE()

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

		RETURN;';
	
	/* ---------------- */
	/* Define Variables */
	/* ---------------- */
	DECLARE @tabShPL_Adoption int
		, @colA_ID int
		, @colA_ID_Pers int
		, @colA_Main_Other_Adopter int
		, @colA_Placement_Date int
		, @colA_SAP_Curtailment_Date int
		, @colA_Partner_Forename int
		, @colA_Partner_Surname int
		, @colA_Partner_Address_1 int
		, @colA_Partner_Address_2 int
		, @colA_Partner_Address_3 int
		, @colA_Partner_Address_4 int
		, @colA_Partner_Postcode int
		, @colA_Partner_NI_Number int
		, @colA_No_NI_Number_Declaration int
		, @colA_Date_Notification_Received int
		, @colA_Intended_ShPL_Start_Date int
		, @colA_Intended_ShPL_End_Date int
		, @colA_SAP_Weeks_Paid int
		, @colA_Total_ShPP_Weeks_Available int
		, @colA_ShPP_Weeks_Employee int
		, @colA_ShPP_Weeks_Partner int
		, @colA_Partner_Employer_Name int
		, @colA_Partner_Employer_Address_1 int
		, @colA_Partner_Employer_Address_2 int
		, @colA_Partner_Employer_Address_3 int
		, @colA_Partner_Employer_Address_4 int
		, @colA_Partner_Employer_Postcode int
		, @colA_Evidence_from_Adoption_Agency int
		, @colA_Declaration_from_Employee int
		, @colA_Declaration_from_Other_Adopter int
		, @colA_Notes int
		, @colA_Payroll_Company_Code int
		, @colA_Staff_Number int
		, @colA_Full_Name int
		, @colA_Trigger_to_Payroll int
		, @colA_Total_SPLIT_Days int
		, @fltrShPL_Adoption int;

	DECLARE @tabShPLA_Leave_Requests int
		, @colAR_ID int
		, @colAR_ID_A int
		, @colAR_Date_of_Request int
		, @colAR_Date_Requested_From int
		, @colAR_Date_Requested_To int
		, @colAR_Binding_Request int
		, @colAR_Consent_from_Other_Adopter int
		, @colAR_Request_Cancelled int
		, @colAR_ShPP_Weeks int
		, @colAR_Notes int
		, @fltrShPLA_Leave_Requests int;

	DECLARE @tabShPLA_SPLIT_Days int
		, @colAS_ID int
		, @colAS_ID_A int
		, @colAS_Start_Date int
		, @colAS_End_Date int
		, @colAS_Reason int
		, @colAS_Notes int
		, @colAS_SPLIT_Days int
		, @fltrShPLA_SPLIT_Days int;

	DECLARE @tabShPL_Birth int
		, @colB_ID int
		, @colB_ID_Pers int
		, @colB_Mother_Father_Partner int
		, @colB_Expected_Birth_Date int
		, @colB_SMP_Curtailment_Date int
		, @colB_Partner_Forename int
		, @colB_Partner_Surname int
		, @colB_Partner_Address_1 int
		, @colB_Partner_Address_2 int
		, @colB_Partner_Address_3 int
		, @colB_Partner_Address_4 int
		, @colB_Partner_Postcode int
		, @colB_Partner_NI_Number int
		, @colB_No_NI_Number_Declaration int
		, @colB_Date_Notification_Received int
		, @colB_Intended_ShPL_Start_Date int
		, @colB_Intended_ShPL_End_Date int
		, @colB_SMP_Weeks_Paid int
		, @colB_Total_ShPP_Weeks_Available int
		, @colB_ShPP_Weeks_Employee int
		, @colB_ShPP_Weeks_Partner int
		, @colB_Partner_Employer_Name int
		, @colB_Partner_Employer_Address_1 int
		, @colB_Partner_Employer_Address_2 int
		, @colB_Partner_Employer_Address_3 int
		, @colB_Partner_Employer_Address_4 int
		, @colB_Partner_Employer_Postcode int
		, @colB_Evidence_of_Birth int
		, @colB_Declaration_from_Employee int
		, @colB_Declaration_from_Other_Parent int
		, @colB_Notes int
		, @colB_Payroll_Company_Code int
		, @colB_Staff_Number int
		, @colB_Full_Name int
		, @colB_Trigger_to_Payroll int
		, @colB_Total_SPLIT_Days int
		, @colB_Actual_Birth_Date int
		, @colB_Location_of_Birth int
		, @fltrShPL_Birth int;

	DECLARE @tabShPLB_Leave_Requests int
		, @colBR_ID int
		, @colBR_ID_B int
		, @colBR_Date_of_Request int
		, @colBR_Date_Requested_From int
		, @colBR_Date_Requested_To int
		, @colBR_Binding_Request int
		, @colBR_Consent_from_Other_Parent int
		, @colBR_Request_Cancelled int
		, @colBR_ShPP_Weeks int
		, @colBR_Notes int
		, @fltrShPLB_Leave_Requests int;

	DECLARE @tabShPLB_SPLIT_Days int
		, @colBS_ID int
		, @colBS_ID_B int
		, @colBS_Start_Date int
		, @colBS_End_Date int
		, @colBS_Reason int
		, @colBS_Notes int
		, @colBS_SPLIT_Days int
		, @fltrShPLB_SPLIT_Days int;

	DECLARE @tabPersonnel_Records int
		, @colPR_Staff_Number int
		, @colPR_Company_Code int
		, @colPR_Full_Name int
		, @colPR_Surname int
		, @colPR_Is_Current_Employee_for_Payroll int;

	DECLARE @tabGlobal_Variables int
		, @colGV_Global_Key int
		, @colGV_Enable_Payroll_Transfer int
		, @colGV_Logic_Value int;

	DECLARE @islocked bit
		, @payrollModule bit
		, @triggerFlag bit
		, @orderID int
		, @screenID int
		, @hScreenID int
		, @exprID int
		, @exprCompID int
		, @scrPersID int
		, @scrAdoptID int
		, @scrBirthID int
		, @access varchar(2)
		, @screenFontName varchar (50)
		, @screenFontSize int
		, @screenLabelHeight int
		, @screenColumnHeight int
		, @enablePayType varchar (6)
		, @cnameID_Pers varchar(6)
		, @cnameShPLA_ID varchar(6)
		, @cnameShPLB_ID varchar(6);

	SET @islocked = 0;
	IF @islocked = 1
		SET @access = 'RO'
	ELSE
		SET @access = 'RW';

	SET @payrollModule = 0;

	SET @tabPersonnel_Records = (SELECT [ParameterValue] FROM [dbo].[ASRSysModuleSetup] WHERE [ModuleKey] = 'MODULE_PERSONNEL' AND [ParameterKey] = 'Param_TablePersonnel');

	SET @screenFontName = (SELECT [SettingValue] FROM [dbo].[ASRSysSystemSettings] WHERE [Section] = 'ScreenDesigner' AND [SettingKey] = 'FontName')
	IF @screenFontName IS NULL
		SET @screenFontName = (SELECT TOP 1 [DfltFontName] FROM [dbo].[ASRSysScreens] WHERE [TableID] = @tabPersonnel_Records ORDER BY [ScreenID]);

	SET @screenFontSize = (SELECT FLOOR([SettingValue]) FROM [dbo].[ASRSysSystemSettings] WHERE [Section] = 'ScreenDesigner' AND [SettingKey] = 'FontSize');
	IF @screenFontSize IS NULL
		SET @screenFontSize = (SELECT TOP 1 [DfltFontSize] FROM [dbo].[ASRSysScreens] WHERE [TableID] = @tabPersonnel_Records ORDER BY [ScreenID]);

	IF @screenFontSize = 8
	BEGIN
		SET @screenLabelHeight = 195;
		SET @screenColumnHeight = 315;
	END
	ELSE
	BEGIN
		SET @screenLabelHeight = 315;
		SET @screenColumnHeight = 435;
	END;		
	
	SET @colPR_Company_Code = (SELECT [ASRColumnID] FROM [dbo].[ASRSysAccordTransferFieldDefinitions] WHERE [TransferTypeID] = 0 AND [TransferFieldID] = 0);
	IF @colPR_Company_Code > 0
	BEGIN
		SET @payrollModule = 1
		SET @colPR_Staff_Number = (SELECT [ASRColumnID] FROM [dbo].[ASRSysAccordTransferFieldDefinitions] WHERE [TransferTypeID] = 0 AND [TransferFieldID] = 1);
		SET @colPR_Is_Current_Employee_for_Payroll = (SELECT [columnID] FROM [dbo].[ASRSysColumns] WHERE [tableID] = @tabPersonnel_Records AND [ColumnName] = 'Is_Current_Employee_for_Payroll');
		IF @colPR_Is_Current_Employee_for_Payroll IS NULL
			SET @triggerFlag = 0
		ELSE
			SET @triggerFlag = 1;
		SET @enablePayType = 'None';
		SET @tabGlobal_Variables = (SELECT [TableID] FROM [dbo].[tbsys_tables] WHERE [TableName] = 'Global_Variables');
		SET @colGV_Global_Key = (SELECT [columnID] FROM [dbo].[ASRSysColumns] WHERE [tableID] = @tabGlobal_Variables AND [ColumnName] = 'Global_Key');
		IF EXISTS (SELECT [Global_Key] FROM [dbo].[tbuser_Global_Variables] WHERE [Global_Key] = 'ENABLEPAY')
		BEGIN
			SET @enablePayType = 'Record';
			SET @colGV_Logic_Value = (SELECT [columnID] FROM [dbo].[ASRSysColumns] WHERE [tableID] = @tabGlobal_Variables AND [ColumnName] = 'Logic_Value');
		END
		ELSE
		BEGIN
			IF EXISTS (SELECT DISTINCT [name] FROM [sys].[all_columns] WHERE [name] = 'Enable_Payroll_Transfer')
			BEGIN
				SET @enablePayType = 'Column';
				SET @colGV_Enable_Payroll_Transfer = (SELECT [columnID] FROM [dbo].[ASRSysColumns] WHERE [tableID] = @tabGlobal_Variables AND [ColumnName] = 'Enable_Payroll_Transfer');
			END
		END
	END
	ELSE
		SET @colPR_Staff_Number = (SELECT [ParameterValue] FROM [dbo].[ASRSysModuleSetup] WHERE [ModuleKey] = 'MODULE_PERSONNEL' AND [ParameterKey] = 'Param_FieldsEmployeeNumber');
	
	SET @colPR_Full_Name = (SELECT [columnID] FROM [dbo].[ASRSysColumns] WHERE [tableID] = @tabPersonnel_Records AND [ColumnName] = 'Full_Name');

	SET @colPR_Surname = (SELECT [ParameterValue] FROM [dbo].[ASRSysModuleSetup] WHERE [ModuleKey] = 'MODULE_PERSONNEL' AND [ParameterKey] = 'Param_FieldsSurname')

	SET @cnameID_Pers = 'ID_' + CONVERT(varchar(3), @tabPersonnel_Records);

	SET @scrPersID = (SELECT TOP 1 h.[parentScreenID]
						FROM [dbo].[ASRSysScreens] s
						INNER JOIN [dbo].[ASRSysHistoryScreens] h
						ON h.[parentScreenID] = s.[ScreenID]
						WHERE s.[TableID] = @tabPersonnel_Records AND s.[SSIntranet] = 0
						ORDER BY s.[ScreenID]);

	/* --------------------------------------------- */
	/* Create Shared Parental Leave (Adoption) Table */
	/* --------------------------------------------- */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPL_Adoption OUTPUT, 'ShPL_Adoption', 2, @islocked, 'C0D22416-060D-4BB9-A304-6B4A1FD9D7BB';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_ID OUTPUT, @tabShPL_Adoption, 'ID', 4, NULL, 0, 0, 0, '', @islocked, '97D11C7D-0AB5-4C02-B336-459D146C3363', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_ID_Pers OUTPUT, @tabShPL_Adoption, @cnameID_Pers, 4, NULL, 0, 0, 0, '', @islocked, 'C11919CA-6A1B-4CC8-9F16-0C073D5EA54D', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Main_Other_Adopter OUTPUT, @tabShPL_Adoption, 'Main_Other_Adopter', 12, 'Select if the employee is the main or other adopter', 13, 0, 0, '', @islocked, '612C5EF9-354D-4F74-94A0-24085C351961', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Placement_Date OUTPUT, @tabShPL_Adoption, 'Placement_Date', 11, 'Enter the child placement date', 0, 0, 0, '', @islocked, '90285290-2235-4229-9669-198B63DBB4A9', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_SAP_Curtailment_Date OUTPUT, @tabShPL_Adoption, 'SAP_Curtailment_Date', 11, 'Enter the SAP curtailment date (mandatory)', 0, 0, 0, '', @islocked, 'E3CCF57C-0E0A-459E-936D-B92ABBA1A73B', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Forename OUTPUT, @tabShPL_Adoption, 'Partner_Forename', 12, 'Enter partner''s forename (mandatory)', 30, 0, 0, '', @islocked, 'A102408D-13F3-4C02-83F9-86B537F286AC', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Surname OUTPUT, @tabShPL_Adoption, 'Partner_Surname', 12, 'Enter partner''s surname (mandatory)', 30, 0, 0, '', @islocked, 'A5BB48DC-6683-45E7-9BBC-1E88C5CD2E1B', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Address_1 OUTPUT, @tabShPL_Adoption, 'Partner_Address_1', 12, 'Enter 1st line of partner''s address (mandatory)', 30, 0, 0, '', @islocked, '64886E4E-8701-427B-ABEC-53F281756197', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Address_2 OUTPUT, @tabShPL_Adoption, 'Partner_Address_2', 12, 'Enter 2nd line of partner''s address (mandatory)', 30, 0, 0, '', @islocked, '6DC0C0E0-B6C6-484A-87DD-872719BEEFB6', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Address_3 OUTPUT, @tabShPL_Adoption, 'Partner_Address_3', 12, 'Enter 3rd line of partner''s address', 30, 0, 0, '', @islocked, '791FB7BA-FFCF-49F9-B79C-3991EC05ECE8', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Address_4 OUTPUT, @tabShPL_Adoption, 'Partner_Address_4', 12, 'Enter 4th line of partner''s address', 30, 0, 0, '', @islocked, '43274A15-E379-4C1D-B613-FB46E2B78436', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Postcode OUTPUT, @tabShPL_Adoption, 'Partner_Postcode', 12, 'Enter partner''s postcode', 8, 0, 0, '', @islocked, '84BC8BB2-F6C7-4A26-986F-EAEFA536A426', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_NI_Number OUTPUT, @tabShPL_Adoption, 'Partner_NI_Number', 12, 'Enter partner''s NI Number', 9, 0, 0, '', @islocked, 'E5DDEBFE-44B4-4D3F-9F2D-4C8C2A94F9E3', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_No_NI_Number_Declaration OUTPUT, @tabShPL_Adoption, 'No_NI_Number_Declaration', -7, 'Check box if no NI Number exists', 0, 0, 0, 'FALSE', @islocked, '71993819-1E84-4504-8F1B-3303E0335E72', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Date_Notification_Received OUTPUT, @tabShPL_Adoption, 'Date_Notification_Received', 11, 'Enter date notification of ShPL received (mandatory)', 0, 0, 0, '', @islocked, 'E9EE72EC-3774-4323-947F-2C04375BDDB0', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Intended_ShPL_Start_Date OUTPUT, @tabShPL_Adoption, 'Intended_ShPL_Start_Date', 11, 'Enter intended ShPL start date', 0, 0, 0, '', @islocked, '35EF640A-C989-4081-923D-2266E25D31DC', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Intended_ShPL_End_Date OUTPUT, @tabShPL_Adoption, 'Intended_ShPL_End_Date', 11, 'Enter intended ShPL end date', 0, 0, 0, '', @islocked, 'A11AFC45-C2DE-4F3D-8451-F5F9E8303230', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_SAP_Weeks_Paid OUTPUT, @tabShPL_Adoption, 'SAP_Weeks_Paid', 4, 'Enter number of SAP weeks paid to main adopter', 2, 0, 0, '0', @islocked, '85D9B401-E671-40C9-8567-7BFBA9A55606', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Total_ShPP_Weeks_Available OUTPUT, @tabShPL_Adoption, 'Total_ShPP_Weeks_Available', 4, '', 2, 0, 0, '39', @islocked, 'A0FE34B7-2CEB-4E74-BDFD-F02A0AB5214A', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_ShPP_Weeks_Employee OUTPUT, @tabShPL_Adoption, 'ShPP_Weeks_Employee', 4, 'Enter number of ShPP weeks to be claimed by employee (mandatory)', 2, 0, 0, '0', @islocked, '413E7D22-190F-4B27-9FB2-C455DCAED247', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_ShPP_Weeks_Partner OUTPUT, @tabShPL_Adoption, 'ShPP_Weeks_Partner', 4, 'Enter number of ShPP weeks to be claimed by employee''s partner (mandatory)', 2, 0, 0, '0', @islocked, '2CDB1FA8-BA9E-44E9-BD29-E358A20D3F50', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Name OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Name', 12, 'Enter name of partner''s employer', 30, 0, 0, '', @islocked, '044A5624-75C8-48E4-A7FB-0BC19A4236AE', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Address_1 OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Address_1', 12, 'Enter 1st line of partner''s employer address', 30, 0, 0, '', @islocked, 'E2180034-FCF2-4C6C-BA34-DD7C1525F4E4', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Address_2 OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Address_2', 12, 'Enter 2nd line of partner''s employer address', 30, 0, 0, '', @islocked, 'CDDC08EF-D886-40BE-84C5-591B3C1333C8', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Address_3 OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Address_3', 12, 'Enter 3rd line of partner''s employer address', 30, 0, 0, '', @islocked, 'BF89ED20-3379-4CA4-AD4C-04EB78FFB26D', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Address_4 OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Address_4', 12, 'Enter 4th line of partner''s employer address', 30, 0, 0, '', @islocked, '1B618EF5-2ABC-4286-8C11-81281BA6FCA4', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Partner_Employer_Postcode OUTPUT, @tabShPL_Adoption, 'Partner_Employer_Postcode', 12, 'Enter partner''s employer postcode', 8, 0, 0, '', @islocked, 'A64F9E9A-A96C-42A2-8480-7B12F4677BE7', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Evidence_from_Adoption_Agency OUTPUT, @tabShPL_Adoption, 'Evidence_from_Adoption_Agency', -4, 'Click button to link/access document', 0, 0, 0, '', @islocked, '82159BD9-8FFA-4DB3-96DC-7838C34BB6EB', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Declaration_from_Employee OUTPUT, @tabShPL_Adoption, 'Declaration_from_Employee', -7, 'Check box when declaration from employee received', 0, 0, 0, 'FALSE', @islocked, 'C2B9CCB7-0AA8-493B-988D-739ABDF61401', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Declaration_from_Other_Adopter OUTPUT, @tabShPL_Adoption, 'Declaration_from_Other_Adopter', -7, 'Check box when declaration from other adopter received', 0, 0, 0, 'FALSE', @islocked, 'FF949621-1C86-4857-BD36-7698CFC4E1B2', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Notes OUTPUT, @tabShPL_Adoption, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, '5B0B025D-03B9-4AFE-89BC-77D4792C4AF8', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Payroll_Company_Code OUTPUT, @tabShPL_Adoption, 'Payroll_Company_Code', 12, '', 2, 0, 0, '', @islocked, 'D5B7D005-7E9F-448E-BB52-E0A7B1BB7C81', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Staff_Number OUTPUT, @tabShPL_Adoption, 'Staff_Number', 12, '', 8, 0, 0, '', @islocked, 'A2E1DD3C-B48C-40BE-914F-390B2AD71D39', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Full_Name OUTPUT, @tabShPL_Adoption, 'Full_Name', 12, '', 40, 0, 0, '', @islocked, '9FAC3754-B994-4BE8-9FE7-C5FDCC120AB8', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Trigger_to_Payroll OUTPUT, @tabShPL_Adoption, 'Trigger_to_Payroll', -7, '', 0, 0, 0, 'FALSE', @islocked, 'A3DD2E40-F08C-422C-AA7D-7FDE7BC81F64', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colA_Total_SPLIT_Days OUTPUT, @tabShPL_Adoption, 'Total_SPLIT_Days', 4, '', 2, 0, 0, '0', @islocked, '4C08C551-1056-44BC-B00F-E80FB973776B', 0, 0, 0;

	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Date_Notification_Received', @tabShPL_Adoption, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colA_Date_Notification_Received, 'F', 0, 1)
			, (@orderID, @colA_Main_Other_Adopter, 'F', 1, 1)
			, (@orderID, @colA_Placement_Date, 'F', 2, 1)
			, (@orderID, @colA_ShPP_Weeks_Employee, 'F', 3, 1)
			, (@orderID, @colA_ShPP_Weeks_Partner, 'F', 4, 1)
			, (@orderID, @colA_Intended_ShPL_Start_Date, 'F', 5, 1)
			, (@orderID, @colA_Intended_ShPL_End_Date, 'F', 6, 1)
			, (@orderID, @colA_Date_Notification_Received, 'O', 1, 0);

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPL_Adoption;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'Statutory Shared Parental Leave (Adoption)', @tabShPL_Adoption, 0, 5760, 11400, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 2800, 5080, @screenLabelHeight * 2, 2625, '(Must add up to Total ShPP Weeks Available)', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 79, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 2240, 5080, @screenLabelHeight, 2595, '(39 - SAP Weeks Paid)', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 76, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2655, 'The Employee is the :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 50, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2145, 'Placement Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 49, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2820, 'SAP Curtailment Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 48, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 6, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 3810, 'SAP Weeks Paid to Main Adopter :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 47, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, NULL, 0, 256, 0, 2240, 280, @screenLabelHeight, 3390, 'Total ShPP Weeks Available :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 46, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, NULL, 0, 256, 0, 2720, 280, @screenLabelHeight, 3060, 'Weeks Employee Claiming :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 45, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 9, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 3660, 'Weeks Other Adopter Claiming :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 44, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 10, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 2940, 'Intended ShPL Start Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 43, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 11, NULL, 0, 256, 0, 4160, 280, @screenLabelHeight, 3030, 'Intended ShPL End Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 42, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 12, @tabShPL_Adoption, @colA_Intended_ShPL_End_Date, 64, 0, 4120, 4200, @screenColumnHeight, 1755, 'ShPL_Adoption.Intended_ShPL_End_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 41, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 13, @tabShPL_Adoption, @colA_Intended_ShPL_Start_Date, 64, 0, 3640, 4200, @screenColumnHeight, 1755, 'ShPL_Adoption.Intended_ShPL_Start_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 40, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 14, @tabShPL_Adoption, @colA_ShPP_Weeks_Partner, 64, 0, 3165, 4200, @screenColumnHeight, 620, 'ShPL_Adoption.ShPP_Weeks_Partner', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 39, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 15, @tabShPL_Adoption, @colA_ShPP_Weeks_Employee, 64, 0, 2680, 4200, @screenColumnHeight, 620, 'ShPL_Adoption.ShPP_Weeks_Employee', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 38, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 16, @tabShPL_Adoption, @colA_Total_ShPP_Weeks_Available, 64, 0, 2200, 4200, @screenColumnHeight, 620, 'ShPL_Adoption.Total_ShPP_Weeks_Available', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 37, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 17, @tabShPL_Adoption, @colA_SAP_Weeks_Paid, 64, 0, 1720, 4200, @screenColumnHeight, 620, 'ShPL_Adoption.SAP_Weeks_Paid', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 36, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 18, @tabShPL_Adoption, @colA_SAP_Curtailment_Date, 64, 0, 1240, 4200, @screenColumnHeight, 1755, 'ShPL_Adoption.SAP_Curtailment_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 35, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 19, @tabShPL_Adoption, @colA_Placement_Date, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPL_Adoption.Placement_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 34, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 20, @tabShPL_Adoption, @colA_Main_Other_Adopter, 2, 0, 285, 4200, @screenColumnHeight, 1980, 'ShPL_Adoption.Main_Other_Adopter', -2147483643, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 33, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 2625, 'Partner''s NI Number :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 28, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 2, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2610, 'Partner''s Forename :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 29, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 3, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2400, 'Partner''s Surname :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 30, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 4, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2190, 'Partner''s Address :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 31, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 5, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 1320, 'Postcode :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 32, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 6, @tabShPL_Adoption, @colA_No_NI_Number_Declaration, 1, 0, 4120, 4200, @screenLabelHeight, 3500, 'No NI Number Declaration', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 27, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 2, 7, @tabShPL_Adoption, @colA_Partner_NI_Number, 64, 0, 3640, 4200, @screenColumnHeight, 1600, 'ShPL_Adoption.Partner_NI_Number', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 26, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 8, @tabShPL_Adoption, @colA_Partner_Postcode, 64, 0, 3160, 4200, @screenColumnHeight, 1400, 'ShPL_Adoption.Partner_Postcode', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 25, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 9, @tabShPL_Adoption, @colA_Partner_Address_4, 64, 0, 2680, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Address_4', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 24, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 10, @tabShPL_Adoption, @colA_Partner_Address_3, 64, 0, 2200, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Address_3', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 23, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 11, @tabShPL_Adoption, @colA_Partner_Address_2, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Address_2', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 22, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 12, @tabShPL_Adoption, @colA_Partner_Address_1, 64, 0, 1240, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Address_1', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 21, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 13, @tabShPL_Adoption, @colA_Partner_Surname, 64, 0, 760, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Surname', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 20, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 14, @tabShPL_Adoption, @colA_Partner_Forename, 64, 0, 280, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Forename', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 19, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 1, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 3600, 'Evidence from Adoption Agency :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 15, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 2, NULL, 0, 256, 0, 2720, 280, @screenLabelHeight, 1605, 'Postcode :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 16, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 3, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 3270, 'Partner''s Employer Address :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 17, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 4, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 3150, 'Partner''s Employer Name :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 18, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 5, @tabShPL_Adoption, @colA_Evidence_from_Adoption_Agency, 8, 0, 3240, 4200, 990, 990, 'ShPL_Adoption.Evidence_from_Adoption_Agency', NULL, NULL, NULL, NULL, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 14, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 6, @tabShPL_Adoption, @colA_Partner_Employer_Postcode, 64, 0, 2680, 4200, @screenColumnHeight, 1400, 'ShPL_Adoption.Partner_Employer_Postcode', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 13, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 7, @tabShPL_Adoption, @colA_Partner_Employer_Address_4, 64, 0, 2200, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Employer_Address_4', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 12, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 8, @tabShPL_Adoption, @colA_Partner_Employer_Address_3, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Employer_Address_3', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 11, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 9, @tabShPL_Adoption, @colA_Partner_Employer_Address_2, 64, 0, 1240, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Employer_Address_2', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 10, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 10, @tabShPL_Adoption, @colA_Partner_Employer_Address_1, 64, 0, 760, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Employer_Address_1', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 11, @tabShPL_Adoption, @colA_Partner_Employer_Name, 64, 0, 280, 4200, @screenColumnHeight, 4200, 'ShPL_Adoption.Partner_Employer_Name', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 1, NULL, 0, 256, 0, 640, 640, @screenLabelHeight * 3, 10170, 'That he/she is the main adopter of the child or the partner of the main adopter, that entitlement criteria for ShPP is satisfied and he/she agrees to inform the company immediately if conditions for entitlement to ShPP cease to be met.', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 2, NULL, 0, 256, 0, 2200, 640, @screenLabelHeight * 4, 10170, 'That he/she has at least 26 weeks employment (employed or self-employed) out of the 66 weeks prior to the 15th week before the relevant matching week, has average earnings of at least £30 during at least 13 of the 66 weeks prior to the relevant week, has curtailed SAP and consents to the employee''s claim to ShPP (as above).', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 3, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 3330, 'Date Notification Received :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 4, @tabShPL_Adoption, @colA_Date_Notification_Received, 64, 0, 3640, 4200, @screenColumnHeight, 1755, 'ShPL_Adoption.Date_Notification_Received', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 5, @tabShPL_Adoption, @colA_Declaration_from_Other_Adopter, 1, 0, 1880, 280, @screenLabelHeight, 3900, 'Declaration from Other Adopter', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 4, 6, @tabShPL_Adoption, @colA_Declaration_from_Employee, 1, 0, 320, 280, @screenLabelHeight, 3495, 'Declaration from Employee', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 5, 1, @tabShPL_Adoption, @colA_Notes, 64, 0, 280, 280, 4245, 10600, 'ShPL_Adoption.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'Eligibility')
			, (@screenID, 2, 'Partner''s Details')
			, (@screenID, 3, 'Supplementary Evidence')
			, (@screenID, 4, 'Declaration')
			, (@screenID, 5, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrPersID, @screenID);
	
	SET @scrAdoptID = @screenID;

	SET @cnameShPLA_ID = 'ID_' + CONVERT(varchar(3), @tabShPL_Adoption);

	/* ------------------------------------------------------------ */
	/* Create Shared Parental Leave (Adoption) Leave Requests Table */
	/* ------------------------------------------------------------ */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPLA_Leave_Requests OUTPUT, 'ShPLA_Leave_Requests', 2, @islocked, 'FB62A23B-67B5-4C4D-AD10-5D743AE6FFAD';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_ID OUTPUT, @tabShPLA_Leave_Requests, 'ID', 4, NULL, 0, 0, 0, '', @islocked, '7D4A96FB-DF3C-4F58-B054-58FE47DA2788', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_ID_A OUTPUT, @tabShPLA_Leave_Requests, @cnameShPLA_ID, 4, NULL, 0, 0, 0, '', @islocked, 'F42CA813-EFE7-496D-AAAE-C9B4283169F5', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Date_of_Request OUTPUT, @tabShPLA_Leave_Requests, 'Date_of_Request', 11, 'Enter date of ShPL request (mandatory)', 0, 0, 0, '', @islocked, '0D6C80BA-86DB-4E86-AF8A-2F767FE8C0E3', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Date_Requested_From OUTPUT, @tabShPLA_Leave_Requests, 'Date_Requested_From', 11, 'Enter date SHPL requested from (mandatory)', 0, 0, 0, '', @islocked, '82AE79E3-1DE2-483A-AF07-B292B03524BF', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Date_Requested_To OUTPUT, @tabShPLA_Leave_Requests, 'Date_Requested_To', 11, 'Enter date SHPL requested to (mandatory)', 0, 0, 0, '', @islocked, 'FFC3E480-2751-4C1D-8123-6405FD3B9B3E', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Binding_Request OUTPUT, @tabShPLA_Leave_Requests, 'Binding_Request', -7, 'Check box if request from employee is binding', 0, 0, 0, 'FALSE', @islocked, '877CD8FD-6EA6-4A86-B006-706F9D6D7405', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Consent_from_Other_Adopter OUTPUT, @tabShPLA_Leave_Requests, 'Consent_from_Other_Adopter', -7, 'Check box when consent received from other adopter', 0, 0, 0, 'FALSE', @islocked, '92200810-FDFD-4AFD-8F28-966130AA9AA0', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Request_Cancelled OUTPUT, @tabShPLA_Leave_Requests, 'Request_Cancelled', -7, 'Check box if rquest has been cancelled', 0, 0, 0, 'FALSE', @islocked, '853CD15F-3AC9-4EFA-B87F-ACE66AB52AA4', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_ShPP_Weeks OUTPUT, @tabShPLA_Leave_Requests, 'ShPP_Weeks', 4, '', 2, 0, 0, '0', @islocked, 'BE17BD92-3078-4A28-9F85-0207F02FAD6D', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAR_Notes OUTPUT, @tabShPLA_Leave_Requests, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, '1415FD8F-4501-4E0C-83A4-E6147FA0FF67', 0, 0, 0;

	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Date_of_Request', @tabShPLA_Leave_Requests, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colAR_Date_of_Request, 'F', 0, 1)
			, (@orderID, @colAR_Date_Requested_From, 'F', 1, 1)
			, (@orderID, @colAR_Date_Requested_To, 'F', 2, 1)
			, (@orderID, @colAR_ShPP_Weeks, 'F', 3, 1)
			, (@orderID, @colAR_Binding_Request, 'F', 4, 1)
			, (@orderID, @colAR_Consent_from_Other_Adopter, 'F', 5, 1)
			, (@orderID, @colAR_Request_Cancelled, 'F', 6, 1)
			, (@orderID, @colAR_Date_of_Request, 'O', 1, 0)
			, (@orderID, @colAR_Date_Requested_From, 'O', 2, 0);

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPLA_Leave_Requests;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'ShPL (Adoption) Leave Requests', @tabShPLA_Leave_Requests, 0, 4770, 11370, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 1755, 'ShPP Weeks :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2430, 'Date Requested To :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 10, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2745, 'Date Requested From :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 11, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2175, 'Date of Request :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 12, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, @tabShPLA_Leave_Requests, @colAR_Request_Cancelled, 1, 0, 3200, 4200, @screenLabelHeight, 2490, 'Request Cancelled', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 6, @tabShPLA_Leave_Requests, @colAR_Consent_from_Other_Adopter, 1, 0, 2720, 4200, @screenLabelHeight, 3645, 'Consent from Other Adopter', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, @tabShPLA_Leave_Requests, @colAR_Binding_Request, 1, 0, 2240, 4200, @screenLabelHeight, 4185, 'Binding Request from Employee', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, @tabShPLA_Leave_Requests, @colAR_ShPP_Weeks, 64, 0, 1720, 4200, @screenColumnHeight, 620, 'ShPLA_Requests.ShPP_Weeks', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 9, @tabShPLA_Leave_Requests, @colAR_Date_Requested_To, 64, 0, 1240, 4200, @screenColumnHeight, 1755, 'ShPLA_Requests.Date_Requested_To', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 10, @tabShPLA_Leave_Requests, @colAR_Date_Requested_From, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPLA_Requests.Date_Requested_From', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 11, @tabShPLA_Leave_Requests, @colAR_Date_of_Request, 64, 0, 280, 4200, @screenColumnHeight, 1755, 'ShPLA_Requests.Date_of_Request', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, @tabShPLA_Leave_Requests, @colAR_Notes, 64, 0, 280, 280, 3270, 10600, 'ShPLA_Requests.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'Leave Request')
			, (@screenID, 2, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrAdoptID, @screenID);

	/* -------------------------------------------------------- */
	/* Create Shared Parental Leave (Adoption) SPLIT Days Table */
	/* -------------------------------------------------------- */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPLA_SPLIT_Days OUTPUT, 'ShPLA_SPLIT_Days', 2, @islocked, '480733E5-0F3D-4418-A576-C1FB484B541E';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_ID OUTPUT, @tabShPLA_SPLIT_Days, 'ID', 4, NULL, 0, 0, 0, '', @islocked, '23500B15-47D1-4C2A-A794-98BBA9B87E64', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_ID_A OUTPUT, @tabShPLA_SPLIT_Days, @cnameShPLA_ID, 4, NULL, 0, 0, 0, '', @islocked, 'DAB40CFC-18DB-433A-99FF-53C751F6419B', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_Start_Date OUTPUT, @tabShPLA_SPLIT_Days, 'Start_Date', 11, 'Enter start SPLIT day (mandatory and unique)', 0, 0, 0, '', @islocked, '8F53E032-52FB-4382-9C6E-858CABFD1AE6', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_End_Date OUTPUT, @tabShPLA_SPLIT_Days, 'End_Date', 11, 'Enter end SPLIT day (mandatory)', 0, 0, 0, '', @islocked, '8189ABE7-B026-46F5-851B-B277754E896A', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_Reason OUTPUT, @tabShPLA_SPLIT_Days, 'Reason', 12, 'Enter reason for SPLIT day(s) (mandatory)', 30, 0, 0, '', @islocked, 'B4819866-7084-4B89-A7F8-140571F9F231', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_Notes OUTPUT, @tabShPLA_SPLIT_Days, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, 'F64E712F-1744-458E-A109-F3127D63052A', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colAS_SPLIT_Days OUTPUT, @tabShPLA_SPLIT_Days, 'SPLIT_Days', 4, '', 2, 0, 0, '0', @islocked, '6B3BD046-F184-4858-A982-DC8DE5369E2F', 0, 0, 0;

	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Start_Date', @tabShPLA_SPLIT_Days, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colAS_Start_Date, 'F', 0, 1)
			, (@orderID, @colAS_End_Date, 'F', 1, 1)
			, (@orderID, @colAS_SPLIT_Days, 'F', 2, 1)
			, (@orderID, @colAS_Reason, 'F', 3, 1)
			, (@orderID, @colAS_Start_Date, 'O', 1, 0)

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPLA_SPLIT_Days;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'ShPL (Adoption) SPLIT Days', @tabShPLA_SPLIT_Days, 0, 3360, 11385, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 1665, 'Start Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 1335, 'End Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 1185, 'Reason :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 1665, 'SPLIT Days :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, @tabShPLA_SPLIT_Days, @colAS_SPLIT_Days, 64, 0, 1240, 4200, @screenColumnHeight, 620, 'ShPLA_SPLIT_Days.SPLIT_Days', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, NULL, 1, NULL, NULL, 0)
			, (@screenID, 1, 6, @tabShPLA_SPLIT_Days, @colAS_Reason, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPLA_SPLIT_Days.Reason', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, @tabShPLA_SPLIT_Days, @colAS_End_Date, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPLA_SPLIT_Days.End_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, @tabShPLA_SPLIT_Days, @colAS_Start_Date, 64, 0, 280, 4200, @screenColumnHeight, 1755, 'ShPLA_SPLIT_Days.Start_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, @tabShPLA_SPLIT_Days, @colAS_Notes, 64, 0, 280, 280, 1860, 10600, 'ShPLA_SPLIT_Days.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'SPLIT Days')
			, (@screenID, 2, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrAdoptID, @screenID);

	/* --------------------------------------------- */
	/* Create Shared Parental Leave (Birth) Table */
	/* --------------------------------------------- */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPL_Birth OUTPUT, 'ShPL_Birth', 2, @islocked, '936EADDE-DD2D-4D80-B292-A09778EEF101';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_ID OUTPUT, @tabShPL_Birth, 'ID', 4, NULL, 0, 0, 0, '', @islocked, 'B2E2ECF0-66A7-42E2-8CAF-A0198AE7F963', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_ID_Pers OUTPUT, @tabShPL_Birth, @cnameID_Pers, 4, NULL, 0, 0, 0, '', @islocked, '448F937A-D02A-471F-BE78-5574D57E5B8F', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Mother_Father_Partner OUTPUT, @tabShPL_Birth, 'Mother_Father_Partner', 12, 'Select if the employee is the mother or father/partner', 14, 0, 0, '', @islocked, 'C914E92C-078A-497F-ABD7-65F9086A5BBC', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Expected_Birth_Date OUTPUT, @tabShPL_Birth, 'Expected_Birth_Date', 11, 'Enter the expected birth (MATB1) date', 0, 0, 0, '', @islocked, 'AE967475-779B-4A1C-8420-95F4C9337226', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_SMP_Curtailment_Date OUTPUT, @tabShPL_Birth, 'SMP_Curtailment_Date', 11, 'Enter the SMP curtailment date (mandatory)', 0, 0, 0, '', @islocked, '77113159-38B9-4755-AC83-7D1976421BD3', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Forename OUTPUT, @tabShPL_Birth, 'Partner_Forename', 12, 'Enter partner''s forename (mandatory)', 30, 0, 0, '', @islocked, '330C7BCA-673E-4FAC-B3D9-E470740E5E84', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Surname OUTPUT, @tabShPL_Birth, 'Partner_Surname', 12, 'Enter partner''s surname (mandatory)', 30, 0, 0, '', @islocked, 'D4D05655-F85B-47D4-9D65-FE0780A591BA', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Address_1 OUTPUT, @tabShPL_Birth, 'Partner_Address_1', 12, 'Enter 1st line of partner''s address (mandatory)', 30, 0, 0, '', @islocked, '196CC640-D0D0-43E5-A4E7-918B3448276C', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Address_2 OUTPUT, @tabShPL_Birth, 'Partner_Address_2', 12, 'Enter 2nd line of partner''s address (mandatory)', 30, 0, 0, '', @islocked, '05A8E9CC-8F19-4363-A49A-18EC217DE15B', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Address_3 OUTPUT, @tabShPL_Birth, 'Partner_Address_3', 12, 'Enter 3rd line of partner''s address', 30, 0, 0, '', @islocked, 'BEAD58BB-7874-4274-8963-8C7DA4E6D99C', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Address_4 OUTPUT, @tabShPL_Birth, 'Partner_Address_4', 12, 'Enter 4th line of partner''s address', 30, 0, 0, '', @islocked, '6789B26D-C581-4DCE-83FE-D3B4D4EB971E', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Postcode OUTPUT, @tabShPL_Birth, 'Partner_Postcode', 12, 'Enter partner''s postcode', 8, 0, 0, '', @islocked, '52E6DC44-40A2-41E7-BB2D-CD082147BAB6', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_NI_Number OUTPUT, @tabShPL_Birth, 'Partner_NI_Number', 12, 'Enter partner''s NI Number', 9, 0, 0, '', @islocked, '78BC58A1-107E-49C3-8AC0-ABE71E3FDEC3', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_No_NI_Number_Declaration OUTPUT, @tabShPL_Birth, 'No_NI_Number_Declaration', -7, 'Check box if no NI Number exists', 0, 0, 0, 'FALSE', @islocked, '7569E7DF-70DA-41D4-BCD4-C5DD43C0C06E', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Date_Notification_Received OUTPUT, @tabShPL_Birth, 'Date_Notification_Received', 11, 'Enter date notification of ShPL received (mandatory)', 0, 0, 0, '', @islocked, 'CE555756-7429-4728-AACD-8D2B98D38C8D', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Intended_ShPL_Start_Date OUTPUT, @tabShPL_Birth, 'Intended_ShPL_Start_Date', 11, 'Enter intended ShPL start date', 0, 0, 0, '', @islocked, '703CE528-F3A1-4577-88E2-0FBBB21522DA', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Intended_ShPL_End_Date OUTPUT, @tabShPL_Birth, 'Intended_ShPL_End_Date', 11, 'Enter intended ShPL end date', 0, 0, 0, '', @islocked, 'C9A0FA74-6B6A-4AE9-B9DE-52ECC1A16690', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_SMP_Weeks_Paid OUTPUT, @tabShPL_Birth, 'SMP_Weeks_Paid', 4, 'Enter number of SMP weeks paid to the mother', 2, 0, 0, '0', @islocked, 'AE7E4903-5308-413B-9061-101FD981630F', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Total_ShPP_Weeks_Available OUTPUT, @tabShPL_Birth, 'Total_ShPP_Weeks_Available', 4, '', 2, 0, 0, '39', @islocked, 'C3B28764-BA05-45AD-AE7B-F703062A6A62', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_ShPP_Weeks_Employee OUTPUT, @tabShPL_Birth, 'ShPP_Weeks_Employee', 4, 'Enter number of ShPP weeks to be claimed by employee (mandatory)', 2, 0, 0, '0', @islocked, 'B217515D-7705-402C-92D6-5DDC14DA58AA', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_ShPP_Weeks_Partner OUTPUT, @tabShPL_Birth, 'ShPP_Weeks_Partner', 4, 'Enter number of ShPP weeks to be claimed by employee''s partner (mandatory)', 2, 0, 0, '0', @islocked, '2D66F1A2-B786-493D-A530-D56FB0194EE0', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Name OUTPUT, @tabShPL_Birth, 'Partner_Employer_Name', 12, 'Enter name of partner''s employer', 30, 0, 0, '', @islocked, '53348182-9D99-48ED-8719-897B9292A145', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Address_1 OUTPUT, @tabShPL_Birth, 'Partner_Employer_Address_1', 12, 'Enter 1st line of partner''s employer address', 30, 0, 0, '', @islocked, '8D317288-5697-40C9-9D0C-B71353056D8B', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Address_2 OUTPUT, @tabShPL_Birth, 'Partner_Employer_Address_2', 12, 'Enter 2nd line of partner''s employer address', 30, 0, 0, '', @islocked, 'D04EB23A-12A3-44C9-9C43-926EE2745DB1', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Address_3 OUTPUT, @tabShPL_Birth, 'Partner_Employer_Address_3', 12, 'Enter 3rd line of partner''s employer address', 30, 0, 0, '', @islocked, '1E5278EA-A8B4-4406-A634-17B842973C13', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Address_4 OUTPUT, @tabShPL_Birth, 'Partner_Employer_Address_4', 12, 'Enter 4th line of partner''s employer address', 30, 0, 0, '', @islocked, 'E706FC9A-EB93-4C1A-94FC-2483E5BD08D6', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Partner_Employer_Postcode OUTPUT, @tabShPL_Birth, 'Partner_Employer_Postcode', 12, 'Enter partner''s employer postcode', 8, 0, 0, '', @islocked, 'D40A3A2E-8EB7-4F85-8AD3-8C25D5D25567', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Evidence_of_Birth OUTPUT, @tabShPL_Birth, 'Evidence_of_Birth', -4, 'Click button to link/access document', 0, 0, 0, '', @islocked, 'C4FA5DBD-954B-4A08-B47C-D96D28E436AC', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Declaration_from_Employee OUTPUT, @tabShPL_Birth, 'Declaration_from_Employee', -7, 'Check box when declaration from employee received', 0, 0, 0, 'FALSE', @islocked, 'DD7C1FD6-02F7-401C-9782-33C8720670F7', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Declaration_from_Other_Parent OUTPUT, @tabShPL_Birth, 'Declaration_from_Other_Parent', -7, 'Check box when declaration from other parent received', 0, 0, 0, 'FALSE', @islocked, '4BFC659C-1838-4CA8-869A-881EE81F0059', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Notes OUTPUT, @tabShPL_Birth, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, '5EC666F8-69EB-450B-B6E6-83EEEA9A6EF5', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Payroll_Company_Code OUTPUT, @tabShPL_Birth, 'Payroll_Company_Code', 12, '', 2, 0, 0, '', @islocked, '7B213BD8-86AC-4998-8ADB-A12D12CDD4D1', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Staff_Number OUTPUT, @tabShPL_Birth, 'Staff_Number', 12, '', 8, 0, 0, '', @islocked, '2ACE9125-4011-4BD5-A897-23273EF756E8', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Full_Name OUTPUT, @tabShPL_Birth, 'Full_Name', 12, '', 40, 0, 0, '', @islocked, 'B0CC4C50-A1EF-471A-8409-1A851FAA98EC', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Trigger_to_Payroll OUTPUT, @tabShPL_Birth, 'Trigger_to_Payroll', -7, '', 0, 0, 0, 'FALSE', @islocked, '7C00CD51-8B90-48CD-8C57-D110D30A6FB7', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Total_SPLIT_Days OUTPUT, @tabShPL_Birth, 'Total_SPLIT_Days', 4, '', 2, 0, 0, '0', @islocked, '9A78E76B-3DAB-4BE4-9AF7-B5E96E2DCB57', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Actual_Birth_Date OUTPUT, @tabShPL_Birth, 'Actual_Birth_Date', 11, 'Enter the actual birth date', 0, 0, 0, '', @islocked, '2819DE27-588C-48CB-8DCE-A6563F7EBCD6', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colB_Location_of_Birth OUTPUT, @tabShPL_Birth, 'Location_of_Birth', 12, 'Enter location of birth', 30, 0, 0, '', @islocked, '63C0CF05-F782-45D5-BD13-9C5A2BB7087A', 0, 0, 0;

	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Date_Notification_Received', @tabShPL_Birth, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colB_Date_Notification_Received, 'F', 0, 1)
			, (@orderID, @colB_Mother_Father_Partner, 'F', 1, 1)
			, (@orderID, @colB_Expected_Birth_Date, 'F', 2, 1)
			, (@orderID, @colB_ShPP_Weeks_Employee, 'F', 3, 1)
			, (@orderID, @colB_ShPP_Weeks_Partner, 'F', 4, 1)
			, (@orderID, @colB_Intended_ShPL_Start_Date, 'F', 5, 1)
			, (@orderID, @colB_Intended_ShPL_End_Date, 'F', 6, 1)
			, (@orderID, @colB_Date_Notification_Received, 'O', 1, 0);

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPL_Birth;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'Statutory Shared Parental Leave (Birth)', @tabShPL_Birth, 0, 6240, 11400, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 3280, 5080, @screenLabelHeight * 2, 3165, '(Must add up to Total ShPP Weeks Available)', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 84, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 2720, 5080, @screenLabelHeight, 2625, '(39 - SMP Weeks Paid)', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 80, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 4640, 280, @screenLabelHeight, 3030, 'Intended ShPL End Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 54, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 4160, 280, @screenLabelHeight, 2940, 'Intended ShPL Start Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 53, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 3525, 'Weeks Other Parent Claiming :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 52, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 6, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 3225, 'Weeks Employee Claiming :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 51, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, NULL, 0, 256, 0, 2720, 280, @screenLabelHeight, 3390, 'Total ShPP Weeks Available :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 50, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, NULL, 0, 256, 0, 2240, 280, @screenLabelHeight, 3255, 'SMP Weeks Paid to Mother :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 49, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 9, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 2820, 'SMP Curtailment Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 48, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 10, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2610, 'Expected Birth Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 47, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 11, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2655, 'The Employee is the :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 46, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 12, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2100, 'Actual Birth Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 45, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 13, @tabShPL_Birth, @colB_Intended_ShPL_End_Date, 64, 0, 4600, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.Intended_ShPL_End_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 44, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 14, @tabShPL_Birth, @colB_Intended_ShPL_Start_Date, 64, 0, 4120, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.Intended_ShPL_Start_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 43, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 15, @tabShPL_Birth, @colB_ShPP_Weeks_Partner, 64, 0, 3640, 4200, @screenColumnHeight, 620, 'ShPL_Birth.ShPP_Weeks_Partner', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 42, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 16, @tabShPL_Birth, @colB_ShPP_Weeks_Employee, 64, 0, 3160, 4200, @screenColumnHeight, 620, 'ShPL_Birth.ShPP_Weeks_Employee', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 41, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 17, @tabShPL_Birth, @colB_Total_ShPP_Weeks_Available, 64, 0, 2680, 4200, @screenColumnHeight, 620, 'ShPL_Birth.Total_ShPP_Weeks_Available', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 40, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 18, @tabShPL_Birth, @colB_SMP_Weeks_Paid, 64, 0, 2200, 4200, @screenColumnHeight, 620, 'ShPL_Birth.SMP_Weeks_Paid', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 39, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 19, @tabShPL_Birth, @colB_SMP_Curtailment_Date, 64, 0, 1720, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.SMP_Curtailment_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 38, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 20, @tabShPL_Birth, @colB_Actual_Birth_Date, 64, 0, 1240, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.Actual_Birth_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 37, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 21, @tabShPL_Birth, @colB_Expected_Birth_Date, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.Expected_Birth_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 36, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 22, @tabShPL_Birth, @colB_Mother_Father_Partner, 2, 0, 285, 4200, @screenColumnHeight, 1980, 'ShPL_Birth.Mother_Father_Partner', -2147483643, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 35, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 1320, 'Postcode :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 30, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 2, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2190, 'Partner''s Address :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 31, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 3, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2400, 'Partner''s Surname :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 32, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 4, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2610, 'Partner''s Forename :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 33, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 5, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 2625, 'Partner''s NI Number :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 34, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 6, @tabShPL_Birth, @colB_No_NI_Number_Declaration, 1, 0, 4120, 4200, @screenLabelHeight, 3500, 'No NI Number Declaration', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 29, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 2, 7, @tabShPL_Birth, @colB_Partner_NI_Number, 64, 0, 3640, 4200, @screenColumnHeight, 1600, 'ShPL_Birth.Partner_NI_Number', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 28, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 8, @tabShPL_Birth, @colB_Partner_Postcode, 64, 0, 3160, 4200, @screenColumnHeight, 1400, 'ShPL_Birth.Partner_Postcode', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 27, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 9, @tabShPL_Birth, @colB_Partner_Address_4, 64, 0, 2680, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Address_4', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 26, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 10, @tabShPL_Birth, @colB_Partner_Address_3, 64, 0, 2200, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Address_3', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 25, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 11, @tabShPL_Birth, @colB_Partner_Address_2, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Address_2', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 24, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 12, @tabShPL_Birth, @colB_Partner_Address_1, 64, 0, 1240, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Address_1', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 23, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 13, @tabShPL_Birth, @colB_Partner_Surname, 64, 0, 760, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Surname', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 22, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 14, @tabShPL_Birth, @colB_Partner_Forename, 64, 0, 280, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Forename', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 21, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 1, NULL, 0, 256, 0, 4360, 280, @screenLabelHeight, 2895, 'Location of Child''s Birth :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 16, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 2, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 3150, 'Partner''s Employer Name :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 17, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 3, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 3270, 'Partner''s Employer Address :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 18, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 4, NULL, 0, 256, 0, 2720, 280, @screenLabelHeight, 1605, 'Postcode :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 19, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 5, NULL, 0, 256, 0, 3200, 280, @screenLabelHeight, 2925, 'Evidence of Child''s Birth :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 20, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 6, @tabShPL_Birth, @colB_Location_of_Birth, 64, 0, 4320, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Location_of_Birth', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 15, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 7, @tabShPL_Birth, @colB_Evidence_of_Birth, 8, 0, 3240, 4200, 990, 990, 'ShPL_Birth.Evidence_of_Birth', NULL, NULL, NULL, NULL, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 14, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 8, @tabShPL_Birth, @colB_Partner_Employer_Postcode, 64, 0, 2680, 4200, @screenColumnHeight, 1400, 'ShPL_Birth.Partner_Employer_Postcode', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 13, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 9, @tabShPL_Birth, @colB_Partner_Employer_Address_4, 64, 0, 2200, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Employer_Address_4', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 12, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 10, @tabShPL_Birth, @colB_Partner_Employer_Address_3, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Employer_Address_3', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 11, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 11, @tabShPL_Birth, @colB_Partner_Employer_Address_2, 64, 0, 1240, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Employer_Address_2', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 10, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 12, @tabShPL_Birth, @colB_Partner_Employer_Address_1, 64, 0, 760, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Employer_Address_1', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 3, 13, @tabShPL_Birth, @colB_Partner_Employer_Name, 64, 0, 280, 4200, @screenColumnHeight, 4200, 'ShPL_Birth.Partner_Employer_Name', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 1, NULL, 0, 256, 0, 3680, 280, @screenLabelHeight, 3330, 'Date Notification Received :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 2, NULL, 0, 256, 0, 2200, 640, @screenLabelHeight * 4, 10410, 'That he/she has at least 26 weeks employment (employed or self-employed) out of the 66 weeks prior to the 15th week before the expected week of birth, has average earnings of at least £30 during at least 13 of the 66 weeks prior to the relevant week, has curtailed SMP/MA and consents to the employee''s claim to ShPP (as above).', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 3, NULL, 0, 256, 0, 640, 640, @screenLabelHeight * 3, 10230, 'That he/she is the mother or father of the child or the partner of the mother, that entitlement criteria for ShPP is satisfied and he/she agrees to inform the company immediately if conditions for entitlement to ShPP cease to be met.', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 4, @tabShPL_Birth, @colB_Date_Notification_Received, 64, 0, 3640, 4200, @screenColumnHeight, 1755, 'ShPL_Birth.Date_Notification_Received', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 4, 5, @tabShPL_Birth, @colB_Declaration_from_Other_Parent, 1, 0, 1880, 280, @screenLabelHeight, 3900, 'Declaration from Other Parent', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 4, 6, @tabShPL_Birth, @colB_Declaration_from_Employee, 1, 0, 320, 280, @screenLabelHeight, 3495, 'Declaration from Employee', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 5, 1, @tabShPL_Birth, @colB_Notes, 64, 0, 280, 280, 4665, 10600, 'ShPL_Birth.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'Eligibility')
			, (@screenID, 2, 'Partner''s Details')
			, (@screenID, 3, 'Supplementary Evidence')
			, (@screenID, 4, 'Declaration')
			, (@screenID, 5, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrPersID, @screenID);
	
	SET @scrBirthID = @screenID;

	SET @cnameShPLB_ID = 'ID_' + CONVERT(varchar(3), @tabShPL_Birth);

	/* ------------------------------------------------------------ */
	/* Create Shared Parental Leave (Birth) Leave Requests Table */
	/* ------------------------------------------------------------ */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPLB_Leave_Requests OUTPUT, 'ShPLB_Leave_Requests', 2, @islocked, 'FBADC2FC-A1A8-4D78-9D06-72425D4B5CA3';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_ID OUTPUT, @tabShPLB_Leave_Requests, 'ID', 4, NULL, 0, 0, 0, '', @islocked, '3000FB38-43A2-40EF-A743-02C4AE15A5A2', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_ID_B OUTPUT, @tabShPLB_Leave_Requests, @cnameShPLB_ID, 4, NULL, 0, 0, 0, '', @islocked, '819F1CB2-7D4F-4BC7-BEEA-D47F6E78B1E8', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Date_of_Request OUTPUT, @tabShPLB_Leave_Requests, 'Date_of_Request', 11, 'Enter date of ShPL request (mandatory)', 0, 0, 0, '', @islocked, 'D5BA821B-C126-472A-872A-6182AE2B5A36', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Date_Requested_From OUTPUT, @tabShPLB_Leave_Requests, 'Date_Requested_From', 11, 'Enter date SHPL requested from (mandatory)', 0, 0, 0, '', @islocked, '5E280AB0-5601-4923-ADD1-E4730B5BB817', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Date_Requested_To OUTPUT, @tabShPLB_Leave_Requests, 'Date_Requested_To', 11, 'Enter date SHPL requested to (mandatory)', 0, 0, 0, '', @islocked, '0F9B4626-536E-4576-AC25-CD475428476A', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Binding_Request OUTPUT, @tabShPLB_Leave_Requests, 'Binding_Request', -7, 'Check box if request from employee is binding', 0, 0, 0, 'FALSE', @islocked, '38F5D0E9-328F-48D3-B232-596EEC203664', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Consent_from_Other_Parent OUTPUT, @tabShPLB_Leave_Requests, 'Consent_from_Other_Parent', -7, 'Check box when consent received from other parent', 0, 0, 0, 'FALSE', @islocked, '383F5F0A-5A35-4F56-87DD-D8C64FFF88D7', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Request_Cancelled OUTPUT, @tabShPLB_Leave_Requests, 'Request_Cancelled', -7, 'Check box if rquest has been cancelled', 0, 0, 0, 'FALSE', @islocked, '1A01FCB8-0E45-4EE2-8C53-92FA82909C9E', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_ShPP_Weeks OUTPUT, @tabShPLB_Leave_Requests, 'ShPP_Weeks', 4, '', 2, 0, 0, '0', @islocked, 'F49BD31F-9E56-4418-86CD-6C914A435963', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBR_Notes OUTPUT, @tabShPLB_Leave_Requests, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, '8B9874B9-46AB-4561-B40C-09525E5B1792', 0, 0, 0;
	
	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Date_of_Request', @tabShPLB_Leave_Requests, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colBR_Date_of_Request, 'F', 0, 1)
			, (@orderID, @colBR_Date_Requested_From, 'F', 1, 1)
			, (@orderID, @colBR_Date_Requested_To, 'F', 2, 1)
			, (@orderID, @colBR_ShPP_Weeks, 'F', 3, 1)
			, (@orderID, @colBR_Binding_Request, 'F', 4, 1)
			, (@orderID, @colBR_Consent_from_Other_Parent, 'F', 5, 1)
			, (@orderID, @colBR_Request_Cancelled, 'F', 6, 1)
			, (@orderID, @colBR_Date_of_Request, 'O', 1, 0)
			, (@orderID, @colBR_Date_Requested_From, 'O', 2, 0);

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPLB_Leave_Requests;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'ShPL (Birth) Leave Requests', @tabShPLB_Leave_Requests, 0, 4770, 11370, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 2175, 'Date of Request :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 2745, 'Date Requested From :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 10, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 2430, 'Date Requested To :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 11, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 1755, 'ShPP Weeks :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 12, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, @tabShPLB_Leave_Requests, @colBR_Request_Cancelled, 1, 0, 3200, 4200, @screenLabelHeight, 2490, 'Request Cancelled', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 6, @tabShPLB_Leave_Requests, @colBR_Consent_from_Other_Parent, 1, 0, 2720, 4200, @screenLabelHeight, 3645, 'Consent from Other Parent', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, @tabShPLB_Leave_Requests, @colBR_Binding_Request, 1, 0, 2240, 4200, @screenLabelHeight, 4185, 'Binding Request from Employee', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, 0, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, @tabShPLB_Leave_Requests, @colBR_ShPP_Weeks, 64, 0, 1720, 4200, @screenColumnHeight, 620, 'ShPLB_Leave_Requests.ShPP_Weeks', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 9, @tabShPLB_Leave_Requests, @colBR_Date_Requested_To, 64, 0, 1240, 4200, @screenColumnHeight, 1755, 'ShPLB_Leave_Requests.Date_Requested_To', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 10, @tabShPLB_Leave_Requests, @colBR_Date_Requested_From, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPLB_Leave_Requests.Date_Requested_From', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 11, @tabShPLB_Leave_Requests, @colBR_Date_of_Request, 64, 0, 280, 4200, @screenColumnHeight, 1755, 'ShPLB_Leave_Requests.Date_of_Request', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, @tabShPLB_Leave_Requests, @colBR_Notes, 64, 0, 280, 280, 3270, 10600, 'ShPLB_Leave_Requests.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'Leave Request')
			, (@screenID, 2, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrBirthID, @screenID);

	/* -------------------------------------------------------- */
	/* Create Shared Parental Leave (Birth) SPLIT Days Table */
	/* -------------------------------------------------------- */
	EXECUTE [dbo].[spshpl_scriptnewtable] @tabShPLB_SPLIT_Days OUTPUT, 'ShPLB_SPLIT_Days', 2, @islocked, 'F4707181-965E-4833-B1CE-FDC93895F1D2';

	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_ID OUTPUT, @tabShPLB_SPLIT_Days, 'ID', 4, NULL, 0, 0, 0, '', @islocked, '5391D292-9671-40F8-86AC-E8A07DA2D89E', 1, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_ID_B OUTPUT, @tabShPLB_SPLIT_Days, @cnameShPLB_ID, 4, NULL, 0, 0, 0, '', @islocked, 'D586966D-CE6B-4DAB-AC15-B366220BA5A8', 0, 1, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_Start_Date OUTPUT, @tabShPLB_SPLIT_Days, 'Start_Date', 11, 'Enter start SPLIT day (mandatory and unique)', 0, 0, 0, '', @islocked, 'A378B069-48B2-43FC-8125-EB95EB7E2172', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_End_Date OUTPUT, @tabShPLB_SPLIT_Days, 'End_Date', 11, 'Enter end SPLIT day (mandatory)', 0, 0, 0, '', @islocked, '6242AFBE-3883-4426-82F9-B218601C0521', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_Reason OUTPUT, @tabShPLB_SPLIT_Days, 'Reason', 12, 'Enter reason for SPLIT day(s) (mandatory)', 30, 0, 0, '', @islocked, 'E0982187-39AA-4B3B-8055-A9BEA775FCC3', 0, 0, 1;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_Notes OUTPUT, @tabShPLB_SPLIT_Days, 'Notes', 12, 'Enter notes (multi-line text)', 2147483646, 0, 0, '', @islocked, '8928A110-9CF9-4474-90C7-3B4FF3FB4633', 0, 0, 0;
	EXECUTE [dbo].[spshpl_scriptnewcolumn] @colBS_SPLIT_Days OUTPUT, @tabShPLB_SPLIT_Days, 'SPLIT_Days', 4, '', 2, 0, 0, '0', @islocked, '40803A0D-00EC-4B0A-8357-E4AC25F6751A', 0, 0, 0;

	SELECT @orderID = MAX([OrderID]) + 1 FROM [dbo].[ASRSysOrders];

	INSERT INTO [dbo].[ASRSysOrders]
		([OrderID], [Name], [TableID], [Type])
		VALUES (@orderID, 'Start_Date', @tabShPLB_SPLIT_Days, 1);
	
	INSERT INTO [dbo].[ASRSysOrderItems]
		([OrderID], [ColumnID], [Type], [Sequence], [Ascending])
		VALUES (@orderID, @colBS_Start_Date, 'F', 0, 1)
			, (@orderID, @colBS_End_Date, 'F', 1, 1)
			, (@orderID, @colBS_SPLIT_Days, 'F', 2, 1)
			, (@orderID, @colBS_Reason, 'F', 3, 1)
			, (@orderID, @colBS_Start_Date, 'O', 1, 0)

	UPDATE [dbo].[tbsys_tables] SET [DefaultOrderID] = @orderID WHERE [TableID] = @tabShPLB_SPLIT_Days;

	EXEC dbo.spsys_setsystemsetting 'autoid', 'orders', @orderID;

	SELECT @screenID = MAX([ScreenID]) + 1 FROM [dbo].[ASRSysScreens];

	INSERT INTO [dbo].[ASRSysScreens]
		([ScreenID], [Name], [TableID], [OrderID], [Height], [Width], [PictureID], [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [GridX], [GridY], [AlignToGrid]
			, [DfltForeColour], [DfltFontName], [DfltFontSize], [DfltFontBold], [DfltFontItalic], [QuickEntry], [SSIntranet])
		VALUES (@screenID, 'ShPL (Birth) SPLIT Days', @tabShPLB_SPLIT_Days, 0, 3360, 11385, 0, @screenFontName, 8, 0, 0, 0, 0, 40, 40, 1, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysControls]
		([ScreenID], [PageNo], [ControlLevel], [TableID], [ColumnID], [ControlType], [ControlIndex], [TopCoord], [LeftCoord], [Height], [Width], [Caption], [BackColor], [ForeColor]
			, [FontName], [FontSize], [FontBold], [FontItalic], [FontStrikeThru], [FontUnderline], [PictureID], [DisplayType], [ContainerType], [ContainerIndex], [TabIndex], [BorderStyle]
			, [Alignment], [ReadOnly], [NavigateTo], [NavigateIn], [NavigateOnSave])
		VALUES (@screenID, 1, 1, NULL, 0, 256, 0, 320, 280, @screenLabelHeight, 1665, 'Start Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 6, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 2, NULL, 0, 256, 0, 800, 280, @screenLabelHeight, 1335, 'End Date :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 7, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 3, NULL, 0, 256, 0, 1760, 280, @screenLabelHeight, 1185, 'Reason :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 8, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 4, NULL, 0, 256, 0, 1280, 280, @screenLabelHeight, 1665, 'SPLIT Days :', -2147483633, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 9, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 5, @tabShPLB_SPLIT_Days, @colBS_SPLIT_Days, 64, 0, 1240, 4200, @screenColumnHeight, 620, 'ShPLB_SPLIT_Days.SPLIT_Days', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 3, NULL, NULL, 1, NULL, NULL, 0)
			, (@screenID, 1, 6, @tabShPLB_SPLIT_Days, @colBS_Reason, 64, 0, 1720, 4200, @screenColumnHeight, 4200, 'ShPLB_SPLIT_Days.Reason', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 4, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 7, @tabShPLB_SPLIT_Days, @colBS_End_Date, 64, 0, 760, 4200, @screenColumnHeight, 1755, 'ShPLB_SPLIT_Days.End_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 2, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 1, 8, @tabShPLB_SPLIT_Days, @colBS_Start_Date, 64, 0, 280, 4200, @screenColumnHeight, 1755, 'ShPLB_SPLIT_Days.Start_Date', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, 0)
			, (@screenID, 2, 1, @tabShPLB_SPLIT_Days, @colBS_Notes, 64, 0, 280, 280, 1860, 10600, 'ShPLB_SPLIT_Days.Notes', 16777215, 0, @screenFontName, @screenFontSize, 0, 0, 0, 0, NULL, NULL, NULL, NULL, 5, NULL, NULL, 0, NULL, NULL, 0);

	INSERT INTO [dbo].[ASRSysPageCaptions]
		([ScreenID], [PageIndexID], [Caption])
		VALUES (@screenID, 1, 'SPLIT Days')
			, (@screenID, 2, 'Notes');

	SELECT @hScreenID = MAX([ID]) + 1 FROM [dbo].[ASRSysHistoryScreens];

	INSERT INTO [dbo].[ASRSysHistoryScreens]
		([ID], [parentScreenID], [historyScreenID])
		VALUES (@hScreenID, @scrBirthID, @screenID);

	/* ---------------------- */
	/* Set extra column flags */
	/* ---------------------- */
	UPDATE [dbo].[tbsys_columns] SET [multiline] = 1 WHERE [columnID] IN (@colA_Notes, @colAR_Notes, @colAS_Notes, @colB_Notes, @colBR_Notes, @colBS_Notes);
	UPDATE [dbo].[tbsys_columns] SET [uniquechecktype] = -2 WHERE [columnID] IN (@colAS_Start_Date, @colBS_Start_Date);
	UPDATE [dbo].[tbsys_columns] SET [controltype] = 2 WHERE [columnID] IN (@colA_Main_Other_Adopter, @colB_Mother_Father_Partner);
	UPDATE [dbo].[tbsys_columns] SET [convertcase] = 1 WHERE [columnID] IN (@colA_Partner_Postcode, @colA_Partner_Employer_Postcode, @colB_Partner_Postcode, @colB_Partner_Employer_Postcode);
	UPDATE [dbo].[tbsys_columns] SET [mask] = 'AA999999S', [Trimming] = 0 WHERE [columnID] IN (@colA_Partner_NI_Number, @colB_Partner_NI_Number);

	/* -------------------- */
	/* Dropdown List Values */
	/* -------------------- */
	INSERT INTO [dbo].[ASRSysColumnControlValues]
		([columnID], [value], [sequence])
		VALUES (@colA_Main_Other_Adopter, 'Main Adopter', 1)
			, (@colA_Main_Other_Adopter, 'Other Adopter', 2)
			, (@colB_Mother_Father_Partner, 'Mother', 1)
			, (@colB_Mother_Father_Partner, 'Father/Partner', 2);

	/* ---------------- */
	/* Define Relations */
	/* ---------------- */
	INSERT INTO [dbo].[ASRSysRelations]
		([ParentID], [ChildID])
		VALUES (@tabPersonnel_Records, @tabShPL_Adoption)
			, (@tabShPL_Adoption, @tabShPLA_Leave_Requests)
			, (@tabShPL_Adoption, @tabShPLA_SPLIT_Days)
			, (@tabPersonnel_Records, @tabShPL_Birth)
			, (@tabShPL_Birth, @tabShPLB_Leave_Requests)
			, (@tabShPL_Birth, @tabShPLB_SPLIT_Days);

	/* ------------ */
	/* Calculations */
	/* ------------ */

	-- ShPL Adoption - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPL_Adoption, 'sa', @access, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, ISNULL(@colPR_Full_Name, @colPR_Surname), 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)

	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPL_Adoption;

	-- ShPL Adoption - Default Staff Number
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Staff_Number', 1, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Staff_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [dfltValueExprID] = @exprID
		WHERE [columnID] = @colA_Staff_Number;

	-- ShPL Adoption - Full Name
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Full_Name', 1, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, ISNULL(@colPR_Full_Name, @colPR_Surname), 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colA_Full_Name;

	-- ShPL Adoption - Total ShPP Weeks Available
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Total_ShPP_Weeks_Available', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);
			
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 39, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPL_Adoption, @colA_SAP_Weeks_Paid, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colA_Total_ShPP_Weeks_Available;

	-- ShPL Adoption - Total SPLIT Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Total_SPLIT_Days', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLA_SPLIT_Days, @colAS_SPLIT_Days, 1, NULL, 4, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colA_Total_SPLIT_Days;

	-- ShPL Adoption - Validate Partner NI Number
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Partner_NI_Number', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<National Insurance Number> Partner NI Number', 1, 0, 0, 3, @exprCompID, '', @tabShPL_Adoption, '', '', 0, 0, 0);
			
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 75, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPL_Adoption, @colA_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Invalid NI Number'
		WHERE [columnID] = @colA_Partner_NI_Number;

	-- ShPL Adoption - Validate Intended ShPP End Date
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Intended_ShPL_End_Date', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Condition> Intended ShPL End Date Populated', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 2, '<Field> Intended ShPL End Date', 4, 0, 0, 3, @exprCompID + 1, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 3, '<If Return> Intended ShPL End Date > Intended ShPL Start Date', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 4, '<Else Return> True', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Adoption, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 4, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 61, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPL_Adoption, @colA_Intended_ShPL_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID + 3, 1, @tabShPL_Adoption, @colA_Intended_ShPL_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 10, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID + 3, 1, @tabShPL_Adoption, @colA_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 6, @exprID + 4, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 3, NULL, NULL, 1, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be after Intended ShPL Start Date'
		WHERE [columnID] = @colA_Intended_ShPL_End_Date;

	-- ShPL Adoption - Validate ShPP Weeks Claimed
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'ShPP_Weeks_Claimed', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);
		
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_ShPP_Weeks_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPL_Adoption, @colA_ShPP_Weeks_Partner, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 7, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 1, @tabShPL_Adoption, @colA_Total_ShPP_Weeks_Available, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
			
	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Weeks to be claimed must be equal to total weeks available'
		WHERE [columnID] = @colA_ShPP_Weeks_Partner;

	-- ShPL Adoption - Validate No NI Number Declaration
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'No_NI_Number_Declaration', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Expression> No NI Number Declaration True and Partner NI Number Empty', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 2, '<Field> Partner NI Number', 1, 0, 0, 3, @exprCompID + 5, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 3, '<Expression> No NI Number Declaration False and Partner NI Number Populated', 3, 0, 0, 3, @exprCompID + 2, '', @tabShPL_Adoption, '', '', 0, 0, 0)
			, (@exprID + 4, '<Field> Partner NI Number', 1, 0, 0, 3, @exprCompID + 10, '', @tabShPL_Adoption, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 6, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 3, @exprID + 1, 1, @tabShPL_Adoption, @colA_No_NI_Number_Declaration, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID + 1, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID + 1, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 16, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 6, @exprID + 2, 1, @tabShPL_Adoption, @colA_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 7, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 13, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 8, @exprID + 3, 1, @tabShPL_Adoption, @colA_No_NI_Number_Declaration, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 9, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 10, @exprID + 3, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 61, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 11, @exprID + 4, 1, @tabShPL_Adoption, @colA_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be checked if NI Number empty and vice versa'
		WHERE [columnID] = @colA_No_NI_Number_Declaration;

	-- ShPL Adoption - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPL Adoption - Default Payroll Company Code
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Payroll_Company_Code', 1, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Company_Code, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		UPDATE [dbo].[tbsys_columns]
			SET [dfltValueExprID] = @exprID
			WHERE [columnID] = @colA_Payroll_Company_Code;
		
		-- ShPL Adoption - Trigger to Payroll Column
		IF @triggerFlag = 1
		BEGIN
			SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
			SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Is_Current_Employee_for_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

			UPDATE [dbo].[tbsys_columns]
				SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
				WHERE [columnID] = @colA_Trigger_to_Payroll;

		END;

		-- ShPL Adoption - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPL_Adoption, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, 1, @colPR_Is_Current_Employee_for_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Other_Adopter, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Adoption, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPL_Adoption = @exprID;

	END;

	-- ShPLA Leave Requests - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name_Intended_ShPL_Start_Date', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Date> Intended ShPL Start Date', 4, 0, 0, 8, @exprCompID + 4, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_Full_Name, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, ' - ShPL Start : ', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 35, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 5, @exprID + 1, 1, @tabShPL_Adoption, @colA_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPLA_Leave_Requests;

	-- ShPLA Leave Requests - Default Date of Request
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'System_Date', 4, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [dfltValueExprID] = @exprID
		WHERE [columnID] = @colAR_Date_of_Request;

	-- ShPLA Leave Requests - ShPP Weeks
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'ShPP_Weeks', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0)
				, (@exprID + 1, '<Start Date> Date Requested From', 4, 0, 0, 1, @exprCompID, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
				, (@exprID + 2, '<End Date> Date Requested To', 4, 0, 0, 1, @exprCompID, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 4, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 7, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colAR_ShPP_Weeks;

	-- ShPLA Leave Requests - Validate Date Requested From
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Date_Requested_From', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Start Date> Date Requested From', 4, 0, 0, 3, @exprCompID, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
			, (@exprID + 2, '<End Date> Date Requested To', 4, 0, 0, 3, @exprCompID, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 16, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 7, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 7, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 6, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Leave must be requested in multiples of 7 days'
		WHERE [columnID] = @colAR_Date_Requested_From;

	-- ShPLA Leave Requests - Validate Date Requested To
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Date_Requested_To', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 10, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPLA_Leave_Requests, @colAR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be after Date Requested From'
		WHERE [columnID] = @colAR_Date_Requested_To;

	-- ShPLA Leave Requests - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPLA Leave Requests - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPLA_Leave_Requests, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_Trigger_to_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Other_Adopter, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_Leave_Requests, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPLA_Leave_Requests = @exprID;

	END;

	-- ShPLA SPLIT Days - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name_Intended_ShPL_Start_Date', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPLA_SPLIT_Days, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Date> Intended ShPL Start Date', 4, 0, 0, 8, @exprCompID + 5, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_Full_Name, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, ' - ShPL Start : ', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 35, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 6, @exprID + 1, 1, @tabShPL_Adoption, @colA_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
			
	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPLA_SPLIT_Days;

	-- ShPLA SPLIT Days - SPLIT Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'SPLIT_Days', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPLA_SPLIT_Days, 'sa', @access, 0, 0, 0)
				, (@exprID + 1, '<Start Date> Start Date', 4, 0, 0, 1, @exprCompID, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0)
				, (@exprID + 2, '<End Date> End Date', 4, 0, 0, 1, @exprCompID, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLA_SPLIT_Days, @colAS_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLA_SPLIT_Days, @colAS_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [columnType] = 2 
		WHERE [columnID] = @colAS_SPLIT_Days;

	-- ShPLA SPLIT Days - Validate End Date
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'End_Date', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLA_SPLIT_Days, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLA_SPLIT_Days, @colAS_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 12, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPLA_SPLIT_Days, @colAS_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be on or after Start Date'
		WHERE [columnID] = @colAS_End_Date;

	-- ShPLA SPLIT Days - Validate Split Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'SPLIT_Days', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLA_SPLIT_Days, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_Total_SPLIT_Days, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 11, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 20, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Total number of SPLIT days must not exceed 20'
		WHERE [columnID] = @colAS_SPLIT_Days;

	-- ShPLA SPLIT Days - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPLA SPLIT Days - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPLA_SPLIT_Days, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabShPL_Adoption, @colA_Trigger_to_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Adoption, @colA_Declaration_from_Other_Adopter, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLA_SPLIT_Days, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPLA_SPLIT_Days = @exprID;

	END;

	-- ShPL Birth - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPL_Birth, 'sa', @access, 0, 0, 0)

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, ISNULL(@colPR_Full_Name, @colPR_Surname), 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)

	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPL_Birth;

	-- ShPL Birth - Default Staff Number
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Staff_Number', 1, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Staff_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [dfltValueExprID] = @exprID
		WHERE [columnID] = @colB_Staff_Number;

	-- ShPL Adoption - Full Name
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Full_Name', 1, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, ISNULL(@colPR_Full_Name, @colPR_Surname), 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colB_Full_Name;

	-- ShPL Birth - Total ShPP Weeks Available
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Total_ShPP_Weeks_Available', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Birth, 'sa', @access, 0, 0, 0);
			
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 39, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPL_Birth, @colB_SMP_Weeks_Paid, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colB_Total_ShPP_Weeks_Available;

	-- ShPL Birth - Total SPLIT Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Total_SPLIT_Days', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLB_SPLIT_Days, @colBS_SPLIT_Days, 1, NULL, 4, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colB_Total_SPLIT_Days;

	-- ShPL Birth - Validate Partner NI Number
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Partner_NI_Number', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Birth, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<National Insurance Number> Partner NI Number', 1, 0, 0, 3, @exprCompID, '', @tabShPL_Birth, '', '', 0, 0, 0);
			
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 75, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPL_Birth, @colB_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Invalid NI Number'
		WHERE [columnID] = @colB_Partner_NI_Number;

	-- ShPL Birth - Validate Intended ShPP End Date
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Intended_ShPL_End_Date', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Birth, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Condition> Intended ShPL End Date Populated', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 2, '<Field> Intended ShPL End Date', 4, 0, 0, 3, @exprCompID + 1, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 3, '<If Return> Intended ShPL End Date > Intended ShPL Start Date', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 4, '<Else Return> True', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Birth, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 4, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 61, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPL_Birth, @colB_Intended_ShPL_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID + 3, 1, @tabShPL_Birth, @colB_Intended_ShPL_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 10, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID + 3, 1, @tabShPL_Birth, @colB_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 6, @exprID + 4, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 3, NULL, NULL, 1, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be after Intended ShPL Start Date'
		WHERE [columnID] = @colB_Intended_ShPL_End_Date;

	-- ShPL Birth - Validate ShPP Weeks Claimed
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'ShPP_Weeks_Claimed', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Birth, 'sa', @access, 0, 0, 0);
		
	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_ShPP_Weeks_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPL_Birth, @colB_ShPP_Weeks_Partner, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 7, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 1, @tabShPL_Birth, @colB_Total_ShPP_Weeks_Available, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
			
	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Weeks to be claimed must be equal to total weeks available'
		WHERE [columnID] = @colB_ShPP_Weeks_Partner;

	-- ShPL Birth - Validate No NI Number Declaration
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'No_NI_Number_Declaration', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPL_Birth, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Expression> No NI Number Declaration True and Partner NI Number Empty', 3, 0, 0, 3, @exprCompID, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 2, '<Field> Partner NI Number', 1, 0, 0, 3, @exprCompID + 5, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 3, '<Expression> No NI Number Declaration False and Partner NI Number Populated', 3, 0, 0, 3, @exprCompID + 2, '', @tabShPL_Birth, '', '', 0, 0, 0)
			, (@exprID + 4, '<Field> Partner NI Number', 1, 0, 0, 3, @exprCompID + 10, '', @tabShPL_Birth, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 6, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 3, @exprID + 1, 1, @tabShPL_Birth, @colB_No_NI_Number_Declaration, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID + 1, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID + 1, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 16, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 6, @exprID + 2, 1, @tabShPL_Birth, @colB_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 7, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 13, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 8, @exprID + 3, 1, @tabShPL_Birth, @colB_No_NI_Number_Declaration, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 9, @exprID + 3, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 10, @exprID + 3, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 61, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 11, @exprID + 4, 1, @tabShPL_Birth, @colB_Partner_NI_Number, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be checked if NI Number empty and vice versa'
		WHERE [columnID] = @colB_No_NI_Number_Declaration;

	-- ShPL Birth - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPL Birth - Default Payroll Company Code
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Payroll_Company_Code', 1, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Company_Code, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		UPDATE [dbo].[tbsys_columns]
			SET [dfltValueExprID] = @exprID
			WHERE [columnID] = @colB_Payroll_Company_Code;
		
		-- ShPL Birth - Trigger to Payroll Column
		IF @triggerFlag = 1
		BEGIN
			SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
			SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabPersonnel_Records, @colPR_Is_Current_Employee_for_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

			UPDATE [dbo].[tbsys_columns]
				SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
				WHERE [columnID] = @colB_Trigger_to_Payroll;

		END;

		-- ShPL Birth - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPL_Birth, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, 1, @colPR_Is_Current_Employee_for_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Other_Parent, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPL_Birth, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPL_Birth = @exprID;

	END;

	-- ShPLB Leave Requests - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name_Intended_ShPL_Start_Date', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Date> Intended ShPL Start Date', 4, 0, 0, 8, @exprCompID + 4, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_Full_Name, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, ' - ShPL Start : ', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 35, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 5, @exprID + 1, 1, @tabShPL_Birth, @colB_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPLB_Leave_Requests;

	-- ShPLB Leave Requests - Default Date of Request
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'System_Date', 4, 0, 0, 4, 0, 'Advanced Business Solutions Standard Default Value', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [dfltValueExprID] = @exprID
		WHERE [columnID] = @colBR_Date_of_Request;

	-- ShPLB Leave Requests - ShPP Weeks
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'ShPP_Weeks', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0)
				, (@exprID + 1, '<Start Date> Date Requested From', 4, 0, 0, 1, @exprCompID, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
				, (@exprID + 2, '<End Date> Date Requested To', 4, 0, 0, 1, @exprCompID, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 4, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 7, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [readOnly] = 1, [columnType] = 2 
		WHERE [columnID] = @colBR_ShPP_Weeks;

	-- ShPLB Leave Requests - Validate Date Requested From
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Date_Requested_From', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Start Date> Date Requested From', 4, 0, 0, 3, @exprCompID, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
			, (@exprID + 2, '<End Date> Date Requested To', 4, 0, 0, 3, @exprCompID, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 16, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 7, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 7, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 6, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Leave must be requested in multiples of 7 days'
		WHERE [columnID] = @colBR_Date_Requested_From;

	-- ShPLB Leave Requests - Validate Date Requested To
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Date_Requested_To', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_To, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 10, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPLB_Leave_Requests, @colBR_Date_Requested_From, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be after Date Requested From'
		WHERE [columnID] = @colBR_Date_Requested_To;

	-- ShPLB Leave Requests - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPLB Leave Requests - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPLB_Leave_Requests, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_Trigger_to_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Other_Parent, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_Leave_Requests, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPLB_Leave_Requests = @exprID;

	END;

	-- ShPLB SPLIT Days - Record Description
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'Name_Intended_ShPL_Start_Date', 1, 0, 0, 8, 0, 'Advanced Business Solutions Standard Record Description', @tabShPLB_SPLIT_Days, 'sa', @access, 0, 0, 0)
			, (@exprID + 1, '<Date> Intended ShPL Start Date', 4, 0, 0, 8, @exprCompID + 5, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_Full_Name, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 3, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, ' - ShPL Start : ', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 4, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 17, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 5, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 35, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 6, @exprID + 1, 1, @tabShPL_Birth, @colB_Intended_ShPL_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
			
	UPDATE [dbo].[tbsys_tables]
		SET [RecordDescExprID] = @exprID
		WHERE [TableID] = @tabShPLB_SPLIT_Days;

	-- ShPLB SPLIT Days - SPLIT Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'SPLIT_Days', 2, 0, 0, 1, 0, 'Advanced Business Solutions Standard Calculation', @tabShPLB_SPLIT_Days, 'sa', @access, 0, 0, 0)
				, (@exprID + 1, '<Start Date> Start Date', 4, 0, 0, 1, @exprCompID, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0)
				, (@exprID + 2, '<End Date> End Date', 4, 0, 0, 1, @exprCompID, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 45, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
			, (@exprCompID + 1, @exprID + 1, 1, @tabShPLB_SPLIT_Days, @colBS_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID + 2, 1, @tabShPLB_SPLIT_Days, @colBS_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [calcExprID] = @exprID, [columnType] = 2 
		WHERE [columnID] = @colBS_SPLIT_Days;

	-- ShPLB SPLIT Days - Validate End Date
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'End_Date', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLB_SPLIT_Days, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPLB_SPLIT_Days, @colBS_End_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 12, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 1, @tabShPLB_SPLIT_Days, @colBS_Start_Date, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Must be on or after Start Date'
		WHERE [columnID] = @colBS_End_Date;

	-- ShPLB SPLIT Days - Validate Split Days
	SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

	INSERT INTO [dbo].[ASRSysExpressions]
		([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
		VALUES (@exprID, 'SPLIT_Days', 3, 0, 0, 3, 0, 'Advanced Business Solutions Standard Validation', @tabShPLB_SPLIT_Days, 'sa', @access, 0, 0, 0);

	INSERT INTO [dbo].[ASRSysExprComponents]
		([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
			, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
			, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
		VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_Total_SPLIT_Days, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 11, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
			, (@exprCompID + 2, @exprID, 4, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 2, NULL, 20, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

	UPDATE [dbo].[tbsys_columns]
		SET [lostFocusExprID] = @exprID, [errorMessage] = 'Total number of SPLIT days must not exceed 20'
		WHERE [columnID] = @colBS_SPLIT_Days;

	-- ShPLB SPLIT Days - Payroll Module Only
	IF @payrollModule = 1
	BEGIN
		-- ShPLB SPLIT Days - Trigger to Payroll Filter
		SELECT @exprID = MAX([ExprID]) + 1 FROM [dbo].[ASRSysExpressions];
		SELECT @exprCompID = MAX([ComponentID]) + 1 FROM [dbo].[ASRSysExprComponents];

		INSERT INTO [dbo].[ASRSysExpressions]
			([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
			VALUES (@exprID, 'Trigger_to_Payroll', 3, 0, 0, 5, 0, 'Advanced Business Solutions Standard Filter', @tabShPLB_SPLIT_Days, 'sa', @access, 0, 0, 0);

		IF @triggerFlag = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID, @exprID, 1, @tabShPL_Birth, @colB_Trigger_to_Payroll, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 1, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		END;

		INSERT INTO [dbo].[ASRSysExprComponents]
			([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
				, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
				, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
			VALUES (@exprCompID + 2, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Employee, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 3, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
				, (@exprCompID + 4, @exprID, 1, @tabShPL_Birth, @colB_Declaration_from_Other_Parent, 1, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);
		
		IF @enablePayType = 'Record'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "ENABLEPAY"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Logic Value', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'ENABLEPAY', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Logic_Value, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;

		IF @enablePayType = 'Column'
		BEGIN
			INSERT INTO [dbo].[ASRSysExpressions]
				([ExprID], [Name], [ReturnType], [ReturnSize], [ReturnDecimals], [Type], [ParentComponentID], [Description], [TableID], [Username], [Access], [ExpandedNode], [ViewInColour], [UtilityID])
				VALUES (@exprID + 1, '<Search Field> Global Variables : Global Key', 101, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 2, '<Search Expression> "Global"', 1, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0)
					, (@exprID + 3, '<Return Field> Global Variables : Enable Payroll Transfer', 103, 0, 0, 5, @exprCompID + 6, '', @tabShPLB_SPLIT_Days, '', '', 0, 0, 0);

			INSERT INTO [dbo].[ASRSysExprComponents]
				([ComponentID], [ExprID], [Type], [FieldTableID], [FieldColumnID], [FieldPassBy], [FieldSelectionTableID], [FieldSelectionRecord], [FieldSelectionLine], [FieldSelectionOrderID], [FieldSelectionFilter]
					, [FunctionID], [CalculationID], [OperatorID], [ValueType], [ValueCharacter], [ValueNumeric], [ValueLogic], [ValueDate], [PromptDescription], [PromptMask], [PromptSize], [PromptDecimals]
					, [FunctionReturnType], [LookupTableID], [LookupColumnID], [FilterID], [ExpandedNode], [PromptDateType], [WorkflowElement], [WorkflowItem], [WorkflowRecord], [WorkflowRecordTableID], [WorkflowElementProperty])
				VALUES (@exprCompID + 5, @exprID, 5, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 5, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 6, @exprID, 2, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 42, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, '', '', -1, 0, NULL)
					, (@exprCompID + 7, @exprID + 1, 1, @tabGlobal_Variables, @colGV_Global_Key, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 8, @exprID + 2, 6, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, 'Global', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, @tabGlobal_Variables, @colGV_Global_Key, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL)
					, (@exprCompID + 9, @exprID + 3, 1, @tabGlobal_Variables, @colGV_Enable_Payroll_Transfer, 2, NULL, 1, 1, 0, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);

		END;
		SET @fltrShPLB_SPLIT_Days = @exprID;

	END;

	-- Reset System Settings
	SELECT @exprID = MAX([ExprID]) FROM [dbo].[ASRSysExpressions];
	SELECT @exprCompID = MAX([ComponentID]) FROM [dbo].[ASRSysExprComponents];
	EXEC dbo.spsys_setsystemsetting 'autoid', 'expressions', @exprID;
	EXEC dbo.spsys_setsystemsetting 'autoid', 'exprcomponents', @exprCompID;

	/* ------------------ */
	/* Shared Table Setup */
	/* ------------------ */
	IF @payrollModule = 1
	BEGIN
		INSERT INTO [dbo].[ASRSysAccordTransferTypes]
			([TransferTypeID], [TransferType], [FilterID], [ASRBaseTableID], [IsVisible], [ForceAsUpdate])
			VALUES (95, 'ShPL (Birth)', @fltrShPL_Birth, @tabShPL_Birth, 1, 0)
				, (96, 'ShPL (Birth) Leave Requests', @fltrShPLB_Leave_Requests, @tabShPLB_Leave_Requests, 1, 0)
				, (97, 'ShPL (Birth) SPLIT Days', @fltrShPLB_SPLIT_Days, @tabShPLB_SPLIT_Days, 1, 0)
				, (98, 'ShPL (Adoption)', @fltrShPL_Adoption, @tabShPL_Adoption, 1, 0)
				, (99, 'ShPL (Adoption) Leave Requests', @fltrShPLA_Leave_Requests, @tabShPLA_Leave_Requests, 1, 0)
				, (100, 'ShPL (Adoption) SPLIT Days', @fltrShPLA_SPLIT_Days, @tabShPLA_SPLIT_Days, 1, 0);

		-- ShPL (Birth)
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 95, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabPersonnel_Records, @colPR_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 95, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabPersonnel_Records, @colPR_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 95, 1, 'Employee is Mother', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Mother_Father_Partner, 0, 'null', 1, 0, 0, 0, 0, 0, 0)
				, (3, 95, 1, 'Employee is Father/Partner', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Mother_Father_Partner, 0, 'null', 1, 0, 0, 0, 0, 0, 0)
				, (4, 95, 1, 'Expected Birth Date', 1, 1, 0, 0, 2, 0, @tabShPL_Birth, @colB_Expected_Birth_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (5, 95, 0, 'Actual Birth Date', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Actual_Birth_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (6, 95, 1, 'SMP Curtailment Date', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_SMP_Curtailment_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (7, 95, 1, 'Partner Forename', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Forename, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (8, 95, 1, 'Partner Surname', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Surname, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (9, 95, 0, 'Partner Address 1', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Address_1, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (10, 95, 0, 'Partner Address 2', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Address_2, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (11, 95, 0, 'Partner Address 3', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Address_3, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (12, 95, 0, 'Partner Address 4', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Address_4, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (13, 95, 0, 'Partner Postcode', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Postcode, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (14, 95, 0, 'Partner NI Number', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_NI_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (15, 95, 0, 'No NI Number Declaration', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_No_NI_Number_Declaration, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (16, 95, 0, 'Partner Employer Name', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Name, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (17, 95, 0, 'Partner Employer Address 1', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Address_1, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (18, 95, 0, 'Partner Employer Address 2', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Address_2, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (19, 95, 0, 'Partner Employer Address 3', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Address_3, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (20, 95, 0, 'Partner Employer Address 4', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Address_4, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (21, 95, 0, 'Partner Employer Postcode', 0, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Partner_Employer_Postcode, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (22, 95, 1, 'Intended ShPL Start Date', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Intended_ShPL_Start_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (23, 95, 0, 'Intended ShPL End Date', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Intended_ShPL_End_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (24, 95, 1, 'ShPP Weeks Claim Employee', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_ShPP_Weeks_Employee, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (25, 95, 1, 'ShPP Weeks Claim Partner', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_ShPP_Weeks_Partner, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (26, 95, 1, 'Declaration from Employee', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Declaration_from_Employee, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (27, 95, 1, 'Declaration from Other Parent', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Declaration_from_Other_Parent, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (28, 95, 1, 'Date Notification Received', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Date_Notification_Received, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (29, 95, 0, 'Location of Birth', 1, 0, 0, 0, 2, 0, @tabShPL_Birth, @colB_Location_of_Birth, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- ShPL (Birth) Leave Requests
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 96, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabShPL_Birth, @colB_Payroll_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 96, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabShPL_Birth, @colB_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 96, 1, 'Date Request Received', 1, 1, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Date_of_Request, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (3, 96, 1, 'Request From Date', 1, 0, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Date_Requested_From, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (4, 96, 1, 'Request To Date', 1, 0, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Date_Requested_To, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (5, 96, 0, 'Request Binding', 1, 0, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Binding_Request, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (6, 96, 0, 'Consent from Other Parent', 1, 0, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Consent_from_Other_Parent, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (7, 96, 0, 'Request Cancelled', 1, 0, 0, 0, 2, 0, @tabShPLB_Leave_Requests, @colBR_Request_Cancelled, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- ShPL (Birth) SPLIT Days
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 97, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabShPL_Birth, @colB_Payroll_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 97, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabShPL_Birth, @colB_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 97, 1, 'Start Date', 1, 1, 0, 0, 2, 0, @tabShPLB_SPLIT_Days, @colBS_Start_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (3, 97, 1, 'End Date', 1, 0, 0, 0, 2, 0, @tabShPLB_SPLIT_Days, @colBS_End_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (4, 97, 1, 'Reason', 1, 0, 0, 0, 2, 0, @tabShPLB_SPLIT_Days, @colBS_Reason, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- ShPL (Adoption)
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 98, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabPersonnel_Records, @colPR_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 98, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabPersonnel_Records, @colPR_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 98, 1, 'Employee is Main Adopter', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Main_Other_Adopter, 0, 'null', 1, 0, 0, 0, 0, 0, 0)
				, (3, 98, 1, 'Employee is Other Adopter', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Main_Other_Adopter, 0, 'null', 1, 0, 0, 0, 0, 0, 0)
				, (4, 98, 1, 'Placement Date', 1, 1, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Placement_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (5, 98, 1, 'SAP Curtailment Date', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_SAP_Curtailment_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (6, 98, 1, 'Partner Forename', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Forename, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (7, 98, 1, 'Partner Surname', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Surname, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (8, 98, 0, 'Partner Address 1', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Address_1, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (9, 98, 0, 'Partner Address 2', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Address_2, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (10, 98, 0, 'Partner Address 3', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Address_3, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (11, 98, 0, 'Partner Address 4', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Address_4, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (12, 98, 0, 'Partner Postcode', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Postcode, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (13, 98, 0, 'Partner NI Number', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_NI_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (14, 98, 0, 'No NI Number Declaration', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_No_NI_Number_Declaration, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (15, 98, 0, 'Partner Employer Name', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Name, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (16, 98, 0, 'Partner Employer Address 1', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Address_1, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (17, 98, 0, 'Partner Employer Address 2', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Address_2, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (18, 98, 0, 'Partner Employer Address 3', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Address_3, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (19, 98, 0, 'Partner Employer Address 4', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Address_4, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (20, 98, 0, 'Partner Employer Postcode', 0, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Partner_Employer_Postcode, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (21, 98, 1, 'Intended ShPL Start Date', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Intended_ShPL_Start_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (22, 98, 0, 'Intended ShPL End Date', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Intended_ShPL_End_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (23, 98, 1, 'ShPP Weeks Claim Employee', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_ShPP_Weeks_Employee, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (24, 98, 1, 'ShPP Weeks Claim Partner', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_ShPP_Weeks_Partner, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (25, 98, 1, 'Declaration from Employee', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Declaration_from_Employee, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (26, 98, 1, 'Declaration from Other Adopter', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Declaration_from_Other_Adopter, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (27, 98, 1, 'Date Notification Received', 1, 0, 0, 0, 2, 0, @tabShPL_Adoption, @colA_Date_Notification_Received, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- ShPL (Adoption) Leave Requests
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 99, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabShPL_Adoption, @colA_Payroll_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 99, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabShPL_Adoption, @colA_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 99, 1, 'Date Request Received', 1, 1, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Date_of_Request, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (3, 99, 1, 'Request From Date', 1, 0, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Date_Requested_From, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (4, 99, 1, 'Request To Date', 1, 0, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Date_Requested_To, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (5, 99, 0, 'Request Binding', 1, 0, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Binding_Request, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (6, 99, 0, 'Consent from Other Adopter', 1, 0, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Consent_from_Other_Adopter, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (7, 99, 0, 'Request Cancelled', 1, 0, 0, 0, 2, 0, @tabShPLA_Leave_Requests, @colAR_Request_Cancelled, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- ShPL (Adoption) SPLIT Days
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [AlwaysTransfer], [IsKeyField], [IsCompanyCode], [IsEmployeeCode], [Direction], [ASRMapType]
				, [ASRTableID], [ASRColumnID], [ASRExprID], [ASRValue], [ConvertData], [IsEmployeeName], [IsDepartmentCode], [IsDepartmentName], [IsPayrollCode], [GroupBy], [PreventModify])
			VALUES (0, 100, 1, 'Company Code', 1, 1, 1, 0, 2, 0, @tabShPL_Adoption, @colA_Payroll_Company_Code, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (1, 100, 1, 'Employee Code', 1, 1, 0, 1, 2, 0, @tabShPL_Adoption, @colA_Staff_Number, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (2, 100, 1, 'Start Date', 1, 1, 0, 0, 2, 0, @tabShPLA_SPLIT_Days, @colAS_Start_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (3, 100, 1, 'End Date', 1, 0, 0, 0, 2, 0, @tabShPLA_SPLIT_Days, @colAS_End_Date, 0, 'null', 0, 0, 0, 0, 0, 0, 0)
				, (4, 100, 1, 'Reason', 1, 0, 0, 0, 2, 0, @tabShPLA_SPLIT_Days, @colAS_Reason, 0, 'null', 0, 0, 0, 0, 0, 0, 0);

		-- Translations
		INSERT INTO [dbo].[ASRSysAccordTransferFieldMappings]
			([TransferID], [FieldID], [HRProValue], [AccordValue])
			VALUES (95, 2, 'Mother', 1)
				, (95, 2, 'Father/Partner', 0)
				, (95, 3, 'Mother', 0)
				, (95, 3, 'Father/Partner', 1)
				, (98, 2, 'Main Adopter', 1)
				, (98, 2, 'Other Adopter', 0)
				, (98, 3, 'Main Adopter', 0)
				, (98, 3, 'Other Adopter', 1);

	END
	ELSE
	BEGIN
		INSERT INTO [dbo].[ASRSysAccordTransferTypes]
			([TransferTypeID], [TransferType], [FilterID], [ASRBaseTableID], [IsVisible], [ForceAsUpdate])
			VALUES (95, 'ShPL (Birth)', 0, 0, 1, 0)
				, (96, 'ShPL (Birth) Leave Requests', 0, 0, 1, 0)
				, (97, 'ShPL (Birth) SPLIT Days', 0, 0, 1, 0)
				, (98, 'ShPL (Adoption)', 0, 0, 1, 0)
				, (99, 'ShPL (Adoption) Leave Requests', 0, 0, 1, 0)
				, (100, 'ShPL (Adoption) SPLIT Days', 0, 0, 1, 0);

		-- ShPL (Birth)
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 95, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 95, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 95, 1, 'Employee is Mother', 0, 0, 2, 0, 1)
				, (3, 95, 1, 'Employee is Father/Partner', 0, 0, 2, 0, 1)
				, (4, 95, 1, 'Expected Birth Date', 0, 0, 2, 1, 1)
				, (5, 95, 0, 'Actual Birth Date', 0, 0, 2, 0, 1)
				, (6, 95, 1, 'SMP Curtailment Date', 0, 0, 2, 0, 1)
				, (7, 95, 1, 'Partner Forename', 0, 0, 2, 0, 1)
				, (8, 95, 1, 'Partner Surname', 0, 0, 2, 0, 1)
				, (9, 95, 0, 'Partner Address 1', 0, 0, 2, 0, 0)
				, (10, 95, 0, 'Partner Address 2', 0, 0, 2, 0, 0)
				, (11, 95, 0, 'Partner Address 3', 0, 0, 2, 0, 0)
				, (12, 95, 0, 'Partner Address 4', 0, 0, 2, 0, 0)
				, (13, 95, 0, 'Partner Postcode', 0, 0, 2, 0, 0)
				, (14, 95, 0, 'Partner NI Number', 0, 0, 2, 0, 0)
				, (15, 95, 0, 'No NI Number Declaration', 0, 0, 2, 0, 0)
				, (16, 95, 0, 'Partner Employer Name', 0, 0, 2, 0, 0)
				, (17, 95, 0, 'Partner Employer Address 1', 0, 0, 2, 0, 0)
				, (18, 95, 0, 'Partner Employer Address 2', 0, 0, 2, 0, 0)
				, (19, 95, 0, 'Partner Employer Address 3', 0, 0, 2, 0, 0)
				, (20, 95, 0, 'Partner Employer Address 4', 0, 0, 2, 0, 0)
				, (21, 95, 0, 'Partner Employer Postcode', 0, 0, 2, 0, 0)
				, (22, 95, 1, 'Intended ShPL Start Date', 0, 0, 2, 0, 1)
				, (23, 95, 0, 'Intended ShPL End Date', 0, 0, 2, 0, 1)
				, (24, 95, 1, 'ShPP Weeks Claim Employee', 0, 0, 2, 0, 1)
				, (25, 95, 1, 'ShPP Weeks Claim Partner', 0, 0, 2, 0, 1)
				, (26, 95, 1, 'Declaration from Employee', 0, 0, 2, 0, 1)
				, (27, 95, 1, 'Declaration from Other Parent', 0, 0, 2, 0, 1)
				, (28, 95, 1, 'Date Notification Received', 0, 0, 2, 0, 1)
				, (29, 95, 0, 'Location of Birth', 0, 0, 2, 0, 1);

		-- ShPL (Birth) Leave Requests
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 96, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 96, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 96, 1, 'Date Request Received', 0, 0, 2, 1, 1)
				, (3, 96, 1, 'Request From Date', 0, 0, 2, 0, 1)
				, (4, 96, 1, 'Request To Date', 0, 0, 2, 0, 1)
				, (5, 96, 0, 'Request Binding', 0, 0, 2, 0, 1)
				, (6, 96, 0, 'Consent from Other Parent', 0, 0, 2, 0, 1)
				, (7, 96, 0, 'Request Cancelled', 0, 0, 2, 0, 1);

		-- ShPL (Birth) SPLIT Days
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 97, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 97, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 97, 1, 'Start Date', 0, 0, 2, 1, 1)
				, (3, 97, 1, 'End Date', 0, 0, 2, 0, 1)
				, (4, 97, 1, 'Reason', 0, 0, 2, 0, 1);

		-- ShPL (Adoption)
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 98, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 98, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 98, 1, 'Employee is Main Adopter', 0, 0, 2, 0, 1)
				, (3, 98, 1, 'Employee is Other Adopter', 0, 0, 2, 0, 1)
				, (4, 98, 1, 'Placement Date', 0, 0, 2, 1, 1)
				, (5, 98, 1, 'SAP Curtailment Date', 0, 0, 2, 0, 1)
				, (6, 98, 1, 'Partner Forename', 0, 0, 2, 0, 1)
				, (7, 98, 1, 'Partner Surname', 0, 0, 2, 0, 1)
				, (8, 98, 0, 'Partner Address 1', 0, 0, 2, 0, 0)
				, (9, 98, 0, 'Partner Address 2', 0, 0, 2, 0, 0)
				, (10, 98, 0, 'Partner Address 3', 0, 0, 2, 0, 0)
				, (11, 98, 0, 'Partner Address 4', 0, 0, 2, 0, 0)
				, (12, 98, 0, 'Partner Postcode', 0, 0, 2, 0, 0)
				, (13, 98, 0, 'Partner NI Number', 0, 0, 2, 0, 0)
				, (14, 98, 0, 'No NI Number Declaration', 0, 0, 2, 0, 0)
				, (15, 98, 0, 'Partner Employer Name', 0, 0, 2, 0, 0)
				, (16, 98, 0, 'Partner Employer Address 1', 0, 0, 2, 0, 0)
				, (17, 98, 0, 'Partner Employer Address 2', 0, 0, 2, 0, 0)
				, (18, 98, 0, 'Partner Employer Address 3', 0, 0, 2, 0, 0)
				, (19, 98, 0, 'Partner Employer Address 4', 0, 0, 2, 0, 0)
				, (20, 98, 0, 'Partner Employer Postcode', 0, 0, 2, 0, 0)
				, (21, 98, 1, 'Intended ShPL Start Date', 0, 0, 2, 0, 1)
				, (22, 98, 0, 'Intended ShPL End Date', 0, 0, 2, 0, 1)
				, (23, 98, 1, 'ShPP Weeks Claim Employee', 0, 0, 2, 0, 1)
				, (24, 98, 1, 'ShPP Weeks Claim Partner', 0, 0, 2, 0, 1)
				, (25, 98, 1, 'Declaration from Employee', 0, 0, 2, 0, 1)
				, (26, 98, 1, 'Declaration from Other Adopter', 0, 0, 2, 0, 1)
				, (27, 98, 1, 'Date Notification Received', 0, 0, 2, 0, 1);

		-- ShPL (Adoption) Leave Requests
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 99, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 99, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 99, 1, 'Date Request Received', 0, 0, 2, 1, 1)
				, (3, 99, 1, 'Request From Date', 0, 0, 2, 0, 1)
				, (4, 99, 1, 'Request To Date', 0, 0, 2, 0, 1)
				, (5, 99, 0, 'Request Binding', 0, 0, 2, 0, 1)
				, (6, 99, 0, 'Consent from Other Adopter', 0, 0, 2, 0, 1)
				, (7, 99, 0, 'Request Cancelled', 0, 0, 2, 0, 1);

		-- ShPL (Adoption) SPLIT Days
		INSERT INTO [dbo].[ASRSysAccordTransferFieldDefinitions]
			([TransferFieldID], [TransferTypeID], [Mandatory], [Description], [IsCompanyCode], [IsEmployeeCode], [Direction], [IsKeyField], [AlwaysTransfer])
			VALUES (0, 100, 1, 'Company Code', 1, 0, 2, 1, 1)
				, (1, 100, 1, 'Employee Code', 0, 1, 2, 1, 1)
				, (2, 100, 1, 'Start Date', 0, 0, 2, 1, 1)
				, (3, 100, 1, 'End Date', 0, 0, 2, 0, 1)
				, (4, 100, 1, 'Reason', 0, 0, 2, 0, 1);

	END;

	/* ------------------ */
	/* Indicate Installed */
	/* ------------------ */
	EXEC dbo.spsys_setsystemsetting 'statutory', 'sharedparentalleave', 1;

END;

/* --------------------------- */
/* Delete dedicated procedures */
/* --------------------------- */
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spshpl_scriptnewtable]') AND xtype = 'P')
	DROP PROCEDURE [dbo].[spshpl_scriptnewtable];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spshpl_scriptnewcolumn]') AND xtype = 'P')
	DROP PROCEDURE [dbo].[spshpl_scriptnewcolumn];

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
PRINT 'Update Script has modified your OpenHR Database to contain Shared Parental Leave'
