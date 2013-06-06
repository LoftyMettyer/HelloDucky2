
/* ----------------------------------------------------- */
/* Update the database from version 2.17 to version 2.18 */
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
	@sSQLVersion nvarchar(20)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 2.17 or 2.18. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '2.17') and (@sDBVersion <> '2.18')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 24 - Adding Accord Transfer Types Table'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[asrSysAccordTransferTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
	
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransferTypes] (
						[TransferTypeID] [int] NULL ,
						[TransferType] [nchar] (20) ,
						[FilterID] [int] NULL,
						[ASRBaseTableID] [int] NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID) VALUES (0, ''Employee'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID) VALUES (1, ''Salary'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID) VALUES (2, ''Allowances'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID) VALUES (3, ''Loans'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID) VALUES (4, ''Deductions'' ,0,0)'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 2 of 24 - Adding Accord Transfer Field Definitions Table'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransferFieldDefinitions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
	
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransferFieldDefinitions] (
			[TransferFieldID] [int] NOT NULL ,
			[TransferTypeID] [int] NOT NULL ,
			[Mandatory] [bit] NOT NULL ,
			[Description] [char] (40) NOT NULL ,
			[AlwaysTransfer] [bit] NOT NULL,
			[IsKeyField] [bit] NOT NULL,
			[IsCompanyCode] [bit] NOT NULL,
			[IsEmployeeCode] [bit] NOT NULL,
			[Direction] [int] NOT NULL,
			[ASRMapType] [int] NULL,
			[ASRTableID] [int] NULL,
			[ASRColumnID] [int] NULL,
			[ASRExprID] [int] NULL,
			[ASRValue] [char] (40) NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,0,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,0,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,0,1,''Employee Surname'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,0,1,''Employee Forenames'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,0,0,''Employee Title'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,0,1,''Employee NI Number'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,0,1,''Employee Date of Birth'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,0,1,''Employee Gender'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,0,0,''Employee Marital Status'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,0,1,''Employee Address Line 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,0,1,''Employee Address Line 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,0,1,''Employee Address Line 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,0,1,''Employee Address Line 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,0,1,''Employee Address Line 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,0,1,''Employee Address Post Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,0,0,''Telephone'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,0,0,''Mobile'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,0,0,''Text Payment Advice'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (18,0,0,''Email'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (19,0,0,''Email Payslip'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (20,0,1,''Employment Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (21,0,1,''Leaving Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (22,0,0,''Payment Frequency'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (23,0,0,''Payment Method'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (24,0,0,''Bank Name'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (25,0,0,''Branch Address Line 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (26,0,0,''Branch Address Line 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (27,0,0,''Branch Address Line 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (28,0,0,''Branch Address Line 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (29,0,0,''Branch Address Post Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (30,0,0,''Branch Sort Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (31,0,0,''Bank Account Name'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (32,0,0,''Bank Account Number'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (33,0,0,''BACS Reference Number'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (34,0,0,''Autopay Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (35,0,0,''Account Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (36,0,0,''Department Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (37,0,0,''Employee Category Code'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (38,0,0,''Nominal Costs Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (39,0,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (40,0,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (41,0,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (42,0,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (43,0,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (44,0,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (45,0,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (46,0,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (47,0,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (48,0,0,''Tax Code + W1/M1 Basis'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (49,0,0,''P45 Previous Employment Taxable Pay'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (50,0,0,''P45 Previous Employment Tax Paid'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (51,0,0,''NI Letter'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (52,0,0,''Full Time Equivalent Hours'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (53,0,0,''Contracted Hours'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (54,0,0,''Part Timer Flag'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (55,0,0,''Director Flag'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (56,0,0,''Director Start Week'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (57,0,0,''Pension Scheme Number 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (58,0,0,''Pension Employee 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (59,0,0,''Pension Employer 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (60,0,0,''Pension AVC 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (61,0,0,''Pension Joining Date 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (62,0,0,''Pension Leaving Date 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (63,0,0,''Pension Policy Number 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (64,0,0,''Pension Scheme Number 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (65,0,0,''Pension Employee 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (66,0,0,''Pension Employer 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (67,0,0,''Pension AVC 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (68,0,0,''Pension Joining Date 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (69,0,0,''Pension Leaving Date 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (70,0,0,''Pension Policy Number 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (71,0,0,''Pension Scheme Number 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (72,0,0,''Pension Employee 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (73,0,0,''Pension Employer 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (74,0,0,''Pension AVC 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (75,0,0,''Pension Joining Date 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (76,0,0,''Pension Leaving Date 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (77,0,0,''Pension Policy Number 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (101,0,0,''User Definable Amount 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (102,0,0,''User Definable Amount 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (103,0,0,''User Definable Amount 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (104,0,0,''User Definable Amount 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (105,0,0,''User Definable Amount 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (106,0,0,''User Definable Amount 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (107,0,0,''User Definable Amount 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (108,0,0,''User Definable Amount 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (109,0,0,''User Definable Amount 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (110,0,0,''User Definable Amount 10'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (111,0,0,''User Definable Amount 11'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (112,0,0,''User Definable Amount 12'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (113,0,0,''User Definable Amount 13'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (114,0,0,''User Definable Amount 14'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (115,0,0,''User Definable Amount 15'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (116,0,0,''User Definable Amount 16'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (117,0,0,''User Definable Amount 17'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (118,0,0,''User Definable Amount 18'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (119,0,0,''User Definable Amount 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (121,0,0,''User Definable Flag 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (122,0,0,''User Definable Flag 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (123,0,0,''User Definable Flag 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (124,0,0,''User Definable Flag 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (125,0,0,''User Definable Flag 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (126,0,0,''User Definable Flag 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (127,0,0,''User Definable Flag 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (128,0,0,''User Definable Flag 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (129,0,0,''User Definable Flag 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (130,0,0,''User Definable Flag 10'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (131,0,0,''User Definable Flag 11'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (132,0,0,''User Definable Flag 12'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (133,0,0,''User Definable Flag 13'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (134,0,0,''User Definable Flag 14'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (135,0,0,''User Definable Flag 15'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (136,0,0,''User Definable Flag 16'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (137,0,0,''User Definable Flag 17'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (138,0,0,''User Definable Flag 18'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (139,0,0,''User Definable Flag 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (141,0,0,''User Definable Date 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (142,0,0,''User Definable Date 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (143,0,0,''User Definable Date 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (144,0,0,''User Definable Date 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (145,0,0,''User Definable Date 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (146,0,0,''User Definable Date 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (147,0,0,''User Definable Date 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (148,0,0,''User Definable Date 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (149,0,0,''User Definable Date 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (150,0,0,''User Definable Date 10'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (151,0,0,''User Definable Date 11'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (152,0,0,''User Definable Date 12'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (153,0,0,''User Definable Date 13'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (154,0,0,''User Definable Date 14'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (155,0,0,''User Definable Date 15'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (156,0,0,''User Definable Date 16'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (157,0,0,''User Definable Date 17'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (158,0,0,''User Definable Date 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (159,0,0,''User Definable Date 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (161,0,0,''User Definable Text 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (162,0,0,''User Definable Text 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (163,0,0,''User Definable Text 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (164,0,0,''User Definable Text 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (165,0,0,''User Definable Text 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (166,0,0,''User Definable Text 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (167,0,0,''User Definable Text 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (168,0,0,''User Definable Text 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (169,0,0,''User Definable Text 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (170,0,0,''User Definable Text 10'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (171,0,0,''User Definable Text 11'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (172,0,0,''User Definable Text 12'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (173,0,0,''User Definable Text 13'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (174,0,0,''User Definable Text 14'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (175,0,0,''User Definable Text 15'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (176,0,0,''User Definable Text 16'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (177,0,0,''User Definable Text 17'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (178,0,0,''User Definable Text 18'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (179,0,0,''User Definable Text 19'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	
		-- Salary History Transfer Types
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,1,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,1,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,1,1,''Contract No'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,1,1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,1,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,1,1,''Amount1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,1,0,''Grade'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,1,0,''Amount2'',0,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,1,0,''Nominal Cost Amount'',0,0,0,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,1,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,1,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,1,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,1,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,1,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,1,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,1,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,1,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,1,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (18,1,0,''Post Id'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

		-- Allowances
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,2,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,2,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,2,1,''Allowance Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,2,1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,2,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,2,1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,2,0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,2,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,2,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,2,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,2,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,2,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,2,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,2,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,2,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,2,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

		-- Loans
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,3,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,3,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,3,1,''Loan Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,3,1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,3,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,3,1,''Period Repayment Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,3,1,''Outstanding Balance'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,3,0,''Repaid Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,3,0,''Reference'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,3,0,''Nominal Account'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,3,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,3,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,3,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,3,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,3,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,3,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,3,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,3,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (18,3,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

		-- Deductions
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,4,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,4,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,4,1,''Allowance Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,4,1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,4,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,4,0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,4,0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,4,0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,4,0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,4,0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,4,0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,4,0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,4,0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,4,0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,4,0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,4,0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,4,0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 3 of 24 - Amending Accord Transfer Definitions Tables'

	-- Update structure of transfer tables - Status Flag
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferTypes')
	and name = 'StatusColumnID'

	if @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes ADD [StatusColumnID] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferTypes SET StatusColumnID = 0'
		EXEC sp_executesql @NVarCommand

	END

	-- Update structure of transfer fields
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferFieldDefinitions')
	and name = 'ConvertData'

	if @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferFieldDefinitions ADD 
					[ConvertData] [bit] NULL'
		EXEC sp_executesql @NVarCommand
	END

	-- Update Data
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 1, Mandatory = 1 WHERE TransferFieldID = 2 AND TransferTypeID = 1'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Nominal Cost Account'' WHERE TransferFieldID = 8 AND TransferTypeID = 1'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 1 WHERE TransferFieldID = 5 AND TransferTypeID = 2'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Nominal Account'' WHERE TransferFieldID = 6 AND TransferTypeID = 2'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 1 WHERE TransferFieldID = 5 AND TransferTypeID = 3'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 1 WHERE TransferFieldID = 6 AND TransferTypeID = 3'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Nominal Account'' WHERE TransferFieldID = 9 AND TransferTypeID = 3'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Deduction Type'' WHERE TransferFieldID = 2 AND TransferTypeID = 4'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 1 WHERE TransferFieldID = 5 AND TransferTypeID = 4'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 0 WHERE TransferFieldID = 6 AND TransferTypeID = 4'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Nominal Account'' WHERE TransferFieldID = 7 AND TransferTypeID = 4'
	EXEC sp_executesql @NVarCommand
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET IsKeyField = 0, Mandatory = 1 WHERE TransferFieldID = 4 AND TransferTypeID = 5'
	EXEC sp_executesql @NVarCommand
	

/* ------------------------------------------------------------- */
PRINT 'Step 4 of 24 - Adding Accord Column Mapping Definitions'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransferFieldMappings]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN

		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransferFieldMappings] (
			[TransferID] [int] NOT NULL ,
			[FieldID] [int] NOT NULL ,
			[HRProValue] [varchar] (100) NOT NULL ,
			[AccordValue] [varchar] (100) NOT NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,7,''Female'',''0'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,7,''Male'',''1'')'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,8,''Single'',''0'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,8,''Married'',''1'')'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''Weekly'',''1'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''2-Weekly'',''2'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''Monthly'',''3'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''4-Weekly'',''4'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''Quarterly'',''5'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''Half Yearly'',''6'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,22,''Yearly'',''7'')'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,23,''Cash or Manual Cheque'',''1'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,23,''Computer Printed Cheque'',''2'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,23,''BACS'',''3'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,23,''CHAPS'',''4'')'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldMappings  (TransferID, FieldID, HRProValue, AccordValue) VALUES (0,23,''BOBS'',''5'')'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 5 of 24 - Adding Accord Transfer Tables'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransactions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransactions] (
			[TransactionID] [int] NOT NULL,
			[TransferType] [smallint] NOT NULL ,
			[TransactionType] [smallint] NOT NULL ,
			[CreatedDateTime] [datetime] NOT NULL ,
			[TransferedDateTime] [datetime] NULL ,
			[Status] [smallint] NOT NULL ,
			[ErrorText] [varchar] (2000) NULL,
			[CompanyCode] [varchar] (255) NULL,
			[EmployeeCode] [varchar] (255) NULL,
			[CreatedUser] [varchar] (100) NOT NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	-- Addition of HRProRecordID
	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysAccordTransactions') and name = 'HRProRecordID'
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransactions ADD [HRProRecordID] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	-- Addition of Archive bit
	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysAccordTransactions') and name = 'Archived'
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransactions ADD [Archived] [bit] NULL'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransactionData]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransactionData] (
			[TransactionID] [int] NOT NULL,
			[FieldID] [smallint] NOT NULL ,
			[OldData] [varchar] (2000) NULL ,
			[NewData] [varchar] (2000) NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransactionWarnings]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransactionWarnings] (
			[TransactionID] [int] NOT NULL,
			[FieldID] [smallint] NOT NULL ,
			[WarningMessage] [varchar] (2000) NULL) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	IF exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysAccordTransactionProcessInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		DROP TABLE [dbo].[ASRSysAccordTransactionProcessInfo]

	SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysAccordTransactionProcessInfo] (
		[SPID] [smallint] NOT NULL ,
		[TransactionID] [numeric](18, 0) NOT NULL,
		[TransferType] [smallint] NOT NULL,
		[RecordID] [int] NOT NULL) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 6 of 24 - Adding Accord Transfer Stored Procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRPopulateAccordTransactionData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRPopulateAccordTransactionData]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRPopulateAccordTransactions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRPopulateAccordTransactions]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransactionData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransactionData]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransaction]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransaction]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPurgeTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPurgeTemp]


	-- spASRAccordPurgeTemp
	SELECT @NVarCommand = 'CREATE PROCEDURE spASRAccordPurgeTemp (
			@piTriggerLevel int,
			@piRecordNo int)
	AS
	BEGIN	
		-- This stored procedure is called from every table trigger and resets the Accord transaction id whenever the trigger level is 1
		IF @piTriggerLevel = 1 DELETE FROM ASRSysAccordTransactionProcessInfo WHERE spid = @@SPID --AND RecordID = @piRecordNo
	END'
	EXEC sp_executesql @NVarCommand


	-- spASRAccordPopulateTransactionData
	SELECT @NVarCommand = 'CREATE PROCEDURE spASRAccordPopulateTransactionData (
		@piTransactionID int,
		@piColumnID int,
		@psOldValue varchar(255),
		@psNewValue varchar(255)
		)
	AS
	BEGIN	
		DECLARE @iRecCount int

		SELECT @iRecCount = COUNT(FieldID) FROM ASRSysAccordTransactionData WHERE @piTransactionID = TransactionID and FieldID = @piColumnID

		-- Insert a record into the Accord Transaction table.	
		IF @iRecCount = 0
			INSERT INTO ASRSysAccordTransactionData
				([TransactionID],[FieldID], [OldData], [NewData])
			VALUES 
				(@piTransactionID,@piColumnID,@psOldValue,@psNewValue)
		ELSE
			UPDATE ASRSysAccordTransactionData SET [OldData] = @psOldValue
				WHERE @piTransactionID = TransactionID and FieldID = @piColumnID
	END'
	EXEC sp_executesql @NVarCommand


	-- spASRAccordPopulateTransaction
	SELECT @NVarCommand = 'CREATE PROCEDURE spASRAccordPopulateTransaction (
	@piTransactionID int OUTPUT,
	@piTransferType int,
	@piTransactionType int,
	@piDefaultStatus int,
	@piHRProRecordID int,
	@iTriggerLevel int)
	AS
	BEGIN	

	-- Return the required user or system setting.
	DECLARE @iCount	integer
	DECLARE @bNewTransaction bit

	SET @piTransactionID = null
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

		-- Insert a record into the Accord Transfer table.
		INSERT INTO ASRSysAccordTransactions
			([TransactionID],[TransferType], [TransactionType], [CreatedUser], [CreatedDateTime], [Status], [HRProRecordID], [Archived])
		VALUES 
			(@piTransactionID, @piTransferType, @piTransactionType, SYSTEM_USER, GETDATE(), @piDefaultStatus, @piHRProRecordID, 0)

		INSERT ASRSysAccordTransactionProcessInfo (SPID, TransactionID,TransferType,RecordID) VALUES (@@SPID, @piTransactionID, @piTransferType, @piHRProRecordID)

	END
	END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 7 of 24 - Adding Accord Transfer Security'

		/* Adding System Permissions for Accord Transfer */	
		DELETE FROM ASRSysPermissionCategories WHERE categoryID = 41
	
		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 41

		IF @iRecCount = 0 
		BEGIN

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(41,
					'ASR Accord Payroll Transfer',
					'',
					10,
					'ACCORD')

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 41

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000C0CCC000000000044CCC040000000000CC000000000000000C000000111000004CC04000999000004CCC00000990000000C500000099000000C50000000990000FCC04C0000099999995C4C00000099000990000000000990F9900000000000990990000000000009999000000000000099900000000000000990000FF470000FE0B0000FF3F0000FFBF00001F1700001F0F00009FCF0000CFCF0000E7C90000F0010000F9CF0000FCCF0000FE4F0000FF0F0000FF8F0000FFCF000000

			DELETE FROM ASRSysPermissionItems WHERE itemid in (145,146,147,148)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (145,'View Transfer Lists',10,41,'VIEWTRANSFER')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (146,'Create Transfers',30,41,'SENDRECORD')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (147,'Archive Transfers',50,41,'DELETE')
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (148,'Block/Unblock Records',20,41,'BLOCK')

		END

		-- Give security to admistrators
		SELECT @iRecCount = count(*)
		FROM ASRSysGroupPermissions
		WHERE itemid IN (145,146,147,148)

		IF @iRecCount = 0 
		BEGIN
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 145, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 146, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 147, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 148, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
		END


/* ------------------------------------------------------------- */
PRINT 'Step 8 of 24 - Adding System Permissions for Outlook Queue'

		/* Adding System Permissions for Match Reports */
		SELECT @iRecCount = count(*)
		FROM ASRSysPermissionCategories
		WHERE categoryID = 40

		IF @iRecCount = 0 
		BEGIN
			--SET IDENTITY_INSERT ASRSysPermissionCategories ON

			/* The record doesn't exist, so create it. */
			INSERT INTO ASRSysPermissionCategories
				(categoryID, 
					description, 
					picture, 
					listOrder, 
					categoryKey)
				VALUES(40,
					'Outlook Calendar Queue',
					'',
					10,
					'OUTLOOKQUEUE')

			--SET IDENTITY_INSERT ASRSysPermissionCategories OFF

			SELECT @ptrval = TEXTPTR(picture) 
			FROM ASRSysPermissionCategories
			WHERE categoryID = 40

			WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C0000000000000000000000000000000000000000000000000008000008000000080800080000000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000000000000000000007F00000000000000FF8F000000000007FF8F8F000000000FF8FF8F8F0000007FF8F8FF8F8F0000FF8FF8F99F8F00077F8F8F98F9F0000CC77F8F9FF9F000000CC77F899F0000000007777F8F0000000000077770000000000000077000000000000000000000000000000000000FFFF0000F9FF0000F07F0000F01F0000E0070000E0010000C0000000C00000008001000080010000C0030000F0030000FC070000FF070000FFCF0000FFFF000000

			DELETE FROM ASRSysPermissionItems WHERE itemid in (144)
			INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
						VALUES (144,'View',10,40,'VIEW')

			DELETE FROM ASRSysGroupPermissions WHERE itemid IN (144)

			INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
				SELECT DISTINCT 144, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		END



/* ------------------------------------------------------------- */
PRINT 'Step 9 of 24 - Amending Email Processing Stored Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASREmailImmediate]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].spASREmailImmediate(@Username varchar(255)) AS
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
					@hResult int,
					@blnEnabled int,
					@RecalculateRecordDesc bit,
					@TableID int,
					@RecipTo varchar(4000),
					@TempText nvarchar(4000)

				/* Loop through all entries which are to be sent */
				DECLARE emailqueue_cursor
				CURSOR LOCAL FAST_FORWARD FOR 
					SELECT QueueID, ASRSysEmailQueue.LinkID, RecordID, ASRSysEmailQueue.ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID
					FROM ASRSysEmailQueue
					INNER JOIN ASRSysEmailLinks ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
					WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
					And (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
					ORDER BY DateDue

				OPEN emailqueue_cursor
				FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

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

					/* Add table name to record descripion if it is a table entry */
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
									EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''
								END
						END
					ELSE IF @TableID > 0
						BEGIN
							SET @sSQL = ''spASRSysEmailAddr''
							IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
								BEGIN
									SELECT @emailDate = getDate()
									EXEC @hResult = @sSQL @RecipTo OUTPUT, @LinkID, 0
									EXEC @hResult = master.dbo.xp_sendmail  @recipients=@RecipTo,  @subject=@columnvalue,  @message=@RecDesc, @no_output=''True''
								END
						END

					IF @hResult = 0
					BEGIN
						UPDATE ASRSysEmailQueue SET DateSent = @emailDate
						WHERE QueueID = @QueueID
					END

					FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc,@RecalculateRecordDesc,@TableID

				END
				CLOSE emailqueue_cursor
				DEALLOCATE emailqueue_cursor

			END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 10 of 24 - Amending Parental Leave Entitlement Calculation'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRParentalLeaveEntitlement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRParentalLeaveEntitlement]

	SELECT @NVarCommand = 'CREATE PROCEDURE dbo.spASRParentalLeaveEntitlement (
			@pdblResult    float OUTPUT,
			@DateOfBirth datetime,
			@AdoptedDate datetime,
			@Disabled bit,
			@Region varchar(8000)
			)
			AS
			BEGIN

			DECLARE @Today datetime
			DECLARE @ChildAge int
			DECLARE @Adopted bit
			DECLARE @YearsOfResponsibility int
			DECLARE @StartDate datetime

			DECLARE @Standard int
			DECLARE @Extended int

			SET @Standard = 65
			SET @Extended = 90
			IF @Region = ''Rep of Ireland''
			BEGIN
				SET @Standard = 70
				SET @Extended = 70
			END


			--Check if we should used the Date of Birth or the Date of Adoption column...
			SET @Adopted = 0
			SET @StartDate = @DateOfBirth
			IF NOT @AdoptedDate IS NULL
			BEGIN
				SET @Adopted = 1
				SET @StartDate = @AdoptedDate
			END

			--Set variables based on this date...
			--(years of responsibility = years since born or adopted)
			SET @Today = getdate()
			EXEC sp_ASRFn_WholeYearsBetweenTwoDates @ChildAge OUTPUT, @DateOfBirth, @Today
			EXEC sp_ASRFn_WholeYearsBetweenTwoDates @YearsOfResponsibility OUTPUT, @StartDate, @Today


			SELECT @pdblResult = CASE
				WHEN @Disabled = 0 And @Adopted = 0 And @ChildAge < 5
					THEN
					@Standard

				WHEN @Disabled = 0 And @Adopted = 1 And @ChildAge < 18
					And @YearsOfResponsibility < 5 THEN
					@Standard

				WHEN @Disabled = 1 And @Adopted = 0 And @ChildAge < 18 
					And DateDiff(d,''12/15/1994'',@DateOfBirth) >= 0 THEN
					@Extended

				WHEN @Disabled = 1 And @Adopted = 1 And @ChildAge < 18 
				And DateDiff(d,''12/15/1994'',@AdoptedDate) >= 0 THEN
					@Extended

				ELSE
					0
				END

			END'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 11 of 24 - Amending Support contact numbers and Support WWW address'
	
	UPDATE ASRSysSystemSettings
	SET SettingValue = '01582 714820'
	WHERE section = 'support'
		AND SettingKey = 'telephone no' 
		AND SettingValue = '01582 714814'

	UPDATE ASRSysSystemSettings
	SET SettingValue = '01582 714814'
	WHERE section = 'support'
		AND SettingKey = 'fax' 
		AND SettingValue = '01582 714820'

	UPDATE ASRSysSystemSettings
		SET SettingValue = 'http://www.asr.co.uk/customer'
		WHERE (SettingKey = 'webpage' and SettingValue = 'http://www.asr.co.uk')
			

/* ------------------------------------------------------------- */
PRINT 'Step 12 of 24 - Dropping redundant ASRSysSSIntranetLinks columns'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		and name = 'viewID'
		
	if @iRecCount > 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks DROP COLUMN viewID'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 24 - Amending Email ID Allocation'

	--Ensure that EMail ID allocation is correct.
	if not exists(select count(*) from asrsyssystemsettings where [section] = 'autoid' and settingkey = 'emailaddress')
	  insert asrsyssystemsettings([section],[settingkey],[settingvalue]) values('autoid','emailaddress',null)

	if exists(select count(*) from asrsysemailaddress)
	  update asrsyssystemsettings set settingvalue = (select max(emailid) from asrsysemailaddress)
	  where [section] = 'autoid' and settingkey = 'emailaddress' and settingvalue is null
	else
	  update asrsyssystemsettings set settingvalue = 0
	  where [section] = 'autoid' and settingkey = 'emailaddress' and settingvalue is null


/* ------------------------------------------------------------- */
PRINT 'Step 14 of 24 - Adding default to Columns table (KB000155)'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DF_ASRSysColumns_QAddressEnabled]') and OBJECTPROPERTY(id, N'IsConstraint') = 1)
	BEGIN
		ALTER TABLE ASRSysColumns
			DROP CONSTRAINT DF_ASRSysColumns_QAddressEnabled
	END

	ALTER TABLE ASRSysColumns ADD CONSTRAINT
		DF_ASRSysColumns_QAddressEnabled DEFAULT (0) FOR QAddressEnabled
	
	UPDATE ASRSysColumns SET QAddressEnabled = 0 WHERE QAddressEnabled IS NULL


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 24 - Child Table Security Permissions (KB000158)'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRIsSysSecMgr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRIsSysSecMgr]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRIsSysSecMgr (
				@psGroupName		sysname,
				@pfSysSecMgr		bit	OUTPUT
			)
			AS
			BEGIN
				DECLARE @iUserGroupID integer

				/* Get the current user''s group ID. */
				SELECT @iUserGroupID = sysusers.gid
				FROM sysusers
				WHERE sysusers.name = @psGroupName

				SELECT @pfSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
				FROM ASRSysGroupPermissions
				INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
				INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
				WHERE sysusers.uid = @iUserGroupID
					AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER'' OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
					AND ASRSysGroupPermissions.permitted = 1
					AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''
			END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 16 of 24 - Convert Character To Numeric stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ConvertCharacterToNumeric]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRFn_ConvertCharacterToNumeric]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_ConvertCharacterToNumeric (
				@pdblResult		float OUTPUT,
				@psStringToConvert  	varchar(8000)
			)
			AS
			BEGIN
				IF (@psStringToConvert is null) OR (len(@psStringToConvert) = 0)
				BEGIN
					SET @pdblResult = 0
				END
				ELSE
				BEGIN
					IF isNumeric(@psStringToConvert) = 1
					BEGIN
						SET @pdblResult = convert(float, convert(money, @psStringToConvert))
					END
					ELSE
					BEGIN
						SET @pdblResult = 0
					END
				END
			END'

	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 17 of 24 - Drop Unique Object stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRDropUniqueObject]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRDropUniqueObject]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRDropUniqueObject(
				@psUniqueObjectName sysname,
				@piType integer)
			AS
			BEGIN
				DECLARE 
					@sCommandString	nvarchar(4000),
					@sCleanUniqueObjectName	sysname
			
				/* Clean the input string parameters. */
				SET @sCleanUniqueObjectName = @psUniqueObjectName
				IF len(@sCleanUniqueObjectName) > 0 SET @sCleanUniqueObjectName = replace(@sCleanUniqueObjectName, '''''''', '''''''''''')
													
				IF (EXISTS (SELECT * 
										FROM sysobjects 
										WHERE name = @psUniqueObjectName))
				BEGIN
					IF @piType = 3 
					BEGIN
						SET @sCommandString = ''DROP TABLE '' + @sCleanUniqueObjectName
					END
			
					IF @piType = 4
					BEGIN
						SET @sCommandString = ''DROP PROCEDURE '' + @sCleanUniqueObjectName
					END 
			
					EXECUTE sp_executesql @sCommandString
			  END
				
				DELETE FROM ASRSysSQLObjects 
				WHERE Name = @psUniqueObjectName 
					AND Type = @piType
					AND Owner = SYSTEM_USER
			END'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 18 of 24 - Amend Send Message Stored Procedure'


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRSendMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRSendMessage]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRSendMessage 
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

		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 19 of 24 - Add column to Imports'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'AsrSysImportName')
	and name = 'IgnoreLastLine'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE AsrSysImportName ADD 
			[IgnoreLastLine] [bit] null'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @NVarCommand = 'UPDATE AsrSysImportName SET IgnoreLastLine = 0 WHERE IgnoreLastLine IS NULL'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 20 of 24 - Adding Self-service Intranet Hidden Groups Table'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysSSIHiddenGroups]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 
			'CREATE TABLE [dbo].[ASRSysSSIHiddenGroups] (
				[LinkID] [int] NOT NULL ,
				[GroupName] [varchar] (256)) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 21 of 24 - Update Module Setup'

	UPDATE ASRSysModuleSetup
	SET parameterType = 'PType_ViewID'
	WHERE parameterType <> 'PType_ViewID' 
	AND parameterKey = 'Param_BulkBookingDefaultView'

/* ------------------------------------------------------------- */
PRINT 'Step 22 of 24 - Update Data Permission Audit'

	UPDATE ASRSysAuditPermissions
	SET Action = 'Deny' WHERE Action = 'Revoke'

/* ------------------------------------------------------------- */
PRINT 'Step 23 of 24 - Update System Permissions Table'

SELECT @iRecCount = COUNT(*)
FROM ASRSysPermissionItems
WHERE itemID = 149

IF @iRecCount = 0
BEGIN
	INSERT INTO ASRSysPermissionItems
		(itemID, description, listOrder, categoryID, itemKey)
	VALUES (149, 'Self-service Intranet', 80, 1, 'SSINTRANET')

	INSERT INTO ASRSysGroupPermissions 
		SELECT 149, groupname, 1 
		FROM ASRSysGroupPermissions
		WHERE (itemID = 4 OR itemID = 100)
			AND permitted = 1
END

UPDATE ASRSysPermissionItems
SET description = 'Data Manager Intranet (multiple record access)'
WHERE itemID = 4

UPDATE ASRSysPermissionItems
SET description = 'Data Manager Intranet (single record access)'
WHERE itemID = 100

UPDATE ASRSysPermissionCategories
SET description = 'Data Manager Intranet'
WHERE categoryID = 19

/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 24 of 24 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.18')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.18.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.18')


SELECT @NVarCommand = 'USE master
GRANT ALL ON master..xp_LoginConfig TO public
GRANT ALL ON master..xp_EnumGroups TO public
GRANT ALL ON master..xp_StartMail TO public
GRANT ALL ON master..xp_SendMail TO public'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE '+@DBName
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.18 Of HR Pro'
