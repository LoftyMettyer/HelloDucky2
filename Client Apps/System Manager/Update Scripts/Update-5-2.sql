/* --------------------------------------------------- */
/* Update the database from version 5.0 to version 5.1 */
/* --------------------------------------------------- */

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
	
DECLARE @sSPCode nvarchar(MAX)


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
IF (@sDBVersion <> '5.1') and (@sDBVersion <> '5.2')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion < 9)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


PRINT 'Step - Updating Paternity calculations'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfstat_ParentalLeaveEntitlement]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement];


	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement] (
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
			@Adopted				bit,
			@YearsOfResponsibility	integer,
			@StartDate				datetime,
			@Standard				integer,
			@Extended				integer;

		SET @Standard = 65;
		SET @Extended = 90;
		SET @Today = GETDATE();
				
		IF @Region = ''Rep of Ireland''
		BEGIN
			SET @Standard = 70;
			SET @Extended = 70;
		END

		IF DATEDIFF(d,''03-08-2013'', @Today) >= 0
		BEGIN
			SET @Standard = 90;
			SET @Extended = 90;
		END

		-- Check if we should used the Date of Birth or the Date of Adoption column...
		SET @Adopted = 0;
		SET @StartDate = @DateOfBirth;
		IF NOT @AdoptedDate IS NULL
		BEGIN
			SET @Adopted = 1;
			SET @StartDate = @AdoptedDate;
		END

		-- Set variables based on this date...
		--( years of responsibility = years since born or adopted)
		SELECT @ChildAge = [dbo].[udfsys_wholeyearsbetweentwodates](@DateOfBirth, @Today);
		SELECT @YearsOfResponsibility = [dbo].[udfsys_wholeyearsbetweentwodates](@StartDate, @Today);

		SELECT @pdblResult = CASE
			WHEN @Disabled = 0 AND @Adopted = 0 AND @ChildAge < 5
				THEN @Standard
			WHEN @Disabled = 0 AND @Adopted = 1 AND @ChildAge < 18
				AND @YearsOfResponsibility < 5 THEN	@Standard
			WHEN @Disabled = 1 AND @Adopted = 0 AND @ChildAge < 18 
				AND DATEDIFF(d,''12/15/1994'',@DateOfBirth) >= 0 THEN @Extended
			WHEN @Disabled = 1 AND @Adopted = 1 AND @ChildAge < 18 
				AND DATEDIFF(d,''12/15/1994'',@AdoptedDate) >= 0 THEN @Extended
			ELSE 0
			END;

		RETURN ISNULL(@pdblResult,0);

	END'
	EXECUTE sp_executeSQL @sSPCode;




/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '5.2';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v5.2')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v5.2 Of OpenHR'
