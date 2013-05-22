CREATE PROCEDURE [dbo].[spASRIntActivateModule](
	@sModule	varchar(50),
	@bLicensed	bit OUTPUT
)
AS
BEGIN
	DECLARE @iCustomerNo		integer,
			@sModuleCode		varchar(100),
			@sValue				varchar(MAX),
			@fNewModuleCode		bit,
			@fNewSettingFound	bit,
			@fOldSettingFound	bit,
			@iModule			int,
			@iSuccessFlag		smallint,
			@objectToken 		integer,
			@hResult 			integer,
			@iValue				integer,
			@sErrorMessage		varchar(MAX),
			@sSQLVersion		int;

	IF @sModule = 'PERSONNEL' SET @iModule = 1;
	IF @sModule = 'RECRUITMENT' SET @iModule = 2;
	IF @sModule = 'ABSENCE' SET @iModule = 4;
	IF @sModule = 'TRAINING' SET @iModule = 8;
	IF @sModule = 'INTRANET' SET @iModule = 16;
	IF @sModule = 'AFD' SET @iModule = 32;
	IF @sModule = 'FULLSYSTEMMANGER' SET @iModule = 64;
	IF @sModule = 'CMG' SET @iModule = 128;
	IF @sModule = 'QADDRESS' SET @iModule = 256;
	IF @sModule = 'ACCORD' SET @iModule = 512;
	IF @sModule = 'WORKFLOW' SET @iModule = 1024;
	IF @sModule = 'VERSIONONE' SET @iModule = 2048;
	IF @sModule = 'MOBILE' SET @iModule = 4096;

	/* Get the module license code. */
	SET @sModuleCode = '';
	SET @fNewModuleCode = 0;
	EXEC [dbo].[sp_ASRIntGetSystemSetting] 'Licence', 'Key', 'moduleCode', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT;
	IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
	BEGIN
		SET @sModuleCode = @sValue;
		SET @fNewModuleCode = @fNewSettingFound;
	END

	SELECT @bLicensed = dbo.udfASRNetIsModuleLicensed(@sModuleCode,@iModule);

END