CREATE PROCEDURE [dbo].[spASRIntGetLicenceInfo](
	@fSelfService		integer,
	@piSuccessFlag		integer			OUTPUT,
	@fIntranetEnabled	bit				OUTPUT,
	@iSSUsers			integer			OUTPUT,
	@iFullUsers			integer			OUTPUT,
	@iSSIUsers			integer			OUTPUT,
	@psErrorMessage		varchar(MAX)	OUTPUT
)
AS
BEGIN
	DECLARE @iCustomerNo		integer,
			@sModuleCode		varchar(100),
			@sValue				varchar(MAX),
			@fNewModuleCode		bit,
			@fNewSettingFound	bit,
			@fOldSettingFound	bit,
			@iSuccessFlag		smallint,
			@objectToken 		integer,
			@hResult 			integer,
			@iValue				integer,
			@iSQLVersion		int;

	SELECT @iSQLVersion = dbo.udfASRSQLVersion();

	/* Get the module license code. */
	SET @sModuleCode = '';
	SET @fNewModuleCode = 0;
	EXEC sp_ASRIntGetSystemSetting 'Licence', 'Key', 'moduleCode', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT;
	IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
	BEGIN
		SET @sModuleCode = @sValue;
		SET @fNewModuleCode = @fNewSettingFound;
	END

	SET @psErrorMessage = '';
	SELECT @fIntranetEnabled = dbo.udfASRNetIsModuleLicensed(@sModuleCode,16);
	IF @fIntranetEnabled = 0
		SET @psErrorMessage = 'You are not licensed to use the intranet module.';
	ELSE
	BEGIN
		SET @piSuccessFlag = 1
		SELECT @iSSIUsers = dbo.udfASRNetGetLicenceKey(@sModuleCode,'SSISUsers');
		SELECT @iFullUsers = dbo.udfASRNetGetLicenceKey(@sModuleCode,'DMIMUsers');
		SELECT @iSSUsers = dbo.udfASRNetGetLicenceKey(@sModuleCode,'DMISUsers');
	END

END