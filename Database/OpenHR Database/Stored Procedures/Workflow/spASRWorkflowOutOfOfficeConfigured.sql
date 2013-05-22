CREATE PROCEDURE [dbo].[spASRWorkflowOutOfOfficeConfigured]
(
    @pfOutOfOfficeConfigured bit output
)
AS
BEGIN
	DECLARE	@iCount	integer;

	-- Check if the SP that checks if the current user is OutOfOffice exists
	SELECT @iCount = COUNT(*)
	FROM sysobjects
	WHERE id = object_id('spASRWorkflowOutOfOfficeCheck')
		AND sysstat & 0xf = 4;

	IF @iCount > 0 
	BEGIN
		-- Check if the SP that sets/resets the current user to be OutOfOffice exists
		SELECT @iCount = COUNT(*)
		FROM sysobjects
		WHERE id = object_id('spASRWorkflowOutOfOfficeSet')
			AND sysstat & 0xf = 4;
	END

	IF @iCount > 0 
	BEGIN
		-- Check if the the Activation column has been defined
		SELECT @iCount = convert(integer, isnull(parameterValue, '0'))
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_WORKFLOW'
			AND parameterKey = 'Param_DelegationActivatedColumn';
	END

	SET @pfOutOfOfficeConfigured = 
	CASE	
		WHEN @iCount > 0 THEN 1
		ELSE 0
	END;
END

