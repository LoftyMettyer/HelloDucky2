CREATE PROCEDURE [dbo].[sp_ASRFn_IsOvernightProcess]
(
    @result integer OUTPUT
)
AS
BEGIN

	DECLARE @iCount					integer;
	DECLARE	@sTempExecString		nvarchar(MAX);
	DECLARE	@sTempParamDefinition	nvarchar(500);
	DECLARE	@sValue					varchar(MAX);

	/* Check if the 'ASRSysSystemSettings' table exists. */
	SELECT @iCount = COUNT(*)
		FROM sysobjects 
		WHERE name = 'ASRSysSystemSettings';
		
	IF @iCount = 1
	BEGIN
		/* The ASRSysSystemSettings table exists. See if the required records exists in it. */
		SET @sTempExecString = 'SELECT @sValue = settingValue' +
			' FROM ASRSysSystemSettings' +
			' WHERE section = ''database''' +
			' AND settingKey = ''updatingdatedependantcolumns''';
		SET @sTempParamDefinition = N'@sValue varchar(MAX) OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT;
	
		IF NOT @sValue IS NULL
		BEGIN
			SET @result = CONVERT(bit, @sValue);
		END
	END
	ELSE
	BEGIN
		SELECT @result = [UpdatingDateDependentColumns] FROM [dbo].[ASRSysConfig];
	END
END