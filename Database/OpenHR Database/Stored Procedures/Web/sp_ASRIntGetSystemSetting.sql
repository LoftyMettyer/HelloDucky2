CREATE PROCEDURE [dbo].[sp_ASRIntGetSystemSetting] (
	@psNewSection		varchar(255),			/* Section value in the new ASRSysSystemSettings table. */
	@psNewKey			varchar(255),			/* Key value in the new ASRSysSystemSettings table. */
	@psOldColumnName	varchar(255),			/* Column name in the old ASRSysConfig table. */
	@psResult			varchar(MAX)	OUTPUT,
	@pfNewSettingFound	bit				OUTPUT,
	@pfOldSettingFound	bit				OUTPUT
)
AS
BEGIN
	DECLARE
		@iCount					integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@sValue					varchar(MAX);

	SET @pfNewSettingFound = 0;
	SET @pfOldSettingFound = 0;
	IF (@psOldColumnName IS NULL) SET @psOldColumnName = '';

	/* Clean the input string parameters. */
	IF len(@psNewSection) > 0 SET @psNewSection = replace(@psNewSection, '''', '''''');
	IF len(@psNewKey) > 0 SET @psNewKey = replace(@psNewKey, '''', '''''');

	/* Check if the 'ASRSysSystemSettings' table exists. */
	SELECT @iCount = count(Name)
	FROM sysobjects 
	WHERE name = 'ASRSysSystemSettings';
		
	IF @iCount = 1
	BEGIN
		/* The ASRSysSystemSettings table exists. See if the required records exists in it. */
		SET @sTempExecString = 'SELECT @sValue = settingValue' +
			' FROM ASRSysSystemSettings' +
			' WHERE section = ''' + @psNewSection + '''' +
			' AND settingKey = ''' + @psNewKey +'''';
		SET @sTempParamDefinition = N'@sValue varchar(MAX) OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT;
	
		IF NOT @sValue IS NULL
		BEGIN
			SET @psResult = @sValue;
			SET @pfNewSettingFound = 1;
		END
	END

	IF @pfNewSettingFound = 0
	BEGIN
		SELECT @iCount = count(syscolumns.name)
		FROM syscolumns 
		INNER JOIN sysobjects ON syscolumns.id = sysobjects.id
		WHERE syscolumns.name = @psOldColumnName
			AND sysobjects.name = 'ASRSysConfig';

		IF @iCount = 1
		BEGIN
			/* Clean the input string parameter. */
			IF len(@psOldColumnName) > 0 SET @psOldColumnName = replace(@psOldColumnName, '''', '''''');

			SET @sTempExecString = 'SELECT @sValue = convert(varchar(8000), ' + @psOldColumnName + ') FROM ASRSysConfig';
			SET @sTempParamDefinition = N'@sValue varchar(100) OUTPUT';
			EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT;

			IF NOT @sValue IS NULL
			BEGIN
				SET @psResult = @sValue;
				SET @pfOldSettingFound = 1;
			END
		END
	END
END