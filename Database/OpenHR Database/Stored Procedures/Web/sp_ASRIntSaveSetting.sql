CREATE PROCEDURE [dbo].[sp_ASRIntSaveSetting] (
	@psSection		varchar(255),
	@psKey			varchar(255),
	@pfUserSetting	bit,
	@psValue		varchar(MAX)	
)
AS
BEGIN
	/* Save the given user or system setting. */
	IF @pfUserSetting = 1
	BEGIN
		DELETE FROM [dbo].[ASRSysUserSettings]
		WHERE section = @psSection
			AND settingKey = @psKey
			AND userName = SYSTEM_USER;

		INSERT INTO [dbo].[ASRSysUserSettings]
			(section, settingKey, settingValue, userName)
		VALUES (@psSection, @psKey, @psValue, SYSTEM_USER);
	END
	ELSE
	BEGIN
		DELETE FROM [dbo].[ASRSysSystemSettings]
		WHERE section = @psSection
			AND settingKey = @psKey;

		INSERT INTO [dbo].[ASRSysSystemSettings]
			(section, settingKey, settingValue)
		VALUES (@psSection, @psKey, @psValue);
	END
END