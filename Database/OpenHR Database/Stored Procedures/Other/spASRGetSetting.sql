CREATE PROCEDURE [dbo].[spASRGetSetting] (
	@psSection		varchar(25),
	@psKey			varchar(255),
	@psDefault		varchar(MAX),
	@pfUserSetting	bit,
	@psResult		varchar(MAX) OUTPUT
)
AS
BEGIN
	/* Return the required user or system setting. */
	DECLARE	@iCount	integer;

	IF @pfUserSetting = 1
	BEGIN
		SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysUserSettings]
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = ISNULL(settingValue , '')
		FROM [dbo].[ASRSysUserSettings]
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;
	END
	ELSE
	BEGIN
		SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysSystemSettings]
		WHERE section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = ISNULL(settingValue , '')
		FROM [dbo].[ASRSysSystemSettings]
		WHERE section = @psSection		
			AND settingKey = @psKey;
	END

	IF @iCount = 0
	BEGIN
		SET @psResult = @psDefault;	
	END
END