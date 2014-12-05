CREATE PROCEDURE [dbo].[spASRIntGetSetting] (
	@psSection		varchar(MAX),
	@psKey			varchar(MAX),
	@psDefault		varchar(MAX),
	@pfUserSetting	bit,
	@psResult		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return the required user or system setting. */
	DECLARE	@iCount	integer;

	IF @pfUserSetting = 1
	BEGIN
		SELECT @iCount = COUNT(userName)
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = settingValue 
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;
	END
	ELSE
	BEGIN
		SELECT @iCount = COUNT(settingKey)
		FROM ASRSysSystemSettings
		WHERE section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = settingValue 
		FROM ASRSysSystemSettings
		WHERE section = @psSection		
			AND settingKey = @psKey;
	END

	IF @iCount = 0
	BEGIN
		SET @psResult = @psDefault;	
	END
END