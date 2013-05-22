CREATE PROCEDURE [dbo].[spASRSaveSetting] (
	@psSection		varchar(50),
	@psKey			varchar(50),
	@psValue		varchar(200)	
)
AS
BEGIN

	/* Save the given system setting. */
	DELETE FROM [dbo].[ASRSysSystemSettings]
	WHERE section = @psSection
		AND [settingKey] = @psKey;
	
	INSERT INTO [dbo].[ASRSysSystemSettings]
		([section], [settingKey], [settingValue])
	VALUES (@psSection, @psKey, @psValue);
	
END