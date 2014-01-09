CREATE PROCEDURE [dbo].[spASRIntGetUniqueExpressionID](
	@settingkey AS varchar(50),
	@settingvalue AS integer OUTPUT)
AS
BEGIN

	SELECT @settingvalue = [SettingValue] FROM [asrsyssystemsettings] WHERE [Section] = 'AUTOID' AND [SettingKey] = @settingkey;

	IF @settingvalue IS NULL
		INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES ('AUTOID', @settingkey, 1);	
	ELSE
	BEGIN
		SET @settingvalue = @settingvalue + 1
		UPDATE ASRSysSystemSettings SET [SettingValue] = @settingvalue  WHERE [Section] ='AUTOID' AND [SettingKey] = @settingkey;
	END

END

