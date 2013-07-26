CREATE PROCEDURE [dbo].[spsys_setsystemsetting](
	@section AS varchar(50),
	@settingkey AS varchar(50),
	@settingvalue AS nvarchar(MAX))
AS
BEGIN
	IF EXISTS(SELECT [SettingValue] FROM [asrsyssystemsettings] WHERE [Section] = @section AND [SettingKey] = @settingkey)
		UPDATE ASRSysSystemSettings SET [SettingValue] = @settingvalue WHERE [Section] = @section AND [SettingKey] = @settingkey;
	ELSE
		INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES (@section, @settingkey, @settingvalue);	
END