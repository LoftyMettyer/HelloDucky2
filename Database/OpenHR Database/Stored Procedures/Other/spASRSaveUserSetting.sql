CREATE PROCEDURE [dbo].[spASRSaveUserSetting]
(
	@sSection		varchar(50),
	@sSettingKey	varchar(50),
	@sSettingValue	varchar(255)
)
AS
BEGIN
	SET NOCOUNT ON;

	IF EXISTS(SELECT [SettingValue] FROM ASRSysUserSettings WHERE Section = @sSection	 AND SettingKey = @sSettingKey AND UserName = System_User)
		UPDATE ASRSysUserSettings SET [SettingValue] = @sSettingValue WHERE Section = @sSection AND SettingKey = @sSettingKey AND UserName = System_User;
	ELSE
		INSERT ASRSysUserSettings ([Section], [SettingKey], [UserName], [SettingValue]) VALUES (@sSection, @sSettingKey, System_User, @sSettingValue);

END
