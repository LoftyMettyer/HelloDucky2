CREATE PROCEDURE [dbo].[spASRGetUserSetting]
	(
		@sSection		varchar(50),
		@sSettingKey	varchar(50),
		@sSettingValue	varchar(255) OUTPUT
	)
	AS
	BEGIN
		SET NOCOUNT ON

		SELECT @sSettingValue = [SettingValue]
		FROM ASRSysUserSettings
		WHERE Section = @sSection
			AND SettingKey = @sSettingKey
			AND UserName = System_User
	END