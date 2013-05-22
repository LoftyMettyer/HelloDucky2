CREATE PROCEDURE [dbo].[spASRGetServerDLLVersion]
(
	@strVersion varchar(255) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON

	DECLARE @objectToken int
	DECLARE @hResult int

	IF EXISTS(SELECT SettingValue
              FROM   ASRSysSystemSettings
              WHERE  [Section] = 'server dll'
                AND  [SettingKey] = 'disable check'
                AND  [SettingValue] = '1')
	BEGIN

		SELECT @strVersion = [SettingValue]
		FROM   ASRSysSystemSettings
		WHERE  [Section] = 'server dll'
		  AND  [SettingKey] = 'minimum version'

	END
	ELSE
	BEGIN
		-- Create Server DLL object
		EXEC @hResult = sp_OACreate 'vbpHRProServer.clsGeneral', @objectToken OUTPUT
		IF @hResult = 0
			EXEC @hResult = sp_OAMethod @objectToken, 'GetVersion', @strVersion OUTPUT
		ELSE
			SET @strVersion = '0.0.0'
				
		EXEC sp_OADestroy @objectToken
			
	END

END
