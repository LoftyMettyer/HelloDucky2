CREATE PROCEDURE [dbo].[spASRSysUpdateFrameworkVersions]
AS
BEGIN

	SET NOCOUNT ON

	DECLARE @clrversion varchar(255)

	BEGIN TRAN

	DELETE FROM ASRSysSystemSettings
	WHERE [Section] = '.NET Framework'
	AND [SettingKey] = 'CLR Version'
	
	SELECT @clrversion = dbo.udfASRNetCLRVersion()

	IF @@ERROR <> 0
	BEGIN
		RAISERROR(N'Unable to detect .NET Framework versions', 16, 1)
		ROLLBACK
	END
	ELSE
	BEGIN
		INSERT INTO ASRSysSystemSettings 
		VALUES ('.NET Framework','CLR Version',@clrversion)

		IF @@ERROR <> 0
		BEGIN
			RAISERROR(N'Unable to update .NET Framework versions', 16, 1)
			ROLLBACK
		END
		ELSE COMMIT TRAN
	END
	
	SET NOCOUNT OFF

END