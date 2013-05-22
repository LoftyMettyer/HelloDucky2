CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]
(
	@iLoginCount	integer OUTPUT,
	@psLoginName	varchar(MAX)
)
AS
BEGIN

	DECLARE @sSQLVersion	integer,
			@Mode			smallint;

	IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id('sp_ASRIntCheckPolls') AND sysstat & 0xf = 4)
	BEGIN
		EXEC sp_ASRIntCheckPolls;
	END

	SELECT @sSQLVersion = dbo.udfASRSQLVersion();
	SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'ProcessAccount' AND [SettingKey] = 'Mode';
	IF @@ROWCOUNT = 0 SET @Mode = 0
	
	IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER('sysadmin') = 1)		
	BEGIN
		SELECT @iLoginCount = dbo.[udfASRNetCountCurrentLogins](@psLoginName);
	END
	ELSE
	BEGIN

		SELECT @iLoginCount = COUNT(*)
		FROM master..sysprocesses p
		WHERE p.program_name LIKE 'OpenHR%'
			AND	p.program_name NOT LIKE 'OpenHR Workflow%'
			AND	p.program_name NOT LIKE 'OpenHR Outlook%'
			AND	p.program_name NOT LIKE 'OpenHR Server.Net%'
			AND	p.program_name NOT LIKE 'OpenHR Intranet Embedding%'
		    AND p.loginame = @psLoginName;
	END

END