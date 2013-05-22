CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
(
	@piCount integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sSQLVersion	integer;
	DECLARE @Mode			smallint;

	IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id('sp_ASRIntCheckPolls') AND sysstat & 0xf = 4)
	BEGIN
		EXEC sp_ASRIntCheckPolls;
	END

	SELECT @sSQLVersion = dbo.udfASRSQLVersion();
	SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'ProcessAccount' AND [SettingKey] = 'Mode';
	IF @@ROWCOUNT = 0 SET @Mode = 0;

	IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER('sysadmin') = 1)
	BEGIN
		SELECT @piCount = dbo.[udfASRNetCountCurrentUsersInApp](APP_NAME());
	END
	ELSE
	BEGIN

		SELECT @piCount = COUNT(p.Program_Name)
		FROM     master..sysprocesses p
		JOIN     master..sysdatabases d
		  ON     d.dbid = p.dbid
		WHERE    p.program_name = APP_NAME()
		  AND    d.name = db_name()
		GROUP BY p.program_name;
	END

END