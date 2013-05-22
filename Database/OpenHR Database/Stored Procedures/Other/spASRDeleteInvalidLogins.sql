CREATE PROCEDURE [dbo].[spASRDeleteInvalidLogins]
	(@pstrDomainName	varchar(100))
AS
BEGIN

	DECLARE @cursLogins cursor,
			@loginName	nvarchar(255);

	-- Are we privileged enough to run this script
	IF IS_SRVROLEMEMBER('securityadmin') = 0 and IS_SRVROLEMEMBER('sysadmin') = 0
	BEGIN
		RETURN 0;
	END

	-- Lets get the invalid accounts into a swish little cursor
	SET @cursLogins = CURSOR LOCAL FAST_FORWARD FOR
	SELECT loginName from master.dbo.syslogins
		WHERE isntname = 1
		AND loginname like @pstrDomainName + '\%'
		AND ((sid <> SUSER_SID(loginname) and SUSER_SID(loginname) is not null)	OR SUSER_SID(loginname) is null);

	-- Now lets get rid of the invalid accounts
	OPEN @cursLogins;
	FETCH NEXT FROM @cursLogins INTO @LoginName;
	WHILE (@@fetch_status = 0)
	BEGIN
		EXEC sp_revokelogin @LoginName;
		FETCH NEXT FROM @cursLogins INTO @LoginName;
	END

	-- Tidy Up
	CLOSE @cursLogins;
	DEALLOCATE @cursLogins;

	RETURN 0;

END