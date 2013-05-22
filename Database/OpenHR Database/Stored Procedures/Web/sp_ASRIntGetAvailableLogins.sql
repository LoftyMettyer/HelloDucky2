CREATE PROCEDURE [dbo].[sp_ASRIntGetAvailableLogins] 
AS
BEGIN

	SELECT name FROM sys.server_principals
	WHERE NOT EXISTS (SELECT sysusers.sid
						FROM sysusers
						WHERE sysusers.sid = sys.server_principals.sid)
	AND TYPE IN ('S', 'U')
	AND name != 'NT AUTHORITY\SYSTEM'
	AND is_disabled = 0

END
