CREATE PROCEDURE [dbo].[spASRGetCurrentUsersFromMaster]
AS
BEGIN

   SET NOCOUNT ON;

   DECLARE @login_time datetime;

	DECLARE @processes TABLE (
		[hostname]		nvarchar(128),
		[loginame]		nvarchar(128),
		program_name	nvarchar(128),
		host_process_id	integer,
		security_id		varbinary(85),
		login_time		datetime,
		spid			smallint,
		uid				integer)

   SELECT TOP 1 @login_time = l.Lock_Time
   FROM ASRSysLock l
		INNER JOIN sys.dm_exec_sessions es on es.session_id = l.spid
         AND es.login_time = l.Login_Time
   WHERE L.Priority < 3
   ORDER BY l.Priority;

   SET @login_time = ISNULL(@login_time, GETDATE());

	-- Fat Clients
   INSERT @processes
	SELECT es.host_name, es.login_name, es.program_name, es.host_process_id
        , es.security_id, es.login_time, es.session_id, 0
   FROM sys.dm_exec_sessions es
   WHERE es.program_name LIKE 'OpenHR%'
     AND es.program_name NOT LIKE 'OpenHR Web%'
     AND es.program_name NOT LIKE 'OpenHR Workflow%'
     AND es.program_name NOT LIKE 'OpenHR Mobile%'
     AND es.program_name NOT LIKE 'OpenHR Outlook%'
     AND es.program_name NOT LIKE 'System Framework Assembly%'
     AND es.program_name NOT LIKE 'OpenHR Intranet Embedding%'
     AND (es.login_Time < @login_time)

	-- Thin clients
   INSERT @processes
	SELECT clientmachine, username, 'OpenHR Web', '', userSID, loginTime, 999, 0
         FROM ASRSysCurrentLogins

	SELECT * FROM @processes ORDER BY loginame;

END