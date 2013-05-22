CREATE PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
(
	@psAppName		varchar(MAX) OUTPUT,
	@psUserName		varchar(MAX)
)
AS
BEGIN

    IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id('sp_ASRIntCheckPolls') AND sysstat & 0xf = 4)
    BEGIN
        EXEC sp_ASRIntCheckPolls;
    END

    SELECT TOP 1 @psAppName = rtrim(p.program_name)
    FROM master..sysprocesses p
    WHERE p.program_name LIKE 'OpenHR%'
		AND	p.program_name NOT LIKE 'OpenHR Workflow%'
		AND	p.program_name NOT LIKE 'OpenHR Outlook%'
		AND	p.program_name NOT LIKE 'OpenHR Server.Net%'
		AND	p.program_name NOT LIKE 'OpenHR Intranet Embedding%'
		AND	p.loginame = @psUsername
    GROUP BY p.hostname
           , p.loginame
           , p.program_name
           , p.hostprocess
    ORDER BY p.loginame;

END