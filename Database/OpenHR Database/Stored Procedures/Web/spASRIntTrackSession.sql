CREATE PROCEDURE dbo.[spASRIntTrackSession](
	@IISServer nvarchar(255),
	@SessionID nvarchar(255),
	@UserName nvarchar(255),
	@SecurityGroup varchar(255),
	@HostName varchar(255),
	@WebArea varchar(20),
	@TrackType tinyint)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @LoginTime	datetime = GETDATE();

	-- Current user tracking
	MERGE INTO ASRSysCurrentSessions AS Target 
	USING (VALUES 
		(@IISServer, @SessionID, @UserName, @HostName, @WebArea) 
	) AS Source (IISServer, SessionID, Username, HostName, WebArea) 
		ON Target.SessionID = Source.SessionID
	WHEN MATCHED AND @TrackType = 1 THEN 
		UPDATE SET webArea = @WebArea, Username = @UserName, HostName = @HostName, IISServer = @IISServer
	WHEN MATCHED AND @TrackType IN (2, 3, 4, 5, 6, 8) THEN
		DELETE
	WHEN NOT MATCHED BY TARGET AND @TrackType = 1 THEN 
		INSERT (IISServer, SessionID, Username, HostName, WebArea)
		VALUES (@IISServer, @SessionID, @UserName, @HostName, @WebArea);

	-- Track in audit log
	INSERT INTO [dbo].[ASRSysAuditAccess]	(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action) 
		VALUES (@LoginTime, @SecurityGroup, @UserName, @HostName, @WebArea
			, CASE @TrackType
				WHEN 1 THEN 'Log In'
				WHEN 2 THEN 'Log Out'
				WHEN 3 THEN 'Forced Log Out'
				WHEN 8 THEN 'Insufficient Licence'
				ELSE 'Session Timeout'
			END);

END