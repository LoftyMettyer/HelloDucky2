CREATE PROCEDURE dbo.spASRDatabaseStatus (
	@message	nvarchar(MAX) OUTPUT)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @SystemStatus	integer,
			@lockUser		varchar(255),
			@lockhost		varchar(255)

	SET @message = '';

	SELECT TOP 1 @SystemStatus= ISNULL(Priority,0), @lockUser = [username], @lockhost = [hostname]
		FROM ASRSysLock
		ORDER BY [Priority];

	IF @SystemStatus = 1
	BEGIN
		SET @message = 'A database update has been started by user ' + @lockUser + ' on machine ' + @lockhost + CHAR(13) + CHAR(13) 
			+ 'Data may be lost until you log off and the update has completed.' + CHAR(13) + CHAR(13) 
			+ 'Please contact your HR system administrator.'
	END
	ELSE IF @SystemStatus = 2
	BEGIN
		SET @message = 'The user ' + @lockUser + ' has locked the database on machine ' + @lockhost + CHAR(13) + CHAR(13) 
			+ 'You are unable to make any changes at this time.' + CHAR(13) + CHAR(13) 
			+ 'Please contact your HR system administrator.'
	
	END


END
