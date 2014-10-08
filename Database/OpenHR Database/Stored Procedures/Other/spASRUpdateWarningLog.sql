CREATE PROCEDURE dbo.spASRUpdateWarningLog(
		@Username			varchar(255),
		@WarningType		integer,
		@WarningRefreshRate	integer,
		@WarnUser			bit OUTPUT)
	AS
	BEGIN

		DECLARE @Today				datetime = GETDATE(),
				@LastWarningDate	datetime;

		SELECT TOP 1 @LastWarningDate = DATEADD(dd, 0, DATEDIFF(dd, 0, WarningDate)) FROM ASRSysWarningsLog
			WHERE Username = @Username AND WarningType = @WarningType
			ORDER BY WarningDate DESC;

		SET @WarnUser = 0;
		IF @LastWarningDate IS NULL OR DATEDIFF(day, @LastWarningDate, DATEDIFF(dd, 0, @Today)) >= @WarningRefreshRate SET @WarnUser = 1

		IF @WarnUser = 1
			INSERT ASRSysWarningsLog (UserName, WarningType, WarningDate) VALUES (@UserName, @WarningType, @Today);

		RETURN @WarnUser;
	END