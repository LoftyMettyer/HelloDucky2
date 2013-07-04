CREATE Procedure spASRIntGetEventLogPurgeDetails
AS
BEGIN
	SET NOCOUNT ON;

	SELECT * FROM ASRSysEventLogPurge;
END
GO

