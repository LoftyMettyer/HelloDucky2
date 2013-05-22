CREATE PROCEDURE [dbo].[spASRIntClearEventLogPurge]
AS
BEGIN

	SET NOCOUNT ON;

	DELETE FROM [dbo].[ASRSysEventLogPurge];
END