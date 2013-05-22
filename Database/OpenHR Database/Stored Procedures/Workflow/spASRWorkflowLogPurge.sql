CREATE PROCEDURE [dbo].[spASRWorkflowLogPurge] 
AS
BEGIN
	EXEC sp_ASRPurgeRecords 'WORKFLOW', 'ASRSysWorkflowInstances', 'completionDateTime';
END

