CREATE PROCEDURE dbo.spASRGetActiveWorkflowStoredDataSteps
AS
BEGIN
	/* Return a recordset of the workflow StoredData steps that need to be actioned by the Workflow service. */
	DECLARE @steps table(ID integer)

	INSERT INTO @steps
	SELECT S.ID
	FROM ASRSysWorkflowInstanceSteps S
	INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
	WHERE S.status = 1
		AND E.type = 5 -- 5 = Stored Data

	UPDATE ASRSysWorkflowInstanceSteps
	SET status = 5 -- In progress
	WHERE ID IN (SELECT ID FROM @steps)

	SELECT S.instanceID AS [instanceID],
		E.ID AS [elementID],
		S.ID AS [stepID]
	FROM ASRSysWorkflowInstanceSteps S
	INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
	WHERE s.ID IN (SELECT ID FROM @steps)
END


