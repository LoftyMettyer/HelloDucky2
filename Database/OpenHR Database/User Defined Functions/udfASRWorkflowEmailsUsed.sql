
CREATE FUNCTION [dbo].[udfASRWorkflowEmailsUsed] (
	@piWorkflowID		integer,
	@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
	@pfDeleteTrigger	bit			-- 1 when the deleted record is the trigger record
)
RETURNS @results TABLE (
		emailID		integer,
		type		integer,
		colExprID	integer
)
AS
BEGIN
	-- Return a table containing the info of any email elements that refer to the 
	-- DeleteTrigger record or the record identified using the given StoredData element.
	DECLARE
		@iBaseTableID		integer,
		@sIdentifier		varchar(8000)

	-- Create a local table variable to hold the results.
	DECLARE @emailsUsed TABLE (
		emailID		integer,
		type		integer,
		colExprID	integer
	)

	-- Get the basic info of the given Workflow/element
	IF @pfDeleteTrigger = 1
	BEGIN
		-- Get the table of the deleted record.
		SELECT @sIdentifier = '',
			@iBaseTableID = isnull(WF.baseTable, 0)
		FROM ASRSysWorkflows WF 
		WHERE WF.ID = @piWorkflowID
	END
	ELSE
	BEGIN
		-- Get the table of the deleted record
		-- and the identifier of the StoredData element.
		SELECT @sIdentifier = isnull(WFE.identifier, ''),
			@iBaseTableID = isnull(WFE.dataTableID, 0)
		FROM ASRSysWorkflowElements WFE
		WHERE WFE.ID = @piElementID
			AND WFE.type = 5 -- StoredData
			AND WFE.dataAction = 2 -- Delete
	END

	----------------------------------------------------------------------------
	-- Determine which fields from the Deleted record are used in Email elements
	-- 1) Email To address
	----------------------------------------------------------------------------
	INSERT INTO @emailsUsed
	SELECT WFE.emailID,
		EA.type,
		CASE
			WHEN EA.type = 1 THEN EA.columnID -- Column
			ELSE EA.exprID -- Calculated
		END
	FROM ASRSysWorkflowElements WFE
	INNER JOIN ASRSysEmailAddress EA ON WFE.emailID = EA.emailID
	WHERE WFE.workflowID = @piWorkflowID
		AND WFE.type = 3 -- Email
		AND EA.tableID = @iBaseTableID
		AND ((EA.type = 1) OR (EA.type = 2))
		AND (((@pfDeleteTrigger = 1) AND (WFE.emailRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (WFE.emailRecord = 1) -- Identified
				AND (WFE.recSelWebFormIdentifier = @sIdentifier)))

	----------------------------------------------------------------------------
	-- 2) Email Copy address
	----------------------------------------------------------------------------
	INSERT INTO @emailsUsed
	SELECT WFE.emailCCID,
		EA.type,
		CASE
			WHEN EA.type = 1 THEN EA.columnID -- Column
			ELSE EA.exprID -- Calculated
		END
	FROM ASRSysWorkflowElements WFE
	INNER JOIN ASRSysEmailAddress EA ON WFE.emailCCID = EA.emailID
	WHERE WFE.workflowID = @piWorkflowID
		AND WFE.type = 3 -- Email
		AND EA.tableID = @iBaseTableID
		AND ((EA.type = 1) OR (EA.type = 2))
		AND (((@pfDeleteTrigger = 1) AND (WFE.emailRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (WFE.emailRecord = 1) -- Identified
				AND (WFE.recSelWebFormIdentifier = @sIdentifier)))

	-- Read and return the results from the local table variable.
	INSERT @results
		SELECT DISTINCT emailID,
			type,
			colExprID
		FROM @emailsUsed

	RETURN
END




