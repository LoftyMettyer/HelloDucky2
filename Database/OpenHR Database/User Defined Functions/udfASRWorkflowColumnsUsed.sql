
CREATE FUNCTION [dbo].[udfASRWorkflowColumnsUsed] (
	@piWorkflowID		integer,
	@piElementID		integer,	-- >0 when the deleted record is the record deleted by the given StoredData element
	@pfDeleteTrigger	bit			-- 1 when the deleted record is the trigger record
)
RETURNS @results TABLE (
	columnID	integer
)
AS
BEGIN
	-- Return a table containing the info of any elements that refer to columns from the 
	-- DeleteTrigger record or the record identified using the given StoredData element.
	DECLARE
		@iBaseTableID		integer,
		@sIdentifier		varchar(8000),
		@iElementType		integer,
		@iEmailType			integer,
		@iEmailColumnID		integer,
		@iEmailExprID		integer,
		@iExprColumnID		integer

	-- Create a local table variable to hold the results.
	DECLARE @columnsUsed TABLE (
		columnID	integer
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
	-- 1) Email items
	----------------------------------------------------------------------------
	INSERT INTO @columnsUsed
	SELECT WEI.dbColumnID
	FROM ASRSysWorkflowElementItems WEI
	INNER JOIN ASRSysWorkflowElements WE ON WEI.elementID = WE.ID
	INNER JOIN ASRSysColumns Cols ON WEI.dbColumnID = Cols.columnID
	WHERE WE.workflowID = @piWorkflowID
		AND WE.type = 3 -- email
		AND WEI.itemType = 1 -- DBValue	
		AND Cols.tableID = @iBaseTableID
		AND (((@pfDeleteTrigger = 1) AND (WEI.dbRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (WEI.dbRecord = 1) -- Identified
				AND (WEI.recSelWebFormIdentifier = @sIdentifier)))

	----------------------------------------------------------------------------
	-- Determine which fields from the Deleted record are used in WebForm elements
	-- 1) WebForm DBValues
	----------------------------------------------------------------------------
	INSERT INTO @columnsUsed
	SELECT WEI.dbColumnID
	FROM ASRSysWorkflowElementItems WEI
	INNER JOIN ASRSysWorkflowElements WE ON WEI.elementID = WE.ID
	INNER JOIN ASRSysColumns Cols ON WEI.dbColumnID = Cols.columnID
	WHERE WE.workflowID = @piWorkflowID
		AND WE.type = 2 -- WebForm
		AND WEI.itemType = 1 -- DBValue	
		AND Cols.tableID = @iBaseTableID
		AND (((@pfDeleteTrigger = 1) AND (WEI.dbRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (WEI.dbRecord = 1) -- Identified
				AND (WEI.WFFormIdentifier = @sIdentifier)))

	----------------------------------------------------------------------------
	-- Determine which fields from the Deleted record are used in StoredData elements
	-- 1) StoredData DBValues
	----------------------------------------------------------------------------
	INSERT INTO @columnsUsed
	SELECT WEC.dbColumnID
	FROM ASRSysWorkflowElementColumns WEC
	INNER JOIN ASRSysWorkflowElements WE ON WEC.elementID = WE.ID
	INNER JOIN ASRSysColumns Cols ON WEC.dbColumnID = Cols.columnID
	WHERE WE.workflowID = @piWorkflowID
		AND WE.type = 5 -- StoredData
		AND WEC.valueType = 2 -- DBValue	
		AND Cols.tableID = @iBaseTableID
		AND (((@pfDeleteTrigger = 1) AND (WEC.dbRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (WEC.dbRecord = 1) -- Identified
				AND (WEC.WFFormIdentifier = @sIdentifier)))

	----------------------------------------------------------------------------
	-- Determine which fields from the Deleted record are used in Expressions
	----------------------------------------------------------------------------
	INSERT INTO @columnsUsed
	SELECT EC.fieldColumnID
	FROM ASRSysExprComponents EC
	INNER JOIN ASRSysExpressions EXPRS ON EC.exprID = EXPRS.exprID
	INNER JOIN ASRSysColumns Cols ON EC.fieldColumnID = Cols.columnID
	WHERE EXPRS.utilityID = @piWorkflowID
		AND EC.type = 12 -- WFField	
		AND Cols.tableID = @iBaseTableID
		AND (((@pfDeleteTrigger = 1) AND (EC.workflowRecord = 4)) -- Triggered
			OR ((@pfDeleteTrigger = 0) 
				AND (EC.workflowRecord = 1) -- Identified
				AND (EC.workflowElement = @sIdentifier)))

	-- Read and return the results from the local table variable.
	INSERT @results
		SELECT DISTINCT columnID
		FROM @columnsUsed

	RETURN
END




