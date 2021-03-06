CREATE FUNCTION [dbo].[udfASRGetSucceedingWorkflowElements] (
	@piElementID	integer,
	@piFlowCode		integer
)
RETURNS @results TABLE (id integer)
AS
BEGIN
	-- Return a table containing the IDs of the elements that succeed the given element
	-- from the given flow.
	-- Connectors are ignored.

	DECLARE @iRowsAdded integer

	-- Create a local table variable to hold the results.
	DECLARE @succeedingElements TABLE (
		elementID	integer PRIMARY KEY CLUSTERED,
		type		integer,
		processed	tinyint default 0
	)

	-- Add details of succeeding elements into the results table.
	-- NB. We skip over the Connector1 elements straight onto the associated Connector 2 elements.
	INSERT INTO @succeedingElements
	SELECT DISTINCT
		CASE 
			WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
			ELSE E.ID
		END, 
		CASE 
			WHEN E.type = 8 THEN 9 -- 9 = Connector 2
			ELSE E.type
		END, 
		0
	FROM ASRSysWorkflowLinks L
	INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
	WHERE L.startElementID = @piElementID
		AND ((L.startOutboundFlowCode = @piFlowCode) OR 
			(@piFlowCode = 0 and L.startOutboundFlowCode = -1))

	SET @iRowsAdded = @@rowcount

	WHILE @iRowsAdded > 0
	BEGIN
		-- If we've just added rows to the results table, process the new rows.
		-- Mark the new rows as 'being processed'.
		UPDATE @succeedingElements
		SET processed = 1
		WHERE processed = 0

		-- Add details of elements that succeed those being processed into the results table.
		-- NB. We skip over the Connector1 elements straight onto the associated Connector 2 elements.
		INSERT INTO @succeedingElements
		SELECT DISTINCT
			CASE 
				WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
				ELSE E.ID
			END, 
			CASE 
				WHEN E.type = 8 THEN 9 -- 9 = Connector 2
				ELSE E.type
			END, 
			0
		FROM ASRSysWorkflowLinks L
		INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
		INNER JOIN @succeedingElements succEl ON L.startElementID = succEl.elementID
		WHERE succEl.processed = 1
			AND succEl.type = 9 -- 9 = Connector 2
			AND CASE 
				WHEN E.type = 8 THEN E.connectionPairID -- 8 = Connector 1
				ELSE E.ID
			END NOT IN (SELECT elementID FROM @succeedingElements)

		SET @iRowsAdded = @@rowcount

		-- Mark the processed rows as 'been processed'.
		UPDATE @succeedingElements
		SET processed = 2
		WHERE processed = 1
	END

	-- Read and return the results from the local table variable.
	INSERT @results
		SELECT elementID
		FROM @succeedingElements
		WHERE type <> 8
			AND type <> 9
	RETURN
END
