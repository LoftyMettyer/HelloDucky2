
CREATE FUNCTION [dbo].[udfASRGetAllPrecedingWorkflowElements] (
	@piElementID	integer
)
RETURNS @results TABLE (id integer)
AS
BEGIN
	-- Return a table containing the IDs of ALL elements that precede the given element.
	-- Connectors are ignored.

	DECLARE @iRowsAdded integer

	-- Create a local table variable to hold the results.
	DECLARE @precedingElements TABLE (
		elementID	integer PRIMARY KEY CLUSTERED,
		type		integer,
		processed	tinyint default 0
	)

	-- Add details of preceding elements into the results table.
	-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
	INSERT INTO @precedingElements
	SELECT DISTINCT
		CASE 
			WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
			ELSE E.ID
		END, 
		CASE 
			WHEN E.type = 9 THEN 8 -- 8 = Connector 1
			ELSE E.type
		END, 
		0
	FROM ASRSysWorkflowLinks L
	INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
	WHERE L.endElementID = @piElementID

	SET @iRowsAdded = @@rowcount

	WHILE @iRowsAdded > 0
	BEGIN
		-- If we've just added rows to the results table, process the new rows.
		-- Mark the new rows as 'being processed'.
		UPDATE @precedingElements
		SET processed = 1
		WHERE processed = 0

		-- Add details of elements that precede those being processed into the results table.
		-- NB. We skip over the Connector2 elements straight onto the associated Connector1 elements.
		INSERT INTO @precedingElements
		SELECT DISTINCT
			CASE 
				WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
				ELSE E.ID
			END, 
			CASE 
				WHEN E.type = 9 THEN 8 -- 8 = Connector 1
				ELSE E.type
			END, 
			0
		FROM ASRSysWorkflowLinks L
		INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
		INNER JOIN @precedingElements precEl ON L.endElementID = precEl.elementID
		WHERE precEl.processed = 1
			AND CASE 
				WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
				ELSE E.ID
			END NOT IN (SELECT elementID FROM @precedingElements)
			AND CASE 
				WHEN E.type = 9 THEN E.connectionPairID -- 9 = Connector 2
				ELSE E.ID
			END <> @piElementID

		SET @iRowsAdded = @@rowcount

		-- Mark the processed rows as 'been processed'.
		UPDATE @precedingElements
		SET processed = 2
		WHERE processed = 1
	END

	-- Read and return the results from the local table variable.
	INSERT @results
		SELECT elementID
		FROM @precedingElements
		WHERE type <> 8
			AND type <> 9
	RETURN
END




