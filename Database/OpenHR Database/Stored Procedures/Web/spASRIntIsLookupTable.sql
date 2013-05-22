CREATE PROCEDURE spASRIntIsLookupTable (
	@piTableID		integer,
	@pfIsLookupTable	bit OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;
	
	SELECT @pfIsLookupTable = 
		CASE
			WHEN tableType = 3 THEN 1
			ELSE 0
		END
	FROM ASRSysTables
	WHERE tableID = @piTableID

	IF @pfIsLookupTable IS null SET @pfIsLookupTable = 0
END
GO

