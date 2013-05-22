CREATE PROCEDURE [dbo].[spASRIntGetDefaultOrder] (
	@piTableID	integer,
	@piOrderID	integer	OUTPUT
)
AS
BEGIN
	SELECT @piOrderID = defaultOrderID
	FROM ASRSysTables
	WHERE tableID = @piTableID;
END