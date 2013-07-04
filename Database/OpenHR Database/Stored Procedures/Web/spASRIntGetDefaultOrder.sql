CREATE PROCEDURE [dbo].[spASRIntGetDefaultOrder] (
	@piTableID	integer,
	@piOrderID	integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT @piOrderID = defaultOrderID
	FROM ASRSysTables
	WHERE tableID = @piTableID;
END