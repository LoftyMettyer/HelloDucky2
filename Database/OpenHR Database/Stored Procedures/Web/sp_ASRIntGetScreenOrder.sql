CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenOrder] (
	@plngOrderID 	int OUTPUT, 
	@plngScreenID	int)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return the order ID of the given screen in the @plngOrderID parameter. */
	DECLARE @lngDefaultOrderID	integer;

	/* Get the order ID, and associated tbale id of the given screen. */
	SELECT @plngOrderID = ASRSysScreens.orderID,
		@lngDefaultOrderID = ASRSysTables.defaultOrderID
	FROM ASRSysScreens
	INNER JOIN ASRSysTables 
		ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.screenID = @plngScreenID;

	/* If no order is defined then use the associated table's default order. */
	IF (@plngOrderID IS NULL) OR (@plngOrderID <= 0)
	BEGIN
		SET @plngOrderID = @lngDefaultOrderID;
	END
END