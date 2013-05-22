CREATE PROCEDURE [dbo].[sp_ASRIntGetTableOrders] (
	@piTableID 		integer, 
	@piViewID 		integer)
AS
BEGIN

	SET NOCOUNT ON;
	
	/* Return a recordset of the orders for the current table/view and order IDs.
		@piTableID = the ID of the table on which the order is based.
		@piViewID = the ID of the view on which the order is based.
	*/

	IF @piViewID > 0 
	BEGIN
		SELECT DISTINCT ASRSysOrders.name AS Name, 
			ASRSysOrders.orderID
		FROM ASRSysOrders
		INNER JOIN ASRSysOrderItems ON ASRSysOrders.orderID = ASRSysOrderItems.orderID
		INNER JOIN ASRSysViewColumns ON ASRSysOrderItems.columnID = ASRSysViewColumns.columnID
		WHERE ASRSysOrders.tableID = @piTableID
			AND ASRSysOrders.[type] = 1
			AND ASRSysViewColumns.inView = 1
			AND ASRSysOrderItems.[type] = 'O'
			AND ASRSysViewColumns.viewID = @piViewID
		ORDER BY ASRSysOrders.name;
	END
	ELSE
	BEGIN
		SELECT name AS Name, orderID
		FROM ASRSysOrders
		WHERE tableID= @piTableID
			AND ASRSysOrders.[type] = 1
		ORDER BY name;
	END
END