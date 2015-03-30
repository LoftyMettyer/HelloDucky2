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
		SELECT DISTINCT o.name AS Name, o.OrderID
		FROM ASRSysOrders o
			INNER JOIN ASRSysOrderItems oi ON o.OrderID = oi.orderID
			INNER JOIN ASRSysViewColumns vc ON oi.columnID = vc.columnID
		WHERE o.tableID = @piTableID
			AND o.[type] = 1 AND vc.inView = 1 AND oi.[type] = 'O' AND vc.viewID = @piViewID
		ORDER BY o.name;
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