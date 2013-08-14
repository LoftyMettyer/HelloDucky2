CREATE PROCEDURE [dbo].[sp_ASRGetOrders] (
	@piViewID	int,
	@piTableID	int)
AS
BEGIN
	SELECT DISTINCT ASRSysOrders.OrderID, 
		ASRSysOrders.name, 
		ASRSysOrders.tableID 
	FROM ASRSysOrders 
	INNER JOIN ASRSysOrderItems ON ASRSysOrders.OrderID = ASRSysOrderItems.orderID 
	INNER JOIN ASRSysViewColumns ON ASRSysOrderItems.columnID = ASRSysViewColumns.columnID 
	WHERE ASRSysOrders.tableID = @piTableID  
		AND ASRSysViewColumns.viewID = @piViewID  
		AND ASRSysViewColumns.inView = 1;
END