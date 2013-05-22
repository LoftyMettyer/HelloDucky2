






CREATE PROCEDURE sp_ASRIntGetOrders (@plngTableID int, @plngViewID int)
AS
BEGIN
	/* Return a recordset of the IDs and names of the orders available for the given table/view. */
	DECLARE @lngTableID		int,
		@lngDefaultOrderID	int

	/* Get the table ID from the view ID (if required). */
	IF @plngTableID > 0 
	BEGIN
		SET @lngTableID = @plngTableID
	END
	ELSE
	BEGIN
		SELECT @lngTableID = ASRSysViews.viewTableID
		FROM ASRSysViews
		WHERE ASRSysViews.viewID = @plngViewID
	END

	/* Create a temporary table to hold our resultset. */
	CREATE TABLE #orderInfo
	(
		orderID			int,
		orderName		sysname,
		defaultOrder		bit
	)

	/* Populate the temporary table with information on the order for the given table. */
   	 INSERT INTO #orderInfo (
		orderID, 
		orderName,
		defaultOrder)	
	(SELECT ASRSysOrders.orderID, 
		ASRSysOrders.name,
		0
	FROM ASRSysOrders
	WHERE ASRSysOrders.tableID = @lngTableID)

	/* Get the table's default order. */
	SELECT @lngDefaultOrderID = ASRSysTables.defaultOrderID
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @lngTableID

	IF @lngDefaultOrderID > 0 
	BEGIN
		UPDATE #orderInfo
		SET defaultOrder = 1 
		WHERE orderID = @lngDefaultOrderID
	END

	/* Return the resultset. */
	SELECT *
	FROM #orderInfo 
	ORDER BY orderName
END







GO

