CREATE PROCEDURE dbo.sp_ASRGetOrderDefinition (
		@piOrderID int) 
AS
BEGIN

		SET NOCOUNT ON;

		-- Return the recordset of order items for the given order.
		SELECT oi.*, c.columnName, c.tableID,	c.dataType, t.tableName,
				c.Size,	c.Decimals,	c.Use1000Separator, c.blankIfZero
		FROM ASRSysOrderItems oi
			INNER JOIN ASRSysColumns c ON oi.columnID = c.columnID
			INNER JOIN ASRSysTables t ON t.tableID = c.tableID
		WHERE oi.orderID = @piOrderID
			AND c.dataType <> -4 AND c.datatype <> -3
		ORDER BY oi.type, oi.sequence;

END
