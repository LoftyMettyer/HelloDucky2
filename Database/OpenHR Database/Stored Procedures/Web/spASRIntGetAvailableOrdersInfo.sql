CREATE PROCEDURE [dbo].[spASRIntGetAvailableOrdersInfo] (
	@plngTableID		integer
)
AS
BEGIN
	SELECT orderid AS [ID], 
		name, 
		'' AS username, 
		'' AS access 
	FROM ASRSysOrders 
	WHERE tableid = @plngTableID  
		AND type = 1 
		ORDER BY [name];
END