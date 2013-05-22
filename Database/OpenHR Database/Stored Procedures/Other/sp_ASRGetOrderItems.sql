

/****** Object:  Stored Procedure dbo.sp_ASRGetOrderItems    Script Date: 16/03/99 09:16:14 ******/
CREATE PROCEDURE sp_ASRGetOrderItems

(
	@sType varchar(1),
	@lOrderID int,
	@lTableID int,
	@lViewID int
) 

AS

IF @lViewID = 0

	SELECT ASRSysOrderItems.*, 
		ASRSysColumns.columnName,
		ASRSysColumns.datatype,
    		ASRSysOrderItems.ColumnID

	FROM	ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.ColumnID = ASRSysColumns.columnID

	WHERE	(ASRSysColumns.tableID = @lTableID) AND 
    		(ASRSysOrderItems.Type = @sType) AND 
	    	(ASRSysOrderItems.OrderID = @lOrderID)

	ORDER BY ASRSysOrderItems.Type, ASRSysOrderItems.Sequence

ELSE

	SELECT	ASRSysOrderItems.*, ASRSysColumns.columnName, 
	    	ASRSysColumns.datatype, ASRSysOrderItems.ColumnID

	FROM	ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.ColumnID = ASRSysColumns.columnID 
	INNER JOIN ASRSysViewColumns ON ASRSysColumns.columnID = ASRSysViewColumns.ColumnID

	WHERE 	(ASRSysColumns.tableID = @lTableID) AND 
		(ASRSysOrderItems.Type = @sType) AND 
		(ASRSysOrderItems.OrderID = @lOrderID) AND 
    		(ASRSysViewColumns.InView = 1) AND 
		(ASRSysViewColumns.ViewID = @lViewID)

	ORDER BY ASRSysOrderItems.Type, ASRSysOrderItems.Sequence



GO

