
CREATE PROCEDURE sp_ASRGetFindOrderItems

(
	@lOrderID int,
	@lTableID int,
	@lViewID int,
	@bUseDef bit
)


AS

IF @bUseDef = 1

BEGIN

	SELECT ASRSysColumns.columnName, ASRSysColumns.DataType FROM ASRSysViewColumns 
	            INNER JOIN ASRSysColumns ON ASRSysViewColumns.columnID = ASRSysColumns.columnID 
	            WHERE ASRSysViewColumns.InView = 1 AND ASRSysViewColumns.ViewID = @lViewID
	            AND ASRSysColumns.datatype <> 4
END

ELSE

IF @lViewID > 0

BEGIN

	SELECT ASRSysOrderItems.*, ASRSysColumns.columnName, ASRSysColumns.DataType FROM 
             		 ASRSysOrderItems INNER JOIN 
	              ASRSysColumns ON ASRSysColumns.tableID = @lTableID  AND ASRSysColumns.columnID = 
	              ASRSysOrderItems.ColumnID INNER Join ASRSysViewColumns ON ASRSysColumns.ColumnID = 
	              ASRSysViewColumns.ColumnID WHERE ASRSysOrderItems.Type = 'F' AND 
	              ASRSysOrderItems.OrderID = @lOrderID  AND ASRSysViewColumns.ViewID = @lViewID
	              AND ASRSysViewColumns.InView = 1 ORDER BY ASRSysOrderItems.Type, ASRSysOrderItems.Sequence

END

ELSE

BEGIN

	SELECT ASRSysOrderItems.*, ASRSysColumns.ColumnName, ASRSysColumns.DataType 
             		FROM ASRSysOrderItems INNER JOIN ASRSysColumns
	             ON ASRSysColumns.TableID= @lTableID 
	             	AND ASRSysColumns.ColumnID=ASRSysOrderItems.ColumnID
	             WHERE Type='F' And OrderID= @lOrderID
             		ORDER BY Type, Sequence

END

GO

