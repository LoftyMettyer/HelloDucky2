CREATE PROCEDURE sp_ASRGetOrderDefinition (
			@piOrderID int) 
		AS
		BEGIN
			/* Return the recordset of order items for the given order. */
			SELECT ASRSysOrderItems.*,
				ASRSysColumns.columnName,
				ASRSysColumns.tableID,
				ASRSysColumns.dataType,
			    	ASRSysTables.tableName,
					ASRSysColumns.Size,
					ASRSysColumns.Decimals,
					ASRSysColumns.Use1000Separator, 
					ASRSysColumns.blankIfZero
			FROM ASRSysOrderItems
			INNER JOIN ASRSysColumns 
				ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
			INNER JOIN ASRSysTables 
				ON ASRSysTables.tableID = ASRSysColumns.tableID
			WHERE ASRSysOrderItems.orderID = @piOrderID
			AND ASRSysColumns.dataType <> -4
			AND ASRSysColumns.datatype <> -3
			ORDER BY ASRSysOrderItems.type, 
				ASRSysOrderItems.sequence
		END





GO