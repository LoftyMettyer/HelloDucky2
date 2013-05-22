CREATE PROCEDURE sp_ASRGetQuickEntry

(
	@lScreenID int,
	@lTableID int
)

 AS

SELECT ASRSysColumns.columnName, ASRSysTables.TableName, 
    ASRSysTables.TableID

FROM ASRSysTables INNER JOIN
    ASRSysColumns ON 
    ASRSysTables.TableID = ASRSysColumns.tableID INNER JOIN
    ASRSysControls ON 
    ASRSysColumns.columnID = ASRSysControls.ColumnID

WHERE ASRSysControls.ScreenID = @lScreenID AND 
    ASRSysControls.TableID <> @lTableID

ORDER BY ASRSysTables.TableName

GO

