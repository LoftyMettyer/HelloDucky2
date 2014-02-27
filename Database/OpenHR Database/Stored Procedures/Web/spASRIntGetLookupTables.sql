CREATE PROCEDURE [dbo].[spASRIntGetLookupTables]
AS
BEGIN
	/* return a recordset of all the lookup tables. */
	SELECT ASRSysTables.TableName, ASRSysTables.TableID 
    FROM ASRSysTables
    WHERE ASRSysTables.TableType = 3
    ORDER BY ASRSysTables.TableName;
END	
