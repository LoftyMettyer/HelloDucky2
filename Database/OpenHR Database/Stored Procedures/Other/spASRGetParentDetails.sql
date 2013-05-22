CREATE PROCEDURE [dbo].[spASRGetParentDetails]
(
	@piBaseTableID		integer,
	@piBaseRecordID		integer,
	@piParent1TableID	integer	OUTPUT,
	@piParent1RecordID	integer	OUTPUT,
	@piParent2TableID	integer	OUTPUT,
	@piParent2RecordID	integer	OUTPUT
)
AS
BEGIN
	-- Return the parent table IDs and related record IDs for the given table and record.
	-- Return 0 if no table/record exists.
	DECLARE
		@sSQL		nvarchar(MAX),
		@sParam		nvarchar(500),
		@sTableName	nvarchar(255);

	SET @piParent1TableID = 0;
	SET @piParent1RecordID = 0;
	SET @piParent2TableID = 0;
	SET @piParent2RecordID = 0;

	SELECT @sTableName = tableName
	FROM [dbo].[ASRSysTables]
	WHERE tableID = @piBaseTableID;

	SELECT TOP 1 @piParent1TableID = isnull(parentID, 0)
	FROM [dbo].[ASRSysRelations]
	WHERE childID = @piBaseTableID
	ORDER BY parentID ASC;

	SELECT TOP 1 @piParent2TableID = isnull(parentID, 0)
	FROM [dbo].[ASRSysRelations] 
	WHERE childID = @piBaseTableID
		AND parentID <> @piParent1TableID
	ORDER BY parentID ASC;

	IF (LEN(@sTableName) > 0) AND (@piBaseRecordID > 0)
	BEGIN
		IF (@piParent1TableID > 0)
		BEGIN
			SET @sSQL = 'SELECT @piParent1RecordID = isnull(ID_' + convert(nvarchar(100), @piParent1TableID) + ',0)'
				+ ' FROM ' + @sTableName
				+ ' WHERE ID = ' + convert(varchar(100), @piBaseRecordID);
			SET @sParam = N'@piParent1RecordID integer OUTPUT';
			EXEC sp_executesql @sSQL, @sParam, @piParent1RecordID OUTPUT;
		END

		IF @piParent2TableID > 0 
		BEGIN
			SET @sSQL = 'SELECT @piParent2RecordID = isnull(ID_' + convert(nvarchar(100), @piParent2TableID) + ',0)'
				+ ' FROM ' + @sTableName
				+ ' WHERE ID = ' + convert(varchar(100), @piBaseRecordID);
			SET @sParam = N'@piParent2RecordID integer OUTPUT';
			EXEC sp_executesql @sSQL, @sParam, @piParent2RecordID OUTPUT;
		END		
	END
END