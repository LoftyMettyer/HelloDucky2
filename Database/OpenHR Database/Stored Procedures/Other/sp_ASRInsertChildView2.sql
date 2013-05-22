
CREATE PROCEDURE sp_ASRInsertChildView2 (
	@plngNewRecordID	int OUTPUT,		/* Output variable to hold the new record ID. */
	@plngTableID		int,			/* ID of the table we're creating a view for. */
	@piType		integer,			/* 0 = OR inter-table join, 1 = AND inter-table join. */
	@psRole		varchar(256))		/* Role name. */
AS
BEGIN
	DECLARE @lngRecordID	int,
		@iCount		int

	SELECT @lngRecordID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @plngTableID
	AND role = @psRole

	IF @lngRecordID IS null
	BEGIN
		/* Insert a record in the ASRSysChildViews table. */
		INSERT INTO ASRSysChildViews2 (tableID, type, role)
		VALUES (@plngTableID, @piType, @psRole)

		/* Get the ID of the inserted record.*/
		SELECT @lngRecordID = MAX(childViewID) FROM ASRSysChildViews2
	END
	ELSE
	BEGIN
		UPDATE ASRSysChildViews2 
		SET type = @piType
		WHERE tableID = @plngTableID
		AND role = @psRole	
	END

	/* Return the new record ID. */
	SET @plngNewRecordID = @lngRecordID
END




GO

