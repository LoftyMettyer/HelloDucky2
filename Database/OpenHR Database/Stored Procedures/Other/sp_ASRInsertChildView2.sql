CREATE PROCEDURE sp_ASRInsertChildView2 (
	@plngNewRecordID	int OUTPUT,		/* Output variable to hold the new record ID. */
	@plngTableID		int,			/* ID of the table we are creating a view for. */
	@piType		integer,			/* 0 = OR inter-table join, 1 = AND inter-table join. */
	@psRole		varchar(256))		/* Role name. */
AS
BEGIN
	DECLARE @lngRecordID	int,
			@iCount		int;

	DECLARE	@outputTable table (childViewId int NOT NULL);

	SELECT @lngRecordID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @plngTableID
	AND role = @psRole;

	IF @lngRecordID IS NULL
	BEGIN
		/* Insert a record in the ASRSysChildViews table. */
		INSERT INTO ASRSysChildViews2 (tableID, type, role)
		OUTPUT inserted.childViewID INTO @outputTable
		VALUES (@plngTableID, @piType, @psRole);

		/* Get the ID of the inserted record.*/
		SELECT @lngRecordID = childViewId FROM @outputTable;
	END
	ELSE
	BEGIN
		UPDATE ASRSysChildViews2 
		SET type = @piType
		WHERE tableID = @plngTableID
		AND role = @psRole;
	END

	/* Return the new record ID. */
	SET @plngNewRecordID = @lngRecordID;
END