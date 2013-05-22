CREATE PROCEDURE [dbo].[sp_ASRIntGetRecordDescription] (
	@piTableID 			integer,
	@piRecordID			integer,
	@piParentTableID	integer,
	@piParentRecordID	integer,
	@psRecDesc			varchar(MAX)	OUTPUT
)
AS
BEGIN
	/* Return the record descriptiuon for the given record, screen, view. */
	DECLARE	@iRecordID			integer,
			@iRecDescID			integer,
			@sEvalRecDesc		varchar(8000),
			@sExecString		nvarchar(MAX),
			@sParamDefinition	nvarchar(500);

	SET @psRecDesc = '';

	/* Return the parent record description if we have a parent record. */
	IF (@piParentTableID > 0) AND  (@piParentRecordID > 0)
	BEGIN
		SET @iRecordID = @piParentRecordID;

		/* Get the parent table's record description ID. */
		SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
		FROM [dbo].[ASRSysTables]
		WHERE ASRSysTables.tableID = @piParentTableID;
	END
	ELSE
	BEGIN
		SET @iRecordID = @piRecordID;
 
		/* Get the table's record description ID. */
		SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
		FROM  [dbo].[ASRSysTables]
		WHERE ASRSysTables.tableID = @piTableID;
	END

	/* Get the record description. */
	IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iRecordID > 0)
	BEGIN
		SET @sExecString = 'exec sp_ASRExpr_' + convert(nvarchar(255), @iRecDescID) + ' @recDesc OUTPUT, @recID';
		SET @sParamDefinition = N'@recDesc varchar(MAX) OUTPUT, @recID integer';
		EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iRecordID;

		IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @psRecDesc = @sEvalRecDesc;
	END
END