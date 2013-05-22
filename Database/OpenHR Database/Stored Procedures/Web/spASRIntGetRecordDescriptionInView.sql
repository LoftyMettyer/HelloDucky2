CREATE PROCEDURE [dbo].[spASRIntGetRecordDescriptionInView] (
	@piViewID 			integer,
	@piTableID 			integer,
	@piRecordID			integer,
	@piParentTableID	integer,
	@piParentRecordID	integer,
	@psRecDesc			varchar(MAX)	OUTPUT,
	@psErrorMessage		varchar(MAX)	OUTPUT
)
AS
BEGIN
	/* Return the record descriptiuon for the given record, screen, view. */
	DECLARE	@iRecordID			integer,
			@iRecDescID			integer,
			@iCount				integer,
			@sEvalRecDesc		varchar(MAX),
			@sExecString		nvarchar(MAX),
			@sParamDefinition	nvarchar(500),
			@sViewName			sysname;

	SET @psRecDesc = '';
	SET @psErrorMessage = '';

	/* Return the parent record description if we have a parent record. */
	IF (@piParentTableID > 0) AND  (@piParentRecordID > 0)
	BEGIN
		SET @iRecordID = @piParentRecordID;

		/* Get the parent table's record description ID. */
		SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
		FROM ASRSysTables
		WHERE ASRSysTables.tableID = @piParentTableID;
	END
	ELSE
	BEGIN
		SET @iRecordID = @piRecordID;
 
		/* Get the table's record description ID. */
		SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
		FROM ASRSysTables 
		WHERE ASRSysTables.tableID = @piTableID;
	END

	IF @iRecordID > 0 
	BEGIN
		/* Check that the given record is still in the given view */
		SELECT @sViewName = viewName
		FROM ASRSysViews
		WHERE viewID = @piViewID;
	
		SET @sExecString = 'SELECT @iCount = COUNT(*) FROM [' + @sViewName + '] WHERE ID = ' + convert(nvarchar(100), @iRecordID);
		SET @sParamDefinition = N'@iCount integer OUTPUT';
		EXEC sp_executesql @sExecString, @sParamDefinition, @iCount OUTPUT;
		
		IF @iCount = 0 
		BEGIN
			SET @psErrorMessage = 'The requested record is not in the current view.';
		END
		ELSE
		BEGIN
			/* Get the record description. */
			IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iRecordID > 0)
			BEGIN
				SET @sExecString = 'exec sp_ASRExpr_' + convert(nvarchar(100), @iRecDescID) + ' @recDesc OUTPUT, @recID';
				SET @sParamDefinition = N'@recDesc varchar(MAX) OUTPUT, @recID integer';
				EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iRecordID;
		
				IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @psRecDesc = @sEvalRecDesc;
			END
		END
	END
END