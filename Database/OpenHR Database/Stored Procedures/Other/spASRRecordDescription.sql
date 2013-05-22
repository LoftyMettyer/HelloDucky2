CREATE PROCEDURE [dbo].[spASRRecordDescription]
(
	@piTableID				integer,
	@piRecordID				integer,
	@psRecordDescription	varchar(MAX)	OUTPUT
)
 AS
BEGIN
	DECLARE @sSQL varchar(MAX),
		@iRecordDescID integer,
		@sRecordDesc varchar(MAX);

	SET @psRecordDescription = '';

	SELECT @iRecordDescID = ISNULL(ASRSysTables.recordDescExprID, 0)
		FROM ASRSysTables
		WHERE ASRSysTables.tableID = @piTableID;

	IF @iRecordDescID > 0 
	BEGIN
		SET @sSQL = 'sp_ASRExpr_' + convert(varchar,@iRecordDescID);
		IF EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @sRecordDesc OUTPUT, @piRecordID;
			SET @psRecordDescription = @sRecordDesc;
		END
	END
END