CREATE PROCEDURE spASRIntGetSingleRecordViewID 
(
		@piTableID		integer OUTPUT,
		@piViewID		integer	OUTPUT
)
AS
BEGIN
	SELECT @piTableID = TableID, @piViewID = ViewID
	FROM ASRSysSSIViews
	WHERE SingleRecordView = 1

	IF @piTableID IS NULL SET @piTableID = 0
	IF @piViewID IS NULL SET @piViewID = 0
END
GO

