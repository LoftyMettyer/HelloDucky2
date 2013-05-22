CREATE PROCEDURE [dbo].[sp_ASRFn_AuditFieldLastChangeDate]
(
	@Result		datetime OUTPUT,
	@ColumnID	integer,
	@RecordID	integer
)
AS
BEGIN
	SET @Result = (SELECT TOP 1 DateTimeStamp FROM [dbo].[ASRSysAuditTrail]
			WHERE ColumnID = @ColumnID And @RecordID = RecordID
			ORDER BY DateTimeStamp DESC);
END