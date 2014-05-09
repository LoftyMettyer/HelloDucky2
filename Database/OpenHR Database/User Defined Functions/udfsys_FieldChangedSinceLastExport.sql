CREATE FUNCTION [dbo].[udfsys_FieldChangedSinceLastExport](
		@columnID	integer,
		@FromDate	datetime,
		@recordID	integer
	)
	RETURNS bit
	AS
	BEGIN

		DECLARE @result	bit = 0;
		
		SELECT @result = CASE WHEN
				EXISTS(SELECT [DateTimeStamp] FROM dbo.[ASRSysAuditTrail]
					WHERE [ColumnID] = @columnid
					AND @recordID = [RecordID] 
					AND [DateTimeStamp] >= @FromDate)
				THEN 1 ELSE 0 END;

		RETURN @result;
	END