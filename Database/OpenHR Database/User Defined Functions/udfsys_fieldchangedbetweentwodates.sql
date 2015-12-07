CREATE FUNCTION [dbo].[udfsys_fieldchangedbetweentwodates](
		@colrefID	varchar(32),
		@fromdate	datetime,
		@todate		datetime,
		@recordID	integer
	)
	RETURNS bit
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @result		bit,
				@tableid	integer,
				@columnid	integer;
		
		SET @tableid = SUBSTRING(@colrefID, 1, 8);
		SET @columnid = SUBSTRING(@colrefID, 10, 8);
		SET @fromdate = DATEADD(dd, 0, DATEDIFF(dd, 0, @fromdate));
		SET @todate = DATEADD(dd, 0, DATEDIFF(dd, 0, @todate));

		SELECT @result = CASE WHEN
				EXISTS(SELECT [DateTimeStamp] FROM dbo.[ASRSysAuditTrail]
					WHERE [ColumnID] = @columnid AND [TableID] = @tableID
					AND @recordID = [RecordID] 
					AND [DateTimeStamp] >= @fromdate AND DateTimeStamp < @todate + 1)
				THEN 1 ELSE 0 END;

		RETURN @result;

	END