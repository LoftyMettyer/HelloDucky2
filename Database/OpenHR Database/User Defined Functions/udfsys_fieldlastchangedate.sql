CREATE FUNCTION [dbo].[udfsys_fieldlastchangedate](
		@colrefID	varchar(32),
		@recordID	integer
	)
	RETURNS datetime
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @result		datetime,
				@tableid	integer,
				@columnid	integer;
		
		SET @tableid = SUBSTRING(@colrefID, 1, 8);
		SET @columnid = SUBSTRING(@colrefID, 10, 8);

		SELECT TOP 1 @result = [DateTimeStamp] FROM dbo.[ASRSysAuditTrail]
			WHERE [ColumnID] = @columnid AND [TableID] = @tableID
				AND @recordID = [RecordID]
			ORDER BY [DateTimeStamp] DESC ;

		RETURN @result;

	END