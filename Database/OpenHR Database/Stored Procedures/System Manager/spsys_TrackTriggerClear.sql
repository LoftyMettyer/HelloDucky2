CREATE PROCEDURE dbo.spsys_TrackTriggerClear(@TableFromID integer)
AS
BEGIN

	DECLARE @buffer varchar(128) = '',
			  @varBin varbinary(128);

    SELECT @buffer += convert(varchar(125),[TableFromId]) + CHAR(2) + convert(varchar(3),[NestLevel]) + CHAR(2) + convert(varchar(3),[ActionType]) + CHAR(3)
		  FROM [InTriggerContext]
		  WHERE [TableFromId] <> @TableFromID

	SET @varBin = convert(varbinary(128), @buffer);
   SET CONTEXT_INFO @varBin;

END