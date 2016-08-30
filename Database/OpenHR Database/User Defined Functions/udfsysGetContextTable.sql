CREATE FUNCTION [dbo].udfsysGetContextTable()
  RETURNS @Context TABLE([TableFromId] integer, [NestLevel] tinyint, [ActionType] tinyint)
  WITH SCHEMABINDING
AS
BEGIN

  DECLARE @buffer varchar(128) = rtrim(replace(convert(varchar(128),CONTEXT_INFO()), char(0), char(32)));
  DECLARE @fPtr1 int = CHARINDEX(CHAR(2),@buffer),
		    @rPtr int = CHARINDEX(CHAR(3),@buffer);
  DECLARE @fPtr2 int = CHARINDEX(CHAR(2),@buffer, @fPtr1+1);
		  
  WHILE @rPtr > 0
  BEGIN

    INSERT INTO @Context
	    SELECT convert(integer, SUBSTRING(@buffer,1,abs(@fPtr1-1))),
			convert(tinyint, SUBSTRING(@buffer, @fPtr1+1, @fPtr2-@fPtr1-1)), 
			convert(tinyint, SUBSTRING(@buffer, @fPtr2+1, @rPtr-@fPtr2-1))
	    WHERE @rPtr > NULLIF(@fPtr1,0)+1;

	SET @buffer = SUBSTRING(@buffer,@rPtr+1,128);
	SET @fPtr1 = CHARINDEX(CHAR(2),@buffer);
	SET @fPtr2 = CHARINDEX(CHAR(2),@buffer, @fPtr1+1);
	SET @rPtr = CHARINDEX(CHAR(3),@buffer);

  END

  RETURN;

END