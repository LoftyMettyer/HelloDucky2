CREATE FUNCTION [dbo].[udfASRConvertNumeric]
(
	@in  decimal(15,4)
  , @dec integer
  , @sep integer
)
RETURNS varchar(MAX)
AS
BEGIN

  DECLARE @out varchar(max)
  DECLARE @out2 varchar(max)

  SET @out = convert(varchar(max),cast(@in as money),@sep)
  SET @out2 = ''

  IF @dec <> 2
  BEGIN
    SET @out = substring(@out,1,CHARINDEX('.',@out)-1)
    IF @dec = 1
      SET @out2 = convert(varchar(max),cast(@in as decimal(15,1)))
    ELSE IF @dec = 3
      SET @out2 = convert(varchar(max),cast(@in as decimal(15,3)))
    ELSE IF @dec = 4
      SET @out2 = convert(varchar(max),cast(@in as decimal(15,4)))
    END

    RETURN @out+substring(@out2,CHARINDEX('.',@out2),8000)

END