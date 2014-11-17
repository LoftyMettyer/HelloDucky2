CREATE FUNCTION [dbo].[udfsys_isovernightprocess] ()
RETURNS bit 
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result bit = 0;
	SELECT @result = ISNULL(settingValue,0) FROM dbo.[ASRSysSystemSettings] WHERE section = 'database' AND settingKey = 'updatingdatedependantcolumns';
	
  RETURN @result;

END
