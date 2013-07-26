CREATE FUNCTION [dbo].[udfsys_getownerid]()
RETURNS uniqueidentifier
AS
BEGIN

	DECLARE @returnval uniqueidentifier;
	SELECT @returnval = [SettingValue]
		FROM dbo.[ASRSysSystemSettings]
		WHERE [Section] = 'database' AND [SettingKey] = 'ownerid';
	RETURN @returnval;

END