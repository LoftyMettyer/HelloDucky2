CREATE FUNCTION [dbo].[udfsys_username]
	(@userid as integer)
RETURNS varchar(255)
WITH SCHEMABINDING
AS
BEGIN

	RETURN SYSTEM_USER;

END