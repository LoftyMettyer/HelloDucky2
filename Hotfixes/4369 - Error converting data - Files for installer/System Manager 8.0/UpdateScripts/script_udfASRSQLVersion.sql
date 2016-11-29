/*
Hotfix Number:1000
Description     :Handles Microsoft changing the way version numbers are returned
Run Type     :2
Version            :8.2
Run Once     :No
Sequence     :2
Database Guid   :None
Checksum :     0x69FD
*/
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRSQLVersion]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRSQLVersion]

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfASRSQLVersion]()
	RETURNS integer
	AS
	BEGIN
		RETURN convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY(''ProductVersion'')))
	END'





