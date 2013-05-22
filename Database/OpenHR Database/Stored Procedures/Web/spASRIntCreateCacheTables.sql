
-- Not called as the local temporary table would only be visible for the duration of this stored procedure,
-- now the command text is called directly from login_submit.asp
CREATE PROCEDURE spASRIntCreateCacheTables
AS
BEGIN

	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@sActualLoginName	varchar(250)

	-- Get the current user's group ID.
	EXEC spASRIntGetActualUserDetails
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT


	-- Create the SysProtects cache table
	IF EXISTS (SELECT 'x' FROM tempdb..sysobjects WHERE type = 'U' and NAME = '#SysProtects')
		DROP TABLE #SysProtects

	CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int)
	INSERT #SysProtects
	SELECT ID, Action, Columns, ProtectType
		FROM sysprotects
		WHERE uid = @iUserGroupID

	CREATE INDEX [IDX_ID] ON #SysProtects(ID)

	-- Create the ASRSysUserSettings





END