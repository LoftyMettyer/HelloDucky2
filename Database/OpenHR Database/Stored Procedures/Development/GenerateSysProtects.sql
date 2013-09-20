DECLARE @iUserGroupID	integer, 
	@sUserGroupName		sysname, 
	@sActualLoginName	varchar(250);

SET @iUserGroupID = 1;		-- Replace this value

EXEC spASRIntGetActualUserDetails 
	@sActualLoginName OUTPUT, 
	@sUserGroupName OUTPUT, 
	@iUserGroupID OUTPUT 
IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL 
	DROP TABLE #SysProtects 
CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) 
	INSERT #SysProtects 
	SELECT ID, Action, Columns, ProtectType 
       FROM sysprotects 
       WHERE uid = @iUserGroupID;
