CREATE PROCEDURE sp_ASRRemoveLock
AS
BEGIN
	IF EXISTS(SELECT * FROM sysobjects WHERE name = 'tmpLock') DELETE FROM tmpLock	
	DROP TABLE tmpLock
END



GO

