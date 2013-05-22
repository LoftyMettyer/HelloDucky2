CREATE Procedure sp_ASRLockDelete (@LockType int)
AS
BEGIN
	DELETE FROM ASRSysLock WHERE Priority = @LockType
END
GO

