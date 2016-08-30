CREATE PROCEDURE [dbo].[spsys_TrackTriggerInsert](@TableFromID integer, @NestLevel tinyint, @actionType tinyint)
AS
BEGIN

   BEGIN TRY

    IF ISNULL(len(@TableFromID),0) = 0
       RAISERROR('Context Key may not by null or empty.',11,1);

    DECLARE @buffer varchar(128) = '';

    SELECT @buffer += convert(varchar(125),[TableFromId]) + CHAR(2) + convert(varchar(3),[NestLevel]) + CHAR(2) + convert(varchar(3),[ActionType]) + CHAR(3)
      FROM [InTriggerContext]
      WHERE [TableFromId] != @TableFromID;

    IF LEN(@buffer) + LEN(@TableFromID) + LEN(@NestLevel)  > 126
       RAISERROR('Context buffer overflow.',11,1);

    IF ISNULL(len(@NestLevel),0) > 0
       SELECT @buffer += convert(varchar(125), @TableFromID) + CHAR(2) + convert(varchar(3),@NestLevel) + CHAR(2) + convert(varchar(3), @actionType) + CHAR(3)

    DECLARE @varbin varbinary(128) = convert(varbinary(128),@buffer);
    SET CONTEXT_INFO @varbin;

  END TRY
  BEGIN CATCH
    DECLARE @ErrMsg nvarchar(4000)=isnull(ERROR_MESSAGE(),'Error caught in setContextValue'), @ErrSeverity int=ERROR_SEVERITY();
  END CATCH

  FINALLY:

  if @ErrSeverity > 0  RAISERROR(@ErrMsg, @ErrSeverity, 1);

  RETURN isnull(len(@buffer),0);

END