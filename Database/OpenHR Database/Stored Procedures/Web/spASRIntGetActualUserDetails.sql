CREATE PROCEDURE [dbo].[spASRIntGetActualUserDetails]
(
		@psUserName sysname OUTPUT,
		@psUserGroup sysname OUTPUT,
		@piUserGroupID integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;
	EXEC dbo.spASRGetActualUserDetails @psUserName OUTPUT, @psUserGroup OUTPUT, @piUserGroupID  OUTPUT, '';

END

