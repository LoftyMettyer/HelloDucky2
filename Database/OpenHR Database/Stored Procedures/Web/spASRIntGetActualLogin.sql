CREATE PROCEDURE [dbo].[spASRIntGetActualLogin]
(
	@psActualLogin		nvarchar(255)	OUTPUT
)
AS
BEGIN

	DECLARE @iFound	integer;

	/* Is this user logged in under a specific login, or as part of a Windows Group login */
	SELECT @iFound = Count(*) FROM sysusers WHERE name = SYSTEM_USER;
	IF (@iFound > 0)
		SET @psActualLogin = SYSTEM_USER;
	ELSE
		SELECT TOP 1 @psActualLogin = name FROM sysusers
			WHERE is_member(Name) & IsNTGroup = 1;

END