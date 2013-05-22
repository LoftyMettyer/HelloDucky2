CREATE PROCEDURE [dbo].[spASRGetWindowsUsers]
(
	@DomainName varchar(200),
	@UserString varchar(MAX) OUTPUT
)
AS
BEGIN
	SELECT @UserString = dbo.udfASRNetGetUsers(@DomainName);
END