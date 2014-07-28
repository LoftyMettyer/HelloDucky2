CREATE PROCEDURE [dbo].[spASRIntGetEmailAddresses]
(@baseTableID int)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT convert(char(10),e.emailid) AS [ID], e.name AS [Name]
		FROM ASRSysEmailAddress e
		WHERE e.tableid = @baseTableID OR e.tableid = 0
		ORDER BY e.name;

END