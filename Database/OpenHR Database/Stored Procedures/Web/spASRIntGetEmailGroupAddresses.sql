CREATE PROCEDURE [dbo].[spASRIntGetEmailGroupAddresses]
	(@EmailGroupID int)
AS
BEGIN

	SET NOCOUNT ON;

	select Fixed from ASRSysEmailAddress
	where EmailID in
	(select EmailDefID from ASRSysEmailGroupItems where EmailGroupID = @EmailGroupID)
	order by [Name];

END