CREATE PROCEDURE [dbo].[spASRIntGetEmailGroupAddresses]
	(@EmailGroupID int)
AS
BEGIN

	select Fixed from ASRSysEmailAddress
	where EmailID in
	(select EmailDefID from ASRSysEmailGroupItems where EmailGroupID = @EmailGroupID)
	order by [Name];

END