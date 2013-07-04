CREATE PROCEDURE [dbo].[spASRIntGetAvailablePicklists] (
	@plngTableID	integer,
	@psUserName		varchar(255)
)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT picklistid AS [ID], 
		name 
	FROM [dbo].[ASRSysPicklistName]
	WHERE tableid = @plngTableID 
		AND (username = @psUserName 
			OR Access <> 'HD') 
	ORDER BY [name];
END