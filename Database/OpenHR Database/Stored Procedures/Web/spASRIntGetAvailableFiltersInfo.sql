CREATE PROCEDURE [dbo].[spASRIntGetAvailableFiltersInfo] (
	@plngTableID	integer,
	@psUserName		varchar(255)
)
AS
BEGIN
	SELECT exprid AS [ID], 
		name, 
		username, 
		access 
	FROM [dbo].[ASRSysExpressions]
	WHERE tableid = @plngTableID 
		AND type = 11 
		AND (returnType = 3 OR type = 10) 
		AND parentComponentID = 0 
		AND (username = @psUserName 
			OR Access <> 'HD') 
	ORDER BY [name];
END