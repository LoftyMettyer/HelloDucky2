CREATE PROCEDURE [dbo].[sp_ASRGetAllObjectNames]
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
	SET NOCOUNT ON;
	SELECT DISTINCT Username FROM ASRSysAllObjectNamesForOpenHRWeb WHERE NOT NULLIF(username,'') = '' ORDER BY username;
END
