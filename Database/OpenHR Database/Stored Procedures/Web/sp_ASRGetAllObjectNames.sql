-- =======================================================
-- Author:		Amrit
-- Create date: 05/Nov/2015
-- Description:	Gets the report/utilities creator users.
-- =======================================================
CREATE PROCEDURE [dbo].[sp_ASRGetAllObjectNames]
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
	SET NOCOUNT ON;
	SELECT DISTINCT Username FROM ASRSysAllObjectNamesForOpenHRWeb WHERE NOT NULLIF(username,'') = '' ORDER BY username;
END
