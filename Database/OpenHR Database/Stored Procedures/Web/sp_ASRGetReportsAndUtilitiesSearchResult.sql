-- ==========================================================================================
-- Author:		Prashant Shah
-- Create date: 05/Nov/2015
-- Description:	Gets the reports/utilities whose definition name mathches the search criteria.
-- ==========================================================================================

CREATE PROCEDURE [dbo].[sp_ASRGetReportsAndUtilitiesSearchResult] (
	@searchText varchar(50) = NULL
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the details with which to populate the intranet defsel grid. */
	DECLARE 
		@sRoleName			varchar(255),
		@sActualUserName	varchar(250),
		@iActualUserGroupID	integer
	
	EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sRoleName OUTPUT,
			@iActualUserGroupID OUTPUT;

	SELECT 
		ASRSysAllObjectNamesForOpenHRWeb.ID,
		ASRSysAllObjectNamesForOpenHRWeb.objecttype,
		ASRSysAllObjectNamesForOpenHRWeb.Name,
		CASE ASRSysAllObjectNamesForOpenHRWeb.objecttype
			WHEN 1 THEN 'CrossTab: ' + ASRSysAllObjectNamesForOpenHRWeb.name 
			WHEN 2 THEN 'Custom Report: ' + ASRSysAllObjectNamesForOpenHRWeb.name 
			WHEN 9 THEN 'Mail Merge: ' + ASRSysAllObjectNamesForOpenHRWeb.name 
			WHEN 17 THEN 'Calender Report: ' + ASRSysAllObjectNamesForOpenHRWeb.name 
			WHEN 35 THEN '9 Box Grid Report: ' + ASRSysAllObjectNamesForOpenHRWeb.name 
		END TextToDisplay, 
		ASRSysAllObjectNamesForOpenHRWeb.description AS [description],
		Access
	FROM ASRSysAllObjectNamesForOpenHRWeb 
		INNER JOIN ASRSysAllObjectAccessForOpenHRWeb ON 
					ASRSysAllObjectNamesForOpenHRWeb.ID = ASRSysAllObjectAccessForOpenHRWeb.ID AND 
					ASRSysAllObjectAccessForOpenHRWeb.groupname = @sRoleName AND 
					(ASRSysAllObjectAccessForOpenHRWeb.access <> 'HD' OR ASRSysAllObjectNamesForOpenHRWeb.userName = SYSTEM_USER) 
	WHERE	ASRSysAllObjectAccessForOpenHRWeb.objecttype = ASRSysAllObjectNamesForOpenHRWeb.objecttype AND 
			ASRSysAllObjectNamesForOpenHRWeb.objecttype IN (1,2,9,17,35) AND 
			ASRSysAllObjectNamesForOpenHRWeb.name LIKE '%' + @searchText + '%'
	ORDER By TextToDisplay

END