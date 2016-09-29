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
		son.ID,
		son.objecttype,
		son.Name,
		CASE son.objecttype
			WHEN 1 THEN 'Cross Tab Report: ' + son.name 
			WHEN 2 THEN 'Custom Report: ' + son.name 
			WHEN 9 THEN 'Mail Merge: ' + son.name 
			WHEN 17 THEN 'Calendar Report: ' + son.name 
			WHEN 35 THEN '9-Box Grid Report: ' + son.name 
			WHEN 38 THEN 'Talent Report: ' + son.name
			WHEN 39 THEN 'Organisation Report: ' + son.name
		END TextToDisplay, 
		son.description AS [description],
		Access
	FROM ASRSysAllObjectNamesForOpenHRWeb  son
		INNER JOIN ASRSysAllObjectAccessForOpenHRWeb soa ON 
					son.ID = soa.ID AND 
					soa.groupname = @sRoleName AND 
					(soa.access <> 'HD' OR son.userName = SYSTEM_USER) 
	WHERE	soa.objecttype = son.objecttype AND 
			son.objecttype IN (1,2,9,17,35,38,39) AND 
			son.name LIKE '%' + @searchText + '%'
	ORDER By TextToDisplay

END