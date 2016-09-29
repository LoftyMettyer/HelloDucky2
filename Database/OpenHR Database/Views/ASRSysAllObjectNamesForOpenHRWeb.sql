CREATE VIEW [dbo].[ASRSysAllObjectNamesForOpenHRWeb]
AS
		SELECT 25 AS [objectType], [ID], [Name], '' AS Username, description FROM ASRSysWorkflows
		UNION
		SELECT CASE [IsBatch] 
				WHEN 0 THEN 29
				WHEN 1 THEN 0
			END	, ID,  Name, Username, description FROM ASRSysBatchJobName
		UNION
		SELECT CASE [IsLabel] 
				WHEN 0 THEN 9
				WHEN 1 THEN 18
			END	AS [objectType],  MailMergeID AS ID, Name, Username, description FROM ASRSysMailMergeName
		UNION
		SELECT 2 AS [objectType], ID, Name, Username, description FROM ASRSysCustomReportsName
		UNION
		SELECT CASE [CrossTabType]
			   WHEN 0 THEN 1
			   WHEN 4 THEN 35
			   END  AS [objectType], CrossTabID AS ID, Name, Username, description FROM ASRSysCrossTab
		UNION		
		SELECT CASE [MatchReportType] 
				WHEN 0 THEN 14 
				WHEN 1 THEN 23
				WHEN 2 THEN 24 
			END	AS [objectType], MatchReportID AS ID, Name, Username, description FROM ASRSysMatchReportName			
		UNION
		SELECT 4 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysExportName
		UNION		
		SELECT 8 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysImportName
		UNION
		SELECT 3 AS [objectType], DataTransferID AS ID, Name, Username, description FROM ASRSysDataTransferName
		UNION
		SELECT CASE [type] 
				WHEN 'A' THEN 5
				WHEN 'D' THEN 6
				WHEN 'U' THEN 7
			END	AS [objectType], [FunctionID] AS ID, Name, Username, description FROM ASRSysGlobalFunctions
		UNION		
		SELECT 15 AS [objectType], 0 AS ID, 'Absence Breakdown', '' AS Username, '' AS Description
		UNION
		SELECT 16 AS [objectType], 0 AS ID, 'Bradford Factor', '' AS Username, '' AS Description
		UNION
		SELECT 17 AS [objectType], ID AS ID, Name, Username, description FROM ASRSysCalendarReports
		UNION		
		SELECT 20 AS [objectType], RecordProfileID AS ID, Name, Username, description FROM ASRSysRecordProfileName
		UNION
		SELECT 30 AS [objectType], 0 AS ID, 'Turnover', '' AS Username, '' AS Description
		UNION
		SELECT 31 AS [objectType], 0 AS ID, 'Stability Index', '' AS Username, '' AS Description
		UNION
		SELECT 38 AS [objectType], ID, Name, Username, description FROM ASRSysTalentReports
		UNION
		SELECT 39 AS [objectType], ID, Name, Username, description FROM ASRSysOrganisationReport
GO


