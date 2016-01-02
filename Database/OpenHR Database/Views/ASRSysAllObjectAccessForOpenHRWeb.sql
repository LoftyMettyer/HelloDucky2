CREATE VIEW [dbo].[ASRSysAllObjectAccessForOpenHRWeb]
	AS
		SELECT CASE b.[IsBatch] 
					WHEN 0 THEN 29
					WHEN 1 THEN 0
				END	AS [objectType], a.* FROM ASRSysBatchJobAccess a
			INNER JOIN ASRSysBatchJobName b ON a.ID = b.ID
		UNION
		SELECT CASE m.[CrossTabType]
					WHEN 0 THEN 1
					WHEN 4 THEN 35
				END	AS [objectType], a.* FROM [ASRSysCrossTabAccess] a
			INNER JOIN ASRSysCrossTab m ON a.ID = m.CrossTabID
		UNION
		SELECT 2 AS [objectType], * FROM [ASRSysCustomReportAccess]
		UNION
		SELECT 3 AS [objectType], * FROM [ASRSysDataTransferAccess]
		UNION
		SELECT 4 AS [objectType], * FROM [ASRSysExportAccess]
		UNION		
		SELECT CASE [type] 
					WHEN 'A' THEN 5
					WHEN 'D' THEN 6
					WHEN 'U' THEN 7
				END	AS [objectType], a.* FROM [ASRSysGlobalAccess] a
			INNER JOIN ASRSysGlobalFunctions g ON a.ID = g.functionID	
		UNION
		SELECT 8 AS [objectType], * FROM [ASRSysImportAccess]
		UNION
		SELECT CASE m.[IsLabel] 
					WHEN 0 THEN 9
					WHEN 1 THEN 18
				END	AS [objectType], a.* FROM ASRSysMailMergeAccess a
			INNER JOIN ASRSysMailMergeName m ON a.ID = m.MailMergeID
		UNION
		SELECT CASE [MatchReportType] 
				WHEN 0 THEN 14 
				WHEN 1 THEN 23
				WHEN 2 THEN 24 
			END	AS [objectType], a.* FROM ASRSysMatchReportAccess a
			INNER JOIN ASRSysMatchReportName m ON a.ID = m.MatchReportID
		UNION
		SELECT 38 AS [objectType], a.* FROM ASRSysTalentReportAccess a
			INNER JOIN ASRSysTalentReports m ON a.ID = m.ID			
		UNION
		SELECT 17 AS [objectType], * FROM ASRSysCalendarReportAccess
		UNION
		SELECT 20 AS [objectType], * FROM [ASRSysRecordProfileAccess]
GO


