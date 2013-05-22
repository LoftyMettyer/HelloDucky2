CREATE PROCEDURE [dbo].[spASRIntGetDocumentManagementTypes]	
	AS
	BEGIN
		SET NOCOUNT ON
		SELECT [DocumentMapID], 
			[Name], 
			[Username], 
			[Access] 
		FROM [ASRSysDocumentManagementTypes] 
		ORDER BY [Name]
	END