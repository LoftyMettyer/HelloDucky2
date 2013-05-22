CREATE PROCEDURE spASRIntGetWorkflowParameters 
(
	@pfWFEnabled			bit	OUTPUT
)
AS
BEGIN
	
	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'WORKFLOW', @pfWFEnabled OUTPUT

END

GO

