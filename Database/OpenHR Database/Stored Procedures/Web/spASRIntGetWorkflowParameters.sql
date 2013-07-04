CREATE PROCEDURE spASRIntGetWorkflowParameters 
(
	@pfWFEnabled			bit	OUTPUT
)
AS
BEGIN
	
	SET NOCOUNT ON;

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'WORKFLOW', @pfWFEnabled OUTPUT;

END

GO

