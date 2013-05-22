CREATE PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]
(
	@psResult	varchar(255) OUTPUT
)
AS
BEGIN
	SET @psResult = 
		CASE 
			WHEN UPPER(LEFT(APP_NAME(), 15)) = 'OPENHR WORKFLOW' THEN 'OpenHR Workflow' 
			ELSE SUSER_SNAME()
		END;
END