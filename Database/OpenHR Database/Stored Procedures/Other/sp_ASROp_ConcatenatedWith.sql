CREATE PROCEDURE [dbo].[sp_ASROp_ConcatenatedWith] 
(
	@psResult 	varchar(MAX) OUTPUT,
	@psString1 	varchar(MAX),
	@psString2	varchar(MAX)
)
AS
BEGIN
	SET @psResult = @psString1 + @psString2;
END