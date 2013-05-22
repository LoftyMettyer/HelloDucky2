CREATE PROCEDURE [dbo].[sp_ASROp_And]
(	
	@pfResult	bit OUTPUT,
	@pfFirst	bit,
	@pfSecond	bit
)
AS
BEGIN
	SET @pfResult = @pfFirst & @pfSecond;
END