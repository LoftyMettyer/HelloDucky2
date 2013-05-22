CREATE PROCEDURE [dbo].[sp_ASRFn_IFThenElse]
(
	@testvalue		bit,
	@date1			datetime,
	@date2			datetime,
	@retdate   		datetime OUTPUT,
	@char1			varchar(MAX),
	@char2			varchar(MAX),
	@retchar   		varchar(MAX) OUTPUT,
	@numeric1		numeric,
	@numeric2		numeric,
	@retnumeric   	numeric OUTPUT,
	@logic1			bit,
	@logic2			bit,
	@retlogic		bit OUTPUT
)
AS
BEGIN

	IF @date1 IS NOT NULL
	BEGIN
		IF @testvalue = 1
		BEGIN
			SET @retdate = @date1;
			SELECT @retdate AS result;
		END
		IF @testvalue = 0
		BEGIN
			SET @retdate = @date2;
			SELECT @retdate AS result;
		END	
	END

	IF @char1 IS NOT NULL
	BEGIN
		IF @testvalue = 1
		BEGIN
			SET @retchar = @char1;
			SELECT @retchar AS result;
		END
		IF @testvalue = 0
		BEGIN
			SET @retchar = @char2;
			SELECT @retchar AS result;
		END	
	END

	IF @numeric1 IS NOT NULL
	BEGIN
		IF @testvalue = 1
		BEGIN
			SET @retnumeric = @numeric1;
			SELECT @retnumeric AS result;
		END
		IF @testvalue = 0
		BEGIN
			SET @retnumeric = @numeric2;
			SELECT @retnumeric AS result;
		END	
	END

	IF @logic1 IS NOT NULL
	BEGIN
		IF @testvalue = 1
		BEGIN
			SET @retlogic = @logic1;
			SELECT @retlogic AS result;
		END
		IF @testvalue = 0
		BEGIN
			SET @retlogic = @logic2;
			SELECT @retlogic AS result;
		END	

	END
END