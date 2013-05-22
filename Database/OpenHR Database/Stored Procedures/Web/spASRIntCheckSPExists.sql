CREATE PROCEDURE [dbo].[spASRIntCheckSPExists]
(
	@psPrefix			varchar(255),
	@plngTableID		integer,
	@pfExists			bit		OUTPUT
)
AS
BEGIN
	DECLARE	@sSPName	varchar(MAX),
			@iCount		integer;

	SET @pfExists = 0;
	SET @sSPName = @psPrefix + convert(varchar(255), @plngTableID);

	IF NOT @sSPName IS null
	BEGIN
		SELECT @iCount = COUNT([Name])
		FROM sysobjects
		WHERE type = 'P'
			AND name = @sSPName;

		IF @iCount > 0 
		BEGIN
			SET @pfExists = 1;
		END
	END
END