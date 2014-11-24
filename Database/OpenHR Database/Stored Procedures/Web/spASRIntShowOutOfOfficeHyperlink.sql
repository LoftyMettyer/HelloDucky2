CREATE PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink]	
	(
		@piTableID		integer,
		@piViewID		integer,
		@pfDisplayHyperlink	bit 	OUTPUT
	)
	AS
	BEGIN

		SET NOCOUNT ON;

		SELECT @pfDisplayHyperlink = WFOutOfOffice
			FROM ASRSysSSIViews
			WHERE (TableID = @piTableID) 
				AND (ViewID = @piViewID);

		SELECT ISNULL(@pfDisplayHyperlink, 0);

	END
GO