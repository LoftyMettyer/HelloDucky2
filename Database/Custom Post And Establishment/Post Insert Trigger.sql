IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Post_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Post_P&E]
GO

CREATE TRIGGER [dbo].[trcustom_Post_P&E] ON [dbo].[tbuser_Post_Records]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	--SELECT Effective_Date FROM inserted



END