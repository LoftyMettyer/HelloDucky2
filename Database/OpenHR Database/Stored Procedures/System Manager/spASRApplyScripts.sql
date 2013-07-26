CREATE PROCEDURE [dbo].[spASRApplyScripts] (@runtype integer)
	AS
	BEGIN
		
		SET NOCOUNT ON;

		DECLARE @NVarCommand nvarchar(MAX);
		DECLARE @changes table(id uniqueidentifier, [file] nvarchar(MAX), [sequence] integer);
		
		-- Collate hotfixes
		INSERT @changes
			SELECT [id], [file], [sequence]
				FROM dbo.[tbsys_scriptedchanges]
				WHERE (runtype = @runtype) AND ([runonce] = 0 OR ([runonce] = 1 AND [lastrundate] IS NULL))
				ORDER BY [sequence];

		-- Build hotixes and apply
		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand + [file]
			FROM @changes
			ORDER BY [sequence];
		EXECUTE sp_executeSQL @NVarCommand;

		-- Mark the hotfixes as complete
		UPDATE [tbsys_scriptedchanges]
			SET [lastrundate] = GETDATE()
			FROM @changes c WHERE c.id = [tbsys_scriptedchanges].id;

	END