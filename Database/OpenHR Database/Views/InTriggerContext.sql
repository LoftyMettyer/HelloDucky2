CREATE VIEW [dbo].InTriggerContext
  WITH SCHEMABINDING
AS
SELECT TOP 16 [TableFromId], [NestLevel], [ActionType]
   FROM dbo.udfsysGetContextTable()

