CREATE PROCEDURE spASRGetWorkflowIDFromName(
	@name varchar(255),
	@id integer OUTPUT)
AS
BEGIN

	IF (SELECT COUNT(id) FROM ASRSysWorkflows WHERE Name = @name) = 1
		SELECT @id = id FROM ASRSysWorkflows WHERE Name = @name;
	ELSE
		SET @id = 0;

END