CREATE PROCEDURE [dbo].[spASRIntGetPersonnelParameters] (
	@piEmployeeTableID	integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	SET @piEmployeeTableID = 0;

	-- Get the EMPLOYEE table information.
	SELECT @piEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel';
	IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0;

END