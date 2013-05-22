CREATE PROCEDURE [dbo].[spASRGetDomains]
	(@DomainString varchar(MAX) OUTPUT)
AS
BEGIN

	SELECT @DomainString = dbo.udfASRNetGetDomains();

END