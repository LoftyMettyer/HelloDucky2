CREATE FUNCTION dbo.[udfsysStringToTable] (           
      @String nvarchar(MAX),
      @delimiter nvarchar(2))
RETURNS @Table TABLE( Splitcolumn nvarchar(MAX)) 
BEGIN

	DECLARE @Xml AS XML;
	SET @Xml = cast(('<A>'+replace(@String,@delimiter,'</A><A>')+'</A>') AS XML);

	INSERT INTO @Table SELECT LTRIM(RTRIM(A.value('.', 'nvarchar(max)'))) AS [Column] FROM @Xml.nodes('A') AS FN(A);
	RETURN;

END