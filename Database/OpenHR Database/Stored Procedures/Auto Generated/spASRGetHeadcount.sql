﻿-- This stored procedure is a code stub and will be regenerated by a System Manager save.
CREATE PROCEDURE spASRGetHeadcount (
	@type		integer,
	@today	datetime)
WITH ENCRYPTION
AS
BEGIN

   SET NOCOUNT ON;

	 SELECT TableID AS ID 
		FROM dbo.tbsys_tables 
		WHERE TableID = 1;

END