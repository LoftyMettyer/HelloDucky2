CREATE PROCEDURE [dbo].[sp_ASRIntGetEmailAddresses]
(@baseTableID int)
AS
BEGIN

  SELECT convert(char(10),ASRSysEmailAddress.emailid)+(case when (ASRSystables.DefaultEmailID = ASRSysEmailAddress.emailid) then '1 ' else '0 ' end)+ASRSysEmailAddress.name AS 'columnDefn'
	  FROM ASRSysEmailAddress
	  LEFT OUTER JOIN ASRSystables ON ASRSystables.tableid = ASRSysEmailAddress.tableid
	  WHERE ASRSysEmailAddress.tableid = @baseTableID OR ASRSysEmailAddress.tableid = 0
	  ORDER BY ASRSysEmailAddress.name;

END