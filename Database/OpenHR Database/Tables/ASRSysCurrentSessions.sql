CREATE TABLE ASRSysCurrentSessions(
  [IISServer]		nvarchar(255) NULL,
	[Username]		nvarchar(128),
	[Hostname]		nvarchar(255),
	[SessionID]		nvarchar(255),
	[loginTime]		datetime,
	[WebArea]	varchar(255)
 )