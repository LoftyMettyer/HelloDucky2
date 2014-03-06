CREATE TABLE ASRSysCurrentLogins(
	[username]		nvarchar(128),
	[usergroup]		nvarchar(255),
	[usergroupid]	integer,
	[userSID]		uniqueidentifier,
	[loginTime]		datetime,
	[application]	varchar(255),
	[clientmachine]		nvarchar(255))